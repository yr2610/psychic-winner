function readConfigFile(confFileName) {
    var confFilePath = fso.BuildPath(fso.GetParentFolderName(filePath), confFileName);
    if (!fso.FileExists(confFilePath)) {
        return {};
    }
    var data = CL.readYAMLFile(confFilePath);

    var postProcess = [];

    //var indexFilePath = fso.BuildPath(fso.GetParentFolderName(filePath), "index.yml");
    //var index = CL.readYAMLFile(indexFilePath);
    //printJSON(index);

    // include文法のパス用
    if (!_.isUndefined(data.includePath)) {
        includePath = includePath.concat(data.includePath);
    }

    function processFunctions(data) {
        if (_.isUndefined(data.$functions)) {
            return;
        }
        var functions = data.$functions;
        delete data.$functions;
        _.forEach(functions, function(value, key) {
            data[key] = Function.call(this, 'return ' + value)();
        });
    }

    // path を include 先のファイル基準の絶対パスに変換
    function processPath(data, baseDirectory) {
        if (_.isUndefined(data.$rootDirectory)) {
            return;
        }

        data.$rootDirectory = fso.BuildPath(baseDirectory, data.$rootDirectory);
    }

    // 循環しないように
    // 循環の対処はしないので、無限ループになる
    function processIncludeFiles(data, baseFile) {
        var baseDirectory = fso.GetParentFolderName(baseFile);

        // XXX: ついでに template_dxl もここで処理
        //if (!_.isUndefined(data.$template_dxl)) {
        //    xmlFilePath = fso.BuildPath(baseDirectory, data.$template_dxl);
        //    delete data.$template_dxl;
        //}

        // XXX: ついでに functions もここで
        processFunctions(data);

        // XXX: クソ実装ではあるけど、path の対処もここで
        processPath(data, baseDirectory);

        if (!_.isUndefined(data.$include)) {
            var includeFiles = data.$include;
            delete data.$include;
            _.forEach(includeFiles, function(value) {
                var includeFilePath = fso.BuildPath(baseDirectory, value);
                var includeData = CL.readYAMLFile(includeFilePath);
                processIncludeFiles(includeData, includeFilePath);
                //_.assign(data, includeData);  // 上書きする
                _.defaults(data, includeData);  // 上書きしない
            });
        }

        if (!_.isUndefined(data.$post_process)) {
            var s = data.$post_process;
            var f = Function.call(this, 'return ' + s)();
            postProcess.push(f);
            delete data.$post_process;
        }
    
    }

    processIncludeFiles(data, confFilePath);

    _.forEach(postProcess, function(f) {
        f(data);
    });

    _.templateSettings = {
        evaluate: /\{\{([\s\S]+?)\}\}/g,
        interpolate: /\{\{=([\s\S]+?)\}\}/g,
        escape: /\{\{-([\s\S]+?)\}\}/g
    };
    
    // テンプレート変数の文字列に他のテンプレート変数が含まれているの対応
    (function(){
        var finished = {};
        var modified;
        var re = /\{\{[\-=]?([\s\S]+?)\}\}/;
        do {
            modified = false;
            _.forEach(data, function(value, key) {
                if (finished[key]) {
                    return;
                }
                //WScript.Echo(typeof value +"\n" + JSON.stringify(value, undefined, 4));
                if (!re.test(value)) {
                    finished[key] = true;
                    return;
                }
                var _compile = _.template(value);
            
                data[key] = _compile(data);
                modified = true;
            });
        } while (modified);

        // 下の階層は一番上の階層の参照のみ対応
        function compileForAllChildren(rootData, data) {
            _.forEach(data, function(value, key) {
                if (typeof value == "object") {
                    if (Array.isArray(value)) {
                        _.forEach(value, function(value, key) {
                            compileForAllChildren(rootData, value);
                        });
                    }
                    else {
                        compileForAllChildren(rootData, value);
                    }
                }
                else if (re.test(value)) {
                    var _compile = _.template(value);

                    data[key] = _compile(rootData);
                }
            });
        }

        compileForAllChildren(data, data);
        
    })();
    
    return  data;
}
