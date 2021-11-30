function parseOneLineComment(srclines) {
    var lines = [];

    _.forEach(srclines, function(lineObj) {
        var line = lineObj.line;
        var cppCommentIndex = line.search(/\s*\/\//);

        if (cppCommentIndex != -1) {
            line = line.slice(0, cppCommentIndex);
            //line = _.trimRight(line);
            if (line === "") {
                return;
            }
            // 行末のコメント
            lineObj.comment = lineObj.line.slice(cppCommentIndex);
            lineObj.line = line;
        }

        lines.push(lineObj);
    });

    return lines;
}

// それぞれ行頭、行末に書かれた <!-- と --> のみ対応
// 入れ子に対応
// C style コメントについてはごくごく簡易的なもの
// 本来 C style コメントは入れ子には対応してないけど、そこまでは対応しない
function parseMultilineComment(srcLines) {
    var lines = [];
    srcLines = new ArrayReader(srcLines);

    var ParseCommentError = function(errorMessage, lineObj) {
        this.errorMessage = errorMessage;
        this.lineObj = lineObj;
    };
    
    function _parseNest(beginRe, endRe, lineObj) {
        if (!beginRe.test(lineObj.line)) {
            return false;
        }

        var beginLine = lineObj;

        for (var commentDepth = 0;;) {
            var line = lineObj.line;

            if (beginRe.test(line)) {
                commentDepth++;
            }
            if (endRe.test(line)) {
                commentDepth--;
            }
            if (commentDepth == 0) {
                break;
            }
            if (srcLines.atEnd) {
                if (commentDepth > 0) {
                    var errorMessage = "コメントが閉じていません";
                    throw new ParseCommentError(errorMessage, beginLine);
                }
            }
            lineObj = srcLines.read();
        }
        return true;
    }

    while (!srcLines.atEnd) {
        var lineObj = srcLines.read();

        if (_parseNest(/^\s*<!--.*/, /.*-->\s*$/, lineObj)) {
            continue;
        }
        if (_parseNest(/^\s*\/\*.*/, /.*\*\/\s*$/, lineObj)) {
            continue;
        }

        lines.push(lineObj);
    }

    return lines;
}


// #define とか #if else endif 的なの
// コメント削除が適用済みのを渡す
// objs: 定義済みマクロ変数
function preProcessConditionalCompile(lines, defines) {
    var srcLines = new ArrayReader(lines);
    var dstLines = [];
    var objs = defines;
    var states = []; // 入れ子対応のためスタックにしておく

    function currentCondtion() {
        for (var i = 0; i < states.length; i++) {
            if (!states[i].cond) {
                return false;
            }
        }
        return true;
    }

    function evalFormula(formula, objs) {
        // XXX: 改行は処理的には不要だけど、デバッグ中に頻繁に出力する都合上付けておく
        var s = "(function(){\n";

        // 宣言されてない変数名は false 扱い
        var ids = formula.trim().match(/([a-zA-Z_]\w*)/g);
        if (ids) {
            var undefs = _.difference(ids, _.keys(objs));
            undefs = _.difference(undefs, ['true', 'false']);
            _.forEach(undefs, function(name) {
                s += "var " + name + "=false;\n";
            });
        }

        for (var name in objs) {
            s += "var " + name + "=" + JSON.stringify(objs[name]) + ";\n";
        }

        s += "return(" + formula + ");})();";
        //WScript.Echo(s);
        return eval(s);
    }

    // parseError にすると ParseError を上書きするようなので parseSharpError にしておく
    function parseSharpError(option, lineObj) {
        if (!currentCondtion()) {
            return;
        }
        var errorMessage = "@error";
        var text = _.trim(option);
        if (text != "") {
            errorMessage += " : '" + text + "'";
        }
        else {
            errorMessage += " が発生しました。";
        }
        throw new ParseError(errorMessage, lineObj);
    }

    function parseDefine(option, lineObj) {
        if (!currentCondtion()) {
            return;
        }

        var name = option.trim();
        if (!/^([a-zA-Z_]\w*)?$/.test(name)) {
            var errorMessage = "@define コマンドの文法が正しくありません。";
            throw new ParseError(errorMessage, lineObj);
        }

        // define の場合は set true 扱い
        parseSet(name + " = true", lineObj);
    }

    function isReservedName(name) {
        var reserved = {
            'break': null,
            'case': null,
            'catch': null,
            'continue': null,
            'debugger': null,
            'default': null,
            'delete': null,
            'do': null,
            'else': null,
            'finally': null,
            'for': null,
            'function': null,
            'if': null,
            'in': null,
            'instanceof': null,
            'new': null,
            'return': null,
            'switch': null,
            'this': null,
            'throw': null,
            'try': null,
            'typeof': null,
            'var': null,
            'void': null,
            'while': null,
            'with': null,

            'true': null,
            'false': null,
            'undefined': null,
            'null': null
        };
        return (name in reserved);
    }

    function parseUndef(option, lineObj) {
        if (!currentCondtion()) {
            return;
        }

        var name = option.trim();
        if (!/^([a-zA-Z_]\w*)?$/.test(name)) {
            var errorMessage = "@undef コマンドの文法が正しくありません。";
            throw new ParseError(errorMessage, lineObj);
        }

        if (!isReservedName(name)) {
            //if (name in objs) {
            delete objs[name];
            //}
        }

        // undef の場合は set false 扱い、にしようと思ったけど undef の後で define で redefine 扱いになるので
        // 素直に削除だけにしておく
        //parseSet(name + " = false", lineObj);
    }

    function parseSet(option, lineObj) {
        if (!currentCondtion()) {
            return;
        }

        var optionMatch = option.trim().match(/^([a-zA-Z_]\w*)\s*=\s*(.+)?$/);
        if (optionMatch === null) {
            var errorMessage = "@set コマンドの文法が正しくありません。";
            throw new ParseError(errorMessage, lineObj);
        }
        //WScript.Echo(JSON.stringify(optionMatch, undefined, 4));
        var name = optionMatch[1];

        if (isReservedName(name)) {
            var errorMessage = "変数名に予約語が使われています。";
            throw new ParseError(errorMessage, lineObj);
        }
        if (name in objs) {
            var errorMessage = "変数 " + name + " はすでに定義されています。";
            throw new ParseError(errorMessage, lineObj);
        }
        var value = optionMatch[2];
        try {
            objs[name] = evalFormula(value, objs);
        }
        catch (e) {
            var errorMessage = '右辺の式 "' + value + '" が不正です。';
            throw new ParseError(errorMessage, lineObj);
        }
        lineObj.define = {
            name: name,
            value: objs[name]
        };
        //WScript.Echo(JSON.stringify(objs, undefined, 4));
    }

    function parseCondition(cond, lineObj) {
        try {
            return evalFormula(cond, objs);
        }
        catch (e) {
            var errorMessage = "条件式が不正です。";
            throw new ParseError(errorMessage, lineObj);
        }
    }

    function parseIf(option, lineObj) {
        var state = {
            cond: false, // 今のフラグ
            elseApplied: false,
            condDisabled: true,  // これが立ってたらつねに false 扱い

            lineObj: lineObj
        };
        if (currentCondtion()) {
            var cond = parseCondition(option.trim(), lineObj);
            state.cond = cond;
            state.condDisabled = cond;
        }

        states.push(state);
    }
    function parseElif(option, lineObj) {
        // いきなり elif 出現エラー
        if (states.length === 0) {
            var errorMessage = "対応する if がありません。";
            throw new ParseError(errorMessage, lineObj);
        }

        var state = _.last(states);
        // 今の階層ですでに else が処理済みならエラー
        if (state.elseApplied) {
            var errorMessage = "elif が else の後に存在します。";
            throw new ParseError(errorMessage, lineObj);
        }

        if (state.condDisabled) {
            state.cond = false;
            return;
        }

        var cond = parseCondition(option.trim(), lineObj);
        if (cond) {
            state.cond = true;
            state.condDisabled = true;
        }
    }
    function parseElse(option, lineObj) {
        // いきなり else 出現エラー
        if (states.length === 0) {
            var errorMessage = "対応する if がありません。";
            throw new ParseError(errorMessage, lineObj);
        }

        var state = _.last(states);
        state.cond = !state.condDisabled;
        state.elseApplied = true;
    }
    function parseEnd(option, lineObj) {
        // いきなり end 出現エラー
        if (states.length === 0) {
            var errorMessage = "対応する if がありません。";
            throw new ParseError(errorMessage, lineObj);
        }
        states.pop();
    }

    try {

    while (!srcLines.atEnd) {
        var lineObj = srcLines.read();
        var line = lineObj.line;
    
        //WScript.Echo(JSON.stringify(lineObj, undefined, 4));
        if (!/^@.*/.test(line)) {
            if (currentCondtion()) {
                dstLines.push(lineObj);
            }
            continue;
        }
        //WScript.Echo(">>>\n" + JSON.stringify(lineObj, undefined, 4));

        var commandMatch = line.match(/^@([a-zA-Z]+)(.+)?$/);

        //var s = "";
        //for (var i = 0; i < commandMatch.length; i++) {
        //    s += i + " : " + commandMatch[i] + "\n";
        //}
        //WScript.Echo(s);
        var command = commandMatch[1];
        var option = commandMatch[2];
        switch (command) {
            case 'define':
                parseDefine(option, lineObj);
                break;
            case 'undef':
                parseUndef(option, lineObj);
                break;
            case 'set':
                parseSet(option, lineObj);
                break;
            case 'if':
                parseIf(option, lineObj);
                break;
            case 'elif':
                parseElif(option, lineObj);
                break;
            case 'else':
                parseElse(option, lineObj);
                break;
            case 'end':
                parseEnd(option, lineObj);
                break;
            case 'error':
                parseSharpError(option, lineObj);
                break;
            default: {
                var errorMessage = "不明の@コマンドです。";
                throw new ParseError(errorMessage, lineObj);
                break;
            }
        }
    }

    if (states.length !== 0) {
        var state = _.last(states);
        var errorMessage = "@if が完結していません。@end が必要です。";
        throw new ParseError(errorMessage, state.lineObj);
    }

    //WScript.Echo(JSON.stringify(objs, undefined, 4));

    // TODO: if が end で終わってないエラー
    // スタックがカラじゃなければエラーとすればOK？
    
    }
    catch (e) {
        if (_.isUndefined(e.lineObj) || _.isUndefined(e.errorMessage)){
            throw e;
        }
        //WScript.Echo(JSON.stringify(e, undefined, 4));
        parseError(e);
    }

    return dstLines;
}

// filePaths: 含まれるすべてのファイルのパス
function preProcess_Recurse(filePath, filePaths, pathStack, defines) {
    filePaths.push(filePath);
    var parentFolderName = fso.GetParentFolderName(filePath);

    _.last(pathStack).parentFolder = parentFolderName;

    stream.Type = adTypeText;
    // UTF-8 BOM なし 専用
    stream.charset = "UTF-8";
    stream.Open();
    stream.LoadFromFile(filePath);
    var allLines = stream.ReadText(adReadAll);
    stream.Close();

    //var path = fso.BuildPath(parentFolderName, image);

    //var lineArray = new ArrayReader(allLines.split(/\r\n|\r|\n/));
    // 空要素も結果に含めたいのでsplitには正規表現を使わないように
    var lineArray = allLines.replace(/\r\n|\r/g, "\n").split("\n");

    // 最初に lineObj にしてしまう
    var lines = [];
    _.forEach(lineArray, function(line, lineNum) {
        if (line === "") {
            return;
        }

        var lineObj = {
            line: line,
            lineNum: 1 + lineNum,   // 1 origin
            filePath: filePath
        };
        lines.push(lineObj);
    });

    try {
        lines = parseOneLineComment(lines);
        lines = parseMultilineComment(lines);
        lines = preProcessConditionalCompile(lines, defines);
    }
    catch (e) {
        if (_.isUndefined(e.lineObj) || _.isUndefined(e.errorMessage)){
            throw e;
        }
        parseError(e);
    }

    //printJSON(lines);
    //WScript.Quit(1);

    var srcLines = new ArrayReader(lines);

    var dstLines = [];

    while (!srcLines.atEnd) {
        var lineObj = srcLines.read();
        var line = lineObj.line;

        if (!_.isUndefined(lineObj.define)) {
            var define = lineObj.define;
            defines[define.name] = define.value;
        }

        var includeMatch = line.match(/^<<\[\s*(.+)\s*\]$/);

        if (includeMatch) {
            var includeFile = includeMatch[1];
            // include 元と同じ場所を探す
            var path = fso.BuildPath(parentFolderName, includeFile);

            if (!fso.FileExists(filePath)) {
                var errorMessage = "include ファイル\n" + includeFile + "\nが存在しません";
                throw new ParseError(errorMessage, lineObj);
            }

            var includeLines = preProcess_Recurse(path, filePaths, pathStack, defines);
            
            dstLines = dstLines.concat(includeLines);

            pathStack.pop();
            continue;
        }
        else {
            dstLines.push(lineObj);
        }
    }

    return dstLines;
}


function preProcess_Recurse_old(filePath, lines, filePaths, pathStack) {
    filePaths.push(filePath);
    var parentFolderName = fso.GetParentFolderName(filePath);

    _.last(pathStack).parentFolder = parentFolderName;

    stream.Type = adTypeText;
    // UTF-8 BOM なし 専用
    stream.charset = "UTF-8";
    stream.Open();
    stream.LoadFromFile(filePath);
    var allLines = stream.ReadText(adReadAll);
    stream.Close();

    //var path = fso.BuildPath(parentFolderName, image);

    //var lineArray = new ArrayReader(allLines.split(/\r\n|\r|\n/));
    // 空要素も結果に含めたいのでsplitには正規表現を使わないように
    var lineArray = allLines.replace(/\r\n|\r/g, "\n").split("\n");

    // 最初に lineObj にしてしまう
    var lines = [];
    _.forEach(lineArray, function(line, lineNum) {
        if (line === "") {
            return;
        }

        var lineObj = {
            line: line,
            lineNum: 1 + lineNum,   // 1 origin
            filePath: filePath
        };
        lines.push(lineObj);
    });

    try {
        lines = parseOneLineComment(lines);
        lines = parseMultilineComment(lines);
        lines = preProcessConditionalCompile(lines);
    }
    catch (e) {
        if (_.isUndefined(e.lineObj) || _.isUndefined(e.errorMessage)){
            throw e;
        }
        parseError(e);
    }

    printJSON(lines);
    WScript.Quit(1);

    lineArray = new ArrayReader(lines);

    while (!lineArray.atEnd) {
        var lineObj = lineArray.read();
        var line = lineObj.line;

        var includeMatch = line.match(/^<<\[\s*(.+)\s*\]$/);
        if (includeMatch) {
            var includeFile = includeMatch[1];
            var path = findIncludeFile(includeFile);

            if (!path) {
                var errorMessage = "include ファイル\n" + includeFile + "\nが存在しません";
                throw new ParseError(errorMessage, lineObj);
            }

            preProcess_Recurse(path, lines, filePaths, pathStack);
            pathStack.pop();
            continue;
        }


        // 何のために作ったか不明
        function getLastLocalPath() {
            for (var i = pathStack.length - 1; i >= 0; --i) {
                var path = pathStack[i];
                if (!path.includePath) {
                    return path.parentFolder;
                }
            }
            // ここにくるはずはない
            return null;
        }

        function findIncludeFileOverride(targetFilePath, pathStack) {
            // override の候補となるパスを返す
            function getOverridePaths(targetFilePath) {
                for (var i = pathStack.length - 1; i >= 0; --i) {
                    var path = pathStack[i];
                    if (!path.includePath) {
                        return pathStack.slice(i);
                    }
                }
                // ここにくるはずはない
                return null;
            }

            // include すべきファイルを見つける
            function findIncludeFilePath(overridePaths) {
                var lastPath = _.last(overridePaths);
                var sentinel = {
                    includePath: (!lastPath.includePath) ? "dummy" : null
                };
                overridePaths.push(sentinel);
                for (var i = 0; i < overridePaths.length - 1; i++) {
                    var overridePath = overridePaths[i];
                    if (overridePath.includePath == overridePaths[i + 1].includePath) {
                        continue;
                    }
                    var path = fso.BuildPath(overridePath.parentFolder, targetFilePath);
                    if (fso.FileExists(path)) {
                        pathStack.push({
                            includePath: overridePath.includePath
                        });
                        return path;
                    }
                }
                // ここにくるはずはない
                return null;
            }

            var overridePaths = getOverridePaths(targetFilePath);

            return findIncludeFilePath(overridePaths);
        }

        function findIncludeFileIncludePath(targetFilePath, pathStack) {
            for (var i = 0; i < includePath.length; i++) {
                var path = fso.BuildPath(includePath[i], targetFilePath);
                if (fso.FileExists(path)) {
                    pathStack.push({
                        includePath: includePath[i]
                    });
                    return path;
                }
            }
            // include path 内に指定のファイルが見つからなかった
            return null;
        }

        function findIncludeFile(targetFilePath) {
            // 最初の2文字が '~/' の場合は override 指定とみなす
            // include 元の同名のファイルを優先して読みに行く
            // home で最後に include した場所から include 方向に向かって、そのパスで最後に inlcude した場所を検索していく
            if (targetFilePath.slice(0, 2) == '~/') {
                targetFilePath = targetFilePath.slice(2);
                return findIncludeFileOverride(targetFilePath, pathStack);
            }

            // 最初の1文字が '/' の場合は include path を順に探して最初に見つかったのを採用
            if (targetFilePath.slice(0, 1) == '/') {
                targetFilePath = targetFilePath.slice(1);
                return findIncludeFileIncludePath(targetFilePath, pathStack);
            }

            // include 元と同じ場所を探す
            var filePath = fso.BuildPath(parentFolderName, targetFilePath);

            if (fso.FileExists(filePath)) {
                pathStack.push({
                    includePath: _.last(pathStack).includePath
                });

                return filePath;
            }

            return null;
        }

        var include = line.match(/^<<\[\s*(.+)\s*\]$/);
        if (include) {
            var includeFile = include[1];
            var path = findIncludeFile(includeFile);

            if (!path) {
                var errorMessage = "include ファイル\n" + includeFile + "\nが存在しません";

                Error(errorMessage, filePath, lineArray.index);
            }

//            var path = fso.BuildPath(parentFolderName, include[1]);
//
//            // ファイルが存在するか確認
//            (function () {
//                var fso = new ActiveXObject("Scripting.FileSystemObject");
//
//                if (!fso.FileExists(path)) {
//                    var relativeFilePath = getRelativePath(path, rootFilePath, fso);
//                    var errorMessage = "include ファイル\n" + relativeFilePath + "\nが存在しません";
//
//                    Error(errorMessage, filePath, lineArray.index);
//                }
//            })();
            
            preProcess_Recurse(path, lines, filePaths, pathStack);
            pathStack.pop();
            continue;
        }

        var lineObj = {
            line: line,
            lineNum: lineArray.index,   // 1 origin
            filePath: filePath
        };
        lines.push(lineObj);
    }
}

// preprocess
// include とかコメント削除とか
// 入れ子の include にも対応
function preProcess(filePath, filePaths) {
    var pathStack = [];
    var defines = {};

    pathStack.push({
        includePath: null
    });

    return preProcess_Recurse(filePath, filePaths, pathStack, defines);
}

