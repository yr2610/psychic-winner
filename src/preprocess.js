function parseOneLineComment(srclines) {
    var lines = [];

    _.forEach(srclines, function(lineObj) {
        var line = lineObj.line;
        var cppCommentIndex = line.search(/\s*\/\//);

        if (cppCommentIndex != -1) {
            if (cppCommentIndex == 0) {
                return;
            }
            // 行末のコメント
            lineObj.comment = line.slice(cppCommentIndex);
            line = line.slice(0, cppCommentIndex);
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

function replaceText(s, data) {
    // XXX: 一旦は正規表現で置換できる程度の仕様にしておく
    var rep_fn = function(m, k) {
        if (!(k in data)) {
            throw k;
        }
        return data[k];
    }

    return s.replace(/\{\{\=\s*([\w\$]+)\s*\}\}/g, rep_fn);
  
//    for (;;) {
//        var templateMatch = s.match(/^(.*)\{\{\=\s*([\w\$]+)\s*\}\}(.*)$/);
//        if (!templateMatch) {
//            break;
//        }
//        var name = templateMatch[2];
//        if (!(name in data)) {
//            throw "no member '" + name + "' in data.";
//        }
//        s = templateMatch[1] + data[name] + data[3];
//    }
//
//    return s;
}

// filename.txt とだけ指定した場合は現在のプロジェクトの source 直下
// projectname:filename.txt と指定すると外部プロジェクトの source 直下
// 外部プロジェクトは root を最優先で検索。次に include path から検索（未対応）
// プロジェクト指定なしの場合 ./filename.txt と指定するとそのファイルからの相対
function parseIncludeFilePath(s, currentProjectPathFromRoot, currentFilePathAbs, variables) {

    var IncludeFilePathError = function(errorMessage) {
        this.errorMessage = errorMessage;
    };

    try {
        s = replaceText(s, variables);
    }
    catch (e) {
        var message = "'" + e + "' を置換できません";
        throw new IncludeFilePathError(message);
    }

    var includeMatch = s.match(/^((\/)?([^:]+):)?(\.\/)?(.+)$/);

    // 無効なパス指定
    if (!includeMatch) {
        var message = "無効なパスです";
        throw new IncludeFilePathError(message);
    }

    var localPath = includeMatch[5];
    var projectDirectoryFromRoot = includeMatch[3];
    // 現在のファイル(include元)からの相対指定
    var relativeFromCurrent = (includeMatch[4] != "");

    if (relativeFromCurrent) {
        if (projectDirectoryFromRoot != "") {
            var message = "外部参照では現在のファイルからの相対指定はできません";
            throw new IncludeFilePathError(message);
        }
        var directoryFromRoot = fso.GetParentFolderName(currentFilePathAbs);

        return {
            projectDirectory: currentProjectPathFromRoot,
            sourceDirectory: directoryFromRoot,
            filePath: fso.BuildPath(directoryFromRoot, localPath)
        };
    }

    // XXX: 当面は root 以下専用
    if (projectDirectoryFromRoot == "") {
        projectDirectoryFromRoot = currentProjectPathFromRoot;
    }
    else {
        // root 指定の有無に関係なく root を優先して読む
        projectDirectoryFromRoot = fso.BuildPath(rootPathAbs, projectDirectoryFromRoot);

        var fromRoot = (includeMatch[2] != "");
        if (!fromRoot) {
            // TODO: root 指定ナシの場合は include path も検索する
        }
    }

    // source 直下からの相対

    var directoryFromRoot = fso.BuildPath(projectDirectoryFromRoot, sourceDirectory);

    var rootPathAbs = conf.$rootDirectory;
    var filePathAbs = fso.BuildPath(directoryFromRoot, localPath);
    filePathAbs = fso.BuildPath(rootPathAbs, filePathAbs);

    return {
        projectDirectory: projectDirectoryFromRoot,
        sourceDirectory: directoryFromRoot,
        filePath: filePathAbs
    };
    //WScript.Echo(localPath);
}

// パラメータは文字列のみの想定
// 文字列以外を渡された場合の動作は不定
// variables は include 元で定義済みの変数
function parseIncludeParameters(s, variables) {
    if (s == "") {
        return {};
    }

    // object を返すには丸括弧が必要らしい
    var params = eval("({" + s + "})");

    // 各パラメータを template 処理
    _.forEach(params, function(value, name) {
        //if (value.indexOf('{{=') == -1) {
        //    return;
        //}
        params[name] = replaceText(value, variables);
        // TODO: システム変数（$currentProject）の処理
        // TODO: $currentProject は root から現在の stack top への相対
    });

    return params;
}

// filePaths: 含まれるすべてのファイルのパス
function preProcess_Recurse(filePathAbs, currentProjectDirectoryFromRoot, defines, templateVariables) {
    // 上書きする（階層が深い方を優先）
    templateVariables = _.assign(templateVariables, { $currentProjectDirectory: currentProjectDirectoryFromRoot });

    stream.Type = adTypeText;
    // UTF-8 BOM なし 専用
    stream.charset = "UTF-8";
    stream.Open();
    stream.LoadFromFile(filePathAbs);
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
            filePath: filePathAbs
        };
        lines.push(lineObj);
    });

    lines = parseOneLineComment(lines);
    lines = parseMultilineComment(lines);
    lines = preProcessConditionalCompile(lines, defines);

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

        var includeMatch = line.match(/^<<\[\s*(.+)\s*\]\s*(\((.+)?\))?$/);

        if (includeMatch) {
            var includeFileString = includeMatch[1];
            var includeOptionString = includeMatch[3];

            try {
                var includeFileInfo = parseIncludeFilePath(includeFileString, currentProjectDirectoryFromRoot, filePathAbs, templateVariables);
            }
            catch (e) {
                throw new ParseError(e.errorMessage, lineObj); 
            }

            try {
                var includeParam = parseIncludeParameters(includeOptionString, templateVariables);
            }
            catch(e) {
                var errorMessage = "include パラメータが不正です。";
                throw new ParseError(errorMessage, lineObj); 
            }
            //printJSON(includeParam);

            // include ファイルに渡す用
            // 上書きする（階層が深い方を優先）
            var localTemplateVariables = _.assign(templateVariables, includeParam);

            localTemplateVariables.$currentProjectDirectory = currentProjectDirectoryFromRoot;
            //printJSON(localTemplateVariables);

            var includeProjectDirectoryFromRoot = includeFileInfo.projectDirectory;

            var path = includeFileInfo.filePath;
            //printJSON(includeFileInfo);

            if (!fso.FileExists(path)) {
                var relativeProjectPath = CL.getRelativePath(conf.$rootDirectory, includeFileInfo.sourceDirectory);
                var relativePath = CL.getRelativePath(includeFileInfo.sourceDirectory, path);
                //var relativePath = includeFileString;

                var errorMessage = "フォルダ\n" + relativeProjectPath + "\nには\nファイル\n" + relativePath + "\nが存在しません";
                throw new ParseError(errorMessage, lineObj);
            }

            var includeLines = preProcess_Recurse(path, includeProjectDirectoryFromRoot, defines, localTemplateVariables);
            
            dstLines = dstLines.concat(includeLines);

            continue;
        }
        else {
            dstLines.push(lineObj);
        }
    }

    return dstLines;
}

// preprocess
// include とかコメント削除とか
// 入れ子の include にも対応
function preProcess(filePathAbs) {
    var defines = {};

    // メインソースファイルのフォルダを現在のプロジェクトフォルダとする
    var entryProject = fso.GetParentFolderName(filePathAbs);
    var projectDirectoryFromRoot = CL.getRelativePath(conf.$rootDirectory, entryProject);

    // TODO: conf.yaml とかで global な変数を指定できるように
    var templateVariables = { };

    try {
        return preProcess_Recurse(filePathAbs, projectDirectoryFromRoot, defines, templateVariables);
    }
    catch (e) {
        if (_.isUndefined(e.lineObj) || _.isUndefined(e.errorMessage)){
            throw e;
        }
        parseError(e);
    }
}

