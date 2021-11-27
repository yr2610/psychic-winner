// preprocess
// include とかコメント削除とか
// 入れ子の include にも対応
function preProcess(filePath, filePaths) {
    var lines = [];
    var pathStack = [];

    pathStack.push({
        includePath: null
    });
    preProcess_Recurse(filePath, lines, filePaths, pathStack);

    return lines;
}

function preProcess_Recurse(filePath, lines, filePaths, pathStack) {
    filePaths.push(filePath);
    var parentFolderName = fso.GetParentFolderName(filePath);

    _.last(pathStack).parentFolder = parentFolderName;

    /**/
    stream.Type = adTypeText;
    // XXX: charset の _autodetect_all が判別に失敗することがある問題がどうにもならないので、超適当に判別
    // 一旦 Shift JIS としてロード
 //   stream.charset = "Shift_JIS";
    // UTF-8 BOM なし 専用
    stream.charset = "UTF-8";
    stream.Open();
    stream.LoadFromFile(filePath);
    var allLines = stream.ReadText(adReadAll);
    stream.Close();

//    // 先頭の文字が Unicode っぽければロードしなおす
//    {
//        var charCode0 = allLines.charCodeAt(0);
//        var charCode1 = allLines.charCodeAt(1);
//        // UTF8 with BOM
//        if (charCode0 === 0x30fb && charCode1 === 0xff7f)
//        {
//            stream.charset = "UTF-8";
//        }
//        else
//        // UTF-16LE, BE
//        if (charCode0 === 0xf8f3 && charCode1 === 0xf8f2 ||
//            charCode0 === 0xf8f2 && charCode1 === 0xf8f3)
//        {
//            stream.charset = "UTF-16";
//        }
//        if (stream.charset !== "Shift_JIS")
//        {
//            stream.Open();
//            stream.LoadFromFile(filePath);
//            allLines = stream.ReadText(adReadAll);
//            stream.Close();
//        }
//    }
    /*/
    // ファイルを読み取り専用で開く
    // とりあえず Unicode(UTF-16)専用
    var file = fso.OpenTextFile(filePath, FORREADING, true, TRISTATE_TRUE);
    var allLines = file.ReadAll();
    file.Close();
    /**/
    //var lineArray = new ArrayReader(allLines.split(/\r\n|\r|\n/));
    // 空要素も結果に含めたいのでsplitには正規表現を使わないように
    var lineArray = new ArrayReader(allLines.replace(/\r\n|\r/g, "\n").split("\n"));

    //var path = fso.BuildPath(parentFolderName, image);

    while (!lineArray.atEnd) {
        var line = lineArray.read();

        if (line === "") {
            continue;
        }

        var cppCommentIndex = line.search(/\/{2,}/);
        if (cppCommentIndex === 0) {
            continue;
        }
        if (cppCommentIndex >= 1) {
            line = line.slice(0, cppCommentIndex);
        }

        // それぞれ行頭、行末に書かれた <!-- と --> のみ対応
        // 入れ子に対応
        function parseMultilineCommments(beginRe, endRe)
        {
            if (!beginRe.test(line))
            {
                return false;
            }

            var beginLineNum = lineArray.index;

            for (var commentDepth = 0;;)
            {
                if (beginRe.test(line))
                {
                    commentDepth++;
                }
                if (endRe.test(line))
                {
                    commentDepth--;
                }
                if (commentDepth == 0)
                {
                    break;
                }
                if (lineArray.atEnd)
                {
                    if (commentDepth > 0)
                    {
                        var errorMessage = "コメントが閉じていません";
                        Error(errorMessage, filePath, beginLineNum);
                    }
                    break;
                }
                line = lineArray.read();
            }
            return true;
        }
        if (parseMultilineCommments(/^\s*<!--.*/, /.*-->\s*$/))
        {
            continue;
        }
        if (parseMultilineCommments(/^\s*\/\*.*/, /.*\*\/\s*$/))
        {
            continue;
        }

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

// #define とか #if else endif 的なの
// include とコメント削除が適用済みのを渡す
function preProcessConditionalCompile(lines) {
    var srcLines = new ArrayReader(lines);
    var dstLines = [];
    var objs = {};
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
