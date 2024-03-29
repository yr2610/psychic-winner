﻿// file の ReadLine(), AtEndOfStream 風の仕様で配列にアクセスするための機構を用意
function ArrayReader(array) { this.__a = array; this.index = 0; this.atEnd = false; }
ArrayReader.prototype.read = function(o) { if (this.atEnd) return null; if (this.index + 1 >= this.__a.length) this.atEnd = true; return this.__a[this.index++]; }

// すべての ID を割り当て直す
var fResetId = false;

var runInCScript = (function() {
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );

    return (fso.getBaseName(WScript.FullName).toLowerCase() === "cscript");
})();

function alert(s) {
    WScript.Echo(s);
}

function printJSON(json) {
    alert(JSON.stringify(json, undefined, 4));
}

function $templateObject(object, data) {
    var json = JSON.stringify(object);
    function replacer(m, k) {
        return data[k];
    }
    json = json.replace(/\{\{([^\}]+)\}\}/g, replacer);

    return JSON.parse(json);
}

function makeLineinfoString(filePath, lineNum) {
    var s = "";

    // ファイル名がない時点で終了
    if (typeof filePath === 'undefined') {
        return s;
    }

    s += "\nファイル:\t" + filePath;

    if (typeof lineNum === 'undefined') {
        return s;
    }

    s += "\n行:\t" + lineNum;

    return s;
}

function getRelativePath(filePath, rootFilePath, fso) {
    if (typeof fso === "undefined") {
        fso = new ActiveXObject( "Scripting.FileSystemObject" );
    }

    var rootFileFolderName = fso.GetParentFolderName(rootFilePath);

    if (!_.startsWith(filePath, rootFileFolderName)) {
        return null;
    }

    return filePath.slice(rootFileFolderName.length + 1);
}

var ParseError = function(errorMessage, lineObj) {
    this.errorMessage = errorMessage;
    this.lineObj = lineObj;
};

// ParseError が引数
function parseError(e) {
    if (_.isUndefined(e.lineObj)) {
        MyError(e.errorMessage);
    }
    else {
        var lineObj = e.lineObj;
        MyError(e.errorMessage, lineObj.filePath, lineObj.lineNum);
    }
}

function MyError(message, filePath, lineNum) {
    if (typeof filePath !== "undefined") {
        var relativeFilePath = getRelativePath(filePath, rootFilePath, fso);
        if (relativeFilePath) {
            filePath = relativeFilePath;
        }

        message += "\n" + makeLineinfoString(filePath, lineNum);
    }

    if (runInCScript) {
        WScript.StdErr.Write(message);
    } else {
        shell.Popup(message, 0, "エラー", ICON_EXCLA);
    }
    WScript.Quit(1);
}

function createRandomId(len, random) {
    if (_.isUndefined(random)) {
        random = Math.random;
    } 

    var c = "abcdefghijklmnopqrstuvwxyz";
    c += c.toUpperCase();
    // 1文字目はアルファベットのみ
    var s = c.charAt(Math.floor(random() * c.length));
    c += "0123456789";
    var cl = c.length;

    for (var i = 1; i < len; i++) {
        s += c.charAt(Math.floor(random() * cl));
    }

    return s;
}

// key が id のリストを渡すと、それと重複しないものを返す
// lenを1、idListに36通りすべてを渡せば簡単に無限ループになるけど、特にそのあたりのチェックとかはしない
function createUid(len, idList) {
    if (_.isUndefined(idList)) {
        idList = {};
    }
    var s = createRandomId(len);

    // idList に存在するものの間は無限ループ
    while (s in idList) {
        s = createRandomId(len);
    }

    return s;
}

if (typeof(global) === 'undefined') {
    global = Function('return this')();
}

var shell = new ActiveXObject("WScript.Shell");
var shellApplication = new ActiveXObject("Shell.Application");
var fso = new ActiveXObject( "Scripting.FileSystemObject" );
var stream = new ActiveXObject("ADODB.Stream");

if (( WScript.Arguments.length != 1 ) ||
    ( WScript.Arguments.Unnamed(0) == "")) {
    MyError("チェックリストのソースファイル（.txt）をドラッグ＆ドロップしてください。");
}

var filePath = WScript.Arguments.Unnamed(0);

if (fso.GetExtensionName(filePath) != "txt") {
    MyError(".txt ファイルをドラッグ＆ドロップしてください。");
}

// Performance を取得
var htmlfile = WSH.CreateObject("htmlfile");
htmlfile.write('<meta http-equiv="x-ua-compatible" content="IE=Edge"/>');
var performance = htmlfile.parentWindow.performance;
htmlfile.close();

// プロジェクトフォルダ内のソース置き場
var sourceDirectoryName = "source";

// バックアップ置き場
var backupDirectoryName = "bak";

// 中間生成ファイル置き場
var intermediateDirectoryName = "intermediate";

var includePath = [];

// メインソースファイルのrootフォルダはデフォルトで最優先で探す
includePath.push(fso.GetParentFolderName(filePath));

// グローバルな設定
// 現状 includePath のみ
// FIXME: 廃止予定
(function(){
    var confFilePath = "conf.yml";
    confFilePath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), confFilePath);
    if (!fso.FileExists(confFilePath)) {
        return;
    }
    var data = CL.readYAMLFile(confFilePath);

    // include文法のパス用
    if (!_.isUndefined(data.includePath)) {
        includePath = includePath.concat(data.includePath);
    }
})();

var confFileName = "conf.yml";
(function(){
    var baseName = fso.GetBaseName(filePath);
    baseName = baseName.replace(/_index$/, "");
    if (baseName != "index") {
        confFileName = baseName + "_" + confFileName;
    }
})();
var conf = readConfigFile(confFileName);

var entryFilePath = filePath;
var entryProject = fso.GetParentFolderName(entryFilePath);
var entryProjectFromRoot = CL.getRelativePath(conf.$rootDirectory, entryProject);

// XXX: entry source からの相対パスを root からの絶対パスに変換
// XXX: 名前が機能を十分に説明してないけど、基本的に source 以下のファイル以外を変換するケースはないと思うので…
function $abspath(path) {
    var entrySourceDirectoryFromRoot = fso.BuildPath(entryProjectFromRoot, sourceDirectoryName);
    return "/" + fso.BuildPath(entrySourceDirectoryFromRoot, path);
}

var allFilePaths = [];

var rootFilePath = filePath;
var srcLines = preProcess(filePath);
srcLines = new ArrayReader(srcLines);

var stack = new Stack();

var kindH = "H";
var kindUL = "UL";

var root = {
    sourceFiles: null, // 場所確保。txt2json 自身がソースファイルを更新するので、最後に取得
    kind: kindH,
    level: 0,
    id: null,   // 場所確保。sourceFiles の last modified date を基に生成した id を埋め込む
    text: "",
    variables: {},
    children: [],

    // 以下はJSON出力前に削除する
    // UID重複チェック用
    // 「複数人で１つのファイルを作成（ID自動生成）してマージしたら衝突」は考慮しなくて良いぐらいの確率だけど、「IDごとコピペして重複」が高頻度で発生する恐れがあるので
    uidList: {}

};
stack.push(root);

// conf から機能を持った変数を移行
(function(){
    if (!_.isUndefined(conf.$templateValues)) {
        _.assign(root.variables, conf.$templateValues);
    }

    // XXX: 無視の方を指定する方が良いか
    var variableList = [
        "templateFilename",
        "ignoreColumnId",
        "outputFilename",
        "projectId",
        "indexSheetname",
        "rootDirectory"
    ];

    _.forEach(variableList, function(key) {
        var value = conf["$" + key];
        if (_.isUndefined(value)) {
            return;
        }
        // XXX: project は render 処理でしか使ってないけど、修正が面倒なのでここで対処
        if (key == "projectId") {
            key = "project";
        }
        root.variables[key] = value;
    });

    if (!_.isUndefined(root.variables.rootDirectory)) {
        // 相対パスに変換
        var basePath = fso.GetParentFolderName(filePath);
        var absolutePath = root.variables.rootDirectory;
        var relativePath = CL.getRelativePath(basePath, absolutePath);

        root.variables.rootDirectory = relativePath;
    }

    if (!_.isUndefined(conf.$input)) {
        var data = conf.$input;

        // 入力欄の順序の宣言
        if (!_.isUndefined(data.order)) {
            stack.peek().columnNames = data.order;
        }
        // デフォルト値
        if (!_.isUndefined(data.defaultValues)) {
            stack.peek().defaultColumnValues = data.defaultValues;
        }

        // 条件付きデフォルト値
        if (!_.isUndefined(data.rules)) {
            var conditionalColumnValues = stack.peek().conditionalColumnValues;
            if (_.isUndefined(conditionalColumnValues)) {
                conditionalColumnValues = [];
            }
    
            _.forEach(data.rules, function(rule) {
                conditionalColumnValues.push({
                    re: new RegExp(rule.condition),
                    columnValues: rule.values
                });
            });
            stack.peek().conditionalColumnValues = conditionalColumnValues;
        }
    }

})();

// root local な project path から project path local なパスを取得
function getProjectLocalPath(filePath, projectDirectory) {
    var projectDirectoryAbs = fso.BuildPath(conf.$rootDirectory, projectDirectory);

    return CL.getRelativePath(projectDirectoryAbs, filePath);
}

// IDがふられてないノード
var noIdNodes = {};

// tree 構築後じゃないと leaf かどうかの判別ができないのと、入力済の ID 間での重複チェックをしたいので、貯めといて最後に ID を割り当てる
function AddNoIdNode(node, lineObj, lineNum, newSrcText) {
    var filePath = lineObj.filePath;
    var projectDirectory = lineObj.projectDirectory;
    var key = projectDirectory + ":" + filePath;

    if (!(key in noIdNodes)) {
        noIdNodes[key] = [];
    }

    var data = {
        node: node,

        lineObj: lineObj,

        lineNum: lineNum,

        // 書き換え後の文字列
        // 文字列は {uid} を含むもの
        // 後で uid を生成して{uid}の位置に埋め込む
        newSrcText: newSrcText
    };

    noIdNodes[key].push(data);
}

var srcTextsToRewrite = {};

function AddSrcTextToRewrite(noIdLineData, newSrcText) {
    var lineObj = noIdLineData.lineObj;
    var filePath = lineObj.filePath;
    var projectDirectory = lineObj.projectDirectory;
    //var filePathProjectLocal = getProjectLocalPath(filePath, projectDirectory);
    var key = projectDirectory + ":" + filePath;
    var lineNum = noIdLineData.lineNum;

    if (!(key in srcTextsToRewrite)) {
        srcTextsToRewrite[key] = {
            filePath: filePath,
            projectDirectory: projectDirectory,
            newTexts: {}
        };
    }

    var newTexts = srcTextsToRewrite[key].newTexts;
    newTexts[lineNum - 1] = newSrcText;
}

function AddChildNode(parent, child) {
    parent.children.push(child);
    child.parent = parent;
}

// 一番近い親を返す
// 自分が存在する前に使いたい都合上、parentとなるnodeを渡す（渡したnodeも検索対象）仕様
function FindParentNode(parent, fun) {
    for (; parent; parent = parent.parent) {
        if (fun(parent)) {
            return parent;
        }
    }
    return null;
}

// 一番近い親の uidList を返す
function FindUidList(parent) {
    var node = FindParentNode(parent, function(node) {
        return node.uidList;
    });

    return node ? node.uidList : null;
}

// tableHeaders 内の ID で最小のものが一番左として連番で検索
function getDataFromTableRow(srcData, parentNode, tableHeaderIds) {
    // data を h1 の tableHeaders の番号に合わせて作り直す
    var data = [];

    // H1は確実に見つかるものとしてOK
    var h1Node = FindParentNode(parentNode, function(node) {
        return (node.kind === kindH && node.level === 1);
    });

    if (typeof tableHeaderIds === 'undefined') {
        /**
        // tableHeaders 内の ID で最小のものが一番左として連番の値
        var minNumber = Infinity;
        for (var i = 0; i < h1Node.tableHeaders.length; i++)
        {
            minNumber = Math.min(minNumber, h1Node.tableHeaders[i].id);
        }
        /*/
        // 「必ず 1 から始まる連番」の方が仕様として素直ですっきりしているか
        var minNumber = 1;
        /**/

        tableHeaderIds = [];
        for (var i = 0; i < srcData.length; i++) {
            tableHeaderIds[i] = minNumber + i;
        }
    }
    for (var i = 0; i < srcData.length; i++)
    {
        if (!srcData[i])
        {
            continue;
        }
        var number = tableHeaderIds[i];
        if (typeof number === 'undefined')
        {
            return null;
        }
        var headerIndex = h1Node.tableHeaders.findIndex(function(element, index, array)
        {
            return (element.id === number);
        });

        if (headerIndex === -1)
        {
            return null;
        }
        
        data[headerIndex] = srcData[i];
    }

    return data;
}

function parseComment(text, lineObj) {
    var projectDirectoryFromRoot = lineObj.projectDirectory;
    var fileParentFolderAbs = sourceLocalPathToAbsolutePath(fso.GetParentFolderName(lineObj.filePath), projectDirectoryFromRoot);
    // 複数行テキストに対応するために .+ じゃなくて [\s\S]+
    var re = /^([\s\S]+)\s+\[\^(.+)\]$/;
    var commentMatch = text.trim().match(re);

    if (!commentMatch) {
        return null;
    }

    var text = commentMatch[1].trim();
    var comment = commentMatch[2].trim();

    comment = comment.replace(/<br>/gi, "\n");
    comment = comment.replace(/\\n/gi, "\n");

    var imageMatch = comment.match(/^\!(.+)\!$/);

    if (!imageMatch) {
        return {
            text: text,
            comment: comment
        };
    }

    var imageFilePath = imageMatch[1];
    
    return {
        text: text,
        imageFilePath: imageFilePath
    };
}

function parseHeading(lineObj) {
    var line = lineObj.line;
    var h = line.match(/^(#+)\s+(.*)$/);

    if (!h) {
        return null;
    }

    var level = h[1].length;
    var text = h[2];

    while (stack.peek().kind != kindH || stack.peek().level >= level) {
        stack.pop();
    }

    var uid = undefined;
    var uidList = undefined;
    var tableHeaders = undefined;
    var url = undefined;
    if (level === 1) {
        var uidMatch = text.match(/^\[#([\w\-]+)\]\s*(.+)$/);
        if (fResetId) {
            uidMatch = null;
        }
        if (uidMatch) {
            uid = uidMatch[1];
            text = uidMatch[2];

            var uidListH1 = FindUidList(stack.peek());
            if (uid in uidListH1) {
                (function(){
                    var uidInfo0 = uidListH1[uid];
                    var errorMessage = "ID '#" + uid + "' が重複しています";
                    errorMessage += makeLineinfoString(uidInfo0.filePath, uidInfo0.lineNum);
                    errorMessage += makeLineinfoString(lineObj.filePath, lineObj.lineNum);

                    throw new ParseError(errorMessage);
                })();
            }
            else {
                uidListH1[uid] = lineObj;
            }
        }

        // シート内での重複だけ確認したいのでここでクリア
        uidList = {};

        tableHeaders = [];
    }
    else {
        while (/.*\s\+\s*$/.test(text)) {
            // 改行の次の行の行頭のスペースは無視するように
            // 厳密にはインデントが揃ってるかちゃんとみるべきだけど、そこまでやるつもりはない
            line = _.trimLeft(srcLines.read().line);
            text = _.trimRight(_.trimRight(text).slice(0, -1));
            text += "\n" + line;
        }

        // １行のみ、行全体以外は対応しない
        var link = text.trim().match(/^\[(.+)\]\((.+)\)$/);
        if (link) {
            text = link[1].trim();
            url = link[2].trim();
        }
    }

    text = text.trim();

    if (text.length > 31) {
        var errorMessage = "シート名が31文字を超えています";

        throw new ParseError(errorMessage, lineObj);
    }
    if (/[\:\\\?\[\]\/\*：￥？［］／＊]/g.test(text)) {
        var errorMessage = "シート名に使用できない文字が含まれています"
        + "\n\nシート名には全角、半角ともに次の文字は使えません"
        + "\n1 ）コロン        ："
        + "\n2 ）円記号        ￥"
        + "\n3 ）疑問符        ？"
        + "\n4 ）角カッコ      [ ]"
        + "\n5 ）スラッシュ     /"
        + "\n6 ）アスタリスク  ＊";

        throw new ParseError(errorMessage, lineObj);
    }
    if (_.find(root.children, function(item) {
        return item.text == text;
    })) {
        var errorMessage = "シート名「" + text + "」はすでに使われています";

        throw new ParseError(errorMessage, lineObj);
    }

    var item = {
        kind: kindH,
        level: level,
        id: uid,
        text: text,
        tableHeaders: tableHeaders,
        tableHeadersNonInputArea: [],
        url: url,
        variables: {},
        children: [],

        // 以下はJSON出力前に削除する
        // UID重複チェック用
        // 「複数人で１つのファイルを作成（ID自動生成）してマージしたら衝突」は考慮しなくて良いぐらいの確率だけど、「IDごとコピペして重複」が高頻度で発生する恐れがあるので
        uidList: uidList,
        lineObj: lineObj
    };
    AddChildNode(stack.peek(), item);
    stack.push(item);
    //WScript.Echo(item.level + "\n" + item.text);

    if (fResetId ||
        level === 1 && !uid) {
        // tree構築後にIDをふる
        var newSrcText = lineObj.line;
        var match = newSrcText.match(/^(#+)(?: \[#[\w\-]+\])?(.*)$/);

        // ID 挿入して書き換え
        newSrcText = match[1] + " [#{uid}]" + match[2];

        if (!_.isUndefined(lineObj.comment)) {
            newSrcText += lineObj.comment;
        }

        AddNoIdNode(item, lineObj, lineObj.lineNum, newSrcText);
    }

    return h;
}

function parseUnorderedList(lineObj) {
    // 行頭に全角スペース、タブがないかのチェック
    (function () {
        var fullwidthSpaceMatch = line.match(/^([\s　]+).*$/);
        if (!fullwidthSpaceMatch) {
            return;
        }
        var regex = /[　\t]/g;
        if (regex.test(fullwidthSpaceMatch[1])) {
            var errorMessage = "行頭に全角スペースもしくはタブ文字が含まれています";
            MyError(errorMessage, lineObj.filePath, lineObj.lineNum);
        }
    })();
    // # とか - とか 1. の後ろにスペースがないかのチェック
    function checkSpaceAfterMark(re) {
        var spaceMatch = line.match(re);
        if (!spaceMatch) {
            return;
        }
        var regex = /^\s+/;
        if (!regex.test(spaceMatch[1])) {
            var errorMessage = "行頭の記号の後ろに半角スペースが必要です";

            throw new ParseError(errorMessage, lineObj);
        }
    }
    checkSpaceAfterMark(/^#+(.+)$/);
    checkSpaceAfterMark(/^\s*[\*\+\-]\.?(.+)$/);
    checkSpaceAfterMark(/^\s*\d+\.(.+)$/);

    var ul = line.match(/^(\s*)([\*\+\-])\s+(.*)$/);

    if (!ul) {
        return null;
    }

    var indent = ul[1].length;
    var text = ul[3];
    var marker = ul[2];

    while (stack.peek().kind == kindUL && stack.peek().indent >= indent) {
        stack.pop();
    }
    if (stack.peek().kind != kindUL && indent > 0) {
        var errorMessage = "一番上の階層のノードがインデントされています";

        throw new ParseError(errorMessage, lineObj);
    }
    if (stack.peek().kind == kindUL && (indent - stack.peek().indent < 2)) {
        var errorMessage = "インデントはスペース 2 個以上必要です";

        throw new ParseError(errorMessage, lineObj);
    }

    var uidMatch = text.match(/^\[#([\w\-]+)\]\s+(.+)$/);
    var uid = undefined;
    if (uidMatch) {
        uid = uidMatch[1];
        text = uidMatch[2];
        {(function(){
            var uidList = FindUidList(stack.peek());
            if (uid in uidList) {
                var uidInfo0 = uidList[uid];
                var errorMessage = "ID '#" + uid + "' が重複しています";
                errorMessage += makeLineinfoString(uidInfo0.filePath, uidInfo0.lineNum);
                errorMessage += makeLineinfoString(lineObj.filePath, lineObj.lineNum);

                throw new ParseError(errorMessage);
            }
            else {
                uidList[uid] = lineObj;
            }
        })();}
    }

    var attributes = void 0;
    while (true) {
        var attributeMatch = text.match(/^\s*\<([A-Za-z_]\w*)\>\(([^\)]+)\)\s*(.+)$/);
        if (attributeMatch === null) {
            break;
        }
        var name = attributeMatch[1];
        var value = attributeMatch[2];
        text = attributeMatch[3];

        if (_.isUndefined(attributes)) {
            attributes = {};
        }
        attributes[name] = value;
    }


    // TODO: leaf 以外で initialValues が設定されていたら削除しておく?
    var initialValues = void 0;

    // 旧仕様も一応残しておく
    while (true) {
        var initialValueMatch = text.match(/^\s*\[#([A-Za-z_]\w*)\]\(([^\)]+)\)\s*(.+)$/);
        if (initialValueMatch === null) {
            break;
        }
        var name = initialValueMatch[1];
        var value = initialValueMatch[2];
        text = initialValueMatch[3];

        if (_.isUndefined(initialValues)) {
            initialValues = {};
        }
        initialValues[name] = value;
    }

    function getLowestColumnNames() {
        for (var i = stack.__a.length - 1; i >= 0; i--) {
            var elem = stack.__a[i];
            if (_.isUndefined(elem.columnNames)) {
                continue;
            }
            return {
                columnNames: elem.columnNames,
                defaultColumnValues: elem.defaultColumnValues
            };
        }
        return null;
    }

    // (foo: 0, bar: "baz") 形式の初期値設定
    (function() {
        var parse = parseColumnValues(text, true);
        if (parse === null) {
            return;
        }

        text = parse.remain;

        if (_.isUndefined(initialValues)) {
            initialValues = {};
        }

        var columnNames;
        var lowestColumnNames = getLowestColumnNames();
        if (lowestColumnNames !== null) {
            columnNames = lowestColumnNames.columnNames;
        }

        parse.columnValues.forEach(function(param, index) {
            var value = _.isUndefined(param.value) ? null : param.value;
            if (!_.isUndefined(param.key)) {
                initialValues[param.key] = value;
                return;
            }
            if (_.isUndefined(columnNames)) {
                var errorMessage = "列名リストが宣言されていません。";
                throw new ParseError(errorMessage, lineObj);
            }
            if (index >= columnNames.length) {
                var errorMessage = "列の初期値が列名リストの範囲外に指定されています。";
                throw new ParseError(errorMessage, lineObj);
            }
            var key = columnNames[index];
            initialValues[key] = value;
        });
        return;

        var match = text.match(/^\s*\(([^\)]+)\)\s*(.+)$/);
        if (match === null) {
            // TODO: デフォルト値が設定されていれば指定がなくてもセット
            return;
        }
        text = match[2];
        initialValues = {};
        var params = match[1].split(',');
        var columnNameIndex = 0;
        params.forEach(function(param) {
            param = _.trim(param);
            var nameValueMatch = param.match(/^([A-Za-z_]\w*)\s*:\s(.+)$/);
            if (nameValueMatch) {
                var name = nameValueMatch[1];
                var value = nameValueMatch[2];
                initialValues[name] = value;
                return;
            }
            // TODO: stackから上にさかのぼって columnNames を見つける
            var columnNames = [];
            if (columnNameIndex >= columnNames.length) {
                // TODO: 範囲外エラー
            }
            initialValues[columnNames[columnNameIndex]] = param;
            // TODO: ダブルクォーテーション対応させる
            columnNameIndex++;
        });
    })();


    while (/.*\s\+\s*$/.test(text)) {
        // 改行の次の行の行頭のスペースは無視するように
        // 厳密にはインデントが揃ってるかちゃんとみるべきだけど、そこまでやるつもりはない
        line = _.trimLeft(srcLines.read().line);
        text = _.trimRight(_.trimRight(text).slice(0, -1));
        text += "\n" + line;
    }

    var commentResult = parseComment(text, lineObj);
    var comment;
    var imageFilePath;
    if (commentResult) {
        text = commentResult.text;
        comment = commentResult.comment;
        imageFilePath = commentResult.imageFilePath;
        //var v = {
        //    text: commentResult.text,
        //    comment: commentResult.comment,
        //    imageFilePath: commentResult.imageFilePath
        //};
        //printJSON(v);
    }

    //var commentMatch = text.trim().match(/^([\s\S]+)\s*\[\^(.+)\]$/);
    //var comment = undefined;
    //if (commentMatch) {
    //    text = commentMatch[1].trim();
    //    comment = commentMatch[2].trim();
    //    comment = comment.replace(/<br>/gi, "\n");
    //}

    // table 形式でデータを記述できるように
    var td = text.match(/^([^\|]+)\|(.*)\|$/);
    var data = undefined;
    if (td) {
        // TODO: 画像対応
        text = td[1].trim();
        data = td[2].split("|");
        for (var i = 0; i < data.length; i++) {
            data[i] = data[i].trim();
        }

        data = getDataFromTableRow(data, stack.peek());

        if (!data) {
            var errorMessage = "シートに該当IDの確認欄がありません";
            throw new ParseError(errorMessage, lineObj);
        }
    }

    // １行のみ、行全体以外は対応しない
    var link = text.trim().match(/^\[(.+)\]\((.+)\)$/);
    var url = undefined;
    if (link) {
        text = link[1].trim();
        url = link[2].trim();
    }

    text = text.trim();

    if (/\t/.test(text)) {
        var errorMessage = "テキストにタブ文字が含まれています";
        throw new ParseError(errorMessage, lineObj);
    }
    if (/\t/.test(comment)) {
        var errorMessage = "コメントにタブ文字が含まれています";
        throw new ParseError(errorMessage, lineObj);
    }

    var item = {
        kind: kindUL,
        indent: indent,
        marker: marker,
        group: -1,   // 場所確保のため一旦追加
        depthInGroup: -1,   // 場所確保のため一旦追加
        id: uid,
        text: text,
        tableData: data,
        comment: comment,
        imageFilePath: imageFilePath,
        initialValues: initialValues,
        attributes: attributes,
        url: url,
        variables: {},
        children: [],

        // 以下はJSON出力前に削除する
        lineObj: lineObj
    };

    AddChildNode(stack.peek(), item);
    stack.push(item);
    //WScript.Echo(ul.length + "\n" + line);

    if (!uidMatch || fResetId) {
        // tree構築後にleafだったらIDをふる
        var newSrcText = lineObj.line;
        var match = newSrcText.match(/^(\s*[\*\+\-])(?: \[#[\w\-]+\]\s+)?(.*)$/);

        // ID 挿入して書き換え
        newSrcText = match[1] + " [#{uid}]" + match[2];

        if (!_.isUndefined(lineObj.comment)) {
            newSrcText += lineObj.comment;
        }

        AddNoIdNode(item, lineObj, lineObj.lineNum, newSrcText);
    }

    return ul;
}

//  ファイルの文字データを一行ずつ読む
while (!srcLines.atEnd) {
    var lineObj = srcLines.read();
    var line = lineObj.line;

    try {
        if (parseHeading(lineObj)) {
            continue;
        }
        if (parseUnorderedList(lineObj)) {
            continue;
        }
    }
    catch (e) {
        parseError(e);
    }

    // "*.", "-.", "+." はチェック項目列の見出しとする
    var headerList = line.match(/^(?:\s*)([\*\+\-])\.\s+(.*)\s*$/);
    if (headerList) {
        //var level = headerList[1].length;
        var marker = headerList[1];
        var text = headerList[2];
        var parent = stack.peek();

        // XXX: 現状は H1 の直下専用
        if (parent.kind === kindH && parent.level === 1) {
            var comment = undefined;
            var commentMatch = text.match(/^(.+)\s*\[\^(.+)\]$/);
            if (commentMatch) {
                text = commentMatch[1].trim();
                comment = commentMatch[2].trim();
                comment = comment.replace(/<br>/gi, "\n");
            }

            var headers = parent.tableHeadersNonInputArea;
            var prevName = (headers.length >= 1) ? headers[headers.length - 1].name : undefined;

            if (text === "" || (prevName && text === prevName)) {
                headers[headers.length - 1].size++;
            }
            else {
                var group = 0;
                if (headers.length >= 1) {
                    group = headers[headers.length - 1].group;
                    if (headers[headers.length - 1].marker !== marker) {
                        group++;
                    }
                }

                var item = {
                    marker: marker,
                    group: group,
                    name: text,
                    comment: comment,
                    size: 1
                };

                headers.push(item);
            }
        }

        continue;
    }


    // 数字は unique ID として扱う
    var ol = line.match(/^\s*(\d+)\.\s+(.*)$/);
    if (ol) {
        (function(){
            var number = parseInt(ol[1], 10);
            var text = ol[2];
            var parent = stack.peek();

            while (/.*\s\+\s*$/.test(text)) {
                // 改行の次の行の行頭のスペースは無視するように
                // 厳密にはインデントが揃ってるかちゃんとみるべきだけど、そこまでやるつもりはない
                line = _.trimLeft(srcLines.read().line);
                text = _.trimRight(_.trimRight(text).slice(0, -1));
                text += "\n" + line;
            }

            if (parent.kind === kindH && parent.level === 1) {
                var comment = undefined;
                var commentMatch = text.trim().match(/^([\s\S]+)\s*\[\^(.+)\]$/);
                if (commentMatch) {
                    text = commentMatch[1].trim();
                    comment = commentMatch[2].trim();
                    if (/<br>/gi.test(comment)) {
                        var errorMessage = "確認欄のコメントでは改行は使えません";
                        MyError(errorMessage, lineObj.filePath, lineObj.lineNum);
                    }
                    // Excel の仕様で、入力時メッセージのタイトルは31文字まで
                    if (comment.length > 32) {
                        var errorMessage = "確認欄のコメントが32文字を超えています";
                        MyError(errorMessage, lineObj.filePath, lineObj.lineNum);
                    }
                }
                
                var item = {
                    name: text,
                    description: comment,
                    id: number
                };

                var headerIndex = parent.tableHeaders.findIndex(function(element, index, array) {
                    return (element.id === number);
                });
                // すでに同じIDの確認欄が存在
                if (headerIndex !== -1) {
                    var errorMessage = "確認欄のID(" + number + ")が重複しています";
                    MyError(errorMessage, lineObj.filePath, lineObj.lineNum);
                }
                else {
                    parent.tableHeaders.push(item);
                }

                return;
            }
            if (parent.kind === kindUL) {
                // H1は確実に見つかるものとしてOK
                var h1Node = FindParentNode(parent, function(node) {
                    return (node.kind === kindH && node.level === 1);
                });

                if (!parent.tableData) {
                    parent.tableData = [];
                }

                var headerIndex = h1Node.tableHeaders.findIndex(function(element, index, array) {
                    return (element.id === number);
                });

                if (headerIndex === -1) {
                    var errorMessage = "シートにID" + number + "の確認欄がありません";
                    MyError(errorMessage, lineObj.filePath, lineObj.lineNum);
                }

                parent.tableData[headerIndex] = text;

                return;
            }
        })();

        continue;
    }
    

    // ol も node にしておこうと思ったけど、leaf は必ず ul であることが前提の作りになっているので、諦める
    var th = line.match(/^\s*\|(.*)\|\s*$/);
    if (th) {
        var parent = stack.peek();
        if (parent.kind != kindH || parent.level != 1) {
            MyError("番号付きリストは H1 の直下以外には作れません");
        }
        parent.tableHeaders = [];
        th = th[1].split("|");
        for (var i = 0; i < th.length; i++) {
            var s = th[i].trim();
            var name_description = s.match(/^\[(.+)\]\(\s*\"(.+)\"\s*\)$/);
            var item = {};
            if (name_description) {
                item.name = name_description[1];
                item.description = name_description[2];
            }
            else {
                item.name = s;
            }
            item.id = i + 1;
            parent.tableHeaders.push(item);
        }
        continue;
    }

    if (/^\s*```yaml\s*$/.test(line)) {
        var topLineObj = lineObj;
        var parent = stack.peek();

        var s = "";
        // ```まで読む
        while (true) {
            lineObj = srcLines.read();
            line = lineObj.line;
            if (/^\s*```\s*$/.test(line)) {
                break;
            }
            s += lineObj.line + "\n";
        }

        var o;

        try {
            o = jsyaml.safeLoad(s);
        }
        catch (e) {
            var errorMessage = "YAML の parse に失敗しました。";
            MyError(errorMessage, lineObj.filePath, lineObj.lineNum);
        }
        //printJSON(o);

        // 関数定義な文字列は関数にする
        _.forEach(o, function (v, k) {
            if (_.isString(v) && v.match(/^function/)) {
                o[k] = Function.call(this, 'return ' + v)();
            }
        });

//        // プリミティブな配列を { $value: value } な配列にする
//        function primitiveArrayToObjectArray(value, key, collection) {
//            // XXX: 要素数 1 以上前提の作り
//            if (_.isArray(value) && !_.isObject(value[1])) {
//                collection[key] = _.map(value, function(n) {
//                    return { $value: n };
//                });
//                _.forEach(collection[key], function(n) {
//                    _.forEach(n, primitiveArrayToObjectArray);
//                });
//            }
//            else if (_.isObject(value)) {
//                _.forEach(value, primitiveArrayToObjectArray);
//            }
//        }
//        _.forEach(o, primitiveArrayToObjectArray);
//        //printJSON(o);
//
//        // $index プロパティをセットする
//        function addIndexProperty(value, key, collection) {
//            if (_.isArray(value)) {
//                _.forEach(collection[key], function(element, index, collection) {
//                    collection[index].$index = index;
//                    _.forEach(collection[index], addIndexProperty);
//                });
//            }
//            if (_.isObject(value)) {
//                _.forEach(value, addIndexProperty);
//            }
//        }
//        _.forEach(o, addIndexProperty);
//        //printJSON(o);

        // 一旦は YAML の場合は記述位置に関係なくシートのrootに持っておくことにする
        var paramNode;
        for (paramNode = parent; paramNode.level != 1; paramNode = paramNode.parent) {
        }

        // TODO: 重複エラー出す
        if (_.isUndefined(paramNode.params)) {
            paramNode.params = {};
        }
        //_.assign(paramNode.params, o);  // 上書きする
        //_.defaults(paramNode.params, o);  // 上書きしない
        // deep merge
        // 同名の配列が宣言された場合は、後から宣言された方で丸ごと上書きされる
        paramNode.params = _.merge(paramNode.params, o, function(a, b) {
            if (_.isArray(a) && _.isArray(b)) {
                return b;
            }
        });
    }

    if (/^\s*```tsv\s*$/.test(line)) {
        var topLineObj = lineObj;
        var parent = stack.peek();
        var hLevel = (parent.kind === kindH) ? (parent.level + 1) : 0;
        var ulIndent = (parent.kind === kindUL) ? (parent.indent + 1) : 0;

        var tsvLines = [];
        // ```まで読む
        while (true) {
            lineObj = srcLines.read();
            line = lineObj.line;
            if (/^\s*```\s*$/.test(line)) {
                break;
            }
            tsvLines.push(lineObj);
        }

        // 文字が含まれる個数を求める
        var counter = function(str,seq) {
            return str.split(seq).length - 1;
        }

        // double quotation で囲まれているセルを含む行を連結
        var lines = [];
        for (var i = 0, l = ""; i < tsvLines.length; i++) {
            l += tsvLines[i].line;
            var n = counter(l, '"');
            if ((n % 2) == 0) {
                lines.push({line: l, lineNum: tsvLines[i].lineNum, originalText: tsvLines[i].line });
                l = "";
                continue;
            }
            l += "\n";
        }

        // 空行を除外
        lines = lines.filter(function(element, index, array) {
            return /\S/.test(element.line);
        });

        // ２行目が区切り行だったら１行目をheaderとみなす
        var tableHeaderIds = undefined;
        if (lines.length >= 3 && /^[-=_\*]{3,}/.test(lines[1].line)) {
            var ids = lines[0].line.match(/\|((\t\d+)+)\s*$/);
            if (ids) {
                tableHeaderIds = ids[1].split("\t").slice(1);
                for (var i = 0; i < tableHeaderIds.length; i++) {
                    tableHeaderIds[i] = parseInt(tableHeaderIds[i], 10);
                }
            }
            lines = lines.slice(2);
        }

        // ２次元配列化
        for (var i = 0; i < lines.length; i++) {
            // これだと " " 内のタブに対応できない
            //var data = lines[i].line.split("\t");
            // これだと空っぽのセルに対応できない
            //var data = lines[i].line.match(/(\"[^\"]+\"|[^\t]+)/g);

            var tabSplitted = lines[i].line.split("\t");
            var data = [];
            for (var j = 0; j < tabSplitted.length; j++) {
                var s = tabSplitted[j];
                if (s.charAt(0) === '"') {
                    // 文字列の末尾が " になるまで連結する
                    while (s.charAt(s.length - 1) !== '"') {
                        j++;
                        if (j >= tabSplitted.length) {
                            // TODO: エラーを出すべき。想定してない形式
                            break;
                        }
                        // どうせexcelではセル内のタブは表示に影響ないので、削除でも問題はないけど一応
                        s += "\t" + tabSplitted[j];
                    }
                }
                data.push(s);
            }

            // 前後の double quotation を削除
            for (var j = 0; j < data.length; j++) {
                if (data[j].charAt(0) === '"' && data[j].charAt(data[j].length - 1) === '"') {
                    data[j] = data[j].slice(1, -1);
                }
            }

            lines[i].data = data;
        }

        // １行目に "||" のセルがあれば、それより左のセルをheaderとする
        // ２行目以降のすべてに対して１行目の "||" が適用される（"||"と同じ列の中身は無視）
        var hNodeSplitColumn = 0;

        // XXX: 先頭に空行はない前提の作り
        // 行頭が区切りというのも想定してないので、1からで
        // 0 の場合は区切りなし
        for (var i = 1; i < lines[0].data.length; i++) {
            if (lines[0].data[i] === "||") {
                hNodeSplitColumn = i;
                break;
            }
        }
        //Error(hNodeSplitColumn +", "+ !hLevel);
        if (hNodeSplitColumn && !hLevel) {
            var errorMessage = "カテゴリーを項目の子階層として追加することはできません";
            MyError(errorMessage, topLineObj.filePath, topLineObj.lineNum);
        }
        // "||" 列を削除
        if (hNodeSplitColumn) {
            for (var i = 0; i < lines.length; i++) {
                lines[i].data.splice(hNodeSplitColumn, 1);
            }
        }

        var commentSplitColumn = 0;
        for (var i = 1; i < lines[0].data.length; i++) {
            if (lines[0].data[i] === "^") {
                commentSplitColumn = i;
                break;
            }
        }
        // "^" 列を削除
        if (commentSplitColumn) {
            for (var i = 0; i < lines.length; i++) {
                lines[i].data.splice(commentSplitColumn, 1);
            }
        }

        // １行目に "|" のセルがあれば、それより右のセルをデータとする
        // ２行目以降のすべてに対して１行目の "|" が適用される（"|"と同じ列の中身は無視）
        var dataSplitColumn = 0;

        // XXX: 先頭に空行はない前提の作り
        // 行頭が区切りというのも想定してないので、1からで
        // 0 の場合は区切りなし
        for (var i = 1; i < lines[0].data.length; i++) {
            if (lines[0].data[i] === "|") {
                dataSplitColumn = i;
                break;
            }
        }
        // "|" 列を削除
        if (dataSplitColumn) {
            for (var i = 0; i < lines.length; i++) {
                lines[i].data.splice(dataSplitColumn, 1);
            }
        }

        var uidColumn = 0;
        // 最終列にIDが存在するか
        // 最初に見つかった行でID列を判断
        for (var i = 0; i < lines.length; i++)
        {
            var line = lines[i].data;

            // 末尾の空白をすべて削除
            for (var j = line.length - 1; j >= 0 && !line[j]; --j)
            {
                line.pop();
            }
            // 空行
            if (line.length === 0)
            {
                continue;
            }
            var lastData = line[line.length - 1];
            if (/^\[#[\w\-]+\]$/.test(lastData))
            {
                uidColumn = line.length - 1;
                break;
            }
        }

        // tree の node を２次元配列で持つ
        for (var i = 0; i < lines.length; i++)
        {
            var data = lines[i].data;
            if (uidColumn)
            {
                var uid = data[uidColumn];

                // idセルを削除
                data = data.slice(0, uidColumn);

                // ID列は [#XXX] 形式という前提。他の形式は認めない
                if (uid)
                {
                    uid = uid.slice(2, -1);
                    {(function(){
                        var uidList = FindUidList(stack.peek());
                        if (uid in uidList)
                        {
                            var uidInfo0 = uidList[uid];
                            var errorMessage = "ID '#" + uid + "' が重複しています";
                            errorMessage += makeLineinfoString(uidInfo0.filePath, uidInfo0.lineNum);
                            errorMessage += makeLineinfoString(lineObj.filePath, lines[i].lineNum);
                            MyError(errorMessage);
                        }
                        else
                        {
                            uidList[uid] = { filePath: lineObj.filePath, lineNum: lines[i].lineNum };
                        }
                    })();}
                    lines[i].id = uid;
                }
            }
            if (dataSplitColumn)
            {
                var tableData = getDataFromTableRow(data.slice(dataSplitColumn), parent, tableHeaderIds);

                if (!tableData)
                {
                    var errorMessage = "シートに該当IDの確認欄がありません";
                    MyError(errorMessage, lineObj.filePath, lines[i].lineNum);
                }

                lines[i].tableData = tableData;
                data = data.slice(0, dataSplitColumn);
            }

            if (commentSplitColumn)
            {
                if (data[commentSplitColumn])
                {
                    lines[i].comment = data[commentSplitColumn];
                }
                data = data.slice(0, commentSplitColumn);
            }
            
            lines[i].nodes = data;
        }

/**/
        function BuildTree_Recurse(lines, hx1, idx, parentNode, y0, y1, x)
        {
            // 最後に検出されたnodeをleafとみなして終了
            if (lines[y0].nodes.length <= x)
            {
                var uid = lines[y0].id;
                if (!uid || fResetId)
                {
                    // leaf確定なので書き換え後のテキストをこの場で作る
                    var newSrcText = lines[y0].originalText;
                    if (!idx)
                    {
                        newSrcText += "\t";
                    }
                    else if (fResetId)
                    {
                        var match = newSrcText.match(/^(.*)\[#[\w\-]+\]$/);
                        if (match)
                        {
                            newSrcText = match[1];
                        }
                    }
                    newSrcText += "[#{uid}]";

                    AddNoIdNode(parentNode, lineObj, lines[y0].lineNum, newSrcText);
                }
                else
                {
                    parentNode.id = uid;
                }

                parentNode.tableData = lines[y0].tableData;
                parentNode.comment = lines[y0].comment;
                return;
            }
            
            for (var y = y0; y < y1;)
            {
                y0 = y;
                var line = lines[y0];
                var text = line.nodes[x];

                // とりあえず node を作って parent に追加
                if (text)
                {
                    var node;
                    if (x < hx1)
                    {
                        node = 
                        {
                            kind: kindH,
                            level: parentNode.level + 1,
                            text: text,
                            variables: {},
                            children: []
                        };
                    }
                    else
                    {
                        var indent = (parentNode.kind === kindUL) ? parentNode.indent + 1 : 0;
                        node = 
                        {
                            kind: kindUL,
                            indent: indent,
                            marker: "-",    // XXX: 一旦 "-" 固定で
                            group: -1,   // 場所確保のため一旦追加
                            depthInGroup: -1,   // 場所確保のため一旦追加
                            id: null,   // 場所確保のため一旦追加
                            text: text,
                            tableData: null,   // 場所確保のため一旦追加
                            comment: null,   // 場所確保のため一旦追加
                            variables: {},
                            children: []
                        };
                    }
                    AddChildNode(parentNode, node);
                    // leaf じゃないので削除
                    //delete parentNode.id;
                    //delete parentNode.tableData;
                }
                else
                {
                    // 空欄の場合は、親を引き継ぐ（スキップ）
                    node = parentNode;
                }

                // textが現れるか最後の行まで空欄をスキップ
                y++;
                for (; y < y1; y++)
                {
                    if (lines[y].nodes[x])
                    {
                        break;
                    }
                }

                BuildTree_Recurse(lines, hx1, idx, node, y0, y, x + 1);
            }
        }

        BuildTree_Recurse(lines, hNodeSplitColumn, uidColumn, parent, 0, lines.length, 0);
/*/

        // 処理中のnodeに対して、親となるnode
        var parentNode = [];
        for (var i = 0; i < lines.length; i++)
        {
            var line = lines[i].data;
            var uid = undefined;
            if (uidColumn > 0)
            {
                // ID列は [#XXX] 形式という前提。他の形式は認めない
                if (line[uidColumn])
                {
                    uid = line[uidColumn].slice(2, -1);
                    {(function(){
                        var uidList = FindUidList(stack.peek());
                        if (uid in uidList)
                        {
                            var uidInfo0 = uidList[uid];
                            var errorMessage = "ID '#" + uid + "' が重複しています";
                            errorMessage += makeLineinfoString(uidInfo0.filePath, uidInfo0.lineNum);
                            errorMessage += makeLineinfoString(lineObj.filePath, lines[i].lineNum);
                            MyError(errorMessage);
                        }
                        else
                        {
                            uidList[uid] = { filePath: lineObj.filePath, lineNum: lines[i].lineNum };
                        }
                    })();}
                }
                // idセルを削除
                line = line.slice(0, uidColumn);
            }

            var numColumns = (dataSplitColumn > 0) ? dataSplitColumn : line.length;

            for (var j = 0; j < numColumns; j++)
            {
                var text = line[j];
                if (!text)
                {
                    continue;
                }

                var data = undefined;
                // 右端の項目なら、右側のセルをtable dataとして取り込む
                if (j === numColumns - 1 && dataSplitColumn > 0)
                {
                    data = line.slice(dataSplitColumn);
                    data = getDataFromTableRow(data, parent);

                    if (!data)
                    {
                        var errorMessage = "シートに該当IDの確認欄がありません";
                        MyError(errorMessage, lineObj.filePath, lines[i].lineNum);
                    }
                }
                
                var item = 
                {
                    kind: kindUL,
                    indent: ulIndent + j,
                    id: (j === numColumns - 1) ? uid : undefined,
                    text: text,
                    tableData: data,
                    variables: {},
                    children: []
                };

                if (j === 0)
                {
                    AddChildNode(parent, item);
                }
                else
                {
                    // ひとつ左の階層の一番下のを親とする
                    AddChildNode(parentNode[j - 1], item);
                }

                parentNode[j] = item;

                if (j === numColumns - 1 && (!uid || fResetId))
                {
                    // tsv の場合はleaf確定なので書き換え後のテキストをこの場で作る
                    var newSrcText = lines[i].originalText;
                    if (!uidColumn)
                    {
                        newSrcText += "\t";
                    }
                    else if (fResetId)
                    {
                        var match = newSrcText.match(/^(.*)(?:\[#[\w\-]+\])$/);
                        if (match)
                        {
                            newSrcText = match[1];
                        }
                    }
                    newSrcText += "[#{uid}]";

                    AddNoIdNode(item, lineObj.filePath, lines[i].lineNum, newSrcText);
                }
            }
        }
/**/

        continue;
    }


    // obsolete
    /*
    var image = line.match(/^(\s*)!\[\]\((.+)\)$/);
    if (image)
    {
        stack.peek().image = image[2];
        continue;
    }
    */

    // 自由にプロパティを追加できるようにしてしまう…
    var property = line.match(/^\s*\[(.+)\]:\s+(.+)$/);
    if (property) {
        stack.peek().variables[_.trim(property[1])] = _.trim(property[2]);
    }

    var ColumnValueError = function(errorMessage, lineObj) {
        this.errorMessage = errorMessage;
        this.lineObj = lineObj;
    };

    // 行頭の (foo: 1, bar) 的な部分を parse
    function parseColumnValues(s, _isValueBase) {
        var isValueBase = _.isUndefined(_isValueBase) ? true : _isValueBase;

        var match = _.trimLeft(s).match(/^\((.+)\)\s+(.*)$/);
        if (!match) {
            match = _.trimLeft(s).match(/^\((.+)\)\s*$/);
        }
        if (!match) {
            return null;
        }
        var remain = match[2];
        var columnValues = [];
        var params = match[1].split(',');
        params.forEach(function(param) {
            param = _.trim(param);
            var keyMatch = param.match(/^[A-Za-z_]\w*$/);
            if (keyMatch) {
                var data = isValueBase ? { value: param } : { key: param };
                columnValues.push(data);
                return;
            }
            var keyValueMatch = param.match(/^([A-Za-z_]\w*)\s*:\s*(.*)$/);
            if (keyValueMatch) {
                var value = keyValueMatch[2];
                if (value.slice(0, 1) === '"' && value.slice(-1) === '"') {
                    value = eval(keyValueMatch[2]);
                }
                columnValues.push({
                    key: keyValueMatch[1],
                    value: value
                });
                return;
            }
            var value = param;
            if (value.slice(0, 1) === '"' && value.slice(-1) === '"') {
                value = eval(param);
            }
            columnValues.push({ value: value });
        });

        return {
            columnValues: columnValues,
            remain: remain
        }
    }

    try {

    // 初期値宣言
    (function() {
        var match = _.trim(line).match(/^\((\(.+\))\)$/);
        if (!match) {
            return;
        }
        line = match[1];

        var parse = parseColumnValues(line, false);
        if (parse === null) {
            return;
        }
        var columnNames = [];
        var defaultValues = {};
        parse.columnValues.forEach(function(param) {
            if (_.isUndefined(param.key)) {
                // key がない場合エラー
                var errorMessage = "列名の順序宣言には列名が必要です。";
                throw new ColumnValueError(errorMessage, lineObj);
            }
            columnNames.push(param.key);
            if (!_.isUndefined(param.value)) {
                defaultValues[param.key] = param.value;
            }
        });
        stack.peek().columnNames = columnNames;
        stack.peek().defaultColumnValues = defaultValues;
    })();

    // デフォルト値を正規表現で指定
    (function() {
        var match = _.trimRight(line).match(/^\/(.+)\/\s+(.+)$/);
        if (!match) {
            return;
        }

        var parse = parseColumnValues(match[2]);
        if (parse === null) {
            return;
        }

        var reString = match[1];
        var re = new RegExp(reString);
        var conditionalColumnValues = stack.peek().conditionalColumnValues;

        if (_.isUndefined(conditionalColumnValues)) {
            conditionalColumnValues = [];
        }

        parse.columnValues.forEach(function(param) {
            // TODO: (error)key, value が両方そろってない場合エラー
            if (_.isUndefined(param.key) || _.isUndefined(param.value)) {
            }
        });

        conditionalColumnValues.push({
            re: re,
            columnValues: parse.columnValues
        });

        stack.peek().conditionalColumnValues = conditionalColumnValues;

    })();

    }
    catch (e) {
        (function (errorMessage, lineObj) {
            if (_.isUndefined(lineObj)) {
                MyError(errorMessage);
            }
            else {
                MyError(errorMessage, lineObj.filePath, lineObj.lineNum);
            }
        })(e.errorMessage, e.lineObj);
    }

}

function getNumLeaves(node)
{
    if (node.children.length == 0)
    {
        return 1;
    }

    var n = 0;
    for (var i = 0; i < node.children.length; i++)
    {
        n += getNumLeaves(node.children[i]);
    }
    return n;
}

function getMaxLevel_Recurse(node, kind, max)
{
    if (node.kind == kind)
    {
        max = Math.max(max, node.level);
    }

    for (var i = 0; i < node.children.length; i++)
    {
        max = Math.max(max, getMaxLevel_Recurse(node.children[i], kind, max));
    }

    return max;
}
function getMaxLevel(node, kind)
{
    return getMaxLevel_Recurse(node, kind, 0);
}

function getString(node)
{
    var s = "";
    s += (node.kind == kindUL) ? "UL" : "H";
    s += node.group + "-" + node.depthInGroup;
    s += "(" + node.children.length + ", " + getNumLeaves(node) + ")";
    s += ": " + node.text;
    s += "\n";
    for (var i = 0; i < node.children.length; i++)
    {
        s += getString(node.children[i]);
    }
    return s;
}

//Error(getString(root));

// 確認欄の指定がないシートはデフォルトで
// TODO: テキストは外から指定できるように
root.children.forEach(function(element, index, array) {
    if (element.tableHeaders.length === 0) {
        element.tableHeaders = [{ name: "確認欄", id: 1 }];
    }
});

// 空っぽのシートがないか確認
root.children.forEach(function(element, index, array) {
    if (element.children.length === 0) {
        var errorMessage = "シート「"+ element.text +"」に項目が存在しません\n※シートには最低１個の項目が必要です";
        var lineObj = element.lineObj;
        MyError(errorMessage, lineObj.filePath, lineObj.lineNum);
    }
});

_.forEach(noIdNodes, function(infos) {
    // id を付与してファイルに書き出すノードを抽出
    infos = infos.filter(function(element, index, array) {
        // H1 ノード
        if (element.node.kind === kindH && element.node.level === 1) {
            return true;
        }

        // leaf ノード
        // ただし '[', ']' は除外
        if (element.node.children.length === 0) {
            return (element.node.text !== '[' && element.node.text !== ']');
        }

        // alias 参照ノード
        if (/^\*[A-Za-z_]\w*\(.*\)$/.test(element.node.text.trim())) {
            return true;
        }

        return false;
    });

    // TODO: 複数箇所で include されてる時に異なる id が振られないように
    _.forEach(infos, function(info) {
        var node = info.node;
        var uidList = FindUidList(node.parent);
        var uid = createUid(8, uidList);
        node.id = uid;

        // ID 挿入して書き換え
        // "{{foo}}" みたいな文法を作ろうとしたら {} に置換されてしまうので、汎用 format ではなく "{uid}" 専用の replace 処理に
        //var newSrcText = info.newSrcText.format({uid: uid});
        var newSrcText = info.newSrcText.replace(/\{uid\}/, uid);

        AddSrcTextToRewrite(info, newSrcText);
    });
});

// 配列を展開
// '[' ']' の node で挟まれた node を配列とみなし、 leaf に ']' ノードの子treeをすべてコピー
// XXX: 配列の階層は一旦考えない（非対応とする）。使えた方が便利？
(function() {
    var ArrayError = function(errorMessage, node) {
        this.errorMessage = errorMessage;
        this.node = node;
    };

    // TODO: alias の同じ function と統合する
    // subTree に対してそのまま cloneDeep を呼ぶと、 parent をさかのぼって tree 全体が clone されるので対処
    function cloneSubTree(srcSubTree) {
        // FIXME: 
        return _.cloneDeep(srcSubTree);
    }

    function unrollArray(parent, index) {
        parent.children[index] = null;
        // 直近の弟の ] ノードを見つける
        var leaves = [];
        var commonChildren = null;

        for (var i = index + 1; i < parent.children.length; i++) {
            var sibling = parent.children[i];

            if (sibling.text === ']') {
                commonChildren = sibling.children;
                parent.children[i] = null;
                break;
            }

            // 同一階層の入れ子の検出は面倒なので対応しない
            // できた方が便利な状況が頻発すれば検討
            if (sibling.text === '[') {
                var errorMessage = "配列の同一階層における入れ子構造はには対応していません。";
                throw new ArrayError(errorMessage, sibling);
            }

            // sibling 以下のすべての leaf を追加
            forAllNodes_Recurse(sibling, parent, i, function(node, parent, index) {
                if (node === null) {
                    return true;
                }
                if (node.children.length === 0) {
                    if (node.text === '[') {
                        unrollArray(parent, index);
                    }
                    else {
                        leaves.push(node);
                    }
                    return true;
                }
            });
        }
        if (commonChildren === null) {
            // 配列が閉じられてないエラー
            var errorMessage = "配列が ']' で閉じられていません。";
            throw new ArrayError(errorMessage, node);
        }

        leaves.forEach(function(leaf, index, array) {
            // leaf に clone した commonChildren を追加
            for (var i = 0; i < commonChildren.length; i++) {
                // まず clone
                var child = cloneSubTree(commonChildren[i]);

                // group とかはこの後処理されるので、マーカーがちゃんとしてれば OK なはず
                forAllNodes_Recurse(child, null, -1, function(node, parent, index) {
                    if (node.children.length === 0) {
                        // id 連結
                        node.id = leaf.id + "_" + node.id;
                    }
                });
    
                AddChildNode(leaf, child);
            }
        });

    }

    try {
        forAllNodes_Recurse(root, null, -1, function(node, parent, index) {
            if (node === null) {
                return true;
            }
            if (parent === null) {
                return;
            }
            if (node.kind !== kindUL) {
                return;
            }
            if (node.text !== '[') {
                if (node.text === ']') {
                    // [ がないのにいきなり ] がきたらエラー
                    var errorMessage = "配列ではない場所に ] が存在します。";
                    throw new ArrayError(errorMessage, node);
                }
                return;
            }

            unrollArray(parent, index);

            return true;
        });
    }
    catch (e) {
        (function (errorMessage, node) {
            var lineObj = node.lineObj;
            if (_.isUndefined(lineObj)) {
                MyError(errorMessage);
            }
            else {
                MyError(errorMessage, lineObj.filePath, lineObj.lineNum);
            }
        })(e.errorMessage, e.node);
    }

    // children の null の要素を削除して shrink
    function shrinkChildrenArrayforAllNodes() {
        forAllNodes_Recurse(root, null, -1, function(node, parent, index) {
            var validChildren = node.children.filter(function(element, index, array) {
                return (element !== null);
            });
            node.children = validChildren;
        });
    }
    shrinkChildrenArrayforAllNodes();

})();

// group と depthInGroup を計算
forAllNodes_Recurse(root, null, -1, function(node, parent, index) {
    if (node.kind !== kindUL) {
        return;
    }

    if (parent.kind === kindUL) {
        var markerChanged = (parent.marker !== node.marker);
        var parentNodeGroup = parent.group;

        node.group = markerChanged ? (parentNodeGroup + 1) : parentNodeGroup;
        node.depthInGroup = markerChanged ? 0 : (parent.depthInGroup + 1);
    }
    else {
        node.group = 0;
        node.depthInGroup = 0;
    }
});

// marker は不要なので削除
CL.deletePropertyForAllNodes(root, "marker");

//function echoJson(obj, name) {
//    var s = JSON.stringify(obj, undefined, 4);
//    if (!_.isUndefined(name)) {
//        s = name + ":\n" + s;
//    }
//    WScript.Echo(s);
//}

// エイリアス埋め込み
// まずはすべてのノードについて調べ、親に登録
(function() {
    var startTime = performance.now();

    var AliasError = function(errorMessage, node) {
        this.errorMessage = errorMessage;
        this.node = node;
    };

    function aliasError(errorMessage, node) {
        var lineObj = node.lineObj;
        if (_.isUndefined(lineObj)) {
            MyError(errorMessage);
        }
        else {
            MyError(errorMessage, lineObj.filePath, lineObj.lineNum);
        }
    }

    function evalParameters(paramsStr, node, currentParameters) {
        paramsStr = paramsStr.trim();
        if (paramsStr == "") {
            return {};
        }

        var referableParams = {};
        if (!_.isUndefined(currentParameters)) {
            _.defaults(referableParams, currentParameters);
        }
        for (var parent = node.parent; !_.isUndefined(parent); parent = parent.parent) {
            if (!_.isUndefined(parent.params)) {
                _.defaults(referableParams, parent.params);
            }
        }

        // XXX: これだと { foo: "bar" } を { foo: "bar" } としてしまうのでNG
        //function replacer(m, k) {
        //    return "referableParams." + k;
        //}
        //paramsStr = paramsStr.replace( /([_A-Za-z]\w*)/g, replacer);
        //alert(paramsStr);

        // XXX: 処理が重すぎる
        // TODO: 計測するべき
        //function parseParams(referableParams, paramsStr) {
        //    var s = "";
        //    _.forEach(referableParams, function(value, key) {
        //        // XXX: key が添字な文字列、value が undefined な値が来ることがあるので対処。理由は調査できてない…
        //        if (_.isUndefined(value)) {
        //            return;
        //        }
        //        s += "var " + key + "=referableParams." + key + ";";
        //    });
        //    s += "var result = [" + paramsStr + "];";
        //    eval(s);
//
        //    return result;
        //}
        function parseParams(referableParams, paramsStr) {
            // _.keys(), _.values() は列挙順は保証されてないので一応自前で詰めておく
            //var keys = _.keys(referableParams);
            //var values = _.values(referableParams);
            var keys = [];
            var values = [];
            _.forEach(referableParams, function(value, key) {
                // XXX: key が添字な文字列、value が undefined な値が来ることがあるので対処。理由は調査できてない…
                if (_.isUndefined(value)) {
                    return;
                }
                keys.push(key);
                values.push(value);
            });
            keys.push("__paramsStr");
            values.push(paramsStr);
            var f = Function(keys.join(","), 'return eval("([" + __paramsStr + "])");');
    
            return f.apply(null, values);
        }
        var paramsArray = parseParams(referableParams, paramsStr);

        // object を返すには丸括弧が必要らしい
        //var paramsArray = eval("([" + paramsStr + "])");
        //printJSON(paramsArray);
        if (paramsArray.length == 1) {
            return paramsArray[0];
        }

        // ここでマージしたものを展開してしまう？
        // 一番長い配列を調べて展開。配列なら index でアクセス。objectならそのまま。先頭から _.defaults() でマージして push
        var maxArrayElem = _.max(paramsArray, function(elem) {
            return _.isArray(elem) ? elem.length : 0;
        });
        var maxArrayLength = _.isArray(maxArrayElem) ? maxArrayElem.length : 1;
        if (maxArrayLength == 0) {
            // TODO: 例外投げる
        }
        var mergedArray = [];
        _.forEach(_.range(maxArrayLength), function(i) {
            var o = {};
            _.forEach(paramsArray, function(elem) {
                if (_.isArray(elem)) {
                    if (i < elem.length) {
                        if (_.isObject(elem[i])) {
                            _.defaults(o, elem[i]);
                        }
                        else {
                            _.defaults(o, {$value: elem[i]});
                        }
                    }
                }
                // 関数が渡された場合、引数に渡された順に実行
                else if (_.isFunction(elem)) {
                    // o はこの関数で作った object なので clone 不要
                    // 関数の中で直接書き換えてもOK
                    o = elem(o);
                }
                else {
                    _.defaults(o, elem);
                }
            });
            mergedArray.push(o);
        });
        // ほぼ意味ないけど、要素数1の場合はobjectを返す
        return (mergedArray.length == 1) ? mergedArray[0] : mergedArray;

//        // 直接 object 渡しの場合は { } で囲む
//        if (/^\{.+\}$/.test(paramsStr)) {
//            // object を返すには丸括弧が必要らしい
//            var params = eval("(" + paramsStr + ")");
//            //params.$params = params;
//            return params;
//        }
//
//        var referableParams = {};
//        if (!_.isUndefined(currentParameters)) {
//            _.defaults(referableParams, currentParameters);
//        }
//        for (var parent = node.parent; !_.isUndefined(parent); parent = parent.parent) {
//            if (!_.isUndefined(parent.params)) {
//                _.defaults(referableParams, parent.params);
//            }
//        }
//        //printJSON(referableParams);
//
//        // TODO: , で split して順に _.default() で集める
//        // TODO: 1個目が配列の場合、2個目以降と扱いを分ける
//        var params = _.get(referableParams, paramsStr);
//        if (!_.isUndefined(params)) {
//            //params.$params = params;
//            return params;
//        }
//
//        // TODO: 該当する名前のパラメータオブジェクトが見つからない場合は例外投げる
//        return {};
    }

    // subTree に対してそのまま cloneDeep を呼ぶと、 parent をさかのぼって tree 全体が clone されるので対処
    function cloneSubTree(srcSubTree) {
//        var rootParent = srcSubTree.parent;
//        CL.deletePropertyForAllNodes(srcSubTree, "parent");
//
//        var dst = _.cloneDeep(srcSubTree);
//
//        //srcSubTree.parent = rootParent;
//        //dst.parent = rootParent;
//
//        forAllNodes_Recurse(srcSubTree, rootParent, -1, function(node, parent, index) {
//            if (node === null) {
//                return true;
//            }
//            node.parent = parent;
//        });
//        forAllNodes_Recurse(dst, rootParent, -1, function(node, parent, index) {
//            if (node === null) {
//                return true;
//            }
//            node.parent = parent;
//        });
//
//        return dst;

        // 自前で tree をたどって全 node を shallow copy
        var dstSubTree = _.assign({}, srcSubTree);

        function _recurse(dstNode, srcNode) {
            dstNode.children = [];
            _.forEach(srcNode.children, function(srcChild) {
                if (srcChild === null) {
                    return;
                }
                var dstChild = _.assign({}, srcChild);
                dstChild.parent = dstNode;
                dstNode.children.push(dstChild);
                _recurse(dstChild, srcChild);
            });
        }

        _recurse(dstSubTree, srcSubTree);

        return dstSubTree;

//        // root node を shallow copy
//        //var dst = Object.assign(srcSubTree);
//        var dst = {};
//        _.assign(dst, srcSubTree);
//
//        dst.parent = null;
//
//        return _.cloneDeep(dst);
    }

    // データを定義できるように
    forAllNodes_Recurse(root, null, -1, function(node, parent, index) { }, function(node, parent, index) {
        if (parent === null) {
            return;
        }
        if (node.kind !== kindUL) {
            return;
        }

        var match = node.text.trim().match(/^&([A-Za-z_]\w*):$/);
        if (match === null) {
            return;
        }

        var paramName = match[1];

        if (!_.isUndefined(parent.params)) {
            // 重複エラー
            if (paramName in parent.params) {
                var errorMessage = "データ名'"+ paramName +"'が重複しています。";
                aliasError(errorMessage, node);
            }
        }
        else {
            parent.params = {};
        }

        var param = [];
        // XXX: 1階層の単純な構成の想定。エラーチェックとかは一切しない
        _.forEach(node.children, function(child) {
            var o = {
                $value: child.text,
                $id: child.id
            };
            param.push(o);
        });
        //printJSON(param);

        parent.params[paramName] = param;

        // 親の children の自分自身を null に
        parent.children[index] = null;
    });

    // すべての alias を tree から取り外し、所属 node にリストアップしておく
    forAllNodes_Recurse(root, null, -1, function(node, parent, index) { }, function(node, parent, index) {
        if (parent === null) {
            return;
        }
        if (node.kind !== kindUL) {
            return;
        }

        var match = node.text.trim().match(/^&([A-Za-z_]\w*)\(\)$/);
        if (match === null) {
            return;
        }

        var subTreeName = match[1];

        if ("subTrees" in parent) {
            // 重複エラー
            if (subTreeName in parent.subTrees) {
                var errorMessage = "エイリアス名'"+ subTreeName +"'が重複しています。";
                aliasError(errorMessage, node);
            }
        }
        else {
            parent.subTrees = {};
        }
        parent.subTrees[subTreeName] = node;

        // node の group 関係を subtree root からの offset 値に
        // 木の中で宣言した場合でも大丈夫なように対応しておく
        var subTreeGroup = node.group;
        var subTreeDepthInGroup = node.depthInGroup;
        forAllNodes_Recurse(node, null, -1, function(node, parent, index) {
            if (node.group === subTreeGroup) {
                // subTreeRoot と同じ group の node の depthInGroup は必ず 1 多いので引いておく
                node.depthInGroup -= subTreeDepthInGroup + 1;
            }
            node.group -= subTreeGroup;
        });

        // 親の children の自分自身を null に
        parent.children[index] = null;

        // TODO: 宣言時に記述する defaultParameter 仕様は廃止の方向で
        // TODO: 上に遡ってアクセスできたものを defaultParameter とする（下層優先）
        // 現状 evalParameters() の仕様が defaultParameter 仕様に合っていない
        //try {
        //    node.defaultParameters = evalParameters(match[2], node);
        //}
        //catch(e) {
        //    var errorMessage = "パラメータが不正です。";
        //    aliasError(errorMessage, node);
        //}
        //printJSON(node.defaultParameters);
    });

    // children の null の要素を削除して shrink
    function shrinkChildrenArray(node, parent, index) {
        forAllNodes_Recurse(node, parent, index,
            function(node, parent, index) {
                if (_.isUndefined(node)) {
                    return true;
                }
                if (node === null) {
                    return true;
                }
                if (node.children.length === 0) {
                    return true;
                }
            },
            function(node, parent, index) {
                var validChildren = node.children.filter(function(element, index, array) {
                    return (element !== null);
                });
                if (validChildren.length === 0) {
                    if (node.kind === kindH) {
                        var errorMessage = "シート「"+ node.text +"」に有効な項目が存在しません\n※子階層がエイリアスのみとなっている可能性があります";
                        var lineObj = node.lineObj;
                        MyError(errorMessage, lineObj.filePath, lineObj.lineNum);
                    }
                    if (parent !== null) {
                        parent.children[index] = null;
                        delete node;
                    }
                    return;
                }
                node.children = validChildren;
            }
        );
    }

    function shrinkChildrenArrayforAllNodes() {
        shrinkChildrenArray(root, null, -1);
    }

    shrinkChildrenArrayforAllNodes();

    // 名前からtreeをさかのぼって見つける
    // なければ null を返す
    function findSubTree_Recurse(subTreeName, node) {
        if (_.isUndefined(node) || node === null) {
            return null;
        }
        if ("subTrees" in node) {
            if (subTreeName in node.subTrees) {
                return node.subTrees[subTreeName];
            }
        }
        return findSubTree_Recurse(subTreeName, node.parent);
    }

    // すべての alias 内の alias 参照を事前に展開しておく
//    forAllNodes_Recurse(root, null, -1, function(node, parent, index) {
//        if (_.isUndefined(node.subTrees)) {
//            return;
//        }
//
//        _.forEach(node.subTrees, function(subTreeRoot, name) {
//
//            verifyReference(subTreeRoot);
//
//            forAllNodes_Recurse(subTreeRoot, null, -1, function(node, parent, index) {
//                if (node === null) {
//                    return true;
//                }
//                var match = node.text.trim().match(/^\*([A-Za-z_]\w*)\((.*)\)$/);
//                if (match === null) {
//                    return;
//                }
//                var subTreeName = match[1];
//
//                try {
//                    var parameters = evalParameters(match[2], node);
//                }
//                catch(e) {
//                    var errorMessage = "パラメータが不正です。";
//                    aliasError(errorMessage, node);
//                }
//        
//                addSubTree(node, index, subTreeName, parameters);
//            });
//
//            shrinkChildrenArray(subTreeRoot, null, -1);
//        });
//    });

    // 一旦 parent を削除
    // subtree だけ deep clone したつもりが parent をさかのぼって tree 全体が丸ごと clone されてしまうので
//    forAllNodes_Recurse(root, null, -1, function(node, parent, index) {
//        if (_.isUndefined(node.subTrees)) {
//            return;
//        }
//        _.forEach(node.subTrees, function(subTreeRoot, name) {
//            //CL.deletePropertyForAllNodes(subTreeRoot, "parent");
//            delete subTreeRoot.parent;
//        });
//    });


    function nodeToString(root) {
        var depth = 0;
        var s = "";
        forAllNodes_Recurse(root, null, -1,
            function(node, parent, index) {
                if (node === null) {
                    return true;
                }
                var indent = _.repeat("    ", depth);
                s += index + " : "
                s += indent + "(" + node.group + " ," + node.depthInGroup + ")  " + node.text + "\n";
                depth++;
            },
            function(node, parent, index) {
                depth--;
            }
        );
        return s;
    }

    // 問題がないか調べる
    // 一度確認した subtree は isValid フラグ立てておく。json 出力前に delete
    function verifyReference(subTreeRoot) {

        function _recurse(subTree, callStack) {
            if (subTree.isValidSubTree) {
                return;
            }

            if (subTree.children.length === 0) {
                var errorMessage = "エイリアスには1個以上の子ノードが必要です。";
                throw new AliasError(errorMessage, subTree);
            }

            for (var i = 0; i < subTree.children.length; i++) {
                if (subTree.children[i].group !== subTree.group) {
                    var errorMessage = "エイリアスの第2階層はグループ切り替えはできません。\nルート（エイリアス名の行）と同じマークにしてください";
                    throw new AliasError(errorMessage, subTree);
                }
            }

            var subTreeName = subTree.text.slice(2, -1);
            var lineObj = subTree.lineObj;
            var callName = subTreeName + ":" + lineObj.filePath + ":" + lineObj.lineNum;
            if (_.indexOf(callStack, callName) >= 0) {
                //WScript.Echo(callStack.toString()+"\n"+callName);
                var errorMessage = "エイリアス'"+ subTreeName +"'に循環参照が存在します。";
                throw new AliasError(errorMessage, subTree);
            }
            callStack.push(callName);

            forAllNodes_Recurse(subTree, null, -1, function(node, parent, index) {
                var match = node.text.trim().match(/^\*([A-Za-z_]\w*)\(.*\)$/);
                if (match === null) {
                    return;
                }
                //if (node.children.length > 0) {
                //    // 参照は leaf 以外は認めないのでエラー
                //    var errorMessage = "エイリアスを参照できるのは葉ノードだけです。";
                //    throw new AliasError(errorMessage, node);
                //}
                var refSubTreeName = match[1];

                var refSubTree = findSubTree_Recurse(refSubTreeName, node.parent);

                // みつからなかった
                if (refSubTree === null) {
                    var errorMessage = "エイリアス'" + refSubTreeName + "'は存在しません。";
                    throw new AliasError(errorMessage, node);
                }

                _recurse(refSubTree, callStack);
            });

            callStack.pop();

            subTree.isValidSubTree = true;
        }

        try {
            _recurse(subTreeRoot, []);
        }
        catch (e) {
            if (_.isUndefined(e.node) || _.isUndefined(e.errorMessage)){
                throw e;
            }
            //WScript.Echo(JSON.stringify(e, undefined, 4));
            aliasError(e.errorMessage, e.node);
        }
    }
    
    // node に sub tree の clone を追加する
    // 展開前の状態で追加
    function addSubTree(targetNode, targetIndex, subTreeName, parameters) {
        function rollArray(targetNode, targetIndex, subTreeName, parameters) {
            var clonedTargetNodes = [];
            _.forEach(parameters, function(element, index) {
                var node = cloneSubTree(targetNode);
                if (!_.isObject(element)) {
                    element = {
                        $value: element
                    };
                }
                var elementId = ("$id" in element) ? element.$id : "i" + index;
                node.id = targetNode.id + "_" + elementId;

                node.tempParams = {
                    $index: index,
                    $value: element.$value
                };

                var paramJSON = JSON.stringify(element);
                //var match = node.text.match(/^\*[A-Za-z_]\w*\((.*)\)$/);
                //var paramName = match[1];

                //node.text = "*" + subTreeName + "(" + paramName + "[" + index + "])" + match[2];
                node.text = "*" + subTreeName + "(" + paramJSON + ")";
                clonedTargetNodes.push(node);
            });

            targetNode.parent.children[targetIndex] = null;
            //Array.prototype.splice.apply(targetNode.parent.children, [targetIndex + 1, 0].concat(clonedTargetNodes));
            var a = targetNode.parent.children;
            var insertedChildren = a.slice(0, targetIndex+1).concat(clonedTargetNodes).concat(a.slice(targetIndex+1));
            insertedChildren[targetIndex] = null;
            targetNode.parent.children = insertedChildren;

            // ここではノードの追加のみ
            // 処理自体はループの後ろでされる想定
        }

        //printJSON(parameters);
        if (_.isArray(parameters)) {
            rollArray(targetNode, targetIndex, subTreeName, parameters);
            return;
        }

        var subTree = findSubTree_Recurse(subTreeName, targetNode.parent);

        // みつからなかった
        if (subTree === null) {
            var errorMessage = "エイリアス'" + subTreeName + "'は存在しません。";
            throw new AliasError(errorMessage, targetNode);
        }

        // まず clone
        subTree = cloneSubTree(subTree);

        // TODO: defaultParameter 仕様は廃止の方向で
        // TODO: テンプレ宣言時のデフォルト値仕様をやめて、この位置から参照可能な定数を入れておく
        //_.forEach(subTree.defaultParameters, function(value, key) {
        //    if (_.isUndefined(parameters[key])) {
        //        parameters[key] = value;
        //    }
        //});

        // 配列展開時に追加したパラメータ
        // $index, $value
        if (!_.isUndefined(targetNode.tempParams)) {
            // primitive array のために必要な対応
            if (!_.isObject(parameters)) {
                parameters = {};
            }
            _.defaults(parameters, targetNode.tempParams);
        }

        // 変数展開
        //if (!_.isEmpty(parameters)) {
        {
            //printJSON(parameters);

            // この後の eval 内でプロパティに直接アクセスできるように
            // primitive array のために必要な対応
            if (_.isObject(parameters)) {
                // XXX: いろいろやったらややこしいことになったので一旦はシンプルに追加
                // XXX: ただこれをやると JSON.stringify が使えなくなるので極力避けたい
                // XXX: ここではローカル変数で宣言して evalParameters にその $params 渡すとか…？
                parameters.$params = _.assign(parameters);

                var s = "";
                _.forEach(parameters, function(value, key) {
                    // XXX: key が添字な文字列、value が undefined な値が来ることがあるので対処。理由は調査できてない…
                    if (_.isUndefined(value)) {
                        return;
                    }
                    //var valueStr = JSON.stringify(value);
                    //alert(valueStr);
                    //if (_.isString(value)) {
                    //    value = "'" + value + "'";
                    //}
                    //if (_.isArray(value)) {
                    //    value = "[" + value + "]";
                    //}
                    //var s = key + "=" + valueStr;
                    s += "var " + key + "=parameters." + key + ";";
                });
                //s += "var $params=_.assign(parameters);"
                eval(s);
            }

            // 省略時はこれを使う
            // 引数1個の場合を想定した仕様だけど、複数あってもエラーにはしないでおく
            var defaultParam = "$value";
            var firstParam = _.find(_.keys(parameters), function(s) {
                return s.substr(0, 1) != "$";
            });
            if (!_.isUndefined(firstParam)) {
                defaultParam = firstParam;
            }
            defaultParam = "{{" + defaultParam + "}}";

            forAllNodes_Recurse(subTree, null, -1, function(node, parent, index) {
                // あるnodeに1個でも false 的なものが渡されたら、それ以下のnode削除
                function replaceVariable(s) {
                    if (!s) {
                        return void(0);
                    }

                    // {{}} のように省略した場合は、引数（無名の場合は $value）を指定したものとみなす
                    // 引数1個の想定の仕様。引数が複数の場合は何が参照されるかは保証はない
                    s = s.replace(/\{\{\s*\}\}/g, defaultParam);

                    var toDelete = false;
                    function replacer(m, k) {
                        if (toDelete) {
                            return "";
                        }
                        //k = k.trim();
                        var parameter;
                        //var match = k.trim().match(/^eval\((.+)\)$/);
                        if (/^[\w\$\.\[\]]+$/.test(k)) {
                            // 変数そのままと . と [] でのアクセスだけ別処理
                            // 少しでも処理が軽くなることを期待。意味なさそうだけど一応
                            parameter = _.get(parameters, k, null);
                        }
                        else {
                            parameter = eval("(" + k + ")");
                        }

                        // 明示的に省略を指定させたいので、未定義は対象外としておくことも考えたが、使い勝手的に省略で削除できる方が便利なので、そうする
                        if (parameter === false || _.isUndefined(parameter) ||  _.isNull(parameter)) {
                            toDelete = true;
                            return "";
                        }
                        if (parameter === true) {
                            return "";
                        }
                        return parameter;
                    }
                    var replaced = s.replace( /\{\{\s*([^\}]+)\s*\}\}/g, replacer);
                    return toDelete ? void(0) : replaced;
                }
                node.text = replaceVariable(node.text);
                if (_.isUndefined(node.text)) {
                    node.parent.children[index] = null;
                    return;
                }
                node.comment = replaceVariable(node.comment);
                node.imageFilePath = replaceVariable(node.imageFilePath);
            });
            shrinkChildrenArray(subTree, null, -1);
        }

        // XXX: node に循環参照があるので JSON.stringify は使えない
        //subTree = JSON.parse(JSON.stringify(subTree));

        // subtree 内の subtree 呼び出し
        forAllNodes_Recurse(subTree, null, -1, function(node, parent, index) {
            if (parent === null) {
                return;
            }
            var match = node.text.trim().match(/^\*([A-Za-z_]\w*)\((.*)\)$/);
            if (match === null) {
                return;
            }
            var subTreeName = match[1];

            try {
                var parsedParameters = evalParameters(match[2], node, parameters);
            }
            catch(e) {
                var errorMessage = "パラメータが不正です。\n\n" + e.message;
                aliasError(errorMessage, node);
            }

            //if (!_.isArray(parsedParameters)) {
            //    if (match[4] != "") {
            //        (function(){
            //            var functionName = match[4];
            //            var f = null;
            //            for (var parent = node.parent; parent !== null; parent = parent.parent) {
            //                if (_.isUndefined(parent.params)) {
            //                    continue;
            //                }
            //                if (functionName in parent.params) {
            //                    var s = parent.params[functionName];
            //                    f = Function.call(this, 'return ' + s)();
            //                    break;
            //                }
            //            }
            //            // TODO: 該当する関数が見つからない時にエラー出す
            //            if (f) {
            //                // 関数内で直接アクセスできるように
            //                if (!_.isUndefined(node.tempParams)) {
            //                    _.forEach(node.tempParams, function(value, key) {
            //                        var valueStr = JSON.stringify(value, undefined, 4);
            //                        var s = key + "=" + valueStr;
            //                        eval(s);
            //                    });
            //                }
            //                
            //                var inputParameters = _.assign({}, parsedParameters);
            //                // FIXME: 一旦全部見えるようにしておく。渡されたものだけ参照可能になるようにするべき
            //                _.defaults(inputParameters, parameters);
            //                var outputParameters = f(inputParameters);
            //                // 触らなかったのはそのまま使えるように
            //                parsedParameters = _.defaults(outputParameters, parsedParameters);
            //            }
            //        })();
            //    }
            //}

            addSubTree(node, index, subTreeName, parsedParameters);
        });

        // subtree の leaf に target の子ノードを追加する
        if (targetNode.children.length > 0) {
            var targetClone = cloneSubTree(targetNode);

            // offset にしておく
            forAllNodes_Recurse(targetClone, null, -1, function(node, parent, index) {
                if (node.group === targetNode.group) {
                    node.depthInGroup -= targetNode.depthInGroup;
                }
                node.group -= targetNode.group;
            });
//            forAllNodes_Recurse(targetClone, null, -1, function(node, parent, index) {
//                if (parent === null) {
//                    return;
//                }
//                var match = node.text.trim().match(/^\*([A-Za-z_]\w*)\((.*)\)$/);
//                if (match !== null) {
//                    var subTreeName = match[1];
//
//                    try {
//                        var parameters = evalParameters(match[2], node);
//                    }
//                    catch(e) {
//                        var errorMessage = "パラメータが不正です。";
//                        aliasError(errorMessage, node);
//                    }
//        
//                    addSubTree(node, index, subTreeName, parameters);
//                }
//            });

            forAllNodes_Recurse(subTree, null, -1, function(node, parent, index) {
                if (node.children.length > 0) {
                    return;
                }
                // 内容は不問
                if (_.has(node, 'attributes.sealed')) {
                    return;
                }
                var subTreeLeaf = node;
                var target = cloneSubTree(targetClone);
                forAllNodes_Recurse(target, null, -1, function(node, parent, index) {
                    if (node === null) {
                        return true;
                    }
                    if (node.group === 0) {
                        node.depthInGroup += subTreeLeaf.depthInGroup;
                    }
                    node.group += subTreeLeaf.group;
                    if (node.children.length === 0) {
                        // id を _ で連結
                        node.id = subTreeLeaf.id + "_" + node.id;
                        return true;
                    }
                });
                subTreeLeaf.children = target.children;
                return true;
            });
        }

        // subTree の 全 node の group と leaf の id を書き換える
        forAllNodes_Recurse(subTree, null, -1, function(node, parent, index) {
            if (node === null) {
                return true;
            }
            // group 関係は subtree root からのオフセットとして扱う
            if (node.group === 0) {
                node.depthInGroup += targetNode.depthInGroup;
            }
            node.group += targetNode.group;
            if (node.children.length === 0) {
                // id を _ で連結
                node.id = targetNode.id + "_" + node.id;
                return true;
            }
        });

        // splice で自分を subTree の children で置き換える
        // ループを正しくたどれるように置き換えでなく、直後に挿入 + 削除予約
        // splice は配列のまま渡せない。spread構文も使えないのでconcatとか使ってやる
        //var insertedChildren = targetNode.parent.children.splice(targetIndex + 1, 0, subTree.children);
        var a = targetNode.parent.children;
        // subTree の parent 書き換え
        for (var i = 0; i < subTree.children.length; i++) {
            if (subTree.children[i] === null) {
                continue;
            }
            subTree.children[i].parent = targetNode.parent;
        }
        var insertedChildren = a.slice(0, targetIndex+1).concat(subTree.children).concat(a.slice(targetIndex+1));
        insertedChildren[targetIndex] = null;
        targetNode.parent.children = insertedChildren;
    }

    // sub tree をインライン展開していく
    forAllNodes_Recurse(root, null, -1, function(node, parent, index) {
        if (node === null) {
            return true;
        }
        // XXX: コピペを何とかする
        var match = node.text.trim().match(/^\*([A-Za-z_]\w*)\((.*)\)$/);
        if (match !== null) {
            var subTreeName = match[1];

            try {
                var parameters = evalParameters(match[2], node, {});
            }
            catch(e) {
                var errorMessage = "パラメータが不正です。\n\n" + e.message;
                aliasError(errorMessage, node);
            }

            // XXX: コピペを何とかする
            //if (!_.isArray(parameters)) {
            //    if (match[4] != "") {
            //        (function(){
            //            var functionName = match[4];
            //            var f = null;
            //            for (var parent = node.parent; parent !== null; parent = parent.parent) {
            //                if (_.isUndefined(parent.params)) {
            //                    continue;
            //                }
            //                if (functionName in parent.params) {
            //                    var s = parent.params[functionName];
            //                    f = Function.call(this, 'return ' + s)();
            //                    break;
            //                }
            //            }
            //            // TODO: 該当する関数が見つからない時にエラー出す
            //            if (f) {
            //                // 関数内で直接アクセスできるように
            //                if (!_.isUndefined(node.tempParams)) {
            //                    _.forEach(node.tempParams, function(value, key) {
            //                        var valueStr = JSON.stringify(value, undefined, 4);
            //                        var s = key + "=" + valueStr;
            //                        eval(s);
            //                    });
            //                }
            //                
            //                var inputParameters = _.assign({}, parameters);
            //                var outputParameters = f(inputParameters);
            //                // 触らなかったのはそのまま使えるように
            //                parameters = _.defaults(outputParameters, parameters);
            //            }
            //        })();
            //    }
            //}

            try {
                addSubTree(node, index, subTreeName, parameters);
            }
            catch (e) {
                if (_.isUndefined(e.node) || _.isUndefined(e.errorMessage)){
                    throw e;
                }
                //WScript.Echo(JSON.stringify(e, undefined, 4));
                aliasError(e.errorMessage, e.node);
            }
        }
    });

    shrinkChildrenArrayforAllNodes();

    // leaf じゃなくなった node の id を削除
    forAllNodes_Recurse(root, null, -1, function(node, parent, index) {
        //if (node === null) {
        //    return true;
        //}
        if (node.kind !== kindUL) {
            return;
        }
        if (node.children.length === 0) {
            return true;
        }
        if (node.id) {
            delete node.id;
        }
    });

    var endTime = performance.now();
    //WScript.Echo(endTime - startTime);
})();

// imageFilePath をエントリープロジェクトからの相対に変換
forAllNodes_Recurse(root, null, -1, function(node, parent, index) {
    if (_.isUndefined(node.imageFilePath)) {
        return;
    }
    var lineObj = node.lineObj;
    var projectDirectoryFromRoot = lineObj.projectDirectory;
    var fileParentFolderAbs = sourceLocalPathToAbsolutePath(fso.GetParentFolderName(lineObj.filePath), projectDirectoryFromRoot);

    // エントリープロジェクトからの相対パスを求める
    function getImageFilePathFromEntryProject(imageFilePath) {
        if (imageFilePath.charAt(0) != "/") {
            //if (imageDirectory) {
            //    // XXX: imageDirectory の仕様は廃止の方がいい
            //    imageFilePath = fso.BuildPath(imageDirectory, imageFilePath);
            //}
            imageFilePath = fso.BuildPath(fileParentFolderAbs, imageFilePath);
        }
        else {
            imageFilePath = getAbsoluteProjectPath(imageFilePath.slice(1));
        }

        return absolutePathToDirectoryLocalPath(imageFilePath, entryProjectFromRoot);
    }

    node.imageFilePath = getImageFilePathFromEntryProject(node.imageFilePath);

    // TODO: entry file のプロジェクトからの相対パスにする
    //imageFilePath = fso.BuildPath(fso.GetParentFolderName(filePath), image);
    
});


try {

// initial values のデフォルト値の処理
(function() {
    // stack
    // すべて番兵を入れておく
    var columnNames = [{value: []}];
    var defaultValues = [{value: {}}];
    var conditionalColumnValues = [{value: []}];
    forAllNodes_Recurse(root, null, -1,
        function(node, parent, index) {
            if (node.children.length !== 0) {
                if (!_.isUndefined(node.columnNames)) {
                    columnNames.push({
                        node: node,
                        value: node.columnNames
                    });
                }

                if (!_.isUndefined(node.defaultColumnValues)) {
                    // 1個親に自分を上書き追加していく感じで
                    var value = _.clone(_.last(defaultValues).value);
                    // node.defaultColumnValues を value に上書き
                    _.forEach(node.defaultColumnValues, function(val, key) {
                        value[key] = val;
                    });
                    defaultValues.push({
                        node: node,
                        value: value
                    });
                }

                function keyValuePairToObject(conditionalColumnValues) {
                    var result = [];

                    conditionalColumnValues.forEach(function(conditionalColumnValue) {
                        var re = conditionalColumnValue.re;
                        if (!_.isArray(conditionalColumnValue.columnValues)) {
                            result.push({
                                re: re,
                                columnValues: conditionalColumnValue.columnValues
                            });
                            return;
                        }

                        var columnValuesData = conditionalColumnValue.columnValues;
                        var columnValues = {};
                        var currentColumnNames = _.last(columnNames).value;
    
                        columnValuesData.forEach(function(element, index) {
                            if (_.isUndefined(element.key)) {
                                if (index >= currentColumnNames.length) {
                                    var errorMessage = "初期値が列名リストの範囲外に設定されています。";
                                    throw new ColumnValueError(errorMessage, lineObj);
                                }
                                var key = currentColumnNames[index]
                                columnValues[key] = element.value;
                            }
                            else {
                                columnValues[element.key] = element.value;
                            }
                        });
    
                        result.push({
                            re: re,
                            columnValues: columnValues
                        });
                    });

                    return result;
                }

                if (!_.isUndefined(node.conditionalColumnValues)) {
                    // 1個親の先頭に自分を追加していく感じで
                    var value = node.conditionalColumnValues.concat(_.last(conditionalColumnValues).value);
                    conditionalColumnValues.push({
                        node: node,
                        value: keyValuePairToObject(value)
                    });
                }

                // XXX: 本来はエラーとすべきだけど、一旦削除するようにしておく
                // XXX: 仕様変更するかも
                if (!_.isUndefined(node.initialValues)) {
                    delete node.initialValues;
                }

            }
            // leaf の場合
            else {
                (function() {
                    // 指定がなくてもデフォルト値は処理する
                    if (_.isUndefined(node.initialValues)) {
                        node.initialValues = {};                    //    return;
                    }

                    _.last(conditionalColumnValues).value.forEach(function(elem, index) {
                        var columnValues = {};
                        _.forEach(elem.columnValues, function(value, key) {
                            if (!(key in node.initialValues)) {
                                columnValues[key] = value;
                            }
                        });
                        if (_.isEmpty(columnValues)) {
                            return;
                        }
                        if (elem.re.test(node.text)) {
                            _.forEach(columnValues, function(value, key) {
                                node.initialValues[key] = value;
                            });
                        }
                    });

                    _.forEach(_.last(defaultValues).value, function(value, key) {
                        if (!(key in node.initialValues)) {
                            node.initialValues[key] = value;
                        }
                    });

                    // 空文字列なら削除
                    _.keys(node.initialValues).forEach(function(key) {
                        if (node.initialValues[key] === "") {
                            delete node.initialValues[key];
                        }
                    });

                    if (_.isEmpty(node.initialValues)) {
                        delete node.initialValues;
                    }

                })();
            }
        },
        function(node, parent, index) {
            function popSameNode(stack, node) {
                if (_.isEmpty(stack)) {
                    return;
                }
                if (_.last(stack).node === node) {
                    stack.pop();
                }
            }
            popSameNode(conditionalColumnValues, node);
            popSameNode(columnNames, node);
            popSameNode(defaultValues, node);
        }
    );
})();

}
catch (e) {
    (function (errorMessage, lineObj) {
        if (_.isUndefined(lineObj)) {
            MyError(errorMessage);
        }
        else {
            MyError(errorMessage, lineObj.filePath, lineObj.lineNum);
        }
    })(e.errorMessage, e.lineObj);
}

function forAllLeaves_Recurse(node, fun) {
    if (node.children.length === 0) {
        if (fun(node)) {
            return true;
        }
    }

    for (var i = 0; i < node.children.length; i++) {
        forAllLeaves_Recurse(node.children[i], fun);
    }

    return false;
}

function forAllNodes_Recurse(node, parent, index, preChildren, postChildren) {
    if (preChildren(node, parent, index)) {
        return true;
    }

    for (var i = 0; i < node.children.length; i++) {
        if (node.children[i] === null) {
            continue;
        }
        forAllNodes_Recurse(node.children[i], node, i, preChildren, postChildren);
    }

    if (!_.isUndefined(postChildren)) {
        postChildren(node, parent, index);
    }

    return false;
}

// 値が null のプロパティを削除
function deleteNullProperty_Recurse(node) {
    if (node === null) {
        return;
    }
    for (var propertyName in node) {
        if (node[propertyName] === null) {
            delete node[propertyName];
        }
    }

    for (var i = 0; i < node.children.length; i++) {
        deleteNullProperty_Recurse(node.children[i]);
    }
}

// 値が null のプロパティ（場所確保用）を削除する
deleteNullProperty_Recurse(root);

// JSON出力前に不要なプロパティを削除する
CL.deletePropertyForAllNodes(root, "uidList");
CL.deletePropertyForAllNodes(root, "columnNames");
CL.deletePropertyForAllNodes(root, "defaultColumnValues");
CL.deletePropertyForAllNodes(root, "conditionalColumnValues");
CL.deletePropertyForAllNodes(root, "lineObj");
CL.deletePropertyForAllNodes(root, "indent");
CL.deletePropertyForAllNodes(root, "parent");
CL.deletePropertyForAllNodes(root, "subTrees");
CL.deletePropertyForAllNodes(root, "isValidSubTree");
CL.deletePropertyForAllNodes(root, "params");
CL.deletePropertyForAllNodes(root, "tempParams");

forAllNodes_Recurse(root, null, -1, function(node, parent, index) {
    var headers = node.tableHeadersNonInputArea;
    if (!headers) {
        return;
    }
    for (var i = 0; i < headers.length; i++) {
        delete headers[i].marker;
    }
});

//function binToHex(binStr) {
//    var xmldom = new ActiveXObject("Microsoft.XMLDOM");
//    var binObj= xmldom.createElement("binObj");
//
//    binObj.dataType = 'bin.hex';
//    binObj.nodeTypedValue = binStr;
//
//    return String(binObj.text);
//}
// 文字コードを stream.charset にセットする文字列形式で返す
// UTF-8 with BOM, UTF-16 BE, LE のみ判定。それ以外は shift JIS を返す
//function GetCharsetFromTextfile(objSt, path)
//{
//    return "UTF-8";
//
//    objSt.type = adTypeBinary;
//    objSt.Open();
//    objSt.LoadFromFile(path);
//    var bytes = objSt.Read(3);
//    var strBOM = binToHex(bytes);
//    objSt.Close();
//
//    if (strBOM === "efbbbf")
//    {
//        return "UTF-8";
//    }
//
//    strBOM = strBOM.substr(0, 4);
//    if (strBOM === "fffe" || strBOM === "feff")
//    {
//        return "UTF-16";
//    }
//
//    return "Shift_JIS";
//}

function getAbsoluteProjectPath(projectPathFromRoot) {
    return fso.BuildPath(conf.$rootDirectory, projectPathFromRoot);
}

function getAbsoluteDirectory(projectPathFromRoot, directoryName) {
    var projectPathAbs = getAbsoluteProjectPath(projectPathFromRoot);

    if (_.isUndefined(directoryName)) {
        return projectPathAbs;
    }    

    return fso.BuildPath(projectPathAbs, directoryName);
}
function getAbsoluteSourceDirectory(projectPathFromRoot) {
    return getAbsoluteDirectory(projectPathFromRoot, sourceDirectoryName);
}

// project local なファイルパスを絶対パスに変換
function directoryLocalPathToAbsolutePath(filePathProjectLocal, projectPathFromRoot, directoryName) {
    var directoryAbs = getAbsoluteDirectory(projectPathFromRoot, directoryName);

    return fso.BuildPath(directoryAbs, filePathProjectLocal);
}
function sourceLocalPathToAbsolutePath(filePathProjectLocal, projectPathFromRoot) {
    return directoryLocalPathToAbsolutePath(filePathProjectLocal, projectPathFromRoot, sourceDirectoryName);
}
function absolutePathToDirectoryLocalPath(filePath, projectPathFromRoot, directoryName) {
    var directoryAbs = getAbsoluteDirectory(projectPathFromRoot, directoryName);

    return CL.getRelativePath(directoryAbs, filePath);
}
// ソースディレクトリからの相対に変換
function absolutePathToSourceLocalPath(filePath, projectPathFromRoot) {
    return absolutePathToDirectoryLocalPath(filePath, projectPathFromRoot, sourceDirectoryName);
}

function getAbsoluteBackupDirectory(projectPathFromRoot) {
    var projectPathAbs = getAbsoluteProjectPath(projectPathFromRoot);

    return fso.BuildPath(projectPathAbs, backupDirectoryName);
}
function getAbsoluteBackupPath(filePathProjectLocal, projectPathFromRoot) {
    var backupDirectoryAbs = getAbsoluteBackupDirectory(projectPathFromRoot);

    return fso.BuildPath(backupDirectoryAbs, filePathProjectLocal);
}

(function(){

// 先に別名でコピーして、それを読みながら、元ファイルを上書きするように
// 元ファイルをリネームだとエディターで開いてる元ファイルが閉じてしまうので
for (var key in srcTextsToRewrite) {
    var noIdLineData = srcTextsToRewrite[key];
    //printJSON(noIdLineData);
    var filePath = noIdLineData.filePath;
    var projectDirectory = noIdLineData.projectDirectory;
    var filePathAbs = sourceLocalPathToAbsolutePath(filePath, projectDirectory);
    var entryFileFolderName = fso.GetParentFolderName(rootFilePath);
    var folderName = fso.GetParentFolderName(filePath);
    //var backupFolderName = fso.BuildPath(entryFileFolderName, "bak");
    var backupFolderName = getAbsoluteBackupDirectory(projectDirectory);
    backupFolderName = fso.BuildPath(backupFolderName, "txt");

    // 何やってたか忘れたので一旦コメントアウト
    //if (folderName !== entryFileFolderName) {
    //    if (_.startsWith(folderName, entryFileFolderName)) {
    //        var backupSubFolderName = folderName.slice(entryFileFolderName.length + 1);
    //        backupFolderName = fso.BuildPath(backupFolderName, backupSubFolderName);
    //    } else {
    //        // XXX: 何かした方が良いんだろうけど、とりあえず何もしない…
    //    }
    //}

    // 最初から filePath を使えば済む話？
    var fileDirectoryAbs = fso.GetParentFolderName(filePathAbs);
    var fileDirectoryFromSource = absolutePathToSourceLocalPath(fileDirectoryAbs, projectDirectory);
    backupFolderName = fso.BuildPath(backupFolderName, fileDirectoryFromSource);
    CL.createFolder(backupFolderName);

    //var backupPath = getAbsoluteBackupPath(filePath, projectDirectory);
    //alert(filePath + "\n" + projectDirectory + "\n" + backupPath);
    var backupFileName = CL.makeBackupFileName(filePathAbs, fso);
    var backupFilePath = fso.BuildPath(backupFolderName, backupFileName);

    fso.CopyFile(filePathAbs, backupFilePath);

    // バックアップファイルを読んで、元ファイルを直接上書き更新
    var s = CL.readTextFileUTF8(filePathAbs);

    // バックアップファイルを１行ずつ読んで、srcTextsToRewriteに行番号が存在すればそちらを、なければそのまま書き出し
    // XXX: あらかじめ改行でjoinして１回で書き込んだ場合との速度差はどの程度か？
    s = s.split(/\r\n|\n|\r/);
    _.forEach(noIdLineData.newTexts, function(newSrcText, lineNum) {
        s[lineNum] = newSrcText;
    });
    s = s.join("\n");

    CL.writeTextFileUTF8(s, filePathAbs);

    srcTextsToRewrite[key] = null;
}
})();

/**
(function(){
var s = "";
for (var filePath in srcTextsToRewrite)
{
    s += "[ " + filePath + " ]\n";
    for (var lineNum in srcTextsToRewrite[filePath])
    {
        var text = srcTextsToRewrite[filePath][lineNum];
        s += lineNum + ": ";
        s += text + "\n";
    }

    s += "\n";
}
Error(s);
})();
/**/

// TODO: leaf じゃない node に ID がふられてたら無駄なので削除


//function getFileInfo(filePath)
//{
//    var fso = new ActiveXObject("Scripting.FileSystemObject");
//    var file = fso.GetFile(filePath);
//    var info = {
//        fileName: fso.GetFileName(filePath),
//        dateLastModified: new Date(file.DateLastModified).toString()
//    };
//
//    return info;
//}

// TODO: root.id 廃止。 commit, update とかで使ってるので修正範囲は広い
(function() {
    // root.children を基に hash を求める
    var k = JSON.stringify(root.children);
    var shaObj = new jsSHA("SHA-256", "TEXT", { encoding: "UTF8" });
    shaObj.update(k);
    root.id = shaObj.getHash("HEX");
})();


var sJson = JSON.stringify(root, undefined, 4);

// 直列な感じにしてみるテスト
// 全部シートを１つの配列にすると json のサイズは半分ぐらいになるけど jsondiffpatch が簡単にスタック食いつぶすっぽい
// シート内だけを配列にすると json のサイズが2/3ぐらいで、 jsondiffpatch でも良い感じで diff がとれるっぽい。ただし、１シートのnode数が多いとスタック食いつぶす危険性はつねにある
// 恐らくサイズの違いの主な要素はインデント（半角スペース）
/*
sJson = (function () {
    function treeToArray(node, nodes, parentIndex) {
        node.parent = parentIndex;
        parentIndex = nodes.length;
        nodes.push(node);
        for (var i = 0; i < node.children.length; i++) {
            treeToArray(node.children[i], nodes, parentIndex);
        }
        delete node.children;
    }

    var tree = JSON.parse(sJson);
    var children = {};
    var childrenOrder = [];
    for (var i = 0; i < tree.children.length; i++) {
        var nodes = [];
        treeToArray(tree.children[i], nodes, -1);
        var id = nodes[0].id;
        children[id] = nodes;
        childrenOrder.push(id);
    }
    delete tree.children;
    tree.childrenOrder = childrenOrder;
    tree.children = children;

    return JSON.stringify(tree, undefined, 4);
})();
*/

var outFilename = fso.GetBaseName(filePath) + ".json";
var outfilePath = fso.BuildPath(fso.GetParentFolderName(filePath), outFilename);

CL.writeTextFileUTF8(sJson, outfilePath);

if (!runInCScript) {
    WScript.Echo("JSONファイル(" + outFilename + ")を出力しました");
}

var updatedFiles = "";
for (var filePath in srcTextsToRewrite) {
    //updatedFiles += "\n" + fso.GetFileName(filePath);
    updatedFiles += "\n" + filePath;
}

if (updatedFiles !== "") {
    WScript.Echo("以下のソースファイルを更新しました" + updatedFiles);
}

WScript.Quit(0);
