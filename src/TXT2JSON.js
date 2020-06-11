// file の ReadLine(), AtEndOfStream 風の仕様で配列にアクセスするための機構を用意
function ArrayReader(array) { this.__a = array; this.index = 0; this.atEnd = false; }
ArrayReader.prototype.read = function(o) { if (this.atEnd) return null; if (this.index + 1 >= this.__a.length) this.atEnd = true; return this.__a[this.index++]; }

// すべての ID を割り当て直す
var fResetId = false;

var runInCScript = (function() {
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );

    return (fso.getBaseName(WScript.FullName).toLowerCase() === "cscript");
})();

function makeLineinfoString(filePath, lineNum)
{
    var s = "";

    // ファイル名がない時点で終了
    if (typeof filePath === 'undefined')
    {
        return s;
    }

    s += "\nファイル:\t" + filePath;

    if (typeof lineNum === 'undefined')
    {
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
        Error(e.errorMessage);
    }
    else {
        var lineObj = e.lineObj;
        Error(e.errorMessage, lineObj.filePath, lineObj.lineNum);
    }
}

function Error(message, filePath, lineNum) {
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

function createRandomId(len, random)
{
    if (_.isUndefined(random)) {
        random = Math.random;
    } 

    var c = "abcdefghijklmnopqrstuvwxyz";
    c += c.toUpperCase();
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
function createUid(len, idList)
{
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

var shell = new ActiveXObject("WScript.Shell");
var shellApplication = new ActiveXObject("Shell.Application");
var fso = new ActiveXObject( "Scripting.FileSystemObject" );
var stream = new ActiveXObject("ADODB.Stream");

if (( WScript.Arguments.length != 1 ) ||
    ( WScript.Arguments.Unnamed(0) == ""))
{
    Error("チェックリストのソースファイル（.txt）をドラッグ＆ドロップしてください。");
}

var filePath = WScript.Arguments.Unnamed(0);

if (fso.GetExtensionName(filePath) != "txt" && fso.GetExtensionName(filePath) != "md")
{
    Error(".txt ファイルをドラッグ＆ドロップしてください。");
}

// Performance を取得
var htmlfile = WSH.CreateObject("htmlfile");
htmlfile.write('<meta http-equiv="x-ua-compatible" content="IE=Edge"/>');
var performance = htmlfile.parentWindow.performance;
htmlfile.close();

// preprocess
// include とかコメント削除とか
// 入れ子の include にも対応
function preProcess(filePath, filePaths)
{
    var lines = [];
    
    preProcess_Recurse(filePath, lines, filePaths);

    return lines;
}
function preProcess_Recurse(filePath, lines, filePaths)
{
    filePaths.push(filePath);
    var parentFolderName = fso.GetParentFolderName(filePath);

    /**/
    stream.Type = adTypeText;
    // XXX: charset の _autodetect_all が判別に失敗することがある問題がどうにもならないので、超適当に判別
    // 一旦 Shift JIS としてロード
    stream.charset = "Shift_JIS";
    stream.Open();
    stream.LoadFromFile(filePath);
    var allLines = stream.ReadText(adReadAll);
    stream.Close();

    // 先頭の文字が Unicode っぽければロードしなおす
    {
        var charCode0 = allLines.charCodeAt(0);
        var charCode1 = allLines.charCodeAt(1);
        // UTF8 with BOM
        if (charCode0 === 0x30fb && charCode1 === 0xff7f)
        {
            stream.charset = "UTF-8";
        }
        else
        // UTF-16LE, BE
        if (charCode0 === 0xf8f3 && charCode1 === 0xf8f2 ||
            charCode0 === 0xf8f2 && charCode1 === 0xf8f3)
        {
            stream.charset = "UTF-16";
        }
        if (stream.charset !== "Shift_JIS")
        {
            stream.Open();
            stream.LoadFromFile(filePath);
            allLines = stream.ReadText(adReadAll);
            stream.Close();
        }
    }
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

        var include = line.match(/^<<\[\s*(.+)\s*\]$/);
        if (include)
        {
            var path = fso.BuildPath(parentFolderName, include[1]);

            // ファイルが存在するか確認
            (function () {
                var fso = new ActiveXObject("Scripting.FileSystemObject");

                if (!fso.FileExists(path)) {
                    var relativeFilePath = getRelativePath(path, rootFilePath, fso);
                    var errorMessage = "include ファイル\n" + relativeFilePath + "\nが存在しません";

                    Error(errorMessage, filePath, lineArray.index);
                }
            })();
            
            preProcess_Recurse(path, lines, filePaths);
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

    function parseDefine(option, lineObj) {
        if (!currentCondtion()) {
            return;
        }

        var name = option.trim()
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

        var name = option.trim()
        if (!/^([a-zA-Z_]\w*)?$/.test(name)) {
            var errorMessage = "@undef コマンドの文法が正しくありません。";
            throw new ParseError(errorMessage, lineObj);
        }

        if (!isReservedName(name)) {
            //if (name in objs) {
            delete objs[name];
            //}
        }

        // undef の場合は set false 扱い
        parseSet(name + " = false", lineObj);
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

var allFilePaths = [];

var rootFilePath = filePath;
var srcLines = preProcess(filePath, allFilePaths);
srcLines = preProcessConditionalCompile(srcLines);
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

    // UID重複チェック用
    // 「複数人で１つのファイルを作成（ID自動生成）してマージしたら衝突」は考慮しなくて良いぐらいの確率だけど、「IDごとコピペして重複」が高頻度で発生する恐れがあるので
    // JSON出力前に削除する
    uidList: {}
};
stack.push(root);

// IDがふられてないノード
var noIdNodes = {};

// tree 構築後じゃないと leaf かどうかの判別ができないのと、入力済の ID 間での重複チェックをしたいので、貯めといて最後に ID を割り当てる
function AddNoIdNode(node, filePath, lineNum, newSrcText)
{
    if (!(filePath in noIdNodes))
    {
        noIdNodes[filePath] = [];
    }

    var data =
    {
        node: node,
        lineNum: lineNum,

        // 書き換え後の文字列
        // 文字列は {uid} を含むもの
        // 後で uid を生成して{uid}の位置に埋め込む
        newSrcText: newSrcText
    };

    noIdNodes[filePath].push(data);
}

var srcTextsToRewrite = {};

function AddSrcTextToRewrite(filePath, lineNum, newText)
{
    if (!(filePath in srcTextsToRewrite))
    {
        srcTextsToRewrite[filePath] = {};
    }

    srcTextsToRewrite[filePath][lineNum] = newText;
}

function AddChildNode(parent, child)
{
    parent.children.push(child);
    child.parent = parent;
}

// 一番近い親を返す
// 自分が存在する前に使いたい都合上、parentとなるnodeを渡す（渡したnodeも検索対象）仕様
function FindParentNode(parent, fun)
{
    for (; parent; parent = parent.parent)
    {
        if (fun(parent))
        {
            return parent;
        }
    }
    return null;
}

// 一番近い親の uidList を返す
function FindUidList(parent)
{
    var node = FindParentNode(parent, function(node)
    {
        return node.uidList;
    });

    return node ? node.uidList : null;
}

// tableHeaders 内の ID で最小のものが一番左として連番で検索
function getDataFromTableRow(srcData, parentNode, tableHeaderIds)
{
    // data を h1 の tableHeaders の番号に合わせて作り直す
    var data = [];

    // H1は確実に見つかるものとしてOK
    var h1Node = FindParentNode(parentNode, function(node)
    {
        return (node.kind === kindH && node.level === 1);
    });

    if (typeof tableHeaderIds === 'undefined')
    {
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
        for (var i = 0; i < srcData.length; i++)
        {
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

//  ファイルの文字データを一行ずつ読む
while (!srcLines.atEnd) {
    var lineObj = srcLines.read();
    var line = lineObj.line;

    var h = line.match(/^(#+)\s+(.*)$/);
    if (h) {
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
                if (uid in uidListH1)
                {(function(){
                    var uidInfo0 = uidListH1[uid];
                    var errorMessage = "ID '#" + uid + "' が重複しています";
                    errorMessage += makeLineinfoString(uidInfo0.filePath, uidInfo0.lineNum);
                    errorMessage += makeLineinfoString(lineObj.filePath, lineObj.lineNum);
                    Error(errorMessage);
                })();}
                else {
                    uidListH1[uid] = lineObj;
                }
            }

            // シート内での重複だけ確認したいのでここでクリア
            uidList = {};

            tableHeaders = [];
        }
        else {
            while (/.* {2}$/.test(text)) {
                // 改行の次の行の行頭のスペースは無視するように
                // 厳密にはインデントが揃ってるかちゃんとみるべきだけど、そこまでやるつもりはない
                line = _.trimLeft(srcLines.read().line);
                text = text.slice(0, -2) + "\n" + line;
            }

            // １行のみ、行全体以外は対応しない
            var link = text.trim().match(/^\[(.+)\]\((.+)\)$/);
            if (link) {
                text = link[1].trim();
                url = link[2].trim();
            }
        }

        if (text.length > 31) {
            var errorMessage = "シート名が31文字を超えています";
            Error(errorMessage, lineObj.filePath, lineObj.lineNum);
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
            Error(errorMessage, lineObj.filePath, lineObj.lineNum);
        }
        if (_.find(root.children, function(item) {
            return item.text == text;
        })) {
            var errorMessage = "シート名「" + text + "」はすでに使われています";
            Error(errorMessage, lineObj.filePath, lineObj.lineNum);
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

            AddNoIdNode(item, lineObj.filePath, lineObj.lineNum, newSrcText);
        }

        continue;
    }

    // 行頭に全角スペースがないかのチェック
    (function () {
        var fullwidthSpaceMatch = line.match(/^([\s　]+).*$/);
        if (!fullwidthSpaceMatch) {
            return;
        }
        var regex = /　/g;
        if (regex.test(fullwidthSpaceMatch[1]))
        {
            var errorMessage = "行頭に全角スペースが含まれています";
            Error(errorMessage, lineObj.filePath, lineObj.lineNum);
        }
    })();
    // # とか - とか 1. の後ろにスペースがないかのチェック
    function checkSpaceAfterMark(re) {
        var spaceMatch = line.match(re);
        if (!spaceMatch) {
            return;
        }
        var regex = /^\s+/;
        if (!regex.test(spaceMatch[1]))
        {
            var errorMessage = "行頭の記号の後ろに半角スペースが必要です";
            Error(errorMessage, lineObj.filePath, lineObj.lineNum);
        }
    }
    checkSpaceAfterMark(/^#+(.+)$/);
    checkSpaceAfterMark(/^\s*[\*\+\-]\.?(.+)$/);
    checkSpaceAfterMark(/^\s*\d+\.(.+)$/);

    var ul = line.match(/^(\s*)([\*\+\-])\s+(.*)$/);
    if (ul) {
        var indent = ul[1].length;
        var text = ul[3];
        var marker = ul[2];

        while (stack.peek().kind == kindUL && stack.peek().indent >= indent) {
            stack.pop();
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
                    Error(errorMessage);
                }
                else {
                    uidList[uid] = lineObj;
                }
            })();}
        }

        // TODO: leaf 以外で initialValues が設定されていたらエラー出す
        var initialValues = void 0;
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

        while (/.* {2}$/.test(text)) {
            // 改行の次の行の行頭のスペースは無視するように
            // 厳密にはインデントが揃ってるかちゃんとみるべきだけど、そこまでやるつもりはない
            line = _.trimLeft(srcLines.read().line);
            text = text.slice(0, -2) + "\n" + line;
        }

        var commentMatch = text.trim().match(/^([\s\S]+)\s*\[\^(.+)\]$/);
        var comment = undefined;
        if (commentMatch) {
            text = commentMatch[1].trim();
            comment = commentMatch[2].trim();
            comment = comment.replace(/<br>/gi, "\n");
        }

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
                Error(errorMessage, lineObj.filePath, lineObj.lineNum);
            }
        }

        // １行のみ、行全体以外は対応しない
        var link = text.trim().match(/^\[(.+)\]\((.+)\)$/);
        var url = undefined;
        if (link) {
            text = link[1].trim();
            url = link[2].trim();
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
            initialValues: initialValues,
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

            AddNoIdNode(item, lineObj.filePath, lineObj.lineNum, newSrcText);
        }

        continue;
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

            while (/.* {2}$/.test(text)) {
                // 改行の次の行の行頭のスペースは無視するように
                // 厳密にはインデントが揃ってるかちゃんとみるべきだけど、そこまでやるつもりはない
                line = _.trimLeft(srcLines.read().line);
                text = text.slice(0, -2) + "\n" + line;
            }

            if (parent.kind === kindH && parent.level === 1) {
                var comment = undefined;
                var commentMatch = text.trim().match(/^([\s\S]+)\s*\[\^(.+)\]$/);
                if (commentMatch) {
                    text = commentMatch[1].trim();
                    comment = commentMatch[2].trim();
                    if (/<br>/gi.test(comment)) {
                        var errorMessage = "確認欄のコメントでは改行は使えません";
                        Error(errorMessage, lineObj.filePath, lineObj.lineNum);
                    }
                    // Excel の仕様で、入力時メッセージのタイトルは31文字まで
                    if (comment.length > 32) {
                        var errorMessage = "確認欄のコメントが32文字を超えています";
                        Error(errorMessage, lineObj.filePath, lineObj.lineNum);
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
                    Error(errorMessage, lineObj.filePath, lineObj.lineNum);
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
                    Error(errorMessage, lineObj.filePath, lineObj.lineNum);
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
            Error("番号付きリストは H1 の直下以外には作れません");
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
            Error(errorMessage, topLineObj.filePath, topLineObj.lineNum);
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
                            Error(errorMessage);
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
                    Error(errorMessage, lineObj.filePath, lines[i].lineNum);
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

                    AddNoIdNode(parentNode, lineObj.filePath, lines[y0].lineNum, newSrcText);
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
                            Error(errorMessage);
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
                        Error(errorMessage, lineObj.filePath, lines[i].lineNum);
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
    if (property)
    {
        stack.peek().variables[property[1]] = property[2];
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
        Error(errorMessage, lineObj.filePath, lineObj.lineNum);
    }
});

(function(){
    for (var filePath in noIdNodes) {
        // id を付与してファイルに書き出すノードを抽出
        var infos = noIdNodes[filePath].filter(function(element, index, array) {
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

        for (var i = 0; i < infos.length; i++) {
            var info = infos[i];
            var node = info.node;
            var uidList = FindUidList(node.parent);
            var uid = createUid(8, uidList);
            node.id = uid;

            // ID 挿入して書き換え
            var newSrcText = info.newSrcText.format({uid: uid});

            AddSrcTextToRewrite(filePath, info.lineNum, newSrcText);
        }
    }
})();

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
                Error(errorMessage);
            }
            else {
                Error(errorMessage, lineObj.filePath, lineObj.lineNum);
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
            Error(errorMessage);
        }
        else {
            Error(errorMessage, lineObj.filePath, lineObj.lineNum);
        }
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

    // すべての alias を tree から取り外し、所属 node にリストアップしておく
    forAllNodes_Recurse(root, null, -1, function(node, parent, index) {
        if (parent === null) {
            return;
        }
        if (node.kind !== kindUL) {
            return;
        }

        var match = node.text.trim().match(/^&([A-Za-z_]\w*)\(.*\)$/);
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

        return true;
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
                        Error(errorMessage, lineObj.filePath, lineObj.lineNum);
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
    forAllNodes_Recurse(root, null, -1, function(node, parent, index) {
        if (_.isUndefined(node.subTrees)) {
            return;
        }

        _.forEach(node.subTrees, function(subTreeRoot, name) {

            verifyReference(subTreeRoot);

            forAllNodes_Recurse(subTreeRoot, null, -1, function(node, parent, index) {
                if (node === null) {
                    return true;
                }
                var match = node.text.trim().match(/^\*([A-Za-z_]\w*)\(.*\)$/);
                if (match === null) {
                    return;
                }
                var subTreeName = match[1];
                addSubTree(node, index, subTreeName);
            });

            shrinkChildrenArray(subTreeRoot, null, -1);
        });
    });

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
    function addSubTree(targetNode, targetIndex, subTreeName) {
        var subTree = findSubTree_Recurse(subTreeName, targetNode.parent);

        // みつからなかった
        if (subTree === null) {
            var errorMessage = "エイリアス'" + subTreeName + "'は存在しません。";
            throw new AliasError(errorMessage, targetNode);
        }

        // まず clone
        subTree = cloneSubTree(subTree);
        // XXX: node に循環参照があるので JSON.stringify は使えない
        //subTree = JSON.parse(JSON.stringify(subTree));

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
            forAllNodes_Recurse(targetClone, null, -1, function(node, parent, index) {
                if (parent === null) {
                    return;
                }
                var match = node.text.trim().match(/^\*([A-Za-z_]\w*)\(.*\)$/);
                if (match !== null) {
                    var subTreeName = match[1];
                    addSubTree(node, index, subTreeName);
                }
            });

            forAllNodes_Recurse(subTree, null, -1, function(node, parent, index) {
                if (node.children.length > 0) {
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
        var match = node.text.trim().match(/^\*([A-Za-z_]\w*)\(.*\)$/);
        if (match !== null) {
            var subTreeName = match[1];

            try {
                addSubTree(node, index, subTreeName);
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
CL.deletePropertyForAllNodes(root, "lineObj");
CL.deletePropertyForAllNodes(root, "indent");
CL.deletePropertyForAllNodes(root, "parent");
CL.deletePropertyForAllNodes(root, "subTrees");
CL.deletePropertyForAllNodes(root, "isValidSubTree");

forAllNodes_Recurse(root, null, -1, function(node, parent, index) {
    var headers = node.tableHeadersNonInputArea;
    if (!headers) {
        return;
    }
    for (var i = 0; i < headers.length; i++) {
        delete headers[i].marker;
    }
});

function binToHex(binStr)
{
    var xmldom = new ActiveXObject("Microsoft.XMLDOM");
    var binObj= xmldom.createElement("binObj");

    binObj.dataType = 'bin.hex';
    binObj.nodeTypedValue = binStr;

    return String(binObj.text);
}
// 文字コードを stream.charset にセットする文字列形式で返す
// UTF-8 with BOM, UTF-16 BE, LE のみ判定。それ以外は shift JIS を返す
function GetCharsetFromTextfile(objSt, path)
{
    objSt.type = adTypeBinary;
    objSt.Open();
    objSt.LoadFromFile(path);
    var bytes = objSt.Read(3);
    var strBOM = binToHex(bytes);
    objSt.Close();

    if (strBOM === "efbbbf")
    {
        return "UTF-8";
    }

    strBOM = strBOM.substr(0, 4);
    if (strBOM === "fffe" || strBOM === "feff")
    {
        return "UTF-16";
    }

    return "Shift_JIS";
}

(function(){
// 先に別名でコピーして、それを読みながら、元ファイルを上書きするように
// 元ファイルをリネームだとエディターで開いてる元ファイルが閉じてしまうので
var streamOut = new ActiveXObject("ADODB.Stream");
for (var filePath in srcTextsToRewrite)
{
    var rootFileFolderName = fso.GetParentFolderName(rootFilePath);
    var folderName = fso.GetParentFolderName(filePath);
    var backupFolderName = fso.BuildPath(rootFileFolderName, "bak");
    backupFolderName = fso.BuildPath(backupFolderName, "txt");
    if (folderName !== rootFileFolderName) {
        if (_.startsWith(folderName, rootFileFolderName)) {
            var backupSubFolderName = folderName.slice(rootFileFolderName.length + 1);
            backupFolderName = fso.BuildPath(backupFolderName, backupSubFolderName);
        } else {
            // XXX: 何かした方が良いんだろうけど、とりあえず何もしない…
        }
    }
    CL.createFolder(backupFolderName);

    var backupFileName = CL.makeBackupFileName(filePath, fso);
    var backupFilePath = fso.BuildPath(backupFolderName, backupFileName);

    fso.CopyFile(filePath, backupFilePath);

    // バックアップファイルを読んで、元ファイルを直接上書き更新
    var inputCharset = GetCharsetFromTextfile(stream, backupFilePath);
    stream.Type = adTypeText;
    stream.charset = inputCharset;
    stream.Open();
    stream.LoadFromFile(filePath);

    streamOut.Type = adTypeText;
    // 元ファイルの文字コードによらず、UTF-8 with BOM固定で
    streamOut.charset = "utf-8";
    streamOut.Open();

    // バックアップファイルを１行ずつ読んで、srcTextsToRewriteに行番号が存在すればそちらを、なければそのまま書き出し
    // XXX: あらかじめ改行でjoinして１回で書き込んだ場合との速度差はどの程度か？
    for (var lineNum = 1; !stream.EOS; lineNum++)
    {
        var text = stream.ReadText(adReadLine);
        if (lineNum in srcTextsToRewrite[filePath])
        {
            text = srcTextsToRewrite[filePath][lineNum];
            delete srcTextsToRewrite[filePath][lineNum];
        }
        streamOut.WriteText(text, adWriteLine);
    }

    stream.Close();

    streamOut.SaveToFile(filePath, adSaveCreateOverWrite);
    streamOut.Close();
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


function getFileInfo(filePath)
{
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var file = fso.GetFile(filePath);
    var info = {
        fileName: fso.GetFileName(filePath),
        dateLastModified: new Date(file.DateLastModified).toString()
    };

    return info;
}

var allFileInfo = [];
for (var i = 0; i < allFilePaths.length; i++)
{
    allFileInfo.push(getFileInfo(allFilePaths[i]));
}

// txt2json 自身がソースファイルを更新するので、最後に
root.sourceFiles = allFileInfo;

var srcTextLastModifiedDate = (function(){
    var maxTime = 0;
    var newestDate = null;
    for (var i = 0; i < root.sourceFiles.length; i++)
    {
        var date = new Date(root.sourceFiles[i].dateLastModified);
        var time = date.getTime();
        if (time > maxTime)
        {
            maxTime = time;
            newestDate = date;
        }
    }
    return newestDate;
})();

(function() {
    var seed = srcTextLastModifiedDate.getTime();
    var random = createXor128(seed);

    // 特に意味はないけど、通常の id のより長めにしておく
    root.id = createRandomId(16, random);
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

var fso = new ActiveXObject( "Scripting.FileSystemObject" );


//  ファイルを書き込み専用で開く
var outFilename = fso.GetBaseName(filePath) + ".json";
var outfilePath = fso.BuildPath(fso.GetParentFolderName(filePath), outFilename);
/**/
stream.Type = adTypeText;
stream.charset = "utf-8";
stream.Open();

var jsonArray = sJson.split("\n");
for (var i = 0; i < jsonArray.length; i++)
{
    stream.WriteText(jsonArray[i], adWriteLine);
}

stream.SaveToFile(outfilePath, adSaveCreateOverWrite);
stream.Close();
/*/
var outfile = fso.OpenTextFile(outfilePath, FORWRITING, true, TRISTATE_FALSE);

var jsonArray = sJson.split("\n");
for (var i = 0; i < jsonArray.length; i++)
{
    outfile.WriteLine(jsonArray[i]);
}

//  ファイルを閉じる
outfile.Close();
/**/

if (!runInCScript) {
    WScript.Echo("JSONファイル(" + outFilename + ")を出力しました");
}

var updatedFiles = "";
for (var filePath in srcTextsToRewrite)
{
    updatedFiles += "\n" + fso.GetFileName(filePath);
}

if (updatedFiles !== "")
{
    WScript.Echo("以下のソースファイルを更新しました" + updatedFiles);
}

WScript.Quit(0);
