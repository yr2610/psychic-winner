function alert(s) {
    WScript.Echo(s);
}

function printJSON(json) {
    alert(stringifyPretty(json));
}

var DROP_KEYS_LIST = [
    "uidList",
    "columnNames",
    "defaultColumnValues",
    "conditionalColumnValues",
    "lineObj",
    "indent",
    "parent",
    "templates",
    "isValidSubTree",
    "params",

    "$args",
    "$params",

    "$get",
    "$set",
    "$defaults",

    "$index1",
    "$count",
    "$isFirst",
    "$isLast",
    "$isOdd",
    "$isEven",

    "marker"    // 削除済みだけど一応
];

function toKeySet(keys) {
  var set = {};
  for (var i = 0; i < keys.length; i++) {
    set[keys[i]] = 1;
  }
  return set;
}

var DEFAULT_DROP_KEYS = toKeySet(DROP_KEYS_LIST);

// --- numeric repeat sugar (lodash 3.10.1 前提) ---
var ENABLE_NUMERIC_REPEAT = (conf && conf.ENABLE_NUMERIC_REPEAT === false) ? false : true;
var MAX_REPEAT = (conf && conf.MAX_REPEAT) || 500;

function toRepeatList(arg){
  if (_.isArray(arg)) return arg;               // 既存の配列はそのまま
  if (!ENABLE_NUMERIC_REPEAT) return null;

  if (typeof arg === "number"){                 // *Foo(5)
    var n = Math.max(0, Math.min(arg, MAX_REPEAT));
    return _.times(n, function(){ return { $value: undefined }; });
  }
  if (arg && typeof arg.$times === "number"){   // *Foo({ $times: 5, ... })
    var m = Math.max(0, Math.min(arg.$times, MAX_REPEAT));
    return _.times(m, _.constant(arg));         // 同一参照でOK（必要なら _.clone に変更）
  }
  return null;
}

function buildLoopMeta(k, total){
  return {
    $index:  k,
    $index1: k + 1,
    $count:  total,
    $isFirst: (k === 0),
    $isLast:  (k === total - 1),
    $isOdd:   ((k % 2) === 1),
    $isEven:  ((k % 2) === 0)
  };
}

// 速い replacer（循環は parent/マップ系を落として断つ）
function makeFastJSONReplacer(dropKeys) {
  var set = dropKeys
    ? (typeof dropKeys.length === "number" ? toKeySet(dropKeys) : dropKeys)
    : DEFAULT_DROP_KEYS;

  return function replacer(key, value) {
    if (typeof key === "string" && key !== "") {
      // 1) 明示ドロップ（循環を断つのが最重要）
      if (set[key]) return undefined;

      // 2) "__" / "$$" 接頭辞を高速判定（regex使わない）
      var c0 = key.charCodeAt(0);
      if (c0 === 95) { // '_'
        if (key.length > 1 && key.charCodeAt(1) === 95) return undefined; // "__"
      } else if (c0 === 36) { // '$'
        if (key.length > 1 && key.charCodeAt(1) === 36) return undefined; // "$$"
      }
    }
    if (typeof value === "function") return undefined;
    return value;
  };
}

var JSON_REPLACER = makeFastJSONReplacer();

// デフォは2スペース。必要に応じて "\t" や 4 に変えてOK
function stringifyPretty(obj, indent) {
    if (indent === undefined) {
        indent = 2;
    }
    return JSON.stringify(obj, JSON_REPLACER, indent);
}

function countLeadingSpaces(s) {
    var match = String(s).match(/^\s*/);
    return match ? match[0].length : 0;
}

function stripTrailingSpaces(text) {
    return text.replace(/[ \t]+$/, "");
}

function appendLineComment(text, lineObj) {
    if (!_.isUndefined(lineObj.comment)) {
        text += lineObj.comment;
    }
    return text;
}

function buildProjectKey(projectDirectory, filePath) {
    return projectDirectory + ":" + filePath;
}

function ensureMapEntry(store, key, initializer) {
    if (!(key in store)) {
        store[key] = initializer();
    }
    return store[key];
}

// ==== 継続吸収（" +" とハンギング）====
// reader: srcLines の reader
// text:   先頭行の本文
// baseline: 本文開始より右なら継続（UL: indent+2, OL: leading+2）
// options: { trimLeft: true, structuralGuard: true, stripRight: true }
function absorbContinuations(reader, text, baseline, options) {
    options = options || {};
    var trimLeft = options.trimLeft !== false;
    var useGuard = options.structuralGuard !== false;
    var stripRight = options.stripRight !== false;

    function unreadOne() {
        if (reader.index > 0) {
            reader.index--;
            reader.atEnd = false;
        }
    }

    function isStructuralStart(s) {
        if (!useGuard) return false;
        var t = String(s).trim();
        return /^#{1,6}\s+/.test(t)      // 見出し
            || /^\s*[\*\+\-]\s+/.test(s) // UL
            || /^\s*\d+\.\s+/.test(s)    // OL
            || /^&[A-Za-z_]\w*\s*\(/.test(t) // &Name( 宣言
            || /^\*[A-Za-z_]\w*\s*\(/.test(t) // *Call(
            || /^@[A-Za-z_]\w*:/.test(t)    // ディレクティブ（@xxx:）
            || /^\s*\[.+\]:\s+.+$/.test(s); // 属性宣言（[key]: value）
    }

    // ---- 1) 既存の " +" 明示継続 ----
    var usedPlus = false;
    // 先頭行が " +" で継続指定されているかを検出
    if (/[ \t]+\+[ \t]*$/.test(text)) {
        usedPlus = true;

        // 末尾の " +" を落とす（改行は触らない）
        text = text.replace(/[ \t]+\+[ \t]*$/, "");

        // 以降は “プラス継続モード” で後続行を吸収する
        var modePlus = true;
        while (modePlus && !reader.atEnd) {
            var rec = reader.read();
            var raw = rec ? rec.line : "";
            var trimmed = String(raw).trim();

            // 1-a) 空行マーカー: 「+」のみ（前後に半角スペース可）
            //     → 空行を 1 行挿入し、継続は継続（次の行も読む）
            //     （オプション化したい場合は options.plusBlankMarker !== false 等で分岐）
            if (trimmed === "+") {
                if (stripRight) {
                    text = stripTrailingSpaces(text); // 直前行の末尾空白だけ除去
                }
                text += "\n"; // 空行を挿入
                modePlus = true;
                continue;
            }

            // 1-b) 本当の空行（空白のみ or 完全な空行）は“継続終了”
            //     → 空行は消費しない（元仕様どおり）。必要ならここで append しても良い。
            if (trimmed.length === 0) {
                unreadOne();
                break;
            }

            // 1-c) 通常の継続行：その行自身が " +" で終わるなら更に継続
            var hasPlus = /[ \t]+\+[ \t]*$/.test(raw);
            var content = hasPlus ? raw.replace(/[ \t]+\+[ \t]*$/, "") : raw;

            var piece = trimLeft ? _.trimLeft(content) : content;
            if (stripRight) {
                piece = stripTrailingSpaces(piece);
            }

            text += "\n" + piece;
            modePlus = hasPlus; // この行が " +" で終わっていれば、次も読む
        }
    }

    // ---- 2) ハンギング継続（" +" を使っていない項目のみ許可）----
    if (!usedPlus) {
        while (!reader.atEnd) {
            var rec = reader.read();
            var s = rec ? rec.line : "";

            if (isStructuralStart(s)) {
                unreadOne();
                break;
            }

            var ls = countLeadingSpaces(s);
            if (s.trim().length === 0 || ls >= baseline) {
                var seg = trimLeft ? _.trimLeft(s) : s;
                if (stripRight) {
                    seg = stripTrailingSpaces(seg);
                }
                // 直前行の末尾空白だけ削る（改行は保持＝空行は維持）
                if (stripRight) {
                    text = stripTrailingSpaces(text);
                }
                text += "\n" + seg;
                continue;
            }

            unreadOne();
            break;
        }
    }

    // ---- 3) 念のため：全行の行末スペースのみ掃除（改行は保持）----
    if (stripRight) {
        text = stripTrailingSpaces(text.replace(/[ \t]+(\r?\n)/g, "$1"));
    }

    return text;
}


var shell = null;
var shellApplication = null;
var fso = null;
var stream = null;
var filePath = (WScript && WScript.Arguments && WScript.Arguments.length > 0)
  ? WScript.Arguments.Unnamed(0) : "";

var conf = null;

function setupEnvironment(force) {
    if (force || !shell) shell = new ActiveXObject("WScript.Shell");
    if (force || !shellApplication) shellApplication = new ActiveXObject("Shell.Application");
    if (force || !fso) fso = new ActiveXObject("Scripting.FileSystemObject");
    if (force || !stream) stream = new ActiveXObject("ADODB.Stream");
}

function showPopup(message, secondsToWait, title, type) {
    if (_.isUndefined(secondsToWait)) secondsToWait = 0;
    if (_.isUndefined(title)) title = "";
    if (_.isUndefined(type)) type = 0;

    for (var attempt = 0; attempt < 2; attempt++) {
        try {
            setupEnvironment(attempt > 0);
            return shell.Popup(message, secondsToWait, title, type);
        } catch (err) {
            shell = null;
            if (attempt === 1) {
                throw err;
            }
        }
    }
}

function parseArgs() {
    if (( WScript.Arguments.length != 1 ) ||
        ( WScript.Arguments.Unnamed(0) == "")) {
        MyError("チェックリストのソースファイル（.txt）をドラッグ＆ドロップしてください。");
    }
    return WScript.Arguments.Unnamed(0);
}

function loadGlobalConfig() {

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

}

function computeRootId() {

    // root.children を基に hash を求める
    //var k = JSON.stringify(root.children);
    var k = _.values(srcTexts).join("\n");
    
    //var startTime = performance.now();
//    var shaObj = new jsSHA("SHA-256", "TEXT", { encoding: "UTF8" });
//    shaObj.update(k);
//    root.id = shaObj.getHash("HEX");
    root.id = getMD5Hash(k);
    //var endTime = performance.now();
    //alert(endTime - startTime);

}

// file の ReadLine(), AtEndOfStream 風の仕様で配列にアクセスするための機構を用意
function ArrayReader(array) { this.__a = array; this.index = 0; this.atEnd = false; }
ArrayReader.prototype.read = function(o) { if (this.atEnd) return null; if (this.index + 1 >= this.__a.length) this.atEnd = true; return this.__a[this.index++]; }

// すべての ID を割り当て直す
var fResetId = false;

var runInCScript = (function() {
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );

    return (fso.getBaseName(WScript.FullName).toLowerCase() === "cscript");
})();

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
        showPopup(message, 0, "エラー", ICON_EXCLA);
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

setupEnvironment();

filePath = parseArgs();

if (fso.GetExtensionName(filePath) != "txt") {
    MyError(".txt ファイルをドラッグ＆ドロップしてください。");
}

var outFilename = fso.GetBaseName(filePath) + ".json";
var outfilePath = fso.BuildPath(fso.GetParentFolderName(filePath), outFilename);

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
loadGlobalConfig();

var confFileName = "conf.yml";
(function() {
    var baseName = fso.GetBaseName(filePath);
    baseName = baseName.replace(/_index$/, "");
    if (baseName != "index") {
        confFileName = baseName + "_" + confFileName;
    }
})();
conf = readConfigFile(confFileName);

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
(function() {
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
    var key = buildProjectKey(projectDirectory, filePath);
    var list = ensureMapEntry(noIdNodes, key, function() { return []; });

    var data = {
        node: node,

        lineObj: lineObj,

        lineNum: lineNum,

        // 書き換え後の文字列
        // 文字列は {uid} を含むもの
        // 後で uid を生成して{uid}の位置に埋め込む
        newSrcText: newSrcText
    };

    list.push(data);
}

var srcTextsToRewrite = {};

function AddSrcTextToRewrite(noIdLineData, newSrcText) {
    var lineObj = noIdLineData.lineObj;
    var filePath = lineObj.filePath;
    var projectDirectory = lineObj.projectDirectory;
    //var filePathProjectLocal = getProjectLocalPath(filePath, projectDirectory);
    var key = buildProjectKey(projectDirectory, filePath);
    var lineNum = noIdLineData.lineNum;

    var entry = ensureMapEntry(srcTextsToRewrite, key, function() {
        return {
            filePath: filePath,
            projectDirectory: projectDirectory,
            newTexts: {}
        };
    });

    var newTexts = entry.newTexts;
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
                (function() {
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
        var baseline = indent + 2; // 記号＋スペースぶん
        text = absorbContinuations(srcLines, text, baseline, { trimLeft: true, structuralGuard: true });

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
        srcHash: null,  // 場所確保
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

        newSrcText = appendLineComment(newSrcText, lineObj);

        AddNoIdNode(item, lineObj, lineObj.lineNum, newSrcText);
    }

    return h;
}

function parseUnorderedList(lineObj, line) {
    if (typeof line === "undefined") {
        line = lineObj.line;
    }
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
        {(function() {
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

    var leading = countLeadingSpaces(line);
    var baseline = leading + 2; // "n. " の見た目の最低値でOK
    text = absorbContinuations(srcLines, text, baseline, { trimLeft: true, structuralGuard: true });

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

        newSrcText = appendLineComment(newSrcText, lineObj);

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
        if (parseUnorderedList(lineObj, line)) {
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
        (function() {
            var number = parseInt(ol[1], 10);
            var text = ol[2];
            var parent = stack.peek();

            // baseline は「その行の本文先頭の桁」。
            // プレーン行なら先頭空白数 + 1 くらいでも可（必要なら厳密化）。
            var leading = countLeadingSpaces(line);
            var baseline = leading + 1;
            text = absorbContinuations(srcLines, text, baseline, { trimLeft: true, structuralGuard: true });

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
        function convertFunctions(o) {
            _.forEach(o, function (v, k) {
                if (_.isString(v) && v.match(/^function/)) {
                    o[k] = Function.call(this, 'return ' + v)();
                } else if (_.isObject(v) && !_.isArray(v)) {
                    convertFunctions(v); // 再帰的に処理
                }
            });
        }
        convertFunctions(o);
        //_.forEach(o, function (v, k) {
        //    if (_.isString(v) && v.match(/^function/)) {
        //        o[k] = Function.call(this, 'return ' + v)();
        //    }
        //});

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
        // ただし次の2種は除外
        // 1) &Name: 宣言（アンカー定義） …… 例: &foo:
        // 2) ディレクティブ @xxx:         …… 例: @init:, @set: など将来拡張含む
        if (element.node.children.length === 0) {
            var t = element.node.text;
            return !/^(?:&[A-Za-z_]\w*:|@[A-Za-z_]\w*:)/.test(t);
        }

        // テンプレート参照ノード
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
        node.lineObj.line = newSrcText;
    });
});

var lastParsedRoot;

(function() {
    // 前回出力したJSONファイルがあれば読む
    if (!fso.FileExists(outfilePath)) {
        return;
    }

    var s = CL.readTextFileUTF8(outfilePath);
    //var startTime = performance.now();
    //lastParsedRoot = JSON.parse(s);
    //var endTime = performance.now();
    //alert(endTime - startTime);

    // parse できるものを parse するならこちらの方が全然速い
    function parseJSON(str) {
        if (str === "") str = '""';
        eval("var p=" + str + ";");
        return p;
    }
    lastParsedRoot = parseJSON(s);
})();

// 「byte配列」から「16進数文字列」
function bytes2hex(bytes) {
    var hex = null;
    // 「DOMDocument」生成
    var doc = new ActiveXObject("Msxml2.DOMDocument");
    // 「DomNode」生成（hex）
    var element = doc.createElement("hex");
    // 「dataType」に「bin.hex」を設定
    element.dataType = "bin.hex";
    // 「nodeTypedValue」に「byte配列」を設定
    element.nodeTypedValue = bytes;
    // 「text」を取得
    hex = element.text;
    // 後処理
    element = null;
    doc = null;
    return hex;
}

function getHash(crypto, input) {
    var encoding = new ActiveXObject("System.Text.UTF8Encoding");
    var bytes = encoding.GetBytes_4(input);
    var hash = crypto.ComputeHash_2(bytes);
    return bytes2hex(hash);
}
function getMD5Hash(input) {
    var crypto = new ActiveXObject("System.Security.Cryptography.MD5CryptoServiceProvider");
    return getHash(crypto, input);
}
function getSHA1Hash(input) {
    var crypto = new ActiveXObject("System.Security.Cryptography.SHA1CryptoServiceProvider");
    return getHash(crypto, input);
}

// preprocess 後、 id 付与後のソーステキストをシートごとにhashで持っておく
var parsedSheetNodeInfos = [];
var reusedSheetNames = {};
var srcTexts;   // XXX: root.id 用に保存しておく…
(function() {
    var children = root.children;
    var src = srcLines.__a;
    var result = {};
    for (var i = 0; i < children.length; i++) {
        var start = src.indexOf(children[i].lineObj);
        var end = (i + 1 < children.length) ? src.indexOf(children[i + 1].lineObj) : src.length;
        var lines = [];
        for (var j = start; j < end; j++) {
            lines.push(src[j].line);
        }
        result[children[i].id] = lines.join("\n");
    }
    srcTexts = result;

    // root には存在せず lastParsedRoot には存在するノードを抽出
    var removedNodesFromLastParse;
    if (lastParsedRoot) {
        removedNodesFromLastParse = _.filter(lastParsedRoot.children, function(node) {
            return !_.some(root.children, function(rootNode) { return rootNode.id === node.id; });
        });
    }

    _.forEach(root.children, function(v, index) {
        var srcSheetText = result[v.id];

        //v.srcHash = getSHA1Hash(srcSheetText);
        v.srcHash = getMD5Hash(srcSheetText);

        function getParsedSheetNode(sheetNode) {
            if (!lastParsedRoot) {
                return;
            }
            var parsedSheetNode = _.find(lastParsedRoot.children, { id: sheetNode.id });
            if (!parsedSheetNode) {
                return;
            }
            if (parsedSheetNode.srcHash && parsedSheetNode.srcHash == sheetNode.srcHash) {
                return parsedSheetNode;
            }
        }

        var parsedSheetNode = getParsedSheetNode(v);
        var sheetNameForWarnings = getPlaceholderWarningSheetName(v);

        // srcHash が同じ sheetNode があれば、そのまま再利用
        if (shouldReuseParsedSheetNode(parsedSheetNode, sheetNameForWarnings)) {
            var info = {
                index: index,
                node: parsedSheetNode
            };
            parsedSheetNodeInfos.push(info);
            reusedSheetNames[sheetNameForWarnings] = true;
            root.children[index] = null;
        }
    });

    // 一旦削除する
    // 「parsedSheet に置き換えする node は処理しない」というのをすべての処理に入れるというのは修正コストが高すぎるので
    root.children = root.children.filter(function(node) {
        return node != null;
    });

    // 更新の場合はメッセージを表示
    (function () {
        if (!lastParsedRoot) {
            // 完全新規っぽい場合は何も表示しない
            return;
        }

        var message = "";

        if (removedNodesFromLastParse.length > 0) {
            message += "次のシートが削除されました。\n\n";
            var removedNodesString = _.map(removedNodesFromLastParse, function(sheetNode) {
                return '* ' + sheetNode.text;
            }).join('\n');
            message += removedNodesString;
            message += "\n\n";
        }

        // キャンセル時には一般的によく使われるとされている値を返しておく
        // 1: 一般的なエラー
        // 2: コマンドライン引数のエラー
        // 3: ファイルが見つからない
        // 4: アクセス権限のエラー
        // 5: ユーザーによるキャンセル        
        if (root.children.length == 0) {
            message += "更新が必要なシートはありません\nJSONファイルを出力しますか？";
            var btnr = showPopup(message, 0, "確認", ICON_QUESTN|BTN_OK_CANCL);
            if (btnr == BTNR_CANCL) {
                WScript.Quit(5);
            }
            return;
        }

        message += "以下のシートを作成・更新します\n\n";
        
        // 抽出した要素のtextプロパティに先頭に「*」をつけて改行で連結
        var formattedString = _.map(root.children, function(sheetNode) {
            return '* ' + sheetNode.text;
        }).join('\n');

        message += formattedString;

        var btnr = showPopup(message, 0, "シート作成・更新", BTN_OK_CANCL);
        if (btnr == BTNR_CANCL) {
            WScript.Quit(5);
        }
    })();
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

var PLACEHOLDER_WARN_ON_UNDEFINED = true;   // まずは警告だけ出す
var PLACEHOLDER_LEGACY_DROP       = true;   // 互換: 裸 {{…}} は falsy で行削除
var PLACEHOLDER_UNDEFINED_IS_ERROR = false; // 将来: 未定義でエラーに切り替える
var Q_BOOL_STRICT                 = false;  // {{?}} を bool 限定にするか
var PLACEHOLDER_WARN_DIALOG_LIMIT = 3;      // ダイアログに表示する警告の上限
var PLACEHOLDER_WARNINGS_FILENAME = "placeholder_warnings.txt";
var PLACEHOLDER_WARNINGS_CACHE_FILENAME = "placeholder_warnings.json";

var previousPlaceholderWarningsBySheet = loadPlaceholderWarningsCache();
var placeholderWarnings = [];
var cachedPlaceholderWarningsMerged = false;

function getPlaceholderWarningSheetName(node) {
    if (!node) {
        return "";
    }

    var current = node;
    while (current && current !== root) {
        if (current.parent === root) {
            var sheetText = current.text;
            if (sheetText == null) {
                return "";
            }
            return _.isString(sheetText) ? sheetText : String(sheetText);
        }
        current = current.parent;
    }

    return "";
}

function shouldReuseParsedSheetNode(parsedSheetNode, sheetName) {
    if (!parsedSheetNode) {
        return false;
    }

    var sheetKey = sheetName || "";
    if (previousPlaceholderWarningsBySheet &&
        previousPlaceholderWarningsBySheet[sheetKey] &&
        previousPlaceholderWarningsBySheet[sheetKey].length > 0) {
        return false;
    }

    return true;
}

function getPlaceholderWarningFilePath(lineObj) {
    if (!lineObj) {
        return void 0;
    }

    var projectDirectory = lineObj.projectDirectory;
    var filePathProjectLocal = lineObj.filePath;
    if (!projectDirectory || !filePathProjectLocal) {
        return filePathProjectLocal;
    }

    try {
        var filePathAbs = sourceLocalPathToAbsolutePath(filePathProjectLocal, projectDirectory);
        var sourceRelative = absolutePathToSourceLocalPath(filePathAbs, projectDirectory);
        if (sourceRelative && !/^\.\./.test(sourceRelative)) {
            return sourceRelative;
        }
        var rootRelative = CL.getRelativePath(conf.$rootDirectory, filePathAbs);
        return rootRelative || filePathProjectLocal;
    } catch (e) {
        return filePathProjectLocal;
    }
}

function isDuplicatePlaceholderWarning(existing, candidate) {
    var existingSheet = existing.sheetName || "";
    var candidateSheet = candidate.sheetName || "";
    if (existingSheet !== candidateSheet) {
        return false;
    }

    if (existing.kind !== candidate.kind) {
        return false;
    }

    var existingFile = existing.filePath || "";
    var candidateFile = candidate.filePath || "";
    var existingLine = existing.lineNum == null ? "" : String(existing.lineNum);
    var candidateLine = candidate.lineNum == null ? "" : String(candidate.lineNum);

    if (existing.kind === "undefinedPlaceholder") {
        return (existing.placeholder || "") === (candidate.placeholder || "") &&
            existingFile === candidateFile &&
            existingLine === candidateLine;
    }

    return (existing.message || "") === (candidate.message || "") &&
        existingFile === candidateFile &&
        existingLine === candidateLine;
}

function storePlaceholderWarning(entry) {
    var stored = _.assign({}, entry);
    var isDuplicate = _.some(placeholderWarnings, function(existing) {
        return isDuplicatePlaceholderWarning(existing, stored);
    });
    if (isDuplicate) {
        return false;
    }
    placeholderWarnings.push(stored);
    return true;
}

function pushPlaceholderWarning(entry, node) {
    var stored = _.assign({}, entry);
    stored.sheetName = getPlaceholderWarningSheetName(node);
    if (node && node.lineObj) {
        var lineObj = node.lineObj;
        stored.lineNum = lineObj.lineNum;
        var filePath = getPlaceholderWarningFilePath(lineObj);
        if (filePath) {
            stored.filePath = filePath;
        }
    }
    storePlaceholderWarning(stored);
}

function formatPlaceholderWarning(entry) {
    var message;
    if (entry.kind === "undefinedPlaceholder") {
        message = [
            "未定義プレースホルダー: " + entry.placeholder
        ];
        if (entry.filePath) {
            message.push("ファイル: " + entry.filePath);
        }
        if (entry.lineNum) {
            message.push("行: " + entry.lineNum);
        }
        message = message.join("\n");
    } else {
        message = entry.message || "";
        if (entry.filePath) {
            message += makeLineinfoString(entry.filePath, entry.lineNum);
        }
    }
    return message;
}

function getPlaceholderWarningsFilePath() {
    var outputDirectory = fso.GetParentFolderName(outfilePath);
    return fso.BuildPath(outputDirectory, PLACEHOLDER_WARNINGS_FILENAME);
}

function clearPlaceholderWarningsFile() {
    var warningsFilePath = getPlaceholderWarningsFilePath();
    if (fso.FileExists(warningsFilePath)) {
        try {
            fso.DeleteFile(warningsFilePath);
        } catch (e) {
            // ignore failures (e.g. read-only file)
        }
    }
}

function getPlaceholderWarningsCacheFilePath() {
    var outputDirectory = fso.GetParentFolderName(outfilePath);
    return fso.BuildPath(outputDirectory, PLACEHOLDER_WARNINGS_CACHE_FILENAME);
}

function clearPlaceholderWarningsCacheFile() {
    var cacheFilePath = getPlaceholderWarningsCacheFilePath();
    if (fso.FileExists(cacheFilePath)) {
        try {
            fso.DeleteFile(cacheFilePath);
        } catch (e) {
            // ignore failures (e.g. read-only file)
        }
    }
}

function loadPlaceholderWarningsCache() {
    var cacheFilePath = getPlaceholderWarningsCacheFilePath();
    if (!fso.FileExists(cacheFilePath)) {
        return {};
    }

    try {
        var content = CL.readTextFileUTF8(cacheFilePath);
        if (!content) {
            return {};
        }
        var parsed = JSON.parse(content);
        var result = {};
        for (var sheetName in parsed) {
            if (!parsed.hasOwnProperty(sheetName)) {
                continue;
            }
            var entries = parsed[sheetName];
            if (!_.isArray(entries)) {
                continue;
            }
            result[sheetName] = _.map(entries, function(entry) {
                var normalized = _.pick(entry, [
                    "kind",
                    "placeholder",
                    "message",
                    "filePath",
                    "lineNum",
                    "sheetName"
                ]);
                if (typeof normalized.sheetName !== "string") {
                    normalized.sheetName = sheetName;
                }
                return normalized;
            });
        }
        return result;
    } catch (e) {
        return {};
    }
}

function mergeCachedPlaceholderWarnings() {
    if (cachedPlaceholderWarningsMerged) {
        return;
    }
    cachedPlaceholderWarningsMerged = true;

    if (!previousPlaceholderWarningsBySheet) {
        previousPlaceholderWarningsBySheet = {};
    }

    var existingSheets = {};
    for (var i = 0; i < placeholderWarnings.length; i++) {
        var sheetKey = placeholderWarnings[i].sheetName || "";
        existingSheets[sheetKey] = true;
    }

    for (var sheetName in reusedSheetNames) {
        if (!reusedSheetNames.hasOwnProperty(sheetName)) {
            continue;
        }
        var sheetKey = sheetName || "";
        if (existingSheets[sheetKey]) {
            continue;
        }
        var cachedEntries = previousPlaceholderWarningsBySheet[sheetKey];
        if (!cachedEntries) {
            continue;
        }
        for (var j = 0; j < cachedEntries.length; j++) {
            storePlaceholderWarning(cachedEntries[j]);
        }
        existingSheets[sheetKey] = true;
    }
}

function groupPlaceholderWarningsBySheet(entries) {
    var order = [];
    var map = {};
    for (var i = 0; i < entries.length; i++) {
        var entry = entries[i];
        var sheetKey = entry.sheetName || "";
        if (!map.hasOwnProperty(sheetKey)) {
            map[sheetKey] = [];
            order.push(sheetKey);
        }
        map[sheetKey].push(entry);
    }
    return { order: order, map: map };
}

function savePlaceholderWarningsCache(groupedData) {
    var cacheFilePath = getPlaceholderWarningsCacheFilePath();
    var hasEntries = false;
    var data = {};

    for (var i = 0; i < groupedData.order.length; i++) {
        var sheetKey = groupedData.order[i];
        var groupEntries = groupedData.map[sheetKey];
        if (!groupEntries || groupEntries.length === 0) {
            continue;
        }
        hasEntries = true;
        data[sheetKey] = _.map(groupEntries, function(entry) {
            var stored = _.pick(entry, [
                "kind",
                "placeholder",
                "message",
                "filePath",
                "lineNum",
                "sheetName"
            ]);
            stored.sheetName = sheetKey;
            return stored;
        });
    }

    if (!hasEntries) {
        clearPlaceholderWarningsCacheFile();
        return;
    }

    try {
        CL.writeTextFileUTF8(JSON.stringify(data, null, 2), cacheFilePath);
    } catch (e) {
        // ignore failures (e.g. read-only file)
    }
}

function finalizePlaceholderWarnings() {
    mergeCachedPlaceholderWarnings();

    if (placeholderWarnings.length === 0) {
        clearPlaceholderWarningsFile();
        clearPlaceholderWarningsCacheFile();
        return null;
    }

    var grouped = groupPlaceholderWarningsBySheet(placeholderWarnings);

    var limit = PLACEHOLDER_WARN_DIALOG_LIMIT;
    if (!_.isNumber(limit) || limit <= 0) {
        limit = placeholderWarnings.length;
    }

    var remaining = Math.min(limit, placeholderWarnings.length);
    var shownCount = 0;
    var shownGroups = [];

    for (var gi = 0; gi < grouped.order.length && remaining > 0; gi++) {
        var sheetKey = grouped.order[gi];
        var groupEntries = grouped.map[sheetKey] || [];
        var groupMessages = [];
        for (var ei = 0; ei < groupEntries.length && remaining > 0; ei++) {
            groupMessages.push(formatPlaceholderWarning(groupEntries[ei]));
            remaining--;
            shownCount++;
        }
        if (groupMessages.length > 0) {
            shownGroups.push({
                sheetName: sheetKey,
                messages: groupMessages
            });
        }
    }

    var dialogParts = [];
    dialogParts.push("警告（全" + placeholderWarnings.length + "件中" + shownCount + "件を表示）");

    for (var sg = 0; sg < shownGroups.length; sg++) {
        var group = shownGroups[sg];
        var sheetDisplayName = group.sheetName ? group.sheetName : "シート不明";
        dialogParts.push("【" + sheetDisplayName + "】");
        dialogParts.push(group.messages.join("\n\n"));
    }

    var moreCount = placeholderWarnings.length - shownCount;

    var undefinedSections = [];
    for (var ui = 0; ui < grouped.order.length; ui++) {
        var sheetName = grouped.order[ui];
        var entries = grouped.map[sheetName] || [];
        var undefinedMessages = [];
        for (var uj = 0; uj < entries.length; uj++) {
            if (entries[uj].kind === "undefinedPlaceholder") {
                undefinedMessages.push(formatPlaceholderWarning(entries[uj]));
            }
        }
        if (undefinedMessages.length === 0) {
            continue;
        }
        var headerName = sheetName ? sheetName : "シート不明";
        undefinedSections.push("[" + headerName + "]");
        undefinedSections.push(undefinedMessages.join("\n\n"));
    }

    if (undefinedSections.length > 0) {
        var warningsFilePath = getPlaceholderWarningsFilePath();
        try {
            CL.writeTextFileUTF8(undefinedSections.join("\n\n") + "\n", warningsFilePath);
        } catch (e) {
            // ignore failures (e.g. read-only file)
        }
        dialogParts.push("未定義プレースホルダーの一覧を " + PLACEHOLDER_WARNINGS_FILENAME + " に出力しました。");
    } else {
        clearPlaceholderWarningsFile();
    }

    if (moreCount > 0) {
        dialogParts.push("... ほか " + moreCount + " 件の警告があります。");
    }

    savePlaceholderWarningsCache(grouped);

    return dialogParts.join("\n\n");
}

// 親スコープをプロトタイプ継承し、現在ノードのレイヤーを上書きで乗せる
function extendScope(parentScope, layer) {
    var base = parentScope || (typeof globalScope !== "undefined" ? globalScope : {});
    var child = Object.create(base);
    if (layer) {
        _.assign(child, layer); // 近いもの勝ち（シャドーイング）
    }
    return child;
}

// キャッシュ付き evaluator（プロトタイプ継承を素直に使える）
var __EVAL_EXPR_CACHE = {};
function evalExprWithScope(expr, scope) {
    var fn = __EVAL_EXPR_CACHE[expr];
    if (!fn) {
        fn = new Function("scope", "with(scope){ return (" + expr + "); }");
        __EVAL_EXPR_CACHE[expr] = fn;
    }
    return fn(scope);
}

// [] の中は数値だけ許可（動的は式評価へフォールバックさせる）
var SIMPLE_PATH_RE = /^[\w$]+(?:\.[\w$]+|\[\d+\])*$/;
var SIMPLE_PATH_LITERAL_MAP = {
    "true": true,
    "false": false,
    "null": null,
    "undefined": void 0
};
function evaluateExprOrPath(expr, scope) {
    // 単純参照（変数/ドット/ブラケット）は _.get でOK（プロトタイプを辿れる）
    if (SIMPLE_PATH_LITERAL_MAP.hasOwnProperty(expr)) {
        return SIMPLE_PATH_LITERAL_MAP[expr];
    }
    if (SIMPLE_PATH_RE.test(expr)) {
        return _.get(scope, expr, void 0);
    }
    // 任意式は with(scope) 評価（継承が効く）
    try {
        return evalExprWithScope(expr, scope);
    } catch (e) {
        return void 0;
    }
}

function isDropByQ(val) {
    if (Q_BOOL_STRICT) {
        if (typeof val !== "boolean") throw new Error("{{?}} は boolean のみ許可: " + val);
        return (val === false);
    }
    // 緩和モード: false/null/undefined で削除
    return (val === false || val === null || val === void 0);
}

function handleUndefinedPlaceholder(placeholderName, node) {
    if (PLACEHOLDER_UNDEFINED_IS_ERROR) {
        throw new ParseError("未定義プレースホルダー: " + placeholderName, node && node.lineObj);
    }
    if (PLACEHOLDER_WARN_ON_UNDEFINED) {
        pushPlaceholderWarning({
            kind: "undefinedPlaceholder",
            placeholder: placeholderName
        }, node);
    }
}

function evalPlaceholderToken(raw, scope, node) {
    var trimmedRaw = (raw || "").trim();
    var s = trimmedRaw;
    var mode = "legacy";
    if (s.charAt(0) === "?") {
        mode = "dropOnFalsy";
        s = s.slice(1).trim();
    }
    if (s.charAt(s.length - 1) === "!") {
        throw new ParseError("プレースホルダー '{{" + trimmedRaw + "}}' の '!' 指定は廃止されました。", node && node.lineObj);
    }

    var val = evaluateExprOrPath(s, scope);

    // 明示の ? : 条件ガード + 簡易プレースホルダ
    // - falsy(false/null/undefined / ※Q_BOOL_STRICTなら厳格) → ノードごと削除
    // - truthy のとき true は空文字、それ以外は通常の文字列化で出力
    if (mode === "dropOnFalsy") {
        var dropQ = isDropByQ(val);
        return {
            drop: dropQ,
            text: dropQ ? "" : (val === true ? "" : String(val))
        };
    }

    if (mode === "legacy" && PLACEHOLDER_LEGACY_DROP) {
        var legacyUndefined = (val === void 0 || val === null);
        var falsyLegacy = (val === false || legacyUndefined);
        if (legacyUndefined) {
            handleUndefinedPlaceholder(s, node);
        }
        return {
            drop: falsyLegacy,
            text: falsyLegacy ? "" : (val === true ? "" : String(val))
        };
    }

    // 警告だけ
    if (val === void 0 || val === null) {
        handleUndefinedPlaceholder(s, node);
    }

    // true は空文字、null/undefined は空
    return {
        drop: false,
        text: (val === true ? "" : (val == null ? "" : String(val)))
    };
}

// defaultParamKey はテンプレ内の "{{}}" 省略キー用（通常ノードでは null）
function replacePlaceholdersInNode(node, scope, defaultParamKey) {
    var defaultToken = defaultParamKey ? "{{" + defaultParamKey + "}}" : null;
    var RE_EMPTY = /\{\{\s*\}\}/g;
    var RE_EXPR  = /\{\{\s*([^\}]+)\s*\}\}/g;

    function applyOnce(s) {
        if (s === void 0 || s === null) return void 0;

        if (defaultToken) s = s.replace(RE_EMPTY, defaultToken);

        var toDelete = false;
        var out = s.replace(RE_EXPR, function(__, expr){
            if (toDelete) return "";

            var hit = evalPlaceholderToken(expr, scope, node);
            if (hit.drop) {
                toDelete = true;
                return "";
            }
            return hit.text;
        });

        return toDelete ? void 0 : out;
    }

    // text/comment/image に適用
    node.text = applyOnce(node.text);
    if (node.text === void 0) return false; // ノード削除
    node.comment = applyOnce(node.comment);
    node.imageFilePath = applyOnce(node.imageFilePath);
    return true;
}

// 必要に応じてシート（root直下）へ擬似引数を渡したい場合はこのフックを実装
function getSheetCallArgs(node) {
    // 例: conf.SHEETS?.byId[node.id] / byName[node.text] を返す… 等
    return null; // まずは無効
}

// ルートから降りながら、親→子にスコープを継承して {{}} を適用
function applyPlaceholdersEverywhere() {
    var scopeStack = [ (typeof globalScope !== "undefined" ? globalScope : {}) ];

    forAllNodes_Recurse(
        root, null, -1,
        // pre
        function(node, parent, index) {
            if (!node) return true;

            var parentScope = scopeStack[scopeStack.length - 1];
            var localScope  = parentScope;

            // 現在ノードの params を反映
            if (node.params) localScope = extendScope(localScope, node.params);

            // root直下（=シート）だけ疑似引数を上乗せしたい場合
            if (parent === root) {
                var callArgs = getSheetCallArgs(node);
                if (callArgs) localScope = extendScope(localScope, callArgs);
            }

            scopeStack.push(localScope);

            // *template(...) 行は展開側で処理するためスキップ
            var m = node.text && node.text.trim().match(/^\*([A-Za-z_]\w*)\((.*)\)$/);
            var skipChildren = false;
            if (!m) {
                var ok = replacePlaceholdersInNode(node, localScope, null);
                if (!ok) {
                    if (parent) {
                        parent.children[index] = null;
                    }
                    skipChildren = true;
                }
            }

            if (skipChildren) {
                scopeStack.pop();
                return true;
            }
        },
        // post
        function() { scopeStack.pop(); }
    );
}

//// 使用例
//var globalScope = { foo: 1 };
//var localScope = Object.create(globalScope);
//localScope.bar = 2;
//
//// 平坦化（継承元も含めて）
//var flatScope = _.assign({}, localScope);
//
//var result = evaluateInScope("foo + bar", flatScope); // → 3
function evaluateInScope(expr, scope) {
    var keys = _.keys(scope);
    var values = _.map(keys, function(k) { return scope[k]; });
    var func = new Function(keys.join(","), "return " + expr + ";");
    return func.apply(null, values);
}

// 一番上の階層の upper snake case なプロパティをシートから閲覧できるようにする
var globalScope = (function(original) {
    if (typeof original === "undefined") return {};

    var keys = _.keys(original);
    var filteredKeys = _.filter(keys, function(key) {
        return /^[A-Z0-9_]+$/.test(key);
    });
    var filtered = _.pick(original, filteredKeys);

    return filtered;
})(conf);
//printJSON(globalScope);

// テンプレート埋め込み
// まずはすべてのノードについて調べ、親に登録
(function() {
    var startTime = performance.now();

    // ===== Errors =====
    var TemplateError = function(errorMessage, node) {
        this.errorMessage = errorMessage;
        this.node = node;
    };

    function templateError(errorMessage, node) {
        var lineObj = node.lineObj;
        if (_.isUndefined(lineObj)) {
            MyError(errorMessage);
        } else {
            MyError(errorMessage, lineObj.filePath, lineObj.lineNum);
        }
    }

    // ===== Parameter Evaluation =====
    var templateParamFnCache = Object.create(null);

    function evalTemplateParameters(paramsStr, node, currentParameters) {
        paramsStr = paramsStr.trim();
        if (paramsStr == "") {
            return {};
        }

        var referableParams = {};
        // まずグローバル（あれば）
        if (typeof globalScope !== "undefined") {
          _.defaults(referableParams, globalScope);
        }
        if (!_.isUndefined(currentParameters)) {
            _.defaults(referableParams, currentParameters);
        }
        for (var parent = node.parent; !_.isUndefined(parent); parent = parent.parent) {
            if (!_.isUndefined(parent.params)) {
                _.defaults(referableParams, parent.params);
            }
        }

        // ★ 呼び出し引数の“丸ごと”参照を提供（互換用）
        attachArgAliases(referableParams, currentParameters);

        // XXX: 処理が重すぎる
        function parseParams(referableParams, paramsStr) {
            // _.keys(), _.values() は列挙順は保証されてないので一応自前で詰めておく
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
            var cacheKey = keys.join("\u0001") + "\u0002" + paramsStr;

            var fn = templateParamFnCache[cacheKey];
            if (!fn) {
                var fnArgs = keys.slice();
                fnArgs.push("__paramsStr");
                fn = Function(fnArgs.join(","), 'return eval("([" + __paramsStr + "])");');
                templateParamFnCache[cacheKey] = fn;
            }

            var fnValues = values.slice();
            fnValues.push(paramsStr);

            return fn.apply(null, fnValues);
        }

        var paramsArray = parseParams(referableParams, paramsStr);

        // object を返すには丸括弧が必要らしい
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
                        } else {
                            _.defaults(o, {$value: elem[i]});
                        }
                    }
                }
                // 関数が渡された場合、引数に渡された順に実行
                else if (_.isFunction(elem)) {
                    // o はこの関数で作った object なので clone 不要
                    // 関数の中で直接書き換えてもOK
                    o = elem(o);
                } else {
                    _.defaults(o, elem);
                }
            });
            mergedArray.push(o);
        });
        // ほぼ意味ないけど、要素数1の場合はobjectを返す
        return (mergedArray.length == 1) ? mergedArray[0] : mergedArray;
    }

    // ===== Tree Utilities =====
    // templateTree に対してそのまま cloneDeep を呼ぶと、 parent をさかのぼって tree 全体が clone されるので対処
    function cloneTemplateTree(srcTemplateTree) {
        // 自前で tree をたどって全 node を shallow copy
        var dstTemplateTree = _.assign({}, srcTemplateTree);

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
        _recurse(dstTemplateTree, srcTemplateTree);
        return dstTemplateTree;
    }

    function nodeToString(root) {
        var depth = 0;
        var s = "";
        forAllNodes_Recurse(root, null, -1,
            function(node, parent, index) {
                if (node === null) {
                    return true;
                }
                var indent = _.repeat("    ", depth);
                s += index + " : ";
                s += indent + "(" + node.group + " ," + node.depthInGroup + ")  " + node.text + "\n";
                depth++;
            },
            function(node, parent, index) {
                depth--;
            }
        );
        return s;
    }

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
                var validChildren = node.children.filter(function(element) {
                    return (element !== null);
                });
                if (validChildren.length === 0) {
                    if (node.kind === kindH) {
                        var errorMessage = "シート「"+ node.text +"」に有効な項目が存在しません\n※子階層がテンプレートのみとなっている可能性があります";
                        var lineObj = node.lineObj;
                        MyError(errorMessage, lineObj.filePath, lineObj.lineNum);
                    }
                    if (parent !== null) {
                        parent.children[index] = null;
                    }
                    return;
                }
                node.children = validChildren;
            }
        );
    }
    function shrinkChildrenArrayForAllNodes() {
        shrinkChildrenArray(root, null, -1);
    }

    // ===== Template Collection =====
    // &NAME: データ定義（params）を親に集約
    function collectTemplateParamsFromULNodes() {
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
                    templateError(errorMessage, node);
                }
            } else {
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

            parent.params[paramName] = param;

            // 親の children の自分自身を null に
            parent.children[index] = null;
        });
    }

    // すべてのテンプレート宣言（&NAME()）を tree から取り外し、所属 node にリストアップ
    // 命名: templates
    function collectTemplateDeclarations() {
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

            var templateName = match[1];

            if ("templates" in parent) {
                // 重複エラー
                if (templateName in parent.templates) {
                    var errorMessage = "テンプレート名'"+ templateName +"'が重複しています。";
                    templateError(errorMessage, node);
                }
            } else {
                parent.templates = {};
            }
            parent.templates[templateName] = node;

            // node の group 関係を template root からの offset 値に
            // 木の中で宣言した場合でも大丈夫なように対応しておく
            var templateGroup = node.group;
            var templateDepthInGroup = node.depthInGroup;
            forAllNodes_Recurse(node, null, -1, function(n, p, i) {
                if (n.group === templateGroup) {
                    // templateRoot と同じ group の node の depthInGroup は必ず 1 多いので引いておく
                    n.depthInGroup -= templateDepthInGroup + 1;
                }
                n.group -= templateGroup;
            });

            // 親の children の自分自身を null に
            parent.children[index] = null;

            // TODO: 宣言時の defaultParameter 仕様は廃止方向
        });
    }

    // 名前から tree をさかのぼって見つける
    // なければ null を返す
    function findTemplate_Recurse(templateName, node) {
        if (_.isUndefined(node) || node === null) {
            return null;
        }
        if ("templates" in node) {
            if (templateName in node.templates) {
                return node.templates[templateName];
            }
        }
        return findTemplate_Recurse(templateName, node.parent);
    }

    // ===== Template Reference Verification =====
    // 問題がないか調べる
    // 一度確認した template は isValid フラグ立て（json 出力前に delete）
    function verifyTemplateReference(templateRoot) {

        function _recurse(templateNode, callStack) {
            if (templateNode.isValidSubTree) {
                // 旧フラグ名を流用（互換）
                return;
            }

            if (templateNode.children.length === 0) {
                var errorMessage = "テンプレートには1個以上の子ノードが必要です。";
                throw new TemplateError(errorMessage, templateNode);
            }

            for (var i = 0; i < templateNode.children.length; i++) {
                if (templateNode.children[i].group !== templateNode.group) {
                    var errorMessage = "テンプレートの第2階層はグループ切り替えはできません。\nルート（テンプレート名の行）と同じマークにしてください";
                    throw new TemplateError(errorMessage, templateNode);
                }
            }

            var templateName = templateNode.text.slice(2, -1); // "&NAME()"
            var lineObj = templateNode.lineObj;
            var callName = templateName + ":" + lineObj.filePath + ":" + lineObj.lineNum;
            if (_.indexOf(callStack, callName) >= 0) {
                var errorMessage = "テンプレート'"+ templateName +"'に循環参照が存在します。";
                throw new TemplateError(errorMessage, templateNode);
            }
            callStack.push(callName);

            forAllNodes_Recurse(templateNode, null, -1, function(n) {
                var match = n.text.trim().match(/^\*([A-Za-z_]\w*)\(.*\)$/);
                if (match === null) {
                    return;
                }
                var refTemplateName = match[1];

                var refTemplate = findTemplate_Recurse(refTemplateName, n.parent);

                // みつからなかった
                if (refTemplate === null) {
                    var errorMessage = "テンプレート'" + refTemplateName + "'は存在しません。";
                    throw new TemplateError(errorMessage, n);
                }

                _recurse(refTemplate, callStack);
            });

            callStack.pop();
            templateNode.isValidSubTree = true; // 互換目的の既存フラグを継続使用
        }

        try {
            _recurse(templateRoot, []);
        } catch (e) {
            if (_.isUndefined(e.node) || _.isUndefined(e.errorMessage)){
                throw e;
            }
            templateError(e.errorMessage, e.node);
        }
    }

    function attachArgAliases(scope, parameters) {
        // 引数がオブジェクト/配列で、かつ空でない時だけ付与
        if (!parameters || typeof parameters !== "object") return;
        var isEmpty = _.isArray(parameters) ? parameters.length === 0 : _.isEmpty(parameters);
        if (isEmpty) return; // 引数が空なら付与しない

        var copy = _.isArray(parameters) ? parameters.slice() : _.assign({}, parameters);
        scope.$args   = copy;
        scope.$params = copy; // 互換のため当面残す
    }

    function execInScope(code, scope) {
        var fn = new Function("scope", "with(scope){\n" + code + "\n}");
        return fn(scope);
    }

    function parseConstLiteral(src) {
        var t = String(src).trim();

        // boolean / null
        if (t === "true") return true;
        if (t === "false") return false;
        if (t === "null") return null;

        // number (int/float/scientific)
        if (/^-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(t)) return Number(t);

        // quoted string
        if (/^"(?:\\.|[^"\\])*"$/.test(t)) return JSON.parse(t);
        if (/^'(?:\\.|[^'\\])*'$/.test(t)) {
            var q = "\"" + t.slice(1, -1)
                .replace(/\\'/g, "'")
                .replace(/\\/g, "\\\\")
                .replace(/"/g, "\\\"") + "\"";
            return JSON.parse(q);
        }

        // fallback: unquoted string (YAML っぽく無引用を許可)
        // ※ 先頭末尾の空白は trim 済み。必要なら引用してください。
        return t;
    }

    function runAnchorDeclarationsGlobally(root) {
        forAllNodes_Recurse(root, null, -1, function(node, parent, index) {
            if (!node || node.kind !== kindUL) return;

            var t = node.text.trim();
            var m = t.match(/^&([A-Za-z_]\w*):\s*([\s\S]*)$/);
            if (!m) return;

            // 子を持つ (&foo_array:) は既存の配列処理に委譲
            if (node.children && node.children.length > 0) return;

            var name = m[1];
            var raw  = m[2];

            // ★ もっとも近い祖先に params を作ってそこへ保存
            var owner = node.parent;
            while (owner && !owner.params) owner = owner.parent;
            if (!owner) owner = parent || root;       // どこにも無ければ親→root に作る
            owner.params = owner.params || {};
            var target = owner.params;

            try {
                target[name] = parseConstLiteral(raw);
            } catch (e) {
                // 無言方針なら何もしない／必要なら最小限のログ
            }

            if (parent) parent.children[index] = null;  // 宣言ノードは削除
        });

        shrinkChildrenArray(root, null, -1);
    }

    function runAnchorDeclarations(tplRoot, scope) {
        function visit(node) {
            if (!node || !node.children) {
                return;
            }

            for (var i = 0; i < node.children.length; i++) {
                var child = node.children[i];
                if (!child) {
                    continue;
                }

                if (child.kind === kindUL && (!child.children || child.children.length === 0)) {
                    var match = child.text && child.text.match(/^&([A-Za-z_]\w*):\s*([\s\S]*)$/);
                    if (match) {
                        try {
                            scope[match[1]] = parseConstLiteral(match[2]);
                        } catch (e) {
                            // 既存方針: サイレントに無視
                        }

                        node.children.splice(i, 1);
                        i--;
                        continue;
                    }
                }

                visit(child);
            }

            if (node.templates) {
                for (var key in node.templates) {
                    if (!Object.prototype.hasOwnProperty.call(node.templates, key)) {
                        continue;
                    }
                    var templateNode = node.templates[key];
                    if (!templateNode) {
                        continue;
                    }
                    var cloned = cloneTemplateTree(templateNode);
                    visit(cloned);
                }
            }
        }

        visit(tplRoot);
    }

    // ----- @init をテンプレート外でも実行する -----
    function runInitDirectivesGlobally(root) {
        if (typeof globalScope === "undefined") { globalScope = {}; }
        var scopeStack = [ globalScope ];
        var snapshotStack = [];

        forAllNodes_Recurse(
            root, null, -1,
            function (node, parent, index) {
                if (!node) return true;

                var parentScope = scopeStack[scopeStack.length - 1];
                var localScope = extendScope(parentScope, node.params || {});
                var before = {};
                for (var key in localScope) {
                    if (!Object.prototype.hasOwnProperty.call(localScope, key)) continue;
                    before[key] = localScope[key];
                }
                scopeStack.push(localScope);
                snapshotStack.push({ scope: localScope, parent: parentScope, before: before });

                // UL の 「@init: <code>」のみ対応（コロン必須 / 子ULは読まない）
                if (node.kind === kindUL) {
                    var t = (node.text || "").trim();
                    var m = t.match(/^@init\s*:\s*([\s\S]*)$/);
                    if (m) {
                        var code = m[1] || "";

                        // --- 永続化先を決定：祖先 params → H1 params → globalScope
                        function findParamOwner(n) {
                            for (var p = n; p; p = p.parent) { if (p && p.params) return p; }
                            return null;
                        }
                        var owner = findParamOwner(node);
                        if (!owner) {
                            var sheet = FindParentNode(node, function (n) { return n && n.kind === "H" && n.level === 1; });
                            if (sheet) { sheet.params = sheet.params || {}; owner = sheet; }
                        }
                        var persistTarget = owner ? owner.params : globalScope;

                        // 実行前スナップショット（own props のみ）
                        var before = {};
                        for (var k in localScope) {
                            if (Object.prototype.hasOwnProperty.call(localScope, k)) {
                                before[k] = localScope[k];
                            }
                        }

                        try {
                            installInitHelpers(localScope);   // $get/$set/$defaults
                            execInScope(code, localScope);    // with(scope){ ... }

                            // 新規/変更されたキーだけ永続層へコピー（$* と関数は除外）
                            for (var key in localScope) {
                                if (!Object.prototype.hasOwnProperty.call(localScope, key)) continue;
                                if (key.charAt(0) === "$") continue;
                                if (typeof localScope[key] === "function") continue;

                                var isNew = !(key in before);
                                var changed = isNew ? true : (before[key] !== localScope[key]);
                                if (isNew || changed) {
                                    persistTarget[key] = localScope[key];
                                }
                            }
                        } catch (e) {
                            // 細かいエラーは出しすぎない方針なので簡素化
                            var lo = node.lineObj;
                            var msg = "init 実行エラー: " + e.message;
                            if (lo) MyError(msg, lo.filePath, lo.lineNum); else MyError(msg);
                        } finally {
                            delete localScope.$get; delete localScope.$set; delete localScope.$defaults;
                        }

                        // 実行済みの @init 行は除去
                        if (parent) parent.children[index] = null;
                    }
                }
            },
            function () {
                var state = snapshotStack.pop();
                var localScope = state.scope;
                var parentScope = state.parent;
                var before = state.before;

                for (var diffKey in localScope) {
                    if (!Object.prototype.hasOwnProperty.call(localScope, diffKey)) continue;
                    if (diffKey.charAt(0) === "$") continue;
                    if (typeof localScope[diffKey] === "function") continue;

                    var isNew = !Object.prototype.hasOwnProperty.call(before, diffKey);
                    var changed = isNew ? true : (before[diffKey] !== localScope[diffKey]);
                    if (isNew || changed) {
                        parentScope[diffKey] = localScope[diffKey];
                    }
                }

                scopeStack.pop();
            }
        );

        // children の null を掃除
        shrinkChildrenArray(root, null, -1);
    }

    function installInitHelpers(scope) {
        scope.$get = function(k){ return (k in scope) ? scope[k] : void 0; };
        scope.$set = function(k,v){ scope[k] = v; };
        scope.$defaults = function(obj){
            if (!obj || typeof obj !== "object") return;
            for (var k in obj) if (obj.hasOwnProperty(k) && !(k in scope)) scope[k] = obj[k];
        };
    }

    function runInitDirectives(tplRoot, scope) {
        function visit(node, currentScope) {
            if (!node) {
                return;
            }

            var localScope = extendScope(currentScope || {}, node.params || {});
            var scopeSnapshot = {};
            for (var key in localScope) {
                if (!Object.prototype.hasOwnProperty.call(localScope, key)) continue;
                scopeSnapshot[key] = localScope[key];
            }

            if (node.children) {
                for (var i = 0; i < node.children.length; i++) {
                    var child = node.children[i];
                    if (!child) {
                        continue;
                    }

                    if (child.kind === kindUL) {
                        var match = (child.text || "").trim().match(/^@init\s*:\s*([\s\S]*)$/);
                        if (match) {
                            var code = match[1] || "";
                            var directiveScope = localScope;
                            var propagateToLocal = false;
                            var before = null;
                            if (child.params) {
                                var extended = extendScope(directiveScope, child.params);
                                propagateToLocal = (extended !== directiveScope);
                                directiveScope = extended;
                                if (propagateToLocal) {
                                    before = {};
                                    for (var snapshotKey in directiveScope) {
                                        if (!Object.prototype.hasOwnProperty.call(directiveScope, snapshotKey)) continue;
                                        before[snapshotKey] = directiveScope[snapshotKey];
                                    }
                                }
                            }
                            try {
                                installInitHelpers(directiveScope);
                                execInScope(code, directiveScope);
                            } catch (e) {
                                templateError("init 実行エラー:\n" + e.message, child);
                            } finally {
                                delete directiveScope.$get; delete directiveScope.$set; delete directiveScope.$defaults;
                                if (propagateToLocal) {
                                    for (var diffKey in directiveScope) {
                                        if (!Object.prototype.hasOwnProperty.call(directiveScope, diffKey)) continue;
                                        if (diffKey.charAt(0) === "$") continue;
                                        if (typeof directiveScope[diffKey] === "function") continue;

                                        var isNew = !Object.prototype.hasOwnProperty.call(before, diffKey);
                                        var changed = isNew ? true : (before[diffKey] !== directiveScope[diffKey]);
                                        if (isNew || changed) {
                                            localScope[diffKey] = directiveScope[diffKey];
                                        }
                                    }
                                }
                            }

                            node.children.splice(i, 1);
                            i--;
                            continue;
                        }
                    }

                    visit(child, localScope);
                }
            }

            if (node.templates) {
                for (var key in node.templates) {
                    if (!Object.prototype.hasOwnProperty.call(node.templates, key)) {
                        continue;
                    }
                    var templateNode = node.templates[key];
                    if (!templateNode) {
                        continue;
                    }
                    var cloned = cloneTemplateTree(templateNode);
                    visit(cloned, localScope);
                }
            }

            if (currentScope) {
                for (var diffKey in localScope) {
                    if (!Object.prototype.hasOwnProperty.call(localScope, diffKey)) continue;
                    if (diffKey.charAt(0) === "$") continue;
                    if (typeof localScope[diffKey] === "function") continue;

                    var isNew = !Object.prototype.hasOwnProperty.call(scopeSnapshot, diffKey);
                    var changed = isNew ? true : (scopeSnapshot[diffKey] !== localScope[diffKey]);
                    if (isNew || changed) {
                        currentScope[diffKey] = localScope[diffKey];
                    }
                }
            }
        }

        visit(tplRoot, scope || {});
    }

    function expandInlineParamArray(targetNode, targetIndex, paramName, callSiteScope) {
        var parent = targetNode.parent;
        if (!parent) {
            return;
        }

        if (targetNode.children && targetNode.children.length > 0) {
            throw new TemplateError("'*" + paramName + "' と子要素は同時に使用できません。", targetNode);
        }

        var scopedValue = callSiteScope && callSiteScope[paramName];
        if (_.isUndefined(scopedValue)) {
            throw new TemplateError("データ'" + paramName + "'は存在しません。", targetNode);
        }

        var list = toRepeatList(scopedValue) || scopedValue;
        if (!_.isArray(list)) {
            throw new TemplateError("データ'" + paramName + "'は配列ではありません。", targetNode);
        }

        var baseId = targetNode.id;
        var replacements = [];

        _.forEach(list, function(element, index) {
            var entry = element;
            if (!_.isObject(entry)) {
                entry = { $value: entry };
            }

            var value = entry.$value;
            if (value === void 0 || value === null) {
                value = "";
            }

            var suffix = entry.$id;
            if (!suffix) {
                suffix = "i" + index;
            }

            var newNode = _.assign({}, targetNode);
            newNode.text = String(value);
            delete newNode.comment;
            delete newNode.imageFilePath;
            delete newNode.tableData;
            delete newNode.initialValues;
            delete newNode.url;
            newNode.variables = {};
            delete newNode.params;
            delete newNode.templates;
            newNode.children = [];
            newNode.id = baseId + "_" + suffix;
            newNode.parent = parent;

            replacements.push(newNode);
        });

        var siblings = parent.children;
        siblings[targetIndex] = null;
        var insertedChildren = siblings.slice(0, targetIndex + 1).concat(replacements).concat(siblings.slice(targetIndex + 1));
        insertedChildren[targetIndex] = null;
        parent.children = insertedChildren;
    }

    // ===== Template Expansion =====
    // node に template の clone を追加する（展開前の状態で追加）
    function addTemplate(targetNode, targetIndex, templateName, parameters, callSiteScope) {

        // パラメータが配列ならローリング展開
        function rollArray(targetNode, targetIndex, templateName, list) {
            var clonedTargetNodes = [];
            var total = list.length;
            _.forEach(list, function(element, index) {
                var node = cloneTemplateTree(targetNode);
                if (!_.isObject(element)) {
                    element = {
                        $value: element
                    };
                }
                var elementId = ("$id" in element) ? element.$id : "i" + index;
                node.id = targetNode.id + "_" + elementId;

                // ループメタを付与
                element = _.assign({}, element, buildLoopMeta(index, total));

                var paramJSON = JSON.stringify(element);
                node.text = "*" + templateName + "(" + paramJSON + ")";
                clonedTargetNodes.push(node);
            });

            targetNode.parent.children[targetIndex] = null;
            var a = targetNode.parent.children;
            var insertedChildren = a.slice(0, targetIndex+1).concat(clonedTargetNodes).concat(a.slice(targetIndex+1));
            insertedChildren[targetIndex] = null;
            targetNode.parent.children = insertedChildren;

            // ここではノードの追加のみ（処理は後段）
        }

        // 数値/{$times:...}/配列 を一律リスト化
        var __list = toRepeatList(parameters) || (_.isArray(parameters) ? parameters : null);
        if (__list) {
            rollArray(targetNode, targetIndex, templateName, __list);
            return;
        }

        var templateRoot = findTemplate_Recurse(templateName, targetNode.parent);

        // みつからなかった
        if (templateRoot === null) {
            var errorMessage = "テンプレート'" + templateName + "'は存在しません。";
            throw new TemplateError(errorMessage, targetNode);
        }

        // まず clone
        templateRoot = cloneTemplateTree(templateRoot);

        // 変数展開（共通 evaluator）
        {
            // 呼び出し地点のスコープに引数を最上段で重ねる
            if (!parameters || typeof parameters !== "object") parameters = {};
            var parametersScopeTop = extendScope(callSiteScope, parameters);

            attachArgAliases(parametersScopeTop, parameters);

            // 省略時はこれを使う（引数1個を想定）
            var defaultParam = "$value";
            var firstParam = _.find(_.keys(parameters), function(s) { return s.substr(0,1) != "$"; });
            if (!_.isUndefined(firstParam)) defaultParam = firstParam;

            runAnchorDeclarations(templateRoot, parametersScopeTop);
            runInitDirectives(templateRoot, parametersScopeTop);

            // ★ テンプレツリー内でも params を積みながら置換
            var tplStack = [ parametersScopeTop ];
            forAllNodes_Recurse(
                templateRoot, null, -1,
                function(n, p, i) {
                    if (!n) return true;
                    var parentScope = tplStack[tplStack.length - 1];
                    var localScope  = n.params ? extendScope(parentScope, n.params) : parentScope;
                    tplStack.push(localScope);
                    var ok = replacePlaceholdersInNode(n, localScope, defaultParam);
                    if (!ok) { n.parent.children[i] = null; return; }
                },
                function(){ tplStack.pop(); }
            );
            shrinkChildrenArray(templateRoot, null, -1);
        }

        // template 内の template 呼び出し（ネスト展開）
        var tplScopeStack = [ parametersScopeTop ];
        forAllNodes_Recurse(templateRoot, null, -1, function(n, p, i) {
            var parentScope = tplScopeStack[tplScopeStack.length - 1];
            var localScope  = n && n.params ? extendScope(parentScope, n.params) : parentScope;
            tplScopeStack.push(localScope);

            if (p === null) {
                return;
            }
            var match = n.text.trim().match(/^\*([A-Za-z_]\w*)\((.*)\)$/);
            if (match === null) {
                return;
            }
            var innerTemplateName = match[1];
            var parsedParameters;
            try {
                parsedParameters = evalTemplateParameters(match[2], n, localScope);
            } catch(e) {
                var errorMessage = "パラメータが不正です。\n\n" + e.message;
                templateError(errorMessage, n);
            }

            addTemplate(n, i, innerTemplateName, parsedParameters, localScope);
        }, function(){ tplScopeStack.pop(); });

        // template の leaf に target の子ノードを追加する
        if (targetNode.children.length > 0) {
            var targetClone = cloneTemplateTree(targetNode);

            // offset にしておく
            forAllNodes_Recurse(targetClone, null, -1, function(n) {
                if (n.group === targetNode.group) {
                    n.depthInGroup -= targetNode.depthInGroup;
                }
                n.group -= targetNode.group;
            });

            forAllNodes_Recurse(templateRoot, null, -1, function(n, p, i) {
                if (n.children.length > 0) {
                    return;
                }
                // 内容は不問
                if (_.has(n, 'attributes.sealed')) {
                    return;
                }
                var templateLeaf = n;
                var target = cloneTemplateTree(targetClone);
                forAllNodes_Recurse(target, null, -1, function(nn) {
                    if (nn === null) {
                        return true;
                    }
                    if (nn.group === 0) {
                        nn.depthInGroup += templateLeaf.depthInGroup;
                    }
                    nn.group += templateLeaf.group;
                    if (nn.children.length === 0) {
                        // id を _ で連結
                        nn.id = templateLeaf.id + "_" + nn.id;
                        return true;
                    }
                });
                templateLeaf.children = target.children;
                return true;
            });
        }

        // template の 全 node の group と leaf の id を書き換える
        forAllNodes_Recurse(templateRoot, null, -1, function(n) {
            if (n === null) {
                return true;
            }
            // group 関係は template root からのオフセットとして扱う
            if (n.group === 0) {
                n.depthInGroup += targetNode.depthInGroup;
            }
            n.group += targetNode.group;
            if (n.children.length === 0) {
                // id を _ で連結
                n.id = targetNode.id + "_" + n.id;
                return true;
            }
        });

        // splice で自分を template の children で置き換える（直後に挿入 + 自分は null 予約）
        var a = targetNode.parent.children;
        // template の parent 書き換え
        for (var j = 0; j < templateRoot.children.length; j++) {
            if (templateRoot.children[j] === null) {
                continue;
            }
            templateRoot.children[j].parent = targetNode.parent;
        }
        var insertedChildren = a.slice(0, targetIndex+1).concat(templateRoot.children).concat(a.slice(targetIndex+1));
        insertedChildren[targetIndex] = null;
        targetNode.parent.children = insertedChildren;
    }

    // テンプレートをインライン展開していく
    function expandAllTemplateCalls() {
        var scopeStack = [ (typeof globalScope !== "undefined" ? globalScope : {}) ];

        forAllNodes_Recurse(
            root, null, -1,
            function(node, parent, index) {
                if (!node) return true;

                var parentScope = scopeStack[scopeStack.length - 1];
                var localScope  = parentScope;
                if (node.params) localScope = extendScope(localScope, node.params);
                scopeStack.push(localScope);

                var trimmedText = node.text && node.text.trim();

                var inlineArrayMatch = trimmedText && trimmedText.match(/^\*([A-Za-z_]\w*)$/);
                if (inlineArrayMatch) {
                    var paramName = inlineArrayMatch[1];
                    try {
                        expandInlineParamArray(node, index, paramName, localScope);
                    } catch (e) {
                        if (_.isUndefined(e.node) || _.isUndefined(e.errorMessage)) throw e;
                        templateError(e.errorMessage, e.node);
                    }
                    return;
                }

                var match = trimmedText && trimmedText.match(/^\*([A-Za-z_]\w*)\((.*)\)$/);
                if (match) {
                    var templateName = match[1];
                    var parameters;
                    try {
                        parameters = evalTemplateParameters(match[2], node, {});
                    } catch(e) {
                        templateError("パラメータが不正です。\n\n" + e.message, node);
                    }

                    if (parameters === null) {
                        if (parent && parent.children) {
                            parent.children[index] = null;
                        }
                        return;
                    }

                    try {
                        // ★ 呼び出し地点のスコープを addTemplate に渡す
                        addTemplate(node, index, templateName, parameters, localScope);
                    } catch (e) {
                        if (_.isUndefined(e.node) || _.isUndefined(e.errorMessage)) throw e;
                        templateError(e.errorMessage, e.node);
                    }
                }
            },
            function() { scopeStack.pop(); }
        );
    }

    // すべての展開・置換が終わった最後に呼ぶ
    function dropInitNodesEverywhere(root) {
        forAllNodes_Recurse(root, null, -1, function(n, p, i) {
            if (!n || n.kind !== kindUL) return;
            var t = (n.text || "").trim();
            if (/^@init\s*:/.test(t)) {
                p.children[i] = null;
            }
        });
        shrinkChildrenArray(root, null, -1);
    }

    // ===== Pipeline =====
    // 1) データ定義（&name:）収集
    collectTemplateParamsFromULNodes();

    // 2) テンプレート宣言（&name()）収集
    collectTemplateDeclarations();

    // 3) null を除去して一次整形
    shrinkChildrenArrayForAllNodes();

    // 3.5) テンプレート外 @init を実行（placeholders 置換の前に）
    runAnchorDeclarationsGlobally(root);
    runInitDirectivesGlobally(root);
    // 4) テンプレ外の {{...}} を適用するプリパス
    applyPlaceholdersEverywhere();

    // （必要に応じてテンプレ宣言内部の参照検証）
    // ※元コードでは全テンプレ事前展開/検証はコメントアウトされていたため、呼び出しは保持しません。
    //    verifyTemplateReference(...) を使いたい場合は、parent.templates を走査して適宜呼び出してください。

    // 5) ルートから *template(...) 呼び出しを順次インライン展開
    expandAllTemplateCalls();

    dropInitNodesEverywhere(root);

    // 6) 再度 null 洗浄
    shrinkChildrenArrayForAllNodes();

    // 7) leaf じゃなくなった node の id を削除
    forAllNodes_Recurse(root, null, -1, function(node, parent, index) {
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

forAllNodes_Recurse(root, null, -1, function(node, parent, index) {
    var headers = node.tableHeadersNonInputArea;
    if (!headers) {
        return;
    }
    for (var i = 0; i < headers.length; i++) {
        delete headers[i].marker;
    }
});

// srcHash が同じだった sheetNode を元の位置に挿入
_.forEach(parsedSheetNodeInfos, function(info) {
    root.children.splice(info.index, 0, info.node);
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
// TODO: commit, update とかは一旦すべて使わないことになったので、いろいろ気にせずやめても良い。無駄に時間食いすぎる
// XXX: とりあえず SHA256 はとんでもなく時間かかるので MD5 に。JSON.stringfy がかかるのでソーステキストに
computeRootId();


var sJson = stringifyPretty(root);

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

CL.writeTextFileUTF8(sJson, outfilePath);

var strUpdatedSrcFiles = (function () {
    if (_.isEmpty(srcTextsToRewrite)) {
        return null;
    }

    var message = "以下のソースファイルに変更を加えました\n\n";
    message += _.map(srcTextsToRewrite, function(value, key) {
        return '* ' + key;
    }).join('\n');

    return message;
})();

var placeholderWarningsMessage = finalizePlaceholderWarnings();

if (!runInCScript) {
    var messageParts = ["JSONファイル(" + outFilename + ")を出力しました"];

    if (strUpdatedSrcFiles) {
        messageParts.push(strUpdatedSrcFiles);
    }
    if (placeholderWarningsMessage) {
        messageParts.push(placeholderWarningsMessage);
    }

    alert(messageParts.join("\n\n---\n\n"));
}
else {
    var cliMessageParts = [];

    // CScript 実行時は更新情報と警告のみ通知する
    if (strUpdatedSrcFiles) {
        cliMessageParts.push(strUpdatedSrcFiles);
    }
    if (placeholderWarningsMessage) {
        cliMessageParts.push(placeholderWarningsMessage);
    }

    if (cliMessageParts.length > 0) {
        alert(cliMessageParts.join("\n\n---\n\n"));
    }
}

WScript.Quit(0);
