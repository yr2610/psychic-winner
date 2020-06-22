// XXX: 高速化のため、現状、 Date には対応してない

function Error(message)
{
    shell.Popup(message, 0, "エラー", ICON_EXCLA);
    WScript.Quit();
}

var shell = new ActiveXObject("WScript.Shell");
var shellApplication = new ActiveXObject("Shell.Application");

// フォルダー内の filepath を取得
function getFolderFiles(folderspec)
{
    var a = [];
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var f = fso.GetFolder(folderspec);
    var fc = new Enumerator(f.files);
    for (; !fc.atEnd(); fc.moveNext())
    {
        a.push(fc.item());
    }
    return a;
}

function getBaseNameFromRevisionedExcelFile(filePath)
{
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );
    var extensionName = fso.GetExtensionName(filePath);
    var baseName = fso.GetBaseName(filePath);

    var match = baseName.match(/^(.+)\-r\d+$/);
    if (match)
    {
        baseName = match[1];
    }

    return baseName;
}

// filepath 配列から repo ファイルを抜き出し
function getFolderRepoFiles(filePath)
{
    var excelBaseName = getBaseNameFromRevisionedExcelFile(filePath);
    var parentFolderName = fso.GetParentFolderName(filePath);
    var repoRe = /^(.+)\-r\d+\.repo$/;

    return getFolderFiles(parentFolderName).filter(function(element, index, array){
        var fileName = fso.GetFileName(element);
        var repoMatch = fileName.match(repoRe);
        return (repoMatch && repoMatch[1] === excelBaseName);
    });
}

// repo filepath 配列から番号が最大のものを返す
function getMaxRepoFile(files)
{
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var repoRe = /^.+\-r(\d+)$/;
    var o = {};
    for (var i = 0; i < files.length; i++)
    {
        var fileName = fso.GetBaseName(files[i]);
        var repoMatch = fileName.match(repoRe);
        o[repoMatch[1]] = files[i];
    }
    var maxRepo = Math.max.apply(null, Object.keys(o));

    return o[maxRepo];
}

function addSheetToEndOfBook(book, name, visible)
{
    var sheet = book.Worksheets.Add();
    sheet.Name = name;
    sheet.Move(null, book.Worksheets(book.Worksheets.Count));
    sheet.Visible = visible;

    return sheet;
}


if (( WScript.Arguments.length != 1 ) ||
    ( WScript.Arguments.Unnamed(0) == ""))
{
    Error("取り込み先のチェックリスト（Excelファイル）をドロップしてください。");
}

var filePath = WScript.Arguments.Unnamed(0);

// 一応 xls* で受け取るようにしておく
(function() {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    if (fso.GetExtensionName(filePath).substring(0, 3) != "xls") {
        Error("Excel ファイルをドラッグ＆ドロップしてください。");
    }
})();

// フラグだけ立てといて、更新不要かどうかの判定後にエラーを出す
// 更新不要とか、その他エラーがあった場合にexcel閉じなくて良いように
var isExcelFileOpened = CL.isFileOpened(filePath);

var fso = new ActiveXObject("Scripting.FileSystemObject");
var repoFiles = getFolderRepoFiles(filePath);

if (repoFiles.length === 0)
{
    var repoFileName = getBaseNameFromRevisionedExcelFile(filePath) + "-rX.repo";
    Error(".repo ファイルがありません。\n\n取り込み先のチェックリストと同じフォルダに\n" + repoFileName + "\nを置いてください。");
}

// 単純に sort じゃダメ
var repoFilePath = getMaxRepoFile(repoFiles);
var xlsFilePath = filePath;

var srcHistory = CL.readJSONFile(repoFilePath);


// TODO: Excelファイルの確認

initializeExcel();
//excel.Visible = true;
//excel.ScreenUpdating = true;

var book = openBookReadOnly(xlsFilePath);

// コンフリクトしてたら何もしない
if (findSheetByName(book, "conflicts"))
{
    finalizeExcel();
    Error("競合を解決してから再度実行してください");
}

var jsonSheet = findSheetByName(book, "JSON");
if (!jsonSheet)
{
    finalizeExcel();
    Error("JSONシートが存在しません");
}

var root = CL.readJSONFromSheet(jsonSheet);

var templateData;
var templateDataSheet = findSheetByName(book, "template.json");
if (templateDataSheet) {
    templateData = CL.ReadJSONFromSheet(templateDataSheet);
}
else {
    // TODO: 不要になったら削除
    CheckRange = CheckRange_old;
}

// 一番若い共通の親を返す
function findCommonParentChangeSetIndex(history0, history1)
{
    var changeSets0 = history0.changeSets;
    var changeSets1 = history1.changeSets;
    var i;

    for (i = 0; i < changeSets0.length && i < changeSets1.length; i++)
    {
        if (changeSets0[i].id !== changeSets1[i].id)
        {
            return i - 1;
        }
    }
    return i - 1;
}

function findNodeById(node, id)
{
    if (node.id === id)
    {
        return node;
    }

    for (var i = 0; i < node.children.length; i++)
    {
        var result = findNodeById(node.children[i], id);
        // id はユニークという前提なので、１つ見つかった時点で終了して良い
        if (result)
        {
            return result;
        }
    }

    return null;
}

// JSON を指定のリビジョンまで戻す
function revertTo(root, history, changeSetIndex)
{
    for (var i = history.changeSets.length - 1; i > changeSetIndex; --i)
    {
        var changeSet = history.changeSets[i];
        for (var sheetIndex = 0; sheetIndex < changeSet.sheetChanges.length; sheetIndex++)
        {
            var sheetChange = changeSet.sheetChanges[sheetIndex];
            var sheetNode = findNodeById(root, sheetChange.sheetId);
            for (var itemIndex = 0; itemIndex < sheetChange.itemChanges.length; itemIndex++)
            {
                var itemChange = sheetChange.itemChanges[itemIndex];
                var node = findNodeById(sheetNode, itemChange.id);
                for (var headerName in itemChange.change)
                {
                    var value = itemChange.change[headerName];

                    // undefined かどうかで処理を分ける必要はない雰囲気
                    node.values[headerName] = value.from;
                }
            }
        }
    }
}


var historySheet = findSheetByName(book, "history");

// XXX: historySheet なしの xlsx への update は一旦実装優先度低で
// TODO: そのまま srcHistory を取り込んで、 data との差分があれば conflict として出力
if (!historySheet)
{
    finalizeExcel();
    Error("historyシートが存在しません");
}


// history の JSON を読み込み
var dstHistory = CL.readJSONFromSheet(historySheet);

var commonParentChangeSetIndex = findCommonParentChangeSetIndex(srcHistory, dstHistory);

// すべて取り込み済み
if (commonParentChangeSetIndex === srcHistory.head)
{
    finalizeExcel();
    shell.Popup("最新の revision です\n更新の必要はありません", 0);

    discardedOldRepoFiles(filePath, dstHistory.head);
    
    WScript.Quit();
}

if (commonParentChangeSetIndex === -1)
{
    finalizeExcel();
    Error("チェックリストのバージョンが異なるため更新できません");
}

if (isExcelFileOpened)
{
    Error("Excelファイルが開いています。\nファイルを閉じてから再度実行してください。");
    finalizeExcel();
    WScript.Quit();
}

function revertCheckSheet(data, changeSet)
{
    var srcSheets = changeSet.changes.checkSheet.sheets;
    var dstSheets = data.checkSheet.sheets;
    for (var sheetId in srcSheets)
    {
        var srcSheet = srcSheets[sheetId];
        if (!(sheetId in dstSheets))
        {
            dstSheets[sheetId] = {
                text: srcSheet.text,
                items: {}
            };
        }
        var dstSheet = dstSheets[sheetId];

        for (var itemId in srcSheet.items)
        {
            var srcItem = srcSheet.items[itemId];

            if (!(itemId in dstSheet.items))
            {
                dstSheet.items[itemId] = {
                    text: srcItem.text,
                    values: {}
                };
            }
            var dstItem = dstSheet.items[itemId];

            for (var header in srcItem.values)
            {
                var value = srcItem.values[header];

                dstItem.values[header] = (value === null) ? undefined : value;
            }
        }
    }
}

// historData を changeSet １個分戻す
function revertHistoryData(data, changeSet)
{
    revertCheckSheet(data, changeSet);

    // TODO: indexSheet の values, variables も
}

// history を指定の revision まで戻す
// history を直接変更する
// とはいっても完全な deep copy ではないので、いろいろ注意
function revertHistoryTo(history, revision)
{
    for (var i = history.changeSets.length - 1; i >= revision + 1; --i)
    {
        var changeSet = history.changeSets.pop();

        revertHistoryData(history.data, changeSet);
    }

    history.head = revision;
}

// 共通の親の状態を取得
revertHistoryTo(dstHistory, commonParentChangeSetIndex);

var commonParentData = dstHistory.data;

// 3way merge を行う
// mine は今の book から取得
// theirs が優先されるように
// book にそのまま反映
// conflicts があれば conflicts シート（シート名に日時）作ってにとりあえず json を出力
// TODO: indexSheet の values, variables も
function mergeAndApplyToSheet(root, book, parentData, theirsData, conflicts)
{
    function addconflict(nodeH1, checkRange, x, y, parent, theirs, mine)
    {
        var sheets = conflicts.checkSheet.sheets;
        if (!(nodeH1.id in sheets))
        {
            sheets[nodeH1.id] = {
                text: nodeH1.text,
                items: {}
            };
        }
        var sheet = sheets[nodeH1.id];
        var leafNode = checkRange.leafNodes[y];
        if (!(leafNode.id in sheet.items))
        {
            sheet.items[leafNode.id] = {
                text: leafNode.text,
                values: {}
            };
        }
        var item = sheet.items[leafNode.id];
        var header = checkRange.headers[x];
        // 空欄を "" にしておく。 null だと分かりにくいと思うのでとりあえず
        item.values[header] = {
            base: (typeof parent === "undefined") ? "" : parent,
            theirs: (typeof theirs === "undefined") ? "" : theirs,
            mine: (typeof mine === "undefined") ? "" : mine
        };
    }

    var parentSheets = parentData.checkSheet.sheets;
    var theirsSheets = theirsData.checkSheet.sheets;

    for (var i = 0; i < root.children.length; i++)
    {
        var nodeH1 = root.children[i];
        var sheetId = nodeH1.id;
        var name = nodeH1.text;
        var sheet = findSheetByName(book, name);
        var checkRange = new CheckRange(nodeH1, sheet, templateData);

        var mine = checkRange.getCheckCellArray2d();
        var parent;
        if (sheetId in parentSheets)
        {
            parent = checkRange.getArray2dFromSheetValues(parentSheets[sheetId].items);
        }
        else
        {
            parent = this.createEmptyArray2d();
        }
        var theirs;
        if (sheetId in theirsSheets)
        {
            theirs = checkRange.getArray2dFromSheetValues(theirsSheets[sheetId].items);
        }
        else
        {
            theirs = this.createEmptyArray2d();
        }

        for (var x = 0; x < checkRange.width; x++)
        {
            if (!checkRange.isTarget[x])
            {
                continue;
            }
            for (var y = 0; y < checkRange.height; y++)
            {
                if (theirs[y][x] === mine[y][x] ||
                    parent[y][x] === mine[y][x])
                {
                    // theirs 採用
                    continue;
                }
                if (theirs[y][x] === parent[y][x])
                {
                    theirs[y][x] = mine[y][x];
                    continue;
                }
                // conflict に追加しつつ theirs 採用
                addconflict(nodeH1, checkRange, x, y, parent[y][x], theirs[y][x], mine[y][x]);
            }
        }

        // merge 結果をシートに反映
        checkRange.setArray2d(theirs);
    }
}

var conflicts = {
    date: (new Date()).toString(),
    checkSheet: {
        sheets: {}
    }
};

mergeAndApplyToSheet(root, book, commonParentData, srcHistory.data, conflicts);

if (!historySheet)
{
    // history という名前のシートを作成
    historySheet = addSheetToEndOfBook(book, "history", false);
}

// history は丸々置き換え
// それまでの履歴は破棄で
CL.writeJSONToSheet(srcHistory, historySheet);

CL.createChangelogSheet(book, srcHistory, root, excel, templateData);

// コンフリクトシート作って書き出し
// TODO: indexSheet 関係も
if (Object.keys(conflicts.checkSheet.sheets).length > 0)
{
    var conflictsSheet = addSheetToEndOfBook(book, "conflicts", true);
    conflictsSheet.Tab.ColorIndex = 3;

    // TODO: conflictsSheet に情報を書き出し

    // 先頭に移動して選択
    conflictsSheet.Move(book.Worksheets(1), null);
    conflictsSheet.Select();

    // XXX: とりあえず JSON そのまま
    CL.writeJSONToSheet(conflicts, conflictsSheet);

    shell.Popup("競合がありました\n競合を解決してください", 0, "競合", ICON_EXCLA);
}

function buildBackupFolderName(folderName) {
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );
    var baseName = fso.GetBaseName(filePath);
    var match = baseName.match(/^(.+)\-r\d+$/);
    if (match)
    {
        baseName = match[1];
    }
    var s = fso.BuildPath("bak", baseName);
    if (typeof folderName !== "undefined") {
        s = fso.BuildPath(s, folderName);
    }

    return s;
}

/**
jsonSheet.Visible = true;
historySheet.Visible = true;
/**/
// いろいろ怖いんで毎回バックアップはとっておく
// ファイル名に rev をつけて別名保存なので、コピーではなく元ファイルを移動で
CL.moveFile(xlsFilePath, buildBackupFolderName("update"));

// save as で revision(head) をファイル名に含める
// すでにファイル名に revision がついていれば置き換えられるように
function getRevisionedExcelBaseFileName(revision, filePath)
{
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );
    var extensionName = fso.GetExtensionName(filePath);
    var baseName = fso.GetBaseName(filePath);
    var parentFolderName = fso.GetParentFolderName(filePath);

    var match = baseName.match(/^(.+)\-r\d+$/);
    if (match)
    {
        baseName = match[1];
    }

    var fileName = baseName + "-r" + revision + "." + extensionName;

    return fso.BuildPath(parentFolderName, fileName);
}

book.SaveAs(getRevisionedExcelBaseFileName(srcHistory.head, filePath));

// Excelは閉じない
excel.Visible = true;
excel.ScreenUpdating = true;

function getFolderRepoFilesToDelete(filePath)
{
    var excelBaseName = getBaseNameFromRevisionedExcelFile(filePath);
    var parentFolderName = fso.GetParentFolderName(filePath);
    var repoRe = /^(.+)\-r\d+\.repo$/;

    return getFolderFiles(parentFolderName).filter(function(element, index, array){
        var fileName = fso.GetFileName(element);
        var repoMatch = fileName.match(repoRe);
        return (repoMatch && repoMatch[1] === excelBaseName);
    });
}

// repo ファイルは不要なので、削除
function discardedOldRepoFiles(filePath, headRevision) {
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );
    var repoFiles = getFolderRepoFiles(filePath);
    var repoRe = /^.+\-r(\d+)\.repo$/;

    // head revision 以下の revision の repository file を抽出
    var filesToDiscard = repoFiles.filter(function(element, index, array){
        var fileName = fso.GetFileName(element);
        var repoMatch = fileName.match(repoRe);
        var revision = parseInt(repoMatch[1]);
        return (revision <= headRevision);
    });

    if (filesToDiscard.length === 0)
    {
        return;
    }

    var s = filesToDiscard.reduce(function(previousValue, currentValue, index, array){
        return previousValue + fso.GetFileName(currentValue) + "\n";
    }, "");

    if (shell.Popup("以下のファイルはもう使いません。削除しますか？\n\n" + s, 0, "確認", ICON_QUESTN|BTN_YES_NO) !== BTNR_YES)
    {
        // XXX: 削除しない場合でも移動だけはしておく方が良い？
        return;
    }

    // fso.DeleteFile でワイルドカードでの削除はしないでおく
    filesToDiscard.forEach(function(element, index, array) {
        element.Delete(true);
    });

}
discardedOldRepoFiles(filePath, srcHistory.head);

WScript.Quit();
