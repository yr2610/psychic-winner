// XXX: 高速化のため、現状、 formula とか Date には対応してない

var verify = true;    // 十分使って問題が一度も起きないようであればverifyしないようにする。こういうのは wsf に書くべきか
var compress = true;    // XXX: 一旦ここで。最終的には設定ファイルとかから読むように

var repositoryFileOnly = true;

function Error(message)
{
    shell.Popup(message, 0, "エラー", ICON_EXCLA);
    WScript.Quit();
}

var shell = new ActiveXObject("WScript.Shell");
var shellApplication = new ActiveXObject("Shell.Application");


if (( WScript.Arguments.length != 1 ) ||
    ( WScript.Arguments.Unnamed(0) == ""))
{
    Error("変更個所を commit したいチェックリスト（Excelファイル）をドロップしてください。");
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
// 更新不要の場合にexcel閉じなくて良いように
var isExcelFileOpened = CL.isFileOpened(filePath);

// TODO: Excelファイルの確認

initializeExcel();
//excel.Visible = true;
//excel.ScreenUpdating = true;

var book = openBookReadOnly(filePath);

// コンフリクトしてたら何もしない
if (findSheetByName(book, "conflicts"))
{
    finalizeExcel();
    Error("conflicts を解決してから再度実行してください");
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


function GetUserName()
{
    var network = new ActiveXObject("WScript.Network");

    // TODO: ユーザー名をダイアログから入力させる。デフォは WScript.Network の UserName で取得

    return network.UserName;
}

var userName = GetUserName();

var history;
var historySheet = findSheetByName(book, "history");
if (!historySheet)
{
    history = {
        head: 0,
        // head の状態
        data: {
            checkSheet: {
                sheets: {}
            }
        },
        // 変更履歴
        changeSets: []
    };
}
else
{
    // history の JSON を読み込み
    history = CL.readJSONFromSheet(historySheet);
}

var newRevision = (!historySheet) ? 0 : history.head + 1;

// Excelの今の状態とheadとの差分を求める
// 差分を求める過程で history.data 用の形式のも作られるので、利用
var sheetChanges = getSheetChanges(root, book, history.data.checkSheet.sheets);

// TODO: index の変更も取り込む
// TODO: getIndexSheetVariables(root, book);
// TODO: getIndexSheetValues(root, book);

// TODO: index にも変更がないという条件追加
if (newRevision > 0 && Object.keys(sheetChanges.changes).length === 0) {
    shell.Popup("変更個所はありません", 0);
    finalizeExcel();
    WScript.Quit();
}

if (isExcelFileOpened)
{
    Error("Excelファイルが開いています。\nファイルを閉じてから再度実行してください。");
    finalizeExcel();
    WScript.Quit();
}

history.data.checkSheet.sheets = sheetChanges.valueses;

var changeSet = {
    revision: newRevision,
    id: CL.createRandomId(16),
    author: userName,
    date: (new Date()).toString(),
    changes: {
        checkSheet: {
            sheets: sheetChanges.changes
        },
        indexSheet: {
            // TODO:
        }
    }
};

// revision 0 からさらに遡る必要はないので持つ必要はない。冗長なだけ
if (changeSet.revision === 0)
{
    delete changeSet.changes;
}

history.changeSets.push(changeSet);
history.head = newRevision;

// TODO: 行数の上限指定できるように
// シートごとに一旦独立した文字列で求めて、行数が上限超えるか都度確認して追加
function changeSetToReadableString(changeSet, data)
{
    var s = "";
    var sheets = changeSet.changes.checkSheet.sheets;
    var dataSheets = data.checkSheet.sheets;
    for (var sheetId in sheets)
    {
        var sheet = sheets[sheetId];
        var dataSheet = dataSheets[sheetId];
        s += "# " + sheet.text + "\n";
        for (var id in sheet.items)
        {
            var item = sheet.items[id];
            var dataItem = dataSheet ? dataSheet.items[id] : undefined;
            s += "- " + item.text + "\n";
            for (var header in item.values)
            {
                var value0 = item.values[header];
                value0 = (value0 === null) ? "" : value0;
                value0 = '"' + value0 + '"';
                var value = dataItem ? dataItem.values[header] : "";
                if (typeof value === "undefined")
                {
                    value = "";
                }
                value = '"' + value + '"';
                s += "  " + header + ": " + value0 + " → " + value + "\n";
            }
        }
        s += "\n";
    }

    return s;
}

// 確認
// XXX: 内容をダイアログに表示すると言う仕様は仮。量が増えたらボタンが表示されない
(function () {
    var message;
    if (history.head === 0) {
        message = "現在の状態を Revision 0 としてバージョン管理を開始します。\nよろしければOKボタンを押してください。\n";
    }
    else {
        if (repositoryFileOnly) {
            message = "以下の変更を Revision " + history.head + " として repository ファイルを出力します。\nよろしければOKボタンを押してください。\n";
        }
        else {
            message = "以下の変更を Revision " + history.head + " としてコミットします。\nよろしければOKボタンを押してください。\n";
        }
        message += "\n";
        message += changeSetToReadableString(changeSet, history.data);
    }

    if (shell.Popup(message, 0, "コミット", BTN_OK_CANCL) !== BTNR_OK)
    {
        finalizeExcel();
        WScript.Quit();
    }
})();

if (!repositoryFileOnly || history.head === 0) {

if (!historySheet)
{
    // history という名前のシートを作成
    historySheet = book.Worksheets.Add();
    historySheet.Name = "history";
    historySheet.Move(null, book.Worksheets(book.Worksheets.Count));
    historySheet.Visible = false;

    // 初回 commit 時は index シートが選択された状態にする
    var indexSheet = CL.getIndexSheet(book, root);
    indexSheet.Select();
}

// history シートを新しいデータで更新
CL.writeJSONToSheet(history, historySheet);

CL.createChangelogSheet(book, history, root, excel, templateData);



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
excel.Visible = true;
excel.ScreenUpdating = true;
/*/
// いろいろ怖いんで毎回バックアップはとっておく
// ファイル名に rev をつけて別名保存なので、コピーではなく元ファイルを移動で
CL.moveFile(filePath, buildBackupFolderName("commit"));
// repo ファイルがあれば一緒に move しておく
(function () {
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );
    var baseName = fso.GetBaseName(filePath);
    var parentFolderName = fso.GetParentFolderName(filePath);
    var repoFilePath = fso.BuildPath(parentFolderName, baseName + ".repo");
    if (fso.FileExists(repoFilePath)) {
        CL.moveFile(repoFilePath, buildBackupFolderName("commit"));
    }
})();

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

book.SaveAs(getRevisionedExcelBaseFileName(history.head, filePath));

}

/**/
finalizeExcel();
/*/
// Excelは閉じない
excel.Visible = true;
excel.ScreenUpdating = true;
/**/

function getHistoryJSONBaseFileName(history)
{
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );
    var baseName = fso.GetBaseName(filePath);
    var match = baseName.match(/^(.+)\-r\d+$/);
    if (match)
    {
        baseName = match[1];
    }

    var revision = history.head;

    return baseName + "-r" + revision + ".repo";
}

var endMessage;

// revision 0 の場合、 repo ファイルは不要
if (history.head > 0) {
    if (repositoryFileOnly) {
        endMessage = "変更を Revision " + history.head + " として ";
    }
    else {
        endMessage = "変更を Revision " + history.head + " としてコミットしました\n";
    }
    (function() {
        var outFilename = getHistoryJSONBaseFileName(history);
        var fso = new ActiveXObject( "Scripting.FileSystemObject" );
        var outfilePath = fso.BuildPath(fso.GetParentFolderName(filePath), outFilename);
        var outString = JSON.stringify(history, undefined, 4);

        if (compress) {
            // TODO: 一番圧縮率がいいのはUTF16圧縮してファイルもUTF16で保存。UTF8で保存だとEncodedURIComponentの方がファイルサイズが小さい
            //var compressor = LZString.compressToUTF16;
            //var decompressor = LZString.decompressFromUTF16;
            //var compressOption = "UTF16";
            var compressor = LZString.compressToEncodedURIComponent;
            var decompressor = LZString.decompressFromEncodedURIComponent;
            var compressOption = "EncodedURIComponent";
            var compressed = compressor(outString);
            if (verify) {
                if (outString !== decompressor(compressed)) {
                    Error("compress failed.");
                }
            }
            var compressData = {
                compress: "LZString",
                option: compressOption,
                data: compressed
            };
            outString = JSON.stringify(compressData, undefined, 4);
        }

        CL.writeTextFileUTF8(outString, outfilePath);

        endMessage += "repository ファイル(" + outFilename + ")を出力しました";
    })();
} else {
    endMessage = "現在の状態を Revision " + history.head + " としてバージョン管理を開始しました";
}

WScript.Echo(endMessage);

WScript.Quit();


// ==============================================

function getIndexSheetVariables(root, book)
{
    var indexSheet = CL.getIndexSheet(book, root);

    // 変数名が _ で始まり、その次が大文字の変数は値を取り込む
    for (var key in root.variables)
    {
        if (!/^_[A-Z].*/.test(key))
        {
            continue;
        }

        var cell = indexSheet.Range(root.variables[key]);

        // XXX: 日付を取得するために Text にしてお茶を濁す。本当はDateをDateとして扱うべき
        //root.variables[key] = cell.Value;
        root.variables[key] = cell.Text;
    }
    
}

function getIndexSheetValues(root, book)
{
    // まだ headerAddress 関係の情報がない頃のフォーマット
    if (!root.headerAddress)
    {
        return;
    }

    var indexSheet = CL.getIndexSheet(book, root);
    var headerRow = indexSheet.Range(root.headerAddress).Row;
    var headerCellColumn = indexSheet.Range(root.headerAddress).Column;
    var shouldSave = {};  // 保存対象外か。column が キー

    var leftHeaderCell = getFirstCellInRow(indexSheet, headerRow);
    var rightHeaderCell = getLastCellInRow(indexSheet, headerRow);
    var headerCells = indexSheet.Range(leftHeaderCell, rightHeaderCell);
    xEach(headerCells, function(c)
    {
        // シート名のセルは当然保存対象外
        if (c.Column === headerCellColumn)
        {
            return;
        }

        // 数式は出力しない
        if (c.Offset(1, 0).HasFormula)
        {
            return;
        }

        var headerName = c.Text;
        // 見出しが空欄の列は保存対象外
        if (!headerName)
        {
            return;
        }

        // header の text が!で挟まれてる列は保存対象外
        if (/^\!.*\!$/.test(headerName))
        {
            return;
        }

        shouldSave[c.Column] = headerName;
    });

    // 最左列は番号、最下行は集計行という前提で
    var leftBottomCell = getLastCellInColumn(indexSheet, leftHeaderCell.Column);
    var idColumnRange = indexSheet.Range(leftHeaderCell.Offset(1, 0), leftBottomCell.Offset(-1, 0));
    // 元のアドレスをキーとして元の位置からの現在の位置までのrowのずれ
    var primalToCurrentH1Addresses = {};
    xEach(idColumnRange, function(c)
    {
        var address = headerCells.Offset(c.Row - headerRow, 0).Address(false, false);
        var primalAddress = headerCells.Offset(c.Value, 0).Address(false, false);

        primalToCurrentH1Addresses[primalAddress] = address;
    });

    for (var i = 0; i < root.children.length; i++)
    {
        var nodeH1 = root.children[i];
        nodeH1.indexSheetValues = {};
        var addressH1 = primalToCurrentH1Addresses[nodeH1.indexSheetAddress];
        var rangeH1 = indexSheet.Range(addressH1);

        xEach(rangeH1, function(c)
        {
            if (!shouldSave[c.Column])
            {
                return;
            }

            var v = c.Value;

            if (!v)
            {
                return;
            }

            if (typeof v === 'date')
            {
                // TODO: JSON.parse() 後に new Date(Date.parse(v)) する必要がある？
                v = (new Date(v)).toDateString();
            }

            var headerName = shouldSave[c.Column];
            nodeH1.indexSheetValues[headerName] = v;
        });

        if (Object.keys(nodeH1.indexSheetValues).length === 0)
        {
            //delete nodeH1.indexSheetValues;
        }

    }
}


// ==================

function getSheetChanges(root, book, sheetsData)
{
    var sheetChanges = {};
    var sheetValueses = {};

    for (var i = 0; i < root.children.length; i++)
    {
        var nodeH1 = root.children[i];
        var name = nodeH1.text;
        var sheet = findSheetByName(book, name);
        var checkRange = new CheckRange(nodeH1, sheet, templateData);
        var itemsData = (nodeH1.id in sheetsData) ? sheetsData[nodeH1.id].items : {};
        var change = checkRange.getChangeFromSheetValues(itemsData);

        if (change)
        {
            sheetChanges[nodeH1.id] = change;
        }

        var sheetValues = checkRange.getSheetValues();
        if (Object.keys(sheetValues.items).length !== 0)
        {
            sheetValueses[nodeH1.id] = sheetValues;
        }
    }

    return {
        changes: sheetChanges,
        valueses: sheetValueses
    };
}

function getSheetsData(root, book)
{
    var sheetsData = {};

    for (var i = 0; i < root.children.length; i++)
    {
        var nodeH1 = root.children[i];
        var name = nodeH1.text;
        var sheet = findSheetByName(book, name);
        var checkRange = new CheckRange(nodeH1, sheet, templateData);
        var itemsData = (nodeH1.id in sheetsData) ? sheetsData[nodeH1.id].items : {};
        var change = checkRange.getChangeFromSheetValues(itemsData);

        if (change)
        {
            sheetsData[nodeH1.id] = change;
        }
    }

    return sheetsData;
}
