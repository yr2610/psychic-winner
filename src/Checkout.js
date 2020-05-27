// XXX: 高速化のため、現状、 formula とか Date には対応してない

function Error(message)
{
    shell.Popup(message, 0, "エラー", ICON_EXCLA);
    WScript.Quit();
}

var shell = new ActiveXObject("WScript.Shell");
var shellApplication = new ActiveXObject("Shell.Application");

if (false) {
var BIF_RETURNONLYFSDIRS = 0x00000001;
var BIF_DONTGOBELOWDOMAIN = 0x00000002;
var BIF_STATUSTEXT = 0x00000004;
var BIF_RETURNFSANCESTORS = 0x00000008;
var BIF_EDITBOX = 0x00000010;
var BIF_VALIDATE = 0x00000020;
var BIF_NEWDIALOGSTYLE = 0x00000040;
var BIF_BROWSEINCLUDEURLS = 0x00000080;
var BIF_USENEWUI =  BIF_EDITBOX | BIF_NEWDIALOGSTYLE;
var BIF_UAHINT = 0x00000100;
var BIF_NONEWFOLDERBUTTON = 0x00000200;
var iOptions = BIF_STATUSTEXT | BIF_VALIDATE | BIF_USENEWUI | BIF_UAHINT;
var ssfWINDOWS = 0x12;
var objFolder = shellApplication.BrowseForFolder(0, "フォルダ選択", iOptions, "\\\\ika");
//var objFolder = shellApplication.BrowseForFolder(0, "フォルダ選択", iOptions, ssfWINDOWS);
if (objFolder != null)
{
    // Add code here.
    Error(objFolder.Items().Item().Path);
}
Error("no folder");
}

if (( WScript.Arguments.length != 1 ) ||
    ( WScript.Arguments.Unnamed(0) == ""))
{
    Error("変更個所を commit したいチェックリスト（Excelファイル）をドロップしてください。");
}

var filePath = WScript.Arguments.Unnamed(0);

var fso = new ActiveXObject( "Scripting.FileSystemObject" );
var isFolder = fso.FolderExists(filePath);
Error(isFolder + ":" + filePath);

// TODO: Excelファイルの確認

initializeExcel();
//excel.Visible = true;
//excel.ScreenUpdating = true;

function OpenBook(path, readOnly)
{
    var updateLinks = 0;

    return excel.Workbooks.Open(filePath, updateLinks, readOnly);
}

function OpenBookReadOnly(path)
{
    var readOnly = true;

    return OpenBook(path, readOnly);
}

function FinalizeExcel()
{
    // Excelを閉じる
    excel.DisplayAlerts = false;    // today() が含まれてると開いただけで更新されるので
    book.Close();
    excel.DisplayAlerts = true;
    excel.Quit();
}

var book = OpenBook(filePath, false);

// コンフリクトしてたら何もしない
if (findSheetByName(book, "conflicts"))
{
    FinalizeExcel();
    Error("conflicts を解決してから再度実行してください");
}

var jsonSheet = findSheetByName(book, "JSON");
if (!jsonSheet)
{
    FinalizeExcel();
    Error("JSONシートが存在しません");
}

var root = CL.readJSONFromSheet(jsonSheet);


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
if (Object.keys(sheetChanges.changes).length === 0)
{
    shell.Popup("変更個所はありません", 0);
    FinalizeExcel();
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
// XXX: 内容をダイアログに表示すると言う仕様は仮。量が増えたらボタンが表示されない可能性高い
(function () {
    var message;
    if (history.head === 0) {
        message = "現在の状態を Revision 0 としてバージョン管理を開始します。\nよろしければOKボタンを押してください。\n";
    }
    else {
        message = "以下の変更を Revision " + history.head + " としてコミットします。\nよろしければOKボタンを押してください。\n";
        message += "\n";
        message += changeSetToReadableString(changeSet, history.data);
    }

    if (shell.Popup(message, 0, "コミット", BTN_OK_CANCL) !== BTNR_OK)
    {
        FinalizeExcel();
        WScript.Quit();
    }
})();


if (!historySheet)
{
    // history という名前のシートを作成
    historySheet = book.Worksheets.Add();
    historySheet.Name = "history";
    historySheet.Move(null, book.Worksheets(book.Worksheets.Count));
    historySheet.Visible = false;
}

// history シートを新しいデータで更新
CL.writeJSONToSheet(history, historySheet);

/**
jsonSheet.Visible = true;
historySheet.Visible = true;
excel.Visible = true;
excel.ScreenUpdating = true;
/*/
// いろいろ怖いんで毎回バックアップはとっておく
// ファイル名に rev をつけて別名保存なのでバックアップ不要
//CL.makeBackupFile(filePath);

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

/**/
FinalizeExcel();
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

var endMessage = "変更を Revision " + history.head + " としてコミットしました\n";

// revision 0 の場合、 repo ファイルは不要
if (history.head > 0) {
    (function() {
        var outFilename = getHistoryJSONBaseFileName(history);
        var fso = new ActiveXObject( "Scripting.FileSystemObject" );
        var outfilePath = fso.BuildPath(fso.GetParentFolderName(filePath), outFilename);

        CL.writeTextFileUTF8(JSON.stringify(history, undefined, 4), outfilePath);

        endMessage += "repository ファイル(" + outFilename + ")を出力しました";
    })();
}

WScript.Echo(endMessage);

WScript.Quit();


// ==============================================

function getIndexSheetVariables(root, book)
{
    var indexSheet = findSheetByName(book, root.variables.sheetname ? root.variables.sheetname : "index");

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

    var indexSheet = findSheetByName(book, root.variables.sheetname ? root.variables.sheetname : "index");
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
        var checkRange = new CheckRange(nodeH1, sheet);
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
        var checkRange = new CheckRange(nodeH1, sheet);
        var itemsData = (nodeH1.id in sheetsData) ? sheetsData[nodeH1.id].items : {};
        var change = checkRange.getChangeFromSheetValues(itemsData);

        if (change)
        {
            sheetsData[nodeH1.id] = change;
        }
    }

    return sheetsData;
}
