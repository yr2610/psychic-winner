function Error(message)
{
    shell.Popup(message, 0, "エラー", ICON_EXCLA);
    WScript.Quit(1);
}

var shell = new ActiveXObject("WScript.Shell");
var shellApplication = new ActiveXObject("Shell.Application");

if (( WScript.Arguments.length != 1 ) ||
    ( WScript.Arguments.Unnamed(0) == ""))
{
    Error("クリーンコピーを作成したいチェックリストをドラッグ＆ドロップしてください。");
}

var filePath = WScript.Arguments.Unnamed(0);

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

if (isExcelFileOpened)
{
    finalizeExcel();
    Error("Excelファイルが開いています。\nファイルを閉じてから再度実行してください。");
}

var jsonSheet = findSheetByName(book, "JSON");
if (!jsonSheet)
{
    Error("JSONシートが存在しません");
}

var root = CL.ReadJSONFromSheet(jsonSheet);

var indexSheet = CL.getIndexSheet(book, root);

var templateData;
var templateDataSheet = findSheetByName(book, "template.json");
if (templateDataSheet) {
    templateData = CL.ReadJSONFromSheet(templateDataSheet);
}


var sheetsToDelete = [
    "JSON",
    "template.json",
    "history",
    "changelog",
    null
];

excel.DisplayAlerts = false;
sheetsToDelete.forEach(function (item, index, array) {
    if (!item) {
        return;
    }
    var sheet = findSheetByName(book, item);
    if (sheet) {
        sheet.Delete();
    }
});
excel.DisplayAlerts = true;

// 表紙を select
indexSheet.Select();

if (templateData) {
    indexSheet.Range(templateData.indexSheet.table.address).Resize(1, 1).Select();
    // TODO: 良い感じのスクロール位置になるように
}

function getCleanCopyFileName(filePath)
{
    var fso = new ActiveXObject( "Scripting.FileSystemObject" );
    var extensionName = fso.GetExtensionName(filePath);
    var baseName = fso.GetBaseName(filePath);
    var parentFolderName = fso.GetParentFolderName(filePath);

    baseName = baseName.replace(/\-r\d+$/, "");
    baseName = baseName.replace(/\-\d+$/, "");

    var yyyymmdd = CL.yyyymmddhhmmss(new Date()).slice(0, -6);
    baseName = baseName.replace(/yyyymmdd/gi, yyyymmdd);
    baseName = baseName.replace(/yymmdd/gi, yyyymmdd.slice(2));
    baseName = baseName.replace(/mmdd/gi, yyyymmdd.slice(4));

    var fileName = baseName + "." + extensionName;

    return fso.BuildPath(parentFolderName, fileName);
}

var outFilePath = getCleanCopyFileName(filePath);

var notSaved = false;
try {
    book.SaveAs(outFilePath);
} catch (e) {
    notSaved = true;
}

excel.DisplayAlerts = false;
book.Final = true;
excel.DisplayAlerts = true;

finalizeExcel();

if (notSaved) {
    WScript.Quit();
}

var fso = new ActiveXObject( "Scripting.FileSystemObject" );

WScript.Echo(fso.GetFileName(outFilePath) + "\nを出力しました");
