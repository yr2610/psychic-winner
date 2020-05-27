var Excel =
{
    xlDown: -4121,
    xlUp: -4162,
    xlToLeft: -4159,
    xlToRight: -4161,

    xlComments: -4144,
    xlFormulas: -4123,
    xlValues: -4163,

    xlA1: 1,
    xlR1C1: -4150,

    // XlLookAt 
    xlWhole: 1,
    xlPart: 2,

    // XlSearchOrder
    xlByColumns: 2,
    xlByRows: 1,

    // XlSearchDirection
    xlNext: 1,
    xlPrevious: 2,

    xlFillCopy: 1,

    xlNone: -4142,

    xlDiagonalDown: 5,
    xlDiagonalUp: 6,
    xlEdgeBottom: 9,
    xlEdgeLeft: 7,
    xlEdgeRight: 10,
    xlEdgeTop: 8,
    xlInsideHorizontal: 12,
    xlInsideVertical: 11,

    xlCenter: -4108,

    xlSrcRange: 1,

    xlYes: 1,
    xlNo: 2,

    xlCellTypeConstants: 2,

    xlFormatFromLeftOrAbove: 0,
    xlFormatFromRightOrBelow: 1,

    xlXXXXXXXX: 0   // dummy
};

var excel;

function initializeExcel()
{
    try
    {
        excel = WScript.CreateObject("ET.Application");
    }
    catch(e)
    {
        excel = WScript.CreateObject("Excel.Application");
    }
}

function finalizeExcel()
{
    // Excelを閉じる
    excel.DisplayAlerts = false;    // today() が含まれてると開いただけで更新されるので
    book.Close();
    excel.DisplayAlerts = true;
    excel.Quit();
}

function openBook(path, readOnly)
{
    var updateLinks = 0;

    return excel.Workbooks.Open(path, updateLinks, readOnly);
}

function openBookReadOnly(path)
{
    var readOnly = true;

    return openBook(path, readOnly);
}


// このあたりはExcelというわけではない気がするけど、一旦ここで
function xEach(objs, f)
{
    for (var obj = new Enumerator(objs); !obj.atEnd(); obj.moveNext())
    {
        f(obj.item());
    }

}

function xFind(objs, f)
{
    for (var obj = new Enumerator(objs); !obj.atEnd(); obj.moveNext())
    {
        if (f(obj.item()))
        {
            return obj.item();
        }
    }
    return null;
}

function findSheetByName(book, sheetName)
{
    return xFind(book.Worksheets, function(sheet){ return (sheet.Name === sheetName); });
}

function getFirstCellInRow(sheet, row)
{
    return sheet.Cells(row, 1).End(Excel.xlToRight);
}

function getLastCellInRow(sheet, row)
{
    return sheet.Cells(row, excel.Columns.Count).End(Excel.xlToLeft);
}

function getLastCellInColumn(sheet, column)
{
    return sheet.Cells(excel.Rows.Count, column).End(Excel.xlUp);
}
