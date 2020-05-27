function Error(message)
{
    shell.Popup(message, 0, "エラー", ICON_EXCLA);
    WScript.Quit();
}

var shell = new ActiveXObject("WScript.Shell");
var shellApplication = new ActiveXObject("Shell.Application");
var fso = new ActiveXObject( "Scripting.FileSystemObject" );

if (WScript.Arguments.length != 2 ||
    WScript.Arguments.Unnamed(0) == "" ||
    WScript.Arguments.Unnamed(1) == "")
{
    Error("以下のファイルを複数選択した状態でドラッグ＆ドロップしてください。\n\n* 取り込み先のチェックリスト（Excelファイル）\n* 取り込み元の入力データ（.savファイル）");
}

var dataFilePath = WScript.Arguments.Unnamed(0); // sav.json
var xlsFilePath = WScript.Arguments.Unnamed(1);

if (fso.GetExtensionName(dataFilePath) !== "sav")
{
    var t = xlsFilePath;
    xlsFilePath = dataFilePath;
    dataFilePath = t;
}

if (fso.GetExtensionName(dataFilePath) != "sav")
{
    Error("以下のファイルがドロップされていません。\n\n* 取り込み元の入力データ（savファイル）");
}

if (fso.GetExtensionName(xlsFilePath) != "xlsx")
{
    Error("以下のファイルがドロップされていません。\n\n* 取り込み先のチェックリスト（Excelファイル）");
}

if (CL.isFileOpened(xlsFilePath))
{
    Error("Excelファイルが開いています。\nファイルを閉じてから再度実行してください。");
    WScript.Quit();
}

//var data = CL.ReadJSONFile(dataFilePath);
var data = CL.readYAMLFile(dataFilePath);

initializeExcel();
//excel.Visible = true;
//excel.ScreenUpdating = true;

var book = openBook(xlsFilePath, false);

var jsonSheet = findSheetByName(book, "JSON");
if (!jsonSheet)
{
    Error("JSONシートが存在しません");
}

var root = CL.ReadJSONFromSheet(jsonSheet);

CL.AddParentPropertyForAllNodes(root);

var templateData;
var templateDataSheet = findSheetByName(book, "template.json");
if (templateDataSheet) {
    templateData = CL.ReadJSONFromSheet(templateDataSheet);
}
else {
    // TODO: 不要になったら削除
    sheetSetValues = sheetSetValues_old;
    indexSheetSetValues = indexSheetSetValues_old;
    indexSheetSetVariables = indexSheetSetVariables_old;
}

function getSrcHeaderToX(dstHeaders, nodeH1, sheetData, notFound)
{
    var dstHeaderToX = {};
    for (var i = 0; i < dstHeaders.length; i++)
    {
        dstHeaderToX[dstHeaders[i]] = i;
    }

    var idToX = {};
    for (var i = 0; i < nodeH1.tableHeaders.length; i++)
    {
        var header = nodeH1.tableHeaders[i];
        idToX[header.id] = dstHeaderToX[header.name];
    }

    // Object.keys(sheetData.items).length === 0 は想定しない
    // すべての item の values には同じ header が含まれている想定
    var key0 = Object.keys(sheetData.items)[0];
    var srcHeaders = Object.keys(sheetData.items[key0].values);
    var srcHeaderToX = {};
    for (var i = 0; i < srcHeaders.length; i++)
    {
        var header = srcHeaders[i];
        var idMatch = header.match(/^(\d+)\..*/);
        if (idMatch)
        {
            var id = parseInt(idMatch[1]);
            if (id in idToX)
            {
                srcHeaderToX[header] = idToX[id];
            }
            else
            {
                notFound.push(header);
            }
        }
        else
        {
            if (header in dstHeaderToX)
            {
                srcHeaderToX[header] = dstHeaderToX[header];
            }
            else
            {
                notFound.push(header);
            }
        }
    }

    return srcHeaderToX;
}

// 2d配列を部分的に取り出した2d配列
// 指定した矩形がsrcをはみ出した場合のチェックとかはしない
function getRectArray2d(src, offsetX, offsetY, width, height)
{
    var dst = new Array(height);

    for (var y = 0, srcY = offsetY; y < height; y++, srcY++)
    {
        dst[y] = src[srcY].slice(offsetX, offsetX + width);
    }

    return dst;
}

function sheetSetValues_old(sheet, nodeH1, sheetData)
{
    var dstLeaves = CL.GetLeafNodes(nodeH1);
    var idToY = {};
    for (var i = 0; i < dstLeaves.length; i++)
    {
        idToY[dstLeaves[i].id] = i;
    }

    var leftHeaderCell = sheet.Range(nodeH1.leftCheckHeaderCellAddress);
    var rightHeaderCell = getLastCellInRow(sheet, leftHeaderCell.Row);
    var headerRange = sheet.Range(leftHeaderCell, rightHeaderCell);
    var dstHeaders = headerRange.Value.toArray();

    var headerNotFound = [];
    var srcHeaderToX = getSrcHeaderToX(dstHeaders, nodeH1, sheetData, headerNotFound);

    var dstCheckRange = leftHeaderCell.Offset(1, 0).Resize(dstLeaves.length, dstHeaders.length);
    var dstArray2d = CL.RangeToValueArray2d(dstCheckRange);

    for (var id in sheetData.items)
    {
        var item = sheetData.items[id];
        var y = idToY[id];

        if (typeof y === "undefined")
        {
            // TODO: 一つでも入力があれば notFound に。 null は delete
            continue;
        }

        for (var header in item.values)
        {
            var x = srcHeaderToX[header];
            if (typeof x === "undefined")
            {
                // TODO: notFound に
                continue;
            }

            var value = item.values[header];
            dstArray2d[y][x] = (value === null) ? undefined : value;
        }
    }

    var hasFormulaColumns = [];
    xEach(headerRange.Offset(1, 0), function(cell)
    {
        // 列が数式かどうか
        hasFormulaColumns.push(cell.HasFormula);
    });

    // bool 反転
    var shouldApply = hasFormulaColumns.map(function(f) {
        return !f;
    });

    applyArrayToRange(dstCheckRange, dstArray2d, shouldApply);

    // 確認欄だけ header 込みで autofit
    // 無条件で確認欄全体をやってしまう
    var rangeToAutofit = leftHeaderCell.Resize(1 + dstLeaves.length, nodeH1.tableHeaders.length);
    rangeToAutofit.Columns.AutoFit();
}
function sheetSetValues(sheet, nodeH1, sheetData, templateData) {
    var table = templateData.checkSheet.table;

    var maxItemWidth = CL.getMaxItemWidth(nodeH1);
    var totalItemWidth = _.sum(maxItemWidth);
    var checkHeaders = CL.getCheckHeaders(nodeH1, table);
    var checkCellsWidth = checkHeaders.length;

    var dstLeaves = CL.GetLeafNodes(nodeH1);
    var dstHeight = dstLeaves.length;

    var idToY = {};
    for (var i = 0; i < dstLeaves.length; i++) {
        idToY[dstLeaves[i].id] = i;
    }
    
    var dstHeaders = checkHeaders.concat(table.other.headers);

    var headerNotFound = [];
    var srcHeaderToX = getSrcHeaderToX(dstHeaders, nodeH1, sheetData, headerNotFound);

    // save 対象が含まれてさえいれば良いので
    var otherCellsWidth = (table.other.indicesToSave.length === 0) ? 0 : _.max(table.other.indicesToSave) + 1;

    var dstColumn = table.ul.column + totalItemWidth;
    var dstWidth = checkCellsWidth + otherCellsWidth;
    var dstRange = sheet.Cells(table.row, dstColumn).Resize(dstHeight, dstWidth);
    var dstArray2d = CL.RangeToValueArray2d(dstRange);

    for (var id in sheetData.items) {
        var item = sheetData.items[id];
        var y = idToY[id];

        if (_.isUndefined(y)) {
            // TODO: 一つでも入力があれば notFound に。 null は delete
            continue;
        }

        for (var header in item.values) {
            var x = srcHeaderToX[header];
            if (_.isUndefined(x)) {
                // TODO: notFound に
                continue;
            }

            var value = item.values[header];
            dstArray2d[y][x] = (value === null) ? undefined : value;
        }
    }

    // check cell の左からの index
    var indicesToSave = _.range(0, checkCellsWidth);
    table.other.indicesToSave.forEach(function(element, index, array) {
        indicesToSave.push(checkCellsWidth + element);
    });

    var shouldApply = [];
    for (var i = 0; i < indicesToSave.length; i++) {
        shouldApply[indicesToSave[i]] = true;
    }

    applyArrayToRange(dstRange, dstArray2d, shouldApply);

    // 確認欄だけ header 込みで autofit
    var rangeToAutofit = sheet.Cells(table.row - 1, dstColumn).Resize(1 + dstHeight, checkCellsWidth);
    rangeToAutofit.Columns.AutoFit();
    // 元の幅より細くならないようにしておく
    var defaultInputColumnWidth = table.input.columnWidth;
    for (var i = 0; i < checkCellsWidth; i++) {
        if (rangeToAutofit.Columns(1 + i).ColumnWidth < defaultInputColumnWidth) {
            rangeToAutofit.Columns(1 + i).ColumnWidth = defaultInputColumnWidth;
        }
    }
}

// 指定された列だけ分割しながらできるだけまとめて入力
// shouldApply は列ごとの入力するしないの bool 配列
function applyArrayToRange(dstRange, srcArray, shouldApply)
{
    var shouldApply = shouldApply.concat(false);   // 番兵

    var offsetWidth = [];
    for (var i = 0, offset = 0; i < shouldApply.length; i++)
    {
        if (!shouldApply[i])
        {
            var width = i - offset;
            if (width > 0)
            {
                offsetWidth.push({
                    offset: offset,
                    width: width
                });
            }
            offset = i + 1;
        }
    }

    var height = srcArray.length;
    for (var i = 0; i < offsetWidth.length; i++)
    {
        var ow = offsetWidth[i];
        var offsetX = ow.offset;
        var width = ow.width;
        range = dstRange.Offset(0, offsetX).Resize(height, width);
        array = getRectArray2d(srcArray, offsetX, 0, width, height);
        range.Value = jsArray2dToSafeArray2d(array);
    }

}

// TODO: CLCommonあたりに移動
// 確認欄より右側の保存対象列について、隣接した列を横方向に連結したグループの配列にしておく
function getOffsetWidthFromIndices(indices)
{
    var f = [];
    for (var i = 0; i < indices.length; i++) {
        f[indices[i]] = true;
    }
    f = f.concat(false);   // 番兵

    var offsetWidth = [];
    for (var i = 0, offset = 0; i < f.length; i++)
    {
        if (!f[i])
        {
            var width = i - offset;
            if (width > 0)
            {
                offsetWidth.push({
                    offset: offset,
                    width: width
                });
            }
            offset = i + 1;
        }
    }

    return offsetWidth;
}


function indexSheetSetValues_old(root, indexSheet, indexSheetData)
{
    var headerRow = indexSheet.Range(root.headerAddress).Row;
    var headerCellColumn = indexSheet.Range(root.headerAddress).Column;
    var leftHeaderCell = getFirstCellInRow(indexSheet, headerRow);
    var rightHeaderCell = getLastCellInRow(indexSheet, headerRow);
    var headerCells = indexSheet.Range(leftHeaderCell, rightHeaderCell);
    var headers = headerCells.Value.toArray();
    var numColumns = headers.length;
    var numRows = root.children.length;
    // 最左列は番号という前提で
    var indices = leftHeaderCell.Offset(1, 0).Resize(numRows, 1).Value;
    indices = (numRows === 1) ? [ indices ] : indices.toArray();
    var dstRange = leftHeaderCell.Offset(1, 0).Resize(numRows, numColumns);
    var dstArray2d = CL.RangeToValueArray2d(dstRange);
    var shouldApply = [];
    xEach(headerCells, function(c)
    {
        // シート名のセルは当然保存対象外
        if (c.Column === headerCellColumn)
        {
            shouldApply.push(false);
            return;
        }

        // 数式は出力しない
        if (c.Offset(1, 0).HasFormula)
        {
            shouldApply.push(false);
            return;
        }

        var headerName = c.Text;
        // 見出しが空欄の列は保存対象外
        if (!headerName)
        {
            shouldApply.push(false);
            return;
        }

        // header の text が!で挟まれてる列は保存対象外
        if (/^\!.*\!$/.test(headerName))
        {
            shouldApply.push(false);
            return;
        }

        shouldApply.push(true);
    });

    // root.children から直接 findIndex とかを呼ぶとエラーが出るので、回避
    // 理由は把握できてない
    function getH1Index(nodeH1) {
        for (var i = 0; i < root.children.length; i++)
        {
            if (root.children[i] === nodeH1)
            {
                return i;
            }
        }
        return -1;
    }

    var idToNodeH1 = {};
    for (var i = 0; i < root.children.length; i++)
    {
        var nodeH1 = root.children[i];
        idToNodeH1[nodeH1.id] = nodeH1;
    }

    var items = indexSheetData.items;
    for (var id in items)
    {
        var item = items[id];
        var nodeH1 = idToNodeH1[id];

        if (typeof nodeH1 === 'undefined')
        {
            // TODO: values がすべて undefined じゃなければ notfound へ
            continue;
        }

        var index = getH1Index(nodeH1);
        //var index = root.children.findIndex(function(element, index, array) {
        //    return (element.id === nodeH1.id);
        //});
        var y = indices.indexOf(index + 1);    // Excel側は 1 origin なので +1
        var arrayX = dstArray2d[y];

        var values = item.values;
        for (header in values)
        {
            var x = headers.indexOf(header);
            if (x === -1)
            {
                // TODO: undefined じゃなければ notFound に
                continue;
            }

            var value = values[header];
            arrayX[x] = (value === null) ? undefined : value;
        }
        
    }

    applyArrayToRange(dstRange, dstArray2d, shouldApply);
}
function indexSheetSetValues(root, indexSheet, indexSheetData, templateData)
{
    var table = templateData.indexSheet.table;

    var height = root.children.length;
    var width = _.max(table.indicesToSave) + 1;
    var dstRange = indexSheet.Cells(table.row, table.column).Resize(height, width);
    var dstArray2d = CL.rangeToValueArray2d(dstRange);

    // XXX: 最左列は番号という前提で
    var h1Indices = CL.array2dTransposed(dstArray2d)[0];

    var idToNodeH1 = {};
    for (var i = 0; i < root.children.length; i++) {
        var nodeH1 = root.children[i];
        idToNodeH1[nodeH1.id] = nodeH1;
    }

    var items = indexSheetData.items;
    for (var id in items) {
        var item = items[id];
        var nodeH1 = idToNodeH1[id];

        if (_.isUndefined(nodeH1)) {
            // TODO: values がすべて undefined じゃなければ notfound へ
            continue;
        }

        // Array.prototype.findIndex() は使えないようなので、 lodash のを使う
        var index = _.findIndex(root.children, function(element) {
            return (element.id === nodeH1.id);
        });
        var y = h1Indices.indexOf(index + 1);    // Excel側は 1 origin なので +1
        var arrayX = dstArray2d[y];

        var values = item.values;
        for (header in values)
        {
            var x = table.headers.indexOf(header);
            if (x === -1)
            {
                // TODO: undefined じゃなければ notFound に
                continue;
            }

            var value = values[header];
            arrayX[x] = (value === null) ? undefined : value;
        }
        
    }

    var shouldApply = [];
    for (var i = 0; i < table.indicesToSave.length; i++) {
        shouldApply[table.indicesToSave[i]] = true;
    }

    applyArrayToRange(dstRange, dstArray2d, shouldApply);
}

function indexSheetSetVariables_old(root, indexSheet, indexSheetData)
{
    var variableData = indexSheetData.variables;
    for (var key in variableData)
    {
        var address = root.variables[key];
        if (typeof address === "undefined")
        {
            // TODO: notFound
            continue;
        }
        var cell = indexSheet.Range(address);

        // XXX: とりあえず Date は面倒なので対応しない
        // XXX: もしやる場合は JSON.parse() 後に new Date(Date.parse(v)) とか…？

        cell.Value = variableData[key];
    }
}
function indexSheetSetVariables(root, indexSheet, indexSheetData, templateData)
{
    var variableData = indexSheetData.variables;
    for (var key in variableData)
    {
        var address = templateData.indexSheet.variables[key].address;
        if (_.isUndefined(address))
        {
            // TODO: notFound
            continue;
        }
        var cell = indexSheet.Range(address);

        // XXX: とりあえず Date は面倒なので対応しない
        // XXX: もしやる場合は JSON.parse() 後に new Date(Date.parse(v)) とか…？

        cell.Value = variableData[key];
    }
}

function bookSetDataToIndexSheet(root, book, data, templateData)
{
    var indexSheet = CL.getIndexSheet(book, root);
    var indexSheetData = data.indexSheet;

    indexSheetSetValues(root, indexSheet, indexSheetData, templateData);
    indexSheetSetVariables(root, indexSheet, indexSheetData, templateData);
}

function bookSetDataToCheckSheets(root, book, data, templateData)
{
    var idToNodeH1 = {};
    for (var i = 0; i < root.children.length; i++) {
        var nodeH1 = root.children[i];
        idToNodeH1[nodeH1.id] = nodeH1;
    }

    for (var id in data.checkSheet.sheets) {
        var sheetData = data.checkSheet.sheets[id];
        var nodeH1 = idToNodeH1[id];

        if (_.isUndefined(nodeH1)) {
            // TODO: notFound に
            continue;
        }

        var sheet = findSheetByName(book, nodeH1.text);

        sheetSetValues(sheet, nodeH1, sheetData, templateData);
    }
}

function bookSetData(root, book, data, templateData)
{
    bookSetDataToCheckSheets(root, book, data, templateData);
    bookSetDataToIndexSheet(root, book, data, templateData);
}

bookSetData(root, book, data, templateData);

// いろいろ怖いんで毎回バックアップはとっておく
var backupFolderName = "bak/load";
CL.makeBackupFile(xlsFilePath, backupFolderName);
book.Save();

// Excelは閉じない
excel.Visible = true;
excel.ScreenUpdating = true;

WScript.Quit();
