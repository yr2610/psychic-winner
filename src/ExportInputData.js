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
    Error("入力データを出力したいチェックリスト（Excelファイル）をドロップしてください。");
}

var filePath = WScript.Arguments.Unnamed(0);

// TODO: Excelファイルの確認

initializeExcel();
//excel.Visible = true;
//excel.ScreenUpdating = true;

var book = openBookReadOnly(filePath);

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
    getIndexSheetValues = getIndexSheetValues_old;
    getCheckSheetValuesFromSheet = getCheckSheetValuesFromSheet_old;
}


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

// Excelから入力を取り込む
var valuesToSave = {
    sourceFile: getFileInfo(filePath),
    checkSheet: {
        sheets: getCheckSheetValues(root, templateData, book)
    },
    indexSheet: {
        items: getIndexSheetValues(root, book, templateData),
        variables: getIndexSheetVariables(root, templateData, book)
    }
};

// Excelを閉じる
finalizeExcel();

//function replacer (k,v) {
// return (typeof v === "undefined") ? null : v;
//}

// 入力取り込み済みの
//var newJson = JSON.stringify(valuesToSave, undefined, 2);

var yamlOptions = {sortKeys:true};
var newJson = jsyaml.safeDump(valuesToSave, yamlOptions);

//var newJson = JSON.stringify(valuesToSave, replacer, 2);

var fso = new ActiveXObject( "Scripting.FileSystemObject" );

function GetUserName()
{
    var network = new ActiveXObject("WScript.Network");

    return network.UserName;
}

function GetUserNameInitial(isFirstNameInitial, isLastNameInitial, separator)
{
    isFirstNameInitial = (typeof isFirstNameInitial === "undefined") ? false : isFirstNameInitial;
    isLastNameInitial = (typeof isLastNameInitial === "undefined") ? false : isLastNameInitial;
    separator = (typeof separator === "undefined") ? "" : separator;

    var userName = GetUserName();

    var userNameMatch = userName.match(/^([A-Za-z]+)_([A-Za-z0-9]+)$/);
    if (!userNameMatch)
    {
        return userName;
    }

    var fn = userNameMatch[2];
    var ln = userNameMatch[1];
    if (isFirstNameInitial) fn = fn.slice(0, 1);
    if (isLastNameInitial) ln = ln.slice(0, 1);

    return fn + separator + ln;
}

//  ファイルを書き込み専用で開く
var file = fso.GetFile(filePath);
var srcFileDateLastModified = new Date(file.DateLastModified);
var date = CL.yyyymmddhhmmss(srcFileDateLastModified).slice(2, 12);

var outFilename = fso.GetBaseName(filePath) + "-" + GetUserNameInitial(true, true) + date + ".sav";
var outfilePath = fso.BuildPath(fso.GetParentFolderName(filePath), outFilename);

CL.writeTextFileUTF8(newJson, outfilePath);

WScript.Echo("savファイル(" + outFilename + ")を出力しました");
WScript.Quit();

// =============================================================


// マージツールを使ってマージされることを想定して（コンフリクト時にマージしやすいように）、
// 以下のような容量軽視な仕様にしておく
// * 横並びのセルを配列でなく、Object で保存（全セル１行に出力でなく、１セルを１行に出力したい）
// * 未入力のセルもすべて出力（{} と１行にまとめられてしまうのを避けるため。以下の「マージに失敗するケース」回避）
// * 丸ごと未入力のシートも出力されるようにする（配列が空っぽになって [] となるのを避けるため。マージ失敗回避）
// また、最低限の可読性のため（だけ）に、シート名と項目の一番右のセルの内容を出力しておく（読み込み時には無視する）
// # 自動マージでJSONが壊れるケース
// ---
// [parent]
// "values": {
//     "備考": "追加配信②（4/26）では対象外",
//     "": null
// },
// [A]
// "values": {},
// [B]
// "values": {
//     "確認欄(ボス)": "○ / Rev.10281 / 品管 伊藤 / 170517",
//     "備考": "追加配信②（4/26）では対象外",
//     "": null
// },
// [merged]
// "values": {},
//     "確認欄(ボス)": "○ / Rev.10281 / 品管 伊藤 / 170517",
// ---
function getCheckSheetValuesFromSheet_old(nodeH1, book)
{
    var name = nodeH1.text;
    var sheet = findSheetByName(book, name);
//        var maxItemWidth = getMaxItemWidth(nodeH1);
//        var totalItemWidth = sum(maxItemWidth);
    var leftHeaderCell = sheet.Range(nodeH1.leftCheckHeaderCellAddress);
    var rightHeaderCell = getLastCellInRow(sheet, leftHeaderCell.Row);
    var headerRange = sheet.Range(leftHeaderCell, rightHeaderCell);
    var numHeaders = headerRange.Columns.Count;
    var leafNodes = CL.getLeafNodes(nodeH1);
    var checkCellRows = leafNodes.length;
    var checkRange = headerRange.Offset(1, 0).Resize(checkCellRows, numHeaders);
    var checkCellArray = CL.RangeToValueArray2d(checkRange);
    var headers = headerRange.Value.toArray();
    var hasFormulaColumns = [];

    xEach(headerRange, function(cell)
    {
        // 列が数式かどうか
        hasFormulaColumns.push(cell.Offset(1, 0).HasFormula);
    });

    (function () {
        for (var i = 0; i < nodeH1.tableHeaders.length; i++)
        {
            var id = nodeH1.tableHeaders[i].id;
            headers[i] = id + ". " + headers[i];
        }
    })();

    var sheetValues = {
        text: nodeH1.text,
        items: {}
    };
    for (var y = 0; y < checkCellRows; y++)
    {
        var node = leafNodes[y];
        var item = {
            text: node.text,
            values: {}
        };
        for (var x = 0; x < numHeaders; x++)
        {
            if (hasFormulaColumns[x])
            {
                continue;
            }
            // JSON stringify で消されないように、とりあえず null を入れておく
            var v = checkCellArray[y][x];
            v = (typeof v === "undefined") ? null : v;
            item.values[headers[x]] = v;
        }
        sheetValues.items[node.id] = item;
    }

    return sheetValues;
}

function getCheckSheetValuesFromSheet(nodeH1, book, templateData)
{
    var table = templateData.checkSheet.table;

    var name = nodeH1.text;
    var sheet = findSheetByName(book, name);
    var maxItemWidth = CL.getMaxItemWidth(nodeH1);
    var totalItemWidth = _.sum(maxItemWidth);
    var checkHeaders = CL.getCheckHeaders(nodeH1, table);
    var checkCellsWidth = checkHeaders.length;

    var leafNodes = CL.getLeafNodes(nodeH1);
    var checkCellRows = leafNodes.length;

    // save 対象が含まれてさえいれば良いので
    var otherCellsWidth = (table.other.indicesToSave.length === 0) ? 0 : _.max(table.other.indicesToSave) + 1;

    var saveRangeColumn = table.ul.column + totalItemWidth;
    var saveRangeWidth = checkCellsWidth + otherCellsWidth;
    var saveRange = sheet.Cells(table.row, saveRangeColumn).Resize(checkCellRows, saveRangeWidth);
    var saveCellArray = CL.RangeToValueArray2d(saveRange);

    // check cell の左からの index
    var indicesToSave = _.range(0, checkCellsWidth);
    table.other.indicesToSave.forEach(function(element, index, array) {
        indicesToSave.push(checkCellsWidth + element);
    });

    var headers = nodeH1.tableHeaders.map(function(n) {
        return n.id + ". " + n.name;
    })
    .concat(table.other.headers);

    var sheetValues = {
        text: nodeH1.text,
        items: {}
    };

    for (var y = 0; y < checkCellRows; y++)
    {
        var node = leafNodes[y];
        var item = {
            text: node.text,
            values: {}
        };
        indicesToSave.forEach(function(x, index, array) {
            // JSON stringify で消されないように、とりあえず null を入れておく
            var v = saveCellArray[y][x];
            item.values[headers[x]] = (_.isUndefined(v)) ? null : v;
        });

        sheetValues.items[node.id] = item;
    }

    return sheetValues;
}

function getCheckSheetValues(root, templateData, book)
{
    var sheets = {};

    for (var i = 0; i < root.children.length; i++)
    {
        var nodeH1 = root.children[i];
        var sheetValues = getCheckSheetValuesFromSheet(nodeH1, book, templateData);

        sheets[nodeH1.id] = sheetValues;
    }

    return sheets;
}

function getIndexSheetVariables(root, templateData, book)
{
    var indexSheet = CL.getIndexSheet(book, root);
    var variables = {};

    var getAddress = function (key) {
        if (!(key in templateData.indexSheet.variables)) {
            return undefined;
        }
        return templateData.indexSheet.variables[key].address;
    };
    // TODO: 不要になったら削除
    if (_.isUndefined(templateData)) {
        getAddress = function (key) {
            return root.variables[key];
        };
    }

    // 変数名が _ で始まり、その次が大文字の変数は値を取り込む
    for (var key in root.variables)
    {
        if (!/^_[A-Z].*/.test(key))
        {
            continue;
        }

        var address = getAddress(key);
        if (_.isUndefined(address)) {
            continue;
        }
        var cell = indexSheet.Range(address);

        // XXX: 日付を取得するために Text にしてお茶を濁す。本当はDateをDateとして扱うべき
        //root.variables[key] = cell.Value;
        variables[key] = cell.Text;
    }
    return variables;
}

// 日付は対応しない
function getIndexSheetValues_old(root, book)
{
    var items = {};

    var indexSheet = CL.getIndexSheet(book, root);
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

    for (var i = 0; i < root.children.length; i++)
    {
        var nodeH1 = root.children[i];
        var sheetValues = {
            text: nodeH1.text,
            values: {}
        };

        var index = getH1Index(nodeH1);
        //var index = root.children.findIndex(function(element, index, array) {
        //    return (element.id === nodeH1.id);
        //});
        var y = indices.indexOf(index + 1);    // Excel側は 1 origin なので +1
        var arrayX = dstArray2d[y];

        for (var x = 0; x < arrayX.length; x++)
        {
            if (!shouldApply[x])
            {
                continue;
            }

            var v = arrayX[x];

            if (typeof v === "undefined")
            {
                v = null;
            }

            sheetValues.values[headers[x]] = v;
        }

        items[nodeH1.id] = sheetValues;
    }

    return items;
}
function getIndexSheetValues(root, book, templateData)
{
    var table = templateData.indexSheet.table;

    var items = {};

    var indexSheet = CL.getIndexSheet(book, root);

    //var minIndex = _.min(table.indicesToSave);
    var maxIndex = _.max(table.indicesToSave);

    var height = root.children.length;
    var width = maxIndex + 1;
    var saveRange = indexSheet.Cells(table.row, table.column).Resize(height, width);
    var saveCellArray2d = CL.rangeToValueArray2d(saveRange);

    // XXX: 最左列は番号という前提で
    var h1Indices = CL.array2dTransposed(saveCellArray2d)[0];

    for (var i = 0; i < root.children.length; i++)
    {
        var nodeH1 = root.children[i];
        var sheetValues = {
            text: nodeH1.text,
            values: {}
        };

        // Array.prototype.findIndex() は使えないようなので、 lodash のを使う
        var index = _.findIndex(root.children, function(element) {
            return (element.id === nodeH1.id);
        });
        var y = h1Indices.indexOf(index + 1);    // Excel側は 1 origin なので +1
        var arrayX = saveCellArray2d[y];

        table.indicesToSave.forEach(function(x, index, array) {
            var v = arrayX[x];

            sheetValues.values[table.headers[x]] = _.isUndefined(v) ? null : v;
        });

        items[nodeH1.id] = sheetValues;
    }

    return items;
}
