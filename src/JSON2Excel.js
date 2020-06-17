function Error(message) {
    shell.Popup(message, 0, "エラー", ICON_EXCLA);
    WScript.Quit();
}

function findAllValuesInRange(range, targetValue)
{
    var cell = range.Find(targetValue, range.Cells(1, 1), Excel.xlValues, Excel.xlWhole, Excel.xlByRows, Excel.xlNext, true, true);
    if (!cell) {
        return null;
    }

    var cells = [ cell ];
    var firstAddress = cell.Address;
    while (true) {
        cell = range.FindNext(cell);
        if (!cell || cell.Address === firstAddress) {
            break;
        }
        cells.push(cell);
    }

    return cells;
}

// TODO: root なしでも記述が文法どおりなら全部パース対象とするべきか
function parseIndexSheet(indexSheet, checkSheet, root) {
    var data = {};

    // usedrange を全部調べて、 $ ではさまれてるの全部保存しておく
    // 現状は、「このシートはそのままで json だけ更新される」という状況はないけど、将来必要になりそうな雰囲気がないこともないので、一応対応
    data.variables = {};
    for (var key in root.variables) {
        var cells = findAllValuesInRange(indexSheet.Cells, "$" + key + "$");

        // セルが見つからない
        if (!cells) {
            continue;
        }

        // １文字目が _ ならアドレスを保存
        if (key.charAt(0) === "_" && cells.length >= 2) {
            var message = "保存対象の変数のセル " + key + " が複数存在します";
            throw (new Error(message));
        }

        data.variables[key] = {};
        data.variables[key].address = cells.map(function(c) {
            return c.Address(false, false);
        });
    }

    var checkSheetName = checkSheet.Name;
    var templateCell = indexSheet.Cells.Find(checkSheetName, indexSheet.Cells(1, 1), Excel.xlValues, Excel.xlWhole, Excel.xlByRows, Excel.xlNext, true);
    if (!templateCell) {
        var message = "シート " + indexSheet.Name + " にチェックシート名 " + checkSheetName + " のセルが存在しません";
        throw (new Error(message));
    }
    var leftHeaderCell = getFirstCellInRow(indexSheet, templateCell.Row - 1);
    var rightHeaderCell = getLastCellInRow(indexSheet, templateCell.Row - 1);
    var headerCells = indexSheet.Range(leftHeaderCell, rightHeaderCell);

    // テーブルの情報を取得しておく
    data.table = {};
    data.table.address = headerCells.Offset(1, 0).Address(false, false);
    data.table.row = headerCells.Offset(1, 0).Row;
    data.table.column = headerCells.Column;
    data.table.headers = headerCells.Value.toArray();
    data.table.columnWidth = [];    // 不要だとは思うけど、念のため
    data.table.mainIndex = templateCell.Column - leftHeaderCell.Column;
    data.table.indicesToSave = [];

    // 保存対象のヘッダーの列（表の左端からのindex）のdictionaryを作成しておく
    xEach(headerCells, function(cell)
    {
        data.table.columnWidth.push(cell.ColumnWidth);

        // 数式の列は対象外
        if (cell.Offset(1, 0).HasFormula)
        {
            return;
        }

        // header の text が!で挟まれてる列は対象外
        if (/^\!.+\!$/.test(cell.Value)) {
            return;
        }

        var index = cell.Column - leftHeaderCell.Column;

        // シート名の列は対象外
        if (index === data.table.mainIndex) {
            return;
        }

        data.table.indicesToSave.push(index);
    });
    
    return data;
}
;;;

function parseCheckSheet(sheet) {
    // セルの挿入とか削除でずれると確認欄とかの列の幅はそのままで内容だけずれる雰囲気なので、何とかする
    // 入力欄の列の幅を保存

    var data = {};

    // 小文字はNGとする
    var cellH1 = sheet.Cells.Find("[H1]", sheet.Cells(1, 1), Excel.xlValues, Excel.xlWhole, Excel.xlByRows, Excel.xlNext, true);
    if (!cellH1)
    {
        var message = "シートに H1 セルが存在しません";
        throw (new Error(message));
    }
    data.h1 = {};
    data.h1.address = cellH1.Address(false, false);

    data.table = {};

    var cellUL = sheet.Cells.Find("[UL]", sheet.Cells(1, 1), Excel.xlValues, Excel.xlWhole, Excel.xlByRows, Excel.xlNext, true);
    if (!cellUL)
    {
        var message = "シートに UL セルが存在しません";
        throw (new Error(message));
    }
    data.table.row = cellUL.Row;

    data.table.ul = {
        column: cellUL.Column,
        columnWidth: cellUL.ColumnWidth
    };

    // 確認欄
    // これは仕様とする
    var cellInput = cellUL.Offset(0, 1);
    var resultColumnID = (function(s) {
        var columnIDMatch = s.match(/^#([A-Za-z_]\w+)$/);
        if (columnIDMatch !== null) {
            return columnIDMatch[1];
        }
        return void 0;
    })(cellInput.Offset(1 - cellInput.Row).Text);

    data.table.input = {
        column: cellInput.Column,
        header: cellInput.Offset(-1, 0).Value,
        columnID: resultColumnID,
        columnWidth: cellInput.ColumnWidth
    };

    // 確認欄より右側の保存対象列の情報
    // TODO: 確認欄のID取得。というより indicesToSave はこの区切りなく管理するべき
    var leftHeaderCell = cellInput.Offset(-1, 1);
    var rightHeaderCell = getLastCellInRow(sheet, leftHeaderCell.Row);
    var headerCells = sheet.Range(leftHeaderCell, rightHeaderCell);

    data.table.other = {
        column: headerCells.Column,
        headers: headerCells.Value.toArray(),
        columnWidth: [],
        columnID: {},
        indicesToSave: []
    };
    xEach(headerCells, function(cell) {
        data.table.other.columnWidth.push(cell.ColumnWidth);

        // 数式の列は対象外
        if (cell.Offset(1, 0).HasFormula) {
            return;
        }

        var columnIDCell = cell.Offset(1 - cell.Row).Text;
        //if (_.isUndefined(columnIDCell)) {
        //    return;
        //}

        // ID が存在する列だけを対象とする
        var columnIDMatch = columnIDCell.match(/^#([A-Za-z_]\w+)$/);
        if (columnIDMatch === null) {
            return;
        }

        var index = cell.Column - leftHeaderCell.Column;
        data.table.other.indicesToSave.push(index);

        var columnID = columnIDMatch[1];
        data.table.other.columnID[columnID] = index;

    });

    return data;
}


function clearMarksInCheckSheet(sheet, data)
{
    // H1 セルは必ず上書きされるのでクリア不要
    //sheet.Range(data.h1.address).SpecialCells(Excel.xlCellTypeConstants).ClearContents();

    var leftTableCell = sheet.Cells(data.table.row, data.table.ul.column);
    var rightTableCell = sheet.Cells(data.table.row, data.table.other.column + data.table.other.headers.length);

    // 数式はクリアしないようにする
    // 定数セルは必ず存在しているので(H2等)、エラー補足はしなくてOK
    sheet.Range(leftTableCell, rightTableCell).SpecialCells(Excel.xlCellTypeConstants).ClearContents();
}

var shell = new ActiveXObject("WScript.Shell");
var shellApplication = new ActiveXObject("Shell.Application");
var fso = new ActiveXObject( "Scripting.FileSystemObject" );
var stream = new ActiveXObject("ADODB.Stream");

// Performance を取得
var htmlfile = WSH.CreateObject("htmlfile");
htmlfile.write('<meta http-equiv="x-ua-compatible" content="IE=Edge"/>');
var performance = htmlfile.parentWindow.performance;
htmlfile.close();

if (( WScript.Arguments.length != 1 ) ||
    ( WScript.Arguments.Unnamed(0) == ""))
{
    Error("Excelのチェックリストを生成する .json ファイルをドロップしてください。");
}

var filePath = WScript.Arguments.Unnamed(0);

if (fso.GetExtensionName(filePath) != "json")
{
    Error(".json ファイルをドロップしてください。");
}

/**/
stream.Type = adTypeText;
stream.charset = "utf-8";
stream.Open();
stream.LoadFromFile(filePath);
var sJSON = stream.ReadText(adReadAll);
stream.Close();
/*/
//  ファイルを読み取り専用で開く
var file = fso.OpenTextFile(filePath, FORREADING, true, TRISTATE_FALSE);

var sJSON = file.Readall();

//  ファイルを閉じる
file.Close();
/**/

var root = JSON.parse(sJSON);


var kindH = "H";
var kindUL = "UL";

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

// ----------------

// 古い仕様に対応
if (root.variables.sheetname) {
    root.variables.indexSheetname = root.variables.sheetname;
    delete root.variables.sheetname;
}

if (root.variables.indexSheetname) {
    if (root.variables.indexSheetname.length >= 32) {
        Error("indexSheetname( " + root.variables.indexSheetname + " ）が32文字以上です。");
    }
    if (/[\:\\\?\[\]\/\*：￥？［］／＊]/g.test(root.variables.indexSheetname)) {
        Error("indexSheetname( " + root.variables.indexSheetname + " ）に使えない文字が含まれています。");
    }
}

/**
// デフォルトではソーステキストと同じ場所にある checklist_template.xlsx という名前のExcelをテンプレートとして使うようにしておく
var defaultTemplateName = "checklist_template.xlsx";
var templatesDirectory = fso.GetParentFolderName(filePath);
/*/
// デフォルトではツール置き場の templates/default.xlsx という名前のExcelをテンプレートとして使うようにしておく
var defaultTemplateName = "default.xlsx";
var templateName = root.variables.templateFilename ? root.variables.templateFilename : defaultTemplateName;
var currentDirectory = fso.GetParentFolderName(WScript.ScriptFullName);
var templatesDirectory = fso.BuildPath(currentDirectory, "templates");
var localTemplatesDirectory = fso.BuildPath(fso.GetParentFolderName(filePath), "templates");
var templatePath = fso.BuildPath(templatesDirectory, templateName);
if (fso.FolderExists(localTemplatesDirectory)) {
    var localTemplatePath = fso.BuildPath(localTemplatesDirectory, templateName);
    if (fso.FileExists(localTemplatePath)) {
        templatePath = localTemplatePath;
    }
}
/**/

// 出力（SaveAs）ファイル名を先に確認
var dstFilename = root.variables.outputFilename ? root.variables.outputFilename : fso.GetBaseName(filePath);
dstFilename += "-" + CL.yyyymmddhhmmss(new Date()).slice(2, -2);
dstFilename += "." + fso.GetExtensionName(templatePath);
var dstFolderName = fso.BuildPath(fso.GetParentFolderName(filePath), "build");
CL.createFolder(dstFolderName);
var dstFilePath = fso.BuildPath(dstFolderName, dstFilename);
try {
    if (root.variables.outputFilename) {
        fso.CreateTextFile(dstFilePath, true);
        fso.DeleteFile(dstFilePath);
    }
}
catch (e) {
    Error("出力ファイル名（ " + root.variables.outputFilename + " ）が不正です。");
}

initializeExcel();
excel.Visible = true;
excel.ScreenUpdating = false;

var book = openBookReadOnly(templatePath);

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

// TODO: 複数のタイプから選択できるようにしたい（ソースファイル内でシート毎に指定）
var templateSheet = book.Worksheets("checksheet");    // XXX:
// TODO: ソースファイルに設定して変更できるようにしたい
var indexSheet = book.Worksheets("index");  // XXX:


var indexNode = root;
if (indexNode.variables.indexSheetname) {
    indexSheet.Name = indexNode.variables.indexSheetname;
}

var indexSheetData = parseIndexSheet(indexSheet, templateSheet, root);
var checkSheetData = parseCheckSheet(templateSheet);
var templateData = {
    indexSheet: indexSheetData,
    checkSheet: checkSheetData
};

// 情報抜き出したらそれに基づいてセルの内容削除
clearMarksInCheckSheet(templateSheet, checkSheetData);

for (var i = 0; i < root.children.length; i++)
{
    var nodeH1 = root.children[i];
    var name = nodeH1.text;
    excel.StatusBar = "シート作成中: " + name;

    if (findSheetByName(book, name))
    {
        // TODO: 上書き生成する(OK)か、今のを残す(Cancel)か確認する
        WScript.Echo("すでにシート [" + name + "] は存在しているためスキップします");
        continue;
    }

    // コピー元が非表示だとアクセスが面倒くさいので、先頭にコピーして表示させてから末尾に移動
    templateSheet.Copy(book.Worksheets(1));
    var sheet = book.Worksheets(1);
    sheet.Visible = true;
    sheet.Name = name;
    sheet.Move(null, book.Worksheets(book.Worksheets.Count));

    render(sheet, nodeH1, checkSheetData);
}


excel.StatusBar = "シート作成中: " + indexSheet.Name;

//excel.ScreenUpdating = true;
{
    // １文字目が大文字の変数はシート内の同名（$FOO$ 形式）のセルに値を埋め込む
    for (var key in root.variables)
    {
        // １文字目が大文字、または _ で始まり、その次が大文字じゃない
        if (!/^_?[A-Z].*/.test(key))
        {
            continue;
        }
        // セルが見つからない
        var cell = indexSheet.Cells.Find("$" + key + "$", sheet.Cells(1, 1), Excel.xlValues, Excel.xlWhole, Excel.xlByRows, Excel.xlNext, true);
        if (!cell)
        {
            continue;
        }

        cell.Value = root.variables[key];
    }

    var templateName = templateSheet.Name;

    var templateCell = indexSheet.Cells.Find(templateName, indexSheet.Cells(1, 1), Excel.xlValues, Excel.xlWhole, Excel.xlByRows, Excel.xlNext, true);
    if (!templateCell)
    {
        Error("シートにテンプレート名 " + templateName + " のセルが存在しません");
    }

    var leftHeaderCell = getFirstCellInRow(indexSheet, templateCell.Row - 1);
    var rightHeaderCell = getLastCellInRow(indexSheet, templateCell.Row - 1);
    var headerCells = indexSheet.Range(leftHeaderCell, rightHeaderCell);

    // ヘッダーのセルのaddressのdictionaryを作成しておく
    var headerCellColumns = {};
    xEach(headerCells, function(cell)
    {
        // 数式の列は対象外
        if (cell.Offset(1, 0).HasFormula)
        {
            return;
        }

        headerCellColumns[cell.Value] = cell.Column;
    });

    // 表の行追加
    // ルートの直下はすべて H1 という前提
    var leftCell = leftHeaderCell.Offset(1, 0);
    var rightCell = rightHeaderCell.Offset(1, 0);
    var tableRow = indexSheet.Range(leftCell, rightCell);

    if (root.children.length >= 2)
    {
        indexSheet.Range(templateCell.Offset(1, 0), templateCell.Offset(root.children.length - 1, 0)).EntireRow.Insert(Excel.xlDown);

        var fillRange = indexSheet.Range(leftCell, rightCell.Offset(root.children.length - 1, 0));
        tableRow.AutoFill(fillRange, Excel.xlFillSeries);
    }
    indexSheet.Select();

    var regExp = new RegExp(templateName, "g");
    var replaceRange = tableRow;
    var columnsToAutofit = {};
    for (var i = 0; i < root.children.length; i++)
    {
        var nodeH1 = root.children[i];
        var name = nodeH1.text;
        //var sheet = book.Worksheets(name);

        var afterName = name;
        // single quoted
        var afterNameSQ = "'" + afterName + "'";

        xEach(replaceRange, function(cell)
        {
            //WScript.Echo(afterNameSQ + ":" + cell.Address(false, false) + ":" + cell.Formula);
            if (cell.HasFormula)
            {
                var formula0 = cell.Formula;
                if (regExp.test(formula0))
                {
                    cell.Formula = formula0.replace(regExp, afterNameSQ);
                }
            }
            else
            {
                var value0 = cell.Value;
                if (value0)
                {
                    if (regExp.test(value0))
                    {
                        cell.Value = value0.replace(regExp, afterName);

                        // この列を記録しておいて、header含めて最後の行までautofit
                        columnsToAutofit[cell.Column] = cell.Row;
                    }
                }

            }

            if (cell.Value === afterName) {
                var targetSheet = findSheetByName(book, afterName);
                // XXX: [H1] のアドレスを保存しておいて、そこを見るように
                var cellH1 = targetSheet.Cells.Find(afterName, sheet.Cells(1, 1), Excel.xlValues, Excel.xlWhole, Excel.xlByRows, Excel.xlNext, true);
                // XXX: とりあえず size, bold だけ…
                var fontH1Size = cellH1.Font.Size;
                var fontH1Bold = cellH1.Font.Bold;
                var targetSubAddress = indexSheet.Name + "!" + cell.Address(true, true);
                cellH1.Hyperlinks.Add(cellH1, "", targetSubAddress, "", cellH1.Value);
                cellH1.Font.Size = fontH1Size;
                cellH1.Font.Bold = fontH1Bold;
            }

            // リンク先の置換
            if (cell.Hyperlinks.Count > 0)
            {
                var subAddress0 = cell.Hyperlinks(1).SubAddress;
                if (subAddress0)
                {
                    var subAddress1 = subAddress0.replace(regExp, afterNameSQ);

                    if (true)
                    {
                        //cell.Hyperlinks.Delete(); // これがなくても問題なさげ（特にごみデータが残ったりとかもないっぽい）
                        indexSheet.Hyperlinks.Add(cell, "", subAddress1, "", cell.Value);
                    }
                    else
                    {
                        // TODO: これだとなぜかすべてのHyperlinkが一斉に書き変わる
                        cell.Hyperlinks(1).SubAddress = subAddress1;
                    }
                }
            }

        });

        replaceRange = replaceRange.Offset(1, 0);

    }

    // autofit
    (function(){
        var row0 = headerCells.Row;
        for (var column in columnsToAutofit)
        {
            var row1 = columnsToAutofit[column];
            column = parseInt(column);
            var cell0 = indexSheet.Cells(row0, column);
            var cell1 = indexSheet.Cells(row1, column);
            indexSheet.Range(cell0, cell1).Columns.AutoFit();
        }
    })();
    
}

//DeleteTemplateRowFromIndexSheet(indexSheet, templateSheet);


function addJSONSheet(object, sheetName) {
    excel.StatusBar = sheetName + " 出力中";
    var sJSON = JSON.stringify(object, undefined, 4);
    var jsonSheet = book.Worksheets.Add();
    jsonSheet.Name = sheetName;
    var lastSheet = book.Worksheets(book.Worksheets.Count);
    var lastSheetVisible = lastSheet.Visible;
    lastSheet.Visible = true;   // 一旦 visible にしておかないと、意図した位置に移動されないっぽい
    jsonSheet.Move(null, lastSheet);
    lastSheet.Visible = lastSheetVisible;
    jsonSheet.Visible = false;

    /**/
    var sJSONArray = sJSON.split("\n");
    var startTime = performance.now();
    //var excelArray = jsArray1DToExcelRangeArray(sJSONArray, sJSONArray.length);
    var excelArray = jsArray1dColumnMajorToSafeArray2d(sJSONArray, sJSONArray.length);
    jsonSheet.Cells(1, 1).Resize(sJSONArray.length, 1) = excelArray;
    var endTime = performance.now();
    //WScript.Echo("1d:\n" + (endTime - startTime));
    /*/
    // 1行1セルで出力するとクソ重いので、一つのセルに、セルの文字数上限32767ギリギリまで詰め込む
    var row;
    for (row = 0; sJSON.length >= 0x7fff; row++)
    {
        var s = sJSON.substr(0, 0x7fff);
        var i = s.lastIndexOf("\n");
        if (i == -1)
        {
            // 0x7fff 文字あるのに改行が一つもない、みたいな特殊すぎるケースには対応しない
            Error("no cr found in JSON.");
        }
        jsonSheet.Cells(1 + row, 1).Value = sJSON.substr(0, i);
        sJSON = sJSON.substr(i + 1);
    }
    jsonSheet.Cells(1 + row, 1).Value = sJSON;
    /**/
}

// tree の状態を残しておく
addJSONSheet(root, "JSON");

// 元のtemplateのexcelの情報も残しておく
addJSONSheet(templateData, "template.json");


indexSheet.Select();

excel.DisplayAlerts = false;
templateSheet.Delete();
excel.DisplayAlerts = true;


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

function addPictureAsComment(cell, path)
{
    // path のファイルが存在しないならエラー
    (function () {
        var fso = new ActiveXObject("Scripting.FileSystemObject");

        if (!fso.FileExists(path)) {
            var relativeFilePath = getRelativePath(path, filePath, fso);
            var errorMessage = "画像ファイル\n" + relativeFilePath + "\nが存在しません";

            finalizeExcel();
            Error(errorMessage);
        }
    })();

    // 画像のサイズを知るために一旦追加する
    var shape = cell.Parent.Shapes.AddPicture(path, false, true, 0, 0, -1, -1);

    var comment = cell.AddComment();
    comment.Visible = false;
    comment.Shape.Fill.UserPicture(path);
    comment.Shape.Height = shape.Height;
    comment.Shape.Width = shape.Width;
    comment.Text(" ");  // XXX: 空文字を渡すと書き換えてくれない仕様っぽいのでダミーの半角スペース

    // サイズの取得が終わった時点で不要
    shape.Delete();
}

// textArray は [row(y)][column(x)] な 2d array を渡す
function renderUL_Recurse(node, y, cellOrigin, widthUL, groupOffset, imagePath, textArray, mergeCellMap)
{
nodeUL: {
        if (node.kind !== kindUL)
        {
            break nodeUL;
        }

        var x = groupOffset[node.group] + node.depthInGroup;
        var cell = cellOrigin.Offset(y, x);

        textArray[y][x] = node.text;

        if (node.comment)
        {
            // コメントが画像ファイル名ならコメントに画像を貼り付ける
            // 画像とコメントとの併用は不可
            //if (/^.+\.(jpg|jpeg|png|gif)$/i.test(node.comment))
            var imageMatch = node.comment.match(/^\!(.+)\!$/);
            if (imageMatch) {
                var image = imageMatch[1];
                if (imagePath) {
                    image = imagePath + "/" + image;
                }

                // ソースファイルからの相対パスにする
                var path = fso.BuildPath(fso.GetParentFolderName(filePath), image);

                addPictureAsComment(cell, path);
            }
            else
            {
                try {
                    cell.AddComment(node.comment);
                } catch (e) {
                    // XXX: 何か謎のエラーが出ることがあるので、対処
                    // XXX: エラーの原因はまったくの不明
                    // XXX: 対処といっても、いろいろやってみて、たまたまうまくいったというだけ。ちゃんとした解決法ではない。何かの拍子にまたエラーになるかも
                    cell.ClearComments();
                    cell.AddComment(node.comment);
                }
                cell.Comment.Shape.TextFrame.AutoSize = true;
            }
        }

        if (node.url)
        {
            cell.Parent.Hyperlinks.Add(cell, node.url);
        }

        var numRows = getNumLeaves(node);
        if (numRows >= 2)
        {
            var cellToMerge = cell.Resize(numRows, 1);

            // セルのマージはしない。横の罫線を消すだけ（見た目だけマージ風）
            cellToMerge.Borders(Excel.xlInsideHorizontal).LineStyle = Excel.xlNone;
        }

        // かぶらなければ別に何でも良いけど、とりあえず
        var mergeCellId = y + ", " + x;
        for (var j = y; j < y + numRows; j++)
        {
            for (var i = x; i < widthUL; i++)
            {
                mergeCellMap[j][i] = mergeCellId;
            }
        }

        // leaf
        if (node.children.length === 0)
        {
            if (node.tableData)
            {
                var valuesOffsetX = widthUL;
                // TODO: チェック列数より多い分は無視する
                for (var i = 0; i < node.tableData.length; i++)
                {
                    var td = node.tableData[i];
                    if (!td)
                    {
                        continue;
                    }

                    var dataCell = cellOrigin.Offset(y, widthUL + i);

                    // 画像ファイル名なら InputMessage の設定はせず画像を貼り付ける
                    //if (/^.+\.(jpg|jpeg|png|gif)$/i.test(td))
                    var imageMatch = td.match(/^\!(.+)\!$/);
                    if (imageMatch)
                    {
                        var image = imageMatch[1];
                        if (imagePath)
                        {
                            image = imagePath + "/" + image;
                        }

                        // ソースファイルからの相対パスにする
                        var path = fso.BuildPath(fso.GetParentFolderName(filePath), image);

                        addPictureAsComment(dataCell, path);
                        
                        continue;
                    }

                    // InputTitle は文字数制限32文字っぽいので、データは InputMessage を使う
                    dataCell.Validation.InputMessage = td;
                    if (dataCell.Validation.InputMessage !== td) {
                        // XXX: Office2016 になったら、 Insert() した場合は、Validation.InputTitle とかに代入しても反映されない謎の状態に。回避方法見つからず
                        // XXX: 回避方法が見つかるまではExcelのコメントで
                        dataCell.AddComment(td);
                        dataCell.Comment.Shape.TextFrame.AutoSize = true;
                    }
                }
            }

            y++;
        }
    }

    for (var i = 0; i < node.children.length; i++)
    {
        y = renderUL_Recurse(node.children[i], y, cellOrigin, widthUL, groupOffset, imagePath, textArray, mergeCellMap);
    }

    return y;
}

function renderInitialValues_Recurse(node, cellOrigin, y) {
    // leaf 以外で initialValues が設定されていても無視

    if (node.children.length > 0) {
        for (var i = 0; i < node.children.length; i++) {
            y = renderInitialValues_Recurse(node.children[i], cellOrigin, y);
        }
        return y;
    }

    if (!_.isUndefined(node.initialValues)) {
        var resultData = templateData.checkSheet.table.input;
        var otherData = templateData.checkSheet.table.other;

        _.forEach(node.initialValues, function(value, key) {
            if (key == resultData.columnID) {
                // XXX: result 欄1列の場合しか対応しない
                cellOrigin.Offset(y).Value = value;
            }
            else if (key in otherData.columnID) {
                // XXX: result 欄1列の場合しか対応しない
                var x = 1 + otherData.columnID[key];
                cellOrigin.Offset(y, x).Value = value;
            }
        });
    }

    return y + 1;
}


function render(sheet, nodeH1, checkSheetData)
{
    sheet.Select();

    var cellH1 = sheet.Range(checkSheetData.h1.address);
    cellH1.Value = sheet.Name;

    var cellUL = sheet.Cells(checkSheetData.table.row, checkSheetData.table.ul.column);

    // チェック欄
    var checkCell = sheet.Cells(checkSheetData.table.row, checkSheetData.table.input.column);

    // セルの挿入とか削除でずれると確認欄とかの列の幅はそのままで内容だけずれる雰囲気なので、何とかする
    // 入力欄の列の幅を保存
//    var inputCellsWidth = [];
//    {
//        // 表の一番右のセル
//        var rightCell = getLastCellInRow(sheet, cellUL.Row - 1);
//        for (var i = 0; i < rightCell.Column - cellUL.Column; i++)
//        {
//            inputCellsWidth.push(cellUL.Offset(0, 1 + i).ColumnWidth);
//        }
//    }

    var maxItemWidth = CL.getMaxItemWidth(nodeH1);
    var totalItemWidth = _.sum(maxItemWidth);

    var groupOffset = [ 0 ];
    for (var i = 1; i < maxItemWidth.length; i++)
    {
        groupOffset[i] = groupOffset[i - 1] + maxItemWidth[i - 1];
    }

    if (totalItemWidth >= 2)
    {
        // insert されたセルの Value は empty のようなので ClearContents は不要
        var insertRange = cellUL.Offset(-1, 1).Resize(2, totalItemWidth - 1);
        //var insertRangeAddress = insertRange.Address(false, false);
        insertRange.Insert(Excel.xlToRight, Excel.xlFormatFromLeftOrAbove);
        //cellUL.Parent.Range(insertRangeAddress).ClearContents();

        // ID 行も同じようにずらしておく
        insertRange = cellUL.Offset(1 - cellUL.Row, 1).Resize(1, totalItemWidth - 1);
        insertRange.Insert(Excel.xlToRight);
    }

    // 確認欄が設定されている場合
    if (nodeH1.tableHeaders)
    {
        var checkHeaderCell = cellUL.Offset(-1, totalItemWidth);
        if (nodeH1.tableHeaders.length >= 2)
        {
            checkHeaderCell.Offset(0, 1).Resize(2, nodeH1.tableHeaders.length - 1).Insert(Excel.xlToRight);
            //checkHeaderCell.Resize(2, 1).AutoFill(checkHeaderCell.Resize(2, nodeH1.tableHeaders.length), Excel.xlFillCopy);
        }
        for (var i = 0; i < nodeH1.tableHeaders.length; i++)
        {
            var tableHeader = nodeH1.tableHeaders[i];
            checkHeaderCell.Offset(0, i).Value = tableHeader.name;
            if (tableHeader.description)
            {
                var checkInputCell = checkHeaderCell.Offset(1, i);
                var validation = checkInputCell.Validation;

                validation.InputTitle = tableHeader.description;
                // XXX: InputTitle だけだと何も表示されない雰囲気なのでダミーで空白１文字を入れておく
                validation.InputMessage = " ";

                if (validation.InputTitle !== tableHeader.description) {
                    // XXX: Office2016 になったら、 Insert() した場合は、Validation.InputTitle とかに代入しても反映されない謎の状態に。回避方法見つからず
                    // XXX: 回避方法が見つかるまではExcelのコメントで
                    checkHeaderCell.Offset(0, i).AddComment(tableHeader.description);
                    checkHeaderCell.Offset(0, i).Comment.Shape.TextFrame.AutoSize = true;
                }
            }
        }
    }

    // ウィンドウ枠の固定のずれをここで修正
    // 列が固定されている場合のみ
    (function(){
        var activeWindow = excel.ActiveWindow;
        if (activeWindow.FreezePanes && activeWindow.SplitColumn > 0)
        {
            /**/
            // template のウィンドウ枠固定の設定はフラグ代わりになってしまってる感はあるけど、よしとする
            activeWindow.SplitRow = 0;
            activeWindow.SplitColumn = 0;
            activeWindow.FreezePanes = false;
            sheet.Cells(cellUL.Row, cellUL.Column + totalItemWidth).Select();
            // select だけだとダメなケースがある（原因はまったくの不明）ので、別のやり方で
            // 非表示の行・列があると SplitColumn, SplitRow はずれるらしい（多分VBAのバグ）
            // XXX: 100% template 依存の処理。時間もないので一旦はこれで
            activeWindow.SplitColumn = activeWindow.ActiveCell.Column - 1 - 1;
            activeWindow.SplitRow = activeWindow.ActiveCell.Row - 1 - 1;
            activeWindow.FreezePanes = true;
            /*/
            var splitRow = activeWindow.SplitRow;
            activeWindow.FreezePanes = false;
            activeWindow.SplitColumn = checkHeaderCell.Column - 1;
            activeWindow.SplitRow = splitRow;
            activeWindow.FreezePanes = true;
            /**/
        }
    })();

    var checkHeaders = CL.getCheckHeaders(nodeH1, checkSheetData.table);
    var checkCellsWidth = checkHeaders.length;

    // 挿入前の幅に戻す
    // 見出しセルをmergeした後にやればあんまりややこしくないのかも。やる場合は、Hなしの場合の考慮とかに気をつけること
    (function(){
        var offset = totalItemWidth - 1 + checkCellsWidth - 1;
        if (offset === 0) {
            return;
        }
        var column = checkSheetData.table.other.column + offset;
        var columnWidth = checkSheetData.table.other.columnWidth;
        for (var i = 0; i < columnWidth.length; i++)
        {
            sheet.Columns(column + i).ColumnWidth = columnWidth[i];
        }
    })();
    {
        var checkHeaderCell = cellUL.Offset(-1, totalItemWidth);
        checkHeaderCell.Resize(1, nodeH1.tableHeaders.length).Columns.EntireColumn.AutoFit();
        // 確認欄の幅は autofit してtemplateのより小さくなる場合はtemplateの幅を採用するように
        for (var i = 0; i < nodeH1.tableHeaders.length; i++) {
            if (checkHeaderCell.Offset(0, i).EntireColumn.ColumnWidth < checkSheetData.table.input.columnWidth) {
                checkHeaderCell.Offset(0, i).EntireColumn.ColumnWidth = checkSheetData.table.input.columnWidth;
            }
        }
    }

    var totalRows = getNumLeaves(nodeH1);
    if (totalRows >= 2)
    {
        // TODO: 今のtemplateには不要だけど念のため、挿入してからオートフィルする作りにしておく
        sheet.Range(cellUL.Offset(1, 0), cellUL.Offset(totalRows - 1, 0)).EntireRow.Insert(Excel.xlDown);

        // オートフィル
        //var copySrcRightCell = cellUL.Offset(0, getMaxLevel(nodeH1, kindH) - 1 + getMaxLevel(nodeH1, kindUL) - 1);
        var copySrcRightCell = getLastCellInRow(sheet, cellUL.Row - 1).Offset(1, 0);
        var copySrcRow = sheet.Range(cellUL, copySrcRightCell);
        var fillRange = sheet.Range(cellUL, copySrcRightCell.Offset(totalRows - 1, 0));
        // TODO: ちゃんとやるなら数式以外のセルを ClearContents するように。とりあえずは clear しなくても問題ないので clear なしで
        //copySrcRow.ClearContents();
        copySrcRow.AutoFill(fillRange, Excel.xlFillCopy);

        // autofillだと数式が参照している範囲に拡張が反映されないので、名前のついた範囲を拡張に合わせて更新
        // TODO: テーブル化はしない
        // TODO: 設定されてる名前を取得して範囲を拡張するように
        // XXX: とりあえず名前固定で
        sheet.Names.Add("check_cell", checkCell.resize(totalRows, checkCellsWidth));

    }

    // headerが設定してある場合は適用
    if (nodeH1.tableHeadersNonInputArea.length > 0)
    {
        (function(){
            var headers = nodeH1.tableHeadersNonInputArea;
            // グループ毎に分ける
            var headersPerGroup = [];
            for (var i = 0; i < headers.length; i++)
            {
                var group = headers[i].group;
                if (typeof headersPerGroup[group] === "undefined")
                {
                    headersPerGroup[group] = [];
                }
                headersPerGroup[group].push(headers[i]);
            }

            // header の方がグループが多い場合は削除
            if (maxItemWidth.length < headersPerGroup.length)
            {
                headersPerGroup = headersPerGroup.slice(0, maxItemWidth.length);
            }

            var actualHeaders = [];
            for (var i = 0, offset = 0; i < headersPerGroup.length; i++)
            {
                var maxWidth = maxItemWidth[i];
                var offsetEnd = offset + maxWidth;
                for (var j = 0;
                    j < headersPerGroup[i].length && offset < offsetEnd;
                    j++)
                {
                    actualHeaders.push({
                        offset: offset,
                        header: headersPerGroup[i][j]
                    });
                    offset += headersPerGroup[i][j].size;                    
                }

                offset = offsetEnd;
            }

            var headerCellOrigin = cellUL.Offset(-1, 0);
            var rangesToMerge = [];  // マージしながらだとoffsetがややこしいので、最後にまとめてマージする用

            for (var i = 0; i < actualHeaders.length; i++)
            {
                var header = actualHeaders[i].header;
                var offset = actualHeaders[i].offset;
                var cell = headerCellOrigin.Offset(0, offset);
                var size = ((i === actualHeaders.length - 1) ? totalItemWidth : actualHeaders[i + 1].offset) - offset;

                cell.Value = header.name;

                if (header.comment)
                {
                    cell.AddComment(header.comment);
                    cell.Comment.Shape.TextFrame.AutoSize = true;
                }

                if (size >= 2)
                {
                    var rangeToMerge = cell.Resize(1, size);
                    rangesToMerge.push(rangeToMerge);
                }
            }
            
            for (var i = 0; i < rangesToMerge.length; i++)
            {
                var rangeToMerge = rangesToMerge[i];
                rangeToMerge.HorizontalAlignment = Excel.xlCenter;
                rangeToMerge.Merge();
            }
        })();
    }
    else if (totalItemWidth >= 2)
    {
        var cellLeft = cellUL.Offset(-1, 0);
        var cellRight = cellLeft.Offset(0, totalItemWidth - 1);
        sheet.Range(cellLeft.Offset(0, 1), cellRight).ClearContents();
        var cellToMerge = sheet.Range(cellLeft, cellRight);
        cellToMerge.HorizontalAlignment = Excel.xlCenter;
        cellToMerge.Merge();
    }

    function new2dArray(n1, n2)
    {
        var array = [];
        for (var i = 0; i < n1; i++)
        {
            // new Array() で作って一度も代入してないと safe array 変換でバグる
            // Array.prototype.push.apply() で新しい配列に入れなおすだけで正常動作するっぽいけど、最初から null 埋めしておく
            array.push(_.fill(Array(n2), null));
        }
        return array;
    }

    // 2d 配列を転置したものを返す
    function transposed2dArray(array)
    {
        var n1 = array[0].length;
        var n2 = array.length;
        var result = new2dArray(n1, n2);
        for (var i = 0; i < n1; i++)
        {
            for (var j = 0; j < n2; j++)
            {
                result[i][j] = array[j][i];
            }
        }
        return result;
    }

    {(function(){
        var textArray = new2dArray(totalRows, totalItemWidth);
        var mergeCellMap = new2dArray(totalRows, totalItemWidth + 1);   // 番兵用に1列多めに確保
        var imagePath = "images";

        if (nodeH1.variables.imagePath) {
            imagePath += "/" + nodeH1.variables.imagePath;
        }

        renderUL_Recurse(nodeH1, 0, cellUL, totalItemWidth, groupOffset, imagePath, textArray, mergeCellMap);

        //var startTime = performance.now();
        //var result = [];
        //for (var t = 0; t < 1000; t++) {
        //    result.push(jsArray2dToSafeArray2d(textArray));
        //}
        //var endTime = performance.now();
        //WScript.Echo("2d:\n" + (endTime - startTime));

        var startTime = performance.now();
        //var maxColumns = 0;
        //textArray.forEach(function(e) { maxColumns = Math.max(maxColumns, e.length); });
        cellUL.Resize(totalRows, totalItemWidth).Value = jsArray2dToSafeArray2d(textArray);
        var endTime = performance.now();
        //WScript.Echo("2d:\n" + (endTime - startTime));
    
        // 初期値が設定されているセルは入力
        renderInitialValues_Recurse(nodeH1, cellUL.Offset(0, totalItemWidth), 0);

        // マージ後だとoffsetがまともに扱えないので、マージ前のrangeを保持
        var rangeToAutoFitColumns = cellUL.Offset(-1, 0).Resize(totalRows + 1, totalItemWidth);

        // 入力があった列を書き出し & autofit
        // マージ前に保存
        var checkCellOrigin = cellUL.Offset(0, totalItemWidth);

        // コメント画像のサイズが AutoFit で崩れるようになったので、対処
        var pictureRects = [];
        for (var i = 0; i < sheet.Comments.Count; i++) {
            var commentShape = sheet.Comments(1+i).Shape;
            pictureRects.push({ width: commentShape.Width, height: commentShape.Height });
        }

        betterAutoFit(cellUL, mergeCellMap);

        MergeULCells(cellUL, mergeCellMap);

        // autofitはセルをマージした後にやる
        //rangeToAutoFitColumns.Columns.AutoFit();
        // FIXME: H2が存在しない場合にデフォの確認欄がautofitされてるっぽい
        cellUL.Resize(totalRows, totalItemWidth).Rows.AutoFit();

        for (var i = 0; i < sheet.Comments.Count; i++) {
            var commentShape = sheet.Comments(1+i).Shape;
            var rect = pictureRects[i];
            commentShape.Width = rect.width;
            commentShape.Height = rect.height;
        }

    })();}
}

// 扱いやすい形に変換
function mergeCellMapToWidthMap(mergeCellMap) {
    var result = [];
    var height = mergeCellMap.length;
    if (height === 0) {
        return;
    }
    var maxWidth = 0;
    for (var y = 0; y < height; y++) {
        var buf = [];
        var width = mergeCellMap[y].length;
        var id0 = mergeCellMap[y][x0];
        var x0 = 0;
        var id0 = mergeCellMap[y][x0];
        buf[x0] = 1;
        maxWidth = Math.max(maxWidth, width);
        for (var x = 1; x < width; x++) {
            var id = mergeCellMap[y][x];
            if (id === id0) {
                buf[x0]++;
                buf[x] = 0;
            }
            else {
                id0 = id;
                x0 = x;
                buf[x0] = 1;
            }
        }
        //buf.pop();  // 番兵を削除
        result.push(buf);
    }
    // 番兵を追加
    result.push(_.fill(Array(maxWidth), 0));
    
    return result;
}

function betterAutoFit(cellOrigin, mergeCellMap) {
    var widthMap = mergeCellMapToWidthMap(mergeCellMap);
    var height = widthMap.length;
    if (height === 0) {
        return;
    }
    var width = 0;
    for (var y = 0; y < height; y++) {
        width = Math.max(width, widthMap[y].length);
    }
    // 番兵の列は除外
    width--;

    //for (var x = 0; x < width; x++) {
    //    var rowList = [];
    //    for (var y = 0; y < height; y++) {
    //        var w = widthMap[y][x];
    //        if (_.isUndefined(w) || w === 0) {
    //            continue;
    //        }
    //        if (_.isUndefined(rowList[w - 1])) {
    //            rowList[w - 1] = [];
    //        }
    //        rowList[w - 1].push(y);
    //    }
    //    // TODO: 比率が高い幅を優先もしくは重視した方が良い感じになるかも
    //    ;;;
    //}

    // 折返しが有効だと幅を広げる方向には AutoFit されないようなので元の大きさに広げておく
    cellOrigin.Resize(1, width).EntireColumn.ColumnWidth = templateData.checkSheet.table.ul.columnWidth;

    var columnWidth = [];
    for (var x = 0; x < width; x++) {
        var range = null;
        for (var y = 0; y < height; y++) {
            var w = widthMap[y][x];
            if (w !== 1) {
                continue;
            }

            var y0 = y;
            y++;
            // 番兵がいるので y < height 不要
            for (;; y++) {
                if (widthMap[y][x] === 1) {
                    continue;
                }
                var subRange = cellOrigin.Offset(y0, x);
                var h = y - y0;
                // 連続してるなら一塊にして扱う
                if (h >= 2) {
                    subRange = subRange.Resize(h, 1);
                }
                range = (range !== null) ? excel.Union(range, subRange) : subRange;
                break;
            }
        }
        if (range !== null) {
            range.Columns.AutoFit();
            columnWidth.push(range.Columns.ColumnWidth);
        }
    }
}

//function betterAutoFit(cellOrigin, mergeCellMap) {
//    var height = mergeCellMap.length;
//    if (height === 0) {
//        return;
//    }
//    var width = mergeCellMap[0].length;
//
//    // まずはマージされないセル基準で幅をautofit
//    var autoFitColumnCells = [];
//    for (var i = 0; i < width - 1; i++) {
//        autoFitColumnCells[i] = [];
//    }
//    for (var y = 0; y < height; y++) {
//        var count = 0;
//        var id0 = mergeCellMap[y][0];
//        for (var x = 1; x < width; x++) {
//            var id = mergeCellMap[y][x];
//            if (id0 === id) {
//                count++;
//                continue;
//            }
//            if (count === 0) {
//                var cell = cellOrigin.Offset(y, x - 1);
//                autoFitColumnCells[x - 1].push(cell);
//            }
//            count = 0;
//            id0 = id;
//        }
//    }
//    for (var i = 0; i < autoFitColumnCells.length; i++) {
//        var cellsList = autoFitColumnCells[i];
//        var range = cellsList[0];
//        for (j = 1; j < cellsList.length; j++) {
//            range = excel.Union(range, cellsList[j]);
//        }
//        range.Columns.AutoFit();
//    }
//    ;;;
//}

// excel の Range.Value.toArray() で取得した配列を a[row(y)][column(x)] な配列に変換
// 処理的にはどうってことないはずなので扱いやすい形に変換してしまう
function rangeValueToArray2d(range)
{
    var rows = range.Rows.Count;
    var array = range.Value.toArray();
    var a = new Array(rows);

    for (var y = 0; y < rows; y++)
    {
        a[y] = [];
    }
    for (var i = 0; i < array.length; )
    {
        for (var y = 0; y < rows; y++)
        {
            a[y].push(array[i++]);
        }
    }

    return a;
}

// mergeCellMap で同じIDが横方向に連続しているセルをマージする
function MergeULCells(cellOrigin, mergeCellMap)
{
    var height = mergeCellMap.length;
    var cellsToMerge = [];  // マージしながらだと offset がややこしくなるので、一旦保存してからマージする

    for (var y = 0; y < height; y++)
    {
        var width = mergeCellMap[y].length;
        var x0 = 0;
        var id = mergeCellMap[y][x0];
        for (var x = 1; x < width; x++)
        {
            if (mergeCellMap[y][x] !== id)
            {
                var mergeWidth = x - x0;
                if (mergeWidth >= 2)
                {
                    cellsToMerge.push(cellOrigin.Offset(y, x0).Resize(1, mergeWidth));
                }
                x0 = x;
                id = mergeCellMap[y][x0];
            }
        }
    }

    for (var i = 0; i < cellsToMerge.length; i++)
    {
        cellsToMerge[i].Merge(true);
    }
}

function DeleteTemplateRowFromIndexSheet(indexSheet, templateSheet)
{
    var templateName = templateSheet.Name;

    var cell = indexSheet.Cells.Find(templateName, indexSheet.Cells(1, 1), Excel.xlValues, Excel.xlWhole, Excel.xlByRows, Excel.xlNext, true);

    indexSheet.Rows(cell.Row).Delete();
}

// OBSOLETE
/*
function AddToIndexSheet(indexSheet, templateSheet, srcSheet)
{
    var templateName = templateSheet.Name;

    var srcCell = indexSheet.Cells.Find(templateName, indexSheet.Cells(1, 1), Excel.xlValues, Excel.xlWhole, Excel.xlByRows, Excel.xlNext, true);
    if (!srcCell)
    {
        Error("シートにテンプレート名 " + templateName + " のセルが存在しません");
    }

    var srcRow = srcCell.Row;
    var rowRange = indexSheet.Rows(srcRow);
    rowRange.Copy();
    rowRange.Offset(1).Insert(Excel.xlDown);

    var firstCell = getFirstCellInRow(indexSheet, srcRow);
    var lastCell = getLastCellInRow(indexSheet, srcRow);
    var replaceRange = indexSheet.Range(firstCell, lastCell);

    var regExp = new RegExp(templateName, "g");

    var afterName = srcSheet.Name;
    var afterNameSQ = "'" + afterName + "'";

    // TODO: リンクの置換
    // TODO: 合計セルに表の拡張が反映されるように（一旦テーブルにして挿入したらテーブル解除で）
    // TODO: シートの順に index に追加されるように
    xEach(replaceRange, function(cell)
    {
        if (cell.HasFormula)
        {
            var formula0 = cell.Formula;
            cell.Formula = formula0.replace(regExp, afterNameSQ);
        }
        else
        {
            var value0 = cell.Value;
            cell.Value = value0.replace(regExp, afterName);

            // TODO: リンクの置換
            // FIXME: 数式じゃないセルが必ずリンクがある前提の作りになってる
            //var subAddress0 = cell.Hyperlinks(1).SubAddress;
            //cell.Hyperlinks(1).SubAddress = subAddress0.replace(regExp, afterNameSQ);
        }
    });

    excel.CutCopyMode = false;

}
*/

// 途中で SaveAs してしまうと、エラーが起きた場合、中途半端な状態の dstFilePath な book のファイルを削除する必要が出てきてしまう
// 最後に SaveAs するようにして（readonly でも編集はできるはずなので）、エラー時は保存せずにそのまま強制的に閉じればシンプルになる
// 元の ActiveWorkBook は自動で閉じられ、そのまま book が SaveAs した book になる雰囲気
book.SaveAs(dstFilePath);

excel.StatusBar = false;
excel.ScreenUpdating = true;
