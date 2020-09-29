// XXX: 高速化のため、現状、 formula とか Date には対応してない

// InputRange とかにするべきだったか
function CheckRange_old(nodeH1, sheet)
{
    var leftHeaderCell = sheet.Range(nodeH1.leftCheckHeaderCellAddress);
    var rightHeaderCell = getLastCellInRow(sheet, leftHeaderCell.Row);
    var headerRange = sheet.Range(leftHeaderCell, rightHeaderCell);
    var numHeaders = headerRange.Columns.Count;
    var leafNodes = CL.getLeafNodes(nodeH1);
    var checkCellRows = leafNodes.length;
    var checkRange = headerRange.Offset(1, 0).Resize(checkCellRows, numHeaders);
    var headers = headerRange.Value.toArray();

    this.checkHeaders = nodeH1.tableHeaders;

    var hasFormulaColumns = [];
    xEach(headerRange.Offset(1, 0), function(cell)
    {
        // 列が数式かどうか
        hasFormulaColumns.push(cell.HasFormula);
    });
    // bool を反転
    var isTarget = hasFormulaColumns.map(function(f) {
        return !f;
    });

    this.checkRange = checkRange;
    this.leafNodes = leafNodes;
    this.headers = headers;
    this.isTarget = isTarget;

    this.width = headers.length;
    this.height = leafNodes.length;

    this.nodeH1 = nodeH1;
    this.sheet = sheet;

    this.checkCellArray = null;

    this.idToY = null;
    this.headerToX = null;

    this.sheetValues = null;
}
function CheckRange(nodeH1, sheet, templateData)
{
    var table = templateData.checkSheet.table;

    var maxItemWidth = CL.getMaxItemWidth(nodeH1);
    var totalItemWidth = _.sum(maxItemWidth);
    var checkHeaders = CL.getCheckHeaders(nodeH1, table);
    var checkCellsWidth = checkHeaders.length;

    var headers = checkHeaders.concat(table.other.headers);

    var leafNodes = CL.getLeafNodes(nodeH1);

    // save 対象が含まれてさえいれば良いので
    var otherCellsWidth = (table.other.indicesToSave.length === 0) ? 0 : _.max(table.other.indicesToSave) + 1;
    var column = table.ul.column + totalItemWidth;
    var width = checkCellsWidth + otherCellsWidth;
    var height = leafNodes.length;
    var checkRange = sheet.Cells(table.row, column).Resize(height, width);

    // check cell の左からの index
    var indicesToSave = _.range(0, checkCellsWidth);
    table.other.indicesToSave.forEach(function(element, index, array) {
        indicesToSave.push(checkCellsWidth + element);
    });

    var isTarget = [];
    for (var i = 0; i < indicesToSave.length; i++) {
        isTarget[indicesToSave[i]] = true;
    }
    
    this.checkHeaders = nodeH1.tableHeaders;

    this.checkRange = checkRange;
    this.leafNodes = leafNodes;
    this.headers = headers;
    this.isTarget = isTarget;

    this.width = width;
    this.height = height;

    this.nodeH1 = nodeH1;
    this.sheet = sheet;

    this.checkCellArray = null;

    this.idToY = null;
    this.headerToX = null;

    this.sheetValues = null;

    this.templateData = templateData;
}

// シートに値をセット
CheckRange.prototype.setArray2d = function(srcArray) {
    // 2d配列を部分的に取り出した2d配列
    // 指定した矩形がsrcをはみ出した場合のチェックとかはしない
    function getSubArray2d(src, offsetX, offsetY, width, height)
    {
        var dst = new Array(height);

        for (var y = 0, srcY = offsetY; y < height; y++, srcY++)
        {
            dst[y] = src[srcY].slice(offsetX, offsetX + width);
        }

        return dst;
    }

    var isTarget = this.isTarget.concat(false);   // 番兵

    var offsetWidth = [];
    for (var i = 0, offset = 0; i < isTarget.length; i++)
    {
        if (!isTarget[i])
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

    var dstRange = this.checkRange;
    var height = this.height;
    for (var i = 0; i < offsetWidth.length; i++)
    {
        var ow = offsetWidth[i];
        var offsetX = ow.offset;
        var width = ow.width;
        var range = dstRange.Offset(0, offsetX).Resize(height, width);
        var subArray = getSubArray2d(srcArray, offsetX, 0, width, height);
        range.Value = jsArray2dToSafeArray2d(subArray);
    }

    // 入力がある確認欄の列は autofit
    // としようと思ったけど、別に「確認欄の列は無条件で autofit」で良いか
    {
        var width = this.checkHeaders.length;
        var rangeToAutofit = dstRange.Offset(-1, 0).Resize(height + 1, width);
        rangeToAutofit.Columns.AutoFit();

        if (this.templateData) {
            var defaultInputColumnWidth = this.templateData.checkSheet.table.input.columnWidth;
            for (var i = 0; i < width; i++) {
                if (rangeToAutofit.Columns(1 + i).ColumnWidth < defaultInputColumnWidth) {
                    rangeToAutofit.Columns(1 + i).ColumnWidth = defaultInputColumnWidth;
                }
            }
        }
    }

    // これは更新が必要
    this.checkCellArray = null;
    this.sheetValues = null;
};

CheckRange.prototype.getCheckCellArray2d = function() {
    if (!this.checkCellArray)
    {
        this.checkCellArray = CL.RangeToValueArray2d(this.checkRange);
    }
    return this.checkCellArray;
};

CheckRange.prototype.getIdToY = function() {
    if (!this.idToY)
    {
        this.idToY = [];

        for (var i = 0; i < this.height; i++)
        {
            this.idToY[this.leafNodes[i].id] = i;
        }
    }
    return this.idToY;
};

CheckRange.prototype.createEmptyArray2d = function() {
    var array = [];
    //array[0] = new Array(this.width);   // 1行目のサイズで２次元配列の大きさを判断することが多いので、1行目だけサイズ分確保
    //for (var y = 1; y < this.height; y++)
    //{
    //    array[y] = [];
    //}
    for (var y = 0; y < this.height; y++) {
        // new Array() で作って一度も代入してないと safe array 変換でバグる
        // Array.prototype.push.apply() で新しい配列に入れなおすだけで正常動作するっぽいけど、最初から null 埋めしておく
        array.push(_.fill(Array(this.width), null));
    }

    return array;
};

CheckRange.prototype.getHeaderToX = function() {
    if (!this.headerToX)
    {
        this.headerToX = [];
        for (var i = 0; i < this.width; i++)
        {
            this.headerToX[this.headers[i]] = i;
        }
    }
    return this.headerToX;
};

CheckRange.prototype.getSheetValues = function() {
    if (!this.sheetValues)
    {
        var nodeH1 = this.nodeH1;
        var checkCellArray = this.getCheckCellArray2d();
        var sheetValues = {
            text: nodeH1.text,
            //headers: nodeH1.tableHeaders,
            items: {}
        };
        for (var y = 0; y < this.height; y++)
        {
            var values = {};

            for (var x = 0; x < this.width; x++)
            {
                if (!this.isTarget[x])
                {
                    continue;
                }
                var v = checkCellArray[y][x];
                //v = (typeof v === "undefined") ? null : v;
                values[this.headers[x]] = v;
            }

            // 空っぽなら追加しない
            if (Object.keys(values).length === 0)
            {
                continue;
            }

            var node = this.leafNodes[y];
            var item = {
                text: node.text,
                values: values
            };
            sheetValues.items[node.id] = item;
        }

        this.sheetValues = sheetValues;
    }
    return this.sheetValues;
};

CheckRange.prototype.getArray2dFromSheetValues = function(items) {
    var array = this.createEmptyArray2d();
    var idToY = this.getIdToY();
    var headerToX = this.getHeaderToX();
    for (var id in items)
    {
        var item = items[id];
        var y = idToY[id];
        var values = items[id].values;
        for (var header in values)
        {
            var x = headerToX[header];
            array[y][x] = values[header];
        }
    }

    return array;
};

// sheetValues の状態(from)からの今のシートの状態(to)の差分を求める
CheckRange.prototype.getChangeFromSheetValues = function(items) {
    var fromArray = this.getArray2dFromSheetValues(items);
    var toArray = this.getCheckCellArray2d();

    var nodeH1 = this.nodeH1;
    var sheetChange = {
        text: nodeH1.text,
        items: {}
    };

    for (var y = 0; y < this.height; y++)
    {
        var values = {};

        for (var x = 0; x < this.width; x++)
        {
            if (!this.isTarget[x])
            {
                continue;
            }
            // JSON stringify で消されないように、とりあえず null を入れておく
            var v = fromArray[y][x];
            if (v !== toArray[y][x])
            {
                var v = (typeof v === "undefined") ? null : v;
                
                values[this.headers[x]] = v;
            }
        }

        // 空っぽなら追加しない
        if (Object.keys(values).length === 0)
        {
            continue;
        }

        var node = this.leafNodes[y];
        var item = {
            text: node.text,
            values: values
        };
        sheetChange.items[node.id] = item;
    }

    if (Object.keys(sheetChange.items).length === 0)
    {
        return null;
    }

    return sheetChange;
};
