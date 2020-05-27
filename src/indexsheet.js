// XXX: 高速化のため、現状、 formula とか Date には対応してない

function replaceEmptyToUndefined(v) {
    return (v === null || v === "") ? undefined : v;
};

function IndexSheetVariables(root, sheet, templateData) {
    var variables = {};

    // 変数名が _ で始まり、その次が大文字の変数は値を取り込む
    for (var key in root.variables) {
        if (!/^_[A-Z].*/.test(key)) {
            continue;
        }

        var variableData = templateData.indexSheet.variables[key];
        if (_.isUndefined(variableData)) {
            continue;
        }

        var address = variableData.address;
        var cell = indexSheet.Range(address);

        variables[key] = {
            cell: cell
        };
    }

    this.variables = variables;
};

// 今のシートの状態を取得
// emptyValue: 空セルの value は undefined として取得されるが、 json にした際に消えてしまうのが都合が悪い場合に使うことを想定
IndexSheetVariables.prototype.getFromSheet = function(emptyValue) {
    var values = {};

    for (var key in this.variables) {
        var variable = this.variables[key];

        // XXX: 日付を取得するために Text にしてお茶を濁す。本当はDateをDateとして扱うべき
        var v = variable.cell.Text;

        values[key] = (_.isUndefined(v)) ? emptyValue : v;
    }

    return values;
};

// シートにセット
IndexSheetVariables.prototype.setToSheet = function(values) {
    for (var key in values) {
        var dst = this.variables[key];
        if (_.isUndefined(dst)) {
            // TODO: notFound
            continue;
        }

        dst.cell.Value = replaceEmptyToUndefined(values[key]);
    }
};

// conflict なら null を返す
// null 渡すの禁止
// TODO: 共通で使える場所に移動
function getMerged(base, theirs, mine) {
    if (theirs === mine || base === mine) {
        return theirs;
    }
    if (theirs === base) {
        return mine;
    }
    return null;
}

// merge されたものを返す
// conflict 時は theirs 採用という仕様にしておく
IndexSheetVariables.prototype.getMerged = function(base, theirs, mine, conflicts) {
    var merged = {};

    var keys = _.union([_.keys(base), _.keys(theirs), _.keys(mine)]);

    for (var key in keys) {
        var b = replaceEmptyToUndefined(base[key]);
        var t = replaceEmptyToUndefined(theirs[key]);
        var m = replaceEmptyToUndefined(mine[key]);
        merged[key] = getMerged(b, t, m);

        // conflict
        if (merged[key] === null) {
            conflicts[key] = {
                base: b,
                theirs: t,
                mine: m
            };
            merged[key] = t;
        }
    }

    return merged;
};

// lhs の状態からの rhs の状態の差分を求める
// 差異がある箇所の rhs 側の値のみ返す
// 要素が一致しているもののみ対応。一旦エラーチェックなし
// 空セルが JSON stringify で消されないように、 undefined の場合の値を指定できるように
// 変更を適用するときに空セルへの変更が存在しない（undefined）と空セルへの変更ができなくなるので
// メンバー関数である必要はないけど、管理上ここに入れておきたい
IndexSheetVariables.prototype.getChanges = function(lhs, rhs, emptyValue) {
    var changes = {};

    var keys = _.union([_.keys(lhs), _.keys(rhs)]);
    for (var key in keys) {
        var r = replaceEmptyToUndefined(rhs[key]);
        var l = replaceEmptyToUndefined(lhs[key]);

        if (r !== l) {
            changes[key] = (_.isUndefined(r)) ? emptyValue : r;
        }
    }

    return changes;
};

// 破壊的（variables を直接変更）
IndexSheetVariables.prototype.applyChanges = function(values, changes) {
    for (var key in changes) {
        values[key] = replaceEmptyToUndefined(changes[key]);
    }
};



function IndexSheetTable(root, sheet, templateData) {
    this.root = root;
    this.templateData = templateData.indexSheet.table;

    var maxIndex = _.max(this.templateData.indicesToSave);

    this.height = root.children.length;
    this.width = maxIndex + 1;

    this.range = sheet.Cells(this.templateData.row, this.templateData.column).Resize(this.height, this.width);

    this.cellArray2d = null;

    // XXX: 最左列は番号という前提で
    this.h1Indices = CL.array2dTransposed(this.getCellArray2d())[0];
};

IndexSheetTable.prototype.getCellArray2d = function() {
    if (!this.cellArray2d) {
        this.cellArray2d = CL.rangeToValueArray2d(this.range);
    }
    return this.cellArray2d;
}

// 今のシートの状態を取得
IndexSheetTable.prototype.getFromSheet = function(emptyValue) {
    var cellArray2d = this.getCellArray2d();
    var root = this.root;
    var table = this.templateData;

    var items = {};

    for (var i = 0; i < root.children.length; i++) {
        var nodeH1 = root.children[i];
        var sheetValues = {
            text: nodeH1.text,
            values: {}
        };

        // Array.prototype.findIndex() は使えないようなので、 lodash のを使う
        var index = _.findIndex(root.children, function(element) {
            return (element.id === nodeH1.id);
        });
        var y = this.h1Indices.indexOf(index + 1);    // Excel側は 1 origin なので +1
        var arrayX = cellArray2d[y];

        table.indicesToSave.forEach(function(x, index, array) {
            var v = arrayX[x];

            sheetValues.values[table.headers[x]] = _.isUndefined(v) ? emptyValue : v;
        });

        items[nodeH1.id] = sheetValues;
    }

    return items;
};

IndexSheetTable.prototype.setToSheet = function(values) {
    var cellArray2d = this.getCellArray2d();
    var root = this.root;
    var table = this.templateData;

    var idToNodeH1 = {};
    root.children.forEach(function (nodeH1) {
        idToNodeH1[nodeH1.id] = nodeH1;
    });

    var items = values.items;
    for (var id in items) {
        var item = items[id];
        var nodeH1 = idToNodeH1[id];

        if (_.isUndefined(nodeH1)) {
            // TODO: values がすべて undefined じゃなければ notfound へ（そもそも src がないなら、失われる情報はないので警告の意味がない）
            continue;
        }

        // Array.prototype.findIndex() は使えないようなので、 lodash のを使う
        var index = _.findIndex(root.children, function(element) {
            return (element.id === nodeH1.id);
        });
        var y = this.h1Indices.indexOf(index + 1);    // Excel側は 1 origin なので +1
        var arrayX = cellArray2d[y];

        var values = item.values;
        for (header in values) {
            var x = table.headers.indexOf(header);
            var value = replaceEmptyToUndefined(values[header]);
            if (x === -1) {
                if (!_.isUndefined(value)) {
                    // TODO: notFound へ
                }
                continue;
            }

            arrayX[x] = value;
        }
    }

    applyArrayToRange(this.range, cellArray2d, getOffsetWidthFromIndexArray(table.indicesToSave));

    // これは更新が必要
    this.cellArray2d = null;
};

// save対象のbool配列から offset width 配列を求める
function getOffsetWidthFromBoolArray(array) {
    var offsetWidth = [];
    var array = array.concat(false);   // 番兵

    for (var i = 0, offset = 0; i < array.length; i++) {
        if (!array[i]) {
            var width = i - offset;
            if (width > 0) {
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

// indicesToSave を渡す想定
// 対象範囲の左端を基準とした昇順な index 配列
function getOffsetWidthFromIndexArray(indexArray) {
    var offsetWidth = [];

    for (var i = 0; i < indexArray.length; ) {
        var i0 = i;
        i++;
        // 値が連続してる間 i を進める
        for (; i < indexArray.length; i++) {
            if (indexArray[i] - indexArray[i - 1] > 1) {
                break;
            }
        }
        offsetWidth.push({
            offset: indexArray[i0],
            width: i - i0
        });
    }

    return offsetWidth;
}

// 2d配列を部分的に取り出した2d配列
// 指定した矩形がsrcをはみ出した場合のチェックとかはしない
function getRectArray2d(src, offsetX, offsetY, width, height) {
    var dst = new Array(height);

    for (var y = 0, srcY = offsetY; y < height; y++, srcY++) {
        dst[y] = src[srcY].slice(offsetX, offsetX + width);
    }

    return dst;
}

// 指定された列だけ分割しながらできるだけまとめて入力
function applyArrayToRange(dstRange, srcArray, offsetWidth) {
    var height = srcArray.length;
    for (var i = 0; i < offsetWidth.length; i++) {
        var ow = offsetWidth[i];
        var offsetX = ow.offset;
        var width = ow.width;
        range = dstRange.Offset(0, offsetX).Resize(height, width);
        array = getRectArray2d(srcArray, offsetX, 0, width, height);
        range.Value = jsArray2dToSafeArray2d(array);
    }
}

function getValuesFromItem(item) {
    return _.isUndefined(item) ? {} : item.values;
}

// base, theirs, mine には同じ構成のデータを渡す
// conflict 時は theirs 採用という仕様にしておく
IndexSheetTable.prototype.getMerged = function(base, theirs, mine, conflicts) {
    var merged = {};
    var sheetIds = _.union([_.keys(base), _.keys(theirs), _.keys(mine)]);

    sheetIds.forEach(function(sheetId) {
        var baseValues = getValuesFromItem(base[sheetId]);
        var theirsValues = getValuesFromItem(theirs[sheetId]);
        var mineValues = getValuesFromItem(mine[sheetId]);
        // _.keys() は undefined を渡すと空の object を返してくれる
        var headers = _.union([_.keys(baseValues), _.keys(theirsValues), _.keys(mineValues)]);

        if (_.isEmpty(headers)) {
            return;
        }

        var mergedValues = {};
        var conflictValues = {};
        headers.forEach(function(header) {
            var b = replaceEmptyToUndefined(baseValues[header]);
            var t = replaceEmptyToUndefined(theirsValues[header]);
            var m = replaceEmptyToUndefined(mineValues[header]);

            mergedValues[header] = getMerged(b, t, m);

            // conflict
            if (mergedValues[header] === null) {
                conflictValues[header] = {
                    base: b,
                    theirs: t,
                    mine: m
                };
                mergedValues[header] = t;
            }
        });

        if (!_.empty(conflictValues)) {
            conflicts[sheetId] = {
                text: theirs[sheetId].text, // FIXME: theirs が存在しない場合は他のから取得
                values: conflictValues
            };
        }

        if (_.every(mergedValues, _.negate(_.isUndefined))) {
            merged[sheetId] = {
                text: theirs[sheetId].text, // FIXME: theirs が存在しない場合は他のから取得
                values: mergedValues
            };
        }

    });

    return merged;
};

IndexSheetTable.prototype.getChanges = function(lhs, rhs, emptyValue) {
    var changes = {};

    var sheetIds = _.union([_.keys(lhs), _.keys(rhs)]);
    sheetIds.forEach(function(sheetId) {
        var lValues = getValuesFromItem(lhs[sheetId]);
        var rValues = getValuesFromItem(rhs[sheetId]);
        var headers = _.union([_.keys(lValues), _.keys(rValues)]);
        var values = {};
        headers.forEach(function(header) {
            var l = replaceEmptyToUndefined(lValues[header]);
            var r = replaceEmptyToUndefined(rValues[header]);

            if (r !== l) {
                values[header] = (_.isUndefined(r)) ? emptyValue : r;
            }
        });

        if (!_.empty(values)) {
            changes[sheetId] = {
                text: rhs[sheetId].text,    // FIXME: rhs[id]がundefinedならlhs[id]に
                values: values
            };
        }
    });

    return changes;
};

// 破壊的（直接変更）
IndexSheetTable.prototype.applyChanges = function(items, changes) {
    _.forIn(changes, function(changesInSheet, sheetId) {
        if (_.isUndefined(items[sheetId])) {
            items[sheetId] = {
                text: changesInSheet.text,
                values: {}
            };
        }
        var dstValues = items[sheetId].values;
        var srcValues = changesInSheet.values;
        _.forIn(srcValues, function(value, header) {
            dstValues[header] = replaceEmptyToUndefined(value);
        });
    });
};



function IndexSheet(root, sheet, templateData)
{
    this.variables = new IndexSheetVariables(root, sheet, templateData);
    this.table = new IndexSheetTable(root, sheet, templateData);


//    var maxItemWidth = CL.getMaxItemWidth(nodeH1);
//    var totalItemWidth = _.sum(maxItemWidth);
//    var checkHeaders = CL.getCheckHeaders(nodeH1, table);
//    var checkCellsWidth = checkHeaders.length;
//
//    var headers = checkHeaders.concat(table.other.headers);
//
//    var leafNodes = CL.getLeafNodes(nodeH1);
//
//    // save 対象が含まれてさえいれば良いので
//    var otherCellsWidth = (table.other.indicesToSave.length === 0) ? 0 : _.max(table.other.indicesToSave) + 1;
//    var column = table.ul.column + totalItemWidth;
//    var width = checkCellsWidth + otherCellsWidth;
//    var height = leafNodes.length;
//    var checkRange = sheet.Cells(table.row, column).Resize(height, width);
//
//    // check cell の左からの index
//    var indicesToSave = _.range(0, checkCellsWidth);
//    table.other.indicesToSave.forEach(function(element, index, array) {
//        indicesToSave.push(checkCellsWidth + element);
//    });
//
//    var isTarget = [];
//    for (var i = 0; i < indicesToSave.length; i++) {
//        isTarget[indicesToSave[i]] = true;
//    }
//    
//    this.checkHeaders = nodeH1.tableHeaders;
//
//    this.checkRange = checkRange;
//    this.leafNodes = leafNodes;
//    this.headers = headers;
//    this.isTarget = isTarget;
//
//    this.width = width;
//    this.height = height;
//
//    this.nodeH1 = nodeH1;
//    this.sheet = sheet;
//
//    this.checkCellArray = null;
//
//    this.idToY = null;
//    this.headerToX = null;
//
//    this.sheetValues = null;
//
//    this.templateData = templateData;
}

IndexSheet.prototype.getFromSheet = function(undefinedValue) {
    return {
        variables: this.variables.getFromSheet(undefinedValue),
        items: this.table.getFromSheet(undefinedValue)
    };
};

IndexSheet.prototype.setToSheet = function(values, isUndefined) {
    this.variables.setToSheet(values.variables, isUndefined);
    this.table.setToSheet(values.items, isUndefined);
};

IndexSheet.prototype.getMerged = function(base, theirs, mine, conflicts, isUndefined) {
    conflicts.variables = {};
    conflicts.items = {};

    return {
        variables: this.variables.getMerged(base.variables, theirs.variables, mine.variables, conflicts.variables, isUndefined),
        items: this.table.getMerged(base.items, theirs.items, mine.items, conflicts.items, isUndefined)
    }
};

IndexSheet.prototype.getChanges = function(lhs, rhs, emptyValue) {
    return {
        variables: this.variables.getChanges(lhs.variables, rhs.variables, emptyValue),
        items: this.table.getChanges(lhs.items, rhs.items, emptyValue)
    };
};

IndexSheet.prototype.applyChanges = function(values, changes) {
    this.variables.applyChanges(values.variables, changes.variables);
    this.table.applyChanges(values.items, changes.items);
};



/*
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
    array[0] = new Array(this.width);   // 1行目のサイズで２次元配列の大きさを判断することが多いので、1行目だけサイズ分確保
    for (var y = 1; y < this.height; y++)
    {
        array[y] = [];
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
*/
