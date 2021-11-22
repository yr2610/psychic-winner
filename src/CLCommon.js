CL = {};

CL.kind = {
  H: "H",
  UL: "UL"
};

CL.readTextFileUTF8 = function (filePath) {
  var stream = new ActiveXObject("ADODB.Stream");

  stream.Type = adTypeText;
  stream.charset = "utf-8";
  stream.Open();
  stream.LoadFromFile(filePath);
  var s = stream.ReadText(adReadAll);
  stream.Close();

  return s;
};

// json をテキストファイルに書き出すのを作ったけど、 stringify は別にやれば済む話なので単に文字列をテキストファイル化するだけな感じで
CL.writeTextFileUTF8 = function (s, outFilePath) {
  var stream = new ActiveXObject("ADODB.Stream");

  stream.Type = adTypeText;
  stream.charset = "utf-8";
  stream.Open();

  s = s.split("\n").join("\r\n"); // メモ帳で開けるように…

  stream.WriteText(s, adWriteLine);

  stream.Position = 0;
  stream.Type = adTypeBinary;
  stream.Position = 3;
  var bytes = stream.Read();
  stream.Position = 0;
  stream.SetEOS();
  stream.Write(bytes);

  stream.SaveToFile(outFilePath, adSaveCreateOverWrite);
  stream.Close();
  stream = null;
};

// 圧縮されてれば展開
CL.decompressJSON = function (json) {
  var o = JSON.parse(json);

  var count = 0;
  while (o.compress) {
    if (o.compress === "LZString") {
      var decompressor;
      if (o.option === "UTF16") {
        decompressor = LZString.decompressFromUTF16;
      }
      else if (o.option === "EncodedURIComponent") {
        decompressor = LZString.decompressFromEncodedURIComponent;
      }
      else {
        decompressor = LZString.decompress;
      }
      json = decompressor(o.data);
      o = JSON.parse(json);
    }
    else {
      Error("invalid compressor.");
      return;
    }
    if (count++ > 10) {
      Error("invalid JSON data.");
      return;
    }
  }

  return {
    json: json,
    object: o
  };
};

CL.readJSONFile = function (jsonFilePath) {
  var s = CL.readTextFileUTF8(jsonFilePath);

  return CL.decompressJSON(s).object;
};
CL.ReadJSONFile = CL.readJSONFile;

CL.readYAMLFile = function (yamlFilePath) {
  var s = CL.readTextFileUTF8(yamlFilePath);

  return jsyaml.safeLoad(s);
};

CL.readJSONFromSheet = function (jsonSheet) {
  var jsonLastCell = getLastCellInColumn(jsonSheet, 1);
  var json;

  if (jsonLastCell.Row >= 2) {
    json = jsonSheet.Range(jsonSheet.Cells(1, 1), jsonLastCell).Value.toArray().join("\n");
  }
  else {
    json = jsonLastCell.Value;
  }

  return JSON.parse(json);
};
CL.ReadJSONFromSheet = CL.readJSONFromSheet;

CL.writeJSONToSheet = function (object, sheet) {
  var sJSON = JSON.stringify(object, undefined, 4);
  var sJSONArray = sJSON.split("\n");
  var excelArray = jsArray1dColumnMajorToSafeArray2d(sJSONArray, sJSONArray.length);

  // まずクリア
  sheet.Cells.ClearContents();

  sheet.Cells(1, 1).Resize(sJSONArray.length, 1) = excelArray;
};
CL.WriteJSONToSheet = CL.writeJSONToSheet;


// 2d 配列を転置したものを返す
CL.array2dTransposed = function (array) {
  var n1 = array[0].length;
  var n2 = array.length;

  var a = new Array(n1);
  for (var i = 0; i < n1; i++) {
    a[i] = new Array(n2);
  }

  for (var i = 0; i < n1; i++) {
    for (var j = 0; j < n2; j++) {
      a[i][j] = array[j][i];
    }
  }

  return a;
};
CL.Array2dTransposed = CL.array2dTransposed;

// excel の Range.Value.toArray() で取得した配列を a[row(y)][column(x)] な配列に変換
// 処理的にはどうってことないはずなので扱いやすい形に変換してしまう
CL.rangeToValueArray2d = function (range) {
  var rows = range.Rows.Count;
  var array = range.Value.toArray();
  var a = new Array(rows);

  for (var y = 0; y < rows; y++) {
    a[y] = [];
  }
  for (var i = 0; i < array.length;) {
    for (var y = 0; y < rows; y++) {
      a[y].push(array[i++]);
    }
  }

  return a;
};
CL.RangeToValueArray2d = CL.rangeToValueArray2d;

// excel の Range.Value.toArray() で取得した配列を a[column(x)][row(y)] な配列に変換
CL.rangeToValueArray2dColumnMajor = function (range) {
  var columns = range.Columns.Count;
  var rows = range.Rows.Count;
  var array = range.Value.toArray();
  var a = new Array(columns);

  for (var x = 0, i = 0; x < columns; x++ , i += rows) {
    a[x] = array.slice(i, i + rows);
  }

  return a;
};
CL.RangeToValueArray2dColumnMajor = CL.rangeToValueArray2dColumnMajor;


// fun は true を返せばそれ以降の traverse を打ち切る
CL.forAllNodes = function (node, parent, fun) {
  if (node === null) {
    return false;
  }
  if (fun(node, parent)) {
    return true;
  }

  for (var i = 0; i < node.children.length; i++) {
    if (CL.forAllNodes(node.children[i], node, fun)) {
      return true;
    }
  }

  return false;
};
CL.ForAllNodes = CL.forAllNodes;

// ULの各グループの幅を配列で返す
CL.getMaxItemWidth = function (node)
{
    var max = [];   // group 毎

    CL.forAllNodes(node, null, function(node) {
        if (node.kind !== CL.kind.UL) {
            return;
        }

        if (typeof max[node.group] === "undefined") {
            max[node.group] = node.depthInGroup + 1;
        }
        else {
            max[node.group] = Math.max(max[node.group], node.depthInGroup + 1);
        }
    });

    return max;
}

CL.getCheckHeaders = function (nodeH1, checkSheetTableData) {
  if (!nodeH1.tableHeaders) {
    return [ checkSheetTableData.input.header ];
  }
  return nodeH1.tableHeaders.map(function(x) {
    return x.name
  });
}


CL.getLeafNodes = function (node) {
  var leaves = [];
  CL.ForAllNodes(node, null, function (node, parent) {
    if (node.children.length === 0) {
      leaves.push(node);
    }
  });
  return leaves;
}
CL.GetLeafNodes = CL.getLeafNodes;

CL.nodeGetNumLeaves = function (node) {
  return CL.GetLeafNodes(node).length;
};
CL.NodeGetNumLeaves = CL.nodeGetNumLeaves;

CL.deletePropertyForAllNodes = function (node, propertyName) {
  CL.ForAllNodes(node, null, function (node, parent) {
    if (propertyName in node) {
      delete node[propertyName];
    }
  });
};
CL.DeletePropertyForAllNodes = CL.deletePropertyForAllNodes;

CL.addParentPropertyForAllNodes = function (node) {
  CL.ForAllNodes(node, null, function (node, parent) {
    node.parent = parent;
  });
};
CL.AddParentPropertyForAllNodes = CL.addParentPropertyForAllNodes;


// ID を基に node を取得
// leaf にしか ID はふられてないので、返る node は leaf になるはずだけど、 leaf 以外が返ったとしても特に問題ない作りのはず
// シート名を変更したい場合もあるはずなので、毎回すべてを検索するべき
// TODO: indexValues 用に level1 の H node 調べる用に maxDepth を渡せるようにしても良いか
CL.FindNodeById = function (node, id) {
  var resultNode = null;
  CL.ForAllNodes(node, null, function (node, parent) {
    if (node.id === id) {
      resultNode = node;
      // id はユニークという前提なので、１つ見つかった時点で終了して良い
      return true;
    }
  });
  return resultNode;
};

// 階層を考慮して id で検索
// idPath 通りの id の並び（idPath の末尾まで一致）の node を返す
// idPath には親から順に格納された配列を渡す
CL.FindNodeByIdPath = function (node, idPath) {
  if (idPath.length === 0) {
    return null;
  }

  var currentIdPath = [];

  function recurse(node) {
    if (node.id) {
      var i = currentIdPath.length;

      if (idPath[i] !== node.id) {
        return null;
      }
      // idPath の末尾まで一致してた
      // id path はユニークという前提なので、１つ見つかった時点で終了して良い
      if (i === idPath.length) {
        return node;
      }

      // push で idPath と同じ長さになる
      if (i + 1 >= idPath.length) {
        return null;
      }

      currentIdPath.push(node.id);
    }

    if (currentIdPath.length < idPath.length)

      for (var i = 0; i < node.children.length; i++) {
        var result = recurse(node.children[i]);

        if (result) {
          return result;
        }
      }

    if (node.id) {
      currentIdPath.pop();
    }

    return null;
  }

  return recurse(node);
};


CL.yyyymmddhhmmss = function (date) {
  // 1桁の数字を0埋めして2桁に
  function zeroPadding(value) {
    return ('0' + value).slice(-2);
    //return (value < 10) ? "0" + value : value;
  }
  var sa =
    [
      date.getFullYear(),
      zeroPadding(date.getMonth() + 1),
      zeroPadding(date.getDate()),
      zeroPadding(date.getHours()),
      zeroPadding(date.getMinutes()),
      zeroPadding(date.getSeconds())
    ];
  return sa.join("");
};

// フォルダが存在しなければ作成
// フォルダ名として作れないパスを渡された場合は無視
CL.createFolder = function (folderPath)
{
  var fso = new ActiveXObject("Scripting.FileSystemObject");

  function recurse(folderPath) {
    var parentFolderPath = fso.GetParentFolderName(folderPath);
    // 少なくともここで対象としているフォルダはファイルが置かれている場所より下の階層なので、rootまで遡ってしまうことは考慮しなくていいけど、一応
    if (parentFolderPath !== "" && !fso.FolderExists(parentFolderPath)) {
      recurse(parentFolderPath);
    }

    if (!fso.FolderExists(folderPath)) {
      try {
        fso.CreateFolder(folderPath);
      } catch (e) {
      }
    }
  }

  recurse(folderPath);
}

// 指定したフォルダ（相対パス。なければ作る）にファイルを移動
CL.moveFile = function (filePath, relativeFolderPath) {
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  var parentFolderPath = fso.GetParentFolderName(filePath);
  var dstFolderPath = fso.BuildPath(parentFolderPath, relativeFolderPath);
  var fileName = fso.GetFileName(filePath);
  var dstFilePath = fso.BuildPath(dstFolderPath, fileName);

  // なければ作る
  CL.createFolder(dstFolderPath);
  fso.MoveFile(filePath, dstFilePath);
};

// DateLastModified をつけたファイル名を生成
CL.makeBackupFileName = function (filePath, fso) {
  if (typeof fso === "undefined") {
    fso = new ActiveXObject("Scripting.FileSystemObject");
  }
  var file = fso.GetFile(filePath);
  var lastModifiedDate = CL.yyyymmddhhmmss(new Date(file.DateLastModified)).slice(2);
  var backupFileName = fso.GetBaseName(filePath) + "-bak" + lastModifiedDate + "." + fso.GetExtensionName(filePath);

  return backupFileName;
};

// ファイルのバックアップ作成
// 更新日時をファイル名に追加したような名前でコピーする
CL.makeBackupFile = function (filePath, relativeFolderPath) {
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  var backupFolderPath = fso.GetParentFolderName(filePath);
  if (typeof relativeFolderPath !== "undefined") {
    backupFolderPath = fso.BuildPath(backupFolderPath, relativeFolderPath);
    CL.createFolder(backupFolderPath);
  }
  var backupFileName = CL.makeBackupFileName(filePath, fso);
  var backupFilePath = fso.BuildPath(backupFolderPath, backupFileName);

  fso.CopyFile(filePath, backupFilePath);
};
CL.MakeBackupFile = CL.makeBackupFile;

// https://dobon.net/vb/dotnet/file/getabsolutepath.html#section4 をそのまま移植
CL.getRelativePath = function (basePath, absolutePath) {
  if (basePath == null || basePath.length == 0) {
      return absolutePath;
  }
  if (absolutePath == null || absolutePath.length == 0) {
      return "";
  }

  var directorySeparatorChar = "\\";
  var parentDirectoryString = ".." + directorySeparatorChar;

  basePath = _.trimRight(basePath, directorySeparatorChar);

  //パスを"\"で分割する
  var basePathDirs = basePath.split(directorySeparatorChar);
  var absolutePathDirs = absolutePath.split(directorySeparatorChar);

  //基準パスと絶対パスで、先頭から共通する部分を探す
  var commonCount = 0;
  for (var i = 0;
      i < basePathDirs.length &&
      i < absolutePathDirs.length &&
      basePathDirs[i].toUpperCase() === absolutePathDirs[i].toUpperCase();
      i++) {
      //共通部分の数を覚えておく
      commonCount++;
  }

  //共通部分がない時
  if (commonCount == 0) {
      return absolutePath;
  }

  //共通部分以降の基準パスのフォルダの深さを取得する
  var baseOnlyCount = basePathDirs.length - commonCount;
  //その数だけ"..\"を付ける
  var buf = _.repeat(parentDirectoryString, baseOnlyCount);

  //共通部分以降の絶対パス部分を追加する
  buf += absolutePathDirs.slice(commonCount).join(directorySeparatorChar);

  return buf;
}


CL.createRandomId = function (len) {
  var c = "abcdefghijklmnopqrstuvwxyz";
  var s = c.charAt(Math.floor(Math.random() * c.length));
  c += "0123456789";
  var cl = c.length;

  for (var i = 1; i < len; i++) {
    s += c.charAt(Math.floor(Math.random() * cl));
  }

  return s;
};

CL.convertUt2Sn = function(unixTimeMillis){ // UNIX時間(ミリ秒)→シリアル値
  var COEFFICIENT = 24 * 60 * 60 * 1000; //日数とミリ秒を変換する係数

  var DATES_OFFSET = 70 * 365 + 17 + 1 + 1; //「1900/1/0」～「1970/1/1」 (日数)
  var MILLIS_DIFFERENCE = 9 * 60 * 60 * 1000; //UTCとJSTの時差 (ミリ秒)

  return (unixTimeMillis + MILLIS_DIFFERENCE) / COEFFICIENT + DATES_OFFSET;
}

CL.yyyymmddhhmmssExcelFormat = function (date) {
  // 1桁の数字を0埋めして2桁に
  function zeroPadding(value) {
    return ('0' + value).slice(-2);
    //return (value < 10) ? "0" + value : value;
  }

  var s = "{0}/{1}/{2} {3}:{4}".format(
    date.getFullYear(),
    zeroPadding(date.getMonth() + 1),
    zeroPadding(date.getDate()),
    zeroPadding(date.getHours()),
    zeroPadding(date.getMinutes())
  );
  
  return s;
};

// sheet に更新履歴を書き出す
CL.renderHistoryToSheet = function(dstSheet, history, root, excel, templateData)
{
  var checkSheetTableRow = 0;
  if (!_.isUndefined(templateData)) {
    checkSheetTableRow = templateData.checkSheet.table.row;
  }

  function getIdToY(leafNodes) {
    var idToY = [];

    for (var i = 0; i < leafNodes.length; i++)
    {
        idToY[leafNodes[i].id] = i;
    }

    return idToY;
  }

  function findH1NodeById(id) {
    var nodeH1 = null;
    CL.forAllNodes(root, null, function (node, parent) {
      if (parent !== root) {
        return false;
      }
      if (node.id === id) {
        nodeH1 = node;
        return true;
      }
      return false;
    });
    return nodeH1;
  }

  var array = [];

  var headers = ["Rev.", "ユーザー", "日付", "シート", "行", "項目", "列", "元", "変更後"];
  array.push(headers);

  // 新しい順に並び替え
  var changeSets = history.changeSets.slice(0).reverse();

  // XXX: 力技で deep copy…
  var data = JSON.parse(JSON.stringify(history.data));

  for (var i = 0; i < changeSets.length; i++)
  {
    var changeSet = changeSets[i];
    if (!changeSet.changes) {
      continue;
    }
    var row = [];
    row.push(changeSet.revision);
    row.push(changeSet.author);
    row.push(CL.yyyymmddhhmmssExcelFormat(new Date(changeSet.date)));
    var sheets = changeSet.changes.checkSheet.sheets;
    for (var sheetId in sheets)
    {
      var sheet = sheets[sheetId];
      var sheetData = data.checkSheet.sheets[sheetId];
      var nodeH1 = findH1NodeById(sheetId);
      var idToY = getIdToY(CL.getLeafNodes(nodeH1));
      row.push(sheet.text);
      for (var id in sheet.items)
      {
        var item = sheet.items[id];
        var itemData = sheetData.items[id];
        row.push(checkSheetTableRow + idToY[id]); // TODO: リンク張る
        row.push(_.trunc(item.text));  // TODO: パスを入れる
        for (var header in item.values)
        {
          var value0 = item.values[header];
          var value1 = itemData.values[header];
          itemData.values[header] = value0;
          row.push(header);
          row.push(value0 === null ? "" : value0);
          row.push(value1);
          array.push(row.slice(0));
          row.pop();
          row.pop();
          row.pop();
        }
        row.pop();
        row.pop();
      }
      row.pop();
    }
  }
  
  excel.ScreenUpdating = false;

  var range = dstSheet.Cells(1, 1).Resize(array.length, headers.length);
  range.Value = jsArray2dToSafeArray2d(array);

  var xlSrcRange = 1;
  var xlYes = 1;
  //var table = dstSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$D$5"), , xlYes).Name = "テーブル";
  //dstSheet.Range("テーブル[#All]").ListObjects("テーブル").TableStyle = "TableStyleMedium2"
  var listObject = dstSheet.ListObjects.Add(xlSrcRange, dstSheet.UsedRange, null, xlYes);
  listObject.TableStyle = "TableStyleMedium2";
  //listObject.ShowTotals = true;

  dstSheet.UsedRange.EntireColumn.AutoFit;

  dstSheet.Cells(1, 1).Resize(1, headers.length).Select();
  excel.ActiveWindow.Zoom = true;
  dstSheet.Cells(2, 1).Select();  // 何となく

  excel.ScreenUpdating = true;
};

CL.addSheetToEndOfBook = function (book, name, visible) {
    var sheet = book.Worksheets.Add();
    sheet.Name = name;
    // 非表示のシートが末尾にある場合は、その非表示のシートより前に移動する仕様っぽい
    sheet.Move(null, book.Worksheets(book.Worksheets.Count));
    sheet.Visible = visible;

    return sheet;
};

// 更新履歴シート作成
CL.createChangelogSheet = function(book, history, root, excel, templateData) {
  var changelogSheet = findSheetByName(book, "changelog");
  // あれば削除して作り直す
  if (changelogSheet)
  {
      excel.DisplayAlerts = false;
      changelogSheet.Delete();
      excel.DisplayAlerts = true;
  }

  // revision 0 の場合は作らない
  if (history.head === 0) {
    return null;
  }

  changelogSheet = CL.addSheetToEndOfBook(book, "changelog", true);
  // 先頭に移動して選択
  // → やっぱり先頭に移動はやめておく
  //changelogSheet.Move(book.Worksheets(1), null);
  changelogSheet.Select();
  CL.renderHistoryToSheet(changelogSheet, history, root, excel, templateData);

  return changelogSheet;
};

// Excel等のファイルがすでに開かれているか判定
CL.isFileOpened = function(filePath) {
  var fso = new ActiveXObject("Scripting.FileSystemObject");
  try {
    // 同じ名前にリネームしてみてエラーになるかどうかで判定
    // ファイルをロックする系のアプリで開かれているかはこれで判定可能とのこと
    fso.MoveFile(filePath, filePath);
  } catch(e) {
    return true;
  }
  return false;
};

CL.getIndexSheet = function(book, root) {
//  if (root.variables.sheetname) {
//    root.variables.indexSheetname = root.variables.sheetname;
//  }
//  var sheetname = root.variables.indexSheetname;
//
//  sheetname = _.isUndefined(sheetname) ? "index" : sheetname;

  var sheetname = root.variables.indexSheetname || root.variables.sheetname || "index";

  return findSheetByName(book, sheetname);
};
