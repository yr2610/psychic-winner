// ただ解凍するだけの
// 開発用

function main() {
    var shell = new ActiveXObject("WScript.Shell");
    var shellApplication = new ActiveXObject("Shell.Application");

    function error(message) {
        shell.Popup(message, 0, "エラー", ICON_EXCLA);
        WScript.Quit();
    }

    function getOutNameFromFile(filePath) {
        var fso = new ActiveXObject( "Scripting.FileSystemObject" );
        var extensionName = fso.GetExtensionName(filePath);
        var baseName = fso.GetBaseName(filePath);
        var parentFolderName = fso.GetParentFolderName(filePath);

        return fso.BuildPath(parentFolderName, baseName + "-decompressed." + extensionName);
    }

    if (( WScript.Arguments.length != 1 ) ||
        ( WScript.Arguments.Unnamed(0) == "")) {
        error("解凍したいJSONファイルをドラッグ＆ドロップしてください。");
    }

    var filePath = WScript.Arguments.Unnamed(0);

    var s;
    try {
        s = CL.readTextFileUTF8(filePath);
    } catch (e) {
        error("JSON ファイルの読み込みに失敗しました。");
    }

    var o;
    try {
        o = JSON.parse(s);
    } catch (e) {
        error("JSON のパースに失敗しました。");
    }

    if (_.isUndefined(o.compress)) {
        WScript.Echo("圧縮形式ではありません。");
        WScript.Quit();
    }

    var json = CL.decompressJSON(s).json;
    var outfilePath = getOutNameFromFile(filePath);

    CL.writeTextFileUTF8(json, outfilePath);

    var fso = new ActiveXObject( "Scripting.FileSystemObject" );
    var outfileName = fso.GetFileName(outfilePath);

    WScript.Echo("解凍済みファイル(" + outfileName + ")を出力しました");

    WScript.Quit();
}

main();
