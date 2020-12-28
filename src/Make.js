function Error(message) {
    shell.Popup(message, 0, "エラー", ICON_EXCLA);
    WScript.Quit(1);
}

var shell = new ActiveXObject("WScript.Shell");
var shellApplication = new ActiveXObject("Shell.Application");

var fso = new ActiveXObject( "Scripting.FileSystemObject" );

var args = WScript.Arguments;
var argsNamed = args.Named;
var argsUnnamed = WScript.Arguments.Unnamed;

if (( args.length != 1 ) ||
    ( argsUnnamed(0) == "")) {
    Error("チェックリストのソースファイル（.txt）をドロップしてください。");
}

var scriptFolderName = fso.getParentFolderName(WScript.ScriptFullName);

//var argsNamed = WScript.Arguments.Named;
//
//if (argsNamed.Exists("o")) {
//    WScript.Echo(argsNamed.Item("o"));
//}

var filePath = '"' + argsUnnamed(0) + '"';

var parserCommand = "cscript /nologo "+ fso.BuildPath(scriptFolderName, "Parse.wsf") + " " + filePath;
//Error(command);

var parserExec = shell.Exec(parserCommand);

//  実行中の間ループ
var cnt = 0;
while (parserExec.Status == 0) {
     WScript.Sleep(100);
     cnt++;

    //  30秒経過したら終了
    if (cnt  >=  300) {
        parserExec.Terminate();
        Error("parser で問題発生（タイムアウト）");
    }
}

if (parserExec.ExitCode != 0) {
    var message;

    if (!parserExec.StdErr.AtEndOfStream) {
        message = parserExec.StdErr.ReadAll();
    }
    else {
        message = "parser でエラーが発生しました";
    }

    Error(message);
}
{
    if (!parserExec.StdOut.AtEndOfStream) {
        WScript.Echo(parserExec.StdOut.ReadAll());
    }
}

// TODO: parser の出力 json 名は外から指定できるようにする
var jsonFilename = fso.GetBaseName(argsUnnamed(0)) + ".json";
var jsonFilePath = fso.BuildPath(fso.GetParentFolderName(argsUnnamed(0)), jsonFilename);
jsonFilePath = '"' + jsonFilePath + '"';

var rendererCommand = "wscript "+ fso.BuildPath(scriptFolderName, "Render.wsf") + " " + jsonFilePath;

var rendererExitCode = shell.Run(rendererCommand, 1, true);

// TODO: 出力先の excel と同名の excel ファイルが開いてるか確認
//var isExcelFileOpened = CL.isFileOpened(filePath);

// TODO: 出力先を指定できるように
// 

//var exec = shell.Exec(cmd); // コマンドを実行
//while (exec.Status == 0) {
//  WScript.Sleep(100);
//}
//
//var r = oe.StdOut.ReadAll();
//
//WScript.Echo(fso.GetFileName(outFilePath) + "\nを出力しました");

WScript.Quit(rendererExitCode);
