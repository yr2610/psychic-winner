// 定数関係をとりあえず全部ここに突っ込んでおく


//  ボタンの種類
var BTN_OK                 = 0;    // [ＯＫ]ボタン
var BTN_OK_CANCL           = 1;    // [ＯＫ][キャンセル]ボタン
var BTN_STOP_RETRI_DISRGRD = 2;    // [中止][再試行][無視]ボタン
var BTN_YES_NO_CANCL       = 3;    // [はい][いいえ][キャンセル]ボタン
var BTN_YES_NO             = 4;    // [はい][いいえ]ボタン
var BTN_RETRI_CANCL        = 5;    // [再試行][キャンセル]ボタン

//  アイコンの種類
var ICON_STOP              = 16;   // [Stop]アイコン
var ICON_QUESTN            = 32;   // [?]アイコン
var ICON_EXCLA             = 48;   // [!]アイコン
var ICON_I                 = 64;   // [i]アイコン

//  押されたボタンごとの戻り値
var BTNR_OK                =  1;   // [ＯＫ]ボタン押下時
var BTNR_CANCL             =  2;   // [キャンセル]ボタン押下時
var BTNR_STOP              =  3;   // [中止]ボタン押下時
var BTNR_RETRI             =  4;   // [再試行]ボタン押下時
var BTNR_DISRGRD           =  5;   // [無視]ボタン押下時
var BTNR_YES               =  6;   // [はい]ボタン押下時
var BTNR_NO                =  7;   // [いいえ]ボタン押下時
var BTNR_NOT               = -1;   // どのボタンも押さなかったとき



//  オープンモード
var FORREADING      = 1;    // 読み取り専用
var FORWRITING      = 2;    // 書き込み専用
var FORAPPENDING    = 8;    // 追加書き込み

//  開くファイルの形式
var TRISTATE_TRUE       = -1;   // Unicode
var TRISTATE_FALSE      =  0;   // ASCII
var TRISTATE_USEDEFAULT = -2;   // システムデフォルト



// 保存データの種類
		// StreamTypeEnum
		// http://msdn.microsoft.com/ja-jp/library/cc389884.aspx
var adTypeBinary = 1; // バイナリ
var adTypeText   = 2; // テキスト

// 読み込み方法
		// StreamReadEnum
		// http://msdn.microsoft.com/ja-jp/library/cc389881.aspx
var adReadAll  = -1; // 全行
var adReadLine = -2; // 一行ごと

// 書き込み方法
		// StreamWriteEnum
		// http://msdn.microsoft.com/ja-jp/library/cc389886.aspx
var adWriteChar = 0; // 改行なし
var adWriteLine = 1; // 改行あり

// ファイルの保存方法
		// SaveOptionsEnum 
		// http://msdn.microsoft.com/ja-jp/library/cc389870.aspx
var adSaveCreateNotExist  = 1; // ない場合は新規作成
var adSaveCreateOverWrite = 2; // ある場合は上書き
