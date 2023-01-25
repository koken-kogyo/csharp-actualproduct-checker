using Oracle.ManagedDataAccess.Client;

namespace ActualProductChecker
{

    class Common
    {
        // プログラムタイトル
        public const string PROGRAM_TITLE = "[KMC006SC] 製造実績ファイル照合";
        public const string PROGRAM_NAME = "ActualProductChecker";

        // 定数定義
        public const string CONFIG_FILE_DB = "ConfigDB.xml";        // データベース設定 ファイル名
        public const string CONFIG_FILE_FS = "ConfigFS.xml";        // ファイルシステム設定 ファイル名
        public const string CONFIG_FILE_SMTP = "ConfigSMTP.xml";    // SMTP設定 ファイル名
        public const string LOG_FILE = PROGRAM_NAME + ".log";       // ログ ファイル名
        public const string ERROR_FILE_LIST = "error_list.txt";     // エラー一覧 ファイル名
        public const string JIDOU_LOG_PATH = @"D:\\IREPOEXE\\exe";  // 自動受入のアプリログの場所
        public const string JIDOU_LOG_FILE = "batch.log";           // 自動受入のアプリログファイル名
        public const int RETRY_MAX = 5;                             // 再試行回数上限

        // i-Reporter 対象工程マスタ (A70)
        public const string TABLE_NAME_A70 = "A70";                 // テーブル名称: A70
        public const string A70_ACTIVE_ACTIVE = "1";                // 有効フラグ (1:有効)
        public const string A70_OPESTAT_OPARATED = "3";             // 運用状態フラグ (3:運用中)

        public const string TABLE_NAME_KM0040 = "KM0040";           // テーブル名称: KM0040
        public const string PGM_ID = "KMC006";                      // プログラムID
        public const string EMAIL_ADRTYPE_TO = "1";                 // 宛先種別 (1:To)
        public const string EMAIL_ADRTYPE_CC = "2";                 // 宛先種別 (2:Cc)
        public const string EMAIL_ADRTYPE_BCC = "3";                // 宛先種別 (3:Bcc)

        public const int CONNECT_SOURCE = 0;                        // 接続先 (設定ファイル[0])
        public const int CONNECT_DEST = 1;                          // 接続先 (設定ファイル[1])
        public const string DI_BACKUP = "backup";                   // DirectoryInfo
        public const string DI_ERROR = "error";                     // DirectoryInfo

        // メールタイトル
        public const string MAIL_SUBJECT = "[自動通知] iRepo製造実績ファイル照合結果";

        // メール本文
        public const string MAIL_BODY_HEADER = 
            "各位\n\n本日の製造実績ファイル照合結果\n\n";

        // メールフッター
        public const string MAIL_BODY_FOOTER = 
            "------------------------------------------------------------------------\n" +
            PROGRAM_TITLE + "\n" +
            "このメールは、プログラムが自動送信したものです。\n" +
            "------------------------------------------------------------------------\n";

        // Common関数

        // 変数定義

        // プロパティ
        public OracleConnection Cnn { get; set; }       // Oracle データ接続 (Managed)
        public DBConfigData DbCd { get; set; }          // データベース設定データ
        public FSConfigData[] FsCd { get; set; }        // ファイルシステム設定データ
        public SMTPConfigData SMTPCd { get; set; }      // SMTP設定データ
        public DBAccess Dba { get; set; }               // データベース アクセス
        public FileAccess Fa { get; set; }              // ファイル アクセス
        public string BaseDir { get; set; }             // 実行ファイルのあるディレクトリ
        public string FsConfFilePath { get; set; }      // ファイル システム設定ファイルへのパス
        public string DbConfFilePath { get; set; }      // データベース設定ファイルへのパス
        public string MailConfFilePath { get; set; }    // メール関連設定ファイルへのパス
    }

    /// <summary>
    /// データベース設定データ クラス
    /// </summary>
    public class DBConfigData
    {
        // プロパティ
        public string User { get; set; }        // ユーザー ID
        public string Passwd { get; set; }      // パスワード
        public string Protocol { get; set; }    // プロトコル
        public string Host { get; set; }        // ホスト名 (IP アドレス)
        public int Port { get; set; }           // ポート番号
        public string ServiceName { get; set; } // サービス名
        public string Schema { get; set; }      // スキーマ
    }

    /// <summary>
    /// ファイルシステム設定データ クラス
    /// </summary>
    public class FSConfigData
    {
        // プロパティ
        public string HostName { get; set; }    // ホスト名
        public string IpAddr { get; set; }      // IP アドレス
        public string UserId { get; set; }      // ユーザー ID
        public string Passwd { get; set; }      // パスワード
        public string ShareName { get; set; }   // 共有名
        public string RootPath { get; set; }    // i-Reporter 完了帳票保存先のルート パス
        public string FileFilter { get; set; }  // 監視対象ファイルのフィルター
        public string CreateSubDir { get; set; }// サブ ディレクトリ作成
    }

    /// <summary>
    /// SMTP設定データ クラス
    /// </summary>
    public class SMTPConfigData
    {
        // プロパティ
        public string Server { get; set; }      // SMTP アドレス
        public int Port { get; set; }           // SMTP ポート番号
        public string LoginUser { get; set; }   // SMTP ログインユーザーID
        public string LoginPass { get; set; }   // SMTP ログインパスワード
        public string From { get; set; }        // 送信者メールアドレス
    }
}
