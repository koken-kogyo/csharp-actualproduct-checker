using System;
using System.IO;
using System.Text;
using System.Xml.Serialization;
using System.Runtime.InteropServices; // DllImport
using DecryptPassword;
using System.Linq;

namespace ActualProductChecker
{
    /// <summary>
    /// ファイル アクセス クラス
    /// </summary>
    class FileAccess
    {

        // Win32エラー・コードを取得する
        public const uint FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000;
        [DllImport("kernel32.dll")]
        public static extern uint FormatMessage(uint dwFlags, IntPtr lpSource, uint dwMessageId, uint dwLanguageId, StringBuilder lpBuffer, int nSize, IntPtr Arguments);

        // 接続切断する Win32 API を宣言
        [DllImport("mpr.dll", EntryPoint = "WNetCancelConnection2", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        public static extern int WNetCancelConnection2(string lpName, Int32 dwFlags, bool fForce);

        // 認証情報を使って接続する Win32 API 宣言
        [DllImport("mpr.dll", EntryPoint = "WNetAddConnection2", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        public static extern int WNetAddConnection2(ref NETRESOURCE lpNetResource, string lpPassword, string lpUsername, Int32 dwFlags);

        // WNetAddConnection2 に渡す接続の詳細情報の構造体
        [StructLayout(LayoutKind.Sequential)]
        // 接続情報
        public struct NETRESOURCE
        {
            public int dwScope; // 列挙の範囲
            public int dwType; // リソースタイプ
            public int dwDisplayType; // 表示オブジェクト
            public int dwUsage; // リソースの使用方法
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpLocalName; // ローカルデバイス名。使わないなら NULL
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpRemoteName; // リモートネットワーク名。使わないなら NULL
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpComment; // ネットワーク内の提供者に提供された文字列
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpProvider; // リソースを所有しているプロバイダ名
        }

        /// <summary>
        /// データベース設定ファイルの逆シリアライズ
        /// </summary>
        /// <param name="cmn">共通クラス</param>
        /// <returns></returns>
        public DBConfigData ReserializeDBConfigFile(Common cmn)
        {
            // 設定ファイル名
            string fileName = @cmn.DbConfFilePath;

            // XmlSerializerオブジェクトを作成
            XmlSerializer serializer = new XmlSerializer(typeof(DBConfigData));

            // 読み込むファイルを開く
            // 逆シリアライズ化
            using (StreamReader sr = new StreamReader(fileName, new System.Text.UTF8Encoding(false)))
            {
                // XMLファイルから読み込み、逆シリアル化する
                DBConfigData obj = (DBConfigData)serializer.Deserialize(sr);

                // ファイルを閉じる
                sr.Close();

                // パスワードをデコード
                if (obj.Passwd != null && obj.Passwd != "")
                {
                    DecryptPasswordClass decryptPasswordClass = new DecryptPasswordClass();
                    string cryptographyPassword = obj.Passwd;
                    string decryptionPassword;
                    if (decryptPasswordClass.DecryptPassword(cryptographyPassword, out decryptionPassword))
                    {
                        // デコードしたパスワードを格納
                        obj.Passwd = decryptionPassword;
                    }
                }
                return obj;
            }
        }

        /// <summary>
        /// ファイル システム設定ファイルの逆シリアライズ
        /// </summary>
        /// <param name="cmn">共通クラス</param>
        /// <returns></returns>
        public FSConfigData[] ReserializeFSConfigFile(Common cmn)
        {
            // 設定ファイル名
            string fileName = @cmn.FsConfFilePath;

            // XmlSerializer オブジェクトを作成
            XmlSerializer serializer = new XmlSerializer(typeof(FSConfigData[]));

            // 逆シリアライズ化
            using (StreamReader sr = new StreamReader(fileName, new System.Text.UTF8Encoding(false)))
            {
                // XMLファイルから読み込み、逆シリアル化する
                FSConfigData[] loadAry;
                loadAry = (FSConfigData[])serializer.Deserialize(sr);

                // ファイルを閉じる
                sr.Close();

                // パスワードをデコード
                DecryptPasswordClass dp = new DecryptPasswordClass();
                for (var i = 0; i < loadAry.Length; i++)
                {
                    if (loadAry[i].Passwd != null && loadAry[i].Passwd != "")
                    {
                        string decryptionPassword;
                        if (dp.DecryptPassword(loadAry[i].Passwd, out decryptionPassword))
                        {
                            // デコードしたパスワードを格納
                            loadAry[i].Passwd = decryptionPassword;
                        }
                    }
                }
                return loadAry;
            }
        }

        /// <summary>
        /// メール関連設定ファイルの逆シリアライズ
        /// </summary>
        /// <param name="cmn">共通クラス</param>
        /// <returns></returns>
        /// [C#]階層化された繰返し要素のXMLをデシリアライズする
        /// https://zero-config.com/dotnet/xmlserializer002.html
        public SMTPConfigData ReserializeMailConfigFile(Common cmn)
        {
            // 設定ファイル名
            string fileName = @cmn.MailConfFilePath;

            // XmlSerializerオブジェクトを作成
            XmlSerializer serializer = new XmlSerializer(typeof(SMTPConfigData));

            // 読み込むファイルを開く
            // 逆シリアライズ化
            using (StreamReader sr = new StreamReader(fileName, new System.Text.UTF8Encoding(false)))
            {
                // XMLファイルから読み込み、逆シリアル化する
                SMTPConfigData obj = (SMTPConfigData)serializer.Deserialize(sr);

                // ファイルを閉じる
                sr.Close();

                // パスワードをデコード
                if (obj.LoginPass != null && obj.LoginPass != "")
                {
                    DecryptPasswordClass decryptPasswordClass = new DecryptPasswordClass();
                    string cryptographyPassword = obj.LoginPass;
                    string decryptionPassword;
                    if (decryptPasswordClass.DecryptPassword(cryptographyPassword, out decryptionPassword))
                    {
                        // デコードしたパスワードを格納
                        obj.LoginPass = decryptionPassword;
                    }
                }
                return obj;
            }
        }

        /// <summary>
        /// サーバーへの接続
        /// </summary>
        /// <param name="_cmn">共通クラス</param>
        /// <param name="_type">宛先種別 (0:Source 1:Dest)</param>
        /// <returns>結果 (true: 正常, false: 異常)</returns>
        public bool ConnectServer(Common _cmn, int _type)
        {
            // 接続情報を設定
            NETRESOURCE sourceNetResource = new NETRESOURCE();
            sourceNetResource.dwScope = 0;
            sourceNetResource.dwType = 1;
            sourceNetResource.dwDisplayType = 0;
            sourceNetResource.dwUsage = 0;
            sourceNetResource.lpLocalName = ""; // ネットワーク ドライブにする場合は "z:" などドライブレター設定  
            sourceNetResource.lpRemoteName = _cmn.FsCd[_type].ShareName;
            sourceNetResource.lpProvider = "";

            int ret = 0;
            try
            {
                // 既に接続している場合があるので一旦切断する
                ret = WNetCancelConnection2(_cmn.FsCd[_type].ShareName, 0, true);
                // 認証情報を使って共有フォルダに接続
                ret = WNetAddConnection2(ref sourceNetResource, _cmn.FsCd[_type].Passwd, _cmn.FsCd[_type].UserId, 0);

                if (ret != 0)
                {
                    StringBuilder message = new StringBuilder(255);
                    FormatMessage(
                        FORMAT_MESSAGE_FROM_SYSTEM,
                        IntPtr.Zero,
                        (uint)ret,
                        0,
                        message,
                        message.Capacity,
                        IntPtr.Zero);
                    Console.WriteLine("[" + _cmn.FsCd[_type].ShareName + "]" + message.ToString()); // Win32エラー・メッセージ
                    Console.Read();
                    return false;
                }
            }
            // 例外発生時
            catch (Exception ex)
            {
                // エラー処理
                Console.WriteLine(ex.Message);
                Console.Read();
                return false;
            }
            var msg = _type == Common.CONNECT_SOURCE ? "Source" : _type == Common.CONNECT_DEST ? "Dest" : "";
            Console.Write("Connected to ");
            Console.WriteLine(msg + "Dir [" + _cmn.FsCd[_type].ShareName + "]");
            return true;
        }

    }
}
