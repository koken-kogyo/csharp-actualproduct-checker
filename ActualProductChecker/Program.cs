using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

namespace ActualProductChecker
{
    internal class Program
    {
        private static bool fales;

        class Target
        {
            public Target(string folder) {
                Folder = folder;
                CntIn = 0;
                CntTg = 0;
                CntOk = 0;
                CntNg = 0;
                CntNothing = 0;
                AttachFiles = new List<string>();
            }

            public bool IsActivate() {
                if (CntOk + CntNg > 0) return true;
                return false;
            }

            public string Folder { get; set; }
            public int CntIn { get; set; }
            public int CntTg { get; set; }
            public int CntOk { get; set; }
            public int CntNg { get; set; }
            public int CntNothing { get; set; }
            public List<string> AttachFiles { get; set; }

        }

        static void Main(string[] args)
        {
            Console.WriteLine(Common.PROGRAM_TITLE + "\n");

            // お膳立て
            var cmn = new Common();                     // 共通クラス
            cmn.DbCd = new DBConfigData();              // データベース設定データ
            cmn.FsCd = new FSConfigData[2];             // 設定データ クラス
            cmn.SMTPCd = new SMTPConfigData();          // SMTP設定データ クラス
            cmn.BaseDir = AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\');          // 実行ファイルのあるディレクトリ
            cmn.DbConfFilePath = Path.Combine(cmn.BaseDir, Common.CONFIG_FILE_DB);      // データベース設定データは実行ファイルと同一ディレクトリ
            cmn.FsConfFilePath = Path.Combine(cmn.BaseDir, Common.CONFIG_FILE_FS);      // ファイルシステム設定データは実行ファイルと同一ディレクトリ
            cmn.MailConfFilePath = Path.Combine(cmn.BaseDir, Common.CONFIG_FILE_SMTP);  // SMTP設定データは実行ファイルと同一ディレクトリ
            cmn.Dba = new DBAccess();                           // データベース アクセス クラス
            cmn.Fa = new FileAccess();                          // ファイル アクセス クラス
            cmn.DbCd = cmn.Fa.ReserializeDBConfigFile(cmn);     // データベース設定ファイル逆シリアライズ
            cmn.FsCd = cmn.Fa.ReserializeFSConfigFile(cmn);     // ファイルシステム設定ファイル逆シリアライズ
            cmn.SMTPCd = cmn.Fa.ReserializeMailConfigFile(cmn); // メール送信者情報設定ファイル逆シリアライズ
            OracleConnection cnn = null;
            if (cmn.Dba.ConnectSchema(cmn, ref cnn))            // データベース接続
            {
                cmn.Cnn = cnn;
            } else { 
                Console.WriteLine("データベースに接続できませんでした．\nメール送信なしでファイルのチェックだけを行いますか？ (y/n)");
                if (Console.ReadKey().Key.ToString().ToLower() != "y") return;
                Console.WriteLine();
                cmn.Cnn = null;
            }



            // データベース関連情報取得
            var myDs = new DataSet();
            var myIRepoList = new List<string>();
            if (cmn.Cnn != null)
            {
                // データベースから送信先アドレスを取得
                if (!cmn.Dba.GetEmailAddress(cmn, Common.EMAIL_ADRTYPE_TO, ref myDs)) return;
                if (!cmn.Dba.GetEmailAddress(cmn, Common.EMAIL_ADRTYPE_CC, ref myDs)) return;
                // データベースから有効かつ運用中の工程名称を取得
                if (!cmn.Dba.GetIRepoKt(cmn, Common.A70_ACTIVE_ACTIVE, Common.A70_OPESTAT_OPARATED, ref myDs)) return;
                myIRepoList = myDs.Tables[Common.TABLE_NAME_A70].AsEnumerable()
                    .Select(row => row.Field<string>("IREPOKTNM")).ToList();
            }



            // サーバーのファイルシステムに接続
            var sourceFolder = @cmn.FsCd[Common.CONNECT_SOURCE].RootPath;
            var backupFolder = @cmn.FsCd[Common.CONNECT_DEST].RootPath + "\\" + Common.DI_BACKUP;
            var errorFolder = @cmn.FsCd[Common.CONNECT_DEST].RootPath + "\\" + Common.DI_ERROR;
            if (!cmn.Fa.ConnectServer(cmn, Common.CONNECT_SOURCE)) return;
            if (!cmn.Fa.ConnectServer(cmn, Common.CONNECT_DEST)) return;
            if (!Directory.Exists(backupFolder))
            {
#if DEBUG
                Console.WriteLine("自動受入バックアップ先フォルダーがみつかりませんでした．");
                Console.ReadKey();
#endif
                WriteLog("自動受入バックアップ先フォルダーがみつかりませんでした．");
                return;
            }
            if (!Directory.Exists(errorFolder))
            {
#if DEBUG
                Console.WriteLine("自動受入エラーフォルダーがみつかりませんでした．");
                Console.ReadKey();
#endif
                WriteLog("自動受入エラーフォルダーがみつかりませんでした．");
                return;
            }



            // [KMC006SC] 製造実績ファイル照合処理開始
            Console.WriteLine();
            var Targets = new List<Target>();
            if (args.Length == 0) //当日処理
            {
                var TargetDt = DateTime.Today; // DateTime.Parse("2020/9/24"); // DateTime.Today; //.AddDays(-1);  // 対象となる処理日付
                // 判定対象フォルダの存在チェック
                if (!Directory.Exists(sourceFolder + "\\" + TargetDt.ToString("yyyyMMdd")))
                {
                    WriteLog("本日の実績登録はありませんでした．");
                }
                else
                {
                    Targets.Add(new Target(TargetDt.ToString("yyyyMMdd")));
                    ActualProductChecker(cmn, myIRepoList, Targets[0]);
                }
            }
            else
            {
                if (args[0] == "/all")
                {
                    DirectoryInfo di = new DirectoryInfo(sourceFolder);
                    DirectoryInfo[] diAlls = di.GetDirectories();
                    foreach (DirectoryInfo d in diAlls)
                    {
                        Targets.Add(new Target(d.Name));
                        ActualProductChecker(cmn, myIRepoList, Targets[Targets.Count() - 1]);
                    }
                }
            }
            
            // 結果をメール送信
            var cIn = Targets.Select(i => i.CntIn).Sum();
            var cTg = Targets.Where(row => row.IsActivate()).Select(col => col.CntTg).Sum();
            var cOk = Targets.Where(row => row.IsActivate()).Select(col => col.CntOk).Sum();
            var cNg = Targets.Where(row => row.IsActivate()).Select(col => col.CntNg).Sum();
            var cNothing = Targets.Where(row => row.IsActivate()).Select(col => col.CntNothing).Sum();
            if (cmn.Cnn != null)
            {
                try
                {
                    MailMessage msg = new MailMessage();
                    msg.From = new MailAddress(cmn.SMTPCd.From); // 送信元("コーケン工業　システム管理者 <yoshiyuki-watanabe@koken-kogyo.co.jp>")
                    foreach (DataRow result in myDs.Tables["To"].Rows)
                    {
                        msg.To.Add(new MailAddress(result["EMAIL"].ToString()));
                    }
                    foreach (DataRow result in myDs.Tables["Cc"].Rows)
                    {
                        msg.CC.Add(new MailAddress(result["EMAIL"].ToString()));
                    }
                    msg.Subject = Common.MAIL_SUBJECT;
                    if (cIn == 0)
                    {
                        msg.Body = Common.MAIL_BODY_HEADER + 
                            "実績登録はありませんでした．\n\n" +
                            Common.MAIL_BODY_FOOTER;
                    }
                    else
                    {
                        msg.Body = Common.MAIL_BODY_HEADER +
                            "入力完了帳票：" + cIn + "件\n" +
                            "　受入対象：" + cTg + "件\n" +
                            "　自動受入成功：" + cOk + "件\n" +
                            "　自動受入失敗：" + cNg + "件\n\n" +
                            "　行方不明となったファイル：" + cNothing + "件\n" +
                            Common.MAIL_BODY_FOOTER;
                        // 自動受入ログの添付（実行中はファイルにアクセス出来ない為複写して添付）
                        if (File.Exists(Common.JIDOU_LOG_PATH + "\\" + Common.JIDOU_LOG_FILE))
                        {
                            var jidoulog = @cmn.BaseDir + "\\" + Common.JIDOU_LOG_FILE;
                            File.Copy(@Common.JIDOU_LOG_PATH + "\\" + Common.JIDOU_LOG_FILE, jidoulog, true);
                            Attachment attachjidou = new Attachment(jidoulog);
                            msg.Attachments.Add(attachjidou);
                        }

                        // 失敗もしくは不明があればファイルを添付する
                        var attachfiles = Targets
                            .Where(t => t.IsActivate() && t.AttachFiles.Count > 0)
                            .SelectMany(a => a.AttachFiles);
                        // 対象が10件以上ある場合はテキストファイルに内容を作成して添付する
                        var listflg = (attachfiles.ToList().Count > 10);
                        var errorlist = "";
                        foreach (var fl in attachfiles)
                        {
                            if (listflg)
                            {
                                errorlist += fl + "\n";
                            }
                            else
                            {
                                var attachment = new System.Net.Mail.Attachment(fl);
                                msg.Attachments.Add(attachment);
                            }
                        }
                        if (errorlist != "")
                        {
                            WriteList(errorlist);
                            System.Net.Mail.Attachment attachment;
                            attachment = new System.Net.Mail.Attachment(
                                @Path.Combine(Directory.GetCurrentDirectory(), Common.ERROR_FILE_LIST));
                            msg.Attachments.Add(attachment);
                        }
                    }

                    SmtpClient sc = new SmtpClient();

                    sc.EnableSsl = true;
                    sc.UseDefaultCredentials = false;
                    sc.Credentials = new System.Net.NetworkCredential(cmn.SMTPCd.LoginUser, cmn.SMTPCd.LoginPass); //uX[HgKn{djM56=xc
                    sc.Host = cmn.SMTPCd.Server; // "140.227.104.69";
                    sc.Port = cmn.SMTPCd.Port; // 587;
                    sc.DeliveryMethod = SmtpDeliveryMethod.Network;

                    // SSL証明書チェックのコールバック設定
                    ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ValidateServerCertificate);

                    // メール送信
                    sc.Send(msg);

                    // データベースクローズ
                    cmn.Dba.CloseSchema(cnn);

                    // 結果ログ出力
                    if (cIn != 0)
                    {
                        var tg = String.Format("{0,3}", cTg);
                        var ok = String.Format("{0,3}", cOk);
                        var mode = args.Length == 0 ? "" : "全件チェック ";
                        WriteLog(mode + $"自動受入 対象:{tg}件 / 正常:{ok}件 / 失敗:{cNg}件 / 不明:{cNothing}件");
                    }

                    Console.WriteLine("Send Mail Compleated !!");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }

        public static bool ValidateServerCertificate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            if (sslPolicyErrors == SslPolicyErrors.None)
                return true; // .NET Frameworkに対して「SSL証明書の使用は問題なし」
            else
            {
                // false の場合は「SSL証明書の使用は不可」
                // 何かしらのチェックを行いtrue / false を判定
                // このプログラムでは true を返却し、信頼されないSSL証明書の問題を回避
                return true; 
            }
        }

        private static void WriteLog(string message)
        {
            var mySW = new StreamWriter(
                @Path.Combine(Directory.GetCurrentDirectory(), Common.LOG_FILE)
                , true
                , System.Text.Encoding.GetEncoding("shift_jis"));
            mySW.WriteLine("[" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "] " + message);
            mySW.Close();
        }

        private static void WriteList(string message)
        {
            var mySW = new StreamWriter(
                @Path.Combine(Directory.GetCurrentDirectory(), Common.ERROR_FILE_LIST)
                , false
                , System.Text.Encoding.GetEncoding("shift_jis"));
            mySW.WriteLine(message);
            mySW.Close();
        }

        private static bool IRepoFileExists(string dirname, string filename)
        {
            var files = Directory.GetFiles(dirname, filename.Replace(".csv", "*.csv"));
            if (files.Length == 0) { return false; }
            return true;
        }

        private static void ActualProductChecker(Common cmn, List<string> irepolist, Target tg)
        {
            var srcFolder = @cmn.FsCd[Common.CONNECT_SOURCE].RootPath + "\\" + tg.Folder;
            var backupFolder = @cmn.FsCd[Common.CONNECT_DEST].RootPath + "\\" + Common.DI_BACKUP;
            var errorFolder = @cmn.FsCd[Common.CONNECT_DEST].RootPath + "\\" + Common.DI_ERROR;

            // SOURCEディレクトリ直下の当日フォルダをフィルター付きで調査
            //Console.WriteLine("[" + srcFolder + "] ディレクトリ内のファイルを調査します");
            var message = "[" + srcFolder + "] ディレクトリ内のファイルを調査します\n";
            var dis = new DirectoryInfo(srcFolder);
            var fiAlls = dis.GetFiles(cmn.FsCd[Common.CONNECT_SOURCE].FileFilter);
            foreach (FileInfo fis in fiAlls)
            {
                tg.CntIn++;
                if (cmn.Cnn == null || irepolist.Contains(fis.Name.Split('_')[0]))
                {
                    tg.CntTg++;
                    message += "  " + fis.Name;
                    if (IRepoFileExists(backupFolder, fis.Name))
                    {
                        tg.CntOk++;                             // backupフォルダに存在
                        message += " -- > ok\n";
                    }
                    else
                    {
                        if (IRepoFileExists(errorFolder, fis.Name))
                        {
                            tg.CntNg++;                         // errorフォルダに存在
                            tg.AttachFiles.Add(fis.FullName);
                            message += " 登録エラー\n";
                        }
                        else
                        {
                            tg.CntNothing++;                    // 転送失敗
                            tg.AttachFiles.Add(fis.FullName);
                            message += " 転送失敗！ファイルが見つかりませんでした．\n";
                        }

                    }
                }
            }
            if (tg.CntOk + tg.CntNg == 0)
            {
                tg.CntNothing = 0;
                tg.AttachFiles = null;
            }
            else
            {
                Console.WriteLine(message);
            }
        }
    }
}