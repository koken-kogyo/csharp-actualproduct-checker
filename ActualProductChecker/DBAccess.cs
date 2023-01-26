using Oracle.ManagedDataAccess.Client;
using System;
using System.Data;
using System.Diagnostics;

namespace ActualProductChecker
{
    /// <summary>
    /// データベース アクセス クラス
    /// </summary>
    class DBAccess
    {
        /// <summary>
        /// Oracle データベース スキーマへの接続
        /// </summary>
        /// <param name="cmn">共通クラス</param>
        /// <param name="cnn">Oracle データベースへの接続クラス</param>
        public bool ConnectSchema(Common cmn, ref OracleConnection cnn)
        {
            string user = cmn.DbCd.User; // ユーザー
            string passwd = cmn.DbCd.Passwd;  // パスワード
                                              //string ds = "oracle"; // データソース
            string ds = "(DESCRIPTION="
                        + "(ADDRESS="
                          + "(PROTOCOL=" + cmn.DbCd.Protocol + ")"
                          + "(HOST=" + cmn.DbCd.Host + ")"
                          + "(PORT=" + cmn.DbCd.Port + ")"
                        + ")"
                        + "(CONNECT_DATA="
                          + "(SERVICE_NAME=" + cmn.DbCd.ServiceName + ")"
                        + ")"
                      + ")";  // Oracle Client を使用せず直接接続する

            // Oracle 接続文字列を組み立てる
            string connectstring = "User Id=" + user + "; "
                                 + "Password=" + passwd + "; "
                                 + "Data Source=" + ds;

            try
            {
                cnn = new OracleConnection(connectstring);

                // Oracle へのコネクションの確立
                cnn.Open();
                Console.Write("Connected to Oracle ");
                Console.WriteLine(cnn.ServerVersion);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Source);
                Console.WriteLine(e.Message + "\n");
                CloseSchema(cmn.Cnn); // 接続を閉じる
                return false;
            }
            return true;
        }

        /// <summary>
        /// Oracle データベース スキーマからの切断
        /// </summary>
        /// <param name="cnn">Oracle データベースへの接続クラス</param>
        public void CloseSchema(OracleConnection cnn)
        {
            // 接続を閉じる
            if (cnn != null)
            {
                cnn.Close();
            }
        }

        /// <summary>
        /// i-Reporter 対象工程取得 (KM1060 i-Reporter 対象工程マスタ)
        /// </summary>
        /// <param name="_cmn">共通クラス</param>
        /// <param name="_active">有効フラグ</param>
        /// <param name="_ktNm">i-Reporter工程名称</param>
        /// <param name="_opeStat">運用状態フラグ</param>
        /// <param name="_dataSet">データセット</param>
        /// <returns>結果 (0: 正常, -1: 異常)</returns>
        public bool GetIRepoKt(Common _cmn, string _active, string _opeStat, ref DataSet _dataSet)
        {
            try
            {
                string sql = "SELECT "
                           + "SEQ, "
                           + "ACTIVE, "
                           + "IREPOKTCD, "
                           + "IREPOKTNM, "
                           + "OPESTAT "
                           + "FROM "
                           + _cmn.DbCd.Schema + "." + Common.TABLE_NAME_KM1060 + " "
                           + "WHERE "
                           + "ACTIVE = '" + _active + "' AND "
                           + "OPESTAT = '" + _opeStat + "'";

                using (OracleCommand myCmd = new OracleCommand(sql, _cmn.Cnn))
                {
                    using (OracleDataAdapter myDa = new OracleDataAdapter(myCmd))
                    {
                        // 結果取得
                        myDa.Fill(_dataSet, Common.TABLE_NAME_KM1060);
                        if (_dataSet.Tables[Common.TABLE_NAME_KM1060].Rows.Count == 0) // 該当データなし
                        {
                            Console.WriteLine("i-Reporter 対象工程がデータベースに登録されていません．");
                            return false;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine("Exception Source = " + e.Source);
                Debug.WriteLine("Exception Message = " + e.Message);
                CloseSchema(_cmn.Cnn); // 接続を閉じる
                return false;
            }
            return true;
        }

        /// <summary>
        /// 送信先メールアドレス取得 (KM0040 アドレス帳マスタ)
        /// </summary>
        /// <param name="_cmn">共通クラス</param>
        /// <param name="_adrtype">宛先種別 (1:To 2:Cc 3:Bcc)</param>
        /// <param name="_dataSet">データセット</param>
        /// <returns>結果 (true: 正常, false: 異常)</returns>
        public bool GetEmailAddress(Common _cmn, string _adrtype, ref DataSet _dataSet)
        {
            try
            {
                string dtName = _adrtype == Common.EMAIL_ADRTYPE_TO ? "To" :
                                _adrtype == Common.EMAIL_ADRTYPE_CC ? "Cc" :
                                _adrtype == Common.EMAIL_ADRTYPE_BCC ? "Bcc" : "不明";
                string sql = "SELECT "
                           + "DESTNM||' '||HONOR||' <'||EMAIL||'>' AS EMAIL "
                           + "FROM "
                           + _cmn.DbCd.Schema + "." + Common.TABLE_NAME_KM0040 + " "
                           + "WHERE "
                           + "PGMID = '" + Common.PGM_ID + "' AND "
                           + "ACTIVE = '1' AND "
                           + "ADRTYPE = '" + _adrtype + "' AND "
                           + "SYSDATE BETWEEN VALDTF AND VALDTT";

                using (OracleCommand myCmd = new OracleCommand(sql, _cmn.Cnn))
                {
                    using (OracleDataAdapter myDa = new OracleDataAdapter(myCmd))
                    {
                        // 結果取得
                        myDa.Fill(_dataSet, dtName);

                        if (_adrtype == Common.EMAIL_ADRTYPE_TO &&
                            _dataSet.Tables[dtName].Rows.Count == 0) // 該当データなし
                        {
                            Console.WriteLine("送信先メールアドレスがデータベースに登録されていません．");
                            return false;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Source);
                Console.WriteLine(e.Message);
                CloseSchema(_cmn.Cnn); // 接続を閉じる
                return false;
            }
            return true;
        }


    }
}
