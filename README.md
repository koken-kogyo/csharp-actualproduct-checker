# ActualProductChecker  

- [KMC006SC] 製造実績ファイル照合  


## 概要  

- ireposv サーバーの 入力完了帳票 と kemsvr2 サーバーの 自動受入 されたファイルを比較照合。  
- 比較照合結果をメールにて配信する。  
- 配信の際、kemsvr2 サーバーの 自動受入ログファイル と 受入エラーとなったファイルを添付。  


## 開発環境  

- Visual Studio Professional 2022 (Visual C# 2022)  


## アプリケーションの種類  

- コンソール アプリ (.NET Framework) C#  


## メンバー  

- y.watanabae  


## プロジェクト構成  

~~~
./
│  .gitignore					# ソース管理除外対象
│  ActualProductChecker.sln			# プロジェクトファイル
│  README.md					# このファイル
│  ReleaseNote.txt				# リリース情報
│  
├─ ActualProductChecker
│  │  Common.cs 				# 共通 クラス
│  │  DBAccess.cs				# データベースアクセ スクラス
│  │  FileAccess.cs				# ファイルアクセス クラス
│  │  Program.cs				# コンソール アプリケーション本体
│  │          
│  └─Properties
│          AssemblyInfo.cs
│          
│          
├─ packages
│      DecryptPassword.dll			# パスワード復号化モジュール
│      Oracle.ManagedDataAccess.dll		# Oracle接続モジュール
│      
├─ settingfiles
│      ConfigDB - KOKEN_1.xml			# データベース設定ファイル
│      ConfigDB - KOKEN_5.xml			# データベース設定ファイル (テスト用)
│      ConfigFS - kemsvr2.xml			# サーバー設定ファイル
│      ConfigFS - Local.xml			# サーバー設定ファイル (テスト用)
│      ConfigFS - pc090n.xml			# サーバー設定ファイル (テスト用)
│      ConfigSMTP - gmail.xml			# SMTP設定ファイル (テスト用)
│      ConfigSMTP - Koken.xml			# SMTP設定ファイル
│      
└─ specification
        [KMC006SC] 製造実績ファイル照合 機能仕様書_Ver.1.0.0.0.xlsx
        
~~~


## データベース参照テーブル  

| Table    | Name                      |  
| :------: | :------------------------ |  
| KM0040   | アドレス帳マスタ          |  
| KM1060   | i-Reporter 対象工程マスタ |  


