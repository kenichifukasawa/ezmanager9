# EZManager9 プロジェクトマップ

**解析日:** 2026-07-01 (コードベース実解析による更新)
**プロジェクト名:** EZManager9 / CATVケーブルテレビ管理システム
**開発会社:** Plot Consultant Office (`VersionCompanyName` in CATV.VBP)
**対象組合:** 勝沼町CATV組合
**現行バージョン:** 9.1.26 (MajorVer=9 / MinorVer=1 / RevisionVer=26 in CATV.VBP)
**エントリーポイント:** `Sub Main` (Module1.bas:13667)

---

## 1. プロジェクト概要

Visual Basic 6 で構築されたレガシーデスクトップアプリケーション。ケーブルテレビ加入者の契約管理・課金管理・データ同期を担う総合管理システム。銀行FDデータ連携、NHK受信料、VoIP、ペイパービューなど複数サービスを包括的に管理する。

起動時にファイルサーバー (`\\k-catv-server\data\共有データ\ezmanager\`) から最新の catv.exe を MyDocuments にコピーして自動更新する仕組みを持つ (Module1.bas:13827)。

---

## 2. ディレクトリ構造

```
ezmanager9/                          # プロジェクトルート
├── CATV.VBP                         # メインプロジェクトファイル (v9.1.26)
├── CATV.VBW                         # IDE ウィンドウレイアウト
├── CATV.INI                         # アプリ設定ファイル (空)
├── catv.exe                         # メイン実行ファイル (4.2 MB)
├── oldcatv.exe                      # 旧バージョン保存用 (3.5 MB)
├── Module1.bas                      # コアビジネスロジック (15,458行)
├── clsCryptFile.cls                 # 暗号化クラス
├── *.frm / *.frx                    # UIフォーム群 (72本)
├── CATV1.ICO / CATV2.ICO           # アプリアイコン
├── settei.mdb                       # アプリ設定DB (ルートコピー)
├── ver.txt / sorita_ver.txt        # バージョン情報
│
├── mente/                           # メンテナンス サブプロジェクト
│   ├── ezmanager_ment.vbp
│   ├── frmmain.frm / frmmain.frx
│   ├── McWinFTP.cls                 # FTPクライアント実装
│   ├── FTPFile.cls                  # FTPファイル操作ラッパー
│   └── 勝沼町CATV組合/              # FTP転送用ステージング
│
├── ver/                             # バージョン管理 サブプロジェクト
│   ├── ezmanager_ver.vbp
│   ├── frmverup.frm
│   └── Module1.bas
│
├── sub/                             # サブコンポーネント
│   └── 勝沼町CATV組合/conbart/
│
├── 勝沼町CATV組合/                  # 組合データディレクトリ (本体) ※gitignore除外
│   ├── catv.mdb   (15.8 MB)        # メイン加入者・契約DB
│   ├── cl.mdb     (86.3 MB)        # 課金・顧客台帳DB
│   ├── koushin.mdb (238.4 MB)      # 更新・取引履歴DB (最大)
│   ├── nhk.mdb    (3.9 MB)         # NHK受信料DB
│   ├── voip.mdb   (1.8 MB)         # VoIPサービスDB
│   ├── ppv.mdb    (188 KB)         # ペイパービューDB
│   ├── out.mdb    (340 KB)         # 出力・エクスポート用DB
│   ├── conbart.mdb (332 KB)        # 契約・銀行口座マッピングDB
│   ├── settei.mdb (292 KB)         # アプリ設定DB (本体)
│   ├── FDBACK/                      # 銀行FDフィードバック (.dat)
│   ├── fdback2/
│   └── conbart/                     # 契約・銀行データCSV
│
└── 配布/                            # 配布パッケージ ※gitignore除外
    ├── 7.00/
    └── 7.15/
```

---

## 3. 技術スタック

以下は CATV.VBP の `Reference=` および `Object=` 行から確認した正確な一覧。

| 分類 | コンポーネント | 用途 |
|------|-------------|------|
| 開発言語 | Visual Basic 6 | |
| データベース | Microsoft Access (MDB) | Jet Engine |
| データアクセス | DAO360.DLL — Microsoft DAO 3.6 Object Library | |
| Officeライブラリ | MSACC.OLB — Microsoft Access 8.0 Object Library | |
| スクリプト | scrrun.dll — Microsoft Scripting Runtime | FileSystemObject使用 |
| OCX — グリッド | Vsflex7D.ocx (VSFlexGrid 7D) | メイングリッドコントロール |
| OCX — グリッド | MSFLXGRD.OCX (MSFlexGrid) | サブグリッド |
| OCX — タブ | TABCTL32.OCX | タブコントロール |
| OCX — コモン | MSCOMCTL.OCX (ListView/TreeView等) | |
| OCX — コモン2 | MSCOMCT2.OCX (DateTimePicker等) | |
| OCX — テキスト | Text60.ocx | テキスト拡張 |
| OCX — バーコード | BarCode.ocx | バーコード印刷 |
| OCX — 帳票 | ZoonFIO.ocx (v26.0) | 印刷・帳票出力 |
| OCX — ネット | MSINET.OCX | インターネット転送 |
| 暗号化 | clsCryptFile.cls | RSA実装 (自作) |
| FTP | mente/McWinFTP.cls | FTP通信 (自作) |
| ランタイム | msvbvm60.dll, dao360.dll | VB6ランタイム |
| OLE | STDOLE2.TLB | OLE Automation |

---

## 4. 実行ファイル一覧

| ファイル | サイズ | バージョン | 役割 |
|---------|--------|-----------|------|
| catv.exe | 4.2 MB | 9.1.26 | メインCATVアプリ (Startup="Sub Main") |
| ezmanager_ment.exe | 114 KB | 7.1.1 | データメンテ・FTP同期 |
| ezmanager_ver.exe | 28 KB | 7.1.1 | バージョン更新管理 |
| ezmanager_sub.exe | 327 KB | — | サブコンポーネント (別プロセス起動) |
| oldcatv.exe | 3.5 MB | — | 旧バージョン保存 |

---

## 5. ソースコード主要ファイル

### Module1.bas (15,458行) — コアモジュール

#### グローバル変数 (Module1.bas:123-222)

| 変数名 | 型 | 説明 |
|--------|-----|------|
| `db` | Database | catv.mdb 接続 |
| `db2` | Database | settei.mdb 接続 |
| `db3` | Database | 第3DB接続 |
| `db4` | Database | 第4DB接続 |
| `db_cl` | Database | cl.mdb 接続 |
| `db_voip` | Database | voip.mdb 接続 |
| `db_ppv` | Database | ppv.mdb 接続 |
| `db_out` | Database | out.mdb 接続 |
| `db_convart` | Database | conbart.mdb 接続 |
| `VBNAME` | String | catv.mdb パス |
| `VBNAME2` | String | settei.mdb パス |
| `vbname_cl` | String | cl.mdb パス |
| `vbname_voip` | String | voip.mdb パス |
| `vbname_out` | String | out.mdb パス |
| `vbname_convart` | String | conbart.mdb パス |
| `vbname_ppv` | String | ppv.mdb パス |
| `vbname3` | String | settei.mdb パス (別参照) |
| `rs` | Recordset | 汎用レコードセット |
| `cellrs` | Recordset | セルレコードセット |
| `sql`, `sql2` | String | SQL文バッファ |
| `bankidtbl()~4` | Array | 銀行IDテーブル |
| `bankcount~4` | Long | 銀行件数 |
| `shuruiidtbl()` | Array | サービス種別IDテーブル |
| `shuruicount` | Integer | サービス種別件数 |
| `sonotaidtbl()` | Array | その他サービスIDテーブル |
| `sonotacount` | Integer | その他サービス件数 |
| `koushindata()` | Long Array | 更新データ |
| `csv_data_keiyaku()` | Array | 契約CSVデータ |
| `csv_data_cl()` | Array | cl.mdb CSVデータ |
| `kidoupassword` | String | DB接続パスワード ("kamifusafusa") |
| `kiidou_mode` | Integer | 起動モード (0=通常, 1=モード1, 2=モード2) |
| `sopa` | String | データディレクトリパス |
| `mapa` | String | アプリディレクトリパス |
| `map_exe_path` | String | ezmanager_map.exe パス |
| `ftp_exe_path` | String | ezmanager_ftp.exe パス |
| `sub_exe_path` | String | ezmanager_sub.exe パス |
| `sagyou_log_path` | String | 作業ログパス (デスクトップ\EZManager_data\log\) |
| `s_kidou_kojinid` | String | 起動時指定個人ID (コマンドライン "k" prefix) |
| `nhk_ryoukin_tsuujou` | Long | NHK通常料金 |
| `nhk_ryoukin_hangaku` | Long | NHK半額料金 |

#### 主要関数・サブルーチン一覧 (Module1.bas より grep 確認)

**起動・DB接続系**
| 関数名 | 行 | 説明 |
|--------|-----|------|
| `Sub Main()` | 13667 | エントリーポイント。認証→スプラッシュ→バージョン確認→DB接続 |
| `Sub kidou_main()` | 15304 | catv.mdb を db に接続 |
| `Function kidou2()` | 15313 | settei.mdb を db2 に接続 |
| `Function kidou_cl()` | 344 | cl.mdb を db_cl に接続 |
| `Function kidou_voip()` | 15330 | voip.mdb を db_voip に接続 |
| `Function kidou_out()` | 15346 | out.mdb を db_out に接続 |
| `Function kidou_conbart()` | 15362 | conbart.mdb を db_convart に接続 |
| `Function kidou_ppv()` | 15377 | ppv.mdb を db_ppv に接続 |
| `Function kidou4()` | 15395 | 第4DB接続 |

**SQL・データアクセス系**
| 関数名 | 行 | 説明 |
|--------|-----|------|
| `Function FcSQlGet()` | 12460 | catv.mdb 汎用レコードセット取得 |
| `Function FcSQlGet2()` | 12531 | settei.mdb クエリ |
| `Function FcSQlGet3()` | 12556 | 第3DB クエリ |
| `Function FcSQlGet4()` | 12581 | 第4DB クエリ |
| `Function FcSQlGet_cl()` | 318 | cl.mdb クエリ |
| `Function FcSQlGet_voip()` | 12606 | voip.mdb クエリ |
| `Function FcSQlGet_out()` | 12631 | out.mdb クエリ |
| `Function FcSQlGet_ppv()` | 12656 | ppv.mdb クエリ |
| `Function soukatsu()` | 15181 | settei テーブル読み書き (koumoku / t1-t20) |
| `Function setting()` | 14211 | 設定値読み書きラッパー |

**契約・課金処理系**
| 関数名 | 行 | 説明 |
|--------|-----|------|
| `Function a_jidoukoushin()` | 7254 | 自動更新処理 (月次バッチ相当) |
| `Function a_genkin_nyuukin()` | 8486 | 現金入金処理 |
| `Function a_genkin_nyuukin2()` | 8330 | 現金入金処理 (バリアント2) |
| `Function a_genkin_nyuukin3()` | 8167 | 現金入金処理 (バリアント3) |
| `Function a_fdd_koushin2()` | 8722 | 銀行FDフィードバック更新 |
| `Function a_shukeiyaku_henkou()` | 10437 | 主契約変更 |
| `Function a_shukeiyaku_sakujo()` | 10276 | 主契約削除 |
| `Function a_shukeiyaku_koushin()` | 10551 | 主契約更新 |
| `Function a_keiyaku_henkou()` | 10362 | 契約変更 |
| `Function a_sonota_touroku()` | 10185 | その他サービス登録 |
| `Function a_sonota_henkou()` | 10106 | その他サービス変更 |
| `Function a_sonota_sakujo()` | 10037 | その他サービス削除 |
| `Function cl_koushin()` | 4155 | cl.mdb 更新 |
| `Function voip_koushin()` | 6288 | VoIP更新 |
| `Function ppv_koushin()` | 6053 | PPV更新 |
| `Function shousaiset()` | 9704 | 支払詳細設定 |
| `Function shousai_main_set()` | 9643 | 支払詳細メイン設定 |
| `Function shukeiyaku_touroku()` | 4084 | 主契約登録 |
| `Function sonotakeiyaku_touroku()` | 4026 | その他サービス契約登録 |
| `Function tainou_tesuuryou_touroku()` | 9515 | 滞納手数料登録 |
| `Function yoteikanri_touroku()` | 424 | 予定管理登録 |
| `Function yotejikkou()` | 3783 | 予定実行 |
| `Function yotejikkou2()` | 3528 | 予定実行 (バリアント2) |

**NHK受信料系**
| 関数名 | 行 | 説明 |
|--------|-----|------|
| `Function a_nhk_touroku()` | 9967 | NHK登録 |
| `Function a_nhk_henkou()` | 9889 | NHK変更 |
| `Function nhk_set()` | 11823 | NHK設定 |
| `Sub nhk_csv_fdd_check()` | 4575 | NHK FDチェック |
| `Sub nhk_csv_fdd_check2()` | 4728 | NHK FDチェック2 |
| `Sub nhk_csv_fdd_check3()` | 4207 | NHK FDチェック3 |
| `Sub nhk_csv_fdd_check4()` | 4459 | NHK FDチェック4 |
| `Sub nhk_zougaku()` | 13402 | NHK増額処理 |
| `Function nhk_gengaku()` | 13539 | NHK減額計算 |

**CSV出力系**
| 関数名 | 行 | 説明 |
|--------|-----|------|
| `Sub csv_csv2()` | 5274 | CSV出力2 |
| `Sub csv_csv3()` | 4979 | CSV出力3 |
| `Sub csv_csv4()` | 4829 | CSV出力4 |
| `Sub csv_csv5()` | 5905 | CSV出力5 (日付指定) |
| `Sub csv_csv6()` | 5752 | CSV出力6 (日付指定) |
| `Sub csv_csv7()` | 5596 | CSV出力7 (日付指定) |
| `Sub csv_csv8()` | 4910 | CSV出力8 |
| `Function csv_csv_cl()` | 5072 | cl.mdb CSV出力 |

**cl.mdbデータ処理系**
| 関数名 | 行 | 説明 |
|--------|-----|------|
| `Sub cl_data_bunkatsu()` | 1352 | cl.mdbデータ分割取込 |
| `Sub cl_data_bunkatsu2()` | 1655 | cl.mdbデータ分割取込2 |
| `Sub cl_data_convert()` | 1984 | cl.mdbデータ変換 |
| `Sub cl_data_convert2()` | 2921 | cl.mdbデータ変換2 |
| `Sub cl_tools_chk()` | 2714 | cl.mdbツールチェック |
| `Sub data_kyoutou_to_katsunuma()` | 755 | 共同→勝沼データ同期 |

**ログ・エラー系**
| 関数名 | 行 | 説明 |
|--------|-----|------|
| `Sub logset()` | 11246 | 操作ログ設定 |
| `Sub logset2()` | 10851 | 操作ログ設定2 |
| `Sub log_sagyou()` | 4957 | 作業ログファイル書き込み |
| `Sub err_write()` | 6463 | エラーDB記録 |
| `Sub err_write2()` | 6477 | エラーファイル記録 |
| `Sub chk_log()` | 10833 | ログチェック |

**UI補助系**
| 関数名 | 行 | 説明 |
|--------|-----|------|
| `Sub keiyakuset()` | 11658 | 契約情報UI設定 |
| `Sub keiyakuset2()` | 11747 | 契約情報UI設定2 |
| `Sub nyuukinrirekiset()` | 14950 | 入金履歴UI設定 |
| `Sub nyuukinrirekiset2()` | 14614 | 入金履歴UI設定2 |
| `Sub nyuukin_shousai_set()` | 15047 | 入金詳細UI設定 |
| `Function hyouji_touroku()` | 11913 | 表示登録 |
| `Sub sbbankset()` | 12962 | 銀行コンボ設定 |
| `Sub sbnbankset2()~6()` | 12997 | 銀行コンボ各種設定 |
| `Sub sbshuruiset()` | 12752 | サービス種別コンボ設定 |
| `Sub sbsonotaset()` | 12782 | その他サービスコンボ設定 |
| `Sub bankset()` | 13298 | 銀行情報設定 |
| `Sub cellset()` | 13358 | セル情報設定 |
| `Sub sonotaset()` | 13173 | その他サービス設定 |
| `Sub stb_grid_set()` | 6103 | STBグリッド設定 |
| `Sub ppv_grid_set()` | 6174 | PPVグリッド設定 |
| `Sub cl_grid_set()` | 289 | cl.mdbグリッド設定 |
| `Sub frmtourokuclear()` | 12682 | 登録フォームクリア |

**システム系**
| 関数名 | 行 | 説明 |
|--------|-----|------|
| `Sub systemend()` | 6413 | システム終了処理 |
| `Sub kakunin_ver()` | 14302 | バージョン確認・更新 |
| `Function koushinchk()` | 723 | 更新チェック |
| `Function rocked()` | 741 | ロック制御 |
| `Sub data_check2()` | 6530 | データチェック |
| `Sub kyousei_shuusei()` | 6490 | 強制修正 |
| `Function flgup()` | 10674 | フラグ更新 |
| `Function flgup2()` | 10642 | フラグ更新2 |
| `Sub data_tsuia()` | 9606 | データ追加 |
| `Sub add_fld_200403()` | 14430 | フィールド追加 (v2004.03) |
| `Function newtbl_make()` | 14580 | テーブル新規作成 |
| `Sub errorfd()` | 15413 | エラーFD処理 |

### clsCryptFile.cls — 暗号化クラス
RSAアルゴリズム実装によるファイル暗号化・復号化。セキュリティキー管理。

### mente/McWinFTP.cls — FTPクライアント
### mente/FTPFile.cls — FTPファイル操作ラッパー

---

## 6. UIフォーム一覧 (72本)

VBP に登録された順序 (CATV.VBP:15-79)。

### メイン・ナビゲーション
| フォーム | 役割 |
|---------|------|
| frmmain.frm | メインダッシュボード。12個のフレームでグループ化されたピクチャボタンによるナビゲーション |
| frmrogo.frm | スプラッシュスクリーン (起動時進捗表示) |
| FRMINFO.frm | アプリ情報・バージョン表示 |
| FRMPASSWD.frm | 起動時パスワード認証 |
| FRMSETTEI.frm | 設定画面 (settei.mdb 管理) |
| frmomachi.frm | 待機・処理中表示 |

### 契約管理
| フォーム | 役割 |
|---------|------|
| frmtouroku.frm | 新規加入者登録 |
| frmtouroku2.frm | 加入者登録 (バリアント2) |
| frmtouroku3.frm | 加入者登録 (バリアント3) |
| frmtourokusentaku.frm | 登録対象選択 |
| frmtourokusentaku2.frm | 登録対象選択2 |
| frmkeiyaku.frm | 契約詳細 |
| frmkeiyaku2.frm | 契約詳細2 |
| frmkeiyaku3.frm | 契約詳細3 |
| frmkeiyaku4.frm | 契約詳細4 |
| frmkeiyaku5.frm | 契約詳細5 |
| frmkeiyaku6.frm | 契約詳細6 |
| frmkeiyakurireki.frm | 契約履歴 |
| frmkeiyakurireki2.frm | 契約履歴2 |

### 課金・支払
| フォーム | 役割 |
|---------|------|
| frmnyuukin.frm | 入金処理 |
| frmkoushin.frm | 更新・課金処理メイン |
| frmkoushin2.frm | 更新・課金処理2 |
| frmbank.frm | 銀行口座管理 |
| frmbank2.frm | 銀行口座管理2 |
| frmkekka.frm | 結果表示 |
| frmkekkaichi.frm | 結果一覧 |
| frmfurikou.frm | 振替処理 |
| frmkari.frm | 仮登録 |

### 検索・選択
| フォーム | 役割 |
|---------|------|
| frmkensaku.frm | 検索インターフェース |
| frmsentaku.frm | 選択・フィルタリング |
| frmjoukyousentaku.frm | 状況別選択 |
| frmrirekisentaku.frm | 履歴別選択 |

### 照会・表示
| フォーム | 役割 |
|---------|------|
| frmjoukyou.frm | 状況表示 |
| frmjoukyou2.frm | 状況表示2 |
| frmsonota.frm | その他サービス情報 |
| frmsonota2.frm | その他サービス情報2 |
| frmkojinshousai.frm | 個人詳細 (cl.mdbグリッド表示含む) |
| frmlogshousai.frm | ログ詳細 |
| frmrireki.frm | 履歴記録 |
| frmyotei.frm | 予定管理 |
| frmshukei.frm | 集計・サマリー |

### データ入出力
| フォーム | 役割 |
|---------|------|
| frmfdd.frm | FDフォーマットデータ処理 |
| frmsouken.frm | 総検・照会 |
| frmsouken2.frm | 総検2 |
| frmsouken3.frm | 総検3 |
| frmsouken4.frm | 総検4 |
| frmsouken5.frm | 総検5 |
| frmsouken6.frm | 総検6 |
| frmsouken7.frm | 総検7 |
| frmcell.frm | セル個別レコード |
| frmcell2.frm | セル個別レコード2 |
| frmja.frm | JAデータ処理 |
| frmdatacheck.frm | データ検証 |
| frmdatacheck2.frm | データ検証2 |

### 特殊サービス
| フォーム | 役割 |
|---------|------|
| frmnhk.frm | NHK受信料管理 |
| frmnhk2.frm | NHK受信料管理2 |
| frmnhkshuukei.frm | NHK集計 |
| frmvoip.frm | VoIPサービス管理 |
| frmstb.frm | STB (セットトップボックス) 管理 |
| frmsaitekika.frm | 最適化・マッチング処理 |

### 設定・管理
| フォーム | 役割 |
|---------|------|
| frmdatapath.frm | データパス設定 |
| frmdir.frm | ディレクトリ選択 |
| frmkanrisha.frm | 管理者管理 |
| frmhiduke.frm | 日付選択ダイアログ |

### 削除
| フォーム | 役割 |
|---------|------|
| frmsakujo.frm | 削除処理 |
| frmsakujo2.frm | 削除処理2 |

### cl.mdb専用
| フォーム | 役割 |
|---------|------|
| frmcl.frm | cl.mdb メイン表示 |
| frmcl2.frm | cl.mdb サブ表示 |

### ver/サブプロジェクト
| フォーム | 役割 |
|---------|------|
| ver/frmverup.frm | バージョンアップ処理 |

---

## 7. メインメニュー構造 (frmmain.frm から確認)

frmmain は 12個の Frame (Frame1-Frame12) でグループ分けされた PictureBox ボタンで構成。以下は Picture1_Click ～ Picture17_Click のコードから確認した遷移。

| ボタン | 遷移先フォーム | 機能 |
|--------|--------------|------|
| Picture1 | frmkeiyaku | 契約管理 |
| Picture2 | frmbank | 銀行口座管理 |
| Picture3 | frmsonota | その他サービス一覧 |
| Picture4 | frmkensaku | 検索 |
| Picture5 | (終了処理) | バックアップ後アプリ終了 |
| Picture6 | frmsaitekika | 最適化 (kiidou_mode=0 必須) |
| Picture7 | FRMSETTEI | 設定 |
| Picture8 | FRMINFO | アプリ情報 |
| Picture9 | frmfdd | FDデータ処理 (kiidou_mode=0 必須) |
| Picture10 | frmkoushin | 更新・課金 (kiidou_mode=0 必須) |
| Picture11 | frmjoukyousentaku | 状況選択 |
| Picture12 | frmkeiyakurireki または ezmanager_sub.exe | 履歴 (chknew.Value で分岐) |
| Picture13 | ezmanager_sub.exe " s" | サブEXE起動 |
| Picture14 | frmsentaku | 選択 |
| Picture15 | frmkoushin2 | 更新2 (kiidou_mode=0 必須) |
| Picture16 | frmtouroku | 新規加入者登録 (kiidou_mode=0 必須) |
| Picture17 | frmyotei | 予定管理 |

`kiidou_mode` が 1 または 2 の場合、一部ボタンはロックされる (起動モード制限)。

---

## 8. データベーススキーマ概要

### catv.mdb — メインDB (db 変数) パスワード: "kamifusafusa"

SQLから確認したテーブル:

| テーブル名 | 主なカラム | 用途 |
|-----------|-----------|------|
| `kojin` | kojinid, name, bankid, f1, f3, f4, sonotakin, jikaikin, jikaishu, jikaibi, seikyuu | 加入者マスタ |
| `keiyaku` | keiyakuid, kojinid, shurui, kingaku, kadou, sadou, shime, kuchisuu | メインサービス契約 |
| `keiyaku2` | keiyakuid, kojinid, sonotaid | その他サービス契約 |
| `shurui` | shuruiid, shuruimei | サービス種別マスタ |
| `sonota` | sonotaid, sonota | その他サービスマスタ |
| `bank` | bankid, bankcode, tel | 銀行口座 |
| `bankmain` | bankcode | 銀行マスタ |
| `cell` | cellid | セル (引落区分?) |
| `log` | kojinid, kekka | 操作ログ |
| `yotei` | yoteiid, kojinid, shurui, kanrenid, kanrenmei, kingaku, kuchisuu, sousuu, itsukara, itsumade, itsu, nanji, adoff, bikou, st | 予定管理 |
| `koushinkiroku` | koushnid | 更新記録 |
| `nhk_m` | kojinid | NHK受信料マスタ |
| `nhk_k` | — | NHK集計 |
| `voip_m` | st_st | VoIPサービスマスタ |
| `settei` | koumoku, t1-t20 | システム設定 (キー=値形式、値は最大20カラム) |
| `renban` | renban | 連番管理 |
| `seikyuu` | — | 請求 |

### settei テーブルの主要キー (soukatsu() / setting() から)

| koumoku 値 | 参照フィールド | 用途 |
|-----------|--------------|------|
| `"startup"` | t3 | バックアップ先パス |
| `"4"` | t4 | データディレクトリパス (sopa) |
| `"7"` | t7 | 起動モード (kiidou_mode) |
| `"renban"` | t16 | 連番値 |

### cl.mdb — 課金台帳DB (db_cl 変数) パスワード: "kamifusafusa"

| テーブル名 | 主なカラム | 用途 |
|-----------|-----------|------|
| `cltbl` | nentsuki, kojinid | 課金台帳 (年月別) |
| `shousai` | keist, keiseikyuu, keinyu, bankid | 支払詳細 |

### その他DB

| DB | 主なテーブル | 用途 |
|----|------------|------|
| nhk.mdb | nhk_m, nhk_k | NHK受信料データ |
| voip.mdb | voip_m (st_st) | VoIPサービスデータ |
| ppv.mdb | keiyaku2 | ペイパービュー契約 |
| out.mdb | voip_m | 出力・エクスポート用ステージング |
| conbart.mdb | — | 契約・銀行口座マッピング |
| settei.mdb | settei | アプリ設定 (db2 / db3 で参照) |

---

## 9. データフロー

```
[起動]
  └→ コマンドライン解析 ("k" + 5桁kojinid で直接加入者起動)
  └→ Windows Build チェック (>28500 で警告)
  └→ FRMPASSWD (パスワード認証)
  └→ frmrogo (スプラッシュ)
  └→ バージョン確認 (kakunin_ver)
  └→ サーバーから catv.exe 自動更新 (\\k-catv-server\data\共有データ\ezmanager\)
  └→ settei.mdb → データパス読込 → catv.mdb / cl.mdb / voip.mdb 等を接続

[通常処理]
  ユーザー操作 (frmmain → 各フォーム)
       ↓
  Module1.bas のビジネスロジック関数
       ↓
  DAO で catv.mdb / cl.mdb / 各DB 更新
       ↓
  ログ記録 (log テーブル / 作業ログファイル)

[データ連携]
  銀行FDデータ (.dat) → FDBACK/ → nhk_csv_fdd_check() / a_fdd_koushin2() → catv.mdb 更新
  catv.mdb → csv_csv*() → CSV出力 → 銀行・外部システム
  cl.mdb → cl_data_convert() → cltbl 集計処理

[終了]
  Picture5_Click → settei.mdb バックアップ → systemend() → アプリ終了
```

---

## 10. 主な機能

- **契約管理:** 新規加入登録・契約変更・解約・履歴管理
- **課金・支払:** 入金記録・銀行口座管理・自動引落・残高管理・滞納手数料
- **データ同期:** FTPベースのデータ同期・銀行FDフィードバック処理
- **特殊サービス:** NHK受信料 (増額/減額/集計)・VoIP・ペイパービュー・STB管理
- **照会・帳票:** 多条件検索・状況別照会・ZoonFIO.ocx による帳票印刷・バーコード出力
- **システム管理:** パスワード認証・管理者設定・kiidou_mode による権限制御
- **外部連携:** 銀行FDデータ交換 (FDBACK/*.dat)・CSV インポート/エクスポート・FTP同期
- **自動更新:** 起動時にファイルサーバーから最新EXEを自動コピー
- **予定管理:** yotei テーブルによる将来の契約変更予約・自動実行

---

## 11. バージョン履歴概要

| バージョン | 時期 | 内容 |
|-----------|------|------|
| 7.00 | 2013年4月 | 配布版 (配布/7.00/) |
| 7.15 | 2013年4月 | アップデート版 (配布/7.15/) |
| 7.1.1 | 2022〜2023年 | mente/ver サブプロジェクト現行版 |
| 9.1.26 | 2026-05-10 | 現行メインアプリ (catv.exe タイムスタンプ) |

---

## 12. アーキテクチャ特性

- **スタイル:** モノリシックデスクトップアプリ + 複数補助実行ファイル (ezmanager_sub.exe, ezmanager_ment.exe, ezmanager_ver.exe)
- **データ:** 複数の Access MDB ファイルに機能別分散 (catv/cl/nhk/voip/ppv/out/conbart/settei)
- **ビジネスロジック:** クライアントサイドの Module1.bas (15,458行) に一元集中
- **セキュリティ:** パスワード認証 (FRMPASSWD) + MDB パスワード ("kamifusafusa") + ファイル暗号化 (clsCryptFile.cls)
- **権限制御:** kiidou_mode (0=フル, 1/2=制限) により特定ボタンをロック
- **通信:** FTP (McWinFTP.cls) + UNCパス (\\k-catv-server\) による LAN 内ファイル共有
- **自動更新:** 起動時サーバーから EXE をコピー (Module1.bas:13827)
- **コマンドライン:** `catv.exe k12345` 形式で特定加入者を直接起動可能
- **言語:** UI・データ・ドキュメントすべて日本語 (Shift-JIS エンコード)
- **エントリーポイント:** `Sub Main` (標準モジュール起動、フォームは Startup="Sub Main" で管理)
