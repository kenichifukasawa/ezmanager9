# EZManager9 プロジェクトマップ

**解析日:** 2026-06-13 (2026-06-21 更新)  
**プロジェクト名:** EZManager9 / CATVケーブルテレビ管理システム  
**開発会社:** プロット コンサルタント オフィス (Plot Consultant Office)  
**対象組合:** 勝沼町CATV組合  
**現行バージョン:** 9.1.26

---

## 1. プロジェクト概要

Visual Basic 6 で構築されたレガシーデスクトップアプリケーション。ケーブルテレビ加入者の契約管理・課金管理・データ同期を担う総合管理システム。銀行FDデータ連携、NHK受信料、VoIP、ペイパービューなど複数サービスを包括的に管理する。

---

## 2. ディレクトリ構造

```
ezmanager9/                          # プロジェクトルート (691 MB)
├── CATV.VBP                         # メインプロジェクトファイル
├── CATV.VBW                         # IDE ウィンドウレイアウト
├── CATV.INI                         # アプリ設定ファイル (空)
├── catv.exe                         # メイン実行ファイル (4.2 MB)
├── oldcatv.exe                      # 旧バージョン保存用 (3.5 MB)
├── Module1.bas                      # コアビジネスロジック (15,458行)
├── clsCryptFile.cls                 # 暗号化クラス
├── *.frm / *.frx                    # UIフォーム群 (72本)
├── CATV1.ICO / CATV2.ICO           # アプリアイコン
├── ver.txt / sorita_ver.txt        # バージョン情報
├── COMB.TXT                         # 加入者データエクスポート (626 KB)
├── KOUFURI.DAT                      # 振替・支払データ (390 KB)
├── IKKATUFD.dat                     # 一括FDデータ (330 KB)
│
├── mente/                           # メンテナンス サブプロジェクト
│   ├── ezmanager_ment.vbp           # メンテプロジェクト (v7.1.1)
│   ├── ezmanager_ment.exe           # メンテ実行ファイル (114 KB)
│   ├── catv.exe                     # 同梱コピー (メインEXE)
│   ├── settei.mdb                   # 設定DBの同梱コピー
│   ├── frmmain.frm                  # メインフォーム
│   ├── McWinFTP.cls                 # FTPクライアント実装
│   ├── FTPFile.cls                  # FTPファイル操作ラッパー
│   └── 勝沼町CATV組合/              # 組合データDB群の同梱コピー (FTP転送用と思われる)
│       ├── catv.mdb / cl.mdb / koushin.mdb / nhk.mdb / voip.mdb
│       ├── ppv.mdb / out.mdb / conbart.mdb / settei.mdb
│       ├── FDBACK/
│       └── conbart/
│
├── ver/                             # バージョン管理 サブプロジェクト
│   ├── ezmanager_ver.vbp            # バージョン管理プロジェクト (v7.1.1)
│   ├── ezmanager_ver.exe            # バージョン管理実行ファイル (28 KB)
│   ├── catv.exe                     # 同梱コピー (メインEXE)
│   └── Module1.bas                  # バージョン管理ロジック
│
├── sub/                             # サブコンポーネント
│   ├── ezmanager_sub.exe            # サブ実行ファイル (327 KB)
│   └── 勝沼町CATV組合/
│       └── conbart/                 # 契約・銀行データCSVのみ
│
├── 勝沼町CATV組合/                  # 組合データディレクトリ (本体)
│   ├── catv.mdb                     # メイン加入者・契約DB (15.8 MB)
│   ├── cl.mdb                       # 課金・顧客台帳DB (86.3 MB)
│   ├── koushin.mdb                  # 更新・取引履歴DB (238.4 MB)
│   ├── nhk.mdb                      # NHK受信料DB (3.9 MB)
│   ├── voip.mdb                     # VoIPサービスDB (1.8 MB, 2025-08-08更新)
│   ├── ppv.mdb                      # ペイパービューDB (188 KB)
│   ├── out.mdb                      # 出力・エクスポート用DB (340 KB)
│   ├── conbart.mdb                  # 契約・銀行口座マッピングDB (332 KB)
│   ├── settei.mdb                   # アプリ設定DB (292 KB)
│   ├── old_catv.mdb                 # catv.mdb 旧版バックアップ (9.7 MB, 2018年)
│   ├── old_koushin.mdb              # koushin.mdb 旧版バックアップ (85.2 MB, 2018年)
│   ├── *.ldb                        # Access ロックファイル (catv/cl/ppv)
│   ├── error.txt                    # エラーログ
│   ├── nhk_202204.csv               # NHKデータエクスポート
│   ├── FDBACK/                      # 銀行FDフィードバックファイル群 (.dat, 2006年〜)
│   ├── fdback2/                     # 銀行FDフィードバック群 第2世代 (.dat, 2019年〜)
│   └── conbart/                     # 契約・銀行データCSV
│       ├── ケーブルネット納入者.csv
│       ├── 加入者名簿.csv / 加入者名簿1.csv / 加入者名簿old.csv
│       ├── ｹｰﾌﾞﾙﾈｯﾄ引落.csv
│       └── ｹｰﾌﾞﾙﾈｯﾄ引落old.csv
│
├── 配布/                            # 配布パッケージ
│   ├── 7.00/                        # v7.00 インストーラ (2013年)
│   │   ├── setup.exe
│   │   ├── SETUP.LST
│   │   ├── catv.CAB
│   │   └── Support/                 # DLL・OCX群
│   └── 7.15/                        # v7.15 インストーラ (2013年)
│       ├── setup.exe
│       └── Support/
│
└── (ドキュメント類)
    ├── 勝沼CATV_データベース.txt
    ├── 団体一括 出力フォーマット.doc
    └── 契約状況_20111014-205528.csv  (契約状況エクスポート 1.2 MB)
```

---

## 3. 技術スタック

| 分類 | 内容 |
|------|------|
| 開発言語 | Visual Basic 6 |
| データベース | Microsoft Access (MDB / Jet Engine) |
| データアクセス | DAO 3.6 Object Library |
| OCXコントロール | TABCTL32.OCX, MSFLXGRD.OCX, Vsflex7D.ocx, BarCode.ocx, ZoonFIO.ocx, Text60.ocx |
| 共通コントロール | MSCOMCTL.OCX, MSCOMCT2.OCX |
| ネットワーク | MSINET.OCX, 独自FTPクライアント (McWinFTP.cls) |
| セキュリティ | clsCryptFile.cls (RSA暗号化実装) |
| ランタイム | msvbvm60.dll, dao360.dll, msjet40.dll |

---

## 4. 実行ファイル一覧

| ファイル | サイズ | バージョン | 役割 |
|---------|--------|-----------|------|
| catv.exe | 4.2 MB | 9.1.26 | メインCATVアプリ |
| ezmanager_ment.exe | 114 KB | 7.1.1 | データメンテ・FTP同期 |
| ezmanager_ver.exe | 28 KB | 7.1.1 | バージョン更新管理 |
| ezmanager_sub.exe | 327 KB | - | サブコンポーネント |
| oldcatv.exe | 3.5 MB | - | 旧バージョン保存 |

---

## 5. ソースコード主要ファイル

### Module1.bas (15,458行) — コアモジュール
- グローバルDB接続変数 (db, db2, db3, db4)
- 銀行ID・カテゴリID・顧客IDのグローバル配列
- Windows API 宣言
- ファイル操作・暗号化ユーティリティ関数
- FTP/ネットワーク通信ハンドラ
- パス設定・設定管理

### clsCryptFile.cls — 暗号化クラス
- ファイル暗号化・復号化
- RSAアルゴリズム実装
- セキュリティキー管理

### mente/McWinFTP.cls — FTPクライアント
### mente/FTPFile.cls — FTPファイル操作ラッパー

---

## 6. UIフォーム一覧 (72本)

### メイン
| フォーム | 役割 |
|---------|------|
| frmmain.frm | メインダッシュボード・ナビゲーション |
| frmrogo.frm | スプラッシュスクリーン |
| FRMINFO.frm | アプリ情報 |

### 契約管理
| フォーム | 役割 |
|---------|------|
| frmtouroku.frm / 2 / 3 | 加入者登録 |
| frmkeiyaku.frm / 2〜6 | 契約詳細・バリアント |
| frmkeiyakurireki.frm / 2 | 契約履歴 |

### 検索・選択
| フォーム | 役割 |
|---------|------|
| frmkensaku.frm | 検索インターフェース |
| frmsentaku.frm | 選択・フィルタリング |
| frmjoukyousentaku.frm | 状況別選択 |
| frmrirekisentaku.frm | 履歴別選択 |
| frmtourokusentaku.frm / 2 | 登録選択 |

### 課金・支払
| フォーム | 役割 |
|---------|------|
| frmnyuukin.frm | 入金処理 |
| frmkoushin.frm / 2 | 更新・変更管理 |
| frmbank.frm / 2 | 銀行口座管理 |
| frmkekka.frm / frmkekkaichi.frm | 結果表示 |

### 照会・表示
| フォーム | 役割 |
|---------|------|
| frmjoukyou.frm | 状況表示 |
| frmsonota.frm / 2 | その他情報 |
| frmkojinshousai.frm | 個人詳細 |
| frmlogshousai.frm | ログ詳細 |

### 設定・管理
| フォーム | 役割 |
|---------|------|
| FRMSETTEI.frm | 設定画面 |
| frmdir.frm | ディレクトリ選択 |
| frmdatapath.frm | データパス設定 |
| FRMPASSWD.frm | パスワード管理 |
| frmkanrisha.frm | 管理者管理 |

### データ入出力・集計
| フォーム | 役割 |
|---------|------|
| frmfdd.frm | FDフォーマットデータ |
| frmsouken.frm / 2〜7 | 総検・照会バリアント |
| frmcell.frm / 2 | 個別レコード |
| frmdatacheck2.frm | データ検証 |
| frmja.frm | JAデータ |
| frmshukei.frm | 集計・サマリー |

### 特殊サービス
| フォーム | 役割 |
|---------|------|
| frmnhk.frm / 2 | NHK受信料管理 |
| frmvoip.frm | VoIPサービス管理 |
| frmsaitekika.frm | 最適化・マッチング |

### ユーティリティ
| フォーム | 役割 |
|---------|------|
| frmsakujo.frm / 2 | 削除処理 |
| frmkari.frm | 仮登録 |
| frmomachi.frm | 待機・保留 |
| frmhiduke.frm | 日付選択 |
| frmfurikou.frm | 振替処理 |
| frmyotei.frm | 予定管理 |
| frmrireki.frm | 履歴記録 |
| frmverup.frm (ver/) | バージョンアップ |

---

## 7. データベーススキーマ概要

| DB | サイズ | 内容 |
|----|--------|------|
| catv.mdb | 15.8 MB | 加入者情報・契約詳細・住所・サービス状態 |
| cl.mdb | 86.3 MB | 課金台帳・顧客元帳・支払履歴・残高 |
| koushin.mdb | 238.4 MB | 更新ログ・取引履歴・変更追跡・監査証跡 (最大DB) |
| nhk.mdb | 3.9 MB | NHK受信料データ・加入者情報・支払状態 |
| voip.mdb | 1.8 MB | VoIP加入者データ・通話ログ・サービス設定 |
| ppv.mdb | 188 KB | ペイパービューコンテンツ・購読管理 |
| out.mdb | 340 KB | 出力・エクスポートデータステージング |
| conbart.mdb | 332 KB | 契約・銀行口座マッピング・引落情報 |
| settei.mdb | 292 KB | アプリ設定・設定パラメータ・ユーザー設定 |
| old_catv.mdb | 9.7 MB | catv.mdb の2018年時点バックアップ |
| old_koushin.mdb | 85.2 MB | koushin.mdb の2018年時点バックアップ |

`mente/勝沼町CATV組合/` 配下にも同名DB群の同梱コピーが存在 (FTP送受信用ステージング領域と推定)。

---

## 8. データフロー

```
ユーザー入力 (Forms)
      ↓
バリデーション・処理 (Module1.bas)
      ↓
DAO データベース操作 (*.mdb)
      ↓
FTP 同期 (ezmanager_ment.exe + McWinFTP.cls)
      ↓
CSV エクスポート (銀行・外部システム連携)
      ↓
FDフィードバック受信 (FDBACK/*.dat → 処理)
```

---

## 9. 主な機能

- **契約管理:** 新規加入登録・契約変更・解約・履歴管理
- **課金・支払:** 入金記録・銀行口座管理・自動引落・残高管理
- **データ同期:** FTPベースのデータ同期・FDフォーマット対応
- **特殊サービス:** NHK受信料・VoIP・ペイパービュー
- **照会・帳票:** 多条件検索・状況別照会・履歴レポート
- **システム管理:** パスワード管理・管理者設定・バージョン管理
- **外部連携:** 銀行FDデータ交換・CSV インポート/エクスポート

---

## 10. バージョン履歴概要

| バージョン | 時期 | 内容 |
|-----------|------|------|
| 7.00 | 2013年4月 | 配布版 |
| 7.15 | 2013年4月 | アップデート版 |
| 9.1.26 | 最新 | 現行メインアプリ |
| 7.1.1 | 2022〜2023年 | mente/ver サブプロジェクト |

---

## 11. アーキテクチャ特性

- **スタイル:** モノリシックデスクトップアプリ + 複数補助実行ファイル
- **データ:** ローカルAccessデータベース中心設計
- **ビジネスロジック:** クライアントサイド (Module1.bas に集中)
- **セキュリティ:** パスワード認証 + ファイル暗号化
- **通信:** FTPのみ (クラウド接続なし)
- **言語:** UI・データ・ドキュメントすべて日本語
