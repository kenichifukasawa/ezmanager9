VERSION 5.00
Begin VB.Form FRMINFO 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "ﾊﾞｰｼﾞｮﾝ情報"
   ClientHeight    =   10665
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7590
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMINFO.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7361.172
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   7127.401
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6360
      MouseIcon       =   "FRMINFO.frx":030A
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "FRMINFO.frx":0C09
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      ToolTipText     =   "指定した請求内容の詳細を表示します。"
      Top             =   510
      Width           =   480
   End
   Begin VB.PictureBox Picture16 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'なし
      Height          =   480
      Left            =   270
      MouseIcon       =   "FRMINFO.frx":1A4B
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "FRMINFO.frx":234A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   270
      Width           =   480
   End
   Begin VB.Frame Frame2 
      Caption         =   "バージョンアップ履歴"
      ClipControls    =   0   'False
      Height          =   6495
      Left            =   210
      TabIndex        =   10
      Top             =   1950
      Width           =   7095
      Begin VB.ListBox lstver 
         Height          =   6000
         Left            =   180
         TabIndex        =   11
         Top             =   330
         Width           =   6705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "この製品は、次の方にライセンスされています。"
      ClipControls    =   0   'False
      Height          =   765
      Left            =   240
      TabIndex        =   5
      Top             =   1110
      Width           =   5235
      Begin VB.Label SSPRAISENSUKAISHA 
         Caption         =   "勝沼CATV株式会社"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1050
         TabIndex        =   6
         Top             =   330
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   6000
      TabIndex        =   0
      Top             =   9465
      Width           =   1310
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "ｼｽﾃﾑ情報(&S)..."
      Height          =   345
      Left            =   6000
      TabIndex        =   1
      Top             =   9915
      Width           =   1310
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "バージョンアップ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   525
      Index           =   2
      Left            =   6060
      TabIndex        =   14
      Top             =   1140
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "サポート：Plot Office (ﾌﾟﾛｯﾄ ｵﾌｨｽ）"
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   8550
      Width           =   4845
   End
   Begin VB.Label Label2 
      Caption         =   "連絡先：E-Mail : info@plot-office.com"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   780
      TabIndex        =   8
      Top             =   8880
      Width           =   3165
   End
   Begin VB.Label Label3 
      Caption         =   "携帯電話 : 090-1504-5341"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   4170
      TabIndex        =   7
      Top             =   8880
      Width           =   2985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '実線 (ふちどり)
      Index           =   1
      X1              =   253.544
      X2              =   6930.2
      Y1              =   6429.378
      Y2              =   6429.378
   End
   Begin VB.Label lblTitle 
      Caption         =   "ｱﾌﾟﾘｹｰｼｮﾝ ﾀｲﾄﾙ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1020
      TabIndex        =   3
      Top             =   240
      Width           =   4485
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   267.63
      X2              =   5478.426
      Y1              =   6439.731
      Y2              =   6439.731
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  '右揃え
      Caption         =   "ﾊﾞｰｼﾞｮﾝ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1680
      TabIndex        =   4
      Top             =   690
      Width           =   3825
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"FRMINFO.frx":2C14
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   330
      TabIndex        =   2
      Top             =   9600
      Width           =   5520
   End
End
Attribute VB_Name = "FRMINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ﾚｼﾞｽﾄﾘ ｷｰ ｾｷｭﾘﾃｨ ｵﾌﾟｼｮﾝ...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' ﾚｼﾞｽﾄﾘ ｷｰ ROOT 型...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode Null 文字で終わる文字列
Const REG_DWORD = 4                      ' 32 ﾋﾞｯﾄ数値

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long





Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdok_Click()
  Unload Me
End Sub

Private Sub Form_Load()


    Left = (Screen.Width - Width) / 2   ' フォームを画面の水平方向にセンタリングします。
    Top = (Screen.Height - Height) / 2  ' フォームを画面の縦方向にセンタリングします。
    

    
    Me.Caption = App.Title & "のﾊﾞｰｼﾞｮﾝ情報"
    lblVersion.Caption = "ﾊﾞｰｼﾞｮﾝ " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = "「EZ Manager 8」"
    
    With lstver
    
                .AddItem "ver 9.1.25 2026/05/09"
     .AddItem " 　CLチェックの不具合の修正"
                .AddItem "ver 9.1.24 2026/04/09"
     .AddItem " 　集計の修正"
    
                .AddItem "ver 9.1.21 2026/03/06"
     .AddItem " 　集計の修正"
    
                .AddItem "ver 9.1.20 2026/02/27"
     .AddItem " 　各種修正"
     
                .AddItem "ver 9.1.19 2025/07/09"
     .AddItem " 　CL分分割時のチェック機能を追加"
     
             .AddItem "ver 9.1.18 2024/04/24"
     .AddItem " 　滞納者手数料加算機能を追加"
     
          .AddItem "ver 9.1.16 2024/03/09"
     .AddItem " 　契約の表記名の項目を修正"
    
    
          .AddItem "ver 9.1.15 2024/02/10"
     .AddItem " 　契約の契約期間調整の項目を追加"
    
          .AddItem "ver 9.1.14 2024/01/27"
     .AddItem " 　契約の表記名の項目を追加"
     
      .AddItem "ver 9.1.13 2023/12/20"
     .AddItem " 　契約の更新情報追加機能の修正"
     
      .AddItem "ver 9.1.11 2023/12/15"
     .AddItem " 　契約の更新情報追加機能の追加"
    
       .AddItem "ver 9.1.10 2023/12/11"
     .AddItem " 　CLのチェック機能の追加"
     .AddItem " 　CLの取り込み機能の修正"
     
      .AddItem "ver 9.1.09 2023/11/25"
     .AddItem " 　契約の更新情報管理機能の追加"
     
     
    .AddItem "ver 9.1.08 2023/11/19"
     .AddItem " 　個人情報の項目の追加"
     
     .AddItem "ver 9.1.00 2023/10/11"
     .AddItem " 　CLの処理機能の修正"

    .AddItem "ver 9.0.98 2023/09/30"
     .AddItem " 　予定管理機能の修正"
    .AddItem "ver 9.0.97 2023/09/08"
     .AddItem " 　簡易バージョンアップ機能追加"
    .AddItem "ver 9.0.89 2023/05/23"
     .AddItem " 　各種修正"
     .AddItem "ver 9.0.86 2023/05/13"
     .AddItem " 　CL登録オプション機能を追加"
                         .AddItem "ver 9.0.85 2023/03/04"
     .AddItem " 　詳細伝票作成・削除機能を追加"
                         .AddItem "ver 9.0.83 2022/10/22"
     .AddItem " 　各種修正"
                         
                         .AddItem "ver 9.0.82 2022/08/26"
     .AddItem " 　月別集計の出力機能の追加"

                         .AddItem "ver 9.0.80 2022/08/09"
     .AddItem " 　月別集計の表示内容の修正"
     .AddItem " 　職員の追加"

                     .AddItem "ver 9.0.73 2022/05/31"
     .AddItem " 　表示内容の修正"

                 .AddItem "ver 9.0.72 2022/05/14"
     .AddItem " 　CLの処理機能の修正"

             .AddItem "ver 9.0.71 2022/04/09"
     .AddItem " 　未収伝票作成機能の追加"
             .AddItem "ver 9.0.69 2022/04/01"
     .AddItem " 　職員機能の修正"
            .AddItem "ver 9.0.67 2022/03/31"
     .AddItem " 　文書管理機能連携の追加"
             .AddItem "ver 9.0.62 2022/03/08"
     .AddItem " 　CLの処理機能の修正"
                 .AddItem "ver 9.0.61 2022/03/03"
     .AddItem " 　ログ管理機能の切り離し修正"

             .AddItem "ver 9.0.53 2022/02/12"
     .AddItem " 　ログの入力機能の修正"
    
             .AddItem "ver 9.0.46 2022/02/08"
     .AddItem " 　CLの処理機能の修正"
     .AddItem " 　予定管理の処理を分離"
     
             .AddItem "ver 9.0.45 2022/01/30"
     .AddItem " 　予定管理の処理を分離"

         .AddItem "ver 9.0.44 2022/01/29"
     .AddItem " 　予定管理の更新処理を修正"
     .AddItem " 　個人データの項目を追加"
         
         .AddItem "ver 9.0.42 2021/11/13"
     .AddItem " 　予定管理の表示内容を修正"
   
        .AddItem "ver 9.0.41 2021/10/28"
     .AddItem " 　光の制限の修正"
     .AddItem " 　個人情報の変更時の制限の修正"
        
        .AddItem "ver 9.0.38 2021/10/16"
     .AddItem " 　予定管理の登録時の終了日の不具合を修正"
     .AddItem " 　FDDデータ更新時の不具合を修正"
    .AddItem " 　TOOLSチェックにNHK請求チェック機能を追加"

    .AddItem "ver 9.0.36 2021/07/24"
     .AddItem " 　キャディックス用データ出力ファイル自動転送機能を追加"
    
    .AddItem "ver 9.0.34 2021/07/01"
     .AddItem " 　キャディックス用データ出力ファイル保存機能を追加"
    
    .AddItem "ver 9.0.33 2021/06/15"
     .AddItem " 　キャディックス用データ出力機能を追加"
    
    .AddItem "ver 9.0.32 2021/06/13"
     .AddItem " 　詳細を修正"
    
    
    .AddItem "ver 9.0.28 2021/05/22"
     .AddItem " 　CLの読込方法（内訳）を修正"
     .AddItem " 　CLの読込チェックを修正"

    .AddItem "ver 9.0.26 2021/05/13"
     .AddItem " 　詳細を修正"

                    .AddItem "ver 9.0.24 2021/04/30"
     .AddItem " 　詳細を修正"


                        .AddItem "ver 9.0.23 2021/04/15"
     .AddItem " 　ケーブルライン集計機能を修正"
     .AddItem " 　ケーブルライン請求書通知機能を修正"

                    .AddItem "ver 9.0.22 2021/04/13"
     .AddItem " 　ケーブルライン集計機能を追加"
     .AddItem " 　ケーブルライン請求書通知機能を追加"

                .AddItem "ver 9.0.20 2021/04/08"
     .AddItem " 　ケーブルライン管理機能を修正"

        .AddItem "ver 9.0.19 2021/04/01"
     .AddItem " 　予定管理機能の更新機能を修正"

            .AddItem "ver 9.0.17 2021/03/27"
     .AddItem " 　ケーブルラインエラー管理機能を追加"
            
            .AddItem "ver 9.0.16 2021/03/26"
     .AddItem " 　ケーブルライン請求管理機能を追加"
     .AddItem " 　ケーブルライン集計管理機能を追加"
     
            
            .AddItem "ver 9.0.15 2021/03/18"
     .AddItem " 　ケーブルライン管理機能を修正"


        .AddItem "ver 9.0.14 2021/03/16"
     .AddItem " 　ケーブルライン管理機能を修正"

     .AddItem "ver 9.0.13 2021/03/12"
     .AddItem " 　個人データの項目の追加"
     
    .AddItem "ver 9.0.11 2021/03/04"
     .AddItem " 　個人データの項目の修正"
     
     .AddItem "ver 9.0.10 2021/02/19"
     .AddItem " 　CSV出力機能に項目を修正"
     .AddItem "ver 9.0.09 2021/02/17"
     .AddItem " 　個人データの項目の追加"
     .AddItem " 　CSV出力機能に項目を追加"
     
    
         .AddItem "ver 9.07 2021/02/12"
     .AddItem " 　CSV出力機能に項目を修正"

     .AddItem "ver 9.06 2021/02/09"
     .AddItem " 　個人データの項目の追加"
     .AddItem " 　CSV出力機能に項目を修正"
     
     
        .AddItem "ver 9.04 2021/02/05"
     .AddItem " 　個人データの項目の追加"
    .AddItem " 　CSV出力機能に項目を修正"

    .AddItem "ver 9.02 2021/01/23"
     .AddItem " 　データ更新機能の修正"

    
    .AddItem "ver 9.0.00 2021/01/19"
     .AddItem " 　予定管理機能を追加"
     .AddItem " 　ケーブルライン管理機能を追加"
     
     .AddItem "ver 8.73 2020/10/13"
     .AddItem " 　NHK料金管理機能を追加"
     
     .AddItem "ver 8.71 2020/09/11"
     .AddItem " 　集計機能の不具合を修正"
     .AddItem " 　個人情報の管理項目を修正"
     
     .AddItem "ver 8.69 2020/09/03"
     .AddItem " 　CSV出力機能に項目を修正"
     
    .AddItem "ver 8.68 2020/08/28"
     .AddItem " 　CSV出力機能に項目を追加"
     
    .AddItem "ver 8.67 2020/08/23"
     .AddItem " 　CATV集計機能を修正"
     
    .AddItem "ver 8.66 2020/08/16"
     .AddItem " 　CATV集計機能を修正"
     
        .AddItem "ver 8.65 2020/08/14"
     .AddItem " 　個人情報の管理項目を修正"
     
        .AddItem "ver 8.64 2020/08/13"
     .AddItem " 　個人情報の管理項目を追加"
     
        .AddItem "ver 8.62 2020/08/09"
     .AddItem " 　起動モードを追加"
     
       .AddItem "ver 8.61 2020/08/07"
     .AddItem " 　個人情報の管理項目を追加"
     .AddItem " 　CSV出力機能に項目を追加"
     
     .AddItem "ver 8.60 2020/08/01"
     .AddItem " 　CATV集計機能を修正"
    .AddItem " 　バージョンアップ機能を修正"
    
    .AddItem "ver 8.58 2020/07/31"
     .AddItem " 　CATV集計機能を修正"
    
     .AddItem "ver 8.57 2020/07/02"
     .AddItem " 　データ更新チェック機能を修正"
    
    
    .AddItem "ver 8.56 2020/06/25"
     .AddItem " 　CATV集計機能を修正"
    
    .AddItem "ver 8.55 2020/06/16"
     .AddItem " 　CATV集計機能を追加"
     
    .AddItem "ver 8.54 2020/06/05"
     .AddItem " 　CSV出力機能の不具合を修正"
     
    .AddItem "ver 8.53 2020/05/08"
     .AddItem " 　NHK更新の不具合を修正"
     
     .AddItem "ver 8.52 2020/05/03"
     .AddItem " 　データ更新時のNHKの不具合を修正"
     .AddItem " 　NHK集計の不具合を修正"
    
      .AddItem "ver 8.51 2020/03/31"
     .AddItem " 　NHK管理機能を追加"
     
        .AddItem "ver 8.50 2019/11/05"
     .AddItem " 　期間別請求データ詳細の出力機能を追加"
       .AddItem " 　期間別入金データ詳細の出力機能を追加"


    .AddItem "ver 8.48 2019/10/31"
     .AddItem " 　作業ログ自動記録機能を追加"
     
     .AddItem "ver 8.45 2019/10/14"
     .AddItem " 　NHKチェック機能をマルチに対応"
     
      .AddItem "ver 8.42 2019/10/13"
     .AddItem " 　NHKチェック機能を追加"
    
     .AddItem "ver 8.41 2019/08/02"
     .AddItem " 　マップ機能を追加"
     
     .AddItem "ver 8.39 2019/07/31"
        .AddItem " 　未収データ出力機能を修正"
    
            .AddItem "ver 8.38 2019/07/11"
        .AddItem " 　各種契約に追加項目の管理機能を追加"
        .AddItem " 　各種契約に出力機能に項目を追加"


        .AddItem "ver 8.36 2019/07/08"
        .AddItem " 　未収データ出力機能を追加"
    
          .AddItem "ver 8.35 2019/05/23"
        .AddItem " 　仮読み込み機能を修正0"
    
    
        .AddItem "ver 8.34 2019/05/18"
        .AddItem " 　繰越の詳細データの条件を修正"
        
        .AddItem "ver 8.33 2019/05/14"
        .AddItem " 　入金確認メッセージの書式を修正"
        
        .AddItem "ver 8.31 2019/05/12"
        .AddItem " 　繰越の詳細データの条件を修正"
        
        .AddItem "ver 8.28 2019/05/05"
        .AddItem " 　滞納者リスト抽出条件を追加"
         .AddItem " 　０更新時の不具合を修正"
        
        .AddItem "ver 8.26 2019/04/26"
        .AddItem " 　仮読み込み機能を修正"
    
        .AddItem "ver 8.25 2019/04/25"
        .AddItem " 　仮読み込み機能を追加"
        .AddItem " 　仮読み込み表示機能を追加"
        .AddItem " 　仮読み込み出力機能を追加"
        
        .AddItem "ver 8.24 2019/04/13"
        .AddItem " 　仮読み込みのUIのサンプル"
    
        .AddItem "ver 8.23 2019/03/31"
        .AddItem " 　職員の追加及び設定の修正"

        .AddItem "ver 8.22 2019/03/30"
        .AddItem " 　ＣＳＶ出力情報の項目の修正"
        .AddItem "ver 8.21 2019/02/01"
        .AddItem " 　ＣＳＶ出力情報の項目を追加"
        .AddItem "ver 8.20 2019/01/15"
        .AddItem " 　検索の条件のチェックの修正"
        .AddItem " 　契約者情報の登録・更新時のチェック機能の修正"
        .AddItem " 　ＣＳＶ出力情報の不具合を修正"
        .AddItem " 　請求入金集計のＣＳＶ出力機能を修正"
        .AddItem " 　請求書扱いのデータのCSV出力機能を追加"
        .AddItem "ver 8.19 2019/01/08"
        .AddItem " 　職員の追加及び設定の修正"
        .AddItem " 　ログからの個人情報へのジャンプ機能を追加"
        .AddItem "ver 8.18 2018/11/09"
        .AddItem " 　集計機能の不具合を修正"
        .AddItem "ver 8.17 2018/11/04"
        .AddItem " 　県民信用用請求データの請求・更新の機能を修正"
        .AddItem "ver 8.16 2018/10/18"
        .AddItem " 　農協用請求データの請求・更新の機能を修正"
        .AddItem "ver 8.15 2018/10/18"
        .AddItem " 　山梨中央用請求データの請求・更新の機能を修正"
        .AddItem "ver 8.14 2018/09/11"
        .AddItem " 　請求データの請求・更新の機能を統合"
        .AddItem "ver 8.07 2017/09/24"
        .AddItem " 　FDDのドライブの設定機能を追加"
        .AddItem "ver 8.00 2017/08/15"
        .AddItem " 　Windows10対応"
        .AddItem "ver 7.64 2017/04/13"
        .AddItem " 　料金変更における修正"
        .AddItem "ver 7.63 2016/12/07"
        .AddItem " 　郵便局の変更における修正"
        .AddItem "ver 7.50 2016/10/11"
        .AddItem " 　TOOLのチェック機能を修正"
        .AddItem "ver 7.46 2016/08/12"
        .AddItem " 　TOOLのチェック機能を修正"
        .AddItem "ver 7.45 2016/05/08"
        .AddItem " 　データを修正（手動修正96）"
        .AddItem "ver 7.44 2016/04/08"
        .AddItem " 　スタッフを修正"
        .AddItem "ver 7.43 2016/04/06"
        .AddItem " 　スタッフを追加"
        .AddItem "ver 7.42 2015/07/28"
        .AddItem " 　サブモード用バックアップ機能を追加"
        .AddItem "ver 7.39 2015/07/01"
        .AddItem " 　HNKの契約で０円時に更新処理の不具合を修正"
        .AddItem "ver 7.35 2015/03/27"
        .AddItem " 　HNKの契約で０円時に更新処理の不具合を修正"
        .AddItem "ver 7.33 2015/03/14"
        .AddItem " 　個人データにCATV・Internetの解約日を追加"
        .AddItem " 　データベース修正（手動修正94）"
        .AddItem "ver 7.30 2015/02/17"
        .AddItem " 　各種集計の年号の不具合を修正"
        .AddItem "ver 7.29 2014/11/11"
        .AddItem " 　データを修正（手動修正93）"
        .AddItem "ver 7.28 2013/08/07"
        .AddItem " 　CSV出力の修正"
        .AddItem "ver 7.27 2013/03/12"
        .AddItem " 　HNKの金額修正"
        .AddItem "ver 7.26 2012/11/2"
        .AddItem " 　更新進行状況表示の修正"
        .AddItem "ver 7.24 2012/10/17"
        .AddItem " 　接続チェックの修正"
        .AddItem "ver 7.23 2012/10/11"
        .AddItem " 　CSV出力エラーの修正"
        .AddItem "ver 7.22 2012/10/09"
        .AddItem " 　ログの修正"
        .AddItem " 　単体入金処理修正"
        .AddItem "ver 7.21 2012/10/09"
        .AddItem " 　ログの文字数を増設"
        .AddItem "ver 7.20 2012/09/24"
        .AddItem " 　データを修正（手動修正91）"
        .AddItem "ver 7.19 2012/09/07"
        .AddItem " 　その他契約単体入金処理機能を追加"
        .AddItem "ver 7.18 2012/08/17"
        .AddItem " 　CSV出力項目追加"
        .AddItem "ver 7.17 2012/08/09"
        .AddItem " 　データを修正（手動修正90）"
        .AddItem "ver 7.16 2012/06/27"
        .AddItem " 　FDDデータ読み込み機能修正"
        .AddItem "ver 7.15 2012/06/16"
        .AddItem " 　FDDデータ読み込みテスト機能追加"
        .AddItem "ver 7.12 2012/06/09"
        .AddItem " 　データを修正（手動修正89）"
        .AddItem "ver 7.10 2012/05/29"
        .AddItem " 　ロゴの抽出機能を追加"
        .AddItem "ver 7.00 2012/05/21"
        .AddItem " 　Windows7対応"
        .AddItem "ver 6.44 2012/04/21"
        .AddItem " 　０更新日を保存・表示するように修正"
        .AddItem " 　０更新日実行時にパスワードを入力するように修正"
        .AddItem " 　０更新の警告日を２０日に修正"
        .AddItem "ver 6.41 2012/04/12"
        .AddItem " 　ＮＨＫ不正チェックを修正"
        .AddItem "ver 6.40 2012/11/09"
        .AddItem " 　デッドロック時のエラー回避を修正"
        .AddItem "ver 6.40 2012/11/09"
        .AddItem " 　デッドロック時のエラー回避を修正"
        .AddItem "ver 6.39 2012/10/04"
        .AddItem " 　NHK料金修正における手動修正"
        .AddItem "ver 6.38 2012/10/02"
        .AddItem " 　ログ内容の削除表示機能を追加"
        .AddItem "ver 6.37 2012/09/08"
        .AddItem " 　ログ内容の項目管理を修正"
        .AddItem "ver 6.36 2012/09/01"
        .AddItem " 　ログ内容の項目を追加修正"
        .AddItem "ver 6.35 2012/08/11"
        .AddItem " 　ログ内容を300文字から500文字に修正"
        .AddItem "ver 6.34 2012/08/09"
        .AddItem " 　データを修正（手動修正83）"
        .AddItem " 　UIの表示名エラーを修正"
        .AddItem "ver 6.33 2012/08/04"
        .AddItem " 　ログ更新時のエラー回避処理を修正"
        .AddItem " 　ファイルのバージョンアップ機能を修正"
        .AddItem "ver 6.32 2012/06/12"
        .AddItem " 　データチェック２を修正"
        .AddItem "ver 6.31 2012/06/08"
        .AddItem " 　VoIP読込の制限機能を追加"
        .AddItem "ver 6.30 2012/06/01"
        .AddItem " 　備考サイズの修正と表示機能を追加"
        .AddItem "ver 6.29 2012/05/09"
        .AddItem " 　データを修正（手動修正82）"
        .AddItem "ver 6.28 2012/05/08"
        .AddItem " 　HNKチェックのエラーログを修正"
        .AddItem " 　HNKチェックの条件（NHK名前）を変更修正"
        .AddItem "ver 6.27 2012/05/05"
        .AddItem " 　NHK契約データチェックシステム修正"
        .AddItem "ver 6.26 2012/05/02"
        .AddItem " 　NHK契約データチェックシステム修正"
        .AddItem "ver 6.25 2012/04/12"
        .AddItem " 　データを修正（手動修正81）"
        .AddItem "ver 6.24 2012/04/03"
        .AddItem " 　NHK契約データチェックシステム追加"
        .AddItem " 　自動更新システム修正"
        .AddItem " 　NHK更新システム修正"
        .AddItem " 　データを修正（手動修正80）"
        .AddItem "ver 6.23 2012/03/12"
        .AddItem " 　データを修正（手動修正79）"
        .AddItem "ver 6.22 2012/02/28"
        .AddItem " 　NHKデータチェックシステム追加（テスト１）"
        .AddItem "ver 6.21 2012/02/10"
        .AddItem " 　データを修正（手動修正78）"
        .AddItem "ver 6.20 2012/01/20"
        .AddItem " 　データチェックシステム修正"
        
        .AddItem "ver 6.19 2011/11/23"
        .AddItem " 　ログ管理のサイズの修正"
        .AddItem "ver 6.18 2011/10/18"
        .AddItem " 　NHKデータチェック機能追加"
        .AddItem "ver 6.17 2011/10/16"
        .AddItem " 　NHKデータ簡易チェック機能追加"
        .AddItem "ver 6.16 2011/10/13"
        .AddItem " 　ログ管理の変更・表示方法の修正"
        .AddItem "ver 6.15 2011/10/13"
        .AddItem " 　データを修正（手動修正77）"
        .AddItem "ver 6.14 2011/10/13"
        .AddItem " 　データを修正（手動修正76）"
        .AddItem "ver 6.13 2011/10/05"
        .AddItem " 　データを修正（手動修正75）"
        .AddItem "ver 6.12 2011/10/01"
        .AddItem " 　データチェックシステム修正"
        .AddItem "ver 6.11 2011/08/14"
        .AddItem " 　データチェックシステム修正"
        .AddItem "ver 6.10 2011/08/13"
        .AddItem " 　データを修正（手動修正74）"
        .AddItem " 　データチェックシステム修正"
        .AddItem "ver 6.09 2011/08/11"
        .AddItem " 　データを修正（手動修正73）"
        .AddItem " 　データチェックシステム修正"
        .AddItem "ver 6.08 2011/08/05"
        .AddItem " 　ログ管理機能の変更機能の追加"
        .AddItem " 　ログ管理機能の表示機能の修正"
        .AddItem "ver 6.07 2011/08/05"
        .AddItem " 　データを修正（手動修正72）"
        .AddItem " 　自動更新システム修正"
        .AddItem "ver 6.06 2011/08/02"
        .AddItem " 　峡東CATVからのデータのチェック機能を修正"
        .AddItem "ver 6.05 2011/08/02"
        .AddItem " 　峡東CATVからのデータのチェック機能を修正"
        .AddItem "ver 6.04 2011/08/01"
        .AddItem " 　峡東CATVからのデータのチェック機能を追加"
        .AddItem "ver 6.03 2011/07/31"
        .AddItem " 　起動時エラーの修正"
        .AddItem "ver 6.02 2011/07/26"
        .AddItem " 　デッドロックの修正"
        .AddItem " 　ログ管理機能追加"
        .AddItem "ver 6.01 2011/07/25"
        .AddItem " 　排他モードの修正"
        .AddItem "ver 6.00 2011/07/24"
        .AddItem " 　クライアントサーバシステムに移行"
        .AddItem "ver 5.09 2011/07/13"
        .AddItem " 　自動更新システム修正"
        .AddItem " 　データチェックシステム修正"
        .AddItem " 　FDD再出力システム修正"
        .AddItem "ver 5.08 2011/07/12"
        .AddItem " 　データチェックシステム戻し"
        .AddItem "ver 5.07 2011/07/01"
        .AddItem " 　データチェックシステム修正"
        .AddItem " 　入金データ更新システム修正"
        .AddItem "ver 5.06 2011/06/15"
        .AddItem " 　データチェックシステム修正"
        .AddItem "ver 5.05 2011/05/18"
        .AddItem " 　FDD出力時メッセージエラー修正"
        .AddItem " 　FDD更新時エラーを修正"
        .AddItem "ver 5.04 2011/05/10"
        .AddItem " 　データを修正（手動修正71）"
        .AddItem "ver 5.03 2011/04/29"
        .AddItem " 　現金自動更新機能を修正"
        .AddItem "ver 5.02 2011/04/22"
        .AddItem " 　契約更新の諸修正"
        .AddItem "ver 5.01 2011/04/08"
        .AddItem " 　契約使用制限の実施"
        .AddItem " 　契約更新の諸修正"
        .AddItem "ver 5.00 2011/04/01"
        .AddItem " 　バージョン５プロトタイプ"
        .AddItem "ver 2.67 2010/10/13"
        .AddItem " 　FDD出力テスト機能を追加"
        .AddItem "ver 2.66 2010/09/30"
        .AddItem " 　データを修正（手動修正68）"
        .AddItem "ver 2.65 2010/08/23"
        .AddItem " 　WindowsXP対応"
        .AddItem "ver 2.64 2010/06/17"
        .AddItem " 　ＦＤＤの出力終了時のチェック機能を追加"
        .AddItem " 　バックアップ時のＮＨＫファイルの圧縮エラー修正"
        .AddItem "ver 2.62 2010/06/02"
        .AddItem " 　データを修正（手動修正67）"
        .AddItem "ver 2.61 2010/05/14"
        .AddItem " 　ＦＤＤ出力システムの修正"
        .AddItem "ver 2.60 2010/04/01"
        .AddItem " 　JAデータ変換システムを追加"
        .AddItem "ver 2.58 2010/03/09"
        .AddItem " 　JAデータ変換システムを修正"
        .AddItem "ver 2.57 2010/02/16"
        .AddItem " 　ＣＳＶデータ出力項目を修正"
        .AddItem "ver 2.53 2010/02/13"
        .AddItem " 　データを修正（手動修正66）"
        .AddItem "ver 2.52 2010/02/07"
        .AddItem " 　ＣＳＶデータ出力項目を追加修正"
        .AddItem "ver 2.50 2009/12/14"
        .AddItem "ver 2.51 2010/02/03"
        .AddItem " 　データを修正（手動修正65）"
        .AddItem "ver 2.50 2009/12/14"
        .AddItem " 　JAデータ変換システムを追加"
        .AddItem "ver 2.49 2009/09/09"
        .AddItem " 　データチェック時データ変更機能を修正"
        .AddItem " 　データを修正（手動修正64）"
        .AddItem "ver 2.48 2009/09/05"
        .AddItem " 　データチェック時データ変更機能を追加"
        .AddItem "ver 2.47 2009/06/09"
        .AddItem " 　データを修正（手動修正63）"
        .AddItem "ver 2.46 2009/04/30"
        .AddItem " 　データを修正（手動修正62）"
        .AddItem "ver 2.44 2009/04/13"
        .AddItem " 　FDD出力機能を修正"
        .AddItem "ver 2.44 2009/04/02"
        .AddItem " 　入金機能を修正"
        .AddItem "ver 2.43 2009/03/29"
        .AddItem " 　更新時エラーを修正（手動修正61）"
        .AddItem "ver 2.41 2009/03/26"
        .AddItem " 　自動バージョンアップ機能を追加"
        .AddItem "ver 2.40 2009/03/25"
        .AddItem " 　自動バックアップ機能を追加"
        .AddItem "ver 2.30 2009/01/10"
        .AddItem " 　入金日付修正機能を修正"
        .AddItem "ver 2.29 2008/11/13"
        .AddItem " 　主契約登録機能を修正）"
        .AddItem "ver 2.28 2008/11/11"
        .AddItem " 　データチェック機能を修正（手動修正60）"
        .AddItem "ver 2.27 2008/10/30"
        .AddItem "  自動繰越更新機能を修正"
        .AddItem "ver 2.26 2008/10/14"
        .AddItem "  FDD出力時データチェック機能を追加"
        .AddItem "ver 2.24 2008/10/12"
        .AddItem "  FDD出力時データチェック機能を修正"
        .AddItem "ver 2.23 2008/10/03"
        .AddItem "  データ修正（手動修正59）"
        .AddItem "ver 2.22 2008/09/03"
        .AddItem "  バージョンアップ機能を追加"
        .AddItem "ver 2.10 2008/06/11"
        .AddItem " 　データチェック機能を修正（手動修正58）"
        .AddItem "ver 2.09 2008/06/07"
        .AddItem "  最新入金日修正機能を追加"
        .AddItem "ver 2.08 2008/05/19"
        .AddItem "  請求入金詳細ＮＨＫ表示機能を追加"
        .AddItem "ver 2.07 2008/04/14"
        .AddItem " 　データチェック機能を修正"
        .AddItem "ver 2.06 2008/04/01"
        .AddItem " 　データチェック機能を修正（手動修正57）"
        .AddItem "ver 2.05 2008/03/31"
        .AddItem " 　データ更新機能を修正"
        .AddItem "ver 2.04 2008/03/18"
        .AddItem " 　データチェック機能を修正"
        .AddItem "ver 2.03 2008/03/06"
        .AddItem " 　主契約を変更時のメッセージを修正"
        .AddItem " 　損金・取り消し金の処理を修正"
        .AddItem "ver 2.02 2007/12/13"
        .AddItem " 　データチェック機能を修正（手動修正56）"
        .AddItem "ver 2.01 2007/12/13"
        .AddItem " 　延滞者一覧のCSV出力機能を追加"
        .AddItem "ver 2.00 2007/12/13"
        .AddItem " 　データチェック機能を修正"
        .AddItem "ver 1.99 2007/12/06"
        .AddItem " 　データ修正（手動修正55）"
        .AddItem "ver 1.98 2007/12/05"
        .AddItem " 　簡易修正機能を修正"
        .AddItem "ver 1.97 2007/11/27"
        .AddItem " 　ふりがな機能を修正"
        .AddItem " 　結果表示機能を修正"
        .AddItem "ver 1.95 2007/11/21"
        .AddItem " 　データチェック機能を修正（手動修正54）"
        .AddItem "ver 1.94 2007/11/13"
        .AddItem " 　主契約変更機能を修正"
        .AddItem "ver 1.92 2007/10/16"
        .AddItem " 　データチェック機能を修正"
        .AddItem "ver 1.92 2007/10/15"
        .AddItem " 　データ修正（手動修正53）"
        .AddItem "ver 1.91 2007/10/13"
        .AddItem " 　データ修正（手動修正52）"
        .AddItem "ver 1.90 2007/10/11"
        .AddItem " 　データ修正（手動修正51）"
        .AddItem "ver 1.89 2007/10/03"
        .AddItem " 　自動更新時エラーを修正"
        .AddItem " 　ＦＤＤ更新時エラーを修正"
        .AddItem "ver 1.88 2007/10/03"
        .AddItem " 　データ修正（手動修正50）"
        .AddItem "ver 1.87 2007/10/02"
        .AddItem " 　契約状況データCSV出力項目を追加"
        .AddItem "ver 1.86 2007/09/28"
        .AddItem " 　契約状況データCSV出力機能を修正"
        .AddItem "ver 1.85 2007/09/26"
        .AddItem " 　その他の契約内容を変更時のエラーを修正"
        .AddItem "ver 1.84 2007/09/25"
        .AddItem " 　契約状況データCSV出力機能を追加"
        .AddItem "ver 1.83 2007/09/13"
        .AddItem " 　FDD更新時エラーを修正"
        .AddItem "ver 1.82 2007/08/01"
        .AddItem " 　細かいを修正（手動修正49）"
        .AddItem "ver 1.81 2007/05/25"
        .AddItem " 　細かいを修正（手動修正48）"
        .AddItem "ver 1.80 2007/05/16"
        .AddItem " 　データチェック機能を修正（手動修正47）"
        .AddItem "ver 1.78 2007/05/01"
        .AddItem " 　FDD更新システムを修正（手動修正45）"
        .AddItem "ver 1.77 2007/05/01"
        .AddItem " 　FDD更新システムを修正（手動修正44）"
        .AddItem "ver 1.76 2007/04/14"
        .AddItem " 　リアル未納者の内訳詳細表示機能を追加"
        .AddItem "ver 1.75 2007/04/13"
        .AddItem " 　取消・損金管理機能に内訳詳細表示機能を追加"
        .AddItem "ver 1.74 2007/04/04"
        .AddItem " 　データチェック機能を修正（手動修正42）"
        .AddItem " 　FDD更新機能を修正"
        .AddItem " 　PPV/VoIP更新機能を修正（手動修正42）"
        .AddItem "ver 1.72 2007/04/02"
        .AddItem " 　自動更新機能を修正（手動修正41）"
        .AddItem " 　取消・損金管理機能を修正"
        .AddItem "ver 1.71 2007/03/30"
        .AddItem " 　取消・損金管理機能を修正"
        .AddItem "ver 1.70 2007/03/23"
        .AddItem " 　取消・損金管理機能を追加（手動修正40）"
        .AddItem "ver 1.67 2007/03/06"
        .AddItem " 　検索非表示機能を追加（手動修正39）"
        .AddItem "ver 1.65 2007/02/28"
        .AddItem " 　金額チェック機能の修正（手動修正38）"
        .AddItem "ver 1.64 2007/01/14"
        .AddItem " 　FDD出力の修正（手動修正37）"
        .AddItem "ver 1.63 2007/01/31"
        .AddItem " 　自動更新日付計算エラーの修正"
        .AddItem " 　自動更新及びFDD更新制御機能を追加"
        .AddItem "ver 1.60 2007/01/15"
        .AddItem " 　契約変更時金額計算エラーの修正"
        .AddItem "ver 1.58 2007/01/09"
        .AddItem " 　自動更新時の日付エラーの修正"
        .AddItem "ver 1.57 2007/01/05"
        .AddItem " 　自動更新時のエラーの修正"
        .AddItem "ver 1.56 2006/12/29"
        .AddItem " 　FDD出力時のエラーの修正（手動修正35）"
        .AddItem "ver 1.55 2006/12/27"
        .AddItem " 　特殊入金時のエラーの修正（手動修正34）"
        .AddItem "ver 1.53 2006/12/14"
        .AddItem " 　FDD更新時エラーの修正（手動修正31,32,33）"
        .AddItem "ver 1.50 2006/12/12"
        .AddItem " 　特殊金額入金機能を追加"
        .AddItem "ver 1.49 2006/12/10"
        .AddItem " 　契約状況変更時エラーの修正（手動修正30）"
        .AddItem "ver 1.47 2006/11/16"
        .AddItem " 　契約作動フィールド追加及び修正（手動修正28/29）"
        .AddItem "ver 1.46 2006/11/13"
        .AddItem "  入金処理で更新処理エラーを修正"
        .AddItem "ver 1.42 2006/10/16"
        .AddItem "  請求額なくNHKのみ契約者の契約更新機能を修正"
        .AddItem "ver 1.40 2006/10/02"
        .AddItem "  契約更新機能を修正"
        .AddItem "ver 1.37 2006/09/14"
        .AddItem "  個人契約金額情報変更機能を修正"
        .AddItem "ver 1.36 2006/09/04"
        .AddItem "  請求入金詳細表示機能を修正"
        .AddItem "ver 1.35 2006/08/29"
        .AddItem "  未入金請求詳細表示機能を追加"
        .AddItem "ver 1.27 2006/07/31"
        .AddItem "  現金入金履歴内容表示エラーを修正"
        .AddItem "ver 1.26 2006/07/31"
        .AddItem "  請求入金集計の詳細内容表示エラーを修正"
        .AddItem "ver 1.25 2006/07/25"
        .AddItem "  新FDD更新システム追加"
        .AddItem "ver 1.21 2006/07/21"
        .AddItem "  新自動更新システム追加"
        .AddItem "ver 1.20 2006/07/11"
        .AddItem "  新FDD出力システムの修正"
        .AddItem "ver 1.19 2006/07/10"
        .AddItem "  新データチェックシステム追加"
        .AddItem "ver 1.18 2006/07/07"
        .AddItem "  新請求入金集計システムに移行"
        .AddItem "  新FDD出力システムに移行"
        .AddItem "  新入金システムに移行"
        .AddItem "ver 1.14 2006/07/02"
        .AddItem "  現金入金抽出機能を追加"
        .AddItem "  FDD更新時の不具合を修正"
        .AddItem "  入金更新時の不具合を修正"
        .AddItem "ver 1.13 2006/06/28"
        .AddItem "  請求書扱いの手動更新時の不具合を修正"
        .AddItem "  請求・入金データエクスポート機能を追加"
        .AddItem "ver 1.10 2006/06/23"
        .AddItem "  PPV/VoIPデータインポート機能を追加"
        .AddItem "ver 1.08 2006/06/13"
        .AddItem "  ＦＤＤ再出力時機能のエラーを修正"
        .AddItem "ver 1.06 2006/06/09"
        .AddItem "  入金詳細情報集計機能を修正"
        .AddItem "  入金情報の日付の変更機能の追加"
        .AddItem "  Ｑ検索の不具合を修正"
        
        .AddItem "ver 1.04 2006/06/07"
        .AddItem "  入金処理システムの不具合を修正"
        
        .AddItem "ver 1.01 2006/06/01"
        .AddItem "  入金処理システムエラーを修正"
        .AddItem "  請求書扱いの手動更新システム追加"
        .AddItem "ver 0.96 2006/05/24"
        .AddItem "  個人情報の口座情報の移行エラーの修正"
        .AddItem "  入金詳細表示エラーの修正"
        .AddItem "  最終振替結果コード表示機能追加"
        
        .AddItem "ver 0.95 2006/05/24"
        .AddItem "  支払契約別検索機能を追加"
        .AddItem "  請求書用手動契約更新機能を追加"
        
        .AddItem "ver 0.92 2006/05/18"
        .AddItem "  銀行別集計機能開始コンバート機能を追加"
        .AddItem "  銀行別集計機能ＦＤＤ出力機能を追加"
        .AddItem "  銀行別集計機能集計機能を追加"
        .AddItem "ver 0.91 2006/05/17"
        .AddItem "  入金詳細表示機能を追加"
        .AddItem "  NHK詳細集計表示機能を追加"
        .AddItem "ver 0.90 2006/05/15"
        .AddItem "  FDD出力時の修正"
        .AddItem "ver 0.88 2006/05/11"
        .AddItem "  メインコンバートの修正"
        .AddItem "ver 0.87 2006/05/10"
        .AddItem "  主契約の項目の修正"
        .AddItem "ver 0.86 2006/05/09"
        .AddItem "  FDD更新時エラーを修正"
        .AddItem "  顧客登録時エラーを修正"
        .AddItem "ver 0.83 2006/05/03"
        .AddItem "  請求・入金集計機能を追加"
        .AddItem "ver 0.79 2006/05/02"
        .AddItem "  滞納者リスト表示・印刷機能を修正"
        .AddItem "  詳細集計用データベースを修正"
        .AddItem "ver 0.77 2006/05/02"
        .AddItem "  加入日のデータを修正"
        .AddItem "  詳細集計用データベースを構築"
        .AddItem "  既存データの再構築"
        .AddItem "  ＦＤＤ出力システム修正"
        .AddItem "  ＦＤＤ更新システム修正"
        .AddItem "ver 0.74 2006/04/23"
        .AddItem "  追加コンバートシステムを修正"
        .AddItem "  NHK集計システムを修正"
        .AddItem "ver 0.73 2006/04/13"
        .AddItem "  管理者モードを修正"
        .AddItem "ver 0.71 2006/04/10"
        .AddItem "   契約一覧を修正"
        .AddItem "   個人データ変更時エラーを修正"
        .AddItem "   データ集計を修正"
        .AddItem "   データコンバートシステムを修正"
        .AddItem "ver 0.65 2006/03/30"
        .AddItem "   データコンバートシステムを追加"
        .AddItem "ver 0.62 2006/03/25"
        .AddItem "   ノード番号・図面番号を追加"
        .AddItem "   メイン画面の色の修正"
        .AddItem "ver 0.61 2006/03/21"
        .AddItem "   郵便局用FDD確認機能を修正"
        .AddItem "ver 0.60 2006/03/20"
        .AddItem "   個人客のフリガナ機能を修正"
        .AddItem "   郵便局用FDD作成機能を修正"
        .AddItem "   金融機関別集計機能を修正"
        .AddItem "   現金入金処理・管理機能を修正"
        .AddItem "   ＨＮＫ処理・管理機能を修正"
        .AddItem "ver 0.58 2006/03/05"
        .AddItem "   個人情報登録時エラーを修正"
        .AddItem "   現金入金処理機能を追加"
        .AddItem "   入金履歴管理機能を修正"
        
        .AddItem "ver 0.57 2006/03/02"
        .AddItem "   個人情報管理に管理ＩＤを追加"
        .AddItem "   課金計算システムを修正"
        .AddItem "   主契約登録システムを修正"
        .AddItem "   主契約変更システムを修正"
        .AddItem "   主契約削除システムを修正"
        .AddItem "   その他契約登録システムを修正"
        .AddItem "   その他契約変更システムを修正"
        .AddItem "   その他契約削除システムを修正"
        .AddItem "ver 0.56 2006/03/01 システム内容の分離"
        .AddItem "ver 0.55 2006/02/28 システム内容及びＧＵＩを変更"
        .AddItem "ver 0.54 2006/02/26 ユーザーインターフェイス変更"
        .AddItem "ver 0.52 2006/02/24 自動バージョンアップ機能を追加"
        .AddItem "ver 0.51 2006/02/23 メンテナンス機能を追加"
        .AddItem "ver 0.50 2006/02/22 プロトタイプ"
    End With
    
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' ﾚｼﾞｽﾄﾘからｼｽﾃﾑ情報ﾌﾟﾛｸﾞﾗﾑのﾊﾟｽ\名前を取得します...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' ﾚｼﾞｽﾄﾘからｼｽﾃﾑ情報ﾌﾟﾛｸﾞﾗﾑのﾊﾟｽ名のみを取得します...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' 既に存在するはずの 32 ﾋﾞｯﾄ ﾊﾞｰｼﾞｮﾝのﾌｧｲﾙを確認します。
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' ｴﾗｰ - ﾌｧｲﾙが見つかりません...
        Else
            GoTo SysInfoErr
        End If
    ' ｴﾗｰ - ﾚｼﾞｽﾄﾘ ｴﾝﾄﾘが見つかりません...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    On Error GoTo 0
    Exit Sub
SysInfoErr:
    MsgBox LoadResString(429), vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' ﾙｰﾌﾟ ｶｳﾝﾀ
    Dim rc As Long                                          ' 戻り値のｺｰﾄﾞ
    Dim hKey As Long                                        ' ｵｰﾌﾟﾝしたﾚｼﾞｽﾄﾘ ｷｰのﾊﾝﾄﾞﾙ
    Dim KeyValType As Long                                  ' ﾚｼﾞｽﾄﾘ ｷｰのﾃﾞｰﾀ型
    Dim tmpVal As String                                    ' ﾚｼﾞｽﾄﾘ ｷｰ値の一時保存領域
    Dim KeyValSize As Long                                  ' ﾚｼﾞｽﾄﾘ ｷｰ変数のｻｲｽﾞ
    '------------------------------------------------------------
    ' ﾙｰﾄ ｷｰ {HKEY_LOCAL_MACHINE...} にあるﾚｼﾞｽﾄﾘ ｷｰを開きます。
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ﾚｼﾞｽﾄﾘ ｷｰを開く
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ﾊﾝﾄﾞﾙ ｴﾗｰ...
    
    tmpVal = String$(1024, 0)                             ' 変数領域の割り当て
    KeyValSize = 1024                                       ' 変数のｻｲｽﾞを記憶
    
    '------------------------------------------------------------
    ' ﾚｼﾞｽﾄﾘ ｷｰ値を取得します...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' ｷｰ値の取得/作成
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ﾊﾝﾄﾞﾙ ｴﾗｰ
    
    tmpVal = Left(tmpVal, InStr(tmpVal, Chr(0)) - 1)
    '------------------------------------------------------------
    ' 変換のために、ｷｰ値の型を調べます...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' ﾃﾞｰﾀ型検索...
    Case REG_SZ                                             ' String ﾚｼﾞｽﾄﾘ ｷｰ ﾃﾞｰﾀ型
        KeyVal = tmpVal                                     ' String 値をｺﾋﾟｰ
    Case REG_DWORD                                          ' Double Word ﾚｼﾞｽﾄﾘ ｷｰ ﾃﾞｰﾀ型
        For i = Len(tmpVal) To 1 Step -1                    ' 各ﾋﾞｯﾄの変換
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Char ごとに値を作成
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Double Word を String に変換
    End Select
    
    GetKeyValue = True                                      ' 正常終了
    rc = RegCloseKey(hKey)                                  ' ﾚｼﾞｽﾄﾘ ｷｰをｸﾛｰｽﾞ
    Exit Function                                           ' 終了
    
GetKeyError:      ' ｴﾗｰ発生後の後始末...
    KeyVal = vbNullString                                            ' 戻り値の値を空文字列に設定
    GetKeyValue = False                                     ' 異常終了
    rc = RegCloseKey(hKey)                                  ' ﾚｼﾞｽﾄﾘ ｷｰをｸﾛｰｽﾞ
End Function

Private Sub Image1_Click()
End Sub




Private Sub lblTitle_Click()
data_kyoutou_to_katsunuma (0)
End Sub

Private Sub Picture16_Click()
    If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    End If


msg = "管理者用のパスワードを入力してください。"
res = InputBox(msg, "管理者パスワード")
Unload Me

If res = "kamifusafusa" Then
    frmkanrisha.Show 1
End If


End Sub

Private Sub Picture3_Click()

 mentefaile = mapa & "ezmanager_versionup.exe"
    
    If Dir(mentefaile) = "" Then
       MsgBox "メンテナンス用プログラムがありません。"
       Exit Sub
    End If
    
    Dim RetVal

    
    RetVal = Shell(mentefaile, 1)   ' 電卓を実行します。
    
    
End Sub
