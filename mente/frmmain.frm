VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "メンテナンス"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14745
   ControlBox      =   0   'False
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmmain.frx":0E42
   ScaleHeight     =   7230
   ScaleWidth      =   14745
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1590
      MouseIcon       =   "frmmain.frx":1C84
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmmain.frx":2583
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   6210
      Width           =   480
   End
   Begin VB.PictureBox Picture16 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   480
      MouseIcon       =   "frmmain.frx":2E4D
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmmain.frx":374C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   6180
      Width           =   480
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   60
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame20 
      Caption         =   "メンテナンス内容の選択"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6900
      Left            =   2610
      TabIndex        =   0
      ToolTipText     =   "テキストボックスに、パスを指定してください。"
      Top             =   150
      Width           =   5175
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   315
         Left            =   4620
         TabIndex        =   25
         Top             =   6420
         Width           =   315
      End
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   420
         TabIndex        =   23
         Top             =   6420
         Width           =   4095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "指定した暗号化データを解凍します。"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   330
         TabIndex        =   22
         Top             =   5520
         Width           =   4035
      End
      Begin VB.Frame Frame2 
         Height          =   1665
         Left            =   360
         TabIndex        =   13
         Top             =   2460
         Width           =   4695
         Begin VB.OptionButton o2 
            Caption         =   "２．EZ Manager メンテナンスシステム"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   21
            Top             =   1020
            Width           =   4185
         End
         Begin VB.OptionButton o2 
            Caption         =   "１．EZ Manager メインシステム"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   3525
         End
         Begin VB.Label lblver 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1470
            TabIndex        =   19
            Top             =   630
            Width           =   1215
         End
         Begin VB.Label lblver 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1470
            TabIndex        =   18
            Top             =   1305
            Width           =   1215
         End
         Begin VB.Label lblver 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   3390
            TabIndex        =   17
            Top             =   630
            Width           =   1215
         End
         Begin VB.Label lblver 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   3390
            TabIndex        =   16
            Top             =   1305
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000B&
            Caption         =   "→"
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
            Index           =   0
            Left            =   2880
            TabIndex        =   15
            Top             =   660
            Width           =   285
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000B&
            Caption         =   "→"
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
            Index           =   1
            Left            =   2880
            TabIndex        =   14
            Top             =   1335
            Width           =   285
         End
      End
      Begin VB.TextBox txtpath2 
         Height          =   345
         Left            =   360
         TabIndex        =   12
         Top             =   4950
         Width           =   4575
      End
      Begin VB.TextBox txtpath 
         Height          =   345
         Left            =   360
         TabIndex        =   11
         Top             =   4560
         Width           =   4575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "データをサーバに転送を実行します。"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   330
         TabIndex        =   10
         Top             =   4230
         Width           =   4035
      End
      Begin VB.OptionButton Option1 
         Caption         =   "最新版へのバージョンアップを実行します。"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   330
         TabIndex        =   4
         Top             =   1680
         Value           =   -1  'True
         Width           =   4035
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "暗号化されたファイルのあるフォルダーを指定してください。同じフォルダーに復元します。"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   420
         TabIndex        =   24
         Top             =   5820
         Width           =   4545
      End
      Begin VB.Label lblzenkai 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   390
         TabIndex        =   9
         Top             =   2160
         Width           =   4545
      End
      Begin VB.Label Label33 
         Caption         =   "「EZManager」をのメンテナンスを実行します。下記のどちらかを選んで実行してください。"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   765
         Left            =   330
         TabIndex        =   1
         Top             =   540
         Width           =   4665
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   6915
      Left            =   7950
      TabIndex        =   2
      Top             =   150
      Width           =   6615
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6300
         Left            =   210
         TabIndex        =   3
         Top             =   300
         Width           =   6165
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ｷｬﾝｾﾙ"
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
      Height          =   225
      Index           =   10
      Left            =   1470
      TabIndex        =   8
      Top             =   6690
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "実行"
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
      Height          =   255
      Index           =   14
      Left            =   210
      TabIndex        =   7
      Top             =   6690
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   2520
      Left            =   150
      Picture         =   "frmmain.frx":458E
      Stretch         =   -1  'True
      Top             =   390
      Width           =   2220
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   975
      Left            =   210
      Top             =   6090
      Width           =   2085
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'噴出し

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" _
        Alias "Shell_NotifyIconA" _
        (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'構造体宣言
Private Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 128
        dwState As Long
        dwStateMask As Long
        szInfo As String * 256
        uTimeoutOrVersion As Long
        szInfoTitle As String * 64
        dwInfoFlags As Long
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_INFO = &H10

'マウス操作
Private Const WM_MOUSEMOVE = &H200
Private Const WM_USER = &H400
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)



'圧縮のためのＡＰＩです。
Private Declare Function Unlha Lib "UNLHA32.DLL" (ByVal hWnd As Long, _
                                          ByVal aCommand As String, _
                                          ByVal BuffName As String, _
                                          ByVal dSize As Long) As Long
                                          
'Shell_NotifyIconで使用する構造体
'Private Type NOTIFYICONDATA
'        cbSize As Long          '構造体のｻｲｽﾞ
'        hWnd As Long            'ﾒｯｾｰｼﾞを受け取るｳｨﾝﾄﾞｳのﾊﾝﾄﾞﾙ
'        uID As Long
'        uFlags As Long          'ﾌﾗｸﾞ
'        uCallbackMessage As Long 'ｺｰﾙﾊﾞｯｸﾒｯｾｰｼﾞ
'        hIcon As Long           '表示するｱｲｺﾝ
'        szTip(64) As Byte       'ﾂｰﾙﾁｯﾌﾟに表示する文字列
        'VB5.0の場合は既存通り
        'szTip As String * 64
        'で良い
'End Type

'ﾀｽｸﾄﾚｲのｱｲｺﾝを操作するAPI
'Private Declare Function Shell_NotifyIcon Lib "shell32.dll" _
        Alias "Shell_NotifyIconA" _
        (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'NOTIFYICONDATAで使用する定数
'Private Const NIM_ADD = &H0      'ｱｲｺﾝを追加
'Private Const NIM_MODIFY = &H1   '情報を更新
'Private Const NIM_DELETE = &H2   'ｱｲｺﾝを削除

'Shell_NotifyIconで使用する定数
'Private Const NIF_ICON = &H2     'ｱｲｺﾝを表示
'Private Const NIF_MESSAGE = &H1  'ﾒｯｾｰｼﾞを返す
'Private Const NIF_TIP = &H4      'ﾂｰﾙﾁｯﾌﾟを表示

'WMﾒｯｾｰｼﾞ
'Private Const WM_MOUSEMOVE = &H200       'ﾏｳｽが動いた
Private Const WM_LBUTTONDOWN = &H201     '左ﾎﾞﾀﾝﾀﾞｳﾝ
'Private Const WM_LBUTTONUP = &H202       '左ﾎﾞﾀﾝｱｯﾌﾟ
'Private Const WM_LBUTTONDBLCLK = &H203   '左ﾀﾞﾌﾞﾙｸﾘｯｸ
'Private Const WM_RBUTTONDOWN = &H204     '右ﾎﾞﾀﾝﾀﾞｳﾝ
'Private Const WM_RBUTTONUP = &H205       '右ﾎﾞﾀﾝｱｯﾌﾟ
'Private Const WM_RBUTTONDBLCLK = &H206   '右ﾀﾞﾌﾞﾙｸﾘｯｸ

'Windowを作成したｽﾚｯﾄﾞを前景に置いてｱｸﾃｨﾌﾞにするAPI
Private Declare Function SetForegroundWindow Lib "user32" _
        (ByVal hWnd As Long) As Long

                                          
                                          

'[フォルダの参照]ダイアログを表示する関数
'Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
   (lpbi As BROWSEINFO) As Long
'ITEMIDLISTをパスに変換する関数
'Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
   (ByVal pidl As Long, _
    ByVal pszPath As String) As Long

'メモリブロックを解放する関数
'Private Declare Sub CoTaskMemFree Lib "ole32.dll" _
   (pv As Any)
   
   
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBROWSEINFO As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)



'[フォルダの参照]ダイアログ情報
Private Type BROWSEINFO
    hWndOwner      As Long
    pidlRoot       As Long
    pszDisplayName As String
    lpszTitle      As String
    ulFlags        As Long
    lpfn           As Long
    lParam         As Long
    iImage         As Long
End Type

Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer
    dwStrucVersionh As Integer
    dwFileVersionMSl As Integer
    dwFileVersionMSh As Integer
    dwFileVersionLSl As Integer
    dwFileVersionLSh As Integer
    dwProductVersionMSl As Integer
    dwProductVersionMSh As Integer
    dwProductVersionLSl As Integer
    dwProductVersionLSh As Integer
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type

Private Type CODEPAGE
    lngLOW As Integer
    lngHIGH As Integer
End Type

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" ( _
    ByVal lptstrFilename As String, _
    ByVal dwHandle As Long, _
    ByVal dwLen As Long, _
    lpData As Any) As Long

Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" ( _
    ByVal lptstrFilename As String, _
    lpdwHandle As Long) As Long

Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" ( _
    pBlock As Any, _
    ByVal lpSubBlock As String, _
    lplpBuffer As Any, _
    puLen As Long) As Long

Private Declare Sub MoveMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    dest As Any, _
    ByVal Source As Long, _
    ByVal length As Long)

Private cnn_r As ADODB.Connection

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const CSIDL_NETWORK = 0 ' &H12
Private Const MAX_PATH2 = 260
Private Const MAX_PATH = 300
Private daijoubu As Integer
Private strFileName(3) As String
Private strFileName2(3) As String
Private verpath As String
Private kidoupassword As String
Private settei_file As String
Private mapa As String
Private settei_res
Private kidou_type2 As Integer

Private kidou_type As Integer
'大域変数の宣言
Private nidNoftyIconData As NOTIFYICONDATA
Private ingAPIReVal As Long     'APIの戻り値を格納

Private cCrypt    As clsCryptFile
Private kidou_surukai As Integer
Private lblver2txt(2) As String











Function FcSQlGet3(hajime As Integer, PrRS As ADODB.Recordset, PrSQL As String, PrMsg As String) As Boolean
If hajime = 1 Then
    data_r_open
End If

    'ＳＱＬ文の実行
    With PrRS
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
    End With
    On Error GoTo ErrSyori
        'Set PrRS = db_r.OpenRecordset(PrSQL, dbOpenDynaset, dbReadOnly)
        'Set PrRS = cnn_r.Execute(PrSQL)
            Set PrRS.ActiveConnection = cnn_r
            PrRS.Open PrSQL
    On Error GoTo 0
    
    '結果の確認
    If PrRS.RecordCount = 0 Then
        'PrMsg = LoadResString(491)
        FcSQlGet3 = False
        Exit Function
    'Else
        'PrRS.MoveLast
        'PrRS.MoveFirst
    End If
    
    FcSQlGet3 = True
    Exit Function
    
ErrSyori:
    'PrMsg = LoadResString(217)
    FcSQlGet3 = False

End Function
Function angouka(angoukamae As String, angoukago As String, kaijo_key As String) As Integer

    
    ' 暗号化
    Set cCrypt = New clsCryptFile
    With cCrypt
        .InFile = Trim(angoukamae)
        .OutFile = Trim(angoukago)
        .Password = kaijo_key
        If .Encrypt Then
            angouka = 1
            Exit Function
        Else
            angouka = 0
            Exit Function
        End If
    End With

End Function
Function fukugen(angoukamae As String, angoukago As String, kaijo_key As String) As Integer

    
    '復元
    Set cCrypt = New clsCryptFile
    With cCrypt
        .InFile = Trim(angoukamae)
        .OutFile = Trim(angoukago)
        .Password = kaijo_key
        If .Decrypt Then
            fukugen = 1
            Exit Function
        Else
            fukugen = 0
            Exit Function
        End If
    End With

End Function

Function asshuku(asshuku_mae As String, asshuku_go As String) As Integer
Dim StrBuff As String  '' * 15885
Dim i
Dim LZHString As String
    'MAIN.MDBをMAINDB.LZHに圧縮する。
      LZHString = "A " & Chr(34) & asshuku_go & Chr(34) & Space(1) & Chr(34) & asshuku_mae & Chr(34)
      i = Unlha(frmmain.hWnd, LZHString, StrBuff, Len(StrBuff))
      If i = 0 Then  '成功
            
            asshuku = 0
      Else    '失敗
          asshuku = -1
      End If

End Function


Sub data_r_open()
            Set cnn_r = New ADODB.Connection
            cnn_r.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & settei_file & _
            ";Persist Security Info=False;Jet OLEDB:Database Password=" & kidoupassword
            cnn_r.Open

End Sub


Function setting(hajime As Integer, id As Long, docchi As Integer, newid As String, owari As Integer) As String
Dim sql_set As String, rs_set As New ADODB.Recordset


If hajime = 1 Then
    'データのロード
    'kidoupassword = ";pwd=" & "kamifusafusa"
    'Set DB = OpenDatabase(VBNAME)
    On Error GoTo errsettingkidou
        'Set db_r = DBEngine.Workspaces(0).OpenDatabase(settei_file, False, False, kidoupassword)
        'ADO
            data_r_open
    On Error GoTo 0
End If

If docchi = 0 Then '読み込み
    sql_set = "select * from settei where id ='1'"
    If FcSQlGet3(0, rs_set, sql_set, "") = False Then
        setting = "-1"
        Exit Function
    Else
        Select Case id
        Case 1
            setting = CStr(rs_set!s1)
        Case 2
            setting = CStr(rs_set!s2)
        Case 3
            setting = CStr(rs_set!s3)
        Case 4
            setting = CStr(rs_set!s4)
        Case 5
            setting = CStr(rs_set!s5)
        Case 6
            setting = CStr(rs_set!s6)
        Case 7
            setting = CStr(rs_set!s7)
        Case 8
            setting = CStr(rs_set!s8)
        Case 9
            setting = CStr(rs_set!s9)
        Case 10
            setting = CStr(rs_set!s10)
        Case 11
            setting = CStr(rs_set!s11)
        Case 12
            setting = CStr(rs_set!s12)
        Case 13
            setting = CStr(rs_set!s13)
        Case 14
            setting = CStr(rs_set!s14)
        Case 15
            setting = CStr(rs_set!s15)
        Case 16
            setting = CStr(rs_set!s16)
        Case 17
            setting = CStr(rs_set!s17)
        Case 18
            setting = CStr(rs_set!s18)
        Case 19
            setting = CStr(rs_set!s19)
        Case 20 '履歴の数
            setting = CStr(rs_set!s20)
        End Select
        rs_set.Close
    End If
    
Else '書き込み
    'sql_setstr = "s" & CStr(id)
    'sql_set = "UPDATE settei SET " & sql_setstr & " = '" & newid & "'" & _
                          " WHERE id = '1'"
    On Error GoTo errrirekikaki2
        'db_r.Execute sql_set, dbSQLPassThrough
        ' Open recordset with data from settei table.
        rs_set.CursorType = adOpenKeyset
        rs_set.LockType = adLockOptimistic
        Set rs_set.ActiveConnection = cnn_r
        rs_set.Open "SELECT * FROM settei where id='1'"

        'rs_set.Open "SELECT * FROM settei where id='1'", cnn_r, , adCmdText
        Select Case id
        Case 1
            rs_set!s1 = newid
        Case 2
            rs_set!s2 = newid
        Case 3
            rs_set!s3 = newid
        Case 4
            rs_set!s4 = newid
        Case 5
            rs_set!s5 = newid
        Case 6
            rs_set!s6 = newid
        Case 7
            rs_set!s7 = newid
        Case 8
            rs_set!s8 = newid
        Case 9
            rs_set!s9 = newid
        Case 10
            rs_set!s10 = newid
        Case 11
            rs_set!s11 = newid
        Case 12
            rs_set!s12 = newid
        Case 13
            rs_set!s13 = newid
        Case 14
            rs_set!s14 = newid
        Case 15
            rs_set!s15 = newid
        Case 16
            rs_set!s16 = newid
        Case 17
            rs_set!s17 = newid
        Case 18
            rs_set!s18 = newid
        Case 19
            rs_set!s19 = newid
        Case 20
            rs_set!s20 = newid
        End Select
        rs_set.Update
    
    On Error GoTo 0
    
End If

If owari = 1 Then
    'db_r.Close
    cnn_r.Close
End If

Exit Function

errsettingkidou:
    setting = "-1"
    Exit Function
errrirekikaki2:
    setting = "-1"
    'rs_set.CancelUpdate
    'cnn_r.RollbackTrans
    Exit Function

End Function

Private Sub Command1_Click()
    Dim udtBrowseInfo As BROWSEINFO
    Dim lngIDList     As Long
    Dim strPath       As String
    
    '[フォルダの参照]ダイアログ情報を設定
    With udtBrowseInfo
        .hWndOwner = Me.hWnd
        .pidlRoot = CSIDL_NETWORK
        .pszDisplayName = String(MAX_PATH2, vbNullChar)
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    '[フォルダの参照]ダイアログを表示
    lngIDList = SHBrowseForFolder(udtBrowseInfo)
    If lngIDList <> 0 Then
        strPath = String(MAX_PATH, vbNullChar)
        'ITEMIDLISTをパスに変換
        SHGetPathFromIDList lngIDList, strPath
        If Left(strPath, 1) = vbNullChar Then
            'パスが空の場合はコンピュータ名を取得
            With udtBrowseInfo
                strPath = "\\" & Left(.pszDisplayName, _
                                      InStr(.pszDisplayName, _
                                            vbNullChar) - 1)
            End With
        Else
            strPath = Left(strPath, _
                           InStr(strPath, _
                                 vbNullChar) - 1)
        End If
        '結果をラベルに表示
        Text1.Text = strPath
        'ITEMIDLISTを解放
        CoTaskMemFree lngIDList
    End If

End Sub

Private Sub Form_Load()

Dim ret

Dim cmdline() As String


    cmdline = Split(Command(), " ") 'スペースで切り分けてコマンドライン引数を配列に格納
    
    Dim ArCount As Integer
    ArCount = UBound(cmdline) '配列サイズを求める
    
    kidou_type = 0
    
    
    Dim i As Integer
    For i = 0 To ArCount
        If cmdline(i) = "-u" Then
            kidou_type = -1
            Exit For
        End If
        If cmdline(i) = "-i" Then
            kidou_type2 = -1
            Exit For
        End If
    Next i
    

If kidou_type = -1 Or kidou_type2 = -1 Then

        Dim trayicon As NOTIFYICONDATA
        
        
        '-------------タスクトレイへ表示-------------
        With trayicon
        .cbSize = Len(trayicon)
        .hWnd = frmmain.hWnd
        .uID = 0&
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = frmmain.Icon
        .szTip = "Cheveux 2007 メンテナンス" & Chr(0)
        End With
        ret = Shell_NotifyIcon(NIM_ADD, trayicon)
    '-------------タスクトレイへ表示END-------------



    'さあ、表示しよう○(^o^)。
    ingAPIReVal = Shell_NotifyIcon(NIM_ADD, nidNoftyIconData)



    'フォームを非表示にする
    frmmain.Visible = False


    If kidou_type2 = -1 Then
        bb_hyouji ("システムが最新かをチェック中・・")
    End If

    
End If






verget (0)


    If kidou_surukai = 0 Then
           'タスクトレイのアイコンを消去
        
        '-------------タスクトレイから削除-------------
        With trayicon
        .cbSize = Len(trayicon)
        .hWnd = frmmain.hWnd
        .uID = 0&
        End With
        
        ret = Shell_NotifyIcon(NIM_DELETE, trayicon)
        '-------------タスクトレイから削除END-------------

        End
    End If



'パスワード
    kidoupassword = "kamifusafusa"
    
mapa = App.Path
If Right(mapa, 1) <> "\" Then
    mapa = mapa & "\"
End If
    
    

'データベースファイルのチェック
    settei_file = mapa & "settei.mdb"

If Dir$(settei_file) = vbNullString Then
   MsgBox "設定ファイルがありません。"
   End
End If

settei_res = setting(1, 4, 0, "", 1)
If settei_res = "-1" Then
    MsgBox "設定の読み出しに失敗しました。004"
 
    End
ElseIf settei_res = "0" Then
    MsgBox "データﾍﾞースの設定がありません。"
    End
Else
    txtpath.Text = settei_res & "\"
End If

Dim root_pass As String

Dim objwsh As Object
Set objwsh = CreateObject("Wscript.Shell")
root_pass = objwsh.SpecialFolders("MyDocuments")

If Right(root_pass, 1) <> "\" Then
   root_pass = root_pass & "\"
End If
Set objwsh = Nothing


txtpath2.Text = root_pass



If kidou_type = -1 Then
    Option1(1).Value = True
    Picture16_Click
End If

If kidou_type2 = -1 Then
    Option1(0).Value = True
    Picture16_Click
End If

End Sub
Sub verget(no As Integer)

Dim lngRet As Long
Dim lngDummy As Long
Dim bBuffer() As Byte
Dim lngLen As Long
Dim lpBuffer As Long
Dim ffi As VS_FIXEDFILEINFO
Dim i As Integer, lblver2(2) As String

' strFileName に取得したいファイル名をセット
If no = 0 Then
    strFileName(0) = App.Path & "\catv.exe"
    strFileName(1) = App.Path & "\ezmanager_ment.exe"
    strFileName(2) = App.Path & "\ezmanager_sub.exe"
Else
    strFileName(0) = verpath & "\catv.exe"
    strFileName(1) = verpath & "\ezmanager_ment.exe"
    strFileName(2) = verpath & "\ezmanager_sub.exe"
End If
strFileName2(0) = "catv.exe"
strFileName2(1) = "ezmanager_ment.exe"
strFileName2(2) = "ezmanager_sub.exe"

For i = 0 To 2
        ' サイズを取得
        lngLen = GetFileVersionInfoSize(strFileName(i), lngDummy)
        If lngLen < 1 Then
            lblver(i).Caption = "Nothing"
        Else
            ' バイトの配列の領域取得
            ReDim bBuffer(lngLen)
            
            ' ファイル バージョン情報を取得
            lngRet = GetFileVersionInfo(strFileName(i), 0&, lngLen, bBuffer(0))
            lngRet = VerQueryValue(bBuffer(0), "\", lpBuffer, lngLen)
            
            ' バイトの処理
            MoveMemory ffi, lpBuffer, Len(ffi)
            
            ' ファイル バージョン
                lblver(i + no).Caption = "Ver." & _
                Format$(ffi.dwFileVersionMSh) & "." & _
                Format$(ffi.dwFileVersionMSl) & "." & _
                Format$(ffi.dwFileVersionLSl)
                lblver2(i) = Format$(ffi.dwFileVersionMSh) & _
                Format$(ffi.dwFileVersionMSl, "00") & _
                Format$(ffi.dwFileVersionLSl, "00")
        End If
        
Next i


kidou_surukai = 0
If kidou_type2 = -1 Then
        If txt_ver_chk = -1 Then
            Unload frmmain
            End
        End If
        
        If lblver2txt(0) >= lblver2(0) Then
            kidou_surukai = 1
        Else
            If lblver2txt(1) > lblver2(1) Then
                kidou_surukai = 1
            End If
        End If
Else
     kidou_surukai = 1
End If

'MsgBox kidou_surukai

End Sub
Function txt_ver_chk() As Integer

Dim servertxt As String, windowsbetsupath As String, setstr As String
Dim ipi As Integer, mojimoji As String, loginname As String, loginpassword As String
Dim root_pass As String, root_pass2 As String

Dim objwsh As Object
Set objwsh = CreateObject("Wscript.Shell")
root_pass = objwsh.SpecialFolders("MyDocuments")

If Right(root_pass, 1) <> "\" Then
   root_pass = root_pass & "\"
End If
Set objwsh = Nothing

root_pass2 = root_pass & "ver.txt"

If Dir(root_pass2) <> "" Then
    Kill root_pass2
End If
    
 servertxt = "ftp://plot-office.com"
windowsbetsupath = "CD www/versionup/ezmanager/new7"
 loginname = "plot-office"
 loginpassword = "Plot8877Ken"




With frmmain.Inet1
On Error GoTo errinet
        .URL = servertxt
        .UserName = loginname
        .Password = loginpassword
        .RequestTimeout = 15
        .Execute , windowsbetsupath
        Do While frmmain.Inet1.StillExecuting = True
            DoEvents
        Loop
        If daijoubu = -1 Then
            txt_ver_chk = -1
            .Execute , "close"
            Screen.MousePointer = 0
            Exit Function
        End If
        setstr = "GET " & "ver.txt" & Space(1) & root_pass2
        .Execute , setstr
        
        Do While frmmain.Inet1.StillExecuting = True
            DoEvents
        Loop
        If daijoubu = -1 Then
            txt_ver_chk = -1
            .Execute , "close"
            Screen.MousePointer = 0
            Exit Function
        End If
            .Execute , "close"
On Error GoTo 0
End With
lblver2txt(0) = ""
lblver2txt(1) = ""
mojimoji = ""
ipi = 0

If Dir(root_pass2) = "" Then
    txt_ver_chk = -1
    Screen.MousePointer = 0
    Exit Function
End If
   Open root_pass & "ver.txt" For Input Lock Read Write As #1
    While Not EOF(1)
        Line Input #1, mojimoji
        ipi = ipi + 1
    Wend
    Close #1

lblver2txt(0) = Mid(Trim(mojimoji), 1, 5)
lblver2txt(1) = Mid(Trim(mojimoji), 6, 5)


 txt_ver_chk = 0
 Exit Function
 
errinet:
            txt_ver_chk = -1
            Screen.MousePointer = 0
            Exit Function
 
 
End Function

Private Sub Inet1_StateChanged(ByVal State As Integer)
Dim strMess As String
Select Case State

Case icError
    ' エラーを説明するテキストを取得します。
      strMess = "ErrorCode: " & Inet1.ResponseCode & _
      " : " & Inet1.ResponseInfo
      MsgBox strMess
      If daijoubu = 1 Then
        MsgBox "違うサーバを選んで再度実行してください。何度か実行しても同じエラーが出るときは、サーバが止まっている可能性があります。少ししてから実行してください。"
      daijoubu = -1
      End If
Case icResponseCompleted


End Select
Screen.MousePointer = 0
End Sub


Sub listmsg(no As Integer, dochi As Integer, Optional hozo As String)

Dim mmsg As String

Select Case no
Case -1
    mmsg = "エラーが発生しました。中止します。"
Case 0
    mmsg = "インターネットに接続を開始します。"
Case 1
    mmsg = "最新版へのバージョンアップを開始します。"
Case 2
    If dochi = 0 Then
        mmsg = "FTPサーバ[plot.dip.jp]に接続を開始します。"
    Else
        mmsg = "FTPサーバ[plot-office.com]に接続を開始します。"
    End If
Case 3
    mmsg = "FTPサーバ[ezmanager/new]に移動します。"
Case 4
    mmsg = "FTPサーバ[ezmanager/new]に移動中..."
Case 5
    mmsg = "FTPサーバ[ezmanager/new]に移動完了"
Case 6
    Select Case dochi
    Case 0
        mmsg = "EZManagerメインシステムの最新版のローカルへのコピーを開始します。"
    Case 1
        mmsg = "EZManagerメンテナンスシステムの最新版のローカルへのコピーを開始します。"
    End Select
Case 7
    Select Case dochi
    Case 0
        mmsg = "EZManagerメインシステムの最新版のローカルへのコピーを開始中..."
    Case 1
        mmsg = "EZManagerメンテナンスシステムの最新版のローカルへのコピーを開始中..."
    End Select
Case 8
    Select Case dochi
    Case 0
        mmsg = "EZManagerメインシステムの最新版のローカルへのコピーが終了しました｡"
    Case 1
        mmsg = "EZManagerメンテナンスシステムの最新版のローカルへのコピーが終了しました｡"
    End Select
Case 9
    If dochi = 0 Then
        mmsg = "FTPサーバ[plot-office.com]に接続を終了しました｡"
    Else
        mmsg = "FTPサーバ[plot.dip.jp]に接続を終了しました｡"
    End If
Case 10
    mmsg = "FTPサーバ[ezmanager/data/katsunuma_catv]に移動します。"
Case 11
    mmsg = "FTPサーバ[ezmanager/data/katsunuma_catv]に移動中..."
Case 12
    mmsg = "FTPサーバ[ezmanager/data/katsunuma_catv]に移動完了"
Case 13
    mmsg = "FTPサーバ[ezmanager/data/katsunuma_catv]に" & hozo & "フォルダーを作成します"
Case 14
    mmsg = "FTPサーバ[ezmanager/data/katsunuma_catv]に" & hozo & "フォルダーを作成しました。"
Case 15
    mmsg = "FTPサーバ[ezmanager/data/katsunuma_catv" & hozo & "]に移動します。"
Case 16
    mmsg = "FTPサーバ[ezmanager/data/katsunuma_catv" & hozo & "]に移動中..."
Case 17
    mmsg = "FTPサーバ[ezmanager/data/katsunuma_catv" & hozo & "]に移動完了"
Case 18
    mmsg = "EZManagerのデータ[" & hozo & "]をサーバへのコピーを開始します。"
Case 19
    mmsg = "EZManagerのデータ[" & hozo & "]をサーバへのコピーを開始中..."
Case 20
    mmsg = "EZManagerのデータ[" & hozo & "]をサーバへのコピーが終了しました。"
Case 21
    mmsg = "EZManagerのデータ[" & hozo & "]の転送準備を開始します。"
Case 22
    mmsg = "EZManagerのデータ[" & hozo & "]の転送準備が完了しました。"



End Select

List1.AddItem mmsg

End Sub
Sub maindata_plot_tensou_jidou()
Dim appath(4) As String, servertxt As String, appath2(4) As String, appath3(4) As String
Dim i As Integer, hozonfile As String, copymotofile(0) As String
Dim angouka_key As String

Dim sspath As String, sspath2 As String

sspath = Trim(txtpath.Text)
If sspath = "" Then
    MsgBox "コピー元のパスが入力されていません。"
    Exit Sub
End If
sspath2 = Trim(txtpath2.Text)
If sspath2 = "" Then
    MsgBox "コピー先のパスが入力されていません。"
    Exit Sub
End If

Screen.MousePointer = 11
Picture2.Enabled = False


appath(0) = sspath & "*.*"
appath2(0) = "all.lzh"
appath3(0) = sspath2 & "all.lzh"
angouka_key = "kamifusafusa"
If Dir(appath3(0)) <> "" Then
    Kill appath3(0)
    DoEvents
End If

If asshuku(appath(0), appath3(0)) = -1 Then
    MsgBox "圧縮時にエラーが発生しました。終了します。"
    Exit Sub
End If

'ファイル削除




On Error GoTo errinet2

With Inet1

                 

        hozonfile = Format(Date, "yyyymmdd") & "-" & Format(Time, "hmm")
        .URL = "ftp://plot-office.com"
        .UserName = "eplot-office"
        .Password = "Plot8877Ken"
        listmsg 2, 0
        .RequestTimeout = 15
        daijoubu = 1
        .Execute , "CD www/backup/ezmanager/"
        listmsg 10, 0
        listmsg 11, 0
        Do While Inet1.StillExecuting = True
            DoEvents
        Loop
        If daijoubu = -1 Then
            listmsg -1, 0
            Picture2.Enabled = True
            .Execute , "close"
            Exit Sub
        End If
        listmsg 12, 0
        daijoubu = 4
        listmsg 13, 0, hozonfile
        .Execute , "MKDIR " & hozonfile
        Do While Inet1.StillExecuting = True
            DoEvents
        Loop
        If daijoubu = -1 Then
            listmsg -1, 0
            Picture2.Enabled = True
            .Execute , "close"
            Exit Sub
        End If
        listmsg 14, 0, hozonfile
        daijoubu = 1
                        
        .Execute , "CD " & hozonfile
        listmsg 15, 0, hozonfile
        listmsg 16, 0, hozonfile
        Do While Inet1.StillExecuting = True
            DoEvents
        Loop
        If daijoubu = -1 Then
            listmsg -1, 0
            Picture2.Enabled = True
            .Execute , "close"
            Exit Sub
        End If
        listmsg 17, 0, hozonfile
            If Dir(appath3(0)) <> "" Then
                 
                 daijoubu = 5
                 
                  listmsg 18, 0, appath2(0)
                 copymotofile(0) = "PUT " & appath3(0) & Space(1) & appath2(0)
                 .Execute , copymotofile(0)
                 listmsg 19, 0, angouka_go_filename
                 
                 Do While Inet1.StillExecuting = True
                     DoEvents
                 Loop
                 If daijoubu = -1 Then
                     listmsg -1, 0
                     Picture2.Enabled = True
                     .Execute , "close"
                     Screen.MousePointer = 0
                     Exit Sub
                 End If
                 daijoubu = 3
                 listmsg 20, 0, appath2(0)
            End If
        
        .Execute , "close"
        Do While Inet1.StillExecuting = True
            DoEvents
        Loop
            listmsg 9, 1
        Screen.MousePointer = 0


On Error GoTo 0
End With
        Unload frmmain
        End

Exit Sub

errinet2:
    Screen.MousePointer = 0
    
    MsgBox "エラーが発生しました。"
        Unload frmmain
        End


End Sub
Sub maindata_plot_tensou()
Dim appath(4) As String, servertxt As String, appath2(4) As String, appath3(4) As String
Dim i As Integer, hozonfile As String, copymotofile(0) As String

Dim sspath As String, sspath2 As String

sspath = Trim(txtpath.Text)
If sspath = "" Then
    MsgBox "コピー元のパスが入力されていません。"
    Exit Sub
End If
sspath2 = Trim(txtpath2.Text)
If sspath2 = "" Then
    MsgBox "コピー先のパスが入力されていません。"
    Exit Sub
End If

Screen.MousePointer = 11
Picture2.Enabled = False


appath(0) = sspath & "*.*"
appath2(0) = "all.lzh"
appath3(0) = sspath2 & "all.lzh"
If Dir(appath3(0)) <> "" Then
    Kill appath3(0)
    DoEvents
End If
If asshuku(appath(0), appath3(0)) = -1 Then
    MsgBox "圧縮時にエラーが発生しました。終了します。"
    Exit Sub
End If

On Error GoTo errinet2

With Inet1

                 

        hozonfile = Format(Date, "yyyymmdd") & "-" & Format(Time, "hmm")
        .URL = "ftp://plot-office.com"
        .UserName = "plot-office"
        .Password = "Plot8877Ken"
        listmsg 2, 0
        .RequestTimeout = 15
        daijoubu = 1
        .Execute , "CD www/backup/ezmanager/"
        listmsg 10, 0
        listmsg 11, 0
        Do While Inet1.StillExecuting = True
            DoEvents
        Loop
        If daijoubu = -1 Then
            listmsg -1, 0
            Picture2.Enabled = True
            .Execute , "close"
            Exit Sub
        End If
        listmsg 12, 0
        daijoubu = 4
        listmsg 13, 0, hozonfile
        .Execute , "MKDIR " & hozonfile
        Do While Inet1.StillExecuting = True
            DoEvents
        Loop
        If daijoubu = -1 Then
            listmsg -1, 0
            Picture2.Enabled = True
            .Execute , "close"
            Exit Sub
        End If
        listmsg 14, 0, hozonfile
        daijoubu = 1
                        
       ' .Execute , "CD " & hozonfile
        listmsg 15, 0, hozonfile
        listmsg 16, 0, hozonfile
        Do While Inet1.StillExecuting = True
            DoEvents
        Loop
        If daijoubu = -1 Then
            listmsg -1, 0
            Picture2.Enabled = True
            .Execute , "close"
            Exit Sub
        End If
        listmsg 17, 0, hozonfile
            If Dir(appath3(0)) <> "" Then
                 
                 daijoubu = 5
                 
                  listmsg 18, 0, appath2(0)
                 copymotofile(0) = "PUT " & appath3(0) & Space(1) & appath2(0)
                 .Execute , copymotofile(0)
                 listmsg 19, 0, appath2(0)
                 
                 Do While Inet1.StillExecuting = True
                     DoEvents
                 Loop
                 If daijoubu = -1 Then
                     listmsg -1, 0
                     Picture2.Enabled = True
                     .Execute , "close"
                     Screen.MousePointer = 0
                     Exit Sub
                 End If
                 daijoubu = 3
                 listmsg 20, 0, appath2(0)
            End If
        
        .Execute , "close"
        Do While Inet1.StillExecuting = True
            DoEvents
        Loop
            listmsg 9, 1
        Screen.MousePointer = 0
        MsgBox "完了しました。"


On Error GoTo 0
End With
Picture2.Enabled = True

Exit Sub

errinet2:
    Screen.MousePointer = 0
    
    MsgBox "エラーが発生しました。"
    Picture2.Enabled = True
    Exit Sub


End Sub

Sub bb_hyouji(shiti_msg As String)

Dim ret As Long
Dim trayicon As NOTIFYICONDATA

'-------------バルーン表示-------------
With trayicon
.cbSize = Len(trayicon)
.uFlags = NIF_INFO
.hWnd = frmmain.hWnd
.szInfoTitle = "Ez Manager2 メンテナンス" & Chr(0)
.szInfo = shiti_msg & Chr(0)
.dwInfoFlags = 1
End With

ret = Shell_NotifyIcon(NIM_MODIFY, trayicon)
'-------------バルーン表示END-------------


End Sub


Private Sub Picture16_Click()

Dim copymotofile(2) As String, copysakifile(2) As String, temsofile(2) As String, hozonfile As String
Dim appath(3) As String, servertxt As String, appath2(3) As String
Dim i As Integer, setstr As String, da1file As String, da2file As String, da3file As String
Dim da1file2 As String, da2file2 As String
Dim trayicon As NOTIFYICONDATA
Dim ret As Long, root_pass As String

Dim objwsh As Object
Set objwsh = CreateObject("Wscript.Shell")
root_pass = objwsh.SpecialFolders("MyDocuments")

If Right(root_pass, 1) <> "\" Then
   root_pass = root_pass & "\"
End If
Set objwsh = Nothing


If Option1(2).Value = True Then
        '復元
        Dim moto_file As String, go_file As String
        If Trim(Text1.Text) = "" Then
            MsgBox "復元元の場所が選択されていません。"
            Exit Sub
        End If
        moto_file = Text1.Text & "\all.lzh3"
        If Dir(moto_file) = "" Then
            MsgBox "復元元のファイルがみつかりません。"
            Exit Sub
        End If
        go_file = Text1.Text & "\all.lzh"
        If Dir(go_file) <> "" Then
            Kill go_file
            DoEvents
        End If
        If fukugen(moto_file, go_file, "kamifusafusa") = 0 Then
            MsgBox "復元に失敗しました。"
            Exit Sub
        Else
            MsgBox "復元しました。"
            Exit Sub
        End If
        Exit Sub
End If

If Option1(1).Value = True Then
    If kidou_type = -1 Then
        maindata_plot_tensou_jidou
    Else
        maindata_plot_tensou
    End If
    Exit Sub
End If


If kidou_type2 = -1 Then
    bb_hyouji ("システムを最新に更新中・・")
End If


'servertxt = "ftp://203.138.49.67"
servertxt = "ftp://plot-office.com"


Screen.MousePointer = 11
Picture2.Enabled = False
daijoubu = 0
listmsg 0, 0





With Inet1
                            
         listmsg 1, 0
         
If o2(0).Value = True Then
'da1file2 = "c:\oldcatv.exe"
da1file2 = root_pass & "oldcatv.exe"

If Dir(da1file2) <> "" Then
    Kill da1file2
    DoEvents
End If

'da1file = "c:\catv.exe"
da1file = root_pass & "catv.exe"
If Dir(da1file) <> "" Then
    Name da1file As da1file2
    DoEvents
End If

On Error GoTo errinet
        .URL = servertxt
        '.UserName = "versionup"
        '.Password = "cheveux"
        .UserName = "plot-office"
        .Password = "Plot8877Ken"
        .RequestTimeout = 15
        daijoubu = 1
            listmsg 2, 1
        '.Execute , "CD ezmanager/new/"
        .Execute , "CD www/versionup/ezmanager/new7"
        listmsg 3, 0
        listmsg 4, 0
        Do While Inet1.StillExecuting = True
            DoEvents
        Loop
        If daijoubu = -1 Then
            listmsg -1, 0
            Picture2.Enabled = True
            .Execute , "close"
            Exit Sub
        End If
        listmsg 5, 0
        'For i = 0 To 1
        i = 0
                daijoubu = 10 + i
                listmsg 6, i
                'setstr = "GET " & strFileName2(i) & Space(1) & "c:\" & strFileName2(i)
                setstr = "GET " & strFileName2(i) & Space(1) & root_pass & strFileName2(i)
                .Execute , setstr
                listmsg 7, i
                
                Do While Inet1.StillExecuting = True
                    DoEvents
                Loop
                listmsg 8, i
                If daijoubu = -1 Then
                    listmsg -1, 0
                    Picture2.Enabled = True
                    .Execute , "close"
                    Exit Sub
                End If
        
Else
    da2file2 = root_pass & "oldezmanager_ment.exe"
    If Dir(da2file2) <> "" Then
        Kill da2file2
        DoEvents
    End If
    
    da2file = root_pass & "ezmanager_ment.exe"
    If Dir(da2file) <> "" Then
        Name da2file As da2file2
        DoEvents
    End If

        .URL = servertxt
        '.UserName = "versionup"
        '.Password = "cheveux"
        .UserName = "plot-office"
        .Password = "Plot8877Ken"
        .RequestTimeout = 15
        daijoubu = 1
            listmsg 2, 1
        '.Execute , "CD ezmanager/new/"
        .Execute , "CD www/versionup/ezmanager/new7"
        listmsg 3, 0
        listmsg 4, 0
        Do While Inet1.StillExecuting = True
            DoEvents
        Loop
        If daijoubu = -1 Then
            listmsg -1, 0
            Picture2.Enabled = True
            .Execute , "close"
            Exit Sub
        End If
        listmsg 5, 0
        
        
       i = 1
                daijoubu = 10 + i
                listmsg 6, i
                setstr = "GET " & strFileName2(i) & Space(1) & root_pass & strFileName2(i)
                .Execute , setstr
                listmsg 7, i
                
                Do While Inet1.StillExecuting = True
                    DoEvents
                Loop
                listmsg 8, i
                If daijoubu = -1 Then
                    listmsg -1, 0
                    Picture2.Enabled = True
                    .Execute , "close"
                    Exit Sub
                End If
        'Next i
        daijoubu = 3
End If
        .Execute , "close"
            listmsg 9, 0
        verget (4)
Screen.MousePointer = 0
If kidou_type2 = -1 Then
           'タスクトレイのアイコンを消去
        
        '-------------タスクトレイから削除-------------
        With trayicon
        .cbSize = Len(trayicon)
        .hWnd = frmmain.hWnd
        .uID = 0&
        End With
        
        ret = Shell_NotifyIcon(NIM_DELETE, trayicon)
        '-------------タスクトレイから削除END-------------

        End
End If
    MsgBox "完了しました。画面の「キャンセル」ボタンを押すと、「ＥＺＭａｎａｇｅｒ」が起動します。"

On Error GoTo 0
End With
Picture2.Enabled = True

Exit Sub

errinet:
    Screen.MousePointer = 0
    
    MsgBox "エラーが発生しました。"
    Picture2.Enabled = True
    Exit Sub

End Sub

Private Sub Picture2_Click()
Dim RetVal, resetfaile As String

Unload Me
resetfaile = "c:\program files\catv\catv.exe"
RetVal = Shell(resetfaile, 1)   ' 電卓を実行します。

End

End Sub

