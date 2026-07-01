VERSION 5.00
Begin VB.Form frmdatapath 
   Caption         =   "起動設定"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   11535
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "起動する種類とデータベースのパスを指定してください。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2805
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "起動の種類"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1080
         TabIndex        =   9
         Top             =   420
         Width           =   7545
         Begin VB.OptionButton optkidou 
            BackColor       =   &H00FFFFFF&
            Caption         =   "サブ起動(Log+光)"
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
            Index           =   2
            Left            =   4860
            TabIndex        =   12
            Top             =   420
            Width           =   2175
         End
         Begin VB.OptionButton optkidou 
            BackColor       =   &H00FFFFFF&
            Caption         =   "メイン起動"
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
            Index           =   0
            Left            =   270
            TabIndex        =   11
            Top             =   420
            Width           =   1815
         End
         Begin VB.OptionButton optkidou 
            BackColor       =   &H00FFFFFF&
            Caption         =   "サブ起動(Logのみ)"
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
            Index           =   1
            Left            =   2325
            TabIndex        =   10
            Top             =   420
            Width           =   2295
         End
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   9240
         MouseIcon       =   "frmdatapath.frx":0000
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmdatapath.frx":08FF
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   6
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   10140
         MouseIcon       =   "frmdatapath.frx":11C9
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmdatapath.frx":1AC8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'なし
         Height          =   480
         Left            =   180
         MouseIcon       =   "frmdatapath.frx":2392
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmdatapath.frx":2C91
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   3
         Top             =   960
         Width           =   480
      End
      Begin VB.CommandButton Command1 
         Caption         =   "...."
         Height          =   375
         Left            =   10380
         TabIndex        =   2
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox txtdatapath 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1140
         TabIndex        =   1
         Top             =   1920
         Width           =   9015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "データベースのパス"
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
         Index           =   14
         Left            =   1110
         TabIndex        =   8
         Top             =   1650
         Width           =   2010
      End
      Begin VB.Label SSCJIKKOU 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "更新"
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
         Left            =   9270
         TabIndex        =   7
         Top             =   1140
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "戻る"
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
         Index           =   13
         Left            =   10200
         TabIndex        =   5
         Top             =   1140
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmdatapath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim udtBrowseInfo As BROWSEINFO
    Dim lngIDList     As Long
    Dim strPath       As String
    
    '[フォルダの参照]ダイアログ情報を設定
    With udtBrowseInfo
        .hWndOwner = Me.hwnd
        .pidlRoot = CSIDL_NETWORK
        .pszDisplayName = String(MAX_PATH2, vbNullChar)
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    '[フォルダの参照]ダイアログを表示
    lngIDList = SHBrowseForFolder(udtBrowseInfo)
    If lngIDList <> 0 Then
        strPath = String(MAX_PATH2, vbNullChar)
        'ITEMIDLISTをパスに変換
        SHGetPathFromIDList lngIDList, strPath
        CoTaskMemFree lngIDList
            strPath = Left(strPath, _
                           InStr(strPath, _
                                 vbNullChar) - 1)
        '結果をラベルに表示
        txtdatapath.Text = strPath
    End If
End Sub



Private Sub Form_Load()
settei_res = setting(1, 7, 0, "", 1)
If settei_res = "-1" Then
    ret = MsgBox("設定の読み出しに失敗しました。004", 16, "EZ Manager")
   ' db3.Close
    End
ElseIf settei_res = "0" Then
    optkidou(0).Value = True
ElseIf settei_res = "1" Then
    optkidou(1).Value = True
ElseIf settei_res = "2" Then
    optkidou(2).Value = True
End If

settei_res = setting(1, 4, 0, "", 1)
If settei_res = "-1" Then
    ret = MsgBox("設定の読み出しに失敗しました。004", 16, "EZ Manager")
    'db3.Close
    End
ElseIf settei_res = "0" Then
    ret = MsgBox("データﾍﾞースの設定がありません。", 16, "EZ Manager")
    End
Else
    txtdatapath.Text = settei_res
End If




End Sub

Private Sub Picture3_Click()
Unload Me

End Sub

Private Sub Picture7_Click()

Dim new_data_path As String, chkpath As String, new_shiyou As String

If optkidou(0).Value = True Then
    new_shiyou = "0"
ElseIf optkidou(1).Value = True Then
    new_shiyou = "1"
ElseIf optkidou(2).Value = True Then
    new_shiyou = "2"
End If

    new_data_path = Trim(txtdatapath.Text)
    
    
    chkpath = new_data_path & "\catv.mdb"
    
    If new_data_path = "" Then
        ret = MsgBox("データパスが入力されていません。", 16, "EZ Manager")
        Exit Sub
    End If
   ' If Dir(chkpath) = "" Then
      '  ret = MsgBox("有効なデータパスではありません。", 16, "EZ Manager")
       ' Exit Sub
    'End If
    
    settei_res = setting(1, 7, 1, new_shiyou, 0)
    If settei_res = "-1" Then
        ret = MsgBox("設定の書き込みに失敗しました。", 16, "EZ Manager")
        db3.Close
        End
    End If
    settei_res = setting(0, 4, 1, new_data_path, 0)
    If settei_res = "-1" Then
        ret = MsgBox("設定の書き込みに失敗しました。", 16, "EZ Manager")
        db3.Close
        End
    End If
    
    
    
ret = MsgBox("データパスと起動モードを設定しました。一度終了します。", 64, "EZ Manager")
End

End Sub
