VERSION 5.00
Begin VB.Form frmdir 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "パスを指定してください。"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   5265
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2820
      MouseIcon       =   "frmdir.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmdir.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   4530
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   4170
      MouseIcon       =   "frmdir.frx":11C9
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmdir.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   4530
      Width           =   480
   End
   Begin VB.TextBox txtpath 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   5000
   End
   Begin VB.DirListBox dirlist 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3150
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4995
   End
   Begin VB.DriveListBox drvlist 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5000
   End
   Begin VB.Label SSCJIKKOU 
      AutoSize        =   -1  'True
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
      Height          =   225
      Left            =   2850
      TabIndex        =   6
      Top             =   5040
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "終了"
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
      Left            =   4200
      TabIndex        =   4
      Top             =   5040
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   1035
      Left            =   120
      Top             =   4350
      Width           =   5000
   End
End
Attribute VB_Name = "frmdir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdok_Click()
End Sub

Private Sub dirlist_Change()

txtpath.Text = dirlist.Path


End Sub

Private Sub drvlist_Change()
On Error GoTo DriveHandler   ' 新しいドライブが選択された場合は、Dir1 ボックスの   ' 表示を更新します。
   dirlist.Path = drvlist.Drive
   Exit Sub      ' エラーが発生した場合は、drvList.Drive の値を
   ' dirList.Path のドライブに戻します。
DriveHandler:
   drvlist.Drive = dirlist.Path
   Exit Sub

End Sub

Private Sub Picture13_Click()
Unload Me
End Sub

Private Sub Picture7_Click()
Dim krishiteipath As String, subpath As String

krishiteipath = Trim(txtpath.Text)

If krishiteipath <> "" Then
        If Dir$(krishiteipath, vbDirectory) = "" Then
            Dim fso
            Set fso = CreateObject("Scripting.FileSystemObject")
            fso.CreateFolder (krishiteipath)
            Set fso = Nothing
        End If
End If
If Trim(txtpath.Text) <> "" Then
    If Right(txtpath.Text, 1) <> "\" Then
        txtpath.Text = txtpath.Text & "\"
    End If
End If
If dirdou = 1 Then
    FRMSETTEI.txtback1.Text = txtpath.Text
ElseIf dirdou = 2 Then
    'FRMSETTEI.txtback2.Text = txtpath.Text
ElseIf dirdou = 3 Then
   ' FRMSETTEI.txtback2.Text = txtpath.Text
ElseIf dirdou = 4 Then
    frmdatapath.txtdatapath.Text = txtpath.Text
  ElseIf dirdou = 5 Then
    frmfdd.txtde.Text = txtpath.Text
ElseIf dirdou = 6 Then
    frmkoushin.txtde.Text = txtpath.Text
ElseIf dirdou = 7 Then
    frmkoushin2.txtde.Text = txtpath.Text
ElseIf dirdou = 8 Then
    FRMSETTEI.txtdrpobox.Text = txtpath.Text
ElseIf dirdou = 9 Then
    FRMSETTEI.txtbunsho.Text = txtpath.Text

End If
dirdou = 0
    Unload frmdir
    Set frmdir = Nothing

End Sub
