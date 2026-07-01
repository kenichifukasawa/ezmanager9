VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMOMACHI 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "処理中......"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5190
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture16 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'なし
      Height          =   480
      Left            =   180
      MouseIcon       =   "frmomachi.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmomachi.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   210
      Width           =   480
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   195
      Left            =   870
      TabIndex        =   1
      Top             =   750
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblmsg 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "FRMOMACHI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Left = (Screen.Width - Width) / 2   ' フォームを画面の水平方向にセンタリングします。
    Top = (Screen.Height - Height) / 2  ' フォームを画面の縦方向にセンタリングします。

End Sub

Private Sub Picture1_Click()

End Sub
