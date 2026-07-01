VERSION 5.00
Begin VB.Form frmverup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "バージョンアップ管理"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   ControlBox      =   0   'False
   Icon            =   "frmverup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   7155
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture16 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'なし
      Height          =   480
      Left            =   600
      MouseIcon       =   "frmverup.frx":0E42
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmverup.frx":1741
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   300
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "処理しています。少々お待ちください......"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1950
      TabIndex        =   0
      Top             =   390
      Width           =   4215
   End
End
Attribute VB_Name = "frmverup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub
