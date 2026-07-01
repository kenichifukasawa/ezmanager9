VERSION 5.00
Begin VB.Form frmnhkshuukei 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "指定期間NHK集計表"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9480
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8340
      MouseIcon       =   "frmnhkshuukei.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmnhkshuukei.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   4740
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
      Left            =   8400
      TabIndex        =   1
      Top             =   5250
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   975
      Left            =   7800
      Top             =   4620
      Width           =   1500
   End
End
Attribute VB_Name = "frmnhkshuukei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Picture13_Click()
Unload Me
End Sub
