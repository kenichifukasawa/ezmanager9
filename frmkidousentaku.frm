VERSION 5.00
Begin VB.Form frmkidousentaku 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "起動選択"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   9405
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "起動するデータベースを選択してください。"
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
      Height          =   1365
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   9255
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   8340
         MouseIcon       =   "frmkidousentaku.frx":0000
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmkidousentaku.frx":08FF
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   7
         Top             =   330
         Width           =   480
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'なし
         Height          =   480
         Left            =   240
         MouseIcon       =   "frmkidousentaku.frx":1741
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmkidousentaku.frx":2040
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   6
         Top             =   480
         Width           =   480
      End
      Begin VB.PictureBox Picture16 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   6405
         MouseIcon       =   "frmkidousentaku.frx":290A
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmkidousentaku.frx":3209
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         ToolTipText     =   "指定条件のデータを抽出表示します。"
         Top             =   330
         Width           =   480
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   7372
         MouseIcon       =   "frmkidousentaku.frx":3B08
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmkidousentaku.frx":4407
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         Top             =   330
         Width           =   480
      End
      Begin VB.ComboBox combkaisha 
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
         Left            =   900
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   1
         Top             =   600
         Width           =   4995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "設定"
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
         Index           =   7
         Left            =   8340
         TabIndex        =   8
         Top             =   870
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "起動"
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
         Left            =   6390
         TabIndex        =   5
         Top             =   870
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
         Left            =   7365
         TabIndex        =   3
         Top             =   870
         Width           =   480
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '不透明
         BorderStyle     =   0  '透明
         Height          =   975
         Index           =   2
         Left            =   6150
         Top             =   240
         Width           =   2985
      End
   End
End
Attribute VB_Name = "frmkidousentaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub combkaisha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture16_Click
End If
End Sub


Private Sub Picture16_Click()

Dim kidou_no As Integer

kidou_no = combkaisha.ListIndex + 1

kidoumei = combkaisha.Text

main2 (kidou_no)

End Sub

Private Sub Picture3_Click()
Unload Me
End

End Sub

Private Sub Picture7_Click()
frmdatapath.Show 1
End Sub
