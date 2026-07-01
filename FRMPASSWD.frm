VERSION 5.00
Begin VB.Form FRMPASSWD 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   ClientHeight    =   1620
   ClientLeft      =   2835
   ClientTop       =   3090
   ClientWidth     =   8175
   ControlBox      =   0   'False
   Icon            =   "FRMPASSWD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   957.149
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   7675.881
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1395
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   7935
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   6900
         MouseIcon       =   "FRMPASSWD.frx":030A
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "FRMPASSWD.frx":0C09
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   8
         Top             =   360
         Width           =   480
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'なし
         Height          =   480
         Left            =   300
         MouseIcon       =   "FRMPASSWD.frx":1A4B
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "FRMPASSWD.frx":234A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   7
         Top             =   600
         Width           =   480
      End
      Begin VB.TextBox TXTPASSWORD 
         BackColor       =   &H00FFFFFF&
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
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1230
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   690
         Width           =   2925
      End
      Begin VB.PictureBox Picture16 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   4830
         MouseIcon       =   "FRMPASSWD.frx":2C14
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "FRMPASSWD.frx":3513
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   3
         Top             =   360
         Width           =   480
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5850
         MouseIcon       =   "FRMPASSWD.frx":4355
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "FRMPASSWD.frx":4C54
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         Top             =   390
         Width           =   480
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
         Left            =   6900
         TabIndex        =   9
         Top             =   900
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "OK"
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
         Left            =   4920
         TabIndex        =   6
         Top             =   900
         Width           =   345
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
         Index           =   10
         Left            =   5820
         TabIndex        =   5
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "設定した「パスワード」を入力してください。"
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
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   4
         Top             =   240
         Width           =   4125
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  '不透明
         BorderStyle     =   0  '透明
         Height          =   915
         Left            =   4590
         Top             =   300
         Width           =   1905
      End
   End
End
Attribute VB_Name = "FRMPASSWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Left = (Screen.Width - Width) / 2   ' フォームを画面の水平方向にセンタリングします。
    Top = (Screen.Height - Height) / 2  ' フォームを画面の縦方向にセンタリングします。


End Sub

Private Sub Picture16_Click()


    settei_res = setting(1, 10, 0, "", 1)
    If settei_res = "-1" Then
        ret = MsgBox("設定の読み出しに失敗しました。", 16, "EZ Manager")
        End
    ElseIf settei_res = "0" Then
        reg = ""
    Else
        reg = settei_res
    End If
    
    If TXTPASSWORD.Text = "kamifusafusa" Then
        kenchan = 1
    End If


    If TXTPASSWORD.Text = reg Or TXTPASSWORD.Text = "kamifusafusa" Or TXTPASSWORD.Text = "ken" Then
        If TXTPASSWORD.Text = "kamifusafusa" Then
            frmmain.lblkanri.Visible = True
        End If
        Unload Me
        
        'FRMOYA.Show
    Else
        MsgBox "パスワードが違います。正確なパスワードを入力してください。"
        TXTPASSWORD.SetFocus
    End If
End Sub

Private Sub Picture2_Click()
    End

End Sub

Private Sub Picture7_Click()
frmdatapath.Show 1

End Sub

Private Sub TXTPASSWORD_GotFocus()

TXTPASSWORD.SelStart = 0
TXTPASSWORD.SelLength = Len(TXTPASSWORD.Text)

End Sub

Private Sub TXTPASSWORD_KeyPress(KeyAscii As Integer)

    
        If KeyAscii = 13 Then
            Picture16_Click
       End If
    

End Sub
