VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmbank 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "金融機関"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   10200
      MouseIcon       =   "frmbank.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmbank.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      ToolTipText     =   "登録処理を行います。"
      Top             =   3660
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   10200
      MouseIcon       =   "frmbank.frx":11C9
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmbank.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   4860
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "金融機関一覧"
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
      Height          =   6045
      Left            =   60
      TabIndex        =   3
      Top             =   120
      Width           =   9555
      Begin MSFlexGridLib.MSFlexGrid bankgrd 
         Height          =   5625
         Left            =   60
         TabIndex        =   4
         ToolTipText     =   "金融機関一覧です。"
         Top             =   240
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   9922
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColor       =   16777215
         BackColorBkg    =   16777215
      End
   End
   Begin VB.Frame SSFTOKUSHU 
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      Height          =   3075
      Left            =   9720
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   480
         MouseIcon       =   "frmbank.frx":2392
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmbank.frx":2C91
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   10
         Top             =   2070
         Width           =   480
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   450
         MouseIcon       =   "frmbank.frx":355B
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmbank.frx":3E5A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         Top             =   1110
         Width           =   480
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "番号"
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
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   1275
         Begin VB.ComboBox CBOREN 
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
            Left            =   90
            Sorted          =   -1  'True
            TabIndex        =   2
            ToolTipText     =   "金融機関番号を指定します。"
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         Caption         =   "削除"
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
         Height          =   285
         Index           =   6
         Left            =   360
         TabIndex        =   12
         Top             =   2580
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "変更"
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
         Index           =   5
         Left            =   450
         TabIndex        =   11
         Top             =   1650
         Width           =   495
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "登録"
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
      Index           =   7
      Left            =   10140
      TabIndex        =   8
      Top             =   4170
      Width           =   600
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
      Left            =   10230
      TabIndex        =   7
      Top             =   5430
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   5985
      Left            =   9750
      Top             =   150
      Width           =   1410
   End
End
Attribute VB_Name = "frmbank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()

    Left = (Screen.Width - Width) / 2   ' フォームを画面の水平方向にセンタリングします。
    Top = (Screen.Height - Height) / 2  ' フォームを画面の縦方向にセンタリングします。
    frmbank.bankgrd.Font.Size = 11
End Sub

Private Sub Picture13_Click()
    Unload frmbank
    Set frmbank = Nothing

End Sub

Private Sub Picture7_Click()
    If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    End If


If Trim(CBOREN.Text) = "" Then
      MsgBox "ＩＤを選択してから、実行してください。"
      CBOREN.SetFocus
      Exit Sub
Else
      sql = "SELECT*FROM bank WHERE bankid = '" & Trim(CBOREN.Text) & "'"
      If FcSQlGet(rs, sql, PrMsg) = False Then
           MsgBox "ＩＤを正確に入力、又は、選択してください。"
           CBOREN.SetFocus
           Exit Sub
      Else
           With frmbank2
               .lbloldbankid.Caption = rs!bankid
               .lbloldbankmei.Caption = rs!bankmei
               .lbloldbankcode.Caption = rs!bankcode
               .lbloldshitencode.Caption = rs!shitencode
               .lblnewbankid.Caption = rs!bankid
               .txtnewbankmei.Text = rs!bankmei
               .txtnewbankcode.ListIndex = FcCmbbankidGet2(rs!bankcode)
               .txtnewshitencode.Text = rs!shitencode
               .Caption = "金融機関の変更"
               .SSCJIKKOU.Caption = "変更"
           End With
           rs.Close
           frmbank.Hide
           frmbank2.Show 1
      End If
End If
End Sub

Private Sub Picture8_Click()
    If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    End If
If Trim(CBOREN.Text) = "" Then
      MsgBox "ＩＤを選択してから、実行してください。"
    CBOREN.SetFocus
      Exit Sub
Else
         sql = "SELECT*FROM bank WHERE bankid = '" & Trim(CBOREN.Text) & "'"
         If FcSQlGet(rs, sql, PrMsg) = False Then
              MsgBox "ＩＤを正確に入力、又は、選択してください。"
              CBOREN.SetFocus
              Exit Sub
         End If
        msg = "「" & rs!bankmei & "」を削除します。"
        res = MsgBox(msg, vbYesNo)
        If res = vbYes Then
          rs.Close
          sql = "SELECT*FROM kojin WHERE bankid = '" & Trim(CBOREN.Text) & "'"
          If FcSQlGet(rs, sql, PrMsg) = True Then
                 MsgBox "この項目は、データの中に使用されていますので、削除できません。"
                rs.Close
                Exit Sub
          End If
          rs.Close
            sql = "DELETE FROM bank WHERE bankid = '" & Trim(CBOREN.Text) & "'"
    On Error GoTo errshousaku2
                db.Execute sql, dbSQLPassThrough
    On Error GoTo 0
            bankset
            MsgBox "削除しました。"
        End If
            
End If
Exit Sub
errshousaku2:
          MsgBox "削除に失敗しました。"
          Exit Sub

End Sub

Private Sub Picture9_Click()
    If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    End If
With frmbank2
    .Caption = "金融機関の登録"
    .SSCJIKKOU.Caption = "登録"
    If soukatsu(0, "renban", 3, "", 1, 1) = False Then
            MsgBox "設定の読み込みに失敗しました。"
    End If
    If reg = "0" Then
        reg = ""
    End If
    
    If reg = vbNullString Then
        .lblnewbankid.Caption = "001"
    Else
        .lblnewbankid.Caption = Format$(reg, "00#")
    End If
    
    frmbank.Hide
    .Show 1
End With

End Sub


