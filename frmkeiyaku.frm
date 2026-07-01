VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmkeiyaku 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "契約"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   12780
      MouseIcon       =   "frmkeiyaku.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmkeiyaku.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   4920
      Width           =   480
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   12780
      MouseIcon       =   "frmkeiyaku.frx":11C9
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmkeiyaku.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      ToolTipText     =   "登録処理を行います。"
      Top             =   3690
      Width           =   480
   End
   Begin VB.Frame SSFTOKUSHU 
      BackColor       =   &H8000000E&
      ClipControls    =   0   'False
      Height          =   3045
      Left            =   12300
      TabIndex        =   2
      Top             =   210
      Width           =   1455
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   480
         MouseIcon       =   "frmkeiyaku.frx":2392
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmkeiyaku.frx":2C91
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   10
         Top             =   1050
         Width           =   480
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   510
         MouseIcon       =   "frmkeiyaku.frx":355B
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmkeiyaku.frx":3E5A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         Top             =   2010
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
         TabIndex        =   3
         Top             =   120
         Width           =   1335
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
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   4
            ToolTipText     =   "種類番号を指定します。"
            Top             =   300
            Width           =   1095
         End
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
         Left            =   480
         TabIndex        =   12
         Top             =   1590
         Width           =   495
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
         Left            =   390
         TabIndex        =   11
         Top             =   2520
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "契約種類一覧"
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
      Height          =   6225
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   12015
      Begin MSFlexGridLib.MSFlexGrid shuruigrd 
         Height          =   5895
         Left            =   60
         TabIndex        =   1
         ToolTipText     =   "契約種類一覧です。"
         Top             =   240
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   10398
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         BackColor       =   16777215
         BackColorBkg    =   16777215
      End
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
      Left            =   12810
      TabIndex        =   8
      Top             =   5490
      Width           =   495
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
      Left            =   12720
      TabIndex        =   6
      Top             =   4230
      Width           =   600
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   5985
      Left            =   12330
      Top             =   210
      Width           =   1410
   End
End
Attribute VB_Name = "frmkeiyaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Left = (Screen.Width - Width) / 2   ' フォームを画面の水平方向にセンタリングします。
    Top = (Screen.Height - Height) / 2  ' フォームを画面の縦方向にセンタリングします。
    frmkeiyaku.shuruigrd.Font.Size = 11
End Sub

Private Sub Picture13_Click()
    Unload frmkeiyaku
    Set frmkeiyaku = Nothing

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
    If CInt(CBOREN.Text) > 95 Then
        MsgBox "この項目は変更できません。"
        Exit Sub
    End If
      sql = "SELECT*FROM shurui WHERE shuruiid = '" & Trim(CBOREN.Text) & "'"
      If FcSQlGet(rs, sql, PrMsg) = False Then
           MsgBox "ＩＤを正確に入力、又は、選択してください。"
           CBOREN.SetFocus
           Exit Sub
      Else
           With frmkeiyaku2
               .lbloldshuruiid.Caption = rs!shuruiid
               .lbloldshuruimei.Caption = rs!shuruimei
               .lbloldsetsumei.Caption = rs!setsumei
               .lbloldtanka.Caption = rs!tanka
               .lblnewshuruiid.Caption = rs!shuruiid
               .txtnewshuruimei.Text = rs!shuruimei
               .txtnewsetsumei.Text = rs!setsumei
               .txtnewtanka.Text = rs!tanka
               kakaku = CLng(rs!tanka)
               If rs!kankei = "0" Then
                    .optoldbunrui(0).Value = True
                    .optnewbunrui(0).Value = True
               Else
                    .optoldbunrui(1).Value = True
                    .optnewbunrui(1).Value = True
               End If
               .Caption = "契約種類の変更"
               .SSCJIKKOU.Caption = "変更"
               .fh.Enabled = True
               '.opthouhou(1).Enabled = False
               If IsNull(rs!iru) Then
               .lblshuukei.Caption = ""
               Else
               Select Case rs!iru
               Case "0"
                .lblshuukei.Caption = "なし"
                .optshuukei(0).Value = True
               Case "1"
                .lblshuukei.Caption = "あり"
                .optshuukei(1).Value = True
               End Select
               End If
               If IsNull(rs!gp) Then
               .lblgroup.Caption = ""
               .txtgroup.Text = ""
               Else
               .lblgroup.Caption = rs!gp
               .txtgroup.Text = rs!gp
               End If
               If IsNull(rs!narabe) Then
               .lblnarabe.Caption = ""
               .txtnarabe.Text = ""
               Else
               .lblnarabe.Caption = rs!narabe
               .txtnarabe.Text = rs!narabe
               End If
               
              If IsNull(rs!hyoukimei) Then
               .lblhyouki.Caption = ""
               .txthyouki.Text = ""
               Else
               .lblhyouki.Caption = rs!hyoukimei
               .txthyouki.Text = rs!hyoukimei
               End If
               
               If IsNull(rs!kasan) Then
               .lblchousei.Caption = ""
               .txtchousei.Text = ""
               Else
               .lblchousei.Caption = rs!kasan
               .txtchousei.Text = rs!kasan
               End If
               
           End With
           rs.Close
           frmkeiyaku.Hide
           frmkeiyaku2.Show 1
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
     If Trim(CBOREN.Text) = "01" Then
        MsgBox "01．ＣＡＴＶ主契約は、削除できません。"
        Exit Sub
     End If
    If CInt(CBOREN.Text) > 95 Then
        MsgBox "この項目は削除できません。"
        Exit Sub
    End If
         sql = "SELECT*FROM shurui WHERE shuruiid = '" & Trim(CBOREN.Text) & "'"
         If FcSQlGet(rs, sql, PrMsg) = False Then
              MsgBox "ＩＤを正確に入力、又は、選択してください。"
              CBOREN.SetFocus
              Exit Sub
         End If
        msg = "「" & rs!shuruimei & "」を削除します。"
        res = MsgBox(msg, vbYesNo)
        If res = vbYes Then
          rs.Close
          sql = "SELECT*FROM keiyaku WHERE shurui = '" & Trim(CBOREN.Text) & "'"
          If FcSQlGet(rs, sql, PrMsg) = True Then
                 MsgBox "この項目は、データの中に使用されていますので、削除できません。"
                rs.Close
                Exit Sub
          End If
            sql = "DELETE FROM shurui WHERE shuruiid = '" & Trim(CBOREN.Text) & "'"
    On Error GoTo errshousaku
                db.Execute sql, dbSQLPassThrough
    On Error GoTo 0
            shuruiset
            MsgBox "削除しました。"
        End If
            
End If
Exit Sub
errshousaku:
          MsgBox "削除に失敗しました。"
          Exit Sub

End Sub

Private Sub Picture9_Click()

    If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    End If

With frmkeiyaku2
    .Caption = "契約種類の登録"
    .SSCJIKKOU.Caption = "登録"
    If soukatsu(0, "renban", 2, "", 1, 1) = False Then
            MsgBox "設定の読み込みに失敗しました。"
    End If
    If reg = "0" Then
        reg = ""
    End If
    If reg = vbNullString Then
        .lblnewshuruiid.Caption = "01"
    Else
        .lblnewshuruiid.Caption = Format$(reg, "0#")
    End If
    
    frmkeiyaku.Hide
    .Show 1
End With


End Sub

Private Sub SSCHENKOU_Click()



End Sub

Private Sub SSCSAKUJO_Click()



End Sub

Private Sub SSCSHUURYOU_Click()
End Sub

Private Sub SSCTOUROKU_Click()
End Sub

