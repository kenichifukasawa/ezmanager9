VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmsonota 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "その他の契約"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   11385
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   10230
      MouseIcon       =   "frmsonota.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmsonota.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   4800
      Width           =   480
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   10230
      MouseIcon       =   "frmsonota.frx":11C9
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmsonota.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      ToolTipText     =   "登録処理を行います。"
      Top             =   3600
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "その他契約種類一覧"
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
      Height          =   6165
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   9555
      Begin MSFlexGridLib.MSFlexGrid sonotaichigrid 
         Height          =   5805
         Left            =   60
         TabIndex        =   4
         ToolTipText     =   "契約種類一覧です。"
         Top             =   240
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   10239
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         BackColor       =   16777215
         BackColorBkg    =   16777215
      End
   End
   Begin VB.Frame SSFTOKUSHU 
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      Height          =   3105
      Left            =   9750
      TabIndex        =   0
      Top             =   90
      Width           =   1425
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   480
         MouseIcon       =   "frmsonota.frx":2392
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmsonota.frx":2C91
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   10
         Top             =   1140
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
         MouseIcon       =   "frmsonota.frx":355B
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmsonota.frx":3E5A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         Top             =   2100
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
         Height          =   855
         Left            =   60
         TabIndex        =   1
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
            Left            =   150
            Sorted          =   -1  'True
            TabIndex        =   2
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
         Top             =   1680
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
         Top             =   2610
         Width           =   675
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
      Left            =   10260
      TabIndex        =   8
      Top             =   5370
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
      Left            =   10170
      TabIndex        =   7
      Top             =   4110
      Width           =   600
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   6045
      Left            =   9780
      Top             =   90
      Width           =   1410
   End
End
Attribute VB_Name = "frmsonota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
frmsonota.sonotaichigrid.Font.Size = 11
End Sub

Private Sub Picture13_Click()
Unload frmsonota
Set frmsonota = Nothing

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
ElseIf Trim(CBOREN.Text) = "99" Then
    MsgBox "『99　返金』は変更できません。"
    Exit Sub
ElseIf Trim(CBOREN.Text) = "98" Then
    MsgBox "『98　前回未納金』は変更できません。"
    Exit Sub
ElseIf Trim(CBOREN.Text) = "97" Then
    MsgBox "『97　PPV課金』は変更できません。"
    Exit Sub
Else
      sql = "SELECT*FROM sonota WHERE sonotaid = '" & Trim(CBOREN.Text) & "'"
      If FcSQlGet(rs, sql, PrMsg) = False Then
           MsgBox "ＩＤを正確に入力、又は、選択してください。"
           CBOREN.SetFocus
           Exit Sub
      Else
           With frmsonota2
               .lbloldsonotaid.Caption = rs!sonotaid
               .lbloldsonotamei.Caption = rs!sonota
               If rs!kotei = "0" Then
                    .lbloldsonotakotei.Caption = "固定"
                    .optkotei(0).Value = True
               Else
                    .lbloldsonotakotei.Caption = "変動"
                    .optkotei(1).Value = True
               End If
               If IsNull(rs!kingaku) Then
                    .lbloldsonotakin.Caption = ""
                    .txtnewsonotakin.Text = ""
               Else
                    .lbloldsonotakin.Caption = rs!kingaku
                    .txtnewsonotakin.Text = rs!kingaku
               End If
               .lblnewsonotaid.Caption = rs!sonotaid
               .txtnewsonotamei.Text = rs!sonota
               
               If rs!kankei = "0" Then
                    .optoldbunrui(0).Value = True
                    .optnewbunrui(0).Value = True
               Else
                    .optnewbunrui(1).Value = True
                    .optoldbunrui(1).Value = True
               End If
               
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
               
               
               .Caption = "その他の契約の変更"
               .SSCJIKKOU.Caption = "変更"
           End With
           rs.Close
           frmsonota.Hide
           frmsonota2.Show 1
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
ElseIf Trim(CBOREN.Text) = "99" Then
    MsgBox "『99　返金』は削除できません。"
    Exit Sub
ElseIf Trim(CBOREN.Text) = "98" Then
    MsgBox "『98　前回未納金』は削除できません。"
    Exit Sub
ElseIf Trim(CBOREN.Text) = "97" Then
    MsgBox "『97　PPV課金』は削除できません。"
    Exit Sub
ElseIf Trim(CBOREN.Text) = "96" Then
    MsgBox "『96　滞納手数料』は削除できません。"
    Exit Sub
Else
         sql = "SELECT*FROM sonota WHERE sonotaid = '" & Trim(CBOREN.Text) & "'"
         If FcSQlGet(rs, sql, PrMsg) = False Then
              MsgBox "ＩＤを正確に入力、又は、選択してください。"
              CBOREN.SetFocus
              Exit Sub
         End If
        msg = "「" & rs!sonota & "」を削除します。"
        res = MsgBox(msg, vbYesNo)
        If res = vbYes Then
          rs.Close
          sql = "SELECT*FROM keiyaku2 WHERE sonotaid = '" & Trim(CBOREN.Text) & "'"
          If FcSQlGet(rs, sql, PrMsg) = True Then
                 MsgBox "この項目は、データの中に使用されていますので、削除できません。"
                rs.Close
                Exit Sub
          End If
          rs.Close
            sql = "DELETE FROM sonota WHERE sonotaid = '" & Trim(CBOREN.Text) & "'"
    On Error GoTo errsonotasaku2
                db.Execute sql, dbSQLPassThrough
    On Error GoTo 0
            sonotaset
            MsgBox "削除しました。"
        End If
            
End If
Exit Sub
errsonotasaku2:
          MsgBox "削除に失敗しました。"
          Exit Sub

End Sub

Private Sub Picture9_Click()

    If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    End If
With frmsonota2
    .Caption = "その他の契約の登録"
    .SSCJIKKOU.Caption = "登録"
    If soukatsu(0, "renban", 5, "", 1, 1) = False Then
            MsgBox "設定の読み込みに失敗しました。"
    End If
    If reg = "0" Then
        reg = ""
    End If
    If reg = vbNullString Then
        .lblnewsonotaid.Caption = "01"
    ElseIf reg = "96" Then
        MsgBox "これ以上のその他の契約項目の作成はできません。管理者に連絡してください。"
        Exit Sub
    Else
        .lblnewsonotaid.Caption = Format$(reg, "0#")
    End If
    
    frmsonota.Hide
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
