VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmcell2 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "セル一覧"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   6780
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5610
      MouseIcon       =   "frmcell2.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmcell2.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
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
      Index           =   1
      Left            =   5640
      MouseIcon       =   "frmcell2.frx":11C9
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmcell2.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   4920
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "セル一覧"
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
      Height          =   6015
      Index           =   5
      Left            =   120
      TabIndex        =   3
      Top             =   90
      Width           =   4845
      Begin MSFlexGridLib.MSFlexGrid cellgrid 
         Height          =   5595
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "金融機関一覧です。"
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   9869
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         BackColorBkg    =   16777215
      End
   End
   Begin VB.Frame SSFTOKUSHU 
      BackColor       =   &H80000009&
      ClipControls    =   0   'False
      Height          =   2955
      Left            =   5130
      TabIndex        =   0
      Top             =   150
      Width           =   1425
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   450
         MouseIcon       =   "frmcell2.frx":2392
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmcell2.frx":2C91
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   10
         Top             =   1020
         Width           =   480
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   480
         MouseIcon       =   "frmcell2.frx":355B
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmcell2.frx":3E5A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         Top             =   1980
         Width           =   480
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H80000009&
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
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   1305
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
            TabIndex        =   2
            ToolTipText     =   "金融機関番号を指定します。"
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
         Left            =   450
         TabIndex        =   12
         Top             =   1560
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
         Left            =   360
         TabIndex        =   11
         Top             =   2490
         Width           =   675
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
      Left            =   5550
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
      Index           =   12
      Left            =   5670
      TabIndex        =   6
      Top             =   5430
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   5985
      Left            =   5160
      Top             =   150
      Width           =   1410
   End
End
Attribute VB_Name = "frmcell2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Picture13_Click(Index As Integer)
Unload Me
End Sub

Private Sub Picture7_Click()
If Trim(CBOREN.Text) = "" Then
      MsgBox "ＩＤを選択してから、実行してください。"
      CBOREN.SetFocus
      Exit Sub
Else
      sql = "SELECT*FROM cell WHERE cellid = '" & Trim(CBOREN.Text) & "'"
      If FcSQlGet(rs, sql, PrMsg) = False Then
           MsgBox "ＩＤを正確に入力、又は、選択してください。"
           CBOREN.SetFocus
           Exit Sub
      Else
           With frmcell
               .lbloldcellid.Caption = rs!cellid
               .lbloldcellname.Caption = rs!cellname
               .lblnewcellid.Caption = rs!cellid
               .txtnewcellname.Text = rs!cellname
               .SSCJIKKOU.Caption = "変更"
           End With
           rs.Close
           frmcell.Show 1
      End If
End If

End Sub

Private Sub Picture8_Click()
If Trim(CBOREN.Text) = "" Then
      MsgBox "ＩＤを選択してから、実行してください。"
    CBOREN.SetFocus
      Exit Sub
Else
         sql = "SELECT*FROM cell WHERE cellid = '" & Trim(CBOREN.Text) & "'"
         If FcSQlGet(rs, sql, PrMsg) = False Then
              MsgBox "ＩＤを正確に入力、又は、選択してください。"
              CBOREN.SetFocus
              Exit Sub
         End If
        msg = "「" & rs!cellname & "」を削除します。"
        res = MsgBox(msg, vbYesNo)
        If res = vbYes Then
          rs.Close
          sql = "SELECT*FROM kojin WHERE cellid = '" & Trim(CBOREN.Text) & "'"
          If FcSQlGet(rs, sql, PrMsg) = True Then
                 MsgBox "この項目は、データの中に使用されていますので、削除できません。"
                rs.Close
                Exit Sub
          End If
          rs.Close
            sql = "DELETE FROM cell WHERE cellid = '" & Trim(CBOREN.Text) & "'"
    On Error GoTo errshousaku2
                db.Execute sql, dbSQLPassThrough
    On Error GoTo 0
            cellset
            MsgBox "削除しました。"
        End If
            
End If
Exit Sub
errshousaku2:
          MsgBox "削除に失敗しました。"
          Exit Sub

End Sub

Private Sub Picture9_Click()
With frmcell
    .SSCJIKKOU.Caption = "登録"
    If soukatsu(0, "renban", 12, "", 1, 1) = False Then
            MsgBox "設定の読み込みに失敗しました。"
    End If
    If reg = "0" Then
        reg = ""
    End If
    
    If reg = vbNullString Then
        .lblnewcellid.Caption = "01"
    Else
        .lblnewcellid.Caption = Format$(reg, "0#")
    End If
    
    .Show 1
End With

End Sub

Private Sub SSCHENKOU_Click()

End Sub

Private Sub SSCSAKUJO_Click()

End Sub

Private Sub SSCTOUROKU_Click()


End Sub
