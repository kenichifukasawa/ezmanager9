VERSION 5.00
Begin VB.Form frmrirekisentaku 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "履歴選択"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   10965
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      Caption         =   "表示したい項目を選択してください。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1605
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   10695
      Begin VB.CheckBox chkhihyouji 
         BackColor       =   &H00FFFFFF&
         Caption         =   "削除伝票非表示"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   7710
         TabIndex        =   17
         Top             =   120
         Value           =   1  'ﾁｪｯｸ
         Width           =   1755
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   6990
         MouseIcon       =   "frmrirekisentaku.frx":0000
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmrirekisentaku.frx":08FF
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   15
         ToolTipText     =   "入金履歴を表示します。"
         Top             =   450
         Width           =   480
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   8235
         MouseIcon       =   "frmrirekisentaku.frx":1741
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmrirekisentaku.frx":2040
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   13
         ToolTipText     =   "入金履歴を表示します。"
         Top             =   450
         Width           =   480
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   4500
         MouseIcon       =   "frmrirekisentaku.frx":2E82
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmrirekisentaku.frx":3781
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   10
         ToolTipText     =   "入金履歴を表示します。"
         Top             =   450
         Width           =   480
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5745
         MouseIcon       =   "frmrirekisentaku.frx":45C3
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmrirekisentaku.frx":4EC2
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         ToolTipText     =   "入金履歴を表示します。"
         Top             =   450
         Width           =   480
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1944
         MouseIcon       =   "frmrirekisentaku.frx":5D04
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmrirekisentaku.frx":6603
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   7
         Top             =   450
         Width           =   480
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   750
         MouseIcon       =   "frmrirekisentaku.frx":6F02
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmrirekisentaku.frx":7801
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   3
         Top             =   450
         Width           =   480
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3240
         MouseIcon       =   "frmrirekisentaku.frx":8100
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmrirekisentaku.frx":89FF
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         Top             =   450
         Width           =   480
      End
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   9480
         MouseIcon       =   "frmrirekisentaku.frx":92FE
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmrirekisentaku.frx":9BFD
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   450
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         Caption         =   "CL"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   2
         Left            =   6720
         TabIndex        =   16
         Top             =   1050
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         Caption         =   "予定履歴"
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
         Index           =   5
         Left            =   8010
         TabIndex        =   14
         Top             =   1050
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         Caption         =   "その他契約"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   4194
         TabIndex        =   12
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         Caption         =   "VoIP"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   1
         Left            =   5532
         TabIndex        =   11
         Top             =   1050
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "取消・損金"
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
         Left            =   1653
         TabIndex        =   8
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "請求履歴"
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
         Left            =   450
         TabIndex        =   6
         Top             =   1050
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "入金履歴"
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
         Left            =   2991
         TabIndex        =   5
         Top             =   1050
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "戻る"
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
         Left            =   9510
         TabIndex        =   4
         Top             =   1050
         Width           =   450
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  '不透明
         BorderStyle     =   0  '透明
         Height          =   1185
         Left            =   180
         Top             =   300
         Width           =   10320
      End
   End
End
Attribute VB_Name = "frmrirekisentaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Picture1_Click()
Dim ssshiduhidu As String

If Trim(frmtouroku2.lblid.Caption) = "" Then
    MsgBox "個人IDを参照できません。"
    Exit Sub
End If

If kidou2 = False Then
    MsgBox "データベースファイルを開けませんでした。"
    Exit Sub
End If
Unload Me
nyuukinrirekiset2 (Format(frmtouroku2.lblid.Caption, "00000"))



End Sub

Private Sub Picture13_Click()
Unload Me

End Sub

Private Sub Picture2_Click()
Dim ssshiduhidu As String

If Trim(frmtouroku2.lblid.Caption) = "" Then
    MsgBox "個人IDを参照できません。"
    Exit Sub
End If

If kidou2 = False Then
    MsgBox "データベースファイルを開けませんでした。"
    Exit Sub
End If
Unload Me
nyuukinrirekiset2 Format(frmtouroku2.lblid.Caption, "00000"), 1

End Sub

Private Sub Picture3_Click()
Unload Me
frmvoip.Show 1

End Sub

Private Sub Picture4_Click()
Unload Me
If kidou2 = False Then
    MsgBox "データベースファイルを開けませんでした。"
    Exit Sub
End If

nyuukinrirekiset (0)






End Sub

Private Sub Picture5_Click()
Dim rirekirs As Recordset, rirekisuu As Long



sql = "SELECT keiyaku2.keiyakuid2, sonota.sonota, keiyaku2.kaishi, keiyaku2.kingaku" & _
    " FROM keiyaku2 LEFT JOIN sonota ON keiyaku2.sonotaid = sonota.sonotaid" & _
    " WHERE (((keiyaku2.kojinid)='" & frmtouroku2.lblid.Caption & "')" & _
    " AND ((keiyaku2.f2)='2')) order by keiyaku2.kaishi desc"
If FcSQlGet(rirekirs, sql, PrMsg) = False Then
    MsgBox "その他の契約履歴はありません。"
    Exit Sub
Else
    frmrireki.Caption = "その他の契約履歴"
    rirekirs.MoveLast
    rirekisuu = rirekirs.RecordCount
    With frmrireki.fggridrireki
        .Font.Size = 11
        .Rows = rirekisuu + 1
        .Cols = 4
        .ColWidth(0) = 950
        .ColWidth(1) = 1350
        .ColWidth(2) = 3000
        .ColWidth(3) = 1000
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "作成日"
        .Col = 2
        .Text = "契約種類"
        .Col = 3
        .Text = "契約金額"
    End With
    rirekirs.MoveFirst
    For i = 1 To rirekisuu
        With frmrireki.fggridrireki
            .Row = i
            .Col = 0
            .Text = rirekirs!keiyakuid2
            .Col = 1
            .Text = Format(rirekirs!kaishi, "@@@@/@@/@@")
            .Col = 2
            .Text = rirekirs!sonota
            .Col = 3
            .Text = rirekirs!kingaku
        End With
       rirekirs.MoveNext
    Next i
    rirekirs.Close
    Unload Me
    frmrireki.Show 1
End If


End Sub

Private Sub Picture6_Click()

If frmmain.chkhihyouji.Value = 1 Then
    frmyotei.lblno.Caption = "1"
Else
     frmyotei.lblno.Caption = ""
End If


Unload Me

Dim s_kojinid As String

s_kojinid = Trim(frmtouroku2.lblid.Caption)

yoteiset (s_kojinid)

frmyotei.Show 1


End Sub

Private Sub Picture7_Click()
Unload Me
frmcl2.Show 1
End Sub
