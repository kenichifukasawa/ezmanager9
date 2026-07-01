VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmcl2 
   Caption         =   "ケーブルライン"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   15330
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "期間指定(請求年月を指定)"
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
      Height          =   735
      Index           =   1
      Left            =   9810
      TabIndex        =   4
      Top             =   300
      Width           =   2715
      Begin VB.ComboBox cmbnen 
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
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   6
         Top             =   300
         Width           =   915
      End
      Begin VB.ComboBox cmbtsuki 
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
         Left            =   1410
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   5
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "年"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1170
         TabIndex        =   8
         Top             =   420
         Width           =   255
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "月"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2130
         TabIndex        =   7
         Top             =   420
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   14280
      MouseIcon       =   "frmcl2.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmcl2.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   360
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   13260
      MouseIcon       =   "frmcl2.frx":11C9
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmcl2.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      ToolTipText     =   "入金履歴を表示します。"
      Top             =   300
      Width           =   480
   End
   Begin VB.Frame Frame3 
      Caption         =   "請求詳細データ"
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
      Height          =   7575
      Left            =   90
      TabIndex        =   0
      Top             =   1320
      Width           =   15075
      Begin MSFlexGridLib.MSFlexGrid grdcl 
         Height          =   7125
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12568
         _Version        =   393216
         Cols            =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label lblkeisan 
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
      Height          =   735
      Left            =   540
      TabIndex        =   11
      Top             =   300
      Width           =   8775
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
      Left            =   14310
      TabIndex        =   10
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "抽出"
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
      Index           =   0
      Left            =   13020
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   1035
      Left            =   180
      Top             =   150
      Width           =   15030
   End
End
Attribute VB_Name = "frmcl2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim in2 As Integer, shitenen As Long

For in2 = 1 To 12
    If in2 < 5 Then
        shitenen = CLng(Format(Date, "yyyy")) - 4 + in2
        frmcl2.cmbnen.AddItem CStr(shitenen)
    End If
    frmcl2.cmbtsuki.AddItem Format(in2, "00")
Next in2

End Sub

Private Sub Picture1_Click()
    Dim n_nen As String, n_tsuki As String, n_hi As String
    Dim rirekirs5 As Recordset, sql5 As String, rirekisuu As Long
    Dim kazekeisan(5) As Double
    
        If cmbnen.ListIndex = -1 Then
            MsgBox "表示検索したい年を選択してください"
            Exit Sub
        Else
            n_nen = cmbnen.Text
        End If
        If cmbtsuki.ListIndex = -1 Then
            MsgBox "表示検索したい月を選択してください"
            Exit Sub
        Else
            n_tsuki = cmbtsuki.Text
        End If
        n_hi = n_nen & n_tsuki

If kidou_cl = False Then
    MsgBox "データベースファイルを開けませんでした。"
    Exit Sub
End If
sql5 = "SELECT *" & _
    " FROM cltbl" & _
    " WHERE kojinid='" & frmtouroku2.lblid.Caption & "' and nentsuki='" & n_hi & "' order by clno"
    
If FcSQlGet_cl(rirekirs5, sql5, PrMsg) <= 0 Then
    frmcl2.grdcl.Rows = 1
    frmcl2.lblkeisan.Caption = ""
    MsgBox "CLの請求データはありません。"
    db_cl.Close
    Exit Sub
Else
    rirekirs5.MoveFirst
    'Do Until rirekirs5.EOF
    rirekirs5.MoveLast
    rirekisuu = rirekirs5.RecordCount
    With frmcl2.grdcl
        .Font.Size = 11
        .Rows = rirekisuu + 1
        .ColWidth(0) = 1200
        .ColWidth(1) = 1600
        .ColWidth(2) = 1600
        .ColWidth(3) = 1600
        .ColWidth(4) = 1100
        .ColWidth(5) = 4100
        .ColWidth(6) = 1200
        .ColWidth(7) = 900
        .ColWidth(8) = 900
        .ColWidth(9) = 900
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "お客コード"
        .Col = 2
        .Text = "電話番号"
        .Col = 3
        .Text = "請求番号"
        .Col = 4
        .Text = "内訳コード"
        .Col = 5
        .Text = "内訳詳細"
        .Col = 6
        .Text = "金額"
        .Col = 7
        .Text = "税区分"
        .Col = 8
        .Text = "通話区分"
        .Col = 9
        .Text = "課金"
    End With
    kazekeisan(0) = 0
    kazekeisan(1) = 0
    kazekeisan(2) = 0
    kazekeisan(3) = 0
    kazekeisan(4) = 0
    rirekirs5.MoveFirst
    For i = 1 To rirekisuu
        With frmcl2.grdcl
            .Row = i
            .Col = 0
            .Text = rirekirs5!clno
            .Col = 1
            .Text = rirekirs5!okyukucode
            .Col = 2
            .Text = rirekirs5!tel
            .Col = 3
            .Text = rirekirs5!seikyuuno
            .Col = 4
            .Text = rirekirs5!uchiwakecode
            .Col = 5
            .Text = rirekirs5!uchiwakemeishou
            .Col = 6
            .Text = CDbl(rirekirs5!kingaku) / 100
            .Col = 7
            Select Case rirekirs5!zeikubun
            Case "F"
                .Text = "不課税"
                If rirekirs5!kakin = "1" Then
                    kazekeisan(1) = kazekeisan(1) + CDbl(rirekirs5!kingaku)
                End If
            Case "H"
                .Text = "非課税"
                If rirekirs5!kakin = "1" Then
                    kazekeisan(1) = kazekeisan(1) + CDbl(rirekirs5!kingaku)
                End If
            Case "M"
                .Text = "免税"
                If rirekirs5!kakin = "1" Then
                    kazekeisan(1) = kazekeisan(1) + CDbl(rirekirs5!kingaku)
                End If
            Case "S"
                .Text = "外税"
                If rirekirs5!kakin = "1" Then
                    kazekeisan(0) = kazekeisan(0) + CDbl(rirekirs5!kingaku)
                End If
            Case "U"
                .Text = "内税"
                If rirekirs5!kakin = "1" Then
                    kazekeisan(1) = kazekeisan(1) + CDbl(rirekirs5!kingaku)
                End If
            Case "Z"
                .Text = "その他"
                If rirekirs5!kakin = "1" Then
                    kazekeisan(1) = kazekeisan(1) + CDbl(rirekirs5!kingaku)
                End If
            Case Else
                .Text = "エラー"
                If rirekirs5!kakin = "1" Then
                    kazekeisan(1) = kazekeisan(1) + CDbl(rirekirs5!kingaku)
                End If
            End Select
            .Col = 8
            If Trim(rirekirs5!tsuuwakubun) = "" Then
                .Text = "非通話"
            Else
                .Text = "通話"
            End If
            .Col = 9
            If Trim(rirekirs5!kakin) = "0" Then
                .Text = "×"
            Else
                .Text = "○"
            End If
        End With
       rirekirs5.MoveNext
    Next i
    rirekirs5.Close
    
    If n_hi <= "201309" Then
        kazekeisan(2) = kazekeisan(0) * 0.05
    ElseIf n_hi <= "201909" Then
        kazekeisan(2) = kazekeisan(0) * 0.08
    Else
        kazekeisan(2) = kazekeisan(0) * 0.1
    End If
    
    kazekeisan(3) = kazekeisan(0) + kazekeisan(2) + kazekeisan(1)
    kazekeisan(4) = Int(kazekeisan(3) / 100)
    frmcl2.lblkeisan.Caption = "課税対象金額：" & kazekeisan(0) / 100 & "円。  消費税：" & kazekeisan(2) / 100 & "円。" & Chr(13) & _
                                "非課税対象金額：" & kazekeisan(1) / 100 & "円。  合計額：" & kazekeisan(3) / 100 & "円。" & Chr(13) & _
                                "最終請求額（端数処理後）：" & kazekeisan(4) & "円。"
                                
    db_cl.Close
End If

End Sub

Private Sub Picture13_Click()
Unload Me

End Sub
