VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7D.ocx"
Begin VB.Form frmnyuukin 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "手動入金処理"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6240
   StartUpPosition =   2  '画面の中央
   Begin VB.CheckBox chktokushu 
      BackColor       =   &H00FFFFFF&
      Caption         =   "特殊入金を使用する"
      CausesValidation=   0   'False
      Height          =   195
      Left            =   3900
      TabIndex        =   9
      Top             =   900
      Width           =   2025
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   4230
      MouseIcon       =   "frmnyuukin.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmnyuukin.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   120
      Width           =   480
   End
   Begin VB.Frame Frame2 
      Caption         =   "更新日"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2715
      Begin VB.CommandButton cmdhiduke 
         Caption         =   "...."
         Height          =   375
         Left            =   1935
         TabIndex        =   6
         Top             =   420
         Width           =   375
      End
      Begin VB.TextBox txtshiteikoushinbi 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   540
         TabIndex        =   5
         Top             =   420
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5280
      MouseIcon       =   "frmnyuukin.frx":1741
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmnyuukin.frx":2040
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   150
      Width           =   480
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3300
      MouseIcon       =   "frmnyuukin.frx":290A
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmnyuukin.frx":3209
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin VSFlex7DAOCtl.VSFlexGrid gridnyuukinshousai 
      Height          =   1935
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   5955
      _cx             =   10504
      _cy             =   3413
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "特殊"
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
      Index           =   0
      Left            =   4230
      TabIndex        =   8
      Top             =   630
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
      Left            =   5310
      TabIndex        =   3
      Top             =   630
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "普通"
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
      Index           =   8
      Left            =   3300
      TabIndex        =   2
      Top             =   630
      Width           =   480
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   1095
      Left            =   3030
      Top             =   90
      Width           =   2910
   End
End
Attribute VB_Name = "frmnyuukin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdend_Click()
End Sub

Private Sub chktokushu_Click()
If chktokushu.Value = 0 Then
    Picture1.Enabled = False
    frmnyuukin.Height = 1710

Else
    Picture1.Enabled = True
    frmnyuukin.Height = 3915

End If
End Sub

Private Sub cmdhiduke_Click()
With frmhiduke
    donohi = 1
    .Top = 3000
    .Left = 3000
    .Show 1
End With

End Sub


Private Sub cmdkoushin_Click()

End Sub

Private Sub Picture1_Click()
Dim sakujosusunoka As Integer, suuchkac4 As Integer, rett
Dim kenshitesakujoid As String, shiteshitesql As String, sentakukingaku As Long
 Dim sentaku As String
Dim kohi As String, kohi3 As String
Dim koushinhouhou As Integer
    
     '入金日
        kohi = Format(txtshiteikoushinbi.Text, "yyyymmdd") '更新日　　新maebi
        If kohi = "" Then
            MsgBox "更新日を指定してください。"
            cmdhiduke.SetFocus
            Exit Sub
        Else
            kohi3 = Format(kohi, "@@@@/@@/@@")
        End If
    '個人ＩＤ
        sentaku = frmtouroku2.lblid.Caption
    '入金方法・・現金
        koushinhouhou = 1 '現金
        If frmtouroku2.lbljoukyou.Caption = "不可" Then
            MsgBox "口座扱いの顧客はＦＤＤが戻るまで更新できません。"
            Exit Sub
        End If

sentakukingaku = 0
sakujosusunoka = 0
For suuchkac4 = 1 To gridnyuukinshousai.Rows - 1
    If gridnyuukinshousai.Cell(flexcpChecked, suuchkac4, 0) = flexChecked Then
        sakujosusunoka = 1
            '入金金額
        sentakukingaku = sentakukingaku + CLng(gridnyuukinshousai.Cell(flexcpText, suuchkac4, 4))
    End If
Next

If sakujosusunoka = 0 Then
     MsgBox "入金したい項目が選択されていません。"
    Exit Sub
Else
    rett = MsgBox("『  " & kohi3 & "  』 に  " & CStr(sentakukingaku) & " 円の入金でよいですか？", vbYesNo)
    If rett = vbNo Then
        MsgBox "キャンセルしました。"
        Exit Sub
    End If
End If

If kidou2 = False Then
    MsgBox "データベースファイルを開けませんでした。"
     Exit Sub
End If

For suuchkac4 = 1 To gridnyuukinshousai.Rows - 1
    If gridnyuukinshousai.Cell(flexcpChecked, suuchkac4, 0) = flexChecked Then
        kenshitesakujoid = gridnyuukinshousai.Cell(flexcpText, suuchkac4, 0)
            'shousaiの更新
                sql = "UPDATE shousai SET  keinyu = '" & kohi & "',keist='1',bankid='" & Space(1) & "'" & _
                            " WHERE shousaiid = '" & kenshitesakujoid & "'"
    
                On Error GoTo errsoukatsunn
                     db2.Execute sql, dbSQLPassThrough
                On Error GoTo 0
    End If
Next

If a_genkin_nyuukin2(sentaku, kohi, CStr(sentakukingaku), CStr(koushinhouhou)) = False Then
    Exit Sub
End If


  'Log
    Dim ss_name As String
    ss_name = frmtouroku2.lblname.Caption
    Dim ss_id As String
        ss_id = frmtouroku2.lblid.Caption
    sagyou_msg = "入金処理(特殊) ：" & kohi3 & "に、" & CStr(sentakukingaku) & " 円の入金を処理しました。"
    log_sagyou sagyou_log_path, sagyou_msg, ss_id, ss_name
    
    
                                        
If hyouji_touroku(sentaku) = False Then
    Exit Sub
End If
MsgBox "更新しました。"
Unload frmnyuukin
Set frmnyuukin = Nothing
Exit Sub

errsoukatsunn:
    MsgBox "更新設定の書き込みに失敗しました。"
    Exit Sub

End Sub

Private Sub Picture10_Click()
Dim sentaku As String, rett
Dim kohi As String, kohi2 As String
Dim koushinhouhou As Integer
Dim koushinkingaku As Long
Dim kohi3 As String

     '入金日
        kohi = Format(txtshiteikoushinbi.Text, "yyyymmdd") '更新日　　新maebi
        If kohi = "" Then
            MsgBox "更新日を指定してください。"
            cmdhiduke.SetFocus
            Exit Sub
        Else
            kohi3 = Format(kohi, "@@@@/@@/@@")
        End If
    '個人ＩＤ
        sentaku = frmtouroku2.lblid.Caption
    '入金方法・・現金
        koushinhouhou = 1 '現金
    '入金金額
        If Trim(frmtouroku2.txtsoukei.Text) = "" Then
            MsgBox "請求金額がない場合は入金できません。"
            Unload Me
            Exit Sub
        Else
            On Error GoTo errkinga
               koushinkingaku = CLng(frmtouroku2.txtsoukei.Text)
            On Error GoTo 0
        End If
    rett = MsgBox("『  " & kohi3 & "  』 に  " & CStr(koushinkingaku) & " 円の入金でよいですか？", vbYesNo)
    If rett = vbNo Then
        MsgBox "キャンセルしました。"
        Exit Sub
    End If
        
     kohi2 = ""
      If Trim(frmtouroku2.lblmaebi.Caption) <> "" Then
           If Format(Trim(frmtouroku2.lblmaebi.Caption), "yyyymm") <> Format(Date, "yyyymm") Then
                    kohi2 = Format(Trim(frmtouroku2.lblmaebi.Caption), "yyyymm")
            End If
    End If
        
        If frmtouroku2.lbljoukyou.Caption = "不可" Then
            MsgBox "口座扱いの顧客はＦＤＤが戻るまで更新できません。"
            Exit Sub
        End If
            
     '請求書の扱いの場合は必要
     'If frmtouroku2.lblseikyuu.Caption = "1" Then
            'データ更新
              If shousaiset(sentaku, kohi, koushinkingaku, kohi2, Space(1)) = -1 Then
                  'MsgBox "詳細データベースへの書き込みに失敗しました。"
                  Screen.MousePointer = 0
                  Exit Sub
              End If
    'End If
              If a_genkin_nyuukin(sentaku, kohi, CStr(koushinkingaku), CStr(koushinhouhou)) = False Then
                  Exit Sub
              End If
                    
                    
            If hyouji_touroku(sentaku) = False Then
                Exit Sub
            End If

    'Log
    Dim ss_name As String
    ss_name = frmtouroku2.lblname.Caption
    sagyou_msg = "入金処理(総額) ：" & kohi3 & "に、" & CStr(koushinkingaku) & " 円の入金を処理しました。"
    log_sagyou sagyou_log_path, sagyou_msg, sentaku, ss_name
    
    
        MsgBox "更新しました。"
        Unload frmnyuukin
        Set frmnyuukin = Nothing

Exit Sub

errkinga:
    MsgBox "請求金額が不正です。"
    Exit Sub

End Sub

Private Sub Picture13_Click()
Unload frmnyuukin
Set frmnyuukin = Nothing

End Sub
