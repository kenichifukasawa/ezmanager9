VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsouken4 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "現金入金履歴"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   14370
      MouseIcon       =   "frmsouken4.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmsouken4.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   270
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   12840
      MouseIcon       =   "frmsouken4.frx":11C9
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmsouken4.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      ToolTipText     =   "入金履歴を表示します。"
      Top             =   270
      Width           =   480
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年月選択"
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
      Height          =   945
      Left            =   180
      TabIndex        =   1
      Top             =   150
      Width           =   2535
      Begin VB.ComboBox combzentsuki2 
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
         Left            =   1440
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   3
         Top             =   420
         Width           =   675
      End
      Begin VB.ComboBox combzennen 
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
         Left            =   240
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   2
         Top             =   420
         Width           =   915
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000004&
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
         Index           =   0
         Left            =   2160
         TabIndex        =   5
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000004&
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
         Index           =   0
         Left            =   1200
         TabIndex        =   4
         Top             =   540
         Width           =   255
      End
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   9825
      Left            =   90
      TabIndex        =   0
      Top             =   1260
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   17330
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblkekka 
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
      Left            =   6360
      TabIndex        =   10
      Top             =   510
      Width           =   5145
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
      Left            =   14400
      TabIndex        =   9
      Top             =   810
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
      Left            =   12600
      TabIndex        =   8
      Top             =   810
      Width           =   975
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   915
      Left            =   3060
      Top             =   210
      Width           =   11970
   End
End
Attribute VB_Name = "frmsouken4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim in2 As Integer, shitenen As Long
Dim keikanenn As Integer
keikanenn = CLng(Format(Date, "yyyy")) - 2006

For i = keikanenn To 0 Step -1
    combzennen.AddItem Format(Date, "yyyy") - i

Next
 combzennen.ListIndex = keikanenn
  
For in2 = 1 To 12
  
    frmsouken4.combzentsuki2.AddItem Format(in2, "00")
Next in2

End Sub

Private Sub Picture1_Click()

Dim nnewnewshiteibi As String, nnewnewshiteibi1 As String, nnewnewshiteibi2 As String
Dim nnewnewshiteibi5 As String, chdata(), sql_ch8 As String, rs_ch8 As Recordset
Dim sql_ch As String, rs_ch As Recordset, nowcount As Long, chusercount As Long, hhhajime As Long
Dim n_kingaku As Long
        
        frmsouken4.lblkekka.Caption = ""
        n_kingaku = 0
        
        
        frmsouken4.lv1.ListItems.Clear
        frmsouken4.lv1.ColumnHeaders.Clear
        DoEvents
        
        If frmsouken4.combzennen.ListIndex = -1 Then
            MsgBox "指定年数を指定してから実行してください。"
            Exit Sub
        End If
        If frmsouken4.combzentsuki2.ListIndex = -1 Then
            MsgBox "指定月数を指定してから実行してください。"
            Exit Sub
        End If
        
        nnewnewshiteibi = Format(frmsouken4.combzennen.Text, "0000") & Format(frmsouken4.combzentsuki2.Text, "00")
        nnewnewshiteibi1 = nnewnewshiteibi & "01"
        nnewnewshiteibi2 = nnewnewshiteibi & "31"
        If frmsouken4.combzentsuki2.Text = "01" Then
            nnewnewshiteibi5 = Format(CLng(frmsouken4.combzennen.Text) - 1, "0000") & "12"
        Else
            nnewnewshiteibi5 = Format(frmsouken4.combzennen.Text, "0000") & Format(CInt(frmsouken4.combzentsuki2.Text) - 1, "00")
        End If
        With frmsouken4
                .lv1.Checkboxes = False
                   .lv1.ColumnHeaders. _
                   Add , , "NO", 500
                   
                   .lv1.ColumnHeaders. _
                   Add , , "入金日", 1400
                   
                   .lv1.ColumnHeaders. _
                   Add , , "契約氏名", 5000
                   
                   .lv1.ColumnHeaders. _
                   Add , , "契約種類", 3200
                   
                   .lv1.ColumnHeaders. _
                   Add , , "入金額", 1400, 1
                   
                   
        End With
        
    ' ListItem オブジェクトを追加するための変数を宣言します。
    Dim itmX As ListItem
With frmsouken4
        
        If kidou2 = False Then
             MsgBox "データベースファイルを開けませんでした。"
             Exit Sub
         End If

       sql_ch = "SELECT * FROM koushinkiroku where houhou='1'" & _
               " and koushinbi between '" & nnewnewshiteibi1 & "' and '" & nnewnewshiteibi2 & "'" & _
               " order by koushinbi"
        If FcSQlGet2(rs_ch, sql_ch, PrMsg) = True Then
               rs_ch.MoveFirst
               While Not rs_ch.EOF
                    ReDim Preserve chdata(7, nowcount)
                      chdata(0, nowcount) = nowcount
                      chdata(1, nowcount) = Format(rs_ch!koushinbi, "@@@@/@@/@@")
                      chdata(4, nowcount) = rs_ch!kingaku
                      
                      n_kingaku = n_kingaku + CLng(rs_ch!kingaku)
                      
                      chdata(5, nowcount) = rs_ch!id
                      chdata(6, nowcount) = rs_ch!kojinid
                        sql_ch8 = "SELECT bank.bankmei, kojin.name, kojin.seikyuu, bankmain.bankname" & _
                            " From (bank RIGHT JOIN kojin ON bank.bankid = kojin.bankid) LEFT JOIN bankmain ON bank.bankcode = bankmain.bankcode" & _
                            " WHERE kojinid ='" & rs_ch!kojinid & "'"
                        If FcSQlGet(rs_ch8, sql_ch8, PrMsg) = False Then
                             chdata(2, nowcount) = "err" 'name
                             chdata(3, nowcount) = "err"
                        Else
                            If IsNull(rs_ch8!Name) Then
                                chdata(2, nowcount) = "err" 'name
                            Else
                                chdata(2, nowcount) = rs_ch8!Name 'name
                            End If
                            If IsNull(rs_ch8!seikyuu) Then
                                chdata(3, nowcount) = "err" 'name
                            Else
                                If Trim(rs_ch8!seikyuu) = "0" Then
                                    chdata(3, nowcount) = rs_ch8!bankname & rs_ch8!bankmei
                                Else
                                    chdata(3, nowcount) = "請求書"
                                End If
                            End If
                            rs_ch8.Close
                        End If

                            
                            nowcount = nowcount + 1
                   rs_ch.MoveNext   ' 次のレコードに移動します。
                Wend
                rs_ch.Close
        End If
        
    chusercount = nowcount - 1
                For hhhajime = 0 To nowcount - 1
                            Set itmX = .lv1.ListItems. _
                            Add(, , CStr(hhhajime + 1))
                            
                            If Trim(chdata(1, hhhajime)) <> "" Then
                                itmX.SubItems(1) = CStr(chdata(1, hhhajime))
                            End If
                            
                            If Trim(chdata(2, hhhajime)) <> "" Then
                                itmX.SubItems(2) = CStr(chdata(2, hhhajime))
                            End If
                            If Trim(chdata(3, hhhajime)) <> "" Then
                                itmX.SubItems(3) = CStr(chdata(3, hhhajime))
                            End If
                            If IsNull(chdata(4, hhhajime)) Or Trim(chdata(4, hhhajime)) <> "" Then
                                itmX.SubItems(4) = CStr(chdata(4, hhhajime))
                            End If
                 Next
                Screen.MousePointer = 0
End With


  frmsouken4.lblkekka.Caption = "入金総合計は、" & Format(CStr(n_kingaku), "\\#,##0;\\-#,##0") & " 円です。"

End Sub

Private Sub Picture13_Click()
Unload Me

End Sub
