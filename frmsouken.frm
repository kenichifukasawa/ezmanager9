VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsouken 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "滞納者データ抽出"
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
   Begin VB.Frame Frame1 
      Caption         =   "抽出条件"
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
      Height          =   1185
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   15045
      Begin VB.Frame Frame4 
         Caption         =   "条件"
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
         Left            =   7830
         TabIndex        =   21
         Top             =   300
         Width           =   2685
         Begin VB.OptionButton optjouken 
            Caption         =   "回数<>0"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   1470
            TabIndex        =   23
            Top             =   330
            Width           =   1125
         End
         Begin VB.OptionButton optjouken 
            Caption         =   "未入<>0"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   180
            TabIndex        =   22
            Top             =   330
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   12840
         MouseIcon       =   "frmsouken.frx":0000
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmsouken.frx":08FF
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   18
         Top             =   270
         Width           =   480
      End
      Begin VB.Frame Frame2 
         Caption         =   "回数"
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
         Left            =   10620
         TabIndex        =   16
         Top             =   300
         Width           =   1125
         Begin VB.ComboBox cmbkaisuu 
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
            Left            =   210
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   17
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "取り扱い"
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
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   4275
         Begin VB.ComboBox cmbbank 
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
            Left            =   1650
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   15
            Top             =   210
            Width           =   2565
         End
         Begin VB.ComboBox cmbshiharai 
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
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   14
            Top             =   210
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "並べ替え"
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
         Left            =   4410
         TabIndex        =   9
         Top             =   300
         Width           =   3405
         Begin VB.OptionButton optnarabe 
            Caption         =   "降順"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   2520
            TabIndex        =   12
            Top             =   300
            Width           =   795
         End
         Begin VB.OptionButton optnarabe 
            Caption         =   "昇順"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   1710
            TabIndex        =   11
            Top             =   300
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.ComboBox cmbnarabe 
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
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   10
            Top             =   240
            Width           =   1515
         End
      End
      Begin VB.PictureBox Picture15 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   13530
         MouseIcon       =   "frmsouken.frx":11FE
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmsouken.frx":1AFD
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   7
         Top             =   270
         Width           =   480
      End
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   14280
         MouseIcon       =   "frmsouken.frx":293F
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmsouken.frx":323E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
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
         Left            =   12030
         MouseIcon       =   "frmsouken.frx":3B08
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmsouken.frx":4407
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   3
         ToolTipText     =   "入金履歴を表示します。"
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lblsql 
         BackColor       =   &H80000009&
         Height          =   195
         Left            =   10500
         TabIndex        =   20
         Top             =   150
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "CSV"
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
         Left            =   12860
         TabIndex        =   19
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         Caption         =   "印刷"
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
         Index           =   14
         Left            =   13420
         TabIndex        =   8
         Top             =   780
         Width           =   795
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
         TabIndex        =   6
         Top             =   780
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
         Left            =   11790
         TabIndex        =   5
         Top             =   780
         Width           =   975
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  '不透明
         BorderStyle     =   0  '透明
         Height          =   915
         Left            =   11790
         Top             =   180
         Width           =   3120
      End
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   9045
      Left            =   120
      TabIndex        =   1
      Top             =   2010
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   15954
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
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
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   150
      TabIndex        =   2
      Top             =   1320
      Width           =   15015
   End
End
Attribute VB_Name = "frmsouken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmbshiharai_Click()
If cmbshiharai.ListIndex <> 0 Then
    cmbbank.Clear
    cmbbank.ListIndex = -1
Else
    sbnbankset4
End If
End Sub


Private Sub Form_Load()

With frmsouken.cmbnarabe
    .AddItem "ＩＤ"
    .AddItem "ふりがな"
    .AddItem "住所"
    .AddItem "電話番号"
    .AddItem "請求額"
    .AddItem "前回入金日"
    .AddItem "ＮＨＫＮＯ"
    .ListIndex = 0
End With
With frmsouken.cmbshiharai
    .AddItem "口座"
    .AddItem "請求書"
    .AddItem "すべて"
    .ListIndex = 2
End With
With cmbkaisuu
    .AddItem ""
    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
    .ListIndex = 0
End With
cmbbank.Clear

End Sub

Private Sub Picture1_Click()

        Dim d1 As Long, d2 As Integer, d3 As String, si2 As String, kakikubun As String, kakishu As String
        Dim shuwhere As String
        Dim kenkenbankid As String
        d1 = CLng(Format(Date, "yyyy"))
        d2 = CInt(Format(Date, "mm"))
        If d2 = 1 Then
            d2 = 12
            d1 = CLng(d1) - 1
        Else
            d2 = d2 - 1
        End If
        d3 = CStr(d1) & Format(d2, "00") & "00"
        frmsouken.lv1.ListItems.Clear
        frmsouken.lv1.ColumnHeaders.Clear
        DoEvents
        
        shuwhere = ""

        '請求方法
        Select Case cmbshiharai.ListIndex
        Case 0 '口座
            If cmbbank.ListIndex <> -1 Then
                kenkenbankid = Mid(cmbbank.Text, 1, 4)
                shuwhere = shuwhere & "(kojin.seikyuu ='0' and bank.bankcode='" & kenkenbankid & "')"
            Else
                shuwhere = shuwhere & "kojin.seikyuu ='0'"
            End If
        Case 1 '請求
              shuwhere = shuwhere & "kojin.seikyuu ='1'"
        End Select
        
        If optjouken(0).Value = True Then
            If cmbkaisuu.ListIndex = 0 Then
                If shuwhere = "" Then
                    shuwhere = " kojin.fkaisuu <> '0'"
                Else
                    shuwhere = shuwhere & " and kojin.fkaisuu <> '0'"
                End If
            Else
                If shuwhere = "" Then
                    shuwhere = " kojin.fkaisuu = '" & CStr(cmbkaisuu.ListIndex) & "'"
                Else
                    shuwhere = shuwhere & " and kojin.fkaisuu = '" & CStr(cmbkaisuu.ListIndex) & "'"
                End If
            End If
        Else
            If cmbkaisuu.ListIndex = 0 Then
                If shuwhere = "" Then
                    shuwhere = " kojin.minyuukin <> '0'"
                Else
                    shuwhere = shuwhere & " and kojin.minyuukin <> '0'"
                End If
            Else
                If shuwhere = "" Then
                    shuwhere = " kojin.soukei <> '0' and kojin.minyuukin = '" & CStr(cmbkaisuu.ListIndex) & "'"
                Else
                    shuwhere = shuwhere & " and kojin.minyuukin <> '0' and kojin.fkaisuu = '" & CStr(cmbkaisuu.ListIndex) & "'"
                End If
            End If
        End If
        
        '***並べ替え****
        Select Case cmbnarabe.ListIndex
        Case 0
            shuwhere = shuwhere & " order by kojin.kojinid"
        Case 1
            shuwhere = shuwhere & " order by kojin.furigana"
        Case 2
            shuwhere = shuwhere & " order by kojin.juusho2"
        Case 3
            shuwhere = shuwhere & " order by kojin.tel1"
        Case 4
            shuwhere = shuwhere & " order by clng(kojin.soukei)"
        Case 5
            shuwhere = shuwhere & " order by kojin.maebi"
        Case 6
            shuwhere = shuwhere & " order by kojin.nhkno"
        End Select
        If optnarabe(0).Value = True Then
            shuwhere = shuwhere & " ASC;"
        Else
            shuwhere = shuwhere & " DESC;"
        End If
        
        With frmsouken
                .lv1.Checkboxes = False
                   .lv1.ColumnHeaders. _
                   Add , , "id", 800
                   
                   .lv1.ColumnHeaders. _
                   Add , , "名前", 2000
                   
                   .lv1.ColumnHeaders. _
                   Add , , "ﾌﾘｶﾞﾅ", 2000
                          
                   .lv1.ColumnHeaders. _
                   Add , , "住所", 2000
                   
                  .lv1.ColumnHeaders. _
                   Add , , "電話番号", 1650
                   
                   .lv1.ColumnHeaders. _
                   Add , , "請求金額", 1100, lvwColumnRight
                
                   .lv1.ColumnHeaders. _
                   Add , , "請求", 900
                   
                   .lv1.ColumnHeaders. _
                   Add , , "B", 400
                   .lv1.ColumnHeaders. _
                   Add , , "C", 400
                   .lv1.ColumnHeaders. _
                   Add , , "前回入金日", 1500
                   .lv1.ColumnHeaders. _
                   Add , , "銀行名", 2000
                   .lv1.ColumnHeaders. _
                   Add , , "最終結果コード", 2000
                   .lv1.ColumnHeaders. _
                   Add , , "郵便番号", 1
        End With
        
       
        
        sql2 = "SELECT kojin.*,bank.bankcode,bankmain.bankname" & _
            " FROM (bank RIGHT JOIN kojin ON bank.bankid = kojin.bankid) LEFT JOIN bankmain ON bank.bankcode = bankmain.bankcode" & _
            " where " & shuwhere
        
        lblsql.Caption = sql2
        If FcSQlGet(rs, sql2, PrMsg) = False Then
            MsgBox "該当する契約内容の契約者はいません。"
            Exit Sub
        Else
' ListItem オブジェクトを追加するための変数を宣言します。
Dim itmX As ListItem
            
                With frmsouken
                kakikubun = "請求チェック 　表示件数"
                kakishu = ""
               ' 最終レコードまで、ListItem オブジェクトを追加します。
               kakikubun = kakikubun & " (" & rs.RecordCount & "件）"
               rs.MoveFirst
               Screen.MousePointer = 11
               While Not rs.EOF
                            Set itmX = .lv1.ListItems. _
                            Add(, , CStr(rs!kojinid))  ' kojinid フィールド。
                          
                             ' AuthorID フィールドが長さ 0 の文字列以外の場合、
                             ' SubItem 1 にこの文字列を設定します。
                            If Not IsNull(rs!Name) Then
                                itmX.SubItems(1) = CStr(rs!Name)
                            End If
                            ' birth フィールドが長さ 0 の文字列以外の場合、
                            ' SubItem 2 にこの文字列を設定します。
                            If Not IsNull(rs!Furigana) Then
                                itmX.SubItems(2) = rs!Furigana
                            End If
                             ' SubItem 1 にこの文字列を設定します。
                            If Not IsNull(rs!juusho2) Then
                                itmX.SubItems(3) = CStr(rs!juusho2)
                            End If
                            ' birth フィールドが長さ 0 の文字列以外の場合、
                            ' SubItem 2 にこの文字列を設定します。
                            If Not IsNull(rs!tel1) Then
                                itmX.SubItems(4) = rs!tel1
                            End If
                                                
                             If Not IsNull(rs!soukei) Then
                                itmX.SubItems(5) = rs!soukei
                            End If
                           
                             If Not IsNull(rs!seikyuu) Then
                                If rs!seikyuu = "0" Then
                                    itmX.SubItems(6) = "口座扱"
                                ElseIf rs!seikyuu = "1" Then
                                    itmX.SubItems(6) = "請求書"
                                Else
                                    itmX.SubItems(6) = "error"
                                End If
                            End If
                            
                             If Not IsNull(rs!f2) Then
                                If rs!f1 = "1" Then
                                    itmX.SubItems(7) = "○"
                                End If
                             End If
                             If IsNull(rs!fkaisuu) Or Trim(rs!fkaisuu) = "" Then
                                    itmX.SubItems(8) = "0"
                             Else
                                itmX.SubItems(8) = Trim(rs!fkaisuu)
                             End If
                             
                             If Not IsNull(rs!maebi) And Trim(rs!maebi) <> "" Then
                                itmX.SubItems(9) = Format(rs!maebi, "@@@@/@@/@@")
                            End If
                            If Not IsNull(rs!bankname) And Trim(rs!bankname) <> "" Then
                                 itmX.SubItems(10) = rs!bankname
                            End If
                             If Not IsNull(rs!seikyuu) Then
                                If rs!kekkacode = "1" Then
                                    itmX.SubItems(11) = "資金不足"
                                ElseIf rs!kekkacode = "2" Then
                                    itmX.SubItems(11) = "取引なし"
                                ElseIf rs!kekkacode = "3" Then
                                    itmX.SubItems(11) = "預金者の都合"
                                ElseIf rs!kekkacode = "4" Then
                                    itmX.SubItems(11) = "振替依頼書なし"
                                ElseIf rs!kekkacode = "8" Then
                                    itmX.SubItems(11) = "委託者の都合"
                                ElseIf rs!kekkacode = "9" Then
                                    itmX.SubItems(11) = "その他"
                                Else
                                    itmX.SubItems(11) = "不明"
                                End If
                            End If
                             If IsNull(rs!yuubin) Or Trim(rs!yuubin) = "" Then
                                    itmX.SubItems(12) = ""
                             Else
                                If Len(Trim(rs!yuubin)) = 7 Then
                                    itmX.SubItems(12) = Format(rs!yuubin, "@@@-@@@@")
                                Else
                                    itmX.SubItems(12) = Trim(rs!yuubin)
                                End If
                             End If
                             
                   rs.MoveNext   ' 次のレコードに移動します。
                Wend
                rs.Close
                Screen.MousePointer = 0
                .Picture15.Enabled = True
                .lblkekka.Caption = kakikubun & kakishu
                End With
            
        End If
Exit Sub

errsuuji50:
    MsgBox "金額は、半角数字を入力してください。"
    Exit Sub

End Sub

Private Sub Picture13_Click()
Unload Me
End Sub

Private Sub Picture15_Click()
    
    Dim appaccess As New Access.Application
    Dim ss As Long
    
    
MsgBox "yet.."
Exit Sub
    
Screen.MousePointer = 11

db.Close

Dim dbapp As Object
    Set dbapp = appaccess.DBEngine.OpenDatabase(VBNAME, False, False, kidoupassword)
With appaccess
    .OpenCurrentDatabase VBNAME, False
    dbapp.Close
    Set dbapp = Nothing
    .Application.Visible = False
    
        'メインレポートのデータソース
        
        .DoCmd.OpenReport "minyu", acViewDesign
        .Reports!minyu.RecordSource = sql2
Screen.MousePointer = 0
        .DoCmd.OpenReport "minyu", acViewNormal
        DoEvents
        .DoCmd.Close acReport, "minyu", acSaveYes
    '.Application.Visible = False
    'Do While appaccess.Reports.Count
    'Reportが開いているのを監視
        DoEvents
    'Loop
    .CloseCurrentDatabase
    .Quit
    Set appaccess = Nothing

    kidou_main
    
    Unload Me
    
End With


End Sub

Private Sub Picture8_Click()
    Dim udtBrowseInfo As BROWSEINFO
    Dim lngIDList     As Long
    Dim strPath       As String
    Dim shiteppp As String
On Error GoTo errcsv
    '[フォルダの参照]ダイアログ情報を設定
    With udtBrowseInfo
        .hWndOwner = Me.hwnd
        .pidlRoot = CSIDL_NETWORK
        .pszDisplayName = String(MAX_PATH2, vbNullChar)
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    '[フォルダの参照]ダイアログを表示
    lngIDList = SHBrowseForFolder(udtBrowseInfo)
    If lngIDList <> 0 Then
        strPath = String(MAX_PATH, vbNullChar)
        'ITEMIDLISTをパスに変換
        SHGetPathFromIDList lngIDList, strPath
        If Left(strPath, 1) = vbNullChar Then
            'パスが空の場合はコンピュータ名を取得
            With udtBrowseInfo
                strPath = "\\" & Left(.pszDisplayName, _
                                      InStr(.pszDisplayName, _
                                            vbNullChar) - 1)
            End With
        Else
            strPath = Left(strPath, _
                           InStr(strPath, _
                                 vbNullChar) - 1)
        End If
        '結果をラベルに表示
       shiteppp = strPath
        'ITEMIDLISTを解放
        CoTaskMemFree lngIDList
    Else
        Exit Sub
    End If




    If shiteppp = "" Then
        MsgBox "保存先を指定してから実行してください。"
        Exit Sub
    End If
    shiteppp = shiteppp & "\滞納者_" & Format(Date, "yyyymmdd") & "-" & Format(Time, "hhmmss") & ".csv"
    csv_csv3 (shiteppp)
Exit Sub
errcsv:
        MsgBox "出力に失敗しました。"
        Exit Sub



End Sub
