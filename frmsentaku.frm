VERSION 5.00
Begin VB.Form frmsentaku 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "条件の選択"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7170
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6060
      MouseIcon       =   "frmsentaku.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmsentaku.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   3990
      Width           =   480
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   4530
      MouseIcon       =   "frmsentaku.frx":11C9
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmsentaku.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   3990
      Width           =   480
   End
   Begin VB.Frame Frame2 
      Caption         =   "並べ替え"
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
      Height          =   1215
      Left            =   150
      TabIndex        =   5
      Top             =   3720
      Width           =   3675
      Begin VB.OptionButton optsort 
         Caption         =   "ﾌﾘｶﾞﾅ順で結果を表示します"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   780
         Width           =   3255
      End
      Begin VB.OptionButton optsort 
         Caption         =   "契約者ID順で結果を表示します"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   3315
      End
   End
   Begin VB.Frame fshiten 
      Caption         =   "金融機関選択"
      Enabled         =   0   'False
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
      Height          =   1815
      Left            =   180
      TabIndex        =   1
      Top             =   1770
      Width           =   6855
      Begin VB.OptionButton optshiten 
         Caption         =   "下記から支店を選択します。"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   780
         Width           =   5415
      End
      Begin VB.OptionButton optshiten 
         Caption         =   "全支店を選択します。"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   5475
      End
      Begin VB.ComboBox comshiten 
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
         Height          =   345
         Left            =   180
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   2
         Top             =   1230
         Width           =   6195
      End
   End
   Begin VB.Frame f1 
      Caption         =   "支払条件選択"
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
      Height          =   1515
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   6855
      Begin VB.ComboBox combank 
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
         TabIndex        =   10
         Top             =   960
         Width           =   6135
      End
      Begin VB.OptionButton optjouken 
         Caption         =   "請求書"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3690
         TabIndex        =   9
         Top             =   390
         Width           =   1335
      End
      Begin VB.OptionButton optjouken 
         Caption         =   "金融機関"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   8
         Top             =   420
         Value           =   -1  'True
         Width           =   1335
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
      Left            =   6090
      TabIndex        =   14
      Top             =   4500
      Width           =   495
   End
   Begin VB.Label SSCJIKKOU 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "実行"
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
      Left            =   4560
      TabIndex        =   13
      Top             =   4500
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   1035
      Left            =   3960
      Top             =   3840
      Width           =   3030
   End
End
Attribute VB_Name = "frmsentaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub combank_Click()

Dim sbi As String

If combank.ListIndex <> -1 Then
    comshiten.Clear
        fshiten.Enabled = True
    optshiten(0).Value = True
    sbi = Mid(combank.Text, 1, 4)
    sql = "SELECT bank.shitencode, bank.bankmei" & _
        " FROM bank INNER JOIN kojin ON bank.bankid = kojin.bankid" & _
        " GROUP BY kojin.seikyuu, bank.bankcode, bank.shitencode, bank.bankmei" & _
         " having (((bank.bankcode)='" & sbi & "') AND ((kojin.seikyuu)='0'));"
    If FcSQlGet(rs, sql, PrMsg) = True Then
        rs.MoveFirst
        Do Until rs.EOF
            a = rs!shitencode & Space(1) & rs!bankmei
            comshiten.AddItem a
            rs.MoveNext
        Loop
        rs.Close
    End If
End If
End Sub


Private Sub Form_Load()
sbnbankset3
End Sub

Private Sub optjouken_Click(Index As Integer)
If optjouken(0).Value = True Then
    combank.Enabled = True
Else
    combank.ListIndex = -1
    combank.Enabled = False
    comshiten.ListIndex = -1
    optshiten(0).Value = True
    fshiten.Enabled = False
End If
End Sub

Private Sub optshiten_Click(Index As Integer)
If optshiten(0).Value = True Then
    comshiten.Enabled = False
Else
    comshiten.Enabled = True
End If
End Sub


Private Sub Picture13_Click()
Unload frmsentaku
Set frmsentaku = Nothing

End Sub

Private Sub Picture21_Click()
End Sub

Private Sub Picture7_Click()
Dim bi As String, si As String, kekkasuu As Long, si2 As String
Dim dono As Integer, kakikubun As String, kakishu As String



' ListItem オブジェクトを追加するための変数を宣言します。
Dim itmX As ListItem



dono = 0
kakikubun = ""
kakishu = ""
        If optjouken(0).Value = True Then   '金融機関検索
                If optsort(0).Value = True Then
                    si2 = " ORDER BY kojin.kojinid ASC;"
                Else
                    si2 = " ORDER BY kojin.furigana ASC;"
                End If
                If combank.ListIndex = -1 Then
                    MsgBox "金融機関名を選択してください。"
                    combank.SetFocus
                    Exit Sub
                End If
                bi = Mid(combank.Text, 1, 4)
                If optshiten(1).Value = True Then
                    If comshiten.ListIndex = -1 Then
                        MsgBox "支店名を選択してください。"
                        comshiten.SetFocus
                        Exit Sub
                    End If
                End If
                If optshiten(0).Value = True Then
                    si = ")"
                Else
                    Dim shitenlen As Integer
                    shitenlen = InStr(1, comshiten.Text, " ")
                    If shitenlen = 0 Then
                        si = " AND ((bank.shitencode)='" & Mid(comshiten.Text, 1, 3) & "'))"
                    Else
                        si = " AND ((bank.shitencode)='" & Mid(comshiten.Text, 1, shitenlen - 1) & "'))"
                    End If
                End If
                
                
                'sql = "SELECT kojin.*" & _
                        " FROM kojin LEFT JOIN bank ON kojin.bankid = bank.bankid" & _
                        " WHERE (((bank.bankcode)='" & bi & "')" & si
                    sql = "SELECT kojin.*,bank.bankcode,bank.shitencode,bank.bankmei" & _
                            " FROM kojin LEFT JOIN bank ON kojin.bankid = bank.bankid" & _
                            " WHERE (((bank.bankcode)='" & bi & "')" & si & si2
                    kakikubun = "金融機関"
                    kakishu = combank.Text & Space(3) & comshiten.Text
        Else    '現金検索
                If optsort(0).Value = True Then
                    si2 = " ORDER BY kojinid ASC"
                Else
                    si2 = " ORDER BY furigana ASC"
                End If
                sql = "SELECT * FROM kojin WHERE seikyuu ='1'" & si2
                dono = 2    '現金
                kakikubun = "請求書"
        End If
                
             ReDim csv_data_keiyaku(10, 0)
            csv_data_k_count = 0

        If FcSQlGet(rs, sql, PrMsg) = False Then
            MsgBox "該当する契約内容の契約者はいません。"
            Exit Sub
        Else
        
        With frmkekka
            If frmsentaku.optjouken(0).Value = True Then   '金融機関検索
            
                  ' ColumnHeaders を追加します。列の幅は、コントロールの幅を
                   .lv1.ColumnHeaders. _
                   Add , , "id", 1000
                   
                   .lv1.ColumnHeaders. _
                   Add , , "名前", 2000
                   
                   .lv1.ColumnHeaders. _
                   Add , , "ﾌﾘｶﾞﾅ", 2000
                   
                   .lv1.ColumnHeaders. _
                   Add , , "支店", 1000
                   
                   .lv1.ColumnHeaders. _
                   Add , , "種目", 700
                   
                      .lv1.ColumnHeaders. _
                   Add , , "口座", 1200
                   
                   .lv1.ColumnHeaders. _
                   Add , , "金額", 1000, lvwColumnRight
                
                   .lv1.ColumnHeaders. _
                   Add , , "更新", 700
                   
            Else  '請求書検索
                  ' ColumnHeaders を追加します。列の幅は、コントロールの幅を
                   .lv1.ColumnHeaders. _
                   Add , , "id", 1000
                   
                   .lv1.ColumnHeaders. _
                   Add , , "名前", 2000
                   
                   .lv1.ColumnHeaders. _
                   Add , , "ﾌﾘｶﾞﾅ", 2000
                          
                      .lv1.ColumnHeaders. _
                   Add , , "郵便番号", 1500
                   
                   .lv1.ColumnHeaders. _
                   Add , , "住所", 2000
                   
                   .lv1.ColumnHeaders. _
                   Add , , "電話番号", 2000
                   
                   .lv1.ColumnHeaders. _
                   Add , , "金額", 1000, lvwColumnRight
                
                   .lv1.ColumnHeaders. _
                   Add , , "更新", 700
                   
                   
                   .lv1.ColumnHeaders. _
                   Add , , "住所１", 0
            End If
          End With
          
          
          
          
          
          
          
          
                With frmkekka
               
               ' 最終レコードまで、ListItem オブジェクトを追加します。
               ' ListItem オブジェクトのテキストには、author フィールドを使用します。
               ' ListItem オブジェクトの SubItem(1) には、AuthorID を使用します。
               ' ListItem オブジェクトの SubItem(2) には、"Year of Birth" フィールドを使用します。
               .lblkensuu.Caption = rs.RecordCount
               
               csv_data_k_count = rs.RecordCount
                 ReDim csv_data_keiyaku(10, csv_data_k_count)
            
               
               
               rs.MoveFirst
               frmsentaku.Hide
               FRMOMACHI.lblmsg.Caption = "一覧を準備中..."
               FRMOMACHI.Show
               DoEvents
               Screen.MousePointer = 11
               
               Dim n_count As Integer
               
               n_count = 1
               While Not rs.EOF
                    If optjouken(0).Value = True Then   '金融機関検索
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
                            
                            If Not IsNull(rs!shitencode) Then
                                itmX.SubItems(3) = rs!shitencode
                            End If
                
                            If Not IsNull(rs!kouzakubun) Then
                                If rs!kouzakubun = "0" Then
                                    itmX.SubItems(4) = "普"
                                Else
                                    itmX.SubItems(4) = "当"
                                End If
                            End If
                            
                            If Not IsNull(rs!kouzabangou) Then
                                itmX.SubItems(5) = rs!kouzabangou
                            End If
                            
                             If Not IsNull(rs!soukei) Then
                                itmX.SubItems(6) = rs!soukei
                            End If
                           
                             If Not IsNull(rs!f1) Then
                                If rs!f1 = "0" Then
                                    itmX.SubItems(7) = "○"
                                Else
                                    itmX.SubItems(7) = "×"
                                End If
                            End If
                    Else
                            Set itmX = .lv1.ListItems. _
                            Add(, , CStr(rs!kojinid))  ' kojinid フィールド。
                          csv_data_keiyaku(0, n_count) = CStr(rs!kojinid)
                             ' AuthorID フィールドが長さ 0 の文字列以外の場合、
                             ' SubItem 1 にこの文字列を設定します。
                            If Not IsNull(rs!Name) Then
                                itmX.SubItems(1) = CStr(rs!Name)
                                 csv_data_keiyaku(1, n_count) = CStr(rs!Name)
                            End If
                            ' birth フィールドが長さ 0 の文字列以外の場合、
                            ' SubItem 2 にこの文字列を設定します。
                            If Not IsNull(rs!Furigana) Then
                                itmX.SubItems(2) = rs!Furigana
                                csv_data_keiyaku(2, n_count) = CStr(rs!Furigana)
                            End If
                            
                            If Not IsNull(rs!yuubin) Then
                                itmX.SubItems(3) = Format(rs!yuubin, "@@@-@@@@")
                                 csv_data_keiyaku(3, n_count) = Format(rs!yuubin, "@@@-@@@@")
                            End If
                            
                            
                             ' SubItem 1 にこの文字列を設定します。
                            If Not IsNull(rs!juusho2) Then
                                itmX.SubItems(4) = CStr(rs!juusho2)
                                 csv_data_keiyaku(4, n_count) = CStr(rs!juusho2)
                            End If
                            ' birth フィールドが長さ 0 の文字列以外の場合、
                            ' SubItem 2 にこの文字列を設定します。
                            If Not IsNull(rs!tel1) Then
                                itmX.SubItems(5) = rs!tel1
                                   csv_data_keiyaku(5, n_count) = CStr(rs!tel1)
                            End If
                                                
                             If Not IsNull(rs!soukei) Then
                                itmX.SubItems(6) = rs!soukei
                                csv_data_keiyaku(6, n_count) = CStr(rs!soukei)
                            End If
                           
                             If Not IsNull(rs!f1) Then
                                If rs!f1 = "0" Then
                                    itmX.SubItems(7) = "○"
                                    csv_data_keiyaku(7, n_count) = "○"
                                Else
                                    itmX.SubItems(7) = "×"
                                     csv_data_keiyaku(7, n_count) = "×"
                                End If
                            End If
                            
                            If Not IsNull(rs!juusho1) Then
                                itmX.SubItems(8) = CStr(rs!juusho1)
                                 csv_data_keiyaku(8, n_count) = CStr(rs!juusho1)
                            End If
                            
                    End If
                   rs.MoveNext   ' 次のレコードに移動します。
                   n_count = n_count + 1
                Wend
                rs.Close
                Screen.MousePointer = 0
                .lblkubun.Caption = kakikubun
                .lblshu.Caption = kakishu
                Unload FRMOMACHI
                Set FRMOMACHI = Nothing
                .Show 1
                End With
            
        End If
End Sub
