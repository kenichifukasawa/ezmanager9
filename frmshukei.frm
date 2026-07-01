VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmshukei 
   Caption         =   "集計"
   ClientHeight    =   10185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10185
   ScaleWidth      =   14940
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   12330
      MouseIcon       =   "frmshukei.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmshukei.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   270
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "集計詳細"
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
      Height          =   8715
      Left            =   90
      TabIndex        =   4
      Top             =   1320
      Width           =   14685
      Begin MSFlexGridLib.MSFlexGrid dgvkari 
         Height          =   8295
         Left            =   60
         TabIndex        =   5
         ToolTipText     =   "金融機関一覧です。"
         Top             =   240
         Width           =   14475
         _ExtentX        =   25532
         _ExtentY        =   14631
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColor       =   16777215
         BackColorBkg    =   16777215
      End
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   13950
      MouseIcon       =   "frmshukei.frx":11FE
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmshukei.frx":1AFD
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   270
      Width           =   480
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3300
      MouseIcon       =   "frmshukei.frx":23C7
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmshukei.frx":2CC6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   270
      Width           =   480
   End
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
      Left            =   1710
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   450
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
      Left            =   510
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   450
      Width           =   915
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "CSV出力"
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
      Left            =   12060
      TabIndex        =   11
      Top             =   870
      Width           =   945
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
      Left            =   13980
      TabIndex        =   9
      Top             =   870
      Width           =   495
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
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
      Height          =   225
      Left            =   3270
      TabIndex        =   8
      Top             =   870
      Width           =   480
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
      Index           =   0
      Left            =   2430
      TabIndex        =   7
      Top             =   570
      Width           =   255
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
      Index           =   0
      Left            =   1470
      TabIndex        =   6
      Top             =   570
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   1155
      Left            =   120
      Top             =   90
      Width           =   14640
   End
End
Attribute VB_Name = "frmshukei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim in2 As Integer, shitenen As Long

For in2 = 1 To 12
    If in2 < 5 Then
        shitenen = CLng(Format(Date, "yyyy")) - 4 + in2
        frmshukei.combzennen.AddItem CStr(shitenen)
    End If
    frmshukei.combzentsuki2.AddItem Format(in2, "00")
Next in2
frmshukei.combzennen.ListIndex = 3



End Sub

Private Sub Picture13_Click()
  Unload frmshukei
   
    Set frmshukei = Nothing
End Sub

Private Sub Picture6_Click()


Dim s_nichiji As String, n_furikou As String, t_furikou As String
Dim s_hi As String, rs_sentaku As Recordset, sql_sentaku As String
Dim rs_fu As Recordset, sql_fu As String
Dim rs_fu2 As Recordset, sql_fu2 As String

Dim csv_data() As String

With frmshukei
    If .combzennen.ListIndex = -1 Then
        MsgBox "表示検索したい年を選択してください"
        Exit Sub
    Else
        n_furikou = .combzennen.Text
    End If
    If .combzentsuki2.ListIndex = -1 Then
        MsgBox "表示検索したい月を選択してください"
        Exit Sub
    Else
        t_furikou = .combzentsuki2.Text
    End If
    
    .dgvkari.Clear
    .dgvkari.Rows = 1

End With



s_hi = n_furikou & t_furikou
s_nichiji = s_hi & "01"

  If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If
    

'sql_sentaku = "SELECT * FROM shousai WHERE keiseikyuu='" & s_hi & "' order by keishurui,keiid"
sql_sentaku = "SELECT * FROM shousai WHERE keiseikyuu='" & s_hi & "' and keishurui='1' and (keiid='06' or keiid='18' or keiid='20') and kuchisuu is not null order by kojinid,keishurui,keiid"

If FcSQlGet2(rs_sentaku, sql_sentaku, PrMsg) = True Then

 On Error GoTo eee1
    datasuu = rs_sentaku.RecordCount
        ReDim Preserve csv_data(14, datasuu)
        csv_data(0, 0) = "NO"
        csv_data(1, 0) = "個人 ID"
        csv_data(2, 0) = "個人名"
        csv_data(3, 0) = "請求年月"
        csv_data(4, 0) = "契約種類"
        csv_data(5, 0) = "契約種類名"
        csv_data(6, 0) = "契約ID"
        csv_data(7, 0) = "契約名"
        csv_data(8, 0) = "金額"
        csv_data(9, 0) = "ST"
        csv_data(10, 0) = "juutaku"
        csv_data(11, 0) = "f3"
        csv_data(12, 0) = "kuchisuu"
        csv_data(13, 0) = "juutaku2"
        
    rs_sentaku.MoveFirst
    sususu = 1
    Do Until rs_sentaku.EOF
            csv_data(0, sususu) = Format(sususu, "00000")
            csv_data(1, sususu) = rs_sentaku!kojinid
            sql_fu = "select name,juutaku,f3,juutaku2 from kojin where kojinid ='" & rs_sentaku!kojinid & "'"
            If FcSQlGet(rs_fu, sql_fu, PrMsg) = True Then
                csv_data(2, sususu) = rs_fu!Name
                 csv_data(10, sususu) = rs_fu!juutaku
                 csv_data(11, sususu) = rs_fu!f3
                 If IsNull(rs_fu!juutaku2) Then
                    csv_data(13, sususu) = ""
                 Else
                    csv_data(13, sususu) = Trim(rs_fu!juutaku2)
                End If
                rs_fu.Close
            Else
                csv_data(2, sususu) = "error"
                csv_data(10, sususu) = "error"
                csv_data(11, sususu) = "error"
                csv_data(13, sususu) = "error"
            End If
            csv_data(3, sususu) = rs_sentaku!keiseikyuu
            csv_data(4, sususu) = rs_sentaku!keishurui
            If rs_sentaku!keishurui = "1" Then
                csv_data(5, sususu) = "主契約"
            ElseIf rs_sentaku!keishurui = "2" Then
                csv_data(5, sususu) = "その他の契約"
            ElseIf rs_sentaku!keishurui = "3" Then
                 csv_data(5, sususu) = "NHKの契約"
            End If
            csv_data(6, sususu) = rs_sentaku!keiid
            csv_data(7, sususu) = ""
            If rs_sentaku!keishurui = "1" Then
                 sql_fu = "select shuruimei from shurui where shuruiid ='" & rs_sentaku!keiid & "'"
                If FcSQlGet(rs_fu2, sql_fu, PrMsg) = True Then
                    csv_data(7, sususu) = rs_fu2!shuruimei
                    rs_fu2.Close
                Else
                    csv_data(7, sususu) = "error"
                End If
            ElseIf rs_sentaku!keishurui = "2" Then
                 sql_fu = "select sonota from sonota where sonotaid ='" & rs_sentaku!keiid & "'"
                If FcSQlGet(rs_fu2, sql_fu, PrMsg) = True Then
                    csv_data(7, sususu) = rs_fu2!sonota
                    rs_fu2.Close
                Else
                    csv_data(7, sususu) = "error"
                End If
            ElseIf rs_sentaku!keishurui = "3" Then
               Select Case rs_sentaku!keiid
               Case "0"
                csv_data(7, sususu) = "口座振替１ヶ月"
               Case "1"
                csv_data(7, sususu) = "口座振替２ヶ月"
               Case "2"
                csv_data(7, sususu) = "口座振替３ヶ月"
               Case "3"
                csv_data(7, sususu) = "口座振替４ヶ月"
               Case "4"
                csv_data(7, sususu) = "口座振替５ヶ月"
               Case "5"
                csv_data(7, sususu) = "口座振替６ヶ月"
                  Case "6"
                csv_data(7, sususu) = "口座振替６ヶ月ハンデ"
                  Case "7"
                csv_data(7, sususu) = "口座振替６ヶ月家族割"
               End Select
    
            End If
            csv_data(8, sususu) = rs_sentaku!keikin
            csv_data(9, sususu) = rs_sentaku!keist
            
            If IsNull(rs_sentaku!kuchisuu) Then
                csv_data(12, sususu) = "0"
            Else
                  csv_data(12, sususu) = rs_sentaku!kuchisuu
            End If
            
        sususu = sususu + 1
        rs_sentaku.MoveNext
    Loop
    rs_sentaku.Close
On Error GoTo 0
End If

Dim shutsu_path As String
'shutsu_path = "C:\Documents and Settings\ken\デスクトップ\testtest.csv"
'ファイル作成
'Open shutsu_path For Output Access Write As 1
'For writecounter = 0 To datasuu
       ' Write #1, csv_data(0, writecounter), csv_data(1, writecounter), csv_data(2, writecounter), csv_data(3, writecounter), csv_data(4, writecounter), csv_data(5, writecounter), csv_data(6, writecounter), csv_data(7, writecounter), csv_data(8, writecounter), csv_data(9, writecounter), csv_data(10, writecounter), csv_data(11, writecounter), csv_data(12, writecounter), csv_data(13, writecounter)
'Next

'Close #1

'全額免除の数の計算
Dim s_zengakumenjo_suu As Integer, rs_yotei4 As Recordset
s_zengakumenjo_suu = 0

sql = "SELECT * FROM keiyaku WHERE kadou ='0' and shurui ='25'"
                        
    If FcSQlGet(rs_yotei4, sql, PrMsg) = True Then
        If rs_yotei4.RecordCount = 0 Then
            'EZManagerに契約がない
              '  Screen.MousePointer = 0
               ' MsgBox "EZManagerには、契約データが登録されていないため、EZManagerにあって請求データにない場合のチェックは行われません。"
        Else
         
            rs_yotei4.MoveFirst
            Do Until rs_yotei4.EOF
                    
                 s_zengakumenjo_suu = s_zengakumenjo_suu + CInt(rs_yotei4!kuchisuu)
             
                rs_yotei4.MoveNext
            Loop
            rs_yotei4.Close
                
           
        End If
    Else
    
    End If




If datasuu = 0 Then
    MsgBox "指定された月は、集計データがありません。"

    Exit Sub
End If


'データを解析する
Dim kaiseki_data() As String


Dim s_x As Integer, s_y As Integer
Dim n_suu As Long, s_juutakuid2 As String

 ReDim kaiseki_data(3, 10)
 
 For i = 1 To datasuu
    Select Case csv_data(6, i)
        Case "06"
            s_x = 0
        Case "18"
            s_x = 1
        Case "20"
            s_x = 2
    End Select
    
     n_suu = CLng(csv_data(12, i))
    
    s_juutakuid2 = csv_data(13, i)
    
    If s_juutakuid2 = "" Then
        Select Case csv_data(10, i)
            Case "0"
                s_y = 0
            Case "1"
                s_y = 1
            Case "2"
                s_y = 2
            Case "3"
                s_y = 3
            Case "4"
                s_y = 4
            Case "5"
                s_y = 5
            Case "6"
                s_y = 6
            Case "7"
                s_y = 7
        End Select
    Else
         Select Case s_juutakuid2
            Case "0"
                s_y = 0
            Case "1"
                s_y = 1
            Case "2"
                s_y = 2
            Case "3"
                s_y = 3
            Case "4"
                s_y = 4
            Case "5"
                s_y = 5
            Case "6"
                s_y = 6
            Case "7"
                s_y = 7
        End Select
        
    End If
    
    
    
    If s_y = 4 Then
       If s_x = 2 Then
            If kaiseki_data(s_x, s_y) = "" Then
                 kaiseki_data(s_x, s_y) = "1"
            Else
                 kaiseki_data(s_x, s_y) = CStr(CLng(kaiseki_data(s_x, s_y)) + 1)
            End If
        Else
            If kaiseki_data(s_x, s_y) = "" Then
                 kaiseki_data(s_x, s_y) = CStr(n_suu)  '"1"
            Else
                 kaiseki_data(s_x, s_y) = CStr(CLng(kaiseki_data(s_x, s_y)) + n_suu)
            End If
        End If
    Else
        If kaiseki_data(s_x, s_y) = "" Then
             kaiseki_data(s_x, s_y) = "1"
        Else
             kaiseki_data(s_x, s_y) = CStr(CLng(kaiseki_data(s_x, s_y)) + 1)
        End If
    End If
    
 Next
 
Dim s_counter As Integer
Dim s_sousuu As Long

With frmshukei.dgvkari
        .Cols = 3
        .Rows = 25 + 2
        .ColWidth(0) = 3000
        .ColWidth(1) = 5000
        .ColWidth(2) = 1400
        
        .Row = 0
        .Col = 0
        .Text = "種類"
        .Col = 1
        .Text = "詳細"
         .Col = 2
        .Text = "件数"
        
        s_counter = 1
        s_sousuu = 0
        
        For i = 0 To 2
            For ii = 0 To 7
            .Row = s_counter
            .Col = 0
            Select Case i
                Case 0
                .Text = "CATV基本料金"
                Case 1
                .Text = "CATV基本料金半額免除"
                Case 2
                .Text = "CATV基本料金集合住宅"
            End Select
            .Col = 1
             .CellAlignment = 1
            Select Case ii
                Case 0
                .Text = "0.一般家庭"
                Case 1
                .Text = "1.一般事務所(31+2)"
                Case 2
                .Text = "2.公共(公民館・消防団)"
                Case 3
                .Text = "3.公共１(病院・福祉センター・ぶどうの丘)"
                Case 4
                If i = 2 Then
                    .Text = "4.集合住宅(福祉施設)"
                Else
                    .Text = "4.集合住宅"
                End If
                Case 5
                .Text = "5.農・事務所"
                Case 6
                .Text = "6.遊覧園"
                Case 7
                .Text = "7.その他(お試し住宅)"
            End Select
            .Col = 2
            .Text = kaiseki_data(i, ii)
            
            
            If kaiseki_data(i, ii) <> "" Then
                s_sousuu = s_sousuu + CLng(kaiseki_data(i, ii))
            End If
            
            s_counter = s_counter + 1
            
            Next ii
        Next
        
         .Row = s_counter
         .Col = 0
         .Text = "CATV基本料金全額免除"
        .Col = 2
        .Text = CStr(s_zengakumenjo_suu)
        
        
         .Row = s_counter + 1
         .Col = 0
         .Text = "*******  合計  ********"
        .Col = 2
        .Text = CStr(s_sousuu + s_zengakumenjo_suu)
        
End With

Exit Sub

eee1:
MsgBox "出力データの読み込み時にエラーが発生しました。"
Exit Sub



End Sub


Private Sub Picture8_Click()

Dim s_nichiji As String, n_furikou As String, t_furikou As String
Dim s_hi As String

With Me
    If .combzennen.ListIndex = -1 Then
        MsgBox "表示検索したい年を選択してください"
        Exit Sub
    Else
        n_furikou = .combzennen.Text
    End If
    If .combzentsuki2.ListIndex = -1 Then
        MsgBox "表示検索したい月を選択してください"
        Exit Sub
    Else
        t_furikou = .combzentsuki2.Text
    End If
    
    
    If .dgvkari.Rows = 0 Then
        MsgBox "出力したいデータがありません。"
        Exit Sub
    End If
End With



s_hi = n_furikou & t_furikou
s_nichiji = s_hi & "01"



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
On Error GoTo 0
On Error GoTo errcsv2

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
On Error GoTo 0




    If shiteppp = "" Then
        MsgBox "保存先を指定してから実行してください。"
        Exit Sub
    End If
    shiteppp = shiteppp & "\月別請求集計[" & s_hi & "]_" & Format(Date, "yyyymmdd") & "-" & Format(Time, "hhmmss") & ".csv"
    csv_csv8 (shiteppp)
Exit Sub
errcsv:
        MsgBox "ダイアログ設定エラー"
        Exit Sub
errcsv2:
        MsgBox "パス選択エラー:" & strPath
        Exit Sub

End Sub
