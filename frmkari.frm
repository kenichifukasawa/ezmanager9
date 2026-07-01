VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmkari 
   Caption         =   "仮読み込み一覧"
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10245
   ScaleWidth      =   14880
   StartUpPosition =   2  '画面の中央
   Begin VB.CheckBox chknomi 
      BackColor       =   &H00FFFFFF&
      Caption         =   "引き落としできなかったもののみ表示"
      Height          =   270
      Left            =   4830
      TabIndex        =   12
      Top             =   510
      Value           =   1  'ﾁｪｯｸ
      Width           =   3465
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   12600
      MouseIcon       =   "frmkari.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmkari.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   210
      Width           =   480
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
      TabIndex        =   6
      Top             =   420
      Width           =   915
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
      TabIndex        =   5
      Top             =   420
      Width           =   675
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3300
      MouseIcon       =   "frmkari.frx":1741
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmkari.frx":2040
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   240
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   13950
      MouseIcon       =   "frmkari.frx":293F
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmkari.frx":323E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   240
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "仮読み込み詳細"
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
      TabIndex        =   0
      Top             =   1350
      Width           =   14685
      Begin MSFlexGridLib.MSFlexGrid dgvkari 
         Height          =   8295
         Left            =   60
         TabIndex        =   1
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
   Begin VB.Label SSCJIKKOU 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "出力"
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
      Left            =   12570
      TabIndex        =   11
      Top             =   780
      Width           =   480
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
      TabIndex        =   9
      Top             =   540
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
      Index           =   0
      Left            =   2430
      TabIndex        =   8
      Top             =   540
      Width           =   255
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
      TabIndex        =   7
      Top             =   840
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
      Left            =   13980
      TabIndex        =   3
      Top             =   810
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   1155
      Left            =   120
      Top             =   60
      Width           =   14640
   End
End
Attribute VB_Name = "frmkari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim in2 As Integer, shitenen As Long

For in2 = 1 To 12
    If in2 < 5 Then
        shitenen = CLng(Format(Date, "yyyy")) - 4 + in2
        frmkari.combzennen.AddItem CStr(shitenen)
    End If
    frmkari.combzentsuki2.AddItem Format(in2, "00")
Next in2
frmkari.combzennen.ListIndex = 3
End Sub

Private Sub Picture13_Click()
   Unload frmkari
   
    Set frmkari = Nothing
End Sub

Private Sub Picture6_Click()

Dim sql_kari As String, rs_kari As Recordset, n_furikou As String, t_furikou As String
Dim hi_kari As String, furikousuu As Long, furicount As Long
Dim sql_fu As String, rs_fu As Recordset, fufu As Integer, fufu2 As Integer, fudata(8, 11) As Long
Dim fy As Integer, fx As Integer, s_str As String

With frmkari
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

s_str = ""
If chknomi.Value = 1 Then
    s_str = " and st<>'0'"

End If

End With


  If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
    
        
        Screen.MousePointer = 0
        
        Exit Sub
    End If


 With frmkari.dgvkari
 
hi_kari = n_furikou & t_furikou

sql_kari = "select * from koushinkiroku2 where koushinbi ='" & hi_kari & "'" & s_str

If FcSQlGet2(rs_kari, sql_kari, PrMsg) = True Then
    furikousuu = rs_kari.RecordCount
    rs_kari.MoveFirst
        .Cols = 14
        .Rows = furikousuu + 1
        .ColWidth(0) = 1200
        .ColWidth(1) = 1000
        .ColWidth(2) = 1400
        .ColWidth(3) = 1100
        .ColWidth(4) = 2000
        .ColWidth(5) = 500
        .ColWidth(6) = 1500
        .ColWidth(7) = 1500
        .ColWidth(8) = 1200
        .ColWidth(9) = 2000
        .ColWidth(10) = 1500
        .ColWidth(11) = 500
        .ColWidth(12) = 500
        .ColWidth(13) = 1200
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "個人ID"
         .Col = 2
        .Text = "個人名"
        .Col = 3
        .Text = "金額"
        .Col = 4
        .Text = "銀行"
        .Col = 5
        .Text = "ST"
         .Col = 6
        .Text = "更新日"
        .Col = 7
        .Text = "フリガナ"
        .Col = 8
        .Text = "郵便番号"
        .Col = 9
        .Text = "住所"
        .Col = 10
        .Text = "電話番号"
        .Col = 11
        .Text = "B"
        .Col = 12
        .Text = "C"
        .Col = 13
        .Text = "前回"
    furicount = 1
    Do Until rs_kari.EOF
        .Row = furicount
        .Col = 0
        .Text = rs_kari!kariid
        .Col = 1
        .Text = rs_kari!kojinid
        .Col = 2
        sql_fu = "select * from kojin where kojinid ='" & rs_kari!kojinid & "'"
        If FcSQlGet(rs_fu, sql_fu, PrMsg) = True Then
            .Text = rs_fu!Name
              .Col = 7
            .Text = rs_fu!Furigana
              .Col = 8
             If Len(Trim(rs_fu!yuubin)) = 7 Then
                .Text = Format(rs_fu!yuubin, "@@@-@@@@")
            Else
                .Text = Trim(rs_fu!yuubin)
            End If
              .Col = 9
            .Text = rs_fu!juusho1 & rs_fu!juusho2
              .Col = 10
            .Text = rs_fu!tel1
              .Col = 11
            .Text = rs_fu!f2
              .Col = 12
            .Text = rs_fu!fkaisuu
              .Col = 13
            .Text = Format(rs_fu!maebi, "@@@@/@@/@@")
            
            rs_fu.Close
        Else
            .Text = "error"
        End If
        .Col = 3
        .Text = rs_kari!kingaku
        .Col = 4
        .CellAlignment = flexAlignRightCenter
        Select Case rs_kari!bankid
        Case "0142"
            .Text = "山梨中央銀行"
        Case "9900"
            .Text = "郵便局"
        Case "1385"
            .Text = "甲府信用金庫"
            Case "1386"
            .Text = "甲府商工信用金庫"
            Case "2377"
            .Text = "山梨県民信用組合"
            Case "5169"
            .Text = "フルーツ山梨農協"
        End Select
        .Col = 5
        .CellAlignment = flexAlignRightCenter
        Select Case rs_kari!st
        Case "0"
            .Text = "OK"
         Case Else '"1"
            .Text = "NG"
        End Select
        .Col = 6
        .Text = Format(rs_kari!koushinbi, "@@@@/@@/@@")
        
        furicount = furicount + 1
        rs_kari.MoveNext
    Loop
    rs_kari.Close
Else

With frmkari.dgvkari
        .Clear
       .Cols = 14
        .Rows = furikousuu + 1
        .ColWidth(0) = 1200
        .ColWidth(1) = 1000
        .ColWidth(2) = 1400
        .ColWidth(3) = 1100
        .ColWidth(4) = 2000
        .ColWidth(5) = 500
        .ColWidth(6) = 1500
        .ColWidth(7) = 1500
        .ColWidth(8) = 1200
        .ColWidth(9) = 2000
        .ColWidth(10) = 1500
        .ColWidth(11) = 500
        .ColWidth(12) = 500
        .ColWidth(13) = 1200
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "個人ID"
         .Col = 2
        .Text = "個人名"
        .Col = 3
        .Text = "金額"
        .Col = 4
        .Text = "銀行"
        .Col = 5
        .Text = "ST"
         .Col = 6
        .Text = "更新日"
        .Col = 7
        .Text = "フリガナ"
        .Col = 8
        .Text = "郵便番号"
        .Col = 9
        .Text = "住所"
        .Col = 10
        .Text = "電話番号"
        .Col = 11
        .Text = "B"
        .Col = 12
        .Text = "C"
        .Col = 13
        .Text = "前回"
 End With
End If
db2.Close



 


End With

End Sub

Private Sub Picture7_Click()
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
    shiteppp = shiteppp & "\滞納者_仮入金_" & Format(Date, "yyyymmdd") & "-" & Format(Time, "hhmmss") & ".csv"
    csv_csv4 (shiteppp)
Exit Sub
errcsv:
        MsgBox "出力に失敗しました。"
        Exit Sub
End Sub
