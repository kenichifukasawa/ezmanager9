VERSION 5.00
Object = "{271E16FB-2B6D-11D4-B06D-0090991097BF}#26.0#0"; "ZoonFIO.ocx"
Begin VB.Form frmkoushin 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "各種更新・確認"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   7875
   StartUpPosition =   2  '画面の中央
   Begin VB.CheckBox chkmushi 
      Caption         =   "電話番号のチェック以外は行わないで、データを読み込みを実行する。"
      Height          =   585
      Left            =   690
      TabIndex        =   20
      Top             =   6930
      Width           =   4605
   End
   Begin VB.OptionButton optkoushin 
      Caption         =   "予定管理のデータの更新を行います。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   9
      Left            =   240
      TabIndex        =   19
      Top             =   3393
      Width           =   5025
   End
   Begin VB.CheckBox chkkaihi 
      Caption         =   "EZManager上になければならない主契約のチェックを無視し、請求しないように回避する。"
      Height          =   585
      Left            =   660
      TabIndex        =   18
      Top             =   6300
      Width           =   4605
   End
   Begin VB.OptionButton optkoushin 
      Caption         =   "CLのエラーデータ(再度戻ってきた)を読み込みます。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   270
      TabIndex        =   17
      Top             =   7620
      Width           =   5025
   End
   Begin VB.OptionButton optkoushin 
      Caption         =   "CLのデータを山梨CATVと勝沼CTAV用に分割します。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   7
      Left            =   240
      TabIndex        =   16
      Top             =   5437
      Width           =   5265
   End
   Begin VB.TextBox txtcl 
      Enabled         =   0   'False
      Height          =   345
      IMEMode         =   4  '全角ひらがな
      Left            =   6150
      TabIndex        =   15
      Top             =   4020
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.OptionButton optkoushin 
      Caption         =   "CLのデータ(上記で作成したもの)を読み込みます。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   240
      TabIndex        =   14
      Top             =   5955
      Width           =   5025
   End
   Begin VB.Frame Frame1 
      Caption         =   "入力先"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   240
      TabIndex        =   10
      Top             =   90
      Width           =   7515
      Begin VB.PictureBox Picture21 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   6630
         MouseIcon       =   "frmkoushin.frx":0000
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmkoushin.frx":08FF
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   12
         Top             =   300
         Width           =   480
      End
      Begin VB.TextBox txtde 
         Height          =   345
         IMEMode         =   4  '全角ひらがな
         Left            =   180
         TabIndex        =   11
         Top             =   600
         Width           =   5925
      End
      Begin VB.Label Label7 
         Alignment       =   2  '中央揃え
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "参照"
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
         Left            =   6615
         TabIndex        =   13
         Top             =   810
         Width           =   495
      End
      Begin VB.Shape Shape15 
         BackStyle       =   1  '不透明
         BorderStyle     =   0  '透明
         Height          =   885
         Left            =   6450
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.OptionButton optkoushin 
      Caption         =   "金融機関用FDDの読み出しチェックを行います。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   240
      TabIndex        =   9
      Top             =   2882
      Width           =   4665
   End
   Begin VB.OptionButton optkoushin 
      Caption         =   "VoIPのデータ(峡東CATVからの)をチェックします。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   4415
      Width           =   5025
   End
   Begin VB.OptionButton optkoushin 
      Caption         =   "VoIPのデータ(峡東CATVからの)を読み込みます。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   4926
      Width           =   5025
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5880
      MouseIcon       =   "frmkoushin.frx":1741
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmkoushin.frx":2040
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   6810
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6930
      MouseIcon       =   "frmkoushin.frx":290A
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmkoushin.frx":3209
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   6870
      Width           =   480
   End
   Begin VB.OptionButton optkoushin 
      Caption         =   "PPVのデータ(山梨CATVからの)を読み込みます。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   3904
      Width           =   5025
   End
   Begin VB.OptionButton optkoushin 
      Caption         =   "金融機関用FDDの内容を更新します。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   2371
      Width           =   4000
   End
   Begin VB.OptionButton optkoushin 
      Caption         =   "金融機関用FDDの内容を確認をします。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1860
      Value           =   -1  'True
      Width           =   5565
   End
   Begin ZoonFIO.ZoonFileIO fd2 
      Left            =   7110
      Top             =   1650
      _ExtentX        =   900
      _ExtentY        =   900
   End
   Begin VB.Label Label1 
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
      Index           =   8
      Left            =   5910
      TabIndex        =   6
      Top             =   7350
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
      Left            =   6930
      TabIndex        =   4
      Top             =   7350
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   855
      Left            =   5700
      Top             =   6780
      Width           =   1890
   End
End
Attribute VB_Name = "frmkoushin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub data_yamanashi_to_katsunuma()

Dim kita_fdd As String, sql_out As String, rs_out As Recordset, shitey As Long, shitem As Long
Dim sql_out2 As String, rs_out2 As Recordset, newtsuwaid As String, resresre
Dim sql_out3 As String, rs_out3 As Recordset, vbname_out2 As String
Dim sql_out4 As String, rs_out4 As Recordset

'kita_fdd = "a:\out.mdb"
kita_fdd = fd_a & "out.mdb"

On Error GoTo errfdd
If Dir(kita_fdd) = "" Then
        MsgBox "データファイルがありません。フロッピーディスクを確認してください。"
        Exit Sub
End If
On Error GoTo 0
On Error Resume Next
vbname_out = sopa & "FDBACK\out"

 If Dir$(vbname_out, vbDirectory) = "" Then
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder (vbname_out)
    Set fso = Nothing
End If
shitey = CLng(Format(Date, "yyyy"))
shitem = CLng(Format(Date, "mm"))
If shitem = 1 Then
   shitey = shitey - 1
   shitem = 12
Else
    shitem = shitem - 1
End If
vbname_out2 = vbname_out & "\" & Format(shitey, "0000") & Format(shitem, "00") & "_out.mdb"
If Dir(vbname_out2) <> "" Then
    MsgBox "すでに更新済みのデータです。"
    Exit Sub
    'resresre = MsgBox("すでに更新済みのデータです。再更新しますか？", vbYesNo)
    'If resresre = vbNo Then
      '  Exit Sub
    'End If
    'Kill vbname_out2
    'DoEvents
End If
vbname_out = vbname_out & "\out.mdb"

If Dir(vbname_out) <> "" Then
    Kill vbname_out
End If

FileCopy kita_fdd, vbname_out

If kidou_out = False Then
    Screen.MousePointer = 0
    MsgBox "データＯＰＥＮに失敗しました。"
    Exit Sub
End If
If kidou_voip = False Then
    Screen.MousePointer = 0
    MsgBox "データＯＰＥＮに失敗しました。"
    Exit Sub
End If
If kidou_ppv = False Then
    Screen.MousePointer = 0
    MsgBox "データＯＰＥＮに失敗しました。"
    Exit Sub
End If

BeginTrans

sql_out = "select * from voip_m"

If FcSQlGet_out(rs_out, sql_out, PrMsg) > 0 Then
    If soukatsu(0, "startup", 9, "", 1, 1) = False Then
        Screen.MousePointer = 0
        Kill vbname_out
        MsgBox "設定の読み込みに失敗しました。"
        Exit Sub
    End If
    If reg = "0" Then
        reg = ""
    End If
    
    If reg = vbNullString Then
        newtsuwaid = "000000001"
    Else
        newtsuwaid = Format$(reg, "00000000#")
    End If
    rs_out.MoveFirst
    Do Until rs_out.EOF
        sql = "INSERT INTO voip_m VALUES('" & newtsuwaid & "'," & _
                       "'" & rs_out!voipno & "','" & rs_out!nentsuki & _
                       "','" & rs_out!koumoku2 & "'," & _
                       rs_out!kingaku & "," & rs_out!zeigaku & "," & rs_out!voipgaku & ",'" & _
                       rs_out!code & "','" & rs_out!kojinid & "','" & rs_out!st_st & "')"
        On Error GoTo errfko
        db_voip.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        newtsuwaid = Format$(CStr(CLng(newtsuwaid) + 1), "00000000#")
        rs_out.MoveNext
    Loop
    rs_out.Close
        newtsuwaid = CStr(CLng(newtsuwaid))
        If soukatsu(1, "startup", 9, newtsuwaid, 1, 1) = False Then
        Screen.MousePointer = 0
            Rollback
            Kill vbname_out
                MsgBox "設定の書き込みに失敗しました。"
                Exit Sub
        End If
End If
sql_out4 = "select * from seikyuu"

If FcSQlGet_out(rs_out4, sql_out4, PrMsg) > 0 Then
    If soukatsu(0, "startup", 6, "", 1, 1) = False Then
        Screen.MousePointer = 0
        Kill vbname_out
        Rollback
        MsgBox "設定の読み込みに失敗しました。"
        Exit Sub
    End If
    If reg = "0" Then
        reg = ""
    End If
    
    If reg = vbNullString Then
        newtsuwaid = "00000001"
    Else
        newtsuwaid = Format$(reg, "0000000#")
    End If
    rs_out4.MoveFirst
    Do Until rs_out4.EOF
        sql = "INSERT INTO seikyuu VALUES('" & newtsuwaid & "'," & _
                       "'" & rs_out4!nentsuki & "','" & rs_out4!code & _
                       "','" & rs_out4!nencode & "','" & rs_out4!hajime & "','" & rs_out4!owari & _
                       "'," & rs_out4!seikyuugaku & "," & rs_out4!zei & ",'" & rs_out4!hi & "')"
        On Error GoTo errfko
        db_voip.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        newtsuwaid = Format$(CStr(CLng(newtsuwaid) + 1), "0000000#")
        rs_out4.MoveNext
    Loop
    rs_out4.Close
        newtsuwaid = CStr(CLng(newtsuwaid))
        If soukatsu(1, "startup", 6, newtsuwaid, 1, 1) = False Then
        Screen.MousePointer = 0
            Rollback
            Kill vbname_out
                MsgBox "設定の書き込みに失敗しました。"
                Exit Sub
        End If
End If

sql_out2 = "select * from ppv"

If FcSQlGet_out(rs_out2, sql_out2, PrMsg) > 0 Then
    If soukatsu(0, "startup", 10, "", 1, 1) = False Then
        Screen.MousePointer = 0
        Rollback
        Kill vbname_out
        MsgBox "設定の読み込みに失敗しました。"
        Exit Sub
    End If
    If reg = "0" Then
        reg = ""
    End If
    
    If reg = vbNullString Then
        newtsuwaid = "00000001"
    Else
        newtsuwaid = Format$(reg, "0000000#")
    End If
    rs_out2.MoveFirst
    Do Until rs_out2.EOF
        sql = "INSERT INTO ppv VALUES('" & newtsuwaid & "','" & rs_out2!nentsuki & "','" & rs_out2!kojinid & _
                       "','" & rs_out2!icno & _
                       "','" & rs_out2!keiyakubi & "','" & rs_out2!shouhincode & _
                       "','" & rs_out2!sabisuid & "','" & rs_out2!housoubi & _
                       "','" & rs_out2!kounyuubi & "','" & rs_out2!shuushuubi & _
                       "','" & rs_out2!meishou & "'," & CLng(rs_out2!kingaku) & _
                       "," & CLng(rs_out2!zeiritsu) & ",'" & rs_out2!rokugaken & _
                       "','" & rs_out2!adaruto & "','" & rs_out2!st & "')"
        On Error GoTo errfko
        db_ppv.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        newtsuwaid = Format$(CStr(CLng(newtsuwaid) + 1), "0000000#")
        rs_out2.MoveNext
    Loop
    rs_out2.Close
        newtsuwaid = CStr(CLng(newtsuwaid))
        If soukatsu(1, "startup", 10, newtsuwaid, 1, 1) = False Then
        Screen.MousePointer = 0
            Rollback
            Kill vbname_out
                MsgBox "設定の書き込みに失敗しました。"
                Exit Sub
        End If
End If

sql_out3 = "select * from keiyaku2"

If FcSQlGet_out(rs_out3, sql_out3, PrMsg) > 0 Then
    rs_out3.MoveFirst
    Do Until rs_out3.EOF
        If voip_koushin(rs_out3!sonotaid, rs_out3!kojinid, rs_out3!kingaku, rs_out3!kaishi) = -1 Then
            Screen.MousePointer = 0
            Kill vbname_out
            Rollback
            Exit Sub
        End If
        rs_out3.MoveNext
    Loop
    rs_out3.Close
End If

db_out.Close
DoEvents

CommitTrans

Name vbname_out As vbname_out2


    'Log
       
        sagyou_msg = "PPV更新：PPVの振替データを読込、更新しました。"
        log_sagyou sagyou_log_path, sagyou_msg


MsgBox "更新完了しました。"
Exit Sub

errfdd:
    MsgBox "フロッピーディスクを確認してください。"
    Resume
errfko:
    MsgBox "データの登録に失敗しました。"
    Rollback
    Kill vbname_out
    Exit Sub

End Sub

Private Sub Form_Load()
If kenchan = 0 Then
    If fddkoushinOK = 1 Then
        optkoushin(1).Enabled = False
        optkoushin(2).Enabled = False
        optkoushin(3).Enabled = False
    End If
End If

Dim DesktopPath As String
' デスクトップのパスを取得
DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
' 表示
'MsgBox DesktopPath
txtde.Text = DesktopPath & "\EZManager_data\入力\"

Dim s_path5 As String, s_path5_1 As String, s_path5_2 As String


s_path5 = DesktopPath & "\EZManager_data\cl\"

'なければ作成
'FileSystemObjectインスタンスを生成
Dim fi As New FileSystemObject

'フォルダの存在確認
If fi.FolderExists(s_path5) Then
    'Call MsgBox("sampleフォルダは存在します")
Else
    'Call MsgBox("sampleフォルダは存在しません")
    ' フォルダ (ディレクトリ) を作成する
    MkDir (s_path5)
    DoEvents
End If

'オブジェクトの解放
Set fi = Nothing


txtcl.Text = s_path5


s_path5_2 = DesktopPath & "\EZManager_data\cl\yamanashi\"

'なければ作成
'FileSystemObjectインスタンスを生成
Dim fi_1 As New FileSystemObject

'フォルダの存在確認
If fi_1.FolderExists(s_path5_2) Then
    'Call MsgBox("sampleフォルダは存在します")
Else
    'Call MsgBox("sampleフォルダは存在しません")
    ' フォルダ (ディレクトリ) を作成する
    MkDir (s_path5_2)
    DoEvents
End If

'オブジェクトの解放
Set fi_1 = Nothing

s_path5_1 = DesktopPath & "\EZManager_data\cl\katsunuma\"

'なければ作成
'FileSystemObjectインスタンスを生成
Dim fi_2 As New FileSystemObject

'フォルダの存在確認
If fi_2.FolderExists(s_path5_1) Then
    'Call MsgBox("sampleフォルダは存在します")
Else
    'Call MsgBox("sampleフォルダは存在しません")
    ' フォルダ (ディレクトリ) を作成する
    MkDir (s_path5_1)
    DoEvents
End If

'オブジェクトの解放
Set fi_2 = Nothing



End Sub

Private Sub Picture10_Click()

'FDの更新処理
Dim motofile As String, sakifile As String, Ret As Integer, kensars As Recordset
Dim hdatas(7) As String, hdatal(8) As Long
Dim ddatas() As String, ddatal() As Long
Dim tdatas(1) As String, tdatal(6) As Long, kongetsu As String
Dim koushinid As String, koushinnewhi As String, newfilemei As String
Dim yuubinkakunin As Integer, bkrs4 As Recordset
    Dim nou As Integer
Dim motofile2 As String
Dim seikyuudata(), shousaidata(), tsuwadata(), FilePath(3)

Dim itakushacode As Integer

Dim s_path As String


yuubinkakunin = 0
nou = 0

frmkoushin.Hide

If optkoushin(3).Value = True Then
        If kidou2 = False Then
            MsgBox "データベースファイルを開けませんでした。"
            Exit Sub
        End If
        sql = "select * from bankkiroku where hidu ='" & Format(Date, "yyyymm") & "'"
        If FcSQlGet2(bkrs4, sql, PrMsg) = False Then
            '未請求状態
            MsgBox "今月の請求が行われていません。"
            Exit Sub
        Else
            If bkrs4.RecordCount <> 4 Then
                MsgBox "今月の請求が全件行われていません。"
                Exit Sub
            Else
                bkrs4.MoveFirst
                Do Until bkrs4.EOF
                    If bkrs4!nyusuu = "0" And bkrs4!fusuu = "0" Then
                        MsgBox "今月の銀行のFDの処理が全件行われていません。"
                        Exit Sub
                    End If
                    bkrs4.MoveNext
                Loop
                bkrs4.Close
                 data_kyoutou_to_katsunuma (0)
                        
            End If
        End If
ElseIf optkoushin(9).Value = True Then
'***************   予定管理の更新　　************************
    
    msg = "0更新し金融機関の更新後に実行できます。0更新、銀行の更新処理は終了しましたか？"
    res = MsgBox(msg, vbYesNo)
    If res = vbNo Then
           Exit Sub
    Else
        msg = "予定管理の更新を開始しますか？"
        res = MsgBox(msg, vbYesNo)
        If res = vbNo Then
               Exit Sub
        End If
    End If
    
    If yotejikkou2(Format(Date, "yyyymm")) = -1 Then
         MsgBox "予定管理の更新中にエラーが発生しました。"
    Else
          MsgBox "予定管理の更新が完了しました。"
    End If
     Screen.MousePointer = 0
    Exit Sub
        
ElseIf optkoushin(7).Value = True Then
    
 ' **************   CL  分割  *********************************
    
    settei_res = setting(1, 14, 0, "", 1)
    If settei_res = "-1" Or settei_res = "0" Then
            MsgBox "DropBoxの設定がありません。"
            Exit Sub
    Else
        s_path = Trim(settei_res)
    End If




    cl_data_bunkatsu (s_path)
    
        
ElseIf optkoushin(6).Value = True Then

'**************   CL   読込 *********************************

    settei_res = setting(1, 14, 0, "", 1)
    If settei_res = "-1" Or settei_res = "0" Then
            MsgBox "DropBoxの設定がありません。"
            Exit Sub
    Else
         s_path = Trim(settei_res)
    End If


 If chkmushi.Value = 0 Then
    cl_data_convert (s_path) 'チェックする
Else
    cl_data_convert2 (s_path) 'チェックしない
End If


ElseIf optkoushin(8).Value = True Then

'**************   CL   読込 *********************************

    settei_res = setting(1, 14, 0, "", 1)
    If settei_res = "-1" Or settei_res = "0" Then
            MsgBox "DropBoxの設定がありません。"
            Exit Sub
    Else
         s_path = Trim(settei_res)
    End If


    cl_data_convert s_path, "1"

ElseIf optkoushin(4).Value = True Then

    data_kyoutou_to_katsunuma (1)
ElseIf optkoushin(2).Value = True Then
        
    data_yamanashi_to_katsunuma
    
    
ElseIf optkoushin(0).Value = True Then

'ここから
       
       
        If Trim(txtde.Text) = "" Then
            MsgBox "入力先の設定がありません。"
            Exit Sub
        End If
        motofile2 = Trim(txtde.Text)
        If Right(motofile2, 1) <> "\" Then
            motofile2 = motofile2 & "\"
        End If
            
        motofile = motofile2 & "KOUFURI.dat"
        On Error GoTo errfdnashi
        If Dir(motofile) = "" Then
           ' motofile = "a:\KOUFURI"
            motofile = motofile2 & "KOUFURI"
            nou = 1
            If Dir(motofile) = "" Then
                'motofile = "a:\JIHARAI"
                motofile = motofile2 & "JIHARAI"
                nou = 2
            End If
        End If
        On Error GoTo 0
        

'ここまで



   
    With fd2
        .File_Name = motofile
        .File_Type = 0
        .Active = True
        Select Case .Err
        Case -9
            MsgBox "データファイルがありません。フロッピーディスクを確認してください。ＺＯＯＭ"
                Exit Sub
        Case -11
            MsgBox "フロッピーディスクがセットされていません。再度入れてから実行してください。ＺＯＯＭ"
            Exit Sub
        End Select
        'ヘッダーレコードの読み込み
        If nou = 0 Then
            Ret = .File_Read("D1D2D1D10X40D4D4X15D3X15D1D7X17", _
                hdatal(0), hdatal(1), hdatal(2), hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3))
        
            '読み込みエラーが発生したら、郵便局を参照する。
             'エラー発生
            If Ret < 0 Then
                If .File_Seek(1) = False Then
                    MsgBox "読み込み位置の移動に失敗しました。ＺＯＯＭ"
                    .Active = False
                    Kill sakifile
                    Exit Sub
                End If
    
                Ret = .File_Read("D1D2X1D2D8X40D4D4X15D3X17D6X4D1X12", _
                    hdatal(0), hdatal(1), hdatas(4), hdatal(2), hdatal(3), _
                    hdatas(0), hdatal(4), hdatal(5), _
                    hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatas(5), hdatal(8), hdatas(3))
                
                yuubinkakunin = 1
            End If
        ElseIf nou = 1 Then
            Ret = .File_Read("D1D2D1D10X40D4D4X15D3X15D1D7X17x2", _
                hdatal(0), hdatal(1), hdatal(2), hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3), hdatas(6))
        Else
            Ret = .File_Read("D1D2D1D10X40D4D4X15D3X15D1D7X17", _
                hdatal(0), hdatal(1), hdatal(2), hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3))
        End If
    
        'エラー発生
        If Ret < 0 Then
            errorfd (Ret)
            Exit Sub
        End If
        
        If hdatal(5) = "1386" Then
            nou = 0
        End If
        
        If nou = 0 Then
            '会社の確認
                sql = "select * from bankmain where bankcode ='" & Format(hdatal(5), "0000") & "'"
                If FcSQlGet(kensars, sql, PrMsg) = False Then
                    MsgBox "フロッピーディスクから金融機関データを読み込めません。"
                    .Active = False
                    Exit Sub
                Else
                'トレーラコードまでのデータ更新
                    If .File_Seek(FileLen(motofile) - 239) = False Then
                        MsgBox "読み込み位置の移動に失敗しました。"
                        Exit Sub
                    End If
                    Ret = .File_Read("D1D6D12D6D12D6D12X65", _
                    tdatal(0), tdatal(1), tdatal(2), tdatal(3), tdatal(4), tdatal(5), tdatal(6), tdatas(0))
                    'エラー発生
                    If Ret < 0 Then
                        errorfd (Ret)
                        Exit Sub
                    End If
            
                    'データの表示
                    If tdatal(3) = 0 And tdatal(4) = 0 And tdatal(5) = 0 And tdatal(6) = 0 Then
                        msg = "このデータは繰越処理されていないデータです。"
                    Else
                        msg = "このデータは繰越されているデータです。"
                    End If
                    msg = msg & Chr(13) & "(" & kensars!bankname & "　の " & Format(Mid(Format(hdatal(4), "0000"), 1, Len(hdatal(4)) - 2), "00") & _
                        "/" & Right(hdatal(4), 2) & "　振り替え日用のデータです）"
                    MsgBox (msg)
                
                    kensars.Close
                End If
                
        ElseIf nou = 1 Then
                
                'トレーラコードまでのデータ更新
                    If .File_Seek(FileLen(motofile) - 243) = False Then
                        MsgBox "読み込み位置の移動に失敗しました。"
                        Exit Sub
                    End If
                    Ret = .File_Read("D1D6D12D6D12D6D12X65x2", _
                    tdatal(0), tdatal(1), tdatal(2), tdatal(3), tdatal(4), tdatal(5), tdatal(6), tdatas(0), tdatas(1))
                    'エラー発生
                    If Ret < 0 Then
                        errorfd (Ret)
                        Exit Sub
                    End If
            
                    'データの表示
                    If tdatal(3) = 0 And tdatal(4) = 0 And tdatal(5) = 0 And tdatal(6) = 0 Then
                        msg = "このデータは繰越処理されていないデータです。"
                    Else
                        msg = "このデータは繰越されているデータです。"
                    End If
                    msg = msg & Chr(13) & "(農業協同組合　の " & Format(Mid(Format(hdatal(4), "0000"), 1, Len(hdatal(4)) - 2), "00") & _
                        "/" & Right(hdatal(4), 2) & "　振り替え日用のデータです）"
                    MsgBox (msg)
        Else
            '会社の確認
                sql = "select * from bankmain where bankcode ='" & Format(hdatal(5), "0000") & "'"
                If FcSQlGet(kensars, sql, PrMsg) = False Then
                    MsgBox "フロッピーディスクから金融機関データを読み込めません。"
                    .Active = False
                    Exit Sub
                Else
                'トレーラコードまでのデータ更新
                    If .File_Seek(FileLen(motofile) - 243) = False Then
                        MsgBox "読み込み位置の移動に失敗しました。"
                        Exit Sub
                    End If
                    Ret = .File_Read("D1D6D12D6D12D6D12X65", _
                    tdatal(0), tdatal(1), tdatal(2), tdatal(3), tdatal(4), tdatal(5), tdatal(6), tdatas(0))
                    'エラー発生
                    If Ret < 0 Then
                        errorfd (Ret)
                        Exit Sub
                    End If
            
                    'データの表示
                    If tdatal(3) = 0 And tdatal(4) = 0 And tdatal(5) = 0 And tdatal(6) = 0 Then
                        msg = "このデータは繰越処理されていないデータです。"
                    Else
                        msg = "このデータは繰越されているデータです。"
                    End If
                    msg = msg & Chr(13) & "(" & kensars!bankname & "　の " & Format(Mid(Format(hdatal(4), "0000"), 1, Len(hdatal(4)) - 2), "00") & _
                        "/" & Right(hdatal(4), 2) & "　振り替え日用のデータです）"
                    MsgBox (msg)
                
                    kensars.Close
                End If
        End If
                
                
        .Active = False
                
                
    End With
    Exit Sub

ElseIf optkoushin(1).Value = True Then
    '****************FDデータのOpen************************************************
    
    Dim s_bankmei As String
    
    With fd2
    
    
    'ここから
       
       
        If Trim(txtde.Text) = "" Then
            MsgBox "入力先の設定がありません。"
            Exit Sub
        End If
        motofile2 = Trim(txtde.Text)
        If Right(motofile2, 1) <> "\" Then
            motofile2 = motofile2 & "\"
        End If
            
          nou = 0
         FRMOMACHI.lblmsg.Caption = "更新の準備中..."
         FRMOMACHI.Show
         DoEvents
         Screen.MousePointer = 11
        
        
        motofile = motofile2 & "KOUFURI.dat"
        On Error GoTo errfdnashi
        If Dir(motofile) = "" Then
           ' motofile = "a:\KOUFURI"
            motofile = motofile2 & "KOUFURI"
            nou = 1
            If Dir(motofile) = "" Then
                'motofile = "a:\JIHARAI"
                motofile = motofile2 & "JIHARAI"
                nou = 2
            End If
        End If
        On Error GoTo 0
        

'ここまで

    
   
            
                .File_Name = motofile
                .File_Type = 0
                .Active = True
                    Select Case .Err
                    Case -9
                        MsgBox "データファイルがありません。フロッピーディスクを確認してください。"
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Exit Sub
                    Case -11
                        MsgBox "フロッピーディスクがセットされていません。再度入れてから実行してください。"
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Exit Sub
                    End Select
                .Active = False
            'FDのコピーをHDDに実施
            sakifile = sopa & "fdback\"
            If Dir$(sakifile, vbDirectory) = "" Then
                Dim fso
                Set fso = CreateObject("Scripting.FileSystemObject")
                fso.CreateFolder (sakifile)
                Set fso = Nothing
            End If
            If Right(sakifile, 1) <> "\" Then
                sakifile = sakifile & "\"
            End If
            
            kongetsu = CInt(Format(Date, "mm"))
            
            Dim koushoufile As String
            
            koushoufile = sakifile
            
            newfilemei = sakifile
            
            sakifile = sakifile & "sumple.dat"
            FileCopy motofile, sakifile
                
            '再度OPEN
                .File_Name = sakifile
                .File_Type = 0
                .Active = True
                    Select Case .Err
                    Case -9, -11
                        MsgBox "フロッピーディスクのデータか、ＨＤＤのデータが破損しています。"
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Kill sakifile
                        Exit Sub
                    End Select
            'ヘッダーレコードの読み込み
                    If nou = 0 Then
                        Ret = .File_Read("D1D2D1D1D9X40D4D4X15D3X15D1D7X17", _
                            hdatal(0), hdatal(1), hdatal(2), itakushacode, hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                            hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3))
                    
                           'エラーが発生したら、郵便局を参照
                         'エラー発生
                        If Ret < 0 Then
                            If .File_Seek(1) = False Then
                                MsgBox "読み込み位置の移動に失敗しました。"
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                
                            Ret = .File_Read("D1D2X1D2D8X40D4D4X15D3X17D6X4D1X12", _
                                hdatal(0), hdatal(1), hdatas(4), hdatal(2), hdatal(3), _
                                hdatas(0), hdatal(4), hdatal(5), _
                                hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatas(5), hdatal(8), hdatas(3))
                            
                            
                            
                            yuubinkakunin = 1
                        End If
                    ElseIf nou = 1 Then
                        Ret = .File_Read("D1D2D1D10X40D4D4X15D3X15D1D7X17x2", _
                            hdatal(0), hdatal(1), hdatal(2), hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                            hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3), hdatas(6))
                    Else
                        Ret = .File_Read("D1D2D1D10X40D4D4X15D3X15D1D7X17", _
                            hdatal(0), hdatal(1), hdatal(2), hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                            hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3))
                    End If
                    
                    
                    If Ret < 0 Then
                        errorfd (Ret)
                        .Active = False
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Kill sakifile
                        Exit Sub
                    End If
                    If hdatal(5) = "1386" Then
                        nou = 0
                        Dim koushoufiledat As String
                        koushoufile = koushoufile & "koufuri"
                        koushoufiledat = koushoufile & ".dat"
                        FileCopy motofile, koushoufile
                        Name koushoufile As koushoufiledat
                        .Active = False
                        Kill sakifile
                        Name koushoufiledat As sakifile
                        .File_Name = sakifile
                        .File_Type = 0
                        .Active = True
                            Select Case .Err
                            Case -9, -11
                                MsgBox "フロッピーディスクのデータか、ＨＤＤのデータが破損しています。"
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End Select
                        
                    End If
                            
                            
            '会社の確認
                
                If nou = 1 Then
                    s_bankmei = "農業協同組合"
                    msg = "金融機関「農業協同組合」の" & Format(Mid(Format(hdatal(4), "0000"), 1, Len(hdatal(4)) - 2), "00") & "/" & Right(hdatal(4), 2) & "データを読み込み、更新しますか？"
                    res = MsgBox(msg, vbYesNo)
                    If res = vbNo Then
                        MsgBox "中止しました。"
                        .Active = False
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Kill sakifile
                        Exit Sub
                    End If
                Else 'If nou = 0 Then
                        sql = "select * from bankmain where bankcode ='" & Format(hdatal(5), "0000") & "'"
                        If FcSQlGet(kensars, sql, PrMsg) = False Then
                            MsgBox "フロッピーディスクから金融機関データを読み込めません。"
                            .Active = False
                            Unload FRMOMACHI
                            Set FRMOMACHI = Nothing
                            Screen.MousePointer = 0
                            Kill sakifile
                            Exit Sub
                        Else
                             s_bankmei = kensars!bankname
                            msg = "金融機関「" & kensars!bankname & "」の" & Format(Mid(Format(hdatal(4), "0000"), 1, Len(hdatal(4)) - 2), "00") & "/" & Right(hdatal(4), 2) & "データを読み込み、更新しますか？"
                            kensars.Close
                            res = MsgBox(msg, vbYesNo)
                            If res = vbNo Then
                                MsgBox "中止しました。"
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                        End If
                
                
                End If
                
                 FRMOMACHI.lblmsg.Caption = "データの読み込みを準備中..."
            DoEvents
                If nou = 1 Then
                    'トレーラコードまでのデータ更新
                    Dim modoriatai As Long
                  On Error GoTo errmodori
                    modoriatai = 243 'CLng(InputBox("数字を入力", "値"))
                    On Error GoTo 0
                    If modoriatai = 244 Or modoriatai = 243 Then
                    
                    Else
                        MsgBox "２４３か２４４の数字を半角で入力してください。"
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Kill sakifile
                        Exit Sub
                    End If
                    If .File_Seek(FileLen(sakifile) - modoriatai) = False Then
                        MsgBox "読み込み位置の移動に失敗しました。"
                        .Active = False
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Kill sakifile
                        Exit Sub
                    End If
                    'レコードカウントの取得
        
                    Ret = .File_Read("D1D6D12D6D12D6D12X65x2", _
                    tdatal(0), tdatal(1), tdatal(2), tdatal(3), tdatal(4), tdatal(5), tdatal(6), tdatas(0), tdatas(1))
                    'エラー発生
                    If Ret < 0 Then
                        errorfd (Ret)
                        .Active = False
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Kill sakifile
                        Exit Sub
                    End If
                    
                        'データの表示
                        If tdatal(3) = 0 And tdatal(4) = 0 And tdatal(5) = 0 And tdatal(6) = 0 Then
                            MsgBox "このデータは繰越処理されていないデータです。"
                            .Active = False
                            Unload FRMOMACHI
                            Set FRMOMACHI = Nothing
                            Screen.MousePointer = 0
                            Kill sakifile
                            Exit Sub
                        End If
                  
                        
                        If .File_Seek(1) = False Then
                            MsgBox "読み込み位置の移動に失敗しました。"
                            .Active = False
                            Unload FRMOMACHI
                            Set FRMOMACHI = Nothing
                            Screen.MousePointer = 0
                            Kill sakifile
                            Exit Sub
                        End If
                            
            Ret = .File_Read("D1D2D1D10X40D4D4X15D3X15D1D7X17x2", _
                hdatal(0), hdatal(1), hdatal(2), hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3), hdatas(6))
                    If Ret < 0 Then
                        errorfd (Ret)
                        .Active = False
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Kill sakifile
                        Exit Sub
                    End If
                        
                 FRMOMACHI.pb.Visible = True
                 FRMOMACHI.pb.Min = 0
                 FRMOMACHI.pb.Max = tdatal(1) + 1
                 FRMOMACHI.pb.Value = 0
                 FRMOMACHI.lblmsg.Caption = "データの読み込み中..."
                 DoEvents
                            ReDim ddatas(tdatal(1), 11) As String
                            ReDim ddatal(tdatal(1), 8) As Long
                      For i = 1 To tdatal(1)
                            Ret = .File_Read("D1D4X15D3X15X4D1D7X30D10D1D20D1X8x2", _
                            ddatal(i, 0), ddatal(i, 1), ddatas(i, 0), ddatal(i, 2), ddatas(i, 1), ddatas(i, 2), _
                            ddatal(i, 3), ddatal(i, 4), ddatas(i, 3), ddatal(i, 5), ddatal(i, 6), _
                            ddatal(i, 7), ddatal(i, 8), ddatas(i, 4), ddatas(i, 11))
                                
                            'エラー発生
                            If Ret < 0 Then
                                errorfd (Ret)
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                                FRMOMACHI.pb.Value = i
                                DoEvents
   
                            
                      Next i
                       
                        
                        
                        
                ElseIf nou = 0 Then
                        'トレーラコードまでのデータ更新
                            If .File_Seek(FileLen(sakifile) - 239) = False Then
                                MsgBox "読み込み位置の移動に失敗しました。"
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                        'レコードカウントを取得
                            Ret = .File_Read("D1D6D12D6D12D6D12X65", _
                            tdatal(0), tdatal(1), tdatal(2), tdatal(3), tdatal(4), tdatal(5), tdatal(6), tdatas(0))
                            'エラー発生
                            If Ret < 0 Then
                                errorfd (Ret)
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                .Active = False
                                Kill sakifile
                                Exit Sub
                            End If
                            
                            If tdatal(3) = 0 And tdatal(4) = 0 And tdatal(5) = 0 And tdatal(6) = 0 Then
                                MsgBox "このデータは繰越処理されていないデータです。"
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                            
                            
                            If .File_Seek(1) = False Then
                                MsgBox "読み込み位置の移動に失敗しました。"
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                            
                            If yuubinkakunin = 0 Then
                                    Ret = .File_Read("D1D2D1D1D9X40D4D4X15D3X15D1D7X17", _
                                        hdatal(0), hdatal(1), hdatal(2), itakushacode, hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                                        hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3))
                            Else
                                    
                                    Ret = .File_Read("D1D2X1D2D8X40D4D4X15D3X17D6X4D1X12", _
                                        hdatal(0), hdatal(1), hdatas(4), hdatal(2), hdatal(3), _
                                        hdatas(0), hdatal(4), hdatal(5), _
                                        hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatas(5), hdatal(8), hdatas(3))
                            End If
                            
                            If Ret < 0 Then
                                errorfd (Ret)
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                 FRMOMACHI.pb.Visible = True
                 FRMOMACHI.pb.Min = 0
                 FRMOMACHI.pb.Max = tdatal(1)
                 FRMOMACHI.pb.Value = 0
                 FRMOMACHI.lblmsg.Caption = "データの読み込み中..."
            DoEvents
                    
                        
                            'ReDim ddatas(tdatal(1) - 1, 10) As String
                            'ReDim ddatal(tdatal(1) - 1, 8) As Long
                            'For I = 0 To tdatal(1) - 1
                        On Error GoTo erryomikomi
                            ReDim ddatas(tdatal(1), 10) As String
                            ReDim ddatal(tdatal(1), 8) As Long
                            For i = 1 To tdatal(1)
                                If yuubinkakunin = 1 Then
                                    Ret = .File_Read("D1D4X15D3X15X2X2X1D7X30D10X1D20D1X4X2X2", _
                                    ddatal(i, 0), ddatal(i, 1), ddatas(i, 0), ddatal(i, 2), ddatas(i, 1), ddatas(i, 2), _
                                    ddatas(i, 5), ddatas(i, 6), _
                                    ddatal(i, 3), ddatas(i, 7), ddatal(i, 4), ddatas(i, 3), ddatal(i, 7), ddatal(i, 8), _
                                    ddatas(i, 9), ddatas(i, 8), ddatas(i, 4))
                                
                                Else
                                    Ret = .File_Read("D1D4X15D3X15X4D1D7X30D10D1D20D1X8", _
                                    ddatal(i, 0), ddatal(i, 1), ddatas(i, 0), ddatal(i, 2), ddatas(i, 1), ddatas(i, 2), _
                                    ddatal(i, 3), ddatal(i, 4), ddatas(i, 3), ddatal(i, 5), ddatal(i, 6), _
                                    ddatal(i, 7), ddatal(i, 8), ddatas(i, 4))
                                End If
                                'エラー発生
                                If Ret < 0 Then
                                    errorfd (Ret)
                                    .Active = False
                                    Unload FRMOMACHI
                                    Set FRMOMACHI = Nothing
                                    Screen.MousePointer = 0
                                    Kill sakifile
                                    Exit Sub
                                End If
                                FRMOMACHI.pb.Value = i
                                DoEvents
            
                            Next i
                          On Error GoTo 0
            Else
                Dim dami_crlf As String
                        'トレーラコードまでのデータ更新
                            If .File_Seek(FileLen(sakifile) - 243) = False Then
                                MsgBox "読み込み位置の移動に失敗しました。"
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                        'レコードカウントを取得
                            Ret = .File_Read("D1D6D12D6D12D6D12X65X2", _
                            tdatal(0), tdatal(1), tdatal(2), tdatal(3), tdatal(4), tdatal(5), tdatal(6), tdatas(0), dami_crlf)
                            'エラー発生
                            If Ret < 0 Then
                                errorfd (Ret)
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                .Active = False
                                Kill sakifile
                                Exit Sub
                            End If
                            
                            If tdatal(3) = 0 And tdatal(4) = 0 And tdatal(5) = 0 And tdatal(6) = 0 Then
                                MsgBox "このデータは繰越処理されていないデータです。"
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                            
                            
                            If .File_Seek(1) = False Then
                                MsgBox "読み込み位置の移動に失敗しました。"
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                            
                            If yuubinkakunin = 0 Then
                                    Ret = .File_Read("D1D2D1D10X40D4D4X15D3X15D1D7X17X2", _
                                        hdatal(0), hdatal(1), hdatal(2), hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                                        hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3), dami_crlf)
                            Else
                                    
                                    Ret = .File_Read("D1D2X1D2D8X40D4D4X15D3X17D6X4D1X12", _
                                        hdatal(0), hdatal(1), hdatas(4), hdatal(2), hdatal(3), _
                                        hdatas(0), hdatal(4), hdatal(5), _
                                        hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatas(5), hdatal(8), hdatas(3))
                            End If
                            
                            If Ret < 0 Then
                                errorfd (Ret)
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                 FRMOMACHI.pb.Visible = True
                 FRMOMACHI.pb.Min = 0
                 FRMOMACHI.pb.Max = tdatal(1)
                 FRMOMACHI.pb.Value = 0
                 FRMOMACHI.lblmsg.Caption = "データの読み込み中..."
            DoEvents
                    
                        
                            'ReDim ddatas(tdatal(1) - 1, 10) As String
                            'ReDim ddatal(tdatal(1) - 1, 8) As Long
                            'For I = 0 To tdatal(1) - 1
                            
                        On Error GoTo erryomikomi
                            ReDim ddatas(tdatal(1), 10) As String
                            ReDim ddatal(tdatal(1), 8) As Long
                            For i = 1 To tdatal(1)
                                If yuubinkakunin = 1 Then
                                    Ret = .File_Read("D1D4X15D3X15X2X2X1D7X30D10X1D20D1X4X2X2", _
                                    ddatal(i, 0), ddatal(i, 1), ddatas(i, 0), ddatal(i, 2), ddatas(i, 1), ddatas(i, 2), _
                                    ddatas(i, 5), ddatas(i, 6), _
                                    ddatal(i, 3), ddatas(i, 7), ddatal(i, 4), ddatas(i, 3), ddatal(i, 7), ddatal(i, 8), _
                                    ddatas(i, 9), ddatas(i, 8), ddatas(i, 4))
                                
                                Else
                                    Ret = .File_Read("D1D4X15D3X15X4D1D7X30D10D1D20D1X8X2", _
                                    ddatal(i, 0), ddatal(i, 1), ddatas(i, 0), ddatal(i, 2), ddatas(i, 1), ddatas(i, 2), _
                                    ddatal(i, 3), ddatal(i, 4), ddatas(i, 3), ddatal(i, 5), ddatal(i, 6), _
                                    ddatal(i, 7), ddatal(i, 8), ddatas(i, 4), dami_crlf)
                                End If
                                'エラー発生
                                If Ret < 0 Then
                                    errorfd (Ret)
                                    .Active = False
                                    Unload FRMOMACHI
                                    Set FRMOMACHI = Nothing
                                    Screen.MousePointer = 0
                                    Kill sakifile
                                    Exit Sub
                                End If
                                FRMOMACHI.pb.Value = i
                                DoEvents
            
                            Next i
                          On Error GoTo 0
            End If
            
            
                        'ファイルを閉じる
                            .Active = False
                        'ファイル名を変更
                         If nou = 1 Then
                            newfilemei = newfilemei & "5169" & Format(Date, "yyyy") & Format(kongetsu, "00") & "r.dat"
                         Else
                            newfilemei = newfilemei & Format(hdatal(5), "0000") & Format(Date, "yyyy") & Format(kongetsu, "00") & "r.dat"
                            End If
                        On Error GoTo erronajifile
                            Name sakifile As newfilemei
                        On Error GoTo 0
            FRMOMACHI.pb.Value = 0
            
If rocked("on") = False Then
        Unload FRMOMACHI
        MsgBox "個人データを更新中のフラッグを立てるのに失敗しました。"
        Set FRMOMACHI = Nothing
        Kill newfilemei
        Screen.MousePointer = 0
        Exit Sub
End If
'*********トランザクション開始**************
BeginTrans
                           
             FRMOMACHI.lblmsg.Caption = "更新中..."
            DoEvents
                        koushinnewhi = Format(Date, "yyyy") & Format(Mid(Format(hdatal(4), "0000"), 1, Len(hdatal(4)) - 2), "00") & Right(hdatal(4), 2)
                        '個人データの処理
                            Dim kaisuu As Long
                            'If nou = 1 Then
                                'kaisuu = dsuu(0)
                            'Else
                                kaisuu = tdatal(1)
                            'End If
                            'For I = 0 To tdatal(1) - 1
                            For i = 1 To kaisuu
                                koushinid = Format(ddatal(i, 7), "00000")
                               
                                If yuubinkakunin = 1 Then
                                    If a_fdd_koushin2(ddatal(i, 8), koushinid, koushinnewhi, 0, ddatal(i, 4)) = 9 Then
                                        Unload FRMOMACHI
                                        MsgBox "個人データを更新中にエラーが発生しました。データを復旧します。" & "error no " & ddatal(i, 7)
                                        Set FRMOMACHI = Nothing
                                        '*******ロールバック*************
                                        Rollback
                                        Kill newfilemei
                                        Screen.MousePointer = 0
                                        If rocked("off") = False Then
                                            MsgBox "個人データを更新中のフラッグの初期化に失敗しました。"
                                        End If
                                        Exit Sub
                                    End If
                                Else
                                    If a_fdd_koushin2(ddatal(i, 8), koushinid, koushinnewhi, 0, ddatal(i, 5)) = 9 Then
                                        Unload FRMOMACHI
                                        MsgBox "個人データを更新中にエラーが発生しました。データを復旧します。" & "error no " & ddatal(i, 7)
                                        Set FRMOMACHI = Nothing
                                        '*******ロールバック*************
                                        Rollback
                                        Kill newfilemei
                                        Screen.MousePointer = 0
                                        If rocked("off") = False Then
                                            MsgBox "個人データを更新中のフラッグの初期化に失敗しました。"
                                        End If
                                        Exit Sub
                                    End If
                                End If
                                
                                
                              
                                
                                    FRMOMACHI.pb.Value = i
                                    DoEvents
            
                            Next i
                            Unload FRMOMACHI
                            Set FRMOMACHI = Nothing
                            Screen.MousePointer = 0
                            'Bank記録作成
                            If kidou2 = False Then
                                MsgBox "データベースファイルを開けませんでした。"
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                '*******ロールバック*************
                                Rollback
                                Kill newfilemei
                                Screen.MousePointer = 0
                                If rocked("off") = False Then
                                    MsgBox "個人データを更新中のフラッグの初期化に失敗しました。"
                                End If
                                Exit Sub
                            End If
                            
                            Dim subnen As String, subtsuki As String, subhiduke As String, subbankid As String
                            Dim bkrs As Recordset
                            
                            subnen = Format(Date, "yyyy")
                            subtsuki = Mid(Format(hdatal(4), "0000"), 1, 2)
                            subhiduke = subnen & subtsuki
                            If nou <> 1 Then
                                subbankid = Format(hdatal(5), "0000")
                            Else
                                subbankid = "5169"
                            End If
                            sql = "select * from bankkiroku where hidu ='" & subhiduke & "' and bankid ='" & subbankid & "'"
                            If FcSQlGet2(bkrs, sql, PrMsg) = False Then
                                        Unload FRMOMACHI
                                        Set FRMOMACHI = Nothing
                                        '*******ロールバック*************
                                        Rollback
                                        Kill newfilemei
                                        MsgBox "銀行の移動記録の更新に失敗しました。"
                                        Screen.MousePointer = 0
                                        If rocked("off") = False Then
                                            MsgBox "個人データを更新中のフラッグの初期化に失敗しました。"
                                        End If
                                        Exit Sub
                            Else
                                '更新
                                    'If nou = 0 Then
                                    
                                        sql = "UPDATE bankkiroku SET nyusuu = '" & CStr(CCur(tdatal(3))) & "'," & _
                                                     "nyukin = '" & CStr(CCur(tdatal(4))) & "'," & _
                                                      "fusuu = '" & CStr(CCur(tdatal(5))) & "'," & _
                                                       "fukin = '" & CStr(CCur(tdatal(6))) & "'" & _
                                                     " WHERE id = '" & bkrs!id & "'"
                                    'Else
                                        'sql = "UPDATE bankkiroku SET nyusuu = '" & CStr(nounyuudata(0)) & "'," & _
                                                     "nyukin = '" & CStr(nounyuudata(1)) & "'," & _
                                                      "fusuu = '" & CStr(nounyuudata(2)) & "'," & _
                                                       "fukin = '" & CStr(nounyuudata(3)) & "'" & _
                                                     " WHERE id = '" & bkrs!id & "'"
                                 
                                    'End If
                              On Error GoTo errm1
                                 db2.Execute sql, dbSQLPassThrough
                              On Error GoTo 0
                                bkrs.Close
                            End If
                            
                            
                            
                            
                            FRMOMACHI.pb.Max = kaisuu
                              FRMOMACHI.pb.Value = 0
                                    DoEvents
                        'For i = 1 To kaisuu
                                'koushinid = Format(ddatal(i, 7), "00000")
                            
                                'If yotejikkou(Format(Date, "yyyymm"), koushinid) = -1 Then
                                       ' Unload FRMOMACHI
                                      '  MsgBox "個人データ(Yotei)を更新中にエラーが発生しました。データを復旧します。" & "error no " & ddatal(i, 7)
                                      '  Set FRMOMACHI = Nothing
                                        '*******ロールバック*************
                                       ' Rollback
                                       ' Kill newfilemei
                                       ' Screen.MousePointer = 0
                                      '  Exit Sub
                               ' End If
                                
                                   ' FRMOMACHI.pb.Value = i - 1
                                   ' DoEvents
                            
                        ' Next i
                            
    '*************トランザクション終了********
    CommitTrans
                            If rocked("off") = False Then
                                MsgBox "個人データを更新中のフラッグの初期化に失敗しました。"
                            End If
                            'db2.Clos
                            
                            'Log
                               
                                sagyou_msg = "FDD更新：" & s_bankmei & "の振替データを読込、更新しました。"
                                log_sagyou sagyou_log_path, sagyou_msg
            
            
                            MsgBox "更新処理が終了しました。"
            
            End With
            
            Exit Sub
            
ElseIf optkoushin(5).Value = True Then

      '****************FDデータのOpen************************************************
      
  


      
      
    With fd2
     
         'ここから
       
       
        If Trim(txtde.Text) = "" Then
            MsgBox "入力先の設定がありません。"
            Exit Sub
        End If
        motofile2 = Trim(txtde.Text)
        If Right(motofile2, 1) <> "\" Then
            motofile2 = motofile2 & "\"
        End If
            
               FRMOMACHI.lblmsg.Caption = "更新の準備中..."
             FRMOMACHI.Show
             DoEvents
             Screen.MousePointer = 11
            
            
        motofile = motofile2 & "KOUFURI.dat"
        On Error GoTo errfdnashi
        If Dir(motofile) = "" Then
           ' motofile = "a:\KOUFURI"
            motofile = motofile2 & "KOUFURI"
            nou = 1
            If Dir(motofile) = "" Then
                'motofile = "a:\JIHARAI"
                motofile = motofile2 & "JIHARAI"
                nou = 2
            End If
        End If
        On Error GoTo 0
        

'ここまで
     
     
                .File_Name = motofile
                .File_Type = 0
                .Active = True
                    Select Case .Err
                    Case -9
                        MsgBox "データファイルがありません。フロッピーディスクを確認してください。"
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Exit Sub
                    Case -11
                        MsgBox "フロッピーディスクがセットされていません。再度入れてから実行してください。"
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Exit Sub
                    End Select
                .Active = False
            'FDのコピーをHDDに実施
            sakifile = sopa & "fdback\"
            If Dir$(sakifile, vbDirectory) = "" Then
       
                Set fso = CreateObject("Scripting.FileSystemObject")
                fso.CreateFolder (sakifile)
                Set fso = Nothing
            End If
            If Right(sakifile, 1) <> "\" Then
                sakifile = sakifile & "\"
            End If
            
            
            
            
            koushoufile = sakifile
            
            newfilemei = sakifile
            
            sakifile = sakifile & "sumple.dat"
            FileCopy motofile, sakifile
                
            '再度OPEN
                .File_Name = sakifile
                .File_Type = 0
                .Active = True
                    Select Case .Err
                    Case -9, -11
                        MsgBox "フロッピーディスクのデータか、ＨＤＤのデータが破損しています。"
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Kill sakifile
                        Exit Sub
                    End Select
            'ヘッダーレコードの読み込み
                    If nou = 0 Then
                        Ret = .File_Read("D1D2D1D1D9X40D4D4X15D3X15D1D7X17", _
                            hdatal(0), hdatal(1), hdatal(2), itakushacode, hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                            hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3))
                    
                           'エラーが発生したら、郵便局を参照
                         'エラー発生
                        If Ret < 0 Then
                            If .File_Seek(1) = False Then
                                MsgBox "読み込み位置の移動に失敗しました。"
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                            
                            Ret = .File_Read("D1D2D1D1D9X40D4D4X15D3X15D1D7X17", _
                            hdatal(0), hdatal(1), hdatal(2), itakushacode, hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                            hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3))
                            
                            
                                'エラー発生
                            If Ret < 0 Then
                                If .File_Seek(1) = False Then
                                    MsgBox "読み込み位置の移動に失敗しました。"
                                    .Active = False
                                    Unload FRMOMACHI
                                    Set FRMOMACHI = Nothing
                                    Screen.MousePointer = 0
                                    Kill sakifile
                                    Exit Sub
                                End If
                                
                                
                    
                                Ret = .File_Read("D1D2X1D2D8X40D4D4X15D3X17D6X4D1X12", _
                                    hdatal(0), hdatal(1), hdatas(4), hdatal(2), hdatal(3), _
                                    hdatas(0), hdatal(4), hdatal(5), _
                                    hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatas(5), hdatal(8), hdatas(3))
                                
                                
                                
                                yuubinkakunin = 1
                            End If
                        End If
                    ElseIf nou = 1 Then
                        Ret = .File_Read("D1D2D1D10X40D4D4X15D3X15D1D7X17x2", _
                            hdatal(0), hdatal(1), hdatal(2), hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                            hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3), hdatas(6))
                    Else
                        Ret = .File_Read("D1D2D1D10X40D4D4X15D3X15D1D7X17", _
                            hdatal(0), hdatal(1), hdatal(2), hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                            hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3))
                    End If
                    
                    
                    If Ret < 0 Then
                        errorfd (Ret)
                        .Active = False
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Kill sakifile
                        Exit Sub
                    End If
                    If hdatal(5) = "1386" Then
                        nou = 0
              
                        koushoufile = koushoufile & "koufuri"
                        koushoufiledat = koushoufile & ".dat"
                        FileCopy motofile, koushoufile
                        Name koushoufile As koushoufiledat
                        .Active = False
                        Kill sakifile
                        Name koushoufiledat As sakifile
                        .File_Name = sakifile
                        .File_Type = 0
                        .Active = True
                            Select Case .Err
                            Case -9, -11
                                MsgBox "フロッピーディスクのデータか、ＨＤＤのデータが破損しています。"
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End Select
                        
                    End If
                            
                            
            '会社の確認
                
                If nou = 1 Then
                    msg = "金融機関「農業協同組合」の" & Format(Mid(Format(hdatal(4), "0000"), 1, Len(hdatal(4)) - 2), "00") & "/" & Right(hdatal(4), 2) & "データをチェックしますか？"
                    res = MsgBox(msg, vbYesNo)
                    If res = vbNo Then
                        MsgBox "中止しました。"
                        .Active = False
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Kill sakifile
                        Exit Sub
                    End If
                Else 'If nou = 0 Then
                        sql = "select * from bankmain where bankcode ='" & Format(hdatal(5), "0000") & "'"
                        If FcSQlGet(kensars, sql, PrMsg) = False Then
                            MsgBox "フロッピーディスクから金融機関データを読み込めません。"
                            .Active = False
                            Unload FRMOMACHI
                            Set FRMOMACHI = Nothing
                            Screen.MousePointer = 0
                            Kill sakifile
                            Exit Sub
                        Else
                            msg = "金融機関「" & kensars!bankname & "」の" & Format(Mid(Format(hdatal(4), "0000"), 1, Len(hdatal(4)) - 2), "00") & "/" & Right(hdatal(4), 2) & "データをチェックしますか？"
                            kensars.Close
                            res = MsgBox(msg, vbYesNo)
                            If res = vbNo Then
                                MsgBox "中止しました。"
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                        End If
                
                
                End If
                
                 FRMOMACHI.lblmsg.Caption = "データの読み込みを準備中..."
            DoEvents
                If nou = 1 Then
                    'トレーラコードまでのデータ更新
           
                  On Error GoTo errmodori
                    modoriatai = 243 'CLng(InputBox("数字を入力", "値"))
                    On Error GoTo 0
                    If modoriatai = 244 Or modoriatai = 243 Then
                    
                    Else
                        MsgBox "２４３か２４４の数字を半角で入力してください。"
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Kill sakifile
                        Exit Sub
                    End If
                    If .File_Seek(FileLen(sakifile) - modoriatai) = False Then
                        MsgBox "読み込み位置の移動に失敗しました。"
                        .Active = False
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Kill sakifile
                        Exit Sub
                    End If
                    'レコードカウントの取得
        
                    Ret = .File_Read("D1D6D12D6D12D6D12X65x2", _
                    tdatal(0), tdatal(1), tdatal(2), tdatal(3), tdatal(4), tdatal(5), tdatal(6), tdatas(0), tdatas(1))
                    'エラー発生
                    If Ret < 0 Then
                        errorfd (Ret)
                        .Active = False
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Kill sakifile
                        Exit Sub
                    End If
                    
                    
                        
                        If .File_Seek(1) = False Then
                            MsgBox "読み込み位置の移動に失敗しました。"
                            .Active = False
                            Unload FRMOMACHI
                            Set FRMOMACHI = Nothing
                            Screen.MousePointer = 0
                            Kill sakifile
                            Exit Sub
                        End If
                            
                    Ret = .File_Read("D1D2D1D10X40D4D4X15D3X15D1D7X17x2", _
                        hdatal(0), hdatal(1), hdatal(2), hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                        hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3), hdatas(6))
                        
                    If Ret < 0 Then
                        errorfd (Ret)
                        .Active = False
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Kill sakifile
                        Exit Sub
                    End If
                        
                 FRMOMACHI.pb.Visible = True
                 FRMOMACHI.pb.Min = 0
                 FRMOMACHI.pb.Max = tdatal(1) + 1
                 FRMOMACHI.pb.Value = 0
                 FRMOMACHI.lblmsg.Caption = "データの読み込み中..."
                 DoEvents
                            ReDim ddatas(tdatal(1), 11) As String
                            ReDim ddatal(tdatal(1), 8) As Long
                      For i = 1 To tdatal(1)
                            Ret = .File_Read("D1D4X15D3X15X4D1D7X30D10D1D20D1X8x2", _
                            ddatal(i, 0), ddatal(i, 1), ddatas(i, 0), ddatal(i, 2), ddatas(i, 1), ddatas(i, 2), _
                            ddatal(i, 3), ddatal(i, 4), ddatas(i, 3), ddatal(i, 5), ddatal(i, 6), _
                            ddatal(i, 7), ddatal(i, 8), ddatas(i, 4), ddatas(i, 11))
                                
                            'エラー発生
                            If Ret < 0 Then
                                errorfd (Ret)
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                                FRMOMACHI.pb.Value = i
                                DoEvents
   
                            
                      Next i
                       
                        
                        
                        
                ElseIf nou = 0 Then
                        'トレーラコードまでのデータ更新
                            If .File_Seek(FileLen(sakifile) - 239) = False Then
                                MsgBox "読み込み位置の移動に失敗しました。"
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                        'レコードカウントを取得
                            Ret = .File_Read("D1D6D12D6D12D6D12X65", _
                            tdatal(0), tdatal(1), tdatal(2), tdatal(3), tdatal(4), tdatal(5), tdatal(6), tdatas(0))
                            'エラー発生
                            If Ret < 0 Then
                                errorfd (Ret)
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                .Active = False
                                Kill sakifile
                                Exit Sub
                            End If
                           
                            
                            
                            If .File_Seek(1) = False Then
                                MsgBox "読み込み位置の移動に失敗しました。"
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                            
                            If yuubinkakunin = 0 Then
                                    Ret = .File_Read("D1D2D1D1D9X40D4D4X15D3X15D1D7X17", _
                                        hdatal(0), hdatal(1), hdatal(2), itakushacode, hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                                        hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3))
                                        
                                         If Ret < 0 Then
                                                If .File_Seek(1) = False Then
                                                    MsgBox "読み込み位置の移動に失敗しました。"
                                                    .Active = False
                                                    Unload FRMOMACHI
                                                    Set FRMOMACHI = Nothing
                                                    Screen.MousePointer = 0
                                                    Kill sakifile
                                                    Exit Sub
                                                End If
                                                
                                                 Ret = .File_Read("D1D2D1D1D9X40D4D4X15D3X15D1D7X17", _
                                                    hdatal(0), hdatal(1), hdatal(2), itakushacode, hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                                                    hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3))
                                        
                                        End If
                            Else
                                    
                                    Ret = .File_Read("D1D2X1D2D8X40D4D4X15D3X17D6X4D1X12", _
                                        hdatal(0), hdatal(1), hdatas(4), hdatal(2), hdatal(3), _
                                        hdatas(0), hdatal(4), hdatal(5), _
                                        hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatas(5), hdatal(8), hdatas(3))
                            End If
                            
                            If Ret < 0 Then
                                errorfd (Ret)
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                 FRMOMACHI.pb.Visible = True
                 FRMOMACHI.pb.Min = 0
                 FRMOMACHI.pb.Max = tdatal(1)
                 FRMOMACHI.pb.Value = 0
                 FRMOMACHI.lblmsg.Caption = "データの読み込み中..."
            DoEvents
                    
                        
                            'ReDim ddatas(tdatal(1) - 1, 10) As String
                            'ReDim ddatal(tdatal(1) - 1, 8) As Long
                            'For I = 0 To tdatal(1) - 1
                        On Error GoTo erryomikomi
                            ReDim ddatas(tdatal(1), 10) As String
                            ReDim ddatal(tdatal(1), 8) As Long
                            For i = 1 To tdatal(1)
                                If i = tdatal(1) Then
                                    'MsgBox "Hit"
                                End If
                                If yuubinkakunin = 1 Then
                                    Ret = .File_Read("D1D4X15D3X15X2X2X1D7X30D10X1D20D1X4X2X2", _
                                    ddatal(i, 0), ddatal(i, 1), ddatas(i, 0), ddatal(i, 2), ddatas(i, 1), ddatas(i, 2), _
                                    ddatas(i, 5), ddatas(i, 6), _
                                    ddatal(i, 3), ddatas(i, 7), ddatal(i, 4), ddatas(i, 3), ddatal(i, 7), ddatal(i, 8), _
                                    ddatas(i, 9), ddatas(i, 8), ddatas(i, 4))
                                
                                Else
                                    Ret = .File_Read("D1D4X15D3X15X4D1D7X30D10D1D20D1X8", _
                                    ddatal(i, 0), ddatal(i, 1), ddatas(i, 0), ddatal(i, 2), ddatas(i, 1), ddatas(i, 2), _
                                    ddatal(i, 3), ddatal(i, 4), ddatas(i, 3), ddatal(i, 5), ddatal(i, 6), _
                                    ddatal(i, 7), ddatal(i, 8), ddatas(i, 4))
                                End If
                                'エラー発生
                                If Ret < 0 Then
                                    errorfd (Ret)
                                    .Active = False
                                    Unload FRMOMACHI
                                    Set FRMOMACHI = Nothing
                                    Screen.MousePointer = 0
                                    Kill sakifile
                                    Exit Sub
                                End If
                                FRMOMACHI.pb.Value = i
                                DoEvents
            
                            Next i
                          On Error GoTo 0
            Else
            
                        'トレーラコードまでのデータ更新
                            If .File_Seek(FileLen(sakifile) - 243) = False Then
                                MsgBox "読み込み位置の移動に失敗しました。"
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                        'レコードカウントを取得
                            Ret = .File_Read("D1D6D12D6D12D6D12X65X2", _
                            tdatal(0), tdatal(1), tdatal(2), tdatal(3), tdatal(4), tdatal(5), tdatal(6), tdatas(0), dami_crlf)
                            'エラー発生
                            If Ret < 0 Then
                                errorfd (Ret)
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                .Active = False
                                Kill sakifile
                                Exit Sub
                            End If
                            
                          
                            
                            If .File_Seek(1) = False Then
                                MsgBox "読み込み位置の移動に失敗しました。"
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                            
                            If yuubinkakunin = 0 Then
                                    Ret = .File_Read("D1D2D1D10X40D4D4X15D3X15D1D7X17X2", _
                                        hdatal(0), hdatal(1), hdatal(2), hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                                        hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3), dami_crlf)
                            Else
                                    
                                    Ret = .File_Read("D1D2X1D2D8X40D4D4X15D3X17D6X4D1X12", _
                                        hdatal(0), hdatal(1), hdatas(4), hdatal(2), hdatal(3), _
                                        hdatas(0), hdatal(4), hdatal(5), _
                                        hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatas(5), hdatal(8), hdatas(3))
                            End If
                            
                            If Ret < 0 Then
                                errorfd (Ret)
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                Kill sakifile
                                Exit Sub
                            End If
                 FRMOMACHI.pb.Visible = True
                 FRMOMACHI.pb.Min = 0
                 FRMOMACHI.pb.Max = tdatal(1)
                 FRMOMACHI.pb.Value = 0
                 FRMOMACHI.lblmsg.Caption = "データの読み込み中..."
            DoEvents
                    
                        
                            'ReDim ddatas(tdatal(1) - 1, 10) As String
                            'ReDim ddatal(tdatal(1) - 1, 8) As Long
                            'For I = 0 To tdatal(1) - 1
                            
                        On Error GoTo erryomikomi
                            ReDim ddatas(tdatal(1), 10) As String
                            ReDim ddatal(tdatal(1), 8) As Long
                            For i = 1 To tdatal(1)
                                If yuubinkakunin = 1 Then
                                    Ret = .File_Read("D1D4X15D3X15X2X2X1D7X30D10X1D20D1X4X2X2", _
                                    ddatal(i, 0), ddatal(i, 1), ddatas(i, 0), ddatal(i, 2), ddatas(i, 1), ddatas(i, 2), _
                                    ddatas(i, 5), ddatas(i, 6), _
                                    ddatal(i, 3), ddatas(i, 7), ddatal(i, 4), ddatas(i, 3), ddatal(i, 7), ddatal(i, 8), _
                                    ddatas(i, 9), ddatas(i, 8), ddatas(i, 4))
                                
                                Else
                                    Ret = .File_Read("D1D4X15D3X15X4D1D7X30D10D1D20D1X8X2", _
                                    ddatal(i, 0), ddatal(i, 1), ddatas(i, 0), ddatal(i, 2), ddatas(i, 1), ddatas(i, 2), _
                                    ddatal(i, 3), ddatal(i, 4), ddatas(i, 3), ddatal(i, 5), ddatal(i, 6), _
                                    ddatal(i, 7), ddatal(i, 8), ddatas(i, 4), dami_crlf)
                                End If
                                'エラー発生
                                If Ret < 0 Then
                                    errorfd (Ret)
                                    .Active = False
                                    Unload FRMOMACHI
                                    Set FRMOMACHI = Nothing
                                    Screen.MousePointer = 0
                                    Kill sakifile
                                    Exit Sub
                                End If
                                FRMOMACHI.pb.Value = i
                                DoEvents
            
                            Next i
                          On Error GoTo 0
            End If
            
            
            'ファイルを閉じる
                .Active = False
            'ファイルを削除
End With


            Unload FRMOMACHI
            Set FRMOMACHI = Nothing
            Screen.MousePointer = 0
             Kill sakifile

            MsgBox "全データの読み込みチェックが完了しました。"
    
    
    
End If






Exit Sub
            
ErrorTrap3:
Screen.MousePointer = 0
    MsgBox "VoIPデータの読み込みに失敗しました。"
    Exit Sub
errvoip:
Screen.MousePointer = 0
    MsgBox "VoIPデータの読み込みデータの書き込みに失敗しました。"
    Rollback
    db_voip.Close
    Exit Sub
errm1:
    '*******ロールバック*************
    Rollback
    Kill newfilemei
    MsgBox "銀行の移動記録の更新に失敗しました。"
    bkrs.Close
    db2.Close
    If rocked("off") = False Then
        MsgBox "個人データを更新中のフラッグの初期化に失敗しました。"
    End If
    Exit Sub

erronajifile:
        MsgBox "このデータを使用して契約データがすでに更新されています。"
        msg = "実際にデータが更新されていない場合（予期せぬエラーが発生した場合）があります。更新されていないですか？"
        res = MsgBox(msg, vbYesNo)
        If res = vbYes Then
            msg = "間違ってこの処理をすると、２回更新することになります。本当に更新されていませんか？"
            res = MsgBox(msg, vbYesNo)
            If res = vbYes Then
                Kill newfilemei
                MsgBox "間違った更新記録を修復しました。再度、更新処理を行ってください。"
             End If
        End If
        Unload FRMOMACHI
        Set FRMOMACHI = Nothing
        Screen.MousePointer = 0
        Exit Sub
errfdnashi:
    MsgBox "金融機関用更新ファイルが見つかりません。確認してから実行してください。"
    Exit Sub
    
errfdnashi2:
    Unload FRMOMACHI
    Set FRMOMACHI = Nothing
    Screen.MousePointer = 0
    MsgBox "フロッピーが入っていません。もう一度入れ直してから実行してください。"
    Exit Sub

erryomikomi:
     MsgBox "データを読み込んでいる際に予期せぬエラーが発生しました。"
    fd2.Active = False
    Unload FRMOMACHI
    Set FRMOMACHI = Nothing
    Screen.MousePointer = 0
    Kill sakifile
    Exit Sub
errmodori:
        MsgBox "正確な数字を入力してください。"
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Kill sakifile
                        Exit Sub




End Sub

Private Sub Picture13_Click()
Unload Me
End Sub

Private Sub Picture21_Click()
dirdou = 6
frmdir.Show 1
End Sub
