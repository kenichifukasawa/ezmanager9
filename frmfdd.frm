VERSION 5.00
Object = "{271E16FB-2B6D-11D4-B06D-0090991097BF}#26.0#0"; "ZoonFIO.ocx"
Begin VB.Form frmfdd 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "FDD出力"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   7680
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      Caption         =   "出力先"
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
      Height          =   1275
      Left            =   180
      TabIndex        =   7
      Top             =   1380
      Width           =   7335
      Begin VB.PictureBox Picture21 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   6390
         MouseIcon       =   "frmfdd.frx":0000
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmfdd.frx":08FF
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         Top             =   300
         Width           =   480
      End
      Begin VB.TextBox txtde 
         Enabled         =   0   'False
         Height          =   345
         IMEMode         =   4  '全角ひらがな
         Left            =   180
         TabIndex        =   8
         Top             =   600
         Width           =   5835
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
         Left            =   6390
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.Shape Shape15 
         BackStyle       =   1  '不透明
         BorderStyle     =   0  '透明
         Height          =   885
         Left            =   6210
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5040
      MouseIcon       =   "frmfdd.frx":1741
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmfdd.frx":2040
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
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
      Left            =   6570
      MouseIcon       =   "frmfdd.frx":2E82
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmfdd.frx":3781
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   330
      Width           =   480
   End
   Begin VB.Frame f1 
      Caption         =   "金融機関選択"
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
      Height          =   1155
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4125
      Begin VB.CheckBox chktest 
         Caption         =   "テスト出力をして実行する"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   6
         Top             =   840
         Width           =   2775
      End
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
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
   End
   Begin ZoonFIO.ZoonFileIO fd1 
      Left            =   0
      Top             =   600
      _ExtentX        =   900
      _ExtentY        =   900
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
      Left            =   5070
      TabIndex        =   5
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
      Left            =   6600
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   1035
      Left            =   4470
      Top             =   180
      Width           =   3030
   End
End
Attribute VB_Name = "frmfdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private jars As Recordset, jarecordaru As Integer
Private toresuu As Long, torekin As Long
Sub fdd_sai_check(s_bankno As String, bi2 As String)

Dim motofile As String, sakifile As String, ret As Integer, kensars As Recordset
Dim hdatas(6) As String, hdatal(8) As Long
Dim ddatas() As String, ddatal() As Long
Dim tdatas(1) As String, tdatal(6) As Long, kongetsu As String
Dim koushinid As String, koushinnewhi As String, newfilemei As String
Dim yuubinkakunin As Integer
    Dim nou As Integer
Dim motofile2 As String
Dim seikyuudata(), shousaidata(), tsuwadata(), FilePath(3)



'ここから



        motofile2 = Trim(txtde.Text)
       If Right(motofile2, 1) <> "\" Then
           motofile2 = motofile2 & "\"
       End If

     
        motofile2 = motofile2 & bi2 & "\"
        
        nou = 0
         If s_bankno = "5169" Or s_bankno = "1386" Then
          '  flopy = "a:\koufuri"
             motofile = motofile2 & "koufuri"
             nou = 1
        ElseIf s_bankno = "9900" Then
            'flopy = "a:\JIHARAI"
             motofile = motofile2 & "JIHARAI"
            nou = 2
        Else
            'flopy = "a:\koufuri.dat"
            motofile = motofile2 & "koufuri.dat"
        End If
        
'ここまで


    With frmfdd.fd1
        .File_Name = motofile
        .File_Type = 0
        .Active = True
        Select Case .Err
        Case -9
      Case -9
            MsgBox "データファイルがありません。フォルダーを確認してください。"
                Exit Sub
        Case -11
            MsgBox "データファイルがありません。手動でチェックを実行してください。"
            Exit Sub
        End Select
        'ヘッダーレコードの読み込み
        If nou = 0 Then
            ret = .File_Read("D1D2D1D10X40D4D4X15D3X15D1D7X17", _
                hdatal(0), hdatal(1), hdatal(2), hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3))
        
            '読み込みエラーが発生したら、郵便局を参照する。
             'エラー発生
            If ret < 0 Then
                If .File_Seek(1) = False Then
                    MsgBox "読み込み位置の移動に失敗しました。"
                    .Active = False
                    Kill sakifile
                    Exit Sub
                End If
    
                ret = .File_Read("D1D2X1D2D8X40D4D4X15D3X17D6X4D1X12", _
                    hdatal(0), hdatal(1), hdatas(4), hdatal(2), hdatal(3), _
                    hdatas(0), hdatal(4), hdatal(5), _
                    hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatas(5), hdatal(8), hdatas(3))
                
                yuubinkakunin = 1
            End If
        ElseIf nou = 1 Then
            ret = .File_Read("D1D2D1D10X40D4D4X15D3X15D1D7X17x2", _
                hdatal(0), hdatal(1), hdatal(2), hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3), hdatas(6))
        Else
            ret = .File_Read("D1D2D1D10X40D4D4X15D3X15D1D7X17", _
                hdatal(0), hdatal(1), hdatal(2), hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3))
        End If
    
        'エラー発生
        If ret < 0 Then
            errorfd (ret)
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
                    ret = .File_Read("D1D6D12D6D12D6D12X65", _
                    tdatal(0), tdatal(1), tdatal(2), tdatal(3), tdatal(4), tdatal(5), tdatal(6), tdatas(0))
                    'エラー発生
                    If ret < 0 Then
                        errorfd (ret)
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
                    ret = .File_Read("D1D6D12D6D12D6D12X65x2", _
                    tdatal(0), tdatal(1), tdatal(2), tdatal(3), tdatal(4), tdatal(5), tdatal(6), tdatas(0), tdatas(1))
                    'エラー発生
                    If ret < 0 Then
                        errorfd (ret)
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
                    ret = .File_Read("D1D6D12D6D12D6D12X65", _
                    tdatal(0), tdatal(1), tdatal(2), tdatal(3), tdatal(4), tdatal(5), tdatal(6), tdatas(0))
                    'エラー発生
                    If ret < 0 Then
                        errorfd (ret)
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



End Sub


Function koufushoukou(bc As String, hk As String) As Integer

Dim fdrs As Recordset, pbv As Long
Screen.MousePointer = 11

'ｓｑｌ文
        sql = "SELECT kojin.*,bank.bankcode,bank.shitencode,bank.bankmei" & _
                " FROM kojin LEFT JOIN bank ON kojin.bankid = bank.bankid" & _
                " WHERE bank.bankcode='" & bc & "'and clng(kojin.soukei) > 0"
'recordset
     If FcSQlGet(fdrs, sql, PrMsg) = False Then
        koufushoukou = 1
        Screen.MousePointer = 0
        Exit Function
     End If
     
'書き込み
 FRMOMACHI.pb.Visible = True
 FRMOMACHI.pb.Min = 0
 FRMOMACHI.pb.Max = fdrs.RecordCount
 FRMOMACHI.pb.Value = 0
 FRMOMACHI.lblmsg.Caption = "ファイルを作成中..."
DoEvents

Dim itakushakodo As Long, itakushamei As String, hikiotoshibi As Integer, headdata(2) As Integer
Dim bankcode As Long, bankmei As String, shitencode As Long, shitenmei As String
Dim shubetsu As Integer, kouzabangou As Long, hdami As String, khi As String

Dim dshitencode As Long, dshitenmei As String, ddami As String, dshubetsu As Integer
Dim dkouzabangou As Long, yokinshamei As String, hikiotoshikin As Long
Dim kokyakuid As Long, ddami2 As String, ddata(2) As Integer

Dim goukeisuu As Long, goukeikingaku As Long, tdami As String, tdata(1) As Integer

Dim edami As String, edata(0) As Integer

'銀行別ヘッダーレコード
Dim ret As Integer
'甲府商工
toresuu = 0
torekin = 0
headdata(0) = 1
headdata(1) = 91
headdata(2) = 1
    
If soukatsu(0, "koufushoukou", 1, "", 1, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        koufushoukou = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    koufushoukou = 3
    Screen.MousePointer = 0
    Exit Function
End If

itakushakodo = CLng(reg)        '委託者コード
If soukatsu(0, "koufushoukou", 2, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        koufushoukou = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    koufushoukou = 3
    Screen.MousePointer = 0
    Exit Function
End If
itakushamei = CStr(reg)   '委託者名

hikiotoshibi = CInt(Mid(hk, 5))   '引落日

bankcode = 1386
bankmei = "ｺｳﾌｼﾖｳｺｳｼﾝｷﾝ"
If soukatsu(0, "koufushoukou", 4, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        koufushoukou = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    koufushoukou = 3
    Screen.MousePointer = 0
    Exit Function
End If
shitencode = CLng(reg) '支店コード
If soukatsu(0, "koufushoukou", 5, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        koufushoukou = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    koufushoukou = 3
    Screen.MousePointer = 0
    Exit Function
End If
shitenmei = CStr(reg) '支店名
If soukatsu(0, "koufushoukou", 6, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        koufushoukou = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    koufushoukou = 3
    Screen.MousePointer = 0
    Exit Function
End If
shubetsu = CInt(reg)  '種別
If soukatsu(0, "koufushoukou", 7, "", 0, 1) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        koufushoukou = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    Screen.MousePointer = 0
    koufushoukou = 3
    Exit Function
End If
kouzabangou = CLng(reg) '口座番号

hdami = Space(17)

ret = fd1.File_Write("D1D2D1DRZ10XLS40DRZ4DRZ4XLS15DRZ3XLS15D1DRZ7X17", _
    headdata(0), headdata(1), headdata(2), itakushakodo, itakushamei, hikiotoshibi, bankcode, bankmei, _
    shitencode, shitenmei, shubetsu, kouzabangou, hdami)


'エラー発生
If ret < 0 Then
    errorfd (ret)
    MsgBox "Head data(" & headdata(0) & "," & headdata(1) & "," & headdata(2) & "," & itakushakodo & _
            itakushamei & "," & hikiotoshibi & "," & bankcode & "," & bankmei & _
            shitencode & "," & shitenmei & "," & shubetsu & "," & kouzabangou & "," & hdami & ")"
    koufushoukou = 2
    Screen.MousePointer = 0
    Exit Function
End If
    pbv = 0
    
    fdrs.MoveFirst
    Do Until fdrs.EOF
        ddata(0) = 2
        If Trim(fdrs!shitencode) = "" Then
            koufushoukou = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            dshitencode = CLng(fdrs!shitencode)
        End If
        dshitenmei = Space(15)
        ddami = Space(4)
        If fdrs!kouzakubun = 0 Then '普通
            dshubetsu = 1
        ElseIf fdrs!kouzakubun = 1 Then '当座
            dshubetsu = 2
        Else
            koufushoukou = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        End If
        If Trim(fdrs!kouzabangou) = "" Then
            koufushoukou = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            dkouzabangou = CLng(fdrs!kouzabangou)
        End If
        If Trim(fdrs!meigi) = "" Then
            koufushoukou = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            yokinshamei = CStr(fdrs!meigi)
        End If
        If Trim(fdrs!soukei) = "" Then
            koufushoukou = 4
            MsgBox fdrs!kojinid & Space(3) & yokinshamei
            Screen.MousePointer = 0
            Exit Function
        Else
            hikiotoshikin = CLng(fdrs!soukei)
             toresuu = toresuu + 1
             torekin = torekin + CLng(fdrs!soukei)

        End If
        ddata(1) = 1
        If Trim(fdrs!kojinid) = "" Then
            koufushoukou = 4
            MsgBox fdrs!kojinid & Space(3) & yokinshamei
            Screen.MousePointer = 0
            Exit Function
        Else
            kokyakuid = CLng(fdrs!kojinid)
        End If
        ddata(2) = 0
        ddami2 = Space(8)
        'データ・レコード
        ret = fd1.File_Write("D1DRZ4XLS15DRZ3X15X4D1DRZ7XLS30DRZ10D1DRZ20D1X8", _
            ddata(0), bankcode, bankmei, dshitencode, dshitenmei, ddami, dshubetsu, _
            dkouzabangou, yokinshamei, hikiotoshikin, ddata(1), kokyakuid, ddata(2), ddami2)

        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            MsgBox "Head data(" & ddata(0) & "," & bankcode & "," & bankmei & "," & dshitencode & _
                    dshitenmei & "," & ddami & "," & dshubetsu & _
                    dkouzabangou & "," & yokinshamei & "," & hikiotoshikin & "," & ddata(1) & _
                    kokyakuid & "," & ddata(2) & "," & ddami2 & ")"
            koufushoukou = 2
            Screen.MousePointer = 0
            Exit Function
        End If
        fdrs.MoveNext
        pbv = pbv + 1
        FRMOMACHI.pb.Value = pbv
        DoEvents
    Loop
    fdrs.Close
    FRMOMACHI.pb.Visible = False
    


'トレーラ・レコード
tdata(0) = 8
tdata(1) = 0
goukeisuu = toresuu
goukeikingaku = torekin
tdami = Space(65)
ret = fd1.File_Write("D1DRZ6DRZ12DRZ6DRZ12DRZ6DRZ12X65", _
    tdata(0), goukeisuu, goukeikingaku, tdata(1), tdata(1), tdata(1), tdata(1), tdami)
    
        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            MsgBox "Head data(" & tdata(0) & "," & goukeisuu & "," & goukeikingaku & "," & tdata(1) & _
                    tdata(1) & "," & tdata(1) & "," & tdata(1) & "," & tdami & ")"
            koufushoukou = 2
            Screen.MousePointer = 0
            Exit Function
        End If
    
'エンド・レコード

edami = Space(119)
edata(0) = 9
ret = fd1.File_Write("D1X119", edata(0), edami)
     
        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            koufushoukou = 2
            Screen.MousePointer = 0
            Exit Function
        End If
     
Screen.MousePointer = 0
End Function



Function yamanami(bc As String, hk As String) As Integer

Dim fdrs As Recordset, pbv As Long
Screen.MousePointer = 11

'ｓｑｌ文
        sql = "SELECT kojin.*,bank.bankcode,bank.shitencode,bank.bankmei" & _
                " FROM kojin LEFT JOIN bank ON kojin.bankid = bank.bankid" & _
                " WHERE bank.bankcode='" & bc & "'and clng(kojin.soukei) > 0"
'recordset
     If FcSQlGet(fdrs, sql, PrMsg) = False Then
        yamanami = 1
        Screen.MousePointer = 0
        Exit Function
     End If
     
'書き込み
 FRMOMACHI.pb.Visible = True
 FRMOMACHI.pb.Min = 0
 FRMOMACHI.pb.Max = fdrs.RecordCount
 FRMOMACHI.pb.Value = 0
 FRMOMACHI.lblmsg.Caption = "ファイルを作成中..."
DoEvents

Dim itakushakodo2 As String
Dim itakushakodo As Long, itakushamei As String, hikiotoshibi As Integer, headdata(2) As Integer
Dim bankcode As Long, bankmei As String, shitencode As Long, shitenmei As String
Dim shubetsu As Integer, kouzabangou As Long, hdami As String, khi As String

Dim dshitencode As Long, dshitenmei As String, ddami As String, dshubetsu As Integer
Dim dkouzabangou As Long, yokinshamei As String, hikiotoshikin As Long
Dim kokyakuid As Long, ddami2 As String, ddata(2) As Integer
Dim kokyakuid2 As String
Dim goukeisuu As Long, goukeikingaku As Long, tdami As String, tdata(1) As Integer

Dim edami As String, edata(0) As Integer

Dim s_itakusha_head As String

'銀行別ヘッダーレコード
Dim ret As Integer
'やまなみ
toresuu = 0
torekin = 0

headdata(0) = 1
headdata(1) = 91
'headdata(2) = 1
headdata(2) = 0 '1  2018/09/11 修正

If soukatsu(0, "yamanami", 1, "", 1, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yamanami = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yamanami = 3
    Screen.MousePointer = 0
    Exit Function
Else
    s_itakusha_head = Mid(reg, 1, 1)
    Select Case s_itakusha_head
    Case "0", "1"
        itakushakodo = CLng(reg)        '委託者コード
    Case Else
        itakushakodo2 = Trim(reg)        '委託者コード
    End Select
End If



If soukatsu(0, "yamanami", 2, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yamanami = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yamanami = 3
    Screen.MousePointer = 0
    Exit Function
End If
itakushamei = CStr(reg)   '委託者名


hikiotoshibi = CInt(Mid(hk, 5))   '引落日

bankcode = 2377
bankmei = Space(15)

If soukatsu(0, "yamanami", 4, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yamanami = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yamanami = 3
    Screen.MousePointer = 0
    Exit Function
End If
shitencode = CLng(reg) '支店コード

If soukatsu(0, "yamanami", 5, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yamanami = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yamanami = 3
    Screen.MousePointer = 0
    Exit Function
End If
shitenmei = CStr(reg) '支店名

If soukatsu(0, "yamanami", 6, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yamanami = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yamanami = 3
    Screen.MousePointer = 0
    Exit Function
End If
shubetsu = CInt(reg)  '種別
    
If soukatsu(0, "yamanami", 7, "", 0, 1) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yamanami = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    Screen.MousePointer = 0
    yamanami = 3
    Exit Function
End If
kouzabangou = CLng(reg) '口座番号

hdami = Space(17)


    Select Case s_itakusha_head
    Case "0", "1"
        ret = fd1.File_Write("D1D2D1DRZ10XLS40DRZ4DRZ4XLS15DRZ3XLS15D1DRZ7X17", _
        headdata(0), headdata(1), headdata(2), itakushakodo, itakushamei, hikiotoshibi, bankcode, bankmei, _
        shitencode, shitenmei, shubetsu, kouzabangou, hdami)

            'エラー発生
        If ret < 0 Then
            errorfd (ret)
            MsgBox "Head data(" & headdata(0) & "," & headdata(1) & "," & headdata(2) & "," & itakushakodo & _
                    itakushamei & "," & hikiotoshibi & "," & bankcode & "," & bankmei & _
                    shitencode & "," & shitenmei & "," & shubetsu & "," & kouzabangou & "," & hdami & ")"
            yamanami = 2
            Screen.MousePointer = 0
            Exit Function
        End If
    Case Else
           '委託者コード修正
        ret = fd1.File_Write("D1D2D1X10XLS40DRZ4DRZ4XLS15DRZ3XLS15D1DRZ7X17", _
        headdata(0), headdata(1), headdata(2), itakushakodo2, itakushamei, hikiotoshibi, bankcode, bankmei, _
        shitencode, shitenmei, shubetsu, kouzabangou, hdami)

            'エラー発生
        If ret < 0 Then
            errorfd (ret)
            MsgBox "Head data(" & headdata(0) & "," & headdata(1) & "," & headdata(2) & "," & itakushakodo2 & _
                    itakushamei & "," & hikiotoshibi & "," & bankcode & "," & bankmei & _
                    shitencode & "," & shitenmei & "," & shubetsu & "," & kouzabangou & "," & hdami & ")"
            yamanami = 2
            Screen.MousePointer = 0
            Exit Function
        End If
           
           
           
           
           
    End Select




    pbv = 0
    
    fdrs.MoveFirst
    Do Until fdrs.EOF
        ddata(0) = 2
        If Trim(fdrs!shitencode) = "" Then
            yamanami = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            dshitencode = CLng(fdrs!shitencode)
        End If
        dshitenmei = Space(15)
        ddami = Space(4)
        If fdrs!kouzakubun = 0 Then '普通
            dshubetsu = 1
        ElseIf fdrs!kouzakubun = 1 Then '当座
            dshubetsu = 2
        Else
            yamanami = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        End If
        If Trim(fdrs!kouzabangou) = "" Then
            yamanami = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            dkouzabangou = CLng(fdrs!kouzabangou)
        End If
        If Trim(fdrs!meigi) = "" Then
            yamanami = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            yokinshamei = CStr(fdrs!meigi)
        End If
        If Trim(fdrs!soukei) = "" Then
            yamanami = 4
            MsgBox fdrs!kojinid & Space(3) & yokinshamei
            Screen.MousePointer = 0
            Exit Function
        Else
            hikiotoshikin = CLng(fdrs!soukei)
             toresuu = toresuu + 1
             torekin = torekin + CLng(fdrs!soukei)
        End If
        ddata(1) = 1
        If Trim(fdrs!kojinid) = "" Then
            yamanami = 4
            MsgBox fdrs!kojinid & Space(3) & yokinshamei
            Screen.MousePointer = 0
            Exit Function
        Else
            kokyakuid = CLng(fdrs!kojinid)
            kokyakuid2 = Trim(fdrs!kojinid)
        End If
        ddata(2) = 0
        ddami2 = Space(8)
        'データ・レコード
        ret = fd1.File_Write("D1DRZ4XLS15DRZ3X15X4D1DRZ7XLS30DRZ10D1DRZ20D1X8", _
            ddata(0), bankcode, bankmei, dshitencode, dshitenmei, ddami, dshubetsu, _
            dkouzabangou, yokinshamei, hikiotoshikin, ddata(1), kokyakuid, ddata(2), ddami2)

        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            MsgBox "Head data(" & ddata(0) & "," & bankcode & "," & bankmei & "," & dshitencode & _
                    dshitenmei & "," & ddami & "," & dshubetsu & _
                    dkouzabangou & "," & yokinshamei & "," & hikiotoshikin & "," & ddata(1) & _
                    kokyakuid & "," & ddata(2) & "," & ddami2 & ")"
            yamanami = 2
            Screen.MousePointer = 0
            Exit Function
        End If
        
        If shousaiset(kokyakuid2, hk, hikiotoshikin, , bc) = -1 Then
            MsgBox "詳細データベースへの書き込みに失敗しました。" & ddata(0) & "," & bankcode & "," & bankmei & "," & dshitencode & _
                    dshitenmei & "," & ddami & "," & dshubetsu & _
                    dkouzabangou & "," & yokinshamei & "," & hikiotoshikin & "," & ddata(1) & _
                    kokyakuid & "," & ddata(2) & "," & ddami2 & ")"
            yamanami = 2
            Screen.MousePointer = 0
            Exit Function
        End If
        
        fdrs.MoveNext
        pbv = pbv + 1
        FRMOMACHI.pb.Value = pbv
        DoEvents
    Loop
    fdrs.Close
    FRMOMACHI.pb.Visible = False
    


'トレーラ・レコード
tdata(0) = 8
tdata(1) = 0
goukeisuu = toresuu
goukeikingaku = torekin
tdami = Space(65)
ret = fd1.File_Write("D1DRZ6DRZ12DRZ6DRZ12DRZ6DRZ12X65", _
    tdata(0), goukeisuu, goukeikingaku, tdata(1), tdata(1), tdata(1), tdata(1), tdami)
'ret = fd1.File_Write("D1DRZ6DRZ12DRZ6DRZ12DRZ6DRZ12X65", _
    tdata(0), goukeisuu, goukeikingaku, goukeisuu, goukeikingaku, tdata(1), tdata(1), tdami)
    
        'エラー発生
        If ret < 0 Then
            errorfd (ret)
             MsgBox "Head data(" & tdata(0) & "," & goukeisuu & "," & goukeikingaku & "," & tdata(1) & _
             tdata(1) & "," & tdata(1) & "," & tdata(1) & "," & tdami & ")"
           yamanami = 2
            Screen.MousePointer = 0
            Exit Function
        End If
    
'エンド・レコード

edami = Space(119)
edata(0) = 9
ret = fd1.File_Write("D1X119", edata(0), edami)
     
        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            yamanami = 2
            Screen.MousePointer = 0
            Exit Function
        End If
     
     

     
     
     
Screen.MousePointer = 0
End Function


Function koufushinkin(bc As String, hk As String) As Integer

Dim fdrs As Recordset, pbv As Long
Screen.MousePointer = 11

'ｓｑｌ文
        sql = "SELECT kojin.*,bank.bankcode,bank.shitencode,bank.bankmei" & _
                " FROM kojin LEFT JOIN bank ON kojin.bankid = bank.bankid" & _
                " WHERE bank.bankcode='" & bc & "'and clng(kojin.soukei) > 0"
'recordset
     If FcSQlGet(fdrs, sql, PrMsg) = False Then
        koufushinkin = 1
        Screen.MousePointer = 0
        Exit Function
     End If
     
'書き込み
 FRMOMACHI.pb.Visible = True
 FRMOMACHI.pb.Min = 0
 FRMOMACHI.pb.Max = fdrs.RecordCount
 FRMOMACHI.pb.Value = 0
 FRMOMACHI.lblmsg.Caption = "ファイルを作成中..."
DoEvents

Dim itakushakodo As Long, itakushamei As String, hikiotoshibi As Integer, headdata(2) As Integer
Dim bankcode As Long, bankmei As String, shitencode As Long, shitenmei As String
Dim shubetsu As Integer, kouzabangou As Long, hdami As String, khi As String

Dim dshitencode As Long, dshitenmei As String, ddami As String, dshubetsu As Integer
Dim dkouzabangou As Long, yokinshamei As String, hikiotoshikin As Long
Dim kokyakuid As Long, ddami2 As String, ddata(2) As Integer

Dim goukeisuu As Long, goukeikingaku As Long, tdami As String, tdata(1) As Integer

Dim edami As String, edata(0) As Integer

'銀行別ヘッダーレコード
Dim ret As Integer
'甲府商工
toresuu = 0
torekin = 0

headdata(0) = 1
headdata(1) = 91
headdata(2) = 1

If soukatsu(0, "koufushinkin", 1, "", 1, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        koufushinkin = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    koufushinkin = 3
    Screen.MousePointer = 0
    Exit Function
End If

itakushakodo = CLng(reg)        '委託者コード

If soukatsu(0, "koufushinkin", 2, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        koufushinkin = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    koufushinkin = 3
    Screen.MousePointer = 0
    Exit Function
End If
itakushamei = CStr(reg)   '委託者名


hikiotoshibi = CInt(Mid(hk, 5))   '引落日

bankcode = 1385
bankmei = Space(15)

If soukatsu(0, "koufushinkin", 4, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        koufushinkin = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    koufushinkin = 3
    Screen.MousePointer = 0
    Exit Function
End If
shitencode = CLng(reg) '支店コード

If soukatsu(0, "koufushinkin", 5, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        koufushinkin = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    koufushinkin = 3
    Screen.MousePointer = 0
    Exit Function
End If
shitenmei = CStr(reg) '支店名

If soukatsu(0, "koufushinkin", 6, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        koufushinkin = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    koufushinkin = 3
    Screen.MousePointer = 0
    Exit Function
End If
shubetsu = CInt(reg)  '種別
    
If soukatsu(0, "koufushinkin", 7, "", 0, 1) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        koufushinkin = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    Screen.MousePointer = 0
    koufushinkin = 3
    Exit Function
End If
kouzabangou = CLng(reg) '口座番号

hdami = Space(17)

ret = fd1.File_Write("D1D2D1DRZ10XLS40DRZ4DRZ4XLS15DRZ3XLS15D1DRZ7X17", _
    headdata(0), headdata(1), headdata(2), itakushakodo, itakushamei, hikiotoshibi, bankcode, bankmei, _
    shitencode, shitenmei, shubetsu, kouzabangou, hdami)


'エラー発生
If ret < 0 Then
    errorfd (ret)
    MsgBox "Head data(" & headdata(0) & "," & headdata(1) & "," & headdata(2) & "," & itakushakodo & _
            itakushamei & "," & hikiotoshibi & "," & bankcode, bankmei & _
            shitencode & "," & shitenmei & "," & shubetsu & "," & kouzabangou & "," & hdami & ")"
    koufushinkin = 2
    Screen.MousePointer = 0
    Exit Function
End If
    pbv = 0
    
    fdrs.MoveFirst
    Do Until fdrs.EOF
        ddata(0) = 2
        If Trim(fdrs!shitencode) = "" Then
            koufushinkin = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            dshitencode = CLng(fdrs!shitencode)
        End If
        dshitenmei = Space(15)
        ddami = Space(4)
        If fdrs!kouzakubun = 0 Then '普通
            dshubetsu = 1
        ElseIf fdrs!kouzakubun = 1 Then '当座
            dshubetsu = 2
        Else
            koufushinkin = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        End If
        If Trim(fdrs!kouzabangou) = "" Then
            koufushinkin = 4
            Screen.MousePointer = 0
            Exit Function
        Else
            dkouzabangou = CLng(fdrs!kouzabangou)
        End If
        If Trim(fdrs!meigi) = "" Then
            koufushinkin = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            yokinshamei = CStr(fdrs!meigi)
        End If
        If Trim(fdrs!soukei) = "" Then
            koufushinkin = 4
            MsgBox fdrs!kojinid & Space(3) & yokinshamei
            Screen.MousePointer = 0
            Exit Function
        Else
            hikiotoshikin = CLng(fdrs!soukei)
             toresuu = toresuu + 1
             torekin = torekin + CLng(fdrs!soukei)
        
        End If
        ddata(1) = 1
        If Trim(fdrs!kojinid) = "" Then
            koufushinkin = 4
            MsgBox fdrs!kojinid & Space(3) & yokinshamei
            Screen.MousePointer = 0
            Exit Function
        Else
            kokyakuid = CLng(fdrs!kojinid)
        End If
        ddata(2) = 0
        ddami2 = Space(8)
        'データ・レコード
        ret = fd1.File_Write("D1DRZ4XLS15DRZ3X15X4D1DRZ7XLS30DRZ10D1DRZ20D1X8", _
            ddata(0), bankcode, bankmei, dshitencode, dshitenmei, ddami, dshubetsu, _
            dkouzabangou, yokinshamei, hikiotoshikin, ddata(1), kokyakuid, ddata(2), ddami2)
    'エラー発生
        If ret < 0 Then
            errorfd (ret)
            MsgBox "Head data(" & ddata(0) & "," & bankcode & "," & bankmei & "," & dshitencode & _
                    dshitenmei & "," & ddami & "," & dshubetsu & "," & _
                    dkouzabangou & "," & yokinshamei & "," & hikiotoshikin & "," & ddata(1) & "," & _
                    kokyakuid & "," & ddata(2) & "," & ddami2 & ")"
            koufushinkin = 2
            Screen.MousePointer = 0
            Exit Function
        End If
        fdrs.MoveNext
        pbv = pbv + 1
        FRMOMACHI.pb.Value = pbv
        DoEvents
    Loop
    fdrs.Close
    FRMOMACHI.pb.Visible = False
    


'トレーラ・レコード
tdata(0) = 8
tdata(1) = 0

goukeisuu = toresuu
goukeikingaku = torekin
tdami = Space(65)
ret = fd1.File_Write("D1DRZ6DRZ12DRZ6DRZ12DRZ6DRZ12X65", _
    tdata(0), goukeisuu, goukeikingaku, tdata(1), tdata(1), tdata(1), tdata(1), tdami)
    
        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            MsgBox "Head data(" & tdata(0) & "," & goukeisuu & "," & goukeikingaku & "," & tdata(1) & _
                    tdata(1) & "," & tdata(1) & "," & tdata(1) & "," & tdami & ")"
            koufushinkin = 2
            Screen.MousePointer = 0
            Exit Function
        End If
    
'エンド・レコード

edami = Space(119)
edata(0) = 9
ret = fd1.File_Write("D1X119", edata(0), edami)
     
        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            koufushinkin = 2
            Screen.MousePointer = 0
            Exit Function
        End If
     
Screen.MousePointer = 0
End Function
Function shitei_shousai_sakujo(code As String, nengetsu As String) As Boolean
    
    Dim sql_saku As String, sql_saku2 As String
    
    If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        shitei_shousai_sakujo = False
        Exit Function
    End If
    
    sql_saku = "delete * from shousai where bankid='" & code & "' and keiseikyuu='" & nengetsu & "' and keist='0'"
    On Error GoTo errsakujo_9992
     db2.Execute sql_saku, dbSQLPassThrough
    On Error GoTo 0
    
    sql = "select kojin.kojinid from kojin LEFT JOIN bank ON kojin.bankid = bank.bankid" & _
                " where bank.bankcode='" & code & "'"
    'recordset
    If FcSQlGet(rs, sql, PrMsg) = True Then
        rs.MoveFirst
        Do Until rs.EOF
            If flgup(0, rs!kojinid) = False Then
                MsgBox "個人データのフラッグを所期できませんでした。初めから再度実行してください。"
                shitei_shousai_sakujo = False
                Exit Function
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If

shitei_shousai_sakujo = True
Exit Function
errsakujo_9992:
        MsgBox "データの初期化に失敗しました。"
        shitei_shousai_sakujo = False
        Exit Function

    
End Function



Function makefd(bid As String, test_shutsu As Integer, bi2 As String) As Integer


Dim fdrs2  As Recordset, fdfilename As String, flopy As String, kongetsu As Integer
Dim fdrs3 As Recordset

If test_shutsu = 1 Then
    MsgBox "出力テストを開始します。"
End If

'FDの確認
 frmfdd.Hide
 FRMOMACHI.lblmsg.Caption = "書き込みの準備中..."
 FRMOMACHI.Show
 DoEvents
 Screen.MousePointer = 11
 
 fdfilename = sopa & "fdback"
 If Dir$(fdfilename, vbDirectory) = "" Then
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder (fdfilename)
    Set fso = Nothing
End If
If Right(fdfilename, 1) <> "\" Then
    fdfilename = fdfilename & "\"
End If

'If Format(Date, "mm") <> "12" Then
    kongetsu = CInt(Format(Date, "mm"))   '+ 1
'Else
'    kongetsu = 1
'End If
fdfilename = fdfilename & Format(bid, "0000") & Format(Date, "yyyy") & Format(kongetsu, "00") & "s.dat"
    
    With fd1
        .File_Name = fdfilename    'ファイル名
        .File_Type = 1    'ファイル書き込み
        .Active = True    'ファイルオープン
        If .Err = -9 Then
            
        ElseIf .Err = -10 Then
            makefd = 11
            .Active = False
            If test_shutsu = 1 Then
                MsgBox "出力処理がすでに行われています。"
                makefd = 15
            Else
                msg = "出力処理がすでに行われています。再度出力しますか？"
                res = MsgBox(msg, vbYesNo)
                If res = vbYes Then
                    If shitei_shousai_sakujo(Format(bid, "0000"), Format(Date, "yyyy") & Format(kongetsu, "00")) = True Then
                        Kill fdfilename
                    End If
                    makefd = 15
                End If
            End If
            Unload FRMOMACHI
            Set FRMOMACHI = Nothing
            Screen.MousePointer = 0
            Exit Function
        ElseIf .Err = -11 Then
            If test_shutsu = 0 Then
                makefd = 12
                .Active = False
                    If shitei_shousai_sakujo(Format(bid, "0000"), Format(Date, "yyyy") & Format(kongetsu, "00")) = True Then
                        Kill fdfilename
                    End If
            End If
            Unload FRMOMACHI
            Set FRMOMACHI = Nothing
            Screen.MousePointer = 0
            Exit Function
        End If
    
    End With
Screen.MousePointer = 0
'ファイルの作成

Dim karitsuki As Integer

BeginTrans

Select Case bid
Case "0142"  '山梨中央銀行
    If soukatsu(0, "yamachuu", 3, "", 1, 1) = False Then
            MsgBox "設定の読み込みに失敗しました。"
            makefd = 4
            Screen.MousePointer = 0
            Exit Function
    End If
    If reg = "0" Then
        reg = ""
    End If
    If reg <> "" Then
        donohi = 5
        karitsuki = CInt(Format(Date, "mm"))
        frmhiduke.MonthView1.Value = Format(Date, "yy") & "/" & Format(karitsuki, "00") & "/" & Format(reg, "00")
        frmhiduke.lbl1.Visible = True
        frmhiduke.Show 1
        Select Case yamachuu(bid, fdhiduke)
        Case 0  '正常
            makefd = 1 '成功
        Case 1  'レコードなし
            makefd = 2
        Case 2  '異常
            makefd = 3
        Case 3  '契約内容登録されていない
            makefd = 4
        Case 4  'レコードセットの値が不正
            makefd = 5
        Case 5 '    集計レコードなし
            makefd = 6
        End Select
    Else
        makefd = 4
    End If
Case "1385"  '甲府信用金庫
    makefd = 13
    Exit Function

    If soukatsu(0, "koufushinkin", 3, "", 1, 1) = False Then
            MsgBox "設定の読み込みに失敗しました。"
            makefd = 4
            Screen.MousePointer = 0
            Exit Function
    End If
    If reg = "0" Then
        reg = ""
    End If
    If reg <> "" Then
        donohi = 5
        karitsuki = CInt(Format(Date, "mm"))
        frmhiduke.MonthView1.Value = Format(Date, "yy") & "/" & Format(karitsuki, "00") & "/" & Format(reg, "00")
        frmhiduke.lbl1.Visible = True
        frmhiduke.Show 1
        Select Case koufushinkin(bid, fdhiduke)
        Case 0  '正常
            makefd = 1 '成功
        Case 1  'レコードなし
            makefd = 2
        Case 2  '異常
            makefd = 3
        Case 3  '契約内容登録されていない
            makefd = 4
        Case 4  'レコードセットの値が不正
            makefd = 5
        Case 5 '    集計レコードなし
            makefd = 6
        End Select
    Else
        makefd = 4
    End If
        
Case "2377"  'やまなみ信用組合
    If soukatsu(0, "yamanami", 3, "", 1, 1) = False Then
            MsgBox "設定の読み込みに失敗しました。"
            makefd = 4
            Screen.MousePointer = 0
            Exit Function
    End If
    If reg = "0" Then
        reg = ""
    End If
    If reg <> "" Then
        donohi = 5
        karitsuki = CInt(Format(Date, "mm"))
        frmhiduke.MonthView1.Value = Format(Date, "yy") & "/" & Format(karitsuki, "00") & "/" & Format(reg, "00")
        frmhiduke.lbl1.Visible = True
        frmhiduke.Show 1
        Select Case yamanami(bid, fdhiduke)
        Case 0  '正常
            makefd = 1 '成功
        Case 1  'レコードなし
            makefd = 2
        Case 2  '異常
            makefd = 3
        Case 3  '契約内容登録されていない
            makefd = 4
        Case 4  'レコードセットの値が不正
            makefd = 5
        Case 5 '    集計レコードなし
            makefd = 6
        End Select
    Else
        makefd = 4
    End If

Case "1386"  '甲府商工信用金庫
    makefd = 13
    Exit Function
    If soukatsu(0, "koufushoukou", 3, "", 1, 1) = False Then
            MsgBox "設定の読み込みに失敗しました。"
            makefd = 4
            Screen.MousePointer = 0
            Exit Function
    End If
    If reg = "0" Then
        reg = ""
    End If
    If reg <> "" Then
        donohi = 5
        karitsuki = CInt(Format(Date, "mm"))
        frmhiduke.MonthView1.Value = Format(Date, "yy") & "/" & Format(karitsuki, "00") & "/" & Format(reg, "00")
        frmhiduke.lbl1.Visible = True
        frmhiduke.Show 1
        Select Case koufushoukou(bid, fdhiduke)
        Case 0  '正常
            makefd = 1 '成功
        Case 1  'レコードなし
            makefd = 2
        Case 2  '異常
            makefd = 3
        Case 3  '契約内容登録されていない
            makefd = 4
        Case 4  'レコードセットの値が不正
            makefd = 5
        Case 5 '    集計レコードなし
            makefd = 6
        End Select
    Else
        makefd = 4
    End If

Case "9900"  '郵便局
    If soukatsu(0, "yuubinkyoku", 7, "", 1, 1) = False Then
            MsgBox "設定の読み込みに失敗しました。"
            makefd = 4
            Screen.MousePointer = 0
            Exit Function
    End If
    If reg = "0" Then
        reg = ""
    End If
    If reg <> "" Then
        donohi = 5
        karitsuki = CInt(Format(Date, "mm"))
        frmhiduke.MonthView1.Value = Format(Date, "yy") & "/" & Format(karitsuki, "00") & "/" & Format(reg, "00")
        frmhiduke.lbl1.Visible = True
        frmhiduke.Show 1
        Select Case yuubinkyoku2(bid, fdhiduke)
        Case 0  '正常
            makefd = 1 '成功
        Case 1  'レコードなし
            makefd = 2
        Case 2  '異常
            makefd = 3
        Case 3  '契約内容登録されていない
            makefd = 4
        Case 4  'レコードセットの値が不正
            makefd = 5
        Case 5 '    集計レコードなし
            makefd = 6
        Case 6 '    口座番号不正
            makefd = 10
        End Select
    Else
        makefd = 4
    End If
Case "5169"  'JA
    If soukatsu(0, "jakihon", 3, "", 1, 1) = False Then
            MsgBox "設定の読み込みに失敗しました。"
            makefd = 4
            Screen.MousePointer = 0
            Exit Function
    End If
    If reg = "0" Then
        reg = ""
    End If
    If reg <> "" Then
    
        donohi = 5
        karitsuki = CInt(Format(Date, "mm"))
        frmhiduke.MonthView1.Value = Format(Date, "yy") & "/" & Format(karitsuki, "00") & "/" & Format(reg, "00")
        frmhiduke.lbl1.Visible = True
        frmhiduke.Show 1
        YAJI = Chr(13) + Chr(10)
        Select Case jayamanashi(bid, fdhiduke)
        Case 0  '正常
            makefd = 1 '成功
        Case 1  'レコードなし
            makefd = 2
        Case 2  '異常
            makefd = 3
        Case 3  '契約内容登録されていない
            makefd = 4
        Case 4  'レコードセットの値が不正
            makefd = 5
        Case 5 '    集計レコードなし
            makefd = 6
        End Select
    Else
        makefd = 4
    End If

Case Else
    makefd = 13
End Select
'ファイルを閉じる
fd1.Active = False
If test_shutsu = 1 Then
    Rollback
    Kill fdfilename
    Unload FRMOMACHI
    Set FRMOMACHI = Nothing
    If makefd <> 1 Then
        MsgBox "出力テストに失敗しました。データを確認してください。"
    Else
        MsgBox "出力テストに成功しました。"
    End If
    Exit Function
End If
If makefd <> 1 Then
    Rollback
    Kill fdfilename
    Unload FRMOMACHI
    Set FRMOMACHI = Nothing
    Exit Function
Else
    CommitTrans
End If





  
  'ここから


        MsgBox "指定フォルダーにデータを書き込みを開始します。"
        
        flopy = Trim(txtde.Text)
        If Right(flopy, 1) <> "\" Then
            flopy = flopy & "\"
        End If
        
        flopy = flopy & bi2 & "\"
        
        
        If bid = "5169" Or bid = "1386" Then
          '  flopy = "a:\koufuri"
             flopy = flopy & "koufuri"
        ElseIf bid = "9900" Then
            'flopy = "a:\JIHARAI"
             flopy = flopy & "JIHARAI"
            '再読み込みをし、ｃｒ＋ｌｆを追加
            If add_crlf_9900(fdfilename) = -1 Then
                Kill fdfilename
                Unload FRMOMACHI
                Set FRMOMACHI = Nothing
                Exit Function
            End If
        Else
            'flopy = "a:\koufuri.dat"
            flopy = flopy & "koufuri.dat"
        End If
        
        
            
        On Error GoTo fdcopy3
            FileCopy fdfilename, flopy
            DoEvents
        On Error GoTo 0
       


'ここまで
  

'*********trans*****************
BeginTrans
'Bank記録作成
           If kidou2 = False Then
                MsgBox "データベースファイルを開けませんでした。"
                        makefd = 4
                        Rollback  '*********roll******
                        Screen.MousePointer = 0
                Exit Function
            End If
            
            sql = "select * from bankkiroku where hidu ='" & Mid(fdhiduke, 1, 6) & "' and bankid ='" & bid & "'"
            If FcSQlGet2(fdrs3, sql, PrMsg) = False Then
                '新規
                Dim bankkirokuid As String, newbankkirokuid As String
                If soukatsu(0, "renban", 7, "", 0, 0) = False Then
                        MsgBox "設定の読み込みに失敗しました。"
                        makefd = 4
                        Rollback  '*********roll******
                        Screen.MousePointer = 0
                        Exit Function
                End If
                If reg = "0" Then
                    reg = ""
                End If
                If reg = vbNullString Then
                    bankkirokuid = "00001"
                Else
                    bankkirokuid = Format$(reg, "0000#")
                End If
                sql = "INSERT INTO bankkiroku VALUES('" & bankkirokuid & "'," & _
                                "'" & Mid(fdhiduke, 1, 6) & "','" & bid & "'," & _
                                "'" & CStr(toresuu) & "','" & CStr(torekin) & "','0','0','0','0')"
            On Error GoTo errbankkiroku
                 db2.Execute sql, dbSQLPassThrough
            On Error GoTo 0
                 newbankkirokuid = CStr(CLng(bankkirokuid) + 1)
                    If soukatsu(1, "renban", 7, newbankkirokuid, 0, 0) = False Then
                            Rollback  '*********roll******
                            MsgBox "設定の書き込みに失敗しました。"
                            Exit Function
                    End If
            
                            
            Else
                '更新
                 sql = "UPDATE bankkiroku SET desuu = '" & CStr(toresuu) & "'," & _
                              "dekin = '" & CStr(torekin) & "'," & _
                              "nyusuu = '0'," & _
                              "nyukin = '0'," & _
                              "fusuu = '0'," & _
                              "fukin = '0'" & _
                              " WHERE id = '" & fdrs3!id & "'"
            On Error GoTo errbankkiroku
                 db2.Execute sql, dbSQLPassThrough
            On Error GoTo 0
                fdrs3.Close
            End If
            
            
            
            db2.Close

            '〆のflagをたてる
             FRMOMACHI.lblmsg.Caption = "個人データを更新中..."
            DoEvents
               Dim pbp As Long
                sql = "select kojin.kojinid from kojin LEFT JOIN bank ON kojin.bankid = bank.bankid" & _
                            " where bank.bankcode='" & bid & "' and kojin.soukei >'0'"
                'recordset
                If FcSQlGet(fdrs2, sql, PrMsg) = False Then
                       makefd = 7
                      Rollback  '*********roll******
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                       Exit Function
                End If
                 FRMOMACHI.pb.Visible = True
                FRMOMACHI.pb.Min = 0
                FRMOMACHI.pb.Max = fdrs2.RecordCount
                pbp = 1
                FRMOMACHI.pb.Value = pbp
                DoEvents
                fdrs2.MoveFirst
                Do Until fdrs2.EOF
                    If flgup(1, fdrs2!kojinid) = False Then
                        makefd = 8
                        Rollback  '*********roll******
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        Exit Function
                    End If
                    fdrs2.MoveNext
                    FRMOMACHI.pb.Value = pbp
                    DoEvents
                    pbp = pbp + 1
                Loop
                fdrs2.Close
                
                'NHK集計データ作成
                Dim sql_nhksakusei As String, rs_nhksakusei As Recordset, shiteinhkbi As String
                Dim conutid As String, newcountid As String, nhk_data_dou As Long, nhk_data_dou2 As Long
                If kidou4 = False Then
                            makefd = 16
                            Rollback  '*********roll******
                            Unload FRMOMACHI
                            Set FRMOMACHI = Nothing
                            Screen.MousePointer = 0
                            MsgBox "ＮＨＫデータＯＰＥＮに失敗しました。"
                            Exit Function
                End If
                
                    If soukatsu(0, "renban", 10, "", 1, 0) = False Then
                        makefd = 16
                        Rollback  '*********roll******
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        MsgBox "設定の読み込みに失敗しました。"
                        db4.Close
                        Exit Function
                    End If
                    If reg = "0" Then
                        conutid = "000001"
                    Else
                        conutid = Format(reg, "000000")
                    End If

                    shiteinhkbi = Mid(fdhiduke, 1, 6) & "00" '引き落とし日
                    sql_nhksakusei = "select * from nhk_k where seikyuubi ='" & shiteinhkbi & "'"
                    nhk_data_dou = FcSQlGet4(rs_nhksakusei, sql_nhksakusei, PrMsg)
                    If nhk_data_dou = 0 Then
                        sql_nhksakusei = "select * from nhk_m where jikai ='" & shiteinhkbi & "'"
                        nhk_data_dou2 = FcSQlGet4(rs_nhksakusei, sql_nhksakusei, PrMsg)
                        If nhk_data_dou2 > 0 Then
                            rs_nhksakusei.MoveFirst
                            Do Until rs_nhksakusei.EOF
                            '集計用データの作成
                                sql_nhksakusei = "INSERT INTO nhk_k VALUES('" & conutid & "'," & _
                                                   "'" & rs_nhksakusei!kojinid & "','" & rs_nhksakusei!jikai & "'," & _
                                                   "'" & rs_nhksakusei!kingaku & "','" & rs_nhksakusei!shurui & "'," & _
                                                   "'0','" & Space(1) & "'," & _
                                                   "'" & Space(1) & "')"
                                On Error GoTo errnhksakusei1
                                    db4.Execute sql_nhksakusei, dbSQLPassThrough
                                On Error GoTo 0

                                newcountid = CStr(CLng(conutid) + 1)
                                conutid = Format(newcountid, "000000")
                                rs_nhksakusei.MoveNext
                            Loop
                         
                            If soukatsu(1, "renban", 10, newcountid, 0, 1) = False Then
                                makefd = 16
                                Rollback  '*********roll******
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
                                    MsgBox "設定の書き込みに失敗しました。"
                                    db4.Close
                                    Exit Function
                            End If
                         
                         ElseIf nhk_data_dou2 = -1 Then
                            makefd = 16
                            Rollback  '*********roll******
                            Unload FRMOMACHI
                            Set FRMOMACHI = Nothing
                            Screen.MousePointer = 0
                            MsgBox "設定の参照に失敗しました。"
                            db4.Close
                            Exit Function
                        End If
                    ElseIf nhk_data_dou = -1 Then
                        makefd = 16
                        Rollback  '*********roll******
                        Unload FRMOMACHI
                        Set FRMOMACHI = Nothing
                        Screen.MousePointer = 0
                        MsgBox "設定の参照に失敗しました。"
                        db4.Close
                        Exit Function

                    End If
                Unload FRMOMACHI
                Set FRMOMACHI = Nothing
                Screen.MousePointer = 0
'************commit****************
CommitTrans
                
Exit Function

fdcopy:
    Screen.MousePointer = 0
    MsgBox "フロッピーディスクを入れてください。"
    Screen.MousePointer = 11
    Resume
fdcopy2:
    Screen.MousePointer = 0
    MsgBox "2枚目のフロッピーディスクを入れてください。"
    Screen.MousePointer = 11
    Resume
fdcopy3:
    Screen.MousePointer = 0
    MsgBox "書き込みに失敗しました。"
    Screen.MousePointer = 11
  Screen.MousePointer = 0
           Exit Function
errbankkiroku:
           makefd = 14
           Rollback  '*********roll******
            Unload FRMOMACHI
            Set FRMOMACHI = Nothing
            Screen.MousePointer = 0
           Exit Function
errnhksakusei1:
           makefd = 16
           Rollback  '*********roll******
            Unload FRMOMACHI
            Set FRMOMACHI = Nothing
            Screen.MousePointer = 0
           Exit Function





End Function
Function add_crlf_9900(hozon_file_path As String) As Integer

Dim data_nunber As Integer
Dim my_data() As String

Dim my_data_len As Integer
Dim data_nunber2 As Integer

Dim my_data2() As String
Dim datacounter As Integer
Dim chkcount As Integer

'*****************一括読み込み********************

On Error GoTo errfff
    Open hozon_file_path For Input Access Read As #1
On Error GoTo 0
data_nunber = 0
Do While Not EOF(1)
    data_nunber = data_nunber + 1
    ReDim Preserve my_data(data_nunber)
    Line Input #1, my_data(data_nunber)
Loop

'For data_nunber = 1 To UBound(my_data)
'    MsgBox my_data(data_nunber)
'Next data_nunber

Close #1

'*******************120で切る**************************
If data_nunber <> 1 Then
    add_crlf_9900 = -1
    MsgBox "読み込み文字列が1つではありません・・"
    Exit Function
End If

my_data_len = Len(my_data(1))
data_nunber2 = Fix(my_data_len / 120)
chkcount = Int((my_data_len / 120) + 0.9)

If data_nunber2 <> chkcount Then
    add_crlf_9900 = -1
    MsgBox "読み込み文字数が１２０で割り切れません。・・"
    Exit Function
End If
For datacounter = 1 To data_nunber2
    ReDim Preserve my_data2(datacounter)
    my_data2(datacounter) = Mid(my_data(1), (120 * (datacounter - 1)) + 1, 120)
Next datacounter

'For datacounter = 1 To UBound(my_data2)
'    MsgBox my_data2(datacounter)
'Next datacounter

'*********************再構成*********************************
On Error GoTo errfff
    Open hozon_file_path For Output Access Write As #1
On Error GoTo 0

For datacounter = 1 To data_nunber2
    Print #1, my_data2(datacounter)
Next datacounter
Close #1

add_crlf_9900 = 1
'MsgBox "ok"
Exit Function

errfff:
    add_crlf_9900 = -1
    MsgBox "ファイルオープンに失敗しました。"
    Exit Function

End Function
Function yuubinkyoku2(bc As String, hk As String) As Integer

Dim fdrs As Recordset, pbv As Long
Screen.MousePointer = 11

'ｓｑｌ文
        sql = "SELECT kojin.*,bank.bankcode,bank.shitencode,bank.bankmei" & _
                " FROM kojin LEFT JOIN bank ON kojin.bankid = bank.bankid" & _
                " WHERE bank.bankcode='" & bc & "'and clng(kojin.soukei) > 0"
'recordset
     If FcSQlGet(fdrs, sql, PrMsg) = False Then
        yuubinkyoku2 = 1
        Screen.MousePointer = 0
        Exit Function
     End If
     
'書き込み
 FRMOMACHI.pb.Visible = True
 FRMOMACHI.pb.Min = 0
 FRMOMACHI.pb.Max = fdrs.RecordCount
 FRMOMACHI.pb.Value = 0
 FRMOMACHI.lblmsg.Caption = "ファイルを作成中..."
DoEvents


Dim itakushakodo As Long, itakushamei As String, hikiotoshibi As Integer, headdata(3) As Integer
Dim bankcode As Long, bankmei As String, hdami3 As String, shitencode As Long
Dim kouzabangou As Long, hdami As String, khi As String, hdami2 As String, newgyoumukubun As String
Dim hdami4 As String

Dim dshitencode As Long, ddami2 As String, ddami As String, ddami3 As String, ddami4 As String
Dim dkouzabangou As Long, yokinshamei As String, hikiotoshikin As Long, ddami7 As String
Dim kokyakuid As Long, ddata(1) As Integer, ddami5 As String, ddami6 As String
Dim kokyakuid2 As String
Dim goukeisuu As Long, goukeikingaku As Long, tdami As String, tdata As Integer, tdami2 As String, tdami3 As String
Dim damidami As Long
Dim edami As String, edata(0) As Integer

'銀行別ヘッダーレコード
Dim ret As Integer
'郵便局
toresuu = 0
torekin = 0

headdata(0) = 1

'契約種別コード（２）

If soukatsu(0, "yuubinkyoku", 1, "", 1, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yuubinkyoku2 = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yuubinkyoku2 = 3
    Screen.MousePointer = 0
    Exit Function
End If
headdata(1) = CLng(reg)

hdami = Space(1)
newgyoumukubun = "1"

'受け入れ事務センターコード（２）

If soukatsu(0, "yuubinkyoku", 2, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yuubinkyoku2 = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yuubinkyoku2 = 3
    Screen.MousePointer = 0
    Exit Function
End If
headdata(2) = CLng(reg)

'事業主番号（８）
If soukatsu(0, "yuubinkyoku", 3, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yuubinkyoku2 = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yuubinkyoku2 = 3
    Screen.MousePointer = 0
    Exit Function
End If

itakushakodo = CLng(reg)        '委託者コード

If soukatsu(0, "yuubinkyoku", 4, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yuubinkyoku2 = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yuubinkyoku2 = 3
    Screen.MousePointer = 0
    Exit Function
End If
itakushamei = CStr(reg)   '委託者名

hikiotoshibi = CInt(Mid(hk, 5))   '引落日

bankcode = 9900
bankmei = "ﾕｳﾋﾞﾝｷﾖｸ"

If soukatsu(0, "yuubinkyoku", 5, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yuubinkyoku2 = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yuubinkyoku2 = 3
    Screen.MousePointer = 0
    Exit Function
End If

shitencode = CLng(reg)

hdami2 = Space(17)

If soukatsu(0, "yuubinkyoku", 6, "", 0, 1) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yuubinkyoku2 = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    Screen.MousePointer = 0
    yuubinkyoku2 = 3
    Exit Function
End If
kouzabangou = CLng(reg) '口座番号
hdami3 = Space(4)
headdata(3) = 1
hdami4 = Space(12)

ret = fd1.File_Write("D1D2X1DRZ2DRZ8XLS40DRZ4DRZ4XLS15DRZ3X17DRZ6X4X1X12", _
    headdata(0), headdata(1), newgyoumukubun, headdata(2), itakushakodo, itakushamei, hikiotoshibi, bankcode, bankmei, _
    shitencode, hdami2, kouzabangou, hdami3, hdami, hdami4)


'エラー発生
If ret < 0 Then
    errorfd (ret)
    MsgBox "Head data(" & headdata(0) & "," & headdata(1) & "," & hdami & "," & headdata(2) & _
            itakushakodo & "," & itakushamei & "," & hikiotoshibi & "," & bankcode & "," & bankmei & _
            shitencode & "," & hdami2 & "," & kouzabangou & "," & hdami3 & "," & headdata(3) & "," & hdami4 & ")"
    yuubinkyoku2 = 2
    Screen.MousePointer = 0
    Exit Function
End If
    pbv = 0
    Dim barlen As Long, karishite As String, kariban As String, banlen As Long
    fdrs.MoveFirst
    Do Until fdrs.EOF
        ddata(0) = 2
        If Trim(fdrs!kouzabangou) = "" Then
            yuubinkyoku2 = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            '通常預金番号
            barlen = InStr(1, fdrs!kouzabangou, "-")
            If barlen = 6 Then
                karishite = Mid(fdrs!kouzabangou, 1, barlen - 1)
                dshitencode = CLng(Mid(karishite, 2, 3))
                kariban = Mid(fdrs!kouzabangou, 7)
                banlen = Len(kariban)
                dkouzabangou = CLng(Mid(kariban, 1, banlen - 1))
            Else 'エラー
                yuubinkyoku2 = 6
                MsgBox fdrs!kojinid & Space(3) & fdrs!Furigana & Space(3) & fdrs!kouzabangou
                Screen.MousePointer = 0
                Exit Function
            End If
        End If
        ddami = Space(15)
        ddami2 = Space(2)
        ddami3 = Space(1)
        
        If Trim(fdrs!meigi) = "" Then
            yuubinkyoku2 = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            yokinshamei = CStr(fdrs!meigi)
        End If
        If Trim(fdrs!soukei) = "" Then
            yuubinkyoku2 = 4
            MsgBox fdrs!kojinid & Space(3) & yokinshamei
            Screen.MousePointer = 0
            Exit Function
        Else
            hikiotoshikin = CLng(fdrs!soukei)
             toresuu = toresuu + 1
             torekin = torekin + CLng(fdrs!soukei)
        
        End If
        ddami4 = Space(1)
        If Trim(fdrs!kojinid) = "" Then
            yuubinkyoku2 = 4
            MsgBox fdrs!kojinid & Space(3) & yokinshamei
            Screen.MousePointer = 0
            Exit Function
        Else
            kokyakuid = CLng(fdrs!kojinid)
            kokyakuid2 = Trim(fdrs!kojinid)
        End If
        ddata(1) = 0
        ddami5 = Space(4)
        ddami6 = Space(2)
        ddami7 = Space(2)
        'データ・レコード
        ret = fd1.File_Write("D1DRZ4XLS15DRZ3X15X2D2X1DRZ7XLS30DRZ10X1DRZ20D1X4X2X2", _
            ddata(0), bankcode, bankmei, dshitencode, ddami, ddami2, headdata(1), ddami3, _
            dkouzabangou, yokinshamei, hikiotoshikin, ddami4, kokyakuid, ddata(1), ddami5, ddami6, ddami7)

        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            MsgBox "Head data(" & ddata(0) & "," & bankcode & "," & bankmei & "," & dshitencode & _
                    ddami & "," & ddami2 & "," & ddami2 & "," & ddami3 & _
                    dkouzabangou & "," & yokinshamei & "," & hikiotoshikin & "," & ddami4 & _
                    kokyakuid & "," & ddata(1) & "," & ddami5 & "," & ddami6 & "," & ddami7 & ")"
            yuubinkyoku2 = 2
            Screen.MousePointer = 0
            Exit Function
        End If
        If shousaiset(kokyakuid2, hk, hikiotoshikin, , bc) = -1 Then
            MsgBox "詳細データベースへの書き込みに失敗しました。" & ddata(0) & "," & bankcode & "," & bankmei & "," & dshitencode & _
                    ddami & "," & ddami2 & "," & ddami2 & "," & ddami3 & _
                    dkouzabangou & "," & yokinshamei & "," & hikiotoshikin & "," & ddami4 & _
                    kokyakuid & "," & ddata(1) & "," & ddami5 & "," & ddami6 & "," & ddami7 & ")"
            yuubinkyoku2 = 2
            Screen.MousePointer = 0
            Exit Function
        End If
        
        fdrs.MoveNext
        pbv = pbv + 1
        FRMOMACHI.pb.Value = pbv
        DoEvents
    Loop
    
    fdrs.Close
    FRMOMACHI.pb.Visible = False
    


'トレーラ・レコード
tdata = 8

goukeisuu = toresuu
goukeikingaku = torekin
tdami = Space(6)
tdami2 = Space(12)
damidami = 0
tdami3 = Space(29)
ret = fd1.File_Write("D1DRZ6DRZ12DRZ6DRZ12DRZ6DRZ12X6X6X6X6X12X29", _
    tdata, goukeisuu, goukeikingaku, damidami, damidami, damidami, damidami, tdami, tdami, tdami, tdami, _
     tdami2, tdami3)
    
        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            yuubinkyoku2 = 2
            MsgBox "Head data(" & tdata & "," & goukeisuu & "," & goukeikingaku & "," & tdami & _
                    tdami2 & "," & tdami & "," & tdami2 & "," & goukeisuu & "," & tdami & _
                    tdami & "," & tdami & "," & tdami2 & "," & tdami3 & ")"
            Screen.MousePointer = 0
            Exit Function
        End If
    
'エンド・レコード
                
edami = Space(119)
edata(0) = 9
ret = fd1.File_Write("D1X119", edata(0), edami)
     
        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            yuubinkyoku2 = 2
            Screen.MousePointer = 0
            Exit Function
        End If
     
Screen.MousePointer = 0
End Function
Function yuubinkyoku(bc As String, hk As String) As Integer

Dim fdrs As Recordset, pbv As Long
Screen.MousePointer = 11

'ｓｑｌ文
        sql = "SELECT kojin.*,bank.bankcode,bank.shitencode,bank.bankmei" & _
                " FROM kojin LEFT JOIN bank ON kojin.bankid = bank.bankid" & _
                " WHERE bank.bankcode='" & bc & "'and clng(kojin.soukei) > 0"
'recordset
     If FcSQlGet(fdrs, sql, PrMsg) = False Then
        yuubinkyoku = 1
        Screen.MousePointer = 0
        Exit Function
     End If
     
'書き込み
 FRMOMACHI.pb.Visible = True
 FRMOMACHI.pb.Min = 0
 FRMOMACHI.pb.Max = fdrs.RecordCount
 FRMOMACHI.pb.Value = 0
 FRMOMACHI.lblmsg.Caption = "ファイルを作成中..."
DoEvents


Dim itakushakodo As Long, itakushamei As String, hikiotoshibi As Integer, headdata(3) As Integer
Dim bankcode As Long, bankmei As String, hdami3 As String, shitencode As Long
Dim kouzabangou As Long, hdami As String, khi As String, hdami2 As String
Dim hdami4 As String

Dim dshitencode As Long, ddami2 As String, ddami As String, ddami3 As String, ddami4 As String
Dim dkouzabangou As Long, yokinshamei As String, hikiotoshikin As Long, ddami7 As String
Dim kokyakuid As Long, ddata(1) As Integer, ddami5 As String, ddami6 As String

Dim goukeisuu As Long, goukeikingaku As Long, tdami As String, tdata As Integer, tdami2 As String, tdami3 As String

Dim edami As String, edata(0) As Integer

'銀行別ヘッダーレコード
Dim ret As Integer
'郵便局
toresuu = 0
torekin = 0

headdata(0) = 1

'契約種別コード（２）

If soukatsu(0, "yuubinkyoku", 1, "", 1, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yuubinkyoku = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yuubinkyoku = 3
    Screen.MousePointer = 0
    Exit Function
End If
headdata(1) = CLng(reg)

hdami = Space(1)

'受け入れ事務センターコード（２）

If soukatsu(0, "yuubinkyoku", 2, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yuubinkyoku = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yuubinkyoku = 3
    Screen.MousePointer = 0
    Exit Function
End If
headdata(2) = CLng(reg)

'事業主番号（８）
If soukatsu(0, "yuubinkyoku", 3, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yuubinkyoku = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yuubinkyoku = 3
    Screen.MousePointer = 0
    Exit Function
End If

itakushakodo = CLng(reg)        '委託者コード

If soukatsu(0, "yuubinkyoku", 4, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yuubinkyoku = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yuubinkyoku = 3
    Screen.MousePointer = 0
    Exit Function
End If
itakushamei = CStr(reg)   '委託者名

hikiotoshibi = CInt(Mid(hk, 5))   '引落日

bankcode = 9900
bankmei = "ﾕｳｾｲｼﾖｳﾁﾖｷﾝｷﾖｸ "

If soukatsu(0, "yuubinkyoku", 5, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yuubinkyoku = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yuubinkyoku = 3
    Screen.MousePointer = 0
    Exit Function
End If

shitencode = CLng(reg)

hdami2 = Space(17)

If soukatsu(0, "yuubinkyoku", 6, "", 0, 1) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yuubinkyoku = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    Screen.MousePointer = 0
    yuubinkyoku = 3
    Exit Function
End If
kouzabangou = CLng(reg) '口座番号
hdami3 = Space(4)
headdata(3) = 1
hdami4 = Space(12)

ret = fd1.File_Write("D1D2X1D2DRZ8XLS40DRZ4DRZ4XLS15D3X17DRZ6X4D1X12", _
    headdata(0), headdata(1), hdami, headdata(2), itakushakodo, itakushamei, hikiotoshibi, bankcode, bankmei, _
    shitencode, hdami2, kouzabangou, hdami3, headdata(3), hdami4)


'エラー発生
If ret < 0 Then
    errorfd (ret)
    MsgBox "Head data(" & headdata(0) & "," & headdata(1) & "," & hdami & "," & headdata(2) & _
            itakushakodo & "," & itakushamei & "," & hikiotoshibi & "," & bankcode & "," & bankmei & _
            shitencode & "," & hdami2 & "," & kouzabangou & "," & hdami3 & "," & headdata(3) & "," & hdami4 & ")"
    yuubinkyoku = 2
    Screen.MousePointer = 0
    Exit Function
End If
    pbv = 0
    Dim barlen As Long, karishite As String, kariban As String, banlen As Long
    fdrs.MoveFirst
    Do Until fdrs.EOF
        ddata(0) = 2
        If Trim(fdrs!kouzabangou) = "" Then
            yuubinkyoku = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            '通常預金番号
            barlen = InStr(1, fdrs!kouzabangou, "-")
            If barlen = 6 Then
                karishite = Mid(fdrs!kouzabangou, 1, barlen - 1)
                dshitencode = CLng(Mid(karishite, 2, 3))
                kariban = Mid(fdrs!kouzabangou, 7)
                banlen = Len(kariban)
                dkouzabangou = CLng(Mid(kariban, 1, banlen - 1))
            Else 'エラー
                yuubinkyoku = 6
                MsgBox fdrs!kojinid & Space(3) & fdrs!Furigana & Space(3) & fdrs!kouzabangou
                Screen.MousePointer = 0
                Exit Function
            End If
        End If
        ddami = Space(15)
        ddami2 = Space(2)
        ddami3 = Space(1)
        
        If Trim(fdrs!meigi) = "" Then
            yuubinkyoku = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            yokinshamei = CStr(fdrs!meigi)
        End If
        If Trim(fdrs!soukei) = "" Then
            yuubinkyoku = 4
            MsgBox fdrs!kojinid & Space(3) & yokinshamei
            Screen.MousePointer = 0
            Exit Function
        Else
            hikiotoshikin = CLng(fdrs!soukei)
             toresuu = toresuu + 1
             torekin = torekin + CLng(fdrs!soukei)
        
        End If
        ddami4 = Space(1)
        If Trim(fdrs!kojinid) = "" Then
            yuubinkyoku = 4
            MsgBox fdrs!kojinid & Space(3) & yokinshamei
            Screen.MousePointer = 0
            Exit Function
        Else
            kokyakuid = CLng(fdrs!kojinid)
        End If
        ddata(1) = 0
        ddami5 = Space(4)
        ddami6 = Space(2)
        ddami7 = Space(2)
        'データ・レコード
        ret = fd1.File_Write("D1DRZ4XLS15DRZ3X15X2X2X1DRZ7XLS30DRZ10X1DRZ20D1X4X2X2", _
            ddata(0), bankcode, bankmei, dshitencode, ddami, ddami2, ddami2, ddami3, _
            dkouzabangou, yokinshamei, hikiotoshikin, ddami4, kokyakuid, ddata(1), ddami5, ddami6, ddami7)

        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            MsgBox "Head data(" & ddata(0) & "," & bankcode & "," & bankmei & "," & dshitencode & _
                    ddami & "," & ddami2 & "," & ddami2 & "," & ddami3 & _
                    dkouzabangou & "," & yokinshamei & "," & hikiotoshikin & "," & ddami4 & _
                    kokyakuid & "," & ddata(1) & "," & ddami5 & "," & ddami6 & "," & ddami7 & ")"
            yuubinkyoku = 2
            Screen.MousePointer = 0
            Exit Function
        End If
        fdrs.MoveNext
        pbv = pbv + 1
        FRMOMACHI.pb.Value = pbv
        DoEvents
    Loop
    
    fdrs.Close
    FRMOMACHI.pb.Visible = False
    


'トレーラ・レコード
tdata = 8

goukeisuu = toresuu
goukeikingaku = torekin
tdami = Space(6)
tdami2 = Space(12)
tdami3 = Space(29)
ret = fd1.File_Write("D1DRZ6DRZ12X6X12X6X12DRZ6X6X6X6X12X29", _
    tdata, goukeisuu, goukeikingaku, tdami, tdami2, tdami, tdami2, goukeisuu, tdami, tdami, tdami, _
     tdami2, tdami3)
    
        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            yuubinkyoku = 2
            MsgBox "Head data(" & tdata & "," & goukeisuu & "," & goukeikingaku & "," & tdami & _
                    tdami2 & "," & tdami & "," & tdami2 & "," & goukeisuu & "," & tdami & _
                    tdami & "," & tdami & "," & tdami2 & "," & tdami3 & ")"
            Screen.MousePointer = 0
            Exit Function
        End If
    
'エンド・レコード
                
edami = Space(119)
edata(0) = 9
ret = fd1.File_Write("D1X119", edata(0), edami)
     
        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            yuubinkyoku = 2
            Screen.MousePointer = 0
            Exit Function
        End If
     
Screen.MousePointer = 0
End Function

Function yamachuu(bc As String, hk As String) As Integer

Dim fdrs As Recordset, pbv As Long
Screen.MousePointer = 11

'ｓｑｌ文
        sql = "SELECT kojin.*,bank.bankcode,bank.shitencode,bank.bankmei" & _
                " FROM kojin LEFT JOIN bank ON kojin.bankid = bank.bankid" & _
                " WHERE bank.bankcode='" & bc & "'and clng(kojin.soukei) > 0"
'recordset
     If FcSQlGet(fdrs, sql, PrMsg) = False Then
        yamachuu = 1
        Screen.MousePointer = 0
        Exit Function
     End If
     
'書き込み
 FRMOMACHI.pb.Visible = True
 FRMOMACHI.pb.Min = 0
 FRMOMACHI.pb.Max = fdrs.RecordCount
 FRMOMACHI.pb.Value = 0
 FRMOMACHI.lblmsg.Caption = "ファイルを作成中..."
DoEvents

Dim itakushakodo As Long, itakushamei As String, hikiotoshibi As Integer, headdata(2) As Integer
Dim bankcode As Long, bankmei As String, shitencode As Long, shitenmei As String
Dim shubetsu As Integer, kouzabangou As Long, hdami As String, khi As String

Dim dshitencode As Long, dshitenmei As String, ddami As String, dshubetsu As Integer
Dim dkouzabangou As Long, yokinshamei As String, hikiotoshikin As Long
Dim kokyakuid As Long, ddami2 As String, ddata(2) As Integer
Dim kokyakuid2 As String
Dim goukeisuu As Long, goukeikingaku As Long, tdami As String, tdata(1) As Integer

Dim edami As String, edata(0) As Integer

'銀行別ヘッダーレコード
Dim ret As Integer
'山梨中央銀行
toresuu = 0
torekin = 0

headdata(0) = 1
headdata(1) = 91
headdata(2) = 0 '1  2018/09/11 修正

If soukatsu(0, "yamachuu", 1, "", 1, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yamachuu = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yamachuu = 3
    Screen.MousePointer = 0
    Exit Function
End If

itakushakodo = CLng(reg)        '委託者コード

If soukatsu(0, "yamachuu", 2, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yamachuu = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yamachuu = 3
    Screen.MousePointer = 0
    Exit Function
End If
itakushamei = CStr(reg)   '委託者名

hikiotoshibi = CInt(Mid(hk, 5))   '引落日

bankcode = 142
bankmei = "ﾔﾏﾅｼﾁｭｳｵｳ"

If soukatsu(0, "yamachuu", 4, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yamachuu = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yamachuu = 3
    Screen.MousePointer = 0
    Exit Function
End If
shitencode = CLng(reg) '支店コード

If soukatsu(0, "yamachuu", 5, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yamachuu = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yamachuu = 3
    Screen.MousePointer = 0
    Exit Function
End If
shitenmei = CStr(reg) '支店名

If soukatsu(0, "yamachuu", 6, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yamachuu = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    yamachuu = 3
    Screen.MousePointer = 0
    Exit Function
End If
shubetsu = CInt(reg)  '種別
    
If soukatsu(0, "yamachuu", 7, "", 0, 1) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        yamachuu = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    Screen.MousePointer = 0
    yamachuu = 3
    Exit Function
End If
kouzabangou = CLng(reg) '口座番号

hdami = Space(17)

ret = fd1.File_Write("D1D2D1DRZ10XLS40DRZ4DRZ4XLS15D3XLS15D1DRZ7X17", _
    headdata(0), headdata(1), headdata(2), itakushakodo, itakushamei, hikiotoshibi, bankcode, bankmei, _
    shitencode, shitenmei, shubetsu, kouzabangou, hdami)


'エラー発生
If ret < 0 Then
    errorfd (ret)
    MsgBox "Head data(" & headdata(0) & "," & headdata(1) & "," & headdata(2) & "," & itakushakodo & _
            itakushamei & "," & hikiotoshibi & "," & bankcode & "," & bankmei & _
            shitencode & "," & shitenmei & "," & shubetsu & "," & kouzabangou & "," & hdami & ")"
    yamachuu = 2
    Screen.MousePointer = 0
    Exit Function
End If
    pbv = 0
    
    fdrs.MoveFirst
    Do Until fdrs.EOF
        ddata(0) = 2
        If Trim(fdrs!shitencode) = "" Then
            yamachuu = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            dshitencode = CLng(fdrs!shitencode)
        End If
        dshitenmei = Space(15)
        ddami = Space(4)
        If fdrs!kouzakubun = 0 Then '普通
            dshubetsu = 1
        ElseIf fdrs!kouzakubun = 1 Then '当座
            dshubetsu = 2
        Else
            yamachuu = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        End If
        If Trim(fdrs!kouzabangou) = "" Then
            yamachuu = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            On Error GoTo erryama
            dkouzabangou = CLng(fdrs!kouzabangou)
            On Error GoTo 0
        End If
        If Trim(fdrs!meigi) = "" Then
            yamachuu = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            yokinshamei = CStr(fdrs!meigi)
        End If
        If Trim(fdrs!soukei) = "" Then
            yamachuu = 4
            MsgBox fdrs!kojinid & Space(3) & yokinshamei
            Screen.MousePointer = 0
            Exit Function
        Else
            hikiotoshikin = CLng(fdrs!soukei)
        
             toresuu = toresuu + 1
             torekin = torekin + CLng(fdrs!soukei)
        End If
        ddata(1) = 1
        If Trim(fdrs!kojinid) = "" Then
            yamachuu = 4
            MsgBox fdrs!kojinid & Space(3) & yokinshamei
            Screen.MousePointer = 0
            Exit Function
        Else
            kokyakuid = CLng(fdrs!kojinid)
            kokyakuid2 = Trim(fdrs!kojinid)
        End If
        ddata(2) = 0
        ddami2 = Space(8)
        'データ・レコード
        ret = fd1.File_Write("D1DRZ4XLS15D3X15X4D1DRZ7XLS30DRZ10D1DRZ20D1X8", _
            ddata(0), bankcode, bankmei, dshitencode, dshitenmei, ddami, dshubetsu, _
            dkouzabangou, yokinshamei, hikiotoshikin, ddata(1), kokyakuid, ddata(2), ddami2)

        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            MsgBox "Head data(" & ddata(0) & "," & bankcode & "," & bankmei & "," & dshitencode & _
                    dshitenmei & "," & ddami & "," & dshubetsu & _
                    dkouzabangou & "," & yokinshamei & "," & hikiotoshikin & "," & ddata(1) & _
                    kokyakuid & "," & ddata(2) & "," & ddami2 & ")"
            yamachuu = 2
            Screen.MousePointer = 0
            Exit Function
        End If
        If kokyakuid2 = "00969" Then
           ' MsgBox "Hit"
        End If
        If shousaiset(kokyakuid2, hk, hikiotoshikin, , bc) = -1 Then
            MsgBox "詳細データベースへの書き込みに失敗しました。" & ddata(0) & "," & bankcode & "," & bankmei & "," & dshitencode & _
                    dshitenmei & "," & ddami & "," & dshubetsu & _
                    dkouzabangou & "," & yokinshamei & "," & hikiotoshikin & "," & ddata(1) & _
                    kokyakuid & "," & ddata(2) & "," & ddami2 & ")"
            yamachuu = 2
            Screen.MousePointer = 0
            Exit Function
        End If
        
        fdrs.MoveNext
        pbv = pbv + 1
        FRMOMACHI.pb.Value = pbv
        DoEvents
    Loop
    fdrs.Close
    FRMOMACHI.pb.Visible = False
    


'トレーラ・レコード
tdata(0) = 8
tdata(1) = 0

goukeisuu = toresuu
goukeikingaku = torekin
tdami = Space(65)
ret = fd1.File_Write("D1DRZ6DRZ12DRZ6DRZ12DRZ6DRZ12X65", _
    tdata(0), goukeisuu, goukeikingaku, tdata(1), tdata(1), tdata(1), tdata(1), tdami)
    
        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            MsgBox "Head data(" & tdata(0) & "," & goukeisuu & "," & goukeikingaku & "," & tdata(1) & _
                    tdata(1) & "," & tdata(1) & "," & tdata(1) & "," & tdami & ")"
            yamachuu = 2
            Screen.MousePointer = 0
            Exit Function
        End If
    
'エンド・レコード

edami = Space(119)
edata(0) = 9
ret = fd1.File_Write("D1X119", edata(0), edami)
     
        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            yamachuu = 2
            Screen.MousePointer = 0
            Exit Function
        End If
     
Screen.MousePointer = 0

Exit Function

erryama:
            MsgBox "口座番号が不正です。契約ID：" & fdrs!kojinid & ",口座番号：" & fdrs!kouzabangou
            yamachuu = 2
            Screen.MousePointer = 0
            Exit Function


End Function
Function jayamanashi(bc As String, hk As String) As Integer

Dim fdrs As Recordset, pbv As Long
Screen.MousePointer = 11

'ｓｑｌ文
        sql = "SELECT kojin.*,bank.bankcode,bank.shitencode,bank.bankmei" & _
                " FROM kojin LEFT JOIN bank ON kojin.bankid = bank.bankid" & _
                " WHERE bank.bankcode='" & bc & "'and clng(kojin.soukei) > 0"
'recordset
     If FcSQlGet(fdrs, sql, PrMsg) = False Then
        jayamanashi = 1
        Screen.MousePointer = 0
        Exit Function
     End If
'書き込み
 FRMOMACHI.pb.Visible = True
 FRMOMACHI.pb.Min = 0
 FRMOMACHI.pb.Max = fdrs.RecordCount
 FRMOMACHI.pb.Value = 0
 FRMOMACHI.lblmsg.Caption = "農協のファイルを作成中..."
DoEvents

Dim itakushakodo As Long, itakushamei As String, hikiotoshibi As Integer, headdata(2) As Integer
Dim bankcode As Long, bankmei As String, shitencode As Long, shitenmei As String
Dim shubetsu As Integer, kouzabangou As Long, hdami As String, khi As String

Dim dshitencode As Long, dshitenmei As String, ddami As String, dshubetsu As Integer
Dim dkouzabangou As Long, yokinshamei As String, hikiotoshikin As Long
Dim kokyakuid As Long, ddami2 As String, ddata(2) As Integer
Dim kokyakuid2 As String
Dim goukeisuu As Long, goukeikingaku As Long, tdami As String, tdata(1) As Integer


'銀行別ヘッダーレコード
Dim ret As Integer
'ＪＡ
toresuu = 0
torekin = 0
headdata(0) = 1
headdata(1) = 91
headdata(2) = 0

If soukatsu(0, "jakihon", 1, "", 1, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        jayamanashi = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    jayamanashi = 3
    Screen.MousePointer = 0
    Exit Function
End If

itakushakodo = CLng(reg)        '委託者コード

If soukatsu(0, "jakihon", 2, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        jayamanashi = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    jayamanashi = 3
    Screen.MousePointer = 0
    Exit Function
End If
itakushamei = CStr(reg)   '委託者名

hikiotoshibi = CInt(Mid(hk, 5))   '引落日

bankcode = 5169
bankmei = Space(15)

If soukatsu(0, "jakihon", 4, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        jayamanashi = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    jayamanashi = 3
    Screen.MousePointer = 0
    Exit Function
End If
shitencode = CLng(reg) '支店コード

If soukatsu(0, "jakihon", 5, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        jayamanashi = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    jayamanashi = 3
    Screen.MousePointer = 0
    Exit Function
End If
shitenmei = CStr(reg) '支店名



If soukatsu(0, "jakihon", 6, "", 0, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        jayamanashi = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    jayamanashi = 3
    Screen.MousePointer = 0
    Exit Function
End If
shubetsu = CInt(reg)  '種別
    
If soukatsu(0, "jakihon", 7, "", 0, 1) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        jayamanashi = 3
        Screen.MousePointer = 0
        Exit Function
End If
If reg = "0" Then
    reg = ""
End If
If reg = "" Then
    Screen.MousePointer = 0
    jayamanashi = 3
    Exit Function
End If
kouzabangou = CLng(reg) '口座番号

hdami = Space(17)

ret = fd1.File_Write("D1D2D1DRZ10XLS40DRZ4DRZ4XLS15DRZ3XLS15D1DRZ7X17x2", _
    headdata(0), headdata(1), headdata(2), itakushakodo, itakushamei, hikiotoshibi, bankcode, bankmei, _
    shitencode, shitenmei, shubetsu, kouzabangou, hdami, YAJI)



'エラー発生
If ret < 0 Then
    errorfd (ret)
    jayamanashi = 2
    Screen.MousePointer = 0
    Exit Function
End If
    pbv = 0
    fdrs.MoveFirst
    Do Until fdrs.EOF
    
        If fdrs!kojinid = "01402" Then
        'MsgBox "01402"
        End If
    
        ddata(0) = 2
        If Trim(fdrs!shitencode) = "" Then
            jayamanashi = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            dshitencode = CLng(fdrs!shitencode)
        End If
        dshitenmei = Space(15)
        ddami = Space(4)
        If fdrs!kouzakubun = 0 Then '普通
            dshubetsu = 1
        ElseIf fdrs!kouzakubun = 1 Then '当座
            dshubetsu = 2
        Else
            jayamanashi = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        End If
        If Trim(fdrs!kouzabangou) = "" Then
            jayamanashi = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            On Error GoTo errjkou
                dkouzabangou = CLng(fdrs!kouzabangou)
            On Error GoTo 0
        End If
        If Trim(fdrs!meigi) = "" Then
            jayamanashi = 4
            MsgBox fdrs!kojinid & Space(3) & fdrs!meigi
            Screen.MousePointer = 0
            Exit Function
        Else
            yokinshamei = CStr(fdrs!meigi)
        End If
        If Trim(fdrs!soukei) = "" Then
            jayamanashi = 4
            MsgBox fdrs!kojinid & Space(3) & yokinshamei
            Screen.MousePointer = 0
            Exit Function
        Else
            hikiotoshikin = CLng(fdrs!soukei)
             toresuu = toresuu + 1
             torekin = torekin + CLng(fdrs!soukei)
        
        End If
        ddata(1) = 1
        If Trim(fdrs!kojinid) = "" Then
            jayamanashi = 4
            MsgBox fdrs!kojinid & Space(3) & yokinshamei
            Screen.MousePointer = 0
            Exit Function
        Else
            kokyakuid = CLng(fdrs!kojinid)
            kokyakuid2 = Trim(fdrs!kojinid)
        End If
        ddata(2) = 0
        ddami2 = Space(8)
        'データ・レコード
        ret = fd1.File_Write("D1DRZ4XLS15DRZ3X15X4D1DRZ7XLS30DRZ10D1DRZ20D1X8x2", _
            ddata(0), bankcode, bankmei, dshitencode, dshitenmei, ddami, dshubetsu, _
            dkouzabangou, yokinshamei, hikiotoshikin, ddata(1), kokyakuid, ddata(2), ddami2, YAJI)

        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            MsgBox kokyakuid & Space(3) & yokinshamei
            jayamanashi = 2
            Screen.MousePointer = 0
            Exit Function
        End If
        
        If shousaiset(kokyakuid2, hk, hikiotoshikin, , bc) = -1 Then
            MsgBox "詳細データベースへの書き込みに失敗しました。"
            jayamanashi = 2
            Screen.MousePointer = 0
            Exit Function
        End If
        
        fdrs.MoveNext
        pbv = pbv + 1
        FRMOMACHI.pb.Value = pbv
        DoEvents
    Loop
    fdrs.Close
    FRMOMACHI.pb.Visible = False
    


'トレーラ・レコード
tdata(0) = 8
tdata(1) = 0

toresuu = toresuu
torekin = torekin
tdami = Space(65)
ret = fd1.File_Write("D1DRZ6DRZ12DRZ6DRZ12DRZ6DRZ12X65x2", _
    tdata(0), toresuu, torekin, tdata(1), tdata(1), tdata(1), tdata(1), tdami, YAJI)
    
        'エラー発生
        If ret < 0 Then
            errorfd (ret)
            jayamanashi = 2
            Screen.MousePointer = 0
            Exit Function
        End If
    
'エンドコード

           Dim edami As String, edata As Integer
           edami = Space(119)
           edata = 9
           ret = fd1.File_Write("D1X119x2", edata, edami, YAJI)

           'エラー発生
           If ret < 0 Then
               errorfd (ret)
               jayamanashi = 2
               Screen.MousePointer = 0
               Exit Function
           End If


Screen.MousePointer = 0
Exit Function
errjkou:
    MsgBox fdrs!kojinid & Space(1) & "の口座番号が不正です。"
    jayamanashi = 2
    Screen.MousePointer = 0
    Exit Function

End Function


Private Sub Form_Load()

sbnbankset5

Dim DesktopPath As String
' デスクトップのパスを取得
DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
' 表示
'MsgBox DesktopPath
Dim s_path As String
Dim s_path2 As String
Dim s_path3 As String
Dim s_path4 As String


s_path = DesktopPath & "\EZManager_data\"
s_path2 = DesktopPath & "\EZManager_data\出力\"
s_path3 = DesktopPath & "\EZManager_data\入力\"
s_path4 = DesktopPath & "\EZManager_data\仮読込\"

txtde.Text = s_path2

'なければ作成
'FileSystemObjectインスタンスを生成
Dim fi As New FileSystemObject

'フォルダの存在確認
If fi.FolderExists(s_path) Then
    'Call MsgBox("sampleフォルダは存在します")
Else
    'Call MsgBox("sampleフォルダは存在しません")
    ' フォルダ (ディレクトリ) を作成する
    MkDir (s_path)
    DoEvents
End If

'オブジェクトの解放
Set fi = Nothing

'FileSystemObjectインスタンスを生成
Dim fi2 As New FileSystemObject

'フォルダの存在確認
If fi2.FolderExists(s_path2) Then
    'Call MsgBox("sampleフォルダは存在します")
Else
    'Call MsgBox("sampleフォルダは存在しません")
    ' フォルダ (ディレクトリ) を作成する
    MkDir (s_path2)
    DoEvents
End If
'オブジェクトの解放
Set fi2 = Nothing

'なければ作成
'FileSystemObjectインスタンスを生成
Dim fi4 As New FileSystemObject

'フォルダの存在確認
If fi4.FolderExists(s_path3) Then
    'Call MsgBox("sampleフォルダは存在します")
Else
    'Call MsgBox("sampleフォルダは存在しません")
    ' フォルダ (ディレクトリ) を作成する
    MkDir (s_path3)
    DoEvents
End If

'オブジェクトの解放
Set fi4 = Nothing

Dim bi As String
Dim kaku_path As String
Dim kaku_path2 As String
'FileSystemObjectインスタンスを生成
Dim fi3 As New FileSystemObject

For i = 0 To combank.ListCount - 1
    bi = Trim(Mid(combank.List(i), 5))
    kaku_path = s_path2 & "\" & bi & "\"
    
    'フォルダの存在確認
    If fi3.FolderExists(kaku_path) Then
        'Call MsgBox("sampleフォルダは存在します")
    Else
        'Call MsgBox("sampleフォルダは存在しません")
        ' フォルダ (ディレクトリ) を作成する
        MkDir (kaku_path)
        DoEvents
    End If
  
Next

'オブジェクトの解放
Set fi3 = Nothing


'FileSystemObjectインスタンスを生成
Dim fi5 As New FileSystemObject

'フォルダの存在確認
If fi5.FolderExists(s_path4) Then
    'Call MsgBox("sampleフォルダは存在します")
Else
    'Call MsgBox("sampleフォルダは存在しません")
    ' フォルダ (ディレクトリ) を作成する
    MkDir (s_path4)
    DoEvents
End If

'オブジェクトの解放
Set fi5 = Nothing



End Sub

Private Sub Picture13_Click()
Unload Me
End Sub

Private Sub Picture21_Click()
dirdou = 5
frmdir.Show 1

End Sub

Private Sub Picture7_Click()

Dim bi As String, testshutsuryoku As Integer
        If combank.ListIndex = -1 Then
            MsgBox "金融機関名を選択してください。"
            combank.SetFocus
            Exit Sub
        End If
        bi = Mid(combank.Text, 1, 4)
           Dim bi2 As String
         bi2 = Trim(Mid(combank.Text, 5))
         
      
        If bi = "0000" Then
            MsgBox "不明金融機関なので、フロッピーディスクの作成はできません。"
            Exit Sub
        End If
        
       
            If Trim(txtde.Text) = "" Then
                MsgBox "出力先の設定がありません。"
                Exit Sub
            End If
        
        
        sql = "select kojin.kojinid from kojin LEFT JOIN bank ON kojin.bankid = bank.bankid" & _
                    " where bank.bankcode='" & bi & "'"
        'recordset
        If FcSQlGet(rs, sql, PrMsg) = False Then
            MsgBox "作成するデータがありません。"
            Exit Sub
        End If
        
        
        
        rs.Close
        
        If chktest.Value = 1 Then
            testshutsuryoku = 1
        Else
            testshutsuryoku = 0
        End If
        
        Select Case makefd(bi, testshutsuryoku, bi2)
        Case 1
        
        
            'Log
           
            sagyou_msg = "FDD出力：" & bi2 & "の請求データを出力しました。"
            log_sagyou sagyou_log_path, sagyou_msg
        
            If testshutsuryoku = 1 Then
                Exit Sub
            End If
            
             
            MsgBox "請求データの作成が完了しました。"
           
            
            fdd_sai_check bi, bi2
            
        Case 2
            MsgBox "作成するデータがありません。"
        Case 3    '異常終了のコード
            
        Case 4
            MsgBox "契約内容が登録されていません。「設定」で契約内容を登録してください。"
        Case 5
            MsgBox "個人の登録データが不正です。"
        Case 6
            MsgBox "集計ができませんでした。"
        Case 7
            MsgBox "予期せぬ事由によってエラーが発生しました。"
        Case 8
            MsgBox "個人情報に「〆」のフラッグの書き込みに失敗しました。"
        Case 9
            'MsgBox "フロッピーディスクを入れてから、再度実行してください。"
        Case 10
            MsgBox "口座番号が不正です。"
        Case 11
            MsgBox "すでにフロッピーへの出力処理が行われています。"
        Case 12
            MsgBox "フロッピーディスクを入れてください。"
        Case 13
            MsgBox "選択した金融機関が不正です。"
        Case 14
            MsgBox "バンク更新記録の作成に失敗しました。"
        Case 15
            MsgBox "初期化しますので、再度実行してください"
        Case 16
            MsgBox "ＮＨＫデータの更新に失敗しました。"
        End Select
        frmfdd.Show 1
        Exit Sub

End Sub


