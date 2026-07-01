VERSION 5.00
Object = "{271E16FB-2B6D-11D4-B06D-0090991097BF}#26.0#0"; "ZoonFIO.ocx"
Begin VB.Form frmkoushin2 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "仮読み込み"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   7770
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6870
      MouseIcon       =   "frmkoushin2.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmkoushin2.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   1620
      Width           =   480
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5880
      MouseIcon       =   "frmkoushin2.frx":11C9
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmkoushin2.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   1590
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "読み込み先"
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
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   7515
      Begin VB.TextBox txtde 
         Height          =   345
         IMEMode         =   4  '全角ひらがな
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   5925
      End
      Begin VB.PictureBox Picture21 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   6630
         MouseIcon       =   "frmkoushin2.frx":2392
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmkoushin2.frx":2C91
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   300
         Width           =   480
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
         TabIndex        =   3
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
   Begin ZoonFIO.ZoonFileIO fd2 
      Left            =   120
      Top             =   1770
      _ExtentX        =   900
      _ExtentY        =   900
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
      Left            =   6900
      TabIndex        =   7
      Top             =   2100
      Width           =   495
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
      Left            =   5880
      TabIndex        =   6
      Top             =   2100
      Width           =   480
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   855
      Left            =   5640
      Top             =   1530
      Width           =   1890
   End
End
Attribute VB_Name = "frmkoushin2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim DesktopPath As String
' デスクトップのパスを取得
DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
' 表示
'MsgBox DesktopPath
txtde.Text = DesktopPath & "\EZManager_data\仮読込\"
End Sub

Private Sub Picture10_Click()

Dim motofile As String, sakifile As String, ret As Integer, kensars As Recordset
Dim hdatas(7) As String, hdatal(8) As Long
Dim ddatas() As String, ddatal() As Long
Dim tdatas(1) As String, tdatal(6) As Long, kongetsu As String
Dim koushinid As String, koushinnewhi As String, newfilemei As String
Dim yuubinkakunin As Integer, bkrs4 As Recordset
    Dim nou As Integer
Dim motofile2 As String
Dim seikyuudata(), shousaidata(), tsuwadata(), FilePath(3)

Dim s_bankmei As String

   With fd2
    
    
    'ここから
       
       Me.Hide
       
        If Trim(txtde.Text) = "" Then
            MsgBox "入力先の設定がありません。"
            Exit Sub
        End If
        motofile2 = Trim(txtde.Text)
        If Right(motofile2, 1) <> "\" Then
            motofile2 = motofile2 & "\"
        End If
            
          nou = 0
         FRMOMACHI.lblmsg.Caption = "仮読み込みの準備中..."
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
            sakifile = sopa & "fdback2\"
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
                        ret = .File_Read("D1D2D1D1D9X40D4D4X15D3X15D1D7X17", _
                            hdatal(0), hdatal(1), hdatal(2), itakushacode, hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                            hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3))
                    
                           'エラーが発生したら、郵便局を参照
                         'エラー発生
                        If ret < 0 Then
                            If .File_Seek(1) = False Then
                                MsgBox "読み込み位置の移動に失敗しました。"
                                .Active = False
                                Unload FRMOMACHI
                                Set FRMOMACHI = Nothing
                                Screen.MousePointer = 0
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
                    
                    
                    If ret < 0 Then
                        errorfd (ret)
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
                    msg = "金融機関「農業協同組合」の" & Format(Mid(Format(hdatal(4), "0000"), 1, Len(hdatal(4)) - 2), "00") & "/" & Right(hdatal(4), 2) & "データを仮読み込みしますか？"
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
                            msg = "金融機関「" & kensars!bankname & "」の" & Format(Mid(Format(hdatal(4), "0000"), 1, Len(hdatal(4)) - 2), "00") & "/" & Right(hdatal(4), 2) & "データを仮読み込みしますか？"
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
                
                 FRMOMACHI.lblmsg.Caption = "データの仮読み込みを準備中..."
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
        
                    ret = .File_Read("D1D6D12D6D12D6D12X65x2", _
                    tdatal(0), tdatal(1), tdatal(2), tdatal(3), tdatal(4), tdatal(5), tdatal(6), tdatas(0), tdatas(1))
                    'エラー発生
                    If ret < 0 Then
                        errorfd (ret)
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
                            
            ret = .File_Read("D1D2D1D10X40D4D4X15D3X15D1D7X17x2", _
                hdatal(0), hdatal(1), hdatal(2), hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3), hdatas(6))
                    If ret < 0 Then
                        errorfd (ret)
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
                            ret = .File_Read("D1D4X15D3X15X4D1D7X30D10D1D20D1X8x2", _
                            ddatal(i, 0), ddatal(i, 1), ddatas(i, 0), ddatal(i, 2), ddatas(i, 1), ddatas(i, 2), _
                            ddatal(i, 3), ddatal(i, 4), ddatas(i, 3), ddatal(i, 5), ddatal(i, 6), _
                            ddatal(i, 7), ddatal(i, 8), ddatas(i, 4), ddatas(i, 11))
                                
                            'エラー発生
                            If ret < 0 Then
                                errorfd (ret)
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
                            ret = .File_Read("D1D6D12D6D12D6D12X65", _
                            tdatal(0), tdatal(1), tdatal(2), tdatal(3), tdatal(4), tdatal(5), tdatal(6), tdatas(0))
                            'エラー発生
                            If ret < 0 Then
                                errorfd (ret)
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
                                    ret = .File_Read("D1D2D1D1D9X40D4D4X15D3X15D1D7X17", _
                                        hdatal(0), hdatal(1), hdatal(2), itakushacode, hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                                        hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3))
                            Else
                                    
                                    ret = .File_Read("D1D2X1D2D8X40D4D4X15D3X17D6X4D1X12", _
                                        hdatal(0), hdatal(1), hdatas(4), hdatal(2), hdatal(3), _
                                        hdatas(0), hdatal(4), hdatal(5), _
                                        hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatas(5), hdatal(8), hdatas(3))
                            End If
                            
                            If ret < 0 Then
                                errorfd (ret)
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
                                    ret = .File_Read("D1D4X15D3X15X2X2X1D7X30D10X1D20D1X4X2X2", _
                                    ddatal(i, 0), ddatal(i, 1), ddatas(i, 0), ddatal(i, 2), ddatas(i, 1), ddatas(i, 2), _
                                    ddatas(i, 5), ddatas(i, 6), _
                                    ddatal(i, 3), ddatas(i, 7), ddatal(i, 4), ddatas(i, 3), ddatal(i, 7), ddatal(i, 8), _
                                    ddatas(i, 9), ddatas(i, 8), ddatas(i, 4))
                                
                                Else
                                    ret = .File_Read("D1D4X15D3X15X4D1D7X30D10D1D20D1X8", _
                                    ddatal(i, 0), ddatal(i, 1), ddatas(i, 0), ddatal(i, 2), ddatas(i, 1), ddatas(i, 2), _
                                    ddatal(i, 3), ddatal(i, 4), ddatas(i, 3), ddatal(i, 5), ddatal(i, 6), _
                                    ddatal(i, 7), ddatal(i, 8), ddatas(i, 4))
                                End If
                                'エラー発生
                                If ret < 0 Then
                                    errorfd (ret)
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
                            ret = .File_Read("D1D6D12D6D12D6D12X65X2", _
                            tdatal(0), tdatal(1), tdatal(2), tdatal(3), tdatal(4), tdatal(5), tdatal(6), tdatas(0), dami_crlf)
                            'エラー発生
                            If ret < 0 Then
                                errorfd (ret)
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
                                    ret = .File_Read("D1D2D1D10X40D4D4X15D3X15D1D7X17X2", _
                                        hdatal(0), hdatal(1), hdatal(2), hdatal(3), hdatas(0), hdatal(4), hdatal(5), _
                                        hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatal(8), hdatas(3), dami_crlf)
                            Else
                                    
                                    ret = .File_Read("D1D2X1D2D8X40D4D4X15D3X17D6X4D1X12", _
                                        hdatal(0), hdatal(1), hdatas(4), hdatal(2), hdatal(3), _
                                        hdatas(0), hdatal(4), hdatal(5), _
                                        hdatas(1), hdatal(6), hdatas(2), hdatal(7), hdatas(5), hdatal(8), hdatas(3))
                            End If
                            
                            If ret < 0 Then
                                errorfd (ret)
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
                                    ret = .File_Read("D1D4X15D3X15X2X2X1D7X30D10X1D20D1X4X2X2", _
                                    ddatal(i, 0), ddatal(i, 1), ddatas(i, 0), ddatal(i, 2), ddatas(i, 1), ddatas(i, 2), _
                                    ddatas(i, 5), ddatas(i, 6), _
                                    ddatal(i, 3), ddatas(i, 7), ddatal(i, 4), ddatas(i, 3), ddatal(i, 7), ddatal(i, 8), _
                                    ddatas(i, 9), ddatas(i, 8), ddatas(i, 4))
                                
                                Else
                                    ret = .File_Read("D1D4X15D3X15X4D1D7X30D10D1D20D1X8X2", _
                                    ddatal(i, 0), ddatal(i, 1), ddatas(i, 0), ddatal(i, 2), ddatas(i, 1), ddatas(i, 2), _
                                    ddatal(i, 3), ddatal(i, 4), ddatas(i, 3), ddatal(i, 5), ddatal(i, 6), _
                                    ddatal(i, 7), ddatal(i, 8), ddatas(i, 4), dami_crlf)
                                End If
                                'エラー発生
                                If ret < 0 Then
                                    errorfd (ret)
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
                        
                        
  'ファイルの有無の確認
  
            If Dir(newfilemei) <> "" Then
                    MsgBox "すでに更新されています。仮読み込みはできません。"
                  Unload FRMOMACHI
                  Set FRMOMACHI = Nothing
                  '*******ロールバック*************
              
                  Kill sakifile
                  Screen.MousePointer = 0
                  
                  Exit Sub
                        
            End If
                        
                        
                        On Error GoTo erronajifile
                            Name sakifile As newfilemei
                        On Error GoTo 0
            FRMOMACHI.pb.Value = 0
            

'*********トランザクション開始**************
BeginTrans
                           
                           
            'Bank記録作成
              If kidou2 = False Then
                  MsgBox "データベースファイルを開けませんでした。"
                  Unload FRMOMACHI
                  Set FRMOMACHI = Nothing
                  '*******ロールバック*************
              
                  Kill newfilemei
                  Screen.MousePointer = 0
                  
                  Exit Sub
              End If
              
                           
             FRMOMACHI.lblmsg.Caption = "更新中..."
            DoEvents
                        koushinnewhi = Format(Date, "yyyy") & Format(Mid(Format(hdatal(4), "0000"), 1, Len(hdatal(4)) - 2), "00") & Right(hdatal(4), 2)
                        '個人データの処理
                            Dim kaisuu As Long
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
                            
                            kaisuu = tdatal(1)
                            
                             Dim koushnid As String, newkoushinid As String
                            
                            If soukatsu(0, "startup", 14, "", 1, 1) = False Then
                                    MsgBox "設定の読み込みに失敗しました。"
                                     Unload FRMOMACHI
                                    Set FRMOMACHI = Nothing
                                    '*******ロールバック*************
                                    Kill newfilemei
                                    Screen.MousePointer = 0
                                    
                                    Exit Sub
                            End If
                            If reg = "0" Then
                                reg = ""
                            End If
                            If reg = vbNullString Then
                                koushnid = "0000000001"
                            ElseIf reg = "999999999" Then
                                MsgBox "これ以上の更新記録の作成はできません。管理者に連絡してください。"
                                 Unload FRMOMACHI
                                    Set FRMOMACHI = Nothing
                                    '*******ロールバック*************
                                    Kill newfilemei
                                    Screen.MousePointer = 0
                                    
                                    Exit Sub
                            Else
                                koushnid = Format$(reg, "000000000#")
                            End If
                             newkoushinid = CStr(CLng(koushnid) + kaisuu)
                            If soukatsu(1, "startup", 14, newkoushinid, 1, 0) = False Then
                                    MsgBox "設定の書き込みに失敗しました。"
                                     Unload FRMOMACHI
                                    Set FRMOMACHI = Nothing
                                    '*******ロールバック*************
                                    Kill newfilemei
                                    Screen.MousePointer = 0
                                    
                                    Exit Sub
                            End If
                            
                            
                            Dim s_id As String, s_nentsuki As String, s_kojinid As String, s_kingaku As String, s_bankid As String, s_st As String
                            Dim s_koushinbi As String
            
                            For i = 1 To kaisuu
                                
                                 '更新記録を残す
  
                                    s_id = koushnid
                                    s_nentsuki = subhiduke
                                    s_koushinbi = Format(Date, "yyyy") & Format(Mid(Format(hdatal(4), "0000"), 1, Len(hdatal(4)) - 2), "00") & Right(hdatal(4), 2)
                                    s_kojinid = Format(ddatal(i, 7), "00000")
                                    If subbankid = "9900" Or subbankid = "5169" Then
                                        s_kingaku = ddatal(i, 5)
                                    Else
                                        s_kingaku = ddatal(i, 4)
                                    End If
                                    s_bankid = subbankid
                                     s_st = ddatal(i, 8)
                                            
                                    
                                sql = "INSERT INTO koushinkiroku2 VALUES('" & s_id & "'," & _
                                               "'" & s_nentsuki & "','" & s_koushinbi & "','" & s_kojinid & "'," & _
                                               "'" & s_kingaku & "','" & s_bankid & "','" & s_st & "')"
                                               
                                On Error GoTo errko2
                                    db2.Execute sql, dbSQLPassThrough
                                On Error GoTo 0
                                
                                koushnid = CStr(CLng(koushnid) + 1)
                                            
                               koushnid = Format$(koushnid, "000000000#")
                               
                                    FRMOMACHI.pb.Value = i
                                    DoEvents
            
                            Next i
                            Unload FRMOMACHI
                            Set FRMOMACHI = Nothing
                            Screen.MousePointer = 0
                           
                            
                          
    '*************トランザクション終了********
    CommitTrans
                          
                          
            'Log
           
            sagyou_msg = "仮読込：" & s_bankmei & "の振替データを仮読み込みしました。"
            log_sagyou sagyou_log_path, sagyou_msg
                          
                            'db2.Close
                            MsgBox "更新処理が終了しました。"
            
            End With
            
            Exit Sub
            
            
errm1:
     Rollback
    Kill newfilemei
    MsgBox "銀行の移動記録の更新に失敗しました。"
    bkrs.Close
    db2.Close
   
    Exit Sub
    
erronajifile:
  MsgBox "一時読み込みデータがすでに存在します。"
       
         Kill sakifile
       
        Unload FRMOMACHI
        Set FRMOMACHI = Nothing
        Screen.MousePointer = 0
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
errfdnashi:
  MsgBox "フロッピーが入っていません。もう一度入れ直してから実行してください。"
    Exit Sub
    
errko2:
        Unload FRMOMACHI
        MsgBox "仮読み込みデータを更新中にエラーが発生しました。データを復旧します。"
        Set FRMOMACHI = Nothing
        '*******ロールバック*************
        Rollback
        Kill newfilemei
        Screen.MousePointer = 0
        
        Exit Sub
    
End Sub

Private Sub Picture13_Click()
Unload Me
End Sub

Private Sub Picture21_Click()
dirdou = 7
frmdir.Show 1
End Sub
