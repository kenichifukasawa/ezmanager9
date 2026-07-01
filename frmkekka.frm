VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmkekka 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "検索結果"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   13860
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1140
      MouseIcon       =   "frmkekka.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmkekka.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   7770
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   9600
      MouseIcon       =   "frmkekka.frx":1741
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmkekka.frx":2040
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   7740
      Width           =   480
   End
   Begin VB.PictureBox Picture15 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   11205
      MouseIcon       =   "frmkekka.frx":2E82
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmkekka.frx":3781
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   7740
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   12810
      MouseIcon       =   "frmkekka.frx":45C3
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmkekka.frx":4EC2
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   7740
      Width           =   480
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   7035
      Left            =   180
      TabIndex        =   0
      Top             =   540
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   12409
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
   Begin VB.Label lblseikyuukingaku 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12900
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "請求書"
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
      Index           =   2
      Left            =   1020
      TabIndex        =   12
      Top             =   8310
      Width           =   720
   End
   Begin VB.Label Label1 
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
      Index           =   1
      Left            =   9600
      TabIndex        =   10
      Top             =   8280
      Width           =   480
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
      Left            =   11070
      TabIndex        =   8
      Top             =   8280
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ｷｬﾝｾﾙ"
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
      Left            =   12690
      TabIndex        =   6
      Top             =   8280
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "件"
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
      Left            =   12120
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblkensuu 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10440
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblshu 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   7815
   End
   Begin VB.Label lblkubun 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   300
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   915
      Left            =   8760
      Top             =   7680
      Width           =   4920
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   915
      Left            =   180
      Top             =   7680
      Width           =   2460
   End
End
Attribute VB_Name = "frmkekka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit














Private Sub Form_Load()
    

' View プロパティを詳細表示 (Report) に設定します。
   lv1.View = lvwReport

End Sub

Private Sub Picture1_Click()

If Trim(lblkubun.Caption) <> "請求書" Then

        Exit Sub
End If


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
 '   shiteppp = shiteppp & "\滞納者_" & Format(Date, "yyyymmdd") & "-" & Format(Time, "hhmmss") & ".csv"



Dim csvsuu As Long, chfileout As String, flopy As String

chfileout = shiteppp & "\" & Trim(lblkubun.Caption) & "_" & Format(Date, "yyyymmdd") & "_" & Format(Time, "hhmmss")
chfileout = chfileout & ".csv"





  
  Open chfileout For Output As #1 'Len = 32000
  Write #1, "NO", "ID", "契約者名", "フリガナ", "郵便番号", "住所", "電話番号", "請求額", "更新"

  For csvsuu = 1 To csv_data_k_count
      Write #1, CStr(Format(csvsuu, "000"));
      Write #1, CStr(csv_data_keiyaku(0, csvsuu));
      Write #1, CStr(csv_data_keiyaku(1, csvsuu));
      Write #1, CStr(csv_data_keiyaku(2, csvsuu));
      Write #1, CStr(csv_data_keiyaku(3, csvsuu));
      Write #1, CStr(csv_data_keiyaku(8, csvsuu)) & CStr(csv_data_keiyaku(4, csvsuu));
      Write #1, CStr(csv_data_keiyaku(5, csvsuu));
      Write #1, CStr(csv_data_keiyaku(6, csvsuu));
       Write #1, CStr(csv_data_keiyaku(7, csvsuu))
  Next
  Close #1
        
        Screen.MousePointer = 0
        MsgBox "作成しました。"
Exit Sub
errcsv:
        MsgBox "出力に失敗しました。"
        Exit Sub
End Sub

Private Sub Picture13_Click()
Unload frmkekka
Set frmkekka = Nothing
Unload frmsentaku
Set frmsentaku = Nothing

End Sub

Private Sub Picture15_Click()
    
    Dim appaccess As New Access.Application
    Dim ss As Long
    
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

Private Sub Picture2_Click()

  Dim s_path As String, y_count As Integer
  

    settei_res = setting(1, 14, 0, "", 1)
    If settei_res = "-1" Or settei_res = "0" Then
            MsgBox "DropBoxの設定がありません。"
            Exit Sub
    Else
        s_path = Trim(settei_res)
    End If
    
     If Right(s_path, 1) = "\" Then
        s_path = Mid(s_path, 1, Len(s_path) - 1)
    End If
    
    Dim k_t As Integer, k_n As Long
    
    If Format(Date, "mm") = "01" Then
        k_t = 12
        k_n = CLng(Format(Date, "yyyy")) - 1
    Else
        k_t = CInt(Format(Date, "mm")) - 1
        k_n = CLng(Format(Date, "yyyy"))
    End If
    
    
    msg = "山梨CATVへケーブルラインの請求データを送信しますか？。"
    res = MsgBox(msg, vbYesNo)
    If res = vbNo Then
        Exit Sub
    End If

    
    
    
    
     Dim s_saki_file As String
     
     s_saki_file = s_path & "\請求書データ\請求書_cl_" & Format(k_n, "0000") & "_" & Format(k_t, "00") & ".txt"
     
     Dim s_seikyuukingaku As String
     
     s_seikyuukingaku = Trim(lblshu.Caption)
     
     Dim s_kensuu As String
     
      s_kensuu = Trim(Me.lblkensuu.Caption)
     
     
      'データを2つに保存
    '初期化
     If Dir(s_saki_file) <> "" Then
         Kill s_saki_file
    End If
   
    
    '作成
    '山梨用データコピー
    Dim writecounter As Integer
    Open s_saki_file For Output Access Write As 1
    
     Write #1, s_seikyuukingaku, s_kensuu & "件"
     
    Write #1, "個人ID・名前", "電話番号"
    
    For writecounter = 0 To csv_data_cl_count - 1
        'Write #1, y_data(1, writecounter), y_data(2, writecounter), y_data(3, writecounter), y_data(4, writecounter), y_data(5, writecounter), y_data(6, writecounter), y_data(7, writecounter), y_data(8, writecounter), y_data(9, writecounter), y_data(10, writecounter), y_data(11, writecounter), y_data(12, writecounter), y_data(13, writecounter), y_data(14, writecounter), y_data(15, writecounter), y_data(16, writecounter), y_data(17, writecounter), y_data(18, writecounter)
    
        Write #1, csv_data_cl(0, writecounter), csv_data_cl(1, writecounter)
    Next
    Close #1

 
                               
                                
                                
    MsgBox "CLの請求書データを作成し、送信しました。"
    
End Sub
