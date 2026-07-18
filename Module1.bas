Attribute VB_Name = "Module1"
Option Explicit
'********************************************************
'kojin 　F1=  0=普通　1=〆
'　　　  F2
'********************************************************
Public Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
    
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const CSIDL_NETWORK = 0 ' &H12
Public Const MAX_PATH2 = 260
Public Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer
    dwStrucVersionh As Integer
    dwFileVersionMSl As Integer
    dwFileVersionMSh As Integer
    dwFileVersionLSl As Integer
    dwFileVersionLSh As Integer
    dwProductVersionMSl As Integer
    dwProductVersionMSh As Integer
    dwProductVersionLSl As Integer
    dwProductVersionLSh As Integer
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type

Public Type CODEPAGE
    lngLOW As Integer
    lngHIGH As Integer
End Type

Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" ( _
    ByVal lptstrFilename As String, _
    ByVal dwHandle As Long, _
    ByVal dwLen As Long, _
    lpData As Any) As Long

Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" ( _
    ByVal lptstrFilename As String, _
    lpdwHandle As Long) As Long

Public Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" ( _
    pBlock As Any, _
    ByVal lpSubBlock As String, _
    lplpBuffer As Any, _
    puLen As Long) As Long
    
Public Declare Sub MoveMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    dest As Any, _
    ByVal Source As Long, _
    ByVal length As Long)


'[フォルダの参照]ダイアログを表示する関数
'Public Declare Function SHBrowseForFolder Lib "shell32.dll" _
   (lpbi As BROWSEINFO) As Long

'ITEMIDLISTをパスに変換する関数
'Public Declare Function SHGetPathFromIDList Lib "shell32.dll" _
   (ByVal pidl As Long, _
    ByVal pszPath As String) As Long

'メモリブロックを解放する関数
'Public Declare Sub CoTaskMemFree Lib "ole32.dll" _
   (pv As Any)
   
   Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBROWSEINFO As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

'[フォルダの参照]ダイアログ情報
Public Type BROWSEINFO
    hWndOwner      As Long
    pidlRoot       As Long
    pszDisplayName As String
    lpszTitle      As String
    ulFlags        As Long
    lpfn           As Long
    lParam         As Long
    iImage         As Long
End Type

       

Public Const LVIS_STATEIMAGEMASK As Long = &HF000 '
 
Public Const MAX_PATH As Long = 260

Public Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type

Public Declare Function ShellExecute Lib _
      "shell32.dll" Alias "ShellExecuteA" _
      (ByVal hwnd As Long, _
      ByVal lpOperation As String, _
      ByVal lpFile As String, _
      ByVal lpParameters As String, _
      ByVal lpDirectory As String, _
      ByVal nShowCmd As Long) As Long
 

Public a                As String       '簡単な文字列変数
Public bankidtbl()                      '金融機関idテーブル
Public bankidtbl2()                      '金融機関idテーブル
Public bankidtbl3()                      '金融機関idテーブル
Public bankidtbl4()                      '金融機関idテーブル
Public bankcount        As Long         '金融機関数
Public bankcount2       As Long        '金融機関数
Public bankcount3       As Long        '金融機関数
Public bankcount4       As Long         '金融機関数
Public db               As Database     'データベース
Public db2               As Database     'データベース
Public db_voip          As Database
Public db_ppv           As Database
Public db_out           As Database
Public db_convart           As Database
Public donohi           As Integer      '日付を指定する
Public dirdou           As Integer      'どのパスを指定か？
Public fdhiduke         As String
Public i                As Long      'Loopをまわす回数
Public henkanmoto       As String       '変更元の値
Public kidoupassword    As String       '起動時のパスワード
Public koushindata()    As Long         '更新後のデータ
Public msg              As String       'メッセージの内容
Public PrMsg            As String       'エラーメッセージ
Public reg              As String
Public res                              'レスポンスの変数
Public rs               As Recordset    'レコードセット
Public shuruicount      As Integer      '契約種類の数
Public shuruiidtbl()                    '契約種類のIDデーブル
Public sopa             As String       'パス
Public sonotacount      As Integer      'その他の契約数
Public sonotaidtbl()                    'その他IDテーブル
Public sql              As String       'sql文
Public sql2             As String
Public tanka            As Long
Public VBNAME           As String       'mdbパス
Public VBNAME2          As String       'mdbパス
Public vbname_voip      As String
Public vbname_out      As String
Public vbname_convart   As String
Public YAJI As String
Public kenchan As Integer
Public furiganahenkou2 As Integer
Public exever As String
Public filever As String
Public backupsaki1 As String
Public backupsaki2 As String
Public kakaku As Long
'マルチ修正
Public vbname3          As String
Public settei_res As String
Public db3 As Database
Public mapa As String
Public Ret
Public db4 As Database
Public vbname4          As String
Public henkoumae_shurui As String
Public henkoumae_kingaku As String
Public henkoumae_nen As String
Public henkoumae_tsuki As String
Public kidoumei As String


Public datacheck_cansel As Integer
Public cellrs As Recordset
Public cellcount As Integer
Public cellidtbl()

Public vbname_ppv As String

Public mentefaile As String
Public fddkoushinOK As Integer
Public kiidou_mode As Integer
Public shitei_shousaiid As String

Public fd_a As String

Public csv_data_keiyaku()
Public csv_data_k_count As Integer

Public map_exe_path As String
Public ftp_exe_path As String

Public sagyou_log_path As String
Public sagyou_msg  As String

Public nhk_ryoukin_tsuujou As Long
Public nhk_ryoukin_hangaku As Long

Public db_cl As Database
Public vbname_cl As String

Public csv_data_cl()
Public csv_data_cl_count As Integer

Public sub_exe_path As String
Public s_kidou_kojinid As String

Public s_tainou_tsuika_kingaku As Long

Private Function ReadRegValue(ByVal regPath As String, ByVal defaultValue As String) As String
    Dim sh As Object
    Dim v
    On Error Resume Next

    Set sh = CreateObject("WScript.Shell")
    v = sh.RegRead(regPath)

    If Err.Number <> 0 Then
        ReadRegValue = defaultValue
        Err.Clear
    Else
        ReadRegValue = CStr(v)
    End If

    Set sh = Nothing
End Function


Private Function GetWindowsVersionInfo() As String
    Dim baseKey As String
   ' Dim productName As String
   ' Dim displayVersion As String
    Dim build As String
    Dim ubr As String
    Dim buildFull As String

    baseKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\"

    'productName = ReadRegValue(baseKey & "ProductName", "")
    'displayVersion = ReadRegValue(baseKey & "DisplayVersion", "")
    build = ReadRegValue(baseKey & "CurrentBuildNumber", "")
    ubr = ReadRegValue(baseKey & "UBR", "")

    If build <> "" And ubr <> "" Then
        buildFull = build & "." & ubr
    Else
        buildFull = build
    End If

    ' 例: Windows 11 Pro 23H2 (OS ビルド 22621.3007)
  '  GetWindowsVersionInfo = Trim(productName & " " & displayVersion & _
    '                             IIf(buildFull <> "", " (OS ビルド " & buildFull & ")", ""))
                                 
                                 
    GetWindowsVersionInfo = IIf(buildFull <> "", buildFull, "")
    
End Function


Sub cll_grid_set(retsu As Integer)
With frmkojinshousai.grdcl
        .Font.Size = 11
        .Rows = retsu + 1
        .Cols = 2
        .ColWidth(0) = 750
        .ColWidth(1) = 1700
        .Row = 0
        .Col = 0
        .Text = "ID"
        .Col = 1
        .Text = "TEL"
End With
End Sub


Sub cl_grid_set(sentakukojinid As String)

Dim sql_stb As String, rs_stb As Recordset, stbcount As Long, i As Long

sql_stb = "select * from cl where kojinid='" & sentakukojinid & "'"
    If FcSQlGet(rs_stb, sql_stb, PrMsg) = False Then
        cll_grid_set (0)
    Else
        rs_stb.MoveLast
        stbcount = rs_stb.RecordCount
        cll_grid_set (stbcount)
        rs_stb.MoveFirst
        With frmkojinshousai.grdcl
            For i = 1 To stbcount
                    .Row = i
                    .Col = 0
                    .Text = rs_stb!clid
                    .Col = 1
                    .Text = rs_stb!tel
               rs_stb.MoveNext
            Next i
            rs_stb.Close
        End With
    End If

End Sub



Function FcSQlGet_cl(PrRS As Recordset, PrSQL As String, PrMsg As String) As Integer

    'ＳＱＬ文の実行
    On Error GoTo ErrSyori
    Set PrRS = db_cl.OpenRecordset(PrSQL, dbOpenDynaset, dbReadOnly)
    On Error GoTo 0
    
    '結果の確認
    If PrRS.RecordCount = 0 Then
        'PrMsg = LoadResString(141)
        FcSQlGet_cl = 0
        Exit Function
    Else
        PrRS.MoveLast
        'PrRS.MoveFirst
    End If
    
    FcSQlGet_cl = 1
    Exit Function
    
ErrSyori:
    PrMsg = "予期せぬエラーが発生しました。"
    FcSQlGet_cl = -1

End Function

Function kidou_cl() As Boolean
'データのロード

kidoupassword = ";pwd=" & "kamifusafusa"
'Set DB = OpenDatabase(VBNAME)
On Error GoTo db4open2
    Set db_cl = DBEngine.Workspaces(0).OpenDatabase(vbname_cl, False, False, kidoupassword)
On Error GoTo 0
kidou_cl = True
Exit Function

db4open2:
    MsgBox "データベースが開けませんでした。一度再起動し、再度実行してください。"
    kidou_cl = False
    Exit Function
End Function

Sub sbshuruiset_2()
'*****************************************
'*   契約種類の名前をコンボボックスに設定      *
'*****************************************

    Dim shuruirs As Recordset
        
    sql = "SELECT * FROM shurui ORDER BY shuruiid"
    
    If FcSQlGet(shuruirs, sql, PrMsg) = False Then
        'MsgBox WMsg
        Exit Sub
    End If
    
    shuruicount = shuruirs.RecordCount
    
    ReDim shuruiidtbl(shuruirs.RecordCount)    '種類ＩＤテーブルの再定義
    shuruirs.MoveFirst
    i = 0
    Do Until shuruirs.EOF
        a = shuruirs!shuruiid & Space(1) & shuruirs!shuruimei
            frmtourokusentaku2.cmbkeiyakushurui.AddItem a    'コンボボックスに項目を追加
        shuruiidtbl(i) = shuruirs!shuruiid         '種類ＩＤテーブルに保存
        shuruirs.MoveNext
        i = i + 1
    Loop

    shuruirs.Close

End Sub

Sub sbsonotaset_2()
'*****************************************
'*   契約種類の名前をコンボボックスに設定      *
'*****************************************

    Dim sonotars As Recordset
        
    sql = "SELECT * FROM sonota ORDER BY sonotaid"
    
    If FcSQlGet(sonotars, sql, PrMsg) = False Then
        'MsgBox WMsg
        Exit Sub
    End If
    
    sonotacount = sonotars.RecordCount
    
    ReDim sonotaidtbl(sonotars.RecordCount)    '種類ＩＤテーブルの再定義
    sonotars.MoveFirst
    i = 0
    Do Until sonotars.EOF
        a = sonotars!sonotaid & Space(1) & sonotars!sonota
            frmtourokusentaku2.cmbsonotakeiyaku.AddItem a    'コンボボックスに項目を追加
        sonotaidtbl(i) = sonotars!sonotaid         '種類ＩＤテーブルに保存
        sonotars.MoveNext
        i = i + 1
    Loop

    sonotars.Close


End Sub

Function yoteikanri_touroku(s_kojinid As String, s_shurui As String, s_kanrenid As String, _
    s_kanrenmei As String, s_kingaku As String, s_kuchisuu As String, s_sousuu As String, _
    s_itsukara As String, s_itsumade As String, s_adoff As String, s_bikou As String) As Boolean

Dim tid As String, tid2 As String
            
            Dim s_itsu As String
            Dim s_nanji As String
            
            s_itsu = Format(Date, "yyyymmdd")
            s_nanji = Format(Time, "hhmmss")
            
            yoteikanri_touroku = False
            
            If soukatsu(0, "renban", 16, "", 1, 1) = False Then
                    MsgBox "設定の読み込みに失敗しました。"
            End If
            If reg = "0" Then
                reg = ""
            End If
    
            If reg = vbNullString Then
                tid = "0000001"
            Else
                tid = Format$(reg, "000000#")
            End If
            
            tid2 = CStr(CLng(tid) + 1)
            If soukatsu(1, "renban", 16, tid2, 1, 1) = False Then
                    MsgBox "設定の書き込みに失敗しました。"
                    Exit Function
            End If
        
            
       
        sql = "INSERT INTO yotei VALUES('" & tid & "','" & s_kojinid & "'," & _
                       "'" & s_shurui & "','" & s_kanrenid & "','" & s_kanrenmei & "'," & _
                       "'" & s_kingaku & "','" & s_kuchisuu & "','" & s_sousuu & "'," & _
                       "'" & s_itsukara & "','" & s_itsumade & "','" & s_itsu & "'," & _
                       "'" & s_nanji & "','" & Space(1) & "','" & Space(1) & "'," & _
                       "'" & Space(1) & "','" & s_adoff & "','0','" & s_bikou & "')"
        
On Error GoTo errk4
        db.Execute sql, dbSQLPassThrough
On Error GoTo 0
       
        
     
        
        MsgBox "登録しました。"
        
        yoteikanri_touroku = True
Exit Function

errk4:
    MsgBox "登録に失敗しました。"
    Exit Function
End Function

Sub yoteiset(Optional s_id As String = "")

'種類の一覧を表示します。

Dim shokuinsuu As Integer

Screen.MousePointer = 11

sql = "SELECT*FROM yotei ORDER BY yoteiid"

sql = "SELECT yotei.*,kojin.name FROM yotei left join kojin on yotei.kojinid = kojin.kojinid"

Dim s_where As String
s_where = ""

If s_id <> "" Then
    s_where = " where yotei.kojinid = '" & s_id & "'"
End If

If frmyotei.lblno.Caption = "1" Then
    If s_where = "" Then
        s_where = " where yotei.del = space(1)"
    Else
         s_where = s_where & " and yotei.del = space(1)"
    End If
End If

sql = sql & s_where & " ORDER BY yotei.yoteiid"


If FcSQlGet(rs, sql, PrMsg) = False Then
    With frmyotei.yoteigrd
        .Rows = shokuinsuu + 1
        .ColWidth(0) = 800
        .ColWidth(1) = 900
        .ColWidth(2) = 2500
        .ColWidth(3) = 900
         .ColWidth(4) = 800
          .ColWidth(5) = 3500
           .ColWidth(6) = 1000
            .ColWidth(7) = 700
             .ColWidth(8) = 700
              .ColWidth(9) = 1100
               .ColWidth(10) = 1100
                .ColWidth(11) = 3000
                .ColWidth(12) = 2000
                 .ColWidth(13) = 1800
                  .ColWidth(14) = 1800
                   .ColWidth(15) = 1200
        .Row = 0
        .Col = 0
        .Text = "操作"
        .Col = 1
        .Text = "ST"
        .Col = 2
        .Text = "顧客"
        .Col = 3
        .Text = "種類"
        .Col = 4
        .Text = "契約ID"
        .Col = 5
        .Text = "契約名"
        .Col = 6
        .Text = "金額"
        .Col = 7
        .Text = "口数"
        .Col = 8
        .Text = "回数"
        .Col = 9
        .Text = "開始"
        .Col = 10
        .Text = "終了"
         .Col = 11
        .Text = "備考"
        .Col = 12
        .Text = "作成"
        .Col = 13
        .Text = "変更"
        .Col = 14
        .Text = "削除"
        .Col = 15
        .Text = "ID"
        
    End With
    
    frmyotei.Picture4.Enabled = False
    frmyotei.Picture5.Enabled = False

   
    Screen.MousePointer = 0
Else
  '  frmshokuinichi.CBOREN.Clear
   
    shokuinsuu = rs.RecordCount
    rs.MoveFirst
    With frmyotei.yoteigrd
        .Rows = shokuinsuu + 1
         .ColWidth(0) = 800
        .ColWidth(1) = 900
        .ColWidth(2) = 2500
        .ColWidth(3) = 900
         .ColWidth(4) = 800
          .ColWidth(5) = 3500
           .ColWidth(6) = 1000
            .ColWidth(7) = 700
             .ColWidth(8) = 700
              .ColWidth(9) = 1100
               .ColWidth(10) = 1100
                .ColWidth(11) = 3000
                .ColWidth(12) = 2000
                 .ColWidth(13) = 1800
                  .ColWidth(14) = 1800
                   .ColWidth(15) = 1200
        .Row = 0
        .Col = 0
        .Text = "操作"
        .Col = 1
        .Text = "ST"
        .Col = 2
        .Text = "顧客"
        .Col = 3
        .Text = "種類"
        .Col = 4
        .Text = "契約ID"
        .Col = 5
        .Text = "契約名"
        .Col = 6
        .Text = "金額"
        .Col = 7
        .Text = "口数"
        .Col = 8
        .Text = "回数"
        .Col = 9
        .Text = "開始"
        .Col = 10
        .Text = "終了"
         .Col = 11
        .Text = "備考"
        .Col = 12
        .Text = "作成"
        .Col = 13
        .Text = "変更"
        .Col = 14
        .Text = "削除"
        .Col = 15
        .Text = "ID"
        
          frmyotei.Picture4.Enabled = True
        frmyotei.Picture5.Enabled = True
    
    End With
    For i = 1 To shokuinsuu
        With frmyotei.yoteigrd
            .Row = i
            .Col = 0
            If rs!addoff = "0" Then
                .Text = "追加"
                .Col = 3
                If rs!shurui = "0" Then
                    .Text = "主契約"
                ElseIf rs!shurui = "1" Then
                    .Text = "その也"
                Else
                    .Text = ""
                End If
            Else
                .Text = "削除"
                .Col = 3
                If rs!shurui = "0" Then
                    .Text = "削除"
                ElseIf rs!shurui = "1" Then
                    .Text = "休止"
                ElseIf rs!shurui = "2" Then
                    .Text = "中止"
                Else
                    .Text = ""
                End If
            End If
            .Col = 1
            If rs!st = "0" Then
                .Text = "未実施"
            Else
                If Trim(rs!del) = "" Then
                    If Trim(rs!h_itsu) = "" Then
                        .Text = "実施済"
                    Else
                        .Text = "中止"
                    End If
                Else
                    .Text = "削除"
                End If
            End If
            .Col = 2
            .CellAlignment = flexAlignLeftCenter
            .Text = Trim(rs!kojinid) & " " & rs!Name
            
               .Col = 4
            .Text = rs!kanrenid
               .Col = 5
            .Text = rs!kanrenmei
               .Col = 6
            .Text = rs!kingaku
               .Col = 7
            .Text = rs!kuchisuu
               .Col = 8
            .Text = rs!sousuu
               .Col = 9
            .Text = rs!itsukara
               .Col = 10
            .Text = rs!itsumade
            
             .Col = 11
             If IsNull(rs!bikou) Then
                .Text = ""
             Else
                .Text = rs!bikou
            End If
             .Col = 12
            .Text = rs!itsu & " " & rs!nanji
            .Col = 13
            .Text = rs!h_itsu & " " & rs!h_nanji
             .Col = 14
            .Text = rs!del
            .Col = 15
            .Text = rs!yoteiid
            
           
          
        End With
       rs.MoveNext
    Next i
    rs.Close
End If
    Screen.MousePointer = 0
    



End Sub

Function koushinchk() As Boolean
    
    
    If soukatsu(0, "renban", 14, "", 1, 1) = False Then
            MsgBox "設定の読み込みに失敗しました。"
            koushinchk = False
            Exit Function
    End If
    If reg = "0" Then
        koushinchk = True
    Else
         MsgBox "更新処理中です。少し経ってから使用してください。ほかのパソコンが使用していないようなら、初期化を行ってください。"
        koushinchk = False
    End If



End Function
Function rocked(onoff As String) As Boolean
    Dim newa As String
    If onoff = "on" Then
       newa = "1"
    ElseIf onoff = "off" Then
       newa = "0"
    End If
    If soukatsu(1, "renban", 14, newa, 1, 1) = False Then
            rocked = False
    Else
            rocked = True
    End If
End Function

Sub data_kyoutou_to_katsunuma(no As Integer)

    '***********データを指定してコンバートする**************************
   ' このプロシージャはセミコロンで区切られた
   ' ファイルからデータを配列内にロードします。
   Dim GetLine As Variant, FilePath(3) As String
   Dim Spot As Integer, Position As Integer, newkojinid As String
   Dim Row As Integer, Col As Integer, newvoipnospot2 As Integer
   Dim RowCount As Integer, ColCount As Integer, ColCount2 As Integer, ColCount3 As Integer
    Dim udtBrowseInfo As BROWSEINFO
    Dim lngIDList     As Long, newzei As Long, newsou As Long, newvoipnospot As Integer, newv1 As String, newv2 As String, newv3 As String
    Dim strPath       As String, kariatai, sql_voip2 As String, newvoipno As String, sql_voip299 As String
    Dim newtsuwaid As String, voipii As Integer, rs_voip As Recordset, rs_voip2 As Recordset, rs_voip299 As Recordset
    Dim rcrc2 As Integer
    
    Dim err_counter As Integer
    
    
    '[フォルダの参照]ダイアログ情報を設定
    With udtBrowseInfo
        .hWndOwner = frmmain.hwnd
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
        FilePath(2) = strPath
        'ITEMIDLISTを解放
        CoTaskMemFree lngIDList
    End If
    If MsgBox("読み込みと登録データのチェックを開始します。", vbYesNo) = vbNo Then
        Exit Sub
    End If
Screen.MousePointer = 11
    If Right(FilePath(2), 1) = "\" Then
        FilePath(2) = Mid(FilePath(2), 1, Len(FilePath(2)) - 1)
    End If
   'FilePath(0) = FilePath(2) & "\請求データサンプル.txt"
   'FilePath(1) = FilePath(2) & "\請求明細データサンプル.txt"
   'FilePath(2) = FilePath(2) & "\通話明細データサンプル.txt"
   FilePath(0) = FilePath(2) & "\bill_002_" & Format(Date, "yyyymm") & ".txt"
   FilePath(1) = FilePath(2) & "\item_002_" & Format(Date, "yyyymm") & ".txt"
   'FilePath(2) = FilePath(2) & "\call_002_" & Format(Date, "yyyymm") & ".txt"
      'On Error GoTo ErrorTrap
      
    'FDのコピーをHDDに実施
    Dim sakifile9 As String, sakifailego As String
    sakifile9 = sopa & "fdback\voip\"
    If Dir$(sakifile9, vbDirectory) = "" Then
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder (sakifile9)
        Set fso = Nothing
    End If
    If Right(sakifile9, 1) <> "\" Then
        sakifile9 = sakifile9 & "\"
    End If
    If Dir(FilePath(0)) = "" Then
        Screen.MousePointer = 0
        MsgBox "今月の「bill」ファイルがありません。"
        Exit Sub
    End If
    If Dir(FilePath(1)) = "" Then
        Screen.MousePointer = 0
        MsgBox "今月の「item」ファイルがありません。"
        Exit Sub
    End If
    
    sakifailego = sakifile9 & "bill_002_" & Format(Date, "yyyymm") & ".txt"
    FileCopy FilePath(0), sakifailego
    sakifailego = sakifile9 & "item_002_" & Format(Date, "yyyymm") & ".txt"
    FileCopy FilePath(1), sakifailego
    
      
      
      
      
      
      
      
On Error GoTo ErrorTrap3
   Open FilePath(0) For Input Lock Read Write As #1
   ColCount = 0
   RowCount = 8
    While Not EOF(1)
        Line Input #1, GetLine
        ColCount = ColCount + 1
    Wend
    Close #1
    
    Open FilePath(0) For Input Lock Read Write As #1
   ReDim seikyuudata(ColCount, RowCount)
   Position = 1
   Row = 0
   'Line Input #1, GetLine
   'Spot = InStr(1, GetLine, ",")
   While Not EOF(1)
      Line Input #1, GetLine
      Col = 1
      
      Spot = InStr(1, GetLine, ",")
      While Spot <> 0
         kariatai = Left$(GetLine, Spot - 1)
         If Col <> 7 And Col <> 6 Then
            If Mid(kariatai, 1, 1) = """" Then
                If Right(kariatai, 1) = """" Then
                    seikyuudata(Row, Col) = Mid(kariatai, 2, Len(kariatai) - 2)
                Else
                    seikyuudata(Row, Col) = Mid(kariatai, 2)
                End If
            Else
                If Right(kariatai, 1) = """" Then
                    seikyuudata(Row, Col) = Mid(kariatai, 1, Len(kariatai) - 1)
                Else
                    seikyuudata(Row, Col) = kariatai
                End If
            End If
        Else
            seikyuudata(Row, Col) = kariatai
        End If
         Col = Col + 1
         GetLine = Right$(GetLine, Len(GetLine) - Spot)
         Spot = InStr(1, GetLine, ",")
      Wend
      If Len(GetLine) <> 0 Then
            If Mid(GetLine, 1, 1) = """" Then
                If Right(GetLine, 1) = """" Then
                    seikyuudata(Row, Col) = Mid(GetLine, 2, Len(GetLine) - 2)
                Else
                    seikyuudata(Row, Col) = Mid(GetLine, 2)
                End If
            Else
                If Right(GetLine, 1) = """" Then
                    seikyuudata(Row, Col) = Mid(GetLine, 1, Len(GetLine) - 1)
                Else
                    seikyuudata(Row, Col) = GetLine
                End If
            End If
      End If
      Row = Row + 1
   Wend
   Close #1

   'On Error GoTo ErrorTrap
   Open FilePath(1) For Input Lock Read Write As #1
   ColCount2 = 0
   RowCount = 5
    While Not EOF(1)
        Line Input #1, GetLine
        ColCount2 = ColCount2 + 1
    Wend
    Close #1
    
    Open FilePath(1) For Input Lock Read Write As #1
   ReDim shousaidata(ColCount2, RowCount)
   Position = 1
   Row = 0
   'Line Input #1, GetLine
   'Spot = InStr(1, GetLine, ",")
   While Not EOF(1)
      Line Input #1, GetLine
      Col = 1
      Spot = InStr(1, GetLine, ",")
      While Spot <> 0
         kariatai = Left$(GetLine, Spot - 1)
         If Col <> 5 Then
            If Mid(kariatai, 1, 1) = """" Then
                If Right(kariatai, 1) = """" Then
                    shousaidata(Row, Col) = Mid(kariatai, 2, Len(kariatai) - 2)
                Else
                    shousaidata(Row, Col) = Mid(kariatai, 2)
                End If
            Else
                If Right(kariatai, 1) = """" Then
                    shousaidata(Row, Col) = Mid(kariatai, 1, Len(kariatai) - 1)
                Else
                    shousaidata(Row, Col) = kariatai
                End If
            End If
        Else
            shousaidata(Row, Col) = kariatai
        End If
         Col = Col + 1
         GetLine = Right$(GetLine, Len(GetLine) - Spot)
         Spot = InStr(1, GetLine, ",")
      Wend
      If Len(GetLine) <> 0 Then
         shousaidata(Row, Col) = GetLine
      End If
      Row = Row + 1
   Wend
   Close #1

'*************未設定チェック*************************
Dim maed As String, chkmore As Integer

maed = ""
chkmore = 0
   For voipii = 0 To ColCount2 - 1
        newvoipno = ""
      If maed <> shousaidata(voipii, 3) Then
        If shousaidata(voipii, 3) <> "合計" And shousaidata(voipii, 3) <> "回線計" Then
                maed = shousaidata(voipii, 3)
                newvoipnospot = InStr(1, shousaidata(voipii, 3), "-")
                If newvoipnospot > 0 Then
                    newv1 = Mid(shousaidata(voipii, 3), 1, newvoipnospot - 1)
                    newvoipnospot2 = InStr(newvoipnospot + 1, shousaidata(voipii, 3), "-")
                    If newvoipnospot2 > 0 Then
                        newv2 = Mid(shousaidata(voipii, 3), newvoipnospot + 1, newvoipnospot2 - newvoipnospot - 1)
                        newv3 = Mid(shousaidata(voipii, 3), newvoipnospot2 + 1)
                        newvoipno = newv1 & newv2 & newv3
                    Else
                        newv2 = Mid(shousaidata(voipii, 3), newvoipnospot + 1)
                        newvoipno = newv1 & newv2
                    End If
                Else
                    newvoipno = shousaidata(voipii, 3)
                End If
                If newvoipno = "" Then
                            MsgBox "データが不正です。"
                            Exit Sub
                End If
                
                    '個人IDの確認
                    sql_voip299 = "select kojinid from kojin where (v1='" & newvoipno & "')" & _
                    "or (v2='" & newvoipno & "') or (v3='" & newvoipno & "')" & _
                    "or (v4='" & newvoipno & "') or (v5='" & newvoipno & "')" & _
                    "or (v6='" & newvoipno & "') or (v7='" & newvoipno & "')" & _
                    "or (v8='" & newvoipno & "') or (v9='" & newvoipno & "')" & _
                    "or (v10='" & newvoipno & "')"
            
                    If FcSQlGet(rs_voip299, sql_voip299, PrMsg) = False Then
                        chkmore = 1
                        Screen.MousePointer = 0
                        MsgBox "番号「" & newvoipno & "」は、登録されていません。確認してから再度実行してください。"
                    End If
        End If
    End If
Next voipii

If chkmore = 1 Then
    MsgBox "登録されていない番号を入力してから再度実行してください。処理を中止します。"
    Exit Sub
End If

    If kidou_voip = False Then
        Exit Sub
    End If
On Error GoTo 0

'*************既存のデータか確認***************
Dim rs_voip20 As Recordset
sql = "select * from seikyuu where nentsuki ='" & seikyuudata(0, 1) & "'"
If FcSQlGet_voip(rs_voip20, sql, PrMsg) = 1 Then
    rs_voip20.Close
    MsgBox "すでに取り込み済のデータです。"
    Screen.MousePointer = 0
    Exit Sub
End If

'********************************************
For voipii = 0 To ColCount - 1
    For rcrc2 = 1 To 8
        If seikyuudata(voipii, rcrc2) = Null Then
            MsgBox "データにNULLがあります。データが不正です。"
            Screen.MousePointer = 0
            Exit Sub
        End If
        If Trim(seikyuudata(voipii, rcrc2)) = "" Then
            MsgBox "データに欠損があります。データが不正です。"
            Screen.MousePointer = 0
            Exit Sub
        End If
    Next
Next
'********************************************
If no = 1 Then
    MsgBox "データに欠損は見つかりませんでした。"
    Screen.MousePointer = 0
    Exit Sub
End If

'*********トランザクション開始**************
BeginTrans
    
    '************請求********************
    If soukatsu(0, "startup", 6, "", 1, 1) = False Then
    Screen.MousePointer = 0
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
    
    err_counter = 1
    
 On Error GoTo errvoip
   
   For voipii = 0 To ColCount - 1
        sql = "INSERT INTO seikyuu VALUES('" & newtsuwaid & "'," & _
                       "'" & seikyuudata(voipii, 1) & "','" & seikyuudata(voipii, 2) & _
                       "','" & seikyuudata(voipii, 3) & "','" & seikyuudata(voipii, 4) & "','" & seikyuudata(voipii, 5) & _
                       "'," & seikyuudata(voipii, 6) & "," & seikyuudata(voipii, 7) & ",'" & seikyuudata(voipii, 8) & "')"
        db_voip.Execute sql, dbSQLPassThrough
        err_counter = err_counter + 1
        newtsuwaid = Format$(CStr(CLng(newtsuwaid) + 1), "0000000#")
    Next voipii

On Error GoTo 0

        newtsuwaid = CStr(CLng(newtsuwaid))
        If soukatsu(1, "startup", 6, newtsuwaid, 1, 1) = False Then
        Screen.MousePointer = 0
                MsgBox "設定の書き込みに失敗しました。"
                Rollback
                Exit Sub
        End If
    '************詳細********************
    If soukatsu(0, "startup", 7, "", 1, 1) = False Then
    Screen.MousePointer = 0
            MsgBox "設定の読み込みに失敗しました。"
            Rollback
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
    
    err_counter = 1
    
 On Error GoTo errvoip2
    
   For voipii = 0 To ColCount2 - 1
        sql = "INSERT INTO shousai VALUES('" & newtsuwaid & "'," & _
                       "'" & shousaidata(voipii, 1) & "','" & shousaidata(voipii, 2) & _
                       "','" & shousaidata(voipii, 3) & "','" & shousaidata(voipii, 4) & "','" & shousaidata(voipii, 5) & "')"
        db_voip.Execute sql, dbSQLPassThrough
        newtsuwaid = Format$(CStr(CLng(newtsuwaid) + 1), "00000000#")
        err_counter = err_counter + 1
    Next voipii
    
On Error GoTo 0
        newtsuwaid = CStr(CLng(newtsuwaid))
        If soukatsu(1, "startup", 7, newtsuwaid, 1, 1) = False Then
        Screen.MousePointer = 0
            Rollback
                MsgBox "設定の書き込みに失敗しました。"
                Exit Sub
        End If
    
     '************詳細********************
    If soukatsu(0, "startup", 9, "", 1, 1) = False Then
    Screen.MousePointer = 0
        Rollback
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
    
    Dim goukeimade As Integer, goukeinashi As Integer
    
    err_counter = 1
    
 On Error GoTo errvoip3
    
   For voipii = 0 To ColCount2 - 1
        newvoipno = ""
        
        If shousaidata(voipii, 3) <> "合計" And shousaidata(voipii, 3) <> "回線計" Then
                newvoipnospot = InStr(1, shousaidata(voipii, 3), "-")
                If newvoipnospot > 0 Then
                    newv1 = Mid(shousaidata(voipii, 3), 1, newvoipnospot - 1)
                    newvoipnospot2 = InStr(newvoipnospot + 1, shousaidata(voipii, 3), "-")
                    If newvoipnospot2 > 0 Then
                        newv2 = Mid(shousaidata(voipii, 3), newvoipnospot + 1, newvoipnospot2 - newvoipnospot - 1)
                        newv3 = Mid(shousaidata(voipii, 3), newvoipnospot2 + 1)
                        newvoipno = newv1 & newv2 & newv3
                    Else
                        newv2 = Mid(shousaidata(voipii, 3), newvoipnospot + 1)
                        newvoipno = newv1 & newv2
                    End If
                Else
                    newvoipno = shousaidata(voipii, 3)
                End If
                voipii = voipii + 1
                goukeinashi = 0
                For goukeimade = 0 To 10
                    If shousaidata(voipii, 3) = "合計" Or shousaidata(voipii, 3) = "回線計" Then
                        goukeinashi = 1
                        Exit For
                    End If
                    voipii = voipii + 1
                Next goukeimade
                If goukeinashi = 0 Then
                    Rollback
                    MsgBox "データが不正です。"
                    Exit Sub
                End If
        End If
        If newvoipno = "" Then
                    Rollback
                    MsgBox "データが不正です。"
                    Exit Sub
        End If
        If shousaidata(voipii, 3) = "合計" Then
            If shousaidata(voipii, 4) = "課税対象金額" Then
                voipii = voipii + 1
            End If
            If shousaidata(voipii, 4) = "非課税対象金額" Then
                voipii = voipii + 1
            End If
            If shousaidata(voipii, 4) = "消費税" Then
                newzei = shousaidata(voipii, 5)
                voipii = voipii + 1
            End If
            If shousaidata(voipii, 4) = "総合計" Then
                newsou = shousaidata(voipii, 5)
            End If
        ElseIf shousaidata(voipii, 3) = "回線計" Then
            If shousaidata(voipii, 4) = "課税対象金額" Then
                voipii = voipii + 1
            End If
            If shousaidata(voipii, 4) = "非課税対象金額" Then
                voipii = voipii + 1
            End If
            If shousaidata(voipii, 4) = "消費税" Then
                newzei = shousaidata(voipii, 5)
                voipii = voipii + 1
            End If
            If shousaidata(voipii, 4) = "回線合計" Then
                newsou = shousaidata(voipii, 5)
                'voipii = voipii + 1
            End If
        End If
        
                    '個人IDの確認
                    sql_voip2 = "select kojinid from kojin where (v1='" & newvoipno & "')" & _
                    "or (v2='" & newvoipno & "') or (v3='" & newvoipno & "')" & _
                    "or (v4='" & newvoipno & "') or (v5='" & newvoipno & "')" & _
                    "or (v6='" & newvoipno & "') or (v7='" & newvoipno & "')" & _
                    "or (v8='" & newvoipno & "') or (v9='" & newvoipno & "')" & _
                    "or (v10='" & newvoipno & "')"
            
                    If FcSQlGet(rs_voip2, sql_voip2, PrMsg) = False Then
                        Screen.MousePointer = 0
                        Rollback
                        MsgBox "番号「" & newvoipno & "」は、登録されていません。確認してから再度実行してください。"
                        Exit Sub
                    Else
                        newkojinid = rs_voip2!kojinid
                        rs_voip2.Close
                    End If
                    
                    
                    
                    sql = "INSERT INTO voip_m VALUES('" & newtsuwaid & "'," & _
                                   "'" & newvoipno & "','" & shousaidata(voipii, 1) & _
                                   "','" & shousaidata(voipii, 4) & "'," & _
                                   shousaidata(voipii, 5) & "," & newzei & "," & newsou & ",'" & _
                                   shousaidata(voipii, 2) & "','" & newkojinid & "','0')"
                    db_voip.Execute sql, dbSQLPassThrough
                    newtsuwaid = Format$(CStr(CLng(newtsuwaid) + 1), "00000000#")
                    
                    
                    
            voipii = voipii + 1
            If shousaidata(voipii, 4) <> "総合計" Then
                voipii = voipii - 1
            End If
                    
                    
                    
           err_counter = err_counter + 1
    Next voipii
    
    
On Error GoTo 0

        newtsuwaid = CStr(CLng(newtsuwaid))
        If soukatsu(1, "startup", 9, newtsuwaid, 1, 1) = False Then
        Screen.MousePointer = 0
        Rollback
                MsgBox "設定の書き込みに失敗しました。"
                Exit Sub
        End If
       
        
 CommitTrans
        
        
  
   MsgBox "読み込みを完了しました。続いて請求データの作成を開始します。"
'*********データリンク*****************

sql = "SELECT*FROM voip_m where st_st='0'"
If FcSQlGet_voip(rs_voip, sql, PrMsg) = False Then
Screen.MousePointer = 0
    MsgBox "リンクするデータはありません。"
    Exit Sub
Else
    rs_voip.MoveFirst
    Do Until rs_voip.EOF
        'ここから
        If voip_koushin("000", rs_voip!kojinid, rs_voip!voipgaku, Format(Date, "yyyymmdd")) = -1 Then
            Screen.MousePointer = 0
            Rollback
            Exit Sub
        End If
            sql = "UPDATE voip_m SET st_st = '1'" & _
                         " WHERE voipid = '" & rs_voip!voipid & "'"
            db_voip.Execute sql, dbSQLPassThrough
        
        rs_voip.MoveNext
    Loop
    rs_voip.Close
End If
db_voip.Close
'*************トランザクション終了********
'CommitTrans


    'Log
       
        sagyou_msg = "VoIP更新：VoIPの振替データを読込、更新しました。"
        log_sagyou sagyou_log_path, sagyou_msg


    Screen.MousePointer = 0
    MsgBox "請求データの作成が完了しました。"

Exit Sub

    
errvoip:
    Rollback
    Screen.MousePointer = 0
    MsgBox "請求データの作成が失敗しました。ファイル：bill.txt　行数：" & err_counter
    Exit Sub
errvoip2:
    Rollback
    Screen.MousePointer = 0
    MsgBox "詳細データの作成が失敗しました。ファイル：item.txt　行数：" & err_counter
    Exit Sub
errvoip3:
    Rollback
    Screen.MousePointer = 0
    MsgBox "個人データの作成が失敗しました。ファイル：item.txt　行数：" & err_counter
    Exit Sub

ErrorTrap3:
    'Rollback
    Screen.MousePointer = 0
    MsgBox "データの作成が失敗しました。２"
    Exit Sub



End Sub
Sub cl_data_bunkatsu(s_path As String)

Dim udtBrowseInfo As BROWSEINFO
    Dim lngIDList     As Long, newzei As Long, newsou As Long, newvoipnospot As Integer, newv1 As String, newv2 As String, newv3 As String
    Dim strPath       As String, kariatai, sql_ppv2 As String, newvoipno As String, sql_ppv299 As String
    Dim newtsuwaid As String, voipii As Integer, rs_ppv As Recordset, rs_ppv2 As Recordset, rs_ppv299 As Recordset, rs_cl As Recordset
    Dim rirekirs55 As Recordset, rirekirs555 As Recordset
    Dim kakunin_atai(4) As Long, jitsu_kaki(5) As Double
    Dim FilePathppv(3), seikyuudata_ppv()
    
    Dim s_yamanashi_filepath As String
    Dim s_katsunuma_filepath As String
    
    Dim s_chk_kikan As String

FilePathppv(2) = s_path

    If Right(FilePathppv(2), 1) = "\" Then
        FilePathppv(2) = Mid(FilePathppv(2), 1, Len(FilePathppv(2)) - 1)
    End If
    Dim k_t As Integer, k_n As Long
    
    If Format(Date, "mm") = "01" Then
        k_t = 12
        k_n = CLng(Format(Date, "yyyy")) - 1
    Else
        k_t = CInt(Format(Date, "mm")) - 1
        k_n = CLng(Format(Date, "yyyy"))
    End If
    
    s_chk_kikan = Format(k_n, "0000") & Format(k_t, "00")
    
    '*******元の読込ファイル（山梨CTAV→勝沼CATV）***********************
   FilePathppv(0) = FilePathppv(2) & "\山梨から勝沼へ\cl_" & Format(k_n, "0000") & "_" & Format(k_t, "00") & "_both.txt"
   
   s_yamanashi_filepath = FilePathppv(2) & "\山梨CATV用データ\cl_" & Format(k_n, "0000") & "_" & Format(k_t, "00") & ".txt"
   s_katsunuma_filepath = FilePathppv(2) & "\勝沼CATV用データ\cl_" & Format(k_n, "0000") & "_" & Format(k_t, "00") & "_katsunuma.txt"
   
  ' FilePathppv(0) = FilePathppv(2) & "\請求データサンプル" & ".txt"
   
   If Dir(FilePathppv(0)) = "" Then
        MsgBox "更新したいファイルがありません。再度確認してください。"
        Exit Sub
   End If
   
    If MsgBox("読み込みと登録データの比較チェックを開始し、データを分割します。よろしいですか？", vbYesNo) = vbNo Then
        Exit Sub
    End If


   
    'FDのコピーをHDDに実施
    Dim sakifile9 As String, sakifailego As String
    sakifile9 = sopa & "fdback\cl\yamanashi\"
    If Dir$(sakifile9, vbDirectory) = "" Then
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder (sakifile9)
        Set fso = Nothing
    End If
    If Right(sakifile9, 1) <> "\" Then
        sakifile9 = sakifile9 & "\"
    End If
    
     '*******保存ファイル（勝沼CATV）***********************
    sakifailego = sakifile9 & "cl_" & Format(k_n, "0000") & "_" & Format(k_t, "00") & "_both.txt"
    
    
    Dim ColCount As Long
    Dim RowCount As Long
    Dim GetLine As String
    Dim Position As Integer
    Dim Row As Integer, Col As Integer, Spot As Integer
  
   If Dir(sakifailego) <> "" Then
      If MsgBox("更新したいファイルは、すでに更新した疑いがあります。よろしいですか？", vbYesNo) = vbNo Then
            Exit Sub
        End If
        
   End If
    
    Screen.MousePointer = 11
    
    'コピー
    On Error GoTo fdcopy
    FileCopy FilePathppv(0), sakifailego
    On Error GoTo 0
    
     DoEvents
    
    'データ読込
    On Error GoTo Errppv1
  ' Open FilePathppv(0) For Input Lock Read Write As #1
   Open sakifailego For Input Lock Read Write As #1
   ColCount = 0
   RowCount = 65
    While Not EOF(1)
        Line Input #1, GetLine
        ColCount = ColCount + 1
    Wend
    Close #1
    
   ' Open FilePathppv(0) For Input Lock Read Write As #1
    Open sakifailego For Input Lock Read Write As #1
   ReDim seikyuudata_ppv(ColCount - 1, RowCount)
   Position = 1
   Row = 0
   While Not EOF(1)
      Line Input #1, GetLine
      Col = 1
      Spot = InStr(1, GetLine, ",")
      While Spot <> 0
         kariatai = Left$(GetLine, Spot - 1)
         If Col <> 12 Then
            If Mid(kariatai, 1, 1) = """" Then
                If Right(kariatai, 1) = """" Then
                    seikyuudata_ppv(Row, Col) = Mid(kariatai, 2, Len(kariatai) - 2)
                Else
                    seikyuudata_ppv(Row, Col) = Mid(kariatai, 2)
                End If
            Else
                If Right(kariatai, 1) = """" Then
                    seikyuudata_ppv(Row, Col) = Mid(kariatai, 1, Len(kariatai) - 1)
                Else
                    seikyuudata_ppv(Row, Col) = kariatai
                End If
            End If
        Else
            seikyuudata_ppv(Row, Col) = kariatai
        End If
         Col = Col + 1
         GetLine = Right$(GetLine, Len(GetLine) - Spot)
         Spot = InStr(1, GetLine, ",")
      Wend
      If Len(GetLine) <> 0 Then
            If Mid(GetLine, 1, 1) = """" Then
                If Right(GetLine, 1) = """" Then
                    seikyuudata_ppv(Row, Col) = Mid(GetLine, 2, Len(GetLine) - 2)
                Else
                    seikyuudata_ppv(Row, Col) = Mid(GetLine, 2)
                End If
            Else
                If Right(GetLine, 1) = """" Then
                    seikyuudata_ppv(Row, Col) = Mid(GetLine, 1, Len(GetLine) - 1)
                Else
                    seikyuudata_ppv(Row, Col) = GetLine
                End If
            End If
      End If
      Row = Row + 1
   Wend
   Close #1
   
    
   
    
    Dim m_data() As String
    Dim m_count As Integer
    
    
    '種類を取得

       
                
        sql = "select * from cl"
        
        If FcSQlGet(rs, sql, PrMsg) = False Then
            Screen.MousePointer = 0
                MsgBox "CLの番号がありません。"
                    
                    If Dir(sakifailego) <> "" Then
                          Kill sakifailego
                     End If
        
                    Exit Sub
        Else
            rs.MoveLast
            m_count = rs.RecordCount
           ReDim m_data(2, m_count)
            rs.MoveFirst
            For i = 0 To m_count - 1
               m_data(0, i) = rs!kojinid
                m_data(1, i) = rs!tel
               rs.MoveNext
            Next i
            rs.Close
        End If
          
    
    
     'データ判別
    
    Dim y_data() As String
    Dim y_count As Integer
    Dim k_data() As String
    Dim k_count As Integer
    
    y_count = 0
    k_count = 0
    
        '比較
        Dim s_tel As String, isi As Integer, s_hit As Integer, isi2 As Integer
        
        For i = 0 To ColCount - 1
            If seikyuudata_ppv(i, 1) = "2" Then
                If seikyuudata_ppv(i, 2) = "5100" Then
                    s_tel = Trim(seikyuudata_ppv(i, 5))
                    If s_tel <> "99999999999999999999999" Then
                        s_hit = 0
                        For isi = 0 To m_count - 1
                            If s_tel = m_data(1, isi) Then
                                s_hit = 1
                                Exit For
                            End If
                        Next
                        If s_hit = 1 Then
                            'ヒットしたので勝沼用
                            'If seikyuudata_ppv(i, 8) <> "CIPE1" Then
                                k_count = k_count + 1
                                ReDim Preserve k_data(65, k_count)
                                For isi2 = 0 To 65
                                    k_data(isi2, k_count - 1) = seikyuudata_ppv(i, isi2)
                                Next
                            'Else
                            
                           ' End If
                        Else
                            'ヒットしないので山梨用
                            y_count = y_count + 1
                            ReDim Preserve y_data(65, y_count)
                            For isi2 = 0 To 65
                                y_data(isi2, y_count - 1) = seikyuudata_ppv(i, isi2)
                            Next
                        End If
                     End If
                End If
            ElseIf seikyuudata_ppv(i, 1) = "1" Then
                If seikyuudata_ppv(i, 7) <> s_chk_kikan Then
                    Screen.MousePointer = 0
   
                    MsgBox "分割しようとしているファイルは、期間が違います。確認してください。"
                    Exit Sub
                End If
            End If
        Next
    
    
    'データを2つに保存
    '初期化
     If Dir(s_yamanashi_filepath) <> "" Then
         Kill s_yamanashi_filepath
    End If
    If Dir(s_katsunuma_filepath) <> "" Then
         Kill s_katsunuma_filepath
    End If
    
    '作成
    '山梨用データコピー
    Dim writecounter As Integer
    Open s_yamanashi_filepath For Output Access Write As 1
    For writecounter = 0 To y_count - 1
            Write #1, y_data(1, writecounter), y_data(2, writecounter), y_data(3, writecounter), y_data(4, writecounter), y_data(5, writecounter), y_data(6, writecounter), y_data(7, writecounter), y_data(8, writecounter), y_data(9, writecounter), y_data(10, writecounter), y_data(11, writecounter), y_data(12, writecounter), y_data(13, writecounter), y_data(14, writecounter), y_data(15, writecounter), y_data(16, writecounter), y_data(17, writecounter), y_data(18, writecounter)
    Next
    Close #1

  '勝沼用データコピー
    Open s_katsunuma_filepath For Output Access Write As 2
    For writecounter = 0 To k_count - 1
            Write #2, k_data(1, writecounter), k_data(2, writecounter), k_data(3, writecounter), k_data(4, writecounter), k_data(5, writecounter), k_data(6, writecounter), k_data(7, writecounter), k_data(8, writecounter), k_data(9, writecounter), k_data(10, writecounter), k_data(11, writecounter), k_data(12, writecounter), k_data(13, writecounter), k_data(14, writecounter), k_data(15, writecounter), k_data(16, writecounter), k_data(17, writecounter), k_data(18, writecounter)
    Next
    Close #2
   
 
    DoEvents
 
    'ファイル削除
    
      If Dir(FilePathppv(0)) <> "" Then
             Kill FilePathppv(0)
        End If
    
    
     Screen.MousePointer = 0
     
     MsgBox "山梨CATVからの請求ファイルを分割しました。"
     
  Exit Sub
  
fdcopy:
   Screen.MousePointer = 0
   
    MsgBox "コピーに失敗しました。"
    Exit Sub
Errppv1:
    Screen.MousePointer = 0
    If Dir(sakifailego) <> "" Then
        Kill sakifailego
   End If
   Screen.MousePointer = 0
    MsgBox "ケーブルラインデータの読み込みデータの書き込みに失敗しました。"
    Exit Sub

End Sub
Sub cl_data_bunkatsu2(s_path As String)

Dim udtBrowseInfo As BROWSEINFO
    Dim lngIDList     As Long, newzei As Long, newsou As Long, newvoipnospot As Integer, newv1 As String, newv2 As String, newv3 As String
    Dim strPath       As String, kariatai, sql_ppv2 As String, newvoipno As String, sql_ppv299 As String
    Dim newtsuwaid As String, voipii As Integer, rs_ppv As Recordset, rs_ppv2 As Recordset, rs_ppv299 As Recordset, rs_cl As Recordset
    Dim rirekirs55 As Recordset, rirekirs555 As Recordset
    Dim kakunin_atai(4) As Long, jitsu_kaki(5) As Double
    Dim FilePathppv(3), seikyuudata_ppv()
    
    Dim s_yamanashi_filepath As String
    Dim s_katsunuma_filepath As String
    

FilePathppv(2) = s_path

    If Right(FilePathppv(2), 1) = "\" Then
        FilePathppv(2) = Mid(FilePathppv(2), 1, Len(FilePathppv(2)) - 1)
    End If
    Dim k_t As Integer, k_n As Long
    
    If Format(Date, "mm") = "01" Then
        k_t = 12
        k_n = CLng(Format(Date, "yyyy")) - 1
    Else
        k_t = CInt(Format(Date, "mm")) - 1
        k_n = CLng(Format(Date, "yyyy"))
    End If
    
    '*******元の読込ファイル（山梨CTAV→勝沼CATV）***********************
   FilePathppv(0) = FilePathppv(2) & "\山梨から勝沼へ\cl_" & Format(k_n, "0000") & "_" & Format(k_t, "00") & "_both.txt"
   
   s_yamanashi_filepath = FilePathppv(2) & "\山梨CATV用データ\cl_" & Format(k_n, "0000") & "_" & Format(k_t, "00") & ".txt"
   s_katsunuma_filepath = FilePathppv(2) & "\勝沼CATV用データ\cl_" & Format(k_n, "0000") & "_" & Format(k_t, "00") & "_katsunuma.txt"
   
  ' FilePathppv(0) = FilePathppv(2) & "\請求データサンプル" & ".txt"
   
   If Dir(FilePathppv(0)) = "" Then
        MsgBox "更新したいファイルがありません。再度確認してください。"
        Exit Sub
   End If
   
    If MsgBox("読み込みと登録データの比較チェックを開始し、データを分割します。よろしいですか？", vbYesNo) = vbNo Then
        Exit Sub
    End If


   
    'FDのコピーをHDDに実施
    Dim sakifile9 As String, sakifailego As String
    sakifile9 = sopa & "fdback\cl\yamanashi\"
    If Dir$(sakifile9, vbDirectory) = "" Then
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder (sakifile9)
        Set fso = Nothing
    End If
    If Right(sakifile9, 1) <> "\" Then
        sakifile9 = sakifile9 & "\"
    End If
    
     '*******保存ファイル（勝沼CATV）***********************
    sakifailego = sakifile9 & "cl_" & Format(k_n, "0000") & "_" & Format(k_t, "00") & "_both.txt"
    
    
    Dim ColCount As Long
    Dim RowCount As Long
    Dim GetLine As String
    Dim Position As Integer
    Dim Row As Integer, Col As Integer, Spot As Integer
  
   If Dir(sakifailego) <> "" Then
      If MsgBox("更新したいファイルは、すでに更新した疑いがあります。よろしいですか？", vbYesNo) = vbNo Then
            Exit Sub
        End If
        
   End If
    
    Screen.MousePointer = 11
    
    'コピー
    On Error GoTo fdcopy
    FileCopy FilePathppv(0), sakifailego
    On Error GoTo 0
    
     DoEvents
    
    'データ読込
    On Error GoTo Errppv1
  ' Open FilePathppv(0) For Input Lock Read Write As #1
   Open sakifailego For Input Lock Read Write As #1
   ColCount = 0
   RowCount = 65
    While Not EOF(1)
        Line Input #1, GetLine
        ColCount = ColCount + 1
    Wend
    Close #1
    
   ' Open FilePathppv(0) For Input Lock Read Write As #1
    Open sakifailego For Input Lock Read Write As #1
   ReDim seikyuudata_ppv(ColCount - 1, RowCount)
   Position = 1
   Row = 0
   While Not EOF(1)
      Line Input #1, GetLine
      Col = 1
      Spot = InStr(1, GetLine, ",")
      While Spot <> 0
         kariatai = Left$(GetLine, Spot - 1)
         If Col <> 12 Then
            If Mid(kariatai, 1, 1) = """" Then
                If Right(kariatai, 1) = """" Then
                    seikyuudata_ppv(Row, Col) = Mid(kariatai, 2, Len(kariatai) - 2)
                Else
                    seikyuudata_ppv(Row, Col) = Mid(kariatai, 2)
                End If
            Else
                If Right(kariatai, 1) = """" Then
                    seikyuudata_ppv(Row, Col) = Mid(kariatai, 1, Len(kariatai) - 1)
                Else
                    seikyuudata_ppv(Row, Col) = kariatai
                End If
            End If
        Else
            seikyuudata_ppv(Row, Col) = kariatai
        End If
         Col = Col + 1
         GetLine = Right$(GetLine, Len(GetLine) - Spot)
         Spot = InStr(1, GetLine, ",")
      Wend
      If Len(GetLine) <> 0 Then
            If Mid(GetLine, 1, 1) = """" Then
                If Right(GetLine, 1) = """" Then
                    seikyuudata_ppv(Row, Col) = Mid(GetLine, 2, Len(GetLine) - 2)
                Else
                    seikyuudata_ppv(Row, Col) = Mid(GetLine, 2)
                End If
            Else
                If Right(GetLine, 1) = """" Then
                    seikyuudata_ppv(Row, Col) = Mid(GetLine, 1, Len(GetLine) - 1)
                Else
                    seikyuudata_ppv(Row, Col) = GetLine
                End If
            End If
      End If
      Row = Row + 1
   Wend
   Close #1
   
    
   
    
    Dim m_data() As String
    Dim m_count As Integer
    
    
    '種類を取得

        sql = "SELECT keiyaku.keiyakuid, keiyaku.kojinid, keiyaku.shurui," & _
                "keiyaku.kaishi, keiyaku.kadou, keiyaku.kuchisuu, keiyaku.kingaku,keiyaku.shime" & _
                " FROM kojin RIGHT JOIN keiyaku ON kojin.kojinid = keiyaku.kojinid" & _
                " WHERE (keiyaku.shurui ='33' or keiyaku.shurui ='35' or keiyaku.shurui ='36' or keiyaku.shurui ='37' or keiyaku.shurui ='38'" & _
                " or keiyaku.shurui ='39' or keiyaku.shurui ='40' or keiyaku.shurui ='41' or keiyaku.shurui ='42' or keiyaku.shurui ='43')" & _
                " and kadou='0' and shime='0'" & _
                " order by kojin.kojinid,keiyaku.keiyakuid"
                
        sql = "SELECT keiyaku.keiyakuid, keiyaku.kojinid, keiyaku.shurui," & _
                "keiyaku.kaishi, keiyaku.kadou, keiyaku.kuchisuu, keiyaku.kingaku,keiyaku.shime" & _
                " FROM kojin RIGHT JOIN keiyaku ON kojin.kojinid = keiyaku.kojinid" & _
                " WHERE (keiyaku.shurui ='33' or keiyaku.shurui ='35' or keiyaku.shurui ='36' or keiyaku.shurui ='37' or keiyaku.shurui ='38'" & _
                " or keiyaku.shurui ='39' or keiyaku.shurui ='40' or keiyaku.shurui ='41' or keiyaku.shurui ='42' or keiyaku.shurui ='43')" & _
                " and kadou='0'" & _
                " order by kojin.kojinid,keiyaku.keiyakuid"
        
        If FcSQlGet(rs, sql, PrMsg) = False Then
            Screen.MousePointer = 0
                MsgBox "CLの主契約がありません。"
                    
                    If Dir(sakifailego) <> "" Then
                          Kill sakifailego
                     End If
        
                    Exit Sub
        Else
            rs.MoveLast
            m_count = rs.RecordCount
           ReDim m_data(10, m_count)
            rs.MoveFirst
            For i = 0 To m_count - 1
               m_data(0, i) = rs!keiyakuid
                m_data(1, i) = rs!shurui
                m_data(2, i) = rs!kuchisuu
                m_data(3, i) = rs!kingaku
                m_data(4, i) = rs!kojinid
               rs.MoveNext
            Next i
            rs.Close
        End If
          
    
         For i = 0 To m_count - 1
         
                 sql = "select * from cl where kojinid = '" & m_data(4, i) & "'"
                
                Set rs_cl = db.OpenRecordset(sql, dbOpenDynaset, dbReadOnly)
                If FcSQlGet(rs_cl, sql, PrMsg) = False Then
                    Screen.MousePointer = 0
                    MsgBox "CLの主契約はありますが、CLの番号が登録されていません。個人ID：" & m_data(4, i)
                    
                    If Dir(sakifailego) <> "" Then
                          Kill sakifailego
                     End If
        
                    Exit Sub
                Else
                    
                     m_data(5, i) = rs_cl!tel
                    
                    rs_cl.Close
                End If
         Next
    
     'データ判別
    
    Dim y_data() As String
    Dim y_count As Integer
    Dim k_data() As String
    Dim k_count As Integer
    
    y_count = 0
    k_count = 0
    
        '比較
        Dim s_tel As String, isi As Integer, s_hit As Integer, isi2 As Integer
        
        For i = 0 To ColCount - 1
            If seikyuudata_ppv(i, 1) = "2" Then
                If seikyuudata_ppv(i, 2) = "5100" Then
                    s_tel = Trim(seikyuudata_ppv(i, 5))
                    If s_tel <> "99999999999999999999999" Then
                        s_hit = 0
                        For isi = 0 To m_count - 1
                            If s_tel = m_data(5, isi) Then
                                s_hit = 1
                                Exit For
                            End If
                        Next
                        If s_hit = 1 Then
                            'ヒットしたので勝沼用
                            'If seikyuudata_ppv(i, 8) <> "CIPE1" Then
                                k_count = k_count + 1
                                ReDim Preserve k_data(65, k_count)
                                For isi2 = 0 To 65
                                    k_data(isi2, k_count - 1) = seikyuudata_ppv(i, isi2)
                                Next
                            'Else
                            
                           ' End If
                        Else
                            'ヒットしないので山梨用
                            y_count = y_count + 1
                            ReDim Preserve y_data(65, y_count)
                            For isi2 = 0 To 65
                                y_data(isi2, y_count - 1) = seikyuudata_ppv(i, isi2)
                            Next
                        End If
                     End If
                End If
            End If
        Next
    
    
    'データを2つに保存
    '初期化
     If Dir(s_yamanashi_filepath) <> "" Then
         Kill s_yamanashi_filepath
    End If
    If Dir(s_katsunuma_filepath) <> "" Then
         Kill s_katsunuma_filepath
    End If
    
    '作成
    '山梨用データコピー
    Dim writecounter As Integer
    Open s_yamanashi_filepath For Output Access Write As 1
    For writecounter = 0 To y_count - 1
            Write #1, y_data(1, writecounter), y_data(2, writecounter), y_data(3, writecounter), y_data(4, writecounter), y_data(5, writecounter), y_data(6, writecounter), y_data(7, writecounter), y_data(8, writecounter), y_data(9, writecounter), y_data(10, writecounter), y_data(11, writecounter), y_data(12, writecounter), y_data(13, writecounter), y_data(14, writecounter), y_data(15, writecounter), y_data(16, writecounter), y_data(17, writecounter), y_data(18, writecounter)
    Next
    Close #1

  '勝沼用データコピー
    Open s_katsunuma_filepath For Output Access Write As 2
    For writecounter = 0 To k_count - 1
            Write #2, k_data(1, writecounter), k_data(2, writecounter), k_data(3, writecounter), k_data(4, writecounter), k_data(5, writecounter), k_data(6, writecounter), k_data(7, writecounter), k_data(8, writecounter), k_data(9, writecounter), k_data(10, writecounter), k_data(11, writecounter), k_data(12, writecounter), k_data(13, writecounter), k_data(14, writecounter), k_data(15, writecounter), k_data(16, writecounter), k_data(17, writecounter), k_data(18, writecounter)
    Next
    Close #2
   
 
    DoEvents
 
    'ファイル削除
    
      If Dir(FilePathppv(0)) <> "" Then
             Kill FilePathppv(0)
        End If
    
    
     Screen.MousePointer = 0
     
     MsgBox "山梨CATVからの請求ファイルを分割しました。"
     
  Exit Sub
  
fdcopy:
   Screen.MousePointer = 0
   
    MsgBox "コピーに失敗しました。"
    Exit Sub
Errppv1:
    Screen.MousePointer = 0
    If Dir(sakifailego) <> "" Then
        Kill sakifailego
   End If
   Screen.MousePointer = 0
    MsgBox "ケーブルラインデータの読み込みデータの書き込みに失敗しました。"
    Exit Sub

End Sub
Sub cl_data_convert(s_path As String, Optional s_err As String = "")
    '***********データを指定してコンバートする**************************
   ' このプロシージャはセミコロンで区切られた
   ' ファイルからデータを配列内にロードします。
   Dim GetLine As Variant
   Dim Spot As Integer, Position As Integer, newkojinid As String
   Dim Row As Integer, Col As Integer, newvoipnospot2 As Integer, newnentsuki As String
   Dim RowCount As Integer, ColCount As Integer, ColCount2 As Integer, ColCount3 As Integer
    Dim udtBrowseInfo As BROWSEINFO
    Dim lngIDList     As Long, newzei As Long, newsou As Long, newvoipnospot As Integer, newv1 As String, newv2 As String, newv3 As String
    Dim strPath       As String, kariatai, sql_ppv2 As String, newvoipno As String, sql_ppv299 As String
    Dim newtsuwaid As String, voipii As Integer, rs_ppv As Recordset, rs_ppv2 As Recordset, rs_ppv299 As Recordset, rs_yotei3 As Recordset, rs_yotei4 As Recordset
    Dim rirekirs55 As Recordset, rirekirs555 As Recordset
    Dim kakunin_atai(4) As Long, jitsu_kaki(5) As Double
    Dim FilePathppv(3), seikyuudata_ppv()
    
Dim s_nentsuki As String

FilePathppv(2) = s_path

    If Right(FilePathppv(2), 1) = "\" Then
        FilePathppv(2) = Mid(FilePathppv(2), 1, Len(FilePathppv(2)) - 1)
    End If
    Dim k_t As Integer, k_n As Long
    
    If Format(Date, "mm") = "01" Then
        k_t = 12
        k_n = CLng(Format(Date, "yyyy")) - 1
    Else
        k_t = CInt(Format(Date, "mm")) - 1
        k_n = CLng(Format(Date, "yyyy"))
    End If
    
    '*******元の読込ファイル（山梨CTAV→勝沼CATV）***********************
    s_nentsuki = Format(k_n, "0000") & Format(k_t, "00")
    
    If s_err = "" Then
        FilePathppv(0) = FilePathppv(2) & "\勝沼CATV用データ\cl_" & Format(k_n, "0000") & "_" & Format(k_t, "00") & "_katsunuma.txt"
    Else
        FilePathppv(0) = FilePathppv(2) & "\エラーデータ\cl_" & Format(k_n, "0000") & "_" & Format(k_t, "00") & "_katsunuma_err.txt"
    End If
    
   If Dir(FilePathppv(0)) = "" Then
        MsgBox "更新したいファイルがありません。再度確認してください。"
        Exit Sub
   End If
   
   If s_err = "" Then
        If MsgBox("読み込みと登録データのチェックを開始します。よろしいですか？", vbYesNo) = vbNo Then
            Exit Sub
        End If
    Else
        If MsgBox("再度戻ってきたエラーデータの読み込みと登録データのチェックを開始します。よろしいですか？", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
Screen.MousePointer = 11
   
    'FDのコピーをHDDに実施
    Dim sakifile9 As String, sakifailego As String
    sakifile9 = sopa & "fdback\cl\katsunuma\"
    If Dir$(sakifile9, vbDirectory) = "" Then
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder (sakifile9)
        Set fso = Nothing
    End If
    If Right(sakifile9, 1) <> "\" Then
        sakifile9 = sakifile9 & "\"
    End If
    
     '*******保存ファイル（勝沼CATV）***********************
     If s_err = "" Then
            sakifailego = sakifile9 & "cl_" & Format(k_n, "0000") & "_" & Format(k_t, "00") & "_katsunuma.txt"
    Else
             sakifailego = sakifile9 & "cl_" & Format(k_n, "0000") & "_" & Format(k_t, "00") & "_katsunuma_err.txt"
    End If
    
    
   If Dir(sakifailego) <> "" Then
        MsgBox "更新したいファイルは、すでに更新した疑いがあります。再度確認してください。"
        Screen.MousePointer = 0
        Exit Sub
   End If
    
    
    'コピー
    FileCopy FilePathppv(0), sakifailego
   


    '読込中
On Error GoTo Errppv1
   Open FilePathppv(0) For Input Lock Read Write As #1
   ColCount = 0
   RowCount = 65
    While Not EOF(1)
        Line Input #1, GetLine
        ColCount = ColCount + 1
    Wend
    Close #1
    
    Open FilePathppv(0) For Input Lock Read Write As #1
   ReDim seikyuudata_ppv(ColCount - 1, RowCount)
   Position = 1
   Row = 0
   While Not EOF(1)
      Line Input #1, GetLine
      Col = 1
      Spot = InStr(1, GetLine, ",")
      While Spot <> 0
         kariatai = Left$(GetLine, Spot - 1)
         If Col <> 12 Then
            If Mid(kariatai, 1, 1) = """" Then
                If Right(kariatai, 1) = """" Then
                    seikyuudata_ppv(Row, Col) = Mid(kariatai, 2, Len(kariatai) - 2)
                Else
                    seikyuudata_ppv(Row, Col) = Mid(kariatai, 2)
                End If
            Else
                If Right(kariatai, 1) = """" Then
                    seikyuudata_ppv(Row, Col) = Mid(kariatai, 1, Len(kariatai) - 1)
                Else
                    seikyuudata_ppv(Row, Col) = kariatai
                End If
            End If
        Else
            seikyuudata_ppv(Row, Col) = kariatai
        End If
         Col = Col + 1
         GetLine = Right$(GetLine, Len(GetLine) - Spot)
         Spot = InStr(1, GetLine, ",")
      Wend
      If Len(GetLine) <> 0 Then
            If Mid(GetLine, 1, 1) = """" Then
                If Right(GetLine, 1) = """" Then
                    seikyuudata_ppv(Row, Col) = Mid(GetLine, 2, Len(GetLine) - 2)
                Else
                    seikyuudata_ppv(Row, Col) = Mid(GetLine, 2)
                End If
            Else
                If Right(GetLine, 1) = """" Then
                    seikyuudata_ppv(Row, Col) = Mid(GetLine, 1, Len(GetLine) - 1)
                Else
                    seikyuudata_ppv(Row, Col) = GetLine
                End If
            End If
      End If
      Row = Row + 1
   Wend
   Close #1
   
   
   
'***************既存データチェック****************************************
Dim chk_sumi_data()
Dim chk_sumi_count As Integer



    Dim rs_yotei2  As Recordset, s_shiteikojinid As String, s_kensakustr As String, s_koumoku As String, s_err_msg As String, s_hikaku_str2 As String, s_hikaku_str3 As String
 Dim chkumu As Integer, iii As Integer
 Dim s_bangou As String
 
   chkumu = 0
   chk_sumi_count = 0
   
    For voipii = 0 To ColCount - 1
        If seikyuudata_ppv(voipii, 1) = "2" And seikyuudata_ppv(voipii, 2) = "5100" Then
               If seikyuudata_ppv(voipii, 5) <> "99999999999999999999999" Then
                        s_shiteikojinid = ""
                        s_hikaku_str2 = ""
                        s_koumoku = ""
                        
                        s_bangou = seikyuudata_ppv(voipii, 5)
                        
                        'CLの電話番号 *****************************************************************************
                        
                        sql_ppv299 = "select * from cl where tel='" & seikyuudata_ppv(voipii, 5) & "'"
                
                        If FcSQlGet(rs_ppv299, sql_ppv299, PrMsg) = False Then
                            Screen.MousePointer = 0
                            MsgBox "CL番号「" & seikyuudata_ppv(voipii, 5) & "」は、登録されていません。確認してから再度実行してください。"
                            chkumu = 1
                            Exit For
                        Else
                            If rs_ppv299.RecordCount <> 1 Then
                                Screen.MousePointer = 0
                                MsgBox "CL番号「" & seikyuudata_ppv(voipii, 5) & "」は、複数登録されています。確認してから再度実行してください。"
                                rs_ppv299.MoveFirst
                                Do Until rs_ppv299.EOF
                                    MsgBox "ID「" & rs_ppv299!kojinid & "」に登録されています。チェックしてください。"
                                    rs_ppv299.MoveNext
                                Loop
                                rs_ppv299.Close
                                chkumu = 1
                                Exit For
                            Else
                                seikyuudata_ppv(voipii, 19) = rs_ppv299!kojinid
                                 s_shiteikojinid = rs_ppv299!kojinid
                                 
                                 
                                 If s_shiteikojinid = "00618" Then
                                     'MsgBox "00618"
                                 End If
                                 
                                 
                                 Dim s_aruno As Integer, i3i As Integer, n_index As Integer
                                 s_aruno = 0
                                For i3i = 0 To chk_sumi_count - 1
                                    n_index = i3i
                                    If chk_sumi_data(0, i3i) = s_shiteikojinid Then
                                        If chk_sumi_data(1, i3i) = seikyuudata_ppv(voipii, 5) Then
                                            s_aruno = 1
                                            Exit For
                                        End If
                                    End If
                                Next
                                If s_aruno = 0 Then
                                        ReDim Preserve chk_sumi_data(3, chk_sumi_count + 1)
                                        
                                        chk_sumi_data(0, chk_sumi_count) = s_shiteikojinid
                                         chk_sumi_data(1, chk_sumi_count) = seikyuudata_ppv(voipii, 5)
                                         chk_sumi_data(2, chk_sumi_count) = "1"
                                         
                                         chk_sumi_count = chk_sumi_count + 1
                                 
                                 End If
                                  
                                  
                            End If
                            
                        End If
                        '******************************************************************************************
                        
                        If seikyuudata_ppv(voipii, 8) <> "CIPE1" Then
                                '課金課の判断
                                If s_shiteikojinid <> "" Then
                                    s_kensakustr = ""
                                    Select Case seikyuudata_ppv(voipii, 8)
                                        Case "CIP5K"
                                            s_kensakustr = "33" '基本料金
                                            s_koumoku = "基本料金"
                                        Case "CIP5M"
                                            s_kensakustr = "35" '番号表示
                                            s_koumoku = "番号表示"
                                        Case "CIP5O"
                                            s_kensakustr = "36" '番号リクエスト
                                            s_koumoku = "番号リクエスト"
                                        Case "CIP5L"
                                            s_kensakustr = "37" 'キャッチ
                                            s_koumoku = "キャッチ"
                                        Case "CIP5P"
                                            s_kensakustr = "38" '着信お断り
                                            s_koumoku = "着信お断り"
                                        Case "CIP5N"
                                            s_kensakustr = "39" '着信転送
                                            s_koumoku = "着信転送"
                                        Case "CIP5Q"
                                            s_kensakustr = "40" 'パック１
                                            s_koumoku = "パック１"
                                        Case "CIP5R"
                                            s_kensakustr = "41" 'パック２
                                            s_koumoku = "パック２"
                                        Case "CIP80"
                                            s_kensakustr = "42" 'パック３
                                            s_koumoku = "パック３"
                                        Case "CIP81"
                                            s_kensakustr = "43" 'パック４
                                            s_koumoku = "パック４"
                                    End Select
                                    
                                    '対応契約があり****************************************************************
                                    If s_kensakustr <> "" Then
                                    
                                        Dim n_kaisuu As Integer, ii3 As Integer
                                        n_kaisuu = 0
                                             For ii3 = 0 To chk_sumi_count - 1
                                                If chk_sumi_data(0, ii3) = s_shiteikojinid Then
                                                    If chk_sumi_data(1, ii3) = s_bangou Then
                                                            n_kaisuu = n_kaisuu + 1
                                                    End If
                                                End If
                                            Next
                                    
                                    
                                        'sql = "SELECT * FROM keiyaku  WHERE kojinid ='" & s_shiteikojinid & "' and sadou ='0' and shime ='0' and shurui='" & s_kensakustr & "'"
                                         sql = "SELECT * FROM keiyaku  WHERE kojinid ='" & s_shiteikojinid & "' and kuchisuu >='" & CStr(n_kaisuu) & "' and kadou ='0' and shime ='0' and shurui='" & s_kensakustr & "'"
                                         
                                        If FcSQlGet(rs_yotei2, sql, PrMsg) = True Then
                                             seikyuudata_ppv(voipii, 20) = "0" '課金なし
                                        Else
                                            '契約がないとき（契約月ではEZManagerには主契約がないが、請求には＋とーがあがってくる）
                                            Dim s_mainasuari As Integer, s_hikaku_str As String
                                            s_mainasuari = 0
                                            s_hikaku_str = ""
                                            Select Case seikyuudata_ppv(voipii, 8)
                                                Case "CIP5K"
                                                    s_hikaku_str = "CIP22" '基本料金
                                                Case "CIP5M"
                                                    s_hikaku_str = "CIP24" '番号表示
                                                Case "CIP5O"
                                                    s_hikaku_str = "CIP26" '番号リクエスト
                                                Case "CIP5L"
                                                    s_hikaku_str = "CIP23" 'キャッチ
                                                Case "CIP5P"
                                                    s_hikaku_str = "CIP27" '着信お断り
                                                Case "CIP5N"
                                                    s_hikaku_str = "CIP25" '着信転送
                                                Case "CIP5Q"
                                                    s_hikaku_str = "CIP28" 'パック１
                                                Case "CIP5R"
                                                    s_hikaku_str = "CIP29" 'パック２
                                                Case "CIP80"
                                                    s_hikaku_str = "CIP82" 'パック３
                                                Case "CIP81"
                                                    s_hikaku_str = "CIP83" 'パック４
                                            End Select
                                            
                                            For iii = 0 To ColCount - 1
                                                If seikyuudata_ppv(iii, 19) = s_shiteikojinid Then
                                                    If seikyuudata_ppv(iii, 8) = s_hikaku_str Then
                                                        s_mainasuari = 1
                                                        Exit For
                                                    End If
                                                End If
                                            Next iii
                                            
                                            If s_mainasuari = 0 Then
                                                'seikyuudata_ppv(voipii, 20) = "1" '課金する
                                                'CLにあってEZにない
                                                If frmkoushin.chkkaihi.Value = 1 Then
                                                        s_err_msg = "顧客ID：" & s_shiteikojinid & "　に、主契約「" & s_koumoku & "」がありませんが、今回は請求を回避します。"
                                                        
                                                        MsgBox s_err_msg
                                                        
                                                        seikyuudata_ppv(voipii, 20) = "0" '課金なし
                                                Else
                                                        s_err_msg = "顧客ID：" & s_shiteikojinid & "　に、主契約「" & s_koumoku & "」がありません。契約内容を確認し、追加後、再度実行してください。"
                                                        
                                                        MsgBox s_err_msg
                                                    
                                                        If Dir(sakifailego) <> "" Then
                                                              Kill sakifailego
                                                        End If
                                                        Screen.MousePointer = 0
                                                        Exit Sub
                                                End If
                                             Else
                                                seikyuudata_ppv(voipii, 20) = "0" '課金なし
                                             End If
                                        End If
                                        
                                        
                                        
                                   '対応契約がない****************************************************************************************************
                                   
                                    Else
                                        Select Case seikyuudata_ppv(voipii, 8)
                                            Case "CIP22", "CIP23", "CIP24", "CIP25", "CIP26", "CIP27", "CIP28", "CIP29", "CIP82", "CIP83"
                                                'ｷｬﾝﾍﾟｰﾝ割引があるのに、EZに有効な契約があるのはおかしい
                                                s_hikaku_str2 = ""
                                                s_hikaku_str3 = ""
                                                Select Case seikyuudata_ppv(voipii, 8)
                                                    Case "CIP22"
                                                        s_hikaku_str2 = "33" '基本料金
                                                        s_hikaku_str3 = "基本料金"
                                                    Case "CIP24"
                                                        s_hikaku_str2 = "35" '番号表示
                                                        s_hikaku_str3 = "番号表示"
                                                    Case "CIP26"
                                                        s_hikaku_str2 = "36" '番号リクエスト
                                                        s_hikaku_str3 = "番号リクエスト"
                                                    Case "CIP23"
                                                        s_hikaku_str2 = "37" 'キャッチ
                                                        s_hikaku_str3 = "キャッチ"
                                                    Case "CIP27"
                                                        s_hikaku_str2 = "38" '着信お断り
                                                        s_hikaku_str3 = "着信お断り"
                                                    Case "CIP25"
                                                        s_hikaku_str2 = "39" '着信転送
                                                        s_hikaku_str3 = "着信転送"
                                                    Case "CIP28"
                                                        s_hikaku_str2 = "40" 'パック１
                                                        s_hikaku_str3 = "パック１"
                                                    Case "CIP29"
                                                        s_hikaku_str2 = "41" 'パック２
                                                        s_hikaku_str3 = "パック２"
                                                    Case "CIP82"
                                                        s_hikaku_str2 = "42" 'パック３
                                                        s_hikaku_str3 = "パック３"
                                                    Case "CIP83"
                                                        s_hikaku_str2 = "43" 'パック４
                                                        s_hikaku_str3 = "パック４"
                                                End Select
                                                
                                               
                                                If s_hikaku_str2 <> "" Then
                                                    sql = "SELECT * FROM keiyaku  WHERE kojinid ='" & s_shiteikojinid & "' and sadou ='0' and shime ='0' and shurui='" & s_hikaku_str2 & "'"
                        
                                                    If FcSQlGet(rs_yotei3, sql, PrMsg) = True Then
                                                        s_err_msg = "顧客ID：" & s_shiteikojinid & "　に、請求元データにキャンペーンの割引があるのに、主契約「" & s_hikaku_str3 & "」があります。契約内容と契約金額を確認してください。"
                                                        
                                                        MsgBox s_err_msg
                                                    
                                                       ' If Dir(sakifailego) <> "" Then
                                                             ' Kill sakifailego
                                                      ' End If
                                                       ' Screen.MousePointer = 0
                                                       ' Exit Sub
                                                    End If
                                                End If
                                                
                                                seikyuudata_ppv(voipii, 20) = "0" '課金なし
                                            Case Else
                                                seikyuudata_ppv(voipii, 20) = "1" '課金する
                                         End Select
                                    End If
                                    '************************************************************************************
                                    
                                  
                                End If
                                
                            Else
                                seikyuudata_ppv(voipii, 20) = "0" '課金なし
                            End If
                    
                End If
        End If
    Next voipii
    
    
    'EZManager のデータ取得
    
    Dim s_ezmanager_cl_data()
    Dim s_ezmanager_cl_count As Integer
    
    Dim s_str As String, i_counter As Integer
    i_counter = 0
    s_ezmanager_cl_count = 0
    
    
    If frmkoushin.chkmushi.Value = 0 Then
    
            s_str = " and (keiyaku.shurui ='33' or  keiyaku.shurui ='35' or keiyaku.shurui ='36' or keiyaku.shurui ='37'" & _
            " or keiyaku.shurui ='38' or keiyaku.shurui ='39' or keiyaku.shurui ='40' or keiyaku.shurui ='41'" & _
            " or keiyaku.shurui ='42' or keiyaku.shurui ='43')"
            
            
            sql = "SELECT keiyaku.*,shurui.* FROM keiyaku left join shurui on keiyaku.shurui=shurui.shuruiid" & _
            " WHERE keiyaku.kadou ='0' and keiyaku.shime ='0'" & s_str
                                
            If FcSQlGet(rs_yotei4, sql, PrMsg) = True Then
                If rs_yotei4.RecordCount = 0 Then
                    'EZManagerに契約がない
                        Screen.MousePointer = 0
                        MsgBox "EZManagerには、契約データが登録されていないため、EZManagerにあって請求データにない場合のチェックは行われません。"
                Else
                    '契約がある（それぞれを請求データと比較する）
                    s_ezmanager_cl_count = rs_yotei4.RecordCount
                    ReDim s_ezmanager_cl_data(5, s_ezmanager_cl_count)
                    
                    rs_yotei4.MoveFirst
                    Do Until rs_yotei4.EOF
                        s_ezmanager_cl_data(0, i_counter) = rs_yotei4!kojinid
                        s_ezmanager_cl_data(1, i_counter) = rs_yotei4!shurui
                        s_ezmanager_cl_data(2, i_counter) = rs_yotei4!shuruimei
                        s_ezmanager_cl_data(3, i_counter) = ""
                         For ii3 = 0 To chk_sumi_count - 1
                            If chk_sumi_data(0, ii3) = rs_yotei4!kojinid Then
                                   s_ezmanager_cl_data(3, i_counter) = chk_sumi_data(1, ii3)
                                   Exit For
                            End If
                        Next
                       s_ezmanager_cl_data(4, i_counter) = ""
                     
                        rs_yotei4.MoveNext
                        i_counter = i_counter + 1
                    Loop
                    rs_yotei4.Close
                        
                   
                End If
            Else
                'EZManagerに契約がない
                  Screen.MousePointer = 0
                    MsgBox "抽出エラーのため、EZManagerにあって請求データにない場合のチェックは行われません。"
            End If
    
    
    
    
            Dim s_st As Integer, s_err2 As String, s_henkan As String
            
            For i = 0 To s_ezmanager_cl_count - 1
                s_st = 0
                s_err2 = "個人ＩＤ：" & s_ezmanager_cl_data(0, i) & "の" & s_ezmanager_cl_data(2, i)
                 For voipii = 0 To ColCount - 1
                        If s_ezmanager_cl_data(3, i) = seikyuudata_ppv(voipii, 5) Then
                            If seikyuudata_ppv(voipii, 5) = "0553-31-9556" Then
                                'MsgBox "OK"
                            End If
                                s_henkan = ""
                                
                                Select Case seikyuudata_ppv(voipii, 8)
                                Case "CIP5K"
                                    s_henkan = "33" '基本料金
                                Case "CIP5M"
                                    s_henkan = "35" '番号表示
                                Case "CIP5O"
                                    s_henkan = "36" '番号リクエスト
                                Case "CIP5L"
                                    s_henkan = "37" 'キャッチ
                                Case "CIP5P"
                                    s_henkan = "38" '着信お断り
                                Case "CIP5N"
                                    s_henkan = "39" '着信転送
                                Case "CIP5Q"
                                    s_henkan = "40" 'パック１
                                Case "CIP5R"
                                    s_henkan = "41" 'パック２
                                Case "CIP80"
                                    s_henkan = "42" 'パック３
                                Case "CIP81"
                                    s_henkan = "43" 'パック４
                                End Select
                                
                            If s_henkan <> "" Then
                                If s_henkan = s_ezmanager_cl_data(1, i) Then
                                    s_st = 1
                                    Exit For
                                End If
                            End If
                    End If
                 Next
                If s_st = 0 Then
                    'MsgBox s_err2 & "　の請求データがありません。取り込み後、契約と金額を確認してください"
                    
                   '  If Dir(sakifailego) <> "" Then
                       '   Kill sakifailego
                      '  End If
                      '  Screen.MousePointer = 0
                   ' Exit Sub
                End If
            Next
    

    End If
    
    
    
  
    If chkumu = 1 Then
        Screen.MousePointer = 0
        MsgBox "登録されていないデータを登録してから再度実行してください。処理を中止します。"
        If Dir(sakifailego) <> "" Then
            Kill sakifailego
        End If
        Screen.MousePointer = 0
        Exit Sub
    End If


If rocked("on") = False Then
    Screen.MousePointer = 0
        MsgBox "個人データを更新中のフラッグを立てるのに失敗しました。"
        If Dir(sakifailego) <> "" Then
            Kill sakifailego
        End If
        Screen.MousePointer = 0
        Exit Sub
End If

'*************ｃｌデータの保存***************
        newnentsuki = s_nentsuki ' seikyuudata_ppv(0, 7)
        
        If soukatsu(0, "renban", 18, "", 1, 1) = False Then
        Screen.MousePointer = 0
                MsgBox "設定の読み込みに失敗しました。"
                If Dir(sakifailego) <> "" Then
                    Kill sakifailego
                End If
                Screen.MousePointer = 0
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
        
        
    If kidou_cl = False Then
        If Dir(sakifailego) <> "" Then
            Kill sakifailego
        End If
        Screen.MousePointer = 0
        Exit Sub
    End If
    
  
        
        For voipii = 0 To ColCount - 1
               If seikyuudata_ppv(voipii, 1) = "2" And seikyuudata_ppv(voipii, 2) = "5100" Then    '各顧客のデータのみ
                    If seikyuudata_ppv(voipii, 5) <> "99999999999999999999999" Then    '消費税の総計を除外
                        If seikyuudata_ppv(voipii, 11) <> "Z" Then  '消費税の分を除外
                            
                                If Trim(seikyuudata_ppv(voipii, 15)) = "" Then
                                    seikyuudata_ppv(voipii, 15) = "0"
                                End If
                                
                                
                                sql = "INSERT INTO cltbl VALUES('" & newtsuwaid & "','" & newnentsuki & "'," & _
                                               "'" & seikyuudata_ppv(voipii, 19) & "','" & seikyuudata_ppv(voipii, 3) & _
                                               "','" & seikyuudata_ppv(voipii, 4) & "','" & seikyuudata_ppv(voipii, 5) & _
                                               "','" & seikyuudata_ppv(voipii, 7) & "','" & seikyuudata_ppv(voipii, 8) & _
                                               "','" & seikyuudata_ppv(voipii, 9) & "','" & seikyuudata_ppv(voipii, 10) & _
                                               "','" & seikyuudata_ppv(voipii, 11) & "'," & seikyuudata_ppv(voipii, 15) & "," & seikyuudata_ppv(voipii, 20) & ",'')"
                                
                                db_cl.Execute sql, dbSQLPassThrough
                                
                                newtsuwaid = Format$(CStr(CLng(newtsuwaid) + 1), "0000000#")
                            
                        End If
                    End If
                End If
        Next
        newtsuwaid = CStr(CLng(newtsuwaid))
        If soukatsu(1, "renban", 18, newtsuwaid, 1, 1) = False Then
        Screen.MousePointer = 0
            Rollback
                MsgBox "設定の書き込みに失敗しました。"
                If Dir(sakifailego) <> "" Then
                    Kill sakifailego
                End If
                Screen.MousePointer = 0
                Exit Sub
        End If

'*************ｃｌデータの請求***************



            sql = "SELECT kojinid" & _
                " FROM cltbl" & _
                " WHERE  nentsuki='" & newnentsuki & "' group by kojinid"
                
            If FcSQlGet_cl(rirekirs55, sql, PrMsg) <= 0 Then
                    MsgBox "請求データ用個人データが見つかりません。管理者に連絡してください。"
                    If Dir(sakifailego) <> "" Then
                    Kill sakifailego
                End If
                    Screen.MousePointer = 0
                    Exit Sub
            Else
                rirekirs55.MoveFirst
                Do Until rirekirs55.EOF
                        jitsu_kaki(0) = 0
                        jitsu_kaki(1) = 0
                        jitsu_kaki(2) = 0
                        jitsu_kaki(3) = 0
                        jitsu_kaki(4) = 0
                        sql = "SELECT *" & _
                            " FROM cltbl" & _
                            " WHERE kojinid='" & rirekirs55!kojinid & "' and nentsuki='" & newnentsuki & "' and kakin='1' order by clno DESC"
                            
                        If FcSQlGet_cl(rirekirs555, sql, PrMsg) > 0 Then
                            
                            rirekirs555.MoveFirst
                            Do Until rirekirs555.EOF
                                If rirekirs555!zeikubun = "S" Then
                                    jitsu_kaki(0) = jitsu_kaki(0) + CDbl(rirekirs555!kingaku)
                                Else
                                    jitsu_kaki(1) = jitsu_kaki(1) + CDbl(rirekirs555!kingaku)
                                End If
                                rirekirs555.MoveNext
                            Loop
                            'jitsu_kaki(2) = jitsu_kaki(0) * 0.05
                            'jitsu_kaki(2) = jitsu_kaki(0) * 0.08
                            jitsu_kaki(2) = jitsu_kaki(0) * 0.1
                            jitsu_kaki(3) = jitsu_kaki(0) + jitsu_kaki(2) + jitsu_kaki(1)
                            jitsu_kaki(4) = Int(jitsu_kaki(3) / 100)
                            If cl_koushin("43", CStr(rirekirs55!kojinid), CStr(jitsu_kaki(4)), newnentsuki & "01") = -1 Then
                                Screen.MousePointer = 0
                                MsgBox "請求データ用詳細データの登録に失敗しました。管理者に連絡してください。"
                                If Dir(sakifailego) <> "" Then
                                    Kill sakifailego
                                End If
                                Exit Sub
                            End If
                        End If
                     rirekirs55.MoveNext
                Loop
                rirekirs55.Close
            End If
            
            

    If rocked("off") = False Then
        MsgBox "個人データを更新中のフラッグの初期化に失敗しました。"
    End If

     If Dir(FilePathppv(0)) <> "" Then
        Kill FilePathppv(0)
   End If
    
    
    Screen.MousePointer = 0
    MsgBox "請求データの作成が完了しました。"


Exit Sub


Errppv1:
    Screen.MousePointer = 0
    If Dir(sakifailego) <> "" Then
        Kill sakifailego
   End If
   Screen.MousePointer = 0
    MsgBox "ケーブルラインデータの読み込みデータの書き込みに失敗しました。"
    Exit Sub

End Sub

Sub cl_tools_chk(newnentsuki As String)


'主契約とその他の契約のデータを取得
Dim s_ezmanager_cl_data() As String, s_ezmanager_cl_count As Integer, s_str As String
Dim rs_yotei4 As Recordset, i_counter As Integer

s_str = " and (keiyaku.shurui ='33' or  keiyaku.shurui ='35' or keiyaku.shurui ='36' or keiyaku.shurui ='37'" & _
            " or keiyaku.shurui ='38' or keiyaku.shurui ='39' or keiyaku.shurui ='40' or keiyaku.shurui ='41'" & _
            " or keiyaku.shurui ='42' or keiyaku.shurui ='43')"
            
            
            sql = "SELECT keiyaku.*,shurui.* FROM keiyaku left join shurui on keiyaku.shurui=shurui.shuruiid" & _
            " WHERE keiyaku.kadou ='0' and keiyaku.shime ='0'" & s_str & " order by keiyaku.kojinid,keiyaku.shurui"
                                
            If FcSQlGet(rs_yotei4, sql, PrMsg) = True Then
                If rs_yotei4.RecordCount = 0 Then
                    'EZManagerに契約がない
                        Screen.MousePointer = 0
                        MsgBox "EZManagerには、契約データが登録されていないため、EZManagerにあって請求データにない場合のチェックは行われません。"
                         Exit Sub
                Else
                    '契約がある（それぞれを請求データと比較する）
                    s_ezmanager_cl_count = rs_yotei4.RecordCount
                    ReDim s_ezmanager_cl_data(5, s_ezmanager_cl_count)
                    
                    i_counter = 0
                    
                    rs_yotei4.MoveFirst
                    Do Until rs_yotei4.EOF
                        s_ezmanager_cl_data(0, i_counter) = rs_yotei4!kojinid
                        s_ezmanager_cl_data(1, i_counter) = rs_yotei4!shurui
                        s_ezmanager_cl_data(2, i_counter) = rs_yotei4!shuruimei
                        s_ezmanager_cl_data(3, i_counter) = rs_yotei4!kuchisuu
                       s_ezmanager_cl_data(4, i_counter) = ""
                        s_ezmanager_cl_data(5, i_counter) = ""
                        rs_yotei4.MoveNext
                        i_counter = i_counter + 1
                    Loop
                    rs_yotei4.Close
                        
                   
                End If
            Else
                'EZManagerに契約がない
                  Screen.MousePointer = 0
                   MsgBox "抽出エラーのため、EZManagerにあって請求データにない場合のチェックは行われません。"
                   Exit Sub
            End If
    




        If kidou_cl = False Then
            MsgBox "データベースファイルを開けませんでした。"
            Exit Sub
        End If


Dim rirekirs55 As Recordset, rirekirs555 As Recordset, rirekirs5555 As Recordset, s_kensakustr As String, n_chantoiru As Integer
Dim s_k_id As String, s_hikaku_str As String

           sql = "SELECT kojinid FROM cltbl WHERE  nentsuki='" & newnentsuki & "' group by kojinid"
                
            If FcSQlGet_cl(rirekirs55, sql, PrMsg) <= 0 Then
                Screen.MousePointer = 0
                Exit Sub
            Else
                rirekirs55.MoveFirst
                Do Until rirekirs55.EOF
                        s_k_id = rirekirs55!kojinid
                        If rirekirs55!kojinid = "00618" Then
                            'MsgBox "618"
                        End If
                        
                        sql = "SELECT *" & _
                            " FROM cltbl" & _
                            " WHERE  kojinid = '" & rirekirs55!kojinid & "' and nentsuki='" & newnentsuki & "' and kakin='0' order by uchiwakecode desc"
                        If FcSQlGet_cl(rirekirs555, sql, PrMsg) <= 0 Then
                            Screen.MousePointer = 0
                            Exit Sub
                        Else
                            rirekirs555.MoveFirst
                            Do Until rirekirs555.EOF
                                    n_chantoiru = 0 '初期化
                                    s_kensakustr = ""
                                  Select Case Trim(rirekirs555!uchiwakecode)
                                        Case "CIP5K"
                                            s_kensakustr = "33" '基本料金
                                        Case "CIP5M"
                                            s_kensakustr = "35" '番号表示
                                        Case "CIP5O"
                                            s_kensakustr = "36" '番号リクエスト
                                        Case "CIP5L"
                                            s_kensakustr = "37" 'キャッチ
                                        Case "CIP5P"
                                            s_kensakustr = "38" '着信お断り
                                        Case "CIP5N"
                                            s_kensakustr = "39" '着信転送
                                        Case "CIP5Q"
                                            s_kensakustr = "40" 'パック１
                                        Case "CIP5R"
                                            s_kensakustr = "41" 'パック２
                                        Case "CIP80"
                                            s_kensakustr = "42" 'パック３
                                        Case "CIP81"
                                            s_kensakustr = "43" 'パック４
                                        Case "CIPE1"
                                            '工事費
                                            n_chantoiru = 1
                                        Case Else
                                           '割り引き
                                           n_chantoiru = 1
                                  End Select
                                  If s_kensakustr <> "" Then
                                        For i = 0 To s_ezmanager_cl_count - 1
                                              If Trim(rirekirs555!kojinid) = s_ezmanager_cl_data(0, i) Then '個人IDが同じ
                                                  If s_kensakustr = s_ezmanager_cl_data(1, i) Then '種類が同じ
                                                      If s_ezmanager_cl_data(4, i) = "" Then 'まだ使ってない
                                                          If s_ezmanager_cl_data(3, i) = "1" Then '契約数１
                                                              s_ezmanager_cl_data(4, i) = "1"
                                                                n_chantoiru = 1
                                                                Exit For
                                                          Else '契約数1以上
                                                              If s_ezmanager_cl_data(5, i) = "" Then '初
                                                                  s_ezmanager_cl_data(5, i) = "1"
                                                                  n_chantoiru = 1
                                                                      Exit For
                                                              Else
                                                                  If CInt(s_ezmanager_cl_data(5, i)) + 1 = CInt(s_ezmanager_cl_data(3, i)) Then
                                                                      s_ezmanager_cl_data(4, i) = "1"
                                                                      n_chantoiru = 1
                                                                      Exit For
                                                                  Else
                                                                      s_ezmanager_cl_data(5, i) = CStr(CInt(s_ezmanager_cl_data(5, i)) + 1)
                                                                      n_chantoiru = 1
                                                                      Exit For
                                                                  End If
                                                              End If
                                                          End If
                                                      End If
                                                  End If
                                              End If
                                        Next
                                  End If
                                  
                                  If n_chantoiru = 0 Then
                                        'いない
                                        'MsgBox "NG"
                                        Select Case Trim(rirekirs555!uchiwakecode)
                                              Case "CIP5K"
                                                    s_hikaku_str = "CIP22" '基本料金
                                                Case "CIP5M"
                                                    s_hikaku_str = "CIP24" '番号表示
                                                Case "CIP5O"
                                                    s_hikaku_str = "CIP26" '番号リクエスト
                                                Case "CIP5L"
                                                    s_hikaku_str = "CIP23" 'キャッチ
                                                Case "CIP5P"
                                                    s_hikaku_str = "CIP27" '着信お断り
                                                Case "CIP5N"
                                                    s_hikaku_str = "CIP25" '着信転送
                                                Case "CIP5Q"
                                                    s_hikaku_str = "CIP28" 'パック１
                                                Case "CIP5R"
                                                    s_hikaku_str = "CIP29" 'パック２
                                                Case "CIP80"
                                                    s_hikaku_str = "CIP82" 'パック３
                                                Case "CIP81"
                                                    s_hikaku_str = "CIP83" 'パック４
                                                Case Else
                                                     MsgBox "[契約]" & Trim(rirekirs555!uchiwakecode) & "、[個人ID]" & s_k_id & " が不正です。管理者に連絡してください。"
                                                    Exit Sub
                                        End Select
                                        sql = "SELECT *" & _
                                            " FROM cltbl" & _
                                            " WHERE  kojinid = '" & s_k_id & "' and nentsuki='" & newnentsuki & "' and uchiwakecode='" & s_hikaku_str & "'"
                                        If FcSQlGet_cl(rirekirs5555, sql, PrMsg) <= 0 Then
                                              MsgBox "[契約]" & Trim(rirekirs555!uchiwakecode) & "、[個人ID]" & s_k_id & " が不正です。"
                                        Else
                                            'OK
                                        
                                        End If
                                  End If
                                  
                                 rirekirs555.MoveNext
                            Loop
                            rirekirs555.Close
                        End If
                        
                     rirekirs55.MoveNext
                Loop
                rirekirs55.Close
            End If
            
            For i = 0 To s_ezmanager_cl_count - 1
               If s_ezmanager_cl_data(4, i) <> "1" Then
                    MsgBox "[契約]" & s_ezmanager_cl_data(2, i) & "、[個人ID]" & s_ezmanager_cl_data(0, i) & " が不正です。"
               End If
            Next
            
            
            MsgBox "終了しました。"

End Sub

Sub cl_data_convert2(s_path As String)
    '***********データを指定してコンバートする**************************
   ' このプロシージャはセミコロンで区切られた
   ' ファイルからデータを配列内にロードします。
   Dim GetLine As Variant
   Dim Spot As Integer, Position As Integer, newkojinid As String
   Dim Row As Integer, Col As Integer, newvoipnospot2 As Integer, newnentsuki As String
   Dim RowCount As Integer, ColCount As Integer, ColCount2 As Integer, ColCount3 As Integer
    Dim udtBrowseInfo As BROWSEINFO
    Dim lngIDList     As Long, newzei As Long, newsou As Long, newvoipnospot As Integer, newv1 As String, newv2 As String, newv3 As String
    Dim strPath       As String, kariatai, sql_ppv2 As String, newvoipno As String, sql_ppv299 As String
    Dim newtsuwaid As String, voipii As Integer, rs_ppv As Recordset, rs_ppv2 As Recordset, rs_ppv299 As Recordset, rs_yotei3 As Recordset, rs_yotei4 As Recordset
    Dim rirekirs55 As Recordset, rirekirs555 As Recordset
    Dim kakunin_atai(4) As Long, jitsu_kaki(5) As Double
    Dim FilePathppv(3), seikyuudata_ppv()
    
Dim s_nentsuki As String

FilePathppv(2) = s_path

    If Right(FilePathppv(2), 1) = "\" Then
        FilePathppv(2) = Mid(FilePathppv(2), 1, Len(FilePathppv(2)) - 1)
    End If
    Dim k_t As Integer, k_n As Long
    
    If Format(Date, "mm") = "01" Then
        k_t = 12
        k_n = CLng(Format(Date, "yyyy")) - 1
    Else
        k_t = CInt(Format(Date, "mm")) - 1
        k_n = CLng(Format(Date, "yyyy"))
    End If
    
    '*******元の読込ファイル（山梨CTAV→勝沼CATV）***********************
    s_nentsuki = Format(k_n, "0000") & Format(k_t, "00")
    
  
        'FilePathppv(0) = FilePathppv(2) & "\勝沼CATV用データ\cl_" & Format(k_n, "0000") & "_" & Format(k_t, "00") & "_katsunuma.txt"
        FilePathppv(0) = FilePathppv(2) & "\勝沼CATV用データ\cl_" & Format(k_n, "0000") & "_" & Format(k_t, "00") & "_katsunuma2.txt"
  
  
    
   If Dir(FilePathppv(0)) = "" Then
        MsgBox "更新したいファイルがありません。再度確認してください。"
        Exit Sub
   End If
   
   
        If MsgBox("読み込みと登録データのチェックを開始します。よろしいですか？", vbYesNo) = vbNo Then
            Exit Sub
        End If
  
    
Screen.MousePointer = 11
   
    'FDのコピーをHDDに実施
    Dim sakifile9 As String, sakifailego As String
    sakifile9 = sopa & "fdback\cl\katsunuma\"
    If Dir$(sakifile9, vbDirectory) = "" Then
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder (sakifile9)
        Set fso = Nothing
    End If
    If Right(sakifile9, 1) <> "\" Then
        sakifile9 = sakifile9 & "\"
    End If
    
     '*******保存ファイル（勝沼CATV）***********************
     
            sakifailego = sakifile9 & "cl_" & Format(k_n, "0000") & "_" & Format(k_t, "00") & "_katsunuma2.txt"
   
    
   If Dir(sakifailego) <> "" Then
        MsgBox "更新したいファイルは、すでに更新した疑いがあります。再度確認してください。"
        Screen.MousePointer = 0
        Exit Sub
   End If
    
    
    'コピー
    FileCopy FilePathppv(0), sakifailego
   


    '読込中
On Error GoTo Errppv1
   Open FilePathppv(0) For Input Lock Read Write As #1
   ColCount = 0
   RowCount = 65
    While Not EOF(1)
        Line Input #1, GetLine
        ColCount = ColCount + 1
    Wend
    Close #1
    
    Open FilePathppv(0) For Input Lock Read Write As #1
   ReDim seikyuudata_ppv(ColCount - 1, RowCount)
   Position = 1
   Row = 0
   While Not EOF(1)
      Line Input #1, GetLine
      Col = 1
      Spot = InStr(1, GetLine, ",")
      While Spot <> 0
         kariatai = Left$(GetLine, Spot - 1)
         If Col <> 12 Then
            If Mid(kariatai, 1, 1) = """" Then
                If Right(kariatai, 1) = """" Then
                    seikyuudata_ppv(Row, Col) = Mid(kariatai, 2, Len(kariatai) - 2)
                Else
                    seikyuudata_ppv(Row, Col) = Mid(kariatai, 2)
                End If
            Else
                If Right(kariatai, 1) = """" Then
                    seikyuudata_ppv(Row, Col) = Mid(kariatai, 1, Len(kariatai) - 1)
                Else
                    seikyuudata_ppv(Row, Col) = kariatai
                End If
            End If
        Else
            seikyuudata_ppv(Row, Col) = kariatai
        End If
         Col = Col + 1
         GetLine = Right$(GetLine, Len(GetLine) - Spot)
         Spot = InStr(1, GetLine, ",")
      Wend
      If Len(GetLine) <> 0 Then
            If Mid(GetLine, 1, 1) = """" Then
                If Right(GetLine, 1) = """" Then
                    seikyuudata_ppv(Row, Col) = Mid(GetLine, 2, Len(GetLine) - 2)
                Else
                    seikyuudata_ppv(Row, Col) = Mid(GetLine, 2)
                End If
            Else
                If Right(GetLine, 1) = """" Then
                    seikyuudata_ppv(Row, Col) = Mid(GetLine, 1, Len(GetLine) - 1)
                Else
                    seikyuudata_ppv(Row, Col) = GetLine
                End If
            End If
      End If
      Row = Row + 1
   Wend
   Close #1
   
   
   
'***************既存データチェック****************************************
Dim chk_sumi_data()
Dim chk_sumi_count As Integer



    Dim rs_yotei2  As Recordset, s_shiteikojinid As String, s_kensakustr As String, s_koumoku As String, s_err_msg As String, s_hikaku_str2 As String, s_hikaku_str3 As String
 Dim chkumu As Integer, iii As Integer
   chkumu = 0
   chk_sumi_count = 0
   
    For voipii = 0 To ColCount - 1
        If seikyuudata_ppv(voipii, 1) = "2" And seikyuudata_ppv(voipii, 2) = "5100" Then
               If seikyuudata_ppv(voipii, 5) <> "99999999999999999999999" Then
                        s_shiteikojinid = ""
                        s_hikaku_str2 = ""
                        s_koumoku = ""
                        
                        'CLの電話番号の有無の確認 *****************************************************************************
                        
                        sql_ppv299 = "select * from cl where tel='" & seikyuudata_ppv(voipii, 5) & "'"
                
                        If FcSQlGet(rs_ppv299, sql_ppv299, PrMsg) = False Then
                            Screen.MousePointer = 0
                            MsgBox "CL番号「" & seikyuudata_ppv(voipii, 5) & "」は、登録されていません。確認してから再度実行してください。"
                            chkumu = 1
                            Exit For
                        Else
                            If rs_ppv299.RecordCount <> 1 Then
                                Screen.MousePointer = 0
                                MsgBox "CL番号「" & seikyuudata_ppv(voipii, 5) & "」は、複数登録されています。確認してから再度実行してください。"
                                rs_ppv299.MoveFirst
                                Do Until rs_ppv299.EOF
                                    MsgBox "ID「" & rs_ppv299!kojinid & "」に登録されています。チェックしてください。"
                                    rs_ppv299.MoveNext
                                Loop
                                rs_ppv299.Close
                                chkumu = 1
                                Exit For
                            Else
                                seikyuudata_ppv(voipii, 19) = rs_ppv299!kojinid
                                 s_shiteikojinid = rs_ppv299!kojinid
                                 
                                 
                                 Dim s_aruno As Integer, i3i As Integer, n_index As Integer
                                 s_aruno = 0
                                For i3i = 0 To chk_sumi_count - 1
                                    n_index = i3i
                                    If chk_sumi_data(0, i3i) = s_shiteikojinid Then
                                        If chk_sumi_data(1, i3i) = seikyuudata_ppv(voipii, 5) Then
                                            s_aruno = 1
                                            Exit For
                                        End If
                                    End If
                                Next
                                If s_aruno = 0 Then
                                        ReDim Preserve chk_sumi_data(3, chk_sumi_count + 1)
                                        
                                        chk_sumi_data(0, chk_sumi_count) = s_shiteikojinid
                                         chk_sumi_data(1, chk_sumi_count) = seikyuudata_ppv(voipii, 5)
                                         chk_sumi_data(2, chk_sumi_count) = "1"
                                         
                                         chk_sumi_count = chk_sumi_count + 1
                                 
                                 End If
                                  
                                  
                            End If
                            
                        End If
                        '******************************************************************************************
                        
                        If seikyuudata_ppv(voipii, 8) <> "CIPE1" Then
                                '課金課の判断
                                If s_shiteikojinid <> "" Then
                                    s_kensakustr = ""
                                    Select Case seikyuudata_ppv(voipii, 8)
                                        Case "CIP5K"
                                            s_kensakustr = "33" '基本料金
                                            s_koumoku = "基本料金"
                                        Case "CIP5M"
                                            s_kensakustr = "35" '番号表示
                                            s_koumoku = "番号表示"
                                        Case "CIP5O"
                                            s_kensakustr = "36" '番号リクエスト
                                            s_koumoku = "番号リクエスト"
                                        Case "CIP5L"
                                            s_kensakustr = "37" 'キャッチ
                                            s_koumoku = "キャッチ"
                                        Case "CIP5P"
                                            s_kensakustr = "38" '着信お断り
                                            s_koumoku = "着信お断り"
                                        Case "CIP5N"
                                            s_kensakustr = "39" '着信転送
                                            s_koumoku = "着信転送"
                                        Case "CIP5Q"
                                            s_kensakustr = "40" 'パック１
                                            s_koumoku = "パック１"
                                        Case "CIP5R"
                                            s_kensakustr = "41" 'パック２
                                            s_koumoku = "パック２"
                                        Case "CIP80"
                                            s_kensakustr = "42" 'パック３
                                            s_koumoku = "パック３"
                                        Case "CIP81"
                                            s_kensakustr = "43" 'パック４
                                            s_koumoku = "パック４"
                                    End Select
                                    
                                    '対応契約があり****************************************************************
                                    If s_kensakustr <> "" Then
                                    
                                        Dim n_kaisuu As Integer, ii3 As Integer
                                        n_kaisuu = 0
                                             For ii3 = 0 To chk_sumi_count - 1
                                                If chk_sumi_data(0, ii3) = s_shiteikojinid Then
                                                        n_kaisuu = n_kaisuu + 1
                                                End If
                                            Next
                                    
                                    
                                        'sql = "SELECT * FROM keiyaku  WHERE kojinid ='" & s_shiteikojinid & "' and sadou ='0' and shime ='0' and shurui='" & s_kensakustr & "'"
                                         sql = "SELECT * FROM keiyaku  WHERE kojinid ='" & s_shiteikojinid & "' and kuchisuu >='" & CStr(n_kaisuu) & "' and kadou ='0' and shime ='0' and shurui='" & s_kensakustr & "'"
                                         
                                        If FcSQlGet(rs_yotei2, sql, PrMsg) = True Then
                                             seikyuudata_ppv(voipii, 20) = "0" '課金なし
                                        Else
                                            '契約がないとき（契約月ではEZManagerには主契約がないが、請求には＋とーがあがってくる）
                                            Dim s_mainasuari As Integer, s_hikaku_str As String
                                            s_mainasuari = 0
                                            s_hikaku_str = ""
                                            Select Case seikyuudata_ppv(voipii, 8)
                                                Case "CIP5K"
                                                    s_hikaku_str = "CIP22" '基本料金
                                                Case "CIP5M"
                                                    s_hikaku_str = "CIP24" '番号表示
                                                Case "CIP5O"
                                                    s_hikaku_str = "CIP26" '番号リクエスト
                                                Case "CIP5L"
                                                    s_hikaku_str = "CIP23" 'キャッチ
                                                Case "CIP5P"
                                                    s_hikaku_str = "CIP27" '着信お断り
                                                Case "CIP5N"
                                                    s_hikaku_str = "CIP25" '着信転送
                                                Case "CIP5Q"
                                                    s_hikaku_str = "CIP28" 'パック１
                                                Case "CIP5R"
                                                    s_hikaku_str = "CIP29" 'パック２
                                                Case "CIP80"
                                                    s_hikaku_str = "CIP82" 'パック３
                                                Case "CIP81"
                                                    s_hikaku_str = "CIP83" 'パック４
                                            End Select
                                            
                                            For iii = 0 To ColCount - 1
                                                If seikyuudata_ppv(iii, 19) = s_shiteikojinid Then
                                                    If seikyuudata_ppv(iii, 8) = s_hikaku_str Then
                                                        s_mainasuari = 1
                                                        Exit For
                                                    End If
                                                End If
                                            Next iii
                                            
                                            If s_mainasuari = 0 Then
                                                'seikyuudata_ppv(voipii, 20) = "1" '課金する
                                                'CLにあってEZにない
                                                If frmkoushin.chkkaihi.Value = 1 Then
                                                        s_err_msg = "顧客ID：" & s_shiteikojinid & "　に、主契約「" & s_koumoku & "」がありませんが、今回は請求を回避します。"
                                                        
                                                        MsgBox s_err_msg
                                                        
                                                        seikyuudata_ppv(voipii, 20) = "0" '課金なし
                                                Else
                                                        s_err_msg = "顧客ID：" & s_shiteikojinid & "　に、主契約「" & s_koumoku & "」がありません。契約内容を確認し、追加後、再度実行してください。"
                                                        
                                                        MsgBox s_err_msg
                                                    
                                                        If Dir(sakifailego) <> "" Then
                                                              Kill sakifailego
                                                        End If
                                                        Screen.MousePointer = 0
                                                        Exit Sub
                                                End If
                                             Else
                                                seikyuudata_ppv(voipii, 20) = "0" '課金なし
                                             End If
                                        End If
                                        
                                        
                                        
                                   '対応契約がない****************************************************************************************************
                                   
                                    Else
                                        Select Case seikyuudata_ppv(voipii, 8)
                                            Case "CIP22", "CIP23", "CIP24", "CIP25", "CIP26", "CIP27", "CIP28", "CIP29", "CIP82", "CIP83"
                                                'ｷｬﾝﾍﾟｰﾝ割引があるのに、EZに有効な契約があるのはおかしい
                                                s_hikaku_str2 = ""
                                                s_hikaku_str3 = ""
                                                Select Case seikyuudata_ppv(voipii, 8)
                                                    Case "CIP22"
                                                        s_hikaku_str2 = "33" '基本料金
                                                        s_hikaku_str3 = "基本料金"
                                                    Case "CIP24"
                                                        s_hikaku_str2 = "35" '番号表示
                                                        s_hikaku_str3 = "番号表示"
                                                    Case "CIP26"
                                                        s_hikaku_str2 = "36" '番号リクエスト
                                                        s_hikaku_str3 = "番号リクエスト"
                                                    Case "CIP23"
                                                        s_hikaku_str2 = "37" 'キャッチ
                                                        s_hikaku_str3 = "キャッチ"
                                                    Case "CIP27"
                                                        s_hikaku_str2 = "38" '着信お断り
                                                        s_hikaku_str3 = "着信お断り"
                                                    Case "CIP25"
                                                        s_hikaku_str2 = "39" '着信転送
                                                        s_hikaku_str3 = "着信転送"
                                                    Case "CIP28"
                                                        s_hikaku_str2 = "40" 'パック１
                                                        s_hikaku_str3 = "パック１"
                                                    Case "CIP29"
                                                        s_hikaku_str2 = "41" 'パック２
                                                        s_hikaku_str3 = "パック２"
                                                    Case "CIP82"
                                                        s_hikaku_str2 = "42" 'パック３
                                                        s_hikaku_str3 = "パック３"
                                                    Case "CIP83"
                                                        s_hikaku_str2 = "43" 'パック４
                                                        s_hikaku_str3 = "パック４"
                                                End Select
                                                
                                               
                                                If s_hikaku_str2 <> "" Then
                                                    sql = "SELECT * FROM keiyaku  WHERE kojinid ='" & s_shiteikojinid & "' and sadou ='0' and shime ='0' and shurui='" & s_hikaku_str2 & "'"
                        
                                                    If FcSQlGet(rs_yotei3, sql, PrMsg) = True Then
                                                        s_err_msg = "顧客ID：" & s_shiteikojinid & "　に、請求元データにキャンペーンの割引があるのに、主契約「" & s_hikaku_str3 & "」があります。契約内容を確認し、削除後、再度実行してください。"
                                                        
                                                        MsgBox s_err_msg
                                                    
                                                        If Dir(sakifailego) <> "" Then
                                                              Kill sakifailego
                                                        End If
                                                        Screen.MousePointer = 0
                                                        Exit Sub
                                                    End If
                                                End If
                                                
                                                seikyuudata_ppv(voipii, 20) = "0" '課金なし
                                            Case Else
                                                seikyuudata_ppv(voipii, 20) = "1" '課金する
                                         End Select
                                    End If
                                    '************************************************************************************
                                    
                                  
                                End If
                                
                            Else
                                seikyuudata_ppv(voipii, 20) = "0" '課金なし
                            End If
                    
                End If
        End If
    Next voipii
    
    
    'EZManager のデータ取得
    
    Dim s_ezmanager_cl_data()
    Dim s_ezmanager_cl_count As Integer
    
    Dim s_str As String, i_counter As Integer
    i_counter = 0
    s_ezmanager_cl_count = 0
    
    
   
    
    
    
  
    If chkumu = 1 Then
        Screen.MousePointer = 0
        MsgBox "登録されていないデータを登録してから再度実行してください。処理を中止します。"
        If Dir(sakifailego) <> "" Then
            Kill sakifailego
        End If
        Screen.MousePointer = 0
        Exit Sub
    End If


If rocked("on") = False Then
    Screen.MousePointer = 0
        MsgBox "個人データを更新中のフラッグを立てるのに失敗しました。"
        If Dir(sakifailego) <> "" Then
            Kill sakifailego
        End If
        Screen.MousePointer = 0
        Exit Sub
End If

'*************ｃｌデータの保存***************
        newnentsuki = s_nentsuki ' seikyuudata_ppv(0, 7)
        
        If soukatsu(0, "renban", 18, "", 1, 1) = False Then
        Screen.MousePointer = 0
                MsgBox "設定の読み込みに失敗しました。"
                If Dir(sakifailego) <> "" Then
                    Kill sakifailego
                End If
                Screen.MousePointer = 0
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
        
        
    If kidou_cl = False Then
        If Dir(sakifailego) <> "" Then
            Kill sakifailego
        End If
        Screen.MousePointer = 0
        Exit Sub
    End If
    
  
        
        For voipii = 0 To ColCount - 1
               If seikyuudata_ppv(voipii, 1) = "2" And seikyuudata_ppv(voipii, 2) = "5100" Then    '各顧客のデータのみ
                    If seikyuudata_ppv(voipii, 5) <> "99999999999999999999999" Then    '消費税の総計を除外
                        If seikyuudata_ppv(voipii, 11) <> "Z" Then  '消費税の分を除外
                            
                                If Trim(seikyuudata_ppv(voipii, 15)) = "" Then
                                    seikyuudata_ppv(voipii, 15) = "0"
                                End If
                                
                                
                                sql = "INSERT INTO cltbl VALUES('" & newtsuwaid & "','" & newnentsuki & "'," & _
                                               "'" & seikyuudata_ppv(voipii, 19) & "','" & seikyuudata_ppv(voipii, 3) & _
                                               "','" & seikyuudata_ppv(voipii, 4) & "','" & seikyuudata_ppv(voipii, 5) & _
                                               "','" & seikyuudata_ppv(voipii, 7) & "','" & seikyuudata_ppv(voipii, 8) & _
                                               "','" & seikyuudata_ppv(voipii, 9) & "','" & seikyuudata_ppv(voipii, 10) & _
                                               "','" & seikyuudata_ppv(voipii, 11) & "'," & seikyuudata_ppv(voipii, 15) & "," & seikyuudata_ppv(voipii, 20) & ",'2')"
                                
                                db_cl.Execute sql, dbSQLPassThrough
                                
                                newtsuwaid = Format$(CStr(CLng(newtsuwaid) + 1), "0000000#")
                            
                        End If
                    End If
                End If
        Next
        newtsuwaid = CStr(CLng(newtsuwaid))
        If soukatsu(1, "renban", 18, newtsuwaid, 1, 1) = False Then
        Screen.MousePointer = 0
            Rollback
                MsgBox "設定の書き込みに失敗しました。"
                If Dir(sakifailego) <> "" Then
                    Kill sakifailego
                End If
                Screen.MousePointer = 0
                Exit Sub
        End If

'*************ｃｌデータの請求***************



            sql = "SELECT kojinid" & _
                " FROM cltbl" & _
                " WHERE  nentsuki='" & newnentsuki & "' and yomikaisuu ='2' group by kojinid"
                
            If FcSQlGet_cl(rirekirs55, sql, PrMsg) <= 0 Then
                    MsgBox "請求データ用個人データが見つかりません。管理者に連絡してください。"
                    If Dir(sakifailego) <> "" Then
                    Kill sakifailego
                End If
                    Screen.MousePointer = 0
                    Exit Sub
            Else
                rirekirs55.MoveFirst
                Do Until rirekirs55.EOF
                        jitsu_kaki(0) = 0
                        jitsu_kaki(1) = 0
                        jitsu_kaki(2) = 0
                        jitsu_kaki(3) = 0
                        jitsu_kaki(4) = 0
                        sql = "SELECT *" & _
                            " FROM cltbl" & _
                            " WHERE kojinid='" & rirekirs55!kojinid & "' and nentsuki='" & newnentsuki & "' and kakin='1' order by clno DESC"
                            
                        If FcSQlGet_cl(rirekirs555, sql, PrMsg) > 0 Then
                           
                            rirekirs555.MoveFirst
                            Do Until rirekirs555.EOF
                                If rirekirs555!zeikubun = "S" Then
                                    jitsu_kaki(0) = jitsu_kaki(0) + CDbl(rirekirs555!kingaku)
                                Else
                                    jitsu_kaki(1) = jitsu_kaki(1) + CDbl(rirekirs555!kingaku)
                                End If
                                rirekirs555.MoveNext
                            Loop
                            'jitsu_kaki(2) = jitsu_kaki(0) * 0.05
                            'jitsu_kaki(2) = jitsu_kaki(0) * 0.08
                            jitsu_kaki(2) = jitsu_kaki(0) * 0.1
                            jitsu_kaki(3) = jitsu_kaki(0) + jitsu_kaki(2) + jitsu_kaki(1)
                            jitsu_kaki(4) = Int(jitsu_kaki(3) / 100)
                            If cl_koushin("43", CStr(rirekirs55!kojinid), CStr(jitsu_kaki(4)), newnentsuki & "01") = -1 Then
                                Screen.MousePointer = 0
                                MsgBox "請求データ用詳細データの登録に失敗しました。管理者に連絡してください。"
                                If Dir(sakifailego) <> "" Then
                                    Kill sakifailego
                                End If
                                Exit Sub
                            End If
                        End If
                     rirekirs55.MoveNext
                Loop
                rirekirs55.Close
            End If
            
            

    If rocked("off") = False Then
        MsgBox "個人データを更新中のフラッグの初期化に失敗しました。"
    End If

     If Dir(FilePathppv(0)) <> "" Then
        Kill FilePathppv(0)
   End If
    
    
    Screen.MousePointer = 0
    MsgBox "請求データの作成が完了しました。"


Exit Sub


Errppv1:
    Screen.MousePointer = 0
    If Dir(sakifailego) <> "" Then
        Kill sakifailego
   End If
   Screen.MousePointer = 0
    MsgBox "ケーブルラインデータの読み込みデータの書き込みに失敗しました。"
    Exit Sub

End Sub

Function yotejikkou2(s_nentsuki2 As String) As Integer

yotejikkou2 = 0



Dim yotei_data() As String
Dim shokuinsuu As Integer


Dim s_nentsuki As String
Dim s_nen As Long, s_tsuki As Integer, s_tsuki2 As Integer, s_nen2 As Long
s_nen = CLng(Mid(s_nentsuki2, 1, 4))
s_tsuki = CInt(Mid(s_nentsuki2, 5, 2))


'********************************
If s_tsuki = 12 Then
    s_tsuki2 = 1
    s_nen2 = s_nen + 1
Else
    s_tsuki2 = s_tsuki + 1
    s_nen2 = s_nen
End If
'*******************************


s_nentsuki = CStr(s_nen2) & Format$(s_tsuki2, "0#")


Dim rs_yotei  As Recordset


'今回に予定があるかの判定
'sql = "SELECT * FROM yotei where kojinid ='" & s_kojinid & "' and itsukara <='" & s_nentsuki & "'  and itsumade >='" & s_nentsuki & "'" & _
        '" and st='0' ORDER BY yoteiid"
        
sql = "SELECT * FROM yotei where  itsukara <='" & s_nentsuki & "'  and itsumade >='" & s_nentsuki & "'" & _
        " and st='0' ORDER BY kojinid,yoteiid"

If FcSQlGet(rs_yotei, sql, PrMsg) = False Then
    rs_yotei.Close
    Exit Function
Else
    shokuinsuu = rs_yotei.RecordCount
    ReDim yotei_data(shokuinsuu, 11)
    rs_yotei.MoveFirst
For i = 0 To shokuinsuu - 1
       
            If rs_yotei!addoff = "0" Then
                yotei_data(i, 0) = "0" ' "追加"
                If rs_yotei!shurui = "0" Then
                    yotei_data(i, 1) = "0" ' "主契約"
                ElseIf rs_yotei!shurui = "1" Then
                    yotei_data(i, 1) = "1" '"その也"
                End If
            Else
                yotei_data(i, 0) = "1" ' "削除"
                If rs_yotei!shurui = "0" Then
                    yotei_data(i, 1) = "0" ' "削除"
                ElseIf rs_yotei!shurui = "1" Then
                    yotei_data(i, 1) = "1" '"休止"
                ElseIf rs_yotei!shurui = "2" Then
                    yotei_data(i, 1) = "2" '"中止"
                End If
            End If
            
         
            yotei_data(i, 2) = rs_yotei!kanrenid
            yotei_data(i, 3) = rs_yotei!kanrenmei
            yotei_data(i, 4) = rs_yotei!kingaku
            yotei_data(i, 5) = rs_yotei!kuchisuu
            yotei_data(i, 6) = rs_yotei!sousuu
           yotei_data(i, 7) = rs_yotei!itsukara
            yotei_data(i, 8) = rs_yotei!itsumade
             yotei_data(i, 9) = rs_yotei!yoteiid
            
            yotei_data(i, 10) = rs_yotei!kojinid
       rs_yotei.MoveNext
    Next i
    rs_yotei.Close
End If

Dim s_keizoku As Integer



Dim kyk As String, kyk2 As String

For i = 0 To shokuinsuu - 1
    If yotei_data(i, 0) = "1" Then ' "削除"
           If yotei_data(i, 1) = "0" Then '削除 "主契約"
           
                 kyk = "0"  'rs_yotei!kadou
                 kyk2 = yotei_data(i, 4)  'rs_yotei!kingaku
                ' kyk3 = rs_yotei!kaishi
                
                 sql = "delete * from keiyaku where keiyakuid ='" & yotei_data(i, 2) & "'"
                On Error GoTo errsakujo
                 db.Execute sql, dbSQLPassThrough
                On Error GoTo 0
                
                 'If a_shukeiyaku_sakujo(s_kojinid, kyk2, kyk) = False Then
                 If a_shukeiyaku_sakujo(yotei_data(i, 10), kyk2, kyk) = False Then
                     MsgBox "料金の更新に失敗しました。"
                      yotejikkou2 = -1
                     Exit Function
                 End If
               
            
            Else   'If yotei_data(i, 1) = "1" Then '休止1　中止2
                '変更
                Dim rs_yotei2  As Recordset, m_kin As String, m_shurui As String, hkingaku As String

                '現在のSTと価格を調べる
                sql = "SELECT * FROM keiyaku  WHERE keiyakuid ='" & yotei_data(i, 2) & "'"

                If FcSQlGet(rs_yotei2, sql, PrMsg) = True Then
                        
                         rs_yotei2.MoveFirst
                        
                        m_kin = rs_yotei2!kingaku
                        m_shurui = rs_yotei2!kadou
                        
                        rs_yotei2.Close
                        
                    
                        Dim hkudou As String
                        If yotei_data(i, 1) = "1" Then
                            hkudou = "1"
                        ElseIf yotei_data(i, 1) = "2" Then
                            hkudou = "2"
                        End If
                        
                        hkingaku = yotei_data(i, 4) 'rs_yotei!kingaku
                        
                        If m_shurui <> hkudou Then
                        
                            sql = "UPDATE keiyaku SET kadou = '" & hkudou & "'" & _
                               " WHERE keiyakuid = '" & yotei_data(i, 2) & "'"
                
                                On Error GoTo errk3
                                     db.Execute sql, dbSQLPassThrough
                                On Error GoTo 0
                                            
                             'If a_shukeiyaku_henkou(s_kojinid, m_kin, hkingaku, m_shurui, hkudou) = False Then
                             If a_shukeiyaku_henkou(yotei_data(i, 10), m_kin, hkingaku, m_shurui, hkudou) = False Then
                                MsgBox "料金の更新に失敗しました。"
                                 yotejikkou2 = -1
                                 Exit Function
                            End If
                        End If

                End If
                
            
            
            End If
            
            'ST更新
             On Error GoTo errsakujo2
                 sql = "UPDATE yotei SET st = '1'" & _
                              " WHERE yoteiid = '" & yotei_data(i, 9) & "'"
                 db.Execute sql, dbSQLPassThrough
             On Error GoTo 0
             
    
    Else  '"追加"
            Dim tkaishi As String, tkingaku As String, tkudou As String, shime As String
            
            tkaishi = Format(Date, "yyyymmdd")
            tkingaku = yotei_data(i, 4)
            tkudou = "0"
            shime = "0"
            
            If yotei_data(i, 1) = "0" Then ' "主契約"
                
                'If shukeiyaku_touroku(s_kojinid, yotei_data(i, 2), yotei_data(i, 5), yotei_data(i, 4), tkaishi, tkudou, shime) = 1 Then
                If shukeiyaku_touroku(yotei_data(i, 10), yotei_data(i, 2), yotei_data(i, 5), yotei_data(i, 4), tkaishi, tkudou, shime) = 1 Then
                
                    'If a_shukeiyaku_koushin(s_kojinid, tkaishi, tkingaku, tkudou, shime) = False Then
                    If a_shukeiyaku_koushin(yotei_data(i, 10), tkaishi, tkingaku, tkudou, shime) = False Then
                        MsgBox "料金の更新に失敗しました。"
                         yotejikkou2 = -1
                        Exit Function
                    Else
                          'ST更新
                            On Error GoTo errsakujo2
                                sql = "UPDATE yotei SET st = '1'" & _
                                             " WHERE yoteiid = '" & yotei_data(i, 9) & "'"
                                db.Execute sql, dbSQLPassThrough
                            On Error GoTo 0
                    End If
                End If
                
    
             Else '"その也"
             
                'If sonotakeiyaku_touroku(s_kojinid, yotei_data(i, 2), tkingaku, tkaishi) = 1 Then
                If sonotakeiyaku_touroku(yotei_data(i, 10), yotei_data(i, 2), tkingaku, tkaishi) = 1 Then
                
                    'If a_sonota_touroku(s_kojinid, tkingaku, tkudou) = False Then
                    If a_sonota_touroku(yotei_data(i, 10), tkingaku, tkudou) = False Then
                       MsgBox "料金の更新に失敗しました。"
                        yotejikkou2 = -1
                       Exit Function
                    Else
                        s_keizoku = 0
                        If s_nentsuki = yotei_data(i, 8) Then
                            s_keizoku = 1
                        End If
                        'ST更新
                        If s_keizoku = 1 Then
                            On Error GoTo errsakujo2
                                sql = "UPDATE yotei SET st = '1'" & _
                                             " WHERE yoteiid = '" & yotei_data(i, 9) & "'"
                                db.Execute sql, dbSQLPassThrough
                            On Error GoTo 0
                        End If
                    End If
                 
                 End If
            End If
            
            
    End If
    
    
     
     
                 
Next


Exit Function

errsakujo:
   MsgBox "契約追加の更新に失敗しました。"
   yotejikkou2 = -1
                     Exit Function
errsakujo2:
    MsgBox "契約削除の更新に失敗しました。"
    yotejikkou2 = -1
                     Exit Function

errk3:
    MsgBox "契約変更の更新に失敗しました。"
    yotejikkou2 = -1
                     Exit Function



End Function


Function yotejikkou(s_nentsuki2 As String, s_kojinid As String) As Integer

yotejikkou = 0



Dim yotei_data() As String
Dim shokuinsuu As Integer


Dim s_nentsuki As String
Dim s_nen As Long, s_tsuki As Integer, s_tsuki2 As Integer, s_nen2 As Long
s_nen = CLng(Mid(s_nentsuki2, 1, 4))
s_tsuki = CInt(Mid(s_nentsuki2, 5, 2))


'********************************
If s_tsuki = 12 Then
    s_tsuki2 = 1
    s_nen2 = s_nen + 1
Else
    s_tsuki2 = s_tsuki + 1
    s_nen2 = s_nen
End If
'*******************************


s_nentsuki = CStr(s_nen2) & Format$(s_tsuki2, "0#")


Dim rs_yotei  As Recordset


'今回に予定があるかの判定
sql = "SELECT * FROM yotei where kojinid ='" & s_kojinid & "' and itsukara <='" & s_nentsuki & "'  and itsumade >='" & s_nentsuki & "'" & _
        " and st='0' ORDER BY yoteiid"

If FcSQlGet(rs_yotei, sql, PrMsg) = False Then
    rs_yotei.Close
    Exit Function
Else
    shokuinsuu = rs_yotei.RecordCount
    ReDim yotei_data(shokuinsuu, 10)
    rs_yotei.MoveFirst
For i = 0 To shokuinsuu - 1
       
            If rs_yotei!addoff = "0" Then
                yotei_data(i, 0) = "0" ' "追加"
                If rs_yotei!shurui = "0" Then
                    yotei_data(i, 1) = "0" ' "主契約"
                ElseIf rs_yotei!shurui = "1" Then
                    yotei_data(i, 1) = "1" '"その也"
                End If
            Else
                yotei_data(i, 0) = "1" ' "削除"
                If rs_yotei!shurui = "0" Then
                    yotei_data(i, 1) = "0" ' "削除"
                ElseIf rs_yotei!shurui = "1" Then
                    yotei_data(i, 1) = "1" '"休止"
                ElseIf rs_yotei!shurui = "2" Then
                    yotei_data(i, 1) = "2" '"中止"
                End If
            End If
            
         
            yotei_data(i, 2) = rs_yotei!kanrenid
            yotei_data(i, 3) = rs_yotei!kanrenmei
            yotei_data(i, 4) = rs_yotei!kingaku
            yotei_data(i, 5) = rs_yotei!kuchisuu
            yotei_data(i, 6) = rs_yotei!sousuu
           yotei_data(i, 7) = rs_yotei!itsukara
            yotei_data(i, 8) = rs_yotei!itsumade
             yotei_data(i, 9) = rs_yotei!yoteiid
            
 
       rs_yotei.MoveNext
    Next i
    rs_yotei.Close
End If

Dim s_keizoku As Integer



Dim kyk As String, kyk2 As String

For i = 0 To shokuinsuu - 1
    If yotei_data(i, 0) = "1" Then ' "削除"
           If yotei_data(i, 1) = "0" Then '削除 "主契約"
           
                 kyk = "0"  'rs_yotei!kadou
                 kyk2 = yotei_data(i, 4)  'rs_yotei!kingaku
                ' kyk3 = rs_yotei!kaishi
                
                 sql = "delete * from keiyaku where keiyakuid ='" & yotei_data(i, 2) & "'"
                On Error GoTo errsakujo
                 db.Execute sql, dbSQLPassThrough
                On Error GoTo 0
                
                 If a_shukeiyaku_sakujo(s_kojinid, kyk2, kyk) = False Then
                     MsgBox "料金の更新に失敗しました。"
                      yotejikkou = -1
                     Exit Function
                 End If
               
            
            Else   'If yotei_data(i, 1) = "1" Then '休止1　中止2
                '変更
                Dim rs_yotei2  As Recordset, m_kin As String, m_shurui As String, hkingaku As String

                '現在のSTと価格を調べる
                sql = "SELECT * FROM keiyaku  WHERE keiyakuid ='" & yotei_data(i, 2) & "'"

                If FcSQlGet(rs_yotei2, sql, PrMsg) = True Then
                        
                         rs_yotei2.MoveFirst
                        
                        m_kin = rs_yotei2!kingaku
                        m_shurui = rs_yotei2!kadou
                        
                        rs_yotei2.Close
                        
                    
                        Dim hkudou As String
                        If yotei_data(i, 1) = "1" Then
                            hkudou = "1"
                        ElseIf yotei_data(i, 1) = "2" Then
                            hkudou = "2"
                        End If
                        
                        hkingaku = yotei_data(i, 4) 'rs_yotei!kingaku
                        
                        If m_shurui <> hkudou Then
                        
                            sql = "UPDATE keiyaku SET kadou = '" & hkudou & "'" & _
                               " WHERE keiyakuid = '" & yotei_data(i, 2) & "'"
                
                                On Error GoTo errk3
                                     db.Execute sql, dbSQLPassThrough
                                On Error GoTo 0
                                            
                             If a_shukeiyaku_henkou(s_kojinid, m_kin, hkingaku, m_shurui, hkudou) = False Then
                                MsgBox "料金の更新に失敗しました。"
                                 yotejikkou = -1
                                 Exit Function
                            End If
                        End If

                End If
                
            
            
            End If
            
            'ST更新
             On Error GoTo errsakujo2
                 sql = "UPDATE yotei SET st = '1'" & _
                              " WHERE yoteiid = '" & yotei_data(i, 9) & "'"
                 db.Execute sql, dbSQLPassThrough
             On Error GoTo 0
             
    
    Else  '"追加"
            Dim tkaishi As String, tkingaku As String, tkudou As String, shime As String
            
            tkaishi = Format(Date, "yyyymmdd")
            tkingaku = yotei_data(i, 4)
            tkudou = "0"
            shime = "0"
            
            If yotei_data(i, 1) = "0" Then ' "主契約"
                
                If shukeiyaku_touroku(s_kojinid, yotei_data(i, 2), yotei_data(i, 5), yotei_data(i, 4), tkaishi, tkudou, shime) = 1 Then
                
                    If a_shukeiyaku_koushin(s_kojinid, tkaishi, tkingaku, tkudou, shime) = False Then
                        MsgBox "料金の更新に失敗しました。"
                         yotejikkou = -1
                        Exit Function
                    Else
                          'ST更新
                            On Error GoTo errsakujo2
                                sql = "UPDATE yotei SET st = '1'" & _
                                             " WHERE yoteiid = '" & yotei_data(i, 9) & "'"
                                db.Execute sql, dbSQLPassThrough
                            On Error GoTo 0
                    End If
                End If
                
    
             Else '"その也"
             
                If sonotakeiyaku_touroku(s_kojinid, yotei_data(i, 2), tkingaku, tkaishi) = 1 Then
                    If a_sonota_touroku(s_kojinid, tkingaku, tkudou) = False Then
                       MsgBox "料金の更新に失敗しました。"
                        yotejikkou = -1
                       Exit Function
                    Else
                        s_keizoku = 0
                        If s_nentsuki = yotei_data(i, 8) Then
                            s_keizoku = 1
                        End If
                        'ST更新
                        If s_keizoku = 1 Then
                            On Error GoTo errsakujo2
                                sql = "UPDATE yotei SET st = '1'" & _
                                             " WHERE yoteiid = '" & yotei_data(i, 9) & "'"
                                db.Execute sql, dbSQLPassThrough
                            On Error GoTo 0
                        End If
                    End If
                 
                 End If
            End If
            
            
    End If
    
    
     
     
                 
Next


Exit Function

errsakujo:
   MsgBox "契約追加の更新に失敗しました。"
   yotejikkou = -1
                     Exit Function
errsakujo2:
    MsgBox "契約削除の更新に失敗しました。"
    yotejikkou = -1
                     Exit Function

errk3:
    MsgBox "契約変更の更新に失敗しました。"
    yotejikkou = -1
                     Exit Function


End Function

Function sonotakeiyaku_touroku(s_kojinid As String, s_shurui As String, s_kingaku As String, s_kaishibi As String) As Integer

Dim tid As String, tshurui As String, tkojinid As String, tkingaku As String
Dim tkaishi As String, tkudou As String, ttid As String


    sonotakeiyaku_touroku = 0
    
            If soukatsu(0, "renban", 6, "", 1, 1) = False Then
                    MsgBox "設定の読み込みに失敗しました。"
                    sonotakeiyaku_touroku = -1
                    Exit Function
            End If
            If reg = "0" Then
                reg = ""
            End If
            If reg = vbNullString Then
                tid = "0000001"
            Else
                tid = Format$(reg, "000000#")
            End If
        
            tshurui = s_shurui ' Mid(.cmbsonotakeiyaku.Text, 1, 3)
            
           
            tkojinid = s_kojinid 'frmtouroku2.lblid.Caption
            
            tkingaku = s_kingaku ' .txtkeiyakukin.Text
            
           
            tkaishi = s_kaishibi ' Format(.txtkaishibi.Text, "yyyymmdd")
           
            tkudou = "0"
            
        sql = "INSERT INTO keiyaku2 VALUES('" & tid & "','" & tkojinid & "'," & _
                       "'" & tshurui & "','" & tkaishi & "','" & tkudou & "','" & tkingaku & "')"
On Error GoTo errk5
        db.Execute sql, dbSQLPassThrough
On Error GoTo 0
        ttid = CStr(CLng(tid) + 1)
        If soukatsu(1, "renban", 6, ttid, 1, 1) = False Then
                MsgBox "設定の書き込みに失敗しました。"
                sonotakeiyaku_touroku = -1
                Exit Function
        End If
        
        sonotakeiyaku_touroku = 1
     Exit Function
     
errk5:
    MsgBox "設定の書き込みに失敗しました。"
                sonotakeiyaku_touroku = -1
                Exit Function
     
     

End Function

Function shukeiyaku_touroku(s_kojinid As String, s_shurui As String, s_kuchisuu As String, s_kingaku As String, s_kashibi As String, s_kubou As String, s_shime As String) As Integer

        Dim tid As String, tshurui As String, tkojinid As String, tkuchisuu As String, tkingaku As String, tkaishi As String
        Dim tkudou As String, shime As String, ttid As String

            If soukatsu(0, "renban", 4, "", 1, 1) = False Then
                    MsgBox "設定の読み込みに失敗しました。"
                    shukeiyaku_touroku = -1
                    Exit Function
            End If
            If reg = "0" Then
                reg = ""
            End If
    
            If reg = vbNullString Then
                tid = "0000001"
            Else
                tid = Format$(reg, "000000#")
            End If
            

            tshurui = s_shurui 'Mid(.cmbkeiyakushurui.Text, 1, 3)
            tkojinid = s_kojinid  'frmtouroku2.lblid.Caption
            '重複の確認
            sql = "select*from keiyaku where kojinid ='" & tkojinid & "' and shurui ='" & tshurui & "' and kadou='0'"
            If FcSQlGet(rs, sql, PrMsg) = True Then
                MsgBox "同じ契約がすでに登録されています。契約者ID:" & tkojinid
                shukeiyaku_touroku = 0
                rs.Close
                Exit Function
            End If
            
            tkuchisuu = s_kuchisuu ' .txtkeiyakusuu.Text
            
            tkingaku = s_kingaku ' .txtkeiyakukin.Text
            
            tkaishi = s_kashibi ' Format(.txtkaishibi.Text, "yyyymmdd")
            
            tkudou = s_kubou
           
            shime = s_shime
           
           
        sql = "INSERT INTO keiyaku VALUES('" & tid & "','" & tkojinid & "'," & _
                       "'" & tshurui & "','" & tkaishi & "','" & tkudou & "'," & _
                       "'" & tkuchisuu & "','" & tkingaku & "','" & shime & "','0')"
        
On Error GoTo errk4
        db.Execute sql, dbSQLPassThrough
On Error GoTo 0
        ttid = CStr(CLng(tid) + 1)
        If soukatsu(1, "renban", 4, ttid, 1, 1) = False Then
                MsgBox "設定の書き込みに失敗しました。"
                shukeiyaku_touroku = -1
                Exit Function
        End If
    
    shukeiyaku_touroku = 1
    
Exit Function

errk4:

    MsgBox "データの書き込みに失敗しました。"
    shukeiyaku_touroku = -1
    Exit Function

End Function



Function cl_koushin(k_trshurui As String, k_kojin As String, k_kingaku As String, k_hiduke As String) As Integer

Dim tid As String, tkojinid As String, tkingaku As String, jurs As Recordset
Dim tkaishi As String, tkudou As String, tshurui As String, ttid As String

    
If soukatsu(0, "renban", 6, "", 1, 1) = False Then

        MsgBox "設定の読み込みに失敗しました。"
End If
If reg = "0" Then
    reg = ""
End If
If reg = vbNullString Then
    tid = "0000001"
Else
    tid = Format$(reg, "000000#")
End If

            tshurui = k_trshurui
            tkojinid = k_kojin
            tkingaku = k_kingaku
            tkaishi = k_hiduke
            tkudou = "0" '可能
        sql = "INSERT INTO keiyaku2 VALUES('" & tid & "','" & tkojinid & "'," & _
                       "'" & tshurui & "','" & tkaishi & "','" & tkudou & "','" & tkingaku & "')"
On Error GoTo errk5
        db.Execute sql, dbSQLPassThrough
On Error GoTo 0
        ttid = CStr(CLng(tid) + 1)
        If soukatsu(1, "renban", 6, ttid, 1, 1) = False Then
                MsgBox "設定の書き込みに失敗しました。"
                cl_koushin = -1
                Exit Function
        End If
        If a_sonota_touroku(tkojinid, tkingaku, tkudou, 1) = False Then
            MsgBox "料金の更新に失敗しました。"
            cl_koushin = -1
            Exit Function
        End If
        
    
        cl_koushin = 1
Exit Function

errk5:
    MsgBox "契約の登録に失敗しました。"
    cl_koushin = -1
    Exit Function

End Function

Sub nhk_csv_fdd_check3()
Dim hikakuid As String, ipi As Long, logcount As Integer, logcount2 As Integer
Dim hikakuid2 As String, sql_n2 As String, rs_n2 As Recordset, sql_n3 As String, rs_n3 As Recordset
Dim m_name As String, m_kin As Long, hikakusurudata(4) As String, hikakusuu(2) As Integer, m_tsukihi As String
Dim sql_n4 As String, rs_n4 As Recordset

Dim reres

reres = MsgBox("NHKデータが登録済み個人のNHK契約データあるかを確認しますか？", vbYesNo)
If reres = vbNo Then
    Exit Sub
End If



If frmnhk2.lblok.Caption = "OK" Then
    logcount = 1
    add_err2 1, 0, "行数", "内容", "NHK側", "CATV側"
'各行のデータをDBと比較
    For ipi = 1 To frmnhk2.gridnhk.Rows - 1
        frmnhk2.gridnhk.Row = ipi
        'IDがあるか
        frmnhk2.gridnhk.Col = 8
        hikakuid = Trim(frmnhk2.gridnhk.Text)
        frmnhk2.gridnhk.Col = 6
        m_name = Trim(frmnhk2.gridnhk.Text)
        frmnhk2.gridnhk.Col = 14
        m_kin = CLng(frmnhk2.gridnhk.Text)
        frmnhk2.gridnhk.Col = 38
        m_tsukihi = Trim(frmnhk2.gridnhk.Text)
        
        If Len(hikakuid) <> 19 Then
            add_err2 logcount + 2, logcount, CStr(ipi), "CATV施設コードが19桁ありません。", hikakuid, ""
            logcount = logcount + 1
        Else
            '契約
            hikakuid2 = Right(hikakuid, 5)
            sql_n2 = "select * from kojin where kojinid='" & hikakuid2 & "'"
            If FcSQlGet(rs_n2, sql_n2, PrMsg) = False Then
                add_err2 logcount + 2, logcount, CStr(ipi), "CATV施設コードが登録されていません。ID:" & hikakuid2 & " " & m_name, "", ""
                logcount = logcount + 1
            Else
                sql_n3 = "select * from nhk_m where kojinid='" & hikakuid2 & "'"
                If FcSQlGet(rs_n3, sql_n3, PrMsg) = False Then
                    add_err2 logcount + 2, logcount, CStr(ipi), "NHK契約が登録されていません。ID:" & hikakuid2 & " " & m_name, "", ""
                    logcount = logcount + 1
                Else
                    If Mid(rs_n3!jikai, 1, 6) = m_tsukihi Then
                        add_err2 logcount + 2, logcount, CStr(ipi), "NHK契約の次回日が不正です。ID:" & hikakuid2 & " " & m_name, m_tsukihi, rs_n3!jikaibi
                        logcount = logcount + 1
                    End If
                    If CLng(rs_n3!kingaku) <> m_kin Then
                        add_err2 logcount + 2, logcount, CStr(ipi), "NHK契約の金額が不正です。ID:" & hikakuid2 & " " & m_name, CStr(m_kin), rs_n3!kingaku
                        logcount = logcount + 1
                    End If
                
                End If
            End If
        End If
        

    Next ipi
    
    Dim keisuu(5) As Long, keikin(5) As Double, c_id As String, arunokana As Integer, c_kin As Double, c_name As String
    Dim hikakudata() As String
    
     keisuu(0) = 0
     keisuu(1) = 0
     keisuu(2) = 0
     keisuu(3) = 0
     keisuu(4) = 0
    keikin(0) = 0
    keikin(1) = 0
    keikin(2) = 0
    keikin(3) = 0
    keikin(4) = 0
    
    ReDim hikakudata(frmnhk2.gridnhk.Rows, 4) As String
    
    For ipi = 1 To frmnhk2.gridnhk.Rows - 1
            frmnhk2.gridnhk.Row = ipi
            frmnhk2.gridnhk.Col = 8
            hikakuid = Trim(frmnhk2.gridnhk.Text)
            hikakuid2 = Right(hikakuid, 5)
            
            hikakudata(ipi - 1, 1) = hikakuid2
            
            frmnhk2.gridnhk.Col = 6
            hikakudata(ipi - 1, 2) = Trim(frmnhk2.gridnhk.Text)
            
            frmnhk2.gridnhk.Col = 14
            hikakudata(ipi - 1, 3) = CLng(frmnhk2.gridnhk.Text)
        
    Next
    
    logcount2 = 1
    add_err3 1, 0, "ID", "内容", "NHK側", "CATV側"
    
    sql_n4 = "SELECT nhk_m.*, kojin.name" & _
            " FROM kojin RIGHT JOIN nhk_m ON kojin.kojinid = nhk_m.kojinid"
    If FcSQlGet(rs_n4, sql_n4, PrMsg) = True Then
        rs_n4.MoveFirst
        Do Until rs_n4.EOF
            keisuu(4) = keisuu(4) + 1
            keikin(4) = keikin(4) + CDbl(rs_n4!kingaku)
            Select Case CDbl(rs_n4!kingaku)
            Case 11890
                keisuu(0) = keisuu(0) + 1
            Case 5345
                keisuu(1) = keisuu(1) + 1
            Case 5195
                keisuu(2) = keisuu(2) + 1
            Case Else
                keisuu(3) = keisuu(3) + 1
                keikin(3) = keikin(3) + CDbl(rs_n4!kingaku)
            End Select
            c_kin = CDbl(rs_n4!kingaku)
            c_id = rs_n4!kojinid
            c_name = rs_n4!Name
            arunokana = 0
            For ipi = 1 To frmnhk2.gridnhk.Rows - 1
                If c_id = hikakudata(ipi - 1, 1) Then
                        m_name = hikakudata(ipi - 1, 2)
                        m_kin = CLng(hikakudata(ipi - 1, 3))
                        If c_kin <> m_kin Then
                            add_err3 logcount2 + 2, logcount2, c_id, "登録済み個人データの金額が不正です。" & c_name, CStr(m_kin), CStr(c_kin)
                            logcount2 = logcount2 + 1
                        End If
                    arunokana = 1
                    Exit For
                End If
            Next
            
            If arunokana = 0 Then
                        add_err3 logcount2 + 2, logcount2, c_id, "登録済み個人データがありません。" & c_name, "", ""
                        logcount2 = logcount2 + 1
            End If
            
            
            
            
            
            rs_n4.MoveNext
         Loop
         rs_n4.Close
         keikin(0) = keisuu(0) * 11890
         keikin(1) = keisuu(1) * 5345
         keikin(2) = keisuu(2) * 5195
    End If
    
    
    If logcount = 1 Then
        MsgBox "エラーはありませんでした。"
    End If
        
        frmnhk2.lbl1m.Caption = keisuu(0)
        frmnhk2.lbl2m.Caption = keisuu(1)
        frmnhk2.lbl3m.Caption = keisuu(2)
        frmnhk2.lbl4m.Caption = keisuu(3)
        frmnhk2.lbl5m.Caption = keisuu(4)
        frmnhk2.lbl11m.Caption = keikin(0)
        frmnhk2.lbl22m.Caption = keikin(1)
        frmnhk2.lbl33m.Caption = keikin(2)
        frmnhk2.lbl44m.Caption = keikin(3)
        frmnhk2.lbl55m.Caption = keikin(4)
        
        frmnhk2.gridkekka2.Rows = logcount
        frmnhk2.fhikakukekka2.Visible = True
        frmnhk2.gridkekka2.ColWidth(0) = 1200
        frmnhk2.gridkekka2.ColWidth(1) = 1200
        frmnhk2.gridkekka2.ColWidth(2) = 7000
        frmnhk2.gridkekka2.ColWidth(3) = 2000
        frmnhk2.gridkekka2.ColWidth(4) = 2000
    
        frmnhk2.gridkekka3.Rows = logcount2
        frmnhk2.gridkekka3.ColWidth(0) = 1200
        frmnhk2.gridkekka3.ColWidth(1) = 1200
        frmnhk2.gridkekka3.ColWidth(2) = 7000
        frmnhk2.gridkekka3.ColWidth(3) = 2000
        frmnhk2.gridkekka3.ColWidth(4) = 2000
    
Else
    MsgBox "データを読み込んでから、実行してください。"
    Exit Sub
End If



    

End Sub
Sub add_err(srows As Integer, srow As Integer, s1 As String, s2 As String, s3 As String, s4 As String)

With frmnhk2
    .gridkekka.Rows = srows
    .gridkekka.Row = srow
    .gridkekka.Col = 0
    If srow <> 0 Then
        .gridkekka.Text = srow - 1
    End If
    .gridkekka.Col = 1
    .gridkekka.Text = s1
    .gridkekka.Col = 2
    .gridkekka.Text = s2
    .gridkekka.Col = 3
    .gridkekka.Text = s3
    .gridkekka.Col = 4
    .gridkekka.Text = s4
End With

End Sub
Sub add_err2(srows As Integer, srow As Integer, s1 As String, s2 As String, s3 As String, s4 As String)

With frmnhk2
    .gridkekka2.Rows = srows
    .gridkekka2.Row = srow
    .gridkekka2.Col = 0
    If srow <> 0 Then
        .gridkekka2.Text = srow
    End If
    .gridkekka2.Col = 1
    .gridkekka2.Text = s1
    .gridkekka2.Col = 2
    .gridkekka2.Text = s2
    .gridkekka2.Col = 3
    .gridkekka2.Text = s3
    .gridkekka2.Col = 4
    .gridkekka2.Text = s4
End With

End Sub
Sub add_err3(srows As Integer, srow As Integer, s1 As String, s2 As String, s3 As String, s4 As String)

With frmnhk2
    .gridkekka3.Rows = srows
    .gridkekka3.Row = srow
    .gridkekka3.Col = 0
    If srow <> 0 Then
        .gridkekka3.Text = srow
    End If
    .gridkekka3.Col = 1
    .gridkekka3.Text = s1
    .gridkekka3.Col = 2
    .gridkekka3.Text = s2
    .gridkekka3.Col = 3
    .gridkekka3.Text = s3
    .gridkekka3.Col = 4
    .gridkekka3.Text = s4
End With

End Sub

Sub nhk_csv_fdd_check4()
Dim hikakuid As String, ipi As Long, logcount As Integer
Dim hikakuid2 As String, sql_n2 As String, rs_n2 As Recordset
Dim m_name As String, m_tel As String, hikakusurudata(6) As String, hikakusuu(3) As Integer


Dim reres


reres = MsgBox("NHKデータが登録済み個人データと比較しますか？", vbYesNo)
If reres = vbNo Then
    Exit Sub
End If


If frmnhk2.lblok.Caption = "OK" Then
    logcount = 1
    add_err 1, 0, "行数", "内容", "NHK側", "CATV側"
'各行のデータをDBと比較
    For ipi = 1 To frmnhk2.gridnhk.Rows - 1
        frmnhk2.gridnhk.Row = ipi
        'IDがあるか
        frmnhk2.gridnhk.Col = 8
        hikakuid = Trim(frmnhk2.gridnhk.Text)
        frmnhk2.gridnhk.Col = 6
        m_name = Trim(frmnhk2.gridnhk.Text)
        frmnhk2.gridnhk.Col = 9
        m_tel = Trim(frmnhk2.gridnhk.Text)
        frmnhk2.gridnhk.Col = 10
        m_tel = m_tel & "-" & Trim(frmnhk2.gridnhk.Text)
        frmnhk2.gridnhk.Col = 11
        m_tel = m_tel & "-" & Trim(frmnhk2.gridnhk.Text)
        
        If Len(hikakuid) <> 19 Then
            add_err logcount + 2, logcount, CStr(ipi), "CATV施設コードが19桁ありません。ID:" & hikakuid, "", ""
            logcount = logcount + 1
        Else
            '名前と電話と住所
            hikakuid2 = Right(hikakuid, 5)
            sql_n2 = "select * from kojin where kojinid='" & hikakuid2 & "'"
            If FcSQlGet(rs_n2, sql_n2, PrMsg) = False Then
                add_err logcount + 2, logcount, CStr(ipi), "CATV施設コードが登録されていません。ID:" & hikakuid2, "", ""
                logcount = logcount + 1
            Else
                If m_name <> Trim(rs_n2!Name) Then
                    If InStr(1, m_name, "　") <> 0 Then
                        If InStr(1, Trim(rs_n2!Name), " ") <> 0 Then
                            hikakusuu(0) = InStr(1, m_name, "　")
                            hikakusurudata(0) = Mid(m_name, 1, hikakusuu(0) - 1)
                            hikakusurudata(1) = Mid(m_name, hikakusuu(0) + 1)
                            hikakusuu(1) = InStr(1, Trim(rs_n2!Name), " ")
                            hikakusurudata(2) = Mid(Trim(rs_n2!Name), 1, hikakusuu(1) - 1)
                            hikakusurudata(3) = Mid(Trim(rs_n2!Name), hikakusuu(1) + 1)
                            If hikakusurudata(0) & hikakusurudata(1) <> hikakusurudata(2) & hikakusurudata(3) Then
                                If m_name <> Trim(rs_n2!nhkname) Then
                                    If InStr(1, Trim(rs_n2!nhkname), " ") <> 0 Then
                                        hikakusuu(2) = InStr(1, Trim(rs_n2!nhkname), " ")
                                        hikakusurudata(4) = Trim(Mid(Trim(rs_n2!nhkname), 1, hikakusuu(1) - 1))
                                        hikakusurudata(5) = Trim(Mid(Trim(rs_n2!nhkname), hikakusuu(1) + 1))
                                        If hikakusurudata(0) & hikakusurudata(1) <> hikakusurudata(4) & hikakusurudata(5) Then
                                            add_err logcount + 2, logcount, CStr(ipi), "名前が違います。ID:" & hikakuid2, m_name, Trim(rs_n2!nhkname)
                                            logcount = logcount + 1
                                        End If
                                    Else
                                        add_err logcount + 2, logcount, CStr(ipi), "名前が違います。ID:" & hikakuid2, m_name, Trim(rs_n2!nhkname)
                                        logcount = logcount + 1
                                    End If
                                 End If
                            End If
                        Else
                            add_err logcount + 2, logcount, CStr(ipi), "名前が違います。ID:" & hikakuid2, m_name, Trim(rs_n2!nhkname)
                            logcount = logcount + 1
                        End If
                    Else
                        add_err logcount + 2, logcount, CStr(ipi), "名前が違います。ID:" & hikakuid2, m_name, Trim(rs_n2!nhkname)
                        logcount = logcount + 1
                    End If
                End If
                If m_tel <> rs_n2!tel1 Then
                    add_err logcount + 2, logcount, CStr(ipi), "電話番号が違います。ID:" & hikakuid2, m_tel, rs_n2!tel1
                    logcount = logcount + 1
                End If
                
                
                
                rs_n2.Close
            End If
        End If
        


    Next ipi
    
    If logcount = 1 Then
        MsgBox "エラーはありませんでした。"
        Exit Sub
    Else
         frmnhk2.gridkekka.Rows = logcount
        frmnhk2.fhikakukekka.Visible = True
        frmnhk2.gridkekka.ColWidth(0) = 1200
        frmnhk2.gridkekka.ColWidth(1) = 1200
        frmnhk2.gridkekka.ColWidth(2) = 5000
        frmnhk2.gridkekka.ColWidth(3) = 3000
        frmnhk2.gridkekka.ColWidth(4) = 3000
    End If
    
    
Else
    MsgBox "データを読み込んでから、実行してください。"
    Exit Sub
End If



End Sub

Sub nhk_csv_fdd_check()
Dim sql_sentaku As String, rs_sentaku As Recordset, sql_ch35 As String, rs_ch35 As Recordset
Dim csv_data(), datasuu2 As Long, sususu As Long, rs_s5 As Recordset, sql_s5 As String
Dim seikyuu_data(), datasuu3 As Long, sususu2 As Long, sumi As Integer, csv_5_count As Long

Dim reres


reres = MsgBox("個人情報に登録済みのNHKの契約データが、すでに請求されているかを確認しますか？", vbYesNo)
If reres = vbNo Then
    Exit Sub
End If


sql_sentaku = "SELECT nhk_m.kojinid, kojin.name, kojin.furigana, kojin.chiikiid," & _
            "kojin.juusho1, kojin.juusho2, kojin.tel1, kojin.tel2, kojin.keitai," & _
            "nhk_m.shurui, nhk_m.n1, nhk_m.n2, nhk_m.jikai,nhk_m.zenkai," & _
            "nhk_m.kingaku,kojin.juutaku,kojin.cellno,kojin.bikou,kojin.setaisuu,kojin.tvsuu" & _
            " FROM kojin RIGHT JOIN nhk_m ON kojin.kojinid = nhk_m.kojinid"

If FcSQlGet(rs_sentaku, sql_sentaku, PrMsg) = False Then
        MsgBox "契約内容データ（CSV）が存在しません。"
        Exit Sub
Else
    datasuu2 = rs_sentaku.RecordCount
    sususu = 0
        ReDim Preserve csv_data(27, datasuu2)
    rs_sentaku.MoveFirst
    Do Until rs_sentaku.EOF
            csv_data(0, sususu) = Format(sususu, "00000")
            csv_data(1, sususu) = rs_sentaku!kojinid
            csv_data(2, sususu) = rs_sentaku!Name
            csv_data(3, sususu) = rs_sentaku!Furigana
            csv_data(4, sususu) = rs_sentaku!chiikiid
            csv_data(5, sususu) = rs_sentaku!juusho1
            csv_data(6, sususu) = rs_sentaku!juusho2
            csv_data(7, sususu) = rs_sentaku!tel1
            csv_data(8, sususu) = rs_sentaku!tel2
            csv_data(9, sususu) = rs_sentaku!keitai
            csv_data(10, sususu) = rs_sentaku!shurui
            Select Case rs_sentaku!shurui
            Case "0"
                csv_data(11, sususu) = "口座振替１ヶ月"
            Case "1"
                csv_data(11, sususu) = "口座振替２ヶ月"
            Case "2"
                csv_data(11, sususu) = "口座振替３ヶ月"
            Case "3"
                csv_data(11, sususu) = "口座振替４ヶ月"
            Case "4"
                csv_data(11, sususu) = "口座振替５ヶ月"
            Case "5"
                csv_data(11, sususu) = "口座振替６ヶ月"
            Case "6"
                csv_data(11, sususu) = "口座振替６ヶ月ハンデ"
            Case "7"
                csv_data(11, sususu) = "口座振替６ヶ月家族割"
            End Select
            csv_data(12, sususu) = ""
            csv_data(13, sususu) = rs_sentaku!n1
            csv_data(14, sususu) = rs_sentaku!kingaku
            csv_data(15, sususu) = rs_sentaku!juutaku
            csv_data(16, sususu) = rs_sentaku!cellno
            csv_data(17, sususu) = rs_sentaku!bikou
            csv_data(18, sususu) = rs_sentaku!zenkai
            csv_data(19, sususu) = rs_sentaku!jikai
            csv_data(20, sususu) = rs_sentaku!setaisuu
            csv_data(21, sususu) = rs_sentaku!tvsuu
        sususu = sususu + 1
        rs_sentaku.MoveNext
    Loop
    rs_sentaku.Close
End If

    If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If
'未入金のデータ
sql_ch35 = "SELECT * FROM shousai" & _
        " WHERE keist='0'" & _
        " and keiid='5' and keishurui='3' order by kojinid"
        
If FcSQlGet2(rs_ch35, sql_ch35, PrMsg) = True Then
        datasuu3 = rs_ch35.RecordCount
        ReDim Preserve seikyuu_data(27, datasuu3)
        
        rs_ch35.MoveFirst
        sususu = 0
        Do Until rs_ch35.EOF
            seikyuu_data(0, sususu) = Format(sususu, "00000")
            seikyuu_data(1, sususu) = rs_ch35!kojinid
            sql_s5 = "select * from kojin where kojinid='" & rs_ch35!kojinid & "'"
            If FcSQlGet(rs_s5, sql_s5, PrMsg) = False Then
                MsgBox "TBL_shousaiの" & rs_ch35!shousaiid & "データが不正です。"
                Exit Sub
            Else
                seikyuu_data(2, sususu) = rs_s5!Name
                seikyuu_data(2, sususu) = rs_s5!Furigana
                rs_s5.Close
            End If
            seikyuu_data(10, sususu) = "5"
            seikyuu_data(14, sususu) = rs_ch35!keikin
            sususu = sususu + 1
            rs_ch35.MoveNext
        Loop
        rs_ch35.Close
Else
        MsgBox "請求内容データ（FDD）が存在しません。"
        Exit Sub
End If


'比較
MsgBox "CSVデータが請求済みデータにあるかを検証します。"
csv_5_count = 0
For sususu = 0 To datasuu2 - 1
    If csv_data(10, sususu) = "5" Then
        csv_5_count = csv_5_count + 1
        sumi = 0
        For sususu2 = 0 To datasuu3 - 1
            If csv_data(1, sususu) = seikyuu_data(1, sususu2) Then
                If csv_data(14, sususu) <> seikyuu_data(14, sususu2) Then
                    reres = MsgBox("登録済みデータと請求済みデータの契約金額が違います。続けますか？顧客ID：" & csv_data(1, sususu) & Space(2) & csv_data(2, sususu), vbYesNo)
                    If reres = vbNo Then
                        Exit Sub
                    End If
                End If
                sumi = 1
                Exit For
            End If
        Next
        If sumi = 0 Then
            reres = MsgBox("登録済みデータが請求予定データにありません。続けますか？顧客ID：" & csv_data(1, sususu) & Space(2) & csv_data(2, sususu), vbYesNo)
            If reres = vbNo Then
                Exit Sub
            End If
        End If
    End If
Next


'MsgBox "請求済みデータがCSVデータにあるかを検証します。"






MsgBox "検証終了。"


End Sub
Sub nhk_csv_fdd_check2()
Dim sql_sentaku As String, rs_sentaku As Recordset, sql_ch35 As String, rs_ch35 As Recordset
Dim csv_data(), datasuu2 As Long, sususu As Long, rs_s5 As Recordset, sql_s5 As String
Dim seikyuu_data(), datasuu3 As Long, sususu2 As Long, sumi As Integer, csv_5_count As Long
Dim reres


reres = MsgBox("個人情報に登録済みのNHKの契約データが、現在請求額に加算されているかを確認しますか？", vbYesNo)
If reres = vbNo Then
    Exit Sub
End If

sql_sentaku = "SELECT nhk_m.kojinid, kojin.name, kojin.furigana, kojin.chiikiid," & _
            "kojin.juusho1, kojin.juusho2, kojin.tel1, kojin.tel2, kojin.keitai," & _
            "nhk_m.shurui, nhk_m.n1, nhk_m.n2, nhk_m.jikai,nhk_m.zenkai," & _
            "nhk_m.kingaku,kojin.juutaku,kojin.cellno,kojin.bikou,kojin.setaisuu,kojin.tvsuu, kojin.f4" & _
            " FROM kojin RIGHT JOIN nhk_m ON kojin.kojinid = nhk_m.kojinid"

If FcSQlGet(rs_sentaku, sql_sentaku, PrMsg) = False Then
        MsgBox "契約内容データ（CSV）が存在しません。"
        Exit Sub
Else
    datasuu2 = rs_sentaku.RecordCount
    sususu = 0
        ReDim Preserve csv_data(27, datasuu2)
    rs_sentaku.MoveFirst
    Do Until rs_sentaku.EOF
            csv_data(0, sususu) = Format(sususu, "00000")
            csv_data(1, sususu) = rs_sentaku!kojinid
            csv_data(2, sususu) = rs_sentaku!Name
            csv_data(3, sususu) = rs_sentaku!Furigana
            csv_data(4, sususu) = rs_sentaku!chiikiid
            csv_data(5, sususu) = rs_sentaku!juusho1
            csv_data(6, sususu) = rs_sentaku!juusho2
            csv_data(7, sususu) = rs_sentaku!tel1
            csv_data(8, sususu) = rs_sentaku!tel2
            csv_data(9, sususu) = rs_sentaku!keitai
            csv_data(10, sususu) = rs_sentaku!shurui
            Select Case rs_sentaku!shurui
            Case "0"
                csv_data(11, sususu) = "口座振替１ヶ月"
            Case "1"
                csv_data(11, sususu) = "口座振替２ヶ月"
            Case "2"
                csv_data(11, sususu) = "口座振替３ヶ月"
            Case "3"
                csv_data(11, sususu) = "口座振替４ヶ月"
            Case "4"
                csv_data(11, sususu) = "口座振替５ヶ月"
            Case "5"
                csv_data(11, sususu) = "口座振替６ヶ月"
            Case "6"
                csv_data(11, sususu) = "口座振替６ヶ月ハンデ"
            Case "7"
                csv_data(11, sususu) = "口座振替６ヶ月家族割"
            End Select
            csv_data(12, sususu) = ""
            csv_data(13, sususu) = rs_sentaku!n1
            csv_data(14, sususu) = rs_sentaku!kingaku
            csv_data(15, sususu) = rs_sentaku!juutaku
            csv_data(16, sususu) = rs_sentaku!cellno
            csv_data(17, sususu) = rs_sentaku!bikou
            csv_data(18, sususu) = rs_sentaku!zenkai
            csv_data(19, sususu) = rs_sentaku!jikai
            csv_data(20, sususu) = rs_sentaku!setaisuu
            csv_data(21, sususu) = rs_sentaku!tvsuu
            csv_data(22, sususu) = rs_sentaku!f4
        sususu = sususu + 1
        rs_sentaku.MoveNext
    Loop
    rs_sentaku.Close
End If


'比較
MsgBox "データが請求予定データにあるかを検証します。"
csv_5_count = 0
For sususu = 0 To datasuu2 - 1
    If csv_data(10, sususu) = "5" Then
        csv_5_count = csv_5_count + 1
        If csv_data(14, sususu) <> csv_data(22, sususu) Then
            reres = MsgBox("登録済みデータが請求予定データにありません。続けますか？顧客ID：" & csv_data(1, sususu) & Space(2) & csv_data(2, sususu), vbYesNo)
            If reres = vbNo Then
                Exit Sub
            End If
        End If
    End If
Next


'MsgBox "請求済みデータがCSVデータにあるかを検証します。"






MsgBox "検証終了。"


End Sub
Sub csv_csv4(shutsu_path As String)

Dim datasuu As Integer, csv_data()


    datasuu = frmkari.dgvkari.Rows
    
    
        ReDim Preserve csv_data(22, datasuu)
        
        csv_data(0, 0) = "NO"
        csv_data(1, 0) = "個人 ID"
        csv_data(2, 0) = "個人名"
        csv_data(3, 0) = "個人フリガナ"
        csv_data(4, 0) = "住所"
        csv_data(5, 0) = "電話番号"
        csv_data(6, 0) = "金額"
        csv_data(7, 0) = "請求"
        csv_data(8, 0) = "B"
        csv_data(9, 0) = "C"
        csv_data(10, 0) = "前回入金日"
        csv_data(11, 0) = "銀行名"
        csv_data(12, 0) = "最終結果コード"
        csv_data(13, 0) = "郵便番号"
        
  
 With frmkari.dgvkari
 For i = 0 To datasuu - 1
        .Row = i
        
            csv_data(0, i) = Format(i + 1, "00000")
        .Col = 1
            csv_data(1, i) = .Text 'rs_sentaku!kojinid
        .Col = 2
            csv_data(2, i) = .Text ' rs_sentaku!Name
       
        .Col = 3
            csv_data(6, i) = .Text 'rs_sentaku!soukei
        
            csv_data(7, i) = "口座扱"
        .Col = 4
            csv_data(11, i) = .Text '  rs_sentaku!bankname
        .Col = 5
                csv_data(12, i) = .Text '"資金不足"
         .Col = 7
            csv_data(3, i) = .Text 'rs_sentaku!Furigana
          .Col = 8
                csv_data(13, i) = .Text ' Format(rs_sentaku!yuubin, "@@@-@@@@")
        .Col = 9
            csv_data(4, i) = .Text 'rs_sentaku!juusho1 & rs_sentaku!juusho2
        .Col = 10
            csv_data(5, i) = .Text 'rs_sentaku!tel1
       .Col = 11
            csv_data(8, i) = .Text 'rs_sentaku!f2
        .Col = 12
            csv_data(9, i) = .Text 'rs_sentaku!fkaisuu
        .Col = 13
            csv_data(10, i) = .Text ' Format(rs_sentaku!maebi, "@@@@/@@/@@")
       
      
Next
End With

Dim writecounter As Integer, writecounter2 As Integer
For writecounter = 1 To datasuu
        For writecounter2 = 1 To 13
          If IsNull(csv_data(writecounter2, writecounter)) Then
            csv_data(writecounter2, writecounter) = ""
          End If
        Next writecounter2
Next writecounter
Open shutsu_path For Output Access Write As 1
For writecounter = 0 To datasuu - 1
        Write #1, csv_data(0, writecounter), csv_data(1, writecounter), csv_data(2, writecounter), csv_data(3, writecounter), csv_data(13, writecounter), csv_data(4, writecounter), csv_data(5, writecounter), csv_data(6, writecounter), csv_data(7, writecounter), csv_data(8, writecounter), csv_data(9, writecounter), csv_data(10, writecounter), csv_data(11, writecounter), csv_data(12, writecounter)
Next

Close #1

MsgBox "指定データのエクスポートが完了しました。"
End Sub

Sub csv_csv8(shutsu_path As String)

Dim datasuu As Integer, csv_data()


    datasuu = frmshukei.dgvkari.Rows
    
    
        ReDim Preserve csv_data(22, datasuu)
        
        csv_data(0, 0) = "種類"
        csv_data(1, 0) = "詳細"
        csv_data(2, 0) = "件数"
  
 With frmshukei.dgvkari
 For i = 0 To datasuu - 1
        .Row = i
        .Col = 0
            csv_data(0, i) = .Text 'Format(i + 1, "00000")
        .Col = 1
            csv_data(1, i) = .Text 'rs_sentaku!kojinid
        .Col = 2
            csv_data(2, i) = .Text ' rs_sentaku!Name
       
       
Next
End With

Dim writecounter As Integer, writecounter2 As Integer
For writecounter = 1 To datasuu
        For writecounter2 = 0 To 2
          If IsNull(csv_data(writecounter2, writecounter)) Then
            csv_data(writecounter2, writecounter) = ""
          End If
        Next writecounter2
Next writecounter
Open shutsu_path For Output Access Write As 1
For writecounter = 0 To datasuu - 1
        Write #1, csv_data(0, writecounter), csv_data(1, writecounter), csv_data(2, writecounter)
Next

Close #1

MsgBox "指定データのエクスポートが完了しました。"
End Sub


Sub log_sagyou(s_filepath As String, s_str As String, Optional s_kojinid As String = "", Optional s_name As String = "")

Dim s_nichiji As String
Dim s_nichiji2 As String

s_nichiji = Format(Date, "yyyymmdd") & " " & Format(Time, "hhmmss")

Dim logfile As String

logfile = s_filepath & "作業記録_" & Format(Date, "yyyymm") & "_log.txt"

'Open logfile For Output Access Write As 1
Open logfile For Append As 1


Write #1, s_nichiji, s_kojinid, s_name, s_str

Close #1


End Sub

Sub csv_csv3(shutsu_path As String)
Dim sql_sentaku As String, rs_sentaku As Recordset, sususu As Long, csv_data()
Dim datasuu As Long, writecounter As Long, kensakukoumokusuu As Integer, writecounter2 As Long
Dim sql_sentaku2 As String, rs_sentaku2 As Recordset, datasuu2 As Long

sql_sentaku = frmsouken.lblsql.Caption

If FcSQlGet(rs_sentaku, sql_sentaku, PrMsg) = True Then
    datasuu = rs_sentaku.RecordCount
        ReDim Preserve csv_data(22, datasuu)
        csv_data(0, 0) = "NO"
        csv_data(1, 0) = "個人 ID"
        csv_data(2, 0) = "個人名"
        csv_data(3, 0) = "個人フリガナ"
        csv_data(4, 0) = "住所"
        csv_data(5, 0) = "電話番号"
        csv_data(6, 0) = "金額"
        csv_data(7, 0) = "請求"
        csv_data(8, 0) = "B"
        csv_data(9, 0) = "C"
        csv_data(10, 0) = "前回入金日"
        csv_data(11, 0) = "銀行名"
        csv_data(12, 0) = "最終結果コード"
        csv_data(13, 0) = "郵便番号"
        
    rs_sentaku.MoveFirst
    sususu = 1
    Do Until rs_sentaku.EOF
            csv_data(0, sususu) = Format(sususu, "00000")
            csv_data(1, sususu) = rs_sentaku!kojinid
            csv_data(2, sususu) = rs_sentaku!Name
            csv_data(3, sususu) = rs_sentaku!Furigana
            csv_data(4, sususu) = rs_sentaku!juusho1 & rs_sentaku!juusho2
            csv_data(5, sususu) = rs_sentaku!tel1
            csv_data(6, sususu) = rs_sentaku!soukei
            If rs_sentaku!seikyuu = "0" Then
                csv_data(7, sususu) = "口座扱"
            ElseIf rs_sentaku!seikyuu = "1" Then
                csv_data(7, sususu) = "請求書"
            Else
                csv_data(7, sususu) = "error"
            End If
            csv_data(8, sususu) = rs_sentaku!f2
            csv_data(9, sususu) = rs_sentaku!fkaisuu
            If Not IsNull(rs_sentaku!maebi) And Trim(rs_sentaku!maebi) <> "" Then
                csv_data(10, sususu) = Format(rs_sentaku!maebi, "@@@@/@@/@@")
            End If
            csv_data(11, sususu) = rs_sentaku!bankname
            If rs_sentaku!kekkacode = "1" Then
                csv_data(12, sususu) = "資金不足"
            ElseIf rs_sentaku!kekkacode = "2" Then
                csv_data(12, sususu) = "取引なし"
            ElseIf rs_sentaku!kekkacode = "3" Then
                csv_data(12, sususu) = "預金者の都合"
            ElseIf rs_sentaku!kekkacode = "4" Then
                csv_data(12, sususu) = "振替依頼書なし"
            ElseIf rs_sentaku!kekkacode = "8" Then
                csv_data(12, sususu) = "委託者の都合"
            ElseIf rs_sentaku!kekkacode = "9" Then
                csv_data(12, sususu) = "その他"
            Else
               csv_data(12, sususu) = "不明"
            End If
            If Len(Trim(rs_sentaku!yuubin)) = 7 Then
                csv_data(13, sususu) = Format(rs_sentaku!yuubin, "@@@-@@@@")
            Else
                csv_data(13, sususu) = Trim(rs_sentaku!yuubin)
            End If
        sususu = sususu + 1
        rs_sentaku.MoveNext
    Loop
    rs_sentaku.Close
Else
    MsgBox "指定データがありません。"
    Exit Sub
End If

For writecounter = 1 To datasuu
        For writecounter2 = 1 To 13
          If IsNull(csv_data(writecounter2, writecounter)) Then
            csv_data(writecounter2, writecounter) = ""
          End If
        Next writecounter2
Next writecounter
Open shutsu_path For Output Access Write As 1
For writecounter = 0 To datasuu + datasuu2
        Write #1, csv_data(0, writecounter), csv_data(1, writecounter), csv_data(2, writecounter), csv_data(3, writecounter), csv_data(13, writecounter), csv_data(4, writecounter), csv_data(5, writecounter), csv_data(6, writecounter), csv_data(7, writecounter), csv_data(8, writecounter), csv_data(9, writecounter), csv_data(10, writecounter), csv_data(11, writecounter), csv_data(12, writecounter)
Next

Close #1

MsgBox "指定データのエクスポートが完了しました。"
End Sub
Function csv_csv_cl(shutsu_path As String) As String

csv_csv_cl = ""

Dim sql_sentaku As String, rs_sentaku As Recordset, sususu As Long, csv_data()

Dim s_err_id As String

Dim datasuu As Long, writecounter As Long, kensakukoumokusuu As Integer, writecounter2 As Long
Dim sql_sentaku2 As String, rs_sentaku2 As Recordset
Dim sql_sentaku3 As String, rs_sentaku3 As Recordset

sql_sentaku = "SELECT keiyaku.kojinid, kojin.name, kojin.furigana, kojin.chiikiid,kojin.yuubin," & _
            "kojin.seikyuu, kojin.bankid, kojin.kouzakubun, kojin.kouzabangou, kojin.meigi," & _
            "kojin.juusho1, kojin.juusho2, kojin.tel1, kojin.tel2, kojin.keitai,kojin.cellid,kojin.hoshoukin," & _
            "keiyaku.shurui, shurui.shuruimei, keiyaku.kadou, keiyaku.kuchisuu," & _
            "keiyaku.kingaku,kojin.juutaku,kojin.cellno,kojin.bikou,kojin.setaisuu,kojin.tvsuu,kojin.h_koujibangou,kojin.h_zumen,kojin.h_koujikanryoubi,kojin.h_bikou" & _
            ",h_listbangou,cl_bangou,v_onu_adress,d_onu_adress,cl_open_date,cl_close_date,cl_keiyaku_name" & _
            " FROM shurui RIGHT JOIN (kojin RIGHT JOIN keiyaku ON kojin.kojinid = keiyaku.kojinid)" & _
            " ON shurui.shuruiid = keiyaku.shurui"



If FcSQlGet(rs_sentaku, sql_sentaku, PrMsg) = True Then
On Error GoTo eee1
    datasuu = rs_sentaku.RecordCount
        ReDim Preserve csv_data(30, datasuu)
        csv_data(0, 0) = "NO"
        csv_data(1, 0) = "個人ID"
        csv_data(2, 0) = "個人名"
        csv_data(3, 0) = "個人フリガナ"
        csv_data(4, 0) = "郵便番号"
        csv_data(5, 0) = "住所１"
        csv_data(6, 0) = "住所２"
        csv_data(7, 0) = "電話番号１"
        csv_data(8, 0) = "電話番号２"
        csv_data(9, 0) = "携帯電話番号"
        csv_data(10, 0) = "種類　ID"
        csv_data(11, 0) = "種類名"
        
        csv_data(12, 0) = "稼働状況"
        
        csv_data(13, 0) = "口数"
        csv_data(14, 0) = "金額"
        csv_data(15, 0) = "住宅区分"
        csv_data(16, 0) = "備考"
        csv_data(17, 0) = "世帯数"
        csv_data(18, 0) = "ＴＶ数"
        csv_data(19, 0) = "工区番号"
        csv_data(20, 0) = "図面"
        csv_data(21, 0) = "工事完了日"
         csv_data(22, 0) = "備考(光)"
         csv_data(23, 0) = "リスト番号"
         csv_data(24, 0) = "クロージャー番号"
         csv_data(25, 0) = "V-ONUアドレス"
         csv_data(26, 0) = "D-ONUアドレス"
          csv_data(27, 0) = "CL開通日"
        csv_data(28, 0) = "CL解約日"
        csv_data(29, 0) = "CL契約者名"
        
    rs_sentaku.MoveFirst
    sususu = 1
    Do Until rs_sentaku.EOF
            csv_data(0, sususu) = Format(sususu, "00000")
            csv_data(1, sususu) = rs_sentaku!kojinid
            s_err_id = rs_sentaku!kojinid
            If rs_sentaku!kojinid = "03111" Then
              '  MsgBox "HIT"
            End If
            csv_data(2, sususu) = rs_sentaku!Name
            csv_data(3, sususu) = rs_sentaku!Furigana
             If Trim(rs_sentaku!yuubin) = "" Then
                 csv_data(4, sususu) = ""
            Else
                csv_data(4, sususu) = Mid(Trim(rs_sentaku!yuubin), 1, 3) & "-" & Mid(Trim(rs_sentaku!yuubin), 4, 4)
            End If
            
            csv_data(5, sususu) = rs_sentaku!juusho1
            
            csv_data(5, sususu) = Replace(csv_data(5, sususu), "＊", "")
            
            csv_data(6, sususu) = rs_sentaku!juusho2
            csv_data(7, sususu) = rs_sentaku!tel1
            csv_data(8, sususu) = rs_sentaku!tel2
            csv_data(9, sususu) = rs_sentaku!keitai
            csv_data(10, sususu) = rs_sentaku!shurui
            csv_data(11, sususu) = rs_sentaku!shuruimei
            
            csv_data(12, sususu) = rs_sentaku!kadou
            
            csv_data(13, sususu) = rs_sentaku!kuchisuu
            csv_data(14, sususu) = rs_sentaku!kingaku
            csv_data(15, sususu) = rs_sentaku!juutaku
            csv_data(16, sususu) = rs_sentaku!bikou
            csv_data(17, sususu) = rs_sentaku!setaisuu
            csv_data(18, sususu) = rs_sentaku!tvsuu
            csv_data(19, sususu) = rs_sentaku!h_koujibangou
            csv_data(20, sususu) = rs_sentaku!h_zumen
            csv_data(21, sususu) = rs_sentaku!h_koujikanryoubi
            csv_data(22, sususu) = rs_sentaku!h_bikou
            
             If Trim(rs_sentaku!h_listbangou) = "" Then
                 csv_data(23, sususu) = ""
            Else
                csv_data(23, sususu) = Trim(rs_sentaku!h_listbangou)
            End If
            
             If Trim(rs_sentaku!cl_bangou) = "" Then
                 csv_data(24, sususu) = ""
            Else
                csv_data(24, sususu) = Trim(rs_sentaku!cl_bangou)
            End If
            
             If Trim(rs_sentaku!v_onu_adress) = "" Then
                 csv_data(25, sususu) = ""
            Else
                csv_data(25, sususu) = Trim(rs_sentaku!v_onu_adress)
            End If
            
            
             If Trim(rs_sentaku!d_onu_adress) = "" Then
                 csv_data(26, sususu) = ""
            Else
                csv_data(26, sususu) = Trim(rs_sentaku!d_onu_adress)
            End If
            
            If IsNull(rs_sentaku!cl_open_date) Then
                 csv_data(27, sususu) = ""
            Else
                csv_data(27, sususu) = Trim(rs_sentaku!cl_open_date)
            End If
            
             If IsNull(rs_sentaku!cl_close_date) Then
                 csv_data(28, sususu) = ""
            Else
                csv_data(28, sususu) = Trim(rs_sentaku!cl_close_date)
            End If
            
              If IsNull(rs_sentaku!cl_keiyaku_name) Then
                 csv_data(29, sususu) = ""
            Else
                csv_data(29, sususu) = Trim(rs_sentaku!cl_keiyaku_name)
            End If
            
            
           
            
        sususu = sususu + 1
        rs_sentaku.MoveNext
    Loop
    rs_sentaku.Close
On Error GoTo 0
End If



On Error GoTo eee3
For writecounter = 1 To datasuu
        For writecounter2 = 1 To 29
          If IsNull(csv_data(writecounter2, writecounter)) Then
            csv_data(writecounter2, writecounter) = ""
        Else
            csv_data(writecounter2, writecounter) = Replace(csv_data(writecounter2, writecounter), vbCr, "")
            csv_data(writecounter2, writecounter) = Replace(csv_data(writecounter2, writecounter), vbCrLf, "")
            csv_data(writecounter2, writecounter) = Replace(csv_data(writecounter2, writecounter), vbLf, "")
          End If
        Next writecounter2
Next writecounter
On Error GoTo 0

On Error GoTo eee4
Open shutsu_path For Output Access Write As 1
For writecounter = 0 To datasuu
        Write #1, csv_data(1, writecounter), csv_data(2, writecounter), csv_data(3, writecounter), csv_data(4, writecounter), csv_data(5, writecounter), csv_data(6, writecounter), csv_data(7, writecounter), csv_data(8, writecounter), csv_data(9, writecounter), csv_data(10, writecounter), csv_data(11, writecounter), csv_data(12, writecounter), csv_data(13, writecounter), csv_data(14, writecounter), csv_data(15, writecounter), csv_data(16, writecounter), csv_data(17, writecounter), csv_data(18, writecounter), csv_data(19, writecounter), csv_data(20, writecounter), csv_data(21, writecounter), csv_data(22, writecounter), csv_data(23, writecounter), _
        csv_data(24, writecounter), csv_data(25, writecounter), csv_data(26, writecounter), csv_data(27, writecounter), csv_data(28, writecounter), csv_data(29, writecounter)
Next

Close #1
On Error GoTo 0

MsgBox "指定データのエクスポートが完了しました。"

csv_csv_cl = shutsu_path

Exit Function

eee1:
MsgBox "個人データ（口座）の出力時にエラーが発生しました。"
Exit Function
eee3:
MsgBox "個人データの整理時にエラーが発生しました。"
Exit Function

eee4:
MsgBox "個人データの書き込み時にエラーが発生しました。"
Exit Function


End Function



Sub csv_csv2(shutsu_path As String)

Dim sql_sentaku As String, rs_sentaku As Recordset, sususu As Long, csv_data()

Dim s_err_id As String

Dim datasuu As Long, writecounter As Long, kensakukoumokusuu As Integer, writecounter2 As Long
Dim sql_sentaku2 As String, rs_sentaku2 As Recordset, datasuu2 As Long
Dim sql_sentaku3 As String, rs_sentaku3 As Recordset
sql_sentaku = "SELECT keiyaku.kojinid, kojin.name, kojin.furigana, kojin.chiikiid,kojin.yuubin," & _
            "kojin.seikyuu, kojin.bankid, kojin.kouzakubun, kojin.kouzabangou, kojin.meigi," & _
            "kojin.juusho1, kojin.juusho2, kojin.tel1, kojin.tel2, kojin.keitai,kojin.cellid,kojin.hoshoukin," & _
            "keiyaku.shurui, shurui.shuruimei, keiyaku.kadou, keiyaku.kuchisuu," & _
            "keiyaku.kingaku,kojin.juutaku,kojin.cellno,kojin.bikou,kojin.setaisuu,kojin.tvsuu,kojin.h_koujibangou,kojin.h_zumen,kojin.h_koujikanryoubi,kojin.h_bikou" & _
            ",h_listbangou,cl_bangou,v_onu_adress,d_onu_adress,cl_open_date,cl_close_date,cl_keiyaku_name" & _
            " FROM shurui RIGHT JOIN (kojin RIGHT JOIN keiyaku ON kojin.kojinid = keiyaku.kojinid)" & _
            " ON shurui.shuruiid = keiyaku.shurui"

If FcSQlGet(rs_sentaku, sql_sentaku, PrMsg) = True Then
On Error GoTo eee1
    datasuu = rs_sentaku.RecordCount
        ReDim Preserve csv_data(41, datasuu)
        csv_data(0, 0) = "NO"
        csv_data(1, 0) = "個人 ID"
        csv_data(2, 0) = "個人名"
        csv_data(3, 0) = "個人フリガナ"
        csv_data(4, 0) = "地域区分"
        csv_data(5, 0) = "住所１"
        csv_data(6, 0) = "住所２"
        csv_data(7, 0) = "電話番号１"
        csv_data(8, 0) = "電話番号２"
        csv_data(9, 0) = "携帯電話番号"
        csv_data(10, 0) = "種類　ID"
        csv_data(11, 0) = "種類名"
        csv_data(12, 0) = "稼動状況（0：稼動　1：非稼動）"
        csv_data(13, 0) = "口数"
        csv_data(14, 0) = "金額"
        csv_data(15, 0) = "住宅区分"
        csv_data(16, 0) = "地域番号"
        csv_data(17, 0) = "備考"
        csv_data(18, 0) = "前回請求"
        csv_data(19, 0) = "次回請求"
        csv_data(20, 0) = "世帯数"
        csv_data(21, 0) = "ＴＶ数"
        csv_data(22, 0) = "口座名義"
        csv_data(23, 0) = "銀行名"
        csv_data(24, 0) = "支店名"
        csv_data(25, 0) = "口座種類(0:普通　以外：当座）"
        csv_data(26, 0) = "口座番号"
        csv_data(27, 0) = "ノード番号"
        csv_data(28, 0) = "図面番号"
        
        csv_data(29, 0) = "工区番号"
        csv_data(30, 0) = "図面"
        csv_data(31, 0) = "工事完了日"
         csv_data(32, 0) = "備考(光)"
         csv_data(33, 0) = "リスト番号"
         csv_data(34, 0) = "郵便番号"
         csv_data(35, 0) = "クロージャー番号"
         csv_data(36, 0) = "V-ONUアドレス"
         csv_data(37, 0) = "D-ONUアドレス"
          csv_data(38, 0) = "CL開通日「"
        csv_data(39, 0) = "CL解約日"
        csv_data(40, 0) = "CL契約者名"
    rs_sentaku.MoveFirst
    sususu = 1
    Do Until rs_sentaku.EOF
            csv_data(0, sususu) = Format(sususu, "00000")
            csv_data(1, sususu) = rs_sentaku!kojinid
            s_err_id = rs_sentaku!kojinid
            If rs_sentaku!kojinid = "03111" Then
              '  MsgBox "HIT"
            End If
            csv_data(2, sususu) = rs_sentaku!Name
            csv_data(3, sususu) = rs_sentaku!Furigana
            csv_data(4, sususu) = rs_sentaku!chiikiid
            csv_data(5, sususu) = rs_sentaku!juusho1
            csv_data(6, sususu) = rs_sentaku!juusho2
            csv_data(7, sususu) = rs_sentaku!tel1
            csv_data(8, sususu) = rs_sentaku!tel2
            csv_data(9, sususu) = rs_sentaku!keitai
            csv_data(10, sususu) = rs_sentaku!shurui
            csv_data(11, sususu) = rs_sentaku!shuruimei
            csv_data(12, sususu) = rs_sentaku!kadou
            csv_data(13, sususu) = rs_sentaku!kuchisuu
            csv_data(14, sususu) = rs_sentaku!kingaku
            csv_data(15, sususu) = rs_sentaku!juutaku
            csv_data(16, sususu) = rs_sentaku!cellno
            csv_data(17, sususu) = rs_sentaku!bikou
            csv_data(18, sususu) = ""
            csv_data(19, sususu) = ""
            csv_data(20, sususu) = rs_sentaku!setaisuu
            csv_data(21, sususu) = rs_sentaku!tvsuu
            If rs_sentaku!seikyuu <> "0" Then
                csv_data(22, sususu) = ""
                csv_data(23, sususu) = ""
                csv_data(24, sususu) = ""
                csv_data(25, sususu) = ""
                csv_data(26, sususu) = ""
                 csv_data(27, sususu) = rs_sentaku!cellid
                csv_data(28, sususu) = rs_sentaku!hoshoukin
            Else
                sql_sentaku3 = "SELECT bank.bankid, bankmain.bankname, bank.bankmei," & _
                    " bank.bankcode, bank.shitencode" & _
                    " FROM bank LEFT JOIN bankmain ON bank.bankcode = bankmain.bankcode" & _
                    " WHERE bank.bankid ='" & rs_sentaku!bankid & "'"
                If FcSQlGet(rs_sentaku3, sql_sentaku3, PrMsg) = False Then
                    csv_data(23, sususu) = "Err:" & rs_sentaku!bankid
                    csv_data(24, sususu) = "Err"
                Else
                    csv_data(23, sususu) = rs_sentaku3!bankname
                    csv_data(24, sususu) = rs_sentaku3!bankmei
                    rs_sentaku3.Close
                End If
                csv_data(22, sususu) = rs_sentaku!meigi
                csv_data(25, sususu) = rs_sentaku!kouzakubun
                csv_data(26, sususu) = rs_sentaku!kouzabangou
                csv_data(27, sususu) = rs_sentaku!cellid
                csv_data(28, sususu) = rs_sentaku!hoshoukin
                
                
            End If
            csv_data(29, sususu) = rs_sentaku!h_koujibangou
            csv_data(30, sususu) = rs_sentaku!h_zumen
            csv_data(31, sususu) = rs_sentaku!h_koujikanryoubi
            csv_data(32, sususu) = rs_sentaku!h_bikou
            
             If Trim(rs_sentaku!h_listbangou) = "" Then
                 csv_data(33, sususu) = ""
            Else
                csv_data(33, sususu) = Trim(rs_sentaku!h_listbangou)
            End If
            
             If Trim(rs_sentaku!cl_bangou) = "" Then
                 csv_data(35, sususu) = ""
            Else
                csv_data(35, sususu) = Trim(rs_sentaku!cl_bangou)
            End If
            
             If Trim(rs_sentaku!v_onu_adress) = "" Then
                 csv_data(36, sususu) = ""
            Else
                csv_data(36, sususu) = Trim(rs_sentaku!v_onu_adress)
            End If
            
            
             If Trim(rs_sentaku!d_onu_adress) = "" Then
                 csv_data(37, sususu) = ""
            Else
                csv_data(37, sususu) = Trim(rs_sentaku!d_onu_adress)
            End If
            
            If IsNull(rs_sentaku!cl_open_date) Then
                 csv_data(38, sususu) = ""
            Else
                csv_data(38, sususu) = Trim(rs_sentaku!cl_open_date)
            End If
            
             If IsNull(rs_sentaku!cl_close_date) Then
                 csv_data(39, sususu) = ""
            Else
                csv_data(39, sususu) = Trim(rs_sentaku!cl_close_date)
            End If
            
              If IsNull(rs_sentaku!cl_keiyaku_name) Then
                 csv_data(40, sususu) = ""
            Else
                csv_data(40, sususu) = Trim(rs_sentaku!cl_keiyaku_name)
            End If
            
            
            If Trim(rs_sentaku!yuubin) = "" Then
                 csv_data(34, sususu) = ""
            Else
                csv_data(34, sususu) = Mid(Trim(rs_sentaku!yuubin), 1, 3) & "-" & Mid(Trim(rs_sentaku!yuubin), 4, 4)
            End If
            
        sususu = sususu + 1
        rs_sentaku.MoveNext
    Loop
    rs_sentaku.Close
On Error GoTo 0
End If

sql_sentaku = "SELECT nhk_m.kojinid, kojin.name, kojin.furigana, kojin.chiikiid,kojin.yuubin," & _
            "kojin.juusho1, kojin.juusho2, kojin.tel1, kojin.tel2, kojin.keitai," & _
            "nhk_m.shurui, nhk_m.n1, nhk_m.n2, nhk_m.jikai,nhk_m.zenkai," & _
            "nhk_m.kingaku,kojin.juutaku,kojin.cellno,kojin.bikou,kojin.setaisuu,kojin.tvsuu" & _
            " FROM kojin RIGHT JOIN nhk_m ON kojin.kojinid = nhk_m.kojinid"

If FcSQlGet(rs_sentaku2, sql_sentaku, PrMsg) = False Then
    If sususu = 1 Then
            MsgBox "作成したいデータが存在しません。"
            Exit Sub
    End If
Else
On Error GoTo eee2
    datasuu2 = rs_sentaku2.RecordCount
    
        ReDim Preserve csv_data(41, datasuu + datasuu2)
    rs_sentaku2.MoveFirst
    Do Until rs_sentaku2.EOF
            s_err_id = rs_sentaku2!kojinid
            If rs_sentaku2!kojinid = "03111" Then
                'MsgBox "HIT"
            End If
            csv_data(0, sususu) = Format(sususu, "00000")
            csv_data(1, sususu) = rs_sentaku2!kojinid
            csv_data(2, sususu) = rs_sentaku2!Name
            csv_data(3, sususu) = rs_sentaku2!Furigana
            csv_data(4, sususu) = rs_sentaku2!chiikiid
            csv_data(5, sususu) = rs_sentaku2!juusho1
            csv_data(6, sususu) = rs_sentaku2!juusho2
            csv_data(7, sususu) = rs_sentaku2!tel1
            csv_data(8, sususu) = rs_sentaku2!tel2
            csv_data(9, sususu) = rs_sentaku2!keitai
            csv_data(10, sususu) = rs_sentaku2!shurui
            Select Case rs_sentaku2!shurui
            Case "0"
                csv_data(11, sususu) = "口座振替１ヶ月"
            Case "1"
                csv_data(11, sususu) = "口座振替２ヶ月"
            Case "2"
                csv_data(11, sususu) = "口座振替３ヶ月"
            Case "3"
                csv_data(11, sususu) = "口座振替４ヶ月"
            Case "4"
                csv_data(11, sususu) = "口座振替５ヶ月"
            Case "5"
                csv_data(11, sususu) = "口座振替６ヶ月"
            Case "6"
                csv_data(11, sususu) = "口座振替６ヶ月ハンデ"
            Case "7"
                csv_data(11, sususu) = "口座振替６ヶ月家族割"
                
            End Select
            csv_data(12, sususu) = ""
            csv_data(13, sususu) = rs_sentaku2!n1
            csv_data(14, sususu) = rs_sentaku2!kingaku
            csv_data(15, sususu) = rs_sentaku2!juutaku
            csv_data(16, sususu) = rs_sentaku2!cellno
            csv_data(17, sususu) = rs_sentaku2!bikou
            csv_data(18, sususu) = rs_sentaku2!zenkai
            csv_data(19, sususu) = rs_sentaku2!jikai
            csv_data(20, sususu) = rs_sentaku2!setaisuu
            csv_data(21, sususu) = rs_sentaku2!tvsuu
                csv_data(22, sususu) = ""
                csv_data(23, sususu) = ""
                csv_data(24, sususu) = ""
                csv_data(25, sususu) = ""
                csv_data(26, sususu) = ""
                csv_data(27, sususu) = ""
                csv_data(28, sususu) = ""
                csv_data(29, sususu) = ""
                csv_data(30, sususu) = ""
                csv_data(31, sususu) = ""
                csv_data(32, sususu) = ""
                
                csv_data(33, sususu) = ""
                
                csv_data(36, sususu) = ""
                csv_data(35, sususu) = ""
                 csv_data(37, sususu) = ""
                 csv_data(38, sususu) = ""

                 csv_data(39, sususu) = ""

                 csv_data(40, sususu) = ""

                
            If Trim(rs_sentaku2!yuubin) = "" Then
                 csv_data(34, sususu) = ""
            Else
                csv_data(34, sususu) = Mid(Trim(rs_sentaku2!yuubin), 1, 3) & "-" & Mid(Trim(rs_sentaku2!yuubin), 4, 4)
            End If
            
            
        sususu = sususu + 1
        rs_sentaku2.MoveNext
    Loop
    rs_sentaku2.Close
    On Error GoTo 0
End If

On Error GoTo eee3
For writecounter = 1 To datasuu + datasuu2
        For writecounter2 = 1 To 37
          If IsNull(csv_data(writecounter2, writecounter)) Then
            csv_data(writecounter2, writecounter) = ""
          End If
        Next writecounter2
Next writecounter
On Error GoTo 0
On Error GoTo eee4
Open shutsu_path For Output Access Write As 1
For writecounter = 0 To datasuu + datasuu2
        Write #1, csv_data(0, writecounter), csv_data(1, writecounter), csv_data(2, writecounter), csv_data(3, writecounter), csv_data(4, writecounter), csv_data(15, writecounter), csv_data(16, writecounter), csv_data(34, writecounter), csv_data(5, writecounter), csv_data(6, writecounter), csv_data(7, writecounter), csv_data(8, writecounter), csv_data(9, writecounter), csv_data(10, writecounter), csv_data(11, writecounter), csv_data(12, writecounter), csv_data(13, writecounter), csv_data(14, writecounter), csv_data(15, writecounter), csv_data(17, writecounter), csv_data(18, writecounter), csv_data(19, writecounter), csv_data(20, writecounter), csv_data(21, writecounter), csv_data(22, writecounter), csv_data(23, writecounter), csv_data(24, writecounter), _
        csv_data(25, writecounter), csv_data(26, writecounter), csv_data(27, writecounter), csv_data(28, writecounter), csv_data(29, writecounter), csv_data(30, writecounter), csv_data(31, writecounter), csv_data(32, writecounter), csv_data(33, writecounter), csv_data(35, writecounter), csv_data(36, writecounter), csv_data(37, writecounter), csv_data(38, writecounter), csv_data(39, writecounter), csv_data(40, writecounter)
Next

Close #1
On Error GoTo 0

MsgBox "指定データのエクスポートが完了しました。"
Exit Sub

eee1:
MsgBox "個人データ（口座）の出力時にエラーが発生しました。"
Exit Sub
eee2:
MsgBox "個人データ（現金）の出力時にエラーが発生しました。:" & s_err_id
Exit Sub
eee3:
MsgBox "個人データの整理時にエラーが発生しました。"
Exit Sub

eee4:
MsgBox "個人データの書き込み時にエラーが発生しました。"
Exit Sub


End Sub
Sub csv_csv7(shutsu_path As String, s_hi As String)



Dim sql_sentaku As String, rs_sentaku As Recordset, sususu As Long, csv_data()
Dim datasuu As Long, writecounter As Long, kensakukoumokusuu As Integer, writecounter2 As Long
Dim sql_sentaku2 As String, rs_sentaku2 As Recordset
Dim sql_sentaku3 As String, rs_sentaku3 As Recordset, rs_fu As Recordset, rs_fu2 As Recordset
Dim sql_fu As String
Dim s_nichiji As String, s_nichiji2 As String

s_nichiji = s_hi & "01"
s_nichiji2 = s_hi & "31"


  If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If
    
'sql_sentaku = "SELECT * FROM shousai where keist='0' or (keist ='1' and keinyu >='" & s_nichiji & "')"

'sql_sentaku = "SELECT * FROM shousai WHERE (((shousai.keist)='0') AND ((shousai.keiseikyuu)<'" & s_hi & "')) OR (((shousai.keist)='1') AND ((shousai.keinyu)>='" & s_nichiji & "') AND ((shousai.keiseikyuu)<'" & s_hi & "'))"

sql_sentaku = "SELECT * FROM shousai WHERE keinyu between '" & s_nichiji & "' and '" & s_nichiji2 & "'"

If FcSQlGet2(rs_sentaku, sql_sentaku, PrMsg) = True Then
On Error GoTo eee1
    datasuu = rs_sentaku.RecordCount
        ReDim Preserve csv_data(10, datasuu)
        csv_data(0, 0) = "NO"
        csv_data(1, 0) = "個人 ID"
        csv_data(2, 0) = "個人名"
        csv_data(3, 0) = "請求年月"
        csv_data(4, 0) = "契約種類"
        csv_data(5, 0) = "契約種類名"
        csv_data(6, 0) = "契約ID"
        csv_data(7, 0) = "契約名"
        csv_data(8, 0) = "金額"
        csv_data(9, 0) = "入金日"
        
    rs_sentaku.MoveFirst
    sususu = 1
    Do Until rs_sentaku.EOF
            csv_data(0, sususu) = Format(sususu, "00000")
            csv_data(1, sususu) = rs_sentaku!kojinid
            sql_fu = "select name from kojin where kojinid ='" & rs_sentaku!kojinid & "'"
            If FcSQlGet(rs_fu, sql_fu, PrMsg) = True Then
                csv_data(2, sususu) = rs_fu!Name
                rs_fu.Close
            Else
                csv_data(2, sususu) = "error"
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
            csv_data(9, sususu) = rs_sentaku!keinyu
         
        sususu = sususu + 1
        rs_sentaku.MoveNext
    Loop
    rs_sentaku.Close
On Error GoTo 0
End If

If datasuu = 0 Then
    
    Exit Sub
End If

On Error GoTo eee3
For writecounter = 1 To datasuu
        For writecounter2 = 1 To 9
          If IsNull(csv_data(writecounter2, writecounter)) Then
            csv_data(writecounter2, writecounter) = ""
          End If
        Next writecounter2
Next writecounter
On Error GoTo 0
On Error GoTo eee4
Open shutsu_path For Output Access Write As 1
For writecounter = 0 To datasuu
        Write #1, csv_data(0, writecounter), csv_data(1, writecounter), csv_data(2, writecounter), csv_data(3, writecounter), csv_data(4, writecounter), csv_data(5, writecounter), csv_data(6, writecounter), csv_data(7, writecounter), csv_data(8, writecounter), csv_data(9, writecounter)
Next

Close #1
On Error GoTo 0

MsgBox "指定データのエクスポートが完了しました。"
Exit Sub

eee1:
MsgBox "個人データ（口座）の出力時にエラーが発生しました。"
Exit Sub

eee3:
MsgBox "個人データの整理時にエラーが発生しました。"
Exit Sub

eee4:
MsgBox "個人データの書き込み時にエラーが発生しました。"
Close #1
Exit Sub





End Sub

Sub csv_csv6(shutsu_path As String, s_hi As String)


Dim sql_sentaku As String, rs_sentaku As Recordset, sususu As Long, csv_data()
Dim datasuu As Long, writecounter As Long, kensakukoumokusuu As Integer, writecounter2 As Long
Dim sql_sentaku2 As String, rs_sentaku2 As Recordset
Dim sql_sentaku3 As String, rs_sentaku3 As Recordset, rs_fu As Recordset, rs_fu2 As Recordset
Dim sql_fu As String
Dim s_nichiji As String

s_nichiji = s_hi & "01"

  If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If
    
'sql_sentaku = "SELECT * FROM shousai where keist='0' or (keist ='1' and keinyu >='" & s_nichiji & "')"

'sql_sentaku = "SELECT * FROM shousai WHERE (((shousai.keist)='0') AND ((shousai.keiseikyuu)<'" & s_hi & "')) OR (((shousai.keist)='1') AND ((shousai.keinyu)>='" & s_nichiji & "') AND ((shousai.keiseikyuu)<'" & s_hi & "'))"

sql_sentaku = "SELECT * FROM shousai WHERE keiseikyuu='" & s_hi & "'"

If FcSQlGet2(rs_sentaku, sql_sentaku, PrMsg) = True Then
On Error GoTo eee1
    datasuu = rs_sentaku.RecordCount
        ReDim Preserve csv_data(10, datasuu)
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
        
    rs_sentaku.MoveFirst
    sususu = 1
    Do Until rs_sentaku.EOF
            csv_data(0, sususu) = Format(sususu, "00000")
            csv_data(1, sususu) = rs_sentaku!kojinid
            sql_fu = "select name from kojin where kojinid ='" & rs_sentaku!kojinid & "'"
            If FcSQlGet(rs_fu, sql_fu, PrMsg) = True Then
                csv_data(2, sususu) = rs_fu!Name
                rs_fu.Close
            Else
                csv_data(2, sususu) = "error"
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
         
        sususu = sususu + 1
        rs_sentaku.MoveNext
    Loop
    rs_sentaku.Close
On Error GoTo 0
End If

If datasuu = 0 Then
    
    Exit Sub
End If

On Error GoTo eee3
For writecounter = 1 To datasuu
        For writecounter2 = 1 To 9
          If IsNull(csv_data(writecounter2, writecounter)) Then
            csv_data(writecounter2, writecounter) = ""
          End If
        Next writecounter2
Next writecounter
On Error GoTo 0
On Error GoTo eee4
Open shutsu_path For Output Access Write As 1
For writecounter = 0 To datasuu
        Write #1, csv_data(0, writecounter), csv_data(1, writecounter), csv_data(2, writecounter), csv_data(3, writecounter), csv_data(4, writecounter), csv_data(5, writecounter), csv_data(6, writecounter), csv_data(7, writecounter), csv_data(8, writecounter), csv_data(9, writecounter)
Next

Close #1
On Error GoTo 0

MsgBox "指定データのエクスポートが完了しました。"
Exit Sub

eee1:
MsgBox "個人データ（口座）の出力時にエラーが発生しました。"
Exit Sub

eee3:
MsgBox "個人データの整理時にエラーが発生しました。"
Exit Sub

eee4:
MsgBox "個人データの書き込み時にエラーが発生しました。"
Close #1
Exit Sub




End Sub


Sub csv_csv5(shutsu_path As String, s_hi As String)

Dim sql_sentaku As String, rs_sentaku As Recordset, sususu As Long, csv_data()
Dim datasuu As Long, writecounter As Long, kensakukoumokusuu As Integer, writecounter2 As Long
Dim sql_sentaku2 As String, rs_sentaku2 As Recordset
Dim sql_sentaku3 As String, rs_sentaku3 As Recordset, rs_fu As Recordset, rs_fu2 As Recordset
Dim sql_fu As String
Dim s_nichiji As String

s_nichiji = s_hi & "01"

  If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If
    
'sql_sentaku = "SELECT * FROM shousai where keist='0' or (keist ='1' and keinyu >='" & s_nichiji & "')"

sql_sentaku = "SELECT * FROM shousai WHERE (((shousai.keist)='0') AND ((shousai.keiseikyuu)<'" & s_hi & "')) OR (((shousai.keist)='1') AND ((shousai.keinyu)>='" & s_nichiji & "') AND ((shousai.keiseikyuu)<'" & s_hi & "'))"


If FcSQlGet2(rs_sentaku, sql_sentaku, PrMsg) = True Then
On Error GoTo eee1
    datasuu = rs_sentaku.RecordCount
        ReDim Preserve csv_data(10, datasuu)
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
        
    rs_sentaku.MoveFirst
    sususu = 1
    Do Until rs_sentaku.EOF
            csv_data(0, sususu) = Format(sususu, "00000")
            csv_data(1, sususu) = rs_sentaku!kojinid
            sql_fu = "select name from kojin where kojinid ='" & rs_sentaku!kojinid & "'"
            If FcSQlGet(rs_fu, sql_fu, PrMsg) = True Then
                csv_data(2, sususu) = rs_fu!Name
                rs_fu.Close
            Else
                csv_data(2, sususu) = "error"
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
         
        sususu = sususu + 1
        rs_sentaku.MoveNext
    Loop
    rs_sentaku.Close
On Error GoTo 0
End If

If datasuu = 0 Then
    
    Exit Sub
End If

On Error GoTo eee3
For writecounter = 1 To datasuu
        For writecounter2 = 1 To 9
          If IsNull(csv_data(writecounter2, writecounter)) Then
            csv_data(writecounter2, writecounter) = ""
          End If
        Next writecounter2
Next writecounter
On Error GoTo 0
On Error GoTo eee4
Open shutsu_path For Output Access Write As 1
For writecounter = 0 To datasuu
        Write #1, csv_data(0, writecounter), csv_data(1, writecounter), csv_data(2, writecounter), csv_data(3, writecounter), csv_data(4, writecounter), csv_data(5, writecounter), csv_data(6, writecounter), csv_data(7, writecounter), csv_data(8, writecounter), csv_data(9, writecounter)
Next

Close #1
On Error GoTo 0

MsgBox "指定データのエクスポートが完了しました。"
Exit Sub

eee1:
MsgBox "個人データ（口座）の出力時にエラーが発生しました。"
Exit Sub

eee3:
MsgBox "個人データの整理時にエラーが発生しました。"
Exit Sub

eee4:
MsgBox "個人データの書き込み時にエラーが発生しました。"
Close #1
Exit Sub


End Sub

Function ppv_koushin(k_trshurui As String, k_kojin As String, k_kingaku As String, k_hiduke As String) As Integer
Dim tid As String, tkojinid As String, tkingaku As String, jurs As Recordset
Dim tkaishi As String, tkudou As String, tshurui As String, ttid As String

    
If soukatsu(0, "renban", 6, "", 1, 1) = False Then

        MsgBox "設定の読み込みに失敗しました。"
End If
If reg = "0" Then
    reg = ""
End If
If reg = vbNullString Then
    tid = "0000001"
Else
    tid = Format$(reg, "000000#")
End If

            tshurui = k_trshurui
            tkojinid = k_kojin
            tkingaku = k_kingaku
            tkaishi = k_hiduke
            tkudou = "0" '可能
        sql = "INSERT INTO keiyaku2 VALUES('" & tid & "','" & tkojinid & "'," & _
                       "'" & tshurui & "','" & tkaishi & "','" & tkudou & "','" & tkingaku & "')"
On Error GoTo errk5
        db.Execute sql, dbSQLPassThrough
On Error GoTo 0
        ttid = CStr(CLng(tid) + 1)
        If soukatsu(1, "renban", 6, ttid, 1, 1) = False Then
                MsgBox "設定の書き込みに失敗しました。"
                ppv_koushin = -1
                Exit Function
        End If
        If a_sonota_touroku(tkojinid, tkingaku, tkudou, 1) = False Then
            MsgBox "料金の更新に失敗しました。"
            ppv_koushin = -1
            Exit Function
        End If
        
    
        ppv_koushin = 1
Exit Function

errk5:
    MsgBox "契約の登録に失敗しました。"
    ppv_koushin = -1
    Exit Function

End Function
Sub stb_grid_set(retsu As Integer)
With frmkojinshousai.grdstb
        .Font.Size = 11
        .Rows = retsu + 1
        .Cols = 6
        .ColWidth(0) = 1500
        .ColWidth(1) = 2500
        .ColWidth(2) = 2500
        .ColWidth(3) = 1000
        .ColWidth(4) = 2500
        .ColWidth(5) = 1000
        .Row = 0
        .Col = 0
        .Text = "管理ＩＤ"
        .Col = 1
        .Text = "ＳＴＢＮＯ"
        .Col = 2
        .Text = "ＢＣＡＳＮＯ"
        .Col = 3
        .Text = "登録"
        .Col = 4
        .Text = "ＣＣＡＳＮＯ"
        .Col = 5
        .Text = "登録"
End With
End Sub
Sub ppv_grid_set2(retsu As Integer)
With frmkojinshousai.gridppvrireki
        .Font.Size = 11
        .Rows = retsu + 1
        .Cols = 12
        .ColWidth(0) = 1100
        .ColWidth(1) = 900
        .ColWidth(2) = 1800
        .ColWidth(3) = 2300
        .ColWidth(4) = 4000
        .ColWidth(5) = 700
        .ColWidth(6) = 1000
        .ColWidth(7) = 700
        .ColWidth(8) = 700
        .ColWidth(9) = 700
        .ColWidth(10) = 1800
        .ColWidth(11) = 1800
        .Row = 0
        .Col = 0
        .Text = "管理ＩＤ"
        .Col = 1
        .Text = "請求年月"
        .Col = 2
        .Text = "購入日"
        .Col = 3
        .Text = "商品コード"
        .Col = 4
        .Text = "商品名"
        .Col = 5
        .Text = "チャンネル"
        .Col = 6
        .Text = "料金"
        .Col = 7
        .Text = "税率"
        .Col = 8
        .Text = "録画権"
        .Col = 9
        .Text = "アダルト"
        .Col = 10
        .Text = "放送開始日"
        .Col = 11
        .Text = "収集日時"
End With
End Sub

Sub ppv_grid_set(sentakukojinid As String)
Dim sql_ppvr As String, rs_ppvr As Recordset, ppvcount As Long, i As Long

sql_ppvr = "select * from ppv where kojinid='" & sentakukojinid & "' order by ppvno DESC,kounyuubi DESC"
If kidou_ppv = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
End If
    
    If FcSQlGet_ppv(rs_ppvr, sql_ppvr, PrMsg) = False Then
        ppv_grid_set2 (0)
    Else
        rs_ppvr.MoveLast
        ppvcount = rs_ppvr.RecordCount
        ppv_grid_set2 (ppvcount)
        rs_ppvr.MoveFirst
        With frmkojinshousai.gridppvrireki
            For i = 1 To ppvcount
                    .Row = i
                    .Col = 0
                    .Text = rs_ppvr!ppvno
                    .Col = 1
                    .Text = rs_ppvr!nentsuki
                    .Col = 2
                    .Text = rs_ppvr!kounyuubi
                    .Col = 3
                    .Text = rs_ppvr!shouhincode
                    .Col = 4
                    .Text = rs_ppvr!meishou
                    .Col = 5
                    .Text = rs_ppvr!sabisuid
                    .Col = 6
                    .Text = rs_ppvr!kingaku
                    .Col = 7
                    .Text = rs_ppvr!zeiritsu
                    .Col = 8
                    If rs_ppvr!rokugaken = "0" Then
                        .Text = "×"
                    ElseIf rs_ppvr!rokugaken = "1" Then
                        .Text = "○"
                    Else
                        .Text = "err"
                    End If
                    .Col = 9
                    If rs_ppvr!adaruto = "0" Then
                        .Text = "×"
                    ElseIf rs_ppvr!adaruto = "1" Then
                        .Text = "○"
                    Else
                        .Text = "err"
                    End If
                    .Col = 10
                    .Text = rs_ppvr!housoubi
                    .Col = 11
                    .Text = rs_ppvr!shuushuubi
               rs_ppvr.MoveNext
            Next i
            rs_ppvr.Close
        End With
    End If

End Sub
Sub stb_grid_set2(sentakukojinid As String)

Dim sql_stb As String, rs_stb As Recordset, stbcount As Long, i As Long

sql_stb = "select * from stb where kojinid='" & sentakukojinid & "'"
    If FcSQlGet(rs_stb, sql_stb, PrMsg) = False Then
        stb_grid_set (0)
        frmkojinshousai.Picture10.Enabled = False
        frmkojinshousai.Picture5.Enabled = False
    Else
        rs_stb.MoveLast
        stbcount = rs_stb.RecordCount
        stb_grid_set (stbcount)
        rs_stb.MoveFirst
        frmkojinshousai.Picture10.Enabled = True
        frmkojinshousai.Picture5.Enabled = True
        With frmkojinshousai.grdstb
            For i = 1 To stbcount
                    .Row = i
                    .Col = 0
                    .Text = rs_stb!stbkanriid
                    .Col = 1
                    .Text = rs_stb!koushinbistbid
                    .Col = 2
                    .Text = rs_stb!bcasno
                    .Col = 3
                    If rs_stb!bcasumu = "0" Then
                        .Text = "×"
                    ElseIf rs_stb!bcasumu = "1" Then
                        .Text = "○"
                    Else
                        .Text = "err"
                    End If
                    .Col = 4
                    .Text = rs_stb!ccasno
                    .Col = 5
                    If rs_stb!ccasumu = "0" Then
                        .Text = "×"
                    ElseIf rs_stb!ccasumu = "1" Then
                        .Text = "○"
                    Else
                        .Text = "err"
                    End If
               rs_stb.MoveNext
            Next i
            rs_stb.Close
        End With
    End If

End Sub


Function voip_koushin(k_trshurui As String, k_kojin As String, k_kingaku As String, k_hiduke As String) As Integer
Dim tid As String, tkojinid As String, tkingaku As String, jurs As Recordset
Dim tkaishi As String, tkudou As String, tshurui As String, ttid As String

    
If soukatsu(0, "renban", 6, "", 1, 1) = False Then

        MsgBox "設定の読み込みに失敗しました。"
End If
If reg = "0" Then
    reg = ""
End If
If reg = vbNullString Then
    tid = "0000001"
Else
    tid = Format$(reg, "000000#")
End If

            tshurui = k_trshurui
            tkojinid = k_kojin
            tkingaku = k_kingaku
            tkaishi = k_hiduke
            tkudou = "0" '可能
        sql = "INSERT INTO keiyaku2 VALUES('" & tid & "','" & tkojinid & "'," & _
                       "'" & tshurui & "','" & tkaishi & "','" & tkudou & "','" & tkingaku & "')"
On Error GoTo errk5
        db.Execute sql, dbSQLPassThrough
On Error GoTo 0
        ttid = CStr(CLng(tid) + 1)
        If soukatsu(1, "renban", 6, ttid, 1, 1) = False Then
                MsgBox "設定の書き込みに失敗しました。"
                voip_koushin = -1
                Exit Function
        End If
        If a_sonota_touroku(tkojinid, tkingaku, tkudou, 1) = False Then
            MsgBox "料金の更新に失敗しました。"
            voip_koushin = -1
            Exit Function
        End If
        
    
        voip_koushin = 1
Exit Function

errk5:
    MsgBox "契約の登録に失敗しました。"
    voip_koushin = -1
    Exit Function

End Function
Sub nhk_shuusei_sub()
                Dim sql_nhksakusei As String, rs_nhksakusei As Recordset, shiteinhkbi As String
                Dim conutid As String, newcountid As String, shiteinhkbi2 As Integer, paspas

                If kidou4 = False Then
                            MsgBox "ＮＨＫデータＯＰＥＮに失敗しました。"
                            Exit Sub
                End If
                shiteinhkbi2 = CInt(Format(Date, "mm"))
                Select Case shiteinhkbi2
                Case 1, 3, 5, 7, 9, 11
                    paspas = InputBox("パスワード", vbYesNo)
                    If paspas = "kamifusafusa" Then
                        shiteinhkbi = InputBox("日付を8桁で。例「20030900」")
                    Else
                        MsgBox "今月はＮＨＫの集計データの更新はできません。"
                        db4.Close
                        Exit Sub
                    End If
                Case Else
                    shiteinhkbi2 = shiteinhkbi2
                    shiteinhkbi = Format(Date, "yyyy") & Format(shiteinhkbi2, "00") & "00" 'InputBox("日付を8桁で。例「20030900」")
                End Select
                                If soukatsu(0, "renban", 10, "", 1, 0) = False Then
                                    MsgBox "設定の読み込みに失敗しました。"
                                    db4.Close
                                    Exit Sub
                                End If
                                If reg = "0" Then
                                    conutid = "000001"
                                Else
                                    conutid = Format(reg, "000000")
                                End If
                
                    sql_nhksakusei = "select * from nhk_k where seikyuubi ='" & shiteinhkbi & "'"
                    If FcSQlGet4(rs_nhksakusei, sql_nhksakusei, PrMsg) = 0 Then
                        sql_nhksakusei = "select * from nhk_m where jikai ='" & shiteinhkbi & "'"
                        If FcSQlGet4(rs_nhksakusei, sql_nhksakusei, PrMsg) > 0 Then
                            rs_nhksakusei.MoveFirst
                            Do Until rs_nhksakusei.EOF
                            '集計用データの作成
                                sql_nhksakusei = "INSERT INTO nhk_k VALUES('" & conutid & "'," & _
                                                   "'" & rs_nhksakusei!kojinid & "','" & rs_nhksakusei!jikai & "'," & _
                                                   "'" & rs_nhksakusei!kingaku & "','" & rs_nhksakusei!shurui & "'," & _
                                                   "'0','" & Space(1) & "'," & _
                                                   "'" & Space(1) & "')"
                                On Error GoTo errnhksakusei10
                                    db4.Execute sql_nhksakusei, dbSQLPassThrough
                                On Error GoTo 0
                                
                                newcountid = CStr(CLng(conutid) + 1)
                                conutid = Format(newcountid, "000000")
                                rs_nhksakusei.MoveNext
                            Loop
                                If soukatsu(1, "renban", 10, newcountid, 0, 1) = False Then
                                        MsgBox "設定の書き込みに失敗しました。"
                                        db4.Close
                                        Exit Sub
                                End If
                            MsgBox "更新しました。"
                        Else
                            MsgBox "次回引き落とし契約がありませんでした。"
                        End If
                    Else
                        MsgBox "すでに契約更新データが作成されています。"
                    End If
                    db4.Close
Exit Sub

errnhksakusei10:
    MsgBox "データの修正に失敗しました。"
    Exit Sub
                

End Sub
Sub systemend()

Dim res_systemend, jikkoufile As String

If soukatsu(0, "startup", 11, "", 1, 0) = False Then
        MsgBox "設定の読み込みに失敗しました。"
        Unload frmmain
        Set frmmain = Nothing
        End
End If
If reg = "0" Then
    reg = ""
End If
jikkoufile = reg

If jikkoufile <> Format(Date, "yyyymmdd") Then

    jikkoufile = App.Path & "\ezmanager_ment.exe"
    If Dir(jikkoufile) = "" Then
        MsgBox "実行ファイルがありません。"
        End
    End If
    If jikkoufile <> "" Then
        jikkoufile = jikkoufile & " -u"
        If soukatsu(1, "startup", 11, Format(Date, "yyyymmdd"), 0, 1) = False Then
                MsgBox "自動バックアップの日付の更新に失敗しました。"
        End If
          Set db2 = Nothing
        Set db = Nothing
        Set db4 = Nothing
        Set db3 = Nothing
        Set db_voip = Nothing
        Set db_out = Nothing
        Set db_convart = Nothing
        Set db_ppv = Nothing
        On Error GoTo errsystemend
            res_systemend = Shell(jikkoufile, 0)
        On Error GoTo 0
    End If
End If

End

Exit Sub

errsystemend:
    MsgBox "指定されたプログラムを実行できませんでした。"
    End

End Sub
Sub err_write(err_shurui As String, err_naiyou As String, err_katego As String, err_old As String, err_new As String)

Dim err_file As String, fno As Integer, jitsu_tsuika_text As String

err_file = sopa & "error.txt"
fno = FreeFile
jitsu_tsuika_text = Format(Date, "yyyymmdd") & "," & Format(Time, "hhmmss") & "," & err_shurui & "," & err_naiyou & "," & err_katego & "," & err_old & "," & err_new

Open err_file For Append As #fno
Print #fno, jitsu_tsuika_text
Close #fno

End Sub

Sub err_write2(err_naiyou As String, err_file As String)

Dim fno As Integer, jitsu_tsuika_text As String

fno = 1
jitsu_tsuika_text = Format(Date, "yyyymmdd") & "," & Format(Time, "hhmmss") & "," & err_naiyou

Open err_file For Append As #fno
Print #fno, jitsu_tsuika_text
Close #fno

End Sub

Sub kyousei_shuusei(shitei_kojinid As String, newdata)

Dim sel_qq As String, newjikaibilen As Long, newdata2 As String
Dim shitsuki As String
            If newdata <> "" Then
                If newdata = "0" Then
                    newdata2 = "0"
                Else
                    newjikaibilen = Len(newdata)
                    If newjikaibilen <> 7 Then
                        MsgBox "入力したデータが７桁ではなく不正です。"
                        Exit Sub
                    End If
                    
                    shitsuki = Right(newdata, 2)
                    If shitsuki = "04" Or shitsuki = "10" Then
                        newdata2 = Format(newdata, "yyyymm")
                    Else
                        MsgBox "次回請求日の月は、「04」、「10」のみです。"
                        Exit Sub
                    End If
                End If
            Else
                MsgBox "次回請求日が入力されていません。"
                Exit Sub
            End If


        sel_qq = "update kojin set kajou='" & newdata2 & "'" & _
                " where kojinid='" & shitei_kojinid & "'"
        
        On Error GoTo errmaebi
         db.Execute sel_qq, dbSQLPassThrough
        On Error GoTo 0
        MsgBox "変更しました。"
Exit Sub
errmaebi:
MsgBox "データの更新に失敗しました。"
Exit Sub
End Sub
Sub data_check2(Optional no As Integer)


Dim sql_comb As String, rs_comb As Recordset, sql_new As String
Dim nowdata(5) As Long, ncncnc As Long, newsoukeikei As Long, newsoukeikei2() As Long, newsoukeikei2suu As Long
Dim resresres, shitebankid As String, nowf1 As Integer, newsoukeikei3 As Long, newsoukeikei3suu As Long
Dim sql_comb2 As String, rs_comb2 As Recordset, tttatai As String, tttatai2 As Long, tototo As Integer
Dim sql_cnhk As String, rs_cnhk As Recordset, mmmsg As String, newsoukeikei_sub As Long, newsoukeikei4 As Long, newsoukeikei5 As Long
Dim sql_cnh2 As String, rs_cnhk2 As Recordset
Dim sql_comb3 As String, rs_comb3 As Recordset, allnyuukin As Long

resresres = MsgBox("個人データのチェック・修正を開始します。", vbYesNo)
If resresres = vbYes Then

    FRMOMACHI.Show
    FRMOMACHI.pb.Visible = True
    FRMOMACHI.lblmsg.Caption = "個人データのチェック・修正中・・"
    DoEvents
    ncncnc = 0
    sql_comb = "select * from kojin order by kojinid"
    If FcSQlGet(rs_comb, sql_comb, PrMsg) = True Then
        FRMOMACHI.pb.Max = rs_comb.RecordCount
        rs_comb.MoveFirst
       Do Until rs_comb.EOF
            If rs_comb!kojinid = "01402" Then
              '  MsgBox "hit"
            End If
            If Trim(rs_comb!minyuukin) = "" Or Trim(rs_comb!minyuukin) = "0" Then
                nowdata(0) = 0
            Else
                nowdata(0) = CLng(rs_comb!minyuukin)
            End If
            If Trim(rs_comb!f4) = "" Or Trim(rs_comb!f4) = "0" Then
                nowdata(1) = 0
            Else
                nowdata(1) = CLng(rs_comb!f4)
            End If
            If Trim(rs_comb!jikaikin) = "" Or Trim(rs_comb!jikaikin) = "0" Then
                nowdata(2) = 0
            Else
                nowdata(2) = CLng(rs_comb!jikaikin)
            End If
            If Trim(rs_comb!sonotakin) = "" Or Trim(rs_comb!sonotakin) = "0" Then
                nowdata(3) = 0
            Else
                nowdata(3) = CLng(rs_comb!sonotakin)
            End If
            If Trim(rs_comb!soukei) = "" Or Trim(rs_comb!soukei) = "0" Then
                nowdata(4) = 0
            Else
                nowdata(4) = CLng(rs_comb!soukei)
            End If
            
            newsoukeikei = nowdata(0) + nowdata(1) + nowdata(2) + nowdata(3)
            
            If newsoukeikei <> nowdata(4) Then
                'MsgBox "data changed! kojinid:" & rs_comb!kojinid
                err_write "ﾃﾞｰﾀﾁｪｯｸ", "kojinid:" & rs_comb!kojinid & " の請求総額を修正", "soukei", CStr(nowdata(4)), CStr(newsoukeikei)
                sql_new = "update kojin set soukei = '" & CStr(newsoukeikei) & "'" & _
                                    " where kojinid ='" & rs_comb!kojinid & "'"
                On Error GoTo verup_new2
                    db.Execute sql_new, dbSQLPassThrough
                On Error GoTo 0
            End If
            '2007/10/17　次回請求の日時のずれのチェック
            If Trim(rs_comb!kajou) <> "0" And Trim(rs_comb!kajou) <> "1" Then
                tttatai2 = Len(rs_comb!kajou)
                If tttatai2 = 6 Then
                    tttatai = Mid(rs_comb!kajou, 5, 2)
                    If tttatai <> "4" And tttatai <> "04" And tttatai <> "10" Then
                        mmmsg = "次回請求日の値がずれています。修正を開始しますか？個人：" & rs_comb!kojinid & "-" & rs_comb!Name & ",値＝" & Trim(rs_comb!kajou)
                        resresres = MsgBox(mmmsg, vbYesNo)
                        If resresres = vbYes Then
                            res = InputBox("正確な次回請求日を入力してください。例）2009/10" & Chr(13) & "また請求日が毎月の場合は、「0」を入力してください。", "新しい次回請求日")
                            If res = "" Then
                                    MsgBox "次回請求日の値が不正です。個人ID：" & rs_comb!kojinid & "-" & rs_comb!Name & ",値＝" & Trim(rs_comb!kajou)
                            Else
                                kyousei_shuusei rs_comb!kojinid, res
                            End If
                        End If
                    End If
                Else
                    MsgBox "次回請求日の値が不正です。個人ID：" & rs_comb!kojinid & ",値＝" & Trim(rs_comb!kajou)
                End If
            End If
            rs_comb.MoveNext
            ncncnc = ncncnc + 1
            FRMOMACHI.pb.Value = ncncnc
            DoEvents
        Loop
        rs_comb.Close
    Else
        MsgBox "error"
    End If
    
    
    If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If

    FRMOMACHI.Show
    FRMOMACHI.lblmsg.Caption = "個人データの契約の内容（未入金と請求済みデータ）のチェック中・・"

    DoEvents
    
    ncncnc = 0
    sql_comb = "select * from kojin order by kojinid"
    If FcSQlGet(rs_comb, sql_comb, PrMsg) = True Then
    FRMOMACHI.pb.Max = rs_comb.RecordCount
        rs_comb.MoveFirst
       Do Until rs_comb.EOF
            If rs_comb!kojinid = "01402" Then
              '  MsgBox "hit"
            End If

             newsoukeikei = 0
            If Trim(rs_comb!minyuukin) = "" Or Trim(rs_comb!minyuukin) = "0" Then
                nowdata(0) = 0
            Else
                nowdata(0) = CLng(rs_comb!minyuukin)
            End If
            If Trim(rs_comb!soukei) = "" Or Trim(rs_comb!soukei) = "0" Then
                nowdata(4) = 0
            Else
                nowdata(4) = CLng(rs_comb!soukei)
            End If
            If Trim(rs_comb!f1) = "" Or Trim(rs_comb!f1) = "0" Or IsNull(rs_comb!f1) Then
                nowf1 = 0
            Else
                nowf1 = CInt(rs_comb!f1)
            End If
            sql_comb2 = "select * from shousai where kojinid='" & rs_comb!kojinid & "' and keist='0'"
            If FcSQlGet2(rs_comb2, sql_comb2, PrMsg) = True Then
                rs_comb2.MoveFirst
                Do Until rs_comb2.EOF
                    newsoukeikei = newsoukeikei + CLng(rs_comb2!keikin)
                    shitebankid = Trim(rs_comb2!bankid)
                    rs_comb2.MoveNext
                Loop
                rs_comb2.Close
            End If
            If nowf1 = 1 Then
                If newsoukeikei <> nowdata(4) Then
                    'shousaiを削除
                        
                        MsgBox "請求金と詳細合計が一致しません。個人ID：" & rs_comb!kojinid & "-" & newsoukeikei & "/" & nowdata(4)
                End If
            Else
                    If newsoukeikei <> nowdata(0) Then
                        'shousaiを削除
                            
                            MsgBox "未入金と詳細合計が一致しません。個人ID：" & rs_comb!kojinid & "-" & newsoukeikei & "/" & nowdata(0)
                    End If
            End If
            rs_comb.MoveNext
            ncncnc = ncncnc + 1
            FRMOMACHI.pb.Value = ncncnc
            DoEvents
        Loop
        rs_comb.Close
    Else
        MsgBox "error"
    End If

    If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If

    FRMOMACHI.Show
    FRMOMACHI.lblmsg.Caption = "個人データの契約の内容（入金と入金済みデータ）のチェック中・・"
    DoEvents
    
    ncncnc = 0
    sql_comb = "select * from koushinkiroku where houhou ='0' or houhou='1' order by kojinid,koushinbi"
    If FcSQlGet2(rs_comb, sql_comb, PrMsg) = True Then
    FRMOMACHI.pb.Max = rs_comb.RecordCount
        rs_comb.MoveFirst
       Do Until rs_comb.EOF
             newsoukeikei = 0
             newsoukeikei_sub = 0
             newsoukeikei3 = 0
             newsoukeikei4 = 0
             newsoukeikei5 = 0
    '入金額*********＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
            If Trim(rs_comb!kingaku) = "" Or Trim(rs_comb!kingaku) = "0" Then
                nowdata(0) = 0
            Else
                nowdata(0) = CLng(rs_comb!kingaku)
            End If
            If rs_comb!kojinid = "01279" Then
                If rs_comb!koushinbi = "20120517" Then
                'MsgBox "hit"
                End If
            End If
    '詳細テーブルの抽出（同じ日の扱い分を）
            sql_comb2 = "select * from shousai where kojinid='" & rs_comb!kojinid & "'" & _
                        " and keinyu='" & rs_comb!koushinbi & "'"
            If FcSQlGet2(rs_comb2, sql_comb2, PrMsg) = True Then
                newsoukeikei3suu = rs_comb2.RecordCount
                ReDim newsoukeikei2(newsoukeikei3suu) As Long
                newsoukeikei2suu = 1
                rs_comb2.MoveFirst
                Do Until rs_comb2.EOF
                    newsoukeikei5 = newsoukeikei5 + CLng(rs_comb2!keikin)
                    If rs_comb2!keist = "1" Then
                        newsoukeikei = newsoukeikei + CLng(rs_comb2!keikin)
                        newsoukeikei_sub = newsoukeikei_sub + CLng(rs_comb2!keikin)
                        shitebankid = Trim(rs_comb2!bankid)
                    ElseIf rs_comb2!keist = "0" Then
                        newsoukeikei3 = newsoukeikei3 + CLng(rs_comb2!keikin)
                        newsoukeikei2(newsoukeikei2suu) = CLng(rs_comb2!keikin)
                        newsoukeikei2suu = newsoukeikei2suu + 1
                    Else
                        newsoukeikei = nowdata(0)
                        newsoukeikei4 = newsoukeikei4 + CLng(rs_comb2!keikin)
                    End If
                    rs_comb2.MoveNext
                Loop
                rs_comb2.Close
            End If
            If newsoukeikei <> nowdata(0) Then
                
                If newsoukeikei5 - newsoukeikei4 <> nowdata(0) Then
                    If newsoukeikei3 <> nowdata(0) Then
                        tototo = 0
                        For newsoukeikei2suu = 1 To newsoukeikei3suu
                            If newsoukeikei2(newsoukeikei2suu) = nowdata(0) Then
                                tototo = 1
                                Exit For
                            End If
                        Next
                        If tototo = 0 Then
                            'shousaiを削除
                            rs_comb.MoveNext
                            If Trim(rs_comb!kingaku) = "" Or Trim(rs_comb!kingaku) = "0" Then
                                nowdata(0) = nowdata(0) + 0
                            Else
                                nowdata(0) = nowdata(0) + CLng(rs_comb!kingaku)
                            End If
                            If newsoukeikei <> nowdata(0) Then
                                If newsoukeikei_sub <> nowdata(0) Then
                                    rs_comb.MovePrevious
                                    MsgBox "入金と詳細合計が一致しません。個人ID：" & rs_comb!kojinid
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            rs_comb.MoveNext
            ncncnc = ncncnc + 1
            FRMOMACHI.pb.Value = ncncnc
            DoEvents
        Loop
        rs_comb.Close
    Else
        MsgBox "error"
    End If

Dim mo_shu_shou As Long, mo_shu_sei As Long, mo_minyu As Long, mo_soukei As Long, rrrres, newsoukei As Long
Dim mo_nhk As Long, motokajou As String


    If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If

    FRMOMACHI.Show
    FRMOMACHI.lblmsg.Caption = "個人データの契約の内容（CATV基本契約とその他のデータ）のチェック中・・"
    DoEvents
    
    ncncnc = 0
    '契約を持った人の抽出
   ' If no <> 1 Then
        sql_comb = "SELECT kojin.f4,kojin.kojinid,kojin.name, kojin.minyuukin, kojin.jikaishu, kojin.jikaikin, kojin.soukei, kojin.kajou" & _
                " FROM keiyaku RIGHT JOIN kojin ON keiyaku.kojinid = kojin.kojinid" & _
                " WHERE ((keiyaku.kadou)='0');"
    
    If FcSQlGet(rs_comb, sql_comb, PrMsg) = True Then
    FRMOMACHI.pb.Max = rs_comb.RecordCount
        rs_comb.MoveFirst
       Do Until rs_comb.EOF
            If rs_comb!kojinid = "02873" Then
               ' MsgBox "hit"
            End If
            newsoukeikei = 0
             mo_shu_shou = CLng(rs_comb!jikaishu)
             mo_shu_sei = CLng(rs_comb!jikaikin)
             mo_minyu = CLng(rs_comb!minyuukin)
             mo_nhk = CLng(rs_comb!f4)
             mo_soukei = CLng(rs_comb!soukei)
             motokajou = Trim(rs_comb!kajou)
            sql_comb2 = "SELECT keiyaku.*" & _
                                " From keiyaku" & _
                                " WHERE (((keiyaku.kojinid)='" & rs_comb!kojinid & "') AND  ((keiyaku.kadou)='0'));"
            If FcSQlGet(rs_comb2, sql_comb2, PrMsg) = True Then
                rs_comb2.MoveFirst
                Do Until rs_comb2.EOF
                    newsoukeikei = newsoukeikei + CLng(rs_comb2!kingaku)
                    rs_comb2.MoveNext
                Loop
                rs_comb2.Close
                allnyuukin = 0
                If newsoukeikei <> mo_shu_sei Then
                    If mo_shu_sei = 0 Then
                            sql_comb3 = "select * from shousai where kojinid='" & rs_comb!kojinid & "'" & _
                                    " and keiseikyuu='" & Format(Date, "yyyymm") & "' and keist='1' and keiid<>'00'"
                            
                            If FcSQlGet2(rs_comb3, sql_comb3, PrMsg) = True Then
                                rs_comb3.MoveFirst
                                Do Until rs_comb3.EOF
                                    allnyuukin = allnyuukin + CLng(rs_comb3!keikin)
                                    rs_comb3.MoveNext
                                Loop
                                rs_comb3.Close
                            End If
                            If newsoukeikei <> allnyuukin Then
                                MsgBox "主契約の金額と請求金額合計が一致しません。確認してください。個人ID：" & rs_comb!kojinid
                            End If
                    Else
                    'shousaiを削除
                     MsgBox "主契約の金額と請求金額合計が一致しません。確認してください。個人ID：" & rs_comb!kojinid
                   End If
                   ' rrrres = MsgBox("主契約の金額と請求金額合計が一致しません。確認してください。修正してよい場合は「Yes」を！！個人ID：" & rs_comb!kojinid, vbYesNo)
                   ' If rrrres = vbYes Then
                   '     newsoukei = newsoukeikei
                    '    If mo_minyu <> 0 Then
                    '        newsoukei = newsoukei + mo_minyu
                   '     End If
                   '     If mo_nhk <> 0 Then
                   '         newsoukei = newsoukei + mo_nhk
                   '     End If
                   '     sql_new = "update kojin set jikaikin = '" & CStr(newsoukeikei) & "',soukei = '" & CStr(newsoukei) & "'" & _
                                            " where kojinid ='" & rs_comb!kojinid & "'"
                   '     On Error GoTo verup_new2
                   '         db.Execute sql_new, dbSQLPassThrough
                   '    On Error GoTo 0
                    'End If
                End If
            Else
                MsgBox "契約データがありません。顧客情報を確認してください。No：" & rs_comb!kojinid
              '  If motokajou = "0" Then
                      '  sql_new = "update kojin set kajou = '1'" & _
                                            " where kojinid ='" & rs_comb!kojinid & "'"
                      '  On Error GoTo verup_new2
                      '      db.Execute sql_new, dbSQLPassThrough
                      '  On Error GoTo 0
                      '  err_write "ﾃﾞｰﾀﾁｪｯｸ", "kojinid:" & rs_comb!kojinid & " の次回請求日を修正", "kajou", motokajou, "1"
              '  End If
            End If
            rs_comb.MoveNext
            ncncnc = ncncnc + 1
            FRMOMACHI.pb.Value = ncncnc
            DoEvents
        Loop
        rs_comb.Close
    Else
        MsgBox "error"
    End If
    
 '   End If

    
    If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If

    FRMOMACHI.Show
    FRMOMACHI.lblmsg.Caption = "個人データの契約の内容２（CATV基本契約とその他のデータ）のチェック中・・"
    DoEvents
    
    ncncnc = 0
    '主契約を持たない人の抽出
    sql_comb = "SELECT kojin.kojinid,kojin.kajou" & _
                " FROM keiyaku RIGHT JOIN kojin ON keiyaku.kojinid = kojin.kojinid" & _
                " WHERE (((keiyaku.kadou)='0') AND ((keiyaku.shurui)<>'06'))" & _
                " group by kojin.kojinid,kojin.kajou order by kojin.kojinid"
    If FcSQlGet(rs_comb, sql_comb, PrMsg) = True Then
    FRMOMACHI.pb.Max = rs_comb.RecordCount
        rs_comb.MoveFirst
       Do Until rs_comb.EOF
            If rs_comb!kojinid = "01402" Then
             '   MsgBox "hit"
            End If
       
            If rs_comb!kajou <> "0" Then
                sql_comb2 = "SELECT * From keiyaku" & _
                            " WHERE keiyaku.kojinid='" & rs_comb!kojinid & "' AND keiyaku.shurui='01' AND keiyaku.kadou='0'"
                If FcSQlGet(rs_comb2, sql_comb2, PrMsg) = False Then
                    'rrrres = MsgBox("CATV主契約が存在しないのに更新日の設定があります。修正しますか？個人ID：" & rs_comb!kojinid, vbYesNo)
                    'If rrrres = vbYes Then
                        sql_new = "update kojin set kajou = '0'" & _
                                            " where kojinid ='" & rs_comb!kojinid & "'"
                        On Error GoTo verup_new2
                            db.Execute sql_new, dbSQLPassThrough
                        On Error GoTo 0
                        err_write "ﾃﾞｰﾀﾁｪｯｸ", "kojinid:" & rs_comb!kojinid & " の次回請求日を修正", "kajou", CStr(rs_comb!kajou), "0"
                    'End If
                End If
            End If
            rs_comb.MoveNext
            ncncnc = ncncnc + 1
            FRMOMACHI.pb.Value = ncncnc
            DoEvents
        Loop
        rs_comb.Close
    Else
        MsgBox "error"
    End If
    
     If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If
    
    
      FRMOMACHI.Show
    FRMOMACHI.lblmsg.Caption = "個人データの契約の内容３（CL基本契約とインターネット割引）のチェック中・・"
    DoEvents
    
    ncncnc = 0
    '主契約にCL基本を持っている人の抽出
    Dim s_st As Integer
    
    sql_comb = "SELECT kojin.kojinid,kojin.kajou" & _
                " FROM keiyaku RIGHT JOIN kojin ON keiyaku.kojinid = kojin.kojinid" & _
                " WHERE keiyaku.kadou='0' and keiyaku.shurui='33'"
                
    If FcSQlGet(rs_comb, sql_comb, PrMsg) = True Then
    FRMOMACHI.pb.Max = rs_comb.RecordCount
        rs_comb.MoveFirst
       Do Until rs_comb.EOF
            If rs_comb!kojinid = "01402" Then
             '   MsgBox "hit"
            End If
       
           
                sql_comb2 = "SELECT * From keiyaku" & _
                            " WHERE keiyaku.kojinid='" & rs_comb!kojinid & "' AND (keiyaku.shurui='27' or keiyaku.shurui='28' or keiyaku.shurui='29' or keiyaku.shurui='34') AND keiyaku.kadou='0'" & _
                            " order by keiyaku.shurui"
                If FcSQlGet(rs_comb2, sql_comb2, PrMsg) = True Then
                        s_st = 0
                        rs_comb2.MoveFirst
                        Do Until rs_comb2.EOF
                            
                            
                            rs_comb2.MoveNext
                        Loop
                        
                End If
           
           
            rs_comb.MoveNext
            ncncnc = ncncnc + 1
            FRMOMACHI.pb.Value = ncncnc
            DoEvents
        Loop
        rs_comb.Close
    Else
        MsgBox "error"
    End If
    

    FRMOMACHI.Show
    FRMOMACHI.lblmsg.Caption = "個人データの契約の内容（詳細）のチェック中・・"
    DoEvents
   
   Dim shshdata As String, shshdata2 As String, stsuki As Integer, snen As Long

     If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If
    ncncnc = 0
    snen = CLng(Format(Date, "yyyy")) - 1
    stsuki = CInt(Format(Date, "mm")) - 1
    shshdata = Format(Date, "yyyy") & Format(Date, "mm")
    If Format(Date, "mm") = "01" Then
        shshdata2 = CStr(snen) & "12"
    Else
        shshdata2 = Format(Date, "yyyy") & Format(stsuki, "00")
    End If
    sql_comb2 = "select * from shousai where keiseikyuu='" & shshdata & "' or  keiseikyuu='" & shshdata2 & "'"
    If FcSQlGet2(rs_comb2, sql_comb2, PrMsg) = True Then
        FRMOMACHI.pb.Max = rs_comb2.RecordCount
        rs_comb2.MoveFirst
        Do Until rs_comb2.EOF
            sql_comb = "select * from shousai where keiseikyuu='" & rs_comb2!keiseikyuu & "' and keiid='" & rs_comb2!keiid & "' and keikin='" & rs_comb2!keikin & "' and kojinid='" & rs_comb2!kojinid & "' and rs_comb2!keishurui='1'"
            If FcSQlGet2(rs_comb, sql_comb, PrMsg) = True Then
                If rs_comb.RecordCount > 1 Then
                    MsgBox rs_comb2!kojinid & " 詳細データが重複しています。"
                End If
            End If
            ncncnc = ncncnc + 1
            FRMOMACHI.pb.Value = ncncnc
            DoEvents
            rs_comb2.MoveNext
        Loop
        rs_comb2.Close
    End If

    
    If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If

    FRMOMACHI.Show
    FRMOMACHI.lblmsg.Caption = "個人データの契約の内容（契約と合計）のチェック中・・"
    DoEvents
    
    ncncnc = 0
    '人の抽出
    sql_comb = "SELECT kojin.*" & _
                " FROM  kojin order by kojinid"
    If FcSQlGet(rs_comb, sql_comb, PrMsg) = True Then
    FRMOMACHI.pb.Max = rs_comb.RecordCount
        rs_comb.MoveFirst
       Do Until rs_comb.EOF
       
            If rs_comb!kojinid = "01402" Then
               ' MsgBox "hit"
            End If
       
            newsoukeikei = 0
             If CLng(rs_comb!jikaishu) <> 0 Then
                    sql_comb2 = "SELECT sum(kingaku) as kinkinkin" & _
                                " From keiyaku" & _
                                " WHERE keiyaku.kojinid='" & rs_comb!kojinid & "'"
                    If FcSQlGet(rs_comb2, sql_comb2, PrMsg) = True Then
                        If IsNull(rs_comb2!kinkinkin) Then
                            MsgBox "主契約小計額と主契約伝票の合計が一致しません。個人ID：" & rs_comb!kojinid
                        Else
                            If CLng(rs_comb!jikaishu) <> CLng(rs_comb2!kinkinkin) Then
                                MsgBox "主契約小計額と主契約伝票の合計が一致しません。個人ID：" & rs_comb!kojinid
                            End If
                        End If
                        rs_comb2.Close
                    End If
             End If
             If CLng(rs_comb!jikaikin) <> 0 Then
                    sql_comb2 = "SELECT sum(kingaku) as kinkinkin3" & _
                                " From keiyaku" & _
                                " WHERE keiyaku.kojinid='" & rs_comb!kojinid & "' and keiyaku.kadou='0' and keiyaku.shime='0'"
                    If FcSQlGet(rs_comb2, sql_comb2, PrMsg) = True Then
                        If IsNull(rs_comb2!kinkinkin3) Then
                            MsgBox "主契約小計額と主契約伝票の合計が一致しません。個人ID：" & rs_comb!kojinid
                        Else
                            If CLng(rs_comb!jikaikin) <> CLng(rs_comb2!kinkinkin3) Then
                                If rs_comb!kajou = "0" Then
                                    MsgBox "主契約小計額と主契約伝票の合計が一致しません。個人ID：" & rs_comb!kojinid
                                Else
                                    If CLng(rs_comb!jikaikin) <> CLng(rs_comb2!kinkinkin3) - 6000 Then
                                        MsgBox "主契約小計額と主契約伝票の合計が一致しません。個人ID：" & rs_comb!kojinid
                                    End If
                                End If
                            End If
                        End If
                        rs_comb2.Close
                    End If
                    sql_comb2 = "SELECT *" & _
                                " From keiyaku" & _
                                " WHERE keiyaku.kojinid='" & rs_comb!kojinid & "' and keiyaku.kadou='0' and keiyaku.shime='0'"
                    If FcSQlGet(rs_comb2, sql_comb2, PrMsg) = True Then
                        If rs_comb2!sadou = "1" Then
                            sql_new = "update keiyaku set sadou = '0'" & _
                                                " where kojinid ='" & rs_comb!kojinid & "'"
                            On Error GoTo verup_new2
                                db.Execute sql_new, dbSQLPassThrough
                            On Error GoTo 0
                            err_write "ﾃﾞｰﾀﾁｪｯｸ", "kojinid:" & rs_comb!kojinid & " の契約のSADOUを修正", "kajou", "1", "0"
                        End If
                       rs_comb2.Close
                    End If
             End If
             
             If CLng(rs_comb!jikaita) <> 0 Then
                    sql_comb2 = "SELECT sum(kingaku) as kinkinkin2" & _
                                " From keiyaku2" & _
                                " WHERE keiyaku2.kojinid='" & rs_comb!kojinid & "' AND keiyaku2.f2='0'"
                    If FcSQlGet(rs_comb2, sql_comb2, PrMsg) = True Then
                        If IsNull(rs_comb2!kinkinkin2) Then
                            MsgBox "その他契約小計額とその他契約伝票の合計が一致しません。個人ID：" & rs_comb!kojinid
                        Else
                            If CLng(rs_comb!jikaita) <> CLng(rs_comb2!kinkinkin2) Then
                                MsgBox "その他契約小計額とその他契約伝票の合計が一致しません。個人ID：" & rs_comb!kojinid
                            End If
                        End If
                        rs_comb2.Close
                    End If
             End If
             
             
            rs_comb.MoveNext
            ncncnc = ncncnc + 1
            FRMOMACHI.pb.Value = ncncnc
            DoEvents
        Loop
        rs_comb.Close
        
        
    'NHKチェック
    
        FRMOMACHI.Show
    FRMOMACHI.lblmsg.Caption = "個人データのNHKの請求のチェック中・・"
    DoEvents
    
    ncncnc = 0
        
    Dim s_nentsuki As String, sql_shousai As String, rs_shousai As Recordset
    
    s_nentsuki = Format(Date, "yyyymm")
    
    If Mid(s_nentsuki, 5, 2) = "04" Or Mid(s_nentsuki, 5, 2) = "10" Then
        If kidou4 = False Then
               MsgBox "データベースファイルを開けませんでした。"
               Exit Sub
           End If
           
        
           
           sql_shousai = "SELECT * FROM nhk_m" & _
                       " WHERE  jikai='" & s_nentsuki & "00" & "'"
          
           If FcSQlGet4(rs_shousai, sql_shousai, PrMsg) > 0 Then
                FRMOMACHI.pb.Max = rs_shousai.RecordCount
              rs_shousai.MoveFirst
                Do Until rs_shousai.EOF
                    'If Trim(rs_shousai!n2) = 0 Then
                      ' MsgBox "NHKの契約があるが、請求されていない可能性があります。チェックしてください。個人ID：" & rs_shousai!kojinid
                   'Else
                        sql_comb = "select * from kojin where kojinid='" & rs_shousai!kojinid & "' and f4='" & rs_shousai!kingaku & "'"
                        If FcSQlGet(rs_comb, sql_comb, PrMsg) = True Then
                            If rs_comb.RecordCount = 0 Then
                                MsgBox "NHKの契約があるが、請求されない可能性があります。チェックしてください。個人ID：" & rs_shousai!kojinid
                            End If
                        Else
                            MsgBox "NHKの契約があるが、請求されない可能性があります。チェックしてください。個人ID：" & rs_shousai!kojinid
                        End If
                        
                   ' End If
                 rs_shousai.MoveNext
                  
                Loop
                    
               rs_shousai.Close
           End If
           
           db4.Close

    End If
    
        
   FRMOMACHI.Show
    FRMOMACHI.lblmsg.Caption = "個人データの契約の内容（契約状況とＮＨＫデータ）のチェック中・・"
    DoEvents
    
    
 ncncnc = 0
    '現金人の抽出
    sql_comb = "SELECT kojin.*" & _
                " FROM  kojin where seikyuu='1' order by kojinid"
    If FcSQlGet(rs_comb, sql_comb, PrMsg) = True Then
    FRMOMACHI.pb.Max = rs_comb.RecordCount
        rs_comb.MoveFirst
       Do Until rs_comb.EOF
       
       
       
            rs_comb.MoveNext
            ncncnc = ncncnc + 1
            FRMOMACHI.pb.Value = ncncnc
            DoEvents
        Loop
        rs_comb.Close
            
   End If
    

        
   'Log
            
         sagyou_msg = "Toolsチェック：Toolsチェックを実施し、完了しました。"
        log_sagyou sagyou_log_path, sagyou_msg
       
        
        
        MsgBox "終了！"
    Else
        MsgBox "error"
    End If
    
    
    
    
    
    
End If
    
    
    
    
    
    


Unload FRMOMACHI
Exit Sub

verup_new2:
    MsgBox "error_sql"
    rs_comb.Close
    Exit Sub
End Sub






Function a_jidoukoushin() As Boolean

'*********************************
'自動個人データの更新
'*********************************

Dim fdkoushinrs  As Recordset, koushinrs As Recordset, tsuginohi As String, n11 As Long, n12 As Integer

Dim moto_soukei As Long, moto_minyuukin As Long, moto_sonota_shoukei As Long, moto_shu_shoukei As Long, oyasumi As Integer
Dim moto_shu_seikyuu As Long, moto_sonota_seikyuu As Long, keisankin As Long, moto_nhk As Long, moto_kuri As String

Dim nhk_id As String, moto_jikaibi As String, t_n As String, t_m As String, t_n2 As Long, t_m2 As Long
Dim seikyuukingaku_nhk As String, seikyuusumi As String, moto_shurui As String

Dim nhk_henkou As Integer, nhk_hikiotoshi As Integer, nhk_resp As Integer, new_seikyuukaishi As String
Dim kasantsuki As Integer, keisan_soukei As Long, new_nhk_kin As Long, koushnid As String
Dim t_new_m As Integer, t_new_n As Long, newjikai As String, new_nhk_shurui As String
Dim motokojinid As String, hikakunohi As String
Dim moto_f_kasuu As Integer, kyounohi As String, imatorihiki As Integer

 'NHK集計データ作成
Dim sql_nhksakusei As String, rs_nhksakusei As Recordset, shiteinhkbi As String
Dim conutid As String, newcountid As String, nhk_data_dou As Long, nhk_data_dou2 As Long

Dim kadou_keiyaku_seikyuugaku As Long, koushinrs2 As Recordset, newsadou As String, nashinashi As Integer
Dim henkanmotoid As String, henkoukuchisuu As Integer, bimyou As Integer, doudoudou As Integer
Dim henkousakikinngaku As Long, ssllss As Long, koushinrs3 As Recordset, subete_kyuushi As Integer


Dim s_koushin_data() As String
Dim s_koushin_count As Integer
Dim s_koushin_data2() As String
Dim s_koushin_count2 As Integer
Dim s_counter As Integer





frmrogo.lblshinkou.Caption = "金額０円の口座契約者の抽出中・・"
DoEvents

'口座の契約者探し
sql = "SELECT * From kojin" & _
    " WHERE seikyuu ='0' and soukei='0'"
If FcSQlGet(fdkoushinrs, sql, PrMsg) = True Then

     s_koushin_count = fdkoushinrs.RecordCount
    ReDim s_koushin_data(s_koushin_count)
    s_counter = 0


    fdkoushinrs.MoveFirst
    Do Until fdkoushinrs.EOF
        motokojinid = fdkoushinrs!kojinid
        
          s_koushin_data(s_counter) = fdkoushinrs!kojinid
             s_counter = s_counter + 1
        
        frmrogo.lblshinkou.Caption = "口座契約者の更新中・・（ID:" & motokojinid & "）"
        DoEvents

        If motokojinid = "00830" Or motokojinid = "00832" Or motokojinid = "0833" Or motokojinid = "02772" Then
            'MsgBox "Hit"
        End If
        
        
        
        '請求総合計    新maekin用
            If Trim(fdkoushinrs!soukei) = "" Then
                moto_soukei = 0
            Else
                moto_soukei = CLng(fdkoushinrs!soukei)
                keisan_soukei = CLng(fdkoushinrs!soukei)
            End If
        '未入金
            If Trim(fdkoushinrs!minyuukin) = "" Then
                moto_minyuukin = 0
            Else
                moto_minyuukin = CLng(fdkoushinrs!minyuukin)
            End If
        'その他小計
            If Trim(fdkoushinrs!jikaita) = "" Then
                moto_sonota_shoukei = 0
            Else
                moto_sonota_shoukei = CLng(fdkoushinrs!jikaita)
            End If
        '主契約小計
            If Trim(fdkoushinrs!jikaishu) = "" Then
                moto_shu_shoukei = 0
            Else
                moto_shu_shoukei = CLng(fdkoushinrs!jikaishu)
            End If
        '主契約請求金
           If Trim(fdkoushinrs!jikaikin) = "" Then
               moto_shu_seikyuu = 0
           Else
               moto_shu_seikyuu = CLng(fdkoushinrs!jikaikin)
           End If
        'その他請求金
           If Trim(fdkoushinrs!sonotakin) = "" Then
               moto_sonota_seikyuu = 0
           Else
               moto_sonota_seikyuu = CLng(fdkoushinrs!sonotakin)
           End If
           
        'NHK金額
            If Trim(fdkoushinrs!f4) = "" Then
                moto_nhk = 0
            Else
                moto_nhk = CLng(fdkoushinrs!f4)
            End If
        '繰越状態
            If Trim(fdkoushinrs!kajou) = "0" Then
                moto_kuri = "0"
                oyasumi = 0
            ElseIf Trim(fdkoushinrs!kajou) = "1" Then
                moto_kuri = "1"
                oyasumi = 1 '6ヶ月おき
            Else
                moto_kuri = Trim(fdkoushinrs!kajou)
                oyasumi = 2
            End If
            moto_kuri = "0"
            oyasumi = 0
        '取引状態
            If Trim(fdkoushinrs!f3) = "1" Then
                imatorihiki = 1
            Else
                imatorihiki = 0
            End If
       '次の日程
            n11 = CLng(Format(Date, "yyyy"))
            n12 = CInt(Format(Date, "mm")) + 6
            If n12 > 12 Then
                n12 = n12 - 12
                n11 = n11 + 1
            End If
            tsuginohi = Format(n11, "0000") & Format(n12, "00")
       '比較の日程
            n11 = CLng(Format(Date, "yyyy"))
            n12 = CInt(Format(Date, "mm")) + 1
            If n12 > 12 Then
                n12 = n12 - 12
                n11 = n11 + 1
            End If
            hikakunohi = Format(n11, "0000") & Format(n12, "00")
        '今日の日
            n11 = CLng(Format(Date, "yyyy"))
            n12 = CInt(Format(Date, "mm"))
            kyounohi = Format(n11, "0000") & Format(n12, "00")
            
        
        '更新処理(NHK)
        If kidou4 = False Then
            a_jidoukoushin = False
            Exit Function
        End If
        nhk_henkou = 0
        new_nhk_kin = 0
        sql = "select*from nhk_m where kojinid='" & motokojinid & "'"
        nhk_resp = FcSQlGet4(koushinrs, sql, PrMsg)
        If nhk_resp > 0 Then
        

            nhk_id = koushinrs!nhkid
            moto_jikaibi = koushinrs!jikai
            t_n = Mid(koushinrs!jikai, 1, 4) '次回年
            t_n2 = CLng(t_n)
            t_m = Mid(koushinrs!jikai, 5, 2) '次回月
            t_m2 = CLng(t_m)
            seikyuukingaku_nhk = CLng(koushinrs!kingaku) '金額
            seikyuusumi = Trim(koushinrs!n2)
            moto_shurui = koushinrs!shurui '契約種類
                    
            '次回請求日が同じなら請求する
                If CLng(Format(Date, "yyyy")) = t_n2 Then
                    If t_m2 <> 1 Then
                        If CInt(Format(Date, "mm")) = t_m2 - 1 Then
                            nhk_henkou = 1
                            new_seikyuukaishi = "1"
                            new_nhk_kin = seikyuukingaku_nhk
                        Else
                            If CInt(Format(Date, "mm")) = t_m2 Then
                                If seikyuusumi = "1" Then
                                    nhk_henkou = 2
                                    new_seikyuukaishi = "0"
                                    new_nhk_kin = 0
                                End If
                            End If
                        End If
                    End If
                Else
                    If CLng(Format(Date, "yyyy")) = t_n2 - 1 Then
                        If t_m2 = 1 Then
                            If CInt(Format(Date, "mm")) = 12 Then
                                nhk_henkou = 1
                                new_seikyuukaishi = "1"
                                new_nhk_kin = seikyuukingaku_nhk
                            End If
                        End If
                    End If
                End If
            koushinrs.Close
                    
            If nhk_henkou = 1 Then
                sql = "UPDATE nhk_m SET " & _
                        "n2 = '" & new_seikyuukaishi & "'" & _
                        " WHERE nhkid = '" & nhk_id & "'"
                db4.Execute sql, dbSQLPassThrough
            ElseIf nhk_henkou = 2 Then
                If Len(tsuginohi) = 6 Then
                    tsuginohi = tsuginohi & "00"
                End If
                sql = "UPDATE nhk_m SET " & _
                        "zenkai = '" & moto_jikaibi & "'," & _
                        "jikai = '" & tsuginohi & "'," & _
                        "n2 = '" & new_seikyuukaishi & "'" & _
                        " WHERE nhkid = '" & nhk_id & "'"
                db4.Execute sql, dbSQLPassThrough
            
            End If
            
        ElseIf nhk_resp = -1 Then
                a_jidoukoushin = False
                db4.Close
                Exit Function
        End If
        db4.Close
        
         '修正箇所
        Dim keisangaku As Long, keisangaku2 As Long, rs_q As Recordset
        nashinashi = 0
        keisangaku = 0
        keisangaku2 = 0
        
         '修正箇所
        If moto_shu_seikyuu = 0 And moto_sonota_seikyuu = 0 And moto_nhk = 0 Then
                
                Dim keiyakusuu As Long, zenlen As Integer, atolen As Integer, motod As String
                 
                sql = "SELECT * FROM keiyaku WHERE kojinid ='" & motokojinid & "' and kadou = '0'"

                If FcSQlGet(rs_q, sql, PrMsg) = False Then
                    nashinashi = 1
                Else
                    rs_q.MoveFirst
                    Do Until rs_q.EOF
                        If Trim(rs_q!kingaku) <> "" Then
                            keisangaku = keisangaku + CLng(rs_q!kingaku)
                        End If
                        rs_q.MoveNext
                    Loop
                    rs_q.Close
                End If
        End If
        
       
                '比較'更新処理
                If imatorihiki = 1 Then
                    If nhk_henkou = 1 Then
                        keisan_soukei = keisan_soukei + new_nhk_kin
                        sql = "update kojin set " & _
                            "f4 = '" & CStr(new_nhk_kin) & "'," & _
                            "soukei= '" & CStr(keisan_soukei) & "'," & _
                            "jikaibi = '" & CStr(kyounohi) & "'" & _
                            " where kojinid ='" & motokojinid & "'"
                    Else
                        sql = "update kojin set " & _
                            "jikaibi = '" & CStr(kyounohi) & "'" & _
                            " where kojinid ='" & motokojinid & "'"
                    End If
                Else
                        If oyasumi = 0 Then
                            If nashinashi = 1 Then
                                If nhk_henkou = 1 Then
                                    sql = "update kojin set " & _
                                        "f4 = '" & CStr(new_nhk_kin) & "'," & _
                                        "soukei= '" & CStr(new_nhk_kin) & "'" & _
                                        " where kojinid ='" & motokojinid & "'"
                                Else
                                    sql = "update kojin set " & _
                                        "jikaibi = '" & CStr(kyounohi) & "'" & _
                                        " where kojinid ='" & motokojinid & "'"
                                End If
                            Else
                                    If keisangaku = moto_shu_shoukei Then
                                        keisan_soukei = new_nhk_kin + moto_shu_shoukei
                                        keisangaku2 = moto_shu_shoukei
                                    Else
                                        keisan_soukei = new_nhk_kin + keisangaku
                                        keisangaku2 = keisangaku
                                    End If
                                   '個人データの変更
                                   sql = "update kojin set " & _
                                       "f4 = '" & CStr(new_nhk_kin) & "'," & _
                                       "kajou = '0'," & _
                                       "f1 = '0'," & _
                                       "jikaikin = '" & CStr(keisangaku2) & "'," & _
                                       "jikaibi = '" & CStr(kyounohi) & "'," & _
                                       "soukei= '" & CStr(keisan_soukei) & "'" & _
                                       " where kojinid ='" & motokojinid & "'"
                                    
                                    newsadou = "0"
                                    
                            End If
                        Else
                            If hikakunohi = moto_kuri Or new_nhk_kin <> 0 Then
                                sql = "SELECT sum(kingaku) as kikikinn" & _
                                            " From keiyaku" & _
                                            " WHERE (((keiyaku.kojinid)='" & motokojinid & "') and ((keiyaku.kadou)='0'));"
                                If FcSQlGet(koushinrs2, sql, PrMsg) = True Then
                                    If IsNull(koushinrs2!kikikinn) Then
                                        kadou_keiyaku_seikyuugaku = 0
                                        koushinrs2.Close
                                    Else
                                        kadou_keiyaku_seikyuugaku = koushinrs2!kikikinn
                                        koushinrs2.Close
                                    End If
                                Else
                                    kadou_keiyaku_seikyuugaku = 0
                                End If
                                keisan_soukei = new_nhk_kin + kadou_keiyaku_seikyuugaku
                                   '個人データの変更
                                   sql = "update kojin set " & _
                                       "f4 = '" & CStr(new_nhk_kin) & "'," & _
                                       "kajou = '1'," & _
                                       "jikaikin = '" & CStr(kadou_keiyaku_seikyuugaku) & "'," & _
                                       "jikaibi = '" & CStr(kyounohi) & "'," & _
                                       "soukei= '" & CStr(keisan_soukei) & "'" & _
                                       " where kojinid ='" & motokojinid & "'"
                            Else
                               If oyasumi <> 1 Then
                                   sql = "update kojin set " & _
                                       "jikaibi = '" & CStr(kyounohi) & "'" & _
                                       " where kojinid ='" & motokojinid & "'"
                               Else
                                   sql = "update kojin set " & _
                                        "kajou = '" & tsuginohi & "'," & _
                                       "jikaibi = '" & CStr(kyounohi) & "'" & _
                                       " where kojinid ='" & motokojinid & "'"
                               End If
                            End If
                            If oyasumi = 1 Then
                                newsadou = "1"
                            Else        '2
                                If hikakunohi = moto_kuri Then
                                    newsadou = "0"
                                Else
                                    newsadou = "1"
                                End If
                            End If
                          End If
                          
                            '作動伝票の更新
                            Dim sql_sadou As String, rs_sadou As Recordset, sql_sadou2 As String
                            sql_sadou = "SELECT keiyaku.keiyakuid" & _
                                        " FROM keiyaku RIGHT JOIN kojin ON keiyaku.kojinid = kojin.kojinid" & _
                                        " WHERE keiyaku.kadou ='0' AND keiyaku.shurui ='01' and kojin.kojinid ='" & motokojinid & "'"
                            If FcSQlGet(rs_sadou, sql_sadou, PrMsg) = True Then
                                sql_sadou2 = "UPDATE keiyaku SET sadou = '" & newsadou & "'" & _
                                              " WHERE keiyakuid = '" & rs_sadou!keiyakuid & "'"
                                On Error GoTo errko1
                                db.Execute sql_sadou2, dbSQLPassThrough
                                On Error GoTo 0
                                rs_sadou.Close
                            End If
                                  
                        End If
                        
                 On Error GoTo errko12
                db.Execute sql, dbSQLPassThrough
                On Error GoTo 0
        
        fdkoushinrs.MoveNext
    Loop
    fdkoushinrs.Close
    

End If



'請求書契約者探し
frmrogo.lblshinkou.Caption = "現金扱いの契約者の抽出中・・"
DoEvents


sql = "SELECT * From kojin" & _
    " WHERE seikyuu <>'0'"
    
new_nhk_kin = 0

If FcSQlGet(fdkoushinrs, sql, PrMsg) = True Then

   
        s_koushin_count2 = fdkoushinrs.RecordCount
        ReDim s_koushin_data2(s_koushin_count2)
        s_counter = 0
        
    fdkoushinrs.MoveFirst
    Do Until fdkoushinrs.EOF
        If fdkoushinrs!kojinid = "03070" Then
           ' MsgBox "fix"
        End If
        
 
        
         s_koushin_data2(s_counter) = fdkoushinrs!kojinid
        s_counter = s_counter + 1

        
       frmrogo.lblshinkou.Caption = "現金扱いの契約者の更新中・・（ID:" & fdkoushinrs!kojinid & "）"
        DoEvents
        
        If fdkoushinrs!soukei <> "0" Then
            If fdkoushinrs!soukei <> fdkoushinrs!minyuukin Then
                If shousaiset(fdkoushinrs!kojinid, Format(Date, "yyyymm"), CLng(fdkoushinrs!soukei), , , "1") = -1 Then
                    MsgBox "詳細データベースへの書き込みに失敗しました。"
                    Screen.MousePointer = 0
                    a_jidoukoushin = False
                    Exit Function
                End If
            End If
        End If

       
        
        motokojinid = fdkoushinrs!kojinid
        '請求総合計    新maekin用
            If Trim(fdkoushinrs!soukei) = "" Then
                moto_soukei = 0
            Else
                moto_soukei = CLng(fdkoushinrs!soukei)
            End If
        '未入金
            If Trim(fdkoushinrs!minyuukin) = "" Then
                moto_minyuukin = 0
            Else
                moto_minyuukin = CLng(fdkoushinrs!minyuukin)
            End If
        'その他小計
            If Trim(fdkoushinrs!jikaita) = "" Then
                moto_sonota_shoukei = 0
            Else
                moto_sonota_shoukei = CLng(fdkoushinrs!jikaita)
            End If
        '主契約小計
            If Trim(fdkoushinrs!jikaishu) = "" Then
                moto_shu_shoukei = 0
            Else
                moto_shu_shoukei = CLng(fdkoushinrs!jikaishu)
            End If
        '主契約請求金
           If Trim(fdkoushinrs!jikaikin) = "" Then
               moto_shu_seikyuu = 0
           Else
               moto_shu_seikyuu = CLng(fdkoushinrs!jikaikin)
           End If
        'その他請求金
           If Trim(fdkoushinrs!sonotakin) = "" Then
               moto_sonota_seikyuu = 0
           Else
               moto_sonota_seikyuu = CLng(fdkoushinrs!sonotakin)
           End If
           
        'NHK金額
            If Trim(fdkoushinrs!f4) = "" Then
                moto_nhk = 0
            Else
                moto_nhk = CLng(fdkoushinrs!f4)
            End If
        '繰越状態
            If Trim(fdkoushinrs!kajou) = "0" Then
                moto_kuri = "0"
                oyasumi = 0
            ElseIf Trim(fdkoushinrs!kajou) = "1" Then
                moto_kuri = "1"
                oyasumi = 1 '6ヶ月おき
            Else
                            moto_kuri = Trim(fdkoushinrs!kajou)
                            oyasumi = 2
            End If
            
            moto_kuri = "0"
            oyasumi = 0
            
           '引き落ちない回数
            If Trim(fdkoushinrs!fkaisuu) = "" Or Trim(fdkoushinrs!fkaisuu) = "0" Then
                moto_f_kasuu = 0
            Else
                moto_f_kasuu = CInt(fdkoushinrs!fkaisuu)
            End If
           
       '比較の日程
            n11 = CLng(Format(Date, "yyyy"))
            n12 = CInt(Format(Date, "mm")) + 1
            If n12 > 12 Then
                n12 = n12 - 12
                n11 = n11 + 1
            End If
            hikakunohi = Format(n11, "0000") & Format(n12, "00")
        '今日の日
            If Trim(fdkoushinrs!jikaibi) = "" Or Trim(fdkoushinrs!jikaibi) = "0" Then
                '次回の日
                n11 = CLng(Format(Date, "yyyy"))
                n12 = CInt(Format(Date, "mm")) + 6
                If n12 > 12 Then
                    n12 = n12 - 12
                    n11 = n11 + 1
                End If
                tsuginohi = Format(n11, "0000") & Format(n12, "00")
                n11 = CLng(Format(Date, "yyyy"))
                n12 = CInt(Format(Date, "mm"))
                kyounohi = Format(n11, "0000") & Format(n12, "00")
            Else
                '次回の日
                n11 = CLng(Mid(Trim(fdkoushinrs!jikaibi), 1, 4))
                n12 = CInt(Mid(Trim(fdkoushinrs!jikaibi), 5, 2)) + 7
                If n12 > 12 Then
                    n12 = n12 - 12
                    n11 = n11 + 1
                End If
                tsuginohi = Format(n11, "0000") & Format(n12, "00")
                n11 = CLng(Mid(Trim(fdkoushinrs!jikaibi), 1, 4))
                n12 = CInt(Mid(Trim(fdkoushinrs!jikaibi), 5, 2)) + 1
                If n12 > 12 Then
                    n12 = 1
                    n11 = n11 + 1
                End If
                kyounohi = Format(n11, "0000") & Format(n12, "00")
            End If
           
        
        'その他の契約の無0のレコード「２」に変更
         sql = "SELECT keiyaku2.keiyakuid2,keiyaku2.kingaku" & _
             " FROM kojin LEFT JOIN keiyaku2 ON kojin.kojinid = keiyaku2.kojinid" & _
             " WHERE (((kojin.kojinid)='" & motokojinid & "') AND ((keiyaku2.f2)='0'));"
         If FcSQlGet(koushinrs, sql, PrMsg) = True Then
             koushinrs.MoveFirst
             Do Until koushinrs.EOF
                 keisankin = koushinrs!kingaku
                 moto_sonota_shoukei = moto_sonota_shoukei - CLng(keisankin)  'その他小計
                 If moto_sonota_seikyuu = 0 Then
                    moto_sonota_seikyuu = 0
                Else
                    moto_sonota_seikyuu = moto_sonota_seikyuu - CLng(keisankin) 'その他請求金
                End If
                 'motodata(8) = motodata(8) - CLng(keisankin)   '総合計　　　計算用
                 sql = "update keiyaku2 set f2='2' where keiyakuid2 ='" & _
                     koushinrs!keiyakuid2 & "'"
                 On Error GoTo errko13
                     db.Execute sql, dbSQLPassThrough
                 On Error GoTo 0
                 koushinrs.MoveNext
             Loop
             koushinrs.Close
         End If
         
         
        'その他の契約の〆1のレコードを「0」に変更
        sql = "SELECT keiyaku2.keiyakuid2,keiyaku2.kingaku" & _
            " FROM kojin LEFT JOIN keiyaku2 ON kojin.kojinid = keiyaku2.kojinid" & _
            " WHERE (((kojin.kojinid)='" & motokojinid & "') AND ((keiyaku2.f2)='1'));"
        If FcSQlGet(koushinrs, sql, PrMsg) = True Then
            koushinrs.MoveFirst
            Do Until koushinrs.EOF
                keisankin = koushinrs!kingaku
                moto_sonota_seikyuu = moto_sonota_seikyuu + CLng(keisankin) 'その他請求金
                'motodata(8) = motodata(8) + CLng(keisankin)   '総合計　　　計算用
                sql = "update keiyaku2 set f2='0' where keiyakuid2 ='" & _
                    koushinrs!keiyakuid2 & "'"
              On Error GoTo errko14
                db.Execute sql, dbSQLPassThrough
              On Error GoTo 0
                koushinrs.MoveNext
            Loop
            koushinrs.Close
        End If
        If moto_shu_seikyuu = 0 And moto_soukei = 0 Then
            sql = "SELECT keiyaku.keiyakuid,keiyaku.kingaku" & _
                " FROM kojin LEFT JOIN keiyaku ON kojin.kojinid = keiyaku.kojinid" & _
                " WHERE (((kojin.kojinid)='" & motokojinid & "') AND ((keiyaku.kadou)='0'));"
            If FcSQlGet(koushinrs, sql, PrMsg) = True Then
                koushinrs.MoveFirst
                Do Until koushinrs.EOF
                    keisankin = koushinrs!kingaku
                    moto_shu_seikyuu = moto_shu_seikyuu + CLng(keisankin) '主契約請求金
                    koushinrs.MoveNext
                Loop
                koushinrs.Close
            End If
        End If
        
        
        '主契約の〆のレコードを「０」に変更
        sql = "SELECT keiyaku.keiyakuid,keiyaku.kingaku" & _
            " FROM kojin LEFT JOIN keiyaku ON kojin.kojinid = keiyaku.kojinid" & _
            " WHERE (((kojin.kojinid)='" & motokojinid & "') AND ((keiyaku.shime)='1'));"
        If FcSQlGet(koushinrs, sql, PrMsg) = True Then
            koushinrs.MoveFirst
            Do Until koushinrs.EOF
                keisankin = koushinrs!kingaku
                moto_shu_seikyuu = moto_shu_seikyuu + CLng(keisankin) '主契約請求金
                'motodata(8) = motodata(8) + CLng(keisankin)   '総合計　　　計算用
                sql = "update keiyaku set shime='0' where keiyakuid ='" & _
                    koushinrs!keiyakuid & "'"
               On Error GoTo errko15
                db.Execute sql, dbSQLPassThrough
               On Error GoTo 0
                koushinrs.MoveNext
            Loop
            koushinrs.Close
        End If
        
        '主契約の９６－９９のレコードを「０１」に変更
        sql = "SELECT keiyaku.keiyakuid,keiyaku.kingaku" & _
            " FROM kojin LEFT JOIN keiyaku ON kojin.kojinid = keiyaku.kojinid" & _
            " WHERE kojin.kojinid='" & motokojinid & "' AND keiyaku.kadou='0'" & _
            " and (keiyaku.shurui='96' or keiyaku.shurui='97' or keiyaku.shurui='98' or keiyaku.shurui='99')"
        If FcSQlGet(koushinrs, sql, PrMsg) = True Then
            koushinrs.MoveFirst
            Do Until koushinrs.EOF
                keisankin = koushinrs!kingaku
                moto_shu_shoukei = moto_shu_shoukei - CLng(keisankin) + 6000 '主契約小計
                sql = "update keiyaku set shurui='01',kingaku='6000',sadou = '0' where keiyakuid ='" & _
                    koushinrs!keiyakuid & "'"
               On Error GoTo errko16
                db.Execute sql, dbSQLPassThrough
               On Error GoTo 0
                sql = "update kojin set jikaishu='" & CStr(moto_shu_shoukei) & "' where kojinid ='" & _
                    motokojinid & "'"
               On Error GoTo errko17
                db.Execute sql, dbSQLPassThrough
               On Error GoTo 0
                koushinrs.MoveNext
            Loop
            koushinrs.Close
        End If
        ssllss = 0
        subete_kyuushi = 0
        If oyasumi = 2 Then
            
            sql = "SELECT kojin.f4,kojin.kojinid,kojin.name, kojin.minyuukin, kojin.jikaishu, kojin.jikaikin, kojin.soukei,keiyaku.keiyakuid,keiyaku.kuchisuu,keiyaku.kingaku" & _
                        " FROM keiyaku RIGHT JOIN kojin ON keiyaku.kojinid = kojin.kojinid" & _
                        " WHERE keiyaku.kadou ='0' AND keiyaku.shurui ='01' and kojin.kojinid ='" & motokojinid & "'"
            If FcSQlGet(koushinrs, sql, PrMsg) = True Then
            '主契約6ヶ月が存在する場合
                    henkanmotoid = koushinrs!keiyakuid
                    henkoukuchisuu = CInt(koushinrs!kuchisuu)
                    sql = "SELECT sum(kingaku) as kikikinkin" & _
                                " From keiyaku" & _
                                " WHERE (((keiyaku.kojinid)='" & motokojinid & "') AND ((keiyaku.shurui)<>'01') AND ((keiyaku.kadou)='0'));"
                    If FcSQlGet(koushinrs2, sql, PrMsg) = True Then
                    'それ以外の稼動する契約が存在する場合
                        If IsNull(koushinrs2!kikikinkin) Or Trim(koushinrs2!kikikinkin) = "" Or Trim(koushinrs2!kikikinkin) = "0" Then
                            sql = "SELECT *" & _
                                        " From keiyaku" & _
                                        " WHERE (((keiyaku.kojinid)='" & motokojinid & "') AND ((keiyaku.shurui)<>'01'));"
                            If FcSQlGet(koushinrs2, sql, PrMsg) = True Then
                                koushinrs2.MoveFirst
                                Do Until koushinrs2.EOF
                                    If koushinrs2!kadou <> 0 Then
                                        ssllss = ssllss + koushinrs2!kingaku
                                    End If
                                    koushinrs2.MoveNext
                                Loop
                                koushinrs2.Close
                            End If
                        Else
                            bimyou = 1
                            'その他の主契約の金額の合計がある場合
                           If moto_kuri = hikakunohi Then
                           '更新するべき月なら
                                    doudoudou = 1
                                     koushinrs2.Close
                                    henkousakikinngaku = henkoukuchisuu * 1000
                                    moto_shu_shoukei = moto_shu_shoukei - (6000 - 1000) * henkoukuchisuu '主契約小計
                                    If moto_kuri = "0" Or moto_kuri = "1" Then
                                        moto_shu_seikyuu = moto_shu_seikyuu + 1000 * henkoukuchisuu '主契約請求金
                                    End If
                                    '6ヶ月を1ヶ月にへんこう
                                    sql = "UPDATE keiyaku SET shurui = '06',kingaku='" & henkousakikinngaku & "',sadou = '0'" & _
                                      " WHERE keiyakuid = '" & henkanmotoid & "'"
                                    On Error GoTo errko1
                                    db.Execute sql, dbSQLPassThrough
                                    On Error GoTo 0
                                     sql = "update kojin set jikaishu='" & CStr(moto_shu_shoukei) & "' where kojinid ='" & _
                                         motokojinid & "'"
                                    On Error GoTo errko1
                                     db.Execute sql, dbSQLPassThrough
                                    On Error GoTo 0
                                    moto_minyuukin = moto_minyuukin + moto_soukei
                            End If
                         End If
                    End If
                     koushinrs.Close
            Else
            '主契約6ヶ月が存在しない場合
                    sql = "SELECT kojin.f4,kojin.kojinid,kojin.name, kojin.minyuukin, kojin.jikaishu, kojin.jikaikin, kojin.soukei,keiyaku.keiyakuid,keiyaku.kuchisuu,keiyaku.kingaku" & _
                        " FROM keiyaku RIGHT JOIN kojin ON keiyaku.kojinid = kojin.kojinid" & _
                        " WHERE keiyaku.kadou ='1' AND keiyaku.shurui ='01' and kojin.kojinid ='" & motokojinid & "'"
                    If FcSQlGet(koushinrs2, sql, PrMsg) = True Then
                    '主契約6ヶ月の休止状態が存在する場合
                        sql = "SELECT sum(kingaku) as kikikinkin" & _
                                    " From keiyaku" & _
                                    " WHERE (((keiyaku.kojinid)='" & motokojinid & "') AND ((keiyaku.shurui)<>'01') AND ((keiyaku.kadou)='0'));"
                        If FcSQlGet(koushinrs3, sql, PrMsg) = False Then
                        'それ以外の稼動する契約が存在しない場合
                            subete_kyuushi = 1
                        Else
                            If koushinrs3!kikikinkin = 0 Or IsNull(koushinrs3!kikikinkin) Then
                                subete_kyuushi = 1
                            End If
                        End If
                    End If
            End If
        End If
        
        
        
        
        
        If fdkoushinrs!soukei <> "0" Then
                moto_f_kasuu = moto_f_kasuu + 1
        End If
                '更新処理
                
                
        If oyasumi = 1 Then '6カ月おき
               ' keisan_soukei = moto_shu_seikyuu + moto_sonota_seikyuu + new_nhk_kin
                'moto_minyuukin = moto_minyuukin + moto_soukei
                moto_minyuukin = moto_soukei
                sql = "update kojin set " & _
                    "minyuukin = '" & CStr(moto_minyuukin) & "'," & _
                    "f4 = '" & CStr(new_nhk_kin) & "'," & _
                    "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                    "fkaisuu = '" & CStr(moto_f_kasuu) & "'," & _
                    "jikaibi = '" & CStr(kyounohi) & "'," & _
                    "jikaikin = '0'," & _
                    "kajou = '" & CStr(tsuginohi) & "'," & _
                    "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                    "soukei= '" & CStr(moto_minyuukin) & "'" & _
                    " where kojinid ='" & motokojinid & "'"
                    
                    newsadou = "1"
        ElseIf oyasumi = 2 Then '6カ月おき
            If moto_kuri <> hikakunohi Then
            'If moto_kuri > kyounohi Then
                'moto_minyuukin = moto_minyuukin + moto_soukei
                keisan_soukei = moto_shu_seikyuu + moto_sonota_seikyuu + new_nhk_kin
                sql = "update kojin set " & _
                    "minyuukin = '" & CStr(moto_minyuukin) & "'," & _
                    "f4 = '" & CStr(new_nhk_kin) & "'," & _
                    "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                    "fkaisuu = '" & CStr(moto_f_kasuu) & "'," & _
                    "jikaibi = '" & CStr(kyounohi) & "'," & _
                    "jikaikin = '0'," & _
                    "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                    "soukei= '" & CStr(moto_minyuukin) & "'" & _
                    " where kojinid ='" & motokojinid & "'"
                    newsadou = "1"
            Else
                'moto_minyuukin = moto_minyuukin + moto_soukei
                 If subete_kyuushi = 1 Then
                    keisan_soukei = 0
                    keisan_soukei = keisan_soukei + moto_minyuukin
                    moto_shu_shoukei = 0
                 Else
                    keisan_soukei = moto_shu_shoukei + moto_sonota_seikyuu + new_nhk_kin
                    keisan_soukei = keisan_soukei + moto_minyuukin
                End If
                '個人データの変更
                If ssllss = 0 Then
                    sql = "update kojin set " & _
                        "minyuukin = '" & CStr(moto_minyuukin) & "'," & _
                        "f4 = '" & CStr(new_nhk_kin) & "'," & _
                        "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                        "fkaisuu = '" & CStr(moto_f_kasuu) & "'," & _
                        "jikaikin = '" & CStr(moto_shu_shoukei) & "'," & _
                        "jikaibi = '" & CStr(kyounohi) & "'," & _
                        "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                        "soukei= '" & CStr(keisan_soukei) & "'," & _
                        "kajou = '1'" & _
                        " where kojinid ='" & motokojinid & "'"
                 Else
                    moto_shu_shoukei = moto_shu_shoukei - ssllss
                    keisan_soukei = keisan_soukei - ssllss
                    sql = "update kojin set " & _
                        "minyuukin = '" & CStr(moto_minyuukin) & "'," & _
                        "f4 = '" & CStr(new_nhk_kin) & "'," & _
                        "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                        "fkaisuu = '" & CStr(moto_f_kasuu) & "'," & _
                        "jikaikin = '" & CStr(moto_shu_shoukei) & "'," & _
                        "jikaibi = '" & CStr(kyounohi) & "'," & _
                        "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                        "soukei= '" & CStr(keisan_soukei) & "'," & _
                        "kajou = '1'" & _
                        " where kojinid ='" & motokojinid & "'"
                 End If
                    newsadou = "0"
            End If
        Else
            keisan_soukei = moto_shu_seikyuu + moto_sonota_seikyuu + new_nhk_kin
            moto_minyuukin = moto_soukei
            keisan_soukei = keisan_soukei + moto_minyuukin
                '個人データの変更
                sql = "update kojin set " & _
                    "minyuukin = '" & CStr(moto_minyuukin) & "'," & _
                    "f4 = '" & CStr(new_nhk_kin) & "'," & _
                    "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                    "fkaisuu = '" & CStr(moto_f_kasuu) & "'," & _
                    "jikaikin = '" & CStr(moto_shu_seikyuu) & "'," & _
                    "jikaibi = '" & CStr(kyounohi) & "'," & _
                    "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                    "soukei= '" & CStr(keisan_soukei) & "'" & _
                    " where kojinid ='" & motokojinid & "'"
                    
                    newsadou = "0"
        End If
        
      On Error GoTo errko18
      db.Execute sql, dbSQLPassThrough
      On Error GoTo 0
                      
        '作動伝票の更新
            sql_sadou = "SELECT keiyaku.keiyakuid" & _
                        " FROM keiyaku RIGHT JOIN kojin ON keiyaku.kojinid = kojin.kojinid" & _
                        " WHERE keiyaku.kadou ='0' AND keiyaku.shurui ='01' and kojin.kojinid ='" & motokojinid & "'"
            If FcSQlGet(rs_sadou, sql_sadou, PrMsg) = True Then
                sql_sadou2 = "UPDATE keiyaku SET sadou = '" & newsadou & "'" & _
                              " WHERE keiyakuid = '" & rs_sadou!keiyakuid & "'"
                On Error GoTo errko19
                db.Execute sql_sadou2, dbSQLPassThrough
                On Error GoTo 0
                rs_sadou.Close
            End If
        fdkoushinrs.MoveNext
    Loop
    fdkoushinrs.Close

    a_jidoukoushin = True

End If


'ここ
For i = 0 To s_koushin_count - 1
     'If yotejikkou(Format(Date, "yyyymm"), s_koushin_data(i)) = -1 Then
     
    'End If
Next

For i = 0 To s_koushin_count2 - 1
     'If yotejikkou(Format(Date, "yyyymm"), s_koushin_data2(i)) = -1 Then
     
    'End If
Next




Exit Function
errko2:
    MsgBox "設定の書き込みに失敗しました。"
    a_jidoukoushin = False
    Exit Function
errsoukatsunn:
    MsgBox "詳細の更新に失敗しました。"
    a_jidoukoushin = False
    Exit Function


errko1:
    MsgBox "設定の書き込みに失敗しました。"
    a_jidoukoushin = False
    Exit Function
errko12:
    MsgBox "設定の書き込みに失敗しました。"
    a_jidoukoushin = False
    Exit Function
errko13:
    MsgBox "設定の書き込みに失敗しました。"
    a_jidoukoushin = False
    Exit Function
errko14:
    MsgBox "設定の書き込みに失敗しました。"
    a_jidoukoushin = False
    Exit Function
errko15:
    MsgBox "設定の書き込みに失敗しました。"
    a_jidoukoushin = False
    Exit Function
errko16:
    MsgBox "設定の書き込みに失敗しました。"
    a_jidoukoushin = False
    Exit Function
errko17:
    MsgBox "設定の書き込みに失敗しました。"
    a_jidoukoushin = False
    Exit Function
errko18:
    MsgBox "設定の書き込みに失敗しました。"
    a_jidoukoushin = False
    Exit Function
errko19:
    MsgBox "設定の書き込みに失敗しました。"
    a_jidoukoushin = False
    Exit Function



End Function
Function a_genkin_nyuukin3(koid As String, kouhi As String, kinga As String, genkou As String) As Boolean
'*********************************
'現金による個人データの更新
'*********************************
Dim sql_gn As String, rs_gn As Recordset, moto_sougoukin As Long, koushnid As String
Dim newmaebi As String, newmaekin As String, newsoukei As String, koushinrs2 As Recordset, newkoushinid As String
Dim nhk_kingaku As Long
                    
Dim moto_soukei As Long, moto_seikyuuhouhou As Long, moto_minyuukin As Long
Dim moto_sonota_shoukei As Long, moto_shu_shoukei As Long, moto_shu_seikyuu As Long
Dim moto_sonota_seikyuu As Long, oldmaebi As String, moto_kuri As String

'もと情報確認
  sql_gn = "select * from kojin where kojinid ='" & koid & "'"
    If FcSQlGet(rs_gn, sql_gn, PrMsg) = False Then
        MsgBox "個人テーブルアクセスエラー"
        a_genkin_nyuukin3 = False
        Exit Function
    Else
        If Trim(rs_gn!maebi) = "" Then
            oldmaebi = Space(1)
        Else
            oldmaebi = Trim(rs_gn!maebi)
        End If
        '新しい前日
        newmaebi = kouhi
        '入金額
        moto_sougoukin = CLng(kinga)
        '請求総合計    新maekin用
            If Trim(rs_gn!soukei) = "" Then
                moto_soukei = 0
            Else
                moto_soukei = CLng(rs_gn!soukei)
            End If
        '請求方法
            If Trim(genkou) = "" Then
                moto_seikyuuhouhou = 0
            Else
                moto_seikyuuhouhou = CLng(genkou)
            End If
        '未入金
            If Trim(rs_gn!minyuukin) = "" Then
                moto_minyuukin = 0
            Else
                moto_minyuukin = CLng(rs_gn!minyuukin)
            End If
        'その他小計
            If Trim(rs_gn!jikaita) = "" Then
                moto_sonota_shoukei = 0
            Else
                moto_sonota_shoukei = CLng(rs_gn!jikaita)
            End If
        '主契約小計
            If Trim(rs_gn!jikaishu) = "" Then
                moto_shu_shoukei = 0
            Else
                moto_shu_shoukei = CLng(rs_gn!jikaishu)
            End If
         '主契約請求金
            If Trim(rs_gn!jikaikin) = "" Then
                moto_shu_seikyuu = 0
            Else
                moto_shu_seikyuu = CLng(rs_gn!jikaikin)
            End If
         'その他請求金
            If Trim(rs_gn!sonotakin) = "" Then
                moto_sonota_seikyuu = 0
            Else
                moto_sonota_seikyuu = CLng(rs_gn!sonotakin)
            End If
        '繰越状態
            If Trim(rs_gn!kajou) = "0" Then '毎回
                moto_kuri = "0"
            ElseIf Trim(rs_gn!kajou) = "1" Then '6ヶ月
                moto_kuri = "1"
            Else
                moto_kuri = Trim(rs_gn!kajou)
            End If
            

            
            moto_sonota_seikyuu = moto_sonota_seikyuu - moto_sougoukin
            moto_sonota_shoukei = moto_sonota_shoukei - moto_sougoukin
            
            newsoukei = moto_soukei - moto_sougoukin
            
            
            sql = "update kojin set  maebi ='" & newmaebi & "'," & _
                    "maekin = '" & CStr(moto_sougoukin) & "'," & _
                    "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                    "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                    "fkaisuu = '0'," & _
                    "soukei= '" & CStr(newsoukei) & "'" & _
                    " where kojinid ='" & koid & "'"
            
        rs_gn.Close
                
        
        On Error GoTo errko1
        db.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        
        '更新記録を残す
            If soukatsu(0, "renban", 8, "", 1, 1) = False Then
                    MsgBox "設定の読み込みに失敗しました。"
                        a_genkin_nyuukin3 = False
                        Exit Function
            End If
            If reg = "0" Then
                reg = ""
            End If
            If reg = vbNullString Then
                koushnid = "00000001"
            ElseIf reg = "9999999" Then
                MsgBox "これ以上の更新記録の作成はできません。管理者に連絡してください。"
                a_genkin_nyuukin3 = False
                Exit Function
            Else
                koushnid = Format$(reg, "000000#")
            End If
            
        
        'koushin.mdbをOPEN
            If kidou2 = False Then
                MsgBox "データベースファイルを開けませんでした。"
                a_genkin_nyuukin3 = False
                Exit Function
            End If
            
            sql = "INSERT INTO koushinkiroku VALUES('" & koushnid & "'," & _
                           "'" & newmaebi & "','" & koid & "'," & _
                           "'" & CStr(moto_sougoukin) & "','" & genkou & "','" & Space(1) & "')"
            On Error GoTo errko2
                db2.Execute sql, dbSQLPassThrough
            On Error GoTo 0
            
            newkoushinid = CStr(CLng(koushnid) + 1)
            If soukatsu(1, "renban", 8, newkoushinid, 1, 0) = False Then
                    MsgBox "設定の書き込みに失敗しました。"
                    a_genkin_nyuukin3 = False
                    Exit Function
            End If
        
            
            
            db2.Close
        a_genkin_nyuukin3 = True
    End If
    
Exit Function

errko1:
    MsgBox "設定の書き込みに失敗しました。"
    a_genkin_nyuukin3 = False
    Exit Function
errko2:
    MsgBox "設定の書き込みに失敗しました。"
    db2.Close
    a_genkin_nyuukin3 = False
    Exit Function
 

End Function
Function a_genkin_nyuukin2(koid As String, kouhi As String, kinga As String, genkou As String) As Boolean
'*********************************
'現金による個人データの更新
'*********************************
Dim sql_gn As String, rs_gn As Recordset, moto_sougoukin As Long, koushnid As String
Dim newmaebi As String, newmaekin As String, newsoukei As String, koushinrs2 As Recordset, newkoushinid As String
Dim nhk_kingaku As Long, newminyuukin As Long
                    
Dim moto_soukei As Long, moto_seikyuuhouhou As Long, moto_minyuukin As Long
Dim moto_sonota_shoukei As Long, moto_shu_shoukei As Long, moto_shu_seikyuu As Long
Dim moto_sonota_seikyuu As Long, oldmaebi As String, moto_kuri As String

'もと情報確認
  sql_gn = "select * from kojin where kojinid ='" & koid & "'"
    If FcSQlGet(rs_gn, sql_gn, PrMsg) = False Then
        MsgBox "個人テーブルアクセスエラー"
        a_genkin_nyuukin2 = False
        Exit Function
    Else
        If Trim(rs_gn!maebi) = "" Then
            oldmaebi = Space(1)
        Else
            oldmaebi = Trim(rs_gn!maebi)
        End If
        '新しい前日
        newmaebi = kouhi
        '入金額
        moto_sougoukin = CLng(kinga)
        '請求総合計    新maekin用
            If Trim(rs_gn!soukei) = "" Then
                moto_soukei = 0
            Else
                moto_soukei = CLng(rs_gn!soukei)
            End If
        '請求方法
            If Trim(genkou) = "" Then
                moto_seikyuuhouhou = 0
            Else
                moto_seikyuuhouhou = CLng(genkou)
            End If
        '未入金
            If Trim(rs_gn!minyuukin) = "" Then
                moto_minyuukin = 0
            Else
                moto_minyuukin = CLng(rs_gn!minyuukin)
            End If
        'その他小計
            If Trim(rs_gn!jikaita) = "" Then
                moto_sonota_shoukei = 0
            Else
                moto_sonota_shoukei = CLng(rs_gn!jikaita)
            End If
        '主契約小計
            If Trim(rs_gn!jikaishu) = "" Then
                moto_shu_shoukei = 0
            Else
                moto_shu_shoukei = CLng(rs_gn!jikaishu)
            End If
         '主契約請求金
            If Trim(rs_gn!jikaikin) = "" Then
                moto_shu_seikyuu = 0
            Else
                moto_shu_seikyuu = CLng(rs_gn!jikaikin)
            End If
         'その他請求金
            If Trim(rs_gn!sonotakin) = "" Then
                moto_sonota_seikyuu = 0
            Else
                moto_sonota_seikyuu = CLng(rs_gn!sonotakin)
            End If
        '繰越状態
            If Trim(rs_gn!kajou) = "0" Then '毎回
                moto_kuri = "0"
            ElseIf Trim(rs_gn!kajou) = "1" Then '6ヶ月
                moto_kuri = "1"
            Else
                moto_kuri = Trim(rs_gn!kajou)
            End If
            
            newminyuukin = moto_minyuukin - moto_sougoukin
            newsoukei = moto_soukei - moto_sougoukin
            sql = "update kojin set  maebi ='" & newmaebi & "'," & _
                    "maekin = '" & CStr(moto_sougoukin) & "'," & _
                    "minyuukin = '" & CStr(newminyuukin) & "'," & _
                    "fkaisuu = '0'," & _
                    "soukei= '" & CStr(newsoukei) & "'" & _
                    " where kojinid ='" & koid & "'"
            
        rs_gn.Close
                
        
        On Error GoTo errko1
        db.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        
        '更新記録を残す
            If soukatsu(0, "renban", 8, "", 1, 1) = False Then
                    MsgBox "設定の読み込みに失敗しました。"
                        a_genkin_nyuukin2 = False
                        Exit Function
            End If
            If reg = "0" Then
                reg = ""
            End If
            If reg = vbNullString Then
                koushnid = "00000001"
            ElseIf reg = "9999999" Then
                MsgBox "これ以上の更新記録の作成はできません。管理者に連絡してください。"
                a_genkin_nyuukin2 = False
                Exit Function
            Else
                koushnid = Format$(reg, "000000#")
            End If
            
        
        'koushin.mdbをOPEN
            If kidou2 = False Then
                MsgBox "データベースファイルを開けませんでした。"
                a_genkin_nyuukin2 = False
                Exit Function
            End If
            
            sql = "INSERT INTO koushinkiroku VALUES('" & koushnid & "'," & _
                           "'" & newmaebi & "','" & koid & "'," & _
                           "'" & CStr(moto_sougoukin) & "','" & genkou & "','" & Space(1) & "')"
            On Error GoTo errko2
                db2.Execute sql, dbSQLPassThrough
            On Error GoTo 0
            
            newkoushinid = CStr(CLng(koushnid) + 1)
            If soukatsu(1, "renban", 8, newkoushinid, 1, 0) = False Then
                    MsgBox "設定の書き込みに失敗しました。"
                    a_genkin_nyuukin2 = False
                    Exit Function
            End If
        
            
            
            db2.Close
        a_genkin_nyuukin2 = True
    End If
    
Exit Function

errko1:
    MsgBox "設定の書き込みに失敗しました。"
    a_genkin_nyuukin2 = False
    Exit Function
errko2:
    MsgBox "設定の書き込みに失敗しました。"
    db2.Close
    a_genkin_nyuukin2 = False
    Exit Function
 

End Function
Function a_genkin_nyuukin(koid As String, kouhi As String, kinga As String, genkou As String) As Boolean
'*********************************
'現金による個人データの更新
'*********************************
Dim sql_gn As String, rs_gn As Recordset, moto_sougoukin As Long, koushnid As String
Dim newmaebi As String, newmaekin As String, newsoukei As String, koushinrs2 As Recordset, newkoushinid As String
Dim nhk_kingaku As Long, koushinrs As Recordset, sonota2hen As Integer
                    
Dim moto_soukei As Long, moto_seikyuuhouhou As Long, moto_minyuukin As Long
Dim moto_sonota_shoukei As Long, moto_shu_shoukei As Long, moto_shu_seikyuu As Long
Dim moto_sonota_seikyuu As Long, oldmaebi As String, moto_kuri As String
Dim shuusei_sougaku As Long, shuusei_shukeiseikyuu As Long

'もと情報確認

sonota2hen = 0

  sql_gn = "select * from kojin where kojinid ='" & koid & "'"
    If FcSQlGet(rs_gn, sql_gn, PrMsg) = False Then
        MsgBox "個人テーブルアクセスエラー"
        a_genkin_nyuukin = False
        Exit Function
    Else
        If Trim(rs_gn!maebi) = "" Then
            oldmaebi = Space(1)
        Else
            oldmaebi = Trim(rs_gn!maebi)
        End If
        '新しい前日
        newmaebi = kouhi
        '入金額
        moto_sougoukin = CLng(kinga)
        '請求総合計    新maekin用
            If Trim(rs_gn!soukei) = "" Then
                moto_soukei = 0
            Else
                moto_soukei = CLng(rs_gn!soukei)
            End If
        '請求方法
            If Trim(genkou) = "" Then
                moto_seikyuuhouhou = 0
            Else
                moto_seikyuuhouhou = CLng(genkou)
            End If
        '未入金
            If Trim(rs_gn!minyuukin) = "" Then
                moto_minyuukin = 0
            Else
                moto_minyuukin = CLng(rs_gn!minyuukin)
            End If
        'その他小計
            If Trim(rs_gn!jikaita) = "" Then
                moto_sonota_shoukei = 0
            Else
                moto_sonota_shoukei = CLng(rs_gn!jikaita)
            End If
        '主契約小計
            If Trim(rs_gn!jikaishu) = "" Then
                moto_shu_shoukei = 0
            Else
                moto_shu_shoukei = CLng(rs_gn!jikaishu)
            End If
         '主契約請求金
            If Trim(rs_gn!jikaikin) = "" Then
                moto_shu_seikyuu = 0
            Else
                moto_shu_seikyuu = CLng(rs_gn!jikaikin)
            End If
         'その他請求金
            If Trim(rs_gn!sonotakin) = "" Then
                moto_sonota_seikyuu = 0
            Else
                moto_sonota_seikyuu = CLng(rs_gn!sonotakin)
                sonota2hen = 1
            End If
        '繰越状態
            If Trim(rs_gn!kajou) = "0" Then '毎回
                moto_kuri = "0"
            ElseIf Trim(rs_gn!kajou) = "1" Then '6ヶ月
                moto_kuri = "1"
            Else
                moto_kuri = Trim(rs_gn!kajou)
            End If
            moto_kuri = "0"
            
       ' If Trim(rs_gn!seikyuu) = "0" Then
       '     shuusei_sougaku = moto_shu_seikyuu
       '     shuusei_shukeiseikyuu = moto_shu_seikyuu
       ' Else
            shuusei_sougaku = 0
            shuusei_shukeiseikyuu = 0
        'End If
            
         'NHK
         If sonota2hen = 1 Then
            If frmtouroku2.lblumu.Caption = "なし" Then '新規
                    sql = "update kojin set  maebi ='" & newmaebi & "'," & _
                            "maekin = '" & CStr(moto_sougoukin) & "'," & _
                            "minyuukin = '0'," & _
                            "jikaikin = '" & shuusei_shukeiseikyuu & "'," & _
                            "jikaita ='0'," & _
                            "fkaisuu = '0'," & _
                            "sonotakin = '0'," & _
                            "soukei= '" & shuusei_sougaku & "'" & _
                            " where kojinid ='" & koid & "'"
            Else
                    sql = "update kojin set  maebi ='" & newmaebi & "'," & _
                            "maekin = '" & CStr(moto_sougoukin) & "'," & _
                            "minyuukin = '0'," & _
                            "jikaikin = '" & shuusei_shukeiseikyuu & "'," & _
                            "jikaita ='0'," & _
                            "fkaisuu = '0'," & _
                            "f4 = '0'," & _
                            "sonotakin = '0'," & _
                            "soukei= '" & shuusei_sougaku & "'" & _
                            " where kojinid ='" & koid & "'"

            End If
        Else
            If frmtouroku2.lblumu.Caption = "なし" Then '新規
                    sql = "update kojin set  maebi ='" & newmaebi & "'," & _
                            "maekin = '" & CStr(moto_sougoukin) & "'," & _
                            "minyuukin = '0'," & _
                            "jikaikin = '" & shuusei_shukeiseikyuu & "'," & _
                            "fkaisuu = '0'," & _
                            "sonotakin = '0'," & _
                            "soukei= '" & shuusei_sougaku & "'" & _
                            " where kojinid ='" & koid & "'"
            Else
                    sql = "update kojin set  maebi ='" & newmaebi & "'," & _
                            "maekin = '" & CStr(moto_sougoukin) & "'," & _
                            "minyuukin = '0'," & _
                            "jikaikin = '" & shuusei_shukeiseikyuu & "'," & _
                            "fkaisuu = '0'," & _
                            "f4 = '0'," & _
                            "sonotakin = '0'," & _
                            "soukei= '" & shuusei_sougaku & "'" & _
                            " where kojinid ='" & koid & "'"

            End If
         End If
        rs_gn.Close
        
        On Error GoTo errko1
        db.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        
        '更新記録を残す
            If soukatsu(0, "renban", 8, "", 1, 1) = False Then
                    MsgBox "設定の読み込みに失敗しました。"
                        a_genkin_nyuukin = False
                        Exit Function
            End If
            If reg = "0" Then
                reg = ""
            End If
            If reg = vbNullString Then
                koushnid = "00000001"
            ElseIf reg = "9999999" Then
                MsgBox "これ以上の更新記録の作成はできません。管理者に連絡してください。"
                a_genkin_nyuukin = False
                Exit Function
            Else
                koushnid = Format$(reg, "000000#")
            End If
            
        
        'koushin.mdbをOPEN
            If kidou2 = False Then
                MsgBox "データベースファイルを開けませんでした。"
                a_genkin_nyuukin = False
                Exit Function
            End If
            
            sql = "INSERT INTO koushinkiroku VALUES('" & koushnid & "'," & _
                           "'" & newmaebi & "','" & koid & "'," & _
                           "'" & CStr(moto_sougoukin) & "','" & genkou & "','" & Space(1) & "')"
            On Error GoTo errko2
                db2.Execute sql, dbSQLPassThrough
            On Error GoTo 0
            
            newkoushinid = CStr(CLng(koushnid) + 1)
            If soukatsu(1, "renban", 8, newkoushinid, 1, 0) = False Then
                    MsgBox "設定の書き込みに失敗しました。"
                    a_genkin_nyuukin = False
                    Exit Function
            End If
        
        'shousaiの更新
            sql = "UPDATE shousai SET  keinyu = '" & newmaebi & "',keist='1'" & _
                        " WHERE kojinid = '" & koid & "' and keist='0'"

            On Error GoTo errsoukatsunn
                 db2.Execute sql, dbSQLPassThrough
            On Error GoTo 0
            
        'その他の契約の無0のレコード「２」に変更
         sql = "SELECT keiyaku2.keiyakuid2,keiyaku2.kingaku" & _
             " FROM kojin LEFT JOIN keiyaku2 ON kojin.kojinid = keiyaku2.kojinid" & _
             " WHERE (((kojin.kojinid)='" & koid & "') AND ((keiyaku2.f2)='0'));"
         If FcSQlGet(koushinrs, sql, PrMsg) = True Then
             koushinrs.MoveFirst
             Do Until koushinrs.EOF
                 sql = "update keiyaku2 set f2='2' where keiyakuid2 ='" & _
                     koushinrs!keiyakuid2 & "'"
                 On Error GoTo errko1
                     db.Execute sql, dbSQLPassThrough
                 On Error GoTo 0
                 koushinrs.MoveNext
             Loop
             koushinrs.Close
            keiyakuset2 (koid)
         End If
            
            
            
        a_genkin_nyuukin = True
    End If
    
Exit Function

errko1:
    MsgBox "設定の書き込みに失敗しました。"
    a_genkin_nyuukin = False
    Exit Function
errko2:
    MsgBox "設定の書き込みに失敗しました。"
    db2.Close
    a_genkin_nyuukin = False
    Exit Function
errsoukatsunn:
    MsgBox "更新設定の書き込みに失敗しました。"
    a_genkin_nyuukin = False
    Exit Function
 
End Function
Function a_fdd_koushin2(fkekka As Long, motokojinid As String, koushinbi As String, houhou As Integer, kinga As Long) As Integer
'*********************************
'FDDによる個人データの更新
'*********************************

Dim fdkoushinrs  As Recordset, koushinrs As Recordset, tsuginohi As String, n11 As Long, n12 As Integer, koushinrs2 As Recordset

Dim moto_soukei As Long, moto_minyuukin As Long, moto_sonota_shoukei As Long, moto_shu_shoukei As Long, oyasumi As Integer
Dim moto_shu_seikyuu As Long, moto_sonota_seikyuu As Long, keisankin As Long, moto_nhk As Long, moto_kuri As String

Dim nhk_id As String, moto_jikaibi As String, t_n As String, t_m As String, t_n2 As Long, t_m2 As Long
Dim seikyuukingaku_nhk As String, seikyuusumi As String, moto_shurui As String

Dim nhk_henkou As Integer, nhk_hikiotoshi As Integer, nhk_resp As Integer, new_seikyuukaishi As String
Dim kasantsuki As Integer, keisan_soukei As Long, new_nhk_kin As Long, koushnid As String
Dim t_new_m As Integer, t_new_n As Long, newjikai As String, new_nhk_shurui As String

Dim moto_f_kasuu As Integer, hikukazu As Integer, henkoushita As Integer, doudoudou As Integer
Dim hikakunohi As String, kyounohi As String, henkanmotoid As String, henkoukuchisuu As String
Dim henkousakikinngaku As Long, newsadou As String, bimyou As Integer
'もと情報確認

'滞納ST
Dim s_tainou_jikkou As Integer

s_tainou_jikkou = 0


  sql = "select * from kojin where kojinid ='" & motokojinid & "'"
    If FcSQlGet(fdkoushinrs, sql, PrMsg) = False Then
        MsgBox "個人テーブルアクセスエラー"
        a_fdd_koushin2 = 9
        Exit Function
    End If
        If motokojinid = "00887" Then
           ' MsgBox "hit"
        End If
        '請求総合計    新maekin　用
            If Trim(fdkoushinrs!soukei) = "" Then
                moto_soukei = 0
            Else
                moto_soukei = CLng(fdkoushinrs!soukei)
            End If
        '未入金
            If Trim(fdkoushinrs!minyuukin) = "" Then
                moto_minyuukin = 0
            Else
                moto_minyuukin = CLng(fdkoushinrs!minyuukin)
            End If
        'その他小計
            If Trim(fdkoushinrs!jikaita) = "" Then
                moto_sonota_shoukei = 0
            Else
                moto_sonota_shoukei = CLng(fdkoushinrs!jikaita)
            End If
        '主契約小計
            If Trim(fdkoushinrs!jikaishu) = "" Then
                moto_shu_shoukei = 0
            Else
                moto_shu_shoukei = CLng(fdkoushinrs!jikaishu)
            End If
        '主契約請求金
           If Trim(fdkoushinrs!jikaikin) = "" Then
               moto_shu_seikyuu = 0
           Else
               moto_shu_seikyuu = CLng(fdkoushinrs!jikaikin)
           End If
        'その他請求金
           If Trim(fdkoushinrs!sonotakin) = "" Then
               moto_sonota_seikyuu = 0
           Else
               moto_sonota_seikyuu = CLng(fdkoushinrs!sonotakin)
           End If
           
        'NHK金額
            If Trim(fdkoushinrs!f4) = "" Then
                moto_nhk = 0
            Else
                moto_nhk = CLng(fdkoushinrs!f4)
            End If
        '繰越状態
            If Trim(fdkoushinrs!kajou) = "0" Then
                moto_kuri = "0"
                oyasumi = 0
            ElseIf Trim(fdkoushinrs!kajou) = "1" Then
                moto_kuri = "1"
                oyasumi = 1 '6ヶ月おき
            Else
                moto_kuri = Trim(fdkoushinrs!kajou)
                oyasumi = 2
            End If
            'ここを変えただけ
            moto_kuri = "0"
            oyasumi = 0
       
       '次の日程
            n11 = CLng(Format(Date, "yyyy"))
            n12 = CInt(Format(Date, "mm")) + 6
            If n12 > 12 Then
                n12 = n12 - 12
                n11 = n11 + 1
            End If
            tsuginohi = Format(n11, "0000") & Format(n12, "00")
       '比較の日程
            n11 = CLng(Format(Date, "yyyy"))
            n12 = CInt(Format(Date, "mm")) + 1
            If n12 > 12 Then
                n12 = n12 - 12
                n11 = n11 + 1
            End If
            hikakunohi = Format(n11, "0000") & Format(n12, "00")
        '今日の日
            n11 = CLng(Format(Date, "yyyy"))
            n12 = CInt(Format(Date, "mm"))
            kyounohi = Format(n11, "0000") & Format(n12, "00")
            
            
           '引き落ちない回数
            If Trim(fdkoushinrs!fkaisuu) = "" Or Trim(fdkoushinrs!fkaisuu) = "0" Then
                moto_f_kasuu = 0
            Else
                moto_f_kasuu = CInt(fdkoushinrs!fkaisuu)
            End If
           
            fdkoushinrs.Close
            
                n11 = CLng(Format(Date, "yyyy"))
                n12 = CInt(Format(Date, "mm")) + 6
                If n12 > 12 Then
                    n12 = n12 - 12
                    n11 = n11 + 1
                End If
                tsuginohi = Format(n11, "0000") & Format(n12, "00")
            
        '更新処理(NHK)
        If kidou4 = False Then
            a_fdd_koushin2 = 9
            Exit Function
        End If
        nhk_henkou = 0
        nhk_hikiotoshi = 0 '初期化　引落できたかの引数
        new_nhk_kin = 0
        sql = "select*from nhk_m where kojinid='" & motokojinid & "'"
        nhk_resp = FcSQlGet4(koushinrs, sql, PrMsg)
        If nhk_resp > 0 Then
        
        If CLng(koushinrs!kingaku) > 0 Then
            nhk_id = koushinrs!nhkid
            moto_jikaibi = koushinrs!jikai
            t_n = Mid(koushinrs!jikai, 1, 4) '次回年
            t_n2 = CLng(t_n)
            t_m = Mid(koushinrs!jikai, 5, 2) '次回月
            t_m2 = CLng(t_m)
            seikyuukingaku_nhk = CLng(koushinrs!kingaku) '金額
            seikyuusumi = Trim(koushinrs!n2)
            moto_shurui = koushinrs!shurui '契約種類
                    
            If seikyuusumi = "0" Then '請求していない場合
                    '次回請求日が同じなら請求する
                        If CLng(Format(Date, "yyyy")) = t_n2 Then
                            If t_m2 <> 1 Then
                                If CInt(Format(Date, "mm")) = t_m2 - 1 Then
                                    nhk_henkou = 1
                                    new_seikyuukaishi = "1"
                                    new_nhk_kin = seikyuukingaku_nhk
                                End If
                            End If
                        Else
                            If CLng(Format(Date, "yyyy")) = t_n2 - 1 Then
                                If t_m2 = 1 Then
                                    If CInt(Format(Date, "mm")) = 12 Then
                                        nhk_henkou = 1
                                        new_seikyuukaishi = "1"
                                        new_nhk_kin = seikyuukingaku_nhk
                                    End If
                                End If
                            End If
                        End If
            Else '請求している場合
                nhk_hikiotoshi = 1
                '契約タイプが6月なら次回はなし
                '違う場合はチェック
                If moto_shurui <> "5" And moto_shurui <> "6" And moto_shurui <> "7" Then
                    new_seikyuukaishi = "0"
                    Select Case moto_shurui
                    Case "0"
                        kasantsuki = 1
                        new_seikyuukaishi = "1"
                        new_nhk_kin = 11890   '11890
                    Case "1"
                        kasantsuki = 2
                    Case "2"
                        kasantsuki = 3
                    Case "3"
                        kasantsuki = 4
                    Case "4"
                        kasantsuki = 5

                    End Select
                    t_new_m = t_m2 + kasantsuki
                    If t_new_m > 12 Then
                        t_new_n = t_n2 + 1
                        t_new_m = t_new_m - 12
                    Else
                        t_new_n = t_n2
                    End If
                    newjikai = Format(t_new_n, "0000") & Format(t_new_m, "00") & "00"
                    new_nhk_shurui = "5"
                    nhk_henkou = 2
                Else
                    t_new_m = t_m2 + 6
                    If t_new_m > 12 Then
                        t_new_n = t_n2 + 1
                        t_new_m = t_new_m - 12
                    Else
                        t_new_n = t_n2
                    End If
                    newjikai = Format(t_new_n, "0000") & Format(t_new_m, "00") & "00"
                    nhk_henkou = 3
                    new_seikyuukaishi = "0"
                End If
                
                
            End If
                    koushinrs.Close
                    If nhk_henkou = 1 Then
                        sql = "UPDATE nhk_m SET " & _
                                "n2 = '" & new_seikyuukaishi & "'" & _
                                " WHERE nhkid = '" & nhk_id & "'"
                        db4.Execute sql, dbSQLPassThrough
                    
                    ElseIf nhk_henkou = 2 Then
                        sql = "UPDATE nhk_m SET " & _
                                "n2 = '" & new_seikyuukaishi & "'," & _
                                "shurui = '" & new_nhk_shurui & "'," & _
                                "kingaku = '" & new_nhk_kin & "'," & _
                                "zenkai = '" & moto_jikaibi & "'," & _
                                "jikai = '" & newjikai & "'" & _
                                " WHERE nhkid = '" & nhk_id & "'"
                        db4.Execute sql, dbSQLPassThrough
                    ElseIf nhk_henkou = 3 Then
                        sql = "UPDATE nhk_m SET " & _
                                "n2 = '" & new_seikyuukaishi & "'," & _
                                "zenkai = '" & moto_jikaibi & "'," & _
                                "jikai = '" & newjikai & "'" & _
                                " WHERE nhkid = '" & nhk_id & "'"
                        db4.Execute sql, dbSQLPassThrough
                    
                    End If
        ElseIf nhk_resp = -1 Then
                a_fdd_koushin2 = 9
                Exit Function
        End If
        End If
        db4.Close
        
        
        'その他の契約の無0のレコード「２」に変更
         sql = "SELECT keiyaku2.keiyakuid2,keiyaku2.kingaku" & _
             " FROM kojin LEFT JOIN keiyaku2 ON kojin.kojinid = keiyaku2.kojinid" & _
             " WHERE (((kojin.kojinid)='" & motokojinid & "') AND ((keiyaku2.f2)='0'));"
         If FcSQlGet(koushinrs, sql, PrMsg) = True Then
             koushinrs.MoveFirst
             Do Until koushinrs.EOF
                 keisankin = koushinrs!kingaku
                 moto_sonota_shoukei = moto_sonota_shoukei - CLng(keisankin)  'その他小計
                 moto_sonota_seikyuu = moto_sonota_seikyuu - CLng(keisankin) 'その他請求金
                 'motodata(8) = motodata(8) - CLng(keisankin)   '総合計　　　計算用
                 sql = "update keiyaku2 set f2='2' where keiyakuid2 ='" & _
                     koushinrs!keiyakuid2 & "'"
                 On Error GoTo errko1
                     db.Execute sql, dbSQLPassThrough
                 On Error GoTo 0
                 koushinrs.MoveNext
             Loop
             koushinrs.Close
         End If
         
         
        'その他の契約の〆1のレコードを「0」に変更
        sql = "SELECT keiyaku2.keiyakuid2,keiyaku2.kingaku" & _
            " FROM kojin LEFT JOIN keiyaku2 ON kojin.kojinid = keiyaku2.kojinid" & _
            " WHERE (((kojin.kojinid)='" & motokojinid & "') AND ((keiyaku2.f2)='1'));"
        If FcSQlGet(koushinrs, sql, PrMsg) = True Then
            koushinrs.MoveFirst
            Do Until koushinrs.EOF
                keisankin = koushinrs!kingaku
                moto_sonota_seikyuu = moto_sonota_seikyuu + CLng(keisankin) 'その他請求金
                'motodata(8) = motodata(8) + CLng(keisankin)   '総合計　　　計算用
                sql = "update keiyaku2 set f2='0' where keiyakuid2 ='" & _
                    koushinrs!keiyakuid2 & "'"
              On Error GoTo errko1
                db.Execute sql, dbSQLPassThrough
              On Error GoTo 0
                koushinrs.MoveNext
            Loop
            koushinrs.Close
        End If
        
        '主契約の〆のレコードを「０」に変更
        sql = "SELECT keiyaku.keiyakuid,keiyaku.kingaku" & _
            " FROM kojin LEFT JOIN keiyaku ON kojin.kojinid = keiyaku.kojinid" & _
            " WHERE (((kojin.kojinid)='" & motokojinid & "') AND ((keiyaku.shime)='1'));"
        If FcSQlGet(koushinrs, sql, PrMsg) = True Then
            koushinrs.MoveFirst
            Do Until koushinrs.EOF
                keisankin = koushinrs!kingaku
                moto_shu_seikyuu = moto_shu_seikyuu + CLng(keisankin) '主契約請求金
                'motodata(8) = motodata(8) + CLng(keisankin)   '総合計　　　計算用
                sql = "update keiyaku set shime='0' where keiyakuid ='" & _
                    koushinrs!keiyakuid & "'"
               On Error GoTo errko1
                db.Execute sql, dbSQLPassThrough
               On Error GoTo 0
                koushinrs.MoveNext
            Loop
            koushinrs.Close
        End If
        
        henkoushita = 0
        '主契約の９６－９９のレコードを「０１」に変更
        sql = "SELECT keiyaku.keiyakuid,keiyaku.kingaku" & _
            " FROM kojin LEFT JOIN keiyaku ON kojin.kojinid = keiyaku.kojinid" & _
            " WHERE kojin.kojinid='" & motokojinid & "' AND keiyaku.kadou='0'" & _
            " and (keiyaku.shurui='96' or keiyaku.shurui='97' or keiyaku.shurui='98' or keiyaku.shurui='99')"
        If FcSQlGet(koushinrs, sql, PrMsg) = True Then
            koushinrs.MoveFirst
            Do Until koushinrs.EOF
                keisankin = koushinrs!kingaku
                moto_shu_shoukei = moto_shu_shoukei - CLng(keisankin) + 6000 '主契約小計
                moto_shu_seikyuu = moto_shu_seikyuu - CLng(keisankin) '主契約請求金
                sql = "update keiyaku set shurui='01',kingaku='6000',sadou = '0' where keiyakuid ='" & _
                    koushinrs!keiyakuid & "'"
               On Error GoTo errko1
                db.Execute sql, dbSQLPassThrough
               On Error GoTo 0
               
               hikukazu = Int(keisankin / 1000)
               
                n11 = CLng(Format(Date, "yyyy"))
                n12 = CInt(Format(Date, "mm")) + hikukazu
                If n12 > 12 Then
                    n12 = n12 - 12
                    n11 = n11 + 1
                End If
                tsuginohi = Format(n11, "0000") & Format(n12, "00")
               
               
               henkoushita = 1
               
                sql = "update kojin set jikaishu='" & CStr(moto_shu_shoukei) & _
                    "',kajou = '" & CStr(tsuginohi) & "' where kojinid ='" & _
                    motokojinid & "'"
               On Error GoTo errko1
                db.Execute sql, dbSQLPassThrough
               On Error GoTo 0
                koushinrs.MoveNext
            Loop
            koushinrs.Close
        End If
        doudoudou = 0
        bimyou = 0
        If oyasumi = 2 Then
            
            sql = "SELECT kojin.f4,kojin.kojinid,kojin.name, kojin.minyuukin, kojin.jikaishu, kojin.jikaikin, kojin.soukei,keiyaku.keiyakuid,keiyaku.kuchisuu,keiyaku.kingaku" & _
                        " FROM keiyaku RIGHT JOIN kojin ON keiyaku.kojinid = kojin.kojinid" & _
                        " WHERE keiyaku.kadou ='0' AND keiyaku.shurui ='01' and kojin.kojinid ='" & motokojinid & "'"
            If FcSQlGet(koushinrs, sql, PrMsg) = True Then
            '主契約6ヶ月が存在する場合
                    henkanmotoid = koushinrs!keiyakuid
                    henkoukuchisuu = CInt(koushinrs!kuchisuu)
                    sql = "SELECT sum(kingaku) as kikikinkin" & _
                                " From keiyaku" & _
                                " WHERE (((keiyaku.kojinid)='" & motokojinid & "') AND ((keiyaku.shurui)<>'01') AND ((keiyaku.kadou)='0'));"
                    If FcSQlGet(koushinrs2, sql, PrMsg) = True Then
                    'それ以外の稼動する契約が存在する場合
                        If IsNull(koushinrs2!kikikinkin) Or Trim(koushinrs2!kikikinkin) = "" Or Trim(koushinrs2!kikikinkin) = "0" Then
                        Else
                            bimyou = 1
                            'その他の主契約の金額の合計がある場合
                           If moto_kuri = hikakunohi Then
                           '更新するべき月なら
                                    doudoudou = 1
                                     koushinrs2.Close
                                    henkousakikinngaku = henkoukuchisuu * 1000
                                    moto_shu_shoukei = moto_shu_shoukei - (6000 - 1000) * henkoukuchisuu '主契約小計
                                        moto_shu_seikyuu = moto_shu_seikyuu + 1000 * henkoukuchisuu '主契約請求金
                                    '6ヶ月を1ヶ月にへんこう
                                    sql = "UPDATE keiyaku SET shurui = '06',kingaku='" & henkousakikinngaku & "',sadou = '0'" & _
                                      " WHERE keiyakuid = '" & henkanmotoid & "'"
                                    On Error GoTo errko1
                                    db.Execute sql, dbSQLPassThrough
                                    On Error GoTo 0
                                     sql = "update kojin set jikaishu='" & CStr(moto_shu_shoukei) & "' where kojinid ='" & _
                                         motokojinid & "'"
                                    On Error GoTo errko1
                                     db.Execute sql, dbSQLPassThrough
                                    On Error GoTo 0
                            End If
                         End If
                    End If
                     koushinrs.Close
            End If
        End If
        
        keisan_soukei = moto_shu_seikyuu + moto_sonota_seikyuu + new_nhk_kin

If fkekka = 0 Then    '全額引き落としができた場合

s_tainou_jikkou = 0


    If oyasumi = 1 Then '6カ月おき
        sql = "update kojin set  maebi ='" & koushinbi & "'," & _
            "minyuukin = '0'," & _
            "maekin = '" & CStr(kinga) & "'," & _
            "f4 = '0'," & _
            "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
            "jikaikin = '0'," & _
            "fkaisuu = '0'," & _
                    "kekkacode = '" & CStr(fkekka) & "'," & _
            "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
            "kajou = '" & CStr(tsuginohi) & "'," & _
            "soukei= '0'" & _
            " where kojinid ='" & motokojinid & "'"
            
            newsadou = "1"
            
    ElseIf oyasumi = 2 Then  '6カ月おきの2回目以降
        If doudoudou = 1 Then
            If moto_kuri <> hikakunohi Then
                sql = "update kojin set  maebi ='" & koushinbi & "'," & _
                    "minyuukin = '0'," & _
                    "maekin = '" & CStr(kinga) & "'," & _
                    "f4 = '" & CStr(new_nhk_kin) & "'," & _
                    "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                    "kekkacode = '" & CStr(fkekka) & "'," & _
                    "fkaisuu = '0'," & _
                    "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                    "soukei= '" & CStr(keisan_soukei) & "'" & _
                    " where kojinid ='" & motokojinid & "'"
                    
                    newsadou = "1"
            Else
                keisan_soukei = moto_shu_seikyuu + new_nhk_kin
                sql = "update kojin set  maebi ='" & koushinbi & "'," & _
                    "minyuukin = '0'," & _
                    "maekin = '" & CStr(kinga) & "'," & _
                    "f4 = '" & CStr(new_nhk_kin) & "'," & _
                    "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                            "kekkacode = '" & CStr(fkekka) & "'," & _
                    "jikaikin = '" & moto_shu_seikyuu & "'," & _
                    "fkaisuu = '0'," & _
                    "kajou = '1'," & _
                    "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                    "soukei= '" & CStr(keisan_soukei) & "'" & _
                    " where kojinid ='" & motokojinid & "'"
                    
                    newsadou = "0"
            End If
        Else
            If moto_kuri <> hikakunohi Then
            'ここだぞ
                If bimyou = 1 Then
                        sql = "update kojin set  maebi ='" & koushinbi & "'," & _
                            "minyuukin = '0'," & _
                            "maekin = '" & CStr(kinga) & "'," & _
                            "f4 = '" & CStr(new_nhk_kin) & "'," & _
                            "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                                    "kekkacode = '" & CStr(fkekka) & "'," & _
                            "jikaikin = '" & CStr(moto_shu_seikyuu) & "'," & _
                            "fkaisuu = '0'," & _
                            "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                            "soukei= '" & CStr(moto_shu_seikyuu) & "'" & _
                            " where kojinid ='" & motokojinid & "'"
                    
                Else
                        sql = "update kojin set  maebi ='" & koushinbi & "'," & _
                            "minyuukin = '0'," & _
                            "maekin = '" & CStr(kinga) & "'," & _
                            "f4 = '" & CStr(new_nhk_kin) & "'," & _
                            "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                                    "kekkacode = '" & CStr(fkekka) & "'," & _
                            "jikaikin = '0'," & _
                            "fkaisuu = '0'," & _
                            "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                            "soukei= '0'" & _
                            " where kojinid ='" & motokojinid & "'"
               End If
                    newsadou = "1"
            Else
                keisan_soukei = keisan_soukei + moto_shu_shoukei
                If doudoudou = 1 Then
                        sql = "update kojin set  maebi ='" & koushinbi & "'," & _
                            "minyuukin = '0'," & _
                            "maekin = '" & CStr(kinga) & "'," & _
                            "f4 = '" & CStr(new_nhk_kin) & "'," & _
                            "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                                    "kekkacode = '" & CStr(fkekka) & "'," & _
                            "jikaikin = '" & moto_shu_shoukei & "'," & _
                            "fkaisuu = '0'," & _
                            "kajou = '0'," & _
                            "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                            "soukei= '" & CStr(keisan_soukei) & "'" & _
                            " where kojinid ='" & motokojinid & "'"
                Else
                        sql = "update kojin set  maebi ='" & koushinbi & "'," & _
                            "minyuukin = '0'," & _
                            "maekin = '" & CStr(kinga) & "'," & _
                            "f4 = '" & CStr(new_nhk_kin) & "'," & _
                            "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                                    "kekkacode = '" & CStr(fkekka) & "'," & _
                            "jikaikin = '" & moto_shu_shoukei & "'," & _
                            "fkaisuu = '0'," & _
                            "kajou = '1'," & _
                            "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                            "soukei= '" & CStr(keisan_soukei) & "'" & _
                            " where kojinid ='" & motokojinid & "'"
                End If
                    newsadou = "0"
            End If
        End If
    Else
        If henkoushita = 1 Then
            sql = "update kojin set  maebi ='" & koushinbi & "'," & _
                "minyuukin = '0'," & _
                "maekin = '" & CStr(kinga) & "'," & _
                "f4 = '" & CStr(new_nhk_kin) & "'," & _
                "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                "jikaikin = '" & CStr(moto_shu_seikyuu) & "'," & _
                        "kekkacode = '" & CStr(fkekka) & "'," & _
                "fkaisuu = '0'," & _
                "kajou = '" & CStr(tsuginohi) & "'," & _
                "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                "soukei= '" & CStr(keisan_soukei) & "'" & _
                " where kojinid ='" & motokojinid & "'"
        Else
            sql = "update kojin set  maebi ='" & koushinbi & "'," & _
                "minyuukin = '0'," & _
                "maekin = '" & CStr(kinga) & "'," & _
                "f4 = '" & CStr(new_nhk_kin) & "'," & _
                "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                "jikaikin = '" & CStr(moto_shu_seikyuu) & "'," & _
                        "kekkacode = '" & CStr(fkekka) & "'," & _
                "fkaisuu = '0'," & _
                "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                "soukei= '" & CStr(keisan_soukei) & "'" & _
                " where kojinid ='" & motokojinid & "'"
        End If
        newsadou = "0"
    End If
        On Error GoTo errko1
            db.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        '更新記録を残す
            If soukatsu(0, "renban", 8, "", 1, 1) = False Then
                    MsgBox "設定の読み込みに失敗しました。"
                        a_fdd_koushin2 = 9
                        Exit Function
            End If
            If reg = "0" Then
                reg = ""
            End If
            If reg = vbNullString Then
                koushnid = "00000001"
            ElseIf reg = "9999999" Then
                MsgBox "これ以上の更新記録の作成はできません。管理者に連絡してください。"
                a_fdd_koushin2 = 9
                Exit Function
            Else
                koushnid = Format$(reg, "000000#")
            End If
            
            Dim koushinkingaku As Long, koushinhouhou As String, newkoushinid As String

            koushinkingaku = kinga
            
            koushinhouhou = CStr(houhou)
                
            'koushin.mdbをOPEN
            If kidou2 = False Then
                MsgBox "データベースファイルを開けませんでした。"
                a_fdd_koushin2 = 9
                Exit Function
            End If
                
            sql = "INSERT INTO koushinkiroku VALUES('" & koushnid & "'," & _
                           "'" & koushinbi & "','" & motokojinid & "'," & _
                           "'" & koushinkingaku & "','" & koushinhouhou & "','" & Space(1) & "')"
            On Error GoTo errko2
                db2.Execute sql, dbSQLPassThrough
            On Error GoTo 0
            newkoushinid = CStr(CLng(koushnid) + 1)
            If soukatsu(1, "renban", 8, newkoushinid, 1, 0) = False Then
                    MsgBox "設定の書き込みに失敗しました。"
                    a_fdd_koushin2 = 9
                    Exit Function
            End If
            '詳細更新
            'ここからここから
            sql = "UPDATE shousai SET  keinyu = '" & koushinbi & "',keist='1'" & _
                        " WHERE kojinid = '" & motokojinid & "' and keist ='0'"

            On Error GoTo errsoukatsunn
                 db2.Execute sql, dbSQLPassThrough
            On Error GoTo 0

               a_fdd_koushin2 = 0
              If nhk_hikiotoshi = 1 Then '請求済み
                nhk_hikiotoshi = 2 '引き落としOK
              End If
Else 'If fkekka = 1 Then  '引き落としできない場合
    moto_f_kasuu = moto_f_kasuu + 1
                '更新処理
                
        '滞納ST
        s_tainou_jikkou = 1
                
                
        If oyasumi = 1 Then '6カ月おき
            keisan_soukei = kinga
                sql = "update kojin set " & _
                    "minyuukin = '" & CStr(kinga) & "'," & _
                    "f4 = '" & CStr(new_nhk_kin) & "'," & _
                    "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                    "fkaisuu = '" & CStr(moto_f_kasuu) & "'," & _
                    "jikaikin = '0'," & _
                    "kajou = '" & CStr(tsuginohi) & "'," & _
                    "kekkacode = '" & CStr(fkekka) & "'," & _
                    "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                    "soukei= '" & CStr(keisan_soukei) & "'" & _
                    " where kojinid ='" & motokojinid & "'"
                    
                    newsadou = "1"
        Else
            If moto_kuri <> hikakunohi Then

                keisan_soukei = keisan_soukei + kinga
                '個人データの変更
                sql = "update kojin set " & _
                    "minyuukin = '" & CStr(kinga) & "'," & _
                    "f4 = '" & CStr(new_nhk_kin) & "'," & _
                    "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                    "kekkacode = '" & CStr(fkekka) & "'," & _
                    "fkaisuu = '" & CStr(moto_f_kasuu) & "'," & _
                    "jikaikin = '" & CStr(moto_shu_seikyuu) & "'," & _
                    "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                    "soukei= '" & CStr(keisan_soukei) & "'" & _
                    " where kojinid ='" & motokojinid & "'"
                   If oyasumi = 0 Then
                        newsadou = "0"
                    Else
                        newsadou = "1"
                    End If
            Else
                keisan_soukei = keisan_soukei + kinga + moto_shu_shoukei
                '個人データの変更
                sql = "update kojin set " & _
                    "minyuukin = '" & CStr(kinga) & "'," & _
                    "f4 = '" & CStr(new_nhk_kin) & "'," & _
                    "jikaita = '" & CStr(moto_sonota_shoukei) & "'," & _
                    "kekkacode = '" & CStr(fkekka) & "'," & _
                    "fkaisuu = '" & CStr(moto_f_kasuu) & "'," & _
                    "jikaikin = '" & CStr(moto_shu_shoukei) & "'," & _
                    "sonotakin = '" & CStr(moto_sonota_seikyuu) & "'," & _
                    "kajou = '0'," & _
                    "soukei= '" & CStr(keisan_soukei) & "'" & _
                    " where kojinid ='" & motokojinid & "'"
                    
                    newsadou = "0"
            End If
        End If
                On Error GoTo errko1
                db.Execute sql, dbSQLPassThrough
                On Error GoTo 0
                                
              If nhk_hikiotoshi = 1 Then '請求済み
                nhk_hikiotoshi = 3 '引き落としNG
              End If
               
               a_fdd_koushin2 = 0
End If


 '滞納ST
     If s_tainou_jikkou = 1 Then
     '落ちなかった場合
            If s_tainou_tsuika_kingaku <> 0 Then
            '項目名の取得
            Dim s_koumokumei As String
            s_koumokumei = sonota_hyoujimei_get("96")
            '金額と導入設定OK
                If tainou_tesuuryou_touroku(motokojinid, s_tainou_tsuika_kingaku, s_koumokumei) = -1 Then
                    MsgBox "延滞手数料の登録に失敗しました。" & motokojinid
                End If
            End If
    End If
    
    
    
    
        '作動伝票の更新
        Dim sql_sadou As String, rs_sadou As Recordset, sql_sadou2 As String
            sql_sadou = "SELECT keiyaku.keiyakuid" & _
                        " FROM keiyaku RIGHT JOIN kojin ON keiyaku.kojinid = kojin.kojinid" & _
                        " WHERE keiyaku.kadou ='0' AND keiyaku.shurui ='01' and kojin.kojinid ='" & motokojinid & "'"
            If FcSQlGet(rs_sadou, sql_sadou, PrMsg) = True Then
                sql_sadou2 = "UPDATE keiyaku SET sadou = '" & newsadou & "'" & _
                              " WHERE keiyakuid = '" & rs_sadou!keiyakuid & "'"
                On Error GoTo errko1
                db.Execute sql_sadou2, dbSQLPassThrough
                On Error GoTo 0
                rs_sadou.Close
            End If


'フラッグを元に戻す
If flgup(0, motokojinid) = False Then
    MsgBox "フラッグの更新に失敗しました。"
    a_fdd_koushin2 = 9
    Exit Function
End If

'NHK＿Kのデータ更新
If nhk_hikiotoshi <> 0 Then
    If kidou4 = False Then
        a_fdd_koushin2 = 9
        Exit Function
    End If
    sql = "select * from nhk_k where kojinid ='" & motokojinid & "' and seikyuubi ='" & moto_jikaibi & "'"
    nhk_resp = FcSQlGet4(koushinrs, sql, PrMsg)
    If nhk_resp > 0 Then
        If nhk_hikiotoshi <> 1 Then '引落OK
            If fkekka = 0 Then
                sql = "update nhk_k set " & _
                    "suteitasu= '2'" & _
                    " where count ='" & koushinrs!Count & "'"
            Else
                sql = "update nhk_k set " & _
                    "suteitasu= '3'" & _
                    " where count ='" & koushinrs!Count & "'"
            End If
             On Error GoTo errnhk1
             db4.Execute sql, dbSQLPassThrough
             On Error GoTo 0
            
        Else
           MsgBox "不正なＮＨＫフラッグが立ちました。"
           a_fdd_koushin2 = 9
           db4.Close
           Exit Function
        End If
    Else
        MsgBox "ＮＨＫデータの更新に失敗しました。"
        a_fdd_koushin2 = 9
       db4.Close
        Exit Function
    End If
    
    db4.Close
    
    
  
    
    
    
End If
Exit Function
errko1:
    MsgBox "設定の書き込みに失敗しました。"
    a_fdd_koushin2 = 9
    Exit Function
errko2:
    MsgBox "設定の書き込みに失敗しました。"
    db2.Close
    a_fdd_koushin2 = 9
    Exit Function
errsoukatsunn:
                    MsgBox "詳細の更新に失敗しました。"
                    a_fdd_koushin2 = 9
                    Exit Function

errnhk1:
    MsgBox "ＮＨＫ更新データの書き込みに失敗しました。"
    db4.Close
    a_fdd_koushin2 = 9
    Exit Function




End Function

Function tainou_tesuuryou_touroku(s_kojinid As String, s_kingaku As Long, s_koumokumei As String) As Integer

tainou_tesuuryou_touroku = 0

Dim tid As String, tkojinid As String, tkingaku As String, jurs As Recordset
Dim tkaishi As String, tkudou As String, tshurui As String, ttid As String
Dim s_name As String


            If soukatsu(0, "renban", 6, "", 1, 1) = False Then
                    MsgBox "設定の読み込みに失敗しました。"
                    tainou_tesuuryou_touroku = -1
                    Exit Function
            End If
            If reg = "0" Then
                reg = ""
            End If
            If reg = vbNullString Then
                tid = "0000001"
            Else
                tid = Format$(reg, "000000#")
            End If
       ' With frmkeiyaku5
            tshurui = "96" 'Mid(.cmbsonotakeiyaku.Text, 1, 2)
            s_name = s_koumokumei  '"延滞手数料" ' .cmbsonotakeiyaku.Text
            tkojinid = s_kojinid  'frmtouroku2.lblid.Caption
            tkingaku = CStr(s_kingaku)  '.txtkeiyakukin.Text
            tkaishi = Format(Date, "yyyymmdd")  'Format(.txtkaishibi.Text, "yyyymmdd")
            tkudou = "0" '"可能"
            
      ' End With
        sql = "INSERT INTO keiyaku2 VALUES('" & tid & "','" & tkojinid & "'," & _
                       "'" & tshurui & "','" & tkaishi & "','" & tkudou & "','" & tkingaku & "')"
On Error GoTo errk5
        db.Execute sql, dbSQLPassThrough
On Error GoTo 0

        ttid = CStr(CLng(tid) + 1)
        If soukatsu(1, "renban", 6, ttid, 1, 1) = False Then
                MsgBox "設定の書き込みに失敗しました。"
                tainou_tesuuryou_touroku = -1
                Exit Function
        End If
        
         '最終更新日更新
            
        
        
        Dim s_kyounohi As String
        
        s_kyounohi = Format(Date, "yyyy/mm/dd")

         sql = "update kojin set  saishuukoushinbi='" & s_kyounohi & "'" & _
                   " where kojinid ='" & s_kojinid & "'"
        
On Error GoTo errsakujo2
        db.Execute sql, dbSQLPassThrough
On Error GoTo 0
         
        
        If a_sonota_touroku(tkojinid, tkingaku, tkudou) = False Then
            MsgBox "料金の更新に失敗しました。" & tkojinid & "の計算を確認してください。"
            tainou_tesuuryou_touroku = -1
        End If
        
    'Log
    Dim ss_name As String
    ss_name = frmtouroku2.lblname.Caption
    Dim ss_id As String
        ss_id = tshurui  'frmtouroku2.lblid.Caption
    sagyou_msg = "その他の契約：契約 ID " & s_name & "の契約を新規登録しました。"
    log_sagyou sagyou_log_path, sagyou_msg, ss_id, ss_name
        

 
Exit Function

errk5:
    MsgBox "契約の登録に失敗しました。"
    tainou_tesuuryou_touroku = -1
    Exit Function
errsakujo2:
    MsgBox "最終更新日の登録に失敗しました。"
    tainou_tesuuryou_touroku = -1
    Exit Function


End Function



Sub data_tsuia()


Dim newkojinid As String, newseikyuubi As String, newshurui As String, newshurui_no As String, newkin As String, newbankid As String


 With frmsaitekika
 
    newkojinid = Trim(.txtkojinid.Text) '"03113"
    
    newseikyuubi = Trim(.txtseikyuunengetsu.Text) ' "202105" ' Mid(newnyukinbi, 1, 6)
    
    newshurui = Trim(.txtshuruiid.Text)  '"1"
    
    newshurui_no = Trim(.txtkeiyakuid.Text) '"35"  '35=番号表示　　39=着信拒否
    
    newkin = Trim(.txtkingaku.Text)  '"440"      '35＝440　　39＝550
    
    If Trim(.txtbankid.Text) = "" Then
        newbankid = Space(1)
    Else
        newbankid = Trim(.txtbankid.Text)
    End If
    
End With

 If shousai_main_set(newkojinid, newseikyuubi, newshurui, newshurui_no, newkin, newbankid, "1") = -1 Then
            MsgBox "契約の追加に失敗しました。"
            Exit Sub
  End If
  
  
  MsgBox "契約の追加しました。"
  

End Sub

Function shousai_main_set(newkojinid As String, shousai_seikyuubi As String, shurui As String, shousai_id As String, shousai_kin As String, newnewbankid As String, newkuchisuu As String) As Integer
    
    Dim newshousaiid As String, newnewshousaiid As String
    
    If kidou2 = False Then
         MsgBox "データベースファイルを開けませんでした。"
        shousai_main_set = -1
        Screen.MousePointer = 0
         Exit Function
     End If
     
        If soukatsu(0, "renban", 15, "", 0, 0) = False Then
                MsgBox "設定の読み込みに失敗しました。"
                shousai_main_set = -1
                Screen.MousePointer = 0
                Exit Function
        End If
        If reg = "0" Then
            reg = ""
        End If
        If reg = vbNullString Then
            newshousaiid = "000000001"
        Else
            newshousaiid = Format$(reg, "00000000#")
        End If
        If Trim(newnewbankid) = "" Then
            newnewbankid = Space(1)
        End If
        
        sql = "INSERT INTO shousai VALUES('" & newshousaiid & "'," & _
                        "'" & shousai_seikyuubi & "','" & shurui & "'," & _
                        "'" & shousai_id & "','" & shousai_kin & "','0','" & Space(1) & "','" & newkojinid & "','" & newnewbankid & "','" & newkuchisuu & "')"
            
        On Error GoTo errbankkiroku
             db2.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        
        shitei_shousaiid = newshousaiid
        
        
        newnewshousaiid = CStr(CLng(newshousaiid) + 1)
        If soukatsu(1, "renban", 15, newnewshousaiid, 0, 0) = False Then
                MsgBox "設定の書き込みに失敗しました。"
                shousai_main_set = -1
                Screen.MousePointer = 0
                Exit Function
        End If
    
    
    
Exit Function

errbankkiroku:
         MsgBox "データの書き込みに失敗しました。"
        shousai_main_set = -1
        shitei_shousaiid = ""
        Screen.MousePointer = 0
         Exit Function

    
End Function
Function shousaiset(idid As String, hiduhidu As String, kinkin As Long, Optional subhiduke As String, Optional newbankid_sub As String, Optional genkin As String) As Integer

Dim sql_shousai As String, rs_shousai As Recordset, kakikomi_data(), newbankid As String
Dim sougou_kingaku As Long, ysuu As Long, ycounter As Long, newjikaibi As Integer, hikakunen As String


'初期化
sougou_kingaku = 0
ysuu = 0
newjikaibi = 0
newbankid = ""
hikakunen = Format(Date, "YYYYMM")
If idid = "00627" Then
    'MsgBox "hit"
End If

'未支払
    sql_shousai = "SELECT bank.bankcode, kojin.*" & _
            " FROM kojin LEFT JOIN bank ON kojin.bankid = bank.bankid" & _
        " WHERE kojinid ='" & idid & "'"
If FcSQlGet(rs_shousai, sql_shousai, PrMsg) = True Then
    sougou_kingaku = sougou_kingaku + CLng(rs_shousai!minyuukin)
    If IsNull(rs_shousai!bankcode) Then
        newbankid = Space(1)
    Else
        newbankid = rs_shousai!bankcode
    End If
    If rs_shousai!kajou <> "" And rs_shousai!kajou <> "0" And rs_shousai!kajou <> "1" Then
        newjikaibi = 1
    End If
    rs_shousai.Close
    'If Trim(newbankid_sub) <> Trim(newbankid) Then
        If kidou2 = False Then
             MsgBox "データベースファイルを開けませんでした。"
            shousaiset = -1
            Screen.MousePointer = 0
             Exit Function
         End If
         'shousaiの更新
        sql = "UPDATE shousai SET bankid='" & newbankid_sub & "'" & _
                    " WHERE kojinid = '" & idid & "' and keist='0'"
    
        On Error GoTo errsoukatsunn99
             db2.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        newbankid = newbankid_sub
    'End If
End If



'主契約
If newjikaibi = 1 Then
    sql_shousai = "SELECT * FROM keiyaku" & _
        " WHERE kojinid ='" & idid & "' and kadou='0' and shurui<>'01'"
Else
    sql_shousai = "SELECT * FROM keiyaku" & _
        " WHERE kojinid ='" & idid & "' and kadou='0' and sadou='0'"
End If
ycounter = 0
If FcSQlGet(rs_shousai, sql_shousai, PrMsg) = True Then
    ysuu = rs_shousai.RecordCount
    ReDim Preserve kakikomi_data(5, ysuu)
    rs_shousai.MoveFirst
    Do Until rs_shousai.EOF
        kakikomi_data(0, ycounter) = Mid(hiduhidu, 1, 6)
        kakikomi_data(1, ycounter) = "1"
        'kakikomi_data(2, ycounter) = rs_shousai!keiyakuid
        kakikomi_data(2, ycounter) = rs_shousai!shurui
        kakikomi_data(3, ycounter) = rs_shousai!kingaku
        kakikomi_data(4, ycounter) = rs_shousai!kuchisuu
        ycounter = ycounter + 1
        sougou_kingaku = sougou_kingaku + CLng(rs_shousai!kingaku)
        rs_shousai.MoveNext
    Loop
    rs_shousai.Close
End If

Dim sql_sude As String, rs_sude As Recordset, karikingaku As Long

'その他の契約
    sql_shousai = "SELECT * FROM keiyaku2" & _
        " WHERE kojinid ='" & idid & "' and f2='0'"
     
karikingaku = 0
If FcSQlGet(rs_shousai, sql_shousai, PrMsg) = True Then
    ysuu = ysuu + rs_shousai.RecordCount
    ReDim Preserve kakikomi_data(5, ysuu)
    rs_shousai.MoveFirst
    Do Until rs_shousai.EOF
        If genkin = "1" Then
            kakikomi_data(0, ycounter) = Mid(hiduhidu, 1, 6)
            kakikomi_data(1, ycounter) = "2"
            'kakikomi_data(2, ycounter) = rs_shousai!keiyakuid2
            kakikomi_data(2, ycounter) = rs_shousai!sonotaid
            kakikomi_data(3, ycounter) = rs_shousai!kingaku
             kakikomi_data(4, ycounter) = "1"
            ycounter = ycounter + 1
            sougou_kingaku = sougou_kingaku + CLng(rs_shousai!kingaku)
            rs_shousai.MoveNext
        Else
            sql_sude = "select * from shousai where keishurui='2' and kojinid ='" & idid & "' and keist ='1'" & _
                            " and keikin ='" & rs_shousai!kingaku & "' and keiid ='" & rs_shousai!sonotaid & "'" & _
                            " and bankid='" & Space(1) & "'and keiseikyuu like '" & hikakunen & "*'"
            If FcSQlGet2(rs_sude, sql_sude, PrMsg) = True Then
                rs_shousai.MoveNext
                rs_sude.Close
            Else
                kakikomi_data(0, ycounter) = Mid(hiduhidu, 1, 6)
                kakikomi_data(1, ycounter) = "2"
                'kakikomi_data(2, ycounter) = rs_shousai!keiyakuid2
                kakikomi_data(2, ycounter) = rs_shousai!sonotaid
                kakikomi_data(3, ycounter) = rs_shousai!kingaku
                 kakikomi_data(4, ycounter) = "1"
                ycounter = ycounter + 1
                sougou_kingaku = sougou_kingaku + CLng(rs_shousai!kingaku)
                rs_shousai.MoveNext
            End If
        End If
    Loop
    rs_shousai.Close
End If

sougou_kingaku = sougou_kingaku + karikingaku
'NHKの契約
    If kidou4 = False Then
        shousaiset = -1
        Exit Function
    End If
    If subhiduke = "" Then
            sql_shousai = "SELECT * FROM nhk_m" & _
                " WHERE kojinid ='" & idid & "' and jikai='" & Mid(hiduhidu, 1, 6) & "00" & "'"
    Else
            sql_shousai = "SELECT * FROM nhk_m" & _
                " WHERE kojinid ='" & idid & "' and zenkai='" & subhiduke & "00" & "'"
    End If
    If FcSQlGet4(rs_shousai, sql_shousai, PrMsg) > 0 Then
        If Trim(rs_shousai!n2) = 1 Then
            ysuu = ysuu + 1
            ReDim Preserve kakikomi_data(5, ysuu)
            kakikomi_data(0, ycounter) = Mid(hiduhidu, 1, 6)
            kakikomi_data(1, ycounter) = "3"
            'kakikomi_data(2, ycounter) = rs_shousai!nhkid
            kakikomi_data(2, ycounter) = rs_shousai!shurui
            kakikomi_data(3, ycounter) = rs_shousai!kingaku
            kakikomi_data(4, ycounter) = rs_shousai!n1
            sougou_kingaku = sougou_kingaku + CLng(rs_shousai!kingaku)
        End If
        rs_shousai.Close
    End If
    db4.Close

'一応金額のチェック
If sougou_kingaku <> kinkin Then
    MsgBox "請求用のデータが構築されていません。データを確認してください。「口座扱いの入金処理」や「現金扱いのＮＨＫ契約あり」場合などで、請求詳細がない場合は、手動で作成してください。" & idid
 
    
    
    ' For ycounter = 0 To ysuu - 1
        ' If shousai_main_set(idid, CStr(kakikomi_data(0, ycounter)), CStr(kakikomi_data(1, ycounter)), CStr(kakikomi_data(2, ycounter)), CStr(kakikomi_data(3, ycounter)), newbankid, CStr(kakikomi_data(4, ycounter))) = -1 Then
            shousaiset = -1
            Exit Function
        ' End If
    ' Next
Else
    For ycounter = 0 To ysuu - 1
        If shousai_main_set(idid, CStr(kakikomi_data(0, ycounter)), CStr(kakikomi_data(1, ycounter)), CStr(kakikomi_data(2, ycounter)), CStr(kakikomi_data(3, ycounter)), newbankid, CStr(kakikomi_data(4, ycounter))) = -1 Then
            shousaiset = -1
            Exit Function
        End If
    Next
End If


shousaiset = 0

Exit Function

errsoukatsunn99:
    MsgBox "過去の詳細データの書き換えに失敗しました。確認後、再度実行してください。" & idid
    shousaiset = -1
    Exit Function

End Function

Function a_nhk_henkou(koid As String, new_gokin As String, kakin As Boolean) As Boolean
'******************************
'NHK契約登録時の個人データの変更
'******************************
Dim rs_shukeiyakukoushin As Recordset
Dim gen_soukei As String
Dim saki_shukeiyakushoukei As String, saki_nhk As String, saki_soukei As String

    sql = "select * from kojin where kojinid ='" & koid & "'"
    If FcSQlGet(rs_shukeiyakukoushin, sql, PrMsg) = False Then
        MsgBox "更新先登録済個人データがありません。"
        a_nhk_henkou = False
        Exit Function
    Else
        '元のデータ
Dim moto_shu_seikyuu As Long, moto_sonota_seikyuu As Long
Dim moto_nhk As Long, moto_soukei As Long, moto_minyu As Long
         '主契約請求金
            If Trim(rs_shukeiyakukoushin!jikaikin) = "" Then
                moto_shu_seikyuu = 0
            Else
                moto_shu_seikyuu = CLng(rs_shukeiyakukoushin!jikaikin)
            End If
         'その他請求金
            If Trim(rs_shukeiyakukoushin!sonotakin) = "" Then
                moto_sonota_seikyuu = 0
            Else
                moto_sonota_seikyuu = CLng(rs_shukeiyakukoushin!sonotakin)
            End If
         '未支払
            If Trim(rs_shukeiyakukoushin!minyuukin) = "" Then
                moto_minyu = 0
            Else
                moto_minyu = CLng(rs_shukeiyakukoushin!minyuukin)
            End If
        'NHK
            If Trim(rs_shukeiyakukoushin!f4) = "" Then
                moto_nhk = 0
            Else
                moto_nhk = CLng(rs_shukeiyakukoushin!f4)
            End If
         moto_soukei = CLng(rs_shukeiyakukoushin!soukei)
            
         '新しい値
        If kakin = True Then '適用範囲内
            saki_nhk = CLng(new_gokin)
            saki_soukei = moto_sonota_seikyuu + moto_shu_seikyuu + moto_minyu + saki_nhk
        Else
            saki_nhk = 0
            saki_soukei = moto_sonota_seikyuu + moto_shu_seikyuu + moto_minyu
        End If
            
        sql = "UPDATE kojin SET f4 = '" & CStr(saki_nhk) & "'," & _
            "soukei = '" & CStr(saki_soukei) & "'" & _
            " WHERE kojinid = '" & koid & "'"
            
        On Error GoTo errnewkin4
          db.Execute sql, dbSQLPassThrough
         On Error GoTo 0
         
         frmtouroku2.txtnhk.Text = saki_nhk
        frmtouroku2.txtsoukei.Text = saki_soukei
        a_nhk_henkou = True
    End If
Exit Function

errnewkin4:
    MsgBox "個人データの更新に失敗しました。"
    a_nhk_henkou = False
    Exit Function






End Function

Function a_nhk_touroku(koid As String, newnhkkin As String) As Boolean
'******************************
'NHK契約登録時の個人データの更新
'******************************
Dim rs_shukeiyakukoushin As Recordset
Dim gen_nhk_hyouji As String, gen_soukei As String
Dim saki_shukeiyakuseikyuu As String, saki_soukei As String

    sql = "select * from kojin where kojinid ='" & koid & "'"
    If FcSQlGet(rs_shukeiyakukoushin, sql, PrMsg) = False Then
        MsgBox "更新先登録済個人データがありません。"
        a_nhk_touroku = False
        Exit Function
    Else
        '元のデータ
Dim moto_shu_seikyuu As Long, moto_sonota_seikyuu As Long
Dim moto_sannnyuu As Long
         '主契約請求金
            If Trim(rs_shukeiyakukoushin!jikaikin) = "" Then
                moto_shu_seikyuu = 0
            Else
                moto_shu_seikyuu = CLng(rs_shukeiyakukoushin!jikaikin)
            End If
         'その他請求金
            If Trim(rs_shukeiyakukoushin!sonotakin) = "" Then
                moto_sonota_seikyuu = 0
            Else
                moto_sonota_seikyuu = CLng(rs_shukeiyakukoushin!sonotakin)
            End If
        '未入金額
            If Trim(rs_shukeiyakukoushin!minyuukin) = "" Then
                moto_sannnyuu = 0
            Else
                moto_sannnyuu = CLng(rs_shukeiyakukoushin!minyuukin)
            End If
        'NHK
        gen_nhk_hyouji = rs_shukeiyakukoushin!f4
        If Trim(gen_nhk_hyouji) = "" Then
            gen_nhk_hyouji = "0"
        End If
        '総合計
        gen_soukei = rs_shukeiyakukoushin!soukei
        
        '新しいデータ
                saki_shukeiyakuseikyuu = CStr(CLng(newnhkkin))
                saki_soukei = CStr(CLng(gen_soukei) + CLng(newnhkkin))
        
        rs_shukeiyakukoushin.Close
        
        sql = "UPDATE kojin SET f4 = '" & CStr(saki_shukeiyakuseikyuu) & "'," & _
            "soukei = '" & CStr(saki_soukei) & "'" & _
            " WHERE kojinid = '" & koid & "'"
        On Error GoTo errnewkin4
          db.Execute sql, dbSQLPassThrough
         On Error GoTo 0
         
         frmtouroku2.txtnhk.Text = saki_shukeiyakuseikyuu
        frmtouroku2.txtsoukei.Text = saki_soukei
        
        a_nhk_touroku = True
    End If
Exit Function

errnewkin4:
    MsgBox "個人データの更新に失敗しました。"
    a_nhk_touroku = False
    Exit Function


End Function
Function a_sonota_sakujo(koid As String, maekin As String) As Boolean
'******************************
'その他契約変更時の個人データの削除
'******************************
Dim rs_shukeiyakukoushin As Recordset
Dim gen_sonotashoukei As String, gen_soukei As String
Dim saki_sonotashoukei As String, saki_sonotaseikyuu As String, saki_soukei As String

    sql = "select * from kojin where kojinid ='" & koid & "'"
    If FcSQlGet(rs_shukeiyakukoushin, sql, PrMsg) = False Then
        MsgBox "更新先登録済個人データがありません。"
        a_sonota_sakujo = False
        Exit Function
    Else
        '元のデータ
Dim moto_shu_seikyuu As Long, moto_sonota_seikyuu As Long
Dim moto_sannnyuu As Long
         '主契約請求金
            If Trim(rs_shukeiyakukoushin!jikaikin) = "" Then
                moto_shu_seikyuu = 0
            Else
                moto_shu_seikyuu = CLng(rs_shukeiyakukoushin!jikaikin)
            End If
         'その他請求金
            If Trim(rs_shukeiyakukoushin!sonotakin) = "" Then
                moto_sonota_seikyuu = 0
            Else
                moto_sonota_seikyuu = CLng(rs_shukeiyakukoushin!sonotakin)
            End If
        '未入金額
            If Trim(rs_shukeiyakukoushin!minyuukin) = "" Then
                moto_sannnyuu = 0
            Else
                moto_sannnyuu = CLng(rs_shukeiyakukoushin!minyuukin)
            End If
        'その他小計
            gen_sonotashoukei = rs_shukeiyakukoushin!jikaita
            If Trim(gen_sonotashoukei) = "" Then
                gen_sonotashoukei = "0"
            End If
        '総合計
            gen_soukei = rs_shukeiyakukoushin!soukei
        '新しいその他小計
        saki_sonotashoukei = CStr(CLng(gen_sonotashoukei) - CLng(maekin))

        '新しいデータ
        saki_sonotaseikyuu = CStr(CLng(moto_sonota_seikyuu) - CLng(maekin))
        saki_soukei = CStr(CLng(gen_soukei) - CLng(maekin))
        rs_shukeiyakukoushin.Close
        sql = "UPDATE kojin SET sonotakin = '" & CStr(saki_sonotaseikyuu) & "'," & _
            "soukei = '" & CStr(saki_soukei) & "',jikaita = '" & CStr(saki_sonotashoukei) & "'" & _
            " WHERE kojinid = '" & koid & "'"
        On Error GoTo errnewkin4
          db.Execute sql, dbSQLPassThrough
         On Error GoTo 0
         frmtouroku2.txtsonotashoukei.Text = saki_sonotashoukei
         frmtouroku2.txtsonotaseikyuu.Text = saki_sonotaseikyuu
        frmtouroku2.txtsoukei.Text = saki_soukei
        a_sonota_sakujo = True
    End If
Exit Function

errnewkin4:
    MsgBox "個人データの更新に失敗しました。"
    a_sonota_sakujo = False
    Exit Function

End Function

Function a_sonota_henkou(koid As String, maekin As String, gokin As String) As Boolean
'******************************
'その他契約変更時の個人データの変更
'******************************
Dim rs_shukeiyakukoushin As Recordset
Dim gen_sonotashoukei As String, gen_soukei As String
Dim saki_sonotashoukei As String, saki_sonotaseikyuu As String, saki_soukei As String

    sql = "select * from kojin where kojinid ='" & koid & "'"
    If FcSQlGet(rs_shukeiyakukoushin, sql, PrMsg) = False Then
        MsgBox "更新先登録済個人データがありません。"
        a_sonota_henkou = False
        Exit Function
    Else
        '元のデータ
Dim moto_shu_seikyuu As Long, moto_sonota_seikyuu As Long
Dim moto_sannnyuu As Long
         '主契約請求金
            If Trim(rs_shukeiyakukoushin!jikaikin) = "" Then
                moto_shu_seikyuu = 0
            Else
                moto_shu_seikyuu = CLng(rs_shukeiyakukoushin!jikaikin)
            End If
         'その他請求金
            If Trim(rs_shukeiyakukoushin!sonotakin) = "" Then
                moto_sonota_seikyuu = 0
            Else
                moto_sonota_seikyuu = CLng(rs_shukeiyakukoushin!sonotakin)
            End If
        '未入金額
            If Trim(rs_shukeiyakukoushin!minyuukin) = "" Then
                moto_sannnyuu = 0
            Else
                moto_sannnyuu = CLng(rs_shukeiyakukoushin!minyuukin)
            End If
        'その他小計
        gen_sonotashoukei = rs_shukeiyakukoushin!jikaita
        If Trim(gen_sonotashoukei) = "" Then
            gen_sonotashoukei = "0"
        End If
        '総合計
        gen_soukei = rs_shukeiyakukoushin!soukei
        
        '新しいその他小計
        saki_sonotashoukei = CStr(CLng(gen_sonotashoukei) - CLng(maekin) + CLng(gokin))

        '新しいデータ
        If maekin <> gokin Then   '価格が違う場合
                '新しいデータ
                saki_sonotaseikyuu = CStr(CLng(moto_sonota_seikyuu) - CLng(maekin) + CLng(gokin))
                saki_soukei = CStr(CLng(gen_soukei) - CLng(maekin) + CLng(gokin))
        Else '価格が同じ場合
            saki_sonotaseikyuu = moto_sonota_seikyuu
            saki_soukei = gen_soukei
        End If
        rs_shukeiyakukoushin.Close
        
        sql = "UPDATE kojin SET sonotakin = '" & CStr(saki_sonotaseikyuu) & "'," & _
            "soukei = '" & CStr(saki_soukei) & "',jikaita = '" & CStr(saki_sonotashoukei) & "'" & _
            " WHERE kojinid = '" & koid & "'"
        On Error GoTo errnewkin4
          db.Execute sql, dbSQLPassThrough
         On Error GoTo 0
         
         frmtouroku2.txtsonotashoukei.Text = saki_sonotashoukei
         frmtouroku2.txtsonotaseikyuu.Text = saki_sonotaseikyuu
        frmtouroku2.txtsoukei.Text = saki_soukei
        a_sonota_henkou = True
    End If
Exit Function

errnewkin4:
    MsgBox "個人データの更新に失敗しました。"
    a_sonota_henkou = False
    Exit Function


End Function

Function a_sonota_touroku(koid As String, sonotakin As String, joutai As String, Optional voipno As Integer) As Boolean


'******************************
'その他契約変更時の個人データの登録
'******************************
Dim rs_shukeiyakukoushin As Recordset
Dim gen_sonotashoukei As String, gen_soukei As String
Dim saki_sonotashoukei As String, saki_sonotaseikyuu As String, saki_soukei As String

    sql = "select * from kojin where kojinid ='" & koid & "'"
    If FcSQlGet(rs_shukeiyakukoushin, sql, PrMsg) = False Then
        MsgBox "更新先登録済個人データがありません。"
        a_sonota_touroku = False
        Exit Function
    Else
        '元のデータ
Dim moto_shu_seikyuu As Long, moto_sonota_seikyuu As Long
Dim moto_sannnyuu As Long
         '主契約請求金
            If Trim(rs_shukeiyakukoushin!jikaikin) = "" Then
                moto_shu_seikyuu = 0
            Else
                moto_shu_seikyuu = CLng(rs_shukeiyakukoushin!jikaikin)
            End If
         'その他請求金
            If Trim(rs_shukeiyakukoushin!sonotakin) = "" Then
                moto_sonota_seikyuu = 0
            Else
                moto_sonota_seikyuu = CLng(rs_shukeiyakukoushin!sonotakin)
            End If
        '未入金額
            If Trim(rs_shukeiyakukoushin!minyuukin) = "" Then
                moto_sannnyuu = 0
            Else
                moto_sannnyuu = CLng(rs_shukeiyakukoushin!minyuukin)
            End If
        
        'その他小計
        gen_sonotashoukei = rs_shukeiyakukoushin!jikaita
        If Trim(gen_sonotashoukei) = "" Then
            gen_sonotashoukei = "0"
        End If
        '総合計
        gen_soukei = rs_shukeiyakukoushin!soukei
        
        
        '新しいその他小計
            saki_sonotashoukei = CStr(CLng(gen_sonotashoukei) + CLng(sonotakin))
            If joutai = "0" Then '普通
                '新しいその他請求
                saki_sonotaseikyuu = CStr(CLng(moto_sonota_seikyuu) + CLng(sonotakin))
                '新しい総合計
                saki_soukei = CStr(CLng(gen_soukei) + CLng(sonotakin))
            Else '〆中
                saki_sonotaseikyuu = moto_sonota_seikyuu
                saki_soukei = gen_soukei
            End If
        rs_shukeiyakukoushin.Close
        
        
        sql = "UPDATE kojin SET sonotakin = '" & CStr(saki_sonotaseikyuu) & "'," & _
            "soukei = '" & CStr(saki_soukei) & "',jikaita = '" & CStr(saki_sonotashoukei) & "'" & _
            " WHERE kojinid = '" & koid & "'"
        On Error GoTo errnewkin4
          db.Execute sql, dbSQLPassThrough
         On Error GoTo 0
         If voipno <> 1 Then
             frmtouroku2.txtsonotashoukei.Text = saki_sonotashoukei
             frmtouroku2.txtsonotaseikyuu.Text = saki_sonotaseikyuu
            frmtouroku2.txtsoukei.Text = saki_soukei
        End If
        
        a_sonota_touroku = True
    End If
Exit Function

errnewkin4:
    MsgBox "個人データの更新に失敗しました。"
    a_sonota_touroku = False
    Exit Function









End Function
Function a_shukeiyaku_sakujo(koid As String, maekin As String, maeshurui As String) As Boolean


'******************************
'主契約変更時の個人データの削除
'******************************
Dim rs_shukeiyakukoushin As Recordset
Dim gen_shukeiyakushoukei As String, gen_soukei As String
Dim saki_shukeiyakushoukei As String, saki_shukeiyakuseikyuu As String, saki_soukei As String

    sql = "select * from kojin where kojinid ='" & koid & "'"
    If FcSQlGet(rs_shukeiyakukoushin, sql, PrMsg) = False Then
        MsgBox "更新先登録済個人データがありません。"
        a_shukeiyaku_sakujo = False
        Exit Function
    Else
        '元のデータ
Dim moto_shu_seikyuu As Long, moto_sonota_seikyuu As Long
Dim moto_sannnyuu As Long
         '主契約請求金
            If Trim(rs_shukeiyakukoushin!jikaikin) = "" Then
                moto_shu_seikyuu = 0
            Else
                moto_shu_seikyuu = CLng(rs_shukeiyakukoushin!jikaikin)
            End If
         'その他請求金
            If Trim(rs_shukeiyakukoushin!sonotakin) = "" Then
                moto_sonota_seikyuu = 0
            Else
                moto_sonota_seikyuu = CLng(rs_shukeiyakukoushin!sonotakin)
            End If
        '未入金額
            If Trim(rs_shukeiyakukoushin!minyuukin) = "" Then
                moto_sannnyuu = 0
            Else
                moto_sannnyuu = CLng(rs_shukeiyakukoushin!minyuukin)
            End If
        '主契約小計
            gen_shukeiyakushoukei = rs_shukeiyakukoushin!jikaishu
            If Trim(gen_shukeiyakushoukei) = "" Then
                gen_shukeiyakushoukei = "0"
            End If
         '総合計
            gen_soukei = rs_shukeiyakukoushin!soukei
        '新しい主契約小計
        saki_shukeiyakushoukei = CStr(CLng(gen_shukeiyakushoukei) - CLng(maekin))
        
        If maeshurui = "0" Then '「稼動」
                '新しいデータ
                saki_shukeiyakuseikyuu = CStr(CLng(moto_shu_seikyuu) - CLng(maekin))
                saki_soukei = CStr(CLng(gen_soukei) - CLng(maekin))
        Else '「中止」「休止」
                '新しいデータ
                saki_shukeiyakuseikyuu = moto_shu_seikyuu
                saki_soukei = gen_soukei
        End If
        
        rs_shukeiyakukoushin.Close
        
        sql = "UPDATE kojin SET jikaikin = '" & CStr(saki_shukeiyakuseikyuu) & "'," & _
            "soukei = '" & CStr(saki_soukei) & "',jikaishu = '" & CStr(saki_shukeiyakushoukei) & "'" & _
            " WHERE kojinid = '" & koid & "'"
        On Error GoTo errnewkin4
          db.Execute sql, dbSQLPassThrough
         On Error GoTo 0
         frmtouroku2.txtshukeiyakushoukei.Text = saki_shukeiyakushoukei
         frmtouroku2.txtshukeiyakuseikyuu.Text = saki_shukeiyakuseikyuu
        frmtouroku2.txtsoukei.Text = saki_soukei
        a_shukeiyaku_sakujo = True
    End If
Exit Function

errnewkin4:
    MsgBox "個人データの更新に失敗しました。"
    a_shukeiyaku_sakujo = False
    Exit Function








End Function

Function a_keiyaku_henkou(koid As String, maekin As String, gokin As String, maeshurui As String, goshurui As String) As Boolean
'******************************
'主契約変更時の個人データの更新
'******************************
Dim rs_shukeiyakukoushin As Recordset
Dim gen_shukeiyakushoukei As String, gen_shukeiyakuseikyuu As String, gen_soukei As String
Dim saki_shukeiyakushoukei As String, saki_shukeiyakuseikyuu As String, saki_soukei As String

    sql = "select * from kojin where kojinid ='" & koid & "'"
    If FcSQlGet(rs_shukeiyakukoushin, sql, PrMsg) = False Then
        MsgBox "更新先登録済個人データがありません。"
        a_keiyaku_henkou = False
        Exit Function
    Else
        '元のデータ
        gen_shukeiyakushoukei = rs_shukeiyakukoushin!jikaishu
        If Trim(gen_shukeiyakushoukei) = "" Then
            gen_shukeiyakushoukei = "0"
        End If
        gen_shukeiyakuseikyuu = rs_shukeiyakukoushin!jikaikin
        If Trim(gen_shukeiyakuseikyuu) = "" Then
            gen_shukeiyakuseikyuu = "0"
        End If
        gen_soukei = rs_shukeiyakukoushin!soukei
        
        saki_shukeiyakushoukei = CStr(CLng(gen_shukeiyakushoukei) - CLng(maekin) + CLng(gokin))
        
        If maeshurui = "0" Then '変更前が「稼動」
            If goshurui = "0" Then '変更後が同じ
                '新しいデータ
                saki_shukeiyakuseikyuu = CStr(CLng(gen_shukeiyakuseikyuu) - CLng(maekin) + CLng(gokin))
                saki_soukei = CStr(CLng(gen_soukei) - CLng(maekin) + CLng(gokin))
            Else '変更後が「中止」「休止」
                '新しいデータ
                saki_shukeiyakuseikyuu = CStr(CLng(gen_shukeiyakuseikyuu) - CLng(maekin))
                saki_soukei = CStr(CLng(gen_soukei) - CLng(maekin))
            End If
        Else '変更前が「中止」「休止」
            If goshurui = "1" Or goshurui = "2" Then '変更後が同じ
                '新しいデータ
                saki_shukeiyakuseikyuu = gen_shukeiyakuseikyuu
                saki_soukei = gen_soukei
            Else '変更後が「稼動」
                '新しいデータ
                saki_shukeiyakuseikyuu = CStr(CLng(gen_shukeiyakuseikyuu) + CLng(gokin))
                saki_soukei = CStr(CLng(gen_soukei) + CLng(gokin))
            End If
        End If
        
        rs_shukeiyakukoushin.Close
        
        sql = "UPDATE kojin SET jikaikin = '" & CStr(saki_shukeiyakuseikyuu) & "'," & _
            "soukei = '" & CStr(saki_soukei) & "',jikaishu = '" & CStr(saki_shukeiyakushoukei) & "'" & _
            " WHERE kojinid = '" & koid & "'"
        On Error GoTo errnewkin4
          db.Execute sql, dbSQLPassThrough
         On Error GoTo 0
         frmtouroku2.txtshukeiyakushoukei.Text = saki_shukeiyakushoukei
         frmtouroku2.txtshukeiyakuseikyuu.Text = saki_shukeiyakuseikyuu
        frmtouroku2.txtsoukei.Text = saki_soukei
        a_keiyaku_henkou = True
    End If
Exit Function

errnewkin4:
    MsgBox "個人データの更新に失敗しました。"
    a_keiyaku_henkou = False
    Exit Function





End Function

Function a_shukeiyaku_henkou(koid As String, maekin As String, gokin As String, maeshurui As String, goshurui As String) As Boolean
'******************************
'主契約変更時の個人データの更新
'******************************
Dim rs_shukeiyakukoushin As Recordset
Dim gen_shukeiyakushoukei As String, gen_soukei As String, kkkajou As String
Dim saki_shukeiyakushoukei As String, saki_shukeiyakuseikyuu As String, saki_soukei As String

    sql = "select * from kojin where kojinid ='" & koid & "'"
    If FcSQlGet(rs_shukeiyakukoushin, sql, PrMsg) = False Then
        MsgBox "更新先登録済個人データがありません。"
        a_shukeiyaku_henkou = False
        Exit Function
    Else
        '元のデータ
Dim moto_shu_seikyuu As Long, moto_sonota_seikyuu As Long
Dim moto_sannnyuu As Long
         '主契約請求金
            If Trim(rs_shukeiyakukoushin!jikaikin) = "" Then
                moto_shu_seikyuu = 0
            Else
                moto_shu_seikyuu = CLng(rs_shukeiyakukoushin!jikaikin)
            End If
         'その他請求金
            If Trim(rs_shukeiyakukoushin!sonotakin) = "" Then
                moto_sonota_seikyuu = 0
            Else
                moto_sonota_seikyuu = CLng(rs_shukeiyakukoushin!sonotakin)
            End If
        '未入金額
            If Trim(rs_shukeiyakukoushin!minyuukin) = "" Then
                moto_sannnyuu = 0
            Else
                moto_sannnyuu = CLng(rs_shukeiyakukoushin!minyuukin)
            End If
        '主契約小計
            gen_shukeiyakushoukei = rs_shukeiyakukoushin!jikaishu
            If Trim(gen_shukeiyakushoukei) = "" Then
                gen_shukeiyakushoukei = "0"
            End If
        '総合計
            gen_soukei = rs_shukeiyakukoushin!soukei
        '次回請求
            kkkajou = Trim(rs_shukeiyakukoushin!kajou)
            
        '新しい主契約小計
        saki_shukeiyakushoukei = CStr(CLng(gen_shukeiyakushoukei) - CLng(maekin) + CLng(gokin))
        
        If maeshurui = "0" Then '変更前が「稼動」
            If goshurui = "0" Then '変更後が同じ
                '新しいデータ
                saki_shukeiyakuseikyuu = CStr(CLng(moto_shu_seikyuu) - CLng(maekin) + CLng(gokin))
                saki_soukei = CStr(CLng(gen_soukei) - CLng(maekin) + CLng(gokin))
            Else '変更後が「中止」「休止」
                '新しいデータ
                If CLng(moto_shu_seikyuu) = 0 Then
                    saki_shukeiyakuseikyuu = "0"
                    If CLng(gen_soukei) = 0 Then
                        saki_soukei = "0"
                    Else
                        If moto_sannnyuu = gen_soukei Then
                            saki_soukei = gen_soukei
                        Else
                            saki_soukei = CStr(CLng(gen_soukei) - CLng(maekin))
                        End If
                    End If
                Else
                    saki_shukeiyakuseikyuu = CStr(CLng(moto_shu_seikyuu) - CLng(maekin))
                    If CLng(gen_soukei) = 0 Then
                        saki_soukei = "0"
                    Else
                        saki_soukei = CStr(CLng(gen_soukei) - CLng(maekin))
                    End If
                End If
                If kkkajou <> "0" And kkkajou <> "1" Then
                    MsgBox "次回請求日を手動で変更してください。"
                End If
            End If
        Else '変更前が「中止」「休止」
            If goshurui = "1" Or goshurui = "2" Then '変更後が同じ
                '新しいデータ
                saki_shukeiyakuseikyuu = moto_shu_seikyuu
                saki_soukei = gen_soukei
            Else '変更後が「稼動」
                '新しいデータ
                saki_shukeiyakuseikyuu = CStr(CLng(moto_shu_seikyuu) + CLng(gokin))
                saki_soukei = CStr(CLng(gen_soukei) + CLng(gokin))
            End If
        End If
        
        rs_shukeiyakukoushin.Close
        sql = "UPDATE kojin SET jikaikin = '" & CStr(saki_shukeiyakuseikyuu) & "'," & _
            "soukei = '" & CStr(saki_soukei) & "',jikaishu = '" & CStr(saki_shukeiyakushoukei) & "'" & _
            " WHERE kojinid = '" & koid & "'"
        On Error GoTo errnewkin4
          db.Execute sql, dbSQLPassThrough
         On Error GoTo 0
         frmtouroku2.txtshukeiyakushoukei.Text = saki_shukeiyakushoukei
         frmtouroku2.txtshukeiyakuseikyuu.Text = saki_shukeiyakuseikyuu
        frmtouroku2.txtsoukei.Text = saki_soukei
        a_shukeiyaku_henkou = True
    End If
Exit Function

errnewkin4:
    MsgBox "個人データの更新に失敗しました。"
    a_shukeiyaku_henkou = False
    Exit Function





End Function
Function a_shukeiyaku_koushin(koid As String, shiteikaishibi As String, koushinkingaku As String, kadou As String, genjou As String) As Boolean
'******************************
'主契約登録時の個人データの更新
'******************************
Dim rs_shukeiyakukoushin As Recordset
Dim gen_shukeiyakushoukei As String, gen_soukei As String
Dim saki_shukeiyakushoukei As String, saki_shukeiyakuseikyuu As String, saki_soukei As String
Dim moto_shu_seikyuu As Long, moto_sonota_seikyuu As Long
Dim moto_sannnyuu As Long, moto_jikai As String, moto_jikaishu As Long
Dim new_jikai As String

    sql = "select * from kojin where kojinid ='" & koid & "'"
    If FcSQlGet(rs_shukeiyakukoushin, sql, PrMsg) = False Then
        MsgBox "更新先登録済個人データがありません。"
        a_shukeiyaku_koushin = False
        Exit Function
    Else
        If genjou = "1" Then '閉めの時
            '主契約の小計
                gen_shukeiyakushoukei = rs_shukeiyakukoushin!jikaishu
                If Trim(gen_shukeiyakushoukei) = "" Then
                    gen_shukeiyakushoukei = "0"
                End If
            '新しい主契約の小計
                saki_shukeiyakushoukei = CStr(CLng(gen_shukeiyakushoukei) + CLng(koushinkingaku))
                sql = "UPDATE kojin SET jikaishu = '" & CStr(saki_shukeiyakushoukei) & "'" & _
                    " WHERE kojinid = '" & koid & "'"
                On Error GoTo errnewkin4
                  db.Execute sql, dbSQLPassThrough
                 On Error GoTo 0
                 frmtouroku2.txtshukeiyakushoukei.Text = saki_shukeiyakushoukei
                  a_shukeiyaku_koushin = True
        Else '普通の時
             '主契約請求金
                If Trim(rs_shukeiyakukoushin!jikaikin) = "" Then
                    moto_shu_seikyuu = 0
                Else
                    moto_shu_seikyuu = CLng(rs_shukeiyakukoushin!jikaikin)
                End If
             'その他請求金
                If Trim(rs_shukeiyakukoushin!sonotakin) = "" Then
                    moto_sonota_seikyuu = 0
                Else
                    moto_sonota_seikyuu = CLng(rs_shukeiyakukoushin!sonotakin)
                End If
            '未入金額
                If Trim(rs_shukeiyakukoushin!minyuukin) = "" Then
                    moto_sannnyuu = 0
                Else
                    moto_sannnyuu = CLng(rs_shukeiyakukoushin!minyuukin)
                End If
            '主契約小計
                gen_shukeiyakushoukei = rs_shukeiyakukoushin!jikaishu
                If Trim(gen_shukeiyakushoukei) = "" Then
                    gen_shukeiyakushoukei = "0"
                End If
            '総合計
                gen_soukei = rs_shukeiyakukoushin!soukei
            
            
            '新しいデータ主契約小計
                saki_shukeiyakushoukei = CStr(CLng(gen_shukeiyakushoukei) + CLng(koushinkingaku))
                If kadou = "0" Then
                    saki_shukeiyakuseikyuu = CStr(CLng(moto_shu_seikyuu) + CLng(koushinkingaku))
                    saki_soukei = CStr(CLng(gen_soukei) + CLng(koushinkingaku))
                Else
                    saki_shukeiyakuseikyuu = moto_shu_seikyuu
                    saki_soukei = gen_soukei
                End If
                rs_shukeiyakukoushin.Close
                    sql = "UPDATE kojin SET jikaikin = '" & CStr(saki_shukeiyakuseikyuu) & "'," & _
                    "soukei = '" & CStr(saki_soukei) & "',jikaishu = '" & CStr(saki_shukeiyakushoukei) & "'" & _
                    " WHERE kojinid = '" & koid & "'"
                On Error GoTo errnewkin4
                  db.Execute sql, dbSQLPassThrough
                 On Error GoTo 0
                 frmtouroku2.txtshukeiyakushoukei.Text = saki_shukeiyakushoukei
                 frmtouroku2.txtshukeiyakuseikyuu.Text = saki_shukeiyakuseikyuu
                frmtouroku2.txtsoukei.Text = saki_soukei
                a_shukeiyaku_koushin = True
          End If
    End If
Exit Function

errnewkin4:
    MsgBox "個人データの更新に失敗しました。"
    a_shukeiyaku_koushin = False
    Exit Function

End Function

Function flgup2(flagtype As Integer, flgkojinid As String)

Dim newjikaibi As String


Select Case flagtype
Case 1
        sql = "UPDATE kojin SET f1 = '1'" & _
            " WHERE kojinid = '" & flgkojinid & "'"
       On Error GoTo errflg
         db.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        
        flgup2 = True
        Exit Function
Case 0
        sql = "UPDATE kojin SET f1 = '0'" & _
            " WHERE kojinid = '" & flgkojinid & "'"
       On Error GoTo errflg
         db.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        
        flgup2 = True
        Exit Function
End Select
errflg:
 flgup2 = False
 Exit Function
End Function
    


Function flgup(flagtype As Integer, flgkojinid As String)

Dim newjikaibi As String


Select Case flagtype
Case 1
        sql = "UPDATE kojin SET f1 = '1'" & _
            " WHERE kojinid = '" & flgkojinid & "'"
       On Error GoTo errflg
         db.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        
        flgup = True
        Exit Function
Case 0
    newjikaibi = Format(Date, "yyyy") & Format(Date, "mm")
        sql = "UPDATE kojin SET f1 = '0',jikaibi = '" & newjikaibi & "'" & _
            " WHERE kojinid = '" & flgkojinid & "'"
       On Error GoTo errflg
         db.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        
        flgup = True
        Exit Function
End Select
errflg:
 flgup = False
 Exit Function
End Function

Function FcCmbshuruiidGet(PrID As String) As Integer

    Dim i As Integer
    
    For i = 0 To shuruicount
        If shuruiidtbl(i) = PrID Then
            Exit For
        End If
    Next
    
    FcCmbshuruiidGet = i

End Function
Function FcCmbcellidGet(PrID As String) As Integer

    Dim i As Integer
    
    For i = 0 To cellcount
        If cellidtbl(i) = PrID Then
            Exit For
        End If
    Next
    
    FcCmbcellidGet = i

End Function

Function FcCmbbankidGet(PrID As String) As Integer

    Dim i As Integer
    
    For i = 0 To bankcount
        If bankidtbl(i) = PrID Then
            Exit For
        End If
    Next
    
    FcCmbbankidGet = i

End Function
Function FcCmbbankidGet2(PrID As String) As Integer

    Dim i As Integer
    
    For i = 0 To bankcount2
        If bankidtbl2(i) = PrID Then
            Exit For
        End If
    Next
    
    FcCmbbankidGet2 = i

End Function

Sub chk_kyuushi(skid As String, Optional onoff As String)

Dim ugoiteru As Long, zenlen As Integer, atolen As Integer, motod As String, gen_kyuu As Integer
Dim motonof3 As String

ugoiteru = 0
motonof3 = ""

'チェック
If onoff = "" Then
    sql = "select f3 from kojin where kojinid ='" & skid & "'"
    If FcSQlGet(rs, sql, PrMsg) = True Then
        motonof3 = rs!f3
        rs.Close
    End If
    
    sql = "SELECT keiyaku.keiyakuid FROM shurui RIGHT JOIN keiyaku ON shurui.shuruiid = keiyaku.shurui" & _
        " WHERE keiyaku.kojinid ='" & skid & "' and keiyaku.kadou='0'" & _
        " order by shurui.shuruiid"

    If FcSQlGet(rs, sql, PrMsg) = True Then
        ugoiteru = 1
        rs.Close
    Else
    
    
    
        kidou4
        sql = "select * from nhk_m where kojinid='" & skid & "'"
        If FcSQlGet4(rs, sql, PrMsg) = 1 Then
            ugoiteru = 1
            rs.Close
        End If
        db4.Close
    End If
Else
    If onoff = "on" Then
        ugoiteru = 0
        motonof3 = "1"
    Else
        ugoiteru = 1
        motonof3 = "0"
    End If
End If

'実行
If motonof3 = CStr(ugoiteru) Then
        If ugoiteru = 1 Then '稼動
               sql = "UPDATE kojin SET f3 = '0'" & _
                          " WHERE kojinid = '" & skid & "'"
                db.Execute sql, dbSQLPassThrough
                frmtouroku2.grdtourokujoukyou.Row = 13
                zenlen = InStr(frmtouroku2.grdtourokujoukyou.Text, "取")
                atolen = InStr(frmtouroku2.grdtourokujoukyou.Text, "契")
                motod = Mid(frmtouroku2.grdtourokujoukyou.Text, 1, zenlen - 2) & "◆取引状況：通常" & Mid(frmtouroku2.grdtourokujoukyou.Text, atolen - 6)
                frmtouroku2.grdtourokujoukyou.Text = motod
                frmtouroku2.grdtourokujoukyou.Row = 17
                frmtouroku2.lblkeiyaku.Caption = "通常"
        Else '「休止」
               sql = "UPDATE kojin SET f3 = '1'" & _
                          " WHERE kojinid = '" & skid & "'"
                 db.Execute sql, dbSQLPassThrough
                frmtouroku2.grdtourokujoukyou.Row = 13
                zenlen = InStr(frmtouroku2.grdtourokujoukyou.Text, "取")
                atolen = InStr(frmtouroku2.grdtourokujoukyou.Text, "契")
                motod = Mid(frmtouroku2.grdtourokujoukyou.Text, 1, zenlen - 2) & "◆取引状況：休止" & Mid(frmtouroku2.grdtourokujoukyou.Text, atolen - 6)
                frmtouroku2.grdtourokujoukyou.Text = motod
                frmtouroku2.grdtourokujoukyou.Row = 17
                frmtouroku2.lblkeiyaku.Caption = "休止"
        End If
End If
        

End Sub
Sub chk_log(kid As String)

Dim chkrs As Recordset

sql = "select*from log where kojinid='" & kid & "'"
If FcSQlGet(chkrs, sql, PrMsg) = True Then
    If chkrs.RecordCount = 0 Then
        frmtouroku2.Label1(1).BackColor = &HFFFFFF
    Else
        frmtouroku2.Label1(1).BackColor = &HFF&
    End If
    chkrs.Close
Else
    frmtouroku2.Label1(1).BackColor = &HFFFFFF
End If


End Sub
Sub logset2(owarinano As Integer, no As Integer, shiteimoji As String)
'種類の一覧を表示します。

Dim lgsuu As Long, zenlen As Integer, atolen As Integer, motod As String

Screen.MousePointer = 11
frmkeiyakurireki.gridlog.Rows = 1
frmkeiyakurireki.gridlog.Clear
 If owarinano = 0 Then
    
        sql = "SELECT log.*,kojin.name FROM log left join kojin on log.kojinid = kojin.kojinid" & _
                " where log.youken like '*" & shiteimoji & "*'"
    
Else
    
        sql = "SELECT log.*,kojin.name FROM log left join kojin on log.kojinid = kojin.kojinid" & _
                " WHERE log.kekka ='0' and log.youken like '*" & shiteimoji & "*'"
    
End If

If frmkeiyakurireki.chksakujo.Value = "0" Then
     
        sql = sql & " and log.sakujo ='0'"
    
End If

sql = sql & " order by log.nichiji DESC,log.jikan DESC"

On Error GoTo errsuuji2

If FcSQlGet(rs, sql, PrMsg) = False Then
    frmkeiyakurireki.gridlog.Rows = 1
    Screen.MousePointer = 0
Else
    rs.MoveLast
    lgsuu = rs.RecordCount
    With frmkeiyakurireki.gridlog
        .Rows = lgsuu + 1
        .ColWidth(0) = 0
        .ColWidth(1) = 1400
        .ColWidth(2) = 800
        .ColWidth(3) = 700
        If no = 0 Then
            .ColWidth(4) = 1300
        Else
            .ColWidth(4) = 0
        End If
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        If no = 0 Then
            .ColWidth(8) = 6000
        Else
            .ColWidth(8) = 7300
        End If
        .ColWidth(9) = 600
        .ColWidth(10) = 0
        .ColWidth(11) = 350
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "日時"
        .Col = 2
        .Text = "時間"
        .Col = 3
        .Text = "担当"
        .Col = 4
        .Text = "契約者名"
        .Col = 5
        .Text = "相手"
        .Col = 6
        .Text = "連絡先"
        .Col = 7
        .Text = "方法"
        .Col = 8
        .Text = "内容"
        .Col = 9
        .Text = "結果"
        .Col = 11
        .Text = "削"
    End With
    rs.MoveFirst
    For i = 1 To lgsuu
        With frmkeiyakurireki.gridlog
            .Row = i
            .Col = 0
            .CellAlignment = 1
            .Text = rs!logid
            .Col = 1
            .CellAlignment = 1
            .Text = Format(rs!nichiji, "@@@@/@@/@@")
            .Col = 2
            .CellAlignment = 1
            .Text = Format(rs!jikan, "@@:@@")
            .Col = 3
            Select Case rs!tantou
            Case "0"
                .Text = "高安"
            Case "1"
                .Text = "岡部"
            Case "2"
                .Text = "鈴木"
                Case "3"
                .Text = "久保田"
                  Case "4"
                .Text = "三森"
                
                  Case "5"
                .Text = "中村"
                Case "6"
                .Text = "志村"
                Case "7"
                .Text = "萩原"
            Case Else
                .Text = ""
            End Select
            .Col = 4
            .CellAlignment = 1
            If IsNull(rs!Name) Then
            .Text = ""
            Else
                .Text = rs!Name
            End If
            .Col = 5
            .CellAlignment = 1
            If IsNull(rs!dare) Then
            .Text = ""
            Else
            .Text = rs!dare
            End If
            .Col = 6
            .CellAlignment = 1
            If IsNull(rs!saki) Then
            .Text = ""
            Else
                .Text = rs!saki
             End If
             .Col = 7
            Select Case Trim(rs!houhou)
            Case "0"
                .Text = "電話"
            Case "1"
                .Text = "来室"
            Case "2"
                .Text = "訪問"
            Case "3"
                .Text = "  "
            Case "4"
                .Text = "着信"
            Case "5"
                .Text = "発信"
            Case Else
                .Text = ""
            End Select
           .Col = 8
           .CellAlignment = 1
            .Text = rs!youken
            .Col = 9
            Select Case Trim(rs!kekka)
            Case "1"
                .Text = "終了"
            Case Else
                .Text = ""
            End Select
           .Col = 10
           If IsNull(rs!kojinid) Then
                .Text = ""
           Else
                .Text = rs!kojinid
            End If
            If Trim(rs!sakujo) = "1" Then
                .Col = 11
                .Text = "×"
            End If
        End With
       rs.MoveNext
    Next i
    rs.Close
End If
    Screen.MousePointer = 0
    

frmkeiyakurireki.lblend.Caption = CStr(owarinano)

Exit Sub

errsuuji2:

Unload frmlogshousai

Unload frmkeiyakurireki
Screen.MousePointer = 0
MsgBox "データの表示時にエラーが発生しました。"
Exit Sub

End Sub
Sub logset2_2(owarinano As Integer, no As Integer, shiteimoji As String)
'種類の一覧を表示します。

Dim lgsuu As Long, zenlen As Integer, atolen As Integer, motod As String

Screen.MousePointer = 11
frmkeiyakurireki2.gridlog.Rows = 1
frmkeiyakurireki2.gridlog.Clear
 If owarinano = 0 Then
    
        sql = "SELECT log.*,kojin.name FROM log left join kojin on log.kojinid = kojin.kojinid" & _
                " where log.youken like '*" & shiteimoji & "*'"
    
Else
    
        sql = "SELECT log.*,kojin.name FROM log left join kojin on log.kojinid = kojin.kojinid" & _
                " WHERE log.kekka ='0' and log.youken like '*" & shiteimoji & "*'"
    
End If

If frmkeiyakurireki2.chksakujo.Value = "0" Then
     
        sql = sql & " and log.sakujo ='0'"
    
End If

sql = sql & " order by log.nichiji DESC,log.jikan DESC"

On Error GoTo errsuuji2

If FcSQlGet(rs, sql, PrMsg) = False Then
    frmkeiyakurireki2.gridlog.Rows = 1
    Screen.MousePointer = 0
Else
    rs.MoveLast
    lgsuu = rs.RecordCount
    With frmkeiyakurireki2.gridlog
        .Rows = lgsuu + 1
        .ColWidth(0) = 0
        .ColWidth(1) = 1400
        .ColWidth(2) = 800
        .ColWidth(3) = 700
        If no = 0 Then
            .ColWidth(4) = 1300
        Else
            .ColWidth(4) = 0
        End If
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        If no = 0 Then
            .ColWidth(8) = 6000
        Else
            .ColWidth(8) = 7300
        End If
        .ColWidth(9) = 600
        .ColWidth(10) = 0
        .ColWidth(11) = 350
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "日時"
        .Col = 2
        .Text = "時間"
        .Col = 3
        .Text = "担当"
        .Col = 4
        .Text = "契約者名"
        .Col = 5
        .Text = "相手"
        .Col = 6
        .Text = "連絡先"
        .Col = 7
        .Text = "方法"
        .Col = 8
        .Text = "内容"
        .Col = 9
        .Text = "結果"
        .Col = 11
        .Text = "削"
    End With
    rs.MoveFirst
    For i = 1 To lgsuu
        With frmkeiyakurireki2.gridlog
            .Row = i
            .Col = 0
            .CellAlignment = 1
            .Text = rs!logid
            .Col = 1
            .CellAlignment = 1
            .Text = Format(rs!nichiji, "@@@@/@@/@@")
            .Col = 2
            .CellAlignment = 1
            .Text = Format(rs!jikan, "@@:@@")
            .Col = 3
            Select Case rs!tantou
            Case "0"
                .Text = "高安"
            Case "1"
                .Text = "岡部"
            Case "2"
                .Text = "鈴木"
                Case "3"
                .Text = "久保田"
                  Case "4"
                .Text = "三森"
                  Case "5"
                .Text = "中村"
                Case "6"
                .Text = "志村"
                Case "7"
                .Text = "萩原"
            Case Else
                .Text = ""
            End Select
            .Col = 4
            .CellAlignment = 1
            If IsNull(rs!Name) Then
            .Text = ""
            Else
                .Text = rs!Name
            End If
            .Col = 5
            .CellAlignment = 1
            If IsNull(rs!dare) Then
            .Text = ""
            Else
            .Text = rs!dare
            End If
            .Col = 6
            .CellAlignment = 1
            If IsNull(rs!saki) Then
            .Text = ""
            Else
                .Text = rs!saki
             End If
             .Col = 7
            Select Case Trim(rs!houhou)
            Case "0"
                .Text = "電話"
            Case "1"
                .Text = "来室"
            Case "2"
                .Text = "訪問"
            Case "3"
                .Text = "  "
            Case "4"
                .Text = "着信"
            Case "5"
                .Text = "発信"
            Case Else
                .Text = ""
            End Select
           .Col = 8
           .CellAlignment = 1
            .Text = rs!youken
            .Col = 9
            Select Case Trim(rs!kekka)
            Case "1"
                .Text = "終了"
            Case Else
                .Text = ""
            End Select
           .Col = 10
           If IsNull(rs!kojinid) Then
                .Text = ""
           Else
                .Text = rs!kojinid
            End If
            If Trim(rs!sakujo) = "1" Then
                .Col = 11
                .Text = "×"
            End If
        End With
       rs.MoveNext
    Next i
    rs.Close
End If
    Screen.MousePointer = 0
    

frmkeiyakurireki2.lblend.Caption = CStr(owarinano)

Exit Sub

errsuuji2:

Unload frmlogshousai

Unload frmkeiyakurireki2
Screen.MousePointer = 0
MsgBox "データの表示時にエラーが発生しました。"
Exit Sub

End Sub


Sub logset(owarinano As Integer, no As Integer, Optional kid As String, Optional sakujosumi As Integer)
'種類の一覧を表示します。

Dim lgsuu As Long, zenlen As Integer, atolen As Integer, motod As String

Screen.MousePointer = 11
frmkeiyakurireki.gridlog.Rows = 1
frmkeiyakurireki.gridlog.Clear
 If owarinano = 0 Then
    If no = 0 Then
        sql = "SELECT log.*,kojin.name FROM log left join kojin on log.kojinid = kojin.kojinid"
    Else
        sql = "SELECT log.*,kojin.name FROM log left join kojin on log.kojinid = kojin.kojinid WHERE log.kojinid ='" & kid & "'"
    End If
Else
    If no = 0 Then
        sql = "SELECT log.*,kojin.name FROM log left join kojin on log.kojinid = kojin.kojinid WHERE log.kekka ='0'"
    Else
        sql = "SELECT log.*,kojin.name FROM log left join kojin on log.kojinid = kojin.kojinid WHERE log.kojinid ='" & kid & "'" & _
                " and log.kekka ='0'"
    End If
End If

If sakujosumi = 0 Then
     If owarinano = 0 And no = 0 Then
        sql = sql & " WHERE log.sakujo ='0'"
     Else
        sql = sql & "and log.sakujo ='0'"
    End If
End If

sql = sql & " order by log.nichiji DESC,log.jikan DESC"

On Error GoTo errsuuji2

If FcSQlGet(rs, sql, PrMsg) = False Then
    frmkeiyakurireki.gridlog.Rows = 1
    Screen.MousePointer = 0
Else
    rs.MoveLast
    lgsuu = rs.RecordCount
    With frmkeiyakurireki.gridlog
        .Rows = lgsuu + 1
        .ColWidth(0) = 0
        .ColWidth(1) = 1400
        .ColWidth(2) = 800
        .ColWidth(3) = 700
        If no = 0 Then
            .ColWidth(4) = 1300
        Else
            .ColWidth(4) = 0
        End If
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        If no = 0 Then
            .ColWidth(8) = 6000
        Else
            .ColWidth(8) = 7300
        End If
        .ColWidth(9) = 600
        .ColWidth(10) = 0
        .ColWidth(11) = 350
        
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "日時"
        .Col = 2
        .Text = "時間"
        .Col = 3
        .Text = "担当"
        .Col = 4
        .Text = "契約者名"
        .Col = 5
        .Text = "相手"
        .Col = 6
        .Text = "連絡先"
        .Col = 7
        .Text = "方法"
        .Col = 8
        .Text = "内容"
        .Col = 9
        .Text = "結果"
        .Col = 11
        .Text = "削"
    End With
    rs.MoveFirst
    For i = 1 To lgsuu
        With frmkeiyakurireki.gridlog
            .Row = i
            .Col = 0
            .CellAlignment = 1
            .Text = rs!logid
            .Col = 1
            .CellAlignment = 1
            .Text = Format(rs!nichiji, "@@@@/@@/@@")
            .Col = 2
            .CellAlignment = 1
            .Text = Format(rs!jikan, "@@:@@")
            .Col = 3
            Select Case rs!tantou
            Case "0"
                .Text = "高安"
            Case "1"
                .Text = "岡部"
            Case "2"
                .Text = "鈴木"
            Case "3"
                .Text = "久保田"
                Case "4"
                .Text = "三森"
                Case "5"
                .Text = "中村"
                Case "6"
                .Text = "志村"
                Case "7"
                .Text = "萩原"
            Case Else
                .Text = ""
            End Select
            .Col = 4
            .CellAlignment = 1
            If IsNull(rs!Name) Then
            .Text = ""
            Else
                .Text = rs!Name
            End If
            .Col = 5
            .CellAlignment = 1
            If IsNull(rs!dare) Then
            .Text = ""
            Else
            .Text = rs!dare
            End If
            .Col = 6
            .CellAlignment = 1
            If IsNull(rs!saki) Then
            .Text = ""
            Else
                .Text = rs!saki
             End If
             .Col = 7
            Select Case Trim(rs!houhou)
            Case "0"
                .Text = "電話"
            Case "1"
                .Text = "来室"
            Case "2"
                .Text = "訪問"
            Case "3"
                .Text = "  "
            Case "4"
                .Text = "着信"
            Case "5"
                .Text = "発信"
            Case Else
                .Text = ""
            End Select
           .Col = 8
           .CellAlignment = 1
            .Text = rs!youken
            .Col = 9
            Select Case Trim(rs!kekka)
            Case "1"
                .Text = "終了"
            Case Else
                .Text = ""
            End Select
           .Col = 10
           If IsNull(rs!kojinid) Then
                .Text = ""
           Else
                .Text = rs!kojinid
            End If
            If Trim(rs!sakujo) = "1" Then
                .Col = 11
                .Text = "×"
            End If
            
         
            
        End With
       rs.MoveNext
    Next i
    rs.Close
End If
    Screen.MousePointer = 0
    

frmkeiyakurireki.lblend.Caption = CStr(owarinano)

Exit Sub

errsuuji2:

Unload frmlogshousai

Unload frmkeiyakurireki
Screen.MousePointer = 0
MsgBox "データの表示時にエラーが発生しました。"
Exit Sub

End Sub
Sub logset_2(owarinano As Integer, no As Integer, Optional kid As String, Optional sakujosumi As Integer)
'種類の一覧を表示します。

Dim lgsuu As Long, zenlen As Integer, atolen As Integer, motod As String

Screen.MousePointer = 11
frmkeiyakurireki2.gridlog.Rows = 1
frmkeiyakurireki2.gridlog.Clear
 If owarinano = 0 Then
    If no = 0 Then
        sql = "SELECT log.*,kojin.name FROM log left join kojin on log.kojinid = kojin.kojinid"
    Else
        sql = "SELECT log.*,kojin.name FROM log left join kojin on log.kojinid = kojin.kojinid WHERE log.kojinid ='" & kid & "'"
    End If
Else
    If no = 0 Then
        sql = "SELECT log.*,kojin.name FROM log left join kojin on log.kojinid = kojin.kojinid WHERE log.kekka ='0'"
    Else
        sql = "SELECT log.*,kojin.name FROM log left join kojin on log.kojinid = kojin.kojinid WHERE log.kojinid ='" & kid & "'" & _
                " and log.kekka ='0'"
    End If
End If

If sakujosumi = 0 Then
     If owarinano = 0 And no = 0 Then
        sql = sql & " WHERE log.sakujo ='0'"
     Else
        sql = sql & "and log.sakujo ='0'"
    End If
End If

sql = sql & " order by log.nichiji DESC,log.jikan DESC"

On Error GoTo errsuuji2

If FcSQlGet(rs, sql, PrMsg) = False Then
    frmkeiyakurireki2.gridlog.Rows = 1
    Screen.MousePointer = 0
Else
    rs.MoveLast
    lgsuu = rs.RecordCount
    With frmkeiyakurireki2.gridlog
        .Clear
        .Refresh
        .Rows = lgsuu + 1
        .ColWidth(0) = 100
        .ColWidth(1) = 1400
        .ColWidth(2) = 800
        .ColWidth(3) = 700
        If no = 0 Then
            .ColWidth(4) = 1300
        Else
            .ColWidth(4) = 0
        End If
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        If no = 0 Then
            .ColWidth(8) = 6000
        Else
            .ColWidth(8) = 7300
        End If
        .ColWidth(9) = 600
        .ColWidth(10) = 0
        .ColWidth(11) = 350
        
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "日時"
        .Col = 2
        .Text = "時間"
        .Col = 3
        .Text = "担当"
        .Col = 4
        .Text = "契約者名"
        .Col = 5
        .Text = "相手"
        .Col = 6
        .Text = "連絡先"
        .Col = 7
        .Text = "方法"
        .Col = 8
        .Text = "内容"
        .Col = 9
        .Text = "結果"
        .Col = 11
        .Text = "削"
    End With
    rs.MoveFirst
    For i = 1 To lgsuu
        With frmkeiyakurireki2.gridlog
            .Row = i
            .Col = 0
            .CellAlignment = 1
            .Text = rs!logid
            .Col = 1
            .CellAlignment = 1
            .Text = Format(rs!nichiji, "@@@@/@@/@@")
            .Col = 2
            .CellAlignment = 1
            .Text = Format(rs!jikan, "@@:@@")
            .Col = 3
            Select Case rs!tantou
            Case "0"
                .Text = "高安"
            Case "1"
                .Text = "岡部"
            Case "2"
                .Text = "鈴木"
            Case "3"
                .Text = "久保田"
                Case "4"
                .Text = "三森"
                Case "5"
                .Text = "中村"
                Case "6"
                .Text = "志村"
                Case "7"
                .Text = "萩原"
            Case Else
                .Text = ""
            End Select
            .Col = 4
            .CellAlignment = 1
            If IsNull(rs!Name) Then
            .Text = ""
            Else
                .Text = rs!Name
            End If
            .Col = 5
            .CellAlignment = 1
            If IsNull(rs!dare) Then
            .Text = ""
            Else
            .Text = rs!dare
            End If
            .Col = 6
            .CellAlignment = 1
            If IsNull(rs!saki) Then
            .Text = ""
            Else
                .Text = rs!saki
             End If
             .Col = 7
            Select Case Trim(rs!houhou)
            Case "0"
                .Text = "電話"
            Case "1"
                .Text = "来室"
            Case "2"
                .Text = "訪問"
            Case "3"
                .Text = "  "
            Case "4"
                .Text = "着信"
            Case "5"
                .Text = "発信"
            Case Else
                .Text = ""
            End Select
           .Col = 8
           .CellAlignment = 1
            .Text = rs!youken
            .Col = 9
            Select Case Trim(rs!kekka)
            Case "1"
                .Text = "終了"
            Case Else
                .Text = ""
            End Select
           .Col = 10
           If IsNull(rs!kojinid) Then
                .Text = ""
           Else
                .Text = rs!kojinid
            End If
            If Trim(rs!sakujo) = "1" Then
                .Col = 11
                .Text = "×"
            End If
            
         
            
        End With
       rs.MoveNext
    Next i
    rs.Close
End If
    Screen.MousePointer = 0
    

frmkeiyakurireki2.lblend.Caption = CStr(owarinano)

Exit Sub

errsuuji2:

Unload frmlogshousai

Unload frmkeiyakurireki2
Screen.MousePointer = 0
MsgBox "データの表示時にエラーが発生しました。"
Exit Sub

End Sub
Sub keiyakuset(kid As String)

'種類の一覧を表示します。

Dim keiyakusuu As Long, zenlen As Integer, atolen As Integer, motod As String
Screen.MousePointer = 11
 frmbank.CBOREN.Clear
 
 
sql = "SELECT shurui.shuruimei, keiyaku.keiyakuid, keiyaku.kojinid," & _
        "keiyaku.kaishi, keiyaku.kadou, keiyaku.kuchisuu, keiyaku.kingaku,keiyaku.shime" & _
        " FROM shurui RIGHT JOIN keiyaku ON shurui.shuruiid = keiyaku.shurui" & _
        " WHERE keiyaku.kojinid ='" & kid & "'" & _
        " order by keiyaku.kadou,shurui.shuruiid"

If FcSQlGet(rs, sql, PrMsg) = False Then
    frmtouroku2.keiyakugrd.Rows = 1
    frmtouroku2.Picture8.Enabled = False
    frmtouroku2.Picture7.Enabled = False
    Screen.MousePointer = 0
Else
    frmtouroku2.Picture8.Enabled = True
    frmtouroku2.Picture7.Enabled = True
    rs.MoveLast
    keiyakusuu = rs.RecordCount
    With frmtouroku2.keiyakugrd
        .Rows = keiyakusuu + 1
        .ColWidth(0) = 1000
        .ColWidth(1) = 4600
        .ColWidth(2) = 600
        .ColWidth(3) = 1000
        .ColWidth(4) = 1300
        .ColWidth(5) = 600
        .ColWidth(6) = 400
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "契約名"
        .Col = 2
        .Text = "口数"
        .Col = 3
        .Text = "金額"
        .Col = 4
        .Text = "開始日"
        .Col = 5
        .Text = "稼動"
        .Col = 6
        .Text = "〆"
    End With
    rs.MoveFirst
    For i = 1 To keiyakusuu
        With frmtouroku2.keiyakugrd
            .Row = i
            .Col = 0
            .Text = rs!keiyakuid
            .Col = 1
            .CellAlignment = 1
            .Text = rs!shuruimei
            .Col = 2
            .Text = rs!kuchisuu
            .Col = 3
            .Text = rs!kingaku
            .Col = 4
            .Text = Format(rs!kaishi, "@@@@/@@/@@")
            .Col = 5
            If rs!kadou = "0" Then
                .Text = "稼動"
            ElseIf rs!kadou = "1" Then
                .Text = "休止"
            Else
                .Text = "中止"
            End If
            .Col = 6
            If rs!shime = "0" Then
                .Text = "無"
            Else
                .Text = "〆"
            End If
        End With
       rs.MoveNext
    Next i
    rs.Close
End If
    Screen.MousePointer = 0
    


End Sub
Sub keiyakuset2(kid As String)

'種類の一覧を表示します。

Dim keiyakusuu2 As Long

Screen.MousePointer = 11
 frmbank.CBOREN.Clear
'sql = "SELECT sonota.sonota, keiyaku2.keiyakuid2, keiyaku2.kojinid," & _
        "keiyaku2.kaishi, keiyaku2.kingaku" & _
        " FROM sonota RIGHT JOIN keiyaku2 ON sonota.sonotaid = keiyaku2.sonotaid" & _
        " WHERE keiyaku2.kojinid ='" & kid & "' and keiyaku2.f2 <>'2'"
sql = "SELECT keiyaku2.keiyakuid2, keiyaku2.kojinid, sonota.sonota,sonota.sonotaid," & _
    " keiyaku2.kaishi, keiyaku2.f2, keiyaku2.kingaku" & _
     " FROM keiyaku2 LEFT JOIN sonota ON keiyaku2.sonotaid = sonota.sonotaid" & _
     " WHERE keiyaku2.kojinid ='" & kid & "' and keiyaku2.f2 <>'2'"

If FcSQlGet(rs, sql, PrMsg) = False Then
    frmtouroku2.sonotagrd.Rows = 1
    frmtouroku2.Picture4.Enabled = False
    frmtouroku2.Picture5.Enabled = False
    Screen.MousePointer = 0
Else
    frmtouroku2.Picture4.Enabled = True
    frmtouroku2.Picture5.Enabled = True
    rs.MoveLast
    keiyakusuu2 = rs.RecordCount
    With frmtouroku2.sonotagrd
        .Rows = keiyakusuu2 + 1
        .ColWidth(0) = 1000
        .ColWidth(1) = 4700
        .ColWidth(2) = 1000
        .ColWidth(3) = 1300
        .ColWidth(4) = 400
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "契約名"
        .Col = 2
        .Text = "金額"
        .Col = 3
        .Text = "作成日"
        .Col = 4
        .Text = "〆"
    End With
    rs.MoveFirst
    For i = 1 To keiyakusuu2
        With frmtouroku2.sonotagrd
            .Row = i
            .Col = 0
            .Text = rs!keiyakuid2
            .Col = 1
            .CellAlignment = 1
            .Text = rs!sonota
            .Col = 2
            .Text = rs!kingaku
            .Col = 3
            .Text = Format(rs!kaishi, "@@@@/@@/@@")
            .Col = 4
            If rs!f2 = "0" Then
                .Text = "無"
            Else
                .Text = "〆"
            End If
        End With
       rs.MoveNext
    Next i
    rs.Close
End If

    Screen.MousePointer = 0
    


End Sub
Function nhk_set(kojinid As String) As Integer

Dim rs_nhk As Recordset, sql_nhk As String, res_nhk As Integer

If kidou4 = False Then
    nhk_set = -1
    Exit Function
End If

sql_nhk = "select * from nhk_m where kojinid='" & kojinid & "'"
res_nhk = FcSQlGet4(rs_nhk, sql_nhk, PrMsg)
With frmtouroku2
If res_nhk = -1 Then
    MsgBox "予期せぬ事由によってエラーが発生しました。"
    nhk_set = -1
    Exit Function
ElseIf res_nhk = 1 Then
    .lblumu.Caption = "あり"
    .lblnhkid.Caption = rs_nhk!nhkid
    .Picture11.Enabled = True
    Select Case rs_nhk!shurui
    Case "0"
        .lblshurui.Caption = "口座引落１ヶ月"
    Case "1"
        .lblshurui.Caption = "口座引落２ヶ月"
    Case "2"
        .lblshurui.Caption = "口座引落３ヶ月"
    Case "3"
        .lblshurui.Caption = "口座引落４ヶ月"
    Case "4"
        .lblshurui.Caption = "口座引落５ヶ月"
    Case "5"
        .lblshurui.Caption = "口座引落６ヶ月"
    Case "6"
        .lblshurui.Caption = "口座引落６ヶ月ハンデ"
    Case "7"
        .lblshurui.Caption = "口座引落６ヶ月家族割"
    End Select
    .lblkuchisuu.Caption = rs_nhk!n1
    .lblkingaku.Caption = rs_nhk!kingaku
    If Trim(rs_nhk!zenkai) = "" Then
        .lblmaebi.Caption = ""
    Else
        .lblmaebi.Caption = Format(Mid(rs_nhk!zenkai, 1, 6), "@@@@/@@")
    End If
    If Trim(rs_nhk!jikai) = "" Then
        .lbljikaibi.Caption = ""
    Else
        .lbljikaibi.Caption = Format(Mid(rs_nhk!jikai, 1, 6), "@@@@/@@")
        If Len(Trim(rs_nhk!jikai)) = 6 Then
            MsgBox "次回月が不正です。ＨＮＫの変更を実行してください。"
        End If
    End If
    If Trim(rs_nhk!n2) = "0" Then
        .lblorikomi.Caption = "0"
    Else
        .lblorikomi.Caption = "1"
    End If
ElseIf res_nhk = 0 Then
    .lblumu.Caption = "なし"
    .Picture11.Enabled = False
    .lblshurui.Caption = ""
    .lblkingaku.Caption = ""
    .lblmaebi.Caption = ""
    .lbljikaibi.Caption = ""
    .lblnhkid.Caption = ""
    .lblkuchisuu.Caption = ""
    .lblorikomi.Caption = "0"
End If
End With
nhk_set = 1
db4.Close

End Function
Function name_get(s_kojinid As String) As String

Dim hyoujirs As Recordset

sql = "select*from kojin where kojinid='" & s_kojinid & "'"

If FcSQlGet(hyoujirs, sql, PrMsg) = False Then
    name_get = "エラー"
Else
    name_get = hyoujirs!Name
End If

 hyoujirs.Close

End Function

Function hyouji_touroku(shiteid As String) As Boolean
Dim hyoujirs As Recordset, kouzak As String, brs As Recordset, kouzashousai As String, hyoujikingaku As String
Dim hyoujikingaku2 As String, hyoujikingaku3 As String, hyoujikingaku4 As String, hyoujikingaku5 As String, hyoujikingaku6 As String
Dim rs_cell As Recordset, sql_cell As String

Dim s_hikari_ok As Integer

Dim s_map As String

s_hikari_ok = 0

sql = "select*from kojin where kojinid='" & shiteid & "'"
        'sql = "SELECT kojin.*, bank.bankmei" & _
            " FROM bank INNER JOIN kojin ON bank.bankid = kojin.bankid" & _
            " WHERE (((kojin.kojinid)='" & shiteid & "'));"
        If FcSQlGet(hyoujirs, sql, PrMsg) = False Then
            MsgBox "予期せぬ事由によってエラーが発生しました。"
            hyouji_touroku = False
            Exit Function
        End If
        frmtouroku2.lblid.Caption = hyoujirs!kojinid
        frmtouroku2.lblname.Caption = hyoujirs!Name
        With frmtouroku2.grdtourokujoukyou
            .Row = 0
            .Col = 0
            .Text = Space(15) & "◆氏名：" & hyoujirs!Name & "(" & hyoujirs!Furigana & ")"
            If Trim(hyoujirs!linkid) <> "" Then
                .Text = .Text & Space(15) & "◆管理ＩＤ：" & Trim(hyoujirs!linkid)
            End If
            .Row = 1
            s_map = ""
            If hyoujirs!yuubin = Space(1) Then
                .Text = "◆住所：" & hyoujirs!juusho1 & hyoujirs!juusho2
                
            Else
                .Text = "◆住所：" & Format(hyoujirs!yuubin, "@@@-@@@@") & Space(1) & hyoujirs!juusho1 & hyoujirs!juusho2
            End If
            s_map = hyoujirs!juusho1 & hyoujirs!juusho2
            
            frmtouroku2.lblmap.Caption = s_map
            
            If hyoujirs!juutaku = "0" Then
                .Text = .Text & Space(5) & "◆住宅区分：一般家庭"
            ElseIf hyoujirs!juutaku = "1" Then
                .Text = .Text & Space(5) & "◆住宅区分：一般事務所"
            ElseIf hyoujirs!juutaku = "2" Then
                .Text = .Text & Space(5) & "◆住宅区分：公共"
            ElseIf hyoujirs!juutaku = "3" Then
                .Text = .Text & Space(5) & "◆住宅区分：公共１"
            ElseIf hyoujirs!juutaku = "4" Then
                .Text = .Text & Space(5) & "◆住宅区分：集合住宅"
            ElseIf hyoujirs!juutaku = "5" Then
                .Text = .Text & Space(5) & "◆住宅区分：農・事務所"
            ElseIf hyoujirs!juutaku = "6" Then
                .Text = .Text & Space(5) & "◆住宅区分：遊覧園"
            Else
                .Text = .Text & Space(5) & "◆住宅区分：その他"
            End If
            
           
            
            
            If hyoujirs!chiikiid = "0" Then
                .Text = .Text & Space(5) & "◆地域区分：Ａ　勝沼１～１４区"
            ElseIf hyoujirs!chiikiid = "1" Then
                .Text = .Text & Space(5) & "◆地域区分：Ｂ　祝　１～１０区"
            ElseIf hyoujirs!chiikiid = "2" Then
                .Text = .Text & Space(5) & "◆地域区分：Ｃ　東雲１～１５区"
            ElseIf hyoujirs!chiikiid = "3" Then
                .Text = .Text & Space(5) & "◆地域区分：Ｄ　菱山１～１０区"
            End If
            .Text = .Text & Space(3) & hyoujirs!cellno
            .Row = 2
            .Text = "◆電話番号１：" & hyoujirs!tel1 & Space(4) & "◆電話番号２：" & hyoujirs!tel2 & Space(4) & "◆携帯電話：" & hyoujirs!keitai & Space(4) & "◆ＦＡＸ番号：" & hyoujirs!fax
            .Row = 3
            .Text = "◆世帯数：" & hyoujirs!setaisuu & Space(4) & "◆ＴＶ数：" & hyoujirs!tvsuu & Space(4) & "◆保証金(HT)：" & hyoujirs!ht4
            
            If hyoujirs!qshiyou = "1" Then
                .Text = .Text & " ◆残金：「Ｑ」あり" & " ◆関連日："
                If Trim(hyoujirs!qhiduke) <> "" Then
                    .Text = .Text & Format(hyoujirs!qhiduke, "@@@@/@@")
                End If
            Else
                .Text = .Text & " ◆残金：「Ｑ」なし" & " ◆関連日："
                 If Trim(hyoujirs!qhiduke) <> "" Then
                    .Text = .Text & Format(hyoujirs!qhiduke, "@@@@/@@")
                End If
            End If
           
            .Row = 4
            If Trim(hyoujirs!bikou) = "" Then
                frmtouroku2.cmbbikou.Visible = False
            Else
                frmtouroku2.cmbbikou.Visible = True
                .Text = "◆備考：" & hyoujirs!bikou
                frmtouroku2.lblbikou.Caption = hyoujirs!bikou
            End If
            
            .Row = 5
            If IsNull(hyoujirs!y_yuubin) Then
                .Text = "◆郵送先："
            Else
                If Trim(hyoujirs!y_yuubin) = "" Then
                   .Text = "◆郵送先："
                Else
                   .Text = "◆郵送先：" & Trim(hyoujirs!y_yuubin)
                End If
            End If
            If IsNull(hyoujirs!y_juusho1) Then
                .Text = .Text
            Else
                If Trim(hyoujirs!y_juusho1) = "" Then
                   .Text = .Text
                Else
                   .Text = .Text & " " & Trim(hyoujirs!y_juusho1)
                End If
            End If
            If IsNull(hyoujirs!y_juusho2) Then
                .Text = .Text
            Else
                If Trim(hyoujirs!y_juusho2) = "" Then
                   .Text = .Text
                Else
                   .Text = .Text & Trim(hyoujirs!y_juusho2)
                End If
            End If
            
            If IsNull(hyoujirs!y_meishou) Then
                .Text = .Text
            Else
                If Trim(hyoujirs!y_meishou) = "" Then
                   .Text = .Text
                Else
                   .Text = .Text & Trim(hyoujirs!y_meishou)
                End If
            End If
            
            If IsNull(hyoujirs!saishuukoushinbi) Then
                .Text = .Text & Space(5) & "◆最終契約更新日："
            Else
                If Trim(hyoujirs!saishuukoushinbi) = "" Then
                   .Text = .Text & Space(5) & "◆最終契約更新日："
                Else
                   .Text = .Text & Space(5) & "◆最終契約更新日：" & Trim(hyoujirs!saishuukoushinbi)
                End If
            End If
            
                
            .Row = 6
            .Text = "◆設置先名：" & hyoujirs!k_name & "(" & hyoujirs!k_furigana & ")"
            
            If IsNull(hyoujirs!h_bikou) Then
                .Text = .Text & Space(5) & "◆備考(光)："
            Else
                .Text = .Text & Space(5) & "◆備考(光)：" & hyoujirs!h_bikou
            End If
            .Row = 7
            .Text = "◆設置先住所：" & hyoujirs!k_juusho
            If Trim(hyoujirs!cellid) = "" Then
                .Text = .Text & Space(5) & "◆ノード番号：未設定"
            Else
                'sql_cell = "SELECT * FROM cell WHERE cellid ='" & hyoujirs!cellid & "'"
                'If FcSQlGet(rs_cell, sql_cell, PrMsg) = False Then
                '    MsgBox "予期せぬ事由によりエラーが発生しました。"
                'Else
                    .Text = .Text & Space(5) & "◆ノード番号：" & Trim(hyoujirs!cellid) & Space(4) & "◆図面番号：" & hyoujirs!hoshoukin
                 '   rs_cell.Close
                'End If
            End If
            
            .Row = 8
            .Text = "◆設置先メモ：" & hyoujirs!k_memo
            
            .Row = 9
            .Text = "◆設置先電話番号：" & hyoujirs!k_tel
            
            If IsNull(hyoujirs!h_koujibangou) Then
                .Text = .Text & Space(5) & "◆工区番号：未設定"
            Else
            If Trim(hyoujirs!h_koujibangou) = "" Then
                .Text = .Text & Space(5) & "◆工区番号：未設定"
            Else
                .Text = .Text & Space(5) & "◆工区番号：" & hyoujirs!h_koujibangou
            End If
            End If
            If IsNull(hyoujirs!h_zumen) Then
                .Text = .Text & Space(5) & "◆図面：未設定"
            Else
            If Trim(hyoujirs!h_zumen) = "" Then
                .Text = .Text & Space(5) & "◆図面：未設定"
            Else
                .Text = .Text & Space(5) & "◆図面：" & hyoujirs!h_zumen
            End If
            End If
            
            If IsNull(hyoujirs!h_listbangou) Then
                .Text = .Text & Space(5) & "◆リスト番号：未設定"
            Else
            If Trim(hyoujirs!h_listbangou) = "" Then
                .Text = .Text & Space(5) & "◆リスト番号：未設定"
            Else
                .Text = .Text & Space(5) & "◆リスト番号：" & hyoujirs!h_listbangou
            End If
            End If
            
            If IsNull(hyoujirs!h_koujikanryoubi) Then
                .Text = .Text & Space(5) & "◆工事完了日：未設定"
            Else
                If Trim(hyoujirs!h_koujikanryoubi) = "" Then
                    .Text = .Text & Space(5) & "◆工事完了日：未設定"
                Else
                    .Text = .Text & Space(5) & "◆工事完了日：" & Format(hyoujirs!h_koujikanryoubi, "@@@@/@@/@@")
                    s_hikari_ok = 1
                End If
            End If
            
            
     
            .Row = 10
             .Text = ""
             
             If IsNull(hyoujirs!cl_bangou) Then
                .Text = "◆クロージャー番号：未設定"
            Else
                If Trim(hyoujirs!cl_bangou) = "" Then
                    .Text = "◆クロージャー番号：未設定"
                Else
                    .Text = "◆クロージャー番号：" & Trim(hyoujirs!cl_bangou)
                End If
            End If
            
             If IsNull(hyoujirs!v_onu_adress) Then
                .Text = .Text & Space(5) & "◆V-ONUアドレス：未設定"
            Else
                If Trim(hyoujirs!v_onu_adress) = "" Then
                    .Text = .Text & Space(5) & "◆V-ONUアドレス：未設定"
                Else
                    .Text = .Text & Space(5) & "◆V-ONUアドレス：" & Trim(hyoujirs!v_onu_adress)
                End If
            End If
            
             If IsNull(hyoujirs!d_onu_adress) Then
                .Text = .Text & Space(5) & "◆D-ONUアドレス：未設定"
            Else
                If Trim(hyoujirs!d_onu_adress) = "" Then
                    .Text = .Text & Space(5) & "◆D-ONUアドレス：未設定"
                Else
                    .Text = .Text & Space(5) & "◆D-ONUアドレス：" & Trim(hyoujirs!d_onu_adress)
                End If
            End If
             
            
            .Row = 11
            '.Text = "◆ホームターミナル：   No1/" & hyoujirs!ht1 & Space(5) & "No2/" & hyoujirs!ht2 & Space(5) & "No3/：" & hyoujirs!ht3
            If Trim(hyoujirs!kanyuubi) = "" Then
                If Trim(hyoujirs!kaiyakubi) = "" Then
                    .Text = "◆ＣＡＴＶ加入日：未設定 解約日：未設定"
                Else
                    .Text = "◆ＣＡＴＶ加入日：未設定 解約日：" & Format(hyoujirs!kaiyakubi, "@@@@/@@/@@")
                End If
            Else
                If Trim(hyoujirs!kaiyakubi) = "" Then
                    .Text = "◆ＣＡＴＶ加入日：" & Format(hyoujirs!kanyuubi, "@@@@/@@/@@") & " 解約日：未設定"
                Else
                    .Text = "◆ＣＡＴＶ加入日：" & Format(hyoujirs!kanyuubi, "@@@@/@@/@@") & " 解約日：" & Format(hyoujirs!kaiyakubi, "@@@@/@@/@@")
                End If
            End If
            If Trim(hyoujirs!kanyuubi2) = "" Then
                If Trim(hyoujirs!kaiyakubi2) = "" Then
                    .Text = .Text & " ◆ＮＥＴ加入日：未設定 解約日：未設定"
                Else
                    .Text = .Text & " ◆ＮＥＴ加入日：未設定 解約日：" & Format(hyoujirs!kaiyakubi2, "@@@@/@@/@@")
                End If
            Else
                If Trim(hyoujirs!kaiyakubi2) = "" Then
                    .Text = .Text & " ◆ＮＥＴ加入日：" & Format(hyoujirs!kanyuubi2, "@@@@/@@/@@") & " 解約日：未設定"
                Else
                    .Text = .Text & " ◆ＮＥＴ加入日：" & Format(hyoujirs!kanyuubi2, "@@@@/@@/@@") & " 解約日：" & Format(hyoujirs!kaiyakubi2, "@@@@/@@/@@")
                End If
            End If
            '.Row = 11
          '  hyoujikingaku3 = Trim(hyoujirs!koujihi)
           ' If hyoujikingaku3 <> "" Then
              '  If InStr(1, hyoujikingaku3, "\") = 0 Then
                  '   hyoujikingaku3 = Format(hyoujikingaku3, "\\#,##0;\\-#,##0")
               ' End If
           ' End If
            'hyoujikingaku4 = Trim(hyoujirs!waribiki)
           ' If hyoujikingaku4 <> "" Then
               ' If InStr(1, hyoujikingaku4, "\") = 0 Then
                '     hyoujikingaku4 = Format(hyoujikingaku4, "\\#,##0;\\-#,##0")
               ' End If
           ' End If
           ' hyoujikingaku5 = Trim(hyoujirs!koujihi2)
           ' If hyoujikingaku5 <> "" Then
               ' If InStr(1, hyoujikingaku5, "\") = 0 Then
                '     hyoujikingaku5 = Format(hyoujikingaku5, "\\#,##0;\\-#,##0")
               ' End If
            'End If
            'hyoujikingaku6 = Trim(hyoujirs!waribiki2)
           ' If hyoujikingaku6 <> "" Then
               ' If InStr(1, hyoujikingaku6, "\") = 0 Then
                   '  hyoujikingaku6 = Format(hyoujikingaku6, "\\#,##0;\\-#,##0")
               ' End If
            'End If
            '.Text = "◆ケーブルモデム１：   No1/" & hyoujirs!it1 & Space(5) & "No2/" & hyoujirs!it2 & Space(5) & "No3/" & hyoujirs!it3
           ' .Text = "◆ＣＡＴＶ工事費（割引金）：" & hyoujikingaku3 & "(" & hyoujikingaku4 & ")" & hyoujirs!riyuu
           ' .Text = .Text & "◆ＮＥＴ工事費（割引金）：" & hyoujikingaku5 & "(" & hyoujikingaku6 & ")" & hyoujirs!riyuu2
           
           
            .Row = 12
            frmtouroku2.lblhouhou.Caption = hyoujirs!seikyuu
            Select Case hyoujirs!seikyuu
            Case "0"
                .Text = "◆支払い方法：口座振替"
                frmtouroku2.lblseikyuu.Caption = "0"
            Case "1"
                .Text = "◆支払い方法：現金"
                frmtouroku2.lblseikyuu.Caption = "1"
           End Select
            If hyoujirs!seikyuu = "0" Then
                kouzashousai = ""
                'sql = "select*from bank where bankid='" & hyoujirs!bankid & "'"
                sql = "SELECT bank.bankid, bankmain.bankname, bank.bankmei," & _
                    " bank.bankcode, bank.shitencode" & _
                    " FROM bank LEFT JOIN bankmain ON bank.bankcode = bankmain.bankcode" & _
                    " WHERE bank.bankid ='" & hyoujirs!bankid & "'"
                If FcSQlGet(brs, sql, PrMsg) = False Then
                    MsgBox "予期せぬ事由によりエラーが発生しました。"
                End If
                
                kouzashousai = "◆指定口座：" & brs!bankname & Space(1) & brs!bankmei
                If kouzashousai = "郵便局 郵便局" Then
                    kouzashousai = "郵便局"
                End If
                If hyoujirs!kouzakubun = "0" Then
                    kouzashousai = kouzashousai & Space(3) & "普通"
                Else
                    kouzashousai = kouzashousai & Space(3) & "当座"
                End If
                kouzashousai = kouzashousai & Space(3) & hyoujirs!kouzabangou & Space(3) & hyoujirs!meigi
                .Text = .Text & Space(5) & kouzashousai
            End If
            .Row = 13
            
            .Text = "◆登録日：" & Format(hyoujirs!sakuseibi, "@@@@/@@/@@")
            
            If hyoujirs!f3 = "0" Then
                .Text = .Text & Space(5) & "◆取引状況：通常"
                frmtouroku2.lblkeiyaku.Caption = "通常"
            Else
                .Text = .Text & Space(5) & "◆取引状況：休止"
                frmtouroku2.lblkeiyaku.Caption = "休止"
            End If
            If hyoujirs!f2 = "0" Then
                .Text = .Text & Space(5) & "◆契約状況：通常"
            Else
                .Text = .Text & Space(5) & "◆契約状況：ﾌﾞﾗｯｸ"
            End If
            If hyoujirs!f1 = "1" Then
                .Text = .Text & Space(5) & "◆更新状況：不可"
                frmtouroku2.lbljoukyou.Caption = "不可"
            Else
                .Text = .Text & Space(5) & "◆更新状況：可能"
                frmtouroku2.lbljoukyou.Caption = "可能"
            End If
            If hyoujirs!fkaisuu = "" Then
                .Text = .Text & Space(5) & "◆未引落回数：0"
            Else
                .Text = .Text & Space(5) & "◆未引落回数：" & hyoujirs!fkaisuu
            End If
            
            .Row = 14
            hyoujikingaku = Trim(hyoujirs!nyuukingaku)
            If hyoujikingaku <> "" Then
                If InStr(1, hyoujikingaku, "\") = 0 Then
                     hyoujikingaku = Format(hyoujikingaku, "\\#,##0;\\-#,##0")
                End If
            End If
            hyoujikingaku2 = Trim(hyoujirs!zankin)
            If hyoujikingaku2 <> "" Then
                If InStr(1, hyoujikingaku2, "\") = 0 Then
                     hyoujikingaku2 = Format(hyoujikingaku2, "\\#,##0;\\-#,##0")
                End If
            End If
            
            If IsNull(hyoujirs!nyuukinbi) Then
                .Text = "◆入金日：未設定"
            Else
                If Trim(hyoujirs!nyuukinbi) = "" Then
                    .Text = "◆入金日：未設定"
                Else
                    .Text = "◆入金日：" & Format(hyoujirs!nyuukinbi, "@@@@/@@/@@")
                End If
            End If
            
            .Text = .Text & " ◆入金額：" & hyoujikingaku & _
                    " ◆入金関連日：" & Format(hyoujirs!kanrennbi, "@@@@/@@") & _
                    " ◆残金：" & hyoujikingaku2
                    
            If IsNull(hyoujirs!saishuuhakkoubi) Then
                .Text = .Text & Space(5) & "◆最終印刷日："
            Else
                If Trim(hyoujirs!saishuukoushinbi) = "" Then
                   .Text = .Text & Space(5) & "◆最終印刷日："
                Else
                   .Text = .Text & Space(5) & "◆最終印刷日：" & Trim(hyoujirs!saishuuhakkoubi)
                End If
            End If
            
            .Row = 15
            
            If IsNull(hyoujirs!cl_open_date) Then
                .Text = "◆CL開通日：未設定"
            Else
                If Trim(hyoujirs!cl_open_date) = "" Then
                    .Text = "◆CL開通日：未設定"
                Else
                    .Text = "◆CL開通日：" & Format(hyoujirs!cl_open_date, "@@@@/@@/@@")
                End If
            End If
            If IsNull(hyoujirs!cl_close_date) Then
                .Text = .Text & Space(5) & "◆CL解約日：未設定"
            Else
                If Trim(hyoujirs!cl_close_date) = "" Then
                    .Text = .Text & Space(5) & "◆CL解約日：未設定"
                Else
                    .Text = .Text & Space(5) & "◆CL解約日：" & Format(hyoujirs!cl_close_date, "@@@@/@@/@@")
                End If
            End If
            
            If IsNull(hyoujirs!h_bspass) Then
                .Text = .Text & Space(5) & "◆BSパス工事日：未設定"
            Else
                If Trim(hyoujirs!h_bspass) = "" Then
                    .Text = .Text & Space(5) & "◆BSパス工事日：未設定"
                Else
                    .Text = .Text & Space(5) & "◆BSパス工事日：" & Format(hyoujirs!h_bspass, "@@@@/@@/@@")
                End If
            End If
            
            If IsNull(hyoujirs!cl_keiyaku_name) Then
                .Text = .Text & Space(5) & "◆CL契約者名：未設定"
            Else
                If Trim(hyoujirs!cl_keiyaku_name) = "" Then
                    .Text = .Text & Space(5) & "◆CL契約者名：未設定"
                Else
                    .Text = .Text & Space(5) & "◆CL契約者名：" & Trim(hyoujirs!cl_keiyaku_name)
                End If
            End If
           
           
            
            '.Text = "◆旧口座：" & hyoujirs!kyukouza & " " & hyoujirs!kyukouza2 & " ◆管理口座：" & hyoujirs!kanrikouza
           ' If hyoujirs!qshiyou = "1" Then
               ' .Text = .Text & " ◆残金：「Ｑ」あり" & " ◆関連日：" & Format(hyoujirs!qhiduke, "@@@@/@@")
            'Else
                '.Text = .Text & " ◆残金：「Ｑ」なし" & " ◆関連日：" & Format(hyoujirs!qhiduke, "@@@@/@@")
           ' End If
            
            .Row = 17
      End With
      With frmtouroku2
        
            '光工事完了
            If s_hikari_ok = 1 Then
                .lblid.BackColor = &H8080FF
            Else
                .lblid.BackColor = &HC0FFFF
            End If
      
        
            If Trim(hyoujirs!maebi) = "" Then
                .txtzenbi.Text = ""
                .txtzenbi_moto.Text = ""
            Else
                .txtzenbi.Text = Format(hyoujirs!maebi, "@@@@/@@/@@")
                .txtzenbi_moto.Text = Trim(hyoujirs!maebi)
            End If
            If hyoujirs!maekin = Space(1) Then
                .txtzenkin.Text = "0"
            Else
                .txtzenkin.Text = hyoujirs!maekin
            End If
            If hyoujirs!f4 = Space(1) Then
                .txtnhk.Text = "0"
            Else
                .txtnhk.Text = hyoujirs!f4
            End If
            If hyoujirs!minyuukin = Space(1) Then
                .txtminyuuruikei.Text = "0"
            Else
                .txtminyuuruikei.Text = hyoujirs!minyuukin
            End If
            If hyoujirs!jikaikin = Space(1) Then
                .txtshukeiyakuseikyuu.Text = "0"
            Else
                .txtshukeiyakuseikyuu.Text = hyoujirs!jikaikin
            End If
            If hyoujirs!soukei = Space(1) Then
                .txtsoukei.Text = "0"
            Else
                .txtsoukei.Text = hyoujirs!soukei
            End If
            If hyoujirs!sonotakin = Space(1) Then
                .txtsonotaseikyuu.Text = "0"
            Else
                .txtsonotaseikyuu.Text = hyoujirs!sonotakin
            End If
            If hyoujirs!jikaishu = Space(1) Then
                .txtshukeiyakushoukei.Text = "0"
            Else
                .txtshukeiyakushoukei.Text = hyoujirs!jikaishu
            End If
            If hyoujirs!jikaita = Space(1) Then
                .txtsonotashoukei.Text = "0"
            Else
                .txtsonotashoukei.Text = hyoujirs!jikaita
            End If
            If Trim(hyoujirs!jikaibi) = "" Or Trim(hyoujirs!jikaibi) = "0" Then
                .txtkoushinsumi.Text = ""
            Else
                .txtkoushinsumi.Text = Format(Trim(hyoujirs!jikaibi), "@@@@/@@")
            End If
            If kenchan = 1 Then
                If Trim(hyoujirs!kajou) = "" Or Trim(hyoujirs!kajou) = "0" Then
                    .txtjikaiseikyuubi.Text = "0"
                ElseIf Trim(hyoujirs!kajou) = "1" Then
                    .txtjikaiseikyuubi.Text = "1"
                Else
                    .txtjikaiseikyuubi.Text = Format(Trim(hyoujirs!kajou), "@@@@/@@")
                End If
            Else
                If Trim(hyoujirs!kajou) = "" Or Trim(hyoujirs!kajou) = "0" Or Trim(hyoujirs!kajou) = "1" Then
                    .txtjikaiseikyuubi.Text = ""
                Else
                    .txtjikaiseikyuubi.Text = Format(Trim(hyoujirs!kajou), "@@@@/@@")
                End If
            End If
        End With
        
        hyouji_touroku = True
        
        hyoujirs.Close

End Function

Function FcSQlGet(PrRS As Recordset, PrSQL As String, PrMsg As String) As Boolean

    'ＳＱＬ文の実行
    On Error GoTo ErrSyori
    Set PrRS = db.OpenRecordset(PrSQL, dbOpenDynaset, dbReadOnly)
   
    
    '結果の確認
    If PrRS.RecordCount = 0 Then
        'PrMsg = LoadResString(141)
        FcSQlGet = False
        Exit Function
    Else
        PrRS.MoveLast
        'PrRS.MoveFirst
    End If
    
     On Error GoTo 0
     
    FcSQlGet = True
    Exit Function
    
ErrSyori:
    
    PrMsg = "予期せぬエラーが発生しました。"
    FcSQlGet = False

End Function

Function kari_get(s_kanjaid As String, n_furikou As String, t_furikou As String) As String

kari_get = ""

Dim sql_kari As String, rs_kari As Recordset

 

Dim hi_kari As String
hi_kari = n_furikou & t_furikou
Dim s_modori As String

sql_kari = "select * from koushinkiroku2 where kojinid = '" & s_kanjaid & "' and koushinbi ='" & hi_kari & "'"
If FcSQlGet2(rs_kari, sql_kari, PrMsg) = True Then
    s_modori = ""
    If rs_kari.RecordCount = 1 Then
        Select Case rs_kari!st
        Case "0"
            '.Text = "OK"
            s_modori = rs_kari!koushinbi2
             s_modori = s_modori & rs_kari!kingaku
            
        Case "1"
            '.Text = "NG"
        End Select
        
    End If
    
    
    rs_kari.Close
    
    kari_get = s_modori

End If





End Function


Function FcSQlGet2(PrRS As Recordset, PrSQL As String, PrMsg As String) As Boolean

    'ＳＱＬ文の実行
    On Error GoTo ErrSyori
    Set PrRS = db2.OpenRecordset(PrSQL, dbOpenDynaset, dbReadOnly)
    On Error GoTo 0
    
    '結果の確認
    If PrRS.RecordCount = 0 Then
        'PrMsg = LoadResString(141)
        FcSQlGet2 = False
        Exit Function
    Else
        PrRS.MoveLast
        'PrRS.MoveFirst
    End If
    
    FcSQlGet2 = True
    Exit Function
    
ErrSyori:
    PrMsg = "予期せぬエラーが発生しました。"
    FcSQlGet2 = False

End Function
Function FcSQlGet3(PrRS As Recordset, PrSQL As String, PrMsg As String) As Boolean

    'ＳＱＬ文の実行
    On Error GoTo ErrSyori
    Set PrRS = db3.OpenRecordset(PrSQL, dbOpenDynaset, dbReadOnly)
    On Error GoTo 0
    
    '結果の確認
    If PrRS.RecordCount = 0 Then
        'PrMsg = LoadResString(141)
        FcSQlGet3 = False
        Exit Function
    Else
        PrRS.MoveLast
        'PrRS.MoveFirst
    End If
    
    FcSQlGet3 = True
    Exit Function
    
ErrSyori:
    PrMsg = "予期せぬエラーが発生しました。"
    FcSQlGet3 = False

End Function
Function FcSQlGet4(PrRS As Recordset, PrSQL As String, PrMsg As String) As Integer

    'ＳＱＬ文の実行
    On Error GoTo ErrSyori
    Set PrRS = db4.OpenRecordset(PrSQL, dbOpenDynaset, dbReadOnly)
    On Error GoTo 0
    
    '結果の確認
    If PrRS.RecordCount = 0 Then
        'PrMsg = LoadResString(141)
        FcSQlGet4 = 0
        Exit Function
    Else
        PrRS.MoveLast
        'PrRS.MoveFirst
    End If
    
    FcSQlGet4 = 1
    Exit Function
    
ErrSyori:
    PrMsg = "予期せぬエラーが発生しました。"
    FcSQlGet4 = -1

End Function
Function FcSQlGet_voip(PrRS As Recordset, PrSQL As String, PrMsg As String) As Integer

    'ＳＱＬ文の実行
    On Error GoTo ErrSyori
    Set PrRS = db_voip.OpenRecordset(PrSQL, dbOpenDynaset, dbReadOnly)
    On Error GoTo 0
    
    '結果の確認
    If PrRS.RecordCount = 0 Then
        'PrMsg = LoadResString(141)
        FcSQlGet_voip = 0
        Exit Function
    Else
        PrRS.MoveLast
        'PrRS.MoveFirst
    End If
    
    FcSQlGet_voip = 1
    Exit Function
    
ErrSyori:
    PrMsg = "予期せぬエラーが発生しました。"
    FcSQlGet_voip = -1

End Function
Function FcSQlGet_out(PrRS As Recordset, PrSQL As String, PrMsg As String) As Integer

    'ＳＱＬ文の実行
    On Error GoTo ErrSyori
    Set PrRS = db_out.OpenRecordset(PrSQL, dbOpenDynaset, dbReadOnly)
    On Error GoTo 0
    
    '結果の確認
    If PrRS.RecordCount = 0 Then
        'PrMsg = LoadResString(141)
        FcSQlGet_out = 0
        Exit Function
    Else
        PrRS.MoveLast
        'PrRS.MoveFirst
    End If
    
    FcSQlGet_out = 1
    Exit Function
    
ErrSyori:
    PrMsg = "予期せぬエラーが発生しました。"
    FcSQlGet_out = -1

End Function
Function FcSQlGet_ppv(PrRS As Recordset, PrSQL As String, PrMsg As String) As Integer

    'ＳＱＬ文の実行
    On Error GoTo ErrSyori
    Set PrRS = db_ppv.OpenRecordset(PrSQL, dbOpenDynaset, dbReadOnly)
    On Error GoTo 0
    
    '結果の確認
    If PrRS.RecordCount = 0 Then
        'PrMsg = LoadResString(141)
        FcSQlGet_ppv = 0
        Exit Function
    Else
        PrRS.MoveLast
        'PrRS.MoveFirst
    End If
    
    FcSQlGet_ppv = 1
    Exit Function
    
ErrSyori:
    PrMsg = "予期せぬエラーが発生しました。"
    FcSQlGet_ppv = -1

End Function

Sub frmtourokuclear()
With frmtouroku
    .txtkeiyakushafuri.Text = ""
    .txtkeiyakushamei.Text = ""
    .txtkeiyakushamei2.Text = ""
    .txtyuubin1.Text = ""
    .txtyuubin2.Text = ""
    .comchiiki.ListIndex = -1
    .comjuutaku.ListIndex = 0
    .txtjuusho1.Text = ""
    .txtjuusho2.Text = ""
    .txttel1.Text = ""
    .txttel2.Text = ""
    .txtkeitai.Text = ""
    .txtfax.Text = ""
    .txtbikou.Text = ""
    .txtsecchifuri.Text = ""
    .txtsecchiname.Text = ""
    .txtsecchijuusho.Text = ""
    .txtsecchimemo.Text = ""
    .txtsecchitel.Text = ""
    .txtsetaisuu.Text = ""
    .txttvsuu.Text = ""
    .txtht1.Text = ""
    .txtht2.Text = ""
    .txtht3.Text = ""
    .txtht4.Text = ""
    .txtit1.Text = ""
    .txtit2.Text = ""
    .txtit3.Text = ""
    .optseikyuuhouhou(0).Value = True
    .cmbbank.ListIndex = -1
    .optkouzashurui(0).Value = True
    .txtkouzabangou.Text = ""
    .txtkouzameigi.Text = ""
    .chkblack.Value = 0
    
    .txtkanyuu.Text = ""
    .txtkanyuu2.Text = ""
    .txtkoujihi.Text = ""
    .txtriyuu.Text = ""
    .txtwaribiki.Text = ""
    .txtnyuukinbi.Text = ""
    .txtnyuukingaku.Text = ""
    .txtnyuukinnen.Text = ""
    .txtnyuukintsuki.Text = ""
    .txtzangaku.Text = ""
    .chkq.Value = 0
    .txtqnen.Text = ""
    .txtqtsuki.Text = ""
    .txtkarijuusho.Text = ""
    .txtkanrikouza.Text = ""
    .txtkyukouza1.Text = ""
    .txtkyukouza2.Text = ""
    For i = 0 To 9
        .txtvoip(i).Text = ""
        .txtvoiphiduke(i).Text = ""
    Next i
    .txtnhkno.Text = ""
    .txtnhkname.Text = ""
    .txtnhkfuri.Text = ""
    .txtkoujihi2.Text = ""
    .txtriyuu2.Text = ""
    .txtwaribiki2.Text = ""
    .combcellid.Text = ""
    .combcellno.ListIndex = -1
    
    .txtlinkid.Text = ""
End With
End Sub
Sub sbshuruiset()
'*****************************************
'*   契約種類の名前をコンボボックスに設定      *
'*****************************************

    Dim shuruirs As Recordset
        
    sql = "SELECT * FROM shurui ORDER BY shuruiid"
    
    If FcSQlGet(shuruirs, sql, PrMsg) = False Then
        'MsgBox WMsg
        Exit Sub
    End If
    
    shuruicount = shuruirs.RecordCount
    
    ReDim shuruiidtbl(shuruirs.RecordCount)    '種類ＩＤテーブルの再定義
    shuruirs.MoveFirst
    i = 0
    Do Until shuruirs.EOF
        a = shuruirs!shuruiid & Space(1) & shuruirs!shuruimei
            frmkeiyaku4.cmbkeiyakushurui.AddItem a    'コンボボックスに項目を追加
        shuruiidtbl(i) = shuruirs!shuruiid         '種類ＩＤテーブルに保存
        shuruirs.MoveNext
        i = i + 1
    Loop

    shuruirs.Close

End Sub
Sub sbsonotaset()
'*****************************************
'*   契約種類の名前をコンボボックスに設定      *
'*****************************************

    Dim sonotars As Recordset
        
    sql = "SELECT * FROM sonota ORDER BY sonotaid"
    
    If FcSQlGet(sonotars, sql, PrMsg) = False Then
        'MsgBox WMsg
        Exit Sub
    End If
    
    sonotacount = sonotars.RecordCount
    
    ReDim sonotaidtbl(sonotars.RecordCount)    '種類ＩＤテーブルの再定義
    sonotars.MoveFirst
    i = 0
    Do Until sonotars.EOF
        a = sonotars!sonotaid & Space(1) & sonotars!sonota
            frmkeiyaku5.cmbsonotakeiyaku.AddItem a    'コンボボックスに項目を追加
        sonotaidtbl(i) = sonotars!sonotaid         '種類ＩＤテーブルに保存
        sonotars.MoveNext
        i = i + 1
    Loop

    sonotars.Close


End Sub
Sub shuruiset()

'種類の一覧を表示します。

Dim shuruisuu As Integer

Screen.MousePointer = 11

sql = "SELECT*FROM shurui ORDER BY shuruiid"
If FcSQlGet(rs, sql, PrMsg) = False Then
    frmkeiyaku.SSFTOKUSHU.Visible = False
    Screen.MousePointer = 0
Else
    frmkeiyaku.CBOREN.Clear
    frmkeiyaku.SSFTOKUSHU.Visible = True
    shuruisuu = rs.RecordCount
    rs.MoveFirst
    With frmkeiyaku.shuruigrd
        .Rows = shuruisuu + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 4500
        .ColWidth(2) = 1000
        .ColWidth(3) = 4500
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
        .ColWidth(9) = 1000
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "種類名"
        .Col = 2
        .Text = "金額"
        .Col = 3
        .Text = "説明"
        .Col = 4
        .Text = "分類"
        .Col = 5
        .Text = "集計"
        .Col = 6
        .Text = "Group"
        .Col = 7
        .Text = "Sort"
         .Col = 8
        .Text = "表記(文章)"
         .Col = 9
        .Text = "契約期間修正"
    End With
    For i = 1 To shuruisuu
        With frmkeiyaku.shuruigrd
            .Row = i
            .Col = 0
            .Text = rs!shuruiid
            .Col = 1
            .CellAlignment = 1
            .Text = rs!shuruimei
            .Col = 2
            .Text = rs!tanka
            .Col = 3
            .CellAlignment = 1
            .Text = rs!setsumei
            .Col = 4
            .CellAlignment = 1
            If rs!kankei = "0" Then
                .Text = "CATV"
            Else
                .Text = "INTER"
            End If
            .Col = 5
            .CellAlignment = 1
            If IsNull(rs!iru) Then
            .Text = ""
            Else
            .Text = rs!iru
            End If
              .Col = 6
            .CellAlignment = 1
            If IsNull(rs!gp) Then
            .Text = ""
            Else
            .Text = rs!gp
            End If
              .Col = 7
            .CellAlignment = 1
             If IsNull(rs!narabe) Then
            .Text = ""
            Else
            .Text = rs!narabe
            End If
               .Col = 8
            
             If IsNull(rs!hyoukimei) Then
            .Text = ""
            Else
            .Text = rs!hyoukimei
            End If
               .Col = 9
            
             If IsNull(rs!kasan) Then
            .Text = ""
            Else
            .Text = rs!kasan
            End If
            
            frmkeiyaku.CBOREN.AddItem rs!shuruiid
        End With
       rs.MoveNext
    Next i
    rs.Close
End If
    Screen.MousePointer = 0
    


End Sub
Sub cell_ad_set()
Dim bankrs As Recordset

Screen.MousePointer = 11
    sql = "SELECT * FROM cell ORDER BY cellid"
    
    If FcSQlGet(cellrs, sql, PrMsg) = False Then
        'MsgBox WMsg
        
       Screen.MousePointer = 0
       Exit Sub
    End If
    
   cellcount = cellrs.RecordCount
    
    ReDim cellidtbl(cellcount)    '種類ＩＤテーブルの再定義
    cellrs.MoveFirst
    i = 0
    Do Until cellrs.EOF
        a = cellrs!cellid & Space(1) & cellrs!cellname
        'frmtouroku.combcellid.AddItem a       'コンボボックスに項目を追加
        cellidtbl(i) = cellrs!cellid         '種類ＩＤテーブルに保存
        cellrs.MoveNext
        i = i + 1
    Loop

    cellrs.Close
    Screen.MousePointer = 0

End Sub

Sub sbbankset()
'*****************************************
'*   銀行の名前をコンボボックスに設定      *
'*****************************************

    Dim bankrs As Recordset
Screen.MousePointer = 11
    'sql = "SELECT * FROM bank ORDER BY bankmei"
sql = "SELECT bank.bankid, bankmain.bankname, bank.bankmei" & _
    " FROM bank LEFT JOIN bankmain ON bank.bankcode = bankmain.bankcode" & _
    " ORDER BY  bankmain.bankname,bank.bankmei;"
    
    If FcSQlGet(bankrs, sql, PrMsg) = False Then
        'MsgBox WMsg
        
       Screen.MousePointer = 0
       Exit Sub
    End If
    
   bankcount = bankrs.RecordCount
    
    ReDim bankidtbl(bankrs.RecordCount)    '種類ＩＤテーブルの再定義
    bankrs.MoveFirst
    i = 0
    Do Until bankrs.EOF
        a = bankrs!bankname & Space(1) & bankrs!bankmei
        frmtouroku.cmbbank.AddItem a       'コンボボックスに項目を追加
        bankidtbl(i) = bankrs!bankid         '種類ＩＤテーブルに保存
        bankrs.MoveNext
        i = i + 1
    Loop

    bankrs.Close
    Screen.MousePointer = 0
End Sub
Sub sbnbankset2()
'*****************************************
'*   銀行の名前をコンボボックスに設定      *
'*****************************************

    Dim bankrs As Recordset
        
    sql = "SELECT * FROM bankmain ORDER BY bankcode"
    
    If FcSQlGet(bankrs, sql, PrMsg) = False Then
        'MsgBox WMsg
        Exit Sub
    End If
    
   bankcount2 = bankrs.RecordCount
    
    ReDim bankidtbl2(bankrs.RecordCount)    '種類ＩＤテーブルの再定義
    bankrs.MoveFirst
    i = 0
    Do Until bankrs.EOF
        a = bankrs!bankcode & Space(1) & bankrs!bankname
        frmbank2.txtnewbankcode.AddItem a       'コンボボックスに項目を追加
        bankidtbl2(i) = bankrs!bankcode        '種類ＩＤテーブルに保存
        bankrs.MoveNext
        i = i + 1
    Loop

    bankrs.Close


End Sub
Sub sbnbankset4()
'*****************************************
'*   銀行の名前をコンボボックスに設定      *
'*****************************************

    Dim bankrs As Recordset
        
    sql = "SELECT * FROM bankmain ORDER BY bankcode"
    frmsouken.cmbbank.Clear
    If FcSQlGet(bankrs, sql, PrMsg) = False Then
        'MsgBox WMsg
        Exit Sub
    End If
    
    bankrs.MoveFirst
    Do Until bankrs.EOF
        a = bankrs!bankcode & Space(1) & bankrs!bankname
        frmsouken.cmbbank.AddItem a       'コンボボックスに項目を追加
        bankrs.MoveNext
    Loop

    bankrs.Close

End Sub
Sub sbnbankset3()
'*****************************************
'*   銀行の名前をコンボボックスに設定      *
'*****************************************

    Dim bankrs As Recordset
        
    sql = "SELECT * FROM bankmain ORDER BY bankcode"
    
    If FcSQlGet(bankrs, sql, PrMsg) = False Then
        'MsgBox WMsg
        Exit Sub
    End If
    
   bankcount3 = bankrs.RecordCount
    
    ReDim bankidtbl3(bankrs.RecordCount)    '種類ＩＤテーブルの再定義
    bankrs.MoveFirst
    i = 0
    Do Until bankrs.EOF
        a = bankrs!bankcode & Space(1) & bankrs!bankname
        frmsentaku.combank.AddItem a       'コンボボックスに項目を追加
        bankidtbl3(i) = bankrs!bankcode        '種類ＩＤテーブルに保存
        bankrs.MoveNext
        i = i + 1
    Loop

    bankrs.Close


End Sub

Sub sbnbankset5()
'*****************************************
'*   銀行の名前をコンボボックスに設定      *
'*****************************************

    Dim bankrs As Recordset
        
    sql = "SELECT * FROM bankmain ORDER BY bankcode"
    
    If FcSQlGet(bankrs, sql, PrMsg) = False Then
        'MsgBox WMsg
        Exit Sub
    End If
    
   bankcount4 = bankrs.RecordCount
    
    ReDim bankidtbl4(bankrs.RecordCount)    '種類ＩＤテーブルの再定義
    bankrs.MoveFirst
    i = 0
    Do Until bankrs.EOF
        a = bankrs!bankcode & Space(1) & bankrs!bankname
        frmfdd.combank.AddItem a       'コンボボックスに項目を追加
        bankidtbl4(i) = bankrs!bankcode        '種類ＩＤテーブルに保存
        bankrs.MoveNext
        i = i + 1
    Loop

    bankrs.Close


End Sub
Sub sbnbankset6()
'*****************************************
'*   銀行の名前をコンボボックスに設定      *
'*****************************************

    Dim bankrs As Recordset
        
    sql = "SELECT * FROM bankmain ORDER BY bankcode"
    frmsouken.cmbbank.Clear
    If FcSQlGet(bankrs, sql, PrMsg) = False Then
        'MsgBox WMsg
        Exit Sub
    End If
    
    bankrs.MoveFirst
    Do Until bankrs.EOF
        a = bankrs!bankcode & Space(1) & bankrs!bankname
        frmsouken2.cmbbankid.AddItem a        'コンボボックスに項目を追加
        bankrs.MoveNext
    Loop
    frmsouken2.cmbbankid.AddItem "       現金"
    frmsouken2.cmbbankid.AddItem "9999   取消・損金"
    bankrs.Close

End Sub

Function sonota_hyoujimei_get(s_no As String) As String

sonota_hyoujimei_get = ""

Dim s_str As String
Dim sonotasuu As Integer
s_str = ""

sql = "SELECT*FROM sonota where sonotaid = '" & s_no & "' ORDER BY sonotaid"
If FcSQlGet(rs, sql, PrMsg) = False Then
    frmsonota.SSFTOKUSHU.Visible = False
    Screen.MousePointer = 0
Else
    frmsonota.CBOREN.Clear
    frmsonota.SSFTOKUSHU.Visible = True
    sonotasuu = rs.RecordCount
    rs.MoveFirst
    
    If sonotasuu > 0 Then
        If IsNull(rs!hyoukimei) Then
        s_str = Trim(rs!sonota)
        Else
        s_str = Trim(rs!hyoukimei)
        End If
    End If
    
    rs.Close
End If
sonota_hyoujimei_get = s_str

End Function

Sub sonotaset()

'種類の一覧を表示します。

Dim sonotasuu As Integer

Screen.MousePointer = 11

sql = "SELECT*FROM sonota ORDER BY sonotaid"
If FcSQlGet(rs, sql, PrMsg) = False Then
    frmsonota.SSFTOKUSHU.Visible = False
    Screen.MousePointer = 0
Else
    frmsonota.CBOREN.Clear
    frmsonota.SSFTOKUSHU.Visible = True
    sonotasuu = rs.RecordCount
    rs.MoveFirst
    With frmsonota.sonotaichigrid
        .Rows = sonotasuu + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 5000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
        .ColWidth(9) = 1000
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "その他の契約名"
        .Col = 2
        .Text = "金額"
        .Col = 3
        .Text = "固定"
        .Col = 4
        .Text = "分類"
         .Col = 5
        .Text = "集計"
        .Col = 6
        .Text = "Group"
        .Col = 7
        .Text = "Sort"
         .Col = 8
        .Text = "表記(文章)"
         .Col = 9
        .Text = "契約期間修正"
    End With
    For i = 1 To sonotasuu
        With frmsonota.sonotaichigrid
            .Row = i
            .Col = 0
            .Text = rs!sonotaid
            .Col = 1
            .Text = rs!sonota
            .Col = 2
            If IsNull(rs!kingaku) Then
            .Text = ""
            Else
            .Text = rs!kingaku
            End If
            .Col = 3
            If rs!kotei = "0" Then
                .Text = "○"
            Else
                .Text = "×"
            End If
            .Col = 4
            If rs!kankei = "0" Then
                .Text = "ＣＡＴＶ"
            ElseIf rs!kankei = "1" Then
                .Text = "ＩＮＴＥＲ"
            Else
                .Text = "その他"
            End If
             .Col = 5
             If IsNull(rs!iru) Then
            .Text = ""
            Else
            .Text = rs!iru
            End If
              .Col = 6
            .CellAlignment = 1
            If IsNull(rs!gp) Then
            .Text = ""
            Else
            .Text = rs!gp
            End If
              .Col = 7
            .CellAlignment = 1
             If IsNull(rs!narabe) Then
            .Text = ""
            Else
            .Text = rs!narabe
            End If
              .Col = 8
            
             If IsNull(rs!hyoukimei) Then
            .Text = ""
            Else
            .Text = rs!hyoukimei
            End If
            
              .Col = 9
            
             If IsNull(rs!kasan) Then
            .Text = ""
            Else
            .Text = rs!kasan
            End If
            
            frmsonota.CBOREN.AddItem rs!sonotaid
        End With
       rs.MoveNext
    Next i
    rs.Close
End If
    Screen.MousePointer = 0
    


End Sub
Sub bankset()

'種類の一覧を表示します。

Dim banksuu As Integer

Screen.MousePointer = 11

'sql = "SELECT*FROM bank ORDER BY bankid"
sql = "SELECT bank.bankid, bankmain.bankname, bank.bankmei, bank.bankcode," & _
    "bank.shitencode" & _
    " FROM bank LEFT JOIN bankmain ON bank.bankcode = bankmain.bankcode" & _
    " ORDER BY bank.bankid;"
If FcSQlGet(rs, sql, PrMsg) = False Then
    frmbank.SSFTOKUSHU.Visible = False
    Screen.MousePointer = 0
Else
    frmbank.CBOREN.Clear
    frmbank.SSFTOKUSHU.Visible = True
    banksuu = rs.RecordCount
    rs.MoveFirst
    With frmbank.bankgrd
        .Rows = banksuu + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 5000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "金融機関名"
        .Col = 2
        .Text = "機関コード"
        .Col = 3
        .Text = "支店コード"
    End With
    For i = 1 To banksuu
        With frmbank.bankgrd
            .Row = i
            .Col = 0
            .Text = rs!bankid
            .Col = 1
            .Text = rs!bankname & Space(1) & rs!bankmei
            .Col = 2
            .Text = rs!bankcode
            .Col = 3
            .Text = rs!shitencode
            frmbank.CBOREN.AddItem rs!bankid
        End With
       rs.MoveNext
    Next i
    rs.Close
End If
    Screen.MousePointer = 0
    


End Sub

Sub cellset()
'種類の一覧を表示します。

Dim banksuu As Integer

Screen.MousePointer = 11

sql = "SELECT*FROM cell ORDER BY cellid"
If FcSQlGet(rs, sql, PrMsg) = False Then
    frmcell2.SSFTOKUSHU.Visible = False
    Screen.MousePointer = 0
Else
    frmcell2.CBOREN.Clear
    frmcell2.SSFTOKUSHU.Visible = True
    banksuu = rs.RecordCount
    rs.MoveFirst
    With frmcell2.cellgrid
        .Rows = banksuu + 1
        .ColWidth(0) = 500
        .ColWidth(1) = 2000
        .Row = 0
        .Col = 0
        .Text = "セルＩＤ"
        .Col = 1
        .Text = "セル名"
    End With
    For i = 1 To banksuu
        With frmcell2.cellgrid
            .Row = i
            .Col = 0
            .Text = rs!cellid
            .Col = 1
            .Text = rs!cellname
            frmcell2.CBOREN.AddItem rs!cellid
        End With
       rs.MoveNext
    Next i
    rs.Close
End If
    Screen.MousePointer = 0
    

End Sub

Sub nhk_zougaku()

                Dim sql_nhksakusei As String, rs_nhksakusei As Recordset, shiteinhkbi As String
                Dim snhkid As String, skojinid As String, skingaku As String, newkingaku As String
                Dim sagaku As Double, jikkou As Integer, hikiotoshi As String, hikiotoshi2 As String
                Dim sql_shuusei As String, rs_shuusei2 As Recordset, sql_new As String
                Dim motosono As Double, motoseikyuu As Double, motosonoshoukei As Double
                Dim sakisono As Double, sakiseikyuu As Double, sakisonoshoukei As Double
                
                If kidou4 = False Then
                            MsgBox "ＮＨＫデータＯＰＥＮに失敗しました。"
                            Exit Sub
                End If
                
                
                        sql_nhksakusei = "select * from nhk_m"
                        If FcSQlGet4(rs_nhksakusei, sql_nhksakusei, PrMsg) > 0 Then
                            rs_nhksakusei.MoveFirst
                            Do Until rs_nhksakusei.EOF
                                
                                jikkou = 0
                                snhkid = rs_nhksakusei!nhkid
                                skojinid = rs_nhksakusei!kojinid
                                skingaku = rs_nhksakusei!kingaku
                                hikiotoshi = rs_nhksakusei!n2
                                
                               Select Case Trim(rs_nhksakusei!kingaku)
                               Case "4040"
                                    newkingaku = "4160"
                                    sagaku = 120
                               Case "11460"
                                    newkingaku = "11815"
                                    sagaku = 355
                               Case "22250"
                                    newkingaku = "22920"
                                    sagaku = 670
                               Case "9925"
                                    newkingaku = "10260"
                                    sagaku = 335
                               Case "1820"
                                    newkingaku = "1880"
                                    sagaku = 60
                               Case "5130"
                                    newkingaku = "5307"
                                    sagaku = 177
                               Case "9925", "4160", "11815", "22920", "10260", "1880", "5307"
                                    jikkou = 1
                               Case Else
                                    MsgBox "不正な金額が登録されています。NHKID:" & snhkid
                                    Exit Sub
                               End Select
                               
                               
                               If jikkou = 0 Then
                                     hikiotoshi2 = "0"
                                     sql_shuusei = "select * from kojin where kojinid ='" & skojinid & "'"
                                     If FcSQlGet(rs_shuusei2, sql_shuusei, PrMsg) = False Then
                                         MsgBox "個人データが見つかりません。NHKID:" & snhkid
                                         Exit Sub
                                     Else
                                         If Trim(rs_shuusei2!sonotakin) <> "" Then
                                             motosono = CDbl(rs_shuusei2!sonotakin)
                                             motoseikyuu = CDbl(rs_shuusei2!soukei)
                                             motosonoshoukei = CDbl(rs_shuusei2!jikaita)
                                         Else
                                             motosono = 0
                                             motoseikyuu = CDbl(rs_shuusei2!soukei)
                                             motosonoshoukei = CDbl(rs_shuusei2!soukei)
                                         End If
                                         
                                       
                                            
                                                If hikiotoshi = "1" Then
                                                      sakisono = motosono + sagaku
                                                      sakiseikyuu = motoseikyuu + sagaku
                                                      sakisonoshoukei = motosonoshoukei + sagaku
                                                      If sakisonoshoukei < 0 Then
                                                          MsgBox "個人データのその他小計額が不正です。KOJINID:" & skojinid
                                                          Exit Sub
                                                      End If
                                                      If sakisono < 0 Then
                                                          MsgBox "個人データのその他請求額が不正です。KOJINID:" & skojinid
                                                          Exit Sub
                                                      End If
                                                      If sakiseikyuu < 0 Then
                                                          MsgBox "個人データの総請求額が不正です。KOJINID:" & skojinid
                                                          Exit Sub
                                                      End If
                                                Else
                                                      sakisono = motosono
                                                      sakiseikyuu = motoseikyuu
                                                      sakisonoshoukei = motosonoshoukei + sagaku
                                                      If sakisonoshoukei < 0 Then
                                                          MsgBox "個人データのその他小計額が不正です。KOJINID:" & skojinid
                                                          Exit Sub
                                                      End If
                                                End If
                                           
                                         
                                     End If
                                     If hikiotoshi2 = "0" Then
                                            sql_new = "update kojin set sonotakin = '" & CStr(sakisono) & "'," & _
                                                                " jikaita = '" & CStr(sakisonoshoukei) & "'," & _
                                                                " soukei = '" & CStr(sakiseikyuu) & "'" & _
                                                                " where kojinid ='" & skojinid & "'"
                                                                
                                            On Error GoTo errnhksakusei10
                                                db.Execute sql_new, dbSQLPassThrough
                                            On Error GoTo 0
                                    End If
                                         
                                 
                                    sql_nhksakusei = "update nhk_m set kingaku = '" & newkingaku & "'" & _
                                                " where nhkid ='" & snhkid & "'"
                                    On Error GoTo errnhksakusei10
                                        db4.Execute sql_nhksakusei, dbSQLPassThrough
                                    On Error GoTo 0
                                
                                
                                End If
                                
                                rs_nhksakusei.MoveNext
                            Loop
                            MsgBox "修正しました。"
                        Else
                            MsgBox "契約がありませんでした。"
                        End If
Exit Sub

errnhksakusei10:
    MsgBox "データの修正に失敗しました。"
    Exit Sub




End Sub
Function nhk_gengaku(s_mae1 As Long, s_mae2 As Long, s_saki1 As Long, s_saki2 As Long, s_sa1 As Long, s_sa2 As Long) As Integer

nhk_gengaku = 0

                Dim sql_nhksakusei As String, rs_nhksakusei As Recordset, shiteinhkbi As String
                Dim snhkid As String, skojinid As String, skingaku As String, newkingaku As String
                Dim sagaku As Long, jikkou As Integer, hikiotoshi As String, hikiotoshi2 As String
                Dim sql_shuusei As String, rs_shuusei2 As Recordset, sql_new As String
                Dim motosono As Long, motoseikyuu As Long, motosonoshoukei As Long
                Dim sakisono As Long, sakiseikyuu As Long, sakisonoshoukei As Long
                
                If kidou4 = False Then
                            MsgBox "ＮＨＫデータＯＰＥＮに失敗しました。"
                            Exit Function
                End If
                
                
                        sql_nhksakusei = "select * from nhk_m"
                        If FcSQlGet4(rs_nhksakusei, sql_nhksakusei, PrMsg) > 0 Then
                            rs_nhksakusei.MoveFirst
                            Do Until rs_nhksakusei.EOF
                                
                                jikkou = 0
                                snhkid = rs_nhksakusei!nhkid
                                skojinid = rs_nhksakusei!kojinid
                                skingaku = rs_nhksakusei!kingaku
                                hikiotoshi = rs_nhksakusei!n2
                                
                               Select Case CLng(Trim(rs_nhksakusei!kingaku))
                               Case s_mae1
                                    newkingaku = CStr(s_saki1)
                                    sagaku = s_sa1
                               Case s_mae2
                                    newkingaku = CStr(s_saki2)
                                    sagaku = s_sa2
                               Case Else
                                    jikkou = 1
                               End Select
                               
                               
                               If jikkou = 0 Then
                                     hikiotoshi2 = "0"
                                     sql_shuusei = "select * from kojin where kojinid ='" & skojinid & "'"
                                     If FcSQlGet(rs_shuusei2, sql_shuusei, PrMsg) = False Then
                                         MsgBox "個人データが見つかりません。NHKID:" & snhkid
                                         Exit Function
                                     Else
                                         If Trim(rs_shuusei2!f4) <> "" Then
                                             motosono = CDbl(rs_shuusei2!f4)
                                             motoseikyuu = CDbl(rs_shuusei2!soukei)
                                             'motosonoshoukei = CDbl(rs_shuusei2!jikaita)
                                         Else
                                             motosono = 0
                                             motoseikyuu = CDbl(rs_shuusei2!soukei)
                                             'motosonoshoukei = CDbl(rs_shuusei2!soukei)
                                         End If
                                         
                                         If skojinid = "" Then
                                            hikiotoshi2 = "1"
                                         Else
                                                If hikiotoshi = "1" Then
                                                      sakisono = motosono + sagaku
                                                      sakiseikyuu = motoseikyuu + sagaku
                                                      'sakisonoshoukei = motosonoshoukei - sagaku
                                                      'If sakisonoshoukei < 0 Then
                                                         ' MsgBox "個人データのその他小計額が不正です。KOJINID:" & skojinid
                                                         ' Exit Sub
                                                      'End If
                                                      If sakisono < 0 Then
                                                          MsgBox "個人データのNHK額が不正です。KOJINID:" & skojinid
                                                          Exit Function
                                                      End If
                                                      If sakiseikyuu < 0 Then
                                                          MsgBox "個人データの総請求額が不正です。KOJINID:" & skojinid
                                                          Exit Function
                                                      End If
                                                Else
                                                      sakisono = motosono
                                                      sakiseikyuu = motoseikyuu
                                                     ' sakisonoshoukei = motosonoshoukei - sagaku
                                                      'If sakisonoshoukei < 0 Then
                                                         ' MsgBox "個人データのその他小計額が不正です。KOJINID:" & skojinid
                                                         ' Exit Sub
                                                      'End If
                                                End If
                                          End If
                                     End If
                                     If hikiotoshi2 = "0" Then
                                            sql_new = "update kojin set f4 = '" & CStr(sakisono) & "'," & _
                                                                " soukei = '" & CStr(sakiseikyuu) & "'" & _
                                                                " where kojinid ='" & skojinid & "'"
                                                                
                                            On Error GoTo errnhksakusei10
                                                db.Execute sql_new, dbSQLPassThrough
                                            On Error GoTo 0
                                    Else
                                         MsgBox "kojinIDが不正です。NHKID:" & snhkid
                                    End If
                                         
                                 
                                    sql_nhksakusei = "update nhk_m set kingaku = '" & newkingaku & "'" & _
                                                " where nhkid ='" & snhkid & "'"
                                    On Error GoTo errnhksakusei10
                                        db4.Execute sql_nhksakusei, dbSQLPassThrough
                                    On Error GoTo 0
                                
                                
                                End If
                                
                                rs_nhksakusei.MoveNext
                            Loop
                            'MsgBox "修正しました。"
                            nhk_gengaku = 1
                            
                        Else
                            'MsgBox "契約がありませんでした。"
                        End If
Exit Function

errnhksakusei10:
    MsgBox "データの修正に失敗しました。"
    Exit Function




End Function

Sub Main()

s_kidou_kojinid = ""

Dim kidouno As String

 Dim cmdline() As String
 
    cmdline = Split(Command(), " ") 'スペースで切り分けてコマンドライン引数を配列に格納
    
    Dim ArCount As Integer
    ArCount = UBound(cmdline) '配列サイズを求める
    
    Dim i As Integer
    For i = 0 To ArCount
        'MsgBox (cmdline(i)) '表示
          kidouno = Mid(cmdline(i), 1, 1)
           Select Case kidouno
            Case "k"
                If Mid(cmdline(i), 2, 5) <> "" Then
                    s_kidou_kojinid = Mid(cmdline(i), 2, 5)
                Else
                    s_kidou_kojinid = "00000"
                End If
           End Select
    Next i

'バージョン情報

Dim s_ver_st As String
s_ver_st = GetWindowsVersionInfo()
'MsgBox (s_ver_st)

Dim s_bild As String, n_bild As Long
s_bild = Mid(s_ver_st, 1, 5)
n_bild = CLng(s_bild)

If n_bild > 28500 Then
    Dim s_ver_msg As String
    s_ver_msg = "問題が発生しました。" & Chr(13) & Chr(13) & "このエラーについて、システム管理者またはサポートに問い合わせてください。エラーコードは、0x800f081f　です。"
    MsgBox (s_ver_msg)
End If



Dim zeorobi As String
mapa = App.Path
If Right(mapa, 1) <> "\" Then
    mapa = mapa & "\"
End If
    
'Mapパス
map_exe_path = mapa & "ezmanager_map.exe"

'FTPパス
ftp_exe_path = mapa & "ezmanager_ftp.exe"

'Subパス
sub_exe_path = mapa & "ezmanager_sub.exe"

'sagyou log path
Dim DesktopPath As String
' デスクトップのパスを取得
DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
sagyou_log_path = DesktopPath & "\EZManager_data\log\"

    
'設定ファイルのみのチェック
'MAINDB.MDBが存在するかの確認
    vbname3 = mapa & "settei.mdb"
If Dir$(vbname3) = vbNullString Then
   MsgBox "設定ファイルがありません。"
   End
End If


If s_kidou_kojinid = "" Then
    kenchan = 0
    Screen.MousePointer = 0
    FRMPASSWD.Show 1
    Screen.MousePointer = 11
    
    
    frmrogo.lblshinkou.Caption = "起動準備中・・"
    frmrogo.Show
    DoEvents

End If


settei_res = setting(1, 7, 0, "", 1)
If settei_res = "-1" Then
    Ret = MsgBox("設定の読み出しに失敗しました。001", 16, "EZ Manager")
    db3.Close
    End
ElseIf settei_res = "0" Then
    kiidou_mode = 0
ElseIf settei_res = "1" Then
    kiidou_mode = 1
ElseIf settei_res = "2" Then
    kiidou_mode = 2
End If

settei_res = setting(1, 4, 0, "", 1)
If settei_res = "-1" Then
    Ret = MsgBox("設定の読み出しに失敗しました。004", 16, "EZ Manager")
    db3.Close
    End
ElseIf settei_res = "0" Then
    Ret = MsgBox("データデースの設定がありません。", 16, "EZ Manager")
    db3.Close
    End
Else
        sopa = settei_res & "\"
        
     
        
End If


'PC名＋バージョン情報を記録
Dim s_pcname As String, s_appver As String
s_pcname = Environ("COMPUTERNAME")
s_appver = App.Major & "." & App.Minor & "." & App.Revision

err_write2 s_pcname & "," & s_appver, sopa & "version_jouhou.txt"



Screen.MousePointer = 11

If s_kidou_kojinid = "" Then

    frmrogo.lblshinkou.Caption = "新バージョンを確認中・・"
    DoEvents
    
    
    'バージョン情報チェック
    kakunin_ver
    
    
    '自動バージョンアップ起動
    
    Dim RetVal, mentefaile As String, fc As Object
    
    'mentefaile = mapa & "ezmanager_ment.exe"
    'If Dir(mentefaile) = "" Then
        'MsgBox "メンテナンス用プログラムがありません。"
    'Else
        'mentefaile = mentefaile & " -i"
        'RetVal = Shell(mentefaile, 0)   ' 電卓を実行します。
    'End If

End If



'New 20200801

On Error GoTo eff

'ファイルサーバ上のファイルを、自分のMyDocにコピー
Dim s_server_file As String
Dim s_server_file2 As String

s_server_file = "\\k-catv-server\data\共有データ\ezmanager" & "\catv.exe"
s_server_file2 = "\\k-catv-server\data\共有データ\ezmanager" & "\ezmanager_sub.exe"

Dim s_mydoc As String
Dim s_mydoc2 As String

Dim objwsh As Object
Set objwsh = CreateObject("Wscript.Shell")
s_mydoc = objwsh.SpecialFolders("MyDocuments")
s_mydoc2 = objwsh.SpecialFolders("MyDocuments")
s_mydoc = s_mydoc & "\catv.exe"
s_mydoc2 = s_mydoc2 & "\ezmanager_sub.exe"

'FileSystemObjectインスタンスを生成
Dim fi As New FileSystemObject


'Log

 sagyou_msg = "catv.exe updata:" & s_server_file & "/" & s_mydoc
 log_sagyou sagyou_log_path, sagyou_msg
        
        
'ファイルの存在確認
If fi.FileExists(s_server_file) Then
   ' Call MsgBox("ファイルは存在します")
   
   If Dir(s_server_file) <> vbNullString Then
        If Dir(s_mydoc) <> vbNullString Then
            Kill s_mydoc
            DoEvents
        End If
        
        FileCopy s_server_file, s_mydoc
        
         sagyou_msg = "catv.exe updata:OK  " & s_server_file & "/" & s_mydoc
        log_sagyou sagyou_log_path, sagyou_msg

    End If
Else
   ' Call MsgBox("ファイルは存在しません")
End If

'オブジェクトの解放
Set fi = Nothing

Dim fi2 As New FileSystemObject

 sagyou_msg = "ezmanager_sub.exe updata:" & s_server_file2 & "/" & s_mydoc2
 log_sagyou sagyou_log_path, sagyou_msg
        

If fi2.FileExists(s_server_file2) Then
   ' Call MsgBox("ファイルは存在します")
   
   If Dir(s_server_file2) <> vbNullString Then
        If Dir(s_mydoc2) <> vbNullString Then
            Kill s_mydoc2
            DoEvents
        End If
        
        FileCopy s_server_file2, s_mydoc2
        
         sagyou_msg = "ezmanager_sub.exe updata:OK  " & s_server_file2 & "/" & s_mydoc2
        log_sagyou sagyou_log_path, sagyou_msg

    End If
Else
   ' Call MsgBox("ファイルは存在しません")
End If

'オブジェクトの解放
Set fi2 = Nothing

On Error GoTo 0

'kenchan = 0
'Screen.MousePointer = 0
'FRMPASSWD.Show 1
'Screen.MousePointer = 11

On Error GoTo errfile2

        
        'MAINDB.MDBが存在するかの確認
            VBNAME = sopa & "catv.mdb"
        If Dir(VBNAME) = vbNullString Or Dir(VBNAME) = "" Then
            
                MsgBox "「catv.mdb」　がありません。ファイルサーバが稼動していないか、ネットワークに接続していない可能性があります。稼動と接続を確認してから再度実行してください。"
                 End
           
        End If
        On Error GoTo 0
        On Error GoTo errfile
        'MAINDB.MDBが存在するかの確認
            VBNAME2 = sopa & "koushin.mdb"
        If Dir$(VBNAME2) = vbNullString Then
           MsgBox "「koushin.mdb」　がありません。ファイルサーバが稼動していないか、ネットワークに接続していない可能性があります。稼動と接続を確認してから再度実行してください。"
           End
        End If
        
        'MAINDB.MDBが存在するかの確認
            vbname4 = sopa & "nhk.mdb"
        If Dir$(vbname4) = vbNullString Then
           MsgBox "「nhk.mdb」　がありません。ファイルサーバが稼動していないか、ネットワークに接続していない可能性があります。稼動と接続を確認してから再度実行してください。"
           End
        End If
        'vbname_voipが存在するかの確認
            vbname_voip = sopa & "voip.mdb"
        If Dir$(vbname_voip) = vbNullString Then
           MsgBox "「voip.mdb」　がありません。ファイルサーバが稼動していないか、ネットワークに接続していない可能性があります。稼動と接続を確認してから再度実行してください。"
           End
        End If
        'vbname_ppvが存在するかの確認
            vbname_ppv = sopa & "ppv.mdb"
        If Dir$(vbname_ppv) = vbNullString Then
           MsgBox "「ppv.mdb」　がありません。ファイルサーバが稼動していないか、ネットワークに接続していない可能性があります。稼動と接続を確認してから再度実行してください。"
           End
        End If
        
        'vbname_ppvが存在するかの確認
            vbname_cl = sopa & "cl.mdb"
        If Dir$(vbname_cl) = vbNullString Then
           MsgBox "「cl.mdb」　がありません。ファイルサーバが稼動していないか、ネットワークに接続していない可能性があります。稼動と接続を確認してから再度実行してください。"
           End
        End If



On Error GoTo 0

kidou_main



'FD
    settei_res = setting(1, 8, 0, "", 1)
    If settei_res = "-1" Then
        Ret = MsgBox("設定の読み込みに失敗しました。", 16, "EZ Manager")
        fd_a = "a:\"
    ElseIf settei_res = "-1" Then
        fd_a = "a:\"
    Else
         fd_a = settei_res
         If Right(fd_a, 1) <> "\" Then
            fd_a = fd_a & "\"
        End If
    End If
    
    
'滞納
    '滞納ST
    s_tainou_tsuika_kingaku = 0
   If soukatsu(0, "startup", 17, "", 1, 0) = False Then
            MsgBox "滞納設定の読み込みに失敗しました。"
    End If
    If reg = "1" Then
        'chktainou.Value = 1
        If soukatsu(0, "startup", 18, "", 0, 1) = False Then
                MsgBox "滞納設定の読み込みに失敗しました。"
        End If
        If reg = "" Then
            reg = "0"
        End If
        
        s_tainou_tsuika_kingaku = CLng(reg)
    End If
    

    
    

If s_kidou_kojinid = "" Then
    
    If kiidou_mode = 0 Then
        
        frmrogo.lblshinkou.Caption = "自動０更新を確認中・・"
        DoEvents
    
    
        '請求書・現金自動更新
        Dim kyounohi As String, chknohi As String, chkgo As Integer
        
        chkgo = 0
        kyounohi = Format(Date, "mmdd")
        
        If soukatsu(0, "startup", 1, "", 1, 1) = False Then
                MsgBox "設定の読み込みに失敗しました。"
            db.Close
            End
        End If
            
        
        If reg = "0" Then
            '15日以前の時
            If Format(Date, "dd") < 20 Then
                chknohi = Format(Date, "mm") & "20"
            Else '15日以降の時
                If CLng(Format(Date, "mm")) + 1 > 12 Then
                    chknohi = "0120"
                Else
                    chknohi = Format(Format(Date, "mm") + 1 & "20", "0000")
                End If
            End If
            If soukatsu(1, "startup", 1, chknohi, 1, 1) = False Then
                    MsgBox "設定の書き込みに失敗しました。"
                db.Close
                End
            End If
        ElseIf reg = "1220" Then
            chknohi = reg
            If Left(chknohi, 2) <> Left(kyounohi, 2) Then
                If Left(kyounohi, 2) = "01" Then
                    chkgo = 1
                End If
            Else
                If kyounohi > "1220" Then
                    chkgo = 1
                End If
            End If
        ElseIf reg = "0120" Then
            chknohi = reg
                If Mid(kyounohi, 1, 2) = "12" Then
                    chkgo = 0
                Else
                    If chknohi <= kyounohi Then
                        chkgo = 1
                    End If
                End If
        Else
            
            chknohi = reg
                If chknohi <= kyounohi Then
                    chkgo = 1
                End If
        End If
        fddkoushinOK = chkgo
        If chkgo = 1 Then
            BeginTrans
            msg = "口座引落（請求額が０円以下）と請求書扱いの内容を更新します。"
            res = MsgBox(msg, vbYesNo)
            If res = vbYes Then
                msg = "再確認！！更新をキャンセルしますか？"
                res = MsgBox(msg, vbYesNo)
                If res = vbYes Then
                    MsgBox "中止しました。次回起動時に再度更新処理を行います。"
                Else
                    res = InputBox("０更新を実行します。起動時パスワードを入力してください。")
                    
                        frmrogo.lblshinkou.Caption = "自動０更新を実行中・・"
                        DoEvents
    
                        If soukatsu(0, "startup", 2, "", 1, 1) = False Then
                                MsgBox "設定の読み込みに失敗しました。キャンセルしました。"
                        Else
                                If reg = "0" Then
                                    reg = "起動時パスワードが設定されていません。"
                                Else
            
                                    If res = reg Then
                                            If a_jidoukoushin = False Then
                                                Rollback
                                                MsgBox "口座引落（請求額が０円以下）と請求書扱いの自動更新に失敗しました。"
                                                End
                                            
                                            Else
                                                zeorobi = Format(Date, "yyyymmdd")
                                                If soukatsu(1, "startup", 12, zeorobi, 1, 1) = False Then
                                                        Rollback
                                                        MsgBox "設定の書き込みに失敗しました。"
                                                    db.Close
                                                    End
                                                End If
                                                If chknohi = "1220" Then
                                                    chknohi = "0120"
                                                Else
                                                    chknohi = Format(Mid(chknohi, 1, 2) + 1 & "20", "0000")
                                                End If
                                                If soukatsu(1, "startup", 1, chknohi, 1, 1) = False Then
                                                        Rollback
                                                        MsgBox "設定の書き込みに失敗しました。"
                                                    db.Close
                                                    End
                                                Else
                                                    fddkoushinOK = 0
                                                    
                                                            'Log
                                                    sagyou_msg = "０更新：０更新の処理が完了しました。"
                                                    log_sagyou sagyou_log_path, sagyou_msg
                                                    
                                                    
                                                End If
                                            End If
                                    Else
                                        MsgBox "パスワードが違います。キャンセルしました。"
                                    End If
                                End If
                        End If
                End If
            Else
                MsgBox "中止しました。次回起動時に再度更新処理を行います。"
            End If
        CommitTrans
        End If
    End If
    


    
    'NHK金額
    
    
        
          If soukatsu(0, "startup", 8, "", 1, 0) = False Then
                MsgBox "設定の読み込みに失敗しました。何度も表示される場合は、管理者に連絡してください。"
                nhk_ryoukin_tsuujou = 0
        End If
        If reg = "0" Then
            'reg = ""
             MsgBox "起動後すぐに、管理者モードでログインし直し、NHK料金の設定を行ってください。"
            nhk_ryoukin_tsuujou = 0
        End If
        nhk_ryoukin_tsuujou = CLng(reg)
        
            If soukatsu(0, "startup", 13, "", 1, 0) = False Then
                MsgBox "設定の読み込みに失敗しました。何度も表示される場合は、管理者に連絡してください。"
                 nhk_ryoukin_hangaku = 0
        End If
        If reg = "0" Then
            'reg = ""
            MsgBox "起動後すぐに、管理者モードでログインし直し、NHK料金の設定を行ってください。"
            nhk_ryoukin_hangaku = 0
        End If
        nhk_ryoukin_hangaku = CLng(reg)

End If


frmmain.Show
Unload frmrogo
Set frmrogo = Nothing
furiganahenkou2 = 0
Screen.MousePointer = 0


If s_kidou_kojinid <> "" Then


If s_kidou_kojinid <> "00000" Then
        If hyouji_touroku(s_kidou_kojinid) = False Then
              Exit Sub
          End If
          If nhk_set(s_kidou_kojinid) = False Then
              Screen.MousePointer = 0
              Exit Sub
          End If
          
          keiyakuset (s_kidou_kojinid)
          keiyakuset2 (s_kidou_kojinid)
          chk_kyuushi (s_kidou_kojinid)
          chk_log (s_kidou_kojinid)
          
          frmtouroku2.Show 1
    
    End If
End If


Exit Sub
MsgBox "コピー時エラー。"
eff:

Resume Next

errfile2:
Resume Next


errfile:
MsgBox "データベースファイルがありません。再度確認してから実行してください。"
End


End Sub
Function setting(hajime As Integer, id As Long, docchi As Integer, newid As String, owari As Integer) As String
Dim sql_set As String, rs_set As Recordset, sql_setstr As String, wMsg As String



If hajime = 1 Then
    'データのロード
    kidoupassword = ";pwd=" & "kamifusafusa"
    'Set DB = OpenDatabase(VBNAME)
    On Error GoTo errsettingkidou
        Set db3 = DBEngine.Workspaces(0).OpenDatabase(vbname3, False, False, kidoupassword)
    On Error GoTo 0
End If

If docchi = 0 Then '読み込み
    sql_set = "select * from settei where id ='1'"
    If FcSQlGet3(rs_set, sql_set, wMsg) = False Then
        setting = "-1"
        Exit Function
    Else
        Select Case id
        Case 1 '起動１
            setting = CStr(rs_set!s1)
        Case 2 '起動２
            setting = CStr(rs_set!s2)
        Case 3 '起動１
            setting = CStr(rs_set!s3)
        Case 4 '起動１パス
            setting = CStr(rs_set!s4)
        Case 5 '起動２パス
            setting = CStr(rs_set!s5)
        Case 6 '起動３パス
            setting = CStr(rs_set!s6)
        Case 7
            setting = CStr(rs_set!s7)
        Case 8
            setting = CStr(rs_set!s8)
        Case 9
            setting = CStr(rs_set!s9)
        Case 10
            setting = CStr(rs_set!s10)
        Case 11
            setting = CStr(rs_set!s11)
        Case 12
            setting = CStr(rs_set!s12)
        Case 13
            setting = CStr(rs_set!s13)
        Case 14
            setting = CStr(rs_set!s14)
        Case 15
            setting = CStr(rs_set!s15)
        Case 16
            setting = CStr(rs_set!s16)
        Case 17
            setting = CStr(rs_set!s17)
        Case 18
            setting = CStr(rs_set!s18)
        Case 19
            setting = CStr(rs_set!s19)
        Case 20 '履歴の数
            setting = CStr(rs_set!s20)
        End Select
        rs_set.Close
    End If
    
Else '書き込み
    sql_setstr = "s" & CStr(id)
    sql_set = "UPDATE settei SET " & sql_setstr & " = '" & newid & "'" & _
                          " WHERE id = '1'"
    On Error GoTo errrirekikaki2
                 db3.Execute sql_set, dbSQLPassThrough
    On Error GoTo 0
    
End If

If owari = 1 Then
    db3.Close
End If

Exit Function

errsettingkidou:
    setting = "-1"
    Exit Function
errrirekikaki2:
    setting = "-1"
    Exit Function

End Function


Sub kakunin_ver()

    Dim lngRet As Long, lets_go_saikidou As Integer
    Dim lngDummy As Long
    Dim bBuffer() As Byte
    Dim lngLen As Long
    Dim lpBuffer As Long
    Dim ffi As VS_FIXEDFILEINFO
    Dim strFileName(4) As String
    Dim strFileName2(4) As String
    Dim i As Integer
    Dim vernokanri(4) As String
    Dim vernokanri2(4) As String
    Dim settei_res3 As String
    Dim mydoc_path As String
Dim root_pass As String

Dim objwsh As Object
Set objwsh = CreateObject("Wscript.Shell")
root_pass = objwsh.SpecialFolders("MyDocuments")

If Right(root_pass, 1) <> "\" Then
   root_pass = root_pass & "\"
End If
Set objwsh = Nothing
    
    
On Error GoTo errf
    
    mydoc_path = root_pass
    
    
    lets_go_saikidou = 0
    ' strFileName に取得したいファイル名をセット
    strFileName(0) = mydoc_path & "catv.exe"
    strFileName(1) = mydoc_path & "ezmanager_ment.exe"
     strFileName(2) = mydoc_path & "ezmanager_sub.exe"
    
    strFileName2(0) = App.Path & "\catv.exe"
    strFileName2(1) = App.Path & "\ezmanager_ment.exe"
    strFileName2(2) = App.Path & "\ezmanager_sub.exe"
    
    For i = 0 To 2
        '管理フォルダーのファイルのﾊﾞｰｼﾞｮﾝ確認（Settei.mdb内の）
        
            'サーバのバージョン
            lngLen = GetFileVersionInfoSize(strFileName(i), lngDummy)
            If lngLen < 1 Then
                'MsgBox "サーバの最新版のバージョンを読み込めませんでした。"
                vernokanri(i) = "00000"
            Else
                ' バイトの配列の領域取得
                ReDim bBuffer(lngLen)
                If Dir(strFileName(i)) <> "" Then
                    ' ファイル バージョン情報を取得
                    lngRet = GetFileVersionInfo(strFileName(i), 0&, lngLen, bBuffer(0))
                    lngRet = VerQueryValue(bBuffer(0), "\", lpBuffer, lngLen)
                    
                    ' バイトの処理
                    MoveMemory ffi, lpBuffer, Len(ffi)
                    
                    ' ファイル バージョン
                        vernokanri(i) = Format$(ffi.dwFileVersionMSh) & _
                        Format$(ffi.dwFileVersionMSl, "00") & _
                        Format$(ffi.dwFileVersionLSl, "00")
                End If
            End If
            
            '実行しているPCのバージョン
            lngLen = GetFileVersionInfoSize(strFileName2(i), lngDummy)
            If lngLen < 1 Then
                MsgBox "現在実行しているバージョンを読み込めませんでした。"
                Exit Sub
            End If
            ' バイトの配列の領域取得
            ReDim bBuffer(lngLen)
            If Dir(strFileName2(i)) <> "" Then
                ' ファイル バージョン情報を取得
                lngRet = GetFileVersionInfo(strFileName2(i), 0&, lngLen, bBuffer(0))
                lngRet = VerQueryValue(bBuffer(0), "\", lpBuffer, lngLen)
                
                ' バイトの処理
                MoveMemory ffi, lpBuffer, Len(ffi)
                
                ' ファイル バージョン
                    vernokanri2(i) = Format$(ffi.dwFileVersionMSh) & _
                    Format$(ffi.dwFileVersionMSl, "00") & _
                    Format$(ffi.dwFileVersionLSl, "00")
            End If
            
            'MsgBox "元：" & vernokanri(i) & "、今：" & vernokanri2(i)
            '比較
            If vernokanri(i) > vernokanri2(i) Then
                Dim RetVal, resetfaile As String
                
                MsgBox "現在起動しているソフトよりも新しいバージョンがあります。バージョンアップします。"
                
                    resetfaile = App.Path & "\ezmanager_ver.exe"
                    If Dir(resetfaile) = "" Then
                        MsgBox "「EZ Manager」管理プログラムがありません。"
                        End
                    End If
                    RetVal = Shell(resetfaile, 1)   ' 電卓を実行します。
                    End
                
            End If
                
            
            
            
            
    Next i
    
    
 On Error GoTo 0
 
 
 Exit Sub
 
errf:
    
  MsgBox "Ver時エラー"
             
        

End Sub


Sub add_fld_200403()
Dim tb As String, fn As String, fcc As Long, ft As Integer, fs As Long
tb = "kojin"
For fcc = 0 To 37
    ft = 1
    fs = 8
    Select Case fcc
    Case 0
        fn = "kanyuubi"
    Case 1
        fn = "kanyuubi2"
    Case 2
        fn = "koujihi"
        fs = 7
    Case 3
        fn = "waribiki"
    Case 4
        fn = "riyuu"
        fs = 50
    Case 5
        fn = "nyuukinbi"
    Case 6
        fn = "nyuukingaku"
        fs = 7
    Case 7
        fn = "kanrennbi"
        fs = 6
    Case 8
        fn = "zankin"
        fs = 7
    Case 9
        fn = "qshiyou"
        fs = 1
    Case 10
        fn = "qhiduke"
        fs = 6
    Case 11
        fn = "karijusho"
        fs = 40
    Case 12
        fn = "v1"
        fs = 11
    Case 13
        fn = "v2"
        fs = 11
    Case 14
        fn = "v3"
        fs = 11
    Case 15
        fn = "v4"
        fs = 11
    Case 16
        fn = "v5"
        fs = 11
    Case 17
        fn = "v6"
        fs = 11
    Case 18
        fn = "v7"
        fs = 11
    Case 19
        fn = "v8"
        fs = 11
    Case 20
        fn = "v9"
        fs = 11
    Case 21
        fn = "v10"
        fs = 11
    Case 22
        fn = "vh1"
    Case 23
        fn = "vh2"
    Case 24
        fn = "vh3"
    Case 25
        fn = "vh4"
    Case 26
        fn = "vh5"
    Case 27
        fn = "vh6"
    Case 28
        fn = "vh7"
    Case 29
        fn = "vh8"
    Case 30
        fn = "vh9"
    Case 31
        fn = "vh10"
    Case 32
        fn = "kyukouza"
        fs = 50
    Case 33
        fn = "kyukouza2"
        fs = 50
    Case 34
        fn = "kanrikouza"
        fs = 50
    Case 35
        fn = "nhkno"
        fs = 10
    Case 36
        fn = "nhkname"
        fs = 60
    Case 37
        fn = "nhkfuri"
        fs = 45
        
    End Select
    If newtbl_make(tb, fn, ft, fs) = False Then
        MsgBox "err"
        Exit Sub
    End If
Next fcc


MsgBox "ok"


End Sub

Sub add_fld_200403_2()
Dim tb As String, fn As String, fcc As Long, ft As Integer, fs As Long
tb = "kojin"
For fcc = 0 To 2
    ft = 1
    fs = 8
    Select Case fcc
    Case 0
        fn = "koujihi2"
        fs = 7
    Case 1
        fn = "waribiki2"
    Case 2
        fn = "riyuu2"
        fs = 50
    End Select
    If newtbl_make(tb, fn, ft, fs) = False Then
        MsgBox "err"
        Exit Sub
    End If
Next fcc


MsgBox "ok"


End Sub


Function newtbl_make(teblename As String, fieldname As String, typename As Integer, Optional sizename As Long) As Boolean

Dim sql_skin As String, rs_skin As Recordset, typename2 As String
Dim sql_skin2 As String, rs_skin2 As Recordset
On Error GoTo mikefld
    If typename = 1 Then
        typename2 = "string"
    Else
        typename2 = "long"
    End If
    sql_skin = "alter table " & teblename & " add column " & fieldname & " " & typename2
    If Not IsMissing(sizename) Then
        sql_skin = sql_skin & "(" & sizename & ")"
    End If
    db.Execute sql_skin
    
    sql_skin2 = "select kojinid from kojin"
    If FcSQlGet(rs_skin2, sql_skin2, PrMsg) = True Then
        rs_skin2.MoveFirst
        Do Until rs_skin2.EOF
            sql_skin = "UPDATE " & teblename & " SET " & fieldname & " = '" & Space(1) & "'" & _
                        "where kojinid ='" & rs_skin2!kojinid & "'"
            db.Execute sql_skin, dbSQLPassThrough
            rs_skin2.MoveNext
        Loop
        rs_skin2.Close
    End If
newtbl_make = True
Exit Function

mikefld:
newtbl_make = False

End Function
Sub nyuukinrirekiset2(ididid As String, Optional nono As Integer)
Dim rirekirs As Recordset, rirekisuu As Long, datadata(), cccsuu2 As Long
Dim rirekirs2 As Recordset

cccsuu2 = 0
sql = "select * from shousai where kojinid ='" & ididid & "'" & _
    " and keist='0' and keishurui='1'" & _
    " order by keiseikyuu"
If FcSQlGet2(rirekirs, sql, PrMsg) = True Then
    
    rirekirs.MoveFirst
    Do Until rirekirs.EOF
        cccsuu2 = cccsuu2 + 1
        ReDim Preserve datadata(7, cccsuu2)
        datadata(0, cccsuu2) = rirekirs!shousaiid
        datadata(1, cccsuu2) = Format(rirekirs!keiseikyuu, "@@@@/@@")
        datadata(2, cccsuu2) = rirekirs!keiid
        datadata(3, cccsuu2) = rirekirs!keikin
        sql = "select shuruimei from shurui where shuruiid='" & rirekirs!keiid & "'"
        If FcSQlGet(rirekirs2, sql, PrMsg) = True Then
            datadata(4, cccsuu2) = rirekirs2!shuruimei
            rirekirs2.Close
        End If
        datadata(5, cccsuu2) = 1
        datadata(6, cccsuu2) = rirekirs!keiid
        rirekirs.MoveNext
    Loop
    rirekirs.Close

End If

sql = "select * from shousai where kojinid ='" & ididid & "'" & _
    " and keist='0' and keishurui='2'" & _
    " order by keiseikyuu"
If FcSQlGet2(rirekirs, sql, PrMsg) = True Then
    rirekirs.MoveFirst
    Do Until rirekirs.EOF
        cccsuu2 = cccsuu2 + 1
        ReDim Preserve datadata(7, cccsuu2)
        datadata(0, cccsuu2) = rirekirs!shousaiid
        datadata(1, cccsuu2) = Format(rirekirs!keiseikyuu, "@@@@/@@")
        datadata(2, cccsuu2) = rirekirs!keiid
        datadata(3, cccsuu2) = rirekirs!keikin
        sql = "select sonota from sonota where sonotaid='" & rirekirs!keiid & "'"
        If FcSQlGet(rirekirs2, sql, PrMsg) = True Then
            datadata(4, cccsuu2) = rirekirs2!sonota
            rirekirs2.Close
        End If
        datadata(5, cccsuu2) = 2
        datadata(6, cccsuu2) = rirekirs!keiid
        rirekirs.MoveNext
    Loop
    rirekirs.Close

End If

'NHKの契約
sql = "select * from shousai where kojinid ='" & ididid & "'" & _
    " and keist='0' and keishurui='3'" & _
    " order by keiseikyuu"
If FcSQlGet2(rirekirs, sql, PrMsg) = True Then
    rirekirs.MoveFirst
    Do Until rirekirs.EOF
        cccsuu2 = cccsuu2 + 1
        ReDim Preserve datadata(7, cccsuu2)
        datadata(0, cccsuu2) = rirekirs!shousaiid
        datadata(1, cccsuu2) = Format(rirekirs!keiseikyuu, "@@@@/@@")
        datadata(2, cccsuu2) = rirekirs!keiid
        datadata(3, cccsuu2) = rirekirs!keikin
        Select Case CInt(rirekirs!keiid)
        Case 0
            datadata(4, cccsuu2) = "ＮＨＫ団体一括１ヶ月"
            
        Case 7
            datadata(4, cccsuu2) = "ＮＨＫ団体一括６ヶ月家族割"
        Case 6
            datadata(4, cccsuu2) = "ＮＨＫ団体一括６ヶ月ハンデ"
        Case 1
            datadata(4, cccsuu2) = "ＮＨＫ団体一括２ヶ月"
        Case 2
            datadata(4, cccsuu2) = "ＮＨＫ団体一括３ヶ月"
        Case 3
            datadata(4, cccsuu2) = "ＮＨＫ団体一括４ヶ月"
        Case 4
            datadata(4, cccsuu2) = "ＮＨＫ団体一括５ヶ月"
        Case 5
            datadata(4, cccsuu2) = "ＮＨＫ団体一括６ヶ月"
        End Select
        datadata(5, cccsuu2) = 3
        datadata(6, cccsuu2) = rirekirs!keiid
        rirekirs.MoveNext
    Loop
    rirekirs.Close
End If

db2.Close



If nono = 1 Then
    frmsouken6.gridnyuukinshousai.Clear
    With frmsouken6.gridnyuukinshousai
        .Font.Size = 11
        .Rows = cccsuu2 + 1
        .Cols = 7
        .ColWidth(0) = 1200
        .ColWidth(1) = 1000
        .ColWidth(2) = 1300
        .ColWidth(3) = 3000
        .ColWidth(4) = 1200
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "請求日"
        .Col = 2
        .Text = "入金日"
        .Col = 3
        .Text = "契約種類"
        .Col = 4
        .Text = "金額"
    End With
    For i = 1 To cccsuu2
        With frmsouken6.gridnyuukinshousai
            .Row = i
            .Col = 0
                    '  .CellChecked = 2
                      .CellPictureAlignment = 1
                    '  .Editable = flexEDKbd
            .Text = datadata(0, i)
            .Col = 1
            .Text = datadata(1, i)
            .Col = 2
            '.Text = Format(hiduhidu, "@@@@/@@/@@")
            .Col = 3
            .Text = datadata(4, i)
            .Col = 4
            .Text = datadata(3, i)
            .Col = 5
            .Text = datadata(5, i)
            .Col = 6
            .Text = datadata(6, i)
        End With
    Next i

    frmsouken6.Show 1
Else
    frmsouken5.gridnyuukinshousai.Clear
    With frmsouken5.gridnyuukinshousai
        .Font.Size = 11
        .Rows = cccsuu2 + 1
        .Cols = 5
        .ColWidth(0) = 1200
        .ColWidth(1) = 1000
        .ColWidth(2) = 1300
        .ColWidth(3) = 3000
        .ColWidth(4) = 1200
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "請求日"
        .Col = 2
        .Text = "入金日"
        .Col = 3
        .Text = "契約種類"
        .Col = 4
        .Text = "金額"
    End With
    For i = 1 To cccsuu2
        With frmsouken5.gridnyuukinshousai
            .Row = i
            .Col = 0
            .Text = datadata(0, i)
            .Col = 1
            .Text = datadata(1, i)
            .Col = 2
            '.Text = Format(hiduhidu, "@@@@/@@/@@")
            .Col = 3
            .Text = datadata(4, i)
            .Col = 4
            .Text = datadata(3, i)
        End With
    Next i

    If kenchan = 1 Then
        frmsouken5.Picture5.Enabled = True
     Else
         frmsouken5.Picture5.Enabled = False
     End If

frmsouken5.Show 1
End If
End Sub
Sub nyuukinrirekiset3(ididid As String)
Dim rirekirs As Recordset, rirekisuu As Long, datadata(), cccsuu2 As Long
Dim rirekirs2 As Recordset

cccsuu2 = 0
sql = "select * from shousai where kojinid ='" & ididid & "'" & _
    " and keist='0' and keishurui='1'" & _
    " order by keiseikyuu"
If FcSQlGet2(rirekirs, sql, PrMsg) = True Then
    
    rirekirs.MoveFirst
    Do Until rirekirs.EOF
        cccsuu2 = cccsuu2 + 1
        ReDim Preserve datadata(5, cccsuu2)
        datadata(0, cccsuu2) = rirekirs!shousaiid
        datadata(1, cccsuu2) = Format(rirekirs!keiseikyuu, "@@@@/@@")
        datadata(2, cccsuu2) = rirekirs!keiid
        datadata(3, cccsuu2) = rirekirs!keikin
        sql = "select shuruimei from shurui where shuruiid='" & rirekirs!keiid & "'"
        If FcSQlGet(rirekirs2, sql, PrMsg) = True Then
            datadata(4, cccsuu2) = rirekirs2!shuruimei
            rirekirs2.Close
        End If
        rirekirs.MoveNext
    Loop
    rirekirs.Close

End If

sql = "select * from shousai where kojinid ='" & ididid & "'" & _
    " and keist='0' and keishurui='2'" & _
    " order by keiseikyuu"
If FcSQlGet2(rirekirs, sql, PrMsg) = True Then
    rirekirs.MoveFirst
    Do Until rirekirs.EOF
        cccsuu2 = cccsuu2 + 1
        ReDim Preserve datadata(5, cccsuu2)
        datadata(0, cccsuu2) = rirekirs!shousaiid
        datadata(1, cccsuu2) = Format(rirekirs!keiseikyuu, "@@@@/@@")
        datadata(2, cccsuu2) = rirekirs!keiid
        datadata(3, cccsuu2) = rirekirs!keikin
        sql = "select sonota from sonota where sonotaid='" & rirekirs!keiid & "'"
        If FcSQlGet(rirekirs2, sql, PrMsg) = True Then
            datadata(4, cccsuu2) = rirekirs2!sonota
            rirekirs2.Close
        End If
        rirekirs.MoveNext
    Loop
    rirekirs.Close

End If

'NHKの契約
sql = "select * from shousai where kojinid ='" & ididid & "'" & _
    " and keist='0' and keishurui='3'" & _
    " order by keiseikyuu"
If FcSQlGet2(rirekirs, sql, PrMsg) = True Then
    rirekirs.MoveFirst
    Do Until rirekirs.EOF
        cccsuu2 = cccsuu2 + 1
        ReDim Preserve datadata(5, cccsuu2)
        datadata(0, cccsuu2) = rirekirs!shousaiid
        datadata(1, cccsuu2) = Format(rirekirs!keiseikyuu, "@@@@/@@")
        datadata(2, cccsuu2) = rirekirs!keiid
        datadata(3, cccsuu2) = rirekirs!keikin
        Select Case CInt(rirekirs!keiid)
        Case 0
            datadata(4, cccsuu2) = "ＮＨＫ団体一括１ヶ月"
        Case 6
            datadata(4, cccsuu2) = "ＮＨＫ団体一括６ヶ月ハンデ"
        Case 7
            datadata(4, cccsuu2) = "ＮＨＫ団体一括６ヶ月家族割"
        Case 1
            datadata(4, cccsuu2) = "ＮＨＫ団体一括２ヶ月"
        Case 2
            datadata(4, cccsuu2) = "ＮＨＫ団体一括３ヶ月"
        Case 3
            datadata(4, cccsuu2) = "ＮＨＫ団体一括４ヶ月"
        Case 4
            datadata(4, cccsuu2) = "ＮＨＫ団体一括５ヶ月"
        Case 5
            datadata(4, cccsuu2) = "ＮＨＫ団体一括６ヶ月"
        End Select
        rirekirs.MoveNext
    Loop
    rirekirs.Close
End If

db2.Close





    frmnyuukin.gridnyuukinshousai.Clear
    With frmnyuukin.gridnyuukinshousai
        .Font.Size = 11
        .Rows = cccsuu2 + 1
        .Cols = 5
        .ColWidth(0) = 1200
        .ColWidth(1) = 1000
        .ColWidth(2) = 0
        .ColWidth(3) = 2800
        .ColWidth(4) = 1200
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "請求日"
        .Col = 2
        .Text = ""
        .Col = 3
        .Text = "契約種類"
        .Col = 4
        .Text = "金額"
    End With
    For i = 1 To cccsuu2
        With frmnyuukin.gridnyuukinshousai
            .Row = i
            .Col = 0
            .CellChecked = 2
            .CellPictureAlignment = flexPicAlignLeftCenter
            .Editable = flexEDKbd
            .Text = datadata(0, i)
            .Col = 1
            .Text = datadata(1, i)
            .Col = 2
            '.Text = Format(hiduhidu, "@@@@/@@/@@")
            .Col = 3
            .Text = datadata(4, i)
            .Col = 4
            .CellPictureAlignment = flexPicAlignLeftCenter
            .Text = datadata(3, i)
        End With
    Next i

frmnyuukin.Show 1


End Sub
Sub nyuukinrirekiset(sekouno As Integer)
Dim rirekirs As Recordset, rirekisuu As Long


'仮入金を表示

Dim s_karinyuukin As String

Dim s_kojinid As String, s_nen As String, s_tsuki As String, s_nentsuki As String

s_kojinid = Trim(frmtouroku2.lblid.Caption)
s_nentsuki = Format(Date, "yyyymmdd")
s_nen = Mid(s_nentsuki, 1, 4)
s_tsuki = Mid(s_nentsuki, 5, 2)

s_karinyuukin = kari_get(s_kojinid, s_nen, s_tsuki)

If sekouno = 0 Then
    With frmrireki
        If s_karinyuukin = "" Then
            .lblmaebi.Caption = ""
            .lblkingaku.Caption = ""
        Else
            .lblmaebi.Caption = Format(Mid(s_karinyuukin, 1, 8), "@@@@/@@/@@")
            .lblkingaku.Caption = Mid(s_karinyuukin, 9)
        End If
    End With
End If
sql = "select * from koushinkiroku where kojinid ='" & frmtouroku2.lblid.Caption & "'" & _
    " order by koushinbi DESC"
If FcSQlGet2(rirekirs, sql, PrMsg) = False Then
    MsgBox "入金履歴はありません。"
    If sekouno = 1 Then
        Unload frmrireki
    End If
    db2.Close
    
    Exit Sub
Else
    frmrireki.Caption = "入金履歴"
    frmrireki.lblkojinid.Caption = frmtouroku2.lblid.Caption
    If sekouno = 1 Then
        frmrireki.fggridrireki.Clear
    End If
    rirekirs.MoveLast
    rirekisuu = rirekirs.RecordCount
    With frmrireki.fggridrireki
        .Font.Size = 11
        .Rows = rirekisuu + 1
        .Cols = 4
        .ColWidth(0) = 1000
        .ColWidth(1) = 1500
        .ColWidth(2) = 1000
        .ColWidth(3) = 1500
        .Row = 0
        .Col = 0
        .Text = "記録ＩＤ"
        .Col = 1
        .Text = "更新日"
        .Col = 2
        .Text = "金額"
        .Col = 3
        .Text = "支払方法"
    End With
    rirekirs.MoveFirst
    For i = 1 To rirekisuu
        With frmrireki.fggridrireki
            .Row = i
            .Col = 0
            .Text = rirekirs!id
            .Col = 1
            .Text = Format(rirekirs!koushinbi, "@@@@/@@/@@")
            .Col = 2
            .Text = rirekirs!kingaku
            .Col = 3
            If rirekirs!houhou = "0" Then
                .Text = "金融機関"
            ElseIf rirekirs!houhou = "1" Then
                .Text = "現金"
            ElseIf rirekirs!houhou = "2" Then
                .Text = "振込み"
            ElseIf rirekirs!houhou = "3" Or rirekirs!houhou = "5" Or rirekirs!houhou = "7" Then
                .Text = "取消金"
            ElseIf rirekirs!houhou = "4" Or rirekirs!houhou = "6" Or rirekirs!houhou = "8" Then
                .Text = "損金"
            End If
        End With
       rirekirs.MoveNext
    Next i
    rirekirs.Close
    db2.Close
    If sekouno = 0 Then
        frmrireki.Show 1
    End If
End If

End Sub
Sub nyuukin_shousai_set(kokoid As String, hiduhidu As String)
Dim rirekirs As Recordset, rirekisuu As Long, datadata(), cccsuu2 As Long
Dim rirekirs2 As Recordset

cccsuu2 = 0
sql = "select * from shousai where kojinid ='" & frmtouroku2.lblid.Caption & "'" & _
    "and keinyu='" & hiduhidu & "' and keishurui='1'" & _
    " order by keiseikyuu"
If FcSQlGet2(rirekirs, sql, PrMsg) = True Then
    
    rirekirs.MoveFirst
    Do Until rirekirs.EOF
        cccsuu2 = cccsuu2 + 1
        ReDim Preserve datadata(5, cccsuu2)
        datadata(0, cccsuu2) = rirekirs!shousaiid
        datadata(1, cccsuu2) = Format(rirekirs!keiseikyuu, "@@@@/@@")
        datadata(2, cccsuu2) = rirekirs!keiid
        datadata(3, cccsuu2) = rirekirs!keikin
        sql = "select shuruimei from shurui where shuruiid='" & rirekirs!keiid & "'"
        If FcSQlGet(rirekirs2, sql, PrMsg) = True Then
            datadata(4, cccsuu2) = rirekirs2!shuruimei
            rirekirs2.Close
        End If
        rirekirs.MoveNext
    Loop
    rirekirs.Close

End If

sql = "select * from shousai where kojinid ='" & frmtouroku2.lblid.Caption & "'" & _
    " and keinyu='" & hiduhidu & "' and keishurui='2'" & _
    " order by keiseikyuu"
If FcSQlGet2(rirekirs, sql, PrMsg) = True Then
    rirekirs.MoveFirst
    Do Until rirekirs.EOF
        cccsuu2 = cccsuu2 + 1
        ReDim Preserve datadata(5, cccsuu2)
        datadata(0, cccsuu2) = rirekirs!shousaiid
        datadata(1, cccsuu2) = Format(rirekirs!keiseikyuu, "@@@@/@@")
        datadata(2, cccsuu2) = rirekirs!keiid
        datadata(3, cccsuu2) = rirekirs!keikin
        sql = "select sonota from sonota where sonotaid='" & rirekirs!keiid & "'"
        If FcSQlGet(rirekirs2, sql, PrMsg) = True Then
            datadata(4, cccsuu2) = rirekirs2!sonota
            rirekirs2.Close
        End If
        rirekirs.MoveNext
    Loop
    rirekirs.Close

End If

'NHKの契約
sql = "select * from shousai where kojinid ='" & frmtouroku2.lblid.Caption & "'" & _
    " and keinyu='" & hiduhidu & "' and keishurui='3'" & _
    " order by keiseikyuu"
If FcSQlGet2(rirekirs, sql, PrMsg) = True Then
    rirekirs.MoveFirst
    Do Until rirekirs.EOF
        cccsuu2 = cccsuu2 + 1
        ReDim Preserve datadata(5, cccsuu2)
        datadata(0, cccsuu2) = rirekirs!shousaiid
        datadata(1, cccsuu2) = Format(rirekirs!keiseikyuu, "@@@@/@@")
        datadata(2, cccsuu2) = rirekirs!keiid
        datadata(3, cccsuu2) = rirekirs!keikin
        Select Case CInt(rirekirs!keiid)
        Case 0
            datadata(4, cccsuu2) = "ＮＨＫ団体一括１ヶ月"
        Case 6
            datadata(4, cccsuu2) = "ＮＨＫ団体一括６ヶ月ハンデ"
        Case 7
            datadata(4, cccsuu2) = "ＮＨＫ団体一括６ヶ月家族割"
        Case 1
            datadata(4, cccsuu2) = "ＮＨＫ団体一括２ヶ月"
        Case 2
            datadata(4, cccsuu2) = "ＮＨＫ団体一括３ヶ月"
        Case 3
            datadata(4, cccsuu2) = "ＮＨＫ団体一括４ヶ月"
        Case 4
            datadata(4, cccsuu2) = "ＮＨＫ団体一括５ヶ月"
        Case 5
            datadata(4, cccsuu2) = "ＮＨＫ団体一括６ヶ月"
        End Select
        rirekirs.MoveNext
    Loop
    rirekirs.Close
End If

db2.Close





    frmrireki.gridnyuukinshousai.Clear
    With frmrireki.gridnyuukinshousai
        .Font.Size = 11
        .Rows = cccsuu2 + 1
        .Cols = 5
        .ColWidth(0) = 1200
        .ColWidth(1) = 1000
        .ColWidth(2) = 1300
        .ColWidth(3) = 3000
        .ColWidth(4) = 1200
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "請求日"
        .Col = 2
        .Text = "入金日"
        .Col = 3
        .Text = "契約種類"
        .Col = 4
        .Text = "金額"
    End With
    For i = 1 To cccsuu2
        With frmrireki.gridnyuukinshousai
            .Row = i
            .Col = 0
            .Text = datadata(0, i)
            .Col = 1
            .Text = datadata(1, i)
            .Col = 2
            .Text = Format(hiduhidu, "@@@@/@@/@@")
            .Col = 3
            .Text = datadata(4, i)
            .Col = 4
            .Text = datadata(3, i)
        End With
    Next i

End Sub

Function soukatsu(kakuno As Integer, koumokumei As String, fno As Integer, suuchi As String, f As Integer, e As Integer) As Boolean

Dim soukatsurs As Recordset, tblmei As String
If f = "1" Then
        If kidou2 = False Then
            MsgBox "データベースファイルを開けませんでした。"
            soukatsu = False
            Exit Function
        End If
End If
If kakuno = "0" Then '読み込み
    sql = "select * from settei where  koumoku ='" & koumokumei & "'"
    If FcSQlGet2(soukatsurs, sql, PrMsg) = False Then
        soukatsu = False
        'db2.Close
    Else
        Select Case fno
        Case 1
            reg = soukatsurs!t1
        Case 2
            reg = soukatsurs!t2
        Case 3
            reg = soukatsurs!t3
        Case 4
            reg = soukatsurs!t4
        Case 5
            reg = soukatsurs!t5
        Case 6
            reg = soukatsurs!t6
        Case 7
            reg = soukatsurs!t7
        Case 8
            reg = soukatsurs!t8
        Case 9
            reg = soukatsurs!t9
        Case 10
            reg = soukatsurs!t10
        Case 11
            reg = soukatsurs!t11
        Case 12
            reg = soukatsurs!t12
        Case 13
            reg = soukatsurs!t13
        Case 14
            reg = soukatsurs!t14
        Case 15
            reg = soukatsurs!t15
        Case 16
            reg = soukatsurs!t16
        Case 17
            reg = soukatsurs!t17
        Case 18
            reg = soukatsurs!t18
        Case 19
            reg = soukatsurs!t19
        Case 20
            reg = soukatsurs!t20
            
        End Select
        soukatsurs.Close
        soukatsu = True
    End If
Else        '書き込み
    Select Case fno
    Case 1
        tblmei = "t1"
    Case 2
        tblmei = "t2"
    Case 3
        tblmei = "t3"
    Case 4
        tblmei = "t4"
    Case 5
        tblmei = "t5"
    Case 6
        tblmei = "t6"
    Case 7
        tblmei = "t7"
    Case 8
        tblmei = "t8"
    Case 9
        tblmei = "t9"
    Case 10
        tblmei = "t10"
    Case 11
        tblmei = "t11"
    Case 12
        tblmei = "t12"
    Case 13
        tblmei = "t13"
    Case 14
        tblmei = "t14"
    Case 15
        tblmei = "t15"
    Case 16
        tblmei = "t16"
    Case 17
        tblmei = "t17"
    Case 18
        tblmei = "t18"
    Case 19
        tblmei = "t19"
    Case 20
        tblmei = "t20"
    End Select
    sql = "UPDATE settei SET " & tblmei & "= '" & suuchi & "' WHERE koumoku = '" & koumokumei & "'"
On Error GoTo errsoukatsu
     db2.Execute sql, dbSQLPassThrough
On Error GoTo 0
    soukatsu = True
End If

If e = "1" Then
    db2.Close
End If
Exit Function

errsoukatsu:
    soukatsu = False
    Exit Function

End Function

Sub kidou_main()
'データのロード

kidoupassword = ";pwd=" & "kamifusafusa"
'Set DB = OpenDatabase(VBNAME)
Set db = DBEngine.Workspaces(0).OpenDatabase(VBNAME, False, False, kidoupassword)


End Sub
Function kidou2() As Boolean
'データのロード

kidoupassword = ";pwd=" & "kamifusafusa"
'Set DB = OpenDatabase(VBNAME)
On Error GoTo db2open
    Set db2 = DBEngine.Workspaces(0).OpenDatabase(VBNAME2, False, False, kidoupassword)
    'DBEngine.SetOption dbPageTimeout, 5000
On Error GoTo 0
kidou2 = True
Exit Function

db2open:
    MsgBox "データベースが開けませんでした。一度再起動し、再度実行してください。"
    kidou2 = False
    Exit Function
End Function
Function kidou_voip() As Boolean
'データのロード

kidoupassword = ";pwd=" & "kamifusafusa"
'Set DB = OpenDatabase(VBNAME)
On Error GoTo db4open2
    Set db_voip = DBEngine.Workspaces(0).OpenDatabase(vbname_voip, False, False, kidoupassword)
On Error GoTo 0
kidou_voip = True
Exit Function

db4open2:
    MsgBox "データベースが開けませんでした。一度再起動し、再度実行してください。"
    kidou_voip = False
    Exit Function
End Function
Function kidou_out() As Boolean
'データのロード

kidoupassword = ";pwd=" & "kamifusafusa"
'Set DB = OpenDatabase(VBNAME)
On Error GoTo db4open2
    Set db_out = DBEngine.Workspaces(0).OpenDatabase(vbname_out, False, False, kidoupassword)
On Error GoTo 0
kidou_out = True
Exit Function

db4open2:
    MsgBox "データベースが開けませんでした。一度再起動し、再度実行してください。"
    kidou_out = False
    Exit Function
End Function
Function kidou_conbart() As Boolean
'データのロード

'Set DB = OpenDatabase(VBNAME)
On Error GoTo db4open2c
    Set db_convart = DBEngine.Workspaces(0).OpenDatabase(vbname_convart, False, False)
On Error GoTo 0
kidou_conbart = True
Exit Function

db4open2c:
    MsgBox "データベースが開けませんでした。一度再起動し、再度実行してください。"
    kidou_conbart = False
    Exit Function
End Function
Function kidou_ppv() As Boolean
'データのロード

kidoupassword = ";pwd=" & "kamifusafusa"
'Set DB = OpenDatabase(VBNAME)
On Error GoTo db4open23
    Set db_ppv = DBEngine.Workspaces(0).OpenDatabase(vbname_ppv, False, False, kidoupassword)
On Error GoTo 0
kidou_ppv = True
Exit Function

db4open23:
    MsgBox "データベースが開けませんでした。一度再起動し、再度実行してください。"
    kidou_ppv = False
    Exit Function
End Function


Function kidou4() As Boolean
'データのロード

kidoupassword = ";pwd=" & "kamifusafusa"
'Set DB = OpenDatabase(VBNAME)
On Error GoTo db4open
    Set db4 = DBEngine.Workspaces(0).OpenDatabase(vbname4, False, False, kidoupassword)
On Error GoTo 0
kidou4 = True
Exit Function

db4open:
    
    MsgBox "データベースが開けませんでした。一度再起動し、再度実行してください。"
    kidou4 = False
    Exit Function
End Function

Sub errorfd(e As Integer)

Dim m As String

Select Case e
Case -1
   m = "フォーマット内で不正な文字を検出しました。"
Case -2
   m = "小数宣言が不正です。"
Case -3
   m = "フォーマットと変数の型が違います。"
Case -4
   m = "フォーマットと変数の数が違います。"
Case -5
   m = "桁数が不正です。"
Case -6
   m = "文字が無効です。"
Case -7
   m = "ファイルがオープンされていません。"
Case -8
   m = "実行時エラー"
Case -9
   m = "読み込みファイルがありません。"
Case -10
   m = "書き込みファイルがすでにあります。"
Case -11
   m = "実行時エラー"
Case -13
   m = "データがありません。"
Case -310
   m = "フォーマット形式が違います。"
Case -311
   m = "数字の右側の０追加・０削除"
Case -312
   m = "フォーマット形式とデータ内容が違います。"
Case -313
   m = "フォーマット形式とデータサイズが違います。"
Case -410
   m = "フォーマット形式が違います。"

End Select

m = m & "errno " & CStr(e)
MsgBox (m)

End Sub
