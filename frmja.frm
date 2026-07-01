VERSION 5.00
Begin VB.Form frmja 
   Caption         =   "JA統合"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1395
   ScaleWidth      =   5505
   StartUpPosition =   2  '画面の中央
   Begin VB.CheckBox chkmeigi 
      Caption         =   "口座名義を更新する"
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
      Left            =   660
      TabIndex        =   6
      Top             =   1020
      Width           =   2535
   End
   Begin VB.OptionButton Option2 
      Caption         =   "データを読み込み更新します。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   5
      Top             =   600
      Width           =   3135
   End
   Begin VB.OptionButton Option1 
      Caption         =   "データを出力します。"
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
      Left            =   180
      TabIndex        =   4
      Top             =   240
      Value           =   -1  'True
      Width           =   2535
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   4680
      MouseIcon       =   "frmja.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmja.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   420
      Width           =   480
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3690
      MouseIcon       =   "frmja.frx":11C9
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmja.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   390
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
      Left            =   4710
      TabIndex        =   3
      Top             =   900
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
      Left            =   3690
      TabIndex        =   2
      Top             =   900
      Width           =   480
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   855
      Left            =   3360
      Top             =   360
      Width           =   2010
   End
End
Attribute VB_Name = "frmja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub ja_data_shutsuryoku(shutsu_path As String)

Dim sql_sentaku As String, rs_sentaku As Recordset, sususu As Long, csv_data()
Dim datasuu As Long, writecounter As Long, kensakukoumokusuu As Integer, writecounter2 As Long
Dim sql_sentaku2 As String, rs_sentaku2 As Recordset

sql_sentaku = "select kojin.*,bank.* from kojin LEFT JOIN bank ON kojin.bankid = bank.bankid" & _
                " where bank.bankcode='5169'"

If FcSQlGet(rs_sentaku, sql_sentaku, PrMsg) = False Then
            ret = MsgBox("作成したいデータが存在しません。", 16, "総合管理システム「EZ Manager」")
            Exit Sub
Else
    datasuu = rs_sentaku.RecordCount
        ReDim csv_data(9, datasuu)
        csv_data(0, 0) = "NO"
        csv_data(1, 0) = "個人 ID"
        csv_data(2, 0) = "個人名"
        csv_data(3, 0) = "口座名義"
        csv_data(4, 0) = "金融機関コード"
        csv_data(5, 0) = "旧店舗コード"
        csv_data(6, 0) = "旧口座番号"
        csv_data(7, 0) = "新店舗コード"
        csv_data(8, 0) = "新口座番号"
        
    rs_sentaku.MoveFirst
    sususu = 1
    Do Until rs_sentaku.EOF
            csv_data(0, sususu) = Format(sususu, "00000")
            csv_data(1, sususu) = rs_sentaku!kojinid
            csv_data(2, sususu) = rs_sentaku!Name
            csv_data(3, sususu) = rs_sentaku!meigi
            csv_data(4, sususu) = rs_sentaku!bankcode
            csv_data(5, sususu) = rs_sentaku!shitencode
            csv_data(6, sususu) = rs_sentaku!kouzabangou
        sususu = sususu + 1
        rs_sentaku.MoveNext
    Loop
    rs_sentaku.Close
    
End If
For writecounter = 1 To datasuu
        For writecounter2 = 1 To 8
          If IsNull(csv_data(writecounter2, writecounter)) Then
            csv_data(writecounter2, writecounter) = ""
          End If
        Next writecounter2
Next writecounter
Open shutsu_path For Output Access Write As 1
For writecounter = 0 To datasuu
        Write #1, csv_data(0, writecounter), csv_data(1, writecounter), csv_data(2, writecounter), csv_data(3, writecounter), csv_data(4, writecounter), csv_data(5, writecounter), csv_data(6, writecounter), csv_data(7, writecounter), csv_data(8, writecounter)
Next

Close #1

ret = MsgBox("指定データのエクスポートが完了しました。", 64, "EZ Manager")

End Sub
Sub ja_data_yomikomi(shutsu_path As String)
   
   Dim GetLine As Variant
   Dim Spot As Integer, Position As Integer, newkojinid As String
   Dim Row As Integer, Col As Integer, newvoipnospot2 As Integer
   Dim RowCount As Integer, ColCount As Integer, ColCount2 As Integer, ColCount3 As Integer
    
    Dim newzei As Long, newsou As Long, newvoipnospot As Integer, newv1 As String, newv2 As String, newv3 As String
    Dim kariatai, sql_voip2 As String, newvoipno As String, sql_voip299 As String
    Dim newtsuwaid As String, voipii As Integer, rs_voip As Recordset, rs_voip2 As Recordset, rs_voip299 As Recordset

If Dir(shutsu_path) = "" Then
        ret = MsgBox("指定ファイルが見つかりませんでした。", 16, "EZ Manager")
        Exit Sub
End If
    If MsgBox("読み込みと登録データのチェックを開始します。", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
Screen.MousePointer = 11
      
      
On Error GoTo ErrorTrap3
   Open shutsu_path For Input Lock Read Write As #1
   ColCount = 0
   RowCount = 10
    While Not EOF(1)
        Line Input #1, GetLine
        ColCount = ColCount + 1
    
    Wend
    Close #1
    
    Open shutsu_path For Input Lock Read Write As #1
   ReDim GetData(ColCount - 1, RowCount)
   Position = 1
   Row = 0
   While Not EOF(1)
      Line Input #1, GetLine
      Col = 0
      Spot = InStr(1, GetLine, ",")
      While Spot <> 0
         kariatai = Left$(GetLine, Spot - 1)
        If Mid(kariatai, 1, 1) = """" Then
            If Right(kariatai, 1) = """" Then
                GetData(Row, Col) = Mid(kariatai, 2, Len(kariatai) - 2)
            Else
                GetData(Row, Col) = Mid(kariatai, 2)
            End If
        Else
            If Right(kariatai, 1) = """" Then
                GetData(Row, Col) = Mid(kariatai, 1, Len(kariatai) - 1)
            Else
                GetData(Row, Col) = kariatai
            End If
        End If
         Col = Col + 1
         GetLine = Right$(GetLine, Len(GetLine) - Spot)
         Spot = InStr(1, GetLine, ",")
      Wend
      If Len(GetLine) <> 0 Then
            If Mid(GetLine, 1, 1) = """" Then
                If Right(GetLine, 1) = """" Then
                    GetData(Row, Col) = Mid(GetLine, 2, Len(GetLine) - 2)
                Else
                    GetData(Row, Col) = Mid(GetLine, 2)
                End If
            Else
                If Right(GetLine, 1) = """" Then
                    GetData(Row, Col) = Mid(GetLine, 1, Len(GetLine) - 1)
                Else
                    GetData(Row, Col) = GetLine
                End If
            End If
      End If
      Row = Row + 1
   Wend
   Close #1


On Error GoTo errvoip
'*********トランザクション開始**************
BeginTrans
    Dim newbbid As String
        For voipii = 1 To ColCount - 1
              If Trim(GetData(voipii, 7)) <> "" Then
                    '個人IDの確認
                    sql_voip2 = "select kojinid from kojin where kojinid ='" & GetData(voipii, 1) & "'"
            
                    If FcSQlGet(rs_voip2, sql_voip2, PrMsg) = False Then
                        Screen.MousePointer = 0
                        Rollback
                        MsgBox "番号" & GetData(voipii, 1) & "「" & GetData(voipii, 2) & "」は、登録されていません。確認してから再度実行してください。"
                        Exit Sub
                    Else
                        rs_voip2.Close
                    End If
                    '個人IDの確認
                    sql_voip2 = "select bankid from bank where bankcode ='5169' and shitencode ='" & Format(GetData(voipii, 7), "000") & "'"
            
                    If FcSQlGet(rs_voip299, sql_voip2, PrMsg) = False Then
                        Screen.MousePointer = 0
                        Rollback
                        MsgBox "支店コード「" & GetData(voipii, 7) & "」は、登録されていません。確認してから再度実行してください。"
                        Exit Sub
                    Else
                        newbbid = rs_voip299!bankid
                        rs_voip299.Close
                    End If
                    
                    
                    
                    
                    sql = "update kojin set kouzabangou = '" & Format(GetData(voipii, 8), "0000000") & "',bankid = '" & newbbid & "'" & _
                            " where kojinid ='" & GetData(voipii, 1) & "'"
                                   
                    db.Execute sql, dbSQLPassThrough
               End If
        Next voipii
    
'*************トランザクション終了********
CommitTrans
    
    
    Screen.MousePointer = 0
    MsgBox "データの作成が完了しました。"

Exit Sub

ErrorTrap3:
        Screen.MousePointer = 0
        Close #1
        MsgBox "データの読み込みに失敗しました。"
        Exit Sub

errvoip:
        Screen.MousePointer = 0
        Rollback
        MsgBox "登録に失敗しました。"
        Exit Sub



End Sub

Sub ja_data_yomikomi2(shutsu_path As String)
   
   Dim GetLine As Variant
   Dim Spot As Integer, Position As Integer, newkojinid As String, hikaku_meigi As String
   Dim Row As Integer, Col As Integer, newvoipnospot2 As Integer, henkousuu As Integer, spmade As Integer
   Dim RowCount As Integer, ColCount As Integer, ColCount2 As Integer, ColCount3 As Integer, newmeigi1 As String, newmeigi2 As String, newmeigi3 As String
    
    Dim newzei As Long, newsou As Long, newvoipnospot As Integer, newv1 As String, newv2 As String, newv3 As String
    Dim kariatai, sql_voip2 As String, newvoipno As String, sql_voip299 As String
    Dim newtsuwaid As String, voipii As Integer, rs_voip As Recordset, rs_voip2 As Recordset, rs_voip299 As Recordset

If Dir(shutsu_path) = "" Then
        ret = MsgBox("指定ファイルが見つかりませんでした。", 16, "EZ Manager")
        Exit Sub
End If
    If MsgBox("読み込みと登録データのチェックを開始します。", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
Screen.MousePointer = 11
      
      
On Error GoTo ErrorTrap3
   Open shutsu_path For Input Lock Read Write As #1
   ColCount = 0
   RowCount = 10
    While Not EOF(1)
        Line Input #1, GetLine
        ColCount = ColCount + 1
    
    Wend
    Close #1
    
    Open shutsu_path For Input Lock Read Write As #1
   ReDim GetData(ColCount - 1, RowCount)
   Position = 1
   Row = 0
   While Not EOF(1)
      Line Input #1, GetLine
      Col = 0
      Spot = InStr(1, GetLine, ",")
      While Spot <> 0
         kariatai = Left$(GetLine, Spot - 1)
        If Mid(kariatai, 1, 1) = """" Then
            If Right(kariatai, 1) = """" Then
                GetData(Row, Col) = Mid(kariatai, 2, Len(kariatai) - 2)
            Else
                GetData(Row, Col) = Mid(kariatai, 2)
            End If
        Else
            If Right(kariatai, 1) = """" Then
                GetData(Row, Col) = Mid(kariatai, 1, Len(kariatai) - 1)
            Else
                GetData(Row, Col) = kariatai
            End If
        End If
         Col = Col + 1
         GetLine = Right$(GetLine, Len(GetLine) - Spot)
         Spot = InStr(1, GetLine, ",")
      Wend
      If Len(GetLine) <> 0 Then
            If Mid(GetLine, 1, 1) = """" Then
                If Right(GetLine, 1) = """" Then
                    GetData(Row, Col) = Mid(GetLine, 2, Len(GetLine) - 2)
                Else
                    GetData(Row, Col) = Mid(GetLine, 2)
                End If
            Else
                If Right(GetLine, 1) = """" Then
                    GetData(Row, Col) = Mid(GetLine, 1, Len(GetLine) - 1)
                Else
                    GetData(Row, Col) = GetLine
                End If
            End If
      End If
      Row = Row + 1
   Wend
   Close #1


On Error GoTo errvoip
'*********トランザクション開始**************

henkousuu = 0
BeginTrans
    Dim newbbid As String
        For voipii = 1 To ColCount - 1
             If Trim(GetData(voipii, 9)) <> "" Then
                    '個人IDの確認
                    sql_voip2 = "select kojinid,meigi from kojin where kojinid ='" & GetData(voipii, 1) & "'"
                    hikaku_meigi = ""
                    If FcSQlGet(rs_voip2, sql_voip2, PrMsg) = False Then
                        Screen.MousePointer = 0
                        Rollback
                        MsgBox "番号" & GetData(voipii, 1) & "「" & GetData(voipii, 2) & "」は、登録されていません。確認してから再度実行してください。"
                        Exit Sub
                    Else
                        hikaku_meigi = Trim(rs_voip2!meigi)
                        rs_voip2.Close
                    End If
                    'めいぎの確認
                    spmade = InStr(Trim(GetData(voipii, 9)), " ")
                    If spmade <> 0 Then
                        newmeigi1 = Mid(Trim(GetData(voipii, 9)), 1, spmade - 1)
                        newmeigi2 = Mid(Trim(GetData(voipii, 9)), spmade + 1)
                        newmeigi3 = newmeigi1 & newmeigi2
                    Else
                        newmeigi3 = Trim(GetData(voipii, 9))
                    End If
                        
                        If hikaku_meigi <> newmeigi3 Then
                        
                            sql = "update kojin set meigi = '" & newmeigi3 & "'" & _
                                    " where kojinid ='" & GetData(voipii, 1) & "'"
                                           
                            db.Execute sql, dbSQLPassThrough
                            henkousuu = henkousuu + 1
                        End If
                    End If
        Next voipii
    
'*************トランザクション終了********
CommitTrans
    
    
    Screen.MousePointer = 0
    MsgBox henkousuu & "件のデータの修正が完了しました。"

Exit Sub

ErrorTrap3:
        Screen.MousePointer = 0
        Close #1
        MsgBox "データの読み込みに失敗しました。"
        Exit Sub

errvoip:
        Screen.MousePointer = 0
        Rollback
        MsgBox "登録に失敗しました。"
        Exit Sub



End Sub

Private Sub Picture10_Click()

Dim shutsuryokufailename As String
Dim udtBrowseInfo As BROWSEINFO
Dim lngIDList     As Long, strPath As String

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
        shutsuryokufailename = strPath
        'ITEMIDLISTを解放
        CoTaskMemFree lngIDList
    End If


If Option1.Value = True Then
    If shutsuryokufailename = "" Then
        Exit Sub
    Else
        shutsuryokufailename = shutsuryokufailename & "\ja_data.txt"
        ja_data_shutsuryoku (shutsuryokufailename)
    End If
Else
    shutsuryokufailename = shutsuryokufailename & "\ja_data_go.txt"
   If chkmeigi.Value = 0 Then
        ja_data_yomikomi (shutsuryokufailename)
    Else
        ja_data_yomikomi2 (shutsuryokufailename)
    End If
End If

End Sub

Private Sub Picture13_Click()
Unload Me
End Sub
