VERSION 5.00
Begin VB.Form frmkanrisha 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "管理者フォーム"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   8925
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command4 
      Caption         =   "追加"
      Height          =   495
      Left            =   150
      TabIndex        =   15
      Top             =   1380
      Width           =   1485
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ST変更"
      Height          =   495
      Left            =   3960
      TabIndex        =   14
      Top             =   1380
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "space(1)追加"
      Height          =   495
      Left            =   5520
      TabIndex        =   13
      Top             =   1380
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "shousai追加"
      Height          =   495
      Left            =   7080
      TabIndex        =   12
      Top             =   1410
      Width           =   1485
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   7320
      MouseIcon       =   "frmkanrisha.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmkanrisha.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   330
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8130
      MouseIcon       =   "frmkanrisha.frx":11C9
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmkanrisha.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   330
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "起動する処理を選択してください。"
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
      Height          =   1065
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6795
      Begin VB.PictureBox Picture16 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'なし
         Height          =   480
         Left            =   300
         MouseIcon       =   "frmkanrisha.frx":2392
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmkanrisha.frx":2C91
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   3
         Top             =   360
         Width           =   480
      End
      Begin VB.ComboBox combkaisha 
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
         Left            =   960
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   1
         Top             =   480
         Width           =   4485
      End
      Begin VB.Label Label4 
         Caption         =   "処理"
         Height          =   255
         Left            =   6240
         TabIndex        =   11
         Top             =   180
         Width           =   495
      End
      Begin VB.Label lbls 
         Caption         =   "0"
         Height          =   255
         Left            =   5610
         TabIndex        =   10
         Top             =   150
         Width           =   585
      End
      Begin VB.Label Label3 
         Caption         =   "処理中・・"
         Height          =   255
         Left            =   5970
         TabIndex        =   9
         Top             =   720
         Width           =   795
      End
      Begin VB.Label shoricount 
         Caption         =   "0"
         Height          =   255
         Left            =   5610
         TabIndex        =   8
         Top             =   450
         Width           =   705
      End
      Begin VB.Label lblkensuu 
         Height          =   255
         Left            =   300
         TabIndex        =   2
         Top             =   1200
         Width           =   975
      End
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
      Index           =   15
      Left            =   7320
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
      Left            =   8190
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   855
      Left            =   7050
      Top             =   270
      Width           =   1740
   End
End
Attribute VB_Name = "frmkanrisha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
Dim newkojinid As String, newseikyuubi As String, newshurui As String, newshurui_no As String, newkin As String, newbankid As String


 
newkojinid = "00015"

newseikyuubi = "202305" ' Mid(newnyukinbi, 1, 6)

newshurui = "1"

newshurui_no = "35"  '35=番号表示　　39=着信拒否

newkin = "440"      '35＝440　　39＝550

newbankid = Space(1)

 If shousai_main_set(newkojinid, newseikyuubi, newshurui, newshurui_no, newkin, newbankid, "1") = -1 Then
            MsgBox "契約の更新に失敗しました。"
            Exit Sub
  End If
  
  
  MsgBox "契約の更新しました。"
  
End Sub

Private Sub Command2_Click()
'請求データの修正
    If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If
    
    sql = "UPDATE shousai SET keinyu =' ' WHERE shousaiid = '000675083'"
    
    On Error GoTo errko1
    db2.Execute sql, dbSQLPassThrough
    On Error GoTo 0
    
     '' sql = "UPDATE shousai SET keinyu =' ' WHERE shousaiid = '000644143'"
    
   ' On Error GoTo errko1
   ' db2.Execute sql, dbSQLPassThrough
  '  On Error GoTo 0
    
    
   
    db2.Close
    
    MsgBox "契約の更新しました。"
    
    Exit Sub
    
errko1:
    
    MsgBox "NG。"
    
    Exit Sub
    
    
    
End Sub

Private Sub Command3_Click()
    If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If
    
    sql = "UPDATE shousai SET keist ='1' WHERE shousaiid = '000647639'"
    
    On Error GoTo errko1
    db2.Execute sql, dbSQLPassThrough
    On Error GoTo 0
    
      sql = "UPDATE shousai SET keist ='1' WHERE shousaiid = '000647640'"
    
    On Error GoTo errko1
    db2.Execute sql, dbSQLPassThrough
    On Error GoTo 0
    
    
      sql = "UPDATE shousai SET keist ='1' WHERE shousaiid = '000647641'"
    
    On Error GoTo errko1
    db2.Execute sql, dbSQLPassThrough
    On Error GoTo 0
   
   
      sql = "UPDATE shousai SET keist ='1' WHERE shousaiid = '000647643'"
    
    On Error GoTo errko1
    db2.Execute sql, dbSQLPassThrough
    On Error GoTo 0
   
    db2.Close
    
    
   
   Dim newkojinid As String, newseikyuubi As String, newshurui As String, newshurui_no As String, newkin As String, newbankid As String


 
newkojinid = "03020"

newseikyuubi = "202104" ' Mid(newnyukinbi, 1, 6)

newshurui = "1"

newshurui_no = "06"

newkin = "2750"      '35＝440　　39＝550

newbankid = Space(1)

 If shousai_main_set(newkojinid, newseikyuubi, newshurui, newshurui_no, newkin, newbankid, "1") = -1 Then
            MsgBox "契約の更新に失敗しました。"
            Exit Sub
  End If
  
  
  newkojinid = "03020"

newseikyuubi = "202105" ' Mid(newnyukinbi, 1, 6)

newshurui = "1"

newshurui_no = "06"

newkin = "2750"      '35＝440　　39＝550

newbankid = Space(1)

 If shousai_main_set(newkojinid, newseikyuubi, newshurui, newshurui_no, newkin, newbankid, "1") = -1 Then
            MsgBox "契約の更新に失敗しました。"
            Exit Sub
  End If
   
   
   newkojinid = "03020"

newseikyuubi = "202105" ' Mid(newnyukinbi, 1, 6)

newshurui = "1"

newshurui_no = "02"

newkin = "2750"      '35＝440　　39＝550

newbankid = Space(1)

 If shousai_main_set(newkojinid, newseikyuubi, newshurui, newshurui_no, newkin, newbankid, "1") = -1 Then
            MsgBox "契約の更新に失敗しました。"
            Exit Sub
  End If
  
  
  
    MsgBox "契約の更新しました。"
    
    Exit Sub
    
errko1:
    
    MsgBox "NG。"
    
    Exit Sub
End Sub

Private Sub Command4_Click()
  Dim sql_666 As String
  
  
 sql_666 = "update kojin set  minyuukin='13750',jikaikin='5500',soukei='19250'" & _
           " where kojinid ='03020'"

On Error GoTo errko1
db.Execute sql_666, dbSQLPassThrough
On Error GoTo 0
  
  
  
  
  If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If
    
    
      Dim newkojinid As String, newseikyuubi As String, newshurui As String, newshurui_no As String, newkin As String, newbankid As String



  
  
  newkojinid = "03020"

newseikyuubi = "202106" ' Mid(newnyukinbi, 1, 6)

newshurui = "1"

newshurui_no = "06"

newkin = "2750"      '35＝440　　39＝550

newbankid = Space(1)

 If shousai_main_set(newkojinid, newseikyuubi, newshurui, newshurui_no, newkin, newbankid, "1") = -1 Then
            MsgBox "契約の更新に失敗しました。"
            Exit Sub
  End If
   
   
   newkojinid = "03020"

newseikyuubi = "202106" ' Mid(newnyukinbi, 1, 6)

newshurui = "1"

newshurui_no = "02"

newkin = "2750"      '35＝440　　39＝550

newbankid = Space(1)

 If shousai_main_set(newkojinid, newseikyuubi, newshurui, newshurui_no, newkin, newbankid, "1") = -1 Then
            MsgBox "契約の更新に失敗しました。"
            Exit Sub
  End If
  
  
  
    MsgBox "契約の更新しました。"
    
    Exit Sub
    
errko1:
    
    MsgBox "NG。"
    
    Exit Sub
End Sub

Private Sub Form_Load()

combkaisha.AddItem "各種データの手動修正89"
combkaisha.AddItem "各種データの手動修正90"
combkaisha.AddItem "各種データの手動修正91"
combkaisha.AddItem "NHK増税（手動修正92）"
combkaisha.AddItem "手動修正93"
combkaisha.AddItem "解約日の追加修正（手動修正94）"
combkaisha.AddItem "手動０更新及びその他の更新"
combkaisha.AddItem "各種データの手動修正96"
combkaisha.AddItem "各種データの手動修正97"
combkaisha.AddItem "各種データの手動修正98"
combkaisha.AddItem "銀行支店統合修正99"

End Sub
Private Sub Picture1_Click()

Dim kakino As Integer

    If soukatsu(0, "startup", 15, "", 1, 1) = False Then
            MsgBox "設定の読み込みに失敗しました。"
    End If

 kakino = 89 + combkaisha.ListIndex

If reg >= kakino Then
    MsgBox "すでにこの修正は実行済みです。"
    Exit Sub
End If

Select Case combkaisha.ListIndex
Case 0
    'shudoushuusei85
Case 1
    'shudoushuusei90
Case 2
    'shudoushuusei91
Case 3
    'nhk_zougaku2
Case 4
    'shudoushuusei93
  Case 5
    'shudoushuusei94

 Case 6
    'shudoushuusei95
Case 7
    'shudoushuusei96
    '
Case 8
    'shudoushuusei97
    '
Case 9
    'shudoushuusei98
Case 10
    shudoushuusei99
    
Case Else
    MsgBox "項目を指定してから再度実行してください。"
    Exit Sub
End Select

    If soukatsu(1, "startup", 15, CStr(kakino), 1, 1) = False Then
        Screen.MousePointer = 0
                MsgBox "設定の書き込みに失敗しました。"
                Exit Sub
    End If

Exit Sub

End Sub
Sub shudoushuusei99()


Dim fdrs2  As Recordset


  '〆のflagをたてる
            
               
     sql = "update kojin set  bankid='061'" & _
               " where bankid ='065'"
    
    On Error GoTo errko1
    db.Execute sql, dbSQLPassThrough
    On Error GoTo 0
                
                
                MsgBox "OK"
                
                Exit Sub
                
                
errko1:
        MsgBox "err"
        Exit Sub
        
        
End Sub

Sub shudoushuusei97()


Dim fdrs2  As Recordset


  '〆のflagをたてる
            
               Dim pbp As Long
                sql = "select kojin.kojinid from kojin LEFT JOIN bank ON kojin.bankid = bank.bankid" & _
                            " where bank.bankcode='0142' and kojin.soukei >'0'"
                'recordset
                If FcSQlGet(fdrs2, sql, PrMsg) = False Then
                       MsgBox "NG"
                       Exit Sub
                End If
                 
                fdrs2.MoveFirst
                Do Until fdrs2.EOF
                    If flgup(1, fdrs2!kojinid) = False Then
                        MsgBox "NG"
                        
                        Exit Sub
                    End If
                    fdrs2.MoveNext
                Loop
                fdrs2.Close
                
                
                MsgBox "OK"

End Sub
Sub shudoushuusei98()

  Dim newkojinid As String, newseikyuubi As String, newshurui As String, newshurui_no As String, newkin As String, newbankid As String


 
newkojinid = "01945"

newseikyuubi = "202303" ' Mid(newnyukinbi, 1, 6)

newshurui = "1"

newshurui_no = "06"  '35=番号表示　　39=着信拒否

newkin = "2750"      '35＝440　　39＝550

newbankid = Space(1)

 If shousai_main_set(newkojinid, newseikyuubi, newshurui, newshurui_no, newkin, newbankid, "1") = -1 Then
            MsgBox "契約の更新に失敗しました。"
            Exit Sub
  End If
  
  newkojinid = "01945"

newseikyuubi = "202303" ' Mid(newnyukinbi, 1, 6)

newshurui = "2"

newshurui_no = "26"  '35=番号表示　　39=着信拒否

newkin = "3300"      '35＝440　　39＝550

newbankid = Space(1)

 If shousai_main_set(newkojinid, newseikyuubi, newshurui, newshurui_no, newkin, newbankid, "1") = -1 Then
            MsgBox "契約の更新に失敗しました。"
            Exit Sub
  End If
  
  MsgBox "契約の更新しました。"


End Sub
Sub shudoushuusei96()

'00534
    '請求データの修正
    If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If
    
    sql = "UPDATE shousai SET keikin ='13500' WHERE shousaiid = '000260034'"
    
    On Error GoTo errko1
    db2.Execute sql, dbSQLPassThrough
    On Error GoTo 0
    
      sql = "UPDATE shousai SET keikin ='13500' WHERE shousaiid = '000410677'"
    
    On Error GoTo errko1
    db2.Execute sql, dbSQLPassThrough
    On Error GoTo 0
    
     sql = "UPDATE shousai SET keist ='0',keinyu =' ' WHERE shousaiid = '000403295'"
    
    On Error GoTo errko1
    db2.Execute sql, dbSQLPassThrough
    On Error GoTo 0
    
    
   sql = "delete * from koushinkiroku where id ='0216996'"
        On Error GoTo errko1
         db2.Execute sql, dbSQLPassThrough
        On Error GoTo 0
    
    db2.Close
    
   

     
    'メイン画面の修正
    
    
     sql = "update kojin set  minyuukin='13500',soukei='30000'" & _
               " where kojinid ='00534'"
    
    On Error GoTo errko1
    db.Execute sql, dbSQLPassThrough
    On Error GoTo 0
    
    
    
'01210

      '入金取り消し
      
      
     sql = "update kojin set  minyuukin='1500',soukei='1500'" & _
               " where kojinid ='01210'"
    
    On Error GoTo errko1
    db.Execute sql, dbSQLPassThrough
    On Error GoTo 0
    
    
 MsgBox "OK"
    Exit Sub

errko1:
        MsgBox "err"
        Exit Sub

End Sub


Sub shudoushuusei95()

'請求書・現金自動更新
    Dim kyounohi As String, chknohi As String, chkgo As Integer
    Dim zeorobi As String

        msg = "口座引落（請求額が０円以下）と請求書扱いの内容を更新します。"
        res = MsgBox(msg, vbYesNo)
        If res = vbYes Then
 

                  
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
                        End If
                    End If
                              
           
        Else
            MsgBox "中止しました。次回起動時に再度更新処理を行います。"
        End If
   
    
End Sub
Sub shudoushuusei94()


'DM履歴テーブル追加
'テーブル追加
'new_data_base_open

On Error GoTo errsudeni
db.Execute "ALTER table kojin ADD COLUMN kaiyakubi TEXT(8)"
db.Execute "ALTER table kojin ADD COLUMN kaiyakubi2 TEXT(8)"

On Error GoTo 0


'db.Close


  MsgBox "OK"
Exit Sub

errsudeni:

  MsgBox "Err"
End Sub
Sub shudoushuusei93()
        If soukatsu(1, "renban", 13, "2470", 1, 1) = False Then
                MsgBox "設定の書き込みに失敗しました。"
                Exit Sub
        End If
        
        MsgBox "OK"
        
        
End Sub

Sub shudoushuusei91()

Dim sql_666 As String, sql_6663 As String, sql_6662 As String, newshusei As Long
Dim rs_666 As Recordset, rs_6662 As Recordset, souda As Long, ssid As String
Dim henkours As Recordset, henkours2 As Recordset, oldatai As Long, newatai As Long
Dim kyusoukei As Long, tid As String, ttid As String, rirekirs As Recordset

Dim id_data(25) As String, iuiu As Integer
Dim koji_data(10) As Long, shitei_shousaiid As String



If kidou2 = False Then
    MsgBox "データベースファイルを開けませんでした。"
    Exit Sub
End If

sql_6663 = "UPDATE shousai SET keiseikyuu ='201308' WHERE shousaiid = '000281853'"

On Error GoTo errko1
db2.Execute sql_6663, dbSQLPassThrough
On Error GoTo 0



db2.Close

MsgBox "OK"
Exit Sub
errko1:
    MsgBox "err"
    Exit Sub


End Sub
Sub shudoushuusei90()

Dim sql_666 As String, sql_6663 As String, sql_6662 As String, newshusei As Long
Dim rs_666 As Recordset, rs_6662 As Recordset, souda As Long, ssid As String
Dim henkours As Recordset, henkours2 As Recordset, oldatai As Long, newatai As Long
Dim kyusoukei As Long, tid As String, ttid As String, rirekirs As Recordset

Dim id_data(25) As String, iuiu As Integer
Dim koji_data(10) As Long, shitei_shousaiid As String


 sql_666 = "update kojin set  minyuukin='177830',soukei='179330'" & _
           " where kojinid ='02905'"

On Error GoTo errko1
db.Execute sql_666, dbSQLPassThrough
On Error GoTo 0





If kidou2 = False Then
    MsgBox "データベースファイルを開けませんでした。"
    Exit Sub
End If

sql_6663 = "UPDATE shousai SET keikin ='1500' WHERE shousaiid = '000273599'"

On Error GoTo errko1
db2.Execute sql_6663, dbSQLPassThrough
On Error GoTo 0


sql_6663 = "UPDATE shousai SET keikin ='1500' WHERE shousaiid = '000277065'"

On Error GoTo errko1
db2.Execute sql_6663, dbSQLPassThrough
On Error GoTo 0


db2.Close

MsgBox "OK"
Exit Sub
errko1:
    MsgBox "err"
    Exit Sub



End Sub
Sub shudoushuusei85()
Dim sql_666 As String, sql_6663 As String, sql_6662 As String, newshusei As Long
Dim rs_666 As Recordset, rs_6662 As Recordset, souda As Long, ssid As String
Dim henkours As Recordset, henkours2 As Recordset, oldatai As Long, newatai As Long
Dim kyusoukei As Long, tid As String, ttid As String, rirekirs As Recordset

Dim id_data(25) As String, iuiu As Integer
Dim koji_data(10) As Long, shitei_shousaiid As String


 sql_666 = "update kojin set  minyuukin='6000',soukei='9000'" & _
           " where kojinid ='02635'"

On Error GoTo errko1
db.Execute sql_666, dbSQLPassThrough
On Error GoTo 0





If kidou2 = False Then
    MsgBox "データベースファイルを開けませんでした。"
    Exit Sub
End If

sql_6663 = "UPDATE shousai SET keikin ='3000' WHERE shousaiid = '000262211'"

On Error GoTo errko1
db2.Execute sql_6663, dbSQLPassThrough
On Error GoTo 0


sql_6663 = "UPDATE shousai SET keikin ='3000' WHERE shousaiid = '000265670'"

On Error GoTo errko1
db2.Execute sql_6663, dbSQLPassThrough
On Error GoTo 0


db2.Close

MsgBox "OK"
Exit Sub
errko1:
    MsgBox "err"
    Exit Sub





End Sub
Sub nhk_zougaku2()


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
                               Case "2020"
                                    newkingaku = "2080"
                                    sagaku = 430
                               Case "4040"
                                    newkingaku = "4160"
                                    sagaku = 430
                               Case "11460"
                                    newkingaku = "11815"
                                    sagaku = 430
                               Case "22250"
                                    newkingaku = "22920"
                                    sagaku = 430
                               Case "5130"
                                    newkingaku = "5307"
                                    sagaku = 65
                               Case "0"
                                    jikkou = 1
                               Case Else
                                    MsgBox "不正な金額が登録されています。NHKID:" & snhkid
                                    Exit Sub
                               End Select
                               
                               
                               If jikkou = 0 Then
                             
                                 
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



Private Sub Picture13_Click()
Unload Me

End Sub

Private Sub Picture16_Click()

Dim newkojinid As String, newseikyuubi As String, newshurui As String, newshurui_no As String, newkin As String, newbankid As String


 
newkojinid = "03113"

newseikyuubi = "202105" ' Mid(newnyukinbi, 1, 6)

newshurui = "1"

newshurui_no = "35"  '35=番号表示　　39=着信拒否

newkin = "440"      '35＝440　　39＝550

newbankid = Space(1)

 If shousai_main_set(newkojinid, newseikyuubi, newshurui, newshurui_no, newkin, newbankid, "1") = -1 Then
            MsgBox "契約の更新に失敗しました。"
            Exit Sub
  End If
  
  
  MsgBox "契約の更新しました。"
  
  
End Sub
