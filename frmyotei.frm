VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmyotei 
   Caption         =   "予定管理"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   11535
      MouseIcon       =   "frmyotei.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmyotei.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   240
      Width           =   480
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   12847
      MouseIcon       =   "frmyotei.frx":11C9
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmyotei.frx":1AC8
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
      Left            =   14160
      MouseIcon       =   "frmyotei.frx":2392
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmyotei.frx":2C91
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   240
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "予定管理一覧"
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
      Height          =   9735
      Left            =   90
      TabIndex        =   0
      Top             =   1200
      Width           =   15015
      Begin MSFlexGridLib.MSFlexGrid yoteigrd 
         Height          =   9135
         Left            =   60
         TabIndex        =   1
         ToolTipText     =   "予定管理一覧です。"
         Top             =   390
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   16113
         _Version        =   393216
         Cols            =   17
         FixedCols       =   0
         BackColor       =   16777215
         BackColorBkg    =   16777215
         SelectionMode   =   1
      End
   End
   Begin VB.Label lblno 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9870
      TabIndex        =   8
      Top             =   510
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "中止"
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
      Left            =   11520
      TabIndex        =   7
      Top             =   750
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "削除"
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
      Index           =   3
      Left            =   12750
      TabIndex        =   6
      Top             =   750
      Width           =   675
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
      Left            =   14160
      TabIndex        =   3
      Top             =   750
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   1005
      Left            =   90
      Top             =   120
      Width           =   15000
   End
End
Attribute VB_Name = "frmyotei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkhihyouji_Click()
            Dim s_kojinid As String

            s_kojinid = Trim(frmtouroku2.lblid.Caption)
            
            yoteiset (s_kojinid)
End Sub

Private Sub Form_Load()

frmyotei.yoteigrd.Font.Size = 11

End Sub

Private Sub Picture13_Click()

    Unload frmyotei
    Set frmyotei = Nothing
    
    
End Sub

Private Sub Picture4_Click()

       
        
         If kiidou_mode <> 0 Then
            MsgBox "このモードでは使用できません。"
            Exit Sub
        End If
            
        Dim h_id As String, s_st As String
        
        frmyotei.yoteigrd.Col = 1
        s_st = frmyotei.yoteigrd.Text
        
        If s_st = "×" Then
             MsgBox "ステータスがすでに無効です。"
                Exit Sub
        End If
        

        frmyotei.yoteigrd.Col = 15
        h_id = frmyotei.yoteigrd.Text
        
        Dim s_name As String
        frmyotei.yoteigrd.Col = 5
        s_name = frmyotei.yoteigrd.Text
        
        
        Dim s_itsu As String, s_nanji As String
        
        s_itsu = Format(Now, "yyyymmdd")
        s_nanji = Format(Now, "hhnnss")
        
        
    sql = "select  * from yotei" & _
        " where yoteiid='" & h_id & "'"
    If FcSQlGet(rs, sql, PrMsg) = False Then
        MsgBox "中止したい項目を正確選択してから、実行してください。"
        Unload frmsakujo
        Set frmsakujo = Nothing
        Exit Sub
    Else
        msg = "予定ID『" & rs!yoteiid & " " & s_name & " 』を中止してもよいですか？"
       
        rs.Close
        
        res = MsgBox(msg, vbYesNo)
        If res = vbYes Then
           ' sql = "delete * from yotei where yoteiid ='" & h_id & "'"
            sql = "UPDATE yotei SET " & _
            "h_itsu = '" & s_itsu & "',h_nanji = '" & s_nanji & "',st = '1'" & _
            " WHERE yoteiid = '" & h_id & "'"
           
           On Error GoTo errsakujo
            db.Execute sql, dbSQLPassThrough
           On Error GoTo 0
           
           Dim s_kojinid As String

            s_kojinid = Trim(frmtouroku2.lblid.Caption)
            
            yoteiset (s_kojinid)
            
            
            MsgBox "中止しました"
        Else
            MsgBox "処理を中止しました。"
            Exit Sub
        End If
    End If
    
    Exit Sub
    
errsakujo:
    MsgBox "中止に失敗しました。"
  Exit Sub
  

End Sub

Private Sub Picture5_Click()

      
        
         If kiidou_mode <> 0 Then
            MsgBox "このモードでは使用できません。"
            Exit Sub
        End If
            
        Dim h_id As String, s_st As String
        
        frmyotei.yoteigrd.Col = 1
        s_st = frmyotei.yoteigrd.Text
        
        If s_st = "×" Then
             MsgBox "ステータスがすでに無効です。"
                Exit Sub
        End If

        frmyotei.yoteigrd.Col = 15
        h_id = frmyotei.yoteigrd.Text
        
          Dim s_name As String
        frmyotei.yoteigrd.Col = 5
        s_name = frmyotei.yoteigrd.Text
        
        
         Dim s_itsu As String, s_nanji As String, s_del As String
        
        s_itsu = Format(Now, "yyyymmdd")
        s_nanji = Format(Now, "hhnnss")
        
        s_del = s_itsu & " " & s_nanji
        
        
    sql = "select  * from yotei" & _
        " where yoteiid='" & h_id & "'"
    If FcSQlGet(rs, sql, PrMsg) = False Then
        MsgBox "削除したい項目を正確選択してから、実行してください。"
        Unload frmsakujo
        Set frmsakujo = Nothing
        Exit Sub
    Else
        msg = "予定ID『" & rs!yoteiid & " " & s_name & " 』を削除してもよいですか？"
       
        rs.Close
        
        res = MsgBox(msg, vbYesNo)
        If res = vbYes Then
           ' sql = "delete * from yotei where yoteiid ='" & h_id & "'"
            sql = "UPDATE yotei SET " & _
            "del = '" & s_del & "',st = '1'" & _
            " WHERE yoteiid = '" & h_id & "'"
           
           On Error GoTo errsakujo
            db.Execute sql, dbSQLPassThrough
           On Error GoTo 0
           
           Dim s_kojinid As String

            s_kojinid = Trim(frmtouroku2.lblid.Caption)
            
            yoteiset (s_kojinid)
            
            
            MsgBox "削除しました"
        Else
            MsgBox "中止しました。"
            Exit Sub
        End If
    End If
    
    Exit Sub
    
errsakujo:
    MsgBox "削除に失敗しました。"
  Exit Sub
  
End Sub
