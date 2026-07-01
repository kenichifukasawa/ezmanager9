VERSION 5.00
Begin VB.Form frmsakujo 
   Caption         =   "削除選択"
   ClientHeight    =   1065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1065
   ScaleWidth      =   4065
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   600
      MouseIcon       =   "frmsakujo.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmsakujo.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   180
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3090
      MouseIcon       =   "frmsakujo.frx":11C9
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmsakujo.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   180
      Width           =   480
   End
   Begin VB.PictureBox Picture17 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1770
      MouseIcon       =   "frmsakujo.frx":2392
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmsakujo.frx":2C91
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblst 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  '実線
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblkingaku 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  '実線
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblnaiyou 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  '実線
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblid 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  '実線
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   390
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "通常削除"
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
      Index           =   6
      Left            =   330
      TabIndex        =   5
      Top             =   660
      Width           =   1005
   End
   Begin VB.Label seisei 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "予定削除"
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
      Left            =   1575
      TabIndex        =   3
      Top             =   660
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "戻る"
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
      Left            =   3150
      TabIndex        =   2
      Top             =   660
      Width           =   450
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   855
      Left            =   90
      Top             =   120
      Width           =   3870
   End
End
Attribute VB_Name = "frmsakujo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture13_Click()

    Unload frmsakujo
    Set frmsakujo = Nothing
    
End Sub

Private Sub Picture17_Click()
  
    
    If Trim(lblst.Caption) <> "稼動" Then
        MsgBox "「稼動」以外の契約は、予定削除できません。"
        Unload frmsakujo
        Exit Sub
    End If

    With frmsakujo2
        .lblkeiyakuid.Caption = Trim(frmsakujo.lblid.Caption)
        .lblkeiyakumei.Caption = Trim(frmsakujo.lblnaiyou.Caption)
        .lblkingaku.Caption = Trim(frmsakujo.lblkingaku.Caption)
  
    Unload frmsakujo
  
    Set frmsakujo = Nothing
  
  
        .Show 1
    End With
    
    
End Sub

Private Sub Picture8_Click()

 

Dim s_id As String, kyk As String, kyk2 As String, kyk3 As String
If kenchan = 0 Then
    If frmtouroku2.lbljoukyou.Caption = "不可" Then
        MsgBox "〆中のデータにつき、削除できません。"
        Unload frmsakujo
        Set frmsakujo = Nothing
        Exit Sub
    End If
End If


Dim s_name As String


With frmtouroku2
    .keiyakugrd.Col = 0
    s_id = .keiyakugrd.Text

    sql = "select  keiyaku.*, shurui.shuruimei from shurui RIGHT JOIN keiyaku ON shurui.shuruiid = keiyaku.shurui" & _
        " where keiyaku.keiyakuid='" & s_id & "'"
    If FcSQlGet(rs, sql, PrMsg) = False Then
        MsgBox "削除したい項目を正確選択してから、実行してください。"
        Exit Sub
    Else
        msg = "契約『" & rs!shuruimei & "』を削除してもよいですか？"
        s_name = rs!shuruimei
        kyk = rs!kadou
        kyk2 = rs!kingaku
        kyk3 = rs!kaishi
        rs.Close
        res = MsgBox(msg, vbYesNo)
        If res = vbYes Then
            '削除
            sql = "delete * from keiyaku where keiyakuid ='" & s_id & "'"
           On Error GoTo errsakujo
            db.Execute sql, dbSQLPassThrough
           On Error GoTo 0
            If a_shukeiyaku_sakujo(.lblid.Caption, kyk2, kyk) = False Then
                MsgBox "料金の更新に失敗しました。"
                Exit Sub
            End If
            
            '最終更新日更新
            
            Dim s_kojinid As String

            s_kojinid = frmtouroku2.lblid.Caption
            
            Dim s_kyounohi As String
            
            s_kyounohi = Format(Date, "yyyy/mm/dd")

             sql = "update kojin set  saishuukoushinbi='" & s_kyounohi & "'" & _
                       " where kojinid ='" & s_kojinid & "'"
            
            On Error GoTo errsakujo2
            db.Execute sql, dbSQLPassThrough
            On Error GoTo 0
    
    
            'Log
            Dim ss_name As String
            ss_name = frmtouroku2.lblname.Caption
            Dim ss_id As String
            ss_id = frmtouroku2.lblid.Caption
    
            sagyou_msg = "主契約：契約 ID " & s_id & " " & s_name & "の契約を削除しました。"
            log_sagyou sagyou_log_path, sagyou_msg, ss_id, ss_name
            
            
            keiyakuset (.lblid.Caption)
            MsgBox "削除しました"
        Else
            MsgBox "中止しました。"
            Exit Sub
        End If
    End If
End With
Exit Sub

errsakujo:
    MsgBox "削除に失敗しました。"
    Exit Sub

errsakujo2:
    MsgBox "削除に失敗しました。"
    Exit Sub



End Sub
