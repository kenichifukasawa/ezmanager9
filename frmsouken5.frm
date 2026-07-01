VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmsouken5 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "未入金請求詳細"
   ClientHeight    =   10725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10725
   ScaleWidth      =   8430
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   7560
      MouseIcon       =   "frmsouken5.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmsouken5.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   9810
      Width           =   480
   End
   Begin MSFlexGridLib.MSFlexGrid gridnyuukinshousai 
      Height          =   9615
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   16960
      _Version        =   393216
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
      Left            =   7470
      TabIndex        =   2
      Top             =   10290
      Width           =   675
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   825
      Left            =   6540
      Top             =   9810
      Width           =   1710
   End
End
Attribute VB_Name = "frmsouken5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub gridnyuukinshousai_DblClick()
Unload Me

End Sub

Private Sub Picture5_Click()

If kiidou_mode <> 0 Then
    MsgBox "このモードでは使用できません。"
    Exit Sub
End If

If frmmain.lblkanri2.Visible = True Then
      MsgBox "サブモードでは更新できません。"
        Exit Sub
End If
Dim h_id As String

    gridnyuukinshousai.Col = 0
    h_id = gridnyuukinshousai.Text
    
    Dim resssres
    If h_id <> "" Then
        resssres = MsgBox("伝票番号" & h_id & "を削除してよろしいですか？", vbYesNo)
        If resssres = vbYes Then
            
             If kidou2 = False Then
                MsgBox "データベースファイルを開けませんでした。"
               Screen.MousePointer = 0
                Exit Sub
            End If
    
            sql = "delete * from shousai where shousaiid ='" & h_id & "'"
           On Error GoTo errsakujo
            db2.Execute sql, dbSQLPassThrough
           On Error GoTo 0
    
    
            MsgBox "削除しました。"
            
            Unload Me
            
        End If
    End If
   Exit Sub

errsakujo:
    MsgBox "削除に失敗しました。"
    Exit Sub


End Sub
