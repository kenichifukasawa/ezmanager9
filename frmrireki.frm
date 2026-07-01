VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmrireki 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "入金履歴"
   ClientHeight    =   9960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   15270
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      Caption         =   "今月の仮入金情報"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   120
      TabIndex        =   7
      Top             =   7740
      Width           =   6765
      Begin VB.Label Label4 
         Caption         =   "円"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   6210
         TabIndex        =   12
         Top             =   450
         Width           =   345
      End
      Begin VB.Label lblmaebi 
         Alignment       =   2  '中央揃え
         BackColor       =   &H8000000E&
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
         Height          =   255
         Left            =   1470
         TabIndex        =   9
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label Label4 
         Caption         =   "入金日："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "入金金額："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   3420
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblkingaku 
         Alignment       =   2  '中央揃え
         BackColor       =   &H8000000E&
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
         Height          =   255
         Left            =   4620
         TabIndex        =   8
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   4920
      MouseIcon       =   "frmrireki.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmrireki.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   8910
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6030
      MouseIcon       =   "frmrireki.frx":11C9
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmrireki.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   8910
      Width           =   480
   End
   Begin MSFlexGridLib.MSFlexGrid fggridrireki 
      Height          =   7485
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   13203
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid gridnyuukinshousai 
      Height          =   9615
      Left            =   6960
      TabIndex        =   4
      Top             =   90
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   16960
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "変更"
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
      Left            =   4950
      TabIndex        =   6
      Top             =   9420
      Width           =   495
   End
   Begin VB.Label lblkojinid 
      Height          =   225
      Left            =   450
      TabIndex        =   3
      Top             =   8850
      Visible         =   0   'False
      Width           =   795
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
      Left            =   6060
      TabIndex        =   2
      Top             =   9420
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   1035
      Left            =   60
      Top             =   8700
      Width           =   6780
   End
End
Attribute VB_Name = "frmrireki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub fggridrireki_DblClick()

Dim ssshiduhidu As String

If Trim(frmrireki.lblkojinid.Caption) = "" Then
    MsgBox "個人IDを参照できません。"
    Exit Sub
End If
With frmrireki.fggridrireki
On Error GoTo errhidu
    .Col = 1
    ssshiduhidu = Format(.Text, "yyyymmdd")
On Error GoTo 0
End With

If kidou2 = False Then
    MsgBox "データベースファイルを開けませんでした。"
    Exit Sub
End If


nyuukin_shousai_set Trim(frmrireki.lblkojinid.Caption), ssshiduhidu


Exit Sub

errhidu:
    MsgBox "調べたい項目を正確にクイックしてください。"
    Exit Sub

End Sub

Private Sub Form_Load()
If kenchan = 1 Then
    Picture10.Enabled = True
Else
    Picture10.Enabled = False
End If

End Sub

Private Sub Picture10_Click()
Dim sa_id As String, pres, newhihihi As String
Dim newhihihi2 As String, shiteirow As Integer

With fggridrireki
    
    .Col = 0
    sa_id = .Text
    If Trim(sa_id) = "" Then
        MsgBox "変更する伝票がないか、指定されていません。"
        Exit Sub
    End If
    .Col = 3
    If .Text = "金融機関" Then
        MsgBox "金融機関の入金に関して、日付は変更できません。"
        Exit Sub
    End If
On Error GoTo errhidu
    .Col = 1
    newhihihi2 = Format(.Text, "yyyymmdd")
On Error GoTo 0

    If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If
    pres = InputBox("新しい日付を入力してください。（例　2006/02/04)", "入金日変更")
    If pres <> "" Then
        If Len(pres) <> 10 Then
            MsgBox "（例　2006/02/04)に習って、半角１０バイトで入力してください。"
            db2.Close
           
            Exit Sub
        End If
        On Error GoTo errnichi
            newhihihi = Format(pres, "yyyymmdd")
        On Error GoTo 0
         sql = "UPDATE koushinkiroku SET koushinbi = '" & newhihihi & "'" & _
                      " WHERE id = '" & sa_id & "'"
        On Error GoTo errsakujo
         db2.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        
         sql = "UPDATE shousai SET keinyu = '" & newhihihi & "'" & _
                      " WHERE kojinid = '" & Trim(frmrireki.lblkojinid.Caption) & "'" & _
                      " and keinyu='" & newhihihi2 & "' and keist='1'"
        On Error GoTo errsakujo
         db2.Execute sql, dbSQLPassThrough
        On Error GoTo 0
    
        If .Row = 1 Then
             sql = "update kojin set maebi='" & newhihihi & "'" & _
                    " where kojinid='" & Trim(frmrireki.lblkojinid.Caption) & "'"
            On Error GoTo errsakujo
             db.Execute sql, dbSQLPassThrough
            On Error GoTo 0
        End If
    
        nyuukinrirekiset (1)
        
    If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If

        nyuukin_shousai_set Trim(frmrireki.lblkojinid.Caption), newhihihi
        
        
        MsgBox "変更が完了しました。一度、EZManagerを再起動し、通常パスワードで起動後、個人情報、入金情報を再度確認をお願いします。"
        
    Else
        MsgBox "中止しました。"
    End If
End With
Exit Sub
errhidu:
    MsgBox "変更したい項目を正確にクイックしてください。"
    Exit Sub
errnichi:
    MsgBox "（例　2006/02/04)に習って、半角１０バイトで入力してください。"
    db2.Close
    Exit Sub
errsakujo:
    MsgBox "変更に失敗しました。"
    db2.Close
    Exit Sub

End Sub

Private Sub Picture13_Click()
Unload frmrireki
Set frmrireki = Nothing

End Sub

Private Sub Picture7_Click()
Dim sa_id As String, pres

MsgBox "Yet・・・"
Exit Sub

With fggridrireki
    
    .Col = 0
    sa_id = .Text
    If Trim(sa_id) = "" Then
        MsgBox "削除する伝票がないか、指定されていません。"
        Exit Sub
    End If
    
    If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If
    pres = MsgBox("本当に削除してよいですか？", vbYesNo)
    If pres = vbYes Then
       sql = "delete * from koushinkiroku where id ='" & sa_id & "'"
        On Error GoTo errsakujo
         db2.Execute sql, dbSQLPassThrough
        On Error GoTo 0
    
        nyuukinrirekiset (1)
    Else
        MsgBox "中止しました。"
    End If
End With
Exit Sub

errsakujo:
    MsgBox "伝票の削除に失敗しました。"
    db2.Close
    Exit Sub

End Sub
