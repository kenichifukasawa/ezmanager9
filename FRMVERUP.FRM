VERSION 5.00
Begin VB.Form frmverup 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ﾊﾞｰｼﾞｮﾝｱｯﾌﾟ中...."
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   7320
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdend 
      Caption         =   "開始"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   7920
      Width           =   1215
   End
   Begin VB.ListBox lstver 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6300
      Left            =   240
      TabIndex        =   0
      Top             =   1140
      Width           =   6675
   End
   Begin VB.Label Label1 
      Caption         =   $"frmverup.frx":0000
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   6675
   End
End
Attribute VB_Name = "frmverup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private honto As Integer

Private Sub cmdend_Click()
If cmdend.Caption = "開始" Then
    honto = 0
        If ver(exever) = True Then
            If soukatsu(1, "startup", 5, exever, 1, 1) = False Then
                frmverup.lstver.AddItem "設定の書き込みに失敗しました。"
                honto = 1
            End If
        Else
            frmverup.lstver.AddItem "エラーにより強制終了します。"
            honto = 1
        End If
        cmdend.Enabled = True
        cmdend.Caption = "終了"
Else
        Unload Me
        If honto = 1 Then
            MsgBox "元の状態に復元し,いったんシステムを終了します。"
            Kill VBNAME
            Kill VBNAME2
            On Error GoTo errfukugen1
                FileCopy backupsaki1, VBNAME
                DoEvents
                FileCopy backupsaki2, VBNAME2
                DoEvents
            On Error GoTo 0
            db.Close
            End
        End If
End If


Exit Sub

errfukugen1:
    MsgBox "復元に失敗しました。管理者までご連絡を...。"
    End
End Sub

