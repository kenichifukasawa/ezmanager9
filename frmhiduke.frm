VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmhiduke 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "日付選択"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4605
   StartUpPosition =   2  '画面の中央
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2220
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   64487425
      CurrentDate     =   36336
   End
   Begin VB.Label lbl1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "振り替え日の確認"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmhiduke.frx":0000
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   1995
   End
End
Attribute VB_Name = "frmhiduke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private shiteibi As String


Private Sub Form_Load()
If donohi <> 5 Then
    frmhiduke.MonthView1.Value = Date
End If
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)


shiteibi = Format(frmhiduke.MonthView1.Value, "yyyy/mm/dd")

If donohi = 1 Then
        frmnyuukin.txtshiteikoushinbi.Text = shiteibi
ElseIf donohi = 2 Then
        frmkeiyaku4.txtkaishibi.Text = shiteibi
ElseIf donohi = 3 Then
        frmkeiyaku5.txtkaishibi.Text = shiteibi
ElseIf donohi = 6 Then
        frmsouken6.txtnichiji.Text = shiteibi
ElseIf donohi = 7 Then
        frmsouken6.txtnichiji2.Text = shiteibi
ElseIf donohi = 5 Then
    msg = "「" & shiteibi & "」でよろしいですか？土日祭日ははずしてください。"
    res = MsgBox(msg, vbYesNo)
    If res = vbYes Then
        fdhiduke = Format(shiteibi, "yyyymmdd")
    Else
        MsgBox "もう一度指定してください"
        Exit Sub
    End If

End If

Unload Me

If donohi = 1 Then
         
ElseIf donohi = 2 Then
   'frmkeiyaku4.txtkeiyakusuu.SetFocus
End If
 
donohi = 0
End Sub

Private Sub MonthView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    shiteibi = Format(frmhiduke.MonthView1.Value, "yyyy/mm/dd")
If donohi = 1 Then
        frmnyuukin.txtshiteikoushinbi.Text = shiteibi
ElseIf donohi = 2 Then
        frmkeiyaku4.txtkaishibi.Text = shiteibi
ElseIf donohi = 3 Then
        frmkeiyaku5.txtkaishibi.Text = shiteibi
ElseIf donohi = 6 Then
        frmsouken6.txtnichiji.Text = shiteibi
ElseIf donohi = 7 Then
        frmsouken6.txtnichiji2.Text = shiteibi
ElseIf donohi = 5 Then
    msg = "「" & shiteibi & "」でよろしいですか？土日祭日ははずしてください。"
    res = MsgBox(msg, vbYesNo)
    If res = vbYes Then
        fdhiduke = Format(shiteibi, "yyyymmdd")
    Else
        MsgBox "もう一度指定してください"
        Exit Sub
    End If

End If

Unload Me

If donohi = 1 Then
         
ElseIf donohi = 2 Then
   'frmkeiyaku4.txtkeiyakusuu.SetFocus
End If
 
donohi = 0
    

End If
End Sub
