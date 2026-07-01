VERSION 5.00
Begin VB.Form frmjoukyou2 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "その他の契約の集計"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   6720
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdend 
      Caption         =   "終了"
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
      Left            =   3240
      TabIndex        =   0
      ToolTipText     =   "終了します。"
      Top             =   5220
      Width           =   3255
   End
End
Attribute VB_Name = "frmjoukyou2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdend_Click()
Unload frmjoukyou2
Set frmjoukyou2 = Nothing

End Sub

Private Sub Form_Load()

End Sub
