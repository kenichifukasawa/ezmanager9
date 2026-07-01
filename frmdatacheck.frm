VERSION 5.00
Begin VB.Form frmdatacheck 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "ÉfÅ[É^É`ÉFÉbÉN"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   7815
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.CommandButton Command1 
      Caption         =   "äÆóπ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6060
      TabIndex        =   3
      Top             =   1590
      Width           =   1605
   End
   Begin VB.Frame Frame1 
      Caption         =   "èÛãµ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.Label lblshinkou 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   270
         TabIndex        =   2
         Top             =   750
         Width           =   7005
      End
      Begin VB.Label lbljoukyou 
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   270
         TabIndex        =   1
         Top             =   270
         Width           =   7005
      End
   End
End
Attribute VB_Name = "frmdatacheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()

End Sub
