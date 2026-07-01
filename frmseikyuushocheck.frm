VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmseikyuushocheck 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "êøãÅèëî≠çsÉ`ÉFÉbÉN"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   14910
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.Frame Frame1 
      Caption         =   "ï¿Ç◊ë÷Ç¶"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   675
      Left            =   120
      TabIndex        =   2
      Top             =   8040
      Width           =   12495
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "ï¿Ç◊ë÷Ç¶é¿çs"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   6
         Top             =   180
         Width           =   2000
      End
      Begin VB.OptionButton optnarabekae 
         Caption         =   "ëOâÒêøãÅì˙èáèá"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   8040
         TabIndex        =   5
         Top             =   240
         Width           =   2000
      End
      Begin VB.OptionButton optnarabekae 
         Caption         =   "ÉtÉäÉKÉièá"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   5970
         TabIndex        =   4
         Top             =   240
         Width           =   2000
      End
      Begin VB.OptionButton optnarabekae 
         Caption         =   "IDèá"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3900
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   2000
      End
      Begin VB.Label Label1 
         Caption         =   "âEÇÃçÄñ⁄ÇëIëÇµÅAÅuï¿Ç◊ë÷Ç¶é¿çsÉ{É^ÉìÅvÇâüÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "∑¨›æŸ"
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
      Left            =   12720
      TabIndex        =   1
      Top             =   8100
      Width           =   2000
   End
   Begin MSFlexGridLib.MSFlexGrid grdseikyuushocheck 
      Height          =   7755
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   13679
      _Version        =   393216
      Cols            =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmseikyuushocheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdend_Click()
Unload Me

End Sub

Private Sub Command1_Click()

Dim narabekae_code As Integer

If optnarabekae(0).Value = True Then
    narabekae_code = 0
ElseIf optnarabekae(1).Value = True Then
    narabekae_code = 1
ElseIf optnarabekae(2).Value = True Then
    narabekae_code = 2
End If
seikyuushohakkou (narabekae_code)

End Sub
