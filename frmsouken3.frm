VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmsouken3 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "è⁄ç◊"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   14610
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   13770
      MouseIcon       =   "frmsouken3.frx":0000
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmsouken3.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   8550
      Width           =   480
   End
   Begin MSFlexGridLib.MSFlexGrid gridseikyuu 
      Height          =   9315
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   4700
      _ExtentX        =   8281
      _ExtentY        =   16431
      _Version        =   393216
      Cols            =   3
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
   Begin MSFlexGridLib.MSFlexGrid gridnyuukin 
      Height          =   9315
      Left            =   4905
      TabIndex        =   1
      Top             =   90
      Width           =   4700
      _ExtentX        =   8281
      _ExtentY        =   16431
      _Version        =   393216
      Cols            =   3
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
   Begin MSFlexGridLib.MSFlexGrid gridminou 
      Height          =   8295
      Left            =   9780
      TabIndex        =   2
      Top             =   90
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   14631
      _Version        =   393216
      Cols            =   3
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "èIóπ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
      Left            =   13800
      TabIndex        =   4
      Top             =   9090
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   915
      Left            =   9780
      Top             =   8490
      Width           =   4650
   End
End
Attribute VB_Name = "frmsouken3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Picture13_Click()
Unload Me
End Sub
