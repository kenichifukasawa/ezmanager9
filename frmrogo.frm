VERSION 5.00
Begin VB.Form frmrogo 
   BorderStyle     =   3  'å≈íË¿ﬁ≤±€∏ﬁ
   ClientHeight    =   5460
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   10395
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmrogo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5220
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   10080
      Begin VB.Label lblshinkou 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3900
         TabIndex        =   5
         Top             =   3870
         Width           =   5565
      End
      Begin VB.Image Image1 
         Height          =   2400
         Left            =   570
         Picture         =   "frmrogo.frx":000C
         Stretch         =   -1  'True
         Top             =   1530
         Width           =   2430
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'âEëµÇ¶
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8640
         TabIndex        =   2
         Top             =   2880
         Width           =   105
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "EZ Manager 9"
         BeginProperty Font 
            Name            =   "Century"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   3660
         TabIndex        =   4
         Top             =   1890
         Width           =   4440
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'âEëµÇ¶
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copyright(C) 2000-2023 Plot Office LLC"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3090
         TabIndex        =   1
         Top             =   4830
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "CATVä«óùÉVÉXÉeÉÄ"
         BeginProperty Font 
            Name            =   "ÇlÇr Çoñæí©"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3570
         TabIndex        =   3
         Top             =   900
         Width           =   3105
      End
   End
End
Attribute VB_Name = "frmrogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = " ﬁ∞ºﬁÆ› " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

