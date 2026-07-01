VERSION 5.00
Begin VB.Form frmtouroku3 
   Caption         =   "åıçHéñä÷åW"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   8760
   StartUpPosition =   1  'µ∞≈∞ Ã´∞—ÇÃíÜâõ
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "è⁄ç◊"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6435
      Left            =   90
      TabIndex        =   4
      Top             =   180
      Width           =   8535
      Begin VB.TextBox txtbspass 
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
         IMEMode         =   2  'µÃ
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   33
         Top             =   5640
         Width           =   1740
      End
      Begin VB.TextBox txtkeiyakushamei 
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
         IMEMode         =   1  'µ›
         Left            =   2220
         MaxLength       =   50
         TabIndex        =   31
         Top             =   4968
         Width           =   4170
      End
      Begin VB.TextBox txtcloseday 
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
         IMEMode         =   2  'µÃ
         Left            =   6300
         MaxLength       =   10
         TabIndex        =   29
         Top             =   4296
         Width           =   1740
      End
      Begin VB.TextBox txtopenday 
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
         IMEMode         =   2  'µÃ
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   26
         Top             =   4296
         Width           =   1740
      End
      Begin VB.TextBox txtdonuadress 
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
         IMEMode         =   2  'µÃ
         Left            =   6030
         MaxLength       =   12
         TabIndex        =   24
         Top             =   2220
         Width           =   1740
      End
      Begin VB.TextBox txtvonuadress 
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
         IMEMode         =   2  'µÃ
         Left            =   2220
         MaxLength       =   12
         TabIndex        =   23
         Top             =   2244
         Width           =   1740
      End
      Begin VB.TextBox txtcl1 
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
         IMEMode         =   2  'µÃ
         Left            =   6030
         MaxLength       =   20
         TabIndex        =   20
         Top             =   1515
         Width           =   2340
      End
      Begin VB.TextBox txtlistbangou 
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
         IMEMode         =   2  'µÃ
         Left            =   2220
         MaxLength       =   3
         TabIndex        =   18
         Top             =   1560
         Width           =   1740
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cl"
         Height          =   315
         Left            =   6240
         TabIndex        =   17
         Top             =   930
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cl"
         Height          =   315
         Left            =   2850
         TabIndex        =   16
         Top             =   930
         Width           =   375
      End
      Begin VB.TextBox txtbikou 
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
         IMEMode         =   1  'µ›
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   14
         Top             =   3612
         Width           =   6420
      End
      Begin VB.TextBox txtkeiyakuid 
         Appearance      =   0  'Ã◊Øƒ
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'Ç»Çµ
         Enabled         =   0   'False
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
         IMEMode         =   2  'µÃ
         Left            =   1710
         TabIndex        =   11
         Top             =   330
         Width           =   1950
      End
      Begin VB.TextBox txtkoujikanryoubi 
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
         IMEMode         =   2  'µÃ
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   7
         Top             =   2928
         Width           =   1740
      End
      Begin VB.ComboBox cmbkoukubangou 
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1710
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   6
         Top             =   945
         Width           =   705
      End
      Begin VB.ComboBox cmbzumen 
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4800
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   5
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BSÉpÉXçHéñì˙ÅF"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   300
         TabIndex        =   34
         Top             =   5700
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CLå_ñÒé“ñºÅF"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   300
         TabIndex        =   32
         Top             =   5026
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CLâñÒì˙ÅF"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4530
         TabIndex        =   30
         Top             =   4350
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CLäJí ì˙ÅF"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   300
         TabIndex        =   27
         Top             =   4352
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "D-ONUÉAÉhÉåÉXÅF"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4080
         TabIndex        =   25
         Top             =   2235
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "V-ONUÉAÉhÉåÉXÅF"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   270
         TabIndex        =   22
         Top             =   2334
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÉNÉçÅ[ÉWÉÉÅ[î‘çÜÅF"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4080
         TabIndex        =   21
         Top             =   1560
         Width           =   1965
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÉäÉXÉgî‘çÜÅF"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   270
         TabIndex        =   19
         Top             =   1662
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "îıçlÅF"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   300
         TabIndex        =   15
         Top             =   3678
         Width           =   975
      End
      Begin VB.Label lblname 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   60
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "å_ñÒÇhÇcÅF"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   270
         TabIndex        =   12
         Top             =   390
         Width           =   1035
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         Caption         =   "çHéñäÆóπì˙ÅF"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   270
         TabIndex        =   10
         Top             =   3006
         Width           =   1935
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFFFF&
         Caption         =   "çHãÊî‘çÜÅF"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   270
         TabIndex        =   9
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label57 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ê}ñ ÅF"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3840
         TabIndex        =   8
         Top             =   960
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6060
      MouseIcon       =   "frmtouroku3.frx":0000
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmtouroku3.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   6960
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   7290
      MouseIcon       =   "frmtouroku3.frx":11C9
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmtouroku3.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   6960
      Width           =   480
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "çHéñäÆóπì˙ÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4440
      TabIndex        =   28
      Top             =   4485
      Width           =   1935
   End
   Begin VB.Label cmdtouroku 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "çXêV"
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
      Left            =   6090
      TabIndex        =   3
      Top             =   7470
      Width           =   480
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
      Left            =   7320
      TabIndex        =   2
      Top             =   7470
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   1095
      Left            =   120
      Top             =   6780
      Width           =   8490
   End
End
Attribute VB_Name = "frmtouroku3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cmbkoukubangou.ListIndex = -1
End Sub

Private Sub Command2_Click()
cmbzumen.ListIndex = -1
End Sub

Private Sub Form_Load()
 
     For cidcid = 1 To 6
        cmbkoukubangou.AddItem CStr(cidcid)
    Next cidcid
    
    
    cmbzumen.AddItem "A7"
    cmbzumen.AddItem "A8"
    cmbzumen.AddItem "B2"
    cmbzumen.AddItem "B3"
    cmbzumen.AddItem "B4"
    cmbzumen.AddItem "B5"
    cmbzumen.AddItem "B6"
    cmbzumen.AddItem "B7"
    cmbzumen.AddItem "B8"
    cmbzumen.AddItem "C2"
    cmbzumen.AddItem "C3"
    cmbzumen.AddItem "C4"
    cmbzumen.AddItem "C5"
    cmbzumen.AddItem "C6"
    cmbzumen.AddItem "C7"
    cmbzumen.AddItem "C8"
    cmbzumen.AddItem "D1"
    cmbzumen.AddItem "D2"
    cmbzumen.AddItem "D3"
    cmbzumen.AddItem "D4"
    cmbzumen.AddItem "D5"
    cmbzumen.AddItem "D6"
    cmbzumen.AddItem "D7"
    cmbzumen.AddItem "D8"
    cmbzumen.AddItem "E2"
    cmbzumen.AddItem "E3"
    cmbzumen.AddItem "E4"
     cmbzumen.AddItem "E5"
    cmbzumen.AddItem "E6"
    cmbzumen.AddItem "E7"
    cmbzumen.AddItem "F2"
    cmbzumen.AddItem "F3"
    cmbzumen.AddItem "F4"
    cmbzumen.AddItem "F5"
    cmbzumen.AddItem "F6"
    cmbzumen.AddItem "G2"
    cmbzumen.AddItem "G3"
    cmbzumen.AddItem "G4"
    cmbzumen.AddItem "G5"
    cmbzumen.AddItem "G6"
    cmbzumen.AddItem "G8"
    cmbzumen.AddItem "G9"
    cmbzumen.AddItem "G10"
     cmbzumen.AddItem "H3"
      cmbzumen.AddItem "H4"
     cmbzumen.AddItem "H5"
     cmbzumen.AddItem "H6"
     cmbzumen.AddItem "H7"
     cmbzumen.AddItem "H8"
     cmbzumen.AddItem "I3"
      cmbzumen.AddItem "I4"
       cmbzumen.AddItem "I5"
        cmbzumen.AddItem "I6"
         cmbzumen.AddItem "I7"
    cmbzumen.AddItem "J2"
    cmbzumen.AddItem "J3"
    cmbzumen.AddItem "J4"
     cmbzumen.AddItem "J5"
      cmbzumen.AddItem "K2"
       cmbzumen.AddItem "K3"
        cmbzumen.AddItem "K4"
         cmbzumen.AddItem "L4"
    
    
End Sub

Private Sub Picture13_Click()
Unload frmtouroku3
Set frmtouroku3 = Nothing

End Sub

Private Sub Picture7_Click()


    Dim s_koujikannryoubi As String
    Dim s_koukubangou As String
    Dim s_zumen As String
    
    Dim s_kojinid As String

    Dim s_name As String
    
    Dim s_bikou As String
    
    Dim s_listbangou As String
    
    Dim s_clbangou As String
    
    Dim s_v_onu As String

    Dim s_d_onu As String
    
    Dim s_open_day As String
    
    Dim s_close_day As String
    
    Dim s_keiyakushamei As String
    
    Dim s_bspass As String
    
With frmtouroku3

     If Trim(.txtkeiyakuid.Text) = "" Then
        MsgBox "ó\ä˙ÇπÇ éñóRÇ…ÇÊÇËÅAÉGÉâÅ[Ç™î≠ê∂ÇµÇ‹ÇµÇΩÅB"
        Exit Sub
    Else
        s_kojinid = Trim(.txtkeiyakuid.Text)
        s_name = Trim(.lblname.Caption)
    End If
    
    If Trim(.txtkoujikanryoubi.Text) = "" Then
        s_koujikannryoubi = Space(1)
    Else
    If Len(Trim(.txtkoujikanryoubi.Text)) = 10 Then
            s_koujikannryoubi = Format(Trim(.txtkoujikanryoubi.Text), "yyyymmdd")
        Else
            MsgBox "çHéñäÆóπì˙ÇÃì˙ïtÇÕÅu2004/02/01ÅvÇÃÇÊÇ§Ç…ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
    End If
    
    If Trim(.txtcl1.Text) = "" Then
            s_clbangou = Space(1)
    Else
        
            s_clbangou = Trim(.txtcl1.Text)
       
    End If
    

        If Trim(.txtvonuadress.Text) = "" Then
            s_v_onu = Space(1)
        Else
            s_v_onu = Trim(.txtvonuadress.Text)
        End If
        
        If Trim(.txtdonuadress.Text) = "" Then
            s_d_onu = Space(1)
        Else
            s_d_onu = Trim(.txtdonuadress.Text)
        End If
        
        

    If .cmbkoukubangou.ListIndex = -1 Then
        s_koukubangou = Space(1)
    Else
        s_koukubangou = Trim(.cmbkoukubangou.Text)
    End If
    
    If .cmbzumen.ListIndex = -1 Then
        s_zumen = Space(1)
    Else
        s_zumen = Trim(.cmbzumen.Text)
    End If


    If Trim(.txtbikou.Text) = "" Then
        s_bikou = Space(1)
    Else
        s_bikou = Trim(.txtbikou.Text)
    End If
    
    If Trim(.txtlistbangou.Text) = "" Then
        s_listbangou = Space(1)
    Else
        s_listbangou = Trim(.txtlistbangou.Text)
    End If
    
    If Trim(.txtopenday.Text) = "" Then
        s_open_day = Space(1)
    Else
        If Len(Trim(.txtopenday.Text)) = 10 Then
            s_open_day = Format(Trim(.txtopenday.Text), "yyyymmdd")
        Else
            MsgBox "CLäJí ì˙ÇÃì˙ïtÇÕÅu2004/02/01ÅvÇÃÇÊÇ§Ç…ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
    End If
    
    If Trim(.txtcloseday.Text) = "" Then
        s_close_day = Space(1)
    Else
        If Len(Trim(.txtcloseday.Text)) = 10 Then
            s_close_day = Format(Trim(.txtcloseday.Text), "yyyymmdd")
        Else
            MsgBox "CLâñÒì˙ÇÃì˙ïtÇÕÅu2004/02/01ÅvÇÃÇÊÇ§Ç…ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
    End If
    
    If Trim(.txtkeiyakushamei.Text) = "" Then
        s_keiyakushamei = Space(1)
    Else
        s_keiyakushamei = Trim(.txtkeiyakushamei.Text)
    End If
    
    If Trim(.txtbspass.Text) = "" Then
        s_bspass = Space(1)
    Else
        If Len(Trim(.txtbspass.Text)) = 10 Then
            s_bspass = Format(Trim(.txtbspass.Text), "yyyymmdd")
        Else
            MsgBox "BSÉpÉXçHéñì˙ÇÃì˙ïtÇÕÅu2004/02/01ÅvÇÃÇÊÇ§Ç…ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
    End If
    
    
    
 
        sql = "UPDATE kojin SET h_koujikanryoubi = '" & s_koujikannryoubi & "',h_koujibangou = '" & s_koukubangou & "'" & _
                ",h_zumen = '" & s_zumen & "',h_bikou = '" & s_bikou & "',h_listbangou = '" & s_listbangou & "'" & _
                ",cl_bangou = '" & s_clbangou & "',v_onu_adress = '" & s_v_onu & "',d_onu_adress = '" & s_d_onu & "'" & _
                ",cl_open_date = '" & s_open_day & "',cl_close_date = '" & s_close_day & "',cl_keiyaku_name = '" & s_keiyakushamei & "',h_bspass = '" & s_bspass & "'" & _
                  " WHERE kojinid = '" & s_kojinid & "'"
                  
                  
     

        On Error GoTo errft2
         db.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        
            'Log
        Dim sss_name As String
        sss_name = frmtouroku2.lblname.Caption
         Dim sss_id As String
        sss_id = frmtouroku2.lblid.Caption
        sagyou_msg = "å⁄ãqèÓïÒÅFå⁄ãq ID " & s_kojinid & " " & s_name & "ÇÃì‡óeÇïœçXÇµÇ‹ÇµÇΩÅB"
        log_sagyou sagyou_log_path, sagyou_msg, sss_id, sss_name
        
        
        MsgBox "ïœçXÇµÇ‹ÇµÇΩÅB"
        
        If hyouji_touroku(s_kojinid) = False Then
            Exit Sub
        End If
        
      Unload frmtouroku3
Set frmtouroku3 = Nothing

End With

Exit Sub
    

errft2:
    MsgBox "çXêVÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
    Exit Sub


End Sub
