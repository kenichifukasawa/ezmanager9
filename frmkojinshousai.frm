VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmkojinshousai 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "å¬êlèÓïÒè⁄ç◊"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleMode       =   0  '’∞ªﬁ∞
   ScaleWidth      =   16607.76
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.Frame Frame5 
      Caption         =   "ÉPÅ[ÉuÉãÉâÉCÉì"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   8070
      TabIndex        =   41
      Top             =   90
      Width           =   6945
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   6060
         MouseIcon       =   "frmkojinshousai.frx":0000
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkojinshousai.frx":08FF
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   48
         Top             =   1215
         Width           =   480
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   6060
         MouseIcon       =   "frmkojinshousai.frx":11C9
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkojinshousai.frx":1AC8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   47
         ToolTipText     =   "ìoò^èàóùÇçsÇ¢Ç‹Ç∑ÅB"
         Top             =   450
         Width           =   480
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   6090
         MouseIcon       =   "frmkojinshousai.frx":2392
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkojinshousai.frx":2C91
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   46
         Top             =   2040
         Width           =   480
      End
      Begin MSFlexGridLib.MSFlexGrid grdcl 
         Height          =   2565
         Left            =   240
         TabIndex        =   42
         Top             =   330
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   4524
         _Version        =   393216
         SelectionMode   =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00FFFFFF&
         Caption         =   "ìoò^"
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
         Height          =   255
         Index           =   4
         Left            =   5850
         TabIndex        =   45
         Top             =   930
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00FFFFFF&
         Caption         =   "çÌèú"
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
         Height          =   255
         Index           =   2
         Left            =   5880
         TabIndex        =   44
         Top             =   2550
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ïœçX"
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
         Index           =   0
         Left            =   6090
         TabIndex        =   43
         Top             =   1725
         Width           =   480
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   2625
         Left            =   5850
         Top             =   270
         Width           =   930
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "PPVégópèÛãµ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3795
      Left            =   60
      TabIndex        =   39
      Top             =   7290
      Width           =   15045
      Begin MSFlexGridLib.MSFlexGrid gridppvrireki 
         Height          =   3375
         Left            =   120
         TabIndex        =   40
         Top             =   300
         Width           =   14835
         _ExtentX        =   26167
         _ExtentY        =   5953
         _Version        =   393216
         Cols            =   12
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "STBÇÃä«óùî‘çÜÇ∆äÌäBî‘çÜÇ∆BCAS/CCAS ÇÃî‘çÜÇ∆äePPVÇÃìoò^ÇÃóLñ≥"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2745
      Left            =   60
      TabIndex        =   31
      Top             =   4470
      Width           =   15045
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   14130
         MouseIcon       =   "frmkojinshousai.frx":355B
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkojinshousai.frx":3E5A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   35
         Top             =   975
         Width           =   480
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   14130
         MouseIcon       =   "frmkojinshousai.frx":4724
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkojinshousai.frx":5023
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   34
         ToolTipText     =   "ìoò^èàóùÇçsÇ¢Ç‹Ç∑ÅB"
         Top             =   210
         Width           =   480
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   14130
         MouseIcon       =   "frmkojinshousai.frx":58ED
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkojinshousai.frx":61EC
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   33
         Top             =   1740
         Width           =   480
      End
      Begin MSFlexGridLib.MSFlexGrid grdstb 
         Height          =   2325
         Left            =   120
         TabIndex        =   32
         Top             =   300
         Width           =   13575
         _ExtentX        =   23945
         _ExtentY        =   4101
         _Version        =   393216
         Cols            =   6
         SelectionMode   =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "STBïœçX"
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
         Index           =   8
         Left            =   13890
         TabIndex        =   38
         Top             =   1470
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00FFFFFF&
         Caption         =   "STBìoò^"
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
         Height          =   255
         Index           =   7
         Left            =   13890
         TabIndex        =   37
         Top             =   690
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00FFFFFF&
         Caption         =   "STBçÌèú"
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
         Height          =   255
         Index           =   3
         Left            =   13890
         TabIndex        =   36
         Top             =   2220
         Width           =   945
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   2415
         Left            =   13830
         Top             =   210
         Width           =   1050
      End
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   14400
      MouseIcon       =   "frmkojinshousai.frx":6AB6
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmkojinshousai.frx":73B5
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   29
      Top             =   3420
      Width           =   480
   End
   Begin VB.Frame Frame10 
      Caption         =   "VoIPÅ@ÇÃî‘çÜÇ∆äJénå_ñÒì˙"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4335
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   5055
      Begin VB.TextBox txtvoiphiduke 
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
         Index           =   9
         Left            =   2910
         MaxLength       =   10
         TabIndex        =   28
         Top             =   3930
         Width           =   1785
      End
      Begin VB.TextBox txtvoiphiduke 
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
         Index           =   8
         Left            =   2910
         MaxLength       =   10
         TabIndex        =   27
         Top             =   3525
         Width           =   1785
      End
      Begin VB.TextBox txtvoiphiduke 
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
         Index           =   7
         Left            =   2910
         MaxLength       =   10
         TabIndex        =   26
         Top             =   3120
         Width           =   1785
      End
      Begin VB.TextBox txtvoiphiduke 
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
         Index           =   6
         Left            =   2910
         MaxLength       =   10
         TabIndex        =   25
         Top             =   2715
         Width           =   1785
      End
      Begin VB.TextBox txtvoiphiduke 
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
         Index           =   5
         Left            =   2910
         MaxLength       =   10
         TabIndex        =   24
         Top             =   2310
         Width           =   1785
      End
      Begin VB.TextBox txtvoiphiduke 
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
         Index           =   4
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   23
         Top             =   1912
         Width           =   1785
      End
      Begin VB.TextBox txtvoiphiduke 
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
         Index           =   3
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   22
         Top             =   1509
         Width           =   1785
      End
      Begin VB.TextBox txtvoiphiduke 
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
         Index           =   2
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1106
         Width           =   1785
      End
      Begin VB.TextBox txtvoiphiduke 
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
         Index           =   1
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   20
         Top             =   703
         Width           =   1785
      End
      Begin VB.TextBox txtvoiphiduke 
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
         Index           =   0
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   19
         Top             =   300
         Width           =   1785
      End
      Begin VB.TextBox txtvoip 
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
         Index           =   9
         Left            =   210
         MaxLength       =   11
         TabIndex        =   18
         Top             =   3930
         Width           =   2500
      End
      Begin VB.TextBox txtvoip 
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
         Index           =   8
         Left            =   210
         MaxLength       =   11
         TabIndex        =   17
         Top             =   3525
         Width           =   2500
      End
      Begin VB.TextBox txtvoip 
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
         Index           =   7
         Left            =   210
         MaxLength       =   11
         TabIndex        =   16
         Top             =   3120
         Width           =   2500
      End
      Begin VB.TextBox txtvoip 
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
         Index           =   6
         Left            =   210
         MaxLength       =   11
         TabIndex        =   15
         Top             =   2715
         Width           =   2500
      End
      Begin VB.TextBox txtvoip 
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
         Index           =   5
         Left            =   210
         MaxLength       =   11
         TabIndex        =   14
         Top             =   2310
         Width           =   2500
      End
      Begin VB.TextBox txtvoip 
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
         Index           =   4
         Left            =   210
         MaxLength       =   11
         TabIndex        =   13
         Top             =   1912
         Width           =   2500
      End
      Begin VB.TextBox txtvoip 
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
         Index           =   3
         Left            =   210
         MaxLength       =   11
         TabIndex        =   12
         Top             =   1509
         Width           =   2500
      End
      Begin VB.TextBox txtvoip 
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
         Index           =   2
         Left            =   210
         MaxLength       =   11
         TabIndex        =   11
         Top             =   1106
         Width           =   2500
      End
      Begin VB.TextBox txtvoip 
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
         Index           =   1
         Left            =   210
         MaxLength       =   11
         TabIndex        =   10
         Top             =   703
         Width           =   2500
      End
      Begin VB.TextBox txtvoip 
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
         Index           =   0
         Left            =   210
         MaxLength       =   11
         TabIndex        =   9
         Top             =   300
         Width           =   2500
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "π∞ÃﬁŸ”√ﬁ—MAC±ƒﬁ⁄Ω"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2115
      Left            =   5310
      TabIndex        =   4
      Top             =   2250
      Width           =   2595
      Begin VB.TextBox txtit1 
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
         Left            =   90
         TabIndex        =   7
         Top             =   390
         Width           =   2300
      End
      Begin VB.TextBox txtit2 
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
         Left            =   90
         TabIndex        =   6
         Top             =   960
         Width           =   2300
      End
      Begin VB.TextBox txtit3 
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
         Left            =   90
         TabIndex        =   5
         Top             =   1530
         Width           =   2300
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Œ∞—≈Ÿ¿∞–≈ŸID"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1995
      Left            =   5280
      TabIndex        =   0
      Top             =   60
      Width           =   2625
      Begin VB.TextBox txtht3 
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
         Left            =   150
         TabIndex        =   3
         Top             =   1440
         Width           =   2300
      End
      Begin VB.TextBox txtht2 
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
         Left            =   150
         TabIndex        =   2
         Top             =   915
         Width           =   2300
      End
      Begin VB.TextBox txtht1 
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
         Left            =   150
         TabIndex        =   1
         Top             =   390
         Width           =   2300
      End
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
      Left            =   14430
      TabIndex        =   30
      Top             =   3900
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   855
      Left            =   8100
      Top             =   3360
      Width           =   6930
   End
End
Attribute VB_Name = "frmkojinshousai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Picture10_Click()

Dim h_id As String

With frmstb
    frmkojinshousai.grdstb.Col = 0
    h_id = grdstb.Text
    sql = "select * from stb where stbkanriid = '" & h_id & "'"
    
    Set rs = db.OpenRecordset(sql, dbOpenDynaset, dbReadOnly)
    If FcSQlGet(rs, sql, PrMsg) = False Then
        MsgBox "ïœçXÇµÇΩÇ¢çÄñ⁄Çê≥ämëIëÇµÇƒÇ©ÇÁÅAé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        Exit Sub
    Else
        .txtstbkanriid.Text = rs!stbkanriid
        .txtstbno.Text = rs!koushinbistbid
        .txtbcasno.Text = rs!bcasno
        .txtccasno.Text = rs!ccasno
        If rs!bcasumu = "0" Then 'å≈íË
            .chkb.Value = 0
        Else
            .chkb.Value = 1
        End If
        If rs!ccasumu = "0" Then 'å≈íË
            .chkc.Value = 0
        Else
            .chkc.Value = 1
        End If
        rs.Close
    End If
    .Label1(15).Caption = "ïœçX"
    frmstb.lblkojinid.Caption = frmtouroku2.lblid.Caption
    .Show 1
End With

End Sub

Private Sub Picture13_Click()
Unload Me
End Sub

Private Sub Picture2_Click()
With frmcl
    If soukatsu(0, "renban", 17, "", 1, 1) = False Then
            MsgBox "ê›íËÇÃì«Ç›çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
    End If
    If reg = "0" Then
        reg = ""
    End If
    If reg = vbNullString Then
        .lblclid.Caption = "0001"
    ElseIf reg = "9999" Then
        MsgBox "Ç±ÇÍà»è„ÇÃCLÇÃèÓïÒÇÃçÏê¨ÇÕÇ≈Ç´Ç‹ÇπÇÒÅBä«óùé“Ç…òAóçÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        Exit Sub
    Else
        .lblclid.Caption = Format$(reg, "000#")
    End If
    frmcl.lblkojinid.Caption = frmtouroku2.lblid.Caption
    .Label1(15).Caption = "ìoò^"
    .Show 1
End With
End Sub

Private Sub Picture3_Click()
Dim nhk_msg As String, nhk_res, h_id As String

With frmcl
    frmkojinshousai.grdcl.Col = 0
    h_id = grdcl.Text
    sql = "select * from cl where clid = '" & h_id & "'"
    
    Set rs = db.OpenRecordset(sql, dbOpenDynaset, dbReadOnly)
    If FcSQlGet(rs, sql, PrMsg) = False Then
        MsgBox "çÌèúÇµÇΩÇ¢çÄñ⁄Çê≥ämëIëÇµÇƒÇ©ÇÁÅAé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        Exit Sub
    Else
            nhk_msg = "CLÇÃìoò^èÓïÒÇçÌèúÇµÇƒÇ‡ó«Ç¢Ç≈Ç∑Ç©ÅH"
             nhk_res = MsgBox(nhk_msg, vbYesNo)
             If nhk_res = vbYes Then
                sql = "delete * from cl where clid ='" & h_id & "'"
                On Error GoTo errsakujo_stb
                 db.Execute sql, dbSQLPassThrough
                On Error GoTo 0
            Else
                MsgBox "íÜé~ÇµÇ‹ÇµÇΩÅB"
                Exit Sub
            End If
            cl_grid_set (frmtouroku2.lblid.Caption)

    End If
End With
Exit Sub

errsakujo_stb:
    MsgBox "CLÇÃèÓïÒÇÃçÌèúÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
    Exit Sub
End Sub

Private Sub Picture4_Click()
Dim h_id As String

With frmcl
    frmkojinshousai.grdcl.Col = 0
    h_id = grdcl.Text
    sql = "select * from cl where clid = '" & h_id & "'"
    
    Set rs = db.OpenRecordset(sql, dbOpenDynaset, dbReadOnly)
    If FcSQlGet(rs, sql, PrMsg) = False Then
        MsgBox "ïœçXÇµÇΩÇ¢çÄñ⁄Çê≥ämëIëÇµÇƒÇ©ÇÁÅAé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        Exit Sub
    Else
        .lblclid.Caption = rs!clid
        .txttelno.Text = rs!tel
        rs.Close
    End If
    .Label1(15).Caption = "ïœçX"
    frmcl.lblkojinid.Caption = frmtouroku2.lblid.Caption
    .Show 1
End With
End Sub

Private Sub Picture5_Click()

Dim nhk_msg As String, nhk_res, h_id As String

With frmstb
    frmkojinshousai.grdstb.Col = 0
    h_id = grdstb.Text
    sql = "select * from stb where stbkanriid = '" & h_id & "'"
    
    Set rs = db.OpenRecordset(sql, dbOpenDynaset, dbReadOnly)
    If FcSQlGet(rs, sql, PrMsg) = False Then
        MsgBox "çÌèúÇµÇΩÇ¢çÄñ⁄Çê≥ämëIëÇµÇƒÇ©ÇÁÅAé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        Exit Sub
    Else
            nhk_msg = "STBÇÃìoò^èÓïÒÇçÌèúÇµÇƒÇ‡ó«Ç¢Ç≈Ç∑Ç©ÅH"
             nhk_res = MsgBox(nhk_msg, vbYesNo)
             If nhk_res = vbYes Then
                sql = "delete * from stb where stbkanriid ='" & h_id & "'"
                On Error GoTo errsakujo_stb
                 db.Execute sql, dbSQLPassThrough
                On Error GoTo 0
            Else
                MsgBox "íÜé~ÇµÇ‹ÇµÇΩÅB"
                Exit Sub
            End If
            stb_grid_set2 (frmtouroku2.lblid.Caption)

    End If
End With
Exit Sub

errsakujo_stb:
    MsgBox "ÇrÇsÇaÇÃèÓïÒÇÃçÌèúÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
    Exit Sub


End Sub

Private Sub Picture9_Click()
With frmstb
    If soukatsu(0, "renban", 11, "", 1, 1) = False Then
            MsgBox "ê›íËÇÃì«Ç›çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
    End If
    If reg = "0" Then
        reg = ""
    End If
    If reg = vbNullString Then
        .txtstbkanriid.Text = "0001"
    ElseIf reg = "9999" Then
        MsgBox "Ç±ÇÍà»è„ÇÃÇªÇÃëºÇÃå_ñÒçÄñ⁄ÇÃçÏê¨ÇÕÇ≈Ç´Ç‹ÇπÇÒÅBä«óùé“Ç…òAóçÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        Exit Sub
    Else
        .txtstbkanriid.Text = Format$(reg, "000#")
    End If
    frmstb.lblkojinid.Caption = frmtouroku2.lblid.Caption
    .Label1(15).Caption = "ìoò^"
    .Show 1
End With
End Sub
