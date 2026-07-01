VERSION 5.00
Begin VB.Form frmkensaku 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "åüçıâÊñ "
   ClientHeight    =   10560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   12345
   StartUpPosition =   1  'µ∞≈∞ Ã´∞—ÇÃíÜâõ
   Begin VB.CheckBox chkend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "åüçıéûîÒï\é¶ê›íËÇÃÇ‡ÇÃÇåüçıÇµÇ»Ç¢"
      Height          =   255
      Left            =   6420
      TabIndex        =   54
      Top             =   10200
      Value           =   1  '¡™Ø∏
      Width           =   3255
   End
   Begin VB.Frame Frame20 
      Caption         =   "éxï•ï˚ñ@ÇÃëIë"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   6300
      TabIndex        =   52
      Top             =   8400
      Width           =   5895
      Begin VB.ComboBox cmbshiharai 
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
         Left            =   1920
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   53
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.Frame Frame19 
      Caption         =   "å_ñÒé“äøéöÇÃì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   6300
      TabIndex        =   50
      Top             =   7380
      Width           =   5895
      Begin VB.TextBox txtname 
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
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   51
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.Frame Frame18 
      Caption         =   "å_ñÒå˚ç¿ÉtÉäÉKÉiÇÃì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   210
      TabIndex        =   48
      Top             =   9480
      Width           =   5895
      Begin VB.TextBox txtkouzaname 
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
         IMEMode         =   6  'îºäp∂¿∂≈
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   49
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   11220
      MouseIcon       =   "frmkensaku.frx":0000
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmkensaku.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   45
      Top             =   9660
      Width           =   480
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   9990
      MouseIcon       =   "frmkensaku.frx":11C9
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmkensaku.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   44
      Top             =   9660
      Width           =   480
   End
   Begin VB.Frame Frame17 
      Caption         =   "NHKå_ñÒî‘çÜÇÃì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   6300
      TabIndex        =   39
      Top             =   5330
      Width           =   5895
      Begin VB.CheckBox chknhk 
         Caption         =   "å_ñÒÇ™Ç†ÇÈ"
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
         Height          =   345
         Left            =   420
         TabIndex        =   18
         Top             =   390
         Width           =   2055
      End
      Begin VB.TextBox txtnhkno 
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
         Left            =   2460
         MaxLength       =   30
         TabIndex        =   19
         Top             =   360
         Width           =   3045
      End
   End
   Begin VB.Frame Frame16 
      Caption         =   "ÉuÉâÉbÉNÇÃì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   6300
      TabIndex        =   38
      Top             =   4306
      Width           =   5895
      Begin VB.CheckBox chkblack 
         Caption         =   "ÉuÉâÉbÉNÉäÉXÉgÇ…ìoò^Ç≥ÇÍÇƒÇ¢ÇÈå⁄ãq"
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
         Height          =   345
         Left            =   420
         TabIndex        =   17
         Top             =   420
         Width           =   4455
      End
   End
   Begin VB.Frame Frame15 
      Caption         =   "VoIPÇÃì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   6300
      TabIndex        =   37
      Top             =   3282
      Width           =   5895
      Begin VB.CheckBox chkvoip 
         Caption         =   "VoIPÇÃî‘çÜÇ™ìoò^Ç≥ÇÍÇƒÇ¢ÇÈå⁄ãq"
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
         Height          =   345
         Left            =   390
         TabIndex        =   16
         Top             =   420
         Width           =   4215
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "ì¸ã‡ä÷òAÇÃì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   6300
      TabIndex        =   36
      Top             =   1234
      Width           =   5895
      Begin VB.CheckBox chkzankin 
         Caption         =   "écäzÇ™Ç†ÇÈå⁄ãq"
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
         Height          =   345
         Left            =   360
         TabIndex        =   10
         Top             =   390
         Width           =   2355
      End
      Begin VB.TextBox txtnyuukintsuki 
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
         Left            =   4980
         MaxLength       =   2
         TabIndex        =   12
         Top             =   300
         Width           =   345
      End
      Begin VB.TextBox txtnyuukinnen 
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
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   11
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label44 
         Caption         =   "åé"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5370
         TabIndex        =   41
         Top             =   480
         Width           =   195
      End
      Begin VB.Label Label43 
         Caption         =   "îN"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4740
         TabIndex        =   40
         Top             =   480
         Width           =   285
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "âºèZèäÇÃì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   6300
      TabIndex        =   35
      Top             =   210
      Width           =   5895
      Begin VB.TextBox txtkarijuusho 
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
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   9
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "ÅuQÅvÇÃì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   6300
      TabIndex        =   34
      Top             =   2258
      Width           =   5895
      Begin VB.CheckBox chkq 
         Caption         =   "ÅuQÅvÇégópÇµÇƒÇ¢ÇÈå⁄ãq"
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
         Height          =   345
         Left            =   390
         TabIndex        =   13
         Top             =   420
         Width           =   2985
      End
      Begin VB.TextBox txtqnen 
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
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   14
         Top             =   330
         Width           =   615
      End
      Begin VB.TextBox txtqtsuki 
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
         Left            =   4980
         MaxLength       =   2
         TabIndex        =   15
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label48 
         Caption         =   "îN"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4740
         TabIndex        =   43
         Top             =   510
         Width           =   285
      End
      Begin VB.Label Label47 
         Caption         =   "åé"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5370
         TabIndex        =   42
         Top             =   510
         Width           =   195
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "ê›íuêÊÉÅÉÇÇÃì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   210
      TabIndex        =   33
      Top             =   8400
      Width           =   5895
      Begin VB.TextBox txtsecchimemo 
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
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   8
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "ê›íuêÊèZèäÇÃì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   210
      TabIndex        =   32
      Top             =   7365
      Width           =   5895
      Begin VB.TextBox txtsecchijuusho 
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
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   7
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "ê›íuêÊÃÿ∂ﬁ≈ÇÃì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   210
      TabIndex        =   31
      Top             =   6345
      Width           =   5895
      Begin VB.TextBox txtsecchifuri 
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
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   6
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "å_ñÒé“î‘ínÇÃì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   210
      TabIndex        =   30
      Top             =   4300
      Width           =   5895
      Begin VB.TextBox txtkenjuu2 
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
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   4
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ï¿Ç◊èáëIë"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   6390
      TabIndex        =   28
      Top             =   9450
      Width           =   3075
      Begin VB.OptionButton optnarabe 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ãÿ∂ﬁ≈èá"
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
         Index           =   1
         Left            =   210
         TabIndex        =   29
         Top             =   270
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton optnarabe 
         BackColor       =   &H00FFFFFF&
         Caption         =   "å⁄ãqIDèá"
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
         Index           =   0
         Left            =   1590
         TabIndex        =   21
         Top             =   270
         Width           =   1365
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "îıçlÇÃì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   210
      TabIndex        =   27
      Top             =   5330
      Width           =   5895
      Begin VB.TextBox txtkenbi 
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
         IMEMode         =   4  'ëSäpÇ–ÇÁÇ™Ç»
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   5
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "å¬êlIDÇÃì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   210
      TabIndex        =   26
      Top             =   180
      Width           =   5895
      Begin VB.TextBox txtkojinid 
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
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   0
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "å_ñÒé“èZèäÇQÇÃì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   210
      TabIndex        =   25
      Top             =   3270
      Width           =   5895
      Begin VB.TextBox txtkenjuu 
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
         IMEMode         =   4  'ëSäpÇ–ÇÁÇ™Ç»
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   3
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "èZëÓãÊï™ÇÃëIë"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   6300
      TabIndex        =   24
      Top             =   6354
      Width           =   5895
      Begin VB.ComboBox comjuutaku 
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
         Left            =   1920
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   20
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ìdòbî‘çÜÇÃì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   210
      TabIndex        =   23
      Top             =   2240
      Width           =   5895
      Begin VB.TextBox txtkentel 
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
         Left            =   1920
         MaxLength       =   12
         TabIndex        =   2
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "å_ñÒé“Ç”ÇËÇ™Ç»ì¸óÕ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   915
      Left            =   210
      TabIndex        =   22
      Top             =   1210
      Width           =   5895
      Begin VB.TextBox txtkenfuri 
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
         IMEMode         =   6  'îºäp∂¿∂≈
         Left            =   1920
         MaxLength       =   45
         TabIndex        =   1
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "∑¨›æŸ"
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
      Left            =   11130
      TabIndex        =   47
      Top             =   10170
      Width           =   750
   End
   Begin VB.Label SSCJIKKOU 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "åüçı"
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
      Left            =   10020
      TabIndex        =   46
      Top             =   10170
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   1065
      Left            =   6300
      Top             =   9420
      Width           =   5880
   End
End
Attribute VB_Name = "frmkensaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdend_Click()

End Sub

Private Sub cmdkensaku_Click()
End Sub

Private Sub comjuutaku_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub

Private Sub Form_Load()
With frmkensaku
    
    .comjuutaku.AddItem "àÍî â∆íÎ"
    .comjuutaku.AddItem "àÍî éññ±èä"
    .comjuutaku.AddItem "åˆã§"
    .comjuutaku.AddItem "åˆã§ÇP"
    .comjuutaku.AddItem "èWçáèZëÓ"
    .comjuutaku.AddItem "î_ÅEéññ±èä"
    .comjuutaku.AddItem "óVóóâÄ"
    .comjuutaku.AddItem "ÇªÇÃëº"
    
    .comjuutaku.ListIndex = -1

    .cmbshiharai.AddItem "êøãÅèëêøãÅ"
    .cmbshiharai.AddItem "å˚ç¿êøãÅ"


End With
End Sub

Private Sub Picture13_Click()
Unload frmkensaku
Set frmkensaku = Nothing

End Sub

Private Sub Picture7_Click()
    
Dim k_furi As String, k_tel As String, wherestr As String, kenrs As Recordset, k_juu As String
Dim k_id As String, k_bi As String, k_nara As String, k_juu2 As String, k_secchi_furi As String
Dim k_secchi_juusho As String, k_secchi_memo As String, k_kariju As String, k_nyu_hi As String
Dim k_nhkno As String, k_kouza As String, k_nana As String, k_q_hi As String

If koushinchk = False Then
    Exit Sub
End If


Screen.MousePointer = 11

kidou_main


With frmkensaku
    k_id = Trim(.txtkojinid.Text)
    k_furi = Trim(.txtkenfuri.Text)
    k_tel = Trim(.txtkentel.Text)
    k_juu = Trim(.txtkenjuu.Text)
    k_juu2 = Trim(.txtkenjuu2.Text)
    k_bi = Trim(.txtkenbi.Text)
    k_secchi_furi = Trim(.txtsecchifuri.Text)
    k_secchi_juusho = Trim(.txtsecchijuusho.Text)
    k_secchi_memo = Trim(.txtsecchimemo.Text)
    k_kariju = Trim(.txtkarijuusho.Text)
    k_nyu_hi = Trim(.txtnyuukinnen.Text) & Trim(.txtnyuukintsuki.Text)
    k_q_hi = Trim(.txtqnen.Text) & Trim(.txtqtsuki.Text)
    k_nhkno = Trim(.txtnhkno.Text)
    k_kouza = Trim(txtkouzaname.Text)
    k_nana = Trim(txtname.Text)
  '  If chknhk.Value = 0 And k_nhkno = "" And chkblack.Value = 0 And chkvoip.Value = 0 And chkzankin.Value = 0 And k_kariju = "" And k_id = "" And k_furi = "" And k_tel = "" And k_juu = "" And k_bi = "" And k_juu2 = "" And k_secchi_furi = "" And k_secchi_juusho = "" And k_secchi_memo = "" Then
  '      MsgBox "Ç«ÇÍÇ©ÇÕÅAç≈í·1Ç¬ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
  '      .txtkojinid.SetFocus
  '      Screen.MousePointer = 0
  '      Exit Sub
  '  End If
    sql = "select*from kojin where "
    wherestr = ""
    If k_id <> "" Then
        wherestr = "kojinid like '*" & k_id & "*'"
    End If
    If k_furi <> "" Then
        If wherestr = "" Then
            wherestr = "furigana like '*" & k_furi & "*'"
        Else
            wherestr = wherestr & "and furigana like '*" & k_furi & "*'"
        End If
    End If
    If k_tel <> "" Then
        If wherestr = "" Then
            wherestr = "tel1 like '*" & k_tel & "*'"
        Else
            wherestr = wherestr & "and tel1 like '*" & k_tel & "*'"
        End If
    End If
    If k_juu <> "" Then
        If wherestr = "" Then
            If k_juu2 <> "" Then
                wherestr = "juusho2 like '*" & k_juu & "*' and juusho2 like '*" & k_juu2 & "*'"
            Else
                wherestr = "juusho2 like '*" & k_juu & "*'"
            End If
        Else
            If k_juu2 <> "" Then
                wherestr = wherestr & "and juusho2 like '*" & k_juu & "*' and juusho2 like '*" & k_juu2 & "*'"
            Else
                wherestr = wherestr & "and juusho2 like '*" & k_juu & "*'"
            End If
        End If
    Else
        If k_juu2 <> "" Then
            If wherestr = "" Then
                wherestr = "juusho2 like '*" & k_juu2 & "*'"
            Else
                wherestr = wherestr & "and juusho2 like '*" & k_juu2 & "*'"
            End If
        End If
    End If
    If k_bi <> "" Then
        If wherestr = "" Then
            wherestr = "bikou like '*" & k_bi & "*'"
        Else
            wherestr = wherestr & "and bikou like '*" & k_bi & "*'"
        End If
    End If
    If k_secchi_furi <> "" Then
        If wherestr = "" Then
            wherestr = "k_furigana like '*" & k_secchi_furi & "*'"
        Else
            wherestr = wherestr & "and k_furigana like '*" & k_secchi_furi & "*'"
        End If
    End If
    If k_secchi_juusho <> "" Then
        If wherestr = "" Then
            wherestr = "k_juusho like '*" & k_secchi_juusho & "*'"
        Else
            wherestr = wherestr & "and k_juusho like '*" & k_secchi_juusho & "*'"
        End If
    End If
    If k_secchi_memo <> "" Then
        If wherestr = "" Then
            wherestr = "k_memo like '*" & k_secchi_memo & "*'"
        Else
            wherestr = wherestr & "and k_memo like '*" & k_secchi_memo & "*'"
        End If
    End If

    If k_kariju <> "" Then
        If wherestr = "" Then
            wherestr = "karijusho like '*" & k_kariju & "*'"
        Else
            wherestr = wherestr & "and karijusho like '*" & k_kariju & "*'"
        End If
    End If
    
    If k_kouza <> "" Then
        If wherestr = "" Then
            wherestr = "meigi like '*" & k_kouza & "*'"
        Else
            wherestr = wherestr & "and meigi like '*" & k_kouza & "*'"
        End If
    End If
    
    If k_nana <> "" Then
        If wherestr = "" Then
            wherestr = "name like '*" & k_nana & "*'"
        Else
            wherestr = wherestr & "and name like '*" & k_nana & "*'"
        End If
    End If
    If chkq.Value = 1 Then
        If Len(k_q_hi) = 6 Then
            If wherestr = "" Then
                wherestr = "qshiyou <> space(1) and qhiduke ='" & k_q_hi & "'"
            Else
                wherestr = wherestr & "and qshiyou <> space(1) and qhiduke ='" & k_q_hi & "'"
            End If
        Else
            MsgBox "ÇpÇÃåéì˙ÇÕó·)2019/01ÇÃÇÊÇ§Ç…ëSÇUåÖÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
    End If
    
    
    If chkzankin.Value = 1 Then
        If Len(k_nyu_hi) = 6 Then
            If wherestr = "" Then
                wherestr = "zankin <> space(1) and nyuukinbi ='" & k_nyu_hi & "'"
            Else
                wherestr = wherestr & "and zankin <> space(1) and nyuukinbi ='" & k_nyu_hi & "'"
            End If
        Else
            If wherestr = "" Then
                wherestr = "zankin <> space(1)"
            Else
                wherestr = wherestr & "and zankin <> space(1)"
            End If
        End If
    End If
    If chkvoip.Value = 1 Then
        If wherestr = "" Then
            wherestr = "v1 <> space(1)"
        Else
            wherestr = wherestr & "and v1 <> space(1)"
        End If
    End If
    If chkblack.Value = 1 Then
        If wherestr = "" Then
            wherestr = "f2 ='1'"
        Else
            wherestr = wherestr & "and f2 ='1'"
        End If
    End If
    If chknhk.Value = 1 Then
        If wherestr = "" Then
            wherestr = "nhkno <> space(1)"
        Else
            wherestr = wherestr & "and nhkno <> space(1)"
        End If
    End If
    If k_nhkno <> "" Then
        If wherestr = "" Then
            wherestr = "nhkno like '*" & k_nhkno & "*'"
        Else
            wherestr = wherestr & "and nhkno like '*" & k_nhkno & "*'"
        End If
    End If
    If .comjuutaku.ListIndex <> -1 Then
        If wherestr = "" Then
            wherestr = "juutaku = '" & CStr(.comjuutaku.ListIndex) & "'"
        Else
            wherestr = wherestr & "and juutaku = '" & CStr(.comjuutaku.ListIndex) & "'"
        End If
    End If
    If optnarabe(0).Value = True Then
        k_nara = ""
    Else
        k_nara = " order by furigana"
    End If
    If cmbshiharai.ListIndex = 0 Then
        If wherestr = "" Then
            wherestr = "seikyuu = '1'"
        Else
            wherestr = wherestr & "and seikyuu = '1'"
        End If
    ElseIf cmbshiharai.ListIndex = 1 Then
        If wherestr = "" Then
            wherestr = "seikyuu = '0'"
        Else
            wherestr = wherestr & "and seikyuu = '0'"
        End If
    End If
    
    If chkend.Value = 1 Then
        If wherestr = "" Then
            wherestr = "eend = '0'"
        Else
            wherestr = wherestr & "and eend = '0'"
        End If
    End If
    
    sql = sql & wherestr & k_nara
    If FcSQlGet(kenrs, sql, PrMsg) = False Then
        Screen.MousePointer = 0
        MsgBox "äYìñÇ∑ÇÈÉfÅ[É^ÇÕìoò^Ç≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB"
        Exit Sub
    Else
        kenrs.MoveLast
        'If kenrs.RecordCount = 1 Then
            'If hyouji_touroku(kenrs!kojinid) = False Then
                'Screen.MousePointer = 0
                'Exit Sub
            'End If
            'If nhk_set(kenrs!kojinid) = -1 Then
                'Screen.MousePointer = 0
                'Exit Sub
            'End If
            'keiyakuset (kenrs!kojinid)
            'keiyakuset2 (kenrs!kojinid)
            'Unload frmkensaku
            'Set frmkensaku = Nothing
            'frmtouroku2.Show 1
        'Else
            kenrs.MoveFirst
            i = 1
            frmkekkaichi.lblkensuu.Caption = kenrs.RecordCount
            With frmkekkaichi.kekkagrd
            .Rows = kenrs.RecordCount + 1
            .ColWidth(0) = 1000
            .ColWidth(1) = 2300
            .ColWidth(2) = 2400
            .ColWidth(3) = 1700
            .ColWidth(4) = 2400
            .ColWidth(5) = 1000
            .ColWidth(6) = 1000
            .Row = 0
            .Col = 0
            .Text = "ÇhÇc"
            .Col = 1
            .Text = "ÉtÉäÉKÉi"
            .Col = 2
            .Text = "å_ñÒé“ñº"
            .Col = 3
            .Text = "ìdòbî‘çÜ"
            .Col = 4
            .Text = "èZèäÇQ"
            .Col = 5
            .Text = "ê¢ë—êî"
            .Col = 6
            .Text = "ÇsÇuêî"
           Do Until kenrs.EOF
                .Row = i
                .Col = 0
                .Text = kenrs!kojinid
                .Col = 1
                .CellAlignment = 1
                .Text = kenrs!Furigana
                .Col = 2
                .CellAlignment = 1
                .Text = kenrs!Name
                .Col = 3
                .CellAlignment = 1
                .Text = kenrs!tel1
                .Col = 4
                .Text = kenrs!juusho2
                .Col = 5
                .Text = kenrs!setaisuu
                .Col = 6
                .Text = kenrs!tvsuu
               i = i + 1
                kenrs.MoveNext
            Loop
            .Col = 0
            .Row = 1
            End With
            kenrs.Close
            Screen.MousePointer = 0
            frmkensaku.Hide
            frmkekkaichi.Show 1
        'End If
    End If
End With

End Sub

Private Sub txtkarijuusho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub

Private Sub txtkenbi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub

Private Sub txtkenfuri_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If
End Sub

Private Sub txtkenjuu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub

Private Sub txtkenjuu2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub

Private Sub txtkentel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub

Private Sub txtkojinid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub

Private Sub txtkouzaname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub

Private Sub txtsecchifuri_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub

Private Sub txtsecchijuusho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub

Private Sub txtsecchimemo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub
