VERSION 5.00
Begin VB.Form frmkeiyaku2 
   BorderStyle     =   3  'ŒÅ’èÀŞ²±Û¸Ş
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11835
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '‰æ–Ê‚Ì’†‰›
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Ì×¯Ä
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '‚È‚µ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   9210
      MouseIcon       =   "frmkeiyaku2.frx":0000
      MousePointer    =   99  'Õ°»Ş°’è‹`
      Picture         =   "frmkeiyaku2.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   30
      Top             =   6510
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ì×¯Ä
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '‚È‚µ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   10740
      MouseIcon       =   "frmkeiyaku2.frx":11C9
      MousePointer    =   99  'Õ°»Ş°’è‹`
      Picture         =   "frmkeiyaku2.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   29
      Top             =   6510
      Width           =   480
   End
   Begin VB.Frame fh 
      Caption         =   "•ÏX•û–@"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6105
      Left            =   8670
      TabIndex        =   26
      Top             =   90
      Width           =   3015
      Begin VB.TextBox txtnoldtanka 
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   2  'µÌ
         Left            =   1290
         MaxLength       =   6
         TabIndex        =   34
         ToolTipText     =   "VŒ_–ñ’P‰¿‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B"
         Top             =   5100
         Width           =   1455
      End
      Begin VB.ComboBox cmbshurui 
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   570
         Style           =   2  'ÄŞÛ¯ÌßÀŞ³İ Ø½Ä
         TabIndex        =   33
         Top             =   3990
         Width           =   2175
      End
      Begin VB.OptionButton opthouhou 
         Caption         =   "¡‚Ü‚Å‚ÌŒ_–ñ‚Ì‹àŠz‚ğA•ÏXŒã‚ÌŒ_–ñ‹àŠz‚É•ÏX‚·‚éiŒ_–ñ‚ª‘¶İ‚ ‚éê‡j"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   1
         Left            =   330
         TabIndex        =   28
         Top             =   2430
         Width           =   2415
      End
      Begin VB.OptionButton opthouhou 
         Caption         =   "¡‚Ü‚Å‚ÌŒ_–ñ‚Ì‹àŠz‚ğ•ÏX‚¹‚¸AV‚µ‚¢Œ_–ñ‚©‚ç•ÏXŒã‚Ì‹àŠz‚ğ“K—p‚·‚é"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   0
         Left            =   330
         TabIndex        =   27
         Top             =   690
         Value           =   -1  'True
         Width           =   2445
      End
      Begin VB.Label Label10 
         Caption         =   "‹Œ’P‰¿"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   300
         TabIndex        =   35
         Top             =   5160
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "‹ŒŒ_–ñÚ×"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3510
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   8355
      Begin VB.Frame Frame3 
         Height          =   555
         Left            =   1800
         TabIndex        =   20
         Top             =   1740
         Width           =   6315
         Begin VB.OptionButton optoldbunrui 
            Caption         =   "ƒCƒ“ƒ^[ƒlƒbƒgŠÖ˜A—¿‹à•ª—Ş"
            BeginProperty Font 
               Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3180
            TabIndex        =   22
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton optoldbunrui 
            Caption         =   "CATVŠÖ˜A—¿‹à•ª—Ş"
            BeginProperty Font 
               Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.Label lblchousei 
         Alignment       =   2  '’†‰›‘µ‚¦
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7110
         TabIndex        =   54
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "’²®"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   6480
         TabIndex        =   55
         Top             =   3090
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "•\‹L(•¶Í)"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   180
         TabIndex        =   51
         Top             =   3060
         Width           =   1185
      End
      Begin VB.Label lblhyouki 
         Alignment       =   2  '’†‰›‘µ‚¦
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   50
         Top             =   3000
         Width           =   4785
      End
      Begin VB.Label lblnarabe 
         Alignment       =   2  '’†‰›‘µ‚¦
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   41
         Top             =   2430
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "•À‚Ñ‡"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   4890
         TabIndex        =   40
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblgroup 
         Alignment       =   2  '’†‰›‘µ‚¦
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3510
         TabIndex        =   39
         Top             =   2430
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ƒOƒ‹[ƒv"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   2595
         TabIndex        =   38
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblshuukei 
         Alignment       =   2  '’†‰›‘µ‚¦
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   37
         Top             =   2430
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "WŒv—L–³"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   36
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "—¿‹à•ª—Ş"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   180
         TabIndex        =   18
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label lbloldsetsumei 
         Alignment       =   2  '’†‰›‘µ‚¦
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   17
         Top             =   1320
         Width           =   7935
      End
      Begin VB.Label Label6 
         Caption         =   "à–¾"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "í—Ş”Ô†"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lbloldshuruiid 
         Alignment       =   2  '’†‰›‘µ‚¦
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "í—Ş–¼"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1620
         TabIndex        =   9
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lbloldshuruimei 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1380
         TabIndex        =   8
         Top             =   540
         Width           =   4875
      End
      Begin VB.Label Label2 
         Caption         =   "’P‰¿"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6300
         TabIndex        =   7
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lbloldtanka 
         Alignment       =   2  '’†‰›‘µ‚¦
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   6
         Top             =   540
         Width           =   1755
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "VŒ_–ñÚ×"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3540
      Left            =   90
      TabIndex        =   3
      Top             =   3840
      Width           =   8355
      Begin VB.TextBox txtchousei 
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   2  'µÌ
         Left            =   7290
         MaxLength       =   1
         TabIndex        =   56
         ToolTipText     =   "VŒ_–ñ’P‰¿‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B"
         Top             =   3060
         Width           =   855
      End
      Begin VB.TextBox txthyouki 
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   1  'µİ
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   53
         ToolTipText     =   "•\‹L–¼‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B"
         Top             =   3060
         Width           =   4935
      End
      Begin VB.Frame Frame5 
         Height          =   585
         Left            =   1440
         TabIndex        =   47
         Top             =   2370
         Width           =   2445
         Begin VB.OptionButton optshuukei 
            Caption         =   "‚ ‚è"
            Height          =   285
            Index           =   1
            Left            =   1260
            TabIndex        =   49
            Top             =   210
            Width           =   795
         End
         Begin VB.OptionButton optshuukei 
            Caption         =   "‚È‚µ"
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   48
            Top             =   210
            Width           =   795
         End
      End
      Begin VB.TextBox txtnarabe 
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   2  'µÌ
         Left            =   7290
         MaxLength       =   2
         TabIndex        =   46
         ToolTipText     =   "VŒ_–ñ’P‰¿‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B"
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtgroup 
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   2  'µÌ
         Left            =   5310
         MaxLength       =   1
         TabIndex        =   45
         ToolTipText     =   "VŒ_–ñ’P‰¿‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B"
         Top             =   2520
         Width           =   855
      End
      Begin VB.Frame Frame4 
         Height          =   555
         Left            =   1800
         TabIndex        =   23
         Top             =   1740
         Width           =   6315
         Begin VB.OptionButton optnewbunrui 
            Caption         =   "CATVŠÖ˜A—¿‹à•ª—Ş"
            BeginProperty Font 
               Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   120
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton optnewbunrui 
            Caption         =   "ƒCƒ“ƒ^[ƒlƒbƒgŠÖ˜A—¿‹à•ª—Ş"
            BeginProperty Font 
               Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3180
            TabIndex        =   24
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.TextBox txtnewsetsumei 
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   1  'µİ
         Left            =   180
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "VŒ_–ñ‚Ìà–¾‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B"
         Top             =   1320
         Width           =   7875
      End
      Begin VB.TextBox txtnewshuruimei 
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   1  'µİ
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   0
         ToolTipText     =   "VŒ_–ñí—Ş–¼‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B"
         Top             =   540
         Width           =   4935
      End
      Begin VB.TextBox txtnewtanka 
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   2  'µÌ
         Left            =   6420
         MaxLength       =   6
         TabIndex        =   1
         ToolTipText     =   "VŒ_–ñ’P‰¿‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B"
         Top             =   540
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "’²®"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   6600
         TabIndex        =   57
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "•\‹L(•¶Í)"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   150
         TabIndex        =   52
         Top             =   3120
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "•À‚Ñ‡"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   6480
         TabIndex        =   44
         Top             =   2580
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ƒOƒ‹[ƒv"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   4275
         TabIndex        =   43
         Top             =   2550
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "WŒv—L–³"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   150
         TabIndex        =   42
         Top             =   2580
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "—¿‹à•ª—Ş"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   180
         TabIndex        =   19
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "à–¾"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   240
         TabIndex        =   16
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label7 
         Caption         =   "í—Ş”Ô†"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   180
         TabIndex        =   15
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "í—Ş–¼"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1620
         TabIndex        =   14
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "’P‰¿"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6390
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblnewshuruiid 
         Alignment       =   2  '’†‰›‘µ‚¦
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   540
         Width           =   1035
      End
   End
   Begin VB.Label SSCJIKKOU 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "“o˜^"
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   9240
      TabIndex        =   32
      Top             =   7020
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "I—¹"
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
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
      Left            =   10770
      TabIndex        =   31
      Top             =   7020
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '•s“§–¾
      BorderStyle     =   0  '“§–¾
      Height          =   1035
      Left            =   8640
      Top             =   6360
      Width           =   3030
   End
End
Attribute VB_Name = "frmkeiyaku2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function henkouhannei(hshuruiid As String, nkingaku As String, old_tanka As String) As Boolean

Dim sql_han As String, rs_han As Recordset, hancount As Long, hankuchi As Integer, hankadou As String
Dim henkokyaid As String, sql_shi As String, rs_shi As Recordset, oldgou As Long
Dim saisuru As Integer, Newgou As Long, sql_n1 As String, tkaishi As String
Dim sql_j1 As String, gen_shukeiyakushoukei As Long, gen_shukeiyakuseikyuu As Long, gen_soukei As Long
Dim saki_shukeiyakushoukei As Long, saki_shukeiyakuseikyuu As Long, saki_soukei As Long

'Œ»İ—LŒø‚ÈŒ_–ñ‚ğŠm”F
Dim sentakushurui As Integer

If Me.cmbshurui.ListIndex = 0 Then
    sentakushurui = 0
ElseIf Me.cmbshurui.ListIndex = 1 Then
'W‡Z‘îˆÈŠO
    sentakushurui = 1
Else
'W‡Z‘î
    sentakushurui = 2
End If

sql_han = "select * from keiyaku where shurui = '" & hshuruiid & "'"


If FcSQlGet(rs_han, sql_han, PrMsg) = False Then
    Screen.MousePointer = 0
    MsgBox "•ÏX‚µ‚½Œ_–ñ‚ÍAŒ»İg—p‚³‚ê‚Ä‚¢‚Ü‚¹‚ñB‹àŠz‚ğ”½‰f‚·‚é•K—v‚ª‚ ‚è‚Ü‚¹‚ñ‚Å‚µ‚½B"
    henkouhannei = False
    Exit Function
Else
    rs_han.MoveLast
    hancount = rs_han.RecordCount
    MsgBox "w" & hancount & "xŒ‚Ìƒf[ƒ^‚ğXV‚µ‚Ü‚·B"
End If
Screen.MousePointer = 11
 tkaishi = Format(Date, "yyyymmdd")
rs_han.MoveFirst
'*********ƒgƒ‰ƒ“ƒUƒNƒVƒ‡ƒ“ŠJn**************
BeginTrans
    Do Until rs_han.EOF
        saisuru = 0 'ÄŒvZ‚·‚é
        hankuchi = CInt(rs_han!kuchisuu)
        hankadou = rs_han!kadou
        henkokyaid = rs_han!kojinid
        'ŒûÀ‚Ìó‘Ô‚ªY‚©‚»‚¤‚Å‚È‚¢‚©‚ğŠm”F‚ÆƒƒbƒZ[ƒW‚†‚P‚O
        sql_shi = "select * from kojin where kojinid = '" & henkokyaid & "'"
        If FcSQlGet(rs_shi, sql_shi, PrMsg) = False Then
            Screen.MousePointer = 0
            MsgBox "•ÏX‚µ‚æ‚¤‚Æ‚µ‚½Œ_–ñ‚ğ’÷Œ‹‚µ‚½Œ_–ñÒ‚ÍAŒ»İ‘¶İ‚µ‚Ü‚¹‚ñBŒÚ‹qNO@" & henkokyaid
            saisuru = 1
        Else
           
            If rs_shi!f1 = "1" Then
                Screen.MousePointer = 0
                MsgBox "‹àŠz‚ğ•ÏX‚µ‚æ‚¤‚Æ‚µ‚½Œ_–ñÒu" & henkokyaid & "/" & rs_shi!Name & "v‚ÌŠÇ—ó‘Ô‚ÍAY‚Å‚·Bè“®‚Å•ÏX‚µ‚Ä‚­‚¾‚³‚¢B"
                saisuru = 1
            End If
            If rs_han!kadou <> "0" Then
                saisuru = 2
            End If
             'Œ³‚Ìƒf[ƒ^
             gen_shukeiyakushoukei = rs_shi!jikaishu 'åŒ_–ñ¬Œv‹àŠz
             If Trim(gen_shukeiyakushoukei) = "" Then
                 gen_shukeiyakushoukei = "0"
             End If
             gen_shukeiyakuseikyuu = rs_shi!jikaikin 'åŒ_–ñ¿‹‹àŠz
             If Trim(gen_shukeiyakuseikyuu) = "" Then
                 gen_shukeiyakuseikyuu = "0"
             End If
             gen_soukei = rs_shi!soukei
             
            
                
            If sentakushurui = 1 Then
                If rs_shi!juutaku = "4" Then
                    saisuru = 1
                End If
            ElseIf sentakushurui = 2 Then
                If rs_shi!juutaku <> "4" Then
                    saisuru = 1
                End If
            End If
        
            rs_shi.Close
            
        End If
        
        If saisuru = 0 Then
            'Œ_–ñ‚ÌŒû”‚É‰‚¶‚ÄŒ_–ñ‚ğ‘‚«Š·‚¦‚é
            oldgou = hankuchi * CLng(old_tanka)
            Newgou = hankuchi * CLng(nkingaku)
            sql_n1 = "update keiyaku set kingaku = '" & Newgou & "'" & _
                    " where keiyakuid = '" & rs_han!keiyakuid & "'"
             On Error GoTo errhan
                  db.Execute sql_n1, dbSQLPassThrough
             On Error GoTo 0
           'Œ_–ñ‚ğ‘‚«Š·‚¦‚È‚ª‚çAÄŒvZ‚ğÀs‚·‚é
            saki_shukeiyakushoukei = CStr(CLng(gen_shukeiyakushoukei) - CLng(oldgou) + CLng(Newgou))
            saki_shukeiyakuseikyuu = CStr(CLng(gen_shukeiyakuseikyuu) - CLng(oldgou) + CLng(Newgou))
            saki_soukei = CStr(CLng(gen_soukei) - CLng(oldgou) + CLng(Newgou))
             
             
             sql_j1 = "UPDATE kojin SET jikaikin = '" & CStr(saki_shukeiyakuseikyuu) & "'," & _
                 "soukei = '" & CStr(saki_soukei) & "',jikaishu = '" & CStr(saki_shukeiyakushoukei) & "'" & _
                 " WHERE kojinid = '" & henkokyaid & "'"
             On Error GoTo errhan
               db.Execute sql_j1, dbSQLPassThrough
              On Error GoTo 0
        ElseIf saisuru = 2 Then
            'Œ_–ñ‚ÌŒû”‚É‰‚¶‚ÄŒ_–ñ‚ğ‘‚«Š·‚¦‚é
            oldgou = hankuchi * CLng(old_tanka)
            Newgou = hankuchi * CLng(nkingaku)
            sql_n1 = "update keiyaku set kingaku = '" & Newgou & "'" & _
                    " where keiyakuid = '" & rs_han!keiyakuid & "'"
             On Error GoTo errhan
                  db.Execute sql_n1, dbSQLPassThrough
             On Error GoTo 0
           'Œ_–ñ‚ğ‘‚«Š·‚¦‚È‚ª‚çAÄŒvZ‚ğÀs‚·‚é
            saki_shukeiyakushoukei = CStr(CLng(gen_shukeiyakushoukei) - CLng(oldgou) + CLng(Newgou))
             
             
             sql_j1 = "UPDATE kojin SET jikaishu = '" & CStr(saki_shukeiyakushoukei) & "'" & _
                 " WHERE kojinid = '" & henkokyaid & "'"
             On Error GoTo errhan
               db.Execute sql_j1, dbSQLPassThrough
              On Error GoTo 0
        End If
        rs_han.MoveNext
    Loop
'*************ƒgƒ‰ƒ“ƒUƒNƒVƒ‡ƒ“I—¹********
CommitTrans
rs_han.Close
Screen.MousePointer = 0
MsgBox "³í‚É•ÏX‚ªI—¹‚µ‚Ü‚µ‚½B"
henkouhannei = True
Exit Function

errhan:
    MsgBox "Œ_–ñ“à—e‚Ì•ÏX‚É¸”s‚µ‚Ü‚µ‚½BŒ³‚Ìó‘Ô‚É–ß‚µ‚Ü‚·B"
    Rollback
    henkouhannei = False
    Screen.MousePointer = 0
    Exit Function


End Function



Private Sub Form_Load()


    Left = (Screen.Width - Width) / 2   ' ƒtƒH[ƒ€‚ğ‰æ–Ê‚Ì…•½•ûŒü‚ÉƒZƒ“ƒ^ƒŠƒ“ƒO‚µ‚Ü‚·B
    Top = (Screen.Height - Height) / 2  ' ƒtƒH[ƒ€‚ğ‰æ–Ê‚Ìc•ûŒü‚ÉƒZƒ“ƒ^ƒŠƒ“ƒO‚µ‚Ü‚·B
   
   cmbshurui.Clear
   cmbshurui.AddItem ("‚·‚×‚Ä")
    cmbshurui.AddItem ("W‡Z‘îˆÈŠO")
     cmbshurui.AddItem ("W‡Z‘î")
     cmbshurui.ListIndex = 0
End Sub


Private Sub Picture13_Click()
    Unload frmkeiyaku2
    Set frmkeiyaku2 = Nothing
    frmkeiyaku.Show 1

End Sub

Private Sub Picture7_Click()
Dim nid As String, nmei As String, ntan As String, nsetsu As String, nnid As String
Dim nkan As String

Dim s_shuukei As String, s_group As String, s_narabe As String

Dim s_hyoukimei As String, s_kasan As String

With frmkeiyaku2

nid = Trim(.lblnewshuruiid.Caption)
nmei = Trim(.txtnewshuruimei.Text)
ntan = Trim(.txtnewtanka.Text)
nsetsu = Trim(.txtnewsetsumei.Text)

s_group = Trim(.txtgroup.Text)
s_narabe = Trim(.txtnarabe.Text)

s_hyoukimei = Trim(.txthyouki.Text)

s_kasan = Trim(txtchousei.Text)

s_shuukei = ""
If .optshuukei(0).Value = True Then
s_shuukei = "0"
ElseIf .optshuukei(1).Value = True Then
s_shuukei = "1"
Else
    MsgBox "WŒv‚ğ•K‚¸‘I‘ğ‚µ‚Ä‚­‚¾‚³‚¢B"
   Exit Sub
End If





If optnewbunrui(0).Value = True Then
    nkan = "0"
Else
    nkan = "1"
End If
If nsetsu = "" Or ntan = "" Or nmei = "" Then
   MsgBox "Še–€‚ğ•K‚¸“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B"
   Exit Sub
End If
 If opthouhou(0).Value = False Then
    If cmbshurui.ListIndex <> 0 Then
        If Trim(txtnoldtanka.Text) = "" Then
            MsgBox "‹Œ‹àŠz‚ğ•K‚¸“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B"
            Exit Sub
        End If
    Else
        If ntan = lbloldtanka.Caption Then
            MsgBox "V‹Œ‹àŠz‚ª“¯‚¶‚È‚Ì‚Åì‹Æ‚ğ’†~‚µ‚Ü‚µ‚½B"
            Exit Sub
        End If
    End If
 End If

 On Error GoTo JIKKOUERROR

 If .SSCJIKKOU.Caption = "“o˜^" Then
     
     sql = "INSERT INTO shurui VALUES('" & nid & "'," & _
                    "'" & nmei & "','" & ntan & "','" & nsetsu & "','" & nkan & "','" & s_shuukei & "','" & s_group & "','" & s_narabe & "','" & s_hyoukimei & "','" & s_kasan & "')"
     db.Execute sql, dbSQLPassThrough
     nnid = CStr(CLng(nid) + 1)
        If soukatsu(1, "renban", 2, nnid, 1, 1) = False Then
                MsgBox "İ’è‚Ì‘‚«‚İ‚É¸”s‚µ‚Ü‚µ‚½B"
                Exit Sub
        End If
     MsgBox "“o˜^‚µ‚Ü‚µ‚½B"
 Else
         sql = "UPDATE shurui SET shuruimei = '" & nmei & "'," & _
                      "tanka = '" & ntan & "',setsumei = '" & nsetsu & "',kankei = '" & nkan & "',iru = '" & s_shuukei & "',gp = '" & s_group & "',narabe = '" & s_narabe & "',hyoukimei = '" & s_hyoukimei & "',kasan = '" & s_kasan & "'" & _
                      " WHERE shuruiid = '" & nid & "'"
    If opthouhou(0).Value = True Then
          db.Execute sql, dbSQLPassThrough
       MsgBox "•ÏX‚µ‚Ü‚µ‚½B"
    Else
        db.Execute sql, dbSQLPassThrough
        
        If cmbshurui.ListIndex = 0 Then
            If henkouhannei(nid, ntan, lbloldtanka.Caption) = False Then
                MsgBox "V‚½‚É•ÏXˆ—‚ğ‚µ‚Ä‚­‚¾‚³‚¢B"
            End If
        Else
            If henkouhannei(nid, ntan, txtnoldtanka.Text) = False Then
                MsgBox "V‚½‚É•ÏXˆ—‚ğ‚µ‚Ä‚­‚¾‚³‚¢B"
            End If
        End If
    End If
 End If
On Error GoTo 0
     
Unload frmkeiyaku2
Set frmkeiyaku2 = Nothing
shuruiset
frmkeiyaku.Show 1

End With
Exit Sub

JIKKOUERROR:
    MsgBox "XV‚É¸”s‚µ‚Ü‚µ‚½¡"
    Exit Sub

End Sub

Private Sub SSCSHUURYOU_Click()
    
End Sub


Private Sub txtnewsetsumei_GotFocus()
txtnewsetsumei.SelStart = 0
txtnewsetsumei.SelLength = Len(txtnewsetsumei.Text)
End Sub

Private Sub txtnewsetsumei_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub

Private Sub txtnewshuruimei_GotFocus()
txtnewshuruimei.SelStart = 0
txtnewshuruimei.SelLength = Len(txtnewshuruimei.Text)
End Sub

Private Sub txtnewshuruimei_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtnewtanka.SetFocus
End If
End Sub

Private Sub txtnewtanka_GotFocus()
txtnewtanka.SelStart = 0
txtnewtanka.SelLength = Len(txtnewtanka.Text)
End Sub

Private Sub txtnewtanka_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtnewsetsumei.SetFocus
End If

End Sub

Private Sub txtnewtanka_LostFocus()
Dim nkakaku As Long

On Error GoTo errnt
    nkakaku = CLng(txtnewtanka.Text)
On Error GoTo 0



Exit Sub

errnt:
    MsgBox "”¼Šp”š‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B"
    txtnewtanka.SetFocus
    Exit Sub
End Sub
