VERSION 5.00
Begin VB.Form frmsonota2 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   ClientHeight    =   7935
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   9825
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8910
      MouseIcon       =   "frmsonota2.frx":0000
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmsonota2.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   29
      Top             =   5670
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8940
      MouseIcon       =   "frmsonota2.frx":11C9
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmsonota2.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   28
      Top             =   6900
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "ãåÇªÇÃëºÇÃå_ñÒè⁄ç◊"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3795
      Left            =   180
      TabIndex        =   6
      Top             =   60
      Width           =   8355
      Begin VB.Frame Frame3 
         Height          =   555
         Left            =   1200
         TabIndex        =   20
         Top             =   1560
         Width           =   7035
         Begin VB.OptionButton optoldbunrui 
            Caption         =   "CATVä÷òAóøã‡ï™óﬁ"
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
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   2295
         End
         Begin VB.OptionButton optoldbunrui 
            Caption         =   "ÉCÉìÉ^Å[ÉlÉbÉgä÷òAóøã‡ï™óﬁ"
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
            Left            =   3180
            TabIndex        =   21
            Top             =   120
            Width           =   3015
         End
      End
      Begin VB.Label lblchousei 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7020
         TabIndex        =   50
         Top             =   3150
         Width           =   1095
      End
      Begin VB.Label lblhyouki 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1380
         TabIndex        =   47
         Top             =   3150
         Width           =   4755
      End
      Begin VB.Label Label13 
         Caption         =   "ï\ãL(ï∂èÕ)"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   120
         TabIndex        =   46
         Top             =   3210
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "èWåvóLñ≥"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Left            =   150
         TabIndex        =   37
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblshuukei 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1290
         TabIndex        =   36
         Top             =   2430
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ÉOÉãÅ[Év"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Left            =   2805
         TabIndex        =   35
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblgroup 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3945
         TabIndex        =   34
         Top             =   2430
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ï¿Ç—èá"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Left            =   5460
         TabIndex        =   33
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblnarabe 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   32
         Top             =   2430
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "óøã‡ï™óﬁ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         TabIndex        =   23
         Top             =   1740
         Width           =   1395
      End
      Begin VB.Label lbloldsonotakotei 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   14
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "ÇªÇÃëºÇÃå_ñÒèÛãµ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "ÇªÇÃëºÇÃå_ñÒî‘çÜ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Left            =   60
         TabIndex        =   12
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label lbloldsonotaid 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "ÇªÇÃëºÇÃå_ñÒñº"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2220
         TabIndex        =   10
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label lbloldsonotamei 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1920
         TabIndex        =   9
         Top             =   540
         Width           =   4875
      End
      Begin VB.Label Label2 
         Caption         =   "ÇªÇÃëºÇÃå_ñÒã‡äz"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4320
         TabIndex        =   8
         Top             =   1200
         Width           =   1875
      End
      Begin VB.Label lbloldsonotakin 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0FFFF&
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
         Left            =   6300
         TabIndex        =   7
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label Label1 
         Caption         =   "í≤êÆ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Left            =   6390
         TabIndex        =   51
         Top             =   3240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "êVÇªÇÃëºÇÃå_ñÒè⁄ç◊"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   180
      TabIndex        =   4
      Top             =   4050
      Width           =   8355
      Begin VB.TextBox txtchousei 
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
         Left            =   7260
         MaxLength       =   1
         TabIndex        =   52
         ToolTipText     =   "êVå_ñÒíPâøÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
         Top             =   3090
         Width           =   855
      End
      Begin VB.TextBox txthyouki 
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
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   48
         ToolTipText     =   "ï\ãLñºÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
         Top             =   3090
         Width           =   4935
      End
      Begin VB.TextBox txtgroup 
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
         Left            =   5280
         MaxLength       =   1
         TabIndex        =   42
         ToolTipText     =   "êVå_ñÒíPâøÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtnarabe 
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
         Left            =   7260
         MaxLength       =   2
         TabIndex        =   41
         ToolTipText     =   "êVå_ñÒíPâøÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
         Top             =   2400
         Width           =   855
      End
      Begin VB.Frame Frame5 
         Height          =   585
         Left            =   1410
         TabIndex        =   38
         Top             =   2250
         Width           =   2445
         Begin VB.OptionButton optshuukei 
            Caption         =   "Ç»Çµ"
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   40
            Top             =   210
            Width           =   795
         End
         Begin VB.OptionButton optshuukei 
            Caption         =   "Ç†ÇË"
            Height          =   285
            Index           =   1
            Left            =   1260
            TabIndex        =   39
            Top             =   210
            Width           =   795
         End
      End
      Begin VB.Frame Frame4 
         Height          =   555
         Left            =   1260
         TabIndex        =   24
         Top             =   1560
         Width           =   6975
         Begin VB.OptionButton optnewbunrui 
            Caption         =   "ÉCÉìÉ^Å[ÉlÉbÉgä÷òAóøã‡ï™óﬁ"
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
            Left            =   3120
            TabIndex        =   26
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton optnewbunrui 
            Caption         =   "CATVä÷òAóøã‡ï™óﬁ"
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
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   120
            Value           =   -1  'True
            Width           =   2295
         End
      End
      Begin VB.Frame f3 
         Height          =   495
         Left            =   2220
         TabIndex        =   19
         Top             =   1020
         Width           =   2175
         Begin VB.OptionButton optkotei 
            Caption         =   "ïœìÆ"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   180
            Width           =   855
         End
         Begin VB.OptionButton optkotei 
            Caption         =   "å≈íË"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Top             =   180
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.TextBox txtnewsonotamei 
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
         Left            =   1980
         MaxLength       =   30
         TabIndex        =   0
         ToolTipText     =   "êVã‡óZã@ä÷ñºÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
         Top             =   540
         Width           =   4935
      End
      Begin VB.TextBox txtnewsonotakin 
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
         Left            =   6420
         MaxLength       =   7
         TabIndex        =   2
         ToolTipText     =   "êVã‡óZã@ä÷ÉRÅ[ÉhÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "í≤êÆ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Left            =   6570
         TabIndex        =   53
         Top             =   3150
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "ï\ãL(ï∂èÕ)"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         TabIndex        =   49
         Top             =   3150
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "èWåvóLñ≥"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Left            =   120
         TabIndex        =   45
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ÉOÉãÅ[Év"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Left            =   4245
         TabIndex        =   44
         Top             =   2430
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ï¿Ç—èá"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Left            =   6450
         TabIndex        =   43
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "óøã‡ï™óﬁ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "ÇªÇÃëºÇÃå_ñÒèÛãµ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   300
         TabIndex        =   18
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "ÇªÇÃëºÇÃå_ñÒî‘çÜ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   60
         TabIndex        =   17
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label5 
         Caption         =   "ÇªÇÃëºÇÃå_ñÒñº"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2220
         TabIndex        =   16
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label Label4 
         Caption         =   "ÇªÇÃëºÇÃå_ñÒã‡äz"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   1200
         Width           =   1875
      End
      Begin VB.Label lblnewsonotaid 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   540
         Width           =   1035
      End
   End
   Begin VB.Label SSCJIKKOU 
      AutoSize        =   -1  'True
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
      Height          =   225
      Left            =   8940
      TabIndex        =   31
      Top             =   6180
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
      Left            =   8970
      TabIndex        =   30
      Top             =   7410
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   7755
      Left            =   8640
      Top             =   90
      Width           =   1050
   End
End
Attribute VB_Name = "frmsonota2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub optkotei_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    txtnewsonotakin.SetFocus
End If
End Sub

Private Sub Picture13_Click()
    Unload frmsonota2
    Set frmsonota2 = Nothing
    frmsonota.Show 1

End Sub

Private Sub Picture7_Click()

Dim nid As String, nmei As String, nkin As String, nkote As String, nnid As String
Dim nkan As String

Dim s_shuukei As String, s_group As String, s_narabe As String

Dim s_hyoukimei As String, s_kasan As String

With frmsonota2

s_group = Trim(.txtgroup.Text)
s_narabe = Trim(.txtnarabe.Text)

s_shuukei = ""
If .optshuukei(0).Value = True Then
s_shuukei = "0"
ElseIf .optshuukei(1).Value = True Then
s_shuukei = "1"
Else
    MsgBox "èWåvÇïKÇ∏ëIëÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
   Exit Sub
End If

s_hyoukimei = Trim(.txthyouki.Text)
s_kasan = Trim(txtchousei.Text)



nid = Trim(.lblnewsonotaid.Caption)
nmei = Trim(.txtnewsonotamei.Text)
If .optkotei(0).Value = True Then
    nkote = "0"
    nkin = Trim(.txtnewsonotakin.Text)
    If nkin = "" Or nkin = "0" Then
        MsgBox "å≈íËÇÃèÍçáÇÕÅAã‡äzÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        .txtnewsonotakin.SetFocus
        Exit Sub
    End If
Else
    nkote = "1"
    nkin = Space(1)
End If

If .optnewbunrui(0).Value = True Then
    nkan = "0"
Else
    nkan = "1"
End If

If nmei = "" Then
   MsgBox "ÇªÇÃëºÇÃå_ñÒñºÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
   .txtnewsonotamei.SetFocus
   Exit Sub
End If


 On Error GoTo JIKKOUERROR3

 If .SSCJIKKOU.Caption = "ìoò^" Then
     
     sql = "INSERT INTO sonota VALUES('" & nid & "'," & _
                    "'" & nmei & "','" & nkin & "','" & nkote & "','" & nkan & "','" & s_shuukei & "','" & s_group & "','" & s_narabe & "','" & s_hyoukimei & "','" & s_kasan & "')"
     
     db.Execute sql, dbSQLPassThrough
     nnid = CStr(CLng(nid) + 1)
        If soukatsu(1, "renban", 5, nnid, 1, 1) = False Then
                MsgBox "ê›íËÇÃèëÇ´çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
                Exit Sub
        End If
     MsgBox "ìoò^ÇµÇ‹ÇµÇΩÅB"
 Else
 
     sql = "UPDATE sonota SET sonota = '" & nmei & "'," & _
                  "kingaku = '" & nkin & "',kotei = '" & nkote & "',kankei = '" & nkan & "',iru = '" & s_shuukei & "',gp = '" & s_group & "',narabe = '" & s_narabe & "',hyoukimei = '" & s_hyoukimei & "',kasan = '" & s_kasan & "'" & _
                  " WHERE sonotaid = '" & nid & "'"
     db.Execute sql, dbSQLPassThrough
    MsgBox "ïœçXÇµÇ‹ÇµÇΩÅB"
   
 End If
On Error GoTo 0
     
Unload frmsonota2
Set frmsonota2 = Nothing
sonotaset
frmsonota.Show 1

End With
Exit Sub

JIKKOUERROR3:
    MsgBox "çXêVÇ…é∏îsÇµÇ‹ÇµÇΩ°"
    Exit Sub


End Sub


Private Sub txtnewsonotakin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If
End Sub

Private Sub txtnewsonotamei_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    optkotei(0).SetFocus
End If
End Sub
