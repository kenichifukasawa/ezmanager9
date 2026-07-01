VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmjoukyou 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "ìoò^èÛãµ"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.Frame Frame1 
      Caption         =   "åªç›ÇÃìoò^èÛãµ"
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
      Height          =   11085
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   15045
      Begin VB.Frame Frame2 
         Caption         =   "ã‚çsï "
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
         Height          =   2400
         Index           =   0
         Left            =   7470
         TabIndex        =   58
         Top             =   8580
         Width           =   7365
         Begin MSFlexGridLib.MSFlexGrid bankbetsugrid 
            Height          =   1995
            Index           =   0
            Left            =   120
            TabIndex        =   59
            Top             =   330
            Width           =   7000
            _ExtentX        =   12356
            _ExtentY        =   3519
            _Version        =   393216
            Cols            =   3
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "ÇªÇÃëºÇÃå_ñÒï "
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
         Height          =   5805
         Index           =   0
         Left            =   7455
         TabIndex        =   44
         Top             =   300
         Width           =   7365
         Begin VB.ComboBox comtsuki 
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
            Index           =   1
            Left            =   6060
            Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
            TabIndex        =   119
            Top             =   300
            Width           =   735
         End
         Begin VB.ComboBox comnen 
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
            Index           =   1
            Left            =   4740
            Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
            TabIndex        =   118
            Top             =   300
            Width           =   915
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Ã◊Øƒ
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'Ç»Çµ
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   5910
            MouseIcon       =   "frmjoukyou.frx":0000
            MousePointer    =   99  '’∞ªﬁ∞íËã`
            Picture         =   "frmjoukyou.frx":08FF
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   62
            Top             =   870
            Width           =   480
         End
         Begin VB.Frame Frame5 
            Caption         =   "ÇªÇÃëºÇÃå_ñÒï "
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
            Height          =   3915
            Index           =   0
            Left            =   210
            TabIndex        =   47
            Top             =   1770
            Width           =   7035
            Begin MSFlexGridLib.MSFlexGrid sonotakeiyakubetsugrid 
               Height          =   3465
               Left            =   120
               TabIndex        =   48
               Top             =   360
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   6112
               _Version        =   393216
               Cols            =   3
            End
         End
         Begin VB.ComboBox comnen 
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
            Left            =   2010
            Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
            TabIndex        =   46
            Top             =   300
            Width           =   915
         End
         Begin VB.ComboBox comtsuki 
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
            Left            =   3330
            Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
            TabIndex        =   45
            Top             =   300
            Width           =   735
         End
         Begin VB.Label SSCJIKKOU 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "ÇªÇÃëºÇÃå_ñÒèWåv"
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
            Left            =   5190
            TabIndex        =   63
            Top             =   1380
            Width           =   1905
         End
         Begin VB.Label Label32 
            Caption         =   "â~"
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
            Height          =   315
            Index           =   0
            Left            =   4350
            TabIndex        =   57
            Top             =   1320
            Width           =   315
         End
         Begin VB.Label Label31 
            Caption         =   "â~"
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
            Height          =   315
            Index           =   0
            Left            =   4380
            TabIndex        =   56
            Top             =   840
            Width           =   315
         End
         Begin VB.Label Label8 
            Caption         =   "èWåvîNåééwíËÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   210
            TabIndex        =   53
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label Label13 
            Caption         =   "îN"
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
            Height          =   315
            Index           =   0
            Left            =   3000
            TabIndex        =   52
            Top             =   420
            Width           =   315
         End
         Begin VB.Label Label17 
            Caption         =   "åéÅ`"
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
            Height          =   315
            Index           =   0
            Left            =   4110
            TabIndex        =   51
            Top             =   420
            Width           =   555
         End
         Begin VB.Label lbl22 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   2850
            TabIndex        =   50
            Top             =   1260
            Width           =   1335
         End
         Begin VB.Label lbl11 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   2880
            TabIndex        =   49
            Top             =   780
            Width           =   1275
         End
         Begin VB.Label Label28 
            Caption         =   "CATVä÷òAëçã‡äzÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   55
            Top             =   810
            Width           =   2115
         End
         Begin VB.Label Label27 
            Caption         =   "≤›¿∞»Øƒä÷òAëçã‡äzÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   54
            Top             =   1290
            Width           =   2415
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'ïsìßñæ
            BorderStyle     =   0  'ìßñæ
            Height          =   975
            Left            =   5070
            Top             =   720
            Width           =   2100
         End
         Begin VB.Label Label13 
            Caption         =   "îN"
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
            Height          =   315
            Index           =   1
            Left            =   5730
            TabIndex        =   120
            Top             =   420
            Width           =   315
         End
         Begin VB.Label Label17 
            Caption         =   "åé"
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
            Height          =   315
            Index           =   1
            Left            =   6840
            TabIndex        =   121
            Top             =   420
            Width           =   300
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "éÂå_ñÒï "
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
         Height          =   2400
         Index           =   0
         Left            =   7455
         TabIndex        =   42
         Top             =   6150
         Width           =   7365
         Begin MSFlexGridLib.MSFlexGrid shukeitakubetsugrid 
            Height          =   2000
            Index           =   0
            Left            =   120
            TabIndex        =   43
            Top             =   300
            Width           =   7000
            _ExtentX        =   12356
            _ExtentY        =   3519
            _Version        =   393216
            Cols            =   3
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "èWåv"
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
         Height          =   10695
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   7185
         Begin VB.PictureBox Picture13 
            Appearance      =   0  'Ã◊Øƒ
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'Ç»Çµ
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   6060
            MouseIcon       =   "frmjoukyou.frx":11FE
            MousePointer    =   99  '’∞ªﬁ∞íËã`
            Picture         =   "frmjoukyou.frx":1AFD
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   60
            Top             =   720
            Width           =   480
         End
         Begin VB.Label lblstbsuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1925
            TabIndex        =   82
            ToolTipText     =   "ìoò^çœÇ›STBê›íuå_ñÒêîÇ≈Ç∑ÅB"
            Top             =   3351
            Width           =   700
         End
         Begin VB.Label lblblack 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1925
            TabIndex        =   79
            ToolTipText     =   "ÉuÉâÉbÉNëŒè€ÇÃÉåÉRÅ[ÉhëçêîÇ≈Ç∑ÅB"
            Top             =   2894
            Width           =   700
         End
         Begin VB.Label lblaptvsuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1920
            TabIndex        =   21
            ToolTipText     =   "ëçTVêîçáåvÇ≈Ç∑ÅB"
            Top             =   3810
            Width           =   705
         End
         Begin VB.Label lblnetkousokuoption 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1920
            TabIndex        =   91
            ToolTipText     =   "çÇë¨ÉIÉvÉVÉáÉìå_ñÒå˚êîçáåvÅiå_ñÒIDÇ™Åu03ÅvÇ≈â“ìÆÇµÇƒÇ¢ÇÈå˚êîÅjÇ≈Ç∑ÅB"
            Top             =   7083
            Width           =   705
         End
         Begin VB.Label lblnetkyuushisuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1920
            TabIndex        =   88
            ToolTipText     =   "ÉCÉìÉ^Å[ÉlÉbÉgãxé~å˚êîçáåvÅiå_ñÒÇhÇcÅu02ÅvÇ≈ãxé~ÇÃå˚êîÅjÇ≈Ç∑ÅB"
            Top             =   6621
            Width           =   705
         End
         Begin VB.Label lblnetkaiyakusuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1920
            TabIndex        =   85
            ToolTipText     =   "ÉCÉìÉ^Å[ÉlÉbÉgâñÒå˚êîçáåvÅiå_ñÒÇhÇcÅu02ÅvÇ≈íÜé~ÇÃå˚êîÅjÇ≈Ç∑ÅB"
            Top             =   6159
            Width           =   705
         End
         Begin VB.Label lblcmsuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   34
            ToolTipText     =   "ÉPÅ[ÉuÉãÉÇÉfÉÄÇPÇ©ÇÁÇRÇÃçáåvêîÇ≈Ç∑ÅB"
            Top             =   8010
            Width           =   705
         End
         Begin VB.Label lblnetnomi 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   3
            ToolTipText     =   "ÉCÉìÉ^Å[ÉlÉbÉgÇÃÇ›Åiê¢ë—êîÇ™1à»è„Ç≈TVêîÇ™0ÇÃÇ‡ÇÃÅjÇÃÉåÉRÅ[ÉhêîÇ≈Ç∑ÅB"
            Top             =   7545
            Width           =   705
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   25
            Left            =   6690
            TabIndex        =   117
            Top             =   8115
            Width           =   315
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   24
            Left            =   6690
            TabIndex        =   116
            Top             =   7710
            Width           =   315
         End
         Begin VB.Label lblkubunsuu2 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   6
            Left            =   5490
            TabIndex        =   115
            ToolTipText     =   "äeèZëÓãÊï™ê¢ë—ëçêîÅiTVêîÇ™1à»è„Ç≈â“ìÆÇµÇƒÇ¢ÇÈê¢ë—êîÅjÇ≈Ç∑ÅB"
            Top             =   7680
            Width           =   1155
         End
         Begin VB.Label lblkubunsuu2 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   7
            Left            =   5490
            TabIndex        =   114
            ToolTipText     =   "äeèZëÓãÊï™ê¢ë—ëçêîÅiTVêîÇ™1à»è„Ç≈â“ìÆÇµÇƒÇ¢ÇÈê¢ë—êîÅjÇ≈Ç∑ÅB"
            Top             =   8085
            Width           =   1155
         End
         Begin VB.Label Label7 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "ÇªÇÃëºê¢ë—êîÅF"
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
            Height          =   315
            Index           =   5
            Left            =   3420
            TabIndex        =   113
            Top             =   8115
            Width           =   2055
         End
         Begin VB.Label Label9 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "óVóóâÄê¢ë—êîÅF"
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
            Height          =   315
            Index           =   5
            Left            =   3450
            TabIndex        =   112
            Top             =   7725
            Width           =   2025
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   23
            Left            =   6690
            TabIndex        =   111
            Top             =   7305
            Width           =   315
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   22
            Left            =   6690
            TabIndex        =   110
            Top             =   6900
            Width           =   315
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   21
            Left            =   6690
            TabIndex        =   109
            Top             =   6480
            Width           =   315
         End
         Begin VB.Label lblkubunsuu2 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   4
            Left            =   5490
            TabIndex        =   108
            ToolTipText     =   "äeèZëÓãÊï™ê¢ë—ëçêîÅiTVêîÇ™1à»è„Ç≈â“ìÆÇµÇƒÇ¢ÇÈê¢ë—êîÅjÇ≈Ç∑ÅB"
            Top             =   6870
            Width           =   1155
         End
         Begin VB.Label lblkubunsuu2 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   5
            Left            =   5490
            TabIndex        =   107
            ToolTipText     =   "äeèZëÓãÊï™ê¢ë—ëçêîÅiTVêîÇ™1à»è„Ç≈â“ìÆÇµÇƒÇ¢ÇÈê¢ë—êîÅjÇ≈Ç∑ÅB"
            Top             =   7275
            Width           =   1155
         End
         Begin VB.Label lblkubunsuu2 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   3
            Left            =   5490
            TabIndex        =   106
            ToolTipText     =   "äeèZëÓãÊï™ê¢ë—ëçêîÅiTVêîÇ™1à»è„Ç≈â“ìÆÇµÇƒÇ¢ÇÈê¢ë—êîÅjÇ≈Ç∑ÅB"
            Top             =   6465
            Width           =   1155
         End
         Begin VB.Label Label7 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "î_ÅEéñã∆èäê¢ë—êîÅF"
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
            Height          =   315
            Index           =   4
            Left            =   3420
            TabIndex        =   105
            Top             =   7312
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "åˆã§á@ê¢ë—êîÅF"
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
            Height          =   315
            Index           =   3
            Left            =   3780
            TabIndex        =   104
            Top             =   6480
            Width           =   1695
         End
         Begin VB.Label Label9 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "èWçáèZëÓê¢ë—êîÅF"
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
            Height          =   315
            Index           =   4
            Left            =   3450
            TabIndex        =   103
            Top             =   6900
            Width           =   2025
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   7
            Left            =   6690
            TabIndex        =   102
            Top             =   6075
            Width           =   315
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   3
            Left            =   6690
            TabIndex        =   101
            Top             =   5670
            Width           =   315
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   2
            Left            =   6690
            TabIndex        =   100
            Top             =   5265
            Width           =   315
         End
         Begin VB.Label lblkubunsuu2 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   1
            Left            =   5490
            TabIndex        =   99
            ToolTipText     =   "äeèZëÓãÊï™ê¢ë—ëçêîÅiTVêîÇ™1à»è„Ç≈â“ìÆÇµÇƒÇ¢ÇÈê¢ë—êîÅjÇ≈Ç∑ÅB"
            Top             =   5655
            Width           =   1155
         End
         Begin VB.Label lblkubunsuu2 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   2
            Left            =   5490
            TabIndex        =   98
            ToolTipText     =   "äeèZëÓãÊï™ê¢ë—ëçêîÅiTVêîÇ™1à»è„Ç≈â“ìÆÇµÇƒÇ¢ÇÈê¢ë—êîÅjÇ≈Ç∑ÅB"
            Top             =   6060
            Width           =   1155
         End
         Begin VB.Label lblkubunsuu2 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   5490
            TabIndex        =   97
            ToolTipText     =   "äeèZëÓãÊï™ê¢ë—ëçêîÅiTVêîÇ™1à»è„Ç≈â“ìÆÇµÇƒÇ¢ÇÈê¢ë—êîÅjÇ≈Ç∑ÅB"
            Top             =   5250
            Width           =   1155
         End
         Begin VB.Label Label7 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "åˆã§ê¢ë—êîÅF"
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
            Height          =   315
            Index           =   3
            Left            =   3420
            TabIndex        =   96
            Top             =   6075
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "àÍî â∆íÎê¢ë—êîÅF"
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
            Height          =   315
            Index           =   2
            Left            =   3420
            TabIndex        =   95
            Top             =   5265
            Width           =   2055
         End
         Begin VB.Label Label9 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "àÍî éññ±èäê¢ë—êîÅF"
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
            Height          =   315
            Index           =   3
            Left            =   3360
            TabIndex        =   94
            Top             =   5685
            Width           =   2115
         End
         Begin VB.Label Label18 
            Caption         =   "çÇë¨µÃﬂºÆ›êîÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   4
            Left            =   150
            TabIndex        =   93
            Top             =   7103
            Width           =   1815
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   20
            Left            =   2730
            TabIndex        =   92
            Top             =   7158
            Width           =   315
         End
         Begin VB.Label Label18 
            Caption         =   "≤›¿∞»Øƒãxé~êîÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   3
            Left            =   150
            TabIndex        =   90
            Top             =   6636
            Width           =   1815
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   19
            Left            =   2730
            TabIndex        =   89
            Top             =   6702
            Width           =   315
         End
         Begin VB.Label Label18 
            Caption         =   "≤›¿∞»ØƒâñÒêîÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   87
            Top             =   6169
            Width           =   1815
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   10
            Left            =   2730
            TabIndex        =   86
            Top             =   6246
            Width           =   315
         End
         Begin VB.Label Label18 
            Caption         =   "STBê›íuë‰êîÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   2
            Left            =   150
            TabIndex        =   84
            Top             =   3351
            Width           =   1815
         End
         Begin VB.Label Label24 
            Caption         =   "å¬"
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
            Height          =   315
            Index           =   18
            Left            =   2730
            TabIndex        =   83
            Top             =   3380
            Width           =   315
         End
         Begin VB.Label Label18 
            Caption         =   "èúñºàµÇ¢êîÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   1
            Left            =   150
            TabIndex        =   81
            Top             =   2894
            Width           =   1815
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   17
            Left            =   2730
            TabIndex        =   80
            Top             =   2920
            Width           =   315
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   16
            Left            =   6690
            TabIndex        =   78
            Top             =   4800
            Width           =   315
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   15
            Left            =   6690
            TabIndex        =   77
            Top             =   4395
            Width           =   315
         End
         Begin VB.Label lblkubunsuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   6
            Left            =   5490
            TabIndex        =   76
            ToolTipText     =   "äeèZëÓãÊï™ÉåÉRÅ[ÉhëçêîÅiTVêîÇ™1à»è„Ç≈â“ìÆÇµÇƒÇ¢ÇÈÉåÉRÅ[ÉhêîÅjÇ≈Ç∑ÅB"
            Top             =   4365
            Width           =   1155
         End
         Begin VB.Label lblkubunsuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   7
            Left            =   5490
            TabIndex        =   75
            ToolTipText     =   "äeèZëÓãÊï™ÉåÉRÅ[ÉhëçêîÅiTVêîÇ™1à»è„Ç≈â“ìÆÇµÇƒÇ¢ÇÈÉåÉRÅ[ÉhêîÅjÇ≈Ç∑ÅB"
            Top             =   4770
            Width           =   1155
         End
         Begin VB.Label Label7 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "ÇªÇÃëºå_ñÒêîÅF"
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
            Height          =   315
            Index           =   2
            Left            =   3420
            TabIndex        =   74
            Top             =   4800
            Width           =   2055
         End
         Begin VB.Label Label9 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "óVóóâÄå_ñÒêîÅF"
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
            Height          =   315
            Index           =   2
            Left            =   3450
            TabIndex        =   73
            Top             =   4410
            Width           =   2025
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   14
            Left            =   6690
            TabIndex        =   72
            Top             =   3990
            Width           =   315
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   13
            Left            =   6690
            TabIndex        =   71
            Top             =   3585
            Width           =   315
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   12
            Left            =   6690
            TabIndex        =   70
            Top             =   3165
            Width           =   315
         End
         Begin VB.Label lblkubunsuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   4
            Left            =   5490
            TabIndex        =   69
            ToolTipText     =   "äeèZëÓãÊï™ÉåÉRÅ[ÉhëçêîÅiTVêîÇ™1à»è„Ç≈â“ìÆÇµÇƒÇ¢ÇÈÉåÉRÅ[ÉhêîÅjÇ≈Ç∑ÅB"
            Top             =   3555
            Width           =   1155
         End
         Begin VB.Label lblkubunsuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   5
            Left            =   5490
            TabIndex        =   68
            ToolTipText     =   "äeèZëÓãÊï™ÉåÉRÅ[ÉhëçêîÅiTVêîÇ™1à»è„Ç≈â“ìÆÇµÇƒÇ¢ÇÈÉåÉRÅ[ÉhêîÅjÇ≈Ç∑ÅB"
            Top             =   3960
            Width           =   1155
         End
         Begin VB.Label lblkubunsuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   3
            Left            =   5490
            TabIndex        =   67
            ToolTipText     =   "äeèZëÓãÊï™ÉåÉRÅ[ÉhëçêîÅiTVêîÇ™1à»è„Ç≈â“ìÆÇµÇƒÇ¢ÇÈÉåÉRÅ[ÉhêîÅjÇ≈Ç∑ÅB"
            Top             =   3150
            Width           =   1155
         End
         Begin VB.Label Label7 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "î_ÅEéñã∆èäå_ñÒêîÅF"
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
            Height          =   315
            Index           =   1
            Left            =   3420
            TabIndex        =   66
            Top             =   3975
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "åˆã§á@å_ñÒêîÅF"
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
            Height          =   315
            Index           =   1
            Left            =   3780
            TabIndex        =   65
            Top             =   3165
            Width           =   1695
         End
         Begin VB.Label Label9 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "èWçáèZëÓå_ñÒêîÅF"
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
            Height          =   315
            Index           =   1
            Left            =   3450
            TabIndex        =   64
            Top             =   3585
            Width           =   2025
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
            Left            =   6090
            TabIndex        =   61
            Top             =   1230
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "ëçå_ñÒêîÅF"
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
            Index           =   0
            Left            =   90
            TabIndex        =   37
            Top             =   300
            Width           =   1635
         End
         Begin VB.Label lblhoshoukinht 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   4845
            TabIndex        =   33
            Top             =   9480
            Width           =   1635
         End
         Begin VB.Label lblintersougaku 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   825
            TabIndex        =   32
            ToolTipText     =   "éÂå_ñÒï èWåvÇÃÉCÉìÉ^Å[ÉlÉbÉgä÷òAÇÃçáåväzÇ≈Ç∑ÅB"
            Top             =   5702
            Width           =   1800
         End
         Begin VB.Label lblcatvsougaku 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   825
            TabIndex        =   31
            ToolTipText     =   "éÂå_ñÒï èWåvÇÃCATVä÷òAÇÃçáåväzÇ≈Ç∑ÅB"
            Top             =   2437
            Width           =   1800
         End
         Begin VB.Label lblbanksuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   1590
            TabIndex        =   30
            ToolTipText     =   "éxï•ï˚ñ@Ç™ã‚çsÇÃÉåÉRÅ[ÉhëçêîÇ≈Ç∑ÅB"
            Top             =   10140
            Width           =   1155
         End
         Begin VB.Label lblgenkinsuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   5130
            TabIndex        =   29
            ToolTipText     =   "éxï•ï˚ñ@Ç™åªã‡ÇÃÉåÉRÅ[ÉhëçêîÇ≈Ç∑ÅB"
            Top             =   10185
            Width           =   1155
         End
         Begin VB.Label Label21 
            Caption         =   "(ëSå_ñÒé“êî/å_ñÒé¿êîÅEãxé~å_ñÒêî)"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   315
            Index           =   0
            Left            =   1620
            TabIndex        =   28
            Top             =   300
            Width           =   3855
         End
         Begin VB.Label Label9 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "àÍî éññ±èäå_ñÒêîÅF"
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
            Height          =   315
            Index           =   0
            Left            =   3360
            TabIndex        =   27
            Top             =   2370
            Width           =   2115
         End
         Begin VB.Label Label2 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "àÍî â∆íÎå_ñÒêîÅF"
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
            Height          =   315
            Index           =   0
            Left            =   3420
            TabIndex        =   26
            Top             =   1950
            Width           =   2055
         End
         Begin VB.Label Label3 
            Caption         =   "ëçê¢ë—êîÅF"
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
            Index           =   0
            Left            =   90
            TabIndex        =   25
            Top             =   1080
            Width           =   1395
         End
         Begin VB.Label Label7 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "åˆã§å_ñÒêîÅF"
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
            Height          =   315
            Index           =   0
            Left            =   3420
            TabIndex        =   24
            Top             =   2760
            Width           =   2055
         End
         Begin VB.Label lblsetaisousuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   23
            ToolTipText     =   "ëSê¢ë—ëçêîÇ≈Ç∑ÅB"
            Top             =   1380
            Width           =   1155
         End
         Begin VB.Label lblkeiyakusousuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   22
            ToolTipText     =   "ëSÉåÉRÅ[ÉhëçêîÇ≈Ç∑ÅB"
            Top             =   660
            Width           =   1155
         End
         Begin VB.Label lblkubunsuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   5490
            TabIndex        =   20
            ToolTipText     =   "äeèZëÓãÊï™ÉåÉRÅ[ÉhëçêîÅiTVêîÇ™1à»è„Ç≈â“ìÆÇµÇƒÇ¢ÇÈÉåÉRÅ[ÉhêîÅjÇ≈Ç∑ÅB"
            Top             =   1935
            Width           =   1155
         End
         Begin VB.Label lblkubunsuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   2
            Left            =   5490
            TabIndex        =   19
            ToolTipText     =   "äeèZëÓãÊï™ÉåÉRÅ[ÉhëçêîÅiTVêîÇ™1à»è„Ç≈â“ìÆÇµÇƒÇ¢ÇÈÉåÉRÅ[ÉhêîÅjÇ≈Ç∑ÅB"
            Top             =   2745
            Width           =   1155
         End
         Begin VB.Label lblkubunsuu 
            Alignment       =   1  'âEëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   1
            Left            =   5490
            TabIndex        =   18
            ToolTipText     =   "äeèZëÓãÊï™ÉåÉRÅ[ÉhëçêîÅiTVêîÇ™1à»è„Ç≈â“ìÆÇµÇƒÇ¢ÇÈÉåÉRÅ[ÉhêîÅjÇ≈Ç∑ÅB"
            Top             =   2340
            Width           =   1155
         End
         Begin VB.Label kyuushiwari 
            Alignment       =   2  'íÜâõëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   1500
            TabIndex        =   17
            ToolTipText     =   "â“ìÆíÜÇÃÉåÉRÅ[ÉhêîÇ∆ãxé~íÜÇÃÉåÉRÅ[ÉhêîÇ≈Ç∑ÅB"
            Top             =   660
            Width           =   3000
         End
         Begin VB.Label lblsousetaiwari 
            Alignment       =   2  'íÜâõëµÇ¶
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   1500
            TabIndex        =   16
            ToolTipText     =   "â“ìÆíÜÇÃê¢ë—êîÇ∆ãxé~íÜÇÃê¢ë—êîÇ≈Ç∑ÅB"
            Top             =   1380
            Width           =   3000
         End
         Begin VB.Label Label19 
            Caption         =   "â~"
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
            Height          =   315
            Index           =   0
            Left            =   2730
            TabIndex        =   15
            Top             =   2460
            Width           =   315
         End
         Begin VB.Label Label20 
            Caption         =   "â~"
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
            Height          =   315
            Index           =   0
            Left            =   2730
            TabIndex        =   14
            Top             =   5790
            Width           =   315
         End
         Begin VB.Label Label4 
            Caption         =   "â~"
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
            Height          =   315
            Index           =   0
            Left            =   6510
            TabIndex        =   13
            Top             =   9540
            Width           =   315
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   0
            Left            =   6690
            TabIndex        =   12
            Top             =   1950
            Width           =   315
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   1
            Left            =   6690
            TabIndex        =   11
            Top             =   2355
            Width           =   315
         End
         Begin VB.Label Label24 
            Caption         =   "å¬"
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
            Height          =   315
            Index           =   4
            Left            =   2730
            TabIndex        =   10
            Top             =   8070
            Width           =   315
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   5
            Left            =   6690
            TabIndex        =   9
            Top             =   2760
            Width           =   315
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   6
            Left            =   2730
            TabIndex        =   8
            Top             =   3840
            Width           =   315
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   8
            Left            =   2850
            TabIndex        =   7
            Top             =   10200
            Width           =   315
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   9
            Left            =   6390
            TabIndex        =   6
            Top             =   10260
            Width           =   315
         End
         Begin VB.Label Label26 
            Caption         =   "TVêîÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   5
            Top             =   3810
            Width           =   765
         End
         Begin VB.Label Label24 
            Caption         =   "åè"
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
            Height          =   315
            Index           =   11
            Left            =   2730
            TabIndex        =   4
            Top             =   7614
            Width           =   315
         End
         Begin VB.Label Label30 
            Caption         =   "≤›¿∞»ØƒÇÃÇ›êîÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   2
            Top             =   7570
            Width           =   1815
         End
         Begin VB.Label Label25 
            Caption         =   "ï€èÿã‡ÅiÇgÇsÅjÇÃçáåvÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   4050
            TabIndex        =   41
            Top             =   9210
            Width           =   2265
         End
         Begin VB.Label Label12 
            Caption         =   "≤›¿∞»Øƒä÷òAëçã‡äzÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   40
            Top             =   5235
            Width           =   2415
         End
         Begin VB.Label Label14 
            Caption         =   "ã‚çsêUë÷êîÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   210
            TabIndex        =   39
            Top             =   10200
            Width           =   1515
         End
         Begin VB.Label Label16 
            Caption         =   "åªã‡éxï•êîÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   3750
            TabIndex        =   38
            Top             =   10230
            Width           =   1515
         End
         Begin VB.Label Label6 
            Caption         =   "π∞ÃﬁŸ”√ﬁ—êîÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   36
            Top             =   8040
            Width           =   1755
         End
         Begin VB.Label Label15 
            Caption         =   "CATVä÷òAëçã‡äzÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   35
            Top             =   1980
            Width           =   2115
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'ïsìßñæ
            BorderStyle     =   0  'ìßñæ
            Height          =   1035
            Left            =   5610
            Top             =   600
            Width           =   1410
         End
      End
   End
End
Attribute VB_Name = "frmjoukyou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdend_Click()
    
End Sub

Private Sub Command1_Click()

End Sub
Sub soushuu2()
Dim shiteihajime As String, shiteiowari As String, intersoukin As Long, catvsoukin As Long
Dim ssssuu As Long, ststr As String

With frmjoukyou
    shiteihajime = .comnen(0).Text & .comtsuki(0).Text & "01"
    shiteiowari = .comnen(1).Text & .comtsuki(1).Text & "31"
'ÇªÇÃÇΩï 
    intersoukin = 0
    catvsoukin = 0
'sql = "SELECT  sonota.sonota, Count(keiyaku2.keiyakuid2) AS suu, Sum(CLng([keiyaku2].[kingaku])) AS kin, sonota.kankei" & _
    " FROM keiyaku2 RIGHT JOIN sonota ON keiyaku2.sonotaid = sonota.sonotaid" & _
    " WHERE keiyaku2.kaishi Between '" & shiteihajime & "' And '" & shiteiowari & "'" & _
    " GROUP BY sonota.sonotaid, sonota.kankei, sonota.sonota" & _
    " HAVING (((sonota.sonotaid)<>'98' And (sonota.sonotaid)<>'99'));"
sql = "SELECT sonota.sonota, Count(keiyaku2.keiyakuid2) AS suu, Sum(CLng([keiyaku2].[kingaku])) AS kin, sonota.kankei" & _
    " FROM (keiyaku2 RIGHT JOIN sonota ON keiyaku2.sonotaid = sonota.sonotaid)" & _
    " LEFT JOIN kojin ON keiyaku2.kojinid = kojin.kojinid" & _
    " WHERE ((keiyaku2.kaishi) Between '" & shiteihajime & "' And '" & shiteiowari & "'" & _
    ")" & _
    " GROUP BY sonota.sonotaid, sonota.kankei, sonota.sonota" & _
    " HAVING (((sonota.sonotaid)<>'98' And (sonota.sonotaid)<>'99'));"

frmjoukyou.sonotakeiyakubetsugrid.Font.Size = 11
intersoukin = 0
catvsoukin = 0
If FcSQlGet(rs, sql, PrMsg) = False Then
    frmjoukyou.sonotakeiyakubetsugrid.Col = 1
Else
    ssssuu = rs.RecordCount
    rs.MoveFirst
    With frmjoukyou.sonotakeiyakubetsugrid
       .Rows = ssssuu + 1
       .ColWidth(0) = 3000
       .ColWidth(1) = 1000
       .ColWidth(2) = 1200
       .Row = 0
       .Col = 0
       .Text = "ÇªÇÃëºÇÃå_ñÒñº"
       .Col = 1
       .Text = "åèêî"
       .Col = 2
       .Text = "ëçäz"
   End With
   i = 1
   Do Until rs.EOF
        With frmjoukyou.sonotakeiyakubetsugrid
            .Row = i
            .Col = 0
            .CellAlignment = 1
            .Text = rs!sonota
            .Col = 1
            .CellAlignment = 7
            .Text = rs!suu
            .Col = 2
            .CellAlignment = 7
            .Text = rs!kin
        End With
        If rs!kankei = "0" Then
            catvsoukin = catvsoukin + CLng(rs!kin)
        ElseIf rs!kankei = "1" Then
            intersoukin = intersoukin + CLng(rs!kin)
        End If
        rs.MoveNext
        i = i + 1
    Loop
    rs.Close
End If

frmjoukyou.lbl11(0).Caption = catvsoukin
frmjoukyou.lbl22(0).Caption = intersoukin







End With


End Sub
Sub soushuu1()
Dim keiyakusousuu As Long, sousetaisuu As Long, dansetaisuu As Long
Dim apsetaisuu As Long, aptvsuu As Long, cmsuu As Long, soubanksuu As Long
Dim sougenkinsuu As Long, souseikyuushosuu As Long, danchisousuu As Long
Dim apartsuu As Long, juutakusousuu As Long, juutakusetaisuu As Long, ssssuu As Long
Dim dansetai(2) As Long, dss(1) As String, sunshurs As Recordset, shukeiyakubutsudata()
Dim intersoukin As Long, catvsoukin As Long, hokin As Long, hokin2 As Long
Dim newnetsuu As Long, newstbsuu As Long, newblacksuu As String, newoptionsuu As Long
Dim netkyu As Long, netchu As Long, newkubunshuukei(7) As Long, suucounter As Integer

frmjoukyou.bankbetsugrid(0).Font.Size = 11
frmjoukyou.shukeitakubetsugrid(0).Font.Size = 11
FRMOMACHI.lblmsg.Caption = "èWåvíÜÇ≈Ç∑..."
FRMOMACHI.pb.Min = 0
FRMOMACHI.pb.Max = 17
FRMOMACHI.pb.Visible = True
FRMOMACHI.Show
FRMOMACHI.pb.Value = 0
DoEvents

Screen.MousePointer = 11

With frmjoukyou


'ëçêî

sql = "select * from kojin"

If FcSQlGet(rs, sql, PrMsg) = False Then
    keiyakusousuu = 0
Else
    keiyakusousuu = rs.RecordCount
    rs.Close
End If
.lblkeiyakusousuu(0).Caption = keiyakusousuu

FRMOMACHI.pb.Value = 1
DoEvents

'ëçêîÅiâ“ìÆÅEãxé~ÅEíÜé~Åj
sql = "SELECT * From kojin" & _
     " where f3 ='0'"
If FcSQlGet(rs, sql, PrMsg) = False Then
    dansetai(0) = 0
Else
    dansetai(0) = rs.RecordCount
    rs.Close
End If
FRMOMACHI.pb.Value = 2
DoEvents
sql = "SELECT * From kojin" & _
        " where f3 ='1'"

If FcSQlGet(rs, sql, PrMsg) = False Then
    dansetai(1) = 0
Else
    dansetai(1) = rs.RecordCount
    rs.Close
End If

dss(0) = "(â“ìÆ:" & dansetai(0) & ",ãxé~:" & dansetai(1) & ")"
.kyuushiwari(0).Caption = dss(0)

'≤›¿∞»ØƒÇÃÇ›êî
sql = "SELECT Count(kojin.kojinid) AS netsousuu From kojin" & _
    " WHERE CLng([setaisuu]>='1') AND CLng([tvsuu]='0')"

If FcSQlGet(rs, sql, PrMsg) = False Then
    newnetsuu = 0
Else
    If IsNull(rs!netsousuu) Then
        newnetsuu = 0
    Else
        newnetsuu = rs!netsousuu
    End If
    rs.Close
End If
 .lblnetnomi(0).Caption = newnetsuu


'ê¢ë—ëçêî
sql = "SELECT sum(clng(setaisuu)) AS sousetaisuu FROM kojin"

If FcSQlGet(rs, sql, PrMsg) = False Then
    sousetaisuu = 0
Else
    If IsNull(rs!sousetaisuu) Then
        sousetaisuu = 0
    Else
        sousetaisuu = rs!sousetaisuu
    End If
    rs.Close
End If
.lblsetaisousuu(0).Caption = sousetaisuu

'ëçê¢ë—êîêîÅiâ“ìÆÅEãxé~ÅEíÜé~Åj
sql = "SELECT Sum(CLng([setaisuu])) AS ds" & _
        " From kojin" & _
        " GROUP BY kojin.f3" & _
        " HAVING kojin.f3 ='0'"
If FcSQlGet(rs, sql, PrMsg) = False Then
    dansetai(0) = 0
Else
    dansetai(0) = rs!ds
    rs.Close
End If
FRMOMACHI.pb.Value = 3
DoEvents
    sql = "SELECT Sum(CLng([setaisuu])) AS ds" & _
        " From kojin" & _
        " GROUP BY kojin.f3" & _
        " HAVING (((kojin.f3)='1'));"
If FcSQlGet(rs, sql, PrMsg) = False Then
    dansetai(1) = 0
Else
    dansetai(1) = rs!ds
    rs.Close
End If
FRMOMACHI.pb.Value = 4
DoEvents

sql = "select count(kojinid) as nnksuu  from kojin where f2='1'"

If FcSQlGet(rs, sql, PrMsg) = False Then
    newblacksuu = 0
Else
    newblacksuu = rs!nnksuu
    rs.Close
End If

.lblblack.Caption = newblacksuu

FRMOMACHI.pb.Value = 5
DoEvents

dss(1) = "(â“ìÆ:" & dansetai(0) & ",ãxé~:" & dansetai(1) & ")"
.lblsousetaiwari(0).Caption = dss(1)

'çÇë¨µÃﬂºÆ›
FRMOMACHI.pb.Value = 6
DoEvents

sql = "select sum(clng([kuchisuu])) as nosuu  from keiyaku where shurui='03' and kadou='0'"

If FcSQlGet(rs, sql, PrMsg) = False Then
    newoptionsuu = 0
Else
    If rs.RecordCount = 0 Then
        newoptionsuu = 0
    Else
        If IsNull(rs!nosuu) Then
            newoptionsuu = 0
        Else
            newoptionsuu = rs!nosuu
        End If
    End If
    rs.Close
End If


.lblnetkousokuoption.Caption = newoptionsuu
'ÉlÉbÉgãxé~
FRMOMACHI.pb.Value = 7
DoEvents

sql = "select sum(clng([kuchisuu])) as nosuu2  from keiyaku where shurui='02' and kadou='1'"

If FcSQlGet(rs, sql, PrMsg) = False Then
    netkyu = 0
Else
    If IsNull(rs!nosuu2) Then
        netkyu = 0
    Else
        netkyu = rs!nosuu2
    End If
    rs.Close
End If

.lblnetkyuushisuu.Caption = netkyu
'ÉlÉbÉgíÜé~
FRMOMACHI.pb.Value = 8
DoEvents

sql = "select sum(clng([kuchisuu])) as nosuu3  from keiyaku where shurui='02' and kadou='2'"

If FcSQlGet(rs, sql, PrMsg) = False Then
    netchu = 0
Else
    If IsNull(rs!nosuu3) Then
    netchu = 0
    Else
    netchu = rs!nosuu3
    End If
    rs.Close
End If

.lblnetkaiyakusuu.Caption = netchu
'ãÊï™ï èWåv
sql = "SELECT Count(kojin.kojinid) AS netkubunsuu, kojin.juutaku From kojin" & _
    " WHERE (((kojin.f3)='0') AND ((kojin.tvsuu)<>'0')) GROUP BY kojin.juutaku;"
For suucounter = 0 To 7
    newkubunshuukei(suucounter) = 0
Next suucounter
If FcSQlGet(rs, sql, PrMsg) = True Then
    rs.MoveFirst
    Do Until rs.EOF
        newkubunshuukei(CInt(rs!juutaku)) = rs!netkubunsuu
        rs.MoveNext
    Loop
    rs.Close
End If

For suucounter = 0 To 7
    .lblkubunsuu(suucounter).Caption = newkubunshuukei(suucounter)
Next suucounter
sql = "SELECT Sum(CLng([setaisuu])) AS netkubunsuu2, kojin.juutaku From kojin" & _
    " WHERE (((kojin.f3)='0') AND ((kojin.tvsuu)<>'0')) GROUP BY kojin.juutaku;"
For suucounter = 0 To 7
    newkubunshuukei(suucounter) = 0
Next suucounter
If FcSQlGet(rs, sql, PrMsg) = True Then
    rs.MoveFirst
    Do Until rs.EOF
        newkubunshuukei(CInt(rs!juutaku)) = rs!netkubunsuu2
        rs.MoveNext
    Loop
    rs.Close
End If

For suucounter = 0 To 7
    .lblkubunsuu2(suucounter).Caption = newkubunshuukei(suucounter)
Next suucounter
'TVêî
FRMOMACHI.pb.Value = 10
DoEvents

        sql = "SELECT sum(clng(tvsuu)) AS tvs FROM kojin"
If FcSQlGet(rs, sql, PrMsg) = False Then
    apsetaisuu = 0
    aptvsuu = 0
Else
    If IsNull(rs!tvs) Then
        aptvsuu = 0
    Else
        aptvsuu = rs!tvs
    End If
    rs.Close
End If
.lblaptvsuu.Caption = aptvsuu

'STB
FRMOMACHI.pb.Value = 11
DoEvents

        sql = "SELECT count(stbkanriid) AS stbstb FROM stb"
If FcSQlGet(rs, sql, PrMsg) = False Then
    newstbsuu = 0
Else
    If IsNull(rs!stbstb) Then
        newstbsuu = 0
    Else
        newstbsuu = rs!stbstb
    End If
    rs.Close
End If

.lblstbsuu.Caption = newstbsuu


'π∞ÃﬁŸ”√ﬁ—ÇÃêî

    sql = "SELECT Count(kojin.it1) AS cm1 From kojin" & _
    " where kojin.it1 <> Space(1)"

If FcSQlGet(rs, sql, PrMsg) = False Then
    cmsuu = 0
Else
    cmsuu = rs!cm1
    rs.Close
End If
    sql = "SELECT Count(kojin.it2) AS cm2 From kojin" & _
    " where kojin.it2 <> Space(1)"
If FcSQlGet(rs, sql, PrMsg) = True Then
    cmsuu = cmsuu + rs!cm2
    rs.Close
End If

    sql = "SELECT Count(kojin.it3) AS cm3 From kojin" & _
    " where kojin.it3 <> Space(1)"
If FcSQlGet(rs, sql, PrMsg) = True Then
    cmsuu = cmsuu + rs!cm3
    rs.Close
End If

.lblcmsuu(0).Caption = cmsuu
FRMOMACHI.pb.Value = 12
DoEvents


'ã‚çsêî

    sql = "SELECT Count(kojin.seikyuu) AS banksousuu FROM kojin" & _
    " where kojin.seikyuu='0'"

If FcSQlGet(rs, sql, PrMsg) = False Then
    soubanksuu = 0
Else
    soubanksuu = rs!banksousuu
    rs.Close
End If
.lblbanksuu(0).Caption = soubanksuu
FRMOMACHI.pb.Value = 13
DoEvents

'åªã‡êî

    sql = "SELECT Count(kojin.seikyuu) AS genkinsousuu FROM kojin" & _
    " where kojin.seikyuu='1'"

If FcSQlGet(rs, sql, PrMsg) = False Then
    sougenkinsuu = 0
Else
    sougenkinsuu = rs!genkinsousuu
    rs.Close
End If

.lblgenkinsuu(0).Caption = sougenkinsuu
FRMOMACHI.pb.Value = 14
DoEvents

'ã‚çsï èWåv
FRMOMACHI.pb.Value = 15
DoEvents
    sql = "SELECT Count(kojin.kojinid) AS bankbetsu, bankmain.bankname, bank.bankmei" & _
    " FROM (bank RIGHT JOIN kojin ON bank.bankid = kojin.bankid)" & _
    " LEFT JOIN bankmain ON bank.bankcode = bankmain.bankcode" & _
    " GROUP BY bankmain.bankname, bank.bankmei, bank.bankcode," & _
    " bank.shitencode, kojin.seikyuu" & _
    " HAVING (((kojin.seikyuu)='0'));"
If FcSQlGet(rs, sql, PrMsg) = False Then
    frmjoukyou.bankbetsugrid(0).Col = 1
    
Else
    With frmjoukyou.bankbetsugrid(0)
        .Rows = rs.RecordCount + 1
        .ColWidth(0) = 2000
        .ColWidth(1) = 2000
        .ColWidth(2) = 700
        .Row = 0
        .Col = 0
        .Text = "ã‡óZã@ä÷ñº"
        .Col = 1
        .Text = "éxìXñº"
        .Col = 2
        .Text = "åèêî"
    End With
    rs.MoveFirst
    For i = 1 To rs.RecordCount
        With frmjoukyou.bankbetsugrid(0)
            .Row = i
            .Col = 0
            .Text = rs!bankname
            .Col = 1
            .Text = rs!bankmei
            .Col = 2
            .Text = rs!bankbetsu
        End With
       rs.MoveNext
    Next i
    rs.Close
End If
FRMOMACHI.pb.Value = 16
DoEvents
'éÂå_ñÒï 
    intersoukin = 0
    catvsoukin = 0
sql = "SELECT shurui.shuruiid, Sum(CLng([kuchisuu])) AS suu, shurui.shuruimei, shurui.kankei" & _
    " FROM (shurui RIGHT JOIN keiyaku ON shurui.shuruiid = keiyaku.shurui) LEFT JOIN kojin" & _
    " ON keiyaku.kojinid = kojin.kojinid" & _
    " GROUP BY shurui.shuruiid, shurui.shuruimei, shurui.kankei, keiyaku.kadou" & _
    " HAVING (((keiyaku.kadou)='0'))" & _
    " ORDER BY shurui.kankei;"
    
If FcSQlGet(rs, sql, PrMsg) = False Then
    frmjoukyou.shukeitakubetsugrid(0).Col = 1
Else
    ssssuu = rs.RecordCount
    ReDim shukeiyakubutsudata(ssssuu, 2)
    rs.MoveFirst
    i = 0
    Do Until rs.EOF
        shukeiyakubutsudata(i, 0) = rs!shuruimei
        shukeiyakubutsudata(i, 1) = rs!suu
        If rs!suu <> 0 Then
            sql = "SELECT shurui.shuruiid, Sum(CLng([kingaku])) AS shukeiyakukin" & _
                " FROM shurui LEFT JOIN keiyaku ON shurui.shuruiid = keiyaku.shurui" & _
                " GROUP BY shurui.shuruiid,keiyaku.kadou" & _
                " HAVING (((shurui.shuruiid)='" & rs!shuruiid & "')) and (((keiyaku.kadou)='0'));"
            If FcSQlGet(sunshurs, sql, PrMsg) = False Then
            
                shukeiyakubutsudata(i, 2) = "err"
            Else
                shukeiyakubutsudata(i, 2) = sunshurs!shukeiyakukin
                If rs!kankei = "0" Then
                    catvsoukin = catvsoukin + sunshurs!shukeiyakukin
                Else
                    intersoukin = intersoukin + sunshurs!shukeiyakukin
                End If
                sunshurs.Close
            End If
        Else
            shukeiyakubutsudata(i, 2) = "0"
        End If
        rs.MoveNext
        i = i + 1
    Loop
    With frmjoukyou.shukeitakubetsugrid(0)
        .Rows = ssssuu + 1
        .ColWidth(0) = 3000
        .ColWidth(1) = 1000
        .ColWidth(2) = 1200
        .Row = 0
        .Col = 0
        .Text = "éÂå_ñÒñº"
        .Col = 1
        .Text = "åèêî"
        .Col = 2
        .Text = "ëçäz"
    End With
    For i = 1 To ssssuu
        With frmjoukyou.shukeitakubetsugrid(0)
            .Row = i
            .Col = 0
            .CellAlignment = 1
            .Text = shukeiyakubutsudata(i - 1, 0)
            .Col = 1
            .CellAlignment = 7
            .Text = shukeiyakubutsudata(i - 1, 1)
            .Col = 2
            .CellAlignment = 7
            .Text = shukeiyakubutsudata(i - 1, 2)
        End With
    Next i
    
End If
'ï€èÿã‡ÇÃçáåv

    sql = "SELECT Sum(CCur([ht4])) AS hk2 From kojin WHERE kojin.ht4<>Space(1)"
If FcSQlGet(rs, sql, PrMsg) = False Then
    hokin2 = 0
Else
    If IsNull(rs!hk2) Then
        hokin2 = 0
    Else
        hokin2 = rs!hk2
    End If
    rs.Close
End If

.lblhoshoukinht.Caption = Format(hokin2, "\\#,##0")


.lblcatvsougaku.Caption = Format(catvsoukin, "\\#,##0")
.lblintersougaku.Caption = Format(intersoukin, "\\#,##0")


Dim keikanenn As Integer
keikanenn = CLng(Format(Date, "yyyy")) - 2006

For i = keikanenn To 0 Step -1
    .comnen(0).AddItem Format(Date, "yyyy") - i
    .comnen(1).AddItem Format(Date, "yyyy") - i

Next
 .comnen(0).ListIndex = keikanenn
  .comnen(1).ListIndex = keikanenn
    '.comnen(0).AddItem Format(Date, "yyyy") - 2
   ' .comnen(0).AddItem Format(Date, "yyyy") - 1
   ' .comnen(0).AddItem Format(Date, "yyyy")
   ' .comnen(0).ListIndex = 2
  '  .comnen(1).AddItem Format(Date, "yyyy") - 2
  '  .comnen(1).AddItem Format(Date, "yyyy") - 1
   ' .comnen(1).AddItem Format(Date, "yyyy")
   ' .comnen(1).ListIndex = 2
    For i = 1 To 12
        .comtsuki(0).AddItem Format(i, "00")
        .comtsuki(1).AddItem Format(i, "00")
    Next i
    .comtsuki(0).ListIndex = CInt(Format(Date, "mm")) - 1
    .comtsuki(1).ListIndex = CInt(Format(Date, "mm")) - 1
    
End With
Screen.MousePointer = 0
Unload FRMOMACHI
Set FRMOMACHI = Nothing
Picture7_Click
End Sub

Private Sub cmdsonotakeiyakushuukei_Click(Index As Integer)
End Sub

Private Sub Form_Load()

    soushuu1
End Sub

Private Sub Picture13_Click()
    Unload frmjoukyou
    Set frmjoukyou = Nothing

End Sub

Private Sub Picture7_Click()

soushuu2

End Sub
