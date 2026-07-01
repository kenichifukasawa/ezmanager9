VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmnhk2 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "NHKÉ`ÉFÉbÉN"
   ClientHeight    =   11490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11490
   ScaleWidth      =   15270
   StartUpPosition =   1  'µ∞≈∞ Ã´∞—ÇÃíÜâõ
   Begin VB.Frame fhikakukekka2 
      Caption         =   "î‰äråãâ "
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   120
      TabIndex        =   44
      Top             =   3180
      Visible         =   0   'False
      Width           =   15015
      Begin VB.Frame Frame3 
         Caption         =   "âêÕÅiEzManagerÉfÅ[É^Ç™NHKì«çûÉfÅ[É^Ç…Ç†ÇÈÇ©ÇåüèÿÅjÅj"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3855
         Left            =   120
         TabIndex        =   46
         Top             =   4200
         Width           =   14775
         Begin MSFlexGridLib.MSFlexGrid gridkekka3 
            Height          =   2445
            Left            =   120
            TabIndex        =   64
            Top             =   1320
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   4313
            _Version        =   393216
            Cols            =   5
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
         Begin VB.Label lbl1m 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1680
            TabIndex        =   61
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label27 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "11890â~êî"
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
            Height          =   255
            Left            =   180
            TabIndex        =   60
            Top             =   300
            Width           =   1260
         End
         Begin VB.Label lbl2m 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5160
            TabIndex        =   59
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label25 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "5345â~êî"
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
            Height          =   255
            Left            =   3660
            TabIndex        =   58
            Top             =   300
            Width           =   1260
         End
         Begin VB.Label lbl3m 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8700
            TabIndex        =   57
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label23 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "5195â~êî"
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
            Height          =   255
            Left            =   7260
            TabIndex        =   56
            Top             =   300
            Width           =   1260
         End
         Begin VB.Label lbl11m 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2520
            TabIndex        =   55
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label lbl22m 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6000
            TabIndex        =   54
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label lbl33m 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   9540
            TabIndex        =   53
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label lbl4m 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   12240
            TabIndex        =   52
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label18 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "ÇªÇÃëºÇÃêî"
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
            Height          =   255
            Left            =   10860
            TabIndex        =   51
            Top             =   300
            Width           =   1260
         End
         Begin VB.Label lbl44m 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   13080
            TabIndex        =   50
            Top             =   300
            Width           =   1140
         End
         Begin VB.Line Line2 
            X1              =   10440
            X2              =   14220
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label lbl55m 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   13080
            TabIndex        =   49
            Top             =   900
            Width           =   1140
         End
         Begin VB.Label Label11 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "çáåvêî"
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
            Height          =   255
            Left            =   10620
            TabIndex        =   48
            Top             =   900
            Width           =   1260
         End
         Begin VB.Label lbl5m 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   12240
            TabIndex        =   47
            Top             =   900
            Width           =   705
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   14340
         MouseIcon       =   "frmnhk2.frx":0000
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmnhk2.frx":08FF
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   45
         Top             =   240
         Width           =   480
      End
      Begin VB.Frame Frame4 
         Caption         =   "âêÕÅiNHKì«çûÉfÅ[É^Ç™EzManagerÉfÅ[É^Ç…Ç†ÇÈÇ©ÇåüèÿÅj"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3795
         Left            =   120
         TabIndex        =   62
         Top             =   360
         Width           =   14775
         Begin MSFlexGridLib.MSFlexGrid gridkekka2 
            Height          =   3405
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   6006
            _Version        =   393216
            Cols            =   5
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
   End
   Begin VB.OptionButton optchk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "å⁄ãqÉfÅ[É^Ç™â¡éZÇ≥ÇÍÇƒÇ¢ÇÈÇ©ÇÉ`ÉFÉbÉN"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   9120
      TabIndex        =   41
      Top             =   900
      Width           =   3975
   End
   Begin VB.OptionButton optchk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "å⁄ãqÉfÅ[É^Ç™êøãÅÇ≥ÇÍÇƒÇ¢ÇÈÇ©ÇÉ`ÉFÉbÉN"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   9120
      TabIndex        =   40
      Top             =   660
      Width           =   3975
   End
   Begin VB.OptionButton optchk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "å⁄ãqÉfÅ[É^Ç∆NHKÉfÅ[É^ÇÃå›ä∑ê´ÇÉ`ÉFÉbÉN"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   9120
      TabIndex        =   39
      Top             =   420
      Width           =   3975
   End
   Begin VB.OptionButton optchk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HNKÉfÅ[É^Ç™êøãÅÇ≥ÇÍÇƒÇ¢ÇÈÇ©ÇÉ`ÉFÉbÉN"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9120
      TabIndex        =   38
      Top             =   180
      Width           =   3975
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   13440
      MouseIcon       =   "frmnhk2.frx":11C9
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmnhk2.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   36
      ToolTipText     =   "å_ñÒÇÃêøãÅèÛãµÇÉ`ÉFÉbÉNÇµÇ‹Ç∑ÅB"
      Top             =   330
      Width           =   480
   End
   Begin VB.Frame fhikakukekka 
      Caption         =   "î‰äråãâ "
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7995
      Left            =   120
      TabIndex        =   34
      Top             =   3120
      Visible         =   0   'False
      Width           =   14895
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   14160
         MouseIcon       =   "frmnhk2.frx":2392
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmnhk2.frx":2C91
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   43
         Top             =   180
         Width           =   480
      End
      Begin MSFlexGridLib.MSFlexGrid gridkekka 
         Height          =   7425
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   13097
         _Version        =   393216
         Cols            =   5
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
   Begin VB.Frame Frame1 
      Caption         =   "ì«Ç›çûÇ›ÉfÅ[É^"
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
      Height          =   9975
      Left            =   120
      TabIndex        =   7
      Top             =   1380
      Width           =   15015
      Begin VB.Frame Frame2 
         Caption         =   "âêÕ(NHKì«çûÉfÅ[É^Åj"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1395
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Width           =   14655
         Begin VB.Label lbl51 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   12480
            TabIndex        =   32
            Top             =   900
            Width           =   705
         End
         Begin VB.Label Label9 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "çáåvêî"
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
            Height          =   255
            Left            =   10860
            TabIndex        =   31
            Top             =   900
            Width           =   1260
         End
         Begin VB.Label lbl52 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   13320
            TabIndex        =   30
            Top             =   900
            Width           =   1140
         End
         Begin VB.Line Line1 
            X1              =   10740
            X2              =   14520
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label7 
            Caption         =   "êøãÅénåéêî"
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
            Height          =   255
            Left            =   3480
            TabIndex        =   29
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label lblseikyuusousuu2 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5280
            TabIndex        =   28
            Top             =   960
            Width           =   1140
         End
         Begin VB.Label lbl42 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   13320
            TabIndex        =   27
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label15 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "ÇªÇÃëºÇÃêî"
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
            Height          =   255
            Left            =   10860
            TabIndex        =   26
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label lbl41 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   12480
            TabIndex        =   25
            Top             =   240
            Width           =   705
         End
         Begin VB.Label lbl32 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   9300
            TabIndex        =   24
            Top             =   960
            Width           =   1140
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   9300
            TabIndex        =   23
            Top             =   600
            Width           =   1140
         End
         Begin VB.Label lbl12 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   9300
            TabIndex        =   22
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label10 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "5195â~êî"
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
            Height          =   255
            Left            =   6840
            TabIndex        =   21
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label lbl31 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8460
            TabIndex        =   20
            Top             =   960
            Width           =   705
         End
         Begin VB.Label Label8 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "5345â~êî"
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
            Height          =   255
            Left            =   6840
            TabIndex        =   19
            Top             =   600
            Width           =   1260
         End
         Begin VB.Label lbl21 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8460
            TabIndex        =   18
            Top             =   600
            Width           =   705
         End
         Begin VB.Label Label6 
            Alignment       =   1  'âEëµÇ¶
            Caption         =   "11890â~êî"
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
            Height          =   255
            Left            =   6840
            TabIndex        =   17
            Top             =   240
            Width           =   1260
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8460
            TabIndex        =   16
            Top             =   240
            Width           =   705
         End
         Begin VB.Label lblseikyuunengetsu 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1740
            TabIndex        =   15
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label4 
            Caption         =   "êøãÅîNåé"
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
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label lblseikyuusousuu 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5280
            TabIndex        =   13
            Top             =   600
            Width           =   1140
         End
         Begin VB.Label lblsousuu 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5280
            TabIndex        =   12
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label Label3 
            Caption         =   "êøãÅîNåéêî"
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
            Height          =   255
            Left            =   3480
            TabIndex        =   11
            Top             =   600
            Width           =   1260
         End
         Begin VB.Label Label2 
            Caption         =   "ëçêîÉåÉRÅ[Éhêî"
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
            Height          =   255
            Left            =   3480
            TabIndex        =   10
            Top             =   240
            Width           =   1785
         End
      End
      Begin MSFlexGridLib.MSFlexGrid gridnhk 
         Height          =   8025
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14155
         _Version        =   393216
         Cols            =   5
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
      Begin VB.Label lblok 
         Height          =   195
         Left            =   14700
         TabIndex        =   33
         Top             =   180
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   7500
      MouseIcon       =   "frmnhk2.frx":355B
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmnhk2.frx":3E25
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   330
      Width           =   480
   End
   Begin VB.TextBox txtdatapath 
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
      IMEMode         =   4  'ëSäpÇ–ÇÁÇ™Ç»
      Left            =   300
      TabIndex        =   4
      ToolTipText     =   "ÉfÅ[É^ÇÃì«çûÉtÉHÉãÉ_Å[ÇéwíËÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
      Top             =   540
      Width           =   7035
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8280
      MouseIcon       =   "frmnhk2.frx":46EF
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmnhk2.frx":4FEE
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      ToolTipText     =   "ÉfÅ[É^Çì«Ç›çûÇ›ÇµÇ‹Ç∑ÅB"
      Top             =   330
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   14370
      MouseIcon       =   "frmnhk2.frx":5E30
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmnhk2.frx":672F
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   330
      Width           =   480
   End
   Begin VB.Label Label73 
      Alignment       =   2  'íÜâõëµÇ¶
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "éQè∆"
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
      Left            =   7500
      TabIndex        =   42
      Top             =   870
      Width           =   495
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "é¿çs"
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
      Left            =   13500
      TabIndex        =   37
      Top             =   870
      Width           =   480
   End
   Begin VB.Label Label71 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÉfÅ[É^ÉpÉX"
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
      Height          =   255
      Left            =   300
      TabIndex        =   6
      Top             =   180
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FFFFFF&
      Caption         =   "ì«çû"
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
      Index           =   0
      Left            =   8040
      TabIndex        =   3
      Top             =   840
      Width           =   975
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
      Left            =   14400
      TabIndex        =   2
      Top             =   870
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   1155
      Left            =   120
      Top             =   120
      Width           =   15000
   End
End
Attribute VB_Name = "frmnhk2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
lblok.Caption = "NG"
End Sub

Private Sub gridkekka_DblClick()
gridkekka.Rows = 2
gridkekka.Clear
fhikakukekka.Visible = False

End Sub

Private Sub Label31_Click()
End Sub

Private Sub Picture1_Click()

Dim nhkfaile As String, nhkdata(), ColCount As Long, RowCount As Long, GetLine As Variant
Dim Row As Long, Col As Long, Spot As Integer, kariatai As String, ii As Long


If Trim(txtdatapath.Text) = "" Then
        ret = MsgBox("ÉtÉHÉãÉ_Å[ÇéwíËÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB", 16, "EZ Manager")
        End
End If

nhkfaile = Trim(txtdatapath.Text)

If Right(nhkfaile, 1) <> "\" Then
    nhkfaile = nhkfaile & "\"
End If

nhkfaile = nhkfaile & "IKKATUFD.dat"

If Dir(nhkfaile) = "" Then
        ret = MsgBox("ÉtÉ@ÉCÉãÇ™Ç†ÇËÇ‹ÇπÇÒÅBë∂ç›Ç∑ÇÈÉtÉHÉãÉ_Å[ÇéwíËÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB", 16, "EZ Manager")
        End
End If



On Error GoTo ErrorTrap3

Open nhkfaile For Input Lock Read Write As #1
ColCount = 0
RowCount = 41
 While Not EOF(1)
     Line Input #1, GetLine
     ColCount = ColCount + 1
 Wend
 Close #1


    
Open nhkfaile For Input Lock Read Write As #1



   ReDim nhkdata(ColCount, RowCount)
   Row = 0
   While Not EOF(1)
      Line Input #1, GetLine
      Col = 1
      
      Spot = InStr(1, GetLine, ",")
      While Spot <> 0
         kariatai = Left$(GetLine, Spot - 1)
         If Col <> 7 And Col <> 6 Then
            If Mid(kariatai, 1, 1) = """" Then
                If Right(kariatai, 1) = """" Then
                    nhkdata(Row, Col) = Mid(kariatai, 2, Len(kariatai) - 2)
                Else
                    nhkdata(Row, Col) = Mid(kariatai, 2)
                End If
            Else
                If Right(kariatai, 1) = """" Then
                    nhkdata(Row, Col) = Mid(kariatai, 1, Len(kariatai) - 1)
                Else
                    nhkdata(Row, Col) = kariatai
                End If
            End If
        Else
            nhkdata(Row, Col) = kariatai
        End If
         Col = Col + 1
         GetLine = Right$(GetLine, Len(GetLine) - Spot)
         Spot = InStr(1, GetLine, ",")
      Wend
      If Len(GetLine) <> 0 Then
            If Mid(GetLine, 1, 1) = """" Then
                If Right(GetLine, 1) = """" Then
                    nhkdata(Row, Col) = Mid(GetLine, 2, Len(GetLine) - 2)
                Else
                    nhkdata(Row, Col) = Mid(GetLine, 2)
                End If
            Else
                If Right(GetLine, 1) = """" Then
                    nhkdata(Row, Col) = Mid(GetLine, 1, Len(GetLine) - 1)
                Else
                    nhkdata(Row, Col) = GetLine
                End If
            End If
      End If
      Row = Row + 1
   Wend
   Close #1

Dim newseikyuusousuu As Long, shiteinenngetsu As String
Dim keisuu(8) As Long, keikin(8) As Long
Dim newseikyuusousuu2 As Long

shiteinenngetsu = Mid(Format(Date, "yyyymm"), 3)
lblseikyuunengetsu.Caption = shiteinenngetsu
shiteinenngetsu = "1110"

newseikyuusousuu = 0
newseikyuusousuu2 = 0
keisuu(0) = 0
keisuu(1) = 0
keisuu(2) = 0
keisuu(3) = 0
keikin(0) = 0
keikin(1) = 0
keikin(2) = 0
keikin(3) = 0

With gridnhk

nhk_grid_set (ColCount)

For i = 1 To ColCount - 1
    .Row = i
    For ii = 0 To 38
        .Col = ii
        If ii = 0 Then
            .Text = i
        Else
            .Text = nhkdata(i, ii)
        End If
        If ii = 14 Then
            Select Case CLng(nhkdata(i, ii))
            Case 11890
                keisuu(0) = keisuu(0) + 1
            Case 5345
                keisuu(1) = keisuu(1) + 1
            Case 5195
                keisuu(2) = keisuu(2) + 1
            Case Else
                keisuu(3) = keisuu(3) + 1
                keikin(3) = keikin(3) + CLng(nhkdata(i, ii))
            End Select
        End If
        If ii = 12 Then
            If Trim(nhkdata(i, ii)) = shiteinenngetsu Then
                newseikyuusousuu2 = newseikyuusousuu2 + 1
            End If
        End If
        
        If ii = 38 Then
            If Trim(nhkdata(i, ii)) = shiteinenngetsu Then
                newseikyuusousuu = newseikyuusousuu + 1
            End If
        End If
    Next ii
Next i

keikin(0) = keisuu(0) * 11890
keikin(1) = keisuu(1) * 5345
keikin(2) = keisuu(2) * 5195
keikin(4) = keikin(0) + keikin(1) + keikin(2) + keikin(3)
keisuu(4) = keisuu(0) + keisuu(1) + keisuu(2) + keisuu(3)
lblsousuu.Caption = ColCount - 1
lblseikyuusousuu.Caption = newseikyuusousuu
lblseikyuusousuu2.Caption = newseikyuusousuu2

lbl11.Caption = keisuu(0)
lbl21.Caption = keisuu(1)
lbl31.Caption = keisuu(2)
lbl41.Caption = keisuu(3)
lbl12.Caption = keikin(0)
lbl22.Caption = keikin(1)
lbl32.Caption = keikin(2)
lbl42.Caption = keikin(3)
lbl51.Caption = keisuu(4)
lbl52.Caption = keikin(4)

End With

lblok.Caption = "OK"

Exit Sub



ErrorTrap3:
    Screen.MousePointer = 0
    MsgBox "ì«çûÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
    lblok.Caption = "NG"

    Exit Sub


End Sub
Sub nhk_grid_set(retsu As Integer)
With frmnhk2.gridnhk
        .Font.Size = 11
        .Rows = retsu
        .Cols = 39
        .ColWidth(0) = 700
        .ColWidth(1) = 1100
        .ColWidth(2) = 1100
        .ColWidth(3) = 1100
        .ColWidth(4) = 1100
        .ColWidth(5) = 2000
        .ColWidth(6) = 2000
        .ColWidth(7) = 2000
        .ColWidth(8) = 1800
        .ColWidth(9) = 1100
        .ColWidth(10) = 1100
        .ColWidth(11) = 1100
        .ColWidth(12) = 1100
        .ColWidth(13) = 1100
        .ColWidth(14) = 2000
        .ColWidth(15) = 1100
        .ColWidth(16) = 1100
        .ColWidth(17) = 1100
        .ColWidth(18) = 1100
        .ColWidth(19) = 1100
        .ColWidth(20) = 1100
        .ColWidth(21) = 1100
        .ColWidth(22) = 1100
        .ColWidth(23) = 1100
        .ColWidth(24) = 1100
        .ColWidth(25) = 1100
        .ColWidth(26) = 1100
        .ColWidth(27) = 1100
        .ColWidth(28) = 1100
        .ColWidth(29) = 1100
        .ColWidth(30) = 1100
        .ColWidth(31) = 1100
        .ColWidth(32) = 1100
        .ColWidth(33) = 1100
        .ColWidth(34) = 2000
        .ColWidth(35) = 1100
        .ColWidth(36) = 1100
        .ColWidth(37) = 1100
        .ColWidth(38) = 1100
        .Row = 0
        .Col = 0
        .Text = "NO"
        .Col = 1
        .Text = "ã«ÉRÅ[Éh"
        .Col = 2
        .Text = "ésãÊåSÉRÅ[Éh"
        .Col = 3
        .Text = "ínàÊA"
        .Col = 4
        .Text = "ínàÊB"
        .Col = 5
        .Text = "ÉJÉiéÅñº"
        .Col = 6
        .Text = "äøéöéÅñº"
        .Col = 7
        .Text = "NHKÇ®ãqólî‘çÜ"
        .Col = 8
        .Text = "CATVé{ê›ÉRÅ[Éh"
        .Col = 9
        .Text = "ésäOã«î‘"
        .Col = 10
        .Text = "ésì‡ã«î‘"
        .Col = 11
        .Text = "â¡ì¸î‘çÜ"
        .Col = 12
        .Text = "êøãÅénåé"
        .Col = 13
        .Text = "èIì¸îNåé"
        .Col = 14
        .Text = "êøãÅã‡äz"
        .Col = 15
        .Text = "ê¥éZã‡äz"
        .Col = 16
        .Text = "å˚êî"
        .Col = 17
        .Text = "éxï•ï˚ñ@"
        .Col = 18
        .Text = "éxï•ÉRÅ[ÉX"
        .Col = 19
        .Text = "ÉTÉCÉNÉãénåé"
        .Col = 20
        .Text = "F2"
        .Col = 21
        .Text = "éüâÒéxï•ÉRÅ[ÉX"
        .Col = 22
        .Text = "BSC"
        .Col = 23
        .Text = "BSM"
        .Col = 24
        .Text = "BST"
        .Col = 25
        .Text = "ínè„C"
        .Col = 26
        .Text = "ínè„M"
        .Col = 27
        .Text = "F3"
        .Col = 28
        .Text = "èïê¨ñ∆èúID"
        .Col = 29
        .Text = "àÍäáéÌï "
        .Col = 30
        .Text = "àÍäáäJénåé"
        .Col = 31
        .Text = "å_ñÒéÛóùîNåé"
        .Col = 32
        .Text = "ãGêﬂÅEäwê∂"
        .Col = 33
        .Text = "ç–äQÅEñ∆èú"
        .Col = 34
        .Text = "ÉiÉJèZèä"
        .Col = 35
        .Text = "ã«ÉRÅ[Éh"
        .Col = 36
        .Text = "Ç®ãqólî‘çÜ"
        .Col = 37
        .Text = "ÉJÉiéÅñº"
        .Col = 38
        .Text = "êøãÅîNåé"
End With
End Sub

Private Sub Picture13_Click()
Unload Me
End Sub
Sub id_chk()

Dim iiii As Long
Dim o_id As String

With gridnhk
For iiii = 1 To .Rows
    .Row = iiii
    .Col = 8
    o_id = Trim(.Text)
    
    
    
Next
End With

End Sub



Private Sub Picture2_Click()
gridkekka.Rows = 1
fhikakukekka.Visible = False
End Sub

Private Sub Picture3_Click()
gridkekka2.Rows = 1
gridkekka3.Rows = 1
fhikakukekka2.Visible = False

End Sub

Private Sub Picture5_Click()


If fhikakukekka2.Visible = True Then
    gridkekka2.Rows = 1
    fhikakukekka2.Visible = False
End If
If fhikakukekka.Visible = True Then
    gridkekka.Rows = 1
    fhikakukekka.Visible = False
End If

If optchk(0).Value = True Then
    nhk_csv_fdd_check3
ElseIf optchk(1).Value = True Then
    nhk_csv_fdd_check4
ElseIf optchk(2).Value = True Then
    nhk_csv_fdd_check

ElseIf optchk(3).Value = True Then
    nhk_csv_fdd_check2
Else
    MsgBox "É`ÉFÉbÉNÇ∑ÇÈéÌóﬁÇëIÇÒÇ≈Ç©ÇÁÅAé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
    Exit Sub
End If

End Sub

Private Sub Picture7_Click()


End Sub

Private Sub Picture9_Click()
    Dim udtBrowseInfo As BROWSEINFO
    Dim lngIDList     As Long
    Dim strPath       As String
    
    '[ÉtÉHÉãÉ_ÇÃéQè∆]É_ÉCÉAÉçÉOèÓïÒÇê›íË
    With udtBrowseInfo
        .hWndOwner = Me.hwnd
        .pidlRoot = CSIDL_NETWORK
        .pszDisplayName = String(MAX_PATH2, vbNullChar)
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    '[ÉtÉHÉãÉ_ÇÃéQè∆]É_ÉCÉAÉçÉOÇï\é¶
    lngIDList = SHBrowseForFolder(udtBrowseInfo)
    If lngIDList <> 0 Then
        strPath = String(MAX_PATH2, vbNullChar)
        'ITEMIDLISTÇÉpÉXÇ…ïœä∑
        SHGetPathFromIDList lngIDList, strPath
        CoTaskMemFree lngIDList
            strPath = Left(strPath, _
                           InStr(strPath, _
                                 vbNullChar) - 1)
        'åãâ ÇÉâÉxÉãÇ…ï\é¶
        txtdatapath.Text = strPath
    End If

End Sub

