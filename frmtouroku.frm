VERSION 5.00
Object = "{E2D000D1-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "Text60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmtouroku 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "0"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15330
   StartUpPosition =   1  'µ∞≈∞ Ã´∞—ÇÃíÜâõ
   Begin TabDlg.SSTab SSTab1 
      Height          =   1935
      Left            =   90
      TabIndex        =   147
      Top             =   4650
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   3413
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "ê›íuêÊèÓïÒ"
      TabPicture(0)   =   "frmtouroku.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame11(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "NHKä÷òA/å_ñÒçXêVèÛãµ"
      TabPicture(1)   =   "frmtouroku.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame11(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "óXëóêÊèÓïÒ"
      TabPicture(2)   =   "frmtouroku.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame11(2)"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "å_ñÒçXêVèÛãµ"
         Height          =   1365
         Left            =   -67440
         TabIndex        =   183
         Top             =   420
         Width           =   7365
         Begin VB.TextBox txtsaishuukoushinbi 
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   184
            Top             =   270
            Width           =   1515
         End
         Begin VB.Label Label58 
            Caption         =   "ç≈èIå_ñÒçXêVì˙ÅF"
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
            Left            =   420
            TabIndex        =   185
            Top             =   330
            Width           =   1935
         End
      End
      Begin VB.Frame Frame11 
         Height          =   1365
         Index           =   2
         Left            =   -74880
         TabIndex        =   150
         Top             =   390
         Width           =   14865
         Begin VB.TextBox txtatena_y 
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
            Left            =   4500
            MaxLength       =   60
            TabIndex        =   186
            Top             =   240
            Width           =   4935
         End
         Begin VB.TextBox txtyuubin2_y 
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
            Left            =   2595
            MaxLength       =   4
            TabIndex        =   179
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtyuubin1_y 
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
            Left            =   1740
            MaxLength       =   3
            TabIndex        =   178
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtjuusho2_y 
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
            Left            =   6750
            TabIndex        =   177
            Top             =   705
            Width           =   6465
         End
         Begin VB.TextBox txtjuusho1_y 
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
            Left            =   1710
            TabIndex        =   176
            Top             =   705
            Width           =   4995
         End
         Begin VB.Label Label59 
            Caption         =   "à∂ñºÅF"
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
            Left            =   3780
            TabIndex        =   187
            Top             =   315
            Width           =   795
         End
         Begin VB.Label Label57 
            Caption         =   "Å|"
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
            Left            =   2295
            TabIndex        =   182
            Top             =   285
            Width           =   255
         End
         Begin VB.Label Label32 
            Caption         =   "óXï÷î‘çÜÅF"
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
            Left            =   240
            TabIndex        =   181
            Top             =   285
            Width           =   1635
         End
         Begin VB.Label Label24 
            Caption         =   "óXëóêÊèZèäÅF"
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
            Left            =   210
            TabIndex        =   180
            Top             =   780
            Width           =   1635
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "NHKä÷òA"
         Height          =   1365
         Index           =   1
         Left            =   -74880
         TabIndex        =   149
         Top             =   420
         Width           =   7365
         Begin VB.TextBox txtnhkname 
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
            Left            =   3660
            TabIndex        =   173
            Top             =   270
            Width           =   2745
         End
         Begin VB.TextBox txtnhkfuri 
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
            Left            =   3660
            TabIndex        =   172
            Top             =   735
            Width           =   2715
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Å´"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6750
            TabIndex        =   170
            ToolTipText     =   "åªç›ÇÃå_ñÒé“ñºÇâ∫ãLÇ…ÉRÉsÅ[ÇµÇ‹Ç∑ÅB"
            Top             =   360
            Width           =   375
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
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   169
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label Label51 
            Caption         =   "ñºëOÅF"
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
            Left            =   2820
            TabIndex        =   175
            Top             =   330
            Width           =   1215
         End
         Begin VB.Label Label50 
            Caption         =   "Ãÿ∂ﬁ≈ÅF"
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
            Left            =   2820
            TabIndex        =   174
            Top             =   795
            Width           =   1215
         End
         Begin VB.Label Label49 
            Caption         =   "î‘çÜÅF"
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
            Left            =   240
            TabIndex        =   171
            Top             =   330
            Width           =   1215
         End
      End
      Begin VB.Frame Frame11 
         Height          =   1455
         Index           =   0
         Left            =   90
         TabIndex        =   148
         Top             =   360
         Width           =   14865
         Begin VB.TextBox txthoshoukin 
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
            Left            =   12990
            MaxLength       =   2
            TabIndex        =   167
            Top             =   930
            Width           =   525
         End
         Begin VB.TextBox combcellid 
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
            Left            =   10920
            MaxLength       =   2
            TabIndex        =   165
            Top             =   930
            Width           =   525
         End
         Begin VB.TextBox txttvsuu 
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
            Left            =   12990
            MaxLength       =   2
            TabIndex        =   163
            Top             =   570
            Width           =   525
         End
         Begin VB.TextBox txtsecchitel 
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
            Left            =   10920
            MaxLength       =   12
            TabIndex        =   160
            Top             =   150
            Width           =   1965
         End
         Begin VB.TextBox txtsetaisuu 
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
            Left            =   10920
            MaxLength       =   2
            TabIndex        =   159
            Top             =   555
            Width           =   525
         End
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
            IMEMode         =   4  'ëSäpÇ–ÇÁÇ™Ç»
            Left            =   810
            MaxLength       =   50
            TabIndex        =   157
            Top             =   990
            Width           =   8595
         End
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
            IMEMode         =   6  'îºäp∂¿∂≈
            Left            =   5550
            MaxLength       =   45
            TabIndex        =   153
            Top             =   180
            Width           =   3855
         End
         Begin VB.TextBox txtsecchiname 
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
            Left            =   810
            MaxLength       =   60
            TabIndex        =   152
            Top             =   165
            Width           =   3855
         End
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
            IMEMode         =   5  'ëSäp∂¿∂≈
            Left            =   810
            MaxLength       =   50
            TabIndex        =   151
            Top             =   600
            Width           =   8595
         End
         Begin VB.Label Label22 
            Caption         =   "ê}ñ î‘çÜÅF"
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
            Left            =   11670
            TabIndex        =   168
            Top             =   1005
            Width           =   1155
         End
         Begin VB.Label Label55 
            Caption         =   "ÉmÅ[Éhî‘çÜÅF"
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
            Left            =   9480
            TabIndex        =   166
            Top             =   1005
            Width           =   1275
         End
         Begin VB.Label Label17 
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   12090
            TabIndex        =   164
            Top             =   645
            Width           =   795
         End
         Begin VB.Label Label21 
            Caption         =   "òAóçêÊÅF"
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
            Left            =   9870
            TabIndex        =   162
            Top             =   210
            Width           =   885
         End
         Begin VB.Label Label16 
            Caption         =   "ê¢ë—êîÅF"
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
            Left            =   9780
            TabIndex        =   161
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "ÉÅÉÇÅF"
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
            Left            =   90
            TabIndex        =   158
            Top             =   1020
            Width           =   795
         End
         Begin VB.Label Label12 
            Caption         =   "Ãÿ∂ﬁ≈ÅF"
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
            Left            =   4830
            TabIndex        =   156
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label13 
            Caption         =   "éÅñºÅF"
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
            Left            =   90
            TabIndex        =   155
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label14 
            Caption         =   "èZèäÅF"
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
            Left            =   90
            TabIndex        =   154
            Top             =   645
            Width           =   795
         End
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "å_ñÒé“èÓïÒ"
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
      Height          =   4515
      Left            =   60
      TabIndex        =   57
      Top             =   30
      Width           =   15105
      Begin VB.TextBox txtmaekubun 
         Alignment       =   2  'íÜâõëµÇ¶
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
         Left            =   8400
         TabIndex        =   144
         Top             =   570
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cl"
         Height          =   315
         Left            =   10890
         TabIndex        =   143
         Top             =   1050
         Width           =   285
      End
      Begin VB.ComboBox comjuutaku2 
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
         Left            =   9540
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   142
         Top             =   1030
         Width           =   1305
      End
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
         Left            =   8190
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   136
         Top             =   1030
         Width           =   1305
      End
      Begin VB.TextBox txtkeiyakuid 
         Alignment       =   2  'íÜâõëµÇ¶
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
         Left            =   9480
         TabIndex        =   135
         Top             =   210
         Width           =   1080
      End
      Begin VB.TextBox txtjuusho1 
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
         Left            =   1230
         TabIndex        =   134
         Top             =   1410
         Width           =   4995
      End
      Begin VB.TextBox txtjuusho2 
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
         Left            =   6240
         TabIndex        =   133
         Top             =   1410
         Width           =   4935
      End
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
         IMEMode         =   4  'ëSäpÇ–ÇÁÇ™Ç»
         Left            =   9450
         TabIndex        =   132
         Top             =   620
         Width           =   4935
      End
      Begin VB.Frame Frame12 
         Caption         =   "CATVä÷òA(è„ÅFäJénì˙Å@â∫ÅFâñÒì˙)"
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
         Height          =   1250
         Left            =   120
         TabIndex        =   78
         Top             =   3150
         Width           =   5200
         Begin VB.TextBox txtkaiyaku 
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
            Left            =   3750
            MaxLength       =   10
            TabIndex        =   139
            Top             =   750
            Width           =   1350
         End
         Begin VB.TextBox txtkanyuu 
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
            Left            =   3750
            MaxLength       =   10
            TabIndex        =   138
            Top             =   300
            Width           =   1350
         End
         Begin VB.TextBox txtwaribiki 
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
            Left            =   2580
            MaxLength       =   7
            TabIndex        =   81
            Top             =   300
            Width           =   915
         End
         Begin VB.TextBox txtriyuu 
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
            Left            =   810
            MaxLength       =   25
            TabIndex        =   80
            Top             =   750
            Width           =   2685
         End
         Begin VB.TextBox txtkoujihi 
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
            Left            =   930
            MaxLength       =   7
            TabIndex        =   79
            Top             =   300
            Width           =   945
         End
         Begin VB.Label Label33 
            Caption         =   "çHéñîÔÅF"
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
            Left            =   60
            TabIndex        =   84
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label38 
            Caption         =   "äÑà¯ÅF"
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
            TabIndex        =   83
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label39 
            Caption         =   "óùóRÅF"
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
            Left            =   90
            TabIndex        =   82
            Top             =   780
            Width           =   1215
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "ì¸ã‡ä÷òA"
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
         Height          =   1250
         Left            =   10770
         TabIndex        =   66
         Top             =   3150
         Width           =   4215
         Begin VB.TextBox txtnyuukinbi 
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
            Left            =   930
            MaxLength       =   10
            TabIndex        =   71
            Top             =   300
            Width           =   1365
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
            Left            =   1020
            MaxLength       =   4
            TabIndex        =   70
            Top             =   780
            Width           =   615
         End
         Begin VB.TextBox txtnyuukingaku 
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
            Left            =   3210
            MaxLength       =   7
            TabIndex        =   69
            Top             =   300
            Width           =   855
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
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   68
            Top             =   780
            Width           =   345
         End
         Begin VB.TextBox txtzangaku 
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
            Left            =   3210
            MaxLength       =   7
            TabIndex        =   67
            Top             =   750
            Width           =   855
         End
         Begin VB.Label Label40 
            Caption         =   "ä÷òAì˙ÅF"
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
            Left            =   60
            TabIndex        =   77
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label41 
            Caption         =   "ì¸ã‡äzÅF"
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
            Left            =   2310
            TabIndex        =   76
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label42 
            Caption         =   "ì¸ã‡ì˙ÅF"
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
            Left            =   30
            TabIndex        =   75
            Top             =   360
            Width           =   1215
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
            Left            =   1680
            TabIndex        =   74
            Top             =   900
            Width           =   285
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
            Left            =   2310
            TabIndex        =   73
            Top             =   900
            Width           =   195
         End
         Begin VB.Label Label45 
            Caption         =   "écäzÅF"
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
            Left            =   2550
            TabIndex        =   72
            Top             =   810
            Width           =   1215
         End
      End
      Begin VB.ComboBox combcellno 
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
         Left            =   6330
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   1035
         Width           =   675
      End
      Begin VB.TextBox txtlinkid 
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
         Left            =   14100
         MaxLength       =   5
         TabIndex        =   123
         Top             =   240
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Caption         =   "ÉCÉìÉ^Å[ÉlÉbÉgä÷òA(è„ÅFäJénì˙Å@â∫ÅFâñÒì˙)"
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
         Height          =   1250
         Left            =   5400
         TabIndex        =   110
         Top             =   3150
         Width           =   5200
         Begin VB.TextBox txtkaiyaku2 
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
            Left            =   3750
            MaxLength       =   10
            TabIndex        =   141
            Top             =   720
            Width           =   1350
         End
         Begin VB.TextBox txtkanyuu2 
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
            Left            =   3750
            MaxLength       =   10
            TabIndex        =   140
            Top             =   270
            Width           =   1350
         End
         Begin VB.TextBox txtkoujihi2 
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
            Left            =   930
            MaxLength       =   7
            TabIndex        =   113
            Top             =   270
            Width           =   945
         End
         Begin VB.TextBox txtriyuu2 
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
            Left            =   810
            MaxLength       =   25
            TabIndex        =   112
            Top             =   720
            Width           =   2745
         End
         Begin VB.TextBox txtwaribiki2 
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
            Left            =   2670
            MaxLength       =   7
            TabIndex        =   111
            Top             =   270
            Width           =   915
         End
         Begin VB.Label Label54 
            Caption         =   "óùóRÅF"
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
            Left            =   90
            TabIndex        =   116
            Top             =   750
            Width           =   1215
         End
         Begin VB.Label Label53 
            Caption         =   "äÑà¯ÅF"
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
            Left            =   2010
            TabIndex        =   115
            Top             =   330
            Width           =   1215
         End
         Begin VB.Label Label52 
            Caption         =   "çHéñîÔÅF"
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
            Left            =   60
            TabIndex        =   114
            Top             =   330
            Width           =   1215
         End
      End
      Begin VB.TextBox txtfax 
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
         Left            =   9540
         MaxLength       =   12
         TabIndex        =   93
         Top             =   1845
         Width           =   1600
      End
      Begin VB.TextBox txtkeitai 
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
         Left            =   6840
         MaxLength       =   13
         TabIndex        =   92
         Top             =   1845
         Width           =   1665
      End
      Begin VB.TextBox txttel2 
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
         Left            =   4125
         MaxLength       =   12
         TabIndex        =   91
         Top             =   1845
         Width           =   1600
      End
      Begin VB.TextBox txttel1 
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
         Left            =   1245
         MaxLength       =   12
         TabIndex        =   90
         Top             =   1845
         Width           =   1600
      End
      Begin VB.TextBox txtkeiyakushafuri 
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
         Left            =   2130
         MaxLength       =   20
         TabIndex        =   2
         Top             =   240
         Width           =   6315
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
         Height          =   855
         IMEMode         =   4  'ëSäpÇ–ÇÁÇ™Ç»
         Left            =   1230
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   89
         Top             =   2250
         Width           =   13785
      End
      Begin VB.TextBox txtkeiyakubi 
         Alignment       =   2  'íÜâõëµÇ¶
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
         Left            =   11520
         TabIndex        =   88
         Top             =   210
         Width           =   1620
      End
      Begin VB.TextBox txtyuubin1 
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
         Left            =   1260
         MaxLength       =   3
         TabIndex        =   87
         Top             =   1035
         Width           =   495
      End
      Begin VB.TextBox txtyuubin2 
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
         Left            =   2115
         MaxLength       =   4
         TabIndex        =   86
         Top             =   1035
         Width           =   615
      End
      Begin VB.ComboBox comchiiki 
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
         Left            =   3945
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   85
         Top             =   1035
         Width           =   2325
      End
      Begin VB.Frame Frame14 
         Caption         =   "èÛãµ"
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
         Height          =   1215
         Left            =   11250
         TabIndex        =   58
         Top             =   960
         Width           =   3765
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
            Height          =   315
            IMEMode         =   2  'µÃ
            Left            =   3090
            MaxLength       =   2
            TabIndex        =   63
            Top             =   240
            Width           =   345
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
            Height          =   315
            IMEMode         =   2  'µÃ
            Left            =   2190
            MaxLength       =   4
            TabIndex        =   62
            Top             =   240
            Width           =   615
         End
         Begin VB.CheckBox chkq 
            Caption         =   "ÅuQÅvÇégópÇ∑ÇÈÅB"
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
            Height          =   345
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   2055
         End
         Begin VB.CheckBox chkkyuushi 
            Caption         =   "åªç›ÅAéÊà¯Ç™Ç†ÇËÇ‹ÇπÇÒÅB"
            Enabled         =   0   'False
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
            Left            =   120
            TabIndex        =   60
            Top             =   540
            Value           =   1  '¡™Ø∏
            Width           =   3075
         End
         Begin VB.CheckBox chkblack 
            Caption         =   "ÉuÉâÉbÉNÉäÉXÉgëŒè€é“Ç≈Ç∑ÅB"
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
            Left            =   120
            TabIndex        =   59
            Top             =   840
            Width           =   3075
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
            Left            =   3480
            TabIndex        =   65
            Top             =   330
            Width           =   195
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
            Left            =   2850
            TabIndex        =   64
            Top             =   330
            Width           =   285
         End
      End
      Begin imText6Ctl.imText txtkeiyakushamei2 
         Height          =   375
         Left            =   5850
         TabIndex        =   1
         Top             =   620
         Width           =   2595
         _Version        =   65536
         _ExtentX        =   4577
         _ExtentY        =   661
         Caption         =   "frmtouroku.frx":0054
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmtouroku.frx":00C2
         Key             =   "frmtouroku.frx":00E0
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   1
         AlignHorizontal =   0
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   0
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   -1
         HighlightText   =   0
         IMEMode         =   4
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin imText6Ctl.imText txtkeiyakushamei 
         Height          =   375
         Left            =   2130
         TabIndex        =   0
         Top             =   620
         Width           =   3195
         _Version        =   65536
         _ExtentX        =   5636
         _ExtentY        =   661
         Caption         =   "frmtouroku.frx":0124
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmtouroku.frx":0192
         Key             =   "frmtouroku.frx":01B0
         BackColor       =   -2147483643
         EditMode        =   0
         ForeColor       =   -2147483640
         ReadOnly        =   0
         ShowContextMenu =   -1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MarginBottom    =   1
         Enabled         =   -1
         MousePointer    =   0
         Appearance      =   1
         BorderStyle     =   1
         AlignHorizontal =   0
         AlignVertical   =   0
         MultiLine       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         AllowSpace      =   -1
         Format          =   ""
         FormatMode      =   1
         AutoConvert     =   -1
         ErrorBeep       =   0
         MaxLength       =   30
         LengthAsByte    =   0
         Text            =   ""
         Furigana        =   -1
         HighlightText   =   0
         IMEMode         =   4
         IMEStatus       =   0
         DropWndWidth    =   0
         DropWndHeight   =   0
         ScrollBarMode   =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
      End
      Begin VB.Label Label6 
         Caption         =   "èZèäÇPÅF"
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
         Left            =   180
         TabIndex        =   137
         Top             =   1500
         Width           =   1635
      End
      Begin VB.Label Label56 
         Caption         =   "ä«óùÇhÇcÅF"
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
         Left            =   13230
         TabIndex        =   124
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label46 
         Caption         =   "âºèZèäÅF"
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
         Left            =   8610
         TabIndex        =   109
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "FAXÅF"
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
         Left            =   8715
         TabIndex        =   108
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "ågë—ÅF"
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
         Left            =   5940
         TabIndex        =   107
         Top             =   1920
         Width           =   690
      End
      Begin VB.Label Label9 
         Caption         =   "ìdòbÇQÅF"
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
         Left            =   3060
         TabIndex        =   106
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "ìdòbÇPÅF"
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
         Left            =   180
         TabIndex        =   105
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "óXï÷î‘çÜÅF"
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
         Left            =   180
         TabIndex        =   104
         Top             =   1080
         Width           =   1635
      End
      Begin VB.Label Label19 
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
         Left            =   180
         TabIndex        =   101
         Top             =   2340
         Width           =   1635
      End
      Begin VB.Label Label7 
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
         Left            =   8640
         TabIndex        =   100
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label20 
         Caption         =   "ì¸óÕì˙ÅF"
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
         Left            =   10650
         TabIndex        =   99
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "ínàÊãÊï™ÅF"
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
         Left            =   2760
         TabIndex        =   98
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "èZëÓãÊï™ÅF"
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
         Left            =   7035
         TabIndex        =   97
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Å|"
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
         Left            =   1815
         TabIndex        =   96
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label36 
         Caption         =   "(ê©)"
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
         Left            =   1650
         TabIndex        =   95
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label37 
         Caption         =   "(ñº)"
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
         Left            =   5370
         TabIndex        =   94
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "å_ñÒé“Ãÿ∂ﬁ≈ÅF"
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
         Left            =   180
         TabIndex        =   102
         Top             =   360
         Width           =   1755
      End
      Begin VB.Label Label2 
         Caption         =   "å_ñÒé“éÅñºÅF"
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
         Left            =   180
         TabIndex        =   103
         Top             =   720
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   14520
      MouseIcon       =   "frmtouroku.frx":01F4
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmtouroku.frx":0AF3
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   121
      Top             =   7950
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   14520
      MouseIcon       =   "frmtouroku.frx":1935
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmtouroku.frx":2234
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   119
      Top             =   10050
      Width           =   480
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   14490
      MouseIcon       =   "frmtouroku.frx":2AFE
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmtouroku.frx":33FD
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   117
      Top             =   9060
      Width           =   480
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
      Height          =   1665
      Left            =   60
      TabIndex        =   52
      Top             =   6720
      Width           =   2475
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
         Left            =   90
         TabIndex        =   55
         Top             =   300
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
         Left            =   90
         TabIndex        =   54
         Top             =   735
         Width           =   2300
      End
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
         Left            =   90
         TabIndex        =   53
         Top             =   1170
         Width           =   2300
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
      Height          =   1665
      Left            =   60
      TabIndex        =   48
      Top             =   8520
      Width           =   2475
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
         TabIndex        =   51
         Top             =   1140
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
         TabIndex        =   50
         Top             =   705
         Width           =   2300
      End
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
         TabIndex        =   49
         Top             =   270
         Width           =   2300
      End
   End
   Begin VB.TextBox txtht4 
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
      TabIndex        =   47
      Top             =   10560
      Width           =   2300
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
      Left            =   2730
      TabIndex        =   26
      Top             =   6720
      Width           =   4725
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
         Left            =   90
         MaxLength       =   11
         TabIndex        =   46
         Top             =   300
         Width           =   2300
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
         Left            =   90
         MaxLength       =   11
         TabIndex        =   45
         Top             =   696
         Width           =   2300
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
         Left            =   90
         MaxLength       =   11
         TabIndex        =   44
         Top             =   1092
         Width           =   2300
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
         Left            =   90
         MaxLength       =   11
         TabIndex        =   43
         Top             =   1488
         Width           =   2300
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
         Left            =   90
         MaxLength       =   11
         TabIndex        =   42
         Top             =   1884
         Width           =   2300
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
         Left            =   90
         MaxLength       =   11
         TabIndex        =   41
         Top             =   2280
         Width           =   2300
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
         Left            =   90
         MaxLength       =   11
         TabIndex        =   40
         Top             =   2676
         Width           =   2300
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
         Left            =   90
         MaxLength       =   11
         TabIndex        =   39
         Top             =   3072
         Width           =   2300
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
         Left            =   90
         MaxLength       =   11
         TabIndex        =   38
         Top             =   3468
         Width           =   2300
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
         Left            =   90
         MaxLength       =   11
         TabIndex        =   37
         Top             =   3870
         Width           =   2300
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
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   36
         Top             =   300
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
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   35
         Top             =   696
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
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   34
         Top             =   1092
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
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   33
         Top             =   1488
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
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   32
         Top             =   1884
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
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   31
         Top             =   2280
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
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   30
         Top             =   2676
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
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   29
         Top             =   3072
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
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   28
         Top             =   3468
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
         Index           =   9
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   27
         Top             =   3870
         Width           =   1785
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "ä«óùå˚ç¿èÓïÒ"
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
      Height          =   705
      Left            =   7710
      TabIndex        =   22
      Top             =   10290
      Width           =   6375
      Begin VB.TextBox txtkanrikouza 
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
         Left            =   1290
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   4950
      End
      Begin VB.Label Label23 
         Caption         =   "ã‡óZã@ä÷ÅF"
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
         Left            =   120
         TabIndex        =   24
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "ãåå˚ç¿èÓïÒÅiå˚ç¿à¯óéÇ©ÇÁåªã‡Ç÷ïœçXéûÅj"
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
      Height          =   1065
      Left            =   7710
      TabIndex        =   17
      Top             =   9180
      Width           =   6375
      Begin VB.TextBox txtkyukouza2 
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
         Left            =   1320
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   615
         Width           =   4900
      End
      Begin VB.TextBox txtkyukouza1 
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
         Left            =   1320
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   4900
      End
      Begin VB.Label Label27 
         Caption         =   "ã‡óZã@ä÷ÅF"
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
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "ÇªÇÃëºÅF"
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
         Left            =   180
         TabIndex        =   20
         Top             =   660
         Width           =   1155
      End
   End
   Begin VB.Frame fseikyuu 
      Caption         =   "êøãÅï˚ñ@"
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
      Height          =   765
      Left            =   7710
      TabIndex        =   14
      Top             =   6630
      Width           =   2775
      Begin VB.TextBox txtmae_shiharai 
         Alignment       =   2  'íÜâõëµÇ¶
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
         Left            =   1320
         TabIndex        =   145
         Top             =   0
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.OptionButton optseikyuuhouhou 
         Caption         =   "êøãÅèë"
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
         Left            =   1590
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   300
         Width           =   1005
      End
      Begin VB.OptionButton optseikyuuhouhou 
         Caption         =   "å˚ç¿êUë÷"
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
         Left            =   210
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   300
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame fkouza 
      Caption         =   "å˚ç¿èÓïÒ"
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
      Height          =   1575
      Left            =   7710
      TabIndex        =   3
      Top             =   7500
      Width           =   6375
      Begin VB.TextBox txtmae_ginkou 
         Alignment       =   2  'íÜâõëµÇ¶
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
         Left            =   6000
         TabIndex        =   146
         Top             =   360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Å´"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5850
         TabIndex        =   25
         ToolTipText     =   "åªç›ÇÃå˚ç¿ÉfÅ[É^Çâ∫ãLÇ…ÉRÉsÅ[ÇµÇ‹Ç∑ÅB"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtkouzameigi 
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
         Left            =   1590
         MaxLength       =   30
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1110
         Width           =   3780
      End
      Begin VB.TextBox txtkouzabangou 
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
         Left            =   3630
         MaxLength       =   14
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   660
         Width           =   1980
      End
      Begin VB.ComboBox cmbbank 
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
         Left            =   1290
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   270
         Width           =   4575
      End
      Begin VB.Frame Frame9 
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
         Height          =   495
         Left            =   930
         TabIndex        =   4
         Top             =   570
         Width           =   1965
         Begin VB.OptionButton optkouzashurui 
            Caption         =   "ïÅí "
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   150
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optkouzashurui 
            Caption         =   "ìñç¿"
            Height          =   315
            Index           =   1
            Left            =   990
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   150
            Width           =   795
         End
      End
      Begin VB.Label Label28 
         Caption         =   "å˚ç¿ñºã`ÅF"
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
         Left            =   150
         TabIndex        =   13
         Top             =   1185
         Width           =   1155
      End
      Begin VB.Label Label29 
         Caption         =   "î‘çÜÅF"
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
         Left            =   2940
         TabIndex        =   12
         Top             =   705
         Width           =   645
      End
      Begin VB.Label Label30 
         Caption         =   "éÌóﬁÅF"
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
         Left            =   180
         TabIndex        =   11
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label31 
         Caption         =   "ã‡óZã@ä÷ÅF"
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
         Left            =   150
         TabIndex        =   10
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame Frame15 
      Caption         =   "ñ¢à¯óéâÒêî"
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
      Height          =   765
      Left            =   10590
      TabIndex        =   126
      Top             =   6630
      Width           =   2235
      Begin VB.TextBox txtfkaisuu 
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
         Left            =   990
         MaxLength       =   1
         TabIndex        =   127
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Label35 
         Caption         =   "âÒ"
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
         Left            =   1770
         TabIndex        =   129
         Top             =   420
         Width           =   345
      End
      Begin VB.Label Label34 
         Caption         =   "âÒêîÅF"
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
         TabIndex        =   128
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame16 
      Caption         =   "éÊÇËàµÇ¢èÛãµ"
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
      Height          =   765
      Left            =   12900
      TabIndex        =   130
      Top             =   6630
      Width           =   2235
      Begin VB.CheckBox chkend 
         Caption         =   "åüçıéûîÒï\é¶"
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
         Left            =   150
         TabIndex        =   131
         Top             =   330
         Width           =   1905
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÉNÉäÉA"
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
      Height          =   285
      Index           =   9
      Left            =   14430
      TabIndex        =   122
      Top             =   8460
      Width           =   675
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
      Left            =   14550
      TabIndex        =   120
      Top             =   10560
      Width           =   495
   End
   Begin VB.Label cmdtouroku 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   14520
      TabIndex        =   118
      Top             =   9570
      Width           =   90
   End
   Begin VB.Label Label25 
      Caption         =   "Œ∞—¿∞–≈ŸÇÃï€èÿã‡"
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
      Left            =   90
      TabIndex        =   56
      Top             =   10290
      Width           =   1785
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   3465
      Left            =   14310
      Top             =   7500
      Width           =   900
   End
End
Attribute VB_Name = "frmtouroku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkblack_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        txtsecchifuri.SetFocus
End If

End Sub

Private Sub cmbbank_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If optkouzashurui(0).Value = True Then
        optkouzashurui(0).SetFocus
    ElseIf optkouzashurui(1).Value = True Then
        optkouzashurui(1).SetFocus
    End If
End If

End Sub




Private Sub combcellid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txthoshoukin.SetFocus
End If

End Sub

Private Sub combcellno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    comjuutaku.SetFocus
End If

End Sub

Private Sub comchiiki_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    combcellno.SetFocus
End If

End Sub

Private Sub comjuutaku_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtjuusho1.SetFocus
End If

End Sub

Private Sub Command1_Click()
If cmbbank.ListIndex = -1 Then
    MsgBox "ã‡óZã@ä÷ñºÇ™ëIëÇ≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB"
    Exit Sub
End If
 txtkyukouza1.Text = cmbbank.Text

If optkouzashurui(0).Value = 1 Then
    txtkyukouza2.Text = "ìñç¿ " & txtkouzabangou.Text & Space(1) & txtkouzameigi.Text
Else
    txtkyukouza2.Text = "ïÅí  " & txtkouzabangou.Text & Space(1) & txtkouzameigi.Text
End If
End Sub

Private Sub Command2_Click()

With frmtouroku
    .txtnhkname.Text = Trim(.txtkeiyakushamei.Text) & Space(1) & Trim(.txtkeiyakushamei2.Text)
    .txtnhkfuri.Text = Trim(.txtkeiyakushafuri.Text)
End With
End Sub

Private Sub Command3_Click()
 comjuutaku2.ListIndex = -1
End Sub

Private Sub Form_Load()

Dim cidcid As Integer

With frmtouroku
    
    .comchiiki.AddItem "Ç`Å@èüè¿ÇPÅ`ÇPÇSãÊ"
    .comchiiki.AddItem "ÇaÅ@èjÅ@ÇPÅ`ÇPÇOãÊ"
    .comchiiki.AddItem "ÇbÅ@ìåâ_ÇPÅ`ÇPÇTãÊ"
    .comchiiki.AddItem "ÇcÅ@ïHéRÇPÅ`ÇPÇOãÊ"

    .comjuutaku.AddItem "àÍî â∆íÎ"
    .comjuutaku.AddItem "àÍî éññ±èä"
    .comjuutaku.AddItem "åˆã§"
    .comjuutaku.AddItem "åˆã§ÇP"
    .comjuutaku.AddItem "èWçáèZëÓ"
    .comjuutaku.AddItem "î_ÅEéññ±èä"
    .comjuutaku.AddItem "óVóóâÄ"
    .comjuutaku.AddItem "ÇªÇÃëº"
    
    .comjuutaku.ListIndex = 0
    
     .comjuutaku2.AddItem "àÍî â∆íÎ"
    .comjuutaku2.AddItem "àÍî éññ±èä"
    .comjuutaku2.AddItem "åˆã§"
    .comjuutaku2.AddItem "åˆã§ÇP"
    .comjuutaku2.AddItem "èWçáèZëÓ"
    .comjuutaku2.AddItem "î_ÅEéññ±èä"
    .comjuutaku2.AddItem "óVóóâÄ"
    .comjuutaku2.AddItem "ÇªÇÃëº"
    
    .comjuutaku2.ListIndex = -1
    
    
    For cidcid = 1 To 15
        .combcellno.AddItem Format(cidcid, "00")
    Next cidcid
    
    
   
    'cell_ad_set
    
End With

sbbankset

End Sub

Private Sub Form_Unload(Cancel As Integer)
furiganahenkou2 = 0

End Sub



Private Sub optkouzashurui_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    txtkouzabangou.SetFocus
End If

End Sub

Private Sub optseikyuuhouhou_Click(Index As Integer)
    If optseikyuuhouhou(0).Value = True Then
        fkouza.Enabled = True
    Else
        fkouza.Enabled = False
    End If
End Sub

Private Sub optseikyuuhouhou_KeyPress(Index As Integer, KeyAscii As Integer)
    
If KeyAscii = 13 Then
        cmbbank.SetFocus
End If

End Sub

Private Sub Picture11_Click()
frmtourokuclear

End Sub

Private Sub Picture13_Click()
Unload frmtouroku
Set frmtouroku = Nothing

End Sub

Private Sub Picture7_Click()
Dim kojindata(42) As String, newid As String, newkojindata(42) As String
Dim nkdi As Integer, newfkaisuu As String
    Dim sknen As Long, sktsuki As Long, chkkingaku As Long, newlinkid As String
    Dim neweend As String, n_catv_kaiyakubi As String, n_internet_kaiyakubi As String
    Dim s_kubun As String


    Dim s_koujikannryoubi As String
    Dim s_koukubangou As String
    Dim s_zumen As String
    Dim s_bikou As String
    
    Dim s_listbangou As String
    
    Dim s_clbangou As String
    Dim s_vonuadress As String
    
    Dim s_donuadress As String
    
    Dim s_open_day As String
    Dim s_close_day As String
    Dim s_keiyakusha As String
    
    Dim s_bspass As String
    
   Dim s_saishuuhakkoubi As String
   
    Dim s_y_yuubin As String
    Dim s_y_juusho1 As String
    Dim s_y_juusho2 As String
    Dim s_y_atena As String
    
     Dim s_saishuukoushinbi As String
    
    
    If kiidou_mode <> 0 Then
        MsgBox "Ç±ÇÃÉÇÅ[ÉhÇ≈ÇÕégópÇ≈Ç´Ç‹ÇπÇÒÅB"
        Exit Sub
    End If

If frmmain.lblkanri2.Visible = True Then
      MsgBox "ÉTÉuÉÇÅ[ÉhÇ≈ÇÕçXêVÇ≈Ç´Ç‹ÇπÇÒÅB"
        Exit Sub
End If
    
For i = 0 To 42
  
        kojindata(i) = Space(1)
   
    newkojindata(i) = Space(1)
Next i
With frmtouroku
   
        s_koujikannryoubi = Space(1)
    
        s_koukubangou = Space(1)
   
        s_zumen = Space(1)
        
        s_bikou = Space(1)
        
        s_listbangou = Space(1)
   
        s_clbangou = Space(1)
        s_vonuadress = Space(1)
        
        s_donuadress = Space(1)
        
        s_open_day = Space(1)
        
        s_close_day = Space(1)
        
        s_keiyakusha = Space(1)
        
        s_bspass = Space(1)
        
        s_saishuuhakkoubi = Space(1)
          s_saishuukoushinbi = Space(1)
    
         s_y_yuubin = Space(1)
         s_y_juusho1 = Space(1)
         s_y_juusho2 = Space(1)
         s_y_atena = Space(1)
    
    If Trim(.txtyuubin1_y.Text) <> "" Then
        If Trim(.txtyuubin2_y.Text) <> "" Then
            s_y_yuubin = Trim(.txtyuubin1_y.Text) & "-" & Trim(.txtyuubin2_y.Text)
        End If
    End If
    
    If Trim(.txtjuusho1_y.Text) <> "" Then
       If Trim(.txtjuusho2_y.Text) <> "" Then
            s_y_juusho1 = Trim(.txtjuusho1_y.Text)
            s_y_juusho2 = Trim(.txtjuusho2_y.Text)
        End If
    End If
    
    If Trim(.txtatena_y.Text) <> "" Then
        s_y_atena = Trim(.txtatena_y.Text)
    End If
    
   If Trim(.txtsaishuukoushinbi.Text) <> "" Then
        s_saishuukoushinbi = Trim(.txtsaishuukoushinbi.Text)
   End If
    

    If Trim(.txtkeiyakuid.Text) = "" Then
        MsgBox "ó\ä˙ÇπÇ éñóRÇ…ÇÊÇËÅAÉGÉâÅ[Ç™î≠ê∂ÇµÇ‹ÇµÇΩÅB"
        Exit Sub
    Else
        kojindata(1) = Trim(.txtkeiyakuid.Text)
    End If
    If Trim(.txtlinkid.Text) = "" Then
        newlinkid = Space(1)
    Else
        newlinkid = Trim(.txtlinkid.Text)
    End If
    
    If Trim(.txtkeiyakubi.Text) = "" Then
        MsgBox "ó\ä˙ÇπÇ éñóRÇ…ÇÊÇËÅAÉGÉâÅ[Ç™î≠ê∂ÇµÇ‹ÇµÇΩÅB"
        Exit Sub
    Else
        kojindata(2) = Format(.txtkeiyakubi.Text, "yyyymmdd")
    End If
    
    If Trim(.txtkeiyakushafuri.Text) = "" Then
        MsgBox "å_ñÒé“Ãÿ∂ﬁ≈Å@Çì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        .txtkeiyakushafuri.SetFocus
        Exit Sub
    Else
        kojindata(3) = Trim(.txtkeiyakushafuri.Text)
    End If
    
    If Trim(.txtkeiyakushamei.Text) = "" And Trim(.txtkeiyakushamei2.Text) = "" Then
        MsgBox "å_ñÒé“ñºÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        .txtkeiyakushamei.SetFocus
        Exit Sub
    Else
        kojindata(4) = Trim(.txtkeiyakushamei.Text) & Space(1) & Trim(.txtkeiyakushamei2.Text)
    End If
    
    If Trim(.txtyuubin1.Text) <> "" Then
        If Len(.txtyuubin1.Text) <> 3 Then
            MsgBox "óXï÷î‘çÜÇÃè„ÇÕÇRåÖÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            .txtyuubin1.SetFocus
            Exit Sub
        Else
            If Trim(.txtyuubin2.Text) = "" Then
                MsgBox "óXï÷î‘çÜÇÃâ∫ÇSåÖÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                .txtyuubin2.SetFocus
                Exit Sub
            Else
                If Len(.txtyuubin2.Text) <> 4 Then
                    MsgBox "óXï÷î‘çÜÇÃâ∫ÇÕÇSåÖÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                    .txtyuubin2.SetFocus
                    Exit Sub
                Else
                    kojindata(5) = Trim(.txtyuubin1.Text) & Trim(.txtyuubin2.Text)
                End If
            End If
        End If
    End If
    
    If .comchiiki.ListIndex = -1 Then
        MsgBox "ínàÊãÊï™ÇëIëÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        .comchiiki.SetFocus
        Exit Sub
    Else
        kojindata(6) = .comchiiki.ListIndex
    End If
    If .comjuutaku.ListIndex = -1 Then
        MsgBox "èZëÓãÊï™ÇëIëÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        .comjuutaku.SetFocus
        Exit Sub
    Else
        kojindata(7) = .comjuutaku.ListIndex
        If .cmdtouroku.Caption <> "ìoò^" Then
            If CStr(.comjuutaku.ListIndex) <> Trim(txtmaekubun.Text) Then
                msg = "èZëÓãÊï™ÇïœçXÇ∑ÇÈÇ∆ÅAâﬂãéÇÃèWåvì‡óeÇ…âeãøÇµÇ‹Ç∑ÅB" & Chr(13) & _
                        "êVÇµÇ¢å_ñÒÇçÏÇÈÇ±Ç∆ÇÇ®ä©ÇﬂÇµÇ‹Ç∑ÅBÉLÉÉÉìÉZÉãÇµÇ‹Ç∑Ç©ÅH"
                res = MsgBox(msg, vbYesNo)
                If res = vbYes Then
                    Exit Sub
                End If
            End If
        End If
    End If
    If .comjuutaku2.ListIndex = -1 Then
        s_kubun = " "
    Else
         s_kubun = .comjuutaku2.ListIndex
    End If
    
    If .combcellid.Text = "" Then
        kojindata(38) = Space(1)
    Else
        kojindata(38) = Trim(.combcellid.Text)
    End If
    If .combcellno.ListIndex = -1 Then
        kojindata(39) = Space(1)
    Else
        kojindata(39) = .combcellno.Text
    End If
    
    
    If Trim(.txtjuusho1.Text) = "" Then
        MsgBox "èZèäÅiåßñºÅEésí¨ë∫ñºÅjÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        .txtjuusho1.SetFocus
        Exit Sub
    Else
        kojindata(8) = Trim(.txtjuusho1.Text)
    End If
    
    If Trim(.txtjuusho2.Text) = "" Then
        MsgBox "èZèäÅiî‘ínÅEÇªÇÃëºÅjÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        .txtjuusho2.SetFocus
        Exit Sub
    Else
        kojindata(9) = Trim(.txtjuusho2.Text)
    End If

    If Trim(.txttel1.Text) <> "" Then
        If Len(.txttel1.Text) = 12 Then
            kojindata(10) = Trim(.txttel1.Text)
        Else
            MsgBox "ìdòbî‘çÜÇÕÅAÅw-ÅxÇì¸ÇÍÇƒÅAÇPÇQåÖÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            .txttel1.SetFocus
            Exit Sub
        End If
    End If

    If Trim(.txttel2.Text) <> "" Then
        If Len(.txttel2.Text) = 12 Then
            kojindata(11) = Trim(.txttel2.Text)
        Else
            MsgBox "ìdòbî‘çÜÇÕÅAÅw-ÅxÇì¸ÇÍÇƒÅAÇPÇQåÖÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            .txttel2.SetFocus
            Exit Sub
        End If
    End If

    If Trim(.txtkeitai.Text) <> "" Then
        If Len(.txtkeitai.Text) = 13 Then
            kojindata(12) = Trim(.txtkeitai.Text)
        Else
            MsgBox "ågë—ìdòbî‘çÜÇÕÅAÅw-ÅxÇì¸ÇÍÇƒÅAÇPÇRåÖÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            .txtkeitai.SetFocus
            Exit Sub
        End If
    End If

    If Trim(.txtfax.Text) <> "" Then
        If Len(.txtfax.Text) = 12 Then
            kojindata(13) = Trim(.txtfax.Text)
        Else
            MsgBox "FAXî‘çÜÇÕÅAÅw-ÅxÇì¸ÇÍÇƒÅAÇPÇQåÖÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            .txtfax.SetFocus
            Exit Sub
        End If
    End If

    If Trim(.txtsecchifuri.Text) <> "" Then
            kojindata(14) = Trim(.txtsecchifuri.Text)
    End If

    If Trim(.txtsecchiname.Text) <> "" Then
            kojindata(15) = Trim(.txtsecchiname.Text)
    End If

    If Trim(.txtsecchijuusho.Text) <> "" Then
            kojindata(16) = Trim(.txtsecchijuusho.Text)
    End If

    If Trim(.txtsecchijuusho.Text) <> "" Then
            kojindata(16) = Trim(.txtsecchijuusho.Text)
    End If

    If Trim(.txtsecchitel.Text) <> "" Then
        If Len(.txtsecchitel.Text) = 12 Then
            kojindata(17) = Trim(.txtsecchitel.Text)
        Else
            MsgBox "ê›íuêÊìdòbî‘çÜÇÕÅAÅw-ÅxÇì¸ÇÍÇƒÅAÇPÇQåÖÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            .txtsecchitel.SetFocus
            Exit Sub
        End If
    End If

    If Trim(.txtsecchimemo.Text) <> "" Then
            kojindata(18) = Trim(.txtsecchimemo.Text)
    End If

    If Trim(.txtsetaisuu.Text) = "" Then
        MsgBox "ê¢ë—êîÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        .txtsetaisuu.SetFocus
        Exit Sub
    Else
        kojindata(19) = Trim(.txtsetaisuu.Text)
    End If
    
    If Trim(.txttvsuu.Text) = "" Then
        MsgBox "ÇsÇuêîÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        .txttvsuu.SetFocus
        Exit Sub
    Else
        kojindata(20) = Trim(.txttvsuu.Text)
    End If

    If Trim(.txthoshoukin.Text) <> "" Then
            kojindata(21) = Trim(.txthoshoukin.Text)
    End If

    If Trim(.txtht1.Text) <> "" Then
            kojindata(22) = Trim(.txtht1.Text)
    End If
    
    If Trim(.txtht2.Text) <> "" Then
            kojindata(23) = Trim(.txtht2.Text)
    End If

    If Trim(.txtht3.Text) <> "" Then
            kojindata(24) = Trim(.txtht3.Text)
    End If

    If Trim(.txtht4.Text) <> "" Then
            kojindata(25) = Trim(.txtht4.Text)
    End If

    If Trim(.txtit1.Text) <> "" Then
            kojindata(26) = Trim(.txtit1.Text)
    End If

    If Trim(.txtit2.Text) <> "" Then
            kojindata(27) = Trim(.txtit2.Text)
    End If

    If Trim(.txtit3.Text) <> "" Then
            kojindata(28) = Trim(.txtit3.Text)
    End If
    kojindata(31) = Space(1)
    kojindata(32) = Space(1)
    kojindata(33) = Space(1)
    
    
    If .optseikyuuhouhou(0).Value = True Then
        kojindata(29) = "0"
        If .cmbbank.ListIndex = -1 Then
            MsgBox "å˚ç¿éwíËÇÃèÍçáÇÕÅAã‡óZã@ä÷ñºÇëIëÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            .cmbbank.SetFocus
            Exit Sub
        End If
        kojindata(30) = bankidtbl(.cmbbank.ListIndex)
        If .optkouzashurui(0).Value = True Then
            kojindata(31) = "0"
        Else
            kojindata(31) = "1"
        End If
        If Trim(.txtkouzabangou.Text) = "" Then
            MsgBox "å˚ç¿ÇéwíËÇÃèÍçáÇÕÅAå˚ç¿î‘çÜÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            .txtkouzabangou.SetFocus
            Exit Sub
        Else
            kojindata(32) = Trim(.txtkouzabangou.Text)
        End If
        If Trim(.txtkouzameigi.Text) = "" Then
            MsgBox "å˚ç¿ÇéwíËÇÃèÍçáÇÕÅAå˚ç¿ñºã`Çì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            .txtkouzameigi.SetFocus
            Exit Sub
        Else
            kojindata(33) = Trim(.txtkouzameigi.Text)
        End If
    ElseIf .optseikyuuhouhou(1).Value = True Then
        kojindata(29) = "1"
        
    End If
    
    Dim s_henkou As Integer
    Dim s_mae_shiharai As String
    Dim s_mae_ginkou As String
    s_mae_shiharai = Trim(.txtmae_shiharai.Text)
    s_mae_ginkou = Trim(.txtmae_ginkou.Text)
    s_henkou = 0
    If s_mae_shiharai <> kojindata(29) Then
        s_henkou = 1
    End If
    If s_mae_ginkou <> kojindata(30) Then
        s_henkou = 1
    End If
    
    
    If s_henkou = 1 Then
        If kenchan = 0 Then
            If frmtouroku2.lbljoukyou.Caption = "ïsâ¬" Then
                    MsgBox "ÅYíÜÇÃÉfÅ[É^Ç…Ç¬Ç´ÅAéxï•Ç¢ï˚ñ@Ç∆ã‚çsèÓïÒÇÃïœçXÇ≈Ç´Ç‹ÇπÇÒÅB"
                    Exit Sub
            End If
        End If
     End If
        
        
    If Trim(.txtbikou.Text) <> "" Then
            kojindata(34) = Trim(.txtbikou.Text)
    End If
    
    If .chkblack.Value = 0 Then
        kojindata(35) = "0"
    Else
        kojindata(35) = "1"
    End If
    If .chkkyuushi.Value = 0 Then
        kojindata(36) = "0"
    Else
        kojindata(36) = "1"
    End If
    If Trim(.txtkanyuu.Text) <> "" Then
        If Len(Trim(.txtkanyuu.Text)) = 10 Then
            newkojindata(0) = Format(Trim(.txtkanyuu.Text), "yyyymmdd")
        Else
            MsgBox "ì˙ïtÇÕÅu2004/02/01ÅvÇÃÇÊÇ§Ç…ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
    End If
    If Trim(.txtkanyuu2.Text) <> "" Then
        If Len(Trim(.txtkanyuu2.Text)) = 10 Then
            newkojindata(1) = Format(Trim(.txtkanyuu2.Text), "yyyymmdd")
        Else
            MsgBox "ì˙ïtÇÕÅu2004/02/01ÅvÇÃÇÊÇ§Ç…ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
    End If
    
    n_catv_kaiyakubi = Space(1)
    n_internet_kaiyakubi = Space(1)
    
     If Trim(.txtkaiyaku.Text) <> "" Then
        If Len(Trim(.txtkaiyaku.Text)) = 10 Then
            n_catv_kaiyakubi = Format(Trim(.txtkaiyaku.Text), "yyyymmdd")
        Else
            MsgBox "ì˙ïtÇÕÅu2004/02/01ÅvÇÃÇÊÇ§Ç…ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
    End If
    
      If Trim(.txtkaiyaku2.Text) <> "" Then
        If Len(Trim(.txtkaiyaku2.Text)) = 10 Then
            n_internet_kaiyakubi = Format(Trim(.txtkaiyaku2.Text), "yyyymmdd")
        Else
            MsgBox "ì˙ïtÇÕÅu2004/02/01ÅvÇÃÇÊÇ§Ç…ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
    End If
    
    
    If Trim(.txtkoujihi2.Text) <> "" Then
        On Error GoTo suuji
            chkkingaku = CLng(Trim(.txtkoujihi2.Text))
        On Error GoTo 0
        If chkkingaku < 0 Then
            MsgBox "ê≥ÇÃêÆêîÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
        newkojindata(38) = Trim(.txtkoujihi2.Text)
    End If
    If Trim(.txtwaribiki2.Text) <> "" Then
        On Error GoTo suuji
            chkkingaku = CLng(Trim(.txtwaribiki2.Text))
        On Error GoTo 0
        If chkkingaku < 0 Then
            MsgBox "ê≥ÇÃêÆêîÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
        newkojindata(39) = Trim(.txtwaribiki2.Text)
    End If
    If Trim(.txtriyuu2.Text) <> "" Then
        newkojindata(40) = Trim(.txtriyuu2.Text)
    End If
    
    
    
    If Trim(.txtkoujihi.Text) <> "" Then
        On Error GoTo suuji
            chkkingaku = CLng(Trim(.txtkoujihi.Text))
        On Error GoTo 0
        If chkkingaku < 0 Then
            MsgBox "ê≥ÇÃêÆêîÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
        newkojindata(2) = Trim(.txtkoujihi.Text)
    End If
    If Trim(.txtwaribiki.Text) <> "" Then
        On Error GoTo suuji
            chkkingaku = CLng(Trim(.txtwaribiki.Text))
        On Error GoTo 0
        If chkkingaku < 0 Then
            MsgBox "ê≥ÇÃêÆêîÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
        newkojindata(3) = Trim(.txtwaribiki.Text)
    End If
    If Trim(.txtriyuu.Text) <> "" Then
        newkojindata(4) = Trim(.txtriyuu.Text)
    End If
    If Trim(.txtnyuukinbi.Text) <> "" Then
        If Len(.txtnyuukinbi.Text) = 10 Then
            newkojindata(5) = Format(Trim(.txtnyuukinbi.Text), "yyyymmdd")
        Else
            MsgBox "ì˙ïtÇÕÅu2004/02/01ÅvÇÃÇÊÇ§Ç…ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
    End If
    If Trim(.txtnyuukingaku.Text) <> "" Then
        On Error GoTo suuji
            chkkingaku = CLng(Trim(.txtnyuukingaku.Text))
        On Error GoTo 0
        If chkkingaku < 0 Then
            MsgBox "ê≥ÇÃêÆêîÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
        newkojindata(6) = Trim(.txtnyuukingaku.Text)
    End If
    If Trim(.txtnyuukinnen.Text) <> "" Then
        If Trim(.txtnyuukintsuki.Text) = "" Then
            MsgBox "åéÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        Else
            On Error GoTo suuji
                sknen = CLng(.txtnyuukinnen.Text)
                sktsuki = CLng(.txtnyuukintsuki.Text)
            On Error GoTo 0
            newkojindata(7) = Trim(.txtnyuukinnen.Text) & Trim(.txtnyuukintsuki.Text)
        End If
    End If
    If Trim(.txtzangaku.Text) <> "" Then
        On Error GoTo suuji
            chkkingaku = CLng(Trim(.txtzangaku.Text))
        On Error GoTo 0
        newkojindata(8) = Trim(.txtzangaku.Text)
    End If
    If .chkq.Value <> 1 Then
        newkojindata(9) = Space(1)
        newkojindata(10) = Space(1)
    Else
        newkojindata(9) = "1"
        If Trim(.txtqnen.Text) = "" Then
            MsgBox "îNÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        Else
            If Trim(.txtqtsuki.Text) = "" Then
                MsgBox "åéÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                Exit Sub
            Else
                On Error GoTo suuji
                    sknen = CLng(.txtqnen.Text)
                    sktsuki = CLng(.txtqtsuki.Text)
                On Error GoTo 0
                newkojindata(10) = Trim(.txtqnen.Text) & Trim(.txtqtsuki.Text)
                If Len(Trim(.txtqnen.Text)) <> 4 Then
                    MsgBox "ÇpÇÃåéÇÕÅAëSÇSåÖÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                    Exit Sub
                End If
                If Len(Trim(.txtqtsuki.Text)) <> 2 Then
                    MsgBox "ÇpÇÃì˙ÇÕÅAëSÇQåÖÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                    Exit Sub
                End If
            End If
        End If
    End If
    If Trim(.txtkarijuusho.Text) <> "" Then
        newkojindata(11) = Trim(.txtkarijuusho.Text)
    End If
    For nkdi = 0 To 9
        If Trim(.txtvoip(nkdi).Text) <> "" Then
            newkojindata(12 + nkdi) = Trim(.txtvoip(nkdi).Text)
        End If
        If Trim(.txtvoiphiduke(nkdi).Text) <> "" Then
            If Len(Trim(.txtvoiphiduke(nkdi).Text)) = 10 Then
                newkojindata(22 + nkdi) = Format(Trim(.txtvoiphiduke(nkdi).Text), "yyyymmdd")
            Else
                MsgBox "ì˙ïtÇÅu2004/12/24ÅvÇÃólÇ…ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                Exit Sub
            End If
        End If
    Next nkdi
    If Trim(.txtkyukouza1.Text) <> "" Then
        newkojindata(32) = Trim(.txtkyukouza1.Text)
    End If
    If Trim(.txtkyukouza2.Text) <> "" Then
        newkojindata(33) = Trim(.txtkyukouza2.Text)
    End If
    If Trim(.txtkanrikouza.Text) <> "" Then
        newkojindata(34) = Trim(.txtkanrikouza.Text)
    End If
    If Trim(.txtnhkno.Text) <> "" Then
        newkojindata(35) = Trim(.txtnhkno.Text)
    End If
    If Trim(.txtnhkname.Text) <> "" Then
        newkojindata(36) = Trim(.txtnhkname.Text)
    End If
    If Trim(.txtnhkfuri.Text) <> "" Then
        newkojindata(37) = Trim(.txtnhkfuri.Text)
    End If
    If Trim(.txtfkaisuu.Text) = "" Then
        newfkaisuu = "0"
    Else
        newfkaisuu = Trim(.txtfkaisuu.Text)
    End If
    
    If chkend.Value = 1 Then
        neweend = "1"
    Else
        neweend = "0"
    End If
    
  
    
    
    If .cmdtouroku.Caption = "ìoò^" Then
        sql = "select*from kojin where tel1='" & kojindata(10) & "' and tel1 <> space(1)"
        If FcSQlGet(rs, sql, PrMsg) = True Then
            msg = "ìØÇ∂ìdòbî‘çÜÇ™ìoò^Ç≥ÇÍÇƒÇ¢Ç‹Ç∑ÅB" & "å_ñÒé“ñºÅw" & rs!Name & "ÅxÇ≈Ç∑ÅB" & Chr(13) & _
                    "ìoò^Ç∑ÇÈèÍçáÇÕÅuYESÅvÇÅAÇµÇ»Ç¢èÍçáÇÕÅuNOÅvÇâüÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            res = MsgBox(msg, vbYesNo)
            If res <> vbYes Then
                Exit Sub
            End If
            rs.Close
        End If
        sql = "INSERT INTO kojin VALUES('" & kojindata(1) & "','" & kojindata(2) & "','" & kojindata(3) & "','" & kojindata(4) & "','" & kojindata(5) & "','" & kojindata(6) & "','" & kojindata(7) & "','" & s_kubun & "'," & _
                    "'" & kojindata(8) & "','" & kojindata(9) & "','" & kojindata(10) & "','" & kojindata(11) & "','" & kojindata(12) & "','" & kojindata(13) & "','" & kojindata(14) & "','" & kojindata(15) & "'," & _
                    "'" & kojindata(16) & "','" & kojindata(17) & "','" & kojindata(18) & "','" & kojindata(19) & "','" & kojindata(20) & "','" & kojindata(21) & "','" & kojindata(22) & "','" & kojindata(23) & "'," & _
                    "'" & kojindata(24) & "','" & kojindata(25) & "','" & kojindata(26) & "','" & kojindata(27) & "','" & kojindata(28) & "','" & kojindata(29) & "','" & kojindata(30) & "','" & kojindata(31) & "'," & _
                    "'" & kojindata(32) & "','" & kojindata(33) & "','" & kojindata(34) & "','" & Space(1) & "','0','0','0','0','0','0','0','0','0','0'," & _
                    "'" & kojindata(35) & "','" & kojindata(36) & "','0'," & _
                    "'" & newkojindata(0) & "','" & newkojindata(1) & "','" & newkojindata(2) & "','" & newkojindata(3) & "'," & _
                    "'" & newkojindata(4) & "','" & newkojindata(5) & "','" & newkojindata(6) & "','" & newkojindata(7) & "'," & _
                    "'" & newkojindata(8) & "','" & newkojindata(9) & "','" & newkojindata(10) & "','" & newkojindata(11) & "'," & _
                    "'" & newkojindata(12) & "','" & newkojindata(13) & "'," & _
                    "'" & newkojindata(14) & "','" & newkojindata(15) & "','" & newkojindata(16) & "','" & newkojindata(17) & "'," & _
                    "'" & newkojindata(18) & "','" & newkojindata(19) & "','" & newkojindata(20) & "','" & newkojindata(21) & "'," & _
                    "'" & newkojindata(22) & "','" & newkojindata(23) & "','" & newkojindata(24) & "','" & newkojindata(25) & "'," & _
                    "'" & newkojindata(26) & "','" & newkojindata(27) & "','" & newkojindata(28) & "','" & newkojindata(29) & "'," & _
                    "'" & newkojindata(30) & "','" & newkojindata(31) & "','" & newkojindata(32) & "','" & newkojindata(33) & "'," & _
                    "'" & newkojindata(34) & "','" & newkojindata(35) & "','" & newkojindata(36) & "','" & newkojindata(37) & "'," & _
                    "'" & newkojindata(38) & "','" & newkojindata(39) & "','" & newkojindata(40) & "','" & kojindata(38) & "'," & _
                    "'" & kojindata(39) & "','" & newlinkid & "','" & newfkaisuu & "','" & Space(1) & "','" & neweend & "','" & n_catv_kaiyakubi & "','" & _
                    n_internet_kaiyakubi & "','" & s_koujikannryoubi & "','" & s_koukubangou & "','" & s_zumen & "','" & s_bikou & "','" & _
                    s_listbangou & "','" & s_clbangou & "','" & s_vonuadress & "','" & s_donuadress & "','" & s_open_day & "','" & s_close_day & "','" & s_keiyakusha & "','" & _
                    s_bspass & "','" & s_saishuuhakkoubi & "','" & s_y_yuubin & "','" & s_y_juusho1 & "','" & s_y_juusho2 & "','" & s_saishuukoushinbi & "','" & s_y_atena & "')"
                    
                    
         
    On Error GoTo errft1
        db.Execute sql, dbSQLPassThrough
    On Error GoTo 0
        newid = CStr(CLng(kojindata(1)) + 1)
        If soukatsu(1, "renban", 1, newid, 1, 1) = False Then
                MsgBox "ê›íËÇÃèëÇ´çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
                Exit Sub
        End If
        
        
        If hyouji_touroku(kojindata(1)) = False Then
            Exit Sub
        End If
        If nhk_set(kojindata(1)) = False Then
            Screen.MousePointer = 0
            Exit Sub
        End If
        
            'Log
        Dim ss_name As String
        ss_name = frmtouroku2.lblname.Caption
         Dim ss_id As String
        ss_id = frmtouroku2.lblid.Caption
        sagyou_msg = "å⁄ãqèÓïÒÅFå⁄ãq ID " & kojindata(1) & " " & kojindata(4) & "ÇêVãKìoò^ÇµÇ‹ÇµÇΩÅB"
        log_sagyou sagyou_log_path, sagyou_msg, ss_id, ss_name

        MsgBox "ìoò^ÇµÇ‹ÇµÇΩÅB"
        Unload frmtouroku
        Set frmtouroku = Nothing
        frmtouroku2.Show 1
    Else
        'If kojindata(29) <> "0" Then
        '    If frmtouroku2.lblumu.Caption = "Ç†ÇË" Then
        '        MsgBox "Ç∑Ç≈Ç…ÇmÇgÇjÇÃå_ñÒÇ™ë∂ç›ÇµÇ‹Ç∑ÅBå˚ç¿å_ñÒà»äOÇÃïœçXÇÕÇ≈Ç´Ç‹ÇπÇÒÅBàÍíUïœçXÇíÜífÇµÅAÇmÇgÇjÇÃå_ñÒÇçÌèúÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        '        Exit Sub
        '    End If
        'End If
        sql = "UPDATE kojin SET sakuseibi = '" & kojindata(2) & "',furigana = '" & kojindata(3) & "', name = '" & kojindata(4) & "'," & _
                  "yuubin = '" & kojindata(5) & "', chiikiid = '" & kojindata(6) & "',juutaku = '" & kojindata(7) & "',juutaku2 = '" & s_kubun & "',juusho1 = '" & kojindata(8) & "'," & _
                  "juusho2 = '" & kojindata(9) & "', tel1 = '" & kojindata(10) & "',tel2 = '" & kojindata(11) & "', keitai = '" & kojindata(12) & "'," & _
                  "fax = '" & kojindata(13) & "',k_furigana = '" & kojindata(14) & "',k_name = '" & kojindata(15) & "', k_juusho = '" & kojindata(16) & "'," & _
                  "k_tel = '" & kojindata(17) & "', k_memo = '" & kojindata(18) & "',setaisuu = '" & kojindata(19) & "',tvsuu = '" & kojindata(20) & "'," & _
                  "hoshoukin = '" & kojindata(21) & "', ht1 = '" & kojindata(22) & "',ht2 = '" & kojindata(23) & "', ht3 = '" & kojindata(24) & "'," & _
                  "ht4 = '" & kojindata(25) & "',it1 = '" & kojindata(26) & "',it2 = '" & kojindata(27) & "', it3 = '" & kojindata(28) & "'," & _
                  "seikyuu = '" & kojindata(29) & "', bankid = '" & kojindata(30) & "',kouzakubun = '" & kojindata(31) & "',kouzabangou = '" & kojindata(32) & "'," & _
                  "meigi = '" & kojindata(33) & "',bikou = '" & kojindata(34) & "',f2 = '" & kojindata(35) & "',kanyuubi = '" & newkojindata(0) & "',kanyuubi2 = '" & newkojindata(1) & "',koujihi = '" & newkojindata(2) & "'," & _
                  "waribiki = '" & newkojindata(3) & "',riyuu = '" & newkojindata(4) & "',nyuukinbi = '" & newkojindata(5) & "'," & _
                  "nyuukingaku = '" & newkojindata(6) & "',kanrennbi = '" & newkojindata(7) & "',zankin = '" & newkojindata(8) & "'," & _
                  "qshiyou = '" & newkojindata(9) & "',qhiduke = '" & newkojindata(10) & "',karijusho = '" & newkojindata(11) & "'," & _
                  "v1 = '" & newkojindata(12) & "',v2 = '" & newkojindata(13) & "',v3 = '" & newkojindata(14) & "',v4 = '" & newkojindata(15) & "',v5 = '" & newkojindata(16) & "',v6 = '" & newkojindata(17) & "'," & _
                  "v7 = '" & newkojindata(18) & "',v8 = '" & newkojindata(19) & "',v9 = '" & newkojindata(20) & "',v10 = '" & newkojindata(21) & "',vh1 = '" & newkojindata(22) & "',vh2 = '" & newkojindata(23) & "'," & _
                  "vh3 = '" & newkojindata(24) & "',vh4 = '" & newkojindata(25) & "',vh5 = '" & newkojindata(26) & "',vh6 = '" & newkojindata(27) & "',vh7 = '" & newkojindata(28) & "'," & _
                  "vh8 = '" & newkojindata(29) & "',vh9 = '" & newkojindata(30) & "',vh10 = '" & newkojindata(31) & "',kyukouza = '" & newkojindata(32) & "',kyukouza2 = '" & newkojindata(33) & "',kanrikouza = '" & newkojindata(34) & "'," & _
                  "nhkno = '" & newkojindata(35) & "',nhkname = '" & newkojindata(36) & "',nhkfuri = '" & newkojindata(37) & "'," & _
                  "koujihi2 = '" & newkojindata(38) & "',waribiki2 = '" & newkojindata(39) & "',riyuu2 = '" & newkojindata(40) & "'," & _
                  "cellid = '" & kojindata(38) & "',cellno = '" & kojindata(39) & "',fkaisuu = '" & newfkaisuu & "',linkid = '" & newlinkid & "'" & _
                  ",eend = '" & neweend & "',kaiyakubi = '" & n_catv_kaiyakubi & "',kaiyakubi2 = '" & n_internet_kaiyakubi & "'" & _
                  ",saishuuhakkoubi = '" & s_saishuuhakkoubi & "',y_yuubin = '" & s_y_yuubin & "',y_juusho1 = '" & s_y_juusho1 & "'" & _
                  ",y_juusho2 = '" & s_y_juusho2 & "',saishuukoushinbi = '" & s_saishuukoushinbi & "',y_meishou = '" & s_y_atena & "'" & _
                  " WHERE kojinid = '" & kojindata(1) & "'"
                  
              

        On Error GoTo errft2
         db.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        
            'Log
        Dim sss_name As String
        sss_name = frmtouroku2.lblname.Caption
         Dim sss_id As String
        sss_id = frmtouroku2.lblid.Caption
        sagyou_msg = "å⁄ãqèÓïÒÅFå⁄ãq ID " & kojindata(1) & " " & kojindata(4) & "ÇÃì‡óeÇïœçXÇµÇ‹ÇµÇΩÅB"
        log_sagyou sagyou_log_path, sagyou_msg, sss_id, sss_name
        
        
        MsgBox "ïœçXÇµÇ‹ÇµÇΩÅB"
        If hyouji_touroku(kojindata(1)) = False Then
            Exit Sub
        End If
        Unload frmtouroku
        Set frmtouroku = Nothing
    End If

End With

Exit Sub
    
errft1:
    MsgBox "ìoò^Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
    Exit Sub
errft2:
    MsgBox "çXêVÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
    Exit Sub
suuji:
    MsgBox "îNÅEåéÇ…ÇÕîºäpêîéöÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
    Exit Sub

    
End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtbikou_GotFocus()
txtbikou.SelStart = 0
txtbikou.SelLength = Len(txtbikou.Text)

End Sub

Private Sub txtbikou_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtsetaisuu.SetFocus
End If

End Sub




Private Sub txtfax_GotFocus()
txtfax.SelStart = 0
txtfax.SelLength = Len(txtfax.Text)

End Sub

Private Sub txtfax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtbikou.SetFocus
End If

End Sub

Private Sub txthoshoukin_GotFocus()
txthoshoukin.SelStart = 0
txthoshoukin.SelLength = Len(txthoshoukin.Text)

End Sub

Private Sub txthoshoukin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtkanyuu.SetFocus
End If

End Sub

Private Sub txthoshoukin_LostFocus()

Dim kakuninsuu As Long
If Trim(txthoshoukin.Text) <> "" Then
    On Error GoTo errsu
        kakuninsuu = CLng(txthoshoukin.Text)
    On Error GoTo 0
End If
Exit Sub

errsu:
    MsgBox "îºäpêîéöÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
    Exit Sub
End Sub

Private Sub txtht1_GotFocus()
txtht1.SelStart = 0
txtht1.SelLength = Len(txtht1.Text)

End Sub

Private Sub txtht1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtit1.SetFocus
End If

End Sub

Private Sub txtht2_GotFocus()
txtht2.SelStart = 0
txtht2.SelLength = Len(txtht2.Text)

End Sub

Private Sub txtht2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtht3.SetFocus
End If

End Sub

Private Sub txtht3_GotFocus()
txtht3.SelStart = 0
txtht3.SelLength = Len(txtht3.Text)

End Sub

Private Sub txtht3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtht4.SetFocus
End If

End Sub

Private Sub txtht4_GotFocus()
txtht4.SelStart = 0
txtht4.SelLength = Len(txtht4.Text)

End Sub

Private Sub txtht4_KeyPress(KeyAscii As Integer)

Dim chkho As Long
If KeyAscii = 13 Then
    If Trim(txtht4.Text) <> "" Then
        On Error GoTo errho
            chkho = CLng(txtht4.Text)
        On Error GoTo 0
    End If
    txtvoip(0).SetFocus
End If
Exit Sub

errho:
    MsgBox "îºäpêîéöÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
    Exit Sub
End Sub

Private Sub txtit1_GotFocus()
txtit1.SelStart = 0
txtit1.SelLength = Len(txtit1.Text)

End Sub

Private Sub txtit1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtht4.SetFocus
End If

End Sub

Private Sub txtit2_GotFocus()
txtit2.SelStart = 0
txtit2.SelLength = Len(txtit2.Text)

End Sub

Private Sub txtit2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtit3.SetFocus
End If

End Sub

Private Sub txtit3_GotFocus()
txtit3.SelStart = 0
txtit3.SelLength = Len(txtit3.Text)

End Sub

Private Sub txtit3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If optseikyuuhouhou(0).Value = True Then
        optseikyuuhouhou(0).SetFocus
    ElseIf optseikyuuhouhou(1).Value = True Then
        optseikyuuhouhou(1).SetFocus
    End If
End If

End Sub

Private Sub txtjuusho1_GotFocus()
txtjuusho1.SelStart = 0
txtjuusho1.SelLength = Len(txtjuusho1.Text)

End Sub

Private Sub txtjuusho1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtjuusho2.SetFocus
End If

End Sub

Private Sub txtjuusho2_GotFocus()
txtjuusho2.SelStart = 0
txtjuusho2.SelLength = Len(txtjuusho2.Text)

End Sub

Private Sub txtjuusho2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txttel1.SetFocus
End If

End Sub

Private Sub txtkanyuu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtkanyuu2.SetFocus
End If

End Sub

Private Sub txtkanyuu2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtkoujihi.SetFocus
End If

End Sub

Private Sub txtkeitai_GotFocus()
txtkeitai.SelStart = 0
txtkeitai.SelLength = Len(txtkeitai.Text)

End Sub

Private Sub txtkeitai_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtfax.SetFocus
End If

End Sub

Private Sub txtkeiyakushafuri_GotFocus()
txtkeiyakushafuri.SelStart = 0
txtkeiyakushafuri.SelLength = Len(txtkeiyakushafuri.Text)
End Sub

Private Sub txtkeiyakushafuri_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtyuubin1.SetFocus
End If
End Sub




Private Sub txtkeiyakushamei_Furigana(Yomi As String)
If cmdtouroku.Caption = "ìoò^" Then
    If furiganahenkou2 <> 0 Then
       
        txtkeiyakushafuri.Text = Yomi
       
       furiganahenkou2 = 0
    Else
        txtkeiyakushafuri.Text = txtkeiyakushafuri.Text & Yomi
    End If
 End If
End Sub

Private Sub txtkeiyakushamei_GotFocus()
txtkeiyakushamei.SelStart = 0
txtkeiyakushamei.SelLength = Len(txtkeiyakushamei.Text)

End Sub

Private Sub txtkeiyakushamei_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtkeiyakushamei2.SetFocus
End If

End Sub

Private Sub txtkeiyakushamei2_Furigana(Yomi As String)
If cmdtouroku.Caption = "ìoò^" Then
    If furiganahenkou2 <> 0 Then
        
        txtkeiyakushafuri.Text = Yomi
       
       furiganahenkou2 = 0
    Else
        txtkeiyakushafuri.Text = txtkeiyakushafuri.Text & Yomi
    End If
  End If
    



End Sub

Private Sub txtkeiyakushamei2_GotFocus()
txtkeiyakushamei2.SelStart = 0
txtkeiyakushamei2.SelLength = Len(txtkeiyakushamei2.Text)

End Sub

Private Sub txtkeiyakushamei2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtkeiyakushafuri.SetFocus
End If

End Sub

Private Sub txtkoujihi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtwaribiki.SetFocus
End If

End Sub

Private Sub txtkoujihi2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtwaribiki2.SetFocus
End If

End Sub

Private Sub txtkouzabangou_GotFocus()
txtkouzabangou.SelStart = 0
txtkouzabangou.SelLength = Len(txtkouzabangou.Text)

End Sub

Private Sub txtkouzabangou_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtkouzameigi.SetFocus
End If

End Sub

Private Sub txtkouzameigi_GotFocus()
If cmdtouroku.Caption <> "ïœçX" Then
    txtkouzameigi.Text = Trim(txtkeiyakushafuri.Text)
End If
txtkouzameigi.SelStart = 0
txtkouzameigi.SelLength = Len(txtkouzameigi.Text)

End Sub

Private Sub txtkouzameigi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub

Private Sub txtnyuukinbi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtnyuukingaku.SetFocus
End If

End Sub

Private Sub txtnyuukingaku_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtnyuukinnen.SetFocus
End If

End Sub

Private Sub txtnyuukinnen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtnyuukintsuki.SetFocus
End If

End Sub

Private Sub txtnyuukintsuki_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtzangaku.SetFocus
End If

End Sub

Private Sub txtriyuu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtkoujihi2.SetFocus
End If

End Sub

Private Sub txtriyuu2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtnyuukinbi.SetFocus
End If

End Sub

Private Sub txtsecchifuri_GotFocus()
txtsecchifuri.SelStart = 0
txtsecchifuri.SelLength = Len(txtsecchifuri.Text)

End Sub

Private Sub txtsecchifuri_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtsecchiname.SetFocus
End If

End Sub

Private Sub txtsecchijuusho_GotFocus()
txtsecchijuusho.SelStart = 0
txtsecchijuusho.SelLength = Len(txtsecchijuusho.Text)

End Sub

Private Sub txtsecchijuusho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtsecchimemo.SetFocus
End If

End Sub

Private Sub txtsecchimemo_GotFocus()
txtsecchimemo.SelStart = 0
txtsecchimemo.SelLength = Len(txtsecchimemo.Text)

End Sub

Private Sub txtsecchimemo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtsecchitel.SetFocus
End If

End Sub

Private Sub txtsecchiname_GotFocus()
txtsecchiname.SelStart = 0
txtsecchiname.SelLength = Len(txtsecchiname.Text)

End Sub

Private Sub txtsecchiname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtsecchijuusho.SetFocus
End If

End Sub

Private Sub txtsecchitel_GotFocus()
txtsecchitel.SelStart = 0
txtsecchitel.SelLength = Len(txtsecchitel.Text)

End Sub

Private Sub txtsecchitel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtsetaisuu.SetFocus
End If

End Sub

Private Sub txtsetaisuu_GotFocus()
txtsetaisuu.SelStart = 0
txtsetaisuu.SelLength = Len(txtsetaisuu.Text)

End Sub

Private Sub txtsetaisuu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txttvsuu.SetFocus
End If

End Sub

Private Sub txttel1_GotFocus()
txttel1.SelStart = 0
txttel1.SelLength = Len(txttel1.Text)

End Sub

Private Sub txttel1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txttel2.SetFocus
End If

End Sub

Private Sub txttel2_GotFocus()
txttel2.SelStart = 0
txttel2.SelLength = Len(txttel2.Text)

End Sub

Private Sub txttel2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtkeitai.SetFocus
End If

End Sub

Private Sub txttvsuu_GotFocus()
txttvsuu.SelStart = 0
txttvsuu.SelLength = Len(txttvsuu.Text)

End Sub

Private Sub txttvsuu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    combcellid.SetFocus
End If

End Sub

Private Sub txtvoip_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    txtvoiphiduke(0).SetFocus
End If
End Sub

Private Sub txtvoiphiduke_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If optseikyuuhouhou(0).Value = True Then
        optseikyuuhouhou(0).SetFocus
    ElseIf optseikyuuhouhou(1).Value = True Then
        optseikyuuhouhou(1).SetFocus
    End If
End If

End Sub

Private Sub txtwaribiki_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtriyuu.SetFocus
End If

End Sub

Private Sub txtwaribiki2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtriyuu2.SetFocus
End If

End Sub

Private Sub txtyuubin1_GotFocus()
txtyuubin1.SelStart = 0
txtyuubin1.SelLength = Len(txtyuubin1.Text)

End Sub

Private Sub txtyuubin1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtyuubin2.SetFocus
End If

End Sub

Private Sub txtyuubin2_GotFocus()
txtyuubin2.SelStart = 0
txtyuubin2.SelLength = Len(txtyuubin2.Text)

End Sub

Private Sub txtyuubin2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    comchiiki.SetFocus
End If

End Sub

Private Sub txtzangaku_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtht1.SetFocus
End If

End Sub
