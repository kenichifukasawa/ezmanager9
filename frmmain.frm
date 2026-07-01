VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "ÇdÇyÅ@ÇlÇÅÇéÇÅÇáÇÖÇí"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   ControlBox      =   0   'False
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.CheckBox chknew 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14850
      TabIndex        =   53
      Top             =   6000
      Value           =   1  '¡™Ø∏
      Width           =   195
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ç≈ìKâªÇ‚ÉäÉXÉgÉAÇ‚ï∂èëä«óùÇé¿çs"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   10410
      TabIndex        =   13
      Top             =   6360
      Width           =   4455
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3270
         MouseIcon       =   "frmmain.frx":08CA
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":11C9
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   54
         Top             =   660
         Width           =   480
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1965
         MouseIcon       =   "frmmain.frx":200B
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":290A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   41
         Top             =   660
         Width           =   480
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   660
         MouseIcon       =   "frmmain.frx":374C
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":404B
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   24
         Top             =   660
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ï∂èëä«óù"
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
         Index           =   18
         Left            =   3030
         TabIndex        =   55
         Top             =   1215
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ê›íË"
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
         Index           =   7
         Left            =   1980
         TabIndex        =   42
         Top             =   1215
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tools"
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
         Index           =   6
         Left            =   630
         TabIndex        =   25
         Top             =   1215
         Width           =   600
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   1005
         Index           =   2
         Left            =   180
         Top             =   570
         Width           =   4035
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ÇªÇÃëºå_ñÒèÓïÒìoò^ÅEïœçXÅEçÌèú"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   4455
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1980
         MouseIcon       =   "frmmain.frx":4E8D
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":578C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   38
         Top             =   660
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÇªÇÃëºå_ñÒàÍóó"
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
         Index           =   3
         Left            =   1440
         TabIndex        =   39
         Top             =   1230
         Width           =   1665
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   1005
         Left            =   180
         Top             =   570
         Width           =   4035
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ÉfÅ[É^ÇÃämîFÅEçXêVÅ@ãyÇ—Å@âºì«Ç›çûÇ›"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   10410
      TabIndex        =   11
      Top             =   2520
      Width           =   4455
      Begin VB.PictureBox Picture15 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3030
         MouseIcon       =   "frmmain.frx":6056
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":6955
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   48
         Top             =   690
         Width           =   480
      End
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1140
         MouseIcon       =   "frmmain.frx":7797
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":8096
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   30
         Top             =   690
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "âºì«çû"
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
         Index           =   16
         Left            =   2940
         TabIndex        =   49
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label Label1 
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
         Index           =   10
         Left            =   1110
         TabIndex        =   31
         Top             =   1260
         Width           =   480
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   1005
         Index           =   6
         Left            =   210
         Top             =   570
         Width           =   4035
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ÉoÅ[ÉWÉáÉìèÓïÒ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   10170
      TabIndex        =   8
      Top             =   8280
      Width           =   4455
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1920
         MouseIcon       =   "frmmain.frx":8ED8
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":97D7
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   26
         Top             =   660
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÉoÅ[ÉWÉáÉìèÓïÒ"
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
         Left            =   1470
         TabIndex        =   27
         Top             =   1230
         Width           =   1515
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   1005
         Index           =   4
         Left            =   210
         Top             =   570
         Width           =   4035
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "êVãKã‡óZã@ä÷èÓïÒìoò^ÅEïœçXÅEçÌèú"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   360
      TabIndex        =   7
      Top             =   6360
      Width           =   4455
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2010
         MouseIcon       =   "frmmain.frx":A619
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":AF18
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   36
         Top             =   660
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ã‡óZã@ä÷àÍóó"
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
         Index           =   2
         Left            =   1590
         TabIndex        =   37
         Top             =   1230
         Width           =   1440
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   1005
         Left            =   240
         Top             =   540
         Width           =   4035
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "äeéÌàÍóóï\é¶ÅEFDçÏê¨"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   10170
      TabIndex        =   6
      Top             =   600
      Width           =   4455
      Begin VB.PictureBox Picture14 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1080
         MouseIcon       =   "frmmain.frx":B7E2
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":C0E1
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   34
         Top             =   690
         Width           =   480
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3090
         MouseIcon       =   "frmmain.frx":C9E0
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":D2DF
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   28
         Top             =   690
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "å_ñÒàÍóó"
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
         Index           =   15
         Left            =   840
         TabIndex        =   35
         Top             =   1260
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÇeÇcÇcçÏê¨"
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
         Index           =   9
         Left            =   2880
         TabIndex        =   29
         Top             =   1260
         Width           =   990
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   1005
         Index           =   5
         Left            =   240
         Top             =   600
         Width           =   4035
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ëççáåüçı"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   480
      TabIndex        =   5
      Top             =   8280
      Width           =   4455
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1920
         MouseIcon       =   "frmmain.frx":E121
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":EA20
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   20
         Top             =   660
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ëççáåüçı"
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
         Index           =   4
         Left            =   1680
         TabIndex        =   21
         Top             =   1230
         Width           =   960
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   1005
         Index           =   0
         Left            =   210
         Top             =   570
         Width           =   4035
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ÉVÉXÉeÉÄèIóπ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   5370
      TabIndex        =   4
      Top             =   8820
      Width           =   4455
      Begin VB.CheckBox chkback 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÉoÉbÉNÉAÉbÉvÇµÇ»Ç¢Ç≈èIóπÇ∑ÇÈÅB"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1320
         TabIndex        =   40
         Top             =   1560
         Value           =   1  '¡™Ø∏
         Width           =   2895
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1920
         MouseIcon       =   "frmmain.frx":F862
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":10161
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   22
         Top             =   630
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
         Index           =   5
         Left            =   1920
         TabIndex        =   23
         Top             =   1200
         Width           =   480
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   1005
         Index           =   1
         Left            =   210
         Top             =   540
         Width           =   4035
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ÉçÉO"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   10650
      TabIndex        =   3
      Top             =   4440
      Width           =   4455
      Begin VB.CheckBox chkhihyouji 
         BackColor       =   &H00FFFFFF&
         Caption         =   "çÌèúì`ï[îÒï\é¶"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   420
         TabIndex        =   52
         Top             =   300
         Value           =   1  '¡™Ø∏
         Width           =   1755
      End
      Begin VB.PictureBox Picture17 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1110
         MouseIcon       =   "frmmain.frx":10FA3
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":118A2
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   750
         Width           =   480
      End
      Begin VB.CheckBox chkend 
         BackColor       =   &H00FFFFFF&
         Caption         =   "èIóπÇï\é¶"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2580
         TabIndex        =   45
         Top             =   300
         Value           =   1  '¡™Ø∏
         Width           =   1455
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2790
         MouseIcon       =   "frmmain.frx":126E4
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":12FE3
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   43
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ó\íË"
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
         Index           =   17
         Left            =   1110
         TabIndex        =   51
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÉçÉOàÍóó"
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
         Index           =   12
         Left            =   2550
         TabIndex        =   44
         Top             =   1290
         Width           =   915
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   1005
         Index           =   3
         Left            =   240
         Top             =   630
         Width           =   4035
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ìoò^èÓïÒÅEêøãÅèëèÓïÒ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   5370
      TabIndex        =   2
      Top             =   150
      Width           =   4455
      Begin VB.PictureBox Picture11 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2010
         MouseIcon       =   "frmmain.frx":138E2
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":141E1
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   32
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00404040&
         Height          =   225
         Index           =   11
         Left            =   2010
         TabIndex        =   33
         Top             =   1290
         Width           =   480
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   1005
         Index           =   7
         Left            =   240
         Top             =   600
         Width           =   4035
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "éÂå_ñÒèÓïÒìoò^ÅEïœçXÅEçÌèú"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   4455
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2010
         MouseIcon       =   "frmmain.frx":15023
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":15922
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   18
         Top             =   630
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "éÂå_ñÒàÍóó"
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
         Index           =   1
         Left            =   1650
         TabIndex        =   19
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   1005
         Left            =   240
         Top             =   540
         Width           =   4035
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "êVãKå⁄ãqèÓïÒìoò^"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   4455
      Begin VB.PictureBox Picture16 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1920
         MouseIcon       =   "frmmain.frx":161EC
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmmain.frx":16AEB
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   16
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ìoò^ÅiêVãKå¬êlèÓïÒÅj"
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
         Index           =   14
         Left            =   1170
         TabIndex        =   17
         Top             =   1170
         Width           =   2190
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   1005
         Left            =   210
         Top             =   510
         Width           =   4035
      End
   End
   Begin VB.Label lblkanri2 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H000000FF&
      Caption         =   "ÉTÉuÉÇÅ[ÉhíÜ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   21.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5310
      TabIndex        =   47
      Top             =   4620
      Visible         =   0   'False
      Width           =   4575
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
      Left            =   5340
      TabIndex        =   46
      Top             =   7380
      Width           =   4440
   End
   Begin VB.Label lbldore 
      Alignment       =   1  'âEëµÇ¶
      BackColor       =   &H00FFFFFF&
      Caption         =   "for èüè¿CATVäîéÆâÔé–"
      BeginProperty Font 
         Name            =   "ÇlÇr ñæí©"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6090
      TabIndex        =   15
      Top             =   8190
      Width           =   3765
   End
   Begin VB.Label lblkanri 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FFFF00&
      Caption         =   "ä«óùé“égópíÜ"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   21.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6060
      TabIndex        =   14
      Top             =   5070
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.Image Image1 
      Height          =   3600
      Left            =   5850
      Picture         =   "frmmain.frx":173B5
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   3480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CATVä«óùºΩ√—"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   5100
      TabIndex        =   9
      Top             =   6930
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Main menu"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   5490
      TabIndex        =   10
      Top             =   2310
      Width           =   4335
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
















Private Sub Form_Load()
frmmain.Caption = "EZ Manager Ver." & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Image1_Click()
With frmmain
    If .Frame1.BackColor = &HFFC0C0 Then
        .Frame1.BackColor = &HC0C0FF
        .Frame2.BackColor = &HC0C0FF
        .Frame3.BackColor = &HC0C0FF
        .Frame4.BackColor = &HC0C0FF
        .Frame5.BackColor = &HC0C0FF
        .Frame6.BackColor = &HC0C0FF
        .Frame7.BackColor = &HC0C0FF
        .Frame8.BackColor = &HC0C0FF
        .Frame9.BackColor = &HC0C0FF
        .Frame10.BackColor = &HC0C0FF
        .Frame11.BackColor = &HC0C0FF
        .Frame12.BackColor = &HC0C0FF
    ElseIf .Frame1.BackColor = &HC0C0FF Then
        .Frame1.BackColor = &HC0FFC0
        .Frame2.BackColor = &HC0FFC0
        .Frame3.BackColor = &HC0FFC0
        .Frame4.BackColor = &HC0FFC0
        .Frame5.BackColor = &HC0FFC0
        .Frame6.BackColor = &HC0FFC0
        .Frame7.BackColor = &HC0FFC0
        .Frame8.BackColor = &HC0FFC0
        .Frame9.BackColor = &HC0FFC0
        .Frame10.BackColor = &HC0FFC0
        .Frame11.BackColor = &HC0FFC0
        .Frame12.BackColor = &HC0FFC0
    Else
        .Frame1.BackColor = &HFFC0C0
        .Frame2.BackColor = &HFFC0C0
        .Frame3.BackColor = &HFFC0C0
        .Frame4.BackColor = &HFFC0C0
        .Frame5.BackColor = &HFFC0C0
        .Frame6.BackColor = &HFFC0C0
        .Frame7.BackColor = &HFFC0C0
        .Frame8.BackColor = &HFFC0C0
        .Frame9.BackColor = &HFFC0C0
        .Frame10.BackColor = &HFFC0C0
        .Frame11.BackColor = &HFFC0C0
        .Frame12.BackColor = &HFFC0C0
    End If
End With
End Sub

Private Sub lblProductName_DblClick()

cl_tools_chk ("202311")

End Sub

Private Sub Picture1_Click()
    shuruiset
    frmkeiyaku.Show 1

End Sub

Private Sub Picture10_Click()
    If kiidou_mode <> 0 Then
        MsgBox "Ç±ÇÃÉÇÅ[ÉhÇ≈ÇÕégópÇ≈Ç´Ç‹ÇπÇÒÅB"
        Exit Sub
    End If
    
    
    If frmmain.lblkanri.Visible = True Then
        frmkoushin.chkkaihi.Enabled = True
        frmkoushin.chkmushi.Enabled = True
    Else
        frmkoushin.chkkaihi.Enabled = False
        frmkoushin.chkmushi.Enabled = False
     End If
        
frmkoushin.Show 1

End Sub

Private Sub Picture11_Click()

If koushinchk = False Then
    Exit Sub
End If

frmjoukyousentaku.Show 1

End Sub

Private Sub Picture12_Click()


    
If chknew.Value = 0 Then
    If chkend.Value = 1 Then
            logset 0, 0
    Else
            logset 1, 0
    End If
    frmkeiyakurireki.cmdchushutsu.Enabled = True
    frmkeiyakurireki.cmdmodoru.Enabled = True
    frmkeiyakurireki.Show 1

Else
    'If chkend.Value = 1 Then
          '  logset_2 0, 0
   ' Else
           ' logset_2 1, 0
   ' End If
   ' frmkeiyakurireki2.lblkojinid.Caption = ""
    'frmkeiyakurireki2.cmdchushutsu.Enabled = True
    'frmkeiyakurireki2.cmdmodoru.Enabled = True
    'frmkeiyakurireki2.Show 1
    
    
        If Dir$(sub_exe_path) = vbNullString Then
           MsgBox "ÉçÉOópé¿çsÉtÉ@ÉCÉãÇ™Ç†ÇËÇ‹ÇπÇÒÅB"
           'Exit Sub
        End If
    
    
        Dim sub_exe_path2 As String
        sub_exe_path2 = sub_exe_path
        'Dim objWshShell As Object
       ' Dim Ret         As Long
       ' Set objWshShell = CreateObject("WScript.Shell")
       ' Ret = objWshShell.Run(sub_exe_path2, 1, True)
        
        Dim processId As Long

        'ìdëÏÇãNìÆÇµÇ‹Ç∑
        processId = Shell(sub_exe_path2, vbNormalFocus)


        End
                
    
 End If

End Sub

Private Sub Picture13_Click()
 
 If Dir$(sub_exe_path) = vbNullString Then
       MsgBox "ÉçÉOópé¿çsÉtÉ@ÉCÉãÇ™Ç†ÇËÇ‹ÇπÇÒÅB"
      ' Exit Sub
    End If
    
    
        Dim sub_exe_path2 As String
        
  
        
        
        sub_exe_path2 = sub_exe_path & " s"
    
    
       ' Dim objWshShell As Object
      '  Dim Ret         As Long
      '  Set objWshShell = CreateObject("WScript.Shell")
        '7Ret = objWshShell.Run(sub_exe_path2, 1, True)
                
        Dim processId As Long
        
        'ìdëÏÇãNìÆÇµÇ‹Ç∑
        processId = Shell(sub_exe_path2, vbNormalFocus)
         

End Sub

Private Sub Picture14_Click()

If koushinchk = False Then
    Exit Sub
End If

frmsentaku.Show 1

End Sub

Private Sub Picture15_Click()
    If kiidou_mode <> 0 Then
        MsgBox "Ç±ÇÃÉÇÅ[ÉhÇ≈ÇÕégópÇ≈Ç´Ç‹ÇπÇÒÅB"
        Exit Sub
    End If
frmkoushin2.Show 1
End Sub

Private Sub Picture16_Click()
    
    If kiidou_mode <> 0 Then
        MsgBox "Ç±ÇÃÉÇÅ[ÉhÇ≈ÇÕégópÇ≈Ç´Ç‹ÇπÇÒÅB"
        Exit Sub
    End If
    
    If soukatsu(0, "renban", 1, "", 1, 1) = False Then
            MsgBox "ê›íËÇÃì«Ç›çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
            Exit Sub
    End If
    If reg = "0" Then
        reg = ""
    End If
    If reg = vbNullString Then
        frmtouroku.txtkeiyakuid.Text = "00001"
    Else
        frmtouroku.txtkeiyakuid.Text = Format$(reg, "0000#")
    End If
    frmtouroku.txtkeiyakubi.Text = Format(Date, "yyyy/mm/dd")

    frmtouroku.txtfkaisuu.Text = "0"
    frmtouroku.Caption = "å¬êlèÓïÒÇÃìoò^âÊñ "
    frmtouroku.cmdtouroku.Caption = "ìoò^"
    frmtouroku.txtjuusho1.Text = "éRóúåßçbèBésèüè¿í¨"
    frmtouroku.txtmaekubun.Text = ""
    
    frmtouroku.Show 1

End Sub

Private Sub Picture17_Click()

If frmmain.chkhihyouji.Value = 1 Then
    frmyotei.lblno.Caption = "1"
Else
     frmyotei.lblno.Caption = ""
End If

yoteiset

frmyotei.Picture4.Enabled = False
frmyotei.Picture5.Enabled = False



frmyotei.Show 1

End Sub

Private Sub Picture2_Click()
    bankset
    frmbank.Show 1

End Sub

Private Sub Picture3_Click()
    sonotaset
    frmsonota.Show 1

End Sub

Private Sub Picture4_Click()


If koushinchk = False Then
    Exit Sub
End If

frmkensaku.Show 1

End Sub

Private Sub Picture5_Click()

Dim bfile1 As String, douomachi As Integer, bsfile1 As String
Dim hihi As String


db.Close


If kiidou_mode <> 0 Then
    Unload frmmain
    Set frmmain = Nothing
    End
    Exit Sub
End If

'db2.Close
 Set db2 = Nothing
 


douomachi = 0
'ÉoÉbÉNÉAÉbÉvêÊ
    If soukatsu(0, "startup", 3, "", 1, 0) = False Then
            MsgBox "ê›íËÇÃì«Ç›çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
            Unload frmmain
            Set frmmain = Nothing
            End
    End If
    If reg = "0" Then
        reg = ""
    End If
    bfile1 = reg
    If bfile1 = "" Then
        douomachi = 0
    Else
        douomachi = 1
    End If
    
  
     Screen.MousePointer = 11
    FRMOMACHI.lblmsg.Caption = "ÉoÉbÉNÉAÉbÉvÇµÇƒÇ¢Ç‹Ç∑...."
    FRMOMACHI.Show
    DoEvents
        
        
    If douomachi <> 0 Then
       
        hihi = "settei.mdb"
        
        bfile1 = bfile1 & hihi
    
        On Error GoTo errbackup
            FileCopy vbname3, bfile1
        On Error GoTo 0
    
       
    End If

    If chkback.Value = 0 Then
       ' systemend
    End If

   
     Unload FRMOMACHI
    Set FRMOMACHI = Nothing
    Screen.MousePointer = 0
        
        
    Unload frmmain
    Set frmmain = Nothing
    
    End

Exit Sub

errbackup:
    MsgBox "ÉoÉbÉNÉAÉbÉvÉÅÉfÉBÉAÇ™ÉZÉbÉgÉAÉbÉvÇ≥ÇÍÇƒÇ»Ç¢ÇΩÇﬂÇ…ÅAÉoÉbÉNÉAÉbÉvÇ≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB"
    
    systemend
    Unload FRMOMACHI
    Set FRMOMACHI = Nothing
    Screen.MousePointer = 0
    Unload frmmain
    Set frmmain = Nothing
    End

End Sub

Private Sub Picture6_Click()
    If kiidou_mode <> 0 Then
        MsgBox "Ç±ÇÃÉÇÅ[ÉhÇ≈ÇÕégópÇ≈Ç´Ç‹ÇπÇÒÅB"
        Exit Sub
    End If
    frmsaitekika.Show 1

End Sub

Private Sub Picture7_Click()
    FRMSETTEI.Show 1

End Sub

Private Sub Picture8_Click()
FRMINFO.Show 1

End Sub

Private Sub Picture9_Click()
    If kiidou_mode <> 0 Then
        MsgBox "Ç±ÇÃÉÇÅ[ÉhÇ≈ÇÕégópÇ≈Ç´Ç‹ÇπÇÒÅB"
        Exit Sub
    End If
frmfdd.Show 1
End Sub
