VERSION 5.00
Begin VB.Form frmdatacheck2 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "ïsê≥ÉfÅ[É^èCê≥"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8325
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.Frame Frame1 
      Caption         =   "ÉfÅ[É^èCê≥ëIë"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   120
      TabIndex        =   39
      Top             =   3900
      Width           =   8055
      Begin VB.TextBox txtk2id 
         Alignment       =   1  'âEëµÇ¶
         Height          =   315
         Left            =   3450
         TabIndex        =   43
         Top             =   600
         Width           =   555
      End
      Begin VB.CheckBox chktsuika 
         Caption         =   "ç∑äzÇëçäzÇ…í«â¡ÇµÅAêøãÅì`ï[Çé©ìÆçƒê∂Ç∑ÇÈÅB"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   41
         Top             =   300
         Width           =   4335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "èCê≥Ç∑ÇÈ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5430
         TabIndex        =   40
         Top             =   300
         Width           =   2520
      End
      Begin VB.Label Label10 
         Caption         =   "ì`ï[ÇçÏê¨Ç∑ÇÈç€ÇÃÅukeiyaku2ÅvÇÃIDÇÕÅA"
         Height          =   255
         Left            =   210
         TabIndex        =   42
         Top             =   660
         Width           =   3165
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "çÏã∆íÜé~"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   37
      Top             =   5130
      Width           =   2500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ÉfÅ[É^éQè∆"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   36
      Top             =   5130
      Width           =   2500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "èCê≥ÇµÇ»Ç¢"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   35
      Top             =   5130
      Width           =   2500
   End
   Begin VB.Frame Frame2 
      Caption         =   "ïsê≥ÉfÅ[É^"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   8085
      Begin VB.TextBox txtshushou 
         Alignment       =   1  'âEëµÇ¶
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
         Left            =   5250
         TabIndex        =   5
         Top             =   1290
         Width           =   1500
      End
      Begin VB.TextBox txtshusei 
         Alignment       =   1  'âEëµÇ¶
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
         Left            =   5250
         TabIndex        =   4
         Top             =   1725
         Width           =   1500
      End
      Begin VB.TextBox txtsoshou 
         Alignment       =   1  'âEëµÇ¶
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
         Left            =   5250
         TabIndex        =   3
         Top             =   2175
         Width           =   1500
      End
      Begin VB.TextBox txtsosei 
         Alignment       =   1  'âEëµÇ¶
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
         Left            =   5250
         TabIndex        =   2
         Top             =   2610
         Width           =   1500
      End
      Begin VB.TextBox txtsoukei 
         Alignment       =   1  'âEëµÇ¶
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
         Left            =   5250
         TabIndex        =   1
         Top             =   3060
         Width           =   1500
      End
      Begin VB.Label lblsagaku 
         Alignment       =   1  'âEëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6840
         TabIndex        =   38
         Top             =   3090
         Width           =   975
      End
      Begin VB.Label lblname 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3810
         TabIndex        =   34
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblid 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   33
         Top             =   330
         Width           =   1125
      End
      Begin VB.Label Label9 
         Alignment       =   2  'íÜâõëµÇ¶
         Caption         =   "å_ñÒé“ñº"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2700
         TabIndex        =   32
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label8 
         Alignment       =   2  'íÜâõëµÇ¶
         Caption         =   "å¬êlÇhÇc"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   31
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'íÜâõëµÇ¶
         Caption         =   "ÉfÅ[É^ï\é¶"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label2 
         Alignment       =   2  'íÜâõëµÇ¶
         Caption         =   "éÂå_ñÒè¨åv"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   29
         Top             =   1290
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   2  'íÜâõëµÇ¶
         Caption         =   "éÂå_ñÒêøãÅ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   28
         Top             =   1740
         Width           =   1245
      End
      Begin VB.Label Label4 
         Alignment       =   2  'íÜâõëµÇ¶
         Caption         =   "ÇªÇÃëºè¨åv"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   27
         Top             =   2190
         Width           =   1245
      End
      Begin VB.Label Label5 
         Alignment       =   2  'íÜâõëµÇ¶
         Caption         =   "ÇªÇÃëºêøãÅ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   26
         Top             =   2640
         Width           =   1245
      End
      Begin VB.Label Label6 
         Alignment       =   2  'íÜâõëµÇ¶
         Caption         =   "ëççáåv"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   25
         Top             =   3090
         Width           =   1245
      End
      Begin VB.Label Label7 
         Alignment       =   2  'íÜâõëµÇ¶
         Caption         =   "ãåÉfÅ[É^"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   24
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label lblkyushushou 
         Alignment       =   1  'âEëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   23
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label lblkyushusei 
         Alignment       =   1  'âEëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   22
         Top             =   1740
         Width           =   1245
      End
      Begin VB.Label lblkyusoshou 
         Alignment       =   1  'âEëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Top             =   2190
         Width           =   1245
      End
      Begin VB.Label lblkyusosei 
         Alignment       =   1  'âEëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Top             =   2640
         Width           =   1245
      End
      Begin VB.Label lblkyusoukei 
         Alignment       =   1  'âEëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Top             =   3090
         Width           =   1245
      End
      Begin VB.Label Label13 
         Alignment       =   2  'íÜâõëµÇ¶
         Caption         =   "çƒåvéZ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   18
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label lblnyushushou 
         Alignment       =   1  'âEëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   17
         Top             =   1290
         Width           =   1245
      End
      Begin VB.Label lblnyushusei 
         Alignment       =   1  'âEëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   16
         Top             =   1740
         Width           =   1245
      End
      Begin VB.Label lblnyusoshou 
         Alignment       =   1  'âEëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   15
         Top             =   2190
         Width           =   1245
      End
      Begin VB.Label lblnyusosei 
         Alignment       =   1  'âEëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   14
         Top             =   2640
         Width           =   1245
      End
      Begin VB.Label lblnyusoukei 
         Alignment       =   1  'âEëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   13
         Top             =   3090
         Width           =   1245
      End
      Begin VB.Label Label19 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4530
         TabIndex        =   12
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblstshushou 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4530
         TabIndex        =   11
         Top             =   1290
         Width           =   495
      End
      Begin VB.Label lblstshusei 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4530
         TabIndex        =   10
         Top             =   1740
         Width           =   495
      End
      Begin VB.Label lblstsoshou 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4530
         TabIndex        =   9
         Top             =   2190
         Width           =   495
      End
      Begin VB.Label lblstsosei 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4530
         TabIndex        =   8
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label lblstsoukei 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4530
         TabIndex        =   7
         Top             =   3090
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   180
         X2              =   7800
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line2 
         X1              =   180
         X2              =   7800
         Y1              =   1650
         Y2              =   1650
      End
      Begin VB.Line Line3 
         X1              =   180
         X2              =   7800
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Line Line4 
         X1              =   180
         X2              =   7800
         Y1              =   2550
         Y2              =   2550
      End
      Begin VB.Line Line5 
         X1              =   180
         X2              =   7800
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line6 
         X1              =   180
         X2              =   7800
         Y1              =   3450
         Y2              =   3450
      End
      Begin VB.Line Line7 
         X1              =   1590
         X2              =   1590
         Y1              =   780
         Y2              =   3450
      End
      Begin VB.Line Line8 
         X1              =   3030
         X2              =   3030
         Y1              =   780
         Y2              =   3450
      End
      Begin VB.Line Line9 
         X1              =   4440
         X2              =   4440
         Y1              =   780
         Y2              =   3450
      End
      Begin VB.Line Line10 
         X1              =   5130
         X2              =   5130
         Y1              =   780
         Y2              =   3450
      End
      Begin VB.Label Label25 
         Alignment       =   2  'íÜâõëµÇ¶
         Caption         =   "êVÇµÇ¢íl(ç∑äz)"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5220
         TabIndex        =   6
         Top             =   840
         Width           =   2385
      End
   End
End
Attribute VB_Name = "frmdatacheck2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Set frmdatacheck2 = Nothing
End Sub

Private Sub Command2_Click()
Dim new1 As String, new2 As String, new3 As String, new4 As String, new5 As String
Dim snew1 As Long, snew2 As Long, snew3 As Long, snew4 As Long, snew5 As Long
Dim sql_s As String, newsagaku As Long
Dim tid As String, tkojinid As String, tkingaku As String, jurs As Recordset
Dim tkaishi As String, tkudou As String, tshurui As String, ttid As String


On Error GoTo errsuuji
newsagaku = CLng(lblsagaku.Caption)
If Trim(txtshushou.Text) = "" Then
    new1 = "0"
Else
    snew1 = CLng(txtshushou.Text)
    new1 = CStr(txtshushou.Text)
End If
If Trim(txtshusei.Text) = "" Then
    new2 = "0"
Else
    snew2 = CLng(txtshusei.Text)
    new2 = CStr(txtshusei.Text)
End If
If Trim(txtsoshou.Text) = "" Then
    new3 = "0"
Else
    snew3 = CLng(txtsoshou.Text)
    new3 = CStr(txtsoshou.Text)
End If
If chktsuika.Value = 1 Then
    snew3 = snew3 + newsagaku
    new3 = CStr(snew3)
End If
If Trim(txtsosei.Text) = "" Then
    new4 = "0"
Else
    snew4 = CLng(txtsosei.Text)
    new4 = CStr(txtsosei.Text)
End If
If chktsuika.Value = 1 Then
    snew4 = snew4 + newsagaku
    new4 = CStr(snew4)
End If
If Trim(txtsoukei.Text) = "" Then
    new5 = "0"
Else
    snew5 = CLng(txtsoukei.Text)
    new5 = CStr(txtsoukei.Text)
End If
If chktsuika.Value = 1 Then
    snew5 = snew5 + newsagaku
    new5 = CStr(snew5)
End If
On Error GoTo 0

sql_s = "update kojin set jikaishu = '" & new1 & "'," & _
        "jikaikin = '" & new2 & "'," & _
        "jikaita = '" & new3 & "'," & _
        "sonotakin = '" & new4 & "'," & _
        "soukei = '" & new5 & "'" & _
       " where kojinid ='" & lblid.Caption & "'"

On Error GoTo verup8
    db.Execute sql_s, dbSQLPassThrough
On Error GoTo 0

'êøãÅì`ï[
If chktsuika.Value = 1 Then

            If soukatsu(0, "renban", 6, "", 1, 1) = False Then
                    MsgBox "ê›íËÇÃì«Ç›çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
                    Exit Sub
            End If
            If reg = "0" Then
                reg = ""
            End If
            If reg = vbNullString Then
                tid = "0000001"
            Else
                tid = Format$(reg, "000000#")
            End If
        With frmdatacheck2
            tshurui = .txtk2id.Text
            tkojinid = .lblid.Caption
            tkingaku = .lblsagaku.Caption
            tkaishi = Format(Date, "yyyymmdd")
            tkudou = "0"
       End With
        sql = "INSERT INTO keiyaku2 VALUES('" & tid & "','" & tkojinid & "'," & _
                       "'" & tshurui & "','" & tkaishi & "','" & tkudou & "','" & tkingaku & "')"
On Error GoTo errk8
        db.Execute sql, dbSQLPassThrough
On Error GoTo 0
        ttid = CStr(CLng(tid) + 1)
        If soukatsu(1, "renban", 6, ttid, 1, 1) = False Then
                MsgBox "ê›íËÇÃèëÇ´çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
                Exit Sub
        End If

End If

Unload Me
Set frmdatacheck2 = Nothing
Exit Sub

errsuuji:
    MsgBox "îºäpÇÃêîéöÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
    Exit Sub
verup8:
    MsgBox "èCê≥Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
    Exit Sub
errk8:
    MsgBox "èCê≥Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
    Exit Sub
End Sub

Private Sub Command3_Click()
sanshousuruyo = 1
If hyouji_touroku(lblid.Caption) = False Then
    Screen.MousePointer = 0
    Exit Sub
End If
If nhk_set(lblid.Caption) = -1 Then
    Screen.MousePointer = 0
    Exit Sub
End If
keiyakuset (lblid.Caption)
keiyakuset2 (lblid.Caption)

frmtouroku2.Picture1.Enabled = True
frmtouroku2.Picture2.Enabled = False
frmtouroku2.Picture3.Enabled = False
frmtouroku2.Picture4.Enabled = False
frmtouroku2.Picture5.Enabled = False
frmtouroku2.Picture6.Enabled = False
frmtouroku2.Picture7.Enabled = False
frmtouroku2.Picture8.Enabled = False
frmtouroku2.Picture9.Enabled = False
frmtouroku2.Picture10.Enabled = False
frmtouroku2.Picture11.Enabled = False
frmtouroku2.Picture13.Enabled = True
frmtouroku2.Picture14.Enabled = True
frmtouroku2.Picture15.Enabled = False
frmtouroku2.Picture16.Enabled = False
frmtouroku2.Picture17.Enabled = False


frmtouroku2.Show 1

End Sub

Private Sub Command4_Click()
Unload Me
Set frmdatacheck2 = Nothing
datacheck_cansel = 1
End Sub

Private Sub Form_Load()
txtk2id.Text = "32"
End Sub
