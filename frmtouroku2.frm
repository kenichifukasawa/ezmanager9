VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7D.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmtouroku2 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "個人情報画面"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   12680
      MouseIcon       =   "frmtouroku2.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmtouroku2.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   95
      ToolTipText     =   "入金処理を行います。"
      Top             =   10260
      Width           =   480
   End
   Begin VB.Frame fbikou 
      BackColor       =   &H00C0E0FF&
      Caption         =   "備考"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   180
      TabIndex        =   85
      Top             =   1080
      Visible         =   0   'False
      Width           =   13695
      Begin VB.PictureBox Picture22 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   12900
         MouseIcon       =   "frmtouroku2.frx":11C9
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmtouroku2.frx":1AC8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   86
         Top             =   540
         Width           =   480
      End
      Begin VB.Label lblbikou 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   88
         Top             =   300
         Width           =   12435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "戻る"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Left            =   12930
         TabIndex        =   87
         Top             =   1140
         Width           =   450
      End
      Begin VB.Shape Shape7 
         BackStyle       =   1  '不透明
         BorderStyle     =   0  '透明
         Height          =   1245
         Left            =   12660
         Top             =   300
         Width           =   945
      End
   End
   Begin VB.CommandButton cmbbikou 
      Caption         =   "..."
      Height          =   315
      Left            =   13500
      TabIndex        =   89
      Top             =   1380
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture17 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   11760
      MouseIcon       =   "frmtouroku2.frx":2392
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmtouroku2.frx":2C91
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   71
      Top             =   10260
      Width           =   480
   End
   Begin VB.PictureBox Picture16 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   13600
      MouseIcon       =   "frmtouroku2.frx":3AD3
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmtouroku2.frx":43D2
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   66
      Top             =   10260
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   14520
      MouseIcon       =   "frmtouroku2.frx":4C9C
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmtouroku2.frx":559B
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   64
      Top             =   10260
      Width           =   480
   End
   Begin VB.Frame Frame2 
      Caption         =   "NHK契約"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   42
      Top             =   4860
      Width           =   11295
      Begin VB.PictureBox Picture11 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   9540
         MouseIcon       =   "frmtouroku2.frx":5E65
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmtouroku2.frx":6764
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   55
         Top             =   210
         Width           =   480
      End
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   10440
         MouseIcon       =   "frmtouroku2.frx":702E
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmtouroku2.frx":792D
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   53
         Top             =   210
         Width           =   480
      End
      Begin VB.Label lblorikomi 
         Alignment       =   2  '中央揃え
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9960
         TabIndex        =   69
         Top             =   180
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "口"
         Height          =   195
         Left            =   5760
         TabIndex        =   63
         Top             =   780
         Width           =   255
      End
      Begin VB.Label lblkuchisuu 
         Alignment       =   2  '中央揃え
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4740
         TabIndex        =   62
         Top             =   720
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "契約数："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   3900
         TabIndex        =   61
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "円"
         Height          =   255
         Left            =   8940
         TabIndex        =   59
         Top             =   780
         Width           =   255
      End
      Begin VB.Label lblnhkid 
         Alignment       =   2  '中央揃え
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   810
         TabIndex        =   57
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         Caption         =   "削除"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Left            =   9420
         TabIndex        =   56
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "変更"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Left            =   10470
         TabIndex        =   54
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblkingaku 
         Alignment       =   2  '中央揃え
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7260
         TabIndex        =   52
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label lblshurui 
         Alignment       =   2  '中央揃え
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1260
         TabIndex        =   51
         Top             =   720
         Width           =   2280
      End
      Begin VB.Label lbljikaibi 
         Alignment       =   2  '中央揃え
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7980
         TabIndex        =   50
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label lblmaebi 
         Alignment       =   2  '中央揃え
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   49
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label lblumu 
         Alignment       =   2  '中央揃え
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3180
         TabIndex        =   48
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "契約の金額："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   6060
         TabIndex        =   47
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "契約の種類："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   46
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "次回請求日："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   6780
         TabIndex        =   45
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "前回請求日："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   44
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "契約の有無："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   1980
         TabIndex        =   43
         Top             =   300
         Width           =   1215
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  '不透明
         BorderStyle     =   0  '透明
         Height          =   855
         Left            =   9240
         Top             =   180
         Width           =   1920
      End
      Begin VB.Label Label4 
         Caption         =   "契約ID："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   58
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "登録情報"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4725
      Left            =   120
      TabIndex        =   21
      Top             =   60
      Width           =   15075
      Begin VB.PictureBox Picture15 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   14190
         MouseIcon       =   "frmtouroku2.frx":81F7
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmtouroku2.frx":8AF6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   97
         Top             =   2445
         Width           =   480
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   14190
         MouseIcon       =   "frmtouroku2.frx":9938
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmtouroku2.frx":A237
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   93
         ToolTipText     =   "登録済個人情報を変更します。"
         Top             =   975
         Width           =   480
      End
      Begin VB.CommandButton Command2 
         Caption         =   "MAP"
         Height          =   315
         Left            =   13380
         TabIndex        =   91
         Top             =   90
         Width           =   495
      End
      Begin VB.PictureBox Picture18 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   14190
         MouseIcon       =   "frmtouroku2.frx":AB01
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmtouroku2.frx":B400
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   73
         ToolTipText     =   "契約履歴を表示します。"
         Top             =   1710
         Width           =   480
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   14190
         MouseIcon       =   "frmtouroku2.frx":C242
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmtouroku2.frx":CB41
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   29
         ToolTipText     =   "登録済個人情報を変更します。"
         Top             =   240
         Width           =   480
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   14190
         MouseIcon       =   "frmtouroku2.frx":D40B
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmtouroku2.frx":DD0A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   26
         ToolTipText     =   "入金履歴を表示します。"
         Top             =   3915
         Width           =   480
      End
      Begin VB.PictureBox Picture14 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   14190
         MouseIcon       =   "frmtouroku2.frx":EB4C
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmtouroku2.frx":F44B
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   24
         ToolTipText     =   "契約履歴を表示します。"
         Top             =   3180
         Width           =   480
      End
      Begin VSFlex7DAOCtl.VSFlexGrid grdtourokujoukyou 
         Height          =   4395
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   13635
         _cx             =   24051
         _cy             =   7752
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   12648447
         ForeColor       =   -2147483640
         BackColorFixed  =   12648447
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   12648447
         BackColorAlternate=   16777215
         GridColor       =   14737632
         GridColorFixed  =   -2147483632
         TreeColor       =   12648447
         FloodColor      =   192
         SheetBorder     =   12648447
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   20
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmtouroku2.frx":1028D
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         Begin VB.Label lblid 
            Alignment       =   2  '中央揃え
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "文書一覧"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Left            =   13950
         TabIndex        =   98
         Top             =   2940
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "光化工事"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Left            =   13980
         TabIndex        =   94
         Top             =   1455
         Width           =   960
      End
      Begin VB.Label lblname 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   13920
         TabIndex        =   92
         Top             =   0
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblmap 
         Alignment       =   2  '中央揃え
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   13710
         TabIndex        =   90
         Top             =   120
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblkeiyaku 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   13800
         TabIndex        =   70
         Top             =   3900
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblseikyuu 
         Alignment       =   2  '中央揃え
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   14760
         TabIndex        =   82
         Top             =   3930
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "個人詳細"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Left            =   13950
         TabIndex        =   74
         Top             =   2190
         Width           =   960
      End
      Begin VB.Label lblhouhou 
         Alignment       =   2  '中央揃え
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   14730
         TabIndex        =   68
         Top             =   4380
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lbljoukyou 
         Alignment       =   2  '中央揃え
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   13950
         TabIndex        =   60
         Top             =   4380
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ログ"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Left            =   14190
         TabIndex        =   28
         Top             =   3678
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "個人変更"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Left            =   13950
         TabIndex        =   25
         Top             =   732
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         Caption         =   "履歴"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Left            =   13950
         TabIndex        =   27
         Top             =   4410
         Width           =   975
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  '不透明
         BorderStyle     =   0  '透明
         Height          =   4485
         Left            =   13860
         Top             =   180
         Width           =   1125
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "その他の契約"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1830
      Left            =   120
      TabIndex        =   15
      Top             =   9240
      Width           =   11295
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   9630
         MouseIcon       =   "frmtouroku2.frx":102B8
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmtouroku2.frx":10BB7
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   34
         ToolTipText     =   "登録処理を行います。"
         Top             =   240
         Width           =   480
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   10470
         MouseIcon       =   "frmtouroku2.frx":11481
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmtouroku2.frx":11D80
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   31
         Top             =   900
         Width           =   480
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   9660
         MouseIcon       =   "frmtouroku2.frx":1264A
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmtouroku2.frx":12F49
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   30
         Top             =   930
         Width           =   480
      End
      Begin MSFlexGridLib.MSFlexGrid sonotagrd 
         Height          =   1455
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   5
         BackColor       =   16777215
         SelectionMode   =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         Caption         =   "登録"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Left            =   9570
         TabIndex        =   35
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         Caption         =   "削除"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   3
         Left            =   10380
         TabIndex        =   33
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "変更"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Left            =   9630
         TabIndex        =   32
         Top             =   1410
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '不透明
         BorderStyle     =   0  '透明
         Height          =   1545
         Left            =   9450
         Top             =   180
         Width           =   1710
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "契約詳細"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3225
      Left            =   120
      TabIndex        =   9
      Top             =   5970
      Width           =   11295
      Begin VB.PictureBox Picture19 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   10440
         MouseIcon       =   "frmtouroku2.frx":13813
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmtouroku2.frx":14112
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   101
         ToolTipText     =   "契約履歴を表示します。"
         Top             =   2370
         Width           =   480
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   10440
         MouseIcon       =   "frmtouroku2.frx":14F54
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmtouroku2.frx":15853
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   40
         ToolTipText     =   "登録処理を行います。"
         Top             =   210
         Width           =   480
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   10440
         MouseIcon       =   "frmtouroku2.frx":1611D
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmtouroku2.frx":16A1C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   37
         Top             =   1650
         Width           =   480
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   10440
         MouseIcon       =   "frmtouroku2.frx":172E6
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmtouroku2.frx":17BE5
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   36
         Top             =   930
         Width           =   480
      End
      Begin MSFlexGridLib.MSFlexGrid keiyakugrd 
         Height          =   2835
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   5001
         _Version        =   393216
         Cols            =   7
         SelectionMode   =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         Caption         =   "特殊"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   20
         Left            =   10380
         TabIndex        =   100
         Top             =   2850
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         Caption         =   "登録"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Left            =   10380
         TabIndex        =   41
         Top             =   690
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         BackColor       =   &H00FFFFFF&
         Caption         =   "削除"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   6
         Left            =   10380
         TabIndex        =   39
         Top             =   2100
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "変更"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Left            =   10410
         TabIndex        =   38
         Top             =   1410
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  '不透明
         BorderStyle     =   0  '透明
         Height          =   3015
         Left            =   10290
         Top             =   180
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "決済情報"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5205
      Left            =   11580
      TabIndex        =   0
      Top             =   4860
      Width           =   3615
      Begin VB.CommandButton cmdflg 
         Caption         =   "FLG"
         Height          =   315
         Left            =   3120
         TabIndex        =   84
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtzenbi_moto 
         Alignment       =   1  '右揃え
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   83
         Top             =   1140
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.TextBox txtkoushinsumi 
         Alignment       =   1  '右揃え
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   80
         Top             =   240
         Width           =   1500
      End
      Begin VB.TextBox txtjikaiseikyuubi 
         Alignment       =   1  '右揃え
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   78
         Top             =   681
         Width           =   1500
      End
      Begin VB.TextBox txtnhk 
         Alignment       =   1  '右揃え
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   76
         Top             =   2445
         Width           =   1500
      End
      Begin VB.TextBox txtshukeiyakuseikyuu 
         Alignment       =   1  '右揃え
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   19
         Top             =   3327
         Width           =   1500
      End
      Begin VB.TextBox txtminyuuruikei 
         Alignment       =   1  '右揃え
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   2004
         Width           =   1500
      End
      Begin VB.TextBox txtsonotashoukei 
         Alignment       =   1  '右揃え
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   3768
         Width           =   1500
      End
      Begin VB.TextBox txtsoukei 
         Alignment       =   1  '右揃え
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   4650
         Width           =   1500
      End
      Begin VB.TextBox txtsonotaseikyuu 
         Alignment       =   1  '右揃え
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   4209
         Width           =   1500
      End
      Begin VB.TextBox txtshukeiyakushoukei 
         Alignment       =   1  '右揃え
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   2886
         Width           =   1500
      End
      Begin VB.TextBox txtzenkin 
         Alignment       =   1  '右揃え
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   1563
         Width           =   1500
      End
      Begin VB.TextBox txtzenbi 
         Alignment       =   1  '右揃え
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   1122
         Width           =   1500
      End
      Begin VB.Label Label9 
         Caption         =   "更新済年月："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   81
         Top             =   330
         Width           =   1635
      End
      Begin VB.Label Label6 
         Caption         =   "次回請求日："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   79
         Top             =   768
         Width           =   1635
      End
      Begin VB.Label Label7 
         Caption         =   "NHK請求額："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   77
         Top             =   2520
         Width           =   1635
      End
      Begin VB.Label Label15 
         Caption         =   "主契約請求金："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   20
         Top             =   3456
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "未支払累計金："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   18
         Top             =   2082
         Width           =   1635
      End
      Begin VB.Label Label3 
         Caption         =   "その他小計："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   14
         Top             =   3894
         Width           =   1635
      End
      Begin VB.Label Label2 
         Caption         =   "今回請求金額："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   12
         Top             =   4770
         Width           =   1635
      End
      Begin VB.Label Label12 
         Caption         =   "その他請求金："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   8
         Top             =   4332
         Width           =   1875
      End
      Begin VB.Label Label11 
         Caption         =   "主契約小計："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2958
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "前回入金金額："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   4
         Top             =   1644
         Width           =   1635
      End
      Begin VB.Label Label10 
         Caption         =   "前回入金日："
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         TabIndex        =   2
         Top             =   1206
         Width           =   1635
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   255
      Left            =   14640
      TabIndex        =   75
      Top             =   10440
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "特殊"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   19
      Left            =   10620
      TabIndex        =   99
      Top             =   8880
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "入金"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   12
      Left            =   12660
      TabIndex        =   96
      Top             =   10800
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "一覧"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Left            =   11760
      TabIndex        =   72
      Top             =   10800
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "変更"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Left            =   13680
      TabIndex        =   67
      Top             =   10800
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "戻る"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Left            =   14580
      TabIndex        =   65
      Top             =   10800
      Width           =   450
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   855
      Left            =   11520
      Top             =   10230
      Width           =   3660
   End
End
Attribute VB_Name = "frmtouroku2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit








Private Sub cmbbikou_Click()



fbikou.Visible = True

End Sub

Private Sub cmdflg_Click()
If flgup2(0, lblid.Caption) = False Then
    MsgBox "フラッグの更新に失敗しました。"
Else
    MsgBox "フラッグの更新に成功しました。個人画面を再度表示し直してください。"
End If

End Sub

Private Sub Command1_Click()
Picture13_Click
End Sub

Private Sub Command2_Click()

Dim s_map_all_path As String
Dim s_adress As String
Dim RetVal

    s_adress = Trim(lblmap.Caption)
    
    If s_adress = "" Then
        MsgBox "住所情報が不正です。管理者に連絡してください。"
        Exit Sub
    End If

    s_map_all_path = map_exe_path & " s=" & s_adress
    
    RetVal = Shell(s_map_all_path, 0)   ' 電卓を実行します。
    

    
End Sub

Private Sub Form_Load()


With frmtouroku2
    .keiyakugrd.Font.Size = 11
    .sonotagrd.Font.Size = 11
    
    .Picture19.Enabled = False
    
    If kenchan = 1 Then
        .txtzenbi.Enabled = True
        .txtzenkin.Enabled = True
        .txtminyuuruikei.Enabled = True
       ' .txtsannyuukingaku.Enabled = True
        .txtshukeiyakuseikyuu.Enabled = True
        .txtshukeiyakushoukei.Enabled = True
        .txtsonotaseikyuu.Enabled = True
        .txtsonotashoukei.Enabled = True
        .txtsoukei.Enabled = True
        .txtjikaiseikyuubi.Enabled = True
        .txtkoushinsumi.Enabled = True
        '.cmdflg.Visible = True
        
        .Picture19.Enabled = True
    End If
    
 
If Dir$(map_exe_path) = vbNullString Then
   .Command2.Enabled = False
Else
   .Command2.Enabled = True
End If
    
    
    
End With
End Sub

Private Sub Picture1_Click()

frmrirekisentaku.Show 1

End Sub

Private Sub Picture10_Click()

    If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    End If
    If frmmain.lblkanri2.Visible = True Then
      MsgBox "サブモードでは更新できません。"
        Exit Sub
End If

Dim sql_nhk2 As String, rs_nhk2 As Recordset
With frmnhk
    If frmtouroku2.lblumu.Caption = "なし" Then '新規
        If frmtouroku2.lblhouhou.Caption <> "0" Then
            MsgBox "ＮＨＫの契約は、口座引落のみです。"
            Exit Sub
        End If
        If soukatsu(0, "renban", 9, "", 1, 1) = False Then
                MsgBox "設定の読み込みに失敗しました。"
                Exit Sub
        End If
        If reg = "0" Then
            reg = ""
        End If
        If reg = vbNullString Then
            .lblnhkid.Caption = "0001"
        Else
            .lblnhkid.Caption = Format(reg, "000#")
        End If
        .combshurui.ListIndex = 0
        .combzennen.Enabled = False
        .combzentsuki.Enabled = False
        .SSCJIKKOU.Caption = "登録"
    Else '変更
        If kenchan = 0 Then
            If frmtouroku2.lbljoukyou.Caption = "不可" Then
                MsgBox "〆中のデータにつき、変更できません。"
                Exit Sub
            End If
        End If
        If kidou4 = False Then
            Exit Sub
        End If
        sql_nhk2 = "select * from nhk_m where nhkid='" & lblnhkid.Caption & "'"
        If FcSQlGet4(rs_nhk2, sql_nhk2, PrMsg) < 1 Then
            MsgBox "ＮＨＫの契約情報がありません。"
            db4.Close
            Exit Sub
        Else
            .lblnhkid.Caption = rs_nhk2!nhkid
            .combzennen.Enabled = True
            .combzentsuki.Enabled = True
            .combzennen.Text = Mid(rs_nhk2!zenkai, 1, 4)
            .combzentsuki.Text = Mid(rs_nhk2!zenkai, 5, 2)
            .combjikainen.Text = Mid(rs_nhk2!jikai, 1, 4)
            henkoumae_nen = Mid(rs_nhk2!jikai, 1, 4)
            .combjikaitsuki.Text = Mid(rs_nhk2!jikai, 5, 2)
            henkoumae_tsuki = Mid(rs_nhk2!jikai, 5, 2)
            .combshurui.ListIndex = rs_nhk2!shurui
            .txtkuchisuu.Text = rs_nhk2!n1
            .txtkingaku.Text = rs_nhk2!kingaku
            'henkoumae_kingaku = rs_nhk2!kingaku
            rs_nhk2.Close
            db4.Close
            .SSCJIKKOU.Caption = "変更"
        End If
    End If
    .Show 1
End With
End Sub

Private Sub Picture11_Click()

    If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    End If
If frmmain.lblkanri2.Visible = True Then
      MsgBox "サブモードでは更新できません。"
        Exit Sub
End If
     
    Dim sql_nhk5 As String, nhk_msg As String, nhk_res
    
    If kenchan = 0 Then
        If frmtouroku2.lbljoukyou.Caption = "不可" Then
            MsgBox "〆中のデータにつき、変更できません。"
            Exit Sub
        End If
    End If
    
    nhk_msg = "ＮＨＫの契約を削除しても良いですか？"
     nhk_res = MsgBox(nhk_msg, vbYesNo)
     If nhk_res = vbYes Then
        If Trim(lblnhkid.Caption) = "" Then
            MsgBox "ＮＨＫＩＤが不正です。"
            Exit Sub
        End If
        sql_nhk5 = "delete * from nhk_m where nhkid ='" & lblnhkid.Caption & "'"
        If kidou4 = False Then
            Exit Sub
        End If
        On Error GoTo errsakujo_nhk
         db4.Execute sql_nhk5, dbSQLPassThrough
        On Error GoTo 0
        db4.Close
        If a_nhk_henkou(frmtouroku2.lblid.Caption, "0", False) = False Then
            MsgBox "料金の更新に失敗しました。"
            Exit Sub
        End If
        
        'Log
        Dim ss_name As String
        Dim ss_id As String
        ss_id = frmtouroku2.lblid.Caption
        ss_name = frmtouroku2.lblname.Caption
        sagyou_msg = "NHK処理：NHK ID " & Trim(lblnhkid.Caption) & "の契約を削除しました。"
        log_sagyou sagyou_log_path, sagyou_msg, ss_id, ss_name
        
        MsgBox "削除しました。"
        If nhk_set(frmtouroku2.lblid.Caption) = -1 Then
            Screen.MousePointer = 0
            Unload frmtouroku2
            Set frmtouroku2 = Nothing
            Exit Sub
        End If
    Else
        MsgBox "中止しました。"
        Exit Sub
    End If
    
    Exit Sub
    
errsakujo_nhk:
    MsgBox "ＮＨＫの契約データの削除に失敗しました。"
    Exit Sub
End Sub



Private Sub Picture12_Click()


 If kiidou_mode = 1 Then
    MsgBox "このモードでは使用できません。"
    Exit Sub
 End If
 
 
 Dim k1 As String, k2 As String, kn1 As String, kn2 As String, klen As Long, knlen As Long

sql = "select * from kojin where kojinid='" & lblid.Caption & "'"
    If FcSQlGet(rs, sql, PrMsg) = False Then
        MsgBox "予期せぬ事由によりエラーが発生しました。"
        Exit Sub
    Else
        With frmtouroku3
            .txtkeiyakuid.Text = rs!kojinid
            
            .lblname.Caption = rs!Name
            
            If IsNull(rs!h_koujibangou) Then
                .cmbkoukubangou.ListIndex = -1
            Else
                If Trim(rs!h_koujibangou) = "" Then
                     .cmbkoukubangou.ListIndex = -1
                Else
                    .cmbkoukubangou.ListIndex = CInt(rs!h_koujibangou) - 1
                End If
            End If
            If IsNull(rs!h_zumen) Then
                .cmbzumen.ListIndex = -1
            Else
            If Trim(rs!h_zumen) = "" Then
                    .cmbzumen.ListIndex = -1
                Else
                    
                     .cmbzumen.Text = Trim(rs!h_zumen)
                End If
            End If
            If IsNull(rs!h_koujikanryoubi) Then
                .txtkoujikanryoubi.Text = ""
            Else
                If Trim(rs!h_koujikanryoubi) = "" Then
                   .txtkoujikanryoubi.Text = ""
                Else
                     .txtkoujikanryoubi.Text = Format(rs!h_koujikanryoubi, "@@@@/@@/@@")
                End If
            End If
            
            If IsNull(rs!h_listbangou) Then
                .txtlistbangou.Text = ""
            Else
                If Trim(rs!h_listbangou) = "" Then
                   .txtlistbangou.Text = ""
                Else
                     .txtlistbangou.Text = Trim(rs!h_listbangou)
                End If
            End If
            
            If IsNull(rs!cl_bangou) Then
                .txtcl1.Text = ""
            Else
                .txtcl1.Text = Trim(rs!cl_bangou)
            End If
            
            If IsNull(rs!v_onu_adress) Then
                .txtvonuadress.Text = ""
            Else
                .txtvonuadress.Text = Trim(rs!v_onu_adress)
            End If
            
             If IsNull(rs!d_onu_adress) Then
                .txtdonuadress.Text = ""
            Else
                .txtdonuadress.Text = Trim(rs!d_onu_adress)
            End If
            
            If IsNull(rs!h_bikou) Then
                .txtbikou.Text = ""
            Else
                If Trim(rs!h_bikou) = "" Then
                   .txtbikou.Text = ""
                Else
                     .txtbikou.Text = Trim(rs!h_bikou)
                End If
            End If
            
              If IsNull(rs!cl_open_date) Then
                .txtopenday.Text = ""
            Else
                If Trim(rs!cl_open_date) = "" Then
                   .txtopenday.Text = ""
                Else
                     .txtopenday.Text = Format(rs!cl_open_date, "@@@@/@@/@@")
                End If
               
            End If
            
              If IsNull(rs!cl_close_date) Then
                .txtcloseday.Text = ""
            Else
                If Trim(rs!cl_close_date) = "" Then
                   .txtcloseday.Text = ""
                Else
                     .txtcloseday.Text = Format(rs!cl_close_date, "@@@@/@@/@@")
                End If
            End If
            
              If IsNull(rs!cl_keiyaku_name) Then
                .txtkeiyakushamei.Text = ""
            Else
                .txtkeiyakushamei.Text = Trim(rs!cl_keiyaku_name)
            End If
            
               If IsNull(rs!h_bspass) Then
                .txtbspass.Text = ""
            Else
                 If Trim(rs!h_bspass) = "" Then
                   .txtbspass.Text = ""
                Else
                     .txtbspass.Text = Format(rs!h_bspass, "@@@@/@@/@@")
                End If
            End If
           
           
            rs.Close
            
            .Show 1
        End With
    End If

 
End Sub

Private Sub Picture13_Click()

    Unload frmkekkaichi
    Set frmkekkaichi = Nothing
    Unload frmtouroku2
    Set frmtouroku2 = Nothing

End Sub

Private Sub Picture14_Click()

Dim hyoujirs As Recordset

sql = "select*from kojin where kojinid='" & lblid.Caption & "'"
If FcSQlGet(hyoujirs, sql, PrMsg) = False Then
    MsgBox "予期せぬ事由によってエラーが発生しました。"
    Exit Sub
Else
    If frmmain.chknew.Value = 0 Then
        frmkeiyakurireki.lblname.Caption = Trim(hyoujirs!Name)
    Else
        'frmkeiyakurireki2.lblname.Caption = Trim(hyoujirs!Name)
    End If
    hyoujirs.Close
End If

If frmmain.chknew.Value = 0 Then






    frmkeiyakurireki.chksakujo.Visible = True
    frmkeiyakurireki.lblkojinid.Caption = lblid.Caption
    
    logset 0, 1, lblid.Caption

    frmkeiyakurireki.Show 1
    
Else

  '  frmkeiyakurireki2.chksakujo.Visible = True
  ' frmkeiyakurireki2.lblkojinid.Caption = lblid.Caption
    
  '  logset_2 0, 1, lblid.Caption

  '  frmkeiyakurireki2.Show 1


  If Dir$(sub_exe_path) = vbNullString Then
       MsgBox "ログ用実行ファイルがありません。"
      ' Exit Sub
    End If
    
    
        Dim sub_exe_path2 As String
        Dim s_kojinid As String
  
        '自動実行　引数
        s_kojinid = Trim(lblid.Caption)
        
        sub_exe_path2 = sub_exe_path & " k" & s_kojinid
    
    
       ' Dim objWshShell As Object
      '  Dim Ret         As Long
      '  Set objWshShell = CreateObject("WScript.Shell")
        '7Ret = objWshShell.Run(sub_exe_path2, 1, True)
                
Dim processId As Long

'電卓を起動します
processId = Shell(sub_exe_path2, vbNormalFocus)


'End


End If


End Sub


Private Sub Picture15_Click()

 If Dir$(sub_exe_path) = vbNullString Then
       MsgBox "ログ用実行ファイルがありません。"
      ' Exit Sub
    End If
    
    
        Dim sub_exe_path2 As String
        Dim s_kojinid As String
  
        '自動実行　引数
        s_kojinid = Trim(lblid.Caption)
        
        sub_exe_path2 = sub_exe_path & " s" & s_kojinid
    
    
       ' Dim objWshShell As Object
      '  Dim Ret         As Long
      '  Set objWshShell = CreateObject("WScript.Shell")
        '7Ret = objWshShell.Run(sub_exe_path2, 1, True)
                
        Dim processId As Long
        
        '電卓を起動します
        processId = Shell(sub_exe_path2, vbNormalFocus)
        
        
        

End Sub

Private Sub Picture16_Click()
Dim newzenbi As String, newzenkin As String, newzenbilen As Integer, testnewzenkin As Long
Dim sel_qq As String, newminyuu As String, newminyuu2 As Long, newsannyuu As String, newsannyuu2 As Long
Dim newshusei As String, newshusei2 As Long, newshushou As String, newshushou2 As Long
Dim newsosei As String, newsosei2 As Long, newsoshou As String, newsoshou2 As Long, newkajou As String, newkajou2 As Long
Dim lenseihi As Integer, newsougou As String, newsougou2 As Long, newsumisumi As String
Dim newsumilen As Integer
Dim karisannyuu As Long, newminou As String, kariminou As Long
Dim rrres, msmsmsms As String, rrresss, newsql As String
Dim newjikaibi As String, newjikaibilen As Integer, newnewnew As String, newnewnew2 As Long
Dim sql_hhh As String, rs_hhh As Recordset


    If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    End If
If frmmain.lblkanri2.Visible = True Then
      MsgBox "サブモードでは更新できません。"
        Exit Sub
End If


If kenchan = 1 Then
        
    msmsmsms = "修正を開始します。よろしいですか？"
    rrres = MsgBox(msmsmsms, vbYesNo)
    If rrres = vbNo Then
        Exit Sub
    End If
Else
    
    MsgBox "権限がありません。"
    Exit Sub
End If
        
        With frmtouroku2
        newjikaibi = Trim(txtjikaiseikyuubi.Text)
        If newjikaibi = "1" Then
            newjikaibi = "1"
        ElseIf newjikaibi = "0" Then
            newjikaibi = "0"
        Else
            If newjikaibi <> "" Then
                newjikaibilen = Len(newjikaibi)
                If newjikaibilen <> 7 Then
                    MsgBox "正確に入力してください。例）2000/01"
                    Exit Sub
                End If
                newjikaibi = Format(newjikaibi, "yyyymm")
            End If
        End If
        newjikaibi = "0"
        newzenbi = Trim(txtzenbi.Text)
        If newzenbi <> "" Then
            newzenbilen = Len(newzenbi)
            If newzenbilen <> 10 Then
                MsgBox "正確に入力してください。例）2000/01/01"
                Exit Sub
            End If
            newzenbi = Format(newzenbi, "yyyymmdd")
        Else
            newzenbi = Space(1)
        End If
        
        newsumisumi = Trim(txtkoushinsumi.Text)
        If newsumisumi <> "" Then
            newsumilen = Len(newsumisumi)
            If newsumilen <> 7 Then
                MsgBox "正確に入力してください。例）2000/01"
                Exit Sub
            End If
            newsumisumi = Format(newsumisumi, "yyyymm")
        Else
            newsumisumi = Space(1)
        End If
        
        
        newzenkin = Trim(txtzenkin.Text)
        On Error GoTo errsuuji
            testnewzenkin = CLng(newzenkin)
        On Error GoTo 0
        
        newminyuu = Trim(.txtminyuuruikei.Text)
        On Error GoTo errsuuji
            newminyuu2 = CLng(newminyuu)
        On Error GoTo 0
        newshusei = Trim(.txtshukeiyakuseikyuu.Text)
        On Error GoTo errsuuji
            newshusei2 = CLng(newshusei)
        On Error GoTo 0
        newshushou = Trim(.txtshukeiyakushoukei.Text)
        On Error GoTo errsuuji
            newshushou2 = CLng(newshushou)
        On Error GoTo 0
        newsosei = Trim(.txtsonotaseikyuu.Text)
        On Error GoTo errsuuji
            newsosei2 = CLng(newsosei)
        On Error GoTo 0
        newsoshou = Trim(.txtsonotashoukei.Text)
        On Error GoTo errsuuji
            newsoshou2 = CLng(newsoshou)
        On Error GoTo 0
        newsougou = Trim(.txtsoukei.Text)
        On Error GoTo errsuuji
            newsougou2 = CLng(newsougou)
        On Error GoTo 0
        
        If newzenbi <> Trim(.txtzenbi_moto.Text) Then
            MsgBox "必ず、管理者モードで起動し、再度個人情報を表示、入金履歴の変更で再度伝票を修正してください。"
        End If
        sel_qq = "update kojin set maebi='" & newzenbi & "'," & _
                "maekin='" & newzenkin & "',minyuukin='" & newminyuu & "'," & _
                "jikaikin='" & newshusei & "',kajou='" & newjikaibi & "'," & _
                "jikaishu='" & newshushou & "',sonotakin='" & newsosei & "'," & _
                "jikaibi='" & newsumisumi & "'," & _
                "jikaita='" & newsoshou & "'," & _
                "soukei='" & newsougou & "'" & _
                " where kojinid='" & lblid.Caption & "'"
        
        On Error GoTo errmaebi
         db.Execute sel_qq, dbSQLPassThrough
        On Error GoTo 0
        .txtzenbi_moto.Text = newzenbi
        MsgBox "変更しました。"
        End With

Exit Sub
errsuuji2:
    MsgBox "IDは正確に入力してください。"
    Exit Sub

errsuuji:
    MsgBox "金額は正確に入力してください。"
    Exit Sub
errmaebi:
    MsgBox "変更できませんでした。"
    Exit Sub
errsan:
    MsgBox "半角数字を入力してください。"
    'txtsannyuukingaku.SetFocus
    Exit Sub

End Sub

Private Sub Picture17_Click()
Unload frmtouroku2
Set frmtouroku2 = Nothing
frmkekkaichi.Show 1
End Sub
Private Sub Picture18_Click()

stb_grid_set2 (frmtouroku2.lblid.Caption)

ppv_grid_set (frmtouroku2.lblid.Caption)

cl_grid_set (frmtouroku2.lblid.Caption)


sql = "select * from kojin where kojinid='" & lblid.Caption & "'"
    If FcSQlGet(rs, sql, PrMsg) = False Then
        MsgBox "予期せぬ事由によりエラーが発生しました。"
        Exit Sub
    Else
        With frmkojinshousai
            .txtvoip(0).Text = rs!v1
            .txtvoip(1).Text = rs!v2
            .txtvoip(2).Text = rs!v3
            .txtvoip(3).Text = rs!v4
            .txtvoip(4).Text = rs!v5
            .txtvoip(5).Text = rs!v6
            .txtvoip(6).Text = rs!v7
            .txtvoip(7).Text = rs!v8
            .txtvoip(8).Text = rs!v9
            .txtvoip(9).Text = rs!v10
            If Trim(rs!vh1) <> "" Then
                .txtvoiphiduke(0).Text = Format(rs!vh1, "@@@@/@@/@@")
            End If
            If Trim(rs!vh2) <> "" Then
                .txtvoiphiduke(1).Text = Format(rs!vh2, "@@@@/@@/@@")
            End If
            If Trim(rs!vh3) <> "" Then
                .txtvoiphiduke(2).Text = Format(rs!vh3, "@@@@/@@/@@")
            End If
            If Trim(rs!vh4) <> "" Then
                .txtvoiphiduke(3).Text = Format(rs!vh4, "@@@@/@@/@@")
            End If
            If Trim(rs!vh5) <> "" Then
                .txtvoiphiduke(4).Text = Format(rs!vh5, "@@@@/@@/@@")
            End If
            If Trim(rs!vh6) <> "" Then
                .txtvoiphiduke(5).Text = Format(rs!vh6, "@@@@/@@/@@")
            End If
            If Trim(rs!vh7) <> "" Then
                .txtvoiphiduke(6).Text = Format(rs!vh7, "@@@@/@@/@@")
            End If
            If Trim(rs!vh8) <> "" Then
                .txtvoiphiduke(7).Text = Format(rs!vh8, "@@@@/@@/@@")
            End If
            If Trim(rs!vh9) <> "" Then
                .txtvoiphiduke(8).Text = Format(rs!vh9, "@@@@/@@/@@")
            End If
            If Trim(rs!vh10) <> "" Then
                .txtvoiphiduke(9).Text = Format(rs!vh10, "@@@@/@@/@@")
            End If
            .txtht1.Text = rs!ht1
            .txtht2.Text = rs!ht2
            .txtht3.Text = rs!ht3
            .txtit1.Text = rs!it1
            .txtit2.Text = rs!it2
            .txtit3.Text = rs!it3
            
            .Show 1
        End With
    End If

End Sub

Private Sub Picture19_Click()


msg = "指定した契約をその金額で、未支払伝票を作成しますか？"

 res = MsgBox(msg, vbYesNo)
 
If res = vbNo Then
    Exit Sub
End If



Dim s_kingaku As String

s_kingaku = Trim(txtsoukei.Text)

If s_kingaku = "" Or s_kingaku = "0" Then
    MsgBox "今回請求金額がありません。"
    Exit Sub
End If

Dim newkojinid As String, newseikyuubi As String, newshurui As String, newshurui_no As String, newkin As String, newbankid As String
Dim h_id As String

            frmtouroku2.keiyakugrd.Col = 0
            h_id = keiyakugrd.Text
            'sql = "select * from keiyaku where keiyakuid = '" & h_id & "'"
            sql = "SELECT keiyaku.keiyakuid, shurui.shuruimei, keiyaku.kaishi,keiyaku.shime," & _
                "keiyaku.kadou, keiyaku.kuchisuu, keiyaku.kingaku, keiyaku.sadou, keiyaku.kojinid, keiyaku.shurui" & _
                " FROM shurui RIGHT JOIN keiyaku ON shurui.shuruiid = keiyaku.shurui" & _
                " WHERE (((keiyaku.keiyakuid)='" & h_id & "'));"
            
            Set rs = db.OpenRecordset(sql, dbOpenDynaset, dbReadOnly)
            If FcSQlGet(rs, sql, PrMsg) = False Then
                MsgBox "作成したい項目を正確選択してから、実行してください。"
                Exit Sub
            Else
                If rs!kadou = "0" Then
                Else
                    MsgBox "作成したい項目は、稼動していません。"
                    Exit Sub
                End If
                
                newkojinid = rs!kojinid
                
                 newshurui_no = rs!shurui
                 
                 newkin = rs!kingaku
                 
                
               
                
            End If


 If s_kingaku <> newkin Then
    MsgBox "作成したい項目の金額は、今回請求金額と違います。"
    Exit Sub
End If


newseikyuubi = Format(Date, "yyyymm") ' "202105" ' Mid(newnyukinbi, 1, 6)

newshurui = "1"
newbankid = Space(1)

 If shousai_main_set(newkojinid, newseikyuubi, newshurui, newshurui_no, newkin, newbankid, "1") = -1 Then
            MsgBox "契約の更新に失敗しました。"
            Exit Sub
  End If
  
  
 
  
  
  MsgBox "未支払伝票を作成しました。"
  
  Exit Sub
  
errko1:
    
    MsgBox "個人の請求と未収金額の更新に失敗しました。手動で行ってください。"
    
    Exit Sub



End Sub

Private Sub Picture2_Click()

    If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    End If
If frmmain.lblkanri2.Visible = True Then
      MsgBox "サブモードでは更新できません。"
        Exit Sub
End If
If Trim(frmtouroku2.txtsoukei.Text) <> "0" Then
    frmnyuukin.Height = 1710
    If Trim(frmtouroku2.lblid.Caption) = "" Then
        MsgBox "個人IDを参照できません。"
        Exit Sub
    End If
    
    If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If
    nyuukinrirekiset3 (Format(frmtouroku2.lblid.Caption, "00000"))
    

    
Else
    MsgBox "請求金額がないために入金処理ができません。"
End If
End Sub

Private Sub Picture22_Click()

fbikou.Visible = False

End Sub

Private Sub Picture3_Click()
Dim k1 As String, k2 As String, kn1 As String, kn2 As String, klen As Long, knlen As Long

sql = "select * from kojin where kojinid='" & lblid.Caption & "'"
    If FcSQlGet(rs, sql, PrMsg) = False Then
        MsgBox "予期せぬ事由によりエラーが発生しました。"
        Exit Sub
    Else
        With frmtouroku
            .Picture11.Enabled = False
            .cmdtouroku.Caption = "変更"
            .Caption = "個人情報の変更画面"
            .txtkeiyakuid.Text = rs!kojinid
            .txtkeiyakubi.Text = Format(rs!sakuseibi, "@@@@/@@/@@")
            .txtkeiyakushafuri.Text = rs!Furigana
            knlen = InStr(1, rs!Name, Space(1), vbTextCompare)
            If knlen <> 0 Then
                kn1 = Mid(rs!Name, 1, knlen)
                .txtkeiyakushamei.Text = Trim(kn1)
                kn2 = Mid(rs!Name, knlen)
                .txtkeiyakushamei2.Text = Trim(kn2)
            Else
                .txtkeiyakushamei.Text = rs!Name
                .txtkeiyakushamei2.Text = Space(1)
            End If
            .txtyuubin1.Text = Mid(rs!yuubin, 1, 3)
            .txtyuubin2.Text = Mid(rs!yuubin, 4, 4)
            If Trim(rs!chiikiid) = "" Then
                    .comchiiki.ListIndex = -1
                    MsgBox "地域区分が設定されていないか、不正です。再設定してください。"
            Else
                If CInt(rs!chiikiid) > 3 Or CInt(rs!chiikiid) < 0 Then
                    .comchiiki.ListIndex = -1
                    MsgBox "地域区分が設定されていないか、不正です。再設定してください。"
                Else
                    .comchiiki.ListIndex = CInt(rs!chiikiid)
                End If
            End If
            If Trim(rs!juutaku) = "" Then
                MsgBox "住宅区分が設定されていないか、不正です。再設定してください。"
                .comjuutaku.ListIndex = -1
                .txtmaekubun.Text = "-1"
            Else
                .comjuutaku.ListIndex = CInt(rs!juutaku)
                .txtmaekubun.Text = rs!juutaku
            End If
            
            If Trim(rs!juutaku2) = "" Then
                .comjuutaku2.ListIndex = -1
            Else
                If IsNull(rs!juutaku2) Then
                     .comjuutaku2.ListIndex = -1
                Else
                    .comjuutaku2.ListIndex = CInt(rs!juutaku2)
                End If
            End If
            
            .txtjuusho1.Text = rs!juusho1
            .txtjuusho2.Text = rs!juusho2
            .txttel1.Text = rs!tel1
            .txttel2.Text = rs!tel2
            .txtkeitai.Text = rs!keitai
            .txtfax.Text = rs!fax
            
            .txtbikou.Text = rs!bikou
            .txtsecchifuri.Text = rs!k_furigana
            .txtsecchiname.Text = rs!k_name
            .txtsecchijuusho.Text = rs!k_juusho
            .txtsecchitel.Text = rs!k_tel
            .txtsecchimemo.Text = rs!k_memo
            .txtsetaisuu.Text = rs!setaisuu
            .txttvsuu.Text = rs!tvsuu
            .txtht1.Text = rs!ht1
            .txtht2.Text = rs!ht2
            .txtht3.Text = rs!ht3
            .txtht4.Text = rs!ht4
            .txtit1.Text = rs!it1
            .txtit2.Text = rs!it2
            .txtit3.Text = rs!it3
            .txthoshoukin.Text = rs!hoshoukin
            If rs!f2 = "0" Then
                .chkblack.Value = 0
            Else
            .chkblack.Value = 1
            End If
            If Trim(rs!cellid) <> "" Then
                .combcellid.Text = Trim(rs!cellid)
            Else
                .combcellid.Text = ""
            End If
            If Trim(rs!linkid) <> "" Then
                .txtlinkid.Text = Trim(rs!linkid)
            Else
                .txtlinkid.Text = ""
            End If
            If Trim(rs!cellno) <> "" Then
                .combcellno.ListIndex = CInt(rs!cellno) - 1
            Else
                .combcellno.ListIndex = -1
            End If
            
            .txtmae_shiharai.Text = ""
            .txtmae_ginkou.Text = ""
            If rs!seikyuu = "0" Or rs!seikyuu = "1" Then
                .optseikyuuhouhou(CInt(rs!seikyuu)).Value = True
                If .optseikyuuhouhou(0).Value = True Then
                    .txtmae_shiharai.Text = "0"
                    .fkouza.Enabled = True
                    .cmbbank.ListIndex = FcCmbbankidGet(rs!bankid)
                    .txtmae_ginkou.Text = rs!bankid
                    .optkouzashurui(CInt(rs!kouzakubun)).Value = True
                    .txtkouzabangou.Text = rs!kouzabangou
                    .txtkouzameigi.Text = rs!meigi
                Else
                    .txtmae_shiharai.Text = "1"
                    .fkouza.Enabled = False
                End If
             End If
                
            If rs!f3 = "0" Then
                .chkkyuushi.Value = 0
            Else
                .chkkyuushi.Value = 1
            End If
            If kenchan = 1 Then
                .chkkyuushi.Enabled = True
            End If
            
            If Trim(rs!kanyuubi) <> "" Then
                .txtkanyuu.Text = Format(Trim(rs!kanyuubi), "@@@@/@@/@@")
            End If
            If Trim(rs!kanyuubi2) <> "" Then
                .txtkanyuu2.Text = Format(Trim(rs!kanyuubi2), "@@@@/@@/@@")
            End If
              If Trim(rs!kaiyakubi) <> "" Then
                .txtkaiyaku.Text = Format(Trim(rs!kaiyakubi), "@@@@/@@/@@")
            End If
            If Trim(rs!kaiyakubi2) <> "" Then
                .txtkaiyaku2.Text = Format(Trim(rs!kaiyakubi2), "@@@@/@@/@@")
            End If
            
            If Trim(rs!koujihi) <> "" Then
                .txtkoujihi.Text = Trim(rs!koujihi)
            End If
            If Trim(rs!waribiki) <> "" Then
                .txtwaribiki.Text = Trim(rs!waribiki)
            End If
            If Trim(rs!riyuu) <> "" Then
                .txtriyuu.Text = Trim(rs!riyuu)
            End If
            If Trim(rs!koujihi2) <> "" Then
                .txtkoujihi2.Text = Trim(rs!koujihi2)
            End If
            If Trim(rs!waribiki2) <> "" Then
                .txtwaribiki2.Text = Trim(rs!waribiki2)
            End If
            If Trim(rs!riyuu2) <> "" Then
                .txtriyuu2.Text = Trim(rs!riyuu2)
            End If
            If Trim(rs!nyuukinbi) <> "" Then
                .txtnyuukinbi.Text = Format(Trim(rs!nyuukinbi), "@@@@/@@/@@")
            End If
            If Trim(rs!nyuukingaku) <> "" Then
                .txtnyuukingaku.Text = Trim(rs!nyuukingaku)
            End If
            If Trim(rs!kanrennbi) <> "" Then
                .txtnyuukinnen.Text = Mid(Trim(rs!kanrennbi), 1, 4)
                .txtnyuukintsuki.Text = Mid(Trim(rs!kanrennbi), 5, 2)
            End If
            If Trim(rs!zankin) <> "" Then
                .txtzangaku.Text = Trim(rs!zankin)
            End If
            If Trim(rs!qshiyou) = "" Or rs!qshiyou = "0" Then
                .chkq.Value = 0
                .txtqnen.Text = ""
                .txtqtsuki.Text = ""
            Else
                .chkq.Value = 1
                .txtqnen.Text = Mid(Trim(rs!qhiduke), 1, 4)
                .txtqtsuki.Text = Mid(Trim(rs!qhiduke), 5, 2)
            End If
            .txtfkaisuu.Enabled = True
            If Trim(rs!fkaisuu) = "" Then
                .txtfkaisuu.Text = "0"
            Else
                .txtfkaisuu.Text = Trim(rs!fkaisuu)
            End If
            If Trim(rs!karijusho) <> "" Then
                .txtkarijuusho.Text = Trim(rs!karijusho)
            End If
            If Trim(rs!v1) <> "" Then
                .txtvoip(0).Text = Trim(rs!v1)
            End If
            If Trim(rs!vh1) <> "" Then
                .txtvoiphiduke(0).Text = Format(Trim(rs!vh1), "@@@@/@@/@@")
            End If
            If Trim(rs!v2) <> "" Then
                .txtvoip(1).Text = Trim(rs!v2)
            End If
            If Trim(rs!vh2) <> "" Then
                .txtvoiphiduke(1).Text = Format(Trim(rs!vh2), "@@@@/@@/@@")
            End If
            If Trim(rs!v3) <> "" Then
                .txtvoip(2).Text = Trim(rs!v3)
            End If
            If Trim(rs!vh3) <> "" Then
                .txtvoiphiduke(2).Text = Format(Trim(rs!vh3), "@@@@/@@/@@")
            End If
            If Trim(rs!v4) <> "" Then
                .txtvoip(3).Text = Trim(rs!v4)
            End If
            If Trim(rs!vh4) <> "" Then
                .txtvoiphiduke(3).Text = Format(Trim(rs!vh4), "@@@@/@@/@@")
            End If
            If Trim(rs!v5) <> "" Then
                .txtvoip(4).Text = Trim(rs!v5)
            End If
            If Trim(rs!vh5) <> "" Then
                .txtvoiphiduke(4).Text = Format(Trim(rs!vh5), "@@@@/@@/@@")
            End If
            If Trim(rs!v6) <> "" Then
                .txtvoip(5).Text = Trim(rs!v6)
            End If
            If Trim(rs!vh6) <> "" Then
                .txtvoiphiduke(5).Text = Format(Trim(rs!vh6), "@@@@/@@/@@")
            End If
            If Trim(rs!v7) <> "" Then
                .txtvoip(6).Text = Trim(rs!v7)
            End If
            If Trim(rs!vh7) <> "" Then
                .txtvoiphiduke(6).Text = Format(Trim(rs!vh7), "@@@@/@@/@@")
            End If
            If Trim(rs!v8) <> "" Then
                .txtvoip(7).Text = Trim(rs!v8)
            End If
            If Trim(rs!vh8) <> "" Then
                .txtvoiphiduke(7).Text = Format(Trim(rs!vh8), "@@@@/@@/@@")
            End If
            If Trim(rs!v9) <> "" Then
                .txtvoip(8).Text = Trim(rs!v9)
            End If
            If Trim(rs!vh9) <> "" Then
                .txtvoiphiduke(8).Text = Format(Trim(rs!vh9), "@@@@/@@/@@")
            End If
            If Trim(rs!v10) <> "" Then
                .txtvoip(9).Text = Trim(rs!v10)
            End If
            If Trim(rs!vh10) <> "" Then
                .txtvoiphiduke(9).Text = Format(Trim(rs!vh10), "@@@@/@@/@@")
            End If
            If Trim(rs!kyukouza) <> "" Then
                .txtkyukouza1.Text = Trim(rs!kyukouza)
            End If
            If Trim(rs!kyukouza2) <> "" Then
                .txtkyukouza2.Text = Trim(rs!kyukouza2)
            End If
            If Trim(rs!kanrikouza) <> "" Then
                .txtkanrikouza.Text = Trim(rs!kanrikouza)
            End If
            If Trim(rs!nhkno) <> "" Then
                .txtnhkno.Text = Trim(rs!nhkno)
            End If
            If Trim(rs!nhkname) <> "" Then
                .txtnhkname.Text = Trim(rs!nhkname)
            End If
            If Trim(rs!nhkfuri) <> "" Then
                .txtnhkfuri.Text = Trim(rs!nhkfuri)
            End If
            
            If rs!eend = "1" Then
                .chkend.Value = 1
            Else
                .chkend.Value = 0
            End If
            
            If IsNull(rs!y_yuubin) Then
                 .txtyuubin1_y.Text = ""
                 .txtyuubin2_y.Text = ""
            Else
                 .txtyuubin1_y.Text = Mid(Trim(rs!y_yuubin), 1, 3)
                 .txtyuubin2_y.Text = Mid(Trim(rs!y_yuubin), 5, 4)
            End If
            
             If IsNull(rs!y_juusho1) Then
                 .txtjuusho1_y.Text = ""
            Else
                 .txtjuusho1_y.Text = Trim(rs!y_juusho1)
            End If
            
             If IsNull(rs!y_juusho2) Then
                 .txtjuusho2_y.Text = ""
            Else
                 .txtjuusho2_y.Text = Trim(rs!y_juusho2)
            End If
            
             If IsNull(rs!y_meishou) Then
                 .txtatena_y.Text = ""
            Else
                 .txtatena_y.Text = Trim(rs!y_meishou)
            End If
            
             If IsNull(rs!saishuukoushinbi) Then
                 .txtsaishuukoushinbi.Text = ""
            Else
                 .txtsaishuukoushinbi.Text = Trim(rs!saishuukoushinbi)
            End If
            
            
            rs.Close
            furiganahenkou2 = 1
            .Show 1
        End With
    End If

End Sub

Private Sub Picture4_Click()

    If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    End If
If frmmain.lblkanri2.Visible = True Then
      MsgBox "サブモードでは更新できません。"
        Exit Sub
End If
Dim h_id As String

With frmkeiyaku6
    frmtouroku2.sonotagrd.Col = 0
    h_id = sonotagrd.Text
    'sql = "select * from keiyaku where keiyakuid = '" & h_id & "'"
    sql = "SELECT keiyaku2.keiyakuid2, sonota.sonota, keiyaku2.kaishi," & _
        "keiyaku2.kingaku,keiyaku2.f2,keiyaku2.sonotaid,sonota.kotei" & _
        " FROM sonota RIGHT JOIN keiyaku2 ON sonota.sonotaid = keiyaku2.sonotaid" & _
        " WHERE (((keiyaku2.keiyakuid2)='" & h_id & "'));"
    
    Set rs = db.OpenRecordset(sql, dbOpenDynaset, dbReadOnly)
    If FcSQlGet(rs, sql, PrMsg) = False Then
        MsgBox "変更したい項目を正確選択してから、実行してください。"
        Exit Sub
    Else
        If kenchan = 0 Then
            If frmtouroku2.lbljoukyou.Caption = "不可" Then
                MsgBox "〆中のデータにつき、変更できません。"
                Exit Sub
            End If
        End If
        .lblkeiyakuid.Caption = rs!keiyakuid2
        .lblsonotaid.Caption = rs!sonotaid
        .lblkeiyaku.Caption = rs!sonota
        .txtkaishibi.Text = Format(rs!kaishi, "@@@@/@@/@@")
        If rs!kotei = "0" Then '固定
            .txtkeiyakukin.Enabled = False
        Else
            .txtkeiyakukin.Enabled = True
        End If
        .txtkeiyakukin.Text = rs!kingaku
        henkoumae_kingaku = rs!kingaku
        rs.Close
    End If
    .Show 1
End With

End Sub

Private Sub Picture5_Click()

    If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    End If
If frmmain.lblkanri2.Visible = True Then
      MsgBox "サブモードでは更新できません。"
        Exit Sub
End If
Dim s_id As String, kyk As String, kyk2 As String, kyk3 As String

Dim s_name As String

With frmtouroku2
    .sonotagrd.Col = 0
    s_id = .sonotagrd.Text

    sql = "select  keiyaku2.*, sonota.sonota from sonota RIGHT JOIN keiyaku2 ON sonota.sonotaid = keiyaku2.sonotaid" & _
        " where keiyaku2.keiyakuid2='" & s_id & "'"
    If FcSQlGet(rs, sql, PrMsg) = False Then
        MsgBox "変更したい項目を正確選択してから、実行してください。"
        Exit Sub
    Else
        If kenchan = 0 Then
            If frmtouroku2.lbljoukyou.Caption = "不可" Then
                MsgBox "〆中のデータにつき、削除できません。"
                Exit Sub
            End If
        End If
            msg = "契約『" & rs!sonota & "』を削除してもよいですか？"
            s_name = rs!sonota
            kyk = .lbljoukyou.Caption
            kyk2 = rs!kingaku
            kyk3 = rs!kaishi
            rs.Close
            res = MsgBox(msg, vbYesNo)
            If res = vbYes Then
                sql = "delete * from keiyaku2 where keiyakuid2 ='" & s_id & "'"
               On Error GoTo errsakujo
                db.Execute sql, dbSQLPassThrough
               On Error GoTo 0
                If a_sonota_sakujo(.lblid.Caption, kyk2) = False Then
                    MsgBox "料金の更新に失敗しました。"
                    Exit Sub
                End If
                
                 '最終更新日更新
            
                Dim s_kojinid As String
    
                s_kojinid = frmtouroku2.lblid.Caption
                
                Dim s_kyounohi As String
                
                s_kyounohi = Format(Date, "yyyy/mm/dd")
    
                 sql = "update kojin set  saishuukoushinbi='" & s_kyounohi & "'" & _
                           " where kojinid ='" & s_kojinid & "'"
                
                On Error GoTo errsakujo2
                db.Execute sql, dbSQLPassThrough
                On Error GoTo 0
                
                'Log
                Dim ss_name As String
                Dim ss_id As String
                ss_id = frmtouroku2.lblid.Caption

                ss_name = frmtouroku2.lblname.Caption
                sagyou_msg = "その他の契約：契約 ID " & s_id & " " & s_name & "の契約を削除しました。"
                log_sagyou sagyou_log_path, sagyou_msg, ss_id, ss_name
        
                
                
                keiyakuset2 (.lblid.Caption)
                MsgBox "削除しました"
            Else
                MsgBox "中止しました。"
                Exit Sub
            End If
    End If
End With
Exit Sub

errsakujo:
    MsgBox "削除に失敗しました。"
    Exit Sub

errsakujo2:
    MsgBox "削除に失敗しました。"
    Exit Sub


End Sub

Private Sub Picture6_Click()

    If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    End If
If frmmain.lblkanri2.Visible = True Then
      MsgBox "サブモードでは更新できません。"
        Exit Sub
End If




With frmtourokusentaku
    .lblno.Caption = "1"
    
    .Show 1
End With


'frmkeiyaku5.Show 1

End Sub

Private Sub Picture7_Click()

    If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    End If
If frmmain.lblkanri2.Visible = True Then
      MsgBox "サブモードでは更新できません。"
        Exit Sub
End If
Dim h_id As String
If kenchan = 0 Then
    If frmtouroku2.lbljoukyou.Caption = "不可" Then
        MsgBox "〆中のデータにつき、変更できません。"
        Exit Sub
    End If
End If
With frmkeiyaku3

            frmtouroku2.keiyakugrd.Col = 0
            h_id = keiyakugrd.Text
            'sql = "select * from keiyaku where keiyakuid = '" & h_id & "'"
            sql = "SELECT keiyaku.keiyakuid, shurui.shuruimei, keiyaku.kaishi,keiyaku.shime," & _
                "keiyaku.kadou, keiyaku.kuchisuu, keiyaku.kingaku, keiyaku.sadou" & _
                " FROM shurui RIGHT JOIN keiyaku ON shurui.shuruiid = keiyaku.shurui" & _
                " WHERE (((keiyaku.keiyakuid)='" & h_id & "'));"
            
            Set rs = db.OpenRecordset(sql, dbOpenDynaset, dbReadOnly)
            If FcSQlGet(rs, sql, PrMsg) = False Then
                MsgBox "変更したい項目を正確選択してから、実行してください。"
                Exit Sub
            Else
                If rs!kadou = "0" Then
                    .optkadou(0).Value = True
                    henkanmoto = "0"
                ElseIf rs!kadou = "1" Then
                    .optkadou(1).Value = True
                    henkanmoto = "1"
                Else
                    .optkadou(2).Value = True
                    henkanmoto = "2"
                End If
                henkoumae_shurui = henkanmoto
                .lblkeiyakuid.Caption = rs!keiyakuid
                .lblkeiyaku.Caption = rs!shuruimei
                .txtkaishibi.Text = Format(rs!kaishi, "@@@@/@@/@@")
                .txtkeiyakusuu.Text = rs!kuchisuu
                .txtkeiyakukin.Text = rs!kingaku
                If kenchan = 1 Then
                .lblsadou.Visible = True
                .cmdsadou.Visible = True
                Else
                .lblsadou.Visible = False
                .cmdsadou.Visible = False
                End If
                .lblsadou.Caption = rs!sadou
                henkoumae_kingaku = rs!kingaku
                tanka = CLng(rs!kingaku) / CLng(rs!kuchisuu)
            End If
            .Show 1
End With

End Sub

Private Sub Picture8_Click()

    If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    End If
If frmmain.lblkanri2.Visible = True Then
      MsgBox "サブモードでは更新できません。"
        Exit Sub
End If




Dim s_id As String
Dim s_naiyou As String
Dim s_kingaku As String
Dim s_st As String

    frmtouroku2.keiyakugrd.Col = 0
    s_id = frmtouroku2.keiyakugrd.Text

 frmtouroku2.keiyakugrd.Col = 1
    s_naiyou = frmtouroku2.keiyakugrd.Text
    
     frmtouroku2.keiyakugrd.Col = 3
    s_kingaku = frmtouroku2.keiyakugrd.Text
    
         frmtouroku2.keiyakugrd.Col = 5
    s_st = frmtouroku2.keiyakugrd.Text


With frmsakujo
    .lblid.Caption = s_id
    .lblnaiyou.Caption = s_naiyou
    .lblkingaku.Caption = s_kingaku
    .lblst.Caption = s_st
    .Show 1
End With








End Sub

Private Sub Picture9_Click()

    If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    End If
    
    If frmmain.lblkanri2.Visible = True Then
      MsgBox "サブモードでは更新できません。"
        Exit Sub
    End If


With frmtourokusentaku
    .lblno.Caption = "0"
    
    .Show 1
End With




'frmkeiyaku4.Show 1




End Sub



