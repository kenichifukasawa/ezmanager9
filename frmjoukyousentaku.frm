VERSION 5.00
Begin VB.Form frmjoukyousentaku 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "状況選択"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   9465
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      Caption         =   "表示したい項目を選択してください。"
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
      Height          =   8295
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   9045
      Begin VB.PictureBox Picture18 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   7140
         MouseIcon       =   "frmjoukyousentaku.frx":0000
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmjoukyousentaku.frx":08FF
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   66
         Top             =   5700
         Width           =   480
      End
      Begin VB.CheckBox chkauto 
         BackColor       =   &H00FFFFFF&
         Caption         =   "自動実行する"
         Height          =   255
         Left            =   4830
         TabIndex        =   65
         Top             =   7830
         Value           =   1  'ﾁｪｯｸ
         Width           =   1425
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5340
         MouseIcon       =   "frmjoukyousentaku.frx":11FE
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmjoukyousentaku.frx":1AFD
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   63
         Top             =   7050
         Width           =   480
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "期間指定(請求年月を指定)"
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
         Height          =   1155
         Index           =   6
         Left            =   390
         TabIndex        =   56
         Top             =   6810
         Width           =   3945
         Begin VB.ComboBox cmbtsuki5 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1440
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   59
            Top             =   420
            Width           =   675
         End
         Begin VB.ComboBox cmbnen5 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   210
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   58
            Top             =   420
            Width           =   915
         End
         Begin VB.PictureBox Picture17 
            Appearance      =   0  'ﾌﾗｯﾄ
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'なし
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2910
            MouseIcon       =   "frmjoukyousentaku.frx":23FC
            MousePointer    =   99  'ﾕｰｻﾞｰ定義
            Picture         =   "frmjoukyousentaku.frx":2CFB
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   57
            Top             =   210
            Width           =   480
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "月"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   2160
            TabIndex        =   62
            Top             =   540
            Width           =   255
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "年"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   1200
            TabIndex        =   61
            Top             =   540
            Width           =   255
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "CL請求データ"
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
            Left            =   2340
            TabIndex        =   60
            Top             =   810
            Width           =   1410
         End
      End
      Begin VB.PictureBox Picture16 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5370
         MouseIcon       =   "frmjoukyousentaku.frx":35FA
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmjoukyousentaku.frx":3EF9
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   54
         Top             =   3270
         Width           =   480
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "指定年月の入金データ出力"
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
         Height          =   1155
         Index           =   5
         Left            =   390
         TabIndex        =   47
         Top             =   5550
         Width           =   3945
         Begin VB.ComboBox cmbtsuki4 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1440
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   50
            Top             =   420
            Width           =   675
         End
         Begin VB.ComboBox cmbnen4 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   180
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   49
            Top             =   420
            Width           =   915
         End
         Begin VB.PictureBox Picture15 
            Appearance      =   0  'ﾌﾗｯﾄ
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'なし
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2910
            MouseIcon       =   "frmjoukyousentaku.frx":47F8
            MousePointer    =   99  'ﾕｰｻﾞｰ定義
            Picture         =   "frmjoukyousentaku.frx":50F7
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   48
            Top             =   210
            Width           =   480
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "月"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   2160
            TabIndex        =   53
            Top             =   540
            Width           =   255
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "年"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   1200
            TabIndex        =   52
            Top             =   540
            Width           =   255
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "入金データ"
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
            Left            =   2610
            TabIndex        =   51
            Top             =   810
            Width           =   1110
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "指定年月の請求データ出力"
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
         Height          =   1155
         Index           =   4
         Left            =   390
         TabIndex        =   40
         Top             =   4281
         Width           =   3945
         Begin VB.PictureBox Picture14 
            Appearance      =   0  'ﾌﾗｯﾄ
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'なし
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2910
            MouseIcon       =   "frmjoukyousentaku.frx":59F6
            MousePointer    =   99  'ﾕｰｻﾞｰ定義
            Picture         =   "frmjoukyousentaku.frx":62F5
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   43
            Top             =   210
            Width           =   480
         End
         Begin VB.ComboBox cmbnen3 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   180
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   42
            Top             =   420
            Width           =   915
         End
         Begin VB.ComboBox cmbtsuki3 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1440
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   41
            Top             =   420
            Width           =   675
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "請求データ"
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
            Left            =   2610
            TabIndex        =   46
            Top             =   810
            Width           =   1110
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "年"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   1200
            TabIndex        =   45
            Top             =   540
            Width           =   255
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "月"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   2160
            TabIndex        =   44
            Top             =   540
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "指定年月の1日のデータ出力"
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
         Height          =   1155
         Index           =   2
         Left            =   390
         TabIndex        =   33
         Top             =   3014
         Width           =   3945
         Begin VB.ComboBox cmbtsuki2 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1440
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   36
            Top             =   420
            Width           =   675
         End
         Begin VB.ComboBox cmbnen2 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   180
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   35
            Top             =   420
            Width           =   915
         End
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'ﾌﾗｯﾄ
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'なし
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2910
            MouseIcon       =   "frmjoukyousentaku.frx":6BF4
            MousePointer    =   99  'ﾕｰｻﾞｰ定義
            Picture         =   "frmjoukyousentaku.frx":74F3
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   34
            Top             =   210
            Width           =   480
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "月"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   2160
            TabIndex        =   39
            Top             =   540
            Width           =   255
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "年"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   38
            Top             =   540
            Width           =   255
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "未収データ"
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
            Left            =   2610
            TabIndex        =   37
            Top             =   810
            Width           =   1110
         End
      End
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   7200
         MouseIcon       =   "frmjoukyousentaku.frx":7DF2
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmjoukyousentaku.frx":86F1
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   31
         Top             =   4440
         Width           =   480
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5340
         MouseIcon       =   "frmjoukyousentaku.frx":8FF0
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmjoukyousentaku.frx":98EF
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   29
         Top             =   5700
         Width           =   480
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5370
         MouseIcon       =   "frmjoukyousentaku.frx":A1EE
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmjoukyousentaku.frx":AAED
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   27
         Top             =   4530
         Width           =   480
      End
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   7650
         MouseIcon       =   "frmjoukyousentaku.frx":B3EC
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmjoukyousentaku.frx":BCEB
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   25
         Top             =   7080
         Width           =   480
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   7230
         MouseIcon       =   "frmjoukyousentaku.frx":C5B5
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmjoukyousentaku.frx":CEB4
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   23
         Top             =   3240
         Width           =   480
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   7320
         MouseIcon       =   "frmjoukyousentaku.frx":D7B3
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmjoukyousentaku.frx":E0B2
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   17
         Top             =   660
         Width           =   480
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5430
         MouseIcon       =   "frmjoukyousentaku.frx":E9B1
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmjoukyousentaku.frx":F2B0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   15
         Top             =   1920
         Width           =   480
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   7290
         MouseIcon       =   "frmjoukyousentaku.frx":FBAF
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmjoukyousentaku.frx":104AE
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   13
         Top             =   1890
         Width           =   480
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'なし
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5430
         MouseIcon       =   "frmjoukyousentaku.frx":10DAD
         MousePointer    =   99  'ﾕｰｻﾞｰ定義
         Picture         =   "frmjoukyousentaku.frx":116AC
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   11
         Top             =   660
         Width           =   480
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "期間指定(請求年月を指定)"
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
         Height          =   1155
         Index           =   1
         Left            =   390
         TabIndex        =   6
         Top             =   1747
         Width           =   3945
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'ﾌﾗｯﾄ
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'なし
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2910
            MouseIcon       =   "frmjoukyousentaku.frx":11FAB
            MousePointer    =   99  'ﾕｰｻﾞｰ定義
            Picture         =   "frmjoukyousentaku.frx":128AA
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   19
            Top             =   210
            Width           =   480
         End
         Begin VB.ComboBox cmbnen 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   210
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   8
            Top             =   420
            Width           =   915
         End
         Begin VB.ComboBox cmbtsuki 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1440
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   7
            Top             =   420
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "VoIP請求データ"
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
            Left            =   2190
            TabIndex        =   20
            Top             =   810
            Width           =   1620
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "年"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   10
            Top             =   540
            Width           =   255
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "月"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   9
            Top             =   540
            Width           =   255
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "期間指定(請求した年を指定)"
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
         Height          =   1155
         Index           =   0
         Left            =   390
         TabIndex        =   1
         Top             =   480
         Width           =   3945
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'ﾌﾗｯﾄ
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'なし
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2910
            MouseIcon       =   "frmjoukyousentaku.frx":131A9
            MousePointer    =   99  'ﾕｰｻﾞｰ定義
            Picture         =   "frmjoukyousentaku.frx":13AA8
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   21
            Top             =   210
            Width           =   480
         End
         Begin VB.ComboBox combzentsuki2 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1320
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   4
            Top             =   390
            Width           =   675
         End
         Begin VB.ComboBox combzennen 
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   2
            Top             =   390
            Width           =   915
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "NHK状況リスト"
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
            Left            =   2190
            TabIndex        =   22
            Top             =   810
            Width           =   1530
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "月"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   2040
            TabIndex        =   5
            Top             =   510
            Width           =   255
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "年"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   3
            Top             =   510
            Width           =   255
         End
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "書類出力"
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
         Left            =   6840
         TabIndex        =   67
         Top             =   6300
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cadix出力"
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
         Left            =   4980
         TabIndex        =   64
         Top             =   7560
         Width           =   1065
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "月別請求集計"
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
         Left            =   4920
         TabIndex        =   55
         Top             =   3780
         Width           =   1440
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "仮読込出力"
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
         Left            =   6810
         TabIndex        =   32
         Top             =   5040
         Width           =   1200
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "最新０更新日"
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
         Left            =   4920
         TabIndex        =   30
         Top             =   6300
         Width           =   1365
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "CSV出力"
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
         Left            =   5100
         TabIndex        =   28
         Top             =   5040
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "終了"
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
         Left            =   7650
         TabIndex        =   26
         Top             =   7590
         Width           =   495
      End
      Begin VB.Label SSCJIKKOU 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "現金入金リスト"
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
         Left            =   6780
         TabIndex        =   24
         Top             =   3750
         Width           =   1530
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "金融機関集計"
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
         Left            =   6900
         TabIndex        =   18
         Top             =   1260
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "請求入金集計"
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
         Left            =   5010
         TabIndex        =   16
         Top             =   2520
         Width           =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "滞納者リスト"
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
         Left            =   6870
         TabIndex        =   14
         Top             =   2490
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "登録状況集計"
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
         Left            =   5010
         TabIndex        =   12
         Top             =   1260
         Width           =   1440
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  '不透明
         BorderStyle     =   0  '透明
         Height          =   7875
         Left            =   150
         Top             =   270
         Width           =   8700
      End
   End
End
Attribute VB_Name = "frmjoukyousentaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Form_Load()
Dim in2 As Integer, shitenen As Long

For in2 = 1 To 12
    If in2 <= 5 Then
        shitenen = CLng(Format(Date, "yyyy")) - 4 + in2
        frmjoukyousentaku.combzennen.AddItem CStr(shitenen)
        frmjoukyousentaku.cmbnen.AddItem CStr(shitenen)
        frmjoukyousentaku.cmbnen2.AddItem CStr(shitenen)
        frmjoukyousentaku.cmbnen3.AddItem CStr(shitenen)
        frmjoukyousentaku.cmbnen4.AddItem CStr(shitenen)
        frmjoukyousentaku.cmbnen5.AddItem CStr(shitenen)
    End If
    frmjoukyousentaku.combzentsuki2.AddItem Format(in2, "00")
    frmjoukyousentaku.cmbtsuki.AddItem Format(in2, "00")
    frmjoukyousentaku.cmbtsuki2.AddItem Format(in2, "00")
    frmjoukyousentaku.cmbtsuki3.AddItem Format(in2, "00")
    frmjoukyousentaku.cmbtsuki4.AddItem Format(in2, "00")
    frmjoukyousentaku.cmbtsuki5.AddItem Format(in2, "00")
Next in2
frmjoukyousentaku.combzennen.ListIndex = 3
frmjoukyousentaku.cmbnen.ListIndex = 3
frmjoukyousentaku.cmbnen2.ListIndex = 3
frmjoukyousentaku.cmbnen3.ListIndex = 3
frmjoukyousentaku.cmbnen4.ListIndex = 3
frmjoukyousentaku.cmbnen5.ListIndex = 3

End Sub




Private Sub Picture1_Click()

    Unload Me
    frmjoukyou.Show 1

End Sub

Private Sub Picture10_Click()
Unload Me
frmkari.Show 1

End Sub

Private Sub Picture11_Click()

 Dim n_nen As String, n_tsuki As String, n_hi As String
        If cmbnen2.ListIndex = -1 Then
            MsgBox "抽出したい年を選択してください"
            Exit Sub
        Else
            n_nen = cmbnen2.Text
        End If
        If cmbtsuki2.ListIndex = -1 Then
            MsgBox "抽出したい月を選択してください"
            Exit Sub
        Else
            n_tsuki = cmbtsuki2.Text
        End If
        n_hi = n_nen & n_tsuki
        
        

    Dim udtBrowseInfo As BROWSEINFO
    Dim lngIDList     As Long
    Dim strPath       As String
    Dim shiteppp As String
On Error GoTo errcsv
    '[フォルダの参照]ダイアログ情報を設定
    With udtBrowseInfo
        .hWndOwner = Me.hwnd
        .pidlRoot = CSIDL_NETWORK
        .pszDisplayName = String(MAX_PATH2, vbNullChar)
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
On Error GoTo 0
On Error GoTo errcsv2

    '[フォルダの参照]ダイアログを表示
    lngIDList = SHBrowseForFolder(udtBrowseInfo)
    If lngIDList <> 0 Then
        strPath = String(MAX_PATH, vbNullChar)
        'ITEMIDLISTをパスに変換
        SHGetPathFromIDList lngIDList, strPath
        If Left(strPath, 1) = vbNullChar Then
            'パスが空の場合はコンピュータ名を取得
            With udtBrowseInfo
                strPath = "\\" & Left(.pszDisplayName, _
                                      InStr(.pszDisplayName, _
                                            vbNullChar) - 1)
            End With
        Else
            strPath = Left(strPath, _
                           InStr(strPath, _
                                 vbNullChar) - 1)
        End If
        '結果をラベルに表示
       shiteppp = strPath
        'ITEMIDLISTを解放
        CoTaskMemFree lngIDList
    Else
        Exit Sub
    End If
On Error GoTo 0




    If shiteppp = "" Then
        MsgBox "保存先を指定してから実行してください。"
        Exit Sub
    End If
    shiteppp = shiteppp & "\未収契約詳細_" & n_hi & "_" & Format(Date, "yyyymmdd") & "-" & Format(Time, "hhmmss") & ".csv"
    
    csv_csv5 shiteppp, n_hi
    
Exit Sub
errcsv:
        MsgBox "ダイアログ設定エラー"
        Exit Sub
errcsv2:
        MsgBox "パス選択エラー:" & strPath
        Exit Sub
        
        
        
End Sub



Private Sub Picture12_Click()

    Dim s_hozonsaki As String, s_filename As String, s_mydoc As String, s_desktop As String
    Dim s_hozonsaki_path As String, shiteppp As String

    Dim objwsh As Object
    Set objwsh = CreateObject("Wscript.Shell")
    s_mydoc = objwsh.SpecialFolders("MyDocuments")
    s_hozonsaki_path = s_mydoc & "\久保田\cadix転送データ\"
    
    shiteppp = mapa & "subs.csv"
        
     If Dir$(s_hozonsaki_path, vbDirectory) = "" Then
             MsgBox "「" & s_hozonsaki_path & "」フォルダーが存在しないため保存できません。"
        Exit Sub
    End If
    
    If Dir$(ftp_exe_path) = vbNullString Then
       MsgBox "FTP転送用実行ファイルがありません。"
       Exit Sub
    End If


    
     msg = "Cadix用のデータをFTP転送しますか？"
    res = MsgBox(msg, vbYesNo)
    If res <> vbYes Then
        Exit Sub
    End If
    
    
    
    
    
    Dim s_path As String
    
    s_path = csv_csv_cl(shiteppp)
    
    
    If s_path <> "" Then
        'ファイルを保存
       
        
        s_filename = "subs_" & Format(Date, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".csv"

        
        s_hozonsaki = s_hozonsaki_path & s_filename

        
        FileCopy shiteppp, s_hozonsaki
        
        
        'FTP転送
        
        Dim ftp_exe_path2 As String
        '自動実行　引数
        If chkauto.Value = 1 Then
            ftp_exe_path2 = ftp_exe_path & " k1"
        Else
            ftp_exe_path2 = ftp_exe_path
        End If
        
        Dim RetVal
        RetVal = Shell(ftp_exe_path2, 1)   ' 電卓を実行します。

  
    End If
    
    

End Sub




Private Sub Picture13_Click()
Unload Me

End Sub

Private Sub Picture14_Click()
Dim n_nen As String, n_tsuki As String, n_hi As String
        If cmbnen3.ListIndex = -1 Then
            MsgBox "抽出したい年を選択してください"
            Exit Sub
        Else
            n_nen = cmbnen3.Text
        End If
        If cmbtsuki3.ListIndex = -1 Then
            MsgBox "抽出したい月を選択してください"
            Exit Sub
        Else
            n_tsuki = cmbtsuki3.Text
        End If
        n_hi = n_nen & n_tsuki
        
        

    Dim udtBrowseInfo As BROWSEINFO
    Dim lngIDList     As Long
    Dim strPath       As String
    Dim shiteppp As String
On Error GoTo errcsv
    '[フォルダの参照]ダイアログ情報を設定
    With udtBrowseInfo
        .hWndOwner = Me.hwnd
        .pidlRoot = CSIDL_NETWORK
        .pszDisplayName = String(MAX_PATH2, vbNullChar)
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
On Error GoTo 0
On Error GoTo errcsv2

    '[フォルダの参照]ダイアログを表示
    lngIDList = SHBrowseForFolder(udtBrowseInfo)
    If lngIDList <> 0 Then
        strPath = String(MAX_PATH, vbNullChar)
        'ITEMIDLISTをパスに変換
        SHGetPathFromIDList lngIDList, strPath
        If Left(strPath, 1) = vbNullChar Then
            'パスが空の場合はコンピュータ名を取得
            With udtBrowseInfo
                strPath = "\\" & Left(.pszDisplayName, _
                                      InStr(.pszDisplayName, _
                                            vbNullChar) - 1)
            End With
        Else
            strPath = Left(strPath, _
                           InStr(strPath, _
                                 vbNullChar) - 1)
        End If
        '結果をラベルに表示
       shiteppp = strPath
        'ITEMIDLISTを解放
        CoTaskMemFree lngIDList
    Else
        Exit Sub
    End If
On Error GoTo 0




    If shiteppp = "" Then
        MsgBox "保存先を指定してから実行してください。"
        Exit Sub
    End If
    shiteppp = shiteppp & "\請求データ詳細_" & n_hi & "_" & Format(Date, "yyyymmdd") & "-" & Format(Time, "hhmmss") & ".csv"
    
    csv_csv6 shiteppp, n_hi
    
Exit Sub
errcsv:
        MsgBox "ダイアログ設定エラー"
        Exit Sub
errcsv2:
        MsgBox "パス選択エラー:" & strPath
        Exit Sub
        
      
End Sub

Private Sub Picture15_Click()
Dim n_nen As String, n_tsuki As String, n_hi As String
        If cmbnen4.ListIndex = -1 Then
            MsgBox "抽出したい年を選択してください"
            Exit Sub
        Else
            n_nen = cmbnen4.Text
        End If
        If cmbtsuki4.ListIndex = -1 Then
            MsgBox "抽出したい月を選択してください"
            Exit Sub
        Else
            n_tsuki = cmbtsuki4.Text
        End If
        n_hi = n_nen & n_tsuki
        
        

    Dim udtBrowseInfo As BROWSEINFO
    Dim lngIDList     As Long
    Dim strPath       As String
    Dim shiteppp As String
On Error GoTo errcsv
    '[フォルダの参照]ダイアログ情報を設定
    With udtBrowseInfo
        .hWndOwner = Me.hwnd
        .pidlRoot = CSIDL_NETWORK
        .pszDisplayName = String(MAX_PATH2, vbNullChar)
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
On Error GoTo 0
On Error GoTo errcsv2

    '[フォルダの参照]ダイアログを表示
    lngIDList = SHBrowseForFolder(udtBrowseInfo)
    If lngIDList <> 0 Then
        strPath = String(MAX_PATH, vbNullChar)
        'ITEMIDLISTをパスに変換
        SHGetPathFromIDList lngIDList, strPath
        If Left(strPath, 1) = vbNullChar Then
            'パスが空の場合はコンピュータ名を取得
            With udtBrowseInfo
                strPath = "\\" & Left(.pszDisplayName, _
                                      InStr(.pszDisplayName, _
                                            vbNullChar) - 1)
            End With
        Else
            strPath = Left(strPath, _
                           InStr(strPath, _
                                 vbNullChar) - 1)
        End If
        '結果をラベルに表示
       shiteppp = strPath
        'ITEMIDLISTを解放
        CoTaskMemFree lngIDList
    Else
        Exit Sub
    End If
On Error GoTo 0




    If shiteppp = "" Then
        MsgBox "保存先を指定してから実行してください。"
        Exit Sub
    End If
    shiteppp = shiteppp & "\入金データ詳細_" & n_hi & "_" & Format(Date, "yyyymmdd") & "-" & Format(Time, "hhmmss") & ".csv"
    
    csv_csv7 shiteppp, n_hi
    
Exit Sub
errcsv:
        MsgBox "ダイアログ設定エラー"
        Exit Sub
errcsv2:
        MsgBox "パス選択エラー:" & strPath
        Exit Sub
        
End Sub

Private Sub Picture16_Click()
Unload Me
frmshukei.Show 1

End Sub

Private Sub Picture17_Click()
Dim kakishu As String, kakikubun As String
Dim itmX As ListItem

Dim s_kensuu As Integer
s_kensuu = 0

Dim n_sousuu As Long

Dim kazekeisan(2) As Double

 ReDim csv_data_cl(3, 0)
csv_data_cl_count = 0

    Dim n_nen As String, n_tsuki As String, n_hi As String
        If cmbnen5.ListIndex = -1 Then
            MsgBox "表示検索したい年を選択してください"
            Exit Sub
        Else
            n_nen = cmbnen5.Text
        End If
        If cmbtsuki5.ListIndex = -1 Then
            MsgBox "表示検索したい月を選択してください"
            Exit Sub
        Else
            n_tsuki = cmbtsuki5.Text
        End If
        n_hi = n_nen & n_tsuki
        
        If kidou_cl = False Then
            MsgBox "データベースファイルを開けませんでした。"
            Exit Sub
        End If
        sql2 = "SELECT *" & _
            " FROM cltbl" & _
            " WHERE  nentsuki='" & n_hi & "' order by kojinid,tel,kakin"

        If FcSQlGet_cl(rs, sql2, PrMsg) = False Then
            MsgBox "CLのデータはありません。"
            db_cl.Close
            Exit Sub
        Else
            
                With frmkekka
                kakikubun = "CL"
                kakishu = "CL請求データ(" & n_nen & "年" & n_tsuki & "月請求分）"
                
               n_sousuu = rs.RecordCount
               
               rs.MoveFirst
               frmjoukyousentaku.Hide
               FRMOMACHI.lblmsg.Caption = "一覧を準備中..."
               FRMOMACHI.Show
               DoEvents
               Screen.MousePointer = 11
    With frmkekka
            .lv1.Checkboxes = False
               .lv1.ColumnHeaders. _
               Add , , "個人情報", 2000
               
              .lv1.ColumnHeaders. _
               Add , , "お客コード", 1700
               
              .lv1.ColumnHeaders. _
               Add , , "電話番号", 1800
                      
               .lv1.ColumnHeaders. _
               Add , , "請求番号", 1500
               
               .lv1.ColumnHeaders. _
               Add , , "内訳コード", 1000
               
               .lv1.ColumnHeaders. _
               Add , , "内訳詳細", 2300
            
               .lv1.ColumnHeaders. _
               Add , , "金額", 1000
               
              .lv1.ColumnHeaders. _
               Add , , "税区分", 700
               
              .lv1.ColumnHeaders. _
               Add , , "通話区分", 700
               
              .lv1.ColumnHeaders. _
               Add , , "課金", 700
               
                .Picture2.Enabled = False
    End With
    
      
        
        Dim s_maeid As String, s_maetel As String
        s_maeid = ""
        s_maetel = ""
        
        Dim n_koko As Double, n_matomete As Double, n_zei As Double, ii As Integer, s_kojinjouhou As String
            
            ii = 0
            
               While Not rs.EOF
                        
                        ii = ii + 1
                        
                        s_kojinjouhou = CStr(rs!kojinid) & " " & name_get(CStr(rs!kojinid))
                            
                            Set itmX = .lv1.ListItems. _
                            Add(, , s_kojinjouhou)  ' kojinid フィールド。
                            
                            If s_maeid <> CStr(rs!kojinid) Then
                                s_kensuu = s_kensuu + 1
                                
                                n_koko = 0
                                n_matomete = 0
                                
                                csv_data_cl_count = csv_data_cl_count + 1
                                 ReDim Preserve csv_data_cl(3, csv_data_cl_count)
                                
                                csv_data_cl(0, csv_data_cl_count - 1) = s_kojinjouhou
                                
                                csv_data_cl(1, csv_data_cl_count - 1) = rs!tel
                            Else
                                If s_maetel <> CStr(rs!tel) Then
                                     s_kensuu = s_kensuu + 1
                                    
                                   ' n_koko = 0
                                   ' n_matomete = 0
                                    
                                    csv_data_cl_count = csv_data_cl_count + 1
                                     ReDim Preserve csv_data_cl(3, csv_data_cl_count)
                                    
                                    csv_data_cl(0, csv_data_cl_count - 1) = s_kojinjouhou
                                    
                                    csv_data_cl(1, csv_data_cl_count - 1) = rs!tel
                                    
                                End If
                            End If
                            s_maeid = CStr(rs!kojinid)
                            s_maetel = CStr(rs!tel)
                          
                            itmX.SubItems(1) = rs!okyukucode
                            itmX.SubItems(2) = rs!tel
                            
                            itmX.SubItems(3) = rs!seikyuuno
                            itmX.SubItems(4) = rs!uchiwakecode
                            itmX.SubItems(5) = rs!uchiwakemeishou
                            itmX.SubItems(6) = CDbl(rs!kingaku) / 100
                            
                            n_koko = CDbl(rs!kingaku) ' / 100
                            
                            
                            
                            n_zei = 0
                             Select Case rs!zeikubun
                                Case "F"
                                    itmX.SubItems(7) = "不課税"
                                Case "H"
                                    itmX.SubItems(7) = "非課税"
                                Case "M"
                                    itmX.SubItems(7) = "免税"
                                Case "S"
                                    itmX.SubItems(7) = "外税"
                                    n_zei = CDbl(rs!kingaku) / 10 '00
                                Case "U"
                                    itmX.SubItems(7) = "内税"
                                Case "Z"
                                    itmX.SubItems(7) = "その他"
                                Case Else
                                    itmX.SubItems(7) = "エラー"
                                End Select
                             
                             
                             If Trim(rs!tsuuwakubun) = "" Then
                                itmX.SubItems(8) = "非通話"
                            Else
                                itmX.SubItems(8) = "通話"
                            End If
                            
                            If Trim(rs!kakin) = "0" Then
                                itmX.SubItems(9) = "×"
                                
                                '個々
                                If rs!uchiwakecode <> "CIPE1" Then
                                    kazekeisan(0) = kazekeisan(0) + Int((n_koko + n_zei) / 100)
                                
                                End If
                                
                                rs.MoveNext   ' 次のレコードに移動します。
                            Else
                                 itmX.SubItems(9) = "○"
                                 'まとめて
                                 If n_sousuu = ii Then
                                    '最終行
                                    If rs!uchiwakecode <> "CIPE1" Then
                                        n_matomete = n_matomete + n_koko
                                    End If
                                    n_zei = n_matomete / 10
                                    kazekeisan(1) = kazekeisan(1) + Int((n_matomete + n_zei) / 100)
                                    
                                    rs.MoveNext   ' 次のレコードに移動します。
                                 Else
                                    If rs!uchiwakecode <> "CIPE1" Then
                                        n_matomete = n_matomete + n_koko
                                    End If
                                    
                                    rs.MoveNext   ' 次のレコードに移動します。
                                    
                                    If s_maeid <> CStr(rs!kojinid) Then
                                        n_zei = n_matomete / 10
                                        kazekeisan(1) = kazekeisan(1) + Int((n_matomete + n_zei) / 100)
                                    End If
                                 End If
                            End If
                                             
                 
                            
                            
                           
                                    
                   'rs.MoveNext   ' 次のレコードに移動します。
                Wend
                rs.Close
                Screen.MousePointer = 0
                kakishu = kakishu & Space(2) & "総請求額　" & CStr(kazekeisan(0) + kazekeisan(1)) & "円"
                .lblkubun.Caption = kakikubun
                .lblshu.Caption = kakishu
                .lblkensuu.Caption = Str(s_kensuu)
                .lblseikyuukingaku.Caption = CStr(kazekeisan(0) + kazekeisan(1))
                Unload FRMOMACHI
                Set FRMOMACHI = Nothing
                db_cl.Close
                
                .Picture2.Enabled = True
                
                .Show 1
                End With
                
        End If
    Exit Sub

End Sub

Private Sub Picture18_Click()

        Unload Me


        If Dir$(sub_exe_path) = vbNullString Then
          MsgBox "ログ用実行ファイルがありません。"
         ' Exit Sub
        End If
    
        Dim sub_exe_path2 As String
  
        '自動実行　引数
        
        sub_exe_path2 = sub_exe_path & " b"
         
         
        Dim processId As Long
        
        '電卓を起動します
        processId = Shell(sub_exe_path2, vbNormalFocus)

        End

End Sub

Private Sub Picture2_Click()
    Unload Me
    frmsouken.Show 1

End Sub

Private Sub Picture3_Click()
Unload Me
    frmsouken2.Show 1

End Sub

Private Sub Picture4_Click()

' ListItem オブジェクトを追加するための変数を宣言します。
Dim itmX As ListItem

Dim kakikubun As String, kakishu As String

        If kidou2 = False Then
                MsgBox "データベースファイルを開けませんでした。"
                Exit Sub
            End If
        
        sql2 = "SELECT * FROM bankkiroku order by hidu DESC,bankid"

        If FcSQlGet2(rs, sql2, PrMsg) = False Then
            MsgBox "金融機関の記録はありません。"
            db2.Close
            Exit Sub
        Else
            
                With frmkekka
                kakikubun = "金融機関記録"
                kakishu = "フロッピーへの出力・更新の記録表示"
               ' 最終レコードまで、ListItem オブジェクトを追加します。
               .lblkensuu.Caption = rs.RecordCount
               rs.MoveFirst
               frmjoukyousentaku.Hide
               FRMOMACHI.lblmsg.Caption = "一覧を準備中..."
               FRMOMACHI.Show
               DoEvents
               Screen.MousePointer = 11
    With frmkekka
            .lv1.Checkboxes = False
               .lv1.ColumnHeaders. _
               Add , , "id", 800
               
              .lv1.ColumnHeaders. _
               Add , , "日付", 1500
               
              .lv1.ColumnHeaders. _
               Add , , "金融機関", 1800
                      
               .lv1.ColumnHeaders. _
               Add , , "請求数", 1000
               
               .lv1.ColumnHeaders. _
               Add , , "請求金額", 1500
               
               .lv1.ColumnHeaders. _
               Add , , "入金数", 1000
            
               .lv1.ColumnHeaders. _
               Add , , "入金金額", 1500
               
              .lv1.ColumnHeaders. _
               Add , , "未入数", 1000
               .lv1.ColumnHeaders. _
               Add , , "未入金額", 1500
    End With
               While Not rs.EOF
                            Set itmX = .lv1.ListItems. _
                            Add(, , CStr(rs!id))  ' kojinid フィールド。
                          
                             ' AuthorID フィールドが長さ 0 の文字列以外の場合、
                             ' SubItem 1 にこの文字列を設定します。
                            ' birth フィールドが長さ 0 の文字列以外の場合、
                            ' SubItem 2 にこの文字列を設定します。
                            If Not IsNull(rs!hidu) Then
                                itmX.SubItems(1) = Format(rs!hidu, "@@@@/@@")
                            End If
                             If Not IsNull(rs!bankid) Then
                                Select Case rs!bankid
                                Case "0142"
                                    itmX.SubItems(2) = "山梨中央銀行"
                                Case "1385"
                                    itmX.SubItems(2) = "甲府信用金庫"
                                Case "1386"
                                    itmX.SubItems(2) = "甲府商工信金"
                                Case "2375"
                                    itmX.SubItems(2) = "やまなみ信組"
                                Case "2377"
                                    itmX.SubItems(2) = "山梨県民信組"
                                Case "9900"
                                    itmX.SubItems(2) = "郵便局"
                                Case "5169"
                                    itmX.SubItems(2) = "農業協同組合"
                                End Select
                            End If
                             ' SubItem 1 にこの文字列を設定します。
                            If Not IsNull(rs!desuu) Then
                                itmX.SubItems(3) = CStr(rs!desuu)
                            End If
                            ' birth フィールドが長さ 0 の文字列以外の場合、
                            ' SubItem 2 にこの文字列を設定します。
                            If Not IsNull(rs!dekin) Then
                                itmX.SubItems(4) = rs!dekin
                            End If
                                                
                             If Not IsNull(rs!nyusuu) Then
                                itmX.SubItems(5) = rs!nyusuu
                            End If
                           
                             If Not IsNull(rs!nyukin) Then
                                itmX.SubItems(6) = rs!nyukin
                            End If
                           
                             If Not IsNull(rs!fusuu) Then
                                itmX.SubItems(7) = rs!fusuu
                            End If
                           
                             If Not IsNull(rs!fukin) Then
                                itmX.SubItems(8) = rs!fukin
                            End If
                            
                   rs.MoveNext   ' 次のレコードに移動します。
                Wend
                rs.Close
                Screen.MousePointer = 0
                .lblkubun.Caption = kakikubun
                .lblshu.Caption = kakishu
                Unload FRMOMACHI
                Set FRMOMACHI = Nothing
                db2.Close
                .Show 1
                End With
                
        End If
    Exit Sub
End Sub

Private Sub Picture5_Click()
Dim kakikubun As String
Dim goukeivoip As Long, kakishu As String
Dim itmX As ListItem

    Dim n_nen As String, n_tsuki As String, n_hi As String
        If cmbnen.ListIndex = -1 Then
            MsgBox "表示検索したい年を選択してください"
            Exit Sub
        Else
            n_nen = cmbnen.Text
        End If
        If cmbtsuki.ListIndex = -1 Then
            MsgBox "表示検索したい月を選択してください"
            Exit Sub
        Else
            n_tsuki = cmbtsuki.Text
        End If
        n_hi = n_nen & n_tsuki
        If kidou_voip = False Then
                MsgBox "データベースファイルを開けませんでした。"
                Exit Sub
            End If
        goukeivoip = 0
        sql2 = "SELECT * FROM seikyuu where nentsuki ='" & n_hi & "' order by seikyuid"

        If FcSQlGet_voip(rs, sql2, PrMsg) = False Then
            MsgBox "VoIPのデータはありません。"
            db_voip.Close
            Exit Sub
        Else
            
                With frmkekka
                kakikubun = "VoIP"
                kakishu = "VoIP請求データ(" & n_nen & "年" & n_tsuki & "月請求分）"
               ' 最終レコードまで、ListItem オブジェクトを追加します。
               .lblkensuu.Caption = rs.RecordCount
               rs.MoveFirst
               frmjoukyousentaku.Hide
               FRMOMACHI.lblmsg.Caption = "一覧を準備中..."
               FRMOMACHI.Show
               DoEvents
               Screen.MousePointer = 11
    With frmkekka
            .lv1.Checkboxes = False
               .lv1.ColumnHeaders. _
               Add , , "VoIPid", 1200
               
              .lv1.ColumnHeaders. _
               Add , , "日付", 1500
               
              .lv1.ColumnHeaders. _
               Add , , "コード", 1800
                      
               .lv1.ColumnHeaders. _
               Add , , "開始日", 1500
               
               .lv1.ColumnHeaders. _
               Add , , "終了日", 1500
               
               .lv1.ColumnHeaders. _
               Add , , "通話料金", 1500
            
               .lv1.ColumnHeaders. _
               Add , , "税額", 1000
               
              .lv1.ColumnHeaders. _
               Add , , "請求期限", 1500
    End With
               While Not rs.EOF
                            Set itmX = .lv1.ListItems. _
                            Add(, , CStr(rs!seikyuid))  ' kojinid フィールド。
                          
                             ' AuthorID フィールドが長さ 0 の文字列以外の場合、
                             ' SubItem 1 にこの文字列を設定します。
                            ' birth フィールドが長さ 0 の文字列以外の場合、
                            ' SubItem 2 にこの文字列を設定します。
                            If Not IsNull(rs!nentsuki) Then
                                itmX.SubItems(1) = Format(rs!nentsuki, "@@@@/@@")
                            End If
                            itmX.SubItems(2) = rs!code
                             ' SubItem 1 にこの文字列を設定します。
                            If Not IsNull(rs!hajime) Then
                                itmX.SubItems(3) = Format(rs!hajime, "@@@@/@@/@@")
                            End If
                            ' birth フィールドが長さ 0 の文字列以外の場合、
                            ' SubItem 2 にこの文字列を設定します。
                            If Not IsNull(rs!owari) Then
                                itmX.SubItems(4) = Format(rs!owari, "@@@@/@@/@@")
                            End If
                                                
                             If Not IsNull(rs!seikyuugaku) Then
                                itmX.SubItems(5) = rs!seikyuugaku
                                goukeivoip = goukeivoip + CLng(rs!seikyuugaku)
                            End If
                           
                             If Not IsNull(rs!zei) Then
                                itmX.SubItems(6) = rs!zei
                            End If
                           
                             If Not IsNull(rs!hi) Then
                                itmX.SubItems(7) = Format(rs!hi, "@@@@/@@/@@")
                            End If
                           
                            
                   rs.MoveNext   ' 次のレコードに移動します。
                Wend
                rs.Close
                Screen.MousePointer = 0
                kakishu = kakishu & Space(2) & "総請求額　" & goukeivoip & "円"
                .lblkubun.Caption = kakikubun
                .lblshu.Caption = kakishu
                Unload FRMOMACHI
                Set FRMOMACHI = Nothing
                db_voip.Close
                .Show 1
                End With
                
        End If
    Exit Sub

End Sub

Private Sub Picture6_Click()
'NHK集計
    nhk_furikoulist

End Sub

Private Sub Picture7_Click()
Unload Me
frmsouken4.Show 1
End Sub

Sub nhk_furikoulist()
Dim sql_furikou As String, rs_furikou As Recordset, n_furikou As String, t_furikou As String
Dim hi_furikou As String, furikousuu As Long, furicount As Long, rouhounai As Integer
Dim sql_fu As String, rs_fu As Recordset, fufu As Integer, fufu2 As Integer, fudata(8, 11) As Long
Dim fy As Integer, fx As Integer

With frmjoukyousentaku
If .combzennen.ListIndex = -1 Then
    MsgBox "表示検索したい年を選択してください"
    Exit Sub
Else
    n_furikou = .combzennen.Text
End If
If .combzentsuki2.ListIndex = -1 Then
    MsgBox "表示検索したい月を選択してください"
    Exit Sub
Else
    t_furikou = .combzentsuki2.Text
End If
If kidou4 = False Then
    Exit Sub
End If
End With
rouhounai = 0
 With frmfurikou.gridfurikou2
        hi_furikou = n_furikou & t_furikou & "00"
        sql_furikou = "SELECT nhk_k.seikyuubi, nhk_k.keiyakushurui, nhk_k.suteitasu," & _
                        "Sum(nhk_k.seikyuukingaku) AS nhkgoukin, Count(nhk_k.count) AS nhkgousuu" & _
                    " From nhk_k" & _
                    " GROUP BY nhk_k.seikyuubi, nhk_k.keiyakushurui, nhk_k.suteitasu" & _
                    " HAVING (((nhk_k.seikyuubi) Like '" & hi_furikou & "'));"
        If FcSQlGet4(rs_furikou, sql_furikou, PrMsg) > 0 Then
            furikousuu = rs_furikou.RecordCount
            rs_furikou.MoveFirst
            .Cols = 8
            .Rows = 12
            For fufu = 1 To 11 '10
                For fufu2 = 2 To 7
                    If fufu <> 10 Then  '９
                        .Row = fufu
                        .Col = fufu2
                        .Text = "0"
                     End If
                Next fufu2
            Next fufu

            .ColWidth(0) = 1000
            .ColWidth(1) = 1300
            .ColWidth(2) = 800
            .ColWidth(3) = 1200
            .ColWidth(4) = 800
            .ColWidth(5) = 1200
            .ColWidth(6) = 800
            .ColWidth(7) = 1200
            .Row = 0
            .Col = 0
            .Text = "年月"
            .Col = 1
            .Text = "種類"
            .Col = 2
            .Text = "請求数"
            .Col = 3
            .Text = "請求額"
            .Col = 4
            .Text = "入金数"
            .Col = 5
            .Text = "入金額"
            .Col = 6
            .Text = "未納数"
            .Col = 7
            .Text = "未納額"
            .Row = 1
            .Col = 0
            .Text = n_furikou & "/" & t_furikou
            .Col = 1
            .Text = "１ヶ月"
            .Row = 2
            .Col = 1
            .Text = "２ヶ月"
            .Row = 3
            .Col = 1
            .Text = "３ヶ月"
            .Row = 4
            .Col = 1
            .Text = "３ヶ月"
            .Row = 5
            .Col = 1
            .Text = "４ヶ月"
            .Row = 6
            .Col = 1
            .Text = "５ヶ月"
            .Row = 7
            .Col = 1
            .Text = "６ヶ月"
            .Row = 8
            .Col = 1
            .Text = "６ヶ月ハンデ"
            .Row = 9
            .Col = 1
            .Text = "６ヶ月家族割"
            .Row = 11
            .Col = 1
            .Text = "総合計"
            Do Until rs_furikou.EOF
                Select Case rs_furikou!keiyakushurui
                Case "0"
                    .Row = 1
                    fy = 0
                Case "1"
                    .Row = 2
                    fy = 1
                Case "2"
                    .Row = 3
                    fy = 2
                Case "3"
                    .Row = 4
                    fy = 3
                Case "4"
                    .Row = 5
                    fy = 4
                Case "5"
                    .Row = 6
                    fy = 5
                Case "6"
                    .Row = 7
                    fy = 6
                Case "7"
                    .Row = 8
                    fy = 7
                End Select
                If rs_furikou!suteitasu = "2" Then
                    .Col = 4
                    fx = 4
                    .Text = rs_furikou!nhkgousuu
                    fudata(fx, fy + 2) = CLng(rs_furikou!nhkgousuu)
                    .Col = 5
                    fx = 5
                    .Text = rs_furikou!nhkgoukin
                    fudata(fx, fy + 2) = CLng(rs_furikou!nhkgoukin)
                ElseIf rs_furikou!suteitasu = "3" Then
                    .Col = 2
                    fx = 6
                    .Text = rs_furikou!nhkgousuu
                    fudata(fx, fy + 2) = CLng(rs_furikou!nhkgousuu)
                    .Col = 3
                    fx = 7
                    .Text = rs_furikou!nhkgoukin
                    fudata(fx, fy + 2) = CLng(rs_furikou!nhkgoukin)
                End If
                rs_furikou.MoveNext
            Loop
            rs_furikou.Close
            For fufu = 1 To 9
                fudata(2, fufu) = fudata(4, fufu) + fudata(6, fufu)
                fudata(3, fufu) = fudata(5, fufu) + fudata(7, fufu)
                .Row = fufu
                .Col = 2
                .Text = fudata(2, fufu)
                .Col = 3
                .Text = fudata(3, fufu)
            Next fufu
            For fufu = 1 To 9
                For fufu2 = 2 To 7
                    .Row = fufu
                    .Col = fufu2
                    .Text = fudata(fufu2, fufu)
                    fudata(fufu2, 11) = fudata(fufu2, 11) + CLng(fudata(fufu2, fufu))
                Next fufu2
            Next fufu
            .Row = 11
            For fufu2 = 2 To 7
                .Col = fufu2
                .Text = fudata(fufu2, 11)
            Next fufu2
            rouhounai = 1
        End If
  End With


 With frmfurikou.gridfurikou
hi_furikou = n_furikou & t_furikou & "00"
sql_furikou = "select*from nhk_k where seikyuubi ='" & hi_furikou & "' and suteitasu='3'"
If FcSQlGet4(rs_furikou, sql_furikou, PrMsg) > 0 Then
    furikousuu = rs_furikou.RecordCount
    rs_furikou.MoveFirst
        .Cols = 5
        .Rows = furikousuu + 1
        .ColWidth(0) = 1000
        .ColWidth(1) = 2000
        .ColWidth(2) = 1400
        .ColWidth(3) = 1300
        .ColWidth(4) = 1600
        .Row = 0
        .Col = 0
        .Text = "ＩＤ"
        .Col = 1
        .Text = "氏名"
        .Col = 2
        .Text = "請求月"
        .Col = 3
        .Text = "種類"
        .Col = 4
        .Text = "金額"
    furicount = 1
    Do Until rs_furikou.EOF
        .Row = furicount
        .Col = 0
        .Text = rs_furikou!kojinid
        .Col = 1
        sql_fu = "select name from kojin where kojinid ='" & rs_furikou!kojinid & "'"
        If FcSQlGet(rs_fu, sql_fu, PrMsg) = True Then
            .Text = rs_fu!Name
            rs_fu.Close
        Else
            .Text = "error"
        End If
        .Col = 2
        .Text = Format(Mid(rs_furikou!seikyuubi, 1, 6), "@@@@/@@")
        .Col = 3
        .CellAlignment = flexAlignRightCenter
        Select Case rs_furikou!keiyakushurui
        Case "0"
            .Text = "１ヶ月"
        Case "1"
            .Text = "２ヶ月"
        Case "2"
            .Text = "３ヶ月"
        Case "3"
            .Text = "４ヶ月"
        Case "4"
            .Text = "５ヶ月"
        Case "5"
            .Text = "６ヶ月"
        Case "6"
            .Text = "６ヶ月ハンデ"
        Case "7"
            .Text = "６ヶ月家族割"
        End Select
        .Col = 4
        .CellAlignment = flexAlignRightCenter
        .Text = rs_furikou!seikyuukingaku
        furicount = furicount + 1
        rs_furikou.MoveNext
    Loop
    rs_furikou.Close

    rouhounai = 1
End If
db4.Close
 If rouhounai = 0 Then
    MsgBox "該当するデータがありません"
    Exit Sub
 End If
    frmfurikou.Show 1


  End With

End Sub

Private Sub Picture8_Click()

    Dim udtBrowseInfo As BROWSEINFO
    Dim lngIDList     As Long
    Dim strPath       As String
    Dim shiteppp As String
On Error GoTo errcsv
    '[フォルダの参照]ダイアログ情報を設定
    With udtBrowseInfo
        .hWndOwner = Me.hwnd
        .pidlRoot = CSIDL_NETWORK
        .pszDisplayName = String(MAX_PATH2, vbNullChar)
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
On Error GoTo 0
On Error GoTo errcsv2

    '[フォルダの参照]ダイアログを表示
    lngIDList = SHBrowseForFolder(udtBrowseInfo)
    If lngIDList <> 0 Then
        strPath = String(MAX_PATH, vbNullChar)
        'ITEMIDLISTをパスに変換
        SHGetPathFromIDList lngIDList, strPath
        If Left(strPath, 1) = vbNullChar Then
            'パスが空の場合はコンピュータ名を取得
            With udtBrowseInfo
                strPath = "\\" & Left(.pszDisplayName, _
                                      InStr(.pszDisplayName, _
                                            vbNullChar) - 1)
            End With
        Else
            strPath = Left(strPath, _
                           InStr(strPath, _
                                 vbNullChar) - 1)
        End If
        '結果をラベルに表示
       shiteppp = strPath
        'ITEMIDLISTを解放
        CoTaskMemFree lngIDList
    Else
        Exit Sub
    End If
On Error GoTo 0




    If shiteppp = "" Then
        MsgBox "保存先を指定してから実行してください。"
        Exit Sub
    End If
    shiteppp = shiteppp & "\契約状況_" & Format(Date, "yyyymmdd") & "-" & Format(Time, "hhmmss") & ".csv"
    csv_csv2 (shiteppp)
Exit Sub
errcsv:
        MsgBox "ダイアログ設定エラー"
        Exit Sub
errcsv2:
        MsgBox "パス選択エラー:" & strPath
        Exit Sub


End Sub

Private Sub Picture9_Click()

    If soukatsu(0, "startup", 12, "", 1, 1) = False Then
            MsgBox "設定の読み込みに失敗しました。"
        db.Close
        End
    End If
        
    
    If reg = "0" Then
        MsgBox "未実行です。"
    Else
        MsgBox "前回の０更新実行日は、" & Format(reg, "@@@@/@@/@@") & " です。"
    End If

End Sub
