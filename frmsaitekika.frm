VERSION 5.00
Begin VB.Form frmsaitekika 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "Tools"
   ClientHeight    =   11475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11475
   ScaleWidth      =   8490
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1710
      MouseIcon       =   "frmsaitekika.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmsaitekika.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   44
      Top             =   10530
      Width           =   480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "予定管理実行　(指定ID)"
      Height          =   915
      Left            =   4620
      TabIndex        =   43
      Top             =   10410
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "旧NHKチェック"
      Enabled         =   0   'False
      Height          =   915
      Left            =   3060
      TabIndex        =   41
      Top             =   10410
      Width           =   1365
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   450
      MouseIcon       =   "frmsaitekika.frx":1741
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmsaitekika.frx":2040
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   28
      Top             =   10500
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   7500
      MouseIcon       =   "frmsaitekika.frx":290A
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmsaitekika.frx":3209
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   10530
      Width           =   480
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6510
      MouseIcon       =   "frmsaitekika.frx":3AD3
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmsaitekika.frx":43D2
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   10500
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "種類の選択"
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
      Height          =   10065
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.OptionButton optsaitekika 
         Caption         =   "チェック"
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
         Index           =   2
         Left            =   180
         TabIndex        =   61
         Top             =   9480
         Width           =   1215
      End
      Begin VB.TextBox txtbankid 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   2  'ｵﾌ
         Left            =   3210
         MaxLength       =   4
         TabIndex        =   59
         Top             =   8910
         Width           =   1095
      End
      Begin VB.TextBox txtshuruiid 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   2  'ｵﾌ
         Left            =   6150
         MaxLength       =   1
         TabIndex        =   57
         ToolTipText     =   "1：主契約　2：その他　3：ＮＨＫ"
         Top             =   8190
         Width           =   525
      End
      Begin VB.TextBox txtkeiyakuid 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   2  'ｵﾌ
         Left            =   6150
         MaxLength       =   2
         TabIndex        =   55
         Top             =   8550
         Width           =   1095
      End
      Begin VB.TextBox txtkingaku 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   2  'ｵﾌ
         Left            =   6150
         MaxLength       =   6
         TabIndex        =   53
         Top             =   8910
         Width           =   1095
      End
      Begin VB.TextBox txtseikyuunengetsu 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   2  'ｵﾌ
         Left            =   3210
         MaxLength       =   6
         TabIndex        =   51
         Top             =   8550
         Width           =   1095
      End
      Begin VB.TextBox txtkojinid 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   2  'ｵﾌ
         Left            =   3210
         MaxLength       =   5
         TabIndex        =   48
         Top             =   8190
         Width           =   1095
      End
      Begin VB.OptionButton optsaitekika 
         Caption         =   "詳細追加"
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
         Index           =   1
         Left            =   180
         TabIndex        =   46
         Top             =   7560
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NHKデータチェック(\EZManager_data\NHK\)"
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
         Height          =   1185
         Index           =   3
         Left            =   1440
         TabIndex        =   30
         Top             =   3570
         Width           =   6495
         Begin VB.TextBox txtid 
            Height          =   315
            Left            =   2610
            MaxLength       =   2
            TabIndex        =   34
            Top             =   300
            Width           =   465
         End
         Begin VB.TextBox txtkin 
            Height          =   315
            Index           =   0
            Left            =   4680
            MaxLength       =   2
            TabIndex        =   33
            Top             =   300
            Width           =   465
         End
         Begin VB.TextBox txtkin 
            Height          =   315
            Index           =   1
            Left            =   5250
            MaxLength       =   2
            TabIndex        =   32
            Top             =   300
            Width           =   465
         End
         Begin VB.TextBox txtname 
            Height          =   315
            Left            =   840
            MaxLength       =   2
            TabIndex        =   31
            Top             =   300
            Width           =   465
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "※金額にカンマが入っている場合は2つ入力して下さい。"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Index           =   13
            Left            =   900
            TabIndex        =   42
            Top             =   780
            Width           =   5475
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "金額"
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
            Index           =   3
            Left            =   4050
            TabIndex        =   40
            Top             =   330
            Width           =   615
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ID"
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
            Index           =   3
            Left            =   2160
            TabIndex        =   39
            Top             =   330
            Width           =   345
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "列目"
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
            Left            =   3210
            TabIndex        =   38
            Top             =   330
            Width           =   525
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "列目"
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
            Left            =   5760
            TabIndex        =   37
            Top             =   330
            Width           =   525
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "列目"
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
            Left            =   1440
            TabIndex        =   36
            Top             =   330
            Width           =   525
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "名前"
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
            Index           =   7
            Left            =   210
            TabIndex        =   35
            Top             =   330
            Width           =   465
         End
      End
      Begin VB.TextBox txtato2 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5430
         MaxLength       =   5
         TabIndex        =   22
         Top             =   6810
         Width           =   1095
      End
      Begin VB.TextBox txtato1 
         Alignment       =   1  '右揃え
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5430
         MaxLength       =   5
         TabIndex        =   20
         Top             =   6240
         Width           =   1095
      End
      Begin VB.OptionButton optsaitekika 
         Caption         =   "NHK料金変更"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   5100
         Width           =   1215
      End
      Begin VB.OptionButton optsaitekika 
         Caption         =   "NHK関係チェック"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   180
         TabIndex        =   9
         Top             =   2670
         Width           =   1215
      End
      Begin VB.OptionButton optsaitekika 
         Caption         =   "フラッグ初期化"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   180
         TabIndex        =   7
         Top             =   1575
         Width           =   1215
      End
      Begin VB.OptionButton optsaitekika 
         Caption         =   "チェック"
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
         Index           =   5
         Left            =   180
         TabIndex        =   5
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CLの請求データに誤りがないか、チェックをします。"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   22
         Left            =   1440
         TabIndex        =   62
         Top             =   9540
         Width           =   6495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "金融機関コード"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   21
         Left            =   1440
         TabIndex        =   60
         Top             =   8910
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "種類ＮＯ(1桁)"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   20
         Left            =   4470
         TabIndex        =   58
         Top             =   8190
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "契約ＮＯ（2桁）"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   4470
         TabIndex        =   56
         Top             =   8550
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "請求額"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   4470
         TabIndex        =   54
         Top             =   8910
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "請求年月（6桁）"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   1440
         TabIndex        =   52
         Top             =   8550
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "個人ＩＤ（5桁）"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   1440
         TabIndex        =   50
         Top             =   8190
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "円"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   7410
         TabIndex        =   49
         Top             =   8880
         Width           =   285
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "請求詳細のテーブルにデータを手動で強制的に追加します。"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   14
         Left            =   1440
         TabIndex        =   47
         Top             =   7440
         Width           =   6495
      End
      Begin VB.Label txtmae1 
         Alignment       =   1  '右揃え
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
         Height          =   285
         Left            =   3540
         TabIndex        =   27
         Top             =   6270
         Width           =   1065
      End
      Begin VB.Label txtmae2 
         Alignment       =   1  '右揃え
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
         Height          =   285
         Left            =   3540
         TabIndex        =   26
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Label lblsa2 
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
         Height          =   285
         Left            =   7200
         TabIndex        =   25
         Top             =   6810
         Width           =   795
      End
      Begin VB.Label lblsa1 
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
         Height          =   285
         Left            =   7200
         TabIndex        =   24
         Top             =   6240
         Width           =   795
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "円"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   6600
         TabIndex        =   23
         Top             =   6810
         Width           =   285
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "円"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   6600
         TabIndex        =   21
         Top             =   6240
         Width           =   285
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "円"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   4800
         TabIndex        =   19
         Top             =   6840
         Width           =   285
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "円"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   4800
         TabIndex        =   18
         Top             =   6270
         Width           =   285
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "差額"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   7200
         TabIndex        =   17
         Top             =   5790
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "新しい金額"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   5340
         TabIndex        =   16
         Top             =   5790
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "現在の金額"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   3510
         TabIndex        =   15
         Top             =   5790
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "家族割り・半額免除"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   14
         Top             =   6840
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "団体一括6ヶ月"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   13
         Top             =   6270
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NHKの価格を変更します。実行するタイミングは、０更新し、FDDのデータを更新した後、銀行請求までの間に行うこと"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   1440
         TabIndex        =   12
         Top             =   5010
         Width           =   6495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NHKからのデータとEZManagerとのデータの双方向チェックを実行します。カンマ区切りのデータを使用してください。"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   1440
         TabIndex        =   10
         Top             =   2610
         Width           =   6495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FDDの出力と更新処理の作業中に立てるフラッグを初期化します。"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   1440
         TabIndex        =   8
         Top             =   1530
         Width           =   6495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "新システムのデータに誤りがないか、チェックをします。不正データはその場で修正できます。"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   1440
         TabIndex        =   6
         Top             =   540
         Width           =   6495
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "メンテナンス"
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
      Left            =   1410
      TabIndex        =   45
      Top             =   11070
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "初期化"
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
      Index           =   0
      Left            =   330
      TabIndex        =   29
      Top             =   11070
      Width           =   720
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
      Left            =   7530
      TabIndex        =   4
      Top             =   11070
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "実行"
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
      Left            =   6510
      TabIndex        =   3
      Top             =   11070
      Width           =   480
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   975
      Left            =   6240
      Top             =   10410
      Width           =   1890
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   975
      Left            =   150
      Top             =   10410
      Width           =   2520
   End
End
Attribute VB_Name = "frmsaitekika"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
  Unload Me
    frmnhk2.Show 1
End Sub

Private Sub Command2_Click()
Dim s_id As String, s_yyyymm As String

s_id = InputBox("個人IDの入力", vbYesNo)

If s_id = "" Then
    Exit Sub
End If


s_yyyymm = InputBox("年月の入力(202104で202105が実行されます)", vbYesNo)

If s_yyyymm = "" Then
    Exit Sub
End If


If yotejikkou(s_yyyymm, s_id) = -1 Then
     Exit Sub
End If
MsgBox "実行完了 "

End Sub

Private Sub Form_Load()

   If soukatsu(0, "startup", 8, "", 1, 0) = False Then
            MsgBox "設定の読み込みに失敗しました。"
    End If
    If reg = "0" Then
        reg = ""
    End If
    txtmae1.Caption = reg
    
        If soukatsu(0, "startup", 13, "", 1, 0) = False Then
            MsgBox "設定の読み込みに失敗しました。"
    End If
    If reg = "0" Then
        reg = ""
    End If
    txtmae2.Caption = reg
    
    
     If kenchan = 1 Then
        optsaitekika(1).Enabled = True
     Else
        optsaitekika(1).Enabled = False
     End If
     
    
    
End Sub

Private Sub Picture1_Click()
   If rocked("off") = False Then
        MsgBox "個人データを更新中のフラッグの初期化に失敗しました。"
    Else
          MsgBox "個人データを更新中のフラッグの初期化をしました。"
    End If
End Sub

Private Sub Picture10_Click()
Dim motofile As String, sakifile As String
Dim motofile2 As String, sakifile2 As String



If optsaitekika(1).Value = True Then    'チェック2

    data_tsuia
    
    Exit Sub

ElseIf optsaitekika(5).Value = True Then    'チェック2
    Unload Me
    data_check2 (1)
    Exit Sub
ElseIf optsaitekika(2).Value = True Then    'チェック2
    Unload Me
    
    Dim snen As String, stsuki As String, s_nentsuki As String
    
    snen = CLng(Format(Date, "yyyy")) - 1
    stsuki = CInt(Format(Date, "mm")) - 1
    If Format(Date, "mm") = "01" Then
        s_nentsuki = CStr(snen) & "12"
    Else
        s_nentsuki = Format(Date, "yyyy") & Format(stsuki, "00")
    End If
     cl_tools_chk s_nentsuki
    Exit Sub
ElseIf optsaitekika(7).Value = True Then     'フラッグ初期化
    If rocked("off") = False Then
        MsgBox "フラッグの初期化に失敗しました。"
        Exit Sub
    Else
        MsgBox "フラッグの初期化しました。"
        Exit Sub
    End If

    
ElseIf optsaitekika(8).Value = True Then     'NHK

            Dim DesktopPath As String
            ' デスクトップのパスを取得
            DesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
            ' 表示
            Dim s_nhk_path As String
            s_nhk_path = DesktopPath & "\EZManager_data\NHK\"
            
            
             If MsgBox("NHKの読み込みと登録データのチェックを開始します。例）nhk_201910.csv ", vbYesNo) = vbNo Then
                    Exit Sub
            End If
            
            
            Dim s_err_path As String
            s_err_path = s_nhk_path & "nhk_log_" & Format(Date, "yyyymm") & ".txt"
            'ファイルのチェック
            s_nhk_path = s_nhk_path & "nhk_" & Format(Date, "yyyymm") & ".csv"
            
            
            If Dir(s_nhk_path) = "" Then
                    MsgBox "データファイルがありません。フォルダーを確認してください。" & s_nhk_path
                    Exit Sub
            End If
            
            Dim s_id_index As Integer
            
            If Trim(txtid.Text) = "" Then
                 MsgBox "IDの行数目を確認してください。"
                Exit Sub
            Else
                s_id_index = CInt(Trim(txtid.Text))
            End If
            
            Dim s_kin_index As Integer, s_kin_index2 As Integer
            
            If Trim(txtkin(0).Text) = "" Then
                 MsgBox "金額の行数目を確認してください。"
                Exit Sub
            Else
            
                s_kin_index = CInt(Trim(txtkin(0).Text))
                
                If Trim(txtkin(1).Text) <> "" Then
                    s_kin_index2 = CInt(Trim(txtkin(1).Text))
                End If
            End If
            
            Dim s_name_index As Integer
            
            If Trim(txtname.Text) = "" Then
                MsgBox "名前の行数目を確認してください。"
                Exit Sub
            Else
                s_name_index = CInt(Trim(txtname.Text))
            End If
            
            
            err_write2 "NHKファイルの読み込み開始・・・", s_err_path
            
            Dim GetLine As Variant
            Dim RowCount As Integer, ColCount As Integer
            Dim Spot As Integer, Position As Integer, newkojinid As String
            Dim Row As Integer, Col As Integer, kariatai
            Dim s_nhk_data() As String, s_nhk_data2() As String
               
               
               
               'ファイルからデータ抽出
               
            
               
            On Error GoTo ErrorTrap3
               Open s_nhk_path For Input Lock Read Write As #1
               ColCount = 0
               RowCount = 100  '25
                While Not EOF(1)
                    Line Input #1, GetLine
                    ColCount = ColCount + 1
                Wend
                Close #1
                
                Open s_nhk_path For Input Lock Read Write As #1
               ReDim s_nhk_data(ColCount, RowCount)
               Position = 1
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
                                s_nhk_data(Row, Col) = Mid(kariatai, 2, Len(kariatai) - 2)
                            Else
                                s_nhk_data(Row, Col) = Mid(kariatai, 2)
                            End If
                        Else
                            If Right(kariatai, 1) = """" Then
                                s_nhk_data(Row, Col) = Mid(kariatai, 1, Len(kariatai) - 1)
                            Else
                                s_nhk_data(Row, Col) = kariatai
                            End If
                        End If
                    Else
                        s_nhk_data(Row, Col) = kariatai
                    End If
                     Col = Col + 1
                     GetLine = Right$(GetLine, Len(GetLine) - Spot)
                     Spot = InStr(1, GetLine, ",")
                  Wend
                  If Len(GetLine) <> 0 Then
                        If Mid(GetLine, 1, 1) = """" Then
                            If Right(GetLine, 1) = """" Then
                                s_nhk_data(Row, Col) = Mid(GetLine, 2, Len(GetLine) - 2)
                            Else
                                s_nhk_data(Row, Col) = Mid(GetLine, 2)
                            End If
                        Else
                            If Right(GetLine, 1) = """" Then
                                s_nhk_data(Row, Col) = Mid(GetLine, 1, Len(GetLine) - 1)
                            Else
                                s_nhk_data(Row, Col) = GetLine
                            End If
                        End If
                  End If
                  Row = Row + 1
               Wend
               Close #1
            
             
            'EZManagerからデータ抽出
            
            err_write2 "EzManagerのデータの読み込み開始・・・", s_err_path
            
            Dim sql_sentaku As String, rs_sentaku2 As Recordset, datasuu2 As Integer, sususu As String
            
            sususu = 1
            
            sql_sentaku = "SELECT nhk_m.kojinid, kojin.name, kojin.furigana, kojin.chiikiid, kojin.nhkno,kojin.nhkname,kojin.nhkfuri," & _
                        "kojin.juusho1, kojin.juusho2, kojin.tel1, kojin.tel2, kojin.keitai," & _
                        "nhk_m.shurui, nhk_m.n1, nhk_m.n2, nhk_m.jikai,nhk_m.zenkai," & _
                        "nhk_m.kingaku,kojin.juutaku,kojin.cellno,kojin.bikou,kojin.setaisuu,kojin.tvsuu" & _
                        " FROM kojin RIGHT JOIN nhk_m ON kojin.kojinid = nhk_m.kojinid"
            
            If FcSQlGet(rs_sentaku2, sql_sentaku, PrMsg) = False Then
                If sususu = 1 Then
                        MsgBox "EzManagerにNHKの請求データが存在しません。"
                        Exit Sub
                End If
            Else
            On Error GoTo eee2
                datasuu2 = rs_sentaku2.RecordCount
                
                    ReDim Preserve s_nhk_data2(29, datasuu2)
                rs_sentaku2.MoveFirst
                Do Until rs_sentaku2.EOF
                        s_nhk_data2(0, sususu) = Format(sususu, "00000")
                        s_nhk_data2(1, sususu) = rs_sentaku2!kojinid
                        s_nhk_data2(2, sususu) = rs_sentaku2!Name
                        s_nhk_data2(3, sususu) = rs_sentaku2!Furigana
                        s_nhk_data2(4, sususu) = rs_sentaku2!chiikiid
                        s_nhk_data2(5, sususu) = rs_sentaku2!juusho1
                        s_nhk_data2(6, sususu) = rs_sentaku2!juusho2
                        s_nhk_data2(7, sususu) = rs_sentaku2!tel1
                        s_nhk_data2(8, sususu) = rs_sentaku2!tel2
                        s_nhk_data2(9, sususu) = rs_sentaku2!keitai
                        s_nhk_data2(10, sususu) = rs_sentaku2!shurui 'NHK_m
                        Select Case rs_sentaku2!shurui 'NHK_m
                        Case "0"
                            s_nhk_data2(11, sususu) = "口座振替１ヶ月"
                        Case "1"
                            s_nhk_data2(11, sususu) = "口座振替２ヶ月"
                        Case "2"
                            s_nhk_data2(11, sususu) = "口座振替３ヶ月"
                        Case "3"
                            s_nhk_data2(11, sususu) = "口座振替４ヶ月"
                        Case "4"
                            s_nhk_data2(11, sususu) = "口座振替５ヶ月"
                        Case "5"
                            s_nhk_data2(11, sususu) = "口座振替６ヶ月"
                        Case "6"
                            s_nhk_data2(11, sususu) = "口座振替６ヶ月ハンデ"
                        Case "7"
                            s_nhk_data2(11, sususu) = "口座振替６ヶ月家族割"
                        End Select
                        s_nhk_data2(12, sususu) = rs_sentaku2!n1     'NHK_m
                        s_nhk_data2(13, sususu) = rs_sentaku2!n2     'NHK_m
                        s_nhk_data2(14, sususu) = rs_sentaku2!kingaku  'NHK_m
                        s_nhk_data2(15, sususu) = rs_sentaku2!juutaku
                        s_nhk_data2(16, sususu) = rs_sentaku2!cellno
                        s_nhk_data2(17, sususu) = rs_sentaku2!bikou
                        s_nhk_data2(18, sususu) = rs_sentaku2!zenkai  'NHK_m
                        s_nhk_data2(19, sususu) = rs_sentaku2!jikai   'NHK_m
                        s_nhk_data2(20, sususu) = rs_sentaku2!setaisuu
                        s_nhk_data2(21, sususu) = rs_sentaku2!tvsuu
                            s_nhk_data2(22, sususu) = rs_sentaku2!nhkno
                            s_nhk_data2(23, sususu) = rs_sentaku2!nhkname
                            s_nhk_data2(24, sususu) = rs_sentaku2!nhkfuri
                            s_nhk_data2(25, sususu) = ""
                            s_nhk_data2(26, sususu) = ""
                            s_nhk_data2(27, sususu) = ""
                            s_nhk_data2(28, sususu) = ""
                    sususu = sususu + 1
                    rs_sentaku2.MoveNext
                Loop
                rs_sentaku2.Close
                On Error GoTo 0
            End If
            
            
            
            
            '比較
            
            Dim s_kokyakuid As String, s_nkno As String, s_kingaku As String, s_shurui As String, s_shurui2 As String, n_kingaku As Long
            Dim s_name As String, ii As Integer
            Dim s_st As Integer, s_msg As String
            
            err_write2 "NHKデータの比較開始・・・", s_err_path
            
                'NHKのデータがあるもので、EZManagerにないもの
                For i = 0 To ColCount - 1
                    s_st = 0
                    
                    If Len(s_nhk_data(i, s_id_index)) <= 5 Then
                        s_kokyakuid = Format$(s_nhk_data(i, s_id_index), "0000#")
                    Else
                        s_kokyakuid = Right(s_nhk_data(i, s_id_index), 5)
                    End If
                    s_name = Trim(s_nhk_data(i, s_name_index))
                     s_nkno = s_nhk_data(i, 6)
                     
                     If s_kin_index2 = 0 Then
                        n_kingaku = CLng(s_nhk_data(i, s_kin_index))
                        s_kingaku = CStr(n_kingaku)
                     Else
                        n_kingaku = CLng(s_nhk_data(i, s_kin_index) & s_nhk_data(i, s_kin_index2))
                        s_kingaku = CStr(n_kingaku)
                     End If
                    s_shurui = s_nhk_data(i, 8)
                    s_shurui2 = s_nhk_data(i, 9)
                    
                    For ii = 0 To sususu - 1
                        If s_kokyakuid = s_nhk_data2(1, ii) Then '顧客IDで検索
                             If s_kingaku = s_nhk_data2(14, ii) Then '金額で検索
                                    s_st = 1
                                    Exit For
                             End If
                        End If
                    Next
                    
                    If s_st = 0 Then
                        s_msg = s_kokyakuid & " " & s_name & "  のNHKデータが、EzManagerにありません。"
                        err_write2 "エラー：" & s_msg, s_err_path
                        ' MsgBox s_msg
                    End If
                Next
                
            err_write2 "EzManagerデータの比較開始・・・", s_err_path
            
            
                 'EZManagerのデータがあるもので、NHKにないもの
                 
                 Dim s_kokyakuid2 As String, n_kingaku2 As Long, s_kingaku2 As String
                 
                 For ii = 0 To sususu - 1
                    s_st = 0
                    s_kokyakuid = s_nhk_data2(1, ii)
                    s_name = s_nhk_data2(2, ii)
                    ' s_nkno = s_nhk_data(i, 6)
                     s_kingaku = s_nhk_data2(14, ii)
                   ' s_shurui = s_nhk_data(i, 8)
                   ' s_shurui2 = s_nhk_data(i, 9)
                   
                    If s_kokyakuid <> "" Then
                       For i = 0 To ColCount - 1
                            If Len(s_nhk_data(i, s_id_index)) <= 5 Then
                                s_kokyakuid2 = Format$(s_nhk_data(i, s_id_index), "0000#")
                            Else
                                s_kokyakuid2 = Right(s_nhk_data(i, s_id_index), 5)
                            End If
                          
                             If s_kin_index2 = 0 Then
                                n_kingaku2 = CLng(s_nhk_data(i, s_kin_index))
                                s_kingaku2 = CStr(n_kingaku2)
                             Else
                                n_kingaku2 = CLng(s_nhk_data(i, s_kin_index) & s_nhk_data(i, s_kin_index2))
                                s_kingaku2 = CStr(n_kingaku2)
                             End If
                             
                           If s_kokyakuid = s_kokyakuid2 Then '顧客IDで検索
                                If s_kingaku = s_kingaku2 Then '金額で検索
                                       s_st = 1
                                       Exit For
                                End If
                           End If
                       Next
                    
                       If s_st = 0 Then
                           s_msg = s_kokyakuid & " " & s_name & " のEzManagerデータが、NHKにありません。"
                           err_write2 "エラー２：" & s_msg, s_err_path
                           ' MsgBox s_msg
                       End If
                    End If
                 Next
                 
                 
                 
                 
            
            err_write2 "各データの比較終了・・・", s_err_path
            
            
            Dim s_nhk_shuukei(3) As Long
            Dim n_nhk_shuukei(3) As Double
            
            Dim s_ez_shuukei(3) As Long
            Dim n_ez_shuukei(3) As Double
            
            '集計
            Dim s_msg2 As String
            
             'NHKのデータ
                For i = 0 To ColCount - 1
                     
                     If s_kin_index2 = 0 Then
                        n_kingaku = CLng(s_nhk_data(i, s_kin_index))
                        s_kingaku = CStr(n_kingaku)
                     Else
                        n_kingaku = CLng(s_nhk_data(i, s_kin_index) & s_nhk_data(i, s_kin_index2))
                        s_kingaku = CStr(n_kingaku)
                     End If
                   
                    Select Case s_kingaku
                    Case nhk_ryoukin_tsuujou
                        s_nhk_shuukei(0) = s_nhk_shuukei(0) + 1
                    Case nhk_ryoukin_hangaku
                        s_nhk_shuukei(1) = s_nhk_shuukei(1) + 1
                    End Select
                    s_nhk_shuukei(2) = s_nhk_shuukei(2) + 1
                Next
                n_nhk_shuukei(0) = s_nhk_shuukei(0) * nhk_ryoukin_tsuujou
                n_nhk_shuukei(1) = s_nhk_shuukei(1) * nhk_ryoukin_hangaku
                n_nhk_shuukei(2) = n_nhk_shuukei(0) + n_nhk_shuukei(1)
            
                s_msg2 = CStr(nhk_ryoukin_tsuujou) & "円　" & CStr(s_nhk_shuukei(0)) & "件　" & CStr(n_nhk_shuukei(0)) & "円"
                s_msg2 = s_msg2 & "  " & CStr(nhk_ryoukin_hangaku) & "円　" & CStr(s_nhk_shuukei(1)) & "件　" & CStr(n_nhk_shuukei(1)) & "円"
                s_msg2 = s_msg2 & "  総計　" & CStr(s_nhk_shuukei(2)) & "件　" & CStr(n_nhk_shuukei(2)) & "円"
                err_write2 "NHK集計：" & s_msg2, s_err_path
            
                
                 'EZManagerのデータ
                For ii = 0 To sususu - 1
                    
                     s_kingaku = s_nhk_data2(14, ii)
                  
                     Select Case s_kingaku
                    Case nhk_ryoukin_tsuujou
                        s_ez_shuukei(0) = s_ez_shuukei(0) + 1
                    Case nhk_ryoukin_hangaku
                        s_ez_shuukei(1) = s_ez_shuukei(1) + 1
                    End Select
                    s_ez_shuukei(2) = s_ez_shuukei(2) + 1
                 Next
                 n_ez_shuukei(0) = s_ez_shuukei(0) * nhk_ryoukin_tsuujou
                n_ez_shuukei(1) = s_ez_shuukei(1) * nhk_ryoukin_hangaku
                n_ez_shuukei(2) = n_ez_shuukei(0) + n_ez_shuukei(1)
            
                s_msg2 = CStr(nhk_ryoukin_tsuujou) & "円　" & CStr(s_ez_shuukei(0)) & "件　" & CStr(n_ez_shuukei(0)) & "円"
                s_msg2 = s_msg2 & "  " & CStr(nhk_ryoukin_hangaku) & "円　" & CStr(s_ez_shuukei(1)) & "件　" & CStr(n_ez_shuukei(1)) & "円"
                s_msg2 = s_msg2 & "  総計　" & CStr(s_ez_shuukei(2)) & "件　" & CStr(n_ez_shuukei(2)) & "円"
                err_write2 "EzManager集計：" & s_msg2, s_err_path
                
                
            MsgBox "チェックが終了しました。"
            
            Exit Sub



ElseIf optsaitekika(0).Value = True Then     'NHK

    Dim s_moto_1 As String
    Dim s_moto_2 As String
    
    Dim s_saki_1 As String
    Dim s_saki_2 As String
    
    Dim s_sa_1 As String
    Dim s_sa_2 As String
    
    Dim n_moto_1 As Long
    Dim n_moto_2 As Long
    
    Dim n_saki_1 As Long
    Dim n_saki_2 As Long
    
    Dim n_sa_1 As Long
    Dim n_sa_2 As Long
    
    
    s_moto_1 = Trim(txtmae1.Caption)
    If s_moto_1 = "" Then
        MsgBox "金額が入力されていません。"
        Exit Sub
    Else
        n_moto_1 = CLng(s_moto_1)
    End If
    
    s_moto_2 = Trim(txtmae2.Caption)
    If s_moto_2 = "" Then
        MsgBox "金額が入力されていません。"
        Exit Sub
    Else
        n_moto_2 = CLng(s_moto_2)
    End If
    
    s_saki_1 = Trim(txtato1.Text)
    If s_saki_1 = "" Then
        MsgBox "金額が入力されていません。"
        Exit Sub
    Else
        n_saki_1 = CLng(s_saki_1)
    End If
    
    s_saki_2 = Trim(txtato2.Text)
    If s_saki_2 = "" Then
        MsgBox "金額が入力されていません。"
        Exit Sub
    Else
        n_saki_2 = CLng(s_saki_2)
    End If
    
    
    s_sa_1 = Trim(lblsa1.Caption)
    If s_sa_1 = "" Then
        MsgBox "差額が計算されていません。"
        Exit Sub
    Else
        n_sa_1 = CLng(s_sa_1)
    End If
    
    s_sa_2 = Trim(lblsa2.Caption)
    If s_sa_2 = "" Then
        MsgBox "差額が計算されていません。"
        Exit Sub
    Else
        n_sa_2 = CLng(s_sa_2)
    End If
    
    Dim pres
    pres = MsgBox("本当に変更してよいですか？一度実行すると元に戻すことはできません。", vbYesNo)
    If pres = vbYes Then
    
      
      If nhk_gengaku(n_moto_1, n_moto_2, n_saki_1, n_saki_2, n_sa_1, n_sa_2) = 1 Then
      
      
              'NHK金額
        
                If soukatsu(1, "startup", 8, s_saki_1, 1, 0) = False Then
                        MsgBox "契約金額1の修正は終わりましたが、新しい金額の設定の書き込みに失敗しました。"
                        Exit Sub
                End If
                
                
                If soukatsu(1, "startup", 13, s_saki_2, 1, 0) = False Then
                        MsgBox "契約金額2の修正は終わりましたが、新しい金額の設定の書き込みに失敗しました。"
                        Exit Sub
                End If
                
                
                
               txtmae1.Caption = s_saki_1
               
                txtmae2.Caption = s_saki_2
                
                txtato1.Text = ""
                
                txtato2.Text = ""
                
                lblsa1.Caption = ""
                
                lblsa2.Caption = ""
                
        
            MsgBox "更新が完了しました。"
      Else
            MsgBox "更新に失敗しました。"
      End If
      
      
      
      
        
    
    
    
    Else
        MsgBox "キャンセルしました。"
    End If
    


End If



Exit Sub


ErrorTrap3:
        MsgBox "ファイルの読み込みに失敗しました。"
        Close #1
        Exit Sub

eee2:

        MsgBox "データの抽出に失敗しました。"
      
        Exit Sub
        
End Sub

Private Sub Picture13_Click()
Unload frmsaitekika
Set frmsaitekika = Nothing

End Sub

Private Sub Picture2_Click()
   'If kiidou_mode <> 0 Then
        MsgBox "このモードでは使用できません。"
        Exit Sub
    'End If
             Dim RetVal
                msg = "ＥＺ　Ｍａｎａｇｅｒの最新版への自動更新、またはメンテナンスのためにデータを転送しますか？"
            res = MsgBox(msg, vbYesNo)
            If res = vbYes Then
                mentefaile = mapa & "ezmanager_ment.exe"
                If Dir(mentefaile) = "" Then
                    MsgBox "メンテナンス用プログラムがありません。"
                    Exit Sub
                End If
                RetVal = Shell(mentefaile, 1)   ' 電卓を実行します。
                End
            End If

End Sub

Private Sub txtato1_KeyPress(KeyAscii As Integer)


If KeyAscii = 13 Then
    
    lblsa1.Caption = ""
    
    Dim s_moto_1 As String
    Dim s_moto_2 As String
    
    Dim s_saki_1 As String
    Dim s_saki_2 As String
    
    Dim s_sa_1 As String
    Dim s_sa_2 As String
    
    Dim n_moto_1 As Long
    Dim n_moto_2 As Long
    
    Dim n_saki_1 As Long
    Dim n_saki_2 As Long
    
    Dim n_sa_1 As Long
    Dim n_sa_2 As Long
    
    
    s_moto_1 = Trim(txtmae1.Caption)
    If s_moto_1 = "" Then
        Exit Sub
    Else
        n_moto_1 = CLng(s_moto_1)
    End If
    
    
    s_saki_1 = Trim(txtato1.Text)
    If s_saki_1 = "" Then
        Exit Sub
    Else
        n_saki_1 = CLng(s_saki_1)
    End If
    
    
    lblsa1.Caption = CStr(n_saki_1 - n_moto_1)
    
    
End If


End Sub

Private Sub txtato2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    lblsa2.Caption = ""
    
    Dim s_moto_1 As String
    Dim s_moto_2 As String
    
    Dim s_saki_1 As String
    Dim s_saki_2 As String
    
    Dim s_sa_1 As String
    Dim s_sa_2 As String
    
    Dim n_moto_1 As Long
    Dim n_moto_2 As Long
    
    Dim n_saki_1 As Long
    Dim n_saki_2 As Long
    
    Dim n_sa_1 As Long
    Dim n_sa_2 As Long
    
    
    s_moto_2 = Trim(txtmae2.Caption)
    If s_moto_2 = "" Then
        Exit Sub
    Else
        n_moto_2 = CLng(s_moto_2)
    End If
    
    
    s_saki_2 = Trim(txtato2.Text)
    If s_saki_2 = "" Then
        Exit Sub
    Else
        n_saki_2 = CLng(s_saki_2)
    End If
    
    
    lblsa2.Caption = CStr(n_saki_2 - n_moto_2)
    
    
End If
End Sub
