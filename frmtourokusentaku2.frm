VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmtourokusentaku2 
   Caption         =   "ó\íËä«óù"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   8850
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "ìoò^ó\íËè⁄ç◊"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4965
      Left            =   180
      TabIndex        =   2
      Top             =   210
      Width           =   8475
      Begin VB.TextBox txtbikou 
         Alignment       =   1  'âEëµÇ¶
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
         IMEMode         =   1  'µ›
         Left            =   1980
         MaxLength       =   100
         TabIndex        =   32
         Top             =   4350
         Width           =   6045
      End
      Begin VB.TextBox txtkingaku 
         Alignment       =   1  'âEëµÇ¶
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
         Left            =   1980
         TabIndex        =   31
         Top             =   2310
         Width           =   1335
      End
      Begin VB.TextBox txtkuchisuu 
         Alignment       =   1  'âEëµÇ¶
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
         Left            =   6120
         TabIndex        =   27
         Top             =   2310
         Width           =   360
      End
      Begin VB.TextBox txtkeiyakusuu 
         Alignment       =   1  'âEëµÇ¶
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
         Left            =   2010
         TabIndex        =   24
         Top             =   3600
         Width           =   360
      End
      Begin VB.ComboBox cmbkeiyakushurui 
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
         Left            =   2010
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   22
         Top             =   1050
         Width           =   5715
      End
      Begin VB.ComboBox cmbsonotakeiyaku 
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
         Left            =   2010
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   21
         Top             =   1170
         Width           =   5715
      End
      Begin VB.ComboBox cmbnen 
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
         Left            =   1980
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   4
         Top             =   3030
         Width           =   915
      End
      Begin VB.ComboBox cmbtsuki 
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
         Left            =   3240
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   3
         Top             =   3030
         Width           =   675
      End
      Begin MSComCtl2.UpDown ud1 
         Height          =   375
         Left            =   2370
         TabIndex        =   23
         Top             =   3600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtkeiyakusuu"
         BuddyDispid     =   196613
         OrigLeft        =   2280
         OrigTop         =   840
         OrigRight       =   2520
         OrigBottom      =   1215
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   6481
         TabIndex        =   28
         Top             =   2310
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtkuchisuu"
         BuddyDispid     =   196612
         OrigLeft        =   2280
         OrigTop         =   840
         OrigRight       =   2520
         OrigBottom      =   1215
         Max             =   99
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   0   'False
      End
      Begin VB.Label Label14 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0E0FF&
         Caption         =   "îıçl"
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
         Height          =   285
         Left            =   420
         TabIndex        =   33
         Top             =   4440
         Width           =   1005
      End
      Begin VB.Label lblshuuryou 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00FFFFFF&
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
         Left            =   5910
         TabIndex        =   6
         Top             =   3570
         Width           =   2055
      End
      Begin VB.Label Label13 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0E0FF&
         Caption         =   "å˚êî"
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
         Height          =   285
         Left            =   4560
         TabIndex        =   30
         Top             =   2370
         Width           =   1005
      End
      Begin VB.Label Label12 
         Caption         =   "å˚"
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
         Left            =   7230
         TabIndex        =   29
         Top             =   2430
         Width           =   255
      End
      Begin VB.Label lbltannka 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00FFFFFF&
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
         Left            =   7230
         TabIndex        =   26
         Top             =   270
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label10 
         Caption         =   "âÒ"
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
         Left            =   3120
         TabIndex        =   25
         Top             =   3720
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0E0FF&
         Caption         =   "å_ñÒ"
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
         Height          =   285
         Left            =   420
         TabIndex        =   18
         Top             =   1230
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0E0FF&
         Caption         =   "éÌóﬁID"
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
         Height          =   285
         Left            =   420
         TabIndex        =   17
         Top             =   510
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0E0FF&
         Caption         =   "å_ñÒã‡äz"
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
         Height          =   285
         Left            =   420
         TabIndex        =   16
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label Label5 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0E0FF&
         Caption         =   "ëçâÒêî"
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
         Height          =   285
         Left            =   450
         TabIndex        =   15
         Top             =   3660
         Width           =   1005
      End
      Begin VB.Label Label7 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0E0FF&
         Caption         =   "äJénîNåé"
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
         Height          =   285
         Left            =   390
         TabIndex        =   14
         Top             =   3090
         Width           =   1005
      End
      Begin VB.Label Label8 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0E0FF&
         Caption         =   "èIóπîNåé"
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
         Height          =   285
         Left            =   4620
         TabIndex        =   13
         Top             =   3630
         Width           =   1515
      End
      Begin VB.Label lblshuruiid 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00FFFFFF&
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
         Left            =   2010
         TabIndex        =   12
         Top             =   420
         Width           =   1035
      End
      Begin VB.Label lblshuruimei 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00FFFFFF&
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
         Left            =   3240
         TabIndex        =   11
         Top             =   435
         Width           =   2055
      End
      Begin VB.Label lblkeiyakuid 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00FFFFFF&
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
         Left            =   1980
         TabIndex        =   10
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label lblkeiyakumei 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00FFFFFF&
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
         Left            =   2970
         TabIndex        =   9
         Top             =   1680
         Width           =   5055
      End
      Begin VB.Label Label6 
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
         Left            =   2940
         TabIndex        =   8
         Top             =   3150
         Width           =   255
      End
      Begin VB.Label Label15 
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
         Left            =   3960
         TabIndex        =   7
         Top             =   3150
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "â~"
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
         Left            =   3600
         TabIndex        =   5
         Top             =   2460
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6420
      MouseIcon       =   "frmtourokusentaku2.frx":0000
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmtourokusentaku2.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   5460
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   7620
      MouseIcon       =   "frmtourokusentaku2.frx":11C9
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmtourokusentaku2.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   5460
      Width           =   480
   End
   Begin VB.Label Label16 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FFFFFF&
      Caption         =   "Å¶äJénîNåéÇ‚èIóπîNåéÇÃéwíËÇ≥ÇÍÇΩîNåéÇÃé¿é{ì˙ÇÕÅAÇªÇÃëOåéññÇÃçXêVéûÇ…é¿çsÇ≥ÇÍÇ‹Ç∑ÅB"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   270
      TabIndex        =   34
      Top             =   5580
      Width           =   5055
   End
   Begin VB.Label Label11 
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
      Left            =   6450
      TabIndex        =   20
      Top             =   6000
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
      Index           =   16
      Left            =   7650
      TabIndex        =   19
      Top             =   5970
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   1035
      Left            =   6000
      Top             =   5310
      Width           =   2580
   End
End
Attribute VB_Name = "frmtourokusentaku2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbkeiyakushurui_Click()
Dim k_rs As Recordset

If cmbkeiyakushurui.ListIndex <> -1 Then
    sql = "select*from shurui where shuruiid='" & Mid(cmbkeiyakushurui.Text, 1, 3) & "'"
    If FcSQlGet(k_rs, sql, PrMsg) = False Then
        MsgBox "ó\ä˙ÇπÇ éñóRÇ…ÇÊÇËÉGÉâÅ[Ç™î≠ê∂ÇµÇ‹ÇµÇΩÅB"
        Exit Sub
    End If
    With frmtourokusentaku2
        .lblkeiyakuid.Caption = k_rs!shuruiid
        .lblkeiyakumei.Caption = k_rs!setsumei
        .txtkingaku.Text = k_rs!tanka
         .txtkingaku.Enabled = False
        .lbltannka.Caption = k_rs!tanka
        .txtkuchisuu.Text = "1"
        .UpDown1.Enabled = True
        .txtkeiyakusuu.Text = "1"
        .ud1.Enabled = False
    End With
    k_rs.Close
End If

End Sub

Private Sub cmbsonotakeiyaku_Click()
Dim k_rs As Recordset

If cmbsonotakeiyaku.ListIndex <> -1 Then
     sql = "select*from sonota where sonotaid='" & Mid(cmbsonotakeiyaku.Text, 1, 3) & "'"
    If FcSQlGet(k_rs, sql, PrMsg) = False Then
        MsgBox "ó\ä˙ÇπÇ éñóRÇ…ÇÊÇËÉGÉâÅ[Ç™î≠ê∂ÇµÇ‹ÇµÇΩÅB"
        Exit Sub
    End If
    With frmtourokusentaku2
        .lblkeiyakuid.Caption = k_rs!sonotaid
        .lblkeiyakumei.Caption = k_rs!sonota
        
        If k_rs!kotei = "0" Then 'å≈íË
            If IsNull(k_rs!kingaku) Then
                .txtkingaku.Text = ""
                .lbltannka.Caption = ""
            Else
                .txtkingaku.Text = k_rs!kingaku
                .lbltannka.Caption = k_rs!kingaku
                End If
                
                .txtkingaku.Enabled = False
                
        Else
            .txtkingaku.Text = ""
            .lbltannka.Caption = ""
            
             .txtkingaku.Enabled = True
             
        End If
        
        .txtkuchisuu.Text = "1"
        .UpDown1.Enabled = False
        .txtkeiyakusuu.Text = "0"
        .ud1.Enabled = True
    End With
    k_rs.Close
End If

    
End Sub

Private Sub Form_Load()
Dim nntt As Integer, karinntt As Long

With Me
 
   For nntt = 0 To 5
        karinntt = CLng(Format(Date, "yyyy")) + nntt
        .cmbnen.AddItem karinntt
   Next nntt
   For nntt = 1 To 12
        .cmbtsuki.AddItem Format(nntt, "00")
   Next nntt
  
End With
End Sub

Private Sub Picture13_Click(Index As Integer)
    Unload frmtourokusentaku2
    Set frmtourokusentaku2 = Nothing
End Sub

Private Sub Picture4_Click()

'Ç±Ç±Ç©ÇÁ

Dim s_jisshi As String, s_jisshi2 As String


If Trim(cmbnen.Text) = "" Then
    MsgBox "îNÇ™ì¸óÕÇ≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB"
    Exit Sub
End If
If Trim(cmbtsuki.Text) = "" Then
    MsgBox "åéÇ™ì¸óÕÇ≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB"
    Exit Sub
End If

s_jisshi = Trim(cmbnen.Text) & Trim(cmbtsuki.Text)

Dim s_itsu As String
            
s_itsu = Format(Date, "yyyymm")
 
If s_jisshi < s_itsu Then
    If Trim(lblshuruiid.Caption) = "0" Then
        MsgBox "îNåéÇ™âﬂãéÇÃÇ‡ÇÃÇÕìoò^Ç≈Ç´Ç‹ÇπÇÒÅB"
        Exit Sub
    Else
        MsgBox "äJénîNåéÇ™âﬂãéÇ≈Ç∑ÅBéÊÇËàµÇ¢Ç…íçà”ÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
    End If
ElseIf s_jisshi = s_itsu Then
    MsgBox "îNåéÇ™ìØàÍÇ≈Ç∑ÅBÇOçXêVÇ∆ÉfÅ[É^çXêVÇ™èIÇÌÇ¡ÇƒÇ¢ÇÈèÍçáÇÕÅAó\íËÇÕçÏìÆÇµÇ‹ÇπÇÒÅB"

End If

Dim s_shurui As String

If Trim(lblshuruiid.Caption) = "0" Then
    'éÂå_ñÒ
    s_shurui = "0"
    
    s_jisshi2 = s_jisshi
Else
    'ÇªÇÃëºÇÃå_ñÒ
    s_shurui = "1"
    
    If Trim(lblshuuryou.Caption) = "" Then
        MsgBox "èIóπîNåéÇ™ï\é¶Ç≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB"
        Exit Sub
    Else
        s_jisshi2 = Mid(Trim(lblshuuryou.Caption), 1, 4) & Mid(Trim(lblshuuryou.Caption), 6, 2)
    End If

End If

Dim s_bikou As String

s_bikou = Trim(txtbikou.Text)
 

Dim s_kojinid As String

s_kojinid = Trim(frmtouroku2.lblid.Caption)





Dim s_kanrenid As String, s_kanrenmei As String, s_kingaku As String

s_kanrenid = Trim(lblkeiyakuid.Caption)
s_kanrenmei = Trim(lblkeiyakumei.Caption)

s_kingaku = Trim(txtkingaku.Text)

If s_kingaku = "" Then
    MsgBox "ã‡äzÇ™ì¸óÕÇ≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB"
    Exit Sub
End If

Dim s_kuchisuuu As String, s_sousuu As String

s_kuchisuuu = Trim(txtkuchisuu.Text)
If s_kuchisuuu = "" Or s_kuchisuuu = "0" Then
    MsgBox "å˚êîÇ™ï\é¶Ç≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB"
    Exit Sub
End If

s_sousuu = Trim(txtkeiyakusuu.Text)
If s_sousuu = "" Or s_sousuu = "0" Then
    MsgBox "âÒêîÇ™ï\é¶Ç≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB"
    Exit Sub
End If




If yoteikanri_touroku(s_kojinid, s_shurui, s_kanrenid, s_kanrenmei, s_kingaku, s_kuchisuuu, s_sousuu, s_jisshi, s_jisshi2, "0", s_bikou) = False Then
    MsgBox "ó\íËÇÃìoò^Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
    Exit Sub
Else
    Unload frmtourokusentaku2
    Set frmtourokusentaku2 = Nothing
    
End If




End Sub

Private Sub txtkeiyakusuu_Change()

lblshuuryou.Caption = ""

If Trim(txtkeiyakusuu.Text) = "" Or Trim(txtkeiyakusuu.Text) = "0" Then
    Exit Sub
End If

Dim n_kaisuu As Integer

n_kaisuu = CInt(Trim(txtkeiyakusuu.Text))

If n_kaisuu = 1 Then
    lblshuuryou.Caption = Trim(cmbnen.Text) & "/" & Trim(cmbtsuki.Text)
    Exit Sub
End If

Dim s_nen As String
Dim n_nen As Long

s_nen = Trim(cmbnen.Text)

If s_nen = "" Then
    Exit Sub
Else
    n_nen = CLng(s_nen)
End If

Dim s_tsuki As String
Dim n_tsuki As Integer

 s_tsuki = Trim(cmbtsuki.Text)
 
If s_tsuki = "" Then
    Exit Sub
Else
    n_tsuki = CInt(s_tsuki)
End If


Dim n_sou_tsuki As Integer

n_sou_tsuki = n_tsuki + n_kaisuu - 1

Dim n_add_nen As Integer
Dim n_add_tsuki As Integer

n_add_nen = Int(n_sou_tsuki / 12)
n_add_tsuki = n_sou_tsuki Mod 12

If n_add_tsuki = 0 Then
    n_add_tsuki = 12
    n_add_nen = n_add_nen - 1
End If

Dim n_tsugi_nen As Long
Dim n_tsugi_tsuki As Integer

n_tsugi_nen = n_nen + n_add_nen
n_tsugi_tsuki = n_add_tsuki

'èIóπì˙ÇçXêV

lblshuuryou.Caption = CStr(n_tsugi_nen) & "/" & Format(n_tsugi_tsuki, "00")

End Sub

Private Sub txtkuchisuu_Change()
If Trim(lbltannka.Caption) = "" Then
    Exit Sub
End If
txtkingaku.Text = CDbl(lbltannka.Caption) * CInt(txtkuchisuu.Text)




End Sub

