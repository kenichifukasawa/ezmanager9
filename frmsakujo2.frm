VERSION 5.00
Begin VB.Form frmsakujo2 
   Caption         =   "ó\íËä«óù"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   7800
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   6540
      MouseIcon       =   "frmsakujo2.frx":0000
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmsakujo2.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   5850
      Width           =   480
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5340
      MouseIcon       =   "frmsakujo2.frx":11C9
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmsakujo2.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   5850
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "å_ñÒó\íËçÌèúè⁄ç◊"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   270
      TabIndex        =   0
      Top             =   270
      Width           =   7335
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   1980
         TabIndex        =   20
         Top             =   3090
         Width           =   4755
         Begin VB.OptionButton optshurui 
            Caption         =   "çÌèú"
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
            Index           =   2
            Left            =   3390
            TabIndex        =   23
            Top             =   300
            Width           =   915
         End
         Begin VB.OptionButton optshurui 
            Caption         =   "íÜé~"
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
            Left            =   1875
            TabIndex        =   22
            Top             =   300
            Width           =   915
         End
         Begin VB.OptionButton optshurui 
            Caption         =   "ãxé~"
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
            Left            =   360
            TabIndex        =   21
            Top             =   300
            Value           =   -1  'True
            Width           =   915
         End
      End
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
         TabIndex        =   17
         Top             =   4320
         Width           =   4725
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
         Left            =   3270
         TabIndex        =   13
         Top             =   2460
         Width           =   735
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
         Left            =   2010
         TabIndex        =   12
         Top             =   2460
         Width           =   915
      End
      Begin VB.Label Label5 
         Alignment       =   2  'íÜâõëµÇ¶
         Caption         =   "éÌóﬁ"
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
         Left            =   480
         TabIndex        =   19
         Top             =   3120
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   2  'íÜâõëµÇ¶
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
         Left            =   450
         TabIndex        =   18
         Top             =   4350
         Width           =   1005
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
         Left            =   3570
         TabIndex        =   16
         Top             =   1920
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
         Left            =   4110
         TabIndex        =   15
         Top             =   2580
         Width           =   255
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
         Left            =   2970
         TabIndex        =   14
         Top             =   2580
         Width           =   255
      End
      Begin VB.Label lblkingaku 
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
         Left            =   1980
         TabIndex        =   11
         Top             =   1770
         Width           =   1395
      End
      Begin VB.Label lblkeiyakumei 
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
         Left            =   1980
         TabIndex        =   10
         Top             =   1140
         Width           =   5055
      End
      Begin VB.Label lblkeiyakuid 
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
         Left            =   1980
         TabIndex        =   9
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label Label7 
         Alignment       =   2  'íÜâõëµÇ¶
         Caption         =   "ó\íËîNåé"
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
         TabIndex        =   4
         Top             =   2520
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   2  'íÜâõëµÇ¶
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
         Left            =   390
         TabIndex        =   3
         Top             =   1860
         Width           =   1005
      End
      Begin VB.Label Label2 
         Alignment       =   2  'íÜâõëµÇ¶
         Caption         =   "å_ñÒì‡óe"
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
         TabIndex        =   2
         Top             =   1230
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'íÜâõëµÇ¶
         Caption         =   "å_ñÒID"
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
         Index           =   0
         Left            =   390
         TabIndex        =   1
         Top             =   600
         Width           =   1005
      End
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
      Left            =   6570
      TabIndex        =   8
      Top             =   6360
      Width           =   495
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
      Left            =   5370
      TabIndex        =   7
      Top             =   6390
      Width           =   480
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   1035
      Left            =   4920
      Top             =   5700
      Width           =   2580
   End
End
Attribute VB_Name = "frmsakujo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
  Unload frmsakujo2
    Set frmsakujo2 = Nothing
End Sub

Private Sub Picture4_Click()

Dim s_jisshi As String

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
    MsgBox "îNåéÇ™âﬂãéÇÃÇ‡ÇÃÇÕìoò^Ç≈Ç´Ç‹ÇπÇÒÅB"
    Exit Sub
ElseIf s_jisshi = s_itsu Then
    MsgBox "îNåéÇ™ìØàÍÇ≈Ç∑ÅBÇOçXêVÇ∆ÉfÅ[É^çXêVÇ™èIÇÌÇ¡ÇƒÇ¢ÇÈèÍçáÇÕÅAó\íËÇÕçÏìÆÇµÇ‹ÇπÇÒÅB"

End If
 

Dim s_kojinid As String

s_kojinid = Trim(frmtouroku2.lblid.Caption)

Dim s_kanrenid As String, s_kanrenmei As String, s_kingaku As String

s_kanrenid = Trim(lblkeiyakuid.Caption)
s_kanrenmei = Trim(lblkeiyakumei.Caption)
s_kingaku = Trim(lblkingaku.Caption)

Dim s_shurui As String

If optshurui(0).Value = True Then
    s_shurui = "1"
Else
    If optshurui(1).Value = True Then
        s_shurui = "2"
    Else
        If optshurui(2).Value = True Then
            s_shurui = "0"
        End If
    End If
End If




Dim s_bikou As String

s_bikou = Trim(txtbikou.Text)


If yoteikanri_touroku(s_kojinid, s_shurui, s_kanrenid, s_kanrenmei, s_kingaku, "1", "1", s_jisshi, s_jisshi, "1", s_bikou) = False Then


End If



 Unload frmsakujo2
    Set frmsakujo2 = Nothing



End Sub
