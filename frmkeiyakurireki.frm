VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmkeiyakurireki 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "ÉçÉO"
   ClientHeight    =   10710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10710
   ScaleWidth      =   14910
   StartUpPosition =   1  'µ∞≈∞ Ã´∞—ÇÃíÜâõ
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   13980
      MouseIcon       =   "frmkeiyakurireki.frx":0000
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmkeiyakurireki.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   25
      Top             =   1500
      Width           =   480
   End
   Begin VB.Frame fnaiyou 
      Caption         =   "ì‡óe"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5925
      Left            =   2520
      TabIndex        =   21
      Top             =   4020
      Visible         =   0   'False
      Width           =   10515
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   9420
         MouseIcon       =   "frmkeiyakurireki.frx":11C9
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkeiyakurireki.frx":1AC8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   23
         Top             =   4800
         Width           =   480
      End
      Begin VB.Label lblrow 
         Height          =   75
         Left            =   7740
         TabIndex        =   27
         Top             =   180
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ñﬂÇÈ"
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
         Index           =   0
         Left            =   9450
         TabIndex        =   24
         Top             =   5340
         Width           =   450
      End
      Begin VB.Label lblnaiyou 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   8595
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   1275
         Left            =   9180
         Top             =   4500
         Width           =   930
      End
   End
   Begin VB.CheckBox chksakujo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "çÌèúÇï\é¶Ç∑ÇÈ"
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
      Left            =   11820
      TabIndex        =   19
      Top             =   660
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.ComboBox cmbhouhou 
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
      Left            =   8640
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   18
      Top             =   540
      Width           =   2595
   End
   Begin VB.TextBox txtsaki 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   2  'µÃ
      Left            =   4800
      MaxLength       =   25
      TabIndex        =   17
      Top             =   600
      Width           =   2595
   End
   Begin VB.TextBox txtdare 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   4  'ëSäpÇ–ÇÁÇ™Ç»
      Left            =   1260
      MaxLength       =   25
      TabIndex        =   0
      Top             =   600
      Width           =   2355
   End
   Begin VB.TextBox txtlog 
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      IMEMode         =   4  'ëSäpÇ–ÇÁÇ™Ç»
      Left            =   300
      MaxLength       =   600
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   960
      Width           =   13395
   End
   Begin VB.ComboBox cmbshain 
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
      Left            =   11460
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   11
      Top             =   120
      Width           =   1635
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   13980
      MouseIcon       =   "frmkeiyakurireki.frx":2392
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmkeiyakurireki.frx":2C91
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   420
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "ÉçÉOàÍóó"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6945
      Left            =   120
      TabIndex        =   3
      Top             =   3660
      Width           =   14655
      Begin VB.CommandButton cmdmodoru 
         Caption         =   "ñﬂÇÈ"
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
         Height          =   315
         Left            =   11880
         TabIndex        =   31
         Top             =   180
         Width           =   795
      End
      Begin VB.CommandButton cmdchushutsu 
         Caption         =   "íäèo"
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
         Height          =   315
         Left            =   10980
         TabIndex        =   30
         Top             =   180
         Width           =   795
      End
      Begin VB.TextBox txtmojiretsu 
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
         IMEMode         =   4  'ëSäpÇ–ÇÁÇ™Ç»
         Left            =   8520
         TabIndex        =   29
         Top             =   180
         Width           =   2295
      End
      Begin MSFlexGridLib.MSFlexGrid gridlog 
         Height          =   6285
         Left            =   60
         TabIndex        =   4
         ToolTipText     =   "ÉçÉOàÍóóÇ≈Ç∑ÅB"
         Top             =   540
         Width           =   14475
         _ExtentX        =   25532
         _ExtentY        =   11086
         _Version        =   393216
         Cols            =   12
         FixedCols       =   0
         BackColor       =   16777215
         BackColorBkg    =   16777215
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         Caption         =   "ì‡óeÇÃï∂éöóÒåüçıÅF"
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
         Left            =   6420
         TabIndex        =   28
         Top             =   220
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   13980
      MouseIcon       =   "frmkeiyakurireki.frx":355B
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmkeiyakurireki.frx":3E5A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label SSCJIKKOU 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ïœçX"
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
      Left            =   14040
      TabIndex        =   26
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label lblend 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   14100
      TabIndex        =   20
      Top             =   60
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label4 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00C0FFFF&
      Caption         =   "ï˚ñ@ÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7560
      TabIndex        =   16
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label3 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00C0FFFF&
      Caption         =   "òAóçêÊÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3840
      TabIndex        =   15
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00C0FFFF&
      Caption         =   "ëäéËÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   300
      TabIndex        =   14
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label5 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00C0FFFF&
      Caption         =   "íSìñÅF"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10560
      TabIndex        =   12
      Top             =   180
      Width           =   795
   End
   Begin VB.Label lbljikan 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8580
      TabIndex        =   10
      Top             =   180
      Width           =   1575
   End
   Begin VB.Label lblnichiji 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6420
      TabIndex        =   9
      Top             =   180
      Width           =   1995
   End
   Begin VB.Label lblname 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1560
      TabIndex        =   8
      Top             =   180
      Width           =   4755
   End
   Begin VB.Label lblkojinid 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   7
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ìoò^"
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
      Left            =   14040
      TabIndex        =   6
      Top             =   960
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ñﬂÇÈ"
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
      Left            =   14040
      TabIndex        =   2
      Top             =   3060
      Width           =   450
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   3435
      Left            =   180
      Top             =   120
      Width           =   14520
   End
End
Attribute VB_Name = "frmkeiyakurireki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chksakujo_Click()

If chksakujo.Value = "1" Then
        If frmkeiyakurireki.lblend.Caption = "0" Then
            If Trim(lblkojinid.Caption) = "" Then
                logset 0, 0, "", 1
            Else
                logset 0, 1, Trim(lblkojinid.Caption), 1
            End If
        Else
            If Trim(lblkojinid.Caption) = "" Then
                logset 1, 0, "", 1
            Else
                logset 1, 1, Trim(lblkojinid.Caption), 1
            End If
        End If
Else
        If frmkeiyakurireki.lblend.Caption = "0" Then
            If Trim(lblkojinid.Caption) = "" Then
                logset 0, 0, "", 0
            Else
                logset 0, 1, Trim(lblkojinid.Caption), 0
            End If
        Else
            If Trim(lblkojinid.Caption) = "" Then
                logset 1, 0, "", 0
            Else
                logset 1, 1, Trim(lblkojinid.Caption), 0
            End If
        End If
End If

End Sub

Private Sub cmbhouhou_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtlog.SetFocus
End If
End Sub



Private Sub cmdchushutsu_Click()
    
    Dim shitei_moji As String
    
    shitei_moji = Trim(txtmojiretsu.Text)
    
    If shitei_moji = "" Then
        MsgBox "íäèoï∂éöÇì¸óÕÇµÇƒÇ©ÇÁé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        Exit Sub
    End If
    

    If frmmain.chkend.Value = 1 Then
            logset2 0, 0, shitei_moji
    Else
            logset2 1, 0, shitei_moji
    End If

End Sub

Private Sub cmdmodoru_Click()
    If frmmain.chkend.Value = 1 Then
            logset 0, 0
    Else
            logset 1, 0
    End If
txtmojiretsu.Text = ""

End Sub

Private Sub Form_Load()


lblnichiji.Caption = Format(Date, "yyyy/mm/dd")

lbljikan.Caption = Format(Time, "hh:mm")

cmbshain.AddItem "çÇà¿"
cmbshain.AddItem "â™ïî"
cmbshain.AddItem "óÈñÿ"
cmbshain.AddItem "ãvï€ìc"
cmbshain.AddItem "éOêX"
cmbshain.AddItem "íÜë∫"

cmbshain.AddItem "éuë∫"
cmbshain.AddItem "îãå¥"

'cmbhouhou.AddItem "ìdòb"
cmbhouhou.AddItem "óàé∫"
cmbhouhou.AddItem "ñKñ‚"
cmbhouhou.AddItem "Å@Å@"
cmbhouhou.AddItem "íÖêM"
cmbhouhou.AddItem "î≠êM"



 settei_res = setting(1, 13, 0, "", 1)
If settei_res = "-1" Then
    Ret = MsgBox("ê›íËÇÃì«Ç›çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB", 16, "EZ Manager")
Else
    cmbshain.ListIndex = CInt(settei_res)
End If
cmbhouhou.ListIndex = 0


txtmojiretsu.Text = ""

End Sub

Private Sub gridlog_DblClick()

Dim h_id As String, s_id As String

Dim s_sentakuretsu As Integer


s_sentakuretsu = gridlog.Col


Select Case s_sentakuretsu
    Case 4
    
            gridlog.Col = 10
            s_id = Trim(gridlog.Text)
            'å_ñÒé“èÓïÒÇ…ÉWÉÉÉìÉv
            If s_id <> "" Then
                    If hyouji_touroku(s_id) = False Then
                        Exit Sub
                    End If
                    If nhk_set(s_id) = False Then
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                    
                    keiyakuset (s_id)
                    keiyakuset2 (s_id)
                    chk_kyuushi (s_id)
                    chk_log (s_id)
                    
                    Unload frmkeiyakurireki
                    Set frmkeiyakurireki = Nothing
                    frmtouroku2.Show 1
            End If
    Case Else
    
            gridlog.Col = 0
            h_id = Trim(gridlog.Text)
            'å_ñÒèÓïÒè⁄ç◊Çï\é¶
            
            lblrow.Caption = gridlog.Row
            
            If h_id = "" Then
                MsgBox "ê≥èÌÇ…ëIëÇ≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÇ≈ÇµÇΩÅB"
                Exit Sub
            End If
            If Len(h_id) <> 8 Then
                MsgBox "ëIëÇ≥ÇÍÇΩIDÇ™ïsê≥Ç≈Ç∑ÅB"
                Exit Sub
            End If
            sql = "SELECT * FROM log WHERE logid ='" & h_id & "'"
            
            Set rs = db.OpenRecordset(sql, dbOpenDynaset, dbReadOnly)
            If FcSQlGet(rs, sql, PrMsg) = False Then
                MsgBox "ëIëÇµÇΩçÄñ⁄Ç™å©Ç¬Ç©ÇËÇ‹ÇπÇÒÅB"
                Exit Sub
            Else
                lblnaiyou.Caption = Trim(rs!youken)
                fnaiyou.Visible = True
                rs.Close
            End If
            
End Select
          
            
            
            
            
End Sub


Private Sub Picture1_Click()
On Error GoTo errsuuji
lblnaiyou.Caption = ""
fnaiyou.Visible = False
gridlog.Row = CInt(lblrow.Caption)
gridlog.Col = 0
gridlog.RowSel = CInt(lblrow.Caption)
gridlog.ColSel = 9
gridlog.SetFocus

Exit Sub

errsuuji:
lblnaiyou.Caption = ""

fnaiyou.Visible = False


End Sub

Private Sub Picture13_Click()
Unload Me
End Sub

Private Sub Picture2_Click()
If frmmain.lblkanri2.Visible = True Then
      MsgBox "ÉTÉuÉÇÅ[ÉhÇ≈ÇÕçXêVÇ≈Ç´Ç‹ÇπÇÒÅB"
        Exit Sub
End If


Dim h_id As String

            gridlog.Col = 0
            h_id = Trim(gridlog.Text)
            
            If h_id = "" Then
                MsgBox "ê≥èÌÇ…ëIëÇ≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÇ≈ÇµÇΩÅB"
                Exit Sub
            End If
            If Len(h_id) <> 8 Then
                MsgBox "ëIëÇ≥ÇÍÇΩIDÇ™ïsê≥Ç≈Ç∑ÅB"
                Exit Sub
            End If
            sql = "SELECT * FROM log WHERE logid ='" & h_id & "'"
            
            Set rs = db.OpenRecordset(sql, dbOpenDynaset, dbReadOnly)
            If FcSQlGet(rs, sql, PrMsg) = False Then
                MsgBox "ëIëÇµÇΩçÄñ⁄Ç™å©Ç¬Ç©ÇËÇ‹ÇπÇÒÅB"
                Exit Sub
            Else
                With frmlogshousai
                    .cmbshain.AddItem "çÇà¿"
                    .cmbshain.AddItem "â™ïî"
                     .cmbshain.AddItem "óÈñÿ"
                      .cmbshain.AddItem "ãvï€ìc"
                      .cmbshain.AddItem "éOêX"
                    .cmbshain.AddItem "íÜë∫"
                    .cmbshain.AddItem "éuë∫"
                    .cmbshain.AddItem "îãå¥"
                      
                    '.cmbhouhou.AddItem "ìdòb"
                    .cmbhouhou.AddItem "óàé∫"
                    .cmbhouhou.AddItem "ñKñ‚"
                    .cmbhouhou.AddItem "Å@Å@"
                    .cmbhouhou.AddItem "íÖêM"
                    .cmbhouhou.AddItem "î≠êM"
                    .cmbkekka.AddItem "â“ìÆ"
                    .cmbkekka.AddItem "èIóπ"
                    .lbllogid.Caption = rs!logid
                    .lblkojinid.Caption = rs!kojinid
                    .txtnichiji.Text = Format(rs!nichiji, "@@@@/@@/@@")
                    .txtjikan.Text = Format(rs!jikan, "@@:@@")
                    .txtlog.Text = Trim(rs!youken)
                    If IsNull(rs!dare) Then
                        .txtdare.Text = ""
                    Else
                        .txtdare.Text = Trim(rs!dare)
                    End If
                    If IsNull(rs!saki) Then
                        .txtsaki.Text = ""
                    Else
                        .txtsaki.Text = Trim(rs!saki)
                     End If
                    .cmbshain.ListIndex = CInt(rs!tantou)
                    If rs!houhou = 0 Then
                        .cmbhouhou.ListIndex = -1
                    Else
                        .cmbhouhou.ListIndex = CInt(rs!houhou) - 1
                    End If
                    .cmbkekka.ListIndex = CInt(rs!kekka)
                    .lblname.Caption = lblname.Caption
                End With
                rs.Close
                frmlogshousai.Show 1
            End If
            

End Sub

Private Sub Picture7_Click()

Dim tid As String, tkojinid As String, tnichiji As String, tjikan As String
Dim tnaiyou As String, ttantou As String, ttid As String, shime As String
Dim tdare As String, tsaki As String, thouhou As String


If frmmain.lblkanri2.Visible = True Then
      MsgBox "ÉTÉuÉÇÅ[ÉhÇ≈ÇÕçXêVÇ≈Ç´Ç‹ÇπÇÒÅB"
        Exit Sub
End If



    If Trim(txtlog.Text) <> "" Then
            If soukatsu(0, "renban", 13, "", 1, 1) = False Then
                    MsgBox "ê›íËÇÃì«Ç›çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
            End If
            If reg = "0" Then
                reg = ""
            End If
    
            If reg = vbNullString Then
                tid = "00000001"
            Else
                tid = Format$(reg, "0000000#")
            End If
        With frmkeiyakurireki
            If Trim(.txtdare.Text) = "" Then
                tdare = Space(1)
            Else
                tdare = Trim(.txtdare.Text)
            End If
            If Trim(.txtsaki.Text) = "" Then
                tsaki = Space(1)
            Else
                tsaki = Trim(.txtsaki.Text)
            End If
            If .cmbhouhou.ListIndex = -1 Then
                MsgBox "ï˚ñ@ÇëIëÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                Exit Sub
            Else
                thouhou = CStr(.cmbhouhou.ListIndex + 1)
            End If
            ttantou = CStr(.cmbshain.ListIndex)
            If Trim(.lblkojinid.Caption) = "" Then
                tkojinid = Space(1)
            Else
                tkojinid = Trim(.lblkojinid.Caption)
            End If
            tnichiji = Format(.lblnichiji.Caption, "yyyymmdd")
            tjikan = Format(.lbljikan.Caption, "hhmm")
            tnaiyou = Trim(.txtlog.Text)
        End With
        sql = "INSERT INTO log VALUES('" & tid & "','" & tkojinid & "'," & _
                       "'" & tnichiji & "','" & tjikan & "','" & ttantou & "'," & _
                       "'" & tnaiyou & "','0','0','" & tdare & "','" & tsaki & "','" & thouhou & "')"
        
On Error GoTo errk4
        db.Execute sql, dbSQLPassThrough
On Error GoTo 0
        ttid = CStr(CLng(tid) + 1)
        If soukatsu(1, "renban", 13, ttid, 1, 1) = False Then
                MsgBox "ê›íËÇÃèëÇ´çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
                Exit Sub
        End If
        If frmkeiyakurireki.lblend.Caption = "0" Then
            If Trim(lblkojinid.Caption) = "" Then
                logset 0, 0
            Else
                logset 0, 1, tkojinid
            End If
        Else
            If Trim(lblkojinid.Caption) = "" Then
                logset 1, 0
            Else
                logset 1, 1, tkojinid
            End If
        End If
        
                'Log
        Dim ss_name As String
        Dim ss_id As String
        ss_id = Trim(tkojinid)
        ss_name = ""
        sagyou_msg = "ÉçÉOè⁄ç◊ÅFÉçÉOÇêVãKìoò^ÇµÇ‹ÇµÇΩÅB"
        log_sagyou sagyou_log_path, sagyou_msg, ss_id, ss_name



        MsgBox "ìoò^ÇµÇ‹ÇµÇΩÅB"
        txtlog.Text = ""
        txtdare.Text = ""
        txtsaki.Text = ""
        cmbhouhou.ListIndex = 0
        lblnichiji.Caption = Format(Date, "yyyy/mm/dd")
        lbljikan.Caption = Format(Time, "hh:mm")
        
    Else
        MsgBox "ì‡óeÇì¸óÕÇµÇƒÇ©ÇÁÅAé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        txtlog.SetFocus
        Exit Sub
    End If
Exit Sub

errk4:
    MsgBox "ÉçÉOÇÃìoò^Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
    Exit Sub
End Sub

Private Sub txtdare_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    txtsaki.SetFocus
End If

End Sub

Private Sub txtsaki_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbhouhou.SetFocus
End If

End Sub
