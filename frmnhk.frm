VERSION 5.00
Begin VB.Form frmnhk 
   BorderStyle     =   1  'ŒÅ’è(Àü)
   Caption         =   "NHKXV‰æ–Ê"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   8100
   StartUpPosition =   2  '‰æ–Ê‚Ì’†‰›
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Ì×¯Ä
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '‚È‚µ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   7140
      MouseIcon       =   "frmnhk.frx":0000
      MousePointer    =   99  'Õ°»Ş°’è‹`
      Picture         =   "frmnhk.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   23
      Top             =   270
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ì×¯Ä
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '‚È‚µ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   7170
      MouseIcon       =   "frmnhk.frx":11C9
      MousePointer    =   99  'Õ°»Ş°’è‹`
      Picture         =   "frmnhk.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   22
      Top             =   1500
      Width           =   480
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "NHKŒ_–ñ“à—e"
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2235
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6675
      Begin VB.TextBox txtkuchisuu 
         Alignment       =   1  '‰E‘µ‚¦
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   2  'µÌ
         Left            =   3480
         MaxLength       =   8
         TabIndex        =   18
         Top             =   1500
         Width           =   735
      End
      Begin VB.ComboBox combjikaitsuki 
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5520
         TabIndex        =   15
         ToolTipText     =   "FD‚Éo—Í‚·‚éŒ‚ğ‘I‘ğ‚µ‚Ä‚­‚¾‚³‚¢B"
         Top             =   720
         Width           =   675
      End
      Begin VB.ComboBox combjikainen 
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4260
         TabIndex        =   14
         ToolTipText     =   "FD‚Éo—Í‚·‚é”N‚ğ‘I‘ğ‚µ‚Ä‚­‚¾‚³‚¢B"
         Top             =   720
         Width           =   915
      End
      Begin VB.ComboBox combzentsuki 
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2880
         TabIndex        =   11
         Top             =   720
         Width           =   675
      End
      Begin VB.ComboBox combzennen 
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1620
         TabIndex        =   10
         Top             =   720
         Width           =   915
      End
      Begin VB.ComboBox combshurui 
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   180
         Style           =   2  'ÄŞÛ¯ÌßÀŞ³İ Ø½Ä
         TabIndex        =   8
         Top             =   1500
         Width           =   2835
      End
      Begin VB.TextBox txtkingaku 
         Alignment       =   1  '‰E‘µ‚¦
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   1  'µİ
         Left            =   4920
         MaxLength       =   8
         TabIndex        =   1
         ToolTipText     =   "•ÏX‚·‚éê‡‚ÍAí—ŞEŒû”‚ğİ’èŒãAÅŒã‚Éİ’è‚µ‚Ä‚­‚¾‚³‚¢B"
         Top             =   1500
         Width           =   1155
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Œ_–ñŒû”"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3240
         TabIndex        =   20
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Œû"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         TabIndex        =   19
         Top             =   1620
         Width           =   315
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Œ"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   17
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "”N"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5220
         TabIndex        =   16
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Œ"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "”N"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2580
         TabIndex        =   12
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "‰~"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6180
         TabIndex        =   9
         Top             =   1620
         Width           =   315
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "NHKŒ_–ñí—Ş"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ÿ‰ñ¿‹“ú"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4320
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "‘O‰ñ¿‹“ú"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   1740
         TabIndex        =   5
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "NHK ID”Ô†"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Œ_–ñ‹àŠz"
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4920
         TabIndex        =   3
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label lblnhkid 
         Alignment       =   2  '’†‰›‘µ‚¦
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   2
         Top             =   660
         Width           =   1035
      End
   End
   Begin VB.Label SSCJIKKOU 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "“o˜^"
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   7170
      TabIndex        =   21
      Top             =   780
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "I—¹"
      BeginProperty Font 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
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
      Left            =   7200
      TabIndex        =   24
      Top             =   2010
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '•s“§–¾
      BorderStyle     =   0  '“§–¾
      Height          =   2235
      Left            =   6870
      Top             =   120
      Width           =   1050
   End
End
Attribute VB_Name = "frmnhk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub combshurui_Click()
If Trim(txtkuchisuu.Text) = "" Then
    txtkuchisuu.Text = "1"
End If


End Sub


Private Sub Form_Load()
Dim nntt As Integer, karinntt As Long

With frmnhk
    .combshurui.AddItem "ŒûÀU‘Ö‚Pƒ–Œ"
    .combshurui.AddItem "ŒûÀU‘Ö‚Qƒ–Œ"
    .combshurui.AddItem "ŒûÀU‘Ö‚Rƒ–Œ"
    .combshurui.AddItem "ŒûÀU‘Ö‚Sƒ–Œ"
    .combshurui.AddItem "ŒûÀU‘Ö‚Tƒ–Œ"
    .combshurui.AddItem "ŒûÀU‘Ö‚Uƒ–Œ"
    .combshurui.AddItem "ŒûÀU‘Ö‚Uƒ–Œƒnƒ“ƒf"
    .combshurui.AddItem "ŒûÀU‘Ö‚Uƒ–Œ‰Æ‘°Š„"
   For nntt = -5 To 1
        karinntt = CLng(Format(Date, "yyyy")) + nntt
        .combzennen.AddItem karinntt
        .combjikainen.AddItem karinntt
   Next nntt
   For nntt = 1 To 12
        .combzentsuki.AddItem Format(nntt, "00")
        .combjikaitsuki.AddItem Format(nntt, "00")
   Next nntt
   txtkuchisuu.Text = "1"
   
End With
End Sub

Private Sub Picture13_Click()
Unload Me

End Sub

Private Sub Picture7_Click()


Dim newid As String, newmaebi As String, newjikaibi As String, nnid As String, flg_nhk As String
Dim newshurui As String, newkingaku As String, sql_nhk3 As String, newkuchisuu As String
Dim newkojinid As String, newmaebi2 As Long, newjikaibi2 As Long, nhkjoukyou As String
Dim newjikaibi3 As Long, kyounen As Long, kyoutsuki As Long, karinhkkingaku As Long
Dim newkakinkakin As Boolean

If frmtouroku2.lbljoukyou.Caption = "•s‰Â" Then
    nhkjoukyou = "1"
Else
    nhkjoukyou = "0"
End If


With frmnhk

    newid = Trim(.lblnhkid.Caption)
    If newid = "" Then
        MsgBox "‚m‚g‚j‚h‚c‚ª•s³‚Å‚·B"
        Exit Sub
    End If
    newkojinid = Trim(frmtouroku2.lblid.Caption)
    If newkojinid = "" Then
        MsgBox "ŒÂl‚h‚c‚ª•s³‚Å‚·B"
        Exit Sub
    End If
On Error GoTo errhiduke
    If Trim(.combzennen.Text) <> "" Then
        If Trim(.combzentsuki.Text) <> "" Then
            newmaebi2 = CLng(Trim(.combzennen.Text))
            newmaebi = Trim(.combzennen.Text)
            newmaebi2 = CLng(Trim(.combzentsuki.Text))
            newmaebi = newmaebi & Trim(.combzentsuki.Text) & "00"
        Else
            MsgBox "”N‚ª“ü—Í‚³‚ê‚Ä‚¢‚ÄAŒ‚ª“ü—Í‚³‚ê‚Ä‚¢‚Ü‚¹‚ñB"
            Exit Sub
        End If
    Else
        If Trim(.combzentsuki.Text) <> "" Then
            MsgBox "Œ‚ª“ü—Í‚³‚ê‚Ä‚¢‚ÄA”N‚ª“ü—Í‚³‚ê‚Ä‚¢‚Ü‚¹‚ñB"
            Exit Sub
        Else
            newmaebi = Space(1)
        End If
    End If
    If Trim(.combjikainen.Text) <> "" Then
        If Trim(.combjikaitsuki.Text) <> "" Then
            newjikaibi2 = CLng(Trim(.combjikainen.Text))
            newjikaibi = Trim(.combjikainen.Text)
            newjikaibi3 = CLng(Trim(.combjikaitsuki.Text))
            newjikaibi = newjikaibi & Trim(.combjikaitsuki.Text) & "00"
         Else
            MsgBox "”N‚ª“ü—Í‚³‚ê‚Ä‚¢‚ÄAŒ‚ª“ü—Í‚³‚ê‚Ä‚¢‚Ü‚¹‚ñB"
            Exit Sub
         End If
    Else
        If Trim(.combjikaitsuki.Text) <> "" Then
            MsgBox "Œ‚ª“ü—Í‚³‚ê‚Ä‚¢‚ÄA”N‚ª“ü—Í‚³‚ê‚Ä‚¢‚Ü‚¹‚ñB"
            Exit Sub
        Else
            MsgBox "Ÿ‰ñ¿‹“úi¿‹ŠJn“új‚ª“ü—Í‚³‚ê‚Ä‚¢‚Ü‚¹‚ñB"
            Exit Sub
        End If
    End If
    '‚¢‚ÂÀs‚·‚é‚©A³Šm‚©H
    kyounen = CLng(Format(Date, "yyyy"))
    kyoutsuki = CLng(Format(Date, "mm"))
    If kyounen = newjikaibi2 Then
        If kyoutsuki = newjikaibi3 Then '“¯‚¶”NA“¯‚¶Œ
                If CLng(Format(Date, "dd")) >= 13 Then
                    MsgBox "¿‹“ú‚ª¡Œ‚Å‚·‚ªA‚·‚Å‚É¡Œ‚Ì¿‹‚ÍI—¹‚µ‚Ä‚¢‚Ü‚·B—ˆŒ‚Ì¿‹‚É•ÏX‚µ‚Ä‚­‚¾‚³‚¢B"
                    Exit Sub
                End If
        End If
        If kyoutsuki = newjikaibi3 Then   'U‘ÖŒ‚ª“¯‚¶
            
        ElseIf kyoutsuki > newjikaibi3 Then   '“¯‚¶”NAU‘ÖŒ‚ªˆÈ‘O
                    MsgBox "¿‹“ú‚ª¡ŒˆÈ‘O‚Å‚·B‚Ü‚½‚·‚Å‚É¡Œ‚Ì¿‹‚ÍI—¹‚µ‚Ä‚¢‚Ü‚·‚Ì‚ÅA—ˆŒ‚Ì¿‹‚É•ÏX‚µ‚Ä‚­‚¾‚³‚¢B"
                    Exit Sub
            
        Else
            If nhkjoukyou = "1" Then
                MsgBox "¿‹“ú‚ª¡ŒˆÈ‘O‚Å‚·B‚Ü‚½‚·‚Å‚É¡Œ‚Ì¿‹‚ÍI—¹‚µ‚Ä‚¢‚Ü‚·‚Ì‚ÅA—ˆŒ‚Ì¿‹‚É•ÏX‚µ‚Ä‚­‚¾‚³‚¢B"
                Exit Sub
            Else
                If newjikaibi3 - kyoutsuki = 1 Then '“¯‚¶”NA‘O‚ÌŒ
                
                Else
                    nhkjoukyou = "1"
                End If
            End If
        End If
    ElseIf kyounen > newjikaibi2 Then '”N‚ªˆÈ‘O‚Ìê‡
        MsgBox "Ÿ‰ñU‘Ö“ú‚ª¡”N‚æ‚èˆÈ‘O‚É‚È‚Á‚Ä‚¢‚Ü‚·B"
        Exit Sub
    Else '—ˆ”NˆÈ~
        nhkjoukyou = "1"
    End If
    
On Error GoTo 0
    newshurui = CStr(.combshurui.ListIndex)
    If newshurui = "-1" Then
        MsgBox "í—Ş‚Ì’l‚ª•s³‚Å‚·B"
        Exit Sub
    End If
    newkingaku = Trim(.txtkingaku.Text)
    If newkingaku = "" Then
        MsgBox "‹àŠz‚Ì’l‚ª•s³‚Å‚·B"
        Exit Sub
    Else
        On Error GoTo errnhksuuji
        karinhkkingaku = CLng(newkingaku)
        On Error GoTo 0
        
    End If
    newkuchisuu = Trim(txtkuchisuu.Text)
    If nhkjoukyou = "1" Then
        flg_nhk = "0"
        newkakinkakin = False
        
    Else
        flg_nhk = "1"
        newkakinkakin = True
    End If
If kidou4 = False Then
    Exit Sub
End If
    
    
    If SSCJIKKOU.Caption = "“o˜^" Then
        newmaebi = Space(1)
        sql_nhk3 = "INSERT INTO nhk_m VALUES('" & newid & "'," & _
                       "'" & newkojinid & "','" & newshurui & "','" & newkingaku & "'," & _
                       "'" & newmaebi & "','" & newjikaibi & "'," & _
                       "'" & newkuchisuu & "','" & flg_nhk & "'," & _
                       "'" & Space(1) & "','" & Space(1) & "'," & _
                       "'" & Space(1) & "','" & Space(1) & "'," & _
                       "'" & Space(1) & "','" & Space(1) & "'," & _
                       "'" & Space(1) & "','" & Space(1) & "')"
        db4.Execute sql_nhk3, dbSQLPassThrough
        nnid = CStr(CLng(newid) + 1)
           If soukatsu(1, "renban", 9, nnid, 1, 1) = False Then
                   MsgBox "İ’è‚Ì‘‚«‚İ‚É¸”s‚µ‚Ü‚µ‚½B"
                   Exit Sub
           End If
        If nhkjoukyou = "0" Then
            If a_nhk_touroku(newkojinid, newkingaku) = False Then
                       MsgBox "—¿‹à‚ÌXV‚É¸”s‚µ‚Ü‚µ‚½B"
                       Exit Sub
            End If
        End If
        
        'Log
        Dim ss_name As String
        ss_name = frmtouroku2.lblname.Caption
        Dim ss_id As String
        ss_id = frmtouroku2.lblid.Caption
        sagyou_msg = "NHKˆ—FNHK ID " & newid & "‚ÌŒ_–ñ‚ğV‹K“o˜^‚µ‚Ü‚µ‚½B"
        log_sagyou sagyou_log_path, sagyou_msg, ss_id, ss_name
        
        MsgBox "“o˜^‚µ‚Ü‚µ‚½B"
    Else
        If newmaebi = "" Then
            newmaebi = Space(1)
        End If
         sql = "UPDATE nhk_m SET n1='" & newkuchisuu & "'," & _
                      "shurui = '" & newshurui & "',kingaku = '" & newkingaku & "'," & _
                      "n2 = '" & flg_nhk & "'," & _
                      "zenkai = '" & newmaebi & "'," & _
                      "jikai = '" & newjikaibi & "'" & _
                      " WHERE nhkid = '" & newid & "'"
         db4.Execute sql, dbSQLPassThrough
         If a_nhk_henkou(newkojinid, newkingaku, newkakinkakin) = False Then
                    MsgBox "—¿‹à‚ÌXV‚É¸”s‚µ‚Ü‚µ‚½B"
                    Exit Sub
         End If
         
         'Log
        Dim sss_name As String
        sss_name = frmtouroku2.lblname.Caption
        Dim sss_id As String
        sss_id = frmtouroku2.lblid.Caption
        sagyou_msg = "NHKˆ—FNHK ID " & newid & "‚ÌŒ_–ñ‚ğ•ÏX‚µ‚Ü‚µ‚½B"
        log_sagyou sagyou_log_path, sagyou_msg, sss_id, sss_name
        
        
        MsgBox "•ÏX‚µ‚Ü‚µ‚½B"
    End If
db4.Close



    If nhk_set(newkojinid) = -1 Then
        Screen.MousePointer = 0
        Unload Me
        Unload frmtouroku2
        Set frmtouroku2 = Nothing
        Exit Sub
    End If

    Unload Me
    
End With
Exit Sub

errhiduke:
    MsgBox "“ú•t‚Í”¼Šp”š‚ğ“ü—Í‚·‚é‚©A‘I‘ğ‚µ‚Ä‚­‚¾‚³‚¢B"
    Exit Sub
errnhksuuji:
    MsgBox "—¿‹à‚Í”¼Šp”š‚ğ“ü—Í‚µ‚Ä‚©‚çÀs‚µ‚Ä‚­‚¾‚³‚¢B"
    Exit Sub


End Sub

Private Sub SSCSHUURYOU_Click()

End Sub

Private Sub SSCJIKKOU2_Click()

End Sub

Private Sub txtkingaku_GotFocus()
txtkingaku.SelStart = 0
txtkingaku.SelLength = Len(txtkingaku.Text)
End Sub

Private Sub txtkingaku_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub

Private Sub txtkuchisuu_GotFocus()
txtkuchisuu.SelStart = 0
txtkuchisuu.SelLength = Len(txtkuchisuu.Text)
End Sub

Private Sub txtkuchisuu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtkingaku.SetFocus
End If
End Sub

