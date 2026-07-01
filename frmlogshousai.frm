VERSION 5.00
Begin VB.Form frmlogshousai 
   Caption         =   "ÉçÉOè⁄ç◊"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   12060
   StartUpPosition =   1  'µ∞≈∞ Ã´∞—ÇÃíÜâõ
   Begin VB.TextBox txtjikan 
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
      Left            =   720
      MaxLength       =   25
      TabIndex        =   21
      ToolTipText     =   "ó·Åj13:23"
      Top             =   4140
      Width           =   1275
   End
   Begin VB.TextBox txtnichiji 
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
      Left            =   720
      MaxLength       =   25
      TabIndex        =   20
      ToolTipText     =   "ó·Åj2011/01/01"
      Top             =   3660
      Width           =   1635
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
      Left            =   1440
      MaxLength       =   25
      TabIndex        =   16
      Top             =   3180
      Width           =   2355
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
      Left            =   4980
      MaxLength       =   25
      TabIndex        =   15
      Top             =   3180
      Width           =   2595
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
      Left            =   3900
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   14
      Top             =   3600
      Width           =   2595
   End
   Begin VB.ComboBox cmbkekka 
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
      Left            =   3900
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   12
      Top             =   4140
      Width           =   3735
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   9180
      MouseIcon       =   "frmlogshousai.frx":0000
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmlogshousai.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   3720
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   10800
      MouseIcon       =   "frmlogshousai.frx":11C9
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmlogshousai.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   3720
      Width           =   480
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
      Left            =   180
      MaxLength       =   600
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   660
      Width           =   11715
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   9990
      MouseIcon       =   "frmlogshousai.frx":2392
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmlogshousai.frx":2C91
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   3720
      Width           =   480
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
      Left            =   7440
      Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
      TabIndex        =   1
      Top             =   240
      Width           =   1635
   End
   Begin VB.Label Label6 
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
      Left            =   480
      TabIndex        =   19
      Top             =   3180
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
      Left            =   4020
      TabIndex        =   18
      Top             =   3180
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
      Left            =   3000
      TabIndex        =   17
      Top             =   3660
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00C0FFFF&
      Caption         =   "åãâ ÅF"
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
      Left            =   3000
      TabIndex        =   13
      Top             =   4200
      Width           =   795
   End
   Begin VB.Label lbllogid 
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
      Left            =   9120
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FFFFFF&
      Caption         =   "çÌèú"
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
      Left            =   9060
      TabIndex        =   10
      Top             =   4230
      Width           =   675
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
      Left            =   10830
      TabIndex        =   8
      Top             =   4230
      Width           =   450
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
      Index           =   5
      Left            =   10035
      TabIndex        =   7
      Top             =   4230
      Width           =   480
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
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1215
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
      Left            =   1440
      TabIndex        =   5
      Top             =   240
      Width           =   4755
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
      Left            =   6540
      TabIndex        =   4
      Top             =   300
      Width           =   795
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   4695
      Left            =   60
      Top             =   120
      Width           =   11940
   End
End
Attribute VB_Name = "frmlogshousai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Picture11_Click()

msg = "ñ{ìñÇ…çÌèúÇµÇƒÇ‡ÇÊÇ¢Ç≈Ç∑Ç©ÅH"
res = MsgBox(msg, vbYesNo)

If res = vbYes Then


     sql = "UPDATE log SET sakujo = '1'" & _
            " WHERE logid = '" & lbllogid.Caption & "'"

On Error GoTo errk3
     db.Execute sql, dbSQLPassThrough
On Error GoTo 0

If Trim(frmkeiyakurireki.lblkojinid.Caption) = "" Then
    logset 0, 0
Else
    logset 0, 0, frmkeiyakurireki.lblkojinid.Caption
End If

End If

Unload Me

     'Log
        Dim ss_name As String
        Dim ss_id As String
        ss_id = ""
        ss_name = ""
        sagyou_msg = "ÉçÉOè⁄ç◊ÅFÉçÉOID " & Trim(lbllogid.Caption) & "ÇÃÉçÉOÇçÌèúÇµÇ‹ÇµÇΩÅB"
        log_sagyou sagyou_log_path, sagyou_msg, ss_id, ss_name

Exit Sub

errk3:
MsgBox "çÌèúÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
Exit Sub


End Sub

Private Sub Picture13_Click()
Unload Me
End Sub

Private Sub Picture7_Click()

Dim tid As String, tkojinid As String, tnichiji As String, tjikan As String
Dim tnaiyou As String, ttantou As String, ttid As String, shime As String
Dim tdare As String, tsaki As String, thouhou As String, tkekka As String
    
    tid = Trim(lbllogid.Caption)
    
    If tid <> "" Then
    
        With frmlogshousai
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
            If Len(.txtnichiji.Text) <> 10 Then
                MsgBox "ì˙ïtÇÃì¸óÕÇ™ïsê≥Ç≈Ç∑ÅBó·Åj2011/02/04"
                Exit Sub
            End If
            If Mid(.txtnichiji.Text, 5, 1) <> "/" Or Mid(.txtnichiji.Text, 8, 1) <> "/" Then
                MsgBox "ì˙ïtÇÃì¸óÕÇ™ïsê≥Ç≈Ç∑ÅBó·Åj2011/02/04"
                Exit Sub
            End If
            tnichiji = Format(.txtnichiji.Text, "yyyymmdd")
            If Len(.txtjikan.Text) <> 5 Then
                MsgBox "éûä‘ÇÃì¸óÕÇ™ïsê≥Ç≈Ç∑ÅBó·Åj13:11"
                Exit Sub
            End If
            If Mid(.txtjikan.Text, 3, 1) <> ":" Then
                MsgBox "éûä‘ÇÃì¸óÕÇ™ïsê≥Ç≈Ç∑ÅBó·Åj13:11"
                Exit Sub
            End If
            tjikan = Format(.txtjikan.Text, "hhmm")
            tkekka = CStr(.cmbkekka.ListIndex)
            tnaiyou = Trim(.txtlog.Text)
        End With
        sql = "UPDATE log SET nichiji = '" & tnichiji & "'" & _
             ",jikan = '" & tjikan & "'" & _
             ",tantou = '" & ttantou & "'" & _
             ",youken = '" & tnaiyou & "'" & _
             ",kekka = '" & tkekka & "'" & _
             ",dare = '" & tdare & "'" & _
             ",saki = '" & tsaki & "'" & _
             ",houhou = '" & thouhou & "'" & _
             " WHERE logid = '" & tid & "'"

        
        
On Error GoTo errk4
        db.Execute sql, dbSQLPassThrough
On Error GoTo 0


                'Log
        Dim ss_name As String
        Dim ss_id As String
        ss_id = Trim(tkojinid)
        ss_name = ""
        sagyou_msg = "ÉçÉOè⁄ç◊ÅFÉçÉOID" & tid & "ÇÃÉçÉOÇçXêVÇµÇ‹ÇµÇΩÅB"
        log_sagyou sagyou_log_path, sagyou_msg, ss_id, ss_name



        MsgBox "ïœçXÇµÇ‹ÇµÇΩÅB"

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
        Unload Me
        Exit Sub
        
    Else
        MsgBox "ÉçÉOIDÇ™ïsê≥Ç≈Ç∑ÅB"
        Unload Me
        Exit Sub
    End If
Exit Sub

errk4:
    MsgBox "ÉçÉOÇÃïœçXÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
    Unload Me
    Exit Sub
End Sub
