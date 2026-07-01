VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmkeiyakurireki2 
   Caption         =   "ÉçÉO"
   ClientHeight    =   10665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   ScaleHeight     =   10665
   ScaleWidth      =   14865
   StartUpPosition =   1  'µ∞≈∞ Ã´∞—ÇÃíÜâõ
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
      Left            =   2550
      TabIndex        =   6
      Top             =   2730
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
         MouseIcon       =   "frmkeiyakurireki2.frx":0000
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkeiyakurireki2.frx":08FF
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   7
         Top             =   4800
         Width           =   480
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
         TabIndex        =   10
         Top             =   360
         Width           =   8595
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
         TabIndex        =   9
         Top             =   5340
         Width           =   450
      End
      Begin VB.Label lblrow 
         Height          =   75
         Left            =   7740
         TabIndex        =   8
         Top             =   180
         Visible         =   0   'False
         Width           =   255
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
      Height          =   10395
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   14655
      Begin VB.CheckBox chkwsh 
         Caption         =   "WSH"
         Height          =   195
         Left            =   8550
         TabIndex        =   23
         Top             =   180
         Value           =   1  '¡™Ø∏
         Width           =   675
      End
      Begin VB.PictureBox Picture11 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   10200
         MouseIcon       =   "frmkeiyakurireki2.frx":11C9
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkeiyakurireki2.frx":1AC8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   21
         Top             =   210
         Width           =   480
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
         Left            =   7140
         TabIndex        =   18
         Top             =   510
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   12600
         MouseIcon       =   "frmkeiyakurireki2.frx":2392
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkeiyakurireki2.frx":2C91
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   15
         Top             =   210
         Width           =   480
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   11460
         MouseIcon       =   "frmkeiyakurireki2.frx":355B
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkeiyakurireki2.frx":3E5A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   13
         Top             =   210
         Width           =   480
      End
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   13740
         MouseIcon       =   "frmkeiyakurireki2.frx":4724
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkeiyakurireki2.frx":5023
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   11
         Top             =   210
         Width           =   480
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
         Left            =   2640
         TabIndex        =   3
         Top             =   420
         Width           =   2295
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
         Left            =   5100
         TabIndex        =   2
         Top             =   420
         Width           =   795
      End
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
         Left            =   6000
         TabIndex        =   1
         Top             =   420
         Width           =   795
      End
      Begin MSFlexGridLib.MSFlexGrid gridlog 
         Height          =   9225
         Left            =   90
         TabIndex        =   4
         ToolTipText     =   "ÉçÉOàÍóóÇ≈Ç∑ÅB"
         Top             =   1050
         Width           =   14475
         _ExtentX        =   25532
         _ExtentY        =   16272
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
         Height          =   255
         Index           =   9
         Left            =   10080
         TabIndex        =   22
         Top             =   750
         Width           =   675
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
         Left            =   3450
         TabIndex        =   20
         Top             =   60
         Width           =   3495
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
         Left            =   2340
         TabIndex        =   19
         Top             =   30
         Width           =   795
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
         Left            =   1260
         TabIndex        =   17
         Top             =   30
         Width           =   795
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
         Left            =   12660
         TabIndex        =   16
         Top             =   750
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
         Left            =   11520
         TabIndex        =   14
         Top             =   750
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
         Left            =   13800
         TabIndex        =   12
         Top             =   750
         Width           =   450
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
         Left            =   540
         TabIndex        =   5
         Top             =   465
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   825
         Left            =   9450
         Top             =   180
         Width           =   5130
      End
   End
End
Attribute VB_Name = "frmkeiyakurireki2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chksakujo_Click()

If chksakujo.Value = "1" Then
        If frmkeiyakurireki2.lblend.Caption = "0" Then
            If Trim(lblkojinid.Caption) = "" Then
                logset_2 0, 0, "", 1
            Else
                logset_2 0, 1, Trim(lblkojinid.Caption), 1
            End If
        Else
            If Trim(lblkojinid.Caption) = "" Then
                logset_2 1, 0, "", 1
            Else
                logset_2 1, 1, Trim(lblkojinid.Caption), 1
            End If
        End If
Else
        If frmkeiyakurireki2.lblend.Caption = "0" Then
            If Trim(lblkojinid.Caption) = "" Then
                logset_2 0, 0, "", 0
            Else
                logset_2 0, 1, Trim(lblkojinid.Caption), 0
            End If
        Else
            If Trim(lblkojinid.Caption) = "" Then
                logset_2 1, 0, "", 0
            Else
                logset_2 1, 1, Trim(lblkojinid.Caption), 0
            End If
        End If
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
            logset2_2 0, 0, shitei_moji
    Else
            logset2_2 1, 0, shitei_moji
    End If


End Sub

Private Sub cmdmodoru_Click()

    If frmmain.chkend.Value = 1 Then
            logset_2 0, 0
    Else
            logset_2 1, 0
    End If
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
                    
                    Unload frmkeiyakurireki2
                    Set frmkeiyakurireki2 = Nothing
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

Private Sub Picture11_Click()

Dim h_id As String

gridlog.Col = 0
h_id = Trim(gridlog.Text)

If h_id = "" Then
    Exit Sub
End If

msg = h_id & " ÇÃÉçÉOÇçÌèúÇµÇƒÇ‡ÇÊÇ¢Ç≈Ç∑Ç©ÅH"
res = MsgBox(msg, vbYesNo)

If res = vbYes Then


             sql = "UPDATE log SET sakujo = '1'" & _
                    " WHERE logid = '" & h_id & "'"
        
        On Error GoTo errk3
             db.Execute sql, dbSQLPassThrough
        On Error GoTo 0
        
        If Trim(frmkeiyakurireki2.lblkojinid.Caption) = "" Then
            logset_2 0, 0
        Else
            logset_2 0, 1, frmkeiyakurireki2.lblkojinid.Caption
        End If
        
      

     'Log
        Dim ss_name As String
        Dim ss_id As String
        ss_id = ""
        ss_name = ""
        sagyou_msg = "ÉçÉOè⁄ç◊ÅFÉçÉOID " & h_id & "ÇÃÉçÉOÇçÌèúÇµÇ‹ÇµÇΩÅB"
        log_sagyou sagyou_log_path, sagyou_msg, ss_id, ss_name
End If



Exit Sub

errk3:
MsgBox "çÌèúÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
Exit Sub

End Sub

Private Sub Picture13_Click()
Unload Me
End Sub

Private Sub Picture2_Click()
    If Dir$(sub_exe_path) = vbNullString Then
       MsgBox "ÉçÉOópé¿çsÉtÉ@ÉCÉãÇ™Ç†ÇËÇ‹ÇπÇÒÅB"
       'Exit Sub
    End If
    
Dim h_id As String

gridlog.Col = 0
h_id = Trim(gridlog.Text)

If h_id = "" Then
    Exit Sub
End If

 Dim s_kojinid As String
  
        'é©ìÆé¿çsÅ@à¯êî
        s_kojinid = Trim(lblkojinid.Caption)
        

        Dim sub_exe_path2 As String
        If s_kojinid = "" Then
          sub_exe_path2 = sub_exe_path & " m" & h_id
        Else
          sub_exe_path2 = sub_exe_path & " k" & s_kojinid & " m" & h_id
        End If
        
        If chkwsh.Value = 1 Then
        
                Dim objWshShell As Object
                Dim Ret         As Long
                Set objWshShell = CreateObject("WScript.Shell")
                Ret = objWshShell.Run(sub_exe_path2, 1, True)
                
                If chksakujo.Value = "1" Then
                        If frmkeiyakurireki2.lblend.Caption = "0" Then
                            If Trim(lblkojinid.Caption) = "" Then
                                logset_2 0, 0, "", 1
                            Else
                                logset_2 0, 1, Trim(lblkojinid.Caption), 1
                            End If
                        Else
                            If Trim(lblkojinid.Caption) = "" Then
                                logset_2 1, 0, "", 1
                            Else
                                logset_2 1, 1, Trim(lblkojinid.Caption), 1
                            End If
                        End If
                Else
                        If frmkeiyakurireki2.lblend.Caption = "0" Then
                            If Trim(lblkojinid.Caption) = "" Then
                                logset_2 0, 0, "", 0
                            Else
                                logset_2 0, 1, Trim(lblkojinid.Caption), 0
                            End If
                        Else
                            If Trim(lblkojinid.Caption) = "" Then
                                logset_2 1, 0, "", 0
                            Else
                                logset_2 1, 1, Trim(lblkojinid.Caption), 0
                            End If
                        End If
                End If

    Else


    
 
        
        Dim RetVal
        
        RetVal = Shell(sub_exe_path2, 1)   ' ìdëÏÇé¿çsÇµÇ‹Ç∑ÅB

   End If
   


End Sub

Private Sub Picture7_Click()


    If Dir$(sub_exe_path) = vbNullString Then
       MsgBox "ÉçÉOópé¿çsÉtÉ@ÉCÉãÇ™Ç†ÇËÇ‹ÇπÇÒÅB"
      ' Exit Sub
    End If
    
    
        Dim sub_exe_path2 As String
        Dim s_kojinid As String
  
        'é©ìÆé¿çsÅ@à¯êî
        s_kojinid = Trim(lblkojinid.Caption)
        
        If s_kojinid <> "" Then
            sub_exe_path2 = sub_exe_path & " k" & s_kojinid
        Else
            sub_exe_path2 = sub_exe_path
        End If
        
       ' sub_exe_path2 = "C:\windows\system32\notepad.exe"
        
        'èWåv
    
        If chkwsh.Value = 1 Then
        
                Dim objWshShell As Object
                Dim Ret         As Long
                Set objWshShell = CreateObject("WScript.Shell")
                Ret = objWshShell.Run(sub_exe_path2, 1, True)
                
                If chksakujo.Value = "1" Then
                        If frmkeiyakurireki2.lblend.Caption = "0" Then
                            If Trim(lblkojinid.Caption) = "" Then
                                logset_2 0, 0, "", 1
                            Else
                                logset_2 0, 1, Trim(lblkojinid.Caption), 1
                            End If
                        Else
                            If Trim(lblkojinid.Caption) = "" Then
                                logset_2 1, 0, "", 1
                            Else
                                logset_2 1, 1, Trim(lblkojinid.Caption), 1
                            End If
                        End If
                Else
                        If frmkeiyakurireki2.lblend.Caption = "0" Then
                            If Trim(lblkojinid.Caption) = "" Then
                                logset_2 0, 0, "", 0
                            Else
                                logset_2 0, 1, Trim(lblkojinid.Caption), 0
                            End If
                        Else
                            If Trim(lblkojinid.Caption) = "" Then
                                logset_2 1, 0, "", 0
                            Else
                                logset_2 1, 1, Trim(lblkojinid.Caption), 0
                            End If
                        End If
                End If

    Else

        
        Dim RetVal
        
        RetVal = Shell(sub_exe_path2, 1)   ' ìdëÏÇé¿çsÇµÇ‹Ç∑ÅB

    End If

    'èWåv
    

End Sub
