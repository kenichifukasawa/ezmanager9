VERSION 5.00
Begin VB.Form frmkeiyaku5 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "ÇªÇÃëºÇÃå_ñÒÇÃìoò^"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   7440
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "å_ñÒçXêV"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2625
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   7275
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
         Left            =   1380
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   0
         Top             =   300
         Width           =   5715
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "å_ñÒê›íË"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1875
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   6915
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Ã◊Øƒ
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'Ç»Çµ
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   3930
            MouseIcon       =   "frmkeiyaku5.frx":0000
            MousePointer    =   99  '’∞ªﬁ∞íËã`
            Picture         =   "frmkeiyaku5.frx":08FF
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   13
            Top             =   900
            Width           =   480
         End
         Begin VB.PictureBox Picture13 
            Appearance      =   0  'Ã◊Øƒ
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'Ç»Çµ
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   5460
            MouseIcon       =   "frmkeiyaku5.frx":11C9
            MousePointer    =   99  '’∞ªﬁ∞íËã`
            Picture         =   "frmkeiyaku5.frx":1AC8
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   12
            Top             =   900
            Width           =   480
         End
         Begin VB.CommandButton cmdhiduke 
            Caption         =   "...."
            Height          =   375
            Left            =   2880
            TabIndex        =   11
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox txtkaishibi 
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
            Left            =   1440
            TabIndex        =   1
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txtkeiyakukin 
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
            Height          =   375
            IMEMode         =   2  'µÃ
            Left            =   1440
            TabIndex        =   2
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label SSCJIKKOU 
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
            Left            =   3960
            TabIndex        =   15
            Top             =   1410
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
            Index           =   13
            Left            =   5490
            TabIndex        =   14
            Top             =   1410
            Width           =   495
         End
         Begin VB.Label lblkeiyakuid 
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
            Height          =   375
            Left            =   180
            TabIndex        =   9
            Top             =   300
            Width           =   1275
         End
         Begin VB.Label lblkeiyakusetsumei 
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
            Height          =   375
            Left            =   1500
            TabIndex        =   8
            Top             =   300
            Width           =   5295
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0FFFF&
            Caption         =   "çÏê¨ì˙ÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   180
            TabIndex        =   7
            Top             =   900
            Width           =   915
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "å_ñÒã‡äzÅF"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   6
            Top             =   1380
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "â~"
            BeginProperty Font 
               Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   3000
            TabIndex        =   5
            Top             =   1440
            Width           =   375
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'ïsìßñæ
            BorderStyle     =   0  'ìßñæ
            Height          =   975
            Left            =   3360
            Top             =   750
            Width           =   3420
         End
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "å_ñÒéÌóﬁÅF"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmkeiyaku5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbsonotakeiyaku_Click()
Dim k_rs As Recordset

If cmbsonotakeiyaku.ListIndex <> -1 Then
    sql = "select*from sonota where sonotaid='" & Mid(cmbsonotakeiyaku.Text, 1, 2) & "'"
    If FcSQlGet(k_rs, sql, PrMsg) = False Then
        MsgBox "ó\ä˙ÇπÇ éñóRÇ…ÇÊÇËÉGÉâÅ[Ç™î≠ê∂ÇµÇ‹ÇµÇΩÅB"
        Exit Sub
    End If
    With frmkeiyaku5
        .lblkeiyakuid.Caption = k_rs!sonotaid
        .txtkaishibi.Text = Format(Date, "yyyy/mm/dd")
        If k_rs!kotei = "0" Then
            .lblkeiyakusetsumei.Caption = "Ç±ÇÃÇªÇÃëºÇÃå_ñÒÇÃã‡äzÇÕå≈íËÇ≈Ç∑"
            .txtkeiyakukin.Text = k_rs!kingaku
            .txtkeiyakukin.Enabled = False
        Else
            .lblkeiyakusetsumei.Caption = "Ç±ÇÃÇªÇÃëºÇÃå_ñÒÇÃã‡äzÇÕïœìÆÇ≈Ç∑"
            .txtkeiyakukin.Text = ""
            .txtkeiyakukin.Enabled = True
        End If
    End With
    k_rs.Close
End If

End Sub

Private Sub cmdclear_Click()

End Sub

Private Sub cmdhiduke_Click()
donohi = 3
frmhiduke.Show 1
End Sub

Private Sub cmdtsuika_Click()
End Sub

Private Sub Form_Load()
sbsonotaset
End Sub

Private Sub Picture13_Click()
Unload frmkeiyaku5
Set frmkeiyaku5 = Nothing

End Sub

Private Sub Picture7_Click()
Dim tid As String, tkojinid As String, tkingaku As String, jurs As Recordset
Dim tkaishi As String, tkudou As String, tshurui As String, ttid As String
Dim s_name As String
    
    If cmbsonotakeiyaku.ListIndex <> -1 Then
            If soukatsu(0, "renban", 6, "", 1, 1) = False Then
                    MsgBox "ê›íËÇÃì«Ç›çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
            End If
            If reg = "0" Then
                reg = ""
            End If
            If reg = vbNullString Then
                tid = "0000001"
            Else
                tid = Format$(reg, "000000#")
            End If
        With frmkeiyaku5
            tshurui = Mid(.cmbsonotakeiyaku.Text, 1, 2)
            s_name = .cmbsonotakeiyaku.Text
            tkojinid = frmtouroku2.lblid.Caption
            tkingaku = .txtkeiyakukin.Text
            If Trim(tkingaku) = "" Then
                MsgBox "å_ñÒã‡äzÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                .txtkeiyakukin.SetFocus
                Exit Sub
            End If
            tkaishi = Format(.txtkaishibi.Text, "yyyymmdd")
            If Trim(tkaishi) = "" Then
                MsgBox "å_ñÒäJénì˙Çì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                .txtkaishibi.SetFocus
                Exit Sub
            End If
            If kenchan = 1 Then
                tkudou = "0"
            Else
                If frmtouroku2.lbljoukyou.Caption = "â¬î\" Then
                    tkudou = "0"
                ElseIf frmtouroku2.lbljoukyou.Caption = "ïsâ¬" Then
                    tkudou = "1"
                End If
            End If
       End With
        sql = "INSERT INTO keiyaku2 VALUES('" & tid & "','" & tkojinid & "'," & _
                       "'" & tshurui & "','" & tkaishi & "','" & tkudou & "','" & tkingaku & "')"
On Error GoTo errk5
        db.Execute sql, dbSQLPassThrough
On Error GoTo 0
        ttid = CStr(CLng(tid) + 1)
        If soukatsu(1, "renban", 6, ttid, 1, 1) = False Then
                MsgBox "ê›íËÇÃèëÇ´çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
                Exit Sub
        End If
        
         'ç≈èIçXêVì˙çXêV
            
                Dim s_kojinid As String
    
                s_kojinid = frmtouroku2.lblid.Caption
                
                Dim s_kyounohi As String
                
                s_kyounohi = Format(Date, "yyyy/mm/dd")
    
                 sql = "update kojin set  saishuukoushinbi='" & s_kyounohi & "'" & _
                           " where kojinid ='" & s_kojinid & "'"
                
                On Error GoTo errsakujo2
                db.Execute sql, dbSQLPassThrough
                On Error GoTo 0
                
                
        keiyakuset2 (tkojinid)
        
        
        If a_sonota_touroku(tkojinid, tkingaku, tkudou) = False Then
            MsgBox "óøã‡ÇÃçXêVÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
        End If
        
    'Log
    Dim ss_name As String
    ss_name = frmtouroku2.lblname.Caption
    Dim ss_id As String
        ss_id = frmtouroku2.lblid.Caption
    sagyou_msg = "ÇªÇÃëºÇÃå_ñÒÅFå_ñÒ ID " & s_name & "ÇÃå_ñÒÇêVãKìoò^ÇµÇ‹ÇµÇΩÅB"
    log_sagyou sagyou_log_path, sagyou_msg, ss_id, ss_name
        
        
        MsgBox "ìoò^ÇµÇ‹ÇµÇΩÅB"
        Unload frmkeiyaku5
        Set frmkeiyaku5 = Nothing
    
    
    Else
        MsgBox "å_ñÒéÌóﬁÇëIëÇµÇƒÇ©ÇÁÅAé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        cmbsonotakeiyaku.SetFocus
        Exit Sub
    End If
Exit Sub

errk5:
    MsgBox "å_ñÒÇÃìoò^Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
    Exit Sub
errsakujo2:
    MsgBox "ç≈èIçXêVì˙ÇÃìoò^Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
    Exit Sub



End Sub

Private Sub txtkeiyakukin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub
