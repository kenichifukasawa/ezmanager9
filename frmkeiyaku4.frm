VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmkeiyaku4 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "å_ñÒì‡óeÇÃìoò^"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7320
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
      ForeColor       =   &H00C00000&
      Height          =   3645
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   7275
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
         Height          =   2745
         Left            =   180
         TabIndex        =   2
         Top             =   720
         Width           =   6915
         Begin VB.PictureBox Picture13 
            Appearance      =   0  'Ã◊Øƒ
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'Ç»Çµ
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   5460
            MouseIcon       =   "frmkeiyaku4.frx":0000
            MousePointer    =   99  '’∞ªﬁ∞íËã`
            Picture         =   "frmkeiyaku4.frx":08FF
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   21
            Top             =   1770
            Width           =   480
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Ã◊Øƒ
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'Ç»Çµ
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   3960
            MouseIcon       =   "frmkeiyaku4.frx":11C9
            MousePointer    =   99  '’∞ªﬁ∞íËã`
            Picture         =   "frmkeiyaku4.frx":1AC8
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   20
            Top             =   1740
            Width           =   480
         End
         Begin MSComCtl2.UpDown ud1 
            Height          =   375
            Left            =   1995
            TabIndex        =   19
            Top             =   900
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtkeiyakusuu"
            BuddyDispid     =   196618
            OrigLeft        =   2280
            OrigTop         =   840
            OrigRight       =   2520
            OrigBottom      =   1215
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   0   'False
         End
         Begin VB.CommandButton cmdhiduke 
            Caption         =   "...."
            Height          =   375
            Left            =   2760
            TabIndex        =   18
            Top             =   1530
            Width           =   375
         End
         Begin VB.TextBox txtkeiyakukin 
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
            Left            =   1380
            TabIndex        =   9
            Top             =   2130
            Width           =   1335
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0FFFF&
            Caption         =   "â“ìÆèÛãµ"
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
            Height          =   825
            Left            =   3360
            TabIndex        =   5
            Top             =   720
            Width           =   3435
            Begin VB.OptionButton optkadou 
               BackColor       =   &H00C0FFFF&
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
               Height          =   255
               Index           =   2
               Left            =   2460
               TabIndex        =   8
               Top             =   390
               Width           =   795
            End
            Begin VB.OptionButton optkadou 
               BackColor       =   &H00C0FFFF&
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
               Height          =   255
               Index           =   1
               Left            =   1320
               TabIndex        =   7
               Top             =   390
               Width           =   795
            End
            Begin VB.OptionButton optkadou 
               BackColor       =   &H00C0FFFF&
               Caption         =   "â“ìÆ"
               BeginProperty Font 
                  Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   6
               Top             =   390
               Value           =   -1  'True
               Width           =   795
            End
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
            Left            =   1380
            TabIndex        =   4
            Top             =   1530
            Width           =   1335
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
            Left            =   1380
            TabIndex        =   3
            Top             =   900
            Width           =   615
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
            TabIndex        =   23
            Top             =   2280
            Width           =   495
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
            TabIndex        =   22
            Top             =   2280
            Width           =   480
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
            Left            =   2940
            TabIndex        =   16
            Top             =   2250
            Width           =   375
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "å˚"
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
            Left            =   2370
            TabIndex        =   15
            Top             =   1020
            Width           =   375
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
            Left            =   120
            TabIndex        =   14
            Top             =   2190
            Width           =   1215
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0FFFF&
            Caption         =   "äJénì˙ÅF"
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
            TabIndex        =   13
            Top             =   1590
            Width           =   915
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "å_ñÒêîÅF"
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
            TabIndex        =   12
            Top             =   960
            Width           =   915
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
            TabIndex        =   11
            Top             =   300
            Width           =   5295
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
            TabIndex        =   10
            Top             =   300
            Width           =   1275
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'ïsìßñæ
            BorderStyle     =   0  'ìßñæ
            Height          =   975
            Left            =   3360
            Top             =   1620
            Width           =   3420
         End
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
         Left            =   1380
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   1
         Top             =   300
         Width           =   5715
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
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmkeiyaku4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbkeiyakushurui_Click()
Dim k_rs As Recordset

If cmbkeiyakushurui.ListIndex <> -1 Then
    sql = "select*from shurui where shuruiid='" & Mid(cmbkeiyakushurui.Text, 1, 2) & "'"
    If FcSQlGet(k_rs, sql, PrMsg) = False Then
        MsgBox "ó\ä˙ÇπÇ éñóRÇ…ÇÊÇËÉGÉâÅ[Ç™î≠ê∂ÇµÇ‹ÇµÇΩÅB"
        Exit Sub
    End If
    With frmkeiyaku4
        .lblkeiyakuid.Caption = k_rs!shuruiid
        .lblkeiyakusetsumei.Caption = k_rs!setsumei
        .txtkaishibi.Text = Format(Date, "yyyy/mm/dd")
        .txtkeiyakukin.Text = k_rs!tanka
        tanka = k_rs!tanka
        .ud1.Enabled = True
        .txtkeiyakusuu.Text = "1"
        .optkadou(0).Value = True
    End With
    k_rs.Close
End If

End Sub

Private Sub cmdclear_Click()

End Sub

Private Sub cmdhiduke_Click()
donohi = 2
frmhiduke.Show 1
End Sub

Private Sub cmdtsuika_Click()
    
End Sub

Private Sub Form_Load()
    sbshuruiset

End Sub

Private Sub Picture13_Click()
Unload frmkeiyaku4
Set frmkeiyaku4 = Nothing

End Sub

Private Sub Picture7_Click()
Dim tid As String, tkojinid As String, tkuchisuu As String, tkingaku As String
Dim tkaishi As String, tkudou As String, tshurui As String, ttid As String, shime As String
Dim s_name As String

Dim s_kyounohi As String
            
s_kyounohi = Format(Date, "yyyy/mm/dd")
            
            
    If cmbkeiyakushurui.ListIndex <> -1 Then
            If soukatsu(0, "renban", 4, "", 1, 1) = False Then
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
        With frmkeiyaku4
            tshurui = Mid(.cmbkeiyakushurui.Text, 1, 2)
            s_name = .cmbkeiyakushurui.Text
            Select Case tshurui
            Case "01", "96", "97", "98", "99"
                MsgBox "åªç›ÅAÇ±ÇÃå_ñÒÇÕégópÇ≈Ç´Ç‹ÇπÇÒÅB"
                Exit Sub
            End Select
            
            tkojinid = frmtouroku2.lblid.Caption
            'èdï°ÇÃämîF
            sql = "select*from keiyaku where kojinid ='" & tkojinid & "' and shurui ='" & tshurui & "'"
            If FcSQlGet(rs, sql, PrMsg) = True Then
                MsgBox "ìØÇ∂å_ñÒÇ™Ç∑Ç≈Ç…ìoò^Ç≥ÇÍÇƒÇ¢Ç‹Ç∑ÅBå_ñÒé“ID:" & tkojinid
                rs.Close
                Exit Sub
            End If
            
            tkuchisuu = .txtkeiyakusuu.Text
            If Trim(tkuchisuu) = "" Then
                MsgBox "å_ñÒå˚êîÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                .txtkeiyakusuu.SetFocus
                Exit Sub
            End If
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
            If .optkadou(0).Value = True Then
                tkudou = "0"
            ElseIf .optkadou(1).Value = True Then
                tkudou = "1"
            Else
                tkudou = "2"
            End If
            If kenchan = 1 Then
                shime = "0"
            Else
                If frmtouroku2.lbljoukyou.Caption = "ïsâ¬" Then
                    shime = "1"
                Else
                    shime = "0"
                End If
            End If
        End With
        sql = "INSERT INTO keiyaku VALUES('" & tid & "','" & tkojinid & "'," & _
                       "'" & tshurui & "','" & tkaishi & "','" & tkudou & "'," & _
                       "'" & tkuchisuu & "','" & tkingaku & "','" & shime & "','0')"
        
On Error GoTo errk4
        db.Execute sql, dbSQLPassThrough
On Error GoTo 0
        ttid = CStr(CLng(tid) + 1)
        If soukatsu(1, "renban", 4, ttid, 1, 1) = False Then
                MsgBox "ê›íËÇÃèëÇ´çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
                Exit Sub
        End If
        
        keiyakuset (tkojinid)
        Dim sql_shuusei As String

        If tshurui = "01" Then
            sql_shuusei = "update kojin set kajou='1',saishuukoushinbi='" & s_kyounohi & "'" & _
                " where kojinid ='" & tkojinid & "'"
            On Error GoTo errk4
             db.Execute sql_shuusei, dbSQLPassThrough
            On Error GoTo 0
        ElseIf tshurui = "06" Then
            sql_shuusei = "update kojin set kajou='0',saishuukoushinbi='" & s_kyounohi & "'" & _
                " where kojinid ='" & tkojinid & "'"
            On Error GoTo errk4
             db.Execute sql_shuusei, dbSQLPassThrough
            On Error GoTo 0
        Else
            sql_shuusei = "update kojin set saishuukoushinbi='" & s_kyounohi & "'" & _
                " where kojinid ='" & tkojinid & "'"
            On Error GoTo errk4
             db.Execute sql_shuusei, dbSQLPassThrough
            On Error GoTo 0
        End If
        
       
            
        
        If a_shukeiyaku_koushin(tkojinid, tkaishi, tkingaku, tkudou, shime) = False Then
            MsgBox "óøã‡ÇÃçXêVÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
            Exit Sub
        End If
        
        
        
                 'Log
                Dim ss_name As String
                ss_name = frmtouroku2.lblname.Caption
                
                        Dim ss_id As String
        ss_id = frmtouroku2.lblid.Caption

                sagyou_msg = "éÂå_ñÒÅFå_ñÒ ID " & s_name & "ÇÃå_ñÒÇêVãKìoò^ÇµÇ‹ÇµÇΩÅB"
                log_sagyou sagyou_log_path, sagyou_msg, ss_id, ss_name
                
                
        
        MsgBox "ìoò^ÇµÇ‹ÇµÇΩÅB"
        Unload frmkeiyaku4
        Set frmkeiyaku4 = Nothing
    
    
    Else
        MsgBox "å_ñÒéÌóﬁÇëIëÇµÇƒÇ©ÇÁÅAé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        cmbkeiyakushurui.SetFocus
        Exit Sub
    End If
Exit Sub

errk4:
    MsgBox "å_ñÒÇÃìoò^Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
    Exit Sub

End Sub

Private Sub txtkeiyakukin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub

Private Sub txtkeiyakusuu_Change()
txtkeiyakukin.Text = tanka * CInt(txtkeiyakusuu.Text)
End Sub
