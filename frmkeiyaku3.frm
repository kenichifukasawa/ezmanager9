VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmkeiyaku3 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "å_ñÒì‡óeïœçXâÊñ "
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   7125
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "å_ñÒê›íËÇÃïœçX"
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
      Height          =   2775
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   6915
      Begin VB.CommandButton cmdsadou 
         Caption         =   "çÏìÆèCê≥"
         Height          =   315
         Left            =   4740
         TabIndex        =   22
         Top             =   2280
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CheckBox chkshime 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ÅY"
         Height          =   345
         Left            =   2790
         TabIndex        =   20
         Top             =   990
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3960
         MouseIcon       =   "frmkeiyaku3.frx":0000
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkeiyaku3.frx":08FF
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   17
         Top             =   1830
         Width           =   480
      End
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5850
         MouseIcon       =   "frmkeiyaku3.frx":11C9
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkeiyaku3.frx":1AC8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   16
         Top             =   1830
         Width           =   480
      End
      Begin MSComCtl2.UpDown ud1 
         Height          =   375
         Left            =   1995
         TabIndex        =   15
         Top             =   930
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
         Enabled         =   -1  'True
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
         IMEMode         =   2  'µÃ
         Left            =   1380
         TabIndex        =   2
         Top             =   2130
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0E0FF&
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
         Height          =   915
         Left            =   3360
         TabIndex        =   3
         Top             =   720
         Width           =   3435
         Begin VB.OptionButton optkadou 
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   7
            Top             =   480
            Width           =   795
         End
         Begin VB.OptionButton optkadou 
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   6
            Top             =   480
            Width           =   795
         End
         Begin VB.OptionButton optkadou 
            BackColor       =   &H00C0E0FF&
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
            TabIndex        =   5
            Top             =   480
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
         TabIndex        =   1
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
         IMEMode         =   2  'µÃ
         Left            =   1380
         TabIndex        =   0
         Top             =   930
         Width           =   615
      End
      Begin VB.Label lblsadou 
         BackColor       =   &H00C0E0FF&
         Height          =   165
         Left            =   5100
         TabIndex        =   21
         Top             =   1920
         Visible         =   0   'False
         Width           =   165
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
         Left            =   3990
         TabIndex        =   19
         Top             =   2340
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
         Left            =   5880
         TabIndex        =   18
         Top             =   2340
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
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
         Left            =   2820
         TabIndex        =   14
         Top             =   2250
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
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
         Left            =   2400
         TabIndex        =   13
         Top             =   1050
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   12
         Top             =   2190
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   11
         Top             =   1590
         Width           =   915
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   10
         Top             =   990
         Width           =   915
      End
      Begin VB.Label lblkeiyaku 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   9
         Top             =   300
         Width           =   5295
      End
      Begin VB.Label lblkeiyakuid 
         Alignment       =   2  'íÜâõëµÇ¶
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   8
         Top             =   300
         Width           =   1275
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   975
         Left            =   3390
         Top             =   1680
         Width           =   3420
      End
   End
End
Attribute VB_Name = "frmkeiyaku3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdend_Click()
End Sub

Private Sub cmdhenkou_Click()
End Sub



Private Sub cmdsadou_Click()
Dim sql_666 As String, sql_6663 As String, sql_6662 As String, souda2 As String
Dim rs_666 As Recordset, rs_6662 As Recordset, souda As String, ssid As String

    ssid = lblkeiyakuid.Caption
    souda = lblsadou.Caption
    If souda = "0" Then
        souda2 = "1"
    Else
        souda2 = "0"
    End If
     sql_666 = "update keiyaku set sadou ='" & souda2 & "'" & _
                " where keiyakuid ='" & ssid & "'"
    On Error GoTo errko1
     db.Execute sql_666, dbSQLPassThrough
    On Error GoTo 0
    
    MsgBox "OK"
    
     lblsadou.Caption = souda2
Exit Sub
errko1:
    MsgBox "err"
    Exit Sub

End Sub

Private Sub Form_Load()
If kenchan = 0 Then
    chkshime.Visible = False
Else
    chkshime.Visible = True
End If
End Sub

Private Sub Picture13_Click()
    Unload frmkeiyaku3
    Set frmkeiyaku3 = Nothing

End Sub

Private Sub Picture7_Click()


Dim hid As String, hkuchisuu As String, hkingaku As String, hkaishi As String
Dim hkudou As String, newshime As String, newsadou As String
Dim s_name As String

    With frmkeiyaku3
        hid = .lblkeiyakuid.Caption
        s_name = .lblkeiyaku.Caption
        hkuchisuu = .txtkeiyakusuu.Text
        If Trim(hkuchisuu) = "" Then
            MsgBox "å_ñÒå˚êîÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            .txtkeiyakusuu.SetFocus
            Exit Sub
        End If
        hkingaku = .txtkeiyakukin.Text
        If Trim(hkingaku) = "" Then
            MsgBox "å_ñÒã‡äzÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            .txtkeiyakukin.SetFocus
            Exit Sub
        End If
        hkaishi = Format(.txtkaishibi.Text, "yyyymmdd")
        If Trim(hkaishi) = "" Then
            MsgBox "å_ñÒäJénì˙Çì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            .txtkaishibi.SetFocus
            Exit Sub
        End If
        newsadou = "1"
        If .optkadou(0).Value = True Then
            newsadou = "0"
            hkudou = "0"
        ElseIf .optkadou(1).Value = True Then
            hkudou = "1"
        Else
            hkudou = "2"
        End If
    End With
    If chkshime.Visible = True Then
        If chkshime.Value = 1 Then
            newshime = "1"
        Else
            newshime = "0"
        End If
        sql = "UPDATE keiyaku SET kuchisuu = '" & hkuchisuu & "'," & _
            "kingaku = '" & hkingaku & "',kaishi = '" & hkaishi & "',kadou = '" & hkudou & "',shime = '" & newshime & "',sadou = '" & newsadou & "'" & _
            " WHERE keiyakuid = '" & hid & "'"
    Else
        sql = "UPDATE keiyaku SET kuchisuu = '" & hkuchisuu & "'," & _
               "kingaku = '" & hkingaku & "',kaishi = '" & hkaishi & "',kadou = '" & hkudou & "',sadou = '" & newsadou & "'" & _
               " WHERE keiyakuid = '" & hid & "'"
    End If
On Error GoTo errk3
     db.Execute sql, dbSQLPassThrough
On Error GoTo 0

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
            
            

    keiyakuset (frmtouroku2.lblid.Caption)
        
    If henkoumae_shurui <> hkudou Then
            If a_shukeiyaku_henkou(frmtouroku2.lblid.Caption, henkoumae_kingaku, hkingaku, henkoumae_shurui, hkudou) = False Then
                MsgBox "óøã‡ÇÃçXêVÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
                Exit Sub
            End If
    Else
        If henkoumae_kingaku <> hkingaku Then
            If frmtouroku2.txtshukeiyakuseikyuu.Text = "0" And frmtouroku2.txtsoukei.Text = "0" And frmtouroku2.txtjikaiseikyuubi.Text <> "" Then
            
            Else
                If a_shukeiyaku_henkou(frmtouroku2.lblid.Caption, henkoumae_kingaku, hkingaku, henkoumae_shurui, hkudou) = False Then
                    MsgBox "óøã‡ÇÃçXêVÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
                    Exit Sub
                End If
            End If
        End If
    End If
    
            'Log
        Dim ss_name As String
        ss_name = frmtouroku2.lblname.Caption
        Dim ss_id As String
        ss_id = frmtouroku2.lblid.Caption

        sagyou_msg = "éÂå_ñÒÅFå_ñÒ ID " & hid & " " & s_name & "ÇÃå_ñÒÇïœçXÇµÇ‹ÇµÇΩÅB"
        log_sagyou sagyou_log_path, sagyou_msg, ss_id, ss_name
        
        
    
    MsgBox "ïœçXÇµÇ‹ÇµÇΩÅB"
    
    Unload frmkeiyaku3
    Set frmkeiyaku3 = Nothing
Exit Sub

errk3:
    MsgBox "å_ñÒÇÃçXêVÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
    Exit Sub
    
errsakujo2:
    MsgBox "ç≈èIå_ñÒì˙ÇÃçXêVÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
    Exit Sub

End Sub

Private Sub txtkeiyakukin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub

Private Sub txtkeiyakusuu_Change()
txtkeiyakukin = tanka * CInt(txtkeiyakusuu.Text)
End Sub

