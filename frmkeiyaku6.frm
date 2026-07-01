VERSION 5.00
Begin VB.Form frmkeiyaku6 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "ÇªÇÃëºÇÃå_ñÒÇÃïœçX"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   7260
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
      Height          =   1815
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   6915
      Begin VB.TextBox txthi 
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
         Height          =   315
         IMEMode         =   7  'ëSäpâpêî
         Left            =   5100
         TabIndex        =   16
         Top             =   540
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.PictureBox pseikyu 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   4980
         MouseIcon       =   "frmkeiyaku6.frx":0000
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkeiyaku6.frx":08FF
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   13
         ToolTipText     =   "ÇªÇÃëºÇÃå_ñÒÇÃÇ›Çì¸ã‡èàóùÇµÇ‹Ç∑ÅB"
         Top             =   870
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.CheckBox chkshime 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ÅY"
         Height          =   345
         Left            =   2910
         TabIndex        =   12
         Top             =   810
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
         Left            =   3990
         MouseIcon       =   "frmkeiyaku6.frx":11C9
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkeiyaku6.frx":1AC8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         Top             =   870
         Width           =   480
      End
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Ã◊Øƒ
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Ç»Çµ
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   6000
         MouseIcon       =   "frmkeiyaku6.frx":2392
         MousePointer    =   99  '’∞ªﬁ∞íËã`
         Picture         =   "frmkeiyaku6.frx":2C91
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   8
         Top             =   870
         Width           =   480
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
         TabIndex        =   2
         Top             =   780
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
         Left            =   1440
         TabIndex        =   1
         Top             =   1260
         Width           =   1335
      End
      Begin VB.Label lblnyukin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ì¸ã‡ì˙ÅF"
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
         Left            =   4140
         TabIndex        =   17
         Top             =   540
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblsonotaid 
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
         Left            =   3180
         TabIndex        =   15
         Top             =   1020
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblseikyu 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ì¸ã‡"
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
         Left            =   5010
         TabIndex        =   14
         Top             =   1380
         Visible         =   0   'False
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
         Left            =   4020
         TabIndex        =   11
         Top             =   1380
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
         Left            =   6030
         TabIndex        =   10
         Top             =   1380
         Width           =   495
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
         TabIndex        =   7
         Top             =   300
         Width           =   1275
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
         TabIndex        =   6
         Top             =   300
         Width           =   5295
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   5
         Top             =   840
         Width           =   915
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
         Left            =   180
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
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
         Left            =   3000
         TabIndex        =   3
         Top             =   1380
         Width           =   375
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'ïsìßñæ
         BorderStyle     =   0  'ìßñæ
         Height          =   975
         Left            =   3420
         Top             =   720
         Width           =   3420
      End
   End
End
Attribute VB_Name = "frmkeiyaku6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdend_Click()

End Sub

Private Sub cmdhenkou_Click()

End Sub

Private Sub Form_Load()
If kenchan = 0 Then
    chkshime.Visible = False
    lblseikyu.Visible = False
    pseikyu.Visible = False
    txthi.Visible = False
    lblnyukin.Visible = False
Else
    chkshime.Visible = True
      lblseikyu.Visible = True
    pseikyu.Visible = True
    txthi.Visible = True
    txthi.Text = Format(Date, "yyyy/mm/dd")
    lblnyukin.Visible = True
End If

End Sub

Private Sub Picture13_Click()
    Unload frmkeiyaku6
    Set frmkeiyaku6 = Nothing

End Sub

Private Sub Picture7_Click()
Dim hid As String, hkingaku As String, hkaishi As String
Dim hkudou As String, newshime As String
Dim s_name As String




    With frmkeiyaku6
        hid = .lblkeiyakuid.Caption
        s_name = .lblkeiyaku.Caption
        
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
        hkudou = frmtouroku2.lbljoukyou.Caption
    End With
    If chkshime.Visible = True Then
        If chkshime.Value = 1 Then
            newshime = "1"
        Else
            newshime = "0"
        End If
        sql = "UPDATE keiyaku2 SET " & _
               "kingaku = '" & hkingaku & "',kaishi = '" & hkaishi & "',f2 = '" & newshime & "'" & _
               " WHERE keiyakuid2 = '" & hid & "'"
    Else
        sql = "UPDATE keiyaku2 SET " & _
               "kingaku = '" & hkingaku & "',kaishi = '" & hkaishi & "'" & _
               " WHERE keiyakuid2 = '" & hid & "'"
    End If
On Error GoTo errk6
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
                
                
    keiyakuset2 (frmtouroku2.lblid.Caption)

        
    If a_sonota_henkou(frmtouroku2.lblid.Caption, henkoumae_kingaku, hkingaku) = False Then
        MsgBox "óøã‡ÇÃçXêVÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
        Exit Sub
    End If
    
    'Log
    Dim ss_name As String
    ss_name = frmtouroku2.lblname.Caption
     Dim ss_id As String
        ss_id = frmtouroku2.lblid.Caption
    sagyou_msg = "ÇªÇÃëºÇÃå_ñÒÅFå_ñÒ ID " & hid & " " & s_name & "ÇÃå_ñÒÇïœçXÇµÇ‹ÇµÇΩÅB"
    log_sagyou sagyou_log_path, sagyou_msg, ss_id, ss_name
                
    
    MsgBox "ïœçXÇµÇ‹ÇµÇΩÅB"
    
    Unload frmkeiyaku6
    Set frmkeiyaku6 = Nothing
Exit Sub

errk6:
    MsgBox "å_ñÒÇÃçXêVÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
    Exit Sub

errsakujo2:
    MsgBox "ç≈èIçXêVì˙ÇÃçXêVÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
    Exit Sub


End Sub

Private Sub pseikyu_Click()
Dim newkojinid As String, newseikyuubi As String, newshurui As String
Dim newkin As String, newbankid As String, newnyukinbi As String, newkeiyaku2id As String

newkeiyaku2id = lblkeiyakuid.Caption
newkojinid = frmtouroku2.lblid.Caption

If Len(txthi.Text) <> 10 Then
    MsgBox "ì¸ã‡ì˙ÇÕÅu2013/01/01ÅvÇÃÇÊÇ§Ç…ÇPÇOåÖÇ≈ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
    Exit Sub
End If
newnyukinbi = Format(txthi.Text, "yyyymmdd")
newseikyuubi = Mid(newnyukinbi, 1, 6)
newshurui = lblsonotaid.Caption
newkin = txtkeiyakukin.Text
newbankid = Space(1)

 If shousai_main_set(newkojinid, newseikyuubi, "2", newshurui, newkin, newbankid, "1") = -1 Then
            MsgBox "å_ñÒÇÃçXêVÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
            Exit Sub
  End If


sql = "UPDATE shousai SET  keinyu = '" & newnyukinbi & "',keist='1',bankid='" & Space(1) & "'" & _
            " WHERE shousaiid = '" & shitei_shousaiid & "'"

     db2.Execute sql, dbSQLPassThrough

sql = "update keiyaku2 set f2='2' where keiyakuid2 ='" & _
                     newkeiyaku2id & "'"
                     

    db.Execute sql, dbSQLPassThrough

If a_genkin_nyuukin3(newkojinid, newnyukinbi, CStr(newkin), "1") = False Then
    Exit Sub
End If
                                        
If hyouji_touroku(newkojinid) = False Then
    Exit Sub
End If

keiyakuset2 (newkojinid)

Unload Me


                
End Sub

Private Sub txtkeiyakukin_GotFocus()
txtkeiyakukin.SelStart = 0
txtkeiyakukin.SelLength = Len(txtkeiyakukin.Text)
End Sub

Private Sub txtkeiyakukin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If
End Sub
