VERSION 5.00
Begin VB.Form frmbank2 
   BorderStyle     =   3  'å≈íË¿ﬁ≤±€∏ﬁ
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8880
      MouseIcon       =   "frmbank2.frx":0000
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmbank2.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   20
      Top             =   2490
      Width           =   480
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8850
      MouseIcon       =   "frmbank2.frx":11C9
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmbank2.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   19
      Top             =   1260
      Width           =   480
   End
   Begin VB.Frame Frame2 
      Caption         =   "êVã‡óZã@ä÷è⁄ç◊"
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
      ForeColor       =   &H00000000&
      Height          =   1635
      Left            =   120
      TabIndex        =   11
      Top             =   1860
      Width           =   8355
      Begin VB.ComboBox txtnewbankcode 
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
         Left            =   1920
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   17
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtnewbankmei 
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
         MaxLength       =   40
         TabIndex        =   0
         ToolTipText     =   "êVã‡óZã@ä÷ñºÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
         Top             =   540
         Width           =   4935
      End
      Begin VB.TextBox txtnewshitencode 
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
         Left            =   6900
         MaxLength       =   7
         TabIndex        =   1
         ToolTipText     =   "êVéxìXÉRÅ[ÉhÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label lblnewbankid 
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
         Left            =   240
         TabIndex        =   16
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "ã‡óZã@ä÷ÉRÅ[Éh"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1140
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "ã‡óZã@ä÷éxìXñº"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "ã‡óZã@ä÷î‘çÜ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "ã‡óZã@ä÷éxìXÉRÅ[Éh"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4800
         TabIndex        =   12
         Top             =   1140
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ãåã‡óZã@ä÷è⁄ç◊"
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
      ForeColor       =   &H00000000&
      Height          =   1635
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   8355
      Begin VB.Label lbloldbankcode 
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
         Left            =   1920
         TabIndex        =   10
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "ã‡óZã@ä÷ÉRÅ[Éh"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   300
         TabIndex        =   9
         Top             =   1140
         Width           =   1515
      End
      Begin VB.Label lbloldbankmei 
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
         Left            =   1920
         TabIndex        =   8
         Top             =   540
         Width           =   4875
      End
      Begin VB.Label Label3 
         Caption         =   "ã‡óZã@ä÷éxìXñº"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2220
         TabIndex        =   7
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label lbloldbankid 
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
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "ã‡óZã@ä÷î‘çÜ"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
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
         Left            =   180
         TabIndex        =   5
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label6 
         Caption         =   "ã‡óZã@ä÷éxìXÉRÅ[Éh"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4740
         TabIndex        =   4
         Top             =   1140
         Width           =   2055
      End
      Begin VB.Label lbloldshitencode 
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
         Left            =   6900
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
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
      Index           =   13
      Left            =   8910
      TabIndex        =   18
      Top             =   3000
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
      Left            =   8880
      TabIndex        =   21
      Top             =   1770
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   3375
      Left            =   8580
      Top             =   180
      Width           =   1050
   End
End
Attribute VB_Name = "frmbank2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
sbnbankset2
End Sub

Private Sub Picture13_Click()
    Unload frmbank2
    Set frmbank2 = Nothing
    frmbank.Show 1

End Sub

Private Sub Picture7_Click()
Dim nid As String, nmei As String, nbc As String, nsc As String, nnid As String

With frmbank2

nid = Trim(.lblnewbankid.Caption)
nmei = Trim(.txtnewbankmei.Text)
nbc = Mid(.txtnewbankcode.Text, 1, 4)
nsc = Trim(.txtnewshitencode.Text)

If nmei = "" Or nbc = "" Or nsc = "" Then
   MsgBox "äeéñçÄÇïKÇ∏ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
   Exit Sub
End If


 On Error GoTo JIKKOUERROR2

 If .SSCJIKKOU.Caption = "ìoò^" Then
     
     sql = "INSERT INTO bank VALUES('" & nid & "'," & _
                    "'" & nmei & "','" & nbc & "','" & nsc & "')"
     db.Execute sql, dbSQLPassThrough
     nnid = CStr(CLng(nid) + 1)
        If soukatsu(1, "renban", 3, nnid, 1, 1) = False Then
                MsgBox "ê›íËÇÃèëÇ´çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
                Exit Sub
        End If
     MsgBox "ìoò^ÇµÇ‹ÇµÇΩÅB"
 Else
 
     sql = "UPDATE bank SET bankmei = '" & nmei & "'," & _
                  "bankcode = '" & nbc & "',shitencode = '" & nsc & "'" & _
                  " WHERE bankid = '" & nid & "'"
     db.Execute sql, dbSQLPassThrough
    MsgBox "ïœçXÇµÇ‹ÇµÇΩÅB"
   
 End If
On Error GoTo 0
     
Unload frmbank2
Set frmbank2 = Nothing
bankset
frmbank.Show 1

End With
Exit Sub

JIKKOUERROR2:
    MsgBox "çXêVÇ…é∏îsÇµÇ‹ÇµÇΩ°"
    Exit Sub

End Sub

Private Sub txtnewbankcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtnewshitencode.SetFocus
End If

End Sub

Private Sub txtnewbankmei_GotFocus()
txtnewbankmei.SelStart = 0
txtnewbankmei.SelLength = Len(txtnewbankmei.Text)

End Sub

Private Sub txtnewbankmei_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtnewbankcode.SetFocus
End If
End Sub

Private Sub txtnewshitencode_GotFocus()
txtnewshitencode.SelStart = 0
txtnewshitencode.SelLength = Len(txtnewshitencode.Text)

End Sub

Private Sub txtnewshitencode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If

End Sub
