VERSION 5.00
Begin VB.Form frmcell 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "ÉZÉãçXêV"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   9690
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8850
      MouseIcon       =   "frmcell.frx":0000
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmcell.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   270
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8880
      MouseIcon       =   "frmcell.frx":11C9
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmcell.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   1410
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "ãåÉZÉãè⁄ç◊"
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
      Height          =   1125
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   8355
      Begin VB.Label Label1 
         Caption         =   "ÉZÉãÇhÇc"
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
         TabIndex        =   9
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lbloldcellid 
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
         Left            =   270
         TabIndex        =   8
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "ÉZÉãñº"
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
      Begin VB.Label lbloldcellname 
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
         TabIndex        =   6
         Top             =   540
         Width           =   4875
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "êVÉZÉãè⁄ç◊"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   1260
      Width           =   8355
      Begin VB.TextBox txtnewcellname 
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
         TabIndex        =   1
         ToolTipText     =   "êVã‡óZã@ä÷ñºÇì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
         Top             =   540
         Width           =   4935
      End
      Begin VB.Label Label7 
         Caption         =   "ÉZÉãÇhÇc"
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
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "ÉZÉãñº"
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
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblnewcellid 
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
         TabIndex        =   2
         Top             =   540
         Width           =   1035
      End
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
      Left            =   8850
      TabIndex        =   10
      Top             =   810
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
      Left            =   8910
      TabIndex        =   13
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   2265
      Left            =   8550
      Top             =   120
      Width           =   1050
   End
End
Attribute VB_Name = "frmcell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Picture13_Click()
Unload Me

End Sub

Private Sub Picture7_Click()
Dim nid As String, nmei As String, nnid As String

With frmcell

nid = Trim(.lblnewcellid.Caption)
nmei = Trim(.txtnewcellname.Text)


If nmei = "" Or nid = "" Then
   MsgBox "äeéñçÄÇïKÇ∏ì¸óÕÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
   Exit Sub
End If


 On Error GoTo JIKKOUERROR2

 If .SSCJIKKOU.Caption = "ìoò^" Then
     
     sql = "INSERT INTO cell VALUES('" & nid & "'," & _
                    "'" & nmei & "')"
     db.Execute sql, dbSQLPassThrough
     nnid = CStr(CLng(nid) + 1)
        If soukatsu(1, "renban", 12, nnid, 1, 1) = False Then
                MsgBox "ê›íËÇÃèëÇ´çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
                Exit Sub
        End If
     MsgBox "ìoò^ÇµÇ‹ÇµÇΩÅB"
 Else
 
     sql = "UPDATE cell SET cellname = '" & nmei & "'" & _
                  " WHERE cellid = '" & nid & "'"
     db.Execute sql, dbSQLPassThrough
    MsgBox "ïœçXÇµÇ‹ÇµÇΩÅB"
   
 End If
On Error GoTo 0
     
Unload frmcell
Set frmcell = Nothing
cellset

End With
Exit Sub

JIKKOUERROR2:
    MsgBox "çXêVÇ…é∏îsÇµÇ‹ÇµÇΩ°"
    Exit Sub


End Sub

Private Sub SSCJIKKOU2_Click()

End Sub

Private Sub SSCSHUURYOU_Click()
End Sub

Private Sub txtnewcellname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Picture7_Click
End If
End Sub
