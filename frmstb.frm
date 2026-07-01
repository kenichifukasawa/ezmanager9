VERSION 5.00
Begin VB.Form frmstb 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "STBçXêVâÊñ "
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6165
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.Frame Frame3 
      Caption         =   "STB"
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
      Height          =   2145
      Left            =   120
      TabIndex        =   9
      Top             =   90
      Width           =   5865
      Begin VB.CheckBox chkc 
         Caption         =   "PPVÇ†ÇË"
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
         Left            =   4500
         TabIndex        =   4
         Top             =   1650
         Width           =   1200
      End
      Begin VB.CheckBox chkb 
         Caption         =   "PPVÇ†ÇË"
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
         Left            =   4500
         TabIndex        =   2
         Top             =   1230
         Width           =   1200
      End
      Begin VB.TextBox txtccasno 
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
         Left            =   1530
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1590
         Width           =   2700
      End
      Begin VB.TextBox txtbcasno 
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
         Left            =   1530
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1160
         Width           =   2700
      End
      Begin VB.TextBox txtstbno 
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
         Left            =   1530
         MaxLength       =   12
         TabIndex        =   0
         Top             =   730
         Width           =   2300
      End
      Begin VB.TextBox txtstbkanriid 
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
         Left            =   1530
         TabIndex        =   10
         Top             =   300
         Width           =   2300
      End
      Begin VB.Label lblkojinid 
         Height          =   225
         Left            =   4560
         TabIndex        =   15
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "CCASÅ@NOÅF"
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
         TabIndex        =   14
         Top             =   1650
         Width           =   1635
      End
      Begin VB.Label Label13 
         Caption         =   "STBä«óùIDÅF"
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
         TabIndex        =   13
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label Label14 
         Caption         =   "STBÅ@NOÅF"
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
         TabIndex        =   12
         Top             =   790
         Width           =   1635
      End
      Begin VB.Label Label15 
         Caption         =   "BCASÅ@NOÅF"
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
         TabIndex        =   11
         Top             =   1220
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   5340
      MouseIcon       =   "frmstb.frx":0000
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmstb.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   2430
      Width           =   480
   End
   Begin VB.PictureBox Picture16 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   4500
      MouseIcon       =   "frmstb.frx":11C9
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmstb.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   2430
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
      Left            =   5400
      TabIndex        =   8
      Top             =   2910
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Index           =   15
      Left            =   4530
      TabIndex        =   7
      Top             =   2940
      Width           =   90
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   855
      Left            =   4260
      Top             =   2370
      Width           =   1740
   End
End
Attribute VB_Name = "frmstb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub Picture13_Click()
Unload Me
End Sub

Private Sub Picture16_Click()

Dim newidid As String, newstbno As String, newbcasno As String, newccasno As String, newkojinid As String
Dim nnid As String, newbbasumu As String, newccasumu As String
    
With frmstb
    newidid = Trim(.txtstbkanriid.Text)
    newstbno = Trim(.txtstbno.Text)
    If newstbno = "" Then
        newstbno = Space(1)
    End If
    newbcasno = Trim(.txtbcasno.Text)
    If newbcasno = "" Then
        newbcasno = Space(1)
    End If
    newccasno = Trim(.txtccasno.Text)
    If newccasno = "" Then
        newccasno = Space(1)
    End If
    If .chkb.Value = 1 Then
        newbbasumu = "1"
    Else
        newbbasumu = "0"
    End If
    If .chkc.Value = 1 Then
        newccasumu = "1"
    Else
        newccasumu = "0"
    End If
        
    newkojinid = .lblkojinid.Caption
End With
On Error GoTo errstb
If Label1(15).Caption = "ìoò^" Then
     sql = "INSERT INTO stb VALUES('" & newidid & "'," & _
                    "'" & newkojinid & "','" & newstbno & "','" & newbcasno & "'" & _
                    ",'" & newbbasumu & "','" & newccasno & "','" & newccasumu & "')"
     db.Execute sql, dbSQLPassThrough
     nnid = CStr(CLng(newidid) + 1)
        If soukatsu(1, "renban", 11, nnid, 1, 1) = False Then
                MsgBox "ê›íËÇÃèëÇ´çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
                Exit Sub
        End If
     MsgBox "ìoò^ÇµÇ‹ÇµÇΩÅB"
 Else
 
     sql = "UPDATE stb SET koushinbistbid = '" & newstbno & "'," & _
                  "bcasno = '" & newbcasno & "',bcasumu = '" & newbbasumu & "'," & _
                  "ccasno = '" & newccasno & "',ccasumu = '" & newccasumu & "'" & _
                  " WHERE stbkanriid = '" & newidid & "'"
     db.Execute sql, dbSQLPassThrough
    MsgBox "ïœçXÇµÇ‹ÇµÇΩÅB"
   

End If
On Error GoTo 0
Unload Me
stb_grid_set2 (frmtouroku2.lblid.Caption)

Exit Sub

errstb:
    MsgBox "STBÇÃçXêVÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
    Unload Me
End Sub

Private Sub txtbcasno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtccasno.SetFocus
End If

End Sub

Private Sub txtstbno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtbcasno.SetFocus
End If
End Sub
