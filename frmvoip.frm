VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmvoip 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "VoIP"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   15270
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.Frame Frame3 
      Caption         =   "êøãÅè⁄ç◊ÉfÅ[É^"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5775
      Left            =   120
      TabIndex        =   10
      Top             =   1350
      Width           =   15075
      Begin MSFlexGridLib.MSFlexGrid gridshousai 
         Height          =   5265
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   9287
         _Version        =   393216
         Cols            =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "êøãÅÉfÅ[É^"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1155
      Left            =   90
      TabIndex        =   9
      Top             =   90
      Width           =   9645
      Begin MSFlexGridLib.MSFlexGrid gridseikyuu 
         Height          =   765
         Left            =   150
         TabIndex        =   11
         Top             =   300
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   1349
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   13290
      MouseIcon       =   "frmvoip.frx":0000
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmvoip.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      ToolTipText     =   "ì¸ã‡óöóÇï\é¶ÇµÇ‹Ç∑ÅB"
      Top             =   330
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   14310
      MouseIcon       =   "frmvoip.frx":1741
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmvoip.frx":2040
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   390
      Width           =   480
   End
   Begin VB.Frame Frame2 
      Caption         =   "ä˙ä‘éwíË(êøãÅîNåéÇéwíË)"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   885
      Index           =   1
      Left            =   9840
      TabIndex        =   0
      Top             =   240
      Width           =   2715
      Begin VB.ComboBox cmbtsuki 
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
         Left            =   1410
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   2
         Top             =   300
         Width           =   675
      End
      Begin VB.ComboBox cmbnen 
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
         Left            =   150
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   1
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "åé"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2130
         TabIndex        =   4
         Top             =   420
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "îN"
         BeginProperty Font 
            Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1170
         TabIndex        =   3
         Top             =   420
         Width           =   255
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FFFFFF&
      Caption         =   "íäèo"
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
      Index           =   0
      Left            =   13050
      TabIndex        =   8
      Top             =   870
      Width           =   975
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
      Left            =   14340
      TabIndex        =   6
      Top             =   870
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   855
      Left            =   12750
      Top             =   300
      Width           =   2490
   End
End
Attribute VB_Name = "frmvoip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub seikyuu_voip_set2(kikan As String, shitei_code As String)
Dim rirekirs As Recordset, rirekisuu As Long



sql = "SELECT *" & _
    " FROM shousai" & _
    " WHERE nentsuki='" & kikan & "' and code ='" & shitei_code & "'"
    
If FcSQlGet_voip(rirekirs, sql, PrMsg) = False Then
    With frmvoip.gridshousai
        .Font.Size = 11
        .Rows = 1
        .ColWidth(0) = 1200
        .ColWidth(1) = 1300
        .ColWidth(2) = 1900
        .ColWidth(3) = 2500
        .ColWidth(4) = 1000
        .Row = 0
        .Col = 0
        .Text = "ÇhÇc"
        .Col = 1
        .Text = "ÉRÅ[Éh"
        .Col = 2
        .Text = "çÄñ⁄"
        .Col = 3
        .Text = "çÄñ⁄ÇQ"
        .Col = 4
        .Text = "ã‡äz"
    End With
    Exit Sub
Else
    rirekirs.MoveLast
    rirekisuu = rirekirs.RecordCount
    With frmvoip.gridshousai
        .Font.Size = 11
        .Rows = rirekisuu + 1
        .ColWidth(0) = 1200
        .ColWidth(1) = 1300
        .ColWidth(2) = 1900
        .ColWidth(3) = 2500
        .ColWidth(4) = 1000
        .Row = 0
        .Col = 0
        .Text = "ÇhÇc"
        .Col = 1
        .Text = "ÉRÅ[Éh"
        .Col = 2
        .Text = "çÄñ⁄"
        .Col = 3
        .Text = "çÄñ⁄ÇQ"
        .Col = 4
        .Text = "ã‡äz"
    End With
    rirekirs.MoveFirst
    For i = 1 To rirekisuu
        With frmvoip.gridshousai
            .Row = i
            .Col = 0
            .Text = rirekirs!shousaiid
            .Col = 1
            .Text = rirekirs!code
            .Col = 2
            .Text = rirekirs!koumoku
            .Col = 3
            .Text = rirekirs!koumoku2
            .Col = 4
            .Text = rirekirs!kingaku
        End With
       rirekirs.MoveNext
    Next i
    rirekirs.Close
End If


End Sub
Sub seikyuu_voip_set(kikan As String, shitei_code As String)
Dim rirekirs As Recordset, rirekisuu As Long



sql = "SELECT *" & _
    " FROM seikyuu" & _
    " WHERE nentsuki='" & kikan & "' and code ='" & shitei_code & "'"
    
If FcSQlGet_voip(rirekirs, sql, PrMsg) = False Then
    With frmvoip.gridseikyuu
        .Font.Size = 11
        .Rows = 1
        .ColWidth(0) = 1200
        .ColWidth(1) = 1300
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1500
        .Row = 0
        .Col = 0
        .Text = "ÇhÇc"
        .Col = 1
        .Text = "ÉRÅ[Éh"
        .Col = 2
        .Text = "äJénì˙"
        .Col = 3
        .Text = "èIóπì˙"
        .Col = 4
        .Text = "êøãÅäz"
        .Col = 5
        .Text = "ê≈äz"
        .Col = 6
        .Text = "éxï•ä˙å¿"
    End With
    Exit Sub
Else
    rirekirs.MoveLast
    rirekisuu = rirekirs.RecordCount
    With frmvoip.gridseikyuu
        .Font.Size = 11
        .Rows = rirekisuu + 1
        .ColWidth(0) = 1100
        .ColWidth(1) = 1300
        .ColWidth(2) = 1300
        .ColWidth(3) = 1300
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1500
        .Row = 0
        .Col = 0
        .Text = "ÇhÇc"
        .Col = 1
        .Text = "ÉRÅ[Éh"
        .Col = 2
        .Text = "äJénì˙"
        .Col = 3
        .Text = "èIóπì˙"
        .Col = 4
        .Text = "êøãÅäz"
        .Col = 5
        .Text = "ê≈äz"
        .Col = 6
        .Text = "éxï•ä˙å¿"
    End With
    rirekirs.MoveFirst
    For i = 1 To rirekisuu
        With frmvoip.gridseikyuu
            .Row = i
            .Col = 0
            .Text = rirekirs!seikyuid
            .Col = 1
            .Text = rirekirs!code
            .Col = 2
            .Text = Format(rirekirs!hajime, "@@@@/@@/@@")
            .Col = 3
            .Text = Format(rirekirs!owari, "@@@@/@@/@@")
            .Col = 4
            .Text = rirekirs!seikyuugaku
            .Col = 5
            .Text = rirekirs!zei
            .Col = 6
            .Text = Format(rirekirs!hi, "@@@@/@@/@@")
        End With
       rirekirs.MoveNext
    Next i
    rirekirs.Close
End If


End Sub

Private Sub Form_Load()
Dim in2 As Integer, shitenen As Long

For in2 = 1 To 12
    If in2 < 5 Then
        shitenen = CLng(Format(Date, "yyyy")) - 4 + in2
        frmvoip.cmbnen.AddItem CStr(shitenen)
    End If
    frmvoip.cmbtsuki.AddItem Format(in2, "00")
Next in2

End Sub

Private Sub Picture1_Click()
    Dim n_nen As String, n_tsuki As String, n_hi As String
    Dim rirekirs5 As Recordset, sql5 As String
    
        If cmbnen.ListIndex = -1 Then
            MsgBox "ï\é¶åüçıÇµÇΩÇ¢îNÇëIëÇµÇƒÇ≠ÇæÇ≥Ç¢"
            Exit Sub
        Else
            n_nen = cmbnen.Text
        End If
        If cmbtsuki.ListIndex = -1 Then
            MsgBox "ï\é¶åüçıÇµÇΩÇ¢åéÇëIëÇµÇƒÇ≠ÇæÇ≥Ç¢"
            Exit Sub
        Else
            n_tsuki = cmbtsuki.Text
        End If
        n_hi = n_nen & n_tsuki

If kidou_voip = False Then
    MsgBox "ÉfÅ[É^ÉxÅ[ÉXÉtÉ@ÉCÉãÇäJÇØÇ‹ÇπÇÒÇ≈ÇµÇΩÅB"
    Exit Sub
End If
sql5 = "SELECT code" & _
    " FROM voip_m" & _
    " WHERE kojinid='" & frmtouroku2.lblid.Caption & "' group by code"
    
If FcSQlGet_voip(rirekirs5, sql5, PrMsg) = False Then
    MsgBox "VoIPÇÃêøãÅÉfÅ[É^ÇÕÇ†ÇËÇ‹ÇπÇÒÅB"
    Exit Sub
Else
    'rirekirs5.MoveFirst
    'Do Until rirekirs5.EOF
        seikyuu_voip_set n_hi, rirekirs5!code
        seikyuu_voip_set2 n_hi, rirekirs5!code
    
        'rirekirs5.MoveNext
    'Loop
    rirekirs5.Close
End If

End Sub

Private Sub Picture13_Click()
Unload Me

End Sub
