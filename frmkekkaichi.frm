VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmkekkaichi 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "åüçıåãâ âÊñ "
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   12990
   StartUpPosition =   1  'µ∞≈∞ Ã´∞—ÇÃíÜâõ
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8640
      MouseIcon       =   "frmkekkaichi.frx":0000
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmkekkaichi.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   8340
      Width           =   480
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   10380
      MouseIcon       =   "frmkekkaichi.frx":1741
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmkekkaichi.frx":2040
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   8340
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   12000
      MouseIcon       =   "frmkekkaichi.frx":290A
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmkekkaichi.frx":3209
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   8340
      Width           =   480
   End
   Begin VB.Frame Frame2 
      Caption         =   "åüçıåãâ "
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
      Height          =   1035
      Left            =   270
      TabIndex        =   2
      Top             =   8160
      Width           =   3135
      Begin VB.Label lblkensuu 
         Alignment       =   2  'íÜâõëµÇ¶
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
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "åèÇ≈Ç∑"
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
         Left            =   2280
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "èWåvåãâ "
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
      Height          =   7815
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   12735
      Begin MSFlexGridLib.MSFlexGrid kekkagrd 
         Height          =   7365
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   12435
         _ExtentX        =   21934
         _ExtentY        =   12991
         _Version        =   393216
         Cols            =   7
         SelectionMode   =   1
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "çƒåüçı"
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
      Left            =   8550
      TabIndex        =   10
      Top             =   8850
      Width           =   720
   End
   Begin VB.Label SSCJIKKOU 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ï\é¶"
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
      Left            =   10350
      TabIndex        =   8
      Top             =   8820
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "∑¨›æŸ"
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
      Left            =   11910
      TabIndex        =   7
      Top             =   8820
      Width           =   750
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   915
      Left            =   3660
      Top             =   8220
      Width           =   9210
   End
End
Attribute VB_Name = "frmkekkaichi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdend_Click()
End Sub


Private Sub cmdhyouji_Click()

End Sub

Private Sub cmdsaiken_Click()
End Sub

Private Sub Form_Load()
    kekkagrd.Font.Size = 11

End Sub

Private Sub kekkagrd_KeyPress(KeyAscii As Integer)
    
    Dim h_id As String

If KeyAscii = 13 Then
    If frmkekkaichi.kekkagrd.Rows <> 1 Then
        If frmkekkaichi.kekkagrd.Row <> 0 Then
            frmkekkaichi.kekkagrd.Col = 0
            h_id = kekkagrd.Text
            
            If hyouji_touroku(h_id) = False Then
                Exit Sub
            End If
            If nhk_set(h_id) = False Then
                Screen.MousePointer = 0
                Exit Sub
            End If
            
            keiyakuset (h_id)
            keiyakuset2 (h_id)
            chk_kyuushi (h_id)
            chk_log (h_id)
            'Unload frmkekkaichi
            'Set frmkekkaichi = Nothing
            frmkekkaichi.Hide
            Unload frmkensaku
            Set frmkensaku = Nothing
            frmtouroku2.Picture17.Enabled = True
            frmtouroku2.Show 1
        End If
    End If
End If

End Sub

Private Sub Picture1_Click()
Unload frmkekkaichi
Set frmkekkaichi = Nothing
frmkensaku.Show 1

End Sub

Private Sub Picture13_Click()
Unload frmkekkaichi
Set frmkekkaichi = Nothing
Unload frmkensaku
Set frmkensaku = Nothing

End Sub

Private Sub Picture7_Click()

Dim h_id As String

frmkekkaichi.kekkagrd.Col = 0
h_id = kekkagrd.Text

If hyouji_touroku(h_id) = False Then
    Exit Sub
End If
If nhk_set(h_id) = False Then
    Screen.MousePointer = 0
    Exit Sub
End If

keiyakuset (h_id)
keiyakuset2 (h_id)
chk_kyuushi (h_id)
chk_log (h_id)
frmkekkaichi.Hide
Unload frmkensaku
Set frmkensaku = Nothing
frmtouroku2.Picture17.Enabled = True
frmtouroku2.Show 1

End Sub
