VERSION 5.00
Begin VB.Form frmcl 
   Caption         =   "ÉPÅ[ÉuÉãÉâÉCÉì"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   4260
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.PictureBox Picture16 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2580
      MouseIcon       =   "frmcl.frx":0000
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmcl.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   1740
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3420
      MouseIcon       =   "frmcl.frx":11C9
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmcl.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   1740
      Width           =   480
   End
   Begin VB.Frame Frame3 
      Caption         =   "ÉPÅ[ÉuÉãÉâÉCÉìì‡óe"
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
      Height          =   1365
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   4065
      Begin VB.TextBox txttelno 
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
         TabIndex        =   1
         Top             =   795
         Width           =   2300
      End
      Begin VB.Label lblclid 
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
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label14 
         Caption         =   "ìdòbî‘çÜÅF"
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
         TabIndex        =   4
         Top             =   855
         Width           =   1635
      End
      Begin VB.Label Label13 
         Caption         =   "CL IDÅF"
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
         TabIndex        =   3
         Top             =   420
         Width           =   1635
      End
      Begin VB.Label lblkojinid 
         Height          =   225
         Left            =   2820
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
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
      Left            =   2580
      TabIndex        =   8
      Top             =   2220
      Width           =   90
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
      Left            =   3480
      TabIndex        =   7
      Top             =   2220
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   855
      Left            =   2340
      Top             =   1680
      Width           =   1740
   End
End
Attribute VB_Name = "frmcl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Picture13_Click()
Unload Me

End Sub

Private Sub Picture16_Click()
Dim newidid As String, newtelno As String, newkojinid As String
Dim nnid As String, newbbasumu As String, newccasumu As String
    
With frmcl
    newidid = Trim(.lblclid.Caption)
    newkojinid = Trim(.lblkojinid.Caption)
    newtelno = Trim(.txttelno.Text)
    If newtelno = "" Then
        MsgBox "ìdòbî‘çÜÇ™ì¸óÕÇ≥ÇÍÇƒÇ¢Ç‹ÇπÇÒÅB"
        Exit Sub
    End If
    
    If Len(newtelno) <> 12 Then
        MsgBox "ìdòbî‘çÜÇÃì¸óÕÇ™ïsê≥Ç≈Ç∑ÅB"
        Exit Sub
    End If
    
    
    
        
End With
On Error GoTo errstb
If Label1(15).Caption = "ìoò^" Then
     sql = "INSERT INTO cl VALUES('" & newidid & "'," & _
                    "'" & newkojinid & "','" & newtelno & "')"
     db.Execute sql, dbSQLPassThrough
     nnid = CStr(CLng(newidid) + 1)
        If soukatsu(1, "renban", 17, nnid, 1, 1) = False Then
                MsgBox "ê›íËÇÃèëÇ´çûÇ›Ç…é∏îsÇµÇ‹ÇµÇΩÅB"
                Exit Sub
        End If
     MsgBox "ìoò^ÇµÇ‹ÇµÇΩÅB"
 Else
 
     sql = "UPDATE cl SET tel = '" & newtelno & "'" & _
                  " WHERE clid = '" & newidid & "'"
     db.Execute sql, dbSQLPassThrough
    MsgBox "ïœçXÇµÇ‹ÇµÇΩÅB"
   

End If
On Error GoTo 0
Unload Me
cl_grid_set (frmtouroku2.lblid.Caption)

Exit Sub

errstb:
    MsgBox "CLÇÃçXêVÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
    Unload Me

End Sub
