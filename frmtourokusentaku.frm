VERSION 5.00
Begin VB.Form frmtourokusentaku 
   Caption         =   "ìoò^ëIë"
   ClientHeight    =   1080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1080
   ScaleWidth      =   4095
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   570
      MouseIcon       =   "frmtourokusentaku.frx":0000
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmtourokusentaku.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      ToolTipText     =   "ìoò^èàóùÇçsÇ¢Ç‹Ç∑ÅB"
      Top             =   180
      Width           =   480
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3090
      MouseIcon       =   "frmtourokusentaku.frx":11C9
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmtourokusentaku.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   180
      Width           =   480
   End
   Begin VB.PictureBox Picture17 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1830
      MouseIcon       =   "frmtourokusentaku.frx":2392
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmtourokusentaku.frx":2C91
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblno 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'é¿ê¸
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label Label1 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FFFFFF&
      Caption         =   "í èÌìoò^"
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
      Height          =   285
      Index           =   6
      Left            =   330
      TabIndex        =   4
      Top             =   660
      Width           =   1005
   End
   Begin VB.Label seisei 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FFFFFF&
      Caption         =   "ó\íËìoò^"
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
      Height          =   285
      Left            =   1575
      TabIndex        =   3
      Top             =   660
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ñﬂÇÈ"
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
      Left            =   3150
      TabIndex        =   2
      Top             =   660
      Width           =   450
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   855
      Left            =   90
      Top             =   120
      Width           =   3870
   End
End
Attribute VB_Name = "frmtourokusentaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture13_Click()
  
  Unload frmtourokusentaku
  
    Set frmtourokusentaku = Nothing
    
End Sub

Private Sub Picture17_Click()
    
    Dim s_no As String
    
    s_no = Trim(Me.lblno.Caption)

    Unload frmtourokusentaku
  
    Set frmtourokusentaku = Nothing
    


    With frmtourokusentaku2
    
        If s_no = "0" Then
            .lblshuruiid.Caption = "0"
            .lblshuruimei.Caption = "éÂå_ñÒ"
            
            'éÂå_ñÒÇÃÉZÉbÉg
           sbshuruiset_2
            
             .cmbkeiyakushurui.Visible = True
            .cmbsonotakeiyaku.Visible = False
        
        
            .txtkeiyakusuu.Enabled = False
        Else
            .lblshuruiid.Caption = "1"
            .lblshuruimei.Caption = "ÇªÇÃëºå_ñÒ"
        
            'ÇªÇÃëºÇÃÉZÉbÉg
            sbsonotaset_2
            
            .cmbkeiyakushurui.Visible = False
            .cmbsonotakeiyaku.Visible = True
            
            .txtkeiyakusuu.Enabled = True
        End If
        
        
      
        
        
        
        .Show 1
    End With
    
End Sub

Private Sub Picture6_Click()

 
    

If Trim(Me.lblno.Caption) = "0" Then
   Unload frmtourokusentaku
  
    Set frmtourokusentaku = Nothing
    frmkeiyaku4.Show 1
Else
   Unload frmtourokusentaku
  
    Set frmtourokusentaku = Nothing
    frmkeiyaku5.Show 1
End If


End Sub
