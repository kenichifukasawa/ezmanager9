VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsouken2 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "êøãÅÅEì¸ã‡èWåvï\"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Visible         =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "ì¸ã‡ï˚ñ@"
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
      Height          =   945
      Left            =   10410
      TabIndex        =   13
      Top             =   120
      Width           =   2205
      Begin VB.ComboBox cmbbankid 
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
         TabIndex        =   14
         Top             =   420
         Width           =   1965
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   12960
      MouseIcon       =   "frmsouken2.frx":0000
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmsouken2.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      ToolTipText     =   "ì¸ã‡óöóÇï\é¶ÇµÇ‹Ç∑ÅB"
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
      Left            =   14490
      MouseIcon       =   "frmsouken2.frx":1741
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmsouken2.frx":2040
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   180
      Width           =   480
   End
   Begin VB.Frame Frame3 
      Caption         =   "îNåéëIë"
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
      Height          =   945
      Left            =   7800
      TabIndex        =   2
      Top             =   120
      Width           =   2535
      Begin VB.ComboBox combzennen 
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
         Left            =   240
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   10
         Top             =   420
         Width           =   915
      End
      Begin VB.ComboBox combzentsuki2 
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
         Left            =   1440
         Style           =   2  'ƒﬁ€ØÃﬂ¿ﬁ≥› ÿΩƒ
         TabIndex        =   9
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000004&
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
         Index           =   0
         Left            =   1200
         TabIndex        =   12
         Top             =   540
         Width           =   255
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000004&
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
         Index           =   0
         Left            =   2160
         TabIndex        =   11
         Top             =   540
         Width           =   255
      End
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   9825
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   17330
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.PictureBox Picture15 
      Appearance      =   0  'Ã◊Øƒ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Ç»Çµ
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   13740
      MouseIcon       =   "frmsouken2.frx":290A
      MousePointer    =   99  '’∞ªﬁ∞íËã`
      Picture         =   "frmsouken2.frx":3209
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "èoóÕ"
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
      Index           =   1
      Left            =   13770
      TabIndex        =   15
      Top             =   720
      Width           =   480
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
      Left            =   12720
      TabIndex        =   8
      Top             =   720
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
      Left            =   14520
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblkekka 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   7635
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'ïsìßñæ
      BorderStyle     =   0  'ìßñæ
      Height          =   915
      Left            =   12690
      Top             =   120
      Width           =   2460
   End
   Begin VB.Label Label1 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FFFFFF&
      Caption         =   "FDD"
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
      Index           =   14
      Left            =   13350
      TabIndex        =   6
      Top             =   720
      Width           =   795
   End
End
Attribute VB_Name = "frmsouken2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private chdata()
Private chusercount As Long
Private sentakusumidesuka As Integer

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Dim in2 As Integer, shitenen As Long

Dim keikanenn As Integer
keikanenn = CLng(Format(Date, "yyyy")) - 2006

For i = keikanenn To 0 Step -1
    combzennen.AddItem Format(Date, "yyyy") - i

Next
 combzennen.ListIndex = keikanenn
  
For in2 = 1 To 12
   
    frmsouken2.combzentsuki2.AddItem Format(in2, "00")
Next in2

sbnbankset6
frmsouken2.cmbbankid.ListIndex = -1
sentakusumidesuka = 0
End Sub




Private Sub lblkekka_DblClick()
     Dim shuwhere As String, karikindesu As Long, karikindesu2 As String, karihiduke As String
        Dim sql_ch As String, sql_ch2 As String, nowcount As Long
        Dim rs_ch As Recordset, rs_ch2 As Recordset, rs_ch3 As Recordset, sql_ch3 As String
        Dim kin_inter As Long, kin_catv As Long, nhkcount As Integer
        Dim kin_inter2 As Long, kin_catv2 As Long, shiteibankid As String
        Dim newnewshiteibi As String, newnewshiteibi1 As String, newnewshiteibi2 As String
        Dim sounyuukingaku As Long, souseikyuukingaku As Long
        Dim sql_ch8 As String, rs_ch8 As Recordset
        Dim sql_ch9 As String, rs_ch9 As Recordset
        Dim newnewshiteibi5 As String, newnewshiteibi52 As String
        Dim sql_sonkin As String, rs_sonkin As Recordset
        
        chusercount = 0
        souseikyuukingaku = 0
        sounyuukingaku = 0
        frmsouken2.lv1.ListItems.Clear
        frmsouken2.lv1.ColumnHeaders.Clear
        DoEvents
        
        If frmsouken2.combzennen.ListIndex = -1 Then
            MsgBox "éwíËîNêîÇéwíËÇµÇƒÇ©ÇÁé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
        If frmsouken2.combzentsuki2.ListIndex = -1 Then
            MsgBox "éwíËåéêîÇéwíËÇµÇƒÇ©ÇÁé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
        If frmsouken2.cmbbankid.ListIndex = -1 Then
            shiteibankid = "-1"
        Else
            shiteibankid = Trim(Mid(frmsouken2.cmbbankid.Text, 1, 4))
        End If
        
        Me.Hide
       FRMOMACHI.lblmsg.Caption = "èWåvÇÃèÄîıíÜ..."
       FRMOMACHI.pb.Visible = True
        FRMOMACHI.pb.Value = 0
        FRMOMACHI.Show
        DoEvents
        Screen.MousePointer = 11
        
        newnewshiteibi = Format(frmsouken2.combzennen.Text, "0000") & Format(frmsouken2.combzentsuki2.Text, "00")
        newnewshiteibi1 = Format(frmsouken2.combzennen.Text, "0000") & Format(frmsouken2.combzentsuki2.Text, "00") & "01"
        newnewshiteibi2 = Format(frmsouken2.combzennen.Text, "0000") & Format(frmsouken2.combzentsuki2.Text, "00") & "31"
        If frmsouken2.combzentsuki2.Text = "01" Then
            newnewshiteibi5 = Format(CLng(frmsouken2.combzennen.Text) - 1, "0000") & "12"
        Else
            newnewshiteibi5 = Format(frmsouken2.combzennen.Text, "0000") & Format(CInt(frmsouken2.combzentsuki2.Text) - 1, "00")
        End If
        With frmsouken2
        If frmsouken2.cmbbankid.ListIndex = -1 Then
                .lv1.Checkboxes = False
                   .lv1.ColumnHeaders. _
                   Add , , "NO", 500
                   
                   .lv1.ColumnHeaders. _
                   Add , , "id", 800
                   
                   .lv1.ColumnHeaders. _
                   Add , , "å_ñÒñº", 5000
                   
                   .lv1.ColumnHeaders. _
                   Add , , "éÌï ", 1200
                   
                   .lv1.ColumnHeaders. _
                   Add , , "éÌóﬁ", 1200
                          
                   .lv1.ColumnHeaders. _
                   Add , , "åJâzäz", 1400, 1
                   
                   .lv1.ColumnHeaders. _
                   Add , , "êøãÅäz", 1400, 1
                          
                   .lv1.ColumnHeaders. _
                   Add , , "ì¸ã‡äz", 1400, 1
                   
                   .lv1.ColumnHeaders. _
                   Add , , "ñ¢î[äz", 1400, 1
        Else
            If shiteibankid = "9999" Then
                .lv1.Checkboxes = False
                   .lv1.ColumnHeaders. _
                   Add , , "NO", 500
                   
                   .lv1.ColumnHeaders. _
                   Add , , "id", 800
                   
                   .lv1.ColumnHeaders. _
                   Add , , "å_ñÒñº", 5000
                   
                   .lv1.ColumnHeaders. _
                   Add , , "éÌï ", 1200
                   
                   .lv1.ColumnHeaders. _
                   Add , , "éÌóﬁ", 1200
                          
                   .lv1.ColumnHeaders. _
                   Add , , "çáåväz", 1400, 1
            
                   .lv1.ColumnHeaders. _
                   Add , , "éÊè¡ã‡äz", 1400, 1
            
                   .lv1.ColumnHeaders. _
                   Add , , "ëπé∏ã‡äz", 1400, 1
            Else
                .lv1.Checkboxes = False
                   .lv1.ColumnHeaders. _
                   Add , , "NO", 500
                   
                   .lv1.ColumnHeaders. _
                   Add , , "id", 800
                   
                   .lv1.ColumnHeaders. _
                   Add , , "å_ñÒñº", 5000
                   
                   .lv1.ColumnHeaders. _
                   Add , , "éÌï ", 1200
                   
                   .lv1.ColumnHeaders. _
                   Add , , "éÌóﬁ", 1200
                          
                   .lv1.ColumnHeaders. _
                   Add , , "ì¸ã‡äz", 1400, 1
             End If
        End If
        End With
        
    ' ListItem ÉIÉuÉWÉFÉNÉgÇí«â¡Ç∑ÇÈÇΩÇﬂÇÃïœêîÇêÈåæÇµÇ‹Ç∑ÅB
    Dim itmX As ListItem, kakikubun As String, hhhajime As Long
         kin_inter = 0
         kin_catv = 0
         kin_inter2 = 0
         kin_catv2 = 0
         nowcount = 1
With frmsouken2
        FRMOMACHI.lblmsg.Caption = "éÌóﬁÇÃèWåvíÜ..."
        DoEvents
       sql_ch = "SELECT * FROM shurui order by shuruiid"
        If FcSQlGet(rs_ch, sql_ch, PrMsg) = True Then
            If kidou2 = False Then
                 MsgBox "ÉfÅ[É^ÉxÅ[ÉXÉtÉ@ÉCÉãÇäJÇØÇ‹ÇπÇÒÇ≈ÇµÇΩÅB"
                    Unload FRMOMACHI
            Set FRMOMACHI = Nothing
            Screen.MousePointer = 0
                 Exit Sub
             End If
             
             
             FRMOMACHI.pb.Min = 0
                FRMOMACHI.pb.Max = rs_ch.RecordCount + 1
             FRMOMACHI.pb.Value = 0
             DoEvents
             
               rs_ch.MoveFirst
               While Not rs_ch.EOF
                    ReDim Preserve chdata(17, nowcount)
                      chdata(0, nowcount) = rs_ch!shuruiid
                      chdata(1, nowcount) = rs_ch!shuruimei
                       If rs_ch!kankei = "0" Then
                            chdata(2, nowcount) = "CATV"
                      Else
                            chdata(2, nowcount) = "Inter"
                      End If
                        nowcount = nowcount + 1
                        
                    rs_ch.MoveNext   ' éüÇÃÉåÉRÅ[ÉhÇ…à⁄ìÆÇµÇ‹Ç∑ÅB
                Wend
                rs_ch.Close
                
                Dim isi As Integer
                
                For isi = 1 To nowcount - 1
                
                      FRMOMACHI.Caption = "èàóùíÜÅEÅE(" & chdata(0, isi) & ")1"
                       DoEvents
                      
                     
                      chdata(3, isi) = "éÂå_ñÒ"
                      'åJâzäz
                            If shiteibankid = "-1" Then
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='1' AND shousai.keiid ='" & chdata(0, isi) & "'"
                            ElseIf shiteibankid = "" Then
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='1' AND shousai.keiid ='" & chdata(0, isi) & "'" & _
                                    " and shousai.bankid='" & Space(1) & "'"
                            Else
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='1' AND shousai.keiid ='" & chdata(0, isi) & "'" & _
                                    " and shousai.bankid='" & shiteibankid & "'"
                            End If
                                
                            If FcSQlGet2(rs_ch8, sql_ch8, PrMsg) = False Then
                                 chdata(4, isi) = 0
                            Else
                                If IsNull(rs_ch8!kurigou) Then
                                    chdata(4, isi) = 0
                                Else
                                    chdata(4, isi) = rs_ch8!kurigou
                                End If
                                rs_ch8.Close
                            End If

                        FRMOMACHI.Caption = "èàóùíÜÅEÅE(" & chdata(0, isi) & ")2"
                       DoEvents
                      'êøãÅäz
                            If shiteibankid = "-1" Then
                                sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                        " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                                        " and keiid='" & chdata(0, isi) & "' and keishurui='1'"
                            ElseIf shiteibankid = "" Then
                                sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                        " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                                        " and keiid='" & chdata(0, isi) & "' and keishurui='1' and bankid='" & Space(1) & "'"
                            Else
                                sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                        " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                                        " and keiid='" & chdata(0, isi) & "' and keishurui='1' and bankid='" & shiteibankid & "'"
                            End If
                            If FcSQlGet2(rs_ch2, sql_ch2, PrMsg) = False Then
                                 chdata(5, isi) = 0
                            Else
                                If IsNull(rs_ch2!gkgk) Then
                                    chdata(5, isi) = 0
                                Else
                                    chdata(5, isi) = rs_ch2!gkgk
                                    
                                End If
                                rs_ch2.Close
                            End If
                              FRMOMACHI.Caption = "èàóùíÜÅEÅE(" & chdata(0, isi) & ")3"
                       DoEvents
                        'ì¸ã‡
                        If shiteibankid = "-1" Then
                            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                    " and keiid='" & chdata(0, isi) & "' and keishurui='1'"
                        ElseIf shiteibankid = "" Then
                            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                    " and keiid='" & chdata(0, isi) & "' and keishurui='1' and bankid='" & Space(1) & "'"
                        ElseIf shiteibankid = "9999" Then
                            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                    " and keiid='" & chdata(0, isi) & "' and keishurui='1' and keist='2'"
                        Else
                            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                    " and keiid='" & chdata(0, isi) & "' and keishurui='1' and bankid='" & shiteibankid & "'"
                        End If
                            If FcSQlGet2(rs_ch3, sql_ch3, PrMsg) = False Then
                                chdata(6, isi) = 0
                            Else
                                If IsNull(rs_ch3!gkgk) Then
                                    chdata(6, isi) = 0
                                Else
                                    chdata(6, isi) = rs_ch3!gkgk
                                    sounyuukingaku = sounyuukingaku + rs_ch3!gkgk
                                End If
                                rs_ch3.Close
                                  FRMOMACHI.Caption = "èàóùíÜÅEÅE(" & chdata(0, isi) & ")3-1"
                       DoEvents
                                If shiteibankid = "9999" Then
                                    sql_sonkin = "select sum(kingaku) as torikeshikin from koushinkiroku" & _
                                                 " where koushinbi between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                                  " and houhou='3' and keiid='" & chdata(0, isi) & "'"
                                    If FcSQlGet2(rs_sonkin, sql_sonkin, PrMsg) = False Then
                                        chdata(8, isi) = 0
                                    Else
                                        If IsNull(rs_sonkin!torikeshikin) Then
                                        chdata(8, isi) = 0
                                        Else
                                        chdata(8, isi) = rs_sonkin!torikeshikin
                                        End If
                                        rs_sonkin.Close
                                    End If
                                      FRMOMACHI.Caption = "èàóùíÜÅEÅE(" & chdata(0, isi) & ")3-2"
                       DoEvents
                                    sql_sonkin = "select sum(kingaku) as sonkin from koushinkiroku" & _
                                                 " where koushinbi between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                                  " and houhou='4' and keiid='" & chdata(0, isi) & "'"
                                    If FcSQlGet2(rs_sonkin, sql_sonkin, PrMsg) = False Then
                                        chdata(9, isi) = 0
                                    Else
                                        If IsNull(rs_sonkin!sonkin) Then
                                        chdata(9, isi) = 0
                                        Else
                                        chdata(9, isi) = rs_sonkin!sonkin
                                        End If
                                        rs_sonkin.Close
                                    End If
                                End If
                            End If
                            If chdata(2, isi) = "Inter" Then
                                kin_inter = kin_inter + CLng(chdata(6, isi))
                            Else
                                kin_catv = kin_catv + CLng(chdata(6, isi))
                            End If
                            
                      'åJâzäz
                        FRMOMACHI.Caption = "èàóùíÜÅEÅE(" & chdata(0, isi) & ")4"
                       DoEvents
                            chdata(7, isi) = chdata(4, isi) + chdata(5, isi) - chdata(6, isi)
                            
                          
                            
                              FRMOMACHI.pb.Value = isi
                                DoEvents
                                
                    Next
             
        End If
        FRMOMACHI.lblmsg.Caption = "å_ñÒÇÃèWåvíÜ..."
        DoEvents
        
  
        Dim nowcount2 As Integer
        nowcount2 = 0
        
       sql_ch = "SELECT * FROM sonota order by sonotaid"
        If FcSQlGet(rs_ch, sql_ch, PrMsg) = True Then
            If kidou2 = False Then
                 MsgBox "ÉfÅ[É^ÉxÅ[ÉXÉtÉ@ÉCÉãÇäJÇØÇ‹ÇπÇÒÇ≈ÇµÇΩÅB"
                    Unload FRMOMACHI
            Set FRMOMACHI = Nothing
            Screen.MousePointer = 0
             
                 Exit Sub
             End If
             
              FRMOMACHI.pb.Min = 0
                FRMOMACHI.pb.Max = rs_ch.RecordCount + 1 + nowcount
             FRMOMACHI.pb.Value = 0
             DoEvents
             
             Dim s_kaishicount As Integer
             s_kaishicount = nowcount
             
               rs_ch.MoveFirst
               While Not rs_ch.EOF
                    ReDim Preserve chdata(17, nowcount)
                      chdata(0, nowcount) = rs_ch!sonotaid
                      chdata(1, nowcount) = rs_ch!sonota
                      If rs_ch!kankei = "0" Then
                            chdata(2, nowcount) = "CATV"
                      Else
                            chdata(2, nowcount) = "Inter"
                      End If
                      
                         nowcount = nowcount + 1
                        
                    rs_ch.MoveNext   ' éüÇÃÉåÉRÅ[ÉhÇ…à⁄ìÆÇµÇ‹Ç∑ÅB
                Wend
                rs_ch.Close
                
               
                For isi = s_kaishicount To nowcount - 1
                      
                      chdata(3, isi) = "ÇªÇÃëº"
                      'åJâzäz
                            If shiteibankid = "-1" Then
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='2' AND shousai.keiid ='" & chdata(0, isi) & "'"
                            ElseIf shiteibankid = "" Then
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='2' AND shousai.keiid ='" & chdata(0, isi) & "'" & _
                                    " and shousai.bankid='" & Space(1) & "'"
                            Else
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='2' AND shousai.keiid ='" & chdata(0, isi) & "'" & _
                                    " and shousai.bankid='" & shiteibankid & "'"
                            End If
                                
                            If FcSQlGet2(rs_ch8, sql_ch8, PrMsg) = False Then
                                 chdata(4, isi) = 0
                            Else
                                If IsNull(rs_ch8!kurigou) Then
                                    chdata(4, isi) = 0
                                Else
                                    chdata(4, isi) = rs_ch8!kurigou
                                End If
                                rs_ch8.Close
                            End If
                      
                      'êøãÅäz
                        If shiteibankid = "-1" Then
                            sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                                    " and keiid='" & chdata(0, isi) & "' and keishurui='2'"
                        ElseIf shiteibankid = "" Then
                            sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                                    " and keiid='" & chdata(0, isi) & "' and keishurui='2' and bankid='" & Space(1) & "'"
                        Else
                            sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                                    " and keiid='" & chdata(0, isi) & "' and keishurui='2' and bankid='" & shiteibankid & "'"
                        End If
                            If FcSQlGet2(rs_ch2, sql_ch2, PrMsg) = False Then
                                 chdata(5, isi) = 0
                            Else
                                If IsNull(rs_ch2!gkgk) Then
                                    chdata(5, isi) = 0
                                Else
                                    chdata(5, isi) = rs_ch2!gkgk
                                    
                                End If
                                rs_ch2.Close
                            End If
                        'ì¸ã‡
                        If shiteibankid = "-1" Then
                            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                    " and keiid='" & chdata(0, isi) & "' and keishurui='2'"
                        ElseIf shiteibankid = "" Then
                            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                    " and keiid='" & chdata(0, isi) & "' and keishurui='2' and bankid='" & Space(1) & "'"
                        ElseIf shiteibankid = "9999" Then
                            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                    " and keiid='" & chdata(0, isi) & "' and keishurui='2' and keist='2'"
                        Else
                            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                    " and keiid='" & chdata(0, isi) & "' and keishurui='2' and bankid='" & shiteibankid & "'"
                        End If
                            If FcSQlGet2(rs_ch3, sql_ch3, PrMsg) = False Then
                                chdata(6, isi) = 0
                            Else
                                If IsNull(rs_ch3!gkgk) Then
                                    chdata(6, isi) = 0
                                Else
                                    chdata(6, isi) = rs_ch3!gkgk
                                    sounyuukingaku = sounyuukingaku + rs_ch3!gkgk
                                End If
                                rs_ch3.Close
                                If shiteibankid = "9999" Then
                                    sql_sonkin = "select sum(kingaku) as torikeshikin from koushinkiroku" & _
                                                 " where koushinbi between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                                  " and houhou='5' and keiid='" & chdata(0, isi) & "'"
                                    If FcSQlGet2(rs_sonkin, sql_sonkin, PrMsg) = False Then
                                        chdata(8, isi) = 0
                                    Else
                                         If IsNull(rs_sonkin!torikeshikin) Then
                                        chdata(8, isi) = 0
                                        Else
                                       chdata(8, isi) = rs_sonkin!torikeshikin
                                        End If
                                        rs_sonkin.Close
                                    End If
                                    sql_sonkin = "select sum(kingaku) as sonkin from koushinkiroku" & _
                                                 " where koushinbi between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                                  " and houhou='6' and keiid='" & chdata(0, isi) & "'"
                                    If FcSQlGet2(rs_sonkin, sql_sonkin, PrMsg) = False Then
                                        chdata(9, isi) = 0
                                    Else
                                        If IsNull(rs_sonkin!sonkin) Then
                                        chdata(9, isi) = 0
                                        Else
                                        chdata(9, isi) = rs_sonkin!sonkin
                                        End If
                                        rs_sonkin.Close
                                    End If
                                End If
                            End If
                            If chdata(2, isi) = "Inter" Then
                                kin_inter2 = kin_inter2 + CLng(chdata(6, isi))
                            Else
                                kin_catv2 = kin_catv2 + CLng(chdata(6, isi))
                            End If
                            
                      'åJâzäz
                            chdata(7, isi) = chdata(4, isi) + chdata(5, isi) - chdata(6, isi)
                            
                            
                            
                            'nowcount = nowcount + 1
                            nowcount2 = nowcount2 + 1
                            
                               FRMOMACHI.pb.Value = nowcount2 - 1
                                DoEvents
              Next
              
        End If
        
    FRMOMACHI.lblmsg.Caption = "HNKÇÃèWåvíÜ..."
        DoEvents
     For nhkcount = 0 To 6
        ReDim Preserve chdata(17, nowcount)
        chdata(0, nowcount) = CStr(nhkcount)
        Select Case nhkcount
        Case 0
            chdata(1, nowcount) = "ÇmÇgÇjÇPÉñåé"
        Case 1
            chdata(1, nowcount) = "ÇmÇgÇjÇQÉñåé"
        Case 2
            chdata(1, nowcount) = "ÇmÇgÇjÇRÉñåé"
        Case 3
            chdata(1, nowcount) = "ÇmÇgÇjÇSÉñåé"
        Case 4
            chdata(1, nowcount) = "ÇmÇgÇjÇTÉñåé"
        Case 5
            chdata(1, nowcount) = "ÇmÇgÇjÇUÉñåé"
        Case 6
            chdata(1, nowcount) = "ÇmÇgÇjÇUÉñåéÉnÉìÉf"
        Case 7
            chdata(1, nowcount) = "ÇmÇgÇjÇUÉñåéâ∆ë∞äÑ"
         End Select
        chdata(2, nowcount) = "ÇmÇgÇj"
        chdata(3, nowcount) = "ÇmÇgÇj"
        
                      'åJâzäz
                            If shiteibankid = "-1" Then
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='3' AND shousai.keiid ='" & CStr(nhkcount) & "'"
                            ElseIf shiteibankid = "" Then
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='3' AND shousai.keiid ='" & CStr(nhkcount) & "'" & _
                                    " and shousai.bankid='" & Space(1) & "'"
                            Else
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='3' AND shousai.keiid ='" & CStr(nhkcount) & "'" & _
                                    " and shousai.bankid='" & shiteibankid & "'"
                            End If
                                
                            If FcSQlGet2(rs_ch8, sql_ch8, PrMsg) = False Then
                                 chdata(4, nowcount) = 0
                            Else
                                If IsNull(rs_ch8!kurigou) Then
                                    chdata(4, nowcount) = 0
                                Else
                                    chdata(4, nowcount) = rs_ch8!kurigou
                                End If
                                rs_ch8.Close
                            End If
        
        
        'êøãÅäz
          If shiteibankid = "-1" Then
              sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                      " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                      " and keiid='" & CStr(nhkcount) & "' and keishurui='3'"
          ElseIf shiteibankid = "" Then
              sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                      " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                      " and keiid='" & CStr(nhkcount) & "' and keishurui='3' and bankid='" & Space(1) & "'"
          Else
              sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                      " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                      " and keiid='" & CStr(nhkcount) & "' and keishurui='3' and bankid='" & shiteibankid & "'"
          End If
              If FcSQlGet2(rs_ch2, sql_ch2, PrMsg) = False Then
                   chdata(5, nowcount) = 0
              Else
                  If IsNull(rs_ch2!gkgk) Then
                      chdata(5, nowcount) = 0
                  Else
                      chdata(5, nowcount) = rs_ch2!gkgk
                      
                  End If
                  rs_ch2.Close
              End If
          'ì¸ã‡
          If shiteibankid = "-1" Then
              sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                      " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                      " and keiid='" & CStr(nhkcount) & "' and keishurui='3'"
          ElseIf shiteibankid = "" Then
              sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                      " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                      " and keiid='" & CStr(nhkcount) & "' and keishurui='3' and bankid='" & Space(1) & "'"
        ElseIf shiteibankid = "9999" Then
            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                    " and keiid='" & CStr(nhkcount) & "' and keishurui='3' and keist='2'"
          Else
              sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                      " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                      " and keiid='" & CStr(nhkcount) & "' and keishurui='3' and bankid='" & shiteibankid & "'"
          End If
              If FcSQlGet2(rs_ch3, sql_ch3, PrMsg) = False Then
                  chdata(6, nowcount) = 0
              Else
                  If IsNull(rs_ch3!gkgk) Then
                      chdata(6, nowcount) = 0
                  Else
                      chdata(6, nowcount) = rs_ch3!gkgk
                      sounyuukingaku = sounyuukingaku + rs_ch3!gkgk
                      souseikyuukingaku = souseikyuukingaku + rs_ch3!gkgk
                  End If
                  rs_ch3.Close
                                If shiteibankid = "9999" Then
                                    sql_sonkin = "select sum(kingaku) as torikeshikin from koushinkiroku" & _
                                                 " where koushinbi between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                                  " and houhou='7' and keiid='" & CStr(nhkcount) & "'"
                                    If FcSQlGet2(rs_sonkin, sql_sonkin, PrMsg) = False Then
                                        chdata(8, nowcount) = 0
                                    Else
                                        If IsNull(rs_sonkin!torikeshikin) Then
                                        chdata(8, nowcount) = 0
                                        Else
                                        chdata(8, nowcount) = rs_sonkin!torikeshikin
                                        End If
                                        rs_sonkin.Close
                                    End If
                                    sql_sonkin = "select sum(kingaku) as sonkin from koushinkiroku" & _
                                                 " where koushinbi between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                                  " and houhou='8' and keiid='" & CStr(nhkcount) & "'"
                                    If FcSQlGet2(rs_sonkin, sql_sonkin, PrMsg) = False Then
                                        chdata(9, nowcount) = 0
                                    Else
                                        If IsNull(rs_sonkin!sonkin) Then
                                        chdata(9, nowcount) = 0
                                        Else
                                        chdata(9, nowcount) = rs_sonkin!sonkin
                                        End If
                                        rs_sonkin.Close
                                    End If
                                End If
              End If
              
                      'åJâzäz
                            chdata(7, nowcount) = chdata(4, nowcount) + chdata(5, nowcount) - chdata(6, nowcount)
              
              nowcount = nowcount + 1
      Next
        
        
        
        
        
        
End With
FRMOMACHI.lblmsg.Caption = "èWåvÇÃï\é¶íÜ..."
        DoEvents

With frmsouken2
    chusercount = nowcount - 1
                For hhhajime = 1 To nowcount - 1
                    If frmsouken2.cmbbankid.ListIndex = -1 Then
                            Set itmX = .lv1.ListItems. _
                            Add(, , CStr(hhhajime))
                            
                            If Trim(chdata(1, hhhajime)) <> "" Then
                                itmX.SubItems(1) = CStr(chdata(0, hhhajime))
                            End If
                            
                            If Trim(chdata(1, hhhajime)) <> "" Then
                                itmX.SubItems(2) = CStr(chdata(1, hhhajime))
                            End If
                            If Trim(chdata(2, hhhajime)) <> "" Then
                                itmX.SubItems(3) = CStr(chdata(2, hhhajime))
                            End If
                            If IsNull(chdata(3, hhhajime)) Or Trim(chdata(3, hhhajime)) <> "" Then
                                itmX.SubItems(4) = CStr(chdata(3, hhhajime))
                            End If
                            If IsNull(chdata(4, hhhajime)) Or Trim(chdata(4, hhhajime)) <> "" Then
                                itmX.SubItems(5) = CStr(chdata(4, hhhajime))
                            End If
                            If IsNull(chdata(5, hhhajime)) Or Trim(chdata(5, hhhajime)) <> "" Then
                                itmX.SubItems(6) = CStr(chdata(5, hhhajime))
                            End If
                            If IsNull(chdata(6, hhhajime)) Or Trim(chdata(6, hhhajime)) <> "" Then
                                itmX.SubItems(7) = CStr(chdata(6, hhhajime))
                            End If
                            If IsNull(chdata(7, hhhajime)) Or Trim(chdata(7, hhhajime)) <> "" Then
                                itmX.SubItems(8) = CStr(chdata(7, hhhajime))
                            End If
                       Else
                            If shiteibankid = "9999" Then
                                    Set itmX = .lv1.ListItems. _
                                    Add(, , CStr(hhhajime))
                                    
                                    If Trim(chdata(1, hhhajime)) <> "" Then
                                        itmX.SubItems(1) = CStr(chdata(0, hhhajime))
                                    End If
                                    
                                    If Trim(chdata(1, hhhajime)) <> "" Then
                                        itmX.SubItems(2) = CStr(chdata(1, hhhajime))
                                    End If
                                    If Trim(chdata(2, hhhajime)) <> "" Then
                                        itmX.SubItems(3) = CStr(chdata(2, hhhajime))
                                    End If
                                    If IsNull(chdata(3, hhhajime)) Or Trim(chdata(3, hhhajime)) <> "" Then
                                        itmX.SubItems(4) = CStr(chdata(3, hhhajime))
                                    End If
                                    If IsNull(chdata(6, hhhajime)) Or Trim(chdata(6, hhhajime)) <> "" Then
                                        itmX.SubItems(5) = CStr(chdata(6, hhhajime))
                                    End If
                                    If IsNull(chdata(8, hhhajime)) Or Trim(chdata(8, hhhajime)) <> "" Then
                                        itmX.SubItems(6) = CStr(chdata(8, hhhajime))
                                    End If
                                    If IsNull(chdata(9, hhhajime)) Or Trim(chdata(9, hhhajime)) <> "" Then
                                        itmX.SubItems(7) = CStr(chdata(9, hhhajime))
                                    End If
                            Else
                                    Set itmX = .lv1.ListItems. _
                                    Add(, , CStr(hhhajime))
                                    
                                    If Trim(chdata(1, hhhajime)) <> "" Then
                                        itmX.SubItems(1) = CStr(chdata(0, hhhajime))
                                    End If
                                    
                                    If Trim(chdata(1, hhhajime)) <> "" Then
                                        itmX.SubItems(2) = CStr(chdata(1, hhhajime))
                                    End If
                                    If Trim(chdata(2, hhhajime)) <> "" Then
                                        itmX.SubItems(3) = CStr(chdata(2, hhhajime))
                                    End If
                                    If IsNull(chdata(3, hhhajime)) Or Trim(chdata(3, hhhajime)) <> "" Then
                                        itmX.SubItems(4) = CStr(chdata(3, hhhajime))
                                    End If
                                    If IsNull(chdata(6, hhhajime)) Or Trim(chdata(6, hhhajime)) <> "" Then
                                        itmX.SubItems(5) = CStr(chdata(6, hhhajime))
                                    End If
                            End If
                       End If
                        
                 Next
                Screen.MousePointer = 0
                .Picture15.Enabled = True
                .lblkekka.Caption = "éÂå_ñÒì¸ã‡çáåvã‡äzÅ@Å@Å@Å@ÉCÉìÉ^Å[ÉlÉbÉgånÅF" & kin_inter & "â~ÅAÇbÇ`ÇsÇuånÅF" & kin_catv & "â~" & Chr(13) & _
                "ÇªÇÃëºå_ñÒì¸ã‡çáåvã‡äzÅ@Å@ÉCÉìÉ^Å[ÉlÉbÉgånÅF" & kin_inter2 & "â~ÅAÇbÇ`ÇsÇuånÅF" & kin_catv2 & "â~" & Chr(13) & _
                "ÇmÇgÇjå_ñÒì¸ã‡çáåvã‡äzÅF" & souseikyuukingaku & "â~  ì¸ã‡ëçäzÅF" & sounyuukingaku & "â~"
                
                
                
           
            
End With
sentakusumidesuka = 1

        Unload FRMOMACHI
            Set FRMOMACHI = Nothing
           
            Screen.MousePointer = 0
            
             Me.Show 1
End Sub

Private Sub lv1_DblClick()

Dim selectid As String, selectshurui As String, shiteibankid5 As String, sql_ch35 As String, sql_s5 As String
Dim newnewshiteibi5 As String, newnewshiteibi15 As String, newnewshiteibi25 As String, rs_ch35 As Recordset
Dim rs_s5 As Recordset, nowgyou As Long, daimei As String


    
    
    

If sentakusumidesuka = 0 Then
    MsgBox "íäèoÇµÇƒÇ©ÇÁé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
    Exit Sub
End If
If frmsouken2.combzennen.ListIndex = -1 Then
    MsgBox "éwíËîNêîÇéwíËÇµÇƒÇ©ÇÁé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
    Exit Sub
End If
If frmsouken2.combzentsuki2.ListIndex = -1 Then
    MsgBox "éwíËåéêîÇéwíËÇµÇƒÇ©ÇÁé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
    Exit Sub
End If
If frmsouken2.cmbbankid.ListIndex = -1 Then
    shiteibankid5 = "-1"
Else
    shiteibankid5 = Trim(Mid(frmsouken2.cmbbankid.Text, 1, 4))
End If

newnewshiteibi5 = Format(frmsouken2.combzennen.Text, "0000") & Format(frmsouken2.combzentsuki2.Text, "00")
newnewshiteibi15 = Format(frmsouken2.combzennen.Text, "0000") & Format(frmsouken2.combzentsuki2.Text, "00") & "01"
newnewshiteibi25 = Format(frmsouken2.combzennen.Text, "0000") & Format(frmsouken2.combzentsuki2.Text, "00") & "31"

selectid = lv1.SelectedItem.SubItems(1)


If Mid(cmbbankid.Text, 1, 4) = "9999" Then
    
    If lv1.SelectedItem.SubItems(5) = "0" Then
        Exit Sub
    End If

        Select Case lv1.SelectedItem.SubItems(4)
        Case "ÇmÇgÇj"
                    sql_ch35 = "SELECT * FROM koushinkiroku" & _
                            " WHERE koushinbi between '" & newnewshiteibi15 & "' and '" & newnewshiteibi25 & "'" & _
                            " and keiid='" & selectid & "' and (houhou='7' or houhou='8')"
        Case "ÇªÇÃëº"
                    sql_ch35 = "SELECT * FROM koushinkiroku" & _
                            " WHERE koushinbi between '" & newnewshiteibi15 & "' and '" & newnewshiteibi25 & "'" & _
                            " and keiid='" & selectid & "' and (houhou='5' or houhou='6')"
        Case "éÂå_ñÒ"
                    sql_ch35 = "SELECT * FROM koushinkiroku" & _
                            " WHERE koushinbi between '" & newnewshiteibi15 & "' and '" & newnewshiteibi25 & "'" & _
                            " and keiid='" & selectid & "' and (houhou='3' or houhou='4')"
        Case Else
            MsgBox "ÅuéÂå_ñÒÅvÇ‹ÇΩÇÕÅuÇªÇÃëºÅvÇ‹ÇΩÇÕÅuÇmÇgÇjÅvÇéwíËÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End Select
        sql_ch35 = sql_ch35 & " order by kojinid"
        If FcSQlGet2(rs_ch35, sql_ch35, PrMsg) = True Then
            With frmsouken7
                .gridsonkin.Rows = rs_ch35.RecordCount + 1
                If shiteibankid5 = "-1" Then
                     daimei = "ÅiéÊè¡ã‡ÅEëπã‡ÅF" & rs_ch35.RecordCount & "åè )"
                Else
                     daimei = "ÅiéÊè¡ã‡ÅEëπã‡ÅF" & rs_ch35.RecordCount & "åè )"
                End If
                .gridsonkin.Row = 0
                .gridsonkin.Col = 0
                .gridsonkin.Text = "å¬êlÇhÇc"
                .gridsonkin.Col = 1
                .gridsonkin.Text = "å_ñÒé“ñº"
                .gridsonkin.Col = 2
                .gridsonkin.Text = "ã‡äz"
                .gridsonkin.Col = 3
                .gridsonkin.Text = "éÌóﬁ"
                .gridsonkin.ColWidth(0) = 800
                .gridsonkin.ColWidth(1) = 2500
                .gridsonkin.ColWidth(2) = 1000
                .gridsonkin.ColWidth(3) = 1000
                rs_ch35.MoveFirst
                nowgyou = 1
                Do Until rs_ch35.EOF
                    .gridsonkin.Row = nowgyou
                    .gridsonkin.Col = 0
                    .gridsonkin.Text = rs_ch35!kojinid
                    .gridsonkin.Col = 1
                    sql_s5 = "select name from kojin where kojinid='" & rs_ch35!kojinid & "'"
                    If FcSQlGet(rs_s5, sql_s5, PrMsg) = False Then
                        .gridsonkin.Text = "ÉGÉâÅ["
                    Else
                        .gridsonkin.Text = rs_s5!Name
                        rs_s5.Close
                    End If
                    .gridsonkin.Col = 2
                    .gridsonkin.Text = rs_ch35!kingaku
                    .gridsonkin.Col = 3
                    Select Case rs_ch35!houhou
                    Case "3", "5", "7"
                        .gridsonkin.Text = "éÊè¡ã‡"
                    Case "4", "6", "8"
                        .gridsonkin.Text = "ëπã‡"
                    End Select
                    nowgyou = nowgyou + 1
                    rs_ch35.MoveNext
                Loop
                rs_ch35.Close
            End With
        Else
                 If shiteibankid5 = "-1" Then
                    daimei = "ÅiéÊè¡ã‡ÅEëπã‡ÅF" & "0åè )"
                 Else
                    daimei = "ÅiéÊè¡ã‡ÅEëπã‡ÅF" & "0åè )"
                End If
        End If
        
        
        
               frmsouken7.Caption = daimei
                frmsouken7.Show 1

Else

        If shiteibankid5 = "-1" Then
            daimei = "è⁄ç◊   " & lv1.SelectedItem.SubItems(2) & "  ( åJâzã‡äzÅF" & lv1.SelectedItem.SubItems(5) & "â~,"
        
        
                    Select Case lv1.SelectedItem.SubItems(4)
                    Case "ÇªÇÃëº"
                            If shiteibankid5 = "-1" Then
                                sql_ch35 = "SELECT * FROM shousai" & _
                                        " WHERE keiseikyuu ='" & newnewshiteibi5 & "'" & _
                                        " and keiid='" & selectid & "' and keishurui='2'"
                            ElseIf shiteibankid5 = "" Then
                                sql_ch35 = "SELECT * FROM shousai" & _
                                        " WHERE keiseikyuu ='" & newnewshiteibi5 & "'" & _
                                        " and keiid='" & selectid & "' and keishurui='2' and bankid='" & Space(1) & "'"
                            Else
                                sql_ch35 = "SELECT * FROM shousai" & _
                                        " WHERE keiseikyuu ='" & newnewshiteibi5 & "'" & _
                                        " and keiid='" & selectid & "' and keishurui='2' and bankid='" & shiteibankid5 & "'"
                            End If
                    Case "éÂå_ñÒ"
                            If shiteibankid5 = "-1" Then
                                sql_ch35 = "SELECT * FROM shousai" & _
                                        " WHERE keiseikyuu ='" & newnewshiteibi5 & "'" & _
                                        " and keiid='" & selectid & "' and keishurui='1'"
                            ElseIf shiteibankid5 = "" Then
                                sql_ch35 = "SELECT * FROM shousai" & _
                                        " WHERE keiseikyuu ='" & newnewshiteibi5 & "'" & _
                                        " and keiid='" & selectid & "' and keishurui='1' and bankid='" & Space(1) & "'"
                            Else
                                sql_ch35 = "SELECT * FROM shousai" & _
                                        " WHERE keiseikyuu ='" & newnewshiteibi5 & "'" & _
                                        " and keiid='" & selectid & "' and keishurui='1' and bankid='" & shiteibankid5 & "'"
                            End If
                    Case "ÇmÇgÇj"
                            If shiteibankid5 = "-1" Then
                                sql_ch35 = "SELECT * FROM shousai" & _
                                        " WHERE keiseikyuu ='" & newnewshiteibi5 & "'" & _
                                        " and keiid='" & selectid & "' and keishurui='3'"
                            ElseIf shiteibankid5 = "" Then
                                sql_ch35 = "SELECT * FROM shousai" & _
                                        " WHERE keiseikyuu ='" & newnewshiteibi5 & "'" & _
                                        " and keiid='" & selectid & "' and keishurui='3' and bankid='" & Space(1) & "'"
                            Else
                                sql_ch35 = "SELECT * FROM shousai" & _
                                        " WHERE keiseikyuu ='" & newnewshiteibi5 & "'" & _
                                        " and keiid='" & selectid & "' and keishurui='3' and bankid='" & shiteibankid5 & "'"
                            End If
                    Case Else
                        MsgBox "ÅuéÂå_ñÒÅvÇ‹ÇΩÇÕÅuÇªÇÃëºÅvÇ‹ÇΩÇÕÅuÇmÇgÇjÅvÇéwíËÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
                        Exit Sub
                    End Select
                    sql_ch35 = sql_ch35 & " order by kojinid"
                    If FcSQlGet2(rs_ch35, sql_ch35, PrMsg) = True Then
                        With frmsouken3
                            .gridseikyuu.Rows = rs_ch35.RecordCount + 1
                            If shiteibankid5 = "-1" Then
                                daimei = daimei & "Å@êøãÅÅF" & rs_ch35.RecordCount & "åè , " & lv1.SelectedItem.SubItems(6) & "â~,"
                            End If
                            .gridseikyuu.Row = 0
                            .gridseikyuu.Col = 0
                            .gridseikyuu.Text = "å¬êlÇhÇc"
                            .gridseikyuu.Col = 1
                            .gridseikyuu.Text = "å_ñÒé“ñº"
                            .gridseikyuu.Col = 2
                            .gridseikyuu.Text = "ã‡äz"
                            .gridseikyuu.ColWidth(0) = 800
                            .gridseikyuu.ColWidth(1) = 2500
                            .gridseikyuu.ColWidth(2) = 1000
                            rs_ch35.MoveFirst
                            nowgyou = 1
                            Do Until rs_ch35.EOF
                                .gridseikyuu.Row = nowgyou
                                .gridseikyuu.Col = 0
                                .gridseikyuu.Text = rs_ch35!kojinid
                                .gridseikyuu.Col = 1
                                sql_s5 = "select name from kojin where kojinid='" & rs_ch35!kojinid & "'"
                                If FcSQlGet(rs_s5, sql_s5, PrMsg) = False Then
                                    .gridseikyuu.Text = "ÉGÉâÅ["
                                Else
                                    .gridseikyuu.Text = rs_s5!Name
                                    rs_s5.Close
                                End If
                                .gridseikyuu.Col = 2
                                .gridseikyuu.Text = rs_ch35!keikin
                                nowgyou = nowgyou + 1
                                rs_ch35.MoveNext
                            Loop
                            rs_ch35.Close
                        End With
                    Else
                        If shiteibankid5 = "-1" Then
                                daimei = daimei & "êøãÅÅF" & "0åè , " & lv1.SelectedItem.SubItems(6) & "â~,"
                        Else
                                daimei = daimei & "êøãÅÅF" & "0åè , " & lv1.SelectedItem.SubItems(5) & "â~,"
                        End If
                    End If
        
        Else
            daimei = "è⁄ç◊   " & lv1.SelectedItem.SubItems(2) & "  ( ì¸ã‡ã‡äzÅF" & lv1.SelectedItem.SubItems(5) & "â~,"
        End If
        
        
        
        Select Case lv1.SelectedItem.SubItems(4)
        
        Case "ÇªÇÃëº"
                If shiteibankid5 = "-1" Then
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keinyu between '" & newnewshiteibi15 & "' and '" & newnewshiteibi25 & "'" & _
                            " and keiid='" & selectid & "' and keishurui='2'"
                ElseIf shiteibankid5 = "" Then
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keinyu between '" & newnewshiteibi15 & "' and '" & newnewshiteibi25 & "'" & _
                            " and keiid='" & selectid & "' and keishurui='2' and bankid='" & Space(1) & "'"
                Else
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keinyu between '" & newnewshiteibi15 & "' and '" & newnewshiteibi25 & "'" & _
                            " and keiid='" & selectid & "' and keishurui='2' and bankid='" & shiteibankid5 & "'"
                End If
        Case "éÂå_ñÒ"
                If shiteibankid5 = "-1" Then
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keinyu between '" & newnewshiteibi15 & "' and '" & newnewshiteibi25 & "'" & _
                            " and keiid='" & selectid & "' and keishurui='1'"
                ElseIf shiteibankid5 = "" Then
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keinyu between '" & newnewshiteibi15 & "' and '" & newnewshiteibi25 & "'" & _
                            " and keiid='" & selectid & "' and keishurui='1' and bankid='" & Space(1) & "'"
                Else
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keinyu between '" & newnewshiteibi15 & "' and '" & newnewshiteibi25 & "'" & _
                            " and keiid='" & selectid & "' and keishurui='1' and bankid='" & shiteibankid5 & "'"
                End If
        Case "ÇmÇgÇj"
                If shiteibankid5 = "-1" Then
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keinyu between '" & newnewshiteibi15 & "' and '" & newnewshiteibi25 & "'" & _
                            " and keiid='" & selectid & "' and keishurui='3'"
                ElseIf shiteibankid5 = "" Then
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keinyu between '" & newnewshiteibi15 & "' and '" & newnewshiteibi25 & "'" & _
                            " and keiid='" & selectid & "' and keishurui='3' and bankid='" & Space(1) & "'"
                Else
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keinyu between '" & newnewshiteibi15 & "' and '" & newnewshiteibi25 & "'" & _
                            " and keiid='" & selectid & "' and keishurui='3' and bankid='" & shiteibankid5 & "'"
                End If
        Case Else
            MsgBox "ÅuéÂå_ñÒÅvÇ‹ÇΩÇÕÅuÇªÇÃëºÅvÇ‹ÇΩÇÕÅuÇmÇgÇjÅvÇéwíËÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End Select
        sql_ch35 = sql_ch35 & " order by kojinid"
        If FcSQlGet2(rs_ch35, sql_ch35, PrMsg) = True Then
            With frmsouken3
                .gridnyuukin.Rows = rs_ch35.RecordCount + 1
                If shiteibankid5 = "-1" Then
                     daimei = daimei & "ì¸ã‡ÅF" & rs_ch35.RecordCount & "åè ÅA"
                Else
                     daimei = daimei & "ì¸ã‡ÅF" & rs_ch35.RecordCount & "åè ÅA"
                End If
                .gridnyuukin.Row = 0
                .gridnyuukin.Col = 0
                .gridnyuukin.Text = "å¬êlÇhÇc"
                .gridnyuukin.Col = 1
                .gridnyuukin.Text = "å_ñÒé“ñº"
                .gridnyuukin.Col = 2
                .gridnyuukin.Text = "ã‡äz"
                .gridnyuukin.ColWidth(0) = 800
                .gridnyuukin.ColWidth(1) = 2500
                .gridnyuukin.ColWidth(2) = 1000
                rs_ch35.MoveFirst
                nowgyou = 1
                Do Until rs_ch35.EOF
                    .gridnyuukin.Row = nowgyou
                    .gridnyuukin.Col = 0
                    .gridnyuukin.Text = rs_ch35!kojinid
                    .gridnyuukin.Col = 1
                    sql_s5 = "select name from kojin where kojinid='" & rs_ch35!kojinid & "'"
                    If FcSQlGet(rs_s5, sql_s5, PrMsg) = False Then
                        .gridnyuukin.Text = "ÉGÉâÅ["
                    Else
                        .gridnyuukin.Text = rs_s5!Name
                        rs_s5.Close
                    End If
                    .gridnyuukin.Col = 2
                    .gridnyuukin.Text = rs_ch35!keikin
                    nowgyou = nowgyou + 1
                    rs_ch35.MoveNext
                Loop
                rs_ch35.Close
            End With
        Else
                 If shiteibankid5 = "-1" Then
                    daimei = daimei & "ì¸ã‡ÅF" & "0åèÅA"
                 Else
                    daimei = daimei & "ì¸ã‡ÅF" & "0åèÅA"
                End If
        End If
        
        
        Select Case lv1.SelectedItem.SubItems(4)
        
        Case "ÇªÇÃëº"
                If shiteibankid5 = "-1" Then
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keist='0'" & _
                            " and keiid='" & selectid & "' and keishurui='2'"
                ElseIf shiteibankid5 = "" Then
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keist='0'" & _
                            " and keiid='" & selectid & "' and keishurui='2' and bankid='" & Space(1) & "'"
                Else
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keist='0'" & _
                            " and keiid='" & selectid & "' and keishurui='2' and bankid='" & shiteibankid5 & "'"
                End If
        Case "éÂå_ñÒ"
                If shiteibankid5 = "-1" Then
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keist='0'" & _
                            " and keiid='" & selectid & "' and keishurui='1'"
                            
                      sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE (((shousai.keist)='1') AND ((shousai.keishurui)='1') AND ((shousai.keiid)='" & selectid & "')" & _
                            " AND ((shousai.keinyu)>'" & newnewshiteibi25 & "') AND ((shousai.keiseikyuu) <= '" & newnewshiteibi5 & "'))" & _
                            " OR (((shousai.keist)='0') AND ((shousai.keishurui)='1') AND ((shousai.keiid)='" & selectid & "') AND ((shousai.keiseikyuu) <= '" & newnewshiteibi5 & "'))"
                ElseIf shiteibankid5 = "" Then
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keist='0'" & _
                            " and keiid='" & selectid & "' and keishurui='1' and bankid='" & Space(1) & "'"
                            
                     sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE (((shousai.keist)='1') AND ((shousai.keishurui)='1') AND ((shousai.keiid)='" & selectid & "')" & _
                            " AND ((shousai.keinyu)>'" & newnewshiteibi25 & "') AND ((shousai.keiseikyuu) <= '" & newnewshiteibi5 & "') and shousai.bankid='" & Space(1) & "')" & _
                            " OR (((shousai.keist)='0') AND ((shousai.keishurui)='1') AND ((shousai.keiid)='" & selectid & "') AND ((shousai.keiseikyuu) <= '" & newnewshiteibi5 & "') and shousai.bankid='" & Space(1) & "')"
                Else
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keist='0'" & _
                            " and keiid='" & selectid & "' and keishurui='1' and bankid='" & shiteibankid5 & "'"
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE (((shousai.keist)='1') AND ((shousai.keishurui)='1') AND ((shousai.keiid)='" & selectid & "')" & _
                            " AND ((shousai.keinyu)>'" & newnewshiteibi25 & "') AND ((shousai.keiseikyuu) <= '" & newnewshiteibi5 & "') and shousai.bankid='" & shiteibankid5 & "')" & _
                            " OR (((shousai.keist)='0') AND ((shousai.keishurui)='1') AND ((shousai.keiid)='" & selectid & "') AND ((shousai.keiseikyuu) <= '" & newnewshiteibi5 & "') and shousai.bankid='" & shiteibankid5 & "')"
                End If
        Case "ÇmÇgÇj"
                If shiteibankid5 = "-1" Then
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keist='0'" & _
                            " and keiid='" & selectid & "' and keishurui='3'"
                ElseIf shiteibankid5 = "" Then
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keist='0'" & _
                            " and keiid='" & selectid & "' and keishurui='3' and bankid='" & Space(1) & "'"
                Else
                    sql_ch35 = "SELECT * FROM shousai" & _
                            " WHERE keist='0'" & _
                            " and keiid='" & selectid & "' and keishurui='3' and bankid='" & shiteibankid5 & "'"
                End If
        Case Else
            MsgBox "ÅuéÂå_ñÒÅvÇ‹ÇΩÇÕÅuÇªÇÃëºÅvÇ‹ÇΩÇÕÅuÇmÇgÇjÅvÇéwíËÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End Select
        sql_ch35 = sql_ch35 & " order by kojinid"
        If FcSQlGet2(rs_ch35, sql_ch35, PrMsg) = True Then
            With frmsouken3
                .gridminou.Rows = rs_ch35.RecordCount + 1
                If shiteibankid5 = "-1" Then
                     daimei = daimei & "ñ¢î[äzÅF" & rs_ch35.RecordCount & "åè )"
                Else
                     daimei = daimei & "ñ¢î[äzÅF" & rs_ch35.RecordCount & "åè )"
                End If
                .gridminou.Row = 0
                .gridminou.Col = 0
                .gridminou.Text = "å¬êlÇhÇc"
                .gridminou.Col = 1
                .gridminou.Text = "å_ñÒé“ñº"
                .gridminou.Col = 2
                .gridminou.Text = "ã‡äz"
                .gridminou.ColWidth(0) = 800
                .gridminou.ColWidth(1) = 2500
                .gridminou.ColWidth(2) = 1000
                rs_ch35.MoveFirst
                nowgyou = 1
                Do Until rs_ch35.EOF
                    .gridminou.Row = nowgyou
                    .gridminou.Col = 0
                    .gridminou.Text = rs_ch35!kojinid
                    .gridminou.Col = 1
                    sql_s5 = "select name from kojin where kojinid='" & rs_ch35!kojinid & "'"
                    If FcSQlGet(rs_s5, sql_s5, PrMsg) = False Then
                        .gridminou.Text = "ÉGÉâÅ["
                    Else
                        .gridminou.Text = rs_s5!Name
                        rs_s5.Close
                    End If
                    .gridminou.Col = 2
                    .gridminou.Text = rs_ch35!keikin
                    nowgyou = nowgyou + 1
                    rs_ch35.MoveNext
                Loop
                rs_ch35.Close
            End With
        Else
                 If shiteibankid5 = "-1" Then
                    daimei = daimei & "ñ¢î[äzÅF" & "0åè )"
                 Else
                    daimei = daimei & "ñ¢î[äzÅF" & "0åè )"
                End If
        End If
        
        
        
               frmsouken3.Caption = daimei
                frmsouken3.Show 1
End If
End Sub

Private Sub Picture1_Click()

        Dim shuwhere As String, karikindesu As Long, karikindesu2 As String, karihiduke As String
        Dim sql_ch As String, sql_ch2 As String, nowcount As Long
        Dim rs_ch As Recordset, rs_ch2 As Recordset, rs_ch3 As Recordset, sql_ch3 As String
        Dim kin_inter As Long, kin_catv As Long, nhkcount As Integer
        Dim kin_inter2 As Long, kin_catv2 As Long, shiteibankid As String
        Dim newnewshiteibi As String, newnewshiteibi1 As String, newnewshiteibi2 As String
        Dim sounyuukingaku As Long, souseikyuukingaku As Long
        Dim sql_ch8 As String, rs_ch8 As Recordset
        Dim sql_ch9 As String, rs_ch9 As Recordset
        Dim newnewshiteibi5 As String, newnewshiteibi52 As String
        Dim sql_sonkin As String, rs_sonkin As Recordset
        
        chusercount = 0
        souseikyuukingaku = 0
        sounyuukingaku = 0
        frmsouken2.lv1.ListItems.Clear
        frmsouken2.lv1.ColumnHeaders.Clear
        DoEvents
        
        If frmsouken2.combzennen.ListIndex = -1 Then
            MsgBox "éwíËîNêîÇéwíËÇµÇƒÇ©ÇÁé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
        If frmsouken2.combzentsuki2.ListIndex = -1 Then
            MsgBox "éwíËåéêîÇéwíËÇµÇƒÇ©ÇÁé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
            Exit Sub
        End If
        If frmsouken2.cmbbankid.ListIndex = -1 Then
            shiteibankid = "-1"
        Else
            shiteibankid = Trim(Mid(frmsouken2.cmbbankid.Text, 1, 4))
        End If
        
        Me.Hide
       FRMOMACHI.lblmsg.Caption = "èWåvÇÃèÄîıíÜ..."
       FRMOMACHI.pb.Visible = True
        FRMOMACHI.pb.Value = 0
        FRMOMACHI.Show
        DoEvents
        Screen.MousePointer = 11
        
        newnewshiteibi = Format(frmsouken2.combzennen.Text, "0000") & Format(frmsouken2.combzentsuki2.Text, "00")
        newnewshiteibi1 = Format(frmsouken2.combzennen.Text, "0000") & Format(frmsouken2.combzentsuki2.Text, "00") & "01"
        newnewshiteibi2 = Format(frmsouken2.combzennen.Text, "0000") & Format(frmsouken2.combzentsuki2.Text, "00") & "31"
        If frmsouken2.combzentsuki2.Text = "01" Then
            newnewshiteibi5 = Format(CLng(frmsouken2.combzennen.Text) - 1, "0000") & "12"
        Else
            newnewshiteibi5 = Format(frmsouken2.combzennen.Text, "0000") & Format(CInt(frmsouken2.combzentsuki2.Text) - 1, "00")
        End If
        With frmsouken2
        If frmsouken2.cmbbankid.ListIndex = -1 Then
                .lv1.Checkboxes = False
                   .lv1.ColumnHeaders. _
                   Add , , "NO", 500
                   
                   .lv1.ColumnHeaders. _
                   Add , , "id", 800
                   
                   .lv1.ColumnHeaders. _
                   Add , , "å_ñÒñº", 5000
                   
                   .lv1.ColumnHeaders. _
                   Add , , "éÌï ", 1200
                   
                   .lv1.ColumnHeaders. _
                   Add , , "éÌóﬁ", 1200
                          
                   .lv1.ColumnHeaders. _
                   Add , , "åJâzäz", 1400, 1
                   
                   .lv1.ColumnHeaders. _
                   Add , , "êøãÅäz", 1400, 1
                          
                   .lv1.ColumnHeaders. _
                   Add , , "ì¸ã‡äz", 1400, 1
                   
                   .lv1.ColumnHeaders. _
                   Add , , "ñ¢î[äz", 1400, 1
        Else
            If shiteibankid = "9999" Then
                .lv1.Checkboxes = False
                   .lv1.ColumnHeaders. _
                   Add , , "NO", 500
                   
                   .lv1.ColumnHeaders. _
                   Add , , "id", 800
                   
                   .lv1.ColumnHeaders. _
                   Add , , "å_ñÒñº", 5000
                   
                   .lv1.ColumnHeaders. _
                   Add , , "éÌï ", 1200
                   
                   .lv1.ColumnHeaders. _
                   Add , , "éÌóﬁ", 1200
                          
                   .lv1.ColumnHeaders. _
                   Add , , "çáåväz", 1400, 1
            
                   .lv1.ColumnHeaders. _
                   Add , , "éÊè¡ã‡äz", 1400, 1
            
                   .lv1.ColumnHeaders. _
                   Add , , "ëπé∏ã‡äz", 1400, 1
            Else
                .lv1.Checkboxes = False
                   .lv1.ColumnHeaders. _
                   Add , , "NO", 500
                   
                   .lv1.ColumnHeaders. _
                   Add , , "id", 800
                   
                   .lv1.ColumnHeaders. _
                   Add , , "å_ñÒñº", 5000
                   
                   .lv1.ColumnHeaders. _
                   Add , , "éÌï ", 1200
                   
                   .lv1.ColumnHeaders. _
                   Add , , "éÌóﬁ", 1200
                          
                   .lv1.ColumnHeaders. _
                   Add , , "ì¸ã‡äz", 1400, 1
             End If
        End If
        End With
        
    ' ListItem ÉIÉuÉWÉFÉNÉgÇí«â¡Ç∑ÇÈÇΩÇﬂÇÃïœêîÇêÈåæÇµÇ‹Ç∑ÅB
    Dim itmX As ListItem, kakikubun As String, hhhajime As Long
         kin_inter = 0
         kin_catv = 0
         kin_inter2 = 0
         kin_catv2 = 0
         nowcount = 1
With frmsouken2
        FRMOMACHI.lblmsg.Caption = "éÌóﬁÇÃèWåvíÜ..."
        DoEvents
       sql_ch = "SELECT * FROM shurui order by shuruiid"
        If FcSQlGet(rs_ch, sql_ch, PrMsg) = True Then
            If kidou2 = False Then
                 MsgBox "ÉfÅ[É^ÉxÅ[ÉXÉtÉ@ÉCÉãÇäJÇØÇ‹ÇπÇÒÇ≈ÇµÇΩÅB"
                    Unload FRMOMACHI
            Set FRMOMACHI = Nothing
            Screen.MousePointer = 0
                 Exit Sub
             End If
             
             
             FRMOMACHI.pb.Min = 0
                FRMOMACHI.pb.Max = rs_ch.RecordCount + 1
             FRMOMACHI.pb.Value = 0
             DoEvents
             
               rs_ch.MoveFirst
               While Not rs_ch.EOF
                    ReDim Preserve chdata(17, nowcount)
                    
                      chdata(0, nowcount) = rs_ch!shuruiid
                      chdata(1, nowcount) = rs_ch!shuruimei
                      chdata(17, nowcount) = rs_ch!tanka
                      'í«â¡
                       If IsNull(rs_ch!iru) Then
                        chdata(16, nowcount) = ""
                        Else
                         chdata(16, nowcount) = rs_ch!iru
                        End If
                        If IsNull(rs_ch!gp) Then
                         chdata(14, nowcount) = ""
                        Else
                         chdata(14, nowcount) = rs_ch!gp
                        End If
                         If IsNull(rs_ch!narabe) Then
                         chdata(15, nowcount) = ""
                        Else
                        chdata(15, nowcount) = rs_ch!narabe
                        End If
                      
                      FRMOMACHI.Caption = "èàóùíÜÅEÅE(" & chdata(0, nowcount) & ")1"
                       DoEvents
                      
                      If rs_ch!kankei = "0" Then
                            chdata(2, nowcount) = "CATV"
                      Else
                            chdata(2, nowcount) = "Inter"
                      End If
                      chdata(3, nowcount) = "éÂå_ñÒ"
                      'åJâzäz
                            If shiteibankid = "-1" Then
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='1' AND shousai.keiid ='" & rs_ch!shuruiid & "'"
                            ElseIf shiteibankid = "" Then
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='1' AND shousai.keiid ='" & rs_ch!shuruiid & "'" & _
                                    " and shousai.bankid='" & Space(1) & "'"
                            Else
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='1' AND shousai.keiid ='" & rs_ch!shuruiid & "'" & _
                                    " and shousai.bankid='" & shiteibankid & "'"
                            End If
                                
                            If FcSQlGet2(rs_ch8, sql_ch8, PrMsg) = False Then
                                 chdata(4, nowcount) = 0
                            Else
                                If IsNull(rs_ch8!kurigou) Then
                                    chdata(4, nowcount) = 0
                                Else
                                    chdata(4, nowcount) = rs_ch8!kurigou
                                End If
                                rs_ch8.Close
                            End If

                        FRMOMACHI.Caption = "èàóùíÜÅEÅE(" & chdata(0, nowcount) & ")2"
                       DoEvents
                      'êøãÅäz
                            If shiteibankid = "-1" Then
                                sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                        " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                                        " and keiid='" & rs_ch!shuruiid & "' and keishurui='1'"
                            ElseIf shiteibankid = "" Then
                                sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                        " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                                        " and keiid='" & rs_ch!shuruiid & "' and keishurui='1' and bankid='" & Space(1) & "'"
                            Else
                                sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                        " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                                        " and keiid='" & rs_ch!shuruiid & "' and keishurui='1' and bankid='" & shiteibankid & "'"
                            End If
                            If FcSQlGet2(rs_ch2, sql_ch2, PrMsg) = False Then
                                 chdata(5, nowcount) = 0
                            Else
                                If IsNull(rs_ch2!gkgk) Then
                                    chdata(5, nowcount) = 0
                                Else
                                    chdata(5, nowcount) = rs_ch2!gkgk
                                    
                                End If
                                rs_ch2.Close
                            End If
                              FRMOMACHI.Caption = "èàóùíÜÅEÅE(" & chdata(0, nowcount) & ")3"
                       DoEvents
                        'ì¸ã‡
                        If shiteibankid = "-1" Then
                            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                    " and keiid='" & rs_ch!shuruiid & "' and keishurui='1'"
                        ElseIf shiteibankid = "" Then
                            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                    " and keiid='" & rs_ch!shuruiid & "' and keishurui='1' and bankid='" & Space(1) & "'"
                        ElseIf shiteibankid = "9999" Then
                            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                    " and keiid='" & rs_ch!shuruiid & "' and keishurui='1' and keist='2'"
                        Else
                            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                    " and keiid='" & rs_ch!shuruiid & "' and keishurui='1' and bankid='" & shiteibankid & "'"
                        End If
                            If FcSQlGet2(rs_ch3, sql_ch3, PrMsg) = False Then
                                chdata(6, nowcount) = 0
                            Else
                                If IsNull(rs_ch3!gkgk) Then
                                    chdata(6, nowcount) = 0
                                Else
                                    chdata(6, nowcount) = rs_ch3!gkgk
                                    sounyuukingaku = sounyuukingaku + rs_ch3!gkgk
                                End If
                                rs_ch3.Close
                                  FRMOMACHI.Caption = "èàóùíÜÅEÅE(" & chdata(0, nowcount) & ")3-1"
                       DoEvents
                                If shiteibankid = "9999" Then
                                    sql_sonkin = "select sum(kingaku) as torikeshikin from koushinkiroku" & _
                                                 " where koushinbi between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                                  " and houhou='3' and keiid='" & rs_ch!shuruiid & "'"
                                    If FcSQlGet2(rs_sonkin, sql_sonkin, PrMsg) = False Then
                                        chdata(8, nowcount) = 0
                                    Else
                                        If IsNull(rs_sonkin!torikeshikin) Then
                                        chdata(8, nowcount) = 0
                                        Else
                                        chdata(8, nowcount) = rs_sonkin!torikeshikin
                                        End If
                                        rs_sonkin.Close
                                    End If
                                      FRMOMACHI.Caption = "èàóùíÜÅEÅE(" & chdata(0, nowcount) & ")3-2"
                       DoEvents
                                    sql_sonkin = "select sum(kingaku) as sonkin from koushinkiroku" & _
                                                 " where koushinbi between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                                  " and houhou='4' and keiid='" & rs_ch!shuruiid & "'"
                                    If FcSQlGet2(rs_sonkin, sql_sonkin, PrMsg) = False Then
                                        chdata(9, nowcount) = 0
                                    Else
                                        If IsNull(rs_sonkin!sonkin) Then
                                        chdata(9, nowcount) = 0
                                        Else
                                        chdata(9, nowcount) = rs_sonkin!sonkin
                                        End If
                                        rs_sonkin.Close
                                    End If
                                End If
                            End If
                            If chdata(2, nowcount) = "Inter" Then
                                kin_inter = kin_inter + CLng(chdata(6, nowcount))
                            Else
                                kin_catv = kin_catv + CLng(chdata(6, nowcount))
                            End If
                            
                      'åJâzäz
                        FRMOMACHI.Caption = "èàóùíÜÅEÅE(" & chdata(0, nowcount) & ")4"
                       DoEvents
                            chdata(7, nowcount) = chdata(4, nowcount) + chdata(5, nowcount) - chdata(6, nowcount)
                            
                            nowcount = nowcount + 1
                            
                              FRMOMACHI.pb.Value = nowcount
                                DoEvents
                   rs_ch.MoveNext   ' éüÇÃÉåÉRÅ[ÉhÇ…à⁄ìÆÇµÇ‹Ç∑ÅB
                Wend
                rs_ch.Close
        End If
        FRMOMACHI.lblmsg.Caption = "å_ñÒÇÃèWåvíÜ..."
        DoEvents
        
        Dim nowcount2 As Integer
        nowcount2 = 0
        
       sql_ch = "SELECT * FROM sonota order by sonotaid"
        If FcSQlGet(rs_ch, sql_ch, PrMsg) = True Then
            If kidou2 = False Then
                 MsgBox "ÉfÅ[É^ÉxÅ[ÉXÉtÉ@ÉCÉãÇäJÇØÇ‹ÇπÇÒÇ≈ÇµÇΩÅB"
                    Unload FRMOMACHI
            Set FRMOMACHI = Nothing
            Screen.MousePointer = 0
             
                 Exit Sub
             End If
             
              FRMOMACHI.pb.Min = 0
                FRMOMACHI.pb.Max = rs_ch.RecordCount + 1
             FRMOMACHI.pb.Value = 0
             DoEvents
             
               rs_ch.MoveFirst
               While Not rs_ch.EOF
                    ReDim Preserve chdata(17, nowcount)
                      chdata(0, nowcount) = rs_ch!sonotaid
                      chdata(1, nowcount) = rs_ch!sonota
                      
                       chdata(17, nowcount) = rs_ch!kingaku
                       
                      'í«â¡
                       If IsNull(rs_ch!iru) Then
                        chdata(16, nowcount) = ""
                        Else
                         chdata(16, nowcount) = rs_ch!iru
                        End If
                        If IsNull(rs_ch!gp) Then
                         chdata(14, nowcount) = ""
                        Else
                         chdata(14, nowcount) = rs_ch!gp
                        End If
                         If IsNull(rs_ch!narabe) Then
                         chdata(15, nowcount) = ""
                        Else
                        chdata(15, nowcount) = rs_ch!narabe
                        End If
                      
                      
                      
                      If rs_ch!kankei = "0" Then
                            chdata(2, nowcount) = "CATV"
                      Else
                            chdata(2, nowcount) = "Inter"
                      End If
                      chdata(3, nowcount) = "ÇªÇÃëº"
                      'åJâzäz
                            If shiteibankid = "-1" Then
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='2' AND shousai.keiid ='" & rs_ch!sonotaid & "'"
                            ElseIf shiteibankid = "" Then
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='2' AND shousai.keiid ='" & rs_ch!sonotaid & "'" & _
                                    " and shousai.bankid='" & Space(1) & "'"
                            Else
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='2' AND shousai.keiid ='" & rs_ch!sonotaid & "'" & _
                                    " and shousai.bankid='" & shiteibankid & "'"
                            End If
                                
                            If FcSQlGet2(rs_ch8, sql_ch8, PrMsg) = False Then
                                 chdata(4, nowcount) = 0
                            Else
                                If IsNull(rs_ch8!kurigou) Then
                                    chdata(4, nowcount) = 0
                                Else
                                    chdata(4, nowcount) = rs_ch8!kurigou
                                End If
                                rs_ch8.Close
                            End If
                      
                      'êøãÅäz
                        If shiteibankid = "-1" Then
                            sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                                    " and keiid='" & rs_ch!sonotaid & "' and keishurui='2'"
                        ElseIf shiteibankid = "" Then
                            sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                                    " and keiid='" & rs_ch!sonotaid & "' and keishurui='2' and bankid='" & Space(1) & "'"
                        Else
                            sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                                    " and keiid='" & rs_ch!sonotaid & "' and keishurui='2' and bankid='" & shiteibankid & "'"
                        End If
                            If FcSQlGet2(rs_ch2, sql_ch2, PrMsg) = False Then
                                 chdata(5, nowcount) = 0
                            Else
                                If IsNull(rs_ch2!gkgk) Then
                                    chdata(5, nowcount) = 0
                                Else
                                    chdata(5, nowcount) = rs_ch2!gkgk
                                    
                                End If
                                rs_ch2.Close
                            End If
                        'ì¸ã‡
                        If shiteibankid = "-1" Then
                            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                    " and keiid='" & rs_ch!sonotaid & "' and keishurui='2'"
                        ElseIf shiteibankid = "" Then
                            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                    " and keiid='" & rs_ch!sonotaid & "' and keishurui='2' and bankid='" & Space(1) & "'"
                        ElseIf shiteibankid = "9999" Then
                            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                    " and keiid='" & rs_ch!sonotaid & "' and keishurui='2' and keist='2'"
                        Else
                            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                    " and keiid='" & rs_ch!sonotaid & "' and keishurui='2' and bankid='" & shiteibankid & "'"
                        End If
                            If FcSQlGet2(rs_ch3, sql_ch3, PrMsg) = False Then
                                chdata(6, nowcount) = 0
                            Else
                                If IsNull(rs_ch3!gkgk) Then
                                    chdata(6, nowcount) = 0
                                Else
                                    chdata(6, nowcount) = rs_ch3!gkgk
                                    sounyuukingaku = sounyuukingaku + rs_ch3!gkgk
                                End If
                                rs_ch3.Close
                                If shiteibankid = "9999" Then
                                    sql_sonkin = "select sum(kingaku) as torikeshikin from koushinkiroku" & _
                                                 " where koushinbi between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                                  " and houhou='5' and keiid='" & rs_ch!sonotaid & "'"
                                    If FcSQlGet2(rs_sonkin, sql_sonkin, PrMsg) = False Then
                                        chdata(8, nowcount) = 0
                                    Else
                                         If IsNull(rs_sonkin!torikeshikin) Then
                                        chdata(8, nowcount) = 0
                                        Else
                                       chdata(8, nowcount) = rs_sonkin!torikeshikin
                                        End If
                                        rs_sonkin.Close
                                    End If
                                    sql_sonkin = "select sum(kingaku) as sonkin from koushinkiroku" & _
                                                 " where koushinbi between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                                  " and houhou='6' and keiid='" & rs_ch!sonotaid & "'"
                                    If FcSQlGet2(rs_sonkin, sql_sonkin, PrMsg) = False Then
                                        chdata(9, nowcount) = 0
                                    Else
                                        If IsNull(rs_sonkin!sonkin) Then
                                        chdata(9, nowcount) = 0
                                        Else
                                        chdata(9, nowcount) = rs_sonkin!sonkin
                                        End If
                                        rs_sonkin.Close
                                    End If
                                End If
                            End If
                            If chdata(2, nowcount) = "Inter" Then
                                kin_inter2 = kin_inter2 + CLng(chdata(6, nowcount))
                            Else
                                kin_catv2 = kin_catv2 + CLng(chdata(6, nowcount))
                            End If
                            
                      'åJâzäz
                            chdata(7, nowcount) = chdata(4, nowcount) + chdata(5, nowcount) - chdata(6, nowcount)
                            
                            
                            
                            nowcount = nowcount + 1
                            nowcount2 = nowcount2 + 1
                            
                               FRMOMACHI.pb.Value = nowcount2
                                DoEvents
                                
                   rs_ch.MoveNext   ' éüÇÃÉåÉRÅ[ÉhÇ…à⁄ìÆÇµÇ‹Ç∑ÅB
                Wend
                rs_ch.Close
        End If
        
    FRMOMACHI.lblmsg.Caption = "HNKÇÃèWåvíÜ..."
        DoEvents
     For nhkcount = 0 To 7
        ReDim Preserve chdata(17, nowcount)
        chdata(0, nowcount) = CStr(nhkcount)
        chdata(17, nowcount) = ""
        Select Case nhkcount
        Case 0
            chdata(1, nowcount) = "ÇmÇgÇjÇPÉñåé"
        Case 1
            chdata(1, nowcount) = "ÇmÇgÇjÇQÉñåé"
        Case 2
            chdata(1, nowcount) = "ÇmÇgÇjÇRÉñåé"
        Case 3
            chdata(1, nowcount) = "ÇmÇgÇjÇSÉñåé"
        Case 4
            chdata(1, nowcount) = "ÇmÇgÇjÇTÉñåé"
        Case 5
            chdata(1, nowcount) = "ÇmÇgÇjÇUÉñåé"
             chdata(17, nowcount) = CStr(nhk_ryoukin_tsuujou)
        Case 6
            chdata(1, nowcount) = "ÇmÇgÇjÇUÉñåéÉnÉìÉf"
            chdata(17, nowcount) = CStr(nhk_ryoukin_hangaku)
        Case 7
            chdata(1, nowcount) = "ÇmÇgÇjÇUÉñåéâ∆ë∞äÑ"
            chdata(17, nowcount) = CStr(nhk_ryoukin_hangaku)
         End Select
        chdata(2, nowcount) = "ÇmÇgÇj"
        chdata(3, nowcount) = "ÇmÇgÇj"
        
                      'åJâzäz
                            If shiteibankid = "-1" Then
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='3' AND shousai.keiid ='" & CStr(nhkcount) & "'"
                            ElseIf shiteibankid = "" Then
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='3' AND shousai.keiid ='" & CStr(nhkcount) & "'" & _
                                    " and shousai.bankid='" & Space(1) & "'"
                            Else
                                sql_ch8 = "SELECT Sum(shousai.keikin) AS kurigou" & _
                                    " From shousai" & _
                                    " WHERE shousai.keiseikyuu <='" & newnewshiteibi5 & "'" & _
                                    " AND (shousai.keinyu>'" & newnewshiteibi & "00" & "' Or shousai.keinyu =Space(1))" & _
                                    " AND shousai.keishurui='3' AND shousai.keiid ='" & CStr(nhkcount) & "'" & _
                                    " and shousai.bankid='" & shiteibankid & "'"
                            End If
                                
                            If FcSQlGet2(rs_ch8, sql_ch8, PrMsg) = False Then
                                 chdata(4, nowcount) = 0
                            Else
                                If IsNull(rs_ch8!kurigou) Then
                                    chdata(4, nowcount) = 0
                                Else
                                    chdata(4, nowcount) = rs_ch8!kurigou
                                End If
                                rs_ch8.Close
                            End If
        
        
        'êøãÅäz
          If shiteibankid = "-1" Then
              sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                      " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                      " and keiid='" & CStr(nhkcount) & "' and keishurui='3'"
          ElseIf shiteibankid = "" Then
              sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                      " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                      " and keiid='" & CStr(nhkcount) & "' and keishurui='3' and bankid='" & Space(1) & "'"
          Else
              sql_ch2 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                      " WHERE keiseikyuu ='" & newnewshiteibi & "'" & _
                      " and keiid='" & CStr(nhkcount) & "' and keishurui='3' and bankid='" & shiteibankid & "'"
          End If
              If FcSQlGet2(rs_ch2, sql_ch2, PrMsg) = False Then
                   chdata(5, nowcount) = 0
              Else
                  If IsNull(rs_ch2!gkgk) Then
                      chdata(5, nowcount) = 0
                  Else
                      chdata(5, nowcount) = rs_ch2!gkgk
                      
                  End If
                  rs_ch2.Close
              End If
          'ì¸ã‡
          If shiteibankid = "-1" Then
              sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                      " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                      " and keiid='" & CStr(nhkcount) & "' and keishurui='3'"
          ElseIf shiteibankid = "" Then
              sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                      " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                      " and keiid='" & CStr(nhkcount) & "' and keishurui='3' and bankid='" & Space(1) & "'"
        ElseIf shiteibankid = "9999" Then
            sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                    " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                    " and keiid='" & CStr(nhkcount) & "' and keishurui='3' and keist='2'"
          Else
              sql_ch3 = "SELECT sum(clng(keikin)) as gkgk FROM shousai" & _
                      " WHERE keinyu between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                      " and keiid='" & CStr(nhkcount) & "' and keishurui='3' and bankid='" & shiteibankid & "'"
          End If
              If FcSQlGet2(rs_ch3, sql_ch3, PrMsg) = False Then
                  chdata(6, nowcount) = 0
              Else
                  If IsNull(rs_ch3!gkgk) Then
                      chdata(6, nowcount) = 0
                  Else
                      chdata(6, nowcount) = rs_ch3!gkgk
                      sounyuukingaku = sounyuukingaku + rs_ch3!gkgk
                      souseikyuukingaku = souseikyuukingaku + rs_ch3!gkgk
                  End If
                  rs_ch3.Close
                                If shiteibankid = "9999" Then
                                    sql_sonkin = "select sum(kingaku) as torikeshikin from koushinkiroku" & _
                                                 " where koushinbi between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                                  " and houhou='7' and keiid='" & CStr(nhkcount) & "'"
                                    If FcSQlGet2(rs_sonkin, sql_sonkin, PrMsg) = False Then
                                        chdata(8, nowcount) = 0
                                    Else
                                        If IsNull(rs_sonkin!torikeshikin) Then
                                        chdata(8, nowcount) = 0
                                        Else
                                        chdata(8, nowcount) = rs_sonkin!torikeshikin
                                        End If
                                        rs_sonkin.Close
                                    End If
                                    sql_sonkin = "select sum(kingaku) as sonkin from koushinkiroku" & _
                                                 " where koushinbi between '" & newnewshiteibi1 & "' and '" & newnewshiteibi2 & "'" & _
                                                  " and houhou='8' and keiid='" & CStr(nhkcount) & "'"
                                    If FcSQlGet2(rs_sonkin, sql_sonkin, PrMsg) = False Then
                                        chdata(9, nowcount) = 0
                                    Else
                                        If IsNull(rs_sonkin!sonkin) Then
                                        chdata(9, nowcount) = 0
                                        Else
                                        chdata(9, nowcount) = rs_sonkin!sonkin
                                        End If
                                        rs_sonkin.Close
                                    End If
                                End If
              End If
              
                      'åJâzäz
                            chdata(7, nowcount) = chdata(4, nowcount) + chdata(5, nowcount) - chdata(6, nowcount)
              
              nowcount = nowcount + 1
      Next
        
        
        
        
        
        
        
        
End With
FRMOMACHI.lblmsg.Caption = "èWåvÇÃï\é¶íÜ..."
        DoEvents
        
        
 'äeåèêîÇåvéZ
 Dim n_tanka As Long, n_kurikoshi As Long, n_seikyuu As Long, n_nyuukin As Long, n_mishuukin As Long
 For hhhajime = 1 To nowcount - 1
    'íPâøÇ™Ç†ÇÈ
    If Trim(chdata(17, hhhajime)) <> "" Then
        n_tanka = CLng(Trim(chdata(17, hhhajime)))
        
        n_kurikoshi = CLng(Trim(chdata(4, hhhajime)))
        n_seikyuu = CLng(Trim(chdata(5, hhhajime)))
        n_nyuukin = CLng(Trim(chdata(6, hhhajime)))
        n_mishuukin = CLng(Trim(chdata(7, hhhajime)))
        
        If n_tanka = 0 Then
            chdata(10, hhhajime) = CStr(n_kurikoshi)
            chdata(11, hhhajime) = CStr(n_seikyuu)
            chdata(12, hhhajime) = CStr(n_nyuukin)
           chdata(13, hhhajime) = CStr(n_mishuukin)
        Else
            If n_kurikoshi = 0 Then
                 chdata(10, hhhajime) = "0"
                  chdata(11, hhhajime) = "0"
                 chdata(12, hhhajime) = "0"
                chdata(13, hhhajime) = "0"
            Else
                chdata(10, hhhajime) = CStr(n_kurikoshi / n_tanka)
                 chdata(11, hhhajime) = CStr(n_seikyuu / n_tanka)
                chdata(12, hhhajime) = CStr(n_nyuukin / n_tanka)
                 chdata(13, hhhajime) = CStr(n_mishuukin / n_tanka)
                
            End If
         End If
        
    Else
        chdata(10, hhhajime) = ""
         chdata(11, hhhajime) = ""
          chdata(12, hhhajime) = ""
           chdata(13, hhhajime) = ""
    End If
 Next
        
        
        

With frmsouken2
    chusercount = nowcount - 1
                For hhhajime = 1 To nowcount - 1
                    If frmsouken2.cmbbankid.ListIndex = -1 Then
                            Set itmX = .lv1.ListItems. _
                            Add(, , CStr(hhhajime))
                            
                            If Trim(chdata(0, hhhajime)) <> "" Then
                                itmX.SubItems(1) = CStr(chdata(0, hhhajime))
                            End If
                            
                            If Trim(chdata(1, hhhajime)) <> "" Then
                                itmX.SubItems(2) = CStr(chdata(1, hhhajime))
                            End If
                            If Trim(chdata(2, hhhajime)) <> "" Then
                                itmX.SubItems(3) = CStr(chdata(2, hhhajime))
                            End If
                            If IsNull(chdata(3, hhhajime)) Or Trim(chdata(3, hhhajime)) <> "" Then
                                itmX.SubItems(4) = CStr(chdata(3, hhhajime))
                            End If
                            If IsNull(chdata(4, hhhajime)) Or Trim(chdata(4, hhhajime)) <> "" Then
                                itmX.SubItems(5) = CStr(chdata(4, hhhajime))
                            End If
                            If IsNull(chdata(5, hhhajime)) Or Trim(chdata(5, hhhajime)) <> "" Then
                                itmX.SubItems(6) = CStr(chdata(5, hhhajime))
                            End If
                            If IsNull(chdata(6, hhhajime)) Or Trim(chdata(6, hhhajime)) <> "" Then
                                itmX.SubItems(7) = CStr(chdata(6, hhhajime))
                            End If
                            If IsNull(chdata(7, hhhajime)) Or Trim(chdata(7, hhhajime)) <> "" Then
                                itmX.SubItems(8) = CStr(chdata(7, hhhajime))
                            End If
                       Else
                            If shiteibankid = "9999" Then
                                    Set itmX = .lv1.ListItems. _
                                    Add(, , CStr(hhhajime))
                                    
                                    If Trim(chdata(0, hhhajime)) <> "" Then
                                        itmX.SubItems(1) = CStr(chdata(0, hhhajime))
                                    End If
                                    
                                    If Trim(chdata(1, hhhajime)) <> "" Then
                                        itmX.SubItems(2) = CStr(chdata(1, hhhajime))
                                    End If
                                    If Trim(chdata(2, hhhajime)) <> "" Then
                                        itmX.SubItems(3) = CStr(chdata(2, hhhajime))
                                    End If
                                    If IsNull(chdata(3, hhhajime)) Or Trim(chdata(3, hhhajime)) <> "" Then
                                        itmX.SubItems(4) = CStr(chdata(3, hhhajime))
                                    End If
                                    If IsNull(chdata(6, hhhajime)) Or Trim(chdata(6, hhhajime)) <> "" Then
                                        itmX.SubItems(5) = CStr(chdata(6, hhhajime))
                                    End If
                                    If IsNull(chdata(8, hhhajime)) Or Trim(chdata(8, hhhajime)) <> "" Then
                                        itmX.SubItems(6) = CStr(chdata(8, hhhajime))
                                    End If
                                    If IsNull(chdata(9, hhhajime)) Or Trim(chdata(9, hhhajime)) <> "" Then
                                        itmX.SubItems(7) = CStr(chdata(9, hhhajime))
                                    End If
                            Else
                                    Set itmX = .lv1.ListItems. _
                                    Add(, , CStr(hhhajime))
                                    
                                    If Trim(chdata(0, hhhajime)) <> "" Then
                                        itmX.SubItems(1) = CStr(chdata(0, hhhajime))
                                    End If
                                    
                                    If Trim(chdata(1, hhhajime)) <> "" Then
                                        itmX.SubItems(2) = CStr(chdata(1, hhhajime))
                                    End If
                                    If Trim(chdata(2, hhhajime)) <> "" Then
                                        itmX.SubItems(3) = CStr(chdata(2, hhhajime))
                                    End If
                                    If IsNull(chdata(3, hhhajime)) Or Trim(chdata(3, hhhajime)) <> "" Then
                                        itmX.SubItems(4) = CStr(chdata(3, hhhajime))
                                    End If
                                    If IsNull(chdata(6, hhhajime)) Or Trim(chdata(6, hhhajime)) <> "" Then
                                        itmX.SubItems(5) = CStr(chdata(6, hhhajime))
                                    End If
                            End If
                       End If
                        
                 Next
                Screen.MousePointer = 0
                .Picture15.Enabled = True
                .lblkekka.Caption = "éÂå_ñÒì¸ã‡çáåvã‡äzÅ@Å@Å@Å@ÉCÉìÉ^Å[ÉlÉbÉgånÅF" & kin_inter & "â~ÅAÇbÇ`ÇsÇuånÅF" & kin_catv & "â~" & Chr(13) & _
                "ÇªÇÃëºå_ñÒì¸ã‡çáåvã‡äzÅ@Å@ÉCÉìÉ^Å[ÉlÉbÉgånÅF" & kin_inter2 & "â~ÅAÇbÇ`ÇsÇuånÅF" & kin_catv2 & "â~" & Chr(13) & _
                "ÇmÇgÇjå_ñÒì¸ã‡çáåvã‡äzÅF" & souseikyuukingaku & "â~  ì¸ã‡ëçäzÅF" & sounyuukingaku & "â~"
                
                
                
           
            
End With
sentakusumidesuka = 1

        Unload FRMOMACHI
            Set FRMOMACHI = Nothing
           
            Screen.MousePointer = 0
            
             Me.Show 1

End Sub

Private Sub Picture13_Click()
Unload Me

End Sub

Private Sub Picture15_Click()

    Dim udtBrowseInfo As BROWSEINFO
    Dim lngIDList     As Long
    Dim strPath       As String
    Dim shiteppp As String
On Error GoTo errcsv
    '[ÉtÉHÉãÉ_ÇÃéQè∆]É_ÉCÉAÉçÉOèÓïÒÇê›íË
    With udtBrowseInfo
        .hWndOwner = Me.hwnd
        .pidlRoot = CSIDL_NETWORK
        .pszDisplayName = String(MAX_PATH2, vbNullChar)
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    '[ÉtÉHÉãÉ_ÇÃéQè∆]É_ÉCÉAÉçÉOÇï\é¶
    lngIDList = SHBrowseForFolder(udtBrowseInfo)
    If lngIDList <> 0 Then
        strPath = String(MAX_PATH, vbNullChar)
        'ITEMIDLISTÇÉpÉXÇ…ïœä∑
        SHGetPathFromIDList lngIDList, strPath
        If Left(strPath, 1) = vbNullChar Then
            'ÉpÉXÇ™ãÛÇÃèÍçáÇÕÉRÉìÉsÉÖÅ[É^ñºÇéÊìæ
            With udtBrowseInfo
                strPath = "\\" & Left(.pszDisplayName, _
                                      InStr(.pszDisplayName, _
                                            vbNullChar) - 1)
            End With
        Else
            strPath = Left(strPath, _
                           InStr(strPath, _
                                 vbNullChar) - 1)
        End If
        'åãâ ÇÉâÉxÉãÇ…ï\é¶
       shiteppp = strPath
        'ITEMIDLISTÇâï˙
        CoTaskMemFree lngIDList
    Else
        Exit Sub
    End If




    If shiteppp = "" Then
        MsgBox "ï€ë∂êÊÇéwíËÇµÇƒÇ©ÇÁé¿çsÇµÇƒÇ≠ÇæÇ≥Ç¢ÅB"
        Exit Sub
    End If
 '   shiteppp = shiteppp & "\ëÿî[é“_" & Format(Date, "yyyymmdd") & "-" & Format(Time, "hhmmss") & ".csv"



Dim csvsuu As Long, chfileout As String, flopy As String

chfileout = shiteppp & "\êøãÅì¸ã‡_" & Trim(combzennen.Text) & Trim(combzentsuki2.Text)
If cmbbankid.ListIndex <> -1 Then
    chfileout = chfileout & Trim(cmbbankid.Text)
End If
chfileout = chfileout & ".csv"
  
  Open chfileout For Output As #1 Len = 32000
  Write #1, "NO", "ID", "å_ñÒñº", "éÌï ", "éÌóﬁ", "åJâzåèêî", "åJâzäz", "êøãÅåèêî", "êøãÅäz", "ì¸ã‡åèêî", "ì¸ã‡äz", "ñ¢î[åèêî", "ñ¢î[äz", "èWåv(0:ïsóv1:ïKóv)", "ÉOÉãÅ[Év", "ï¿Ç—èá"

  For csvsuu = 1 To chusercount
      Write #1, CStr(Format(csvsuu, "000"));
      Write #1, CStr(chdata(0, csvsuu));
      Write #1, CStr(chdata(1, csvsuu));
      Write #1, CStr(chdata(2, csvsuu));
      Write #1, CStr(chdata(3, csvsuu));
      
      Write #1, CStr(chdata(10, csvsuu));
      
      Write #1, CLng(chdata(4, csvsuu));
      
      Write #1, CStr(chdata(11, csvsuu));
      
      Write #1, CLng(chdata(5, csvsuu));
      
      Write #1, CStr(chdata(12, csvsuu));
      
      Write #1, CLng(chdata(6, csvsuu));
      
      Write #1, CStr(chdata(13, csvsuu));
      
      Write #1, CLng(chdata(7, csvsuu));
      
       Write #1, CStr(chdata(16, csvsuu));
        Write #1, CStr(chdata(14, csvsuu));
         Write #1, CStr(chdata(15, csvsuu))
  Next
  Close #1
        
        Screen.MousePointer = 0
        MsgBox "çÏê¨ÇµÇ‹ÇµÇΩÅB"
Exit Sub
errcsv:
        MsgBox "èoóÕÇ…é∏îsÇµÇ‹ÇµÇΩÅB"
       Close #1
        Exit Sub
End Sub
