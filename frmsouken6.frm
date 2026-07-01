VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7D.ocx"
Begin VB.Form frmsouken6 
   Caption         =   "取り消し・損金処理"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   8400
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   7560
      MouseIcon       =   "frmsouken6.frx":0000
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmsouken6.frx":08FF
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      ToolTipText     =   "入金処理を行います。"
      Top             =   2100
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "対処方法"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   6975
      Begin VB.TextBox txtnichiji2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   420
         TabIndex        =   18
         Top             =   1380
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "...."
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   1380
         Width           =   375
      End
      Begin VB.TextBox txtnichiji 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   420
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdhiduke 
         Caption         =   "...."
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtnyuukin 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   2  'ｵﾌ
         Left            =   4500
         MaxLength       =   7
         TabIndex        =   12
         ToolTipText     =   "新金融機関コードを入力してください。"
         Top             =   1380
         Width           =   1695
      End
      Begin VB.TextBox txtsonkin 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   2  'ｵﾌ
         Left            =   4500
         MaxLength       =   7
         TabIndex        =   6
         ToolTipText     =   "新金融機関コードを入力してください。"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txttorikeshikin 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   2  'ｵﾌ
         Left            =   4500
         MaxLength       =   7
         TabIndex        =   4
         ToolTipText     =   "新金融機関コードを入力してください。"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "入金"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "円"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6420
         TabIndex        =   13
         Top             =   1500
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "円"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6420
         TabIndex        =   9
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "円"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6420
         TabIndex        =   8
         Top             =   420
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "損金"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "請求取消金額"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
   End
   Begin VSFlex7DAOCtl.VSFlexGrid gridnyuukinshousai 
      Height          =   1755
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8175
      _cx             =   14420
      _cy             =   3096
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   7560
      MouseIcon       =   "frmsouken6.frx":11C9
      MousePointer    =   99  'ﾕｰｻﾞｰ定義
      Picture         =   "frmsouken6.frx":1AC8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   3060
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "処理実行"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   12
      Left            =   7320
      TabIndex        =   11
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "終了"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      Left            =   7590
      TabIndex        =   1
      Top             =   3600
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '不透明
      BorderStyle     =   0  '透明
      Height          =   1935
      Left            =   7260
      Top             =   1980
      Width           =   1080
   End
End
Attribute VB_Name = "frmsouken6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdhiduke_Click()
donohi = 6
frmhiduke.Show 1

End Sub

Private Sub Command1_Click()
donohi = 7
frmhiduke.Show 1

End Sub

Private Sub Picture13_Click()
Unload Me
End Sub

Private Sub Picture2_Click()

Dim newtorikeshikin As Long, chktorikeshikin As String
Dim newsonkin As Long, chksonkin As String
Dim newkeikin As Long, chkkeikin As String
Dim newkeikin2 As Long, chkkeikin2 As String
Dim koushnid As String, newmaebi As String
Dim koid As String, newhouhou As String, newhouhou2 As String
Dim newkoushinid As String, oldkeiid As String
Dim newkeiid As String
Dim newminyuukin As Long, chkminyuukin As String
Dim newsoukei As Long, chksoukei As String
Dim newminyuukin2 As Long
Dim newsoukei2 As Long, resssres, resresres3, newkajoukoushin As String
Dim chknyuukin As String, newnyuukin As Long
Dim torikeshibi As String, sonkinbi As String, nyuukinbi As String

resssres = MsgBox("取消金・損金の処理を実行してよろしいですか？", vbYesNo)
If resssres = vbYes Then
    resssres = MsgBox("金額の確認はまだですよね？", vbYesNo)
    If resssres = vbNo Then
        resssres = MsgBox("でも再確認はまだですよね？", vbYesNo)
        If resssres = vbYes Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
Else
    Exit Sub
End If

    newkajoukoushin = ""
    koid = frmtouroku2.lblid.Caption
    newmaebi = Format(Date, "yyyymmdd")
'伝票選択チェック
    frmsouken6.gridnyuukinshousai.Col = 0
    chkkeikin2 = gridnyuukinshousai.Text
    On Error GoTo errsuuji2
        newkeikin2 = CLng(chkkeikin2)
    On Error GoTo 0
'5行目チェック
    frmsouken6.gridnyuukinshousai.Col = 5
    oldkeiid = gridnyuukinshousai.Text
    Select Case oldkeiid
    Case "1"
        newhouhou = "3"
        newhouhou2 = "4"
    Case "2"
        newhouhou = "5"
        newhouhou2 = "6"
    Case "3"
        newhouhou = "7"
        newhouhou2 = "8"
    End Select
    frmsouken6.gridnyuukinshousai.Col = 6
    newkeiid = gridnyuukinshousai.Text
'選択伝票金額チェック
    frmsouken6.gridnyuukinshousai.Col = 4
    chkkeikin = gridnyuukinshousai.Text
    On Error GoTo errsuuji2
        newkeikin = CLng(chkkeikin)
    On Error GoTo 0
    If newkeikin = 6000 Then
        newkajoukoushin = "0"
    Else
        If frmtouroku2.txtjikaiseikyuubi.Text <> "" Then
            resresres3 = MsgBox("次回自動更新日時が設定されています。初期化しますか？", vbYesNo)
            If resresres3 = vbYes Then
                newkajoukoushin = "0"
            End If
        End If
    End If
'取消金額チェック
    chktorikeshikin = Trim(txttorikeshikin.Text)
    If chktorikeshikin = "" Then
        chktorikeshikin = "0"
    End If
    On Error GoTo errsuuji
        newtorikeshikin = CLng(chktorikeshikin)
    On Error GoTo 0
    If newtorikeshikin < 0 Then
        MsgBox "正の整数を入力してください。"
        Exit Sub
    End If
    If newtorikeshikin <> 0 Then
        torikeshibi = Trim(txtnichiji.Text)
        If torikeshibi = "" Then
            MsgBox "取消日を選択してください。"
            Exit Sub
        Else
            torikeshibi = Format(torikeshibi, "yyyymmdd")
        End If
    End If
'損金金額チェック
    chksonkin = Trim(txtsonkin.Text)
    If chksonkin = "" Then
        chksonkin = "0"
    End If
    On Error GoTo errsuuji
        newsonkin = CLng(chksonkin)
    On Error GoTo 0
    If newsonkin < 0 Then
        MsgBox "正の整数を入力してください。"
        Exit Sub
    End If
    If newsonkin <> 0 Then
        sonkinbi = Trim(txtnichiji.Text)
        If sonkinbi = "" Then
            MsgBox "損金日を選択してください。"
            Exit Sub
        Else
            sonkinbi = Format(sonkinbi, "yyyymmdd")
        End If
    End If
'入金金額チェック
    chknyuukin = Trim(txtnyuukin.Text)
    If chknyuukin = "" Then
        chknyuukin = "0"
    End If
    On Error GoTo errsuuji
        newnyuukin = CLng(chknyuukin)
    On Error GoTo 0
    If newnyuukin < 0 Then
        MsgBox "正の整数を入力してください。"
        Exit Sub
    End If
    If newnyuukin <> 0 Then
        nyuukinbi = Trim(txtnichiji2.Text)
        If nyuukinbi = "" Then
            MsgBox "入金日を選択してください。"
            Exit Sub
        Else
            nyuukinbi = Format(nyuukinbi, "yyyymmdd")
        End If
    End If

'未入金と総計
    chkminyuukin = Trim(frmtouroku2.txtminyuuruikei.Text)
    chksoukei = Trim(frmtouroku2.txtsoukei.Text)
    On Error GoTo errsuuji
        newminyuukin = CLng(chkminyuukin)
        newsoukei = CLng(chksoukei)
    On Error GoTo 0
    
'金額の互換チェック
    If newkeikin <> newtorikeshikin + newsonkin + newnyuukin Then
        MsgBox "伝票の金額と入力した取消金額と損金と入金の合計が合いません。"
        Exit Sub
    End If
    newminyuukin2 = newminyuukin - newkeikin
    newsoukei2 = newsoukei - newkeikin
    
'伝票の有無の確認
    '損金と取り消し金がないとき（入金のみ）
    If chktorikeshikin = "0" And chksonkin = "0" Then
        MsgBox "損金と取り消し金が設定されていません。入金のみの場合は、「入金」－「特殊入金」を使用してください。"
        Exit Sub
    End If

'shousaiのデータ変更
    If kidou2 = False Then
        MsgBox "データベースファイルを開けませんでした。"
        Exit Sub
    End If

BeginTrans

    'sql = "UPDATE shousai SET  keinyu = '" & newmaebi & "',keist='2',bankid='" & Space(1) & "'" & _
                " WHERE shousaiid = '" & chkkeikin2 & "'"
    If torikeshibi = "" Then
        sql = "UPDATE shousai SET  keinyu = '" & sonkinbi & "',keist='2',bankid='" & Space(1) & "'" & _
                    " WHERE shousaiid = '" & chkkeikin2 & "'"
    Else
        sql = "UPDATE shousai SET  keinyu = '" & torikeshibi & "',keist='2',bankid='" & Space(1) & "'" & _
                    " WHERE shousaiid = '" & chkkeikin2 & "'"
    End If
    On Error GoTo errsoukatsunn
         db2.Execute sql, dbSQLPassThrough
    On Error GoTo 0
'kojinのデータ変更
    If newkajoukoushin = "0" Then
        sql = "UPDATE kojin SET  minyuukin = '" & CStr(newminyuukin2) & "',soukei='" & CStr(newsoukei2) & "'" & _
                ",kajou='" & CStr(newkajoukoushin) & "'" & _
                    " WHERE kojinid = '" & koid & "'"
    Else
        sql = "UPDATE kojin SET  minyuukin = '" & CStr(newminyuukin2) & "',soukei='" & CStr(newsoukei2) & "'" & _
                    " WHERE kojinid = '" & koid & "'"
    End If
    On Error GoTo errsoukatsunn
         db.Execute sql, dbSQLPassThrough
    On Error GoTo 0
    frmtouroku2.txtminyuuruikei.Text = CStr(newminyuukin2)
    frmtouroku2.txtsoukei.Text = CStr(newsoukei2)
    
'koushinkirokuの追加
        If newnyuukin <> 0 Then
            If soukatsu(0, "renban", 8, "", 0, 0) = False Then
                    MsgBox "設定の読み込みに失敗しました。"
                    Exit Sub
            End If
            If reg = "0" Then
                reg = ""
            End If
            If reg = vbNullString Then
                koushnid = "00000001"
            ElseIf reg = "9999999" Then
                MsgBox "これ以上の更新記録の作成はできません。管理者に連絡してください。"
                Exit Sub
            Else
                koushnid = Format$(reg, "000000#")
            End If
           ' sql = "INSERT INTO koushinkiroku VALUES('" & koushnid & "'," & _
                           "'" & newmaebi & "','" & koid & "'," & _
                           "'" & chknyuukin & "','1','" & Space(1) & "')"
        
            sql = "INSERT INTO koushinkiroku VALUES('" & koushnid & "'," & _
                           "'" & nyuukinbi & "','" & koid & "'," & _
                           "'" & chknyuukin & "','1','" & Space(1) & "')"
            On Error GoTo errko2
                db2.Execute sql, dbSQLPassThrough
            On Error GoTo 0
            
            newkoushinid = CStr(CLng(koushnid) + 1)
            If soukatsu(1, "renban", 8, newkoushinid, 0, 0) = False Then
                    MsgBox "設定の書き込みに失敗しました。"
                    Exit Sub
            End If
            
            sql = "update kojin set  maebi ='" & nyuukinbi & "'," & _
                                       "maekin = '" & CStr(chknyuukin) & "'," & _
                                       "fkaisuu = '0'" & _
                                       " where kojinid ='" & koid & "'"
            On Error GoTo errko2
                db.Execute sql, dbSQLPassThrough
            On Error GoTo 0
    End If

    If newtorikeshikin <> 0 Then
            If soukatsu(0, "renban", 8, "", 0, 0) = False Then
                    MsgBox "設定の読み込みに失敗しました。"
                    Exit Sub
            End If
            If reg = "0" Then
                reg = ""
            End If
            If reg = vbNullString Then
                koushnid = "00000001"
            ElseIf reg = "9999999" Then
                MsgBox "これ以上の更新記録の作成はできません。管理者に連絡してください。"
                Exit Sub
            Else
                koushnid = Format$(reg, "000000#")
            End If
           ' sql = "INSERT INTO koushinkiroku VALUES('" & koushnid & "'," & _
                           "'" & newmaebi & "','" & koid & "'," & _
                           "'" & chktorikeshikin & "','" & newhouhou & "','" & newkeiid & "')"
        
            sql = "INSERT INTO koushinkiroku VALUES('" & koushnid & "'," & _
                           "'" & torikeshibi & "','" & koid & "'," & _
                           "'" & chktorikeshikin & "','" & newhouhou & "','" & newkeiid & "')"
            On Error GoTo errko2
                db2.Execute sql, dbSQLPassThrough
            On Error GoTo 0
            
            newkoushinid = CStr(CLng(koushnid) + 1)
            If soukatsu(1, "renban", 8, newkoushinid, 0, 0) = False Then
                    MsgBox "設定の書き込みに失敗しました。"
                    Exit Sub
            End If
    End If
    If newsonkin <> 0 Then
            If soukatsu(0, "renban", 8, "", 0, 0) = False Then
                    MsgBox "設定の読み込みに失敗しました。"
                    Exit Sub
            End If
            If reg = "0" Then
                reg = ""
            End If
            If reg = vbNullString Then
                koushnid = "00000001"
            ElseIf reg = "9999999" Then
                MsgBox "これ以上の更新記録の作成はできません。管理者に連絡してください。"
                Exit Sub
            Else
                koushnid = Format$(reg, "000000#")
            End If
            'sql = "INSERT INTO koushinkiroku VALUES('" & koushnid & "'," & _
                           "'" & newmaebi & "','" & koid & "'," & _
                           "'" & chksonkin & "','" & newhouhou2 & "','" & newkeiid & "')"
        
            sql = "INSERT INTO koushinkiroku VALUES('" & koushnid & "'," & _
                           "'" & sonkinbi & "','" & koid & "'," & _
                           "'" & chksonkin & "','" & newhouhou2 & "','" & newkeiid & "')"
            On Error GoTo errko2
                db2.Execute sql, dbSQLPassThrough
            On Error GoTo 0
            
            newkoushinid = CStr(CLng(koushnid) + 1)
            If soukatsu(1, "renban", 8, newkoushinid, 0, 0) = False Then
                    MsgBox "設定の書き込みに失敗しました。"
                    Exit Sub
            End If
    End If


CommitTrans

MsgBox "処理が完了しました。"
Unload frmsouken6
If newkajoukoushin = "0" Then
    frmtouroku2.txtjikaiseikyuubi.Text = ""
End If
If newnyuukin <> 0 Then
            If hyouji_touroku(koid) = False Then
                Exit Sub
            End If
            Unload Me
End If

Exit Sub

errsuuji:

    MsgBox "半角数字を入力してください。"
    Exit Sub
errsuuji2:

    MsgBox "請求済み伝票を選択してください。"
    Exit Sub
errko2:
    MsgBox "更新記録のデータ更新に失敗しました。"
    Rollback
    Exit Sub
errsoukatsunn:
    MsgBox "請求記録のデータ更新に失敗しました。"
    Rollback
    Exit Sub

End Sub
