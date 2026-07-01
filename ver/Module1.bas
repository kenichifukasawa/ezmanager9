Attribute VB_Name = "Module1"
Option Explicit
'読み込みのためのAPIです。
Public Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
ByVal lpDefault As String, ByVal lpReturnedString As String, _
ByVal nSize As Long, ByVal lpFileName As String) As Long



Public strFileName(4) As String
Public strFileName2(4) As String
Public copysakifile(4) As String



Sub Main()

    Dim settei_res3 As String
    Dim arunokana As String, ret, ververst As Integer


Dim root_pass As String

Dim objwsh As Object
Set objwsh = CreateObject("Wscript.Shell")
root_pass = objwsh.SpecialFolders("MyDocuments")

If Right(root_pass, 1) <> "\" Then
   root_pass = root_pass & "\"
End If
Set objwsh = Nothing


frmverup.Show
DoEvents

With frmverup
    '管理パス取得


    
    Screen.MousePointer = 11
        
    '総合のパス設定
    
    'MyDoc
         arunokana = root_pass
    
    
    ' strFileName に取得したいファイル名をセット
    strFileName(0) = App.Path & "\catv.exe"
    strFileName(1) = App.Path & "\ezmanager_ment.exe"
     strFileName(3) = App.Path & "\ezmanager_sub.exe"
     
    strFileName2(0) = arunokana & "catv.exe"
    strFileName2(1) = arunokana & "ezmanager_ment.exe"
     strFileName2(3) = arunokana & "ezmanager_sub.exe"
     
    copysakifile(0) = App.Path & "\oldcatv.exe"
    copysakifile(1) = App.Path & "\oldezmanager_ment.exe"
    copysakifile(3) = App.Path & "\oldezmanager_sub.exe"
    ververst = 0
    Dim i As Integer
    
    'For i = 0 To 1
        i = 0
        If Dir(strFileName2(i)) <> "" Then
            'OLDがあれば削除
            If Dir(copysakifile(i)) <> "" Then
                Kill copysakifile(i)
                DoEvents
            End If
            '現行をOLDに変更
            If Dir(strFileName(i)) <> "" Then
                Name strFileName(i) As copysakifile(i)
                DoEvents
            End If
            '最新版をルートにコピー
            FileCopy strFileName2(i), strFileName(i)
            DoEvents
            ververst = ververst + 1
        Else
            MsgBox "バージョンアップ元のファイルがありません。ファイル名「" & strFileName2(i) & "」"
        End If
    'Next i
    DoEvents
    
            i = 3
        If Dir(strFileName2(i)) <> "" Then
            'OLDがあれば削除
            If Dir(copysakifile(i)) <> "" Then
                Kill copysakifile(i)
                DoEvents
            End If
            '現行をOLDに変更
            If Dir(strFileName(i)) <> "" Then
                Name strFileName(i) As copysakifile(i)
                DoEvents
            End If
            '最新版をルートにコピー
            FileCopy strFileName2(i), strFileName(i)
            DoEvents
            ververst = ververst + 1
        Else
            MsgBox "バージョンアップ元のファイルがありません。ファイル名「" & strFileName2(i) & "」"
        End If
    'Next i
    DoEvents
    
    
    
    Screen.MousePointer = 0
    If ververst <> 0 Then
        MsgBox "ファイル" & ververst & "個のバージョンアップが終了しました。"
    End If
    Dim RetVal, resetfaile As String

    RetVal = Shell(strFileName(0), 1)   ' 電卓を実行します。
   
    

    End

End With
End Sub

