VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FTPFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'********************************************************************
'*  McWinFTPクラスへの検索ファイル情報の組み込みメンバークラス
'*-------------------------------------------------------------------
'*  McWinFTPクラスを使用時には必ず本クラスモジュールも同じプロジェクト
'*  に組み込んでください。
'********************************************************************

'リモートファイル情報
Public FileName As String               ' ファイル名
Public FileSize As Double               ' ファイルサイズ(バイト)
Public FileDate As Date                 ' 日付
Public FileType As VBA.VbFileAttribute  ' ファイル属性

' ディレクトリ内のファイル検索が終了したフラグ
Public SearchEnd    As Boolean

