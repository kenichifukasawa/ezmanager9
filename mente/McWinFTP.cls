VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "McWinFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'********************************************************************
'*  McWinFTP    Version 0.92
'*                                      Copyright(C)1998-1999, Sugi.
'*-------------------------------------------------------------------
'*
'*  【 動作確認環境 】
'*      ・WinNT4S/SP3 + IE3.02 (サーバー)
'*      ・WinNT4W/SP3 + IE3.02 + VB5EE/SP3　(クライアント)
'*　　  ・Win98 + Plus!98 + OP + VB6EE/SP2　(クライアント)
'*          現在、環境が無いため上記の環境でしかテストを行っておりません。
'*
'*  【　使用条件　】
'*      このクラスモジュールはフリーウェアです｡
'*      著作権はSugi(sugisita@mitene.or.jp)が保持しますが､
'*      プロジェクトへの組み込み、配布等は自由にしていただいて構いません。
'*      ただし、書籍及びBBS等、不特定多数に配布される可能性のある場合は
'*      事前にメールをお願いします｡
'*      バグ､動作確認環境などについてもメールをいただけると幸いです｡
'*
'*  【　注意事項　】
'*      必ず FTPFILE.CLS を同じプロジェクトに組み込んでください。
'*
'*  【　修正履歴　】
'*      1999-03-26: Version 0.9
'*      1999-09-30: Version 0.91 ファイル日付が9時間ずれるバグを修正
'*      1999-12-15: Version 0.92 FtpGetFile使用時にキャッシュ情報をリロードする
'*                               ように変更(VersionUpのAKIさんより情報提供)
'*
'*  なお、最新モジュールの公開は「み～くんパパの仕事部屋」にて行っています。
'*      URL http://www.mitene.or.jp/~sugisita/family/papa/vb/
'*
'********************************************************************
' クラスモジュール情報
Private Const CLS_NAME = "McWinFTP Class Module"
Private Const CLS_VERNO = "0.9 b"
Private Const CLS_COPY = "Copyright(C)1999, Sugi."

'********************************************************************
'   WIN32 API 定義
'********************************************************************
Private Const MAX_PATH = 260

Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_OFFLINE = &H1000

Private Type FILETIME
    dwLowDateTime   As Long
    dwHighDateTime  As Long
End Type

Private Type SYSTEMTIME
    wYear           As Integer
    wMonth          As Integer
    wDayOfWeek      As Integer
    wDay            As Integer
    wHour           As Integer
    wMinute         As Integer
    wSecond         As Integer
    wMilliseconds   As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" ( _
        lpFileTime As FILETIME, _
        lpLocalFileTime As FILETIME) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32" ( _
        lpFileTime As FILETIME, _
        lpSystemTime As SYSTEMTIME) As Long

Private Declare Function FormatMessage Lib "kernel32" _
    Alias "FormatMessageA" ( _
        ByVal dwFlags As Long, _
        lpSource As Any, _
        ByVal dwMessageId As Long, _
        ByVal dwLanguageId As Long, _
        ByVal lpBuffer As String, _
        ByVal nSize As Long, _
        Arguments As Long) As Long

        Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
        Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
        Private Const FORMAT_MESSAGE_FROM_STRING = &H400
        Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
        Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
        Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
        Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
        
        Private Const LANG_NEUTRAL = &H0
        Private Const SUBLANG_DEFAULT = &H1

'********************************************************************
'   WININET.DLL
'
'********************************************************************

' 本クラスモジュールでは特に使用していませんが、以下のような制限があります。
Private Const INTERNET_MAX_HOST_NAME_LENGTH = 256
Private Const INTERNET_MAX_USER_NAME_LENGTH = 128
Private Const INTERNET_MAX_PASSWORD_LENGTH = 128

' デフォルトのFTPポート番号 (最大値は65535)
Private Const INTERNET_DEFAULT_FTP_PORT = 21

Private Const ERROR_NO_MORE_FILES = 18
Private Const INTERNET_ERROR_BASE = 12000

Private Const INTERNET_FLAG_RELOAD As Long = &H80000000     ' ADD:1999-12-15


'--------------------------------------------------------------------
'   WININET関連のエラーコード定義(兼、一覧)
'   ※　後述しているようにエラー処理は必要最低限しか記述していないため
'       必要ならばここの定数を使用してください。
'       (エラーの内容に関してはSDK等のドキュメントに記載されています。)
'--------------------------------------------------------------------
Private Const ERROR_INTERNET_OUT_OF_HANDLES = (INTERNET_ERROR_BASE + 1)
Private Const ERROR_INTERNET_TIMEOUT = (INTERNET_ERROR_BASE + 2)
Private Const ERROR_INTERNET_EXTENDED_ERROR = (INTERNET_ERROR_BASE + 3)
Private Const ERROR_INTERNET_INTERNAL_ERROR = (INTERNET_ERROR_BASE + 4)
Private Const ERROR_INTERNET_INVALID_URL = (INTERNET_ERROR_BASE + 5)
Private Const ERROR_INTERNET_UNRECOGNIZED_SCHEME = (INTERNET_ERROR_BASE + 6)
Private Const ERROR_INTERNET_NAME_NOT_RESOLVED = (INTERNET_ERROR_BASE + 7)
Private Const ERROR_INTERNET_PROTOCOL_NOT_FOUND = (INTERNET_ERROR_BASE + 8)
Private Const ERROR_INTERNET_INVALID_OPTION = (INTERNET_ERROR_BASE + 9)
Private Const ERROR_INTERNET_BAD_OPTION_LENGTH = (INTERNET_ERROR_BASE + 10)
Private Const ERROR_INTERNET_OPTION_NOT_SETTABLE = (INTERNET_ERROR_BASE + 11)
Private Const ERROR_INTERNET_SHUTDOWN = (INTERNET_ERROR_BASE + 12)
Private Const ERROR_INTERNET_INCORRECT_USER_NAME = (INTERNET_ERROR_BASE + 13)
Private Const ERROR_INTERNET_INCORRECT_PASSWORD = (INTERNET_ERROR_BASE + 14)
Private Const ERROR_INTERNET_LOGIN_FAILURE = (INTERNET_ERROR_BASE + 15)
Private Const ERROR_INTERNET_INVALID_OPERATION = (INTERNET_ERROR_BASE + 16)
Private Const ERROR_INTERNET_OPERATION_CANCELLED = (INTERNET_ERROR_BASE + 17)
Private Const ERROR_INTERNET_INCORRECT_HANDLE_TYPE = (INTERNET_ERROR_BASE + 18)
Private Const ERROR_INTERNET_INCORRECT_HANDLE_STATE = (INTERNET_ERROR_BASE + 19)
Private Const ERROR_INTERNET_NOT_PROXY_REQUEST = (INTERNET_ERROR_BASE + 20)
Private Const ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = (INTERNET_ERROR_BASE + 21)
Private Const ERROR_INTERNET_BAD_REGISTRY_PARAMETER = (INTERNET_ERROR_BASE + 22)
Private Const ERROR_INTERNET_NO_DIRECT_ACCESS = (INTERNET_ERROR_BASE + 23)
Private Const ERROR_INTERNET_NO_CONTEXT = (INTERNET_ERROR_BASE + 24)
Private Const ERROR_INTERNET_NO_CALLBACK = (INTERNET_ERROR_BASE + 25)
Private Const ERROR_INTERNET_REQUEST_PENDING = (INTERNET_ERROR_BASE + 26)
Private Const ERROR_INTERNET_INCORRECT_FORMAT = (INTERNET_ERROR_BASE + 27)
Private Const ERROR_INTERNET_ITEM_NOT_FOUND = (INTERNET_ERROR_BASE + 28)
Private Const ERROR_INTERNET_CANNOT_CONNECT = (INTERNET_ERROR_BASE + 29)
Private Const ERROR_INTERNET_CONNECTION_ABORTED = (INTERNET_ERROR_BASE + 30)
Private Const ERROR_INTERNET_CONNECTION_RESET = (INTERNET_ERROR_BASE + 31)
Private Const ERROR_INTERNET_FORCE_RETRY = (INTERNET_ERROR_BASE + 32)
Private Const ERROR_INTERNET_INVALID_PROXY_REQUEST = (INTERNET_ERROR_BASE + 33)
Private Const ERROR_INTERNET_NEED_UI = (INTERNET_ERROR_BASE + 34)

Private Const ERROR_INTERNET_HANDLE_EXISTS = (INTERNET_ERROR_BASE + 36)
Private Const ERROR_INTERNET_SEC_CERT_DATE_INVALID = (INTERNET_ERROR_BASE + 37)
Private Const ERROR_INTERNET_SEC_CERT_CN_INVALID = (INTERNET_ERROR_BASE + 38)
Private Const ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR = (INTERNET_ERROR_BASE + 39)
Private Const ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR = (INTERNET_ERROR_BASE + 40)
Private Const ERROR_INTERNET_MIXED_SECURITY = (INTERNET_ERROR_BASE + 41)
Private Const ERROR_INTERNET_CHG_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 42)
Private Const ERROR_INTERNET_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 43)
Private Const ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = (INTERNET_ERROR_BASE + 44)
Private Const ERROR_INTERNET_INVALID_CA = (INTERNET_ERROR_BASE + 45)
Private Const ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP = (INTERNET_ERROR_BASE + 46)
Private Const ERROR_INTERNET_ASYNC_THREAD_FAILED = (INTERNET_ERROR_BASE + 47)
Private Const ERROR_INTERNET_REDIRECT_SCHEME_CHANGE = (INTERNET_ERROR_BASE + 48)
Private Const ERROR_INTERNET_DIALOG_PENDING = (INTERNET_ERROR_BASE + 49)
Private Const ERROR_INTERNET_RETRY_DIALOG = (INTERNET_ERROR_BASE + 50)
Private Const ERROR_INTERNET_HTTPS_HTTP_SUBMIT_REDIR = (INTERNET_ERROR_BASE + 52)
Private Const ERROR_INTERNET_INSERT_CDROM = (INTERNET_ERROR_BASE + 53)

'--------------------------------------------------------------------
'   API関数の定義
'--------------------------------------------------------------------
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" ( _
        ByVal lpszAgent As String, _
        ByVal dwAccessType As Long, _
        ByVal lpszProxy As String, _
        ByVal lpszProxyBypass As String, _
        ByVal dwFlags As Long) As Long

        ' Access Types for InternetOpen()
        Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0                   '// use registry configuration
        Private Const INTERNET_OPEN_TYPE_DIRECT = 1                      '// direct to net
        Private Const INTERNET_OPEN_TYPE_PROXY = 3                       '// via named proxy
        Private Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY = 4 '// prevent using java/script/INS

Private Declare Function InternetCloseHandle Lib "wininet.dll" ( _
        ByVal hInet As Long) As Long

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" ( _
        ByVal hInternet As Long, _
        ByVal lpszRemoteHostName As String, _
        ByVal nRemoteHostPort As Long, _
        ByVal lpszUserName As String, _
        ByVal lpszPassword As String, _
        ByVal dwService As Long, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Long

        ' dwService
        Private Const INTERNET_SERVICE_FTP = 1
        ' dwFlags
        Private Const INTERNET_FLAG_PASSIVE = &H8000000
        ' dwContext
        Private Const INTERNET_NO_CALLBACK = 0
        
Private Declare Function FtpCreateDirectory Lib "wininet.dll" _
    Alias "FtpCreateDirectoryA" ( _
        ByVal hConnect As Long, _
        lpszDirectory As Any) As Long

Private Declare Function FtpRemoveDirectory Lib "wininet.dll" _
    Alias "FtpRemoveDirectoryA" ( _
        ByVal hConnect As Long, _
        lpszDirectory As Any) As Long

Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" _
    Alias "FtpSetCurrentDirectoryA" ( _
        ByVal hConnect As Long, _
        ByVal lpszDirectory As String) As Long

Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" _
    Alias "FtpGetCurrentDirectoryA" ( _
        ByVal hConnect As Long, _
        lpszCurrentDirectory As Any, _
        lpdwCurrentDirectory As Any) As Long

Private Declare Function FtpFindFirstFile Lib "wininet.dll" _
    Alias "FtpFindFirstFileA" ( _
        ByVal hConnect As Long, _
        ByVal lpszSearchFile As String, _
        lpFindFileData As WIN32_FIND_DATA, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Long

Private Declare Function InternetFindNextFile Lib "wininet.dll" _
    Alias "InternetFindNextFileA" ( _
        ByVal hFind As Long, _
        lpvFindData As WIN32_FIND_DATA) As Long

Private Declare Function FtpGetFile Lib "wininet.dll" _
    Alias "FtpGetFileA" ( _
        ByVal hConnect As Long, _
        lpszRemoteFile As Any, _
        lpszNewFile As Any, _
        ByVal fFailIfExists As Long, _
        ByVal dwFlagsAndAttributes As Long, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Long

Private Declare Function FtpPutFile Lib "wininet.dll" _
    Alias "FtpPutFileA" ( _
        ByVal hFtpSession As Long, _
        lpszLocalFile As Any, _
        lpszNewRemoteFile As Any, _
        ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Long
        
        ' dwFlags
        Private Const INTERNET_FLAG_TRANSFER_ASCII = &H1
        Private Const INTERNET_FLAG_TRANSFER_BINARY = &H2

Private Declare Function FtpRenameFile Lib "wininet.dll" _
    Alias "FtpRenameFileA" ( _
        ByVal hConnect As Long, _
        lpszExisting As Any, _
        lpszNew As Any) As Long

Private Declare Function FtpDeleteFile Lib "wininet.dll" _
    Alias "FtpDeleteFileA" ( _
        ByVal hConnect As Long, _
        lpszFileName As Any) As Long

Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" _
    Alias "InternetGetLastResponseInfoA" ( _
        lpdwError As Long, _
        lpszBuffer As Any, _
        lpdwBufferLength As Long) As Long

'--------------------------------------------------------------------
'   列挙型
'--------------------------------------------------------------------
' ファイル転送方式
Public Enum enmTrans
    ftpBinary           ' バイナリ
    ftpAscii            ' アスキー
End Enum

'--------------------------------------------------------------------
'   イベント
'--------------------------------------------------------------------
' エラーイベント
Public Event FTPFrr(ByVal ErrNumber As Long, ByVal ErrString As String)

'--------------------------------------------------------------------
'   パブリック変数
'--------------------------------------------------------------------
Public RemoteFile       As New FTPFile

'--------------------------------------------------------------------
'   ローカル変数
'--------------------------------------------------------------------
' プロパティ用
Private m_hInternet     As Long     ' インターネットハンドル
Private m_hFTP          As Long     ' FTPセッションハンドル
Private m_RemotePort    As Long     ' ポート
Private m_RemoteHost    As String   ' サーバー名
Private m_User          As String   ' ユーザー
Private m_Password      As String   ' パスワード
Private m_Passive       As Long     ' パッシブモード

' その他
Private hFind           As Long

Public Property Get ClassName() As String
    ClassName = CLS_NAME
End Property

Public Property Get Version() As String
    Version = CLS_VERNO
End Property

Public Property Get Copyright() As String
    Copyright = CLS_COPY
End Property

Public Property Get hInternet() As Long
    hInternet = m_hInternet
End Property

Public Property Get hFTP() As Long
    hFTP = m_hFTP
End Property

Public Property Get RemotePort() As Long
    RemotePort = m_RemotePort
End Property

Public Property Let RemotePort(nRemotePort As Long)
    m_RemotePort = nRemotePort
End Property

Public Property Get RemoteHost() As String
    RemoteHost = IIf((m_RemoteHost = vbNullString), "", m_RemoteHost)
End Property

Public Property Let RemoteHost(szRemoteHost As String)
    m_RemoteHost = IIf(szRemoteHost = "", vbNullString, szRemoteHost)
End Property

Public Property Get User() As String
    User = IIf((m_User = vbNullString), "", m_User)
End Property

Public Property Let User(szUser As String)
    m_User = IIf(szUser = "", vbNullString, szUser)
End Property

Public Property Get Password() As String
    Password = IIf((m_Password = vbNullString), "", m_Password)
End Property

Public Property Let Password(szPassword As String)
    m_Password = IIf(szPassword = "", vbNullString, szPassword)
End Property

Public Property Get PassiveFlag() As Boolean
    PassiveFlag = IIf((m_Passive = INTERNET_FLAG_PASSIVE), True, False)
End Property

Public Property Let PassiveFlag(bPassive As Boolean)
    m_Passive = IIf(bPassive, INTERNET_FLAG_PASSIVE, &H0)
End Property

'********************************************************************
'   カレントディレクトリの取得
'
'********************************************************************
Public Property Get CurDir() As String
    
    Dim bytDir(MAX_PATH)    As Byte
    Dim lngRet  As Long
    Dim lngLen  As Long
    Dim strBuf  As String
    
    lngLen = MAX_PATH
    lngRet = FtpGetCurrentDirectory(m_hFTP, bytDir(0), lngLen)
    If lngRet <> 0& Then
        strBuf = StrConv(bytDir(), vbUnicode)
        CurDir = Left(strBuf, InStr(1, strBuf, vbNullChar) - 1)
    Else
        Call RaiseFTPErrEvents(Err.LastDllError)
    End If
    
End Property

'********************************************************************
'   カレントディレクトリの指定
'
'********************************************************************
Public Property Let CurDir(szDir As String)
    
    Dim lngRet  As Long
    Dim strDir  As String
    
    strDir = szDir
    ' 1バイト目が"\"または"/"の場合は取り除く
    If Left(strDir, 1) = "\" Or Left(strDir, 1) = "/" Then
        strDir = Mid(strDir, 2)
    End If
    lngRet = FtpSetCurrentDirectory(m_hFTP, strDir)
    If lngRet = 0& Then
        Call RaiseFTPErrEvents(Err.LastDllError)
    End If

End Property

'********************************************************************
'   クラスの初期化処理
'       1.インターネットハンドルの取得
'       2.デフォルトのポート番号を設定
'********************************************************************
Private Sub Class_Initialize()
    ' 初期化処理を行い、インターネットハンドルを取得する
    m_hInternet = InternetOpen(vbNullString, INTERNET_OPEN_TYPE_PRECONFIG, _
                               vbNullString, vbNullString, 0&)
    If m_hInternet = 0& Then
        Call RaiseFTPErrEvents(Err.LastDllError)
    End If
    ' デフォルトのポートを設定する
    m_RemotePort = INTERNET_DEFAULT_FTP_PORT
End Sub

'********************************************************************
'   クラスの終了処理
'       1.FTPセッションハンドルの開放
'       2.インターネットハンドルの開放
'********************************************************************
Private Sub Class_Terminate()
    Set RemoteFile = Nothing
    Call CloseSession
    ' インターネットハンドルの開放
    If InternetCloseHandle(m_hInternet) = 0& Then
        Call RaiseFTPErrEvents(Err.LastDllError)
    End If
End Sub

'********************************************************************
'   FTPセッションを開く
'
'********************************************************************
Public Function OpenSession()
    m_hFTP = InternetConnect(m_hInternet, _
                             m_RemoteHost, m_RemotePort, _
                             m_User, m_Password, _
                             INTERNET_SERVICE_FTP, _
                             m_Passive, _
                             INTERNET_NO_CALLBACK)
    If m_hFTP = 0& Then
        Call RaiseFTPErrEvents(Err.LastDllError)
    End If
End Function

'********************************************************************
'   FTPセッションを閉じる
'
'********************************************************************
Public Function CloseSession()
    ' 要求ハンドルの開放
    Call CloseREQ
    If m_hFTP <> 0& Then
        If InternetCloseHandle(m_hFTP) = 0& Then
            Call RaiseFTPErrEvents(Err.LastDllError)
        Else
            m_hFTP = 0&
        End If
    End If
End Function

'********************************************************************
'   要求ハンドルの開放
'
'********************************************************************
Private Function CloseREQ()
    If hFind <> 0& Then
        If InternetCloseHandle(hFind) = 0& Then
            Call RaiseFTPErrEvents(Err.LastDllError)
        Else
            hFind = 0&
        End If
    End If
End Function

'********************************************************************
'   ファイルの転送  PUT(ローカル --> ホスト)
'
'   (引数)
'   szLocalFile         -   転送するローカルファイル名(フルパス)
'   szRemoteFile        -   転送先で作成されるファイル名
'   nTransfer(省略可)   -   ftpBinary=バイナリモード
'                           ftpAscii=アスキーモード
'   (戻り値)
'       0以外の場合は正常終了
'********************************************************************
Public Function PutFile(szLocalFile As String, szRemoteFile As String, _
                        Optional nTransfer As enmTrans = ftpBinary) As Long
    Dim bytLocalFile()  As Byte
    Dim bytRemoteFile() As Byte
    Dim lngRet      As Long
    Dim lngTransfer As Long
    If nTransfer = ftpAscii Then
        lngTransfer = INTERNET_FLAG_TRANSFER_ASCII
    Else
        lngTransfer = INTERNET_FLAG_TRANSFER_BINARY
    End If
    bytLocalFile = StrConv(szLocalFile & vbNullChar, vbFromUnicode)
    bytRemoteFile = StrConv(szRemoteFile & vbNullChar, vbFromUnicode)
    lngRet = FtpPutFile(m_hFTP, bytLocalFile(0), bytRemoteFile(0), _
                        lngTransfer, INTERNET_NO_CALLBACK)
    If lngRet = 0& Then
        Call RaiseFTPErrEvents(Err.LastDllError)
    End If
    PutFile = lngRet
End Function

'********************************************************************
'   ファイルの転送  GET(ホスト --> ローカル)
'
'   (引数)
'   szRemoteFile        -   転送するリモートファイル
'   szLocalFile         -   転送先のローカルファイル名(フルパス)
'   bFailExists         -   True = ローカルファイルが存在する場合は
'                                  エラーとする。
'                           False = 常に上書きコピー
'   nTransfer(省略可)   -   ftpBinary=バイナリモード
'                           ftpAscii=アスキーモード
'   (戻り値)
'       0以外の場合は正常終了
'********************************************************************
Public Function GetFile(szRemoteFile As String, szLocalFile As String, _
                        bFailExists As Boolean, Optional nTransfer As enmTrans = ftpBinary) As Long
    Dim bytRemoteFile() As Byte
    Dim bytLocalFile()  As Byte
    Dim lngRet      As Long
    Dim lngTransfer As Long
    If nTransfer = ftpAscii Then
        lngTransfer = INTERNET_FLAG_TRANSFER_ASCII
    Else
        lngTransfer = INTERNET_FLAG_TRANSFER_BINARY
    End If
    bytRemoteFile = StrConv(szRemoteFile & vbNullChar, vbFromUnicode)
    bytLocalFile = StrConv(szLocalFile & vbNullChar, vbFromUnicode)

    lngRet = FtpGetFile(m_hFTP, bytRemoteFile(0), bytLocalFile(0), _
                        bFailExists * -1, 0, lngTransfer Or INTERNET_FLAG_RELOAD, INTERNET_NO_CALLBACK)     ' UPDATE:1999-12-15
                        
    If lngRet = 0& Then
        Call RaiseFTPErrEvents(Err.LastDllError)
    End If
    GetFile = lngRet
End Function

'********************************************************************
'   ファイルのリネーム
'
'   (引数)
'   szExistFile         -   現在のファイル名
'   szNewFile           -   変更後のファイル名
'
'   (戻り値)
'       0以外の場合は正常終了
'********************************************************************
Public Function RenameFile(szExistFile As String, szNewFile As String) As Long
    Dim bytExistFile()  As Byte
    Dim bytNewFile()    As Byte
    Dim lngRet      As Long
    bytExistFile = StrConv(szExistFile & vbNullChar, vbFromUnicode)
    bytNewFile = StrConv(szNewFile & vbNullChar, vbFromUnicode)
    lngRet = FtpRenameFile(m_hFTP, bytExistFile(0), bytNewFile(0))
    If lngRet = 0& Then
        Call RaiseFTPErrEvents(Err.LastDllError)
    End If
    RenameFile = lngRet
End Function

'********************************************************************
'   ファイルの削除
'
'   (引数)
'   szRemoteFile        -   削除するサーバー側のファイル名
'
'   (戻り値)
'       0以外の場合は正常終了
'********************************************************************
Public Function DeleteFile(szRemoteFile As String) As Long
    Dim bytRemoteFile() As Byte
    Dim lngRet      As Long
    bytRemoteFile = StrConv(szRemoteFile & vbNullChar, vbFromUnicode)
    lngRet = FtpDeleteFile(m_hFTP, bytRemoteFile(0))
    If lngRet = 0& Then
        Call RaiseFTPErrEvents(Err.LastDllError)
    End If
    DeleteFile = lngRet
End Function

'********************************************************************
'   ディレクトリの作成
'
'   (引数)
'   szRemoteDir         -   作成するディレクトリ名
'
'   (戻り値)
'       0以外の場合は正常終了
'********************************************************************
Public Function CreateDir(szRemoteDir As String) As Long
    Dim bytRemoteDir() As Byte
    Dim lngRet      As Long
    bytRemoteDir = StrConv(szRemoteDir & vbNullChar, vbFromUnicode)
    lngRet = FtpCreateDirectory(m_hFTP, bytRemoteDir(0))
    If lngRet = 0& Then
        Call RaiseFTPErrEvents(Err.LastDllError)
    End If
    CreateDir = lngRet
End Function

'********************************************************************
'   ディレクトリの削除
'
'   (引数)
'   szRemoteDir         -   削除するディレクトリ名
'
'   (戻り値)
'       0以外の場合は正常終了
'********************************************************************
Public Function RemoveDir(szRemoteDir As String) As Long
    Dim bytRemoteDir() As Byte
    Dim lngRet      As Long
    bytRemoteDir = StrConv(szRemoteDir & vbNullChar, vbFromUnicode)
    lngRet = FtpRemoveDirectory(m_hFTP, bytRemoteDir(0))
    If lngRet = 0& Then
        Call RaiseFTPErrEvents(Err.LastDllError)
    End If
    RemoveDir = lngRet
End Function
    
'********************************************************************
'   ファイルの検索
'
'   (引数)
'   szSearchString  -   ファイル検索文字列
'
'********************************************************************
Public Sub FindFile(szSearchString As String)
    
    Dim FileInfo    As WIN32_FIND_DATA
    
    Call CloseREQ
    Set RemoteFile = Nothing
    Set RemoteFile = New FTPFile
    RemoteFile.SearchEnd = False
    With FileInfo
    .cFileName = String(MAX_PATH, vbNullChar)
    hFind = FtpFindFirstFile(m_hFTP, szSearchString, FileInfo, 0&, 0&)
    If hFind = 0& Then
        Call RaiseFTPErrEvents(Err.LastDllError)
        RemoteFile.SearchEnd = True
    Else
        ' ファイル情報をセット
        Call SetFileInfo(FileInfo)
    End If
    End With
    
End Sub

'********************************************************************
'   次のファイルを検索
'
'********************************************************************
Public Sub NextFile()
    
    Dim FileInfo    As WIN32_FIND_DATA
    Dim lngRet      As Long
    
    If RemoteFile.SearchEnd Then
        Exit Sub
    End If
    With FileInfo
    .cFileName = String(MAX_PATH, vbNullChar)
    lngRet = InternetFindNextFile(hFind, FileInfo)
    If lngRet = 0 Then
        Call RaiseFTPErrEvents(Err.LastDllError)
    Else
        ' ファイル情報をセット
        Call SetFileInfo(FileInfo)
    End If
    End With
    
End Sub

'********************************************************************
'   ファイル情報の格納
'
'   (引数)
'   FileInfo    -   FTPFindFirstFile 及び InternetFindNextFile で
'                   取得した検索ファイル情報構造体
'********************************************************************
Private Sub SetFileInfo(FileInfo As WIN32_FIND_DATA)
    
    Const DWORDMAX  As Double = 4294967297#
    
    Dim pFatDate    As Long
    Dim pFatTime    As Long
    Dim LocalTime   As FILETIME
    Dim SysTime     As SYSTEMTIME
    
    RemoteFile.SearchEnd = False
    
    ' リモートファイル名を格納
    RemoteFile.FileName = Left(FileInfo.cFileName, InStr(1, FileInfo.cFileName, vbNullChar) - 1)
    
    ' リモートファイルの日付を格納
        
    '<<<<<< 99.09.30 バグ修正 >>>>>>>>
    ' (ファイルの時間は世界標準時よりローカル日付に変換)
    'FileTimeToLocalFileTime FileInfo.ftLastWriteTime, LocalTime
    'FileTimeToSystemTime LocalTime, SysTime
    
    FileTimeToSystemTime FileInfo.ftLastWriteTime, SysTime
    RemoteFile.FileDate = CDate(SysTime.wYear & "/" & SysTime.wMonth & "/" & SysTime.wDay & " " & _
                          SysTime.wHour & ":" & SysTime.wMinute & ":" & SysTime.wSecond)
    ' リモートファイルのサイズを格納
    RemoteFile.FileSize = FileInfo.nFileSizeLow
    If FileInfo.nFileSizeHigh > 0 Then
        RemoteFile.FileSize = RemoteFile.FileSize + FileInfo.nFileSizeHigh * DWORDMAX
    End If
    ' リモートファイルの属性を格納
    Select Case FileInfo.dwFileAttributes
    Case FILE_ATTRIBUTE_NORMAL
        RemoteFile.FileType = vbNormal
    Case FILE_ATTRIBUTE_DIRECTORY
        RemoteFile.FileType = vbDirectory
    Case Else
        RemoteFile.FileType = vbNormal
    End Select

End Sub

'********************************************************************
'   FTPエラーイベントの起動
'
'   (引数)
'   FTPErrNumber    -   FTPエラーコード
'********************************************************************
Private Sub RaiseFTPErrEvents(ByVal FTPErrNumber As Long)
    
    Dim strErrMsg   As String
    
    Dim lngExtendErr    As Long
    Dim bytExtendErr(255)   As Byte
    Dim strExtendErr    As String
    Dim lngBufLen       As Long
    
    '======================================================
    ' 各エラーコード毎にメッセージ取得
    '======================================================
    Select Case FTPErrNumber
    '======================================================
    ' ファイルが存在しない場合
    '------------------------------------------------------
    '   RemoteFileオブジェクトのSearchEndプロパティをTrueに
    '   設定し、エラーイベントとしては起動せずに抜ける
    '======================================================
    Case ERROR_NO_MORE_FILES
        RemoteFile.SearchEnd = True
        Exit Sub
    '======================================================
    ' サーバーが拡張エラーを返した場合
    '------------------------------------------------------
    '   InternetGetLastResponseInfo 関数で別途エラー情報を
    '   取得する。
    '======================================================
    Case ERROR_INTERNET_EXTENDED_ERROR
        lngExtendErr = FTPErrNumber
        lngBufLen = 255
        InternetGetLastResponseInfo FTPErrNumber, bytExtendErr(0), lngBufLen
        strErrMsg = StrConv(bytExtendErr(), vbUnicode)
        strErrMsg = Left(strErrMsg, InStr(1, strErrMsg, vbNullChar) - 1)
    
    '======================================================
    ' WININET関連のエラー処理
    '------------------------------------------------------
    ' 現在は必要最低限のコードしか処理していません。
    ' これ以外にも必要な場合は別途記述して下さい。
    ' CONST定義はこのモジュール内に全て記述してあります。
    '======================================================
    Case ERROR_INTERNET_NAME_NOT_RESOLVED
        strErrMsg = "サーバー名を解決できませんでした。"
        
    '======================================================
    ' 通常のAPIエラー処理
    '======================================================
    Case Else
        strErrMsg = GetErrMsg(Err.LastDllError)
    End Select
        
    '======================================================
    ' WININET関連で処理していないエラーメッセージが発生した
    ' 場合はとりあえずエラーメッセージを出力。
    '======================================================
    If strErrMsg = "" Then
        strErrMsg = "エラーが発生しています。"
    End If
    
    ' クラスのエラーイベント(FTPErrイベント)を起動
    RaiseEvent FTPFrr(FTPErrNumber, strErrMsg)
    
End Sub

'********************************************************************
'   DLLのエラーメッセージを取得
'
'   (引数)
'   ErrNumber       -   GetLastError API または
'                       ErrオブジェクトのLastDllErrorで取得した
'                       DLL呼び出し側のスレッドが持つエラーコード
'********************************************************************
Private Function GetErrMsg(ByVal ErrNumber As Long) As String
    Dim lngRet      As Long
    Dim strMsgBuf   As String * 256
    lngRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                           FORMAT_MESSAGE_IGNORE_INSERTS Or _
                           FORMAT_MESSAGE_MAX_WIDTH_MASK, _
                           ByVal 0&, _
                           ErrNumber, _
                           MakeLangID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
                           ByVal strMsgBuf, _
                           255&, _
                           ByVal 0&)
                           ' ※ MakeLangIDではデフォルトの言語を取得
    GetErrMsg = Left(strMsgBuf, lngRet)
    
End Function

'********************************************************************
'   LangIDの作成
'   (引数)
'   usPrimaryLanguage   -   プライマリ言語ID
'   usSubLanguage       -   セカンダリ言語ID
'                           (ID値の詳細はSDKにて確認の事)
'********************************************************************
Private Function MakeLangID(ByVal usPrimaryLanguage As Integer, _
                            ByVal usSubLanguage As Long) As Long
    On Error GoTo Err_MakeLangID
    MakeLangID = usSubLanguage * 2 ^ 10 + usPrimaryLanguage
    Exit Function
Err_MakeLangID:
    MsgBox Err.Number & vbCrLf, vbCritical, App.Title
End Function


