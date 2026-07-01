VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCryptFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'====================================================================================
' 定数宣言
'====================================================================================

'// ALG_ID crackers
'#define GET_ALG_CLASS(x)                (x & (7 << 13))
'#define GET_ALG_TYPE(x)                 (x & (15 << 9))
'#define GET_ALG_SID(x)                  (x & (511))

'// Algorithm classes
Private Const ALG_CLASS_ANY = 0                ' (0)
Private Const ALG_CLASS_SIGNATURE = 8192       ' (1 << 13)
Private Const ALG_CLASS_MSG_ENCRYPT = 16384    ' (2 << 13)
Private Const ALG_CLASS_DATA_ENCRYPT = 24576   ' (3 << 13)
Private Const ALG_CLASS_HASH = 32768           ' (4 << 13)
Private Const ALG_CLASS_KEY_EXCHANGE = 40960   ' (5 << 13)

'// Algorithm types
Private Const ALG_TYPE_ANY = 0                 ' (0)
Private Const ALG_TYPE_DSS = 512               ' (1 << 9)
Private Const ALG_TYPE_RSA = 1024              ' (2 << 9)
Private Const ALG_TYPE_BLOCK = 1536            ' (3 << 9)
Private Const ALG_TYPE_STREAM = 2048           ' (4 << 9)

'// Generic sub-ids
Private Const ALG_SID_ANY = 0                  ' (0)

'// Some RSA sub-ids
Private Const ALG_SID_RSA_ANY = 0
Private Const ALG_SID_RSA_PKCS = 1
Private Const ALG_SID_RSA_MSATWORK = 2
Private Const ALG_SID_RSA_ENTRUST = 3
Private Const ALG_SID_RSA_PGP = 4

'// Some DSS sub-ids
'//
Private Const ALG_SID_DSS_ANY = 0
Private Const ALG_SID_DSS_PKCS = 1
Private Const ALG_SID_DSS_DMS = 2

'// Block cipher sub ids
'// DES sub_ids
Private Const ALG_SID_DES = 1
Private Const ALG_SID_3DES = 3
Private Const ALG_SID_DESX = 4
Private Const ALG_SID_IDEA = 5
Private Const ALG_SID_CAST = 6
Private Const ALG_SID_SAFERSK64 = 7
Private Const ALD_SID_SAFERSK128 = 8

'// KP_MODE
Private Const CRYPT_MODE_CBCI = 6       '// ANSI CBC Interleaved
Private Const CRYPT_MODE_CFBP = 7       '// ANSI CFB Pipelined
Private Const CRYPT_MODE_OFBP = 8       '// ANSI OFB Pipelined
Private Const CRYPT_MODE_CBCOFM = 9     '// ANSI CBC + OF Masking
Private Const CRYPT_MODE_CBCOFMI = 10   '// ANSI CBC + OFM Interleaved

'// RC2 sub-ids
Private Const ALG_SID_RC2 = 2

'// Stream cipher sub-ids
Private Const ALG_SID_RC4 = 1
Private Const ALG_SID_SEAL = 2

'// Hash sub ids
Private Const ALG_SID_MD2 = 1
Private Const ALG_SID_MD4 = 2
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA = 4
Private Const ALG_SID_MAC = 5
Private Const ALG_SID_RIPEMD = 6
Private Const ALG_SID_RIPEMD160 = 7
Private Const ALG_SID_SSL3SHAMD5 = 8

'// Our silly example sub-id
Private Const ALG_SID_EXAMPLE = 80

'// algorithm identifier definitions
Private Const CALG_MD2 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2)
Private Const CALG_MD4 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4)
Private Const CALG_MD5 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5)
Private Const CALG_SHA = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA)
Private Const CALG_MAC = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MAC)
Private Const CALG_RSA_SIGN = (ALG_CLASS_SIGNATURE Or ALG_TYPE_RSA Or ALG_SID_RSA_ANY)
Private Const CALG_DSS_SIGN = (ALG_CLASS_SIGNATURE Or ALG_TYPE_DSS Or ALG_SID_DSS_ANY)
Private Const CALG_RSA_KEYX = (ALG_CLASS_KEY_EXCHANGE Or ALG_TYPE_RSA Or ALG_SID_RSA_ANY)
Private Const CALG_DES = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_DES)
Private Const CALG_RC2 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_RC2)
Private Const CALG_RC4 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4)
Private Const CALG_SEAL = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_SEAL)

'// dwFlags definitions for CryptAquireContext
Private Const CRYPT_VERIFYCONTEXT = &HF0000000
Private Const CRYPT_NEWKEYSET = &H8
Private Const CRYPT_DELETEKEYSET = &H10

'// dwFlag definitions for CryptGenKey
Private Const CRYPT_EXPORTABLE = &H1
Private Const CRYPT_USER_PROTECTED = &H2
Private Const CRYPT_CREATE_SALT = &H4
Private Const CRYPT_UPDATE_KEY = &H8

'// exported key blob definitions
Private Const SIMPLEBLOB = &H1
Private Const PUBLICKEYBLOB = &H6
Private Const PRIVATEKEYBLOB = &H7

Private Const AT_KEYEXCHANGE = 1
Private Const AT_SIGNATURE = 2

Private Const CRYPT_USERDATA = 1

'// dwParam
Private Const KP_IV = 1           '// Initialization vector
Private Const KP_SALT = 2         '// Salt value
Private Const KP_PADDING = 3      '// Padding values
Private Const KP_MODE = 4         '// Mode of the cipher
Private Const KP_MODE_BITS = 5    '// Number of bits to feedback
Private Const KP_PERMISSIONS = 6  '// Key permissions DWORD
Private Const KP_ALGID = 7        '// Key algorithm
Private Const KP_BLOCKLEN = 8     '// Block size of the cipher

'// KP_PADDING
Private Const PKCS5_PADDING = 1   '// PKCS 5 (sec 6.2) padding method

'// KP_MODE
Private Const CRYPT_MODE_CBC = 1  '// Cipher block chaining
Private Const CRYPT_MODE_ECB = 2  '// Electronic code book
Private Const CRYPT_MODE_OFB = 3  '// Output feedback mode
Private Const CRYPT_MODE_CFB = 4  '// Cipher feedback mode
Private Const CRYPT_MODE_CTS = 5  '// Ciphertext stealing mode

'// KP_PERMISSIONS
Private Const CRYPT_ENCRYPT = &H1 '// Allow encryption
Private Const CRYPT_DECRYPT = &H2 '// Allow decryption
Private Const CRYPT_EXPORT = &H4  '// Allow key to be exported
Private Const CRYPT_READ = &H8    '// Allow parameters to be read
Private Const CRYPT_WRITE = &H10  '// Allow parameters to be set
Private Const CRYPT_MAC = &H20    '// Allow MACs to be used with key

Private Const HP_ALGID = &H1      '// Hash algorithm
Private Const HP_HASHVAL = &H2    '// Hash value
Private Const HP_HASHSIZE = &H4   '// Hash value size

'#define CRYPT_FAILED            FALSE
'#define CRYPT_SUCCEED           TRUE

'#define RCRYPT_SUCCEEDED(rt)     ((rt) == CRYPT_SUCCEED)
'#define RCRYPT_FAILED(rt)        ((rt) == CRYPT_FAILED)

'//
'// CryptGetProvParam
'//
Private Const PP_ENUMALGS = 1
Private Const PP_ENUMCONTAINERS = 2
Private Const PP_IMPTYPE = 3
Private Const PP_NAME = 4
Private Const PP_VERSION = 5
Private Const PP_CONTAINER = 6

Private Const CRYPT_FIRST = 1
Private Const CRYPT_NEXT = 2

Private Const CRYPT_IMPL_HARDWARE = 1
Private Const CRYPT_IMPL_SOFTWARE = 2
Private Const CRYPT_IMPL_MIXED = 3
Private Const CRYPT_IMPL_UNKNOWN = 4

'//
'// CryptSetProvParam
'//
Private Const PP_CLIENT_HWND = 1

Private Const PROV_RSA_FULL = 1
Private Const PROV_RSA_SIG = 2
Private Const PROV_DSS = 3
Private Const PROV_FORTEZZA = 4
Private Const PROV_MS_EXCHANGE = 5
Private Const PROV_SSL = 6

'//
'//STT defined Providers
'//
Private Const PROV_STT_MER = 7
Private Const PROV_STT_ACQ = 8
Private Const PROV_STT_BRND = 9
Private Const PROV_STT_ROOT = 10
Private Const PROV_STT_ISS = 11

'====================================================================================
' API関数宣言
'====================================================================================

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const OPEN_EXISTING = 3
Private Const CREATE_ALWAYS = 2

'ファイル作成
Private Declare Function CreateFile Lib "KERNEL32" Alias "CreateFileA" ( _
                                        ByVal lpFileName As String, _
                                        ByVal dwDesiredAccess As Long, _
                                        ByVal dwShareMode As Long, _
                                        lpSecurityAttributes As Long, _
                                        ByVal dwCreationDisposition As Long, _
                                        ByVal dwFlagsAndAttributes As Long, _
                                        ByVal hTemplateFile As Long) As Long

'ファイルサイズ取得
Private Declare Function GetFileSize Lib "KERNEL32" ( _
                                        ByVal hFile As Long, _
                                        lpFileSizeHigh As Long) As Long

'ファイル読み込み
Private Declare Function ReadFile Lib "KERNEL32" ( _
                                        ByVal hFile As Long, _
                                        lpBuffer As Any, _
                                        ByVal nNumberOfBytesToRead As Long, _
                                        lpNumberOfBytesRead As Long, _
                                        lpOverlapped As Long) As Long

'ファイル書き込み
Private Declare Function WriteFile Lib "KERNEL32" ( _
                                        ByVal hFile As Long, _
                                        lpBuffer As Any, _
                                        ByVal nNumberOfBytesToWrite As Long, _
                                        lpNumberOfBytesWritten As Long, _
                                        lpOverlapped As Long) As Long
'ハンドルクローズ
Private Declare Function CloseHandle Lib "KERNEL32" ( _
                                        ByVal hObject As Long) As Long

'------------------------------------------------------------------------------------
' CSPハンドル取得
'------------------------------------------------------------------------------------
' [O] phProv       : 取得したCSPハンドル
' [I] pszContainer : コンテナ名(NULL=デフォルトコンテナ)
' [I] pszProvider  : プロバイダ名(NULL=デフォルトプロバイダ)
' [I] dwProvType   : プロバイダタイプ(CryptSetProvParam 定数)
' [I] dwFlags      : フラグ(dwFlags definitions for CryptAquireContext 参照)
'
'           戻り値 : 成功 = True (0以外)
'                  : 失敗 = False(0)
'
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
                                        phProv As Long, _
                                        ByVal pszContainer As String, _
                                        ByVal pszProvider As String, _
                                        ByVal dwProvType As Long, _
                                        ByVal dwFlags As Long) As Long

'------------------------------------------------------------------------------------
' ランダムキー生成
'------------------------------------------------------------------------------------
' [I] phProv  : CSPハンドル
' [I] Algid   : アルゴリズムID(algorithm identifier definitions 参照)
' [I] dwFlags : キー生成タイプフラグ(dwFlag definitions for CryptGenKey 参照)
' [O] phKey   : 生成したキーハンドル
'
'      戻り値 : 成功 = True (0以外)
'             : 失敗 = False(0)
'
Private Declare Function CryptGenKey Lib "advapi32.dll" ( _
                                        ByVal hProv As Long, _
                                        ByVal Algid As Long, _
                                        ByVal dwFlags As Long, _
                                        phKey As Long) As Long

'------------------------------------------------------------------------------------
' ユーザーキーペアハンドル取得
'------------------------------------------------------------------------------------
' [I] phProv    : CSPハンドル
' [I] dwKeySpec : 取得するキーの種類
'               : - AT_KEYEXCHANGE = Exchangeキー
'               : - AT_SIGNATURE   = 署名キー
' [O] phUserKey : 生成したユーザーキーハンドル
'
'        戻り値 : 成功 = True (0以外)
'               : 失敗 = False(0)
'
Private Declare Function CryptGetUserKey Lib "advapi32.dll" ( _
                                        ByVal hProv As Long, _
                                        ByVal dwKeySpec As Long, _
                                        phUserKey As Long) As Long

'------------------------------------------------------------------------------------
' CSPからキーをエクスポート
'------------------------------------------------------------------------------------
' [ I ] hKey       : エクスポートされるキーハンドル
' [ I ] hExpKey    : エクスポートキーハンドル
' [ I ] dwBlobType : エクスポートされるkey blobタイプ(exported key blob definitions 参照)
' [ I ] dwFlags    : 0で予約されています
' [I/O] pbData     : key blobを受け取るバッファポインタ(NULLの場合サイズ取得)
' [ O ] pdwDataLen : pbDataバッファサイズ
'
'           戻り値 : 成功 = True (0以外)
'                  : 失敗 = False(0)
'
Private Declare Function CryptExportKey Lib "advapi32.dll" ( _
                                        ByVal hKey As Long, _
                                        ByVal hExpKey As Long, _
                                        ByVal dwBlobType As Long, _
                                        ByVal dwFlags As Long, _
                                        pbData As Any, _
                                        pdwDataLen As Long) As Long

'------------------------------------------------------------------------------------
' キーハンドル開放
'------------------------------------------------------------------------------------
' [I] hKey : 開放するキーハンドル
'
'   戻り値 : 成功 = True (0以外)
'          : 失敗 = False(0)
'
Private Declare Function CryptDestroyKey Lib "advapi32.dll" ( _
                                        ByVal hKey As Long) As Long

'------------------------------------------------------------------------------------
' ハッシュ・オブジェク作成
'------------------------------------------------------------------------------------
' [I] hProv   : CSPハンドル
' [I] Algid   : ハッシュ・アルゴリズムID(algorithm identifier definitions 参照)
' [I] hKey    : ハッシュ・アルゴリズムがHMACとかMACアルゴリズムのように組み合わせハッ
'             : シュの場合指定します。鍵なしアルゴリズムの場合は 0
' [I] dwFlags : 0で予約されています
' [I] phHash  : 生成されたハッシュ・オブジェクトハンドル
'
'      戻り値 : 成功 = True (0以外)
'             : 失敗 = False(0)
'
Private Declare Function CryptCreateHash Lib "advapi32.dll" ( _
                                        ByVal hProv As Long, _
                                        ByVal Algid As Long, _
                                        ByVal hKey As Long, _
                                        ByVal dwFlags As Long, _
                                        phHash As Long) As Long

'------------------------------------------------------------------------------------
' ハッシュ・オブジェクトへデータ供給
'------------------------------------------------------------------------------------
' [I] hHash   : ハッシュ・オブジェクトハンドル
' [I] pbData  : ハッシュされるデータ
' [I] cbData  : ハッシュされるデータサイズ
' [I] dwFlags : The Microsoft Base Cryptographic Provider では無視されます
'
'      戻り値 : 成功 = True (0以外)
'             : 失敗 = False(0)
'
Private Declare Function CryptHashData Lib "advapi32.dll" ( _
                                        ByVal hHash As Long, _
                                        ByVal pbData As String, _
                                        ByVal cbData As Long, _
                                        ByVal dwFlags As Long) As Long

'------------------------------------------------------------------------------------
' 暗号キー生成
'------------------------------------------------------------------------------------
' [ I ] hProv     : CSPハンドル
' [ I ] Algid     : アルゴリズムID(algorithm identifier definitions 参照)
' [ I ] hBaseData : ハッシュ・オブジェクトハンドル
' [ I ] dwFlags   : 生成されるキータイプID
' [I/O] phKey     : 生成されたキーハンドル
'
'          戻り値 : 成功 = True (0以外)
'                 : 失敗 = False(0)
'
Private Declare Function CryptDeriveKey Lib "advapi32.dll" ( _
                                        ByVal hProv As Long, _
                                        ByVal Algid As Long, _
                                        ByVal hBaseData As Long, _
                                        ByVal dwFlags As Long, _
                                        phKey As Long) As Long

'------------------------------------------------------------------------------------
' 暗号化
'------------------------------------------------------------------------------------
' [ I ] hKey     : 暗号化キーハンドル
' [ I ] hHash    : ハッシュオブジェクトハンドル(ハッシュを利用しない場合は 0)
' [ I ] Final    : 最後のセクションであるかどうか(True/False)
' [ I ] dwFlags  : 0で予約されています
' [I/O] pbData   : 暗号化データ。暗号化が成功すると暗号されたデータが格納されます
' [I/O] pcbData  : 暗号化データサイズ。暗号化が成功すると暗号化されたデータサイズが格納されます
' [ I ] cbBuffer : pbDataバッファのバイト数
'
'           戻り値 : 成功 = True (0以外)
'                  : 失敗 = False(0)
'
Private Declare Function CryptEncrypt Lib "advapi32.dll" ( _
                                        ByVal hKey As Long, _
                                        ByVal hHash As Long, _
                                        ByVal Final As Long, _
                                        ByVal dwFlags As Long, _
                                        pbData As Any, _
                                        pcbData As Long, _
                                        ByVal cbBuffer As Long) As Long

'------------------------------------------------------------------------------------
' 複合化
'------------------------------------------------------------------------------------
' [ I ] hKey       : 複合化キーハンドル
' [ I ] hHash      : ハッシュオブジェクトハンドル(ハッシュを利用しない場合は 0)
' [ I ] Final      : 最後のセクションであるかどうか(True/False)
' [ I ] dwFlags    : 0で予約されています
' [I/O] pbData     : 複合化データ。複合化が成功すると複合されたデータが格納されます
' [I/O] pdwDataLen : 複合化データサイズ。複合化が成功すると複合化されたデータサイズが格納されます
'
'           戻り値 : 成功 = True (0以外)
'                  : 失敗 = False(0)
'
Private Declare Function CryptDecrypt Lib "advapi32.dll" ( _
                                       ByVal hKey As Long, _
                                       ByVal hHash As Long, _
                                       ByVal Final As Long, _
                                       ByVal dwFlags As Long, _
                                       pbData As Any, _
                                       pdwDataLen As Long) As Long

'------------------------------------------------------------------------------------
' CSPやコンテナキーハンドル開放
'------------------------------------------------------------------------------------
' [I] phProv  : CSPやコンテナキーハンドル
' [I] dwFlags : 0で予約されています
'
'      戻り値 : 成功 = True (0以外)
'             : 失敗 = False(0)
'
Private Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
                                        ByVal hProv As Long, _
                                        ByVal dwFlags As Long) As Long

'------------------------------------------------------------------------------------
' ハッシュ・オブジェクト破棄
'------------------------------------------------------------------------------------
' [I] hHash  : ハッシュオブジェクトハンドル
'
'     戻り値 : 成功 = True (0以外)
'            : 失敗 = False(0)
'
Private Declare Function CryptDestroyHash Lib "advapi32.dll" ( _
                                        ByVal hHash As Long) As Long


'------------------------------------------------------------------------------------
' CSP 暗号キーインポート
'------------------------------------------------------------------------------------
Private Declare Function CryptImportKey Lib "advapi32.dll" ( _
                                        ByVal hProv As Long, _
                                        pbData As Any, _
                                        ByVal dwDataLen As Long, _
                                        ByVal hPubKey As Long, _
                                        ByVal dwFlags As Long, _
                                        phKey As Long) As Long

'====================================================================================
' クラスイベント
'====================================================================================

'------------------------------------------------------------------------------------
' 暗号・複合化中イベント
'------------------------------------------------------------------------------------
' CryptBytes : 処理済バイト
' FileBytes  : ファイルバイト(サイズ)
Public Event Crypting(ByVal CryptBytes As Long, ByVal FileBytes As Long)

'------------------------------------------------------------------------------------
' 暗号・複合化完了イベント
'------------------------------------------------------------------------------------
Public Event Crypted()

'====================================================================================
' プロパティ変数
'====================================================================================

Private m_strInFile     As String       ' 入力ファイル
Private m_strOutFile    As String       ' 出力ファイル
Private m_strPassword   As String       ' パスワード

Private Const ENCRYPT_BLOCK_SIZE = 2        ' デフォルトブロックサイズ
Private Const ENCRYPT_ALGORITHM = CALG_RC4  ' デフォルトアルゴリズム

'****************************************************************************************
' プライベートプロシジャ
'****************************************************************************************


'
'----------------------------------------------------------------------------------------
' ファイルの暗号化
'----------------------------------------------------------------------------------------
'
' [ I ] strFileIn   : 暗号化する元ファイル(既存ファイル)
' [ I ] strFileOut  : 暗号化されたファイル(新規に作成される)
' [ I ] strPassword : パスワード(""の場合ランダムセッションキー使用)
'
'       戻り値      : True  = 成功
'                   : False = 失敗
'
Private Function EncryptFile(ByVal strFileIn As String, _
                             ByVal strFileOut As String, _
                             ByVal strPassword As String) As Boolean
    Dim hProv           As Long ' プロバイダハンドル
    Dim hKey            As Long ' セッションキーハンドル
    Dim hXchgKey        As Long ' パブリックキー
    Dim dwKeyBlobLen    As Long ' エクスポートキーサイズ
    Dim pbKeyBlob()     As Byte ' エクスポートキー
    Dim hHash           As Long ' ハッシュハンドル
    Dim dwBlockLen      As Long ' ブロックサイズ
    Dim dwBufferLen     As Long ' バッファサイズ
    Dim pbBuffer()      As Byte ' バッファ
    Dim dwByte          As Long ' 実際のバッファサイズ
    Dim dwWriten        As Long ' 実際に書き込まれたサイズ
    Dim lngEOF          As Long ' ファイル終端フラグ
    Dim hIn             As Long ' 入力ファイルハンドル
    Dim hOut            As Long ' 出力ファイルハンドル
    Dim lngFileSize     As Long ' ファイルサイズ(進捗状況用)
    Dim lngEncryptSize  As Long ' 処理済みサイズ(進捗状況用)
    
    ' 戻り値初期化
    EncryptFile = False

    ' 入力ファイルのチェック & オープン
    If Dir(strFileIn) = "" Then GoTo done:
    hIn = CreateFile(strFileIn, _
                     GENERIC_READ, _
                     FILE_SHARE_READ, _
                     ByVal 0, _
                     OPEN_EXISTING, _
                     FILE_ATTRIBUTE_NORMAL, _
                     ByVal 0)
                     
    ' 出力ファイルのチェック & オープン
    If Dir(strFileOut) <> "" Then GoTo done:
    hOut = CreateFile(strFileOut, _
                      GENERIC_WRITE, _
                      FILE_SHARE_READ, _
                      ByVal 0, _
                      CREATE_ALWAYS, _
                      FILE_ATTRIBUTE_NORMAL, _
                      ByVal 0)
        
    ' デフォルト・プロパイダーハンドル取得
    If Not CBool(CryptAcquireContext(hProv, _
                                     ByVal vbNullString, _
                                     ByVal vbNullString, _
                                     PROV_RSA_FULL, 0)) Then

        If Not CBool(CryptAcquireContext(hProv, _
                                         ByVal vbNullString, _
                                         ByVal vbNullString, _
                                         PROV_RSA_FULL, _
                                         CRYPT_NEWKEYSET)) Then
            MsgBox "CryptAcquireContext Error!"
            GoTo done:
        End If
    End If

    If strPassword = "" Then
        ' ランダム・セッションキー作成
        If Not CBool(CryptGenKey(hProv, ENCRYPT_ALGORITHM, CRYPT_EXPORTABLE, hKey)) Then
            MsgBox "CryptGenKey Error!"
            GoTo done:
        End If
        ' パブリックキー取得
        If Not CBool(CryptGetUserKey(hProv, AT_KEYEXCHANGE, hXchgKey)) Then
            MsgBox "CryptGetUserKey Error!"
            GoTo done:
        End If
        ' キーメモリー割り当て
        If Not CBool(CryptExportKey(hKey, hXchgKey, SIMPLEBLOB, 0, ByVal 0, dwKeyBlobLen)) Then
            MsgBox "CryptExportKey Error!"
            GoTo done:
        End If
        If dwKeyBlobLen = 0 Then
            MsgBox "メモリー割り当て Error!"
            GoTo done:
        Else
            ReDim pbKeyBlob(dwKeyBlobLen - 1)
        End If
        ' セッションキーエクスポート
        If Not CBool(CryptExportKey(hKey, hXchgKey, SIMPLEBLOB, 0, pbKeyBlob(0), dwKeyBlobLen)) Then
            MsgBox "CryptExportKey Error!"
            GoTo done:
        End If
        ' パブリックキー開放
        Call CryptDestroyKey(hXchgKey)
        hXchgKey = 0
        ' キーサイズ書き込み
        Call WriteFile(hOut, dwKeyBlobLen, Len(dwKeyBlobLen), dwWriten, ByVal 0)
        ' キー書き込み
        Call WriteFile(hOut, pbKeyBlob(0), dwKeyBlobLen, dwWriten, ByVal 0)
    Else
        ' ハッシュ作成
        If Not CBool(CryptCreateHash(hProv, CALG_MD5, 0, 0, hHash)) Then
            MsgBox "CryptCreateHash Error!"
            GoTo done:
        End If
        ' パスワードデータハッシュ
        If Not CBool(CryptHashData(hHash, strPassword, Len(strPassword), 0)) Then
            MsgBox "CryptHashData Error!"
            GoTo done:
        End If
        ' ハッシュよりセッションキー取得
        If Not CBool(CryptDeriveKey(hProv, ENCRYPT_ALGORITHM, hHash, 0, hKey)) Then
            MsgBox "CryptDeriveKey Error!"
            GoTo done:
        End If
        ' ハッシュ削除
        Call CryptDestroyHash(hHash)
        hHash = 0
    End If

    ' 入力ファイルサイズ取得
    lngFileSize = GetFileSize(hIn, ByVal 0)
    
    ' 暗号化する1ブロックバイト数
    dwBlockLen = 1000 - 1000 Mod ENCRYPT_BLOCK_SIZE

    ' ブロックサイズ
    If ENCRYPT_BLOCK_SIZE > 1 Then
        dwBufferLen = dwBlockLen + ENCRYPT_BLOCK_SIZE
    Else
        dwBufferLen = dwBlockLen + 1
    End If

    ' メモリー割り当て
    ReDim pbBuffer(dwBufferLen - 1)

    ' 暗号化開始
    Do
        ' 入力ファイルの読み込み
        Call ReadFile(hIn, pbBuffer(0), dwBufferLen, dwByte, ByVal 0)
        ' ファイル終端チェック
        If dwByte < dwBufferLen Then lngEOF = 1
        ' Encrypt
        If Not CBool(CryptEncrypt(hKey, 0, lngEOF, 0, pbBuffer(0), dwByte, dwBufferLen)) Then
            MsgBox "CryptEncrypt Error!"
            GoTo done:
        End If
        ' 出力ファイルへ書き込み
        Call WriteFile(hOut, pbBuffer(0), dwByte, dwWriten, ByVal 0)
        ' 進捗状況イベント
        lngEncryptSize = lngEncryptSize + dwByte
        RaiseEvent Crypting(lngEncryptSize, lngFileSize)
    Loop While Not CBool(lngEOF)
    ' 戻り値セット
    EncryptFile = True
done:
    ' ハンドル開放
    Call CloseHandle(hIn)
    Call CloseHandle(hOut)
    Call CryptDestroyKey(hKey)
    Call CryptDestroyKey(hXchgKey)
    Call CryptDestroyHash(hHash)
    Call CryptReleaseContext(hProv, 0)
    ' 完了イベント
    RaiseEvent Crypted
End Function

'
'----------------------------------------------------------------------------------------
' ファイルの複合化
'----------------------------------------------------------------------------------------
'
' [ I ] strFileIn   : 複合化する元ファイル(暗号化されているファイル)
' [ I ] strFileOut  : 複合化されたファイル(新規に作成される)
' [ I ] strPassword : パスワード(""の場合ランダムセッションキー使用)
'
'       戻り値      : True  = 成功
'                   : False = 失敗
'
Private Function DecryptFile(ByVal strFileIn As String, _
                             ByVal strFileOut As String, _
                             ByVal strPassword As String) As Boolean
    Dim hProv           As Long ' プロバイダハンドル
    Dim hKey            As Long ' セッションキーハンドル
    Dim hXchgKey        As Long ' パブリックキー
    Dim dwKeyBlobLen    As Long ' エクスポートキーサイズ
    Dim pbKeyBlob()     As Byte ' エクスポートキー
    Dim hHash           As Long ' ハッシュハンドル
    Dim dwBlockLen      As Long ' ブロックサイズ
    Dim dwBufferLen     As Long ' バッファサイズ
    Dim pbBuffer()      As Byte ' バッファ
    Dim dwByte          As Long ' 実際のバッファサイズ
    Dim dwWriten        As Long ' 実際に書き込まれたサイズ
    Dim lngEOF          As Long ' ファイル終端フラグ
    Dim hIn             As Long ' 入力ファイルハンドル
    Dim hOut            As Long ' 出力ファイルハンドル
    Dim lngFileSize     As Long ' ファイルサイズ(進捗状況用)
    Dim lngEncryptSize  As Long ' 処理済みサイズ(進捗状況用)
    
    
    ' 戻り値初期化
    DecryptFile = False

    ' 入力ファイルのチェック & オープン
    If Dir(strFileIn) = "" Then GoTo done:
    hIn = CreateFile(strFileIn, _
                     GENERIC_READ, _
                     FILE_SHARE_READ, _
                     ByVal 0, _
                     OPEN_EXISTING, _
                     FILE_ATTRIBUTE_NORMAL, _
                     ByVal 0)

    ' 出力ファイルのチェック & オープン
    If Dir(strFileOut) <> "" Then GoTo done:
    hOut = CreateFile(strFileOut, _
                      GENERIC_WRITE, _
                      FILE_SHARE_READ, _
                      ByVal 0, _
                      CREATE_ALWAYS, _
                      FILE_ATTRIBUTE_NORMAL, _
                      ByVal 0)

    ' デフォルト・プロパイダーハンドル取得
    If Not CBool(CryptAcquireContext(hProv, _
                                     ByVal vbNullString, _
                                     ByVal vbNullString, _
                                     PROV_RSA_FULL, 0)) Then
        MsgBox "CryptAcquireContext Error!"
        GoTo done:
    End If
    
    
    If strPassword = "" Then
        ' 解読セッションキーサイズ
        Call ReadFile(hIn, dwKeyBlobLen, Len(dwKeyBlobLen), dwByte, ByVal 0)
        ' 解読セッションキーメモリー確保
        ReDim pbKeyBlob(dwKeyBlobLen - 1)
        ' 解読セッションキー読み込み
        Call ReadFile(hIn, pbKeyBlob(0), dwKeyBlobLen, dwByte, ByVal 0)
        ' CSP 暗号キーインポート
        If Not CBool(CryptImportKey(hProv, pbKeyBlob(0), dwKeyBlobLen, 0, 0, hKey)) Then
            MsgBox "CryptImportKey Error!"
            GoTo done:
        End If
    Else
        ' ハッシュ作成
        If Not CBool(CryptCreateHash(hProv, CALG_MD5, 0, 0, hHash)) Then
            MsgBox "CryptCreateHash Error!"
            GoTo done:
        End If
        ' パスワードデータハッシュ
        If Not CBool(CryptHashData(hHash, strPassword, Len(strPassword), 0)) Then
            MsgBox "CryptHashData Error!"
            GoTo done:
        End If
        ' ハッシュよりセッションキー取得
        If Not CBool(CryptDeriveKey(hProv, ENCRYPT_ALGORITHM, hHash, 0, hKey)) Then
            MsgBox "CryptDeriveKey Error!"
            GoTo done:
        End If
        ' ハッシュ削除
        Call CryptDestroyHash(hHash)
        hHash = 0
    End If

    ' 入力ファイルサイズ取得
    lngFileSize = GetFileSize(hIn, ByVal 0)
    
    ' 複合化する1ブロックバイト数
    dwBlockLen = 1000 - 1000 Mod ENCRYPT_BLOCK_SIZE

    ' ブロックサイズ
    If ENCRYPT_BLOCK_SIZE > 1 Then
        dwBufferLen = dwBlockLen + ENCRYPT_BLOCK_SIZE
    Else
        dwBufferLen = dwBlockLen + 1
    End If

    ' メモリー割り当て
    ReDim pbBuffer(dwBufferLen - 1)

    ' 複合化開始
    Do
        ' 入力ファイルの読み込み
        Call ReadFile(hIn, pbBuffer(0), dwBufferLen, dwByte, ByVal 0)
        ' ファイル終端チェック
        If dwByte < dwBufferLen Then lngEOF = 1
        ' Decrypt
        If Not CBool(CryptDecrypt(hKey, 0, lngEOF, 0, pbBuffer(0), dwByte)) Then
            MsgBox "CryptDecrypt Error!"
            GoTo done:
        End If
        ' 出力ファイルに書き込み
        Call WriteFile(hOut, pbBuffer(0), dwByte, dwWriten, ByVal 0)
        ' 進捗状況イベント
        lngEncryptSize = lngEncryptSize + dwByte
        RaiseEvent Crypting(lngEncryptSize, lngFileSize)
    Loop While Not CBool(lngEOF)
    ' 戻り値セット
    DecryptFile = True
done:
    ' ハンドル開放
    Call CloseHandle(hIn)
    Call CloseHandle(hOut)
    Call CryptDestroyKey(hKey)
    Call CryptDestroyHash(hHash)
    Call CryptReleaseContext(hProv, 0)
    ' 完了イベント
    RaiseEvent Crypted
End Function

'****************************************************************************************
' メソッド
'****************************************************************************************

' 複合化
'
Public Function Decrypt() As Boolean
    Decrypt = DecryptFile(m_strInFile, m_strOutFile, m_strPassword)
End Function

' 暗号化
'
Public Function Encrypt() As Boolean
    Encrypt = EncryptFile(m_strInFile, m_strOutFile, m_strPassword)
End Function

'****************************************************************************************
' プロパティ
'****************************************************************************************

' 入力ファイル
'
Public Property Get InFile() As String
    InFile = m_strInFile
End Property

Public Property Let InFile(ByVal vNewValue As String)
    m_strInFile = vNewValue
End Property

' 出力ファイル
'
Public Property Get OutFile() As String
    OutFile = m_strOutFile
End Property

Public Property Let OutFile(ByVal vNewValue As String)
    m_strOutFile = vNewValue
End Property

' パスワード
'
Public Property Get Password() As String
    Password = m_strPassword
End Property

Public Property Let Password(ByVal vNewValue As String)
    m_strPassword = vNewValue
End Property

'****************************************************************************************
' クラス
'****************************************************************************************

Private Sub Class_Initialize()
End Sub
