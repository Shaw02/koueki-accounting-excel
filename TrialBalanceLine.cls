VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrialBalanceLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================
' 試算表1行
'===============================

Private m_Key As ledgerKey

Private m_Account           As account

Private m_SubAccountCode    As Long             ' 補助科目コード
Private m_SubAccountCate    As String           ' 会計区分（公1/公2/公共/収/法/etc..）
Private m_SubAccountName    As String           ' 補助科目名（VLookup）
Private m_NetAssetType      As NetAssetType     ' 指定（総勘定元帳のみ）

Private m_BeginningBalance As Currency
Private m_Debit            As Currency
Private m_Credit           As Currency
Private m_EndingBalance    As Currency

'===============================
' Key
'===============================
Public Property Set key(ByVal v As ledgerKey)
    Set m_Key = v
End Property

Public Property Get key() As ledgerKey
    Set key = m_Key
End Property

Public Property Get KeyString() As String
    KeyString = m_Key.key
End Property


'==============================================================================
'   勘定科目クラスを取得
'==============================================================================
Public Property Get account() As account
    Set account = m_Account
End Property

'==============================================================================
'   勘定科目コードを取得
'==============================================================================
Public Property Get AccountCode() As Long
    AccountCode = m_Account.code
End Property

'==============================================================================
'   勘定科目名を取得
'==============================================================================
Public Property Get MajorAccount() As String
    MajorAccount = m_Account.MajorAccount
End Property

'==============================================================================
'   勘定科目名を取得
'==============================================================================
Public Property Get MiddleAccount() As String
    MiddleAccount = m_Account.MiddleAccount
End Property

'==============================================================================
'   補助科目コードを取得
'==============================================================================
Public Property Get SubAccountCode() As Long
    SubAccountCode = m_SubAccountCode
End Property

'==============================================================================
'   区分を取得
'==============================================================================
Public Property Get SubAccountCate() As String
    SubAccountCate = m_SubAccountCate
End Property

'==============================================================================
'   補助科目名を取得
'==============================================================================
Public Property Get SubAccountName() As String
    SubAccountName = m_SubAccountName
End Property

'==============================================================================
'   一般／指定を取得
'==============================================================================
Public Property Get NetAssetType() As NetAssetType
    NetAssetType = m_NetAssetType
End Property

'==============================================================================
'   指定か？を取得
'==============================================================================
Public Property Get Class() As String
    If m_NetAssetType = natDesignated Then
        Class = "（指定）"
    Else
        Class = ""
    End If
End Property

'==============================================================================
'   決算取得
'==============================================================================
Public Property Let BeginningBalance(ByVal v As Currency): m_BeginningBalance = v: End Property
Public Property Get BeginningBalance() As Currency: BeginningBalance = m_BeginningBalance: End Property

Public Property Let Debit(ByVal v As Currency): m_Debit = v: End Property
Public Property Get Debit() As Currency: Debit = m_Debit: End Property

Public Property Let Credit(ByVal v As Currency): m_Credit = v: End Property
Public Property Get Credit() As Currency: Credit = m_Credit: End Property

Public Property Let EndingBalance(ByVal v As Currency): m_EndingBalance = v: End Property
Public Property Get EndingBalance() As Currency: EndingBalance = m_EndingBalance: End Property

Public Property Get BeginningBalanceForDisplay() As Currency

    If Me.account.FS = "BS" Then
        BeginningBalanceForDisplay = Me.BeginningBalance
    Else
        BeginningBalanceForDisplay = 0
    End If

End Property

'==============================================================================
'       初期化
'------------------------------------------------------------------------------
'   Contents:
'       勘定科目情報を初期化します
'
'   Input:
'       acc             勘定科目
'       subCode         補助科目コード
'       subCate         会計区分
'       subName         補助科目名
'       NetAssetType    一般／指定
'
'==============================================================================
Public Sub Initialize( _
        ByVal acc As account, _
        ByVal SubCode As Long, _
        ByVal subCate As String, _
        ByVal subName As String, _
        ByVal NetAssetType As NetAssetType)

    Set m_Account = acc
    m_SubAccountCode = SubCode
    m_SubAccountCate = subCate
    m_SubAccountName = subName
    m_NetAssetType = NetAssetType

End Sub

'===============================
' 加算（集約）
'===============================
Public Sub AddAmount( _
    ByVal beginning As Currency, _
    ByVal Debit As Currency, _
    ByVal Credit As Currency, _
    ByVal ending As Currency)

    '--------------------------------------------
    ' 将来の配賦対応用
    '
    ' 同一キー（勘定科目＋補助科目＋指定区分＋α）に
    ' 複数の元帳データが集約される場合に使用する
    '
    ' 現時点では LedgerAccount が一意のため
    ' 呼ばれないが、将来必ず必要になる
    '--------------------------------------------

    m_BeginningBalance = m_BeginningBalance + beginning
    m_Debit = m_Debit + Debit
    m_Credit = m_Credit + Credit
    m_EndingBalance = m_EndingBalance + ending

End Sub

