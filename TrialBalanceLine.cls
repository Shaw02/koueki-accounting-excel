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

Private m_AccountCode    As Long
Private m_AccountName    As String

Private m_FS             As String
Private m_Category       As String
Private m_SubCategory    As String

Private m_Sector         As String
Private m_Project        As String
Private m_FundType       As NetAssetType

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

'===============================
' Property
'===============================
Public Property Let AccountCode(ByVal v As Long): m_AccountCode = v: End Property
Public Property Get AccountCode() As Long: AccountCode = m_AccountCode: End Property

Public Property Let AccountName(ByVal v As String): m_AccountName = v: End Property
Public Property Get AccountName() As String: AccountName = m_AccountName: End Property

Public Property Let FS(ByVal v As String): m_FS = v: End Property
Public Property Get FS() As String: FS = m_FS: End Property

Public Property Let Category(ByVal v As String): m_Category = v: End Property
Public Property Get Category() As String: Category = m_Category: End Property

Public Property Let SubCategory(ByVal v As String): m_SubCategory = v: End Property
Public Property Get SubCategory() As String: SubCategory = m_SubCategory: End Property

Public Property Let Sector(ByVal v As String): m_Sector = v: End Property
Public Property Get Sector() As String: Sector = m_Sector: End Property

Public Property Let Project(ByVal v As String): m_Project = v: End Property
Public Property Get Project() As String: Project = m_Project: End Property

Public Property Let FundType(ByVal v As NetAssetType): m_FundType = v: End Property
Public Property Get FundType() As NetAssetType: FundType = m_FundType: End Property

Public Property Let BeginningBalance(ByVal v As Currency): m_BeginningBalance = v: End Property
Public Property Get BeginningBalance() As Currency: BeginningBalance = m_BeginningBalance: End Property

Public Property Let Debit(ByVal v As Currency): m_Debit = v: End Property
Public Property Get Debit() As Currency: Debit = m_Debit: End Property

Public Property Let Credit(ByVal v As Currency): m_Credit = v: End Property
Public Property Get Credit() As Currency: Credit = m_Credit: End Property

Public Property Let EndingBalance(ByVal v As Currency): m_EndingBalance = v: End Property
Public Property Get EndingBalance() As Currency: EndingBalance = m_EndingBalance: End Property

Public Property Get BeginningBalanceForDisplay() As Currency

    If Me.FS = "BS" Then
        BeginningBalanceForDisplay = Me.BeginningBalance
    Else
        BeginningBalanceForDisplay = 0
    End If

End Property

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

