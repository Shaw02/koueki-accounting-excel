VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NetAssetLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================
'   純資産管理1行（事業別）
'===============================

Private m_Key               As LedgerKey

Private m_SubAccountCode    As Long
Private m_SubAccountCate    As String
Private m_SubAccountName    As String
Private m_NetAssetType      As NetAssetType

Private m_BeginningBalance  As Currency
Private m_Debit             As Currency   '増加（収益）
Private m_Credit            As Currency   '減少（費用）
Private m_EndingBalance     As Currency

'****************************************
'*      Property                        *
'****************************************
Public Property Get SubAccountCode() As Long: SubAccountCode = m_SubAccountCode: End Property
Public Property Get SubAccountCate() As String: SubAccountCate = m_SubAccountCate: End Property
Public Property Get SubAccountName() As String: SubAccountName = m_SubAccountName: End Property
Public Property Get NetAssetType() As NetAssetType: NetAssetType = m_NetAssetType: End Property

Public Property Let BeginningBalance(ByVal v As Currency): m_BeginningBalance = v: End Property
Public Property Get BeginningBalance() As Currency: BeginningBalance = m_BeginningBalance: End Property

Public Property Let debit(ByVal v As Currency): m_Debit = v: End Property
Public Property Get debit() As Currency: debit = m_Debit: End Property

Public Property Let credit(ByVal v As Currency): m_Credit = v: End Property
Public Property Get credit() As Currency: credit = m_Credit: End Property

Public Property Let EndingBalance(ByVal v As Currency): m_EndingBalance = v: End Property
Public Property Get EndingBalance() As Currency: EndingBalance = m_EndingBalance: End Property

'===============================
' 初期化
'===============================
Public Sub Initialize( _
        ByVal subCode As Long, _
        ByVal subCate As String, _
        ByVal subName As String, _
        ByVal NetAssetType As NetAssetType)

    m_SubAccountCode = subCode
    m_SubAccountCate = subCate
    m_SubAccountName = subName
    m_NetAssetType = NetAssetType

End Sub

'===============================
'   期首
'===============================
Public Sub SetBeginning(ByVal amount As Currency)
    m_BeginningBalance = amount
    m_Debit = 0
    m_Credit = 0
End Sub

'===============================
'   加算（集約）
'===============================
Public Sub AddAmount( _
    ByVal debit As Currency, _
    ByVal credit As Currency)

    m_Debit = m_Debit + debit
    m_Credit = m_Credit + credit

End Sub

'===============================
'   期末計算
'===============================
Public Sub CalcEnding()
    m_EndingBalance = m_BeginningBalance - m_Debit + m_Credit
End Sub

