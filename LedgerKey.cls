VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LedgerKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'   元帳キー
'------------------------------------------------------------------------------
'   Contents:
'       総勘定元帳・補助元帳を識別するキー
'
'       勘定科目
'       補助科目
'       指定区分
'
'==============================================================================

Option Explicit

Private m_AccountCode As Long
Private m_SubAccountCode As Long
Private m_Class As NetAssetType

'==============================================================================
'   初期化
'==============================================================================
Public Sub Initialize( _
    ByVal AccountCode As Long, _
    ByVal SubAccountCode As Long, _
    ByVal ClassType As NetAssetType)

    m_AccountCode = AccountCode
    m_SubAccountCode = SubAccountCode
    m_Class = ClassType

End Sub

'==============================================================================
'   勘定科目コード取得
'==============================================================================
Public Property Get AccountCode() As Long
    AccountCode = m_AccountCode
End Property

'==============================================================================
'   補助科目コード取得
'==============================================================================
Public Property Get SubAccountCode() As Long
    SubAccountCode = m_SubAccountCode
End Property

'==============================================================================
'   指定区分取得
'==============================================================================
Public Property Get ClassType() As NetAssetType
    ClassType = m_Class
End Property

'==============================================================================
'   Dictionaryキー文字列
'==============================================================================
Public Function key() As String

    key = CStr(m_AccountCode) & "|" & _
          CStr(m_SubAccountCode) & "|" & _
          CStr(m_Class)

End Function


