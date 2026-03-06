VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LedgerAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'   勘定科目元帳 ／ 各勘定科目
'------------------------------------------------------------------------------
'   Contents:
'       勘定科目ごとの元帳を保持するクラス
'
'       ・LedgerLine を内部に保持
'       ・残高式元帳に対応
'
'==============================================================================

Option Explicit

Private m_AccountCode       As Long             ' 勘定科目コード
Private m_MajorAccount      As String           ' 大科目（VLookup）
Private m_MiddleAccount     As String           ' 中科目（VLookup）
Private m_SubAccountCode    As Variant          ' 補助科目コード
Private m_SubAccountName    As String           ' 補助科目名（VLookup）
Private m_NetAssetType      As NetAssetType     ' 指定（総勘定元帳のみ）

Private m_Lines As Collection                   ' 各行（LedgerLine）を保持する

Private m_last  As Currency                     ' 前期期末残高  （資産・負債だったら、前期繰越にもなる）
Private m_Balance As Currency                   ' 期末残高  （来期期初残高）

'==============================================================================
'   クラス初期化
'==============================================================================
Private Sub Class_Initialize()

    Set m_Lines = New Collection
    m_Balance = 0
    m_last = 0

End Sub


'==============================================================================
'   勘定科目コード取得
'==============================================================================
Public Property Get AccountCode() As Long
    AccountCode = m_AccountCode
End Property


'==============================================================================
'   勘定科目名取得
'==============================================================================
Public Property Get MajorAccount() As String
    MajorAccount = m_MajorAccount
End Property

'==============================================================================
'   勘定科目名取得
'==============================================================================
Public Property Get MiddleAccount() As String
    MiddleAccount = m_MiddleAccount
End Property

'==============================================================================
'   補助科目コード取得
'==============================================================================
Public Property Get SubAccountCode() As Long
    SubAccountCode = m_SubAccountCode
End Property

'==============================================================================
'   補助科目名取得
'==============================================================================
Public Property Get SubAccountName() As String
    SubAccountName = m_SubAccountName
End Property

'==============================================================================
'   指定
'==============================================================================
Public Property Get Class() As String
    If m_NetAssetType = natDesignated Then
        Class = "（指定）"
    Else
        Class = ""
    End If
End Property


'==============================================================================
'   元帳行取得
'==============================================================================
Public Property Get Lines() As Collection
    Set Lines = m_Lines
End Property


'==============================================================================
'       勘定科目初期化
'------------------------------------------------------------------------------
'   Contents:
'       勘定科目情報を初期化します
'
'   Input:
'       accountCode
'       accountName
'       subCode
'       subName
'       classType
'
'==============================================================================
Public Sub Initialize( _
        ByVal AccountCode As Long, _
        ByVal MajorAccount As String, _
        ByVal MiddleAccount As String, _
        ByVal subCode As Variant, _
        ByVal subName As String, _
        ByVal NetAssetType As NetAssetType)

    m_AccountCode = AccountCode
    m_MajorAccount = MajorAccount
    m_MiddleAccount = MiddleAccount
    m_SubAccountCode = subCode
    m_SubAccountName = subName
    m_NetAssetType = NetAssetType

End Sub


'==============================================================================
'       前期繰越登録
'------------------------------------------------------------------------------
'   Contents:
'       元帳の最初に前期繰越を追加します
'
'   Input:
'       balance     前期残高
'
'==============================================================================
Public Sub AddOpeningBalance(ByVal balance As Currency)

    Dim line As LedgerLine
    Set line = New LedgerLine

    '昨年度実績
    m_last = balance
    
    '資産＆負債の場合は、前期繰越を作る
    If m_AccountCode <> 0 And m_AccountCode < 30000 Then
        Call line.Initialize( _
            DateSerial(Year(Date), 4, 1), _
            "", _
            "", _
            "前期繰越", _
            IIf(balance > 0, balance, 0), _
            IIf(balance < 0, -balance, 0), _
            m_NetAssetType)
    
        m_Balance = balance
    
        line.SetBalance m_Balance
    
        m_Lines.Add line
    
    End If

End Sub


'==============================================================================
'       仕訳追加
'------------------------------------------------------------------------------
'   Contents:
'       元帳に仕訳を追加します
'
'   Input:
'       entryDate
'       counterAccount
'       summary
'       debitAmount
'       creditAmount
'
'==============================================================================
Public Sub AddLine( _
        ByVal entryDate As Variant, _
        ByVal counterMajorAccount As String, _
        ByVal counterMiddleAccount As String, _
        ByVal summary As String, _
        ByVal debitAmount As Currency, _
        ByVal creditAmount As Currency)

    Dim line As LedgerLine
    Set line = New LedgerLine

    '--------------------------------
    ' 残高更新
    '--------------------------------
    m_Balance = m_Balance + debitAmount
    m_Balance = m_Balance - creditAmount

    '--------------------------------
    ' LedgerLine 作成
    '--------------------------------
    Call line.Initialize( _
        entryDate, _
        counterMajorAccount, _
        counterMiddleAccount, _
        summary, _
        debitAmount, _
        creditAmount, _
        m_NetAssetType)

    line.SetBalance m_Balance

    m_Lines.Add line

End Sub


'==============================================================================
'       行数取得
'------------------------------------------------------------------------------
'   Output:
'       元帳行数
'
'==============================================================================
Public Property Get Count() As Long
    Count = m_Lines.Count
End Property


'==============================================================================
'       元帳行取得
'------------------------------------------------------------------------------
'   Input:
'       index
'
'   Output:
'       LedgerLine
'
'==============================================================================
Public Property Get line(ByVal index As Long) As LedgerLine

    Set line = m_Lines(index)

End Property

