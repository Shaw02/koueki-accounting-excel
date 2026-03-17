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

Private m_Account           As Account          ' 勘定科目マスターより

Private m_SubAccountCode    As Variant          ' 補助科目コード
Private m_SubAccountCate    As String           ' 補助科目コード
Private m_SubAccountName    As String           ' 補助科目名（VLookup）
Private m_NetAssetType      As NetAssetType     ' 指定（総勘定元帳のみ）

Private m_Lines As Collection                   ' 各行（LedgerLine）を保持する

Private m_BeginningBalance  As Currency         ' 期初残高  （資産・負債だったら、前期繰越にもなる）
Private m_EndingBalance     As Currency         ' 期末残高  （来期期初残高）
Private m_Debit             As Currency         ' 借方合計
Private m_Credit            As Currency         ' 貸方合計

'==============================================================================
'   勘定科目コード取得
'==============================================================================
Public Property Get AccountCode() As Long
    AccountCode = m_Account.code
End Property


'==============================================================================
'   勘定科目名取得
'==============================================================================
Public Property Get MajorAccount() As String
    MajorAccount = m_Account.MajorAccount
End Property

'==============================================================================
'   勘定科目名取得
'==============================================================================
Public Property Get MiddleAccount() As String
    MiddleAccount = m_Account.MiddleAccount
End Property

'==============================================================================
'   補助科目コード取得
'==============================================================================
Public Property Get SubAccountCode() As Variant
    SubAccountCode = m_SubAccountCode
End Property

'==============================================================================
'   区分
'==============================================================================
Public Property Get SubAccountCate() As String
    SubAccountCate = m_SubAccountCate
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
'   期初残高
'==============================================================================
Public Property Get BeginningBalance() As Currency
    BeginningBalance = m_BeginningBalance
End Property


'==============================================================================
'   期末残高
'==============================================================================
Public Property Get EndingBalance() As Currency
    EndingBalance = m_EndingBalance
End Property


'==============================================================================
'   借方合計
'==============================================================================
Public Property Get Debit() As Currency
    Debit = m_Debit
End Property


'==============================================================================
'   貸方合計
'==============================================================================
Public Property Get Credit() As Currency
    Credit = m_Credit
End Property

'==============================================================================
'   試算表用
'==============================================================================
Public Property Get TrialBalanceAmount() As Currency

    Select Case m_Account.FS

        Case "資産", "費用"
            TrialBalanceAmount = m_EndingBalance

        Case "負債", "純資産", "収益"
            TrialBalanceAmount = -m_EndingBalance

    End Select

End Property


'==============================================================================
'   カテゴリーの取得
'==============================================================================
Public Property Get FS() As String
    FS = m_Account.FS
End Property

Public Property Get Category() As String
    Category = m_Account.Category
End Property

Public Property Get SubCategory() As String
    SubCategory = m_Account.SubCategory
End Property



'==============================================================================
'   クラス初期化
'==============================================================================
Private Sub Class_Initialize()

    Set m_Lines = New Collection
    
    m_BeginningBalance = 0
    m_EndingBalance = 0
    m_Debit = 0
    m_Credit = 0

End Sub


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
        ByVal acc As Account, _
        ByVal subCode As Variant, _
        ByVal subCate As String, _
        ByVal subName As String, _
        ByVal NetAssetType As NetAssetType)

    Set m_Account = acc
    m_SubAccountCode = subCode
    m_SubAccountCate = subCate
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
    m_BeginningBalance = balance
    
    '資産＆負債の場合は、前期繰越を作る
    If m_Account.code <> 0 And m_Account.code < 30000 Then
        If balance <> 0 Then
            Call line.InitializeBegin( _
                DateSerial(Year(Date), 4, 1), _
                "前期繰越", _
                IIf(balance > 0, balance, 0), _
                IIf(balance < 0, -balance, 0))
        
            m_EndingBalance = balance
            
            line.SetBalance (balance)
        
            m_Lines.Add line
        End If
    
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
        ByVal counterAccount As Account, _
        ByVal summary As String, _
        ByVal debitAmount As Currency, _
        ByVal creditAmount As Currency, _
        Optional ByVal isMulti As Boolean = False)

    Dim line As LedgerLine
    Set line = New LedgerLine

    '--------------------------------
    ' 残高更新
    '--------------------------------
    m_EndingBalance = m_EndingBalance + debitAmount
    m_EndingBalance = m_EndingBalance - creditAmount

    m_Debit = m_Debit + debitAmount
    m_Credit = m_Credit + creditAmount

    '--------------------------------
    ' LedgerLine 作成
    '--------------------------------
    Call line.Initialize( _
        entryDate, _
        counterAccount, _
        summary, _
        debitAmount, _
        creditAmount, _
        isMulti)

    line.SetBalance m_EndingBalance

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

