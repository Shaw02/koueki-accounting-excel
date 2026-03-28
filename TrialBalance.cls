VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrialBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Lines As Collection
Private m_Map   As Object ' Dictionary（将来の配賦用）

'==============================================================================
'   Internal state clear
'==============================================================================
Private Sub Class_Initialize()

    Call Clear

End Sub

Public Sub Clear()

    Set m_Lines = New Collection
    Set m_Map = CreateObject("Scripting.Dictionary")

End Sub

Public Property Get Lines() As Collection
    Set Lines = m_Lines
End Property

'==============================================================================
'   汎用集計     （※補助科目を集計したクラスのみ、呼出し可）
'------------------------------------------------------------------------------
'   Input:
'       fundType        一般／指定
'       targetGroup    BS / PL
'       targetColumn   BEGIN / END / FLOW
'
'   Output:
'       集計金額
'
'==============================================================================
Public Function SumAmount( _
    ByVal fundType As NetAssetType, _
    ByVal targetGroup As AccountFSKind, _
    ByVal targetColumn As String _
) As Currency

    Dim line As TrialBalanceLine
    Dim total As Currency
    
    For Each line In m_Lines
        
        '--------------------------------
        ' 純資産は除外
        If line.account.CategoryEnum = ackNetAsset Then GoTo ContinueLoop

        '--------------------------------
        ' 一般／指定
        If line.NetAssetType <> fundType Then GoTo ContinueLoop
        
        '--------------------------------
        ' 集計行（PL / BS）
        If line.account.FSEnum <> targetGroup Then GoTo ContinueLoop
        
        '--------------------------------
        ' 列
        total = total + line.Amount(targetGroup, targetColumn)
        
ContinueLoop:
    Next
    
    SumAmount = total

End Function

'==============================================================================
'   昨年度の残高試算表　読み込み
'------------------------------------------------------------------------------
'   Input:
'
'==============================================================================
Public Sub SetLastResult( _
    ByVal account As account, _
    ByVal SubAccountCode As Long, _
    ByVal SubAccountCate As String, _
    ByVal SubAccountName As String, _
    ByVal Class As NetAssetType, _
    ByVal BeginningBalance As Currency, _
    ByVal debit As Currency, _
    ByVal credit As Currency, _
    ByVal EndingBalance As Currency)

    Dim line As TrialBalanceLine
    
    '--------------------------------
    ' Key（Ledgerと完全一致）
    '--------------------------------
    Dim keyObj As ledgerKey
    Dim key As String
    
    Set keyObj = New ledgerKey
    Call keyObj.Initialize(account.code, SubAccountCode, Class)
    key = keyObj.key

    '--------------------------------
    ' ライン作成
    '--------------------------------
    If m_Map.Exists(key) Then
        Err.Raise vbObjectError + 3001, "TrialBalance:SetLastResult", key & ":が重複しています。"
    Else

        Set line = New TrialBalanceLine

        '--- Key
        Set line.key = keyObj

        '--- 科目
        Call line.Initialize( _
            account, _
            SubAccountCode, _
            SubAccountCate, _
            SubAccountName, _
            Class)

        '--- 金額（残高試算表）
        line.BeginningBalance = BeginningBalance
        line.debit = debit
        line.credit = credit
        line.EndingBalance = EndingBalance

        m_Map.Add key, line
        m_Lines.Add line

    End If

End Sub

'==============================================================================
'   元帳から今期の残高試算表の作成
'------------------------------------------------------------------------------
'   Input:
'       ledgers     元帳
'
'==============================================================================
Public Sub Build(ByVal ledgers As Object)

    Dim acc As LedgerAccount
    Dim line As TrialBalanceLine
    Dim keyObj As ledgerKey
    Dim key As String
    Dim it As Variant

    For Each it In ledgers
        Set acc = ledgers.Item(it)

        '--------------------------------
        ' Key（Ledgerと完全一致）
        '--------------------------------
        Set keyObj = New ledgerKey
        keyObj.Initialize acc.AccountCode, acc.SubAccountCode, acc.NetAssetType
        
        key = keyObj.key

        '--------------------------------
        ' 現時点：1対1なので基本は新規作成
        ' （将来：配賦で複数→AddAmount使用）
        '--------------------------------
        If m_Map.Exists(key) Then

            Set line = m_Map(key)
            Call line.AddAmount( _
                acc.BeginningBalance, _
                acc.debit, _
                acc.credit, _
                acc.EndingBalance)

        Else

            Set line = New TrialBalanceLine

            '--- Key
            Set line.key = keyObj

            '--- 科目
            Call line.Initialize( _
                acc.account, _
                acc.SubAccountCode, _
                acc.SubAccountCate, _
                acc.SubAccountName, _
                acc.NetAssetType)

            '--- 金額（残高試算表）
            line.BeginningBalance = acc.BeginningBalance
            line.debit = acc.debit
            line.credit = acc.credit
            line.EndingBalance = acc.EndingBalance

            m_Map.Add key, line
            m_Lines.Add line

        End If

    Next

End Sub

'==============================================================================
'   ソート
'==============================================================================
Public Sub SortByAccountCode()

    Dim sortedDict As Object
    Dim key As Variant
    
    '--------------------------------
    ' Dictionaryをソート
    '--------------------------------
    Set sortedDict = SortDictionaryByKey(m_Map)
    
    '--------------------------------
    ' Collectionを作り直す
    '--------------------------------
    Set m_Lines = New Collection
    
    For Each key In sortedDict.Keys
        m_Lines.Add sortedDict(key)
    Next key

End Sub
