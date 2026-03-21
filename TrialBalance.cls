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

Private Sub Class_Initialize()

    Set m_Lines = New Collection
    Set m_Map = CreateObject("Scripting.Dictionary")

End Sub

Public Property Get Lines() As Collection
    Set Lines = m_Lines
End Property

'===============================
' 残高試算表 作成
'===============================
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
        ' Sector（空は共通）
        '--------------------------------
        Dim Sector As String
        If Trim(acc.SubAccountCate) = "" Then
            Sector = "共通"
        Else
            Sector = acc.SubAccountCate
        End If

        '--------------------------------
        ' Project（PLのみ）
        '--------------------------------
        Dim Project As String
        If acc.FS = "収益" Or acc.FS = "費用" Then
            Project = acc.SubAccountName
        Else
            Project = ""
        End If

        '--------------------------------
        ' 現時点：1対1なので基本は新規作成
        ' （将来：配賦で複数→AddAmount使用）
        '--------------------------------
        If m_Map.Exists(key) Then

            Set line = m_Map(key)
            line.AddAmount _
                acc.BeginningBalance, _
                acc.Debit, _
                acc.Credit, _
                acc.EndingBalance

        Else

            Set line = New TrialBalanceLine

            '--- Key
            Set line.key = keyObj

            '--- 科目
            line.AccountCode = acc.AccountCode
            line.AccountName = acc.MajorAccount
            line.FS = acc.FS
            line.Category = acc.Category
            line.SubCategory = acc.SubCategory

            '--- 補助
            line.Sector = Sector
            line.Project = Project
            line.FundType = acc.NetAssetType

            '--- 金額（残高試算表）
            line.BeginningBalance = acc.BeginningBalance
            line.Debit = acc.Debit
            line.Credit = acc.Credit
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
    
    For Each key In sortedDict.keys
        m_Lines.Add sortedDict(key)
    Next key

End Sub
