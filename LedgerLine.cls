VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LedgerLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'   勘定科目元帳 ／ 各勘定科目 ／ １行
'------------------------------------------------------------------------------
'   Contents:
'       総勘定元帳・補助元帳 共通の１行レコード
'       残高式元帳
'
'==============================================================================

Option Explicit

Private m_Date As Variant                   '日付
Private m_CounterAccount As Account         '相手科目
Private m_Summary As String                 '摘要

Private m_Debit As Currency                 '借方            '
Private m_Credit As Currency                '貸方
Private m_Balance As Currency               '残高

'==============================================================================
'   日付取得
'==============================================================================
Public Property Get entryDate() As Variant
    entryDate = m_Date
End Property


'==============================================================================
'   相手科目取得
'==============================================================================
Public Property Get counterMajorAccount() As String
    If m_CounterAccount Is Nothing Then
        counterMajorAccount = ""
    Else
        counterMajorAccount = m_CounterAccount.MajorAccount
    End If
End Property


'==============================================================================
'   相手科目取得
'==============================================================================
Public Property Get counterMiddleAccount() As String
    If m_CounterAccount Is Nothing Then
        counterMiddleAccount = ""
    Else
        counterMiddleAccount = m_CounterAccount.MiddleAccount
    End If
End Property


'==============================================================================
'   摘要取得
'==============================================================================
Public Property Get summary() As String
    summary = m_Summary
End Property


'==============================================================================
'   借方取得
'==============================================================================
Public Property Get Debit() As Currency
    Debit = m_Debit
End Property


'==============================================================================
'   貸方取得
'==============================================================================
Public Property Get Credit() As Currency
    Credit = m_Credit
End Property


'==============================================================================
'   残高取得
'==============================================================================
Public Property Get balance() As Currency
    balance = m_Balance
End Property

'==============================================================================
'       元帳行初期化    （前期繰越）
'------------------------------------------------------------------------------
'   Contents:
'       元帳１行を初期化します
'
'   Input:
'       entryDate
'       summary
'       debitAmount
'       creditAmount
'
'==============================================================================
Public Sub InitializeBegin( _
        ByVal entryDate As Variant, _
        ByVal summary As String, _
        ByVal debitAmount As Currency, _
        ByVal creditAmount As Currency)

    '--------------------------------
    ' 借方貸方チェック
    '--------------------------------
    
    If debitAmount <> 0 And creditAmount <> 0 Then
        Err.Raise vbObjectError + 4001, _
                  "LedgerLine.Initialize", _
                  "借方と貸方の両方に金額が入力されています"
    End If
    
    If debitAmount < 0 Then
        Err.Raise vbObjectError + 4002, _
                  "LedgerLine.Initialize", _
                  "借方の金額が負数です"
    End If

    If creditAmount < 0 Then
        Err.Raise vbObjectError + 4002, _
                  "LedgerLine.Initialize", _
                  "貸方の金額が負数です"
    End If

    m_Date = entryDate
    m_Summary = summary
    
    m_Debit = debitAmount
    m_Credit = creditAmount
    
End Sub


'==============================================================================
'       元帳行初期化
'------------------------------------------------------------------------------
'   Contents:
'       元帳１行を初期化します
'
'   Input:
'       entryDate
'       counterAccount
'       summary
'       debitAmount
'       creditAmount
'
'==============================================================================
Public Sub Initialize( _
        ByVal entryDate As Variant, _
        ByVal counterAccount As Account, _
        ByVal summary As String, _
        ByVal debitAmount As Currency, _
        ByVal creditAmount As Currency)

    '--------------------------------
    ' 借方貸方チェック
    '--------------------------------
    
    If debitAmount <> 0 And creditAmount <> 0 Then
        Err.Raise vbObjectError + 4001, _
                  "LedgerLine.Initialize", _
                  "借方と貸方の両方に金額が入力されています"
    End If
    
    If debitAmount < 0 Then
        Err.Raise vbObjectError + 4002, _
                  "LedgerLine.Initialize", _
                  "借方の金額が負数です"
    End If

    If creditAmount < 0 Then
        Err.Raise vbObjectError + 4002, _
                  "LedgerLine.Initialize", _
                  "貸方の金額が負数です"
    End If

    m_Date = entryDate
    Set m_CounterAccount = counterAccount
    m_Summary = summary
    
    m_Debit = debitAmount
    m_Credit = creditAmount
    
End Sub


'==============================================================================
'       残高設定
'------------------------------------------------------------------------------
'   Contents:
'       LedgerAccount から残高を設定
'
'==============================================================================
Public Sub SetBalance(ByVal balance As Currency)

    m_Balance = balance

End Sub


'==============================================================================
'       貸借判定
'------------------------------------------------------------------------------
'   Contents:
'       残高から貸借を判定
'
'   Output:
'       借 / 貸
'
'==============================================================================
Public Function DebitCreditLabel() As String

    If m_Balance < 0 Then
        DebitCreditLabel = "貸"
    Else
        DebitCreditLabel = "借"
    End If

End Function


'==============================================================================
'       残高絶対値
'------------------------------------------------------------------------------
'   Contents:
'       Excel出力用
'
'==============================================================================
Public Function BalanceAbs() As Currency

    BalanceAbs = Abs(m_Balance)

End Function

