VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccountingEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'   仕訳の１レコード
'==============================================================================
Option Explicit

'==============================================================================
'   メンバー変数
'==============================================================================
Private m_EntryDate         As Variant       ' 日付
Private m_Debit             As entrySide     ' 借方
Private m_Credit            As entrySide     ' 貸方
Private m_Summary           As String        ' 摘要

'==============================================================================
'   クラスの初期化
'==============================================================================
Private Sub Class_Initialize()
    Set m_Debit = New entrySide
    Set m_Credit = New entrySide
End Sub

'==============================================================================
'   クラスの解放
'==============================================================================
Private Sub Class_Terminate()
    Set m_Debit = Nothing
    Set m_Credit = Nothing
End Sub

'==============================================================================
'   メンバー変数への代入
'==============================================================================
Public Sub Initialize( _
    ByVal entryDate As Date, _
    ByVal D_AccountCode As Variant, _
    ByVal D_MajorAccount As String, _
    ByVal D_MiddleAccount As String, _
    ByVal D_SubAccountCode As Variant, _
    ByVal D_SubAccountName As String, _
    ByVal D_Class As Variant, _
    ByVal D_Amount As Currency, _
    ByVal C_AccountCode As Variant, _
    ByVal C_MajorAccount As String, _
    ByVal C_MiddleAccount As String, _
    ByVal C_SubAccountCode As Variant, _
    ByVal C_SubAccountName As String, _
    ByVal C_Class As Variant, _
    ByVal C_Amount As Currency, _
    ByVal summaryText As String)
    
    '------------------------------
    ' 代入
    '------------------------------
    m_EntryDate = entryDate
    
    Call m_Debit.Initialize( _
        D_AccountCode, _
        D_MajorAccount, _
        D_MiddleAccount, _
        D_SubAccountCode, _
        D_SubAccountName, _
        D_Class, _
        D_Amount)
    
    Call m_Credit.Initialize( _
        C_AccountCode, _
        C_MajorAccount, _
        C_MiddleAccount, _
        C_SubAccountCode, _
        C_SubAccountName, _
        C_Class, _
        C_Amount)

    m_Summary = summaryText
    
    '------------------------------
    ' 金額チェック
    '------------------------------
    '負数を使っていないか？
    If D_Amount < 0 Then
        Err.Raise vbObjectError + 2002, "AccountingEntry:Initialize", "借方の金額が負数です"
    End If
    
    If C_Amount < 0 Then
        Err.Raise vbObjectError + 2003, "AccountingEntry:Initialize", "貸方の金額が負数です"
    End If
    
    '貸借一致チェック
    If (D_Amount <> C_Amount) Then
        Err.Raise vbObjectError + 2004, "AccountingEntry:Initialize", "貸方と借方の金額が不一致です"
    End If

End Sub

'==============================================================================
'   メンバー変数取得
'==============================================================================
Public Property Get entryDate() As Variant
    entryDate = m_EntryDate
End Property

Public Property Get Debit() As entrySide
    Set Debit = m_Debit
End Property

Public Property Get Credit() As entrySide
    Set Credit = m_Credit
End Property

Public Property Get Summary() As String
    Summary = m_Summary
End Property

