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

'==============================
' 仕訳帳（Sheet2）列定義
'==============================
Private Const COL_DATE          As Long = 1

' --- 借方 ---
Private Const COL_DR_ACC_NO_I   As Long = 2
Private Const COL_DR_ACC_NO_S   As Long = 3
Private Const COL_DR_ACC_NO_S2  As Long = 4
Private Const COL_DR_SUB_NO_I   As Long = 5
Private Const COL_DR_SUB_NO_S   As Long = 6
Private Const COL_DR_CLASS      As Long = 7   ' 借方：一般/指定区分
Private Const COL_DR_MONEY      As Long = 8

' --- 貸方 ---
Private Const COL_CR_ACC_NO_I   As Long = 9
Private Const COL_CR_ACC_NO_S   As Long = 10
Private Const COL_CR_ACC_NO_S2  As Long = 11
Private Const COL_CR_SUB_NO_I   As Long = 12
Private Const COL_CR_SUB_NO_S   As Long = 13
Private Const COL_CR_CLASS      As Long = 14  ' 貸方：一般/指定区分
Private Const COL_CR_MONEY      As Long = 15

' --- 摘要 ---
Private Const COL_SUMMARY       As Long = 16

'==============================
' メンバー
'==============================
Private m_EntryDate         As Variant       ' 日付
Private m_Debit             As entrySide     ' 借方
Private m_Credit            As entrySide     ' 貸方
Private m_Summary           As String        ' 摘要

'==============================
' メンバー
'==============================
Public Sub Initialize( _
    ByVal entryDate As Date, _
    ByVal debitSide As entrySide, _
    ByVal creditSide As entrySide, _
    ByVal summaryText As String)

    If debitSide Is Nothing Then
        Err.Raise vbObjectError + 1000, , "借方未設定"
    End If
    
    If creditSide Is Nothing Then
        Err.Raise vbObjectError + 1001, , "貸方未設定"
    End If
    
    If debitSide.Amount <> creditSide.Amount Then
        Err.Raise vbObjectError + 1002, , "貸借不一致"
    End If
    
    m_EntryDate = entryDate
    Set m_Debit = debitSide
    Set m_Credit = creditSide
    m_Summary = summaryText
End Sub

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

'==============================================================================
'       仕訳帳読み込み
'------------------------------------------------------------------------------
'   Contents:
'       仕訳帳から、１行分上のメンバー変数に読み込みます。
'
'   Input:
'       y       入力Ｙ座標
'
'   Output:
'       entry   仕訳帳 １行分
'
'==============================================================================
Public Sub readJournal(y As Variant)

    'シート依存の部分は、移動
    '====================
    ' 年月日
    '====================
    m_EntryDate = Sheet2.Cells(y, COL_DATE)

    '====================
    ' 借方
    '====================
    m_Debit.AccountCode = Sheet2.Cells(y, COL_DR_ACC_NO_I)
    m_Debit.MajorAccount = Sheet2.Cells(y, COL_DR_ACC_NO_S)
    m_Debit.MiddleAccount = Sheet2.Cells(y, COL_DR_ACC_NO_S2)

    If m_Debit.MiddleAccount <> "" Then
        m_Debit.MajorAccount = m_Debit.MajorAccount & "－" & m_Debit.MiddleAccount
    End If

    m_Debit.SubAccountCode = Sheet2.Cells(y, COL_DR_SUB_NO_I)
    m_Debit.SubAccountName = Sheet2.Cells(y, COL_DR_SUB_NO_S)
    m_Debit.Amount = Sheet2.Cells(y, COL_DR_MONEY)
    
    If Trim(Sheet2.Cells(y, COL_DR_CLASS)) = "指定" Then
        m_Debit.Class = 1
    Else
        m_Debit.Class = 0
    End If

    '====================
    ' 貸方
    '====================
    m_Credit.AccountCode = Sheet2.Cells(y, COL_CR_ACC_NO_I)
    m_Credit.MajorAccount = Sheet2.Cells(y, COL_CR_ACC_NO_S)
    m_Credit.MiddleAccount = Sheet2.Cells(y, COL_CR_ACC_NO_S2)

    If m_Credit.MiddleAccount <> "" Then
         m_Credit.MajorAccount = m_Credit.MajorAccount & "－" & m_Credit.MiddleAccount
    End If
    
    m_Credit.SubAccountCode = Sheet2.Cells(y, COL_CR_SUB_NO_I)
    m_Credit.SubAccountName = Sheet2.Cells(y, COL_CR_SUB_NO_S)
    m_Credit.Amount = Sheet2.Cells(y, COL_CR_MONEY)
    
    If Trim(Sheet2.Cells(y, COL_CR_CLASS)) = "指定" Then
        m_Credit.Class = 1
    Else
        m_Credit.Class = 0
    End If

    '====================
    ' 摘要
    '====================
    m_Summary = Sheet2.Cells(y, COL_SUMMARY)

    '====================
    ' 内容チェック
    '====================
    '------------------------------
    ' 金額チェック
    '------------------------------
    '負数を使っていないか？
    If m_Debit.Amount < 0 Then
        Err.Raise vbObjectError + 2100, "AccountingEntry:readJournal", "仕訳帳" & Str(y) & "行目：借方の金額が負数です"
    End If
    
    If m_Credit.Amount < 0 Then
        Err.Raise vbObjectError + 2101, "AccountingEntry:readJournal", "仕訳帳" & Str(y) & "行目：貸方の金額が負数です"
    End If
    
    '貸借一致チェック
    If (m_Debit.Amount <> m_Credit.Amount) Then
        Err.Raise vbObjectError + 2102, "AccountingEntry:readJournal", "仕訳帳" & Str(y) & "行目：貸方と借方の金額が不一致です"
    End If
    
    '------------------------------
    ' ■To Do 補助科目チェック  これは、entrySide に移す。
    '------------------------------
    '資産じゃない場合、且つ、補助科目が未入力でないか、チェック
    If (m_Debit.AccountCode >= 20000) And (IsNull(m_Debit.SubAccountCode) Or IsEmpty(m_Debit.SubAccountCode)) Then
        Err.Raise vbObjectError + 2201, "AccountingEntry:readJournal", "仕訳帳" & Str(y) & "行目：貸方に補助科目が必要です"
    End If
    If (m_Credit.AccountCode >= 20000) And (IsNull(m_Credit.SubAccountCode) Or IsEmpty(m_Credit.SubAccountCode)) Then
        Err.Raise vbObjectError + 2202, "AccountingEntry:readJournal", "仕訳帳" & Str(y) & "行目：借方に補助科目が必要です"
    End If
    
    If (m_Debit.Class = 1) And (IsNull(m_Debit.SubAccountCode) Or IsEmpty(m_Debit.SubAccountCode)) Then
        Err.Raise vbObjectError + 2203, "AccountingEntry:readJournal", "仕訳帳" & Str(y) & "行目：（貸方）指定の場合は補助科目が必要です"
    End If
    If (m_Credit.Class = 1) And (IsNull(m_Credit.SubAccountCode) Or IsEmpty(m_Credit.SubAccountCode)) Then
        Err.Raise vbObjectError + 2204, "AccountingEntry:readJournal", "仕訳帳" & Str(y) & "行目：（借方）指定の場合は、補助科目が必要です"
    End If

End Sub
