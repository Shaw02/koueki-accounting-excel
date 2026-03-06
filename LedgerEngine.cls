VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LedgerEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'   元帳エンジン
'------------------------------------------------------------------------------
'   Contents:
'       仕訳帳から
'       ・総勘定元帳
'       ・補助元帳
'       を生成する
'
'       LedgerAccount を管理する
'
'==============================================================================
Option Explicit

'勘定科目コード定義
Private Const idNetAssets_End = 31000      '一般正味財産期末残高
Private Const idNetAssets_Begin = 31100    '一般正味財産期首残高
Private Const idNetAssets_Diff = 31200     '当期一般正味財産増減額
Private Const idSpNetAssets_End = 32000    '指定正味財産期末残高
Private Const idSpNetAssets_Begin = 32100  '指定正味財産期首残高
Private Const idSpNetAssets_Diff = 32200   '当期指定正味財産増減額

'-------------------------------
' 正味財産（前期）
'-------------------------------
Private m_iNetAssets_End_p   As Currency   '一般正味財産期末残高
Private m_iNetAssets_Begin_p As Currency   '一般正味財産期首残高
Private m_iNetAssets_Diff_p  As Currency   '当期一般正味財産増減額
    
Private m_iSpNetAssets_End_p   As Currency '指定正味財産期末残高
Private m_iSpNetAssets_Begin_p As Currency '指定正味財産期首残高
Private m_iSpNetAssets_Diff_p  As Currency '当期指定正味財産増減額
    
'-------------------------------
' 正味財産（当期）
'-------------------------------
Private m_iNetAssets_End   As Currency     '一般正味財産期末残高
Private m_iNetAssets_Begin As Currency     '一般正味財産期首残高
Private m_iNetAssets_Diff  As Currency     '当期一般正味財産増減額
    
Private m_iSpNetAssets_End   As Currency   '指定正味財産期末残高
Private m_iSpNetAssets_Begin As Currency   '指定正味財産期首残高
Private m_iSpNetAssets_Diff  As Currency   '当期指定正味財産増減額

Private m_GeneralLedger As Object
Private m_SubLedger As Object


'==============================================================================
'   初期化
'==============================================================================
Private Sub Class_Initialize()

    Set m_GeneralLedger = CreateObject("Scripting.Dictionary")
    Set m_SubLedger = CreateObject("Scripting.Dictionary")
    
    '正味財産の集計用
    m_iNetAssets_Begin = 0                    '一般正味財産期首残高
    m_iNetAssets_End = 0                      '一般正味財産期末残高
    m_iNetAssets_Diff = 0                     '当期一般正味財産増減額
    
    m_iSpNetAssets_Begin = 0                  '指定正味財産期首残高
    m_iSpNetAssets_End = 0                    '指定正味財産期末残高
    m_iSpNetAssets_Diff = 0                   '当期指定正味財産増減額
    
End Sub


'==============================================================================
'   元帳取得
'------------------------------------------------------------------------------
'   指定キーの元帳を取得
'   無ければ新規作成
'
'==============================================================================
Private Function GetOrCreateLedger( _
    ByRef dict As Object, _
    ByVal keyStr As String, _
    ByVal entry As entrySide) As LedgerAccount

    If dict.Exists(keyStr) Then
        
        Set GetOrCreateLedger = dict(keyStr)
        
    Else
        
        Dim acc As LedgerAccount
        Set acc = New LedgerAccount
        
        Call acc.Initialize( _
            entry.AccountCode, _
            entry.MajorAccount, _
            entry.MiddleAccount, _
            entry.SubAccountCode, _
            entry.SubAccountName, _
            entry.Class)
        
        dict.Add keyStr, acc
        
        Set GetOrCreateLedger = acc
        
    End If

End Function

'==============================================================================
'   前期繰越追加
'------------------------------------------------------------------------------
'   Contents:
'       前期実績1行を元帳へ転記
'
'   Input:
'       entry   entrySide
'
'==============================================================================
Public Sub AddOpening(entry As entrySide)

    Dim key As ledgerKey
    Set key = New ledgerKey
    
    key.Initialize _
        entry.AccountCode, _
        entry.SubAccountCode, _
        entry.Class
    
    Dim acc As LedgerAccount
    If (IsNull(entry.SubAccountCode) Or IsEmpty(entry.SubAccountCode) Or entry.SubAccountCode = 0) And (entry.Class = 0) Then
        
        '前年度の増減額・期首残高・期末残高か？
        Select Case entry.AccountCode
            Case idNetAssets_End
                m_iNetAssets_End_p = entry.amount
                m_iNetAssets_Begin = m_iNetAssets_End_p         '昨年期末残高は、今期期初残高
            Case idNetAssets_Begin
                m_iNetAssets_Begin_p = entry.amount
            Case idNetAssets_Diff
                m_iNetAssets_Diff_p = entry.amount
            Case idSpNetAssets_End
                m_iSpNetAssets_End_p = entry.amount
                m_iSpNetAssets_Begin = m_iSpNetAssets_End_p     '昨年期末残高は、今期期初残高
            Case idSpNetAssets_Begin
                m_iSpNetAssets_Begin_p = entry.amount
            Case idSpNetAssets_Diff
                m_iSpNetAssets_Diff_p = entry.amount
            Case Else
                '上以外は、リスト化する
                Set acc = GetOrCreateLedger(m_GeneralLedger, key.key, entry)
                acc.AddOpeningBalance (entry.amount)
        End Select
    Else
        Set acc = GetOrCreateLedger(m_SubLedger, key.key, entry)
        acc.AddOpeningBalance (entry.amount)
    End If

End Sub

'==============================================================================
'   期末の計算
'------------------------------------------------------------------------------
Public Sub Final()
    m_iNetAssets_End = m_iNetAssets_Begin + m_iNetAssets_Diff
    m_iSpNetAssets_End = m_iSpNetAssets_Begin + m_iSpNetAssets_Diff

    Call test

End Sub

'==============================================================================
'   仕訳追加
'------------------------------------------------------------------------------
'   Contents:
'       仕訳帳1行を元帳へ転記
'
'   Input:
'       entry   AccountingEntry
'
'==============================================================================
Public Sub AddJournal(entry As AccountingEntry)

    '------------------------------
    '補助元帳（借方）
    '------------------------------
    
    Dim keyD As ledgerKey
    Set keyD = New ledgerKey
    
    Call keyD.Initialize( _
        entry.Debit.AccountCode, _
        entry.Debit.SubAccountCode, _
        entry.Debit.Class)
    
    Dim accD As LedgerAccount
    Set accD = GetOrCreateLedger(m_SubLedger, keyD.key, entry.Debit)
    
    Call accD.AddLine( _
        entry.entryDate, _
        entry.Credit.MajorAccount, _
        entry.Credit.MiddleAccount, _
        entry.summary, _
        entry.Debit.amount, _
        0)
    
    
    '------------------------------
    '補助元帳（貸方）
    '------------------------------
    
    Dim keyC As ledgerKey
    Set keyC = New ledgerKey
    
    Call keyC.Initialize( _
        entry.Credit.AccountCode, _
        entry.Credit.SubAccountCode, _
        entry.Credit.Class)
    
    Dim accC As LedgerAccount
    Set accC = GetOrCreateLedger(m_SubLedger, keyC.key, entry.Credit)

    Call accC.AddLine( _
        entry.entryDate, _
        entry.Debit.MajorAccount, _
        entry.Debit.MiddleAccount, _
        entry.summary, _
        0, _
        entry.Credit.amount)
    
    
    '------------------------------
    '総勘定元帳（借方）
    '------------------------------
    
    Call keyD.Initialize(entry.Debit.AccountCode, 0, 0)
    
    Set accD = GetOrCreateLedger(m_GeneralLedger, keyD.key, entry.Debit)
    
    Call accD.AddLine( _
        entry.entryDate, _
        entry.Credit.MajorAccount, _
        entry.Credit.MiddleAccount, _
        entry.summary, _
        entry.Debit.amount, _
        0)
    
    If (entry.Debit.AccountCode >= 40000) And (entry.Debit.AccountCode < 80000) Then
        If entry.Debit.Class = 1 Then
            '指定正味財産
            m_iSpNetAssets_Diff = m_iSpNetAssets_Diff - entry.Debit.amount
        Else
            '一般正味財産
            m_iNetAssets_Diff = m_iNetAssets_Diff - entry.Debit.amount
        End If
    End If

    '------------------------------
    '総勘定元帳（貸方）
    '------------------------------
    
    Call keyC.Initialize(entry.Credit.AccountCode, 0, 0)
    
    Set accC = GetOrCreateLedger(m_GeneralLedger, keyC.key, entry.Credit)
    
    Call accC.AddLine( _
        entry.entryDate, _
        entry.Debit.MajorAccount, _
        entry.Debit.MiddleAccount, _
        entry.summary, _
        0, _
        entry.Credit.amount)

    If (entry.Credit.AccountCode >= 40000) And (entry.Credit.AccountCode < 80000) Then
        If entry.Credit.Class = 1 Then
            '指定正味財産
            m_iSpNetAssets_Diff = m_iSpNetAssets_Diff + entry.Credit.amount
        Else
            '一般正味財産
            m_iNetAssets_Diff = m_iNetAssets_Diff + entry.Credit.amount
        End If
    End If

End Sub


'==============================================================================
'   総勘定元帳取得
'==============================================================================

Public Property Get GeneralLedger() As Object

    Set GeneralLedger = m_GeneralLedger

End Property


'==============================================================================
'   補助元帳取得
'==============================================================================

Public Property Get SubLedger() As Object

    Set SubLedger = m_SubLedger

End Property


'==============================================================================
'   テスト　総勘定元帳
'==============================================================================
Public Sub test()

    '----------------------------------
    '   変数宣言
    '----------------------------------
    Dim ws As Worksheet
    Dim key As Variant
    Dim acc As LedgerAccount
    Dim line As LedgerLine
    Dim y As Long
    
    
    '----------------------------------
    '   初期化
    '----------------------------------
    Set ws = ThisWorkbook.Worksheets("TEST")
    
    '----------------------------------
    '   シートクリア
    '----------------------------------
    Call ClearSheet(ws, 1)

'Private m_GeneralLedger As Object
'Private m_SubLedger As Object
    
    For Each key In m_GeneralLedger
        Set acc = m_GeneralLedger.Item(key)
        Debug.Print "科目：" & str(acc.AccountCode) & acc.MajorAccount & "-" & acc.MiddleAccount
        
        For Each line In acc.Lines
            Debug.Print line.entryDate & "|" & line.counterMajorAccount & "-" & line.counterMiddleAccount & "|" & line.summary & "|" & line.Debit & "|" & line.Credit & "|" & line.balance
        Next line
    Next key
    
    For Each key In m_SubLedger
        Set acc = m_SubLedger.Item(key)
        Debug.Print "科目：" & str(acc.AccountCode) & acc.MajorAccount & "-" & acc.MiddleAccount
        Debug.Print "補助科目：" & acc.SubAccountName & acc.Class
       
        For Each line In acc.Lines
            Debug.Print line.entryDate & "|" & line.counterMajorAccount & "-" & line.counterMiddleAccount & "|" & line.summary & "|" & line.Debit & "|" & line.Credit & "|" & line.balance
        Next line
    Next key
    

End Sub

'==============================================================================
'      シートのクリア
'------------------------------------------------------------------------------
'   Contents:
'       指定のシートをクリアします。
'   Input:
'       ws  クリアするシート
'   Output:
'       startRow    開始行
'==============================================================================
Private Sub ClearSheet(ws As Worksheet, startRow As Long)

    '----------------------------------
    '   変数宣言
    '----------------------------------
    Dim lastRow As Long
    
    '----------------------------------
    '   シートクリア
    '----------------------------------
    With ws
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRow >= startRow Then
            .Rows(startRow & ":" & lastRow).Delete
        End If
    End With

End Sub

