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
Private Const idNetAssets_End = 31000       '一般正味財産期末残高
Private Const idNetAssets_Begin = 31100     '一般正味財産期首残高
Private Const idNetAssets_Diff = 31200      '当期一般正味財産増減額
Private Const idSpNetAssets_End = 32000     '指定正味財産期末残高
Private Const idSpNetAssets_Begin = 32100   '指定正味財産期首残高
Private Const idSpNetAssets_Diff = 32200    '当期指定正味財産増減額

'総勘定元帳の出力Ｙ座標
Private Const styGeneralLedger = 5          '総勘定元帳の開始位置

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

Private m_GeneralLedger As Object           '総勘定元帳
Private m_SubLedger As Object               '補助元帳

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

'===============================
' 正味財産（前期）
'===============================
Public Property Get NetAssets_End_p() As Currency
    NetAssets_End_p = m_iNetAssets_End_p
End Property

Public Property Get NetAssets_Begin_p() As Currency
    NetAssets_Begin_p = m_iNetAssets_Begin_p
End Property

Public Property Get NetAssets_Diff_p() As Currency
    NetAssets_Diff_p = m_iNetAssets_Diff_p
End Property

Public Property Get SpNetAssets_End_p() As Currency
    SpNetAssets_End_p = m_iSpNetAssets_End_p
End Property

Public Property Get SpNetAssets_Begin_p() As Currency
    SpNetAssets_Begin_p = m_iSpNetAssets_Begin_p
End Property

Public Property Get SpNetAssets_Diff_p() As Currency
    SpNetAssets_Diff_p = m_iSpNetAssets_Diff_p
End Property

'===============================
' 正味財産（当期）
'===============================
Public Property Get NetAssets_End() As Currency
    NetAssets_End = m_iNetAssets_End
End Property

Public Property Get NetAssets_Begin() As Currency
    NetAssets_Begin = m_iNetAssets_Begin
End Property

Public Property Get NetAssets_Diff() As Currency
    NetAssets_Diff = m_iNetAssets_Diff
End Property

Public Property Get SpNetAssets_End() As Currency
    SpNetAssets_End = m_iSpNetAssets_End
End Property

Public Property Get SpNetAssets_Begin() As Currency
    SpNetAssets_Begin = m_iSpNetAssets_Begin
End Property

Public Property Get SpNetAssets_Diff() As Currency
    SpNetAssets_Diff = m_iSpNetAssets_Diff
End Property

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
'   総勘定元帳の科目一覧取得
'==============================================================================
Public Function GetAccounts() As Collection

    Dim col As New Collection
    Dim key As Variant

    For Each key In m_GeneralLedger
        col.Add m_GeneralLedger(key)
    Next key

    Set GetAccounts = col

End Function


'==============================================================================
'   指定キーの元帳を取得    無ければ新規作成
'------------------------------------------------------------------------------
'   Input
'       dict            総勘定元帳 or 補助元帳の dict
'       keyStr          キー
'       entry           行
'
'   Output
'       LedgerAccount   元帳のオブジェクト
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
            entry.Account, _
            entry.SubAccountCode, _
            entry.SubAccountCate, _
            entry.SubAccountName, _
            entry.Class)
        
        dict.Add keyStr, acc
        
        Set GetOrCreateLedger = acc
        
    End If

End Function

'==============================================================================
'   前期実績1行を元帳へ転記
'------------------------------------------------------------------------------
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
'   仕訳帳1行を元帳へ転記
'------------------------------------------------------------------------------
'   Input:
'       entry   AccountingEntry
'
'==============================================================================
Public Sub AddJournal(entry As AccountingEntry)
    
    Dim d As entrySide
    Dim c As entrySide

    Dim counterAcc As Account
    Dim isMulti As Boolean

    '--------------------------------------
    ' 諸口チェック
    '--------------------------------------
    isMulti = False
    If entry.Credits.Count = 1 Then
        Set counterAcc = entry.Credits(1).Account
    Else
        Set counterAcc = Nothing
        If entry.Credits.Count > 1 Then
            isMulti = True
        End If
    End If
    
    '--------------------------------------
    ' 借方ループ
    '--------------------------------------
    For Each d In entry.Debits
    
        Dim accD As LedgerAccount
        Dim keyD As ledgerKey
        Set keyD = New ledgerKey
        
        '==============================
        ' 総勘定元帳（借方）
        '==============================
        Call keyD.Initialize(d.AccountCode, 0, 0)
        
        Set accD = GetOrCreateLedger(m_GeneralLedger, keyD.key, d)
        
        Call accD.AddLine( _
            d.Row, _
            d.ParentRow, _
            entry.entryDate, _
            counterAcc, _
            entry.summary, _
            d.amount, _
            0, _
            isMulti)
        
        If (d.AccountCode >= 40000) And (d.AccountCode < 80000) Then
            If d.Class = natDesignated Then
                '指定正味財産
                m_iSpNetAssets_Diff = m_iSpNetAssets_Diff - d.amount
            Else
                '一般正味財産
                m_iNetAssets_Diff = m_iNetAssets_Diff - d.amount
            End If
        End If
    
        '==============================
        ' 補助元帳（借方）
        '==============================
        Call keyD.Initialize( _
            d.AccountCode, _
            d.SubAccountCode, _
            d.Class)
        
        Set accD = GetOrCreateLedger(m_SubLedger, keyD.key, d)
        
        Call accD.AddLine( _
            d.Row, _
            d.ParentRow, _
            entry.entryDate, _
            counterAcc, _
            entry.summary, _
            d.amount, _
            0, _
            isMulti)
    Next
        
    '--------------------------------------
    ' 諸口チェック
    '--------------------------------------
    isMulti = False
    If entry.Debits.Count = 1 Then
        Set counterAcc = entry.Debits(1).Account
    Else
        Set counterAcc = Nothing
        If entry.Debits.Count > 1 Then
            isMulti = True
        End If
    End If
    '--------------------------------------
    ' 貸方ループ
    '--------------------------------------
    For Each c In entry.Credits
        
        Dim accC As LedgerAccount
        Dim keyC As ledgerKey
        Set keyC = New ledgerKey
        
        '==============================
        ' 総勘定元帳（貸方）
        '==============================
        Call keyC.Initialize(c.AccountCode, 0, 0)
        
        Set accC = GetOrCreateLedger(m_GeneralLedger, keyC.key, c)
        
        Call accC.AddLine( _
            c.Row, _
            c.ParentRow, _
            entry.entryDate, _
            counterAcc, _
            entry.summary, _
            0, _
            c.amount, _
            isMulti)
    
        If (c.AccountCode >= 40000) And (c.AccountCode < 80000) Then
            If c.Class = 1 Then
                '指定正味財産
                m_iSpNetAssets_Diff = m_iSpNetAssets_Diff + c.amount
            Else
                '一般正味財産
                m_iNetAssets_Diff = m_iNetAssets_Diff + c.amount
            End If
        End If
        
        '==============================
        ' 補助元帳（借方）
        '==============================
        Call keyC.Initialize( _
            c.AccountCode, _
            c.SubAccountCode, _
            c.Class)
        
        Set accC = GetOrCreateLedger(m_SubLedger, keyC.key, c)
        
        Call accC.AddLine( _
            c.Row, _
            c.ParentRow, _
            entry.entryDate, _
            counterAcc, _
            entry.summary, _
            0, _
            c.amount, _
            isMulti)

    Next
End Sub

'==============================================================================
'   期末の計算
'------------------------------------------------------------------------------
Public Sub Final()
    m_iNetAssets_End = m_iNetAssets_Begin + m_iNetAssets_Diff
    m_iSpNetAssets_End = m_iSpNetAssets_Begin + m_iSpNetAssets_Diff

    Call SortLedgers

    Call OutputSheet

End Sub

'==============================================================================
'   元帳ソート
'==============================================================================
Public Sub SortLedgers()

    Set m_GeneralLedger = SortDictionaryByKey(m_GeneralLedger)
    Set m_SubLedger = SortDictionaryByKey(m_SubLedger)

End Sub


'==============================================================================
'   Dictionary Key ソート
'==============================================================================
Private Function SortDictionaryByKey(ByVal dict As Object) As Object

    Dim keys As Variant
    Dim i As Long
    Dim newDict As Object
    
    keys = dict.keys
    
    If UBound(keys) <= 0 Then
        Set SortDictionaryByKey = dict
        Exit Function
    End If
    
    QuickSortKeys keys, LBound(keys), UBound(keys)
    
    Set newDict = CreateObject("Scripting.Dictionary")
    
    For i = LBound(keys) To UBound(keys)
        newDict.Add keys(i), dict(keys(i))
    Next
    
    Set SortDictionaryByKey = newDict

End Function

'==============================================================================
'   クイックソート
'==============================================================================
Private Sub QuickSortKeys(arr As Variant, ByVal first As Long, ByVal last As Long)

    Dim i As Long
    Dim j As Long
    Dim pivot As String
    Dim tmp As String
    
    i = first
    j = last
    
    pivot = arr((first + last) \ 2)
    
    Do While i <= j
    
        Do While arr(i) < pivot
            i = i + 1
        Loop
        
        Do While arr(j) > pivot
            j = j - 1
        Loop
        
        If i <= j Then
        
            tmp = arr(i)
            arr(i) = arr(j)
            arr(j) = tmp
            
            i = i + 1
            j = j - 1
            
        End If
        
    Loop
    
    If first < j Then QuickSortKeys arr, first, j
    If i < last Then QuickSortKeys arr, i, last

End Sub

'==============================================================================
'   総勘定元帳 ＆ 補助元帳出力
'==============================================================================
Public Sub OutputSheet()

    Dim MaxColumn As Long
    MaxColumn = 9

    '----------------------------------
    '   変数宣言
    '----------------------------------
    Dim wsGL As Worksheet
    Dim wsSL As Worksheet
    Dim key As Variant
    Dim acc As LedgerAccount
    Dim line As LedgerLine
    Dim y As Long
    
    '----------------------------------
    '   初期化
    '----------------------------------
    Set wsGL = ThisWorkbook.Worksheets("②総勘定元帳")
    Set wsSL = ThisWorkbook.Worksheets("③補助元帳")
    
    
    '----------------------------------
    '   シートクリア
    '----------------------------------
    Call ClearSheet(wsGL, styGeneralLedger)
    Call ClearSheet(wsSL, styGeneralLedger)

    '----------------------------------
    '   総勘定元帳
    '----------------------------------
    y = styGeneralLedger
    With wsGL
    
         .Columns("A").NumberFormat = "yyyy/mm/dd"
    
        '数値表示形式（小数点なし）
        .Range("E:F, H:H").NumberFormat = "#,##0;△#,##0;;@"
        .Range("E:F, H:H").HorizontalAlignment = xlRight
        .Range("G:G").HorizontalAlignment = xlCenter
        .Range("I:I").HorizontalAlignment = xlCenter
    
        For Each key In m_GeneralLedger
            Set acc = m_GeneralLedger.Item(key)
            If (acc.Lines.Count > 0) And ((acc.AccountCode < 30000) Or (acc.AccountCode >= 40000)) Then
                y = y + 1
                
                '--------------------------
                '勘定科目の出力
                .Range(.Cells(y, 1), .Cells(y, 8)).Merge
                If acc.MiddleAccount <> "" Then
                    .Cells(y, 1) = "科目：" & acc.MajorAccount & " － " & acc.MiddleAccount
                Else
                    .Cells(y, 1) = "科目：" & acc.MajorAccount
                End If
                .Cells(y, 1).Font.Underline = True
                .Cells(y, 1).Font.size = 16
                .Cells(y, 1).HorizontalAlignment = xlCenter
                y = y + 1
            
                '--------------------------
                '勘定科目コードの出力
                .Cells(y, 1) = "勘定科目コード：" & CStr(acc.AccountCode)
                y = y + 1
                
                '--------------------------
                '項目行
                .Cells(y, 1) = "日付"
                .Cells(y, 2) = "相手大科目"
                .Cells(y, 3) = "相手中科目"
                .Cells(y, 4) = "摘要"
                .Cells(y, 5) = "借方"
                .Cells(y, 6) = "貸方"
                .Cells(y, 7) = "貸／借"
                .Cells(y, 8) = "残高"
                .Cells(y, 9) = "仕訳帳行"
                
                .Range(.Cells(y, 1), .Cells(y, MaxColumn)).Font.Bold = True
                .Range(.Cells(y, 1), .Cells(y, MaxColumn)).HorizontalAlignment = xlCenter
                
                .Range(.Cells(y, 1), .Cells(y, MaxColumn)).Borders.LineStyle = xlContinuous
                .Range(.Cells(y, 1), .Cells(y, MaxColumn)).Borders(xlEdgeTop).LineStyle = xlDouble
                .Range(.Cells(y, 1), .Cells(y, MaxColumn)).Borders(xlEdgeBottom).Weight = xlMedium
                y = y + 1
                
                '--------------------------
                '行
                For Each line In acc.Lines
                    .Cells(y, 1).Resize(1, MaxColumn).value = Array( _
                        line.entryDate, _
                        line.counterMajorAccount, _
                        line.counterMiddleAccount, _
                        line.summary, _
                        IIf(line.Debit = 0, "", line.Debit), _
                        IIf(line.Credit = 0, "", line.Credit), _
                        line.DebitCreditLabel, _
                        line.BalanceAbs, _
                        IIf(line.Row = 0, "", "=HYPERLINK(""#'①仕訳帳'!A" & line.Row & """,""" & line.Row & """)"))
                    .Range(.Cells(y, 1), .Cells(y, MaxColumn)).Borders.LineStyle = xlContinuous
                    y = y + 1
                Next line
                y = y + 1
            End If
        Next key
    End With
    
    '----------------------------------
    '   補助元帳
    '----------------------------------
    y = styGeneralLedger
    With wsSL
         .Columns("A").NumberFormat = "yyyy/mm/dd"
    
        '数値表示形式（小数点なし）
        .Range("E:F, H:H").NumberFormat = "#,##0;△#,##0;;@"
        .Range("E:F, H:H").HorizontalAlignment = xlRight
        .Range("G:G").HorizontalAlignment = xlCenter
        .Range("I:I").HorizontalAlignment = xlCenter
        
        For Each key In m_SubLedger
            Set acc = m_SubLedger.Item(key)
            If (acc.Lines.Count > 0) And ((acc.AccountCode < 30000) Or (acc.AccountCode >= 40000)) Then
                y = y + 1
                
                '--------------------------
                '勘定科目の出力
                .Range(.Cells(y, 1), .Cells(y, 8)).Merge
                If acc.MiddleAccount <> "" Then
                    .Cells(y, 1) = "科目：" & acc.MajorAccount & " － " & acc.MiddleAccount
                Else
                    .Cells(y, 1) = "科目：" & acc.MajorAccount
                End If
                .Cells(y, 1).Font.Underline = True
                .Cells(y, 1).Font.size = 16
                .Cells(y, 1).HorizontalAlignment = xlCenter
                y = y + 1
            
                '--------------------------
                '補助勘定科目の出力
                .Range(.Cells(y, 1), .Cells(y, 8)).Merge
                .Cells(y, 1) = "補助科目：［" & acc.SubAccountCate & "］" & acc.SubAccountName & acc.Class
                .Cells(y, 1).Font.Underline = True
                .Cells(y, 1).Font.size = 14
                .Cells(y, 1).HorizontalAlignment = xlCenter
                y = y + 2
                
                '--------------------------
                '勘定科目コードの出力
                .Cells(y, 1) = "勘定科目コード：" & CStr(acc.AccountCode)
                .Cells(y, 3) = "補助科目コード：" & CStr(acc.SubAccountCode)
                y = y + 1
                
                '--------------------------
                '項目行
                .Cells(y, 1) = "日付"
                .Cells(y, 2) = "相手大科目"
                .Cells(y, 3) = "相手中科目"
                .Cells(y, 4) = "摘要"
                .Cells(y, 5) = "借方"
                .Cells(y, 6) = "貸方"
                .Cells(y, 7) = "貸／借"
                .Cells(y, 8) = "残高"
                .Cells(y, 9) = "仕訳帳行"
                
                .Range(.Cells(y, 1), .Cells(y, MaxColumn)).Font.Bold = True
                .Range(.Cells(y, 1), .Cells(y, MaxColumn)).HorizontalAlignment = xlCenter
                
                .Range(.Cells(y, 1), .Cells(y, MaxColumn)).Borders.LineStyle = xlContinuous
                .Range(.Cells(y, 1), .Cells(y, MaxColumn)).Borders(xlEdgeTop).LineStyle = xlDouble
                .Range(.Cells(y, 1), .Cells(y, MaxColumn)).Borders(xlEdgeBottom).Weight = xlMedium
                y = y + 1
                
                '--------------------------
                '行
                For Each line In acc.Lines
                    .Cells(y, 1).Resize(1, MaxColumn).value = Array( _
                        line.entryDate, _
                        line.counterMajorAccount, _
                        line.counterMiddleAccount, _
                        line.summary, _
                        IIf(line.Debit = 0, "", line.Debit), _
                        IIf(line.Credit = 0, "", line.Credit), _
                        line.DebitCreditLabel, _
                        line.BalanceAbs, _
                        IIf(line.Row = 0, "", "=HYPERLINK(""#'①仕訳帳'!A" & line.Row & """,""" & line.Row & """)"))
                    .Range(.Cells(y, 1), .Cells(y, MaxColumn)).Borders.LineStyle = xlContinuous
                    y = y + 1
                Next line
                y = y + 1
            End If
        Next key
    End With

End Sub

'==============================================================================
'      指定のシートをクリアします。
'------------------------------------------------------------------------------
'   Input:
'       ws  クリアするシート
'   Output:
'       startRow    開始行
'==============================================================================
Private Sub ClearSheet(ws As Worksheet, StartRow As Long)

    '----------------------------------
    '   変数宣言
    '----------------------------------
    Dim lastRow As Long
    
    '----------------------------------
    '   シートクリア
    '----------------------------------
    With ws
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRow >= StartRow Then
            .Rows(StartRow & ":" & lastRow).Delete
        End If
    End With

End Sub

