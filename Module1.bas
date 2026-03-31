Attribute VB_Name = "Module1"
Option Explicit

'=======================================================
'       定数定義
'-------------------------------------------------------

'勘定科目コード定義
Public Const idNetAssets_End = 31000    '一般正味財産期末残高
Public Const idNetAssets_Begin = 31100  '一般正味財産期首残高
Public Const idNetAssets_Diff = 31200   '当期一般正味財産増減額
Public Const idSpNetAssets_End = 32000    '指定正味財産期末残高
Public Const idSpNetAssets_Begin = 32100  '指定正味財産期首残高
Public Const idSpNetAssets_Diff = 32200   '当期指定正味財産増減額

'勘定科目コード定義
Public Const idNetAssets = 31000    '一般正味財産期末残高
Public Const idSpNetAssets = 32000  '指定正味財産期末残高

Public Const iThisX = 6             '当年度
Public Const iLastX = 7             '前年度

'==============================================================================
'   期初日付を取得
'==============================================================================
Public Function GetFiscalYearStart() As Date
    GetFiscalYearStart = ThisWorkbook.Worksheets("設定＆使い方").Cells(14, 3).value
End Function

'==============================================================================
'   金額判定
'------------------------------------------------------------------------------
'   Input:
'       str     ソートする連想配列
'
'   Output:
'               文字だったら0、数値だったら金額
'
'==============================================================================
Public Function amount(ByVal str As Variant) As Currency
    If IsNumeric(str) And Not IsEmpty(str) Then
        amount = CCur(str) ' 数字ならそのまま
    Else
        amount = 0 ' 文字なら0
    End If
End Function

'==============================================================================
' Dictionary Key ソート（共通）
'------------------------------------------------------------------------------
'   Input:
'       dict    ソートする連想配列
'
'   Output:
'               ソート後の連想配列
'
'==============================================================================
Public Function SortDictionaryByKey(ByVal dict As Object) As Object

    Dim Keys As Variant
    Dim i As Long
    Dim newDict As Object

    Set newDict = CreateObject("Scripting.Dictionary")

    If dict Is Nothing Then
        Set SortDictionaryByKey = newDict
        Exit Function
    End If

    If dict.Count = 0 Then
        Set SortDictionaryByKey = newDict
        Exit Function
    End If

    Keys = dict.Keys

    If UBound(Keys) <= 0 Then
        newDict.Add Keys(0), dict(Keys(0))
        Set SortDictionaryByKey = newDict
        Exit Function
    End If

    Call QuickSortKeys(Keys, LBound(Keys), UBound(Keys))

    For i = LBound(Keys) To UBound(Keys)
        newDict.Add Keys(i), dict(Keys(i))
    Next

    Set SortDictionaryByKey = newDict

End Function

'==============================================================================
' クイックソート（共通）
'------------------------------------------------------------------------------
'   Input:
'       arr     ソートする配列
'       first   ソートする範囲　開始位置
'       last    ソートする範囲　終了位置
'
'==============================================================================
Public Sub QuickSortKeys(arr As Variant, ByVal first As Long, ByVal last As Long)

    Dim i As Long
    Dim j As Long
    Dim pivot As Variant
    Dim tmp As Variant
    
    i = first
    j = last
    
    pivot = arr((first + last) \ 2)
    
    Do While i <= j
    
        Do While CompareSortKeys(arr(i), pivot) < 0
            i = i + 1
        Loop
        
        Do While CompareSortKeys(arr(j), pivot) > 0
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

'------------------------------------------------------------------------------
Private Function CompareSortKeys(ByVal a As Variant, ByVal b As Variant) As Long

    Dim aStr As String
    Dim bStr As String
    
    aStr = CStr(a)
    bStr = CStr(b)

    Dim aParts() As String
    Dim bParts() As String

    If TryParseLedgerKey(aStr, aParts) And TryParseLedgerKey(bStr, bParts) Then
    
        CompareSortKeys = CompareLongValues(ToLongSafe(aParts(0)), ToLongSafe(bParts(0)))
        If CompareSortKeys <> 0 Then Exit Function

        CompareSortKeys = CompareLongValues(ToLongSafe(aParts(1)), ToLongSafe(bParts(1)))
        If CompareSortKeys <> 0 Then Exit Function

        CompareSortKeys = CompareLongValues(ToLongSafe(aParts(2)), ToLongSafe(bParts(2)))
        Exit Function
        
    End If

    CompareSortKeys = StrComp(aStr, bStr, vbTextCompare)

End Function

'------------------------------------------------------------------------------
Private Function TryParseLedgerKey(ByVal key As String, ByRef parts() As String) As Boolean

    If InStr(1, key, "|", vbBinaryCompare) = 0 Then Exit Function

    parts = Split(key, "|")
    
    If UBound(parts) <> 2 Then Exit Function
    
    If Not IsNumeric(parts(0)) Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function
    If Not IsNumeric(parts(2)) Then Exit Function

    TryParseLedgerKey = True

End Function

'------------------------------------------------------------------------------
Private Function ToLongSafe(ByVal v As String) As Long

    If IsNumeric(v) Then
        ToLongSafe = CLng(v)
    Else
        ToLongSafe = 0
    End If

End Function

'------------------------------------------------------------------------------
Private Function CompareLongValues(ByVal a As Long, ByVal b As Long) As Long

    If a < b Then
        CompareLongValues = -1
    ElseIf a > b Then
        CompareLongValues = 1
    Else
        CompareLongValues = 0
    End If

End Function

'==============================================================================
'      指定のシートをクリアします。
'------------------------------------------------------------------------------
'   Input:
'       ws  クリアするシート
'
'   Output:
'       startRow    開始行
'
'==============================================================================
Public Sub ClearSheet(ws As Worksheet, StartRow As Long)

    Dim lastRowA As Long
    Dim lastRowB As Long
    Dim lastRowC As Long
    Dim lastRow As Long
    
    With ws
    
        lastRowA = .Cells(.Rows.Count, "A").End(xlUp).Row
        lastRowB = .Cells(.Rows.Count, "B").End(xlUp).Row
        lastRowC = .Cells(.Rows.Count, "C").End(xlUp).Row
        
        '最大値を採用
        lastRow = Application.WorksheetFunction.Max(lastRowA, lastRowB, lastRowC)
        
        If lastRow >= StartRow Then
            .Rows(StartRow & ":" & lastRow).Delete
        End If
        
    End With

End Sub

'=======================================================
'       勘定科目マスターの読み込み
'-------------------------------------------------------
'   Contents:
'       勘定科目マスター
'
'   Input:
'       Sheet "(DB)勘定科目マスタ"
'
'=======================================================
Public Function LoadAccountMaster(ws As Worksheet) As AccountMaster

    Dim master As AccountMaster
    Set master = New AccountMaster

    Dim y As Long
    Dim lastRow As Long
    Dim v As Variant

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For y = 2 To lastRow
    
        v = ws.Cells(y, 1).value
        
        If Not IsEmpty(v) Then
        
            If IsNumeric(v) Then
                If CLng(v) = -1 Then Exit For
            End If
            
            Dim acc As account
            Set acc = New account
    
            acc.Initialize _
                ws.Cells(y, 1).value, _
                ws.Cells(y, 2).value, _
                ws.Cells(y, 3).value, _
                ws.Cells(y, 4).value, _
                ws.Cells(y, 5).value, _
                ws.Cells(y, 6).value, _
                ws.Cells(y, 7).value, _
                ws.Cells(y, 8).value, _
                ws.Cells(y, 9).value, _
                ws.Cells(y, 10).value, _
                ws.Cells(y, 11).value
                
            master.AddAccount acc
        End If

    Next y

    Set LoadAccountMaster = master

End Function

'=======================================================
'       集計    メインルーチン
'-------------------------------------------------------
'   Contents:
'       仕訳帳から、
'       総勘定元帳、正味財産増減計算書、貸借対照表を
'       生成します。
'
'   Input:
'       Sheet "仕訳帳"   …　仕訳帳を記入しておくこと。
'
'   Output:
'       Sheet "総勘定元帳"
'       Sheet "正味財産増減計算書"
'       Sheet "貸借対照表"
'
'=======================================================
Sub main()

    '==================================================
    'Phase [0]  初期化
    '--------------------------------------------------

    '勘定科目マスター読み込み
    Dim master As AccountMaster
    Set master = LoadAccountMaster(Sheets("(DB)勘定科目マスタ"))

    '仕訳帳読み込み用
    Dim entry As AccountingEntry
    Set entry = New AccountingEntry
    
    '前期実績の読み込み用
    Dim entrySide As entrySide
    Set entrySide = New entrySide

    '仕訳帳
    Dim db As Journal
    Set db = New Journal                    'クラス生成時に仕訳帳も読んでいる
    
    '総勘定元帳＆補助元帳
    Dim le As LedgerEngine
    Set le = New LedgerEngine

    '試算表
    Dim tb As TrialBalanceEngine
    Set tb = New TrialBalanceEngine

    Dim tbLine As TrialBalanceLine

    '財務諸表
    Dim fsReport As FinancialStatement
    Set fsReport = New FinancialStatement
    
    '廃止予定
    Dim FS As FinancialStatements
    Set FS = New FinancialStatements

    Dim it As Variant           'for each 用
    
   
    '==================================================
    'Phase [1]  前期実績 ＆ 仕訳帳 ⇒ 総勘定元帳への転記
    '--------------------------------------------------
    
    '---------------------------------------
    '[1]-(1) 「前年度実績」⇒「総勘定元帳」＆「補助元帳」に転記
    '---------------------------------------
    '① 前期実績を、試算表クラスに読み込み
    Call tb.Read_LastResult(master)
    
    '② 期初残高を総勘定元帳へ転記
    For Each tbLine In tb.LastGeneralTrialBalance.Lines
        Call le.AddOpening(tbLine, True)
    Next
    
    '③ 期初残高を補助元帳へ転記
    For Each tbLine In tb.LastSubTrialBalance.Lines
        Call le.AddOpening(tbLine, False)
    Next
    
    '---------------------------------------
    '[1]-(2) 「仕訳帳」⇒「総勘定元帳」＆「補助元帳」に転記
    '---------------------------------------
    '① 仕訳帳を、仕訳帳クラスに読み込み
    Call db.readJournal(master)
    
    '② 仕訳帳を、元帳へ転記
    For Each it In db.Items
        Set entry = it
        Call le.AddJournal(entry)
    Next it
    
    '---------------------------------------
    '[1]-(3) 「総勘定元帳」＆「補助元帳」をシートに出力
    '---------------------------------------
    Call le.Final



    '==================================================
    'Phase [2]  試算表を作成
    '--------------------------------------------------
    '① 元帳から、試算表クラスを構成
    Call tb.Build(le)
    
    '② 残高試算表として出力
    Call tb.OutputSheet
    
    ' Net asset business consistency check
    Call tb.ValidateNetAssetConsistency


    '==================================================
    'Phase [3]  財務諸表 ＆ 今期実績を作成
    '--------------------------------------------------
    Call fsReport.BuildAndOutput(tb)

    '---------------------------------------
    '[3]-(1) 資産・負債・収益・費用
    '---------------------------------------
    '   純資産の前期実績は、試算表クラスに入っています。 ←   ※■廃止予定
    For Each tbLine In tb.LastGeneralTrialBalance.Lines
        Call FS.OutFinancialStatements(tbLine.account.code, iLastX, tbLine.EndingBalance)
    Next
    
    For Each tbLine In tb.GeneralTrialBalance.Lines
        Call FS.OutFinancialStatements(tbLine.account.code, iThisX, tbLine.EndingBalance)
    Next

    '---------------------------------------
    '[3]-(2) 正味財産の部
    '---------------------------------------
    '=======================================
    ' 前期実績（iLastX） → Property Get版を使用
    '=======================================
    Call FS.OutFinancialStatements(idNetAssets_End, iLastX, tb.LastGeneralNetEnd)
    Call FS.OutFinancialStatements(idNetAssets_Begin, iLastX, tb.LastGeneralNetBegin)
    Call FS.OutFinancialStatements(idNetAssets_Diff, iLastX, tb.LastGeneralNetChange)
    Call FS.OutFinancialStatements(idSpNetAssets_End, iLastX, tb.LastDesignatedNetEnd)
    Call FS.OutFinancialStatements(idSpNetAssets_Begin, iLastX, tb.LastDesignatedNetBegin)
    Call FS.OutFinancialStatements(idSpNetAssets_Diff, iLastX, tb.LastDesignatedNetChange)
    
    '=======================================
    ' 今期実績（iThisX） → NetAssetLine 取得関数版を使用
    '=======================================
    Call FS.OutFinancialStatements(idNetAssets_End, iThisX, tb.GeneralNetAssetEnd)
    Call FS.OutFinancialStatements(idNetAssets_Begin, iThisX, tb.GeneralNetAssetBegin)
    Call FS.OutFinancialStatements(idNetAssets_Diff, iThisX, tb.GeneralNetAssetChange)
    Call FS.OutFinancialStatements(idSpNetAssets_End, iThisX, tb.DesignatedNetAssetEnd)
    Call FS.OutFinancialStatements(idSpNetAssets_Begin, iThisX, tb.DesignatedNetAssetBegin)
    Call FS.OutFinancialStatements(idSpNetAssets_Diff, iThisX, tb.DesignatedNetAssetChange)


    MsgBox ("正常終了しました")

End Sub
