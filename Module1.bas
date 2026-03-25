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

Public Const iThisX = 6             '当年度
Public Const iLastX = 7             '前年度

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
Public Function Amount(ByVal str As Variant) As Currency
    
    If IsNumeric(str) And Not IsEmpty(str) Then
        Amount = CCur(str) ' 数字ならそのまま
    Else
        Amount = 0 ' 文字なら0
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

    Dim keys As Variant
    Dim i As Long
    Dim newDict As Object
    
    keys = dict.keys
    
    If UBound(keys) <= 0 Then
        Set SortDictionaryByKey = dict
        Exit Function
    End If
    
    Call QuickSortKeys(keys, LBound(keys), UBound(keys))
    
    Set newDict = CreateObject("Scripting.Dictionary")
    
    For i = LBound(keys) To UBound(keys)
        newDict.Add keys(i), dict(keys(i))
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
    y = 2

    Do
        If ws.Cells(y, 1).value <> "" Then
        
            Dim acc As account
            Set acc = New account
    
            acc.Initialize _
                ws.Cells(y, 1), _
                ws.Cells(y, 2), _
                ws.Cells(y, 3), _
                ws.Cells(y, 4), _
                ws.Cells(y, 5), _
                ws.Cells(y, 6), _
                ws.Cells(y, 7), _
                ws.Cells(y, 8), _
                ws.Cells(y, 9), _
                ws.Cells(y, 10)
    
            master.AddAccount acc
        End If
        If ws.Cells(y, 1).value = -1 Then Exit Do
        y = y + 1

    Loop

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
    Dim db As journal
    Set db = New journal                    'クラス生成時に仕訳帳も読んでいる
    
    '総勘定元帳＆補助元帳
    Dim le As LedgerEngine
    Set le = New LedgerEngine

    '試算表
    Dim tb As TrialBalanceEngine
    Set tb = New TrialBalanceEngine

    Dim tbLine As TrialBalanceLine

    '財務諸表
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



    '==================================================
    'Phase [3]  財務諸表 ＆ 今期実績を作成
    '--------------------------------------------------

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
    Call FS.OutFinancialStatements(idNetAssets_End, iLastX, tb.LastGeneralTrialBalance.GeneralNetEnd)
    Call FS.OutFinancialStatements(idNetAssets_Begin, iLastX, tb.LastGeneralTrialBalance.GeneralNetBegin)
    Call FS.OutFinancialStatements(idNetAssets_Diff, iLastX, tb.LastGeneralTrialBalance.GeneralNetChange)
    Call FS.OutFinancialStatements(idSpNetAssets_End, iLastX, tb.LastGeneralTrialBalance.DesignatedNetEnd)
    Call FS.OutFinancialStatements(idSpNetAssets_Begin, iLastX, tb.LastGeneralTrialBalance.DesignatedNetBegin)
    Call FS.OutFinancialStatements(idSpNetAssets_Diff, iLastX, tb.LastGeneralTrialBalance.DesignatedNetChange)

    
    Call FS.OutFinancialStatements(idNetAssets_End, iThisX, tb.GeneralTrialBalance.GeneralNetEnd)
    Call FS.OutFinancialStatements(idNetAssets_Begin, iThisX, tb.GeneralTrialBalance.GeneralNetBegin)
    Call FS.OutFinancialStatements(idNetAssets_Diff, iThisX, tb.GeneralTrialBalance.GeneralNetChange)
    Call FS.OutFinancialStatements(idSpNetAssets_End, iThisX, tb.GeneralTrialBalance.DesignatedNetEnd)
    Call FS.OutFinancialStatements(idSpNetAssets_Begin, iThisX, tb.GeneralTrialBalance.DesignatedNetBegin)
    Call FS.OutFinancialStatements(idSpNetAssets_Diff, iThisX, tb.GeneralTrialBalance.DesignatedNetChange)


    MsgBox ("正常終了しました")

End Sub
