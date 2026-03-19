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

'実績の読み込みＹ座標
Public Const styPerformance = 7     '実績の開始位置

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
        
            Dim acc As Account
            Set acc = New Account
    
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
'       集計
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
    Call db.readJournal(master)             '仕訳帳チェック
    
    '総勘定元帳＆補助元帳
    Dim gl As Ledger
    Set gl = New Ledger
    
    Dim le As LedgerEngine
    Set le = New LedgerEngine

    '財務諸表
    Dim FS As FinancialStatements
    Set FS = New FinancialStatements

    Dim it As Variant           'for each 用
    Dim key As Variant          'for each 用
   
    Dim yInput As Long
    
    '==================================================
    'Phase [1]  前期実績 ＆ 仕訳帳 ⇒ 総勘定元帳への転記
    '--------------------------------------------------
    
    '---------------------------------------
    '[1]-(1) 「前年度実績」⇒「総勘定元帳」＆「補助元帳」に転記
    '---------------------------------------
    yInput = styPerformance
    Do
        '前期実績を1行読み込み
        Call entrySide.ReadResults(yInput, master)
        
        '勘定科目コードの記載が無かったら、検索終了
        If entrySide.AccountCode = -1 Then Exit Do
        
        Call le.AddOpening(entrySide)
        
        '金額が0でない場合
        If entrySide.amount <> 0 Then
            '副科目コードに何も書いていなければ
            If Not entrySide.IsSubAcc Then
                                                              
                '財務諸表（前年度）への出力
                Call FS.OutFinancialStatements(entrySide.AccountCode, iLastX, entrySide.amount)

            End If
        End If
        
        '次の行へ
        yInput = yInput + 1
    Loop
    
    '---------------------------------------
    '[1]-(2) 「仕訳帳」⇒「総勘定元帳」＆「補助元帳」に転記
    '---------------------------------------
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

    ' ■ To Do  試算表



    '==================================================
    'Phase [3]  財務諸表 ＆ 今期実績を作成
    '--------------------------------------------------

    '数値表示形式（小数点なし）
    Sheet11.Range("H:K").NumberFormat = "#,##0;△#,##0;;@"
    Sheet11.Range("H:K").HorizontalAlignment = xlRight

    '---------------------------------------
    '[3]-(1) 資産・負債・収益・費用
    '---------------------------------------

    '今期実績への出力ｙ座標
    Dim yOutPerformance
    Dim acc As LedgerAccount
    
    '今期実績への出力ｙ座標
    yOutPerformance = styPerformance

    For Each key In le.GeneralLedger
        Set acc = le.GeneralLedger.Item(key)
        
        Call FS.OutFinancialStatements(acc.AccountCode, iThisX, acc.EndingBalance)
        '勘定科目は出力する。
        Sheet11.Cells(yOutPerformance, 1) = acc.AccountCode
        Sheet11.Cells(yOutPerformance, 2) = acc.MajorAccount
        Sheet11.Cells(yOutPerformance, 3) = acc.MiddleAccount
        Sheet11.Cells(yOutPerformance, 4) = ""
        Sheet11.Cells(yOutPerformance, 5) = ""
        Sheet11.Cells(yOutPerformance, 6) = ""
        Sheet11.Cells(yOutPerformance, 7) = ""              'ここは、一般も指定も無い。
        If acc.AccountCode < 40000 Then
            Sheet11.Cells(yOutPerformance, 8) = acc.BeginningBalance
        End If
        Sheet11.Cells(yOutPerformance, 9) = acc.Debit
        Sheet11.Cells(yOutPerformance, 10) = acc.Credit
        Sheet11.Cells(yOutPerformance, 11) = acc.EndingBalance
       
        yOutPerformance = yOutPerformance + 1

    Next key

    For Each key In le.SubLedger
        Set acc = le.SubLedger.Item(key)
        
        '勘定科目は出力する。
        Sheet11.Cells(yOutPerformance, 1) = acc.AccountCode
        Sheet11.Cells(yOutPerformance, 2) = acc.MajorAccount
        Sheet11.Cells(yOutPerformance, 3) = acc.MiddleAccount
        Sheet11.Cells(yOutPerformance, 4) = acc.SubAccountCode
        Sheet11.Cells(yOutPerformance, 5) = acc.SubAccountCate
        Sheet11.Cells(yOutPerformance, 6) = acc.SubAccountName
        Sheet11.Cells(yOutPerformance, 7) = acc.Class              'ここは、一般も指定も無い。
        If acc.AccountCode < 40000 Then
            Sheet11.Cells(yOutPerformance, 8) = acc.BeginningBalance
        End If
        Sheet11.Cells(yOutPerformance, 9) = acc.Debit
        Sheet11.Cells(yOutPerformance, 10) = acc.Credit
        Sheet11.Cells(yOutPerformance, 11) = acc.EndingBalance
        
        yOutPerformance = yOutPerformance + 1

    Next key

    '---------------------------------------
    '[3]-(2) 正味財産の部
    '---------------------------------------
    
    Call FS.OutFinancialStatements(idNetAssets_End, iThisX, le.NetAssets_End)
    Call FS.OutFinancialStatements(idNetAssets_Begin, iThisX, le.NetAssets_Begin)
    Call FS.OutFinancialStatements(idNetAssets_Diff, iThisX, le.NetAssets_Diff)
    Call FS.OutFinancialStatements(idSpNetAssets_End, iThisX, le.SpNetAssets_End)
    Call FS.OutFinancialStatements(idSpNetAssets_Begin, iThisX, le.SpNetAssets_Begin)
    Call FS.OutFinancialStatements(idSpNetAssets_Diff, iThisX, le.SpNetAssets_Diff)

    '勘定科目　集計用
    '   1次 科目
    '   2次
    '       0   勘定科目コード
    '       1   前年度 繰越金
    '       2   借方
    '       3   貸方
    '       4   残高
    '       5   補助簿の配列先頭
    '       6   補助簿の配列終了
    '       11  勘定科目名（文字列データ）
    
    '今期実績への出力

    
    Sheet11.Cells(yOutPerformance, 1) = idNetAssets_Diff
    Sheet11.Cells(yOutPerformance, 2) = "当期一般正味財産増減額"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = ""
    Sheet11.Cells(yOutPerformance, 6) = ""
    Sheet11.Cells(yOutPerformance, 11) = le.NetAssets_Diff
    yOutPerformance = yOutPerformance + 1
    
    Sheet11.Cells(yOutPerformance, 1) = idNetAssets_Begin
    Sheet11.Cells(yOutPerformance, 2) = "一般正味財産期首残高"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = ""
    Sheet11.Cells(yOutPerformance, 6) = ""
    Sheet11.Cells(yOutPerformance, 11) = le.NetAssets_Begin
    yOutPerformance = yOutPerformance + 1
            
    Sheet11.Cells(yOutPerformance, 1) = idNetAssets_End
    Sheet11.Cells(yOutPerformance, 2) = "一般正味財産期末残高"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = ""
    Sheet11.Cells(yOutPerformance, 6) = ""
    Sheet11.Cells(yOutPerformance, 11) = le.NetAssets_End
    yOutPerformance = yOutPerformance + 1

    Sheet11.Cells(yOutPerformance, 1) = idSpNetAssets_Diff
    Sheet11.Cells(yOutPerformance, 2) = "当期指定正味財産増減額"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = ""
    Sheet11.Cells(yOutPerformance, 6) = ""
    Sheet11.Cells(yOutPerformance, 11) = le.SpNetAssets_Diff
    yOutPerformance = yOutPerformance + 1
    
    Sheet11.Cells(yOutPerformance, 1) = idSpNetAssets_Begin
    Sheet11.Cells(yOutPerformance, 2) = "指定正味財産期首残高"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = ""
    Sheet11.Cells(yOutPerformance, 6) = ""
    Sheet11.Cells(yOutPerformance, 11) = le.SpNetAssets_Begin
    yOutPerformance = yOutPerformance + 1
            
    Sheet11.Cells(yOutPerformance, 1) = idSpNetAssets_End
    Sheet11.Cells(yOutPerformance, 2) = "指定正味財産期末残高"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = ""
    Sheet11.Cells(yOutPerformance, 6) = ""
    Sheet11.Cells(yOutPerformance, 11) = le.SpNetAssets_End
    yOutPerformance = yOutPerformance + 1

    MsgBox ("正常終了しました")

End Sub
