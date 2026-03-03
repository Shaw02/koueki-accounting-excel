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

'総勘定元帳の出力Ｙ座標
Public Const styGeneralLedger = 5   '総勘定元帳の開始位置

'実績の読み込みＹ座標
Public Const styPerformance = 7     '実績の開始位置

'=======================================================
'       クイックソート
'-------------------------------------------------------
'   Contents:
'       クイックソートします。
'
'   Input:
'       Data()  ソートする配列（２次元）の参照
'       n       配列の列数
'       key     何列目をソートするか？
'       low     どっから
'       high    どこまで
'
'   Output:
'       Data()  ソート後
'
'=======================================================
Sub QuickSort(ByRef Data() As Variant, ByVal n As Variant, ByVal Key As Long, ByVal low As Long, ByVal high As Long)

    Dim i As Long
    Dim l As Long
    Dim r As Long
    l = low
    r = high

    Dim pivot As Variant
    pivot = Data((low + high) \ 2, Key)

    Dim temp As Variant
    
    Do While (l <= r)
        Do While (Data(l, Key) < pivot And l < high)
            l = l + 1
        Loop
        Do While (pivot < Data(r, Key) And r > low)
            r = r - 1
        Loop
    
        If (l <= r) Then
            For i = 0 To n Step 1
                temp = Data(l, i)
                Data(l, i) = Data(r, i)
                Data(r, i) = temp
            Next
            l = l + 1
            r = r - 1
        End If
    Loop
    
    If (low < r) Then
        Call QuickSort(Data, n, Key, low, r)
    End If
    If (l < high) Then
        Call QuickSort(Data, n, Key, l, high)
    End If

End Sub
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
    
    '--------------------------
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
    '       7   勘定科目名（文字列データ）大科目
    '       8   勘定科目名（文字列データ）中科目
    '       9   一般の残高
    '       10  指定の残高
    Dim Account(99999, 11) As Variant
    
    '--------------------------
    '補助科目　集計用
    '   1次 科目
    '   2次
    '       0   補助科目コード
    '       1   前年度 繰越金
    '       2   借方
    '       3   貸方
    '       4   残高
    '       5   補助科目名（文字列データ）
    '       6   一般／指定
    Dim SubAccount(99999, 7) As Variant
   
    '--------------------------
    '初期化
    
    '仕訳帳読み込み用
    Dim entry As AccountingEntry
    Set entry = New AccountingEntry
    
    '前期実績の読み込み用
    Dim entrySide As entrySide
    Set entrySide = New entrySide

    '仕訳帳
    Dim db As journal
    Set db = New journal

    Call db.readJournal

    '総勘定元帳＆補助元帳
    Dim gl As Ledger
    Set gl = New Ledger

    '財務諸表
    Dim fs As FinancialStatements
    Set fs = New FinancialStatements

    '--------------------------
    '科目
    Dim cntAccount As Long      '勘定科目数
    Dim cntSubAccount As Long   '勘定科目数
    Dim stCntSubAccount As Long
    Dim EndCntSubAccount As Long
    Dim iSub As Long
    
    Dim yInput As Long
    Dim i As Long
    
    Dim it As Variant           'for each 用
    
    Dim fThereIs As Boolean
    
    '-------------------------------
    ' 正味財産（前期）
    '-------------------------------
    Dim iNetAssets_End_p   As Currency   '一般正味財産期末残高
    Dim iNetAssets_Begin_p As Currency   '一般正味財産期首残高
    Dim iNetAssets_Diff_p  As Currency   '当期一般正味財産増減額
    
    Dim iSpNetAssets_End_p   As Currency '指定正味財産期末残高
    Dim iSpNetAssets_Begin_p As Currency '指定正味財産期首残高
    Dim iSpNetAssets_Diff_p  As Currency '当期指定正味財産増減額
    
    '-------------------------------
    ' 正味財産（当期）
    '-------------------------------
    Dim iNetAssets_End   As Currency     '一般正味財産期末残高
    Dim iNetAssets_Begin As Currency     '一般正味財産期首残高
    Dim iNetAssets_Diff  As Currency     '当期一般正味財産増減額
    
    Dim iSpNetAssets_End   As Currency   '指定正味財産期末残高
    Dim iSpNetAssets_Begin As Currency   '指定正味財産期首残高
    Dim iSpNetAssets_Diff  As Currency   '当期指定正味財産増減額
    
    
    '==================================================
    'Phase [1]  勘定科目のリストを作成  （このフェーズは、まだ集計しない。リスト作成のみ）
    '--------------------------------------------------
    
    '---------------------------------------
    '[1]-(1) 「前年度実績」に記載の勘定科目を抽出
    '         同時に、前期繰越を取得
    '---------------------------------------
    yInput = styPerformance
    Do
        '前期実績を1行読み込み
        entrySide.ReadResults (yInput)
        
        '勘定科目コードの記載が無かったら、検索終了
        If IsNull(entrySide.AccountCode) Or IsEmpty(entrySide.AccountCode) Then Exit Do
        
        'すでに勘定科目があるか検索
        i = 0
        fThereIs = True
        While (i < cntAccount)
            If (Account(i, 0) = entrySide.AccountCode) Then
                fThereIs = False
            End If
            i = i + 1
        Wend
        '初出の科目だったら
        If (fThereIs = True) And (entrySide.Amount <> 0) Then
            '副科目コードに何も書いていなければ
            If IsNull(entrySide.SubAccountCode) Or IsEmpty(entrySide.SubAccountCode) Then
                                                              
                '前年度の増減額・期首残高・期末残高か？
                Select Case entrySide.AccountCode
                    Case idNetAssets_End
                        iNetAssets_End_p = entrySide.Amount
                    Case idNetAssets_Begin
                        iNetAssets_Begin_p = entrySide.Amount
                    Case idNetAssets_Diff
                        iNetAssets_Diff_p = entrySide.Amount
                    Case idSpNetAssets_End
                        iSpNetAssets_End_p = entrySide.Amount
                    Case idSpNetAssets_Begin
                        iSpNetAssets_Begin_p = entrySide.Amount
                    Case idSpNetAssets_Diff
                        iSpNetAssets_Diff_p = entrySide.Amount
                    Case Else
                        '上以外は、リスト化する
                End Select
                        
                        '■To Do 財務諸表出力をまとめれたら、上 Switch の Case Else に移動
                        Account(cntAccount, 0) = entrySide.AccountCode
                        Account(cntAccount, 1) = entrySide.Amount   '前年度繰越金
                        Account(cntAccount, 7) = entrySide.MajorAccount
                        cntAccount = cntAccount + 1
                            
                '-------------------------------------------------
                '■To Do 財務諸表出力は、まとめる。
                '財務諸表（前年度）への出力
                Call fs.OutFinancialStatements(entrySide.AccountCode, iLastX, entrySide.Amount)
                '-------------------------------------------------

            End If
        End If
        
        '次の行へ
        yInput = yInput + 1
    Loop
    
    '---------------------------------------
    '[1]-(2) 「仕訳帳」に記載の勘定科目を抽出
    '---------------------------------------
    For Each it In db.Items
        Set entry = it
        
        '年月日の記載が無かったら、検索終了
        If IsNull(entry.entryDate) Or IsEmpty(entry.entryDate) Then Exit For
        
        'すでに勘定科目があるか検索[借方]
        i = 0
        fThereIs = True
        While (i < cntAccount)
            If (Account(i, 0) = entry.Debit.AccountCode) Then
                fThereIs = False
            End If
            i = i + 1
        Wend
        If (fThereIs = True) Then
            Account(cntAccount, 0) = entry.Debit.AccountCode
            Account(cntAccount, 1) = 0  '前期繰越
            Account(cntAccount, 7) = entry.Debit.MajorAccount
            cntAccount = cntAccount + 1
        End If
        
        'すでに勘定科目があるか検索[貸方]
        i = 0
        fThereIs = True
        While (i < cntAccount)
            If (Account(i, 0) = entry.Credit.AccountCode) Then
                fThereIs = False
            End If
            i = i + 1
        Wend
        If (fThereIs = True) Then
            Account(cntAccount, 0) = entry.Credit.AccountCode
            Account(cntAccount, 1) = 0  '前期繰越
            Account(cntAccount, 7) = entry.Credit.MajorAccount
            cntAccount = cntAccount + 1
        End If
    Next
    
    
    '---------------------------------------
    '[1]-(3) ソート
    '---------------------------------------
    Call QuickSort(Account(), 10, 0, 0, cntAccount - 1)
    
    
    
    
    
    '==================================================
    'Phase [2]  総勘定元帳・補助元帳の生成　しながら、科目毎の集計
    '--------------------------------------------------
    
    '---------------------------------------
    '[2]-(0) 初期設定
    '---------------------------------------
    Sheet3.Activate
    
    '正味財産の集計用
    iNetAssets_Begin = iNetAssets_End_p     '一般正味財産期首残高
    iNetAssets_End = 0                      '一般正味財産期末残高
    iNetAssets_Diff = 0                     '当期一般正味財産増減額
    
    iSpNetAssets_Begin = iSpNetAssets_End_p '指定正味財産期首残高
    iSpNetAssets_End = 0                    '指定正味財産期末残高
    iSpNetAssets_Diff = 0                   '当期指定正味財産増減額
    
    '総勘定元帳への出力ｙ座標
    Dim yOutput
    Dim ySubOutput
    Dim yOutPerformance
    
    yOutput = styGeneralLedger
    ySubOutput = styGeneralLedger
    
    '今期実績への出力ｙ座標
    yOutPerformance = styPerformance

    '---------------------------------------
    '[2]-(1) リスト化された勘定科目の全項目について、
    '        各勘定科目の「貸方」、「借方」、「残高」を集計する
    '---------------------------------------
    i = 0
    While (i < cntAccount)
        
        '-------------------------------
        '1) 処理中の勘定科目に属す「補助科目」を前期実績から検索しリスト化する
        '-------------------------------
        '補助科目の配列開始位置
        stCntSubAccount = cntSubAccount
        
        '前期実績の開始位置
        yInput = styPerformance
        Do
            '前期実績を1行読み込み
            entrySide.ReadResults (yInput)

            '勘定科目コードの記載が無かったら、前期実績の検索終了
            If IsNull(entrySide.AccountCode) Or IsEmpty(entrySide.AccountCode) Then Exit Do

            '同じ勘定科目かチェック
            If Account(i, 0) = entrySide.AccountCode Then

                '補助科目コードに何か書かれているかチェック
                If IsNull(entrySide.SubAccountCode) Or IsEmpty(entrySide.SubAccountCode) Then
                Else
                    'すでに補助科目 且つ 同一の一般／指定があるか検索
                    iSub = stCntSubAccount
                    fThereIs = True
                    While (iSub < cntSubAccount)
                        If ((SubAccount(iSub, 0) = entrySide.SubAccountCode)) And (SubAccount(iSub, 6) = entrySide.Class) Then
                            fThereIs = False
                        End If
                        iSub = iSub + 1
                    Wend
    
                    '前期繰越があれば、補助科目を処理用配列に追加
                    If (fThereIs = True) And (entrySide.Amount <> 0) Then
                        '貸借対照表科目の場合、リスト化する
                        If (entrySide.AccountCode < 40000) Then
                            SubAccount(cntSubAccount, 0) = entrySide.SubAccountCode
                            SubAccount(cntSubAccount, 1) = entrySide.Amount   '前年度繰越金
                            SubAccount(cntSubAccount, 5) = entrySide.SubAccountName
                            SubAccount(cntSubAccount, 6) = entrySide.Class
                            cntSubAccount = cntSubAccount + 1
                        End If
                    End If
                End If
            End If

            yInput = yInput + 1
        Loop
    
        '-------------------------------
        '2) 「仕訳帳」から、「総勘定元帳」へ転記 ＆ 貸方、借方、残高の集計
        '   同時に、処理中の勘定科目に属す「補助科目」を検索しリスト化する
        '   （※「正味財産（貸借対照表勘定）」を除く）
        '-------------------------------
        If (Account(i, 0) < 30000) Or (Account(i, 0) >= 40000) Then
    
            yOutput = yOutput + 1
            
            '--------------------------
            '総勘定元帳へ科目名、等出力
            '1行目は勘定科目
            Sheet3.Range(Cells(yOutput, 1), Cells(yOutput, 7)).Merge
            
            Sheet3.Cells(yOutput, 1) = Account(i, 7)
            Sheet3.Cells(yOutput, 1).Font.Underline = True
            Sheet3.Cells(yOutput, 1).Font.size = 16
            Sheet3.Cells(yOutput, 1).HorizontalAlignment = xlCenter
            yOutput = yOutput + 1
            
            '--------------------------
            '2行目は勘定科目コード
            Sheet3.Cells(yOutput, 1) = "勘定科目コード：" & Str(Account(i, 0))
            yOutput = yOutput + 1
        
            '--------------------------
            '3行目は項目行
            Sheet3.Cells(yOutput, 1) = "日付"
            Sheet3.Cells(yOutput, 2) = "相手科目"
            Sheet3.Cells(yOutput, 3) = "摘要"
            Sheet3.Cells(yOutput, 4) = "借方"
            Sheet3.Cells(yOutput, 5) = "貸方"
            Sheet3.Cells(yOutput, 6) = "貸／借"
            Sheet3.Cells(yOutput, 7) = "残高"
            
            Sheet3.Range(Cells(yOutput, 1), Cells(yOutput, 7)).Font.Bold = True
            Sheet3.Range(Cells(yOutput, 1), Cells(yOutput, 7)).HorizontalAlignment = xlCenter
            
            Sheet3.Range(Cells(yOutput, 1), Cells(yOutput, 7)).Borders.LineStyle = xlContinuous
            Sheet3.Range(Cells(yOutput, 1), Cells(yOutput, 7)).Borders(xlEdgeTop).LineStyle = xlDouble
            Sheet3.Range(Cells(yOutput, 1), Cells(yOutput, 7)).Borders(xlEdgeBottom).Weight = xlMedium
            yOutput = yOutput + 1
        
            '--------------------------
            '4行目～データ
            
            Account(i, 2) = 0   '借方
            Account(i, 3) = 0   '借方
            Account(i, 4) = 0   '残高
            
            '---------------
            '前期繰越
            
            '前期繰越があり、且つ、「資産」・「負債」の勘定科目である場合、
            '総勘定元帳に「前期繰越」（前年度繰越金）を出力
            If (Account(i, 1) <> 0) And (Account(i, 0) < 30000) Then
                If Account(i, 1) > 0 Then
                    '借方の場合
                    Account(i, 2) = Account(i, 1)
                    Account(i, 4) = Account(i, 1)
                    Sheet3.Cells(yOutput, 4) = Account(i, 1)
                    Sheet3.Cells(yOutput, 5) = ""
                    Sheet3.Cells(yOutput, 6) = "借"
                    Sheet3.Cells(yOutput, 7) = Account(i, 1)
                ElseIf Account(i, 1) < 0 Then
                    '貸方の場合
                    Account(i, 3) = -Account(i, 1)
                    Account(i, 4) = Account(i, 1)
                    Sheet3.Cells(yOutput, 4) = ""
                    Sheet3.Cells(yOutput, 5) = -Account(i, 1)
                    Sheet3.Cells(yOutput, 6) = "貸"
                    Sheet3.Cells(yOutput, 7) = -Account(i, 1)
                End If
                '共通
                Sheet3.Cells(yOutput, 1) = Sheet1.Cells(14, 3)  '期首
                Sheet3.Cells(yOutput, 2) = ""
                Sheet3.Cells(yOutput, 3) = "前期繰越"
                Sheet3.Range(Cells(yOutput, 1), Cells(yOutput, 7)).Borders.LineStyle = xlContinuous
                yOutput = yOutput + 1
            End If
            
            '---------------
            '仕訳帳
            For Each it In db.Items
                Set entry = it

                '年月日の記載が無かったら、検索終了
                If IsNull(entry.entryDate) Or IsEmpty(entry.entryDate) Then Exit For
    
                '補助科目
                entrySide.SubAccountCode = 0
                
                '借方に集計中の勘定科目が記載されていた場合
                If Account(i, 0) = entry.Debit.AccountCode Then
                    Account(i, 2) = Account(i, 2) + entry.Debit.Amount
                    Account(i, 4) = Account(i, 4) + entry.Debit.Amount
                    If Account(i, 0) >= 40000 And Account(i, 0) < 80000 Then
                        If entry.Debit.Class = 1 Then
                            '指定正味財産
                            iSpNetAssets_Diff = iSpNetAssets_Diff - entry.Debit.Amount
                        Else
                            '一般正味財産
                            iNetAssets_Diff = iNetAssets_Diff - entry.Debit.Amount
                        End If
                    End If
                    
                    Sheet3.Cells(yOutput, 1) = entry.entryDate
                    Sheet3.Cells(yOutput, 2) = entry.Credit.MajorAccount
                    Sheet3.Cells(yOutput, 3) = entry.Summary
                    Sheet3.Cells(yOutput, 4) = entry.Debit.Amount
                    Sheet3.Cells(yOutput, 5) = ""
                    If Account(i, 4) < 0 Then
                        Sheet3.Cells(yOutput, 6) = "貸"
                    Else
                        Sheet3.Cells(yOutput, 6) = "借"
                    End If
                    Sheet3.Cells(yOutput, 7) = Abs(Account(i, 4))
                    Sheet3.Range(Cells(yOutput, 1), Cells(yOutput, 7)).Borders.LineStyle = xlContinuous
                    yOutput = yOutput + 1
                
                    '補助科目コードに何か書かれているかチェック
                    If IsNull(entry.Debit.SubAccountCode) Or IsEmpty(entry.Debit.SubAccountCode) Then
                    Else
                        entrySide.SubAccountCode = entry.Debit.SubAccountCode
                        entrySide.SubAccountName = entry.Debit.SubAccountName
                        entrySide.Class = entry.Debit.Class
                    End If
                
                End If
     
                '貸方に集計中の勘定科目が記載されていた場合
                If Account(i, 0) = entry.Credit.AccountCode Then
                    Account(i, 3) = Account(i, 3) + entry.Credit.Amount
                    Account(i, 4) = Account(i, 4) - entry.Credit.Amount
                    If Account(i, 0) >= 40000 And Account(i, 0) < 80000 Then
                        If entry.Credit.Class = 1 Then
                            '指定正味財産
                            iSpNetAssets_Diff = iSpNetAssets_Diff + entry.Credit.Amount
                        Else
                            '一般正味財産
                            iNetAssets_Diff = iNetAssets_Diff + entry.Credit.Amount
                        End If
                    End If
                    
                    Sheet3.Cells(yOutput, 1) = entry.entryDate
                    Sheet3.Cells(yOutput, 2) = entry.Debit.MajorAccount
                    Sheet3.Cells(yOutput, 3) = entry.Summary
                    Sheet3.Cells(yOutput, 4) = ""
                    Sheet3.Cells(yOutput, 5) = entry.Credit.Amount
                    If Account(i, 4) < 0 Then
                        Sheet3.Cells(yOutput, 6) = "貸"
                    Else
                        Sheet3.Cells(yOutput, 6) = "借"
                    End If
                    Sheet3.Cells(yOutput, 7) = Abs(Account(i, 4))
                    Sheet3.Range(Cells(yOutput, 1), Cells(yOutput, 7)).Borders.LineStyle = xlContinuous
                    yOutput = yOutput + 1
                
                    '補助科目コードに何か書かれているかチェック
                    If IsNull(entry.Credit.SubAccountCode) Or IsEmpty(entry.Credit.SubAccountCode) Then
                    Else
                        entrySide.SubAccountCode = entry.Credit.SubAccountCode
                        entrySide.SubAccountName = entry.Credit.SubAccountName
                        entrySide.Class = entry.Credit.Class
                    End If
                
                End If
                
                If (entrySide.SubAccountCode <> 0) Then
                    '既に既出の補助科目かチェック
                    'すでに補助科目があるか検索
                    iSub = stCntSubAccount
                    fThereIs = True
                    While (iSub < cntSubAccount)
                        If ((SubAccount(iSub, 0) = entrySide.SubAccountCode)) And (SubAccount(iSub, 6) = entrySide.Class) Then
                            fThereIs = False
                        End If
                        iSub = iSub + 1
                    Wend
                    '補助科目を処理用配列に追加
                    If fThereIs = True Then
                        SubAccount(cntSubAccount, 0) = entrySide.SubAccountCode
                        SubAccount(cntSubAccount, 1) = 0    '前期繰越
                        SubAccount(cntSubAccount, 5) = entrySide.SubAccountName
                        SubAccount(cntSubAccount, 6) = entrySide.Class
                        cntSubAccount = cntSubAccount + 1
                    End If
                End If
            
            Next
            
            '-------------------------------------------------
            '■To Do 財務諸表出力は、まとめる。
            '財務諸表への出力
            Call fs.OutFinancialStatements(Account(i, 0), iThisX, CCur(Account(i, 4)))
            '-------------------------------------------------
            
            '勘定科目は出力する。
            Sheet11.Cells(yOutPerformance, 1) = Account(i, 0)
            Sheet11.Cells(yOutPerformance, 2) = Account(i, 7)
            Sheet11.Cells(yOutPerformance, 3) = ""
            Sheet11.Cells(yOutPerformance, 4) = ""
            Sheet11.Cells(yOutPerformance, 5) = ""              'ここは、一般も指定も無い。
            Sheet11.Cells(yOutPerformance, 6) = Account(i, 4)
            yOutPerformance = yOutPerformance + 1
            
            '終わり
            yOutput = yOutput + 1
        End If
        
        
        '-------------------------------
        '3) リスト化された、処理中の勘定科目に属す「補助科目」について、
        '   「仕訳帳」から、「補助元帳」へ転記 ＆ 貸方、借方、残高の集計
        '-------------------------------
        
        EndCntSubAccount = cntSubAccount - 1
        Account(i, 5) = stCntSubAccount
        Account(i, 6) = EndCntSubAccount
        
        '補助元帳の生成が必要か（補助科目有り）？
        If stCntSubAccount <= EndCntSubAccount Then
            
            Sheet4.Activate
            
            'ソート
            Call QuickSort(SubAccount(), 6, 0, Account(i, 5), Account(i, 6))

            '登録された勘定科目の分、繰り返す
            iSub = stCntSubAccount
            While (iSub < cntSubAccount)
            
                ySubOutput = ySubOutput + 1
                
                '--------------------------
                '総勘定元帳へ科目名、等出力
                '1行目は勘定科目
                Sheet4.Range(Cells(ySubOutput, 1), Cells(ySubOutput, 7)).Merge
                Sheet4.Cells(ySubOutput, 1) = "科目：" & Account(i, 7)
                Sheet4.Cells(ySubOutput, 1).Font.Underline = True
                Sheet4.Cells(ySubOutput, 1).Font.size = 16
                Sheet4.Cells(ySubOutput, 1).HorizontalAlignment = xlCenter
                
                ySubOutput = ySubOutput + 1
                
                Sheet4.Range(Cells(ySubOutput, 1), Cells(ySubOutput, 7)).Merge
                If SubAccount(iSub, 6) = 1 Then
                    Sheet4.Cells(ySubOutput, 1) = "補助科目：" & SubAccount(iSub, 5) & "(指定)"
                Else
                    Sheet4.Cells(ySubOutput, 1) = "補助科目：" & SubAccount(iSub, 5)
                End If
                Sheet4.Cells(ySubOutput, 1).Font.Underline = True
                Sheet4.Cells(ySubOutput, 1).Font.size = 14
                Sheet4.Cells(ySubOutput, 1).HorizontalAlignment = xlCenter
                
                ySubOutput = ySubOutput + 2
                
                '--------------------------
                '2行目は勘定科目コード
                Sheet4.Cells(ySubOutput, 1) = "勘定科目コード：" & Str(Account(i, 0))
                Sheet4.Cells(ySubOutput, 3) = "補助科目コード：" & Str(SubAccount(iSub, 0))
                ySubOutput = ySubOutput + 1
            
                '--------------------------
                '3行目は項目行
                Sheet4.Cells(ySubOutput, 1) = "日付"
                Sheet4.Cells(ySubOutput, 2) = "相手科目"
                Sheet4.Cells(ySubOutput, 3) = "摘要"
                Sheet4.Cells(ySubOutput, 4) = "借方"
                Sheet4.Cells(ySubOutput, 5) = "貸方"
                Sheet4.Cells(ySubOutput, 6) = "貸／借"
                Sheet4.Cells(ySubOutput, 7) = "残高"
                
                Sheet4.Range(Cells(ySubOutput, 1), Cells(ySubOutput, 7)).Font.Bold = True
                Sheet4.Range(Cells(ySubOutput, 1), Cells(ySubOutput, 7)).HorizontalAlignment = xlCenter
                
                Sheet4.Range(Cells(ySubOutput, 1), Cells(ySubOutput, 7)).Borders.LineStyle = xlContinuous
                Sheet4.Range(Cells(ySubOutput, 1), Cells(ySubOutput, 7)).Borders(xlEdgeTop).LineStyle = xlDouble
                Sheet4.Range(Cells(ySubOutput, 1), Cells(ySubOutput, 7)).Borders(xlEdgeBottom).Weight = xlMedium
                ySubOutput = ySubOutput + 1

            
                '--------------------------
                '4行目～データ
                
                SubAccount(iSub, 2) = 0   '借方
                SubAccount(iSub, 3) = 0   '借方
                SubAccount(iSub, 4) = 0   '残高
                
                '---------------
                '前期繰越
                
                '前期繰越があり、且つ、「資産」・「負債」の勘定科目である場合、
                '元帳に「前期繰越」（前年度繰越金）を出力
                If (SubAccount(iSub, 1) <> 0) And (Account(i, 0) < 30000) Then
                    If SubAccount(iSub, 1) > 0 Then
                        '借方の場合
                        SubAccount(iSub, 2) = SubAccount(iSub, 1)
                        SubAccount(iSub, 4) = SubAccount(iSub, 1)
                        Sheet4.Cells(ySubOutput, 4) = SubAccount(iSub, 1)
                        Sheet4.Cells(ySubOutput, 5) = ""
                        Sheet4.Cells(ySubOutput, 6) = "借"
                        Sheet4.Cells(ySubOutput, 7) = SubAccount(iSub, 1)
                    ElseIf SubAccount(iSub, 1) < 0 Then
                        '貸方の場合
                        SubAccount(iSub, 3) = -SubAccount(iSub, 1)
                        SubAccount(iSub, 4) = SubAccount(iSub, 1)
                        Sheet4.Cells(ySubOutput, 4) = ""
                        Sheet4.Cells(ySubOutput, 5) = -SubAccount(iSub, 1)
                        Sheet4.Cells(ySubOutput, 6) = "貸"
                        Sheet4.Cells(ySubOutput, 7) = -SubAccount(iSub, 1)
                    End If
                    '共通
                    Sheet4.Cells(ySubOutput, 1) = Sheet1.Cells(14, 3)  '期首
                    Sheet4.Cells(ySubOutput, 2) = ""
                    Sheet4.Cells(ySubOutput, 3) = "前期繰越"
                    Sheet4.Range(Cells(ySubOutput, 1), Cells(ySubOutput, 7)).Borders.LineStyle = xlContinuous
                    ySubOutput = ySubOutput + 1

                End If
                
                '---------------
                '仕訳帳
                For Each it In db.Items
                    Set entry = it
                    
                    '年月日の記載が無かったら、検索終了
                    If IsNull(entry.entryDate) Or IsEmpty(entry.entryDate) Then Exit For
                   
                    '借方に集計中の勘定科目が記載されていた場合
                    If (Account(i, 0) = entry.Debit.AccountCode) And (SubAccount(iSub, 0) = entry.Debit.SubAccountCode) And (SubAccount(iSub, 6) = entry.Debit.Class) Then
                        SubAccount(iSub, 2) = SubAccount(iSub, 2) + entry.Debit.Amount
                        SubAccount(iSub, 4) = SubAccount(iSub, 4) + entry.Debit.Amount
                        
                        Sheet4.Cells(ySubOutput, 1) = entry.entryDate
                        Sheet4.Cells(ySubOutput, 2) = entry.Credit.MajorAccount  '相手科目
                        Sheet4.Cells(ySubOutput, 3) = entry.Summary
                        Sheet4.Cells(ySubOutput, 4) = entry.Debit.Amount
                        Sheet4.Cells(ySubOutput, 5) = ""
                        If SubAccount(iSub, 4) < 0 Then
                            Sheet4.Cells(ySubOutput, 6) = "貸"
                        Else
                            Sheet4.Cells(ySubOutput, 6) = "借"
                        End If
                        Sheet4.Cells(ySubOutput, 7) = Abs(SubAccount(iSub, 4))
                        Sheet4.Range(Cells(ySubOutput, 1), Cells(ySubOutput, 7)).Borders.LineStyle = xlContinuous
                        ySubOutput = ySubOutput + 1
                    End If
                    
                    '貸方に集計中の勘定科目が記載されていた場合
                    If (Account(i, 0) = entry.Credit.AccountCode) And (SubAccount(iSub, 0) = entry.Credit.SubAccountCode) And (SubAccount(iSub, 6) = entry.Credit.Class) Then
                        SubAccount(iSub, 3) = SubAccount(iSub, 3) + entry.Credit.Amount
                        SubAccount(iSub, 4) = SubAccount(iSub, 4) - entry.Credit.Amount
                        
                        Sheet4.Cells(ySubOutput, 1) = entry.entryDate
                        Sheet4.Cells(ySubOutput, 2) = entry.Debit.MajorAccount   '相手科目
                        Sheet4.Cells(ySubOutput, 3) = entry.Summary
                        Sheet4.Cells(ySubOutput, 4) = ""
                        Sheet4.Cells(ySubOutput, 5) = entry.Credit.Amount
                        If SubAccount(iSub, 4) < 0 Then
                            Sheet4.Cells(ySubOutput, 6) = "貸"
                        Else
                            Sheet4.Cells(ySubOutput, 6) = "借"
                        End If
                        Sheet4.Cells(ySubOutput, 7) = Abs(SubAccount(iSub, 4))
                        Sheet4.Range(Cells(ySubOutput, 1), Cells(ySubOutput, 7)).Borders.LineStyle = xlContinuous
                        ySubOutput = ySubOutput + 1
                    End If
                    
                    yInput = yInput + 1
                Next
                
                '補助科目は、残高が０でない場合に出力する。
                If SubAccount(iSub, 4) <> 0 Then
                    '今期実績への出力
                    Sheet11.Cells(yOutPerformance, 1) = Account(i, 0)       '勘定科目コード
                    Sheet11.Cells(yOutPerformance, 2) = Account(i, 7)       '勘定科目名
                    Sheet11.Cells(yOutPerformance, 3) = SubAccount(iSub, 0) '補助科目コード
                    Sheet11.Cells(yOutPerformance, 4) = SubAccount(iSub, 5) '補助科目名
                    Sheet11.Cells(yOutPerformance, 5) = SubAccount(iSub, 6)
                    Sheet11.Cells(yOutPerformance, 6) = SubAccount(iSub, 4) '金額
                    yOutPerformance = yOutPerformance + 1
                End If
                
                '基本財産への充当額
                '一般正味財産から充当
                '指定正味財産から充当
                
                
                
                
                
                '特定資産への充当額
                '一般正味財産から充当
                '指定正味財産から充当
                
                
                
                
                
                
                
                '終わり
                ySubOutput = ySubOutput + 1
                
                iSub = iSub + 1
            Wend
        
            Sheet3.Activate

        End If
        
        i = i + 1
    
    Wend


    '---------------------------------------
    '[2]-(2) 正味財産の部
    '---------------------------------------
    
    '■■■■To Do: 以下項目の表示
    '（うち基本財産への充当額）
    '（うち特定資産への充当額）
    
    '一般正味財産増減計算書へ出力
    iNetAssets_End = iNetAssets_Begin + iNetAssets_Diff
    iSpNetAssets_End = iSpNetAssets_Begin + iSpNetAssets_Diff
    
    Call fs.OutFinancialStatements(idNetAssets_End, iThisX, iNetAssets_End)
    Call fs.OutFinancialStatements(idNetAssets_Begin, iThisX, iNetAssets_Begin)
    Call fs.OutFinancialStatements(idNetAssets_Diff, iThisX, iNetAssets_Diff)
    Call fs.OutFinancialStatements(idSpNetAssets_End, iThisX, iSpNetAssets_End)
    Call fs.OutFinancialStatements(idSpNetAssets_Begin, iThisX, iSpNetAssets_Begin)
    Call fs.OutFinancialStatements(idSpNetAssets_Diff, iThisX, iSpNetAssets_Diff)

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
    '       7   勘定科目名（文字列データ）
    
    '今期実績への出力
    Account(cntAccount, 0) = idNetAssets_Diff
    Account(cntAccount, 1) = iNetAssets_Diff_p
    Account(cntAccount, 4) = iNetAssets_Diff
    Account(cntAccount, 7) = "当期一般正味財産増減額"
    cntAccount = cntAccount + 1

    Account(cntAccount, 0) = idNetAssets_Begin
    Account(cntAccount, 1) = iNetAssets_Begin_p
    Account(cntAccount, 4) = iNetAssets_Begin
    Account(cntAccount, 7) = "一般正味財産期首残高"
    cntAccount = cntAccount + 1

    Account(cntAccount, 0) = idNetAssets_End
    Account(cntAccount, 1) = iNetAssets_End_p
    Account(cntAccount, 4) = iNetAssets_End
    Account(cntAccount, 7) = "一般正味財産期末残高"
    cntAccount = cntAccount + 1
    
    Account(cntAccount, 0) = idSpNetAssets_Diff
    Account(cntAccount, 1) = iSpNetAssets_Diff_p
    Account(cntAccount, 4) = iSpNetAssets_Diff
    Account(cntAccount, 7) = "当期一般正味財産増減額"
    cntAccount = cntAccount + 1

    Account(cntAccount, 0) = idSpNetAssets_Begin
    Account(cntAccount, 1) = iSpNetAssets_Begin_p
    Account(cntAccount, 4) = iSpNetAssets_Begin
    Account(cntAccount, 7) = "一般正味財産期首残高"
    cntAccount = cntAccount + 1

    Account(cntAccount, 0) = idSpNetAssets_End
    Account(cntAccount, 1) = iSpNetAssets_End_p
    Account(cntAccount, 4) = iSpNetAssets_End
    Account(cntAccount, 7) = "一般正味財産期末残高"
    cntAccount = cntAccount + 1

    
    Sheet11.Cells(yOutPerformance, 1) = idNetAssets_Diff
    Sheet11.Cells(yOutPerformance, 2) = "当期一般正味財産増減額"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = ""
    Sheet11.Cells(yOutPerformance, 6) = iNetAssets_Diff
    yOutPerformance = yOutPerformance + 1
    
    Sheet11.Cells(yOutPerformance, 1) = idNetAssets_Begin
    Sheet11.Cells(yOutPerformance, 2) = "一般正味財産期首残高"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = ""
    Sheet11.Cells(yOutPerformance, 6) = iNetAssets_Begin
    yOutPerformance = yOutPerformance + 1
            
    Sheet11.Cells(yOutPerformance, 1) = idNetAssets_End
    Sheet11.Cells(yOutPerformance, 2) = "一般正味財産期末残高"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = ""
    Sheet11.Cells(yOutPerformance, 6) = iNetAssets_End
    yOutPerformance = yOutPerformance + 1

    Sheet11.Cells(yOutPerformance, 1) = idSpNetAssets_Diff
    Sheet11.Cells(yOutPerformance, 2) = "当期指定正味財産増減額"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = ""
    Sheet11.Cells(yOutPerformance, 6) = iSpNetAssets_Diff
    yOutPerformance = yOutPerformance + 1
    
    Sheet11.Cells(yOutPerformance, 1) = idSpNetAssets_Begin
    Sheet11.Cells(yOutPerformance, 2) = "指定正味財産期首残高"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = ""
    Sheet11.Cells(yOutPerformance, 6) = iSpNetAssets_Begin
    yOutPerformance = yOutPerformance + 1
            
    Sheet11.Cells(yOutPerformance, 1) = idSpNetAssets_End
    Sheet11.Cells(yOutPerformance, 2) = "指定正味財産期末残高"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = ""
    Sheet11.Cells(yOutPerformance, 6) = iSpNetAssets_End
    yOutPerformance = yOutPerformance + 1

End Sub
