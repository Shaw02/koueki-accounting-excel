Attribute VB_Name = "Module1"
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

'財務諸表（貸借対照表、正味財産増減計算書）の座標
Public Const styFinancialStatements = 6 '財務諸表の開始位置

Public Const iThisX = 6             '当年度
Public Const iLastX = 7             '前年度

'仕訳帳の読み込みＹ座標
Public Const styJournal = 8         '仕訳帳の開始位置
    
'総勘定元帳の出力Ｙ座標
Public Const styGeneralLedger = 5   '総勘定元帳の開始位置

'実績の読み込みＹ座標
Public Const styPerformance = 7     '実績の開始位置

'==============================
' 仕訳帳（Sheet2）列定義
'==============================
Public Const JOURNAL_HEADER_ROW As Long = 7

Public Const COL_DATE As Long = 1

' --- 借方 ---
Public Const COL_CR_ACC_NO_I   As Long = 2
Public Const COL_CR_ACC_NO_S   As Long = 3
Public Const COL_CR_ACC_NO_S2  As Long = 4
Public Const COL_CR_SUB_NO_I   As Long = 5
Public Const COL_CR_SUB_NO_S   As Long = 6
Public Const COL_CR_MONEY      As Long = 7

' --- 貸方 ---
Public Const COL_DR_ACC_NO_I   As Long = 8
Public Const COL_DR_ACC_NO_S   As Long = 9
Public Const COL_DR_ACC_NO_S2  As Long = 10
Public Const COL_DR_SUB_NO_I   As Long = 11
Public Const COL_DR_SUB_NO_S   As Long = 12
Public Const COL_DR_MONEY      As Long = 13

' --- 摘要 ---
Public Const COL_SUMMARY       As Long = 14

'=======================================================
'       グローバル変数　宣言
'-------------------------------------------------------
'仕訳帳からの読み込み用
Public iYear                        '日付
Public sSummary As String           '摘要
       
        '借方
Public iCrAccountNo                 '勘定科目コード
Public sCrAccountNo As String       '勘定科目名
Public iCrSubAccountNo              '補助科目コード
Public sCrSubAccountNo As String    '補助科目名
Public sCrClass As String           '借方区分（一般／指定）
Public iCrMoney                     '金額
        
        '貸方
Public iDrAccountNo                 '勘定科目コード
Public sDrAccountNo As String       '勘定科目名
Public iDrSubAccountNo              '補助科目コード
Public sDrSubAccountNo As String    '補助科目名
Public sDrClass As String           '貸方区分（一般／指定）
Public iDrMoney                     '金額

'=======================================================
'       総勘定元帳・今期実績のクリア
'-------------------------------------------------------
Sub GeneralAccountClear()
'
' GeneralAccountClear Macro
'
' Keyboard Shortcut: Ctrl+g
'
    '今期実績
    Sheet11.Activate
    Sheet11.Range("A7", Sheet11.Range("A7").SpecialCells(xlCellTypeLastCell)).Select
    Selection.EntireRow.Delete
    Sheet11.Range("A7").Select
    
    '総勘定元帳
    Sheet3.Activate
    Sheet3.Range("A5", Sheet3.Range("A5").SpecialCells(xlCellTypeLastCell)).Select
    Selection.EntireRow.Delete
    Sheet3.Range("A5").Select

    '補助元帳
    Sheet4.Activate
    Sheet4.Range("A5", Sheet4.Range("A5").SpecialCells(xlCellTypeLastCell)).Select
    Selection.EntireRow.Delete
    Sheet4.Range("A5").Select

End Sub
'=======================================================
'       財務諸表の０初期化
'-------------------------------------------------------
Sub FinancialStatementsClear()

    '貸借対照表へ
    yOutBS = styFinancialStatements
    Do
        If IsNull(Sheet6.Cells(yOutBS, 1)) Or IsEmpty(Sheet6.Cells(yOutBS, 1)) Then
        Else
            If Sheet6.Cells(yOutBS, 1) = -1 Then
                Exit Do
            Else
                Sheet6.Cells(yOutBS, 6) = 0
                Sheet6.Cells(yOutBS, 7) = 0
            End If
        End If
        yOutBS = yOutBS + 1
    Loop
    
    '正味財産増減計算書へ
    yOutPL = styFinancialStatements
    Do
        If IsNull(Sheet5.Cells(yOutPL, 1)) Or IsEmpty(Sheet5.Cells(yOutPL, 1)) Then
        Else
            If Sheet5.Cells(yOutPL, 1) = -1 Then
                Exit Do
            Else
                Sheet5.Cells(yOutPL, 6) = 0
                Sheet5.Cells(yOutPL, 7) = 0
            End If
        End If
        yOutPL = yOutPL + 1
    Loop

End Sub

'=======================================================
'       仕訳帳読み込み
'-------------------------------------------------------
'   Contents:
'       仕訳帳から、１行分、上のグローバル変数に読み込みます。
'
'   Input:
'       y       入力Ｙ座標
'
'   Output:
'       ＜以下グローバル変数＞
'       iYear           日付
'       iCrAccountNo    借方勘定科目コード
'       sCrAccountNo    借方勘定科目
'       iCrSubAccountNo 借方補助科目コード
'       sCrSubAccountNo 借方補助科目
'       iCrMoney        借方金額
'       iDrAccountNo    貸方勘定科目コード
'       sDrAccountNo    貸方勘定科目
'       iDrSubAccountNo 貸方補助科目コード
'       sDrSubAccountNo 貸方補助科目
'       iDrMoney        貸方金額
'       sSummary        摘要
'
'=======================================================
Sub ReadJournal(y As Variant)

    '====================
    ' 年月日
    '====================
    iYear = Sheet2.Cells(y, COL_DATE)

    '====================
    ' 借方
    '====================
    iCrAccountNo = Sheet2.Cells(y, COL_CR_ACC_NO_I)
    sCrAccountNo = Sheet2.Cells(y, COL_CR_ACC_NO_S)

    If Sheet2.Cells(y, COL_CR_ACC_NO_S2) <> "" Then
        sCrAccountNo = sCrAccountNo & "－" & Sheet2.Cells(y, COL_CR_ACC_NO_S2)
    End If

    iCrSubAccountNo = Sheet2.Cells(y, COL_CR_SUB_NO_I)
    sCrSubAccountNo = Sheet2.Cells(y, COL_CR_SUB_NO_S)
    sCrClass = "一般"     ' ← 後で「一般／指定」に拡張
    iCrMoney = Sheet2.Cells(y, COL_CR_MONEY)

    '====================
    ' 貸方
    '====================
    iDrAccountNo = Sheet2.Cells(y, COL_DR_ACC_NO_I)
    sDrAccountNo = Sheet2.Cells(y, COL_DR_ACC_NO_S)

    If Sheet2.Cells(y, COL_DR_ACC_NO_S2) <> "" Then
        sDrAccountNo = sDrAccountNo & "－" & Sheet2.Cells(y, COL_DR_ACC_NO_S2)
    End If

    iDrSubAccountNo = Sheet2.Cells(y, COL_DR_SUB_NO_I)
    sDrSubAccountNo = Sheet2.Cells(y, COL_DR_SUB_NO_S)
    sDrClass = "一般"     ' ← 同上
    iDrMoney = Sheet2.Cells(y, COL_DR_MONEY)

    '====================
    ' 摘要
    '====================
    sSummary = Sheet2.Cells(y, COL_SUMMARY)

End Sub
'=======================================================
'       財務諸表へ出力
'-------------------------------------------------------
'   Contents:
'       正味財産増減計算書、若しくは貸借対照表の
'       指定の勘定科目コードの欄に、金額を出力します。
'
'   Input:
'       id      勘定科目コード
'       x       出力Ｘ座標
'       iMoney  金額
'
'   Output:
'       Sheet "正味財産増減計算書"
'       Sheet "貸借対照表"
'
'=======================================================
Sub OutFinancialStatements(idAccount As Variant, iColumn As Variant, iMoney As Variant)

    '財務諸表への出力
    If (idAccount < 30000) Then
        '貸借対照表へ
        yOutBS = 6
        flagErr = True
        Do
            If IsNull(Sheet6.Cells(yOutBS, 1)) Or IsEmpty(Sheet6.Cells(yOutBS, 1)) Then
            Else
                If Sheet6.Cells(yOutBS, 1) = -1 Then
                    Exit Do
                ElseIf Sheet6.Cells(yOutBS, 1) = idAccount Then
                    If idAccount >= 10000 And idAccount < 20000 Then
                        '借方
                        Sheet6.Cells(yOutBS, iColumn) = iMoney
                    Else
                        '貸方
                        Sheet6.Cells(yOutBS, iColumn) = -iMoney
                    End If
                    flagErr = False
                End If
            End If
            yOutBS = yOutBS + 1
        Loop
        If flagErr = True Then
            MsgBox "貸借対照表に勘定科目がありません。：勘定科目コード＝" & Str(idAccount)
            End
        End If
        
    Else
        '正味財産増減計算書へ
        yOutPL = 6
        flagErr = True
        Do
            If IsNull(Sheet5.Cells(yOutPL, 1)) Or IsEmpty(Sheet5.Cells(yOutPL, 1)) Then
            Else
                If Sheet5.Cells(yOutPL, 1) = -1 Then
                    Exit Do
                ElseIf Sheet5.Cells(yOutPL, 1) = idAccount Then
                    If idAccount >= 30000 And idAccount < 40000 Then
                        '集計結果
                        Sheet5.Cells(yOutPL, iColumn) = iMoney
                    ElseIf idAccount >= 40000 And idAccount < 50000 Then
                        '一般／経常収益
                        Sheet5.Cells(yOutPL, iColumn) = -iMoney
                    ElseIf idAccount >= 50000 And idAccount < 60000 Then
                        '一般／経常費用
                        Sheet5.Cells(yOutPL, iColumn) = iMoney
                    ElseIf idAccount >= 60000 And idAccount < 70000 Then
                        '一般／経常外収益
                        Sheet5.Cells(yOutPL, iColumn) = -iMoney
                    ElseIf idAccount >= 70000 And idAccount < 80000 Then
                        '一般／経常外費用
                        Sheet5.Cells(yOutPL, iColumn) = iMoney
                    ElseIf idAccount >= 80000 And idAccount < 90000 Then
                        '指定／収益
                        Sheet5.Cells(yOutPL, iColumn) = -iMoney
                    ElseIf idAccount >= 90000 And idAccount < 100000 Then
                        '指定／費用
                        Sheet5.Cells(yOutPL, iColumn) = iMoney
                    End If
                    flagErr = False
                End If
            End If
            yOutPL = yOutPL + 1
        Loop
        If flagErr = True Then
            MsgBox "正味財産増減計算書に勘定科目がありません。：勘定科目コード＝" & Str(idAccount)
            End
        End If

    End If

End Sub
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
    '       7   勘定科目名（文字列データ）
    Dim Account(99999, 8) As Variant
    
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
    Dim SubAccount(99999, 6) As Variant
   
    '--------------------------
    '科目
    cntAccount = 0          '勘定科目数
    cntSubAccount = 0       '補助科目数
    
    '--------------------------
    '正味財産の集計用
    iNetAssets_End = 0      '一般正味財産期末残高
    iNetAssets_Begin = 0    '一般正味財産期首残高
    iNetAssets_Diff = 0     '当期一般正味財産増減額
    iSpNetAssets_End = 0    '指定正味財産期末残高
    iSpNetAssets_Begin = 0  '指定正味財産期首残高
    iSpNetAssets_Diff = 0   '当期指定正味財産増減額
    
    '--------------------------
    '初期化
    
    '財務諸表のクリア
    Call FinancialStatementsClear
    
    '総勘定元帳 をクリア
    Call GeneralAccountClear
    
    '==================================================
    'Phase [1]  勘定科目のリストを作成
    '--------------------------------------------------
    
    '---------------------------------------
    '[1]-(1) 「前年度実績」に記載の勘定科目を抽出
    '         同時に、前期繰越を取得
    '---------------------------------------
    yInput = styPerformance
    Do
        '前期実績を1行読み込み
        iAccountNo = Sheet10.Cells(yInput, 1)
        sAccountNo = Sheet10.Cells(yInput, 2)
        iSubAccountNo = Sheet10.Cells(yInput, 3)
        sSubAccountNo = Sheet10.Cells(yInput, 4)
        iMoney = Sheet10.Cells(yInput, 5)
        
        '勘定科目コードの記載が無かったら、検索終了
        If IsNull(iAccountNo) Or IsEmpty(iAccountNo) Then Exit Do
        
        'すでに勘定科目があるか検索
        i = 0
        fThereIs = True
        While (i < cntAccount)
            If (Account(i, 0) = iAccountNo) Then
                fThereIs = False
            End If
            i = i + 1
        Wend
        '初出の科目だったら
        If (fThereIs = True) And (iMoney <> 0) Then
            '副科目コードに何も書いていなければ
            If IsNull(iSubAccountNo) Or IsEmpty(iSubAccountNo) Then
                                
                '貸借対照表科目の場合、リスト化する
                If (iAccountNo < 40000) Then
                    'リスト化
                    Account(cntAccount, 0) = iAccountNo
                    Account(cntAccount, 1) = iMoney   '前年度繰越金
                    Account(cntAccount, 7) = sAccountNo
                    cntAccount = cntAccount + 1
                End If
                
                Select Case iAccountNo
                    Case idNetAssets_End
                        iNetAssets_End = iMoney
                    Case idNetAssets_Begin
                        iNetAssets_Begin = iMoney
                    Case idNetAssets_Diff
                        iNetAssets_Diff = iMoney
                    Case idSpNetAssets_End
                        iSpNetAssets_End = iMoney
                    Case idSpNetAssets_Begin
                        iSpNetAssets_Begin = iMoney
                    Case idSpNetAssets_Diff
                        iSpNetAssets_Diff = iMoney
                End Select
                            
                '財務諸表（前年度）への出力
                Call OutFinancialStatements(iAccountNo, iLastX, iMoney)

            End If
        End If
        
        '次の行へ
        yInput = yInput + 1
    Loop
    
    '---------------------------------------
    '[1]-(2) 「仕訳帳」に記載の勘定科目を抽出
    '---------------------------------------
    
    '仕訳帳のチェック用
    iTotalCrMoney = 0
    iTotalDrMoney = 0

    yInput = styJournal
    Do
        
        '仕訳帳を読み込み
        ReadJournal (yInput)
        
        '年月日の記載が無かったら、検索終了
        If IsNull(iYear) Or IsEmpty(iYear) Then Exit Do
        
        iTotalCrMoney = iTotalCrMoney + iCrMoney
        iTotalDrMoney = iTotalDrMoney + iDrMoney

        'すでに勘定科目があるか検索[借方]
        i = 0
        fThereIs = True
        While (i < cntAccount)
            If (Account(i, 0) = iCrAccountNo) Then
                fThereIs = False
            End If
            i = i + 1
        Wend
        If (fThereIs = True) Then
            Account(cntAccount, 0) = iCrAccountNo
            Account(cntAccount, 1) = 0  '前期繰越
            Account(cntAccount, 7) = sCrAccountNo
            cntAccount = cntAccount + 1
        End If
        
        'すでに勘定科目があるか検索[貸方]
        i = 0
        fThereIs = True
        While (i < cntAccount)
            If (Account(i, 0) = iDrAccountNo) Then
                fThereIs = False
            End If
            i = i + 1
        Wend
        If (fThereIs = True) Then
            Account(cntAccount, 0) = iDrAccountNo
            Account(cntAccount, 1) = 0  '前期繰越
            Account(cntAccount, 7) = sDrAccountNo
            cntAccount = cntAccount + 1
        End If
    
        'チェック
        
        '資産じゃない場合、且つ、補助科目が未入力でないか、チェック
        If (iCrAccountNo >= 20000) And (IsNull(iCrSubAccountNo) Or IsEmpty(iCrSubAccountNo)) Then
            MsgBox "仕訳帳 " & Str(yInput) & " 行目：借方に補助科目が必要です"
        End If
        If (iDrAccountNo >= 20000) And (IsNull(iDrSubAccountNo) Or IsEmpty(iDrSubAccountNo)) Then
            MsgBox "仕訳帳 " & Str(yInput) & " 行目：貸方に補助科目が必要です"
        End If
        If (iTotalCrMoney <> iTotalDrMoney) Then
            MsgBox "仕訳帳 " & Str(yInput) & " 行目：貸方金額と借方金額の合計が一致しません。"
            End
        End If
        
        yInput = yInput + 1
    Loop
    
    
    '---------------------------------------
    '[1]-(3) ソート
    '---------------------------------------
    Call QuickSort(Account(), 7, 0, 0, cntAccount - 1)
    
    
    
    '==================================================
    'Phase [2]  総勘定元帳・補助元帳の生成　しながら、科目毎の集計
    '--------------------------------------------------
    
    '---------------------------------------
    '[2]-(0) 初期設定
    '---------------------------------------
    Sheet3.Activate
    
    '正味財産の集計用
    iNetAssets_Begin = iNetAssets_End       '一般正味財産期首残高
    iNetAssets_End = 0                      '一般正味財産期末残高
    iNetAssets_Diff = 0                     '当期一般正味財産増減額
    
    iSpNetAssets_Begin = iSpNetAssets_End   '指定正味財産期首残高
    iSpNetAssets_End = 0                    '指定正味財産期末残高
    iSpNetAssets_Diff = 0                   '当期指定正味財産増減額
    
    '総勘定元帳への出力ｙ座標
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
        
        yInput = styPerformance
        Do
            '前期実績読み込み
            iAccountNo = Sheet10.Cells(yInput, 1)
            sAccountNo = Sheet10.Cells(yInput, 2)
            iSubAccountNo = Sheet10.Cells(yInput, 3)
            sSubAccountNo = Sheet10.Cells(yInput, 4)
            iMoney = Sheet10.Cells(yInput, 5)

            '勘定科目コードの記載が無かったら、前期実績の検索終了
            If IsNull(iAccountNo) Or IsEmpty(iAccountNo) Then Exit Do

            '同じ勘定科目かチェック
            If Account(i, 0) = iAccountNo Then

                '補助科目コードに何か書かれているかチェック
                If IsNull(iSubAccountNo) Or IsEmpty(iSubAccountNo) Then
                Else
                    'すでに補助科目があるか検索
                    iSub = stCntSubAccount
                    fThereIs = True
                    While (iSub < cntSubAccount)
                        If (SubAccount(iSub, 0) = iSubAccountNo) Then
                            fThereIs = False
                        End If
                        iSub = iSub + 1
                    Wend
    
                    '前期繰越があれば、補助科目を処理用配列に追加
                    If (fThereIs = True) And (iMoney <> 0) Then
                        '貸借対照表科目の場合、リスト化する
                        If (iAccountNo < 40000) Then
                            SubAccount(cntSubAccount, 0) = iSubAccountNo
                            SubAccount(cntSubAccount, 1) = iMoney   '前年度繰越金
                            SubAccount(cntSubAccount, 5) = sSubAccountNo
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
            Sheet3.Cells(yOutput, 1).Font.Size = 16
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
            
            '仕訳帳の先頭Ｙ座標
            yInput = styJournal
            
            Do
                '仕訳帳を読み込み
                ReadJournal (yInput)
                
                '年月日の記載が無かったら、検索終了
                If IsNull(iYear) Or IsEmpty(iYear) Then Exit Do
    
                '補助科目
                iSubAccountNo = 0
                
                '借方に集計中の勘定科目が記載されていた場合
                If Account(i, 0) = iCrAccountNo Then
                    Account(i, 2) = Account(i, 2) + iCrMoney
                    Account(i, 4) = Account(i, 4) + iCrMoney
                    
                    Sheet3.Cells(yOutput, 1) = iYear
                    Sheet3.Cells(yOutput, 2) = sDrAccountNo
                    Sheet3.Cells(yOutput, 3) = sSummary
                    Sheet3.Cells(yOutput, 4) = iCrMoney
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
                    If IsNull(iCrSubAccountNo) Or IsEmpty(iCrSubAccountNo) Then
                    Else
                        iSubAccountNo = iCrSubAccountNo
                        sSubAccountNo = sCrSubAccountNo
                    End If
                
                End If
     
                '貸方に集計中の勘定科目が記載されていた場合
                If Account(i, 0) = iDrAccountNo Then
                    Account(i, 3) = Account(i, 3) + iDrMoney
                    Account(i, 4) = Account(i, 4) - iDrMoney
                    
                    Sheet3.Cells(yOutput, 1) = iYear
                    Sheet3.Cells(yOutput, 2) = sCrAccountNo
                    Sheet3.Cells(yOutput, 3) = sSummary
                    Sheet3.Cells(yOutput, 4) = ""
                    Sheet3.Cells(yOutput, 5) = iDrMoney
                    If Account(i, 4) < 0 Then
                        Sheet3.Cells(yOutput, 6) = "貸"
                    Else
                        Sheet3.Cells(yOutput, 6) = "借"
                    End If
                    Sheet3.Cells(yOutput, 7) = Abs(Account(i, 4))
                    Sheet3.Range(Cells(yOutput, 1), Cells(yOutput, 7)).Borders.LineStyle = xlContinuous
                    yOutput = yOutput + 1
                
                    '補助科目コードに何か書かれているかチェック
                    If IsNull(iDrSubAccountNo) Or IsEmpty(iDrSubAccountNo) Then
                    Else
                        iSubAccountNo = iDrSubAccountNo
                        sSubAccountNo = sDrSubAccountNo
                    End If
                
                End If
                
                If (iSubAccountNo <> 0) Then
                    '既に既出の補助科目かチェック
                    'すでに補助科目があるか検索
                    iSub = stCntSubAccount
                    fThereIs = True
                    While (iSub < cntSubAccount)
                        If (SubAccount(iSub, 0) = iSubAccountNo) Then
                            fThereIs = False
                        End If
                        iSub = iSub + 1
                    Wend
                    '補助科目を処理用配列に追加
                    If fThereIs = True Then
                        SubAccount(cntSubAccount, 0) = iSubAccountNo
                        SubAccount(cntSubAccount, 1) = 0    '前期繰越
                        SubAccount(cntSubAccount, 5) = sSubAccountNo
                        cntSubAccount = cntSubAccount + 1
                    End If
                End If
                
                yInput = yInput + 1
            Loop
            
            '今期実績への出力
            Sheet11.Cells(yOutPerformance, 1) = Account(i, 0)
            Sheet11.Cells(yOutPerformance, 2) = Account(i, 7)
            Sheet11.Cells(yOutPerformance, 3) = ""
            Sheet11.Cells(yOutPerformance, 4) = ""
            Sheet11.Cells(yOutPerformance, 5) = Account(i, 4)
            yOutPerformance = yOutPerformance + 1
            
            ' NOTE:
            ' 正味財産の区分は、将来は「仕訳の区分（一般／指定）」で判定する想定
            ' 現在は勘定コード帯で代用
            If Account(i, 0) >= 40000 And Account(i, 0) < 80000 Then
                '一般正味財産
                iNetAssets_Diff = iNetAssets_Diff - Account(i, 4)
            ElseIf Account(i, 0) >= 80000 And Account(i, 0) < 100000 Then
                '指定正味財産
                iSpNetAssets_Diff = iSpNetAssets_Diff - Account(i, 4)
            End If

            '財務諸表への出力
            Call OutFinancialStatements(Account(i, 0), iThisX, Account(i, 4))
            
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
            Call QuickSort(SubAccount(), 5, 0, Account(i, 5), Account(i, 6))

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
                Sheet4.Cells(ySubOutput, 1).Font.Size = 16
                Sheet4.Cells(ySubOutput, 1).HorizontalAlignment = xlCenter
                
                ySubOutput = ySubOutput + 1
                
                Sheet4.Range(Cells(ySubOutput, 1), Cells(ySubOutput, 7)).Merge
                Sheet4.Cells(ySubOutput, 1) = "補助科目：" & SubAccount(iSub, 5)
                Sheet4.Cells(ySubOutput, 1).Font.Underline = True
                Sheet4.Cells(ySubOutput, 1).Font.Size = 14
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
                
                '仕訳帳の先頭Ｙ座標
                yInput = styJournal
                
                Do
                    '仕訳帳を読み込み
                    ReadJournal (yInput)
                    
                    '年月日の記載が無かったら、検索終了
                    If IsNull(iYear) Or IsEmpty(iYear) Then Exit Do
                   
                    '借方に集計中の勘定科目が記載されていた場合
                    If (Account(i, 0) = iCrAccountNo) And (SubAccount(iSub, 0) = iCrSubAccountNo) Then
                        SubAccount(iSub, 2) = SubAccount(iSub, 2) + iCrMoney
                        SubAccount(iSub, 4) = SubAccount(iSub, 4) + iCrMoney
                        
                        Sheet4.Cells(ySubOutput, 1) = iYear
                        Sheet4.Cells(ySubOutput, 2) = sDrAccountNo
                        Sheet4.Cells(ySubOutput, 3) = sSummary
                        Sheet4.Cells(ySubOutput, 4) = iCrMoney
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
                    If (Account(i, 0) = iDrAccountNo) And (SubAccount(iSub, 0) = iDrSubAccountNo) Then
                        SubAccount(iSub, 3) = SubAccount(iSub, 3) + iDrMoney
                        SubAccount(iSub, 4) = SubAccount(iSub, 4) - iDrMoney
                        
                        Sheet4.Cells(ySubOutput, 1) = iYear
                        Sheet4.Cells(ySubOutput, 2) = sCrAccountNo
                        Sheet4.Cells(ySubOutput, 3) = sSummary
                        Sheet4.Cells(ySubOutput, 4) = ""
                        Sheet4.Cells(ySubOutput, 5) = iDrMoney
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
                Loop
                
                '今期実績への出力
                Sheet11.Cells(yOutPerformance, 1) = Account(i, 0)
                Sheet11.Cells(yOutPerformance, 2) = Account(i, 7)
                Sheet11.Cells(yOutPerformance, 3) = SubAccount(iSub, 0)
                Sheet11.Cells(yOutPerformance, 4) = SubAccount(iSub, 5)
                Sheet11.Cells(yOutPerformance, 5) = SubAccount(iSub, 4)
                yOutPerformance = yOutPerformance + 1
                
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
    Call OutFinancialStatements(idNetAssets_End, iThisX, iNetAssets_End)
    Call OutFinancialStatements(idNetAssets_Begin, iThisX, iNetAssets_Begin)
    Call OutFinancialStatements(idNetAssets_Diff, iThisX, iNetAssets_Diff)
    Call OutFinancialStatements(idSpNetAssets_End, iThisX, iSpNetAssets_End)
    Call OutFinancialStatements(idSpNetAssets_Begin, iThisX, iSpNetAssets_Begin)
    Call OutFinancialStatements(idSpNetAssets_Diff, iThisX, iSpNetAssets_Diff)

    '今期実績への出力
    Sheet11.Cells(yOutPerformance, 1) = idNetAssets_Diff
    Sheet11.Cells(yOutPerformance, 2) = "当期一般正味財産増減額"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = iNetAssets_Diff
    yOutPerformance = yOutPerformance + 1
    
    Sheet11.Cells(yOutPerformance, 1) = idNetAssets_Begin
    Sheet11.Cells(yOutPerformance, 2) = "一般正味財産期首残高"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = iNetAssets_Begin
    yOutPerformance = yOutPerformance + 1
            
    Sheet11.Cells(yOutPerformance, 1) = idNetAssets_End
    Sheet11.Cells(yOutPerformance, 2) = "一般正味財産期末残高"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = iNetAssets_End
    yOutPerformance = yOutPerformance + 1

    Sheet11.Cells(yOutPerformance, 1) = idSpNetAssets_Diff
    Sheet11.Cells(yOutPerformance, 2) = "当期指定正味財産増減額"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = iSpNetAssets_Diff
    yOutPerformance = yOutPerformance + 1
    
    Sheet11.Cells(yOutPerformance, 1) = idSpNetAssets_Begin
    Sheet11.Cells(yOutPerformance, 2) = "指定正味財産期首残高"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = iSpNetAssets_Begin
    yOutPerformance = yOutPerformance + 1
            
    Sheet11.Cells(yOutPerformance, 1) = idSpNetAssets_End
    Sheet11.Cells(yOutPerformance, 2) = "指定正味財産期末残高"
    Sheet11.Cells(yOutPerformance, 3) = ""
    Sheet11.Cells(yOutPerformance, 4) = ""
    Sheet11.Cells(yOutPerformance, 5) = iSpNetAssets_End
    yOutPerformance = yOutPerformance + 1

End Sub
