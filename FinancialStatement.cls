VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FinancialStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************
'*      定数定義                        *
'****************************************
Private Const MaxColumn As Long = 5                             '出力する最大列（E列まで）
Private Const START_ROW As Long = 5                             'ヘッダー開始行
Private Const MASTER_SHEET_NAME As String = "財務諸表マスタ"    'マスタシート名

'****************************************
'*      メンバー変数                    *
'****************************************
Private m_Master As FinancialStatementMaster                    '財務諸表マスタクラス

'==============================================================================
'       初期化処理
'------------------------------------------------------------------------------
'   Contents:
'       クラス生成時に財務諸表マスタクラスを初期化する
'
'==============================================================================
Private Sub Class_Initialize()
    Set m_Master = New FinancialStatementMaster
End Sub

'==============================================================================
'       キー生成
'------------------------------------------------------------------------------
'   Contents:
'       lineID と 会計区分(category) を元にDictionary用キーを生成する
'       categoryが空の場合は "*"（全体）として扱う
'
'   Input:
'       lineID     財務諸表行ID
'       category   会計区分
'
'   Output:
'       Dictionaryキー文字列
'
'==============================================================================
Private Function BuildKey(ByVal lineID As String, ByVal category As String) As String
    If Trim$(category) = "" Then category = "*"
    BuildKey = UCase$(Trim$(lineID)) & "|" & UCase$(Trim$(category))
End Function

'==============================================================================
'       財務諸表の生成～出力
'------------------------------------------------------------------------------
'   Contents:
'       ・マスタ読込
'       ・試算表→財務諸表用データ変換
'       ・純資産データ追加
'       ・集計処理（2段階）
'       ・計算式解決
'       ・Excelへ出力
'
'   Input:
'       tbEngine        試算表クラス
'
'==============================================================================
Public Sub BuildAndOutput(ByVal tbEngine As TrialBalanceEngine)

    Dim wsMaster As Worksheet
    Set wsMaster = ThisWorkbook.Worksheets(MASTER_SHEET_NAME)
    
    'マスタ読込
    Call m_Master.LoadFromWorksheet(wsMaster)

    Dim thisAmounts As Object
    Dim lastAmounts As Object

    '試算表→財務諸表マップ生成
    Set thisAmounts = BuildStatementAmountMap(tbEngine.SubTrialBalance)
    Set lastAmounts = BuildStatementAmountMap(tbEngine.LastSubTrialBalance)
    
    '純資産データ追加
    Call AddNetAssetAmounts(thisAmounts, tbEngine, False)
    Call AddNetAssetAmounts(lastAmounts, tbEngine, True)

    '① 勘定科目の集計
    Call RollupAggregateTo(thisAmounts, False)
    Call RollupAggregateTo(lastAmounts, False)

    '② 計算式の解決
    Call ResolveFormulaAmounts(thisAmounts)
    Call ResolveFormulaAmounts(lastAmounts)

    '③ 計算式を含めた集計
    Call RollupAggregateTo(thisAmounts, True)
    Call RollupAggregateTo(lastAmounts, True)
        
    '出力
    Call OutputAll(thisAmounts, lastAmounts)

End Sub

'==============================================================================
'       試算表 → 財務諸表金額マップ生成
'------------------------------------------------------------------------------
'   Contents:
'       試算表の各行を財務諸表マスタのFS_IDへマッピングし、
'       Dictionary(lineID|category → 金額)を生成する
'
'   Input:
'       tb      試算表
'
'   Output:
'       Dictionary
'
'==============================================================================
Private Function BuildStatementAmountMap(ByVal tb As TrialBalance) As Object

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim line As TrialBalanceLine
    Dim lineIDs As Variant
    Dim lineID As String
    Dim accountingCategory As String
    Dim amount As Currency
    Dim i As Long
    Dim mapped As Boolean
    
    For Each line In tb.Lines
        amount = line.amount()
        mapped = False

        '勘定科目に紐づくFS_IDを分解
        If Trim$(line.account.fsLineIDs) <> "" Then
            lineIDs = Split(line.account.fsLineIDs, ";")
            For i = LBound(lineIDs) To UBound(lineIDs)
                lineID = UCase$(Trim$(CStr(lineIDs(i))))
                If lineID <> "" Then
                    
                    'マスタ存在チェック
                    If Not ExistsLineID(lineID) Then
                        Err.Raise vbObjectError + 7001, _
                            "BuildStatementAmountMap", _
                            "Undefined FS_ID mapping in account master: " & lineID
                    End If

                    mapped = True
                    If amount <> 0 Then
                        accountingCategory = UCase$(Trim$(line.SubAccountCate))
                        '全体(*)への加算
                        Call AddAmountToMap(result, lineID, "*", amount)
                        '区分別加算
                        If accountingCategory <> "" Then
                            Call AddAmountToMap(result, lineID, accountingCategory, amount)
                        End If
                    End If
                End If
            Next i
        End If

        '未マッピングエラー
        If Not mapped Then
            Err.Raise vbObjectError + 7002, _
                "BuildStatementAmountMap", _
                "TrialBalanceLine is not mapped to FS_ID. AccountCode=" & CStr(line.AccountCode)
        End If
    Next

    Set BuildStatementAmountMap = result

End Function

'==============================================================================
'       FS_ID存在チェック
'------------------------------------------------------------------------------
'   Contents:
'       指定されたlineIDがマスタに存在するか確認する
'
'==============================================================================
Private Function ExistsLineID(ByVal lineID As String) As Boolean

    Dim line As FinancialStatementMasterLine
    Dim targetID As String

    targetID = UCase$(Trim$(lineID))

    For Each line In m_Master.Lines
        If UCase$(Trim$(line.lineID)) = targetID Then
            ExistsLineID = True
            Exit Function
        End If
    Next

    ExistsLineID = False

End Function

'==============================================================================
'       金額加算処理
'------------------------------------------------------------------------------
'   Contents:
'       Dictionaryに対してキー単位で金額を加算する
'
'==============================================================================
Private Sub AddAmountToMap( _
    ByVal accountMap As Object, _
    ByVal lineID As String, _
    ByVal accountingCategory As String, _
    ByVal amount As Currency)

    Dim key As String
    key = UCase$(Trim$(lineID)) & "|" & UCase$(Trim$(accountingCategory))

    If Not accountMap.Exists(key) Then
        accountMap.Add key, CCur(0)
    End If
    accountMap(key) = CCur(accountMap(key)) + amount
        
    Debug.Print key & amount

End Sub

'==============================================================================
'       純資産（期首・期末・増減）のマップ追加
'------------------------------------------------------------------------------
'   Contents:
'       ・貸借対照表用（期末）
'       ・注記用（期首・増減）
'       の両方で使えるように、純資産をフルセットで投入する
'
'       TrialBalanceEngine が保持する純資産情報を、
'       財務諸表マスタの指定IDへ直接マッピングする
'
'   Input:
'       accountMap     出力用Dictionary（lineID|category → 金額）
'       tbEngine       試算表エンジン
'       isLastYear     前年データかどうか
'
'==============================================================================
Private Sub AddNetAssetAmounts( _
    ByVal accountMap As Object, _
    ByVal tbEngine As TrialBalanceEngine, _
    ByVal isLastYear As Boolean)

    If isLastYear Then

        '==============================
        ' 一般純資産
        '==============================
        Call AddAmountToMap(accountMap, "BS_NET_GENERAL_BEGIN", "*", tbEngine.LastGeneralNetBegin)
        Call AddAmountToMap(accountMap, "BS_NET_GENERAL_END", "*", tbEngine.LastGeneralNetEnd)
        Call AddAmountToMap(accountMap, "BS_NET_GENERAL_DIFF", "*", tbEngine.LastGeneralNetChange)

        '==============================
        ' 指定純資産
        '==============================
        Call AddAmountToMap(accountMap, "BS_NET_DESIG_BEGIN", "*", tbEngine.LastDesignatedNetBegin)
        Call AddAmountToMap(accountMap, "BS_NET_DESIG_END", "*", tbEngine.LastDesignatedNetEnd)
        Call AddAmountToMap(accountMap, "BS_NET_DESIG_DIFF", "*", tbEngine.LastDesignatedNetChange)

    Else

        '==============================
        ' 一般純資産
        '==============================
        Call AddAmountToMap(accountMap, "BS_NET_GENERAL_BEGIN", "*", tbEngine.GeneralNetAssetBegin)
        Call AddAmountToMap(accountMap, "BS_NET_GENERAL_END", "*", tbEngine.GeneralNetAssetEnd)
        Call AddAmountToMap(accountMap, "BS_NET_GENERAL_DIFF", "*", tbEngine.GeneralNetAssetChange)

        '==============================
        ' 指定純資産
        '==============================
        Call AddAmountToMap(accountMap, "BS_NET_DESIG_BEGIN", "*", tbEngine.DesignatedNetAssetBegin)
        Call AddAmountToMap(accountMap, "BS_NET_DESIG_END", "*", tbEngine.DesignatedNetAssetEnd)
        Call AddAmountToMap(accountMap, "BS_NET_DESIG_DIFF", "*", tbEngine.DesignatedNetAssetChange)

    End If

End Sub

'==============================================================================
'       財務諸表出力
'------------------------------------------------------------------------------
'   Contents:
'       財務諸表マスタの定義順に従い、各シートへ出力する
'       ・当期／前期／増減を算出
'       ・表示可否判定
'       ・罫線／ゼブラ装飾
'
'   Input:
'       thisAmounts    当期金額マップ
'       lastAmounts    前期金額マップ
'
'==============================================================================
Private Sub OutputAll(ByVal thisAmounts As Object, ByVal lastAmounts As Object)

    Dim line As FinancialStatementMasterLine
    
    Dim vThis As Currency
    Dim vLast As Currency
    Dim vDiff As Currency

    'シート毎の出力行管理（SheetName → 次行）
    Dim nextRowBySheet As Object
    Set nextRowBySheet = CreateObject("Scripting.Dictionary")

    'シート初期化
    Call InitializeOutputSheets(nextRowBySheet)

    'マスタ順に出力
    For Each line In m_Master.Lines
    
        If line.enabled Then
    
            '金額取得
            vThis = GetAmountByLine(thisAmounts, line)
            vLast = GetAmountByLine(lastAmounts, line)
            vDiff = vThis - vLast
            
            '------------------------------------------
            ' シート出力
            '------------------------------------------
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Worksheets(line.TargetSheet)
    
            Dim y As Long
            y = CLng(nextRowBySheet(ws.Name))
            
            '科目名出力
            ws.Cells(y, 2) = line.LineLabel
            ws.Cells(y, 2).HorizontalAlignment = line.LineLabelAligh
    
            '------------------------------------------
            ' 表示制御
            '------------------------------------------
            If Not ShouldSkipOutputLine(line, vThis, vLast) Then
                
                'タイトル行は金額を表示しない
                If line.LineType = "TITLE" Then
                    ws.Cells(y, 3).value = ""
                    ws.Cells(y, 4).value = ""
                    ws.Cells(y, 5).value = ""
                Else
                    ws.Cells(y, 3).value = vThis
                    ws.Cells(y, 4).value = vLast
                    ws.Cells(y, 5).value = vDiff
                End If
                                    
                '------------------------------------------
                ' 見た目調整
                '------------------------------------------
                '横罫線（帳簿形式）
                ws.Range(ws.Cells(y, 2), ws.Cells(y, MaxColumn)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
                'ゼブラ（偶数行）
                If y Mod 2 = 0 Then
                    ws.Range(ws.Cells(y, 2), ws.Cells(y, MaxColumn)).Interior.Color = rgbAliceBlue
                End If

                '次行へ
                nextRowBySheet(ws.Name) = y + 1
            End If
        End If
    Next

End Sub

'==============================================================================
'       出力シート初期化
'------------------------------------------------------------------------------
'   Contents:
'       ・対象シートのクリア
'       ・ヘッダー設定
'       ・シート毎の出力開始行を初期化
'
'==============================================================================
Private Sub InitializeOutputSheets(ByVal nextRowBySheet As Object)

    Dim line As FinancialStatementMasterLine
    Dim ws As Worksheet

    For Each line In m_Master.Lines
        If line.enabled Then
            
            'シート単位で一度だけ初期化
            If Not nextRowBySheet.Exists(line.TargetSheet) Then
            
                Set ws = ThisWorkbook.Worksheets(line.TargetSheet)
                
                '出力開始行設定
                nextRowBySheet.Add line.TargetSheet, START_ROW + 1
                
                '既存データクリア
                Call ClearSheet(ws, START_ROW)
                
                'ヘッダー設定
                Call SetupSheetHeader(ws, line.statementID)
                
            End If
        End If
    Next

End Sub

'==============================================================================
'       シートヘッダー設定
'------------------------------------------------------------------------------
'   Contents:
'       ・列見出し（科目／当期／前期／増減）
'       ・数値フォーマット
'       ・罫線／背景色
'
'==============================================================================
Private Sub SetupSheetHeader(ByVal ws As Worksheet, ByVal statementID As String)

    '数値フォーマット（△表示）
    ws.Range("C:E").NumberFormat = "_ * #,##0_ ;_ * △#,##0_ ;_ * ""-""??_ ;_ @_ "
    ws.Range("C:E").HorizontalAlignment = xlRight

    'ヘッダー項目
    ws.Cells(START_ROW, 2) = "科目"
    ws.Cells(START_ROW, 3) = "当期"
    ws.Cells(START_ROW, 4) = "前期"
    ws.Cells(START_ROW, 5) = "増減"

    'ヘッダー装飾
    With ws.Range(ws.Cells(START_ROW, 2), ws.Cells(START_ROW, MaxColumn))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = rgbLavender
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With

End Sub

'==============================================================================
'       表示可否判定
'------------------------------------------------------------------------------
'   Contents:
'       行を出力するか判定する
'       ・TITLE は常に表示
'       ・ACCOUNT系は前期・当期とも0なら非表示
'
'==============================================================================
Private Function ShouldSkipOutputLine(ByVal line As FinancialStatementMasterLine, ByVal vThis As Currency, ByVal vLast As Currency) As Boolean

    If line.LineType = "TITLE" Or line.LineType = "ACCOUNTS" Then
        ShouldSkipOutputLine = False
        Exit Function
    End If

    'ゼロ行非表示
    If (line.LineType = "ACCOUNT" Or line.LineType = "ACCOUNTS") And vThis = 0 And vLast = 0 Then
        ShouldSkipOutputLine = True
        Exit Function
    End If
       
    ShouldSkipOutputLine = False  'デフォルトは表示
    
End Function

'==============================================================================
'       金額取得
'------------------------------------------------------------------------------
'   Contents:
'       lineID + category に対応する金額をDictionaryから取得する
'       存在しない場合は0を返す
'
'==============================================================================
Private Function GetAmountByLine(ByVal amountMap As Object, ByVal line As FinancialStatementMasterLine) As Currency

    Dim key As String
    key = BuildKey(line.lineID, line.accountingCategory)

    If amountMap.Exists(key) Then
        GetAmountByLine = CCur(amountMap(key))
    Else
        GetAmountByLine = 0
    End If

End Function


'==============================================================================
'       計算式行の金額解決
'------------------------------------------------------------------------------
'   Contents:
'       LineType = FORMULA / CALC の行について、
'       定義された数式を評価し、Dictionaryへ格納する
'
'       ※前提：
'           ・事前にACCOUNT系の集計が完了していること
'
'==============================================================================
Private Sub ResolveFormulaAmounts(ByVal values As Object)

    Dim line As FinancialStatementMasterLine
    Dim key As String

    For Each line In m_Master.Lines
        If line.enabled Then
            If line.LineType = "FORMULA" Or line.LineType = "CALC" Then
                'キー生成（lineID|category）
                key = BuildKey(line.lineID, line.accountingCategory)
                '数式評価
                values(key) = ResolveFormula(line.formula, values)
            End If
        End If
    Next

End Sub

'==============================================================================
'       集計処理（Rollup）
'------------------------------------------------------------------------------
'   Contents:
'       AggregateTo に基づき、子→親へ金額を加算する
'
'       includeFormula により対象を切替：
'           False : ACCOUNT系のみ集計（第1段階）
'           True  : FORMULA/CALCも含める（第2段階）
'
'==============================================================================
Private Sub RollupAggregateTo(ByVal amountMap As Object, ByVal includeFormula As Boolean)

    Dim line As FinancialStatementMasterLine
    Dim childType As String

    For Each line In m_Master.Lines
        If line.enabled Then
            childType = UCase$(Trim$(line.LineType))

            'タイトル行は対象外、かつ親が定義されている場合のみ
            If childType <> "TITLE" And Trim$(line.AggregateTo) <> "" Then
                
                '------------------------------------------
                ' 対象判定（2段階集計）
                '------------------------------------------
                If (includeFormula And (childType = "FORMULA" Or childType = "CALC")) _
                    Or ((Not includeFormula) And (childType = "ACCOUNT" Or childType = "ACCOUNTS")) Then
                    Call RollupLineAmount(line, amountMap)
                End If
            End If
        End If
    Next

End Sub

'==============================================================================
'       1行分の集計処理
'------------------------------------------------------------------------------
'   Contents:
'       子行の金額を取得し、AggregateTo で指定された親行へ加算する
'
'       ・子が存在しない場合は0として扱う
'       ・親が未初期化の場合は0で初期化
'
'==============================================================================
Private Sub RollupLineAmount(ByVal child As FinancialStatementMasterLine, ByVal amountMap As Object)

    Dim parentLine As FinancialStatementMasterLine
    Dim childKey As String
    Dim parentKey As String
    Dim childAmount As Currency

    '------------------------------------------
    ' 子の金額取得
    '------------------------------------------
    childKey = BuildKey(child.lineID, child.accountingCategory)
    If amountMap.Exists(childKey) Then
        childAmount = CCur(amountMap(childKey))
    Else
        childAmount = 0
    End If

    '親未定義なら終了
    If Trim$(child.AggregateTo) = "" Then
        Exit Sub
    End If

    '------------------------------------------
    ' 親行取得
    '------------------------------------------
    Set parentLine = GetAggregateParentLine(child.AggregateTo, child.accountingCategory)
    parentKey = BuildKey(parentLine.lineID, parentLine.accountingCategory)

    '------------------------------------------
    ' 親へ加算
    '------------------------------------------
    If Not amountMap.Exists(parentKey) Then
        amountMap.Add parentKey, CCur(0)
    End If
    
    amountMap(parentKey) = CCur(amountMap(parentKey)) + childAmount

End Sub

'==============================================================================
'       集計先（親行）取得
'------------------------------------------------------------------------------
'   Contents:
'       AggregateTo で指定された親行を取得する
'
'       優先順位：
'           ① 同一category
'           ② "*"（全体）
'
'       ※存在しない場合はエラー
'
'==============================================================================
Private Function GetAggregateParentLine(ByVal parentLineID As String, ByVal category As String) As FinancialStatementMasterLine

    Dim normalizedParent As String
    normalizedParent = UCase$(Trim$(parentLineID))

    '未指定チェック
    If normalizedParent = "" Then
        Err.Raise vbObjectError + 7102, _
            "GetAggregateParentLine", _
            "AggregateTo target is blank."
    End If

    '① category一致
    If m_Master.Exists(normalizedParent, category) Then
        Set GetAggregateParentLine = m_Master.GetLine(normalizedParent, category)
        Exit Function
    End If

    '② 全体("*")
    If m_Master.Exists(normalizedParent, "*") Then
        Set GetAggregateParentLine = m_Master.GetLine(normalizedParent, "*")
        Exit Function
    End If

    '存在しない場合エラー
    Err.Raise vbObjectError + 7103, _
        "GetAggregateParentLine", _
        "AggregateTo target does not exist: " & normalizedParent & "|" & UCase$(Trim$(category))

End Function

'==============================================================================
'       数式評価処理（コアロジック）
'------------------------------------------------------------------------------
'   Contents:
'       数式内の [FS_ID|category] を実際の金額に置換し、
'       Evaluate で計算する
'
'       対応形式：
'           [PL_BIZ]
'           [PL_BIZ|公1]
'
'       解決ロジック：
'           ① 完全一致（lineID|category）
'           ② "*" の場合は同一lineIDの合算
'           ③ 見つからない場合はエラー
'
'==============================================================================
Private Function ResolveFormula(ByVal formula As String, ByVal values As Object) As Currency

    '空なら0
    If Trim$(formula) = "" Then
        ResolveFormula = 0
        Exit Function
    End If

    Dim expr As String
    expr = formula

    '------------------------------------------
    ' トークン抽出（[xxx]）
    '------------------------------------------
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    With re
        .Pattern = "\[([^\]]+)\]"
        .Global = True
    End With

    Dim matches As Object
    Set matches = re.Execute(expr)

    Dim m As Object
    For Each m In matches
        
        Dim token As String
        token = m.SubMatches(0)
        
        Dim lineID As String
        Dim category As String
        
        '------------------------------------------
        ' [PL_BIZ|公1] or [PL_BIZ]
        '------------------------------------------
        If InStr(token, "|") > 0 Then
            lineID = UCase$(Trim$(Split(token, "|")(0)))
            category = UCase$(Trim$(Split(token, "|")(1)))
        Else
            lineID = UCase$(Trim$(token))
            category = "*"
        End If

        If category = "" Then category = "*"

        '------------------------------------------
        ' マスタ存在チェック
        '------------------------------------------
        If category = "*" Then
            If Not ExistsLineID(lineID) Then
                Err.Raise vbObjectError + 3002, _
                    "ResolveFormula", _
                    "Undefined FS_ID in formula: [" & token & "]"
            End If
        Else
            If Not m_Master.Exists(lineID, category) Then
                Err.Raise vbObjectError + 3003, _
                    "ResolveFormula", _
                    "Undefined FS_ID/category in formula: [" & token & "]"
            End If
        End If
        
        Dim key As String
        key = BuildKey(lineID, category)
        
        Dim val As Currency
        
        '------------------------------------------
        ' ① 完全一致
        '------------------------------------------
        If values.Exists(key) Then
            val = values(key)
        
        '------------------------------------------
        ' ② "*"（合算）
        '------------------------------------------
        ElseIf category = "*" Then
            
            Dim k As Variant
            val = 0
            
            For Each k In values.Keys
                If Left$(UCase$(CStr(k)), Len(UCase$(lineID)) + 1) = UCase$(lineID) & "|" Then
                    val = val + values(k)
                End If
            Next
        
        '------------------------------------------
        ' ③ エラー
        '------------------------------------------
        Else
            Err.Raise vbObjectError + 3001, _
                      "ResolveFormula", _
                      "キーが存在しません: [" & token & "]"
        End If
        
        '数式文字列に置換
        expr = Replace(expr, "[" & token & "]", CStr(val))
        
    Next

    'スペース除去
    expr = Replace(expr, " ", "")
    
    '最終評価
    ResolveFormula = CCur(Application.Evaluate(expr))

End Function

