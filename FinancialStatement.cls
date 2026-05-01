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
Private Const START_ROW As Long = 2                             'ヘッダー開始行
Private Const MASTER_SHEET_NAME As String = "財務諸表マスタ"    'マスタシート名

'****************************************
'*      メンバー変数                    *
'****************************************
Private m_AccMaster As AccountMaster                            '勘定科目マスタ
Private m_Master As FinancialStatementMaster                    '財務諸表マスタ管理クラス
Private m_ColumnKeys As FinancialStatementColumnKeys            '出力列キー管理クラス
Private m_thisAmounts As FinancialAmountMap                     '今期財務諸表
Private m_lastAmounts As FinancialAmountMap                     '前期財務諸表

'==============================================================================
'       初期化処理
'------------------------------------------------------------------------------
'   Contents:
'       FinancialStatement クラス生成時の初期化処理。
'
'       財務諸表マスタ管理クラス、および
'       出力列キー管理クラスを生成する。
'
'       ※ 実データの読込は BuildAndOutput 実行時に行う。
'
'==============================================================================
Private Sub Class_Initialize()
    Set m_Master = New FinancialStatementMaster
    Set m_ColumnKeys = New FinancialStatementColumnKeys
    Set m_thisAmounts = New FinancialAmountMap
    Set m_lastAmounts = New FinancialAmountMap
End Sub

'==============================================================================
'       勘定科目マスタ設定
'------------------------------------------------------------------------------
'   Contents:
'       財務諸表生成時に使用する勘定科目マスタを設定する。
'
'       設定された勘定科目マスタは、
'       集計処理・表示制御・勘定判定等で使用される。
'
'   Input:
'       AccMaster     勘定科目マスタクラス
'
'==============================================================================
Public Sub Set_Master(AccMaster As AccountMaster)
    Set m_AccMaster = AccMaster
End Sub

'==============================================================================
'       財務諸表生成メイン処理
'------------------------------------------------------------------------------
'   Contents:
'       試算表データをもとに、財務諸表用データを生成し、
'       各帳票をExcelへ出力する。
'
'       本クラスにおけるメインエントリポイントであり、
'       以下の流れで処理を実行する。
'
'       ■ 処理フロー
'           ① 財務諸表マスタ読込
'           ② 出力列キー初期化
'           ③ 試算表 → 財務諸表金額マップ変換
'           ④ 純資産データ補完
'           ⑤ 勘定科目ベース集計
'           ⑥ 計算式解決
'           ⑦ 計算式反映後の再集計
'           ⑧ 財務諸表出力
'
'       ■ 設計ポイント
'           ・実データ集計と計算式集計を分離している
'           ・計算式解決後に再集計を行う構造としている
'           ・将来的な内部取引消去、配賦、会計振替等の
'             拡張ポイントを確保している
'
'   Input:
'       tbEngine       試算表生成済みエンジンクラス
'
'==============================================================================
Public Sub BuildAndOutput(ByVal tbEngine As TrialBalanceEngine)

    Dim wsMaster As Worksheet
    Set wsMaster = ThisWorkbook.Worksheets(MASTER_SHEET_NAME)
    
    'マスタ読込
    Call m_Master.LoadFromWorksheet(wsMaster)

    '会計区分・補助科目抽出
    Call m_ColumnKeys.Initialize(tbEngine.SubTrialBalance)
    
    '試算表→財務諸表マップ生成
    Call BuildStatementAmountMap(m_thisAmounts, tbEngine.SubTrialBalance)
    Call BuildStatementAmountMap(m_lastAmounts, tbEngine.LastSubTrialBalance)
    
    '純資産データ追加
    Call AddNetAssetAmounts(m_thisAmounts, tbEngine, False)
    Call AddNetAssetAmounts(m_lastAmounts, tbEngine, True)

    '① 勘定科目の集計
    Call RollupAggregateTo(m_thisAmounts, False)
    Call RollupAggregateTo(m_lastAmounts, False)

    '② 計算式の解決
    Call ResolveFormulaAmounts(m_thisAmounts)
    Call ResolveFormulaAmounts(m_lastAmounts)

    '③ 計算式を含めた集計
    Call RollupAggregateTo(m_thisAmounts, True)
    Call RollupAggregateTo(m_lastAmounts, True)
        
    '出力
    Call OutputAll(m_thisAmounts, m_lastAmounts)

End Sub

'==============================================================================
'       財務諸表金額マップ生成
'------------------------------------------------------------------------------
'   Contents:
'       試算表データから、財務諸表用の金額マップを生成する。
'
'       試算表の各行について、
'
'           勘定科目コード
'           会計区分
'
'       をもとに、財務諸表マスタ上の対応行へ金額を加算する。
'
'       生成される Dictionary は、
'
'           Key   : BuildKey(lineID, category)
'           Value : 金額
'
'       の形式で保持される。
'
'       ■ 処理概要
'           ① 試算表を1行ずつ走査
'           ② 勘定科目コードから財務諸表行を取得
'           ③ 会計区分別キーを生成
'           ④ 金額を加算
'
'       ■ 注意
'           本処理では、まだ Aggregate 集計や
'           計算式解決は行わない。
'
'           あくまで「元データ投入フェーズ」である。
'
'   Input:
'       subTB       補助元帳レベル試算表
'
'   Output:
'       amounts     財務諸表金額マップ
'
'==============================================================================
Private Sub BuildStatementAmountMap(ByVal accountMap As FinancialAmountMap, ByVal tb As TrialBalance)

    Dim tbLine As TrialBalanceLine

    For Each tbLine In tb.lines

        Call ApplyMasterLineMatching(accountMap, tbLine)
    
    Next tbLine

End Sub

'==============================================================================
'       純資産データ追加
'------------------------------------------------------------------------------
'   Contents:
'       試算表には存在しない純資産関連項目について、
'       財務諸表表示用の金額を追加する。
'
'       主に、
'
'           ・一般正味財産
'           ・指定正味財産
'
'       の期首・期末残高等を補完する目的で使用する。
'
'       ■ 設計意図
'           純資産系は通常の収益費用集計とは
'           性質が異なるため、
'           BuildStatementAmountMap とは分離している。
'
'       ■ 注意
'           本処理は財務諸表表示用データ補完であり、
'           TrialBalance 自体を書き換えるものではない。
'
'   Input:
'       amounts         財務諸表金額マップ
'       tbEngine        試算表エンジン
'       isLastYear      前年度データかどうか
'
'==============================================================================
Private Sub AddNetAssetAmounts( _
    ByVal accountMap As FinancialAmountMap, _
    ByVal tbEngine As TrialBalanceEngine, _
    ByVal isLastYear As Boolean)

    If isLastYear Then

        '==============================
        ' 一般純資産
        Call AddAmountToMatchedMasterLines(accountMap, idNetAssets_Begin, natGeneral, tbEngine.LastGeneralNetBegin)
        Call AddAmountToMatchedMasterLines(accountMap, idNetAssets_End, natGeneral, tbEngine.LastGeneralNetEnd)
 
        '==============================
        ' 指定純資産
        Call AddAmountToMatchedMasterLines(accountMap, idNetAssets_Begin, natDesignated, tbEngine.LastDesignatedNetBegin)
        Call AddAmountToMatchedMasterLines(accountMap, idNetAssets_End, natDesignated, tbEngine.LastDesignatedNetEnd)
        
    Else

        '==============================
        ' 一般純資産
        Call AddAmountToMatchedMasterLines(accountMap, idNetAssets_Begin, natGeneral, tbEngine.GeneralNetAssetBegin)
        Call AddAmountToMatchedMasterLines(accountMap, idNetAssets_End, natGeneral, tbEngine.GeneralNetAssetEnd)
 
        '==============================
        ' 指定純資産
        Call AddAmountToMatchedMasterLines(accountMap, idNetAssets_Begin, natDesignated, tbEngine.DesignatedNetAssetBegin)
        Call AddAmountToMatchedMasterLines(accountMap, idNetAssets_End, natDesignated, tbEngine.DesignatedNetAssetEnd)

    End If

End Sub

'==============================================================================
'       純資産データを試算表形式へ変換してマスタ突合せへ投入
'------------------------------------------------------------------------------
'   Contents:
'       純資産残高データを TrialBalanceLine へ変換し、
'       通常の試算表データと同じ突合せ処理へ投入する。
'
'       ■ この関数の役割
'           ・純資産データ専用の分岐処理を減らす
'           ・試算表データと同一ロジックへ統一する
'
'       ■ 処理概要
'           ① 仮想 TrialBalanceLine を生成
'           ② 勘定科目マスタから勘定取得
'           ③ 金額設定
'           ④ ApplyMasterLineMatching へ投入
'
'   Input:
'       amountMap      出力用Dictionary（lineID|category → 金額）
'       code           勘定科目コード
'       nat            純資産区分（一般／指定）
'       amount         残高
'
'==============================================================================
Private Sub AddAmountToMatchedMasterLines(ByVal amountMap As FinancialAmountMap, ByVal code As Long, ByVal nat As NetAssetType, ByVal amount As Currency)

    Dim tbLine As TrialBalanceLine
    Set tbLine = New TrialBalanceLine
    
    Call tbLine.Initialize(m_AccMaster.GetAccount(code), 0, "", "", nat)
    Call tbLine.AddAmount(0, 0, amount, amount)

    Call ApplyMasterLineMatching(amountMap, tbLine)

End Sub

'==============================================================================
'       試算表行と財務諸表マスタ行の突合せ
'------------------------------------------------------------------------------
'   Contents:
'       試算表1行（tbLine）が、財務諸表マスタのどの行へ属するか判定し、
'       一致した行へ金額を加算する。
'
'       ■ 処理概要
'           ① 財務諸表マスタ全行を走査
'           ② ACCOUNT / ACCOUNTS 行のみ対象
'           ③ Aggregate 条件判定
'           ④ Filter 条件判定
'           ⑤ 金額マップへ加算
'
'       ■ 分類処理
'           ・固定分類
'           ・EACH展開
'           ・注記用分類
'           に対応している。
'
'       ■ EACH の役割
'           補助科目や事業単位等へ、明細分解して出力するために使用する。
'
'==============================================================================
Private Sub ApplyMasterLineMatching(ByVal amountMap As FinancialAmountMap, ByVal tbLine As TrialBalanceLine)

    Dim masterLine As FinancialStatementMasterLine
    Dim detailKey As String
    Dim filterTokens As Variant
    Dim filterToken As Variant
    Dim amount As Currency
    
    amount = tbLine.amount

    '------------------------------------------
    ' 試算表1行 × マスタ全行 を突き合わせる
    '------------------------------------------
    For Each masterLine In m_Master.lines
        If masterLine.enabled Then
            If masterLine.lineType = "ACCOUNT" Or masterLine.lineType = "ACCOUNTS" Then

                'ACCOUNT / ACCOUNTS：
                '実データ（試算表）から直接金額を拾う行のみ対象
                If IsAggregateMatch(masterLine, tbLine) And IsFilterMatch(masterLine, tbLine) Then
                    '------------------------------------------
                    ' マスタ定義に基づき、この試算表行が
                    ' この財務諸表行に属するか判定
                    '------------------------------------------
                    
                    Call amountMap.AddAmount(masterLine.lineID, "*", amount)

                    If UCase$(Trim$(masterLine.filterValue)) = "EACH" Or UCase$(Trim$(masterLine.statementID)) = "PL_NOTE1" Then
                        '------------------------------------------
                        ' EACH指定：
                        ' 試算表の値ごとに明細分解（補助科目・事業など）
                        '------------------------------------------
                        If UCase$(Trim$(masterLine.statementID)) = "PL_NOTE1" Then
                            detailKey = GetFilterKey(CATEGORY_NET_ASSET, tbLine)
                        Else
                            detailKey = GetFilterKey(masterLine.filterField, tbLine)
                        End If
                        If detailKey <> "" And detailKey <> "*" Then
                            Call amountMap.AddAmount(masterLine.lineID, detailKey, amount)
                        End If
                    ElseIf UCase$(Trim$(masterLine.filterValue)) <> "*" And UCase$(Trim$(masterLine.filterValue)) <> "" Then
                        '------------------------------------------
                        ' 固定カテゴリ指定：
                        ' 例）"一般,指定" のような指定に従って振り分け
                        '------------------------------------------
                        filterTokens = Split(masterLine.filterValue, ",")
                        For Each filterToken In filterTokens
                            detailKey = UCase$(Trim$(CStr(filterToken)))
                            If detailKey <> "" Then
                                Call amountMap.AddAmount(masterLine.lineID, detailKey, amount)
                            End If
                        Next filterToken
                    End If
                End If
            End If
        End If
    Next masterLine

End Sub

'==============================================================================
'       集計対象判定（Aggregate / WHERE句）
'------------------------------------------------------------------------------
'   Contents:
'       試算表1行（tbLine）が、
'       財務諸表マスタ行（masterLine）の「集計対象」に該当するか判定する。
'
'       ■ この関数の役割
'           ・どのデータを拾うかを決める（SQLのWHERE句相当）
'           ・分類（Filter）とは独立した「母集団の決定」
'
'       ■ 判定の流れ（重要）
'           ① aggregateField を正規化（Trim + UCase）
'           ② 未指定の場合は Terget_ID をデフォルト適用
'           ③ 定数ごとに対象フィールドを切替
'           ④ CsvContains により一致判定
'
'       ■ 使用定数
'           Terget_ID         ：財務諸表IDベース
'           Terget_Sub_Class  ：中区分ベース
'           Terget_Acc_Code   ：勘定科目コードベース
'
'       ■ 注意（重要）
'           ・aggregateField 未指定時は Terget_ID が適用される
'           ・区切り文字はフィールドごとに異なる（; / ,）
'           ・未対応フィールドは False（安全側）
'
'   Input:
'       masterLine     財務諸表マスタ行
'       tbLine         試算表行
'
'   Output:
'       True           集計対象
'       False          集計対象外
'
'==============================================================================
Private Function IsAggregateMatch(ByVal masterLine As FinancialStatementMasterLine, ByVal tbLine As TrialBalanceLine) As Boolean

    Dim fieldName As String

    fieldName = UCase$(Trim$(masterLine.aggregateField))
    If fieldName = "" Then fieldName = Terget_ID

    Select Case fieldName

        Case Terget_ID
            IsAggregateMatch = CsvContains(tbLine.account.fsLineIDs, masterLine.aggregateValue, ";")

        Case Terget_Sub_Class
            IsAggregateMatch = CsvContains(tbLine.account.SubCategory, masterLine.aggregateValue, ",")

        Case Terget_Acc_Code
            IsAggregateMatch = CsvContains(tbLine.account.code, masterLine.aggregateValue, ",")

        Case Else
            IsAggregateMatch = False

    End Select

End Function

'==============================================================================
'       分類対象判定（Filter / GROUP BY前段）
'------------------------------------------------------------------------------
'   Contents:
'       試算表1行（tbLine）が、
'       財務諸表マスタ行の「分類対象」として扱われるか判定する。
'
'       ■ この関数の役割
'           ・Aggregateで抽出されたデータを
'             「分類に使うかどうか」を決める
'
'       ■ Aggregateとの違い
'           Aggregate：拾うか（WHERE）
'           Filter   ：分類するか（GROUP BY対象）
'
'       ■ 判定仕様
'           ・"" / "*" / "EACH" → 常に True
'           ・それ以外 → 実際の値と一致する場合のみ True
'
'       ■ 処理の流れ
'           ① filterValue 正規化
'           ② 特殊値（*, EACH）チェック
'           ③ GetFilterKeyで実値取得
'           ④ CsvContainsで一致判定
'
'       ■ 注意
'           ・EACH の実際の分解は別処理（ここではTrueのみ返す）
'           ・GetFilterKey が "" の場合は対象外
'
'   Input:
'       masterLine     財務諸表マスタ行
'       tbLine         試算表行
'
'   Output:
'       True           分類対象
'       False          分類対象外
'
'==============================================================================
Private Function IsFilterMatch(ByVal masterLine As FinancialStatementMasterLine, ByVal tbLine As TrialBalanceLine) As Boolean

    Dim filterValue As String
    filterValue = UCase$(Trim$(masterLine.filterValue))

    If filterValue = "" Or filterValue = "*" Or filterValue = "EACH" Then
        IsFilterMatch = True
        Exit Function
    End If

    Dim actual As String
    actual = GetFilterKey(masterLine.filterField, tbLine)

    If actual = "" Then
        IsFilterMatch = False
    Else
        IsFilterMatch = CsvContains(actual, filterValue, ",")
    End If

End Function

'==============================================================================
'       分類キー生成（GROUP BYキー）
'------------------------------------------------------------------------------
'   Contents:
'       試算表行（tbLine）から、
'       財務諸表出力用の分類キーを生成する。
'
'       ■ この関数の役割
'           ・分類単位（accountingCategory）を決定
'           ・Dictionaryキーの一部になる
'
'       ■ 重要な位置づけ
'           Aggregate → WHERE
'           Filter    → GROUP BY条件
'           本関数    → GROUP BYキー生成
'
'       ■ 分岐ロジック
'           filterField に応じて取得値を切替
'
'       ■ 注意
'           ・戻り値は最終的に UCase + Trim で正規化
'           ・"" は「分類対象外」として扱われる
'
'   Input:
'       filterField    分類項目
'       tbLine         試算表行
'
'   Output:
'       GetFilterKey   分類キー
'
'==============================================================================
Private Function GetFilterKey(ByVal filterField As String, ByVal tbLine As TrialBalanceLine) As String

    Dim key As String
    Select Case UCase$(Trim$(filterField))
        Case CATEGORY_NET_ASSET
            If tbLine.NetAssetType = natDesignated Then
                key = "指定"
            Else
                key = "一般"
            End If

        Case CATEGORY_ACCOUNT_SUM
            key = GetPurposeCategory(tbLine.subAccountCate)

        Case CATEGORY_ACCOUNT
            key = UCase$(Trim$(tbLine.subAccountCate))
            If key = "" Then key = "共通"

        Case CATEGORY_PROJECT
            key = CStr(tbLine.SubAccountCode)

        Case Else
            key = "*"
    End Select

    GetFilterKey = UCase$(Trim$(key))

End Function

'==============================================================================
'       目的区分判定
'------------------------------------------------------------------------------
'   Contents:
'       補助科目区分の先頭文字から、目的区分（公益／収益／法人）を判定する
'
'   Input:
'       subAccountCate     補助科目区分文字列
'
'   Output:
'       GROUP_PUBLIC_LABEL     公益（先頭が「公」）
'       GROUP_REVENUE_LABEL    収益（先頭が「収」）
'       GROUP_CORPORATE_LABEL  法人（上記以外）
'
'==============================================================================
Private Function GetPurposeCategory(ByVal subAccountCate As String) As String

    Dim s As String
    s = UCase$(Trim$(subAccountCate))

    If Left$(s, 1) = "公" Then
        GetPurposeCategory = GROUP_PUBLIC_LABEL
    ElseIf Left$(s, 1) = "収" Then
        GetPurposeCategory = GROUP_REVENUE_LABEL
    Else
        GetPurposeCategory = GROUP_CORPORATE_LABEL
    End If

End Function

'==============================================================================
'       CSV一致判定ユーティリティ
'------------------------------------------------------------------------------
'   Contents:
'       sourceValue（区切り文字付き）と、
'       expectedValues（カンマ区切り）との一致を判定する。
'
'       ■ この関数の役割
'           ・複数値フィールドの一致判定を共通化
'           ・Aggregate / Filter の両方で使用
'
'       ■ 判定仕様
'           ・両方を分解して全組み合わせ比較
'           ・1つでも一致すれば True
'
'       ■ 正規化
'           ・UCase + Trim により比較
'
'       ■ 注意
'           ・部分一致ではなく「完全一致」
'           ・空文字は無視
'
'   Input:
'       sourceValue        比較元文字列
'       expectedValues     比較対象文字列（カンマ区切り）
'       sourceDelimiter    sourceValue 側区切り文字
'
'   Output:
'       True               一致あり
'       False              一致なし
'
'==============================================================================
Private Function CsvContains(ByVal sourceValue As String, ByVal expectedValues As String, ByVal sourceDelimiter As String) As Boolean

    Dim srcItems As Variant
    Dim expItems As Variant
    Dim src As Variant
    Dim exp As Variant
    Dim srcToken As String
    Dim expToken As String

    srcItems = Split(UCase$(Trim$(CStr(sourceValue))), sourceDelimiter)
    expItems = Split(UCase$(Trim$(CStr(expectedValues))), ",")

    For Each src In srcItems
        srcToken = UCase$(Trim$(CStr(src)))
        If srcToken <> "" Then
            For Each exp In expItems
                expToken = UCase$(Trim$(CStr(exp)))
                If expToken <> "" Then
                    If srcToken = expToken Then
                        CsvContains = True
                        Exit Function
                    End If
                End If
            Next exp
        End If
    Next src

    CsvContains = False

End Function

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
Private Sub OutputAll(ByVal thisAmounts As FinancialAmountMap, ByVal lastAmounts As FinancialAmountMap)

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
    For Each line In m_Master.lines
    
        If line.enabled Then
    
            '金額取得
            vThis = thisAmounts.GetAmount(line.lineID, line.filterValue)
            vLast = lastAmounts.GetAmount(line.lineID, line.filterValue)
            vDiff = vThis - vLast
            
            '------------------------------------------
            ' シート出力
            '------------------------------------------
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Worksheets(line.TargetSheet)
    
            Dim y As Long
            Dim y_inc As Long
            y = CLng(nextRowBySheet(ws.Name))
            
            'ヘッダー出力？
            y_inc = TryOutputSpecialLine(ws, y, line)
            If y_inc <> 0 Then
                nextRowBySheet(ws.Name) = y + y_inc
                GoTo ContinueLineLoop
            End If
            
            '科目名出力
            ws.Cells(y, 2) = line.LineLabel
            ws.Cells(y, 2).HorizontalAlignment = line.LineLabelAligh
    
            '------------------------------------------
            ' 表示制御
            '------------------------------------------
            If Not ShouldSkipOutputLine(line, vThis, vLast, thisAmounts, lastAmounts) Then
                
                If UCase$(Trim$(line.filterValue)) = "EACH" Or UCase$(Trim$(line.statementID)) = "PL_NOTE1" Then
                    'EACHの場合
                    Call OutputEachColumns(ws, y, line, thisAmounts)
                ElseIf line.lineType = "TITLE" Then
                    'タイトル行は金額を表示しない
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
                Dim rowLastCol As Long
                rowLastCol = GetOutputRowLastColumn(ws, y)
                ws.Range(ws.Cells(y, 2), ws.Cells(y, rowLastCol)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                ws.Range(ws.Cells(y, 3), ws.Cells(y, rowLastCol)).NumberFormat = "_ * #,##0_ ;_ * △#,##0_ ;_ * ""-""??_ ;_ @_ "
                ws.Range(ws.Cells(y, 3), ws.Cells(y, rowLastCol)).HorizontalAlignment = xlRight
                
                'ゼブラ（偶数行）
                If y Mod 2 = 0 Then
                    ws.Range(ws.Cells(y, 2), ws.Cells(y, rowLastCol)).Interior.Color = rgbAliceBlue
                End If

                '次行へ
                nextRowBySheet(ws.Name) = y + 1
            End If
        End If
ContinueLineLoop:
    Next

End Sub

'==============================================================================
'       出力行最終列取得
'------------------------------------------------------------------------------
'   Contents:
'       指定行における実際の最終使用列を取得する。
'
'   Input:
'       ws      対象ワークシート
'       y       対象行
'
'   Output:
'       最終使用列番号
'
'==============================================================================
Private Function GetOutputRowLastColumn(ByVal ws As Worksheet, ByVal y As Long) As Long

    Dim lastCol As Long
    lastCol = ws.Cells(y, ws.columns.Count).End(xlToLeft).column
    If lastCol < MaxColumn Then lastCol = MaxColumn
    GetOutputRowLastColumn = lastCol

End Function

'==============================================================================
'       行別区分列出力
'------------------------------------------------------------------------------
'   Contents:
'       EACH指定された行について、
'       区分別（会計区分／財源区分／事業区分）の横展開出力を行う。
'
'       出力モードに応じて、
'       m_ColumnKeys に保持された列定義を参照し、各区分の金額を横方向へ出力する。
'
'       また、帳票種別に応じて、
'           ・内部取引等消去列
'           ・合計列
'           ・配賦基準列
'       を自動付与する。
'
'       【重要】
'       会計区分系帳票では、「公益」「収益」「法人」の順序を固定する。
'       列データが存在しない場合でも、公益法人会計基準の様式維持のため、空列を出力する。
'
'   Input:
'       ws              出力対象シート
'       y               出力行
'       line            財務諸表行定義
'       thisAmounts     金額マップ
'
'==============================================================================
Private Sub OutputEachColumns(ByVal ws As Worksheet, ByVal y As Long, ByVal line As FinancialStatementMasterLine, ByVal thisAmounts As FinancialAmountMap)

    Dim mode As String
    Dim columns As Collection
    Dim col As Variant
    Dim c As Long
    Dim totalValue As Currency
    Dim amountValue As Currency
    Dim categoryKey As String

    mode = ResolveOutputCategoryMode(line)
    Set columns = m_ColumnKeys.GetColumns(mode)

    c = 3
    totalValue = 0

    If mode = CATEGORY_ACCOUNT Or mode = CATEGORY_ACCOUNT_SUM Or mode = CATEGORY_PROJECT Then
        Dim targetGroups As Variant
        Dim grp As Variant
        Dim i As Long
        Dim parts() As String
        Dim hasGroupColumn As Boolean

        '------------------------------------------
        ' 会計区分固定順
        '------------------------------------------
        targetGroups = Array(GROUP_PUBLIC_LABEL, GROUP_REVENUE_LABEL, GROUP_CORPORATE_LABEL)

        For Each grp In targetGroups
            hasGroupColumn = False
            '------------------------------------------
            ' 対象グループ列抽出
            '------------------------------------------
            For i = 1 To columns.Count
                parts = Split(CStr(columns(i)), "|")
                If UBound(parts) >= 0 And CStr(parts(0)) = CStr(grp) Then
                    categoryKey = NormalizeColumnCategoryKey(CStr(columns(i)), mode)
                    amountValue = thisAmounts.GetAmount(line.lineID, categoryKey)
                    ws.Cells(y, c).value = amountValue
                    totalValue = totalValue + amountValue
                    c = c + 1
                    hasGroupColumn = True
                End If
            Next i

            '------------------------------------------
            ' データ未存在時も空列維持
            '------------------------------------------
            If Not hasGroupColumn Then
                ws.Cells(y, c).value = 0
                c = c + 1
            End If
        Next grp
    Else
        '------------------------------------------
        ' 通常列出力
        '------------------------------------------
        For Each col In columns
            categoryKey = NormalizeColumnCategoryKey(CStr(col), mode)
            amountValue = thisAmounts.GetAmount(line.lineID, categoryKey)
            ws.Cells(y, c).value = amountValue
            totalValue = totalValue + amountValue
            c = c + 1
        Next col
    End If

    '------------------------------------------
    ' 内部取引等消去列
    '------------------------------------------
    If UCase$(Trim$(line.statementID)) = "PL_NOTE2" Then
        ws.Cells(y, c).value = 0     '将来対応予定
        c = c + 1
    End If

    '------------------------------------------
    ' 合計列
    '------------------------------------------
    ws.Cells(y, c).value = totalValue

    '------------------------------------------
    ' 配賦基準列
    '------------------------------------------
    If UCase$(Trim$(line.statementID)) = "PL_NOTE3" Then
        ws.Cells(y, c + 1).value = "配賦基準："
    End If

End Sub

'==============================================================================
'       出力区分モード解決
'------------------------------------------------------------------------------
'   Contents:
'       財務諸表行定義から、横軸出力時の区分モードを決定する。
'
'       主な用途：
'           ・ヘッダー生成
'           ・EACH列出力
'           ・数式対象category生成
'
'       PL_NOTE1（財源区分別内訳）のみ、強制的に CATEGORY_NET_ASSET とする。
'
'   Input:
'       line        財務諸表行定義
'
'   Output:
'       区分モード
'
'==============================================================================
Private Function ResolveOutputCategoryMode(ByVal line As FinancialStatementMasterLine) As String

    If UCase$(Trim$(line.statementID)) = "PL_NOTE1" Then
        ResolveOutputCategoryMode = CATEGORY_NET_ASSET
    Else
        ResolveOutputCategoryMode = GetHeaderCategoryModeFromFilterField(line.filterField, CATEGORY_ACCOUNT)
    End If

End Function

'==============================================================================
'       列定義からcategoryキー取得
'------------------------------------------------------------------------------
'   Contents:
'       columnKey の構造から、集計用 category キーを抽出する。
'       columnKey は mode に応じて可変長構造を持つ。
'
'       例：
'
'           CATEGORY_ACCOUNT
'               公益|公1
'
'           CATEGORY_PROJECT
'               公益|公1|PRJ001|演奏会
'
'       本関数は、mode ごとに適切な集計キーを返却する。
'
'   Input:
'       columnKey      列定義キー
'       mode           区分モード
'
'   Output:
'       categoryキー
'
'==============================================================================
Private Function NormalizeColumnCategoryKey(ByVal columnKey As String, ByVal mode As String) As String

    Dim parts() As String
    parts = Split(columnKey, "|")

    Select Case mode
        Case CATEGORY_ACCOUNT
            If UBound(parts) >= 1 Then
                NormalizeColumnCategoryKey = UCase$(Trim$(parts(1)))
            Else
                NormalizeColumnCategoryKey = UCase$(Trim$(columnKey))
            End If
        Case CATEGORY_PROJECT
            If UBound(parts) >= 2 Then
                NormalizeColumnCategoryKey = UCase$(Trim$(parts(2)))
            Else
                NormalizeColumnCategoryKey = UCase$(Trim$(columnKey))
            End If
        Case Else
            NormalizeColumnCategoryKey = UCase$(Trim$(parts(0)))
    End Select

End Function

'==============================================================================
'       出力シート初期化
'------------------------------------------------------------------------------
'   Contents:
'       ・対象シートのクリア
'       ・ヘッダー設定
'       ・シート毎の出力開始行を初期化
'
'   Input:
'       nextRowBySheet     シート別出力行管理Dictionary
'
'==============================================================================
Private Sub InitializeOutputSheets(ByVal nextRowBySheet As Object)

    Dim line As FinancialStatementMasterLine
    Dim ws As Worksheet

    For Each line In m_Master.lines
        If line.enabled Then
            
            'シート単位で一度だけ初期化
            If Not nextRowBySheet.Exists(line.TargetSheet) Then
            
                Set ws = ThisWorkbook.Worksheets(line.TargetSheet)
                
                '出力開始行設定
                nextRowBySheet.Add line.TargetSheet, START_ROW + 1
                
                '既存データクリア
                Call ClearSheet(ws, START_ROW)
                
            End If
        End If
    Next

End Sub

'==============================================================================
'       ヘッダー設定
'------------------------------------------------------------------------------
'   Contents:
'       ・列見出し（科目／当期／前期／増減）
'       ・数値フォーマット
'       ・罫線／背景色
'
'   Input:
'       ws      出力対象シート
'       y       出力行
'       line    財務諸表行
'
'   Output:
'       消費行数
'           0 = 通常行
'           1以上 = 特殊行として処理済
'
'==============================================================================
Private Function TryOutputSpecialLine(ByVal ws As Worksheet, ByVal y As Long, ByVal line As FinancialStatementMasterLine) As Long

    Select Case UCase$(Trim$(line.lineType))
        Case "HEADER_MAIN"
            TryOutputSpecialLine = OutputHeaderMain(ws, y, line)
        Case "HEADER_PL_NOTE1"
            TryOutputSpecialLine = OutputNotePLFundHeader(ws, y, line)
        Case "HEADER_PL_NOTE2"
            TryOutputSpecialLine = OutputNotePLAccountProjectHeader(ws, y, line)
        Case "HEADER_PL_NOTE3"
            TryOutputSpecialLine = OutputNotePLNatureHeader(ws, y, line)
        Case "TEXT"
            ws.Cells(y, 2).value = line.LineLabel
            TryOutputSpecialLine = 1
        Case Else
            TryOutputSpecialLine = 0
    End Select

End Function

'==============================================================================
'       貸借対照表 ＆ 活動計算書    共通ヘッダー
'------------------------------------------------------------------------------
Private Function OutputHeaderMain(ByVal ws As Worksheet, ByVal y As Long, ByVal line As FinancialStatementMasterLine) As Long

    With ws

        '========================
        ' 1行目：タイトル
        '========================
        .Cells(y, 2).value = thisYear & "年度 " & line.LineLabel
        .Range(.Cells(y, 2), .Cells(y, MaxColumn)).Merge
    
        With .Cells(y, 2).Font
            .Name = "Meiryo UI"
            .Size = 20
            .Bold = True                 '太字
            .Underline = xlUnderlineStyleSingle   '下線
        End With
    
        .Cells(y, 2).HorizontalAlignment = xlCenter
    
        '========================
        ' 2行目：期間
        '========================
        If UCase$(Trim$(line.statementID)) = "BS" Then
            .Cells(y + 1, 2).value = str_end_date
        Else
            .Cells(y + 1, 2).value = str_this_period
        End If
    
        .Range(.Cells(y + 1, 2), .Cells(y + 1, MaxColumn)).Merge
    
        With .Cells(y + 1, 2).Font
            .Name = "Meiryo UI"
            .Size = 11
        End With
    
        .Cells(y + 1, 2).HorizontalAlignment = xlCenter
    
        '========================
        ' 3行目：単位
        '========================
        .Cells(y + 2, MaxColumn).value = "（単位：円）"
        .Cells(y + 2, MaxColumn).HorizontalAlignment = xlRight
    
        '========================
        ' 4行目：ヘッダ
        '========================
        Call SetupSheetHeader(ws, line.statementID, y + 3, MaxColumn)

    End With

    OutputHeaderMain = 4

End Function

'==============================================================================
'       シートヘッダー設定
'------------------------------------------------------------------------------
'   Contents:
'       ・列見出し（科目／当期／前期／増減）
'       ・数値フォーマット
'       ・罫線／背景色
'
'==============================================================================
Private Sub SetupSheetHeader(ByVal ws As Worksheet, ByVal statementID As String, Optional ByVal headerRow As Long = START_ROW, Optional ByVal maxCol As Long = MaxColumn)

    With ws

        'ヘッダー項目
        .Cells(headerRow, 2) = "科目"
        .Cells(headerRow, 3) = "当期"
        .Cells(headerRow, 4) = "前期"
        .Cells(headerRow, 5) = "増減"
    
        'ヘッダー装飾
        With .Range(.Cells(headerRow, 2), .Cells(headerRow, maxCol))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = rgbLavender
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With

    End With

End Sub

'==============================================================================
'       活動計算書の注記    （1）活動計算書財源区分別内訳
'------------------------------------------------------------------------------
Private Function OutputNotePLFundHeader(ByVal ws As Worksheet, ByVal y As Long, ByVal line As FinancialStatementMasterLine) As Long

    Const x_size As Long = 5
    
    Dim y_inc As Long
    Dim headerRow As Long

    y_inc = 0
    
    With ws
    
        '========================
        ' １行目：  注記タイトル
        '========================
        .Cells(y + y_inc, 2).value = line.LineLabel
        With .Cells(y + y_inc, 2).Font
            .Name = "Meiryo UI"
            .Bold = True
        End With
        
        '========================
        ' ２行目：  単位
        '========================
        .Cells(y + y_inc, x_size).value = "（単位：円）"
        .Cells(y + y_inc, x_size).HorizontalAlignment = xlRight
        y_inc = y_inc + 1
        
        '========================
        ' ３行目：  ヘッダ（財源区分）
        '========================
        headerRow = y + y_inc
        .Cells(y + y_inc, 2).value = "科目"
        .Cells(y + y_inc, 3).value = "一般"
        .Cells(y + y_inc, 4).value = "指定"
        .Cells(y + y_inc, 5).value = "合計"

        'ヘッダー装飾
        Call ApplyHeaderStyle(ws, headerRow, headerRow, 2, x_size)
        
        y_inc = y_inc + 1
    
    End With
    
    OutputNotePLFundHeader = y_inc

End Function

'==============================================================================
'       活動計算書の注記    （2）活動計算書会計区分及び事業区分別内訳
'------------------------------------------------------------------------------
Private Function OutputNotePLAccountProjectHeader(ByVal ws As Worksheet, ByVal y As Long, ByVal line As FinancialStatementMasterLine) As Long

    Dim y_inc As Long
    Dim headerRow As Long
    Dim headerDepth As Long
    Dim mode As String
    Dim columns As Collection
    Dim c As Long
    Dim i As Long
    Dim parts() As String
    
    Dim grp As String
    Dim targetGroups As Variant
    
    Dim it As Variant
    
    '表示項目
    targetGroups = Array(GROUP_PUBLIC_LABEL, GROUP_REVENUE_LABEL, GROUP_CORPORATE_LABEL)

    mode = GetHeaderCategoryModeFromFilterField(line.filterField, CATEGORY_ACCOUNT)

    Set columns = m_ColumnKeys.GetColumns(mode)

    y_inc = 0

    With ws

        '========================
        ' １行目：  注記タイトル
        '========================
        .Cells(y + y_inc, 2).value = line.LineLabel
        With .Cells(y + y_inc, 2).Font
            .Name = "Meiryo UI"
            .Bold = True
        End With
        y_inc = y_inc + 1
        
        '========================
        ' ２行目：  単位
        '========================
        .Cells(y + y_inc, 3 + columns.Count + 1).value = "（単位：円）"
        .Cells(y + y_inc, 3 + columns.Count + 1).HorizontalAlignment = xlRight
        y_inc = y_inc + 1
        
        '========================
        ' ３行目：  会計
        '========================
        headerRow = y + y_inc
        headerDepth = IIf(mode = CATEGORY_PROJECT, 3, 2)

        .Cells(headerRow, 2).value = "科目"
        If headerDepth > 1 Then .Range(.Cells(headerRow, 2), .Cells(headerRow + headerDepth - 1, 2)).Merge

        c = 3
        Dim groupCols As Collection
        Dim col As Variant
        Dim hasData As Boolean
        Dim startCol As Long

        For Each it In targetGroups
            grp = CStr(it)
            
            Set groupCols = New Collection
            
            '------------------------------------------
            ' このグループの列抽出
            '------------------------------------------
            For i = 1 To columns.Count
                parts = Split(CStr(columns(i)), "|")
                If parts(0) = grp Then
                    groupCols.Add columns(i)
                End If
            Next i
            
            startCol = c
            
            '------------------------------------------
            ' データあり
            '------------------------------------------
            If groupCols.Count > 0 Then
                
                For Each col In groupCols
                    parts = Split(CStr(col), "|")
                    
                    'グループ
                    .Cells(headerRow, c).value = parts(0)
                    
                    '区分
                    .Cells(headerRow + 1, c).value = parts(1)
                    
                    '★補助科目名表示
                    If headerDepth = 3 Then
                        If UBound(parts) >= 3 Then
                            .Cells(headerRow + 2, c).value = parts(3)
                        Else
                            .Cells(headerRow + 2, c).value = parts(2)
                        End If
                    End If
                    
                    c = c + 1
                Next
                
           
            '------------------------------------------
            ' ★データなし（ここが今回の本丸）
            '------------------------------------------
            Else
                
                'グループだけ1列出す
                .Cells(headerRow, c).value = grp
            
                '★ここが重要：縦方向マージ
                If headerDepth > 1 Then
                    .Range(.Cells(headerRow, c), _
                           .Cells(headerRow + headerDepth - 1, c)).Merge
                End If
                
                c = c + 1
            
            End If
            
        Next it

        '------------------------------------------
        ' 末尾列
        '------------------------------------------
        .Cells(headerRow, c).value = "内部取引等消去"
        If headerDepth > 1 Then
            .Range(.Cells(headerRow, c), .Cells(headerRow + headerDepth - 1, c)).Merge
        End If
        
        .Cells(headerRow, c + 1).value = "合計"
        If headerDepth > 1 Then
            .Range(.Cells(headerRow, c + 1), .Cells(headerRow + headerDepth - 1, c + 1)).Merge
        End If
        
        '==========================================
        ' ★ここ修正：横マージは最後にまとめてやる
        '==========================================
        Application.DisplayAlerts = False
        
        '① グループ行（1段目）
        Call MergeSameCategoryCells(ws, headerRow, 3, c - 1)
        
        '② 会計区分行（2段目）
        If headerDepth >= 2 Then
            Call MergeSameCategoryCells(ws, headerRow + 1, 3, c - 1)
        End If
        
        Application.DisplayAlerts = True
        
        Call ApplyHeaderStyle(ws, headerRow, headerRow + headerDepth - 1, 2, c + 1)
        
        y_inc = y_inc + headerDepth

    End With

    OutputNotePLAccountProjectHeader = y_inc

End Function


'==============================================================================
'       活動計算書の注記    （3）事業費及び管理費の形態別区分
'------------------------------------------------------------------------------
Private Function OutputNotePLNatureHeader(ByVal ws As Worksheet, ByVal y As Long, ByVal line As FinancialStatementMasterLine) As Long
    
    Dim y_inc As Long
    Dim headerRow As Long
    Dim headerDepth As Long
    Dim mode As String
    Dim columns As Collection
    Dim c As Long
    Dim i As Long
    Dim parts() As String
    
    Dim grp As Variant
    Dim targetGroups As Variant
    Dim groupCols As Collection
    Dim col As Variant
    Dim startCol As Long

    '★固定グループ
    targetGroups = Array(GROUP_PUBLIC_LABEL, GROUP_REVENUE_LABEL, GROUP_CORPORATE_LABEL)

    mode = GetHeaderCategoryModeFromFilterField(line.filterField, CATEGORY_ACCOUNT_SUM)

    Set columns = m_ColumnKeys.GetColumns(mode)

    y_inc = 0

    With ws

        '========================
        ' タイトル
        '========================
        .Cells(y + y_inc, 2).value = line.LineLabel
        .Cells(y + y_inc, 2).Font.Bold = True
        y_inc = y_inc + 1
        
        '========================
        ' 単位
        '========================
        .Cells(y + y_inc, 10).value = "（単位：円）"
        .Cells(y + y_inc, 10).HorizontalAlignment = xlRight
        y_inc = y_inc + 1
        
        '========================
        ' ヘッダー
        '========================
        headerRow = y + y_inc
        
        Select Case mode
            Case CATEGORY_PROJECT: headerDepth = 3
            Case CATEGORY_ACCOUNT: headerDepth = 2
            Case Else: headerDepth = 1
        End Select

        .Cells(headerRow, 2).value = "科目"
        If headerDepth > 1 Then
            .Range(.Cells(headerRow, 2), .Cells(headerRow + headerDepth - 1, 2)).Merge
        End If

        c = 3

        '------------------------------------------
        ' ★グループ単位で処理
        '------------------------------------------
        For Each grp In targetGroups
            
            Set groupCols = New Collection
            
            'このグループの列抽出
            For i = 1 To columns.Count
                parts = Split(CStr(columns(i)), "|")
                If parts(0) = grp Then
                    groupCols.Add columns(i)
                End If
            Next i
            
            startCol = c
            
            '------------------------------------------
            ' データあり
            '------------------------------------------
            If groupCols.Count > 0 Then
                
                For Each col In groupCols
                    
                    parts = Split(CStr(col), "|")
                    
                    'グループ
                    .Cells(headerRow, c).value = parts(0)

                    '区分
                    If headerDepth >= 2 Then
                        .Cells(headerRow + 1, c).value = IIf(UBound(parts) >= 1, parts(1), "")
                    End If
                    
                    '★補助科目名
                    If headerDepth = 3 Then
                        If UBound(parts) >= 3 Then
                            .Cells(headerRow + 2, c).value = parts(3)
                        ElseIf UBound(parts) >= 2 Then
                            .Cells(headerRow + 2, c).value = parts(2)
                        End If
                    End If
                    
                    c = c + 1
                Next
                
            
            '------------------------------------------
            ' ★データなし（←ここ重要）
            '------------------------------------------
            Else
                
                .Cells(headerRow, c).value = grp
                
                '★縦マージ
                If headerDepth > 1 Then
                    With .Range(.Cells(headerRow, c), _
                                .Cells(headerRow + headerDepth - 1, c))
                        .Merge
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                    End With
                End If
                
                c = c + 1
                
            End If
            
        Next grp

        '------------------------------------------
        ' 末尾列
        '------------------------------------------
        .Cells(headerRow, c).value = "合計"
        .Cells(headerRow, c + 1).value = "配賦基準"
        
        If headerDepth > 1 Then
            .Range(.Cells(headerRow, c), .Cells(headerRow + headerDepth - 1, c)).Merge
            .Range(.Cells(headerRow, c + 1), .Cells(headerRow + headerDepth - 1, c + 1)).Merge
        End If
        
        '==========================================
        ' ★ここ修正：横マージは最後にまとめてやる
        '==========================================
        Application.DisplayAlerts = False
        
        '① グループ行
        Call MergeSameCategoryCells(ws, headerRow, 3, c - 1)
        
        '② 会計区分行
        If headerDepth >= 2 Then
            Call MergeSameCategoryCells(ws, headerRow + 1, 3, c - 1)
        End If
        
        Application.DisplayAlerts = True
        
        Call ApplyHeaderStyle(ws, headerRow, headerRow + headerDepth - 1, 2, c + 1)
        
        y_inc = y_inc + headerDepth
    
    End With

    OutputNotePLNatureHeader = y_inc

End Function

'==============================================================================
'       同一カテゴリセルの横方向マージ
'------------------------------------------------------------------------------
'   Contents:
'       指定行のセルを左から走査し、同一値が連続する範囲を横方向にマージする
'       ただし、空白セルはマージ対象としない
'
'   Input:
'       ws          対象ワークシート
'       row         対象行
'       colStart    開始列
'       colEnd      終了列
'
'   Output:
'       なし（対象範囲のセルを直接マージする）
'
'==============================================================================
Private Sub MergeSameCategoryCells(ByVal ws As Worksheet, ByVal row As Long, ByVal colStart As Long, ByVal colEnd As Long)

    Dim c As Long
    Dim startCol As Long
    Dim currentVal As String
    Dim nextVal As String
    
    startCol = colStart
    currentVal = CStr(ws.Cells(row, colStart).value)
    
    For c = colStart + 1 To colEnd + 1
        
        If c <= colEnd Then
            
            nextVal = CStr(ws.Cells(row, c).value)
            
            '★ここ修正（空白はマージしない）
            If nextVal = currentVal And currentVal <> "" Then
                '継続
            Else
                If c - startCol > 1 Then
                    ws.Range(ws.Cells(row, startCol), ws.Cells(row, c - 1)).Merge
                End If
                
                startCol = c
                currentVal = nextVal
            End If
            
        Else
            If c - startCol > 1 And currentVal <> "" Then
                ws.Range(ws.Cells(row, startCol), ws.Cells(row, c - 1)).Merge
            End If
        End If
        
    Next c

End Sub

'==============================================================================
'       ヘッダー区分モード取得（フィルタ項目ベース）
'------------------------------------------------------------------------------
'   Contents:
'       フィルタ項目（filterField）の値に応じて、
'       ヘッダーの区分モードを決定する
'
'       指定されたフィルタ項目が以下のいずれかに一致する場合、
'       対応する区分モードをそのまま返す
'
'           CATEGORY_NET_ASSET
'           CATEGORY_PROJECT
'           CATEGORY_ACCOUNT
'           CATEGORY_ACCOUNT_SUM
'
'       上記以外の場合は、デフォルトモード（defaultMode）を返す
'
'   Input:
'       filterField     フィルタ項目
'       defaultMode     デフォルト区分モード
'
'   Output:
'       ヘッダー区分モード
'
'==============================================================================
Private Function GetHeaderCategoryModeFromFilterField(ByVal filterField As String, ByVal defaultMode As String) As String

    Select Case Trim$(filterField)
        Case CATEGORY_NET_ASSET:    GetHeaderCategoryModeFromFilterField = CATEGORY_NET_ASSET
        Case CATEGORY_PROJECT:      GetHeaderCategoryModeFromFilterField = CATEGORY_PROJECT
        Case CATEGORY_ACCOUNT:      GetHeaderCategoryModeFromFilterField = CATEGORY_ACCOUNT
        Case CATEGORY_ACCOUNT_SUM:  GetHeaderCategoryModeFromFilterField = CATEGORY_ACCOUNT_SUM
        Case Else: GetHeaderCategoryModeFromFilterField = defaultMode
    End Select

End Function

'==============================================================================
'       ヘッダースタイル適用
'------------------------------------------------------------------------------
'   Contents:
'       指定範囲に対して、ヘッダー用の書式（中央揃え・背景色・罫線等）を適用する
'
'   Input:
'       ws          対象ワークシート
'       rowFrom     開始行
'       rowTo       終了行
'       colFrom     開始列
'       colTo       終了列
'
'==============================================================================
Private Sub ApplyHeaderStyle(ByVal ws As Worksheet, ByVal rowFrom As Long, ByVal rowTo As Long, ByVal colFrom As Long, ByVal colTo As Long)

    With ws.Range(ws.Cells(rowFrom, colFrom), ws.Cells(rowTo, colTo))
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
Private Function ShouldSkipOutputLine(ByVal line As FinancialStatementMasterLine, ByVal vThis As Currency, ByVal vLast As Currency, ByVal thisAmounts As FinancialAmountMap, ByVal lastAmounts As FinancialAmountMap) As Boolean

    Dim isEach As Boolean
    isEach = (UCase$(Trim$(line.filterValue)) = "EACH" Or UCase$(Trim$(line.statementID)) = "PL_NOTE1")

    If line.lineType = "TITLE" _
        Or line.lineType = "ACCOUNTS" _
        Or line.lineType = "HEADER_MAIN" _
        Or line.lineType = "TEXT" Then
        
        ShouldSkipOutputLine = False
        Exit Function
    End If

    'ゼロ行非表示
    If line.lineType = "ACCOUNT" Or line.lineType = "ACCOUNTS" Then
        If isEach Then
            ShouldSkipOutputLine = ShouldSkipEachAccountLine(line, thisAmounts, lastAmounts)
            If ShouldSkipOutputLine Then Exit Function
        ElseIf vThis = 0 And vLast = 0 Then
            ShouldSkipOutputLine = True
            Exit Function
        End If
    End If
       
    ShouldSkipOutputLine = False  'デフォルトは表示
    
End Function

'==============================================================================
'       EACH行ゼロ判定
'------------------------------------------------------------------------------
'   Contents:
'       EACH列展開対象行について、全カテゴリが0かどうかを判定する。
'       いずれか1カテゴリでも値が存在する場合は表示対象とする。
'
'   Input:
'       line            財務諸表行
'       thisAmounts     当期金額Map
'       lastAmounts     前期金額Map
'
'   Output:
'       True    表示不要（全カテゴリ0）
'       False   表示対象
'
'==============================================================================
Private Function ShouldSkipEachAccountLine(ByVal line As FinancialStatementMasterLine, ByVal thisAmounts As FinancialAmountMap, ByVal lastAmounts As FinancialAmountMap) As Boolean

    Dim mode As String
    Dim columns As Collection
    Dim col As Variant
    Dim categoryKey As String
    Dim thisValue As Currency
    Dim lastValue As Currency

    mode = ResolveOutputCategoryMode(line)
    Set columns = m_ColumnKeys.GetColumns(mode)

    For Each col In columns
        categoryKey = NormalizeColumnCategoryKey(CStr(col), mode)
        thisValue = thisAmounts.GetAmount(line.lineID, categoryKey)
        lastValue = lastAmounts.GetAmount(line.lineID, categoryKey)
        If thisValue <> 0 Or lastValue <> 0 Then
            ShouldSkipEachAccountLine = False
            Exit Function
        End If
    Next col

    ShouldSkipEachAccountLine = True

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
Private Sub ResolveFormulaAmounts(ByVal values As FinancialAmountMap)

    Dim line As FinancialStatementMasterLine
    Dim category As Variant
    Dim categories As Collection
    
    For Each line In m_Master.lines
        If line.enabled Then
            If line.lineType = "FORMULA" Or line.lineType = "CALC" Then
                Set categories = GetFormulaTargetCategories(line, values)
                For Each category In categories
                    values.SetAmount line.lineID, CStr(category), ResolveFormulaByCategory(line.formula, values, CStr(category))
                Next category
            End If
        End If
    Next

End Sub

'==============================================================================
'       数式対象カテゴリ取得
'------------------------------------------------------------------------------
'   Contents:
'       FORMULA / CALC 行について、数式評価対象となるカテゴリ一覧を取得する。
'
'       【戻り値保証】
'       最低1件 "*" を返却する。
'
'   Input:
'       line        財務諸表行
'       values      金額Map
'
'   Output:
'       評価対象カテゴリCollection
'
'==============================================================================
Private Function GetFormulaTargetCategories(ByVal line As FinancialStatementMasterLine, ByVal values As FinancialAmountMap) As Collection

    Dim categories As Collection
    Dim mode As String
    Dim columns As Collection
    Dim col As Variant
    Dim categoryKey As String
    Dim seen As Object

    If UCase$(Trim$(line.filterValue)) <> "EACH" And UCase$(Trim$(line.statementID)) <> "PL_NOTE1" Then
        Set categories = values.GetCategories(line.lineID)
        If categories.Count = 0 Then categories.Add "*"
        Set GetFormulaTargetCategories = categories
        Exit Function
    End If

    Set categories = New Collection
    Set seen = CreateObject("Scripting.Dictionary")

    mode = ResolveOutputCategoryMode(line)
    Set columns = m_ColumnKeys.GetColumns(mode)

    For Each col In columns
        categoryKey = NormalizeColumnCategoryKey(CStr(col), mode)
        If categoryKey <> "" Then
            categoryKey = UCase$(Trim$(categoryKey))
            If Not seen.Exists(categoryKey) Then
                seen.Add categoryKey, True
                categories.Add categoryKey
            End If
        End If
    Next col

    If categories.Count = 0 Then categories.Add "*"
    Set GetFormulaTargetCategories = categories

End Function

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
'   Input:
'       amounts             財務諸表金額マップ
'       isFormulaAfter      数式解決後かどうか
'
'==============================================================================
Private Sub RollupAggregateTo(ByVal amountMap As FinancialAmountMap, ByVal includeFormula As Boolean)

    Dim line As FinancialStatementMasterLine
    Dim childType As String

    For Each line In m_Master.lines
        If line.enabled Then
            childType = UCase$(Trim$(line.lineType))

            'タイトル行は対象外、かつ親が定義されている場合のみ
            If childType <> "TITLE" And Trim$(line.aggregateTo) <> "" Then
                
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
'       親行へのロールアップ処理
'------------------------------------------------------------------------------
'   Contents:
'       子行（child）の金額を親行へ加算する。
'
'       処理内容：
'           ・child の金額を取得
'           ・parent のキーを生成
'           ・親へ加算
'
'       【前提】
'       ・child → parent の関係はマスタで定義済
'       ・amountMap には child の値が既に存在している
'
'       【重要】
'       ・この処理により「ツリー構造の集計」が実現される
'       ・複数階層でも再帰的に適用可能な設計
'
'       【注意】
'       ・Key構造は BuildKey に依存
'       ・accountingCategory 単位で独立して集計される
'
'   Input:
'       child       子行（FinancialStatementMasterLine）
'       amountMap   金額マップ（Dictionary）
'
'==============================================================================
Private Sub RollupLineAmount(ByVal child As FinancialStatementMasterLine, ByVal amountMap As FinancialAmountMap)

    Dim parentLine As FinancialStatementMasterLine
    Dim childAmount As Currency
    Dim category As Variant

    For Each category In amountMap.GetCategories(child.lineID)
        '------------------------------------------
        ' 子の金額取得
        '------------------------------------------
        childAmount = amountMap.GetAmount(child.lineID, CStr(category))

        '親未定義なら次
        If Trim$(child.aggregateTo) <> "" Then
        
            '------------------------------------------
            ' 親行取得
            Set parentLine = GetAggregateParentLine(child.aggregateTo, CStr(category))
            
            '------------------------------------------
            ' 親へ加算
            amountMap.AddAmount parentLine.lineID, CStr(category), childAmount
        End If
    Next category
    
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
'       数式評価処理
'------------------------------------------------------------------------------
Private Function ResolveFormulaByCategory(ByVal formula As String, ByVal values As FinancialAmountMap, ByVal defaultCategory As String) As Currency

    Dim expr As String
    expr = formula

    If Trim$(defaultCategory) = "" Then defaultCategory = "*"

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

        If InStr(token, "|") > 0 Then
            lineID = UCase$(Trim$(Split(token, "|")(0)))
            category = UCase$(Trim$(Split(token, "|")(1)))
        Else
            lineID = UCase$(Trim$(token))
            category = UCase$(Trim$(defaultCategory))
        End If

        If category = "" Then category = "*"

        Dim val As Currency
        val = values.GetAmount(lineID, category)

        expr = Replace(expr, "[" & token & "]", CStr(val))
    Next

    expr = Replace(expr, " ", "")
    ResolveFormulaByCategory = CCur(Application.Evaluate(expr))

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

    For Each line In m_Master.lines
        If UCase$(Trim$(line.lineID)) = targetID Then
            ExistsLineID = True
            Exit Function
        End If
    Next

    ExistsLineID = False

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
Private Function ResolveFormula(ByVal formula As String, ByVal values As FinancialAmountMap) As Currency

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
        
        Dim val As Currency
        
        '------------------------------------------
        ' ① 完全一致
        '------------------------------------------
        If values.Exists(lineID, category) Then
            val = values.GetAmount(lineID, category)
        
        '------------------------------------------
        ' ② "*"（合算）
        '------------------------------------------
        ElseIf category = "*" Then
            
            Dim k As Variant
            val = 0
            
            For Each k In values.Keys
                If Left$(UCase$(CStr(k)), Len(UCase$(lineID)) + 1) = UCase$(lineID) & "|" Then
                    val = val + CCur(values.GetAmount(lineID, Mid$(UCase$(CStr(k)), Len(UCase$(lineID)) + 2)))
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


