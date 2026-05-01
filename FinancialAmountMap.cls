VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FinancialAmountMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************
'*      メンバー変数                    *
'****************************************
Private m_acc_Master As AccountMaster                            '勘定科目マスタ
Private m_fs_master As FinancialStatementMaster                    '財務諸表マスタ管理クラス
Private m_Map As Object

'==============================================================================
'       初期化処理
'------------------------------------------------------------------------------
'   Contents:
'       FinancialAmountMap クラス生成時の初期化処理。
'
'==============================================================================
Private Sub Class_Initialize()
    Set m_Map = CreateObject("Scripting.Dictionary")
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
Public Sub Set_Master(fs_master As FinancialStatementMaster, acc_master As AccountMaster)
    Set m_fs_master = fs_master
    Set m_acc_Master = acc_master
End Sub

'==============================================================================
'       財務諸表金額マップ用キー生成
'------------------------------------------------------------------------------
'   Contents:
'       財務諸表金額マップ（Dictionary）で使用する
'       一意キーを生成する。
'
'       ■ キー形式
'           lineID | category
'
'       ■ 正規化仕様
'           ・前後空白除去（Trim）
'           ・大文字統一（UCase）
'           ・category が空の場合は "*" に補正
'
'       ■ 用途
'           財務諸表行 + 会計区分 等の組み合わせで
'           金額を管理するために使用する。
'
'       ■ 注意
'           キー仕様変更時は、
'           集計系処理全体へ影響するため注意すること。
'
'   Input:
'       lineID         財務諸表行ID
'       category       会計区分・事業区分等
'
'   Output:
'       BuildKey       正規化済みキー文字列
'
'==============================================================================
Private Function BuildKey(ByVal lineID As String, ByVal category As String) As String
    Dim normalizedLineID As String
    Dim normalizedCategory As String

    normalizedLineID = UCase$(Trim$(lineID))
    normalizedCategory = UCase$(Trim$(category))
    If normalizedCategory = "" Then normalizedCategory = "*"

    BuildKey = normalizedLineID & "|" & normalizedCategory
End Function

'==============================================================================
'       存在チェック
'------------------------------------------------------------------------------
Public Function Exists(ByVal lineID As String, ByVal category As String) As Boolean
    Exists = m_Map.Exists(BuildKey(lineID, category))
End Function

'==============================================================================
'       key 取得
'==============================================================================
Public Function Keys() As Variant
    Keys = m_Map.Keys
End Function

'==============================================================================
'       行ID配下カテゴリ取得
'------------------------------------------------------------------------------
'   Contents:
'       指定lineIDに紐づくカテゴリ一覧を、金額Mapから抽出する。
'
'       金額MapのKey形式：  lineID|category
'
'       処理内容：          ・指定lineID配下を走査
'                           ・category部分を抽出
'                           ・重複除去
'
'       戻り値には、"*"（全体カテゴリ）を最低1件含める。
'
'       本処理は、以下の処理等で使用される。
'           ・Rollup集計
'           ・数式カテゴリ解決
'           ・親子集計
'
'       【重要】
'       amountMap に実際に存在するカテゴリを基準とするため、
'       マスタ定義とは独立して動的にカテゴリを取得可能。
'
'   Input:
'       lineID      対象財務諸表行ID
'
'   Output:
'       category一覧Collection
'
'==============================================================================
Private Function GetCategories(ByVal lineID As String) As Collection
    Dim result As New Collection
    Dim seen As Object
    Dim key As Variant
    Dim normalizedLineID As String
    Dim suffix As String
    Dim category As String

    Set seen = CreateObject("Scripting.Dictionary")
    normalizedLineID = UCase$(Trim$(lineID))

    For Each key In m_Map.Keys
        suffix = UCase$(CStr(key))
        If Left$(suffix, Len(normalizedLineID) + 1) = normalizedLineID & "|" Then
            category = Mid$(suffix, Len(normalizedLineID) + 2)
            If category = "" Then category = "*"
            If Not seen.Exists(category) Then
                seen.Add category, True
                result.Add category
            End If
        End If
    Next

    If Not seen.Exists("*") Then result.Add "*"

    Set GetCategories = result
End Function

'==============================================================================
'       行ID＋区分別金額取得
'------------------------------------------------------------------------------
'   Contents:
'       lineID と category の組み合わせから、amountMap に格納された金額を取得する。
'
'   Input:
'       lineID          財務諸表ID
'       category        区分キー
'
'   Output:
'       金額            キー未存在時は0を返却する。
'
'==============================================================================
Public Function GetAmount(ByVal lineID As String, ByVal category As String) As Currency
    Dim key As String
    key = BuildKey(lineID, category)

    If m_Map.Exists(key) Then
        GetAmount = CCur(m_Map(key))
    Else
        GetAmount = 0
    End If
End Function

'==============================================================================
'       金額マップ代入処理（最終集計単位）
'------------------------------------------------------------------------------
'   Contents:
'       Dictionary（lineID|category）に金額を代入する。
'
'       ■ この関数の役割
'           ・全ての集計結果がここに集約される
'           ・最終的な財務諸表の数値を構成する基盤
'
'       ■ 動作
'           ・キーがなければ新規作成（0初期化）
'
'       ■ キー構造
'           lineID | accountingCategory
'
'       ■ 注意（重要）
'           ・キー生成ロジックはここに固定されている
'           ・ここが壊れると全体集計が壊れる
'
'   Input:
'       lineID                 財務諸表行ID
'       accountingCategory     分類キー
'       amount                 加算金額
'
'==============================================================================
Public Sub SetAmount(ByVal lineID As String, ByVal category As String, ByVal value As Currency)
    Dim key As String
    key = BuildKey(lineID, category)

    If m_Map.Exists(key) Then
        m_Map(key) = value
    Else
        m_Map.Add key, value
    End If
End Sub

'==============================================================================
'       金額マップ加算処理（最終集計単位）
'------------------------------------------------------------------------------
'   Contents:
'       Dictionary（lineID|category）に金額を加算する。
'
'   Input:
'       lineID                 財務諸表行ID
'       accountingCategory     分類キー
'       amount                 加算金額
'
'==============================================================================
Public Sub AddAmount(ByVal lineID As String, ByVal category As String, ByVal value As Currency)
    Call SetAmount(lineID, category, GetAmount(lineID, category) + value)
End Sub
   
'==============================================================================
'       Returns True if any amount exists for the line across target categories.
'------------------------------------------------------------------------------
'   Contents:
'       EACH対象に金額があるかチェック
'
'   Input:
'       line                    財務諸表行
'       columnKeys              分類キー
'       lastAmounts             前期（オプション）
'                               指定した場合は、前期を含めてチェックする
'
'==============================================================================
Public Function HasAnyAmount(ByVal line As FinancialStatementMasterLine, ByVal columnKeys As FinancialStatementColumnKeys, Optional ByVal lastAmounts As FinancialAmountMap) As Boolean

    Dim mode As String
    Dim columns As Collection
    Dim col As Variant
    Dim categoryKey As String
    Dim thisValue As Currency
    Dim lastValue As Currency

    mode = line.OutputCategoryMode
    Set columns = columnKeys.GetColumns(mode)

    For Each col In columns
        categoryKey = columnKeys.ResolveCategoryKey(CStr(col), mode)
        thisValue = Me.GetAmount(line.lineID, categoryKey)
        If Not lastAmounts Is Nothing Then
            lastValue = lastAmounts.GetAmount(line.lineID, categoryKey)
        Else
            lastValue = 0
        End If
        If thisValue <> 0 Or lastValue <> 0 Then
            HasAnyAmount = True
            Exit Function
        End If
    Next col

    HasAnyAmount = False

End Function

'******************************************************************************
'*                                                                            *
'*                          マップ作成                                        *
'*                                                                            *
'******************************************************************************
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
'==============================================================================
Public Sub BuildStatementAmountMap(ByVal tb As TrialBalance)

    Dim tbLine As TrialBalanceLine

    For Each tbLine In tb.lines

        Call ApplyMasterLineMatching(tbLine)
    
    Next tbLine

End Sub

'==============================================================================
'       純資産データ追加
'------------------------------------------------------------------------------
'   Contents:
'       試算表には存在しない純資産関連項目について、財務諸表表示用の金額を追加する。
'
'       主に、
'           ・一般正味財産
'           ・指定正味財産
'       の期首・期末残高等を補完する目的で使用する。
'
'       ■ 設計意図
'           純資産系は通常の収益費用集計とは性質が異なるため、
'           BuildStatementAmountMap とは分離している。
'
'       ■ 注意
'           本処理は財務諸表表示用データ補完であり、
'           TrialBalance 自体を書き換えるものではない。
'
'   Input:
'       cur_general_start       期初残高
'       cur_general_end         期末残高
'       cur_desig_start         期初残高
'       cur_desig_end           期末残高
'
'==============================================================================
Public Sub AddNetAssetAmounts( _
    ByVal cur_general_start As Currency, _
    ByVal cur_general_end As Currency, _
    ByVal cur_desig_start As Currency, _
    ByVal cur_desig_end As Currency)

    '==============================
    ' 一般純資産
    Call AddAmountToMatchedMasterLines(idNetAssets_Begin, natGeneral, cur_general_start)
    Call AddAmountToMatchedMasterLines(idNetAssets_End, natGeneral, cur_general_end)
    
    '==============================
    ' 指定純資産
    Call AddAmountToMatchedMasterLines(idNetAssets_Begin, natDesignated, cur_desig_start)
    Call AddAmountToMatchedMasterLines(idNetAssets_End, natDesignated, cur_desig_end)

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
'   Input:
'       code           勘定科目コード
'       nat            純資産区分（一般／指定）
'       amount         残高
'
'==============================================================================
Private Sub AddAmountToMatchedMasterLines(ByVal code As Long, ByVal nat As NetAssetType, ByVal amount As Currency)

    Dim tbLine As TrialBalanceLine
    Set tbLine = New TrialBalanceLine
    
    Call tbLine.Initialize(m_acc_Master.GetAccount(code), 0, "", "", nat)
    Call tbLine.AddAmount(0, 0, amount, amount)

    Call ApplyMasterLineMatching(tbLine)

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
Private Sub ApplyMasterLineMatching(ByVal tbLine As TrialBalanceLine)

    Dim masterLine As FinancialStatementMasterLine
    Dim detailKey As String
    Dim filterTokens As Variant
    Dim filterToken As Variant
    Dim amount As Currency
    
    amount = tbLine.amount

    '------------------------------------------
    ' 試算表1行 × マスタ全行 を突き合わせる
    '------------------------------------------
    For Each masterLine In m_fs_master.lines
        If masterLine.enabled Then
            If masterLine.lineType = "ACCOUNT" Or masterLine.lineType = "ACCOUNTS" Then

                'ACCOUNT / ACCOUNTS：
                '実データ（試算表）から直接金額を拾う行のみ対象
                If IsAggregateMatch(masterLine, tbLine) And IsFilterMatch(masterLine, tbLine) Then
                    '------------------------------------------
                    ' マスタ定義に基づき、この試算表行が
                    ' この財務諸表行に属するか判定
                    '------------------------------------------
                    
                    Call AddAmount(masterLine.lineID, "*", amount)

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
                            Call AddAmount(masterLine.lineID, detailKey, amount)
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
                                Call AddAmount(masterLine.lineID, detailKey, amount)
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

        Case Terget_ID          '財務諸表IDベース   （デフォルト）
            IsAggregateMatch = CsvContains(tbLine.account.fsLineIDs, masterLine.aggregateValue, ";")

        Case Terget_Sub_Class   '中区分ベース
            IsAggregateMatch = CsvContains(tbLine.account.SubCategory, masterLine.aggregateValue, ",")

        Case Terget_Acc_Code    '勘定科目コードベース
            IsAggregateMatch = CsvContains(tbLine.account.code, masterLine.aggregateValue, ",")

        Case Else
            IsAggregateMatch = False

    End Select

End Function

'==============================================================================
'       分類対象判定（Filter / GROUP BY前段）
'------------------------------------------------------------------------------
'   Contents:
'       試算表1行（tbLine）が、財務諸表マスタ行の「分類対象」として扱われるか判定する。
'
'       ■ この関数の役割
'           ・Aggregateで抽出されたデータを「分類に使うかどうか」を決める（SQLのGROUP BY条件）
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
'       試算表行（tbLine）から、財務諸表出力用の分類キーを生成する（SQLのGROUP BYキー）
'
'       ■ この関数の役割
'           ・分類単位（accountingCategory）を決定
'           ・Dictionaryキーの一部になる
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

'******************************************************************************
'*                                                                            *
'*                          集計処理                                          *
'*                                                                            *
'******************************************************************************
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
'       isFormulaAfter      数式解決後かどうか
'
'==============================================================================
Public Sub RollupAggregateTo(ByVal includeFormula As Boolean)

    Dim line As FinancialStatementMasterLine
    Dim childType As String

    For Each line In m_fs_master.lines
        If line.enabled Then
            childType = UCase$(Trim$(line.lineType))

            'タイトル行は対象外、かつ親が定義されている場合のみ
            If childType <> "TITLE" And Trim$(line.aggregateTo) <> "" Then
                
                '------------------------------------------
                ' 対象判定（2段階集計）
                '------------------------------------------
                If (includeFormula And (childType = "FORMULA" Or childType = "CALC")) _
                    Or ((Not includeFormula) And (childType = "ACCOUNT" Or childType = "ACCOUNTS")) Then
                    Call RollupLineAmount(line)
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
'
'==============================================================================
Private Sub RollupLineAmount(ByVal child As FinancialStatementMasterLine)

    Dim parentLine As FinancialStatementMasterLine
    Dim childAmount As Currency
    Dim category As Variant

    For Each category In GetCategories(child.lineID)
        '------------------------------------------
        ' 子の金額取得
        '------------------------------------------
        childAmount = GetAmount(child.lineID, CStr(category))

        '親未定義なら次
        If Trim$(child.aggregateTo) <> "" Then
        
            '------------------------------------------
            ' 親行取得
            Set parentLine = GetAggregateParentLine(child.aggregateTo, CStr(category))
            
            '------------------------------------------
            ' 親へ加算
            AddAmount parentLine.lineID, CStr(category), childAmount
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
    If m_fs_master.Exists(normalizedParent, category) Then
        Set GetAggregateParentLine = m_fs_master.GetLine(normalizedParent, category)
        Exit Function
    End If

    '② 全体("*")
    If m_fs_master.Exists(normalizedParent, "*") Then
        Set GetAggregateParentLine = m_fs_master.GetLine(normalizedParent, "*")
        Exit Function
    End If

    '存在しない場合エラー
    Err.Raise vbObjectError + 7103, _
        "GetAggregateParentLine", _
        "AggregateTo target does not exist: " & normalizedParent & "|" & UCase$(Trim$(category))

End Function

'******************************************************************************
'*                                                                            *
'*                          計算処理                                          *
'*                                                                            *
'******************************************************************************
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
Public Sub ResolveFormulaAmounts(ByVal m_ColumnKeys As FinancialStatementColumnKeys)

    Dim line As FinancialStatementMasterLine
    Dim category As Variant
    Dim categories As Collection
    
    For Each line In m_fs_master.lines
        If line.enabled Then
            If line.lineType = "FORMULA" Or line.lineType = "CALC" Then
                Set categories = GetFormulaTargetCategories(line, m_ColumnKeys)
                For Each category In categories
                    SetAmount line.lineID, CStr(category), EvaluateFormula(line.formula, CStr(category))
                Next category
            End If
        End If
    Next

End Sub

'==============================================================================
'       Evaluate formula
'------------------------------------------------------------------------------
'   Contents:
'       Replace [FS_ID] / [FS_ID|category] tokens with values and evaluate.
'
'==============================================================================
Private Function EvaluateFormula(ByVal formula As String, ByVal defaultCategory As String) As Currency
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
        Dim lineID As String
        Dim category As String
        Dim val As Currency

        token = m.SubMatches(0)
        Call ParseReferenceToken(token, defaultCategory, lineID, category)

        If category = "" Then category = "*"

        val = GetAmount(lineID, category)

        expr = Replace(expr, "[" & token & "]", CStr(val))
    Next

    expr = Replace(expr, " ", "")
    EvaluateFormula = CCur(Application.Evaluate(expr))
End Function

'==============================================================================
'       Parse reference token
'------------------------------------------------------------------------------
'   Contents:
'       Parse a formula reference token into lineID/category.
'       Supported: FS_ID or FS_ID|category.
'       If category is omitted, defaultCategory is used.
'
'==============================================================================
Private Sub ParseReferenceToken(ByVal token As String, ByVal defaultCategory As String, ByRef lineID As String, ByRef category As String)
    If InStr(token, "|") > 0 Then
        lineID = UCase$(Trim$(Split(token, "|")(0)))
        category = UCase$(Trim$(Split(token, "|")(1)))
    Else
        lineID = UCase$(Trim$(token))
        category = UCase$(Trim$(defaultCategory))
    End If

    If category = "" Then category = "*"
End Sub

'==============================================================================
'       Get categories used for FORMULA/CALC evaluation.
'==============================================================================
Public Function GetFormulaTargetCategories(ByVal line As FinancialStatementMasterLine, ByVal columnKeys As FinancialStatementColumnKeys) As Collection

    Dim categories As Collection
    Dim mode As String
    Dim columns As Collection
    Dim col As Variant
    Dim categoryKey As String
    Dim seen As Object

    If Not line.IsEachLine Then
        Set categories = GetCategories(line.lineID)
        If categories.Count = 0 Then categories.Add "*"
        Set GetFormulaTargetCategories = categories
        Exit Function
    End If

    Set categories = New Collection
    Set seen = CreateObject("Scripting.Dictionary")

    mode = line.OutputCategoryMode
    Set columns = columnKeys.GetColumns(mode)

    For Each col In columns
        categoryKey = columnKeys.ResolveCategoryKey(CStr(col), mode)
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

