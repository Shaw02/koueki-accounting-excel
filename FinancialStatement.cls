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
Private m_Master As FinancialStatementMaster                    '財務諸表マスタクラス
Private m_ColumnKeys As FinancialStatementColumnKeys

'==============================================================================
'       初期化処理
'------------------------------------------------------------------------------
'   Contents:
'       クラス生成時に財務諸表マスタクラスを初期化する
'
'==============================================================================
Private Sub Class_Initialize()
    Set m_Master = New FinancialStatementMaster
    Set m_ColumnKeys = New FinancialStatementColumnKeys
End Sub

'==============================================================================
'       キー生成（lineID + category）
'------------------------------------------------------------------------------
'   Contents:
'       財務諸表金額マップで使用するキーを生成する。
'
'       ■ 形式
'           lineID | category
'
'       ■ 仕様
'           ・category が空の場合は "*" に補正
'           ・大文字 + Trim により正規化
'
'       ■ 注意
'           ・キー仕様は全体で統一されている前提
'==============================================================================
Private Function BuildKey(ByVal lineID As String, ByVal category As String) As String
    If Trim$(category) = "" Then category = "*"
    BuildKey = UCase$(Trim$(lineID)) & "|" & UCase$(Trim$(category))
End Function

'==============================================================================
'       財務諸表生成メイン処理
'------------------------------------------------------------------------------
'   Contents:
'       試算表データをもとに財務諸表を生成し、Excelへ出力する。
'
'       ■ 処理フロー
'           ① マスタ読込
'           ② 試算表 → 金額マップ変換
'           ③ 純資産補完
'           ④ 勘定ベース集計
'           ⑤ 計算式解決
'           ⑥ 再集計
'           ⑦ 出力
'
'       ■ 設計ポイント
'           ・2段階集計（実データ → 計算式反映）
'==============================================================================
Public Sub BuildAndOutput(ByVal tbEngine As TrialBalanceEngine)

    Dim thisAmounts As Object
    Dim lastAmounts As Object
    
    Dim wsMaster As Worksheet
    Set wsMaster = ThisWorkbook.Worksheets(MASTER_SHEET_NAME)
    
    'マスタ読込
    Call m_Master.LoadFromWorksheet(wsMaster)

    '会計区分・補助科目抽出
    Call m_ColumnKeys.Initialize(tbEngine.SubTrialBalance)
    
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
'       試算表 → 財務諸表金額マップ変換
'------------------------------------------------------------------------------
'   Contents:
'       試算表データを、財務諸表行単位の金額へ変換する。
'
'       ■ 出力
'           Dictionary（lineID|category → 金額）
'
'       ■ 処理概要
'           ・試算表 × マスタ を突き合わせ
'           ・Aggregate で抽出
'           ・Filter で分類
'           ・AddAmountToMap で蓄積
'
'       ■ 分類パターン
'           "*" / EACH / 固定指定
'==============================================================================
Private Function BuildStatementAmountMap(ByVal tb As TrialBalance) As Object

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim tbLine As TrialBalanceLine
    Dim masterLine As FinancialStatementMasterLine
    Dim amount As Currency
    Dim detailKey As String
    Dim filterTokens As Variant
    Dim filterToken As Variant

    For Each tbLine In tb.lines
        amount = tbLine.amount()

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
                        Call AddAmountToMap(result, masterLine.lineID, "*", amount)

                        If UCase$(Trim$(masterLine.filterValue)) = "EACH" Then
                            '------------------------------------------
                            ' EACH指定：
                            ' 試算表の値ごとに明細分解（補助科目・事業など）
                            '------------------------------------------
                            detailKey = GetFilterKey(masterLine.filterField, tbLine)
                            If detailKey <> "" And detailKey <> "*" Then
                                Call AddAmountToMap(result, masterLine.lineID, detailKey, amount)
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
                                    Call AddAmountToMap(result, masterLine.lineID, detailKey, amount)
                                End If
                            Next filterToken
                        End If
                    End If
                End If
            End If
        Next masterLine
    Next tbLine

    Set BuildStatementAmountMap = result

End Function

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
'       ■ 設計ポイント
'           ・マスタ設定のみで抽出条件を変更可能
'           ・一致ロジックは CsvContains に集約
'
'       ■ 注意（重要）
'           ・aggregateField 未指定時は Terget_ID が適用される
'           ・区切り文字はフィールドごとに異なる（; / ,）
'           ・未対応フィールドは False（安全側）
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
'       金額マップ加算処理（最終集計単位）
'------------------------------------------------------------------------------
'   Contents:
'       Dictionary（lineID|category）に金額を加算する。
'
'       ■ この関数の役割
'           ・全ての集計結果がここに集約される
'           ・最終的な財務諸表の数値を構成する基盤
'
'       ■ 動作
'           ・キーがなければ新規作成（0初期化）
'           ・既存なら加算
'
'       ■ キー構造
'           lineID | accountingCategory
'
'       ■ 注意（重要）
'           ・キー生成ロジックはここに固定されている
'           ・ここが壊れると全体集計が壊れる
'
'==============================================================================
Private Sub AddAmountToMap( _
    ByVal accountMap As Object, _
    ByVal lineID As String, _
    ByVal accountingCategory As String, _
    ByVal amount As Currency)

    '------------------------------------------
    ' 一意キー生成（ロジックは BuildKey に集約）
    Dim key As String
    key = UCase$(Trim$(lineID)) & "|" & UCase$(Trim$(accountingCategory))

    '------------------------------------------
    ' 既存キーなら加算、なければ新規作成
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
        Call AddAmountToMap(accountMap, "BS_NET_START", "一般", tbEngine.LastGeneralNetBegin)
        Call AddAmountToMap(accountMap, "BS_NET_GENERAL_END", "一般", tbEngine.LastGeneralNetEnd)
        
        Call AddAmountToMap(accountMap, "BS_NET_START", "*", tbEngine.LastGeneralNetBegin)
        Call AddAmountToMap(accountMap, "BS_NET_GENERAL_END", "*", tbEngine.LastGeneralNetEnd)

        '==============================
        ' 指定純資産
        '==============================
        Call AddAmountToMap(accountMap, "BS_NET_START", "指定", tbEngine.LastDesignatedNetBegin)
        Call AddAmountToMap(accountMap, "BS_NET_DESIG_END", "指定", tbEngine.LastDesignatedNetEnd)
        
        Call AddAmountToMap(accountMap, "BS_NET_START", "*", tbEngine.LastDesignatedNetBegin)
        Call AddAmountToMap(accountMap, "BS_NET_DESIG_END", "*", tbEngine.LastDesignatedNetEnd)

    Else

        '==============================
        ' 一般純資産
        '==============================
        Call AddAmountToMap(accountMap, "BS_NET_START", "一般", tbEngine.GeneralNetAssetBegin)
        Call AddAmountToMap(accountMap, "BS_NET_GENERAL_END", "一般", tbEngine.GeneralNetAssetEnd)

        Call AddAmountToMap(accountMap, "BS_NET_START", "*", tbEngine.GeneralNetAssetBegin)
        Call AddAmountToMap(accountMap, "BS_NET_GENERAL_END", "*", tbEngine.GeneralNetAssetEnd)

        '==============================
        ' 指定純資産
        '==============================
        Call AddAmountToMap(accountMap, "BS_NET_START", "指定", tbEngine.DesignatedNetAssetBegin)
        Call AddAmountToMap(accountMap, "BS_NET_DESIG_END", "指定", tbEngine.DesignatedNetAssetEnd)

        Call AddAmountToMap(accountMap, "BS_NET_START", "*", tbEngine.DesignatedNetAssetBegin)
        Call AddAmountToMap(accountMap, "BS_NET_DESIG_END", "*", tbEngine.DesignatedNetAssetEnd)

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
    For Each line In m_Master.lines
    
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
            If Not ShouldSkipOutputLine(line, vThis, vLast) Then
                
                'タイトル行は金額を表示しない
                If line.lineType = "TITLE" Then
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
ContinueLineLoop:
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

        '数値フォーマット（△表示）
        .Range("C:E").NumberFormat = "_ * #,##0_ ;_ * △#,##0_ ;_ * ""-""??_ ;_ @_ "
        .Range("C:E").HorizontalAlignment = xlRight
    
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
Private Function ShouldSkipOutputLine(ByVal line As FinancialStatementMasterLine, ByVal vThis As Currency, ByVal vLast As Currency) As Boolean

    If line.lineType = "TITLE" _
        Or line.lineType = "ACCOUNTS" _
        Or line.lineType = "HEADER_MAIN" _
        Or line.lineType = "TEXT" _
        Or line.lineType = "NOTE_PL_FUND" _
        Or line.lineType = "NOTE_PL_ACCOUNT_PROJECT" _
        Or line.lineType = "NOTE_PL_NATURE" Then
        
        ShouldSkipOutputLine = False
        Exit Function
    End If

    'ゼロ行非表示
    If (line.lineType = "ACCOUNT" Or line.lineType = "ACCOUNTS") And vThis = 0 And vLast = 0 Then
        ShouldSkipOutputLine = True
        Exit Function
    End If
       
    ShouldSkipOutputLine = False  'デフォルトは表示
    
End Function

'==============================================================================
'
'------------------------------------------------------------------------------
'   Contents:
'
'
'==============================================================================
Private Function GetLineCategoriesFromAmountMap(ByVal values As Object, ByVal lineID As String) As Collection

    Dim result As New Collection
    Dim seen As Object
    Dim key As Variant
    Dim normalizedLineID As String
    Dim suffix As String
    Dim category As String

    Set seen = CreateObject("Scripting.Dictionary")
    normalizedLineID = UCase$(Trim$(lineID))

    For Each key In values.Keys
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

    Set GetLineCategoriesFromAmountMap = result

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
    key = BuildKey(line.lineID, line.filterValue)

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
    Dim category As Variant
    Dim key As String

    For Each line In m_Master.lines
        If line.enabled Then
            If line.lineType = "FORMULA" Or line.lineType = "CALC" Then
                For Each category In GetLineCategoriesFromAmountMap(values, line.lineID)
                    'キー生成（lineID|category）
                    key = BuildKey(line.lineID, CStr(category))
                    '数式評価
                    values(key) = ResolveFormulaByCategory(line.formula, values, CStr(category))
                Next category
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
Private Sub RollupLineAmount(ByVal child As FinancialStatementMasterLine, ByVal amountMap As Object)

    Dim parentLine As FinancialStatementMasterLine
    Dim childKey As String
    Dim parentKey As String
    Dim childAmount As Currency
    Dim category As Variant

    
    For Each category In GetLineCategoriesFromAmountMap(amountMap, child.lineID)
        '------------------------------------------
        ' 子の金額取得
        '------------------------------------------
        childKey = BuildKey(child.lineID, CStr(category))
        If amountMap.Exists(childKey) Then
            childAmount = CCur(amountMap(childKey))
        Else
            childAmount = 0
        End If

        '親未定義なら次
        If Trim$(child.aggregateTo) <> "" Then
        
            '------------------------------------------
            ' 親行取得
            '------------------------------------------
            Set parentLine = GetAggregateParentLine(child.aggregateTo, CStr(category))
            parentKey = BuildKey(parentLine.lineID, CStr(category))
    
            '------------------------------------------
            ' 親へ加算
            '------------------------------------------
            If Not amountMap.Exists(parentKey) Then amountMap.Add parentKey, CCur(0)
            amountMap(parentKey) = CCur(amountMap(parentKey)) + childAmount
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
Private Function ResolveFormulaByCategory(ByVal formula As String, ByVal values As Object, ByVal defaultCategory As String) As Currency

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

        Dim key As String
        key = BuildKey(lineID, category)

        Dim val As Currency
        If values.Exists(key) Then
            val = CCur(values(key))
        ElseIf category <> "*" And values.Exists(BuildKey(lineID, "*")) Then
            val = CCur(values(BuildKey(lineID, "*")))
        Else
            val = 0
        End If

        expr = Replace(expr, "[" & token & "]", CStr(val))
    Next

    expr = Replace(expr, " ", "")
    ResolveFormulaByCategory = CCur(Application.Evaluate(expr))

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


