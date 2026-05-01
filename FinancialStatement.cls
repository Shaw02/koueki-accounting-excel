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
    
    Dim wsMaster As Worksheet
    Set wsMaster = ThisWorkbook.Worksheets(MASTER_SHEET_NAME)
    
    'マスタ読込
    Call m_Master.LoadFromWorksheet(wsMaster)

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

    Call m_thisAmounts.Set_Master(m_Master, m_AccMaster)
    Call m_lastAmounts.Set_Master(m_Master, m_AccMaster)

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
'           ① 出力列キー初期化
'           ② 試算表 → 財務諸表金額マップ変換
'           ③ 純資産データ補完
'           ④ 勘定科目ベース集計
'           ⑤ 計算式解決
'           ⑥ 計算式反映後の再集計
'           ⑦ 財務諸表出力
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

    '① 会計区分・補助科目抽出
    Call m_ColumnKeys.Initialize(tbEngine.SubTrialBalance)
    
    '② 試算表→財務諸表マップ生成
    Call m_thisAmounts.BuildStatementAmountMap(tbEngine.SubTrialBalance)
    Call m_lastAmounts.BuildStatementAmountMap(tbEngine.LastSubTrialBalance)
    
    '③ 純資産データ追加
    Call m_thisAmounts.AddNetAssetAmounts( _
        tbEngine.GeneralNetAssetBegin, _
        tbEngine.GeneralNetAssetEnd, _
        tbEngine.DesignatedNetAssetBegin, _
        tbEngine.DesignatedNetAssetEnd)
    
    Call m_lastAmounts.AddNetAssetAmounts( _
        tbEngine.LastGeneralNetBegin, _
        tbEngine.LastGeneralNetEnd, _
        tbEngine.LastDesignatedNetBegin, _
        tbEngine.LastDesignatedNetEnd)
    
    '④ 勘定科目の集計
    Call m_thisAmounts.RollupAggregateTo(False)
    Call m_lastAmounts.RollupAggregateTo(False)

    '⑤ 計算式の解決
    Call m_thisAmounts.ResolveFormulaAmounts(m_ColumnKeys)
    Call m_lastAmounts.ResolveFormulaAmounts(m_ColumnKeys)

    '⑥ 計算式を含めた集計
    Call m_thisAmounts.RollupAggregateTo(True)
    Call m_lastAmounts.RollupAggregateTo(True)
        
    '⑦ 出力
    Call OutputAll(m_thisAmounts, m_lastAmounts)

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
                
                If line.IsEachLine Then
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

    mode = line.OutputCategoryMode
    Set columns = m_ColumnKeys.GetColumns(mode)

    c = 3
    totalValue = 0

    If mode = CATEGORY_ACCOUNT Or mode = CATEGORY_ACCOUNT_SUM Or mode = CATEGORY_PROJECT Then
        Dim targetGroups As Variant
        Dim grp As Variant
        Dim i As Long
        Dim hasGroupColumn As Boolean

        '------------------------------------------
        ' 会計区分固定順
        '------------------------------------------
        targetGroups = m_ColumnKeys.GetDisplayGroups()

        For Each grp In targetGroups
            hasGroupColumn = False
            '------------------------------------------
            ' 対象グループ列抽出
            '------------------------------------------
            For i = 1 To columns.Count
                If CStr(m_ColumnKeys.GetGroupLabel(CStr(columns(i)))) = CStr(grp) Then
                    categoryKey = m_ColumnKeys.ResolveCategoryKey(CStr(columns(i)), mode)
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
            categoryKey = m_ColumnKeys.ResolveCategoryKey(CStr(col), mode)
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
    
    Dim grp As String
    Dim targetGroups As Variant
    
    Dim it As Variant
    
    '表示項目
    targetGroups = m_ColumnKeys.GetDisplayGroups()
    mode = line.OutputCategoryMode

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
                If m_ColumnKeys.GetGroupLabel(CStr(columns(i))) = grp Then
                    groupCols.Add columns(i)
                End If
            Next i
            
            startCol = c
            
            '------------------------------------------
            ' データあり
            '------------------------------------------
            If groupCols.Count > 0 Then
                
                For Each col In groupCols
                    
                    'グループ
                    .Cells(headerRow, c).value = m_ColumnKeys.GetGroupLabel(CStr(col))
                    
                    '区分
                    .Cells(headerRow + 1, c).value = m_ColumnKeys.GetCategoryLabel(CStr(col))
                    
                    '補助科目名表示
                    If headerDepth = 3 Then
                        .Cells(headerRow + 2, c).value = m_ColumnKeys.GetProjectLabel(CStr(col))
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
    
    Dim grp As Variant
    Dim targetGroups As Variant
    Dim groupCols As Collection
    Dim col As Variant
    Dim startCol As Long

    '★固定グループ
    targetGroups = m_ColumnKeys.GetDisplayGroups()
    mode = line.OutputCategoryMode

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
                If m_ColumnKeys.GetGroupLabel(CStr(columns(i))) = grp Then
                    groupCols.Add columns(i)
                End If
            Next i
            
            startCol = c
            
            '------------------------------------------
            ' データあり
            '------------------------------------------
            If groupCols.Count > 0 Then
                
                For Each col In groupCols
                    
                    'グループ
                    .Cells(headerRow, c).value = m_ColumnKeys.GetGroupLabel(CStr(col))

                    '区分
                    If headerDepth >= 2 Then
                        .Cells(headerRow + 1, c).value = m_ColumnKeys.GetCategoryLabel(CStr(col))
                    End If
                    
                    '補助科目名
                    If headerDepth = 3 Then
                        .Cells(headerRow + 2, c).value = m_ColumnKeys.GetProjectLabel(CStr(col))
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
    isEach = line.IsEachLine

    If line.AlwaysVisible Then
        ShouldSkipOutputLine = False
        Exit Function
    End If

    'ゼロ行非表示
    If line.lineType = "ACCOUNT" Or line.lineType = "ACCOUNTS" Then
        If isEach Then
            ShouldSkipOutputLine = Not thisAmounts.HasAnyAmount(line, m_ColumnKeys)
            If ShouldSkipOutputLine Then Exit Function
        ElseIf vThis = 0 And vLast = 0 Then
            ShouldSkipOutputLine = True
            Exit Function
        End If
    End If
       
    ShouldSkipOutputLine = False  'デフォルトは表示
    
End Function

