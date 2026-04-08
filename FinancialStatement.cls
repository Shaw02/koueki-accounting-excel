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
Private Const CATEGORY_EACH_ACCOUNT_SUM As String = "EACH_ACCOUNT_SUM"
Private Const CATEGORY_EACH_ACCOUNT As String = "EACH_ACCOUNT"
Private Const CATEGORY_EACH_PROJECT As String = "EACH_PROJECT"
Private Const CATEGORY_EACH_FUND As String = "EACH_FUND"
Private Const GROUP_PUBLIC As Long = 1
Private Const GROUP_REVENUE As Long = 2
Private Const GROUP_CORPORATE As Long = 3
Private Const GROUP_PUBLIC_LABEL As String = "公益目的事業会計"
Private Const GROUP_REVENUE_LABEL As String = "収益事業会計"
Private Const GROUP_CORPORATE_LABEL As String = "法人会計"

'****************************************
'*      メンバー変数                    *
'****************************************
Private m_Master As FinancialStatementMaster                    '財務諸表マスタクラス
Private m_ColumnKeysByCategory As Object

'==============================================================================
'       初期化処理
'------------------------------------------------------------------------------
'   Contents:
'       クラス生成時に財務諸表マスタクラスを初期化する
'
'==============================================================================
Private Sub Class_Initialize()
    Set m_Master = New FinancialStatementMaster
    Set m_ColumnKeysByCategory = CreateObject("Scripting.Dictionary")
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

    Dim thisAmounts As Object
    Dim lastAmounts As Object

    
    Dim wsMaster As Worksheet
    Set wsMaster = ThisWorkbook.Worksheets(MASTER_SHEET_NAME)
    
    'マスタ読込
    Call m_Master.LoadFromWorksheet(wsMaster)

    '会計区分・補助科目抽出
    Call BuildColumnKeysByCategory(tbEngine.SubTrialBalance)
    
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

    For Each line In m_Master.Lines
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

    mode = UCase$(Trim$(line.accountingCategory))
    If mode <> CATEGORY_EACH_PROJECT Then mode = CATEGORY_EACH_ACCOUNT

    Set columns = GetNoteColumns(mode)

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
        headerDepth = IIf(mode = CATEGORY_EACH_PROJECT, 3, 2)

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

    mode = UCase$(Trim$(line.accountingCategory))
    If mode <> CATEGORY_EACH_ACCOUNT_SUM And mode <> CATEGORY_EACH_ACCOUNT And mode <> CATEGORY_EACH_PROJECT Then
        mode = CATEGORY_EACH_ACCOUNT_SUM
    End If

    Set columns = GetNoteColumns(mode)

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
            Case CATEGORY_EACH_PROJECT: headerDepth = 3
            Case CATEGORY_EACH_ACCOUNT: headerDepth = 2
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
'       補助科目・会計区分の列キー生成
'------------------------------------------------------------------------------
'   Contents:
'       試算表から会計区分・補助科目（事業）を抽出し、
'       注記用の横軸キーを生成する
'
'       ※費用・収益のみ対象とする
'
'==============================================================================
Private Sub BuildColumnKeysByCategory(ByVal tb As TrialBalance)

    Dim accountSeen As Object
    Dim projectSeen As Object
    Dim groupSeen As Object
    Dim line As TrialBalanceLine
    Dim code As String
    Dim cate As String
    Dim proj As String
    Dim grp As String

    Set m_ColumnKeysByCategory = CreateObject("Scripting.Dictionary")

    Set accountSeen = CreateObject("Scripting.Dictionary")
    Set projectSeen = CreateObject("Scripting.Dictionary")
    Set groupSeen = CreateObject("Scripting.Dictionary")

    m_ColumnKeysByCategory.Add CATEGORY_EACH_ACCOUNT, New Collection
    m_ColumnKeysByCategory.Add CATEGORY_EACH_PROJECT, New Collection
    m_ColumnKeysByCategory.Add CATEGORY_EACH_ACCOUNT_SUM, New Collection

    For Each line In tb.Lines

        '------------------------------------------
        ' ■対象判定：費用・収益のみ
        '------------------------------------------
        If (Not line.account.FS = "PL") Or line.IsZero Then
            GoTo ContinueLine
        End If

        code = Trim$(line.SubAccountCode)   '★追加：補助科目コード
        cate = Trim$(line.SubAccountCate)
        proj = Trim$(line.SubAccountName)

        If cate <> "" Then

            '------------------------------------------
            ' 会計区分
            '------------------------------------------
            If Not accountSeen.Exists(UCase$(cate)) Then
                accountSeen.Add UCase$(cate), True
                m_ColumnKeysByCategory(CATEGORY_EACH_ACCOUNT).Add _
                    GetAccountGroupLabel(cate) & "|" & cate
            End If

            '------------------------------------------
            ' グループ
            '------------------------------------------
            grp = GetAccountGroupLabel(cate)
            If Not groupSeen.Exists(UCase$(grp)) Then
                groupSeen.Add UCase$(grp), True
                m_ColumnKeysByCategory(CATEGORY_EACH_ACCOUNT_SUM).Add grp
            End If

            '------------------------------------------
            ' 事業（★ここ改修）
            '------------------------------------------
            If proj <> "" Then
                Dim projectKey As String
                projectKey = UCase$(cate) & "|" & UCase$(code)

                If Not projectSeen.Exists(projectKey) Then
                    projectSeen.Add projectKey, True

                    ' ★コードを第3要素に追加
                    m_ColumnKeysByCategory(CATEGORY_EACH_PROJECT).Add _
                        GetAccountGroupLabel(cate) & "|" & cate & "|" & code & "|" & proj
                End If
            End If

        End If

ContinueLine:
    Next

    Call EnsureMandatoryAccountGroupColumns(groupSeen, m_ColumnKeysByCategory(CATEGORY_EACH_ACCOUNT_SUM))
    Call SortColumnKeysByCategory

End Sub

Private Function GetAccountGroupLabel(ByVal accountCategory As String) As String

    Dim key As String
    key = Left$(Trim$(accountCategory), 1)

    Select Case key
        Case "公"
            GetAccountGroupLabel = GROUP_PUBLIC_LABEL
        Case "収"
            GetAccountGroupLabel = GROUP_REVENUE_LABEL
        Case "法", ""
            GetAccountGroupLabel = GROUP_CORPORATE_LABEL
        Case Else
            GetAccountGroupLabel = GROUP_CORPORATE_LABEL
    End Select

End Function

Private Sub EnsureMandatoryAccountGroupColumns(ByVal groupSeen As Object, ByVal target As Collection)

    Call AddAccountGroupIfMissing(groupSeen, target, GROUP_PUBLIC_LABEL)
    Call AddAccountGroupIfMissing(groupSeen, target, GROUP_REVENUE_LABEL)
    Call AddAccountGroupIfMissing(groupSeen, target, GROUP_CORPORATE_LABEL)

End Sub

Private Sub AddAccountGroupIfMissing(ByVal groupSeen As Object, ByVal target As Collection, ByVal groupLabel As String)

    If Not groupSeen.Exists(UCase$(groupLabel)) Then
        groupSeen.Add UCase$(groupLabel), True
        target.Add groupLabel
    End If

End Sub

Private Sub SortColumnKeysByCategory()

    Dim sorted As Collection

    Set sorted = SortColumnKeyCollection(m_ColumnKeysByCategory(CATEGORY_EACH_ACCOUNT_SUM), CATEGORY_EACH_ACCOUNT_SUM)
    Set m_ColumnKeysByCategory(CATEGORY_EACH_ACCOUNT_SUM) = sorted

    Set sorted = SortColumnKeyCollection(m_ColumnKeysByCategory(CATEGORY_EACH_ACCOUNT), CATEGORY_EACH_ACCOUNT)
    Set m_ColumnKeysByCategory(CATEGORY_EACH_ACCOUNT) = sorted

    Set sorted = SortColumnKeyCollection(m_ColumnKeysByCategory(CATEGORY_EACH_PROJECT), CATEGORY_EACH_PROJECT)
    Set m_ColumnKeysByCategory(CATEGORY_EACH_PROJECT) = sorted

End Sub

Private Function SortColumnKeyCollection(ByVal source As Collection, ByVal mode As String) As Collection

    Dim sorted As Collection
    Dim values() As String
    Dim i As Long
    Dim j As Long
    Dim tmp As String

    Set sorted = New Collection

    If source Is Nothing Then
        Set SortColumnKeyCollection = sorted
        Exit Function
    End If

    If source.Count = 0 Then
        Set SortColumnKeyCollection = sorted
        Exit Function
    End If

    ReDim values(1 To source.Count)
    For i = 1 To source.Count
        values(i) = CStr(source(i))
    Next

    For i = 1 To UBound(values) - 1
        For j = i + 1 To UBound(values)
            If ShouldSwapColumnKey(values(i), values(j), mode) Then
                tmp = values(i)
                values(i) = values(j)
                values(j) = tmp
            End If
        Next
    Next

    For i = 1 To UBound(values)
        sorted.Add values(i)
    Next

    Set SortColumnKeyCollection = sorted

End Function

Private Function ShouldSwapColumnKey(ByVal leftKey As String, ByVal rightKey As String, ByVal mode As String) As Boolean

    Dim leftParts() As String
    Dim rightParts() As String
    Dim cmp As Long

    leftParts = Split(leftKey, "|")
    rightParts = Split(rightKey, "|")

    '------------------------------------------
    ' グループ比較
    '------------------------------------------
    cmp = CompareGroupLabel(leftParts(0), rightParts(0))
    If cmp <> 0 Then
        ShouldSwapColumnKey = (cmp > 0)
        Exit Function
    End If

    If mode = CATEGORY_EACH_ACCOUNT_SUM Then
        ShouldSwapColumnKey = False
        Exit Function
    End If

    '------------------------------------------
    ' 会計区分比較
    '------------------------------------------
    cmp = CompareAccountCategoryByGroup(GetPartValue(leftParts, 1), GetPartValue(rightParts, 1), leftParts(0))
    If cmp <> 0 Then
        ShouldSwapColumnKey = (cmp > 0)
        Exit Function
    End If

    '------------------------------------------
    ' ★補助科目ソート（コード優先）
    '------------------------------------------
    If mode = CATEGORY_EACH_PROJECT Then
        ShouldSwapColumnKey = (UCase$(GetPartValue(leftParts, 2)) > UCase$(GetPartValue(rightParts, 2)))
    Else
        ShouldSwapColumnKey = False
    End If

End Function

Private Function GetPartValue(ByRef parts() As String, ByVal index As Long) As String

    If index <= UBound(parts) Then
        GetPartValue = parts(index)
    Else
        GetPartValue = ""
    End If

End Function

Private Function CompareGroupLabel(ByVal leftLabel As String, ByVal rightLabel As String) As Long

    CompareGroupLabel = GroupPriorityFromLabel(leftLabel) - GroupPriorityFromLabel(rightLabel)

End Function

Private Function GroupPriorityFromLabel(ByVal groupLabel As String) As Long

    Select Case UCase$(Trim$(groupLabel))
        Case UCase$(GROUP_PUBLIC_LABEL)
            GroupPriorityFromLabel = GROUP_PUBLIC
        Case UCase$(GROUP_REVENUE_LABEL)
            GroupPriorityFromLabel = GROUP_REVENUE
        Case Else
            GroupPriorityFromLabel = GROUP_CORPORATE
    End Select

End Function

Private Function CompareAccountCategoryByGroup(ByVal leftCategory As String, ByVal rightCategory As String, ByVal groupLabel As String) As Long

    Dim leftShared As Boolean
    Dim rightShared As Boolean
    Dim leftNumber As Long
    Dim rightNumber As Long

    leftShared = (InStr(leftCategory, "共") > 0)
    rightShared = (InStr(rightCategory, "共") > 0)

    If leftShared Xor rightShared Then
        If leftShared Then
            CompareAccountCategoryByGroup = 1
        Else
            CompareAccountCategoryByGroup = -1
        End If
        Exit Function
    End If

    leftNumber = ExtractCategoryNumber(leftCategory)
    rightNumber = ExtractCategoryNumber(rightCategory)

    If leftNumber <> rightNumber Then
        CompareAccountCategoryByGroup = leftNumber - rightNumber
    Else
        CompareAccountCategoryByGroup = StrComp(UCase$(leftCategory), UCase$(rightCategory), vbTextCompare)
    End If

End Function

Private Function ExtractCategoryNumber(ByVal categoryName As String) As Long

    Dim i As Long
    Dim ch As String
    Dim digits As String

    If InStr(categoryName, "共") > 0 Then
        ExtractCategoryNumber = 999
        Exit Function
    End If

    digits = ""
    For i = 1 To Len(categoryName)
        ch = Mid$(categoryName, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        End If
    Next

    If digits = "" Then
        ExtractCategoryNumber = 0
    Else
        ExtractCategoryNumber = CLng(digits)
    End If

End Function

Private Function GetNoteColumns(ByVal mode As String) As Collection

    If m_ColumnKeysByCategory Is Nothing Then
        Set m_ColumnKeysByCategory = CreateObject("Scripting.Dictionary")
    End If

    If m_ColumnKeysByCategory.Exists(mode) Then
        Set GetNoteColumns = m_ColumnKeysByCategory(mode)
    Else
        Set GetNoteColumns = New Collection
    End If

End Function

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
            If line.lineType = "FORMULA" Or line.lineType = "CALC" Then
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
            childType = UCase$(Trim$(line.lineType))

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

