VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FinancialStatementMasterLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************
'*      メンバー変数                    *
'****************************************
Private m_StatementID As String           '帳票ID（貸借対照表、活動計算書など）
Private m_TargetSheet As String           '出力先シート名
Private m_LineID As String                '財務諸表行ID（内部識別子）
Private m_AggregateField As String
Private m_AggregateValue As String
Private m_FilterField As String
Private m_FilterValue As String
Private m_LineLabel As String             '表示用ラベル（科目名）
Private m_LineLabelAligh As XlHAlign      'ラベルの配置（左寄せなど）
Private m_LineType As String              '行種別（TITLE / ACCOUNT / FORMULA / CALC 等）
Private m_Formula As String               '計算式（FORMULA / CALC用）
Private m_AggregateTo As String           '集計先（親行ID）
Private m_Enabled As Boolean              '有効フラグ（出力対象か）

'------------------------------------------------------------------------------
'   LineType（行種別）
'
'   ■概要
'       財務諸表および注記における「行の役割」を定義する。
'       表示専用行／データ行／計算行／特殊レイアウト行を区別する。
'
'------------------------------------------------------------------------------
'   ■定義値
'
'   【① 表示専用（非集計）】
'
'       TITLE
'           ・見出し行
'           ・金額は表示しない
'           ・集計対象外
'
'       TEXT
'           ・任意テキスト出力行
'           ・例：「財務諸表に対する注記」など
'           ・金額は持たない
'
'
'   【② データ行（試算表連動）】
'
'       ACCOUNT
'           ・勘定科目（単体）
'           ・試算表から直接マッピングされる
'           ・集計対象（子→親へ加算）
'
'       ACCOUNTS
'           ・勘定科目グループ（複数のACCOUNTをまとめる）
'           ・ACCOUNTと同様に集計対象
'           ・ゼロ非表示などの制御単位としても使用
'
'
'   【③ 計算行】
'
'       FORMULA
'           ・数式による計算行
'           ・[FS_ID] 形式で他行を参照可能
'           ・ResolveFormula により計算される
'           ・第2段階の集計対象
'
'       CALC
'           ・補助計算用（FORMULAと同様）
'           ・中間計算専用
'           ・最終出力しない用途でも利用可
'
'
'   【④ 帳票ヘッダー】
'
'       HEADER_MAIN
'           ・貸借対照表／活動計算書のタイトル部を出力
'
'
'   【⑤ 注記ヘッダー（活動計算書）】
'
'       HEADER_PL_NOTE1
'           ・（１）財源区分別内訳表
'           ・横軸：一般／指定／合計
'
'
'       HEADER_PL_NOTE2
'           ・（２）会計区分及び事業区分別内訳
'
'           ■横軸（m_FilterField により切替）
'
'               CATEGORY_ACCOUNT
'                   ・会計区分単位（公1、公2 等）
'
'               CATEGORY_PROJECT
'                   ・補助科目（演奏会）単位
'
'           ■追加列
'               ・内部取引等消去（現時点は0固定）
'               ・合計
'
'
'       HEADER_PL_NOTE3
'           ・（３）事業費・管理費の形態別区分
'
'           ■横軸（m_FilterField により切替）
'
'               CATEGORY_ACCOUNT_SUM
'                   ・会計区分（集約単位）で表示
'                   ・例：公益目的事業会計／収益事業等会計／法人会計
'                   ・簡易表示用
'
'               CATEGORY_ACCOUNT
'                   ・会計区分単位（簡略表示）
'
'               CATEGORY_PROJECT
'                   ・補助科目（演奏会）単位
'
'           ■追加列
'               ・配賦基準（現時点は表示のみ）
'               ・合計
'
'
'------------------------------------------------------------------------------
'   ■注意事項
'
'       ・TITLE / TEXT は AggregateTo を持たない前提
'
'       ・ACCOUNT / ACCOUNTS は試算表と紐づく前提
'
'       ・FORMULA / CALC は formula 必須
'
'       ・HEADER_* 系はレイアウト専用行であり、
'         金額計算・集計処理の対象外とする
'
'       ・HEADER_PL_NOTE2 / NOTE3 は横軸可変（列数が動的）となる
'
'------------------------------------------------------------------------------

'==============================================================================
'       初期化処理
'------------------------------------------------------------------------------
'   Contents:
'       ワークシートから読み込んだ1行分のデータを、
'       財務諸表マスタ行として初期化する
'
'       ■主な処理
'       ・文字列のTrim / 正規化（大文字化）
'       ・accountingCategoryの補完（空 → "*"）
'       ・AggregateTo の自己参照チェックはProperty側で実施
'
'   Input:
'       statementID        帳票ID
'       TargetSheet        出力先シート名
'       lineID             行ID
'       accountingCategory 会計区分
'       LineLabel          表示ラベル
'       LineLabelAligh     表示位置
'       LineType           行種別
'       formula            数式
'       AggregateTo        集計先
'       enabled            有効フラグ
'
'==============================================================================
Public Sub Initialize( _
    ByVal statementID As String, _
    ByVal TargetSheet As String, _
    ByVal lineID As String, _
    ByVal aggregateField As String, _
    ByVal aggregateValue As String, _
    ByVal filterField As String, _
    ByVal filterValue As String, _
    ByVal LineLabel As String, _
    ByVal LineLabelAligh As XlHAlign, _
    ByVal lineType As String, _
    ByVal formula As String, _
    ByVal aggregateTo As String, _
    ByVal enabled As Boolean)

    m_StatementID = Trim$(statementID)
    m_TargetSheet = Trim$(TargetSheet)
    m_LineID = UCase$(Trim$(lineID))

    m_AggregateField = UCase$(Trim$(aggregateField))
    m_AggregateValue = NormalizeCsv(aggregateValue)
    m_FilterField = UCase$(Trim$(filterField))

    m_FilterValue = NormalizeCsv(filterValue)
    If m_FilterValue = "" Then m_FilterValue = "*"

    m_LineLabel = LineLabel
    m_LineLabelAligh = LineLabelAligh
    m_LineType = UCase$(Trim$(lineType))
    m_Formula = Trim$(formula)

    'AggregateTo はPropertyでバリデーション
    Me.aggregateTo = aggregateTo

    m_Enabled = enabled

End Sub

Private Function NormalizeCsv(ByVal value As String) As String
    Dim items As Variant
    Dim item As Variant
    Dim token As String
    Dim result As String
    Dim seen As Object

    Set seen = CreateObject("Scripting.Dictionary")
    items = Split(CStr(value), ",")

    For Each item In items
        token = UCase$(Trim$(CStr(item)))
        If token <> "" Then
            If Not seen.Exists(token) Then
                seen.Add token, True
                If result <> "" Then result = result & ","
                result = result & token
            End If
        End If
    Next item

    NormalizeCsv = result
End Function

'****************************************
'*      Property                        *
'****************************************

'帳票IDを取得する
Public Property Get statementID() As String: statementID = m_StatementID: End Property

'出力先シート名を取得する
Public Property Get TargetSheet() As String: TargetSheet = m_TargetSheet: End Property

'財務諸表行IDを取得する
Public Property Get lineID() As String: lineID = m_LineID: End Property

Public Property Get aggregateField() As String: aggregateField = m_AggregateField: End Property

Public Property Get aggregateValue() As String
    
    If (m_AggregateField = "" Or m_AggregateField = Terget_ID) And m_AggregateValue = "" Then
        aggregateValue = m_LineID
    Else
        aggregateValue = m_AggregateValue
    End If
    
End Property

Public Property Get filterField() As String: filterField = m_FilterField: End Property
Public Property Get filterValue() As String: filterValue = m_FilterValue: End Property

'表示ラベル（科目名）を取得する
Public Property Get LineLabel() As String: LineLabel = m_LineLabel: End Property

'ラベルの配置（左寄せ／中央など）を取得する
Public Property Get LineLabelAligh() As XlHAlign: LineLabelAligh = m_LineLabelAligh: End Property

'行種別を取得する（TITLE / ACCOUNT / FORMULA / CALC 等）
Public Property Get lineType() As String: lineType = m_LineType: End Property

'計算式を取得する（FORMULA / CALC用）
Public Property Get formula() As String: formula = m_Formula: End Property

'集計先（親行ID）を取得する
Public Property Get aggregateTo() As String: aggregateTo = m_AggregateTo: End Property

'集計先（親行ID）を設定する
Public Property Let aggregateTo(ByVal value As String)

    Dim normalized As String
    normalized = UCase$(Trim$(value))

    '------------------------------------------
    ' 自己参照チェック
    '------------------------------------------
    ' 自分自身を親に設定することは禁止
    ' → 無限ループ／集計破壊を防ぐ
    If normalized <> "" Then
        If normalized = m_LineID Then
            Err.Raise vbObjectError + 6401, _
                "FinancialStatementMasterLine.AggregateTo", _
                "AggregateTo cannot point to itself: " & normalized
        End If
    End If

    m_AggregateTo = normalized

End Property

'有効フラグを取得する（出力対象かどうか）
Public Property Get enabled() As Boolean: enabled = m_Enabled: End Property

