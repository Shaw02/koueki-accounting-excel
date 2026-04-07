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
Private m_AccountingCategory As String    '会計区分（公1、公2など）※空の場合は "*"
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
'       財務諸表上での行の役割を定義する
'
'   ■定義値
'       TITLE
'           ・見出し行
'           ・金額は表示しない
'           ・集計対象外
'
'       ACCOUNT
'           ・勘定科目（単体）
'           ・試算表から直接マッピングされる行
'           ・集計対象（子→親へ加算）
'
'       ACCOUNTS
'           ・勘定科目グループ（複数のACCOUNTをまとめる用途）
'           ・ACCOUNTと同様に集計対象
'           ・表示制御（ゼロ非表示など）で使われる場合あり
'
'       FORMULA
'           ・数式による計算行
'           ・[FS_ID] 形式で他行を参照可能
'           ・ResolveFormula により計算される
'           ・第2段階の集計対象
'
'       CALC
'           ・補助計算用（FORMULAと同様）
'           ・内部的な中間計算として使用可能
'           ・最終出力しない用途でも利用可
'
'   ■注意事項
'       ・TITLE は AggregateTo を持たない前提
'       ・FORMULA / CALC は formula 必須
'       ・ACCOUNT / ACCOUNTS は試算表と紐づく前提
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
    ByVal accountingCategory As String, _
    ByVal LineLabel As String, _
    ByVal LineLabelAligh As XlHAlign, _
    ByVal LineType As String, _
    ByVal formula As String, _
    ByVal AggregateTo As String, _
    ByVal enabled As Boolean)

    m_StatementID = Trim$(statementID)
    m_TargetSheet = Trim$(TargetSheet)
    m_LineID = Trim$(lineID)
    
    '会計区分は空の場合 "*"（全体）として扱う
    m_AccountingCategory = UCase$(Trim$(accountingCategory))
    If m_AccountingCategory = "" Then
        m_AccountingCategory = "*"
    End If

    m_LineLabel = LineLabel
    m_LineLabelAligh = LineLabelAligh
    m_LineType = UCase$(Trim$(LineType))
    m_Formula = Trim$(formula)

    'AggregateTo はPropertyでバリデーション
    Me.AggregateTo = AggregateTo

    m_Enabled = enabled

End Sub

'****************************************
'*      Property                        *
'****************************************

'帳票IDを取得する
Public Property Get statementID() As String: statementID = m_StatementID: End Property

'出力先シート名を取得する
Public Property Get TargetSheet() As String: TargetSheet = m_TargetSheet: End Property

'財務諸表行IDを取得する
Public Property Get lineID() As String: lineID = m_LineID: End Property

'会計区分を取得する（空の場合は "*"）
Public Property Get accountingCategory() As String: accountingCategory = m_AccountingCategory: End Property

'表示ラベル（科目名）を取得する
Public Property Get LineLabel() As String: LineLabel = m_LineLabel: End Property

'ラベルの配置（左寄せ／中央など）を取得する
Public Property Get LineLabelAligh() As XlHAlign: LineLabelAligh = m_LineLabelAligh: End Property

'行種別を取得する（TITLE / ACCOUNT / FORMULA / CALC 等）
Public Property Get LineType() As String: LineType = m_LineType: End Property

'計算式を取得する（FORMULA / CALC用）
Public Property Get formula() As String: formula = m_Formula: End Property

'集計先（親行ID）を取得する
Public Property Get AggregateTo() As String: AggregateTo = m_AggregateTo: End Property

'集計先（親行ID）を設定する
Public Property Let AggregateTo(ByVal value As String)

    Dim normalized As String
    normalized = UCase$(Trim$(value))

    '------------------------------------------
    ' 自己参照チェック
    '------------------------------------------
    ' 自分自身を親に設定することは禁止
    ' → 無限ループ／集計破壊を防ぐ
    If normalized <> "" Then
        If normalized = UCase$(Trim$(m_LineID)) Then
            Err.Raise vbObjectError + 6401, _
                "FinancialStatementMasterLine.AggregateTo", _
                "AggregateTo cannot point to itself: " & normalized
        End If
    End If

    m_AggregateTo = normalized

End Property

'有効フラグを取得する（出力対象かどうか）
Public Property Get enabled() As Boolean: enabled = m_Enabled: End Property

