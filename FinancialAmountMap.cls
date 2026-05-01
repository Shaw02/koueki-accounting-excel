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
'       行ID＋区分別金額取得
'------------------------------------------------------------------------------
'   Contents:
'       lineID と category の組み合わせから、
'       amountMap に格納された金額を取得する。
'
'       キー未存在時は0を返却する。
'
'   Input:
'       lineID         財務諸表ID
'       category       区分キー
'
'   Output:
'       金額
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
'       行ID配下カテゴリ取得
'------------------------------------------------------------------------------
'   Contents:
'       指定lineIDに紐づくカテゴリ一覧を、金額Mapから抽出する。
'
'       金額MapのKey形式：
'           lineID|category
'
'       処理内容：
'           ・指定lineID配下を走査
'           ・category部分を抽出
'           ・重複除去
'
'       戻り値には、
'           "*"（全体カテゴリ）
'       を最低1件含める。
'
'       本処理は、
'           ・Rollup集計
'           ・数式カテゴリ解決
'           ・親子集計
'       等で使用される。
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
Public Function GetCategories(ByVal lineID As String) As Collection
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

