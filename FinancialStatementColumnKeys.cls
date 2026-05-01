VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FinancialStatementColumnKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************
'*      メンバー変数                    *
'****************************************
Private m_ColumnKeysByCategory As Object

'==============================================================================
'       初期化処理
'------------------------------------------------------------------------------
'   Contents:
'
'
'==============================================================================
Public Sub Initialize(ByVal tb As TrialBalance)

    Call BuildColumnKeysByCategory(tb)

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

    m_ColumnKeysByCategory.Add CATEGORY_ACCOUNT, New Collection
    m_ColumnKeysByCategory.Add CATEGORY_PROJECT, New Collection
    m_ColumnKeysByCategory.Add CATEGORY_ACCOUNT_SUM, New Collection
    m_ColumnKeysByCategory.Add CATEGORY_NET_ASSET, New Collection
    m_ColumnKeysByCategory(CATEGORY_NET_ASSET).Add "一般"
    m_ColumnKeysByCategory(CATEGORY_NET_ASSET).Add "指定"
    For Each line In tb.lines
        If Not IsCollectibleProfitAndLossLine(line) Then
            GoTo ContinueLine
        End If

        code = Trim$(line.SubAccountCode)
        cate = Trim$(line.subAccountCate)
        proj = Trim$(line.SubAccountName)

        If cate <> "" Then
            If Not accountSeen.Exists(UCase$(cate)) Then
                accountSeen.Add UCase$(cate), True
                m_ColumnKeysByCategory(CATEGORY_ACCOUNT).Add _
                    GetAccountGroupLabel(cate) & "|" & cate
            End If

            grp = GetAccountGroupLabel(cate)
            If Not groupSeen.Exists(UCase$(grp)) Then
                groupSeen.Add UCase$(grp), True
                m_ColumnKeysByCategory(CATEGORY_ACCOUNT_SUM).Add grp
            End If

            If proj <> "" Then
                Dim projectKey As String
                projectKey = UCase$(cate) & "|" & UCase$(code)

                If Not projectSeen.Exists(projectKey) Then
                    projectSeen.Add projectKey, True
                    m_ColumnKeysByCategory(CATEGORY_PROJECT).Add _
                        GetAccountGroupLabel(cate) & "|" & cate & "|" & code & "|" & proj
                End If
            End If

        End If
ContinueLine:
    Next

    Call EnsureMandatoryAccountGroupColumns(groupSeen, m_ColumnKeysByCategory(CATEGORY_ACCOUNT_SUM))
    Call SortColumnKeysByCategory

End Sub

'==============================================================================
'       収集対象の活動計算書行判定
'------------------------------------------------------------------------------
'   Contents:
'       活動計算書(PL)かつゼロ金額ではない行のみ収集対象とする。
'
'==============================================================================
Private Function IsCollectibleProfitAndLossLine(ByVal line As TrialBalanceLine) As Boolean

    If line Is Nothing Then Exit Function
    If line.account Is Nothing Then Exit Function

    IsCollectibleProfitAndLossLine = (line.account.FS = "PL") And (Not line.IsZero)

End Function

'==============================================================================
'       勘定グループラベル取得
'------------------------------------------------------------------------------
'   Contents:
'       勘定区分文字列の先頭文字から、グループラベル（公益／収益／法人）を判定する
'
'   Input:
'       accountCategory    勘定区分文字列
'
'   Output:
'       GROUP_PUBLIC_LABEL     公益（先頭が「公」）
'       GROUP_REVENUE_LABEL    収益（先頭が「収」）
'       GROUP_CORPORATE_LABEL  法人（上記以外、「法」および空文字含む）
'
'==============================================================================
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

'==============================================================================
'       必須グループ列の補完
'------------------------------------------------------------------------------
'   Contents:
'       グループ列に、必須の3区分（公益／収益／法人）が存在するよう補完する
'
'   Input:
'       groupSeen      既存グループ管理用Dictionary
'       target         出力対象コレクション
'
'==============================================================================
Private Sub EnsureMandatoryAccountGroupColumns(ByVal groupSeen As Object, ByVal target As Collection)

    Call AddAccountGroupIfMissing(groupSeen, target, GROUP_PUBLIC_LABEL)
    Call AddAccountGroupIfMissing(groupSeen, target, GROUP_REVENUE_LABEL)
    Call AddAccountGroupIfMissing(groupSeen, target, GROUP_CORPORATE_LABEL)

End Sub

'==============================================================================
'       グループ存在チェック＆追加
'------------------------------------------------------------------------------
'   Contents:
'       指定グループが未登録の場合、DictionaryとCollectionへ追加する
'
'   Input:
'       groupSeen      既存グループ管理用Dictionary
'       target         出力対象コレクション
'       groupLabel     追加対象グループラベル
'
'==============================================================================
Private Sub AddAccountGroupIfMissing(ByVal groupSeen As Object, ByVal target As Collection, ByVal groupLabel As String)

    If Not groupSeen.Exists(UCase$(groupLabel)) Then
        groupSeen.Add UCase$(groupLabel), True
        target.Add groupLabel
    End If

End Sub


'==============================================================================
'       Resolve aggregate category key from columnKey.
'==============================================================================
Public Function ResolveCategoryKey(ByVal columnKey As String, ByVal mode As String) As String

    Dim parts() As String
    parts = Split(columnKey, "|")

    Select Case mode
        Case CATEGORY_ACCOUNT
            If UBound(parts) >= 1 Then
                ResolveCategoryKey = UCase$(Trim$(parts(1)))
            Else
                ResolveCategoryKey = UCase$(Trim$(columnKey))
            End If
        Case CATEGORY_PROJECT
            If UBound(parts) >= 2 Then
                ResolveCategoryKey = UCase$(Trim$(parts(2)))
            Else
                ResolveCategoryKey = UCase$(Trim$(columnKey))
            End If
        Case Else
            ResolveCategoryKey = UCase$(Trim$(parts(0)))
    End Select

End Function

'==============================================================================
'       Return columns grouped by display group.
'==============================================================================
Public Function GetGroupedColumns(ByVal mode As String) As Object

    Dim result As Object
    Dim cols As Collection
    Dim col As Variant
    Dim parts() As String
    Dim grp As String

    Set result = CreateObject("Scripting.Dictionary")
    Set cols = GetColumns(mode)

    For Each col In cols
        parts = Split(CStr(col), "|")
        grp = parts(0)
        If Not result.Exists(grp) Then
            Set result(grp) = New Collection
        End If
        result(grp).Add CStr(col)
    Next col

    Set GetGroupedColumns = result

End Function

'==============================================================================
'       Return display groups in fixed order.
'==============================================================================
Public Function GetDisplayGroups() As Variant
    GetDisplayGroups = Array(GROUP_PUBLIC_LABEL, GROUP_REVENUE_LABEL, GROUP_CORPORATE_LABEL)
End Function

Public Function GetGroupLabel(ByVal columnKey As String) As String
    GetGroupLabel = GetPartByIndex(columnKey, 0)
End Function

Public Function GetCategoryLabel(ByVal columnKey As String) As String
    GetCategoryLabel = GetPartByIndex(columnKey, 1)
End Function

Public Function GetProjectLabel(ByVal columnKey As String) As String
    Dim p3 As String
    p3 = GetPartByIndex(columnKey, 3)
    If p3 <> "" Then
        GetProjectLabel = p3
    Else
        GetProjectLabel = GetPartByIndex(columnKey, 2)
    End If
End Function

Private Function GetPartByIndex(ByVal columnKey As String, ByVal index As Long) As String
    Dim parts() As String
    parts = Split(columnKey, "|")
    If index <= UBound(parts) Then
        GetPartByIndex = parts(index)
    Else
        GetPartByIndex = ""
    End If
End Function

'==============================================================================
'       列キーのカテゴリ別ソート
'------------------------------------------------------------------------------
'   Contents:
'       各カテゴリ（ACCOUNT_SUM / ACCOUNT / PROJECT）ごとに、
'       列キーコレクションをソートして再設定する
'
'==============================================================================
Private Sub SortColumnKeysByCategory()

    Dim sorted As Collection

    Set sorted = SortColumnKeyCollection(m_ColumnKeysByCategory(CATEGORY_ACCOUNT_SUM), CATEGORY_ACCOUNT_SUM)
    Set m_ColumnKeysByCategory(CATEGORY_ACCOUNT_SUM) = sorted

    Set sorted = SortColumnKeyCollection(m_ColumnKeysByCategory(CATEGORY_ACCOUNT), CATEGORY_ACCOUNT)
    Set m_ColumnKeysByCategory(CATEGORY_ACCOUNT) = sorted

    Set sorted = SortColumnKeyCollection(m_ColumnKeysByCategory(CATEGORY_PROJECT), CATEGORY_PROJECT)
    Set m_ColumnKeysByCategory(CATEGORY_PROJECT) = sorted

End Sub

'==============================================================================
'       列キーコレクションのソート
'------------------------------------------------------------------------------
'   Contents:
'       列キー（"|"区切り文字列）を配列に変換し、
'       指定モードに応じたルールで並び替えた後、Collectionとして返す
'
'   Input:
'       source      元コレクション
'       mode        ソートモード
'
'   Output:
'       ソート済コレクション
'
'==============================================================================
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

'==============================================================================
'       列キー交換判定
'------------------------------------------------------------------------------
'   Contents:
'       2つの列キーを比較し、並び替えが必要か判定する
'
'       判定順序：
'           ① グループ順
'           ② 会計区分順（ACCOUNT_SUM以外）
'           ③ 補助科目順（PROJECTのみ）
'
'   Input:
'       leftKey     左側キー
'       rightKey    右側キー
'       mode        ソートモード
'
'   Output:
'       True：入れ替え
'       False：維持
'
'==============================================================================
Private Function ShouldSwapColumnKey(ByVal leftKey As String, ByVal rightKey As String, ByVal mode As String) As Boolean

    Dim leftParts() As String
    Dim rightParts() As String
    Dim cmp As Long

    leftParts = Split(leftKey, "|")
    rightParts = Split(rightKey, "|")

    cmp = CompareGroupLabel(leftParts(0), rightParts(0))
    If cmp <> 0 Then
        ShouldSwapColumnKey = (cmp > 0)
        Exit Function
    End If

    If mode = CATEGORY_ACCOUNT_SUM Then
        ShouldSwapColumnKey = False
        Exit Function
    End If

    cmp = CompareAccountCategoryByGroup(GetPartValue(leftParts, 1), GetPartValue(rightParts, 1), leftParts(0))
    If cmp <> 0 Then
        ShouldSwapColumnKey = (cmp > 0)
        Exit Function
    End If

    If mode = CATEGORY_PROJECT Then
        ShouldSwapColumnKey = (UCase$(GetPartValue(leftParts, 2)) > UCase$(GetPartValue(rightParts, 2)))
    Else
        ShouldSwapColumnKey = False
    End If

End Function

'==============================================================================
'       キー分割値取得
'------------------------------------------------------------------------------
'   Contents:
'       分割済配列から指定インデックスの値を取得する
'       範囲外の場合は空文字を返す
'
'   Input:
'       parts   分割済配列
'       index   取得位置
'
'   Output:
'       指定要素または空文字
'
'==============================================================================
Private Function GetPartValue(ByRef parts() As String, ByVal index As Long) As String

    If index <= UBound(parts) Then
        GetPartValue = parts(index)
    Else
        GetPartValue = ""
    End If

End Function

'==============================================================================
'       グループラベル比較
'------------------------------------------------------------------------------
'   Contents:
'       グループラベルを優先順位に変換し、差分を返す
'
'   Input:
'       leftLabel   左ラベル
'       rightLabel  右ラベル
'
'   Output:
'       比較結果（数値）
'
'==============================================================================
Private Function CompareGroupLabel(ByVal leftLabel As String, ByVal rightLabel As String) As Long

    CompareGroupLabel = GroupPriorityFromLabel(leftLabel) - GroupPriorityFromLabel(rightLabel)

End Function

'==============================================================================
'       グループ優先順位取得
'------------------------------------------------------------------------------
'   Contents:
'       グループラベルを内部優先順位（数値）へ変換する
'
'   Input:
'       groupLabel  グループラベル
'
'   Output:
'       優先順位（数値）
'
'==============================================================================
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

'==============================================================================
'       会計区分比較（グループ内）
'------------------------------------------------------------------------------
'   Contents:
'       同一グループ内での会計区分の並び順を決定する
'
'       判定順序：
'           ① 共通区分（"共"）は後ろ
'           ② 数値順
'           ③ 文字列比較
'
'   Input:
'       leftCategory    左区分
'       rightCategory   右区分
'       groupLabel      グループラベル
'
'   Output:
'       比較結果（数値）
'
'==============================================================================
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

'==============================================================================
'       区分番号抽出
'------------------------------------------------------------------------------
'   Contents:
'       区分名から数値部分を抽出し、ソート用数値として返す
'       「共」を含む場合は最大値（999）とする
'
'   Input:
'       categoryName    区分名
'
'   Output:
'       数値
'
'==============================================================================
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

'==============================================================================
'       注記列取得
'------------------------------------------------------------------------------
'   Contents:
'       指定モードに対応する列キーコレクションを取得する
'       未存在の場合は空コレクションを返す
'
'   Input:
'       mode    カテゴリモード
'
'   Output:
'       列キーコレクション
'
'==============================================================================
Public Function GetColumns(ByVal category As String) As Collection

    If m_ColumnKeysByCategory Is Nothing Then
        Set m_ColumnKeysByCategory = CreateObject("Scripting.Dictionary")
    End If

    If m_ColumnKeysByCategory.Exists(category) Then
        Set GetColumns = m_ColumnKeysByCategory(category)
    Else
        Set GetColumns = New Collection
    End If

End Function


