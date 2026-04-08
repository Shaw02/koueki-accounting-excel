VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FinancialStatementMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MASTER_START_ROW As Long = 2

Private m_Lines As Collection
Private m_LineMap As Object

Private Sub Class_Initialize()
    Set m_Lines = New Collection
    Set m_LineMap = CreateObject("Scripting.Dictionary")
End Sub

Public Sub Clear()
    Set m_Lines = New Collection
    Set m_LineMap = CreateObject("Scripting.Dictionary")
End Sub

Public Property Get Lines() As Collection
    Set Lines = m_Lines
End Property

'==============================================================================
'   キー生成関数
'------------------------------------------------------------------------------
'   Contents:
'       財務諸表マスタにおける内部キーを生成する。
'
'       ■設計要件
'       ・lineID単独では重複が発生するため使用不可
'       ・同一lineIDでも、会計区分（公1、公2等）ごとに別行として扱う必要がある
'       ・そのため「lineID + accountingCategory」を一意キーとする
'
'       ■例
'           PL_BIZ | 公1
'           PL_BIZ | 公2
'
'       ■備考
'           accountingCategoryが空の場合は "*" として扱う
'
'==============================================================================
Private Function BuildKey(ByVal lineID As String, ByVal category As String) As String

    Dim cat As String
    
    cat = UCase$(Trim$(category))
    If cat = "" Then cat = "*"
    
    BuildKey = UCase$(Trim$(lineID)) & "|" & cat

End Function

'==============================================================================
'   存在確認
'==============================================================================
Public Function Exists(ByVal lineID As String, Optional ByVal category As String = "*") As Boolean
    Exists = m_LineMap.Exists(BuildKey(lineID, category))
End Function

'==============================================================================
'   行取得
'==============================================================================
Public Function GetLine( _
    ByVal lineID As String, _
    Optional ByVal category As String = "*") _
    As FinancialStatementMasterLine

    Dim key As String
    key = BuildKey(lineID, category)

    If Not m_LineMap.Exists(key) Then
        Err.Raise vbObjectError + 6201, _
            "FinancialStatementMaster.GetLine", _
            "Line not found: " & key
    End If

    Set GetLine = m_LineMap(key)

End Function

'==============================================================================
'   ワークシートから読込
'==============================================================================
Public Sub LoadFromWorksheet(ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim y As Long

    Call Clear

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For y = MASTER_START_ROW To lastRow
        If Trim$(CStr(ws.Cells(y, 1).value)) <> "" Then
            Call AppendLineFromRow(ws, y)
        End If
    Next

    'マスタのチェック
    Call ValidateMasterLight
    
End Sub

'==============================================================================
'   行追加（ワークシート1行 → クラス）
'==============================================================================
Private Sub AppendLineFromRow(ByVal ws As Worksheet, ByVal y As Long)

    Dim line As FinancialStatementMasterLine
    Set line = New FinancialStatementMasterLine

    Call line.Initialize( _
        CStr(ws.Cells(y, 1).value), _
        CStr(ws.Cells(y, 2).value), _
        CStr(ws.Cells(y, 3).value), _
        CStr(ws.Cells(y, 4).value), _
        CStr(ws.Cells(y, 5).value), _
        ws.Cells(y, 5).HorizontalAlignment, _
        CStr(ws.Cells(y, 6).value), _
        CStr(ws.Cells(y, 7).value), _
        CStr(ws.Cells(y, 8).value), _
        ParseBoolean(ws.Cells(y, 9).value, True))
    
    Call AddLine(line)

End Sub

'==============================================================================
'   行登録
'------------------------------------------------------------------------------
'   Contents:
'       財務諸表マスタの1行を内部コレクションへ登録する
'
'       ■設計要件
'       ・lineIDは重複可能（例：PL_BIZ）
'       ・ただし、lineID＋accountingCategoryの組み合わせは一意
'       ・同一キーが存在した場合は設計ミスとみなしエラーとする
'
'       ■目的
'       ・会計区分別表示（公1、公2など）を実現する
'       ・勘定科目マスタのIDと財務諸表表示を分離する
'
'==============================================================================
Public Sub AddLine(ByVal line As FinancialStatementMasterLine)

    Dim key As String
    key = BuildKey(line.lineID, line.accountingCategory)

    If m_LineMap.Exists(key) Then
        Err.Raise vbObjectError + 6202, _
            "FinancialStatementMaster.AddLine", _
            "Duplicate Line: " & key
    End If

    m_LineMap.Add key, line
    m_Lines.Add line

End Sub

'==============================================================================
'   数値変換
'==============================================================================
Private Function ParseLong(ByVal v As Variant, ByVal defaultValue As Long) As Long
    If IsNumeric(v) Then
        ParseLong = CLng(v)
    Else
        ParseLong = defaultValue
    End If
End Function

'==============================================================================
'   Boolean変換
'==============================================================================
Private Function ParseBoolean(ByVal v As Variant, ByVal defaultValue As Boolean) As Boolean

    Dim s As String
    s = UCase$(Trim$(CStr(v)))

    If s = "" Then
        ParseBoolean = defaultValue
    ElseIf s = "1" Or s = "TRUE" Or s = "Y" Then
        ParseBoolean = True
    ElseIf s = "0" Or s = "FALSE" Or s = "N" Then
        ParseBoolean = False
    Else
        ParseBoolean = defaultValue
    End If

End Function

'==============================================================================
'       マスタ簡易検証
'------------------------------------------------------------------------------
'   Contents:
'       読み込んだ財務諸表マスタの整合性をチェックする
'
'==============================================================================
Private Sub ValidateMasterLight()

    Call ValidateAggregateToExists
    Call ValidateFormulaReferences
    Call ValidateCircularReference
End Sub

'==============================================================================
'       AggregateTo 存在チェック
'------------------------------------------------------------------------------
'   Contents:
'       各行の AggregateTo（集計先）が正しく定義されているかを検証する
'
'       ■チェック内容
'       ・AggregateTo が指定されている場合
'           → 同一category または "*" に該当する親行が存在すること
'
'       ■エラー条件
'           ・指定された親行がマスタに存在しない場合
'
'       ■目的
'           ・Rollup処理時の実行時エラー防止
'           ・マスタ定義ミスの早期検出
'
'==============================================================================
Private Sub ValidateAggregateToExists()

    Dim line As FinancialStatementMasterLine

    For Each line In m_Lines

        If Trim$(line.AggregateTo) <> "" Then

            If Not Me.Exists(line.AggregateTo, line.accountingCategory) _
               And Not Me.Exists(line.AggregateTo, "*") Then

                Err.Raise vbObjectError + 9002, _
                    "FinancialStatementMaster.ValidateAggregateToExists", _
                    "AggregateTo未定義: " & line.lineID & " → " & line.AggregateTo
            End If

        End If

    Next

End Sub

'==============================================================================
'       Formula参照チェック
'------------------------------------------------------------------------------
'   Contents:
'       FORMULA / CALC 行に定義された数式内の参照が、
'       財務諸表マスタ上に正しく存在するかを検証する
'
'       ■対応形式
'           [FS_ID]
'           [FS_ID|category]
'
'       ■チェック内容
'           ・lineID が存在するか
'           ・category 指定時は該当組み合わせが存在するか
'
'       ■エラー条件
'           ・参照先のFS_IDが存在しない
'           ・category付き参照で該当行が存在しない
'
'       ■目的
'           ・ResolveFormula 実行時のエラー防止
'           ・数式定義ミスの事前検出
'
'==============================================================================
Private Sub ValidateFormulaReferences()

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    With re
        .Pattern = "\[([^\]]+)\]"
        .Global = True
    End With

    Dim line As FinancialStatementMasterLine
    Dim matches As Object
    Dim m As Object

    For Each line In m_Lines

        If line.lineType = "FORMULA" Or line.lineType = "CALC" Then

            Set matches = re.Execute(line.formula)

            For Each m In matches

                Dim token As String
                token = m.SubMatches(0)

                Dim lineID As String
                Dim category As String

                If InStr(token, "|") > 0 Then
                    lineID = UCase$(Split(token, "|")(0))
                    category = UCase$(Split(token, "|")(1))
                Else
                    lineID = UCase$(token)
                    category = "*"
                End If

                If category = "*" Then
                    If Not Me.Exists(lineID) Then
                        Err.Raise vbObjectError + 9003, _
                            "FinancialStatementMaster.ValidateFormulaReferences", _
                            "Formula参照不正: [" & token & "]"
                    End If
                Else
                    If Not Me.Exists(lineID, category) Then
                        Err.Raise vbObjectError + 9004, _
                            "FinancialStatementMaster.ValidateFormulaReferences", _
                            "Formula参照不正: [" & token & "]"
                    End If
                End If

            Next

        End If

    Next

End Sub

'==============================================================================
'       循環参照チェック（AggregateTo）
'------------------------------------------------------------------------------
'   Contents:
'       AggregateTo の親子関係に循環が無いかを検証する
'
'       ■チェック例（NG）
'           A → B
'           B → C
'           C → A
'
'       ■仕様
'           ・DFS（深さ優先探索）で検出
'           ・同一スタック上に同じノードが現れたら循環
'
'==============================================================================
Private Sub ValidateCircularReference()

    Dim visited As Object
    Set visited = CreateObject("Scripting.Dictionary")

    Dim stack As Object
    Set stack = CreateObject("Scripting.Dictionary")

    Dim line As FinancialStatementMasterLine

    For Each line In m_Lines
        Call ValidateCircularReference_DFS(line, visited, stack)
    Next

End Sub

'==============================================================================
'       DFSによる循環検出
'------------------------------------------------------------------------------
'   Contents:
'       再帰的に親を辿り、循環を検出する
'
'   Input:
'       line        現在の行
'       visited     訪問済みノード
'       stack       現在の探索スタック
'
'==============================================================================
Private Sub ValidateCircularReference_DFS( _
    ByVal line As FinancialStatementMasterLine, _
    ByVal visited As Object, _
    ByVal stack As Object)

    Dim key As String
    key = BuildKey(line.lineID, line.accountingCategory)

    '------------------------------------------
    ' スタックに存在 → 循環
    '------------------------------------------
    If stack.Exists(key) Then
        Err.Raise vbObjectError + 9005, _
            "FinancialStatementMaster.ValidateCircularReference", _
            "循環参照が検出されました: " & key
    End If

    '------------------------------------------
    ' 既訪問ならスキップ
    '------------------------------------------
    If visited.Exists(key) Then Exit Sub

    visited.Add key, True
    stack.Add key, True

    '------------------------------------------
    ' 親へ遷移
    '------------------------------------------
    If Trim$(line.AggregateTo) <> "" Then

        Dim parent As FinancialStatementMasterLine

        'GetLineはエラーを投げる仕様なのでここではそのまま使う
        If Me.Exists(line.AggregateTo, line.accountingCategory) Then
            Set parent = Me.GetLine(line.AggregateTo, line.accountingCategory)
        ElseIf Me.Exists(line.AggregateTo, "*") Then
            Set parent = Me.GetLine(line.AggregateTo, "*")
        Else
            'ここは別チェックで拾っているが保険
            Err.Raise vbObjectError + 9006, _
                "FinancialStatementMaster.ValidateCircularReference", _
                "親が存在しません: " & line.AggregateTo
        End If

        Call ValidateCircularReference_DFS(parent, visited, stack)

    End If

    '------------------------------------------
    ' スタックから除外
    '------------------------------------------
    stack.Remove key

End Sub
