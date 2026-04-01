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
        ParseLong(ws.Cells(y, 7).value, 1), _
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

