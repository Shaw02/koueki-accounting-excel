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

Public Function Exists(ByVal lineID As String) As Boolean
    Exists = m_LineMap.Exists(UCase$(Trim$(lineID)))
End Function

Public Function GetLine(ByVal lineID As String) As FinancialStatementMasterLine
    Dim key As String
    key = UCase$(Trim$(lineID))

    If Not m_LineMap.Exists(key) Then
        Err.Raise vbObjectError + 6201, "FinancialStatementMaster.GetLine", "LineID not found: " & lineID
    End If

    Set GetLine = m_LineMap(key)
End Function

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

Public Sub AddLine(ByVal line As FinancialStatementMasterLine)

    Dim key As String
    key = UCase$(line.lineID)

    If m_LineMap.Exists(key) Then
        Err.Raise vbObjectError + 6202, "FinancialStatementMaster.AddLine", "Duplicate LineID: " & line.lineID
    End If

    m_LineMap.Add key, line
    m_Lines.Add line

End Sub

Private Function ParseLong(ByVal v As Variant, ByVal defaultValue As Long) As Long
    If IsNumeric(v) Then
        ParseLong = CLng(v)
    Else
        ParseLong = defaultValue
    End If
End Function

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


