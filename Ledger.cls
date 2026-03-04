VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'   総勘定元帳・補助元帳
'==============================================================================
Option Explicit

Private Const ROW_CR_START As Long = 7
Private Const ROW_GL_START As Long = 5
Private Const ROW_SL_START As Long = 5




'==============================================================================
'   クラスの初期化
'==============================================================================
Private Sub Class_Initialize()
    Call SheetClear          '総勘定元帳・今期実績のクリア
End Sub

'==============================================================================
'   クラスの解放
'==============================================================================
Private Sub Class_Terminate()

End Sub

'==============================================================================
'       総勘定元帳・今期実績のクリア
'==============================================================================
Sub SheetClear()
'
    '----------------------------------
    '   変数宣言
    '----------------------------------
    Dim wsCR As Worksheet
    Dim wsGL As Worksheet
    Dim wsSL As Worksheet
    
    '----------------------------------
    '   初期化
    '----------------------------------
    Set wsCR = ThisWorkbook.Worksheets("今期実績")
    Set wsGL = ThisWorkbook.Worksheets("②総勘定元帳")
    Set wsSL = ThisWorkbook.Worksheets("③補助元帳")
    
    '----------------------------------
    '   シートクリア
    '----------------------------------
    Call ClearSheet(wsCR, ROW_CR_START)
    Call ClearSheet(wsGL, ROW_GL_START)
    Call ClearSheet(wsSL, ROW_SL_START)

End Sub

'==============================================================================
'      シートのクリア
'------------------------------------------------------------------------------
'   Contents:
'       指定のシートをクリアします。
'   Input:
'       ws  クリアするシート
'   Output:
'       startRow    開始行
'==============================================================================
Private Sub ClearSheet(ws As Worksheet, startRow As Long)

    '----------------------------------
    '   変数宣言
    '----------------------------------
    Dim lastRow As Long
    
    '----------------------------------
    '   シートクリア
    '----------------------------------
    With ws
        lastRow = .Cells(.Rows.count, "A").End(xlUp).Row
        If lastRow >= startRow Then
            .Rows(startRow & ":" & lastRow).Delete
        End If
    End With

End Sub
