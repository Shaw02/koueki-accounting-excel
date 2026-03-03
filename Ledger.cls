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

'=======================================================
'       総勘定元帳・今期実績のクリア
'-------------------------------------------------------
Sub SheetClear()
'
' GeneralAccountClear Macro
'
' Keyboard Shortcut: Ctrl+g
'
    '今期実績
    Sheet11.Activate
    Sheet11.Range("A7", Sheet11.Range("A7").SpecialCells(xlCellTypeLastCell)).Select
    Selection.EntireRow.Delete
    Sheet11.Range("A7").Select
    
    '総勘定元帳
    Sheet3.Activate
    Sheet3.Range("A5", Sheet3.Range("A5").SpecialCells(xlCellTypeLastCell)).Select
    Selection.EntireRow.Delete
    Sheet3.Range("A5").Select

    '補助元帳
    Sheet4.Activate
    Sheet4.Range("A5", Sheet4.Range("A5").SpecialCells(xlCellTypeLastCell)).Select
    Selection.EntireRow.Delete
    Sheet4.Range("A5").Select

End Sub

