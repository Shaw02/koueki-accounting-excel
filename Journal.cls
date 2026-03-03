VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Journal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'   仕訳帳
'==============================================================================
Option Explicit

'==============================
' 仕訳帳（Sheet2）列定義
'==============================

Private Const COL_DATE          As Long = 1

'仕訳帳の読み込みＹ座標
Private Const styJournal = 8         '仕訳帳の開始位置

Private record() As AccountingEntry
Private size As Long

Private fErr As Boolean             'エラー発生状況

'==============================================================================
'   クラスの初期化
'==============================================================================
Private Sub Class_Initialize()
    size = 0
    fErr = False
    'Call readJournal
End Sub

'==============================================================================
'   クラスの解放
'==============================================================================
Private Sub Class_Terminate()
    Dim i As Long
    For i = 0 To size - 1
        Set record(i) = Nothing
    Next
End Sub

'==============================================================================
'   サイズの取得
'==============================================================================
Public Function getSize() As Long
    getSize = size
End Function

'==============================================================================
'   ｙ行目の仕訳帳の内容を取得
'==============================================================================
Public Function getRecord(y As Long) As AccountingEntry
    Set getRecord = record(y)
End Function

Public Property Get Items() As Variant
    Items = record
End Property

'==============================================================================
'   仕訳帳の読み込み
'==============================================================================
Public Sub readJournal()

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim y As Long
    Dim i As Long
    
    Dim iTotalCrMoney As Currency
    Dim iTotalDrMoney As Currency
    
    Set ws = ThisWorkbook.Worksheets("①仕訳帳")
    
    '--------------------------------------
    ' 最終行取得（日付列基準）
    '--------------------------------------
    lastRow = ws.Cells(ws.Rows.count, COL_DATE).End(xlUp).Row
    
    ' 開始行より上しかデータがない場合は終了
    If lastRow < styJournal Then
        size = 0
        Exit Sub
    End If
    
    '--------------------------------------
    ' 配列を一括確保（0ベース）
    '--------------------------------------
    size = lastRow - styJournal + 1
    ReDim record(0 To size - 1)
    
    iTotalCrMoney = 0
    iTotalDrMoney = 0
    
    '--------------------------------------
    ' 仕訳帳読み込み
    '--------------------------------------
    i = 0
    For y = styJournal To lastRow
        
        ' AccountingEntry 生成
        Set record(i) = New AccountingEntry
        
        ' 1レコード読み込み（行番号を渡す）
        record(i).readJournal (y)
        
        ' 日付が空なら終了（途中空白対策）
        If IsNull(record(i).entryDate) Or IsEmpty(record(i).entryDate) Then
            size = i   ' 実際の件数に調整
            ReDim Preserve record(0 To size - 1)
            Exit For
        End If
        
        iTotalDrMoney = iTotalDrMoney + record(i).Debit.Amount
        iTotalCrMoney = iTotalCrMoney + record(i).Credit.Amount
        
        i = i + 1
    Next y

    '------------------------------
    ' 貸借一致 最終チェック
    '------------------------------
    If (iTotalCrMoney <> iTotalDrMoney) Then
        Err.Raise vbObjectError + 1100, "Journal:readJournal", "仕訳帳：借方と借方の合計金額が一致しません"
    End If
    
    If fErr = True Then
        MsgBox "処理を停止します。"
        End
    End If
    
    Exit Sub
    
ErrHandler:

    MsgBox "仕訳帳入力エラー：" & vbCrLf & _
           Err.Description, _
           vbExclamation, "仕訳帳入力エラー"

    fErr = True

    Resume Next

End Sub





