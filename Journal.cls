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

'==============================================================================
'   定数定義
'==============================================================================

'--------------------------------------
'仕訳帳の列
Private Const COL_DATE          As Long = 1
' --- 借方 ---
Private Const COL_DR_ACC_NO_I   As Long = 2
Private Const COL_DR_ACC_NO_S   As Long = 3
Private Const COL_DR_ACC_NO_S2  As Long = 4
Private Const COL_DR_SUB_NO_I   As Long = 5
Private Const COL_DR_SUB_NO_S   As Long = 6
Private Const COL_DR_CLASS      As Long = 7   ' 借方：一般/指定区分
Private Const COL_DR_MONEY      As Long = 8
' --- 貸方 ---
Private Const COL_CR_ACC_NO_I   As Long = 9
Private Const COL_CR_ACC_NO_S   As Long = 10
Private Const COL_CR_ACC_NO_S2  As Long = 11
Private Const COL_CR_SUB_NO_I   As Long = 12
Private Const COL_CR_SUB_NO_S   As Long = 13
Private Const COL_CR_CLASS      As Long = 14  ' 貸方：一般/指定区分
Private Const COL_CR_MONEY      As Long = 15
' --- 摘要 ---
Private Const COL_SUMMARY       As Long = 16

'仕訳帳の読み込みＹ座標
Private Const styJournal = 8            '仕訳帳の開始位置

'==============================================================================
'   メンバー変数
'==============================================================================
Private m_record() As AccountingEntry   '仕訳帳
Private m_size As Long                  '仕訳帳のデータ行数

Private y As Long                       '現在読み込み中のy座標
Private m_Err As Boolean                'エラー発生有無
Private m_strErr As String              'エラー文字列（蓄積する）

'==============================================================================
'   クラスの初期化
'==============================================================================
Private Sub Class_Initialize()
    m_size = 0
    m_Err = False
    m_strErr = ""
    Call readJournal
End Sub

'==============================================================================
'   クラスの解放
'==============================================================================
Private Sub Class_Terminate()
    Dim i As Long
    For i = 0 To m_size - 1
        Set m_record(i) = Nothing
    Next
End Sub

'==============================================================================
'   サイズの取得
'==============================================================================
Public Function getSize() As Long
    getSize = m_size
End Function

'==============================================================================
'   ｙ行目の仕訳帳の内容を取得
'==============================================================================
Public Function getRecord(y As Long) As AccountingEntry
    Set getRecord = m_record(y)
End Function

Public Property Get Items() As Variant
    Items = m_record
End Property

'==============================================================================
'   仕訳帳の読み込み
'==============================================================================
Public Sub readJournal()

    On Error GoTo ErrHandler
    
    '----------------------------------
    '   変数宣言
    '----------------------------------
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    
    Dim iTotalCrMoney As Currency
    Dim iTotalDrMoney As Currency
    
    '----------------------------------
    '   初期化
    '----------------------------------
    Set ws = ThisWorkbook.Worksheets("①仕訳帳")
    lastRow = ws.Cells(ws.Rows.Count, COL_DATE).End(xlUp).Row
    
    iTotalCrMoney = 0
    iTotalDrMoney = 0
    
    '--------------------------------------
    ' 最終行取得（日付列基準）
    '--------------------------------------
    ' 開始行より上しかデータがない場合は終了
    If lastRow < styJournal Then
        m_size = 0
        Exit Sub
    End If
    
    '--------------------------------------
    ' 配列を一括確保（0ベース）
    '--------------------------------------
    m_size = lastRow - styJournal + 1
    ReDim m_record(0 To m_size - 1)
    
    '--------------------------------------
    ' 仕訳帳読み込み
    '--------------------------------------
    i = 0
    For y = styJournal To lastRow
        
        ' AccountingEntry 生成
        Set m_record(i) = New AccountingEntry
        
        ' 1レコード読み込み（行番号を渡す）
        Call m_record(i).Initialize( _
            ws.Cells(y, COL_DATE), _
            ws.Cells(y, COL_DR_ACC_NO_I), _
            ws.Cells(y, COL_DR_ACC_NO_S), _
            ws.Cells(y, COL_DR_ACC_NO_S2), _
            ws.Cells(y, COL_DR_SUB_NO_I), _
            ws.Cells(y, COL_DR_SUB_NO_S), _
            ws.Cells(y, COL_DR_CLASS), _
            ws.Cells(y, COL_DR_MONEY), _
            ws.Cells(y, COL_CR_ACC_NO_I), _
            ws.Cells(y, COL_CR_ACC_NO_S), _
            ws.Cells(y, COL_CR_ACC_NO_S2), _
            ws.Cells(y, COL_CR_SUB_NO_I), _
            ws.Cells(y, COL_CR_SUB_NO_S), _
            ws.Cells(y, COL_CR_CLASS), _
            ws.Cells(y, COL_CR_MONEY), _
            ws.Cells(y, COL_SUMMARY))
        
        ' 日付が空なら終了（途中空白対策）
        If IsNull(m_record(i).entryDate) Or IsEmpty(m_record(i).entryDate) Then
            m_size = i   ' 実際の件数に調整
            ReDim Preserve m_record(0 To m_size - 1)
            Exit For
        End If
        
        iTotalDrMoney = iTotalDrMoney + m_record(i).Debit.amount
        iTotalCrMoney = iTotalCrMoney + m_record(i).Credit.amount
        
        i = i + 1
    Next y

    '------------------------------
    ' 貸借一致 最終チェック
    '------------------------------
    If (iTotalCrMoney <> iTotalDrMoney) Then
        Err.Raise vbObjectError + 5000, "Journal:readJournal", "仕訳帳：借方と借方の合計金額が一致しません"
    End If
    
    If m_Err = True Then
        MsgBox "《仕訳帳入力エラー》" & vbCrLf & "仕訳帳の以下の行を修正してください：" & vbCrLf & _
           m_strErr, _
           vbExclamation, "仕訳帳入力エラー"
        End
    End If
    
    Exit Sub
    
ErrHandler:

    If m_Err = True Then
        m_strErr = m_strErr & vbCrLf
    Else
        m_Err = True
    End If

    m_strErr = m_strErr & str(y) & "行目：" & Err.Description

    Resume Next

End Sub





