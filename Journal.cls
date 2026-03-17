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
Private Const COL_DR_SUB_NO_C   As Long = 6
Private Const COL_DR_SUB_NO_S   As Long = 7
Private Const COL_DR_CLASS      As Long = 8   ' 借方：一般/指定区分
Private Const COL_DR_MONEY      As Long = 9
' --- 貸方 ---
Private Const COL_CR_ACC_NO_I   As Long = 10
Private Const COL_CR_ACC_NO_S   As Long = 11
Private Const COL_CR_ACC_NO_S2  As Long = 12
Private Const COL_CR_SUB_NO_I   As Long = 13
Private Const COL_CR_SUB_NO_C   As Long = 14
Private Const COL_CR_SUB_NO_S   As Long = 15
Private Const COL_CR_CLASS      As Long = 16  ' 貸方：一般/指定区分
Private Const COL_CR_MONEY      As Long = 17
' --- 摘要 ---
Private Const COL_SUMMARY       As Long = 18

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
Public Sub readJournal(master As AccountMaster)

    On Error GoTo ErrHandler
    
    '----------------------------------
    '   変数宣言
    '----------------------------------
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    
    Dim d_account As Long
    Dim c_account As Long
    
    Dim iTotalCrMoney As Currency
    Dim iTotalDrMoney As Currency
    
    Dim currentEntry As AccountingEntry     '現在の仕訳
    
    '----------------------------------
    '   初期化
    '----------------------------------
    Set ws = ThisWorkbook.Worksheets("①仕訳帳")
    lastRow = Application.WorksheetFunction.Max( _
        ws.Cells(ws.Rows.Count, COL_DATE).End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, COL_DR_MONEY).End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, COL_CR_MONEY).End(xlUp).Row)
    
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
    Set currentEntry = Nothing
    
    For y = styJournal To lastRow
        
        Dim vDate As Variant
        vDate = ws.Cells(y, COL_DATE)
        
        d_account = ws.Cells(y, COL_DR_ACC_NO_I)
        c_account = ws.Cells(y, COL_CR_ACC_NO_I)
        
        '--------------------------------------
        ' 空行スキップ
        '--------------------------------------
        If IsEmpty(vDate) And d_account = 0 And c_account = 0 Then
            GoTo ContinueLoop
        End If
        
        '--------------------------------------
        ' 新しい仕訳
        '--------------------------------------
        If Not IsEmpty(vDate) Then
            
            ' 前の仕訳を確定
            If Not currentEntry Is Nothing Then
                
                If currentEntry.TotalDebit <> currentEntry.TotalCredit Then
                    Err.Raise vbObjectError + 5000, , _
                        "貸借不一致"
                End If
                
                Set m_record(i) = currentEntry
                
                iTotalDrMoney = iTotalDrMoney + currentEntry.TotalDebit
                iTotalCrMoney = iTotalCrMoney + currentEntry.TotalCredit
                
                i = i + 1
            End If
            
            ' 新規作成
            Set currentEntry = New AccountingEntry
            
            Call currentEntry.SetHeader( _
                vDate, _
                ws.Cells(y, COL_SUMMARY) _
            )
            
        Else
            If currentEntry Is Nothing Then
                Err.Raise vbObjectError + 4001, , "日付なし開始エラー"
            End If
        End If
        
        '--------------------------------------
        ' 借方チェック
        '--------------------------------------
        If d_account <> 0 Then
            If Not master.Exists(d_account) Then
                Err.Raise vbObjectError + 3001, _
                    "Journal.class", _
                    "借方:勘定科目コード" & d_account & "は入力できません。"
            End If
            
            Dim d As entrySide
            Set d = New entrySide
            
            Call d.Initialize( _
                master.GetAccount(d_account), _
                ws.Cells(y, COL_DR_SUB_NO_I), _
                ws.Cells(y, COL_DR_SUB_NO_C), _
                ws.Cells(y, COL_DR_SUB_NO_S), _
                ws.Cells(y, COL_DR_CLASS), _
                ws.Cells(y, COL_DR_MONEY), _
                "借方" _
            )
            
            currentEntry.AddDebit d
        End If
        
        '--------------------------------------
        ' 貸方チェック
        '--------------------------------------
        If c_account <> 0 Then
            If Not master.Exists(c_account) Then
                Err.Raise vbObjectError + 3001, _
                    "Journal.class", _
                    "貸方:勘定科目コード" & c_account & "は入力できません。"
            End If
            
            Dim c As entrySide
            Set c = New entrySide
            
            Call c.Initialize( _
                master.GetAccount(c_account), _
                ws.Cells(y, COL_CR_SUB_NO_I), _
                ws.Cells(y, COL_CR_SUB_NO_C), _
                ws.Cells(y, COL_CR_SUB_NO_S), _
                ws.Cells(y, COL_CR_CLASS), _
                ws.Cells(y, COL_CR_MONEY), _
                "貸方" _
            )
            
            currentEntry.AddCredit c
        End If
        
ContinueLoop:
    Next y
    
    '--------------------------------------
    ' 最終仕訳
    '--------------------------------------
    If Not currentEntry Is Nothing Then
        
        If currentEntry.TotalDebit <> currentEntry.TotalCredit Then
            Err.Raise vbObjectError + 5000, , "貸借不一致"
        End If
        
        Set m_record(i) = currentEntry
        
        iTotalDrMoney = iTotalDrMoney + currentEntry.TotalDebit
        iTotalCrMoney = iTotalCrMoney + currentEntry.TotalCredit
        
        i = i + 1
    End If
    
    '--------------------------------------
    ' サイズ調整
    '--------------------------------------
    m_size = i
    ReDim Preserve m_record(0 To m_size - 1)
    
    '--------------------------------------
    ' 最終貸借チェック
    '--------------------------------------
    If iTotalCrMoney <> iTotalDrMoney Then
        Err.Raise vbObjectError + 5000, _
            "Journal:readJournal", _
            "仕訳帳：借方と貸方の合計金額が一致しません"
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

    m_strErr = m_strErr & CStr(y) & "行目：" & Err.Description

    Resume Next

End Sub

'==============================================================================
'   仕訳帳チェック
'==============================================================================
Public Sub ValidateEntries(ByVal master As AccountMaster)

    Dim it As Variant

    For Each it In m_record

        '借方
        Call it.Debit.CheckEntry(master)
        Call it.Credit.CheckEntry(master)

    Next

End Sub
