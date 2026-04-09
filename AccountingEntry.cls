VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccountingEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'   仕訳の１レコード
'==============================================================================
Option Explicit

'==============================================================================
'   メンバー変数
'==============================================================================
Private m_EntryDate         As Variant      ' 日付
Private m_Debits            As Collection   ' 借方
Private m_Credits           As Collection   ' 貸方
Private m_Summary           As String       ' 摘要
Private m_StartRow          As Long         ' 仕訳帳の行

Private m_TotalDebit        As Currency
Private m_TotalCredit       As Currency

'==============================================================================
'   Property
'==============================================================================
Public Property Get entryDate() As Variant
    entryDate = m_EntryDate
End Property

Public Property Get Debits() As Collection
    Set Debits = m_Debits
End Property

Public Property Get Credits() As Collection
    Set Credits = m_Credits
End Property

Public Property Get summary() As String
    summary = m_Summary
End Property

Public Property Get StartRow() As Long
    StartRow = m_StartRow
End Property

'==============================================================================
'   クラスの初期化
'==============================================================================
Private Sub Class_Initialize()
    Set m_Debits = New Collection
    Set m_Credits = New Collection
    
    m_TotalDebit = 0
    m_TotalCredit = 0
End Sub

'==============================================================================
'   クラスの解放
'==============================================================================
Private Sub Class_Terminate()

End Sub

'==============================================================================
'   仕訳ヘッダー設定
'------------------------------------------------------------------------------
'   Input:
'       d   日付
'       s   摘要
'       r   仕訳行
'==============================================================================
Public Sub SetHeader(ByVal d As Variant, ByVal s As String, ByVal r As Long)

    m_EntryDate = d     '日付
    m_Summary = s       '摘要
    m_StartRow = r      '仕訳帳の行

End Sub

'==============================================================================
'   借方追加
'==============================================================================
Public Sub AddDebit(ByVal side As entrySide)

    If side Is Nothing Then
        Err.Raise vbObjectError + 4101, _
                  "AccountingEntry.AddDebit", _
                  "借方明細が Nothing です。"
    End If
    
    If side.amount < 0 Then
        Err.Raise vbObjectError + 4102, _
                  "AccountingEntry.AddDebit", _
                  "借方金額が負数です。"
    End If
    
    m_TotalDebit = m_TotalDebit + side.amount
    m_Debits.Add side

End Sub

'==============================================================================
'   貸方追加
'==============================================================================
Public Sub AddCredit(ByVal side As entrySide)

    If side Is Nothing Then
        Err.Raise vbObjectError + 4103, _
                  "AccountingEntry.AddCredit", _
                  "貸方明細が Nothing です。"
    End If
    
    If side.amount < 0 Then
        Err.Raise vbObjectError + 4104, _
                  "AccountingEntry.AddCredit", _
                  "貸方金額が負数です。"
    End If
    
    m_TotalCredit = m_TotalCredit + side.amount
    m_Credits.Add side

End Sub

'==============================================================================
'   借方合計を取得
'==============================================================================
Public Function TotalDebit() As Currency

    TotalDebit = m_TotalDebit

End Function

'==============================================================================
'   貸方合計を取得
'==============================================================================
Public Function TotalCredit() As Currency

    TotalCredit = m_TotalCredit

End Function
