VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LedgerEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'   元帳エンジン
'------------------------------------------------------------------------------
'   Contents:
'       仕訳帳から
'       ・総勘定元帳
'       ・補助元帳
'       を生成する
'
'       LedgerAccount を管理する
'
'==============================================================================
Option Explicit

'総勘定元帳の出力Ｙ座標
Private Const styGeneralLedger = 5          '総勘定元帳の開始位置

Private m_GeneralLedger As Object           '総勘定元帳
Private m_SubLedger As Object               '補助元帳

'==============================================================================
'   総勘定元帳取得
'==============================================================================
Public Property Get GeneralLedger() As Object
    Set GeneralLedger = m_GeneralLedger
End Property

'==============================================================================
'   補助元帳取得
'==============================================================================
Public Property Get SubLedger() As Object
    Set SubLedger = m_SubLedger
End Property

'==============================================================================
'   初期化
'==============================================================================
Private Sub Class_Initialize()

    Set m_GeneralLedger = CreateObject("Scripting.Dictionary")
    Set m_SubLedger = CreateObject("Scripting.Dictionary")
    
End Sub

'==============================================================================
'   総勘定元帳の科目一覧取得
'==============================================================================
Public Function GetAccounts() As Collection

    Dim col As New Collection
    Dim key As Variant

    For Each key In m_GeneralLedger
        col.Add m_GeneralLedger(key)
    Next key

    Set GetAccounts = col

End Function

'==============================================================================
'   指定キーの元帳を取得    無ければ新規作成
'------------------------------------------------------------------------------
'   Input
'       dict            総勘定元帳 or 補助元帳の dict
'       keyStr          キー
'       entry           行
'
'   Output
'       LedgerAccount   元帳のオブジェクト
'
'==============================================================================
Private Function GetOrCreateLedger( _
    ByRef dict As Object, _
    ByVal keyStr As String, _
    ByVal entry As entrySide) As LedgerAccount

    If dict.Exists(keyStr) Then
        
        Set GetOrCreateLedger = dict(keyStr)
        
    Else
        
        Dim acc As LedgerAccount
        Set acc = New LedgerAccount
        
        Call acc.Initialize( _
            entry.account, _
            entry.SubAccountCode, _
            entry.SubAccountCate, _
            entry.SubAccountName, _
            entry.Class)
        
        dict.Add keyStr, acc
        
        Set GetOrCreateLedger = acc
        
    End If

End Function

'==============================================================================
'   前期実績1行を元帳へ転記
'------------------------------------------------------------------------------
'   Input:
'       entry   entrySide
'
'==============================================================================
Public Sub AddOpening(tbLine As TrialBalanceLine, isGeneral As Boolean)
    
    Dim key As ledgerKey
    Set key = tbLine.key
    
    Dim acc As LedgerAccount
    
    '試算表１行をこのクラスに変換する
    Dim entry As entrySide
    Set entry = New entrySide
    
    Call entry.Initialize( _
        tbLine.account, _
        tbLine.SubAccountCode, _
        tbLine.SubAccountCate, _
        tbLine.SubAccountName, _
        tbLine.NetAssetType, _
        tbLine.EndingBalance, _
        "借方", 0, 0)

    If isGeneral Then
        Set acc = GetOrCreateLedger(m_GeneralLedger, key.key, entry)
        acc.AddOpeningBalance (entry.Amount)
    Else
        Set acc = GetOrCreateLedger(m_SubLedger, key.key, entry)
        acc.AddOpeningBalance (entry.Amount)
    End If

End Sub

'==============================================================================
'   仕訳帳１行を、総勘定元帳＆補助元帳（両方）に転記
'------------------------------------------------------------------------------
'   Input:
'       entry   AccountingEntry
'
'==============================================================================
Public Sub AddJournal(entry As AccountingEntry)
    
    Dim d As entrySide
    Dim c As entrySide

    Dim counterAcc As account
    Dim isMulti As Boolean

    '--------------------------------------
    ' 諸口チェック
    '--------------------------------------
    isMulti = False
    If entry.Credits.Count = 1 Then
        Set counterAcc = entry.Credits(1).account
    Else
        Set counterAcc = Nothing
        If entry.Credits.Count > 1 Then
            isMulti = True
        End If
    End If
    
    '--------------------------------------
    ' 借方ループ
    '--------------------------------------
    For Each d In entry.Debits
    
        Dim accD As LedgerAccount
        Dim keyD As ledgerKey
        Set keyD = New ledgerKey
        
        '==============================
        ' 総勘定元帳（借方）
        '==============================
        Call keyD.Initialize(d.AccountCode, 0, 0)
        
        Set accD = GetOrCreateLedger(m_GeneralLedger, keyD.key, d)
        
        Call accD.AddLine( _
            d.Row, _
            d.ParentRow, _
            entry.entryDate, _
            counterAcc, _
            entry.summary, _
            d.Amount, _
            0, _
            isMulti)
        
        '==============================
        ' 補助元帳（借方）
        '==============================
        Call keyD.Initialize( _
            d.AccountCode, _
            d.SubAccountCode, _
            d.Class)
        
        Set accD = GetOrCreateLedger(m_SubLedger, keyD.key, d)
        
        Call accD.AddLine( _
            d.Row, _
            d.ParentRow, _
            entry.entryDate, _
            counterAcc, _
            entry.summary, _
            d.Amount, _
            0, _
            isMulti)
    Next
        
    '--------------------------------------
    ' 諸口チェック
    '--------------------------------------
    isMulti = False
    If entry.Debits.Count = 1 Then
        Set counterAcc = entry.Debits(1).account
    Else
        Set counterAcc = Nothing
        If entry.Debits.Count > 1 Then
            isMulti = True
        End If
    End If
    '--------------------------------------
    ' 貸方ループ
    '--------------------------------------
    For Each c In entry.Credits
        
        Dim accC As LedgerAccount
        Dim keyC As ledgerKey
        Set keyC = New ledgerKey
        
        '==============================
        ' 総勘定元帳（貸方）
        '==============================
        Call keyC.Initialize(c.AccountCode, 0, 0)
        
        Set accC = GetOrCreateLedger(m_GeneralLedger, keyC.key, c)
        
        Call accC.AddLine( _
            c.Row, _
            c.ParentRow, _
            entry.entryDate, _
            counterAcc, _
            entry.summary, _
            0, _
            c.Amount, _
            isMulti)
    
        '==============================
        ' 補助元帳（借方）
        '==============================
        Call keyC.Initialize( _
            c.AccountCode, _
            c.SubAccountCode, _
            c.Class)
        
        Set accC = GetOrCreateLedger(m_SubLedger, keyC.key, c)
        
        Call accC.AddLine( _
            c.Row, _
            c.ParentRow, _
            entry.entryDate, _
            counterAcc, _
            entry.summary, _
            0, _
            c.Amount, _
            isMulti)

    Next
End Sub

'==============================================================================
'   期末の計算
'------------------------------------------------------------------------------
Public Sub Final()
    
    Call SortLedgers

    Call OutputSheet

End Sub

'==============================================================================
'   元帳ソート
'==============================================================================
Public Sub SortLedgers()

    Set m_GeneralLedger = SortDictionaryByKey(m_GeneralLedger)
    Set m_SubLedger = SortDictionaryByKey(m_SubLedger)

End Sub

'==============================================================================
'   総勘定元帳 ＆ 補助元帳出力
'==============================================================================
Public Sub OutputSheet()

    Const MaxColumn As Long = 9

    '----------------------------------
    '   変数宣言
    '----------------------------------
    Dim wsGL As Worksheet
    Dim wsSL As Worksheet
    Dim key As Variant
    Dim acc As LedgerAccount
    Dim line As LedgerLine
    
    Dim y As Long
    Dim sty As Long
    
    '----------------------------------
    '   初期化
    '----------------------------------
    Set wsGL = ThisWorkbook.Worksheets("②総勘定元帳")
    Set wsSL = ThisWorkbook.Worksheets("③補助元帳")
    
    '----------------------------------
    '   シートクリア
    '----------------------------------
    Call ClearSheet(wsGL, styGeneralLedger)
    Call ClearSheet(wsSL, styGeneralLedger)

    '----------------------------------
    '   総勘定元帳
    '----------------------------------
    y = styGeneralLedger
    With wsGL
    
         .Columns("A").NumberFormat = "yyyy/mm/dd"
    
        '数値表示形式（小数点なし）
        .Range("E:F, H:H").NumberFormat = "#,##0;△#,##0;;@"
        .Range("E:F, H:H").HorizontalAlignment = xlRight
        .Range("G:G").HorizontalAlignment = xlCenter
        .Range("I:I").HorizontalAlignment = xlCenter
    
        For Each key In m_GeneralLedger
            Set acc = m_GeneralLedger.Item(key)
            If (acc.Lines.Count > 0) And ((acc.AccountCode < 30000) Or (acc.AccountCode >= 40000)) Then
                y = y + 1
                
                '--------------------------
                '勘定科目の出力
                .Range(.Cells(y, 1), .Cells(y, 8)).Merge
                If acc.MiddleAccount <> "" Then
                    .Cells(y, 1) = "科目：" & acc.MajorAccount & " － " & acc.MiddleAccount
                Else
                    .Cells(y, 1) = "科目：" & acc.MajorAccount
                End If
                .Cells(y, 1).Font.Underline = True
                .Cells(y, 1).Font.size = 16
                .Cells(y, 1).HorizontalAlignment = xlCenter
                y = y + 1
            
                '--------------------------
                '勘定科目コードの出力
                .Cells(y, 1) = "勘定科目コード：" & CStr(acc.AccountCode)
                y = y + 1
                
                '--------------------------
                '項目行
                .Cells(y, 1) = "日付"
                .Cells(y, 2) = "相手大科目"
                .Cells(y, 3) = "相手中科目"
                .Cells(y, 4) = "摘要"
                .Cells(y, 5) = "借方"
                .Cells(y, 6) = "貸方"
                .Cells(y, 7) = "貸／借"
                .Cells(y, 8) = "残高"
                .Cells(y, 9) = "仕訳帳行"
                
                With .Range(.Cells(y, 1), .Cells(y, MaxColumn))
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Interior.Color = rgbLavender
                    .Font.Bold = True
                    .Borders.LineStyle = xlContinuous
                    .Borders(xlEdgeTop).LineStyle = xlDouble
                    .Borders(xlEdgeBottom).Weight = xlMedium
                End With
                y = y + 1
                sty = y
                '--------------------------
                '行
                For Each line In acc.Lines
                    .Cells(y, 1).Resize(1, MaxColumn).value = Array( _
                        line.entryDate, _
                        line.counterMajorAccount, _
                        line.counterMiddleAccount, _
                        line.summary, _
                        IIf(line.debit = 0, "-", line.debit), _
                        IIf(line.credit = 0, "-", line.credit), _
                        line.DebitCreditLabel, _
                        line.BalanceAbs, _
                        IIf(line.Row = 0, "", "=HYPERLINK(""#'①仕訳帳'!A" & line.Row & """,""" & line.Row & """)"))
                    .Range(.Cells(y, 1), .Cells(y, MaxColumn)).Borders.LineStyle = xlContinuous
                    
                    'ゼブラ
                    If y Mod 2 = 0 Then
                        .Range(.Cells(y, 1), .Cells(y, MaxColumn)).Interior.Color = rgbAliceBlue
                    End If
                    
                    y = y + 1
                Next line
                '--------------------------
                '合計行
                .Cells(y, 1) = "合計"
                .Cells(y, 5).Formula = "=SUM(E" & sty & ":E" & y - 1 & ")"
                .Cells(y, 6).Formula = "=SUM(F" & sty & ":F" & y - 1 & ")"
        
                With .Range(.Cells(y, 1), .Cells(y, MaxColumn))
                    .Font.Bold = True
                    .Interior.Color = rgbLightGoldenrodYellow
                    .Borders(xlEdgeTop).LineStyle = xlDouble
                End With
                
                y = y + 2
            End If
        Next key
    End With
    
    '----------------------------------
    '   補助元帳
    '----------------------------------
    y = styGeneralLedger
    With wsSL
         .Columns("A").NumberFormat = "yyyy/mm/dd"
    
        '数値表示形式（小数点なし）
        .Range("E:F, H:H").NumberFormat = "#,##0;△#,##0;;@"
        .Range("E:F, H:H").HorizontalAlignment = xlRight
        .Range("G:G").HorizontalAlignment = xlCenter
        .Range("I:I").HorizontalAlignment = xlCenter
        
        For Each key In m_SubLedger
            Set acc = m_SubLedger.Item(key)
            If (acc.Lines.Count > 0) And ((acc.AccountCode < 30000) Or (acc.AccountCode >= 40000)) Then
                y = y + 1
                
                '--------------------------
                '勘定科目の出力
                .Range(.Cells(y, 1), .Cells(y, 8)).Merge
                If acc.MiddleAccount <> "" Then
                    .Cells(y, 1) = "科目：" & acc.MajorAccount & " － " & acc.MiddleAccount
                Else
                    .Cells(y, 1) = "科目：" & acc.MajorAccount
                End If
                .Cells(y, 1).Font.Underline = True
                .Cells(y, 1).Font.size = 16
                .Cells(y, 1).HorizontalAlignment = xlCenter
                y = y + 1
            
                '--------------------------
                '補助勘定科目の出力
                .Range(.Cells(y, 1), .Cells(y, 8)).Merge
                .Cells(y, 1) = "補助科目：［" & acc.SubAccountCate & "］" & acc.SubAccountName & acc.Class
                .Cells(y, 1).Font.Underline = True
                .Cells(y, 1).Font.size = 14
                .Cells(y, 1).HorizontalAlignment = xlCenter
                y = y + 2
                
                '--------------------------
                '勘定科目コードの出力
                .Cells(y, 1) = "勘定科目コード：" & CStr(acc.AccountCode)
                .Cells(y, 3) = "補助科目コード：" & CStr(acc.SubAccountCode)
                y = y + 1
                
                '--------------------------
                '項目行
                .Cells(y, 1) = "日付"
                .Cells(y, 2) = "相手大科目"
                .Cells(y, 3) = "相手中科目"
                .Cells(y, 4) = "摘要"
                .Cells(y, 5) = "借方"
                .Cells(y, 6) = "貸方"
                .Cells(y, 7) = "貸／借"
                .Cells(y, 8) = "残高"
                .Cells(y, 9) = "仕訳帳行"
                
                With .Range(.Cells(y, 1), .Cells(y, MaxColumn))
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Interior.Color = rgbLavender
                    .Font.Bold = True
                    .Borders.LineStyle = xlContinuous
                    .Borders(xlEdgeTop).LineStyle = xlDouble
                    .Borders(xlEdgeBottom).Weight = xlMedium
                End With
                y = y + 1
                sty = y
                
                '--------------------------
                '行
                For Each line In acc.Lines
                    .Cells(y, 1).Resize(1, MaxColumn).value = Array( _
                        line.entryDate, _
                        line.counterMajorAccount, _
                        line.counterMiddleAccount, _
                        line.summary, _
                        IIf(line.debit = 0, "-", line.debit), _
                        IIf(line.credit = 0, "-", line.credit), _
                        line.DebitCreditLabel, _
                        line.BalanceAbs, _
                        IIf(line.Row = 0, "", "=HYPERLINK(""#'①仕訳帳'!A" & line.Row & """,""" & line.Row & """)"))
                    .Range(.Cells(y, 1), .Cells(y, MaxColumn)).Borders.LineStyle = xlContinuous
                                        
                    'ゼブラ
                    If y Mod 2 = 0 Then
                        .Range(.Cells(y, 1), .Cells(y, MaxColumn)).Interior.Color = rgbAliceBlue
                    End If

                    y = y + 1
                Next line
                '--------------------------
                '合計行
                .Cells(y, 1) = "合計"
                .Cells(y, 5).Formula = "=SUM(E" & sty & ":E" & y - 1 & ")"
                .Cells(y, 6).Formula = "=SUM(F" & sty & ":F" & y - 1 & ")"
        
                With .Range(.Cells(y, 1), .Cells(y, MaxColumn))
                    .Font.Bold = True
                    .Interior.Color = rgbLightGoldenrodYellow
                    .Borders(xlEdgeTop).LineStyle = xlDouble
                End With
                
                y = y + 2
            End If
        Next key
    End With

End Sub

