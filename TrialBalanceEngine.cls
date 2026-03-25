VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrialBalanceEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'昨年度の試算表
Private m_TB_General As TrialBalance            '総勘定元帳ベース
Private m_TB_Sub As TrialBalance                '補助元帳ベース

'今年度の試算表
Private m_TB_LastGeneral As TrialBalance        '総勘定元帳ベース
Private m_TB_LastSub As TrialBalance            '補助元帳ベース

'==============================================================================
'   試算表クラスの取得
'==============================================================================
Public Property Get GeneralTrialBalance() As TrialBalance
    Set GeneralTrialBalance = m_TB_General
End Property

Public Property Get SubTrialBalance() As TrialBalance
    Set SubTrialBalance = m_TB_Sub
End Property

Public Property Get LastGeneralTrialBalance() As TrialBalance
    Set LastGeneralTrialBalance = m_TB_LastGeneral
End Property

Public Property Get LastSubTrialBalance() As TrialBalance
    Set LastSubTrialBalance = m_TB_LastSub
End Property

'==============================================================================
'   初期化
'==============================================================================
Private Sub Class_Initialize()

    Set m_TB_General = New TrialBalance
    Set m_TB_Sub = New TrialBalance
    
    Set m_TB_LastGeneral = New TrialBalance
    Set m_TB_LastSub = New TrialBalance

End Sub

'==============================================================================
'   前期実績から試算表を作成
'==============================================================================
Public Sub Read_LastResult(master As AccountMaster)

    Dim lastRowA As Long
    Dim lastRowB As Long
    Dim lastRowC As Long
    Dim lastRow As Long
    
    Dim y As Long
    Dim str As Variant
    Dim subAcc As Variant
    
    Dim Class As NetAssetType
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("前期実績")
    
    With ws
    
        lastRowA = .Cells(.Rows.Count, "A").End(xlUp).Row
        lastRowB = .Cells(.Rows.Count, "B").End(xlUp).Row
        lastRowC = .Cells(.Rows.Count, "C").End(xlUp).Row
        
        '最大値を採用
        lastRow = Application.WorksheetFunction.Max(lastRowA, lastRowB, lastRowC)
        
        For y = styPerformance To lastRow
            Debug.Print y
            str = .Cells(y, 1)
            If Not (IsNull(str) Or IsEmpty(str) Or str = "") Then
            
                If Not master.Exists(str) Then
                    Err.Raise vbObjectError + 3001, _
                    "TrialBalanceEngine.Read_LastResult", _
                    "貸方:勘定科目コード" & str & "は入力できません。"
                End If
                
                subAcc = .Cells(y, 4)
                If IsNull(subAcc) Or IsEmpty(subAcc) Or subAcc = "" Then
                    '--------------------------
                    ' 総勘定元帳ベース  （補助科目の記載が無い場合）
                    Call m_TB_LastGeneral.SetLastResult( _
                        master.GetAccount(str), _
                        0, _
                        "", _
                        "", _
                        Class, _
                        .Cells(y, 8), _
                        .Cells(y, 9), _
                        .Cells(y, 10), _
                        .Cells(y, 11))
                
                Else
                    '--------------------------
                    ' 補助元帳ベース    （補助科目の記載が有る場合）
                    If .Cells(y, 7) = "指定" Then
                        Class = natDesignated
                    Else
                        Class = natGeneral
                    End If
                    Call m_TB_LastSub.SetLastResult( _
                        master.GetAccount(str), _
                        .Cells(y, 4), _
                        .Cells(y, 5), _
                        .Cells(y, 6), _
                        Class, _
                        .Cells(y, 8), _
                        .Cells(y, 9), _
                        .Cells(y, 10), _
                        .Cells(y, 11))
                
                End If
                
            End If
            
        Next
    
    End With

End Sub
'==============================================================================
'   元帳から試算表を作成
'==============================================================================
Public Sub Build(ByVal engine As LedgerEngine)

    '--------------------------
    ' 総勘定元帳ベース
    Call m_TB_General.Build(engine.GeneralLedger)
    
    '--------------------------
    ' 総勘定元帳ベース
    Call m_TB_Sub.Build(engine.SubLedger)

End Sub

'==============================================================================
'   残高試算表 出力
'==============================================================================
Public Sub OutputSheet()

    Const START_ROW As Long = 5
    Const MAX_COL As Long = 11

    Dim ws As Worksheet
    Dim y As Long
    Dim sty As Long
    Dim line As TrialBalanceLine

    '----------------------------------
    ' ソート
    '----------------------------------
    Call m_TB_General.SortByAccountCode
    Call m_TB_Sub.SortByAccountCode
   
    '----------------------------------
    ' 初期化
    '----------------------------------
    Set ws = ThisWorkbook.Worksheets("④試算表")
    
    Call ClearSheet(ws, START_ROW)
    
    y = START_ROW

    With ws

        '==============================
        ' 列幅（帳票感）
        '==============================
        .Columns("B").ColumnWidth = 14
        .Columns("C").ColumnWidth = 18
        .Columns("D").ColumnWidth = 10
        .Columns("E").ColumnWidth = 10
        .Columns("F").ColumnWidth = 22
        .Columns("G:K").ColumnWidth = 14

        '==============================
        ' 数値書式
        '==============================
        .Range("H:K").NumberFormat = "_ * #,##0_ ;_ * △#,##0_ ;_ * ""-""??_ ;_ @_ "
        .Range("H:K").HorizontalAlignment = xlRight

        '********************************
        '*      残高試算表　まとめ      *
        '********************************
        '==============================
        ' 見出し（1行目）
        '==============================
        .Cells(y, 2) = "勘定科目"
        .Range(.Cells(y, 2), .Cells(y, 3)).Merge
        
        .Cells(y, 4) = "補助科目"
        .Range(.Cells(y, 4), .Cells(y, 6)).Merge
    
        .Cells(y, 7) = "一般／指定"
        .Range(.Cells(y, 7), .Cells(y + 1, 7)).Merge
    
        .Cells(y, 8) = "期首残高"
        .Range(.Cells(y, 8), .Cells(y + 1, 8)).Merge
    
        .Cells(y, 9) = "借方合計"
        .Range(.Cells(y, 9), .Cells(y + 1, 9)).Merge
    
        .Cells(y, 10) = "貸方合計"
        .Range(.Cells(y, 10), .Cells(y + 1, 10)).Merge
    
        .Cells(y, 11) = "期末残高"
        .Range(.Cells(y, 11), .Cells(y + 1, 11)).Merge
    
        '==============================
        ' 見出し（2行目）
        '==============================
        .Cells(y + 1, 2) = "大科目"
        .Cells(y + 1, 3) = "中科目"
        .Cells(y + 1, 4) = "コード"
        .Cells(y + 1, 5) = "区分"
        .Cells(y + 1, 6) = "補助科目"

        '==============================
        ' 見出しスタイル
        '==============================
        With .Range(.Cells(y, 2), .Cells(y + 1, 11))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = rgbLavender
            .Font.Bold = True
            
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            
            '外枠太線
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With

        y = y + 2
        sty = y

        '==============================
        ' データ部
        '==============================
        For Each line In m_TB_General.Lines

            .Cells(y, 1) = line.AccountCode
            .Cells(y, 2) = line.MajorAccount
            .Cells(y, 3) = line.MiddleAccount
            .Cells(y, 4) = ""
            .Cells(y, 5) = ""
            .Cells(y, 6) = ""
            .Cells(y, 7) = ""
            .Cells(y, 8) = IIf(line.BeginningBalanceForDisplay = 0, "-", line.BeginningBalanceForDisplay)
            .Cells(y, 9) = IIf(line.Debit = 0, "-", line.Debit)
            .Cells(y, 10) = IIf(line.Credit = 0, "-", line.Credit)
            .Cells(y, 11) = IIf(line.EndingBalance = 0, "-", line.EndingBalance)

            'インデント（階層）
            .Cells(y, 2).IndentLevel = 0
            .Cells(y, 3).IndentLevel = 1
            .Cells(y, 6).IndentLevel = 2

            '横罫線のみ（帳簿感）
            .Range(.Cells(y, 2), .Cells(y, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous

            'ゼブラ
            If y Mod 2 = 0 Then
                .Range(.Cells(y, 2), .Cells(y, 11)).Interior.Color = rgbAliceBlue
            End If

            y = y + 1

        Next line

        '==============================
        ' 合計行
        '==============================
        .Cells(y, 2) = "合計"

        .Cells(y, 8).Formula = "=SUM(H" & sty & ":H" & y - 1 & ")"
        .Cells(y, 9).Formula = "=SUM(I" & sty & ":I" & y - 1 & ")"
        .Cells(y, 10).Formula = "=SUM(J" & sty & ":J" & y - 1 & ")"
        .Cells(y, 11).Formula = "=SUM(K" & sty & ":K" & y - 1 & ")"

        With .Range(.Cells(y, 2), .Cells(y, 11))
            .Font.Bold = True
            .Interior.Color = rgbLightGoldenrodYellow
            .Borders(xlEdgeTop).LineStyle = xlDouble
        End With


        y = y + 3
        sty = y

        '********************************
        '*      残高試算表　まとめ      *
        '********************************
        '==============================
        ' 見出し（1行目）
        '==============================
        .Cells(y, 2) = "勘定科目"
        .Range(.Cells(y, 2), .Cells(y, 3)).Merge
        
        .Cells(y, 4) = "補助科目"
        .Range(.Cells(y, 4), .Cells(y, 6)).Merge
    
        .Cells(y, 7) = "一般／指定"
        .Range(.Cells(y, 7), .Cells(y + 1, 7)).Merge
    
        .Cells(y, 8) = "期首残高"
        .Range(.Cells(y, 8), .Cells(y + 1, 8)).Merge
    
        .Cells(y, 9) = "借方合計"
        .Range(.Cells(y, 9), .Cells(y + 1, 9)).Merge
    
        .Cells(y, 10) = "貸方合計"
        .Range(.Cells(y, 10), .Cells(y + 1, 10)).Merge
    
        .Cells(y, 11) = "期末残高"
        .Range(.Cells(y, 11), .Cells(y + 1, 11)).Merge
    
        '==============================
        ' 見出し（2行目）
        '==============================
        .Cells(y + 1, 2) = "大科目"
        .Cells(y + 1, 3) = "中科目"
        .Cells(y + 1, 4) = "コード"
        .Cells(y + 1, 5) = "区分"
        .Cells(y + 1, 6) = "補助科目"

        '==============================
        ' 見出しスタイル
        '==============================
        With .Range(.Cells(y, 2), .Cells(y + 1, 11))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = rgbLavender
            .Font.Bold = True
            
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            
            '外枠太線
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With

        y = y + 2

        '==============================
        ' データ部
        '==============================
        For Each line In m_TB_Sub.Lines

            .Cells(y, 1) = line.AccountCode
            .Cells(y, 2) = line.MajorAccount
            .Cells(y, 3) = line.MiddleAccount
            .Cells(y, 4) = line.SubAccountCode
            .Cells(y, 5) = line.SubAccountCate
            .Cells(y, 6) = line.SubAccountName
            If line.NetAssetType = natDesignated Then
                .Cells(y, 7) = "指定"
            Else
                .Cells(y, 7) = "一般"
            End If
            .Cells(y, 8) = IIf(line.BeginningBalanceForDisplay = 0, "-", line.BeginningBalanceForDisplay)
            .Cells(y, 9) = IIf(line.Debit = 0, "-", line.Debit)
            .Cells(y, 10) = IIf(line.Credit = 0, "-", line.Credit)
            .Cells(y, 11) = IIf(line.EndingBalance = 0, "-", line.EndingBalance)

            'インデント（階層）
            .Cells(y, 2).IndentLevel = 0
            .Cells(y, 3).IndentLevel = 1
            .Cells(y, 6).IndentLevel = 2

            '横罫線のみ（帳簿感）
            .Range(.Cells(y, 2), .Cells(y, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous

            'ゼブラ
            If y Mod 2 = 0 Then
                .Range(.Cells(y, 2), .Cells(y, 11)).Interior.Color = rgbAliceBlue
            End If

            y = y + 1

        Next line

        '==============================
        ' 合計行
        '==============================
        .Cells(y, 2) = "合計"

        .Cells(y, 8).Formula = "=SUM(H" & sty & ":H" & y - 1 & ")"
        .Cells(y, 9).Formula = "=SUM(I" & sty & ":I" & y - 1 & ")"
        .Cells(y, 10).Formula = "=SUM(J" & sty & ":J" & y - 1 & ")"
        .Cells(y, 11).Formula = "=SUM(K" & sty & ":K" & y - 1 & ")"

        With .Range(.Cells(y, 2), .Cells(y, 11))
            .Font.Bold = True
            .Interior.Color = rgbLightGoldenrodYellow
            .Borders(xlEdgeTop).LineStyle = xlDouble
        End With

    End With

End Sub
