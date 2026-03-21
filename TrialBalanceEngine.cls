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

Private m_TB_General As TrialBalance
Private m_TB_Sub As TrialBalance

Public Property Get GeneralTrialBalance() As TrialBalance
    Set GeneralTrialBalance = m_TB_General
End Property

Public Property Get SubTrialBalance() As TrialBalance
    Set SubTrialBalance = m_TB_Sub
End Property

Private Sub Class_Initialize()

    Set m_TB_General = New TrialBalance
    Set m_TB_Sub = New TrialBalance

End Sub

'===============================
' 試算表 作成
'===============================
Public Sub Build(ByVal engine As LedgerEngine)

    ' 総勘定元帳ベース
    Call m_TB_General.Build(engine.GeneralLedger)
    
    ' 補助元帳ベース
    Call m_TB_Sub.Build(engine.SubLedger)

End Sub

'==============================================================================
'   残高試算表 出力（Engine側）
'==============================================================================
Public Sub OutputSheet()

    Const START_ROW As Long = 7
    Const MAX_COL As Long = 8

    Dim ws As Worksheet
    Dim y As Long
    Dim line As TrialBalanceLine

    '----------------------------------
    ' ソート（追加）
    '----------------------------------
    Call m_TB_General.SortByAccountCode
    Call m_TB_Sub.SortByAccountCode
   
    '----------------------------------
    '   初期化
    '----------------------------------
    Set ws = ThisWorkbook.Worksheets("TEST")
    
    '----------------------------------
    '   シートクリア
    '----------------------------------
    Call ClearSheet(ws, START_ROW)
    
    '----------------------------------
    '   シート出力
    '----------------------------------
    y = START_ROW

    With ws

        .Range("E:H").NumberFormat = "#,##0;△#,##0;;@"
        .Range("E:H").HorizontalAlignment = xlRight
        .Columns("A:D").HorizontalAlignment = xlLeft

        .Cells(y, 1) = "科目"
        .Cells(y, 2) = "区分"
        .Cells(y, 3) = "補助"
        .Cells(y, 4) = "一般／指定"
        .Cells(y, 5) = "期首残高"
        .Cells(y, 6) = "借方合計"
        .Cells(y, 7) = "貸方合計"
        .Cells(y, 8) = "期末残高"

        .Range(.Cells(y, 1), .Cells(y, MAX_COL)).Font.Bold = True
        .Range(.Cells(y, 1), .Cells(y, MAX_COL)).HorizontalAlignment = xlCenter
        .Range(.Cells(y, 1), .Cells(y, MAX_COL)).Borders.LineStyle = xlContinuous
        .Range(.Cells(y, 1), .Cells(y, MAX_COL)).Borders(xlEdgeTop).LineStyle = xlDouble
        .Range(.Cells(y, 1), .Cells(y, MAX_COL)).Borders(xlEdgeBottom).Weight = xlMedium

        y = y + 1

        For Each line In m_TB_General.Lines

            .Cells(y, 1) = line.AccountName
            .Cells(y, 2) = line.Sector
            .Cells(y, 3) = line.Project

            If line.FundType = natDesignated Then
                .Cells(y, 4) = "指定"
            Else
                .Cells(y, 4) = "一般"
            End If

            If line.FS = "BS" Then
                .Cells(y, 5) = IIf(line.BeginningBalance = 0, "", line.BeginningBalance)
            End If
            .Cells(y, 6) = IIf(line.Debit = 0, "", line.Debit)
            .Cells(y, 7) = IIf(line.Credit = 0, "", line.Credit)
            .Cells(y, 8) = IIf(line.EndingBalance = 0, "", line.EndingBalance)

            .Range(.Cells(y, 1), .Cells(y, MAX_COL)).Borders.LineStyle = xlContinuous

            y = y + 1

        Next line

        .Cells(y, 1) = "合計"
        .Cells(y, 1).Font.Bold = True

        .Cells(y, 5).Formula = "=SUM(E" & START_ROW + 1 & ":E" & y - 1 & ")"
        .Cells(y, 6).Formula = "=SUM(F" & START_ROW + 1 & ":F" & y - 1 & ")"
        .Cells(y, 7).Formula = "=SUM(G" & START_ROW + 1 & ":G" & y - 1 & ")"
        .Cells(y, 8).Formula = "=SUM(H" & START_ROW + 1 & ":H" & y - 1 & ")"

        .Range(.Cells(y, 1), .Cells(y, MAX_COL)).Borders.LineStyle = xlContinuous
        .Range(.Cells(y, 1), .Cells(y, MAX_COL)).Borders(xlEdgeTop).LineStyle = xlDouble

    End With

End Sub
