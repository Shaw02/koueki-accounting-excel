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

Private Const styPerformance = 5

'昨年度の試算表
Private m_TB_General As TrialBalance            '総勘定元帳ベース
Private m_TB_Sub As TrialBalance                '補助元帳ベース

'今年度の試算表
Private m_TB_LastGeneral As TrialBalance        '総勘定元帳ベース
Private m_TB_LastSub As TrialBalance            '補助元帳ベース

'純資産（事業別）
Private m_NetAssets As Object                   'Dictionary

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

    '純資産の保持用
    Set m_NetAssets = CreateObject("Scripting.Dictionary")

End Sub

'==============================================================================
'   純資産勘定か判定
'==============================================================================
Private Function IsNetAssetAccount(ByVal code As Long) As Boolean

    Select Case code
        Case idNetAssets_End, idNetAssets_Begin, idNetAssets_Diff, _
             idSpNetAssets_End, idSpNetAssets_Begin, idSpNetAssets_Diff
            IsNetAssetAccount = True
        Case Else
            IsNetAssetAccount = False
    End Select

End Function

'==============================================================================
'   純資産LINE取得 or 作成
'==============================================================================
Private Function GetOrCreateNetAsset( _
    ByVal subCode As Long, _
    ByVal subCate As String, _
    ByVal subName As String, _
    ByVal nat As NetAssetType) As NetAssetLine

    Dim key As String
    key = subCode & "_" & nat

    If Not m_NetAssets.Exists(key) Then
        
        Dim na As NetAssetLine
        Set na = New NetAssetLine
        
        Call na.Initialize(subCode, subCate, subName, nat)
        
        m_NetAssets.Add key, na
        
    End If

    Set GetOrCreateNetAsset = m_NetAssets(key)

End Function

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
    
    Dim begin As Currency
    Dim debit As Currency
    Dim credit As Currency
    Dim ending As Currency
    
    Dim na As NetAssetLine   '★純資産用
        
    Call m_TB_LastGeneral.Clear
    Call m_TB_LastSub.Clear
    
    '----------------------------------
    ' 純資産初期化
    '----------------------------------
    Set m_NetAssets = CreateObject("Scripting.Dictionary")
    
    With ws
    
        .Columns("B").ColumnWidth = 14
        .Columns("C").ColumnWidth = 18
        .Columns("D").ColumnWidth = 10
        .Columns("E").ColumnWidth = 10
        .Columns("F").ColumnWidth = 22
        .Columns("G:K").ColumnWidth = 14
        
        lastRowA = .Cells(.Rows.Count, "A").End(xlUp).Row
        lastRowB = .Cells(.Rows.Count, "B").End(xlUp).Row
        lastRowC = .Cells(.Rows.Count, "C").End(xlUp).Row
        
        '最大値を採用
        lastRow = Application.WorksheetFunction.Max(lastRowA, lastRowB, lastRowC)
        
        For y = styPerformance To lastRow
        
            str = .Cells(y, 1)
            
            If Not (IsNull(str) Or IsEmpty(str) Or str = "") Then
            
                If Not master.Exists(str) Then
                    Err.Raise vbObjectError + 3001, _
                    "TrialBalanceEngine.Read_LastResult", _
                    "貸方:勘定科目コード" & str & "は入力できません。"
                End If
                
                begin = amount(.Cells(y, 8))
                debit = amount(.Cells(y, 9))
                credit = amount(.Cells(y, 10))
                ending = amount(.Cells(y, 11))
                
                '全て０だったら飛ばす
                If Not ((begin = 0) And (debit = 0) And (credit = 0) And (ending = 0)) Then
                
                    subAcc = .Cells(y, 4)
                    
                    '----------------------------------
                    ' 純資産（30000番台）の場合
                    '----------------------------------
                    If str = idNetAssets Or str = idSpNetAssets Then
                    
                        '==============================
                        ' 一般正味財産
                        '==============================
                        If str = idNetAssets Then
                        
                            '補助科目なし（0）のみ対象
                            If IsNull(subAcc) Or IsEmpty(subAcc) Or subAcc = "" Or subAcc = 0 Then
                            
                                Set na = GetOrCreateNetAsset( _
                                    0, _
                                    "", _
                                    "共通", _
                                    natGeneral)
                                
                                '期末残高のみ保持
                                Call na.SetBeginning(ending)
                            
                            End If
                        
                        '==============================
                        ' 指定正味財産
                        '==============================
                        ElseIf str = idSpNetAssets Then
                        
                            Set na = GetOrCreateNetAsset( _
                                .Cells(y, 4), _
                                .Cells(y, 5), _
                                .Cells(y, 6), _
                                natDesignated)
                            
                            '期末残高のみ保持
                            Call na.SetBeginning(ending)
                        
                        End If
                        
                        '元帳には入れない
                        GoTo ContinueLoop
                    
                    End If
                    
                    '----------------------------------
                    ' 通常（資産・負債など）
                    '----------------------------------
                    If IsNull(subAcc) Or IsEmpty(subAcc) Or subAcc = "" Then
                    
                        '--------------------------
                        ' 総勘定元帳ベース
                        '--------------------------
                        Call m_TB_LastGeneral.SetLastResult( _
                            master.GetAccount(str), _
                            0, _
                            "", _
                            "", _
                            natGeneral, _
                            begin, _
                            debit, _
                            credit, _
                            ending)
                    
                    Else
                    
                        '--------------------------
                        ' 補助元帳ベース
                        '--------------------------
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
                            begin, _
                            debit, _
                            credit, _
                            ending)
                    
                    End If
                
                End If
            End If

ContinueLoop:
        Next
    
    End With

End Sub

'==============================================================================
'   元帳から試算表を作成
'==============================================================================
Public Sub Build(ByVal engine As LedgerEngine)

    Call m_TB_General.Clear
    Call m_TB_Sub.Clear
    
    '--------------------------
    ' 総勘定元帳ベース
    Call m_TB_General.Build(engine.GeneralLedger)
    
    '--------------------------
    ' 総勘定元帳ベース
    Call m_TB_Sub.Build(engine.SubLedger)

    Call BuildNetAssets

End Sub

'==============================================================================
'   純資産集計
'------------------------------------------------------------------------------
'   Contents:
'       補助元帳ベースの試算表から純資産情報を生成
'
'   方針：
'       ・一般／指定で処理を分けない
'       ・差は Read_LastResult 側で吸収
'       ・総合（補助科目=0）も同時に集計する
'
'==============================================================================
Private Sub BuildNetAssets()

    Dim line As TrialBalanceLine
    Dim na As NetAssetLine
    Dim naTotal As NetAssetLine   '★総合用

    '----------------------------------
    ' 補助元帳ベースから生成
    '----------------------------------
    For Each line In m_TB_Sub.Lines

        '==============================
        ' PLのみ対象
        '==============================
        If line.account.CategoryEnum = ackRevenue _
        Or line.account.CategoryEnum = ackExpense Then

            '==============================
            ' 個別（補助科目あり）
            '==============================
            Set na = GetOrCreateNetAsset( _
                line.SubAccountCode, _
                line.SubAccountCate, _
                line.SubAccountName, _
                line.NetAssetType)

            na.AddAmount line.debit, line.credit

            '==============================
            ' 総合（補助科目=0）
            '==============================
            Set naTotal = GetOrCreateNetAsset( _
                0, _
                "", _
                "総合", _
                line.NetAssetType)

            naTotal.AddAmount line.debit, line.credit

        End If

    Next line

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
        y = OutputHeading(ws, y)
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
            .Cells(y, 8) = IIf(line.BeginningBalance = 0, "-", line.BeginningBalance)
            .Cells(y, 9) = IIf(line.debit = 0, "-", line.debit)
            .Cells(y, 10) = IIf(line.credit = 0, "-", line.credit)
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

        '********************************
        '*      残高試算表　まとめ      *
        '********************************
        y = OutputHeading(ws, y)
        sty = y

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
            .Cells(y, 8) = IIf(line.BeginningBalance = 0, "-", line.BeginningBalance)
            .Cells(y, 9) = IIf(line.debit = 0, "-", line.debit)
            .Cells(y, 10) = IIf(line.credit = 0, "-", line.credit)
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
        
        '********************************
        '*      純資産出力              *
        '********************************
        Dim na As NetAssetLine
        Dim v As Variant

        '========================================
        ' 一般正味財産
        '========================================
        .Cells(y, 2) = "【一般正味財産】"
        .Cells(y, 2).Font.Bold = True
        y = OutputHeading(ws, y + 1)
        sty = y
        '------------------------------
        ' 個別（sub ≠ 0）
        '------------------------------
        For Each v In m_NetAssets.Keys

            Set na = m_NetAssets(v)

            If na.NetAssetType = natGeneral Then

                If na.SubAccountCode <> 0 Then

                    na.CalcEnding

                    .Cells(y, 1) = idNetAssets
                    .Cells(y, 2) = "一般正味財産"
                    .Cells(y, 4) = na.SubAccountCode
                    .Cells(y, 5) = na.SubAccountCate
                    .Cells(y, 6) = na.SubAccountName
                    .Cells(y, 7) = "一般"

                    .Cells(y, 8) = IIf(na.BeginningBalance = 0, "-", na.BeginningBalance)
                    .Cells(y, 9) = IIf(na.debit = 0, "-", na.debit)
                    .Cells(y, 10) = IIf(na.credit = 0, "-", na.credit)
                    .Cells(y, 11) = IIf(na.EndingBalance = 0, "-", na.EndingBalance)
            
                    '横罫線のみ（帳簿感）
                    .Range(.Cells(y, 2), .Cells(y, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        
                    'ゼブラ
                    If y Mod 2 = 0 Then
                        .Range(.Cells(y, 2), .Cells(y, 11)).Interior.Color = rgbAliceBlue
                    End If

                    y = y + 1

                End If

            End If

        Next

        '------------------------------
        ' 総合（sub = 0）
        '------------------------------
        Set na = GetOrCreateNetAsset(0, "", "総合", natGeneral)
        na.CalcEnding

        .Cells(y, 1) = idNetAssets
        .Cells(y, 2) = "一般正味財産"
        .Cells(y, 4) = ""
        .Cells(y, 5) = ""
        .Cells(y, 6) = "総合"
        .Cells(y, 7) = "一般"

        .Cells(y, 8) = IIf(na.BeginningBalance = 0, "-", na.BeginningBalance)
        .Cells(y, 9) = IIf(na.debit = 0, "-", na.debit)
        .Cells(y, 10) = IIf(na.credit = 0, "-", na.credit)
        .Cells(y, 11) = IIf(na.EndingBalance = 0, "-", na.EndingBalance)

        With .Range(.Cells(y, 2), .Cells(y, 11))
            .Font.Bold = True
            .Interior.Color = rgbLightGoldenrodYellow
            .Borders(xlEdgeTop).LineStyle = xlDouble
        End With

        
        y = y + 3

        '========================================
        ' 指定正味財産
        '========================================
        .Cells(y, 2) = "【指定正味財産】"
        .Cells(y, 2).Font.Bold = True
        y = OutputHeading(ws, y + 1)
        sty = y
        '------------------------------
        ' 個別（sub ≠ 0）
        '------------------------------
        For Each v In m_NetAssets.Keys

            Set na = m_NetAssets(v)

            If na.NetAssetType = natDesignated Then

                If na.SubAccountCode <> 0 Then

                    na.CalcEnding

                    .Cells(y, 1) = idSpNetAssets
                    .Cells(y, 2) = "指定正味財産"
                    .Cells(y, 4) = na.SubAccountCode
                    .Cells(y, 5) = na.SubAccountCate
                    .Cells(y, 6) = na.SubAccountName
                    .Cells(y, 7) = "指定"

                    .Cells(y, 8) = IIf(na.BeginningBalance = 0, "-", na.BeginningBalance)
                    .Cells(y, 9) = IIf(na.debit = 0, "-", na.debit)
                    .Cells(y, 10) = IIf(na.credit = 0, "-", na.credit)
                    .Cells(y, 11) = IIf(na.EndingBalance = 0, "-", na.EndingBalance)

                    '横罫線のみ（帳簿感）
                    .Range(.Cells(y, 2), .Cells(y, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        
                    'ゼブラ
                    If y Mod 2 = 0 Then
                        .Range(.Cells(y, 2), .Cells(y, 11)).Interior.Color = rgbAliceBlue
                    End If

                    y = y + 1

                End If

            End If

        Next

        '------------------------------
        ' 総合（sub = 0）
        '------------------------------
        Set na = GetOrCreateNetAsset(0, "", "総合", natDesignated)
        na.CalcEnding

        .Cells(y, 1) = idSpNetAssets
        .Cells(y, 2) = "指定正味財産"
        .Cells(y, 4) = ""
        .Cells(y, 5) = ""
        .Cells(y, 6) = "合計"
        .Cells(y, 7) = "指定"

        .Cells(y, 8) = IIf(na.BeginningBalance = 0, "-", na.BeginningBalance)
        .Cells(y, 9) = IIf(na.debit = 0, "-", na.debit)
        .Cells(y, 10) = IIf(na.credit = 0, "-", na.credit)
        .Cells(y, 11) = IIf(na.EndingBalance = 0, "-", na.EndingBalance)

        With .Range(.Cells(y, 2), .Cells(y, 11))
            .Font.Bold = True
            .Interior.Color = rgbLightGoldenrodYellow
            .Borders(xlEdgeTop).LineStyle = xlDouble
        End With
    
    End With

End Sub

Private Function OutputHeading(ws As Worksheet, y As Long)
        
    With ws
    
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
    
    End With
    
    OutputHeading = y + 2

End Function
