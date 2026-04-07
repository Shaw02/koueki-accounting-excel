VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FinancialStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MaxColumn As Long = 5
Private Const START_ROW As Long = 5
Private Const MASTER_SHEET_NAME As String = "財務諸表マスタ"

Private m_Master As FinancialStatementMaster

Private Sub Class_Initialize()
    Set m_Master = New FinancialStatementMaster
End Sub

Private Function BuildKey(ByVal lineID As String, ByVal category As String) As String
    If Trim$(category) = "" Then category = "*"
    BuildKey = UCase$(Trim$(lineID)) & "|" & UCase$(Trim$(category))
End Function

'==============================================================================
'       財務諸表の生成～出力
'------------------------------------------------------------------------------
'   Input:
'       tbEngine        試算表クラス
'
'==============================================================================
Public Sub BuildAndOutput(ByVal tbEngine As TrialBalanceEngine)

    Dim wsMaster As Worksheet
    Set wsMaster = ThisWorkbook.Worksheets(MASTER_SHEET_NAME)
    
    Call m_Master.LoadFromWorksheet(wsMaster)

    Dim thisAmounts As Object
    Dim lastAmounts As Object

    Set thisAmounts = BuildStatementAmountMap(tbEngine.SubTrialBalance)
    Set lastAmounts = BuildStatementAmountMap(tbEngine.LastSubTrialBalance)
    
    Call AddNetAssetAmounts(thisAmounts, tbEngine, False)
    Call AddNetAssetAmounts(lastAmounts, tbEngine, True)

    Call RollupAggregateTo(thisAmounts, False)
    Call RollupAggregateTo(lastAmounts, False)

    Call ResolveFormulaAmounts(thisAmounts)
    Call ResolveFormulaAmounts(lastAmounts)

    Call RollupAggregateTo(thisAmounts, True)
    Call RollupAggregateTo(lastAmounts, True)
        
    Call OutputAll(thisAmounts, lastAmounts)

End Sub

Private Function BuildStatementAmountMap(ByVal tb As TrialBalance) As Object

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim line As TrialBalanceLine
    Dim lineIDs As Variant
    Dim lineID As String
    Dim accountingCategory As String
    Dim amount As Currency
    Dim i As Long
    Dim mapped As Boolean
    
    For Each line In tb.Lines
        amount = line.amount()
        mapped = False

        If Trim$(line.account.fsLineIDs) <> "" Then
            lineIDs = Split(line.account.fsLineIDs, ";")
            For i = LBound(lineIDs) To UBound(lineIDs)
                lineID = UCase$(Trim$(CStr(lineIDs(i))))
                If lineID <> "" Then
                    If Not ExistsLineID(lineID) Then
                        Err.Raise vbObjectError + 7001, _
                            "BuildStatementAmountMap", _
                            "Undefined FS_ID mapping in account master: " & lineID
                    End If

                    mapped = True
                    If amount <> 0 Then
                        accountingCategory = UCase$(Trim$(line.SubAccountCate))
                        Call AddAmountToMap(result, lineID, "*", amount)
                        If accountingCategory <> "" Then
                            Call AddAmountToMap(result, lineID, accountingCategory, amount)
                        End If
                    End If
                End If
            Next i
        End If

        If Not mapped Then
            Err.Raise vbObjectError + 7002, _
                "BuildStatementAmountMap", _
                "TrialBalanceLine is not mapped to FS_ID. AccountCode=" & CStr(line.AccountCode)
        End If
    Next

    Set BuildStatementAmountMap = result

End Function

Private Function ExistsLineID(ByVal lineID As String) As Boolean

    Dim line As FinancialStatementMasterLine
    Dim targetID As String

    targetID = UCase$(Trim$(lineID))

    For Each line In m_Master.Lines
        If UCase$(Trim$(line.lineID)) = targetID Then
            ExistsLineID = True
            Exit Function
        End If
    Next

    ExistsLineID = False

End Function

Private Sub AddAmountToMap( _
    ByVal accountMap As Object, _
    ByVal lineID As String, _
    ByVal accountingCategory As String, _
    ByVal amount As Currency)

    Dim key As String
    key = UCase$(Trim$(lineID)) & "|" & UCase$(Trim$(accountingCategory))

    If Not accountMap.Exists(key) Then
        accountMap.Add key, CCur(0)
    End If
    accountMap(key) = CCur(accountMap(key)) + amount
        
    Debug.Print key & amount

End Sub

'==============================================================================
'       純資産（期首・期末・増減）のマップ追加
'------------------------------------------------------------------------------
'   Contents:
'       ・貸借対照表用（期末）
'       ・注記用（期首・増減）
'       の両方で使えるように、純資産をフルセットで投入する
'
'   Input:
'       accountMap     出力用Dictionary
'       tbEngine       試算表エンジン
'       isLastYear     前年データかどうか
'
'==============================================================================
Private Sub AddNetAssetAmounts( _
    ByVal accountMap As Object, _
    ByVal tbEngine As TrialBalanceEngine, _
    ByVal isLastYear As Boolean)

    If isLastYear Then

        '==============================
        ' 一般純資産
        '==============================
        Call AddAmountToMap(accountMap, "BS_NET_GENERAL_BEGIN", "*", tbEngine.LastGeneralNetBegin)
        Call AddAmountToMap(accountMap, "BS_NET_GENERAL_END", "*", tbEngine.LastGeneralNetEnd)
        Call AddAmountToMap(accountMap, "BS_NET_GENERAL_DIFF", "*", tbEngine.LastGeneralNetChange)

        '==============================
        ' 指定純資産
        '==============================
        Call AddAmountToMap(accountMap, "BS_NET_DESIG_BEGIN", "*", tbEngine.LastDesignatedNetBegin)
        Call AddAmountToMap(accountMap, "BS_NET_DESIG_END", "*", tbEngine.LastDesignatedNetEnd)
        Call AddAmountToMap(accountMap, "BS_NET_DESIG_DIFF", "*", tbEngine.LastDesignatedNetChange)

    Else

        '==============================
        ' 一般純資産
        '==============================
        Call AddAmountToMap(accountMap, "BS_NET_GENERAL_BEGIN", "*", tbEngine.GeneralNetAssetBegin)
        Call AddAmountToMap(accountMap, "BS_NET_GENERAL_END", "*", tbEngine.GeneralNetAssetEnd)
        Call AddAmountToMap(accountMap, "BS_NET_GENERAL_DIFF", "*", tbEngine.GeneralNetAssetChange)

        '==============================
        ' 指定純資産
        '==============================
        Call AddAmountToMap(accountMap, "BS_NET_DESIG_BEGIN", "*", tbEngine.DesignatedNetAssetBegin)
        Call AddAmountToMap(accountMap, "BS_NET_DESIG_END", "*", tbEngine.DesignatedNetAssetEnd)
        Call AddAmountToMap(accountMap, "BS_NET_DESIG_DIFF", "*", tbEngine.DesignatedNetAssetChange)

    End If

End Sub

'==============================================================================
'       財務諸表出力
'==============================================================================
Private Sub OutputAll(ByVal thisAmounts As Object, ByVal lastAmounts As Object)

    Dim line As FinancialStatementMasterLine
    
    Dim vThis As Currency
    Dim vLast As Currency
    Dim vDiff As Currency

    Dim nextRowBySheet As Object
    Set nextRowBySheet = CreateObject("Scripting.Dictionary")

    Call InitializeOutputSheets(nextRowBySheet)

    For Each line In m_Master.Lines
    
        If line.enabled Then
    
            vThis = GetAmountByLine(thisAmounts, line)
            vLast = GetAmountByLine(lastAmounts, line)
            vDiff = vThis - vLast
            
            '------------------------------------------
            ' シート出力
            '------------------------------------------
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Worksheets(line.TargetSheet)
    
            Dim y As Long
            y = CLng(nextRowBySheet(ws.Name))
            
            ws.Cells(y, 2) = line.LineLabel
            ws.Cells(y, 2).HorizontalAlignment = line.LineLabelAligh
    
            'タイトル、集計項目でない無い場合、且つ、前期・当期残高が0の場合は、表示を省略する。
            If Not ShouldSkipOutputLine(line, vThis, vLast) Then
                If line.LineType = "TITLE" Then
                    ws.Cells(y, 3).value = ""
                    ws.Cells(y, 4).value = ""
                    ws.Cells(y, 5).value = ""
                Else
                    ws.Cells(y, 3).value = vThis
                    ws.Cells(y, 4).value = vLast
                    ws.Cells(y, 5).value = vDiff
                End If
                                    
                '横罫線のみ（帳簿感）
                ws.Range(ws.Cells(y, 2), ws.Cells(y, MaxColumn)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
                'ゼブラ
                If y Mod 2 = 0 Then
                    ws.Range(ws.Cells(y, 2), ws.Cells(y, MaxColumn)).Interior.Color = rgbAliceBlue
                End If

                nextRowBySheet(ws.Name) = y + 1
            End If
        End If
    Next


End Sub

'==============================================================================
'       ヘッダー出力
'==============================================================================
Private Sub InitializeOutputSheets(ByVal nextRowBySheet As Object)

    Dim line As FinancialStatementMasterLine
    Dim ws As Worksheet

    For Each line In m_Master.Lines
        If line.enabled Then
            If Not nextRowBySheet.Exists(line.TargetSheet) Then
                Set ws = ThisWorkbook.Worksheets(line.TargetSheet)
                nextRowBySheet.Add line.TargetSheet, START_ROW + 1
                Call ClearSheet(ws, START_ROW)
                Call SetupSheetHeader(ws, line.statementID)
            End If
        End If
    Next

End Sub

'==============================================================================
'       ヘッダー出力
'==============================================================================
Private Sub SetupSheetHeader(ByVal ws As Worksheet, ByVal statementID As String)

    ws.Range("C:E").NumberFormat = "_ * #,##0_ ;_ * △#,##0_ ;_ * ""-""??_ ;_ @_ "
    ws.Range("C:E").HorizontalAlignment = xlRight

    ws.Cells(START_ROW, 2) = "科目"
    ws.Cells(START_ROW, 3) = "当期"
    ws.Cells(START_ROW, 4) = "前期"
    ws.Cells(START_ROW, 5) = "増減"

    With ws.Range(ws.Cells(START_ROW, 2), ws.Cells(START_ROW, MaxColumn))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = rgbLavender
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With

End Sub

'==============================================================================
'       表示するか？
'==============================================================================
Private Function ShouldSkipOutputLine(ByVal line As FinancialStatementMasterLine, ByVal vThis As Currency, ByVal vLast As Currency) As Boolean

    

    If line.LineType = "TITLE" Or line.LineType = "ACCOUNTS" Then
        ShouldSkipOutputLine = False
        Exit Function
    End If

    If (line.LineType = "ACCOUNT" Or line.LineType = "ACCOUNTS") And vThis = 0 And vLast = 0 Then
        ShouldSkipOutputLine = True
        Exit Function
    End If
       
    ShouldSkipOutputLine = False  '←これ必須
    
End Function

'==============================================================================
'       金額取得
'==============================================================================
Private Function GetAmountByLine(ByVal amountMap As Object, ByVal line As FinancialStatementMasterLine) As Currency

    Dim key As String
    key = BuildKey(line.lineID, line.accountingCategory)

    If amountMap.Exists(key) Then
        GetAmountByLine = CCur(amountMap(key))
    Else
        GetAmountByLine = 0
    End If

End Function


Private Sub ResolveFormulaAmounts(ByVal values As Object)

    Dim line As FinancialStatementMasterLine
    Dim key As String

    For Each line In m_Master.Lines
        If line.enabled Then
            If line.LineType = "FORMULA" Or line.LineType = "CALC" Then
                key = BuildKey(line.lineID, line.accountingCategory)
                values(key) = ResolveFormula(line.formula, values)
            End If
        End If
    Next

End Sub

Private Sub RollupAggregateTo(ByVal amountMap As Object, ByVal includeFormula As Boolean)

    Dim line As FinancialStatementMasterLine
    Dim childType As String

    For Each line In m_Master.Lines
        If line.enabled Then
            childType = UCase$(Trim$(line.LineType))

            If childType <> "TITLE" And Trim$(line.AggregateTo) <> "" Then
                If (includeFormula And (childType = "FORMULA" Or childType = "CALC")) _
                    Or ((Not includeFormula) And (childType = "ACCOUNT" Or childType = "ACCOUNTS")) Then
                    Call RollupLineAmount(line, amountMap)
                End If
            End If
        End If
    Next

End Sub

Private Sub RollupLineAmount(ByVal child As FinancialStatementMasterLine, ByVal amountMap As Object)

    Dim currentLine As FinancialStatementMasterLine
    Dim parentLine As FinancialStatementMasterLine
    Dim visited As Object
    Dim currentAggregateTo As String
    Dim childKey As String
    Dim parentKey As String
    Dim childAmount As Currency

    Set visited = CreateObject("Scripting.Dictionary")
    Set currentLine = child

    childKey = BuildKey(child.lineID, child.accountingCategory)
    If amountMap.Exists(childKey) Then
        childAmount = CCur(amountMap(childKey))
    Else
        childAmount = 0
    End If

    Do While Trim$(currentLine.AggregateTo) <> ""
        currentAggregateTo = UCase$(Trim$(currentLine.AggregateTo))

        If visited.Exists(currentAggregateTo) Then
            Err.Raise vbObjectError + 7101, _
                "RollupLineAmount", _
                "Circular AggregateTo reference detected: " & currentAggregateTo
        End If
        visited.Add currentAggregateTo, True

        Set parentLine = GetAggregateParentLine(currentAggregateTo, currentLine.accountingCategory)
        parentKey = BuildKey(parentLine.lineID, parentLine.accountingCategory)

        If Not amountMap.Exists(parentKey) Then
            amountMap.Add parentKey, CCur(0)
        End If
        amountMap(parentKey) = CCur(amountMap(parentKey)) + childAmount

        Set currentLine = parentLine
    Loop

End Sub

Private Function GetAggregateParentLine(ByVal parentLineID As String, ByVal category As String) As FinancialStatementMasterLine

    Dim normalizedParent As String
    normalizedParent = UCase$(Trim$(parentLineID))

    If normalizedParent = "" Then
        Err.Raise vbObjectError + 7102, _
            "GetAggregateParentLine", _
            "AggregateTo target is blank."
    End If

    If m_Master.Exists(normalizedParent, category) Then
        Set GetAggregateParentLine = m_Master.GetLine(normalizedParent, category)
        Exit Function
    End If

    If m_Master.Exists(normalizedParent, "*") Then
        Set GetAggregateParentLine = m_Master.GetLine(normalizedParent, "*")
        Exit Function
    End If

    Err.Raise vbObjectError + 7103, _
        "GetAggregateParentLine", _
        "AggregateTo target does not exist: " & normalizedParent & "|" & UCase$(Trim$(category))

End Function


Private Function ResolveFormula(ByVal formula As String, ByVal values As Object) As Currency

    If Trim$(formula) = "" Then
        ResolveFormula = 0
        Exit Function
    End If

    Dim expr As String
    expr = formula

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    With re
        .Pattern = "\[([^\]]+)\]"
        .Global = True
    End With

    Dim matches As Object
    Set matches = re.Execute(expr)

    Dim m As Object
    For Each m In matches
        
        Dim token As String
        token = m.SubMatches(0)
        
        Dim lineID As String
        Dim category As String
        
        '------------------------------------------
        ' [PL_BIZ|公1] or [PL_BIZ]
        '------------------------------------------
        If InStr(token, "|") > 0 Then
            lineID = UCase$(Trim$(Split(token, "|")(0)))
            category = UCase$(Trim$(Split(token, "|")(1)))
        Else
            lineID = UCase$(Trim$(token))
            category = "*"
        End If

        If category = "" Then category = "*"

        If category = "*" Then
            If Not ExistsLineID(lineID) Then
                Err.Raise vbObjectError + 3002, _
                    "ResolveFormula", _
                    "Undefined FS_ID in formula: [" & token & "]"
            End If
        Else
            If Not m_Master.Exists(lineID, category) Then
                Err.Raise vbObjectError + 3003, _
                    "ResolveFormula", _
                    "Undefined FS_ID/category in formula: [" & token & "]"
            End If
        End If
        
        Dim key As String
        key = BuildKey(lineID, category)
        
        Dim val As Currency
        
        '------------------------------------------
        ' ① 完全一致
        '------------------------------------------
        If values.Exists(key) Then
            val = values(key)
        
        '------------------------------------------
        ' ② "*"（合算）
        '------------------------------------------
        ElseIf category = "*" Then
            
            Dim k As Variant
            val = 0
            
            For Each k In values.Keys
                If Left$(UCase$(CStr(k)), Len(UCase$(lineID)) + 1) = UCase$(lineID) & "|" Then
                    val = val + values(k)
                End If
            Next
        
        '------------------------------------------
        ' ③ エラー
        '------------------------------------------
        Else
            Err.Raise vbObjectError + 3001, _
                      "ResolveFormula", _
                      "キーが存在しません: [" & token & "]"
        End If
        
        expr = Replace(expr, "[" & token & "]", CStr(val))
        
    Next

    expr = Replace(expr, " ", "")
    ResolveFormula = CCur(Application.Evaluate(expr))

End Function

