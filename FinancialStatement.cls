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

Private m_Master As FinancialStatementMaster

Private Sub Class_Initialize()
    Set m_Master = New FinancialStatementMaster
End Sub

Public Sub BuildAndOutput(ByVal tbEngine As TrialBalanceEngine)

    Dim wsMaster As Worksheet
    Set wsMaster = ThisWorkbook.Worksheets("財務諸表マスタ")

    Call m_Master.LoadFromWorksheet(wsMaster)

    Dim thisAmounts As Object
    Dim lastAmounts As Object

    Set thisAmounts = BuildAccountAmountMap(tbEngine.SubTrialBalance)
    Set lastAmounts = BuildAccountAmountMap(tbEngine.LastSubTrialBalance)

    Call OutputAll(thisAmounts, lastAmounts)

End Sub

Private Function BuildAccountAmountMap(ByVal tb As TrialBalance) As Object

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim line As TrialBalanceLine
    Dim lineIDs As Variant
    Dim lineID As String
    Dim amount As Currency
    Dim i As Long

    For Each line In tb.Lines
        If line.account.FSEnum = afsBS Then
            amount = line.amount(afsBS, "END")
        ElseIf line.account.FSEnum = afsPL Then
            amount = line.amount(afsPL, "FLOW")
        Else
            amount = 0
        End If

        If amount = 0 Then GoTo ContinueLoop

        If Trim$(line.account.fsLineIDs) <> "" Then
            lineIDs = Split(line.account.fsLineIDs, ";")
            For i = LBound(lineIDs) To UBound(lineIDs)
                lineID = UCase$(Trim$(CStr(lineIDs(i))))
                If lineID <> "" Then
                    If Not result.Exists(lineID) Then
                        result.Add lineID, CCur(0)
                    End If
                    result(lineID) = CCur(result(lineID)) + amount
                End If
            Next i
        End If

ContinueLoop:
    Next

    Set BuildAccountAmountMap = result

End Function

Private Sub OutputAll(ByVal thisAmounts As Object, ByVal lastAmounts As Object)

    Dim valuesThis As Object
    Dim valuesLast As Object

    Dim wsBS As Worksheet
    Dim wsPL As Worksheet
    Dim wsPL2 As Worksheet
    
    Set wsBS = ThisWorkbook.Worksheets("貸借対照表")
    Set wsPL = ThisWorkbook.Worksheets("活動計算書")
    Set wsPL2 = ThisWorkbook.Worksheets("活動計算書の注記")
    
    Set valuesThis = CreateObject("Scripting.Dictionary")
    Set valuesLast = CreateObject("Scripting.Dictionary")

    Dim line As FinancialStatementMasterLine
    Dim vThis As Currency
    Dim vLast As Currency

    Dim y As Long
    y = 5

    '==============================
    ' 数値書式
    '==============================
    wsBS.Range("C:D").NumberFormat = "_ * #,##0_ ;_ * △#,##0_ ;_ * ""-""??_ ;_ @_ "
    wsBS.Range("C:D").HorizontalAlignment = xlRight
    
    wsPL.Range("C:D").NumberFormat = "_ * #,##0_ ;_ * △#,##0_ ;_ * ""-""??_ ;_ @_ "
    wsPL.Range("C:D").HorizontalAlignment = xlRight
    
    wsPL2.Range("C:D").NumberFormat = "_ * #,##0_ ;_ * △#,##0_ ;_ * ""-""??_ ;_ @_ "
    wsPL2.Range("C:D").HorizontalAlignment = xlRight


    For Each line In m_Master.Lines

        If line.enabled Then
            vThis = ResolveLineAmount(line, thisAmounts, valuesThis)
            vLast = ResolveLineAmount(line, lastAmounts, valuesLast)

            valuesThis(UCase$(line.lineID)) = vThis
            valuesLast(UCase$(line.lineID)) = vLast

            Dim ws As Worksheet
            Set ws = ThisWorkbook.Worksheets(line.TargetSheet)

            ws.Cells(y, 2) = line.LineLabel
            ws.Cells(y, 2).HorizontalAlignment = line.LineLabelAligh
            If line.LineType = "TITLE" Then
                ws.Cells(y, 3).value = ""
                ws.Cells(y, 4).value = ""
            Else
                ws.Cells(y, 3).value = vThis
                ws.Cells(y, 4).value = vLast
            End If
            y = y + 1
        End If
    Next

End Sub

Private Function ResolveLineAmount( _
    ByVal line As FinancialStatementMasterLine, _
    ByVal accountMap As Object, _
    ByVal resolvedValues As Object) As Currency

    Select Case line.LineType
        Case "ACCOUNT", "ACCOUNTS"
            ResolveLineAmount = ResolveAccountAmount(line, accountMap) * line.signFactor
        Case "FORMULA", "CALC"
            ResolveLineAmount = ResolveFormula(line.formula, resolvedValues) * line.signFactor
        Case Else
            ResolveLineAmount = 0
    End Select

End Function

Private Function ResolveAccountAmount(ByVal def As FinancialStatementMasterLine, ByVal accountMap As Object) As Currency

    Dim code As String

    code = UCase$(Trim$(def.lineID))
    If accountMap.Exists(code) Then
        ResolveAccountAmount = CCur(accountMap(code))
    Else
        ResolveAccountAmount = 0
    End If
    Exit Function

End Function

Private Function ResolveFormula(ByVal formula As String, ByVal resolvedValues As Object) As Currency

    If Trim$(formula) = "" Then
        ResolveFormula = 0
        Exit Function
    End If

    Dim expr As String
    expr = formula

    '------------------------------------------
    ' 未解決キーのチェック
    '------------------------------------------
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    With re
        .Pattern = "\[([^\]]+)\]"   ' [KEY] を抽出
        .Global = True
    End With
    
    Dim matches As Object
    Set matches = re.Execute(expr)
    
    Dim m As Object
    For Each m In matches
        Dim key As String
        key = m.SubMatches(0)
        
        If Not resolvedValues.Exists(key) Then
            Err.Raise vbObjectError + 1000, _
                      "ResolveFormula", _
                      "キーが存在しません: [" & key & "]"
        End If
    Next

    '------------------------------------------
    ' 置換処理
    '------------------------------------------
    Dim k As Variant
    For Each k In resolvedValues.Keys
        expr = Replace(expr, "[" & k & "]", CStr(CCur(resolvedValues(k))))
    Next

    expr = Replace(expr, " ", "")

    '------------------------------------------
    ' 評価
    '------------------------------------------
    ResolveFormula = CCur(Application.Evaluate(expr))

End Function

