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
    Set wsMaster = ThisWorkbook.Worksheets("財務諸表マスタ")
    
    Call m_Master.LoadFromWorksheet(wsMaster)

    Dim thisAmounts As Object
    Dim lastAmounts As Object

    Set thisAmounts = BuildAccountAmountMap(tbEngine.SubTrialBalance)
    Set lastAmounts = BuildAccountAmountMap(tbEngine.LastSubTrialBalance)

    Call AddNetAssetAmounts(thisAmounts, tbEngine, False)
    Call AddNetAssetAmounts(lastAmounts, tbEngine, True)
    
    Call OutputAll(thisAmounts, lastAmounts)

End Sub

Private Function BuildAccountAmountMap(ByVal tb As TrialBalance) As Object

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim line As TrialBalanceLine
    Dim lineIDs As Variant
    Dim lineID As String
    Dim accountingCategory As String
    Dim amount As Currency
    Dim i As Long

    For Each line In tb.Lines
        amount = line.amount()

        If amount = 0 Then GoTo ContinueLoop

        If Trim$(line.account.fsLineIDs) <> "" Then
            lineIDs = Split(line.account.fsLineIDs, ";")
            For i = LBound(lineIDs) To UBound(lineIDs)
                lineID = UCase$(Trim$(CStr(lineIDs(i))))
                If lineID <> "" Then
                    accountingCategory = UCase$(Trim$(line.SubAccountCate))
                    Call AddAmountToMap(result, lineID, "*", amount)
                    If accountingCategory <> "" Then
                        Call AddAmountToMap(result, lineID, accountingCategory, amount)
                    End If
                End If
            Next i
        End If

ContinueLoop:
    Next

    Set BuildAccountAmountMap = result

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

    Const START_ROW As Long = 5
    
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

    Dim nextRowBySheet As Object
    Set nextRowBySheet = CreateObject("Scripting.Dictionary")

    nextRowBySheet(wsBS.Name) = START_ROW + 1
    nextRowBySheet(wsPL.Name) = START_ROW + 1
    nextRowBySheet(wsPL2.Name) = START_ROW + 1

    Call ClearSheet(wsBS, START_ROW)
    Call ClearSheet(wsPL, START_ROW)
    Call ClearSheet(wsPL2, START_ROW)

    '==============================
    ' 数値書式
    '==============================
    
    '貸借対照表
    With wsBS
        .Range("C:D").NumberFormat = "_ * #,##0_ ;_ * △#,##0_ ;_ * ""-""??_ ;_ @_ "
        .Range("C:D").HorizontalAlignment = xlRight
        
        .Cells(START_ROW, 2) = "科目"
        .Cells(START_ROW, 3) = "当期"
        .Cells(START_ROW, 4) = "前期"
        With .Range(.Cells(START_ROW, 2), .Cells(START_ROW, 4))
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
    
    '活動計算書
    With wsPL
        .Range("C:D").NumberFormat = "_ * #,##0_ ;_ * △#,##0_ ;_ * ""-""??_ ;_ @_ "
        .Range("C:D").HorizontalAlignment = xlRight
        
        .Cells(START_ROW, 2) = "科目"
        .Cells(START_ROW, 3) = "当期"
        .Cells(START_ROW, 4) = "前期"
        With .Range(.Cells(START_ROW, 2), .Cells(START_ROW, 4))
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
    
    
    '活動計算書の注記
    wsPL2.Range("C:D").NumberFormat = "_ * #,##0_ ;_ * △#,##0_ ;_ * ""-""??_ ;_ @_ "
    wsPL2.Range("C:D").HorizontalAlignment = xlRight

    For Each line In m_Master.Lines
    
        If line.enabled Then
    
            vThis = ResolveLineAmount(line, thisAmounts, valuesThis)
            vLast = ResolveLineAmount(line, lastAmounts, valuesLast)
    
            '------------------------------------------
            ' キー生成（重要）
            '------------------------------------------
            Dim key As String
            key = BuildKey(line.lineID, line.accountingCategory)
    
            valuesThis(key) = vThis
            valuesLast(key) = vLast
    
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
            If Not ((line.LineType = "ACCOUNT" Or line.LineType = "ACCOUNTS") And vThis = 0 And vLast = 0) Then
                If line.LineType = "TITLE" Then
                    ws.Cells(y, 3).value = ""
                    ws.Cells(y, 4).value = ""
                Else
                    ws.Cells(y, 3).value = vThis
                    ws.Cells(y, 4).value = vLast
                End If
                                    
                '横罫線のみ（帳簿感）
                ws.Range(ws.Cells(y, 2), ws.Cells(y, 4)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
                'ゼブラ
                If y Mod 2 = 0 Then
                    ws.Range(ws.Cells(y, 2), ws.Cells(y, 4)).Interior.Color = rgbAliceBlue
                End If

                nextRowBySheet(ws.Name) = y + 1
            End If
        End If
    Next


End Sub

Private Function ResolveLineAmount( _
    ByVal line As FinancialStatementMasterLine, _
    ByVal accountMap As Object, _
    ByVal resolvedValues As Object) As Currency

    Select Case line.LineType
        Case "ACCOUNT", "ACCOUNTS"
            ResolveLineAmount = ResolveAccountAmount(line, accountMap)
        Case "FORMULA", "CALC"
            ResolveLineAmount = ResolveFormula(line.formula, resolvedValues)
        Case Else
            ResolveLineAmount = 0
    End Select

End Function

'==============================================================================
'       勘定金額取得（マスタ準拠版）
'------------------------------------------------------------------------------
'   Contents:
'       ・マスタに存在しないカテゴリ → エラー
'       ・マスタに存在するが金額無し → 0円
'
'==============================================================================
Private Function ResolveAccountAmount( _
    ByVal def As FinancialStatementMasterLine, _
    ByVal accountMap As Object) As Currency

    Dim lineID As String
    Dim category As String
    Dim key As String

    lineID = UCase$(Trim$(def.lineID))
    category = UCase$(Trim$(def.accountingCategory))

    If category = "" Then category = "*"

    key = lineID & "|" & category

    '------------------------------------------
    ' ① マスタに存在するかチェック（最重要）
    '------------------------------------------
    If Not m_Master.Exists(lineID, category) Then
        
        Err.Raise vbObjectError + 2003, _
                  "ResolveAccountAmount", _
                  "未定義の会計区分です: lineID=" & lineID & _
                  ", category=" & category
        
    End If

    '------------------------------------------
    ' ② 金額取得（無ければ0）
    '------------------------------------------
    If accountMap.Exists(key) Then
        ResolveAccountAmount = CCur(accountMap(key))
    Else
        ResolveAccountAmount = 0
    End If

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
            lineID = Split(token, "|")(0)
            category = Split(token, "|")(1)
        Else
            lineID = token
            category = "*"
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

