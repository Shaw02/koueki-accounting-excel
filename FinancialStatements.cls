VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FinancialStatements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'   財務諸表
'==============================================================================
Option Explicit

'財務諸表（貸借対照表、正味財産増減計算書）の座標
Private Const styFinancialStatements = 6 '財務諸表の開始位置


'==============================================================================
'   クラスの初期化
'==============================================================================
Private Sub Class_Initialize()
    Call SheetClear          '総勘定元帳・今期実績のクリア
End Sub

'==============================================================================
'   クラスの解放
'==============================================================================
Private Sub Class_Terminate()

End Sub

'==============================================================================
'       財務諸表の０初期化
'==============================================================================
Sub SheetClear()

    '----------------------------------
    '   変数宣言
    '----------------------------------
    Dim wsBS As Worksheet
    Dim wsPL As Worksheet
    
    '----------------------------------
    '   初期化
    '----------------------------------
    Set wsBS = ThisWorkbook.Worksheets("④貸借対照表")
    Set wsPL = ThisWorkbook.Worksheets("⑤正味財産増減計算書")
        
    Dim yOutBS As Long
    Dim yOutPL As Long

    '貸借対照表へ
    yOutBS = styFinancialStatements
    Do
        If IsNull(wsBS.Cells(yOutBS, 1)) Or IsEmpty(wsBS.Cells(yOutBS, 1)) Then
        Else
            If wsBS.Cells(yOutBS, 1) = -1 Then
                Exit Do
            Else
                wsBS.Cells(yOutBS, 6) = 0
                wsBS.Cells(yOutBS, 7) = 0
            End If
        End If
        yOutBS = yOutBS + 1
    Loop
    
    '正味財産増減計算書へ
    yOutPL = styFinancialStatements
    Do
        If IsNull(wsPL.Cells(yOutPL, 1)) Or IsEmpty(wsPL.Cells(yOutPL, 1)) Then
        Else
            If wsPL.Cells(yOutPL, 1) = -1 Then
                Exit Do
            Else
                wsPL.Cells(yOutPL, 6) = 0
                wsPL.Cells(yOutPL, 7) = 0
            End If
        End If
        yOutPL = yOutPL + 1
    Loop

End Sub

'==============================================================================
'       財務諸表へ出力
'------------------------------------------------------------------------------
'   Contents:
'       正味財産増減計算書、若しくは貸借対照表の
'       指定の勘定科目コードの欄に、金額を出力します。
'
'   Input:
'       id      勘定科目コード
'       x       出力Ｘ座標
'       iMoney  金額
'
'   Output:
'       Sheet "正味財産増減計算書"
'       Sheet "貸借対照表"
'
'==============================================================================
Sub OutFinancialStatements(idAccount As Variant, iColumn As Variant, iMoney As Currency)

    '----------------------------------
    '   変数宣言
    '----------------------------------
    Dim wsBS As Worksheet
    Dim wsPL As Worksheet
    
    Dim yOutBS As Long
    Dim yOutPL As Long
    Dim flagErr As Boolean

    '----------------------------------
    '   初期化
    '----------------------------------
    Set wsBS = ThisWorkbook.Worksheets("④貸借対照表")
    Set wsPL = ThisWorkbook.Worksheets("⑤正味財産増減計算書")
    
    '財務諸表への出力
    If (idAccount < 30000) Then
        '貸借対照表へ
        yOutBS = 6
        flagErr = True
        Do
            If IsNull(wsBS.Cells(yOutBS, 1)) Or IsEmpty(wsBS.Cells(yOutBS, 1)) Then
            Else
                If wsBS.Cells(yOutBS, 1) = -1 Then
                    Exit Do
                ElseIf wsBS.Cells(yOutBS, 1) = idAccount Then
                    If idAccount >= 10000 And idAccount < 20000 Then
                        '借方
                        wsBS.Cells(yOutBS, iColumn) = iMoney
                    Else
                        '貸方
                        wsBS.Cells(yOutBS, iColumn) = -iMoney
                    End If
                    flagErr = False
                End If
            End If
            yOutBS = yOutBS + 1
        Loop
        If flagErr = True Then
            Err.Raise vbObjectError + 6101, _
                "FinancialStatements.OutFinancialStatements", _
                "貸借対照表に勘定科目が存在しません。勘定科目コード：" & CStr(idAccount)
        End If
        
    Else
        '正味財産増減計算書へ
        yOutPL = 6
        flagErr = True
        Do
            If IsNull(wsPL.Cells(yOutPL, 1)) Or IsEmpty(wsPL.Cells(yOutPL, 1)) Then
            Else
                If wsPL.Cells(yOutPL, 1) = -1 Then
                    Exit Do
                ElseIf wsPL.Cells(yOutPL, 1) = idAccount Then
                    If idAccount >= 30000 And idAccount < 40000 Then
                        '集計結果
                        wsPL.Cells(yOutPL, iColumn) = iMoney
                    ElseIf idAccount >= 40000 And idAccount < 50000 Then
                        '経常収益
                        wsPL.Cells(yOutPL, iColumn) = -iMoney
                    ElseIf idAccount >= 50000 And idAccount < 60000 Then
                        '経常費用
                        wsPL.Cells(yOutPL, iColumn) = iMoney
                    ElseIf idAccount >= 60000 And idAccount < 70000 Then
                        '経常外収益
                        wsPL.Cells(yOutPL, iColumn) = -iMoney
                    ElseIf idAccount >= 70000 And idAccount < 80000 Then
                        '経常外費用
                        wsPL.Cells(yOutPL, iColumn) = iMoney
                    End If
                    flagErr = False
                End If
            End If
            yOutPL = yOutPL + 1
        Loop
        If flagErr = True Then
            Err.Raise vbObjectError + 6102, _
                "FinancialStatements.OutFinancialStatements", _
                "正味財産増減計算書に勘定科目が存在しません。勘定科目コード：" & CStr(idAccount)
        End If

    End If

End Sub

