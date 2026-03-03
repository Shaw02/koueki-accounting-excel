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

'=======================================================
'       財務諸表の０初期化
'-------------------------------------------------------
Sub SheetClear()

    Dim yOutBS As Long
    Dim yOutPL As Long

    '貸借対照表へ
    yOutBS = styFinancialStatements
    Do
        If IsNull(Sheet6.Cells(yOutBS, 1)) Or IsEmpty(Sheet6.Cells(yOutBS, 1)) Then
        Else
            If Sheet6.Cells(yOutBS, 1) = -1 Then
                Exit Do
            Else
                Sheet6.Cells(yOutBS, 6) = 0
                Sheet6.Cells(yOutBS, 7) = 0
            End If
        End If
        yOutBS = yOutBS + 1
    Loop
    
    '正味財産増減計算書へ
    yOutPL = styFinancialStatements
    Do
        If IsNull(Sheet5.Cells(yOutPL, 1)) Or IsEmpty(Sheet5.Cells(yOutPL, 1)) Then
        Else
            If Sheet5.Cells(yOutPL, 1) = -1 Then
                Exit Do
            Else
                Sheet5.Cells(yOutPL, 6) = 0
                Sheet5.Cells(yOutPL, 7) = 0
            End If
        End If
        yOutPL = yOutPL + 1
    Loop

End Sub

'=======================================================
'       財務諸表へ出力
'-------------------------------------------------------
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
'=======================================================
Sub OutFinancialStatements(idAccount As Variant, iColumn As Variant, iMoney As Currency)

    Dim yOutBS As Long
    Dim yOutPL As Long
    Dim flagErr As Boolean

    '財務諸表への出力
    If (idAccount < 30000) Then
        '貸借対照表へ
        yOutBS = 6
        flagErr = True
        Do
            If IsNull(Sheet6.Cells(yOutBS, 1)) Or IsEmpty(Sheet6.Cells(yOutBS, 1)) Then
            Else
                If Sheet6.Cells(yOutBS, 1) = -1 Then
                    Exit Do
                ElseIf Sheet6.Cells(yOutBS, 1) = idAccount Then
                    If idAccount >= 10000 And idAccount < 20000 Then
                        '借方
                        Sheet6.Cells(yOutBS, iColumn) = iMoney
                    Else
                        '貸方
                        Sheet6.Cells(yOutBS, iColumn) = -iMoney
                    End If
                    flagErr = False
                End If
            End If
            yOutBS = yOutBS + 1
        Loop
        If flagErr = True Then
            MsgBox "貸借対照表に勘定科目がありません。：勘定科目コード＝" & Str(idAccount)
            End
        End If
        
    Else
        '正味財産増減計算書へ
        yOutPL = 6
        flagErr = True
        Do
            If IsNull(Sheet5.Cells(yOutPL, 1)) Or IsEmpty(Sheet5.Cells(yOutPL, 1)) Then
            Else
                If Sheet5.Cells(yOutPL, 1) = -1 Then
                    Exit Do
                ElseIf Sheet5.Cells(yOutPL, 1) = idAccount Then
                    If idAccount >= 30000 And idAccount < 40000 Then
                        '集計結果
                        Sheet5.Cells(yOutPL, iColumn) = iMoney
                    ElseIf idAccount >= 40000 And idAccount < 50000 Then
                        '経常収益
                        Sheet5.Cells(yOutPL, iColumn) = -iMoney
                    ElseIf idAccount >= 50000 And idAccount < 60000 Then
                        '経常費用
                        Sheet5.Cells(yOutPL, iColumn) = iMoney
                    ElseIf idAccount >= 60000 And idAccount < 70000 Then
                        '経常外収益
                        Sheet5.Cells(yOutPL, iColumn) = -iMoney
                    ElseIf idAccount >= 70000 And idAccount < 80000 Then
                        '経常外費用
                        Sheet5.Cells(yOutPL, iColumn) = iMoney
                    End If
                    flagErr = False
                End If
            End If
            yOutPL = yOutPL + 1
        Loop
        If flagErr = True Then
            MsgBox "正味財産増減計算書に勘定科目がありません。：勘定科目コード＝" & Str(idAccount)
            End
        End If

    End If

End Sub

