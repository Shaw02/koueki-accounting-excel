VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FinancialStatementMasterLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_StatementID As String
Private m_TargetSheet As String
Private m_LineID As String
Private m_AccountingCategory As String
Private m_LineLabel As String
Private m_LineLabelAligh As XlHAlign
Private m_LineType As String
Private m_Formula As String
Private m_Enabled As Boolean

Public Sub Initialize( _
    ByVal StatementID As String, _
    ByVal TargetSheet As String, _
    ByVal lineID As String, _
    ByVal accountingCategory As String, _
    ByVal LineLabel As String, _
    ByVal LineLabelAligh As XlHAlign, _
    ByVal LineType As String, _
    ByVal formula As String, _
    ByVal enabled As Boolean)

    m_StatementID = Trim$(StatementID)
    m_TargetSheet = Trim$(TargetSheet)
    m_LineID = Trim$(lineID)
    
    m_AccountingCategory = UCase$(Trim$(accountingCategory))
    If m_AccountingCategory = "" Then
        m_AccountingCategory = "*"
    End If

    m_LineLabel = LineLabel
    m_LineLabelAligh = LineLabelAligh
    m_LineType = UCase$(Trim$(LineType))
    m_Formula = Trim$(formula)
    m_Enabled = enabled

End Sub

Public Property Get StatementID() As String: StatementID = m_StatementID: End Property
Public Property Get TargetSheet() As String: TargetSheet = m_TargetSheet: End Property
Public Property Get lineID() As String: lineID = m_LineID: End Property
Public Property Get accountingCategory() As String: accountingCategory = m_AccountingCategory: End Property
Public Property Get LineLabel() As String: LineLabel = m_LineLabel: End Property
Public Property Get LineLabelAligh() As XlHAlign: LineLabelAligh = m_LineLabelAligh: End Property
Public Property Get LineType() As String: LineType = m_LineType: End Property
Public Property Get formula() As String: formula = m_Formula: End Property
Public Property Get enabled() As Boolean: enabled = m_Enabled: End Property


