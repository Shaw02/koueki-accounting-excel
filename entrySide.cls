VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "entrySide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'   勘定科目のレコード
'==============================================================================
Option Explicit

Public m_AccountCode       As Variant     ' 勘定科目コード
Public m_MajorAccount      As String      ' 大科目（VLookup）
Public m_MiddleAccount     As String      ' 中科目（VLookup）
Public m_SubAccountCode    As Variant     ' 補助科目コード
Public m_SubAccountName    As String      ' 補助科目名（VLookup）
Public m_Class             As Integer     ' 0:一般(含む空白) / 1:指定
Public m_Amount            As Currency    ' 金額

'=========================
' Property
'=========================
Public Property Get AccountCode() As Variant
    AccountCode = m_AccountCode
End Property

Public Property Let AccountCode(ByVal value As Variant)
    m_AccountCode = value
End Property


Public Property Get MajorAccount() As String
    MajorAccount = m_MajorAccount
End Property

Public Property Let MajorAccount(ByVal value As String)
    m_MajorAccount = value
End Property


Public Property Get MiddleAccount() As String
    MiddleAccount = m_MiddleAccount
End Property

Public Property Let MiddleAccount(ByVal value As String)
    m_MiddleAccount = value
End Property


Public Property Get SubAccountCode() As Variant
    SubAccountCode = m_SubAccountCode
End Property

Public Property Let SubAccountCode(ByVal value As Variant)
    m_SubAccountCode = value
End Property


Public Property Get SubAccountName() As String
    SubAccountName = m_SubAccountName
End Property

Public Property Let SubAccountName(ByVal value As String)
    m_SubAccountName = value
End Property


Public Property Get Class() As Integer
    Class = m_Class
End Property

Public Property Let Class(ByVal value As Integer)
    m_Class = value
End Property


Public Property Get Amount() As Currency
    Amount = m_Amount
End Property

Public Property Let Amount(ByVal value As Currency)
    m_Amount = value
End Property

'==============================================================================
'   クラスの初期化
'==============================================================================
Private Sub Class_Initialize()
    
End Sub

'==============================================================================
'   クラスの解放
'==============================================================================
Private Sub Class_Terminate()

End Sub

'==============================================================================
'       実績読み込み
'------------------------------------------------------------------------------
'   Contents:
'       実績から、１行分、上のメンバー変数に読み込みます。
'
'   Input:
'       y       入力Ｙ座標
'
'   Output:
'       entrySide   仕訳帳 １行分
'
'==============================================================================
Public Sub ReadResults(y As Variant)

    AccountCode = Sheet10.Cells(y, 1)
    MajorAccount = Sheet10.Cells(y, 2)
    MiddleAccount = ""
    SubAccountCode = Sheet10.Cells(y, 3)
    SubAccountName = Sheet10.Cells(y, 4)
    Class = Sheet10.Cells(y, 5)
    Amount = Sheet10.Cells(y, 6)
    
End Sub

