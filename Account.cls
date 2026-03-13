VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'   勘定科目
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
'   メンバー変数
'------------------------------------------------------------------------------
Private m_Code          As Long
Private m_LargeClass    As String
Private m_SmallClass    As String
Private m_MajorAccount  As String
Private m_MiddleAccount As String
Private m_FS            As String
Private m_Category      As String
Private m_SubCategory   As String
Private m_Postable      As Boolean

'------------------------------------------------------------------------------
'   Property
'------------------------------------------------------------------------------
Public Property Get code() As Long
    code = m_Code
End Property

Public Property Get largeClass() As String
    largeClass = m_LargeClass
End Property

Public Property Get smallClass() As String
    smallClass = m_SmallClass
End Property

Public Property Get MajorAccount() As String
    MajorAccount = m_MajorAccount
End Property

Public Property Get MiddleAccount() As String
    MiddleAccount = m_MiddleAccount
End Property

Public Property Get FS() As String
    FS = m_FS
End Property

Public Property Get Category() As String
    Category = m_Category
End Property

Public Property Get SubCategory() As String
    SubCategory = m_SubCategory
End Property

Public Property Get postable() As Boolean
    postable = m_Postable
End Property

'------------------------------------------------------------------------------
'   初期化
'------------------------------------------------------------------------------
Public Sub Initialize( _
    ByVal code As Long, _
    ByVal largeClass As String, _
    ByVal smallClass As String, _
    ByVal major As String, _
    ByVal middle As String, _
    ByVal FS As String, _
    ByVal Category As String, _
    ByVal SubCategory As String, _
    ByVal postable As String)

    m_Code = code
    m_LargeClass = largeClass
    m_SmallClass = smallClass
    m_MajorAccount = major
    m_MiddleAccount = middle
    m_FS = FS
    m_Category = Category
    m_SubCategory = SubCategory

    Select Case Trim(postable)
        Case "集計", ""
            m_Postable = False
        Case "入力"
            m_Postable = True
        Case Else
            Err.Raise vbObjectError + 1001, _
                "Account.Class", _
                "勘定科目マスターの入力区分が不正です: " & postable
    End Select

End Sub

