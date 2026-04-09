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
Private m_Code          As Long                 '勘定科目コード
Private m_LargeClass    As String               '勘定科目　大区分
Private m_SmallClass    As String               '勘定科目　小区分
Private m_MajorAccount  As String               '勘定科目　大科目
Private m_MiddleAccount As String               '勘定科目　中科目
Private m_FS            As String               'FS区分 （BS/PL）
Private m_FSEnum        As AccountFSKind        '
Private m_Category      As String               '区分   （資産／負債／純資産／収益／費用）
Private m_CategoryEnum  As AccountCategoryKind  '
Private m_SubCategory   As String               '小区分（流動資産／固定資産／流動負債／固定負債／経常収益／経常外収益／経常費用／経常外費用）
Private m_Postable      As Boolean              '仕訳帳 （入力用／集計用）
Private m_SubAccount    As Boolean              '補助科目は必須？
Private m_FSLineIDs     As String               '財務諸表ID

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

Public Property Get FSEnum() As AccountFSKind
    FSEnum = m_FSEnum
End Property

Public Property Get category() As String
    category = m_Category
End Property

Public Property Get CategoryEnum() As AccountCategoryKind
    CategoryEnum = m_CategoryEnum
End Property

Public Property Get SubCategory() As String
    SubCategory = m_SubCategory
End Property

Public Property Get postable() As Boolean
    postable = m_Postable
End Property

Public Property Get subAccount() As Boolean
    subAccount = m_SubAccount
End Property

Public Property Get fsLineIDs() As String
    fsLineIDs = m_FSLineIDs
End Property
'------------------------------------------------------------------------------
'   FS区分  （BS/PL）
'------------------------------------------------------------------------------
Private Function ParseFS(ByVal v As String) As AccountFSKind

    Select Case UCase$(Trim$(v))
        Case "BS", "B/S"
            ParseFS = afsBS
        Case "PL", "P/L"
            ParseFS = afsPL
        Case ""
            ParseFS = afsUnknown
        Case Else
            Err.Raise vbObjectError + 1001, _
                "Account.Class", _
                "FS区分が不正です: " & v
    End Select

End Function

'------------------------------------------------------------------------------
'   区分   （資産／負債／純資産／収益／費用）
'------------------------------------------------------------------------------
Private Function ParseCategory(ByVal v As String) As AccountCategoryKind

    Select Case Trim$(v)
        Case "資産"
            ParseCategory = ackAsset
            
        Case "負債"
            ParseCategory = ackLiability
            
        Case "純資産", "指定正味財産"
            ParseCategory = ackNetAsset
            
        Case "収益"
            ParseCategory = ackRevenue
            
        Case "費用"
            ParseCategory = ackExpense
            
        Case ""
            ParseCategory = ackUnknown
        
        Case Else
            Err.Raise vbObjectError + 1001, _
                "Account.Class", _
                "勘定科目区分が不正です: " & v
    End Select

End Function

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
    ByVal category As String, _
    ByVal SubCategory As String, _
    ByVal postable As String, _
    ByVal subAccount As String, _
    Optional ByVal fsLineIDs As String = "")

    m_Code = code
    m_LargeClass = largeClass
    m_SmallClass = smallClass
    m_MajorAccount = major
    m_MiddleAccount = middle
    m_FS = FS
    m_FSEnum = ParseFS(FS)
    m_Category = category
    m_CategoryEnum = ParseCategory(category)
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

    Select Case Trim(subAccount)
        Case "不要", ""
            m_SubAccount = False
        Case "必要", "要"
            m_SubAccount = True
        Case Else
            Err.Raise vbObjectError + 1001, _
                "Account.Class", _
                "勘定科目マスターの入力区分が不正です: " & postable
    End Select
    
    m_FSLineIDs = Trim$(fsLineIDs)

End Sub

