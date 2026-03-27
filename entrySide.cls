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

'==============================================================================
'   メンバー変数
'==============================================================================
Private m_Account           As account      ' 勘定科目
Private m_IsSubAcc          As Boolean      ' 補助科目有り？
Private m_SubAccountCode    As Long         ' 補助科目コード
Private m_SubAccountCate    As String       ' 会計区分（公1/公2/公共/収/法/etc..）
Private m_SubAccountName    As String       ' 補助科目名（VLookup）
Private m_NetAssetType      As NetAssetType ' 一般／指定
Private m_Amount            As Currency     ' 金額
Private m_Side              As Boolean      ' True=Debit False=Credit
Private m_Row               As Long         'この明細の行
Private m_ParentRow         As Long         '仕訳先頭行（任意）

'==============================================================================
'   Property
'==============================================================================
Public Property Get AccountCode() As Variant
    AccountCode = m_Account.code
End Property

Public Property Get account() As account
    Set account = m_Account
End Property

Public Property Get SubAccountCode() As Variant
    SubAccountCode = m_SubAccountCode
End Property

Public Property Get subCode() As Long
    subCode = m_SubAccountCode
End Property

Public Property Let SubAccountCode(ByVal value As Variant)

    If IsNull(value) Or IsEmpty(value) Then
        m_IsSubAcc = False
        m_SubAccountCode = 0
        Exit Property
    End If

    Dim strValue As String
    strValue = Trim(CStr(value))

    If strValue = "" Then
        m_IsSubAcc = False
        m_SubAccountCode = 0
        Exit Property
    End If

    If Not IsNumeric(strValue) Then
        Err.Raise vbObjectError + 3003, _
                  "entrySide.Class", _
                  "SubAccountCode is not numeric: " & strValue
    End If

    Dim v As Long
    v = CLng(strValue)

    If v < 0 Then
        Err.Raise vbObjectError + 3003, _
                  "entrySide.Class", _
                  "SubAccountCode is minus : " & strValue
    Else
        m_IsSubAcc = True
        m_SubAccountCode = v
    End If

End Property

Public Property Get side() As String
    Select Case m_Side
        Case True
            side = "借方"
        Case False
            side = "貸方"
        Case Else
            Err.Raise vbObjectError + 3003, _
              "entrySide.Class", _
              "値が不正です。" & m_Side
    End Select
End Property

Public Property Let side(ByVal value As String)
    Select Case value
        Case True, "借方"
            m_Side = True
        Case False, "貸方"
            m_Side = False
        Case Else
            Err.Raise vbObjectError + 3003, _
              "entrySide.Class", _
              "値が不正です。" & value
    End Select
End Property

Public Property Get SubAccountCate() As Variant
    SubAccountCate = m_SubAccountCate
End Property

Public Property Let SubAccountCate(ByVal value As Variant)

    Dim strValue As String
    strValue = Trim(CStr(value))
    
    Select Case strValue
        Case "", " ", "　", "共", "共通"
            m_SubAccountCate = "共通"
        Case "公共", "公益共通"
            m_SubAccountCate = "公共"
        Case "公1", "公益1"
            m_SubAccountCate = "公1"
        Case "公2", "公益2"
            m_SubAccountCate = "公2"
        Case "公3", "公益3"
            m_SubAccountCate = "公3"
        Case "収共", "収益共通"
            m_SubAccountCate = "収共"
        Case "収1", "収益1"
            m_SubAccountCate = "収1"
        Case "収2", "収益2"
            m_SubAccountCate = "収2"
        Case "収3", "収益3"
            m_SubAccountCate = "収3"
        Case "法", "法人"
            m_SubAccountCate = "法"
        Case Else
            Err.Raise vbObjectError + 3003, _
                      "entrySide.Class", _
                      "会計区分が不正です: " & strValue
    End Select

End Property

Public Property Get SubAccountName() As String
    If m_SubAccountCode = 0 Then
        SubAccountName = "共通"
    Else
        SubAccountName = m_SubAccountName
    End If
End Property

Public Property Let SubAccountName(ByVal value As String)
    m_SubAccountName = value
End Property

Public Property Get Class() As Variant
    Class = m_NetAssetType
End Property

Public Property Let Class(ByVal value As Variant)

    Dim strValue As String
    strValue = Trim(CStr(value))
    
    Select Case strValue
        Case natGeneral, "", " ", "　", "一般"
            m_NetAssetType = natGeneral
            
        Case natDesignated, "指定"
            m_NetAssetType = natDesignated
            
        Case Else
            Err.Raise vbObjectError + 3003, _
                      "entrySide.Class", _
                      "一般指定区分が不正です: " & strValue
    End Select

End Property

Public Property Get amount() As Currency
    amount = m_Amount
End Property

Public Property Let amount(ByVal value As Currency)
    m_Amount = value
End Property

Public Property Get Row() As Long
    Row = m_Row
End Property

Public Property Get ParentRow() As Long
    ParentRow = m_ParentRow
End Property

Public Property Get IsSubAcc() As Boolean
    IsSubAcc = m_IsSubAcc
End Property

'==============================================================================
'   クラスの初期化
'==============================================================================
Private Sub Class_Initialize()
    m_NetAssetType = natGeneral
End Sub

'==============================================================================
'   クラスの解放
'==============================================================================
Private Sub Class_Terminate()

End Sub

'==============================================================================
'   初期化
'------------------------------------------------------------------------------
'   Input:
'       acc                 勘定科目
'       SubAccountCode      補助科目コード
'       SubAccountCate      会計区分
'       SubAccountName      補助科目名
'       Class               一般／指定
'       amount              金額
'       side                借方／貸方
'       Row                 仕訳帳の行
'       ParentRow           仕訳帳開始行
'==============================================================================
Public Sub Initialize( _
    ByVal acc As account, _
    ByVal SubAccountCode As Variant, _
    ByVal SubAccountCate As Variant, _
    ByVal SubAccountName As String, _
    ByVal Class As Variant, _
    ByVal amount As Currency, _
    ByVal side As String, _
    ByVal Row As Long, _
    ByVal ParentRow As Long)

    '------------------------------
    ' 代入
    '------------------------------
    Set m_Account = acc
    Me.SubAccountCode = SubAccountCode
    Me.SubAccountCate = SubAccountCate
    Me.SubAccountName = SubAccountName
    Me.Class = Class
    Me.amount = amount
    Me.side = side
    m_Row = Row
    m_ParentRow = ParentRow
    
    '------------------------------
    ' 補助科目チェック
    '------------------------------
    '資産じゃない場合、且つ、補助科目が未入力でないか、チェック
    If acc.subAccount And (Not m_IsSubAcc) Then
        Err.Raise vbObjectError + 3001, "entrySide:Initialize", Me.side & ":補助科目が必要です"
    End If
    If (m_NetAssetType = natDesignated) And (Not m_IsSubAcc) Then
        Err.Raise vbObjectError + 3002, "entrySide:Initialize", Me.side & ":指定の場合、補助科目が必要です"
    End If

End Sub

'==============================================================================
'   仕訳帳の読み込み
'==============================================================================
Public Sub CheckEntry()

    '入力可能チェック
    If Not m_Account.postable Then

        Err.Raise vbObjectError + 3002, _
        "Journal.CheckEntrySide", _
        Me.side & ":勘定科目コード" & m_Account.code & "は入力できません。"

    End If

End Sub

