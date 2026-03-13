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
Private m_Account           As Account
Private m_SubAccountCode    As Variant     ' 補助科目コード
Private m_SubAccountCate    As String      ' 会計区分
Private m_SubAccountName    As String      ' 補助科目名（VLookup）
Private m_NetAssetType      As NetAssetType
Private m_Amount            As Currency    ' 金額
Private m_side              As String

'==============================================================================
'   Property
'==============================================================================
Public Property Get AccountCode() As Variant
    AccountCode = m_Account.code
End Property

Public Property Get Account() As Account
    Set Account = m_Account
End Property

Public Property Get SubAccountCode() As Variant
    SubAccountCode = m_SubAccountCode
End Property

Public Property Let SubAccountCode(ByVal value As Variant)
    m_SubAccountCode = value
End Property


Public Property Get SubAccountCate() As Variant
    SubAccountCate = m_SubAccountCate
End Property

Public Property Let SubAccountCate(ByVal value As Variant)

    Dim strValue As String
    strValue = Trim(CStr(value))
    
    Select Case strValue
        Case "", " ", "　", "公共", "公益共通"
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
        Case "共", "共通"
            m_SubAccountCate = "共通"
        Case Else
            Err.Raise vbObjectError + 3003, _
                      "entrySide.Class", _
                      "会計区分が不正です: " & strValue
    End Select

End Property

Public Property Get SubAccountName() As String
    SubAccountName = m_SubAccountName
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
'   メンバー変数への代入
'==============================================================================
Public Sub Initialize( _
    ByVal acc As Account, _
    ByVal SubAccountCode As Variant, _
    ByVal SubAccountCate As Variant, _
    ByVal SubAccountName As String, _
    ByVal Class As Variant, _
    ByVal amount As Currency, _
    ByVal side As String)

    '------------------------------
    ' 代入
    '------------------------------
    Set m_Account = acc
    Me.SubAccountCode = SubAccountCode
    Me.SubAccountCate = SubAccountCate
    Me.SubAccountName = SubAccountName
    Me.Class = Class
    Me.amount = amount
    m_side = side
    
    '------------------------------
    ' 補助科目チェック
    '------------------------------
    '資産じゃない場合、且つ、補助科目が未入力でないか、チェック
    If (acc.code >= 20000) And (IsNull(SubAccountCode) Or IsEmpty(SubAccountCode)) Then
        Err.Raise vbObjectError + 3001, "entrySide:Initialize", m_side & ":補助科目が必要です"
    End If
    If (m_NetAssetType = natDesignated) And (IsNull(SubAccountCode) Or IsEmpty(SubAccountCode)) Then
        Err.Raise vbObjectError + 3002, "entrySide:Initialize", m_side & ":指定の場合、補助科目が必要です"
    End If

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
Public Sub ReadResults(y As Variant, master As AccountMaster)

    '----------------------------------
    '   変数宣言
    '----------------------------------
    Dim ws As Worksheet
    Dim str As Variant
    
    Set ws = ThisWorkbook.Worksheets("前期実績")

    str = ws.Cells(y, 1)
    If IsNull(str) Or IsEmpty(str) Then
        Set m_Account = master.GetAccount(-1)
        Me.SubAccountCode = 0
        Me.SubAccountCate = ""
        Me.SubAccountName = ""
        Me.Class = ""
        Me.amount = 0
    Else
       If Not master.Exists(str) Then
            Err.Raise vbObjectError + 3001, _
            "entrySide.class", _
            m_side & ":勘定科目コード" & str & "は入力できません。"
        Else
           Set m_Account = master.GetAccount(str)
            Me.SubAccountCode = ws.Cells(y, 4)
            Me.SubAccountCate = ws.Cells(y, 5)
            Me.SubAccountName = ws.Cells(y, 6)
            Me.Class = ws.Cells(y, 7)
            Me.amount = ws.Cells(y, 11)
        End If
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
        m_side & ":勘定科目コード" & m_Account.code & "は入力できません。"

    End If

End Sub




