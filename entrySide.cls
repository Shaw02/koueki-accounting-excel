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
Private m_AccountCode       As Long        ' 勘定科目コード
Private m_MajorAccount      As String      ' 大科目（VLookup）
Private m_MiddleAccount     As String      ' 中科目（VLookup）
Private m_SubAccountCode    As Variant     ' 補助科目コード
Private m_SubAccountName    As String      ' 補助科目名（VLookup）
Private m_NetAssetType      As NetAssetType
Private m_Amount            As Currency    ' 金額

'==============================================================================
'   Property
'==============================================================================
Public Property Get AccountCode() As Variant
    AccountCode = m_AccountCode
End Property

Public Property Let AccountCode(ByVal value As Variant)

    '----------------------------------
    ' Null / Empty チェック
    '----------------------------------
    If IsNull(value) Or IsEmpty(value) Then
        Err.Raise vbObjectError + 1001, _
                  "entrySide.AccountCode", _
                  "勘定科目コードが未入力です。"
    End If
    
    '----------------------------------
    ' 文字列化して前後空白除去
    '----------------------------------
    Dim strCode As String
    strCode = Trim(CStr(value))
    
    If strCode = "" Then
        Err.Raise vbObjectError + 1002, _
                  "entrySide.AccountCode", _
                  "勘定科目コードが空白です。"
    End If
    
    '----------------------------------
    ' 数値のみチェック
    '----------------------------------
    If Not strCode Like "#####" Then
        Err.Raise vbObjectError + 1003, _
                  "entrySide.AccountCode", _
                  "勘定科目コードは5桁の数字で入力してください。"
    End If
    
    '----------------------------------
    ' 数値変換
    '----------------------------------
    m_AccountCode = CLng(strCode)

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
                      "区分が不正です: " & strValue
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
    ByVal AccountCode As Variant, _
    ByVal MajorAccount As String, _
    ByVal MiddleAccount As String, _
    ByVal SubAccountCode As Variant, _
    ByVal SubAccountName As String, _
    ByVal Class As Variant, _
    ByVal amount As Currency)

    '------------------------------
    ' 代入
    '------------------------------
    Me.AccountCode = AccountCode
    Me.MajorAccount = MajorAccount
    Me.MiddleAccount = MiddleAccount
    Me.SubAccountCode = SubAccountCode
    Me.SubAccountName = SubAccountName
    Me.Class = Class
    Me.amount = amount
    
    '------------------------------
    ' 補助科目チェック
    '------------------------------
    '資産じゃない場合、且つ、補助科目が未入力でないか、チェック
    If (AccountCode >= 20000) And (IsNull(SubAccountCode) Or IsEmpty(SubAccountCode)) Then
        Err.Raise vbObjectError + 3001, "entrySide:Initialize", "補助科目が必要です"
    End If
    If (m_NetAssetType = natDesignated) And (IsNull(SubAccountCode) Or IsEmpty(SubAccountCode)) Then
        Err.Raise vbObjectError + 3002, "entrySide:Initialize", "指定の場合、補助科目が必要です"
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
Public Sub ReadResults(y As Variant)

    '----------------------------------
    '   変数宣言
    '----------------------------------
    Dim ws As Worksheet
    Dim str As Variant
    
    Set ws = ThisWorkbook.Worksheets("前期実績")

    str = ws.Cells(y, 1)
    If IsNull(str) Or IsEmpty(str) Then
        Me.AccountCode = 99999
        Me.MajorAccount = ""
        Me.MiddleAccount = ""
        Me.SubAccountCode = 0
        Me.SubAccountName = ""
        Me.Class = ""
        Me.amount = 0
    Else
        Me.AccountCode = str
        Me.MajorAccount = ws.Cells(y, 2)
        Me.MiddleAccount = ws.Cells(y, 3)
        Me.SubAccountCode = ws.Cells(y, 4)
        Me.SubAccountName = ws.Cells(y, 5)
        Me.Class = ws.Cells(y, 6)
        Me.amount = ws.Cells(y, 7)
    End If
    
End Sub

