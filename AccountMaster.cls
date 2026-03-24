VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccountMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
'   勘定科目マスター
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
'   メンバー変数
'------------------------------------------------------------------------------
Private m_Accounts As Collection

'------------------------------------------------------------------------------
'   初期化
'------------------------------------------------------------------------------
Private Sub Class_Initialize()

    Set m_Accounts = New Collection

End Sub

'------------------------------------------------------------------------------
'   解放
'------------------------------------------------------------------------------
Private Sub Class_Terminate()

    Set m_Accounts = Nothing

End Sub

'------------------------------------------------------------------------------
'   追加
'------------------------------------------------------------------------------
Public Sub AddAccount(ByVal acc As account)

    m_Accounts.Add acc, CStr(acc.code)

End Sub

'------------------------------------------------------------------------------
'   取得
'------------------------------------------------------------------------------
Public Function GetAccount(ByVal code As Variant) As account

    Set GetAccount = m_Accounts.Item(CStr(code))

End Function

'------------------------------------------------------------------------------
'   存在チェック
'------------------------------------------------------------------------------
Public Function Exists(ByVal code As Variant) As Boolean

    On Error GoTo ErrHandle

    Dim tmp As account
    Set tmp = m_Accounts.Item(CStr(code))

    Exists = True
    Exit Function

ErrHandle:

    Exists = False

End Function

'------------------------------------------------------------------------------
'   件数
'------------------------------------------------------------------------------
Public Property Get Count() As Long

    Count = m_Accounts.Count

End Property
