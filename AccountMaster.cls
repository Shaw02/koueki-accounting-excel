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
'   勘定科目マスター（Dictionary版）
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
'   メンバー変数
'------------------------------------------------------------------------------
Private m_dict As Object   ' Scripting.Dictionary

'------------------------------------------------------------------------------
'   初期化
'------------------------------------------------------------------------------
Private Sub Class_Initialize()

    Set m_dict = CreateObject("Scripting.Dictionary")

End Sub

'------------------------------------------------------------------------------
'   解放
'------------------------------------------------------------------------------
Private Sub Class_Terminate()

    Set m_dict = Nothing

End Sub

'------------------------------------------------------------------------------
'   追加
'------------------------------------------------------------------------------
'   Input:
'       acc     勘定科目
'
'   説明:
'       勘定科目をマスターに登録する。
'       同一コードが存在する場合はエラーとする。
'
'------------------------------------------------------------------------------
Public Sub AddAccount(ByVal acc As account)

    If acc Is Nothing Then
        Err.Raise vbObjectError + 2003, _
                  "AccountMaster.AddAccount", _
                  "勘定科目オブジェクトが Nothing です。"
    End If
    
    Dim key As String
    key = CStr(acc.code)

    If m_dict.Exists(key) Then
        Err.Raise vbObjectError + 2002, _
                  "AccountMaster.AddAccount", _
                  "勘定科目コードが重複しています: " & key
    End If

    m_dict.Add key, acc

End Sub

'------------------------------------------------------------------------------
'   取得
'------------------------------------------------------------------------------
'   Input:
'       code    勘定科目コード
'
'   Output:
'               勘定科目オブジェクト
'
'   説明:
'       指定したコードの勘定科目を取得する。
'       存在しない場合はエラーとする。
'
'------------------------------------------------------------------------------
Public Function GetAccount(ByVal code As Variant) As account

    Dim key As String
    key = CStr(code)

    If Not m_dict.Exists(key) Then
        Err.Raise vbObjectError + 2001, _
                  "AccountMaster.GetAccount", _
                  "勘定科目が存在しません: " & key
    End If

    Set GetAccount = m_dict(key)

End Function

'------------------------------------------------------------------------------
'   存在チェック
'------------------------------------------------------------------------------
'   Input:
'       code    勘定科目コード
'
'   Output:
'               True: 存在する
'               False: 存在しない
'
'------------------------------------------------------------------------------
Public Function Exists(ByVal code As Variant) As Boolean

    Exists = m_dict.Exists(CStr(code))

End Function

'------------------------------------------------------------------------------
'   件数
'------------------------------------------------------------------------------
'   Output:
'               登録されている勘定科目数
'
'------------------------------------------------------------------------------
Public Property Get Count() As Long

    Count = m_dict.Count

End Property

'------------------------------------------------------------------------------
'   キー一覧取得
'------------------------------------------------------------------------------
'   Output:
'               勘定科目コード配列（Variant配列）
'
'   説明:
'       全件ループ用にキー配列を返す。
'
'------------------------------------------------------------------------------
Public Property Get Keys() As Variant

    Keys = m_dict.Keys

End Property

'------------------------------------------------------------------------------
'   キー指定取得（ループ用）
'------------------------------------------------------------------------------
'   Input:
'       key     勘定科目コード
'
'   Output:
'               勘定科目オブジェクト
'
'------------------------------------------------------------------------------
Public Function GetByKey(ByVal key As Variant) As account

    Dim keyStr As String
    keyStr = CStr(key)
    
    If Not m_dict.Exists(keyStr) Then
        Err.Raise vbObjectError + 2001, _
                  "AccountMaster.GetByKey", _
                  "勘定科目が存在しません: " & keyStr
    End If

    Set GetByKey = m_dict(keyStr)

End Function

