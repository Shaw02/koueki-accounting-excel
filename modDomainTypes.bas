Attribute VB_Name = "modDomainTypes"
Public Enum NetAssetType
    natGeneral = 0
    natDesignated = 1
End Enum

Public Enum AccountFSKind
    afsUnknown = 0
    afsBS = 1           '貸借対照表
    afsPL = 2           '活動計算書
End Enum

Public Enum AccountCategoryKind
    ackUnknown = 0
    ackAsset = 1        '資産
    ackLiability = 2    '負債
    ackNetAsset = 3     '純資産
    ackRevenue = 4      '収益
    ackExpense = 5      '費用
End Enum

'財務諸表マスタ 対象
Public Const Terget_ID              As String = "ID"
Public Const Terget_Sub_Class       As String = "CLASS"
Public Const Terget_Acc_Code        As String = "CODE"

'財務諸表マスタ フィルター区分
Public Const CATEGORY_ACCOUNT_SUM   As String = "目的区分"
Public Const CATEGORY_ACCOUNT       As String = "会計区分"
Public Const CATEGORY_PROJECT       As String = "補助区分"
Public Const CATEGORY_NET_ASSET     As String = "純資産区分"

