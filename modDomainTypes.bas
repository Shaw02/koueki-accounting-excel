Attribute VB_Name = "modDomainTypes"
Public Enum NetAssetType
    natGeneral = 0
    natDesignated = 1
End Enum

Public Enum AccountFSKind
    afsUnknown = 0
    afsBS = 1
    afsPL = 2
End Enum

Public Enum AccountCategoryKind
    ackUnknown = 0
    ackAsset = 1
    ackLiability = 2
    ackNetAsset = 3
    ackRevenue = 4
    ackExpense = 5
End Enum
