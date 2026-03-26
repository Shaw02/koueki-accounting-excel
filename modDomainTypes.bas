Attribute VB_Name = "modDomainTypes"
Public Enum NetAssetType
    natGeneral = 0
    natDesignated = 1
End Enum

Public Enum AccountFSKind
    afsUnknown = 0
    afsBS = 1           '‘İØ‘ÎÆ•\
    afsPL = 2           'Šˆ“®ŒvZ‘
End Enum

Public Enum AccountCategoryKind
    ackUnknown = 0
    ackAsset = 1        '‘Y
    ackLiability = 2    '•‰Â
    ackNetAsset = 3     'ƒ‘Y
    ackRevenue = 4      'û‰v
    ackExpense = 5      '”ï—p
End Enum

