Attribute VB_Name = "mTREN"
'Type of database platform usage
Public Enum EN_TRDatabasePlatform
    ACCESS = 0
    SQLServer = 1
End Enum

'Type of database action
Public Enum EN_TRDatabaseAction
    Insert = 0
    Update = 1
    Delete = 2
End Enum

'Type of data input
Public Enum EN_TRDataType
    Text = 0
    Number = 1
    Float = 2
    Date = 3
    Bool = 4
End Enum

'Type of default Language
Public Enum EN_TRLanguage
    English = 0
    Thai = 1
End Enum

'Type of Dialog Message
Public Enum EN_TRMessageType
    Information = 0
    Exclamation = 1
    Question = 2
    Critical = 3
    Confirmation = 4
End Enum
