Attribute VB_Name = "mTRCS"
'//System constant
Public Const tCS_TRPrjName As String = "Ada Train"      'Project Name
Public Const tCS_TRDefUser As String = "System"     'Default user

'//Database constant
Public Const tCS_TRDbName As String = "AdaTrain"      'Default database name
Public Const tCS_TRDbUser As String = "sa"    'Default DB User
Public Const tCS_TRDbPwd As String = "P@ssw0rd"    'Default DB Password

'//Provider constant
Public Const tCS_TRPrvSQL As String = "SQLOLEDB.1"        'Default Provider Name SQL SERVER
Public Const tCS_TRDbPrvMSAC As String = "Microsoft.JET.OLEDB.4.0"        'Default Provider Name MS ACCESS

'//Template Constant
Public Const tCS_TRConSQL As String = "Provider={0};Data Source={1};User ID={2};Password={3};Initial Catalog={4};Persist Security Info=False"       'Default Connectionstring Template for MSSQL
Public Const tCS_TRConMSAC As String = "Provider={0};Data Source={1};User ID={2};Password={3};Persist Security Info=False"       'Default Connectionstring Template for MSACCESS

'//SQL constant
Public Const tCS_TRSQLCst = "SELECT * FROM TTRMCst"
Public Const tCS_TRSQLCstNew = "SELECT max(FTCstCode) as tmp FROM TTRMCst WHERE FTCstCode LIKE 'C-____'"
Public Const tCS_TRSQLPdt = "SELECT * FROM TTRMPdt"
Public Const tCS_TRSQLPdtGrp = "SELECT * FROM TTRMPdtGrp"
Public Const tCS_TRSQLPdtGrpNew = "SELECT max(FTPdtGrpCode) as tmp FROM TTRMPdtGrp WHERE FTPdtGrpCode LIKE '___'"
Public Const tCS_TRSQLSpn = "SELECT * FROM TTRMSpn"
Public Const tCS_TRSQLSpnNew = "SELECT max(FTSpnCode) as tmp FROM TTRMSpn WHERE FTSpnCode LIKE 'E-____'"
