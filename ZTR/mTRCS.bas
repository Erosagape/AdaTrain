Attribute VB_Name = "mTRCS"
'//System constant
Public Const tCS_TRProjectName As String = "Ada Train"      'Project Name
Public Const tCS_TRDefaultUser As String = "System"     'Default user

'//Database constant
Public Const tCS_TRDatabaseName As String = "AdaTrain"      'Default database name
Public Const tCS_TRDatabaseUser As String = "sa"    'Default DB User
Public Const tCS_TRDatabasePassword As String = "P@ssw0rd"    'Default DB Password

'//Provider constant
Public Const tCS_TRDbProviderSQL As String = "SQLOLEDB.1"        'Default Provider Name SQL SERVER
Public Const tCS_TRDbProviderACCESS As String = "Microsoft.JET.OLEDB.4.0"        'Default Provider Name MS ACCESS

'//Template Constant
Public Const tCS_TRConnectionSQLTemplate As String = "Provider={0};Data Source={1};User ID={2};Password={3};Initial Catalog={4};Persist Security Info=False"       'Default Connectionstring Template for MSSQL
Public Const tCS_TRConnectionACCESSTemplate As String = "Provider={0};Data Source={1};User ID={2};Password={3};Persist Security Info=False"       'Default Connectionstring Template for MSACCESS

'//SQL constant
Public Const tCS_TRSQLCustomer = "SELECT * FROM TTRMCst"
Public Const tCS_TRSQLCustonerNew = "SELECT max(FTCstCode) as tmp FROM TTRMCst WHERE FTCstCode LIKE 'C-____'"
