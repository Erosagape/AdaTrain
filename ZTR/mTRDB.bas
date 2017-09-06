Attribute VB_Name = "mTRDB"
Option Explicit
Sub Main()

    Call mTRSP.SP_SETxVariable(mTRCS.tCS_TRDefaultUser, ACCESS, English, True)
    
    If mTRVB.oVB_TRDatabaseConnection.State = adStateOpen Then
    
        wTRMain.Show
    
    Else
        
        Call mTRSP.SP_SHOWbMessage(mTRMS.tMS_TRConnectFAIL, Critical)
        Set mTRVB.oVB_TRDatabaseConnection = Nothing
        
    End If
    
End Sub

