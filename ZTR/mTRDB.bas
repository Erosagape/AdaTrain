Attribute VB_Name = "mTRDB"
Option Explicit
Sub Main()

    Call SP_SETxVariable(mTRCS.tCS_TRDefUser, ACCESS, English, True)
    
    If mTRVB.oVB_TRDbCon.State = adStateOpen Then
    
        wTRMain.Show
    
    Else
        
        Call SP_SHOWbMessage(mTRMS.tMS_0001, Critical)
        Set mTRVB.oVB_TRDbCon = Nothing
        
    End If
    
End Sub

