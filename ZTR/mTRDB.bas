Attribute VB_Name = "mTRDB"
Option Explicit
Sub Main()

    Call SP_SETxVariable(tCS_TRDefUser, ACCESS, English, True)
    
    If oVB_TRDbCon.State = adStateOpen Then
    
        wTRMain.Show
    
    Else
        
        Call SP_SHOWbMessage(tMS_0001, Critical)
        Set oVB_TRDbCon = Nothing
        
    End If
    
End Sub

