Attribute VB_Name = "mTRSP"
Option Explicit
Public Sub SP_SETxVariable(ptUser As String, poDbType As EN_TRDatabasePlatform, poLang As EN_TRLanguage, pbTestMode As Boolean)

    mTRVB.oVB_TRCurrentUser = ptUser
    mTRVB.oVB_TRDatabaseType = poDbType
    mTRVB.oVB_TRCurrentLang = poLang
    
    If pbTestMode = True Then
        Set mTRVB.oVB_TRDatabaseConnection = mTRSP.SP_DAToGetConnTEST()
    End If
    
    mTRVB.oVB_TRDatabaseConnection.Open
    
End Sub
Public Function SP_DATtGetConnStr(ptUserID As String, ptPassword As String, ptServerName As String, ptDatabaseName As String) As String
    
    Dim tProvider As String
    Dim tUserID As String
    Dim tPassword As String
    Dim tServerName  As String
    Dim tConnStr As String
    Dim tDatabase As String
    
    tUserID = ptUserID
    tPassword = ptPassword
    tServerName = ptServerName
    tDatabase = ptDatabaseName
    
    If mTRVB.oVB_TRDatabaseType = EN_TRDatabasePlatform.SQLServer Then
    
        tProvider = mTRCS.tCS_TRDbProviderSQL
        
        tConnStr = mTRCS.tCS_TRConnectionSQLTemplate
        tConnStr = Replace(tConnStr, "{0}", tProvider)
        tConnStr = Replace(tConnStr, "{1}", tServerName)
        tConnStr = Replace(tConnStr, "{2}", tUserID)
        tConnStr = Replace(tConnStr, "{3}", tPassword)
        tConnStr = Replace(tConnStr, "{4}", tDatabase)
        
    End If
    
    If mTRVB.oVB_TRDatabaseType = EN_TRDatabasePlatform.ACCESS Then
        tProvider = mTRCS.tCS_TRDbProviderACCESS
        tServerName = App.Path & "\" & tDatabase & ".mdb"
        
        tConnStr = mTRCS.tCS_TRConnectionACCESSTemplate
        tConnStr = Replace(tConnStr, "{0}", tProvider)
        tConnStr = Replace(tConnStr, "{1}", tServerName)
        tConnStr = Replace(tConnStr, "{2}", "Admin")
        tConnStr = Replace(tConnStr, "{3}", "")
            
    End If
            
    SP_DATtGetConnStr = tConnStr
    
End Function
Public Function SP_DAToGetConnTEST() As ADODB.Connection

    Dim oDbConn As New ADODB.Connection
    
    oDbConn.ConnectionString = mTRSP.SP_DATtGetConnStr(mTRCS.tCS_TRDatabaseUser, mTRCS.tCS_TRDatabasePassword, ".", mTRCS.tCS_TRDatabaseName)
    oDbConn.CursorLocation = adUseClient
    
    Set SP_DAToGetConnTEST = oDbConn

End Function
Public Function SP_TBLoGetFromSQL(poDbConn As ADODB.Connection, ptSQL As String) As ADODB.Recordset

    Dim oTable As New ADODB.Recordset
    oTable.Open ptSQL, poDbConn, adOpenDynamic, adLockOptimistic
    
    Set SP_TBLoGetFromSQL = oTable
    
End Function
Public Function SP_GETtNewCustomer(poDbConn As ADODB.Connection) As String
    
    Dim oRs As ADODB.Recordset
    Set oRs = poDbConn.Execute(mTRCS.tCS_TRSQLCustonerNew)
    
    If oRs.EOF = False Then
        SP_GETtNewCustomer = "C-" & Format(CInt(Right(oRs.Fields(0).Value, 4)) + 1, "0000")
    Else
        SP_GETtNewCustomer = "C-0001"
    End If
    
    oRs.Close
    
End Function
Public Function SP_TBLoGetCustomer(poDbConn As ADODB.Connection) As ADODB.Recordset
   
    Set SP_TBLoGetCustomer = SP_TBLoGetFromSQL(poDbConn, mTRCS.tCS_TRSQLCustomer)
    
End Function
Public Sub SP_SQLxSetLogTBL(poAction As EN_TRDatabaseAction, ptTableName As String, ptWhere As String, poDbConn As ADODB.Connection)

    Dim tSql As String
    Dim tLogType As String
    
    Select Case poAction
    Case EN_TRDatabaseAction.Insert
        tLogType = "Ins"
    Case EN_TRDatabaseAction.Update
        tLogType = "Upd"
    Case Else
        tLogType = ""
    End Select
        
    If tLogType <> "" Then
    
        Dim tSQLDate As String
        Dim tSQLTime As String
        
        tSql = " Update " & ptTableName & " set "
        tSql = tSql & " FDDate" & tLogType & "={0}, "
        tSql = tSql & " FTTime" & tLogType & "={1}, "
        tSql = tSql & " FTWho" & tLogType & "='" & mTRVB.oVB_TRCurrentUser & "'"
        
        If tLogType = "Ins" Then
            tSql = tSql & " ,FDDateUpd={0}, "
            tSql = tSql & " FTTimeUpd={1}, "
            tSql = tSql & " FTWhoUpd='" & mTRVB.oVB_TRCurrentUser & "'"
        End If
        
        If mTRVB.oVB_TRDatabaseType = ACCESS Then
        
                tSql = Replace(tSql, "{0}", "Format(Now(),'yyyy-MM-dd')")
                tSql = Replace(tSql, "{1}", "Format(Now(),'HH:mm:ss')")
                
        End If
        
        If mTRVB.oVB_TRDatabaseType = SQLServer Then
        
                tSql = Replace(tSql, "{0}", "Convert(date,GetDate())")
                tSql = Replace(tSql, "{1}", "Convert(time,GetDate())")
                
        End If
                
        tSql = tSql & " WHERE " & ptWhere
        
        poDbConn.Execute tSql
    
    End If
End Sub

Public Function SP_SHOWbMessage(ptMsgCode As String, poMsgType As EN_TRMessageType) As Boolean
    
    If InStr(1, ptMsgCode, ";") = 0 Then ptMsgCode = ptMsgCode & ";" & ptMsgCode
    
    Dim oMsgStyle As VbMsgBoxStyle
    
    Select Case poMsgType
        Case EN_TRMessageType.Information
            oMsgStyle = vbInformation + vbOKOnly
        Case EN_TRMessageType.Exclamation
            oMsgStyle = vbExclamation + vbOKOnly
        Case EN_TRMessageType.Critical
            oMsgStyle = vbCritical + vbOKOnly
        Case EN_TRMessageType.Question
            oMsgStyle = vbQuestion + vbOKCancel
        Case EN_TRMessageType.Confirmation
            oMsgStyle = vbExclamation + vbOKCancel
    End Select
    
    Dim nLangIndex As Integer
    nLangIndex = 0
    
    If mTRVB.oVB_TRCurrentLang = Thai Then
        nLangIndex = 1
    End If
    
    If mTRVB.oVB_TRCurrentLang = English Then
        nLangIndex = 0
    End If
    
    Dim oMsgResult As VbMsgBoxResult
    oMsgResult = MsgBox(Split(ptMsgCode, ";")(nLangIndex), oMsgStyle, mTRCS.tCS_TRProjectName)
    
    SP_SHOWbMessage = IIf(oMsgResult = vbOK, True, False)
    
End Function

Public Sub SP_CTLxSetFocus(ByRef oCtl As Object)
    
    oCtl.SelStart = 0
    oCtl.SelLength = Len(oCtl.Text)

End Sub
Public Function SP_SQLbRunCommand(poDbConn As ADODB.Connection, ptSQLText As String) As Boolean
On Error GoTo Err:
    
    Dim oCmd As New ADODB.Command
    
    oCmd.ActiveConnection = poDbConn
    oCmd.CommandType = adCmdText
    oCmd.CommandText = ptSQLText
                        
    oCmd.Execute
    
    SP_SQLbRunCommand = True
    Exit Function
    
Err:
    SP_SQLbRunCommand = False
    
End Function
Public Function SP_SQLtFormatText(ptValue As String, poDataType As EN_TRDataType) As String
    On Error GoTo Err:
    Dim tStr As String
    tStr = ptValue
    
    If mTRVB.oVB_TRDatabaseType = ACCESS Then
           Select Case poDataType
           Case EN_TRDataType.Number
                    tStr = "CInt(" & ptValue & ")"
           Case EN_TRDataType.Float
                    tStr = "CDbl(" & ptValue & ")"
           Case EN_TRDataType.Date
                    tStr = "CDate('" & ptValue & "')"
           Case EN_TRDataType.Bool
                    tStr = "CBool(" & ptValue & ")"
           Case Else
                    tStr = "'" & Replace(tStr, "'", "''") & "'"
           End Select
    End If
    
    If mTRVB.oVB_TRDatabaseType = SQLServer Then
           Select Case poDataType
           Case EN_TRDataType.Number
                    tStr = ptValue
           Case EN_TRDataType.Date
                    tStr = "Convert(date,'" & Format(CDate(ptValue), "yyyy-MM-dd") & "')"
           Case Else
                    tStr = "'" & Replace(tStr, "'", "''") & "'"
           End Select
    End If
    
    SP_SQLtFormatText = tStr
    
    Exit Function
Err:
    SP_SQLtFormatText = "NULL"
End Function
