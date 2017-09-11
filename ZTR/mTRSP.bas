Attribute VB_Name = "mTRSP"
Option Explicit
Public Sub SP_SETxVariable(ptUser As String, poDbType As EN_TRDbType, poLang As EN_TRLang, pbTestMode As Boolean)

    mTRVB.tVB_TRUser = ptUser
    mTRVB.eVB_TRDbType = poDbType
    mTRVB.eVB_TRLang = poLang
    
    If pbTestMode = True Then
        Set mTRVB.oVB_TRDbCon = SP_DAToGetConnTEST()
    End If
    
    mTRVB.oVB_TRDbCon.Open
    
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
    
    If mTRVB.eVB_TRDbType = EN_TRDbType.SQLServer Then
    
        tProvider = mTRCS.tCS_TRPrvSQL
        
        tConnStr = mTRCS.tCS_TRConSQL
        tConnStr = Replace(tConnStr, "{0}", tProvider)
        tConnStr = Replace(tConnStr, "{1}", tServerName)
        tConnStr = Replace(tConnStr, "{2}", tUserID)
        tConnStr = Replace(tConnStr, "{3}", tPassword)
        tConnStr = Replace(tConnStr, "{4}", tDatabase)
        
    End If
    
    If mTRVB.eVB_TRDbType = EN_TRDbType.ACCESS Then
        tProvider = mTRCS.tCS_TRDbPrvMSAC
        tServerName = App.Path & "\" & tDatabase & ".mdb"
        
        tConnStr = mTRCS.tCS_TRConMSAC
        tConnStr = Replace(tConnStr, "{0}", tProvider)
        tConnStr = Replace(tConnStr, "{1}", tServerName)
        tConnStr = Replace(tConnStr, "{2}", "Admin")
        tConnStr = Replace(tConnStr, "{3}", "")
            
    End If
            
    SP_DATtGetConnStr = tConnStr
    
End Function
Public Function SP_DAToGetConnTEST() As ADODB.Connection

    Dim oDbConn As New ADODB.Connection
    
    oDbConn.ConnectionString = SP_DATtGetConnStr(mTRCS.tCS_TRDbUser, mTRCS.tCS_TRDbPwd, ".", mTRCS.tCS_TRDbName)
    oDbConn.CursorLocation = adUseClient
    
    Set SP_DAToGetConnTEST = oDbConn

End Function
Public Function SP_TBLoGetFROMSQL(poDbConn As ADODB.Connection, ptSQL As String) As ADODB.Recordset

    Dim oTable As New ADODB.Recordset
    oTable.Open ptSQL, poDbConn, adOpenDynamic, adLockOptimistic
    
    Set SP_TBLoGetFROMSQL = oTable
    
End Function
Public Function SP_GETtNewSalesPerson(poDbConn As ADODB.Connection) As String
    
    Dim oRs As ADODB.Recordset
    Set oRs = poDbConn.Execute(mTRCS.tCS_TRSQLSpnNew)
    
    If oRs.EOF = False Then
        SP_GETtNewSalesPerson = "E-" & Format(CInt(Right(oRs.Fields(0).Value, 4)) + 1, "0000")
    Else
        SP_GETtNewSalesPerson = "E-0001"
    End If
    
    oRs.Close
    
End Function
Public Function SP_GETtNewCustomer(poDbConn As ADODB.Connection) As String
    
    Dim oRs As ADODB.Recordset
    Set oRs = poDbConn.Execute(mTRCS.tCS_TRSQLCstNew)
    
    If oRs.EOF = False Then
        SP_GETtNewCustomer = "C-" & Format(CInt(Right(oRs.Fields(0).Value, 4)) + 1, "0000")
    Else
        SP_GETtNewCustomer = "C-0001"
    End If
    
    oRs.Close
    
End Function
Public Function SP_GETtNewProductGroup(poDbConn As ADODB.Connection) As String
    
    Dim oRs As ADODB.Recordset
    Set oRs = poDbConn.Execute(mTRCS.tCS_TRSQLPdtGrpNew)
    
    If oRs.EOF = False Then
        SP_GETtNewProductGroup = Format(CInt(Right(oRs.Fields(0).Value, 4)) + 1, "000")
    Else
        SP_GETtNewProductGroup = "001"
    End If
    
    oRs.Close
    
End Function

Public Function SP_DATtGetInput(poForm As Form, ptCtlType As String, ptName As String) As String
    '//special case field
    Select Case ptName
    Case "FDDateIns", "FDDateUpd"
        SP_DATtGetInput = SP_SQLtFormatText(Format(Now, "yyyy-MM-dd"), Date)
    Case "FTTimeIns", "FTTimeUpd"
        SP_DATtGetInput = SP_SQLtFormatText(Format(Now, "HH:mm:ss"), Text)
    Case "FTWhoIns", "FTWhoUpd"
        SP_DATtGetInput = SP_SQLtFormatText(mTRVB.tVB_TRUser, Text)
    Case Else
        '//find control name
            Dim oCtl As Control
            Set oCtl = poForm.Controls(ptCtlType & ptName)
            
            Dim oType As EN_TRDataType
            Dim tValue As String
            oType = Text
            tValue = ""
            '//default value
            Select Case Mid(ptName, 1, 2)
            Case "FD"
                tValue = "1900-01-01"
                oType = Date
            Case "FC"
                tValue = "0.00"
                oType = Float
            Case "FN"
                tValue = "0.00"
                oType = Number
            Case "FB"
                tValue = "0"
                oType = Bool
            End Select
            
            '//find value base on control type
            If Not oCtl Is Nothing Then
                Select Case ptCtlType
                Case "odt", "orb"
                    tValue = oCtl.Value
                Case "ocb"
                    tValue = oCtl.BoundText
                Case "ock"
                    tValue = oCtl.Checked
                Case Else
                    tValue = oCtl.Text
                End Select
            End If
            
            SP_DATtGetInput = SP_SQLtFormatText(tValue, oType)
    End Select
End Function
Public Function SP_TBLoGetCustomer(poDbConn As ADODB.Connection) As ADODB.Recordset
   
    Set SP_TBLoGetCustomer = SP_TBLoGetFROMSQL(poDbConn, mTRCS.tCS_TRSQLCst)
    
End Function
Public Function SP_TBLoGetProduct(poDbConn As ADODB.Connection) As ADODB.Recordset
   
    Set SP_TBLoGetProduct = SP_TBLoGetFROMSQL(poDbConn, mTRCS.tCS_TRSQLPdt)
    
End Function
Public Function SP_TBLoGetProductGroup(poDbConn As ADODB.Connection) As ADODB.Recordset
   
    Set SP_TBLoGetProductGroup = SP_TBLoGetFROMSQL(poDbConn, mTRCS.tCS_TRSQLPdtGrp)
    
End Function
Public Function SP_TBLoGetSalesPerson(poDbConn As ADODB.Connection) As ADODB.Recordset
   
    Set SP_TBLoGetSalesPerson = SP_TBLoGetFROMSQL(poDbConn, mTRCS.tCS_TRSQLSpn)
    
End Function
Public Sub SP_SQLxSetLogTBL(poAction As EN_TRDbAction, ptTableName As String, ptWhere As String, poDbConn As ADODB.Connection)

    Dim tSql As String
    Dim tLogType As String
    
    Select Case poAction
    Case EN_TRDbAction.Insert
        tLogType = "Ins"
    Case EN_TRDbAction.Update
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
        tSql = tSql & " FTWho" & tLogType & "='" & mTRVB.tVB_TRUser & "'"
        
        If tLogType = "Ins" Then
            tSql = tSql & " ,FDDateUpd={0}, "
            tSql = tSql & " FTTimeUpd={1}, "
            tSql = tSql & " FTWhoUpd='" & mTRVB.tVB_TRUser & "'"
        End If
        
        If mTRVB.eVB_TRDbType = ACCESS Then
        
                tSql = Replace(tSql, "{0}", "Format(Now(),'yyyy-MM-dd')")
                tSql = Replace(tSql, "{1}", "Format(Now(),'HH:mm:ss')")
                
        End If
        
        If mTRVB.eVB_TRDbType = SQLServer Then
        
                tSql = Replace(tSql, "{0}", "Convert(date,GetDate())")
                tSql = Replace(tSql, "{1}", "Convert(time,GetDate())")
                
        End If
                
        tSql = tSql & " WHERE " & ptWhere
        
        poDbConn.Execute tSql
    
    End If
End Sub

Public Function SP_SHOWbMessage(ptMsgCode As String, poMsgType As EN_TRMsgType) As Boolean
    
    If InStr(1, ptMsgCode, ";") = 0 Then ptMsgCode = ptMsgCode & ";" & ptMsgCode
    
    Dim oMsgStyle As VbMsgBoxStyle
    
    Select Case poMsgType
        Case EN_TRMsgType.Information
            oMsgStyle = vbInformation + vbOKOnly
        Case EN_TRMsgType.Exclamation
            oMsgStyle = vbExclamation + vbOKOnly
        Case EN_TRMsgType.Critical
            oMsgStyle = vbCritical + vbOKOnly
        Case EN_TRMsgType.Question
            oMsgStyle = vbQuestion + vbOKCancel
        Case EN_TRMsgType.Confirmation
            oMsgStyle = vbExclamation + vbOKCancel
    End Select
    
    Dim nLangIndex As Integer
    nLangIndex = 0
    
    If mTRVB.eVB_TRLang = Thai Then
        nLangIndex = 1
    End If
    
    If mTRVB.eVB_TRLang = English Then
        nLangIndex = 0
    End If
    
    Dim oMsgResult As VbMsgBoxResult
    oMsgResult = MsgBox(Split(ptMsgCode, ";")(nLangIndex), oMsgStyle, mTRCS.tCS_TRPrjName)
    
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
Public Sub SP_CTLxSetDataCbo(poCbo As DataCombo, ptSQLStr As String, poDbConn As ADODB.Connection)
    
    Dim oRs As ADODB.Recordset
    Set oRs = poDbConn.Execute(ptSQLStr)
    
    Set poCbo.RowSource = oRs
    poCbo.BoundColumn = oRs.Fields(0).Name
    poCbo.ListField = oRs.Fields(1).Name
    poCbo.DataField = oRs.Fields(0).Name

End Sub
Public Function SP_SQLtFormatText(ptValue As String, poDataType As EN_TRDataType) As String
    On Error GoTo Err:
    Dim tStr As String
    tStr = ptValue
    
    If mTRVB.eVB_TRDbType = ACCESS Then
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
'   Text = 0
'    Number = 1
'    Float = 2
'    Date = 3
'    Bool = 4
    If mTRVB.eVB_TRDbType = SQLServer Then
           Select Case poDataType
           Case EN_TRDataType.Date
                    tStr = "Convert(date,'" & Format(CDate(ptValue), "yyyy-MM-dd") & "')"
           Case EN_TRDataType.Text
                    tStr = "'" & Replace(tStr, "'", "''") & "'"
           Case Else
                    tStr = ptValue
           End Select
    End If
    
    SP_SQLtFormatText = tStr
    
    Exit Function
Err:
    SP_SQLtFormatText = "NULL"
End Function
