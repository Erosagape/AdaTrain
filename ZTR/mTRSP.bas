Attribute VB_Name = "mTRSP"
Option Explicit
Public Sub SP_SETxVariable(ptUser As String, poDbType As EN_TRDbType, poLang As EN_TRLang, pbTestMode As Boolean)
'-----------------------------------------------------------------
'   Created by*Puttipong 2017-09-06
'   Cmt : ������Ѻ��˹���� Global Variabe �������ҧ Connection
'   Call :
'       ptUser =���� user �Ѩ�غѹ
'       poDbType = Enum �������ͧ Database �� EN_TRDbType ��Сͺ
'       poLang = Enum ���һѨ�غ� �� EN_TRLang
'       pbTestMode = True ����� Database Test
'                                 False ����� Database ��ԧ
'   Ret :
'
'-----------------------------------------------------------------
    tVB_TRUser = ptUser
    eVB_TRDbType = poDbType
    eVB_TRLang = poLang
    
    If pbTestMode = True Then
        Set oVB_TRDbCon = SP_GEToConnTEST()
    End If
    
    oVB_TRDbCon.Open
    
End Sub
Public Function SP_GETtConnStr(ptUserID As String, ptPassword As String, ptServerName As String, ptDatabaseName As String) As String
'-----------------------------------------------------------------
'   Created by*Puttipong 2017-09-06
'   Cmt : ������Ѻ���ҧ Connection �ҡ����÷���к�
'   Call :
'       ptUserID =���� Database User
'       ptPassword =���� Database Password
'       ptServerName = ����/Path �ͧ Database Server
'       ptDatabaseName =���� Database
'   Return :
'       String �� Connection String
'-----------------------------------------------------------------
    
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
    
    If eVB_TRDbType = EN_TRDbType.SQLServer Then
    
        tProvider = tCS_TRPrvSQL
        
        tConnStr = tCS_TRConSQL
        tConnStr = Replace(tConnStr, "{0}", tProvider)
        tConnStr = Replace(tConnStr, "{1}", tServerName)
        tConnStr = Replace(tConnStr, "{2}", tUserID)
        tConnStr = Replace(tConnStr, "{3}", tPassword)
        tConnStr = Replace(tConnStr, "{4}", tDatabase)
        
    End If
    
    If eVB_TRDbType = EN_TRDbType.ACCESS Then
        tProvider = tCS_TRDbPrvMSAC
        tServerName = App.Path & "\" & tDatabase & ".mdb"
        
        tConnStr = tCS_TRConMSAC
        tConnStr = Replace(tConnStr, "{0}", tProvider)
        tConnStr = Replace(tConnStr, "{1}", tServerName)
        tConnStr = Replace(tConnStr, "{2}", "Admin")
        tConnStr = Replace(tConnStr, "{3}", "")
            
    End If
            
    SP_GETtConnStr = tConnStr
    
End Function
Public Function SP_GEToConnTEST() As ADODB.Connection
'-----------------------------------------------------------------
'   Created by*Puttipong 2017-09-06
'   Cmt : ������Ѻ���¡����� Connection Test
'   Call :
'   Return :
'       Connection Test (ADODB.Connection)
'-----------------------------------------------------------------
    
    On Error GoTo ErrHandle:
    
    Dim oDbConn As New ADODB.Connection
    
    oDbConn.ConnectionString = SP_GETtConnStr(tCS_TRDbUser, tCS_TRDbPwd, ".", tCS_TRDbName)
    oDbConn.CursorLocation = adUseClient
    
    Set SP_GEToConnTEST = oDbConn
    
    Exit Function
    
ErrHandle:

    Set SP_GEToConnTEST = New ADODB.Connection
    
End Function
Public Function SP_GEToData(poDbConn As ADODB.Connection, ptAliasName As String) As ADODB.Recordset
'-----------------------------------------------------------------
'   Created by*Puttipong 2017-09-11
'   Cmt : ������Ѻ�֧������ Table �ҡ Database ���������
'   Call :
'       poDbConn = Connection �������
'       ptAliasName = �����ͧ͢ Table ���д֧������
'   Return :
'       DataTable �ͧ Table ������͡ (ADODB.RecordSet)
'-----------------------------------------------------------------

Select Case ptAliasName
Case "Cst"
    Set SP_GEToData = SP_GEToTbl(poDbConn, tCS_TRSQLCst)
Case "Pdt"
    Set SP_GEToData = SP_GEToTbl(poDbConn, tCS_TRSQLPdt)
Case "PdtGrp"
    Set SP_GEToData = SP_GEToTbl(poDbConn, tCS_TRSQLPdtGrp)
Case "Spn"
    Set SP_GEToData = SP_GEToTbl(poDbConn, tCS_TRSQLSpn)
Case Else
    Set SP_GEToData = New ADODB.Recordset
End Select
End Function
Public Function SP_GEToTbl(poDbConn As ADODB.Connection, ptSQL As String) As ADODB.Recordset
'-----------------------------------------------------------------
'   Created by*Puttipong 2017-09-11
'   Cmt :
'        ������Ѻ�֧������ Table �ҡ Database �������� SQL ����Դ��� RecordSet ����ö�����ʹ Update ��
'   Call :
'       poDbConn = Connection �������
'       ptSQL = ����� SQL Select From Where ����ͧ���
'   Return :
'       DataTable �ͧ Table ������͡ (ADODB.RecordSet)
'-----------------------------------------------------------------

    On Error GoTo ErrHandle:
    
    Dim oRs As New ADODB.Recordset
    oRs.Open ptSQL, poDbConn, adOpenDynamic, adLockOptimistic
    
    Set SP_GEToTbl = oRs
    Exit Function
    
ErrHandle:

    Set SP_GEToTbl = New ADODB.Recordset

End Function
Public Function SP_EXECoSQL(poDbConn As ADODB.Connection, ptSQL As String) As ADODB.Recordset
'-----------------------------------------------------------------
'   Created by*Puttipong 2017-09-11
'   Cmt : ������Ѻ�֧������ Table �ҡ Database �������� SQL ���ԴẺ Read only
'   Call :
'       poDbConn = Connection �������
'       ptSQL = ����� SQL
'   Return :
'       DataTable �ͧ Table ������͡ (ADODB.RecordSet)
'-----------------------------------------------------------------

    On Error GoTo ErrHandle:
    
    Set SP_EXECoSQL = poDbConn.Execute(ptSQL)
    Exit Function
    
ErrHandle:

    Set SP_EXECoSQL = New ADODB.Recordset
End Function
Public Function SP_GETtNewID(poDbConn As ADODB.Connection, ptAliasName As String) As String
'-----------------------------------------------------------------
'   Created by*Puttipong 2017-09-11
'   Cmt : ������Ѻ�Ҥ�� Running Number ����ش�ҡ������ Table
'   Call :
'       poDbConn = Connection �������
'       ptAliasName = �����ͧ͢ Table �����Ң�����
'   Return :
'       String �ͧ Running ������ Format ����
'-----------------------------------------------------------------

Dim oRs As ADODB.Recordset
Select Case ptAliasName
    Case "Cst"
    
        Set oRs = SP_EXECoSQL(poDbConn, tCS_TRSQLCstNew)
        If oRs.EOF = False Then
            SP_GETtNewID = "C-" & Format(CInt(Right(oRs.Fields(0).Value, 4)) + 1, "0000")
        Else
            SP_GETtNewID = "C-0001"
        End If
        oRs.Close
        
    Case "Spn"
    
        Set oRs = SP_EXECoSQL(poDbConn, tCS_TRSQLSpnNew)
        If oRs.EOF = False Then
            SP_GETtNewID = "E-" & Format(CInt(Right(oRs.Fields(0).Value, 4)) + 1, "0000")
        Else
            SP_GETtNewID = "E-0001"
        End If
        oRs.Close
        
    Case "PdtGrp"
        
        Set oRs = SP_EXECoSQL(poDbConn, tCS_TRSQLPdtGrpNew)
        If oRs.EOF = False Then
            SP_GETtNewID = Format(CInt(Right(oRs.Fields(0).Value, 4)) + 1, "000")
        Else
            SP_GETtNewID = "001"
        End If
        oRs.Close
Case Else
    SP_GETtNewID = ""
End Select
End Function
Public Function SP_GETtSQLDefault(poActionType As EN_TRDbAction) As String
'-----------------------------------------------------------------
'   Created by*Puttipong 2017-09-11
'   Cmt : ������Ѻ���ҧ����� SQL Insert/Update ����Ѻ Field �ҵðҹ�ͧ Ada
'   Call :
'       poActionType =�������ͧ����� SQL Insert ���� SQL Update �� EN_TRDbAction
'   Return :
'       String ����� SQL �������к�
'-----------------------------------------------------------------

Dim tFldNme As String
Dim tFldVal As String
Dim tSQL As String
Dim tExceptField As String
tFldNme = ""
tFldVal = ""
tSQL = ""
tExceptField = "FDDateIns,FTTimeIns,FTWhoIns"

Dim oArr As Variant
oArr = Split(mTRCS.tCS_TRDbFldDef, ",")

Dim nIdx As Integer
For nIdx = 0 To UBound(oArr)
    If poActionType = Insert Then
    
        If tFldNme <> "" Then tFldNme = tFldNme & ","
        tFldNme = tFldNme & oArr(nIdx)
        
        If tFldVal <> "" Then tFldVal = tFldVal & ","
        tFldVal = tFldVal & SP_GETtSQLDefVal("" & oArr(nIdx))
    Else
        tFldNme = oArr(nIdx)
        If InStr(1, tExceptField, tFldNme) <= 0 Then

            tFldVal = SP_GETtSQLDefVal(tFldNme)
            If tSQL <> "" Then tSQL = tSQL & "," & vbCrLf
            tSQL = tSQL & tFldNme & "=" & tFldVal
        End If
    End If
Next nIdx

If tSQL = "" Then
    tSQL = tFldNme & ";" & tFldVal
End If

SP_GETtSQLDefault = tSQL

End Function
Public Function SP_GETtSQLDefVal(ptName As String) As String
'-----------------------------------------------------------------
'   Created by*Puttipong 2017-09-11
'   Cmt : ������Ѻ�Ҥ�� Default �ͧ���� Field ����ͧ���
'   Call :
'       ptName = ���Ϳ�Ŵ����к�
'   Return :
'       ��� Default �ͧ��Ŵ����к���ٻẺ Database Format ����
'-----------------------------------------------------------------

    '//special case field
    Select Case ptName
    Case "FDDateIns", "FDDateUpd"
        SP_GETtSQLDefVal = SP_GETtSQLFormat(Format(Now, "yyyy-MM-dd"), Date)
    Case "FTTimeIns", "FTTimeUpd"
        SP_GETtSQLDefVal = SP_GETtSQLFormat(Format(Now, "HH:mm:ss"), Text)
    Case "FTWhoIns", "FTWhoUpd"
        SP_GETtSQLDefVal = SP_GETtSQLFormat(tVB_TRUser, Text)
    Case Else
        SP_GETtSQLDefVal = SP_GETtSQLFormat(Format(Now, "yyyy-MM-dd"), Date)
    End Select
End Function
Public Sub SP_SETxLogTbl(poAction As EN_TRDbAction, ptTableName As String, ptWhere As String, poDbConn As ADODB.Connection)
'-----------------------------------------------------------------
'   Created by*Puttipong 2017-09-11
'   Cmt : ������Ѻ Run ����� Update Field �ҵðҹ�ͧ Ada �� Table ����к���е�� Where Clause
'   Call :
'       poAction = �������ͧ����� SQL Insert ���� Update �� EN_TRDbAction
'       ptTableName = ���� Table ���� Update
'       ptWhere = ����� SQL Where Clause �ͧ�Ƿ��� Update
'   Return :
'       -
'-----------------------------------------------------------------

    Dim tSQL As String
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
        
        tSQL = " Update " & ptTableName & " set "
        tSQL = tSQL & " FDDate" & tLogType & "={0}, "
        tSQL = tSQL & " FTTime" & tLogType & "={1}, "
        tSQL = tSQL & " FTWho" & tLogType & "='" & tVB_TRUser & "'"
        
        If tLogType = "Ins" Then
            tSQL = tSQL & " ,FDDateUpd={0}, "
            tSQL = tSQL & " FTTimeUpd={1}, "
            tSQL = tSQL & " FTWhoUpd='" & tVB_TRUser & "'"
        End If
        
        If eVB_TRDbType = ACCESS Then
        
                tSQL = Replace(tSQL, "{0}", "Format(Now(),'yyyy-MM-dd')")
                tSQL = Replace(tSQL, "{1}", "Format(Now(),'HH:mm:ss')")
                
        End If
        
        If eVB_TRDbType = SQLServer Then
        
                tSQL = Replace(tSQL, "{0}", "Convert(date,GetDate())")
                tSQL = Replace(tSQL, "{1}", "Convert(time,GetDate())")
                
        End If
                
        tSQL = tSQL & " WHERE " & ptWhere
        
        poDbConn.Execute tSQL
    
    End If
End Sub

Public Function SP_SHOWbMessage(ptMsgCode As String, poMsgType As EN_TRMsgType) As Boolean
'-----------------------------------------------------------------
'   Created by*Puttipong 2017-09-06
'   Cmt : ������Ѻ�ʴ������� Message Box �ҵðҹ�ͧ Window
'   Call :
'       poMsgCode = ���ʢͧ��ͤ���� module MS
'       ptMsgType = �������ͧ Message Type �� EN_TRMsgType
'                              Information,Critical ,Confirmation ������Ѻ Alert ������Һ
'                              Question,Exclamation ������Ѻ�׹�ѹ
'   Return :
'       MsgBoxResult �� vbYes/vbNo/vbOk/vbCancel �繵�
'-----------------------------------------------------------------
    
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
    
    If eVB_TRLang = Thai Then
        nLangIndex = 1
    End If
    
    If eVB_TRLang = English Then
        nLangIndex = 0
    End If
    
    Dim oMsgResult As VbMsgBoxResult
    oMsgResult = MsgBox(Split(ptMsgCode, ";")(nLangIndex), oMsgStyle, tCS_TRPrjName)
    
    SP_SHOWbMessage = IIf(oMsgResult = vbOK, True, False)
    
End Function

Public Sub SP_SETxCtlSelect(ByRef poCtl As Object)
'-----------------------------------------------------------------
'   Created by*Puttipong 2017-09-06
'   Cmt : ������Ѻ Highlight �����ŵ͹ Set Focus
'   Call :
'       poCtl = Control ����ͧ���
'
'   Return :
'       -
'-----------------------------------------------------------------

    poCtl.SelStart = 0
    poCtl.SelLength = Len(poCtl.Text)

End Sub
Public Function SP_EXECbSQL(poDbConn As ADODB.Connection, ptSQLText As String) As Boolean
'-----------------------------------------------------------------
'   Created by*Puttipong 2017-09-11
'   Cmt : ������Ѻ�ѹ����� SQL ����ͧ������׹�������ѹ��������� Error �蹾ǡ Insert/update/delete
'   Call :
'       poDbConn = Connection �������
'       ptSQLText= ����� SQL
'   Return :
'       True �������ö�ѹ��ҹ����� Error
'       False ��� Error
'-----------------------------------------------------------------

On Error GoTo ErrHandle:
    
    Dim oCmd As New ADODB.Command
    
    oCmd.ActiveConnection = poDbConn
    oCmd.CommandType = adCmdText
    oCmd.CommandText = ptSQLText
                        
    oCmd.Execute
    
    SP_EXECbSQL = True
    Exit Function
    
ErrHandle:
    SP_EXECbSQL = False
    
End Function
Public Sub SP_SETxCtlDCbo(poCbo As DataCombo, ptSQLStr As String, poDbConn As ADODB.Connection)
'-----------------------------------------------------------------
'   Created by*Puttipong 2017-09-11
'   Cmt : ������Ѻ��˹����� DataCombo ������͡
'   Call :
'       poCbo = object DataCombo ���С�˹����
'       ptSQLStr = ����� SQL ��ͧ�� Select 2 ��Ŵ����ҧ��� ����� field ����á�� key ��� field ������ text
'       poDbConn = Connection �������
'   Return :
'       -
'-----------------------------------------------------------------
    
    Dim oRs As ADODB.Recordset
    Set oRs = poDbConn.Execute(ptSQLStr)
    
    Set poCbo.RowSource = oRs
    poCbo.BoundColumn = oRs.Fields(0).Name
    poCbo.ListField = oRs.Fields(1).Name
    poCbo.DataField = oRs.Fields(0).Name

End Sub
Public Function SP_GETtSQLFormat(ptValue As String, poDataType As EN_TRDataType) As String
'-----------------------------------------------------------------
'   Created by*Puttipong 2017-09-06
'   Cmt : ������Ѻ��� Format ��� Data ���͹����㹤���� SQL Insert/Update
'   Call :
'       ptValue = ������
'       poDataType = �����������ŷ����ŧ �� EN_TRDataType
'                           �� Date,Text,Bool,Float,Number �繵�
'   Return :
'       String ������ Format ��� SQL Database �������������
'-----------------------------------------------------------------

    On Error GoTo ErrHandle:
    Dim tStr As String
    tStr = ptValue
    
    If eVB_TRDbType = ACCESS Then
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
    If eVB_TRDbType = SQLServer Then
           Select Case poDataType
           Case EN_TRDataType.Date
                    tStr = "Convert(date,'" & Format(CDate(ptValue), "yyyy-MM-dd") & "')"
           Case EN_TRDataType.Text
                    tStr = "'" & Replace(tStr, "'", "''") & "'"
           Case Else
                    tStr = ptValue
           End Select
    End If
    
    SP_GETtSQLFormat = tStr
    
    Exit Function
    
ErrHandle:
    SP_GETtSQLFormat = "NULL"
End Function
