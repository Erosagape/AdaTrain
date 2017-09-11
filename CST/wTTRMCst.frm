VERSION 5.00
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7u.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form wTTRMCst 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Customer"
   ClientHeight    =   10920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10920
   ScaleWidth      =   7755
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3360
      TabIndex        =   39
      Top             =   10440
      Width           =   1215
   End
   Begin VB.CommandButton ocmSearch 
      Caption         =   "&Find"
      Height          =   375
      Left            =   6480
      TabIndex        =   30
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox otbCliteria 
      Height          =   375
      Left            =   2760
      TabIndex        =   29
      Top             =   480
      Width           =   3735
   End
   Begin VB.ComboBox ocbSearch 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   480
      Width           =   1575
   End
   Begin VB.CheckBox ockNonActive 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Non-Active"
      Height          =   255
      Left            =   6000
      TabIndex        =   26
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton ocmExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   10440
      Width           =   1095
   End
   Begin VB.Frame ofaMain 
      Caption         =   "Customer Information"
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   7575
      Begin VB.TextBox otbLastUpdate 
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   4320
         Width           =   3615
      End
      Begin VB.CommandButton ocmAdd 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   2640
         TabIndex        =   21
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton ocmDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   1440
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox otbFCCstChqBal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox otbFCCstARBal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox otbFCCreditLimit 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4200
         TabIndex        =   12
         Top             =   2520
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker odtFDBirthDate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   11
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   106102787
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin VB.CommandButton ocmSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox otbFTRemark 
         Height          =   975
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   19
         Tag             =   "FTRemark"
         Top             =   3240
         Width           =   4935
      End
      Begin VB.ListBox olbFTCstPriceLv 
         Height          =   840
         ItemData        =   "wTTRMCst.frx":0000
         Left            =   240
         List            =   "wTTRMCst.frx":0010
         TabIndex        =   18
         Top             =   3240
         Width           =   615
      End
      Begin VB.Frame ofaFTCstStatus 
         Caption         =   "Status"
         Height          =   735
         Left            =   5520
         TabIndex        =   14
         Top             =   2280
         Width           =   1935
         Begin VB.OptionButton orbInactive 
            Caption         =   "Inactive"
            Height          =   375
            Left            =   960
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton orbActive 
            Caption         =   "Active"
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox otbFTCstFax 
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox otbFTCstTel 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox otbFTCstAddress 
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   7215
      End
      Begin VB.TextBox otbFTCstName 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox otbFTCstCode 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Tag             =   "FTCstName"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label olaFCCstChqBal 
         Caption         =   "Cheque Pending"
         Height          =   375
         Left            =   6000
         TabIndex        =   36
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label olaFCCstARBal 
         Caption         =   "A/R Balance"
         Height          =   375
         Left            =   6000
         TabIndex        =   34
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label olaFCCreditLimit 
         Caption         =   "Credit (THB)"
         Height          =   375
         Left            =   4200
         TabIndex        =   33
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label olaFDBirthDate 
         Caption         =   "Birth Date"
         Height          =   375
         Left            =   2760
         TabIndex        =   32
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label olbTotalRec 
         Alignment       =   1  'Right Justify
         Caption         =   "Count Records"
         Height          =   255
         Left            =   5280
         TabIndex        =   31
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label olaFTRemark 
         Caption         =   "Remark"
         Height          =   375
         Left            =   960
         TabIndex        =   23
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label olaFTCstPriceLv 
         Caption         =   "Set Price"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label olaFTCstFax 
         Caption         =   "Customer Fax"
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label olaFTCstTel 
         Caption         =   "Customer Tel"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label olaFTCstAddress 
         Caption         =   "Customer address"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label olaFTCstName 
         Caption         =   "Customer Name"
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label olaFTCstCode 
         Caption         =   "Customer Code*"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin VSFlex7UCtl.VSFlexGrid ogdMain 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   7455
      _cx             =   13150
      _cy             =   7858
      _ConvInfo       =   -1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label olaSearch 
      BackStyle       =   0  'Transparent
      Caption         =   "Search By"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label olaCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "wTTRMCst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oW_DbConn As ADODB.Connection
Dim bW_CancelTab As Boolean



Private Sub SP_DATxSetProperty()
    '//set default max length of control FROM data dictionary
    Me.otbFTCstCode.MaxLength = 15
    Me.otbFTCstName.MaxLength = 100
    Me.otbFTCstAddress.MaxLength = 100
    Me.otbFTRemark.MaxLength = 100
    Me.otbFTCstTel.MaxLength = 10
    Me.otbFTCstFax.MaxLength = 10

    Call SP_DATxClearForm
    
End Sub
Private Function SP_SQLtGetValue(ptFieldName As String) As String
    Dim tValue As String
    Dim bFound As Boolean
    '//special case field
    Select Case ptFieldName
    Case "FTCstPriceLv"
        tValue = SP_SQLtFormatText(olbFTCstPriceLv.ListIndex + 1, Number)
        bFound = True
    Case "FTCstStatus"
        tValue = SP_SQLtFormatText(IIf(orbActive.Value = True, "0", "1"), Number)
        bFound = True
    Case "FCCreditLimit"
        tValue = SP_SQLtFormatText(otbFCCreditLimit.Text, Float)
        bFound = True
    End Select
    '/normal case field
    If bFound = False Then
        Select Case Mid(ptFieldName, 1, 2)
        Case "FD"
            tValue = SP_DATtGetInput(Me, "odt", ptFieldName)
        Case Else
            tValue = SP_DATtGetInput(Me, "otb", ptFieldName)
        End Select
    End If
    
    If tValue = "" Then tValue = "NULL"
    
    SP_SQLtGetValue = tValue
End Function
Private Function SP_SQLtForShow() As String
    '//prepare data for grid view
    Dim tSql As String
    
    tSql = "SELECT FTCstCode,FTCstName,FTCstAddress,FTCstTel,FTCstFax,FTRemark,"
    tSql = tSql & "FTCstStatus,FTCstPriceLv,FDBirthDate,FCCreditLimit,FCCstARBal,FCCstCHQBal,"
    tSql = tSql & "FDDateUpd,FTTimeUpd,FTWhoUpd,FDDateIns,FTTimeIns,FTWhoIns"
    tSql = tSql & " FROM TTRMCst "
    
    '//search cliteria default
    If ockNonActive.Value = vbChecked Then
        tSql = tSql & " WHERE FTCstStatus<>'0' "
    Else
        tSql = tSql & " WHERE FTCstStatus='0' "
    End If
    
    If otbCliteria.Text <> "" Then
        '//if search based FROM all fields or selected field to search
        If ocbSearch.ListIndex > 0 Then
            tSql = tSql & " AND " & ocbSearch.Text & " LIKE '%" & otbCliteria.Text & "%'"
        Else
            tSql = tSql & " AND ("
            tSql = tSql & "FTCstCode LIKE '%" & otbCliteria.Text & "%'  OR "
            tSql = tSql & "FTCstName LIKE '%" & otbCliteria.Text & "%'  OR "
            tSql = tSql & "FTCstAddress LIKE '%" & otbCliteria.Text & "%'  OR "
            tSql = tSql & "FTCstTel LIKE '%" & otbCliteria.Text & "%'  OR "
            tSql = tSql & "FTCstFax LIKE '%" & otbCliteria.Text & "%'  OR "
            tSql = tSql & "FDBirthDate LIKE '%" & otbCliteria.Text & "%'  OR "
            tSql = tSql & "FCCreditLimit LIKE '%" & otbCliteria.Text & "%'  OR "
            tSql = tSql & "FTCstPriceLv='" & otbCliteria.Text & "'  OR "
            tSql = tSql & "FTRemark LIKE '%" & otbCliteria.Text & "%' "
            tSql = tSql & ")"
        End If
    
    End If
    
    tSql = tSql & " ORDER BY FTCstCode "
    
    SP_SQLtForShow = tSql
End Function
Private Sub SP_DATxReadData(poTbl As ADODB.Recordset)
    If ocbSearch.ListCount = 0 Then
        '//load combobox field list for user selected to find
        Dim tExceptField As String
        tExceptField = ",FTCstStatus,FCCstARBal,FCCstCHQBal,FDDateUpd,FTTimeUpd,FTWhoUpd,FDDateIns,FTTimeIns,FTWhoIns"
        
        ocbSearch.AddItem "(All Data)"
        Dim nIdx As Integer
        For nIdx = 0 To poTbl.Fields.Count - 1
        
            If InStr(1, tExceptField, poTbl.Fields(nIdx).Name) <= 0 Then
                ocbSearch.AddItem poTbl.Fields(nIdx).Name
            End If
                       
        Next nIdx
                
        ocbSearch.ListIndex = 0
    
    End If
    
    olbTotalRec.Caption = "Found =" & poTbl.RecordCount & " Record(s)"
    
End Sub
Private Sub SP_DATxShowGrid()
    '//load main grid
    On Error GoTo Err:
    
    Dim oTbl As ADODB.Recordset
    Set oTbl = SP_TBLoGetFROMSQL(oW_DbConn, SP_SQLtForShow)
    
    Set ogdMain.DataSource = oTbl
    ogdMain.SelectionMode = flexSelectionByRow
    '//read data to search
    Call SP_DATxReadData(oTbl)
    
    Exit Sub

Err:

    Call SP_SHOWbMessage(Err.Description, Critical)
    
End Sub

Private Sub Command1_Click()
    Dim oCont As Control

    For Each oCont In Me.Controls
        Select Case TypeName(oCont)
            Case "TextBox"
            Case "ComboBox"
            Case "DataCombo"
            Case "DTPicker"
            Case "ListBox"
            Case "OptionButton"
            Case "CheckBox"
            
        End Select
     Next

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    'change enter to tab to move between objects
    If KeyCode = 13 And bW_CancelTab = False Then
    
        Dim oShell As Object
        Set oShell = CreateObject("WScript.Shell")
        oShell.SendKeys "{TAB}"
        Set oShell = Nothing
        
    End If

End Sub
Private Sub Form_Load()
    '//prepare connection
    Set oW_DbConn = mTRVB.oVB_TRDbCon
    '//show grid and set defaults
    Call SP_DATxShowGrid
    Call SP_DATxSetProperty
  
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If SP_SHOWbMessage(mTRMS.tMS_0003, Question) = False Then
        
        Cancel = 1
    
    End If

End Sub
Private Sub ockNonActive_Click()

    Call SP_DATxShowGrid
    
End Sub
Private Sub ocmAdd_Click()

    Call SP_DATxClearForm
    otbFTCstCode.SetFocus
    
End Sub
Private Sub ocmDelete_Click()
    '//delete data
    If SP_SHOWbMessage(mTRMS.tMS_0005, Confirmation) = False Then Exit Sub
    If SP_TBLbDeleteData() = True Then
    
        Call SP_DATxClearForm
        Call SP_DATxShowGrid
        Call SP_SHOWbMessage(mTRMS.tMS_0010, Exclamation)
    End If

End Sub
Private Sub ocmExit_Click()

    Unload Me
    
End Sub
Private Sub ocmSave_Click()
    '//Save Data
    If SP_DATbCheckValidate() = False Then Exit Sub
    If SP_SHOWbMessage(mTRMS.tMS_0004, Question) = False Then Exit Sub
        
    'If SP_TBLbUpdateData() = True Then     'using Recordset.update instead of SQL Command
    If SP_TBLbSaveData() = True Then        'using SQL Command for saving data
        
        Call SP_DATxShowGrid
        Call SP_DATxLoadFROMDB
        Call SP_SHOWbMessage(mTRMS.tMS_0007, Information)
        
    End If

End Sub
Private Function SP_TBLbDeleteData() As Boolean
    On Error GoTo Err:
    
    '//delete using parameters
    Dim tSQLDelete As String
    tSQLDelete = "DELETE FROM TTRMCst WHERE FTCstCode=?"
    
    Dim oCmd As New ADODB.Command
    oCmd.ActiveConnection = oW_DbConn
    oCmd.CommandType = adCmdText
    oCmd.CommandText = tSQLDelete
    
    oCmd.Parameters.Append oCmd.CreateParameter("p1", adVarChar, adParamInput, 15, Me.otbFTCstCode.Text)
    oCmd.Execute
       
    SP_TBLbDeleteData = True
    
    Exit Function

Err:
    
    SP_TBLbDeleteData = False
    Call SP_SHOWbMessage(Err.Description, Critical)
End Function
Private Function SP_DATbCheckValidate() As Boolean
    
    Dim bValid As Boolean
    bValid = True
    'if key blank then ask for input or create new automatically
    If Trim(Me.otbFTCstCode.Text) = "" Then
        bValid = SP_SHOWbMessage(mTRMS.tMS_0006, Question)
        If Not bValid Then Exit Function
    End If
    'begin check validation data
    bValid = False
    
    Dim tMsg As String
    tMsg = ""
    
    If Trim(Me.otbFTCstName.Text) = "" Then
        tMsg = tMsg & Split(mTRMS.tMS_0013, ";")(mTRVB.eVB_TRLang) & vbCrLf
    End If
    
    If Trim(Me.otbFTCstAddress.Text) = "" Then
        tMsg = tMsg & Split(mTRMS.tMS_0014, ";")(mTRVB.eVB_TRLang) & vbCrLf
    End If
    
    If Trim(Me.otbFTCstTel.Text) = "" Then
        tMsg = tMsg & Split(mTRMS.tMS_0015, ";")(mTRVB.eVB_TRLang) & vbCrLf
    Else
        If IsNumeric(Me.otbFTCstTel) = False Then
            tMsg = tMsg & Split(mTRMS.tMS_0016, ";")(mTRVB.eVB_TRLang) & vbCrLf
        End If
    End If
    
    If Trim(Me.otbFTCstFax.Text) = "" Then
        tMsg = tMsg & Split(mTRMS.tMS_0017, ";")(mTRVB.eVB_TRLang) & vbCrLf
    Else
        If IsNumeric(Me.otbFTCstFax) = False Then
            tMsg = tMsg & Split(mTRMS.tMS_0018, ";")(mTRVB.eVB_TRLang) & vbCrLf
        End If
    End If
    
    If Year(Me.odtFDBirthDate.Value) = 1900 Then
        tMsg = tMsg & Split(mTRMS.tMS_0019, ";")(mTRVB.eVB_TRLang) & vbCrLf
    End If
    
    If IsNumeric(Me.otbFCCreditLimit.Text) = False Then
        tMsg = tMsg & Split(mTRMS.tMS_0020, ";")(mTRVB.eVB_TRLang) & vbCrLf
    End If
    
    If tMsg = "" Then
        bValid = True
    Else
        Call SP_SHOWbMessage(tMsg, Exclamation)
    End If
    
    SP_DATbCheckValidate = bValid
    
End Function
Private Function SP_TBLbSaveData() As Boolean
    On Error GoTo Err:
'----------------------------------------------------------
'   *Puttipong 2017-09-07
'   บันทึกข้อมูลลูกค้าลง Database
'   Call:   -
'
'   Ret:    True = บันทึกสำเร็จ
'               False = บันทึกไม่สำเร็จ
'
'----------------------------------------------------------

    Dim bSuccess As Boolean
    bSuccess = False
    '//generate new code if not input
    If otbFTCstCode.Text = "" Then
        otbFTCstCode.Text = SP_GETtNewCustomer(oW_DbConn)
    End If
    '//read structure
    Dim oRec As ADODB.Recordset
    Set oRec = oW_DbConn.Execute("SELECT * FROM TTRMCst WHERE (1=0)")
    
    Dim nIdx As Integer
    Dim tVal As String
    Dim tSQLColName As String
    Dim tSQLValue As String
    Dim tSQLUpdate As String
    '//Read data FROM form to SQL Command
    For nIdx = 0 To oRec.Fields.Count - 1
        
        tVal = SP_SQLtGetValue(oRec.Fields(nIdx).Name)
        
        tSQLColName = tSQLColName & IIf(tSQLColName <> "", ",", "") & oRec.Fields(nIdx).Name
        tSQLValue = tSQLValue & IIf(tSQLValue <> "", ",", "") & tVal
        tSQLUpdate = tSQLUpdate & IIf(tSQLUpdate <> "", ",", "") & oRec.Fields(nIdx).Name & "=" & tVal

    Next nIdx
    oRec.Close
    
    '//Generate SQL Command
    tSQLColName = "INSERT INTO TTRMCst (" & tSQLColName & ") VALUES "
    tSQLValue = "(" & tSQLValue & ")"
    
    tSQLUpdate = Replace("UPDATE TTRMCst SET " & tSQLUpdate & " WHERE FTCstCode=?", "=?", "='" & otbFTCstCode.Text & "'")
    
    '//Try to insert first and then update if insert failed
    If SP_SQLbRunCommand(oW_DbConn, tSQLColName & tSQLValue) = False Then
        bSuccess = SP_SQLbRunCommand(oW_DbConn, tSQLUpdate)
    Else
        bSuccess = True
    End If
    
    SP_TBLbSaveData = bSuccess
    
    Exit Function
    
Err:

    SP_TBLbSaveData = False
    Call SP_SHOWbMessage(Err.Description, Critical)
    
End Function
Private Function SP_TBLbUpdateData() As Boolean
    '//using Recordset update instread of SQL Command for prevent String exception
    On Error GoTo Err:
    '//Generate new code
    If otbFTCstCode.Text = "" Then
        otbFTCstCode.Text = SP_GETtNewCustomer(oW_DbConn)
    End If
    '//Read FROM Database and filter for key input
    Dim oRs As ADODB.Recordset
    Set oRs = SP_TBLoQueryData()
    
    Dim oAction As EN_TRDbAction
    With oRs
        If .EOF = True Then
            '//if not found then set flag to insert
            .AddNew
            .Fields("FCCstARBal").Value = 0
            .Fields("FCCstChqBal").Value = 0
            
            oAction = Insert
        Else
            '//if found set flag to update
            oAction = Update
        End If
        
        '//read input data into fields
        .Fields("FTCstCode").Value = otbFTCstCode.Text
        .Fields("FTCstName").Value = otbFTCstName.Text
        .Fields("FTCstAddress").Value = otbFTCstAddress.Text
        .Fields("FTCstTel").Value = otbFTCstTel.Text
        .Fields("FTCstFax").Value = otbFTCstFax.Text
        .Fields("FTCstStatus").Value = IIf(orbActive.Value = True, 0, 1)
        .Fields("FTCstPriceLv").Value = olbFTCstPriceLv.ListIndex + 1
        .Fields("FTRemark").Value = otbFTRemark.Text
        .Fields("FCCreditLimit").Value = otbFCCreditLimit.Text
        .Fields("FDBirthDate").Value = Format(CDate(odtFDBirthDate.Value), "yyyy-MM-dd")
        
        .Update
        .Close
    
    End With
    '//update for log
    Call SP_SQLxSetLogTBL(oAction, "TTRMCst", "FTCstCode='" & otbFTCstCode.Text & "'", oW_DbConn)

    SP_TBLbUpdateData = True
    
    Exit Function

Err:

    Call SP_SHOWbMessage(Err.Description, Critical)
    SP_TBLbUpdateData = False

End Function
Private Sub SP_DATxClearForm(Optional ptCode As String = "")
    '//clear form
    Me.otbFTCstCode.Text = ptCode
    Me.otbFTCstName.Text = ""
    Me.otbFTCstAddress.Text = ""
    Me.otbFTCstTel.Text = ""
    Me.otbFTCstFax.Text = ""
    Me.otbFTRemark.Text = ""
    Me.orbActive.Value = True
    Me.olbFTCstPriceLv.ListIndex = 0
    Me.otbFCCstARBal.Text = "0"
    Me.otbFCCstChqBal.Text = "0"
    Me.otbFCCreditLimit.Text = "0"
    Me.odtFDBirthDate.Value = Me.odtFDBirthDate.MinDate
    Me.otbLastUpdate.Text = ""
    
    Me.otbFTCstCode.Locked = False
    Me.otbFTCstCode.BackColor = vbWhite

End Sub
Private Function SP_TBLoQueryData() As ADODB.Recordset
    '//Query database for current key input
    Dim oTbl As New ADODB.Recordset
    
    Set oTbl = SP_TBLoGetCustomer(oW_DbConn)
    oTbl.Filter = "[FTCstCode]='" & otbFTCstCode.Text & "'"
    
    Set SP_TBLoQueryData = oTbl
    
End Function
Private Sub SP_DATxLoadFROMDB()
    On Error GoTo Err:
        
    '//Read FROM data when user input key
    Dim oTbl As ADODB.Recordset
    Set oTbl = SP_TBLoQueryData()
    
    If oTbl.EOF = True Then
    
        '//if not found then clear form and show msgbox data not found
        Call SP_DATxClearForm(Me.otbFTCstCode.Text)
        
        oTbl.Close
        Set oTbl = Nothing
            
        Call SP_SHOWbMessage(mTRMS.tMS_0011, Exclamation)
    
        Exit Sub
        
    End If
    '//if data found then read FROM database and put into controls
    With oTbl
        Me.otbFTCstCode.Text = "" & .Fields("FTCstCode").Value
        Me.otbFTCstName.Text = "" & .Fields("FTCstName").Value
        Me.otbFTCstAddress.Text = "" & .Fields("FTCstAddress").Value
        Me.otbFTCstTel.Text = "" & .Fields("FTCstTel").Value
        Me.otbFTCstFax.Text = "" & .Fields("FTCstFax").Value
        Me.otbFTRemark.Text = "" & .Fields("FTRemark").Value
        Me.orbActive.Value = IIf(.Fields("FTCstStatus").Value = "0", True, False)
        Me.orbInactive.Value = IIf(.Fields("FTCstStatus").Value = "0", False, True)
        Me.olbFTCstPriceLv.ListIndex = CInt(.Fields("FTCstPriceLv").Value) - 1
        
        Me.ocmDelete.Enabled = Me.orbInactive.Value
        
        If Not IsDate("" & .Fields("FDBirthDate").Value) Then
            Me.odtFDBirthDate.Value = Me.odtFDBirthDate.MinDate
        Else
            Me.odtFDBirthDate.Value = CDate(.Fields("FDBirthDate").Value)
        End If
        
        Me.otbFCCreditLimit.Text = Format(.Fields("FCCreditLimit").Value, "#0.00#")
        Me.otbFCCstARBal.Text = Format(.Fields("FCCstARBal").Value, "#0.00#")
        Me.otbFCCstChqBal.Text = Format(.Fields("FCCstChqBal").Value, "#0.00#")
        Me.otbLastUpdate.Text = "Last Update By " & .Fields("FTWhoUpd").Value & " On " & .Fields("FDDateUpd").Value & " " & .Fields("FTTimeUpd").Value

    End With

    Me.otbFTCstCode.Locked = True
    Me.otbFTCstCode.BackColor = &H8000000F

    oTbl.Close
    Set oTbl = Nothing
    Exit Sub
    
Err:

    If oTbl.State = 1 Then oTbl.Close
    Set oTbl = Nothing
    
    Call SP_SHOWbMessage(Err.Description, Critical)
    
End Sub
Private Sub ocmSearch_Click()

    Call SP_DATxShowGrid
    
End Sub
Private Sub ogdMain_Click()
    
        Call ogdMain_RowColChange
        otbFTCstName.SetFocus
            
End Sub
Private Sub ogdMain_RowColChange()
    
    If ogdMain.Row > 0 Then

        Me.otbFTCstCode.Text = ogdMain.TextMatrix(ogdMain.Row, 1)
        Call SP_DATxLoadFROMDB
        
    End If

End Sub

Private Sub otbCliteria_GotFocus()
    
    Call SP_CTLxSetFocus(Me.ActiveControl)
    
End Sub
Private Sub otbFCCreditLimit_GotFocus()
    
    Call SP_CTLxSetFocus(Me.ActiveControl)
    
End Sub
Private Sub otbFTCstAddress_GotFocus()
    
    bW_CancelTab = True
    'Call SP_CTLxSetFocus(Me.ActiveControl)
    
End Sub
Private Sub otbFTCstAddress_LostFocus()
    
    bW_CancelTab = False
    
End Sub
Private Sub otbFTCstCode_GotFocus()

    Call SP_CTLxSetFocus(Me.ActiveControl)
    
End Sub
Private Sub otbFTCstCode_LostFocus()

        If Trim(otbFTCstCode.Text) = "" Then Exit Sub
        
        Call SP_DATxLoadFROMDB

End Sub
Private Sub otbFTCstFax_GotFocus()
    
    Call SP_CTLxSetFocus(Me.ActiveControl)

End Sub
Private Sub otbFTCstName_GotFocus()

    Call SP_CTLxSetFocus(Me.ActiveControl)
    
End Sub
Private Sub otbFTCstTel_GotFocus()

    Call SP_CTLxSetFocus(Me.ActiveControl)
    
End Sub
Private Sub otbFTRemark_GotFocus()
    
    bW_CancelTab = True
    'Call SP_CTLxSetFocus(Me.ActiveControl)
    
End Sub
Private Sub otbFTRemark_LostFocus()

    bW_CancelTab = False

End Sub
