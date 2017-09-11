VERSION 5.00
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form wTTRMPdt 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Product"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9990
   ScaleWidth      =   7935
   WindowState     =   2  'Maximized
   Begin VB.CommandButton ocmExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   240
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   9480
      Width           =   1095
   End
   Begin VB.Frame ofaProductInfo 
      Caption         =   "Product Information"
      Height          =   4335
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   7575
      Begin VB.CommandButton ocmSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton ocmDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   1320
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton ocmAdd 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   2520
         TabIndex        =   30
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox otbLastUpdate 
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3720
         Width           =   3615
      End
      Begin VB.TextBox otbFCPriceSale4 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6120
         TabIndex        =   26
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox otbFCPriceSale3 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6120
         TabIndex        =   24
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox otbFCPriceSale2 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6120
         TabIndex        =   22
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox otbFCPriceSale1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6120
         TabIndex        =   20
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox otbFTRemark 
         Height          =   1455
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   19
         Tag             =   "FTCstName"
         Top             =   2040
         Width           =   4695
      End
      Begin MSDataListLib.DataCombo ocbFTPdtGroup 
         Height          =   315
         Left            =   3960
         TabIndex        =   12
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox otbFTPdtName 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Tag             =   "FTCstName"
         Top             =   1320
         Width           =   5775
      End
      Begin VB.TextBox otbFTPdtBarcode 
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Tag             =   "FTCstName"
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox otbFTPdtCode 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Tag             =   "FTCstName"
         Top             =   600
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo ocbFTPdtUnit 
         Height          =   315
         Left            =   6120
         TabIndex        =   16
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label olaFCPriceSale4 
         Caption         =   "Sale Price 4"
         Height          =   375
         Left            =   5040
         TabIndex        =   27
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label olaFCPriceSale3 
         Caption         =   "Sale Price 3"
         Height          =   375
         Left            =   5040
         TabIndex        =   25
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label olaFCPriceSale2 
         Caption         =   "Sale Price 2"
         Height          =   375
         Left            =   5040
         TabIndex        =   23
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label olaFCPriceSale1 
         Caption         =   "Sale Price 1"
         Height          =   375
         Left            =   5040
         TabIndex        =   21
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label olaFTRemark 
         Caption         =   "Product Remark"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label olaFTPdtUnit 
         Caption         =   "Product Unit"
         Height          =   255
         Left            =   6120
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label olbPdtGroup 
         Caption         =   "Product Group"
         Height          =   255
         Left            =   3960
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label olaFTPdtName 
         Caption         =   "Product Name"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label olbTotalRec 
         Alignment       =   1  'Right Justify
         Caption         =   "Count Records"
         Height          =   255
         Left            =   5520
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label olaFTPdtBarcode 
         Caption         =   "Product Barcode"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label olaFTPdtCode 
         Caption         =   "Product Code*"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.ComboBox ocbSearch 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox otbCliteria 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton ocmSearch 
      Caption         =   "&Find"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VSFlex7UCtl.VSFlexGrid ogdMain 
      Height          =   3855
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      Width           =   7455
      _cx             =   13150
      _cy             =   6800
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
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label olaCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Data"
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
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "wTTRMPdt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oW_DbConn As ADODB.Connection
Dim bW_CancelTab As Boolean
Private Function SP_SQLtForShow() As String

    Dim tSql As String
    tSql = "SELECT FTPdtCode, FTPdtBarCode, FTPdtName, FTPdtUnit, FTPdtGroup, FCPriceSale1, FCPriceSale2, FCPriceSale3, FCPriceSale4, FTRemark, "
    tSql = tSql & " FDDateUpd, FTTimeUpd, FTWhoUpd, FDDateIns, FTTimeIns, FTWhoIns"
    tSql = tSql & " FROM TTRMPdt "
    
    If otbCliteria.Text <> "" Then
        '//if search based FROM all fields or selected field to search
        If ocbSearch.ListIndex > 0 Then
            tSql = tSql & " WHERE " & ocbSearch.Text & " LIKE '%" & otbCliteria.Text & "%'"
        Else
            tSql = tSql & " WHERE ("
            tSql = tSql & "FTPdtCode LIKE '%" & otbCliteria.Text & "%'  OR "
            tSql = tSql & "FTPdtName LIKE '%" & otbCliteria.Text & "%'  OR "
            tSql = tSql & "FTPdtUnit LIKE '%" & otbCliteria.Text & "%'  OR "
            tSql = tSql & "FTPdtGroup LIKE '%" & otbCliteria.Text & "%'  OR "
            tSql = tSql & "FCPriceSale1 LIKE '%" & otbCliteria.Text & "'  OR "
            tSql = tSql & "FCPriceSale2 LIKE '%" & otbCliteria.Text & "'  OR "
            tSql = tSql & "FCPriceSale3 LIKE '%" & otbCliteria.Text & "'  OR "
            tSql = tSql & "FCPriceSale4 LIKE '%" & otbCliteria.Text & "'  OR "
            tSql = tSql & "FTRemark LIKE '%" & otbCliteria.Text & "%' "
            tSql = tSql & ")"
        End If
    
    End If
    
    tSql = tSql & " ORDER BY FTPdtCode "
    
    SP_SQLtForShow = tSql
End Function
Private Function SP_TBLoQueryData(Optional ptFieldKey As String = "FTPdtCode") As ADODB.Recordset
    '//Query database for current key input
    Dim oTbl As New ADODB.Recordset
    
    Set oTbl = SP_TBLoGetProduct(oW_DbConn)
    oTbl.Filter = "[" & ptFieldKey & "]='" & otbFTPdtCode.Text & "'"
    
    Set SP_TBLoQueryData = oTbl
    
End Function
Private Sub SP_DATxReadData(poTbl As ADODB.Recordset)
    If ocbSearch.ListCount = 0 Then
        '//load combobox field list for user selected to find
        Dim tExceptField As String
        tExceptField = ",FDDateUpd,FTTimeUpd,FTWhoUpd,FDDateIns,FTTimeIns,FTWhoIns"
        
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
Private Sub SP_DATxSetProperty()
    '//set default max length of control FROM data dictionary
    Me.otbFTPdtCode.MaxLength = 15
    Me.otbFTPdtBarcode.MaxLength = 50
    Me.otbFTPdtName.MaxLength = 100
    Me.otbFTRemark.MaxLength = 100
    
    Call SP_CTLxSetDataCbo(ocbFTPdtGroup, "SELECT FTPdtGrpCode,FTPdtGrpName FROM TTRMPdtGrp ", oW_DbConn)
    Call SP_CTLxSetDataCbo(ocbFTPdtUnit, "SELECT FTUntCode,FTUntName FROM TTRMUnit ", oW_DbConn)
    
    Call SP_DATxClearForm
    
End Sub
Private Sub SP_DATxClearForm(Optional ptCode As String = "")
    '//clear form
    Me.otbFTPdtCode.Text = ptCode
    Me.otbFTPdtName.Text = ""
    Me.otbFTPdtBarcode.Text = ""
    Me.ocbFTPdtGroup.Text = ""
    Me.ocbFTPdtUnit.Text = ""
    Me.otbFTRemark.Text = ""
    Me.otbFCPriceSale1.Text = "0"
    Me.otbFCPriceSale2.Text = "0"
    Me.otbFCPriceSale3.Text = "0"
    Me.otbFCPriceSale4.Text = "0"
    Me.otbLastUpdate.Text = ""
    
    Me.otbFTPdtCode.Locked = False
    Me.otbFTPdtCode.BackColor = vbWhite
    
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

Private Sub ocmAdd_Click()
    Call SP_DATxClearForm
    otbFTPdtCode.SetFocus
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
Private Function SP_TBLbDeleteData() As Boolean
    On Error GoTo Err:
    
    '//delete using parameters
    Dim tSQLDelete As String
    tSQLDelete = "DELETE FROM TTRMPdt WHERE FTPdtCode=?"
    
    Dim oCmd As New ADODB.Command
    oCmd.ActiveConnection = oW_DbConn
    oCmd.CommandType = adCmdText
    oCmd.CommandText = tSQLDelete
    
    oCmd.Parameters.Append oCmd.CreateParameter("p1", adVarChar, adParamInput, 15, Me.otbFTPdtCode.Text)
    oCmd.Execute
       
    SP_TBLbDeleteData = True
    
    Exit Function

Err:
    
    SP_TBLbDeleteData = False
    Call SP_SHOWbMessage(Err.Description, Critical)
End Function

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

Private Sub ocmSearch_Click()
    Call SP_DATxShowGrid
End Sub
Private Sub ogdMain_Click()
    
        Call ogdMain_RowColChange
        otbFTPdtBarcode.SetFocus
        
End Sub
Private Sub ogdMain_RowColChange()
    
    If ogdMain.Row > 0 Then

        Me.otbFTPdtCode.Text = ogdMain.TextMatrix(ogdMain.Row, 1)
        Call SP_DATxLoadFROMDB
        
    End If

End Sub
Private Sub SP_DATxLoadFROMDB()
    On Error GoTo Err:
   
    '//Read FROM data when user input key
    Dim oTbl As ADODB.Recordset
    Set oTbl = SP_TBLoQueryData()
    
    If oTbl.EOF = True Then
    
        '//if not found then clear form and show msgbox data not found
        Call SP_DATxClearForm(Me.otbFTPdtCode.Text)
        
        oTbl.Close
        Set oTbl = Nothing
            
        Call SP_SHOWbMessage(mTRMS.tMS_0011, Exclamation)
    
        Exit Sub
        
    End If
    '//if data found then read FROM database and put into controls
    With oTbl
        Me.otbFTPdtCode.Text = "" & .Fields("FTPdtCode").Value
        Me.otbFTPdtBarcode.Text = "" & .Fields("FTPdtBarCode").Value
        Me.otbFTPdtName.Text = "" & .Fields("FTPdtName").Value
        Me.otbFTRemark.Text = "" & .Fields("FTRemark").Value
        Me.otbFCPriceSale1.Text = Format(.Fields("FCPriceSale1").Value, "#0.00#")
        Me.otbFCPriceSale2.Text = Format(.Fields("FCPriceSale2").Value, "#0.00#")
        Me.otbFCPriceSale3.Text = Format(.Fields("FCPriceSale3").Value, "#0.00#")
        Me.otbFCPriceSale4.Text = Format(.Fields("FCPriceSale4").Value, "#0.00#")
        Me.otbLastUpdate.Text = "Last Update By " & .Fields("FTWhoUpd").Value & " On " & .Fields("FDDateUpd").Value & " " & .Fields("FTTimeUpd").Value
        Me.ocbFTPdtUnit.BoundText = "" & .Fields("FTPdtUnit").Value
        Me.ocbFTPdtGroup.BoundText = "" & .Fields("FTPdtGroup").Value
        
    End With

    Me.otbFTPdtCode.Locked = True
    Me.otbFTPdtCode.BackColor = &H8000000F

    oTbl.Close
    Set oTbl = Nothing
    Exit Sub
    
Err:

    If oTbl.State = 1 Then oTbl.Close
    Set oTbl = Nothing
    
    Call SP_SHOWbMessage(Err.Description, Critical)
    
End Sub
Private Function SP_DATbCheckValidate() As Boolean
    
    Dim bValid As Boolean
    'begin check validation data
    bValid = False
    
    Dim tMsg As String
    tMsg = ""
    
    If Trim(Me.otbFTPdtCode.Text) = "" Then
        tMsg = tMsg & Split("Product Code must be input;คุณยังไม่ได้ระบุรหัสสินค้า", ";")(mTRVB.eVB_TRLang) & vbCrLf
    End If
    
    If Trim(Me.otbFTPdtBarcode.Text) = "" Then
        tMsg = tMsg & Split("Product Barcode must be input;คุณยังไม่ได้ระบุรหัสบาร์โค้ดสินค้า", ";")(mTRVB.eVB_TRLang) & vbCrLf
    End If
    
    If Trim(Me.otbFTPdtName.Text) = "" Then
        tMsg = tMsg & Split("Product Name must be input;คุณยังไม่ได้ระบุชื่อสินค้า", ";")(mTRVB.eVB_TRLang) & vbCrLf
    End If
       
    If Trim(Me.ocbFTPdtGroup.Text) = "" Then
        tMsg = tMsg & Split("Product Group must be input;คุณยังไม่ได้ระบุกลุ่มสินค้า", ";")(mTRVB.eVB_TRLang) & vbCrLf
    End If
    
    If Trim(Me.ocbFTPdtUnit.Text) = "" Then
        tMsg = tMsg & Split("Product Unit must be input;คุณยังไม่ได้ระบุหน่วยสินค้า", ";")(mTRVB.eVB_TRLang) & vbCrLf
    End If
       
    Dim nStep As Integer
    For nStep = 1 To 4
        If Trim(Me.Controls("otbFCPriceSale" & nStep).Text) = "" Then
            tMsg = tMsg & Split("Product Price " & nStep & " not yet input;ราคาระดับ " & nStep & " ยังไมได้ใส่ข้อมูล", ";")(mTRVB.eVB_TRLang) & vbCrLf
        Else
            If IsNumeric(Me.Controls("otbFCPriceSale" & nStep).Text) = False Then
                tMsg = tMsg & Split("Product Price " & nStep & " Must be number;ราคาระดับ " & nStep & " ต้องเป็นตัวเลข", ";")(mTRVB.eVB_TRLang) & vbCrLf
            End If
        End If
    Next nStep
    
    If tMsg = "" Then
        bValid = True
    Else
        Call SP_SHOWbMessage(tMsg, Exclamation)
    End If
    
    SP_DATbCheckValidate = bValid
    
End Function
Private Function SP_TBLbSaveData() As Boolean
    On Error GoTo Err:
    
    Dim bSuccess As Boolean
    bSuccess = False
    
    SP_TBLbSaveData = bSuccess
    
    Exit Function
    
Err:

    SP_TBLbSaveData = False
    Call SP_SHOWbMessage(Err.Description, Critical)
    
End Function

Private Sub otbCliteria_GotFocus()
    Call SP_CTLxSetFocus(Me.ActiveControl)
End Sub

Private Sub otbFCPriceSale1_GotFocus()
    Call SP_CTLxSetFocus(Me.ActiveControl)
End Sub

Private Sub otbFCPriceSale2_GotFocus()
    Call SP_CTLxSetFocus(Me.ActiveControl)
End Sub

Private Sub otbFCPriceSale3_GotFocus()
    Call SP_CTLxSetFocus(Me.ActiveControl)
End Sub

Private Sub otbFCPriceSale4_GotFocus()
    Call SP_CTLxSetFocus(Me.ActiveControl)
End Sub

Private Sub otbFTPdtBarcode_GotFocus()
    Call SP_CTLxSetFocus(Me.ActiveControl)
End Sub

Private Sub otbFTPdtCode_GotFocus()
    Call SP_CTLxSetFocus(Me.ActiveControl)
End Sub

Private Sub otbFTPdtCode_LostFocus()
    
    If Trim(otbFTPdtCode.Text) = "" Then Exit Sub
    Call SP_DATxLoadFROMDB
    
End Sub

Private Sub otbFTPdtName_GotFocus()
    Call SP_CTLxSetFocus(Me.ActiveControl)
End Sub

Private Sub otbFTRemark_GotFocus()
    bW_CancelTab = True
End Sub

Private Sub otbFTRemark_LostFocus()
    bW_CancelTab = False
End Sub
