VERSION 5.00
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7u.ocx"
Begin VB.Form wTTRMSpn 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Sales Person"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   7815
   WindowState     =   2  'Maximized
   Begin VB.Frame ofaProductInfo 
      Caption         =   "Sales Person Information"
      Height          =   3135
      Left            =   120
      TabIndex        =   14
      Top             =   5040
      Width           =   7575
      Begin VB.TextBox otbFTRemark 
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   7095
      End
      Begin VB.TextBox otbFTSpnName 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   600
         Width           =   5415
      End
      Begin VB.TextBox otbFTSpnCode 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox otbLastUpdate 
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2520
         Width           =   3615
      End
      Begin VB.CommandButton ocmAdd 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton ocmDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton ocmSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label olaFTSpnName 
         Caption         =   "Sales Person Name"
         Height          =   255
         Left            =   1920
         TabIndex        =   18
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label olaFTSpnCode 
         Caption         =   "Sales Person Code"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label olaFTRemark 
         Caption         =   "Remark"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label olbTotalRec 
         Alignment       =   1  'Right Justify
         Caption         =   "Count Records"
         Height          =   255
         Left            =   5520
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton ocmExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton ocmSearch 
      Caption         =   "&Find"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox otbCliteria 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.ComboBox ocbSearch 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VSFlex7UCtl.VSFlexGrid ogdMain 
      Height          =   3855
      Left            =   120
      TabIndex        =   3
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
   Begin VB.Label olaCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Person"
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
      TabIndex        =   13
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label olaSearch 
      BackStyle       =   0  'Transparent
      Caption         =   "Search By"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "wTTRMSpn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oW_DbConn As ADODB.Connection
Dim bW_CancelTab As Boolean
Const tW_TblName As String = "TTRMSpn"

Private Function SP_SQLtForShow() As String

    Dim tSql As String
    tSql = "SELECT FTSpnCode,FTSpnName,FTRemark,"
    tSql = tSql & "FDDateUpd,FTTimeUpd,FTWhoUpd,FDDateIns,FTTimeIns,FTWhoIns "
    tSql = tSql & " FROM " & tW_TblName
    If otbCliteria.Text <> "" Then
        '//if search based FROM all fields or selected field to search
        If ocbSearch.ListIndex > 0 Then
            tSql = tSql & " WHERE " & ocbSearch.Text & " LIKE '%" & otbCliteria.Text & "%'"
        Else
            tSql = tSql & " WHERE ("
            tSql = tSql & "FTSpnCode LIKE '%" & otbCliteria.Text & "%' OR "
            tSql = tSql & "FTSpnName LIKE '%" & otbCliteria.Text & "%' OR "
            tSql = tSql & "FTRemark LIKE '%" & otbCliteria.Text & "%' "
            tSql = tSql & ")"
        End If
    
    End If
    
    tSql = tSql & " ORDER BY 1"
    
    SP_SQLtForShow = tSql
End Function
Private Function SP_TBLoQueryData(Optional ptFieldKey As String = "FTSpnCode") As ADODB.Recordset
    '//Query database for current key input
    Dim oTbl As New ADODB.Recordset
    
    Set oTbl = SP_TBLoGetSalesPerson(oW_DbConn)
    oTbl.Filter = "[" & ptFieldKey & "]='" & otbFTSpnCode.Text & "'"
    
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
    
    Me.otbFTSpnCode.MaxLength = 15
    Me.otbFTSpnName.MaxLength = 100
    Me.otbFTRemark.MaxLength = 100
        
    Call SP_DATxClearForm
    
End Sub
Private Sub SP_DATxClearForm(Optional ptCode As String = "")
    '//clear form
    Me.otbFTSpnCode.Text = ptCode
    Me.otbFTSpnName.Text = ""
    Me.otbFTRemark.Text = ""
    Me.otbLastUpdate.Text = ""
    
    Me.otbFTSpnCode.Locked = False
    Me.otbFTSpnCode.BackColor = vbWhite
    
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
    otbFTSpnCode.SetFocus
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
    tSQLDelete = "DELETE FROM " & tW_TblName & " WHERE FTSpnCode=?"
    
    Dim oCmd As New ADODB.Command
    oCmd.ActiveConnection = oW_DbConn
    oCmd.CommandType = adCmdText
    oCmd.CommandText = tSQLDelete
    
    oCmd.Parameters.Append oCmd.CreateParameter("p1", adVarChar, adParamInput, 15, Me.otbFTSpnCode.Text)
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
        otbFTSpnName.SetFocus
        
End Sub
Private Sub ogdMain_RowColChange()
    
    If ogdMain.Row > 0 Then
        Me.otbFTSpnCode.Text = ogdMain.TextMatrix(ogdMain.Row, 1)
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
        Call SP_DATxClearForm(Me.otbFTSpnCode.Text)
        
        oTbl.Close
        Set oTbl = Nothing
            
        Call SP_SHOWbMessage(mTRMS.tMS_0011, Exclamation)
    
        Exit Sub
        
    End If
    '//if data found then read FROM database and put into controls
    With oTbl
        Me.otbFTSpnCode.Text = "" & .Fields("FTSpnCode").Value
        Me.otbFTSpnName.Text = "" & .Fields("FTSpnName").Value
        Me.otbFTRemark.Text = "" & .Fields("FTRemark").Value
        Me.otbLastUpdate.Text = "Last Update By " & .Fields("FTWhoUpd").Value & " On " & .Fields("FDDateUpd").Value & " " & .Fields("FTTimeUpd").Value
        
    End With

    Me.otbFTSpnCode.Locked = True
    Me.otbFTSpnCode.BackColor = &H8000000F

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
    bValid = True
    'if key blank then ask for input or create new automatically
    If Trim(Me.otbFTSpnCode.Text) = "" Then
        bValid = SP_SHOWbMessage(mTRMS.tMS_0006, Question)
        If Not bValid Then Exit Function
    End If
    'begin check validation data
    bValid = False
    
    Dim tMsg As String
    tMsg = ""
    If Trim(otbFTSpnName.Text) = "" Then
        tMsg = tMsg & Split("Sales Person Name must be input;คุณยังไม่ได้ระบุชื่อพนักงานขาย", ";")(mTRVB.eVB_TRLang) & vbCrLf
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

Private Sub otbFTSpnCode_GotFocus()
    Call SP_CTLxSetFocus(Me.ActiveControl)
End Sub

Private Sub otbFTSpnCode_LostFocus()
    If Trim(otbFTSpnCode.Text) = "" Then Exit Sub
    Call SP_DATxLoadFROMDB
End Sub

Private Sub otbFTSpnName_GotFocus()
    Call SP_CTLxSetFocus(Me.ActiveControl)
End Sub

Private Sub otbFTRemark_GotFocus()
    bW_CancelTab = True
End Sub

Private Sub otbFTRemark_LostFocus()
    bW_CancelTab = False
End Sub
