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
         Format          =   122945539
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

Const tW_TblName As String = "TTRMCst"
Const tW_FldList As String = "FTCstCode,FTCstName,FTCstAddress,FTCstTel,FTCstFax,FTRemark,FTCstStatus,FTCstPriceLv,FDBirthDate,FCCreditLimit,FCCstARBal,FCCstCHQBal"
Const tW_PrimaryKey As String = "[FTCstCode] ='{0}'"
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandle:
    'change enter to tab to move between objects
    If KeyCode = 13 And bW_CancelTab = False Then
    
        Dim oShell As Object
        Set oShell = CreateObject("WScript.Shell")
        oShell.SendKeys "{TAB}"
        Set oShell = Nothing
        Exit Sub
    End If
    
    Select Case KeyCode
    Case vbKeyF2
        Call ocmAdd_Click
    Case vbKeyF3
        Call ocmSearch_Click
    Case vbKeyF8
        Call ocmDelete_Click
    Case vbKeyF9
        Call ocmSave_Click
    Case vbKeyEscape
        Call ocmExit_Click
    End Select
    
    Exit Sub
    
ErrHandle:
    
End Sub
Private Sub Form_Load()
    '//prepare connection
    Set oW_DbConn = oVB_TRDbCon
    '//show grid and set defaults
    Call W_SETxGridData
    Call W_SETxProperty
    Call W_SETxClearData

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If SP_SHOWbMessage(tMS_0003, Question) = False Then
        
        Cancel = 1
    
    End If

End Sub

Private Sub ocmSearch_Click()

    Call W_SETxGridData
    
End Sub
Private Sub ogdMain_Click()
    
        Call ogdMain_RowColChange
        otbFTCstName.SetFocus
            
End Sub
Private Sub ogdMain_RowColChange()
    
    If ogdMain.Row > 0 Then

        Me.otbFTCstCode.Text = ogdMain.TextMatrix(ogdMain.Row, 1)
        Call W_LOADxData
        
    End If

End Sub
Private Sub otbCliteria_GotFocus()
    
    Call SP_SETxCtlSelect(Me.ActiveControl)
    
End Sub
Private Sub otbFCCreditLimit_GotFocus()

    Call SP_SETxCtlSelect(Me.ActiveControl)
    
End Sub
Private Sub otbFTCstAddress_GotFocus()

    bW_CancelTab = True
    'Call SP_SETxCtlSelect(Me.ActiveControl)

End Sub
Private Sub otbFTCstAddress_LostFocus()
    
    bW_CancelTab = False

End Sub
Private Sub otbFTCstCode_GotFocus()
    
    Call SP_SETxCtlSelect(Me.ActiveControl)

End Sub
Private Sub otbFTCstCode_LostFocus()
        
        If Trim(otbFTCstCode.Text) = "" Then Exit Sub
        Call W_LOADxData

End Sub
Private Sub otbFTCstFax_GotFocus()
    
    Call SP_SETxCtlSelect(Me.ActiveControl)

End Sub
Private Sub otbFTCstName_GotFocus()
    
    Call SP_SETxCtlSelect(Me.ActiveControl)

End Sub
Private Sub otbFTCstTel_GotFocus()
    
    Call SP_SETxCtlSelect(Me.ActiveControl)

End Sub
Private Sub otbFTRemark_GotFocus()
    
    bW_CancelTab = True
    'Call SP_SETxCtlSelect(Me.ActiveControl)

End Sub
Private Sub otbFTRemark_LostFocus()
    bW_CancelTab = False
End Sub

Private Sub ockNonActive_Click()

    Call W_SETxGridData
    
End Sub
Private Sub ocmAdd_Click()

    Call W_SETxClearData
    otbFTCstCode.SetFocus
    
End Sub
Private Sub ocmDelete_Click()
    '//delete data
    If SP_SHOWbMessage(tMS_0005, Confirmation) = False Then Exit Sub
    If W_DELbData() = True Then
    
        Call W_SETxClearData
        Call W_SETxGridData
        Call SP_SHOWbMessage(tMS_0010, Exclamation)
    End If

End Sub
Private Sub ocmExit_Click()

    Unload Me
    
End Sub
Private Sub ocmSave_Click()
    '//Save Data
    If W_CHKbInputData() = False Then Exit Sub
    If SP_SHOWbMessage(tMS_0004, Question) = False Then Exit Sub
    '//generate new code here if user input blank
    If Trim(otbFTCstCode.Text) = "" Then
        otbFTCstCode.Text = SP_GETtNewID(oW_DbConn, "Cst")
    End If
    'If W_TBLbUpdateData() = True Then     'using Recordset.update instead of SQL Command
    If W_TBLbSaveData() = True Then        'using SQL Command for saving data
        
        Call W_SETxGridData
        Call W_LOADxData
        Call SP_SHOWbMessage(tMS_0007, Information)
        
    End If

End Sub
Private Function W_DELbData() As Boolean
'----------------------------------------------------------
'   *Puttipong 2017-09-11
'   ใชสำหรับลบช้อมูลจาก Database
'----------------------------------------------------------

    On Error GoTo ErrHandle:
    
    '//delete using parameters
    Dim tSQLDelete As String
    tSQLDelete = "DELETE FROM " & tW_TblName & "  WHERE " & W_SQLtPrimaryKey
           
    SP_EXECoSQL oW_DbConn, tSQLDelete
           
    W_DELbData = True
    
    Exit Function

ErrHandle:
    
    W_DELbData = False
    Call SP_SHOWbMessage(Err.Description, Critical)
End Function
Private Function W_CHKbInputData() As Boolean
'----------------------------------------------------------
'   *Puttipong 2017-09-11
'   ตรวจสอบข้อมูลก่อนเซฟ
'   Call:   -
'
'   Ret:    True = พร้อมบันทึกข้อมูล
'               False = ยังมีข้อมูลที่ต้องตรวจสอบอยู่
'
'----------------------------------------------------------
    Dim bValid As Boolean
    bValid = True
    'if key blank then ask for input or create new automatically
    If Trim(Me.otbFTCstCode.Text) = "" Then
        bValid = SP_SHOWbMessage(tMS_0006, Question)
        If Not bValid Then Exit Function
    End If
    'begin check validation data
    bValid = False
    
    Dim tMsg As String
    tMsg = ""
    
    If Trim(Me.otbFTCstName.Text) = "" Then
        tMsg = tMsg & Split(tMS_0013, ";")(eVB_TRLang) & vbCrLf
    End If
    
    If Trim(Me.otbFTCstAddress.Text) = "" Then
        tMsg = tMsg & Split(tMS_0014, ";")(eVB_TRLang) & vbCrLf
    End If
    
    If Trim(Me.otbFTCstTel.Text) = "" Then
        tMsg = tMsg & Split(tMS_0015, ";")(eVB_TRLang) & vbCrLf
    Else
        If IsNumeric(Me.otbFTCstTel) = False Then
            tMsg = tMsg & Split(tMS_0016, ";")(eVB_TRLang) & vbCrLf
        End If
    End If
    
    If Trim(Me.otbFTCstFax.Text) = "" Then
        tMsg = tMsg & Split(tMS_0017, ";")(eVB_TRLang) & vbCrLf
    Else
        If IsNumeric(Me.otbFTCstFax) = False Then
            tMsg = tMsg & Split(tMS_0018, ";")(eVB_TRLang) & vbCrLf
        End If
    End If
    
    If Year(Me.odtFDBirthDate.Value) = 1900 Then
        tMsg = tMsg & Split(tMS_0019, ";")(eVB_TRLang) & vbCrLf
    End If
    
    If IsNumeric(Me.otbFCCreditLimit.Text) = False Then
        tMsg = tMsg & Split(tMS_0020, ";")(eVB_TRLang) & vbCrLf
    End If
    
    If tMsg = "" Then
        bValid = True
    Else
        Call SP_SHOWbMessage(tMsg, Exclamation)
    End If
    
    W_CHKbInputData = bValid
    
End Function
Private Function W_TBLbSaveData() As Boolean
    On Error GoTo ErrHandle:
'----------------------------------------------------------
'   *Puttipong 2017-09-11
'   บันทึกข้อมูลลูกค้าลง Database
'   Call:   -
'
'   Ret:    True = บันทึกสำเร็จ
'               False = บันทึกไม่สำเร็จ
'
'----------------------------------------------------------

    Dim bSuccess As Boolean
    bSuccess = False

    '//read structure
    Dim oCtl As Control
    
    Dim tSQLIns As String
    Dim tSQLInsFld As String
    Dim tSQLInsVal As String
    Dim tSQLUpd As String
    Dim tVal As String
    
    tSQLIns = SP_GETtSQLDefault(Insert)
    tSQLInsFld = "" & Split(tSQLIns, ";")(0)
    tSQLInsVal = "" & Split(tSQLIns, ";")(1)
    
    tSQLUpd = SP_GETtSQLDefault(Update)
                    
    For Each oCtl In Me.Controls
        If oCtl.Tag <> "" Then
            tVal = ""
            Select Case TypeName(oCtl)
                Case "TextBox"
                    tVal = SP_GETtSQLFormat(oCtl.Text, Text)
                Case "ComboBox"
                    tVal = SP_GETtSQLFormat(oCtl.Text, Text)
                Case "DataCombo"
                    tVal = SP_GETtSQLFormat(oCtl.BoundText, Text)
                Case "DTPicker"
                    tVal = SP_GETtSQLFormat(oCtl.Value, Date)
                Case "ListBox"
                    tVal = SP_GETtSQLFormat(oCtl.Text, Text)
                Case "OptionButton"
                    tVal = SP_GETtSQLFormat(IIf(oCtl.Value = True, 1, 2), Text)
                Case "CheckBox"
                    tVal = SP_GETtSQLFormat(IIf(oCtl.Checked = vbChecked, 1, 2), Text)
            End Select
            tSQLInsFld = tSQLInsFld & "," & oCtl.Tag
            tSQLInsVal = tSQLInsVal & "," & tVal
            tSQLUpd = tSQLUpd & "," & oCtl.Tag & "=" & tVal
        End If
     Next
    '//Generate SQL Command
    tSQLIns = "INSERT INTO " & tW_TblName & " (" & tSQLInsFld & ") "
    tSQLIns = tSQLIns & vbCrLf & "VALUES (" & tSQLInsVal & ")"
    
    tSQLUpd = "UPDATE " & tW_TblName & "  SET " & vbCrLf & tSQLUpd
    tSQLUpd = tSQLUpd & vbCrLf & " WHERE " & W_SQLtPrimaryKey()
    
    '//Try to insert first and then update if insert failed
    Dim oTbl As ADODB.Recordset
    Set oTbl = W_GEToData()
    If oTbl.EOF = False Then
        bSuccess = SP_EXECbSQL(oW_DbConn, tSQLUpd)
    Else
        bSuccess = SP_EXECbSQL(oW_DbConn, tSQLIns)
    End If
    oTbl.Close
    
    W_TBLbSaveData = bSuccess
    
    Exit Function
    
ErrHandle:

    W_TBLbSaveData = False
    Call SP_SHOWbMessage(Err.Description, Critical)
    
End Function
Private Function W_TBLbUpdateData() As Boolean
'----------------------------------------------------------
'   *Puttipong 2017-09-11
'   ใช้ Save ข้อมูลโดยการ Assign ข้อมูลลง Recordset ตรงๆ โดยไม่ต้อง Run คำสั่ง SQL Insert/update
'----------------------------------------------------------
    '//using Recordset update instread of SQL Command for prevent String exception
    On Error GoTo ErrHandle:
    '//Generate new code
    If otbFTCstCode.Text = "" Then
        otbFTCstCode.Text = SP_GETtNewID(oW_DbConn, "Cst")
    End If
    '//Read FROM Database and filter for key input
    Dim oRs As ADODB.Recordset
    Set oRs = W_GEToData()
    
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
        .Fields("FTCstStatus").Value = IIf(orbActive.Value = True, 1, 2)
        .Fields("FTCstPriceLv").Value = olbFTCstPriceLv.ListIndex + 1
        .Fields("FTRemark").Value = otbFTRemark.Text
        .Fields("FCCreditLimit").Value = otbFCCreditLimit.Text
        .Fields("FDBirthDate").Value = Format(CDate(odtFDBirthDate.Value), "yyyy-MM-dd")
        
        .Update
        .Close
    
    End With
    '//update for log
    Call SP_SETxLogTbl(oAction, tW_TblName, W_SQLtPrimaryKey, oW_DbConn)

    W_TBLbUpdateData = True
    
    Exit Function

ErrHandle:

    Call SP_SHOWbMessage(Err.Description, Critical)
    W_TBLbUpdateData = False

End Function
Private Sub W_SETxClearData(Optional ptCode As String = "")
'----------------------------------------------------------
'   *Puttipong 2017-09-11
'   ใชล้างหน้าจอและกำหนดค่าเริ่มต้นของ Control ใหม่
'   Call:
'       pCode = รหัสที่จะแสดงข้อมูล ใส่ค่าว่างได้
'
'----------------------------------------------------------
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
Private Function W_GEToData() As ADODB.Recordset
'----------------------------------------------------------
'   *Puttipong 2017-09-11
'   ใช้ query ข้อมูลจาก primary key
'
'   Ret :
'       RecordSet ของข้อมูล
'----------------------------------------------------------
On Error GoTo ErrHandle:

    '//Query database for current key input
    Dim oTbl As New ADODB.Recordset
    
    Set oTbl = SP_GEToData(oW_DbConn, "Cst")
    oTbl.Filter = W_SQLtPrimaryKey()
    
    Set W_GEToData = oTbl
    Exit Function
ErrHandle:
    Set W_GEToData = New ADODB.Recordset
End Function
Private Sub W_LOADxData()
'----------------------------------------------------------
'   *Puttipong 2017-09-11
'   ใช้ query ข้อมูลจาก primary key แล้วนำมาแสดงใน form
'----------------------------------------------------------
    On Error GoTo ErrHandle:
        
    '//Read FROM data when user input key
    Dim oTbl As ADODB.Recordset
    Set oTbl = W_GEToData()
    
    If oTbl.EOF = True Then
    
        '//if not found then clear form and show msgbox data not found
        Call W_SETxClearData(Me.otbFTCstCode.Text)
        
        oTbl.Close
        Set oTbl = Nothing
            
        Call SP_SHOWbMessage(tMS_0011, Exclamation)
    
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
        Me.orbActive.Value = IIf(.Fields("FTCstStatus").Value = "1", True, False)
        Me.orbInactive.Value = IIf(.Fields("FTCstStatus").Value = "1", False, True)
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
    
ErrHandle:

    If oTbl.State = 1 Then oTbl.Close
    Set oTbl = Nothing
    
    Call SP_SHOWbMessage(Err.Description, Critical)
    
End Sub
Private Sub W_SETxProperty()
'----------------------------------------------------------
'   *Puttipong 2017-09-11
'   ใช้กำหนดค่าเริ่มต้นและ property ต่างๆ ของ Control ใน Form
'----------------------------------------------------------
    Me.otbFTCstCode.MaxLength = 15
    Me.otbFTCstName.MaxLength = 100
    Me.otbFTCstAddress.MaxLength = 100
    Me.otbFTRemark.MaxLength = 100
    Me.otbFTCstTel.MaxLength = 10
    Me.otbFTCstFax.MaxLength = 10
    Me.otbFCCreditLimit.MaxLength = 15
    
    Me.otbFTCstCode.Tag = "FTCstCode"
    Me.otbFTCstName.Tag = "FTCstName"
    Me.otbFTCstAddress.Tag = "FTCstAddress"
    Me.otbFTRemark.Tag = "FTRemark"
    Me.otbFTCstTel.Tag = "FTCstTel"
    Me.otbFTCstFax.Tag = "FTCstFax"
    Me.otbFCCreditLimit.Tag = "FCCreditLimit"
    Me.odtFDBirthDate.Tag = "FDBirthDate"
    Me.olbFTCstPriceLv.Tag = "FTCstPriceLv"
    Me.orbActive.Tag = "FTCstStatus"
    
    Call W_SETxCliteria
    
End Sub
Private Function W_SQLtPrimaryKey() As String
'----------------------------------------------------------
'   *Puttipong 2017-09-11
'   ใช้อ่านค่า Key จาก Form เพื่อใช้เป็น WHERE Clause ตอน Query ข้อมูลใน Database
'   Call -
'   Ret : String เป็นคำสั่ง SQL ที่ไม่มี Where จากตัวแปร tW_PrimaryKey
'----------------------------------------------------------
    Dim tSQL As String
    tSQL = Replace(tW_PrimaryKey, "{0}", otbFTCstCode.Text)
    
    W_SQLtPrimaryKey = tSQL
    
End Function
Private Function W_SQLtSearch() As String
'----------------------------------------------------------
'   *Puttipong 2017-09-11
'   ใช้สร้างคำสั่ง SQL ในการ Query ข้อมูลตามที่ User ป้อนเงื่อนไขการค้นหา
'----------------------------------------------------------

    Dim tSQL As String
    '//ใส่เงื่อนไขที่จะค้นหาพื้นฐานของ form ถ้ามี
    If ockNonActive.Value = vbChecked Then
        tSQL = " WHERE FTCstStatus='2' "
    Else
        tSQL = " WHERE FTCstStatus='1' "
    End If
    '//ถ้า user ไม่เลือกฟิลด์ข้อมูลจะเป็นการค้นหาทั้งหมด หรือไม่ก้เฉพาะฟิลด์ที่ถูกเลือก
    If otbCliteria.Text <> "" Then
        If ocbSearch.ListIndex > 0 Then
            tSQL = tSQL & " AND " & ocbSearch.Text & " LIKE '%" & otbCliteria.Text & "%'"
        Else
            tSQL = tSQL & " AND ("
            
            Dim tCliteria As String
            Dim nIdx As Integer
            For nIdx = 1 To ocbSearch.ListCount - 1
                If tCliteria <> "" Then tCliteria = tCliteria & " OR "
                tCliteria = tCliteria & ocbSearch.List(nIdx) & " LIKE '%" & otbCliteria.Text & "%'"
            Next nIdx
            
            tSQL = tSQL & tCliteria
            tSQL = tSQL & ")"
        End If
    
    End If
    tSQL = tSQL & " ORDER BY 1"
    
    W_SQLtSearch = tSQL
    
End Function
Private Function W_SQLtGrid() As String
'----------------------------------------------------------
'   *Puttipong 2017-09-11
'   ใช้สร้างคำสั่ง SQL สำหรับแสดงผลใน Grid ข้อมูลตามที่ User ค้นหาตามเงื่อนไขที่ระบุใน Form
'   Call -
'   Ret : String เป็นคำสั่ง SQL ที่มาจากการอ่านค่าข้อมูลใน Form แล้ว
'----------------------------------------------------------
    Dim tSQL As String
    
    tSQL = "SELECT " & tW_FldList
    tSQL = tSQL & ",FDDateUpd,FTTimeUpd,FTWhoUpd,FDDateIns,FTTimeIns,FTWhoIns"
    tSQL = tSQL & vbCrLf & " FROM  " & tW_TblName
    tSQL = tSQL & vbCrLf & W_SQLtSearch()
    
    W_SQLtGrid = tSQL
End Function
Private Sub W_SETxCliteria()
'----------------------------------------------------------
'   *Puttipong 2017-09-11
'   ใช้กำหนดค่าให้ Control ที่จะให้  User ค้นหาข้อมูล
'----------------------------------------------------------
    
    '//โหลด Combo ชื่อฟิลด์เพื่อให้ user เลือก
    '//โดยอาจซ่อนบางฟิลด์ก็ได้โดยใส่ในตัวแปร tExceptField
    If ocbSearch.ListCount = 0 Then
        Dim tFld As String
        Dim nIdx As Integer
        
        Dim tExceptField As String
        tExceptField = "FTCstStatus,FCCstARBal,FCCstCHQBal"
        
        ocbSearch.AddItem "(All Data)"
        
        Dim aList As Variant
        aList = Split(tW_FldList, ",")
        For nIdx = 0 To UBound(aList) - 1
            If InStr(1, tExceptField, aList(nIdx)) <= 0 Then
                ocbSearch.AddItem aList(nIdx)
            End If
        Next nIdx
        
        ocbSearch.ListIndex = 0
    
    End If
    
End Sub
Private Sub W_SETxGridData()
'----------------------------------------------------------
'   *Puttipong 2017-09-11
'   ใช้สำหรับ Load ข้อมูลที่ Query ได้มาแสดงใน Grid
'----------------------------------------------------------

    On Error GoTo ErrHandle:
    
    Dim oTbl As ADODB.Recordset
    Set oTbl = SP_EXECoSQL(oW_DbConn, W_SQLtGrid)
    
    Set ogdMain.DataSource = oTbl
    ogdMain.SelectionMode = flexSelectionByRow
    
    olbTotalRec.Caption = "Found =" & oTbl.RecordCount & " Record(s)"
    
    Exit Sub

ErrHandle:

    Call SP_SHOWbMessage(Err.Description, Critical)
    
End Sub
