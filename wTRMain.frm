VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm wTRMain 
   BackColor       =   &H8000000C&
   ClientHeight    =   4080
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7065
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar ostGlobal 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3705
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu omnMain 
      Caption         =   "Main"
      Begin VB.Menu omnCst 
         Caption         =   "Customer"
      End
      Begin VB.Menu omnProduct 
         Caption         =   "Product"
      End
      Begin VB.Menu omnPdtGroup 
         Caption         =   "Product Group"
      End
      Begin VB.Menu omnSpn 
         Caption         =   "Sales Person"
      End
      Begin VB.Menu omnUnit 
         Caption         =   "Unit"
      End
      Begin VB.Menu oln1 
         Caption         =   "-"
      End
      Begin VB.Menu omnExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu omnLanguage 
      Caption         =   "Language"
      Begin VB.Menu omnENLang 
         Caption         =   "English"
      End
      Begin VB.Menu omnTHLang 
         Caption         =   "ไทย"
      End
   End
   Begin VB.Menu omnDbSelect 
      Caption         =   "Database"
      Begin VB.Menu omnAccess 
         Caption         =   "Access"
      End
      Begin VB.Menu omnMSSQL 
         Caption         =   "Sql Server"
      End
   End
End
Attribute VB_Name = "wTRMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Load()
    '//Show Global variable
    Me.Caption = tCS_TRPrjName
    
    Me.ostGlobal.Panels(1).Text = IIf(eVB_TRLang = Thai, "ภาษาไทย", "English")
    Me.ostGlobal.Panels(2).Text = tVB_TRUser
    Me.ostGlobal.Panels(3).Text = SP_GETtConnStr(tCS_TRDbUser, tCS_TRDbPwd, ".", tCS_TRDbName)
    Me.WindowState = vbMaximized
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Check if any form opened alert that user must close all windows first before ask to exit program
    If wTRMain.ActiveForm Is Nothing Then
        If SP_SHOWbMessage(tMS_0002, Confirmation) = False Then
            Cancel = 1
        End If
    Else
        Call SP_SHOWbMessage(tMS_0012, Exclamation)
        Cancel = 1
    End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'clear connection
    If oVB_TRDbCon.State = adStateOpen Then
        oVB_TRDbCon.Close
    End If
    
    Set oVB_TRDbCon = Nothing

End Sub

Private Sub omnAccess_Click()
If Me.ActiveForm Is Nothing Then
    Call SP_SETxVariable(tCS_TRDefUser, ACCESS, eVB_TRLang, True)
    Me.ostGlobal.Panels(3).Text = SP_GETtConnStr(tCS_TRDbUser, tCS_TRDbPwd, ".", tCS_TRDbName)
Else
    Call SP_SHOWbMessage("Cannot Change Database While others window opened;ไม่สามารถเปลี่ยนฐานข้อมูลได้เพราะมีหน้าจอเปิดอยู่", Critical)
End If
End Sub

Private Sub omnCst_Click()
    
    If oVB_TRDbCon.State = adStateOpen Then
        
        Load wTTRMCst
        wTTRMCst.Show
    
    End If

End Sub

Private Sub omnENLang_Click()
    'Change language
    eVB_TRLang = English
    Me.ostGlobal.Panels(1).Text = IIf(eVB_TRLang = Thai, "ภาษาไทย", "English")
End Sub

Private Sub omnExit_Click()
    'if click menu then check all windows must close before exit
    If Me.ActiveForm Is Nothing Then
        Unload Me
    Else
        Call SP_SHOWbMessage(tMS_0012, Exclamation)
    End If
End Sub

Private Sub omnMSSQL_Click()
If Me.ActiveForm Is Nothing Then
    Call SP_SETxVariable(tCS_TRDefUser, SQLServer, eVB_TRLang, True)
    Me.ostGlobal.Panels(3).Text = SP_GETtConnStr(tCS_TRDbUser, tCS_TRDbPwd, ".", tCS_TRDbName)
Else
    Call SP_SHOWbMessage("Cannot Change Database While others window opened;ไม่สามารถเปลี่ยนฐานข้อมูลได้เพราะมีหน้าจอเปิดอยู่", Critical)
End If
End Sub

Private Sub omnPdtGroup_Click()
    If oVB_TRDbCon.State = adStateOpen Then
        
        'Load wTTRMPdtGrp
       ' wTTRMPdtGrp.Show
    
    End If
End Sub

Private Sub omnProduct_Click()
    If oVB_TRDbCon.State = adStateOpen Then
        
        'Load wTTRMPdt
       ' wTTRMPdt.Show
    
    End If
End Sub

Private Sub omnSpn_Click()
    If oVB_TRDbCon.State = adStateOpen Then
        
       ' Load wTTRMSpn
       ' wTTRMSpn.Show
    
    End If
End Sub

Private Sub omnTHLang_Click()
    'Change Language
    eVB_TRLang = Thai
    Me.ostGlobal.Panels(1).Text = IIf(eVB_TRLang = Thai, "ภาษาไทย", "English")
    
End Sub

Private Sub omnUnit_Click()
    If oVB_TRDbCon.State = adStateOpen Then
        
        'Load wTTRMUnit
       ' wTTRMUnit.Show
    
    End If
End Sub
