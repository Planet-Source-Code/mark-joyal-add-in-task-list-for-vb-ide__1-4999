VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Tasks 
   ClientHeight    =   13980
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   14850
   _ExtentX        =   26194
   _ExtentY        =   24659
   _Version        =   393216
   Description     =   "A List of tasks that need to be completed."
   DisplayName     =   "Task List"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Tasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mcbMenuCommandBar As Office.CommandBarControl
Dim mdocAddIn As docAddIn
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Const guidTasks$ = "AB3075C1-B54F-11d3-941A-00A0CC547B23}"

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    
    Set gVBInstance = Application
    Set gwinWindow = gVBInstance.Windows.CreateToolWindow(AddInInst, "TaskList.docAddIn", "Tasks", guidTasks, mdocAddIn)
    Set mcbMenuCommandBar = AddToAddInCommandBar("Task List")
    Set Me.MenuHandler = gVBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    gwinWindow.Visible = True

End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    mcbMenuCommandBar.Delete
    Set mdocAddIn = Nothing
End Sub

Sub Show()
    gwinWindow.Visible = True
End Sub

Sub hide()
    gwinWindow.Visible = False
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    Set cbMenu = gVBInstance.CommandBars("Tools")
    If cbMenu Is Nothing Then
        Exit Function
    End If
    
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    If CommandBarControl.Caption = "Task List" Then
        Me.Show
    End If
End Sub
