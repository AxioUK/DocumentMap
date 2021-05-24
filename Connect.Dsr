VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   8595
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   14145
   _ExtentX        =   24950
   _ExtentY        =   15161
   _Version        =   393216
   Description     =   $"Connect.dsx":0000
   DisplayName     =   "Document Map Add-In"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "None"
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSafe     =   -1  'True
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Dim ToolbarIcon1 As Office.CommandBarControl
Private WithEvents ToolbarIcon1Events As CommandBarEvents
Attribute ToolbarIcon1Events.VB_VarHelpID = -1

Const CommandBarTitle = "Document Map Window"
Const guidTasks$ = "AB3075C1-B54F-11d3-941A-00A0CC547B23"

Dim cbMenuPopup As CommandBarPopup
Dim cbPopMenuItem1 As CommandBarButton
Private WithEvents cbPopMenuItem1events As CommandBarEvents
Attribute cbPopMenuItem1events.VB_VarHelpID = -1
Dim cbPopMenuItem2 As CommandBarButton
Private WithEvents cbPopMenuItem2events As CommandBarEvents
Attribute cbPopMenuItem2events.VB_VarHelpID = -1
Dim cbPopMenuItem3 As CommandBarButton
Private WithEvents cbPopMenuItem3events As CommandBarEvents
Attribute cbPopMenuItem3events.VB_VarHelpID = -1
Dim cbPopMenuItem4 As CommandBarButton
Private WithEvents cbPopMenuItem4events As CommandBarEvents
Attribute cbPopMenuItem4events.VB_VarHelpID = -1





Dim gwinWindow As VBIDE.Window
Dim mUserDoc As UserDoc

Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    gwinWindow.Visible = False
   
End Sub

Sub Show()
    On Error Resume Next
    
    If (gwinWindow Is Nothing) Then
        Set gwinWindow = VBInstance.Windows.CreateToolWindow(Me.VBInstance.Addins("DocumentMapAddIn.Connect"), "DocumentMapAddIn.UserDoc", "Document Map", "{CDA313D0-AFE0-01A3-B621-591AF24708E1}", mUserDoc)
        If Not (gwinWindow Is Nothing) Then
            Set mUserDoc.VBInstance = VBInstance
            Set mUserDoc.Connect = Me
            mUserDoc.UserDocumentLoad
            gwinWindow.Visible = True
        End If
    Else
        gwinWindow.Visible = True
    End If
    
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    'Save the vb instance
    Set VBInstance = Application
    
    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show

    Else
        'Standard VB AddIn Menu
        Set mcbMenuCommandBar = AddToAddInCommandBar(CommandBarTitle)
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
        
        'Create toolbar buton
        Set ToolbarIcon1 = AddButton(CommandBarTitle, 101)
        If Not (ToolbarIcon1 Is Nothing) Then Set ToolbarIcon1Events = VBInstance.Events.CommandBarEvents(ToolbarIcon1)
        
        OverrideRightMenu
        
    End If
    On Error Resume Next
    
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            Me.Show
        End If
    End If
  
    Exit Sub
    
error_handler:
    MsgBox Err.Description
End Sub




'------------------------------------------------------
'This method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    If Not (mcbMenuCommandBar Is Nothing) Then mcbMenuCommandBar.Delete
    If Not (ToolbarIcon1 Is Nothing) Then ToolbarIcon1.Delete
    
    'Rest of menus
    If Not (cbPopMenuItem1 Is Nothing) Then cbPopMenuItem1.Delete
    If Not (cbPopMenuItem2 Is Nothing) Then cbPopMenuItem2.Delete
    If Not (cbPopMenuItem3 Is Nothing) Then cbPopMenuItem3.Delete
    If Not (cbPopMenuItem4 Is Nothing) Then cbPopMenuItem4.Delete
    If Not (cbMenuPopup Is Nothing) Then cbMenuPopup.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload gwinWindow
    Set gwinWindow = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If
End Sub





'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
Dim cbMenu As Object

On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    Set cbMenuCommandBar = VBInstance.CommandBars.FindControl(msoControlButton, , CommandBarTitle)
    If (cbMenuCommandBar Is Nothing) Then
        'add it to the command bar
        Set cbMenuCommandBar = cbMenu.Controls.Add(1, , , , Temporary:=True)
        'set the caption
        cbMenuCommandBar.caption = sCaption
        cbMenuCommandBar.Tag = sCaption
    End If
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:
End Function

Private Function AddButton(caption As String, resImg As Long) As Office.CommandBarControl
Dim cbMenu As CommandBarButton
Dim orgData As String
Dim ipict As IPictureDisp
On Error GoTo Err:
    
    If VBInstance.CommandBars.Count = 0 Then VBInstance.CommandBars.Add Temporary:=True

    'Check if button already existed (usefull in development)
    Set cbMenu = VBInstance.CommandBars(1).FindControl(1, , caption)
    If (cbMenu Is Nothing) Then
        'Add it
        VBInstance.CommandBars(1).Visible = True
        Set cbMenu = VBInstance.CommandBars(1).Controls.Add(1, , , , Temporary:=True) ', , , VBInstance.CommandBars(2).Controls.Count)
        cbMenu.caption = caption
        cbMenu.Tag = caption 'For finding it
        
        'Set icon
        If (resImg = 0) Then
            'No image
        ElseIf (resImg < 0) Then
            'Native icon
            cbMenu.FaceId = Abs(resImg)
        Else
            'Resource icon
            
            'Backup Clipboard
            orgData = Clipboard.GetText
            Clipboard.Clear
            Set ipict = LoadResPicture(resImg, vbResBitmap)
            If ipict Is Nothing Then
                MsgBox "Failed to load res picture: " & resImg
            Else
                Clipboard.SetData ipict
                cbMenu.PasteFace
            End If
            Set AddButton = cbMenu
            
            Clipboard.Clear
            If Len(orgData) > 0 Then Clipboard.SetText orgData
            
        End If
    End If
    
    Exit Function
Err:
    Debug.Print "AddButton: " & caption & " Err: " & Err.Description
End Function

Private Sub ToolbarIcon1Events_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Show
End Sub


Private Sub OverrideRightMenu()
Dim cb As CommandBar, Tag$

On Error GoTo Err:
    
    If VBInstance.CommandBars.Count = 0 Then VBInstance.CommandBars.Add Temporary:=True

    Set cb = VBInstance.CommandBars("Code Window")
    'Set cb = VBInstance.CommandBars("Code Window (Break)")
    'cb.Controls(9).caption = "W&Alternate"
    
    Tag$ = "DocumentMap.RightClick"
    Set cbMenuPopup = cb.FindControl(msoControlPopup, , Tag$)
    If (cbMenuPopup Is Nothing) Then
        Set cbMenuPopup = cb.Controls.Add(msoControlPopup, , , 4, Temporary:=True)
        cbMenuPopup.BeginGroup = True
        cbMenuPopup.caption = "Document Map"
        cbMenuPopup.Tag = Tag$
    End If
        
    Tag$ = "Add Mark 1"
    Set cbPopMenuItem1 = cbMenuPopup.CommandBar.FindControl(1, , Tag$)
    If (cbPopMenuItem1 Is Nothing) Then
        Set cbPopMenuItem1 = cbMenuPopup.Controls.Add(1, , , , True)
        cbPopMenuItem1.caption = Tag$
        cbPopMenuItem1.Tag = Tag$
        cbPopMenuItem1.FaceId = 39 'Blue Right Arrow
    End If
    Set cbPopMenuItem1events = VBInstance.Events.CommandBarEvents(cbPopMenuItem1)

    Tag$ = "Add Mark 2"
    Set cbPopMenuItem2 = cbMenuPopup.CommandBar.FindControl(1, , Tag$)
    If (cbPopMenuItem2 Is Nothing) Then
        Set cbPopMenuItem2 = cbMenuPopup.Controls.Add(1, , , , True)
        cbPopMenuItem2.caption = Tag$
        cbPopMenuItem2.Tag = Tag$
        cbPopMenuItem2.FaceId = 133 'Green Right Arrow
    End If
    Set cbPopMenuItem2events = VBInstance.Events.CommandBarEvents(cbPopMenuItem2)

    Tag$ = "Add Mark 3"
    Set cbPopMenuItem3 = cbMenuPopup.CommandBar.FindControl(1, , Tag$)
    If (cbPopMenuItem3 Is Nothing) Then
        Set cbPopMenuItem3 = cbMenuPopup.Controls.Add(1, , , , True)
        cbPopMenuItem3.caption = Tag$
        cbPopMenuItem3.Tag = Tag$
        cbPopMenuItem3.FaceId = 1812 'Yellow Right Arrow
    End If
    Set cbPopMenuItem3events = VBInstance.Events.CommandBarEvents(cbPopMenuItem3)

    Tag$ = "Add Line Mark"
    Set cbPopMenuItem4 = cbMenuPopup.CommandBar.FindControl(1, , Tag$)
    If (cbPopMenuItem4 Is Nothing) Then
        Set cbPopMenuItem4 = cbMenuPopup.Controls.Add(1, , , , True)
        cbPopMenuItem4.caption = Tag$
        cbPopMenuItem4.Tag = Tag$
        cbPopMenuItem4.FaceId = 613 'Line
        
    End If
    Set cbPopMenuItem4events = VBInstance.Events.CommandBarEvents(cbPopMenuItem4)

    Exit Sub
Err:
    Debug.Print "OverrideRightMenu Err: " & Err.Description
End Sub

Private Sub cbPopMenuItem1events_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    InsertLine "'*1"
End Sub
Private Sub cbPopMenuItem2events_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    InsertLine "'*2"
End Sub
Private Sub cbPopMenuItem3events_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    InsertLine "'*3"
End Sub
Private Sub cbPopMenuItem4events_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    InsertLine "'*-"
End Sub

'Thanks to yokesee@vbforums for pointing how to do this!
Private Sub InsertLine(s$)
Dim lngStartLine As Long
Dim lngStartColumn As Long
Dim lngEndLine As Long
Dim lngEndColumn As Long
      ' Only add a new line if a code pane is present.
      If VBInstance.CodePanes.Count > 0 Then
          ' Retrieve the starting line of the
          ' selection in active code pane.
          VBInstance.ActiveCodePane.GetSelection lngStartLine, lngStartColumn, lngEndLine, lngEndColumn
          ' Add a line at the location that is
          ' retrieved in the GetSelection statement.
          VBInstance.ActiveCodePane.CodeModule.InsertLines lngStartLine, s$
      End If
End Sub




