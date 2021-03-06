Attribute VB_Name = "Module1"
Option Explicit

' Store WndProcs
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" ( _
                ByVal hWnd As Long, _
                ByVal lpString As String) As Long

Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" ( _
                ByVal hWnd As Long, _
                ByVal lpString As String, _
                ByVal hData As Long) As Long

Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" ( _
                ByVal hWnd As Long, _
                ByVal lpString As String) As Long

' Hooking
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
                ByVal lpPrevWndFunc As Long, _
                ByVal hWnd As Long, _
                ByVal Msg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
                ByVal hWnd As Long, _
                ByVal Msg As Long, _
                wParam As Any, _
                lParam As Any) As Long

' Position Checking
Private Declare Function GetWindowRect Lib "user32" ( _
                ByVal hWnd As Long, _
                lpRect As RECT) As Long
                
Private Declare Function GetParent Lib "user32" ( _
                ByVal hWnd As Long) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A
Private Const CB_GETDROPPEDSTATE = &H157

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

' ================================================
' Check Messages
' ================================================
Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo Err
Dim MouseKeys As Long
Dim Rotation As Long
Dim Xpos As Long
Dim Ypos As Long
Dim fFrm As Form

    Select Case Lmsg
    
      Case WM_MOUSEWHEEL
      
        Rotation = wParam / 65536
        'Not needed
        'MouseKeys = wParam And 65535
        'Xpos = lParam And 65535
        'Ypos = lParam / 65536
        
        Set fFrm = GetForm(Lwnd)
        If fFrm Is Nothing Then
            'Nothing...
        Else
            ' Form active and it's over the form so fire mousewheel
            fFrm.MouseWheel MouseKeys, Rotation, Xpos, Ypos
        End If
    End Select
    
    WindowProc = CallWindowProc(GetProp(Lwnd, "PrevWndProc"), Lwnd, Lmsg, wParam, lParam)
Err:
End Function

' ================================================
' Hook / UnHook
' ================================================
Public Sub WheelHook(ByVal hWnd As Long)
  On Error Resume Next
  If (InIDE) Then Exit Sub
  SetProp hWnd, "PrevWndProc", SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub WheelUnHook(ByVal hWnd As Long)
  On Error Resume Next
  If (InIDE) Then Exit Sub
  SetWindowLong hWnd, GWL_WNDPROC, GetProp(hWnd, "PrevWndProc")
  RemoveProp hWnd, "PrevWndProc"
End Sub

' ================================================
' Window Checks
' ================================================
Public Function IsOver(ByVal hWnd As Long, ByVal lX As Long, ByVal lY As Long) As Boolean
  Dim rectCtl As RECT
  GetWindowRect hWnd, rectCtl
  With rectCtl
    IsOver = (lX >= .Left And lX <= .Right And lY >= .Top And lY <= .Bottom)
  End With
End Function

Private Function GetForm(ByVal hWnd As Long) As Form
  For Each GetForm In Forms
    If GetForm.hWnd = hWnd Then Exit Function
  Next GetForm
  Set GetForm = Nothing
End Function

Public Function InIDE() As Boolean
    InIDE = CBool(App.LogMode = 0)
End Function


