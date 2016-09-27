Attribute VB_Name = "envSubClass"
' SubClassing Module constructed from example provided by Garrett Sever
' on www.VisualBasicForum.com

Option Explicit

' A hi-lo type for breaking them up
Private Type HILOWord
  LoWord As Integer
  HiWord As Integer
End Type

' Used to store and retrieve process addresses and handles against a window's handle in
'  the internal Windows database.
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

' Copies memory blocks. I bet you never would have guessed.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
' We need this one to redirect our windows messages... I.e. subclass.
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
' Used to invoke the original/default process for a window.
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' Used to get the parent of the control
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
' Subclassing stuff
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const GWL_WNDPROC As Long = (-4)            ' Used by SetWindowLong to start subclassing
Private Const WM_CREATE As Long = &H1               ' Sent when a window or control is created.
Private Const WM_DESTROY As Long = &H2              ' Sent when a window is being... destroyed.

