Attribute VB_Name = "envGeneral"
Option Explicit

Public GridSize As Long
Public ExitCode As Long
Public hWndProg As Long
Public DoCreateObject As Boolean


Public Const SHCNE_ASSOCCHANGED = &H8000000
Public Const SHCNF_IDLIST = &H0
Public Const PROCESS_ALL_ACCESS& = &H1F0FFF
Public Const STILL_ACTIVE& = &H103&
Public Const INFINITE& = &HFFFF
Public Const WM_SYSCOMMAND = &H112

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Sub SHChangeNotify Lib "shell32" (ByVal wEventId As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)

Public DbgCounter As Single

Sub StartCounter()
    DbgCounter = GetTickCount
End Sub

Function EndCounter() As Single
    EndCounter = (GetTickCount - DbgCounter) / 1000
End Function

