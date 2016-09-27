Attribute VB_Name = "envVirtualFiles"
Option Explicit

Public Const EX_ENTRY = 0
Public Const EX_MODULE = 1
Public Const EX_DIALOG = 2

Type TYPE_VIRTUAL_FILE
    Name As String
    Extension As Long
    Content As String
    Used As Boolean
End Type

Public VirtualFiles() As TYPE_VIRTUAL_FILE

Sub InitVirtualFiles()
    ReDim VirtualFiles(0) As TYPE_VIRTUAL_FILE
End Sub

Sub CreateVirtualFile(Name As String, Extension As Long, Content As String)
    ReDim Preserve VirtualFiles(UBound(VirtualFiles) + 1) As TYPE_VIRTUAL_FILE
    VirtualFiles(UBound(VirtualFiles)).Name = Name
    VirtualFiles(UBound(VirtualFiles)).Extension = Extension
    VirtualFiles(UBound(VirtualFiles)).Content = Content
    VirtualFiles(UBound(VirtualFiles)).Used = True
End Sub

Function VirtualFileExists(Name As String) As Boolean
    Dim i As Integer
    For i = 0 To UBound(VirtualFiles)
        If VirtualFiles(i).Name = Name Then
            VirtualFileExists = True
            Exit Function
        End If
    Next i
End Function

Function ChangeVirtualFileName(Name As String, ToName As String) As Boolean
    Dim i As Integer
    If Not Name = ToName Then If VirtualFileExists(ToName) Then MsgBox "'" & ToName & "' is already used": Exit Function
    For i = 0 To UBound(VirtualFiles)
        If VirtualFiles(i).Name = Name Then
            VirtualFiles(i).Name = ToName
            ChangeVirtualFileName = True
            Exit Function
        End If
    Next i
End Function

Function GetVirtualFileExtension(Name As String) As Long
    Dim i As Integer
    For i = 0 To UBound(VirtualFiles)
        If VirtualFiles(i).Name = Name Then
            GetVirtualFileExtension = VirtualFiles(i).Extension
            Exit Function
        End If
    Next i
    GetVirtualFileExtension = -1
End Function

Function GetVirtualFileContent(Name As String)
    Dim i As Integer
    For i = 0 To UBound(VirtualFiles)
        If VirtualFiles(i).Name = Name Then
            GetVirtualFileContent = VirtualFiles(i).Content
            Exit Function
        End If
    Next i
End Function

Function DeleteVirtualFile(Name As String)
    Dim i As Integer
    For i = 0 To UBound(VirtualFiles)
        If VirtualFiles(i).Name = Name Then
            VirtualFiles(i).Used = False
            Exit Function
        End If
    Next i
End Function

Function SetVirtualFileContent(Name As String, Content As String)
    Dim i As Integer
    For i = 0 To UBound(VirtualFiles)
        If VirtualFiles(i).Name = Name Then
            VirtualFiles(i).Content = Content
            Exit Function
        End If
    Next i
End Function

