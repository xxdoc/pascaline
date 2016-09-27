Attribute VB_Name = "comSummary"
Option Explicit

Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Errors As Long
Public pError As Boolean
Public Summary As String
Public ShowSummary As Boolean
Public sFileToRun As String
Public LenIncludes As Long
Public LenProcModules As Long
        
Sub InitSummary()
    Summary = ""
    Errors = 0
    pError = False
    LenIncludes = 0
End Sub

Sub ErrMessage(Text As String)
    pError = True: Errors = Errors + 1
    Dim i As Integer
    LenProcModules = 0
    For i = 1 To UBound(VirtualFiles)
        If VirtualFiles(i).Name = CurrentModule Then
            LenProcModules = LenProcModules + Len("module " & Chr(34) & VirtualFiles(i).Name & Chr(34) & ";" & vbNewLine)
            Exit For
        Else
            LenProcModules = LenProcModules + Len(VirtualFiles(i).Content) + Len("module " & Chr(34) & VirtualFiles(i).Name & Chr(34) & ";" & vbNewLine)
            End If
    Next i
    
    Summary = Summary & "-> " & Text & " [" & CurrentModule & ".Line:" & GetLineNumber(Position - 1 - LenIncludes - LenProcModules) & "]" & vbCrLf
    'If Not IsCmdCompile Then frmMain.FindOrCreateTab CurrentModule
    'If Not IsCmdCompile Then frmMain.Code.ErrorSelectLineByposition Position - 1 - LenIncludes - LenProcModules
End Sub

Sub InfMessage(Text As String)
    If Not IsCmdCompile Then
    frmMain.txtLog.Text = frmMain.txtLog.Text & Replace(Text, vbNewLine, "") & vbCrLf
    End If
    Summary = Summary & Text & vbCrLf
End Sub

Sub WriteSummary(Text As String)
    If ShowSummary = False And pError = False And IsDLL = False And bLibrary = False Then 'ShellExecute frmMain.hwnd, "open", sFileToRun, "", App.Path & "\Binary", 1: Exit Sub
        Dim ProgramID As Long
        sFileToRun = """" & sFileToRun & """"
        ProgramID = Shell(sFileToRun, vbNormalFocus)
        hWndProg = OpenProcess(PROCESS_ALL_ACCESS, False, ProgramID)
        Exit Sub
    End If
    With frmMain
        .txtLog.Text = .txtLog.Text & Text & vbCrLf
        .txtLog.Text = .txtLog.Text & CStr(Errors & " errors occured.") & vbCrLf
    End With
        If pError = True Or IsDLL = True Then frmInfo.cmdAction.Caption = "Back..": frmInfo.cmdAlwaysBack.Visible = False Else frmInfo.cmdAction.Caption = "Run..": frmInfo.cmdAlwaysBack.Visible = True
        
        If bLibrary = True Then frmInfo.cmdAction.Caption = "Back..":
        If Not IsCmdCompile Then frmInfo.Show 1, frmMain
End Sub

Function GetLineNumber(CurrentPosition As Long)
    Dim ActualLine As Integer
    Dim i As Long
    
    ActualLine = 1
    For i = 1 To CurrentPosition
        If Mid$(Source, i, 2) = vbCrLf Then
            ActualLine = ActualLine + 1
        End If
    Next i
    GetLineNumber = ActualLine
End Function



