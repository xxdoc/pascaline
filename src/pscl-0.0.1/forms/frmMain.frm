VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Pascaline"
   ClientHeight    =   6810
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   11220
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox tbExplorer 
      Align           =   4  'Align Right
      Height          =   5775
      Left            =   9285
      ScaleHeight     =   5715
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   0
      Width           =   1935
      Begin VB.ListBox lstProject 
         Height          =   1035
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox tbar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   250
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   11220
      TabIndex        =   3
      Top             =   6555
      Width           =   11220
      Begin Pascaline.ProgressBar progBar 
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   1575
         _extentx        =   2778
         _extenty        =   450
         appearance      =   1
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ready"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   0
         Width           =   675
      End
   End
   Begin VB.PictureBox tbLog 
      Align           =   2  'Align Bottom
      Height          =   780
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   11160
      TabIndex        =   1
      Top             =   5775
      Width           =   11220
      Begin VB.TextBox txtLog 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "frmMain.frx":0E42
         Top             =   0
         Width           =   11175
      End
   End
   Begin VB.Timer tmrApplicationRuntime 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2640
      Top             =   600
   End
   Begin Pascaline.Editor txtEditor 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _extentx        =   3413
      _extenty        =   2566
      text            =   ""
      font            =   "frmMain.frx":0E55
   End
   Begin Pascaline.UnicodeDialog comdlg 
      Left            =   2160
      Top             =   600
      _extentx        =   847
      _extenty        =   847
      fileflags       =   2621444
      folderflags     =   323
      filecustomfilter=   "frmMain.frx":0E81
      filedefaultextension=   "frmMain.frx":0EA1
      filefilter      =   "frmMain.frx":0EC1
      fileopentitle   =   "frmMain.frx":0EE1
      filesavetitle   =   "frmMain.frx":0F01
      foldermessage   =   "frmMain.frx":0F21
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewProject 
         Caption         =   "&New Project..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Project..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuspec1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveProject 
         Caption         =   "&Save Project"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAsProject 
         Caption         =   "Save &As Project..."
      End
      Begin VB.Menu mnuspec2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "C&opy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuspec3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "Replace"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuCompiler 
      Caption         =   "&Compiler"
      Begin VB.Menu mnuLinkRun 
         Caption         =   "Link && Run"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuspec5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompile 
         Caption         =   "Compile..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelps 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuDocumentation 
         Caption         =   "Documentation"
      End
      Begin VB.Menu mnuspec4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Pascaline"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public isDirty As Boolean
Public lModuleID As Long
Public lResourceID As Long
Public AutoTemplates As Boolean
Public RunEnabled As Boolean

Private Sub Form_Load()

    isDirty = False: RunEnabled = True
    CheckForAssociation
    
    If GetSetting("Pascaline", "Settings", "Summary", False) = True Then
        ShowSummary = True
    Else
        ShowSummary = False
    End If
    
    AutoTemplates = GetSetting("Pascaline", "Settings", "AutoTemplates", True)
    
    
    If GetSetting("Pascaline", "Settings", "Maximized", False) = True Then
        frmMain.WindowState = vbMaximized
    ElseIf GetSetting("Pascaline", "Settings", "Minimized", False) = True Then
        'Nothing
    Else
        frmMain.Move GetSetting("Pascaline", "Settings", "Left", 400), _
                     GetSetting("Pascaline", "Settings", "Top", 400), _
                     GetSetting("Pascaline", "Settings", "Width", 11000), _
                     GetSetting("Pascaline", "Settings", "Height", 7200)
    End If
    
    frmMain.Show: DoEvents
    
    If Command$ <> "" And Not Mid(Command$, 1, 2) = "/c" Then
        comdlg.FileName = Command$: DoEvents
        OpenProject True
        'SelectEntryFile
    Else
        NewProject True
        'SelectEntryFile
    End If

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtEditor.Height = Me.ScaleHeight - tbLog.Height - tbar.Height
    txtEditor.Width = Me.ScaleWidth - tbExplorer.Width
    txtLog.Height = tbLog.ScaleHeight
    txtLog.Width = Me.ScaleWidth - 50
    lstProject.Height = tbExplorer.Height
    lstProject.Width = tbExplorer.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "Pascaline", "Settings", "AutoTemplates", AutoTemplates
    If frmMain.WindowState = vbMaximized Then
        SaveSetting "Pascaline", "Settings", "Maximized", True
    ElseIf frmMain.WindowState = vbMinimized Then
        SaveSetting "Pascaline", "Settings", "vbMinimized", True
    Else
        SaveSetting "Pascaline", "Settings", "Maximized", False
        SaveSetting "Pascaline", "Settings", "Left", frmMain.Left
        SaveSetting "Pascaline", "Settings", "Top", frmMain.Top
        SaveSetting "Pascaline", "Settings", "Width", frmMain.Width
        SaveSetting "Pascaline", "Settings", "Height", frmMain.Height
    End If
    If CheckUnsaved = True Then Cancel = True: Exit Sub
    End
End Sub

Private Sub lstProject_Click()
    Dim NodeText As String
    
    If lstProject.ListCount = 0 Then Exit Sub
    
    NodeText = lstProject.List(lstProject.ListIndex)
    NodeText = Replace(NodeText, "[MODULE] - ", "")
    NodeText = Replace(NodeText, "[ENTRY] - ", "")
    NodeText = Replace(NodeText, "[RESOURCE] - ", "")
    Debug.Print lstProject.List(lstProject.ListIndex)
    Debug.Print "result : " & NodeText
    frmMain.txtEditor.Text = GetVirtualFileContent(NodeText)
    frmMain.txtEditor.Enabled = True
    If GetVirtualFileExtension(NodeText) = EX_DIALOG Then
        'Show Design Form
        
    Else
        'Hide Design Form
    End If
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show , Me
End Sub

Private Sub mnuCompile_Click()
   MakeExecuteable
End Sub

Private Sub mnuCopy_Click()
    Clipboard.SetText txtEditor.SelText
End Sub

Private Sub mnuCut_Click()
    Clipboard.SetText txtEditor.SelText
    txtEditor.Text = Replace(txtEditor.Text, txtEditor.SelText, "")
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuHelps_Click()
Shell "explorer.exe vaxza.ga/pascaline/help", vbNormalFocus
End Sub

Private Sub mnuLinkRun_Click()
    Dim CompilePath As String
    
            If comdlg.FileName = "" Then
                If MsgBox("File is not saved. Do you want to save it now?", _
                          vbInformation + vbYesNo, "Pascaline") = vbYes Then
                    If (SaveProject(True) = False) Then
                        GoTo CannotCompileNotSaved
                    End If
                Else
                    GoTo CannotCompileNotSaved
                End If
            End If
           
            Dim i As Integer
            
            RunEnabled = True
            tmrApplicationRuntime.Enabled = True:
            frmMain.Caption = Right(comdlg.FileName, Len(comdlg.FileName) - InStrRev(comdlg.FileName, "\", -1, vbTextCompare)) & " - Pascaline Compiler 3.0.1 [Compiling..]"
            Screen.MousePointer = 13
           
            CompilePath = Left(comdlg.FileName, Len(comdlg.FileName) - Len(Right(comdlg.FileName, Len(comdlg.FileName) - InStrRev(comdlg.FileName, "\", -1, vbTextCompare))))
            Compile CompilePath & Switch(Right(comdlg.FileName, Len(comdlg.FileName) - InStrRev(comdlg.FileName, "\", -1, vbTextCompare)) <> "", Mid$(Right(comdlg.FileName, Len(comdlg.FileName) - InStrRev(comdlg.FileName, "\", -1, vbTextCompare)), 1, InStr(1, Right(comdlg.FileName, Len(comdlg.FileName) - InStrRev(comdlg.FileName, "\", -1, vbTextCompare)), ".", vbTextCompare)) & "exe", Right(comdlg.FileName, Len(comdlg.FileName) - InStrRev(comdlg.FileName, "\", -1, vbTextCompare)) = "", "noname.exe"), True
            Screen.MousePointer = vbNormal
            AppType = 0 'because stop button checks this
            'RunEnabled = False

CannotCompileNotSaved:
End Sub

Private Sub mnuNewProject_Click()
    NewProject True
End Sub

Private Sub mnuOpen_Click()
    OpenProject
End Sub

Private Sub mnuPaste_Click()
    txtEditor.SelText = Clipboard.GetText
End Sub

Private Sub mnuReplace_Click()
frmReplace.Show
End Sub

Private Sub mnuSaveAsProject_Click()
    SaveProject True
End Sub

Private Sub mnuSaveProject_Click()
    SaveProject False
End Sub

Private Sub mnuSelectAll_Click()
    txtEditor.SelStart = 0
    txtEditor.SelLength = Len(txtEditor.Text)
End Sub

Function SaveProject(SaveAs As Boolean) As Boolean

    On Error GoTo SaveCancelError
    
     Dim comsave As UnicodeFileDialog
     
    SaveProject = False
    
    If SaveAs = False Then If comdlg.FileName <> "" Then GoTo SaveNow
    
    Set comsave = New UnicodeFileDialog
    
    With comsave
        .Filter = "Pascaline Worksheet Project (*.pwp)|*.pwp|All Files (*.*)|*.*"
        .ShowSave Me.hWnd
    
        If Dir(.FileName) <> "" Then
            If MsgBox("File already exist! Overwrite it anyway?", vbQuestion + vbYesNo) = vbNo Then
                Exit Function
            Else
                GoTo SaveNow
            End If
        End If
    End With
    
If comsave.FileName <> "" Then
    comdlg.FileName = comsave.FileName
    comsave.FileName = ""
End If
SaveNow:

Dim i As Long: Dim FileNum As Long
    
    FileNum = FreeFile
    
    If Dir(comdlg.FileName) <> "" Then Kill comdlg.FileName
    
    Open comdlg.FileName For Binary As #FileNum
        
        Put #FileNum, , UBound(VirtualFiles)
        For i = 1 To UBound(VirtualFiles): Put #FileNum, , VirtualFiles(i): Next i
    
    Close #FileNum
    
    frmMain.Caption = Right(comdlg.FileName, Len(comdlg.FileName) - InStrRev(comdlg.FileName, "\", -1, vbTextCompare)) & " - Pascaline"
    isDirty = False
    SaveProject = True

SaveCancelError:
    SaveProject = False
    Exit Function
End Function
Public Sub OpenProject(Optional NoCD As Boolean)
    On Error GoTo CancelOpen
    
    If NoCD = True Then GoTo lNoCD
    
    If CheckUnsaved = True Then Exit Sub
    
    
    With comdlg
        .FileFilter = "Pascaline Worksheet Project (*.pwp)|*.pwp|Visia Project Files (*.via)|*.via|Linley Project Files (*.lnl)|*.lnl|All Files (*.*)|*.*"
        .ShowOpen
    End With
    
lNoCD:
    If comdlg.FileName = "" Then Exit Sub
    If InStr(1, comdlg.FileName, Chr$(34)) <> 0 Then
        comdlg.FileName = Mid$(comdlg.FileName, InStr(1, comdlg.FileName, Chr$(34)) + 1, InStr(InStr(1, comdlg.FileName, Chr$(34)) + 1, comdlg.FileName, Chr$(34)) - InStr(1, comdlg.FileName, Chr$(34)) - 1)
    End If
    frmMain.Caption = Right(comdlg.FileName, Len(comdlg.FileName) - InStrRev(comdlg.FileName, "\", -1, vbTextCompare)) & " - Pascaline"
    'Open File
    
    Dim i As Long: Dim Ident As String: Dim FileNum As Long: Dim NumberOfItems As Long
    
    On Error GoTo InvalidFileType

    FileNum = FreeFile
    
    MakeInitControls True
    InitVirtualFiles
    
    Open comdlg.FileName For Binary As #FileNum
        Get #FileNum, , NumberOfItems
        ReDim VirtualFiles(UBound(VirtualFiles) + NumberOfItems) As TYPE_VIRTUAL_FILE
        For i = 1 To NumberOfItems
            Get #FileNum, , VirtualFiles(UBound(VirtualFiles) - NumberOfItems + i)
            If VirtualFiles(UBound(VirtualFiles) - NumberOfItems + i).Extension = EX_MODULE Then
                lstProject.AddItem "[MODULE] - " & VirtualFiles(UBound(VirtualFiles) - NumberOfItems + i).Name
            ElseIf VirtualFiles(UBound(VirtualFiles) - NumberOfItems + i).Extension = EX_ENTRY Then
                lstProject.AddItem "[ENTRY] - " & VirtualFiles(UBound(VirtualFiles) - NumberOfItems + i).Name
            ElseIf VirtualFiles(UBound(VirtualFiles) - NumberOfItems + i).Extension = EX_DIALOG Then
                lstProject.AddItem "[RESOURCE] - " & VirtualFiles(UBound(VirtualFiles) - NumberOfItems + i).Name
            End If
        Next i
    Close #1
    
    isDirty = False
    Unload frmNew
    Exit Sub
InvalidFileType:
    MsgBox "Error while loading file '" & comdlg.FileName & "'", vbCritical, "Pascaline"
    
CancelOpen:
    MakeInitControls
End Sub
Public Sub CheckForAssociation()
    Dim IsAssigned As String
    If InStr(1, CheckFileAssociation("pwp"), "3.0.0", vbTextCompare) = 0 Then
        DeleteFileAssociation "pwp"
    End If
    If CheckFileAssociation("pwp") <> "" Then Exit Sub
    MakeFileAssociation "pwp", App.Path & "\", App.EXEName, "", "project.ico"
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub

Function CheckUnsaved() As Boolean
    If isDirty = True Then
        Select Case MsgBox("Project has been changed do you want to save your current project?", vbInformation + vbYesNoCancel)
            Case vbYes: SaveProject False
            Case vbCancel: CheckUnsaved = True
        End Select
    End If
End Function

Function GetLineNumberByCarret(CarretPosition As Long)
    Dim ActualLine As Integer
    Dim CodeSource As String
    Dim i As Long
    CodeSource = txtEditor.Text
    
    ActualLine = 1
    For i = 1 To CarretPosition
        If Mid$(CodeSource, i, 2) = vbCrLf Then
            ActualLine = ActualLine + 1
        End If
    Next i
    GetLineNumberByCarret = ActualLine
End Function

Public Sub MakeExecuteable()
Dim comdlgexe As UnicodeFileDialog
    On Error GoTo ExecCancelError
    
    Set comdlgexe = New UnicodeFileDialog
    With comdlgexe
    .Filter = "Executable Files (*.exe)|*.exe|Dynamic Link Library (*.dll)|*.dll|"
    .CancelError = True
    .ShowSave Me.hWnd
    End With
    
    If Dir(comdlgexe.FileName) <> "" Then
        If MsgBox("File already exists! Overwrite anyway?", vbYesNo + vbCritical) = vbNo Then
            Exit Sub
        End If
    End If
    
    Compile comdlg.FileName, False
    isDirty = False
ExecCancelError:

End Sub

Sub NewProject(CreateNew As Boolean)
    If CheckUnsaved = True Then Exit Sub
    txtEditor.Text = ""
    comdlg.FileName = ""
    Me.Caption = "Unsaved - Pascaline"
    'NewTemplate
    frmNew.Show 1
    isDirty = False
End Sub

Sub MakeInitControls(Optional OpenProject As Boolean)
On Error Resume Next

    InitVirtualFiles
    InitScanObjects
    
    txtEditor.Visible = True

    lstProject.Clear
    
    If Not OpenProject Then lstProject.AddItem "Entry Point"
    If Not OpenProject Then CreateVirtualFile "Entry Point", EX_ENTRY, ""
       
    On Error Resume Next
    If Not OpenProject Then txtEditor.SetFocus
    
End Sub
Private Sub tmrApplicationRuntime_Timer()
    Dim i As Long
    
    GetExitCodeProcess hWndProg, ExitCode

    If ExitCode = STILL_ACTIVE& Then
        If RunEnabled = True Then
            frmMain.Caption = Right(comdlg.FileName, Len(comdlg.FileName) - InStrRev(comdlg.FileName, "\", -1, vbTextCompare)) & " - Visia Compiler 4.8.7 [Running]"
            txtEditor.Enabled = False: txtEditor.BackColor = &H8000000F: DoEvents
            mnuFile.Enabled = False: mnuEdit.Enabled = False: mnuCompile.Enabled = False: mnuHelp.Enabled = False: mnuEdit.Enabled = False
            RunEnabled = False
        End If
    Else
        If RunEnabled = False Then
            frmMain.Caption = Right(comdlg.FileName, Len(comdlg.FileName) - InStrRev(comdlg.FileName, "\", -1, vbTextCompare)) & " - Visia Compiler 4.8.7"
            txtEditor.Enabled = True: txtEditor.BackColor = &H80000005
            If hWndProg <> 0 Then CloseHandle hWndProg: hWndProg = 0
            mnuFile.Enabled = True: mnuEdit.Enabled = True: mnuCompile.Enabled = True: mnuHelp.Enabled = True: mnuEdit.Enabled = True
            RunEnabled = True
            frmMain.txtLog.Text = frmMain.txtLog.Text & "Ready .." & vbCrLf
            On Error Resume Next
            frmMain.SetFocus: txtEditor.SetFocus
            tmrApplicationRuntime.Enabled = False
        End If
    End If
End Sub

