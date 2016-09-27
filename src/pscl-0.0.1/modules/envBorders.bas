Attribute VB_Name = "envBorders"
Option Explicit

Public Declare Function GetWindowLong _
                         Lib "user32" _
                         Alias "GetWindowLongA" _
                         (ByVal hWnd As Long, _
                          ByVal nIndex As Long) _
                         As Long

Public Declare Function SetWindowLong _
                         Lib "user32" _
                         Alias "SetWindowLongA" _
                         (ByVal hWnd As Long, _
                          ByVal nIndex As Long, _
                          ByVal dwNewLong As Long) _
                         As Long

Public Declare Function SetWindowPos _
                         Lib "user32" _
                         (ByVal hWnd As Long, _
                          ByVal hWndInsertAfter As Long, _
                          ByVal X As Long, _
                          ByVal Y As Long, _
                          ByVal cx As Long, _
                          ByVal cy As Long, _
                          ByVal wFlags As Long) _
                         As Long

Public Const GWL_EXSTYLE = (-20)
'Public Const WS_EX_CLIENTEDGE = &H200
'Public Const WS_EX_STATICEDGE = &H20000

'Public Const SWP_FRAMECHANGED = &H20
'Public Const SWP_NOMOVE = &H2
'Public Const SWP_NOOWNERZORDER = &H200
'Public Const SWP_NOSIZE = &H1
'Public Const SWP_NOZORDER = &H4

Public Enum ControlType
    ctAll = 0
    ctCheckBox = 1
    ctComboBox = 2
    ctCommandButton = 3
    ctDirListBox = 4
    ctDriveListBox = 5
    ctFileListBox = 6
    ctFrame = 7
    ctHScrollBar = 8
    ctImage = 9
    ctImageCombo = 10
    ctLine = 11
    ctListBox = 12
    ctListView = 13
    ctOptionButton = 14
    ctPictureBox = 15
    ctProgressBar = 16
    ctPropertyPage = 17
    ctShape = 18
    ctStatusBar = 19
    ctTabStrip = 20
    ctTextBox = 21
    ctToolbar = 22
    ctTreeView = 23
    ctVScrollBar = 24
End Enum

Public Function AddOfficeBorders(frmForm As Form, _
                                 ctControlType As ControlType, _
                                 Optional blnNoBorderStyle As Boolean, _
                                 Optional blnNoAppearance As Boolean, _
                                 Optional strMsgBoxTitle As String, _
                                 Optional blnErr_ShowFriendly As Boolean, _
                                 Optional blnErr_ShowCritical As Boolean) _
                                As Long
On Error GoTo err_AddOfficeBorders    'initiate error handler
    AddOfficeBorders = 0    'set default return
    
On Error Resume Next
    
    Dim ctlControl      As Control
    
    For Each ctlControl In frmForm.Controls
        Select Case ctControlType
            Case 0: SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 1: If TypeOf ctlControl Is CheckBox Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 2: If TypeOf ctlControl Is ComboBox Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 3: If TypeOf ctlControl Is CommandButton Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 4: If TypeOf ctlControl Is DirListBox Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 5: If TypeOf ctlControl Is DriveListBox Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 6: If TypeOf ctlControl Is FileListBox Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 7: If TypeOf ctlControl Is Frame Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 8: If TypeOf ctlControl Is HScrollBar Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 9: If TypeOf ctlControl Is Image Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 10: If TypeOf ctlControl Is ImageCombo Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 11: If TypeOf ctlControl Is Line Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 12: If TypeOf ctlControl Is ListBox Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 13: If TypeOf ctlControl Is ListView Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 14: If TypeOf ctlControl Is OptionButton Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 15: If TypeOf ctlControl Is PictureBox Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 16: If TypeOf ctlControl Is ProgressBar Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 17: If TypeOf ctlControl Is PropertyPage Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 18: If TypeOf ctlControl Is Shape Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 19: If TypeOf ctlControl Is StatusBar Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 20: If TypeOf ctlControl Is TabStrip Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 21: If TypeOf ctlControl Is TextBox Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 22: If TypeOf ctlControl Is Toolbar Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 23: If TypeOf ctlControl Is TreeView Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
            Case 24: If TypeOf ctlControl Is VScrollBar Then SetOfficeBorder ctlControl, blnNoBorderStyle, blnNoAppearance
        End Select
    Next ctlControl
    
    AddOfficeBorders = 1
    
    Exit Function
err_AddOfficeBorders:    'error handler
    AddOfficeBorders = -1    'set internal error return
    'send message to immediate window
    Debug.Print Now & " | Function: & AddOfficeBorders & | Error: #" & _
                Err.Number & vbTab & Err.Description
    'if we want to show critical messages to the user
    If blnErr_ShowCritical = True Then
        'notify the user
        MsgBox "Error: #" & Err.Number & vbTab & Err.Description & _
               vbCrLf & vbCrLf & Now, _
               vbOKOnly + vbCritical, _
               strMsgBoxTitle & " [Function: AddOfficeBorders" & "]"
    End If
    Err.Clear    'clear the error object
On Error Resume Next
    'Cleanup
    
End Function

Public Function SetOfficeBorder(ByVal ctlControl As Control, _
                                Optional blnNoBorderStyle As Boolean, _
                                Optional blnNoAppearance As Boolean, _
                                Optional strMsgBoxTitle As String, _
                                Optional blnErr_ShowFriendly As Boolean, _
                                Optional blnErr_ShowCritical As Boolean) _
                               As Long
On Error GoTo err_SetOfficeBorder    'initiate error handler
    SetOfficeBorder = 0    'set default return
    
    Dim lngRetVal       As Long
    
    'Retrieve the current border style
    lngRetVal = GetWindowLong(ctlControl.hWnd, GWL_EXSTYLE)
    
    'Calculate border style to use
    lngRetVal = lngRetVal Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
    
On Error Resume Next
    If blnNoBorderStyle Then ctlControl.BorderStyle = 0
    If blnNoAppearance Then ctlControl.Appearance = 0
On Error GoTo err_SetOfficeBorder
    
    'Apply the changes
    SetWindowLong ctlControl.hWnd, GWL_EXSTYLE, lngRetVal
    SetWindowPos ctlControl.hWnd, _
                 0, 0, 0, 0, 0, _
                 SWP_NOMOVE Or _
                 SWP_NOSIZE Or _
                 SWP_NOOWNERZORDER Or _
                 SWP_NOZORDER Or _
                 SWP_FRAMECHANGED
    
    SetOfficeBorder = 1
    
    Exit Function
err_SetOfficeBorder:    'error handler
    SetOfficeBorder = -1    'set internal error return
    'send message to immediate window
    Debug.Print Now & " | Function: & SetOfficeBorder & | Error: #" & _
                Err.Number & vbTab & Err.Description
    'if we want to show critical messages to the user
    If blnErr_ShowCritical = True Then
        'notify the user
        MsgBox "Error: #" & Err.Number & vbTab & Err.Description & _
               vbCrLf & vbCrLf & Now, _
               vbOKOnly + vbCritical, _
               strMsgBoxTitle & " [Function: SetOfficeBorder" & "]"
    End If
    Err.Clear    'clear the error object
On Error Resume Next
    'Cleanup
    
End Function


