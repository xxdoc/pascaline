Attribute VB_Name = "envObjScan"
Option Explicit

Type TYPE_SCAN_OBJECT
    Category As String
    Parent As String
    Position As Long
End Type

Public ScanObjects() As TYPE_SCAN_OBJECT

Sub InitScanObjects()
    ReDim ScanObjects(0) As TYPE_SCAN_OBJECT
End Sub

Sub AddScanObject(Category As String, Parent As String, Position As Long)
    ReDim Preserve ScanObjects(UBound(ScanObjects) + 1) As TYPE_SCAN_OBJECT
    ScanObjects(UBound(ScanObjects)).Category = Category
    ScanObjects(UBound(ScanObjects)).Parent = Parent
    ScanObjects(UBound(ScanObjects)).Position = Position
End Sub

Function IsCategoryUsed(Category As String) As Boolean
    Dim i As Long
    For i = 1 To UBound(ScanObjects)
        If ScanObjects(i).Category = Category Then
            IsCategoryUsed = True
            Exit Function
        End If
    Next i
End Function

Sub SelectObjectByScan()
    Dim i As Long
    Dim ii As Long
    Dim cFrame As String
    Dim sSource As String
    sSource = frmMain.txtEditor.Text
    'i = frmMain.txtCode.GetCarretPos + 2
    While i > 0
        If Mid$(sSource, i, 5) = "frame" Then
            i = i + 5
            While Mid$(sSource, i, 1) <> "("
                cFrame = cFrame & Mid$(sSource, i, 1)
                DoEvents
                If i >= Len(sSource) Then GoTo ObjectNothingFound
                i = i + 1
            Wend
            'If frmMain.cmbParent.Text = Trim(cFrame) Then Exit Sub
            'frmMain.SelectObjectObject "Frames"
            'frmMain.SelectObjectParent Trim(cFrame)
            Exit Sub
        ElseIf Mid$(sSource, i, 6) = "export" Then
            i = i + 6
            While Mid$(sSource, i, 1) <> "("
                cFrame = cFrame & Mid$(sSource, i, 1)
                DoEvents
                If i >= Len(sSource) Then GoTo ObjectNothingFound
                i = i + 1
            Wend
            'If frmMain.cmbParent.Text = Trim(cFrame) Then Exit Sub
            'frmMain.SelectObjectObject "Exports"
            'frmMain.SelectObjectParent Trim(cFrame)
            Exit Sub
        ElseIf Mid$(sSource, i, 5) = "entry" Then
            i = i + 5
            'If frmMain.cmbParent.Text = "Entry" Then Exit Sub
            'frmMain.SelectObjectObject "General"
            'frmMain.SelectObjectParent "Entry"
            Exit Sub
        ElseIf Mid$(sSource, i, 6) = "import" Then
            i = i + 6
            
            While Mid$(sSource, i, 1) = " "
                i = i + 1
                If i >= Len(sSource) Then GoTo ObjectNothingFound
                DoEvents
            Wend
            
            While Mid$(sSource, i, 1) <> " "
                cFrame = cFrame & Mid$(sSource, i, 1)
                DoEvents
                If i >= Len(sSource) Then GoTo ObjectNothingFound
                i = i + 1
            Wend
            'If frmMain.cmbParent.Text = Trim(cFrame) Then Exit Sub
            'frmMain.SelectObjectObject "Imports"
            'frmMain.SelectObjectParent Trim(cFrame)
            Exit Sub
        ElseIf Mid$(sSource, i, 4) = "end;" Or _
               Mid$(sSource, i, 4) = "end." Then
            'frmMain.cmbObject.Text = ""
            'frmMain.cmbParent.Text = ""
            Exit Sub
        ElseIf Mid$(sSource, i, 3) = "lib" Then
            'frmMain.cmbObject.Text = ""
            'frmMain.cmbParent.Text = ""
            Exit Sub
        End If
        i = i - 1
        If i <= 1 Then GoTo ObjectNothingFound
        DoEvents
    Wend
On Error Resume Next
    frmMain.txtEditor.SetFocus
    Exit Sub
ObjectNothingFound:
On Error Resume Next
    'frmMain.cmbObject.Text = ""
    'frmMain.cmbParent.Text = ""
    'frmMain.Code.SetFocus
End Sub

Sub CodeScan()
    Dim i As Long
    Dim sCode As String
    Dim sIdent As String
    InitScanObjects
    sCode = frmMain.txtEditor.Text
    i = 1
    'frmMain.cmbParent.ComboItems.Clear
    'frmMain.cmbObject.ComboItems.Clear
    While i < Len(sCode)
        
        If Mid$(sCode, i, 5) = "frame" Then
            i = i + 5
            While Mid$(sCode, i, 1) <> "("
                sIdent = sIdent & Mid$(sCode, i, 1)
                i = i + 1
                If i >= Len(sCode) Then Exit Sub
                DoEvents
            Wend
            AddScanObject "Frames", Trim(sIdent), i
        ElseIf Mid$(sCode, i, 6) = "export" Then
            i = i + 6
            While Mid$(sCode, i, 1) <> "("
                sIdent = sIdent & Mid$(sCode, i, 1)
                i = i + 1
                If i >= Len(sCode) Then Exit Sub
                DoEvents
            Wend
            AddScanObject "Exports", Trim(sIdent), i
        ElseIf Mid$(sCode, i, 5) = "entry" Then
            i = i + 5
            AddScanObject "General", "Entry", i
        ElseIf Mid$(sCode, i, 6) = "import" Then
            i = i + 6
            
            While Mid$(sCode, i, 1) = " "
                i = i + 1
                If i >= Len(sCode) Then Exit Sub
                DoEvents
            Wend
            
            While Mid$(sCode, i, 1) <> " "
                sIdent = sIdent & Mid$(sCode, i, 1)
                i = i + 1
                If i >= Len(sCode) Then Exit Sub
                DoEvents
            Wend
            AddScanObject "Imports", Trim(sIdent), i
        End If
        
        sIdent = ""
        i = i + 1
        DoEvents
    Wend

    'If IsCategoryUsed("General") Then frmMain.cmbObject.ComboItems.Add , , "General", "Object"
    'If IsCategoryUsed("Frames") Then frmMain.cmbObject.ComboItems.Add , , "Frames", "Object"
    'If IsCategoryUsed("Exports") Then frmMain.cmbObject.ComboItems.Add , , "Exports", "Object"
    'If IsCategoryUsed("Imports") Then frmMain.cmbObject.ComboItems.Add , , "Imports", "Object"
    
    On Error Resume Next
    ''frmMain.cmbObject.Text = "": frmMain.cmbParent.Text = ""
    'frmMain.cmbObject.ComboItems(1).Selected = True
    'frmMain.cmbObject_Click
End Sub

Sub AddScanObjectsToCombo(Category As String)
    Dim i As Long
    'frmMain.cmbParent.ComboItems.Clear
    'frmMain.cmbParent.Text = ""
    If Category = "General" Then
        For i = 1 To UBound(ScanObjects)
            If ScanObjects(i).Category = "General" Then
                'frmMain.cmbParent.ComboItems.Add , , ScanObjects(i).Parent, "Parent"
            End If
        Next i
    ElseIf Category = "Frames" Then
        For i = 1 To UBound(ScanObjects)
            If ScanObjects(i).Category = "Frames" Then
                'frmMain.cmbParent.ComboItems.Add , , ScanObjects(i).Parent, "Parent"
            End If
        Next i
    ElseIf Category = "Exports" Then
        For i = 1 To UBound(ScanObjects)
            If ScanObjects(i).Category = "Exports" Then
                'frmMain.cmbParent.ComboItems.Add , , ScanObjects(i).Parent, "Parent"
            End If
        Next i
    ElseIf Category = "Imports" Then
        For i = 1 To UBound(ScanObjects)
            If ScanObjects(i).Category = "Imports" Then
                'frmMain.cmbParent.ComboItems.Add , , ScanObjects(i).Parent, "Parent"
            End If
        Next i
    End If
    'frmMain.cmbParent.ComboItems(1).Selected = True
End Sub
