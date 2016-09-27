Attribute VB_Name = "comResources"
Option Explicit

Type TYPE_RES_RESOURCE
    Value As Long
    FileName As String
    SymbolName As String
End Type

Type TYPE_RES_ITEM
    Value As Long
    RE() As TYPE_RES_RESOURCE
End Type

Type TYPE_RES_DIRECTORY
    Value As Long
    IT() As TYPE_RES_ITEM
End Type

Public D() As TYPE_RES_DIRECTORY
Public lUniqueBMP As Long

Sub InitResources()
    lUniqueBMP = 0
    ReDim D(0) As TYPE_RES_DIRECTORY
End Sub

'--------------------------------------------------------------
'Parsing And Adding Resources
'--------------------------------------------------------------

Sub DeclareBitmap()
    Dim i As Long
    Dim Ident As String
    Dim sFile As String
    
    lUniqueBMP = lUniqueBMP + 1
    Ident = Identifier
    Symbol ","
    sFile = StringExpression
    Terminator
    
    i = FindDirIDByType(2): If i = -1 Then i = AddResourceDirectory(2)
    AddResourceItem i, lUniqueBMP
    AddResourceResource i, GetDirItemUBound(CLng(i)), 2, sFile, Ident
    
    CodeBlock
End Sub

Function GetDirItemUBound(DirID As Integer) As Long
    GetDirItemUBound = UBound(D(DirID).IT)
End Function

Function FindDirIDByType(DirType As Integer) As Long
    Dim i As Integer
    For i = 1 To UBound(D)
        If D(i).Value = DirType Then
            FindDirIDByType = i
            Exit Function
        End If
    Next i
    FindDirIDByType = -1
End Function

Function AddResourceDirectory(Value As Long) As Long
    ReDim Preserve D(UBound(D) + 1) As TYPE_RES_DIRECTORY
    D(UBound(D)).Value = Value
    ReDim D(UBound(D)).IT(0) As TYPE_RES_ITEM
    AddResourceDirectory = UBound(D)
End Function

Sub AddResourceItem(DirID As Long, Value As Long)
    ReDim Preserve D(DirID).IT(UBound(D(DirID).IT) + 1) As TYPE_RES_ITEM
    D(DirID).IT(UBound(D(DirID).IT)).Value = Value
    ReDim D(DirID).IT(UBound(D(DirID).IT)).RE(0) As TYPE_RES_RESOURCE
End Sub

Sub AddResourceResource(DirID As Long, ItemID As Long, Value As Long, Optional FileName As String, Optional SymbolName As String)
    ReDim Preserve D(DirID).IT(ItemID).RE(UBound(D(DirID).IT(ItemID).RE) + 1) As TYPE_RES_RESOURCE
    D(DirID).IT(ItemID).RE(UBound(D(DirID).IT(ItemID).RE)).Value = Value
    D(DirID).IT(ItemID).RE(UBound(D(DirID).IT(ItemID).RE)).FileName = FileName
    D(DirID).IT(ItemID).RE(UBound(D(DirID).IT(ItemID).RE)).SymbolName = SymbolName
End Sub

'--------------------------------------------------------------
'Generate
'--------------------------------------------------------------

Sub AddResSymbol(Name As String)
    AddSymbol Name, OffsetOf(".rsrc"), Resource, ST_LABEL
End Sub

Sub AddResResource(Name As String)
    AddSectionDWord 0
    AddFixup Name, OffsetOf(".rsrc"), Resource, OffsetOf(".rsrc") + 4
    AddSectionDWord 0
End Sub

Sub AddResSubDirectory(Name As String, DirType As Long)
    AddSectionDWord DirType
    AddFixup Name, OffsetOf(".rsrc"), Resource, &H80000000 + OffsetOf(".rsrc") + 4
    AddSectionDWord 0
End Sub

Sub GenerateResources()
    Dim i As Integer: Dim ii As Integer: Dim iii As Integer
    
    If UBound(D) = 0 Then Exit Sub
    CurrentSection = ".rsrc"
    AddResSymbol "resource_root"
    AddSectionDWord 0
    AddSectionDWord 0
    AddSectionDWord 0
    AddSectionWord 0
    AddSectionWord UBound(D)
    
    For i = 1 To UBound(D)
        AddResSubDirectory "Directory_" & i, D(i).Value
    Next i
    
    For i = 1 To UBound(D)
        AddResSymbol "Directory_" & i
        AddSectionDWord 0
        AddSectionDWord 0
        AddSectionDWord 0
        AddSectionWord 0
        AddSectionWord UBound(D(i).IT)
        For ii = 1 To UBound(D(i).IT)
            AddResSubDirectory "ID_" & ii, D(i).IT(ii).Value
        Next ii
    Next i
    
    For i = 1 To UBound(D)
        For ii = 1 To UBound(D(i).IT)
            AddResSymbol "ID_" & ii
            AddSectionDWord 0
            AddSectionDWord 0
            AddSectionDWord 0
            AddSectionWord 0
            AddSectionWord UBound(D(i).IT(ii).RE)
            For iii = 1 To UBound(D(i).IT(ii).RE)
                AddResResource "ID_" & ii & "_resource_" & iii
            Next iii
        Next ii
    Next i
    
    For i = 1 To UBound(D)
        For ii = 1 To UBound(D(i).IT)
            For iii = 1 To UBound(D(i).IT(ii).RE)
                If Mid$(D(i).IT(ii).RE(iii).FileName, 1, 1) = "\" Then
                    D(i).IT(ii).RE(iii).FileName = Left(frmMain.comdlg.FileName, InStrRev(frmMain.comdlg.FileName, "\", Len(frmMain.comdlg.FileName), vbTextCompare) - 1) & D(i).IT(ii).RE(iii).FileName
                End If
                If Dir(D(i).IT(ii).RE(iii).FileName) = "" Then ErrMessage "file '" & D(i).IT(ii).RE(iii).FileName & "' does not exist!": Exit Sub
                AddResSymbol "ID_" & ii & "_resource_" & iii
                ChooseResource i, ii, iii
            Next iii
        Next ii
    Next i
    
    For i = 1 To UBound(D)
        For ii = 1 To UBound(D(i).IT)
            For iii = 1 To UBound(D(i).IT(ii).RE)
                If SymbolExists(D(i).IT(ii).RE(iii).SymbolName) Then ErrMessage "resource '" & D(i).IT(ii).RE(iii).SymbolName & "' already exists!": Exit Sub
                WriteResource i, ii, iii
            Next iii
        Next ii
    Next i
    
End Sub

Sub ChooseResource(DirID As Integer, ItemID As Integer, ResID As Integer)
    Select Case D(DirID).IT(ItemID).RE(ResID).Value
        Case 2:
            AddFixup D(DirID).IT(ItemID).RE(ResID).SymbolName, OffsetOf(".rsrc"), Resource
            AddSectionDWord 0
            AddSectionDWord FileLen(D(DirID).IT(ItemID).RE(ResID).FileName) - &HE
            AddSectionDWord 0
            AddSectionDWord 0
        Case 5:
        
        Case 3:
            Dim lSize As Long
            MsgBox D(DirID).IT(ItemID).RE(ResID).FileName
            lSize = DWordFromFile(D(DirID).IT(ItemID).RE(ResID).FileName, 14)
            AddFixup D(DirID).IT(ItemID).RE(ResID).SymbolName, OffsetOf(".rsrc"), Resource
            AddResourceDWord 0
            AddResourceDWord lSize
            AddResourceDWord 0
            AddResourceDWord 0
        Case Else
    End Select
End Sub

Sub WriteResource(DirID As Integer, ItemID As Integer, ResID As Integer)
    Dim i As Long
    Dim filec As Byte
    Select Case D(DirID).IT(ItemID).RE(ResID).Value
        Case 2
            AddSymbol D(DirID).IT(ItemID).RE(ResID).SymbolName, OffsetOf(".rsrc"), Resource, ST_RESOURCE
            Open D(DirID).IT(ItemID).RE(ResID).FileName For Binary As #1
                Seek #1, &HE + 1
                While Not EOF(1)
                    Get #1, , filec
                    AddSectionByte filec
                    i = i + 1
                Wend
            Close #1
            'ResAlign
        Case 3
            Dim lPos As Long: Dim lSize As Long
            lPos = DWordFromFile(D(DirID).IT(ItemID).RE(ResID).FileName, 18)
            lSize = DWordFromFile(D(DirID).IT(ItemID).RE(ResID).FileName, 14)
            AddSymbol D(DirID).IT(ItemID).RE(ResID).SymbolName, OffsetOf(".rsrc"), Resource, ST_RESOURCE
            Open D(DirID).IT(ItemID).RE(ResID).FileName For Binary As #1
                Seek #1, &HE + 1
                For i = lPos To lPos + lSize - 1
                    Get #1, , filec
                    AddSectionByte filec
                Next i
            Close #1
        Case 14:
            AddFixup "$" & ItemID & "_header_" & ResID, OffsetOf(".rsrc"), Resource
            AddResourceDWord 6 + (1 * 14)
            AddResourceDWord 0
            AddResourceDWord 0
            AddSymbol "$" & ItemID & "_header_" & ResID, OffsetOf(".rsrc"), Resource, ST_RESOURCE
            AddResourceWord 0
            AddResourceWord 1
            AddResourceWord LoWord(1)
            Open D(DirID).IT(ItemID).RE(ResID).FileName For Binary As #1
                Seek #1, &HE + 1
                For i = lPos To 16 + 22 - 1
                    Get #1, , filec
                    AddSectionByte filec
                Next i
            Close #1
        Case Else
    End Select
End Sub

Function DWordFromFile(sFileName As String, lPosition As Long) As Long
    Dim FileNum As Long, i As Long, lTmp As Long: lTmp = 0
    FileNum = FreeFile
    Open sFileName For Binary As #FileNum
    Seek #FileNum, lPosition + 1
    Get #FileNum, , lTmp
    Close #FileNum
    DWordFromFile = lTmp
End Function
