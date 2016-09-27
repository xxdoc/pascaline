Attribute VB_Name = "comImports"
Option Explicit

Type TYPE_IMPORT
    Name As String
    Alias As String
    Library As String
    pCount As Long
    Used As Boolean
End Type

Public Imports() As TYPE_IMPORT

Sub InitImports()
    ReDim Imports(0) As TYPE_IMPORT
End Sub

Function IsImportUsed(AliasID As Long) As Boolean
    If Imports(AliasID).Used = True Then
        IsImportUsed = True
    Else
        IsImportUsed = False
    End If
End Function

Function ImportExists(Name As String) As Boolean
    Dim i As Long
    For i = 0 To UBound(Imports)
        If Imports(i).Name = Name Then
            ImportExists = True
            Exit Function
        End If
    Next i
End Function

Sub AddImport(Name As String, Library As String, pCount As Long, Optional Alias As String, Optional Used As Boolean = False)
    If ImportExists(Name) Then Exit Sub
    ReDim Preserve Imports(UBound(Imports) + 1) As TYPE_IMPORT
    Imports(UBound(Imports)).Name = Name
    Imports(UBound(Imports)).pCount = pCount
    Imports(UBound(Imports)).Library = Library
    Imports(UBound(Imports)).Used = Used
    
    If Alias <> "" Then
        Imports(UBound(Imports)).Alias = Alias
    Else
        Imports(UBound(Imports)).Alias = Name
    End If
    
End Sub

Function SetImportUsed(Name As String, Offset As Long)
    Dim i As Long
    For i = 1 To UBound(Imports)
        If Imports(i).Alias = Name Then
            Imports(i).Used = True
            AddRelocation Offset
            Exit Function
        End If
    Next i
End Function

Function ImportPCountByName(Name As String) As Long
    Dim i As Long
    For i = 1 To UBound(Imports)
        If Imports(i).Alias = Name Then
            ImportPCountByName = Imports(i).pCount
            Exit Function
        End If
    Next i
End Function

Function IsImport(Ident As String) As Boolean
    Dim i As Long
    For i = 1 To UBound(Imports)
        If Imports(i).Alias = Ident Then
            IsImport = True
            Exit Function
        End If
    Next i
End Function

Sub GenerateImportTable(Optional NoSymbols As Boolean)
    Dim i As Long: Dim ii As Long: Dim Duplicate As Boolean: Dim Libraries() As String
    
    CurrentSection = ".idata"
    ReDim Libraries(0) As String
    
    If UBound(Imports) = 0 Then Exit Sub
    
    For i = 1 To UBound(Imports)
        
        While Not IsImportUsed(i)
            i = i + 1
            If i > UBound(Imports) Then Exit For
        Wend
    
        For ii = 1 To UBound(Libraries)
            If UCase(Libraries(ii)) = UCase(Imports(i).Library) Then
                Duplicate = True
            End If
        Next ii
        If Duplicate = False Then
            ReDim Preserve Libraries(UBound(Libraries) + 1) As String
            Libraries(UBound(Libraries)) = UCase(Imports(i).Library)
        End If
        
        Duplicate = False
        
    Next i
    
    For i = 1 To UBound(Libraries)
        AddSectionDWord &H0
        AddSectionDWord &H0
        AddSectionDWord &H0
        AddFixup Libraries(i) & "_NAME", OffsetOf(".idata"), Import
        AddSectionDWord &H0
        AddFixup Libraries(i) & "_TABLE", OffsetOf(".idata"), Import
        AddSectionDWord &H0
    Next i
    
    If UBound(Libraries) > 0 Then
        AddSectionDWord &H0
        AddSectionDWord &H0
        AddSectionDWord &H0
        AddSectionDWord &H0
        AddSectionDWord &H0
    End If
    
    For i = 1 To UBound(Libraries)
        AddSymbol Libraries(i) & "_TABLE", OffsetOf(".idata"), Import
        For ii = 1 To UBound(Imports)
            If UCase(Imports(ii).Library) = UCase(Libraries(i)) Then
                If Imports(ii).Used = True Then
                    AddSymbol Imports(ii).Alias, OffsetOf(".idata"), Import, ST_IMPORT
                    AddFixup Imports(ii).Name & "_ENTRY", OffsetOf(".idata"), Import
                    AddSectionDWord &H0
                End If
            End If
        Next ii
        AddSectionDWord &H0
    Next i
    
    For i = 1 To UBound(Libraries)
        AddSymbol Libraries(i) & "_NAME", OffsetOf(".idata"), Import
        For ii = 1 To Len(Libraries(i))
            AddSectionByte CByte(Asc(Mid$(UCase(Libraries(i)), ii, 1)))
        Next ii
        AddSectionByte 0
    Next i
    
    For i = 1 To UBound(Imports)
        If Imports(i).Used = True Then
            AddSymbol Imports(i).Name & "_ENTRY", OffsetOf(".idata"), Import
            AddSectionWord 0
            For ii = 1 To Len(Imports(i).Name)
                AddSectionByte CByte(Asc(Mid$(Imports(i).Name, ii, 1)))
            Next ii
            AddSectionByte 0
        End If
    Next i
    
End Sub

Sub ImportLibrary()
    Dim i As Long
    Dim lIID As Long
    Dim Ident As String
    Dim FileNum As Long
    Dim NumberOfItems As Long
    
    'On Error GoTo InvalidLibrary
    
    Call SkipBlank: Ident = StringExpression
    
    If Dir(App.Path & "\include\" & Ident) = "" Then ErrMessage "cannot include '" & Ident & "'. check your include folder.": Exit Sub
    
    FileNum = FreeFile
    
    Open App.Path & "\include\" & Ident For Binary As #FileNum
        Get #FileNum, , NumberOfItems
        ReDim Preserve Imports(UBound(Imports) + NumberOfItems) As TYPE_IMPORT
        For i = 1 To NumberOfItems
            Get #FileNum, , Imports(UBound(Imports) - NumberOfItems + i)
            Imports(UBound(Imports) - NumberOfItems + i).Used = False
        Next i
        
        Get #FileNum, , NumberOfItems
        ReDim Preserve Types(UBound(Types) + NumberOfItems) As TYPE_TYPE
        For i = 1 To NumberOfItems
            Get #FileNum, , Types(UBound(Types) - NumberOfItems + i)
        Next i
        
        Get #FileNum, , NumberOfItems
        ReDim Preserve Constants(UBound(Constants) + NumberOfItems) As TYPE_CONSTANT
        For i = 1 To NumberOfItems
            Get #FileNum, , Constants(UBound(Constants) - NumberOfItems + i)
        Next i
    Close #FileNum
    
    If IsSymbol(",") Then Symbol ",": ImportLibrary: Exit Sub
    Terminator
    CodeBlock
    Exit Sub
InvalidLibrary:
    ErrMessage "Invalid Library Format => " & Ident
    Close #FileNum
End Sub

Sub ExportLibrary()
    Dim i As Long
    Dim FileNum As Long
    Dim sFileName As String
    
    If frmMain.comdlg.FileName = "" Then ErrMessage "File is not saved => cannot build library.": Exit Sub
    
    InfMessage "Building Library.."
    
    FileNum = FreeFile
    sFileName = frmMain.comdlg.FileName
    sFileName = Left(sFileName, Len(sFileName) - 4)
    sFileName = sFileName & ".lib"
    
    If Dir(sFileName) <> "" Then Kill sFileName
    
    Open sFileName For Binary As #FileNum
        
        Put #FileNum, , UBound(Imports)
            For i = 1 To UBound(Imports): Put #FileNum, , Imports(i): Next i
        Put #FileNum, , UBound(Types)
            For i = 1 To UBound(Types): Put #FileNum, , Types(i): Next i
        Put #FileNum, , UBound(Constants)
            For i = 1 To UBound(Constants): Put #FileNum, , Constants(i): Next i
            
    Close #FileNum
    
    InfMessage "library compiled. => " & sFileName
    frmMain.RunEnabled = False
    WriteSummary Summary
End Sub
