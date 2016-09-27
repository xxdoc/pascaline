Attribute VB_Name = "comCompiler"
Option Explicit

Public bLibrary As Boolean
Public LibraryName As String
Public NameDLL As String
Public IsCmdCompile As Boolean

Sub Compile(sFile As String, bRun As Boolean)
    'On Error GoTo CompileFailed
    InitSummary
    InfMessage "Initializing Modules .."
        
    InitSections
    InitSymbols
    InitResources
    InitImports
    InitExports
    InitFixups
    InitFrames
    InitData
    InitTypes
    InitParser
    
    InfMessage "Parsing .."
    Parse
    
    If IsDLL Then NameDLL = Right(sFile, Len(sFile) - InStrRev(sFile, "\", -1, vbTextCompare)): NameDLL = Left(NameDLL, Len(NameDLL) - 3): NameDLL = NameDLL & "DLL"
    If pError = True Then WriteSummary Summary: Exit Sub
    If bLibrary = True Then ExportLibrary: Exit Sub
    InitLinker
    If pError = True Then WriteSummary Summary: Exit Sub
    DoFixups
    Link sFile, bRun
    Exit Sub
CompileFailed:
    pError = True
    ErrMessage "Internal Error -> " & Err.Description
    WriteSummary Summary
End Sub

Sub AddCodeByte(Value As Byte)
'Warning can be wrong if no standard CreateSection were applied! .code = Section ID 2
    ReDim Preserve Section(2).Bytes(UBound(Section(2).Bytes) + 1) As Byte
    Section(2).Bytes(UBound(Section(2).Bytes)) = Value
End Sub

Sub AddCodeWord(Value As Integer)
    AddCodeByte LoByte(Value)
    AddCodeByte HiByte(Value)
End Sub

Sub AddCodeSingle(Value As Single)
    Dim B1 As Byte: Dim B2 As Byte
    Dim B3 As Byte: Dim B4 As Byte
    
    Open App.Path & "\single.dmp" For Binary As #1
        Put #1, 1, Value
        Get #1, 1, B4: Get #1, 2, B3
        Get #1, 3, B2: Get #1, 4, B1
    Close #1
    
    AddCodeByte B4: AddCodeByte B3
    AddCodeByte B2: AddCodeByte B1
    
    Kill App.Path & "\single.dmp"
End Sub

Sub AddCodeDWord(Value As Long)
    AddCodeWord LoWord(Value)
    AddCodeWord HiWord(Value)
End Sub

Sub AddDataByte(Value As Byte)
'Warning can be wrong if no standard CreateSection were applied! .data = Section ID 1
    ReDim Preserve Section(1).Bytes(UBound(Section(1).Bytes) + 1) As Byte
    Section(1).Bytes(UBound(Section(1).Bytes)) = Value
End Sub

Sub AddDataWord(Value As Integer)
    AddDataByte LoByte(Value)
    AddDataByte HiByte(Value)
End Sub

Sub AddDataSingle(Value As Single)
    Dim B1 As Byte: Dim B2 As Byte
    Dim B3 As Byte: Dim B4 As Byte
    
    Open App.Path & "\single.dmp" For Binary As #1
        Put #1, 1, Value
        Get #1, 1, B4: Get #1, 2, B3
        Get #1, 3, B2: Get #1, 4, B1
    Close #1
    
    AddDataByte B4: AddDataByte B3
    AddDataByte B2: AddDataByte B1
    
    Kill App.Path & "\single.dmp"
End Sub

Sub AddDataDWord(Value As Long)
    'AddSectionNameDWord ".data", Value
    AddDataWord LoWord(Value)
    AddDataWord HiWord(Value)
End Sub

'------------------------------------------------------------------------
Sub AddImportByte(Value As Byte)
    AddSectionNameByte ".idata", Value
End Sub

Sub AddImportWord(Value As Integer)
    AddSectionNameWord ".idata", Value
End Sub

Sub AddImportDWord(Value As Long)
    AddSectionNameDWord ".idata", Value
End Sub

Sub AddExportByte(Value As Byte)
    AddSectionNameByte ".edata", Value
End Sub

Sub AddExportWord(Value As Integer)
    AddSectionNameWord ".edata", Value
End Sub

Sub AddExportDWord(Value As Long)
    AddSectionNameDWord ".edata", Value
End Sub

Sub AddResourceByte(Value As Byte)
    AddSectionNameByte ".rsrc", Value
End Sub

Sub AddResourceWord(Value As Integer)
    AddSectionNameWord ".rsrc", Value
End Sub

Sub AddResourceDWord(Value As Long)
    AddSectionNameDWord ".rsrc", Value
End Sub

Sub AddRelocationByte(Value As Byte)
    AddSectionNameByte ".reloc", Value
End Sub

Sub AddRelocationWord(Value As Integer)
    AddSectionNameWord ".reloc", Value
End Sub

Sub AddRelocationDWord(Value As Long)
    AddSectionNameDWord ".reloc", Value
End Sub

