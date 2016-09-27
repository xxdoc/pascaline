Attribute VB_Name = "comLinker"
Option Explicit

Enum ENUM_APP_TYPE
    GUI = 2
    CUI = 3
End Enum

Enum ENUM_SECTION_TYPE
    Data = 1
    Code = 2
    Import = 3
    Export = 4
    Resource = 5
    Linker = 6
    Relocate = 7
End Enum

Enum ENUM_SECTION_CHARACTERISTICS
    CH_CODE = &H20
    CH_INITIALIZED_DATA = &H40
    CH_UNINITIALIZED_DATA = &H80
    CH_MEM_DISCARDABLE = &H2000000
    CH_MEM_NOT_CHACHED = &H4000000
    CH_MEM_NOT_PAGED = &H8000000
    CH_MEM_SHARED = &H10000000
    CH_MEM_EXECUTE = &H20000000
    CH_MEM_READ = &H40000000
    CH_MEM_WRITE = &H80000000
End Enum

Type TYPE_SECTION
    Name As String
    Bytes() As Byte
    SectionType As ENUM_SECTION_TYPE
    Characteristics As ENUM_SECTION_CHARACTERISTICS
End Type

Public Section() As TYPE_SECTION
Public SizeOfHeader As Integer
Public AppType As ENUM_APP_TYPE
Public SizeOfAllSectionsBefore As Long
Public SizeOfAllSectionsBeforeRaw As Long

Sub InitLinker()
    GenerateResources
    GenerateImportTable
    GenerateExportTable
    If IsDLL Then WriteRelocations
    OutputDOSHeader
    OutputDOSStub
    OutputPEHeader
    OutputSectionTable
    OutputSections
End Sub

Sub InitSections()
    ReDim Section(0) As TYPE_SECTION
    ReDim Section(0).Bytes(0)
    Section(0).Name = ".link"
    Section(UBound(Section)).SectionType = 0
    Section(UBound(Section)).Characteristics = 0
End Sub


Sub Link(sFile As String, Run As Boolean)
    Dim i As Double
    
    On Error GoTo LinkFail
    
    If IsDLL Then sFile = Left(sFile, Len(sFile) - 3): sFile = sFile & "DLL"
    If Dir(sFile) <> "" Then Kill sFile
    
    Open sFile For Binary As #1
        For i = 1 To UBound(Section(0).Bytes)
            Put #1, , Section(0).Bytes(i)
        Next i
    Close #1
    
    If IsDLL Then
        InfMessage "dynamic link library compiled. " & vbCrLf & _
                   EndCounter & " seconds. " & UBound(Section(0).Bytes) & " bytes written."
        frmMain.RunEnabled = False
    Else
        InfMessage "application compiled. " & vbCrLf & _
                   EndCounter & " seconds. " & UBound(Section(0).Bytes) & " bytes written."
    End If
    
    sFileToRun = sFile
    WriteSummary Summary
    Exit Sub
LinkFail:
    ErrMessage "linking process failed -> proccess already running.": WriteSummary Summary
    Close #1
End Sub

Sub OutputDOSHeader()
    OutputDWord &H805A4D
    OutputDWord &H1
    OutputDWord &H100004
    OutputDWord &HFFFF
    OutputDWord &H140
    OutputDWord &H0
    OutputDWord &H40
    OutputDWord &H0
    OutputDWord &H0
    OutputDWord &H0
    OutputDWord &H0
    OutputDWord &H0
    OutputDWord &H0
    OutputDWord &H0
    OutputDWord &H0
    OutputDWord &H80
End Sub

Sub OutputDOSStub()
    OutputDWord &HEBA1F0E
    OutputDWord &HCD09B400
    OutputDWord &H4C01B821
    OutputDWord &H687421CD
    OutputDWord &H70207369
    OutputDWord &H72676F72
    OutputDWord &H63206D61
    OutputDWord &H6F6E6E61
    OutputDWord &H65622074
    OutputDWord &H6E757220
    OutputDWord &H206E6920
    OutputDWord &H20534F44
    OutputDWord &H65646F6D
    OutputDWord &H240A0D2E
    OutputDWord &H0
    OutputDWord &H0
End Sub

Sub OutputPEHeader()
    
    OutputDWord &H4550                     'Signature = "PE"
    OutputWord &H14C                       'Machine 0x014C;i386
    OutputWord NumberOfSections            'NumberOfSections = 4
    OutputDWord &H0                        'TimeDateStamp
    OutputDWord &H0                        'PointerToSymbolTable = 0
    OutputDWord &H0                        'NumberOfSymbols = 0
    OutputWord &HE0                        'SizeOfOptionalHeader
    If IsDLL Then
        OutputWord &H210E                      'Characteristics
    Else
        OutputWord &H818F                      'Characteristics
    End If
    
    OutputWord &H10B                       'Magic
    OutputByte &H5                         'MajorLinkerVersion
    OutputByte &H0                         'MinerLinkerVersion
    DeclareAttribute "SizeOfCode"          'SizeOfCode
    DeclareAttribute "SizeOfInitializedData"    'SizeOfInitializedData
    DeclareAttribute "SizeOfUnInitializedData"  'SizeOfUnInitializedData
    DeclareAttribute "AddressOfEntryPoint" 'AddressOfEntryPoint
    DeclareAttribute "BaseOfCode"          'BaseOfCode
    DeclareAttribute "BaseOfData"          'BaseOfData
    OutputDWord &H400000                   'ImageBase
    OutputDWord &H1000                     'SectionAlignment
    OutputDWord &H200                      'FileAlignment
    OutputWord &H1                         'MajorOSVersion
    OutputWord &H0                         'MinorOSVersion
    OutputWord &H0                         'MajorImageVersion
    OutputWord &H0                         'MinorImageVersion
    OutputWord &H4                         'MajorSubSystemVerion
    OutputWord &H0                         'MinorSubSystemVerion
    OutputDWord &H0                        'Win32VersionValue
    DeclareAttribute "SizeOfImage"         'SizeOfImage
    DeclareAttribute "SizeOfHeaders"       'SizeOfHeaders
    OutputDWord &H0                        'CheckSum
    OutputWord CInt(AppType)               'SubSystem = 2:GUI; 3:CUI
    OutputWord &H0                         'DllCharacteristics
    OutputDWord &H10000                    'SizeOfStackReserve
    OutputDWord &H10000                    'SizeOfStackCommit
    OutputDWord &H10000                    'SizeOfHeapReserve
    OutputDWord &H0                        'SizeOfHeapRCommit
    OutputDWord &H0                        'LoaderFlags
    OutputDWord &H10                       'NumberOfDataDirectories
    
    DeclareAttribute "ExportTable.Entry"
    DeclareAttribute "ExportTable.Size"
    
    DeclareAttribute "ImportTable.Entry"
    DeclareAttribute "ImportTable.Size"
    
    DeclareAttribute "ResourceTable.Entry"
    DeclareAttribute "ResourceTable.Size"
   
    OutputDWord &H0: OutputDWord &H0      'Exception_Table
    OutputDWord &H0: OutputDWord &H0      'Certificate_Table

    DeclareAttribute "RelocationTable.Entry"
    DeclareAttribute "RelocationTable.Size"
    
    OutputDWord &H0: OutputDWord &H0      'Debug_Data
    OutputDWord &H0: OutputDWord &H0      'Architecture
    OutputDWord &H0: OutputDWord &H0      'Global_PTR
    OutputDWord &H0: OutputDWord &H0      'TLS_Table
    OutputDWord &H0: OutputDWord &H0      'Load_Config_Table
    OutputDWord &H0: OutputDWord &H0      'BoundImportTable
    OutputDWord &H0: OutputDWord &H0      'ImportAddressTable
    OutputDWord &H0: OutputDWord &H0      'DelayImportDescriptor
    OutputDWord &H0: OutputDWord &H0      'COMplusRuntimeHeader
    OutputDWord &H0: OutputDWord &H0      'Reserved
End Sub

Sub OutputSectionTable()
    Dim i As Integer
    Dim ni As Integer
    For i = 1 To UBound(Section)
        
        If UBound(Section(i).Bytes) = 0 Then GoTo SkipSectionST
        
        'Output 8 bytes for name
        For ni = 1 To 8
            If ni > Len(Section(i).Name) Then
                OutputByte &H0
            Else
                OutputByte Asc(Mid$(Section(i).Name, ni, 1))
            End If
        Next ni
        
        DeclareAttribute Section(i).Name & ".VirtualSize"
        DeclareAttribute Section(i).Name & ".VirtualAddress"
        DeclareAttribute Section(i).Name & ".SizeOfRawData"
        DeclareAttribute Section(i).Name & ".PointerToRawData"
        DeclareAttribute Section(i).Name & ".PointerToRelocations"
        OutputDWord &H0                        'PointerToLinenumbers
        OutputWord &H0                         'NumberOfRelocations
        OutputWord &H0                         'NumberOfLinenumbers
        OutputDWord Section(i).Characteristics 'Characteristics
SkipSectionST:
    Next i

    Dim ii As Integer: Dim HowBig As Integer
    
    SizeOfHeader = UBound(Section(0).Bytes)
    For ii = 0 To SizeOfHeader + 512 Step 512
        HowBig = ii
    Next ii
    
    For ii = SizeOfHeader To HowBig - 1
        AddSectionNameByte ".link", &H0
    Next ii
    
    SizeOfHeader = UBound(Section(0).Bytes)
    FixAttribute "SizeOfHeaders", CLng(SizeOfHeader)
    SizeOfAllSectionsBeforeRaw = SizeOfHeader
End Sub

Sub FixTableEntry(SectionID As Integer)
    Select Case Section(SectionID).SectionType
        Case Code: FixAttribute "AddressOfEntryPoint", SizeOfAllSectionsBefore
        Case Import: FixAttribute "ImportTable.Entry", SizeOfAllSectionsBefore
        Case Export: FixAttribute "ExportTable.Entry", SizeOfAllSectionsBefore
        Case Resource: FixAttribute "ResourceTable.Entry", SizeOfAllSectionsBefore
        Case Relocate: FixAttribute "RelocationTable.Entry", SizeOfAllSectionsBefore
    End Select
End Sub

Sub FixTableSize(SectionID As Integer, Size As Long)
    Select Case Section(SectionID).SectionType
        Case Import: FixAttribute "ImportTable.Size", Size
        Case Export: FixAttribute "ExportTable.Size", Size
        Case Resource: FixAttribute "ResourceTable.Size", Size
        Case Relocate: FixAttribute "RelocationTable.Size", Size
    End Select
End Sub

Sub OutputSections()
    Dim i As Integer
    Dim ii As Long
    Dim PhysicalSize As Long
    
    SizeOfAllSectionsBefore = &H1000
    
    For i = 1 To UBound(Section)
        
        If UBound(Section(i).Bytes) = 0 Then GoTo SkipSectionOS
        
        FixAttribute Section(i).Name & ".VirtualSize", UBound(Section(i).Bytes)
        
        FixTableSize i, UBound(Section(i).Bytes)
        
        PhysicalSize = PhysicalSizeOf(Section(i).Bytes)
        
        For ii = UBound(Section(i).Bytes) To PhysicalSize - 1
            AddSectionNameByte Section(i).Name, &H0
        Next ii
        
        FixTableEntry i
        
        If Section(i).Name = ".reloc" Then FixAttribute ".code.PointerToRelocations", SizeOfAllSectionsBefore
        FixAttribute Section(i).Name & ".VirtualAddress", SizeOfAllSectionsBefore
        SizeOfAllSectionsBefore = SizeOfAllSectionsBefore + VirtualSizeOf(Section(i).Bytes)
        
        FixAttribute Section(i).Name & ".PointerToRawData", SizeOfAllSectionsBeforeRaw
        SizeOfAllSectionsBeforeRaw = SizeOfAllSectionsBeforeRaw + PhysicalSize
        
        FixAttribute Section(i).Name & ".SizeOfRawData", PhysicalSize
        
        For ii = 1 To UBound(Section(i).Bytes)
            AddSectionNameByte ".link", Section(i).Bytes(ii)
        Next ii
        
        For ii = 0 To &H1000& * &HFFFF& Step &H1000
            If PhysicalSize = ii Then SizeOfAllSectionsBefore = SizeOfAllSectionsBefore - &H1000
        Next ii
        
SkipSectionOS:

    Next i
    
    FixAttribute "SizeOfImage", SizeOfAllSectionsBefore
    
End Sub

Function PhysicalSizeOf(Value() As Byte, Optional ExtraSub As Long) As Long
    Dim i As Long
    If UBound(Value) = 0 Then PhysicalSizeOf = 0: Exit Function
    For i = 0 To UBound(Value) + 512 - ExtraSub Step 512
        PhysicalSizeOf = i
    Next i
End Function

Function VirtualSizeOf(Value() As Byte, Optional ExtraSub As Long) As Long
    Dim i As Long
    If UBound(Value) = 0 Then
        VirtualSizeOf = 0
    Else
        For i = &H1000 To &H1000& * &HFFFF& Step &H1000
            If i > (UBound(Value) - ExtraSub) Then
                VirtualSizeOf = i
                Exit For
            End If
        Next i
    End If
End Function

Sub DeclareAttribute(Name As String)
    AddSymbol Name, UBound(Section(0).Bytes), Linker
    OutputDWord &H0
End Sub

Sub FixAttribute(Name As String, Value As Long)
    FixDWord GetSymbolOffset(Name), CLng(Value)
End Sub

Function OffsetOf(Name As String) As Long
    Dim i As Byte
    For i = 1 To UBound(Section)
        If Section(i).Name = Name Then
            OffsetOf = UBound(Section(i).Bytes): Exit Function
        End If
    Next i
End Function

Function FixDWord(Offset As Long, Value As Long)
    Section(0).Bytes(Offset + 1) = LoByte(LoWord(Value))
    Section(0).Bytes(Offset + 2) = HiByte(LoWord(Value))
    Section(0).Bytes(Offset + 3) = LoByte(HiWord(Value))
    Section(0).Bytes(Offset + 4) = HiByte(HiWord(Value))
End Function

Function NumberOfSections() As Integer
    Dim i As Byte
    For i = 1 To UBound(Section)
        If UBound(Section(i).Bytes) > 0 Then
            NumberOfSections = NumberOfSections + 1
        End If
    Next i
End Function

Function SectionExists(Name As String) As Boolean
    Dim i As Byte
    For i = 1 To UBound(Section)
        If Section(i).Name = Name Then
            SectionExists = True: Exit Function
        End If
    Next i
End Function

Function SectionID(Name As String) As Byte
    Dim i As Byte
    For i = 1 To UBound(Section)
        If Section(i).Name = Name Then
            SectionID = i: Exit Function
        End If
    Next i
End Function

Sub OutputByte(Value As Byte)
    AddSectionNameByte ".link", Value
End Sub

Sub OutputWord(Value As Integer)
    AddSectionNameWord ".link", Value
End Sub

Sub OutputDWord(Value As Long)
    AddSectionNameDWord ".link", Value
End Sub

Sub AddSectionByte(Value As Byte)
    Dim lID As Integer
    lID = GetSectionIDByName(CurrentSection)
    ReDim Preserve Section(lID).Bytes(UBound(Section(lID).Bytes) + 1) As Byte
    Section(lID).Bytes(UBound(Section(lID).Bytes)) = Value
End Sub

Sub AddSectionWord(Value As Integer)
    AddSectionByte LoByte(Value)
    AddSectionByte HiByte(Value)
End Sub

Sub AddSectionDWord(Value As Long)
    AddSectionWord LoWord(Value)
    AddSectionWord HiWord(Value)
End Sub

Sub AddSectionNameByte(Name As String, Value As Byte)
    Dim lID As Integer
    lID = GetSectionIDByName(Name)
    ReDim Preserve Section(lID).Bytes(UBound(Section(lID).Bytes) + 1) As Byte
    Section(lID).Bytes(UBound(Section(lID).Bytes)) = Value
End Sub

Sub AddSectionNameWord(Name As String, Value As Integer)
    AddSectionNameByte Name, LoByte(Value)
    AddSectionNameByte Name, HiByte(Value)
End Sub

Sub AddSectionNameDWord(Name As String, Value As Long)
    AddSectionNameWord Name, LoWord(Value)
    AddSectionNameWord Name, HiWord(Value)
End Sub

Sub AddSectionNameSingle(Name As String, Value As Single)
    Dim B1 As Byte: Dim B2 As Byte
    Dim B3 As Byte: Dim B4 As Byte
    
    Open App.Path & "\single.dmp" For Binary As #1
        Put #1, 1, Value
        Get #1, 1, B4: Get #1, 2, B3
        Get #1, 3, B2: Get #1, 4, B1
    Close #1
    
    AddSectionNameByte Name, B4: AddSectionNameByte Name, B3
    AddSectionNameByte Name, B2: AddSectionNameByte Name, B1
    
    Kill App.Path & "\single.dmp"
End Sub

Function GetSectionIDByName(Name As String) As Long
    Dim i As Byte
    For i = 0 To UBound(Section)
        If Section(i).Name = Name Then
            GetSectionIDByName = i
            Exit Function
        End If
    Next i
    ErrMessage "section '" & Name & "' does not exist!"
End Function

