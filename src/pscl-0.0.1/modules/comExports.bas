Attribute VB_Name = "comExports"
Option Explicit

Type TYPE_EXPORT
    Name As String
    Ordinal As Integer
End Type

Type TYPE_RELOCATION
    Offset As Long
End Type

Public IsDLL As Boolean
Public Exports() As TYPE_EXPORT
Public Relocations() As TYPE_RELOCATION

Sub InitExports()
    IsDLL = False
    ReDim Exports(0) As TYPE_EXPORT
    ReDim Relocations(0) As TYPE_RELOCATION
End Sub

Sub AddExport(Name As String)
    ReDim Preserve Exports(UBound(Exports) + 1) As TYPE_EXPORT
    Exports(UBound(Exports)).Name = Name
End Sub

Sub AddRelocation(Offset As Long)
    ReDim Preserve Relocations(UBound(Relocations) + 1) As TYPE_RELOCATION
    Relocations(UBound(Relocations)).Offset = Offset
End Sub

Sub WriteRelocations()
    Dim i As Integer
    CurrentSection = ".reloc"
   
    AddSectionDWord VirtualAddressOf(Code)
    AddFixup "Reloc_Last", OffsetOf(".reloc"), Relocate, (VirtualAddressOf(Relocate)) * (-1#)
    AddSectionDWord &H0
    
    For i = 1 To UBound(Relocations)
        If Relocations(i).Offset <> 0 Then
            AddSectionWord CInt(Relocations(i).Offset + &H3000)
        End If
    Next i
    
    AddSymbol "Reloc_Last", OffsetOf(".reloc"), Relocate
End Sub

Sub SortExports()
    Dim i As Integer
    Dim Elements() As String
    
    ReDim Elements(UBound(Exports)) As String
    
    For i = 1 To UBound(Exports)
        Elements(i) = Exports(i).Name
    Next i
    
    SortStringArray Elements, 1, UBound(Elements)
    
    For i = 1 To UBound(Elements)
        Exports(i).Name = Elements(i)
        Exports(i).Ordinal = i
    Next i
End Sub

Sub GenerateExportTable()
    Dim i As Integer
    Dim ii As Integer
    
    If UBound(Exports) = 0 Then Exit Sub
    
    CurrentSection = ".edata"
    
    SortExports
    
    AddSectionDWord 0
    AddSectionDWord 0
    AddSectionDWord 0
    AddFixup "DLL_NAME", OffsetOf(".edata"), Export
    AddSectionDWord 0
    AddSectionDWord 1
    
    AddSectionDWord UBound(Exports)
    AddSectionDWord UBound(Exports)
    AddFixup "ADDRESSES_TABLE", OffsetOf(".edata"), Export
    AddSectionDWord 0
    AddFixup "NAMES_TABLE", OffsetOf(".edata"), Export
    AddSectionDWord 0
    AddFixup "ORDINAL_TABLE", OffsetOf(".edata"), Export
    AddSectionDWord 0
    
    AddSymbol "ADDRESSES_TABLE", OffsetOf(".edata"), Export, ST_EXPORT
    For i = 1 To UBound(Exports)
        AddFixup Exports(i).Name & ".Address", OffsetOf(".edata"), Export
        AddSectionDWord 0
    Next i
    
    AddSymbol "NAMES_TABLE", OffsetOf(".edata"), Export, ST_EXPORT
    For i = 1 To UBound(Exports)
        AddFixup "_" & Exports(i).Name, OffsetOf(".edata"), Export
        AddSectionDWord 0
    Next i

    AddSymbol "ORDINAL_TABLE", OffsetOf(".edata"), Export, ST_EXPORT
    For i = 1 To UBound(Exports)
        AddSectionWord Exports(i).Ordinal - 1
    Next i
    
    AddSymbol "DLL_NAME", OffsetOf(".edata"), Export, ST_EXPORT
    For ii = 1 To Len(NameDLL)
        AddSectionByte CByte(Asc(Mid$(UCase(NameDLL), ii, 1)))
    Next ii
    AddSectionByte 0
    
    For i = 1 To UBound(Exports)
        AddSymbol "_" & Exports(i).Name, OffsetOf(".edata"), Export, ST_EXPORT
        For ii = 1 To Len(Exports(i).Name)
            AddSectionByte CByte(Asc(Mid$(Exports(i).Name, ii, 1)))
        Next ii
        AddSectionByte 0
    Next i
End Sub


