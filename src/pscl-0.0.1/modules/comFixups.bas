Attribute VB_Name = "comFixups"
Option Explicit

Type TYPE_FIXUP
    Name As String
    Offset As Long
    Value As Long
    ExtraAdd As Long
    Section As ENUM_SECTION_TYPE
    Deleted As Boolean
End Type

Public Fixups() As TYPE_FIXUP

Sub InitFixups()
    ReDim Fixups(0) As TYPE_FIXUP
End Sub

Sub AddCodeFixup(Name As String)
    AddFixup Name, OffsetOf(".code"), Code, &H400000
    AddCodeDWord 0
End Sub

Sub DeleteFixup(Name As String)
    Dim i As Long
    For i = 0 To UBound(Fixups)
        If Fixups(i).Name = Name Then
            Fixups(i).Deleted = True
            Exit Sub
        End If
    Next i
End Sub

Sub AddFixup(Name As String, Offset As Long, Section As ENUM_SECTION_TYPE, Optional ExtraAdd As Long)
    ReDim Preserve Fixups(UBound(Fixups) + 1) As TYPE_FIXUP
    Fixups(UBound(Fixups)).Name = Name
    Fixups(UBound(Fixups)).Offset = Offset
    Fixups(UBound(Fixups)).Section = Section
    Fixups(UBound(Fixups)).ExtraAdd = ExtraAdd
End Sub

Function LinkerFix(Offset As Long, Value As Long)
    Section(0).Bytes(Offset + 1) = LoByte(LoWord(Value))
    Section(0).Bytes(Offset + 2) = HiByte(LoWord(Value))
    Section(0).Bytes(Offset + 3) = LoByte(HiWord(Value))
    Section(0).Bytes(Offset + 4) = HiByte(HiWord(Value))
End Function

Function PhysicalAddressOf(EST As ENUM_SECTION_TYPE) As Long
    Dim i As Byte
    PhysicalAddressOf = SizeOfHeader
    For i = 1 To UBound(Section)
        If Section(i).SectionType = EST Then Exit Function
        PhysicalAddressOf = PhysicalAddressOf + PhysicalSizeOf(Section(i).Bytes, 1)
    Next i
End Function

Function VirtualAddressOf(EST As ENUM_SECTION_TYPE) As Long
    Dim i As Byte
    VirtualAddressOf = &H1000
    For i = 1 To UBound(Section)
        If Section(i).SectionType = EST Then Exit Function
        VirtualAddressOf = VirtualAddressOf + VirtualSizeOf(Section(i).Bytes, 1)
    Next i
End Function

Sub DoFixups()
    Dim i As Integer: Dim ii As Integer: Dim Found As Boolean
    
    For i = 1 To UBound(Fixups)
        If Fixups(i).Deleted = True Then GoTo SkipFixup
        For ii = 1 To UBound(Symbols)
            If Symbols(ii).IsProto Then GoTo SkipSymbol
            If Fixups(i).Name = Symbols(ii).Name Then
                If Symbols(ii).SymType = ST_LABEL Or _
                   Symbols(ii).SymType = ST_LOCAL_DWORD Or _
                   Symbols(ii).SymType = ST_LOCAL_SINGLE Or _
                   Symbols(ii).SymType = ST_LOCAL_STRING Or _
                   Symbols(ii).SymType = ST_FRAME Then
                    LinkerFix PhysicalAddressOf(Fixups(i).Section) + Fixups(i).Offset, _
                              Symbols(ii).Offset - Fixups(i).Offset - 4 + Fixups(i).ExtraAdd
            
                Else
                    LinkerFix PhysicalAddressOf(Fixups(i).Section) + Fixups(i).Offset, _
                              VirtualAddressOf(Symbols(ii).Section) + Symbols(ii).Offset + Fixups(i).ExtraAdd
                End If
                Found = True
                GoTo FixupFound
            End If
SkipSymbol:
        Next ii
FixupFound:
        If Found = False Then
            If InStr(1, Fixups(i).Name, ".HeapHandle", vbTextCompare) Then
                ErrMessage "'" & Left(Fixups(i).Name, InStr(1, Fixups(i).Name, ".", vbTextCompare) - 1) & "' is not an array.": Exit Sub
            ElseIf InStr(1, Fixups(i).Name, ".PtrToArray", vbTextCompare) Then
                ErrMessage "'" & Left(Fixups(i).Name, InStr(1, Fixups(i).Name, ".", vbTextCompare) - 1) & "' is not an array.": Exit Sub
            ElseIf InStr(1, Fixups(i).Name, "AddressToString", vbTextCompare) Then
                ErrMessage "cannot compare string with value": Exit Sub
            Else
                ErrMessage "symbol '" & Fixups(i).Name & "' doesn't exist. ": Exit Sub
            End If
        End If
        Found = False
        If Not IsCmdCompile Then
        frmMain.lblStatus.Caption = "Fixups.. (" & CInt(i / UBound(Fixups) * 100) & "% done..)"
        frmMain.progBar.Value = CInt(i / UBound(Fixups) * 100)
        End If
SkipFixup:
    Next i
End Sub
