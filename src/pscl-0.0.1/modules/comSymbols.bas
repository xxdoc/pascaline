Attribute VB_Name = "comSymbols"
Option Explicit

Enum ENUM_SYMBOL_TYPE
    ST_RVA = 1
    ST_LABEL = 2
    ST_DWORD = 3
    ST_WORD = 4
    ST_BYTE = 5
    ST_US_DWORD = 6
    ST_US_WORD = 7
    ST_US_BYTE = 8
    ST_STRING = 9
    ST_TYPE = 10
    ST_IMPORT = 11
    ST_EXPORT = 12
    ST_RESOURCE = 13
    ST_FRAME = 14
    ST_LOCAL_DWORD = 15
    ST_LOCAL_STRING = 16
    ST_LOCAL_SINGLE = 17
    ST_SINGLE = 18
End Enum

Type TYPE_SYMBOL
    Name As String
    Offset As Long
    Section As ENUM_SECTION_TYPE
    SymType As ENUM_SYMBOL_TYPE
    IsProto As Boolean
End Type

Type TYPE_CONSTANT
    Name As String
    Value As String
End Type

Public Constants() As TYPE_CONSTANT
Public Symbols() As TYPE_SYMBOL

Sub InitSymbols()
    ReDim Symbols(0) As TYPE_SYMBOL
    ReDim Constants(0) As TYPE_CONSTANT
End Sub

Sub AddSymbol(Name As String, Offset As Long, Section As ENUM_SECTION_TYPE, Optional SymType As ENUM_SYMBOL_TYPE = 1, Optional IsProto As Boolean)
    If SymbolExists(Name) Then ErrMessage "symbol '" & Name & "' already exists": Exit Sub
    ReDim Preserve Symbols(UBound(Symbols) + 1) As TYPE_SYMBOL
    Symbols(UBound(Symbols)).Name = Name
    Symbols(UBound(Symbols)).Offset = Offset
    Symbols(UBound(Symbols)).Section = Section
    Symbols(UBound(Symbols)).SymType = SymType
    Symbols(UBound(Symbols)).IsProto = IsProto
End Sub

Function SymbolExists(Name As String) As Boolean
    Dim i As Integer
    For i = 1 To UBound(Symbols)
        If Symbols(i).Name = Name Then
            If Not Symbols(i).IsProto Then
                SymbolExists = True
                Exit Function
            End If
        End If
    Next i
End Function

Function GetSymbolOffset(Name As String) As Long
    Dim i As Long
    For i = 1 To UBound(Symbols)
        If Symbols(i).Name = Name Then
            GetSymbolOffset = Symbols(i).Offset
            Exit Function
        End If
    Next i
    ErrMessage "symbol '" & Name & "' does not exist!": Exit Function
End Function

Function GetSymbolSpace(Ident As String) As Long
    Dim i As Integer
    For i = 1 To UBound(Symbols)
        If Symbols(i).Name = Ident Then
            'GetSymbolSpace = Symbols(i).Space
            Exit Function
        End If
    Next i
End Function

Function GetSymbolSize(Ident As String) As Long
    Dim i As Integer
    
    For i = 1 To UBound(Symbols)
        If Symbols(i).Name = Ident Then
            If Symbols(i).SymType = ST_DWORD Or Symbols(i).SymType = ST_US_DWORD Or Symbols(i).SymType = ST_SINGLE Then
                GetSymbolSize = 4
            ElseIf Symbols(i).SymType = ST_WORD Or Symbols(i).SymType = ST_US_WORD Then
                GetSymbolSize = 2
            ElseIf Symbols(i).SymType = ST_BYTE Or Symbols(i).SymType = ST_US_BYTE Then
                GetSymbolSize = 1
            ElseIf Symbols(i).SymType = ST_STRING Then
                'GetSymbolSize = Symbols(i).Space
            Else
                ErrMessage "unknown symbol type '" & Ident & "'"
            End If
            Exit Function
        End If
    Next i
End Function

Function GetSymbolType(Ident As String) As ENUM_SYMBOL_TYPE
    Dim i As Integer
    For i = 1 To UBound(Symbols)
        If Symbols(i).Name = Ident Then
            GetSymbolType = Symbols(i).SymType
            Exit Function
        End If
    Next i
End Function

Function GetSymbolID(Ident As String) As Long
    Dim i As Integer
    For i = 1 To UBound(Symbols)
        If Symbols(i).Name = Ident Then
            GetSymbolID = i
            Exit Function
        End If
    Next i
End Function

Function GetConstant(Name As String) As Long
    Dim i As Integer
    If Name = "" Then Exit Function
    For i = 1 To UBound(Constants)
        If Name = Constants(i).Name Then
            GetConstant = Constants(i).Value
            Exit Function
        End If
    Next i
    ErrMessage "unknown constant '" & Name & "'"
End Function

Sub AddConstant(Name As String, Value As String)
    ReDim Preserve Constants(UBound(Constants) + 1) As TYPE_CONSTANT
    Constants(UBound(Constants)).Name = Name
    Constants(UBound(Constants)).Value = Value
End Sub
