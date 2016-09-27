Attribute VB_Name = "comTypes"
Option Explicit

Type TYPE_TYPE
    Name As String
    Source As String
End Type

Public Types() As TYPE_TYPE
Public CurrentType As String
Public TypesLeft As Long

Sub InitTypes()
    CurrentType = ""
    TypesLeft = 0
    ReDim Types(0) As TYPE_TYPE
End Sub

Sub AssignType(Ident As String, AsIdent As String)
    Dim i As Integer
    Dim ii As Integer
    Dim myType As String
    Dim myIdent As String
    Dim myLastPos As Long
    Terminator
    
    For i = 1 To UBound(Types)
        If Types(i).Name = AsIdent Then
            AddSymbol Ident, OffsetOf(".data"), Data, ST_TYPE
            InsertSource Types(i).Source & "}"
            LenIncludes = LenIncludes + Len(Types(i).Source)
            myType = Ident
            CurrentType = Ident
            TypesLeft = 0
            While Not IsSymbol("}")
                myIdent = Identifier
                If IsType(myIdent) Then
                    myType = myIdent
                    myIdent = Identifier
                    CurrentType = CurrentType & "." & myIdent
                    Terminator
                    For ii = 1 To UBound(Types)
                        If Types(ii).Name = myType Then
                            AddSymbol CurrentType, OffsetOf(".data"), Data, ST_TYPE
                            InsertSource Types(ii).Source & "}"
                            LenIncludes = LenIncludes + Len(Types(ii).Source)
                            TypesLeft = TypesLeft + 1
                        End If
                    Next ii
                Else
                    VariableBlock myIdent, False, True
                End If
                If Position = myLastPos Then ErrMessage "expected '}' but could not process '" & AsIdent & "'.": Exit Sub
                If Position >= Len(Source) Then ErrMessage "expected '}' but found end of code.": Exit Sub
                myLastPos = Position
                SkipBlank
                DoEvents
                If IsSymbol("}") And TypesLeft > 0 Then Skip: TypesLeft = TypesLeft - 1: CurrentType = Ident
            Wend
            If IsSymbol("}") Then Skip
            CurrentType = ""
            SkipBlank
            CodeBlock
        End If
    Next i
End Sub

Function IsAssignedType(Ident As String) As Boolean
    Dim i As Integer
    For i = 1 To UBound(Symbols)
        If Symbols(i).Name = Ident Then
            If Symbols(i).SymType = ST_TYPE Then
                IsAssignedType = True
                Exit Function
            End If
        End If
    Next i
End Function

Function IsType(Ident As String) As Boolean
    Dim i As Integer
    For i = 1 To UBound(Types)
        If Types(i).Name = Ident Then
            IsType = True
            Exit Function
        End If
    Next i
End Function

