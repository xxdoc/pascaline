Attribute VB_Name = "comFrames"
Option Explicit

Type TYPE_FRAME
    Name As String
    Declares As String
    ReturnAs As String
    Property As Boolean
End Type

Public IsCallFrame As Boolean
Public Frames() As TYPE_FRAME
Public ArgCount As Long
Public CurrentFrame As String
Public fcUniqueID As Long

Sub InitFrames()
    fcUniqueID = 0
    ReDim Frames(0) As TYPE_FRAME
End Sub

Sub AddFrameDeclare(VarName As String)
    Frames(UBound(Frames)).Declares = Frames(UBound(Frames)).Declares & VarName & ","
End Sub

Sub AddFrame(Name As String, Optional IsProperty As Boolean)
    ReDim Preserve Frames(UBound(Frames) + 1) As TYPE_FRAME
    Frames(UBound(Frames)).Name = Name
    Frames(UBound(Frames)).Property = IsProperty
    CurrentFrame = Name
End Sub

Function IsLocalVariable(Ident As String) As Boolean
    Dim i As Integer
    For i = 1 To UBound(Symbols)
        If Symbols(i).Name = CurrentFrame & "." & Ident Then
            IsLocalVariable = True
            Exit Function
        End If
    Next i
End Function

Function GetFrameIDByName(Name As String) As Long
    Dim i As Integer
    For i = 1 To UBound(Frames)
        If Frames(i).Name = Name Then
            GetFrameIDByName = i
            Exit Function
        End If
    Next i
End Function

Function GetReturnType(Name As String)
    Dim i As Integer
    For i = 1 To UBound(Frames)
        If Frames(i).Name = Name Then
            GetReturnType = Frames(i).ReturnAs
            Exit Function
        End If
    Next i
End Function

Sub CallFrame(Ident As String, Optional FromExpression As Boolean)
    Dim i As Integer
    Dim fID As Long
    Dim iLabel As Long
    Dim FrameDeclares As Variant
    
    IsCallFrame = True
    fID = GetFrameIDByName(Ident)
   
    FrameDeclares = Split(Frames(fID).Declares, ",")
    ReverseParams 'UBound(FrameDeclares)
    
    Symbol "("
    For i = UBound(FrameDeclares) - 1 To 0 Step -1
        If GetSymbolType(Frames(fID).Name & "." & FrameDeclares(i)) = ST_LOCAL_DWORD Or _
           GetSymbolType(Frames(fID).Name & "." & FrameDeclares(i)) = ST_LOCAL_SINGLE Then
            Expression
        ElseIf GetSymbolType(Frames(fID).Name & "." & FrameDeclares(i)) = ST_LOCAL_STRING Then
            Expression
        Else
            Expression Frames(fID).Name & "." & FrameDeclares(i)
            PushContent Frames(fID).Name & "." & FrameDeclares(i)
        End If
        If IsSymbol(",") Then Position = Position + 1
    Next i
    
    Symbol ")"
    IsCallFrame = False
    If Not FromExpression Then Terminator
    ExprCall Ident
    If Not FromExpression Then CodeBlock
End Sub

Sub CallProperty(Ident As String, Optional FromExpression As Boolean)
    Dim i As Integer
    Dim fID As Long
    Dim iLabel As Long
    Dim FrameDeclares As Variant
    
    IsCallFrame = True
    SkipBlank
    If Not FromExpression Then
        Symbol "="
        Expression "$Intern.Property"
        PushContent "$Intern.Property"
        Terminator
    End If
    IsCallFrame = False
    If CurrentFrame <> Ident Then
        'if Form1.Top = 20 is called outside a frame the routine will be raised. Else Form1.Top will be just assigned
        ExprCall Ident
    End If
    If Not FromExpression Then CodeBlock
End Sub

Function IsFrame(Name As String) As Boolean
    Dim i As Integer
    For i = 1 To UBound(Frames)
        If Frames(i).Name = Name Then
            IsFrame = True
            Exit Function
        End If
    Next i
End Function

Function IsProperty(Name As String) As Boolean
    Dim i As Integer
    For i = 1 To UBound(Frames)
        If Frames(i).Name = Name Then
            If Frames(i).Property Then
                IsProperty = True
                Exit Function
            End If
        End If
    Next i
End Function

Sub StatementReturn()
    Symbol "("
    Expression "$Intern.Return"
    PushContent "$Intern.Return"
    ExprJump CurrentFrame & ".end"
    Symbol ")"
    Terminator
    CodeBlock
End Sub

Sub DeclareFrame(Optional IsExport As Boolean, Optional NoCodeBlock As Boolean, Optional IsProto As Boolean, Optional IsProperty As Boolean)
    Dim pCount As Long
    Dim Ident As String
    Dim fAlias As String
    Dim Name As String
    Dim Method As String
    Dim IdentII As String
    Dim RetAs As String
    
    ArgCount = 0
    
    If IsProperty Then
        Method = Identifier
        If Method <> "set" And _
           Method <> "get" Then
            ErrMessage "property 'set'/'get'"
        Else
            Method = "." & Method
        End If
    End If
    
    Ident = Identifier
    AddFrame Ident & Method, IsProperty
   
    Symbol "(":
NextDeclare:
    IdentII = Identifier
    If IdentII <> "" Then
        VariableBlock IdentII, True
        ArgCount = ArgCount + 1
    End If
    
    If IsSymbol(",") Then
        Symbol (","): GoTo NextDeclare
    ElseIf IsSymbol(")") Then
        Symbol (")")
        If IsIdent("as") Then
            SkipIdent
            RetAs = Identifier
            If RetAs = "dword" Or RetAs = "single" Or RetAs = "string" Then
                Frames(UBound(Frames)).ReturnAs = RetAs
            Else
                ErrMessage "'" & RetAs & "' is not allowed as return type for '" & CurrentFrame & "'": Exit Sub
            End If
            SkipBlank
        End If
        Terminator
    Else
        ErrMessage "unexpected '" & Mid$(Source, Position, 1) & "'": Exit Sub
    End If
    
    If IsProto Then
        AddSymbol Ident & Method, OffsetOf(".code"), 0, ST_FRAME, True
        AddSymbol Ident & Method & ".Address", OffsetOf(".code"), Code, ST_DWORD, True
    Else
        AddSymbol Ident & Method, OffsetOf(".code"), 0, ST_FRAME
        AddSymbol Ident & Method & ".Address", OffsetOf(".code"), Code, ST_DWORD
    End If
    
    If Not IsProto Then If IsExport Then AddExport Ident & Method
    
    If Not IsProto Then
        StartFrame
        CodeBlock
        AddSymbol Ident & Method & ".end", OffsetOf(".code"), 0, ST_FRAME
        EndProc
        EndFrame ArgCount * 4
        If Not NoCodeBlock Then CodeBlock
    End If
End Sub

Sub EndProc()
    If IsIdent("end") Then
        SkipIdent
        Terminator
        CurrentFrame = ""
        DoEvents
        If Not IsCmdCompile Then frmMain.lblStatus.Caption = "Parsing.. (" & CInt(Position / Len(Source) * 100) & "% done.. | Position: " & Position & " )"
        Exit Sub
    Else
        ErrMessage "could not find end of '" & CurrentFrame & "'": Exit Sub
        Exit Sub
    End If
End Sub

Sub Align4(SectionName As String)
    Dim i As Long
    If OffsetOf(SectionName) = Int(OffsetOf(SectionName) / 4) * 4 Then Exit Sub
    For i = OffsetOf(SectionName) To Int((OffsetOf(SectionName) / 4) + 1) * 4 - 1
        AddSectionNameByte SectionName, 0
    Next i
End Sub
