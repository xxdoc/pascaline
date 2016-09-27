Attribute VB_Name = "comParser"
Option Explicit

Public Source As String
Public Position As Long
Public WithIdent As String
Public EntryPoint As String
Public UnsignedDeclare As Boolean

Sub InitParser()
    Dim i As Integer
    AppType = 0: pError = False: bLibrary = False: UnsignedDeclare = False
    CompareOne = "": CompareTwo = "": Source = ""
    For i = 1 To UBound(VirtualFiles)
        If Not VirtualFiles(i).Extension = EX_DIALOG Then
            Source = Source & "module " & Chr(34) & VirtualFiles(i).Name & Chr(34) & ";" & vbNewLine & VirtualFiles(i).Content & vbNewLine
        End If
    Next i
    Source = Replace(Source, vbTab, " ", 1, -1, vbTextCompare)
    Source = Replace(Source, " _ " & vbCrLf, "     ", 1, -1, vbTextCompare)
    Source = Replace(Source, " _" & vbCrLf, "    ", 1, -1, vbTextCompare)
    Source = Replace(Source, "_" & vbCrLf, "   ", 1, -1, vbTextCompare)
    Source = Source & vbNewLine
    Position = 1
End Sub

Sub Parse()
    StartCounter
    CurrentModule = "": CurrentFrame = "": CurrentType = "": EntryPoint = ""
    If Not bLibrary Then
        CreateSection ".data", Data, CH_INITIALIZED_DATA + CH_MEM_READ + CH_MEM_WRITE
        CreateSection ".code", Code, CH_CODE + CH_MEM_READ + CH_MEM_EXECUTE
        CreateSection ".idata", Import, CH_INITIALIZED_DATA + CH_MEM_READ + CH_MEM_WRITE
        CreateSection ".edata", Export, CH_INITIALIZED_DATA + CH_MEM_READ
        CreateSection ".rsrc", Resource, CH_INITIALIZED_DATA + CH_MEM_READ
        CreateSection ".reloc", Relocate, CH_MEM_DISCARDABLE + CH_INITIALIZED_DATA
        
        AssignProtoTypes
        DirectiveModule
        DirectiveApplication
        
        AddImport "lstrcpyA", "KERNEL32.DLL", 2, "lstrcpy"
        AddImport "lstrcmpA", "KERNEL32.DLL", 2, "lstrcmp"
        AddImport "wsprintfA", "USER32.DLL", -1, "Format"
        AddImport "ExitProcess", "KERNEL32.DLL", 1
        AddImport "GetModuleHandleA", "KERNEL32.DLL", 1, "GetModuleHandle"
        AddImport "HeapCreate", "KERNEL32.DLL", 3
        AddImport "HeapAlloc", "KERNEL32.DLL", 3
        AddImport "HeapDestroy", "KERNEL32.DLL", 1
        AddImport "RtlMoveMemory", "KERNEL32.DLL", 3, "MoveMemory"
        AddImport "MessageBoxA", "USER32.DLL", 4, "MessageBox"
        
        
        DeclareDataDWord "Instance", 0
        DeclareDataDWord "$Intern.Property", 0
        DeclareDataDWord "$Intern.Compare.One", 0
        DeclareDataDWord "$Intern.Compare.Two", 0
        DeclareDataDWord "$Intern.Float", 0
        DeclareDataDWord "$Intern.Array", 0
        DeclareDataDWord "$Intern.Loop", 0
        DeclareDataDWord "$Intern.Count", 0
        DeclareDataDWord "$Intern.Return", 0

        AddConstant "TRUE", -1: AddConstant "FALSE", 0
        AddConstant "NULL", 0
        
        AddCodeWord &H6A: InvokeByName "GetModuleHandle"
        AddCodeByte &HA3: AddCodeFixup "Instance"
        
        If Not IsDLL Then
            If EntryPoint = "" Then
                ExprCall "$entry"
            Else
                ExprCall EntryPoint
                Push 0
                InvokeByName "ExitProcess"
            End If
        Else
            InitializeDLL
        End If
    End If
    Call CodeBlock: If Not bLibrary And Not IsDLL Then EntryBlock: Call CodeBlock
End Sub

Sub CodeBlock()
    Dim Ident As String
    
    Ident = Identifier
    
    If Ident = "" Or pError = True Then Exit Sub
    
    Select Case LCase(Ident)
        Case "import": DeclareImport
        Case "const": DeclareConstant
        Case "type": DeclareType
        Case "frame": DeclareFrame
        Case "property": DeclareFrame False, False, False, True
        Case "export": DeclareFrame True
        Case "return": StatementReturn
        Case "if": StatementIf
        Case "while": StatementWhile
        Case "for": StatementFor
        Case "loop": StatementLoop
        Case "goto": StatementGoto
        Case "jump": StatementGoto
        Case "include": StatementInclude
        Case "library": StatementInclude
        Case "local": DeclareLocal
        Case "preserve": StatementPreserve
        Case "reserve": StatementReserve
        Case "destroy": StatementDestroy
        Case "direct": StatementDirect
        Case "bytes": StatementBytes
        Case "with": StatementWith
        Case "ubound": StatementUBound
        Case "lbound": StatementLBound
        Case "bitmap": DeclareBitmap
        Case "module": Position = Position - 6: DirectiveModule: CodeBlock: Exit Sub
        Case "end": Position = Position - 3: Exit Sub
        Case "end.": Position = Position - 4: Exit Sub
        Case "entry": Position = Position - 6: Exit Sub
        Case Else
            If IsImport(Ident) Then
                CallImport Ident
            ElseIf IsLocalVariable(Ident) Then
                EvalLocalVariable Ident
            ElseIf IsProperty(Ident & ".set") Then
                CallProperty Ident & ".set"
            ElseIf IsFrame(Ident) Then
                CallFrame Ident
            ElseIf IsVariable(Ident) Then
                EvalVariable Ident
            Else
                VariableBlock Ident
            End If
    End Select

   
End Sub

Sub VariableBlock(Ident As String, Optional FrameExpression As Boolean, Optional NoCodeBlock As Boolean)

    If Ident = "" Or pError = True Then Exit Sub
    
    Select Case LCase(Ident)
        Case "signed": UnsignedDeclare = False: Ident = Identifier: VariableBlock Ident, FrameExpression, NoCodeBlock: Exit Sub
        Case "unsigned": UnsignedDeclare = True: Ident = Identifier: VariableBlock Ident, FrameExpression, NoCodeBlock: Exit Sub
        Case "byte": DeclareVariable CurrentType, "byte", FrameExpression, NoCodeBlock
        Case "bool": DeclareVariable CurrentType, "byte", FrameExpression, NoCodeBlock
        Case "word": DeclareVariable CurrentType, "word", FrameExpression, NoCodeBlock
        Case "dword": DeclareVariable CurrentType, "dword", FrameExpression, NoCodeBlock
        Case "single": DeclareVariable CurrentType, "single", FrameExpression, NoCodeBlock
        Case "string": DeclareString CurrentType, FrameExpression, NoCodeBlock
        Case "boolean": DeclareVariable CurrentType, "byte", FrameExpression, NoCodeBlock
        Case Else
            If IsType(Ident) Then
                AssignType Identifier, Ident
            Else
                ErrMessage "unknown identifier -> '" & Ident & "'": Exit Sub
            End If
    End Select
    
    UnsignedDeclare = False
End Sub

Function Identifier() As String
    Dim Value As String
    SkipBlank
    Value = Mid$(Source, Position, 1)
    If Value = "." And WithIdent <> "" Then Identifier = WithIdent
    While (UCase(Value) >= "A" And _
           UCase(Value) <= "Z") Or _
           Value = "." Or _
           Value = "_"
            While IsNumeric(Mid$(Source, Position + 1, 1))
                Identifier = Identifier & Mid$(Source, Position, 1)
                Position = Position + 1
            Wend
            Identifier = Identifier & Mid$(Source, Position, 1)
            Position = Position + 1
            Value = Mid$(Source, Position, 1)
    Wend
    If IsSymbol(":") Then DeclareLabel Identifier: Identifier = Identifier()
End Function

Sub Skip(Optional NumberOfChars As Integer)
    Position = Position + 1 + NumberOfChars
End Sub

Sub SkipBlank()
    Dim Value As String
    Dim Value2 As String
    Value = Mid$(Source, Position, 1)
    Value2 = Mid$(Source, Position, 2)
    While Value = " " Or _
          Value = vbCr Or _
          Value = vbLf Or _
          Value = vbTab Or _
          Value2 = "//"
          If Value2 = "//" Then
            While Mid$(Source, Position, 2) <> vbCrLf
                Position = Position + 1: If Position >= Len(Source) Then ErrMessage "found end of code": Exit Sub
            Wend
          End If
        Position = Position + 1
        Value = Mid$(Source, Position, 1)
        Value2 = Mid$(Source, Position, 2)
    Wend
End Sub

Sub SkipIdent()
    Call Identifier
End Sub

Sub Symbol(Value As String)
    SkipBlank
    If Mid$(Source, Position, 1) = Value Then
        Position = Position + 1
    Else
        ErrMessage "expected symbol '" & Value & "' but found '" & Mid$(Source, Position, 1) & "'": Exit Sub
    End If
End Sub

Function IsIdent(Word As String) As Boolean
    SkipBlank
    If Mid$(Source, Position, Len(Word)) = Word Then IsIdent = True
End Function

Function IsSymbol(Value As String) As Boolean
    If Mid$(Source, Position, Len(Value)) = Value Then IsSymbol = True
End Function

Sub Blank()
    If Mid$(Source, Position, 1) = " " Then
        Position = Position + 1
    Else
        ErrMessage "expected blank ' ' but found '" & Mid$(Source, Position, 1) & "'": Exit Sub
    End If
End Sub

Sub Terminator()
    If Mid$(Source, Position, 1) = ";" Then
        Position = Position + 1
    Else
        ErrMessage "expected terminator (;) but found '" & Mid$(Source, Position, 1) & "'": Exit Sub
    End If
End Sub

Function IsVariable(Name As String) As Boolean
    Dim i As Integer
    For i = 0 To UBound(Symbols)
        If Symbols(i).Name = Name Then
            If Symbols(i).SymType = ST_BYTE Or _
               Symbols(i).SymType = ST_WORD Or _
               Symbols(i).SymType = ST_DWORD Or _
               Symbols(i).SymType = ST_SINGLE Or _
               Symbols(i).SymType = ST_US_BYTE Or _
               Symbols(i).SymType = ST_US_WORD Or _
               Symbols(i).SymType = ST_US_DWORD Or _
               Symbols(i).SymType = ST_STRING Then
                IsVariable = True
                Exit Function
            End If
            Exit Function
        End If
    Next i
End Function

Function IsEndOfCode(Value As Long) As Boolean
    If Value > Len(Source) Then
        ErrMessage "found end of code. but expected ')' or ','": IsEndOfCode = True: Exit Function
    End If
End Function

Function IsVariableExpression() As Boolean
    If (UCase(Mid$(Source, Position, 1)) >= "A" And _
           UCase(Mid$(Source, Position, 1)) <= "Z") Then
        IsVariableExpression = True
    End If
End Function

Function IsStringExpression() As Boolean
    If Mid$(Source, Position, 1) = Chr(34) Then
        IsStringExpression = True
    End If
End Function

Function IsFloatExpression() As Boolean
    Dim i As Integer
    Dim StrFloat As String
    Dim OPosition As Long
    OPosition = Position
    While IsNumeric(Mid$(Source, Position, 1)) Or _
                 Mid$(Source, Position, 1) = "-" Or _
                 Mid$(Source, Position, 1) = "."
        StrFloat = StrFloat & Mid$(Source, Position, 1)
        Position = Position + 1
    Wend
    Position = OPosition
    If InStr(1, StrFloat, ".", vbTextCompare) <> 0 Then
        IsFloatExpression = True
    End If
End Function

Function IsNumberExpression() As Boolean
    Dim i As Integer
    If IsNumeric(Mid$(Source, Position, 1)) Or _
                 Mid$(Source, Position, 1) = "-" Or _
                 Mid$(Source, Position, 1) = "$" Then
        IsNumberExpression = True
    ElseIf IsConstantExpression Then
        For i = 1 To UBound(Constants)
            If IsIdent(Constants(i).Name) Then
                IsNumberExpression = True
                Exit Function
            End If
        Next i
    End If
End Function

Function IsConstantExpression() As Boolean
    If (UCase(Mid$(Source, Position, 1)) >= "A" And _
           UCase(Mid$(Source, Position, 1)) <= "Z") Then
        IsConstantExpression = True
    ElseIf IsSymbol("[") Then
        IsConstantExpression = True
    End If
End Function

Function NumberExpression() As Variant
    Dim i As Integer
    Dim CToHex As Boolean
    Dim Str2Hex As String
    SkipBlank
    If IsSymbol("$") Then
        Symbol "$"
        CToHex = True
        While IsNumeric(Mid$(Source, Position, 1)) Or Mid$(Source, Position, 1) = "-" Or _
              Mid$(Source, Position, 1) = "A" Or Mid$(Source, Position, 1) = "B" Or _
              Mid$(Source, Position, 1) = "C" Or Mid$(Source, Position, 1) = "D" Or _
              Mid$(Source, Position, 1) = "E" Or Mid$(Source, Position, 1) = "F"
                NumberExpression = NumberExpression & Mid$(Source, Position, 1)
                Position = Position + 1
        Wend
    Else
        While IsNumeric(Mid$(Source, Position, 1)) Or Mid$(Source, Position, 1) = "-" Or Mid$(Source, Position, 1) = "."
                Dim IsNegative As Boolean
                Dim IsFloatRec As Boolean
                Dim AfterPoint As String
                If Mid$(Source, Position, 1) = "." Then
                    Position = Position + 1
                    While IsNumeric(Mid$(Source, Position, 1))
                        AfterPoint = AfterPoint & Mid$(Source, Position, 1)
                        Position = Position + 1
                    Wend
                    Position = Position - 1
                    If CSng("0.1") = 0.1 Then           'Check for American System
                        AfterPoint = "0." & AfterPoint
                    ElseIf CSng("0,1") = 0.1 Then       'Check for German System
                        AfterPoint = "0," & AfterPoint
                    End If
                    NumberExpression = CSng(NumberExpression) + CSng(AfterPoint)
                    If IsNegative = True Then NumberExpression = NumberExpression * (-1)
                ElseIf Mid$(Source, Position, 1) = "-" Then
                    IsNegative = True
                Else
                    NumberExpression = NumberExpression & Mid$(Source, Position, 1)
                End If
                Position = Position + 1
        Wend
        If AfterPoint = "" Then
            If IsNegative Then
                NumberExpression = NumberExpression * (-1)
            End If
        End If
    End If
    
    If CToHex = True Then NumberExpression = CLng("&H" & NumberExpression)
    If NumberExpression = 0 Then
        For i = 1 To UBound(Constants)
            If IsIdent(Constants(i).Name) Then
                NumberExpression = GetConstant(Identifier)
                Exit Function
            End If
        Next i
    End If
End Function

Sub InsertSource(sISource As String)
    Dim Header As String: Dim Footer As String
    Header = Mid$(Source, 1, Position - 1)
    Footer = Mid$(Source, Position, Len(Source) - Position + 1)
    Source = Header & sISource & Footer
End Sub

Function ConstantExpression() As Long
    If IsSymbol("[") Then
        Symbol "["
        While Not IsSymbol("]")
            ConstantExpression = NumberExpression
            If IsSymbol("+") Then
                Symbol "+"
                ConstantExpression = ConstantExpression + NumberExpression
            ElseIf IsSymbol("-") Then
                Symbol "-"
                ConstantExpression = ConstantExpression - NumberExpression
            ElseIf IsSymbol("|") Then
                Symbol "|"
                If IsSymbol("!") Then
                    Symbol "!"
                    ConstantExpression = ConstantExpression Or Not NumberExpression
                Else
                    ConstantExpression = ConstantExpression Or NumberExpression
                End If
            ElseIf IsSymbol("&") Then
                Symbol "&"
                If IsSymbol("!") Then
                    Symbol "!"
                    ConstantExpression = ConstantExpression And Not NumberExpression
                Else
                    ConstantExpression = ConstantExpression And NumberExpression
                End If
            ElseIf IsSymbol("~") Then
                Symbol "~"
                ConstantExpression = ConstantExpression Xor NumberExpression
            Else
                ErrMessage "invalid constant value": Exit Function
            End If
            If Position >= Len(Source) Then ErrMessage "found end of code. but expected ')' or ','": Exit Function
        Wend
        Symbol "]"
    Else
    ConstantExpression = GetConstant(Identifier)
    End If
End Function

Function VariableExpression() As String
    SkipBlank
    VariableExpression = Identifier
End Function

Function StringExpression() As String
    Dim Value As String
    SkipBlank
    Symbol Chr(34)
    Value = Mid$(Source, Position, 1)
    While Value <> Chr(34)
        StringExpression = StringExpression & Mid$(Source, Position, 1)
        Position = Position + 1
        Value = Mid$(Source, Position, 2)
        If Value = "\n" Then Position = Position + 2: StringExpression = StringExpression & vbCrLf
        If Value = "\t" Then Position = Position + 2: StringExpression = StringExpression & vbTab
        Value = Mid$(Source, Position, 1)
        If Value = vbCr Or _
           Value = "" Then
            ErrMessage "unterminated string": Exit Function
        End If
    Wend
    Symbol Chr(34)
End Function





