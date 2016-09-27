Attribute VB_Name = "comSyntax"
Option Explicit

Public CurrentSection As String
Public CurrentModule As String

Sub DirectiveModule()
    Dim ModName As String
    If IsIdent("module") Then
        SkipIdent
        CurrentModule = StringExpression
        Terminator
    End If
End Sub

Sub DirectiveApplication()
    If IsIdent("application") Then
        bLibrary = False
        SkipIdent
        If IsIdent("PE") Then
            SkipIdent
            If IsIdent("GUI") Then
                SkipIdent
                AppType = GUI
            ElseIf IsIdent("CUI") Then
                SkipIdent
                AppType = CUI
            Else
                ErrMessage "invalid format '" & Identifier & "' expected 'GUI' or 'CUI'": Exit Sub
            End If
            SkipBlank
            If IsIdent("DLL") Then
                SkipIdent
                IsDLL = True
            End If
            If IsIdent("entry") Then
                DeclareEntryPoint
            Else
                Terminator
            End If
        Else
            ErrMessage "expected 'PE' but found '" & Identifier & "'": Exit Sub
        End If
    ElseIf IsIdent("library") Then
        SkipIdent
        bLibrary = True
        SkipBlank
        LibraryName = StringExpression
        Terminator
    Else
        ErrMessage "expected 'application' or 'library' but found '" & Identifier & "'": Exit Sub
    End If
End Sub

Sub EntryBlock()
    Dim Ident As String
    If Not EntryPoint = "" Or EntryPoint = "entry" Then Exit Sub
    Ident = Identifier
    If Ident = "entry" Then
        AddSymbol "$entry", OffsetOf(".code"), Code, ST_LABEL
    Else
        ErrMessage "expected 'entry' but found '" & Ident & "'": Exit Sub
    End If
    CodeBlock
    Ident = Identifier
    If Not Ident = "end." Then
        ErrMessage "expected 'end.' but found '" & Ident & "'": Exit Sub
    Else
        Push 0
        InvokeByName "ExitProcess"
    End If
End Sub

Sub DeclareEntryPoint()
    SkipIdent
    EntryPoint = Identifier
    Terminator
End Sub

Sub DirectiveSection()
    Dim Name As String
    Dim Ident As String
    Dim ST As ENUM_SECTION_TYPE
    Dim CH As ENUM_SECTION_CHARACTERISTICS
    
    Name = StringExpression
    If SectionExists(Name) Then GoTo DirSectionExists
    Blank
    Ident = Identifier
    
    Select Case LCase(Ident)
        Case "data": ST = Data: CH = CH + CH_INITIALIZED_DATA
        Case "code": ST = Code: CH = CH + CH_CODE
        Case "import": ST = Import
        Case "export": ST = Export
        Case "resource": ST = Resource
        Case Else
            ErrMessage "invalid section type '" & Ident & "'": Exit Sub
    End Select
    
    If IsSymbol(" ") Then
JCharacteristic:
        Blank
        Ident = Identifier
    
        Select Case LCase(Ident)
            Case "code": CH = CH + CH_CODE
            Case "data": CH = CH + CH_INITIALIZED_DATA
            Case "udata": CH = CH + CH_UNINITIALIZED_DATA
            Case "discardable": CH = CH + CH_MEM_DISCARDABLE
            Case "executable": CH = CH + CH_MEM_EXECUTE
            Case "notchached": CH = CH + CH_MEM_NOT_CHACHED
            Case "notpaged": CH = CH + CH_MEM_NOT_PAGED
            Case "readable": CH = CH + CH_MEM_READ
            Case "shared": CH = CH + CH_MEM_SHARED
            Case "writeable": CH = CH + CH_MEM_WRITE
            Case Else
                ErrMessage "invalid characteristic '" & Ident & "'": Exit Sub
        End Select
    End If
    If IsSymbol(" ") Then GoTo JCharacteristic
    CurrentSection = Name
DirSectionExists:
    CreateSection Name, ST, CH
    Terminator
    CodeBlock
End Sub

Sub CreateSection(Name As String, SectionType As ENUM_SECTION_TYPE, Characteristics As ENUM_SECTION_CHARACTERISTICS)
    Dim i As Integer
    If SectionExists(Name) Then
        CurrentSection = Section(SectionID(Name)).Name
    Else
        ReDim Preserve Section(UBound(Section) + 1) As TYPE_SECTION
        ReDim Section(UBound(Section)).Bytes(0)
        Section(UBound(Section)).Name = Name
        Section(UBound(Section)).SectionType = SectionType
        Section(UBound(Section)).Characteristics = Characteristics
    End If
End Sub

Sub DeclareLabel(Name As String)
    AddSymbol Name, OffsetOf(".code"), Code, ST_LABEL
    Symbol ":"
End Sub

Sub StatementGoto()
    ExprJump Identifier
    Terminator
    CodeBlock
End Sub

Sub DeclareString(Optional CurrentType As String, Optional FrameExpression As Boolean, Optional NoCodeBlock As Boolean)
    Dim FullName As String: Dim Ident As String: Dim Space As Long: Dim Value As String
    
    Ident = Identifier
    FullName = Switch(CurrentType = "", Ident, CurrentType <> "", CurrentType & "." & Ident)
    
    If IsSymbol("=") Then Symbol "=": Value = StringExpression Else: Value = ""
    If IsSymbol("[") Then Symbol "[": Space = NumberExpression: Symbol "]" Else: Space = 256
        
    If IsSymbol("(") Then
        Symbol "("
        If Not IsSymbol(")") Then
            If CurrentFrame = "" Then
                ErrMessage "you cannot dimension the array outside of a frame. use '()' instead and ""reserve " & FullName & "([Size])"" inside a frame."
                Exit Sub
            End If
        End If
        DeclareDataString FullName, Value, Space
        ReserveArray FullName, NumberExpression
        Symbol ")"
    End If
    
    If FrameExpression = False Then
        Terminator
        If Not SymbolExists(FullName) Then
            DeclareDataString FullName, Value, Space
        End If
    Else
        If Not SymbolExists(CurrentFrame & "." & FullName) Then
            AddSymbol CurrentFrame & "." & FullName, 8 + (ArgCount * 4), 0, ST_LOCAL_STRING
            AddFrameDeclare FullName
        End If
    End If
    
    If Not NoCodeBlock Then CodeBlock
End Sub

Sub DeclareVariable(Optional CurrentType As String, Optional Size As String, Optional FrameExpression As Boolean, Optional NoCodeBlock As Boolean)
    Dim FullName As String: Dim Value As Single: Dim Ident As String
    
    Ident = Identifier
    
    FullName = Switch(CurrentType = "", Ident, CurrentType <> "", CurrentType & "." & Ident)
    
    If IsSymbol("=") Then Symbol "=": Value = NumberExpression Else: Value = 0
    
    If IsSymbol("(") Then
        Symbol "("
        If Not IsSymbol(")") Then If CurrentFrame = "" Then ErrMessage "you cannot dimension the array outside of a frame. use '()' instead and ""reserve " & FullName & "([Size])"" inside a frame.": Exit Sub
        If UnsignedDeclare Then
            Select Case Size
                Case "byte": DeclareDataUnsignedByte FullName, CByte(Value)
                Case "word": DeclareDataUnsignedWord FullName, CInt(Value)
                Case Else: ErrMessage "invalid size '" & FullName & "'": Exit Sub
            End Select
        Else
            Select Case Size
                Case "byte": DeclareDataByte FullName, CByte(Value)
                Case "word": DeclareDataWord FullName, CInt(Value)
                Case "dword": DeclareDataDWord FullName, CLng(Value)
                Case "single": DeclareDataSingle FullName, CSng(Value)
                Case Else: ErrMessage "invalid size '" & FullName & "'": Exit Sub
            End Select
        End If
        ReserveArray FullName, NumberExpression
        Symbol ")"
    End If
    
    If FrameExpression = False Then
        If Not SymbolExists(FullName) Then
            If UnsignedDeclare Then
                Select Case Size
                    Case "byte": DeclareDataUnsignedByte FullName, CByte(Value)
                    Case "word": DeclareDataUnsignedWord FullName, CInt(Value)
                    Case Else: ErrMessage "invalid size '" & FullName & "'": Exit Sub
                End Select
            Else
                Select Case Size
                    Case "byte": DeclareDataByte FullName, CByte(Value)
                    Case "word": DeclareDataWord FullName, CInt(Value)
                    Case "dword": DeclareDataDWord FullName, CLng(Value)
                    Case "single": DeclareDataSingle FullName, CSng(Value)
                    Case Else: ErrMessage "invalid size '" & FullName & "'": Exit Sub
                End Select
            End If
        End If
        If IsSymbol(",") Then Symbol ",": DeclareVariable CurrentType, Size, FrameExpression: Exit Sub
        Terminator
    Else
        If Not SymbolExists(CurrentFrame & "." & FullName) Then
            Select Case Size
                Case "single": AddSymbol CurrentFrame & "." & FullName, 8 + (ArgCount * 4), 0, ST_LOCAL_SINGLE
                Case Else: AddSymbol CurrentFrame & "." & FullName, 8 + (ArgCount * 4), 0, ST_LOCAL_DWORD
            End Select
            AddFrameDeclare Ident
        End If
    End If
    If Not NoCodeBlock Then CodeBlock
End Sub

Sub DeclareLocal()
    Dim Ident As String
    Dim IdentII As String
    Dim Value As Variant
    Dim Space As Long
    Dim ArrayValue As Long
    
    Ident = Identifier
    IdentII = Identifier
    
    If CurrentFrame = "" Then ErrMessage "cannot declare local variable '" & IdentII & "' outside of a frame": Exit Sub
    
    If Ident = "byte" Or _
       Ident = "word" Or _
       Ident = "bool" Or _
       Ident = "dword" Or _
       Ident = "boolean" Then
        AddSymbol CurrentFrame & "." & IdentII, 8 + (ArgCount * 4), 0, ST_LOCAL_DWORD
        ArgCount = ArgCount + 1
    ElseIf Ident = "single" Then
        AddSymbol CurrentFrame & "." & IdentII, 8 + (ArgCount * 4), 0, ST_LOCAL_SINGLE
        ArgCount = ArgCount + 1
    ElseIf Ident = "string" Then
        If IsSymbol("[") Then
            Symbol "["
            Space = NumberExpression
            Symbol "]"
        Else
            Space = 256
        End If
        AddSymbol CurrentFrame & "." & IdentII, 8 + (ArgCount * 4), 0, ST_LOCAL_STRING
        eUniqueID = eUniqueID + 1
        DeclareDataString "Local.String" & eUniqueID, "", Space
        MovEAXAddress "Local.String" & eUniqueID
        'mov [ebp+number],eax
        AddCodeWord &H8589
        AddCodeDWord 8 + (ArgCount * 4)
        ArgCount = ArgCount + 1
    Else
        ErrMessage "expected identifier 'byte','word','dword','single' or 'string' but found" & Ident: Exit Sub
    End If
    
    Terminator
        

    CodeBlock
End Sub

Sub DeclareConstant()
    Dim Name As String
    Name = Identifier
    Symbol "="
    SkipBlank
    If IsStringExpression Then
        AddConstant Name, StringExpression
    ElseIf IsNumberExpression Then
        AddConstant Name, NumberExpression
    ElseIf IsConstantExpression Then
        AddConstant Name, ConstantExpression
    Else
        ErrMessage "invalid constant value. '" & Name & " '": Exit Sub
    End If
    Terminator
    CodeBlock
End Sub

Sub DeclareType()
    Dim Name As String
    Dim Ident As String
    Dim TypeSource As String
    
    ReDim Preserve Types(UBound(Types) + 1) As TYPE_TYPE
    Types(UBound(Types)).Name = Identifier
    
    Symbol "{"
    
    While Not IsSymbol("}")
        SkipBlank
        If IsIdent("string") Or IsIdent("dword") Or IsIdent("word") Or _
           IsIdent("byte") Or IsIdent("bool") Or IsIdent("boolean") Or IsIdent("single") Then
            TypeSource = TypeSource & Identifier & " " & Identifier
            If IsSymbol("[") Then
                Symbol "[": TypeSource = TypeSource & "["
                TypeSource = TypeSource & NumberExpression & "]"
                Symbol "]"
            ElseIf IsSymbol("(") Then
                Symbol "(": TypeSource = TypeSource & "("
                Symbol ")": TypeSource = TypeSource & ")"
            End If
            TypeSource = TypeSource & ";"
            Terminator
            SkipBlank
        Else
            Ident = Identifier
            If IsType(Ident) Then
                TypeSource = TypeSource & Ident & " " & Identifier
                TypeSource = TypeSource & ";"
                Terminator
                SkipBlank
            Else
                Symbol "}"
                Exit Sub
            End If
        End If
    Wend
    Types(UBound(Types)).Source = TypeSource
    Symbol "}"
    CodeBlock
End Sub

Sub StatementInclude()
    Dim FileName As String: FileName = StringExpression
    Position = Position - (Len(FileName)) - 2
    If Right(FileName, 4) = ".lib" Then
        ImportLibrary
    Else
        IncludeFile
    End If
End Sub

Sub IncludeFile()
    Dim i As Integer: Dim Files() As String: Dim Content As String
    
    ReDim Files(0) As String
IncludeMore:
    ReDim Preserve Files(UBound(Files) + 1) As String
    Files(UBound(Files)) = StringExpression
    If IsSymbol(",") Then Position = Position + 1: SkipBlank: GoTo IncludeMore
    Terminator
    
    For i = UBound(Files) To 1 Step -1
        If Dir(App.Path & "\include\" & Files(i)) = "" Then ErrMessage "cannot include '" & Files(i) & "'. check your include folder.": Exit Sub
        Open App.Path & "\include\" & Files(i) For Binary As #1
            Content = Space(LOF(1))
            Get #1, , Content
        Close #1
        InsertSource Content
        LenIncludes = LenIncludes + Len(Content)
    Next i
    CodeBlock
End Sub

Sub StatementIf()
    Dim iID As Long: iID = iID + lUniqueID: lUniqueID = lUniqueID + 1
    Dim elseifcount As Long
    
    Symbol "("
    
    Expression "$Intern.Compare.One"
    Expression "$Intern.Compare.Two"
    
    If IsStringCompare Then
        ExprCompareS "$Intern.Compare.One", "$Intern.Compare.Two"
    Else
        ExprCompare "$Intern.Compare.One", "$Intern.Compare.Two"
    End If
    
    ChooseRelation iID, "$else"
    Symbol ")":
    Symbol "{"
        AddSymbol "$then" & iID, OffsetOf(".code"), Code, ST_LABEL
        CodeBlock
        ExprJump "$out" & iID
    Symbol "}"
    SkipBlank
    If IsIdent("else") Then
        SkipIdent
        Symbol "{"
        AddSymbol "$else" & iID, OffsetOf(".code"), Code, ST_LABEL
        CodeBlock
        Symbol "}"
    Else
        AddSymbol "$else" & iID, OffsetOf(".code"), Code, ST_LABEL
    End If
        AddSymbol "$out" & iID, OffsetOf(".code"), Code, ST_LABEL
    CodeBlock
End Sub

Sub StatementWhile()
    Dim wID As Long: wID = wID + lUniqueID: lUniqueID = lUniqueID + 1
    Symbol "("
    
    AddSymbol "$swhile" & wID, OffsetOf(".code"), Code, ST_LABEL
   
    Expression "$Intern.Compare.One"
    Expression "$Intern.Compare.Two"
    
    If IsStringCompare Then
        ExprCompareS "$Intern.Compare.One", "$Intern.Compare.Two"
    Else
        ExprCompare "$Intern.Compare.One", "$Intern.Compare.Two"
    End If
    
    ChooseRelation wID, "$endwhile"
    
    Symbol ")"
    Symbol "{"
    
    AddSymbol "$while" & wID, OffsetOf(".code"), Code, ST_LABEL
    CodeBlock
    ExprJump "$swhile" & wID
    Symbol "}"
    SkipBlank
    
    AddSymbol "$endwhile" & wID, OffsetOf(".code"), Code, ST_LABEL
    CodeBlock
End Sub

Sub StatementFor()
    Dim fID As Long: fID = fID + lUniqueID: lUniqueID = lUniqueID + 1
    Dim Ident As String: Dim sExpression As String
    
    Symbol "("
    
    Ident = Identifier
    SkipBlank
    If IsSymbol("(") Then
        SetArray Ident
    Else
        If IsVariable(Ident) Then
            EvalVariable Ident, True
        ElseIf IsLocalVariable(Ident) Then
            EvalLocalVariable Ident, True
        End If
    End If
    
    IsStringCompare = False
    AddSymbol "$sfor" & fID, OffsetOf(".code"), Code, ST_LABEL
    
    Expression "$Intern.Compare.One"
    Expression "$Intern.Compare.Two"

    If IsStringCompare Then
        ExprCompareS "$Intern.Compare.One", "$Intern.Compare.Two"
    Else
        ExprCompare "$Intern.Compare.One", "$Intern.Compare.Two"
    End If
    

    ChooseRelation fID, "$endfor"
    Terminator
       
    While Not IsSymbol(")")
        sExpression = sExpression & Mid$(Source, Position, 1)
        
        If IsSymbol("(") Then
            Position = Position + 1
            While Not IsSymbol(")")
                sExpression = sExpression & Mid$(Source, Position, 1)
                Position = Position + 1
            Wend
                sExpression = sExpression & Mid$(Source, Position, 1)
        End If
        If Position >= Len(Source) Then ErrMessage "found end of code. but expected ')' or ','": Exit Sub
        Position = Position + 1
    Wend
    Symbol ")"
    Symbol "{"
    AddSymbol "$for" & fID, OffsetOf(".code"), Code, ST_LABEL
    CodeBlock
    InsertSource sExpression & ";": CodeBlock
    ExprJump "$sfor" & fID
    Symbol "}"
    SkipBlank
    AddSymbol "$endfor" & fID, OffsetOf(".code"), Code, ST_LABEL
    CodeBlock
End Sub

Sub StatementLoop()
    Dim Ident As String
    Dim Mode As String
    Dim iID As Long
    
    iID = iID + lUniqueID
    lUniqueID = lUniqueID + 1
    
    Mode = Identifier
    
    If Mode = "until" Then
        Symbol "("
        Expression "$Intern.Compare.One"
        Expression "$Intern.Compare.Two"
        Symbol ")"
        
        AddSymbol "$loop" & iID, OffsetOf(".code"), Code, ST_LABEL
        Symbol "{"
        CodeBlock
        Symbol "}"
        
        ExprCompare "$Intern.Compare.One", "$Intern.Compare.Two"
        ChooseRelation iID, "$loop"
        
        AddSymbol "$loopout" & iID, OffsetOf(".code"), Code, ST_LABEL
    ElseIf Mode = "down" Or Mode = "" Then
        Symbol "("
        Expression
        PopECX
        If IsSymbol(",") Then Skip: Ident = Identifier
        Symbol ")"
        
        AddSymbol "$loop" & iID, OffsetOf(".code"), Code, ST_LABEL
        
        Symbol "{"
        CodeBlock
        Symbol "}": DecECX
        
        If Ident <> "" Then
            'mov [Variable],ecx
            AddCodeWord &HD89: AddCodeFixup Ident
        End If
        
        'cmp ecx,0
        AddCodeWord &HF983: AddCodeByte 0
        ExprJA "$loop" & iID
    ElseIf Mode = "up" Then
        AddCodeByte &HB9: AddCodeDWord 0 'mov ecx,0
        
        Symbol "("
        Expression "$Intern.Count"
        If IsSymbol(",") Then Skip: Ident = Identifier
        Symbol ")"
        
        AddSymbol "$loop" & iID, OffsetOf(".code"), Code, ST_LABEL
        
        Symbol "{"
        CodeBlock
        Symbol "}": IncECX
        
        If Ident <> "" Then
            'mov [Variable],ecx
            AddCodeWord &HD89: AddCodeFixup Ident
        End If
        
        'cmp ecx,[variable]
        AddCodeWord &HD3B: AddCodeFixup "$Intern.Count"
        ExprJL "$loop" & iID
    Else
        ErrMessage "expected loop 'up' or 'down' but found '" & Mode & "'": Exit Sub
    End If

    CodeBlock
End Sub

Sub StatementDirect()
    Dim Ident As String
    Dim AddrIdent As String
    
    Dim Variable As String
    
    Symbol "["
    Ident = Identifier
NextDirect:
    If Ident = "single" Then
        AddCodeSingle CSng(NumberExpression)
    ElseIf Ident = "dword" Then
        AddCodeDWord CLng(NumberExpression)
    ElseIf Ident = "word" Then
        AddCodeWord LoWord(NumberExpression)
    ElseIf Ident = "byte" Then
        AddCodeByte LoByte(LoWord(NumberExpression))
    ElseIf Ident = "address" Then
        AddrIdent = Identifier
        AddCodeFixup AddrIdent
    Else
        ErrMessage "data type must be specified 'single', 'dword', 'word', 'byte', 'address'"
    End If
    SkipBlank
    If IsSymbol(",") Then Position = Position + 1: GoTo NextDirect
    Symbol "]"
    Terminator
    CodeBlock
End Sub

Sub StatementBytes()
    Dim Ident As String
    Dim bByte As Long
    Ident = Identifier
    Symbol "["
NextBytes:
    AddDataByte NumberExpression
    If IsSymbol("@") Then Position = Position + 1: AddSymbol Ident, OffsetOf(".data"), Data, ST_DWORD
    If IsSymbol(",") Then Position = Position + 1: GoTo NextBytes
    Symbol "]"
    Terminator
    CodeBlock
End Sub

Sub EvalVariable(Name As String, Optional OnlySet As Boolean)
    SkipBlank
    
    If IsSymbol("(") Then
        SetArray Name
        Terminator
        CodeBlock
        Exit Sub
    End If
    
    If IsSymbol("=") Then
        Symbol "="
        Expression Name
    ElseIf IsSymbol("+") Then
        Symbol "+"
        If IsSymbol("+") Then
            Symbol "+"
            'inc [Variable]
            AddCodeWord &H5FF
            AddCodeFixup Name
        Else
            'add [Name],Value
            AddCodeWord &H581
            AddCodeFixup Name
            AddCodeDWord NumberExpression
        End If
    ElseIf IsSymbol("-") Then
        Symbol "-"
        If IsSymbol("-") Then
            Symbol "-"
            'dec [Variable]
            AddCodeWord &HDFF
            AddCodeFixup Name
        Else
            'sub [Name],Value
            AddCodeWord &H2D81
            AddCodeFixup Name
            AddCodeDWord NumberExpression
        End If
    End If
    Terminator
    If Not OnlySet Then CodeBlock
End Sub

Sub EvalLocalVariable(Name As String, Optional OnlySet As Boolean)
    Dim iLabel As Long
    SkipBlank
    If IsSymbol("=") Then
        Symbol "="
        Expression
        PopEAX
    ElseIf IsSymbol("+") Then
        Symbol "+"
        'mov eax, [ebp+number]
        AddCodeWord &H858B
        AddCodeDWord GetSymbolOffset(CurrentFrame & "." & Name)
        AddCodeByte &H5
        If IsSymbol("+") Then
            Symbol "+": AddCodeDWord &H1
        Else
            AddCodeDWord NumberExpression
        End If
    ElseIf IsSymbol("-") Then
        Symbol "-"
        AddCodeWord &H858B
        AddCodeDWord GetSymbolOffset(CurrentFrame & "." & Name)
        AddCodeByte &H2D
        If IsSymbol("-") Then
            Symbol "-": AddCodeDWord &H1
        Else
            AddCodeDWord NumberExpression
        End If
    End If
    
    'mov [ebp+number], eax
    AddCodeWord &H8589
    AddCodeDWord GetSymbolOffset(CurrentFrame & "." & Name)
    Terminator
    If Not OnlySet Then CodeBlock
End Sub

Sub StatementWith()
    WithIdent = Identifier
    Symbol "{"
        CodeBlock
    Symbol "}"
    WithIdent = ""
    CodeBlock
End Sub

Sub DeclareImport()
    Dim Ident As String
    Dim OIdent As String
    Dim FunctionName As String
    Dim FunctionAlias As String
    Dim Library As String
    Dim ParamCount As Long
    
    FunctionAlias = ""
    Ident = Identifier
    OIdent = Identifier
    
    If OIdent = "alias" Then
        FunctionAlias = Ident
        FunctionName = Identifier
        OIdent = Identifier
    Else
        FunctionAlias = Ident
        If OIdent = "ascii" Then
            OIdent = Identifier
            FunctionName = Ident & "A"
        ElseIf OIdent = "unicode" Then
            OIdent = Identifier
            FunctionName = Ident & "W"
        Else
            FunctionName = Ident
        End If
    End If
    
    If OIdent = "lib" Or OIdent = "library" Then
        Library = StringExpression
    Else
        ErrMessage "expected 'lib' but found '" & Ident & "'": Exit Sub
        Exit Sub
    End If
    
    If IsSymbol(",") Then
        Position = Position + 1: ParamCount = NumberExpression
    Else
        ParamCount = 0
    End If
        
    Terminator
    
    AddImport FunctionName, Library, ParamCount, FunctionAlias

    CodeBlock
    
End Sub
