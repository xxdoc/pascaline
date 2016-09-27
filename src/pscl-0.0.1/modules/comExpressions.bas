Attribute VB_Name = "comExpressions"
Option Explicit

Public IsFloat As Boolean
Public Relation As String
Public eUniqueID As Long
Public CompareOne As String
Public CompareTwo As String
Public Assignment As String
Public BracketsOpen As Byte
Public EvaluateCount As Long
Public IsStringCompare As Boolean

Sub ChooseRelation(iID As Long, LabelElse As String)
    If Relation = "=" Then
        ExprJNE LabelElse & iID
    ElseIf Relation = "!=" Then
        ExprJE LabelElse & iID
    ElseIf Relation = "<>" Then
        ExprJE LabelElse & iID
    ElseIf Relation = "<" Then
        ExprJAE LabelElse & iID
    ElseIf Relation = ">" Then
        ExprJLE LabelElse & iID
    ElseIf Relation = ">=" Then
        ExprJL LabelElse & iID
    ElseIf Relation = "<=" Then
        ExprJA LabelElse & iID
    ElseIf Relation = "=>" Then
        ExprJL LabelElse & iID
    ElseIf Relation = "=<" Then
        ExprJA LabelElse & iID
    End If
End Sub

Function IsOperatorAdd() As Boolean
    IsOperatorAdd = IsSymbol("+") Or IsSymbol("-") Or IsIdent("add") Or IsIdent("sub")
End Function

Function IsOperatorBool() As Boolean
    IsOperatorBool = IsSymbol("|") Or IsSymbol("~") Or IsSymbol("&") Or IsIdent("or") Or IsIdent("xor") Or IsIdent("and")
End Function

Function IsOperatorMul() As Boolean
    IsOperatorMul = IsSymbol("*") Or IsSymbol("/") Or IsSymbol("%") Or IsSymbol(">>") Or IsSymbol("<<") Or IsIdent("mul") Or IsIdent("div") Or IsIdent("mod") Or IsIdent("shr") Or IsIdent("shl")
End Function

Function IsOperatorRelation() As Boolean
    IsOperatorRelation = IsSymbol("=") Or IsSymbol("!=") Or IsSymbol("<>") Or IsSymbol(">=") Or IsSymbol("<=") Or IsSymbol("=>") Or IsSymbol("=<") Or IsSymbol(">") Or IsSymbol("<")
End Function

Sub Expression(Optional AssignTo As String)
    SkipBlank
    Assignment = AssignTo
    EvaluateCount = 0
    IsFloat = False
    Call EvalRelation
    If Not AssignTo = "" Then
        If AssignTo = "$Intern.Compare.One" And CompareOne <> "" Then Exit Sub
        If AssignTo = "$Intern.Compare.Two" And CompareTwo <> "" Then Exit Sub
        If GetSymbolSize(AssignTo) = 1 Then
            PopEDX
            AddCodeWord &H1588  'mov [variable],dl
            AddFixup AssignTo, OffsetOf(".code"), Code, &H400000
            AddCodeDWord 0
        ElseIf GetSymbolSize(AssignTo) = 2 Then
            PopEDX
            AddCodeWord &H8966  'mov [variable],cx
            AddCodeByte &H15
            AddFixup AssignTo, OffsetOf(".code"), Code, &H400000
            AddCodeDWord 0
        Else
            PopEAX
            AssignEAX AssignTo
        End If
    End If
End Sub

Sub EvalRelation()
    SkipBlank
    Call EvalBool
    While IsOperatorRelation
        If IsSymbol("<>") Then
            Relation = "<>": Position = Position + 2: Exit Sub
        ElseIf IsSymbol(">=") Then
            Relation = ">=": Position = Position + 2: Exit Sub
        ElseIf IsSymbol("<=") Then
            Relation = "<=": Position = Position + 2: Exit Sub
        ElseIf IsSymbol("=>") Then
            Relation = "=>": Position = Position + 2: Exit Sub
        ElseIf IsSymbol("=<") Then
            Relation = "=<": Position = Position + 2: Exit Sub
        ElseIf IsSymbol("=") Then
            Relation = "=": Position = Position + 1: Exit Sub
        ElseIf IsSymbol(">") Then
            Relation = ">": Position = Position + 1: Exit Sub
        ElseIf IsSymbol("<") Then
            Relation = "<": Position = Position + 1: Exit Sub
        ElseIf IsSymbol("!=") Then
            Relation = "!=": Position = Position + 2: Exit Sub
        End If
    Wend
End Sub

Sub EvalBool()
    SkipBlank
    Call EvalExpression
    While IsOperatorBool
        If IsSymbol("|") Or IsIdent("or") Then
            Call SkipIdent: Position = Position + 1: EvalExpression: ExprOr
        ElseIf IsSymbol("~") Or IsIdent("xor") Then
            Call SkipIdent: Position = Position + 1: EvalExpression: ExprXor
        ElseIf IsSymbol("&") Or IsIdent("and") Then
            Call SkipIdent: Position = Position + 1: EvalExpression: ExprAnd
        End If
        PushEAX
    Wend
End Sub

Sub EvalExpression()
    SkipBlank
    Call EvalTerm
    While IsOperatorAdd
        If IsSymbol("+") Or IsIdent("add") Then
            Call SkipIdent: Position = Position + 1: EvalTerm
            If IsFloat Then ExprFloatAdd Else ExprAdd
        ElseIf IsSymbol("-") Or IsIdent("sub") Then
            Call SkipIdent: Position = Position + 1: EvalTerm
            If IsFloat Then ExprFloatSub Else ExprSub
        End If
        PushEAX
    Wend
End Sub

Sub EvalTerm()
    SkipBlank
    Call EvalFactor
    While IsOperatorMul
        If IsSymbol("*") Or IsIdent("mul") Then
            Call SkipIdent: Position = Position + 1: EvalFactor
            If IsFloat Then ExprFloatMul Else ExprMul
        ElseIf IsSymbol("/") Or IsIdent("div") Then
            Call SkipIdent: Position = Position + 1: EvalFactor
            If IsFloat Then ExprFloatDiv Else ExprDiv
        ElseIf IsSymbol("%") Or IsIdent("mod") Then
            Call SkipIdent: Position = Position + 1: EvalFactor
            If IsFloat Then ExprFloatMod Else ExprMod
        ElseIf IsSymbol("<<") Then
            Call SkipIdent: Position = Position + 2: EvalFactor: ExprShl
        ElseIf IsIdent("shl") Then
            Call SkipIdent: Position = Position + 1: EvalFactor: ExprShl
        ElseIf IsSymbol(">>") Then
            Call SkipIdent: Position = Position + 2: EvalFactor: ExprShr
        ElseIf IsIdent("shr") Then
            Call SkipIdent: Position = Position + 1: EvalFactor: ExprShr
        End If
        PushEAX
    Wend
End Sub

Sub EvalFactor()
    Dim IsNot As Boolean: Dim IsPtr As Boolean: Dim Ident As String: Dim myAssign As String
    IsPtr = IsSymbol("^"): IsNot = IsSymbol("!")
    If IsPtr Or IsNot Then Position = Position + 1
    myAssign = Assignment: SkipBlank
    
    If EvaluateCount > 0 Then If CompareOne <> "" Then PushContent CompareOne: CompareOne = "": If CompareTwo <> "" Then PushContent CompareTwo: CompareTwo = ""
    
    IsFloat = False: IsStringCompare = False
    
    If IsFloatExpression Then
        PushF NumberExpression: IsFloat = True
    ElseIf IsNumberExpression Then
        Push NumberExpression
    ElseIf IsStringExpression Then
        IsStringCompare = True: dUniqueID = dUniqueID + 1
        Dim StringValue As String: StringValue = StringExpression
        DeclareDataString "$UniqueString" & dUniqueID, StringValue, Len(StringValue)
        
        If IsCallFrame Then GoTo GetAddressOfString
        
        If GetSymbolType(myAssign) = ST_DWORD Or GetSymbolType(myAssign) = ST_WORD Or GetSymbolType(myAssign) = ST_BYTE Or GetSymbolType(myAssign) = ST_SINGLE Or _
           GetSymbolType(myAssign) = ST_US_DWORD Or GetSymbolType(myAssign) = ST_US_WORD Or GetSymbolType(myAssign) = ST_US_BYTE Then
            PushAddress "$UniqueString" & dUniqueID
        ElseIf GetSymbolType(myAssign) = ST_STRING Then
GetAddressOfString:
            If IsCallFrame Then
                PushAddress "$UniqueString" & dUniqueID
            Else
                PushAddress "$UniqueString" & dUniqueID
                PushAddress myAssign
                InvokeByName "lstrcpy"
                PushContent myAssign
            End If
        Else
            PushAddress "$UniqueString" & dUniqueID
        End If
    ElseIf IsSymbol("(") Then
        Symbol "("
        Expression
        Symbol ")"
    ElseIf IsSymbol(")") Then
        Exit Sub
    ElseIf IsSymbol(",") Then
        Exit Sub
    ElseIf IsSymbol(";") Then
        Exit Sub
    ElseIf IsSymbol("@") Then
        Symbol "@"
        Dim VarIdentII As String
        VarIdentII = Identifier
        If GetSymbolType(CurrentFrame & "." & VarIdentII) = ST_LOCAL_DWORD Or _
           GetSymbolType(CurrentFrame & "." & VarIdentII) = ST_LOCAL_SINGLE Then
           AddCodeByte &H55 'push ebp
           Push GetSymbolOffset(CurrentFrame & "." & VarIdentII)
           ExprAdd
           PushEAX
        ElseIf GetSymbolType(CurrentFrame & "." & VarIdentII) = ST_LOCAL_STRING Then
            AddCodeWord &H858D
            AddCodeDWord GetSymbolOffset(CurrentFrame & "." & VarIdentII)
            AddCodeWord &H8B
            PushEAX
        Else
            If GetSymbolType(VarIdentII) = ST_FRAME Then
                VarIdentII = VarIdentII & ".Address"
            End If
            If GetSymbolType(VarIdentII) = ST_STRING Then
                PushAddress VarIdentII
                'Call popeax: AddCodeWord &H8B: PushEAX
            Else
                PushAddress VarIdentII
            End If
        End If
    Else
        Ident = Identifier
        If IsImport(Ident) Then
            CallImport Ident, True
            PushEAX
        ElseIf IsVariable(Ident) Then
                If IsSymbol("(") Then
                    Symbol "("
                    GetArray Ident
                    Symbol ")"
                ElseIf myAssign = "$Intern.Compare.One" And GetSymbolSize(Ident) = 4 And EvaluateCount = 0 Then
                    CompareOne = Ident
                ElseIf myAssign = "$Intern.Compare.Two" And GetSymbolSize(Ident) = 4 And EvaluateCount = 0 Then
                    CompareTwo = Ident
                Else
                    If GetSymbolType(Ident) = ST_STRING And GetSymbolType(myAssign) = ST_STRING Then
                        IsStringCompare = True
                        PushContent Ident
                    ElseIf GetSymbolType(Ident) = ST_STRING Then
                        IsStringCompare = True
                        PushAddress Ident
                    Else
                        If GetSymbolType(Ident) = ST_BYTE Then
                            AddCodeByte &HF
                            AddCodeByte &HBE
                            AddCodeByte &H5
                        ElseIf GetSymbolType(Ident) = ST_US_BYTE Then
                            AddCodeByte &HF
                            AddCodeByte &HB6
                            AddCodeByte &H5
                        ElseIf GetSymbolType(Ident) = ST_WORD Then
                            AddCodeByte &HF
                            AddCodeByte &HBF
                            AddCodeByte &H5
                        ElseIf GetSymbolType(Ident) = ST_US_WORD Then
                            AddCodeByte &HF
                            AddCodeByte &HB7
                            AddCodeByte &H5
                        ElseIf GetSymbolType(Ident) = ST_DWORD Then
                            AddCodeByte &HA1
                        ElseIf GetSymbolType(Ident) = ST_US_DWORD Then
                            AddCodeByte &HA1
                        ElseIf GetSymbolType(Ident) = ST_SINGLE Then
                            IsFloat = True
                            AddCodeByte &HA1
                        End If
                        AddFixup Ident, OffsetOf(".code"), Code, &H400000
                        AddCodeDWord 0
                        PushEAX
WasFloat:
                    End If
            End If
        ElseIf IsLocalVariable(Ident) Then
            If GetSymbolType(CurrentFrame & "." & Ident) = ST_LOCAL_STRING And myAssign <> "" Then
                AddCodeWord &H858D
                AddCodeDWord GetSymbolOffset(CurrentFrame & "." & Ident)
                PushEAX
            If GetSymbolType(myAssign) = ST_STRING Then AddCodeWord &H8B
                PushEAX
                PushAddress myAssign
                InvokeByName "lstrcpy"
                PushContent myAssign
            ElseIf GetSymbolType(CurrentFrame & "." & Ident) = ST_LOCAL_STRING Then
                
                AddCodeWord &H858D
                AddCodeDWord GetSymbolOffset(CurrentFrame & "." & Ident)
                AddCodeWord &H8B
                PushEAX
            ElseIf GetSymbolType(CurrentFrame & "." & Ident) = ST_LOCAL_DWORD Then
                'mov eax,[ebp+number]
                AddCodeWord &H858D
                AddCodeDWord GetSymbolOffset(CurrentFrame & "." & Ident)
                'mov eax,[eax]
                AddCodeWord &H8B
                PushEAX
            ElseIf GetSymbolType(CurrentFrame & "." & Ident) = ST_LOCAL_SINGLE Then
                IsFloat = True
                AddCodeWord &H858D
                AddCodeDWord GetSymbolOffset(CurrentFrame & "." & Ident)
                'mov eax,[eax]
                AddCodeWord &H8B
                PushEAX
            End If
        ElseIf IsAssignedType(Ident) Then
            PushAddress Ident
        ElseIf IsProperty(Ident & ".get") Then
            CallProperty Ident & ".get", True
            PushEAX
            IsStringCompare = False
        ElseIf IsFrame(Ident) Then
            CallFrame Ident, True
            Select Case GetReturnType(Ident)
                Case "single": IsFloat = True
                Case "string": IsStringCompare = True
                Case "dword": IsStringCompare = False: IsFloat = False
                Case "property": IsStringCompare = False: IsFloat = False
                Case Else
            End Select
            If GetReturnType(Ident) = "single" Then IsFloat = True
            If GetSymbolType(myAssign) = ST_STRING Then
                PushEAX
                PushAddress myAssign
                InvokeByName "lstrcpy"
                PushContent myAssign
            Else
                PushEAX
                IsStringCompare = False
            End If
        ElseIf Ident <> "" Then
            Position = Position - Len(Ident)
            CodeBlock
        Else
            If Ident = "" Then
                ErrMessage "unknown symbol '" & Mid$(Source, Position, 1) & "'": Exit Sub
            Else
                ErrMessage "unknown identifier '" & Ident & "'": Exit Sub
            End If
        End If
    End If
    While IsSymbol(" ")
        Position = Position + 1
    Wend
    If IsPtr Then PopEAX: AddCodeWord &H8B: PushEAX 'mov eax,[eax]
    If IsNot Then ExprNot: PushEAX
    EvaluateCount = EvaluateCount + 1
End Sub

Sub CallImport(Ident As String, Optional FromExpression As Boolean)
    Dim pCount As Integer
    pCount = ImportPCountByName(Ident)
    If pCount = -1 Then pCount = UserDefinedParameters
    ReverseParams
    Symbol "("
    While pCount > 0
        SkipBlank
        Expression
        If pCount > 1 Then Symbol ","
        pCount = pCount - 1
    Wend
    
    Symbol ")"
    If Not FromExpression Then Terminator
    InvokeByName Ident
    If Not FromExpression Then CodeBlock
End Sub

Function ReverseParams() As String
    Dim Header As String
    Dim Footer As String
    Dim Content As String
    Dim OPosition As Long
    
    OPosition = Position
    BracketsOpen = 0
    If IsSymbol("(") Then Position = Position + 1
    Content = RevParams
    Header = Mid$(Source, 1, OPosition)
    Footer = Mid$(Source, Position, Len(Source) - Position + 1)
    Source = Header & Content & Footer
    Position = OPosition
End Function

Function ParamsBrackets() As String
    While Not IsSymbol(")")
        If IsSymbol("(") Then Position = Position + 1: ParamsBrackets = ParamsBrackets & "(" & ParamsBrackets()
        ParamsBrackets = ParamsBrackets & Mid$(Source, Position, 1)
        Position = Position + 1
        If Position >= Len(Source) Then ErrMessage "found end of code. but expected ')' or ','": Exit Function
    Wend
End Function

Function RevParams() As String
    Dim i As Integer
    Dim Params() As String
    Dim strExpr As String
    
    ReDim Params(0) As String
    
    While Not IsSymbol(")")
        If IsSymbol("(") Then
            Position = Position + 1
            strExpr = strExpr & "("
            strExpr = strExpr & ParamsBrackets
        End If
        If IsSymbol(Chr(34)) Then
            strExpr = strExpr & Mid$(Source, Position, 1): Position = Position + 1
            While Not Mid$(Source, Position, 1) = Chr(34)
                strExpr = strExpr & Mid$(Source, Position, 1)
                Position = Position + 1
                If Position >= Len(Source) Then ErrMessage "found end of code. but expected ')' or ','": Exit Function
            Wend
        End If
        If IsSymbol(",") Then
            If Mid$(strExpr, 1, 1) = "," Then strExpr = Mid$(strExpr, 2, Len(strExpr))
            ReDim Preserve Params(UBound(Params) + 1) As String
            Params(UBound(Params)) = strExpr
            strExpr = ""
        End If
        strExpr = strExpr & Mid$(Source, Position, 1)
        Position = Position + 1
        If Position >= Len(Source) Then ErrMessage "found end of code. but expected ')' or ','": Exit Function
    Wend
Done:
    If Mid$(strExpr, 1, 1) = "," Then strExpr = Mid$(strExpr, 2, Len(strExpr))
    ReDim Preserve Params(UBound(Params) + 1) As String
    Params(UBound(Params)) = strExpr
    
    Dim strReversed As String
    Dim strOriginal As String
   
    For i = UBound(Params) To 1 Step -1
        strReversed = strReversed & Params(i) & Switch(i > 1, ",")
    Next i
    
    RevParams = strReversed
End Function

Function UserDefinedParameters() As Long
    Dim i As Long: Dim InStringExpr As Boolean

    i = Position
    InStringExpr = False
    UserDefinedParameters = 1
    While Mid$(Source, i, 1) <> ")"
        If Mid$(Source, i, 1) = Chr(34) Then
            If InStringExpr = False Then
                InStringExpr = True
            Else
                InStringExpr = False
            End If
            i = i + 1
        End If
        If Mid$(Source, i, 1) = "," Then
            If InStringExpr = False Then
                UserDefinedParameters = UserDefinedParameters + 1
            End If
        End If
        i = i + 1
        If Position >= Len(Source) Then ErrMessage "found end of code. but expected ')' or ','": Exit Function
    Wend
End Function

