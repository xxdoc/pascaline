Attribute VB_Name = "comAssembler"
Option Explicit

Sub PopEAX()
    If Not OptimizeAble("PopEAX") Then
        AddCodeByte &H58
    End If
End Sub

Sub PopEBX()
    AddCodeByte &H5B
End Sub

Sub PopECX()
    AddCodeByte &H59
End Sub

Sub PopEDX()
    AddCodeByte &H5A
End Sub

Sub PushEAX()
    AddCodeByte &H50
End Sub

Sub PushECX()
    AddCodeByte &H51
End Sub

Sub DecECX()
    AddCodeByte &H49
End Sub

Sub IncECX()
    AddCodeByte &H41
End Sub

Sub StoreECX()
    'mov [variable],ecx
    AddCodeWord &HD89
    AddCodeFixup "$Intern.Loop"
End Sub

Sub RestoreECX()
    'mov ecx,[variable]
    AddCodeWord &HD8B
    AddCodeFixup "$Intern.Loop"
End Sub

Sub Push(Optional Value As Long)
    'push dword Value
    AddCodeByte &H68
    AddCodeDWord Value
End Sub

Sub PushF(Optional Value As Single)
    'push dword Value
    AddCodeByte &H68
    AddCodeSingle Value
End Sub

Sub PushContent(Variable As String)
    'push [Variable]
    AddCodeWord &H35FF
    AddFixup Variable, OffsetOf(".code"), Code, &H400000
    AddRelocation OffsetOf(".code")
    AddCodeDWord &H0
End Sub

Sub PushAddress(Variable As String)
    'push Variable
    AddCodeByte &H68
    AddFixup Variable, OffsetOf(".code"), Code, &H400000
    AddRelocation OffsetOf(".code")
    AddCodeDWord &H0
End Sub

Sub PushFloatEAX()
    fUniqueID = fUniqueID + 1
    DeclareDataSingle "$Float" & fUniqueID, 0
    AssignEAX "$Float" & fUniqueID
    PushFloatContent "$Float" & fUniqueID
End Sub

Sub PushFloatEDX()
    fUniqueID = fUniqueID + 1
    DeclareDataSingle "$Float" & fUniqueID, 0
    AssignEDX "$Float" & fUniqueID
    PushFloatContent "$Float" & fUniqueID
End Sub

Sub PushFloat(Value As Single)
    fUniqueID = fUniqueID + 1
    DeclareDataSingle "$Float" & fUniqueID, Value
    PushFloatContent "$Float" & fUniqueID
End Sub

Sub PushFloatContent(Name As String)
    AddCodeWord &H5D9
    AddCodeFixup Name
End Sub

Sub Invoke()
    AddCodeWord &H15FF
End Sub

Sub InvokeByName(Name As String)
    Call Invoke
    SetImportUsed Name, OffsetOf(".code")
    AddFixup Name, OffsetOf(".code"), Code, &H400000
    AddCodeDWord &H0
End Sub

Sub ExprAdd()
    Call PopEDX: PopEAX
    'add eax,edx
    AddCodeWord &HD001
End Sub

Sub ExprFloatAdd()
    Call PopEDX: PopEAX
    PushFloatEAX
    PushFloatEDX
    'faddp
    AddCodeWord &HC1DE
    'fstp variable
    AddCodeWord &H1DD9
    AddCodeFixup "$Intern.Float"
    MovEAX "$Intern.Float"
End Sub

Sub ExprFloatSub()
    Call PopEDX: PopEAX
    PushFloatEAX
    PushFloatEDX
    'fsubp
    AddCodeWord &HE9DE
    'fstp variable
    AddCodeWord &H1DD9
    AddCodeFixup "$Intern.Float"
    MovEAX "$Intern.Float"
End Sub

Sub ExprFloatMul()
    Call PopEDX: PopEAX
    PushFloatEAX
    PushFloatEDX
    'fmulp
    AddCodeWord &HC9DE
    'fstp variable
    AddCodeWord &H1DD9
    AddCodeFixup "$Intern.Float"
    MovEAX "$Intern.Float"
End Sub

Sub ExprFloatDiv()
    Call PopEDX: PopEAX
    PushFloatEAX
    PushFloatEDX
    'fdivp
    AddCodeWord &HF9DE
    'fstp variable
    AddCodeWord &H1DD9
    AddCodeFixup "$Intern.Float"
    MovEAX "$Intern.Float"
End Sub

Sub ExprFloatMod()
    Call PopEDX: PopEAX
    PushFloatEAX
    PushFloatEDX
    'fprem
    AddCodeWord &HF8D9
    'fstp variable
    AddCodeWord &H1DD9
    AddCodeFixup "$Intern.Float"
    MovEAX "$Intern.Float"
End Sub

Sub ExprSub()
    Call PopEDX: PopEAX
    'sub eax,edx
    AddCodeWord &HD029
End Sub

Sub ExprDiv()
    Call PopEDX: PopEAX
    'mov ebx,edx
    AddCodeWord &HD389
    'mov edx,0
    AddCodeByte &HBA
    AddCodeDWord &H0
    'idiv ebx
    AddCodeWord &HFBF7
End Sub

Sub ExprMul()
    Call PopEDX: PopEAX
    'mov ebx,edx
    AddCodeWord &HD389
    'mul ebx
    AddCodeWord &HE3F7
End Sub

Sub ExprMod()
    Call PopEDX: PopEAX
    'mov ebx,edx
    AddCodeWord &HD389
    'mov edx,0
    AddCodeByte &HBA
    AddCodeDWord &H0
    'idiv ebx
    AddCodeWord &HFBF7
    'mov eax,edx
    AddCodeWord &HC28B
End Sub

Sub ExprShl()
    Call PopECX: PopEAX
    'shl eax,cl
    AddCodeWord &HE0D3
End Sub

Sub ExprShr()
    Call PopECX: PopEAX
    'shl eax,cl
    AddCodeWord &HE8D3
End Sub

Sub ExprAnd()
    Call PopEBX: PopEAX
    'and eax,ebx
    AddCodeWord &HC323
End Sub

Sub ExprOr()
    Call PopEBX: PopEAX
    'or eax,ebx
    AddCodeWord &HC30B
End Sub

Sub ExprXor()
    Call PopEBX: PopEAX
    'xor eax,ebx
    AddCodeWord &HC333
End Sub

Sub ExprNeg()
    AddCodeWord &HD8F7
End Sub

Sub ExprNot()
    AddCodeWord &HD0F7
End Sub

Sub MovEAX(Name As String)
    'mov eax,[name]
    AddCodeByte &HA1: AddCodeFixup Name
End Sub

Sub MovEAXAddress(Name As String)
    'mov eax,name
    AddCodeByte &HB8: AddCodeFixup Name
End Sub

Sub MovEDX(Name As String)
    'mov edx,[name]
    AddCodeWord &H158B: AddCodeFixup Name
End Sub

Sub ExprCompare(Variable As String, Variable2 As String)
    If CompareOne <> "" And CompareTwo = "" Then
        MovEAX CompareOne: MovEDX Variable2
    ElseIf CompareOne = "" And CompareTwo <> "" Then
        MovEAX Variable:  MovEDX CompareTwo
    ElseIf CompareOne <> "" And CompareTwo <> "" Then
        MovEAX CompareOne: MovEDX CompareTwo
    Else
        MovEAX Variable: MovEDX Variable2
    End If
    'cmp eax,edx
    AddCodeWord &HD039
    CompareOne = "": CompareTwo = ""
End Sub

Sub ExprCompareS(Variable As String, Variable2 As String)
    PushContent Variable2: PushContent Variable
    InvokeByName "lstrcmp"
    'cmp eax,0
    AddCodeByte &H3D: AddCodeDWord &H0
End Sub

Sub AssignEAX(Variable As String)
    'mov [variable],eax
    AddCodeByte &HA3
    AddFixup Variable, OffsetOf(".code"), Code, &H400000
    AddCodeDWord 0
End Sub

Sub AssignEDX(Variable As String)
    'mov [variable],edx
    AddCodeWord &H1589
    AddFixup Variable, OffsetOf(".code"), Code, &H400000
    AddCodeDWord 0
End Sub

Sub ExprJE(Name As String)
    AddCodeWord &H840F
    AddFixup Name, OffsetOf(".code"), Code
    AddCodeDWord &H0
End Sub

Sub ExprJNE(Name As String)
    AddCodeWord &H850F
    AddFixup Name, OffsetOf(".code"), Code
    AddCodeDWord &H0
End Sub

Sub ExprJL(Name As String)
    AddCodeWord &H8C0F
    AddFixup Name, OffsetOf(".code"), Code
    AddCodeDWord &H0
End Sub

Sub ExprJLE(Name As String)
    AddCodeByte &HF
    AddCodeByte &H8E
    AddFixup Name, OffsetOf(".code"), Code
    AddCodeDWord &H0
End Sub

Sub ExprJA(Name As String)
    AddCodeWord &H8F0F
    AddFixup Name, OffsetOf(".code"), Code
    AddCodeDWord &H0
End Sub

Sub ExprJAE(Name As String)
    AddCodeWord &H8D0F
    AddFixup Name, OffsetOf(".code"), Code
    AddCodeDWord &H0
End Sub

Sub ExprJump(Name As String)
    AddCodeByte &HE9
    AddFixup Name, OffsetOf(".code"), Code
    AddCodeDWord &H0
End Sub

Sub ExprLoop(Name As String)
    'loop label
    AddCodeByte &HE2
    AddCodeByte &HFF - CByte(OffsetOf(".code") - GetSymbolOffset(Name))
End Sub

Sub ExprCall(Name As String)
    AddCodeByte &HE8
    AddFixup Name, OffsetOf(".code"), Code
    AddCodeDWord &HFFFFFFFF
End Sub

Sub StartFrame()
    AddCodeByte &H55
    AddCodeByte &H89
    AddCodeByte &HE5
End Sub

Sub EndFrame(Value As Integer)
    AddCodeByte &HC9
    AddCodeByte &HC2
    AddCodeWord Value
End Sub

Sub InitializeDLL()
    StartFrame
    AddCodeByte &HB8
    AddCodeDWord &H1
    EndFrame &HC
End Sub

