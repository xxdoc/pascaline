Attribute VB_Name = "comMemory"
Option Explicit

Sub ReserveArray(Ident As String, Size As Long)
    DeclareDataDWord Ident & ".PtrToArray", 0
    DeclareDataDWord Ident & ".HeapHandle", 0
    DeclareDataDWord Ident & ".ubound", 0
    DeclareDataDWord Ident & ".lbound", 0
    DeclareDataDWord Ident & ".Array", 0
    
    If CurrentFrame <> "" Then
        Push 0: Push 0: Push 0
        InvokeByName "HeapCreate"                           'Create Heap
        AssignEAX Ident & ".HeapHandle"                     'Save Handle of HeapCreate
    
        Push GetSymbolSize(Ident) * Size: Push 8            'Size in Bytes: 8 = ZeroMemory
        PushContent Ident & ".HeapHandle"                   'Push Handle of HeapCreate
        InvokeByName "HeapAlloc"                            'Allocate Heap
        AssignEAX Ident & ".PtrToArray"                     'Save Individual Ptr to Allocated Memory for Ident variable
        
        Push 0: PopEAX: AssignEAX Ident & ".lbound"
        Push Size: PopEAX: AssignEAX Ident & ".ubound"
    End If
    
End Sub

Sub GetArray(Ident As String)
    Expression Ident & ".Array"
    Push GetSymbolSize(Ident)
    PushContent Ident & ".Array"
    Push GetSymbolSize(Ident)
    ExprMul
    PushEAX
    PushContent Ident & ".PtrToArray"
    ExprAdd
    PushEAX
    PushAddress Ident
    
    InvokeByName "MoveMemory": PushContent Ident
End Sub

Sub SetArray(Name As String)
    Symbol "("
    Expression "$Intern.Array"
    Symbol ")"
    Symbol "="
    Expression Name
    
    Push GetSymbolSize(Name)
    PushAddress Name
    PushContent "$Intern.Array"
    Push GetSymbolSize(Name)
    ExprMul
    PushEAX
    PushContent Name & ".PtrToArray"
    ExprAdd
    PushEAX
    
    InvokeByName "MoveMemory"
End Sub

Sub StatementPreserve()
    Dim Ident As String
    Ident = Identifier
    Symbol "("
    Expression "$Intern.Array"
    Symbol ")"
    Terminator
    
    PushContent "$Intern.Array": Push 1
    PushContent Ident & ".HeapHandle"
    InvokeByName "HeapAlloc"
    
    PushContent "$Intern.Array": PopEAX: AssignEAX Ident & ".ubound"
    
    CodeBlock
End Sub

Sub StatementReserve()
    Dim Ident As String
    Ident = Identifier
    Symbol "("
    Expression "$Intern.Array"
    Symbol ")"
    Terminator
    
    Push 0: Push 0: Push 0
    InvokeByName "HeapCreate"                           'Create Heap
    AssignEAX Ident & ".HeapHandle"                     'Save Handle of HeapCreate
    
    PushContent "$Intern.Array"
    Push GetSymbolSize(Ident)
    ExprMul
    PushEAX
    Push 8                                              '8 = ZeroMemory
    PushContent Ident & ".HeapHandle"                   'Push Handle of HeapCreate
    InvokeByName "HeapAlloc"                            'Allocate Heap
    AssignEAX Ident & ".PtrToArray"                     'Save Individual Ptr to Allocated Memory for Ident variable
    
    PushContent "$Intern.Array": PopEAX: AssignEAX Ident & ".ubound"
    
    CodeBlock
End Sub

Sub StatementDestroy()
    Dim Ident As String
    Ident = Identifier
    Terminator
    If SymbolExists(Ident & ".HeapHandle") Then
        PushContent Ident & ".HeapHandle"
        InvokeByName "HeapDestroy"                          'Destroy Heap
        Push 0: PopEAX
        AssignEAX Ident & ".ubound"
    Else
        ErrMessage "cannot destroy '" & Ident & "' has not been reserved before.": Exit Sub
    End If
    CodeBlock
End Sub

Sub StatementUBound()
    Dim Ident As String
    Symbol "("
    Ident = Identifier
    Symbol ")"
    PushContent Ident & ".ubound"
End Sub

Sub StatementLBound()
    Dim Ident As String
    Symbol "("
    Ident = Identifier
    Symbol ")"
    PushContent Ident & ".lbound"
End Sub

