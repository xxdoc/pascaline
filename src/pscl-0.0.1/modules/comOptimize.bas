Attribute VB_Name = "comOptimize"
Option Explicit

Function OptimizeAble(ByInstruction As String, Optional Variable As String) As Boolean
    Dim B1 As Byte: Dim B2 As Byte: Dim B3 As Byte: Dim B4 As Byte
    
    On Error GoTo NotOptimize
    
    Select Case ByInstruction
        Case "PopEAX"
            If Section(2).Bytes(UBound(Section(2).Bytes)) = &H50 Then
                'Kill PushEAX and do not add PopEAX
                ReDim Preserve Section(2).Bytes(UBound(Section(2).Bytes) - 1) As Byte
                OptimizeAble = True
            ElseIf Section(2).Bytes(UBound(Section(2).Bytes) - 4) = &H68 Then
                'If a Push Number is made and then pop eax this can be -> mov eax,Number
                B1 = Section(2).Bytes(UBound(Section(2).Bytes) - 3)
                B2 = Section(2).Bytes(UBound(Section(2).Bytes) - 2)
                B3 = Section(2).Bytes(UBound(Section(2).Bytes) - 1)
                B4 = Section(2).Bytes(UBound(Section(2).Bytes))
                ReDim Preserve Section(2).Bytes(UBound(Section(2).Bytes) - 5) As Byte
                'mov eax,Value
                AddCodeByte &HB8
                AddCodeByte B1
                AddCodeByte B2
                AddCodeByte B3
                AddCodeByte B4
                OptimizeAble = True
            End If
        Case Else
    End Select
NotOptimize:
End Function
