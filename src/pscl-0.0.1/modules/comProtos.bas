Attribute VB_Name = "comProtos"
Option Explicit

Sub AssignProtoTypes()
    Dim OPosition As Long
    
    OPosition = Position
    
    While Position <= Len(Source)
        If Mid$(Source, Position, 5) = "frame" Then
            Call SkipIdent: SkipBlank
            DeclareFrame False, False, True
        ElseIf Mid$(Source, Position, 8) = "property" Then
            Call SkipIdent: SkipBlank
            DeclareFrame False, False, True, True
        ElseIf Mid$(Source, Position, 6) = "export" Then
            Call SkipIdent: SkipBlank
            DeclareFrame True, False, True
        ElseIf Mid$(Source, Position, 2) = "//" Then
            While Mid$(Source, Position, 2) <> vbCrLf
                Position = Position + 1
            Wend
        End If
        Position = Position + 1
        If Position >= Len(Source) Then GoTo ProtoDone
    Wend
    
ProtoDone:
    Position = OPosition
End Sub
