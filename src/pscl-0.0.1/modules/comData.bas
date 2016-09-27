Attribute VB_Name = "comData"
Option Explicit

Public lUniqueID As Long
Public sUniqueID As Long
Public dUniqueID As Long
Public fUniqueID As Long

Sub InitData()
    lUniqueID = 0
    sUniqueID = 0
    dUniqueID = 0
    fUniqueID = 0
End Sub

Sub DeclareDataSingle(Name As String, Value As Single)
    AddSymbol Name, OffsetOf(".data"), Data, ST_SINGLE
    AddDataSingle Value
End Sub

Sub DeclareDataDWord(Name As String, Value As Long)
    AddSymbol Name, OffsetOf(".data"), Data, ST_DWORD
    AddDataDWord Value
End Sub

Sub DeclareDataWord(Name As String, Value As Integer)
    AddSymbol Name, OffsetOf(".data"), Data, ST_WORD
    AddDataWord Value
End Sub

Sub DeclareDataByte(Name As String, Value As Byte)
    AddSymbol Name, OffsetOf(".data"), Data, ST_BYTE
    AddDataByte Value
End Sub

Sub DeclareDataUnsignedDWord(Name As String, Value As Long)
    AddSymbol Name, OffsetOf(".data"), Data, ST_US_DWORD
    AddDataDWord Value
End Sub

Sub DeclareDataUnsignedWord(Name As String, Value As Integer)
    AddSymbol Name, OffsetOf(".data"), Data, ST_US_WORD
    AddDataWord Value
End Sub

Sub DeclareDataUnsignedByte(Name As String, Value As Byte)
    AddSymbol Name, OffsetOf(".data"), Data, ST_US_BYTE
    AddDataByte Value
End Sub

Sub DeclareDataString(Name As String, Text As String, Optional Space As Long)
    Dim i As Integer
    AddSymbol Name, OffsetOf(".data"), Data, ST_STRING
    
    For i = 1 To Len(Text): AddDataByte Asc(Mid$(Text, i, 1)): Next i
    
    If Space > 0 Then
        For i = Len(Text) To Space
            AddDataByte &H0
        Next i
    End If
    
    AddDataByte &H0
End Sub

