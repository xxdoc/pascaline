Attribute VB_Name = "comMath"
Option Explicit

Type TYPE_QWORD
    Value As Currency
End Type

Private Type TYPE_LOHIQWORD
    lLoDWord As Long
    lHiDWord As Long
End Type

Function HiByte(ByVal iWord As Integer) As Byte
    HiByte = (iWord And &HFF00&) \ &H100
End Function

Function LoByte(ByVal iWord As Integer) As Byte
    LoByte = iWord And &HFF
End Function

Function HiWord(lDword As Long) As Integer
    HiWord = (lDword And &HFFFF0000) \ &H10000
End Function

Function LoWord(lDword As Long) As Integer
    If lDword And &H8000& Then
        LoWord = lDword Or &HFFFF0000
    Else
        LoWord = lDword And &HFFFF&
    End If
End Function

Function HiSWord(lDword As Single) As Integer
    HiSWord = (lDword And &HFFFF0000) \ &H10000
End Function

Function LoSWord(lDword As Single) As Integer
    If lDword And &H8000& Then
        LoSWord = lDword Or &HFFFF0000
    Else
        LoSWord = lDword And &HFFFF&
    End If
End Function

Function LoDWord(ByVal cQWord As Currency) As Long
    Dim QWord As TYPE_QWORD: Dim LoHiQword As TYPE_LOHIQWORD
    QWord.Value = cQWord / 10000
    LSet LoHiQword = QWord
    LoDWord = LoHiQword.lLoDWord
End Function

Function HiDWord(ByVal cQWord As Currency) As Long
    Dim QWord As TYPE_QWORD: Dim LoHiQword As TYPE_LOHIQWORD
    QWord.Value = cQWord / 10000
    LSet LoHiQword = QWord
    HiDWord = LoHiQword.lHiDWord
End Function

Function MakeQWord(ByVal lHiDWord As Long, ByVal lLoDWord As Long) As Currency
    Dim QWord As TYPE_QWORD: Dim LoHiQword As TYPE_LOHIQWORD
    LoHiQword.lHiDWord = lHiDWord: LoHiQword.lLoDWord = lLoDWord
    LSet QWord = LoHiQword: MakeQWord = QWord.Value * 10000
End Function

