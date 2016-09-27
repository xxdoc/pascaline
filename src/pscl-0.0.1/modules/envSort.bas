Attribute VB_Name = "envSort"

Option Explicit
Option Base 0        'If comparing this code to other sorting code for speed,
Option Compare Text  '    be sure to test under equal Option Compare conditions


'===================================================================================
'                                SORTING MODULE
'
'  Author:  John Korejwa  <korejwa@tiac.net>
'  Version: 29/DEC/2002
'
'  Resubmitted to PSC on 8/November/2003
'
'  Description:
'    This module applies the QuickSort algorithm for sorting an array of values.
'
'    Quicksort is the fastest known general sorting algorithm for large arrays.
'    However, once the number of elements in a partitioned subarray is smaller than
'    some threshhold, Insertion Sort becomes faster.  So this code uses QuickSort
'    for large subarrays, and Insertion Sort for small subarrays.
'
'    Quicksort is generally implemented recursively, but this code is non-recursive.
'    To avoid recursion, a very simple Stack is used within the sorting procedure.
'
'    It is possible to get access to pointers to strings using the CopyMemory API
'    call and the undocumented VarPtr() function.  Accessing pointers speeds things
'    up enormously, particularly when swapping two strings.
'
'===================================================================================


Private Const QTHRESH As Long = 7          'Threshhold for switching from QuickSort to Insertion Sort
Private Const MinLong As Long = &H80000000

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Sub SwapStrings(String1 As String, String2 As String)
    Static SwpVal As Long
    CopyMemory SwpVal, ByVal VarPtr(String1), 4
    CopyMemory ByVal VarPtr(String1), ByVal VarPtr(String2), 4
    CopyMemory ByVal VarPtr(String2), SwpVal, 4
End Sub
Private Sub SwapLong(a As Long, b As Long)
    Static c As Long
    c = a
    a = b
    b = c
End Sub


Public Sub SortStringArray(TheArray() As String, Optional LowerBound As Long = MinLong, Optional UpperBound As Long = MinLong)
    Dim f          As Long   'Subarray Minimum
    Dim g          As Long   'Subarray Maximum
    Dim h          As Long   'Subarray Middle
    Dim i          As Long   'Subarray Low  Scan Index
    Dim j          As Long   'Subarray High Scan Index

    Dim s(1 To 64) As Long   'Stack space for pending Subarrays
    Dim t          As Long   'Stack pointer

    Dim swp        As String 'Swap variable

    If LowerBound = MinLong Then f = LBound(TheArray) Else f = LowerBound
    If UpperBound = MinLong Then g = UBound(TheArray) Else g = UpperBound

    t = 0
    Do
        If g - f < QTHRESH Then       'Insertion Sort this small subarray
            For j = f + 1 To g
                CopyMemory ByVal VarPtr(swp), ByVal VarPtr(TheArray(j)), 4 ' swp = TheArray(j)
                For i = j - 1 To f Step -1
                    If TheArray(i) <= swp Then Exit For
                    CopyMemory ByVal VarPtr(TheArray(i + 1)), ByVal VarPtr(TheArray(i)), 4 ' TheArray(i + 1) = TheArray(i)
                Next i
                CopyMemory ByVal VarPtr(TheArray(i + 1)), ByVal VarPtr(swp), 4 ' TheArray(i + 1) = swp
            Next j
            If t = 0 Then Exit Do     'Finished sorting <<<
            g = s(t)                  'Pop stack and begin new partitioning round
            f = s(t - 1)
            t = t - 2
        Else                          'QuickSort this large subarray
            h = (f + g) \ 2
            SwapStrings TheArray(h), TheArray(f + 1)
            If TheArray(f) > TheArray(g) Then SwapStrings TheArray(f), TheArray(g)
            If TheArray(f + 1) > TheArray(g) Then SwapStrings TheArray(f + 1), TheArray(g)
            If TheArray(f) > TheArray(f + 1) Then SwapStrings TheArray(f), TheArray(f + 1)

            i = f + 1                 'Initialize pointers for partitioning
            j = g                     'swp is partitioning element
            CopyMemory ByVal VarPtr(swp), ByVal VarPtr(TheArray(f + 1)), 4 ' swp = TheArray(f + 1)
            Do
                Do                    'Scan up to find element > swp
                  i = i + 1
                Loop While TheArray(i) < swp
                Do                    'Scan down to find element < swp
                    j = j - 1
                Loop While TheArray(j) > swp
                If j < i Then Exit Do 'Scan Elements crossed ... Partitioning complete
                SwapStrings TheArray(i), TheArray(j)
            Loop

            'Insert partitioning element
            CopyMemory ByVal VarPtr(TheArray(f + 1)), ByVal VarPtr(TheArray(j)), 4 ' TheArray(f + 1) = TheArray(j)
            CopyMemory ByVal VarPtr(TheArray(j)), ByVal VarPtr(swp), 4 ' TheArray(j) = swp

            t = t + 2 'Push larger subarray onto stack; Sort smaller subarray first
            If g - i + 1 >= j - f Then
                s(t) = g
                s(t - 1) = i
                g = j - 1
            Else
                s(t) = j - 1
                s(t - 1) = f
                f = i
            End If
        End If
    Loop

    CopyMemory ByVal VarPtr(swp), 0&, 4  'Clear the string pointer.  This is necessary,
                                         'especially if this code is run under Win NT 4.0
End Sub


Public Sub SortLongArray(TheArray() As Long, Optional LowerBound As Long = MinLong, Optional UpperBound As Long = MinLong)
    Dim f          As Long   'Subarray Minimum
    Dim g          As Long   'Subarray Maximum
    Dim h          As Long   'Subarray Middle
    Dim i          As Long   'Subarray Low  Scan Index
    Dim j          As Long   'Subarray High Scan Index

    Dim s(1 To 64) As Long   'Stack space for pending Subarrays
    Dim t          As Long   'Stack pointer

    Dim swp        As Long   'Swap variable

    If LowerBound = MinLong Then f = LBound(TheArray) Else f = LowerBound
    If UpperBound = MinLong Then g = UBound(TheArray) Else g = UpperBound

    t = 0
    Do
        If g - f < QTHRESH Then       'Insertion Sort this small subarray
            For j = f + 1 To g
                swp = TheArray(j)
                For i = j - 1 To f Step -1
                    If TheArray(i) <= swp Then Exit For
                    TheArray(i + 1) = TheArray(i)
                Next i
                TheArray(i + 1) = swp
            Next j
            If t = 0 Then Exit Do     'Finished sorting <<<
            g = s(t)                  'Pop stack and begin new partitioning round
            f = s(t - 1)
            t = t - 2
        Else                          'QuickSort this large subarray
            h = (f + g) \ 2
            SwapLong TheArray(h), TheArray(f + 1)
            If TheArray(f) > TheArray(g) Then SwapLong TheArray(f), TheArray(g)
            If TheArray(f + 1) > TheArray(g) Then SwapLong TheArray(f + 1), TheArray(g)
            If TheArray(f) > TheArray(f + 1) Then SwapLong TheArray(f), TheArray(f + 1)

            i = f + 1                 'Initialize pointers for partitioning
            j = g                     'swp is partitioning element
            swp = TheArray(f + 1)
            Do
                Do                    'Scan up to find element > swp
                  i = i + 1
                Loop While TheArray(i) < swp
                Do                    'Scan down to find element < swp
                    j = j - 1
                Loop While TheArray(j) > swp
                If j < i Then Exit Do 'Scan Elements crossed ... Partitioning complete
                SwapLong TheArray(i), TheArray(j)
            Loop

            TheArray(f + 1) = TheArray(j) 'Insert partitioning element
            TheArray(j) = swp

            t = t + 2 'Push larger subarray onto stack; Sort smaller subarray first
            If g - i + 1 >= j - f Then
                s(t) = g
                s(t - 1) = i
                g = j - 1
            Else
                s(t) = j - 1
                s(t - 1) = f
                f = i
            End If
        End If
    Loop

End Sub


Public Sub SortStringIndexArray(TheArray() As String, TheIndex() As Long, Optional LowerBound As Long = MinLong, Optional UpperBound As Long = MinLong)
    Dim f          As Long
    Dim g          As Long
    Dim h          As Long
    Dim i          As Long
    Dim j          As Long

    Dim s(1 To 64) As Long
    Dim t          As Long

    Dim swp        As String
    Dim indxt      As Long

    If LowerBound = MinLong Then f = LBound(TheIndex) Else f = LowerBound
    If UpperBound = MinLong Then g = UBound(TheIndex) Else g = UpperBound

    t = 0
    Do
        If g - f < QTHRESH Then
            For j = f + 1 To g
                indxt = TheIndex(j)
                CopyMemory ByVal VarPtr(swp), ByVal VarPtr(TheArray(indxt)), 4 ' swp = TheArray(indxt)
                For i = j - 1 To f Step -1
                    If TheArray(TheIndex(i)) <= swp Then Exit For
                    TheIndex(i + 1) = TheIndex(i)
                Next i
                TheIndex(i + 1) = indxt
            Next j
            If t = 0 Then Exit Do
            g = s(t)
            f = s(t - 1)
            t = t - 2
        Else
            h = (f + g) \ 2
            SwapLong TheIndex(h), TheIndex(f + 1)

            If TheArray(TheIndex(f)) > TheArray(TheIndex(g)) Then SwapLong TheIndex(f), TheIndex(g)
            If TheArray(TheIndex(f + 1)) > TheArray(TheIndex(g)) Then SwapLong TheIndex(f + 1), TheIndex(g)
            If TheArray(TheIndex(f)) > TheArray(TheIndex(f + 1)) Then SwapLong TheIndex(f), TheIndex(f + 1)

            i = f + 1
            j = g
            indxt = TheIndex(f + 1)
            CopyMemory ByVal VarPtr(swp), ByVal VarPtr(TheArray(indxt)), 4 ' swp = TheArray(indxt)
            Do
                Do
                  i = i + 1
                Loop While TheArray(TheIndex(i)) < swp
                Do
                    j = j - 1
                Loop While TheArray(TheIndex(j)) > swp
                If j < i Then Exit Do
                SwapLong TheIndex(i), TheIndex(j)
            Loop

            TheIndex(f + 1) = TheIndex(j)
            TheIndex(j) = indxt

            t = t + 2
            If g - i + 1 >= j - f Then
                s(t) = g
                s(t - 1) = i
                g = j - 1
            Else
                s(t) = j - 1
                s(t - 1) = f
                f = i
            End If
        End If
    Loop

    CopyMemory ByVal VarPtr(swp), 0&, 4  'Clear the string pointer.  This is necessary.
                                         'especially if this code is run under Win NT 4.0
End Sub


Public Sub SortLongIndexArray(TheArray() As Long, TheIndex() As Long, Optional LowerBound As Long = MinLong, Optional UpperBound As Long = MinLong)
    Dim f          As Long
    Dim g          As Long
    Dim h          As Long
    Dim i          As Long
    Dim j          As Long

    Dim s(1 To 64) As Long
    Dim t          As Long

    Dim swp        As Long
    Dim indxt      As Long

    If LowerBound = MinLong Then f = LBound(TheIndex) Else f = LowerBound
    If UpperBound = MinLong Then g = UBound(TheIndex) Else g = UpperBound

    t = 0
    Do
        If g - f < QTHRESH Then
            For j = f + 1 To g
                indxt = TheIndex(j)
                swp = TheArray(indxt)
                For i = j - 1 To f Step -1
                    If TheArray(TheIndex(i)) <= swp Then Exit For
                    TheIndex(i + 1) = TheIndex(i)
                Next i
                TheIndex(i + 1) = indxt
            Next j
            If t = 0 Then Exit Do
            g = s(t)
            f = s(t - 1)
            t = t - 2
        Else
            h = (f + g) \ 2
            SwapLong TheIndex(h), TheIndex(f + 1)

            If TheArray(TheIndex(f)) > TheArray(TheIndex(g)) Then SwapLong TheIndex(f), TheIndex(g)
            If TheArray(TheIndex(f + 1)) > TheArray(TheIndex(g)) Then SwapLong TheIndex(f + 1), TheIndex(g)
            If TheArray(TheIndex(f)) > TheArray(TheIndex(f + 1)) Then SwapLong TheIndex(f), TheIndex(f + 1)

            i = f + 1
            j = g
            indxt = TheIndex(f + 1)
            swp = TheArray(indxt)
            Do
                Do
                  i = i + 1
                Loop While TheArray(TheIndex(i)) < swp
                Do
                    j = j - 1
                Loop While TheArray(TheIndex(j)) > swp
                If j < i Then Exit Do
                SwapLong TheIndex(i), TheIndex(j)
            Loop

            TheIndex(f + 1) = TheIndex(j)
            TheIndex(j) = indxt

            t = t + 2
            If g - i + 1 >= j - f Then
                s(t) = g
                s(t - 1) = i
                g = j - 1
            Else
                s(t) = j - 1
                s(t - 1) = f
                f = i
            End If
        End If
    Loop

End Sub

