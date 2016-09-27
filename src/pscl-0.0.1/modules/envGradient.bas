Attribute VB_Name = "envGradient"

'-------------------------------------------------------------------------
' Source Code : 'SmoothGradient'
' Upgraded    : Jim Jose
' email       : jimjosev33@yahoo.com
' Credits     : Based on an Excellent thought by 'paul_turcksin' to use
'               CopyMemory to bypass Color Bytes to TRIVERTEX Structure
' Comment     : Please give full Credit to 'paul_turcksin' for his great work.
'             : If you want to know more about TRIVERTEX color bypassing tequnique
'             : view his code at http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=52141&lngWId=1
' Purpose     : DrawGradient ( Horizontal/Vertical/Any Reapts )
' Argument    : Fastest/Smoothest/Simplest/Compact of all Gradients in PSC
'-------------------------------------------------------------------------

Option Explicit

Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, ByRef pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Integer
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Const COLOR_BTNFACE = 15

Private Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Public Enum GradientDirection
    [GR_Fill_None] = -1
    [gr_Fill_Horizontal] = 0
    [GR_Fill_Vertical] = 1
End Enum

'-------------------------------------------------------------------------
' Procedure  : SmoothGradient
' Upgraded   : Jim Jose
' Credits    : Based on an Excellent thought by 'paul_turcksin' to use
'              CopyMemory to fill Color Bytes to TRIVERTEX Structure
' Input      : hdcObject, GradientType, Start/End col, RepeatFactor
' OutPut     : Done?
' Purpose    : DrawGradient ( Horizontal/Vertical/Any Reapts )
'-------------------------------------------------------------------------

Public Function SmoothGradient(ByVal gHDC As Long, _
                                 ByVal gStartColor As Long, _
                                 ByVal gEndColor As Long, _
                                 ByVal gX As Double, _
                                 ByVal gY As Double, _
                                 ByVal gWidth As Double, _
                                 ByVal gHeight As Double, _
                                 ByVal gType As GradientDirection, _
                                 Optional ByVal gLeft2Right As Boolean = True, _
                                 Optional ByVal gRepeat As Long = 0) As Boolean
' The Variables
On Error GoTo Handle
Dim X               As Long
Dim tmpCol          As Long
Dim grdRect         As GRADIENT_RECT
Dim grdVertex(1)    As TRIVERTEX
Dim grdByteClrs(3)  As Byte
Dim grdByteVert(7)  As Byte

    ' Check if we need to fill
    If gType = GR_Fill_None Then Exit Function
    If gType = gr_Fill_Horizontal Then gWidth = gWidth / (gRepeat + 1) Else gHeight = gHeight / (gRepeat + 1)
    
    For X = 0 To gRepeat
    ' Check If the Fill is From Left to Right
    If gLeft2Right Then
        ' Init vertices : Set Position : Define Size
        grdVertex(0).X = gX: grdVertex(1).X = gX + gWidth
        grdVertex(0).Y = gY: grdVertex(1).Y = gY + gHeight
    Else
        ' Init vertices : Set Position : Define Size
        grdVertex(0).X = gX + gWidth: grdVertex(1).X = gX
        grdVertex(0).Y = gY + gHeight: grdVertex(1).Y = gY
    End If
   
    ' Init vertices :colors, initial
    CopyMemory grdByteClrs(0), gStartColor, &H4
    grdByteVert(1) = grdByteClrs(0)   ' Red
    grdByteVert(3) = grdByteClrs(1)   ' Green
    grdByteVert(5) = grdByteClrs(2)   ' Blue
    CopyMemory grdVertex(0).Red, grdByteVert(0), &H8

    ' Init vertices :colors, final
    CopyMemory grdByteClrs(0), gEndColor, &H4
    grdByteVert(1) = grdByteClrs(0)   ' Red
    grdByteVert(3) = grdByteClrs(1)   ' Green
    grdByteVert(5) = grdByteClrs(2)   ' Blue
    CopyMemory grdVertex(1).Red, grdByteVert(0), &H8

    ' Init gradient rect
    grdRect.UpperLeft = 0
    grdRect.LowerRight = 1

    ' Fill the DC
    GradientFill gHDC, grdVertex(0), 2, grdRect, 1, gType
    
    ' Proceed for Repeated Gradient
    tmpCol = gStartColor: gStartColor = gEndColor: gEndColor = tmpCol
    If gType = GR_Fill_Vertical Then gY = gY + gHeight Else gX = gX + gWidth
    
    Next X
    
Handle:
End Function




