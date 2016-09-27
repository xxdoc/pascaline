VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' =========================================================================
' vbAccelerator Image List Control Demonstrator
' Copyright © 1998 Steve McMahon (steve@dogma.demon.co.uk)
'
' Implements an Image List control in VB using COMCTL32.DLL
'
' Visit vbAccelerator at www.dogma.demon.co.uk
'
' Modified www.cyberactivex.com
' Consolidated AddFromHandle into single call.
' =========================================================================
'
' -----------
' API
' -----------
' General:

Private Declare Function ImageList_Create Lib "comctl32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_AddMasked Lib "comctl32" (ByVal hImageList As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_GetBkColor Lib "comctl32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_AddIcon Lib "comctl32" (ByVal himl As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_GetIcon Lib "comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal fuFlags As Long) As Long
Private Declare Function ImageList_GetImageCount Lib "comctl32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long
' ImageList functions:
' Draw:

' Messages:
Private Const ILC_MASK = &H1&

' -----------
' ENUMS
' -----------
Public Enum ImageTypes
    IMAGE_BITMAP = 0
    IMAGE_ICON = 1
    IMAGE_CURSOR = 2
End Enum

Public Enum eilColourDepth
    ILC_COLOR = &H0
    ILC_COLOR4 = &H4
    ILC_COLOR8 = &H8
    ILC_COLOR16 = &H10
    ILC_COLOR24 = &H18
    ILC_COLOR32 = &H20
End Enum

'ILC_MIRROR 'Version 6.00. Microsoft® Windows® can be mirrored to display languages such as Hebrew or Arabic that read right-to-left. If the image list is created on a mirrored version of Windows, then the images in the lists are mirrored, that is, they are flipped so they display from right to left. Use this flag on a mirrored version of Windows to instruct the image list not to automatically mirror images.
'ILC_PERITEMMIRROR'Version 6.00. Specify this flag if ILC_MIRROR is used on an image list that contains a strip of images. ILC_MIRROR must be specified for this flag to have any effect.

' ------------------
' Private variables:
' ------------------
Private m_hIml          As Long
Private m_lIconSizeX    As Long
Private m_lIconSizeY    As Long
Private m_TransColor    As Long
Private m_eColourDepth  As eilColourDepth

Public Property Get BackColor() As Long

    If (himl <> 0) Then
        BackColor = ImageList_GetBkColor(himl)
    End If

End Property

Public Property Get IconSizeX() As Long

    IconSizeX = m_lIconSizeX

End Property

Public Property Get IconSizeY() As Long

    IconSizeY = m_lIconSizeY

End Property

'Purpose:Consolidate code so only one call is needed.

Public Function CreateFromHandle( _
                                 ByVal eDepth As eilColourDepth, _
                                 ByVal lSizeX As Long, _
                                 ByVal lSizeY As Long, _
                                 ByVal hImage As Long, _
                                 Optional ByVal lBackColor As OLE_COLOR = -1) As Boolean

Dim lR               As Long

    m_eColourDepth = eDepth
    m_lIconSizeX = lSizeX
    m_lIconSizeY = lSizeY
    m_TransColor = lBackColor
    Destroy
    ' Create the Imagelist:
    m_hIml = ImageList_Create(lSizeX, lSizeY, ILC_MASK Or eDepth, 4, 4)
    If (m_hIml <> 0) And (m_hIml <> -1) Then
        ' Adds an image or series of images from a GDI image handle.
        If (hImage <> 0) Then
            ' And add it to the image list:
            lR = ImageList_AddMasked(himl, hImage, lBackColor)
        Else
            lR = -1
        End If
    Else
        m_hIml = 0
        lR = -1
    End If
    CreateFromHandle = (lR <> -1)

End Function

'Purpose:Kill the image list if we have one.

Public Sub Destroy()

    If (himl <> 0) Then
        ImageList_Destroy himl
        m_hIml = 0
    End If

End Sub

Public Sub Draw(Index As Long, hDc As Long, x As Long, y As Long)

'i, hdcDst, x, y, fStyle

    ImageList_Draw m_hIml, Index, hDc, x, y, 0

End Sub

'Purpose:Returns the ImageList handle.

Public Property Get himl() As Long

    himl = m_hIml

End Property

Public Property Get ImageCount() As Integer

' Returns the number of images in the ImageList:

    If (himl <> 0) Then
        ImageCount = ImageList_GetImageCount(himl)
    End If

End Property

Private Sub Class_Initialize()

    m_lIconSizeX = 16
    m_lIconSizeY = 16
    m_eColourDepth = ILC_COLOR

End Sub

Private Sub Class_Terminate()

    Destroy

End Sub
