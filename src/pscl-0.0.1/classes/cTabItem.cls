VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' ***************************************************************************************
' * Project  | GpTabStrip                                                               *
' *----------|--------------------------------------------------------------------------*
' * Version  | V1.0                                                                     *
' *----------|--------------------------------------------------------------------------*
' * Author   | Genghis Khan(GuangJian Guo)                                              *
' *----------|--------------------------------------------------------------------------*
' * WebSite  | http://www.itkhan.com                                                    *
' *----------|--------------------------------------------------------------------------*
' * MailTo   | webmaster@itkhan.com                                                     *
' *----------|--------------------------------------------------------------------------*
' * Date     | 13 April 2003                                                            *
' ***************************************************************************************

' ======================================================================================
' Enums
' ======================================================================================

' 标头文字对齐方式
Public Enum GPTAB_ALIGNMENT_METHOD
    GpTabCaptionLeft = 0          '/* 左对齐 */
    GpTabCaptionRight = 1         '/* 右对齐 */
    GpTabCaptionCenter = 2        '/* 置中 */
End Enum

' 标头图标显示位置
Public Enum GPTAB_ICONALIGN_METHOD
    GpTabIconLeft = 0             '/* 靠左显示 */
    GpTabIconRight = 1            '/* 靠右显示 */
End Enum

' ======================================================================================
' Private variables:
' ======================================================================================

Private m_strCaption            As String
Private m_lngIndex              As Long
Private m_lngIcon               As Long
Private m_lngTop                As Long
Private m_lngLeft               As Long
Private m_lngHeight             As Long
Private m_lngDefaultWidth       As Long
Private m_lngWidth              As Long
Private m_strKey                As String
Private m_strTag                As String
Private m_blnSelected           As Boolean
Private m_oleBackColor          As OLE_COLOR
Private m_oleForeColor          As OLE_COLOR
Private m_udtAlignment          As GPTAB_ALIGNMENT_METHOD
Private m_udtIconAlign          As GPTAB_ICONALIGN_METHOD

' ======================================================================================
' Events
' ======================================================================================

Public Event AlignmentChanged()
Public Event CaptionChanged()
Public Event IconChanged()
Public Event IconAlignChanged()
Public Event SelectedChanged()
Public Event KeyChanged(ByVal OldKey As String) '/* 校验修改的Key是否唯一值 */

Public Property Get Alignment() As GPTAB_ALIGNMENT_METHOD
    Alignment = m_udtAlignment
End Property

Public Property Let Alignment(ByVal NewValue As GPTAB_ALIGNMENT_METHOD)
    If m_udtAlignment <> NewValue Then
       m_udtAlignment = NewValue
       RaiseEvent AlignmentChanged
    End If
End Property

Public Property Get Caption() As String
    Caption = m_strCaption
End Property

Public Property Let Caption(ByVal NewValue As String)
    If m_strCaption <> NewValue Then
       m_strCaption = NewValue
       RaiseEvent CaptionChanged
    End If
End Property

Friend Property Get DefaultWidth() As Long
    DefaultWidth = m_lngDefaultWidth
End Property

Friend Property Let DefaultWidth(ByVal NewValue As Long)
    m_lngDefaultWidth = NewValue
End Property

Friend Property Get Height() As Long
    Height = m_lngHeight
End Property

Friend Property Let Height(ByVal NewValue As Long)
    m_lngHeight = NewValue
End Property

Public Property Get Icon() As Long
    Icon = m_lngIcon
End Property

Public Property Let Icon(ByVal NewValue As Long)
    If m_lngIcon <> NewValue Then
       m_lngIcon = NewValue
       RaiseEvent IconChanged
    End If
End Property

Public Property Get IconAlign() As GPTAB_ICONALIGN_METHOD
    IconAlign = m_udtIconAlign
End Property

Public Property Let IconAlign(ByVal NewValue As GPTAB_ICONALIGN_METHOD)
    If m_udtIconAlign <> NewValue Then
       m_udtIconAlign = NewValue
       RaiseEvent IconAlignChanged
    End If
End Property

Public Property Get Index() As Long
    Index = m_lngIndex
End Property

Public Property Get Key() As String
    Key = m_strKey
End Property

Public Property Let Key(ByVal NewValue As String)
    Dim strOld As String
    
    If m_strKey <> NewValue Then
       strOld = m_strKey
       m_strKey = NewValue
       RaiseEvent KeyChanged(strOld)
    End If
End Property

Friend Property Get Left() As Long
    Left = m_lngLeft
End Property

Friend Property Let Left(ByVal NewValue As Long)
    m_lngLeft = NewValue
End Property

Public Property Get Selected() As Boolean
    Selected = m_blnSelected
End Property

Public Property Let Selected(ByVal NewValue As Boolean)
    If m_blnSelected <> NewValue Then
       m_blnSelected = NewValue
       RaiseEvent SelectedChanged
    End If
End Property

Friend Property Let TabIndex(ByVal NewValue As Long)
    m_lngIndex = NewValue
End Property

Public Property Get Tag() As String
    Tag = m_strTag
End Property

Public Property Let Tag(ByVal NewValue As String)
    m_strTag = NewValue
End Property

Friend Property Get Top() As Long
    Top = m_lngTop
End Property

Friend Property Let Top(ByVal NewValue As Long)
    m_lngTop = NewValue
End Property

Friend Property Get Width() As Long
    Width = m_lngWidth
End Property

Friend Property Let Width(ByVal NewValue As Long)
    m_lngWidth = NewValue
End Property

Private Sub Class_Initialize()
    m_udtAlignment = GpTabCaptionCenter
    m_udtIconAlign = GpTabIconLeft
End Sub

