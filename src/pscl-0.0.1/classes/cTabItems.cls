VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabItems"
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
' Constants
' ======================================================================================
Private Const MODULE_NAME = "CTabItems"

' ======================================================================================
' Private variables:
' ======================================================================================
Private m_blnAddTabitem           As Boolean
Private m_colTabs                 As Collection
Private WithEvents m_clsTab       As cTabItem
Attribute m_clsTab.VB_VarHelpID = -1

' ======================================================================================
' Events
' ======================================================================================
Public Event TabAddNew()
Public Event TabCaptionChanged(ByVal Index As Long)
Public Event TabAlignmentChanged(ByVal Index As Long)
Public Event TabIconChanged(ByVal Index As Long)
Public Event TabIconAlignChanged(ByVal Index As Long)
Public Event TabSelectedChanged(ByVal Index As Long)
Public Event TabRemove()

Public Function Add(Optional ByVal Index As Long = -1, _
                    Optional ByVal Key As String = "", _
                    Optional ByVal Caption As String = "", _
                    Optional ByVal Icon As Long = -1) As cTabItem
    Dim lngI        As Long
    Dim lngindex    As Long
    
    If Not IsValidTabItemIndex(Index) Then Exit Function
    If Not IsValidTabItemKey(Key) Then Exit Function
    m_blnAddTabitem = True
    If (Index <= 0) Then
       lngindex = m_colTabs.Count + 1
    Else
       lngindex = Index
    End If
    
    Set m_clsTab = New cTabItem
    With m_clsTab
         .TabIndex = lngindex
         .Key = Key
         .Caption = Caption
         .Icon = Icon
    End With
    
    If m_colTabs.Count > 0 Then
       If lngindex <= m_colTabs.Count Then
          m_colTabs.Add m_clsTab, , , lngindex
          Call ReCreateIndex
       Else
          m_colTabs.Add m_clsTab
       End If
    Else
       m_colTabs.Add m_clsTab
    End If
    Set Add = m_clsTab
    RaiseEvent TabAddNew
    m_blnAddTabitem = False
End Function

Public Sub Clear()
    Call Class_Initialize
    RaiseEvent TabRemove
End Sub

Public Property Get Count() As Long
    Count = m_colTabs.Count
End Property

Private Function GetSubitemIndex(ByVal Key As Variant) As Long
    Dim lngI                    As Long
    Dim lngindex                As Long
    
    If IsNumeric(Key) Then
       lngindex = CLng(Key)
       If lngindex <= 0 Then
          Err.Raise 380, App.EXEName & ".CTabItems"
       End If
    Else
       Key = CStr(Key)
       For lngI = 1 To m_colTabs.Count
           If (m_colTabs(lngI).Key = Key) Then
              lngindex = lngI
              Exit For
           End If
       Next lngI
       If lngindex <= 0 Then
          Err.Raise 35601, App.EXEName & ".CTabItems"
       End If
    End If
ByeBye:
    GetSubitemIndex = lngindex
End Function

Private Function IsValidTabItemIndex(ByVal Index As String) As Boolean
    If Index = 0 Then
       Err.Raise 9, App.EXEName & ".CTabItems"
       Exit Function
    End If
    If m_colTabs.Count > 0 Then
       If Index > m_colTabs.Count + 1 Then
          Err.Raise 9, App.EXEName & ".CTabItems"
          Exit Function
       End If
    Else
       If Index >= 0 And Index <> 1 Then
          Err.Raise 9, App.EXEName & ".CTabItems"
          Exit Function
       End If
    End If
    IsValidTabItemIndex = True
End Function

Private Function IsValidTabItemKey(ByVal Key As String, Optional Index As Long = 0) As Boolean
    Dim lngI As Long
    
    If (Key <> "") Then
       For lngI = 1 To m_colTabs.Count
          If lngI <> Index Then
             If (m_colTabs(lngI).Key = Key) Then
                Err.Raise 457, App.EXEName & ".CTabItems"
                Exit Function
             End If
          End If
       Next
    End If
    IsValidTabItemKey = True
End Function

' ƒ¨»œ Ù–‘
Public Property Get Item(ByVal Index As Variant) As cTabItem
    Dim lngindex As Long
    On Error Resume Next
    lngindex = GetSubitemIndex(Index)
    Set m_clsTab = m_colTabs(Index)
    Set Item = m_clsTab
End Property

Private Sub ReCreateIndex()
    Dim lngI               As Long
    Dim lngListCount       As Long
    Dim clsCuItem          As cTabItem
    
    lngListCount = m_colTabs.Count
    If lngListCount <= 0 Then Exit Sub
    For lngI = 1 To lngListCount
        Set clsCuItem = m_colTabs(lngI)
        clsCuItem.TabIndex = lngI
    Next
    Set clsCuItem = Nothing
End Sub

Public Sub Remove(ByVal Index As Long)
    Dim lngI As Long
    
    m_colTabs.Remove Index
    If Index <= m_colTabs.Count Then
       For lngI = Index To m_colTabs.Count
          m_colTabs(lngI).Index = lngI
       Next
    End If
    RaiseEvent TabRemove
End Sub

Private Sub Class_Initialize()
    Call Class_Terminate
    Set m_colTabs = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colTabs = Nothing
    Set m_clsTab = Nothing
End Sub

Private Sub m_clsTab_AlignmentChanged()
    If m_blnAddTabitem Then Exit Sub
    RaiseEvent TabAlignmentChanged(m_clsTab.Index)
End Sub

Private Sub m_clsTab_CaptionChanged()
    If m_blnAddTabitem Then Exit Sub
    RaiseEvent TabCaptionChanged(m_clsTab.Index)
End Sub

Private Sub m_clsTab_IconAlignChanged()
    If m_blnAddTabitem Then Exit Sub
    RaiseEvent TabIconAlignChanged(m_clsTab.Index)
End Sub

Private Sub m_clsTab_IconChanged()
    If m_blnAddTabitem Then Exit Sub
    RaiseEvent TabIconChanged(m_clsTab.Index)
End Sub

Private Sub m_clsTab_KeyChanged(ByVal OldKey As String)
    If m_blnAddTabitem Then Exit Sub
    If IsValidTabItemKey(m_clsTab.Key, m_clsTab.Index) Then
       '
    End If
End Sub

Private Sub m_clsTab_SelectedChanged()
    If m_blnAddTabitem Then Exit Sub
    RaiseEvent TabSelectedChanged(m_clsTab.Index)
End Sub


