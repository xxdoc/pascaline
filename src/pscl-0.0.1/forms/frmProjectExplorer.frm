VERSION 5.00
Begin VB.Form frmProjectExplorer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Explorer"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProjectExplorer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstProject 
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "frmProjectExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    lstProject.Height = Me.ScaleHeight
    lstProject.Width = Me.ScaleWidth
End Sub

Private Sub lstProject_Click()
    Dim NodeText As String
    
    If lstProject.ListCount = 0 Then Exit Sub
    
    NodeText = lstProject.List(lstProject.ListIndex)
    NodeText = Replace(NodeText, "[MODULE] - ", "")
    NodeText = Replace(NodeText, "[ENTRY] - ", "")
    NodeText = Replace(NodeText, "[RESOURCE] - ", "")
    Debug.Print lstProject.List(lstProject.ListIndex)
    Debug.Print "result : " & NodeText
    frmMain.txtEditor.Text = GetVirtualFileContent(NodeText)
    frmMain.txtEditor.Enabled = True
    If GetVirtualFileExtension(NodeText) = EX_DIALOG Then
        'Show Design Form
        
    Else
        'Hide Design Form
    End If
End Sub
