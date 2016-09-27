VERSION 5.00
Begin VB.Form frmReplace 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Replace..."
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtReplace 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   4695
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Replace To :"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find What :"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReplace_Click()
If MsgBox("All string with '" & txtFind.Text & "' will be replaced, Are you sure?", vbQuestion + vbYesNo) = vbYes Then
    frmMain.txtEditor.Text = Replace(frmMain.txtEditor.Text, txtFind.Text, txtReplace.Text)
    MsgBox "Replaced successfully", vbInformation
End If
End Sub

Private Sub txtFind_Change()
    If txtFind.Text <> "" And txtReplace.Text <> "" Then
        cmdReplace.Enabled = True
    Else
        cmdReplace.Enabled = False
    End If
End Sub

Private Sub txtReplace_Change()
    If txtFind.Text <> "" And txtReplace.Text <> "" Then
        cmdReplace.Enabled = True
    Else
        cmdReplace.Enabled = False
    End If
End Sub
