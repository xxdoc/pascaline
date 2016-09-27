VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Project"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4305
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
   ScaleHeight     =   2880
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open existing..."
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton cmdEmpty 
      Caption         =   "Blank Project"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   4095
   End
   Begin VB.CommandButton cmdConsole 
      Caption         =   "Windows Console "
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   4095
   End
   Begin VB.CommandButton cmdDLL 
      Caption         =   "Dynamic Link Library"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4095
   End
   Begin VB.CommandButton cmdGUI 
      Caption         =   "Windows GUI"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select New Project :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1665
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub TemplateGUI()
    frmMain.MakeInitControls
    SetVirtualFileContent "Entry Point", _
                "application PE GUI;" & vbCrLf & vbCrLf & _
                "import MessageBox ascii lib ""USER32.DLL"",4;" & vbCrLf & vbCrLf & _
                "entry" & vbCrLf & vbCrLf & _
                vbTab & "MessageBox(0,""Hello World!"",""Pascaline"",$20);" & vbCrLf & _
                vbCrLf & _
                "end."
    Unload Me
End Sub

Sub TemplateDLL()
    frmMain.MakeInitControls
    SetVirtualFileContent "Entry Point", _
                "application PE GUI DLL;" & vbCrLf & vbCrLf & _
                "export IsInitialized();" & vbCrLf & _
                vbTab & "return(TRUE);" & vbCrLf & _
                "end;" & vbCrLf
    Unload Me
End Sub

Sub TemplateCUI()
    frmMain.MakeInitControls
    SetVirtualFileContent "Entry Point", _
                "application PE CUI;" & vbCrLf & vbCrLf & _
                "include ""Windows.inc"", ""Console.inc"";" & vbCrLf & vbCrLf & _
                "entry" & vbCrLf & vbCrLf & _
                vbTab & "Console.Init(""Pascaline"");" & vbCrLf & _
                vbTab & "Console.Write(""Hello World!"");" & vbCrLf & _
                vbTab & "Console.Read();" & vbCrLf & _
                vbCrLf & _
                "end."
    Unload Me
End Sub

Private Sub cmdConsole_Click()
TemplateCUI
End Sub

Private Sub cmdDLL_Click()
TemplateDLL
End Sub

Private Sub cmdEmpty_Click()
frmMain.MakeInitControls
End Sub

Private Sub cmdGUI_Click()
TemplateGUI
End Sub

Private Sub cmdOpen_Click()
frmMain.OpenProject
End Sub

Private Sub Form_Load()
    frmMain.txtEditor.Text = ""
End Sub
