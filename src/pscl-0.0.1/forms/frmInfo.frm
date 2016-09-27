VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1185
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   2565
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   2565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAction 
      Caption         =   "Action"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAlwaysBack 
      Caption         =   "Back.."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compile Success!"
      Height          =   195
      Left            =   660
      TabIndex        =   0
      Top             =   120
      Width           =   1245
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAction_Click()
    Select Case cmdAction.Caption
        Case "Back..": frmMain.RunEnabled = False: Unload Me
        Case "Run.."
        Dim ProgramID As Long
        sFileToRun = """" & sFileToRun & """"
        ProgramID = Shell(sFileToRun, vbNormalFocus)
        hWndProg = OpenProcess(PROCESS_ALL_ACCESS, False, ProgramID)
        Unload Me
    End Select
End Sub

Private Sub cmdAlwaysBack_Click()
    frmMain.progBar.Value = 0
    frmMain.lblStatus.Caption = "Ready"
    frmMain.RunEnabled = False
    frmMain.txtEditor.Enabled = True
    Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    cmdAction.Left = frmInfo.Width - cmdAction.Width - 200
    cmdAction.Top = frmInfo.Height - cmdAction.Height - 200
    cmdAlwaysBack.Left = 50
    cmdAlwaysBack.Top = frmInfo.Height - cmdAlwaysBack.Height - 200
    rtfSummary.Width = frmInfo.Width - 200
    rtfSummary.Height = frmInfo.Height - 1600
    shpHead.Width = frmInfo.Width - 200
End Sub

