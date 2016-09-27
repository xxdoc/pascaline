VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Splash"
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7290
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
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Splash"
   MaxButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   7290
      TabIndex        =   0
      Top             =   2445
      Width           =   7290
      Begin VB.Label lblLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.vaxza.ga"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5640
         MouseIcon       =   "frmSplash.frx":0E42
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   60
         Width           =   1575
      End
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © Dani Pragustia"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   60
         Width           =   1980
      End
   End
   Begin VB.Timer tmrLoading 
      Interval        =   2000
      Left            =   6720
      Top             =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   1800
      X2              =   6360
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Left            =   -240
      Top             =   -120
      Width           =   975
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   855
      Left            =   -120
      Top             =   -360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Left            =   6600
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait..."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   990
   End
   Begin VB.Image Image2 
      Height          =   1080
      Left            =   1680
      Picture         =   "frmSplash.frx":0F94
      Top             =   600
      Width           =   4530
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1080
      Picture         =   "frmSplash.frx":384E
      Top             =   720
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   855
      Left            =   6480
      Top             =   2040
      Width           =   975
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim CmdParams As Variant
    Dim SourceCon As String
    Dim SourFile As String
    Dim DestFile As String
    Dim cdCMD As UnicodeFileDialog
    Set cdCMD = New UnicodeFileDialog
    If Command() <> "" Then
        CmdParams = Split(Command(), " ")
        If UBound(CmdParams) = 1 Then
            Timer1.Enabled = False
            If CmdParams(0) = "/c" Then
                SourFile = CmdParams(1)
                InitVirtualFiles
                If Not Dir(CmdParams(1)) <> "" Then MsgBox "File '" & CmdParams(1) & "' does not exist.": End
                Open SourFile For Binary As #1
                    SourceCon = Space(LOF(1))
                    Get #1, , SourceCon
                Close #1
                CreateVirtualFile "Entry Point", EX_ENTRY, SourceCon
            
                On Error GoTo CmdExeError
                With cdCMD
                .Filter = "Executable Files|*.exe"
                .CancelError = True
                .ShowSave Me.hWnd
                End With
                
                If Dir(cdCMD.FileName) <> "" Then
                    If MsgBox("File already exists! Overwrite?", vbYesNo + vbCritical) = vbNo Then
                        Exit Sub
                    End If
                End If
                IsCmdCompile = True
                Compile cdCMD.FileName, False
CmdExeError:
                End
            Else
                MsgBox "Usage: /c source.* destination.*"
            End If
        End If
        Unload Me
        frmMain.Show
    End If
End Sub

Private Sub lblLink_Click()
Shell "explorer http://www.vaxza.ga", vbNormalFocus
End Sub

Private Sub tmrLoading_Timer()
    Unload Me
    frmMain.Show
End Sub
