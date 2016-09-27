VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Pascaline"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Visit Website"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.ListBox lstBundle 
      Height          =   645
      ItemData        =   "frmAbout.frx":0E42
      Left            =   120
      List            =   "frmAbout.frx":0E4C
      TabIndex        =   3
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label lblNotice 
      Caption         =   $"frmAbout.frx":0EA1
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2016 Dani Pragustia"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2385
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For 32-bit Windows Development"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pascaline 2016.1"
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
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0F54
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Shell "explorer.exe www.vaxza.ga", vbNormalFocus
End Sub
