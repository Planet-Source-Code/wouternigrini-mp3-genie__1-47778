VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mp3 Genie Settings"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraGeneral 
      Caption         =   "General"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtArtistCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Text            =   "2"
         Top             =   1080
         Width           =   285
      End
      Begin VB.CheckBox chkGroup 
         Caption         =   "&Group mp3's into Artist folders"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2895
      End
      Begin VB.CheckBox chkRemove 
         Caption         =   "&Remove mp3 from list when added"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "mp3's of the same artist was found"
         Enabled         =   0   'False
         Height          =   390
         Index           =   1
         Left            =   2160
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Group artists after"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "&Defaults"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDefault_Click()
    'Load the default settings
    
End Sub

Private Sub cmdOk_Click()
    'Unload the current form
    Unload Me
End Sub

Private Sub Form_Terminate()
    'Free memory
    Set frmSettings = Nothing
End Sub

