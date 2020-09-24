VERSION 5.00
Begin VB.Form frmTag 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modify Mp3 Tag"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   720
      MaxLength       =   30
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtArtist 
      Height          =   285
      Left            =   720
      MaxLength       =   30
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   195
      Index           =   1
      Left            =   285
      TabIndex        =   2
      Top             =   600
      Width           =   345
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Artist:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   390
   End
End
Attribute VB_Name = "frmTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    On Error Resume Next
    
    Dim strFileName As String
    Dim strPath As String
    Dim strFileTAG As String
    Dim iListIndex As Integer
    
    'If txtArtist.Text = vbNullString And txtTitle.Text = vbNullString Then cmdCancel_Click
    
    'Set the fields
    gMp3Details.Artist = txtArtist.Text
    gMp3Details.Title = txtTitle.Text
    
    With frmMain
        'Set the temp variables
        strFileName = .lstFilename.Text
        strPath = .lstFullPath.Text
        strFileTAG = RTrim(gMp3Details.Artist) & " - " & RTrim(gMp3Details.Title)
        iListIndex = .lstFileTag.ListIndex
        
        'Save the TAG
        Call SetMp3Info(frmMain.lstFullPath)
        
        'Remove the mp3 from the list
        .lstFilename.RemoveItem (iListIndex)
        .lstFileTag.RemoveItem (iListIndex)
        .lstFullPath.RemoveItem (iListIndex)
        
        'Add the new items to update the lists
        .lstFilename.AddItem strFileName
        .lstFileTag.AddItem strFileTAG
        .lstFullPath.AddItem strPath
        
        .lstFilename.ListIndex = .lstFilename.ListCount - 1
    End With
    
    'Close the TAG dialog
    Unload Me
End Sub

Private Sub Form_Load()
    txtArtist.Text = GetArtist(frmMain.lstFileTag.Text)
    txtTitle.Text = GetTitle(frmMain.lstFileTag.Text)
    
End Sub

Private Sub Form_Terminate()
    Set frmTag = Nothing
End Sub

Private Sub txtArtist_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtTitle_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub
