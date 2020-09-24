VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP3 Genie"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   0
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   534
      TabIndex        =   19
      Top             =   0
      Width           =   8070
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   285
         Index           =   2
         Left            =   765
         TabIndex        =   22
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   240
         Index           =   1
         Left            =   585
         TabIndex        =   21
         Top             =   450
         Width           =   165
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "v.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   20
         Top             =   480
         Width           =   225
      End
   End
   Begin VB.ListBox lstFullPath 
      Height          =   3180
      Left            =   2760
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog dlgAdd 
      Left            =   5880
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Mp3 Filenames"
      Height          =   3975
      Index           =   1
      Left            =   4080
      TabIndex        =   5
      Top             =   840
      Width           =   3855
      Begin VB.CommandButton cmdPlay 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   3480
         Width           =   495
      End
      Begin VB.ListBox lstFilename 
         Height          =   3180
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3615
      End
      Begin MediaPlayerCtl.MediaPlayer mpPreview 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
         AudioStream     =   -1
         AutoSize        =   0   'False
         AutoStart       =   -1  'True
         AnimationAtStart=   -1  'True
         AllowScan       =   -1  'True
         AllowChangeDisplaySize=   -1  'True
         AutoRewind      =   0   'False
         Balance         =   0
         BaseURL         =   ""
         BufferingTime   =   5
         CaptioningID    =   ""
         ClickToPlay     =   -1  'True
         CursorType      =   0
         CurrentPosition =   -1
         CurrentMarker   =   0
         DefaultFrame    =   ""
         DisplayBackColor=   0
         DisplayForeColor=   16777215
         DisplayMode     =   0
         DisplaySize     =   4
         Enabled         =   -1  'True
         EnableContextMenu=   -1  'True
         EnablePositionControls=   -1  'True
         EnableFullScreenControls=   0   'False
         EnableTracker   =   -1  'True
         Filename        =   ""
         InvokeURLs      =   -1  'True
         Language        =   -1
         Mute            =   0   'False
         PlayCount       =   1
         PreviewMode     =   0   'False
         Rate            =   1
         SAMILang        =   ""
         SAMIStyle       =   ""
         SAMIFileName    =   ""
         SelectionStart  =   -1
         SelectionEnd    =   -1
         SendOpenStateChangeEvents=   -1  'True
         SendWarningEvents=   -1  'True
         SendErrorEvents =   -1  'True
         SendKeyboardEvents=   0   'False
         SendMouseClickEvents=   0   'False
         SendMouseMoveEvents=   0   'False
         SendPlayStateChangeEvents=   -1  'True
         ShowCaptioning  =   0   'False
         ShowControls    =   0   'False
         ShowAudioControls=   -1  'True
         ShowDisplay     =   0   'False
         ShowGotoBar     =   0   'False
         ShowPositionControls=   -1  'True
         ShowStatusBar   =   0   'False
         ShowTracker     =   -1  'True
         TransparentAtStart=   0   'False
         VideoBorderWidth=   0
         VideoBorderColor=   0
         VideoBorder3D   =   0   'False
         Volume          =   -600
         WindowlessVideo =   0   'False
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Mp3 Tags - 0 files"
      Height          =   3975
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3855
      Begin VB.CommandButton cmdModify 
         Caption         =   "Modify &Tag"
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   3480
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   3480
         Width           =   495
      End
      Begin VB.ListBox lstFileTag 
         Height          =   3180
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3615
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   3480
         Width           =   495
      End
   End
   Begin VB.Frame fraFrame 
      Height          =   735
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   5295
      Begin VB.CommandButton cmdSettings 
         Caption         =   "&Settings..."
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "File to TAG"
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdOption 
         Caption         =   "TAG to File"
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdFix 
      Caption         =   "Fix It!"
      Default         =   -1  'True
      Height          =   640
      Left            =   7283
      MaskColor       =   &H8000000F&
      Picture         =   "frmMain.frx":8C22
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4890
      Width           =   640
   End
   Begin MSComctlLib.ProgressBar pbrProcess 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame fraFrame 
      Caption         =   "0 Mp3's to process"
      Height          =   3135
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   5535
      Width           =   7815
      Begin MSComctlLib.ListView lvwProcess 
         Height          =   2775
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         PictureAlignment=   4
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TAG"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Option"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Filename"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Full Path"
            Object.Width           =   10583
         EndProperty
      End
   End
   Begin VB.Label lblEMail 
      AutoSize        =   -1  'True
      Caption         =   "wnigrini@softhome.net"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   210
      Left            =   120
      TabIndex        =   23
      Top             =   8670
      Width           =   1680
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Copyright © 1999 - 2003, Wouter Nigrini."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   5040
      TabIndex        =   18
      Top             =   8670
      Width           =   2910
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu mnuListClear 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private mItem As ListItem
    Private mblnStop As Boolean

Private Sub cmdAdd_Click()
    On Error GoTo CancelAdd
    
    Dim strFiles() As String
    Dim iX As Integer
    
    With dlgAdd
        .CancelError = True
        .DialogTitle = "Add Mp3's..."
        .InitDir = "c:"
        .Filter = "Mp3 Files (*.mp3)|*.mp3|Other Files (*.*)|*.*"
        .MaxFileSize = 32000
        .Flags = cdlOFNExplorer + cdlOFNAllowMultiselect + cdlOFNNoLongNames
        .ShowOpen
       
        strFiles = Split(.FileName, Chr(0))
        
        If UBound(strFiles) > 1 Then
           For iX = 1 To UBound(strFiles)
               lstFilename.AddItem strFiles(iX)
               If Right$(strFiles(0), 1) = "\" Then
                  lstFullPath.AddItem strFiles(0) & strFiles(iX)
               Else
                  lstFullPath.AddItem strFiles(0) & "\" & strFiles(iX)
               End If
               
               Call GetMp3Info(strFiles(iX))
               
               lstFileTag.AddItem RTrim(gMp3Details.Artist) & " - " & RTrim(gMp3Details.Title)
           Next iX
        Else
           lstFilename.AddItem .FileTitle
           lstFullPath.AddItem .FileName
           
           Call GetMp3Info(.FileName)
           
           lstFileTag.AddItem RTrim(gMp3Details.Artist) & " - " & RTrim(gMp3Details.Title)
        End If
    End With
    
    'Set the caption of the frame
    fraFrame(0).Caption = "Mp3 Tags - " & lstFileTag.ListCount & " files"
    
    'Select the first mp3 in the list
    lstFilename.ListIndex = 0
    
    Exit Sub
CancelAdd:
    If Err.Number = 32755 Then
       'Cancelled
    End If
End Sub

Private Sub cmdFix_Click()
    On Error Resume Next
    
    Dim iX As Integer
    Dim strPath As String
    Dim strFileTitle As String
    Dim strOption As String
    
    'Only start processing when there are items in the Processing list
    If lvwProcess.ListItems.Count = 0 Then
       MsgBox "First add mp3's to the processing list before attempting to fix any", vbInformation, "Fix It!"
       Exit Sub
    End If
    
    Select Case cmdFix.Caption
        Case "Fix It!"
            mblnStop = False
            
            cmdFix.Caption = "Stop!"
            
            pbrProcess.Visible = True
            fraFrame(2).Visible = False
            
            pbrProcess.Max = lvwProcess.ListItems.Count
            
            For iX = 1 To lvwProcess.ListItems.Count
                'Stop processing if the stop button is clicked
                If mblnStop = True Then Exit For
                
                'Set the values
                gMp3Details.Artist = GetArtist(lvwProcess.ListItems(iX))
                gMp3Details.Title = GetTitle(lvwProcess.ListItems(iX) & ".mp3")
                strOption = lvwProcess.ListItems(iX).SubItems(1)
                strFileTitle = lvwProcess.ListItems(iX).SubItems(2)
                strPath = lvwProcess.ListItems(iX).SubItems(3)
                    
                pbrProcess.Value = iX
                If strOption = "Fixed" Then
                   gMp3Details.Artist = GetArtist(strFileTitle)
                   gMp3Details.Title = GetTitle(strFileTitle)
                
                   Call SetMp3Info(strPath)
                   
                Else
                   'Rename the current file to the new one
                   Name strPath As Left$(strPath, Len(strPath) - Len(strFileTitle)) & RTrim(gMp3Details.Artist) & " - " & RTrim(gMp3Details.Title) & ".mp3"
                   
                   lvwProcess.ListItems(iX).SubItems(2) = RTrim(gMp3Details.Artist) & " - " & RTrim(gMp3Details.Title) & ".mp3"
                End If
                
                lvwProcess.ListItems(iX).SubItems(1) = "Complete"
                
                DoEvents
            Next iX
            
            pbrProcess.Visible = False
            fraFrame(2).Visible = True
            
            cmdFix.Caption = "Fix It!"
            
        Case "Stop!"
            'Stop processing
            mblnStop = True
                        
            cmdFix.Caption = "Fix It!"
                        
            pbrProcess.Visible = False
            fraFrame(2).Visible = True
    End Select
End Sub

Private Sub cmdModify_Click()
    'Check if any files exist in the list
    If lstFilename.ListCount = 0 Then Exit Sub
    
    frmTag.Show vbModal
End Sub

Private Sub cmdClear_Click()
    'Clear both lists
    lstFileTag.Clear
    lstFilename.Clear
    
    'Set the caption of the frame
    fraFrame(0).Caption = "Mp3 Tags - " & lstFileTag.ListCount & " files"
    
End Sub

Private Sub cmdOption_Click(Index As Integer)
    Dim ItemX As ListItem
    
    'Check if there is any items in the two lists
    If lstFilename.ListCount <> 0 Then
        If lstFilename.ListIndex = -1 Then lstFilename.ListIndex = 0
    
        If Index = 1 Then
           Set ItemX = lvwProcess.ListItems.Add(, , FixTag(lstFilename.Text))
           ItemX.SubItems(1) = "Fixed"
        Else
           Set ItemX = lvwProcess.ListItems.Add(, , lstFileTag.Text)
           ItemX.SubItems(1) = cmdOption(Index).Caption
        End If
        
        ItemX.SubItems(2) = lstFilename.Text
        ItemX.SubItems(3) = lstFullPath.Text
        
        If frmSettings.chkRemove.Value = vbChecked Then
           cmdRemove_Click
        End If
    End If
    
    'Set the caption of the frame
    fraFrame(3).Caption = lvwProcess.ListItems.Count & " Mp3's to process"

End Sub

Private Sub cmdPlay_Click()
    On Error GoTo SelectItem
    
    Select Case cmdPlay.Caption
        Case "4"
            mpPreview.FileName = lstFilename.Text
            mpPreview.Play
            
            cmdPlay.Caption = ";"
        Case ";"
            mpPreview.Stop
            
            cmdPlay.Caption = "4"
    End Select
        
    Exit Sub
SelectItem:
    MsgBox "Please select a file first or add mp3's to the list before attempting" _
         & vbNewLine & "to preview a file", vbInformation, "Mp3 Preview"
End Sub

Private Sub cmdRemove_Click()
    Dim iPreviousIndex As Integer
    
    'If the are items in the list
    If lstFilename.ListCount <> 0 Then
       
       If lstFilename.ListIndex = -1 Then lstFilename.ListIndex = 0
       
       iPreviousIndex = lstFileTag.ListIndex
       
       'Remove the selected item from both lists
       lstFileTag.RemoveItem (iPreviousIndex)
       lstFilename.RemoveItem (iPreviousIndex)
       lstFullPath.RemoveItem (iPreviousIndex)
       
       If iPreviousIndex = lstFilename.ListCount Then
          iPreviousIndex = lstFilename.ListCount - 1
       End If
           
       'Move to the previous item
       lstFilename.ListIndex = iPreviousIndex
    End If
    
    'Set the caption of the frame
    fraFrame(0).Caption = "Mp3 Tags - " & lstFileTag.ListCount & " files"
End Sub

Private Sub cmdSettings_Click()
    frmSettings.Show vbModal
End Sub

Private Sub Form_Activate()
    cmdAdd.SetFocus
End Sub

Private Sub Form_Load()
    'Set the version number
    lblVersion(0).Caption = "v." & App.Major
    lblVersion(1).Caption = "." & App.Minor
    lblVersion(2).Caption = "." & App.Revision
    
    'Set the progressbar colors
    Call SetPBColor(pbrProcess.hwnd, 245, 220, 39, 245, 163, 7)
End Sub

Private Sub Form_Terminate()
    Set frmMain = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Unload the settings form
    Unload frmSettings
End Sub

Private Sub lblEMail_Click()
    'Open the default mail program
    Call Shell("start mailto:wnigrini@softhome.net", vbNormalFocus)
End Sub

Private Sub lstFileTag_Click()
    'Synchronize the two listbox items
    lstFilename.ListIndex = lstFileTag.ListIndex
    
    lstFullPath.ListIndex = lstFilename.ListIndex
End Sub

Private Sub lstFileTag_DblClick()
    'Chnage the Tag of the selected file
    cmdModify_Click
End Sub

Private Sub lstFilename_Click()
    'Synchronize the two listbox items
    lstFileTag.ListIndex = lstFilename.ListIndex
End Sub

Private Sub lstFilename_DblClick()
    'Play the selected file
    cmdPlay_Click
End Sub

Private Sub lvwProcess_DblClick()
    Select Case mItem.SubItems(1)
        Case "TAG to File"
            mItem.SubItems(1) = "Fixed"
        Case "Fixed"
            mItem.SubItems(1) = "TAG to File"
    End Select
End Sub

Private Sub lvwProcess_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set mItem = Item
End Sub

Private Sub lvwProcess_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Check for right mouse button
    If Button = vbRightButton Then Me.PopupMenu mnuList
       
End Sub

Private Sub mnuListClear_Click()
    'Clear the process list
    lvwProcess.ListItems.Clear
    
    'Set the caption of the frame
    fraFrame(3).Caption = "0 Mp3's to process"
End Sub
