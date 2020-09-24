Attribute VB_Name = "modEngine"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const CCM_FIRST = &H2000
Public Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Public Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
Public Const WM_USER = &H400
Public Const PBM_SETBARCOLOR = (WM_USER + 9)

Public Type Info
    Artist As String * 30
    Title As String * 30
End Type

Public gMp3Details As Info
Private mblnReadError As Boolean

Public Sub GetMp3Info(ByVal strFileName As String)
    Dim iFileNum As Integer
    Dim lFilePos As Long
    Dim strData As String * 128
    Dim iA As Integer
    Dim strTitle As String
    Dim strArtist As String
    Dim strTemp As String
    
    'Check if the file is valid
    If CheckFile(strFileName) = True Then
       iFileNum = FreeFile
       lFilePos = FileLen(strFileName) - 127
       
       If lFilePos > 0 Then
          'Open the file and read the data
          Open strFileName For Binary As #iFileNum
          Get #iFileNum, lFilePos, strData
          Close #iFileNum
       End If
       
       gMp3Details.Artist = vbNullString
       gMp3Details.Title = vbNullString
       
       'Get the TAG info
       If Left(strData, 3) = "TAG" Then
          strTitle = Mid$(strData, 4, 30)
          For iA = 1 To Len(strTitle)
              If Mid(strTitle, iA, 1) <> Chr(0) Then
                 strTemp = strTemp & Mid(strTitle, iA, 1)
              End If
          Next iA
          gMp3Details.Title = strTemp
          
          strTemp = ""
          strArtist = Mid$(strData, 34, 30)
          For iA = 1 To Len(strArtist)
              If Mid(strArtist, iA, 1) <> Chr(0) Then
                 strTemp = strTemp & Mid(strArtist, iA, 1)
              End If
          Next iA
          gMp3Details.Artist = strTemp
       End If
    Else
       MsgBox "Info Unavailible"
    End If
End Sub

Public Function FixTag(ByVal strFileName As String)
    FixTag = GetArtist(strFileName) & " - " & GetTitle(strFileName)
End Function

Public Sub SetMp3Info(ByVal strFileName As String)
    Dim iFileNum As Integer

    iFileNum = FreeFile
    Open strFileName For Binary Access Write As #iFileNum
        Seek #iFileNum, FileLen(strFileName) - 127
        Put #iFileNum, , "TAG"
        Put #iFileNum, , gMp3Details.Title
        Put #iFileNum, , gMp3Details.Artist
    Close #iFileNum
End Sub

Public Function GetTitle(strTag As String) As String
    Dim iPos As Integer
    
    If mblnReadError = False Then
       iPos = InStr(1, strTag, " - ") + 2
        
       GetTitle = Right$(strTag, Len(strTag) - iPos)
        
       'Remove the mp3 extention from the string
       If Right(GetTitle, 4) = ".mp3" Then
          GetTitle = Left(GetTitle, Len(GetTitle) - 4)
       End If
    Else
       GetTitle = InputBox("Please enter the Title due to reading difficulty", "Title")
    End If
    
End Function

Public Function GetArtist(strTag As String) As String
    On Error GoTo Artist
    
    Dim iPos As Integer
    Dim strArtist As String
    
    mblnReadError = False
    
    iPos = InStr(1, strTag, " - ")
    GetArtist = Left$(strTag, iPos - 1)
    
    Exit Function
Artist:
    GetArtist = InputBox("Please enter Artist due to reading difficulty", "Artist")
    
    mblnReadError = True
End Function

Public Sub SetPBColor(pbBar As Long, bRed As Integer, bGreen As Integer, bBlue As Integer, fRed As Integer, fGreen As Integer, fBlue As Integer)
    SendMessage pbBar, PBM_SETBKCOLOR, 0, ByVal RGB(bRed, bGreen, bBlue)
    SendMessage pbBar, PBM_SETBARCOLOR, 0, ByVal RGB(fRed, fGreen, fBlue)
End Sub

Private Function CheckFile(ByVal strFileName As String) As Boolean
    If LCase(Right(strFileName, 3)) = "mp3" Then
       CheckFile = True
    Else
       CheckFile = False
    End If
End Function

'Private Function GetInfo(ByVal sFilename) As Info
'    Dim i As Info
'    GetInfo = i
'    Dim s
'    s = sFilename
'
'
'    If InStrRev(s, "\") > 0 Then 'it's a full path
'        s = Mid(s, InStrRev(s, "\") + 1)
'    End If
'
'    'drop extension
'    s = Left(s, InStrRev(s, ".", , vbTextCompare) - 1)
'    s = Replace(Trim(s), " ", " ")
'    s = Trim(s)
'
'
'
'    If CountItems(s, " ") < 1 Then
'        i.sTitle = Replace(s, "_", " ")
'        GetInfo = i
'        Exit Function
'    End If
'
'    s = Trim(Replace(s, "_", " "))
'
'
'    If Left(s, 1) = "(" And CountItems(s, "-") < 3 Then
'        i.sArtist = Mid(s, 2, InStr(s, ")") - 2)
'        s = Trim(Mid(s, InStr(s, ")") + 1))
'
'
'        If Left(s, 1) = "-" Then 'grab title
'            i.sTitle = Trim(Mid(s, 2))
'        Else 'grab title anyway
'
'
'            If InStr(s, "-") > 0 Then
'                i.sAlbum = Mid(s, InStr(s, "-") + 1)
'                i.sTitle = Left(s, InStr(s, "-") - 1)
'            Else
'                i.sTitle = Trim(s)
'            End If
'        End If
'    Else
'        Dim aThings
'        Dim l
'        aThings = Split(s, "- ")
'
'
'        For l = 0 To UBound(aThings)
'
'
'            If Not IsNumeric(aThings(l)) Then
'
'
'                If i.sArtist = "" Then
'                    i.sArtist = aThings(l)
'                Else
'
'
'                    If IsNumeric(aThings(l - 1)) Then ' title
'
'
'                        If i.sTitle = "" Then
'                            i.sTitle = aThings(l)
'                        End If
'                    ElseIf i.sAlbum = "" Then
'                        i.sAlbum = aThings(l)
'                    End If
'                End If
'            End If
'        Next ' i
'
'    End If
'
'    i.sArtist = Replace(Replace(i.sArtist, "(", ""), ")", "")
'
'
'
'    If Left(s, 1) <> "(" And i.sTitle = "" And (InStr(sFilename, "\") <> InStrRev(sFilename, "\")) Then
'        ' recurse
'        GetInfo = GetInfo(FixDir(sFilename))
'    Else
'        GetInfo = i
'    End If
'End Function
'
