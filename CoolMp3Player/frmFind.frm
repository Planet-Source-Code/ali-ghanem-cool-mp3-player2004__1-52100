VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search For Mp3"
   ClientHeight    =   4395
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7320
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   7260
      TabIndex        =   1
      Top             =   4140
      Width           =   7320
   End
   Begin VB.ListBox List1 
      Height          =   3465
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7335
   End
   Begin VB.Menu mnuFindFiles 
      Caption         =   "Search For MP3"
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "Add MP3s"
   End
   Begin VB.Menu mnuClose 
      Caption         =   "Close"
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)


Dim PicHeight%, hLB&, filespec$, UseFileSpec%
Dim TotalDirs%, TotalFiles%, Running%

Dim WFD As WIN32_FIND_DATA, hItem&, hFile&

Const vbBackslash = "\"
Const vbAllFiles = "*.*"
Const vbKeyDot = 46
Private Sub Form_Load()
    ScaleMode = vbPixels
    PicHeight% = Picture1.Height
    hLB& = List1.hwnd
    SendMessage hLB&, LB_INITSTORAGE, 30000&, ByVal 30000& * 200
    Move (Screen.Width - Width) * 0.5, (Screen.Height - Height) * 0.5
End Sub
 
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Running% Then Running% = False
End Sub

Private Sub Form_Resize()
    MoveWindow hLB&, 0, 0, ScaleWidth, ScaleHeight - PicHeight%, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFind = Nothing
    Unload Me
End Sub

Private Sub mnuAdd_Click()
Dim i As Long
Dim tag As mp3Tag


For i = 0 To List1.ListCount - 1
Set item = Form1.ListView1.ListItems.Add(, , List1.List(i))


    If getMp3Tag(List1.List(i), tag) = True Then
        item.SubItems(1) = Trim(tag.album)
        item.SubItems(2) = Trim(tag.artist)
        item.SubItems(3) = Trim(tag.title)
        item.SubItems(4) = genreSearch(tag.genre)
        item.SubItems(5) = Trim(tag.year)
        item.SubItems(6) = Trim(tag.comment)
    End If
Next

End Sub

Private Sub mnuClose_Click()
Unload Me
End Sub

Private Sub mnuFindFiles_Click()
Dim drv1 As String
    If Running% Then: Running% = False: Exit Sub
  drv1 = InputBox("Choose Path Or Drive ..", , "C:\")
  
    Dim drvbitmask&, maxpwr%, pwr%
    On Error Resume Next
    
    filespec$ = "*.mp3"
               
    If Len(filespec$) = 0 Then Exit Sub
    
    MousePointer = 11
    Running% = True
    UseFileSpec% = True
    mnuFindFiles.Caption = "&Stop!"
    mnuFolderInfo.Enabled = False
    List1.Clear
    
    
    drvbitmask& = GetLogicalDrives()
    If drvbitmask& Then
        
        maxpwr% = Int(Log(drvbitmask&) / Log(2))
        For pwr% = 0 To 1
            If Running% And (2 ^ pwr% And drvbitmask&) Then _
                Call SearchDirs(drv1)
        Next
    End If
    
    Running% = False
    UseFileSpec% = False
    mnuFindFiles.Caption = "&Search For MP3s"
    mnuFolderInfo.Enabled = True
    MousePointer = 0

    Picture1.Cls
    Picture1.Print "Find File(s): " & List1.ListCount & " items found matching " & """" & filespec$ & """"
    Beep
    
End Sub

Private Sub mnuFolderInfo_Click()

    If Running% Then: Running% = False: Exit Sub
    
    Dim searchpath$
    On Error Resume Next

    searchpath$ = InputBox("Enter a valid explicit path:", "Folder Info", "C:\")
    If Len(searchpath$) < 2 Then Exit Sub
    If Mid$(searchpath$, 2, 1) <> ":" Then Exit Sub
    
    If Right$(searchpath$, 1) <> vbBackslash Then searchpath$ = searchpath$ & vbBackslash
    If FindClose(FindFirstFile(searchpath$ & vbAllFiles, WFD)) = False Then
        MsgBox searchpath$, vbInformation, "Path is invalid": Exit Sub
    End If

    MousePointer = 11
    Running% = True
    mnuFolderInfo.Caption = "&Stop!"
    mnuFindFiles.Enabled = False
    List1.Clear

    TotalDirs% = 0
    TotalFiles% = 0
    Call SearchDirs(searchpath$)
    
    Running% = False
    mnuFolderInfo.Caption = "&Folder Info..."
    mnuFindFiles.Enabled = True
    Picture1.Cls
    MousePointer = 0

    MsgBox "Total folders: " & vbTab & TotalDirs% & vbCrLf & _
                 "Total files: " & vbTab & TotalFiles%, , _
                 "Folder Info for: " & searchpath$
    
End Sub
 Private Sub SearchDirs(curpath$)

    Dim dirs%, dirbuf$(), i%
    
    Picture1.Cls
    Picture1.Print "Searching " & curpath$
    
    DoEvents
    If Not Running% Then Exit Sub
    
    hItem& = FindFirstFile(curpath$ & vbAllFiles, WFD)
    If hItem& <> INVALID_HANDLE_VALUE Then
        
        Do
            If (WFD.dwFileAttributes And vbDirectory) Then
                
                ' If not a  "." or ".." DOS subdir...
                If Asc(WFD.cFileName) <> vbKeyDot Then
                    ' This is executed in the mnuFindFiles_Click()
                    ' call though it isn't used...
                    TotalDirs% = TotalDirs% + 1
                    ' This is the heart of a recursive proc...
                    ' Cache the subdirs of the current dir in the 1 based array.
                    ' This proc calls itself below for each subdir cached in the array.
                    ' (re-allocating the array only once every 10 itinerations improves speed)
                    If (dirs% Mod 10) = 0 Then ReDim Preserve dirbuf$(dirs% + 10)
                    dirs% = dirs% + 1
                    dirbuf$(dirs%) = Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
                End If
            
            ' File size and attribute tests can be used here, i.e:
            ' ElseIf (WFD.dwFileAttributes And vbHidden) = False Then  'etc...
            
            ' Get a total file count for mnuFolderInfo_Click()
            ElseIf Not UseFileSpec% Then
                TotalFiles% = TotalFiles% + 1
            End If
        
        ' Get the next subdir or file
        Loop While FindNextFile(hItem&, WFD)
        
        ' Close the search handle
        Call FindClose(hItem&)
    
    End If

    ' When UseFileSpec% is set mnuFindFiles_Click(),
    ' SearchFileSpec() is called & each folder must be
    ' searched a second time.
    If UseFileSpec% Then
        ' Turning off painting speeds things quite a bit...
        ' Speed also would be vastly improved if the redrawing
        ' & scrolling were placed in a Timer event...
        SendMessage hLB&, WM_SETREDRAW, 0, 0
        Call SearchFileSpec(curpath$)
        ' Keeps the currently found items scrolled into view...
        SendMessage hLB&, WM_VSCROLL, SB_BOTTOM, 0
        SendMessage hLB&, WM_SETREDRAW, 1, 0
    End If
    
    ' Recursively call this proc & iterate through each subdir cached above.
    For i% = 1 To dirs%: SearchDirs curpath$ & dirbuf$(i%) & vbBackslash: Next i%
  
End Sub

Private Sub SearchFileSpec(curpath$)   ' curpath$ is passed w/ trailing "\"
' This procedure *only*  finds files in the
' current folder that match the FileSpec$
    
    hFile& = FindFirstFile(curpath$ & filespec$, WFD)
    If hFile& <> INVALID_HANDLE_VALUE Then
        
        Do
            ' Use DoEvents here since we're loading a ListBox and
            ' there could be hundreds of files matching the FileSpec$
            DoEvents
            If Not Running% Then Exit Sub
            
            ' The ListBox's Sorted property is initially set to False.
            ' Set it to True and see how things slow down a bit...
            SendMessage hLB&, LB_ADDSTRING, 0, _
                ByVal curpath$ & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
        
        ' Get the next file matching the FileSpec$
        Loop While FindNextFile(hFile&, WFD)
        
        ' Close the search handle
        Call FindClose(hFile&)
    
    End If

End Sub



