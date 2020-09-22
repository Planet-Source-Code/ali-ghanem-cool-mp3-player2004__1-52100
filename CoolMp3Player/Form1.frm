VERSION 5.00
Object = "{05589FA0-C356-11CE-BF01-00AA0055595A}#2.0#0"; "AMOVIE.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Cool MP3 Player      -      Created By Ali Ghanem         -      2004"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Height          =   810
      Left            =   14950
      Picture         =   "Form1.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   30
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   3840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   2400
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   3960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   5400
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2655
   End
   Begin AMovieCtl.ActiveMovie AM 
      Height          =   420
      Left            =   2400
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   741
      ShowDisplay     =   0   'False
      AutoStart       =   -1  'True
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   6120
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   0
      TabIndex        =   2
      Top             =   810
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "lstFileName"
         Text            =   "File Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "lstAlbum"
         Text            =   "Album"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "lstArtist"
         Text            =   "Artist"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "lstTitle"
         Text            =   "Title"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "lstGenre"
         Text            =   "Genre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "lstYear"
         Text            =   "Year"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "lstComment"
         Text            =   "Comment"
         Object.Width           =   4762
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1429
      ButtonWidth     =   1482
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   23
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Files"
            Key             =   "tbAddFiles"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Dir"
            Key             =   "tbAddDir"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Play"
            Key             =   "tbPlay"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "tbStop"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pause"
            Key             =   "tbPause"
            ImageIndex      =   21
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "tbBack"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Next"
            Key             =   "tbNext"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Option"
            Key             =   "tbOption"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Volume"
            Key             =   "tbVolume"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save List"
            Key             =   "tbSaveList"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Load List"
            Key             =   "tbLoadList"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear List"
            Key             =   "tbClearList"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Properties"
            Key             =   "tbProperties"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lyrics"
            Key             =   "tbLyrics"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "tbFindMP3"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "tbAbout"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "tbExit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8400
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0922
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1AD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":23B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3564
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4718
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4FF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":58CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":64C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6D9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7674
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7F4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8828
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9102
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":99DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A2B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A5D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AEAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B1C4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SW_MINIMIZE = 6
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long



Private Sub AM_Timer()
Text1.Text = CInt(AM.CurrentPosition)
Text2.Text = CInt(AM.Duration)
Text3.Text = SecToMin(Text1.Text)
frmOptions.Label8.Caption = SecToMin(Text1.Text)
If CInt(AM.CurrentPosition) + 0.1 >= CInt(AM.Duration) Then
AM.Stop
NextSong2
End If
End Sub

Private Sub Command1_Click()
SetWindowPos Me.hwnd, 0, 0, 0, 0, 0, 6
End Sub





Private Sub Form_Load()
populateGenreList
AM.Left = -6000
frmSplash.Show 0, Form1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case (x \ Screen.TwipsPerPixelX)
   
    Case &H203 '   Left Button
    Case &H201 '   Right Button
      Form2.Show
    End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
ListView1.Width = Form1.ScaleWidth
ListView1.Height = Form1.ScaleHeight - 800
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call shell_notifyicon(NIM_DELETE, TryIcon)
End
End Sub
Private Sub ListView1_DblClick()
On Error Resume Next
AM.FileName = Form1.Caption
Toolbar1.Buttons(4).Value = tbrPressed
ListView1.SelectedItem.ForeColor = vbRed
KO = ListView1.SelectedItem.Index
Text4.Text = KO
End Sub

Private Sub ListView1_ItemClick(ByVal item As MSComctlLib.ListItem)
Form1.Caption = item.Text
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
If KeyCode = vbKeyDelete Then
    ListView1.ListItems.Remove ListView1.SelectedItem.Index
End If
End Sub
Private Sub Slider1_Scroll()
AM.Volume = Slider1.Value
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key

Case "tbAbout"
frmSplash.Show

Case "tbExit"
msg = MsgBox("Are You Sure You Want To Leave ?", vbOKCancel + vbExclamation, "Cool MP3 Player")
If msg = vbOK Then
End
End If

Case "tbAddDir"
AddDir
On Error Resume Next
Form1.Caption = ListView1.ListItems(1).Text

Case "tbAddFiles"
AddFiles
On Error Resume Next
Form1.Caption = ListView1.ListItems(1).Text

Case "tbStop"
On Error Resume Next
AM.Stop
Toolbar1.Buttons(4).Value = tbrUnpressed
Toolbar1.Buttons(6).Value = tbrUnpressed

Case "tbPause"
On Error Resume Next
AM.Pause
Toolbar1.Buttons(4).Value = tbrUnpressed
If Toolbar1.Buttons(6).Value = tbrPressed Then
Toolbar1.Buttons(6).Value = tbrPressed
Else
Toolbar1.Buttons(6).Value = tbrPressed
End If

Case "tbPlay"
On Error Resume Next
AM.Run
Toolbar1.Buttons(6).Value = tbrUnpressed
If Toolbar1.Buttons(4).Value = tbrPressed Then
Toolbar1.Buttons(4).Value = tbrPressed
Else
Toolbar1.Buttons(4).Value = tbrPressed
End If

Case "tbNext"
On Error Resume Next
AM.CurrentPosition = AM.CurrentPosition + 5

Case "tbBack"
On Error Resume Next
AM.CurrentPosition = AM.CurrentPosition - 5
'**************************************************

Case "tbOption"
frmOptions.Show 0, Form1

Case "tbProperties"
On Error Resume Next
AM.Stop
AM.FileName = ""
Toolbar1.Buttons(4).Value = tbrUnpressed
frmProperties.Show

Case "tbVolume"
Dim S1 As String
S1 = Space(50)
GetWindowsDirectory S1, Len(S1)

s2 = RTrim(Left(S1, 10)) & "\" & "system32\sndvol32.exe"
Shell s2, vbNormalFocus

Case "tbLoadList"
LoadList

Case "tbSaveList"
SaveList

Case "tbClearList"
re = MsgBox("Are You Sure", vbYesNo + vbInformation, "Cool Mp3 Info")
If re = vbYes Then
ListView1.ListItems.Clear
End If

Case "tbLyrics"
frmLyrics.Show 0, Form1

Case "tbFindMP3"
frmFind.Show 1, Form1


End Select
End Sub
Function GetFiles(filespec As String, Optional Attributes As VbFileAttribute)
On Error Resume Next
fl = Dir(filespec & "*.mp3", Attributes)
Do While Len(fl)
Set item = ListView1.ListItems.Add(, , filespec & fl)
st = a & fl

Dim tag As mp3Tag
Dim endOfFile As Long

    If getMp3Tag(st, tag) Then
        item.SubItems(1) = Trim(tag.album)
        item.SubItems(2) = Trim(tag.artist)
        item.SubItems(3) = Trim(tag.title)
        item.SubItems(4) = genreSearch(tag.genre)
        item.SubItems(5) = Trim(tag.year)
        item.SubItems(6) = Trim(tag.comment)
        
   Else
  End If
   
    fl = Dir
    Loop

End Function


Sub AddFiles()
On Error GoTo FullText
Dim S1
CD.FileName = ""

  Dim a As Variant
    Dim P As Integer
   
    
    CD.Flags = cdlOFNExplorer + cdlOFNAllowMultiselect
    CD.Filter = "MP3 Files (*.mp3)|*.mp3| Wave Files (*.wav)|*.wav|"
    
    CD.InitDir = "C:\Documents and Settings\Doraid Ghanem\My Documents\My Music\Radio\"
    CD.MaxFileSize = 5000
    CD.ShowOpen
    
   If CD.FileName = "" Then Exit Sub
    
   
    a = Split(CD.FileName, vbNullChar)
     
 
     

    For P = 1 To UBound(a)

    Dim tag As mp3Tag
    
    If InStr(1, a(0) + "\" + a(P), "\\", vbTextCompare) <> 0 Then
    Set item = ListView1.ListItems.Add(, , a(0) + a(P))
    


    If getMp3Tag(a(0) + a(P), tag) Then
        item.SubItems(1) = Trim(tag.album)
        item.SubItems(2) = Trim(tag.artist)
        item.SubItems(3) = Trim(tag.title)
        item.SubItems(4) = genreSearch(tag.genre)
        item.SubItems(5) = Trim(tag.year)
        item.SubItems(6) = Trim(tag.comment)
    End If
    
    
    Else
    
    Set item = ListView1.ListItems.Add(, , a(0) + "\" + a(P))
    If getMp3Tag(a(0) + "\" + a(P), tag) Then
        item.SubItems(1) = Trim(tag.album)
        item.SubItems(2) = Trim(tag.artist)
        item.SubItems(3) = Trim(tag.title)
        item.SubItems(4) = genreSearch(tag.genre)
        item.SubItems(5) = Trim(tag.year)
        item.SubItems(6) = Trim(tag.comment)
    End If
    
    End If
    
    
     
    Next
    On Error Resume Next
    
If P = 1 Then
Set item = ListView1.ListItems.Add(, , CD.FileName)

    If getMp3Tag(CD.FileName, tag) Then
        item.SubItems(1) = Trim(tag.album)
        item.SubItems(2) = Trim(tag.artist)
        item.SubItems(3) = Trim(tag.title)
        item.SubItems(4) = genreSearch(tag.genre)
        item.SubItems(5) = Trim(tag.year)
        item.SubItems(6) = Trim(tag.comment)
    End If

End If

Exit Sub
FullText:
MsgBox "You Can Select 15 Files Each Once", vbCritical
End Sub

Sub AddDir()
a = BrowseForFolder("C:\Documents and Settings\Doraid Ghanem\My Documents\My Music\Radio\", Me.hwnd, "&Select a directory:")
If Right$(a, 1) <> "\" Then
a = a & "\"
End If
If a = "\" Then Exit Sub
GetFiles a, vbNormal
End Sub


Sub SaveList()
If ListView1.ListItems.Count = 0 Then
MsgBox "You Have To Choose Songs Before ....", vbCritical
Exit Sub
End If

Dim F
Dim n As Long
F = FreeFile
CD.FileName = ""
CD.Filter = "Songs List (*.lst)|*.lst|"
CD.ShowSave
If CD.FileName = "" Then Exit Sub
Open CD.FileName For Output As #F
  
  For n = 1 To ListView1.ListItems.Count

  Print #F, ListView1.ListItems(n).Text
  
  Next
  
Close #F
MsgBox "Songs List Was Saved ...", vbInformation
End Sub

Sub LoadList()
On Error Resume Next
Dim F
Dim S1 As String
Dim tag As mp3Tag
Dim endOfFile As Long


F = FreeFile
CD.FileName = ""
CD.Filter = "Songs List (*.lst)|*.lst|"
CD.ShowOpen

If CD.FileName = "" Then Exit Sub

Open CD.FileName For Input As #F

Do While Not EOF(F)
Input #F, S1
Set item = ListView1.ListItems.Add(, , S1)



Loop
    
Close #F

'###################Load Info####################'
Static n As Integer

Do Until n = ListView1.ListItems.Count
n = n + 1
Set item = ListView1.ListItems(n)


    If getMp3Tag(Trim(ListView1.ListItems(n).Text), tag) Then
    
        item.SubItems(1) = Trim(tag.album)
        item.SubItems(2) = Trim(tag.artist)
        item.SubItems(3) = Trim(tag.title)
        item.SubItems(4) = genreSearch(tag.genre)
        item.SubItems(5) = Trim(tag.year)
        item.SubItems(6) = Trim(tag.comment)
  
    End If
Loop


End Sub

Sub NextSong()
If ListView1.SelectedItem.Index = ListView1.ListItems.Count Then
ListView1.ListItems(1).Selected = True
ListView1.ListItems(1).ForeColor = vbRed
Form1.Caption = ListView1.ListItems(1).Text
AM.Stop
AM.FileName = Form1.Caption
AM.Run
Toolbar1.Buttons(4).Value = tbrPressed
Else
ListView1.ListItems.item(ListView1.SelectedItem.Index + 1).Selected = True

Form1.Caption = ListView1.ListItems.item(ListView1.SelectedItem.Index).Text

ListView1.ListItems.item(ListView1.SelectedItem.Index).ForeColor = vbRed
AM.Stop
AM.FileName = ListView1.SelectedItem.Text
AM.Run
Toolbar1.Buttons(4).Value = tbrPressed

End If
End Sub

Public Function SecToMin(vSeconds As Long) As String
If vSeconds < 60 Then
    SecToMin = "0:" & Format(vSeconds, "00")
    Exit Function
End If

If (vSeconds Mod 60) = 0 Then
    vSeconds = vSeconds / 60
    SecToMin = vSeconds & ":00"
    Exit Function
End If


Dim vData As Long

vData = (vSeconds Mod 60)

vSeconds = vSeconds - vData

vSeconds = vSeconds / 60

SecToMin = vSeconds & ":" & Format(vData, "00")
End Function

Sub NextSong2()
If KO = ListView1.ListItems.Count Then
KO = 1
Text4.Text = KO
ListView1.ListItems(1).Selected = True
ListView1.ListItems(1).ForeColor = vbRed
Form1.Caption = ListView1.ListItems(1).Text
'AM.Stop
AM.FileName = ListView1.ListItems(1).Text
AM.Run
Toolbar1.Buttons(4).Value = tbrPressed
Else
KO = KO + 1
Text4.Text = KO
ListView1.ListItems.item(KO).Selected = True

Form1.Caption = ListView1.ListItems.item(KO).Text

ListView1.ListItems.item(KO).ForeColor = vbRed
AM.FileName = ListView1.ListItems(KO).Text
AM.Run
Toolbar1.Buttons(4).Value = tbrPressed

End If
End Sub

