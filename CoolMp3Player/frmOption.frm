VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOption.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "MP3 Info"
      Height          =   4215
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   4095
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   23
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   22
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   21
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1200
         MaxLength       =   40
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   19
         Top             =   2280
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         TabIndex        =   18
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Title :"
         Height          =   195
         Left            =   480
         TabIndex        =   29
         Top             =   480
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Artist :"
         Height          =   195
         Left            =   360
         TabIndex        =   28
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Album :"
         Height          =   195
         Left            =   360
         TabIndex        =   27
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Year :"
         Height          =   195
         Left            =   480
         TabIndex        =   26
         Top             =   2280
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Comment :"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Genre :"
         Height          =   195
         Left            =   360
         TabIndex        =   24
         Top             =   3600
         Width           =   540
      End
   End
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Archive"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7080
      TabIndex        =   16
      Top             =   3840
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Hidden"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5880
      TabIndex        =   15
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Caption         =   "Properties"
      Height          =   4215
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   3855
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Read Only"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Size :"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   390
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Path :"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Type :"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Accessed :"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   3120
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Created :"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   2640
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Modified :"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   2160
         Width           =   705
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update MP3 Info"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1935
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo SaveProblem
Dim tag As mp3Tag

tag.title = Trim(Text1.Text)
tag.artist = Trim(Text2.Text)
tag.album = Trim(Text3.Text)
tag.year = Trim(Text4.Text)
tag.comment = Trim(Text5.Text)
tag.genre = getGenreTagCode(Combo1.Text)
tag.tagID = Text1.tag
If putMp3Tag(Form1.Caption, tag) Then
End If
MsgBox "New MP3 Info Saved ....  ", vbInformation, "MP3 Info"
Exit Sub
SaveProblem:
MsgBox "You Can Not Save in Read Only File..", vbCritical
Exit Sub
End Sub



Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    populateGenreList
    
    For i = 0 To 125
        Combo1.AddItem (genreList(i).genre)
    Next i
    
   Command1.Picture = Me.Icon
   Command2.Picture = Form1.ImageList1.ListImages(19).Picture
On Error Resume Next
LoadInfo
End Sub

Sub LoadInfo()
    Dim endOfFile As Long
    Dim tag As mp3Tag
           
    If getMp3Tag(Trim(Form1.Caption), tag) Then
        Text1.tag = "TAG"
        Text1.Text = Trim(tag.title)
        Text2.Text = Trim(tag.artist)
        Text3.Text = Trim(tag.album)
        Text4.Text = Trim(tag.year)
        Text5.Text = Trim(tag.comment)
        Combo1.Text = genreSearch(tag.genre)
    Else
        Text1.tag = ""
        Text1.Text = "No Tag Found"
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Combo1.Text = ""
    End If
Dim fso As New FileSystemObject
Dim fo As File

Set fo = fso.GetFile(Form1.Caption)
Text6.Text = fo.DateLastModified
Text7.Text = fo.DateCreated
Text8.Text = fo.DateLastAccessed
Text9.Text = fo.Type
Text10.Text = fo.Path
Text12.Text = fo.Size

art = fo.Attributes
Select Case art
Case 32
Check3.Value = 1
Check1.Value = 0
Check2.Value = 0
Case 1
Check1.Value = 1
Check2.Value = 0
Check3.Value = 0
Case 2
Check2.Value = 2
Case 33
Check1.Value = 1
Check3.Value = 1
Case 34
Check1.Value = 1
Check2.Value = 1
Case 3
Check1.Value = 1
Check2.Value = 1
Case 35
Check1.Value = 1
Check2.Value = 1
Check3.Value = 1

End Select
End Sub

