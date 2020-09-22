VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Slider Slider3 
      Height          =   390
      Left            =   240
      TabIndex        =   12
      Top             =   3600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   688
      _Version        =   393216
      BorderStyle     =   1
      Min             =   -4000
      Max             =   1
      TickStyle       =   3
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2640
      Top             =   2400
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   3600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Max             =   2
      SelectRange     =   -1  'True
      TickStyle       =   3
      Value           =   1
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Volume Bar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   480
      TabIndex        =   13
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mute Sound"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   4680
      TabIndex        =   11
      Top             =   2040
      Width           =   1380
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Position Bar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   285
      Left            =   3720
      TabIndex        =   10
      Top             =   3240
      Width           =   1410
   End
   Begin VB.Image Image3 
      Height          =   720
      Left            =   5160
      Picture         =   "frmOptions.frx":08CA
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Position"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3120
      TabIndex        =   6
      Top             =   2520
      Width           =   1875
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Song Duration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   240
      Width           =   1665
   End
   Begin VB.Image Image2 
      Height          =   1920
      Left            =   2640
      Picture         =   "frmOptions.frx":114A
      Top             =   600
      Width           =   1920
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   5280
      TabIndex        =   8
      Top             =   2400
      Width           =   1125
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   5280
      TabIndex        =   7
      Top             =   120
      Width           =   1125
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   3975
      Left            =   2520
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Right"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1800
      TabIndex        =   4
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Center"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   870
      TabIndex        =   3
      Top             =   2640
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sound Balance"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   480
      Picture         =   "frmOptions.frx":2F54
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1440
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   3975
      Left            =   120
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Const V3 = -10000
Const V2 = 10000
Const V1 = 0

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
On Error Resume Next
Label7.Caption = Form1.SecToMin(Form1.AM.Duration)
Slider2.Min = 1
Slider2.Max = Form1.AM.Duration
End Sub

Private Sub Form_Load()
On Error Resume Next
Slider1.Value = 1
Form1.AM.Balance = 1
Me.Left = Form1.Width - Me.Width - 200
Me.Top = Form1.Height - Me.Height - 200
End Sub

Private Sub Image3_Click()
frmProperties.Show
frmProperties.Command1.Enabled = False
End Sub

Private Sub Slider1_Scroll()
On Error Resume Next
Select Case Slider1.Value

Case 0
Form1.AM.Balance = V2

Case 1
Form1.AM.Balance = V1

Case 2
Form1.AM.Balance = V3

End Select
End Sub
Private Sub Slider2_Scroll()
Form1.AM.CurrentPosition = Slider2.Value
End Sub

Private Sub Slider3_Scroll()
On Error Resume Next
Form1.AM.Volume = Slider3.Value
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Slider2.Value = Form1.AM.CurrentPosition
End Sub
