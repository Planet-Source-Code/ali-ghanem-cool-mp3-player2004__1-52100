VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4890
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dont Forget to leave comments And votes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2640
      TabIndex        =   10
      Top             =   3900
      Width           =   4455
   End
   Begin VB.Image Image3 
      Height          =   1920
      Left            =   120
      Picture         =   "frmSplash.frx":000C
      Top             =   1080
      Width           =   1920
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6240
      Picture         =   "frmSplash.frx":C856
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Made By Ali Ghanem"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   915
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   2010
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COOL MP3 Player"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1725
      Left            =   2160
      TabIndex        =   8
      Top             =   1080
      Width           =   3825
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   600
      Left            =   2520
      Top             =   3720
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   120
      Top             =   4440
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   6000
      Picture         =   "frmSplash.frx":D120
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ali-gn@lycos.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4560
      TabIndex        =   7
      Top             =   120
      Width           =   2505
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ghanem Soft Corporation ™"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   510
      Left            =   1395
      TabIndex        =   6
      Top             =   465
      Width           =   5910
   End
   Begin VB.Label lblLicenseTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LicenseTo All Developers On PSC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   2835
   End
   Begin VB.Label lblProductName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "COOL MP3 Player"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1725
      Left            =   2040
      TabIndex        =   4
      Top             =   1020
      Width           =   3825
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For Windows 95,98,NT,XP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   270
      TabIndex        =   3
      Top             =   3120
      Width           =   4095
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5760
      TabIndex        =   2
      Top             =   2940
      Width           =   1215
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "i made this good work to be a contest winner but i cant without your votes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   4500
      Width           =   7095
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright  2004 - 2005"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5280
      TabIndex        =   0
      Top             =   3300
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   240
      Left            =   240
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
End Sub

Private Sub Form_Paint()
'    Const cPi = 3.1415926
    Dim intLineWidth As Integer
    intLineWidth = 5
    ' 'save scale mode
    Dim intSaveScaleMode As Integer
    intSaveScaleMode = frmSplash.ScaleMode
    frmSplash.ScaleMode = 3
    Dim intScaleWidth As Integer
    Dim intScaleHeight As Integer
    intScaleWidth = frmSplash.ScaleWidth
    intScaleHeight = frmSplash.ScaleHeight
    ' 'clear form
    frmSplash.Cls
    ' 'draw white lines
    frmSplash.Line (0, intScaleHeight)-(intLineWidth, 0), &HFFFFFF, BF
    frmSplash.Line (0, intLineWidth)-(intScaleWidth, 0), &HFFFFFF, BF
    ' 'draw grey lines
    frmSplash.Line (intScaleWidth, 0)-(intScaleWidth - intLineWidth, intScaleHeight), &H808080, BF
    frmSplash.Line (intScaleWidth, intScaleHeight - intLineWidth)-(0, intScaleHeight), &H808080, BF
    ' 'draw triangles(actually circles) at corners
    Dim intCircleWidth As Integer
    intCircleWidth = Sqr(intLineWidth * intLineWidth + intLineWidth * intLineWidth)
    frmSplash.FillStyle = 0
    frmSplash.FillColor = QBColor(15)
    frmSplash.Circle (intLineWidth, intScaleHeight - intLineWidth), intCircleWidth, QBColor(15), _
    -3.1415926, -3.90953745777778 '-180 * cPi / 180, -224 * cPi / 180
    frmSplash.Circle (intScaleWidth - intLineWidth, intLineWidth), intCircleWidth, QBColor(15), _
    -0.78539815, -1.5707963 ' -45 * cPi / 180, -90 * cPi / 180
    ' 'draw black frame
    frmSplash.Line (0, intScaleHeight)-(0, 0), 0
    frmSplash.Line (0, 0)-(intScaleWidth - 1, 0), 0
    frmSplash.Line (intScaleWidth - 1, 0)-(intScaleWidth - 1, intScaleHeight - 1), 0
    frmSplash.Line (0, intScaleHeight - 1)-(intScaleWidth - 1, intScaleHeight - 1), 0
    frmSplash.ScaleMode = intSaveScaleMode
End Sub

Private Sub Image1_Click()
ShellExecute hwnd, "open", "mailto:ali-gn@lycos.com", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub imgLogo_Click()
 
End Sub

Private Sub Image3_Click()
Unload Me

End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub lblCompanyProduct_Click()
Unload Me
End Sub

Private Sub lblProductName_Click()
Unload Me
End Sub

