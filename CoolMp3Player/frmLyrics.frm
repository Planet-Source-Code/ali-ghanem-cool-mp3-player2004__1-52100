VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLyrics 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lyrics : Here You Can Save Songs Words and Sing It ...."
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLyrics.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   5910
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   6585
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   2469
            MinWidth        =   2469
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "20/02/2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
            Key             =   "sbRecNum"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      DataField       =   "Songname"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   4
      Top             =   960
      Width           =   3855
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   1429
      ButtonWidth     =   1032
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "tbNew"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            Key             =   "tbFind"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "tbDelete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "First"
            Key             =   "tbFirst"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "tbBack"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Next"
            Key             =   "tbNext"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Last"
            Key             =   "tbLast"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fonts"
            Key             =   "tbFonts"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   8
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbBold"
                  Text            =   "Bold"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbItalic"
                  Text            =   "Italic"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbUnderline"
                  Text            =   "Underline"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbStrikethru"
                  Text            =   "Strikethru"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbLeft"
                  Text            =   "Align To Left"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbCenter"
                  Text            =   "Align To Center"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbRight"
                  Text            =   "Align To Right"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Color"
            Key             =   "tbColor"
            ImageIndex      =   9
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbRed"
                  Text            =   "Red"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbBlue"
                  Text            =   "Blue"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbBlack"
                  Text            =   "Black"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbYellow"
                  Text            =   "Yellow"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbGreen"
                  Text            =   "Green"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLyrics.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLyrics.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLyrics.frx":14BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLyrics.frx":1D98
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLyrics.frx":20B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLyrics.frx":298C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLyrics.frx":3266
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLyrics.frx":3580
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLyrics.frx":3E5A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RText1 
      DataField       =   "Lyrics"
      DataSource      =   "Data1"
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8281
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmLyrics.frx":4734
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DataField       =   "Num"
      DataSource      =   "Data1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Song Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lyrics"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   930
   End
End
Attribute VB_Name = "frmLyrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Data1_Reposition()
Label3.Caption = Data1.Recordset.AbsolutePosition + 1
StatusBar1.Panels(3).Text = "Record Number  :" & "  " & Label3.Caption
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path + "\Songs.mdb"
Data1.RecordSource = "Table1"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key

Case "tbColor"
Form1.CD.ShowColor
RText1.SelColor = Form1.CD.Color

Case "tbFirst"
FirstRecord
Case "tbLast"
LastRecord
Case "tbNext"
NextRecord
Case "tbBack"
PrevRecord

Case "tbNew"
AddRecord
Case "tbDelete"
DelRecord
Case "tbFind"
FindRecNum

End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key

Case "tbBold"
  If RText1.SelBold = True Then
  RText1.SelBold = False
  Else
  RText1.SelBold = True
  End If
  
Case "tbItalic"
  If RText1.SelItalic = True Then
  RText1.SelItalic = False
  Else
  RText1.SelItalic = True
  End If
  
Case "tbUnderline"
  If RText1.SelUnderline = True Then
  RText1.SelUnderline = False
  Else
  RText1.SelUnderline = True
  End If
  
Case "tbStrikethru"
  If RText1.SelStrikeThru = True Then
  RText1.SelStrikeThru = False
  Else
  RText1.SelStrikeThru = True
  End If
  
Case "tbRed"
RText1.SelColor = vbRed

Case "tbBlue"
RText1.SelColor = vbBlue

Case "tbGreen"
RText1.SelColor = vbGreen

Case "tbYellow"
RText1.SelColor = vbYellow

Case "tbBlack"
RText1.SelColor = vbBlack

Case "tbLeft"
RText1.SelAlignment = 0
Case "tbCenter"
RText1.SelAlignment = 2
Case "tbRight"
RText1.SelAlignment = 1
End Select

End Sub


Sub AddRecord()
On Error Resume Next
Data1.Recordset.AddNew
End Sub

Sub DelRecord()
On Error Resume Next
msg1 = MsgBox("Are You Sure To Delete ?", vbYesNo + vbExclamation, "Delete")
If Data1.Recordset.RecordCount = 1 Then
MsgBox "You Can Not Delete The First Record", vbCritical
Exit Sub
End If
If msg1 = vbYes Then
Data1.Recordset.Delete
Data1.Recordset.MovePrevious
End If
End Sub

Sub NextRecord()
On Error Resume Next
With Data1
   If Not .Recordset.EOF Then .Recordset.MoveNext
   If .Recordset.EOF And .Recordset.RecordCount > 0 Then
   .Refresh
   .Recordset.MoveFirst
   End If
End With
End Sub

Sub PrevRecord()
On Error Resume Next
With Data1
  If Not .Recordset.BOF Then .Recordset.MovePrevious
  If .Recordset.BOF And .Recordset.RecordCount > 0 Then
    .Refresh
    .Recordset.MoveLast
  End If
End With
End Sub

Sub FirstRecord()
Data1.Recordset.MoveFirst
End Sub

Sub LastRecord()
Data1.Recordset.MoveLast
End Sub

Sub FindRecNum()
Dim varName7 As Variant
Dim strBkMark7 As String
varName7 = InputBox("Enter Record Number Please ...", "Find Author", "1")
If varName7 = "" Then
Exit Sub
Else
varName7 = "'" & varName7 & "'" ' fix for string
End If
With Me.Data1.Recordset
strBkMark7 = .Bookmark
.FindFirst "Num LIKE " & varName7
If .NoMatch Then
.Bookmark = strBkMark7
MsgBox "Unable to find Num LIKE [" & varName7 & "]", vbCritical, "Find Error"
End If
End With
End Sub
