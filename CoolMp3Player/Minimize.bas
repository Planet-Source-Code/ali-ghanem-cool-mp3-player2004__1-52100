Attribute VB_Name = "Minimize"
Public Declare Function shell_notifyicon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_ICON = &H2
Public Const WM_COMMNOTIFY = &H44
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const MAX_TOOLTIP As Integer = 64

'Êã ÊÚÑíÝ äãØ ÎÇÕ
Type NOTIFYICONDATA
    cbsize As Long
    hwind As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * MAX_TOOLTIP
End Type

Public TryIcon As NOTIFYICONDATA
