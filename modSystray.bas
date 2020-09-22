Attribute VB_Name = "modSystray"
' Description: Routines to place this little app in the System Tray
' Modified   : 7/11/2001 by B Battles

Option Explicit

'user defined type required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
    cbSize           As Long
    hWnd             As Long
    uId              As Long
    uFlags           As Long
    uCallBackMessage As Long
    hIcon            As Long
    szTip            As String * 64
End Type

'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

Public nID           As NOTIFYICONDATA

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
