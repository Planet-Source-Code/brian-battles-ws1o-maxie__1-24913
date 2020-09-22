Attribute VB_Name = "modMaximizeIE"
' Module     : modMaximizeIE
' Description: makes API call find and active instances of Internet Explorer,
'               then maximize or minimize all of them
' Procedures : fEnumWindows()
'              fEnumWindowsCallBack(ByVal hWnd As Long, ByVal lpData As Long)
'              MaximizeInternetExplorer()
'
' Modified   : 7/11/2001 by B Battles

Option Explicit

Public Const MAX_PATH = 260
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6

Private bieMinimize  As Boolean

Public Type POINTAPI
    X                As Long
    Y                As Long
End Type

Public Type RECT
    Left             As Long
    Top              As Long
    Right            As Long
    Bottom           As Long
End Type

Public Type WINDOWPLACEMENT
    Length           As Long
    Flags            As Long
    ShowCmd          As Long
    ptMinPosition    As POINTAPI
    ptMaxPosition    As POINTAPI
    rcNormalPosition As RECT
End Type

Private Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpWndPl As WINDOWPLACEMENT) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Function fEnumWindows(Optional ieMinimize As Boolean = False) As Boolean
    
    Dim hWnd As Long
    
    ' The EnumWindows function enumerates all top-level windows
    ' on the screen by passing the handle of each window, in turn,
    ' to an application-defined callback function. EnumWindows
    ' continues until the last top-level window is enumerated or
    ' the callback function returns FALSE
    
    On Error GoTo Err_fEnumWindows
    
    ' if minimized was passed in as a parameter, set the flag to True
    bieMinimize = ieMinimize
    Call EnumWindows(AddressOf fEnumWindowsCallBack, hWnd)
    
Exit_fEnumWindows:
    
    On Error GoTo 0
    Exit Function
    
Err_fEnumWindows:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modMaximizeIE" & " during " & "fEnumWindows" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_fEnumWindows
    End Select
    
End Function
Public Function fEnumWindowsCallBack(ByVal hWnd As Long, ByVal lpData As Long) As Long
    
    Dim lResult    As Long
    Dim sWndName   As String
    Dim sClassName As String
    Dim WinEst     As WINDOWPLACEMENT
    Dim Rtn        As Long
    
    ' This callback function is called by Windows (from the EnumWindows
    ' API call) for EVERY window that exists. It populates the aWindowList
    ' array with a list of windows we are interested in.
    
    On Error GoTo Err_fEnumWindowsCallBack
    
    fEnumWindowsCallBack = 1
    sClassName = Space$(MAX_PATH)
    sWndName = Space$(MAX_PATH)
    lResult = GetClassName(hWnd, sClassName, MAX_PATH)
    sClassName = Left$(sClassName, lResult)
    ' is this an Internet Explorer window?
    If InStr(sClassName, "IEFrame") Then
        'initialize the structure
        WinEst.Length = Len(WinEst)
            If bieMinimize Then
                WinEst.ShowCmd = SW_MINIMIZE
            Else
                WinEst.ShowCmd = SW_MAXIMIZE
            End If
        'set the new window placement (minimized or maximized)
        Rtn = SetWindowPlacement(hWnd, WinEst)
    End If
    lResult = GetWindowText(hWnd, sWndName, MAX_PATH)
    
Exit_fEnumWindowsCallBack:
    
    On Error GoTo 0
    Exit Function
    
Err_fEnumWindowsCallBack:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In modMaximizeIE" & " during " & "fEnumWindowsCallBack" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_fEnumWindowsCallBack
    End Select
    
End Function
Sub MaximizeInternetExplorer(Optional b_Minimize As Boolean = False)
    
    On Error GoTo Err_MaximizeInternetExplorer
    
    ' just run the routines that enumerate
    ' the windows and minimizes or maximizes IE
    If b_Minimize Then
        fEnumWindows b_Minimize
    Else
        fEnumWindows
    End If
    'End
    
Exit_MaximizeInternetExplorer:
    
    On Error GoTo 0
    Exit Sub
    
Err_MaximizeInternetExplorer:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In Module1" & " during " & "MaximizeInternetExplorer" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_MaximizeInternetExplorer
    End Select
    
End Sub
