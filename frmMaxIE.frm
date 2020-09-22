VERSION 5.00
Begin VB.Form frmMaxIE 
   Caption         =   "                 Maximize All IE Windows"
   ClientHeight    =   975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5565
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaxIE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Brian Battles WS1O 2001"
      Height          =   210
      Left            =   1575
      TabIndex        =   1
      Top             =   735
      Width           =   2115
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maximize All Internet Explorer Windows Loads To Systray At Startup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   5535
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuMaxIE 
      Caption         =   "&Maximize IE"
   End
   Begin VB.Menu mnuMinIE 
      Caption         =   "Mi&nimize IE"
   End
   Begin VB.Menu mnupopuptray 
      Caption         =   "&Systray"
      Begin VB.Menu mnupopupMaxIE 
         Caption         =   "&MaximizeIE"
      End
      Begin VB.Menu mnupopupMinIE 
         Caption         =   "Mi&nimize IE"
      End
      Begin VB.Menu mnupopuprestore 
         Caption         =   "&Restore Me"
      End
      Begin VB.Menu mnupopmin 
         Caption         =   "M&inimize Me"
      End
      Begin VB.Menu mnupopupSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnupopupexit 
         Caption         =   "E&xit (Unload)  Me"
      End
   End
End
Attribute VB_Name = "frmMaxIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Description: Parks in the System Tray unless clicked by user
' Procedures : Form_Load()
'              Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'              Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'              Form_Resize()
'              Form_Unload(Cancel As Integer)
'              HideTheForm()
'              mnuMaxIE_Click()
'              mnupopmin_Click()
'              mnupopupexit_Click()
'              mnupopupMaxIE_Click()
'              mnupopuprestore_Click()
'              mPopExit_Click()
'
' Modified   : 7/11/2001 by B Battles

Option Explicit

Dim Result As Long
Dim Msg    As Long
Private Sub Form_Load()
        
    On Error GoTo Err_Form_Load
    
    HideTheForm
    
Exit_Form_Load:
    
    Exit Sub
    
Err_Form_Load:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during Form_Load, in frmLoadToSystray", vbInformation, "Advisory"
            Resume Exit_Form_Load
    End Select
    
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'this procedure receives the
    'callbacks from the System Tray icon
    
    On Error GoTo Err_Form_MouseMove
    
    Dim Result As Long
    Dim Msg    As Long
    
    'the value of X will vary depending on the scalemode setting
    If Me.ScaleMode = vbPixels Then
        Msg = X
    Else
        Msg = X / Screen.TwipsPerPixelX
    End If
    Select Case Msg
        Case WM_LBUTTONUP        '514 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hWnd)
            Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hWnd)
            Me.Show
        Case WM_RBUTTONUP        '517 display popup menu
            Result = SetForegroundWindow(Me.hWnd)
            Me.PopupMenu Me.mnupopuptray
    End Select
    
Exit_Form_MouseMove:
    
    Exit Sub
    
Err_Form_MouseMove:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during Form_MouseMove, in frmLoadToSystray", vbInformation, "Advisory"
            Resume Exit_Form_MouseMove
    End Select
    
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Err_Form_MouseUp
    
    If Me.WindowState = vbMinimized Then
        Me.Visible = False
        Exit Sub
    End If
    If Me.WindowState <> vbMinimized Then
        Me.Visible = True
        Exit Sub
    End If
    
Exit_Form_MouseUp:
    
    Exit Sub
    
Err_Form_MouseUp:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during Form_MouseUp, in frmLoadToSystray", vbInformation, "Advisory"
            Resume Exit_Form_MouseUp
    End Select
    
End Sub
Private Sub Form_Resize()
    
    ' assures that the minimized window Is Hidden
    
    On Error GoTo Err_Form_Resize
    
    If Me.WindowState = vbMinimized Then
        Me.Hide
    End If
    
Exit_Form_Resize:
    
    Exit Sub
    
Err_Form_Resize:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during Form_Resize, in frmLoadToSystray", vbInformation, "Advisory"
            Resume Exit_Form_Resize
    End Select
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    ' removes the icon from the systray
    
    On Error GoTo Err_Form_Unload
    
    Shell_NotifyIcon NIM_DELETE, nID
        
Exit_Form_Unload:
    
    End
    Exit Sub
    
Err_Form_Unload:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during Form_Unload, in frmLoadToSystray", vbInformation, "Advisory"
            Resume Exit_Form_Unload
    End Select
    
End Sub
Private Sub HideTheForm()
    
    On Error GoTo Err_HideTheForm
    
    Me.Show
    Me.Refresh
    With nID
        .cbSize = Len(nID)
        .hWnd = Me.hWnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Click for More Info..." & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nID
    ' for a tray start up
    Me.Visible = False
    
Exit_HideTheForm:
    
    Exit Sub
    
Err_HideTheForm:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during HideTheForm, in frmLoadToSystray", vbInformation, "Advisory"
            Resume Exit_HideTheForm
    End Select
    
End Sub
Private Sub mnuMaxIE_Click()
    
    On Error GoTo Err_mnuMaxIE_Click
    
    MaximizeInternetExplorer
    mnupopmin_Click
    
Exit_mnuMaxIE_Click:
    
    Exit Sub
    
Err_mnuMaxIE_Click:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during mnuMaxIE_Click, in frmLoadToSystray", vbInformation, "Advisory"
            Resume Exit_mnuMaxIE_Click
    End Select
    
End Sub
Private Sub mnuMinIE_Click()
   
    On Error GoTo Err_mnuMinIE_Click

    MaximizeInternetExplorer True
    mnupopmin_Click

Exit_mnuMinIE_Click:
    
    On Error GoTo 0
    Exit Sub

Err_mnuMinIE_Click:
     
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmMaxIE" & " during " & "mnuMinIE_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_mnuMinIE_Click
    End Select

End Sub
Private Sub mnupopmin_Click()
    
    'minimize
    
    On Error GoTo Err_mnupopmin_Click
    
    If Me.WindowState <> vbMinimized Then
        Me.Hide
    End If
    
Exit_mnupopmin_Click:
    
    Exit Sub
    
Err_mnupopmin_Click:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during mnupopmin_Click, in frmLoadToSystray", vbInformation, "Advisory"
            Resume Exit_mnupopmin_Click
    End Select
    
End Sub
Private Sub mnupopupexit_Click()
    
    ' called when user clicks the popup menu Exit command
    
    On Error GoTo Err_mnupopupexit_Click
    
    Unload Me
    
Exit_mnupopupexit_Click:
    
    Exit Sub
    
Err_mnupopupexit_Click:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during mnupopupexit_Click, in frmLoadToSystray", vbInformation, "Advisory"
            Resume Exit_mnupopupexit_Click
    End Select
    
End Sub
Private Sub mnupopupMaxIE_Click()
    
    On Error GoTo Err_mnupopupMaxIE_Click
    
    MaximizeInternetExplorer
    mnupopmin_Click
    
Exit_mnupopupMaxIE_Click:
    
    Exit Sub
    
Err_mnupopupMaxIE_Click:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during mnupopupMaxIE_Click, in frmLoadToSystray", vbInformation, "Advisory"
            Resume Exit_mnupopupMaxIE_Click
    End Select
    
End Sub
Private Sub mnupopupMinIE_Click()

    On Error GoTo Err_mnupopupMinIE_Click

    MaximizeInternetExplorer True
    mnupopmin_Click

Exit_mnupopupMinIE_Click:
    
    On Error GoTo 0
    Exit Sub

Err_mnupopupMinIE_Click:
     
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmMaxIE" & " during " & "mnupopupMinIE_Click" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_mnupopupMinIE_Click
    End Select

End Sub
Private Sub mnupopuprestore_Click()
    'called when the user clicks
    'the popup menu Restore command
    
    On Error GoTo Err_mnupopuprestore_Click
    
    Dim Result As Long
    
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hWnd)
    Me.Show
    
Exit_mnupopuprestore_Click:
    
    Exit Sub
    
Err_mnupopuprestore_Click:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during mnupopuprestore_Click, in frmLoadToSystray", vbInformation, "Advisory"
            Resume Exit_mnupopuprestore_Click
    End Select
    
End Sub
