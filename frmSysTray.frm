VERSION 5.00
Begin VB.Form frmSysTray 
   BorderStyle     =   0  'None
   ClientHeight    =   735
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   2055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Menu mnu 
      Caption         =   "Profiles"
      Begin VB.Menu mnuProfile 
         Caption         =   "Profile1"
         Index           =   0
      End
      Begin VB.Menu mnuProfilesSerp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuSerp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public WithEvents FSys As Form
Attribute FSys.VB_VarHelpID = -1
Public Event Click(ClickWhat As String)
Public Event TIcon(F As Form)
Public Event MenuClicked(mnuIndex As Long)

Private nid As NOTIFYICONDATA
Private LastWindowState As Integer



Public Property Let Tooltip(Value As String)
   nid.szTip = Value & vbNullChar
End Property

Public Property Get Tooltip() As String
   Tooltip = nid.szTip
End Property

Public Property Let TrayIcon(Value As StdPicture)
    On Error Resume Next
    ' Value can be a picturebox, image, form or string
    
    Me.Icon = Value
    RaiseEvent TIcon(Me)

    UpdateIcon NIM_MODIFY
End Property

Private Sub Form_Load()
    Me.Visible = False
    Tooltip = App.EXEName
    UpdateIcon NIM_ADD

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim Result As Long
   Dim msg As Long
   
   ' The Form_MouseMove is intercepted to give systray mouse events.
   If Me.ScaleMode = vbPixels Then
      msg = X
   Else
      msg = X / Screen.TwipsPerPixelX
   End If
      
   Select Case msg
      Case WM_RBUTTONDBLCLK
         RaiseEvent Click("RBUTTONDBLCLK")
      Case WM_RBUTTONDOWN
         RaiseEvent Click("RBUTTONDOWN")
      Case WM_RBUTTONUP
         RaiseEvent Click("RBUTTONUP")
         PopupMenu mnu
      Case WM_LBUTTONDBLCLK
         RaiseEvent Click("LBUTTONDBLCLK")
         mRestore_Click
      Case WM_LBUTTONDOWN
         RaiseEvent Click("LBUTTONDOWN")
      Case WM_LBUTTONUP
         RaiseEvent Click("LBUTTONUP")
      Case WM_MBUTTONDBLCLK
         RaiseEvent Click("MBUTTONDBLCLK")
      Case WM_MBUTTONDOWN
         RaiseEvent Click("MBUTTONDOWN")
      Case WM_MBUTTONUP
         RaiseEvent Click("MBUTTONUP")
      Case WM_MOUSEMOVE
         RaiseEvent Click("MOUSEMOVE")
      Case Else
         RaiseEvent Click("OTHER....: " & Format$(msg))
   End Select
End Sub

Private Sub FSys_Resize()
   ' Event generated my main form. WindowState is stored in LastWindowState, so that
   ' it may be re- set when the menu item "Restore" is selected.
   If (FSys.WindowState <> vbMinimized) Then LastWindowState = FSys.WindowState
End Sub

Private Sub FSys_Unload(Cancel As Integer)
   ' Important: remove icon from tray, and unload this form when
   ' the main form is unloaded.
   
   If Cancel = True Then Exit Sub
   
   UpdateIcon NIM_DELETE
   Unload Me
End Sub


Private Sub mRestore_Click()
   ' Don't "restore"  FSys is visible and not minimized.
   If (FSys.Visible And FSys.WindowState <> vbMinimized) Then Exit Sub
   ' Restore LastWindowState
   FSys.WindowState = vbNormal
   FSys.Visible = True
   SetForegroundWindow FSys.hwnd
End Sub

Private Sub UpdateIcon(Value As Long)
   ' Used to add, modify and delete icon.
   With nid
      .cbSize = Len(nid)
      .hwnd = Me.hwnd
      .uID = vbNull
      .uFlags = NIM_DELETE Or NIF_TIP Or NIM_MODIFY
      .uCallbackMessage = WM_MOUSEMOVE
      .hIcon = Me.Icon
   End With
   Shell_NotifyIcon Value, nid
End Sub

Public Sub MeQueryUnload(ByRef F As Form, Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = vbFormControlMenu Then
      ' Cancel by setting Cancel = 1, minimize and hide main window.
      Cancel = 1
      F.WindowState = vbMinimized
      F.Hide
   End If
End Sub

Public Sub MeResize(ByRef F As Form)

    Select Case F.WindowState

        Case vbNormal, vbMaximized
            ' Store LastWindowState
            LastWindowState = F.WindowState

        Case vbMinimized
            F.Hide

    End Select

End Sub

'Private Sub TmrFlash_Timer()
'   ' Change icon.
'   Static LastIconWasFlash1 As Boolean
'   LastIconWasFlash1 = Not LastIconWasFlash1
'   Select Case LastIconWasFlash1
'      Case True
'         Me.Icon = Flash2
'      Case Else
'         Me.Icon = Flash1
'   End Select
'   RaiseEvent TIcon(Me)
'   UpdateIcon NIM_MODIFY
'End Sub

Private Sub mnuExit_Click()
    exitTheProgram = True
    Unload FSys
End Sub

Private Sub mnuProfile_Click(Index As Integer)
    RaiseEvent MenuClicked(Index * 1)
End Sub

Private Sub mnuRestore_Click()
    FSys.WindowState = vbNormal
    FSys.Visible = True
End Sub
