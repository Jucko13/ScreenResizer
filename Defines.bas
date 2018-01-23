Attribute VB_Name = "Defines"


Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)



Declare Function EnumDisplaySettings _
 Lib "user32" Alias "EnumDisplaySettingsA" ( _
 ByVal lpszDeviceName As String, _
 ByVal iModeNum As Long, _
 ByRef lpDevMode As devMode) As Boolean
 
Declare Function ChangeDisplaySettings _
 Lib "user32" Alias "ChangeDisplaySettingsA" ( _
 ByRef lpDevMode As devMode, _
 ByVal dwflags As Long) As Long
 

Private Type POINTL
    X As Long
    Y As Long
End Type

Type devMode
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmPosition As POINTL
    dmDisplayOrientation As Long
    dmDisplayFixedOutput As Long
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmLogPixels As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

'Type devMode
'    dmDeviceName As String * 32
'    dmSpecVersion As Integer
'    dmDriverVersion As Integer
'    dmSize As Integer
'    dmDriverExtra As Integer
'    dmFields As Long
'    dmOrientation As Integer
'    dmPaperSize As Integer
'    dmPaperLength As Integer
'    dmPaperWidth As Integer
'    dmScale As Integer
'    dmCopies As Integer
'    dmDefaultSource As Integer
'    dmPrintQuality As Integer
'    dmColor As Integer
'    dmDuplex As Integer
'    dmYResolution As Integer
'    dmTTOption As Integer
'    dmCollate As Integer
'    dmFormName As String * 32
'    dmUnusedPadding As Integer
'    dmBitsPerPel As Integer
'    dmPelsWidth As Long
'    dmPelsHeight As Long
'    dmDisplayFlags As Long
'    dmDisplayFrequency As Long
'End Type

Declare Function ChangeDisplaySettingsEx Lib "user32" Alias "ChangeDisplaySettingsExA" _
        (ByVal lpszDeviceName As String, ByRef lpDevMode As devMode, ByVal hWnd As Long, _
        ByVal dwflags As Long, ByVal lParam As Long) As Long
    
    
Declare Function EnumDisplayDevices Lib "user32" _
   Alias "EnumDisplayDevicesA" _
  (ByVal lpDevice As Any, _
   ByVal iDevNum As Long, _
   lpDisplayDevice As DISPLAY_DEVICE, _
   ByVal dwflags As Long) As Long
   
   
Global Const CDS_UPDATEREGISTRY = &H1
Global Const CDS_TEST = &H4
Global Const CDS_RESET = &H40000000


Global Const DISP_CHANGE_SUCCESSFUL = 0
Global Const DISP_CHANGE_RESTART = 1

Global Const DM_BITSPERPEL = &H40000
Global Const DM_PELSWIDTH = &H80000
Global Const DM_PELSHEIGHT = &H100000
Global Const DM_DISPLAYFREQUENCY = &H400000


Type DISPLAY_DEVICE
   cb As Long
   DeviceName(0 To 31) As Byte
   DeviceString(0 To 127) As Byte
   StateFlags As Long
   DeviceID(0 To 127) As Byte
   DeviceKey(0 To 127) As Byte
End Type



Sub main()
    'uEnableMouseHooks = True
    uDontDrawDots = True
    
    frmMain.Show
    

End Sub


