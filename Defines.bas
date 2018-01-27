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
 ByVal dwFlags As Long) As Long
 

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
        (ByVal lpszDeviceName As String, ByRef lpDevMode As devMode, ByVal hwnd As Long, _
        ByVal dwFlags As Long, ByVal lParam As Long) As Long
    
    
Declare Function EnumDisplayDevices Lib "user32" _
   Alias "EnumDisplayDevicesA" _
  (ByVal lpDevice As Any, _
   ByVal iDevNum As Long, _
   lpDisplayDevice As DISPLAY_DEVICE, _
   ByVal dwFlags As Long) As Long
   
   
Global Const CDS_UPDATEREGISTRY     As Long = &H1
Global Const CDS_TEST               As Long = &H4
Global Const CDS_RESET              As Long = &H40000000
Global Const CDS_SET_PRIMARY        As Long = &H10
Global Const CDS_NORESET            As Long = &H10000000
Global Const CDS_FORCE              As Long = &H80000000

Global Const DISP_CHANGE_SUCCESSFUL As Long = 0
Global Const DISP_CHANGE_RESTART    As Long = 1

Global Const DM_ORIENTATION = &H1 ' PRINTER
Global Const DM_PAPERSIZE = &H2 ' PRINTER
Global Const DM_PAPERLENGTH = &H4 ' PRINTER
Global Const DM_PAPERWIDTH = &H8 ' PRINTER
Global Const DM_SCALE = &H10 ' PRINTER
Global Const DM_POSITION = &H20
Global Const DM_NUP = &H40
Global Const DM_DISPLAYORIENTATION = &H80 ' DISPLAY -- XP only
Global Const DM_COPIES = &H100 ' PRINTER
Global Const DM_DEFAULTSOURCE = &H200 ' PRINTER
Global Const DM_PRINTQUALITY = &H400 ' PRINTER
Global Const DM_COLOR = &H800 ' PRINTER
Global Const DM_DUPLEX = &H1000 ' PRINTER
Global Const DM_YRESOLUTION = &H2000 ' PRINTER
Global Const DM_TTOPTION = &H4000 ' PRINTER
Global Const DM_COLLATE = &H8000 ' PRINTER
Global Const DM_FORMNAME = &H10000 ' PRINTER
Global Const DM_LOGPIXELS = &H20000
Global Const DM_BITSPERPEL = &H40000 ' DISPLAY
Global Const DM_PELSWIDTH = &H80000 ' DISPLAY
Global Const DM_PELSHEIGHT = &H100000 ' DISPLAY
Global Const DM_DISPLAYFLAGS = &H200000 ' DISPLAY
Global Const DM_DISPLAYFREQUENCY = &H400000 ' DISPLAY
Global Const DM_ICMMETHOD = &H800000
Global Const DM_ICMINTENT = &H1000000
Global Const DM_MEDIATYPE = &H2000000
Global Const DM_DITHERTYPE = &H4000000
Global Const DM_PANNINGWIDTH = &H8000000
Global Const DM_PANNINGHEIGHT = &H10000000
Global Const DM_DISPLAYFIXEDOUTPUT = &H20000000 ' XP only

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
    
    Load frmMain
    

End Sub


