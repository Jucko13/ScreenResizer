VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   Caption         =   "Screen Resizer by Ricardo de Roode v1.0"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8790
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   478
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   586
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrErrorHide 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4050
      Top             =   585
   End
   Begin Project1.uTextBox utxtError 
      Height          =   525
      Left            =   3540
      TabIndex        =   30
      Top             =   450
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   926
      BackgroundColor =   7667560
      BorderColor     =   49152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   49152
      ConsoleColors   =   0   'False
      RowLineColor    =   255
      HideCursor      =   -1  'True
      AutoResize      =   -1  'True
   End
   Begin Project1.uToolTip uttHelp 
      Height          =   420
      Left            =   4710
      TabIndex        =   10
      Top             =   330
      Visible         =   0   'False
      Width           =   780
      _ExtentX        =   79
      _ExtentY        =   53
   End
   Begin Project1.uButton ubtnScanMonitors 
      Height          =   330
      Left            =   135
      TabIndex        =   2
      Top             =   135
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   582
      BackgroundColor =   33023
      ForeColor       =   16777215
      MouseOverBackgroundColor=   8438015
      FocusColor      =   0
      BackgroundColorDisabled=   12632256
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      MouseOverBackgroundColorDisabled=   12632256
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "Scan Monitors"
      Border          =   0   'False
      BorderAnimation =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.uFrame ufrmSavedSettings 
      Height          =   6870
      Index           =   0
      Left            =   5505
      TabIndex        =   0
      Top             =   60
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   12118
      BackgroundColor =   4210752
      ForeColor       =   16777215
      Caption         =   "Saved Settings"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Project1.uListBox ulstSavedMonitors 
         Height          =   4335
         Left            =   135
         TabIndex        =   1
         Top             =   930
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   7646
         BackgroundColor =   33023
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         Text            =   ""
         SelectionBackgroundColor=   8438015
         SelectionBorderColor=   8438015
         SelectionForeColor=   16777215
         ItemHeight      =   57
      End
      Begin Project1.uButton ubtnSaveCurrent 
         Height          =   555
         Left            =   135
         TabIndex        =   11
         Top             =   255
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   979
         BackgroundColor =   33023
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8438015
         FocusColor      =   0
         BackgroundColorDisabled=   12632256
         BorderColorDisabled=   0
         ForeColorDisabled=   0
         CaptionBorderColorDisabled=   0
         FocusColorDisabled=   0
         FocusVisible    =   0   'False
         Caption         =   "Save Current Monitor Setup"
         Border          =   0   'False
         BorderAnimation =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionOffsetTop=   -1
      End
      Begin Project1.uButton ubtnSetSavedResolution 
         Height          =   555
         Left            =   720
         TabIndex        =   15
         Top             =   6180
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   979
         BackgroundColor =   33023
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8438015
         FocusColor      =   0
         BackgroundColorDisabled=   12632256
         BorderColorDisabled=   0
         ForeColorDisabled=   0
         MouseOverBackgroundColorDisabled=   12632256
         CaptionBorderColorDisabled=   0
         FocusColorDisabled=   0
         FocusVisible    =   0   'False
         Caption         =   "Set Resolution"
         Border          =   0   'False
         BorderAnimation =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.uButton ubntDeleteSave 
         Height          =   555
         Left            =   135
         TabIndex        =   16
         Top             =   6180
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   979
         BackgroundColor =   33023
         ForeColor       =   16777215
         MouseOverBackgroundColor=   8438015
         FocusColor      =   0
         BackgroundColorDisabled=   12632256
         BorderColorDisabled=   0
         ForeColorDisabled=   0
         MouseOverBackgroundColorDisabled=   12632256
         CaptionBorderColorDisabled=   0
         FocusColorDisabled=   0
         FocusVisible    =   0   'False
         Caption         =   "Del"
         Border          =   0   'False
         BorderAnimation =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.uFrame ufrmSavedSettings 
         Height          =   735
         Index           =   4
         Left            =   135
         TabIndex        =   17
         Top             =   5310
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1296
         BackgroundColor =   4210752
         ForeColor       =   16777215
         Caption         =   "Options for selected item"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Project1.uCheckBox uchkLoadOnStartup 
            Height          =   450
            Left            =   60
            TabIndex        =   18
            Top             =   225
            Visible         =   0   'False
            Width           =   2745
            _ExtentX        =   1455
            _ExtentY        =   794
            BackgroundColor =   4210752
            Border          =   0   'False
            BorderThickness =   2
            Caption         =   "Load startup"
            CaptionOffsetLeft=   10
            CheckBorderThickness=   0
            CheckSelectionColor=   33023
            CheckSize       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
            AutoSize        =   0   'False
         End
      End
   End
   Begin Project1.uFrame ufrmSavedSettings 
      Height          =   6045
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   885
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   10663
      BackgroundColor =   4210752
      ForeColor       =   16777215
      Caption         =   "Monitors Found"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox picDisplays 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   135
         ScaleHeight     =   85
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   325
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   750
         Width           =   4875
         Begin VB.PictureBox picDisplay 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H000040C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   645
            Index           =   0
            Left            =   255
            ScaleHeight     =   43
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   76
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   225
            Visible         =   0   'False
            Width           =   1140
         End
      End
      Begin Project1.uFrame ufrmSavedSettings 
         Height          =   1335
         Index           =   2
         Left            =   135
         TabIndex        =   4
         Top             =   2100
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   2355
         BackgroundColor =   4210752
         ForeColor       =   16777215
         Caption         =   "Display Information"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Project1.uButton ubtnSetPrimary 
            Height          =   270
            Left            =   3645
            TabIndex        =   31
            Top             =   1050
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   476
            BackgroundColor =   33023
            ForeColor       =   16777215
            MouseOverBackgroundColor=   8438015
            FocusColor      =   0
            BackgroundColorDisabled=   12632256
            BorderColorDisabled=   0
            ForeColorDisabled=   0
            MouseOverBackgroundColorDisabled=   12632256
            CaptionBorderColorDisabled=   0
            FocusColorDisabled=   0
            FocusVisible    =   0   'False
            Caption         =   "Set Primary"
            Border          =   0   'False
            BorderAnimation =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Offset: x:0 y:0"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   12
            Top             =   1035
            Width           =   4455
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "VideoCard:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   765
            Width           =   4485
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Colors:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   495
            Width           =   4425
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hardware:"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   225
            Width           =   945
         End
      End
      Begin Project1.uDropDown udrpMonitors 
         Height          =   375
         Left            =   135
         TabIndex        =   5
         Top             =   255
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   661
         BackgroundColor =   33023
         BorderColor     =   33023
         ForeColor       =   16777215
         SelectionBackgroundColor=   33023
         SelectionBorderColor=   33023
         BackgroundColorDisabled=   12632256
         SelectionBackgroundColorDisabled=   16777215
         SelectionBorderColorDisabled=   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         Border          =   0   'False
         VisibleItems    =   6
      End
      Begin Project1.uFrame ufrmSavedSettings 
         Height          =   2385
         Index           =   3
         Left            =   135
         TabIndex        =   19
         Top             =   3525
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   4207
         BackgroundColor =   4210752
         ForeColor       =   16777215
         Caption         =   "Set new Resolution"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Project1.uTextBox utxtOffsetX 
            Height          =   330
            Left            =   1935
            TabIndex        =   27
            Top             =   1065
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   582
            BackgroundColor =   33023
            BorderColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
            Border          =   0   'False
            Border          =   0   'False
         End
         Begin Project1.uDropDown udrpResolution 
            Height          =   330
            Left            =   1935
            TabIndex        =   20
            Top             =   255
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   582
            BackgroundColor =   33023
            ForeColor       =   16777215
            SelectionBackgroundColor=   33023
            SelectionBorderColor=   33023
            BackgroundColorDisabled=   12632256
            SelectionBackgroundColorDisabled=   16777215
            SelectionBorderColorDisabled=   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            Border          =   0   'False
            VisibleItems    =   6
         End
         Begin Project1.uDropDown udrpRefreshRate 
            Height          =   330
            Left            =   1935
            TabIndex        =   21
            Top             =   660
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   582
            BackgroundColor =   33023
            ForeColor       =   16777215
            SelectionBackgroundColor=   33023
            SelectionBorderColor=   33023
            BackgroundColorDisabled=   12632256
            SelectionBackgroundColorDisabled=   16777215
            SelectionBorderColorDisabled=   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            Border          =   0   'False
            VisibleItems    =   6
         End
         Begin Project1.uButton ubtnSetResolution 
            Height          =   375
            Left            =   1935
            TabIndex        =   24
            Top             =   1875
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   661
            BackgroundColor =   33023
            ForeColor       =   16777215
            MouseOverBackgroundColor=   8438015
            FocusColor      =   0
            BackgroundColorDisabled=   12632256
            BorderColorDisabled=   0
            ForeColorDisabled=   0
            MouseOverBackgroundColorDisabled=   12632256
            CaptionBorderColorDisabled=   0
            FocusColorDisabled=   0
            FocusVisible    =   0   'False
            Caption         =   "Set Resolution"
            Border          =   0   'False
            BorderAnimation =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Project1.uTextBox utxtOffsetY 
            Height          =   330
            Left            =   1935
            TabIndex        =   28
            Top             =   1470
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   582
            BackgroundColor =   33023
            BorderColor     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Consolas"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
            Border          =   0   'False
            Border          =   0   'False
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "4: OffsetY"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   4
            Left            =   135
            TabIndex        =   29
            Top             =   1515
            Width           =   1725
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "3: OffsetX"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   3
            Left            =   135
            TabIndex        =   26
            Top             =   1110
            Width           =   1725
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "3: Profit"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   2
            Left            =   135
            TabIndex        =   25
            Top             =   1935
            Width           =   1725
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "1: Resolution"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   23
            Top             =   315
            Width           =   1515
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "2: Refresh Rate"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   1
            Left            =   135
            TabIndex        =   22
            Top             =   705
            Width           =   1725
         End
      End
   End
   Begin Project1.uButton ubtnRefreshSetup 
      Height          =   330
      Left            =   135
      TabIndex        =   32
      Top             =   540
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   582
      BackgroundColor =   33023
      ForeColor       =   16777215
      MouseOverBackgroundColor=   8438015
      FocusColor      =   0
      BackgroundColorDisabled=   12632256
      BorderColorDisabled=   0
      ForeColorDisabled=   0
      MouseOverBackgroundColorDisabled=   12632256
      CaptionBorderColorDisabled=   0
      FocusColorDisabled=   0
      FocusVisible    =   0   'False
      Caption         =   "Refresh Setup"
      Border          =   0   'False
      BorderAnimation =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatus 
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   2160
      TabIndex        =   9
      Top             =   135
      Width           =   3300
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type MonitorType
    data As DISPLAY_DEVICE
    displayResolutions() As devMode
    scannedForModes As Boolean
    displayResolutionCurrent As devMode
End Type

Private Type MonitorAndSettingType
    data As DISPLAY_DEVICE
    displayResolutionCurrent As devMode
End Type

Dim monitors() As MonitorType
Dim monitorCount As Long

Private Type ProgramSaveType
    monitorSaveData() As MonitorAndSettingType
    saveName As String
    loadOnStartup As Boolean
    isSave As Boolean
End Type

Dim savedResolution(0 To 4) As ProgramSaveType

Private Function TrimNull(Item As String)

    Dim pos As Integer
   
   'double check that there is a chr$(0) in the string
    pos = InStr(Item, Chr$(0))
    If pos Then
       TrimNull = Left$(Item, pos - 1)
    Else
       TrimNull = Item
    End If
  
End Function

Sub refreshCurrentResolution(Optional scanModes As Boolean = True)
    Dim i As Long
    Dim sDeviceName As String
    
    For i = 0 To UBound(monitors)
        sDeviceName = TrimNull(StrConv(monitors(i).data.DeviceName, vbUnicode))
        
        EnumDisplaySettings sDeviceName, -1, monitors(i).displayResolutionCurrent
    Next i
    
    calculateDisplayPositions scanModes, udrpMonitors.ListIndex
    
End Sub

Sub detectAllMonitors(Optional scanModes As Boolean = True)
    Static isScanning As Boolean
    If isScanning Then Exit Sub
    isScanning = True
    ubtnScanMonitors.Enabled = False
    
    
    Dim sDeviceName As String
    Dim dev As DISPLAY_DEVICE
    Dim deviceIndex As Long
    Dim enumIndex As Long
    Dim deviceModeIndex As Long
    Dim deviceModeIndexReal As Long
    
    Dim enumResults As Long
    
    
    Dim totalModes As Long
    Dim totalMonitors As Long
    Dim totalModesRefreshCounter As Long
    lblStatus.Caption = "Monitors: 0" & vbCrLf & "Modes: 0"
    
    totalMonitors = 1
    
    udrpMonitors.Clear
    
    Dim c As Long
    
    ReDim monitors(0)
    
    monitors(0).data.cb = Len(monitors(0).data)
    
    Do While EnumDisplayDevices(0&, enumIndex, monitors(deviceIndex).data, 0&)
        
        sDeviceName = TrimNull(StrConv(monitors(deviceIndex).data.DeviceName, vbUnicode))
        
        'get current display mode
        EnumDisplaySettings sDeviceName, -1, monitors(deviceIndex).displayResolutionCurrent
        
        ReDim monitors(deviceIndex).displayResolutions(0)
        
        monitors(deviceIndex).scannedForModes = False
        
        'start of display modes
        deviceModeIndex = 0
        deviceModeIndexReal = 0
        Do
            enumResults = EnumDisplaySettings(sDeviceName, deviceModeIndexReal, monitors(deviceIndex).displayResolutions(deviceModeIndex))
            If enumResults <= 0 Then
                deviceModeIndex = deviceModeIndex - 1
                totalModes = totalModes - 1
                If deviceModeIndex >= 0 Then ReDim Preserve monitors(deviceIndex).displayResolutions(deviceModeIndex)
                
                Exit Do
            End If
            
            'Debug.Print enumResults
            
            totalModesRefreshCounter = totalModesRefreshCounter + 1
            If totalModesRefreshCounter > 150 Then
                lblStatus.Caption = "Monitors: " & totalMonitors & vbCrLf & "Modes: " & totalModes
                DoEvents
                totalModesRefreshCounter = 0
                'Exit Do
            End If
            
            totalModes = totalModes + 1
            
            deviceModeIndexReal = deviceModeIndexReal + 1
            If monitors(deviceIndex).displayResolutions(deviceModeIndex).dmBitsPerPel = 32 Then
                
                deviceModeIndex = deviceModeIndex + 1
                ReDim Preserve monitors(deviceIndex).displayResolutions(deviceModeIndex)
            End If
            
            If scanModes = False Then Exit Do
            'monitors(deviceModeIndex).data.cb = Len(monitors(deviceIndex).data)
        Loop
        'end of display modes
            
        monitors(deviceIndex).scannedForModes = scanModes
        
        If deviceModeIndex > 0 Then
            If scanModes Then
                'add display to the list
                udrpMonitors.AddItem sDeviceName
            End If
            
            'increase index for next monitor to detect
            deviceIndex = deviceIndex + 1
            totalMonitors = totalMonitors + 1
            
            ReDim Preserve monitors(deviceIndex)
            monitors(deviceIndex).data.cb = Len(monitors(deviceIndex).data)
        End If
        
        
        enumIndex = enumIndex + 1
    Loop
    
    totalMonitors = totalMonitors - 1
    deviceIndex = deviceIndex - 1
    If deviceIndex >= 0 Then ReDim Preserve monitors(deviceIndex)
    
    udrpMonitors.ItemsVisible = IIf(udrpMonitors.ListCount < 1, 1, udrpMonitors.ListCount)
    
    monitorCount = UBound(monitors) + 1
    lblStatus.Caption = "Monitors: " & totalMonitors & vbCrLf & "Modes: " & totalModes
    If scanModes = False Then
        lblStatus.Caption = lblStatus.Caption & vbCrLf & "Click scan to unlock"

    End If
    
    setMonitorManipulationEnabled scanModes
    
    If scanModes Then testSaveMonitorModes
    
    ubtnScanMonitors.Enabled = True
    isScanning = False
    
    
End Sub

Sub setMonitorManipulationEnabled(Enabled As Boolean)
    udrpMonitors.Enabled = Enabled
    udrpResolution.Enabled = Enabled
    udrpRefreshRate.Enabled = Enabled
    ubtnSetResolution.Enabled = Enabled
    ubtnRefreshSetup.Enabled = Enabled
    
    If Enabled Then
        utxtOffsetX.BackgroundColor = udrpResolution.BackgroundColor
        utxtOffsetY.BackgroundColor = udrpResolution.BackgroundColor
        
        utxtOffsetX.ForeColor = udrpResolution.ForeColor
        utxtOffsetY.ForeColor = udrpResolution.ForeColor
        
    Else
        utxtOffsetX.BackgroundColor = udrpResolution.BackgroundColorDisabled
        utxtOffsetY.BackgroundColor = udrpResolution.BackgroundColorDisabled
        
        utxtOffsetX.ForeColor = udrpResolution.ForeColorDisabled
        utxtOffsetY.ForeColor = udrpResolution.ForeColorDisabled
        
    End If
    
    
    calculateDisplayPositions Enabled
    
    
End Sub


Sub calculateDisplayPositions(Optional scanModes As Boolean = True, Optional selectedIndex As Long = -1)
    Dim i As Long
    
    Dim xMin As Long
    Dim yMin As Long
    Dim xMax As Long
    Dim yMax As Long
    
    For i = 0 To UBound(monitors)
        With monitors(i).displayResolutionCurrent
            
            If .dmPosition.X < xMin Then xMin = .dmPosition.X
            If .dmPosition.Y < yMin Then yMin = .dmPosition.Y
            
            If .dmPosition.X + .dmPelsWidth > xMax Then xMax = .dmPosition.X + .dmPelsWidth
            If .dmPosition.Y + .dmPelsHeight > yMax Then yMax = .dmPosition.Y + .dmPelsHeight
        End With
    Next i
    
    Dim xCenter As Long
    Dim yCenter As Long
    
    Dim H As Long
    Dim W As Long

    
    W = xMax - xMin
    H = yMax - yMin
    
    xCenter = picDisplays.ScaleWidth / 2
    yCenter = picDisplays.ScaleHeight / 2
    
    Dim xScale As Double
    Dim yScale As Double
    Dim rescale As Double
    
    xScale = (picDisplays.ScaleWidth - 3) / W
    yScale = (picDisplays.ScaleHeight - 3) / H
    
    If xScale < yScale Then
        rescale = xScale
    Else
        rescale = yScale
    End If
    
    While picDisplay.Count <= UBound(monitors)
        i = picDisplay.Count
        Load picDisplay(i)
        picDisplay(i).BorderStyle = 0
        picDisplay(i).Visible = True
    Wend
    
    Dim xText As Long
    Dim yText As Long
    Dim tx As Long
    Dim ty As Long
    Dim printableIndex As Long
    Dim tmpDev As devMode
    
    For i = 0 To UBound(monitors)
        picDisplay(i).Visible = True
        picDisplay(i).BackColor = IIf(scanModes, IIf(i = selectedIndex, &H80C0FF, &H80FF&), &HC0C0C0)
        
        printableIndex = ReturnNonAlpha(StrConv(monitors(i).data.DeviceName, vbUnicode))
        
        tmpDev = monitors(i).displayResolutionCurrent
        With picDisplay(i)
            .Left = xCenter - W / 2 * rescale + Fix((tmpDev.dmPosition.X - xMin) * rescale) + 1
            .width = (tmpDev.dmPelsWidth * rescale) - 2
            
            .Top = yCenter - H / 2 * rescale + Fix((tmpDev.dmPosition.Y - yMin) * rescale)
            .Height = (tmpDev.dmPelsHeight * rescale) - 2
            
            
            .FontSize = Fix(.ScaleHeight / 2)
            xText = .ScaleWidth / 2 - .TextWidth(i & "") / 2
            yText = .ScaleHeight / 2 - .TextHeight(i & "") / 2 - 1
            
            For tx = -1 To 1
                For ty = -1 To 1
                    If tx <> 0 And ty <> 0 Then
                        .CurrentX = xText + tx
                        .CurrentY = yText + ty
                        .ForeColor = IIf(scanModes, IIf(i = selectedIndex, &H80FF&, &H50FF), &H808080)
                        picDisplay(i).Print printableIndex & ""
                    End If
                Next ty
            Next tx
            
            .CurrentX = xText
            .CurrentY = yText
            .ForeColor = IIf(scanModes, &HFFFFFF, &HA0A0A0)
            picDisplay(i).Print printableIndex & ""
            
            
            If tmpDev.dmPosition.X = 0 And tmpDev.dmPosition.Y = 0 Then 'primary monitor
                .CurrentX = 3
                .CurrentY = 0
                .FontSize = Fix(.ScaleHeight / 6)
                .ForeColor = IIf(scanModes, &HFFFFFF, &HA0A0A0)
                picDisplay(i).Print "P"
            End If
            
        End With
    Next i
    
    
End Sub

Public Function ReturnNonAlpha(ByVal sString As String) As String
   Dim i As Integer
   For i = 1 To Len(sString)
       If Mid(sString, i, 1) Like "[0-9]" Then
           ReturnNonAlpha = ReturnNonAlpha + Mid(sString, i, 1)
       End If
   Next i
   
   If ReturnNonAlpha = "" Then ReturnNonAlpha = "0"
End Function

Private Sub Form_Load()
    
    uttHelp.setForm Me
 
    uttHelp.Add ubtnScanMonitors.hWnd, "This scans the system for connected monitors" & vbCrLf & "and their supported modes."
   
    uttHelp.StartTimer
    
    
    ubtnSaveCurrent.Caption = "Save Current" & vbCrLf & "Monitor Setup"
    
    uchkLoadOnStartup.Caption = "Load saved settings" & vbCrLf & "on program start."
    
    
    
    ulstSavedMonitors.setTabStop 0, 5
    
    'testLoadMonitorModes
    
    loadMonitors
    
    detectAllMonitors False
    
    lblStatus.BackStyle = 0
    'ulstSavedMonitors.AddItem "1920x1080 @ 144Hz" & vbCrLf & "1920x1080 @ 144Hz" & vbCrLf & "1920x1080 @ 144Hz"
    ' detectAllMonitors
    
    'showMessage "Could not change display 1"
    
   
    Dim i As Long
    
    For i = 0 To 4
        If savedResolution(i).isSave Then
            If savedResolution(i).loadOnStartup Then
                setSavedResolution i
                Exit For
            End If
        End If
    Next i
   

End Sub



Sub showMessage(Message As String, Optional isError As Boolean = False)
    tmrErrorHide.Enabled = False
    
    With utxtError
        .Text = " " & Message
        .Redraw
        .Left = Me.ScaleWidth / 2 - .width / 2
        .Top = Me.ScaleHeight / 2 - .Height / 2
        .Visible = True
        If isError Then
            tmrErrorHide.Interval = 1000
            .BackgroundColor = &H8080FF
            .ForeColor = &HFF&
            .BorderColor = &HFF&
        Else
            tmrErrorHide.Interval = 2000
            .BackgroundColor = &H74FF68
            .ForeColor = &HC000&
            .BorderColor = &HC000&
        End If
        
    End With
    
    tmrErrorHide.Enabled = True
End Sub


Private Sub Form_Unload(cancel As Integer)
    On Error Resume Next
    
    Unload frmIdentify
End Sub

Private Sub lblInfo_DblClick(Index As Integer)
    Clipboard.Clear
    
    Clipboard.SetText lblInfo(Index).Caption
    showMessage "Text Copied!"
End Sub

Private Sub picDisplay_Click(Index As Integer)
    udrpMonitors.ListIndex = Index
End Sub

Private Sub picDisplay_DblClick(Index As Integer)

    Dim printableIndex As Long
    
    printableIndex = ReturnNonAlpha(StrConv(monitors(Index).data.DeviceName, vbUnicode))
    
    With monitors(Index).displayResolutionCurrent
        frmIdentify.customShow printableIndex, .dmPosition.X, .dmPosition.Y, .dmPelsWidth, .dmPelsHeight
    End With
    
End Sub

Private Sub tmrErrorHide_Timer()
    utxtError.Visible = False
    tmrErrorHide.Enabled = False
End Sub

Private Sub ubntDeleteSave_Click(Button As Integer, X As Single, Y As Single)
    Dim Index As Long
    
    Index = ulstSavedMonitors.ListIndex
    If Index < 0 Then Exit Sub
    
    Index = ulstSavedMonitors.ItemData(Index)
    
    savedResolution(Index).isSave = False
    Erase savedResolution(Index).monitorSaveData
    
    saveMonitors
    
    loadMonitors
    
    ulstSavedMonitors_ItemChange ulstSavedMonitors.ListIndex
End Sub

Private Sub ubtnRefreshSetup_Click(Button As Integer, X As Single, Y As Single)
    If Not ubtnRefreshSetup.Enabled Then Exit Sub
    
    refreshCurrentResolution udrpMonitors.Enabled
    showMessage "Done!"
End Sub

Private Sub ubtnSaveCurrent_Click(Button As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim devModes() As devMode
    
    Dim displayValue As String
    
    If monitorCount <= 0 Then Exit Sub
    
    Dim newIndex As Long
    
    newIndex = -1
    For i = 0 To 4
        If savedResolution(i).isSave = False Then
            newIndex = i
            Exit For
        End If
    Next i
    
    If newIndex = -1 Then
        MsgBox "No more save space available"
        Exit Sub
    End If
    
    savedResolution(newIndex).isSave = True
    savedResolution(newIndex).saveName = InputBox("Enter the description for this save:")
    ReDim savedResolution(newIndex).monitorSaveData(0 To monitorCount - 1)
    
    For i = 0 To monitorCount - 1
        savedResolution(newIndex).monitorSaveData(i).displayResolutionCurrent = monitors(i).displayResolutionCurrent
        savedResolution(newIndex).monitorSaveData(i).data = monitors(i).data
    Next i
    
    ulstSavedMonitors.AddItem displayValue
    
    saveMonitors

    loadMonitors
    
End Sub


Sub testLoadMonitorModes()
    Dim nFile As Long
    Dim m As MonitorType
    
    nFile = FreeFile
    
    Open App.Path & "\monitorModes.bin" For Binary Access Read As #nFile
        ReDim monitors(0 To 2)
        Get #nFile, , monitors
    Close nFile
    
    'Debug.Print UBound(monitors)
    setMonitorManipulationEnabled True
    
    Dim i As Long
    
    udrpMonitors.Clear
    
    For i = 0 To 2
        udrpMonitors.AddItem TrimNull(StrConv(monitors(i).data.DeviceName, vbUnicode))
    Next i
    
    udrpMonitors.ItemsVisible = 3
End Sub


Sub testSaveMonitorModes()
    Dim nFile As Long
    nFile = FreeFile
    
    Open App.Path & "\monitorModes.bin" For Binary Access Write As #nFile
        Put #nFile, , monitors
    Close nFile
End Sub


Sub saveMonitors()
    Dim nFile As Long
    nFile = FreeFile
    
    Open App.Path & "\monitors.bin" For Binary Access Write As #nFile
        Put #nFile, , savedResolution
    Close nFile
End Sub

Sub loadMonitors()
    Dim nFile As Long
    nFile = FreeFile
    
    Dim i As Long
    Dim j As Long
    
    Dim displayValue As String
    ulstSavedMonitors.Clear

    'ReDim devModes(0 To 2)
    
    Open App.Path & "\monitors.bin" For Binary Access Read Write As #nFile
        If (LOF(nFile) > 0) Then
            Get #nFile, , savedResolution
            
            On Error GoTo endit:
            For i = 0 To 4
                If savedResolution(i).isSave Then
                    displayValue = savedResolution(i).saveName & vbCrLf & (UBound(savedResolution(i).monitorSaveData) + 1) & " monitor(s)" & vbCrLf
                    
                    For j = 0 To UBound(savedResolution(i).monitorSaveData)
                        If j > 0 Then displayValue = displayValue & ", "
                        With savedResolution(i).monitorSaveData(j)
                            displayValue = displayValue & .displayResolutionCurrent.dmPelsWidth & "x" & .displayResolutionCurrent.dmPelsHeight & " @ " & .displayResolutionCurrent.dmDisplayFrequency & "Hz"
                        End With
                        
                    Next j
                    
                    ulstSavedMonitors.AddItem displayValue, i
                End If
                
            Next i
            
            
            
    
        End If
    Close nFile
endit:
    

End Sub

Private Sub ubtnScanMonitors_Click(Button As Integer, X As Single, Y As Single)
    If Not ubtnScanMonitors.Enabled Then Exit Sub
    showMessage "Could take some time..."
    detectAllMonitors
    showMessage "Done!"
End Sub

Private Sub ubtnSetPrimary_Click(Button As Integer, X As Single, Y As Single)
    If Not ubtnSetPrimary.Enabled Then Exit Sub
    
    Dim monitorIndex As Long
    Dim sDeviceName As String
    
    monitorIndex = udrpMonitors.ListIndex
    If monitorIndex = -1 Then Exit Sub
    
    sDeviceName = TrimNull(StrConv(monitors(monitorIndex).data.DeviceName, vbUnicode))
    
    monitors(monitorIndex).displayResolutionCurrent.dmPosition.X = 0
    monitors(monitorIndex).displayResolutionCurrent.dmPosition.Y = 0
    monitors(monitorIndex).displayResolutionCurrent.dmFields = DM_POSITION
    
    Select Case ChangeDisplaySettingsEx(sDeviceName, monitors(monitorIndex).displayResolutionCurrent, 0, CDS_UPDATEREGISTRY Or CDS_SET_PRIMARY Or CDS_NORESET, 0)
        Case DISP_CHANGE_SUCCESSFUL
            showMessage "GOOD"
        Case DISP_CHANGE_RESTART
            showMessage "restart"
        Case Else
            showMessage "bad", True
    End Select
    
    Debug.Print "click"
End Sub

Private Sub ubtnSetResolution_Click(Button As Integer, X As Single, Y As Single)
    Dim monitorIndex As Long
    Dim devModeIndex As Long
    Dim sDeviceName As String
    
    
    monitorIndex = udrpMonitors.ListIndex
    If monitorIndex = -1 Then Exit Sub
    
    devModeIndex = udrpRefreshRate.ListIndex
    If devModeIndex = -1 Then Exit Sub
    
    Dim d As devMode
    
    devModeIndex = udrpRefreshRate.ItemData(devModeIndex)
    
    sDeviceName = TrimNull(StrConv(monitors(monitorIndex).data.DeviceName, vbUnicode))
    d = monitors(monitorIndex).displayResolutions(devModeIndex)
    
    If Not IsNumeric(utxtOffsetX.Text) Or Not IsNumeric(utxtOffsetY.Text) Then Exit Sub
    
    d.dmPosition.X = Val(utxtOffsetX.Text)
    d.dmPosition.Y = Val(utxtOffsetY.Text)
    
    d.dmFields = DM_POSITION
    
    Debug.Print d.dmPelsWidth & "x" & d.dmPelsHeight & " @ " & d.dmDisplayFrequency
    
    Select Case setResolution(sDeviceName, d)
        Case 0
            showMessage "Resolution set for " & sDeviceName
        Case 1
            showMessage "Restart required for " & sDeviceName, True
        Case -1
            showMessage "Could not change " & sDeviceName, True
        
    End Select
    refreshCurrentResolution
    
    
    
End Sub

Private Function setResolution(DeviceName As String, dev As devMode) As Long
    Select Case ChangeDisplaySettingsEx(DeviceName, dev, 0, CDS_UPDATEREGISTRY, 0)
        Case DISP_CHANGE_SUCCESSFUL
            Debug.Print DeviceName & " succeeded"
            setResolution = 0
        Case DISP_CHANGE_RESTART
            Debug.Print DeviceName & " needs a restart"
            setResolution = 1
        Case Else
            Debug.Print DeviceName & " could not change"
            setResolution = -1
    End Select
End Function

Private Sub ubtnSetSavedResolution_Click(Button As Integer, X As Single, Y As Single)
    Dim Index As Long
    
    
    Index = ulstSavedMonitors.ListIndex
    If Index = -1 Then Exit Sub
    
    
    
    setSavedResolution Index
End Sub

Sub setSavedResolution(Index As Long)
    Dim i As Long
    Dim DeviceName As String
    On Error GoTo endit:
    
    For i = 0 To UBound(savedResolution(Index).monitorSaveData)
        DeviceName = TrimNull(StrConv(savedResolution(Index).monitorSaveData(i).data.DeviceName, vbUnicode))
        
        setResolution DeviceName, savedResolution(Index).monitorSaveData(i).displayResolutionCurrent
        
    Next i
    
endit:
    refreshCurrentResolution udrpMonitors.Enabled
End Sub

Private Sub uchkLoadOnStartup_ActivateNextState(u_Cancel As Boolean, u_NewState As uCheckboxConstants)
    Dim Index As Long
    
    Index = ulstSavedMonitors.ListIndex
    If Index = -1 Then Exit Sub
    
    Index = ulstSavedMonitors.ItemData(Index)
    
    Dim i As Long
    
    For i = 0 To 4
        If savedResolution(i).isSave Then
            savedResolution(i).loadOnStartup = False
        End If
    Next i
    
    If u_NewState = u_unChecked Then
        u_NewState = u_Checked
    Else
        u_NewState = u_unChecked
    End If
    
    savedResolution(Index).loadOnStartup = IIf(u_NewState = u_Checked, True, False)

    saveMonitors
End Sub

Private Sub udrpMonitors_ItemChange(ItemIndex As Long)
    Dim i As Long
    Dim doubles As String
    Dim singles As String
    
    If ItemIndex = -1 Then Exit Sub
    
    With monitors(ItemIndex).displayResolutionCurrent
        lblInfo(0).Caption = "Hardware: " & TrimNull(StrConv(monitors(ItemIndex).data.DeviceID, vbUnicode))
        lblInfo(1).Caption = "Colors: " & .dmBitsPerPel & "-bits"
        lblInfo(2).Caption = "VideoCard: " & TrimNull(StrConv(monitors(ItemIndex).data.DeviceString, vbUnicode))
        lblInfo(3).Caption = "Offset: x:" & .dmPosition.X & " y:" & .dmPosition.Y
        
        If .dmPosition.X = 0 And .dmPosition.Y = 0 Then
            ubtnSetPrimary.Enabled = False
        Else
            ubtnSetPrimary.Enabled = True
        End If
    End With
    
    
    
    
    calculateDisplayPositions True, ItemIndex
    
    udrpRefreshRate.Clear
    udrpResolution.Clear
    udrpResolution.RedrawPause
    
    
    Dim selectionIndex As Long
    selectionIndex = -1
    
    For i = 0 To UBound(monitors(ItemIndex).displayResolutions)
        With monitors(ItemIndex).displayResolutions(i)
            singles = .dmPelsWidth & "x" & .dmPelsHeight & " "
            If InStr(1, doubles, singles) = 0 Then
                If monitors(ItemIndex).displayResolutionCurrent.dmPelsWidth = .dmPelsWidth And monitors(ItemIndex).displayResolutionCurrent.dmPelsWidth = .dmPelsWidth Then
                    selectionIndex = udrpResolution.AddItem(singles, i)
                Else
                    udrpResolution.AddItem singles, i
                End If
                
                doubles = doubles & singles
            End If
            
        End With
    Next i
    
    udrpResolution.RedrawResume
    
    utxtOffsetX.Text = monitors(ItemIndex).displayResolutionCurrent.dmPosition.X
    utxtOffsetY.Text = monitors(ItemIndex).displayResolutionCurrent.dmPosition.Y
    
    
    If selectionIndex <> -1 Then udrpResolution.ListIndex = selectionIndex
    
End Sub

Private Sub udrpResolution_ItemChange(ItemIndex As Long)
    If ItemIndex = -1 Then Exit Sub
    
    Dim lWidth As Long
    Dim lHeight As Long
    Dim i As Long
    Dim selectedDevMode As devMode
    Dim doubles As String
    Dim singles As String
    Dim selectionIndex As Long
    selectionIndex = -1
    
    udrpRefreshRate.Clear
    
    Dim lCurrent As Long
    
    lCurrent = udrpMonitors.ListIndex
    If lCurrent = -1 Then Exit Sub
    
    selectedDevMode = monitors(lCurrent).displayResolutions(udrpResolution.ItemData(ItemIndex))
    
    
    
    For i = 0 To UBound(monitors(lCurrent).displayResolutions)
        With monitors(lCurrent).displayResolutions(i)
            If selectedDevMode.dmPelsHeight = .dmPelsHeight And selectedDevMode.dmPelsWidth = .dmPelsWidth Then
                singles = .dmDisplayFrequency & " "
                
                If InStr(1, doubles, singles) = 0 Then
                    If monitors(lCurrent).displayResolutionCurrent.dmDisplayFrequency = .dmDisplayFrequency Then
                        selectionIndex = udrpRefreshRate.AddItem(singles, i)
                    Else
                        udrpRefreshRate.AddItem singles, i
                    End If
                    doubles = doubles & singles
                End If
            
            End If
        End With
    Next i
    
    udrpRefreshRate.ItemsVisible = udrpRefreshRate.ListCount
    
    If selectionIndex <> -1 Then udrpRefreshRate.ListIndex = selectionIndex
    
End Sub



Private Sub ulstSavedMonitors_DblClick()
    ubtnSetSavedResolution_Click 0, 0, 0
End Sub

Private Sub ulstSavedMonitors_ItemChange(ItemIndex As Long)
    Dim Index As Long
    
    Index = ulstSavedMonitors.ListIndex
    If Index = -1 Then
        uchkLoadOnStartup.Visible = False
        Exit Sub
    Else
        uchkLoadOnStartup.Visible = True
    End If
    
    Index = ulstSavedMonitors.ItemData(Index)
    uchkLoadOnStartup.Value = IIf(savedResolution(Index).loadOnStartup, u_Checked, u_unChecked)
    
    
End Sub










Private Sub utxtError_Click(ByVal charIndex As Long, ByVal charRow As Long)
    tmrErrorHide_Timer
End Sub

Private Sub utxtOffsetX_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    Dim tmpText As String
    
    Select Case KeyCode
        Case vbKeyV
            If Shift = 2 Then
                tmpText = Clipboard.GetText
                If Not IsNumeric(tmpText) Then
                    KeyCode = 0
                    Shift = 0
                End If
            Else
                KeyCode = 0
                Shift = 0
            End If
        
        Case vbKeyA, vbKeyC, vbKeyX
            If Not (Shift = 2) Then
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKey0 To vbKey9, 189 '189 = minus
            If Not (Shift = 0) Then
                KeyCode = 0
                Shift = 0
            End If
            
        Case vbKeyBack, vbKeyDelete, vbKeySubtract, vbKeyNumpad0 To vbKeyNumpad9
        
        
        Case Else
            KeyCode = 0
            Shift = 0
            
    End Select
    
End Sub

Private Sub utxtOffsetY_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    utxtOffsetX_KeyDown KeyCode, Shift
End Sub
