VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Screen Resizer by Ricardo de Roode v1.0"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   Begin Project1.uTextBox utxtError 
      Height          =   525
      Left            =   3540
      TabIndex        =   29
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
   Begin VB.PictureBox picPreviewBackground 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1470
      Left            =   2655
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   234
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   3510
      Begin VB.PictureBox picPreview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1230
         Left            =   120
         ScaleHeight     =   82
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   226
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   120
         Width           =   3390
         Begin VB.PictureBox picPreviewDisplay 
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
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   1140
         End
      End
   End
   Begin Project1.uListBox ulstSavedMonitors 
      Height          =   3795
      Left            =   6150
      TabIndex        =   39
      Top             =   765
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   6694
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
      ItemHeight      =   50
   End
   Begin Project1.uFrame ufrmSavedSettings 
      Height          =   5895
      Index           =   0
      Left            =   6015
      TabIndex        =   0
      Top             =   60
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   10398
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
      Begin Project1.uButton ubtnSaveCurrent 
         Height          =   330
         Left            =   135
         TabIndex        =   10
         Top             =   255
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   582
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
         Height          =   330
         Left            =   870
         TabIndex        =   14
         Top             =   5430
         Width           =   2460
         _ExtentX        =   4339
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
         Height          =   330
         Left            =   135
         TabIndex        =   15
         Top             =   5430
         Width           =   615
         _ExtentX        =   1085
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
         TabIndex        =   16
         Top             =   4575
         Width           =   3195
         _ExtentX        =   5636
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
            TabIndex        =   17
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
   Begin VB.Timer tmrErrorHide 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4050
      Top             =   585
   End
   Begin Project1.uToolTip uttHelp 
      Height          =   420
      Left            =   4710
      TabIndex        =   9
      Top             =   330
      Visible         =   0   'False
      Width           =   780
      _ExtentX        =   79
      _ExtentY        =   53
   End
   Begin Project1.uButton ubtnScanMonitors 
      Height          =   330
      Left            =   135
      TabIndex        =   1
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
      Height          =   5700
      Index           =   1
      Left            =   135
      TabIndex        =   2
      Top             =   885
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   10054
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
         ScaleWidth      =   361
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   750
         Width           =   5415
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
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   225
            Visible         =   0   'False
            Width           =   1140
         End
      End
      Begin Project1.uFrame ufrmSavedSettings 
         Height          =   1065
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   2100
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   1879
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
            TabIndex        =   30
            Top             =   705
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
            Left            =   2040
            TabIndex        =   11
            Top             =   495
            Width           =   2685
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "VideoCard:"
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   7
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
            TabIndex        =   6
            Top             =   495
            Width           =   2055
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hardware:"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   225
            Width           =   945
         End
      End
      Begin Project1.uDropDown udrpMonitors 
         Height          =   375
         Left            =   135
         TabIndex        =   4
         Top             =   255
         Width           =   5415
         _ExtentX        =   9551
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
         Height          =   2340
         Index           =   3
         Left            =   135
         TabIndex        =   18
         Top             =   3225
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   4128
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
            TabIndex        =   26
            Top             =   1065
            Width           =   1410
            _ExtentX        =   2487
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
            TabIndex        =   19
            Top             =   255
            Width           =   3345
            _ExtentX        =   5900
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
            TabIndex        =   20
            Top             =   660
            Width           =   3345
            _ExtentX        =   5900
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
            Height          =   330
            Left            =   1935
            TabIndex        =   23
            Top             =   1875
            Width           =   3345
            _ExtentX        =   5900
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
            Left            =   3870
            TabIndex        =   27
            Top             =   1065
            Width           =   1410
            _ExtentX        =   2487
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
         Begin Project1.uDropDown udrpOrientation 
            Height          =   330
            Left            =   1935
            TabIndex        =   34
            Top             =   1470
            Width           =   3345
            _ExtentX        =   5900
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
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "Y:"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   6
            Left            =   3630
            TabIndex        =   33
            Top             =   1110
            Width           =   435
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "X:"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   5
            Left            =   1695
            TabIndex        =   32
            Top             =   1110
            Width           =   390
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "4: Orientation:"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   4
            Left            =   135
            TabIndex        =   28
            Top             =   1515
            Width           =   1725
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "3: Offset"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   3
            Left            =   135
            TabIndex        =   25
            Top             =   1110
            Width           =   1035
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "3: Profit"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   2
            Left            =   135
            TabIndex        =   24
            Top             =   1920
            Width           =   1725
         End
         Begin VB.Label lblSteps 
            BackStyle       =   0  'Transparent
            Caption         =   "1: Resolution"
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   22
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
            TabIndex        =   21
            Top             =   705
            Width           =   1725
         End
      End
   End
   Begin Project1.uButton ubtnRefreshSetup 
      Height          =   330
      Left            =   135
      TabIndex        =   31
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
   Begin Project1.uCheckBox uchkStartWithWindows 
      Height          =   450
      Left            =   120
      TabIndex        =   38
      Top             =   6660
      Width           =   5700
      _ExtentX        =   1455
      _ExtentY        =   794
      BackgroundColor =   4210752
      Border          =   0   'False
      BorderThickness =   2
      Caption         =   "Start with windows (in system tray)."
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
   Begin Project1.uFrame ufrmSavedSettings 
      Height          =   990
      Index           =   5
      Left            =   6015
      TabIndex        =   40
      Top             =   6090
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   1746
      BackgroundColor =   4210752
      ForeColor       =   16777215
      Caption         =   "Move To System tray:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Project1.uOptionBox uoptMinimize 
         Height          =   330
         Index           =   0
         Left            =   90
         TabIndex        =   41
         Top             =   225
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   582
         BackgroundColor =   4210752
         Border          =   0   'False
         Caption         =   "When clicking minimize"
         CaptionOffsetLeft=   10
         CheckBorderColor=   16777215
         CheckSelectionColor=   33023
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
      End
      Begin Project1.uOptionBox uoptMinimize 
         Height          =   330
         Index           =   1
         Left            =   90
         TabIndex        =   42
         Top             =   585
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   582
         BackgroundColor =   4210752
         Border          =   0   'False
         Caption         =   "When clicking on close"
         CaptionOffsetLeft=   10
         CheckBorderColor=   16777215
         CheckSelectionColor=   33023
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
      End
   End
   Begin VB.Label lblStatus 
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   2160
      TabIndex        =   8
      Top             =   135
      Width           =   3750
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
    displayResolutionCurrent As devMode
    
    displayResolutions() As devMode
    scannedForModes As Boolean
End Type

Private Type MonitorAndSettingType
    data As DISPLAY_DEVICE
    displayResolutionCurrent As devMode
End Type

Private Type ProgramSaveType
    monitorSaveData() As MonitorAndSettingType
    saveName As String
    loadOnStartup As Boolean
    isSave As Boolean
End Type


Dim monitors() As MonitorType
Dim monitorCount As Long
Dim savedResolution(0 To 4) As ProgramSaveType

Private WithEvents frmSystemTray As frmSysTray
Attribute frmSystemTray.VB_VarHelpID = -1



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
    
    calculateDisplayPositions -1, scanModes, udrpMonitors.ListIndex
    
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
    udrpOrientation.Enabled = Enabled
    
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
    
    
    calculateDisplayPositions -1, Enabled
    
    
End Sub

Private Sub getMinMaxOffset(ByRef dev As devMode, ByRef xMin As Long, ByRef yMin As Long, ByRef xMax As Long, ByRef yMax As Long)
    With dev
        If .dmPosition.X < xMin Then xMin = .dmPosition.X
        If .dmPosition.Y < yMin Then yMin = .dmPosition.Y
        
        If .dmPosition.X + .dmPelsWidth > xMax Then xMax = .dmPosition.X + .dmPelsWidth
        If .dmPosition.Y + .dmPelsHeight > yMax Then yMax = .dmPosition.Y + .dmPelsHeight
    End With
End Sub

Public Function HasIndex(ControlArray As Object, ByVal Index As Integer) As Boolean
    HasIndex = (VarType(ControlArray(Index)) <> vbObject)
End Function

Private Sub calculateDisplayPositions(previewWindowIndex As Long, Optional scanModes As Boolean = True, Optional selectedIndex As Long = -1)
    Dim i As Long
    
    Dim xMin As Long
    Dim yMin As Long
    Dim xMax As Long
    Dim yMax As Long
    
    
    Dim container As PictureBox
    Dim displayRectangles As Object
    
    If previewWindowIndex > -1 Then
        Set container = picPreview
        Set displayRectangles = picPreviewDisplay
        
        For i = 0 To UBound(savedResolution(previewWindowIndex).monitorSaveData)
            getMinMaxOffset savedResolution(previewWindowIndex).monitorSaveData(i).displayResolutionCurrent, xMin, yMin, xMax, yMax
        Next i
        
    Else
        Set container = picDisplays
        Set displayRectangles = picDisplay
        
        For i = 0 To UBound(monitors)
            getMinMaxOffset monitors(i).displayResolutionCurrent, xMin, yMin, xMax, yMax
        Next i
    End If
    
    

    
    Dim xCenter As Long
    Dim yCenter As Long
    
    Dim H As Long
    Dim W As Long

    
    W = xMax - xMin
    H = yMax - yMin
    
    xCenter = container.ScaleWidth / 2
    yCenter = container.ScaleHeight / 2
    
    Dim xScale As Double
    Dim yScale As Double
    Dim rescale As Double
    
    xScale = (container.ScaleWidth - 4) / W
    yScale = (container.ScaleHeight - 4) / H
    
    If xScale < yScale Then
        rescale = xScale
    Else
        rescale = yScale
    End If
    
    Dim monCount As Long
    
    If previewWindowIndex > -1 Then
        monCount = UBound(savedResolution(previewWindowIndex).monitorSaveData)
    Else
        monCount = UBound(monitors)
    End If
    
    For i = 0 To max(monCount, picDisplay.Count - 1)
        If Not HasIndex(displayRectangles, i) Then
            Load displayRectangles(i)
        End If
        displayRectangles(i).BorderStyle = 0
        displayRectangles(i).Visible = False
    Next i
    
    Dim xText As Long
    Dim yText As Long
    Dim tx As Long
    Dim ty As Long
    Dim printableIndex As Long
    Dim tmpDev As devMode
    
    For i = 0 To monCount
        displayRectangles(i).Visible = True
        displayRectangles(i).BackColor = IIf(scanModes, IIf(i = selectedIndex, &H80C0FF, &H80FF&), &HC0C0C0)
        
        If previewWindowIndex > -1 Then
            printableIndex = ReturnNonAlpha(StrConv(savedResolution(previewWindowIndex).monitorSaveData(i).data.DeviceName, vbUnicode))
            tmpDev = savedResolution(previewWindowIndex).monitorSaveData(i).displayResolutionCurrent
        Else
            printableIndex = ReturnNonAlpha(StrConv(monitors(i).data.DeviceName, vbUnicode))
            tmpDev = monitors(i).displayResolutionCurrent
        End If
        
        
        
        With displayRectangles(i)
            .Left = xCenter - Fix(W / 2 * rescale) + Fix((tmpDev.dmPosition.X - xMin) * rescale) + 1
            .width = Fix(tmpDev.dmPelsWidth * rescale) - 2
            
            .Top = yCenter - Fix(H / 2 * rescale) + Fix((tmpDev.dmPosition.Y - yMin) * rescale) + 1
            .Height = Fix(tmpDev.dmPelsHeight * rescale) - 2
            
            
            .FontSize = Fix(min(.ScaleHeight, .ScaleWidth) / 2)
            xText = .ScaleWidth / 2 - .TextWidth(i & "") / 2
            yText = .ScaleHeight / 2 - .TextHeight(i & "") / 2 - 1
            
            For tx = -1 To 1
                For ty = -1 To 1
                    If tx <> 0 And ty <> 0 Then
                        .CurrentX = xText + tx
                        .CurrentY = yText + ty
                        .ForeColor = IIf(scanModes, IIf(i = selectedIndex, &H80FF&, &H50FF), &H808080)
                        displayRectangles(i).Print printableIndex & ""
                    End If
                Next ty
            Next tx
            
            .CurrentX = xText
            .CurrentY = yText
            .ForeColor = IIf(scanModes, &HFFFFFF, &HA0A0A0)
            displayRectangles(i).Print printableIndex & ""
            
            
            If tmpDev.dmPosition.X = 0 And tmpDev.dmPosition.Y = 0 Then 'primary monitor
                .CurrentX = 3
                .CurrentY = 0
                .FontSize = Fix(min(.ScaleHeight, .ScaleWidth) / 6)
                .ForeColor = IIf(scanModes, &HFFFFFF, &HA0A0A0)
                displayRectangles(i).Print "P"
            End If
            
        End With
    Next i
    
    
End Sub

Function min(a As Long, b As Long) As Long
    min = IIf(a < b, a, b)
End Function

Function max(a As Long, b As Long) As Long
    max = IIf(a > b, a, b)
End Function

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
    Load frmSysTray
    
    Set frmSystemTray = frmSysTray
    Set frmSystemTray.FSys = Me
    frmSystemTray.TrayIcon = Me.Icon
    
    uttHelp.setForm Me
 
    uttHelp.Add ubtnScanMonitors.hwnd, "This scans the system for connected monitors" & vbCrLf & "and their supported modes."
   
    uttHelp.StartTimer
    
    
    'ubtnSaveCurrent.Caption = "Save Current" & vbCrLf & "Monitor Setup"
    
    uchkLoadOnStartup.Caption = "Load saved settings" & vbCrLf & "on program start."
    
    udrpOrientation.AddItem "Landscape:   0°    ", 0
    udrpOrientation.AddItem "Landscape: 180°    ", 2
    udrpOrientation.AddItem "Portrait:   90° CCW", 1
    udrpOrientation.AddItem "Portrait:   90°  CW", 3
    
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
   
    uchkStartWithWindows.Value = IIf(WillRunAtStartup("ScreenResizer"), u_Checked, u_unChecked)
    
    Debug.Print Command
    
    If Command = "startup" Then
        Me.Visible = False
    Else
        Me.Visible = True
    End If
    
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


Private Sub Form_Resize()
    If uoptMinimize(0).Value = u_Selected Then
        If Me.WindowState = vbMinimized Then
            Me.Visible = False
        End If
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Unload frmIdentify
End Sub

Private Sub frmSystemTray_MenuClicked(mnuIndex As Long)
    ulstSavedMonitors.ListIndex = mnuIndex
    ubtnSetSavedResolution_Click 0, 0, 0
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
    'hide all the menus
    For i = 0 To frmSysTray.mnuProfile.Count - 1
        frmSysTray.mnuProfile(i).Visible = False
    Next i
    
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
                    
                    If Not HasIndex(frmSysTray.mnuProfile, i) Then
                        Load frmSysTray.mnuProfile(i)
                    End If
                    frmSysTray.mnuProfile(i).Caption = savedResolution(i).saveName
                    frmSysTray.mnuProfile(i).Visible = True
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
    
    
    If ubtnSetResolution.Enabled = False Then Exit Sub
    
    monitorIndex = udrpMonitors.ListIndex
    If monitorIndex = -1 Then Exit Sub
    
    devModeIndex = udrpRefreshRate.ListIndex
    If devModeIndex = -1 Then Exit Sub
    
    Dim d As devMode
    
    devModeIndex = udrpRefreshRate.ItemData(devModeIndex)
    
    sDeviceName = TrimNull(StrConv(monitors(monitorIndex).data.DeviceName, vbUnicode))
    d = monitors(monitorIndex).displayResolutions(devModeIndex)
    
    
    Debug.Print d.dmDeviceName
    Debug.Print d.dmSpecVersion
    Debug.Print d.dmDriverVersion
    Debug.Print d.dmSize
    Debug.Print d.dmDriverExtra
    Debug.Print d.dmFields
    Debug.Print d.dmPosition.X
    Debug.Print d.dmPosition.Y
    Debug.Print d.dmDisplayOrientation
    Debug.Print d.dmDisplayFixedOutput
    Debug.Print d.dmColor
    Debug.Print d.dmDuplex
    Debug.Print d.dmYResolution
    Debug.Print d.dmTTOption
    Debug.Print d.dmCollate
    Debug.Print d.dmFormName
    Debug.Print d.dmLogPixels
    Debug.Print d.dmBitsPerPel
    Debug.Print d.dmPelsWidth
    Debug.Print d.dmPelsHeight
    Debug.Print d.dmDisplayFlags
    Debug.Print d.dmDisplayFrequency
    
    
    
    If Not IsNumeric(utxtOffsetX.Text) Or Not IsNumeric(utxtOffsetY.Text) Then Exit Sub
    
    d.dmPosition.X = Val(utxtOffsetX.Text)
    d.dmPosition.Y = Val(utxtOffsetY.Text)
    
    d.dmDisplayOrientation = udrpOrientation.ItemData(udrpOrientation.ListIndex)
    
    Dim tmpWidth As Long
    
    If (d.dmDisplayOrientation And 1) = 1 Then
        
        If d.dmPelsWidth > d.dmPelsHeight Then
            tmpWidth = d.dmPelsWidth
            d.dmPelsWidth = d.dmPelsHeight
            d.dmPelsHeight = tmpWidth
        End If
    Else
        If d.dmPelsWidth < d.dmPelsHeight Then
            tmpWidth = d.dmPelsWidth
            d.dmPelsWidth = d.dmPelsHeight
            d.dmPelsHeight = tmpWidth
        End If
    End If
    
    
    d.dmFields = 544997536 'DM_POSITION Or DM_DISPLAYORIENTATION Or DM_PELSHEIGHT Or DM_PELSWIDTH Or DM_DISPLAYFLAGS Or DM_DISPLAYFREQUENCY Or DM_BITSPERPEL Or DM_DISPLAYFIXEDOUTPUT
    
    'Debug.Print d.dmPelsWidth & "x" & d.dmPelsHeight & " @ " & d.dmDisplayFrequency & " " & d.dmDisplayFixedOutput
    

    
    
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
    Dim res As Long
    
    res = ChangeDisplaySettingsEx(DeviceName, dev, 0, CDS_UPDATEREGISTRY Or CDS_FORCE, 0)
    
    Select Case res
        Case DISP_CHANGE_SUCCESSFUL
            Debug.Print DeviceName & " succeeded"
            setResolution = 0
        Case DISP_CHANGE_RESTART
            Debug.Print DeviceName & " needs a restart"
            setResolution = 1
        Case Else
            Debug.Print DeviceName & " could not change. Error: " & res
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
    Dim d As devMode
    
    On Error GoTo endit:
    
    For i = 0 To UBound(savedResolution(Index).monitorSaveData)
        DeviceName = TrimNull(StrConv(savedResolution(Index).monitorSaveData(i).data.DeviceName, vbUnicode))
        
        setResolution DeviceName, savedResolution(Index).monitorSaveData(i).displayResolutionCurrent
        
        d = savedResolution(Index).monitorSaveData(i).displayResolutionCurrent
        
        If DeviceName = "\\.\DISPLAY2" Then
            Debug.Print d.dmDeviceName
            Debug.Print d.dmSpecVersion
            Debug.Print d.dmDriverVersion
            Debug.Print d.dmSize
            Debug.Print d.dmDriverExtra
            Debug.Print d.dmFields
            Debug.Print d.dmPosition.X
            Debug.Print d.dmPosition.Y
            Debug.Print d.dmDisplayOrientation
            Debug.Print d.dmDisplayFixedOutput
            Debug.Print d.dmColor
            Debug.Print d.dmDuplex
            Debug.Print d.dmYResolution
            Debug.Print d.dmTTOption
            Debug.Print d.dmCollate
            Debug.Print d.dmFormName
            Debug.Print d.dmLogPixels
            Debug.Print d.dmBitsPerPel
            Debug.Print d.dmPelsWidth
            Debug.Print d.dmPelsHeight
            Debug.Print d.dmDisplayFlags
            Debug.Print d.dmDisplayFrequency
        End If
        
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

Private Sub uchkStartWithWindows_ActivateNextState(u_Cancel As Boolean, u_NewState As uCheckboxConstants)
    If u_NewState = u_unChecked Then
        u_NewState = u_Checked
    Else
        u_NewState = u_unChecked
    End If
    
    
    SetRunAtStartup "ScreenResizer", App.Path, u_NewState = u_Checked
End Sub

Private Sub udrpMonitors_ItemChange(itemIndex As Long)
    Dim i As Long
    Dim doubles As String
    Dim singles As String
    
    If itemIndex = -1 Then Exit Sub
    
    With monitors(itemIndex).displayResolutionCurrent
        lblInfo(0).Caption = "Hardware: " & TrimNull(StrConv(monitors(itemIndex).data.DeviceID, vbUnicode))
        lblInfo(1).Caption = "Colors: " & .dmBitsPerPel & "-bits"
        lblInfo(2).Caption = "VideoCard: " & TrimNull(StrConv(monitors(itemIndex).data.DeviceString, vbUnicode))
        lblInfo(3).Caption = "Offset: x:" & .dmPosition.X & " y:" & .dmPosition.Y
        
        If .dmPosition.X = 0 And .dmPosition.Y = 0 Then
            ubtnSetPrimary.Enabled = False
        Else
            ubtnSetPrimary.Enabled = True
        End If
    End With
    
    
    
    
    calculateDisplayPositions -1, True, itemIndex
    
    udrpRefreshRate.Clear
    udrpResolution.Clear
    udrpResolution.RedrawPause
    
    
    Dim selectionIndex As Long
    selectionIndex = -1
    
    For i = 0 To UBound(monitors(itemIndex).displayResolutions)
        With monitors(itemIndex).displayResolutions(i)
            singles = .dmPelsWidth & "x" & .dmPelsHeight & " "
            If InStr(1, doubles, singles) = 0 Then
                If monitors(itemIndex).displayResolutionCurrent.dmPelsWidth = .dmPelsWidth And monitors(itemIndex).displayResolutionCurrent.dmPelsWidth = .dmPelsWidth Then
                    selectionIndex = udrpResolution.AddItem(singles, i)
                Else
                    udrpResolution.AddItem singles, i
                End If
                
                doubles = doubles & singles
            End If
            
        End With
    Next i
    
    udrpResolution.RedrawResume
    
    utxtOffsetX.Text = monitors(itemIndex).displayResolutionCurrent.dmPosition.X
    utxtOffsetY.Text = monitors(itemIndex).displayResolutionCurrent.dmPosition.Y
    
    udrpOrientation.ListIndex = monitors(itemIndex).displayResolutionCurrent.dmDisplayOrientation
    
    If selectionIndex <> -1 Then udrpResolution.ListIndex = selectionIndex
    
End Sub

Private Sub udrpResolution_ItemChange(itemIndex As Long)
    If itemIndex = -1 Then Exit Sub
    
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
    
    selectedDevMode = monitors(lCurrent).displayResolutions(udrpResolution.ItemData(itemIndex))
    
    
    
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

Private Sub ulstSavedMonitors_ItemChange(itemIndex As Long)
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










Private Sub ulstSavedMonitors_MouseEnter()
    picPreviewBackground.Visible = True
End Sub

Private Sub ulstSavedMonitors_MouseLeave()
    picPreviewBackground.Visible = False
End Sub

Private Sub ulstSavedMonitors_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, itemIndex As Long)
    Static isMovingOver As Boolean
    If isMovingOver Then Exit Sub
    
    isMovingOver = True
    
    If itemIndex = -1 Then
        ulstSavedMonitors_MouseLeave
        isMovingOver = False
        Exit Sub
    Else
        ulstSavedMonitors_MouseEnter
    End If
    
    Dim newTop As Long
    newTop = Y + ulstSavedMonitors.Top - picPreviewBackground.Height / 2
    If newTop < ulstSavedMonitors.Top + 1 Then newTop = ulstSavedMonitors.Top + 1
    
    If newTop > ulstSavedMonitors.Top + picPreviewBackground.Height - 1 Then newTop = ulstSavedMonitors.Top + picPreviewBackground.Height - 1
    
    picPreviewBackground.Top = newTop
    
    
    calculateDisplayPositions itemIndex, True
    
    DoEvents
    isMovingOver = False
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
