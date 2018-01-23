VERSION 5.00
Begin VB.Form frmIdentify 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7710
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
   ScaleHeight     =   447
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrVanish 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2355
      Top             =   2295
   End
End
Attribute VB_Name = "frmIdentify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hwnd As Long, _
                ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
                
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal crKey As Long, _
                ByVal bAlpha As Byte, _
                ByVal dwFlags As Long) As Long
 
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2



Private Sub Form_Click()
    Me.Visible = False
    tmrVanish.Enabled = False
End Sub

Sub setColorHack()
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, vbCyan, 0&, LWA_COLORKEY
End Sub

Sub customShow(index As Long, positionX As Long, positionY As Long, width As Long, height As Long)
    Me.Visible = False
    tmrVanish.Enabled = False
    Me.Picture = LoadPicture
    
    Me.BackColor = vbCyan
    
    Me.Left = Screen.TwipsPerPixelX * positionX
    Me.Top = Screen.TwipsPerPixelY * positionY
    Me.width = Screen.TwipsPerPixelX * width
    Me.height = Screen.TwipsPerPixelY * height
    
    Dim xText As Long
    Dim yText As Long
    
    Dim tx As Long
    Dim ty As Long
    
    Me.FontSize = 350 'Fix(Me.ScaleHeight / 2)
    xText = Me.ScaleWidth / 2 - Me.TextWidth(index & "") / 2
    yText = Me.ScaleHeight / 2 - Me.TextHeight(index & "") / 2 - 1
    
    For tx = -3 To 3
        For ty = -3 To 3
            If tx <> 0 And ty <> 0 Then
                Me.CurrentX = xText + tx
                Me.CurrentY = yText + ty
                Me.ForeColor = &H80FF&
                Me.Print index & ""
            End If
        Next ty
    Next tx
    
    Me.CurrentX = xText
    Me.CurrentY = yText
    Me.ForeColor = vbWhite
    Me.Print index & ""
    
    
    
    SetTopMostWindow Me.hwnd, True
    
    setColorHack
    tmrVanish.Enabled = True
    Me.Visible = True
End Sub

Private Sub Label1_Click()

End Sub

Private Sub tmrVanish_Timer()
    Me.Visible = False
    tmrVanish.Enabled = False
End Sub
