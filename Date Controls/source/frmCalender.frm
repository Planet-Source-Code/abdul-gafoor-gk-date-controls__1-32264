VERSION 5.00
Begin VB.Form frmCalendar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   148
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   191
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin DateControls.gkMonth clndr 
      Height          =   2220
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   3916
      Appearance      =   0
      BeginProperty DayHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DayHeaderFormat =   0
      ShowLines       =   0   'False
      PreMonthBackColor=   -2147483643
      PostMonthBackColor=   -2147483643
      TitleAppearance =   0
      TitleFontSize   =   8.25
      DayHeaderFontSize=   8.25
      CurrentMonthFontName=   "Arial"
      CurrentMonthFontSize=   8.25
      PreMonthFontName=   "Arial"
      PreMonthFontSize=   8.25
      PostMonthFontName=   "Arial"
      PostMonthFontSize=   8.25
      ActiveDayFontName=   "Arial"
      ActiveDayFontSize=   8.25
      ActiveDayFontItalic=   0   'False
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ================================================================================
'// Copyright © 2001-2002 by Abdul Gafoor.GK
'// ================================================================================
Option Explicit

Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private mblnCanceled As Boolean

Public Property Get Canceled() As Boolean
    Canceled = mblnCanceled
End Property

Public Sub ActivateForm(ByVal lpLeft As Long, ByVal lpTop As Long, _
                        ByVal lpRight As Long, ByVal lpBottom As Long)
    Dim lngLeft     As Long
    Dim lngTop      As Long

    lngLeft = lpLeft - 2
    lngTop = lpBottom + 2
    '// if calendar goes left off-screen, move it to left
    If (lngLeft < 0) Then lngLeft = 0
    '// if calendar goes right off-screen, move it to left
    If ((lngLeft + clndr.Width) > (Screen.Width / Screen.TwipsPerPixelX)) Then
        lngLeft = (Screen.Width / Screen.TwipsPerPixelX) - clndr.Width
    End If
    '// if form goes down off-screen, show it above the date picker
    If ((lngTop + clndr.Height) > (Screen.Height / Screen.TwipsPerPixelY)) Then
        lngTop = lngTop - (clndr.Height + (lpBottom - lpTop) + 4)
    End If
    
    '// position the form
    Call SetWindowPos(Me.hwnd, 0&, lngLeft, lngTop, clndr.Width, clndr.Height, SWP_NOZORDER Or SWP_FRAMECHANGED)
    '// capture all mouse movements (including off-window mouse movements)
    Call SetCapture(clndr.hwnd)
    '// finally show the form modally
    Me.Show vbModal
End Sub

Private Sub clndr_MenuClosed(MenuType As Integer)
    If (GetCapture() <> clndr.hwnd) Then
        Call SetCapture(clndr.hwnd)
    End If
End Sub

Private Sub clndr_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If IsMouseOnForm(x, y) Then
        If (Button = vbLeftButton) Then
            If (clndr.HitTest(CLng(x), CLng(y)) = dtCalendarDate) Then
                Call ReleaseCapture
                Call Form_KeyDown(vbKeyReturn, 0)
            End If
        End If
        '// since any mouse click will release capture, set mouse capture again
        Call SetCapture(clndr.hwnd)
    Else
        '// outside of the form.  release mouse capture and unload it
        Call ReleaseCapture
        Call Form_KeyDown(vbKeyEscape, 0)
    End If
End Sub

Private Sub clndr_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (GetCapture() <> clndr.hwnd) Then
        Call SetCapture(clndr.hwnd)
    End If
End Sub

Private Sub Form_Activate()
    clndr.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape: mblnCanceled = True
        Case vbKeyReturn: mblnCanceled = False
        Case Else: Exit Sub
    End Select
    
    Me.Hide
End Sub

Private Function IsMouseOnForm(ByVal x As Long, ByVal y As Long) As Boolean
    IsMouseOnForm = ((x >= 0) And (x <= clndr.Width)) And _
                    ((y >= 0) And (y <= clndr.Height))
End Function

Private Sub Form_Load()
    Me.KeyPreview = True
End Sub
