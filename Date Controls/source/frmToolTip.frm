VERSION 5.00
Begin VB.Form frmToolTip 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   ForeColor       =   &H80000017&
   LinkTopic       =   "Form1"
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   258
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type SIZE
    cx As Long
    cy As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const HWND_TOP = 0
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40

Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_CALCRECT = &H400

Private Const BDR_RAISEDINNER = &H4
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const DEF_FIRST_DELAY = 1
Private Const DEF_TIP_DELAY = 4

Private mstrTipText             As String
Private mlngTmr                 As Long
Private mblnTipVisible          As Boolean
Private WithEvents mclsTimer    As cTimer
Attribute mclsTimer.VB_VarHelpID = -1

Public Sub DisplayToolTip(strToolTip As String)
    Dim udtSize As SIZE
    
    '// intialize timer object, if it is not done
    If mclsTimer Is Nothing Then Set mclsTimer = New cTimer
    '// save tip text
    mstrTipText = strToolTip
    '// set timer interval
    mclsTimer.Interval = 500
    '// enable it
    mclsTimer.Enabled = True
    '// save start timer
    mlngTmr = VBA.Timer
End Sub

Public Sub HideToolTip()
    mclsTimer.Enabled = False
    mblnTipVisible = False
    Me.Hide
End Sub

Private Sub Form_Terminate()
    If Not mclsTimer Is Nothing Then
        Set mclsTimer = Nothing
    End If
End Sub

Private Sub mclsTimer_Tick()
    If Not mblnTipVisible Then
        If (VBA.Timer >= (mlngTmr + DEF_FIRST_DELAY)) Then
            Call DisplayForm
        End If
    Else
        If (VBA.Timer >= (mlngTmr + (DEF_FIRST_DELAY + DEF_TIP_DELAY))) Then
            Call HideToolTip
        End If
    End If
End Sub

Private Sub DisplayForm()
    Dim udtPt   As POINTAPI
    Dim udtSize As SIZE
    Dim udtRct  As RECT
    
    '// get current mouse position
    Call GetCursorPos(udtPt)
    '// set an arbitrary rectangle
    Call SetRect(udtRct, 0, 0, Me.ScaleWidth, Me.ScaleHeight)
    '// calculate the height and width need to print the tooltip text
    Call DrawText(Me.hdc, mstrTipText, CLng(Len(mstrTipText)), udtRct, DT_CALCRECT)
    '// reset rectangle to fit the tooltip text
    Call SetRect(udtRct, 0, 0, udtRct.Right + 8, udtRct.Bottom + 6)
    '// print the text in the form
    Call DrawText(Me.hdc, mstrTipText, Len(mstrTipText), udtRct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
    '// draw the edge of the form little bit raised
    Call DrawEdge(Me.hdc, udtRct, BDR_RAISEDINNER, BF_RECT)
    '// position the form and show it
    Call SetWindowPos(Me.hwnd, HWND_TOP, (udtPt.x + 2), (udtPt.y + 18), _
                    (udtRct.Right - udtRct.Left), (udtRct.Bottom - udtRct.Top), _
                    SWP_NOACTIVATE Or SWP_SHOWWINDOW)
    
    mblnTipVisible = True
End Sub
