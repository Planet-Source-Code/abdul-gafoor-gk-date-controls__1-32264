VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMemoryDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------
' cMemoryDC.cls
'----------------------------------------------------------------------
' Implementation file for cMemoryDC class
' This class represents an off screen DC that is useful
' for enabling flick-free and smooth repainting of things
' like controls.
'----------------------------------------------------------------------
' Copyright (c) 1996, Microsoft Corporation
'----------------------------------------------------------------------
Option Explicit

'// ======================================================================
'// API Declares
'// ======================================================================
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal lngWidth As Long, ByVal lngHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal lngWidth As Long, ByVal lngHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lngWidth As Long, ByVal lngHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal lngWidth As Long, ByVal lngHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function InvertRectAPI Lib "user32" Alias "InvertRect" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function CreatePatternBrush& Lib "gdi32" (ByVal hBitmap As Long)
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As Long, ByVal lpOutput As Long, lpInitData As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetCurrentPositionEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function TextOutAPI Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, ByVal lpDx As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Private Declare Function LineToAPI Lib "gdi32" Alias "LineTo" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function RectangleAPI Lib "gdi32" Alias "Rectangle" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function FillRectAPI Lib "user32" Alias "FillRect" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal lngWidth As Long, ByVal crColor As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DrawFocusRectAPI Lib "user32" Alias "DrawFocusRect" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function DrawTextAPI Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal ColorIn As Long, ByVal hPal As Long, ByRef RGBColorOut As Long)
#If UNICODE Then
    Private Declare Function DrawState Lib "user32" Alias "DrawStateW" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
#Else
    Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
#End If

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type SIZE
    cx As Long
    cy As Long
End Type

'// ======================================================================
'// API Text alignment
'// ======================================================================
Private Const TA_BASELINE& = 24
Private Const TA_BOTTOM& = 8
Private Const TA_CENTER& = 6
Private Const TA_LEFT& = 0
Private Const TA_NOUPDATECP& = 0
Private Const TA_UPDATECP& = 1
Private Const TA_RIGHT& = 2
Private Const TA_TOP& = 0

'// ======================================================================
'// API Pen Constants
'// ======================================================================
Private Const PS_SOLID& = 0

'// ======================================================================
'// API Color Constants
'// ======================================================================
Private Const COLOR_SYSTEM = &H80000000
Private Const COLOR_SYSTEM_INDEX = &HFF
Private Const COLOR_BTNFACE& = 15
Private Const COLOR_BTNHIGHLIGHT& = 20
Private Const COLOR_BTNHILIGHT& = COLOR_BTNHIGHLIGHT
Private Const COLOR_BTNSHADOW& = 16
Private Const COLOR_BTNTEXT& = 18
Private Const COLOR_3DDKSHADOW& = 21
Private Const COLOR_3DFACE& = COLOR_BTNFACE
Private Const COLOR_3DHIGHLIGHT& = COLOR_BTNHIGHLIGHT
Private Const COLOR_3DHILIGHT& = COLOR_BTNHIGHLIGHT
Private Const COLOR_3DLIGHT& = 22
Private Const COLOR_3DSHADOW& = COLOR_BTNSHADOW

'// ======================================================================
'// API DrawEdge Constants
'// ======================================================================
Private Const BORDER_WIDTH As Long = 0
Private Const BDR_INNER& = &HC
Private Const BDR_OUTER& = &H3
Private Const BDR_RAISED& = &H5
Private Const BDR_RAISEDINNER& = &H4
Private Const BDR_RAISEDOUTER& = &H1
Private Const BDR_SUNKEN& = &HA
Private Const BDR_SUNKENINNER& = &H8
Private Const BDR_SUNKENOUTER& = &H2
Private Const EDGE_BUMP& = &H9&
Private Const EDGE_ETCHED& = &H6&
Private Const EDGE_RAISED& = &H5&
Private Const EDGE_SUNKEN& = &HA&
Private Const BF_ADJUST& = &H2000
Private Const BF_BOTTOM& = &H8
Private Const BF_BOTTOMLEFT& = &H9
Private Const BF_BOTTOMRIGHT& = &HC
Private Const BF_DIAGONAL& = &H10
Private Const BF_FLAT& = &H4000
Private Const BF_LEFT& = &H1
Private Const BF_MIDDLE& = &H800
Private Const BF_MONO& = &H8000&
Private Const BF_RECT& = &HF
Private Const BF_RIGHT& = &H4
Private Const BF_SOFT& = &H1000
Private Const BF_TOP& = &H2
Private Const BF_TOPLEFT& = &H3
Private Const BF_TOPRIGHT& = &H6

'// ======================================================================
'// API DrawFrameControl Constants
'// ======================================================================
Private Const DFC_BUTTON& = 4
Private Const DFC_CAPTION& = 1
Private Const DFC_MENU& = 2
Private Const DFC_SCROLL& = 3
Private Const DFCS_ADJUSTRECT = &H2000
Private Const DFCS_BUTTON3STATE& = &H8
Private Const DFCS_BUTTONCHECK& = &H0
Private Const DFCS_BUTTONPUSH = &H10
Private Const DFCS_BUTTONRADIO& = &H4
Private Const DFCS_BUTTONRADIOIMAGE& = &H1
Private Const DFCS_BUTTONRADIOMASK& = &H2
Private Const DFCS_CAPTIONCLOSE& = &H0
Private Const DFCS_CAPTIONHELP& = &H4
Private Const DFCS_CAPTIONMAX& = &H2
Private Const DFCS_CAPTIONMIN& = &H1
Private Const DFCS_CAPTIONRESTORE& = &H3
Private Const DFCS_CHECKED& = &H400
Private Const DFCS_FLAT& = &H4000
Private Const DFCS_INACTIVE& = &H100
Private Const DFCS_MENUARROW& = &H0
Private Const DFCS_MENUARROWRIGHT& = &H4
Private Const DFCS_MENUBULLET& = &H2
Private Const DFCS_MENUCHECK& = &H1
Private Const DFCS_MONO& = &H8000&
Private Const DFCS_PUSHED& = &H200
Private Const DFCS_SCROLLCOMBOBOX& = &H5
Private Const DFCS_SCROLLDOWN& = &H1
Private Const DFCS_SCROLLLEFT& = &H2
Private Const DFCS_SCROLLRIGHT& = &H3
Private Const DFCS_SCROLLSIZEGRIP& = &H8
Private Const DFCS_SCROLLSIZEGRIPRIGHT& = &H10
Private Const DFCS_SCROLLUP& = &H0

'// ======================================================================
'// API DrawText Constants
'// ======================================================================
Private Const DT_BOTTOM& = &H8
Private Const DT_CALCRECT& = &H400
Private Const DT_CENTER& = &H1
Private Const DT_EXPANDTABS& = &H40
Private Const DT_EXTERNALLEADING& = &H200
Private Const DT_LEFT& = &H0
Private Const DT_NOCLIP& = &H100
Private Const DT_NOPREFIX& = &H800
Private Const DT_RIGHT& = &H2
Private Const DT_SINGLELINE& = &H20
Private Const DT_TOP& = &H0
Private Const DT_VCENTER& = &H4
Private Const DT_WORDBREAK& = &H10

Private Const OPAQUE& = 2
Private Const SRCAND& = &H8800C6
Private Const SRCCOPY& = &HCC0020
Private Const SRCERASE& = &H440328
Private Const SRCINVERT& = &H660046
Private Const DSTINVERT& = &H550009
Private Const SRCPAINT& = &HEE0086
Private Const COLOR_HIGHLIGHT& = 13
Private Const WHITENESS& = &HFF0062
Private Const BLACKNESS& = &H42

Private Const ETO_OPAQUE = 2
Private Const ETO_CLIPPED = 4

Private Const COLOR_SCROLLBAR = 0
Private Const COLOR_BACKGROUND = 1
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_MENU = 4
Private Const COLOR_WINDOW = 5
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_WINDOWTEXT = 8
Private Const COLOR_CAPTIONTEXT = 9
Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_INACTIVEBORDER = 11
Private Const COLOR_APPWORKSPACE = 12
'Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
'Private Const COLOR_BTNFACE = 15
'Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_GRAYTEXT = 17
'Private Const COLOR_BTNTEXT = 18
Private Const COLOR_INACTIVECAPTIONTEXT = 19
'Private Const COLOR_BTNHIGHLIGHT = 20

'// ======================================================================
'// Class Constants
'// ======================================================================
Private Const m_lSUNKEN_OFFSET = 1

'// ======================================================================
'// Button States
'// ======================================================================
Private Const giFLATTENED = 0
Private Const giRAISED = 1
Private Const giSUNKEN = 3
Private Const giDISABLED = 4

'// ======================================================================
'// VB Errors
'// ======================================================================
Private Const giOBJECT_VARIABLE_NOT_SET = 91
Private Const giINVALID_PICTURE = 481
Private Const giDLL_FUNCTION_NOT_FOUND = 453

'// ======================================================================
'// Public Enumerations
'// ======================================================================
Public Enum mdcTextAlignment
    mdcBaseLine = TA_BASELINE
    mdcBottom = TA_BOTTOM
    mdcCenter = TA_CENTER
    mdcLeft = TA_LEFT
    mdcRight = TA_RIGHT
    mdcTop = TA_TOP
    mdcNoCursorPositionUpdate = TA_NOUPDATECP
    mdcUpdateCursorPosition = TA_UPDATECP
End Enum

Public Enum mdcBorderStyle
    'mdcDropShadow = 0
    mdcBump = 1
    mdcEtched = 2
    mdcRaised = 3
    mdcSunken = 4
    mdcThinRaised = 5
    mdcThinSunken = 6
End Enum

Public Enum mdcComponentType
    mdcButton = DFC_BUTTON
    mdcCaption = DFC_CAPTION
    mdcMenu = DFC_MENU
    mdcScroll = DFC_SCROLL
End Enum

Public Enum mdcComponentStyle
    mdcButton3State = DFCS_BUTTON3STATE
    mdcButtonCheck = DFCS_BUTTONCHECK
    mdcButtonPush = DFCS_BUTTONPUSH
    mdcButtonRadio = DFCS_BUTTONRADIO
    mdcCaptionClose = DFCS_CAPTIONCLOSE
    mdcCaptionHelp = DFCS_CAPTIONHELP
    mdcCaptionMax = DFCS_CAPTIONMAX
    mdcCaptionMin = DFCS_CAPTIONMIN
    mdcCaptionRestore = DFCS_CAPTIONRESTORE
    mdcMenuArrow = DFCS_MENUARROW
    mdcMenuBullet = DFCS_MENUBULLET
    mdcMenuCheck = DFCS_MENUCHECK
    mdcScrollComboBox = DFCS_SCROLLCOMBOBOX
    mdcScrollDown = DFCS_SCROLLDOWN
    mdcScrollleft = DFCS_SCROLLLEFT
    mdcScrollright = DFCS_SCROLLRIGHT
    mdcScrollSizeGrip = DFCS_SCROLLSIZEGRIP
    mdcScrollUp = DFCS_SCROLLUP
End Enum

Public Enum mdcComponentState
    mdcStateNone = 0
    mdcStateChecked = DFCS_CHECKED
    mdcStateFlat = DFCS_FLAT
    mdcStateInactive = DFCS_INACTIVE
    mdcStateMono = DFCS_MONO
    mdcStatePushed = DFCS_PUSHED
End Enum

Public Enum mdcTextFormatting
    mdcTextAlignBottom = DT_BOTTOM
    mdcTextCalRectangle = DT_CALCRECT
    mdcTextMultiLineCenter = DT_CENTER
    mdcTextExpandTabs = DT_EXPANDTABS
    mdcTextExternalLeading = DT_EXTERNALLEADING
    mdcTextAlignLeft = DT_LEFT
    mdcTextNoClipping = DT_NOCLIP
    mdcTextNoPrefix = DT_NOPREFIX
    mdcTextAlignRight = DT_RIGHT
    mdcTextSingleLine = DT_SINGLELINE
    mdcTextAlignTop = DT_TOP
    mdcTextSingleLineCenter = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    mdcTextWordBreak = DT_WORDBREAK
End Enum

Public Enum bmpBorderStates
    'Button States
    bmpFLATTENED = 0
    bmpRAISED = 1
    bmpSUNKEN = 3
    bmpDISABLED = 4
End Enum

Public Enum mdcCaptionAlignments
    mdcCenterCenter
    mdcCenterLeft
    mdcCenterRight
    mdcTopCenter
    mdcTopLeft
    mdcTopRight
    mdcBottomCenter
    mdcBottomLeft
    mdcBottomright
End Enum

Public Enum mdcAppearances
    mdcappFlat = 0
    mdcappSunken
    mdcappThin
    mdcappEtched
    mdcappRaised
    mdcappNoLines
    mdcappSelected
End Enum

'// ======================================================================
'//  Private Data Members
'// ======================================================================
Private mlngMemoryDC        As Long     '// Off-screen HDC
Private mlngActualDC        As Long     '// Actual HDC of the control
Private mlngOldBmp          As Long     '// hBmp of the old bitmap in the off-sceen DC
Private mfntCurrent         As IFont    '// Font to use when drawing text
Private mlngOldFont         As Long     '// hFont of the old font in the off-screen dc
Private mlngErrNumber       As Long
Private mlngLineColor       As Long

Private mlngActualWidth     As Long     '// Width of the control's surface
Private mlngActualHeight    As Long     '// Height of the control's surface

Private mlng3DFaceCol       As Long     '// Color to use for the 3d face
Private mlng3DHighlightCol  As Long     '// Color to use for the 3d highlight
Private mlng3DShadowCol     As Long     '// Color to use for the 3d shadow

Private Sub Class_Terminate()
'// ======================================================================
'// Purpose:  Called when the object is destroyed--do clean-up work
'// Inputs:   None
'// Outputs:  None
'// ======================================================================
    Dim hBmp As Long
    
    '// If our handles are NULL then just get out
    If (mlngMemoryDC <> 0) Then
        '// Select the old font back into the off-screen dc
        Call SelectObject(mlngMemoryDC, mlngOldFont)
        '// Select the old bitmap back into the off-screen DC
        hBmp = SelectObject(mlngMemoryDC, mlngOldBmp)
        '// Delete the bitmap we were using
        Call DeleteObject(hBmp)
        '// and now delete the off-screen DC to totally clean up
        Call DeleteDC(mlngMemoryDC)
    End If
End Sub

'// ======================================================================
'// Purpose:  To get and let the current background color of the DC
'// ======================================================================
Public Property Get BackColor() As OLE_COLOR
    '// Assert that we are initialized
    Debug.Assert (mlngMemoryDC <> 0)
    '// Return the current background color
    BackColor = GetBkColor(mlngMemoryDC)
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
    Dim lngColor As Long
    
    '// Assert that we are initialized
    Debug.Assert (mlngMemoryDC <> 0)
    '// Set the new background color
    Call OleTranslateColor(vData, 0&, lngColor)
    Call SetBkColor(mlngMemoryDC, lngColor)
End Property

'// ======================================================================
'// Purpose:  To get and let the current fore color of the DC
'// ======================================================================
Public Property Get ForeColor() As OLE_COLOR
    '// Assert that we are initialized
    Debug.Assert (mlngMemoryDC <> 0)
    '// Return the current Text color
    ForeColor = GetTextColor(mlngMemoryDC)
End Property

Public Property Let ForeColor(ByVal vData As OLE_COLOR)
    Dim lngColor As Long
    
    '// Assert that we are initialized
    Debug.Assert (mlngMemoryDC <> 0)
    '// Set the new text color
    Call OleTranslateColor(vData, 0&, lngColor)
    Call SetTextColor(mlngMemoryDC, lngColor)
End Property

'// ======================================================================
'//  Purpose:  To get and let the line color of the DC
'// ======================================================================
Public Property Get LineColor() As OLE_COLOR
    '// Assert that we are initialized
    Debug.Assert (mlngMemoryDC <> 0)
    '// Return the current line color
    LineColor = mlngLineColor
End Property

Public Property Let LineColor(ByVal vData As OLE_COLOR)
    Dim lngColor As Long
    
    '// Assert that we are initialized
    Debug.Assert (mlngMemoryDC <> 0)
    '// Set the new line color
    Call OleTranslateColor(vData, 0&, lngColor)
    mlngLineColor = lngColor
End Property

'// ======================================================================
'// Purpose:  To get and set the current font to use on the DC
'// ======================================================================
Public Property Get Font() As StdFont
    '// Just return the reference we currently are holding
    Set Font = mfntCurrent
End Property

Public Property Set Font(vData As StdFont)
    '// Make sure we're initialized first
    '// Must call Initialize before setting the font!
    Debug.Assert (mlngMemoryDC <> 0)
    
    '// Below we will set a local member variable equal to the
    '// object passed in.  Even though the type passed in is a
    '// StdFont, our member variable is of type IFont.  A StdFont
    '// can be casted (QI) to an IFont, and the IFont interface gives
    '// us access to the hFont property, which we need in order to
    '// set the current font of the off-screen device context.
    
    '// if this is the first time the font is being set,
    '// grab the existing hFont handle so we can restore it
    '// before deleting the DC
    If mfntCurrent Is Nothing Then
        Set mfntCurrent = vData
        mlngOldFont = SelectObject(mlngMemoryDC, mfntCurrent.hFont)
    Else
        Set mfntCurrent = vData
        '// If this is being set to Nothing, restore the old font
        If mfntCurrent Is Nothing Then
            Call SelectObject(mlngMemoryDC, mlngOldFont)
        Else
            Call SelectObject(mlngMemoryDC, mfntCurrent.hFont)
        End If
    End If
End Property

Public Sub StartDrawing( _
                lngHdc As Long, _
                lngWidth As Long, _
                lngHeight As Long, _
                Optional ByVal blnCopySourceDC As Boolean = False)
'// ======================================================================
'//  Purpose:  To initialize the object with the screen DC from which we
'//            will create the off-screen DC
'//  Inputs:   The user control
'//  Outputs:  none
'// ======================================================================
    Dim hBmp As Long
    
    '// Assert that the inputs are valid
    '// And that we haven't already called Initialize
    Debug.Assert (lngHdc <> 0)
    Debug.Assert (mlngMemoryDC = 0)
    '// Store the HDC of the control in our private variable
    mlngActualDC = lngHdc
    '// Capture the width and height of the control
    mlngActualWidth = lngWidth
    mlngActualHeight = lngHeight
    '// Create the off-sceen DC
    mlngMemoryDC = CreateCompatibleDC(mlngActualDC)
    '// Create a compatible bitmap from the control DC
    '// that is the same size as the control itself
    hBmp = CreateCompatibleBitmap(mlngActualDC, mlngActualWidth, mlngActualHeight)
    '// Select that new bitmap into the off-screen DC
    '// and save the old bitmap handle so we can reselect
    '// it back in before we destroy the off-screen DC
    mlngOldBmp = SelectObject(mlngMemoryDC, hBmp)
    
    Me.BackColor = GetBkColor(mlngActualDC)
    Me.ForeColor = GetTextColor(mlngActualDC)
    
    If blnCopySourceDC Then
        '// Use bitblt to blast the contents of the control to the memory DC
        Call BitBlt(mlngMemoryDC, 0, 0, mlngActualWidth, _
                    mlngActualHeight, mlngActualDC, 0, 0, SRCCOPY)
    Else
        '// Clear the Memory DC
        Call FillRect(0&, 0&, mlngActualWidth, mlngActualHeight, Me.BackColor)
    End If
End Sub

Public Sub StopDrawing( _
                Optional lngLeft As Long = 0, _
                Optional lngTop As Long = 0, _
                Optional lngWidth As Long = -1, _
                Optional lngHeight As Long = -1)
'// ======================================================================
'//  Purpose:  Blasts the contents of the off-screen DC to the control's
'              on-screen surface
'//  Inputs:   none
'//  Outputs:  none
'// ======================================================================
    If (lngWidth = -1) Then lngWidth = mlngActualWidth
    If (lngHeight = -1) Then lngHeight = mlngActualHeight
    '// Use bitblt to blast the contents of the off-screen dc to the control
    Call BitBlt(mlngActualDC, lngLeft, lngTop, lngWidth, lngHeight, _
                mlngMemoryDC, lngLeft, lngTop, SRCCOPY)
End Sub

Public Sub FillRect( _
            ByVal lngLeft As Long, _
            ByVal lngTop As Long, _
            ByVal lngWidth As Long, _
            ByVal lngHeight As Long, _
            ByVal lngColor As Long, _
            Optional strCaption As String = "", _
            Optional enuCaptionAlign As mdcCaptionAlignments = mdcCenterCenter)
'// ======================================================================
'//  Purpose:  To fill a rectangle on the off-screen DC with a specified
'//            color in a fast way
'//  Inputs:   The rectangle to fill and color to use
'//  Outputs:  none
'// ======================================================================
    Dim lngX            As Long     '// X for drawing caption text
    Dim lngY            As Long     '// Y for drawing caption text
    Dim udtR            As RECT     '// Rect struct to pass to the GDI
    Dim udtTextExtent   As SIZE     '// Pixel size of caption
    Dim lngOrgColor     As Long
    
    '// Assert that we've been initialized already
    Debug.Assert (mlngMemoryDC <> 0)
    '// and check the inputs
    Call OleTranslateColor(lngColor, 0&, lngColor)
    Debug.Assert (lngColor >= 0)
    '// Set the back color of the DC to the color desired
    lngOrgColor = Me.BackColor
    Me.BackColor = lngColor
    '// Calculate the caption X and Y (centered) if the caption
    '// is not an empty string
    If (Len(strCaption) > 0) Then
        '// Get the pixel width of the Caption
        Call GetTextExtentPoint32(mlngMemoryDC, strCaption, Len(strCaption), udtTextExtent)
        '// Determine the X value based on the alignment chosen
        Select Case enuCaptionAlign
            Case mdcCenterCenter, mdcTopCenter, mdcBottomCenter
                lngX = ((lngWidth - udtTextExtent.cx) \ 2) + lngLeft
            Case mdcCenterRight, mdcTopRight, mdcBottomright
                lngX = lngWidth - BORDER_WIDTH - udtTextExtent.cx + lngLeft
            Case mdcCenterLeft, mdcTopLeft, mdcBottomLeft
                lngX = lngLeft + BORDER_WIDTH
        End Select
        '// Determine the Y value base on the alignment chosen
        Select Case enuCaptionAlign
            Case mdcCenterCenter, mdcCenterRight, mdcCenterLeft
                lngY = ((lngHeight - udtTextExtent.cy) \ 2) + lngTop
            Case mdcTopCenter, mdcTopLeft, mdcTopRight
                lngY = lngTop + BORDER_WIDTH
            Case mdcBottomCenter, mdcBottomLeft, mdcBottomright
                lngY = lngHeight - BORDER_WIDTH - udtTextExtent.cy + lngTop
        End Select
    End If
    '// Assign the input values to the rect struct
    With udtR
        .Left = lngLeft
        .Top = lngTop
        .Right = lngWidth + lngLeft
        .Bottom = lngHeight + lngTop
    End With
    '// ExtTextOut is one of the fastest ways to fill a rectangular
    '// area on a DC and is used here to fill our rect
    Call ExtTextOut(mlngMemoryDC, lngX, lngY, ETO_OPAQUE + ETO_CLIPPED, _
                    udtR, strCaption, Len(strCaption), 0)
    
    Me.BackColor = lngOrgColor
End Sub

Public Sub Draw3DRect( _
                ByVal lngLeft As Long, _
                ByVal lngTop As Long, _
                ByVal lngWidth As Long, _
                ByVal lngHeight As Long, _
                Optional strCaption As String = "", _
                Optional enuCaptionAlign As mdcCaptionAlignments = mdcCenterCenter, _
                Optional enuAppearance As mdcAppearances = mdcappRaised)
'// ======================================================================
'//  Purpose:  To draw a 3D looking rectangle on the off-screen DC
'//  Inputs:   The rectangle to make 3d and optionally a caption to
'//            display centered in the rect
'//  Outputs:  none
'// ======================================================================
    Dim lngLowerRightCol    As Long     '// Color to use for the lower right
    Dim lngUpperLeftCol     As Long     '// Color to use for the upper left
    
    '// Assert that we've been initialized already
    '// and check the inputs
    Debug.Assert (mlngMemoryDC <> 0)
    '// If we haven't gotten the system colors for 3d effects
    '// get them first
    If (mlng3DFaceCol = 0) Then
        mlng3DFaceCol = GetSysColor(COLOR_BTNFACE)
        mlng3DHighlightCol = GetSysColor(COLOR_BTNHIGHLIGHT)
        mlng3DShadowCol = GetSysColor(COLOR_BTNSHADOW)
    End If
    '// Set the lower-right and upper-left colors based on the
    '// Desired enuAppearance
    Select Case enuAppearance
        Case mdcappFlat
            lngLowerRightCol = mlng3DShadowCol
            lngUpperLeftCol = mlng3DShadowCol
        Case mdcappRaised
            lngLowerRightCol = mlng3DShadowCol
            lngUpperLeftCol = mlng3DHighlightCol
        Case mdcappSunken
            lngLowerRightCol = mlng3DHighlightCol
            lngUpperLeftCol = mlng3DShadowCol
        Case mdcappSelected
            lngLowerRightCol = mlng3DHighlightCol
            lngUpperLeftCol = vbBlack
    End Select
    '// Fill the rect with the shadow color (or hightlight if sunken)
    If (enuAppearance <> mdcappNoLines) Then
        Call Me.FillRect(lngLeft, lngTop, lngWidth, lngHeight, lngLowerRightCol)
    End If
    '// Now pull the right and bottom edges in by 1 pixel
    lngWidth = lngWidth - 1
    lngHeight = lngHeight - 1
    '// Fill the rect with the 3d highlight color (or shadow if sunken)
    If (enuAppearance <> mdcappNoLines) Then
        Call Me.FillRect(lngLeft, lngTop, lngWidth, lngHeight, lngUpperLeftCol)
    End If
    '// Finally pull in the left and top edges by 1 pixel
    lngLeft = lngLeft + 1
    lngTop = lngTop + 1
    lngWidth = lngWidth - 1
    lngHeight = lngHeight - 1
    '// Change the color to the 3d face color
    '// and fill the rect passing the desired caption
    Call Me.FillRect(lngLeft, lngTop, lngWidth, lngHeight, _
                Me.BackColor, strCaption, enuCaptionAlign)
    '// If the Appearance setting was Selected, invert the rect
    If (enuAppearance = mdcappSelected) Then
        Call InvertRect(lngLeft, lngTop, lngWidth, lngHeight)
    End If
End Sub

Public Sub InvertRect( _
                lngLeft As Long, _
                lngTop As Long, _
                lngWidth As Long, _
                lngHeight As Long)
'// ======================================================================
'//  Purpose:  To invert a particular rect on the bitmap
'//  Inputs:   The area to invert
'//  Outputs:  none
'// ======================================================================
    Dim udtR    As RECT
    
    With udtR
        .Left = lngLeft
        .Top = lngTop
        .Right = lngLeft + lngWidth
        .Bottom = lngTop + lngHeight
    End With
    Call InvertRectAPI(mlngMemoryDC, udtR)
End Sub

Public Sub DrawBorder( _
                ByVal lngLeft As Long, _
                ByVal lngTop As Long, _
                ByVal lngWidth As Long, _
                ByVal lngHeight As Long, _
                ByVal enuStyle As mdcBorderStyle, _
                Optional ByVal blnFlat As Boolean = False, _
                Optional ByVal blnMono As Boolean = False, _
                Optional ByVal blnSoft As Boolean = False)
'// ======================================================================
'//  Purpose:  To draw a various kinds border around a rect
'//  Inputs:   The rect structure
'//  Outputs:  none
'// ======================================================================
    Dim udtR        As RECT
    Dim lngFlag     As Long
    Dim lngEdge     As Long
    
    With udtR
        .Left = lngLeft
        .Top = lngTop
        .Right = lngLeft + lngWidth
        .Bottom = lngTop + lngHeight
    End With
    
    Select Case enuStyle
        Case mdcBump: lngEdge = EDGE_BUMP
        Case mdcEtched: lngEdge = EDGE_ETCHED
        Case mdcSunken: lngEdge = EDGE_SUNKEN
        Case mdcRaised: lngEdge = EDGE_RAISED
        Case mdcThinRaised: lngEdge = BDR_SUNKENOUTER
        Case mdcThinSunken: lngEdge = BDR_RAISEDINNER
        Case Else: lngEdge = EDGE_SUNKEN
    End Select
    
    If blnSoft Then lngFlag = lngFlag + BF_SOFT
    If blnMono Then lngFlag = lngFlag + BF_MONO
    If blnFlat Then lngFlag = lngFlag + BF_FLAT
    '// modified by Abdul Gafoor.GK to remove softline when set as Flat
    If (lngFlag <> BF_FLAT) Then
        lngFlag = lngFlag Or BF_RECT
    End If
    
    mlngErrNumber = DrawEdge(mlngMemoryDC, udtR, lngEdge, lngFlag)
End Sub

Public Sub DrawFrame( _
            ByVal lngLeft As Long, _
            ByVal lngTop As Long, _
            ByVal lngWidth As Long, _
            ByVal lngHeight As Long, _
            Optional ByVal lngBorderColor As OLE_COLOR = -1)
'// ======================================================================
'//  Purpose:  To draw a frame around a rect
'//  Inputs:   The rect structure
'//  Outputs:  none
'// ======================================================================
    Dim udtR        As RECT
    Dim lngColor    As Long
    
    With udtR
        .Left = lngLeft
        .Top = lngTop
        .Right = lngLeft + lngWidth
        .Bottom = lngTop + lngHeight
    End With
    '// Create a brush from the color passed
    lngColor = CreateSolidBrush(lngBorderColor)
    '// Draw the frame
    mlngErrNumber = FrameRect(mlngMemoryDC, udtR, lngColor)
    '// Delete the brush object to free memory
    Call DeleteObject(lngColor)
End Sub

Public Function DrawText( _
                    ByVal lngLeft As Long, _
                    ByVal lngTop As Long, _
                    ByVal lngWidth As Long, _
                    ByVal lngHeight As Long, _
                    ByVal strText As String, _
                    Optional ByVal enuFormat As mdcTextFormatting, _
                    Optional ByVal lngForeColor As SystemColorConstants = -1) As Long
'// ======================================================================
'//  Purpose:  To draw a caption
'//  Inputs:   The rect structure and caption
'//  Outputs:  none
'// ======================================================================
    Dim lngFormat           As Long
    Dim lngCurrentTextColor As Long
    Dim udtR                As RECT
    
    With udtR
        .Left = lngLeft
        .Top = lngTop
        .Right = lngLeft + lngWidth
        .Bottom = lngTop + lngHeight
    End With
    
    lngFormat = enuFormat
    
    If (lngForeColor <> -1) Then
        lngCurrentTextColor = GetTextColor(mlngMemoryDC)
        If lngForeColor And COLOR_SYSTEM Then
            lngForeColor = GetSysColor(lngForeColor And COLOR_SYSTEM_INDEX)
        End If
        mlngErrNumber = SetTextColor(mlngMemoryDC, lngForeColor)
    End If
    
    DrawText = DrawTextAPI(mlngMemoryDC, strText, Len(strText), udtR, lngFormat)
    
    If Not IsMissing(lngCurrentTextColor) Then
        mlngErrNumber = SetTextColor(mlngMemoryDC, lngCurrentTextColor)
    End If
End Function


Public Function DrawTextDirect( _
                    hdc As Long, _
                    ByVal lngLeft As Long, _
                    ByVal lngTop As Long, _
                    ByVal lngWidth As Long, _
                    ByVal lngHeight As Long, _
                    ByVal strText As String, _
                    Optional ByVal enuFormat As mdcTextFormatting, _
                    Optional ByVal lngForeColor As SystemColorConstants = -1) As Long
'// ======================================================================
'//  Purpose:  To draw a caption direct in an device context
'//  Inputs:   The rect structure and caption
'//  Outputs:  none
'// ======================================================================
    Dim lngFormat           As Long
    Dim lngCurrentTextColor As Long
    Dim udtR                As RECT
    
    With udtR
        .Left = lngLeft
        .Top = lngTop
        .Right = lngLeft + lngWidth
        .Bottom = lngTop + lngHeight
    End With
    
    lngFormat = enuFormat
    
    If (lngForeColor <> -1) Then
        lngCurrentTextColor = GetTextColor(hdc)
        If lngForeColor And COLOR_SYSTEM Then
            lngForeColor = GetSysColor(lngForeColor And COLOR_SYSTEM_INDEX)
        End If
        mlngErrNumber = SetTextColor(hdc, lngForeColor)
    End If
    
    DrawTextDirect = DrawTextAPI(hdc, strText, Len(strText), udtR, lngFormat)
    
    If Not IsMissing(lngCurrentTextColor) Then
        mlngErrNumber = SetTextColor(hdc, lngCurrentTextColor)
    End If
End Function





''Set rgbColor to -1 if you wish the rectangle not to be filled in
'Public Sub FillRect(nLeft As Long, nTop As Long, nWidth As Long, nHeight As Long, rgbColor As Long, Optional sCaption As String = "", Optional CaptionAlign As Integer = caCenterCenter)
'
'    Dim nX As Long              'X for drawing caption text
'    Dim nY As Long              'Y for drawing caption text
'    Dim rc As RECT              'rect struct to pass to the GDI
'    Dim szTextExtent As SIZE    'pixel size of caption
'    Dim lColor As Long
'    Dim lOrgColor As Long
'
'    'assert that we've been initialized already
'    'and check the inputs
'    Debug.Assert mclsMemoryDC.hdc <> 0
'
'    If rgbColor < 0 Then
'        OleTranslateColor rgbColor, 0, lColor
'        rgbColor = lColor
'    End If
'
'    'set the back color of the DC to the color desired
'    lOrgColor = mclsMemoryDC.BackColor
'    mclsMemoryDC.BackColor = rgbColor
'
'    'calculate the caption X and Y (centered) if the caption
'    'is not an empty string
'    If Len(sCaption) > 0 Then
'
'        'get the pixel width of the Caption
'        GetTextExtentPoint32 mclsMemoryDC.hdc, sCaption, Len(sCaption), szTextExtent
'
'        'determine the X value based on the alignment chosen
'        Select Case CaptionAlign
'            Case caCenterCenter, caTopCenter, caBottomCenter
'                nX = ((nWidth - szTextExtent.cx) \ 2) + nLeft
'
'            Case caCenterRight, caTopRight, caBottomright
'                nX = nWidth - BORDER_WIDTH - szTextExtent.cx + nLeft
'
'            Case caCenterLeft, caTopLeft, caBottomLeft
'                nX = nLeft + BORDER_WIDTH
'        End Select
'
'        'determine the Y value base on the alignment chosen
'        Select Case CaptionAlign
'            Case caCenterCenter, caCenterRight, caCenterLeft
'                nY = ((nHeight - szTextExtent.cy) \ 2) + nTop
'
'            Case caTopCenter, caTopLeft, caTopRight
'                nY = nTop + BORDER_WIDTH
'
'            Case caBottomCenter, caBottomLeft, caBottomright
'                nY = nHeight - BORDER_WIDTH - szTextExtent.cy + nTop
'
'        End Select
'    End If 'caption is not ""
'
'    'assign the input values to the rect struct
'    rc.Left = nLeft
'    rc.Top = nTop
'    rc.Right = nWidth + nLeft
'    rc.Bottom = nHeight + nTop
'
'    'ExtTextOut is one of the fastest ways to fill a rectangular
'    'area on a DC and is used here to fill our rect
'    ExtTextOut mclsMemoryDC.hdc, nX, nY, ETO_OPAQUE + ETO_CLIPPED, rc, sCaption, Len(sCaption), 0
'
'    mclsMemoryDC.BackColor = lOrgColor
'
'End Sub 'FillRect()
