VERSION 5.00
Begin VB.UserControl gkMonth 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   PropertyPages   =   "gkMonth.ctx":0000
   ScaleHeight     =   157
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   181
   ToolboxBitmap   =   "gkMonth.ctx":0033
End
Attribute VB_Name = "gkMonth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'// ================================================================================
'// Copyright © 2001-2002 by Abdul Gafoor.GK
'// ================================================================================
Option Explicit

Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As Any) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Const MF_STRING = &H0
Private Const MF_BYCOMMAND = &H0
Private Const MF_BYPOSITION = &H400

Private Const TPM_NONOTIFY = &H80              '// Don't send any notification msgs
Private Const TPM_RETURNCMD = &H100

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

Private Type CellInformation
    lForeColor As Long
    lBackColor As Long
    tLocation As RECT
End Type


Public Enum dtAppearanceConstants
    dtFlat = 0
    dt3D
    dtThin
    dtEtched
End Enum

Public Enum dtBorderStyleConstants
    dtNone
    dtFixedSingle
End Enum

Public Enum dtDayHeaderFormats
    dtSingleLetter = 0
    dtMedium
    dtFullName
End Enum

Public Enum dtMonthConstants
    dtJanuary = 1
    dtFebruary
    dtMarch
    dtApril
    dtMay
    dtJune
    dtJuly
    dtAugust
    dtSeptember
    dtOctober
    dtNovember
    dtDecember
End Enum

Public Enum dtDaysOfTheWeek
    dtSunday = 1
    dtMonday
    dtTuesday
    dtWednesday
    dtThursday
    dtFriday
    dtSaturday
End Enum

Public Enum dtCalendarHitTestAreas
    dtCalendarDate
    dtInvalid
End Enum

Private Enum MenuType
    MonthMenu = 1
    YearMenu
End Enum

'// internal constants
Private Const DEF_CALENDAR_ROWS = 6
Private Const DEF_CALENDAR_COLS = 7
Private Const DEF_CELL_UBOUND = 5
Private Const DEF_TOP_MARGIN = 20                     'The top margin starting point for drawing the calendar. Also accounts for the comboboxes.
Private Const DEF_LEFT_MARGIN = 2                     'The left margin starting point for drawing the calendar
Private Const DEF_YEAR_WIDTH = 60
Private Const DEF_BUTTON_WIDTH = 10
Private Const DEF_YEAR_MENU_COUNT = 20

'// Internal variables
Private mudtCellInfo()                  As CellInformation      'Deminsioned as (m_nPeriodRows, DEF_CALENDAR_COLS)
Private mintCellWidth                   As Integer      'The width in pixels for a calendar cell
Private mintCellHeight                  As Integer      'The height in pixels for a calendar cell
Private mintCalLeft                     As Integer
Private mintCalTop                      As Integer
Private mintCalRight                    As Integer
Private mintCalBottom                   As Integer
Private mdtCalenderStart                As Date
Private mdtCalenderEnd                  As Date
Private mdtMonthStart                   As Date
Private mdtMonthEnd                     As Date
Private mintEdgeOffset                  As Integer
Private mintCurCellId                   As Integer
Private mblnToolTipVisible              As Boolean
Private mblnMenuVisible                 As Boolean
Private mfrmToolTip                     As frmToolTip
Private mlngMonthMenu                   As Long
Private mudtMonthArea                   As RECT
Private mlngYearMenu                    As Long
Private mudtYearArea                    As RECT

'// Default Property Values:
Private Const mdefShowToolTip = True
Private Const mdefDateTipFormat = "Dddd, mmmm dd, yyyy"
Private Const mdefMinDate As Date = "01/01/1900"
Private Const mdefMaxDate As Date = "31/12/2099"
Private Const mdefAutoRefresh = True
Private Const mdefShowDayHeader = True
Private Const mdefShowLines = True
Private Const mdefShowFocusRect = True
Private Const mdefActiveDayForeColor = SystemColorConstants.vbButtonText
Private Const mdefActiveDayBackColor = SystemColorConstants.vbWindowBackground
Private Const mdefPreMonthForeColor = SystemColorConstants.vbGrayText
Private Const mdefCurrentMonthForeColor = SystemColorConstants.vbButtonText
Private Const mdefPostMonthForeColor = SystemColorConstants.vbGrayText
Private Const mdefPreMonthBackColor = SystemColorConstants.vbInactiveBorder
Private Const mdefCurrentMonthBackColor = SystemColorConstants.vbWindowBackground
Private Const mdefPostMonthBackColor = SystemColorConstants.vbInactiveBorder
Private Const mdefAppearance = dt3D
Private Const mdefBorderStyle = dtFixedSingle
Private Const mdefLineColor = SystemColorConstants.vbButtonShadow
Private Const mdefDayHeaderFormat = dtMedium
Private Const mdefFirstDayOfWeek = dtSunday
Private Const mdefDayHeaderBackColor = SystemColorConstants.vbButtonFace
Private Const mdefDayHeaderForeColor = SystemColorConstants.vbButtonText
Private Const mdefTitleForeColor = SystemColorConstants.vbButtonText
Private Const mdefTitleBackColor = SystemColorConstants.vbButtonFace
Private Const mdefTitleAppearance = dtAppearanceConstants.dt3D

'// Property Variables:
Private mblnShowToolTip                 As Boolean
Private mstrDateTipFormat               As String
Private mdtCurrentDate                  As Date
Private mdtMinDate                      As Date
Private mdtMaxDate                      As Date
Private mblnAutoRefresh                 As Boolean
Private mblnShowLines                   As Boolean
Private mblnShowDayHeader               As Boolean
Private mblnShowFocusRect               As Boolean
Private mintCurrentYear                 As Integer
Private mfntTitleFont                   As New StdFont
Private mfntActiveDayFont               As New StdFont
Private mfntDayHeaderFont               As New StdFont
Private mfntCurrentMonthFont            As New StdFont
Private mfntPreMonthFont                As New StdFont
Private mfntPostMonthFont               As New StdFont
Private mlngActiveDayForeColor          As OLE_COLOR
Private mlngActiveDayBackColor          As OLE_COLOR
Private mlngPreMonthForeColor           As OLE_COLOR
Private mlngCurrentMonthForeColor       As OLE_COLOR
Private mlngPostMonthForeColor          As OLE_COLOR
Private mlngPreMonthBackColor           As OLE_COLOR
Private mlngCurrentMonthBackColor       As OLE_COLOR
Private mlngPostMonthBackColor          As OLE_COLOR
Private mlngLineColor                   As OLE_COLOR
Private mlngDayHeaderBackColor          As OLE_COLOR
Private mlngDayHeaderForeColor          As OLE_COLOR
Private mlngTitleForeColor              As OLE_COLOR
Private mlngTitleBackColor              As OLE_COLOR
Private menuAppearance                  As dtAppearanceConstants
Private menuBorderStyle                 As dtBorderStyleConstants
Private menuCurrentMonth                As dtMonthConstants
Private menuDayHeaderFormat             As dtDayHeaderFormats
Private menuFirstDayOfWeek              As dtDaysOfTheWeek
Private menuTitleAppearance             As dtAppearanceConstants

Public Event WillChangeDate(OldDate As Date, NewDate As Date, Cancel As Boolean)
Public Event DateChanged(OldDate As Date, NewDate As Date)
Public Event WillOpenMenu(MenuType As Integer, Cancel As Boolean)
Public Event MenuClosed(MenuType As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event Click()
Attribute Click.VB_MemberFlags = "200"
Public Event DblClick()

Public Property Get Appearance() As dtAppearanceConstants
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Appearance = menuAppearance
End Property

Public Property Let Appearance(ByVal vData As dtAppearanceConstants)
    menuAppearance = vData
    PropertyChanged "Appearance"
    '// Refresh edge offset
    mintEdgeOffset = GetEdgeOffset()
    '// Refresh control to reflect the changes
    Call Me.Refresh
End Property

Public Property Let BorderStyle(ByVal vData As dtBorderStyleConstants)
    menuBorderStyle = vData
    PropertyChanged "BorderStyle"
    '// Refresh edge offset
    mintEdgeOffset = GetEdgeOffset()
    '// Refresh control to reflect the changes
    If (menuAppearance = dtFlat) Then
        Call Me.Refresh
    End If
End Property

Public Property Get BorderStyle() As dtBorderStyleConstants
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = menuBorderStyle
End Property

Public Property Get AutoRefresh() As Boolean
Attribute AutoRefresh.VB_Description = "If true, control will be redrawn each time a property is set."
Attribute AutoRefresh.VB_ProcData.VB_Invoke_Property = ";Misc"
    AutoRefresh = mblnAutoRefresh
End Property

Public Property Let AutoRefresh(ByVal vData As Boolean)
    mblnAutoRefresh = vData
    PropertyChanged "AutoRefresh"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Let ShowDayHeader(ByVal vData As Boolean)
    mblnShowDayHeader = vData
    PropertyChanged "ShowDayHeader"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get ShowDayHeader() As Boolean
Attribute ShowDayHeader.VB_Description = "Trun on/off showing week day names."
    ShowDayHeader = mblnShowDayHeader
End Property

Public Property Set TitleFont(vData As StdFont)
Attribute TitleFont.VB_Description = "The font attributes that are used to display the title."
Attribute TitleFont.VB_ProcData.VB_Invoke_PropertyPutRef = "StandardFont;Font"
    Set mfntTitleFont = vData
    PropertyChanged "TitleFont"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get TitleFont() As StdFont
    Set TitleFont = mfntTitleFont
End Property

Public Property Get TitleFontName() As String
Attribute TitleFontName.VB_MemberFlags = "400"
    TitleFontName = mfntTitleFont.Name
End Property

Public Property Let TitleFontName(ByVal vData As String)
    mfntTitleFont.Name = vData
    PropertyChanged "TitleFontName"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get TitleFontSize() As Long
Attribute TitleFontSize.VB_MemberFlags = "400"
    TitleFontSize = mfntTitleFont.SIZE
End Property

Public Property Let TitleFontSize(ByVal vData As Long)
    mfntTitleFont.SIZE = vData
    PropertyChanged "TitleFontSize"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get TitleFontBold() As Boolean
Attribute TitleFontBold.VB_MemberFlags = "400"
    TitleFontBold = mfntTitleFont.Bold
End Property

Public Property Let TitleFontBold(ByVal vData As Boolean)
    mfntTitleFont.Bold = vData
    PropertyChanged "TitleFontBold"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get TitleFontItalic() As Boolean
Attribute TitleFontItalic.VB_MemberFlags = "400"
    TitleFontItalic = mfntTitleFont.Italic
End Property

Public Property Let TitleFontItalic(ByVal vData As Boolean)
    mfntTitleFont.Italic = vData
    PropertyChanged "TitleFontItalic"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get TitleFontUnderline() As Boolean
Attribute TitleFontUnderline.VB_MemberFlags = "400"
    TitleFontUnderline = mfntTitleFont.Underline
End Property

Public Property Let TitleFontUnderline(ByVal vData As Boolean)
    mfntTitleFont.Underline = vData
    PropertyChanged "TitleFontUnderline"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get TitleFontStrikethrough() As Boolean
Attribute TitleFontStrikethrough.VB_MemberFlags = "400"
    TitleFontStrikethrough = mfntTitleFont.Strikethrough
End Property

Public Property Let TitleFontStrikethrough(ByVal vData As Boolean)
    mfntTitleFont.Strikethrough = vData
    PropertyChanged "TitleFontStrikethrough"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get TitleForeColor() As OLE_COLOR
Attribute TitleForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TitleForeColor = mlngTitleForeColor
End Property

Public Property Let TitleForeColor(ByVal vData As OLE_COLOR)
    mlngTitleForeColor = vData
    PropertyChanged "TitleForeColor"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get TitleBackColor() As OLE_COLOR
Attribute TitleBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TitleBackColor = mlngTitleBackColor
End Property

Public Property Let TitleBackColor(ByVal vData As OLE_COLOR)
    mlngTitleBackColor = vData
    PropertyChanged "TitleBackColor"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Set DayHeaderFont(ByVal vData As StdFont)
Attribute DayHeaderFont.VB_Description = "The font attributes that are used to display the day of week header."
Attribute DayHeaderFont.VB_ProcData.VB_Invoke_PropertyPutRef = "StandardFont;Font"
    Set mfntDayHeaderFont = vData
    PropertyChanged "DayHeaderFont"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get DayHeaderFont() As StdFont
    Set DayHeaderFont = mfntDayHeaderFont
End Property

Public Property Get DayHeaderFontName() As String
Attribute DayHeaderFontName.VB_MemberFlags = "400"
    DayHeaderFontName = mfntDayHeaderFont.Name
End Property

Public Property Let DayHeaderFontName(ByVal vData As String)
    mfntDayHeaderFont.Name = vData
    PropertyChanged "DayHeaderFontName"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get DayHeaderFontSize() As Long
Attribute DayHeaderFontSize.VB_MemberFlags = "400"
    DayHeaderFontSize = mfntDayHeaderFont.SIZE
End Property

Public Property Let DayHeaderFontSize(ByVal vData As Long)
    mfntDayHeaderFont.SIZE = vData
    PropertyChanged "DayHeaderFontSize"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get DayHeaderFontBold() As Boolean
Attribute DayHeaderFontBold.VB_MemberFlags = "400"
    DayHeaderFontBold = mfntDayHeaderFont.Bold
End Property

Public Property Let DayHeaderFontBold(ByVal vData As Boolean)
    mfntDayHeaderFont.Bold = vData
    PropertyChanged "DayHeaderFontBold"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get DayHeaderFontItalic() As Boolean
Attribute DayHeaderFontItalic.VB_MemberFlags = "400"
    DayHeaderFontItalic = mfntDayHeaderFont.Italic
End Property

Public Property Let DayHeaderFontItalic(ByVal vData As Boolean)
    mfntDayHeaderFont.Italic = vData
    PropertyChanged "DayHeaderFontItalic"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get DayHeaderFontUnderline() As Boolean
Attribute DayHeaderFontUnderline.VB_MemberFlags = "400"
    DayHeaderFontUnderline = mfntDayHeaderFont.Underline
End Property

Public Property Let DayHeaderFontUnderline(ByVal vData As Boolean)
    mfntDayHeaderFont.Underline = vData
    PropertyChanged "DayHeaderFontUnderline"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get DayHeaderFontStrikethrough() As Boolean
Attribute DayHeaderFontStrikethrough.VB_MemberFlags = "400"
    DayHeaderFontStrikethrough = mfntDayHeaderFont.Strikethrough
End Property

Public Property Let DayHeaderFontStrikethrough(ByVal vData As Boolean)
    mfntDayHeaderFont.Strikethrough = vData
    PropertyChanged "DayHeaderFontStrikethrough"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get DayHeaderBackColor() As OLE_COLOR
Attribute DayHeaderBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DayHeaderBackColor = mlngDayHeaderBackColor
End Property

Public Property Let DayHeaderBackColor(ByVal vData As OLE_COLOR)
    mlngDayHeaderBackColor = vData
    PropertyChanged "DayHeaderBackColor"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get DayHeaderForeColor() As OLE_COLOR
Attribute DayHeaderForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DayHeaderForeColor = mlngDayHeaderForeColor
End Property

Public Property Let DayHeaderForeColor(ByVal vData As OLE_COLOR)
    mlngDayHeaderForeColor = vData
    PropertyChanged "DayHeaderForeColor"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get FirstDayOfWeek() As dtDaysOfTheWeek
Attribute FirstDayOfWeek.VB_Description = "Indicates which day of the week will be shown as first day of the week."
Attribute FirstDayOfWeek.VB_ProcData.VB_Invoke_Property = ";Behavior"
    FirstDayOfWeek = menuFirstDayOfWeek
End Property

Public Property Let FirstDayOfWeek(ByVal vData As dtDaysOfTheWeek)
    If (vData >= dtSunday) And (vData <= dtSaturday) Then
        menuFirstDayOfWeek = vData
        PropertyChanged "FirstDayOfWeek"
        Call PopulateDates
        If mblnAutoRefresh Then Call Me.Refresh
    End If
End Property

Public Property Get DayHeaderFormat() As dtDayHeaderFormats
Attribute DayHeaderFormat.VB_ProcData.VB_Invoke_Property = ";Misc"
    DayHeaderFormat = menuDayHeaderFormat
End Property

Public Property Let DayHeaderFormat(ByVal vData As dtDayHeaderFormats)
    menuDayHeaderFormat = vData
    PropertyChanged "DayHeaderFormat"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get ShowLines() As Boolean
Attribute ShowLines.VB_ProcData.VB_Invoke_Property = "Appearance"
    ShowLines = mblnShowLines
End Property

Public Property Let ShowLines(ByVal vData As Boolean)
    mblnShowLines = vData
    PropertyChanged "ShowLines"
    If (menuAppearance = dtFlat) Then mintEdgeOffset = GetEdgeOffset()
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get LineColor() As OLE_COLOR
Attribute LineColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    LineColor = mlngLineColor
End Property

Public Property Let LineColor(ByVal vData As OLE_COLOR)
    mlngLineColor = vData
    PropertyChanged "LineColor"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get CurrentMonthFont() As StdFont
Attribute CurrentMonthFont.VB_Description = "The font attributes used for displaying the days in the calendar grid\n."
Attribute CurrentMonthFont.VB_ProcData.VB_Invoke_Property = "StandardFont;Font"
    Set CurrentMonthFont = mfntCurrentMonthFont
End Property

Public Property Set CurrentMonthFont(ByVal vData As StdFont)
    Set mfntCurrentMonthFont = vData
    PropertyChanged "CurrentMonthFont"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get CurrentMonthFontName() As String
Attribute CurrentMonthFontName.VB_MemberFlags = "400"
    CurrentMonthFontName = mfntCurrentMonthFont.Name
End Property

Public Property Let CurrentMonthFontName(ByVal vData As String)
    mfntCurrentMonthFont.Name = vData
    PropertyChanged "CurrentMonthFontName"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get CurrentMonthFontSize() As Long
Attribute CurrentMonthFontSize.VB_MemberFlags = "400"
    CurrentMonthFontSize = mfntCurrentMonthFont.SIZE
End Property

Public Property Let CurrentMonthFontSize(ByVal vData As Long)
    mfntCurrentMonthFont.SIZE = vData
    PropertyChanged "CurrentMonthFontSize"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get CurrentMonthFontBold() As Boolean
Attribute CurrentMonthFontBold.VB_MemberFlags = "400"
    CurrentMonthFontBold = mfntCurrentMonthFont.Bold
End Property

Public Property Let CurrentMonthFontBold(ByVal vData As Boolean)
    mfntCurrentMonthFont.Bold = vData
    PropertyChanged "CurrentMonthFontBold"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get CurrentMonthFontItalic() As Boolean
Attribute CurrentMonthFontItalic.VB_MemberFlags = "400"
    CurrentMonthFontItalic = mfntCurrentMonthFont.Italic
End Property

Public Property Let CurrentMonthFontItalic(ByVal vData As Boolean)
    mfntCurrentMonthFont.Italic = vData
    PropertyChanged "CurrentMonthFontItalic"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get CurrentMonthFontUnderline() As Boolean
Attribute CurrentMonthFontUnderline.VB_MemberFlags = "400"
    CurrentMonthFontUnderline = mfntCurrentMonthFont.Underline
End Property

Public Property Let CurrentMonthFontUnderline(ByVal vData As Boolean)
    mfntCurrentMonthFont.Underline = vData
    PropertyChanged "CurrentMonthFontUnderline"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get CurrentMonthFontStrikethrough() As Boolean
Attribute CurrentMonthFontStrikethrough.VB_MemberFlags = "400"
    CurrentMonthFontStrikethrough = mfntCurrentMonthFont.Strikethrough
End Property

Public Property Let CurrentMonthFontStrikethrough(ByVal vData As Boolean)
    mfntCurrentMonthFont.Strikethrough = vData
    PropertyChanged "CurrentMonthFontStrikethrough"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PreMonthFont() As StdFont
Attribute PreMonthFont.VB_ProcData.VB_Invoke_Property = "StandardFont;Font"
    Set PreMonthFont = mfntPreMonthFont
End Property

Public Property Set PreMonthFont(ByVal vData As StdFont)
    Set mfntPreMonthFont = vData
    PropertyChanged "PreMonthFont"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PreMonthFontName() As String
Attribute PreMonthFontName.VB_MemberFlags = "400"
    PreMonthFontName = mfntPreMonthFont.Name
End Property

Public Property Let PreMonthFontName(ByVal vData As String)
    mfntPreMonthFont.Name = vData
    PropertyChanged "PreMonthFontName"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PreMonthFontSize() As String
Attribute PreMonthFontSize.VB_MemberFlags = "400"
    PreMonthFontSize = mfntPreMonthFont.SIZE
End Property

Public Property Let PreMonthFontSize(ByVal vData As String)
    mfntPreMonthFont.SIZE = vData
    PropertyChanged "PreMonthFontSize"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PreMonthFontBold() As Boolean
Attribute PreMonthFontBold.VB_MemberFlags = "400"
    PreMonthFontBold = mfntPreMonthFont.Bold
End Property

Public Property Let PreMonthFontBold(ByVal vData As Boolean)
    mfntPreMonthFont.Bold = vData
    PropertyChanged "PreMonthFontBold"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PreMonthFontItalic() As Boolean
Attribute PreMonthFontItalic.VB_MemberFlags = "400"
    PreMonthFontItalic = mfntPreMonthFont.Italic
End Property

Public Property Let PreMonthFontItalic(ByVal vData As Boolean)
    mfntPreMonthFont.Italic = vData
    PropertyChanged "PreMonthFontItalic"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PreMonthFontUnderline() As Boolean
Attribute PreMonthFontUnderline.VB_MemberFlags = "400"
    PreMonthFontUnderline = mfntPreMonthFont.Underline
End Property

Public Property Let PreMonthFontUnderline(ByVal vData As Boolean)
    mfntPreMonthFont.Underline = vData
    PropertyChanged "PreMonthFontUnderline"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PreMonthFontStrikethrough() As Boolean
Attribute PreMonthFontStrikethrough.VB_MemberFlags = "400"
    PreMonthFontStrikethrough = mfntPreMonthFont.Strikethrough
End Property

Public Property Let PreMonthFontStrikethrough(ByVal vData As Boolean)
    mfntPreMonthFont.Strikethrough = vData
    PropertyChanged "PreMonthFontStrikethrough"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PostMonthFont() As StdFont
Attribute PostMonthFont.VB_ProcData.VB_Invoke_Property = "StandardFont;Font"
    Set PostMonthFont = mfntPostMonthFont
End Property

Public Property Set PostMonthFont(ByVal vData As StdFont)
    Set mfntPostMonthFont = vData
    PropertyChanged "PostMonthFont"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PostMonthFontName() As String
Attribute PostMonthFontName.VB_MemberFlags = "400"
    PostMonthFontName = mfntPostMonthFont.Name
End Property

Public Property Let PostMonthFontName(ByVal vData As String)
    mfntPostMonthFont.Name = vData
    PropertyChanged "PostMonthFontName"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PostMonthFontSize() As String
Attribute PostMonthFontSize.VB_MemberFlags = "400"
    PostMonthFontSize = mfntPostMonthFont.SIZE
End Property

Public Property Let PostMonthFontSize(ByVal vData As String)
    mfntPostMonthFont.SIZE = vData
    PropertyChanged "PostMonthFontSize"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PostMonthFontBold() As Boolean
Attribute PostMonthFontBold.VB_MemberFlags = "400"
    PostMonthFontBold = mfntPostMonthFont.Bold
End Property

Public Property Let PostMonthFontBold(ByVal vData As Boolean)
    mfntPostMonthFont.Bold = vData
    PropertyChanged "PostMonthFontBold"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PostMonthFontItalic() As Boolean
Attribute PostMonthFontItalic.VB_MemberFlags = "400"
    PostMonthFontItalic = mfntPostMonthFont.Italic
End Property

Public Property Let PostMonthFontItalic(ByVal vData As Boolean)
    mfntPostMonthFont.Italic = vData
    PropertyChanged "PostMonthFontItalic"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PostMonthFontUnderline() As Boolean
Attribute PostMonthFontUnderline.VB_MemberFlags = "400"
    PostMonthFontUnderline = mfntPostMonthFont.Underline
End Property

Public Property Let PostMonthFontUnderline(ByVal vData As Boolean)
    mfntPostMonthFont.Underline = vData
    PropertyChanged "PostMonthFontUnderline"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PostMonthFontStrikethrough() As Boolean
Attribute PostMonthFontStrikethrough.VB_MemberFlags = "400"
    PostMonthFontStrikethrough = mfntPostMonthFont.Strikethrough
End Property

Public Property Let PostMonthFontStrikethrough(ByVal vData As Boolean)
    mfntPostMonthFont.Strikethrough = vData
    PropertyChanged "PostMonthFontStrikethrough"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PreMonthForeColor() As OLE_COLOR
Attribute PreMonthForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PreMonthForeColor = mlngPreMonthForeColor
End Property

Public Property Let PreMonthForeColor(ByVal vData As OLE_COLOR)
    mlngPreMonthForeColor = vData
    PropertyChanged "PreMonthForeColor"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PreMonthBackColor() As OLE_COLOR
Attribute PreMonthBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PreMonthBackColor = mlngPreMonthBackColor
End Property

Public Property Let PreMonthBackColor(ByVal vData As OLE_COLOR)
    mlngPreMonthBackColor = vData
    PropertyChanged "PreMonthBackColor"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get CurrentMonthForeColor() As OLE_COLOR
Attribute CurrentMonthForeColor.VB_Description = "The foreground color used to display the current month calendar cells"
Attribute CurrentMonthForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CurrentMonthForeColor = mlngCurrentMonthForeColor
End Property

Public Property Let CurrentMonthForeColor(ByVal vData As OLE_COLOR)
    mlngCurrentMonthForeColor = vData
    PropertyChanged "CurrentMonthForeColor"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get CurrentMonthBackColor() As OLE_COLOR
Attribute CurrentMonthBackColor.VB_Description = "The background color used to display the current month calendar cells"
Attribute CurrentMonthBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CurrentMonthBackColor = mlngCurrentMonthBackColor
End Property

Public Property Let CurrentMonthBackColor(ByVal vData As OLE_COLOR)
    mlngCurrentMonthBackColor = vData
    PropertyChanged "CurrentMonthBackColor"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PostMonthForeColor() As OLE_COLOR
Attribute PostMonthForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PostMonthForeColor = mlngPostMonthForeColor
End Property

Public Property Let PostMonthForeColor(ByVal vData As OLE_COLOR)
    mlngPostMonthForeColor = vData
    PropertyChanged "PostMonthForeColor"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get PostMonthBackColor() As OLE_COLOR
Attribute PostMonthBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PostMonthBackColor = mlngPostMonthBackColor
End Property

Public Property Let PostMonthBackColor(ByVal vData As OLE_COLOR)
    mlngPostMonthBackColor = vData
    PropertyChanged "PostMonthBackColor"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get ActiveDayFont() As StdFont
Attribute ActiveDayFont.VB_ProcData.VB_Invoke_Property = "StandardFont;Font"
    Set ActiveDayFont = mfntActiveDayFont
End Property

Public Property Set ActiveDayFont(ByVal vData As StdFont)
    Set mfntActiveDayFont = vData
    PropertyChanged "ActiveDayFont"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get ActiveDayFontName() As String
    ActiveDayFontName = mfntActiveDayFont.Name
End Property

Public Property Let ActiveDayFontName(ByVal vData As String)
Attribute ActiveDayFontName.VB_Description = "For the currently selected date"
Attribute ActiveDayFontName.VB_MemberFlags = "400"
    mfntActiveDayFont.Name = vData
    PropertyChanged "ActiveDayFontName"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get ActiveDayFontSize() As Long
Attribute ActiveDayFontSize.VB_Description = "For the currently selected date"
Attribute ActiveDayFontSize.VB_MemberFlags = "400"
    ActiveDayFontSize = mfntActiveDayFont.SIZE
End Property

Public Property Let ActiveDayFontSize(ByVal vData As Long)
    mfntActiveDayFont.SIZE = vData
    PropertyChanged "ActiveDayFontSize"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get ActiveDayFontBold() As Boolean
Attribute ActiveDayFontBold.VB_Description = "For the currently selected date"
Attribute ActiveDayFontBold.VB_MemberFlags = "400"
    ActiveDayFontBold = mfntActiveDayFont.Bold
End Property

Public Property Let ActiveDayFontBold(ByVal vData As Boolean)
    mfntActiveDayFont.Bold = vData
    PropertyChanged "ActiveDayFontBold"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get ActiveDayFontItalic() As Boolean
Attribute ActiveDayFontItalic.VB_Description = "For the currently selected date"
Attribute ActiveDayFontItalic.VB_MemberFlags = "400"
    ActiveDayFontItalic = mfntActiveDayFont.Italic
End Property

Public Property Let ActiveDayFontItalic(ByVal vData As Boolean)
    mfntActiveDayFont.Italic = vData
    PropertyChanged "ActiveDayFontItalic"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get ActiveDayFontUnderline() As Boolean
Attribute ActiveDayFontUnderline.VB_Description = "For the currently selected date"
Attribute ActiveDayFontUnderline.VB_MemberFlags = "400"
    ActiveDayFontUnderline = mfntActiveDayFont.Underline
End Property

Public Property Let ActiveDayFontUnderline(ByVal vData As Boolean)
    mfntActiveDayFont.Underline = vData
    PropertyChanged "ActiveDayFontUnderline"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get ActiveDayFontStrikethrough() As Boolean
Attribute ActiveDayFontStrikethrough.VB_Description = "For the currently selected date"
Attribute ActiveDayFontStrikethrough.VB_MemberFlags = "400"
    ActiveDayFontStrikethrough = mfntActiveDayFont.Strikethrough
End Property

Public Property Let ActiveDayFontStrikethrough(ByVal vData As Boolean)
    mfntActiveDayFont.Strikethrough = vData
    PropertyChanged "ActiveDayFontStrikethrough"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get ActiveDayForeColor() As OLE_COLOR
Attribute ActiveDayForeColor.VB_Description = "The foregound color to display the currenlty selected date"
Attribute ActiveDayForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ActiveDayForeColor = mlngActiveDayForeColor
End Property

Public Property Let ActiveDayForeColor(ByVal vData As OLE_COLOR)
    mlngActiveDayForeColor = vData
    PropertyChanged "ActiveDayForeColor"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get ActiveDayBackColor() As OLE_COLOR
Attribute ActiveDayBackColor.VB_Description = "The background color to display the currenlty selected date"
Attribute ActiveDayBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ActiveDayBackColor = mlngActiveDayBackColor
End Property

Public Property Let ActiveDayBackColor(ByVal vData As OLE_COLOR)
    mlngActiveDayBackColor = vData
    PropertyChanged "ActiveDayBackColor"
    If mblnAutoRefresh Then Call Me.Refresh
End Property

Public Property Get CurrentYear() As Integer
Attribute CurrentYear.VB_Description = "The year value for the currently active period."
Attribute CurrentYear.VB_ProcData.VB_Invoke_Property = ";Date"
    CurrentYear = mintCurrentYear
End Property

Public Property Let CurrentYear(ByVal vData As Integer)
    Dim dtOld As Date
    Dim dtNew As Date
    Dim blnCancel As Boolean
    
    dtOld = mdtCurrentDate
    dtNew = DateSerial(vData, Month(mdtCurrentDate), Day(mdtCurrentDate))
    '// If date goes past current month set date to last day of current month
    If (Month(dtOld) <> Month(dtNew)) Then
        dtNew = DateAdd("d", -1, DateAdd("m", 1, DateSerial(vData, Month(dtOld), 1)))
    End If
    
    If (dtNew <> dtOld) And IsInRange(dtNew) Then
        RaiseEvent WillChangeDate(dtOld, dtNew, blnCancel)
        If blnCancel Then Exit Property
        
        mintCurrentYear = vData
        mdtCurrentDate = dtNew
        Call PopulateDates
        RaiseEvent DateChanged(dtOld, dtNew)
        
        PropertyChanged "CurrentYear"
        If mblnAutoRefresh Then Call Me.Refresh
        Call RefreshYearMenu
    End If
End Property

Public Property Get CurrentMonth() As dtMonthConstants
Attribute CurrentMonth.VB_Description = "The month value for the currently selected calendar cell."
Attribute CurrentMonth.VB_ProcData.VB_Invoke_Property = ";Date"
    CurrentMonth = menuCurrentMonth
End Property

Public Property Let CurrentMonth(ByVal vData As dtMonthConstants)
    Dim dtOld As Date
    Dim dtNew As Date
    Dim blnCancel As Boolean
    
    If (vData >= dtJanuary) And (vData <= dtDecember) Then      '// Just to make sure
        dtOld = mdtCurrentDate
        dtNew = DateSerial(Year(mdtCurrentDate), vData, Day(mdtCurrentDate))
        '// If date goes past current month set date to last day of current month
        If (vData <> Month(dtNew)) Then
            dtNew = DateAdd("d", -1, DateAdd("m", 1, DateSerial(Year(mdtCurrentDate), vData, 1)))
        End If
        
        If (dtNew <> dtOld) And IsInRange(dtNew) Then
            RaiseEvent WillChangeDate(dtOld, dtNew, blnCancel)
            If blnCancel Then Exit Property
            
            menuCurrentMonth = vData
            mdtCurrentDate = dtNew
            Call PopulateDates
            RaiseEvent DateChanged(dtOld, dtNew)
            
            PropertyChanged "CurrentMonth"
            If mblnAutoRefresh Then Call Me.Refresh
        End If
    End If
End Property

Public Property Get CurrentDate() As Date
Attribute CurrentDate.VB_Description = "The date value for the currently selected calendar cell"
Attribute CurrentDate.VB_ProcData.VB_Invoke_Property = ";Date"
Attribute CurrentDate.VB_MemberFlags = "200"
    CurrentDate = mdtCurrentDate
End Property

Public Property Let CurrentDate(ByVal vData As Date)
    Dim dtOld As Date
    Dim blnCancel As Boolean
    
    If IsDate(vData) Then
        dtOld = mdtCurrentDate
        If (vData <> dtOld) And IsInRange(vData) Then
            RaiseEvent WillChangeDate(dtOld, vData, blnCancel)
            If blnCancel Then Exit Property
            
            mdtCurrentDate = vData
            menuCurrentMonth = Month(vData)
            mintCurrentYear = Year(vData)
            Call PopulateDates
            RaiseEvent DateChanged(dtOld, vData)
            
            PropertyChanged "CurrentDate"
        End If
        '// This line is written outside of If condition
        '// only to show focus rectangle for the first time.
        If mblnAutoRefresh Then Call Me.Refresh
    End If
End Property

Public Property Get MinDate() As Date
Attribute MinDate.VB_Description = "Minimum date, which can be selected"
Attribute MinDate.VB_ProcData.VB_Invoke_Property = ";Date"
    MinDate = mdtMinDate
End Property

Public Property Let MinDate(ByVal vData As Date)
    mdtMinDate = vData
    PropertyChanged "MinDate"
End Property

Public Property Get MaxDate() As Date
Attribute MaxDate.VB_Description = "Maximum date, which can be selected"
Attribute MaxDate.VB_ProcData.VB_Invoke_Property = ";Date"
    MaxDate = mdtMaxDate
End Property

Public Property Let MaxDate(ByVal vData As Date)
    mdtMaxDate = vData
    PropertyChanged "MaxDate"
End Property

Public Property Get ShowToolTip() As Boolean
Attribute ShowToolTip.VB_Description = "Turn on/off showing tooltip."
Attribute ShowToolTip.VB_ProcData.VB_Invoke_Property = "Appearance"
    ShowToolTip = mblnShowToolTip
End Property

Public Property Let ShowToolTip(ByVal vData As Boolean)
    mblnShowToolTip = vData
    PropertyChanged "ShowToolTip"
End Property

Public Property Get DateTipFormat() As String
Attribute DateTipFormat.VB_Description = "The format to use when displaying the date in the tooltip."
Attribute DateTipFormat.VB_ProcData.VB_Invoke_Property = "Appearance"
    DateTipFormat = mstrDateTipFormat
End Property

Public Property Let DateTipFormat(ByVal vData As String)
    mstrDateTipFormat = vData
    PropertyChanged "DateTipFormat"
End Property

Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = mblnShowFocusRect
End Property

Public Property Let ShowFocusRect(ByVal vData As Boolean)
    If Ambient.UserMode Then Err.Raise 382
    mblnShowFocusRect = vData
    PropertyChanged "ShowFocusRect"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get TitleAppearance() As dtAppearanceConstants
Attribute TitleAppearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TitleAppearance = menuTitleAppearance
End Property

Public Property Let TitleAppearance(ByVal vData As dtAppearanceConstants)
    If (vData = dt3D) Or (vData = dtFlat) Then
        menuTitleAppearance = vData
        PropertyChanged "TitleAppearance"
        '// Refresh control to reflect the changes
        If mblnAutoRefresh Then Call Me.Refresh
    End If
End Property

Public Sub About()
Attribute About.VB_UserMemId = -552
    frmAbout.ActivateForm "MONTH"
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_EnterFocus()
    If mblnShowFocusRect Then Call DisplayFocusRect
End Sub

Private Sub UserControl_ExitFocus()
    Call Me.Refresh
End Sub

Private Sub UserControl_Initialize()
    UserControl.ScaleMode = vbPixels
End Sub

Private Sub UserControl_InitProperties()
    menuAppearance = mdefAppearance
    menuBorderStyle = mdefBorderStyle
    mblnShowDayHeader = mdefShowDayHeader
    mlngDayHeaderBackColor = mdefDayHeaderBackColor
    mlngDayHeaderForeColor = mdefDayHeaderForeColor
    menuFirstDayOfWeek = mdefFirstDayOfWeek
    menuDayHeaderFormat = mdefDayHeaderFormat
    mblnShowLines = mdefShowLines
    mlngLineColor = mdefLineColor
    mlngPreMonthBackColor = mdefPreMonthBackColor
    mlngCurrentMonthBackColor = mdefCurrentMonthBackColor
    mlngPostMonthBackColor = mdefPostMonthBackColor
    mlngPreMonthForeColor = mdefPreMonthForeColor
    mlngCurrentMonthForeColor = mdefCurrentMonthForeColor
    mlngPostMonthForeColor = mdefPostMonthForeColor
    mdtCurrentDate = Date
    menuCurrentMonth = Month(mdtCurrentDate)
    mintCurrentYear = Year(mdtCurrentDate)
    mlngActiveDayForeColor = mdefActiveDayForeColor
    mlngActiveDayBackColor = mdefActiveDayBackColor
    mblnAutoRefresh = mdefAutoRefresh
    mintEdgeOffset = GetEdgeOffset()
    mdtMinDate = mdefMinDate
    mdtMaxDate = mdefMaxDate
    mblnShowToolTip = mdefShowToolTip
    mstrDateTipFormat = mdefDateTipFormat
    mlngTitleForeColor = mdefTitleForeColor
    mlngTitleBackColor = mdefTitleBackColor
    mblnShowFocusRect = mdefShowFocusRect
    menuTitleAppearance = mdefTitleAppearance
    
    With mfntTitleFont
        .Name = "Tahoma"
        .SIZE = 8
        .Bold = True
        .Italic = False
        .Underline = False
        .Strikethrough = False
    End With
    
    With mfntDayHeaderFont
        .Name = "Tahoma"
        .SIZE = 8
        .Bold = False
        .Italic = False
        .Underline = False
        .Strikethrough = False
    End With
    
    With mfntCurrentMonthFont
        .Name = "Tahoma"
        .SIZE = 8
        .Bold = False
        .Italic = False
        .Underline = False
        .Strikethrough = False
    End With

    With mfntPreMonthFont
        .Name = "Tahoma"
        .SIZE = 8
        .Bold = False
        .Italic = False
        .Underline = False
        .Strikethrough = False
    End With

    With mfntPostMonthFont
        .Name = "Tahoma"
        .SIZE = 8
        .Bold = False
        .Italic = False
        .Underline = False
        .Strikethrough = False
    End With

    With mfntActiveDayFont
        .Name = "Tahoma"
        .SIZE = 8
        .Bold = True
        .Italic = False
        .Underline = False
        .Strikethrough = False
    End With

    Call PopulateDates
    Call Me.Refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intDays As Integer
    
    '// Make sure that tooltip is not visible
    Call DisplayToolTip(False)
    
    Select Case KeyCode
        Case vbKeyLeft: intDays = (-1)
        Case vbKeyUp: intDays = (-7)
        Case vbKeyRight: intDays = 1
        Case vbKeyDown: intDays = 7
        Case vbKeyPageUp: intDays = (-1) * Day(DateAdd("d", (-1), mdtMonthStart))
        Case vbKeyPageDown: intDays = Day(mdtMonthEnd)
    End Select

    CurrentDate = DateAdd("d", intDays, mdtCurrentDate)
    If mblnShowFocusRect Then Call DisplayFocusRect
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not mblnMenuVisible Then
        RaiseEvent MouseDown(Button, Shift, x, y)
        
        If (Button = vbRightButton) Then
            If IsMouseInMonthArea(x, y) Then
                Call ShowMenu(MonthMenu)
            ElseIf IsMouseInYearArea(x, y) Then
                Call ShowMenu(YearMenu)
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim intRow As Integer, intCol As Integer
    Dim intCellId As Integer
    Dim dtDate As Date
    
    If Not mblnMenuVisible Then
        RaiseEvent MouseMove(Button, Shift, x, y)
        
        If GetMouseCell(x, y, intRow, intCol, dtDate) Then
            intCellId = GetCellId(intRow, intCol)
            If Not (mintCurCellId = intCellId) Then
                '// If tooltip is being shown, hide it
                If mblnToolTipVisible Then Call DisplayToolTip(False)
                '// Show tool tip
                Call DisplayToolTip(True, Format$(dtDate, mstrDateTipFormat))
                '// Update current cell id
                mintCurCellId = intCellId
            End If
        Else
            mintCurCellId = 0
            '// If tooltip is being shown, hide it
            Call DisplayToolTip(False)
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim intRow As Integer, intCol As Integer
    Dim intCellId As Integer
    Dim dtDate As Date
    
    If Not mblnMenuVisible Then
        RaiseEvent MouseUp(Button, Shift, x, y)
        
        If (Button = vbLeftButton) Then
            If GetMouseCell(x, y, intRow, intCol, dtDate) Then
                If IsInRange(dtDate) Then
                    '// Make sure that tooltip is not visible
                    Call DisplayToolTip(False)
                    '// Change current date
                    CurrentDate = dtDate
                    '// Show a focus rectangle around the current date
                    If mblnShowFocusRect Then Call DisplayFocusRect
                End If
            Else
                '// Redraw the control to remove the focus
                Call Me.Refresh
            End If
        End If
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    menuAppearance = PropBag.ReadProperty("Appearance", mdefAppearance)
    menuBorderStyle = PropBag.ReadProperty("BorderStyle", mdefBorderStyle)
    mintEdgeOffset = GetEdgeOffset()
    mblnAutoRefresh = PropBag.ReadProperty("AutoRefresh", mdefAutoRefresh)
    mblnShowDayHeader = PropBag.ReadProperty("ShowDayHeader", mdefShowDayHeader)
    mlngDayHeaderBackColor = PropBag.ReadProperty("DayHeaderBackColor", mdefDayHeaderBackColor)
    mlngDayHeaderForeColor = PropBag.ReadProperty("DayHeaderForeColor", mdefDayHeaderForeColor)
    menuFirstDayOfWeek = PropBag.ReadProperty("FirstDayOfWeek", mdefFirstDayOfWeek)
    menuDayHeaderFormat = PropBag.ReadProperty("DayHeaderFormat", mdefDayHeaderFormat)
    mblnShowLines = PropBag.ReadProperty("ShowLines", mdefShowLines)
    mintEdgeOffset = GetEdgeOffset()
    mlngLineColor = PropBag.ReadProperty("LineColor", mdefLineColor)
    mlngPreMonthForeColor = PropBag.ReadProperty("PreMonthForeColor", mdefPreMonthForeColor)
    mlngPreMonthBackColor = PropBag.ReadProperty("PreMonthBackColor", mdefPreMonthBackColor)
    mlngCurrentMonthForeColor = PropBag.ReadProperty("CurrentMonthForeColor", mdefCurrentMonthForeColor)
    mlngCurrentMonthBackColor = PropBag.ReadProperty("CurrentMonthBackColor", mdefCurrentMonthBackColor)
    mlngPostMonthForeColor = PropBag.ReadProperty("PostMonthForeColor", mdefPostMonthForeColor)
    mlngPostMonthBackColor = PropBag.ReadProperty("PostMonthBackColor", mdefPostMonthBackColor)
    mdtCurrentDate = PropBag.ReadProperty("CurrentDate", Date)
    menuCurrentMonth = PropBag.ReadProperty("CurrentMonth", Month(mdtCurrentDate))
    mintCurrentYear = PropBag.ReadProperty("CurrentYear", Year(mdtCurrentDate))
    mlngActiveDayForeColor = PropBag.ReadProperty("ActiveDayForeColor", mdefActiveDayForeColor)
    mlngActiveDayBackColor = PropBag.ReadProperty("ActiveDayBackColor", mdefActiveDayBackColor)
    mdtMinDate = PropBag.ReadProperty("MinDate", mdefMinDate)
    mdtMaxDate = PropBag.ReadProperty("MaxDate", mdefMaxDate)
    mblnShowToolTip = PropBag.ReadProperty("ShowToolTip", mdefShowToolTip)
    mstrDateTipFormat = PropBag.ReadProperty("DateTipFormat", mdefDateTipFormat)
    mlngTitleForeColor = PropBag.ReadProperty("TitleForeColor", mdefTitleForeColor)
    mlngTitleBackColor = PropBag.ReadProperty("TitleBackColor", mdefTitleBackColor)
    mblnShowFocusRect = PropBag.ReadProperty("ShowFocusRect", mdefShowFocusRect)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    menuTitleAppearance = PropBag.ReadProperty("TitleAppearance", dt3D)
    
    '// title font
    With mfntTitleFont
        .Name = PropBag.ReadProperty("TitleFontName", "Tahoma")
        .SIZE = PropBag.ReadProperty("TitleFontSize", 8)
        .Bold = PropBag.ReadProperty("TitleFontBold", True)
        .Italic = PropBag.ReadProperty("TitleFontItalic", False)
        .Underline = PropBag.ReadProperty("TitleFontUnderline", False)
        .Strikethrough = PropBag.ReadProperty("TitleFontStrikethrough", False)
    End With
    
    '// day header font
    With mfntDayHeaderFont
        .Name = PropBag.ReadProperty("DayHeaderFontName", "Tahoma")
        .SIZE = PropBag.ReadProperty("DayHeaderFontSize", 8)
        .Bold = PropBag.ReadProperty("DayHeaderFontBold", False)
        .Italic = PropBag.ReadProperty("DayHeaderFontItalic", False)
        .Underline = PropBag.ReadProperty("DayHeaderFontUnderline", False)
        .Strikethrough = PropBag.ReadProperty("DayHeaderFontStrikethrough", False)
    End With
    
    '// Current month font
    With mfntCurrentMonthFont
        .Name = PropBag.ReadProperty("CurrentMonthFontName", "Tahoma")
        .SIZE = PropBag.ReadProperty("CurrentMonthFontSize", 8)
        .Bold = PropBag.ReadProperty("CurrentMonthFontBold", False)
        .Italic = PropBag.ReadProperty("CurrentMonthFontItalic", False)
        .Underline = PropBag.ReadProperty("CurrentMonthFontUnderline", False)
        .Strikethrough = PropBag.ReadProperty("CurrentMonthFontStrikethrough", False)
    End With
    
    '// Previous month font
    With mfntPreMonthFont
        .Name = PropBag.ReadProperty("PreMonthFontName", "Tahoma")
        .SIZE = PropBag.ReadProperty("PreMonthFontSize", 8)
        .Bold = PropBag.ReadProperty("PreMonthFontBold", False)
        .Italic = PropBag.ReadProperty("PreMonthFontItalic", False)
        .Underline = PropBag.ReadProperty("PreMonthFontUnderline", False)
        .Strikethrough = PropBag.ReadProperty("PreMonthFontStrikethrough", False)
    End With
    
    '// Post month font
    With mfntPostMonthFont
        .Name = PropBag.ReadProperty("PostMonthFontName", "Tahoma")
        .SIZE = PropBag.ReadProperty("PostMonthFontSize", 8)
        .Bold = PropBag.ReadProperty("PostMonthFontBold", False)
        .Italic = PropBag.ReadProperty("PostMonthFontItalic", False)
        .Underline = PropBag.ReadProperty("PostMonthFontUnderline", False)
        .Strikethrough = PropBag.ReadProperty("PostMonthFontStrikethrough", False)
    End With
    
    '// Active Day font
    With mfntActiveDayFont
        .Name = PropBag.ReadProperty("ActiveDayFontName", "Tahoma")
        .SIZE = PropBag.ReadProperty("ActiveDayFontSize", 8)
        .Bold = PropBag.ReadProperty("ActiveDayFontBold", True)
        .Italic = PropBag.ReadProperty("ActiveDayFontItalic", True)
        .Underline = PropBag.ReadProperty("ActiveDayFontUnderline", False)
        .Strikethrough = PropBag.ReadProperty("ActiveDayFontStrikethrough", False)
    End With
    
    Call PopulateDates
    Call Me.Refresh
End Sub

Private Sub UserControl_Resize()
    Call Me.Refresh
End Sub

Private Sub UserControl_Terminate()
    '// release memory
    Set mfntActiveDayFont = Nothing
    Set mfntDayHeaderFont = Nothing
    Set mfntCurrentMonthFont = Nothing
    Set mfntTitleFont = Nothing
    Set mfntPreMonthFont = Nothing
    Set mfntPostMonthFont = Nothing
    
    If Not mfrmToolTip Is Nothing Then Set mfrmToolTip = Nothing
    If IsMenu(mlngMonthMenu) Then Call DestroyMenu(mlngMonthMenu)
    If IsMenu(mlngYearMenu) Then Call DestroyMenu(mlngYearMenu)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Appearance", menuAppearance, mdefAppearance)
    Call PropBag.WriteProperty("BorderStyle", menuBorderStyle, mdefBorderStyle)
    Call PropBag.WriteProperty("AutoRefresh", mblnAutoRefresh, mdefAutoRefresh)
    Call PropBag.WriteProperty("ShowDayHeader", mblnShowDayHeader, mdefShowDayHeader)
    Call PropBag.WriteProperty("DayHeaderFont", mfntDayHeaderFont, UserControl.Font)
    Call PropBag.WriteProperty("DayHeaderBackColor", mlngDayHeaderBackColor, mdefDayHeaderBackColor)
    Call PropBag.WriteProperty("DayHeaderForeColor", mlngDayHeaderForeColor, mdefDayHeaderForeColor)
    Call PropBag.WriteProperty("FirstDayOfWeek", menuFirstDayOfWeek, mdefFirstDayOfWeek)
    Call PropBag.WriteProperty("DayHeaderFormat", menuDayHeaderFormat, mdefDayHeaderFormat)
    Call PropBag.WriteProperty("ShowLines", mblnShowLines, mdefShowLines)
    Call PropBag.WriteProperty("LineColor", mlngLineColor, mdefLineColor)
    Call PropBag.WriteProperty("PreMonthForeColor", mlngPreMonthForeColor, mdefPreMonthForeColor)
    Call PropBag.WriteProperty("PreMonthBackColor", mlngPreMonthBackColor, mdefPreMonthBackColor)
    Call PropBag.WriteProperty("CurrentMonthForeColor", mlngCurrentMonthForeColor, mdefCurrentMonthForeColor)
    Call PropBag.WriteProperty("CurrentMonthBackColor", mlngCurrentMonthBackColor, mdefCurrentMonthBackColor)
    Call PropBag.WriteProperty("PostMonthForeColor", mlngPostMonthForeColor, mdefPostMonthForeColor)
    Call PropBag.WriteProperty("PostMonthBackColor", mlngPostMonthBackColor, mdefPostMonthBackColor)
    Call PropBag.WriteProperty("CurrentDate", mdtCurrentDate, Date)
    Call PropBag.WriteProperty("CurrentMonth", menuCurrentMonth, Month(mdtCurrentDate))
    Call PropBag.WriteProperty("CurrentYear", mintCurrentYear, Year(mdtCurrentDate))
    Call PropBag.WriteProperty("ActiveDayForeColor", mlngActiveDayForeColor, mdefActiveDayForeColor)
    Call PropBag.WriteProperty("ActiveDayBackColor", mlngActiveDayBackColor, mdefActiveDayBackColor)
    Call PropBag.WriteProperty("MinDate", mdtMinDate, mdefMinDate)
    Call PropBag.WriteProperty("MaxDate", mdtMaxDate, mdefMaxDate)
    Call PropBag.WriteProperty("ShowToolTip", mblnShowToolTip, mdefShowToolTip)
    Call PropBag.WriteProperty("DateTipFormat", mstrDateTipFormat, mdefDateTipFormat)
    Call PropBag.WriteProperty("TitleForeColor", mlngTitleForeColor, mdefTitleForeColor)
    Call PropBag.WriteProperty("TitleBackColor", mlngTitleBackColor, mdefTitleBackColor)
    Call PropBag.WriteProperty("ShowFocusRect", mblnShowFocusRect, mdefShowFocusRect)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("TitleAppearance", menuTitleAppearance, dt3D)
    
    '// Title font
    With mfntTitleFont
        Call PropBag.WriteProperty("TitleFontName", .Name, "Tahoma")
        Call PropBag.WriteProperty("TitleFontSize", .SIZE, 8)
        Call PropBag.WriteProperty("TitleFontBold", .Bold, True)
        Call PropBag.WriteProperty("TitleFontItalic", .Italic, False)
        Call PropBag.WriteProperty("TitleFontUnderline", .Underline, False)
        Call PropBag.WriteProperty("TitleFontStrikethrough", .Strikethrough, False)
    End With
    
    '// Day header font
    With mfntDayHeaderFont
        Call PropBag.WriteProperty("DayHeaderFontName", .Name, "Tahoma")
        Call PropBag.WriteProperty("DayHeaderFontSize", .SIZE, 8)
        Call PropBag.WriteProperty("DayHeaderFontBold", .Bold, False)
        Call PropBag.WriteProperty("DayHeaderFontItalic", .Italic, False)
        Call PropBag.WriteProperty("DayHeaderFontUnderline", .Underline, False)
        Call PropBag.WriteProperty("DayHeaderFontStrikethrough", .Strikethrough, False)
    End With
    
    '// Current month font
    With mfntCurrentMonthFont
        Call PropBag.WriteProperty("CurrentMonthFontName", .Name, "Tahoma")
        Call PropBag.WriteProperty("CurrentMonthFontSize", .SIZE, 8)
        Call PropBag.WriteProperty("CurrentMonthFontBold", .Bold, False)
        Call PropBag.WriteProperty("CurrentMonthFontItalic", .Italic, False)
        Call PropBag.WriteProperty("CurrentMonthFontUnderline", .Underline, False)
        Call PropBag.WriteProperty("CurrentMonthFontStrikethrough", .Strikethrough, False)
    End With

    '// Previous month font
    With mfntPreMonthFont
        Call PropBag.WriteProperty("PreMonthFontName", .Name, "Tahoma")
        Call PropBag.WriteProperty("PreMonthFontSize", .SIZE, 8)
        Call PropBag.WriteProperty("PreMonthFontBold", .Bold, False)
        Call PropBag.WriteProperty("PreMonthFontItalic", .Italic, False)
        Call PropBag.WriteProperty("PreMonthFontUnderline", .Underline, False)
        Call PropBag.WriteProperty("PreMonthFontStrikethrough", .Strikethrough, False)
    End With

    '// Post month font
    With mfntPostMonthFont
        Call PropBag.WriteProperty("PostMonthFontName", .Name, "Tahoma")
        Call PropBag.WriteProperty("PostMonthFontSize", .SIZE, 8)
        Call PropBag.WriteProperty("PostMonthFontBold", .Bold, False)
        Call PropBag.WriteProperty("PostMonthFontItalic", .Italic, False)
        Call PropBag.WriteProperty("PostMonthFontUnderline", .Underline, False)
        Call PropBag.WriteProperty("PostMonthFontStrikethrough", .Strikethrough, False)
    End With
    
    '// Active day font
    With mfntActiveDayFont
        Call PropBag.WriteProperty("ActiveDayFontName", .Name, "Tahoma")
        Call PropBag.WriteProperty("ActiveDayFontSize", .SIZE, 8)
        Call PropBag.WriteProperty("ActiveDayFontBold", .Bold, True)
        Call PropBag.WriteProperty("ActiveDayFontItalic", .Italic, True)
        Call PropBag.WriteProperty("ActiveDayFontUnderline", .Underline, False)
        Call PropBag.WriteProperty("ActiveDayFontStrikethrough", .Strikethrough, False)
    End With
End Sub

Public Sub Refresh()
'// This procedure contains codes from www.vbaccelerator.com
'// Thanks to Steve McMahon

    Dim tr                  As RECT
    Dim clsMDC              As New cMemoryDC
    Dim intOffset           As Integer
    Dim intRow              As Integer
    Dim intCol              As Integer
    Dim intHeaderHeight     As Integer
    Dim strCaption          As String
    Dim intFirstDayOfWeek   As Integer
    Dim dtCurrentDate       As Date
    Dim intCounter          As Integer
    Dim lngEdgeStyle        As Long
    Dim intTitleHeight      As Integer
    Dim intTopMargin        As Integer
    Dim i As Integer, j As Integer
    
    ReDim mudtCellInfo(1 To DEF_CALENDAR_ROWS, 1 To DEF_CALENDAR_COLS)
    
    '// Calculate height needed to show title
    intTitleHeight = GetTitleHeight()
    '// Set top margin for calender
    intOffset = IIf(mblnShowLines, 1, 0)
    intTopMargin = intTitleHeight + intOffset
    
    With clsMDC
        '// Start drawing
        Call .StartDrawing(hdc, ScaleWidth, ScaleHeight)
        
        Set .Font = UserControl.Font
        '// Draw edge
        Select Case menuAppearance
            Case dt3D, dtFlat: lngEdgeStyle = mdcSunken
            Case dtEtched: lngEdgeStyle = mdcEtched
            Case dtThin: lngEdgeStyle = mdcThinRaised
        End Select
        Call .DrawBorder(0, 0, ScaleWidth, ScaleHeight, lngEdgeStyle, (menuAppearance = dtFlat))
        '// If edge style is flat and border style is fixed single, draw a frame
        If (menuAppearance = dtFlat) And (menuBorderStyle = dtFixedSingle) Then
            Call .DrawFrame(0, 0, ScaleWidth, ScaleHeight, SystemColorConstants.vbWindowFrame)
        End If
        
        '// Determine the Width of our calendar cells
        '// With a flat or noline border the borders are
        '// not as thick as the 3D version
        mintCellWidth = (ScaleWidth - 2 * mintEdgeOffset) \ DEF_CALENDAR_COLS
        '// Set the beginning of the first calendar row
        intRow = intTopMargin + mintEdgeOffset
        '// save the left and right positions of calenders's dates part
        mintCalLeft = mintEdgeOffset
        '// Display the days of the week header, if enabled
        If mblnShowDayHeader Then
            .LineColor = mlngLineColor
            
            Set .Font = mfntDayHeaderFont
            .BackColor = mlngDayHeaderBackColor
            .ForeColor = mlngDayHeaderForeColor
            
            intOffset = IIf(mblnShowLines, 3, 4)
            mintCellHeight = TextHeight("Wednesday") + intOffset
            intHeaderHeight = mintCellHeight
            
            intFirstDayOfWeek = CInt(menuFirstDayOfWeek)
            intCol = mintEdgeOffset
            
            For i = 1 To DEF_CALENDAR_COLS
                Select Case menuDayHeaderFormat
                    Case dtSingleLetter
                        strCaption = Left$(Format$(intFirstDayOfWeek, "Ddd"), 1)
                    Case dtMedium
                        strCaption = Format$(intFirstDayOfWeek, "Ddd")
                    Case dtFullName
                        strCaption = Format$(intFirstDayOfWeek, "Dddd")
                End Select
                
                intOffset = IIf(mblnShowLines, 1, 2)
                Call .Draw3DRect(intCol, intRow, _
                                 (mintCellWidth + intOffset), _
                                 (mintCellHeight + intOffset), strCaption, _
                                 mdcCenterCenter, IIf(mblnShowLines, mdcappFlat, mdcappNoLines))
                
                intCol = intCol + mintCellWidth
                intFirstDayOfWeek = intFirstDayOfWeek + 1
                If (intFirstDayOfWeek > 7) Then intFirstDayOfWeek = 1
            Next i
        End If
        
        '// If the days of the week are being displayed then
        '// account for them in our cell height calculations
        If mblnShowDayHeader Then
            intOffset = IIf(mblnShowLines, 3, 2)
            mintCellHeight = (ScaleHeight - (intTopMargin + (2 * mintEdgeOffset) + mintCellHeight)) \ DEF_CALENDAR_ROWS
            intRow = TextHeight("Wednesday") + intOffset + (intTopMargin + mintEdgeOffset)
        Else
            mintCellHeight = (ScaleHeight - (intTopMargin + (2 * mintEdgeOffset))) \ DEF_CALENDAR_ROWS
            intRow = intTopMargin + mintEdgeOffset
        End If
        
        '// Save top position of calnder's dates part
        mintCalTop = intRow
        
        '// Create the grid and display the day value
        dtCurrentDate = mdtCalenderStart
        Set .Font = mfntCurrentMonthFont
        For i = 1 To DEF_CALENDAR_ROWS
            intCol = mintEdgeOffset
            For j = 1 To DEF_CALENDAR_COLS
                '// Get the next date value that will be displayed
                dtCurrentDate = DateAdd("d", intCounter, mdtCalenderStart)
                
                '// Determine which background and foreground colors to use
                If (dtCurrentDate < mdtMonthStart) Then
                    Set .Font = mfntPreMonthFont
                    .BackColor = mlngPreMonthBackColor
                    .ForeColor = mlngPreMonthForeColor
                ElseIf (dtCurrentDate > mdtMonthEnd) Then
                    Set .Font = mfntPostMonthFont
                    .BackColor = mlngPostMonthBackColor
                    .ForeColor = mlngPostMonthForeColor
                Else
                    Set .Font = mfntCurrentMonthFont
                    .BackColor = mlngCurrentMonthBackColor
                    .ForeColor = mlngCurrentMonthForeColor
                End If
                
                '// Set properties for active day
                If (dtCurrentDate = mdtCurrentDate) Then
                    Set .Font = mfntActiveDayFont
                    .ForeColor = mlngActiveDayForeColor
                    .BackColor = mlngActiveDayBackColor
                End If
                
                intOffset = IIf(mblnShowLines, 1, 2)
                Call .Draw3DRect(intCol, intRow, mintCellWidth + intOffset, _
                                 mintCellHeight + intOffset, Day(dtCurrentDate), _
                                 mdcCenterCenter, IIf(mblnShowLines, mdcappFlat, mdcappNoLines))
                
                '// Save cell information
                With mudtCellInfo(i, j)
                    .lForeColor = clsMDC.ForeColor
                    .lBackColor = clsMDC.BackColor
                    With .tLocation
                        .Left = intCol
                        .Top = intRow
                        .Right = intCol + (mintCellWidth + intOffset)
                        .Bottom = intRow + (mintCellHeight + intOffset)
                    End With
                End With
                
                '// Set back font in case of active day
                If (dtCurrentDate = mdtCurrentDate) Then
                    Set .Font = mfntCurrentMonthFont
                End If
                
                intCol = intCol + mintCellWidth
                intCounter = intCounter + 1
            Next j
            intRow = intRow + mintCellHeight
        Next i
        
        '// Save right and bottom position of calender's dates part
        mintCalRight = intCol
        mintCalBottom = intRow
        
        '// Roll back back color
        .BackColor = mlngTitleBackColor
        
        Set .Font = mfntTitleFont
        '// Fill backgroud with title back color
        intOffset = mintEdgeOffset + 1
        Call .FillRect(intOffset, intOffset, _
                    mintCellWidth * DEF_CALENDAR_COLS, intTitleHeight, mlngTitleBackColor)
        '// Calculate year area
        intOffset = IIf((menuAppearance = dtFlat) And Not mblnShowLines, 1, 0)
        With tr
            .Left = (intCol - DEF_YEAR_WIDTH) + intOffset
            .Top = mintEdgeOffset + intOffset
            .Right = intCol
            .Bottom = mintEdgeOffset + intTitleHeight
        End With
        '// Save year area
        Call CopyRect(mudtYearArea, tr)
        '// Print year
        Call .DrawText(tr.Left, tr.Top, (tr.Right - tr.Left), (tr.Bottom - tr.Top), mintCurrentYear, mdcTextSingleLineCenter, mlngTitleForeColor)
        '// Draw year section
        Call .DrawBorder(tr.Left, tr.Top, (tr.Right - tr.Left), (tr.Bottom - tr.Top), mdcThinRaised, (menuTitleAppearance = dtFlat))
        '// Calculate month area
        With tr
            .Right = .Left - 2
            .Left = mintEdgeOffset + intOffset
        End With
        '// Save month area
        Call CopyRect(mudtMonthArea, tr)
        '// Print month name
        If (menuCurrentMonth > 0) Then
            Call .DrawText(tr.Left, tr.Top, (tr.Right - tr.Left), (tr.Bottom - tr.Top), MonthName(menuCurrentMonth), mdcTextSingleLineCenter, mlngTitleForeColor)
        End If
        '// Draw month section
        Call .DrawBorder(tr.Left, tr.Top, (tr.Right - tr.Left), (tr.Bottom - tr.Top), mdcThinRaised, (menuTitleAppearance = dtFlat))
        
        '// All done drawing our control so lets display it
        Call .StopDrawing(0&, 0&, ScaleWidth, ScaleHeight)
    End With
End Sub

Public Function HitTest(x As Long, y As Long, Optional dDate As Date) As dtCalendarHitTestAreas
    Dim intRow As Integer
    Dim intCol As Integer
    
    If ((x > mintCalLeft) And (x < mintCalRight)) And _
       ((y > mintCalTop) And (y < mintCalBottom)) Then
        '// Calculate current row
        intRow = ((y - mintCalTop) \ mintCellHeight) + 1
        '// Calculate current column
        intCol = ((x - mintCalLeft) \ mintCellWidth) + 1
        '// Get the date in cell
        dDate = DateAdd("d", (intRow - 1) * DEF_CALENDAR_COLS + (intCol - 1), mdtCalenderStart)
        '// Return hit test area
        HitTest = dtCalendarDate
    Else
        HitTest = dtInvalid
    End If
End Function

Private Sub PopulateDates()
    Dim intPrevFirstDay As Integer
    Dim intPrevLastDay As Integer
    Dim intPrevYear As Integer

    '// get first day of the current month
    mdtMonthStart = DateSerial(mintCurrentYear, menuCurrentMonth, 1)
    '// get last day of the current month
    mdtMonthEnd = DateAdd("d", -1, DateAdd("m", 1, mdtMonthStart))
    
    '// Set calendar start date
    Call GetPrevMonthDays(intPrevFirstDay, intPrevLastDay, intPrevYear)
    '// Check to see if the period start falls on the first day
    If (intPrevLastDay = (-1)) Then
        '// Falls on the FirstDay so lets add a week to the calendar start
        '// date so that the user will be able to select a date from the
        '// previous period.
        mdtCalenderStart = DateAdd("d", -7, mdtMonthStart)
    Else
        mdtCalenderStart = DateSerial(intPrevYear, Month(DateAdd("m", -1, mdtMonthStart)), intPrevFirstDay)
    End If
End Sub

Private Sub GetPrevMonthDays( _
                    intFirstDay As Integer, _
                    intLastDay As Integer, _
                    intYear As Integer)
    
    Dim intColDayOne    As Integer      '// Column of 1st day of cur month
    Dim dtTemp          As Date         '// Temp date
    
    '// Construct a date to do date math
    dtTemp = DateSerial(mintCurrentYear, menuCurrentMonth, 1)
    '// Determine the column of the first day of the current month
    intColDayOne = Weekday(dtTemp, menuFirstDayOfWeek)
    
    '// If the first day of the current month is in column 1, we
    '// don't need to paint any days from the prev month, so return
    '// zeros and -1 for the first and last value
    If (intColDayOne = 1) Then
        intFirstDay = 0
        intLastDay = -1
    Else
        '// If there are days to paint, calculate the last and
        '// first day using date math
        dtTemp = DateAdd("d", (-1), dtTemp)
        intLastDay = VBA.Day(dtTemp)
        
        dtTemp = DateAdd("d", -(intColDayOne - 2), dtTemp)
        intFirstDay = VBA.Day(dtTemp)
    
        dtTemp = DateAdd("d", -(intColDayOne - 2), dtTemp)
        intYear = VBA.Year(dtTemp)
    End If
End Sub

Private Function GetEdgeOffset() As Integer
    Dim intOffset As Integer
    
    Select Case menuAppearance
        Case dt3D: intOffset = 2
        Case dtEtched, dtThin: intOffset = 1
        Case dtFlat
            Select Case menuBorderStyle
                Case dtNone: intOffset = 0
                Case dtFixedSingle: intOffset = IIf(mblnShowLines, 1, 0)
            End Select
    End Select
    
    GetEdgeOffset = intOffset
End Function

Private Function GetMouseCell(ByVal sngXPos As Single, ByVal sngYPos As Single, _
                              intRow As Integer, intCol As Integer, dtDate As Date) As Boolean
    
    If ((sngXPos > mintCalLeft) And (sngXPos < mintCalRight)) And _
       ((sngYPos > mintCalTop) And (sngYPos < mintCalBottom)) Then
        '// Calculate current row
        intRow = ((sngYPos - mintCalTop) \ mintCellHeight) + 1
        '// Calculate current column
        intCol = ((sngXPos - mintCalLeft) \ mintCellWidth) + 1
        '// Get the date in cell
        dtDate = DateAdd("d", (intRow - 1) * DEF_CALENDAR_COLS + (intCol - 1), mdtCalenderStart)
        '// Return true
        GetMouseCell = True
    End If
End Function

Private Function GetCellId(intRow As Integer, intCol As Integer) As Integer
    GetCellId = ((intRow - 1) * DEF_CALENDAR_COLS) + intCol
End Function

Private Function IsInRange(ByVal dtDate As Date) As Boolean
    IsInRange = ((dtDate >= mdtMinDate) And (dtDate <= mdtMaxDate))
End Function

Private Sub DisplayFocusRect()
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim intCell     As Integer
    Dim udtRct      As RECT
    
    If mblnShowFocusRect Then
        intCell = DateDiff("d", mdtCalenderStart, mdtCurrentDate)
        intRow = (intCell \ DEF_CALENDAR_COLS) + 1
        intCol = (intCell Mod DEF_CALENDAR_COLS) + 1
        
        Call CopyRect(udtRct, mudtCellInfo(intRow, intCol).tLocation)
        Call InflateRect(udtRct, (-2), (-2))
        Call DrawFocusRect(hdc, udtRct)
    End If
End Sub

Private Sub DisplayToolTip(blnShow As Boolean, Optional strToolTip As String)
    If mblnShowToolTip Then
        If blnShow Then
            If mfrmToolTip Is Nothing Then
                Set mfrmToolTip = New frmToolTip
            End If
            Call mfrmToolTip.DisplayToolTip(strToolTip)
            Call SetCapture(hwnd)
        Else
            If mblnToolTipVisible Then
                Call mfrmToolTip.HideToolTip
                Set mfrmToolTip = Nothing
                If (GetCapture() = hwnd) Then Call ReleaseCapture
            End If
        End If
        mblnToolTipVisible = blnShow
    End If
End Sub

Private Sub CreateMonthMenu()
    Dim i As Integer
    
    If Not IsMenu(mlngMonthMenu) Then
        mlngMonthMenu = CreatePopupMenu()
        For i = dtJanuary To dtDecember
            Call AppendMenu(mlngMonthMenu, MF_STRING, CLng(i), MonthName(CLng(i)))
        Next i
    End If
End Sub

Private Sub CreateYearMenu()
    Dim intYear     As Integer
    Dim i           As Integer
    
    '// If year menu is not created already, create it first.
    If Not IsMenu(mlngYearMenu) Then mlngYearMenu = CreatePopupMenu()
    '// Calculate start year
    intYear = mintCurrentYear - (DEF_YEAR_MENU_COUNT / 2)
    '// Create menu items
    For i = intYear To (intYear + DEF_YEAR_MENU_COUNT)
        Call AppendMenu(mlngYearMenu, MF_STRING, CLng(i), CStr(i))
    Next i
End Sub

Private Sub RefreshYearMenu()
    Dim intYear     As Integer
    Dim i           As Integer

    If IsMenu(mlngYearMenu) Then
        '// Update menu items
        intYear = mintCurrentYear - (DEF_YEAR_MENU_COUNT / 2)
        For i = intYear To (intYear + DEF_YEAR_MENU_COUNT)
            Call ModifyMenu(mlngYearMenu, (i - intYear), MF_BYPOSITION, CLng(i), CStr(i))
        Next i
    Else
        '// If year menu is not created already, create it.
        Call CreateYearMenu
    End If
End Sub

Private Sub ShowMenu(lngMenuType As MenuType)
    Dim blnCancel   As Boolean
    Dim udtPt       As POINTAPI
    Dim lngMenuId   As Long
    
    '// Raise event
    RaiseEvent WillOpenMenu(CInt(lngMenuType), blnCancel)
    If blnCancel Then Exit Sub
    '// Set flag indicating that menu is being shown
    mblnMenuVisible = True
    '// Get current mouse position
    Call GetCursorPos(udtPt)
    '// Show menu according to the type
    If (lngMenuType = MonthMenu) Then
        '// If menu is not created already, create it first.
        If Not IsMenu(mlngMonthMenu) Then Call CreateMonthMenu
        '// Show month menu
        lngMenuId = TrackPopupMenu(mlngMonthMenu, TPM_NONOTIFY Or TPM_RETURNCMD, _
                    udtPt.x, udtPt.y, 0&, hwnd, ByVal 0&)
        '// If user has selected an item from the menu, change current month
        If (lngMenuId > 0) Then CurrentMonth = lngMenuId
    Else
        '// If menu is not created already, create it first.
        If Not IsMenu(mlngYearMenu) Then Call CreateYearMenu
        '// Show month menu
        lngMenuId = TrackPopupMenu(mlngYearMenu, TPM_NONOTIFY Or TPM_RETURNCMD, _
                    udtPt.x, udtPt.y, 0&, hwnd, ByVal 0&)
        '// If user has selected an item from the menu, change current year
        If (lngMenuId > 0) Then CurrentYear = lngMenuId
    End If
    '// Trun off the flag, indicating that menu is closed
    mblnMenuVisible = False
    '// Raise event
    RaiseEvent MenuClosed(CInt(lngMenuType))
End Sub

Private Function IsMouseInMonthArea(x As Single, y As Single) As Boolean
    Dim tr          As RECT
    Dim blnStatus   As Boolean
    
    Call CopyRect(tr, mudtMonthArea)
    blnStatus = ((x >= tr.Left) And (x <= tr.Right)) And _
                ((y >= tr.Top) And (y <= tr.Bottom))
    IsMouseInMonthArea = blnStatus
End Function

Private Function IsMouseInYearArea(x As Single, y As Single) As Boolean
    Dim tr          As RECT
    Dim blnStatus   As Boolean
    
    Call CopyRect(tr, mudtYearArea)
    blnStatus = ((x >= tr.Left) And (x <= tr.Right)) And _
                ((y >= tr.Top) And (y <= tr.Bottom))
    IsMouseInYearArea = blnStatus
End Function

Private Function GetTitleHeight() As Integer
    Dim intTitleHeight  As Integer
    Dim fntTemp         As StdFont
    
    '// Save current font settings
    Set fntTemp = UserControl.Font
    '// Calculate height needed to show title
    Set UserControl.Font = mfntTitleFont
    intTitleHeight = CInt(UserControl.TextHeight("January"))
    '// Add some space on top and bottom
    intTitleHeight = intTitleHeight + 4
    '// Set old font settings
    Set UserControl.Font = fntTemp
    
    GetTitleHeight = intTitleHeight
End Function
