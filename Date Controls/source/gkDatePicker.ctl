VERSION 5.00
Begin VB.UserControl gkDatePicker 
   AutoRedraw      =   -1  'True
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   BeginProperty Font 
      Name            =   "Marlett"
      Size            =   9
      Charset         =   2
      Weight          =   500
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   PropertyPages   =   "gkDatePicker.ctx":0000
   ScaleHeight     =   48
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   150
   ToolboxBitmap   =   "gkDatePicker.ctx":0041
   Begin VB.TextBox txt 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   510
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   960
   End
End
Attribute VB_Name = "gkDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'// ================================================================================
'// Copyright © 2001-2002 by Abdul Gafoor.GK
'// ================================================================================
Option Explicit

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function MessageBoxIndirect Lib "user32" Alias "MessageBoxIndirectA" (lpMsgBoxParams As MSGBOXPARAMS) As Long
Private Type MSGBOXPARAMS
    cbSize As Long
    hwndOwner As Long
    hInstance As Long
    lpszText As String
    lpszCaption As String
    dwStyle As Long
    lpszIcon As String
    dwContextHelpId As Long
    lpfnMsgBoxCallback As Long
    dwLanguageId As Long
End Type

Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOWNORMAL = 1
Private Const LOCALE_SLONGDATE = &H20        '// Long date format string
Private Const LOCALE_SSHORTDATE = &H1F       '// Short date format string

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum dtFormatConstants
    dtLongDate = 1
    dtShortDate = 2
    dtCustom = 3
End Enum

Public Enum dtFormatWhenEditConstants
    [dd/mm/yyyy] = 1
    [mm/dd/yyyy] = 2
    [dd/yyyy/mm] = 3
    [mm/yyyy/dd] = 4
    [yyyy/dd/mm] = 5
    [yyyy/mm/dd] = 6
End Enum

Private Const SEPERATORS = "/-\:"
Private Const PLACE_HOLDERS = "_#- "

Private mblnButtonDown                  As Boolean
Private mblnInButton                    As Boolean
Private mblnCalVisible                  As Boolean
Private mudtButtonArea                  As RECT
Private mstrMask                        As String
Private mstrText                        As String
Private mvarOldDate                     As Variant
Private menuCurrentSection              As dtSectionConstants
Private frm                             As frmCalendar
Private WithEvents clndr                As gkMonth
Attribute clndr.VB_VarHelpID = -1
Private mstrFormat                      As String
Private mstrFormatWhenEdit              As String

Private Enum dtSectionConstants
    dtDaySection
    dtMonthSection
    dtYearSection
    dtInvalid
End Enum

Private Enum dtDateValidationConstants
    dtValidDate
    dtInvalidDate
    dtNullDate
End Enum

Private Const mdefAppearance = dt3D
Private Const mdefBorderStyle = dtFixedSingle
Private Const mdefCalendarTitleBackColor = SystemColorConstants.vbButtonFace
Private Const mdefCalendarTitleForeColor = SystemColorConstants.vbButtonText
Private Const mdefCalendarMonthForeColor = SystemColorConstants.vbButtonText
Private Const mdefCalendarMonthBackColor = SystemColorConstants.vbWindowBackground
Private Const mdefCalendarActiveDayForeColor = SystemColorConstants.vbButtonText
Private Const mdefCalendarActiveDayBackColor = SystemColorConstants.vbWindowBackground
Private Const mdefCalendarDayHeaderForeColor = SystemColorConstants.vbButtonText
Private Const mdefCalendarDayHeaderBackColor = SystemColorConstants.vbButtonFace
Private Const mdefMinDate As Date = "01/01/1901"
Private Const mdefMaxDate As Date = "31/12/2099"
Private Const mdefCustomFormat = "dd/MM/yyyy"
Private Const mdefFormat = dtFormatConstants.dtShortDate
Private Const mdefAllowNull = True
Private Const mdefFormatWhenEdit = dtFormatWhenEditConstants.[dd/mm/yyyy]
Private Const mdefSeperator = "/"
Private Const mdefPlaceHolder = "_"

Private menuAppearance                  As dtAppearanceConstants
Private menuBorderStyle                 As dtBorderStyleConstants
Private mlngCalendarTitleBackColor      As OLE_COLOR
Private mlngCalendarTitleForeColor      As OLE_COLOR
Private mlngCalendarMonthForeColor      As OLE_COLOR
Private mlngCalendarMonthBackColor      As OLE_COLOR
Private mlngCalendarActiveDayForeColor  As OLE_COLOR
Private mlngCalendarActiveDayBackColor  As OLE_COLOR
Private mlngCalendarDayHeaderForeColor  As OLE_COLOR
Private mlngCalendarDayHeaderBackColor  As OLE_COLOR
Private mdtMaxDate                      As Date
Private mdtMinDate                      As Date
Private mvarValue                       As Variant
Private mstrCustomFormat                As String
Private menuFormat                      As dtFormatConstants
Private mblnAllowNull                   As Boolean
Private menuFormatWhenEdit              As dtFormatWhenEditConstants
Private mstrSeperator                   As String
Private mstrPlaceHolder                 As String

Public Event Click()
Public Event DblClick()
Public Event Change()
Attribute Change.VB_MemberFlags = "200"
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)


Public Property Get Appearance() As dtAppearanceConstants
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Appearance = menuAppearance
End Property

Public Property Let Appearance(ByVal vData As dtAppearanceConstants)
    If (vData = dt3D) Or (vData = dtFlat) Then
        menuAppearance = vData
        PropertyChanged "Appearance"
        Call Refresh
    End If
End Property

Public Property Get BorderStyle() As dtBorderStyleConstants
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = menuBorderStyle
End Property

Public Property Let BorderStyle(ByVal vData As dtBorderStyleConstants)
    menuBorderStyle = vData
    PropertyChanged "BorderStyle"
    Call Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = txt.ForeColor
End Property

Public Property Let ForeColor(ByVal vData As OLE_COLOR)
    txt.ForeColor() = vData
    PropertyChanged "ForeColor"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = txt.BackColor
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
    txt.BackColor() = vData
    PropertyChanged "BackColor"
    Call Refresh
End Property

Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = txt.Font
End Property

Public Property Set Font(ByVal vData As Font)
    Set txt.Font = vData
    PropertyChanged "Font"
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = txt.FontBold
End Property

Public Property Let FontBold(ByVal vData As Boolean)
    txt.FontBold() = vData
    PropertyChanged "FontBold"
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = txt.FontItalic
End Property

Public Property Let FontItalic(ByVal vData As Boolean)
    txt.FontItalic() = vData
    PropertyChanged "FontItalic"
End Property

Public Property Get FontName() As String
Attribute FontName.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontName.VB_MemberFlags = "400"
    FontName = txt.FontName
End Property

Public Property Let FontName(ByVal vData As String)
    txt.FontName() = vData
    PropertyChanged "FontName"
End Property

Public Property Get FontSize() As Single
Attribute FontSize.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = txt.FontSize
End Property

Public Property Let FontSize(ByVal vData As Single)
    txt.FontSize() = vData
    PropertyChanged "FontSize"
End Property

Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontStrikethru.VB_MemberFlags = "400"
    FontStrikethru = txt.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal vData As Boolean)
    txt.FontStrikethru() = vData
    PropertyChanged "FontStrikethru"
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = txt.FontUnderline
End Property

Public Property Let FontUnderline(ByVal vData As Boolean)
    txt.FontUnderline() = vData
    PropertyChanged "FontUnderline"
End Property

Public Property Get CalendarTitleBackColor() As OLE_COLOR
Attribute CalendarTitleBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CalendarTitleBackColor = mlngCalendarTitleBackColor
End Property

Public Property Let CalendarTitleBackColor(ByVal vData As OLE_COLOR)
    mlngCalendarTitleBackColor = vData
    PropertyChanged "CalendarTitleBackColor"
End Property

Public Property Get CalendarTitleForeColor() As OLE_COLOR
Attribute CalendarTitleForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CalendarTitleForeColor = mlngCalendarTitleForeColor
End Property

Public Property Let CalendarTitleForeColor(ByVal vData As OLE_COLOR)
    mlngCalendarTitleForeColor = vData
    PropertyChanged "CalendarTitleForeColor"
End Property

Public Property Get CalendarMonthForeColor() As OLE_COLOR
Attribute CalendarMonthForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CalendarMonthForeColor = mlngCalendarMonthForeColor
End Property

Public Property Let CalendarMonthForeColor(ByVal vData As OLE_COLOR)
    mlngCalendarMonthForeColor = vData
    PropertyChanged "CalendarMonthForeColor"
End Property

Public Property Get CalendarMonthBackColor() As OLE_COLOR
Attribute CalendarMonthBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CalendarMonthBackColor = mlngCalendarMonthBackColor
End Property

Public Property Let CalendarMonthBackColor(ByVal vData As OLE_COLOR)
    mlngCalendarMonthBackColor = vData
    PropertyChanged "CalendarMonthBackColor"
End Property

Public Property Get CalendarActiveDayForeColor() As OLE_COLOR
Attribute CalendarActiveDayForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CalendarActiveDayForeColor = mlngCalendarActiveDayForeColor
End Property

Public Property Let CalendarActiveDayForeColor(ByVal vData As OLE_COLOR)
    mlngCalendarActiveDayForeColor = vData
    PropertyChanged "CalendarActiveDayForeColor"
End Property

Public Property Get CalendarActiveDayBackColor() As OLE_COLOR
Attribute CalendarActiveDayBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CalendarActiveDayBackColor = mlngCalendarActiveDayBackColor
End Property

Public Property Let CalendarActiveDayBackColor(ByVal vData As OLE_COLOR)
    mlngCalendarActiveDayBackColor = vData
    PropertyChanged "CalendarActiveDayBackColor"
End Property

Public Property Get CalendarDayHeaderForeColor() As OLE_COLOR
Attribute CalendarDayHeaderForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CalendarDayHeaderForeColor = mlngCalendarDayHeaderForeColor
End Property

Public Property Let CalendarDayHeaderForeColor(ByVal vData As OLE_COLOR)
    mlngCalendarDayHeaderForeColor = vData
    PropertyChanged "CalendarDayHeaderForeColor"
End Property

Public Property Get CalendarDayHeaderBackColor() As OLE_COLOR
Attribute CalendarDayHeaderBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CalendarDayHeaderBackColor = mlngCalendarDayHeaderBackColor
End Property

Public Property Let CalendarDayHeaderBackColor(ByVal vData As OLE_COLOR)
    mlngCalendarDayHeaderBackColor = vData
    PropertyChanged "CalendarDayHeaderBackColor"
End Property

Public Property Get MaxDate() As Date
    MaxDate = mdtMaxDate
End Property

Public Property Let MaxDate(ByVal vData As Date)
    If (vData > mdtMinDate) Then
        mdtMaxDate = vData
        PropertyChanged "MaxDate"
    End If
End Property

Public Property Get MinDate() As Date
    MinDate = mdtMinDate
End Property

Public Property Let MinDate(ByVal vData As Date)
    If (vData < mdtMaxDate) Then
        mdtMinDate = vData
        PropertyChanged "MinDate"
    End If
End Property

Public Property Get Value() As Variant
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "34"
    Value = mvarValue
End Property

Public Property Let Value(ByVal vData As Variant)
    If IsDate(vData) Then
        mvarValue = CDate(vData)
    ElseIf (Len(CStr(vData)) = 0) Then
        mvarValue = IIf(mblnAllowNull, vData, Date)
    Else
        Exit Property
    End If
    
    If (Len(CStr(mvarValue)) > 0) Then
        txt.Text = VBA.Format$(mvarValue, mstrFormat)
        mstrText = VBA.Format$(mvarValue, mstrFormatWhenEdit)
    Else
        mstrText = GetMaskString()
        txt.Text = IIf(Ambient.UserMode, mstrText, "")
    End If
    
    PropertyChanged "Value"
End Property

Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_ProcData.VB_Invoke_Property = ";Behavior"
    RightToLeft = txt.RightToLeft
End Property

Public Property Let RightToLeft(ByVal vData As Boolean)
    txt.RightToLeft() = vData
    PropertyChanged "RightToLeft"
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Locked = txt.Locked
End Property

Public Property Let Locked(ByVal vData As Boolean)
    txt.Locked() = vData
    PropertyChanged "Locked"
End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_UserMemId = -515
Attribute hwnd.VB_MemberFlags = "400"
    hwnd = txt.hwnd
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal vData As Boolean)
    UserControl.Enabled() = vData
    txt.Enabled() = vData
    PropertyChanged "Enabled"
    Call Refresh
End Property

Public Property Get CustomFormat() As String
Attribute CustomFormat.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CustomFormat = mstrCustomFormat
End Property

Public Property Let CustomFormat(ByVal vData As String)
    mstrCustomFormat = vData
    PropertyChanged "CustomFormat"
End Property

Public Property Get Format() As dtFormatConstants
Attribute Format.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Format = menuFormat
End Property

Public Property Let Format(ByVal vData As dtFormatConstants)
    menuFormat = vData
    PropertyChanged "Format"
End Property

Public Property Get AllowNull() As Boolean
Attribute AllowNull.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AllowNull = mblnAllowNull
End Property

Public Property Let AllowNull(ByVal vData As Boolean)
    mblnAllowNull = vData
    PropertyChanged "AllowNull"
    
    If Not mblnAllowNull Then
        If (Value = "") Then
            Value = Date
        End If
    End If
End Property

Public Property Get FormatWhenEdit() As dtFormatWhenEditConstants
Attribute FormatWhenEdit.VB_ProcData.VB_Invoke_Property = ";Behavior"
    FormatWhenEdit = menuFormatWhenEdit
End Property

Public Property Let FormatWhenEdit(ByVal vData As dtFormatWhenEditConstants)
    menuFormatWhenEdit = vData
    PropertyChanged "FormatWhenEdit"
    mstrMask = GetMaskString()
End Property

Public Property Get Seperator() As String
Attribute Seperator.VB_ProcData.VB_Invoke_Property = ";Misc"
    Seperator = mstrSeperator
End Property

Public Property Let Seperator(ByVal vData As String)
    If Ambient.UserMode Then Err.Raise 382      '// In run-time raise an error
    If (Len(vData) = 1) Then
        If (vData <> mstrPlaceHolder) Then
            If InStr(1, SEPERATORS, vData) Then
                mstrSeperator = vData
                PropertyChanged "Seperator"
            End If
        End If
    End If
End Property

Public Property Get PlaceHolder() As String
Attribute PlaceHolder.VB_ProcData.VB_Invoke_Property = ";Misc"
    PlaceHolder = mstrPlaceHolder
End Property

Public Property Let PlaceHolder(ByVal vData As String)
    If Ambient.UserMode Then Err.Raise 382      '// In run-time raise an error
    If (Len(vData) = 1) Then
        If (vData <> mstrSeperator) Then
            If InStr(1, PLACE_HOLDERS, vData) Then
                mstrPlaceHolder = vData
                PropertyChanged "PlaceHolder"
            End If
        End If
    End If
End Property

Public Property Get OLEDragMode() As OLEDragConstants
    OLEDragMode = txt.OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal vData As OLEDragConstants)
    txt.OLEDragMode() = vData
    PropertyChanged "OLEDragMode"
End Property

Public Property Get OLEDropMode() As OLEDropConstants
    OLEDropMode = txt.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal vData As OLEDropConstants)
    txt.OLEDropMode() = vData
    PropertyChanged "OLEDropMode"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = txt.MousePointer
End Property

Public Property Let MousePointer(ByVal vData As MousePointerConstants)
    txt.MousePointer() = vData
    PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = txt.MouseIcon
End Property

Public Property Set MouseIcon(ByVal vData As Picture)
    Set txt.MouseIcon = vData
    PropertyChanged "MouseIcon"
End Property

Public Function IsNull() As Boolean
    IsNull = IsEmpty(Me.Value) Or (Me.Value = "")
End Function

Public Sub About()
Attribute About.VB_UserMemId = -552
    frmAbout.ActivateForm "DATEPICKER"
End Sub

Private Sub clndr_DateChanged(OldDate As Date, NewDate As Date)
    Value = NewDate
End Sub

Private Sub txt_Change()
    RaiseEvent Change
End Sub

Private Sub txt_GotFocus()
    Dim dtDate As Date

    With txt
        .MaxLength = Len(mstrFormatWhenEdit)
        If (Len(CStr(Value)) > 0) Then
            Call GetDate(dtDate)
            .Text = VBA.Format$(dtDate, mstrFormatWhenEdit)
            mvarOldDate = dtDate
        Else
            If mblnAllowNull Then
                .Text = mstrMask
                mvarOldDate = ""
            Else
                .Text = VBA.Format$(VBA.Date, mstrFormatWhenEdit)
                mvarOldDate = .Text
            End If
        End If
        mstrText = .Text
    End With
    Call MakeSelection(GetCurrentSection())
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    
    If (KeyCode = vbKeyF4) Or ((KeyCode = vbKeyDown) And (Shift = vbAltMask)) Then
        Call ShowCalendar
    ElseIf (KeyCode = vbKeyDelete) Or (KeyCode = vbKeyBack) Then
        Call DeleteNumber(KeyCode)
    ElseIf (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Then
        Call MoveInsertionPoint(KeyCode)
    ElseIf (KeyCode = vbKeyUp) Or (KeyCode = vbKeyDown) Then
        Call ChangeSectionValue(KeyCode)
    ElseIf (KeyCode = vbKeyHome) Then
        txt.SelStart = 0
        Call MakeSelection(GetCurrentSection())
    ElseIf (KeyCode = vbKeyEnd) Then
        txt.SelStart = txt.MaxLength
        Call MakeSelection(GetCurrentSection())
    End If
    
    KeyCode = 0
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    
    If (KeyAscii >= vbKeySpace) Then
        If (KeyAscii >= vbKey0) And (KeyAscii <= vbKey9) Then
            Call InsertNumber(Chr$(KeyAscii))
        ElseIf (KeyAscii = Asc(mstrSeperator)) Then
            Call MoveToNextSection
        End If
    End If
    
    KeyAscii = 0
End Sub

Private Sub txt_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    
    Select Case KeyCode
        Case vbKeyInsert
            Value = VBA.Date
            Call txt_GotFocus
    End Select
End Sub

Private Sub txt_LostFocus()
    Dim strDate As String
    Dim dtDate As Date
    Dim enuDtType As dtDateValidationConstants
    
    With txt
        .MaxLength = 0
        
        Select Case GetDate(dtDate)
            Case dtValidDate
                Value = dtDate
            Case dtInvalidDate
                Value = mvarOldDate
            Case dtNullDate
                Value = IIf(mblnAllowNull, "", mvarOldDate)
        End Select
    End With
End Sub

Private Sub txt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
    If (GetCurrentSection() <> menuCurrentSection) Then Call MakeSelection
End Sub

Private Sub txt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub txt_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub txt_Click()
    RaiseEvent Click
End Sub

Private Sub txt_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub txt_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Public Sub OLEDrag()
    txt.OLEDrag
End Sub

Private Sub txt_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub txt_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub txt_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub txt_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub txt_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_InitProperties()
    menuAppearance = mdefAppearance
    menuBorderStyle = mdefBorderStyle
    mlngCalendarTitleBackColor = mdefCalendarTitleBackColor
    mlngCalendarTitleForeColor = mdefCalendarTitleForeColor
    mlngCalendarMonthForeColor = mdefCalendarMonthForeColor
    mlngCalendarMonthBackColor = mdefCalendarMonthBackColor
    mlngCalendarActiveDayForeColor = mdefCalendarActiveDayForeColor
    mlngCalendarActiveDayBackColor = mdefCalendarActiveDayBackColor
    mlngCalendarDayHeaderForeColor = mdefCalendarDayHeaderForeColor
    mlngCalendarDayHeaderBackColor = mdefCalendarDayHeaderBackColor
    mdtMaxDate = mdefMaxDate
    mdtMinDate = mdefMinDate
    mstrSeperator = mdefSeperator
    mstrPlaceHolder = mdefPlaceHolder
    menuFormat = mdefFormat
    mstrCustomFormat = mdefCustomFormat
    mblnAllowNull = mdefAllowNull
    menuFormatWhenEdit = mdefFormatWhenEdit
    mstrFormat = GetFormat()
    mstrFormatWhenEdit = GetEditFormatString()
    Value = Date
    
    Call Refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
    
    If (Button = vbLeftButton) Then
        If IsMouseInButtonArea(x, y) Then
            mblnButtonDown = True
            Call Refresh
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
    
    If (Button = vbLeftButton) Then
        If mblnButtonDown Then
            mblnButtonDown = False
            Call Refresh
        End If
        
        If IsMouseInButtonArea(x, y) Then
            Call ShowCalendar
        End If
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    menuAppearance = PropBag.ReadProperty("Appearance", mdefAppearance)
    menuBorderStyle = PropBag.ReadProperty("BorderStyle", mdefBorderStyle)
    txt.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txt.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txt.FontBold = PropBag.ReadProperty("FontBold", Ambient.Font.Bold)
    txt.FontItalic = PropBag.ReadProperty("FontItalic", Ambient.Font.Italic)
    txt.FontName = PropBag.ReadProperty("FontName", Ambient.Font.Name)
    txt.FontSize = PropBag.ReadProperty("FontSize", Ambient.Font.SIZE)
    txt.FontStrikethru = PropBag.ReadProperty("FontStrikethru", Ambient.Font.Strikethrough)
    txt.FontUnderline = PropBag.ReadProperty("FontUnderline", Ambient.Font.Underline)
    mlngCalendarTitleForeColor = PropBag.ReadProperty("CalendarTitleForeColor", mdefCalendarTitleForeColor)
    mlngCalendarTitleBackColor = PropBag.ReadProperty("CalendarTitleBackColor", mdefCalendarTitleBackColor)
    mlngCalendarMonthForeColor = PropBag.ReadProperty("CalendarMonthForeColor", mdefCalendarMonthForeColor)
    mlngCalendarMonthBackColor = PropBag.ReadProperty("CalendarMonthBackColor", mdefCalendarMonthBackColor)
    mlngCalendarActiveDayForeColor = PropBag.ReadProperty("CalendarActiveDayForeColor", mdefCalendarActiveDayForeColor)
    mlngCalendarActiveDayBackColor = PropBag.ReadProperty("CalendarActiveDayBackColor", mdefCalendarActiveDayBackColor)
    mlngCalendarDayHeaderForeColor = PropBag.ReadProperty("CalendarDayHeaderForeColor", mdefCalendarDayHeaderForeColor)
    mlngCalendarDayHeaderBackColor = PropBag.ReadProperty("CalendarDayHeaderBackColor", mdefCalendarDayHeaderBackColor)
    mdtMaxDate = PropBag.ReadProperty("MaxDate", mdefMaxDate)
    mdtMinDate = PropBag.ReadProperty("MinDate", mdefMinDate)
    txt.Locked = PropBag.ReadProperty("Locked", False)
    txt.RightToLeft = PropBag.ReadProperty("RightToLeft", False)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    txt.Enabled = PropBag.ReadProperty("Enabled", True)
    mstrCustomFormat = PropBag.ReadProperty("CustomFormat", mdefCustomFormat)
    mblnAllowNull = PropBag.ReadProperty("AllowNull", mdefAllowNull)
    mstrSeperator = PropBag.ReadProperty("Seperator", mdefSeperator)
    mstrPlaceHolder = PropBag.ReadProperty("PlaceHolder", mdefPlaceHolder)
    txt.OLEDragMode = PropBag.ReadProperty("OLEDragMode", OLEDragConstants.vbOLEDragManual)
    txt.OLEDropMode = PropBag.ReadProperty("OLEDropMode", OLEDropConstants.vbOLEDropNone)
    txt.MousePointer = PropBag.ReadProperty("MousePointer", MousePointerConstants.vbDefault)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    
    menuFormat = PropBag.ReadProperty("Format", mdefFormat)
    mstrFormat = GetFormat()
    
    menuFormatWhenEdit = PropBag.ReadProperty("FormatWhenEdit", mdefFormatWhenEdit)
    mstrFormatWhenEdit = GetEditFormatString()
    mstrMask = GetMaskString()
    
    mvarValue = PropBag.ReadProperty("Value", Date)
    If (Len(CStr(mvarValue)) = 0) Then
        If mblnAllowNull Then
            txt.Text = IIf(Ambient.UserMode, GetMaskString(), "")
        Else
            mvarValue = VBA.Date
            txt.Text = VBA.Format$(mvarValue, mstrFormat)
            mstrText = VBA.Format$(mvarValue, mstrFormatWhenEdit)
        End If
    Else
        txt.Text = VBA.Format$(mvarValue, mstrFormat)
        mstrText = VBA.Format$(mvarValue, mstrFormatWhenEdit)
    End If
    
    Call Refresh
End Sub

Private Sub UserControl_Resize()
    Call Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Appearance", menuAppearance, mdefAppearance)
    Call PropBag.WriteProperty("BorderStyle", menuBorderStyle, mdefBorderStyle)
    Call PropBag.WriteProperty("ForeColor", txt.ForeColor, &H80000008)
    Call PropBag.WriteProperty("BackColor", txt.BackColor, &H80000005)
    Call PropBag.WriteProperty("FontBold", txt.FontBold, Ambient.Font.Bold)
    Call PropBag.WriteProperty("FontItalic", txt.FontItalic, Ambient.Font.Italic)
    Call PropBag.WriteProperty("FontName", txt.FontName, Ambient.Font.Name)
    Call PropBag.WriteProperty("FontSize", txt.FontSize, Ambient.Font.SIZE)
    Call PropBag.WriteProperty("FontStrikethru", txt.FontStrikethru, Ambient.Font.Strikethrough)
    Call PropBag.WriteProperty("FontUnderline", txt.FontUnderline, Ambient.Font.Underline)
    Call PropBag.WriteProperty("CalendarTitleBackColor", mlngCalendarTitleBackColor, mdefCalendarTitleBackColor)
    Call PropBag.WriteProperty("CalendarTitleForeColor", mlngCalendarTitleForeColor, mdefCalendarTitleForeColor)
    Call PropBag.WriteProperty("CalendarMonthForeColor", mlngCalendarMonthForeColor, mdefCalendarMonthForeColor)
    Call PropBag.WriteProperty("CalendarMonthBackColor", mlngCalendarMonthBackColor, mdefCalendarMonthBackColor)
    Call PropBag.WriteProperty("CalendarActiveDayForeColor", mlngCalendarActiveDayForeColor, mdefCalendarActiveDayForeColor)
    Call PropBag.WriteProperty("CalendarActiveDayBackColor", mlngCalendarActiveDayBackColor, mdefCalendarActiveDayBackColor)
    Call PropBag.WriteProperty("CalendarDayHeaderForeColor", mlngCalendarDayHeaderForeColor, mdefCalendarDayHeaderForeColor)
    Call PropBag.WriteProperty("CalendarDayHeaderBackColor", mlngCalendarDayHeaderBackColor, mdefCalendarDayHeaderBackColor)
    Call PropBag.WriteProperty("MaxDate", mdtMaxDate, mdefMaxDate)
    Call PropBag.WriteProperty("MinDate", mdtMinDate, mdefMinDate)
    Call PropBag.WriteProperty("Locked", txt.Locked, False)
    Call PropBag.WriteProperty("RightToLeft", txt.RightToLeft, False)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Format", menuFormat, mdefFormat)
    Call PropBag.WriteProperty("CustomFormat", mstrCustomFormat, mdefCustomFormat)
    Call PropBag.WriteProperty("AllowNull", mblnAllowNull, mdefAllowNull)
    Call PropBag.WriteProperty("FormatWhenEdit", menuFormatWhenEdit, mdefFormatWhenEdit)
    Call PropBag.WriteProperty("Seperator", mstrSeperator, mdefSeperator)
    Call PropBag.WriteProperty("PlaceHolder", mstrPlaceHolder, mdefPlaceHolder)
    Call PropBag.WriteProperty("Value", mvarValue, Date)
    Call PropBag.WriteProperty("OLEDragMode", txt.OLEDragMode, OLEDragConstants.vbOLEDragManual)
    Call PropBag.WriteProperty("OLEDropMode", txt.OLEDropMode, OLEDropConstants.vbOLEDropNone)
    Call PropBag.WriteProperty("MousePointer", txt.MousePointer, MousePointerConstants.vbDefault)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
End Sub

Private Sub Refresh()
    Dim tr              As RECT
    Dim intButWidth     As Integer
    Dim intOffset       As Integer
    Dim bln3D           As Boolean
    Dim blnButStatus    As Boolean
    Dim clsMDC          As New cMemoryDC
    
    intButWidth = 18
    
    intOffset = IIf((menuAppearance = dt3D), 3, 2)
    With txt
        .Left = intOffset
        .Top = intOffset
        .Width = ScaleWidth - (intButWidth + intOffset * 2)
        .Height = ScaleHeight - intOffset * 2
    End With
    
    With clsMDC
        '// Start drawing in memory dc
        Call .StartDrawing(hdc, ScaleWidth, ScaleHeight)
        
        '// Set font for the device context
        Set .Font = UserControl.Font
        '// Fill background of control with backcolor of text
        Call .FillRect(0, 0, ScaleWidth, ScaleHeight, txt.BackColor)
        
        '// Check for 3d state
        bln3D = (menuAppearance = dt3D)
        '// Cave button status according to 3D state
        blnButStatus = IIf(bln3D, mblnButtonDown, True)
        '// Set position of drop-down button
        intOffset = IIf(bln3D, 2, 0)
        With tr
            .Left = ScaleWidth - intButWidth
            .Top = intOffset
            .Right = ScaleWidth - intOffset
            .Bottom = ScaleHeight - intOffset
        End With
        '// Save button area
        Call CopyRect(mudtButtonArea, tr)
        '// Print drop-down sign on the button
        Call .FillRect(tr.Left, tr.Top, tr.Right - tr.Left, tr.Bottom - tr.Top, vbButtonFace)
        '// Print drop-down sign
        intOffset = IIf(bln3D And mblnButtonDown, 1, 0)
        If UserControl.Enabled Then
            .ForeColor = vbButtonText
            Call .DrawText((tr.Left + intOffset + 1), (tr.Top + intOffset), _
                        (tr.Right - tr.Left), (tr.Bottom - tr.Top), _
                        "6", mdcTextSingleLineCenter)
        End If
        '// Draw the button
        Call .DrawBorder(tr.Left, tr.Top, tr.Right - tr.Left, tr.Bottom - tr.Top, mdcRaised, blnButStatus, , blnButStatus)
        
        '// Draw border according to appearance setting
        Call .DrawBorder(0&, 0&, ScaleWidth, ScaleHeight, mdcSunken, (menuAppearance = dtFlat))
        '// If edge style is flat and border style is fixed single, draw a frame
        If (menuAppearance = dtFlat) And (menuBorderStyle = dtFixedSingle) Then
            Call .DrawFrame(0, 0, ScaleWidth, ScaleHeight, SystemColorConstants.vbWindowFrame)
        End If
        
        Call .StopDrawing(0&, 0&, ScaleWidth, ScaleHeight)
        
        '// Draw disabled dropdown sign if the control is in disabled state
        If Not UserControl.Enabled Then
            Call .DrawTextDirect(UserControl.hdc, (tr.Left + intOffset) + 1, _
                        (tr.Top + intOffset) + 1, (tr.Right - tr.Left), _
                        (tr.Bottom - tr.Top), "6", mdcTextSingleLineCenter, _
                        SystemColorConstants.vb3DHighlight)
            Call .DrawTextDirect(UserControl.hdc, (tr.Left + intOffset), _
                        (tr.Top + intOffset), (tr.Right - tr.Left), _
                        (tr.Bottom - tr.Top), "6", mdcTextSingleLineCenter, _
                        SystemColorConstants.vb3DShadow)
        End If
    End With
    
    '// Refresh the control, just to make sure
    UserControl.Refresh
End Sub

Private Function IsMouseInButtonArea(ByVal x As Long, ByVal y As Long) As Boolean
    Dim tr          As RECT
    Dim blnStatus   As Boolean
    
    Call CopyRect(tr, mudtButtonArea)
    blnStatus = ((x >= tr.Left) And (x <= tr.Right)) And _
                ((y >= tr.Top) And (y <= tr.Bottom))
    IsMouseInButtonArea = blnStatus
End Function

Private Sub ShowCalendar()
    Dim varValue    As Variant
    Dim udtPos      As RECT
    
    '// Call 'LostFocus' event of textbox since it is behaving strangely sometimes
    Call txt_LostFocus
    
    '// Save old date
    varValue = Me.Value
    '// Get current position of the control
    Call GetWindowRect(hwnd, udtPos)
    '// Load calendar form and set intitial properties
    Set frm = New frmCalendar
    Load frm
    With frm
        Set clndr = .clndr
        With .clndr
            '// Turn off tooltip since tooltip form is not being shown modally
            .ShowToolTip = False
            '// Set other properties
            .AutoRefresh = False
            .TitleBackColor = mlngCalendarTitleBackColor
            .TitleForeColor = mlngCalendarTitleForeColor
            .CurrentMonthBackColor = mlngCalendarMonthBackColor
            .CurrentMonthForeColor = mlngCalendarMonthForeColor
            .DayHeaderBackColor = mlngCalendarDayHeaderBackColor
            .DayHeaderForeColor = mlngCalendarDayHeaderForeColor
            .DayHeaderFormat = dtSingleLetter
            .ActiveDayBackColor = mlngCalendarActiveDayBackColor
            .ActiveDayForeColor = mlngCalendarActiveDayForeColor
            .MinDate = mdtMinDate
            .MaxDate = mdtMaxDate
            .CurrentDate = IIf((Len(CStr(Me.Value)) > 0), Me.Value, Date)
            .AutoRefresh = True
        End With
        '// show form
        Call .ActivateForm(udtPos.Left, udtPos.Top, udtPos.Right, udtPos.Bottom)
        '// If not canceled then set selected date.
        '// Otherwise restore old date
        If Not .Canceled Then
            Me.Value = .clndr.CurrentDate
        Else
            Me.Value = varValue
        End If
    End With
    '// Unload form to release memory used by the form
    Unload frm
    
    '// Call 'GotFocus' event of textbox since it is behaving strangely sometimes
    Call txt_GotFocus
    
    Set frm = Nothing
    Set clndr = Nothing
End Sub

Private Function GetMaskString() As String
    Dim i()         As Integer
    Dim strMask     As String
    
    ReDim i(1 To 3)
    Select Case menuFormatWhenEdit
        Case [dd/mm/yyyy], [mm/dd/yyyy]
            i(1) = 2: i(2) = 2: i(3) = 4
        Case [dd/yyyy/mm], [mm/yyyy/dd]
            i(1) = 2: i(2) = 4: i(3) = 2
        Case [yyyy/dd/mm], [yyyy/mm/dd]
            i(1) = 4: i(2) = 2: i(3) = 2
    End Select
    
    GetMaskString = String$(i(1), mstrPlaceHolder) & mstrSeperator & _
                    String$(i(2), mstrPlaceHolder) & mstrSeperator & _
                    String$(i(3), mstrPlaceHolder)
End Function

Private Sub InsertNumber(ByVal sChar As String)
    Dim intStart    As Integer
    Dim intEnd      As Integer
    Dim strText     As String
    Dim strNewText  As String
    Dim intPos      As Integer
    Dim i           As Integer
    
    With txt
        '// If insertion point is at maximum length, exit the procedure
        If (.SelStart = .MaxLength) Then Exit Sub
        '// If some text has been selected, delete it first
        If (.SelLength > 0) Then Call DeleteSelection
        '// Get current section information
        Call GetSectionInfo(GetCurrentSection(), intStart, intEnd)
        '// Get position of insertion point from the beginning of current section
        intPos = .SelStart - intStart
        '// Get section text
        strText = Mid$(mstrText, intStart + 1, (intEnd - intStart))
        '// Parse the section text to form new text
        strNewText = Mid$(strText, 1, intPos) & sChar & _
                     Replace(strText, mstrPlaceHolder, "", intPos + 1, 1)
        strNewText = Left$(strNewText, Len(strText))
        '// Apply the parsed string in main date string
        Mid$(mstrText, intStart + 1, Len(strNewText)) = strNewText
        '// Update text box
        .Text = mstrText
        .SelStart = intPos + intStart + 1
    End With
    
    Call MoveInsertionPoint(0)
End Sub

Private Sub DeleteNumber(ByVal iDeleteMode As Integer)
    Dim intStart    As Integer
    Dim intEnd      As Integer
    Dim strText     As String
    Dim strNewText  As String
    Dim intPos      As Integer

    With txt
        '// If a selection has been made, delete it
        If (.SelLength > 0) Then
            Call DeleteSelection
            Exit Sub
        End If
        
        If (iDeleteMode = vbKeyDelete) Then
            '// If insertion point is at the end, exit the function
            If (.SelStart = .MaxLength) Then Exit Sub
            '// If the next letter is a seperator, move insertion point one step forward
            If (Mid$(mstrText, .SelStart + 1, 1) = mstrSeperator) Then
                .SelStart = .SelStart + 1
                menuCurrentSection = GetCurrentSection()
            End If
            '// Get current section information
            Call GetSectionInfo(GetCurrentSection(), intStart, intEnd)
            '// Get position of insertion point from the beginning of current section
            intPos = .SelStart - intStart
            '// Get section text
            strText = Mid$(mstrText, intStart + 1, (intEnd - intStart))
            '// Parse the section text to form new text
            strNewText = Mid$(strText, 1, intPos) & _
                         Mid$(strText, intPos + 2, (intEnd - intStart) - intPos)
            strNewText = strNewText & mstrPlaceHolder
            '// Apply the parsed string in main date string
            Mid$(mstrText, intStart + 1, Len(strNewText)) = strNewText
        Else
            '// If insertion point is at the end, exit the function
            If (.SelStart = 0) Then Exit Sub
            '// If the letter just before is a seperator, move insertion point one step backward
            If (Mid$(mstrText, .SelStart, 1) = mstrSeperator) Then
                .SelStart = .SelStart - 1
                menuCurrentSection = GetCurrentSection()
            End If
            '// Get current section information
            Call GetSectionInfo(GetCurrentSection(), intStart, intEnd)
            '// Get position of insertion point from the beginning of current section
            intPos = .SelStart - intStart
            '// Get section text
            strText = Mid$(mstrText, intStart + 1, (intEnd - intStart))
            '// Parse the section text to form new text
            strNewText = Mid$(strText, 1, intPos - 1) & _
                         Mid$(strText, intPos + 1, (intEnd - intStart) - intPos)
            strNewText = strNewText & mstrPlaceHolder
            '// Apply the parsed string in main date string
            Mid$(mstrText, intStart + 1, Len(strNewText)) = strNewText
        End If
        
        .Text = mstrText
        .SelStart = intPos + intStart - IIf((iDeleteMode = vbKeyBack), 1, 0)
    End With
End Sub

Private Sub DeleteSelection()
    Dim strDel      As String
    Dim strChar     As String
    Dim intPos      As Integer
    Dim intLen      As Integer
    Dim i           As Integer
    
    With txt
        intPos = .SelStart
        intLen = .SelLength
        
        For i = (intPos + 1) To (intPos + intLen)
            strChar = Mid$(mstrText, i, 1)
            strDel = strDel & IIf((strChar = mstrSeperator), mstrSeperator, mstrPlaceHolder)
        Next i
        Mid$(mstrText, intPos + 1, intLen) = strDel
        
        .Text = mstrText
        .SelStart = intPos
    End With
End Sub

Private Sub MakeSelection( _
                Optional eSection As dtSectionConstants = dtInvalid, _
                Optional ByVal iStart As Integer = -1, _
                Optional ByVal iEnd As Integer = -1)
    
    Dim intStart    As Integer
    Dim intEnd      As Integer
    
    If (iStart = -1) Or (iEnd = -1) Then
        If (eSection = dtInvalid) Then
            Call GetSectionInfo(GetCurrentSection(), intStart, intEnd)
        Else
            Call GetSectionInfo(eSection, intStart, intEnd)
        End If
    Else
        intStart = iStart
        intEnd = iEnd
    End If
    
    txt.SelStart = intStart
    txt.SelLength = (intEnd - intStart)
    
    menuCurrentSection = GetCurrentSection()
End Sub

Private Function GetCurrentSection() As dtSectionConstants
    Dim strCurPosChar   As String
    Dim intPos          As Integer
    
    With txt
        If (.SelLength > 0) Then
            If (InStr(1, .SelText, mstrSeperator) > 0) Then
                GetCurrentSection = dtInvalid
                Exit Function
            End If
        End If
        
        intPos = .SelStart '+ IIf((.SelStart = .MaxLength), 0, 1)
        If (.SelStart <> .MaxLength) Or (.SelStart = 0) Then
            intPos = intPos + 1
        End If
        strCurPosChar = Mid$(mstrFormatWhenEdit, intPos, 1)
        If (strCurPosChar = mstrSeperator) Then
            strCurPosChar = Mid$(mstrFormatWhenEdit, intPos - 1, 1)
        End If
    End With
    
    Select Case strCurPosChar
        Case "d": GetCurrentSection = dtDaySection
        Case "m": GetCurrentSection = dtMonthSection
        Case "y": GetCurrentSection = dtYearSection
    End Select
End Function

Private Sub MoveInsertionPoint(ByVal iMoveMode As Integer)
    Dim intPos          As Integer
    Dim intDirection    As Integer
    Dim enuSection      As dtSectionConstants
    
    With txt
        '// Find the current position of insertion point
        intPos = .SelStart
        '// Get the direction to move the insertion point
        intDirection = IIf((iMoveMode = vbKeyRight) Or (iMoveMode = 0), 1, -1)
        '// If insertion point has to be moved, validate & move
        If (iMoveMode <> 0) Then
            If (.SelLength = 0) Then
                .SelStart = .SelStart + intDirection
            Else
                If (iMoveMode = vbKeyRight) Then
                    If ((.SelStart + .SelLength) <> .MaxLength) Then
                        .SelStart = (.SelStart + .SelLength) + intDirection
                    End If
                Else
                    If (.SelStart <> 0) Then
                        .SelStart = .SelStart + intDirection
                    End If
                End If
            End If
        End If
        '// If insertion point is at the end of current section,
        '// move the point to next section
        If (Mid$(mstrText, .SelStart + 1, 1) = mstrSeperator) And (.SelLength = 0) Then
            .SelStart = .SelStart + intDirection
        End If
        '// If current section differs from old section,
        '// select the current section.
        enuSection = GetCurrentSection()
        If (enuSection = dtInvalid) Then Exit Sub
        If (enuSection <> menuCurrentSection) Then
            If (.SelStart <> 0) And (.SelStart <> .MaxLength) Then
                Call MakeSelection(enuSection)
            End If
        End If
    End With
End Sub

Private Sub ChangeSectionValue(ByVal iChangeMode As Integer)
    Dim enuSection      As dtSectionConstants
    Dim intDirection    As Integer
    Dim intStart        As Integer
    Dim intEnd          As Integer
    Dim intValue        As Integer
    
    enuSection = GetCurrentSection()
    If (enuSection = dtInvalid) Then Exit Sub
    
    intDirection = IIf((iChangeMode = vbKeyUp), 1, -1)
    Select Case enuSection
        Case dtDaySection
            Call GetSectionInfo(dtDaySection, intStart, intEnd, intValue)
            intValue = IIf((intValue = 0), Day(Date), intValue + intDirection)
            If (intValue > 31) Then
                intValue = 1
            ElseIf (intValue < 1) Then
                intValue = 31
            End If
            Mid$(mstrText, intStart + 1, (intEnd - intStart)) = VBA.Format$(intValue, "00")
        Case dtMonthSection
            Call GetSectionInfo(dtMonthSection, intStart, intEnd, intValue)
            intValue = IIf((intValue = 0), Month(Date), intValue + intDirection)
            If (intValue > 12) Then
                intValue = 1
            ElseIf (intValue < 1) Then
                intValue = 12
            End If
            Mid$(mstrText, intStart + 1, (intEnd - intStart)) = VBA.Format$(intValue, "00")
        Case dtYearSection
            Call GetSectionInfo(dtYearSection, intStart, intEnd, intValue)
            intValue = IIf((intValue = 0), Year(Date), intValue + intDirection)
            If (intValue > Year(mdtMaxDate)) Then
                intValue = Year(mdtMinDate)
            ElseIf (intValue < Year(mdtMinDate)) Then
                intValue = Year(mdtMaxDate)
            End If
            Mid$(mstrText, intStart + 1, (intEnd - intStart)) = VBA.Format$(intValue, "0000")
    End Select
    
    txt.Text = mstrText
    Call MakeSelection(, intStart, intEnd)
End Sub

Private Sub GetSectionInfo( _
                eSection As dtSectionConstants, _
                Optional iStart As Integer, _
                Optional iEnd As Integer, _
                Optional iValue As Integer)
    Dim strVal          As String
    Dim strCurPosChar   As String
    Dim intPos          As Integer
    
    Select Case eSection
        Case dtDaySection
            iStart = InStr(1, mstrFormatWhenEdit, "d") - 1
            iEnd = iStart + 2
        Case dtMonthSection
            iStart = InStr(1, mstrFormatWhenEdit, "m") - 1
            iEnd = iStart + 2
        Case dtYearSection
            iStart = InStr(1, mstrFormatWhenEdit, "y") - 1
            iEnd = iStart + 4
    End Select
    
    strVal = Mid$(mstrText, iStart + 1, (iEnd - iStart))
    strVal = Replace(strVal, mstrPlaceHolder, "")
    iValue = CInt(Val(Trim(strVal)))
End Sub

Private Sub MoveToNextSection()
    Dim intPos  As Integer
    
    With txt
        intPos = InStr(.SelStart + 1, mstrText, mstrSeperator)
        If (intPos > 0) Then
            .SelStart = intPos + 1
            Call MakeSelection(GetCurrentSection())
        End If
    End With
End Sub

Private Function GetEditFormatString() As String
    Dim strTmp          As String
    
    Select Case menuFormatWhenEdit
        Case [dd/mm/yyyy]: strTmp = "dd" & mstrSeperator & "mm" & mstrSeperator & "yyyy"
        Case [mm/dd/yyyy]: strTmp = "mm" & mstrSeperator & "dd" & mstrSeperator & "yyyy"
        Case [dd/yyyy/mm]: strTmp = "dd" & mstrSeperator & "yyyy" & mstrSeperator & "mm"
        Case [mm/yyyy/dd]: strTmp = "mm" & mstrSeperator & "yyyy" & mstrSeperator & "dd"
        Case [yyyy/dd/mm]: strTmp = "yyyy" & mstrSeperator & "dd" & mstrSeperator & "mm"
        Case [yyyy/mm/dd]: strTmp = "yyyy" & mstrSeperator & "mm" & mstrSeperator & "dd"
    End Select
    GetEditFormatString = strTmp
End Function

Private Function GetDate(dDate As Date) As dtDateValidationConstants
    Dim intDay          As Integer
    Dim intMonth        As Integer
    Dim intYear         As Integer
    Dim intLastDay      As Integer
    Dim intNewMonth     As Integer
    
    Call GetSectionInfo(dtDaySection, , , intDay)
    Call GetSectionInfo(dtMonthSection, , , intMonth)
    Call GetSectionInfo(dtYearSection, , , intYear)
    
    If (intDay = 0) And (intMonth = 0) And (intYear = 0) Then
        GetDate = dtNullDate
    ElseIf (intDay >= 1) And ((intMonth >= 1) And (intMonth <= 12)) Then
        '// If the year is not valid, take current year
        If (intYear = 0) Then intYear = CInt(Year(Date))
        '// Find the last day of entered month
        intNewMonth = CInt(Month(DateSerial(intYear, intMonth, intDay)))
        intLastDay = CInt(Day(DateAdd("d", -1, DateAdd("m", 1, DateSerial(intYear, intMonth, 1)))))
        '// Generate date
        If (intDay <= intLastDay) And (intMonth = intNewMonth) Then
            dDate = DateSerial(intYear, intMonth, intDay)
            If IsInRange(dDate) Then
                GetDate = dtValidDate
            Else
                GetDate = dtInvalidDate
            End If
        Else
            GetDate = dtInvalidDate
        End If
    Else
        GetDate = dtInvalidDate
    End If
End Function

Private Function IsInRange(dDate As Date) As Boolean
    IsInRange = (dDate >= mdtMinDate) And (dDate <= mdtMaxDate)
End Function

Private Function GetFormat() As String
    Dim strFormat As String
    
    If (menuFormat = dtShortDate) Or (menuFormat = dtLongDate) Then
        strFormat = Space$(128)
        Call GetLocaleInfo(GetSystemDefaultLCID(), _
                IIf((menuFormat = dtShortDate), LOCALE_SSHORTDATE, LOCALE_SLONGDATE), _
                strFormat, 128&)
        strFormat = Left$(strFormat, InStr(1, strFormat, vbNullChar) - 1)
    Else
        strFormat = mstrCustomFormat
    End If
    
    GetFormat = strFormat
End Function
