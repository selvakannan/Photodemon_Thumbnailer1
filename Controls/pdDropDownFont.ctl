VERSION 5.00
Begin VB.UserControl pdDropDownFont 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5610
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   374
   ToolboxBitmap   =   "pdDropDownFont.ctx":0000
   Begin PhotoDemon.pdListBoxOD lbPrimary 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   2566
      _ExtentY        =   661
   End
End
Attribute VB_Name = "pdDropDownFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Font-specific Drop Down control 2.0
'Copyright 2016-2024 by Tanner Helland
'Created: 01/June/16
'Last updated: 29/April/20
'Last update: migrate remainder of UI rendering to pd2D
'
'This is a basic dropdown control, with no edit box functionality (by design).  It is very similar in construction to
' the pdListBoxOD object, including its reliance on a separate pdListSupport class for managing its data.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This control raises much fewer events than a standard ListBox, by design
Public Event Click()

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()
Public Event SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, ByRef newTargetHwnd As Long)

'Positioning the dynamically raised listview window is a bit hairy; we use APIs so we can position things correctly
' in the screen's coordinate space (even on high-DPI displays)
Private Declare Function GetWindowRect Lib "user32" (ByVal srcHWnd As Long, ByRef dstRectL As RectL) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal ptrToRect As Long, ByVal bErase As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const WS_EX_WINDOWEDGE As Long = &H100&
Private Const WS_EX_TOPMOST As Long = &H8&
Private Const WS_EX_PALETTEWINDOW As Long = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
Private m_WindowStyleHasBeenSet As Boolean
Private m_OriginalWindowBits As Long, m_OriginalWindowBitsEx As Long
Private m_popupRectCopy As RectL

Private Declare Sub SetWindowPos Lib "user32" (ByVal targetHWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SWP_NOACTIVATE As Long = &H10

'When the popup listbox is raised, we subclass the parent control.  If it is moved or sized or clicked, we automatically
' unload the dropdown listview.  (This workaround is necessary for modal dialogs, among other things.)
Implements ISubclass
Private m_ParentHWnd As Long
Private Const WM_ENTERSIZEMOVE As Long = &H231
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_WINDOWPOSCHANGING As Long = &H46&

'Font size of the dropdown (and corresponding listview).  This controls all rendering metrics, so please don't change
' it at run-time.  Also, note that the optional caption fontsize is a totally different property that can (and should)
' be set independently.
Private m_FontSize As Single

'In a normal dropdown, we can use the same metrics for both the list and the box itself.
' For a font box, however, the list entries must be taller than the dropdown area.
' As such, we let the (more intelligent) list manager handle the list box metrics,
' while we manually control the dropdown box's height.
Private m_IdealControlHeight As Single

'The font drop-down box is unique in that its list area is deliberately wider than the dropdown area (which tends to
' be quite small, on account of the tool options area being snug).  When the font size of the control changes,
' we find the longest font name, and use that as the basis for our list box width.
Private m_LargestWidth As Single

'Padding around the currently selected list item when painted to the combo box.  These values are also added to the
' default font metrics to arrive at a default control size.
Private Const COMBO_PADDING_HORIZONTAL As Single = 4!
Private Const COMBO_PADDING_VERTICAL As Single = 2!

'Change this value to control the maximum number of visible items in the dropped box.  (Note that it's technically
' this value + 1, with the +1 representing the currently selected item.)
Private Const NUM_ITEMS_VISIBLE As Long = 10

'The rectangle where the combo portion of the control is actually rendered
Private m_ComboRect As RectF, m_MouseInComboRect As Boolean

'When the control receives focus via keyboard (e.g. NOT by mouse events), we draw a focus rect to help orient the user.
Private m_FocusRectActive As Boolean

'When the popup listbox is visible, this is set to TRUE.  (Also, as a failsafe the list box hWnd is cached.)
Private m_PopUpVisible As Boolean, m_PopUpHwnd As Long

'Current background color; (background color is used for the 1px border around the button, and it should always match
' our parent control).
Private m_UseCustomBackgroundColor As Boolean, m_BackgroundColor As OLE_COLOR

'List box support class.  Handles data storage and coordinate math for rendering, but for this control, we primarily
' use the data storage aspect.  (Note that when the combo box is clicked and the corresponding listbox window is raised,
' we hand a copy of this class over to the list view so it can clone it and mirror our data.)
Private WithEvents listSupport As pdListSupport
Attribute listSupport.VB_VarHelpID = -1

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'If something forces us to release our subclass while in the midst of the subclass proc, we want to delay the request until
' the subclass exits.  If we don't do this, PD will crash.
Private m_InSubclassNow As Boolean, m_SubclassActive As Boolean
Private WithEvents m_SubclassReleaseTimer As pdTimer
Attribute m_SubclassReleaseTimer.VB_VarHelpID = -1

'String stack that mirrors the current program font cache.
Private m_listOfFonts As pdStringStack

'This UC will be generating an enormous amount of fonts.  We attempt to alleviate this burden by maintaining a persistent collection of the
' past N fonts we've created, on the assumption that we can reuse them at least a few times as the user scrolls the dropdown.
Private m_FontCollection As pdFontCollection

'Preview string to demo each font face.  This is arbitrary, and currently set during the Initialize event.
' (Adding additional scripts is on the TODO list!)
Private m_Text_Default As String
Private m_Text_EN As String
Private m_Text_CJK As String
Private m_Text_Arabic As String
Private m_Text_Hebrew As String

'DrawText functions
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long

'Formatting constants for DrawText
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_CALCRECT = &H400
Private Const DT_NOPREFIX = &H800

'Local list of themable colors.  This list includes all potential colors used by this class, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDDROPDOWNFONT_COLOR_LIST
    [_First] = 0
    PDDD_Background = 0
    PDDD_ComboFill = 1
    PDDD_ComboBorder = 2
    PDDD_DropDownCaption = 3
    PDDD_DropArrow = 4
    PDDD_ListCaption = 5
    PDDD_ListBorder = 6
    [_Last] = 6
    [_Count] = 7
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_DropDownFont
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'Workaround for VB6 quirks; see VBHacks.InControlArray()
Public Function IsChildInControlArray(ByRef ctlChild As Object) As Boolean
    IsChildInControlArray = Not UserControl.Controls(ctlChild.Name) Is ctlChild
End Function

'Initialize the combo box.  This must be called once, by the caller, prior to display.  The combo box will internally cache its
' own copy of the font list, and if for some reason the list changes, this function can be called again to reset the font list.
Public Sub InitializeFontList()
    Me.Clear
    Fonts.GetCopyOfSystemFontList m_listOfFonts
    CopyFontsToListManager
End Sub

'Duplicate a given string inside the API combo box.  We don't actually use this copy of the string (we use our own, so we can support Unicode),
' but this provides a fallback for accessibility technology.
Private Sub CopyFontsToListManager()

    listSupport.SetAutomaticRedraws False
        
    'Iterate through the string stack, adding fonts as we go
    Dim i As Long
    For i = 0 To m_listOfFonts.GetNumOfStrings - 1
        listSupport.AddItem m_listOfFonts.GetString(i), i
    Next i
    
    listSupport.SetAutomaticRedraws True
    
End Sub

'BackgroundColor and BackColor are different properties.  BackgroundColor should always match the color of the parent control,
' while BackColor controls the actual button fill (and can be anything you want).
Public Property Get BackgroundColor() As OLE_COLOR
    BackgroundColor = m_BackgroundColor
End Property

Public Property Let BackgroundColor(ByVal newColor As OLE_COLOR)
    If (m_BackgroundColor <> newColor) Then
        m_BackgroundColor = newColor
        RedrawBackBuffer
    End If
End Property

Public Property Get UseCustomBackgroundColor() As Boolean
    UseCustomBackgroundColor = m_UseCustomBackgroundColor
End Property

Public Property Let UseCustomBackgroundColor(ByVal newSetting As Boolean)
    If (newSetting <> m_UseCustomBackgroundColor) Then
        m_UseCustomBackgroundColor = newSetting
        RedrawBackBuffer
    End If
End Property

'Caption is handled just like the common control label's caption property.  It is valid at design-time, and any translation,
' if present, will not be processed until run-time.
' IMPORTANT NOTE: only the ENGLISH caption is returned.  I don't have a reason for returning a translated caption (if any),
'                  but I can revisit in the future if it ever becomes relevant.
Public Property Get Caption() As String
    Caption = ucSupport.GetCaptionText()
End Property

Public Property Let Caption(ByRef newCaption As String)
    ucSupport.SetCaptionText newCaption
    PropertyChanged "Caption"
End Property

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    RedrawBackBuffer
    PropertyChanged "Enabled"
End Property

'Font settings other than size are not supported.  If you want specialized per-item rendering, use an owner-drawn list box
Public Property Get FontSize() As Single
    FontSize = m_FontSize
End Property

Public Property Let FontSize(ByVal newSize As Single)
    
    m_FontSize = newSize
    
    'A *ton* of rendering metrics are tied to the current font size.  All must be refreshed upon a change.
    m_LargestWidth = 0
    m_IdealControlHeight = Fonts.GetDefaultStringHeight(m_FontSize) + COMBO_PADDING_VERTICAL * 2
    lbPrimary.ListItemHeight = Fonts.GetDefaultStringHeight(m_FontSize) * 2 + 2
    listSupport.DefaultItemHeight = lbPrimary.ListItemHeight
    
    PropertyChanged "FontSize"
    
End Property

'Font settings other than size are not supported.  If you want specialized per-item rendering, use an owner-drawn list box
Public Property Get FontSizeCaption() As Single
    FontSizeCaption = ucSupport.GetCaptionFontSize()
End Property

Public Property Let FontSizeCaption(ByVal newSize As Single)
    ucSupport.SetCaptionFontSize newSize
    PropertyChanged "FontSizeCaption"
End Property

Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'To support high-DPI settings properly, we expose some specialized move+size functions
Public Function GetLeft() As Long
    GetLeft = ucSupport.GetControlLeft
End Function

Public Sub SetLeft(ByVal newLeft As Long)
    ucSupport.RequestNewPosition newLeft, , True
End Sub

Public Function GetTop() As Long
    GetTop = ucSupport.GetControlTop
End Function

Public Sub SetTop(ByVal newTop As Long)
    ucSupport.RequestNewPosition , newTop, True
End Sub

Public Function GetWidth() As Long
    GetWidth = ucSupport.GetControlWidth
End Function

Public Sub SetWidth(ByVal newWidth As Long)
    ucSupport.RequestNewSize newWidth, , True
End Sub

'Use this helper function to automatically set the dropdown control's width, according to the width of its longest text entry.
Public Sub SetWidthAutomatically()

    Dim newWidth As Long, testWidth As Long
    newWidth = 0
    
    If (listSupport.ListCount > 0) Then
    
        Dim i As Long
        For i = 0 To listSupport.ListCount - 1
            testWidth = Fonts.GetDefaultStringWidth(listSupport.List(i, True), m_FontSize)
            If (testWidth > newWidth) Then newWidth = testWidth
        Next i
    
    Else
        newWidth = FixDPI(100)
    End If
    
    'The drop-down arrow's size is fixed, and we also add in the width of the scrollbar (which may be relevant for
    ' some lists)
    newWidth = newWidth + FixDPI(36)
    ucSupport.RequestNewSize newWidth, , True
    
End Sub

Public Function GetHeight() As Long
    GetHeight = ucSupport.GetControlHeight
End Function

Public Sub SetHeight(ByVal newHeight As Long)
    ucSupport.RequestNewSize , newHeight, True
End Sub

Public Sub SetPositionAndSize(ByVal newLeft As Long, ByVal newTop As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ucSupport.RequestFullMove newLeft, newTop, newWidth, newHeight, True
End Sub

'Listbox-specific functions and subs.  Most of these simply relay the request to the listSupport object, and it will
' raise redraw requests as relevant.
Public Sub AddItem(Optional ByVal srcItemText As String = vbNullString, Optional ByVal itemIndex As Long = -1, Optional ByVal hasTrailingSeparator As Boolean = False, Optional ByVal itemHeight As Long = -1)
    listSupport.AddItem srcItemText, itemIndex, hasTrailingSeparator, itemHeight
End Sub

Public Sub Clear()
    listSupport.Clear
    Set m_listOfFonts = New pdStringStack
End Sub

Public Function GetDefaultItemHeight() As Long
    GetDefaultItemHeight = listSupport.DefaultItemHeight
End Function

Public Function List(ByVal itemIndex As Long, Optional ByVal returnTranslatedText As Boolean = False) As String
    List = listSupport.List(itemIndex, returnTranslatedText)
End Function

Public Function ListCount() As Long
    ListCount = listSupport.ListCount
End Function

Public Property Get ListIndex() As Long
    ListIndex = listSupport.ListIndex
End Property

Public Property Let ListIndex(ByVal newIndex As Long)
    listSupport.ListIndex = newIndex
End Property

Public Function ListIndexByString(ByRef srcString As String, Optional ByVal compareMode As VbCompareMethod = vbBinaryCompare) As Long
    ListIndexByString = listSupport.ListIndexByString(srcString, compareMode)
End Function

Public Sub RemoveItem(ByVal itemIndex As Long)
    listSupport.RemoveItem itemIndex
End Sub

Private Sub lbPrimary_Click()
    
    'Mirror any changes to the base dropdown control, then hide the list box
    Me.ListIndex = lbPrimary.ListIndex
    HideListBox
    
    'Restore the focus to the base combo box
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI Me.hWnd
    
End Sub

Private Sub lbPrimary_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, itemTextEn As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToRectF As Long)
    
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'Cache colors in advance, so we can simply reuse them in the inner loop
    Dim itemFillColor As Long, itemFillBorderColor As Long, itemFontColor As Long
    itemFillColor = m_Colors.RetrieveColor(PDDD_ComboFill, Me.Enabled, itemIsSelected, itemIsHovered)
    itemFillBorderColor = m_Colors.RetrieveColor(PDDD_ListBorder, Me.Enabled, itemIsSelected, itemIsHovered)
    itemFontColor = m_Colors.RetrieveColor(PDDD_ListCaption, Me.Enabled, itemIsSelected, itemIsHovered)
    
    'Grab the rendering rect
    Dim tmpRectF As RectF
    CopyMemoryStrict VarPtr(tmpRectF), ptrToRectF, 16&
    
    'Paint the fill and border
    Dim cSurface As pd2DSurface, cBrush As pd2DBrush, cPen As pd2DPen
    Set cSurface = New pd2DSurface: Set cBrush = New pd2DBrush: Set cPen = New pd2DPen
    cSurface.WrapSurfaceAroundDC bufferDC
    cSurface.SetSurfaceAntialiasing P2_AA_None
    cSurface.SetSurfaceCompositing P2_CM_Overwrite
    
    cBrush.SetBrushColor itemFillColor
    PD2D.FillRectangleF_FromRectF cSurface, cBrush, tmpRectF
    
    cPen.SetPenColor itemFillBorderColor
    cPen.SetPenWidth 1!
    PD2D.DrawRectangleF_FromRectF cSurface, cPen, tmpRectF
    
    Set cPen = Nothing: Set cBrush = Nothing: Set cSurface = Nothing
    
    'Paint the font name in the default UI font
    Dim tmpFont As pdFont, textPadding As Single
    Set tmpFont = Fonts.GetMatchingUIFont(m_FontSize)
    textPadding = COMBO_PADDING_HORIZONTAL
    
    Dim tmpString As String
    tmpString = m_listOfFonts.GetString(itemIndex)
    
    tmpFont.SetFontColor itemFontColor
    tmpFont.AttachToDC bufferDC
    tmpFont.SetTextAlignment vbLeftJustify
    tmpFont.FastRenderTextWithClipping tmpRectF.Left + textPadding, tmpRectF.Top + COMBO_PADDING_VERTICAL, tmpRectF.Width - COMBO_PADDING_HORIZONTAL, tmpRectF.Height - COMBO_PADDING_VERTICAL, tmpString, False, True, False
    tmpFont.ReleaseFromDC
    
    'Next, we want to draw a font preview.  Instead of using a pdFont object, we handle this manually, as there are unique layout needs
    ' depending on the associated font.
    
    'Start by creating this font, as necessary
    Dim fontIndex As Long
    fontIndex = m_FontCollection.AddFontToCache(tmpString, m_FontSize + 4!)

    'Retrieve a handle to the created font
    Dim fontHandle As Long
    fontHandle = m_FontCollection.GetFontHandleByPosition(fontIndex)
    
    'Select the font into the target DC
    Dim oldFont As Long
    oldFont = SelectObject(bufferDC, fontHandle)
    
    'Start by retrieving the width of the font name.  We know that this will be less than 1/2 the total width of the rect,
    ' because we created the rect size using the drawn length of the longest font name!
    Dim fontNameWidth As Long
    Dim tmpRectMeasure As RECT
    tmpFont.DrawTextWrapper StrPtr(tmpString), Len(tmpString), tmpRectMeasure, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX Or DT_CALCRECT
    fontNameWidth = (tmpRectMeasure.Right - tmpRectMeasure.Left)
    
    'Generate a destination rect, inside which we will right-align the text.
    Dim previewRect As RECT
    With previewRect
        
        'For the left boundary, we use the larger of...
        ' 1) the length of the font name (as drawn in the UI font), plus a few extra pixels for padding
        ' 2) the halfway point in the drop-down area
        Dim calcLeft As Long, calcLeftAlternate As Long
        calcLeft = tmpRectF.Left + 4 + fontNameWidth + Interface.FixDPI(32)
        calcLeftAlternate = tmpRectF.Left + 4 + ((tmpRectF.Left + tmpRectF.Width) - tmpRectF.Left - 8) \ 2
        
        If (calcLeft > calcLeftAlternate) Then .Left = calcLeftAlternate Else .Left = calcLeft
        
        'Right/top/bottom are all self-explanatory
        .Right = tmpRectF.Left + tmpRectF.Width - COMBO_PADDING_HORIZONTAL
        .Top = tmpRectF.Top
        .Bottom = tmpRectF.Top + tmpRectF.Height
        
    End With
    
    'Create sample text based on the scripts supported by this font.  If no special scripts are supported,
    ' default English text is used.
    '
    'Note that this behavior can be overridden by the "Interface" performance property
    Dim sampleText As String
    If (g_InterfacePerformance <> PD_PERF_FASTEST) Then
    
        Dim tmpProperties As PD_FONT_PROPERTY
        If m_FontCollection.GetFontPropertiesByPosition(fontIndex, tmpProperties) Then
        
            If tmpProperties.Supports_CJK Then
                sampleText = m_Text_Default & " " & m_Text_CJK
            ElseIf tmpProperties.Supports_Arabic Then
                sampleText = m_Text_Default & " " & m_Text_Arabic
            ElseIf tmpProperties.Supports_Hebrew Then
                sampleText = m_Text_Default & " " & m_Text_Hebrew
            ElseIf tmpProperties.Supports_Latin Then
                sampleText = m_Text_Default & " " & m_Text_EN
            Else
                sampleText = m_Text_Default
            End If
            
        Else
            sampleText = m_Text_Default
        End If
        
    Else
        sampleText = m_Text_Default
    End If
    
    'Render right-aligned preview text
    If (LenB(sampleText) <> 0) Then DrawText bufferDC, StrPtr(sampleText), Len(sampleText), previewRect, DT_RIGHT Or DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX
    
    'Release our font
    SelectObject bufferDC, oldFont
                        
End Sub

Private Sub listSupport_Click()
    RaiseEvent Click
End Sub

'When the list manager detects that an action requires the list to be redrawn (like adding a new item), it will raise
' this event.  Whether or not we respond depends on several factors, like whether the user control is currently visible,
' or whether the update actually changed the ListIndex (which is the only thing this front-facing portion of the
' dropdown cares about).
Private Sub listSupport_RedrawNeeded()
    If ucSupport.AmIVisible Then RedrawBackBuffer True
End Sub

'If a subclassis active, this timer will repeatedly try to kill it.  Do not enable it until you are certain the subclass
' needs to be released.  (This is used as a failsafe if we cannot immediately release the subclass when focus is lost.)
Private Sub m_SubclassReleaseTimer_Timer()
    If (Not m_InSubclassNow) Then
        m_SubclassReleaseTimer.StopTimer
        RemoveSubclass
    End If
End Sub

Private Sub ucSupport_GotFocusAPI()
    m_FocusRectActive = True
    RedrawBackBuffer
    RaiseEvent GotFocusAPI
End Sub

Private Sub ucSupport_KeyDownSystem(ByVal Shift As ShiftConstants, ByVal whichSysKey As PD_NavigationKey, markEventHandled As Boolean)
    
    'Enter/Esc get reported directly to the system key handler.  Note that we track the return, because TRUE
    ' means the key was successfully forwarded to the relevant handler.  (If FALSE is returned, no control
    ' accepted the keypress, meaning we should forward the event down the line.)
    markEventHandled = NavKey.NotifyNavKeypress(Me, whichSysKey, Shift)
    
End Sub

Private Sub ucSupport_LostFocusAPI()
    If m_PopUpVisible Then HideListBox
    m_FocusRectActive = False
    RedrawBackBuffer
    RaiseEvent LostFocusAPI
End Sub

Private Sub ucSupport_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    UpdateMousePosition x, y
    If m_MouseInComboRect And (Me.ListCount > 1) Then RaiseListBox
End Sub

Private Sub ucSupport_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    If m_PopUpVisible Then
        lbPrimary.NotifyKeyDown Shift, vkCode, markEventHandled
    Else
        listSupport.NotifyKeyDown Shift, vkCode, markEventHandled
    End If
End Sub

Private Sub ucSupport_KeyUpCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
    listSupport.NotifyKeyUp Shift, vkCode, markEventHandled
End Sub

Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    UpdateMousePosition x, y
    RedrawBackBuffer
End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    UpdateMousePosition x, y
End Sub

Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    UpdateMousePosition -100, -100
End Sub

Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    UpdateMousePosition x, y
End Sub

Private Sub ucSupport_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    RedrawBackBuffer
End Sub

Private Sub UpdateMousePosition(ByVal mouseX As Single, ByVal mouseY As Single)
    
    Dim mouseCheck As Boolean
    mouseCheck = PDMath.IsPointInRectF(mouseX, mouseY, m_ComboRect)
    
    If m_MouseInComboRect <> mouseCheck Then
        m_MouseInComboRect = mouseCheck
        If m_MouseInComboRect Then ucSupport.RequestCursor IDC_HAND Else ucSupport.RequestCursor IDC_DEFAULT
        RedrawBackBuffer
    End If
    
End Sub

'Unlike a regular listview, where the mousewheel results in pixel-level content scrolling, a closed dropdown scrolls actual
' list values one-at-a-time on each wheel motion.
Private Sub ucSupport_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    listSupport.NotifyMouseWheelVertical Button, Shift, x, y, scrollAmount
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then UpdateControlLayout Else RedrawBackBuffer
End Sub

Private Sub ucSupport_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    RaiseEvent SetCustomTabTarget(shiftTabWasPressed, newTargetHwnd)
End Sub

Private Sub ucSupport_VisibilityChange(ByVal newVisibility As Boolean)
    If newVisibility Then
        listSupport.SetAutomaticRedraws True, True
    Else
        If m_PopUpVisible Then HideListBox
    End If
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestCaptionSupport False
    ucSupport.RequestExtraFunctionality True, True
    ucSupport.SpecifyRequiredKeys VK_DOWN, VK_UP, VK_PAGEDOWN, VK_PAGEUP, VK_HOME, VK_END, VK_RETURN, VK_SPACE, VK_ESCAPE
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDDROPDOWNFONT_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDDropDownFont", colorCount
    If (Not PDMain.IsProgramRunning()) Then UpdateColorList
    
    'Initialize a helper list class; it manages the actual list data, and a bunch of rendering and layout decisions
    Set listSupport = New pdListSupport
    listSupport.SetAutomaticRedraws False
    listSupport.ListSupportMode = PDLM_DropDown
    
    'Prep font-specific managers and renderers
    Set m_listOfFonts = New pdStringStack
    
    'Initialize our font collection.  This is used to store a copy of each font face, as it's encountered, which we use to render preview
    ' text on the right side of the font dropdown.
    Set m_FontCollection = New pdFontCollection
    m_FontCollection.SetCacheSize NUM_ITEMS_VISIBLE * 3 + 1
    m_FontCollection.SetExtendedPropertyCaching True
    
    'Create demo strings, to be rendered in the drop-down using the current font face
    m_Text_Default = "AaBbCc 123"
    m_Text_EN = "Sample"
    m_Text_Arabic = ChrW$(&H639) & ChrW$(&H64A) & ChrW$(&H646) & ChrW$(&H629)
    m_Text_Hebrew = ChrW$(&H5D3) & ChrW$(&H5D5) & ChrW$(&H5BC) & ChrW$(&H5D2) & ChrW$(&H5DE) & ChrW$(&H5B8) & ChrW$(&H5D4)
    
    'CJK char previews are handled specially, in UpdateAgainstCurrentTheme - look there for details.
    
End Sub

Private Sub UserControl_InitProperties()
    BackgroundColor = vbWhite
    UseCustomBackgroundColor = False
    Caption = vbNullString
    Enabled = True
    FontSize = 10
    FontSizeCaption = 12
End Sub

'At run-time, painting is handled by the support class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_BackgroundColor = .ReadProperty("BackgroundColor", vbWhite)
        m_UseCustomBackgroundColor = .ReadProperty("UseCustomBackgroundColor", False)
        Caption = .ReadProperty("Caption", vbNullString)
        Enabled = .ReadProperty("Enabled", True)
        FontSize = .ReadProperty("FontSize", 10)
        FontSizeCaption = .ReadProperty("FontSizeCaption", 12)
    End With
End Sub

Private Sub UserControl_Resize()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_Terminate()
    'As a failsafe, immediately release the popup box.  (If we don't do this, PD will crash.)
    If m_PopUpVisible Then HideListBox
    If Not (m_SubclassReleaseTimer Is Nothing) Then m_SubclassReleaseTimer.StopTimer
    SafelyRemoveSubclass
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackgroundColor", m_BackgroundColor, vbWhite
        .WriteProperty "UseCustomBackgroundColor", m_UseCustomBackgroundColor, False
        .WriteProperty "Caption", Me.Caption, vbNullString
        .WriteProperty "Enabled", Me.Enabled, True
        .WriteProperty "FontSize", Me.FontSize, 10
        .WriteProperty "FontSizeCaption", ucSupport.GetCaptionFontSize, 12
    End With
End Sub

Private Sub RaiseListBox()
    
    On Error GoTo UnexpectedListBoxTrouble
    
    If (Not ucSupport.AmIVisible) Or (Not ucSupport.AmIEnabled) Or (Not PDMain.IsProgramRunning()) Then Exit Sub
    
    'We first want to retrieve this control instance's window coordinates *in the screen's coordinate space*.
    ' (We need this to know how to position the listbox element.)
    Dim myRect As RectL
    GetWindowRect Me.hWnd, myRect
    
    'We now want to figure out the idealized coordinates for the pop-up rect.  I prefer an OSX / Windows 10 approach to
    ' positioning, where the currently selected item (.ListIndex) is positioned directly over the underlying combo box,
    ' with neighboring entries positioned above and/or below, as relevant.
    Dim popupRect As RectF, topOfListIndex As Single
    
    'To construct this rect, we start by calculating the position of the .ListIndex item itself
    With popupRect
        If ucSupport.IsCaptionActive Then
            .Left = myRect.Left + FixDPI(8)
            .Top = myRect.Top + (ucSupport.GetCaptionBottom + 2)
        Else
            .Left = myRect.Left
            .Top = myRect.Top
        End If
        .Width = myRect.Right - .Left
        .Height = listSupport.DefaultItemHeight
    End With
    
    topOfListIndex = popupRect.Top
    
    'Specific to the font dropdown, we are now going to calculate a new width (as necessary).  This is required because
    ' the list contains additional font preview data that the closed box does not.
    Dim i As Long
    If (m_LargestWidth = 0) Then
        
        'Create a temporary DIB so we don't have to constantly re-select the font into a DC of its own making.
        Dim tmpDC As Long
        tmpDC = GDI.GetMemoryDC()
        
        'Font names are rendered in the current UI font
        Dim curFont As pdFont
        Set curFont = Fonts.GetMatchingUIFont(m_FontSize)
        curFont.AttachToDC tmpDC
        
        'Find the longest font name
        Dim tmpWidth As Long
        For i = 0 To m_listOfFonts.GetNumOfStrings - 1
            tmpWidth = curFont.GetWidthOfString(m_listOfFonts.GetString(i))
            If (tmpWidth > m_LargestWidth) Then m_LargestWidth = tmpWidth
        Next i
        
        curFont.ReleaseFromDC
        GDI.FreeMemoryDC tmpDC
        
        'The "best" width of the dropdown is a little sketchy, due to the font previews on the right.  At present,
        ' Use the width of the largest font name (which can only be 32 chars), multiplied by 2 (so an equal amount of size is allotted for
        ' the preview), plus a few extra pixels for padding, so a long font name with a long font preview don't "smash" together.
        m_LargestWidth = m_LargestWidth * 2.35
        
    End If
    
    popupRect.Width = m_LargestWidth
    
    'Next, we want to determine how many preceding and trailing entries are in the list.  (We keep a running tally of how
    ' many items theoretically appear in the current list, because we want to make sure that at least a certain amount are
    ' visible in the dropdown, if possible.)  These are purposefully declared as singles, as you'll see in subsequent steps.
    Dim amtPreceding As Single, amtTrailing As Single
    If (Me.ListIndex > 0) Then amtPreceding = Me.ListIndex Else amtPreceding = 0
    
    If Me.ListIndex >= (Me.ListCount - 1) Then
        amtTrailing = 0
    ElseIf Me.ListIndex < 0 Then
        amtTrailing = Me.ListCount - 1
    Else
        amtTrailing = (Me.ListCount - 1) - Me.ListIndex
    End If
    
    'If the *total* possible amount of items is larger than the previously set NUM_ITEMS_VISIBLE constant, reduce the
    ' numbers proportionally.
    Dim amtToReduceList As Long
    If amtPreceding + amtTrailing > NUM_ITEMS_VISIBLE Then
    
        amtToReduceList = (amtPreceding + amtTrailing) - NUM_ITEMS_VISIBLE
        
        'This step may look weird, but conceptually, it's very simple.  We want to repeatedly reduce the size of the
        ' largest group of dropdown items - either the preceding or trailing group - until one of two things happens:
        ' 1) the two groups are equal in size, or
        ' 2) we reach our "amount to reduce list" target
        ' If (1) is reached before (2), we switch to reducing both groups by one element on each iteration
        Do
        
            If amtPreceding > amtTrailing Then
                amtPreceding = amtPreceding - 1
                amtToReduceList = amtToReduceList - 1
            ElseIf amtTrailing > amtPreceding Then
                amtTrailing = amtTrailing - 1
                amtToReduceList = amtToReduceList - 1
            Else
                amtPreceding = amtPreceding - 1
                amtTrailing = amtTrailing - 1
                amtToReduceList = amtToReduceList - 2
            End If
        
        Loop While amtToReduceList > 0
        
        'We now know exactly how many items we can display above and below the current entry, with a maximum of
        ' NUM_ITEMS_VISIBLE if possible.
        
    End If
    
    'Convert the preceding and trailing list item counts into pixel measurements, and add them to our target rect.
    Dim sizeChange As Single
    If (amtPreceding > 0) Then
        sizeChange = amtPreceding * listSupport.DefaultItemHeight
        popupRect.Top = popupRect.Top - sizeChange
        popupRect.Height = popupRect.Height + sizeChange
    End If
    
    If (amtTrailing > 0) Then
        sizeChange = amtTrailing * listSupport.DefaultItemHeight
        popupRect.Height = popupRect.Height + sizeChange
    End If
    
    'We now want to make sure the popup box doesn't lie off-screen.  Check each dimension in turn, and note that changing
    ' the vertical position of the listbox also changes the pixel-based position of the active .ListIndex within the box.
    If (popupRect.Top < g_Displays.GetDesktopTop) Then
        sizeChange = g_Displays.GetDesktopTop - popupRect.Top
        popupRect.Top = g_Displays.GetDesktopTop
        topOfListIndex = topOfListIndex + sizeChange
    Else
        
        Dim estimatedDesktopBottom As Long
        estimatedDesktopBottom = (g_Displays.GetDesktopTop + g_Displays.GetDesktopHeight) - g_Displays.GetTaskbarHeight
        
        If (popupRect.Top + popupRect.Height > estimatedDesktopBottom) Then
            sizeChange = (popupRect.Top + popupRect.Height) - estimatedDesktopBottom
            popupRect.Top = popupRect.Top - sizeChange
            topOfListIndex = topOfListIndex - sizeChange
        End If
        
    End If

    If (popupRect.Left < g_Displays.GetDesktopLeft) Then
        sizeChange = g_Displays.GetDesktopLeft - popupRect.Left
        popupRect.Left = g_Displays.GetDesktopLeft
    ElseIf (popupRect.Left + popupRect.Width > g_Displays.GetDesktopLeft + g_Displays.GetDesktopWidth) Then
        sizeChange = (popupRect.Left + popupRect.Width) - (g_Displays.GetDesktopLeft + g_Displays.GetDesktopWidth)
        popupRect.Left = popupRect.Left - sizeChange
    End If
    
    'We now have an idealized position rect for the list.  Because listbox scrollbars work in pixel increments, we can now
    ' convert the position of the active .ListIndex item from screen coords into relative coords.
    topOfListIndex = topOfListIndex - popupRect.Top
    
    'The list box is now ready to go.  Before displaying it, we want to convert the listbox to a floating toolbox window
    ' and a bare child of the desktop (hWnd = 0).  This allows the listbox to be positioned outside our boundary rect.
    m_PopUpHwnd = lbPrimary.hWnd
    m_ParentHWnd = UserControl.Parent.hWnd
    If (Not m_WindowStyleHasBeenSet) Then
        m_WindowStyleHasBeenSet = True
        m_OriginalWindowBits = g_WindowManager.GetWindowLongWrapper(m_PopUpHwnd)
        m_OriginalWindowBitsEx = g_WindowManager.GetWindowLongWrapper(m_PopUpHwnd, True)
    End If
    
    'Now we are ready to display the window.  Make it a top-level window (SetParent null) and apply any other relevant
    ' window styles.  The top-level window is especially important, as it allows the listbox to be positioned outside
    ' the boundary rect of this control.
    SetParent m_PopUpHwnd, 0&
    g_WindowManager.SetWindowLongWrapper m_PopUpHwnd, WS_EX_PALETTEWINDOW, False, True
    
    'Normally, you need to reset the popup and child flags when you make a window top-level.  Unfortunately, this breaks
    ' the window terribly, and I'm not sure why; it's probably an internal VB thing.  At any rate, the current solution
    ' seems to work, so we ignore this for now.
    'g_WindowManager.SetWindowLongWrapper m_PopUpHwnd, WS_CHILD, True, False
    'g_WindowManager.SetWindowLongWrapper m_PopUpHwnd, WS_POPUP, False, False
    
    'Move the listbox into position *but do not display it*
    With popupRect
        SetWindowPos m_PopUpHwnd, 0&, .Left, .Top, .Width, .Height, SWP_NOACTIVATE
    End With
    
    'We also need to cache the popup rect's position; when the listbox is closed, we will manually invalidate windows
    ' beneath it (only on certain OS + theme combinations; Aero handles this correctly).
    With m_popupRectCopy
        .Left = popupRect.Left
        .Top = popupRect.Top
        .Right = popupRect.Left + popupRect.Width
        .Bottom = popupRect.Top + popupRect.Height
    End With
    
    'Clone our list's contents; note that we cannot do this until *after* the list size has been established, as the
    ' scroll bar's maximum value is contingent on the available pixel size of the dropdown.
    lbPrimary.CloneExternalListSupport listSupport, topOfListIndex, PDLM_LB_Inside_DD
    
    'Now we can show the window
    With popupRect
        SetWindowPos m_PopUpHwnd, 0&, .Left, .Top, .Width, .Height, SWP_SHOWWINDOW
    End With
    
    'One last thing: because this is a (fairly?  mostly?  extremely?) hackish way to emulate a combo box, we need to cover the
    ' case where the user selects outside the raised list box, but *not* on an object that can receive focus (e.g. an exposed
    ' section of an underlying form).  Focusable objects are taken care of automatically, because a LostFocus event will fire,
    ' but non-focusable clicks are problematic.  To solve this, we subclass our parent control and watch for mouse events.
    ' Also, since we're subclassing the control anyway, we'll also hide the ListBox if the parent window is moved.
    If (m_ParentHWnd <> 0) And PDMain.IsProgramRunning() Then
        
        'Make sure we're not currently trying to release a previous subclass attempt
        Dim subclassActive As Boolean: subclassActive = False
        If Not (m_SubclassReleaseTimer Is Nothing) Then
            If m_SubclassReleaseTimer.IsActive Then
                m_SubclassReleaseTimer.StopTimer
                subclassActive = True
            End If
        End If
        
        If (Not subclassActive) And (Not m_SubclassActive) Then
            VBHacks.StartSubclassing m_ParentHWnd, Me
            m_SubclassActive = True
        End If
        
    End If
    
    'As an additional failsafe, we also notify the central UserControl tracker that a list box is active.
    ' If any other PD control receives focus, that tracker will automatically unload our list box as well,
    ' "just in case"
    UserControls.NotifyDropDownChangeState Me.hWnd, m_PopUpHwnd, True
    
    m_PopUpVisible = True
    
    Exit Sub
    
UnexpectedListBoxTrouble:
    PDDebug.LogAction "WARNING!  pdDropDownFont.RaiseListBox failed because of Err # " & Err.Number & ", " & Err.Description
    
End Sub

Private Sub HideListBox()

    If m_PopUpVisible And (m_PopUpHwnd <> 0) Then
        
        'Notify the central UserControl tracker that our list box is now inactive.
        UserControls.NotifyDropDownChangeState Me.hWnd, m_PopUpHwnd, False
        
        m_PopUpVisible = False
        SetParent m_PopUpHwnd, Me.hWnd
        If (m_OriginalWindowBits <> 0) Then g_WindowManager.SetWindowLongWrapper m_PopUpHwnd, m_OriginalWindowBits, , , True
        If (m_OriginalWindowBitsEx <> 0) Then g_WindowManager.SetWindowLongWrapper m_PopUpHwnd, m_OriginalWindowBits, , True, True
        g_WindowManager.SetVisibilityByHWnd m_PopUpHwnd, False
        
        m_PopUpHwnd = 0
        
        'If Aero theming is not active, hiding the list box may cause windows beneath the current one to render incorrectly.
        If (OS.IsVistaOrLater And (Not g_WindowManager.IsDWMCompositionEnabled)) Then
            InvalidateRect 0&, VarPtr(m_popupRectCopy), 0&
        End If
        
        'Note that termination may result in the client site not being available.  If this happens, we simply want
        ' to continue; the subclasser will handle clean-up automatically.
        SafelyRemoveSubclass
        
    End If
    
End Sub

'If a hook exists, uninstall it.  DO NOT CALL THIS FUNCTION if the class is currently inside the hook proc.
Private Sub RemoveSubclass()
    On Error GoTo UnsubclassUnnecessary
    If ((m_ParentHWnd <> 0) And m_SubclassActive) Then
        VBHacks.StopSubclassing m_ParentHWnd, Me
        m_ParentHWnd = 0
        m_SubclassActive = False
    End If
UnsubclassUnnecessary:
End Sub

'Release the edit box's keyboard hook.  In some circumstances, we can't do this immediately, so we set a timer that will
' release the hook as soon as the system allows.
Private Sub SafelyRemoveSubclass()
    If m_InSubclassNow Then
        If (m_SubclassReleaseTimer Is Nothing) Then Set m_SubclassReleaseTimer = New pdTimer
        m_SubclassReleaseTimer.Interval = 16
        m_SubclassReleaseTimer.StartTimer
    Else
        RemoveSubclass
    End If
End Sub

'Whenever a control property changes that affects control size or layout (including internal changes, like caption adjustments),
' call this function to recalculate the control's internal layout
Private Sub UpdateControlLayout()
    
    'This control auto-sizes its height to match the current font.  To make it a different size, adjust the padding
    ' constants at the top of this module.
    Dim desiredControlHeight As Long
    If ucSupport.IsCaptionActive Then desiredControlHeight = ucSupport.GetCaptionBottom + 2 Else desiredControlHeight = 0
    desiredControlHeight = desiredControlHeight + m_IdealControlHeight + COMBO_PADDING_VERTICAL * 2
    
    'Apply the new height to this UC instance, as necessary
    If ucSupport.GetControlHeight <> desiredControlHeight Then
        ucSupport.RequestNewSize , desiredControlHeight, True
        Exit Sub
    End If
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Next, determine the positioning of the caption, if present.  (ucSupport.GetCaptionBottom tells us where the
    ' caption text ends vertically.)
    If ucSupport.IsCaptionActive Then
        
        'The dropdown area is placed relative to the caption
        With m_ComboRect
            .Left = FixDPI(8)
            .Top = ucSupport.GetCaptionBottom + 3
            .Width = (bWidth - 2) - .Left
            .Height = (bHeight - 2) - .Top
        End With
        
    'If there's no caption, allow the clickable portion to fill the entire control
    Else
        
        With m_ComboRect
            .Left = 1
            .Top = 1
            .Width = (bWidth - 2) - .Left
            .Height = (bHeight - 2) - .Top
        End With
        
    End If
    
    'Notify the list manager of our new size.  (Note that this isn't necessary from a rendering standpoint, as we don't
    ' render a normal list-type UI to the dropdown - but the listSupport class won't raise Redraw events if it has an
    ' invalid rendering rect.)
    listSupport.NotifyParentRectF m_ComboRect
    
    'With all size metrics handled, we can now paint the back buffer
    RedrawBackBuffer True
    
End Sub

'Primary rendering function.  Note that ucSupport handles a number of rendering duties (like maintaining a back buffer for us).
Private Sub RedrawBackBuffer(Optional ByVal redrawImmediately As Boolean = False)
    
    'Figure out which background color to use.  This is normally determined by theme, but individual buttons also allow
    ' a custom .BackColor property (important if this instance lies atop a non-standard background, like a command bar).
    Dim finalBackColor As Long
    If m_UseCustomBackgroundColor Then finalBackColor = m_BackgroundColor Else finalBackColor = m_Colors.RetrieveColor(PDDD_Background, Me.Enabled)
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, finalBackColor)
    If (bufferDC = 0) Then Exit Sub
    
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Thanks to the v7.0 theming overhaul, it's completely safe to retrieve colors in the IDE, so we no longer
    ' need to handle these specially.
    Dim ddColorBorder As Long, ddColorFill As Long, ddColorText As Long, ddColorArrow As Long
    ddColorBorder = m_Colors.RetrieveColor(PDDD_ComboBorder, Me.Enabled, False, m_MouseInComboRect Or m_FocusRectActive)
    ddColorFill = m_Colors.RetrieveColor(PDDD_ComboFill, Me.Enabled, False, m_MouseInComboRect Or m_FocusRectActive)
    ddColorText = m_Colors.RetrieveColor(PDDD_DropDownCaption, Me.Enabled, False, m_MouseInComboRect Or m_FocusRectActive)
    ddColorArrow = m_Colors.RetrieveColor(PDDD_DropArrow, Me.Enabled, False, m_MouseInComboRect Or m_FocusRectActive)
    
    If PDMain.IsProgramRunning() Then
        
        'pd2D is used for all UI rendering
        Dim cSurface As pd2DSurface: Set cSurface = New pd2DSurface
        cSurface.WrapSurfaceAroundDC bufferDC
        cSurface.SetSurfaceAntialiasing P2_AA_None
        cSurface.SetSurfacePixelOffset P2_PO_Normal
        
        'First, fill the combo area interior with the established fill color
        Dim cBrush As pd2DBrush: Set cBrush = New pd2DBrush
        cBrush.SetBrushColor ddColorFill
        PD2D.FillRectangleF_FromRectF cSurface, cBrush, m_ComboRect
        
        'A border is always drawn around the control; its size and color vary by hover state, however.
        Dim borderWidth As Single
        If m_MouseInComboRect Or m_FocusRectActive Then borderWidth = 3 Else borderWidth = 1
        
        Dim cPen As pd2DPen: Set cPen = New pd2DPen
        cPen.SetPenWidth borderWidth
        cPen.SetPenColor ddColorBorder
        cPen.SetPenLineJoin P2_LJ_Miter
        cPen.SetPenLineCap P2_LC_Round
        PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_ComboRect
        
        'Next, the right-aligned arrow.  (We need its measurements to know where to restrict the caption's length.)
        Dim ptList(0 To 2) As PointFloat
        ptList(0).x = m_ComboRect.Left + m_ComboRect.Width - Interface.FixDPIFloat(16)
        ptList(0).y = m_ComboRect.Top + (m_ComboRect.Height / 2) - Interface.FixDPIFloat(1)
        
        ptList(2).x = m_ComboRect.Left + m_ComboRect.Width - Interface.FixDPIFloat(8)
        ptList(2).y = ptList(0).y
        
        ptList(1).x = ptList(0).x + (ptList(2).x - ptList(0).x) / 2
        ptList(1).y = ptList(0).y + Interface.FixDPIFloat(3)
        
        cPen.SetPenColor ddColorArrow
        cPen.SetPenWidth 2
        cPen.SetPenLineJoin P2_LJ_Round
        
        cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
        cSurface.SetSurfacePixelOffset P2_PO_Half
        
        PD2D.DrawLinesF_FromPtF cSurface, cPen, 3, VarPtr(ptList(0))
        
        'For an OSX-type look, we can mirror the arrow across the control's center line, then draw it again;
        ' I personally prefer this behavior (as the list box may extend up or down), but I'm not sold on implementing
        ' it just yet, because it's out of place next to regular Windows drop-downs...
        Set cSurface = Nothing
        
        'Finally, paint the caption, and restrict its length to the available dropdown space
        If (Me.ListIndex <> -1) Then
        
            Dim arrowLeftLimit As Single
            arrowLeftLimit = ptList(0).x - Interface.FixDPI(2)
            
            Dim tmpFont As pdFont
            Set tmpFont = Fonts.GetMatchingUIFont(Me.FontSize)
            tmpFont.SetFontColor ddColorText
            tmpFont.SetTextAlignment vbLeftJustify
            tmpFont.AttachToDC bufferDC
            
            With m_ComboRect
                tmpFont.FastRenderTextWithClipping .Left + COMBO_PADDING_HORIZONTAL, .Top + COMBO_PADDING_VERTICAL, arrowLeftLimit, .Height, listSupport.List(Me.ListIndex, False), True, True
            End With
            
            tmpFont.ReleaseFromDC
            Set tmpFont = Nothing
            
        End If
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint redrawImmediately
    
End Sub

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDDD_Background, "Background", IDE_WHITE
        .LoadThemeColor PDDD_ComboFill, "ComboFill", IDE_WHITE
        .LoadThemeColor PDDD_ComboBorder, "ComboBorder", IDE_GRAY
        .LoadThemeColor PDDD_DropDownCaption, "Caption", IDE_GRAY
        .LoadThemeColor PDDD_DropArrow, "DropArrow", IDE_GRAY
        .LoadThemeColor PDDD_ListCaption, "ListCaption", IDE_GRAY
        .LoadThemeColor PDDD_ListBorder, "ListBorder", IDE_GRAY
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    
    If ucSupport.ThemeUpdateRequired Then
                
        'For CJK chars (which are one shared group in OpenType), we try to customize the sample indicator
        ' to the currently active language.  Default to Simplified Chinese if the language engine is unavailable.
        If (g_Language Is Nothing) Then
            m_Text_CJK = ChrW$(&H6837&) & ChrW$(&H672C&)
        Else
            
            'Retrieve the user's current language setting, and use that to guide our sample chars
            Dim curLocale As String
            curLocale = g_Language.GetCurrentLanguage(True)
            
            Select Case LCase$(curLocale)
                
                Case "ja-jp"
                    m_Text_CJK = ChrW$(&H8A66&) & ChrW$(&H6599&)
                    
                Case "ko-kr"
                    m_Text_CJK = ChrW$(&HACAC&) & ChrW$(&HBCF8&)
                    
                Case "zh-tw"
                    m_Text_CJK = ChrW$(&H7BC4&) & ChrW$(&H4F8B&)
                    
                'All others default to simplified Chinese
                Case Else
                    m_Text_CJK = ChrW$(&H6837&) & ChrW$(&H672C&)
                    
            End Select
            
        
        End If
        
        UpdateColorList
        
        listSupport.UpdateAgainstCurrentTheme
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
        lbPrimary.UpdateAgainstCurrentTheme
        
    End If
    
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub

Private Function ISubclass_WindowMsg(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
    
    m_InSubclassNow = True
    
    'If certain events occur in our parent window, and our list box is visible, release it
    If m_PopUpVisible Then
        If (uiMsg = WM_ENTERSIZEMOVE) Or (uiMsg = WM_WINDOWPOSCHANGING) Then
            HideListBox
        ElseIf (uiMsg = WM_LBUTTONDOWN) Or (uiMsg = WM_RBUTTONDOWN) Or (uiMsg = WM_MBUTTONDOWN) Then
            HideListBox
        ElseIf (uiMsg = WM_NCDESTROY) Then
            HideListBox
            Set m_SubclassReleaseTimer = Nothing
            VBHacks.StopSubclassing hWnd, Me
            m_ParentHWnd = 0
        End If
    End If
    
    'Never eat parent window messages; just peek at them
    ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    
    m_InSubclassNow = False
    
End Function
