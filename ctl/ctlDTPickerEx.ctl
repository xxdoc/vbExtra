VERSION 5.00
Begin VB.UserControl DTPickerEx 
   ClientHeight    =   432
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3768
   HasDC           =   0   'False
   PropertyPages   =   "ctlDTPickerEx.ctx":0000
   ScaleHeight     =   432
   ScaleWidth      =   3768
   ToolboxBitmap   =   "ctlDTPickerEx.ctx":0045
End
Attribute VB_Name = "DTPickerEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' -------------------------------------------------------------------------
' Autor:    Leandro I. Ascierto
' Web:      www.leandroascierto.com.ar
' Fecha:    Domingo, 08 de Noviembre de 2009
' History:
'           26/10/2010
'               Call DoEvents in NM_SETFOCUS (Textbox control cause hide MONTHCAL)
'               Put Focus in SysDateTimePick32 on UserControl_GotFocus
'           28/02/2011
'               Implement Text Back Color
' -------------------------------------------------------------------------

' The above are annotations from the original author
' But the control has been modified since then.


Implements ISubclass

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

Private Type NMHDR
    hwndFrom                    As Long
    idfrom                      As Long
    code                        As Long
End Type

Private Type SYSTEMTIME
    wYear                       As Integer
    wMonth                      As Integer
    wDayOfWeek                  As Integer
    wDay                        As Integer
    wHour                       As Integer
    wMinute                     As Integer
    wSecond                     As Integer
    wMilliseconds               As Integer
End Type

Private Type NMDATETIMECHANGE
    NMHDR                       As NMHDR
    Flags                       As Long
    ST                          As SYSTEMTIME
End Type

'Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetDoubleClickTime Lib "user32" () As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function EnableWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Private Const DTM_FIRST                     As Long = &H1000
Private Const DTM_GETSYSTEMTIME             As Long = (DTM_FIRST + 1)
Private Const DTM_SETSYSTEMTIME             As Long = (DTM_FIRST + 2)
Private Const DTM_SETRANGE                  As Long = (DTM_FIRST + 4)
Private Const DTM_SETFORMATA                As Long = (DTM_FIRST + 5)
Private Const DTM_SETMCCOLOR                As Long = (DTM_FIRST + 6)
Private Const DTM_GETIDEALSIZE              As Long = (DTM_FIRST + 15)
Private Const DTM_CLOSEMONTHCAL As Long = (DTM_FIRST + 13)
Private Const DTM_GETMONTHCAL As Long = (DTM_FIRST + 8)

Private Const DTS_SHORTDATEFORMAT           As Long = &H0
Private Const DTS_UPDOWN                    As Long = &H1
Private Const DTS_SHOWNONE                  As Long = &H2
Private Const DTS_LONGDATEFORMAT            As Long = &H4
Private Const DTS_TIMEFORMAT                As Long = &H9

Private Const DTN_FIRST                     As Long = (-760)
Private Const DTN_DATETIMECHANGE            As Long = (DTN_FIRST + 1)
Private Const DTN_DROPDOWN As Long = (DTN_FIRST + 6)
Private Const DTN_CLOSEUP As Long = (DTN_FIRST + 7)

Private Const NM_FIRST As Long = 0
Private Const NM_SETFOCUS As Long = (NM_FIRST - 7)

Private Const GDT_NONE                      As Long = 1
Private Const GDT_VALID                     As Long = 0

Private Const GDTR_MAX                      As Long = &H2
Private Const GDTR_MIN                      As Long = &H1

Private Const MK_SHIFT As Long = &H4
Private Const MK_CONTROL As Long = &H8
Private Const MK_LBUTTON As Long = &H1
Private Const MK_RBUTTON As Long = &H2
Private Const MK_MBUTTON As Long = &H10

Private Enum MCSC
    MCSC_TEXT = 1
    MCSC_TITLEBK = 2
    MCSC_TITLETEXT = 3
    MCSC_MONTHBK = 4
    MCSC_TRAILINGTEXT = 5
End Enum

Public Enum vbExDTPickerFormatConstants
    dtpLongDate = 0
    dtpShortDate = 1
    dtpTime = 2
    dtpCustom = 3
End Enum

Public Enum vbExCC2MousePointerConstants
    cc2Default = 0  'Default
    cc2Arrow = 1  'Arrow mouse pointer
    cc2Cross = 2  'Cross mouse pointer
    cc2IBeam = 3  'I-Beam mouse pointer
    cc2Icon = 4  'Icon mouse pointer
    cc2Size = 5  'Size mouse pointer
    cc2SizeNESW = 6  'Size NE SW mouse pointer
    cc2SizeNS = 7  'Size N S mouse pointer
    cc2SizeNWSE = 8  'Size NW SE mouse pointer
    cc2SizeEW = 9  'Size W E mouse pointer
    cc2UpArrow = 10  'Up arrow mouse pointer
    cc2Hourglass = 11  'Hourglass mouse pointer
    cc2NoDrop = 12  'No drop mouse pointer
    cc2ArrowHourglass = 13  'Arrow and Hourglass mouse pointer
    cc2ArrowQuestion = 14  'Arrow and Question mark mouse pointer
    cc2SizeAll = 15  'Size all mouse pointer
    cc2Custom = 99  'Custom mouse pointer icon specified by the MouseIcon property
End Enum

Private Const WS_CHILD                      As Long = &H40000000
Private Const WS_OVERLAPPED                 As Long = &H0&
Private Const WS_VISIBLE                    As Long = &H10000000
Private Const WS_EX_CLIENTEDGE              As Long = &H200&
Private Const WS_EX_LEFT                    As Long = &H0&
Private Const WS_EX_LTRREADING              As Long = &H0&
Private Const WS_EX_RIGHTSCROLLBAR          As Long = &H0&
Private Const WS_DISABLED                   As Long = &H8000000
Private Const WM_ERASEBKGND                 As Long = &H14
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MOUSEMOVE As Long = &H200

Private Const WM_SYSKEYDOWN As Long = &H104

Private Const WM_EVENTDATECHANGE As Long = (WM_USER + 1000)

Private Const WM_GETTEXT                    As Long = &HD
Private Const WM_GETFONT                    As Long = &H31
Private Const WM_SETFONT                    As Long = &H30
Private Const WM_NOTIFY                     As Long = &H4E
Private Const WM_KEYDOWN                    As Long = &H100
Private Const WM_KEYUP                      As Long = &H101
Private Const WM_CHAR                       As Long = &H102


Private Const cDTPickerMinDate                    As Date = "01/01/1601"
Private Const cDTPickerMaxDate                    As Date = "31/12/9999"

Public Event Change()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event DropDown()
Public Event CloseUp()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event Click()
Public Event DblClick()
Public Event CalendarClick()

Private mDTPickerHwnd                       As Long
Private tSYSTIME                    As SYSTEMTIME

Private mMinDate                    As Date
Private mMaxDate                    As Date
Private mValue                   As Variant
Private mUpDown                     As Boolean
Private mCheckBox                   As Boolean
Private mEnabled As Boolean

Private mTextBackColor              As Long
Private mCalendarBackColor                  As Long
Private mCalendarForeColor                  As Long
Private mCalendarTitleBackColor             As Long
Private mCalendarTitleForeColor             As Long
Private mCalendarTrailingForeColor          As Long

Private mCustomFormat               As String
Private mFormat                     As vbExDTPickerFormatConstants
Private mDroppedDown As Boolean

Private mMousePointer As Long
Private mMouseIcon As StdPicture

Private hBrush                      As Long

Private mSubclassedUsc As Boolean
Private mSubclassedCtl As Boolean
Private mUserControlHwnd As Long
Private mAmbientUserMode As Boolean
Private mMouseIsDown As Boolean
Private mCalendarMouseDown As Boolean
Private mLastClickTime
Private mHwndCalendar As Long
Private mCalendarSubclassed As Long

Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1


Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    Select Case iMsg
        Case WM_NOTIFY, WM_EVENTDATECHANGE
            ISubclass_MsgResponse = emrPostProcess
        Case Else
            ISubclass_MsgResponse = emrPreprocess
    End Select
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, wParam As Long, lParam As Long, bConsume As Boolean) As Long
    Static sInside As Boolean
    
    If hWnd = mUserControlHwnd Then
        If iMsg = WM_NOTIFY Then
            Dim NM As NMDATETIMECHANGE
            CopyMemory NM, ByVal lParam, LenB(NM)

            If NM.NMHDR.code = DTN_DATETIMECHANGE Then
                If lParam <> 1308108 And lParam <> 1307924 Then
                    If NM.Flags = GDT_VALID Then
                        With NM.ST
                            mValue = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
                        End With
                    Else
                        If mCheckBox Then
                            mValue = Null
                        Else
                            mValue = ""
                        End If
                    End If
                    PostMessage mUserControlHwnd, WM_EVENTDATECHANGE, 0&, 0&
                End If
            ElseIf NM.NMHDR.code = DTN_DROPDOWN Then
'                If Not sInside Then
'                    sInside = True
'                    DoEvents
'                    sInside = False
'                End If
                mDroppedDown = True
                mHwndCalendar = SendMessage(mDTPickerHwnd, DTM_GETMONTHCAL, 0, ByVal 0&)
                If mHwndCalendar <> 0 Then
                    SubclassCalendar
                End If
                RaiseEvent DropDown
            ElseIf NM.NMHDR.code = DTN_CLOSEUP Then
                mDroppedDown = False
                mHwndCalendar = SendMessage(mDTPickerHwnd, DTM_GETMONTHCAL, 0, ByVal 0&)
                If mHwndCalendar <> 0 Then
                    UnsubclassCalendar
                End If
                RaiseEvent CloseUp
            ElseIf NM.NMHDR.code = NM_SETFOCUS Then
                If Not sInside Then
                    sInside = True
                    DoEvents
                    sInside = False
                End If
            End If
        ElseIf iMsg = WM_EVENTDATECHANGE Then
            RaiseEvent_Change
        End If
    Else
        Dim x As Single
        Dim y As Single
        
        Select Case iMsg
            Case WM_ERASEBKGND
                Dim Rec As RECT
                Call GetClientRect(hWnd, Rec)
                Call FillRect(wParam, Rec, hBrush)
                'bHandled = True
                bConsume = True
            Case WM_KEYDOWN
                RaiseEvent KeyDown(wParam And &H7FFF&, pvShiftState())
            Case WM_CHAR
                RaiseEvent KeyPress(wParam And &H7FFF&)
                If InStr("0123456789", ChrW(wParam)) = 0 Then
                    bConsume = True
                End If
                Case WM_KEYUP
                RaiseEvent KeyUp(wParam And &H7FFF&, pvShiftState())
            Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
                If hWnd = mDTPickerHwnd Then
                    x = UserControl.ScaleX(GetLowWord(lParam), vbPixels, vbTwips)
                    y = UserControl.ScaleY(GetHighWord(lParam), vbPixels, vbTwips)
                    RaiseEvent MouseDown(IIf(iMsg = WM_LBUTTONDOWN, vbLeftButton, IIf(iMsg = WM_MBUTTONDOWN, vbMiddleButton, vbRightButton)), GetShiftFromwParam(wParam), x, y)
                    mMouseIsDown = True '(iMsg = WM_LBUTTONDOWN)
                Else ' calendar
                    mCalendarMouseDown = True
                End If
            Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
                If hWnd = mDTPickerHwnd Then
                    x = UserControl.ScaleX(GetLowWord(lParam), vbPixels, vbTwips)
                    y = UserControl.ScaleY(GetHighWord(lParam), vbPixels, vbTwips)
                    RaiseEvent MouseUp(IIf(iMsg = WM_LBUTTONUP, vbLeftButton, IIf(iMsg = WM_MBUTTONUP, vbMiddleButton, vbRightButton)), GetShiftFromwParam(wParam), x, y)
                    If mMouseIsDown Then
    '                    If iMsg = WM_LBUTTONUP Then
                            If (x >= 0) Then
                                If (x <= UserControl.ScaleWidth) Then
                                    If (y >= 0) Then
                                        If (y <= UserControl.ScaleHeight) Then
                                            If (mLastClickTime > 0) And (((Timer - mLastClickTime) * 1000) <= GetDoubleClickTime) Then
                                                RaiseEvent DblClick
                                            Else
                                                mLastClickTime = Timer
                                                RaiseEvent Click
                                            End If
                                        End If
                                    End If
                                End If
                            End If
    '                    End If
                    End If
                    mMouseIsDown = False
                Else    ' calendar
                    If mCalendarMouseDown Then
                        If mHwndCalendar <> 0 Then
                            Dim iRect As RECT
                            Dim iPt As POINTAPI
                            
                            iPt.x = GetLowWord(lParam)
                            iPt.y = GetHighWord(lParam)
                            
                            GetWindowRect mHwndCalendar, iRect
                            iRect.Right = iRect.Right - iRect.Left
                            iRect.Bottom = iRect.Bottom - iRect.Top
                            iRect.Left = 0
                            iRect.Top = 0
                            If iPt.x >= iRect.Left Then
                                If iPt.x <= iRect.Right Then
                                    If iPt.y >= iRect.Top Then
                                        If iPt.y <= iRect.Bottom Then
                                            RaiseEvent CalendarClick
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    mCalendarMouseDown = False
                End If
            Case WM_MOUSEMOVE
                x = UserControl.ScaleX(GetLowWord(lParam), vbPixels, vbTwips)
                y = UserControl.ScaleY(GetHighWord(lParam), vbPixels, vbTwips)
                RaiseEvent MouseMove(GetMouseBurttonFromwParam(wParam), GetShiftFromwParam(wParam), x, y)
        End Select
    End If
End Function

Private Sub mFont_FontChanged(ByVal PropertyName As String)
    UpdateFont
End Sub

Private Sub UserControl_EnterFocus()
    SetFocusAPI mDTPickerHwnd
End Sub

Private Sub UserControl_GotFocus()
    SetFocusAPI mDTPickerHwnd
End Sub

'----------------------------------------------------------------------------------------------------------------
'UserControl Envents
'----------------------------------------------------------------------------------------------------------------

Private Sub UserControl_Initialize()
    InitCommonControls
End Sub

Private Sub UserControl_Terminate()
    UnsubclassCalendar
    If mSubclassedUsc Then
        DetachMessage Me, mUserControlHwnd, WM_NOTIFY
        DetachMessage Me, mUserControlHwnd, WM_EVENTDATECHANGE
        mSubclassedUsc = False
    End If
    If mSubclassedCtl Then
        DetachMessage Me, mDTPickerHwnd, WM_ERASEBKGND
        DetachMessage Me, mDTPickerHwnd, WM_KEYDOWN
        DetachMessage Me, mDTPickerHwnd, WM_CHAR
        DetachMessage Me, mDTPickerHwnd, WM_KEYUP
        DetachMessage Me, mDTPickerHwnd, WM_LBUTTONDOWN
        DetachMessage Me, mDTPickerHwnd, WM_LBUTTONUP
        DetachMessage Me, mDTPickerHwnd, WM_MBUTTONDOWN
        DetachMessage Me, mDTPickerHwnd, WM_MBUTTONUP
        DetachMessage Me, mDTPickerHwnd, WM_RBUTTONDOWN
        DetachMessage Me, mDTPickerHwnd, WM_RBUTTONUP
        DetachMessage Me, mDTPickerHwnd, WM_MOUSEMOVE
        mSubclassedCtl = False
    End If
    If mDTPickerHwnd Then DestroyWindow mDTPickerHwnd
    If hBrush Then DeleteObject hBrush
End Sub

Private Sub UserControl_InitProperties()
    Enabled = True
    mMinDate = cDTPickerMinDate
    mMaxDate = cDTPickerMaxDate
    Me.TextBackColor = vbWindowBackground
    mCalendarBackColor = vbWindowBackground
    mCalendarForeColor = vbButtonText
    mCalendarTitleBackColor = vbActiveTitleBar
    mCalendarTitleForeColor = vbActiveTitleBarText
    mCalendarTrailingForeColor = vbGrayText
    mFormat = dtpShortDate
    Set UserControl.Font = Ambient.Font
    If Ambient.UserMode Then Set mFont = UserControl.Font
    UserControl.Size 1200, 300

    mAmbientUserMode = Ambient.UserMode
'    If mAmbientUserMode Then
'        On Error Resume Next
'        If GetWindowClassName(UserControl.Parent.hWnd) = "ThunderUserControlDC" Then
'            If Err.Number = 0 Then
'                mAmbientUserMode = False
'            End If
'        End If
'        On Error GoTo 0
'    End If
    
    pvCreate
    
    If mAmbientUserMode Then
        mUserControlHwnd = UserControl.hWnd
        AttachMessage Me, mUserControlHwnd, WM_NOTIFY
        AttachMessage Me, mUserControlHwnd, WM_EVENTDATECHANGE
        mSubclassedUsc = True
    End If
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width < 800 Then UserControl.Width = 800
    SetWindowPos mDTPickerHwnd, 0, 0, 0, UserControl.ScaleWidth / Screen.TwipsPerPixelX, UserControl.ScaleHeight / Screen.TwipsPerPixelY, 0&
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim iVar As Variant
    
    With PropBag
        mMinDate = .ReadProperty("MinDate", cDTPickerMinDate)
        mMaxDate = .ReadProperty("MaxDate", cDTPickerMaxDate)
        mUpDown = .ReadProperty("UpDown", False)
        mCheckBox = .ReadProperty("CheckBox", False)
        mEnabled = .ReadProperty("Enabled", True)
        UserControl.Enabled = mEnabled
        
        Me.TextBackColor = .ReadProperty("TextBackColor", vbWindowBackground)
        mCalendarBackColor = .ReadProperty("CalendarBackColor", vbWindowBackground)
        mCalendarForeColor = .ReadProperty("CalendarForeColor", vbButtonText)
        mCalendarTitleBackColor = .ReadProperty("CalendarTitleBackColor", vbActiveTitleBar)
        mCalendarTitleForeColor = .ReadProperty("CalendarTitleForeColor", vbActiveTitleBarText)
        mCalendarTrailingForeColor = .ReadProperty("CalendarTrailingForeColor", vbGrayText)
        
        iVar = .ReadProperty("Value", vbNullString)
        If IsDate(iVar) Then
            mValue = CDate(iVar)
        Else
            If mCheckBox Then
                mValue = Null
            Else
                mValue = iVar
            End If
        End If
        mCustomFormat = .ReadProperty("CustomFormat", vbNullString)
        mFormat = .ReadProperty("Format", dtpShortDate)
        On Error Resume Next
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        If Err.Number Then
            Set UserControl.Font = Ambient.Font
        End If
        On Error GoTo 0
        If Ambient.UserMode Then Set mFont = UserControl.Font
        MousePointer = .ReadProperty("MousePointer ", cc2Default)
        Set MouseIcon = .ReadProperty("MouseIcon ", Nothing)
    End With

    mAmbientUserMode = Ambient.UserMode
'    If mAmbientUserMode Then
'        On Error Resume Next
'        If GetWindowClassName(UserControl.Parent.hWnd) = "ThunderUserControlDC" Then
'            If Err.Number = 0 Then
'                mAmbientUserMode = False
'            End If
'        End If
'        On Error GoTo 0
'    End If
    
    Call pvCreate
    
    If mAmbientUserMode Then
        mUserControlHwnd = UserControl.hWnd
        AttachMessage Me, mUserControlHwnd, WM_NOTIFY
        AttachMessage Me, mUserControlHwnd, WM_EVENTDATECHANGE
        mSubclassedUsc = True
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "MinDate", mMinDate, cDTPickerMinDate
        .WriteProperty "MaxDate", mMaxDate, cDTPickerMaxDate
        .WriteProperty "UpDown", mUpDown, False
        .WriteProperty "CheckBox", mCheckBox, False
        .WriteProperty "Enabled", mEnabled, True
        
        .WriteProperty "TextBackColor", mTextBackColor, vbWindowBackground
        .WriteProperty "CalendarBackColor", mCalendarBackColor, vbWindowBackground
        .WriteProperty "CalendarForeColor", mCalendarForeColor, vbButtonText
        .WriteProperty "CalendarTitleBackColor", mCalendarTitleBackColor, vbActiveTitleBar
        .WriteProperty "CalendarTitleForeColor", mCalendarTitleForeColor, vbActiveTitleBarText
        .WriteProperty "CalendarTrailingForeColor", mCalendarTrailingForeColor, vbGrayText
        
        .WriteProperty "Value", mValue, vbNullString
        .WriteProperty "CustomFormat", mCustomFormat, vbNullString
        .WriteProperty "Format", mFormat, dtpShortDate
        .WriteProperty "Font", UserControl.Font, Ambient.Font
        .WriteProperty "MousePointer", mMousePointer, cc2Default
        .WriteProperty "MouseIcon", mMouseIcon, Nothing
    End With
End Sub

'----------------------------------------------------------------------------------------------------------------
'Private Functions and Sub
'----------------------------------------------------------------------------------------------------------------

Private Function pvCreate()

    Dim lStyle As Long
    Dim lExStyle As Long

    lStyle = WS_CHILD Or WS_OVERLAPPED Or WS_VISIBLE
    If mCheckBox Then lStyle = lStyle Or DTS_SHOWNONE
    If mUpDown Then lStyle = lStyle Or DTS_UPDOWN
    If Not mEnabled Then lStyle = lStyle Or WS_DISABLED
    
    Select Case mFormat
        Case dtpLongDate
            lStyle = lStyle Or DTS_LONGDATEFORMAT
        Case dtpShortDate
            lStyle = lStyle Or DTS_SHORTDATEFORMAT
        Case dtpTime
            lStyle = lStyle Or DTS_TIMEFORMAT
        Case Else
            If mCustomFormat <> "" Then
                If (InStr(mCustomFormat, "d") = 0) And (InStr(mCustomFormat, "M") = 0) And (InStr(mCustomFormat, "y") = 0) Then
                    If (InStr(mCustomFormat, "h") > 0) Or (InStr(mCustomFormat, "m") > 0) Or (InStr(mCustomFormat, "n") > 0) Or (InStr(mCustomFormat, "s") > 0) Then
                        lStyle = lStyle Or DTS_TIMEFORMAT
                    End If
                End If
            End If
    End Select

    lExStyle = WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR
    
    If mAmbientUserMode Then
        If mSubclassedCtl Then
            DetachMessage Me, mDTPickerHwnd, WM_ERASEBKGND
            DetachMessage Me, mDTPickerHwnd, WM_KEYDOWN
            DetachMessage Me, mDTPickerHwnd, WM_CHAR
            DetachMessage Me, mDTPickerHwnd, WM_KEYUP
            DetachMessage Me, mDTPickerHwnd, WM_LBUTTONDOWN
            DetachMessage Me, mDTPickerHwnd, WM_LBUTTONUP
            DetachMessage Me, mDTPickerHwnd, WM_MBUTTONDOWN
            DetachMessage Me, mDTPickerHwnd, WM_MBUTTONUP
            DetachMessage Me, mDTPickerHwnd, WM_RBUTTONDOWN
            DetachMessage Me, mDTPickerHwnd, WM_RBUTTONUP
            DetachMessage Me, mDTPickerHwnd, WM_MOUSEMOVE
            mSubclassedCtl = False
        End If
    End If
    If mDTPickerHwnd <> 0 Then
        DestroyWindow mDTPickerHwnd
    End If
    
    mDTPickerHwnd = CreateWindowEx(lExStyle, "SysDateTimePick32", "", lStyle, 0, 0, UserControl.ScaleWidth / Screen.TwipsPerPixelX, UserControl.ScaleHeight / Screen.TwipsPerPixelY, UserControl.hWnd, 0&, App.hInstance, ByVal 0&)
    
    If mAmbientUserMode Then
        If mDTPickerHwnd <> 0 Then
            AttachMessage Me, mDTPickerHwnd, WM_ERASEBKGND
            AttachMessage Me, mDTPickerHwnd, WM_KEYDOWN
            AttachMessage Me, mDTPickerHwnd, WM_CHAR
            AttachMessage Me, mDTPickerHwnd, WM_KEYUP
            AttachMessage Me, mDTPickerHwnd, WM_LBUTTONDOWN
            AttachMessage Me, mDTPickerHwnd, WM_LBUTTONUP
            AttachMessage Me, mDTPickerHwnd, WM_MBUTTONDOWN
            AttachMessage Me, mDTPickerHwnd, WM_MBUTTONUP
            AttachMessage Me, mDTPickerHwnd, WM_RBUTTONDOWN
            AttachMessage Me, mDTPickerHwnd, WM_RBUTTONUP
            AttachMessage Me, mDTPickerHwnd, WM_MOUSEMOVE
            mSubclassedCtl = True
        End If
    End If
    
    If mCalendarBackColor <> vbWindowBackground Then pvChangeColor MCSC_MONTHBK, mCalendarBackColor
    If mCalendarForeColor <> vbButtonText Then pvChangeColor MCSC_TEXT, mCalendarForeColor
    If mCalendarTitleBackColor <> vbActiveTitleBar Then pvChangeColor MCSC_TITLEBK, mCalendarTitleBackColor
    If mCalendarTitleForeColor <> vbActiveTitleBarText Then pvChangeColor MCSC_TITLETEXT, mCalendarTitleForeColor
    If mCalendarTrailingForeColor <> vbGrayText Then pvChangeColor MCSC_TRAILINGTEXT, mCalendarTrailingForeColor

    If mCustomFormat <> vbNullString And mFormat = dtpCustom Or mFormat = dtpTime Then
        SendMessage mDTPickerHwnd, DTM_SETFORMATA, 0&, ByVal mCustomFormat
    End If
    
    Call pvSetRange
    
    Me.Value = mValue

End Function

Private Function IsValidDate(vDate As Variant) As Boolean
    Dim iDate As Date
    
    On Local Error GoTo ErrFalse
    If mFormat = dtpTime Then
        IsValidDate = IsDate(vDate)
        If Not IsValidDate Then
            If vDate = 0 Then
                IsValidDate = True
            End If
        End If
    ElseIf IsDate(vDate) Then
        If (DateValue(vDate) >= mMinDate) And (DateValue(vDate) <= mMaxDate) Then
            IsValidDate = True
        End If
    Else
        iDate = CDate(vDate)
        If IsDate(iDate) Then
            IsValidDate = True
        End If
    End If
    Exit Function
ErrFalse:
'Debug.Print Err.Description
End Function

Private Function pvGetSysTime() As Long
    pvGetSysTime = SendMessage(mDTPickerHwnd, DTM_GETSYSTEMTIME, 0&, tSYSTIME)
End Function

Private Sub pvSetDateTime(vDate As Variant)
    With tSYSTIME
        .wDay = VBA.DateTime.Day(vDate)
        .wMonth = VBA.DateTime.Month(vDate)
        .wYear = VBA.DateTime.Year(vDate)
        .wHour = VBA.DateTime.Hour(vDate)
        .wMinute = VBA.DateTime.Minute(vDate)
        .wSecond = VBA.DateTime.Second(vDate)
    End With
End Sub

Private Sub pvChangeColor(wParam As MCSC, oColor As Long)
    Dim lColor As Long
    OleTranslateColor oColor, 0, lColor
    Call SendMessage(mDTPickerHwnd, DTM_SETMCCOLOR, wParam, ByVal lColor)
End Sub

Private Sub pvSetRange()
    Dim tST(1) As SYSTEMTIME
    
    If mFormat <> dtpTime Then
        tST(0).wDay = VBA.Day(mMinDate)
        tST(0).wMonth = VBA.DateTime.Month(mMinDate)
        tST(0).wYear = VBA.DateTime.Year(mMinDate)
        tST(0).wHour = VBA.DateTime.Hour(mMinDate)
        tST(0).wMinute = VBA.DateTime.Minute(mMinDate)
        tST(0).wSecond = VBA.DateTime.Second(mMinDate)
    
        tST(1).wDay = VBA.DateTime.Day(mMaxDate)
        tST(1).wMonth = VBA.DateTime.Month(mMaxDate)
        tST(1).wYear = VBA.DateTime.Year(mMaxDate)
        tST(1).wHour = VBA.DateTime.Hour(mMaxDate)
        tST(1).wMinute = VBA.DateTime.Minute(mMaxDate)
        tST(1).wSecond = VBA.DateTime.Second(mMaxDate)
    End If
    Call SendMessage(mDTPickerHwnd, DTM_SETRANGE, GDTR_MIN + GDTR_MAX, tST(0))
End Sub

'----------------------------------------------------------------------------------------------------------------
'Public Function
'----------------------------------------------------------------------------------------------------------------

Public Function GetCaption() As String
    Dim lLength As Long
    Dim sText As String

    sText = Space(255)
    lLength = SendMessage(mDTPickerHwnd, WM_GETTEXT, 255, ByVal sText)
    If lLength Then GetCaption = Left(sText, lLength)

End Function

Public Function Day() As Integer
    Day = IIf(pvGetSysTime = GDT_VALID, tSYSTIME.wDay, -1)
End Function

Public Function Month() As Integer
    Month = IIf(pvGetSysTime = GDT_VALID, tSYSTIME.wMonth, -1)
End Function

Public Function Year() As Integer
    Year = IIf(pvGetSysTime = GDT_VALID, tSYSTIME.wYear, -1)
End Function

Public Function Hour() As Integer
    Hour = IIf(pvGetSysTime = GDT_VALID, tSYSTIME.wHour, -1)
End Function

Public Function Minute() As Integer
    Minute = IIf(pvGetSysTime = GDT_VALID, tSYSTIME.wMinute, -1)
End Function

Public Function Second() As Integer
    Second = IIf(pvGetSysTime = GDT_VALID, tSYSTIME.wSecond, -1)
End Function

Public Function Milliseconds() As Integer
    Milliseconds = IIf(pvGetSysTime = GDT_VALID, tSYSTIME.wMilliseconds, -1)
End Function

Public Function DayOfWeek() As Integer
    DayOfWeek = IIf(pvGetSysTime = GDT_VALID, tSYSTIME.wDayOfWeek, -1)
End Function

'----------------------------------------------------------------------------------------------------------------
'Public Property Get and Let
'----------------------------------------------------------------------------------------------------------------

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get hWndDTPicker() As Long
    hWndDTPicker = mDTPickerHwnd
End Property

Public Property Get UpDown() As Boolean
    UpDown = mUpDown
End Property

Public Property Let UpDown(ByVal nValue As Boolean)
    If mUpDown <> nValue Then
        mUpDown = nValue
        PropertyChanged "UpDown"
        Call pvCreate
    End If
End Property

Public Property Get CheckBox() As Boolean
    CheckBox = mCheckBox
End Property

Public Property Let CheckBox(ByVal nValue As Boolean)
    If mCheckBox <> nValue Then
        mCheckBox = nValue
        PropertyChanged "CheckBox"
        Call pvCreate
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal nValue As Boolean)
    If mEnabled <> nValue Then
        mEnabled = nValue
        UserControl.Enabled = nValue
        EnableWindow mDTPickerHwnd, nValue
        PropertyChanged "Enabled"
    End If
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal nFont As StdFont)
    Set UserControl.Font = nFont
    If Ambient.UserMode Then Set mFont = UserControl.Font
    UpdateFont
    PropertyChanged "Font"
End Property

Private Sub UpdateFont()
    Dim lFont As Long
    
    lFont = SendMessage(UserControl.hWnd, WM_GETFONT, 0&, 0&)
    Call SendMessage(mDTPickerHwnd, WM_SETFONT, lFont, 1)
End Sub

Public Property Let Font(ByVal nFont As StdFont)
    Set Font = nFont
End Property

Public Property Get Format() As vbExDTPickerFormatConstants
    Format = mFormat
End Property

Public Property Let Format(ByVal nValue As vbExDTPickerFormatConstants)
    Dim iFormat_Prev As vbExDTPickerFormatConstants
    
    If nValue <> mFormat Then
        iFormat_Prev = mFormat
        mFormat = nValue
        If mFormat = dtpTime Then
            If Not IsNull(mValue) Then
                If IsDate(mValue) Then
                    If (mValue < mMinDate) Or (mValue > mMaxDate) Then
                        If (Date >= mMinDate) And (Date <= mMaxDate) Then
                            mValue = Date + DatePart("h", mValue) + DatePart("n", mValue) + DatePart("s", mValue)
                        Else
                            mValue = mMinDate + DatePart("h", mValue) + DatePart("n", mValue) + DatePart("s", mValue)
                        End If
                    End If
                End If
            End If
        ElseIf iFormat_Prev = dtpTime Then
            If Not IsValidDate(mValue) Then
                mValue = Null
            End If
        End If
        PropertyChanged "Format"
        pvCreate
    End If
End Property

Public Property Get CustomFormat() As String
    CustomFormat = mCustomFormat
End Property

Public Property Let CustomFormat(ByVal nValue As String)
    If nValue <> mCustomFormat Then
        mCustomFormat = nValue
        PropertyChanged "CustomFormat"
        If mFormat = dtpCustom Then
            pvCreate
        Else
            If mFormat = dtpCustom Then
                SendMessage mDTPickerHwnd, DTM_SETFORMATA, 0&, ByVal mCustomFormat
            End If
        End If
    End If
End Property

Public Property Get MinDate() As Date
    MinDate = mMinDate
End Property

Public Property Let MinDate(ByVal nValue As Date)
    If nValue <> mMinDate Then
        If nValue >= cDTPickerMinDate Then
            mMinDate = nValue
            Call pvSetRange
            PropertyChanged "MinDate"
        Else
            RaiseError 380, TypeName(Me), "The minimum date must be equal or greater than " & cDTPickerMinDate
        End If
    End If
End Property

Public Property Get MaxDate() As Date
    MaxDate = mMaxDate
End Property

Public Property Let MaxDate(ByVal nValue As Date)
    If nValue <> mMaxDate Then
        If nValue <= cDTPickerMaxDate Then
            mMaxDate = nValue
            Call pvSetRange
            PropertyChanged "MaxDate"
        Else
            RaiseError 380, TypeName(Me), "The maximum date must be equal or lower than " & cDTPickerMaxDate
        End If
    End If
End Property


Public Property Get TextBackColor() As OLE_COLOR
    TextBackColor = mTextBackColor
End Property

Public Property Let TextBackColor(ByVal nValue As OLE_COLOR)
    Dim lColor As Long
    
    If mTextBackColor <> nValue Then
        mTextBackColor = nValue
        If hBrush <> 0 Then DeleteObject hBrush
        OleTranslateColor mTextBackColor, 0, lColor
        hBrush = CreateSolidBrush(lColor)
        PropertyChanged "TextBackColor"
        If mDTPickerHwnd Then
            Dim Rec As RECT
            
            GetClientRect mDTPickerHwnd, Rec
            InvalidateRect mDTPickerHwnd, Rec, 1
        End If
    End If
End Property

Public Property Get CalendarBackColor() As OLE_COLOR
    CalendarBackColor = mCalendarBackColor
End Property

Public Property Let CalendarBackColor(ByVal nValue As OLE_COLOR)
    If mCalendarBackColor <> nValue Then
        mCalendarBackColor = nValue
        PropertyChanged "CalendarBackColor"
        pvChangeColor MCSC_MONTHBK, mCalendarBackColor
    End If
End Property

Public Property Get CalendarForeColor() As OLE_COLOR
    CalendarForeColor = mCalendarForeColor
End Property

Public Property Let CalendarForeColor(ByVal nValue As OLE_COLOR)
    If mCalendarForeColor <> nValue Then
        mCalendarForeColor = nValue
        PropertyChanged "CalendarForeColor"
        pvChangeColor MCSC_TEXT, mCalendarForeColor
    End If
End Property

Public Property Get CalendarTitleBackColor() As OLE_COLOR
    CalendarTitleBackColor = mCalendarTitleBackColor
End Property

Public Property Let CalendarTitleBackColor(ByVal nValue As OLE_COLOR)
    If mCalendarTitleBackColor <> nValue Then
        mCalendarTitleBackColor = nValue
        PropertyChanged "CalendarTitleBackColor"
        pvChangeColor MCSC_TITLEBK, mCalendarTitleBackColor
    End If
End Property

Public Property Get CalendarTitleForeColor() As OLE_COLOR
    CalendarTitleForeColor = mCalendarTitleForeColor
End Property

Public Property Let CalendarTitleForeColor(ByVal nValue As OLE_COLOR)
    If mCalendarTitleForeColor <> nValue Then
        mCalendarTitleForeColor = nValue
        PropertyChanged "CalendarTitleForeColor"
        pvChangeColor MCSC_TITLETEXT, mCalendarTitleForeColor
    End If
End Property

Public Property Get CalendarTrailingForeColor() As OLE_COLOR
    CalendarTrailingForeColor = mCalendarTrailingForeColor
End Property

Public Property Let CalendarTrailingForeColor(ByVal nValue As OLE_COLOR)
    If mCalendarTrailingForeColor <> nValue Then
        mCalendarTrailingForeColor = nValue
        PropertyChanged "CalendarTrailingForeColor"
        pvChangeColor MCSC_TRAILINGTEXT, mCalendarTrailingForeColor
    End If
End Property

Public Property Get Value() As Variant
Attribute Value.VB_Description = "Devuelve o establece la fecha o tiempo actual."
Attribute Value.VB_MemberFlags = "200"
    Dim tST As SYSTEMTIME
    If SendMessage(mDTPickerHwnd, DTM_GETSYSTEMTIME, GDT_NONE, tST) = GDT_VALID Then
        If mFormat = dtpTime Then
            Value = VBA.DateTime.TimeSerial(tST.wHour, tST.wMinute, tST.wSecond)
        Else
            Value = VBA.DateTime.DateSerial(tST.wYear, tST.wMonth, tST.wDay)
        End If
    Else
        If mCheckBox Then
            Value = Null
        Else
            Value = ""
        End If
    End If
End Property

Public Property Let Value(ByVal nValue As Variant)
'    If nValue <> mValue Then
        If IsValidDate(nValue) Then
            mValue = nValue
            mValue = CDate(nValue)
            pvSetDateTime mValue
            Call SendMessage(mDTPickerHwnd, DTM_SETSYSTEMTIME, GDT_VALID, tSYSTIME)
        Else
            If nValue = "" Or IsNull(nValue) Then
                If mCheckBox Then
                    mValue = Null
                    Call SendMessage(mDTPickerHwnd, DTM_SETSYSTEMTIME, GDT_NONE, tSYSTIME)
                Else
                    mValue = ""
                    pvSetDateTime Now2
                    Call SendMessage(mDTPickerHwnd, DTM_SETSYSTEMTIME, GDT_VALID, tSYSTIME)
                End If
            Else
                RaiseError 380, TypeName(Me)
                Exit Property
            End If
        End If
        PropertyChanged "Value"
        RaiseEvent_Change
'    End If
End Property

Public Property Get GetIdealHeight() As Long
    If mDTPickerHwnd Then
        Dim pt As POINTAPI
        Call SendMessage(mDTPickerHwnd, DTM_GETIDEALSIZE, 0, pt)
        GetIdealHeight = ScaleY(pt.y, vbPixels, ScaleMode)
    End If
End Property

Public Property Get GetIdealWith() As Long
    If mDTPickerHwnd Then
        Dim pt As POINTAPI
        Call SendMessage(mDTPickerHwnd, DTM_GETIDEALSIZE, 0, pt)
        GetIdealWith = ScaleX(pt.x, vbPixels, ScaleMode)
    End If
End Property

' .....................
' Codigo del ucListView de Carles P.V.
Private Function pvShiftState() As Integer
    Dim lS As Integer
    If (GetAsyncKeyState(vbKeyShift) < 0) Then lS = lS Or vbShiftMask
    If (GetAsyncKeyState(vbKeyMenu) < 0) Then lS = lS Or vbAltMask
    If (GetAsyncKeyState(vbKeyControl) < 0) Then lS = lS Or vbCtrlMask
    pvShiftState = lS
End Function
        
Private Sub RaiseEvent_Change()
    Static sLast
    
    If IsNull(mValue) <> IsNull(sLast) Then
        RaiseEvent Change
        sLast = mValue
    ElseIf mValue <> sLast Then
        RaiseEvent Change
        sLast = mValue
    End If
End Sub


Public Property Get MousePointer() As vbExCC2MousePointerConstants
    MousePointer = mMousePointer
End Property

Public Property Let MousePointer(ByVal nValue As vbExCC2MousePointerConstants)
    If mMousePointer <> nValue Then
        On Error Resume Next
        mMousePointer = nValue
        UserControl.MousePointer = mMousePointer
        PropertyChanged "MousePointer"
    End If
End Property

Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = mMouseIcon
End Property

Public Property Set MouseIcon(ByVal nMouseIcon As IPictureDisp)
    Set mMouseIcon = nMouseIcon
    Set UserControl.MouseIcon = mMouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Let MouseIcon(ByVal nMouseIcon As IPictureDisp)
    Set MouseIcon = nMouseIcon
End Property


Public Property Get Object() As Object
    Set Object = Me
End Property


Public Property Get DroppedDown() As Boolean
Attribute DroppedDown.VB_MemberFlags = "400"
    DroppedDown = mDroppedDown
End Property

Public Property Let DroppedDown(ByVal nValue As Boolean)
    If Not Ambient.UserMode Then Exit Property
    If nValue <> mDroppedDown Then
        If mDTPickerHwnd <> 0 Then
            If nValue Then
                SendMessage mDTPickerHwnd, WM_SYSKEYDOWN, vbKeyDown, ByVal 0&
            Else
                SendMessage mDTPickerHwnd, DTM_CLOSEMONTHCAL, 0, ByVal 0&
                If mHwndCalendar <> 0 Then
                    If IsWindow(mHwndCalendar) <> 0 Then
                        If IsWindowVisible(mHwndCalendar) <> 0 Then
                            ShowWindow mHwndCalendar, SW_HIDE
                        End If
                    End If
                End If
            End If
        End If
    End If
End Property

Private Function GetLowWord(ByVal nLong As Long) As Long
    GetLowWord = nLong And &H7FFF&
    If (nLong And &H8000&) <> 0 Then
        GetLowWord = GetLowWord Or &HFFFF8000
    End If
End Function

Private Function GetHighWord(ByVal nLong As Long) As Long
    GetHighWord = (nLong And &H7FFF0000) \ &H10000
    If (nLong And &H80000000) <> 0 Then
        GetHighWord = GetHighWord Or &HFFFF8000
    End If
End Function

Private Function GetShiftFromwParam(ByVal wParam As Long) As ShiftConstants
    If (wParam And MK_SHIFT) <> 0 Then
        GetShiftFromwParam = vbShiftMask
    End If
    If (wParam And MK_CONTROL) <> 0 Then
        GetShiftFromwParam = GetShiftFromwParam Or vbCtrlMask
    End If
    If GetKeyState(vbKeyMenu) < 0 Then
        GetShiftFromwParam = GetShiftFromwParam Or vbAltMask
    End If
End Function

Private Function GetMouseBurttonFromwParam(ByVal wParam As Long) As MouseButtonConstants
    If (wParam And MK_LBUTTON) <> 0 Then
        GetMouseBurttonFromwParam = vbLeftButton
    End If
    If (wParam And MK_RBUTTON) <> 0 Then
        GetMouseBurttonFromwParam = GetMouseBurttonFromwParam Or vbRightButton
    End If
    If (wParam And MK_MBUTTON) <> 0 Then
        GetMouseBurttonFromwParam = GetMouseBurttonFromwParam Or vbMiddleButton
    End If
End Function

Public Sub Refresh()
    UserControl.Refresh
    RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Private Function Now2() As Date
    Now2 = Now
    If Now2 < mMinDate Then
        Now2 = MinDate
    ElseIf Now2 > mMaxDate Then
        Now2 = mMaxDate
    End If
End Function

Private Sub SubclassCalendar()
    If (Not mCalendarSubclassed) And (mHwndCalendar <> 0) Then
        AttachMessage Me, mHwndCalendar, WM_LBUTTONDOWN
        AttachMessage Me, mHwndCalendar, WM_LBUTTONUP
        AttachMessage Me, mHwndCalendar, WM_DESTROY
        mCalendarSubclassed = True
    End If
End Sub

Private Sub UnsubclassCalendar()
    If (mCalendarSubclassed) And (mHwndCalendar <> 0) Then
        DetachMessage Me, mHwndCalendar, WM_LBUTTONDOWN
        DetachMessage Me, mHwndCalendar, WM_LBUTTONUP
        DetachMessage Me, mHwndCalendar, WM_DESTROY
        mCalendarSubclassed = False
    End If
End Sub
