VERSION 5.00
Begin VB.UserControl ctlProgressBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   DrawStyle       =   2  'Dot
   HasDC           =   0   'False
   PropertyPages   =   "ctlProgressBar.ctx":0000
   ScaleHeight     =   150
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
   ToolboxBitmap   =   "ctlProgressBar.ctx":0037
End
Attribute VB_Name = "ctlProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Taken from Krool's Common Controls Replacement:
' URL: http://www.vbforums.com/showthread.php?698563-CommonControls-%28Replacement-of-the-MS-common-controls

Option Explicit
#If False Then
    Private PrbOrientationHorizontal, PrbOrientationVertical
    Private PrbScrollingStandard, PrbScrollingSmooth
    Private PrbStateInProgress, PrbStateError, PrbStatePaused
    Private PrbTaskBarStateNone, PrbTaskBarStateMarquee, PrbTaskBarStateInProgress, PrbTaskBarStateError, PrbTaskBarStatePaused
#End If
Public Enum CCRightToLeftModeConstants
    CCRightToLeftModeNoControl = 0
    CCRightToLeftModeVBAME = 1
    CCRightToLeftModeSystemLocale = 2
    CCRightToLeftModeUserLocale = 3
    CCRightToLeftModeOSLanguage = 4
End Enum
Public Enum OLEDropModeConstants
    OLEDropModeNone = vbOLEDropNone
    OLEDropModeManual = vbOLEDropManual
End Enum
Public Enum PrbOrientationConstants
    PrbOrientationHorizontal = 0
    PrbOrientationVertical = 1
End Enum
Public Enum PrbScrollingConstants
    PrbScrollingStandard = 0
    PrbScrollingSmooth = 1
End Enum
Private Const PBST_NORMAL As Long = 1
Private Const PBST_ERROR As Long = 2
Private Const PBST_PAUSED As Long = 3
Public Enum PrbStateConstants
    PrbStateInProgress = PBST_NORMAL
    PrbStateError = PBST_ERROR
    PrbStatePaused = PBST_PAUSED
End Enum
Private Const TBPF_NOPROGRESS As Long = 0
Private Const TBPF_INDETERMINATE As Long = 1
Private Const TBPF_NORMAL As Long = 2
Private Const TBPF_ERROR As Long = 4
Private Const TBPF_PAUSED As Long = 8
Public Enum PrbTaskBarStateConstants
    PrbTaskBarStateNone = TBPF_NOPROGRESS
    PrbTaskBarStateMarquee = TBPF_INDETERMINATE
    PrbTaskBarStateInProgress = TBPF_NORMAL
    PrbTaskBarStateError = TBPF_ERROR
    PrbTaskBarStatePaused = TBPF_PAUSED
End Enum
Private Enum VTableInterfaceConstants
    VTableInterfaceControl = 1
    VTableInterfaceInPlaceActiveObject = 2
    VTableInterfacePerPropertyBrowsing = 3
    VTableInterfaceEnumeration = 4
End Enum
Private Enum VTableIndexITaskBarList3Constants
    ' Ignore : ITaskBarList3QueryInterface = 1
    ' Ignore : ITaskBarList3AddRef = 2
    ' Ignore : ITaskBarList3Release = 3
    VTableIndexITaskBarList3HrInit = 4
    ' Ignore : ITaskBarList3AddTab = 5
    ' Ignore : ITaskBarList3DeleteTab = 6
    ' Ignore : ITaskBarList3ActivateTab = 7
    ' Ignore : ITaskBarList3SetActiveAlt = 8
    ' Ignore : ITaskBarList3MarkFullscreenWindow = 9
    VTableIndexITaskBarList3SetProgressValue = 10
    VTableIndexITaskBarList3SetProgressState = 11
    ' Ignore : ITaskBarList3RegisterTab = 12
    ' Ignore : ITaskBarList3UnregisterTab = 13
    ' Ignore : ITaskBarList3SetTabOrder = 14
    ' Ignore : ITaskBarList3SetTabActive = 15
    ' Ignore : ITaskBarList3ThumbBarAddButtons = 16
    ' Ignore : ITaskBarList3ThumbBarUpdateButtons = 17
    ' Ignore : ITaskBarList3ThumbBarSetImageList = 18
    ' Ignore : ITaskBarList3SetOverlayIcon = 19
    ' Ignore : ITaskBarList3SetThumbnailTooltip = 20
    ' Ignore : ITaskBarList3SetThumbnailClip = 21
End Enum
Private Type PBRANGE
    Min                  As Long
    Max                  As Long
End Type
Private Type TINITCOMMONCONTROLSEX
    dwSize               As Long
    dwICC                As Long
End Type
Private Type DLLVERSIONINFO
    cbSize               As Long
    dwMajor              As Long
    dwMinor              As Long
    dwBuildNumber        As Long
    dwPlatformID         As Long
End Type
Private Type TLOCALESIGNATURE
    lsUsb(0 To 15) As Byte
    lsCsbDefault(0 To 1) As Long
    lsCsbSupported(0 To 1) As Long
End Type
Private Const GWL_STYLE As Long = (-16)
Private Type RECT
    Left                 As Long
    Top                  As Long
    Right                As Long
    Bottom               As Long
End Type
Private Type TOOLINFO
    cbSize               As Long
    uFlags               As Long
    hWnd                 As Long
    uId                  As Long
    RC                   As RECT
    hInst                As Long
    lpszText             As Long
    lParam               As Long
End Type
Private Const CC_STDCALL As Long = 4
Private Const E_POINTER As Long = &H80004003
Private Const E_INVALIDARG As Long = &H80070057

Private Const STAP_ALLOW_CONTROLS As Long = (1 * (2 ^ 1))
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, ByRef pCLSID As Any) As Long
Private Declare Function CoCreateInstance Lib "ole32" (ByRef rclsid As Any, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByRef riid As Any, ByRef ppv As IUnknown) As Long
Private Declare Function GetAncestor Lib "user32" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorW" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32" (ByRef ICCEX As TINITCOMMONCONTROLSEX) As Long
Private Declare Function ActivateVisualStyles Lib "uxtheme" Alias "SetWindowTheme" (ByVal hWnd As Long, Optional ByVal pszSubAppName As Long = 0, Optional ByVal pszSubIdList As Long = 0) As Long
Private Declare Function RemoveVisualStyles Lib "uxtheme" Alias "SetWindowTheme" (ByVal hWnd As Long, Optional ByRef pszSubAppName As String = " ", Optional ByRef pszSubIdList As String = " ") As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal Color As Long, ByVal hPal As Long, ByRef RGBResult As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function FreeLibrary Lib "Kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function IsThemeActive Lib "uxtheme" () As Long
Private Declare Function IsAppThemed Lib "uxtheme" () As Long
Private Declare Function GetThemeAppProperties Lib "uxtheme" () As Long
Private Declare Function DllGetVersion Lib "comctl32" (ByRef pdvi As DLLVERSIONINFO) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemDefaultLangID Lib "Kernel32" () As Integer
Private Declare Function GetUserDefaultLangID Lib "Kernel32" () As Integer
Private Declare Function GetUserDefaultUILanguage Lib "Kernel32" () As Integer
Private Declare Function GetLocaleInfo Lib "Kernel32" ()
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As IUnknown, ByVal oVft As Long, ByVal CallConv As Long, ByVal vtReturn As Integer, ByVal cActuals As Long, ByVal prgvt As Long, ByVal prgpvarg As Long, ByRef pvargResult As Variant) As Long

Private Const ICC_PROGRESS_CLASS As Long = &H20
Private Const CLSID_ITaskBarList As String = "{56FDF344-FD6D-11D0-958A-006097C9A090}"
Private Const IID_ITaskBarList3 As String = "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"
Private Const CLSCTX_INPROC_SERVER As Long = 1, S_OK As Long = 0
Private Const RDW_UPDATENOW As Long = &H100, RDW_INVALIDATE As Long = &H1, RDW_ERASE As Long = &H4, RDW_ALLCHILDREN As Long = &H80
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD  As Long = &H40000000
Private Const WS_EX_STATICEDGE As Long = &H20000
Private Const WS_EX_LAYOUTRTL As Long = &H400000
Private Const SW_HIDE   As Long = &H0
Private Const GA_ROOT   As Long = 2
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_SETCURSOR As Long = &H20, HTCLIENT As Long = 1
Private Const CCM_FIRST As Long = &H2000
Private Const CCM_SETBKCOLOR As Long = (CCM_FIRST + 1)
Private Const WM_USER   As Long = &H400
Private Const PBM_SETBKCOLOR As Long = CCM_SETBKCOLOR
Private Const PBM_SETPOS As Long = (WM_USER + 2)
Private Const PBM_DELTAPOS As Long = (WM_USER + 3)
Private Const PBM_SETSTEP As Long = (WM_USER + 4)
Private Const PBM_STEPIT As Long = (WM_USER + 5)
Private Const PBM_SETRANGE32 As Long = (WM_USER + 6)
Private Const PBM_GETRANGE As Long = (WM_USER + 7)
Private Const PBM_GETPOS As Long = (WM_USER + 8)
Private Const PBM_SETBARCOLOR As Long = (WM_USER + 9)
Private Const PBM_SETMARQUEE As Long = (WM_USER + 10)
Private Const PBM_GETSTEP As Long = (WM_USER + 13)
Private Const PBM_SETSTATE As Long = (WM_USER + 16)
Private Const PBM_GETSTATE As Long = (WM_USER + 17)
Private Const PBS_SMOOTH As Long = &H1
Private Const PBS_VERTICAL As Long = &H4
Private Const PBS_MARQUEE As Long = &H8
Private Const PBS_SMOOTHREVERSE As Long = &H10
Implements ISubclass

Private ProgressBarHandle As Long
Private ProgressBarITaskBarList3 As IUnknown
Private ProgressBarIsClick As Boolean
Private DispIDMousePointer As Long
Private PropVisualStyles As Boolean
Private PropMousePointer As Integer, PropMouseIcon As IPictureDisp
Private PropRightToLeft As Boolean
Private PropRightToLeftLayout As Boolean
Private PropRightToLeftMode As CCRightToLeftModeConstants
Private PropRange       As PBRANGE
Private PropValue       As Long
Private PropStep        As Integer, PropStepAutoReset As Boolean
Private PropMarquee     As Boolean
Private PropMarqueeAnimation As Boolean, PropMarqueeSpeed As Long
Private PropOrientation As PrbOrientationConstants
Private PropScrolling   As PrbScrollingConstants
Private PropSmoothReverse As Boolean
Private PropBackColor   As OLE_COLOR
Private PropForeColor   As OLE_COLOR
Private PropState       As PrbStateConstants
Private ShellModHandle  As Long, ShellModCount As Long
Private mSubclassed     As Boolean

Private Sub IPerPropertyBrowsingVB_GetDisplayString(ByRef Handled As Boolean, ByVal DispID As Long, ByRef DisplayName As String)
    If DispID = DispIDMousePointer Then
        Call ComCtlsIPPBSetDisplayStringMousePointer(PropMousePointer, DisplayName)
        Handled = True
    End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedStrings(ByRef Handled As Boolean, ByVal DispID As Long, ByRef StringsOut() As String, ByRef CookiesOut() As Long)
    If DispID = DispIDMousePointer Then
        Call ComCtlsIPPBSetPredefinedStringsMousePointer(StringsOut(), CookiesOut())
        Handled = True
    End If
End Sub

Private Sub IPerPropertyBrowsingVB_GetPredefinedValue(ByRef Handled As Boolean, ByVal DispID As Long, ByVal Cookie As Long, ByRef Value As Variant)
    If DispID = DispIDMousePointer Then
        Value = Cookie
        Handled = True
    End If
End Sub

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrPreprocess
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, wParam As Long, lParam As Long, bConsume As Boolean) As Long
    Select Case iMsg
        Case WM_SETCURSOR
            If LoWord(lParam) = HTCLIENT Then
                If MousePointerID(PropMousePointer) <> 0 Then
                    SetCursor LoadCursor(0, MousePointerID(PropMousePointer))
                    bConsume = True
                    ISubclass_WindowProc = 1
                    Exit Function
                ElseIf PropMousePointer = 99 Then
                    If Not PropMouseIcon Is Nothing Then
                        SetCursor PropMouseIcon.Handle
                        bConsume = True
                        ISubclass_WindowProc = 1
                        Exit Function
                    End If
                End If
            End If
    End Select
    'ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
    Select Case iMsg
        Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN, WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
            Dim x                As Single
            Dim y                As Single
            x = UserControl.ScaleX(Get_X_lParam(lParam), vbPixels, vbTwips)
            y = UserControl.ScaleY(Get_Y_lParam(lParam), vbPixels, vbTwips)
            Select Case iMsg
                Case WM_LBUTTONDOWN
                    RaiseEvent MouseDown(vbLeftButton, GetShiftStateFromParam(wParam), x, y)
                    ProgressBarIsClick = True
                Case WM_MBUTTONDOWN
                    RaiseEvent MouseDown(vbMiddleButton, GetShiftStateFromParam(wParam), x, y)
                    ProgressBarIsClick = True
                Case WM_RBUTTONDOWN
                    RaiseEvent MouseDown(vbRightButton, GetShiftStateFromParam(wParam), x, y)
                    ProgressBarIsClick = True
                Case WM_MOUSEMOVE
                    RaiseEvent MouseMove(GetMouseStateFromParam(wParam), GetShiftStateFromParam(wParam), x, y)
                Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
                    Select Case iMsg
                        Case WM_LBUTTONUP
                            RaiseEvent MouseUp(vbLeftButton, GetShiftStateFromParam(wParam), x, y)
                        Case WM_MBUTTONUP
                            RaiseEvent MouseUp(vbMiddleButton, GetShiftStateFromParam(wParam), x, y)
                        Case WM_RBUTTONUP
                            RaiseEvent MouseUp(vbRightButton, GetShiftStateFromParam(wParam), x, y)
                    End Select
                    If ProgressBarIsClick = True Then
                        ProgressBarIsClick = False
                        If (x >= 0 And x <= UserControl.Width) And (y >= 0 And y <= UserControl.Height) Then RaiseEvent Click
                    End If
            End Select
    End Select

End Function

Private Sub UserControl_Initialize()
    Call ComCtlsLoadShellMod
    Call ComCtlsInitCC(ICC_PROGRESS_CLASS)
    'Call SetVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
'    DispIDMousePointer = GetDispID(Me, "MousePointer")
End Sub

Private Sub UserControl_InitProperties()
    PropVisualStyles = True
    PropMousePointer = 0: Set PropMouseIcon = Nothing
    PropRightToLeft = Ambient.RightToLeft
    PropRightToLeftLayout = False
    PropRightToLeftMode = CCRightToLeftModeVBAME
    If PropRightToLeft = True Then Me.RightToLeft = True
    PropRange.Min = 0
    PropRange.Max = 100
    PropValue = 0
    PropStep = 10
    PropStepAutoReset = True
    PropMarquee = False
    PropMarqueeAnimation = False
    PropMarqueeSpeed = 80
    PropOrientation = PrbOrientationHorizontal
    PropScrolling = PrbScrollingStandard
    PropSmoothReverse = False
    PropBackColor = vbButtonFace
    PropForeColor = vbHighlight
    PropState = PrbStateInProgress
    Call CreateProgressBar
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        PropVisualStyles = .ReadProperty("VisualStyles", True)
        Me.Enabled = .ReadProperty("Enabled", True)
        Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
        PropMousePointer = .ReadProperty("MousePointer", 0)
        Set PropMouseIcon = .ReadProperty("MouseIcon", Nothing)
        PropRightToLeft = .ReadProperty("RightToLeft", False)
        PropRightToLeftLayout = .ReadProperty("RightToLeftLayout", False)
        PropRightToLeftMode = .ReadProperty("RightToLeftMode", CCRightToLeftModeVBAME)
        If PropRightToLeft = True Then Me.RightToLeft = True
        PropRange.Min = .ReadProperty("Min", 0)
        PropRange.Max = .ReadProperty("Max", 100)
        PropValue = .ReadProperty("Value", 0)
        PropStep = .ReadProperty("Step", 1)
        PropStepAutoReset = .ReadProperty("StepAutoReset", True)
        PropMarquee = .ReadProperty("Marquee", False)
        PropMarqueeAnimation = .ReadProperty("MarqueeAnimation", False)
        PropMarqueeSpeed = .ReadProperty("MarqueeSpeed", 80)
        PropOrientation = .ReadProperty("Orientation", PrbOrientationHorizontal)
        PropScrolling = .ReadProperty("Scrolling", PrbScrollingStandard)
        PropSmoothReverse = .ReadProperty("SmoothReverse", PropSmoothReverse)
        PropBackColor = .ReadProperty("BackColor", vbButtonFace)
        PropForeColor = .ReadProperty("ForeColor", vbHighlight)
        PropState = .ReadProperty("State", PrbStateInProgress)
    End With
    Call CreateProgressBar
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "VisualStyles", PropVisualStyles, True
        .WriteProperty "Enabled", Me.Enabled, True
        .WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
        .WriteProperty "MousePointer", PropMousePointer, 0
        .WriteProperty "MouseIcon", PropMouseIcon, Nothing
        .WriteProperty "RightToLeft", PropRightToLeft, False
        .WriteProperty "RightToLeftLayout", PropRightToLeftLayout, False
        .WriteProperty "RightToLeftMode", PropRightToLeftMode, CCRightToLeftModeVBAME
        .WriteProperty "Min", PropRange.Min, 0
        .WriteProperty "Max", PropRange.Max, 100
        .WriteProperty "Value", PropValue, 0
        .WriteProperty "Step", PropStep, 1
        .WriteProperty "StepAutoReset", PropStepAutoReset, True
        .WriteProperty "Marquee", PropMarquee, False
        .WriteProperty "MarqueeAnimation", PropMarqueeAnimation, False
        .WriteProperty "MarqueeSpeed", PropMarqueeSpeed, 80
        .WriteProperty "Orientation", PropOrientation, PrbOrientationHorizontal
        .WriteProperty "Scrolling", PropScrolling, PrbScrollingStandard
        .WriteProperty "SmoothReverse", PropSmoothReverse, False
        .WriteProperty "BackColor", PropBackColor, vbButtonFace
        .WriteProperty "ForeColor", PropForeColor, vbHighlight
        .WriteProperty "State", PropState, PrbStateInProgress
    End With
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(x, vbPixels, vbContainerPosition), UserControl.ScaleY(y, vbPixels, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, UserControl.ScaleX(x, vbPixels, vbContainerPosition), UserControl.ScaleY(y, vbPixels, vbContainerPosition), State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Public Sub OLEDrag()
    UserControl.OLEDrag
End Sub

Private Sub UserControl_Resize()
    Static LastHeight As Single, LastWidth As Single, LastAlign As AlignConstants
    Static InProc As Boolean
    If InProc = True Then Exit Sub
    InProc = True
    With UserControl.Extender
        Select Case .Align
            Case LastAlign
            Case vbAlignNone
            Case vbAlignTop, vbAlignBottom
                Select Case LastAlign
                    Case vbAlignLeft, vbAlignRight
                        .Height = LastWidth
                End Select
            Case vbAlignLeft, vbAlignRight
                Select Case LastAlign
                    Case vbAlignTop, vbAlignBottom
                        .Width = LastHeight
                End Select
        End Select
        LastHeight = .Height
        LastWidth = .Width
        LastAlign = .Align
    End With
    With UserControl
        If DPICorrectionFactor() <> 1 Then
            .Extender.Move .Extender.Left + .ScaleX(1, vbPixels, vbContainerPosition), .Extender.Top + .ScaleY(1, vbPixels, vbContainerPosition)
            .Extender.Move .Extender.Left - .ScaleX(1, vbPixels, vbContainerPosition), .Extender.Top - .ScaleY(1, vbPixels, vbContainerPosition)
        End If
        If ProgressBarHandle <> 0 Then MoveWindow ProgressBarHandle, 0, 0, .ScaleWidth, .ScaleHeight, 1
    End With
    InProc = False
End Sub

Private Sub UserControl_Terminate()
    'Call RemoveVTableSubclass(Me, VTableInterfacePerPropertyBrowsing)
    Call DestroyProgressBar
    Call ComCtlsReleaseShellMod
End Sub

Public Property Get Name() As String
    Name = Ambient.DisplayName
End Property

Public Property Get Tag() As String
    Tag = Extender.Tag
End Property

Public Property Let Tag(ByVal Value As String)
    Extender.Tag = Value
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

Public Property Get Container() As Object
    Set Container = Extender.Container
End Property

Public Property Set Container(ByVal Value As Object)
    Set Extender.Container = Value
End Property

Public Property Get Left() As Single
    Left = Extender.Left
End Property

Public Property Let Left(ByVal Value As Single)
    Extender.Left = Value
End Property

Public Property Get Top() As Single
    Top = Extender.Top
End Property

Public Property Let Top(ByVal Value As Single)
    Extender.Top = Value
End Property

Public Property Get Width() As Single
    Width = Extender.Width
End Property

Public Property Let Width(ByVal Value As Single)
    Extender.Width = Value
End Property

Public Property Get Height() As Single
    Height = Extender.Height
End Property

Public Property Let Height(ByVal Value As Single)
    Extender.Height = Value
End Property

Public Property Get Visible() As Boolean
    Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal Value As Boolean)
    Extender.Visible = Value
End Property

Public Property Get ToolTipText() As String
    ToolTipText = Extender.ToolTipText
End Property

Public Property Let ToolTipText(ByVal Value As String)
    Extender.ToolTipText = Value
End Property

Public Property Get Align() As Integer
    Align = Extender.Align
End Property

Public Property Let Align(ByVal Value As Integer)
    Extender.Align = Value
End Property

Public Property Get DragIcon() As IPictureDisp
    Set DragIcon = Extender.DragIcon
End Property

Public Property Let DragIcon(ByVal Value As IPictureDisp)
    Extender.DragIcon = Value
End Property

Public Property Set DragIcon(ByVal Value As IPictureDisp)
    Set Extender.DragIcon = Value
End Property

Public Property Get DragMode() As Integer
    DragMode = Extender.DragMode
End Property

Public Property Let DragMode(ByVal Value As Integer)
    Extender.DragMode = Value
End Property

Public Sub Drag(Optional ByRef Action As Variant)
    If IsMissing(Action) Then Extender.Drag Else Extender.Drag Action
End Sub

Public Sub ZOrder(Optional ByRef Position As Variant)
    If IsMissing(Position) Then Extender.ZOrder Else Extender.ZOrder Position
End Sub

Public Property Get hWnd() As Long
    hWnd = ProgressBarHandle
End Property

Public Property Get hWndUserControl() As Long
    hWndUserControl = UserControl.hWnd
End Property

Public Property Get VisualStyles() As Boolean
    VisualStyles = PropVisualStyles
End Property

Public Property Let VisualStyles(ByVal Value As Boolean)
    PropVisualStyles = Value
    If ProgressBarHandle <> 0 And EnabledVisualStyles() = True Then
        Dim dwExStyle        As Long, dwExStyleOld As Long
        dwExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
        dwExStyleOld = dwExStyle
        If PropVisualStyles = True Then
            ActivateVisualStyles ProgressBarHandle
            If (dwExStyle And WS_EX_STATICEDGE) = WS_EX_STATICEDGE Then dwExStyle = dwExStyle And Not WS_EX_STATICEDGE
        Else
            RemoveVisualStyles ProgressBarHandle
            If Not (dwExStyle And WS_EX_STATICEDGE) = WS_EX_STATICEDGE Then dwExStyle = dwExStyle Or WS_EX_STATICEDGE
        End If
        If dwExStyle <> dwExStyleOld Then
            SetWindowLong ProgressBarHandle, GWL_EXSTYLE, dwExStyle
            Call ComCtlsFrameChanged(ProgressBarHandle)
        End If
        Me.Refresh
    End If
    UserControl.PropertyChanged "VisualStyles"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    UserControl.Enabled = Value
    UserControl.PropertyChanged "Enabled"
End Property

Public Property Get OLEDropMode() As OLEDropModeConstants
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal Value As OLEDropModeConstants)
    Select Case Value
        Case OLEDropModeNone, OLEDropModeManual
            UserControl.OLEDropMode = Value
        Case Else
            Err.Raise 380
    End Select
    UserControl.PropertyChanged "OLEDropMode"
End Property

Public Property Get MousePointer() As Integer
    MousePointer = PropMousePointer
End Property

Public Property Let MousePointer(ByVal Value As Integer)
    Select Case Value
        Case 0 To 16, 99
            PropMousePointer = Value
        Case Else
            Err.Raise 380
    End Select
    UserControl.PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = PropMouseIcon
End Property

Public Property Let MouseIcon(ByVal Value As IPictureDisp)
    Set Me.MouseIcon = Value
End Property

Public Property Set MouseIcon(ByVal Value As IPictureDisp)
    If Value Is Nothing Then
        Set PropMouseIcon = Nothing
    Else
        If Value.Type = vbPicTypeIcon Or Value.Handle = 0 Then
            Set PropMouseIcon = Value
        Else
            If Ambient.UserMode = False Then
                MsgBox "Invalid property value", vbCritical + vbOKOnly
                Exit Property
            Else
                Err.Raise 380
            End If
        End If
    End If
    UserControl.PropertyChanged "MouseIcon"
End Property

Public Property Get RightToLeft() As Boolean
    RightToLeft = PropRightToLeft
End Property

Public Property Let RightToLeft(ByVal Value As Boolean)
    PropRightToLeft = Value
    UserControl.RightToLeft = PropRightToLeft
    Call ComCtlsCheckRightToLeft(PropRightToLeft, UserControl.RightToLeft, PropRightToLeftMode)
    Dim dwMask           As Long
    If PropRightToLeft = True And PropRightToLeftLayout = True Then dwMask = WS_EX_LAYOUTRTL
    If Ambient.UserMode = True Then Call ComCtlsSetRightToLeft(UserControl.hWnd, dwMask)
    If ProgressBarHandle <> 0 Then Call ComCtlsSetRightToLeft(ProgressBarHandle, dwMask)
    UserControl.PropertyChanged "RightToLeft"
End Property

Public Property Get RightToLeftLayout() As Boolean
    RightToLeftLayout = PropRightToLeftLayout
End Property

Public Property Let RightToLeftLayout(ByVal Value As Boolean)
    PropRightToLeftLayout = Value
    Me.RightToLeft = PropRightToLeft
    UserControl.PropertyChanged "RightToLeftLayout"
End Property

Public Property Get RightToLeftMode() As CCRightToLeftModeConstants
    RightToLeftMode = PropRightToLeftMode
End Property

Public Property Let RightToLeftMode(ByVal Value As CCRightToLeftModeConstants)
    Select Case Value
        Case CCRightToLeftModeNoControl, CCRightToLeftModeVBAME, CCRightToLeftModeSystemLocale, CCRightToLeftModeUserLocale, CCRightToLeftModeOSLanguage
            PropRightToLeftMode = Value
        Case Else
            Err.Raise 380
    End Select
    Me.RightToLeft = PropRightToLeft
    UserControl.PropertyChanged "RightToLeftMode"
End Property

Public Property Get Min() As Long
    If ProgressBarHandle <> 0 Then
        Min = SendMessage(ProgressBarHandle, PBM_GETRANGE, 1, ByVal 0&)
    Else
        Min = PropRange.Min
    End If
End Property

Public Property Let Min(ByVal Value As Long)
    If Value < Me.Max Then
        PropRange.Min = Value
        PropRange.Max = Me.Max
        If PropValue < PropRange.Min Then PropValue = PropRange.Min
    Else
        If Ambient.UserMode = False Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
    If ProgressBarHandle <> 0 Then SendMessage ProgressBarHandle, PBM_SETRANGE32, PropRange.Min, ByVal PropRange.Max
    UserControl.PropertyChanged "Min"
End Property

Public Property Get Max() As Long
    If ProgressBarHandle = 0 Then
        Max = SendMessage(ProgressBarHandle, PBM_GETRANGE, 0, ByVal 0&)
    Else
        Max = PropRange.Max
    End If
End Property

Public Property Let Max(ByVal Value As Long)
    If Value > Me.Min Then
        PropRange.Min = Me.Min
        PropRange.Max = Value
        If PropValue > PropRange.Max Then PropValue = PropRange.Max
    Else
        If Ambient.UserMode = False Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
    If ProgressBarHandle <> 0 Then SendMessage ProgressBarHandle, PBM_SETRANGE32, PropRange.Min, ByVal PropRange.Max
    UserControl.PropertyChanged "Max"
End Property

Public Property Get Value() As Long
    If ProgressBarHandle <> 0 Then
        Value = SendMessage(ProgressBarHandle, PBM_GETPOS, 0, ByVal 0&)
    Else
        Value = PropValue
    End If
End Property

Public Property Let Value(ByVal NewValue As Long)
    If NewValue > Me.Max Then
        NewValue = Me.Max
    ElseIf NewValue < Me.Min Then
        NewValue = Me.Min
    End If
    PropValue = NewValue
    If ProgressBarHandle <> 0 Then SendMessage ProgressBarHandle, PBM_SETPOS, PropValue, ByVal 0&
    UserControl.PropertyChanged "Value"
End Property

Public Property Get Step() As Long
    If ProgressBarHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
        Step = SendMessage(ProgressBarHandle, PBM_GETSTEP, 0, ByVal 0&)
    Else
        Step = PropStep
    End If
End Property

Public Property Let Step(ByVal Value As Long)
    PropStep = Value
    If ProgressBarHandle <> 0 Then SendMessage ProgressBarHandle, PBM_SETSTEP, PropStep, ByVal 0&
    UserControl.PropertyChanged "Step"
End Property

Public Property Get StepAutoReset() As Boolean
    StepAutoReset = PropStepAutoReset
End Property

Public Property Let StepAutoReset(ByVal Value As Boolean)
    PropStepAutoReset = Value
    UserControl.PropertyChanged "StepAutoReset"
End Property

Public Property Get Marquee() As Boolean
    Marquee = PropMarquee
End Property

Public Property Let Marquee(ByVal Value As Boolean)
    PropMarquee = Value
    If ProgressBarHandle <> 0 Then Call ReCreateProgressBar
    UserControl.PropertyChanged "Marquee"
End Property

Public Property Get MarqueeAnimation() As Boolean
    MarqueeAnimation = PropMarqueeAnimation
End Property

Public Property Let MarqueeAnimation(ByVal Value As Boolean)
    PropMarqueeAnimation = Value
    If ProgressBarHandle <> 0 And ComCtlsSupportLevel() >= 1 Then SendMessage ProgressBarHandle, PBM_SETMARQUEE, IIf(PropMarqueeAnimation = True, 1, 0), ByVal PropMarqueeSpeed
    UserControl.PropertyChanged "MarqueeAnimation"
End Property

Public Property Get MarqueeSpeed() As Long
    MarqueeSpeed = PropMarqueeSpeed
End Property

Public Property Let MarqueeSpeed(ByVal Value As Long)
    If Value > 0 Then
        PropMarqueeSpeed = Value
    Else
        If Ambient.UserMode = False Then
            MsgBox "Invalid property value", vbCritical + vbOKOnly
            Exit Property
        Else
            Err.Raise 380
        End If
    End If
    If ProgressBarHandle <> 0 And ComCtlsSupportLevel() >= 1 Then SendMessage ProgressBarHandle, PBM_SETMARQUEE, IIf(PropMarqueeAnimation = True, 1, 0), ByVal PropMarqueeSpeed
    UserControl.PropertyChanged "MarqueeSpeed"
End Property

Public Property Get Orientation() As PrbOrientationConstants
    Orientation = PropOrientation
End Property

Public Property Let Orientation(ByVal Value As PrbOrientationConstants)
    Select Case Value
        Case PrbOrientationHorizontal, PrbOrientationVertical
            With UserControl
                If .Extender.Align = vbAlignNone And PropOrientation <> Value Then
                    If DPICorrectionFactor() <> 1 Then
                        .Extender.Move .Extender.Left, .Extender.Top, .Extender.Height, .Extender.Width
                    Else
                        .Size .ScaleX(.ScaleHeight, vbPixels, vbTwips), .ScaleY(.ScaleWidth, vbPixels, vbTwips)
                    End If
                End If
            End With
            PropOrientation = Value
        Case Else
            Err.Raise 380
    End Select
    If ProgressBarHandle <> 0 Then Call ReCreateProgressBar
    UserControl.PropertyChanged "Orientation"
End Property

Public Property Get Scrolling() As PrbScrollingConstants
    Scrolling = PropScrolling
End Property

Public Property Let Scrolling(ByVal Value As PrbScrollingConstants)
    Select Case Value
        Case PrbScrollingStandard, PrbScrollingSmooth
            PropScrolling = Value
        Case Else
            Err.Raise 380
    End Select
    If ProgressBarHandle <> 0 Then Call ReCreateProgressBar
    UserControl.PropertyChanged "Scrolling"
End Property

Public Property Get SmoothReverse() As Boolean
    SmoothReverse = PropSmoothReverse
End Property

Public Property Let SmoothReverse(ByVal Value As Boolean)
    PropSmoothReverse = Value
    If ProgressBarHandle <> 0 And ComCtlsSupportLevel() >= 1 Then Call ReCreateProgressBar
    UserControl.PropertyChanged "SmoothReverse"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = PropBackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
    PropBackColor = Value
    If ProgressBarHandle <> 0 Then SendMessage ProgressBarHandle, PBM_SETBKCOLOR, 0, ByVal WinColor(PropBackColor)
    UserControl.PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = PropForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
    PropForeColor = Value
    If ProgressBarHandle <> 0 Then SendMessage ProgressBarHandle, PBM_SETBARCOLOR, 0, ByVal WinColor(PropForeColor)
    UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get State() As PrbStateConstants
    If ProgressBarHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
        State = SendMessage(ProgressBarHandle, PBM_GETSTATE, 0, ByVal 0&)
    Else
        State = PropState
    End If
End Property

Public Property Let State(ByVal Value As PrbStateConstants)
    Select Case Value
        Case PrbStateInProgress, PrbStateError, PrbStatePaused
            PropState = Value
        Case Else
            Err.Raise 380
    End Select
    If ProgressBarHandle <> 0 And ComCtlsSupportLevel() >= 2 Then
        SendMessage ProgressBarHandle, PBM_SETSTATE, PropState, ByVal 0&
    End If
    UserControl.PropertyChanged "State"
End Property

Private Sub CreateProgressBar()
    If ProgressBarHandle <> 0 Then Exit Sub
    Dim dwStyle          As Long, dwExStyle As Long
    dwStyle = WS_CHILD Or WS_VISIBLE
    If PropRightToLeft = True And PropRightToLeftLayout = True Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
    If PropOrientation = PrbOrientationVertical Then dwStyle = dwStyle Or PBS_VERTICAL
    If PropScrolling = PrbScrollingSmooth Then dwStyle = dwStyle Or PBS_SMOOTH
    If ComCtlsSupportLevel() >= 1 Then
        If PropMarquee = True Then dwStyle = dwStyle Or PBS_MARQUEE
        If PropSmoothReverse = True Then dwStyle = dwStyle Or PBS_SMOOTHREVERSE
    End If
    ProgressBarHandle = CreateWindowEx(dwExStyle, StrPtr("msctls_progress32"), StrPtr("Progress Bar"), dwStyle, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
    If ProgressBarHandle <> 0 Then SendMessage ProgressBarHandle, PBM_SETRANGE32, PropRange.Min, ByVal PropRange.Max
    Me.VisualStyles = PropVisualStyles
    Me.Value = PropValue
    Me.Step = PropStep
    Me.MarqueeAnimation = PropMarqueeAnimation
    Me.BackColor = PropBackColor
    Me.ForeColor = PropForeColor
    Me.State = PropState
    If Ambient.UserMode = True Then
        If ProgressBarHandle <> 0 Then
            '        Call ComCtlsSetSubclass(ProgressBarHandle, Me, 0)
            mSubclassed = True
            AttachMessage Me, ProgressBarHandle, WM_SETCURSOR
            AttachMessage Me, ProgressBarHandle, WM_LBUTTONDOWN
            AttachMessage Me, ProgressBarHandle, WM_MBUTTONDOWN
            AttachMessage Me, ProgressBarHandle, WM_RBUTTONDOWN
            AttachMessage Me, ProgressBarHandle, WM_MOUSEMOVE
            AttachMessage Me, ProgressBarHandle, WM_LBUTTONUP
            AttachMessage Me, ProgressBarHandle, WM_MBUTTONUP
            AttachMessage Me, ProgressBarHandle, WM_RBUTTONUP
        End If
    End If
End Sub

Private Sub ReCreateProgressBar()
    If Ambient.UserMode = True Then
        Dim Locked           As Boolean
        Locked = CBool(LockWindowUpdate(UserControl.hWnd) <> 0)
        Call DestroyProgressBar
        Call CreateProgressBar
        Call UserControl_Resize
        If Locked = True Then LockWindowUpdate 0
        Me.Refresh
    Else
        Call DestroyProgressBar
        Call CreateProgressBar
        Call UserControl_Resize
    End If
End Sub

Private Sub DestroyProgressBar()
    If ProgressBarHandle = 0 Then Exit Sub
    If mSubclassed Then
        DetachMessage Me, ProgressBarHandle, WM_SETCURSOR
        DetachMessage Me, ProgressBarHandle, WM_LBUTTONDOWN
        DetachMessage Me, ProgressBarHandle, WM_MBUTTONDOWN
        DetachMessage Me, ProgressBarHandle, WM_RBUTTONDOWN
        DetachMessage Me, ProgressBarHandle, WM_MOUSEMOVE
        DetachMessage Me, ProgressBarHandle, WM_LBUTTONUP
        DetachMessage Me, ProgressBarHandle, WM_MBUTTONUP
        DetachMessage Me, ProgressBarHandle, WM_RBUTTONUP
        mSubclassed = False
    End If
    'Call ComCtlsRemoveSubclass(ProgressBarHandle, Me)
    ShowWindow ProgressBarHandle, SW_HIDE
    SetParent ProgressBarHandle, 0
    DestroyWindow ProgressBarHandle
    ProgressBarHandle = 0
End Sub

Public Sub Refresh()
    UserControl.Refresh
    RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
End Sub

Public Sub StepIt()
    If ProgressBarHandle = 0 Then Exit Sub
    If PropStepAutoReset = True Then
        SendMessage ProgressBarHandle, PBM_STEPIT, 0, ByVal 0&
        PropValue = Me.Value
    Else
        If Me.Value + Me.Step <= Me.Max Then
            SendMessage ProgressBarHandle, PBM_STEPIT, 0, ByVal 0&
            PropValue = Me.Value
        Else
            Me.Value = Me.Max
        End If
    End If
End Sub

Public Sub Increment(ByVal Delta As Long)
    If ProgressBarHandle <> 0 Then
        SendMessage ProgressBarHandle, PBM_DELTAPOS, Delta, ByVal 0&
        PropValue = Me.Value
    End If
End Sub

Private Function DPICorrectionFactor() As Single
    Static Done As Boolean, Value As Single
    If Done = False Then
        Value = Screen.TwipsPerPixelX / ((96 / DPI_X()) * 15)
        Done = True
    End If
End Function

Private Function DPI_X() As Long
    Const LOGPIXELSX As Long = 88
    Dim hDCScreen        As Long
    hDCScreen = GetDC(0)
    If hDCScreen <> 0 Then
        DPI_X = GetDeviceCaps(hDCScreen, LOGPIXELSX)
        ReleaseDC 0, hDCScreen
    End If
End Function

Private Sub ComCtlsLoadShellMod()
    If (ShellModHandle Or ShellModCount) = 0 Then ShellModHandle = LoadLibrary(StrPtr("Shell32.dll"))
    ShellModCount = ShellModCount + 1
End Sub

Private Sub ComCtlsInitCC(ByVal ICC As Long)
    Dim ICCEX            As TINITCOMMONCONTROLSEX
    With ICCEX
        .dwSize = LenB(ICCEX)
        .dwICC = ICC
    End With
    InitCommonControlsEx ICCEX
End Sub

Private Sub ComCtlsIPPBSetDisplayStringMousePointer(ByVal MousePointer As Integer, ByRef DisplayName As String)
    Select Case MousePointer
        Case 0: DisplayName = "0 - Default"
        Case 1: DisplayName = "1 - Arrow"
        Case 2: DisplayName = "2 - Cross"
        Case 3: DisplayName = "3 - I-Beam"
        Case 4: DisplayName = "4 - Hand"
        Case 5: DisplayName = "5 - Size"
        Case 6: DisplayName = "6 - Size NE SW"
        Case 7: DisplayName = "7 - Size N S"
        Case 8: DisplayName = "8 - Size NW SE"
        Case 9: DisplayName = "9 - Size W E"
        Case 10: DisplayName = "10 - Up Arrow"
        Case 11: DisplayName = "11 - Hourglass"
        Case 12: DisplayName = "12 - No Drop"
        Case 13: DisplayName = "13 - Arrow and Hourglass"
        Case 14: DisplayName = "14 - Arrow and Question"
        Case 15: DisplayName = "15 - Size All"
        Case 16: DisplayName = "16 - Arrow and CD"
        Case 99: DisplayName = "99 - Custom"
    End Select
End Sub

Private Sub ComCtlsIPPBSetPredefinedStringsMousePointer(ByRef StringsOut() As String, ByRef CookiesOut() As Long)
    ReDim StringsOut(0 To (17 + 1)) As String
    ReDim CookiesOut(0 To (17 + 1)) As Long
    StringsOut(0) = "0 - Default": CookiesOut(0) = 0
    StringsOut(1) = "1 - Arrow": CookiesOut(1) = 1
    StringsOut(2) = "2 - Cross": CookiesOut(2) = 2
    StringsOut(3) = "3 - I-Beam": CookiesOut(3) = 3
    StringsOut(4) = "4 - Hand": CookiesOut(4) = 4
    StringsOut(5) = "5 - Size": CookiesOut(5) = 5
    StringsOut(6) = "6 - Size NE SW": CookiesOut(6) = 6
    StringsOut(7) = "7 - Size N S": CookiesOut(7) = 7
    StringsOut(8) = "8 - Size NW SE": CookiesOut(8) = 8
    StringsOut(9) = "9 - Size W E": CookiesOut(9) = 9
    StringsOut(10) = "10 - Up Arrow": CookiesOut(10) = 10
    StringsOut(11) = "11 - Hourglass": CookiesOut(11) = 11
    StringsOut(12) = "12 - No Drop": CookiesOut(12) = 12
    StringsOut(13) = "13 - Arrow and Hourglass": CookiesOut(13) = 13
    StringsOut(14) = "14 - Arrow and Question": CookiesOut(14) = 14
    StringsOut(15) = "15 - Size All": CookiesOut(15) = 15
    StringsOut(16) = "16 - Arrow and CD": CookiesOut(16) = 16
    StringsOut(17) = "99 - Custom": CookiesOut(17) = 99
End Sub

Private Function LoWord(ByVal DWord As Long) As Integer
    If DWord And &H8000& Then
        LoWord = DWord Or &HFFFF0000
    Else
        LoWord = DWord And &HFFFF&
    End If
End Function

Private Function MousePointerID(ByVal MousePointer As Integer) As Long
    Select Case MousePointer
        Case vbArrow
            Const IDC_ARROW As Long = 32512
            MousePointerID = IDC_ARROW
        Case vbCrosshair
            Const IDC_CROSS As Long = 32515
            MousePointerID = IDC_CROSS
        Case vbIbeam
            Const IDC_IBEAM As Long = 32513
            MousePointerID = IDC_IBEAM
        Case vbIconPointer ' Obselete, replaced Icon with Hand
            Const IDC_HAND As Long = 32649
            MousePointerID = IDC_HAND
        Case vbSizePointer, vbSizeAll
            Const IDC_SIZEALL As Long = 32646
            MousePointerID = IDC_SIZEALL
        Case vbSizeNESW
            Const IDC_SIZENESW As Long = 32643
            MousePointerID = IDC_SIZENESW
        Case vbSizeNS
            Const IDC_SIZENS As Long = 32645
            MousePointerID = IDC_SIZENS
        Case vbSizeNWSE
            Const IDC_SIZENWSE As Long = 32642
            MousePointerID = IDC_SIZENWSE
        Case vbSizeWE
            Const IDC_SIZEWE As Long = 32644
            MousePointerID = IDC_SIZEWE
        Case vbUpArrow
            Const IDC_UPARROW As Long = 32516
            MousePointerID = IDC_UPARROW
        Case vbHourglass
            Const IDC_WAIT As Long = 32514
            MousePointerID = IDC_WAIT
        Case vbNoDrop
            Const IDC_NO As Long = 32648
            MousePointerID = IDC_NO
        Case vbArrowHourglass
            Const IDC_APPSTARTING As Long = 32650
            MousePointerID = IDC_APPSTARTING
        Case vbArrowQuestion
            Const IDC_HELP As Long = 32651
            MousePointerID = IDC_HELP
        Case 16
            Const IDC_WAITCD As Long = 32663 ' Undocumented
            MousePointerID = IDC_WAITCD
    End Select
End Function

Private Function Get_X_lParam(ByVal lParam As Long) As Long
    Get_X_lParam = lParam And &H7FFF&
    If lParam And &H8000& Then Get_X_lParam = Get_X_lParam Or &HFFFF8000
End Function

Private Function Get_Y_lParam(ByVal lParam As Long) As Long
    Get_Y_lParam = (lParam And &H7FFF0000) \ &H10000
    If lParam And &H80000000 Then Get_Y_lParam = Get_Y_lParam Or &HFFFF8000
End Function

Private Function GetShiftStateFromParam(ByVal wParam As Long) As ShiftConstants
    Const MK_SHIFT As Long = &H4, MK_CONTROL As Long = &H8
    If (wParam And MK_SHIFT) = MK_SHIFT Then GetShiftStateFromParam = vbShiftMask
    If (wParam And MK_CONTROL) = MK_CONTROL Then GetShiftStateFromParam = GetShiftStateFromParam Or vbCtrlMask
    If GetKeyState(vbKeyMenu) < 0 Then GetShiftStateFromParam = GetShiftStateFromParam Or vbAltMask
End Function

Private Function GetMouseStateFromParam(ByVal wParam As Long) As MouseButtonConstants
    Const MK_LBUTTON As Long = &H1, MK_RBUTTON As Long = &H2, MK_MBUTTON As Long = &H10
    If (wParam And MK_LBUTTON) = MK_LBUTTON Then GetMouseStateFromParam = vbLeftButton
    If (wParam And MK_RBUTTON) = MK_RBUTTON Then GetMouseStateFromParam = GetMouseStateFromParam Or vbRightButton
    If (wParam And MK_MBUTTON) = MK_MBUTTON Then GetMouseStateFromParam = GetMouseStateFromParam Or vbMiddleButton
End Function

Private Sub ComCtlsReleaseShellMod()
    ShellModCount = ShellModCount - 1
    If ShellModCount = 0 And ShellModHandle <> 0 Then
        FreeLibrary ShellModHandle
        ShellModHandle = 0
    End If
End Sub

Private Function EnabledVisualStyles() As Boolean
    If GetComCtlVersion() >= 6 Then
        If IsThemeActive() <> 0 Then
            If IsAppThemed() <> 0 Then
                EnabledVisualStyles = True
            ElseIf (GetThemeAppProperties() And STAP_ALLOW_CONTROLS) <> 0 Then
                EnabledVisualStyles = True
            End If
        End If
    End If
End Function

Private Function GetComCtlVersion() As Long
    Static Done As Boolean, Value As Long
    If Done = False Then
        Dim Version          As DLLVERSIONINFO
        On Error Resume Next
        Version.cbSize = LenB(Version)
        If DllGetVersion(Version) = S_OK Then Value = Version.dwMajor
        Done = True
    End If
    GetComCtlVersion = Value
End Function

Private Sub ComCtlsFrameChanged(ByVal hWnd As Long)
    Const SWP_FRAMECHANGED As Long = &H20, SWP_NOMOVE As Long = &H2, SWP_NOOWNERZORDER As Long = &H200, SWP_NOSIZE As Long = &H1, SWP_NOZORDER As Long = &H4
    SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub

Private Sub ComCtlsCheckRightToLeft(ByRef Value As Boolean, ByVal UserControlValue As Boolean, ByVal ModeValue As CCRightToLeftModeConstants)
    If Value = False Then Exit Sub
    Select Case ModeValue
        Case CCRightToLeftModeNoControl
        Case CCRightToLeftModeVBAME
            Value = UserControlValue
        Case CCRightToLeftModeSystemLocale, CCRightToLeftModeUserLocale, CCRightToLeftModeOSLanguage
            Const LOCALE_FONTSIGNATURE As Long = &H58, SORT_DEFAULT As Long = &H0
            Dim LangID           As Integer, LCID As Long, LocaleSig As TLOCALESIGNATURE
            Select Case ModeValue
                Case CCRightToLeftModeSystemLocale
                    LangID = GetSystemDefaultLangID()
                Case CCRightToLeftModeUserLocale
                    LangID = GetUserDefaultLangID()
                Case CCRightToLeftModeOSLanguage
                    LangID = GetUserDefaultUILanguage()
            End Select
            LCID = (SORT_DEFAULT * &H10000) Or LangID
            If GetLocaleInfo(LCID, LOCALE_FONTSIGNATURE, VarPtr(LocaleSig), (LenB(LocaleSig) / 2)) <> 0 Then
                ' Unicode subset bitfield 0 to 127. Bit 123 = Layout progress, horizontal from right to left
                Value = CBool((LocaleSig.lsUsb(15) And (2 ^ (4 - 1))) <> 0)
            End If
    End Select
End Sub

Private Sub ComCtlsSetRightToLeft(ByVal hWnd As Long, ByVal dwMask As Long)
    Const WS_EX_LAYOUTRTL As Long = &H400000, WS_EX_RTLREADING As Long = &H2000, WS_EX_RIGHT As Long = &H1000, WS_EX_LEFTSCROLLBAR As Long = &H4000
    ' WS_EX_LAYOUTRTL will take care of both layout and reading order with the single flag and mirrors the window.
    Dim dwExStyle        As Long
    dwExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    If (dwExStyle And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Then dwExStyle = dwExStyle And Not WS_EX_LAYOUTRTL
    If (dwExStyle And WS_EX_RTLREADING) = WS_EX_RTLREADING Then dwExStyle = dwExStyle And Not WS_EX_RTLREADING
    If (dwExStyle And WS_EX_RIGHT) = WS_EX_RIGHT Then dwExStyle = dwExStyle And Not WS_EX_RIGHT
    If (dwExStyle And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR Then dwExStyle = dwExStyle And Not WS_EX_LEFTSCROLLBAR
    If (dwMask And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Then dwExStyle = dwExStyle Or WS_EX_LAYOUTRTL
    If (dwMask And WS_EX_RTLREADING) = WS_EX_RTLREADING Then dwExStyle = dwExStyle Or WS_EX_RTLREADING
    If (dwMask And WS_EX_RIGHT) = WS_EX_RIGHT Then dwExStyle = dwExStyle Or WS_EX_RIGHT
    If (dwMask And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR Then dwExStyle = dwExStyle Or WS_EX_LEFTSCROLLBAR
    Const WS_POPUP As Long = &H80000000
    If (GetWindowLong(hWnd, GWL_STYLE) And WS_POPUP) = 0 Then
        SetWindowLong hWnd, GWL_EXSTYLE, dwExStyle
        InvalidateRect hWnd, ByVal 0&, 1
        Call ComCtlsFrameChanged(hWnd)
    Else
        ' ToolTip control supports only the WS_EX_LAYOUTRTL flag.
        ' Set TTF_RTLREADING flag when dwMask contains WS_EX_RTLREADING, though WS_EX_RTLREADING will not be actually set.
        If (dwExStyle And WS_EX_RTLREADING) = WS_EX_RTLREADING Then dwExStyle = dwExStyle And Not WS_EX_RTLREADING
        If (dwExStyle And WS_EX_RIGHT) = WS_EX_RIGHT Then dwExStyle = dwExStyle And Not WS_EX_RIGHT
        If (dwExStyle And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR Then dwExStyle = dwExStyle And Not WS_EX_LEFTSCROLLBAR
        SetWindowLong hWnd, GWL_EXSTYLE, dwExStyle
        Const TTM_SETTOOLINFOW As Long = (WM_USER + 54)
        Const TTM_SETTOOLINFO As Long = TTM_SETTOOLINFOW
        Const TTM_GETTOOLCOUNT As Long = (WM_USER + 13)
        Const TTM_ENUMTOOLSW As Long = (WM_USER + 58)
        Const TTM_ENUMTOOLS As Long = TTM_ENUMTOOLSW
        Const TTM_UPDATE As Long = (WM_USER + 29)
        Const TTF_RTLREADING As Long = &H4
        Dim i                As Long, TI As TOOLINFO, Buffer As String
        With TI
            .cbSize = LenB(TI)
            Buffer = String(80, vbNullChar)
            .lpszText = StrPtr(Buffer)
            For i = 1 To SendMessage(hWnd, TTM_GETTOOLCOUNT, 0, ByVal 0&)
                If SendMessage(hWnd, TTM_ENUMTOOLS, i - 1, ByVal VarPtr(TI)) <> 0 Then
                    If (dwMask And WS_EX_LAYOUTRTL) = WS_EX_LAYOUTRTL Or (dwMask And WS_EX_RTLREADING) = 0 Then
                        If (.uFlags And TTF_RTLREADING) = TTF_RTLREADING Then .uFlags = .uFlags And Not TTF_RTLREADING
                    Else
                        If (.uFlags And TTF_RTLREADING) = 0 Then .uFlags = .uFlags Or TTF_RTLREADING
                    End If
                    SendMessage hWnd, TTM_SETTOOLINFO, 0, ByVal VarPtr(TI)
                    SendMessage hWnd, TTM_UPDATE, 0, ByVal 0&
                End If
            Next i
        End With
    End If
End Sub

Private Function ComCtlsSupportLevel() As Byte
    Static Done As Boolean, Value As Byte
    If Done = False Then
        Dim Version          As DLLVERSIONINFO
        On Error Resume Next
        Version.cbSize = LenB(Version)
        Const S_OK As Long = &H0
        If DllGetVersion(Version) = S_OK Then
            If Version.dwMajor = 6 And Version.dwMinor = 0 Then
                Value = 1
            ElseIf Version.dwMajor > 6 Or (Version.dwMajor = 6 And Version.dwMinor > 0) Then
                Value = 2
            End If
        End If
        Done = True
    End If
    ComCtlsSupportLevel = Value
End Function

Private Function WinColor(ByVal Color As Long, Optional ByVal hPal As Long) As Long
    If OleTranslateColor(Color, hPal, WinColor) <> 0 Then WinColor = -1
End Function

Private Function VTableCall(ByVal RetType As VbVarType, ByVal OLEInstance As IUnknown, ByVal Entry As Long, ParamArray ArgList() As Variant) As Variant
    Entry = Entry - 1
    Debug.Assert Not (Entry < 0 Or OLEInstance Is Nothing)
    Dim VarArgList       As Variant, HResult As Long
    VarArgList = ArgList
    If UBound(VarArgList) > -1 Then
        Dim i                As Long, ArrVarType() As Integer, ArrVarPtr() As Long
        ReDim ArrVarType(LBound(VarArgList) To UBound(VarArgList)) As Integer
        ReDim ArrVarPtr(LBound(VarArgList) To UBound(VarArgList)) As Long
        For i = LBound(VarArgList) To UBound(VarArgList)
            ArrVarType(i) = VarType(VarArgList(i))
            ArrVarPtr(i) = VarPtr(VarArgList(i))
        Next i
        HResult = DispCallFunc(OLEInstance, Entry * 4, CC_STDCALL, RetType, i, VarPtr(ArrVarType(0)), VarPtr(ArrVarPtr(0)), VTableCall)
    Else
        HResult = DispCallFunc(OLEInstance, Entry * 4, CC_STDCALL, RetType, 0, 0, 0, VTableCall)
    End If
    Select Case HResult
        Case S_OK
        Case E_INVALIDARG
            Err.Raise Number:=HResult, Description:="One of the arguments was invalid"
        Case E_POINTER
            Err.Raise Number:=HResult, Description:="Function address was null"
        Case Else
            Err.Raise HResult
    End Select
End Function

