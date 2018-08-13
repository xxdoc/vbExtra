VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolTipEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ISubclass

Private Const WM_DESTROY As Long = &H2&
Private Const WM_WINDOWPOSCHANGED As Long = &H47&
Private Const WM_MOVE As Long = &H3&

Public Event Closed()
Public Event BeforeClose()

Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_HWNDPARENT As Long = (-8)
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

''Windows API Constants
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000

''Windows API Types
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

''Tooltip Window Constants
Private Const TTS_NOPREFIX = &H2
'Private Const TTF_TRANSPARENT = &H100
'Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = (WM_USER + 4)
'Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
'Private Const TTF_IDISHWND = &H1
'Private Const TTM_SETDelayTimeSeconds = (WM_USER + 3)
'Private Const TTDT_AUTOPOP = 2
'Private Const TTDT_INITIAL = 3
Private Const TTS_CLOSE As Long = &H80

Private Const TTM_TRACKPOSITION = (WM_USER + 18)
Private Const TTM_TRACKACTIVATE As Long = (WM_USER + 17)
Private Const TTF_ABSOLUTE As Long = &H80
Private Const TTF_TRACK As Long = &H20

Private Const TOOLTIPS_CLASSA = "tooltips_class32"

''Tooltip Window Types
Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    hWnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type


Public Enum vbExBalloonTooltipIconConstants
    vxTTNoIcon = 0
    vxTTIconInfo = 1
    vxTTIconWarning = 2
    vxTTIconError = 3
End Enum

Public Enum vbExBalloonTooltipStyleConstants
    vxTTStandard
    vxTTBalloon
End Enum

Private Type POINTAPI
    x As Long
    y As Long
End Type


'local variable(s) to hold property value(s)
Private mBackColor As Variant
Private mForeColor As Variant
Private mTitle As String
Private mIcon As vbExBalloonTooltipIconConstants
Private mStyle As vbExBalloonTooltipStyleConstants
Private mTipText As String
Private mDelayTimeSeconds As Variant
Private mVisibleTimeSeconds As Variant
Private mPositionX As Variant
Private mPositionY As Variant
Private mCloseButton As Boolean
Private mWidth As Variant
Private mRestrictMouseMoveToTwips As Long
Private mTag As String

'private data
Private mTTHwnd As Long ' hWnd of the tooltip
Private mParentHwnd As Long ' hWnd of the window the tooltip attached to
Private mTi As TOOLINFO
Private mOldOwner As Long
Private mTTShown As Boolean
Private mDetached As Boolean
Private mShowTime As Single
Private mStartTime As Single
Private mMouseStartPosition As POINTAPI
Private mDelayTimeSecondsSng As Single
Private mVisibleTimeSecondsSng As Single
Private mParentOriginalRectPos As RECT
Private mTTOriginalRectPos As RECT

Private WithEvents mTmrDelay As cTimer
Attribute mTmrDelay.VB_VarHelpID = -1
Private WithEvents mTmrClose As cTimer
Attribute mTmrClose.VB_VarHelpID = -1
Private WithEvents mTmrTrackMouse As cTimer
Attribute mTmrTrackMouse.VB_VarHelpID = -1

Private mTTCollection As cToolTipExCollection

Public Property Let PositionX(ByVal nValue As Variant)
    If Not IsEmpty(nValue) Then
        mPositionX = Val(nValue)
    End If
End Property

Public Property Get PositionX() As Variant
   PositionX = mPositionX
End Property

Public Property Let PositionY(ByVal nValue As Variant)
    If Not IsEmpty(nValue) Then
        mPositionY = Val(nValue)
    End If
End Property

Public Property Get PositionY() As Variant
   PositionY = mPositionY
End Property

Public Property Let CloseButton(ByVal nValue As Boolean)
   mCloseButton = nValue
End Property

Public Property Get CloseButton() As Boolean
   CloseButton = mCloseButton
End Property

Public Property Let Width(ByVal nValue As Variant)
   If IsEmpty(nValue) Then Exit Property
   mWidth = Val(nValue)
End Property

Public Property Get Width() As Variant
   Width = mWidth
End Property

Public Property Let RestrictMouseMoveToTwips(ByVal nValue As Long)
   mRestrictMouseMoveToTwips = nValue
End Property

Public Property Get RestrictMouseMoveToTwips() As Long
   RestrictMouseMoveToTwips = mRestrictMouseMoveToTwips
End Property

Public Property Let Style(ByVal nValue As vbExBalloonTooltipStyleConstants)
   mStyle = nValue
End Property

Public Property Get Style() As vbExBalloonTooltipStyleConstants
   Style = mStyle
End Property

Public Function Create(ByVal nParentHwnd As Long) As Boolean
    Dim lWinStyle As Long
    Dim iAbsolutePosition As Boolean
    Dim iPt As POINTAPI
    Dim iLng As Long
    Dim iPositionX As Variant
    Dim iPositionY As Variant
    
    If IsWindowVisible(nParentHwnd) = 0 Then Exit Function
    
    mParentHwnd = nParentHwnd
    
    If Not IsEmpty(mPositionX) Or Not IsEmpty(mPositionY) Then
        iAbsolutePosition = True
        If IsEmpty(mPositionX) Or IsEmpty(mPositionY) Then
            GetCursorPos iPt
            ScreenToClient mParentHwnd, iPt
            If IsEmpty(mPositionX) Then
                iPositionX = iPt.x * Screen.TwipsPerPixelX
            Else
                iPositionX = mPositionX
            End If
            If IsEmpty(mPositionY) Then
                iPositionY = iPt.y * Screen.TwipsPerPixelY
            Else
                iPositionY = mPositionY
            End If
        Else
            iPositionX = mPositionX
            iPositionY = mPositionY
        End If
    End If
    
    CloseTip
    
    lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    
    If mCloseButton Then
        lWinStyle = lWinStyle Or TTS_CLOSE
        
        If IsEmpty(mDelayTimeSeconds) Then
            mDelayTimeSecondsSng = 0 ' default
        Else
            mDelayTimeSecondsSng = mDelayTimeSeconds
        End If
        If IsEmpty(mVisibleTimeSeconds) Then
            mVisibleTimeSecondsSng = 120 ' default
        Else
            mVisibleTimeSecondsSng = mVisibleTimeSeconds
        End If
    Else
        If IsEmpty(mDelayTimeSeconds) Then
            mDelayTimeSecondsSng = 0.5 ' default
        Else
            mDelayTimeSecondsSng = mDelayTimeSeconds
        End If
        If IsEmpty(mVisibleTimeSeconds) Then
            mVisibleTimeSecondsSng = 15 ' default
        Else
            mVisibleTimeSecondsSng = mVisibleTimeSeconds
        End If
    End If
    
    ''create baloon style if desired
    If mStyle = vxTTBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON
    
    mTTHwnd = CreateWindowEx(0&, _
        TOOLTIPS_CLASSA, _
        vbNullString, _
        lWinStyle, _
        CW_USEDEFAULT, _
        CW_USEDEFAULT, _
        CW_USEDEFAULT, _
        CW_USEDEFAULT, _
        0&, _
        0&, _
        App.hInstance, _
        0&)
               
    If Not IsEmpty(mWidth) Then
        If mWidth > 0 Then
            SendMessageLong mTTHwnd, TTM_SETMAXTIPWIDTH, 0, CLng(mWidth / Screen.TwipsPerPixelX)
        Else
            SendMessageLong mTTHwnd, TTM_SETMAXTIPWIDTH, 0, Screen.Width / Screen.TwipsPerPixelX / 2 ' to enable multiline and limit the width to half the screen size
        End If
    Else
        SendMessageLong mTTHwnd, TTM_SETMAXTIPWIDTH, 0, Screen.Width / Screen.TwipsPerPixelX / 2 ' to enable multiline and limit the width to half the screen size
    End If
               
    ''now set our tooltip info structure
    With mTi
        If iAbsolutePosition Then
           .lFlags = TTF_TRACK Or TTF_ABSOLUTE
        Else
           .lFlags = TTF_SUBCLASS  'Or TTF_TRACK Or TTF_ABSOLUTE
        End If
        
        ''set the hWnd prop to our parent control's hWnd
        .hWnd = mParentHwnd
        .lId = mParentHwnd '0
        .hInstance = App.hInstance
        '.lpstr = ALREADY SET
        '.lpRect = lpRect
        .lSize = Len(mTi)
    End With
    
    ''add the tooltip structure
    SendMessage mTTHwnd, TTM_ADDTOOLA, 0&, mTi
    
    ''if we want a title or we want an icon
    If mTitle <> "" Then
        SendMessage mTTHwnd, TTM_SETTITLE, CLng(mIcon), ByVal mTitle
    Else
         If (mIcon <> vxTTNoIcon) Or mCloseButton Then
            SendMessage mTTHwnd, TTM_SETTITLE, CLng(mIcon), ByVal " "
         End If
    End If
    
    If Not IsEmpty(mBackColor) Then
        TranslateColor CLng(mBackColor), 0&, iLng
        SendMessage mTTHwnd, TTM_SETTIPBKCOLOR, iLng, 0&
    End If
    If Not IsEmpty(mForeColor) Then
        TranslateColor CLng(mForeColor), 0&, iLng
        SendMessage mTTHwnd, TTM_SETTIPTEXTCOLOR, iLng, 0&
    End If
    
    If iAbsolutePosition Then
        iPt.x = iPositionX / Screen.TwipsPerPixelX
        iPt.y = iPositionY / Screen.TwipsPerPixelY
        ClientToScreen mParentHwnd, iPt
        SendMessageLong mTTHwnd, TTM_TRACKPOSITION, 0&, MakeLong(iPt.x, iPt.y)
    End If
    
    mOldOwner = SetOwner(mTTHwnd, mParentHwnd)
    
    AttachMessage Me, mTTHwnd, WM_WINDOWPOSCHANGED
    AttachMessage Me, mTTHwnd, WM_DESTROY
    AttachMessage Me, mParentHwnd, WM_WINDOWPOSCHANGED
    AttachMessage Me, mParentHwnd, WM_DESTROY
    AttachMessage Me, mParentHwnd, WM_MOVE
    
    mDetached = False
    
    If mDelayTimeSecondsSng = 0 Then
        SendMessage mTTHwnd, TTM_TRACKACTIVATE, 1&, mTi
    Else
        Set mTmrDelay = New cTimer
        mTmrDelay.Interval = 100
        mStartTime = Timer
    End If
    
    If mRestrictMouseMoveToTwips > 0 Then
        If Not (mCloseButton And (mRestrictMouseMoveToTwips = 1000)) Then
            GetCursorPos mMouseStartPosition
            Set mTmrTrackMouse = New cTimer
            mTmrTrackMouse.Interval = 100
        End If
    End If
    
    GetWindowRect mParentHwnd, mParentOriginalRectPos
    
    Create = mTTHwnd <> 0
End Function

Private Function MakeLong(ByVal wLow As Integer, ByVal wHigh As Integer) As Long
    MakeLong = wHigh * &H10000 + wLow
End Function

Public Property Let Icon(ByVal nValue As vbExBalloonTooltipIconConstants)
   mIcon = nValue
   If mTTHwnd <> 0 And mTitle <> Empty And mIcon <> vxTTNoIcon Then
      SendMessage mTTHwnd, TTM_SETTITLE, CLng(mIcon), ByVal mTitle
   End If
End Property

Public Property Get Icon() As vbExBalloonTooltipIconConstants
   Icon = mIcon
End Property

Public Property Let BackColor(ByVal nValue As Variant)
   If Not IsNumeric(nValue) Then Exit Property
   mBackColor = nValue
   If mTTHwnd <> 0 Then
      SendMessage mTTHwnd, TTM_SETTIPBKCOLOR, mBackColor, 0&
   End If
End Property

Public Property Get BackColor() As Variant
   BackColor = mBackColor
End Property

Public Property Let ForeColor(ByVal nValue As Variant)
   If Not IsNumeric(nValue) Then Exit Property
   mForeColor = nValue
   If mTTHwnd <> 0 Then
      SendMessage mTTHwnd, TTM_SETTIPTEXTCOLOR, mForeColor, 0&
   End If
End Property

Public Property Get ForeColor() As Variant
   ForeColor = mForeColor
End Property

Public Property Let Title(ByVal nValue As String)
   mTitle = nValue
   If mTTHwnd <> 0 And mTitle <> Empty And mIcon <> vxTTNoIcon Then
      SendMessage mTTHwnd, TTM_SETTITLE, CLng(mIcon), ByVal mTitle
   End If
End Property

Public Property Get Title() As String
   Title = mTitle
End Property

Public Property Let TipText(ByVal nValue As String)
   mTipText = nValue
   mTi.lpStr = nValue
   If mTTHwnd <> 0 Then
      SendMessage mTTHwnd, TTM_UPDATETIPTEXTA, 0&, mTi
   End If
End Property

Public Property Get TipText() As String
   TipText = mTipText
End Property

Private Sub Class_Terminate()
   CloseTip
End Sub

Public Sub CloseTip()
    RaiseEvent BeforeClose
    If mTTHwnd <> 0 Then
        SetOwner mTTHwnd, mOldOwner
        DestroyWindow mTTHwnd
        mTTHwnd = 0
    End If
    DestroyTimers
End Sub

Private Sub DestroyTimers()
    If Not mTmrDelay Is Nothing Then
        mTmrDelay.Interval = 0
        Set mTmrDelay = Nothing
    End If
    If Not mTmrClose Is Nothing Then
        mTmrClose.Interval = 0
        Set mTmrClose = Nothing
    End If
    If Not mTmrTrackMouse Is Nothing Then
        mTmrTrackMouse.Interval = 0
        Set mTmrTrackMouse = Nothing
    End If
End Sub

Public Property Get VisibleTimeSeconds() As Variant
   VisibleTimeSeconds = mVisibleTimeSeconds
End Property

Public Property Let VisibleTimeSeconds(ByVal nValue As Variant)
   If IsEmpty(nValue) Then Exit Property
   mVisibleTimeSeconds = nValue
End Property

Public Property Get DelayTimeSeconds() As Variant
   DelayTimeSeconds = mDelayTimeSeconds
End Property

Public Property Let DelayTimeSeconds(ByVal nValue As Variant)
   If IsEmpty(nValue) Then Exit Property
   mDelayTimeSeconds = nValue
End Property

Public Property Get ToolTipHwnd() As Long
    ToolTipHwnd = mTTHwnd
End Property

Public Property Get ParentHwnd() As Long
    ParentHwnd = mParentHwnd
End Property

Private Function SetOwner(ByVal HwndWindow, ByVal HwndofOwner) As Long
    SetOwner = SetWindowLong(HwndWindow, GWL_HWNDPARENT, HwndofOwner)
End Function

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrPreprocess
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Dim iRect As RECT
    
    Select Case iMsg
        Case WM_DESTROY
            RaiseEvent BeforeClose
            DetachMessage Me, mTTHwnd, WM_DESTROY
            If Not mDetached Then
                DetachMessage Me, mTTHwnd, WM_WINDOWPOSCHANGED
                DetachMessage Me, mParentHwnd, WM_WINDOWPOSCHANGED
                DetachMessage Me, mParentHwnd, WM_DESTROY
                DetachMessage Me, mParentHwnd, WM_MOVE
            End If
            DestroyTimers
            If Not mTTCollection Is Nothing Then
                mTTCollection.Remove mTTHwnd
                Set mTTCollection = Nothing
            End If
            RaiseEvent Closed
        Case WM_WINDOWPOSCHANGED
            If hWnd = mTTHwnd Then
                If IsWindowVisible(mTTHwnd) = 0 Then
                    If mTTShown Then
                        DetachMessage Me, mTTHwnd, WM_WINDOWPOSCHANGED
                        DetachMessage Me, mParentHwnd, WM_WINDOWPOSCHANGED
                        DetachMessage Me, mParentHwnd, WM_DESTROY
                        DetachMessage Me, mParentHwnd, WM_MOVE
                        mDetached = True
                        mTTShown = False
                        CloseTip
                    End If
                Else
                    If Not mTTShown Then
                        mTTShown = True
                        Set mTmrClose = New cTimer
                        mTmrClose.Interval = 100
                        mShowTime = Timer
                        GetWindowRect mTTHwnd, mTTOriginalRectPos
                    End If
                End If
            Else ' parent
                If IsWindowVisible(mParentHwnd) = 0 Then
                    DetachMessage Me, mTTHwnd, WM_WINDOWPOSCHANGED
                    DetachMessage Me, mParentHwnd, WM_WINDOWPOSCHANGED
                    DetachMessage Me, mParentHwnd, WM_DESTROY
                    DetachMessage Me, mParentHwnd, WM_MOVE
                    mDetached = True
                    mTTShown = False
                    CloseTip
                End If
            End If
        Case WM_MOVE
            GetWindowRect mParentHwnd, iRect
            MoveWindow mTTHwnd, mTTOriginalRectPos.Left + iRect.Left - mParentOriginalRectPos.Left, mTTOriginalRectPos.Top + iRect.Top - mParentOriginalRectPos.Top, mTTOriginalRectPos.Right - mTTOriginalRectPos.Left, mTTOriginalRectPos.Bottom - mTTOriginalRectPos.Top, 1&
    End Select
    
'    ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
End Function

Private Sub mTmrClose_ThatTime()
    If (Timer - mShowTime) >= mVisibleTimeSecondsSng Then
        mTmrClose.Interval = 0
        Set mTmrClose = Nothing
        CloseTip
    End If
End Sub

Private Sub mTmrDelay_ThatTime()
    If (Timer - mStartTime) >= mDelayTimeSecondsSng Then
        mTmrDelay.Interval = 0
        Set mTmrDelay = Nothing
        SendMessage mTTHwnd, TTM_TRACKACTIVATE, 1&, mTi
    End If
End Sub

Friend Property Set TTCollection(nCol As cToolTipExCollection)
    Set mTTCollection = nCol
End Property

Public Sub Reset()
    If mTTShown Then
        mShowTime = Timer
    End If
End Sub

Private Sub mTmrTrackMouse_ThatTime()
    Dim iMpos As POINTAPI
    Dim iPixX As Long
    Dim iPixY As Long
    Dim iClose As Boolean
    
    GetCursorPos iMpos
    iPixX = mRestrictMouseMoveToTwips / Screen.TwipsPerPixelX
    iPixY = mRestrictMouseMoveToTwips / Screen.TwipsPerPixelY
    
    If iPixX = 0 Then
        iClose = True
    End If
    If Abs(iMpos.x - mMouseStartPosition.x) > iPixX Then
        iClose = True
    End If
    If Abs(iMpos.y - mMouseStartPosition.y) > iPixY Then
        iClose = True
    End If
    If iClose Then
        mTmrTrackMouse.Interval = 0
        Set mTmrTrackMouse = Nothing
        CloseTip
    End If
End Sub

Public Sub RaiseEventClosed()
Attribute RaiseEventClosed.VB_MemberFlags = "40"
    RaiseEvent Closed
End Sub

Public Sub RaiseEventBeforeClose()
    RaiseEvent BeforeClose
End Sub


Public Property Let Tag(nValue As String)
    mTag = nValue
End Property

Public Property Get Tag() As String
    Tag = mTag
End Property
