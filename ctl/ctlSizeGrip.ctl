VERSION 5.00
Begin VB.UserControl SizeGrip 
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MousePointer    =   8  'Size NW SE
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ToolboxBitmap   =   "ctlSizeGrip.ctx":0000
   Begin VB.Timer tmrResizingContWithParent 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   432
      Top             =   180
   End
   Begin VB.Image imgDesignMode 
      Height          =   228
      Left            =   1116
      Picture         =   "ctlSizeGrip.ctx":0312
      Top             =   2052
      Width           =   228
   End
   Begin VB.Image img21pix 
      Height          =   252
      Left            =   720
      Picture         =   "ctlSizeGrip.ctx":07C8
      Top             =   2052
      Width           =   252
   End
   Begin VB.Image img17pix 
      Height          =   204
      Left            =   432
      Picture         =   "ctlSizeGrip.ctx":0D4A
      Top             =   2052
      Width           =   204
   End
   Begin VB.Image img14pix 
      Height          =   168
      Left            =   180
      Picture         =   "ctlSizeGrip.ctx":1100
      Top             =   2052
      Width           =   168
   End
End
Attribute VB_Name = "SizeGrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Type COLORS_RGB
    r As Long
    G As Long
    b As Long
End Type

Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const SWP_NOACTIVATE = &H10&
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOMOVE As Long = &H2

Private Const SW_HIDE = 0
Private Const SW_SHOW = 5

Private Const SM_CXSIZEFRAME = 32&
Private Const SM_CYSIZEFRAME = 33&
Private Const SM_CXEDGE = 45
Private Const SM_CYEDGE = 46

Private Const WM_SYSCOMMAND = &H112&
Private Const SC_SIZE_SE  As Long = &HF008&

Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private WithEvents mPic As PictureBox
Attribute mPic.VB_VarHelpID = -1

Private mAutoResizeContainerAtCorner As Boolean
Private mAdditionalBorderSpace As Long

Private mBackColor As Long
Private mWidth As Long
Private mHeight As Long
Private mSettingImage As Boolean
Private mUserControlHwnd As Long
Private mContainerHwnd As Long
Private mParentHwnd As Long
Private mContainerIsNotParent As Boolean
Private mControlIsAtBottomRightCorner As Boolean


Private Sub mForm_Resize()
    Dim iRectC As RECT
    Dim iRectP As RECT
    Dim iHide As Boolean
    
    If mForm.WindowState = vbNormal Then
        UserControl.Width = ScaleX(mWidth, vbPixels, vbTwips)
        UserControl.Height = ScaleY(mHeight, vbPixels, vbTwips)
        PositionUserControl
        ShowWindow mUserControlHwnd, SW_SHOW
    Else
        If IsWindowVisible(mUserControlHwnd) <> 0 Then
            If mContainerIsNotParent Then
                iHide = mControlIsAtBottomRightCorner
            Else
                iHide = True
            End If
            If iHide Then
                ShowWindow mUserControlHwnd, SW_HIDE
            End If
        End If
    End If
    
    If mAutoResizeContainerAtCorner Then
        If mContainerIsNotParent Then
            If tmrResizingContWithParent.Enabled Or mControlIsAtBottomRightCorner Then
                GetWindowRect mContainerHwnd, iRectC
                GetWindowRect mParentHwnd, iRectP
                SetWindowPos mContainerHwnd, 0&, 0&, 0&, iRectP.Right - iRectC.Left - GetSystemMetrics(SM_CXSIZEFRAME), iRectP.Bottom - iRectC.Top - GetSystemMetrics(SM_CYSIZEFRAME), SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE
                PositionUserControl
            End If
        End If
    End If
    If mContainerIsNotParent Then
        mControlIsAtBottomRightCorner = ControlIsAtBottomRightCorner
    End If
End Sub

Private Sub mPic_Resize()
    If mForm.WindowState = vbNormal Then
        PositionUserControl
    End If
End Sub

Private Sub tmrResizingContWithParent_Timer()
    If (GetAsyncKeyState(vbKeyLButton) = 0) Then
        tmrResizingContWithParent.Enabled = False
    End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    Dim iBackColor As Long
    
    If Not mContainerIsNotParent Then
        If PropertyName = "BackColor" Then
            iBackColor = -1
            On Error Resume Next
            iBackColor = UserControl.Extender.Container.BackColor
            If iBackColor = -1 Then
                iBackColor = UserControl.Parent.BackColor
                If iBackColor = -1 Then
                    iBackColor = vbButtonFace
                End If
            End If
            BackColor = iBackColor
            On Error GoTo 0
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    mWidth = 19
    mHeight = 19
End Sub

Private Sub UserControl_InitProperties()
    Dim iBackColor As Long
    
    iBackColor = -1
    On Error Resume Next
    iBackColor = UserControl.Extender.Container.BackColor
    If iBackColor = -1 Then
        iBackColor = UserControl.Parent.BackColor
        If iBackColor = -1 Then
            iBackColor = vbButtonFace
        End If
    Else
        mContainerHwnd = UserControl.Extender.Container.hWnd
    End If
    If mContainerHwnd = 0 Then
        mContainerHwnd = UserControl.Parent.hWnd
    End If
    mParentHwnd = UserControl.Parent.hWnd
    mContainerIsNotParent = (mContainerHwnd <> mParentHwnd) And (mContainerHwnd <> 0)
    On Error GoTo 0
    mUserControlHwnd = UserControl.hWnd
    mAutoResizeContainerAtCorner = True
    mAdditionalBorderSpace = 0
    mBackColor = iBackColor
    SetImage
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim iHwnd As Long
    
    If mContainerIsNotParent Then
        If ControlIsAtBottomRightCorner Then
            iHwnd = mParentHwnd
            tmrResizingContWithParent.Enabled = False
            tmrResizingContWithParent.Enabled = True
        End If
        If iHwnd = 0 Then
            iHwnd = mContainerHwnd
        End If
    Else
        iHwnd = mParentHwnd
    End If
    
    ReleaseCapture
    PostMessage iHwnd, WM_SYSCOMMAND, SC_SIZE_SE, 0&
End Sub

Private Function ControlIsAtBottomRightCorner() As Boolean
    Dim iRectC As RECT
    Dim iRectP As RECT

    GetWindowRect mContainerHwnd, iRectC
    GetWindowRect mParentHwnd, iRectP
    If iRectC.Bottom >= iRectP.Bottom - (GetSystemMetrics(SM_CYSIZEFRAME) + 1) Then
        If iRectC.Right >= iRectP.Right - (GetSystemMetrics(SM_CXSIZEFRAME) + 1) Then
            ControlIsAtBottomRightCorner = True
        End If
    End If

End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim iCont As Object
    Dim iBackColor As Long
    
    iBackColor = -1
    On Error Resume Next
    iBackColor = UserControl.Extender.Container.BackColor
    If iBackColor = -1 Then
        iBackColor = UserControl.Parent.BackColor
        If iBackColor = -1 Then
            iBackColor = vbButtonFace
        End If
    End If
    mBackColor = iBackColor
    On Error GoTo 0
    
    mAutoResizeContainerAtCorner = PropBag.ReadProperty("AutoResizeContainerAtCorner", True)
    mAdditionalBorderSpace = PropBag.ReadProperty("AdditionalBorderSpace", 0)
    
    If Ambient.UserMode Then
        On Error Resume Next
        mUserControlHwnd = UserControl.hWnd
        mContainerHwnd = UserControl.Extender.Container.hWnd
        mParentHwnd = UserControl.Parent.hWnd
        If mContainerHwnd = 0 Then
            mContainerHwnd = UserControl.Parent.hWnd
        End If
        Set iCont = UserControl.Extender.Container
        On Error GoTo 0
        mContainerIsNotParent = (mContainerHwnd <> mParentHwnd) And (mContainerHwnd <> 0)
        
        If TypeOf iCont Is PictureBox Then
            Set mPic = UserControl.Extender.Container
            Set mForm = Parent
        Else
            Set mForm = Parent
        End If
    End If
    
    SetImage
End Sub

Private Sub UserControl_Resize()
    If Not mPic Is Nothing Then Exit Sub
    If Not mSettingImage Then
        If Ambient.UserMode Then
            UserControl.Width = ScaleX(mWidth, vbPixels, vbTwips)
            UserControl.Height = ScaleY(mHeight, vbPixels, vbTwips)
        Else
            UserControl.Width = ScaleX(19, vbPixels, vbTwips)
            UserControl.Height = ScaleY(19, vbPixels, vbTwips)
        End If
        PositionUserControl
    End If
End Sub

Private Sub UserControl_Show()
    If Not Ambient.UserMode Then
        PositionUserControl
    End If
End Sub

Private Sub PositionUserControl()
    Dim iRectC As RECT
    Dim iRectCtl As RECT
    Dim iHwnd As Long
    Dim iAmbientUserMode As Boolean
    
    On Error Resume Next
    iAmbientUserMode = Ambient.UserMode
    On Error GoTo 0
    
    iHwnd = mContainerHwnd
    If iHwnd = 0 Then
        iHwnd = mParentHwnd
    End If
    
    If iHwnd <> 0 And iAmbientUserMode And mContainerIsNotParent Then
        GetWindowRect iHwnd, iRectC
        iRectC.Right = iRectC.Right - iRectC.Left
        iRectC.Bottom = iRectC.Bottom - iRectC.Top
        GetWindowRect mUserControlHwnd, iRectCtl
        iRectCtl.Right = iRectCtl.Right - iRectCtl.Left
        iRectCtl.Bottom = iRectCtl.Bottom - iRectCtl.Top
        SetWindowPos mUserControlHwnd, 0&, iRectC.Right - iRectCtl.Right - GetSystemMetrics(SM_CXEDGE) - mAdditionalBorderSpace, iRectC.Bottom - iRectCtl.Bottom - GetSystemMetrics(SM_CYEDGE) - mAdditionalBorderSpace, 0&, 0&, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
    Else
        Dim iCtl As Control
        
        Err.Clear
        On Error Resume Next
        Set iCtl = Parent.Controls(Ambient.DisplayName)
        iCtl.Left = iCtl.Container.ScaleWidth - iCtl.Container.ScaleX(mWidth, vbPixels, iCtl.Container.ScaleMode) - mAdditionalBorderSpace * Screen.TwipsPerPixelX
        iCtl.Top = iCtl.Container.ScaleHeight - iCtl.Container.ScaleY(mHeight, vbPixels, iCtl.Container.ScaleMode) - mAdditionalBorderSpace * Screen.TwipsPerPixelY
        If Err.Number = 0 Then Exit Sub
        
        On Error GoTo EExit
        iCtl.Left = iCtl.Container.Width - UserControl.ScaleX(mWidth, vbPixels, vbTwips)
        iCtl.Top = iCtl.Container.Height - UserControl.ScaleY(mHeight, vbPixels, vbTwips)
        iCtl.ZOrder
    End If
    
EExit:
End Sub

Public Property Let BackColor(nValue As OLE_COLOR)
    If nValue <> mBackColor Then
        mBackColor = nValue
        PropertyChanged "BackColor"
        SetImage
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_MemberFlags = "200"
    BackColor = mBackColor
End Property

Private Sub UserControl_Terminate()
    Set mForm = Nothing
    Set mPic = Nothing
    tmrResizingContWithParent.Enabled = False
End Sub

Private Function RGBColor(nColor As Long) As COLORS_RGB
    If nColor < 0 Then Exit Function
    RGBColor.r = nColor And 255
    RGBColor.G = (nColor \ 256) And 255
    RGBColor.b = (nColor \ 65536) And 255
End Function

Private Sub SetImage()
    Dim y As Long
    Dim x As Long
    Dim iDarkColor As Long
    Dim iLightColor As Long
    Dim iBackColorRGB As COLORS_RGB
    Dim iAuxRGBColor As COLORS_RGB
    Dim iColor As Long
    Dim iBackColor As Long
    Dim iAmbientUserMode As Boolean
    Dim iPxColor As Long
    Dim iTx As Single
    Dim iPic As StdPicture
    
    mSettingImage = True
    iAmbientUserMode = Ambient.UserMode
    TranslateColor mBackColor, 0, iBackColor
    If iAmbientUserMode Then
        iTx = Screen.TwipsPerPixelX
        If iTx >= 15 Then
            Set iPic = img14pix.Picture
            mWidth = 14
            mHeight = 14
        ElseIf iTx >= 12 Then
            Set iPic = img17pix.Picture
            mWidth = 17
            mHeight = 17
        ElseIf iTx >= 10 Then
            Set iPic = img21pix.Picture
            mWidth = 21
            mHeight = 21
        ElseIf iTx >= 7 Then ' 192 DPI
            Set iPic = StretchPicNN(img14pix.Picture, 2)
            mWidth = 28
            mHeight = 28
        ElseIf iTx >= 6 Then
            Set iPic = StretchPicNN(img17pix.Picture, 2)
            mWidth = 34
            mHeight = 34
        ElseIf iTx >= 5 Then
            Set iPic = StretchPicNN(img21pix.Picture, 2)
            mWidth = 42
            mHeight = 42
        ElseIf iTx >= 4 Then  ' 289 a 360 DPI
            Set iPic = StretchPicNN(img17pix.Picture, 3)
            mWidth = 51
            mHeight = 51
        ElseIf iTx >= 3 Then   ' 361 a 480 DPI
            Set iPic = StretchPicNN(img21pix.Picture, 3)
            mWidth = 63
            mHeight = 63
        ElseIf iTx >= 2 Then   ' 481 a 720 DPI
            Set iPic = StretchPicNN(img21pix.Picture, 5)
            mWidth = 105
            mHeight = 105
        Else ' mayor a 720 DPI
            Set iPic = StretchPicNN(img21pix.Picture, 10)
            mWidth = 210
            mHeight = 210
        End If
    Else
        mWidth = 19
        mHeight = 19
        Set iPic = imgDesignMode.Picture
    End If
    UserControl.AutoRedraw = True
    UserControl.BackColor = iBackColor
    Set UserControl.Picture = Nothing
    UserControl.Cls
    
    iBackColorRGB = RGBColor(iBackColor)
    
    iAuxRGBColor.r = iBackColorRGB.r - 52
    iAuxRGBColor.G = iBackColorRGB.G - 52
    iAuxRGBColor.b = iBackColorRGB.b - 52
    If iAuxRGBColor.r < 0 Then iAuxRGBColor.r = 0
    If iAuxRGBColor.G < 0 Then iAuxRGBColor.G = 0
    If iAuxRGBColor.b < 0 Then iAuxRGBColor.b = 0
    iDarkColor = RGB(iAuxRGBColor.r, iAuxRGBColor.G, iAuxRGBColor.b)
    
    iAuxRGBColor.r = iBackColorRGB.r + 52
    iAuxRGBColor.G = iBackColorRGB.G + 52
    iAuxRGBColor.b = iBackColorRGB.b + 52
    If iAuxRGBColor.r > 255 Then iAuxRGBColor.r = 255
    If iAuxRGBColor.G > 255 Then iAuxRGBColor.G = 255
    If iAuxRGBColor.b > 255 Then iAuxRGBColor.b = 255
    iLightColor = RGB(iAuxRGBColor.r, iAuxRGBColor.G, iAuxRGBColor.b)
    
    UserControl.Width = ScaleX(mWidth * 2, vbPixels, vbTwips)
    UserControl.Height = ScaleY(mHeight, vbPixels, vbTwips)
    UserControl.PaintPicture iPic, mWidth, 0
    
    For y = 0 To UserControl.ScaleHeight - 1
        For x = mWidth To UserControl.ScaleWidth - 1
            iPxColor = GetPixel(UserControl.hDC, x, y)
            Select Case iPxColor
                Case 14215660
                    iColor = iBackColor
                Case 10597816, 10728632
                    iColor = iDarkColor
                Case 16777215
                    iColor = iLightColor
                Case Else
                    iColor = iPxColor
            End Select
            SetPixel UserControl.hDC, x - mWidth, y, iColor
        Next x
    Next y
    UserControl.Width = ScaleX(mWidth, vbPixels, vbTwips)
    Set iPic = UserControl.Image
    UserControl.Cls
    UserControl.PaintPicture iPic, 0, 0
    Set iPic = UserControl.Image
    UserControl.Cls
    Set UserControl.Picture = iPic
    mSettingImage = False
End Sub

Public Sub Refresh()
    mForm_Resize
End Sub

Private Function StretchPicNN(nPic As StdPicture, nFactor As Long) As StdPicture
    Dim iPicInfo As BITMAP
    Dim PicSizeW As Long
    Dim PicSizeH As Long
    Dim iW As Long
    Dim iH As Long
    
    iW = UserControl.Width
    iH = UserControl.Height
    
    GetObjectAPI nPic.Handle, Len(iPicInfo), iPicInfo
    PicSizeW = iPicInfo.bmWidth
    PicSizeH = iPicInfo.bmHeight
    
    UserControl.Width = PicSizeW * nFactor * Screen.TwipsPerPixelX
    UserControl.Height = PicSizeH * nFactor * Screen.TwipsPerPixelY
    
    UserControl.PaintPicture nPic, 0, 0, PicSizeW * nFactor, PicSizeH * nFactor
    
    Set StretchPicNN = UserControl.Image
    UserControl.Cls

    UserControl.Width = iW
    UserControl.Height = iH

End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "AutoResizeContainerAtCorner", mAutoResizeContainerAtCorner, True
    PropBag.WriteProperty "AdditionalBorderSpace", mAdditionalBorderSpace, 0
End Sub


Public Property Get AutoResizeContainerAtCorner() As Boolean
    AutoResizeContainerAtCorner = mAutoResizeContainerAtCorner
End Property

Public Property Let AutoResizeContainerAtCorner(nValue As Boolean)
    If nValue <> mAutoResizeContainerAtCorner Then
        mAutoResizeContainerAtCorner = nValue
        PropertyChanged ("AutoResizeContainerAtCorner")
    End If
End Property


Public Property Get AdditionalBorderSpace() As Long
    AdditionalBorderSpace = mAdditionalBorderSpace
End Property

Public Property Let AdditionalBorderSpace(nValue As Long)
    If nValue <> mAdditionalBorderSpace Then
        If nValue > 15 Then Err.Raise 1011, TypeName(Me), "The AdditionalBorderSpace value is too high."
        mAdditionalBorderSpace = nValue
        PositionUserControl
        PropertyChanged ("AdditionalBorderSpace")
    End If
End Property

