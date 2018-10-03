VERSION 5.00
Begin VB.UserControl ToolBarDA 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   PropertyPages   =   "ctlToolBarDA.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ctlToolBarDA.ctx":002A
   Begin VB.Timer tmrFirstResize 
      Interval        =   1
      Left            =   90
      Top             =   510
   End
   Begin vbExtra.ButtonExNoFocus btnButton 
      Height          =   330
      Index           =   0
      Left            =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   656
      _ExtentY        =   572
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      UseMaskCOlor    =   -1  'True
   End
End
Attribute VB_Name = "ToolBarDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Private Enum efnToolBarDAAlignConstants
    efnTBAlignNone = 0
    efnTBAlignTop = 1
    efnTBAlignBottom = 2
End Enum

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long

Public Event ButtonClick(Button As ToolBarDAButton)
Public Event Click()
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event Resize()

Private mButtons As ToolBarDAButtons
Private mMaskColor As Long
Private mUseMaskColor As Boolean
Private mBackColor As Long
Private mDecreaseButtonHeight As Long
Private mAvailableWidth As Long
Private mAlign As efnToolBarDAAlignConstants
Private mLeaveBorderSpace As Boolean
Private mShowToolTipTextWhenDisabled As Boolean
Private mAutoSize As Boolean

Private mPicHeight16 As Long
Private mPicHeight20 As Long
Private mPicHeight24 As Long
Private mPicHeight30 As Long
Private mPicHeight36 As Long

Private mInterButtonSpaceInPixels As Long
Private mButtonWidth As Long
Private mVisibleButtonsCount As Long
Private mUsercontrolHeight As Long
Private mUsercontrolWidth As Long
Private mRedraw As Boolean
Private mIconsSize As vbExToolbarDAIconsSizeConstants
Private mRefreshPending As Boolean
Private mAllButtonsWidth As Long
Private mHiddenButtonsCount As Long
Private mLastAvailableWidth As Long
Private mNeedToHide As Boolean
Private mHiddenAtWidth As Long
Private mRefreshing As Boolean
Private mHwndParent As Long
Private mUserControlHwnd As Long
Private mAutoHiddden As Boolean

Private Sub btnButton_Click(Index As Integer)
    Dim iButton As ToolBarDAButton
    Static sInside As Boolean
    Dim c As Long
    Dim iButton2 As ToolBarDAButton
    
    If sInside Then Exit Sub
    sInside = True
    If Not mRefreshing Then
        Set iButton = GetButtonByButtonIndex(Val(btnButton(Index).Tag))
        
        If (iButton.Style = vxTBButtonGroup) Or (iButton.Style = vxTBCheck) Then
            iButton.Checked = btnButton(Index).Value
        End If
        If (iButton.Style = vxTBButtonGroup) Then
            If iButton.Checked Then
                c = iButton.Index
                If c > 0 Then
                    Do
                        c = c - 1
                        If c = 0 Then Exit Do
                        Set iButton2 = mButtons(c)
                        If iButton2.Style <> vxTBButtonGroup Then Exit Do
                        iButton2.Checked = False
                    Loop
                End If
                c = iButton.Index
                If c < (mButtons.Count + 1) Then
                    Do
                        c = c + 1
                        If c = (mButtons.Count + 1) Then Exit Do
                        Set iButton2 = mButtons(c)
                        If iButton2.Style <> vxTBButtonGroup Then Exit Do
                        iButton2.Checked = False
                    Loop
                End If
            End If
        End If
        sInside = False
        RaiseEvent ButtonClick(iButton)
    End If
    sInside = False
End Sub

Private Sub tmrFirstResize_Timer()
    tmrFirstResize.Enabled = False
    UserControl_Resize
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    Set mButtons = New ToolBarDAButtons
    mAvailableWidth = -1
    mRedraw = True
    InitGlobal
End Sub

Private Sub UserControl_InitProperties()
    Dim iButton As ToolBarDAButton
    
    mIconsSize = vxIconsAppDefault
    mMaskColor = &HFF00FF
    mUseMaskColor = True
    mBackColor = vbButtonFace
    mShowToolTipTextWhenDisabled = False
    mAutoSize = True
    Set iButton = New ToolBarDAButton
    
    If Ambient.UserMode Then
        On Error Resume Next
        mHwndParent = Parent.hWnd
        mUserControlHwnd = UserControl.hWnd
        On Error GoTo 0
        If mHwndParent <> 0 Then
            AttachMessage Me, mHwndParent, WM_SIZE
        End If
    End If
    
    mButtons.AddButtonObject iButton
    SetPicHeights
    Refresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim c As Long
    Dim iLng As Long
    Dim iButton As ToolBarDAButton
    
    mIconsSize = PropBag.ReadProperty("IconsSize", vxIconsAppDefault)
    mLeaveBorderSpace = PropBag.ReadProperty("LeaveBorderSpace", False)
    mShowToolTipTextWhenDisabled = PropBag.ReadProperty("ShowToolTipTextWhenDisabled", False)
    mAutoSize = PropBag.ReadProperty("AutoSize", True)
    btnButton(0).ShowToolTipTextWhenDisabled = mShowToolTipTextWhenDisabled
    
    mPicHeight16 = PropBag.ReadProperty("PicHeight16", 16)
    mPicHeight20 = PropBag.ReadProperty("PicHeight20", 20)
    mPicHeight24 = PropBag.ReadProperty("PicHeight24", 24)
    mPicHeight30 = PropBag.ReadProperty("PicHeight30", 30)
    mPicHeight36 = PropBag.ReadProperty("PicHeight36", 36)

    mMaskColor = PropBag.ReadProperty("MaskColor", &HFF00FF)
    mUseMaskColor = PropBag.ReadProperty("UseMaskColor", True)
    mBackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    mDecreaseButtonHeight = PropBag.ReadProperty("DecreaseButtonHeight", 0)
    iLng = PropBag.ReadProperty("ButtonsCount", 0)
    If iLng > 100 Then iLng = 100
    If iLng < 1 Then iLng = 1

    For c = 1 To iLng
        Set iButton = New ToolBarDAButton
        iButton.Checked = PropBag.ReadProperty("ButtonChecked" & CStr(c), False)
        iButton.Enabled = PropBag.ReadProperty("ButtonEnabled" & CStr(c), True)
        iButton.Index = c
        iButton.Key = PropBag.ReadProperty("ButtonKey" & CStr(c), "")
        iButton.Tag = PropBag.ReadProperty("ButtonTag" & CStr(c), "")
        Set iButton.Pic16 = PropBag.ReadProperty("ButtonPic16" & CStr(c), Nothing)
        Set iButton.Pic20 = PropBag.ReadProperty("ButtonPic20" & CStr(c), Nothing)
        Set iButton.Pic24 = PropBag.ReadProperty("ButtonPic24" & CStr(c), Nothing)
        Set iButton.Pic30 = PropBag.ReadProperty("ButtonPic30" & CStr(c), Nothing)
        Set iButton.Pic36 = PropBag.ReadProperty("ButtonPic36" & CStr(c), Nothing)
        Set iButton.Pic16Alt = PropBag.ReadProperty("ButtonPic16Alt" & CStr(c), Nothing)
        Set iButton.Pic20Alt = PropBag.ReadProperty("ButtonPic20Alt" & CStr(c), Nothing)
        Set iButton.Pic24Alt = PropBag.ReadProperty("ButtonPic24Alt" & CStr(c), Nothing)
        Set iButton.Pic30Alt = PropBag.ReadProperty("ButtonPic30Alt" & CStr(c), Nothing)
        Set iButton.Pic36Alt = PropBag.ReadProperty("ButtonPic36Alt" & CStr(c), Nothing)
        iButton.UseAltPic = PropBag.ReadProperty("UseAltPic" & CStr(c), False)
        iButton.Width = PropBag.ReadProperty("ButtonWidth" & CStr(c), 360)
        iButton.Style = PropBag.ReadProperty("ButtonStyle" & CStr(c), vxTBDefault)
        iButton.ToolTipText = PropBag.ReadProperty("ButtonToolTipText" & CStr(c), "")
        iButton.Visible = PropBag.ReadProperty("ButtonVisible" & CStr(c), True)
        iButton.OrderToHide = PropBag.ReadProperty("ButtonOrderToHide" & CStr(c), 0)
        If iButton.Key = "" Then
            mButtons.AddButtonObject iButton
        Else
            On Error Resume Next
            mButtons.AddButtonObject iButton, iButton.Key
            On Error GoTo 0
        End If
    Next c
    
    If Ambient.UserMode Then
        On Error Resume Next
        mHwndParent = Parent.hWnd
        mUserControlHwnd = UserControl.hWnd
        On Error GoTo 0
        If mHwndParent <> 0 Then
            AttachMessage Me, mHwndParent, WM_SIZE
        End If
    End If
    
    btnButton(0).ButtonStyle = gToolbarsButtonsStyle
    Refresh
End Sub

Private Sub UserControl_Resize()
    Static sIn As Boolean
    
    If sIn Then Exit Sub
    sIn = True
        
    Select Case UserControl.Extender.Align
        Case vbAlignLeft, vbAlignRight
            UserControl.Parent.Controls(Ambient.DisplayName).Align = vbAlignTop
    End Select
    
    If mAlign <> UserControl.Extender.Align Then
        Select Case UserControl.Extender.Align
            Case vbAlignTop, vbAlignRight, vbAlignNone
                Align1 = UserControl.Extender.Align
            Case Else
                Align1 = efnTBAlignTop
        End Select
    End If
    
    If mUsercontrolWidth = -1 Then
        UserControl.Width = 0
    ElseIf mAlign = efnTBAlignNone Then
        If (mUsercontrolHeight <> 0) And (mUsercontrolWidth <> 0) Then
            On Error Resume Next
            If mAutoSize Then
                UserControl.Width = mUsercontrolWidth
            End If
            If mLeaveBorderSpace Then
                UserControl.Height = mUsercontrolHeight + Screen.TwipsPerPixelY * 4
            Else
                UserControl.Height = mUsercontrolHeight
            End If
            On Error GoTo 0
        End If
    Else
        PositionControl
    End If
    RaiseEvent Resize
    sIn = False
End Sub

Private Sub PositionControl()
    UserControl.Height = mUsercontrolHeight + Screen.TwipsPerPixelY * 4
    If mAutoSize Then
        If mUsercontrolWidth > 0 Then
            UserControl.Width = mUsercontrolWidth
        End If
    End If
End Sub

Private Sub UserControl_Show()
    If mRefreshPending Then
        Refresh
    End If
End Sub

Private Sub UserControl_Terminate()
    Set mButtons = Nothing
    If mHwndParent <> 0 Then
        DetachMessage Me, mHwndParent, WM_SIZE
        mHwndParent = 0
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim c As Long
    
    PropBag.WriteProperty "IconsSize", mIconsSize, vxIconsAppDefault
    PropBag.WriteProperty "LeaveBorderSpace", mLeaveBorderSpace, False
    PropBag.WriteProperty "ShowToolTipTextWhenDisabled", mShowToolTipTextWhenDisabled, False
    PropBag.WriteProperty "AutoSize", mAutoSize, True
    
    PropBag.WriteProperty "PicHeight16", mPicHeight16, 16
    PropBag.WriteProperty "PicHeight20", mPicHeight20, 20
    PropBag.WriteProperty "PicHeight24", mPicHeight24, 24
    PropBag.WriteProperty "PicHeight30", mPicHeight30, 30
    PropBag.WriteProperty "PicHeight36", mPicHeight36, 36

    PropBag.WriteProperty "MaskColor", mMaskColor, &HFF00FF
    PropBag.WriteProperty "UseMaskColor", mUseMaskColor, True
    PropBag.WriteProperty "BackColor", mBackColor, vbButtonFace
    PropBag.WriteProperty "DecreaseButtonHeight", mDecreaseButtonHeight, 0
    PropBag.WriteProperty "ButtonsCount", mButtons.Count, 0

    For c = 1 To mButtons.Count
        PropBag.WriteProperty "ButtonChecked" & CStr(c), mButtons(c).Checked, False
        PropBag.WriteProperty "ButtonEnabled" & CStr(c), mButtons(c).Enabled, True
        PropBag.WriteProperty "ButtonKey" & CStr(c), mButtons(c).Key, ""
        PropBag.WriteProperty "ButtonTag" & CStr(c), mButtons(c).Tag, ""
        PropBag.WriteProperty "ButtonPic16" & CStr(c), mButtons(c).Pic16, Nothing
        PropBag.WriteProperty "ButtonPic20" & CStr(c), mButtons(c).Pic20, Nothing
        PropBag.WriteProperty "ButtonPic24" & CStr(c), mButtons(c).Pic24, Nothing
        PropBag.WriteProperty "ButtonPic30" & CStr(c), mButtons(c).Pic30, Nothing
        PropBag.WriteProperty "ButtonPic36" & CStr(c), mButtons(c).Pic36, Nothing
        PropBag.WriteProperty "ButtonPic16Alt" & CStr(c), mButtons(c).Pic16Alt, Nothing
        PropBag.WriteProperty "ButtonPic20Alt" & CStr(c), mButtons(c).Pic20Alt, Nothing
        PropBag.WriteProperty "ButtonPic24Alt" & CStr(c), mButtons(c).Pic24Alt, Nothing
        PropBag.WriteProperty "ButtonPic30Alt" & CStr(c), mButtons(c).Pic30Alt, Nothing
        PropBag.WriteProperty "ButtonPic36Alt" & CStr(c), mButtons(c).Pic36Alt, Nothing
        PropBag.WriteProperty "UseAltPic" & CStr(c), mButtons(c).UseAltPic, False
        PropBag.WriteProperty "ButtonWidth" & CStr(c), mButtons(c).Width, 360
        PropBag.WriteProperty "ButtonStyle" & CStr(c), mButtons(c).Style, vxTBDefault
        PropBag.WriteProperty "ButtonToolTipText" & CStr(c), mButtons(c).ToolTipText, ""
        PropBag.WriteProperty "ButtonVisible" & CStr(c), mButtons(c).Visible, True
        PropBag.WriteProperty "ButtonOrderToHide" & CStr(c), mButtons(c).OrderToHide, 0
    Next c

End Sub


Public Property Let IconsSize(nValue As vbExToolbarDAIconsSizeConstants)
    Dim iHwnd As Long
    
    If nValue <> mIconsSize Then
'        SetWindowRedraw UserControl.Extender.Container.hWnd, False
        PropertyChanged "IconsSize"
        mIconsSize = nValue
        Refresh
        On Error Resume Next
        UserControl.Parent.Refresh
        UserControl.Extender.Container.Refresh
        
        iHwnd = UserControl.Extender.Container.hWnd
        If iHwnd = 0 Then
            iHwnd = UserControl.Parent.hWnd
        End If
        If iHwnd <> 0 Then
            RedrawWindow iHwnd, ByVal 0&, 0&, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_INTERNALPAINT Or RDW_ALLCHILDREN
        End If
        
'        SetWindowRedraw UserControl.Extender.Container.hWnd, True ', True
'        InvalidateRectAsNull UserControl.Extender.Container.hWnd, 0&, 1&
'        UpdateWindow UserControl.Extender.Container.hWnd
'        InvalidateRectAsNull UserControl.Parent.hWnd, 0&, 1&
'        UpdateWindow UserControl.Parent.hWnd
    End If
End Property

Public Property Get IconsSize() As vbExToolbarDAIconsSizeConstants
    IconsSize = mIconsSize
End Property


Public Property Let MaskColor(nColor As OLE_COLOR)
    If nColor <> mMaskColor Then
        mMaskColor = nColor
        PropertyChanged "MaskColor"
        Refresh
    End If
End Property

Public Property Get MaskColor() As OLE_COLOR
    MaskColor = mMaskColor
End Property


Public Property Let BackColor(nColor As OLE_COLOR)
    If nColor <> mBackColor Then
        mBackColor = nColor
        PropertyChanged "BackColor"
        Refresh
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property


Public Property Let UseMaskColor(nValue As Boolean)
    If nValue <> mUseMaskColor Then
        mUseMaskColor = nValue
        PropertyChanged "UseMaskColor"
        Refresh
    End If
End Property

Public Property Get UseMaskColor() As Boolean
    UseMaskColor = mUseMaskColor
End Property


Public Property Let DecreaseButtonHeight(nValue As Long)
    If nValue <> mDecreaseButtonHeight Then
        mDecreaseButtonHeight = nValue
        PropertyChanged "DecreaseButtonHeight"
        Refresh
    End If
End Property

Public Property Get DecreaseButtonHeight() As Long
    DecreaseButtonHeight = mDecreaseButtonHeight
End Property


Public Property Let AvailableWidth(nValue As Long)
    If nValue <> mAvailableWidth Then
        mAvailableWidth = nValue
        PropertyChanged "AvailableWidth"
        
        If (mLastAvailableWidth = -1) Or (mAvailableWidth = -1) Then
            Refresh
        ElseIf (mAvailableWidth < UserControl.Width) Then
            If (Not mNeedToHide) Or (mAvailableWidth - mHiddenAtWidth) >= (mButtonWidth / 4) Then
                Refresh
            End If
        ElseIf (mHiddenButtonsCount > 0) Or mNeedToHide Then
            If Abs(mAvailableWidth - mLastAvailableWidth) >= (mButtonWidth / 4) Then
                Refresh
            End If
        End If
    
    End If
End Property

Public Property Get AvailableWidth() As Long
    AvailableWidth = mAvailableWidth
End Property


Public Property Get Buttons() As ToolBarDAButtons
    Set Buttons = mButtons
End Property

Friend Property Set Buttons(nButtons As ToolBarDAButtons)
    Set mButtons = nButtons
    PropertyChanged "Buttons"
    SetPicHeights
    Refresh
End Property

Private Sub SetPicHeights()
    Dim c As Long
    Dim iButton As ToolBarDAButton
    Dim iTempDC As Long
    Dim iOldObject As Long
    Dim iPicInfo As BITMAP

    If mButtons.Count = 0 Then
        mPicHeight16 = 16
        mPicHeight20 = 20
        mPicHeight24 = 24
        mPicHeight30 = 30
        mPicHeight36 = 36
    Else
        mPicHeight16 = 0
        mPicHeight20 = 0
        mPicHeight24 = 0
        mPicHeight30 = 0
        mPicHeight36 = 0

        For c = 1 To mButtons.Count
            Set iButton = mButtons(c)

            If Not iButton.Pic16 Is Nothing Then
                iTempDC = CreateCompatibleDC(UserControl.hDC)
                iOldObject = SelectObject(iTempDC, iButton.Pic16.Handle)
                GetObjectAPI iButton.Pic16.Handle, Len(iPicInfo), iPicInfo
                If iPicInfo.bmHeight > mPicHeight16 Then
                    mPicHeight16 = iPicInfo.bmHeight
                End If
                SelectObject iTempDC, iOldObject
                DeleteDC iTempDC
            End If
            If Not iButton.Pic20 Is Nothing Then
                iTempDC = CreateCompatibleDC(UserControl.hDC)
                iOldObject = SelectObject(iTempDC, iButton.Pic20.Handle)
                GetObjectAPI iButton.Pic20.Handle, Len(iPicInfo), iPicInfo
                If iPicInfo.bmHeight > mPicHeight20 Then
                    mPicHeight20 = iPicInfo.bmHeight
                End If
                SelectObject iTempDC, iOldObject
                DeleteDC iTempDC
            End If
            If Not iButton.Pic24 Is Nothing Then
                iTempDC = CreateCompatibleDC(UserControl.hDC)
                iOldObject = SelectObject(iTempDC, iButton.Pic24.Handle)
                GetObjectAPI iButton.Pic24.Handle, Len(iPicInfo), iPicInfo
                If iPicInfo.bmHeight > mPicHeight24 Then
                    mPicHeight24 = iPicInfo.bmHeight
                End If
                SelectObject iTempDC, iOldObject
                DeleteDC iTempDC
            End If
            If Not iButton.Pic30 Is Nothing Then
                iTempDC = CreateCompatibleDC(UserControl.hDC)
                iOldObject = SelectObject(iTempDC, iButton.Pic30.Handle)
                GetObjectAPI iButton.Pic30.Handle, Len(iPicInfo), iPicInfo
                If iPicInfo.bmHeight > mPicHeight30 Then
                    mPicHeight30 = iPicInfo.bmHeight
                End If
                SelectObject iTempDC, iOldObject
                DeleteDC iTempDC
            End If
            If Not iButton.Pic36 Is Nothing Then
                iTempDC = CreateCompatibleDC(UserControl.hDC)
                iOldObject = SelectObject(iTempDC, iButton.Pic36.Handle)
                GetObjectAPI iButton.Pic36.Handle, Len(iPicInfo), iPicInfo
                If iPicInfo.bmHeight > mPicHeight36 Then
                    mPicHeight36 = iPicInfo.bmHeight
                End If
                SelectObject iTempDC, iOldObject
                DeleteDC iTempDC
            End If
        Next c

        If mPicHeight16 < 16 Then mPicHeight16 = 16
        If mPicHeight20 < 20 Then mPicHeight20 = 20
        If mPicHeight24 < 24 Then mPicHeight24 = 24
        If mPicHeight30 < 30 Then mPicHeight30 = 30
        If mPicHeight36 < 36 Then mPicHeight36 = 36
    End If

End Sub

Private Function GetButtonsHeight() As Long
'    GetButtonsHeight = GetPicHeight * Screen.TwipsPerPixelY * 1.4
'    GetButtonsHeight = GetPicHeight * Screen.TwipsPerPixelY + 90 * 15 / Screen.TwipsPerPixelY
    GetButtonsHeight = GetPicHeight * Screen.TwipsPerPixelY + GetNormalPicHeight * Screen.TwipsPerPixelY * 0.36
End Function

Private Function GetPicHeight() As Long
    Dim iTx As Single
    Dim iIconsSize As vbExToolbarDAIconsSizeConstants
    
    iTx = Screen.TwipsPerPixelX
    iIconsSize = mIconsSize
    If iIconsSize = vxIconsAppDefault Then
        iIconsSize = gToolbarsDefaultIconsSize
    End If
    
    Select Case True
        Case iTx >= 15 ' 96 DPI
            If iIconsSize = vxIconsSmall Then
                GetPicHeight = mPicHeight16
            ElseIf iIconsSize = vxIconsBig Then
                GetPicHeight = mPicHeight24
            Else
                GetPicHeight = mPicHeight20
            End If
        Case iTx >= 12 ' 120 DPI
            If iIconsSize = vxIconsSmall Then
                GetPicHeight = mPicHeight20
            ElseIf iIconsSize = vxIconsBig Then
                GetPicHeight = mPicHeight30
            Else
                GetPicHeight = mPicHeight24
            End If
        Case iTx >= 10 ' 144 DPI
            If iIconsSize = vxIconsSmall Then
                GetPicHeight = mPicHeight24
            ElseIf iIconsSize = vxIconsBig Then
                GetPicHeight = mPicHeight36
            Else
                GetPicHeight = mPicHeight30
            End If
        Case Else
            If iIconsSize = vxIconsSmall Then
                GetPicHeight = mPicHeight16 * 15 / iTx
            ElseIf iIconsSize = vxIconsBig Then
                GetPicHeight = mPicHeight24 * 15 / iTx
            Else
                GetPicHeight = mPicHeight20 * 15 / iTx
            End If
    End Select
End Function

Private Function GetNormalPicHeight() As Long
    Dim iTx As Single
    Dim iIconsSize As vbExToolbarDAIconsSizeConstants
    
    iTx = Screen.TwipsPerPixelX
    iIconsSize = mIconsSize
    If iIconsSize = vxIconsAppDefault Then
        iIconsSize = gToolbarsDefaultIconsSize
    End If
    
    Select Case True
        Case iTx >= 15 ' 96 DPI
            If iIconsSize = vxIconsSmall Then
                GetNormalPicHeight = 16
            ElseIf iIconsSize = vxIconsBig Then
                GetNormalPicHeight = 24
            Else
                GetNormalPicHeight = 20
            End If
        Case iTx >= 12 ' 120 DPI
            If iIconsSize = vxIconsSmall Then
                GetNormalPicHeight = 20
            ElseIf iIconsSize = vxIconsBig Then
                GetNormalPicHeight = 30
            Else
                GetNormalPicHeight = 24
            End If
        Case iTx >= 10 ' 144 DPI
            If iIconsSize = vxIconsSmall Then
                GetNormalPicHeight = 24
            ElseIf iIconsSize = vxIconsBig Then
                GetNormalPicHeight = 36
            Else
                GetNormalPicHeight = 30
            End If
        Case Else
            If iIconsSize = vxIconsSmall Then
                GetNormalPicHeight = 16 * 15 / iTx
            ElseIf iIconsSize = vxIconsBig Then
                GetNormalPicHeight = 24 * 15 / iTx
            Else
                GetNormalPicHeight = 20 * 15 / iTx
            End If
    End Select
End Function


Public Sub Refresh()
    Dim c As Long
    Dim c2 As Long
    Dim iButton As ToolBarDAButton
    Dim iButtonHeight As Long
    Dim iCurrentLeft As Long
    Dim iIconsSize As vbExToolbarDAIconsSizeConstants
    Dim iWidthToGain As Long
    Dim iWidthGained As Long
    Dim iVisibleButtonsCount As Long
   
    If Not mRedraw Then
        mRefreshPending = True
        Exit Sub
    End If
    mRefreshPending = False
    
    mRefreshing = True
    On Error GoTo TheExit:
    SetWindowRedraw UserControl.hWnd, False
    UserControl.BackColor = mBackColor
    
    iIconsSize = mIconsSize
    If iIconsSize = vxIconsAppDefault Then
        iIconsSize = gToolbarsDefaultIconsSize
    End If
    
    For c = 1 To mButtons.Count
        mButtons(c).SetParentToolBarDAAndButtonControl Nothing, Nothing
    Next c

    iButtonHeight = GetButtonsHeight
    mInterButtonSpaceInPixels = 2 / 16 * GetPicHeight
    mButtonWidth = iButtonHeight + mInterButtonSpaceInPixels * Screen.TwipsPerPixelX
    mVisibleButtonsCount = 0
    
    mHiddenButtonsCount = 0
    For c = 1 To mButtons.Count
        mButtons(c).Hidden = False
    Next c
    mNeedToHide = False
    If mAvailableWidth > -1 Then
    'If False Then
        For c = 1 To mButtons.Count
            Set iButton = mButtons(c)
            If iButton.Visible Then
                Select Case iButton.Style
                    Case vxTBSeparator
                        iCurrentLeft = iCurrentLeft + 120
                        iButton.Width = 120
                    Case vxTBPlaceholder
                        iCurrentLeft = iCurrentLeft + iButton.Width
                        iButton.Width = iButton.Width
                    Case Else
                        iCurrentLeft = iCurrentLeft + mButtonWidth
                        iButton.Width = mButtonWidth
                End Select
            End If
        Next c
        mAllButtonsWidth = iCurrentLeft - mInterButtonSpaceInPixels * Screen.TwipsPerPixelX
        If mAvailableWidth < mAllButtonsWidth Then
            iWidthToGain = mAllButtonsWidth - mAvailableWidth
            iWidthGained = 0
            For c = 1 To mButtons.Count * 3
                For c2 = 1 To mButtons.Count
                    If mButtons(c2).OrderToHide = c Then
                        mButtons(c2).Hidden = True
                        iWidthGained = iWidthGained + mButtons(c2).Width
                        mHiddenButtonsCount = mHiddenButtonsCount + 1
                    End If
                Next c2
                If iWidthGained >= iWidthToGain Then Exit For
            Next c
            If iWidthGained < iWidthToGain Then
                mNeedToHide = True
                mHiddenAtWidth = mAvailableWidth
            End If
        End If
    End If
    If mAvailableWidth < -1 Then
        mNeedToHide = True
        mHiddenAtWidth = mAvailableWidth
    End If
    iVisibleButtonsCount = 0
    For c = 1 To mButtons.Count
        If Not mButtons(c).Hidden Then
            iVisibleButtonsCount = iVisibleButtonsCount + 1
        End If
    Next c
    If iVisibleButtonsCount = 0 Then
        mNeedToHide = True
        mHiddenAtWidth = mAvailableWidth
    End If
    mLastAvailableWidth = mAvailableWidth
    If mNeedToHide Then
        For c = 0 To btnButton.Count - 1
            btnButton(c).Visible = False
        Next c
        For c = 1 To mButtons.Count
            Set iButton = mButtons(c)
            iButton.SetParentToolBarDAAndButtonControl Me, Nothing
        Next c
        UserControl.Width = 0
        mUsercontrolWidth = -1
        GoTo TheExit:
    End If
    For c = 0 To btnButton.Count - 1
        btnButton(c).Redraw = False
    Next c
    
    iCurrentLeft = 0
    c2 = 0
    For c = 1 To mButtons.Count
        Set iButton = mButtons(c)
        iButton.Index = c
        
        If iButton.Visible And Not iButton.Hidden Then
            mVisibleButtonsCount = mVisibleButtonsCount + 1
            iButton.Left = iCurrentLeft
            Select Case iButton.Style
                Case vxTBSeparator
                    iButton.Width = 120
                    iCurrentLeft = iCurrentLeft + 120
                Case vxTBPlaceholder
                    iCurrentLeft = iCurrentLeft + iButton.Width
                    iButton.SetParentToolBarDAAndButtonControl Me, Nothing
                Case Else
                    iButton.Width = mButtonWidth
                    c2 = c2 + 1
                    If c2 > btnButton.Count Then
                        Load btnButton(c2 - 1)
                    End If
                    
                    btnButton(c2 - 1).BackColor = mBackColor
                    btnButton(c2 - 1).ButtonStyle = gToolbarsButtonsStyle
                    btnButton(c2 - 1).MaskColor = mMaskColor
                    btnButton(c2 - 1).UseMaskColor = mUseMaskColor
    
                    btnButton(c2 - 1).Tag = iButton.Index
                    
                    If (mAlign = efnTBAlignNone) And (Not mLeaveBorderSpace) Then
                        btnButton(c2 - 1).Move iCurrentLeft, 0, iButtonHeight, iButtonHeight - mDecreaseButtonHeight
                    Else
                        btnButton(c2 - 1).Move iCurrentLeft, Screen.TwipsPerPixelY * 2, iButtonHeight, iButtonHeight - mDecreaseButtonHeight
                    End If
    
                    btnButton(c2 - 1).CheckBoxMode = (iButton.Style = vxTBCheck) Or (iButton.Style = vxTBButtonGroup)
                    
                    iButton.IconSize = iIconsSize
                    If iButton.UseAltPic Then
                        If iIconsSize = vxIconsSmall Then
                            Set btnButton(c2 - 1).Pic16 = iButton.Pic16Alt
                            Set btnButton(c2 - 1).Pic20 = iButton.Pic20Alt
                            Set btnButton(c2 - 1).Pic24 = iButton.Pic24Alt
                        ElseIf iIconsSize = vxIconsBig Then
                            Set btnButton(c2 - 1).Pic16 = iButton.Pic24Alt
                            Set btnButton(c2 - 1).Pic20 = iButton.Pic30Alt
                            Set btnButton(c2 - 1).Pic24 = iButton.Pic36Alt
                        Else
                            Set btnButton(c2 - 1).Pic16 = iButton.Pic20Alt
                            Set btnButton(c2 - 1).Pic20 = iButton.Pic24Alt
                            Set btnButton(c2 - 1).Pic24 = iButton.Pic30Alt
                        End If
                    Else
                        If iIconsSize = vxIconsSmall Then
                            Set btnButton(c2 - 1).Pic16 = iButton.Pic16
                            Set btnButton(c2 - 1).Pic20 = iButton.Pic20
                            Set btnButton(c2 - 1).Pic24 = iButton.Pic24
                        ElseIf iIconsSize = vxIconsBig Then
                            Set btnButton(c2 - 1).Pic16 = iButton.Pic24
                            Set btnButton(c2 - 1).Pic20 = iButton.Pic30
                            Set btnButton(c2 - 1).Pic24 = iButton.Pic36
                        Else
                            Set btnButton(c2 - 1).Pic16 = iButton.Pic20
                            Set btnButton(c2 - 1).Pic20 = iButton.Pic24
                            Set btnButton(c2 - 1).Pic24 = iButton.Pic30
                        End If
                    End If
                    btnButton(c2 - 1).ToolTipText = iButton.ToolTipText
                    
                    btnButton(c2 - 1).Enabled = iButton.Enabled
                    btnButton(c2 - 1).Visible = True
                    If (iButton.Style = vxTBCheck) Or (iButton.Style = vxTBButtonGroup) Then
                        btnButton(c2 - 1).Value = iButton.Checked
                    End If
                    
                    iButton.SetParentToolBarDAAndButtonControl Me, btnButton(c2 - 1)
                    iCurrentLeft = iCurrentLeft + mButtonWidth
                    mUsercontrolHeight = iButtonHeight - DecreaseButtonHeight
                    UserControl.Height = mUsercontrolHeight - DecreaseButtonHeight
            End Select
        Else
            iButton.SetParentToolBarDAAndButtonControl Me, Nothing
        End If
    Next c
    
    For c = 0 To btnButton.Count - 1
        btnButton(c).Redraw = True
    Next c
    
    SetWindowRedraw UserControl.hWnd, True
    If mAutoHiddden Then
        ShowWindow UserControl.hWnd, SW_SHOW
        mAutoHiddden = False
    End If
    If mButtons.Count > 0 Then
        mUsercontrolWidth = iCurrentLeft - mInterButtonSpaceInPixels * Screen.TwipsPerPixelX
        If mAlign = efnTBAlignNone Then
            If (UserControl.Width < mUsercontrolWidth) Or mAutoSize Then
                UserControl.Width = mUsercontrolWidth
            End If
        End If
    Else
        mUsercontrolWidth = 0
    End If
    If mAvailableWidth = 0 Then
        mAllButtonsWidth = mUsercontrolWidth
    End If
    If c2 > 0 Then
        For c = (c2 + 1) To btnButton.Count
            Unload btnButton(c - 1)
        Next c
    End If
    mRefreshing = False
    
    Exit Sub
TheExit:
    If mNeedToHide Then
        ShowWindow UserControl.hWnd, SW_HIDE
        mAutoHiddden = True
    End If
    SetWindowRedraw UserControl.hWnd, True
    mRefreshing = False
End Sub

Private Function GetButtonByButtonIndex(nIndex As Long) As ToolBarDAButton
    Dim c As Long

    If nIndex = 0 Then Exit Function
    EnsureDrawn
    For c = 1 To mButtons.Count
        If mButtons(c).Index = nIndex Then
            Set GetButtonByButtonIndex = mButtons(c)
            Exit For
        End If
    Next c
    
'    If mRedraw Then
'        UserControl.Refresh
'    End If
End Function

Public Function GetButtonByKey(nKey As String) As ToolBarDAButton
    Dim c As Long
    
    EnsureDrawn
    For c = 1 To mButtons.Count
        If mButtons(c).Key = nKey Then
            Set GetButtonByKey = mButtons(c)
            Exit For
        End If
    Next c
End Function


Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property


Public Property Let Redraw(nValue As Boolean)
    If nValue <> mRedraw Then
        mRedraw = nValue
        If mRedraw Then
            If mRefreshPending Then
                Refresh
            End If
            UserControl.Refresh
            SetWindowRedraw UserControl.hWnd, True
        Else
            SetWindowRedraw UserControl.hWnd, False
        End If
    End If
End Property

Public Property Get Redraw() As Boolean
    Redraw = mRedraw
End Property

Public Property Get Sized() As Boolean
    Sized = Not tmrFirstResize.Enabled
End Property

Public Property Get ButtonWidth() As Long
    ButtonWidth = mButtonWidth
End Property

Public Property Get VisibleButtonsCount() As Long
    EnsureDrawn
    VisibleButtonsCount = mVisibleButtonsCount
End Property

Public Property Get AllButtonsWidth() As Long
    AllButtonsWidth = mAllButtonsWidth
End Property

Public Property Get NeedToHide() As Boolean
    NeedToHide = mNeedToHide
End Property

Friend Sub ClickButton(nIndex As Long)
    Dim iButton As ToolBarDAButton
    Static sInside As Boolean
    Dim c As Long
    Dim iButton2 As ToolBarDAButton
    
    If sInside Then Exit Sub
    sInside = True
    
    If Not mRefreshing Then
        Set iButton = mButtons(nIndex)
        If (iButton.Style = vxTBButtonGroup) Then
            If iButton.Checked Then
                c = iButton.Index
                If c > 0 Then
                    Do
                        c = c - 1
                        If c = 0 Then Exit Do
                        Set iButton2 = mButtons(c)
                        If iButton2.Style <> vxTBButtonGroup Then Exit Do
                        iButton2.Checked = False
                    Loop
                End If
                c = iButton.Index
                If c < mButtons.Count Then
                    Do
                        c = c + 1
                        If c = (mButtons.Count + 1) Then Exit Do
                        Set iButton2 = mButtons(c)
                        If iButton2.Style <> vxTBButtonGroup Then Exit Do
                        iButton2.Checked = False
                    Loop
                End If
            End If
        End If
        RaiseEvent ButtonClick(iButton)
    End If
    sInside = False
End Sub


Private Property Let Align1(nValue As efnToolBarDAAlignConstants)
    If nValue <> mAlign Then
        mAlign = nValue
        PositionControl
        PropertyChanged "Align"
    End If
End Property

Private Property Get Align1() As efnToolBarDAAlignConstants
    Align1 = mAlign
End Property


Public Property Let LeaveBorderSpace(nValue As Boolean)
    If nValue <> mLeaveBorderSpace Then
        mLeaveBorderSpace = nValue
        PropertyChanged "LeaveBorderSpace"
        Refresh
        UserControl_Resize
    End If
End Property

Public Property Get LeaveBorderSpace() As Boolean
    LeaveBorderSpace = mLeaveBorderSpace
End Property


Public Property Let ShowToolTipTextWhenDisabled(nValue As Boolean)
    Dim c As Long
    
    If nValue <> mShowToolTipTextWhenDisabled Then
        mShowToolTipTextWhenDisabled = nValue
        For c = 0 To btnButton.Count - 1
            btnButton(c).ShowToolTipTextWhenDisabled = mShowToolTipTextWhenDisabled
        Next c
        PropertyChanged "ShowToolTipTextWhenDisabled"
    End If
End Property

Public Property Get ShowToolTipTextWhenDisabled() As Boolean
    ShowToolTipTextWhenDisabled = mShowToolTipTextWhenDisabled
End Property


Public Property Let AutoSize(nValue As Boolean)
    If nValue <> mAutoSize Then
        mAutoSize = nValue
        If mAutoSize Then
            EnsureDrawn
            UserControl.Width = mUsercontrolWidth
        End If
        PropertyChanged "AutoSize"
    End If
End Property

Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_MemberFlags = "200"
    AutoSize = mAutoSize
End Property

Friend Sub EnsureDrawn()
    Dim iRedrawDisabled As Boolean
    
    If tmrFirstResize.Enabled Then
        tmrFirstResize_Timer
    End If
    If mRefreshPending Then
        If Not mRedraw Then
            iRedrawDisabled = True
            mRedraw = True
        End If
        Refresh
        If iRedrawDisabled Then
            mRedraw = False
        End If
    End If
End Sub

Friend Sub DoRefresh()
    Refresh
End Sub

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrPreprocess
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Select Case iMsg
        Case WM_SIZE
            PositionControl
    End Select
End Function

