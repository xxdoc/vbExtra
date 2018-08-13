VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cComboHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISubclass

'Private Const CB_GETITEMHEIGHT As Long = &H154
Private Const CB_GETMINVISIBLE As Long = &H1702&

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Declare Function GetFocus Lib "user32" () As Long

'Private Const LOGPIXELSX As Long = 88
'Private Const LOGPIXELSY As Long = 90
'Private Const LF_FACESIZE = 32

'Private Const WM_CTLCOLORLISTBOX As Long = &H134&
Private Const WM_SETREDRAW As Long = &HB&
Private Const WM_DELETEITEM As Long = &H2D&
Private Const WM_NCHITTEST As Long = &H84&
Private Const WM_SETFOCUS As Long = &H7&

Private Declare Function SendMessageLong Lib "user32" Alias _
        "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long

'Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

' Properties
Private mAutoSizeList As Boolean
Private mShowFullTextOnMouseOver As Boolean

Private mComboPtr As Long
Attribute mComboPtr.VB_VarHelpID = -1
Private mSubclassed As Boolean
Private mListCount As Long
Private mComboEditHwnd As Long
Private WithEvents mTimerToolTip As cTimer
Attribute mTimerToolTip.VB_VarHelpID = -1
Private mToolTipIsVisible As Boolean
Private mToolTipPic As PictureBox
Private WithEvents mToolTipLabel As Label
Attribute mToolTipLabel.VB_VarHelpID = -1
Private mComboHwnd As Long
Private mFormParentHwnd As Long
Private mParentUsercontrolControlsPtr As Long
Private mFirstElement As String
Private mMiddleElement As String
Private mLastElement As String
Private mComboStyle As Long
Private mComboListHwnd As Long
Private mProperSizeDropDownWidthCalled As Boolean
Private mParent As Object

Public Sub SetCombo(nCombo As Object, nParent As ComboFn, nParentUsercontrolControlsPtr As Long, nAutoSizeList As Boolean, nShowFullTextOnMouseOver As Boolean)
    Dim iCboInf As COMBOBOXINFO
    
    Set mParent = nParent
    mComboPtr = ObjPtr(nCombo)
    mComboStyle = GetCombo.Style
    mParentUsercontrolControlsPtr = nParentUsercontrolControlsPtr

    mComboHwnd = GetCombo.hWnd
    mFormParentHwnd = GetParentFormHwnd(GetCombo.hWnd)
    
    ' if the ComboFn control is placed inside an Usercontrol, the Ambient.UserMode reports True, but it's still design time
    If GetWindowClassName(mFormParentHwnd) = "IDEOwner" Then Exit Sub
    
    mAutoSizeList = nAutoSizeList
    mShowFullTextOnMouseOver = nShowFullTextOnMouseOver And (mComboStyle <> 1)
    
    iCboInf.cbSize = Len(iCboInf)
    GetComboBoxInfo GetCombo.hWnd, iCboInf
    mComboEditHwnd = iCboInf.hwndEdit
    mComboListHwnd = iCboInf.hwndList
    
    Attach
End Sub

Private Sub AutoSizeDropDownWidth()
    Dim LRet As Long
    Dim iComboWidth As Single
    Dim iDC As Long
    Dim c As Long
    Dim iAuxWidth As Long
    Dim iLargerWidth As Long
    Dim iRc As RECT
    Dim iLOGFONT As LOGFONT
    Dim iFontHandle As Long
    Dim iOldFont As Long
    Dim iMiddleItem As Long
    Dim iText As String
    Dim iP As POINTAPI
    Dim iShowsScrollBar As Boolean
    Dim iMaxItemsWithoutScrollBar As Long
    
    On Error GoTo errorHandler
    
    If TypeName(GetCombo) <> "ComboBox" Then Exit Sub
    
    iDC = CreateCompatibleDC(GetDC(GetCombo.hWnd))
    If iDC = 0 Then Exit Sub
    
    mListCount = GetCombo.ListCount
    If mListCount > 0 Then
        iMiddleItem = mListCount / 2
        mFirstElement = GetCombo.List(0)
        mLastElement = GetCombo.List(GetCombo.ListCount - 1)
        mMiddleElement = GetCombo.List(iMiddleItem)
    Else
        mFirstElement = ""
        mLastElement = ""
        mMiddleElement = ""
        GetWindowRect GetCombo.hWnd, iRc
        LRet = SendMessageLong(GetCombo.hWnd, CB_SETDROPPEDWIDTH, iRc.Right - iRc.Left, 0)
    End If
    If mListCount = 0 Then Exit Sub
    
    iLOGFONT = StdFontToLogFont_Screen(iDC, GetCombo.Font)
    iFontHandle = CreateFontIndirect(iLOGFONT)
    iOldFont = SelectObject(iDC, iFontHandle)
    
    'Get the width of the largest item
    For c = 0 To mListCount - 1
        iText = GetCombo.List(c)
        GetTextExtentPoint32 iDC, iText, Len(iText), iP
    
        'adjust the number added (20 in this case to
        'achieve desired right margin
        iAuxWidth = iP.x
        If (iAuxWidth > iLargerWidth) Then
            iLargerWidth = iAuxWidth
        End If
    Next
    
    iLargerWidth = iLargerWidth + GetSystemMetrics(SM_CXEDGE) * 2
'    iShowsScrollBar = (GetWindowLong(mComboListHwnd, GWL_STYLE) And WS_VSCROLL) = WS_VSCROLL
'    iShowsScrollBar = False
'    If Not iShowsScrollBar Then
        'GetWindowRect mComboListHwnd, iRc2
        'If (iRc2.Bottom - iRc2.Top) > 4 Then
            'If (iRc2.Bottom - iRc2.Top) < (SendMessage(GetCombo.hWnd, CB_GETITEMHEIGHT, 0, ByVal 0) * mListCount) Then
'            ToClipboard SendMessage(GetCombo.hWnd, CB_GETMINVISIBLE, 0, ByVal 0)
    If IsWindows7OrMore Then
        If IsThemed Then
            iMaxItemsWithoutScrollBar = SendMessageLong(GetCombo.hWnd, CB_GETMINVISIBLE, 0&, 0&)
        Else
            iMaxItemsWithoutScrollBar = 30
        End If
    Else
        iMaxItemsWithoutScrollBar = 8
    End If
    
    If mListCount > iMaxItemsWithoutScrollBar Then
        iShowsScrollBar = True
    End If
        'Else
        '    GoTo errorHandler
        'End If
'    End If
    If iShowsScrollBar Then
        iLargerWidth = iLargerWidth + GetSystemMetrics(SM_CXVSCROLL)
    End If
     
    GetWindowRect GetCombo.hWnd, iRc
    iComboWidth = iRc.Right - iRc.Left ' SendMessageLong(GetCombo.hWnd, CB_GETDROPPEDWIDTH, 0, 0)
    
    If iComboWidth > iLargerWidth Then 'current drop-down width is sufficient
        If SendMessageLong(GetCombo.hWnd, CB_GETDROPPEDWIDTH, 0, 0) <> iComboWidth Then
            LRet = SendMessageLong(GetCombo.hWnd, CB_SETDROPPEDWIDTH, iComboWidth, 0)
        End If
        GoTo errorHandler
    End If
     
    'don't allow drop-down width to
    'exceed screen.width
    If (iLargerWidth + iRc.Left) > (Screen.Width \ Screen.TwipsPerPixelX) Then
        iLargerWidth = (Screen.Width \ Screen.TwipsPerPixelX) - iRc.Left
    End If
    
    LRet = SendMessageLong(GetCombo.hWnd, CB_SETDROPPEDWIDTH, iLargerWidth, 0)
    
errorHandler:
    On Error Resume Next
    If iOldFont <> 0 Then
        SelectObject iDC, iOldFont
    End If
    DeleteDC iDC
    DeleteObject iFontHandle
End Sub

Private Sub Attach()
'    AttachMessage Me, GetCombo.hWnd, WM_CTLCOLORLISTBOX
    AttachMessage Me, mComboHwnd, WM_SETREDRAW
    AttachMessage Me, mComboHwnd, WM_DELETEITEM
    AttachMessage Me, mComboHwnd, WM_NCHITTEST
    AttachMessage Me, mComboHwnd, WM_SETFOCUS
    AttachMessage Me, mComboHwnd, WM_LBUTTONDOWN
    AttachMessage Me, mComboHwnd, WM_DESTROY
    AttachMessage Me, mFormParentHwnd, WM_MOVE
    AttachMessage Me, mFormParentHwnd, WM_SIZE
    If mComboHwnd <> mFormParentHwnd Then
        AttachMessage Me, mFormParentHwnd, WM_DESTROY
    End If
    mSubclassed = True
End Sub

Private Sub Detach()
    If mSubclassed Then
'        DetachMessage Me, mComboHwnd, WM_CTLCOLORLISTBOX
        DetachMessage Me, mComboHwnd, WM_SETREDRAW
        DetachMessage Me, mComboHwnd, WM_DELETEITEM
        DetachMessage Me, mComboHwnd, WM_NCHITTEST
        DetachMessage Me, mComboHwnd, WM_SETFOCUS
        DetachMessage Me, mComboHwnd, WM_LBUTTONDOWN
        DetachMessage Me, mComboHwnd, WM_DESTROY
        DetachMessage Me, mFormParentHwnd, WM_MOVE
        DetachMessage Me, mFormParentHwnd, WM_SIZE
        If mComboHwnd <> mFormParentHwnd Then
            DetachMessage Me, mFormParentHwnd, WM_DESTROY
        End If
        mSubclassed = False
    End If
End Sub

Private Sub Class_Terminate()
    Finish
End Sub

Private Sub Finish()
    Detach
    If Not mTimerToolTip Is Nothing Then
        mTimerToolTip.Interval = 0
        Set mTimerToolTip = Nothing
    End If
    If mToolTipIsVisible Then
        RemoveToolTip
    End If
    Set mParent = Nothing
End Sub

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrPreprocess
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Dim iCancel As Boolean
    
    Select Case iMsg
        Case WM_SETREDRAW, WM_DELETEITEM
            If mAutoSizeList Then
                TestWhetherToUpdateDropDownWidth
            End If
'            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
        Case WM_NCHITTEST
            If mAutoSizeList Then
                TestWhetherToUpdateDropDownWidth
            End If
            If mShowFullTextOnMouseOver Then
                TestWhetherToShowToolTip
            End If
'            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
        Case WM_SETFOCUS
            If mAutoSizeList Then
                TestWhetherToUpdateDropDownWidth
            End If
'            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
        Case WM_MOVE, WM_SIZE
            If mProperSizeDropDownWidthCalled Then
                'ProperSizeDropDownWidth
            Else
                mListCount = -1
            End If
            'ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
        Case WM_DESTROY
            Finish
        Case WM_LBUTTONDOWN
            mParent.RaiseEvent_AboutToDropDown GetCombo.Name, iCancel
            If iCancel Then
                bConsume = True
'                ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            End If
    End Select
    
End Function

Private Sub TestWhetherToShowToolTip()
    Dim iHwndFocus As Long
    
    If Not mTimerToolTip Is Nothing Then Exit Sub
    iHwndFocus = GetFocus
    If iHwndFocus = mComboEditHwnd Then Exit Sub
'    If iHwndFocus = mComboHwnd Then Exit Sub
    
    Set mTimerToolTip = New cTimer
    mTimerToolTip.Interval = 500
    
End Sub

Private Sub mTimerToolTip_ThatTime()
    Dim iHwndFocus As Long
    Dim iEndToolTip As Boolean
    Dim iControlTextWidth As Long
    Dim iR As RECT
    Dim iWinStyleInfo As Long
    
    iHwndFocus = GetFocus
    If (iHwndFocus = mComboEditHwnd) Then ' Or (iHwndFocus = GetCombo.hWnd)
        iEndToolTip = True
    Else
        If MouseIsOverControl(GetCombo) Then
            If IsDropped Then
                iEndToolTip = True
            Else
                iControlTextWidth = ControlTextWidth(GetCombo)
                If iControlTextWidth > ControlWidth(GetCombo) Then
                    If Not mToolTipIsVisible Then
                        mToolTipIsVisible = True
                        On Error Resume Next
                        ParentUsercontrolControls.Add "VB.PictureBox", GetCombo.Name & "_" & "ToolTipPic"
                        Set mToolTipPic = ParentUsercontrolControls(GetCombo.Name & "_" & "ToolTipPic")
                        mToolTipPic.Appearance = 0
                        ParentUsercontrolControls.Add "VB.Label", GetCombo.Name & "_" & "ToolTipLabel"
                        Set mToolTipLabel = ParentUsercontrolControls(GetCombo.Name & "_" & "ToolTipLabel")
                        Set mToolTipLabel.Container = mToolTipPic
                        Set mToolTipLabel.Font = GetCombo.Font
                        mToolTipLabel.ForeColor = GetCombo.ForeColor
                        mToolTipPic.BackColor = vbInfoBackground ' &H80000018
                        mToolTipLabel.BackColor = vbInfoBackground ' &H80000018
                        
                        GetWindowRect GetCombo.hWnd, iR
                        If mComboStyle = 2 Then ' DropDownList
                            SetWindowPos mToolTipPic.hWnd, 0, iR.Left + GetSystemMetrics(SM_CXEDGE), iR.Top + GetSystemMetrics(SM_CXEDGE), iControlTextWidth + 5, (iR.Bottom - iR.Top) - GetSystemMetrics(SM_CXEDGE) * 2, SWP_NOACTIVATE 'Or SWP_SHOWWINDOW
                            mToolTipLabel.Top = (GetSystemMetrics(SM_CYEDGE) - GetSystemMetrics(SM_CYBORDER)) * Screen.TwipsPerPixelY
                            mToolTipLabel.Left = (GetSystemMetrics(SM_CXEDGE) - GetSystemMetrics(SM_CXBORDER)) * Screen.TwipsPerPixelX
                        Else
                            SetWindowPos mToolTipPic.hWnd, 0, iR.Left + GetSystemMetrics(SM_CXEDGE), iR.Top + GetSystemMetrics(SM_CXEDGE), iControlTextWidth + 5, (iR.Bottom - iR.Top) - GetSystemMetrics(SM_CXEDGE) * 2, SWP_NOACTIVATE 'Or SWP_SHOWWINDOW
                            mToolTipLabel.Top = (GetSystemMetrics(SM_CYEDGE) - GetSystemMetrics(SM_CYBORDER) * 2) * Screen.TwipsPerPixelY
                            mToolTipLabel.Left = (GetSystemMetrics(SM_CXEDGE) - GetSystemMetrics(SM_CXBORDER) * 2) * Screen.TwipsPerPixelX
                        End If
                        mToolTipLabel.Width = mToolTipPic.Width
                        mToolTipLabel.Caption = GetCombo.Text
                        mToolTipLabel.Visible = True
                        iWinStyleInfo = GetWindowLong(mToolTipPic.hWnd, GWL_EXSTYLE)
                        iWinStyleInfo = iWinStyleInfo Or WS_EX_TOOLWINDOW
                        SetWindowLong mToolTipPic.hWnd, GWL_EXSTYLE, iWinStyleInfo
                        SetParent mToolTipPic.hWnd, 0
                        MySetProp mToolTipPic.hWnd, "TTPic", 1
                        mToolTipPic.ZOrder
                        mToolTipPic.Visible = True
                    End If
                Else
                    iEndToolTip = True
                End If
            End If
        Else
            iEndToolTip = True
        End If
    End If
    
    If iEndToolTip Then
        If Not mTimerToolTip Is Nothing Then
            mTimerToolTip.Interval = 0
            Set mTimerToolTip = Nothing
        End If
        If mToolTipIsVisible Then
            RemoveToolTip
        End If
    End If
End Sub

Private Sub RemoveToolTip()
    On Error Resume Next
    mToolTipPic.Visible = False
    SetParent mToolTipPic.hWnd, GetCombo.Parent.hWnd
    ParentUsercontrolControls.Remove GetCombo.Name & "_" & "ToolTipPic"
    ParentUsercontrolControls.Remove GetCombo.Name & "_" & "ToolTipLabel"
    MyRemoveProp mToolTipPic.hWnd, "TTPic"
    mToolTipIsVisible = False
End Sub

Public Property Get IsDropped() As Boolean
    Dim iR As RECT
    
    If IsWindows2000OrMore Then
        IsDropped = CBool(SendMessage(GetCombo.hWnd, CB_GETDROPPEDSTATE, 0&, 0&))
    Else
        GetWindowRect GetCombo.hWnd, iR
        IsDropped = GetWindowClassName(WindowFromPoint(iR.Left + 1, iR.Bottom + 5)) = "ComboLBox"
        If Not IsDropped Then
            IsDropped = GetWindowClassName(WindowFromPoint(iR.Left + 1, iR.Top - 5)) = "ComboLBox"
        End If
    End If
End Property

Private Property Get ParentUsercontrolControls() As Object
    Dim iTempObj As Object
    
    If mParentUsercontrolControlsPtr = 0 Then Exit Property
    CopyMemory iTempObj, mParentUsercontrolControlsPtr, 4
    Set ParentUsercontrolControls = iTempObj
    CopyMemory iTempObj, 0&, 4
End Property


Private Property Get GetCombo() As Object
    Dim iTempObj As Object
    
    If mComboPtr = 0 Then Exit Property
    CopyMemory iTempObj, mComboPtr, 4
    Set GetCombo = iTempObj
    CopyMemory iTempObj, 0&, 4
End Property

Private Sub mToolTipLabel_Click()
    On Error Resume Next
    RemoveToolTip
    GetCombo.SetFocus
    mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, GetMessageExtraInfo()
    mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, GetMessageExtraInfo()
End Sub


Public Property Get AutoSizeList() As Boolean
    AutoSizeList = mAutoSizeList
End Property

Public Property Let AutoSizeList(nValue As Boolean)
    Dim iRc As RECT
    
    If nValue <> mAutoSizeList Then
        mAutoSizeList = nValue
        If mAutoSizeList Then
            mListCount = -1
        Else
            GetWindowRect GetCombo.hWnd, iRc
            SendMessageLong GetCombo.hWnd, CB_SETDROPPEDWIDTH, iRc.Right - iRc.Left, 0
        End If
    End If
End Property


Public Property Get ShowFullTextOnMouseOver() As Boolean
    ShowFullTextOnMouseOver = mShowFullTextOnMouseOver
End Property

Public Property Let ShowFullTextOnMouseOver(nValue As Boolean)
    If nValue <> mShowFullTextOnMouseOver Then
        mShowFullTextOnMouseOver = nValue
    End If
End Property

Public Sub ProperSizeDropDownWidth()
    AutoSizeDropDownWidth
    mProperSizeDropDownWidthCalled = True
End Sub

Private Sub TestWhetherToUpdateDropDownWidth()
    Dim iUpdate As Boolean
    Dim iMiddleItem As Long
    
    On Error GoTo TheExit:
    
    If mListCount <> GetCombo.ListCount Then
        iUpdate = True
    Else
        If GetCombo.ListCount > 0 Then
            If GetCombo.List(0) <> mFirstElement Then
                iUpdate = True
            Else
                If GetCombo.List(GetCombo.ListCount - 1) <> mLastElement Then
                    iUpdate = True
                Else
                    iMiddleItem = mListCount / 2
                    If GetCombo.List(iMiddleItem) <> mMiddleElement Then
                        iUpdate = True
                    End If
                End If
            End If
        End If
    End If
    
    If iUpdate Then
        AutoSizeDropDownWidth
    End If
    
TheExit:
End Sub


Public Function GetNeededWidth()
    Dim iDC As Long
    Dim c As Long
    Dim iAuxWidth As Long
    Dim iWidth As Long
    Dim iLOGFONT As LOGFONT
    Dim iFontHandle As Long
    Dim iOldFont As Long
    Dim iText As String
    Dim iP As POINTAPI
    Dim iListCount As Long
    
    On Error GoTo errorHandler
    
    If TypeName(GetCombo) <> "ComboBox" Then Exit Function
    
    iDC = CreateCompatibleDC(GetDC(GetCombo.hWnd))
    If iDC = 0 Then Exit Function
    
    iListCount = GetCombo.ListCount
    
    iLOGFONT = StdFontToLogFont_Screen(iDC, GetCombo.Font)
    iFontHandle = CreateFontIndirect(iLOGFONT)
    iOldFont = SelectObject(iDC, iFontHandle)
    
    'Get the width of the largest item
    For c = 0 To iListCount - 1
        iText = GetCombo.List(c)
        GetTextExtentPoint32 iDC, iText, Len(iText), iP
        
        'adjust the number added (20 in this case to
        'achieve desired right margin
        iAuxWidth = iP.x
        If (iAuxWidth > iWidth) Then
            iWidth = iAuxWidth
        End If
    Next
    
    iWidth = iWidth + GetSystemMetrics(SM_CXEDGE) * 8
    If ComboHasDropDownButton(GetCombo) Then
        iWidth = iWidth + GetSystemMetrics(SM_CXVSCROLL) - GetSystemMetrics(SM_CXEDGE)
    End If
    
    GetNeededWidth = iWidth * Screen.TwipsPerPixelX
    
errorHandler:
    On Error Resume Next
    If iOldFont <> 0 Then
        SelectObject iDC, iOldFont
    End If
    DeleteDC iDC
    DeleteObject iFontHandle
End Function

