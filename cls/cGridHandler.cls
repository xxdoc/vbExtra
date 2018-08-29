VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGridHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISubclass

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const EM_GETLINECOUNT = &HBA&

Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Const WM_KILLFOCUS As Long = &H8&
Private Const WM_LBUTTONDBLCLK As Long = &H203

Private Declare Function GetDoubleClickTime Lib "user32" () As Long
Private Const GAKS_KEYDOWN = &H8000

Private Const IDC_HAND = 32649&
Private Const IDC_WAIT = 65557

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

' properties
Private mStretchColumnsWidthsToFill As Boolean
Private mEnableOrderByColumns As Boolean
Private mGridsFlatAppearance As Boolean
Private mThereAreHiddenCols As Boolean
Private mInitialOrderColumn As Long
Private mInitialOrderDescending As Boolean
Private mSameDataGroupedInColumns As Boolean
Private mShowToolTipsOnLongerCellTexts As Boolean
Private mDoNotRememberOrder As Boolean
Private mShowToolTipsForOrderColumns As Boolean
Private mAllowTextEdition As Boolean
Private mTextEditionLocked As Boolean

Private mGrid As Object
Private mGridNameWithParent As String
Private mGridName As String
Private mRememberUserPrintingPreferences As gfnRememberUserPrintingPreferences
Private mReportID As String
Private mIsTypeSupported As Boolean ' if the object is supoported for full features (like to change the columns sizes)

Private mParentFlexFnObjectPtr As Long
Private mGridFormHwnd As Long
Private mGridHwnd As Long
Private mGridSubclassed As Boolean
Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private mGridParent As Object
Private WithEvents mTimerUpdateUI As cTimer
Attribute mTimerUpdateUI.VB_VarHelpID = -1
Private WithEvents mTimerGridUpdated As cTimer
Attribute mTimerGridUpdated.VB_VarHelpID = -1
Private WithEvents mTimerToolTipOnLargeTexts As cTimer
Attribute mTimerToolTipOnLargeTexts.VB_VarHelpID = -1
Private WithEvents mTimerDoNotStoreCellDataStill As cTimer
Attribute mTimerDoNotStoreCellDataStill.VB_VarHelpID = -1
Private WithEvents mTimerSetFocusToEditCellTextBox As cTimer
Attribute mTimerSetFocusToEditCellTextBox.VB_VarHelpID = -1
Private WithEvents mTimerEditCell As cTimer
Attribute mTimerEditCell.VB_VarHelpID = -1
Private WithEvents mTimerCheckScroll As cTimer
Attribute mTimerCheckScroll.VB_VarHelpID = -1
Private WithEvents mTimerCheckRedraw As cTimer
Attribute mTimerCheckRedraw.VB_VarHelpID = -1
Private mGridValuesFromClientStored As Boolean
Private mGridCols As Long
Private mGridRows As Long
Private mColumnsWidthsFromClient() As Long ' Width of columns as came from client program, columns with default width have -1
Private mColumnsWidthsFromClientLiteral() As Long 'Width of columns as came from client program, columns with default width has literal width in Twips
Private mColumnsHiddenByClientProgram() As Boolean 'Whether the columns were set visible (width <> 0) by the client program
Private mColumnsHeaders() As String
Private mColumnsHeadersForOrder() As String
Private mAllColHeadersAreEmpty As Boolean
Private mColumnsWidthsActual() As Long 'Width of columns as Left by the component the last time
Private mColumnsHiddenByComponent() As Boolean
Private mOriginalMergeCells As Long

Private mAuxColWidths() As Long
Private mAuxColsHeaders() As String
Private mAuxColsHeadersForOrder() As String
Private mDefaultColWidth As Long
Private mColsHeadersChanged As Boolean
Private mHandIconHandle As Long

Private mColumnsRightPos() As Long
Private mColumnsRightPosSet As Boolean
Private mFixedHeight As Long
Private mFixedWidth As Long
Private mLastColumnClicked As Long
Private mLastOrderDescending As Boolean
Private mGridWasOrderedByUser As Boolean
Private mLastOrderColumnRetrived As Boolean
Private mLastOrderColumnHeader As String
Private mNeedToUpdateOrderByColumns As Boolean
Private mData00 As String
Private mData01 As String
Private mData10 As String
Private mData11 As String

' for Flat Appearance
Private mpicGridFlatLeft As PictureBox
Private mpicGridFlatRight As PictureBox
Private mpicGridFlatTop As PictureBox
Private mpicGridFlatBottom As PictureBox
Private mpicGridFlatLeft_Name As String
Private mpicGridFlatRight_Name As String
Private mpicGridFlatTop_Name As String
Private mpicGridFlatBottom_Name As String
Private mBorderWidth As Long
Private mBorderColor As Long
Private mGridOriginalBorderStyle As Integer
Private mGridOriginalGridLines As Integer
Private mGridOriginalGridLinesFixed As Integer
Private mGridOriginalBorderStyleSet As Boolean

Private mRowUnderMouse As Long
Private mColUnderMouse As Long
Private mToolTipPic As PictureBox
Private WithEvents mToolTipLabel As Label
Attribute mToolTipLabel.VB_VarHelpID = -1
Private mToolTipIsVisible As Boolean
Private mRowToolTip As Long
Private mColToolTip As Long
Private mCellForeColor() As Long
Private mCellFontIndex() As Long
Private mCellDataStored As Boolean
Private mFontsAttributes() As cFontAttributes
Private mOldCursor As Long
'Private WithEvents mTimerCursorChanged As cTimer
Private mUpdateHiddenCols As Boolean
Private mValuesFromClientStoredSuccesfully As Boolean
Private mScrollBars As Long
Private mTooManyEmptyRowsAtTheEnd As Boolean
Private mTimerGridUpdated_Count As Long
Private mEditingCell As Boolean
Private WithEvents mEditCellTextBox As TextBox
Attribute mEditCellTextBox.VB_VarHelpID = -1
Private mEditCellCol As Long
Private mEditCellRow As Long
Private mEditCellHighLight As Long
Private mEditCellTopRow As Long
Private mEditCellLeftCol As Long
Private mEditCellTextBoxOldParent As Long
Private mUCControls As Object
Private mLastOrderChanged As Boolean
Private mCountNoRedraw As Long
Private mError As Boolean
Private mLastTextToFind As String

Private Enum efnFilterAligment
    efnFilterAligmentHorizontal
    efnFilterAligmentVertical
End Enum


Public Sub SetGrid(nGrid As Object, nParentObject As FlexFnObject, nStretchColumnsWidthsToFill As Boolean, nEnableOrderByColumns As Boolean, nGridsFlatAppearance As Boolean, nRememberUserPrintingPreferences As gfnRememberUserPrintingPreferences, nReportID As String, nIsTypeSupported As Boolean, nInitialOrderColumn As Long, nInitialOrderDescending As Boolean, nSameDataGroupedInColumns As Boolean, nBorderColor As Long, nBorderWidth As Long, nShowToolTipsOnLongerCellTexts As Boolean, nDoNotRememberOrder As Boolean, nShowToolTipsForOrderColumns As Boolean, nAllowTextEdition As Boolean, nTextEditionLocked As Boolean)
    mError = False
    If Not mGrid Is Nothing Then Exit Sub
    
    On Error GoTo SetError
    
    Set mGrid = nGrid
    mRememberUserPrintingPreferences = nRememberUserPrintingPreferences
    mReportID = nReportID
    mGridName = mGrid.Name
    On Error Resume Next
    mGridName = mGridName & "_" & mGrid.Index
    On Error GoTo SetError
    mGridNameWithParent = ControlNameWithParent(mGrid)
    mIsTypeSupported = nIsTypeSupported
    mInitialOrderColumn = nInitialOrderColumn
    mInitialOrderDescending = nInitialOrderDescending
    mSameDataGroupedInColumns = nSameDataGroupedInColumns
    mShowToolTipsOnLongerCellTexts = nShowToolTipsOnLongerCellTexts
    mAllowTextEdition = nAllowTextEdition
    mTextEditionLocked = nTextEditionLocked
    mDoNotRememberOrder = nDoNotRememberOrder
    mShowToolTipsForOrderColumns = nShowToolTipsForOrderColumns
    mScrollBars = mGrid.ScrollBars
    
    mParentFlexFnObjectPtr = ObjPtr(nParentObject)
    mGridFormHwnd = GetParentFormHwnd(nGrid.hWnd)
    mGridHwnd = nGrid.hWnd
    mStretchColumnsWidthsToFill = nStretchColumnsWidthsToFill
    mEnableOrderByColumns = nEnableOrderByColumns
    mGridsFlatAppearance = nGridsFlatAppearance
    mBorderColor = nBorderColor
    mBorderWidth = nBorderWidth
    
    Set mGridParent = mGrid.Parent
    If TypeOf mGrid.Parent Is Form Then
        Set mForm = mGrid.Parent
    End If
    mOriginalMergeCells = mGrid.MergeCells
    
    If mGridsFlatAppearance Then
        SetFlatAppearance
    End If
    
    SubclassGrid
    
    Set mTimerCheckRedraw = New cTimer
    mTimerCheckRedraw.Interval = 100
    
    Exit Sub

SetError:
    mError = True
End Sub

Public Property Get Error() As Boolean
    Error = mError
End Property

Private Property Get Keyname() As String
    Select Case mRememberUserPrintingPreferences
        Case gfnRememberByReportIDIfAvailable
            Keyname = mGridNameWithParent
            If Trim$(mReportID) <> "" Then
                Keyname = Keyname & "_" & Trim$(mReportID)
            End If
        Case gfnRememberByGridName
            Keyname = mGridNameWithParent
        Case gfnRememberGlobally
            Keyname = "Global"
    End Select
    Keyname = Base64Encode(Keyname)
End Property

Private Sub SubclassGrid()
    If Not mGridSubclassed Then
        mGridSubclassed = True
        AttachMessage Me, mGridHwnd, WM_SIZE
        AttachMessage Me, mGridHwnd, WM_MOVE
        AttachMessage Me, mGridHwnd, WM_PAINT
        AttachMessage Me, mGridHwnd, WM_DESTROY
        AttachMessage Me, mGridFormHwnd, WM_DESTROY
        AttachMessage Me, mGridHwnd, WM_MOUSEMOVE
        AttachMessage Me, mGridHwnd, WM_LBUTTONDOWN
        AttachMessage Me, mGridHwnd, WM_LBUTTONDBLCLK
        AttachMessage Me, mGridHwnd, WM_ERASEBKGND
        AttachMessage Me, mGridHwnd, WM_WINDOWPOSCHANGED
    End If
End Sub

Private Sub UnsubclassGrid()
    If mGridSubclassed Then
        mGridSubclassed = False
        DetachMessage Me, mGridHwnd, WM_SIZE
        DetachMessage Me, mGridHwnd, WM_MOVE
        DetachMessage Me, mGridHwnd, WM_PAINT
        DetachMessage Me, mGridHwnd, WM_DESTROY
        DetachMessage Me, mGridFormHwnd, WM_DESTROY
        DetachMessage Me, mGridHwnd, WM_MOUSEMOVE
        DetachMessage Me, mGridHwnd, WM_LBUTTONDOWN
        DetachMessage Me, mGridHwnd, WM_LBUTTONDBLCLK
        DetachMessage Me, mGridHwnd, WM_ERASEBKGND
        DetachMessage Me, mGridHwnd, WM_WINDOWPOSCHANGED
    End If
End Sub

Public Property Get GridFormHwnd() As Long
    GridFormHwnd = mGridFormHwnd
End Property

Public Property Get GridHwnd() As Long
    GridHwnd = mGridHwnd
End Property

Public Property Get Grid() As Object
    Set Grid = mGrid
End Property

Private Property Get ParentFlexFnObject() As FlexFnObject
    Dim iTempObj As Object
    
    If mParentFlexFnObjectPtr = 0 Then Exit Property
    CopyMemory iTempObj, mParentFlexFnObjectPtr, 4
    Set ParentFlexFnObject = iTempObj
    CopyMemory iTempObj, 0&, 4
End Property


Public Property Get StretchColumnsWidthsToFill() As Boolean
    StretchColumnsWidthsToFill = mStretchColumnsWidthsToFill
End Property

Public Property Let StretchColumnsWidthsToFill(nValue As Boolean)
    If nValue <> mStretchColumnsWidthsToFill Then
        mStretchColumnsWidthsToFill = nValue
        If mGridValuesFromClientStored Then
            UpdateGridColumnsWidthsStretched
        End If
    End If
End Property


Public Property Get EnableOrderByColumns() As Boolean
    EnableOrderByColumns = mEnableOrderByColumns
End Property

Public Property Let EnableOrderByColumns(nValue As Boolean)
    If nValue <> mEnableOrderByColumns Then
        mEnableOrderByColumns = nValue
        If mEnableOrderByColumns Then
            If mGridValuesFromClientStored Then
                OrderGridByColumn
            End If
        End If
    End If
End Property


Public Property Get GridsFlatAppearance() As Boolean
    GridsFlatAppearance = mGridsFlatAppearance
End Property

Public Property Let GridsFlatAppearance(nValue As Boolean)
    If nValue <> mGridsFlatAppearance Then
        mGridsFlatAppearance = nValue
        SetFlatAppearance
    End If
End Property


Public Property Get RememberUserPrintingPreferences() As gfnRememberUserPrintingPreferences
    RememberUserPrintingPreferences = mRememberUserPrintingPreferences
End Property

Public Property Let RememberUserPrintingPreferences(nValue As gfnRememberUserPrintingPreferences)
    If nValue <> mRememberUserPrintingPreferences Then
        mRememberUserPrintingPreferences = nValue
    End If
End Property


Public Property Get InitialOrderColumn() As Long
    InitialOrderColumn = mInitialOrderColumn
End Property

Public Property Let InitialOrderColumn(nValue As Long)
    If nValue <> mInitialOrderColumn Then
        mInitialOrderColumn = nValue
        If Not mGridWasOrderedByUser Then
            mLastColumnClicked = nValue
        End If
        If mGridValuesFromClientStored Then
            If mEnableOrderByColumns Then
                OrderGridByColumn
'            Else
'                OrderGridByColumn mInitialOrderColumn, mInitialOrderDescending
            End If
        End If
    End If
End Property


Public Property Get InitialOrderDescending() As Long
    InitialOrderDescending = mInitialOrderDescending
End Property

Public Property Let InitialOrderDescending(nValue As Long)
    If nValue <> mInitialOrderDescending Then
        mInitialOrderDescending = nValue
        If mEnableOrderByColumns Then
            If Not mGridWasOrderedByUser Then
                mLastOrderDescending = nValue
            End If
            If mGridValuesFromClientStored Then
                OrderGridByColumn
            End If
'        Else
'            OrderGridByColumn mInitialOrderColumn, mInitialOrderDescending
        End If
    End If
End Property


Public Property Get SameDataGroupedInColumns() As Boolean
    SameDataGroupedInColumns = mSameDataGroupedInColumns
End Property

Public Property Let SameDataGroupedInColumns(nValue As Boolean)
    If nValue <> mSameDataGroupedInColumns Then
        mSameDataGroupedInColumns = nValue
        UpdateDataGrouped
    End If
End Property


Public Property Get ShowToolTipsOnLongerCellTexts() As Long
    ShowToolTipsOnLongerCellTexts = mShowToolTipsOnLongerCellTexts
End Property

Public Property Let ShowToolTipsOnLongerCellTexts(nValue As Long)
    If nValue <> mShowToolTipsOnLongerCellTexts Then
        mShowToolTipsOnLongerCellTexts = nValue
        If mShowToolTipsOnLongerCellTexts Then
            StoreCellsData
        End If
    End If
End Property


Public Property Get AllowTextEdition() As Long
    AllowTextEdition = mAllowTextEdition
End Property

Public Property Let AllowTextEdition(nValue As Long)
    If nValue <> mAllowTextEdition Then
        mAllowTextEdition = nValue
    End If
End Property


Public Property Get TextEditionLocked() As Long
    TextEditionLocked = mTextEditionLocked
End Property

Public Property Let TextEditionLocked(nValue As Long)
    If nValue <> mTextEditionLocked Then
        mTextEditionLocked = nValue
    End If
End Property


Public Property Get DoNotRememberOrder() As Long
    DoNotRememberOrder = mDoNotRememberOrder
End Property

Public Property Let DoNotRememberOrder(nValue As Long)
    If nValue <> mDoNotRememberOrder Then
        mDoNotRememberOrder = nValue
'        If mDoNotRememberOrder Then
            
'        End If
    End If
End Property


Public Property Get ShowToolTipsForOrderColumns() As Long
    ShowToolTipsForOrderColumns = mShowToolTipsForOrderColumns
End Property

Public Property Let ShowToolTipsForOrderColumns(nValue As Long)
    If nValue <> mShowToolTipsForOrderColumns Then
        mShowToolTipsForOrderColumns = nValue
    End If
End Property


Public Property Get ReportID() As Boolean
    ReportID = mReportID
End Property

Public Property Let ReportID(nValue As Boolean)
    If nValue <> mReportID Then
        mReportID = nValue
    End If
End Property


Private Sub Class_Initialize()
    mHandIconHandle = LoadCursor(ByVal 0&, IDC_HAND)
    mLastColumnClicked = -1
End Sub

Private Sub Class_Terminate()
    DestroyCursor mHandIconHandle
    Finish
End Sub

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrConsume
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Static sLastWidth As Long
    Dim iRect As RECT
    Dim iAuxWidth As Long
    Dim x As Long
    Dim y As Long
    Dim iCol As Long
    Dim iMouseIsOnHeaders As Boolean
    Dim iStr As String
    Dim iIgnore As Boolean
    Dim iAuxBool As Boolean
    Dim iShowTip As Boolean
    Dim iMouseDown As Boolean
    Dim iGridNeedsToBeUpdated As Boolean
    
    Static sLastY As Long
    Static sLastX As Long
    Static sLastHwndTTT As Long
    Static sMouseOverGrid As Boolean
    
    If Not mGridSubclassed Then
        ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
        Exit Function
    End If
    
    'Debug.Print mGridName, GetMessageName(iMsg)
    
    iMouseDown = GetAsyncKeyState(vbLeftButton) And GAKS_KEYDOWN
    
    If iMouseDown Then
        If Not sMouseOverGrid Then
            sMouseOverGrid = MouseIsOverGrid
        End If
        If sMouseOverGrid Or MouseIsOverForm Then
            If Not ((iMsg = WM_LBUTTONDOWN) Or (iMsg = WM_SIZE) Or (iMsg = WM_MOVE) Or (iMsg = WM_DESTROY) Or (iMsg = WM_ERASEBKGND) Or (iMsg = WM_WINDOWPOSCHANGED) Or (iMsg = WM_KILLFOCUS) Or (iMsg = WM_LBUTTONDBLCLK)) Then
                ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
                Exit Function
            End If
        End If
    Else
        sMouseOverGrid = False
    End If
    
'    Debug.Print mGridName, GetMessageName(iMsg), "  " & iMouseDown, sMouseOverGrid, MouseIsOverGrid
    
    If mEnableOrderByColumns Or mShowToolTipsOnLongerCellTexts Then
        If (iMsg = WM_MOUSEMOVE) Or (iMsg = WM_LBUTTONDOWN) Then
            x = lParam And &HFFFF&
            y = lParam \ &H10000
            iMouseIsOnHeaders = y < mFixedHeight
        End If
    End If
    
    Select Case iMsg
        Case WM_MOUSEMOVE
            If mEnableOrderByColumns Then
                If iMouseIsOnHeaders Then
                    iCol = GetColumnByPosX(x)
                    iShowTip = False
                    If iCol > -1 Then
                        If (mColumnsHeadersForOrder(iCol) <> "  ") Then
                            If (mColumnsHeadersForOrder(iCol) <> "") Or mAllColHeadersAreEmpty Then
                                iShowTip = True
                            End If
                        End If
                    End If
                    If GetCursor <> IDC_WAIT Then
                        If iCol > -1 Then
                            If (mColumnsHeadersForOrder(iCol) <> "  ") Then
                                If GetCursor <> mHandIconHandle Then
                                    mOldCursor = SetCursor(mHandIconHandle)
                                End If
                            End If
                            If (y = sLastY) And (x = sLastX) Then
                                iIgnore = True
                            End If
                            sLastY = y
                            sLastX = x
                            If iShowTip Then
                                If mShowToolTipsForOrderColumns Then
                                    If mAllColHeadersAreEmpty Then
                                        iStr = GetLocalizedString(efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString1)
                                    Else
                                        If Trim$(mColumnsHeadersForOrder(iCol)) = "" Then
                                            iStr = GetLocalizedString(efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString1)
                                        Else
                                            iStr = GetLocalizedString(efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString2) & " " & mColumnsHeadersForOrder(iCol)
                                        End If
                                    End If
                                    If mLastColumnClicked = iCol Then
                                        If mLastOrderDescending Xor mLastOrderChanged Then
                                            iStr = iStr & " (" & GetLocalizedString(efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString3) & ")"
                                        Else
                                            iStr = iStr & " (" & GetLocalizedString(efnGUIStr_cGridHandler_ISubclass_Windowproc_OrderByColumnsString4) & ")"
                                        End If
                                    End If
                                    If mGrid.ToolTipText <> iStr Then
                                        If iIgnore Then
                                            mGrid.ToolTipText = ""
                                            sLastHwndTTT = 0
                                        Else
                                            mGrid.ToolTipText = iStr
                                            sLastHwndTTT = hWnd
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If GetCursor <> IDC_WAIT Then
                            If mOldCursor = 0 Then
                                mGrid.MousePointer = vbDefault
                            Else
                                SetCursor mOldCursor
                            End If
                            If sLastHwndTTT = hWnd Then
                                sLastHwndTTT = 0
                                mGrid.ToolTipText = ""
                            End If
                        End If
                    End If
                Else
                    If GetCursor <> IDC_WAIT Then
                        If mOldCursor = 0 Then
                            mGrid.MousePointer = vbDefault
                        Else
                            SetCursor mOldCursor
                        End If
                        If sLastHwndTTT = hWnd Then
                            sLastHwndTTT = 0
                            mGrid.ToolTipText = ""
                        End If
                    End If
                End If
            End If
            If mShowToolTipsOnLongerCellTexts Then
                If Not iMouseIsOnHeaders Then
                    If mTimerToolTipOnLargeTexts Is Nothing Then
                        Set mTimerToolTipOnLargeTexts = New cTimer
                    Else
                        mTimerToolTipOnLargeTexts.Interval = 0
                    End If
                    mTimerToolTipOnLargeTexts.Interval = 1000
                    GetRowColUnderMouse mRowUnderMouse, mColUnderMouse
                End If
            End If
        Case WM_LBUTTONDOWN
            If mEnableOrderByColumns Then
                If iMouseIsOnHeaders Then
                    If GetCursor <> IDC_WAIT Then
                        iCol = GetColumnByPosX(x)
                        If iCol > -1 Then
                            If (mColumnsHeadersForOrder(iCol) <> "  ") Then
                                ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
                                OrderGridByColumn iCol
                                mGrid.ToolTipText = ""
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
            
            If mAllowTextEdition Then
                If mGridValuesFromClientStored And Not mNeedToUpdateOrderByColumns Then
                    If Not mEditingCell Then
                        If mTimerEditCell Is Nothing Then
                            If mGrid.SelectionMode = flexSelectionFree Then
                                If (mGrid.Row = mGrid.MouseRow) Or (mGrid.HighLight = flexHighlightNever) Then
                                    If (mGrid.col = mGrid.MouseCol) Or (mGrid.HighLight = flexHighlightNever) Then
                                        If mTimerEditCell Is Nothing Then
                                            If mGrid.Row = mGrid.MouseRow Then
                                                mEditCellRow = mGrid.MouseRow
                                            Else
                                                mEditCellRow = -1
                                            End If
                                            If mGrid.col = mGrid.MouseCol Then
                                                mEditCellCol = mGrid.MouseCol
                                            Else
                                                mEditCellCol = -1
                                            End If
                                            mEditCellHighLight = mGrid.HighLight
                                            Set mTimerEditCell = New cTimer
                                            mTimerEditCell.Interval = GetDoubleClickTime
                                        End If
                                    End If
                                End If
                            ElseIf mGrid.SelectionMode = flexSelectionByRow Then
                                If (mGrid.Row = mGrid.MouseRow) Or (mGrid.HighLight = flexHighlightNever) Then
                                    If mTimerEditCell Is Nothing Then
                                        If mGrid.Row = mGrid.MouseRow Then
                                            mEditCellRow = mGrid.MouseRow
                                        Else
                                            mEditCellRow = -1
                                        End If
                                        If mGrid.col = mGrid.MouseCol Then
                                            mEditCellCol = mGrid.MouseCol
                                        Else
                                            mEditCellCol = -1
                                        End If
                                        mEditCellHighLight = mGrid.HighLight
                                        Set mTimerEditCell = New cTimer
                                        mTimerEditCell.Interval = GetDoubleClickTime
                                    End If
                                End If
                            Else
                                If (mGrid.col = mGrid.MouseCol) Or (mGrid.HighLight = flexHighlightNever) Then
                                    If mTimerEditCell Is Nothing Then
                                        If mGrid.Row = mGrid.MouseRow Then
                                            mEditCellRow = mGrid.MouseRow
                                        Else
                                            mEditCellRow = -1
                                        End If
                                        If mGrid.col = mGrid.MouseCol Then
                                            mEditCellCol = mGrid.MouseCol
                                        Else
                                            mEditCellCol = -1
                                        End If
                                        mEditCellHighLight = mGrid.HighLight
                                        Set mTimerEditCell = New cTimer
                                        mTimerEditCell.Interval = GetDoubleClickTime
                                    End If
                                End If
                            End If
                        Else ' doubleclick (it doesn't work)
                            mTimerEditCell.Interval = 0
                            Set mTimerEditCell = Nothing
                        End If
                    Else
                        If (mGrid.Row <> mGrid.MouseRow) Or (mGrid.col <> mGrid.MouseCol) Then
                            RemoveEditCell
                        End If
                    End If
                End If
            End If
        Case WM_LBUTTONDBLCLK
            If mAllowTextEdition Then
                If mGridValuesFromClientStored And Not mNeedToUpdateOrderByColumns Then
                    If Not mEditingCell Then
                        If Not mTimerEditCell Is Nothing Then
                            mTimerEditCell.Interval = 0
                            Set mTimerEditCell = Nothing
                        End If
                    End If
                End If
            End If
        Case WM_SIZE, WM_MOVE
            If mGridsFlatAppearance Then
                PositionBorderPicBoxes
            End If
            If mGridValuesFromClientStored Then
                If mStretchColumnsWidthsToFill Then
                    GetWindowRect hWnd, iRect
                    iAuxWidth = iRect.Right - iRect.Left
                    If sLastWidth <> iAuxWidth Then
                    sLastWidth = iAuxWidth
                        UpdateGridColumnsWidthsStretched
                    End If
                End If
            End If
        Case WM_PAINT
            'mShape.Visible = IsWindowVisible(hWnd) <> 0
            If mTimerUpdateUI Is Nothing Then
                Set mTimerUpdateUI = New cTimer
                mTimerUpdateUI.Interval = 1
            Else
                If mTimerUpdateUI.Interval = 1 Then
                    mTimerUpdateUI.Interval = 0
                    Set mTimerUpdateUI = Nothing
                    Set mTimerUpdateUI = New cTimer
                    mTimerUpdateUI.Interval = 1
                End If
            End If
            If mTimerGridUpdated Is Nothing Then
                Set mTimerGridUpdated = New cTimer
                mTimerGridUpdated.Interval = 50
                CheckWhetherGridNeedsUpdate
                ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
                Exit Function
            End If
            If mTimerGridUpdated.Tag = "2" Then
                mTimerGridUpdated.Interval = 0
                Set mTimerGridUpdated = Nothing
            Else
                If mEnableOrderByColumns Then
                    If Not mDoNotRememberOrder Then
                        If mGrid.Redraw And mGrid.Visible Then
                            iGridNeedsToBeUpdated = GridNeedsToBeUpdated
                            If iGridNeedsToBeUpdated Then
                                If mGridWasOrderedByUser Then
                                    mTimerGridUpdated.Tag = "1"
                                    mGrid.Redraw = False
                                Else
                                    If iGridNeedsToBeUpdated Then
                                        mTimerGridUpdated.Tag = "1"
                                        mGrid.Redraw = False
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                mTimerGridUpdated.Interval = 50
            End If
            If mGridValuesFromClientStored Then
                If mStretchColumnsWidthsToFill Then
                    GetWindowRect hWnd, iRect
                    iAuxWidth = iRect.Right - iRect.Left
                    If sLastWidth <> iAuxWidth Then
                        sLastWidth = iAuxWidth
                        UpdateGridColumnsWidthsStretched
                    End If
                End If
            End If
        Case WM_ERASEBKGND
            If mGridsFlatAppearance Then
                iAuxBool = IsWindowVisible(hWnd) <> 0
                mpicGridFlatLeft.Visible = iAuxBool
                mpicGridFlatRight.Visible = iAuxBool
                mpicGridFlatTop.Visible = iAuxBool
                mpicGridFlatBottom.Visible = iAuxBool
            End If
            If mEditingCell Then
                If mTimerCheckScroll Is Nothing Then
                    Set mTimerCheckScroll = New cTimer
                    mTimerCheckScroll.Interval = 1
                End If
            End If
        Case WM_WINDOWPOSCHANGED
            If mGridsFlatAppearance Then
                ShowHideBorder
            End If
        Case WM_DESTROY
            Finish
        Case WM_KILLFOCUS
            If mEditingCell Then
                RemoveEditCell
            End If
    End Select
    ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
End Function

Private Sub mEditCellTextBox_DblClick()
    RemoveEditCell
    DoEvents
    Sleep 1
    mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, GetMessageExtraInfo()
    Sleep 1
    mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, GetMessageExtraInfo()
    If Not mTimerEditCell Is Nothing Then
        mTimerEditCell.Interval = 0
        Set mTimerEditCell = Nothing
    End If
    Sleep 100
    DoEvents
    Sleep 1
    mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, GetMessageExtraInfo()
    Sleep 1
    mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, GetMessageExtraInfo()
    If Not mTimerEditCell Is Nothing Then
        mTimerEditCell.Interval = 0
        Set mTimerEditCell = Nothing
    End If
    Sleep 100
    DoEvents
    Sleep 1
    mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, GetMessageExtraInfo()
    Sleep 1
    mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, GetMessageExtraInfo()

End Sub

Private Sub mEditCellTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        RemoveEditCell
    End If
End Sub

Private Sub mForm_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Finish
    End If
End Sub

Private Sub mTimerCheckRedraw_ThatTime()
    Dim iDoCount As Boolean
    
    On Error GoTo TheExit:
    If IsWindow(mGridHwnd) <> 0 Then
        If IsWindowVisible(mGridHwnd) <> 0 Then
            If Not mGrid Is Nothing Then
                If mGrid.Redraw = False Then
                    iDoCount = True
                End If
            End If
        End If
    End If
    
    If iDoCount Then
        mCountNoRedraw = mCountNoRedraw + 1
        If mCountNoRedraw > 10 Then
            mGrid.Redraw = True
            mCountNoRedraw = 0
        End If
    Else
        mCountNoRedraw = 0
    End If
    
TheExit:

End Sub

Private Sub mTimerCheckScroll_ThatTime()
    mTimerCheckScroll.Interval = 0
    Set mTimerCheckScroll = Nothing
    If (mGrid.TopRow <> mEditCellTopRow) Or (mGrid.LeftCol <> mEditCellLeftCol) Then
        RemoveEditCell
    End If
End Sub

Private Sub mTimerSetFocusToEditCellTextBox_ThatTime()
    If Not mEditCellTextBox Is Nothing Then
        If GetForegroundWindow = mGridFormHwnd Then
            SetFocusAPI mEditCellTextBox.hWnd
        End If
    End If
    mTimerSetFocusToEditCellTextBox.Interval = 0
    Set mTimerSetFocusToEditCellTextBox = Nothing
End Sub

'Private Sub mTimerCursorChanged_ThatTime()
'
'    If GetCursor <> mHandIconHandle Then
'        mTimerCursorChanged.Interval = 0
'        Set mTimerCursorChanged = Nothing
'    Else
'        If Not MouseIsOverControl(mGrid) Then
'            If GetCursor <> IDC_WAIT Then
'                mTimerCursorChanged.Interval = 0
'                Set mTimerCursorChanged = Nothing
'                If mOldCursor = 0 Then
'                    mGrid.MousePointer = vbDefault
'                Else
'                    SetCursor mOldCursor
'                End If
'            End If
'        End If
'    End If
'End Sub

Private Sub mTimerDoNotStoreCellDataStill_ThatTime()
    mTimerDoNotStoreCellDataStill.Interval = 0
    Set mTimerDoNotStoreCellDataStill = Nothing
    If Not mCellDataStored Then
        StoreCellsData
    End If
End Sub

Private Sub mTimerEditCell_ThatTime()
    Dim iDoit As Boolean
    
'    If mGrid.HighLight = mEditCellHighLight Then
'        If mGrid.SelectionMode = flexSelectionByRow Then
'            If mGrid.MouseCol = mEditCellCol Then
'                iDoit = True
'            End If
'        ElseIf mGrid.SelectionMode = flexSelectionByColumn Then
'            If mGrid.MouseRow = mEditCellRow Then
'                iDoit = True
'            End If
'        Else
'        Debug.Print mGrid.MouseCol, mEditCellCol, mGrid.MouseRow, mEditCellRow
        If (mGrid.MouseCol = mEditCellCol) And (mGrid.MouseRow = mEditCellRow) Then
            iDoit = True
        End If
'        End If
'    End If
    mTimerEditCell.Interval = 0
    Set mTimerEditCell = Nothing
    If iDoit Then
        EditCell
    End If
End Sub

Private Sub mTimerToolTipOnLargeTexts_ThatTime()
    Dim iRowUnderMouse As Long
    Dim iColUnderMouse As Long
    
    If (mRowUnderMouse = -1) Or (mColUnderMouse = -1) Then
        If mToolTipIsVisible Then
            RemoveToolTip
        End If
        mTimerToolTipOnLargeTexts.Interval = 0
        Set mTimerToolTipOnLargeTexts = Nothing
        Exit Sub
    End If
    
    GetRowColUnderMouse iRowUnderMouse, iColUnderMouse
    If (iRowUnderMouse = mRowUnderMouse) And (iColUnderMouse = mColUnderMouse) Then
        If mToolTipIsVisible Then
            If (mRowToolTip <> mRowUnderMouse) Or (mColToolTip <> mColUnderMouse) Then
                RemoveToolTip
            Else
                mTimerToolTipOnLargeTexts.Interval = 50
                Exit Sub
            End If
        End If
        TestWhetherToShowToolTip
        If mToolTipIsVisible Then
            If mTimerToolTipOnLargeTexts Is Nothing Then
                RemoveToolTip
            Else
                mTimerToolTipOnLargeTexts.Interval = 50
            End If
        Else
            If mTimerToolTipOnLargeTexts Is Nothing Then Exit Sub
            mTimerToolTipOnLargeTexts.Interval = 0
            Set mTimerToolTipOnLargeTexts = Nothing
        End If
    Else
        If mToolTipIsVisible Then
            RemoveToolTip
        End If
        mTimerToolTipOnLargeTexts.Interval = 0
        Set mTimerToolTipOnLargeTexts = Nothing
        mRowUnderMouse = iRowUnderMouse
        mColUnderMouse = iColUnderMouse
    End If
End Sub

Private Sub mToolTipLabel_Click()
    On Error Resume Next
    RemoveToolTip
    mGrid.SetFocus
    mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, GetMessageExtraInfo()
    mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, GetMessageExtraInfo()
    If Not mTimerToolTipOnLargeTexts Is Nothing Then
        mTimerToolTipOnLargeTexts.Interval = 2000
    End If
End Sub

Private Sub mTimerGridUpdated_ThatTime()
    mTimerGridUpdated_Count = mTimerGridUpdated_Count + 1
    If mTimerGridUpdated_Count >= 10 Then
        If mTimerGridUpdated.Tag = "1" Then
            mTimerGridUpdated.Tag = "2"
            mGrid.Redraw = True
            mGrid.Refresh
        End If
        If Not mTimerGridUpdated Is Nothing Then
            mTimerGridUpdated.Interval = 0
        End If
        mTimerGridUpdated_Count = 0
    End If
  '  Set mTimerGridUpdated = Nothing
    CheckWhetherGridNeedsUpdate
End Sub

Private Sub mTimerUpdateUI_ThatTime()
    If (mTimerUpdateUI.Interval <> 20) And mGridSubclassed Then
        mTimerUpdateUI.Interval = 20
        ParentFlexFnObject.NotifyUIToBeUpdated
    Else
        mTimerUpdateUI.Interval = 0
        Set mTimerUpdateUI = Nothing
    End If
End Sub

Private Sub Finish()
    UnsubclassGrid
    If Not mTimerUpdateUI Is Nothing Then
        mTimerUpdateUI.Interval = 0
        Set mTimerUpdateUI = Nothing
    End If
    If mToolTipIsVisible Then
        RemoveToolTip
    End If
    If mEditingCell Then
        RemoveEditCell
    End If
    If Not mTimerGridUpdated Is Nothing Then
        mTimerGridUpdated.Interval = 0
        Set mTimerGridUpdated = Nothing
    End If
    If Not mTimerToolTipOnLargeTexts Is Nothing Then
        mTimerToolTipOnLargeTexts.Interval = 0
        Set mTimerToolTipOnLargeTexts = Nothing
    End If
    If Not mTimerDoNotStoreCellDataStill Is Nothing Then
        mTimerDoNotStoreCellDataStill.Interval = 0
        Set mTimerDoNotStoreCellDataStill = Nothing
    End If
    If Not mTimerSetFocusToEditCellTextBox Is Nothing Then
        mTimerSetFocusToEditCellTextBox.Interval = 0
        Set mTimerSetFocusToEditCellTextBox = Nothing
    End If
    If Not mTimerEditCell Is Nothing Then
        mTimerEditCell.Interval = 0
        Set mTimerEditCell = Nothing
    End If
    If Not mTimerCheckScroll Is Nothing Then
        mTimerCheckScroll.Interval = 0
        Set mTimerCheckScroll = Nothing
    End If
    If Not mTimerCheckRedraw Is Nothing Then
        mTimerCheckRedraw.Interval = 0
        Set mTimerCheckRedraw = Nothing
    End If
    If Not mUCControls Is Nothing Then
        Set mUCControls = Nothing
    End If
    Set mForm = Nothing
    Set mGridParent = Nothing
End Sub

Private Sub StoreGridValuesFromClient()
    Dim c As Long
    Dim iLng As Long
    Dim r As Long
    Dim iStr As String
    
    mGridCols = mGrid.Cols
    mGridRows = mGrid.Rows
    mEditCellHighLight = mGrid.HighLight
    
    'If mGridRows > 20000 Then Exit Sub
    If mEditingCell Then
        RemoveEditCell
    End If
    
    If (mGrid.FixedCols > mGrid.Cols) And (mGrid.FixedRows > mGrid.Rows) Then
        mData00 = mGrid.TextMatrix(mGrid.FixedRows, mGrid.FixedCols)
        mData01 = mGrid.TextMatrix(mGrid.FixedRows, mGrid.Cols - 1)
        mData10 = mGrid.TextMatrix(mGrid.Rows - 1, mGrid.FixedCols)
        mData11 = mGrid.TextMatrix(mGrid.Rows - 1, mGrid.Cols - 1)
    End If
    
    ReDim Preserve mColumnsWidthsFromClient(mGridCols - 1)
    ReDim Preserve mColumnsWidthsFromClientLiteral(mGridCols - 1)
    ReDim mColumnsHiddenByClientProgram(mGridCols - 1)
    ReDim mColumnsHeaders(mGridCols - 1)
    ReDim mColumnsHeadersForOrder(mGridCols - 1)
    
    ReDim Preserve mColumnsWidthsActual(mGridCols - 1)
    ReDim mColumnsHiddenByComponent(mGridCols - 1)
    
    ReDim mColumnsRightPos(mGridCols - 1)
    
    SetAuxGridColumsWidths
    SetAuxGridColumsHeaders True
    For c = 0 To mGridCols - 1
        If mColumnsWidthsActual(c) <> mGrid.ColWidth(c) Then
            mColumnsWidthsFromClientLiteral(c) = mAuxColWidths(c)
            mColumnsWidthsFromClient(c) = mGrid.ColWidth(c)
        End If
        mColumnsHiddenByClientProgram(c) = mColumnsWidthsFromClient(c) = 0
        mColumnsHeaders(c) = mAuxColsHeaders(c)
        mColumnsHeadersForOrder(c) = mAuxColsHeadersForOrder(c)
        mColumnsWidthsActual(c) = mColumnsWidthsFromClientLiteral(c)
    Next c

    iLng = 0
    For r = 0 To mGrid.FixedRows - 1
        iLng = iLng + mGrid.RowHeight(r)
    Next r
    
    On Error Resume Next
    Err.Clear
    mFixedHeight = mGrid.Parent.ScaleY(iLng, vbTwips, vbPixels)
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    
    iLng = 0
    For c = 0 To mGrid.FixedCols - 1
        iLng = iLng + mGrid.ColWidth(c)
    Next c
    mFixedWidth = mGrid.Parent.ScaleX(iLng, vbTwips, vbPixels)
    
    mColumnsRightPosSet = False
    
    If Not mDoNotRememberOrder Then
        If Not mLastOrderColumnRetrived Then
            iStr = Base64Decode(GetSetting(AppNameForRegistry, "Preferences", Keyname & "_OrderByCol", ""))
            If iStr = "" Then
                mLastColumnClicked = mInitialOrderColumn
            Else
                If Len(iStr) > 2 Then
                    iLng = GetColIndexByColHeader(Left$(iStr, Len(iStr) - 2))
                    If iLng > -1 Then
                        mLastColumnClicked = iLng
                        mLastOrderDescending = (Right$(iStr, 1) = "D")
                        mGridWasOrderedByUser = True
                    End If
                Else
                    mLastColumnClicked = mInitialOrderColumn
                End If
            End If
            mLastOrderColumnRetrived = True
            If mLastColumnClicked > -1 Then
                mLastOrderColumnHeader = mColumnsHeaders(mLastColumnClicked)
            End If
        Else
            If EnableOrderByColumns Then
                SetOrderByLastColHeader
            End If
        End If
    Else
        If mEnableOrderByColumns Then
            If mInitialOrderColumn > -1 Then
                OrderGridByColumn mInitialOrderColumn, mInitialOrderDescending, True
            End If
        End If
    End If
    
    mCellDataStored = False
    
    If mShowToolTipsOnLongerCellTexts Then
        If mTimerDoNotStoreCellDataStill Is Nothing Then
            Set mTimerDoNotStoreCellDataStill = New cTimer
            mTimerDoNotStoreCellDataStill.Interval = 1000
            StoreCellsData
        Else
            mTimerDoNotStoreCellDataStill.Interval = 0
            Set mTimerDoNotStoreCellDataStill = Nothing
            Set mTimerDoNotStoreCellDataStill = New cTimer
            mTimerDoNotStoreCellDataStill.Interval = 1000
        End If
    End If
    
    mLastTextToFind = ""
    mValuesFromClientStoredSuccesfully = True
    mCountNoRedraw = 0
    
End Sub

Private Sub SetOrderByLastColHeader()
    Dim iLng As Long
    
    If mLastOrderColumnHeader <> "" Then
        If mLastColumnClicked < (mGridCols) Then
            If mLastColumnClicked > -1 Then
                iLng = GetColIndexByColHeader(mLastOrderColumnHeader)
                If iLng > -1 Then
                    If mColumnsHiddenByComponent(iLng) Then
                        mLastColumnClicked = -1
                        RemoveOrderArrowsFromGrid
                    Else
                        If (iLng <> mLastColumnClicked) Then
                            mLastColumnClicked = iLng
                            RemoveOrderArrowsFromGrid
                        End If
                    End If
                Else
                    mLastColumnClicked = -1
                    RemoveOrderArrowsFromGrid
                End If
            End If
        Else
            mLastColumnClicked = -1
            RemoveOrderArrowsFromGrid
        End If
    End If
End Sub

Private Sub SetAuxGridColumsWidths()
    Dim iRedraw_Prev As Boolean
    Dim iCol_Prev As Long
    Dim iColSel_Prev As Long
    Dim c As Long
    
    ReDim mAuxColWidths(mGrid.Cols - 1)
    iRedraw_Prev = mGrid.Redraw
    iCol_Prev = mGrid.col
    iColSel_Prev = mGrid.ColSel
    mGrid.Redraw = False
    
    For c = 0 To mGrid.Cols - 1
        If mGrid.ColWidth(c) <> -1 Then
            mAuxColWidths(c) = mGrid.ColWidth(c)
        Else
            If mDefaultColWidth = 0 Then
                If mGrid.Rows > 0 Then
                    mGrid.col = c
                    mDefaultColWidth = mGrid.CellWidth + Screen.TwipsPerPixelX
                Else
                    mDefaultColWidth = 960
                End If
            End If
            mAuxColWidths(c) = mDefaultColWidth
        End If
    Next c
    
    On Error Resume Next
    mGrid.col = iCol_Prev
    mGrid.ColSel = iColSel_Prev
    mGrid.Redraw = iRedraw_Prev Or mGrid.Redraw
End Sub

Private Sub SetAuxGridColumsHeaders(Optional nIncludeForOrder As Boolean)
    Dim c As Long
    Dim iStr As String
    Dim r As Long
    Dim iThereAreDefinedColHeaders As Boolean
    Dim iThereAreUndefinedColHeaders As Boolean
    
    On Error GoTo TheExit:

    ReDim mAuxColsHeaders(mGrid.Cols - 1)
    If nIncludeForOrder Then
        ReDim mAuxColsHeadersForOrder(mGrid.Cols - 1)
    End If
    
    mAllColHeadersAreEmpty = True
    For c = 0 To mGrid.Cols - 1
        iStr = ""
        For r = 0 To mGrid.FixedRows - 1
            iStr = iStr & mGrid.TextMatrix(r, c) & " "
        Next r
        If Len(iStr) > 0 Then
            iStr = Left$(iStr, Len(iStr) - 1)
        End If
        If Trim$(iStr) <> "" Then
            iStr = Replace(iStr, vbCrLf, " ")
            iStr = Replace(iStr, vbCr, " ")
            iStr = Replace(iStr, vbLf, " ")
            mAuxColsHeaders(c) = iStr
            If nIncludeForOrder Then
                mAuxColsHeadersForOrder(c) = iStr
            End If
            iThereAreDefinedColHeaders = True
            mAllColHeadersAreEmpty = False
        Else
            mAuxColsHeaders(c) = "Col. " & c + 1
            If nIncludeForOrder Then
                mAuxColsHeadersForOrder(c) = iStr
            End If
        End If
    Next c
    If nIncludeForOrder Then
        If Not iThereAreDefinedColHeaders Then
            If iThereAreUndefinedColHeaders Then
                For c = 0 To mGrid.Cols - 1
                    mAuxColsHeadersForOrder(c) = mAuxColsHeaders(c)
                Next c
            End If
        End If
    End If
    
TheExit:
End Sub

Private Sub SetGridVisibleColumnsWithDataFromRegistry()
    Dim c As Long
    Dim iStr As String
    Dim iKeyName As String
    Dim x As Long
    
    If Not mIsTypeSupported Then Exit Sub
    
    iKeyName = Keyname
    c = 1
    iStr = Base64Decode(GetSetting(AppNameForRegistry, "Preferences\" & iKeyName, "HideCol_" & CStr(c), ""))
    mThereAreHiddenCols = False
    Do Until iStr = ""
        x = GetColIndexByColHeader(iStr)
        If x > -1 Then
            mGrid.ColWidth(x) = 0
            mColumnsHiddenByComponent(x) = True
            mColumnsWidthsActual(x) = 0
            mColumnsRightPosSet = False
            mThereAreHiddenCols = True
        End If
        c = c + 1
        iStr = Base64Decode(GetSetting(AppNameForRegistry, "Preferences\" & iKeyName, "HideCol_" & CStr(c), ""))
    Loop
    If mLastColumnClicked > -1 Then
        If mColumnsHiddenByComponent(mLastColumnClicked) Then
            mLastColumnClicked = -1
        End If
    End If
End Sub

Private Sub SetGridVisibleColumnsWithDataFromVariable()
    Dim c As Long
    
    If Not mIsTypeSupported Then Exit Sub
    For c = 0 To mGrid.Cols - 1
        If mColumnsHiddenByComponent(c) Then
            mGrid.ColWidth(c) = 0
            mColumnsWidthsActual(c) = 0
            mColumnsRightPosSet = False
        End If
    Next c
End Sub

Private Function GetColIndexByColHeader(nColHeader As String)
    Dim c As Long
    
    GetColIndexByColHeader = -1
    For c = 0 To mGrid.Cols - 1
        If LCase$(mAuxColsHeaders(c)) = LCase$(nColHeader) Then
            GetColIndexByColHeader = c
            Exit Function
        End If
    Next c
End Function

Public Property Get ThereAreHiddenCols() As Boolean
    ThereAreHiddenCols = mThereAreHiddenCols
End Property

Private Sub CheckWhetherGridNeedsUpdate()
    
    If Not mIsTypeSupported Then Exit Sub
    
    If Not mGridValuesFromClientStored Then
        StoreGridValuesFromClient
        SetGridVisibleColumnsWithDataFromRegistry
        If mStretchColumnsWidthsToFill Then
            UpdateGridColumnsWidthsStretched
        End If
        If mSameDataGroupedInColumns Then
            UpdateDataGrouped
        End If
        If mEnableOrderByColumns Then
            If mDoNotRememberOrder Then
                mLastColumnClicked = -1
                OrderGridByColumn mInitialOrderColumn, mInitialOrderDescending, True
            Else
                OrderGridByColumn
            End If
'        Else
'            If mInitialOrderColumn <> -1 Then
'                OrderGridByColumn mInitialOrderColumn, mInitialOrderDescending
'            End If
        End If
    Else
        If GridNeedsToBeUpdated Then
            StoreGridValuesFromClient
            If mColsHeadersChanged Or mUpdateHiddenCols Then
                SetGridVisibleColumnsWithDataFromRegistry
            Else
                If mThereAreHiddenCols Then
                    SetGridVisibleColumnsWithDataFromVariable
                End If
            End If
            If mStretchColumnsWidthsToFill Then
                UpdateGridColumnsWidthsStretched
            End If
            If mSameDataGroupedInColumns Then
                UpdateDataGrouped
            End If
            If mEnableOrderByColumns Then
'                If mDoNotRememberOrder Then
'                    mLastColumnClicked = -1
'                    OrderGridByColumn mInitialOrderColumn, mInitialOrderDescending
'                Else
                    OrderGridByColumn
'                End If



'            Else
'                If mInitialOrderColumn <> -1 Then
'                    OrderGridByColumn mInitialOrderColumn, mInitialOrderDescending
'                End If
            End If
        Else
            If mNeedToUpdateOrderByColumns Then
                If mEnableOrderByColumns Then
'                    If mDoNotRememberOrder Then
'                        mLastColumnClicked = -1
'                        OrderGridByColumn mInitialOrderColumn, mInitialOrderDescending
'                    Else
                        OrderGridByColumn
'                    End If
                    
'                Else
'                    If mInitialOrderColumn <> -1 Then
'                        OrderGridByColumn mInitialOrderColumn, mInitialOrderDescending
'                    End If
                
                End If
            End If
        End If
        If mUpdateHiddenCols Then
            mUpdateHiddenCols = False
            If mThereAreHiddenCols Then
                SetGridVisibleColumnsWithDataFromVariable
                If mStretchColumnsWidthsToFill Then
                    UpdateGridColumnsWidthsStretched
                End If
            End If
        End If
    End If
    mGridValuesFromClientStored = mValuesFromClientStoredSuccesfully
End Sub

Private Function GridNeedsToBeUpdated() As Boolean
    Dim c  As Long
    
    If Not IsWindow(mGridHwnd) <> 0 Then Exit Function
    
    If mGrid.Cols <> mGridCols Then
        GridNeedsToBeUpdated = True
        mColsHeadersChanged = True
    Else
        If mGrid.Rows <> mGridRows Then
            GridNeedsToBeUpdated = True
        End If
        SetAuxGridColumsHeaders
        mColsHeadersChanged = False
        For c = 0 To mGrid.Cols - 1
            If Not GridNeedsToBeUpdated Then
                If mGrid.ColWidth(c) = -1 Then
                    If mColumnsWidthsFromClient(c) <> -1 Then
                        GridNeedsToBeUpdated = True
                    End If
                Else
                    If mGrid.ColWidth(c) <> mColumnsWidthsActual(c) Then
                        GridNeedsToBeUpdated = True
                    End If
                End If
            End If
            If mAuxColsHeaders(c) <> mColumnsHeaders(c) Then
                GridNeedsToBeUpdated = True
                mColsHeadersChanged = True
            End If
            If mColumnsHiddenByComponent(c) And (mGrid.ColWidth(c) <> 0) Then
'                GridNeedsToBeUpdated = True
'                mColsHeadersChanged = True
                mUpdateHiddenCols = True
            End If
        Next c
        mNeedToUpdateOrderByColumns = False
        If Not GridNeedsToBeUpdated Then
            If mEnableOrderByColumns Or (mInitialOrderColumn <> -1) Or mShowToolTipsOnLongerCellTexts Then
                If mGrid.Rows <> mGridRows Then
                    mNeedToUpdateOrderByColumns = True
                Else
                    If (mGrid.FixedCols < mGrid.Cols) And (mGrid.FixedRows < mGrid.Rows) Then
                        If mGrid.TextMatrix(mGrid.FixedRows, mGrid.FixedCols) <> mData00 Then
                            mNeedToUpdateOrderByColumns = True
                        ElseIf mGrid.TextMatrix(mGrid.FixedRows, mGrid.Cols - 1) <> mData01 Then
                            mNeedToUpdateOrderByColumns = True
                        ElseIf mGrid.TextMatrix(mGrid.Rows - 1, mGrid.FixedCols) <> mData10 Then
                            mNeedToUpdateOrderByColumns = True
                        ElseIf mGrid.TextMatrix(mGrid.Rows - 1, mGrid.Cols - 1) <> mData11 Then
                            mNeedToUpdateOrderByColumns = True
                        End If
                    End If
                    If mNeedToUpdateOrderByColumns Then
                        mData00 = mGrid.TextMatrix(mGrid.FixedRows, mGrid.FixedCols)
                        mData01 = mGrid.TextMatrix(mGrid.FixedRows, mGrid.Cols - 1)
                        mData10 = mGrid.TextMatrix(mGrid.Rows - 1, mGrid.FixedCols)
                        mData11 = mGrid.TextMatrix(mGrid.Rows - 1, mGrid.Cols - 1)
                    End If
                End If
                mGridRows = mGrid.Rows
            End If
        End If
    End If
End Function
    

Public Sub UnHideAllCols()
    Dim c As Long
    Dim iAuxThereWereColsHidden As Boolean
    
    If Not mIsTypeSupported Then Exit Sub
    
    If Not mGridValuesFromClientStored Then
        CheckWhetherGridNeedsUpdate
    End If
    If Not mThereAreHiddenCols Then Exit Sub
    
    For c = 0 To mGrid.Cols - 1
        If mColumnsHiddenByComponent(c) Then
            mGrid.ColWidth(c) = mColumnsWidthsFromClient(c)
            mColumnsWidthsActual(c) = mColumnsWidthsFromClientLiteral(c)
            mColumnsRightPosSet = False
            mColumnsHiddenByComponent(c) = False
            iAuxThereWereColsHidden = True
        End If
    Next c
    mThereAreHiddenCols = False
    On Error Resume Next
    DeleteSetting AppNameForRegistry, "Preferences\" & Keyname
    On Error GoTo 0
    If iAuxThereWereColsHidden Then
        If mStretchColumnsWidthsToFill Then
            UpdateGridColumnsWidthsStretched
        End If
    End If
End Sub

Public Sub HideCol(nCol As Long)
    
    If Not mIsTypeSupported Then Exit Sub
    
    If Not mGridValuesFromClientStored Then
        CheckWhetherGridNeedsUpdate
    End If
    If (nCol < 0) Or (nCol > (mGrid.Cols - 1)) Then Exit Sub
    
    mGrid.ColWidth(nCol) = 0
    mColumnsHiddenByComponent(nCol) = True
    mColumnsWidthsActual(nCol) = 0
    mColumnsRightPosSet = False
    mThereAreHiddenCols = True
    UpdateHiddenColsDataInRegistry
    If mStretchColumnsWidthsToFill Then
        UpdateGridColumnsWidthsStretched
    End If
End Sub

Public Sub UpdateGridColumnsWidthsStretched()
    Dim c As Long
    Dim iTotalWidth As Long
    Dim iAvailableWidth As Long
    Dim iFactor As Single
    Dim iRedraw_Prev As Boolean
    Dim iThereAreHiddenColumns As Boolean
    Dim iUpdated As Boolean
    Const cHorizontalSB = 1
    Dim iHasHowSB As Boolean
    
    If Not mIsTypeSupported Then Exit Sub
'    If Not mStretchColumnsWidthsToFill Then Exit Sub
    
    iRedraw_Prev = mGrid.Redraw
    mGrid.Redraw = False
    On Error GoTo TheExit:
    
    iAvailableWidth = mGrid.Parent.ScaleX(mGrid.Width, mGrid.Parent.ScaleMode, vbTwips)
    If mGrid.BorderStyle <> 0 Then
        iAvailableWidth = iAvailableWidth - 4 * Screen.TwipsPerPixelX
    End If
    If GridShowsVerticalScrollBar Then
        iAvailableWidth = iAvailableWidth - (GetSystemMetrics(SM_CXVSCROLL)) * Screen.TwipsPerPixelX
    End If
    
    For c = 0 To mGrid.Cols - 1
        If Not mColumnsHiddenByClientProgram(c) Then
            If Not mColumnsHiddenByComponent(c) Then
                iTotalWidth = iTotalWidth + mColumnsWidthsFromClientLiteral(c)
            Else
                iThereAreHiddenColumns = True
            End If
        Else
            iThereAreHiddenColumns = True
        End If
    Next c
'    If iThereAreHiddenColumns Then
'        iTotalWidth = iTotalWidth + Screen.TwipsPerPixelX  ' for some problem saw on the MSH grid, that adds the horizontal scrollbar
'    End If
    
    If iAvailableWidth > iTotalWidth Then
        If (mScrollBars And cHorizontalSB) <> 0 Then
            mGrid.ScrollBars = mGrid.ScrollBars And Not cHorizontalSB
        End If
    Else
        mGrid.ScrollBars = mScrollBars
    End If
    
    If iTotalWidth <> 0 Then
        iFactor = Round(iAvailableWidth - Screen.TwipsPerPixelX / 2) / iTotalWidth
        For c = 0 To mGrid.Cols - 1
            If Not mColumnsHiddenByClientProgram(c) Then
                If Not mColumnsHiddenByComponent(c) Then
                    If (iFactor > 1) And mStretchColumnsWidthsToFill Then
                        If mGrid.ColWidth(c) <> 0 Then
                            If Not iUpdated Then
                                iUpdated = mGrid.ColWidth(c) <> mColumnsWidthsFromClientLiteral(c) * iFactor
                            End If
                            mGrid.ColWidth(c) = mColumnsWidthsFromClientLiteral(c) * iFactor
                            mColumnsWidthsActual(c) = mGrid.ColWidth(c)
                        Else
                            mColumnsWidthsActual(c) = 0
                        End If
                    Else
                        If mGrid.ColWidth(c) <> 0 Then
                            mGrid.ColWidth(c) = mColumnsWidthsFromClient(c)
                            mColumnsWidthsActual(c) = mColumnsWidthsFromClientLiteral(c)
                        Else
                            mColumnsWidthsActual(c) = 0
                        End If
                    End If
                    mColumnsRightPosSet = False
                End If
            End If
        Next c
    End If

TheExit:
    
    iHasHowSB = CBool(mGrid.ScrollBars And cHorizontalSB)
    If Not iHasHowSB Then
        mGrid.ScrollBars = mGrid.ScrollBars Or cHorizontalSB
    End If
    mGrid.LeftCol = mGrid.FixedCols
    If Not iHasHowSB Then
        mGrid.ScrollBars = mGrid.ScrollBars And Not cHorizontalSB
    End If
    
    mGrid.Redraw = iRedraw_Prev Or mGrid.Redraw
    ParentFlexFnObject.EventRaise "AfterSettingColumnsWidths", mGrid.Name, CLng(iUpdated), 0
End Sub

Private Sub UpdateHiddenColsDataInRegistry()
    Dim c As Long
    Dim x As Long
    Dim iKeyName As String
    Dim iAuxSavedHiddenColsHeaders() As String
    Dim iStr As String
    
    iKeyName = Keyname
    ReDim iAuxSavedHiddenColsHeaders(0)
    x = 1
    Do
        iStr = Base64Decode(GetSetting(AppNameForRegistry, "Preferences\" & iKeyName, "HideCol_" & CStr(x), ""))
        If iStr = "" Then Exit Do
        If Not ColHeaderExist(iStr) Then
            ReDim Preserve iAuxSavedHiddenColsHeaders(UBound(iAuxSavedHiddenColsHeaders) + 1)
            iAuxSavedHiddenColsHeaders(UBound(iAuxSavedHiddenColsHeaders)) = iStr
        End If
        x = x + 1
    Loop
    
    x = 1
    On Error Resume Next
    DeleteSetting AppNameForRegistry, "Preferences\" & iKeyName
    On Error GoTo 0
    If mThereAreHiddenCols Then
        For c = 0 To mGrid.Cols - 1
            If mColumnsHiddenByComponent(c) Then
                SaveSetting AppNameForRegistry, "Preferences\" & iKeyName, "HideCol_" & CStr(x), Base64Encode(mColumnsHeaders(c))
                x = x + 1
            End If
        Next c
    End If
    For c = 1 To UBound(iAuxSavedHiddenColsHeaders)
        SaveSetting AppNameForRegistry, "Preferences\" & iKeyName, "HideCol_" & CStr(x), Base64Encode(iAuxSavedHiddenColsHeaders(c))
        x = x + 1
    Next c
End Sub

Private Function ColHeaderExist(nColHeader As String) As Boolean
    Dim c As Long
    
    For c = 0 To UBound(mColumnsHeaders)
        If mColumnsHeaders(c) = nColHeader Then
            ColHeaderExist = True
            Exit Function
        End If
    Next c
End Function

Public Sub ShowConfigGridColumns()
    Dim ifrmSelectColumns As New frmSelectColumns
    Dim c As Long
    Dim iChanged As Boolean
    
    If Not mGridValuesFromClientStored Then
        CheckWhetherGridNeedsUpdate
    End If
    ifrmSelectColumns.SetData mColumnsHeaders, mColumnsHiddenByClientProgram, mColumnsHiddenByComponent
    ShowModal ifrmSelectColumns
    mThereAreHiddenCols = False
    For c = 0 To UBound(mColumnsHiddenByComponent)
        If mColumnsHiddenByComponent(c) <> ifrmSelectColumns.ColumnsHidden(c) Then
            mColumnsHiddenByComponent(c) = ifrmSelectColumns.ColumnsHidden(c)
            iChanged = True
        End If
        If mColumnsHiddenByComponent(c) Then
            mThereAreHiddenCols = True
        End If
    Next c
    Set ifrmSelectColumns = Nothing
    If iChanged Then
        UpdateHiddenColsDataInRegistry
        For c = 0 To UBound(mColumnsHiddenByComponent)
            If mColumnsHiddenByComponent(c) Then
                mGrid.ColWidth(c) = 0
                mColumnsWidthsActual(c) = 0
            Else
                mGrid.ColWidth(c) = mColumnsWidthsFromClient(c)
                mColumnsWidthsActual(c) = mColumnsWidthsFromClientLiteral(c)
            End If
            mColumnsRightPosSet = False
        Next c
        If mEnableOrderByColumns Then
            SetOrderByLastColHeader
        End If
        If mStretchColumnsWidthsToFill Then
            UpdateGridColumnsWidthsStretched
        End If
    End If
End Sub

Private Sub SetColumnsRightPos()
    Dim c As Long
    
    On Error GoTo TheExit:
    
    mColumnsRightPos(0) = mGrid.Parent.ScaleX(mColumnsWidthsActual(0), vbTwips, vbPixels)
    For c = 1 To mGridCols - 1
        mColumnsRightPos(c) = mColumnsRightPos(c - 1) + mGrid.Parent.ScaleX(mColumnsWidthsActual(c), vbTwips, vbPixels)
    Next c
    mColumnsRightPosSet = True
    
TheExit:
End Sub

Private Function GetColumnByPosX(nXPos As Long)
    Dim c As Long
    Dim iLC As Long
    Dim iFC As Long
    Dim iFixedColsWidth As Long
    Dim iColWidthHiddenAtLeft As Long
    
    GetColumnByPosX = -1
    If Not mColumnsRightPosSet Then
        SetColumnsRightPos
    End If
    iFC = mGrid.FixedCols
    
    If iFC > 0 Then
        If nXPos < mColumnsRightPos(iFC - 1) Then
            For c = 0 To iFC - 1
                If nXPos < mColumnsRightPos(c) Then
                    GetColumnByPosX = c
                    Exit For
                End If
            Next c
        End If
    End If
    If GetColumnByPosX = -1 Then
        iLC = mGrid.LeftCol
        If iFC > 0 Then
            iFixedColsWidth = mColumnsRightPos(iFC - 1)
        End If
        If iLC > 0 Then
            iColWidthHiddenAtLeft = mColumnsRightPos(iLC - 1)
        End If
        For c = iLC To mGrid.Cols - 1
            If nXPos < (mColumnsRightPos(c) - iColWidthHiddenAtLeft + iFixedColsWidth) Then
                GetColumnByPosX = c
                Exit For
            End If
        Next c
    End If
End Function

Public Sub OrderGridByColumn(Optional nOrderByColumn As Long = -1, Optional nOrderDescending, Optional nDoNotSave As Boolean)
    Dim iKeyName As String
    Dim iOrderCol As Long
    Dim iOrderDescending As Boolean
    Dim iRedraw As Boolean
    Dim c As Long
    Dim iPrevRow As Long
    Dim iPrevCol As Long
    Dim iPrevRowSel As Long
    Dim iPrevColSel As Long
    Dim iPic As StdPicture
    Dim iColClicked As Long
    Dim iRowsRemoved As Long
    Dim iStr As String
    Dim iAuxLng As Long
    Dim iOrderDescending_Orig As Boolean
    
    If (nOrderByColumn = -1) And (mLastColumnClicked = -1) Then Exit Sub
    If (nOrderByColumn = -1) And mDoNotRememberOrder Then Exit Sub
    If nOrderByColumn > (mGridCols - 1) Then Exit Sub
    
    iKeyName = Keyname
    
    If nOrderByColumn = -1 Then
        iOrderCol = mLastColumnClicked
        iOrderDescending = mLastOrderDescending
    Else
        mGridWasOrderedByUser = True
        iOrderCol = nOrderByColumn
        If IsMissing(nOrderDescending) Then
            iOrderDescending = mLastOrderDescending
            If mLastColumnClicked = nOrderByColumn Then
                iOrderDescending = Not iOrderDescending
            Else
                iOrderDescending = False
            End If
        Else
            iOrderDescending = nOrderDescending
        End If
        iStr = mColumnsHeaders(iOrderCol)
        mLastOrderColumnHeader = iStr
        If iOrderDescending Then
            iStr = iStr & "_D"
        Else
            iStr = iStr & "_A"
        End If
        If Not nDoNotSave Then
            SaveSetting AppNameForRegistry, "Preferences", iKeyName & "_OrderByCol", Base64Encode(iStr)
        End If
    End If
    
    ' we have in iOrderCol the col to order, and in iOrderDescending if it's ascending or descending
    If (iOrderCol < 0) Or (iOrderCol > (mGrid.Cols - 1)) Then
        iOrderCol = 0
    End If
    
    iColClicked = iOrderCol
    iAuxLng = CLng(iOrderDescending)
    iOrderDescending_Orig = iOrderDescending
    ParentFlexFnObject.EventRaise "BeforeOrderingByColumn", mGrid.Name, iColClicked, iOrderCol, iAuxLng
    iOrderDescending = CBool(iAuxLng)
    
    If iOrderCol <> -1 Then
    
        mLastColumnClicked = iColClicked
        mLastOrderDescending = iOrderDescending_Orig
        mLastOrderChanged = iOrderDescending <> iOrderDescending_Orig
        
        iRedraw = mGrid.Redraw
        iPrevRow = mGrid.Row
        iPrevCol = mGrid.col
        iPrevRowSel = mGrid.RowSel
        iPrevColSel = mGrid.ColSel
        
        mGrid.Redraw = False
        mGrid.Cols = mGrid.Cols + 1
        mGrid.ColWidth(mGrid.Cols - 1) = 0
        mGrid.TextMatrix(iPrevRow, mGrid.Cols - 1) = "1"
        mGrid.TextMatrix(iPrevRowSel, mGrid.Cols - 1) = Val(mGrid.TextMatrix(iPrevRowSel, mGrid.Cols - 1)) + 2
        
        mGrid.Row = 0
        For c = 0 To mGrid.Cols - 1
            mGrid.col = c
            Set mGrid.CellPicture = Nothing
        Next c
        
        If iOrderDescending Then
            Set iPic = GetGridArrowDownImage(mGrid)
        Else
            Set iPic = GetGridArrowUpImage(mGrid)
        End If
        
        mGrid.col = iColClicked
        mGrid.CellPictureAlignment = flexAlignRightCenter
        Set mGrid.CellPicture = iPic
        
        iRowsRemoved = RemoveEmptyRowsAtTheEnd
        If mTooManyEmptyRowsAtTheEnd Then
'            mGrid.Row = iPrevRow
'            mGrid.Col = iPrevCol
'            mGrid.RowSel = iPrevRowSel
'            mGrid.ColSel = iPrevColSel
            If iPrevRow <> 0 Then
                RestoreSelection iPrevCol, iPrevColSel
            End If
            mGrid.Cols = mGrid.Cols - 1
            mGrid.Redraw = iRedraw Or mGrid.Redraw
            Exit Sub
        End If
        
        If iOrderCol > (mGrid.Cols - 1) Then
            mGrid.Redraw = iRedraw Or mGrid.Redraw
            Exit Sub
        End If
        mGrid.col = iOrderCol
        mGrid.ColSel = iOrderCol ' mGrid.Cols - 1
        If iOrderDescending Then
            mGrid.Sort = flexSortGenericDescending
        Else
            mGrid.Sort = flexSortGenericAscending
        End If
    End If
    
    AddEmptyRowsAtTheEnd iRowsRemoved
    
'    mGrid.Row = iPrevRow
'    mGrid.Col = iPrevCol
'    mGrid.RowSel = iPrevRowSel
'    mGrid.ColSel = iPrevColSel
    If (iPrevRow <> 0) And (iPrevColSel <> iPrevCol) Then
        RestoreSelection iPrevCol, iPrevColSel
    Else
        mGrid.Row = 0
        mGrid.TopRow = mGrid.FixedRows
        mGrid.col = iPrevCol
    End If
    mGrid.Cols = mGrid.Cols - 1
    mGrid.Redraw = iRedraw Or mGrid.Redraw
    
    ParentFlexFnObject.EventRaise "AfterOrderingByColumn", mGrid.Name, iColClicked, iOrderCol, CLng(iOrderDescending_Orig)
    
    mNeedToUpdateOrderByColumns = False
End Sub

Private Sub RestoreSelection(nPrevCol As Long, nPrevColSel As Long)
    Dim r As Long
    Dim iCol As Long
    Dim iSel1 As Boolean
    Dim iSel2 As Boolean
    Dim iR As Long
    Dim iLng As Long
    
    iR = mGrid.Row
    iCol = mGrid.Cols - 1
    For r = mGrid.FixedRows To mGrid.Rows - 1
        If mGrid.TextMatrix(r, iCol) <> "" Then
            Select Case mGrid.TextMatrix(r, iCol)
                Case "1"
                    mGrid.Row = r
                    iSel1 = True
                    If iSel2 Then Exit For
                Case "2"
                    mGrid.RowSel = r
                    iSel2 = True
                    If iSel1 Then Exit For
                Case "3"
                    mGrid.Row = r
                    mGrid.RowSel = r
                    Exit For
            End Select
        End If
    Next r
    mGrid.col = nPrevCol
    mGrid.ColSel = nPrevColSel
    If mGrid.Row <> iR Then
        If mGrid.HighLight <> flexHighlightNever Then
        
            If Not mGrid.RowIsVisible(mGrid.Row) Then
                iLng = mGrid.Row - 3
                If iLng < mGrid.FixedRows Then iLng = mGrid.FixedRows
                mGrid.TopRow = iLng
            Else
                If mGrid.Row < mGrid.Rows - 1 Then
                    If Not mGrid.RowIsVisible(mGrid.Row + 1) Then
                        iLng = mGrid.TopRow + 1
                        If iLng > mGrid.Rows - 1 Then iLng = mGrid.Rows - 1
                        mGrid.TopRow = iLng
                    End If
                Else
                    If mGrid.Row = mGrid.Rows - 1 Then
                        If mGrid.RowPos(mGrid.Row) + mGrid.RowHeight(mGrid.Row) > mGrid.Height Then
                            iLng = mGrid.TopRow + 1
                            If iLng > mGrid.Rows - 1 Then iLng = mGrid.Rows - 1
                            mGrid.TopRow = iLng
                        End If
                    End If
                End If
                If mGrid.Row > 1 Then
                    If Not mGrid.RowIsVisible(mGrid.Row - 1) Then
                        iLng = mGrid.TopRow - 1
                        If iLng < mGrid.FixedRows Then iLng = mGrid.FixedRows
                        mGrid.TopRow = iLng
                    End If
                End If
            End If
        
        End If
    End If
End Sub

Private Function RemoveEmptyRowsAtTheEnd() As Long
    Dim c As Long
    Dim r As Long
    Dim iEmpty As Boolean
    
    'Exit Function
    mTooManyEmptyRowsAtTheEnd = False
    r = mGrid.Rows - 1
    Do
        iEmpty = True
        For c = 0 To mGrid.Cols - 1
            If mGrid.TextMatrix(r, c) <> "" Then
                iEmpty = False
            End If
        Next c
        If iEmpty Then
            If r <= (mGrid.FixedRows) Then Exit Do
            RemoveEmptyRowsAtTheEnd = RemoveEmptyRowsAtTheEnd + 1
            r = r - 1
            If RemoveEmptyRowsAtTheEnd > 2000 Then
                RemoveEmptyRowsAtTheEnd = 0
                mTooManyEmptyRowsAtTheEnd = True
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    mGrid.Rows = mGrid.Rows - RemoveEmptyRowsAtTheEnd
End Function

Private Sub AddEmptyRowsAtTheEnd(nRows As Long)
'    Dim c As Long
'
'    For c = 1 To nRows
'        mGrid.AddItem ""
'    Next c
    mGrid.Rows = mGrid.Rows + nRows
End Sub

Private Sub UpdateDataGrouped()
    Dim c As Long
    Dim iRedraw As Boolean
    
    iRedraw = mGrid.Redraw
    mGrid.Redraw = False
    If mSameDataGroupedInColumns Then
        mGrid.MergeCells = flexMergeRestrictColumns
        For c = 0 To mGrid.Cols - 1
            mGrid.MergeCol(c) = mSameDataGroupedInColumns
        Next c
    Else
        mGrid.MergeCells = mOriginalMergeCells
        For c = 0 To mGrid.Cols - 1
            mGrid.MergeCol(c) = False
        Next c
    End If
    mGrid.Redraw = iRedraw Or mGrid.Redraw
End Sub

Private Sub RemoveOrderArrowsFromGrid()
    Dim iRedraw As Boolean
    Dim c As Long
    Dim iPrevRow As Long
    Dim iPrevCol As Long
    Dim iPrevRowSel As Long
    Dim iPrevColSel As Long
    
    iRedraw = mGrid.Redraw
    iPrevRow = mGrid.Row
    iPrevCol = mGrid.col
    iPrevRowSel = mGrid.RowSel
    iPrevColSel = mGrid.ColSel
    
    mGrid.Redraw = False
    
    mGrid.Row = 0
    For c = 0 To mGrid.Cols - 1
        mGrid.col = c
        Set mGrid.CellPicture = Nothing
    Next c

    mGrid.Row = iPrevRow
    mGrid.col = iPrevCol
    mGrid.RowSel = iPrevRowSel
    mGrid.ColSel = iPrevColSel
    mGrid.Redraw = iRedraw Or mGrid.Redraw

End Sub


Public Property Get BorderColor() As Long
    BorderColor = mBorderColor
End Property

Public Property Let BorderColor(nValue As Long)
    If nValue <> mBorderColor Then
        mBorderColor = nValue
        If mGridsFlatAppearance Then
            SetFlatAppearance
        End If
    End If
End Property


Public Property Get BorderWidth() As Long
    BorderWidth = mBorderWidth
End Property

Public Property Let BorderWidth(nValue As Long)
    If nValue <> mBorderWidth Then
        mBorderWidth = nValue
        If mGridsFlatAppearance Then
            SetFlatAppearance
        End If
    End If
End Property


Private Sub SetFlatAppearance()
    If mGridsFlatAppearance Then
        If Not mGridOriginalBorderStyleSet Then
            mGridOriginalBorderStyleSet = True
            mGridOriginalBorderStyle = mGrid.BorderStyle
            mGridOriginalGridLines = mGrid.GridLines
            mGridOriginalGridLinesFixed = mGrid.GridLinesFixed
        End If
        
        mGrid.BorderStyle = 0
        mGrid.GridLines = flexGridFlat             ' 1
        mGrid.GridLinesFixed = flexGridFlat    ' 1
        
        If mBorderColor = -1 Then
            mBorderColor = mGrid.GridColor
        End If
        If mBorderWidth = -1 Then
            mBorderWidth = mGrid.GridLineWidth
        End If
        
        mpicGridFlatLeft_Name = "picBorderLeft_" & mGridName
        mpicGridFlatRight_Name = "picBorderRight_" & mGridName
        mpicGridFlatTop_Name = "picBorderTop_" & mGridName
        mpicGridFlatBottom_Name = "picBorderBottom_" & mGridName

        On Error Resume Next
        
        Err.Clear
        Set mpicGridFlatLeft = mGrid.Parent.Controls.Add("VB.PictureBox", mpicGridFlatLeft_Name)
        If Err.Number = 727 Then
            Set mpicGridFlatLeft = mGrid.Parent.Controls(mpicGridFlatLeft_Name)
        End If
        mpicGridFlatLeft.BackColor = mBorderColor
        mpicGridFlatLeft.Width = mGrid.Parent.ScaleX(mBorderWidth, vbPixels, mGrid.Parent.ScaleMode)
        Set mpicGridFlatLeft.Container = mGrid.Container
        mpicGridFlatLeft.TabStop = False
        mpicGridFlatLeft.ZOrder
        
        Err.Clear
        Set mpicGridFlatRight = mGrid.Parent.Controls.Add("VB.PictureBox", mpicGridFlatRight_Name)
        If Err.Number = 727 Then
            Set mpicGridFlatRight = mGrid.Parent.Controls(mpicGridFlatRight_Name)
        End If
        mpicGridFlatRight.BackColor = mBorderColor
        mpicGridFlatRight.Width = mGrid.Parent.ScaleX(mBorderWidth, vbPixels, mGrid.Parent.ScaleMode)
        Set mpicGridFlatRight.Container = mGrid.Container
        mpicGridFlatRight.TabStop = False
        mpicGridFlatRight.ZOrder
        
        Err.Clear
        Set mpicGridFlatTop = mGrid.Parent.Controls.Add("VB.PictureBox", mpicGridFlatTop_Name)
        If Err.Number = 727 Then
            Set mpicGridFlatTop = mGrid.Parent.Controls(mpicGridFlatTop_Name)
        End If
        mpicGridFlatTop.BackColor = mBorderColor
        mpicGridFlatTop.Height = mGrid.Parent.ScaleY(mBorderWidth, vbPixels, mGrid.Parent.ScaleMode)
        Set mpicGridFlatTop.Container = mGrid.Container
        mpicGridFlatTop.TabStop = False
        mpicGridFlatTop.ZOrder
        
        Err.Clear
        Set mpicGridFlatBottom = mGrid.Parent.Controls.Add("VB.PictureBox", mpicGridFlatBottom_Name)
        If Err.Number = 727 Then
            Set mpicGridFlatBottom = mGrid.Parent.Controls(mpicGridFlatBottom_Name)
        End If
        mpicGridFlatBottom.BackColor = mBorderColor
        mpicGridFlatBottom.Height = mGrid.Parent.ScaleY(mBorderWidth, vbPixels, mGrid.Parent.ScaleMode)
        Set mpicGridFlatBottom.Container = mGrid.Container
        mpicGridFlatBottom.TabStop = False
        mpicGridFlatBottom.ZOrder
        
        On Error GoTo 0
        
        PositionBorderPicBoxes
        
    Else
        
        mGrid.BorderStyle = mGridOriginalBorderStyle
        mGrid.GridLines = mGridOriginalGridLines
        mGrid.GridLinesFixed = mGridOriginalGridLinesFixed
        
        On Error Resume Next
        
        Set mpicGridFlatLeft = Nothing
        mGrid.Parent.Controls.Remove (mpicGridFlatLeft_Name)
        Set mpicGridFlatRight = Nothing
        mGrid.Parent.Controls.Remove (mpicGridFlatRight_Name)
        Set mpicGridFlatTop = Nothing
        mGrid.Parent.Controls.Remove (mpicGridFlatTop_Name)
        Set mpicGridFlatBottom = Nothing
        mGrid.Parent.Controls.Remove (mpicGridFlatBottom_Name)
        
        On Error GoTo 0
    
    End If
End Sub

Private Sub PositionBorderPicBoxes()
    Dim mOnePixelXInContainerUnits As Single
    Dim mOnePixelYInContainerUnits As Single
    
    'mShape.Move mControl.Left - mBorderWidth * Screen.TwipsPerPixelX, mControl.Top - mBorderWidth * Screen.TwipsPerPixelY, mControl.Width + (2 * mBorderWidth) * Screen.TwipsPerPixelX, mControl.Height + (2 * mBorderWidth) * Screen.TwipsPerPixelY
    On Error Resume Next
    
    mOnePixelXInContainerUnits = mGrid.Container.ScaleX(100, vbPixels, mGrid.Container.ScaleMode) / 100
    If mOnePixelXInContainerUnits = 0 Then
        mOnePixelXInContainerUnits = mGrid.Parent.ScaleX(100, vbPixels, mGrid.Parent.ScaleMode) / 100
    End If
    If mOnePixelXInContainerUnits = 0 Then
        mOnePixelXInContainerUnits = Screen.TwipsPerPixelX
    End If
    
    mOnePixelYInContainerUnits = mGrid.Container.ScaleY(100, vbPixels, mGrid.Container.ScaleMode) / 100
    If mOnePixelYInContainerUnits = 0 Then
        mOnePixelYInContainerUnits = mGrid.Parent.ScaleY(100, vbPixels, mGrid.Parent.ScaleMode) / 100
    End If
    If mOnePixelYInContainerUnits = 0 Then
        mOnePixelYInContainerUnits = Screen.TwipsPerPixelY
    End If
    
    On Error GoTo 0
    
    mpicGridFlatLeft.Move mGrid.Left - mBorderWidth * mOnePixelXInContainerUnits, mGrid.Top - mBorderWidth * mOnePixelYInContainerUnits, mBorderWidth * mOnePixelXInContainerUnits, mGrid.Height + mBorderWidth * mOnePixelYInContainerUnits * 2
    mpicGridFlatRight.Move mGrid.Left + mGrid.Width, mGrid.Top - mBorderWidth * mOnePixelYInContainerUnits, mBorderWidth * mOnePixelXInContainerUnits, mGrid.Height + mBorderWidth * mOnePixelYInContainerUnits * 2
    mpicGridFlatTop.Move mGrid.Left - mBorderWidth * mOnePixelXInContainerUnits, mGrid.Top - mBorderWidth * mOnePixelYInContainerUnits, mGrid.Width + mBorderWidth * mOnePixelXInContainerUnits * 2, mBorderWidth * mOnePixelYInContainerUnits
    mpicGridFlatBottom.Move mGrid.Left - mBorderWidth * mOnePixelXInContainerUnits, mGrid.Top + mGrid.Height, mGrid.Width + mBorderWidth * mOnePixelXInContainerUnits * 2, mBorderWidth * mOnePixelYInContainerUnits
End Sub

Private Sub GetRowColUnderMouse(ByRef nRowUnderMouse As Long, ByRef nColUnderMouse As Long)
    Dim iRowHeight1 As Long
    Dim iColWidth1 As Long
    Dim iPm As POINTAPI
    Dim iRect As RECT
    Dim iLng As Long
    Dim iFound As Boolean
    Dim iMousePosInGrid As POINTAPI
    Dim iSusbstracted As Boolean
    
    On Error GoTo TheExit:
    
    If Not mGridValuesFromClientStored Then
        StoreGridValuesFromClient
    End If
    
    GetCursorPos iPm
    GetWindowRect mGridHwnd, iRect
    
    If ((mGrid.Rows - mGrid.FixedRows) = 0) Or ((mGrid.Cols - mGrid.FixedCols) = 0) Then
        nRowUnderMouse = -1
        nColUnderMouse = -1
        Exit Sub
    End If
    iMousePosInGrid.x = iPm.x - iRect.Left
    iMousePosInGrid.y = iPm.y - iRect.Top
    
    iRowHeight1 = mGrid.RowHeight(mGrid.FixedRows) / Screen.TwipsPerPixelY
    iLng = (iMousePosInGrid.y - mFixedHeight) / iRowHeight1 + mGrid.TopRow - 1
    
    If iLng < mGrid.FixedRows Then iLng = mGrid.FixedRows
    If iLng < mGrid.TopRow Then iLng = mGrid.TopRow
    If iLng > (mGridRows - 1) Then iLng = mGridRows - 1
    
    If iLng < 0 Then
        nRowUnderMouse = -1
        nColUnderMouse = -1
        Exit Sub
    End If
    
    Do Until iFound
        If (iMousePosInGrid.y * Screen.TwipsPerPixelY + mGrid.RowPos(mGrid.TopRow) - mFixedHeight * Screen.TwipsPerPixelY) < mGrid.RowPos(iLng) Then
            iLng = iLng - 1
            iSusbstracted = True
            If iLng < mGrid.FixedRows Then
                iLng = -1
                Exit Do
            End If
        Else
            If (iMousePosInGrid.y * Screen.TwipsPerPixelY + mGrid.RowPos(mGrid.TopRow) - mFixedHeight * Screen.TwipsPerPixelY) > (mGrid.RowPos(iLng) + mGrid.RowHeight(iLng)) Then
                If iSusbstracted Then Exit Do
                iLng = iLng + 1
                If iLng > (mGridRows - 1) Then
                    iLng = -1
                    Exit Do
                End If
            Else
                Exit Do
            End If
        End If
    Loop
    nRowUnderMouse = iLng
    iSusbstracted = False
    
    iColWidth1 = mGrid.ColWidth(mGrid.FixedCols) / Screen.TwipsPerPixelX
    iLng = (iMousePosInGrid.x - mFixedWidth) / iColWidth1 + mGrid.LeftCol - 1
    
    If iLng < mGrid.FixedCols Then iLng = mGrid.FixedCols
    If iLng < mGrid.LeftCol Then iLng = mGrid.LeftCol
    If iLng > (mGridCols - 1) Then iLng = mGridCols - 1
    
    If iLng < 0 Then
        nRowUnderMouse = -1
        nColUnderMouse = -1
        Exit Sub
    End If
    
    Do Until iFound
        If (iMousePosInGrid.x * Screen.TwipsPerPixelX + mGrid.ColPos(mGrid.LeftCol) - mFixedWidth * Screen.TwipsPerPixelX) < mGrid.ColPos(iLng) Then
            iLng = iLng - 1
            iSusbstracted = True
            If iLng < mGrid.FixedCols Then
                iLng = -1
                Exit Do
            End If
        Else
            If (iMousePosInGrid.x * Screen.TwipsPerPixelX + mGrid.ColPos(mGrid.LeftCol) - mFixedWidth * Screen.TwipsPerPixelX) > (mGrid.ColPos(iLng) + mGrid.ColWidth(iLng)) Then
                If iSusbstracted Then Exit Do
                iLng = iLng + 1
                If iLng > (mGridCols - 1) Then
                    iLng = -1
                    Exit Do
                End If
            Else
                Exit Do
            End If
        End If
    Loop
    nColUnderMouse = iLng

TheExit:

End Sub

Private Sub StoreCellsData()
    Dim iFontChanged As Boolean
    Dim iCurrentFontIndex As Long
    Dim iFontFound As Boolean
    
    Dim iFontName_Prev As String
    Dim iFontSize_Prev As Single
    Dim iFontBold_prev As Boolean
    Dim iFontItalic_Prev As Boolean
    Dim iFontStrikethrough_Prev As Boolean
    Dim iFontUnderline_Prev As Boolean
    Dim iFontWidth_Prev As Long
    Dim iGrid_CellFontStrikeThrough As Boolean
    
    Dim r As Long
    Dim c As Long
    Dim f As Long
    
'    Dim iRedraw_Prev As Boolean
    Dim iCol_Prev As Long
    Dim iColSel_Prev As Long
    Dim iRow_Prev As Long
    Dim iRowSel_Prev As Long
    Dim iAFHwnd As Long
    
    On Error GoTo TheExit:
    If mCellDataStored Then Exit Sub
    If Not mGridValuesFromClientStored Then Exit Sub
    If mGrid.Redraw = False Then Exit Sub
    If mGridRows > 5000 Then Exit Sub
    If mGridRows = 0 Then Exit Sub
    
    CheckWhetherGridNeedsUpdate

    mCellDataStored = True
    
    ReDim mCellFontIndex(mGridRows - 1, mGridCols - 1)
    ReDim mCellForeColor(mGridRows - 1, mGridCols - 1)
    ReDim mFontsAttributes(0)
    
    iAFHwnd = GetActiveFormHwnd
'    iRedraw_Prev = mGrid.Redraw
    
    iRow_Prev = mGrid.Row
    iRowSel_Prev = mGrid.RowSel
    iCol_Prev = mGrid.col
    iColSel_Prev = mGrid.ColSel
    
    mGrid.Redraw = False
    
    For r = mGrid.FixedRows To mGridRows - 1
        mGrid.Row = r
        For c = mGrid.FixedCols To mGridCols - 1
            mGrid.col = c
            mCellForeColor(r, c) = mGrid.CellForeColor
            
            On Error Resume Next
            Err.Clear
            iGrid_CellFontStrikeThrough = mGrid.CellFontStrikeThrough
            If Err.Number Then
                iGrid_CellFontStrikeThrough = iFontStrikethrough_Prev
            End If
            If Not mGrid.Redraw Then On Error GoTo TheExit:
            
            iFontChanged = False
            If mGrid.CellFontName <> iFontName_Prev Then
                iFontChanged = True
            ElseIf mGrid.CellFontSize <> iFontSize_Prev Then
                iFontChanged = True
            ElseIf mGrid.CellFontBold <> iFontBold_prev Then
                iFontChanged = True
            ElseIf mGrid.CellFontItalic <> iFontItalic_Prev Then
                iFontChanged = True
            ElseIf mGrid.CellFontUnderline <> iFontUnderline_Prev Then
                iFontChanged = True
            ElseIf iGrid_CellFontStrikeThrough <> iFontStrikethrough_Prev Then
                iFontChanged = True
            Else
                On Error Resume Next
                Err.Clear
                If mGrid.CellFontWidth <> iFontWidth_Prev Then
                    If Err.Number = 0 Then
                        iFontChanged = True
                    End If
                End If
                If Not mGrid.Redraw Then On Error GoTo TheExit:
            End If
            
            If iFontChanged Then
                iFontName_Prev = mGrid.CellFontName
                iFontSize_Prev = mGrid.CellFontSize
                iFontBold_prev = mGrid.CellFontBold
                iFontItalic_Prev = mGrid.CellFontItalic
                iFontUnderline_Prev = mGrid.CellFontUnderline
                On Error Resume Next
                iFontStrikethrough_Prev = mGrid.CellFontStrikeThrough
                iFontWidth_Prev = mGrid.CellFontWidth
                If Not mGrid.Redraw Then On Error GoTo TheExit:
                
                iFontFound = False
                For f = 1 To UBound(mFontsAttributes)
                    If mFontsAttributes(f).Name = iFontName_Prev Then
                        If mFontsAttributes(f).Size = iFontSize_Prev Then
                            If mFontsAttributes(f).Bold = iFontBold_prev Then
                                If mFontsAttributes(f).Italic = iFontItalic_Prev Then
                                    If mFontsAttributes(f).Underline = iFontUnderline_Prev Then
                                        If mFontsAttributes(f).Strikethrough = iFontStrikethrough_Prev Then
                                            If mFontsAttributes(f).Width = iFontWidth_Prev Then
                                                iFontFound = True
                                                iCurrentFontIndex = f
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next f
                
                If Not iFontFound Then
                    iCurrentFontIndex = UBound(mFontsAttributes) + 1
                    ReDim Preserve mFontsAttributes(iCurrentFontIndex)
                    Set mFontsAttributes(iCurrentFontIndex) = New cFontAttributes
                    
                    mFontsAttributes(iCurrentFontIndex).Name = iFontName_Prev
                    mFontsAttributes(iCurrentFontIndex).Size = iFontSize_Prev
                    mFontsAttributes(iCurrentFontIndex).Bold = iFontBold_prev
                    mFontsAttributes(iCurrentFontIndex).Italic = iFontItalic_Prev
                    mFontsAttributes(iCurrentFontIndex).Underline = iFontUnderline_Prev
                    mFontsAttributes(iCurrentFontIndex).Strikethrough = iFontStrikethrough_Prev
                    mFontsAttributes(iCurrentFontIndex).Width = iFontWidth_Prev
                End If
            End If
            
            mCellFontIndex(r, c) = iCurrentFontIndex
        Next c
        
        If r Mod 1000 = 0 Then
            DoEvents
            If mGrid.Row <> r Then
                mCellDataStored = False
                GoTo TheExit:
            End If
            If GetActiveFormHwnd <> iAFHwnd Then
                mCellDataStored = False
                GoTo TheExit:
            End If
        End If
    Next r
    
    mGrid.Row = iRow_Prev
    mGrid.RowSel = iRowSel_Prev
    mGrid.col = iCol_Prev
    mGrid.ColSel = iColSel_Prev
    mGrid.Redraw = True '  iRedraw_Prev Or mGrid.Redraw
    Exit Sub
    
TheExit:
    On Error GoTo -1
    On Error Resume Next
    mGrid.Redraw = True
End Sub

Private Sub TestWhetherToShowToolTip()
    Dim iFont As New StdFont
    Dim iLng As Long
    Dim iWordWrap As Boolean
    Dim iRectCell As RECT
    Dim iRectCalc As RECT
    Dim iDC As Long
    Dim iLOGFONT As LOGFONTW
    Dim iFontHandle As Long
    Dim iOldFont As Long
    Dim iText As String
    Dim iP As POINTAPI
    Dim iEP As DRAWTEXTPARAMS
    Dim iPlaceToolTip As Boolean
    Dim iCellForeColor As Long
    Dim iR As RECT
    Dim iWinStyleInfo As Long
    Dim iFontIndex As Long
    Dim iAuxHeight As Long
    Dim iStr As String
    
    On Error GoTo errorHandler:
    
    iText = mGrid.TextMatrix(mRowUnderMouse, mColUnderMouse)
    
    If Trim$(iText) = "" Then Exit Sub
    
    If Not mCellDataStored Then
        StoreCellsData
    End If
    
    iFontIndex = mCellFontIndex(mRowUnderMouse, mColUnderMouse)
    
    iFont.Name = mFontsAttributes(iFontIndex).Name
    iFont.Size = mFontsAttributes(iFontIndex).Size
    iFont.Bold = mFontsAttributes(iFontIndex).Bold
    iFont.Italic = mFontsAttributes(iFontIndex).Italic
    iFont.Underline = mFontsAttributes(iFontIndex).Underline
    iFont.Strikethrough = mFontsAttributes(iFontIndex).Strikethrough
    
    iLng = -1
    On Error Resume Next
    iLng = mGrid.ColWordWrapOption(mColUnderMouse)
    If iLng = -1 Then
        iWordWrap = mGrid.WordWrap
    Else
        iWordWrap = (iLng And 1) = 1
    End If
    On Error GoTo 0
    
    iRectCell.Right = mGrid.ColWidth(mColUnderMouse) / Screen.TwipsPerPixelX
    iRectCell.Bottom = mGrid.RowHeight(mRowUnderMouse) / Screen.TwipsPerPixelY
    
    iDC = CreateCompatibleDC(GetDC(mGridHwnd))
    If iDC = 0 Then Exit Sub
    
    iLOGFONT = StdFontToLogFont_Screen(iDC, iFont)
    iFontHandle = CreateFontIndirectW(iLOGFONT)
    iOldFont = SelectObject(iDC, iFontHandle)
    
    iRectCalc.Right = iRectCell.Right
    
    If iWordWrap Then
        iEP.cbSize = Len(iEP)
        Call DrawTextEx(iDC, StrPtr(iText), Len(iText), iRectCalc, DT_EXPANDTABS Or DT_TABSTOP Or DT_NOPREFIX Or DT_WORDBREAK Or DT_CALCRECT, iEP)
        If iRectCalc.Bottom > iRectCell.Bottom Then
            iPlaceToolTip = True
        End If
    Else
        GetTextExtentPoint32 iDC, iText, Len(iText), iP
        If iP.x > (iRectCell.Right - GetSystemMetrics(SM_CXEDGE) * 2) Then
            iPlaceToolTip = True
        End If
    End If
    
    If iPlaceToolTip Then
        mToolTipIsVisible = True
        On Error Resume Next
        mGridParent.Controls.Add "VB.PictureBox", mGrid.Name & "_" & "ToolTipPic"
        Set mToolTipPic = mGridParent.Controls(mGrid.Name & "_" & "ToolTipPic")
        mToolTipPic.Appearance = 0
        mGridParent.Controls.Add "VB.Label", mGrid.Name & "_" & "ToolTipLabel"
        Set mToolTipLabel = mGridParent.Controls(mGrid.Name & "_" & "ToolTipLabel")
        Set mToolTipLabel.Container = mToolTipPic
        Set mToolTipLabel.Font = iFont
        iCellForeColor = mCellForeColor(mRowUnderMouse, mColUnderMouse)
        TranslateColor iCellForeColor, 0, iCellForeColor
        If ColorBrightNess(iCellForeColor) > 50 Then
            iCellForeColor = vbBlack
        End If
        mToolTipLabel.ForeColor = iCellForeColor
        mToolTipPic.BackColor = vbInfoBackground ' &H80000018
        mToolTipLabel.BackColor = vbInfoBackground ' &H80000018
        
        GetWindowRect mGridHwnd, iR
        iR.Left = iR.Left + mGrid.ColPos(mColUnderMouse) / Screen.TwipsPerPixelX
        iR.Top = iR.Top + mGrid.RowPos(mRowUnderMouse) / Screen.TwipsPerPixelY
        iR.Right = iR.Left + iRectCell.Right
        iR.Bottom = iR.Top + iRectCell.Bottom
        
        mToolTipLabel.Top = ((mGrid.GridLineWidth - 1) / 2 - 101 + (100 - iFont.Size / 1.15 + 10 / 1.15)) * Screen.TwipsPerPixelX
        
        mToolTipLabel.Left = ((mGrid.GridLineWidth - 1) / 2 + 2) * Screen.TwipsPerPixelY
        If iWordWrap Then
            mToolTipLabel.Top = mToolTipLabel.Top + Screen.TwipsPerPixelY
            mToolTipLabel.WordWrap = True
            SetWindowPos mToolTipPic.hWnd, 0, iR.Left + GetSystemMetrics(SM_CXEDGE) - 2, iR.Top, iR.Right - iR.Left, (iRectCalc.Bottom - iRectCalc.Top) + GetSystemMetrics(SM_CXEDGE) * 2, SWP_NOACTIVATE    'Or SWP_SHOWWINDOW
            mToolTipLabel.Width = mToolTipPic.Width - 6 * Screen.TwipsPerPixelX
            mToolTipLabel.Height = mToolTipPic.Height
        Else
            If (iR.Left + iP.x) > (Screen.Width / Screen.TwipsPerPixelX) Then
                iRectCalc.Left = 0
                iRectCalc.Right = Screen.Width / Screen.TwipsPerPixelX - iR.Left - 6
                iEP.cbSize = Len(iEP)
                Call DrawTextEx(iDC, StrPtr(iText), Len(iText), iRectCalc, DT_EXPANDTABS Or DT_TABSTOP Or DT_NOPREFIX Or DT_WORDBREAK Or DT_CALCRECT, iEP)
                
                mToolTipLabel.WordWrap = True
                SetWindowPos mToolTipPic.hWnd, 0, iR.Left + GetSystemMetrics(SM_CXEDGE) - 2, iR.Top, Screen.Width / Screen.TwipsPerPixelX - iR.Left, (iRectCalc.Bottom - iRectCalc.Top) + GetSystemMetrics(SM_CXEDGE) * 2, SWP_NOACTIVATE    'Or SWP_SHOWWINDOW
                mToolTipLabel.Top = mToolTipLabel.Top + Screen.TwipsPerPixelY * 2
                mToolTipLabel.Width = mToolTipPic.Width - 6 * Screen.TwipsPerPixelX
                mToolTipLabel.Height = mToolTipPic.Height
            Else
                Set mToolTipPic.Font = mToolTipLabel.Font
                iAuxHeight = mToolTipPic.TextHeight(iText) / Screen.TwipsPerPixelY + GetSystemMetrics(SM_CYEDGE) * 3
                If iAuxHeight > (iR.Bottom - iR.Top) Then
                    iR.Bottom = iR.Top + iAuxHeight
                End If
                SetWindowPos mToolTipPic.hWnd, 0, iR.Left + GetSystemMetrics(SM_CXEDGE) - 2, iR.Top + GetSystemMetrics(SM_CXEDGE), iP.x + 5, (iR.Bottom - iR.Top) - GetSystemMetrics(SM_CXEDGE) * 2, SWP_NOACTIVATE 'Or SWP_SHOWWINDOW
                mToolTipLabel.Width = mToolTipPic.Width
            End If
        End If
        
        mToolTipLabel.Caption = iText
        iStr = mToolTipLabel.Caption
        iStr = Replace(iStr, "?", "")
        iStr = Replace(iStr, " ", "")
        If iStr <> "" Then
            mToolTipLabel.Visible = True
            iWinStyleInfo = GetWindowLong(mToolTipPic.hWnd, GWL_EXSTYLE)
            iWinStyleInfo = iWinStyleInfo Or WS_EX_TOOLWINDOW
            SetWindowLong mToolTipPic.hWnd, GWL_EXSTYLE, iWinStyleInfo
            SetParent mToolTipPic.hWnd, 0
            MySetProp mToolTipPic.hWnd, "TTPic", 1
            mToolTipPic.ZOrder
            mToolTipPic.Visible = True
        End If
                
        mRowToolTip = mRowUnderMouse
        mColToolTip = mColUnderMouse
    End If
    
    If iOldFont <> 0 Then
        SelectObject iDC, iOldFont
    End If
    DeleteDC iDC
    DeleteObject iFontHandle
        
    Exit Sub

errorHandler:
    Dim iCame As Boolean
    
    If Not iCame Then
        mCellDataStored = False
        StoreCellsData
        iCame = True
        Resume
    End If
End Sub

Private Sub RemoveToolTip()
    On Error Resume Next
    mToolTipPic.Visible = False
    SetParent mToolTipPic.hWnd, mGridParent.hWnd
    mGridParent.Controls.Remove mGrid.Name & "_" & "ToolTipPic"
    mGridParent.Controls.Remove mGrid.Name & "_" & "ToolTipLabel"
    MyRemoveProp mToolTipPic.hWnd, "TTPic"
    mToolTipIsVisible = False
End Sub

Private Sub ShowHideBorder()
    Dim iVisible As Boolean
    
    iVisible = IsWindowVisible(mGridHwnd) <> 0
    mpicGridFlatLeft.Visible = iVisible
    mpicGridFlatRight.Visible = iVisible
    mpicGridFlatTop.Visible = iVisible
    mpicGridFlatBottom.Visible = iVisible
    
End Sub

Private Function GridShowsVerticalScrollBar() As Boolean
    Dim r As Long
    Dim iHeight As Long
    Dim iGridHeight As Long
    
    If mGlobals.IsShowingVerticalScrollBar(mGrid) Then
        GridShowsVerticalScrollBar = True
    Else
        iGridHeight = mGrid.Height
        
        For r = 0 To mGrid.Rows - 1
            iHeight = iHeight + mGrid.RowHeight(r)
            If iHeight >= iGridHeight Then
                GridShowsVerticalScrollBar = True
                Exit Function
            End If
        Next r
    End If
    
End Function

Private Function MouseIsOverGrid() As Boolean
    Dim iP As POINTAPI
    Dim iR As RECT
    Dim iHwndWindowUnderMouse As Long
    
    GetCursorPos iP
    GetWindowRect mGridHwnd, iR
    
    iHwndWindowUnderMouse = WindowFromPoint(iP.x, iP.y)
    
    If (iHwndWindowUnderMouse = mGridHwnd) Or (MyGetProp(iHwndWindowUnderMouse, "TTPic") <> 0) Then
        If iP.y > iR.Top Then
            If iP.y < iR.Bottom Then
                If iP.x > iR.Left Then
                    iR.Right = iR.Right - GetSystemMetrics(SM_CXEDGE) * 2
                    If iP.x < iR.Right Then
                        MouseIsOverGrid = True
                    End If
                End If
            End If
        End If
    End If
End Function

Private Function MouseIsOverForm() As Boolean
    Dim iP As POINTAPI
    Dim iR As RECT
    Dim iHwndWindowUnderMouse As Long
    
    GetCursorPos iP
    GetWindowRect mGridHwnd, iR
    
    iHwndWindowUnderMouse = WindowFromPoint(iP.x, iP.y)
    
    MouseIsOverForm = iHwndWindowUnderMouse = mGridFormHwnd
End Function

Private Sub EditCell()
    Dim iLng As Long
    Dim iWordWrap As Boolean
    Dim iAuxLeft As Long
    Dim iAuxTop As Long
    Dim iAuxRedraw As Boolean
    Dim iAuxCol As Long
    Dim iAuxRow As Long
    Dim iAuxColSel As Long
    Dim iAuxRowSel As Long
    Dim iCtl As Control
    Dim iAuxFont As StdFont
    Dim iAuxHeight As Long
    Dim iCancel As Long
    Dim iStr As String
    Dim iAuxWW As Boolean
    
    ParentFlexFnObject.EventRaise "BeforeTextEdit", mGrid.Name, iCancel, 0&
    If iCancel <> 0 Then
'        mAllowTextEdition = False
        Exit Sub
    End If
    
    iLng = -1
    On Error Resume Next
    iLng = mGrid.ColWordWrapOption(mGrid.MouseCol)
    If iLng = -1 Then
        iWordWrap = mGrid.WordWrap
    Else
        iWordWrap = (iLng And 1) = 1
    End If
    
    iAuxRedraw = mGrid.Redraw
    mGrid.Redraw = False
    iAuxCol = mGrid.col
    iAuxRow = mGrid.Row
    iAuxColSel = mGrid.ColSel
    iAuxRowSel = mGrid.RowSel
    
    If mGrid.SelectionMode = flexSelectionByRow Then
        mEditCellCol = mGrid.MouseCol
        mGrid.col = mEditCellCol
        mEditCellRow = mGrid.Row
    ElseIf mGrid.SelectionMode = flexSelectionByColumn Then
        mEditCellCol = mGrid.col
        mEditCellRow = mGrid.MouseRow
        mGrid.Row = mEditCellRow
    Else
        mEditCellCol = mGrid.col
        mEditCellRow = mGrid.Row
    End If
    
    If mUCControls Is Nothing Then
        Set mUCControls = ParentFlexFnObject.GetControls
    End If
    If mUCControls Is Nothing Then
        Exit Sub
    End If
    Err.Clear
    mUCControls.Add App.EXEName & ".uscTextBox", mGrid.Name & "_" & "EditCellTextBox"
    'Debug.Print Err.Number, Err.Description
    
    If Not mGrid.ColIsVisible(mGrid.col) Then GoTo Salida:
    If mGrid.ColWidth(mGrid.col) = 0 Then GoTo Salida:
    
    Set iCtl = mUCControls(mGrid.Name & "_" & "EditCellTextBox")
    If Not iCtl Is Nothing Then
        mEditingCell = True
    Else
        GoTo Salida:
    End If
    iCtl.MultiLine = iWordWrap
    Set mEditCellTextBox = iCtl.TextBoxControl
    
    mEditCellTextBox.Alignment = FilterAlignment(mGrid.CellAlignment, efnFilterAligmentHorizontal)
    mEditCellTextBox.Visible = False
    mEditCellTextBoxOldParent = SetParent(mEditCellTextBox.hWnd, mGrid.Container.hWnd)
    If mEditCellTextBoxOldParent = 0 Then Exit Sub
    mEditCellTextBox.Locked = mTextEditionLocked
    mEditCellTextBox.Appearance = 0
    mEditCellTextBox.BorderStyle = 0
    mEditCellTextBox.Font.Bold = mGrid.CellFontBold
    mEditCellTextBox.Font.Name = mGrid.CellFontName
    mEditCellTextBox.Font.Size = mGrid.CellFontSize
    mEditCellTextBox.Font.Bold = mGrid.CellFontBold
    mEditCellTextBox.Font.Italic = mGrid.CellFontItalic
    mEditCellTextBox.Font.Underline = mGrid.CellFontUnderline
    mEditCellTextBox.Font.Strikethrough = mGrid.CellFontStrikeThrough
    mEditCellTextBox.ForeColor = mGrid.CellForeColor
'    Debug.Print mGrid.Row, mGrid.col, "&H" & Hex(mGrid.CellBackColor)
    mEditCellTextBox.BackColor = mGrid.CellBackColor
    
    iAuxLeft = (mGrid.GridLineWidth - 1)
    iAuxTop = mGrid.GridLineWidth
    If iAuxLeft < 2 Then iAuxLeft = 2
    If iAuxTop < 2 Then iAuxTop = 2
    mEditCellTextBox.Text = mGrid.TextMatrix(mEditCellRow, mEditCellCol)
    iStr = mEditCellTextBox.Text
    iStr = Replace(iStr, "?", "")
    iStr = Replace(iStr, " ", "")
    If iStr = "" Then
        RemoveEditCell
        GoTo Salida:
    End If
    
    iAuxWW = False
    If iWordWrap Then
        iAuxWW = SendMessageLong(mEditCellTextBox.hWnd, EM_GETLINECOUNT, 0&, 0&) > 1
        'Debug.Print SendMessageLong(mEditCellTextBox.hWnd, EM_GETLINECOUNT, 0&, 0&)
    End If
    If iAuxWW Then
        mEditCellTextBox.Move mGrid.Left + mGrid.ColPos(mEditCellCol) + iAuxLeft * Screen.TwipsPerPixelX, mGrid.Top + mGrid.RowPos(mEditCellRow) + iAuxTop * Screen.TwipsPerPixelY, mGrid.ColWidth(mEditCellCol) - (iAuxLeft * 2) * Screen.TwipsPerPixelX, mGrid.RowHeight(mEditCellRow) - (iAuxTop * 2) * Screen.TwipsPerPixelY
    Else
        Set iAuxFont = mGrid.Parent.Font
        Set mGrid.Parent.Font = mEditCellTextBox.Font
        iAuxHeight = mGrid.Parent.TextHeight(mGrid.TextMatrix(mEditCellRow, mEditCellCol)) + 2 * Screen.TwipsPerPixelY
        If iAuxHeight > mGrid.RowHeight(mEditCellRow) Then
            iAuxHeight = mGrid.RowHeight(mEditCellRow)
        End If
        Set mGrid.Parent.Font = iAuxFont
        
        Select Case FilterAlignment(mGrid.CellAlignment, efnFilterAligmentVertical)
            Case 0 ' top
                mEditCellTextBox.Move mGrid.Left + mGrid.ColPos(mEditCellCol) + iAuxLeft * Screen.TwipsPerPixelX, mGrid.Top + mGrid.RowPos(mEditCellRow) + iAuxTop * Screen.TwipsPerPixelY, mGrid.ColWidth(mEditCellCol) - (iAuxLeft * 2) * Screen.TwipsPerPixelX, iAuxHeight - (iAuxTop * 2) * Screen.TwipsPerPixelY
            Case 1 ' Center
                mEditCellTextBox.Move mGrid.Left + mGrid.ColPos(mEditCellCol) + iAuxLeft * Screen.TwipsPerPixelX, mGrid.Top + mGrid.RowPos(mEditCellRow) + (mGrid.RowHeight(mEditCellRow) - iAuxHeight) / 2, mGrid.ColWidth(mEditCellCol) - (iAuxLeft * 2) * Screen.TwipsPerPixelX, iAuxHeight - (iAuxTop * 2) * Screen.TwipsPerPixelY
            Case 2 ' Bottom
                mEditCellTextBox.Move mGrid.Left + mGrid.ColPos(mEditCellCol) + iAuxLeft * Screen.TwipsPerPixelX, mGrid.Top + mGrid.RowPos(mEditCellRow) + iAuxTop * Screen.TwipsPerPixelY + (mGrid.RowHeight(mEditCellRow) - iAuxHeight) - 2 * Screen.TwipsPerPixelY, mGrid.ColWidth(mEditCellCol) - (iAuxLeft * 2) * Screen.TwipsPerPixelX, iAuxHeight - (iAuxTop * 2) * Screen.TwipsPerPixelY
        End Select
    End If
    mEditCellTextBox.Visible = True
    mEditCellTextBox.ZOrder
    mEditCellTextBox.SelStart = 0
    mEditCellTextBox.SelLength = Len(mEditCellTextBox.Text)
    Set mTimerSetFocusToEditCellTextBox = New cTimer
    mTimerSetFocusToEditCellTextBox.Interval = 1
    AttachMessage Me, mEditCellTextBox.hWnd, WM_KILLFOCUS
    mEditCellTextBox.Tag = mGrid.HighLight
    mGrid.HighLight = 0
    
Salida:
    mGrid.col = iAuxCol
    mGrid.Row = iAuxRow
    mGrid.ColSel = iAuxColSel
    mGrid.RowSel = iAuxRowSel
    mGrid.Redraw = iAuxRedraw Or mGrid.Redraw
    
    mEditCellTopRow = mGrid.TopRow
    mEditCellLeftCol = mGrid.LeftCol
    
End Sub

Private Sub RemoveEditCell()
    mEditingCell = False
    On Error Resume Next
    DetachMessage Me, mEditCellTextBox.hWnd, WM_KILLFOCUS
    
'    If GetFocus = mEditCellTextBox.hWnd Then
    SetFocusAPI mGrid.hWnd
    mEditCellTextBox.Visible = False
    If mEditCellTextBoxOldParent <> 0 Then
        SetParent mEditCellTextBox.hWnd, mEditCellTextBoxOldParent
        mEditCellTextBoxOldParent = 0
    End If
'    End If
    If Not mTextEditionLocked Then
        If mEditCellTextBox.Text <> mGrid.TextMatrix(mEditCellRow, mEditCellCol) Then
            mGrid.TextMatrix(mEditCellRow, mEditCellCol) = mEditCellTextBox.Text
            ParentFlexFnObject.EventRaise "CellTextChange", mGrid.Name, mEditCellRow, mEditCellCol
            mGrid.HighLight = Val(mEditCellTextBox.Tag)
            Set mEditCellTextBox = Nothing
        End If
    End If
    If Not mUCControls Is Nothing Then
        mUCControls.Remove mGrid.Name & "_" & "EditCellTextBox"
    End If
    If Not mTimerSetFocusToEditCellTextBox Is Nothing Then
        mTimerSetFocusToEditCellTextBox.Interval = 0
        Set mTimerSetFocusToEditCellTextBox = Nothing
    End If
    If Not mTimerEditCell Is Nothing Then
        mTimerEditCell.Interval = 0
        Set mTimerEditCell = Nothing
    End If
End Sub

Private Function FilterAlignment(nFlexAlignment As Long, nFilter As efnFilterAligment, Optional nConsiderCurrentCell As Boolean = True) As Long
    If nFilter = efnFilterAligmentHorizontal Then
        Select Case nFlexAlignment
            Case flexAlignCenterBottom ' 5
                FilterAlignment = 2 ' Center
            Case flexAlignCenterCenter ' 4
                FilterAlignment = 2 ' Center
            Case flexAlignCenterTop ' 3
                FilterAlignment = 2 ' Center
            Case flexAlignGeneral ' 9
                If nConsiderCurrentCell Then
                    If IsNumeric(mGrid.Text) Then
                        FilterAlignment = 1 ' Right$
                    Else
                        FilterAlignment = 0 ' Left$
                    End If
                Else
                    FilterAlignment = 9
                End If
            Case flexAlignLeftBottom ' 2
                FilterAlignment = 0 ' Left$
            Case flexAlignLeftCenter  ' 1
                FilterAlignment = 0 ' Left$
            Case flexAlignLeftTop ' 0
                FilterAlignment = 0 ' Left$
            Case flexAlignRightBottom ' 8
                FilterAlignment = 1 ' Right$
            Case flexAlignRightCenter ' 7
                FilterAlignment = 1 ' Right$
            Case flexAlignRightTop ' 6
                FilterAlignment = 1 ' Right$
        End Select
    Else
        Select Case nFlexAlignment
            Case flexAlignCenterBottom ' 5
                FilterAlignment = 2 ' Bottom
            Case flexAlignCenterCenter ' 4
                FilterAlignment = 1 ' Center
            Case flexAlignCenterTop ' 3
                FilterAlignment = 0 ' Top
            Case flexAlignGeneral ' 9
                FilterAlignment = 1 ' center
            Case flexAlignLeftBottom ' 2
                FilterAlignment = 2 ' Bottom
            Case flexAlignLeftCenter  ' 1
                FilterAlignment = 1 ' Center
            Case flexAlignLeftTop ' 0
                FilterAlignment = 0 ' Top
            Case flexAlignRightBottom ' 8
                FilterAlignment = 2 ' Bottom
            Case flexAlignRightCenter ' 7
                FilterAlignment = 1 ' Center
            Case flexAlignRightTop ' 6
                FilterAlignment = 0 ' Top
        End Select
    End If
End Function

Public Sub UpdateOrderByColumn()
    If Not mGridValuesFromClientStored Then
        StoreGridValuesFromClient
    End If
    OrderGridByColumn mLastColumnClicked, mLastOrderDescending, True
End Sub

Public Property Get GetOrderColumn() As Long
    If Not mGridValuesFromClientStored Then
        StoreGridValuesFromClient
    End If
    GetOrderColumn = mLastColumnClicked
End Property

Public Property Get GetOrderColumnDescending() As Boolean
    If Not mGridValuesFromClientStored Then
        StoreGridValuesFromClient
    End If
    GetOrderColumnDescending = mLastOrderDescending
End Property

Public Property Get TimerCheckRedrawEnabled() As Boolean
    TimerCheckRedrawEnabled = Not mTimerCheckRedraw Is Nothing
End Property

Public Property Let TimerCheckRedrawEnabled(nValue As Boolean)
    If nValue <> TimerCheckRedrawEnabled Then
        If nValue Then
            Set mTimerCheckRedraw = New cTimer
            mTimerCheckRedraw.Interval = 100
        Else
            mTimerCheckRedraw.Interval = 0
            Set mTimerCheckRedraw = Nothing
        End If
    End If
End Property

Public Property Get EditingCell() As Boolean
    EditingCell = mEditingCell
End Property

Public Sub CopyTextEditingCell()
    Dim iTxt As TextBox
    Dim iUniClipboard As UniClipboard
    
    On Error GoTo Salida:
    Set iTxt = mEditCellTextBox
    If iTxt.SelLength > 0 Then
        Set iUniClipboard = New UniClipboard
        iUniClipboard.SetText iTxt.SelText
    End If
    
Salida:
End Sub


Public Property Let LastTextToFind(nText As String)
    mLastTextToFind = nText
End Property

Public Property Get LastTextToFind() As String
    LastTextToFind = mLastTextToFind
End Property
