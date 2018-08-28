VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlexFnObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Private Declare Function CLSIDFromProgID Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As Any) As Long
Private Declare Function StringFromCLSID Lib "ole32.dll" (pCLSID As Any, lpszProgID As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const SW_SHOWMAXIMIZED = 3

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Event TextFound(ByVal Row As Long, ByVal col As Long, ByVal Text As String)
Public Event BeforeOrderingByColumn(ByVal GridName As String, ByVal ColumnClicked As Long, ByRef OrderByThisColumn As Long, Descending As Boolean)
Public Event AfterOrderingByColumn(ByVal GridName As String, ByVal ColumnClicked As Long, ByVal OrderedByCol As Long, ByVal Descending As Boolean)
Public Event BeforePrintGrid(ByVal GridName As String)
Public Event StartPage(ByVal GridName As String)
Public Event EndPage(ByVal GridName As String)
Public Event StartDocument(ByVal GridName As String)
Public Event EndDocument(ByVal GridName As String)
Public Event UpdateUI()
Public Event WorkGridChange(GridName As String)
Public Event PersonalizeDefaultReportStyle(GridName As String, nRS As GridReportStyle)
Public Event BeforeCopyingToClipboard(ByVal GridName As String, ByRef TextBefore As String, ByRef TextAfter As String)
Public Event AfterSettingColumnsWidths(ByVal GridName As String, WidthsWereAdjusted As Boolean)
Public Event CellTextChange(ByVal GridName As String, Row As Long, col As Long)
Public Event GetUCControls(nUCC As Object)
Public Event BeforeTextEdit(ByVal GridName As String, Cancel As Boolean)
Public Event OrientationChange(ByVal NewOrientation As Long)
Public Event DocPrinted()

Private WithEvents mPrintFnObject As PrintFnObject
Attribute mPrintFnObject.VB_VarHelpID = -1
Private WithEvents mMouseWheelNotifierObject As MouseWheelNotifierObject
Attribute mMouseWheelNotifierObject.VB_VarHelpID = -1
Private mGridPrintingData As cGridPrintingData

Public Enum gfnOrientation
    gfnPrinterDefault = 0&
    gfnPortrait = 1&
    gfnLandscape = 2&
    gfnDecideOnReportWidth = 3&
End Enum

Public Enum gfnRememberUserPrintingPreferences
    gfnRememberByReportIDIfAvailable = 0&
    gfnRememberByGridName = 1&
    gfnRememberGlobally = 2&
    gfnDontRemember = 3&
End Enum

Public Enum gfnCopyToClipboardModeOptions
    gfnLetUserToSelect = 0&
    gfnSeparateWithTabs = 1&
    gfnSeparateWithSpacesForFixedFont = 2&
    gfnSeparateWithSpacesForVariableFont = 3&
    gfnSeparateWithSpecialCharacters = 4&
End Enum

Public Enum gfnAfterSaveActionSettings
    gfnDoNothing = 0&
    gfnShowConfirmationMessage = 1&
    gfnOfferToOpen = 2&
End Enum

Public Enum gfnScrollWithMouseWheelSettings
    gfnScrollDisabled = 0&
    gfnScrollEnabled = 1&
    gfnScrollOnlyOnFocus = 2&
End Enum

Private mReportID As String
Private mFileName As String
Private mHeading As String
Private mSubheading As String
Private mMiddleText As String
Private mFinalText As String
Private mDefaultFolderPath As String
Private mDefaultReportStyle As Long
Private mMinScalePercent As Long
Private mMaxScalePercent As Long

Private mCurrentGrid As Control
Private mCurrentGridHandler As cGridHandler
Private mCurrentGridNameWithParent As String

' print format settings
Private mRememberUserPrintingPreferences As gfnRememberUserPrintingPreferences

Private mGridReportStyle As New GridReportStyle
Private WithEvents mPrintGridFormatSettings As PrintGridFormatSettings
Attribute mPrintGridFormatSettings.VB_VarHelpID = -1

Private mOrientation As Long
Private mEnableAutoOrientation As Boolean

Private mPrintFormatSettingsLoaded As Boolean

Private mAuxAllowAutoOrientation As Boolean
Private mUserChangedScale As Boolean
Private mOrientationChangedToLandscapeAutomatically As Boolean
Private mDefaultReportOrientation As Long

Private mUsingDefaultReportStyle As Boolean

Private mSave_EnableAutoOrientation As Boolean
Private mSave_ScalePercent As Boolean
Private mSave_ColorMode  As Boolean
Private mSave_GridAlign  As Boolean
Private mSave_PrintPageNumbers  As Boolean
Private mSave_PageNumbersPosition  As Boolean
Private mSave_PageNumbersFormatLong  As Boolean
Private mSave_PageNumbersFont  As Boolean
Private mSave_PageNumbersForeColor  As Boolean
Private mSave_HeadingFont  As Boolean
Private mSave_HeadingFontColor  As Boolean
Private mSave_SubheadingFont  As Boolean
Private mSave_SubheadingFontColor  As Boolean
Private mSave_OtherTextsFont  As Boolean
Private mSave_OtherTextsFontColor  As Boolean

Private mCopyToClipboardMode As Long
Private mSpecialSeparatorCharacters As String
Private mAfterSaveAction As Long
Private mMergeCellsExcel As Boolean

Private mScrollWithMouseWheel As Long
Private mIgnoreEmptyRowsAtTheEnd  As Boolean

Private mEnableOrderByColumns As Boolean
Private mStretchColumnsWidthsToFill As Boolean
Private mGridsFlatAppearance As Boolean
Private mInitialOrderColumn As Long
Private mInitialOrderDescending As Boolean
Private mSameDataGroupedInColumns As Boolean
Private mGridName As String
Private mBorderColor As Long
Private mBorderWidth As Long
Private mShowToolTipsOnLongerCellTexts As Boolean
Private mDoNotRememberOrder As Boolean
Private mShowToolTipsForOrderColumns As Boolean
Private mAllowTextEdition As Boolean
Private mTextEditionLocked As Boolean
Private mShowCopyConfirmationMessage As Boolean
Private mCopyConfirmationMessage As String
Private mAutoSelect125PercentScaleOnSmallGrids As Boolean

Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private mGridParent As Object
Private mGridExplicitelySet As Boolean

Private mGridPrev As Object
Private mReportIDPrev As String
Private mHeadingPrev As String
Private mSubheadingPrev As String
Private mMiddleTextPrev As String
Private mFinalTextPrev As String
Private mOrientationPrev As Long
Private mScalePercentPrev As Long
Private mFileNamePrev As String
Private mDefaultFolderPathPrev As String
Private mCopyToClipboardModePrev As Long
Private mSpecialSeparatorCharactersPrev As String

Private mLastReportKeyname As String
Private mLastReporScalePercent As Long
Private mGridHandlersCollection As Collection
Private mFormHwnd As Long
Private mFormSubclassed As Boolean
Private mGridControlsStored As Boolean
Private mStoringGridControls As Boolean
Private WithEvents mTimerFormSet As cTimer
Attribute mTimerFormSet.VB_VarHelpID = -1
Private mLastControlsCount As Long
Private WithEvents TimerUpdateUI As cTimer
Attribute TimerUpdateUI.VB_VarHelpID = -1
Private mAmbientUserMode As Boolean
Private mHwndGridHandledIndividually As Long
Private mGridExplicitelySetIsTypeSupported As Boolean

Private Property Set CurrentGrid(nGrid As Object)
    If Not nGrid Is mCurrentGrid Then
        Set mCurrentGrid = nGrid
        If Not mCurrentGrid Is Nothing Then
            Set mCurrentGridHandler = GetGridHandler(mCurrentGrid)
            mCurrentGridNameWithParent = Base64Encode(ControlNameWithParent(mCurrentGrid))
        Else
            Set mCurrentGridHandler = Nothing
            mCurrentGridNameWithParent = ""
        End If
    End If
End Property

Private Property Get CurrentGrid() As Object
    Set CurrentGrid = mCurrentGrid
End Property


Public Property Set Grid(nGrid As Object)
    If Not CurrentGrid Is nGrid Then
        If mGridExplicitelySet Then
            If Not CurrentGrid Is Nothing Then
                RemoveGridHandler CurrentGrid
            End If
        End If
        Set CurrentGrid = nGrid
        mGridExplicitelySet = Not CurrentGrid Is Nothing
        If mGridExplicitelySet Then
            SetGridAsHandledIndividually
        End If
        If Not CurrentGrid Is Nothing Then
            AddGridHandler CurrentGrid
            If mScrollWithMouseWheel <> 0 Then
                Set mMouseWheelNotifierObject = New MouseWheelNotifierObject
                mMouseWheelNotifierObject.SetForm , GetParentFormHwnd(CurrentGrid.hWnd)
            End If
        End If
    End If
End Property

Public Property Get Grid() As Object
    Set Grid = CurrentGrid
End Property


Public Property Get ReportID() As String
    ReportID = mReportID
End Property

Public Property Let ReportID(nValue As String)
    Dim iGH As cGridHandler
    
    If nValue <> mReportID Then
        mReportID = nValue
        ResetReportSpecificVariables
        For Each iGH In mGridHandlersCollection
            iGH.ReportID = mReportID
        Next
    End If

End Property


Public Property Get FileName() As String
    FileName = mFileName
End Property

Public Property Let FileName(nValue As String)
    If nValue <> mFileName Then
        mFileName = nValue
    End If
End Property


Public Property Get Heading() As String
    Heading = mHeading
End Property

Public Property Let Heading(nValue As String)
    If nValue <> mHeading Then
        mHeading = nValue
    End If
End Property


Public Property Get Subheading() As String
    Subheading = mSubheading
End Property

Public Property Let Subheading(nValue As String)
    If nValue <> mSubheading Then
        mSubheading = nValue
    End If
End Property


Public Property Get MiddleText() As String
    MiddleText = mMiddleText
End Property

Public Property Let MiddleText(nValue As String)
    If nValue <> mMiddleText Then
        mMiddleText = nValue
    End If
End Property


Public Property Get FinalText() As String
    FinalText = mFinalText
End Property

Public Property Let FinalText(nValue As String)
    If nValue <> mFinalText Then
        mFinalText = nValue
    End If
End Property


Public Property Get DefaultFolderPath() As String
    DefaultFolderPath = mDefaultFolderPath
End Property

Public Property Let DefaultFolderPath(nValue As String)
    If nValue <> mDefaultFolderPath Then
        mDefaultFolderPath = nValue
    End If
End Property

Private Sub Class_Initialize()
    mDefaultReportStyle = 2
    
    Set mPrintFnObject = New PrintFnObject
    Set mPrintGridFormatSettings = New PrintGridFormatSettings
    
    
    mMinScalePercent = cPrintPreviewDefaultMinScale
    mMaxScalePercent = cPrintPreviewDefaultMaxScale
    mPrintGridFormatSettings.ScalePercent = 100
    
    mPrintFnObject.FormatButtonVisible = True
    mPrintFnObject.AllowUserChangeScale = True
    
    Set mGridReportStyle = GetGridReportStyle("Default" & CStr(mDefaultReportStyle))
    
    mCopyToClipboardMode = gfnSeparateWithTabs
    
    mScrollWithMouseWheel = gfnScrollEnabled
    mIgnoreEmptyRowsAtTheEnd = True
    mSpecialSeparatorCharacters = "|"
    mAfterSaveAction = 2
    mMergeCellsExcel = True
    
    Set mGridHandlersCollection = New Collection
    
    mStretchColumnsWidthsToFill = True
    mEnableOrderByColumns = False
    mGridsFlatAppearance = False
    mInitialOrderColumn = -1
    mInitialOrderDescending = False
    
    mBorderColor = &HC0C0C0
    mBorderWidth = 1
    
    mShowCopyConfirmationMessage = True
    mCopyConfirmationMessage = "Se copió el texto"
    mAutoSelect125PercentScaleOnSmallGrids = False
    
    mPrintFnObject.PrintPrevUseAltScaleIcons = True
    mOrientation = gfnDecideOnReportWidth
End Sub

Private Sub Class_Terminate()
    FinishForm
    Set mGridHandlersCollection = Nothing
    If Not TimerUpdateUI Is Nothing Then
        TimerUpdateUI.Interval = 0
        Set TimerUpdateUI = Nothing
    End If
    RemoveGridAsHandledIndividually
End Sub

Private Sub mForm_Load()
    StoreGridControls
End Sub

Private Sub mForm_Unload(Cancel As Integer)
    If Cancel = 0 Then
        FinishForm
    End If
End Sub

Private Sub mPrintFnObject_AfterShowingPageSetupDialog()
    If mPrintGridFormatSettings.ColorMode <> mPrintFnObject.ColorMode Then
        mPrintGridFormatSettings.ColorMode = mPrintFnObject.ColorMode
        If (mPrintFnObject.DefaultColorMode <> 0) Then
            If mPrintFnObject.ColorMode <> mPrintFnObject.DefaultColorMode Then
                mPrintGridFormatSettings.ColorMode = mPrintFnObject.ColorMode
                mSave_ColorMode = True
            End If
        End If
    End If
    SaveMarginSettings
End Sub

Private Sub mPrintFnObject_AfterShowingPrinterDialog()
    If mPrintGridFormatSettings.ColorMode <> mPrintFnObject.ColorMode Then
        mPrintGridFormatSettings.ColorMode = mPrintFnObject.ColorMode
        If (mPrintFnObject.DefaultColorMode <> 0) Then
            If mPrintFnObject.ColorMode <> mPrintFnObject.DefaultColorMode Then
                mPrintGridFormatSettings.ColorMode = mPrintFnObject.ColorMode
                mSave_ColorMode = True
            End If
        End If
    End If
    SaveUserSettings
    If Not mUsingDefaultReportStyle Then
        SaveGridReportStyle
    End If
End Sub

Private Sub mPrintFnObject_DocPrinted(ByVal DocKey As String)
    RaiseEvent DocPrinted
End Sub

Private Sub mPrintFnObject_PrepareDoc(Cancel As Boolean, ByVal DocKey As String)
    DrawGrid Cancel
End Sub

Private Sub mPrintFnObject_ScaleChange(NewScalePercent As Integer)
    Dim iDSPLarger As Boolean
    Dim iDSPSmaller As Boolean
    
    If mPrintGridFormatSettings.ScalePercent < NewScalePercent Then
        iDSPLarger = True
    End If
    If mPrintGridFormatSettings.ScalePercent > NewScalePercent Then
        iDSPSmaller = True
    End If
    
    mPrintGridFormatSettings.ScalePercent = NewScalePercent
    
    If mOrientationChangedToLandscapeAutomatically Then
        If mEnableAutoOrientation And iDSPLarger Then
            If mPrintFnObject.Orientation = vbPRORPortrait Then
                mAuxAllowAutoOrientation = True
            End If
        End If
        If mEnableAutoOrientation And iDSPSmaller Then
            If mPrintFnObject.Orientation = vbPRORLandscape Then
                mPrintFnObject.Orientation = vbPRORPortrait
                mDefaultReportOrientation = vbPRORPortrait
                mAuxAllowAutoOrientation = True
            End If
        End If
    End If
    
    mPrintFnObject.ScalePercent = mPrintGridFormatSettings.ScalePercent
    mUserChangedScale = True
    If (RememberUserPrintingPreferences = gfnRememberByReportIDIfAvailable) Or (RememberUserPrintingPreferences = gfnRememberByGridName) Then
        mSave_ScalePercent = True
    End If
    
End Sub

' show format options
Private Sub mPrintFnObject_FormatOptionsClick(ByRef Canceled As Boolean)
    Dim ifrmPrintGridFormatOptions As New frmPrintGridFormatOptions
    Dim iDSPLarger As Boolean
    Dim iDSPSmaller As Boolean
    Dim iEAOPrev As Boolean
    
    Dim iEnableAutoOrientationPrev As Boolean
    Dim iScalePercent As Long
    Dim iColorMode  As cdeColorModeConstants
    Dim iGridAlign  As Long
    Dim iPrintPageNumbers  As Boolean
    Dim iPageNumbersPosition  As Long
    Dim iPageNumbersFormatLong  As Long
    Dim iPageNumbersFont  As StdFont
    Dim iPageNumbersForeColor  As Long
    Dim iHeadingFont  As StdFont
    Dim iHeadingFontColor  As Long
    Dim iSubheadingFont  As StdFont
    Dim iSubheadingFontColor  As Boolean
    Dim iOtherTextsFont  As StdFont
    Dim iOtherTextsFontColor  As Boolean
    
    EnsurePrintFormatSettingsLoaded
    Set ifrmPrintGridFormatOptions.Grid = mCurrentGrid
    Set ifrmPrintGridFormatOptions.PrintFnObject = mPrintFnObject
    Set ifrmPrintGridFormatOptions.FlexFnObject = Me
    
    iEnableAutoOrientationPrev = mEnableAutoOrientation
    iScalePercent = mPrintGridFormatSettings.ScalePercent
    
    iColorMode = mPrintGridFormatSettings.ColorMode
    iGridAlign = mPrintGridFormatSettings.GridAlign
    iPrintPageNumbers = mPrintGridFormatSettings.PrintPageNumbers
    iPageNumbersPosition = mPrintGridFormatSettings.PageNumbersPosition
    iPageNumbersFormatLong = mPrintGridFormatSettings.PageNumbersFormatLong
    Set iPageNumbersFont = CloneFont(mPrintGridFormatSettings.PageNumbersFont)
    iPageNumbersForeColor = mPrintGridFormatSettings.PageNumbersForeColor
    Set iHeadingFont = CloneFont(mPrintGridFormatSettings.HeadingFont)
    iHeadingFontColor = mPrintGridFormatSettings.HeadingFontColor
    Set iSubheadingFont = CloneFont(mPrintGridFormatSettings.SubheadingFont)
    iSubheadingFontColor = mPrintGridFormatSettings.SubheadingFontColor
    Set iOtherTextsFont = CloneFont(mPrintGridFormatSettings.OtherTextsFont)
    iOtherTextsFontColor = mPrintGridFormatSettings.OtherTextsFontColor
    
    ifrmPrintGridFormatOptions.MinScalePercent = mMinScalePercent
    ifrmPrintGridFormatOptions.MaxScalePercent = mMaxScalePercent
    ifrmPrintGridFormatOptions.ScalePercent = mPrintGridFormatSettings.ScalePercent
    If Not mPrintFnObject.AllowUserChangeScale Then
        ifrmPrintGridFormatOptions.cboScalePercent.Visible = False
        ifrmPrintGridFormatOptions.lblScalePercent.Visible = False
    End If
    ifrmPrintGridFormatOptions.ColorMode = mPrintGridFormatSettings.ColorMode
    ifrmPrintGridFormatOptions.GridAlign = mPrintGridFormatSettings.GridAlign
    ifrmPrintGridFormatOptions.PrintPageNumbers = mPrintGridFormatSettings.PrintPageNumbers
    ifrmPrintGridFormatOptions.PageNumbersPosition = mPrintGridFormatSettings.PageNumbersPosition
    ifrmPrintGridFormatOptions.PageNumbersFormatLong = mPrintGridFormatSettings.PageNumbersFormatLong
    Set ifrmPrintGridFormatOptions.PageNumbersFont = mPrintGridFormatSettings.PageNumbersFont
    ifrmPrintGridFormatOptions.PageNumbersForeColor = mPrintGridFormatSettings.PageNumbersForeColor
    Set ifrmPrintGridFormatOptions.HeadingFont = mPrintGridFormatSettings.HeadingFont
    ifrmPrintGridFormatOptions.HeadingFontColor = mPrintGridFormatSettings.HeadingFontColor
    Set ifrmPrintGridFormatOptions.SubheadingFont = mPrintGridFormatSettings.SubheadingFont
    ifrmPrintGridFormatOptions.SubheadingFontColor = mPrintGridFormatSettings.SubheadingFontColor
    Set ifrmPrintGridFormatOptions.OtherTextsFont = mPrintGridFormatSettings.OtherTextsFont
    ifrmPrintGridFormatOptions.OtherTextsFontColor = mPrintGridFormatSettings.OtherTextsFontColor
    ifrmPrintGridFormatOptions.DisplayHeadingFont = Trim$(mHeading) <> ""
    ifrmPrintGridFormatOptions.DisplaySubheadingFont = Trim$(mSubheading) <> ""
    ifrmPrintGridFormatOptions.DisplayOtherTextsFont = (Trim$(mMiddleText) <> "") Or (Trim$(mFinalText) <> "")
    If ifrmPrintGridFormatOptions.DisplayHeadingFont Then
        ifrmPrintGridFormatOptions.HeadingSampleText = Trim$(mHeading)
    End If
    If ifrmPrintGridFormatOptions.DisplaySubheadingFont Then
        ifrmPrintGridFormatOptions.SubheadingSampleText = Trim$(mSubheading)
    End If
    If ifrmPrintGridFormatOptions.DisplayOtherTextsFont Then
        If Trim$(mMiddleText) <> "" Then
            ifrmPrintGridFormatOptions.OtherTextsSampleText = Trim$(mMiddleText)
        Else
            ifrmPrintGridFormatOptions.OtherTextsSampleText = Trim$(mFinalText)
        End If
    End If
    Set ifrmPrintGridFormatOptions.GridReportStyle = mGridReportStyle
    
    ifrmPrintGridFormatOptions.EnableAutoOrientation = mEnableAutoOrientation
    
    ifrmPrintGridFormatOptions.Show vbModal
    
    If ifrmPrintGridFormatOptions.OKPressed And ifrmPrintGridFormatOptions.Changed Then
        If mPrintGridFormatSettings.ScalePercent < ifrmPrintGridFormatOptions.ScalePercent Then
            iDSPLarger = True
        End If
        If mPrintGridFormatSettings.ScalePercent > ifrmPrintGridFormatOptions.ScalePercent Then
            iDSPSmaller = True
        End If
        
        mPrintGridFormatSettings.ScalePercent = ifrmPrintGridFormatOptions.ScalePercent
        mPrintGridFormatSettings.ColorMode = ifrmPrintGridFormatOptions.ColorMode
        mPrintGridFormatSettings.GridAlign = ifrmPrintGridFormatOptions.GridAlign
        mPrintGridFormatSettings.PrintPageNumbers = ifrmPrintGridFormatOptions.PrintPageNumbers
        mPrintGridFormatSettings.PageNumbersPosition = ifrmPrintGridFormatOptions.PageNumbersPosition
        mPrintGridFormatSettings.PageNumbersFormatLong = ifrmPrintGridFormatOptions.PageNumbersFormatLong
        Set mPrintGridFormatSettings.PageNumbersFont = ifrmPrintGridFormatOptions.PageNumbersFont
        mPrintGridFormatSettings.PageNumbersForeColor = ifrmPrintGridFormatOptions.PageNumbersForeColor
        Set mPrintGridFormatSettings.HeadingFont = ifrmPrintGridFormatOptions.HeadingFont
        mPrintGridFormatSettings.HeadingFontColor = ifrmPrintGridFormatOptions.HeadingFontColor
        Set mPrintGridFormatSettings.SubheadingFont = ifrmPrintGridFormatOptions.SubheadingFont
        mPrintGridFormatSettings.SubheadingFontColor = ifrmPrintGridFormatOptions.SubheadingFontColor
        Set mPrintGridFormatSettings.OtherTextsFont = ifrmPrintGridFormatOptions.OtherTextsFont
        mPrintGridFormatSettings.OtherTextsFontColor = ifrmPrintGridFormatOptions.OtherTextsFontColor
        
        Set mGridReportStyle = ifrmPrintGridFormatOptions.GridReportStyle
        mUsingDefaultReportStyle = mGridReportStyle.Tag <> "Save"
        
        iEAOPrev = mEnableAutoOrientation
        mEnableAutoOrientation = ifrmPrintGridFormatOptions.EnableAutoOrientation
        If (mEnableAutoOrientation And iDSPLarger) Or (mEnableAutoOrientation And Not iEAOPrev) Then
            If mPrintFnObject.Orientation = vbPRORPortrait Then
                mAuxAllowAutoOrientation = True
            End If
        End If
        If (mEnableAutoOrientation And iDSPSmaller) Or (mEnableAutoOrientation And Not iEAOPrev) Then
            If mPrintFnObject.Orientation = vbPRORLandscape Then
                mPrintFnObject.Orientation = vbPRORPortrait
                mDefaultReportOrientation = vbPRORPortrait
                mAuxAllowAutoOrientation = True
            End If
        End If
        
        mSave_EnableAutoOrientation = mSave_EnableAutoOrientation Or (iEnableAutoOrientationPrev <> mEnableAutoOrientation)
        mSave_ScalePercent = mSave_ScalePercent Or (iScalePercent <> mPrintGridFormatSettings.ScalePercent)
        mUserChangedScale = mUserChangedScale Or (iScalePercent <> mPrintGridFormatSettings.ScalePercent)
        mSave_ColorMode = mSave_ColorMode Or (iColorMode <> mPrintGridFormatSettings.ColorMode)
        mSave_GridAlign = mSave_GridAlign Or (iGridAlign <> mPrintGridFormatSettings.GridAlign)
        mSave_PrintPageNumbers = mSave_PrintPageNumbers Or (iPrintPageNumbers <> mPrintGridFormatSettings.PrintPageNumbers)
        mSave_PageNumbersPosition = mSave_PageNumbersPosition Or (iPageNumbersPosition <> mPrintGridFormatSettings.PageNumbersPosition)
        mSave_PageNumbersFormatLong = mSave_PageNumbersFormatLong Or (iPageNumbersFormatLong <> mPrintGridFormatSettings.PageNumbersFormatLong)
        mSave_PageNumbersFont = mSave_PageNumbersFont Or (Not FontsAreEqual(iPageNumbersFont, mPrintGridFormatSettings.PageNumbersFont))
        mSave_PageNumbersForeColor = mSave_PageNumbersForeColor Or (iPageNumbersForeColor <> mPrintGridFormatSettings.PageNumbersForeColor)
        mSave_HeadingFont = mSave_HeadingFont Or (Not FontsAreEqual(iHeadingFont, mPrintGridFormatSettings.HeadingFont))
        mSave_HeadingFontColor = mSave_HeadingFontColor Or (iHeadingFontColor <> mPrintGridFormatSettings.HeadingFontColor)
        mSave_SubheadingFont = mSave_SubheadingFont Or (Not FontsAreEqual(iSubheadingFont, mPrintGridFormatSettings.SubheadingFont))
        mSave_SubheadingFontColor = mSave_SubheadingFontColor Or (iSubheadingFontColor <> mPrintGridFormatSettings.SubheadingFontColor)
        mSave_OtherTextsFont = mSave_OtherTextsFont Or (Not FontsAreEqual(iOtherTextsFont, mPrintGridFormatSettings.OtherTextsFont))
        mSave_OtherTextsFontColor = mSave_OtherTextsFontColor Or (iOtherTextsFontColor <> mPrintGridFormatSettings.OtherTextsFontColor)
        
        SetFormatOptionsSettingsOnPrintFnObject
    Else
        Canceled = True
    End If
    
End Sub

Public Sub PrintNow(Optional nGrid As Object, Optional nGridName, Optional nReportID, Optional nHeading, Optional nSubheading, Optional nMiddleText, Optional nFinalText, Optional nOrientation As gfnOrientation = -1, Optional nScalePercent)
    ShowPrint nGrid, nGridName, nReportID, nHeading, nSubheading, nMiddleText, nFinalText, nOrientation, nScalePercent, True
End Sub

Public Sub ShowPrint(Optional nGrid As Object, Optional nGridName, Optional nReportID, Optional nHeading, Optional nSubheading, Optional nMiddleText, Optional nFinalText, Optional nOrientation As gfnOrientation = -1, Optional nScalePercent, Optional nPrintWithDefaultSettings As Boolean)
    
    SavePreviousGeneralSettings
    SetCurrentGrid nGrid, nGridName
    If CurrentGrid Is Nothing Then
        MsgBox "Grid not found"
        RestorePreviousGeneralSettings
        Exit Sub
    End If
    
    If Not IsMissing(nReportID) Then
        ReportID = nReportID
    End If
    If Not IsMissing(nHeading) Then
        Heading = nHeading
    End If
    If Not IsMissing(nSubheading) Then
        Subheading = nSubheading
    End If
    If Not IsMissing(nMiddleText) Then
        MiddleText = nMiddleText
    End If
    If Not IsMissing(nFinalText) Then
        FinalText = nFinalText
    End If
    If nOrientation <> -1 Then
        Orientation = nOrientation
    End If
    If Not IsMissing(nScalePercent) Then
        ScalePercent = nScalePercent
    End If
    
    If Keyname <> mLastReportKeyname Then
        ResetReportSpecificVariables
    Else
        If IsMissing(nScalePercent) Then
            If mLastReporScalePercent <> 0 Then
                ScalePercent = mLastReporScalePercent
            End If
        End If
    End If
    mLastReportKeyname = Keyname
    
    EnsurePrintFormatSettingsLoaded
    
    If nPrintWithDefaultSettings Then
        mPrintFnObject.PrintNow
    Else
        mPrintFnObject.ShowPrint
    End If
    
    Set mGridPrintingData = Nothing
    
    mLastReporScalePercent = ScalePercent
    
    RestorePreviousGeneralSettings
    ResetPrinter2
End Sub
    
Public Sub ShowPrintPreview(Optional nGrid As Object, Optional nGridName, Optional nReportID, Optional nHeading, Optional nSubheading, Optional nMiddleText, Optional nFinalText, Optional nOrientation As gfnOrientation = -1, Optional nScalePercent)
    
    SavePreviousGeneralSettings
    SetCurrentGrid nGrid, nGridName
    If CurrentGrid Is Nothing Then
        MsgBox "Grid not found"
        RestorePreviousGeneralSettings
        Exit Sub
    End If
    
    If Not IsMissing(nReportID) Then
        ReportID = nReportID
    End If
    If Not IsMissing(nHeading) Then
        Heading = nHeading
    End If
    If Not IsMissing(nSubheading) Then
        Subheading = nSubheading
    End If
    If Not IsMissing(nMiddleText) Then
        MiddleText = nMiddleText
    End If
    If Not IsMissing(nFinalText) Then
        FinalText = nFinalText
    End If
    If nOrientation <> -1 Then
        Orientation = nOrientation
    End If
    If Not IsMissing(nScalePercent) Then
        ScalePercent = nScalePercent
    End If
    
    If Keyname <> mLastReportKeyname Then
        ResetReportSpecificVariables
    Else
        If IsMissing(nScalePercent) Then
            If mLastReporScalePercent <> 0 Then
                ScalePercent = mLastReporScalePercent
            End If
        End If
    End If
    mLastReportKeyname = Keyname
    
    EnsurePrintFormatSettingsLoaded
    
    mPrintFnObject.MinScalePercent = mMinScalePercent
    mPrintFnObject.MaxScalePercent = mMaxScalePercent
    
    mPrintFnObject.ShowPrintPreview
    
    Set mGridPrintingData = Nothing
    
    mLastReporScalePercent = ScalePercent
    
    RestorePreviousGeneralSettings
    ResetPrinter2
End Sub

Public Sub SaveAsExcelFile(Optional nGrid As Object, Optional nGridName, Optional nReportID, Optional nFileName, Optional nHeading, Optional nSubheading, Optional nMiddleText, Optional nFinalText, Optional nDefaultFolderPath)
    SavePreviousGeneralSettings
    SetCurrentGrid nGrid, nGridName
    If CurrentGrid Is Nothing Then
        MsgBox "Grid not found"
        RestorePreviousGeneralSettings
        Exit Sub
    End If
    
    If Not IsMissing(nDefaultFolderPath) Then
        DefaultFolderPath = nDefaultFolderPath
    End If
    If Not IsMissing(nReportID) Then
        ReportID = nReportID
    End If
    If Not IsMissing(nFileName) Then
        FileName = nFileName
    End If
    If Not IsMissing(nHeading) Then
        Heading = nHeading
    End If
    If Not IsMissing(nSubheading) Then
        Subheading = nSubheading
    End If
    If Not IsMissing(nMiddleText) Then
        MiddleText = nMiddleText
    End If
    If Not IsMissing(nFinalText) Then
        FinalText = nFinalText
    End If
    
    Call SaveGridAsExcelFile
    
    RestorePreviousGeneralSettings
End Sub

Public Sub CopyToClipboard(Optional nGrid As Object, Optional nGridName, Optional nCopyToClipboardMode As gfnCopyToClipboardModeOptions = -1, Optional nSpecialSeparatorCharacters)
    Dim iFrm As frmClipboardCopiedMessage
    Dim iCopied As Boolean
    
    SavePreviousGeneralSettings
    SetCurrentGrid nGrid, nGridName
    If CurrentGrid Is Nothing Then
        MsgBox "Grid not found"
        RestorePreviousGeneralSettings
        Exit Sub
    End If
    
    If nCopyToClipboardMode <> -1 Then
        mCopyToClipboardMode = nCopyToClipboardMode
    End If
    If Not IsMissing(nSpecialSeparatorCharacters) Then
        mSpecialSeparatorCharacters = nSpecialSeparatorCharacters
    End If
    
    iCopied = CopyGridToClipboard
    
    RestorePreviousGeneralSettings
    
    If iCopied And mShowCopyConfirmationMessage Then
        Set iFrm = New frmClipboardCopiedMessage
        iFrm.lblMessage.Caption = mCopyConfirmationMessage
        iFrm.ShowMessage
        If IsFormLoaded(iFrm) Then
            Unload iFrm
        End If
        Set iFrm = Nothing
    End If
End Sub

Public Sub FindText(Optional nGrid As Object, Optional nGridName, Optional TextToFind, Optional nFindNext As Boolean)
    SavePreviousGeneralSettings
    SetCurrentGrid nGrid, nGridName
    If CurrentGrid Is Nothing Then
        MsgBox "Grid not found"
        RestorePreviousGeneralSettings
        Exit Sub
    End If
    
    FindTextInGrid TextToFind, nFindNext
    
    RestorePreviousGeneralSettings
End Sub

Public Sub ShowConfigColumns(Optional nGrid As Object, Optional nGridName)
    SavePreviousGeneralSettings
    SetCurrentGrid nGrid, nGridName
    If CurrentGrid Is Nothing Then
        MsgBox "Grid not found"
        RestorePreviousGeneralSettings
        Exit Sub
    End If
    
    mCurrentGridHandler.ShowConfigGridColumns
    
    RestorePreviousGeneralSettings
End Sub

Public Property Let RememberUserPrintingPreferences(nValue As gfnRememberUserPrintingPreferences)
    Dim iGH As cGridHandler
    
    If nValue <> mRememberUserPrintingPreferences Then
        mRememberUserPrintingPreferences = nValue
        For Each iGH In mGridHandlersCollection
            iGH.RememberUserPrintingPreferences = mRememberUserPrintingPreferences
        Next
    End If
End Property

Public Property Get RememberUserPrintingPreferences() As gfnRememberUserPrintingPreferences
    RememberUserPrintingPreferences = mRememberUserPrintingPreferences
End Property


Private Sub LoadUserSettings()
    Dim iStyleID As String
    Dim iAuxLng As Long
    Dim iColor1 As Long
    Dim iColor2 As Long
    Dim iStr As String
    Dim iUnits As Long
    
    If RememberUserPrintingPreferences = gfnDontRemember Then Exit Sub
    If mCurrentGrid Is Nothing Then Exit Sub
    
    iAuxLng = GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_ScalePercent", 0)
    If iAuxLng <> 0 Then
        mPrintGridFormatSettings.ScalePercent = iAuxLng
    End If
    If (RememberUserPrintingPreferences = gfnRememberByReportIDIfAvailable) Or (RememberUserPrintingPreferences = gfnRememberByGridName) Then
        mUserChangedScale = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_UserChangedScale", 0))
    End If
    mPrintGridFormatSettings.ColorMode = CLng(GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_ColorMode", mPrintGridFormatSettings.ColorMode))
    mPrintGridFormatSettings.GridAlign = GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_GridAlign", mPrintGridFormatSettings.GridAlign)
    mPrintGridFormatSettings.PrintPageNumbers = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_PrintPageNumbers", CLng(mPrintGridFormatSettings.PrintPageNumbers)))
    mPrintGridFormatSettings.PageNumbersPosition = GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_PageNumbersPosition", mPrintGridFormatSettings.PageNumbersPosition)
    mPrintGridFormatSettings.PageNumbersFormatLong = GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_PageNumbersFormatLong", mPrintGridFormatSettings.PageNumbersFormatLong)
    Set mPrintGridFormatSettings.PageNumbersFont = GetSettingFont(AppNameForRegistry, "PrintingSettings", Keyname & "_PageNumbersFont", mPrintGridFormatSettings.PageNumbersFont)
    mPrintGridFormatSettings.PageNumbersForeColor = GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_PageNumbersForeColor", mPrintGridFormatSettings.PageNumbersForeColor)
    Set mPrintGridFormatSettings.HeadingFont = GetSettingFont(AppNameForRegistry, "PrintingSettings", Keyname & "_HeadingFont", mPrintGridFormatSettings.HeadingFont)
    mPrintGridFormatSettings.HeadingFontColor = GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_HeadingFontColor", mPrintGridFormatSettings.HeadingFontColor)
    Set mPrintGridFormatSettings.SubheadingFont = GetSettingFont(AppNameForRegistry, "PrintingSettings", Keyname & "_SubheadingFont", mPrintGridFormatSettings.SubheadingFont)
    mPrintGridFormatSettings.SubheadingFontColor = GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_SubheadingFontColor", mPrintGridFormatSettings.SubheadingFontColor)
    Set mPrintGridFormatSettings.OtherTextsFont = GetSettingFont(AppNameForRegistry, "PrintingSettings", Keyname & "_OtherTextsFont", mPrintGridFormatSettings.OtherTextsFont)
    mPrintGridFormatSettings.OtherTextsFontColor = GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_OtherTextsFontColor", mPrintGridFormatSettings.OtherTextsFontColor)
    
    iStyleID = GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_GridReportStyle", "")
    If iStyleID = "" Then
        mUsingDefaultReportStyle = True
        Set mGridReportStyle = GetGridReportStyle("Default" & CStr(mDefaultReportStyle))
        mGridReportStyle.Tag = "DoNotSave"
        If mCurrentGrid.GridLines = 0 Then
            mGridReportStyle.PrintRowsLines = False
            mGridReportStyle.PrintColumnsDataLines = False
            mGridReportStyle.RowsLinesColor = vbWhite
        End If
        iColor1 = mCurrentGrid.GridColor
        iColor2 = mCurrentGrid.BackColor
        TranslateColor iColor1, 0, iColor1
        TranslateColor iColor2, 0, iColor2
        If iColor1 = iColor2 Then
            mGridReportStyle.PrintRowsLines = False
            mGridReportStyle.PrintColumnsDataLines = False
            mGridReportStyle.RowsLinesColor = vbWhite
        End If
        If Not mCurrentGrid Is Nothing Then
            iStr = mCurrentGrid.Name
        End If
        RaiseEvent PersonalizeDefaultReportStyle(iStr, mGridReportStyle)
    Else
        Set mGridReportStyle = GetGridReportStyle(iStyleID)
    End If
    If mGridReportStyle Is Nothing Then
        Set mGridReportStyle = GetGridReportStyle("Default" & CStr(mDefaultReportStyle))
    End If
    
    If mOrientation = gfnDecideOnReportWidth Then
        iAuxLng = GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_EnableAutoOrientation", 20)
        If iAuxLng <> 20 Then
            mEnableAutoOrientation = CBool(iAuxLng)
        Else
            mEnableAutoOrientation = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "Global" & "_EnableAutoOrientation", (CLng(mOrientation = gfnDecideOnReportWidth))))
        End If
        mAuxAllowAutoOrientation = mEnableAutoOrientation
        If RememberUserPrintingPreferences <> gfnRememberGlobally Then
            If Not mEnableAutoOrientation Then
                mPrintFnObject.Orientation = GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_Orientation", mPrintFnObject.Orientation)
                mDefaultReportOrientation = mPrintFnObject.Orientation
            End If
        End If
    End If
    
    iUnits = mPrintFnObject.Units
    mPrintFnObject.Units = vbMillimeters
    mPrintFnObject.LeftMargin = GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_LeftMargin", mPrintFnObject.LeftMargin)
    mPrintFnObject.TopMargin = GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_TopMargin", mPrintFnObject.TopMargin)
    mPrintFnObject.RightMargin = GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_RightMargin", mPrintFnObject.RightMargin)
    mPrintFnObject.BottomMargin = GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_BottomMargin", mPrintFnObject.BottomMargin)
    mPrintFnObject.Units = iUnits
End Sub

Private Sub EnsurePrintFormatSettingsLoaded()
    If Not mPrintFormatSettingsLoaded Then
        LoadUserSettings
        SetFormatOptionsSettingsOnPrintFnObject
        mPrintFormatSettingsLoaded = True
    End If
End Sub

Private Sub SaveUserSettings()
    If mCurrentGrid Is Nothing Then Exit Sub
    
    If mSave_EnableAutoOrientation Then
        SaveSetting AppNameForRegistry, "PrintingSettings", "Global" & "_EnableAutoOrientation", CLng(mEnableAutoOrientation)
        mSave_EnableAutoOrientation = False
        
        If (GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_AutoOrientationAutoDisabled", 10) <> 10) Then
            DeleteSetting AppNameForRegistry, "PrintingSettings", Keyname & "_AutoOrientationAutoDisabled"
        End If
        If (GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_EnableAutoOrientation", 10) <> 10) Then
            DeleteSetting AppNameForRegistry, "PrintingSettings", Keyname & "_EnableAutoOrientation"
        End If
    Else
        If RememberUserPrintingPreferences <> gfnRememberGlobally Then
            If mPrintFnObject.Orientation <> mDefaultReportOrientation Then
                If mEnableAutoOrientation Then
                    mEnableAutoOrientation = False
                    SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_EnableAutoOrientation", CLng(mEnableAutoOrientation)
                    SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_AutoOrientationAutoDisabled", CLng(True)
                Else
                    If mPrintFnObject.Orientation = vbPRORLandscape Then
                        If CBool(GetSetting(AppNameForRegistry, "PrintingSettings", Keyname & "_AutoOrientationAutoDisabled", 0)) Then
                            DeleteSetting AppNameForRegistry, "PrintingSettings", Keyname & "_AutoOrientationAutoDisabled"
                            mEnableAutoOrientation = True
                            SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_EnableAutoOrientation", CLng(mEnableAutoOrientation)
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    
    If RememberUserPrintingPreferences = gfnDontRemember Then Exit Sub
    
    If RememberUserPrintingPreferences <> gfnRememberGlobally Then
        If mPrintFnObject.Orientation <> mDefaultReportOrientation Then
            SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_Orientation", mPrintFnObject.Orientation
        End If
    End If
    
    If mSave_ScalePercent Then
        If (mPrintGridFormatSettings.ScalePercent >= 60) And (mPrintGridFormatSettings.ScalePercent <= 200) Then
            SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_ScalePercent", mPrintGridFormatSettings.ScalePercent
        End If
        mSave_ScalePercent = False
    End If
    If mSave_ColorMode Then
        If mPrintGridFormatSettings.ColorMode = vbPRCMPrinterDefault Then
            DeleteSetting AppNameForRegistry, "PrintingSettings", Keyname & "_ColorMode"
        Else
            SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_ColorMode", mPrintGridFormatSettings.ColorMode
        End If
        mSave_ColorMode = False
    End If
    If mSave_GridAlign Then
        SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_GridAlign", mPrintGridFormatSettings.GridAlign
        mSave_GridAlign = False
    End If
    If mSave_PrintPageNumbers Then
        SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_PrintPageNumbers", CLng(mPrintGridFormatSettings.PrintPageNumbers)
        mSave_PrintPageNumbers = False
    End If
    If mSave_PageNumbersPosition Then
        SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_PageNumbersPosition", mPrintGridFormatSettings.PageNumbersPosition
        mSave_PageNumbersPosition = False
    End If
    If mSave_PageNumbersFormatLong Then
        SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_PageNumbersFormatLong", mPrintGridFormatSettings.PageNumbersFormatLong
        mSave_PageNumbersFormatLong = False
    End If
    If mSave_PageNumbersFont Then
        SaveSettingFont AppNameForRegistry, "PrintingSettings", Keyname & "_PageNumbersFont", mPrintGridFormatSettings.PageNumbersFont
        mSave_PageNumbersFont = True
    End If
    If mSave_PageNumbersForeColor Then
        SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_PageNumbersForeColor", mPrintGridFormatSettings.PageNumbersForeColor
        mSave_PageNumbersForeColor = False
    End If
    If mSave_HeadingFont Then
        SaveSettingFont AppNameForRegistry, "PrintingSettings", Keyname & "_HeadingFont", mPrintGridFormatSettings.HeadingFont
        mSave_HeadingFont = False
    End If
    If mSave_HeadingFontColor Then
        SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_HeadingFontColor", mPrintGridFormatSettings.HeadingFontColor
        mSave_HeadingFontColor = False
    End If
    If mSave_SubheadingFont Then
        SaveSettingFont AppNameForRegistry, "PrintingSettings", Keyname & "_SubheadingFont", mPrintGridFormatSettings.SubheadingFont
        mSave_SubheadingFont = False
    End If
    If mSave_SubheadingFontColor Then
        SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_SubheadingFontColor", mPrintGridFormatSettings.SubheadingFontColor
        mSave_SubheadingFontColor = False
    End If
    If mSave_OtherTextsFont Then
        SaveSettingFont AppNameForRegistry, "PrintingSettings", Keyname & "_OtherTextsFont", mPrintGridFormatSettings.OtherTextsFont
        mSave_OtherTextsFont = False
    End If
    If mSave_OtherTextsFontColor Then
        SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_OtherTextsFontColor", mPrintGridFormatSettings.OtherTextsFontColor
        mSave_OtherTextsFontColor = False
    End If

    If mUserChangedScale Then
        SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_UserChangedScale", CLng(mUserChangedScale)
    End If

End Sub

Private Sub SaveMarginSettings()
    Dim iUnits As Long
    
    iUnits = mPrintFnObject.Units
    mPrintFnObject.Units = vbMillimeters
    SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_LeftMargin", mPrintFnObject.LeftMargin
    SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_TopMargin", mPrintFnObject.TopMargin
    SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_RightMargin", mPrintFnObject.RightMargin
    SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_BottomMargin", mPrintFnObject.BottomMargin
    mPrintFnObject.Units = iUnits
End Sub

Public Property Get Keyname() As String
Attribute Keyname.VB_MemberFlags = "40"
    Dim iGrid As Object
    Dim iCurrentGridNameWithParent As String
    
    Select Case mRememberUserPrintingPreferences
        Case gfnRememberByReportIDIfAvailable
            If mCurrentGridNameWithParent = "" Then
                Set iGrid = GetGridControl
                If Not iGrid Is Nothing Then
                    iCurrentGridNameWithParent = Base64Encode(ControlNameWithParent(iGrid))
                End If
            Else
                iCurrentGridNameWithParent = mCurrentGridNameWithParent
            End If
            If iCurrentGridNameWithParent <> "" Then
                Keyname = mCurrentGridNameWithParent
                If Trim$(mReportID) <> "" Then
                    Keyname = Keyname & "_" & Base64Encode(Trim$(mReportID))
                End If
            Else
                If Trim$(mReportID) <> "" Then
                    Keyname = Base64Encode(Trim$(mReportID))
                End If
            End If
        Case gfnRememberByGridName
            If mCurrentGridNameWithParent <> "" Then
                Keyname = mCurrentGridNameWithParent
            Else
                Keyname = "DefaultGridSetttings"
            End If
        Case gfnRememberGlobally
            Keyname = "Global"
    End Select
End Property

Private Sub FindTextInGrid(TextToFind, Optional nFindNext As Boolean)
    Dim iTextToFind As String
    Dim r As Long
    Dim c As Long
    Dim iCellBackColor As Long
    Dim iPass As Long
    Static sLastText As String
    Static sLastRow As Long
    Static sLastCol As Long
    Static sLastCurrentRow As Long
    Static sLastCurrentCol As Long
    Dim iRedraw As Boolean
    Dim iPrevRow As Long
    Dim iPrevCol As Long
    Dim iPrevRowSel As Long
    Dim iPrevColSel As Long
    Dim iHighLight As Long
    
    If Not IsMissing(TextToFind) Then
        iTextToFind = CStr(TextToFind)
    Else
        If nFindNext Then
            If Not mCurrentGridHandler Is Nothing Then
                iTextToFind = mCurrentGridHandler.LastTextToFind
            End If
        End If
        If iTextToFind = "" Then
            iTextToFind = InputBox(GetLocalizedString(efnGUIStr_FlexFnObject_FindTextInGrid_InputBoxEnterTextMessage), GetLocalizedString(efnGUIStr_FlexFnObject_FindTextInGrid_InputBoxEnterTextMessageTitle), sLastText)
            If iTextToFind = "" Then Exit Sub
            sLastText = iTextToFind
        End If
    End If
    iTextToFind = Trim$(LCase$(iTextToFind))
    If Not mCurrentGridHandler Is Nothing Then
        mCurrentGridHandler.LastTextToFind = iTextToFind
    End If
    
    iRedraw = mCurrentGrid.Redraw
    iPrevRow = mCurrentGrid.Row
    iPrevCol = mCurrentGrid.col
    iPrevRowSel = mCurrentGrid.RowSel
    iPrevColSel = mCurrentGrid.ColSel
    On Error Resume Next
    iHighLight = mCurrentGrid.HighLight
    On Error GoTo 0
    
    If (sLastCurrentRow = iPrevRow) And (sLastCurrentCol = iPrevCol) Then
        'If mCurrentGrid.Row <> 0 Then
            mCurrentGrid.Row = sLastRow
            mCurrentGrid.col = sLastCol
        'End If
    Else
        mCurrentGrid.col = iPrevCol
        mCurrentGrid.Row = iPrevRow
    End If
    
    sLastCurrentRow = iPrevRow
    sLastCurrentCol = iPrevCol
    
    mCurrentGrid.Redraw = False
    
    
    iPass = 1
    If iTextToFind <> "" Then
        r = mCurrentGrid.Row - 1
        c = mCurrentGrid.col
Again:
        Do Until r = mCurrentGrid.Rows - 1
            r = r + 1
            If mCurrentGrid.RowHeight(r) <> 0 Then
                Do Until c = mCurrentGrid.Cols - 1
                    c = c + 1
                    If mCurrentGrid.ColWidth(c) <> 0 Then
                        If InStr(LCase$(mCurrentGrid.TextMatrix(r, c)), iTextToFind) > 0 Then
                            mCurrentGrid.col = c
                            sLastCol = c
                            If r > 1 Then
                                If mCurrentGrid.TopRow < (r - 1) Then
                                   mCurrentGrid.TopRow = r - 1
                                End If
                            Else
                                mCurrentGrid.TopRow = r
                            End If
                            mCurrentGrid.Row = r
                            sLastRow = r
                            If Not mCurrentGrid.RowIsVisible(r) Then
                                If r > 1 Then
                                    mCurrentGrid.TopRow = r - 1
                                End If
                            End If
                            If Not mCurrentGrid.ColIsVisible(c) Then
                                If c > 1 Then
                                    mCurrentGrid.LeftCol = c - 1
                                End If
                            End If
                            
                            iCellBackColor = mCurrentGrid.CellBackColor
                            mCurrentGrid.HighLight = 0
                            mCurrentGrid.Redraw = iRedraw
                            
                            mCurrentGrid.CellBackColor = vbRed
                            mCurrentGrid.Refresh
                            Sleep 250
                            mCurrentGrid.CellBackColor = iCellBackColor
                            mCurrentGrid.Refresh
                            Sleep 250
                            mCurrentGrid.CellBackColor = vbRed
                            mCurrentGrid.Refresh
                            Sleep 250
                            mCurrentGrid.CellBackColor = iCellBackColor
                            mCurrentGrid.Refresh
                            Sleep 250
                            mCurrentGrid.CellBackColor = vbRed
                            mCurrentGrid.Refresh
                            Sleep 250
                            mCurrentGrid.CellBackColor = iCellBackColor
                            mCurrentGrid.Refresh
                            Sleep 250
                            mCurrentGrid.HighLight = iHighLight
                            'mCurrentGrid.RowSel = r
                            'mCurrentGrid.ColSel = c
                            
                            RaiseEvent TextFound(r, c, iTextToFind)
                            GoTo TheExit:
                        End If
                    End If
                    If iPass = 2 Then
                        If r = mCurrentGrid.Row And c > mCurrentGrid.col Then Exit Do
                    End If
                Loop
            End If
            If iPass = 2 Then
                If r = mCurrentGrid.Row And c > mCurrentGrid.col Then Exit Do
            End If
            c = -1
        Loop
        If iPass = 1 Then
            iPass = 2
            r = mCurrentGrid.FixedRows - 1
            c = -1
            GoTo Again:
        End If
        mCurrentGrid.Redraw = iRedraw
        MsgBox GetLocalizedString(efnGUIStr_FlexFnObject_FindTextInGrid_MsgboxTextNotFound), , ClientProductName
    End If
    
TheExit:
'    mCurrentGrid.row = iPrevRow
'    mCurrentGrid.col = iPrevCol
'    mCurrentGrid.RowSel = iPrevRowSel
'    mCurrentGrid.ColSel = iPrevColSel

End Sub

Private Function SaveGridAsExcelFile() As String
    Dim iPath1 As String
    Dim iExcelFile As New cExcelFile
    Dim iCellsFontNumber() As Long
    Dim iCellsAlignment() As Long
    Dim iFontsInGrid() As cFontAttributes
    Dim iHeadingFont As Long
    Dim iSubheadingFont As Long
    Dim iOtherTextsFont As Long
    Dim iFontAttrib As New cFontAttributes
    Dim iAuxArray()
    Dim c As Long
    Dim c2 As Long
    Dim iExcelFonts() As cFontAttributes
    Dim iExcelFontsCount As Long
    Dim iCellsTexts() As String
    Dim iTextColor As Long
    Dim iBackColor As Long
    Dim iDlg As New CommonDialogExObject
    Dim iAuxPictureBox As Control
    Dim iColumnNumber As Long
    Dim iOffsetY As Long
    Dim iMergeCols As Boolean
    Dim iMergeRows As Boolean
    Dim iCw As Long
    Dim iColumns() As Long
    Dim iCcv As Long
    Dim iAuxLng As Long
    Dim iFileName As String
    Dim iVisibleCols As Long
    Dim iCol1 As Long
    Dim iCol2 As Long
    Dim iStrs() As String
    Dim iGH As cGridHandler
    Dim iRows As Long
    Dim iLastVisibleText As String
    Dim c3 As Long
    
    Dim iRedraw As Boolean
    Dim iPrevRow As Long
    Dim iPrevCol As Long
    Dim iPrevRowSel As Long
    Dim iPrevColSel As Long
    
    On Error GoTo ErrorExit:
    
    If (FileName2 = "") Then
        iFileName = GetLocalizedString(efnGUIStr_mGlobals_ValidFileName_DefaultFileName)
    Else
        iFileName = FileName2
    End If
    
    iPath1 = GetProgramDocumentsFolder & iFileName & ".xls"

Again:
    iDlg.Flags = cdlOFNPathMustExist
    iDlg.FileName = iPath1
    iDlg.Filter = "Archivos de Excel (*.xls)|*.xls"
    iDlg.CancelError = True
    On Error Resume Next
    Set iGH = GetGridHandler(mCurrentGrid)
    If Not iGH Is Nothing Then iGH.TimerCheckRedrawEnabled = False
    iDlg.ShowSave
    If Not iGH Is Nothing Then iGH.TimerCheckRedrawEnabled = True
    
'    DoEvents
    If Err.Number <> 32755 Then 'si no canceló
        iPath1 = iDlg.FileName
        If LCase$(Right$(iPath1, 4)) <> ".xls" Then
            iPath1 = iPath1 & ".xls"
        End If
        
        If FileExists(iPath1) Then
            Select Case MsgBox(GetLocalizedString(efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox1a) & " " & iPath1 & " " & GetLocalizedString(efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox1b), vbYesNoCancel + vbExclamation, ClientProductName)
            Case vbYes
                ' sigue nomás
                Kill iPath1
                Do Until Not FileExists(iPath1)
                    Select Case MsgBox(GetLocalizedString(efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox2a) & vbCrLf & GetLocalizedString(efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox2b), vbExclamation + vbYesNoCancel, ClientProductName)
                        Case vbYes
                            ' nada
                        Case vbNo
                            GoTo Again:
                        Case Else
                            Exit Function
                    End Select
                    Kill iPath1
                Loop
            Case vbNo
                GoTo Again:
            Case vbCancel
                Exit Function
            End Select
        End If
        On Error GoTo ErrorExit:
        'SaveSetting AppNameForRegistry, "Preferences", "DocsFolder", GetFolder(iPath1)
        SaveProgramDocumentsFolder GetFolder(iPath1)
        
        iRedraw = mCurrentGrid.Redraw
        iPrevRow = mCurrentGrid.Row
        iPrevCol = mCurrentGrid.col
        iPrevRowSel = mCurrentGrid.RowSel
        iPrevColSel = mCurrentGrid.ColSel
        mCurrentGrid.Redraw = False
        
        ' prepara los datos
'        mCurrentGrid.Visible = False
         
        Dim ContCol As Long
        Dim ContFil As Long
        
        Screen.MousePointer = vbHourglass
        ReDim iCellsFontNumber(mCurrentGrid.Cols - 1, mCurrentGrid.Rows - 1)
        ReDim iCellsAlignment(mCurrentGrid.Cols - 1, mCurrentGrid.Rows - 1)
        ReDim iCellsTexts(mCurrentGrid.Cols - 1, mCurrentGrid.Rows - 1)
        ReDim iFontsInGrid(0)
        ReDim iColumns(mCurrentGrid.Cols - 1)
        
        iCol1 = 0
        iCol2 = 0
        iVisibleCols = 0
        For c = 0 To mCurrentGrid.Cols - 1
            If mCurrentGrid.ColWidth(c) <> 0 Then
                iVisibleCols = iVisibleCols + 1
                If iCol1 = 0 Then
                    iCol1 = c + 1
                Else
                    If iCol2 = 0 Then
                        iCol2 = c + 1
                    End If
                End If
            End If
        Next c
        If iCol1 = 0 Then iCol1 = 1
        If iCol2 = 0 Then iCol2 = 1
        
        iRows = 0
        For ContFil = 0 To mCurrentGrid.Rows - 1
            If mCurrentGrid.RowHeight(ContFil) <> 0 Then
                iRows = iRows + 1
            End If
        Next ContFil
        ' asigna una matriz y un array, iFontsInGrid para la lista de fuentes en la grilla, y iCellsFontNumber para saber a cual corresponde cada celda
        iCcv = -1
        For ContCol = 0 To mCurrentGrid.Cols - 1
            mCurrentGrid.col = ContCol
            If mCurrentGrid.ColWidth(ContCol) <> 0 Then
                iCcv = iCcv + 1
                iColumns(ContCol) = iCcv
                For ContFil = 0 To mCurrentGrid.Rows - 1
                    If mCurrentGrid.RowHeight(ContFil) <> 0 Then
                        mCurrentGrid.Row = ContFil
                        iFontAttrib.Name = mCurrentGrid.CellFontName
                        iFontAttrib.Size = mCurrentGrid.CellFontSize
                        iFontAttrib.Italic = mCurrentGrid.CellFontItalic
                        iFontAttrib.Bold = mCurrentGrid.CellFontBold
                        iFontAttrib.Underline = mCurrentGrid.CellFontUnderline
                        
                        iCellsFontNumber(ContCol, ContFil) = AddExcelFontToArray(iFontAttrib, iFontsInGrid)
                        
                        ' graba las alineaciones en otro array
                        iCellsAlignment(ContCol, ContFil) = mCurrentGrid.CellAlignment
                        Dim iTextoCelda As String
                        Dim No_Hacer As Boolean
                        
                        ' Guarda los textos
                        iTextColor = mCurrentGrid.ForeColor
                        iBackColor = mCurrentGrid.BackColor
                        iAuxLng = mCurrentGrid.CellForeColor
                        If iAuxLng <> 0 Then
                            iTextColor = mCurrentGrid.CellForeColor
                        End If
                        iAuxLng = mCurrentGrid.CellBackColor
                        If iAuxLng <> 0 Then
                            iBackColor = mCurrentGrid.CellBackColor
                        End If
                        If iTextColor <> iBackColor Then
                            iTextoCelda = mCurrentGrid.TextMatrix(ContFil, ContCol)
                            No_Hacer = False
                            
                            iMergeCols = ((mCurrentGrid.MergeCells = flexMergeFree) Or (mCurrentGrid.MergeCells = flexMergeRestrictColumns)) And mMergeCellsExcel
                            iMergeRows = ((mCurrentGrid.MergeCells = flexMergeFree) Or (mCurrentGrid.MergeCells = flexMergeRestrictRows)) And mMergeCellsExcel
                            
                            If iMergeRows And mCurrentGrid.MergeRow(ContFil) Then
                                ' MergeRows
                                If ContCol > 0 Then
'                                    Debug.Print mCurrentGrid.TextMatrix(ContFil, ContCol - 1)
'                                    Debug.Print iTextoCelda
                                    iLastVisibleText = ""
                                    For c3 = ContCol - 1 To 0 Step -1
                                        If mCurrentGrid.ColWidth(c3) <> 0 Then
                                            iLastVisibleText = mCurrentGrid.TextMatrix(ContFil, c3)
                                            Exit For
                                        End If
                                    Next c3
                                    If mCurrentGrid.TextMatrix(ContFil, ContCol) = iLastVisibleText Then
                                        No_Hacer = True
                                    End If
                                End If
                            End If
                            If iMergeCols And mCurrentGrid.MergeCol(ContCol) Then
                                ' MergeCols
                                If ContFil > 0 Then
                                    iLastVisibleText = ""
                                    For c3 = ContFil - 1 To 0 Step -1
                                        If mCurrentGrid.RowHeight(c3) <> 0 Then
                                            iLastVisibleText = mCurrentGrid.TextMatrix(c3, ContCol)
                                            Exit For
                                        End If
                                    Next c3
                                    If mCurrentGrid.TextMatrix(ContFil, ContCol) = iLastVisibleText Then
                                        No_Hacer = True
                                    End If
                                End If
                            End If
                            If Not No_Hacer Then
                                iCellsTexts(ContCol, ContFil) = mCurrentGrid.TextMatrix(ContFil, ContCol)
                            End If
                        Else
                            iCellsTexts(ContCol, ContFil) = " "
                        End If
                    End If
                Next ContFil
            Else
                iColumns(ContCol) = -1
            End If
        Next ContCol
        ' Asigna las otras fuentes
        If Trim$(mHeading) <> "" Then
            iFontAttrib.Name = mPrintGridFormatSettings.HeadingFont.Name
            iFontAttrib.Size = mPrintGridFormatSettings.HeadingFont.Size
            iFontAttrib.Bold = mPrintGridFormatSettings.HeadingFont.Bold
            iFontAttrib.Italic = mPrintGridFormatSettings.HeadingFont.Italic
            iFontAttrib.Underline = mPrintGridFormatSettings.HeadingFont.Underline
            iFontAttrib.Strikethrough = mPrintGridFormatSettings.HeadingFont.Strikethrough
            iFontAttrib.ForeColor = mPrintGridFormatSettings.HeadingFontColor
            iHeadingFont = AddExcelFontToArray(iFontAttrib, iFontsInGrid)
        End If
        If Trim$(mSubheading) <> "" Then
            iFontAttrib.Name = mPrintGridFormatSettings.SubheadingFont.Name
            iFontAttrib.Size = mPrintGridFormatSettings.SubheadingFont.Size
            iFontAttrib.Bold = mPrintGridFormatSettings.SubheadingFont.Bold
            iFontAttrib.Italic = mPrintGridFormatSettings.SubheadingFont.Italic
            iFontAttrib.Underline = mPrintGridFormatSettings.SubheadingFont.Underline
            iFontAttrib.Strikethrough = mPrintGridFormatSettings.SubheadingFont.Strikethrough
            iFontAttrib.ForeColor = mPrintGridFormatSettings.SubheadingFontColor
            iSubheadingFont = AddExcelFontToArray(iFontAttrib, iFontsInGrid)
        End If
        If Trim$(mMiddleText) <> "" Or Trim$(mFinalText) <> "" Then
            iFontAttrib.Name = mPrintGridFormatSettings.OtherTextsFont.Name
            iFontAttrib.Size = mPrintGridFormatSettings.OtherTextsFont.Size
            iFontAttrib.Bold = mPrintGridFormatSettings.OtherTextsFont.Bold
            iFontAttrib.Italic = mPrintGridFormatSettings.OtherTextsFont.Italic
            iFontAttrib.Underline = mPrintGridFormatSettings.OtherTextsFont.Underline
            iFontAttrib.Strikethrough = mPrintGridFormatSettings.OtherTextsFont.Strikethrough
            iFontAttrib.ForeColor = mPrintGridFormatSettings.OtherTextsFontColor
            iOtherTextsFont = AddExcelFontToArray(iFontAttrib, iFontsInGrid)
        End If
        
        ReDim iAuxArray(UBound(iFontsInGrid), 4)
        If UBound(iFontsInGrid) <= 4 Then
            
            iExcelFontsCount = UBound(iFontsInGrid)
            ReDim iExcelFonts(iExcelFontsCount)
            ' Asigna la tabla de conversión directamente. el elemento 3 de las segunda dimensión de iAuxArray es la tabla de conversión entre la lista de fuentes de iFontsInGrid y las 4 disponibles que es el maximo que acepta el formato de archivos de excel 2.1.
            ' También asigna las fuentes que se van a usar para el archivo de Excel
            For c = 1 To UBound(iFontsInGrid)
                iAuxArray(c, 2) = c ' el elemento 2 de la segunda dimensión se usa para la tabla de conversión con iFontsInGrid
                iAuxArray(c, 4) = c ' el elemento 4 de la segunda dimensión se usa para la tabla de conversión con iExcelFonts
                Set iExcelFonts(c) = New cFontAttributes
                iExcelFonts(c).Name = iFontsInGrid(c).Name
                iExcelFonts(c).Size = iFontsInGrid(c).Size
                iExcelFonts(c).Italic = iFontsInGrid(c).Italic
                iExcelFonts(c).Bold = iFontsInGrid(c).Bold
                iExcelFonts(c).Underline = iFontsInGrid(c).Underline
            Next
            
        Else
            ' mide el ancho de las fuentes
            On Error Resume Next
            mCurrentGrid.Parent.Controls.Add "VB.PictureBox", "iAuxPicGGCE"
            On Error GoTo ErrorExit:
            Set iAuxPictureBox = mCurrentGrid.Parent.Controls("iAuxPicGGCE")
            
            For c = 1 To UBound(iFontsInGrid)
                iAuxPictureBox.FontName = iFontsInGrid(c).Name
                If iFontsInGrid(c).Size < 1 Then
                    iAuxPictureBox.FontSize = 1
                Else
                    iAuxPictureBox.FontSize = iFontsInGrid(c).Size
                End If
                iAuxPictureBox.FontItalic = iFontsInGrid(c).Italic
                iAuxPictureBox.FontBold = iFontsInGrid(c).Bold
                iAuxPictureBox.FontUnderLine = iFontsInGrid(c).Underline
                iAuxArray(c, 0) = iAuxPictureBox.TextWidth("Texto Prueba")
            Next
            mCurrentGrid.Parent.Controls.Remove "iAuxPicGGCE"
            
            ' cuenta en cuantas celdas es usada cada fuente
            For ContCol = 0 To mCurrentGrid.Cols - 1
                If iColumns(ContCol) <> -1 Then
                    For ContFil = 0 To mCurrentGrid.Rows - 1
                        If mCurrentGrid.RowHeight(ContFil) <> 0 Then
                            iAuxArray(iCellsFontNumber(ContCol, ContFil), 1) = iAuxArray(iCellsFontNumber(ContCol, ContFil), 1) + 1
                        End If
                    Next ContFil
                End If
            Next ContCol
            
            ' busca las dos mas usadas
            Dim AuxPrimeraMasUsada As Long
            Dim AuxSegundaMasUsada As Long
            
            For c = 1 To UBound(iAuxArray)
                If AuxSegundaMasUsada = 0 Then
                    If iAuxArray(c, 1) < iAuxArray(AuxPrimeraMasUsada, 1) Then
                        AuxSegundaMasUsada = c
                    End If
                End If
                If iAuxArray(c, 1) > iAuxArray(AuxSegundaMasUsada, 1) Then
                    If iAuxArray(c, 1) < iAuxArray(AuxPrimeraMasUsada, 1) Then
                        AuxSegundaMasUsada = c
                    End If
                End If
                If iAuxArray(c, 1) > iAuxArray(AuxPrimeraMasUsada, 1) Then
                    AuxSegundaMasUsada = AuxPrimeraMasUsada
                    AuxPrimeraMasUsada = c
                End If
            Next
            
            ' las dos mas usadas ya las asigna
            iAuxArray(AuxPrimeraMasUsada, 3) = AuxPrimeraMasUsada
            iAuxArray(AuxSegundaMasUsada, 3) = AuxSegundaMasUsada
            
            ' busca las dos fuentes mas diferentes a las dos que ya están asignadas entre las que quedan sin asignar
            Dim AuxSng1 As Single
            Dim AuxSng2 As Single
            Dim AuxSng3 As Single
            Dim AuxPrimeraMasDiferente As Long
            Dim AuxSegundaMasDiferente As Long
            Dim AuxValorPrimeraMasDiferente As Single
            Dim AuxValorSegundaMasDiferente As Single
            
            AuxSng1 = 0
            AuxSng2 = 0
            AuxSng3 = 0
            AuxPrimeraMasDiferente = 0
            AuxSegundaMasDiferente = 0
            AuxValorPrimeraMasDiferente = 0
            AuxValorSegundaMasDiferente = 0
            ' busca la primera mas diferente
            For c = 1 To UBound(iAuxArray)
                If c <> AuxPrimeraMasUsada And c <> AuxSegundaMasUsada Then
                    AuxSng1 = DifferenceRelationBetweenFonts(iFontsInGrid(c), CSng(iAuxArray(c, 0)), iFontsInGrid(AuxPrimeraMasUsada), CSng(iAuxArray(AuxPrimeraMasUsada, 0)))
                    AuxSng2 = DifferenceRelationBetweenFonts(iFontsInGrid(c), CSng(iAuxArray(c, 0)), iFontsInGrid(AuxSegundaMasUsada), CSng(iAuxArray(AuxSegundaMasUsada, 0)))
                    ' busca la menor diferencia y la deja en AuxSng1
                    If AuxSng2 < AuxSng1 Then
                        AuxSng1 = AuxSng2
                    End If
                    If AuxSng1 > AuxValorPrimeraMasDiferente Then
                        AuxPrimeraMasDiferente = c
                        AuxValorPrimeraMasDiferente = AuxSng1
                    End If
                End If
            Next
            ' busca la segunda mas diferente
            For c = 1 To UBound(iAuxArray)
                If c <> AuxPrimeraMasUsada And c <> AuxSegundaMasUsada And c <> AuxSegundaMasDiferente Then
                    AuxSng1 = DifferenceRelationBetweenFonts(iFontsInGrid(c), CSng(iAuxArray(c, 0)), iFontsInGrid(AuxPrimeraMasUsada), CSng(iAuxArray(AuxPrimeraMasUsada, 0)))
                    AuxSng2 = DifferenceRelationBetweenFonts(iFontsInGrid(c), CSng(iAuxArray(c, 0)), iFontsInGrid(AuxSegundaMasUsada), CSng(iAuxArray(AuxSegundaMasUsada, 0)))
                    AuxSng3 = DifferenceRelationBetweenFonts(iFontsInGrid(c), CSng(iAuxArray(c, 0)), iFontsInGrid(AuxPrimeraMasDiferente), CSng(iAuxArray(AuxPrimeraMasDiferente, 0)))
                    ' busca la menor diferencia y la deja en AuxSng1
                    If AuxSng2 < AuxSng1 Then
                        AuxSng1 = AuxSng2
                    End If
                    If AuxSng3 < AuxSng1 Then
                        AuxSng1 = AuxSng3
                    End If
                    If AuxSng1 > AuxValorSegundaMasDiferente Then
                        AuxSegundaMasDiferente = c
                        AuxValorSegundaMasDiferente = AuxSng1
                    End If
                End If
            Next
            
            ' asigna las dos mas diferentes
            iAuxArray(AuxPrimeraMasDiferente, 3) = AuxPrimeraMasDiferente
            iAuxArray(AuxSegundaMasDiferente, 3) = AuxSegundaMasDiferente
            
            ' asigna las que faltan a las mas similares de las que ya están
            Dim AuxMasparecida As Long
            Dim AuxValorMasParecida As Single
            ' repasa las que faltan asignar y...
            For c = 1 To UBound(iFontsInGrid)
                If IsEmpty(iAuxArray(c, 3)) Then ' si es una que no está asignada
                    '... busca en las ya asignadas la mas parecida
                    AuxMasparecida = 0
                    AuxValorMasParecida = 0
                    For c2 = 1 To UBound(iFontsInGrid)
                        If c2 = AuxPrimeraMasUsada Or c2 = AuxSegundaMasUsada Or c2 = AuxPrimeraMasDiferente Or c2 = AuxSegundaMasDiferente Then ' si es una que está asignada
                            AuxSng1 = DifferenceRelationBetweenFonts(iFontsInGrid(c), CSng(iAuxArray(c, 0)), iFontsInGrid(c2), CSng(iAuxArray(c2, 0)))
                            If AuxSng1 < AuxValorMasParecida Or AuxMasparecida = 0 Then
                                AuxMasparecida = c2
                                AuxValorMasParecida = AuxSng1
                            End If
                        End If
                    Next
                    ' la asigna
                    iAuxArray(c, 3) = AuxMasparecida
                End If
            Next
            
            Dim AuxNumeroFuenteYaAsignada(4) As Long
            
            iExcelFontsCount = 4
            ReDim iExcelFonts(4)
            Dim Encontrada As Boolean
            Dim Contador3 As Long
            Dim AuxNumeroFuenteAAsignar As Long
            
            Contador3 = 0
            For c = 1 To UBound(iFontsInGrid)
                Encontrada = False
                For c2 = 1 To 4
                    ' busca en la lista de las ya asignadas para ver si esta
                    If AuxNumeroFuenteYaAsignada(c2) = iAuxArray(c, 3) Then
                        Encontrada = True
                        AuxNumeroFuenteAAsignar = c2
                        Exit For
                    End If
                Next
                If Not Encontrada Then ' si no la encunetra la pone en la lista
                    Contador3 = Contador3 + 1
                    AuxNumeroFuenteAAsignar = Contador3
                    AuxNumeroFuenteYaAsignada(AuxNumeroFuenteAAsignar) = iAuxArray(c, 3)
                    Set iExcelFonts(AuxNumeroFuenteAAsignar) = New cFontAttributes
                    iExcelFonts(AuxNumeroFuenteAAsignar).Name = iFontsInGrid(iAuxArray(c, 3)).Name
                    iExcelFonts(AuxNumeroFuenteAAsignar).Size = iFontsInGrid(iAuxArray(c, 3)).Size
                    iExcelFonts(AuxNumeroFuenteAAsignar).Italic = iFontsInGrid(iAuxArray(c, 3)).Italic
                    iExcelFonts(AuxNumeroFuenteAAsignar).Bold = iFontsInGrid(iAuxArray(c, 3)).Bold
                    iExcelFonts(AuxNumeroFuenteAAsignar).Underline = iFontsInGrid(iAuxArray(c, 3)).Underline
                End If
                iAuxArray(c, 4) = AuxNumeroFuenteAAsignar ' el elemento 4 de la segunda dimensión se usa para la tabla de conversión con iExcelFonts
            Next
        End If
        
        ' genera el archivo
        With iExcelFile
            .CreateFile iPath1
            .PrintGridLines = True
            
            .SetMargin xlsTopMargin, 1   ' en  pulgadas
            .SetMargin xlsLeftMargin, 0.5
            .SetMargin xlsRightMargin, 0.5
            .SetMargin xlsBottomMargin, 1
            
            Dim AuxFormatoFuente As XLSFontFormatting
            For c = 1 To iExcelFontsCount
                AuxFormatoFuente = xlsNoFormat
                If iExcelFonts(c).Italic Then
                    AuxFormatoFuente = AuxFormatoFuente + xlsItalic
                End If
                If iExcelFonts(c).Bold Then
                    AuxFormatoFuente = AuxFormatoFuente + xlsBold
                End If
                If iExcelFonts(c).Underline Then
                    AuxFormatoFuente = AuxFormatoFuente + xlsUnderline
                End If
                .SetFont iExcelFonts(c).Name, iExcelFonts(c).Size, AuxFormatoFuente         'font c -1
            Next
            
            ' asigna los anchos de las columnas
            For c = 0 To mCurrentGrid.Cols - 1
                If iColumns(c) <> -1 Then
                    iCw = mCurrentGrid.ColWidth(c)
                    If iCw = -1 Then
                        iCw = mCurrentGrid.Font.Size * Screen.TwipsPerPixelX * 8
                    End If
                    .SetColumnWidth CByte(iColumns(c) + 1), CByte(iColumns(c) + 1), iCw / 91
                End If
            Next
            
            ' ahora convierte los valores que tiene almacenado en la matriz que dice que fuente tiene cada celda
            If mHeading <> "" Then
                iOffsetY = 2
                If iVisibleCols > 2 Then
                    iColumnNumber = iCol2
                Else
                    iColumnNumber = iCol1
                End If
                .WriteValue xlsText, ExcelFont(CLng(iAuxArray(iHeadingFont, 4))), xlsLeftAlign, xlsNormal, 1, iColumnNumber, mHeading
            Else
                iOffsetY = 0
            End If
            
            If mSubheading <> "" Then
                If InStr(mSubheading, vbCrLf) > 0 Then
                    iStrs = Split(mSubheading, vbCrLf)
                    iOffsetY = iOffsetY + 1
                    For c = 0 To UBound(iStrs)
                        iOffsetY = iOffsetY + 1
                        iColumnNumber = 1
                        .WriteValue xlsText, ExcelFont(CLng(iAuxArray(iSubheadingFont, 4))), xlsLeftAlign, xlsNormal, iOffsetY - 1, iColumnNumber, iStrs(c)
                    Next c
                Else
                    iOffsetY = iOffsetY + 2
                    iColumnNumber = 1
                    .WriteValue xlsText, ExcelFont(CLng(iAuxArray(iSubheadingFont, 4))), xlsLeftAlign, xlsNormal, iOffsetY - 1, iColumnNumber, mSubheading
                End If
            End If
            
            If mMiddleText <> "" Then
                If InStr(mMiddleText, vbCrLf) > 0 Then
                    iStrs = Split(mMiddleText, vbCrLf)
                    iOffsetY = iOffsetY + 1
                    For c = 0 To UBound(iStrs)
                        iOffsetY = iOffsetY + 1
                        iColumnNumber = 1
                        .WriteValue xlsText, ExcelFont(CLng(iAuxArray(iOtherTextsFont, 4))), xlsLeftAlign, xlsNormal, iOffsetY - 1, iColumnNumber, iStrs(c)
                    Next c
                Else
                    iOffsetY = iOffsetY + 2
                    iColumnNumber = 1
                    .WriteValue xlsText, ExcelFont(CLng(iAuxArray(iOtherTextsFont, 4))), xlsLeftAlign, xlsNormal, iOffsetY - 1, iColumnNumber, mMiddleText
                End If
            End If
            
            For ContCol = 0 To mCurrentGrid.Cols - 1
                If iColumns(ContCol) <> -1 Then
                    For ContFil = 0 To mCurrentGrid.Rows - 1
                        If mCurrentGrid.RowHeight(ContFil) <> 0 Then
                            iCellsFontNumber(ContCol, ContFil) = ExcelFont(CLng(iAuxArray(iCellsFontNumber(ContCol, ContFil), 4)))
                            
                            ' convierte tambien las alineaciones de las celdas
                            Dim AuxCellAlignment As Long
                            If iCellsAlignment(ContCol, ContFil) <> 0 Then
                                AuxCellAlignment = iCellsAlignment(ContCol, ContFil)
                            Else
                                If ContFil < mCurrentGrid.FixedRows Then
                                    AuxCellAlignment = mCurrentGrid.FixedAlignment(ContCol)
                                Else
                                    AuxCellAlignment = mCurrentGrid.ColAlignment(ContCol)
                                End If
                            End If
                            
                            ' contantes de alineacion de Excel:
                            ' xlsGeneralAlign = 0
                            ' xlsLeftAlign = 1
                            ' xlsCentreAlign = 2
                            ' xlsrightAlign = 3
                            ' xlsFillCell = 4
                            ' xlsLeftBorder = 8
                            ' xlsRightBorder = 16
                            ' xlsTopBorder = 32
                            ' xlsBottomBorder = 64
                            ' xlsShaded = 128
                            
                            Select Case AuxCellAlignment
                                Case 0 'flexAlignLeftTop    0   Izquierda arriba.
                                    iCellsAlignment(ContCol, ContFil) = xlsLeftAlign
                                Case 1 'flexAlignLeftCenter 1   Izquierda centro. (Predeterminado para las cadenas)
                                    iCellsAlignment(ContCol, ContFil) = xlsLeftAlign
                                Case 2 'flexAlignLeftBottom 2   Izquierda abajo.
                                    iCellsAlignment(ContCol, ContFil) = xlsLeftAlign
                                Case 3 'flexAlignCenterTop  3   Centro arriba.
                                    iCellsAlignment(ContCol, ContFil) = xlsCentreAlign
                                Case 4 'flexAlignCenterCenter   4   Centrado centro.
                                    iCellsAlignment(ContCol, ContFil) = xlsCentreAlign
                                Case 5 'flexAlignCenterBottom   5   Centro abajo.
                                    iCellsAlignment(ContCol, ContFil) = xlsCentreAlign
                                Case 6 'flexAlignRightTop   6   Derecha arriba.
                                    iCellsAlignment(ContCol, ContFil) = xlsrightAlign
                                Case 7 'flexAlignRightCenter    7   Derecha centro. (Predeterminado para los números)
                                    iCellsAlignment(ContCol, ContFil) = xlsrightAlign
                                Case 8 'flexAlignRightBottom    8   Derecha abajo.
                                    iCellsAlignment(ContCol, ContFil) = xlsrightAlign
                                Case 9 'flexAlignGeneral    9   General.
                                    iCellsAlignment(ContCol, ContFil) = xlsGeneralAlign
                            End Select
                            
                            ' pone los textos en el archivo
                            If mCurrentGrid.ColWidth(ContCol) <> 0 Then
                                If iCellsTexts(ContCol, ContFil) <> "" Then
                                    iCellsTexts(ContCol, ContFil) = Replace(iCellsTexts(ContCol, ContFil), vbCrLf, vbLf)
                                    .WriteValue xlsText, iCellsFontNumber(ContCol, ContFil), iCellsAlignment(ContCol, ContFil), xlsNormal, ContFil + 1 + iOffsetY, iColumns(ContCol) + 1, iCellsTexts(ContCol, ContFil)
                                End If
                            End If
                        End If
                    Next ContFil
                End If
            Next ContCol
            
            iOffsetY = iOffsetY + iRows + 1
            
            If mFinalText <> "" Then
                If InStr(mFinalText, vbCrLf) > 0 Then
                    iStrs = Split(mFinalText, vbCrLf)
                    iOffsetY = iOffsetY + 1
                    For c = 0 To UBound(iStrs)
                        iOffsetY = iOffsetY + 1
                        iColumnNumber = 1
                        .WriteValue xlsText, ExcelFont(CLng(iAuxArray(iOtherTextsFont, 4))), xlsLeftAlign, xlsNormal, iOffsetY - 1, iColumnNumber, iStrs(c)
                    Next c
                Else
                    iOffsetY = iOffsetY + 2
                    iColumnNumber = 1
                    .WriteValue xlsText, ExcelFont(CLng(iAuxArray(iOtherTextsFont, 4))), xlsLeftAlign, xlsNormal, iOffsetY - 1, iColumnNumber, mFinalText
                End If
            End If
            
            .CloseFile
        End With
    '    mCurrentGrid.Visible = True
        Screen.MousePointer = vbDefault
        SaveGridAsExcelFile = iPath1
        
        mCurrentGrid.Row = iPrevRow
        mCurrentGrid.col = iPrevCol
        mCurrentGrid.RowSel = iPrevRowSel
        mCurrentGrid.ColSel = iPrevColSel
        mCurrentGrid.Redraw = iRedraw
        
        If mAfterSaveAction = gfnOfferToOpen Then
            If IsExcelInstalled Then
                If MsgBox(GetLocalizedString(efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox3), vbYesNo + vbQuestion, ClientProductName) = vbYes Then
                    Dim lHWnd As Long
                    Dim iRet As Long
                    
                    iRet = ShellExecute(lHWnd, "OPEN", SaveGridAsExcelFile, "", "", SW_SHOWMAXIMIZED)
                End If
            End If
        Else
            If AfterSaveAction = gfnShowConfirmationMessage Then
                MsgBox GetLocalizedString(efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox4) & " " & SaveGridAsExcelFile, , ClientProductName
            End If
        End If
    
    Else
        SaveGridAsExcelFile = ""
    End If
    
    Exit Function
    
ErrorExit:
    MsgBox GetLocalizedString(efnGUIStr_FlexFnObject_SaveGridAsExcelFile_MsgBox5) & " " & Err.Description, vbExclamation
End Function

Private Function AddExcelFontToArray(nFont As cFontAttributes, ByRef nFontsArray As Variant) As Long
    Dim c As Long
    Dim iAdd As Boolean
    Dim iSimilarity As Single
    Dim iBestSimilarity As Single
    Dim iMoreSimilar As Long
    
    If UBound(nFontsArray) < 4 Then
        iAdd = True
        For c = 1 To UBound(nFontsArray)
            If nFont.Italic = nFontsArray(c).Italic Then
                If nFont.Bold = nFontsArray(c).Bold Then
                    If nFont.Underline = nFontsArray(c).Underline Then
                        If nFont.Size = nFontsArray(c).Size Then
                            If nFont.Name = nFontsArray(c).Name Then
                                If nFont.Strikethrough = nFontsArray(c).Strikethrough Then
                                    iAdd = False
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
        If iAdd Then
            c = UBound(nFontsArray) + 1
            ReDim Preserve nFontsArray(c)
            Set nFontsArray(c) = New cFontAttributes
            nFontsArray(c).Name = nFont.Name
            nFontsArray(c).Size = nFont.Size
            nFontsArray(c).Italic = nFont.Italic
            nFontsArray(c).Bold = nFont.Bold
            nFontsArray(c).Underline = nFont.Underline
        End If
        AddExcelFontToArray = c
    Else
        ' find the more similar font already existent
        For c = 1 To UBound(nFontsArray)
            iSimilarity = nFont.Size / nFontsArray(c).Size
            If nFont.Bold Then
                iSimilarity = iSimilarity * 1.2
            End If
            If nFontsArray(c).Bold Then
                iSimilarity = iSimilarity / 1.2
            End If
            If iSimilarity > 1 Then
                iSimilarity = 1 / iSimilarity
            End If
            If nFont.Underline And nFontsArray(c).Underline Then
                iSimilarity = iSimilarity * 1.2
            Else
                If nFont.Underline Or nFontsArray(c).Underline Then
                    iSimilarity = iSimilarity / 1.3
                End If
            End If
            If nFont.Italic And nFontsArray(c).Italic Then
                iSimilarity = iSimilarity * 1.2
            Else
                If nFont.Italic Or nFontsArray(c).Italic Then
                    iSimilarity = iSimilarity / 1.3
                End If
            End If
            If nFont.Strikethrough And nFontsArray(c).Strikethrough Then
                iSimilarity = iSimilarity * 1.6
            Else
                If nFont.Strikethrough Or nFontsArray(c).Strikethrough Then
                    iSimilarity = iSimilarity / 2
                End If
            End If
            
            If iSimilarity > iBestSimilarity Then
                iMoreSimilar = c
                iBestSimilarity = iSimilarity
            End If
        Next
        
        AddExcelFontToArray = iMoreSimilar
    End If
End Function

Private Function DifferenceRelationBetweenFonts(nFuente1 As cFontAttributes, nAncho1 As Single, nFuente2 As cFontAttributes, nAncho2 As Single) As Single
    Dim iR As Single
    Dim iFactor As Single
    
    iFactor = 1
    If nAncho1 > nAncho2 Then
        iR = nAncho1 / nAncho2
    Else
        iR = nAncho2 / nAncho1
    End If
    If nFuente1.Name <> nFuente2.Name Then
        iFactor = iFactor + 0.15
    End If
    If nFuente1.Italic <> nFuente2.Italic Then
        iFactor = iFactor + 0.15
    End If
    If nFuente1.Bold <> nFuente2.Bold Then
        iFactor = iFactor + 0.15
    End If
    If nFuente1.Underline <> nFuente2.Underline Then
        iFactor = iFactor + 0.3
    End If
    If nFuente1.Strikethrough <> nFuente2.Strikethrough Then
        iFactor = iFactor + 0.6
    End If
    DifferenceRelationBetweenFonts = iR * iFactor
End Function

Private Function ExcelFont(nValue As Long) As XLSCellFont
    Select Case nValue
        Case 1
            ExcelFont = xlsFont0
        Case 2
            ExcelFont = xlsFont1
        Case 3
            ExcelFont = xlsFont2
        Case 4
            ExcelFont = xlsFont3
        Case Else
            Err.Raise 2362, , "Value de fuente Excel inválido"
    End Select
End Function

Private Function IsAppPresent(ProgID As String) As Boolean
    IsAppPresent = Not (Replace(Replace(Replace(Replace(ProgIdToCLSID(ProgID), "0", ""), "-", ""), "{", ""), "}", "") = "")
End Function

Private Function IsExcelInstalled() As Boolean
    Static sValue As Long
    
    If sValue = 0 Then
        sValue = IIf(IsAppPresent("Excel.Application"), 1, -1)
    End If
    IsExcelInstalled = sValue = 1
End Function

' Convert a ProgID (such as "Word.Application") into the
' string representation of its CLSID
Function ProgIdToCLSID(ByVal ProgID As String) As String
    Dim pResult As Long, pChar As Long
    Dim char As Integer, Length As Long
    ' No need to use a special UDT
    Dim GUID(15) As Byte
    
    ' get the CLSID in binary form
    CLSIDFromProgID StrPtr(ProgID), GUID(0)
    ' convert to a string, get pointer to result
    StringFromCLSID GUID(0), pResult
    ' find the terminating null char
    pChar = pResult - 2
    Do
    pChar = pChar + 2
    CopyMemory char, ByVal pChar, 2
    Loop While char
    ' now get the entire string in one operation
    Length = pChar - pResult
    ' no need for a temporary string
    ProgIdToCLSID = Space$(Length \ 2)
    CopyMemory ByVal StrPtr(ProgIdToCLSID), ByVal pResult, Length
    ' release the memory allocated to the string
    CoTaskMemFree pResult
End Function

Private Function CopyGridToClipboard() As Boolean
    Dim iSeparatorCharacter As String
    Dim r As Long
    Dim c As Long
    Dim iText As String
    Dim iLongestColumnText() As Long
    Dim iLng As Long
    Dim iStr As String
    Dim ifrmCopyGridTextOptions As New frmCopyGridTextOptions
    Dim iContinue As Boolean
    Dim iMPg As Long
    Dim iMPs As Long
    Dim iConcat As New SmartConcat
    Dim iCopyColumn() As Boolean
    Dim iLastCopyColumn As Long
    Dim iTextWidths() As Single
    Dim iSpaceWidth As Single
    Dim iLongestColumnTextWidth() As Single
    Dim iColPosRigth() As Single
    Dim iTextWidthCopied() As Single
    Dim iLen As Single
    Dim iFontPrev As StdFont
    Dim iFont As StdFont
    Dim iTextMatrix()
    Dim iCellLines() As String
    Dim iPos1 As Long
    Dim iNeedNormalization As Boolean
    Const cMaxFixedTextLenght As Long = 50
    Dim iStr2 As String
    Dim iColPosRightOfLastRow() As Single
    Dim iPreviousColumnRight As Single
    Dim iSpaceLength As Single
    Dim iMaxVariableFontTextWidth As Single
    Dim iPosSng As Single
    Dim iAuxTextWidth As Single
    Dim iAuxTextWidth2 As Single
    Dim iAuxLng As Long
    Dim iTextBefore As String
    Dim iTextAfter As String
    Dim iGH As cGridHandler
    
    If Not mCurrentGrid Is Nothing Then
        iStr = mCurrentGrid.Name
    End If
    RaiseEvent BeforeCopyingToClipboard(iStr, iTextBefore, iTextAfter)
    iStr = ""
    
    ReDim iCopyColumn(mCurrentGrid.Cols - 1)
    If mCopyToClipboardMode = 0 Then
        ifrmCopyGridTextOptions.SpecialSeparatorCharacters = mSpecialSeparatorCharacters
        ifrmCopyGridTextOptions.SetGrid mCurrentGrid, Keyname
        
        Set iGH = GetGridHandler(mCurrentGrid)
        If Not iGH Is Nothing Then iGH.TimerCheckRedrawEnabled = False
        ShowModal ifrmCopyGridTextOptions
        If Not iGH Is Nothing Then iGH.TimerCheckRedrawEnabled = True
        If ifrmCopyGridTextOptions.OKPressed Then
            iContinue = True
            mCopyToClipboardMode = ifrmCopyGridTextOptions.CopyToClipboardMode
            mSpecialSeparatorCharacters = ifrmCopyGridTextOptions.SpecialSeparatorCharacters
            For c = 0 To mCurrentGrid.Cols - 1
                iCopyColumn(c) = ifrmCopyGridTextOptions.CopyColumn(c)
                If iCopyColumn(c) Then
                    iLastCopyColumn = c
                End If
            Next c
            If mCopyToClipboardMode = gfnSeparateWithSpacesForVariableFont Then
                Set iFont = ifrmCopyGridTextOptions.TargetFont
            End If
        End If
    Else
        If mCopyToClipboardMode = gfnSeparateWithSpacesForVariableFont Then
            Set iFont = New StdFont
            Set iFont = mCurrentGrid.Font
        End If
        iContinue = True
        For c = 0 To mCurrentGrid.Cols - 1
            iCopyColumn(c) = True
        Next c
        iLastCopyColumn = mCurrentGrid.Cols - 1
    End If
    If iContinue Then
        CopyGridToClipboard = True
        If mCopyToClipboardMode = gfnSeparateWithTabs Then
            iSeparatorCharacter = vbTab
        Else
            iSeparatorCharacter = mSpecialSeparatorCharacters
        End If
        iMPs = Screen.MousePointer
        Screen.MousePointer = vbHourglass
        iMPg = mCurrentGrid.MousePointer
        mCurrentGrid.MousePointer = vbHourglass
        iTextMatrix = CopyGridTextMatrix(mCurrentGrid)
        Select Case mCopyToClipboardMode
            Case gfnSeparateWithTabs, gfnSeparateWithSpecialCharacters
                For r = 0 To UBound(iTextMatrix, 1)
                    For c = 0 To mCurrentGrid.Cols - 1
                        If iCopyColumn(c) Then
                            If mCurrentGrid.ColWidth(c) <> 0 Then
                                iConcat.AddString CStr(iTextMatrix(r, c))
                                If c <> iLastCopyColumn Then
                                    iConcat.AddString iSeparatorCharacter
                                End If
                            End If
                        End If
                    Next c
                    iConcat.AddString vbCrLf
                Next r
            Case gfnSeparateWithSpacesForFixedFont
                
                iNeedNormalization = False
                ' too long cell text must be splitted into lines
                For r = 0 To UBound(iTextMatrix, 1)
                    For c = 0 To mCurrentGrid.Cols - 1
                        'If the text of the cell is longer than the max allowed for a column
                        If (Len(iTextMatrix(r, c)) > cMaxFixedTextLenght) Or (InStr(iTextMatrix(r, c), vbCrLf) > 0) Then
                            iNeedNormalization = True
                            iStr = iTextMatrix(r, c)
                            
                            ' remove line feeds at the beginning, if any
                            Do While Left$(iStr, 2) = vbCrLf
                                If Len(iStr) > 2 Then
                                    iStr = Mid$(iStr, 3)
                                Else
                                    iStr = ""
                                End If
                            Loop
                            
                            ReDim iCellLines(0)
                            ' separate in lines because of line feeds
                            iPos1 = InStr(iStr, vbCrLf)
                            Do While (iPos1 <> 0) And (iPos1 <= cMaxFixedTextLenght)
                                ReDim Preserve iCellLines(UBound(iCellLines) + 1)
                                iCellLines(UBound(iCellLines)) = Left$(iStr, iPos1 - 1)
                                If Len(iStr) > (iPos1 + 1) Then
                                    iStr = Mid$(iStr, iPos1 + 2)
                                    iPos1 = InStr(iStr, vbCrLf)
                                Else
                                    iStr = ""
                                    iPos1 = 0
                                End If
                            Loop
                            
                            Do Until Len(iStr) <= cMaxFixedTextLenght
                                ReDim Preserve iCellLines(UBound(iCellLines) + 1)
                                iStr2 = Left$(iStr, cMaxFixedTextLenght)
                                iPos1 = InStrRev(iStr2, " ")
                                If iPos1 > 0 Then
                                    iLng = iPos1 - 1
                                Else
                                    iLng = cMaxFixedTextLenght
                                End If
                                iCellLines(UBound(iCellLines)) = Left$(iStr, iLng)
                                iStr = LTrim(Mid$(iStr, iLng + 1))
                                
                                ' separate in lines because of line feeds
                                iPos1 = InStr(iStr, vbCrLf)
                                Do While (iPos1 <> 0) And (iPos1 <= cMaxFixedTextLenght)
                                    ReDim Preserve iCellLines(UBound(iCellLines) + 1)
                                    iCellLines(UBound(iCellLines)) = Left$(iStr, iPos1 - 1)
                                    If Len(iStr) > (iPos1 + 1) Then
                                        iStr = Mid$(iStr, iPos1 + 2)
                                        iPos1 = InStr(iStr, vbCrLf)
                                    Else
                                        iStr = ""
                                        iPos1 = 0
                                    End If
                                Loop
                            Loop
                            If iStr <> "" Then
                                ReDim Preserve iCellLines(UBound(iCellLines) + 1)
                                iCellLines(UBound(iCellLines)) = iStr
                            End If
                            
                            iTextMatrix(r, c) = iCellLines
                        End If
                    Next c
                Next r
                If iNeedNormalization Then
                    iTextMatrix = GridTextMatrixNormalized(iTextMatrix)
                End If
                
                ReDim iLongestColumnText(mCurrentGrid.Cols - 1)
                For c = 0 To mCurrentGrid.Cols - 1
                    If iCopyColumn(c) Then
                        If mCurrentGrid.ColWidth(c) <> 0 Then
                            For r = 0 To UBound(iTextMatrix, 1)
                                iStr = iTextMatrix(r, c)
                                iLng = Len(iStr)
                                If iLng > iLongestColumnText(c) Then
                                    iLongestColumnText(c) = iLng
                                End If
                            Next r
                        End If
                    End If
                Next c
                
                For r = 0 To UBound(iTextMatrix, 1)
                    For c = 0 To mCurrentGrid.Cols - 1
                        If iCopyColumn(c) Then
                            If mCurrentGrid.ColWidth(c) <> 0 Then
                                iStr = iTextMatrix(r, c)
                                iConcat.AddString iStr
                                iConcat.AddString Space$(iLongestColumnText(c) - Len(iStr) + 5)
                            End If
                        End If
                    Next c
                    iConcat.AddString vbCrLf
                Next r
            Case gfnSeparateWithSpacesForVariableFont
                
                iNeedNormalization = False
                
                ' store the widths of Grid texts
                Set iFontPrev = mCurrentGrid.Parent.Font
                Set mCurrentGrid.Parent.Font = iFont
                iMaxVariableFontTextWidth = mCurrentGrid.Parent.TextWidth(String$(cMaxFixedTextLenght, "a"))
                iSpaceWidth = mCurrentGrid.Parent.TextWidth(Space$(100)) / 100
                ReDim iTextWidths(UBound(iTextMatrix, 1), mCurrentGrid.Cols - 1)
                For r = 0 To UBound(iTextMatrix, 1)
                    For c = 0 To mCurrentGrid.Cols - 1
                        If iCopyColumn(c) Then
                            iTextWidths(r, c) = mCurrentGrid.Parent.TextWidth(iTextMatrix(r, c))
                        End If
                    Next c
                Next r
                
                ' too long cell text must be splitted into lines
                For r = 0 To UBound(iTextMatrix, 1)
                    For c = 0 To mCurrentGrid.Cols - 1
                        'If the text of the cell is longer than the max allowed for a column
                        If (iTextWidths(r, c) > iMaxVariableFontTextWidth) Or (InStr(iTextMatrix(r, c), vbCrLf) > 0) Then
                            
                            iNeedNormalization = True
                            iStr = iTextMatrix(r, c)
                            
                            ' remove line feeds at the beginning, if any
                            Do While Left$(iStr, 2) = vbCrLf
                                If Len(iStr) > 2 Then
                                    iStr = Mid$(iStr, 3)
                                Else
                                    iStr = ""
                                End If
                            Loop
                            
                            ReDim iCellLines(0)
                            ' separate in lines because of line feeds
                            iPos1 = InStr(iStr, vbCrLf)
                            If iPos1 > 0 Then
                                iPosSng = mCurrentGrid.Parent.TextWidth(Left$(iStr, iPos1 - 1))
                            Else
                                iPosSng = 0
                            End If
                            Do While (iPos1 <> 0) And (iPosSng <= iMaxVariableFontTextWidth)
                                ReDim Preserve iCellLines(UBound(iCellLines) + 1)
                                iCellLines(UBound(iCellLines)) = Left$(iStr, iPos1 - 1)
                                If Len(iStr) > (iPos1 + 1) Then
                                    iStr = Mid$(iStr, iPos1 + 2)
                                    iPos1 = InStr(iStr, vbCrLf)
                                    If iPos1 > 0 Then
                                        iPosSng = mCurrentGrid.Parent.TextWidth(Left$(iStr, iPos1 - 1))
                                    Else
                                        iPosSng = 0
                                    End If
                                Else
                                    iStr = ""
                                    iPos1 = 0
                                End If
                            Loop
                            
                            iAuxTextWidth = mCurrentGrid.Parent.TextWidth(iStr)
                            Do Until iAuxTextWidth <= iMaxVariableFontTextWidth
                                ReDim Preserve iCellLines(UBound(iCellLines) + 1)
                                iAuxLng = Len(iStr) / iAuxTextWidth * iMaxVariableFontTextWidth
                                iStr2 = Left$(iStr, iAuxLng)
                                iAuxTextWidth2 = mCurrentGrid.Parent.TextWidth(iStr)
                                If iAuxTextWidth2 < iMaxVariableFontTextWidth Then
                                    Do Until iAuxTextWidth2 >= iMaxVariableFontTextWidth
                                        iAuxLng = iAuxLng + 1
                                        iStr2 = Left$(iStr, iAuxLng)
                                        iAuxTextWidth2 = mCurrentGrid.Parent.TextWidth(iStr2)
                                    Loop
                                End If
                                Do Until iAuxTextWidth2 <= iMaxVariableFontTextWidth
                                     iAuxLng = iAuxLng - 1
                                    iStr2 = Left$(iStr, iAuxLng)
                                    iAuxTextWidth2 = mCurrentGrid.Parent.TextWidth(iStr2)
                                Loop
                                
                                iPos1 = InStrRev(iStr2, " ")
                                If iPos1 > 0 Then
                                    iLng = iPos1 - 1
                                Else
                                    iLng = iAuxLng
                                End If
                                
                                iCellLines(UBound(iCellLines)) = Left$(iStr, iLng)
                                iStr = LTrim(Mid$(iStr, iLng + 1))
                                
                                ' separate in lines because of line feeds
                                iPos1 = InStr(iStr, vbCrLf)
                                If iPos1 > 0 Then
                                    iPosSng = mCurrentGrid.Parent.TextWidth(Left$(iStr, iPos1 - 1))
                                Else
                                    iPosSng = 0
                                End If
                                Do While (iPos1 <> 0) And (iPosSng <= iMaxVariableFontTextWidth)
                                    ReDim Preserve iCellLines(UBound(iCellLines) + 1)
                                    iCellLines(UBound(iCellLines)) = Left$(iStr, iPos1 - 1)
                                    If Len(iStr) > (iPos1 + 1) Then
                                        iStr = Mid$(iStr, iPos1 + 2)
                                        iPos1 = InStr(iStr, vbCrLf)
                                        If iPos1 > 0 Then
                                            iPosSng = mCurrentGrid.Parent.TextWidth(Left$(iStr, iPos1 - 1))
                                        Else
                                            iPosSng = 0
                                        End If
                                    Else
                                        iStr = ""
                                        iPos1 = 0
                                    End If
                                Loop
                                iAuxTextWidth = mCurrentGrid.Parent.TextWidth(iStr)
                            Loop
                            If iStr <> "" Then
                                ReDim Preserve iCellLines(UBound(iCellLines) + 1)
                                iCellLines(UBound(iCellLines)) = iStr
                            End If
                            
                            iTextMatrix(r, c) = iCellLines
                            
                        End If
                    Next c
                Next r
                
                If iNeedNormalization Then
                    iTextMatrix = GridTextMatrixNormalized(iTextMatrix, iTextWidths)
                    For r = 0 To UBound(iTextMatrix, 1)
                        For c = 0 To mCurrentGrid.Cols - 1
                            If iTextWidths(r, c) = 0 Then
                                If iTextMatrix(r, c) <> "" Then
                                    iTextWidths(r, c) = mCurrentGrid.Parent.TextWidth(iTextMatrix(r, c))
                                End If
                            End If
                        Next c
                    Next r
                End If
                Set mCurrentGrid.Parent.Font = iFontPrev
                
                ' find the longest texts for each column
                ReDim iLongestColumnTextWidth(mCurrentGrid.Cols - 1)
                For c = 0 To mCurrentGrid.Cols - 1
                    If iCopyColumn(c) Then
                        If mCurrentGrid.ColWidth(c) <> 0 Then
                            For r = 0 To UBound(iTextMatrix, 1)
                                If iTextWidths(r, c) > iLongestColumnTextWidth(c) Then
                                    iLongestColumnTextWidth(c) = iTextWidths(r, c)
                                End If
                            Next r
                        End If
                    End If
                Next c
                
                ReDim iColPosRigth(mCurrentGrid.Cols - 1)
                For c = 0 To mCurrentGrid.Cols - 1
                    If c = 0 Then
                        iColPosRigth(c) = iLongestColumnTextWidth(c) + iSpaceWidth * 5
                    Else
                        iColPosRigth(c) = iColPosRigth(c - 1) + iLongestColumnTextWidth(c) + iSpaceWidth * 5
                    End If
                Next c
                
                ' add the necessary spaces to each text
                ReDim iTextWidthCopied(UBound(iTextMatrix, 1))
                ReDim iColPosRightOfLastRow(mCurrentGrid.Cols - 1)
                For r = 0 To UBound(iTextMatrix, 1)
                    For c = 0 To mCurrentGrid.Cols - 1
                        If iCopyColumn(c) Then
                            If mCurrentGrid.ColWidth(c) <> 0 Then
                                iStr = iTextMatrix(r, c)
                                iConcat.AddString iStr
                                If c = 0 Then
                                    iLen = Round((iColPosRigth(c) - iTextWidths(r, c)) / iSpaceWidth)
                                Else
                                    iLen = Round((iColPosRigth(c) - iTextWidthCopied(r) - iTextWidths(r, c)) / iSpaceWidth)
                                End If
                                iTextWidthCopied(r) = iTextWidthCopied(r) + iTextWidths(r, c)
                                iStr = Space$(iLen)
                                iSpaceLength = iSpaceWidth * iLen
                                iTextWidthCopied(r) = iTextWidthCopied(r) + iSpaceWidth * iLen
                                
                                If r = 0 Then
                                    iColPosRightOfLastRow(c) = iColPosRigth(c)
                                End If
                                If c > 0 Then
                                    iPreviousColumnRight = iColPosRigth(c - 1)
                                Else
                                    iPreviousColumnRight = 0
                                End If
                                
                                ' check if adding one more space is better
                                If Abs((iPreviousColumnRight + iTextWidths(r, c) + iSpaceLength + iSpaceWidth) - iColPosRightOfLastRow(c)) < Abs((iPreviousColumnRight + iTextWidths(r, c) + iSpaceLength) - iColPosRightOfLastRow(c)) Then
                                        iStr = iStr & " "
                                        iSpaceLength = iSpaceLength + iSpaceWidth
                                        iTextWidthCopied(r) = iTextWidthCopied(r) + iSpaceWidth
                                End If
                                
                                If Trim$(iTextMatrix(r, c)) <> "" Then
                                    If c > 0 Then
                                        iColPosRightOfLastRow(c) = iColPosRightOfLastRow(c - 1) + iTextWidths(r, c) + iSpaceLength
                                    Else
                                        iColPosRightOfLastRow(c) = iTextWidths(r, c) + iSpaceLength
                                    End If
                                End If
                                
                                iConcat.AddString iStr
                            End If
                        End If
                    Next c
                    iConcat.AddString vbCrLf
                Next r
                
        End Select
        
        iText = iConcat.GenerateCurrentString
        
        If iTextBefore <> "" Then
            iText = iTextBefore & vbCrLf & iText
        End If
        If iTextAfter <> "" Then
            iText = iText & vbCrLf & iTextAfter
        End If
        ClipboardCopyUnicode iText
        Screen.MousePointer = iMPs
        mCurrentGrid.MousePointer = iMPg
    End If
    Set ifrmCopyGridTextOptions = Nothing
    
End Function

Private Function CopyGridTextMatrix(nGrid As Object) As Variant
    Dim r As Long
    Dim c As Long
    Dim iCopyGridTextMatrix()
    Dim iCopyGridTextMatrix2()
    Dim r2 As Long
    Dim iMergeRows As Boolean
    Dim iMergeCols As Boolean
    Dim iCopy As Boolean
    Dim iLastVisibleText As String
    Dim c2 As Long
    
    iMergeCols = (nGrid.MergeCells = flexMergeFree) Or (nGrid.MergeCells = flexMergeRestrictColumns)
    iMergeRows = (nGrid.MergeCells = flexMergeFree) Or (nGrid.MergeCells = flexMergeRestrictRows)
    
    ReDim iCopyGridTextMatrix(nGrid.Rows - 1, nGrid.Cols - 1)
    r2 = -1
    For r = 0 To nGrid.Rows - 1
        If nGrid.RowHeight(r) <> 0 Then
            r2 = r2 + 1
            For c = 0 To nGrid.Cols - 1
                iCopy = True
                If iMergeRows And nGrid.MergeRow(r) Then
                    If c > 0 Then
                        iLastVisibleText = ""
                        For c2 = c - 1 To 0 Step -1
                            If nGrid.ColWidth(c2) <> 0 Then
                                iLastVisibleText = nGrid.TextMatrix(r, c2)
                                Exit For
                            End If
                        Next c2
                        If nGrid.TextMatrix(r, c) = iLastVisibleText Then
                            iCopy = False
                        End If
                    End If
                End If
                If iCopy Then
                    If iMergeCols And nGrid.MergeCol(c) Then
                        If r > 0 Then
                            iLastVisibleText = ""
                            For c2 = r - 1 To 0 Step -1
                                If nGrid.RowHeight(c2) <> 0 Then
                                    iLastVisibleText = nGrid.TextMatrix(c2, c)
                                    Exit For
                                End If
                            Next c2
                            If nGrid.TextMatrix(r, c) = iLastVisibleText Then
                                iCopy = False
                            End If
                        End If
                    End If
                End If
                If iCopy Then
                    iCopyGridTextMatrix(r2, c) = nGrid.TextMatrix(r, c)
                End If
            Next c
        End If
    Next r
    
    If r2 = -1 Then r2 = 0
    
    ReDim iCopyGridTextMatrix2(r2, nGrid.Cols - 1)
    For r = 0 To r2
        For c = 0 To nGrid.Cols - 1
            iCopyGridTextMatrix2(r, c) = iCopyGridTextMatrix(r, c)
        Next c
    Next r
    
    CopyGridTextMatrix = iCopyGridTextMatrix2
End Function

Private Function GridTextMatrixNormalized(nTextMatrix, Optional nTextWidths) As Variant
    Dim r As Long
    Dim c As Long
    Dim iGridTextMatrixNormalized()
    Dim iInvertedMatrix()
    Dim iRows As Long
    Dim iCols As Long
    Dim iShift As Long
    Dim iMultipleLines As Boolean
    Dim l As Long
    Dim iAddALine As Boolean
    Dim c2 As Long
    Dim iTextWidthsOld() As Single
    Dim iInvertedTextWidthsNew() As Single
    Dim iTW As Boolean
    
    iRows = UBound(nTextMatrix, 1) + 1
    iCols = UBound(nTextMatrix, 2) + 1
    
    If Not IsMissing(nTextWidths) Then
        iTW = True
        ReDim iInvertedTextWidthsNew(iCols - 1, iRows - 1)
        ReDim iTextWidthsOld(iRows - 1, iCols - 1)
        For r = 0 To iRows - 1
            For c = 0 To iCols - 1
                iTextWidthsOld(r, c) = nTextWidths(r, c)
            Next c
        Next r
    End If
    
    ReDim iInvertedMatrix(iCols - 1, iRows - 1)
    
    For r = 0 To iRows - 1
        iMultipleLines = False
        For c = 0 To iCols - 1
            If IsArray(nTextMatrix(r, c)) Then
                iInvertedMatrix(c, r + iShift) = nTextMatrix(r, c)(1)
                iMultipleLines = True
            Else
                iInvertedMatrix(c, r + iShift) = nTextMatrix(r, c)
                If iTW Then
                    iInvertedTextWidthsNew(c, r + iShift) = iTextWidthsOld(r, c)
                End If
            End If
        Next c
        If iMultipleLines Then
'            If iTW Then
'                For c2 = 0 To iCols - 1
'                    iInvertedTextWidthsNew(c2, r + iShift) = 0
'                Next c2
'            End If
            l = 1
            Do
                l = l + 1
                iAddALine = False
                For c2 = 0 To iCols - 1
                    If IsArray(nTextMatrix(r, c2)) Then
                        If UBound(nTextMatrix(r, c2)) >= l Then
                            iAddALine = True
                            Exit For
                        End If
                    End If
                Next c2
                If iAddALine Then
                    iShift = iShift + 1
                    ReDim Preserve iInvertedMatrix(iCols - 1, iRows - 1 + iShift)
                    If iTW Then
                        ReDim Preserve iInvertedTextWidthsNew(iCols - 1, iRows - 1 + iShift)
                    End If
                    For c2 = 0 To iCols - 1
                        If IsArray(nTextMatrix(r, c2)) Then
                            If UBound(nTextMatrix(r, c2)) >= l Then
                                iInvertedMatrix(c2, r + iShift) = nTextMatrix(r, c2)(l)
                            End If
                        End If
                    Next c2
                Else
                    Exit Do
                End If
            Loop
        End If
    Next r
    
    ReDim iGridTextMatrixNormalized(UBound(iInvertedMatrix, 2), UBound(iInvertedMatrix, 1))
    If iTW Then
        ReDim nTextWidths(UBound(iInvertedMatrix, 2), UBound(iInvertedMatrix, 1))
    End If
    For r = 0 To UBound(iInvertedMatrix, 2)
        For c = 0 To UBound(iInvertedMatrix, 1)
            iGridTextMatrixNormalized(r, c) = iInvertedMatrix(c, r)
            If iTW Then
                nTextWidths(r, c) = iInvertedTextWidthsNew(c, r)
            End If
        Next c
    Next r
    
    GridTextMatrixNormalized = iGridTextMatrixNormalized
    
End Function

Private Sub SetFormatOptionsSettingsOnPrintFnObject()
    mPrintFnObject.ColorMode = mPrintGridFormatSettings.ColorMode
    mPrintFnObject.PrintPageNumbers = mPrintGridFormatSettings.PrintPageNumbers
    Set mPrintFnObject.PageNumbersFont = mPrintGridFormatSettings.PageNumbersFont
    mPrintFnObject.PageNumbersPosition = mPrintGridFormatSettings.PageNumbersPosition
    mPrintFnObject.PageNumbersForeColor = mPrintGridFormatSettings.PageNumbersForeColor
    mPrintFnObject.PageNumbersFormat = mPrintFnObject.GetPredefinedPageNumbersFormatString(mPrintGridFormatSettings.PageNumbersFormatLong)
    mPrintFnObject.ScalePercent = mPrintGridFormatSettings.ScalePercent
End Sub

Public Function GetPredefinedPageNumbersFormatString(nIndex As Long) As String
    GetPredefinedPageNumbersFormatString = mPrintFnObject.GetPredefinedPageNumbersFormatString(nIndex)
End Function

Public Property Get GetPredefinedPageNumbersFormatStringsCount() As Long
Attribute GetPredefinedPageNumbersFormatStringsCount.VB_MemberFlags = "40"
    GetPredefinedPageNumbersFormatStringsCount = mPrintFnObject.GetPredefinedPageNumbersFormatStringsCount
End Property

Public Function GetGridReportStyle(nStyleID As String) As GridReportStyle
    Set GetGridReportStyle = New GridReportStyle
    
    Select Case nStyleID
        Case "Default1"
            GetGridReportStyle.Tag = "Default1"
            
            GetGridReportStyle.LineWidth = 3
            GetGridReportStyle.PrintHeadersBackground = True
            GetGridReportStyle.HeadersBackgroundColor = RGB(239, 239, 239)
            GetGridReportStyle.PrintFixedColsBackground = False
            GetGridReportStyle.PrintOtherBackgrounds = True
            
            GetGridReportStyle.PrintOuterBorder = True
            GetGridReportStyle.OuterBorderColor = RGB(31, 31, 31)
            GetGridReportStyle.PrintHeadersBorder = False
            GetGridReportStyle.HeadersBorderColor = RGB(207, 207, 207)
            GetGridReportStyle.PrintColumnsDataLines = True
            GetGridReportStyle.ColumnsDataLinesColor = RGB(207, 207, 207)
            GetGridReportStyle.PrintColumnsHeadersLines = True
            GetGridReportStyle.ColumnsHeadersLinesColor = RGB(207, 207, 207)
            GetGridReportStyle.PrintRowsLines = True
            GetGridReportStyle.RowsLinesColor = RGB(207, 207, 207)
            GetGridReportStyle.PrintHeadersSeparatorLine = True
            GetGridReportStyle.LineWidthHeadersSeparatorLine = 3
        
        Case "Default2"
            GetGridReportStyle.Tag = "Default" & CStr(mDefaultReportStyle)
            
            GetGridReportStyle.LineWidth = 3
            GetGridReportStyle.PrintHeadersBackground = True
            GetGridReportStyle.HeadersBackgroundColor = RGB(244, 249, 255)
            GetGridReportStyle.PrintFixedColsBackground = False
            GetGridReportStyle.PrintOtherBackgrounds = True
            
            GetGridReportStyle.PrintOuterBorder = False
            GetGridReportStyle.OuterBorderColor = RGB(40, 5, 207)
            GetGridReportStyle.PrintHeadersBorder = True
            GetGridReportStyle.HeadersBorderColor = RGB(29, 4, 145)
            GetGridReportStyle.PrintColumnsDataLines = True
            GetGridReportStyle.ColumnsDataLinesColor = RGB(177, 162, 253)
            GetGridReportStyle.PrintColumnsHeadersLines = True
            GetGridReportStyle.ColumnsHeadersLinesColor = RGB(177, 162, 253)
            GetGridReportStyle.PrintRowsLines = False
            GetGridReportStyle.RowsLinesColor = RGB(230, 224, 254)
            GetGridReportStyle.PrintHeadersSeparatorLine = True
            GetGridReportStyle.LineWidthHeadersSeparatorLine = 3
        
        Case "Default3"
            GetGridReportStyle.Tag = "Default3"
            
            GetGridReportStyle.LineWidth = 3
            GetGridReportStyle.PrintHeadersBackground = True
            GetGridReportStyle.HeadersBackgroundColor = RGB(244, 249, 255)
            GetGridReportStyle.PrintFixedColsBackground = False
            GetGridReportStyle.PrintOtherBackgrounds = True
            
            GetGridReportStyle.PrintOuterBorder = False
            GetGridReportStyle.OuterBorderColor = RGB(40, 5, 207)
            GetGridReportStyle.PrintHeadersBorder = False
            GetGridReportStyle.HeadersBorderColor = RGB(29, 4, 145)
            GetGridReportStyle.PrintColumnsDataLines = True
            GetGridReportStyle.ColumnsDataLinesColor = RGB(177, 162, 253)
            GetGridReportStyle.PrintColumnsHeadersLines = True
            GetGridReportStyle.ColumnsHeadersLinesColor = RGB(177, 162, 253)
            GetGridReportStyle.PrintRowsLines = False
            GetGridReportStyle.RowsLinesColor = RGB(230, 224, 254)
            GetGridReportStyle.PrintHeadersSeparatorLine = True
            GetGridReportStyle.LineWidthHeadersSeparatorLine = 10
        
        Case "Default4"
            GetGridReportStyle.Tag = "Default4"
            
            GetGridReportStyle.LineWidth = 3
            GetGridReportStyle.PrintHeadersBackground = True
            GetGridReportStyle.HeadersBackgroundColor = RGB(252, 243, 243)
            GetGridReportStyle.PrintFixedColsBackground = False
            GetGridReportStyle.PrintOtherBackgrounds = True
            
            GetGridReportStyle.PrintOuterBorder = False
            GetGridReportStyle.OuterBorderColor = RGB(139, 40, 35)
            GetGridReportStyle.PrintHeadersBorder = True
            GetGridReportStyle.HeadersBorderColor = RGB(139, 40, 35)
            GetGridReportStyle.PrintColumnsDataLines = True
            GetGridReportStyle.ColumnsDataLinesColor = RGB(237, 197, 194)
            GetGridReportStyle.PrintColumnsHeadersLines = True
            GetGridReportStyle.ColumnsHeadersLinesColor = RGB(237, 197, 194)
            GetGridReportStyle.PrintRowsLines = False
            GetGridReportStyle.RowsLinesColor = RGB(247, 230, 227)
            GetGridReportStyle.PrintHeadersSeparatorLine = True
            GetGridReportStyle.LineWidthHeadersSeparatorLine = 3
        
        Case "Default5"
            GetGridReportStyle.Tag = "Default5"
            
            GetGridReportStyle.LineWidth = 3
            GetGridReportStyle.PrintHeadersBackground = True
            GetGridReportStyle.HeadersBackgroundColor = RGB(244, 249, 255)
            GetGridReportStyle.PrintFixedColsBackground = False
            GetGridReportStyle.PrintOtherBackgrounds = True
            
            GetGridReportStyle.PrintOuterBorder = True
            GetGridReportStyle.OuterBorderColor = RGB(40, 5, 207)
            GetGridReportStyle.PrintHeadersBorder = False
            GetGridReportStyle.HeadersBorderColor = RGB(40, 5, 207)
            GetGridReportStyle.PrintColumnsDataLines = True
            GetGridReportStyle.ColumnsDataLinesColor = RGB(177, 162, 253)
            GetGridReportStyle.PrintColumnsHeadersLines = True
            GetGridReportStyle.ColumnsHeadersLinesColor = RGB(40, 5, 207)
            GetGridReportStyle.PrintRowsLines = False
            GetGridReportStyle.RowsLinesColor = RGB(230, 224, 254)
            GetGridReportStyle.PrintHeadersSeparatorLine = True
            GetGridReportStyle.LineWidthHeadersSeparatorLine = 3
        
        Case "Default6"
            GetGridReportStyle.Tag = "Default6"
            
            GetGridReportStyle.LineWidth = 3
            GetGridReportStyle.PrintHeadersBackground = True
            GetGridReportStyle.HeadersBackgroundColor = RGB(244, 249, 255)
            GetGridReportStyle.PrintFixedColsBackground = False
            GetGridReportStyle.PrintOtherBackgrounds = True
            
            GetGridReportStyle.PrintOuterBorder = True
            GetGridReportStyle.OuterBorderColor = RGB(40, 5, 207)
            GetGridReportStyle.PrintHeadersBorder = False
            GetGridReportStyle.HeadersBorderColor = RGB(40, 5, 207)
            GetGridReportStyle.PrintColumnsDataLines = True
            GetGridReportStyle.ColumnsDataLinesColor = RGB(177, 162, 253)
            GetGridReportStyle.PrintColumnsHeadersLines = True
            GetGridReportStyle.ColumnsHeadersLinesColor = RGB(177, 162, 253)
            GetGridReportStyle.PrintRowsLines = True
            GetGridReportStyle.RowsLinesColor = RGB(230, 224, 254)
            GetGridReportStyle.PrintHeadersSeparatorLine = True
            GetGridReportStyle.LineWidthHeadersSeparatorLine = 3
            
        Case "Default7"
            GetGridReportStyle.Tag = "Default7"
            
            GetGridReportStyle.LineWidth = 3
            GetGridReportStyle.PrintHeadersBackground = True
            GetGridReportStyle.HeadersBackgroundColor = RGB(252, 243, 243)
            GetGridReportStyle.PrintFixedColsBackground = False
            GetGridReportStyle.PrintOtherBackgrounds = True
            
            GetGridReportStyle.PrintOuterBorder = True
            GetGridReportStyle.OuterBorderColor = RGB(139, 40, 35)
            GetGridReportStyle.PrintHeadersBorder = False
            GetGridReportStyle.HeadersBorderColor = RGB(139, 40, 35)
            GetGridReportStyle.PrintColumnsDataLines = True
            GetGridReportStyle.ColumnsDataLinesColor = RGB(237, 197, 194)
            GetGridReportStyle.PrintColumnsHeadersLines = True
            GetGridReportStyle.ColumnsHeadersLinesColor = RGB(237, 197, 194)
            GetGridReportStyle.PrintRowsLines = True
            GetGridReportStyle.RowsLinesColor = RGB(247, 230, 227)
            GetGridReportStyle.PrintHeadersSeparatorLine = True
            GetGridReportStyle.LineWidthHeadersSeparatorLine = 3
        
        Case Else
            If Val(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintOuterBorder", -44)) <> -44 Then
                GetGridReportStyle.Tag = nStyleID
                
                GetGridReportStyle.PrintOuterBorder = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintOuterBorder", -1))
                GetGridReportStyle.PrintHeadersBorder = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintHeadersBorder", -1))
                GetGridReportStyle.PrintColumnsDataLines = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintColumnsDataLines", -1))
                GetGridReportStyle.PrintColumnsHeadersLines = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintColumnsHeadersLines", -1))
                GetGridReportStyle.PrintRowsLines = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintRowsLines", -1))
                GetGridReportStyle.PrintHeadersSeparatorLine = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintHeadersSeparatorLine", -1))
                GetGridReportStyle.PrintHeadersBackground = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintHeadersBackground", -1))
                GetGridReportStyle.PrintFixedColsBackground = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintFixedColsBackground", 0))
                GetGridReportStyle.PrintOtherBackgrounds = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintOtherBackgrounds", -1))
                GetGridReportStyle.LineWidth = GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_LineWidth", 3)
                GetGridReportStyle.LineWidthHeadersSeparatorLine = GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_LineWidthHeadersSeparatorLine", 3)
            
                GetGridReportStyle.OuterBorderColor = Val(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_OuterBorderColor", RGB(31, 31, 31)))
                GetGridReportStyle.ColumnsDataLinesColor = Val(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_ColumnsDataLinesColor", RGB(207, 207, 207)))
                GetGridReportStyle.ColumnsHeadersLinesColor = Val(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_ColumnsHeadersLinesColor", RGB(207, 207, 207)))
                GetGridReportStyle.RowsLinesColor = Val(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_RowsLinesColor", RGB(207, 207, 207)))
                GetGridReportStyle.HeadersBorderColor = Val(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_HeadersBorderColor", RGB(207, 207, 207)))
                GetGridReportStyle.HeadersBackgroundColor = Val(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_HeadersBackgroundColor", RGB(239, 239, 239)))
            End If
    End Select
End Function

Private Sub SaveGridReportStyle()
    Dim iStyleID As String
    Dim iNumberNewCustomStyle As Long
    
    iStyleID = GetStyleID(mGridReportStyle, iNumberNewCustomStyle)
    
    If iStyleID = "" Then
        If iNumberNewCustomStyle > 10 Then
            iNumberNewCustomStyle = 1
            iNumberNewCustomStyle = Val(GetSetting(AppNameForRegistry, "PrintingSettings", "LastCustomGridReportStyleNumber", 0)) + 1
            If iNumberNewCustomStyle > 10 Then
                iNumberNewCustomStyle = 1
            End If
            SaveSetting AppNameForRegistry, "PrintingSettings", "LastCustomGridReportStyleNumber", iNumberNewCustomStyle
        End If
        iStyleID = "Custom" & iNumberNewCustomStyle
        
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_LineWidth", mGridReportStyle.LineWidth
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintHeadersBackground", CLng(mGridReportStyle.PrintHeadersBackground)
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_HeadersBackgroundColor", mGridReportStyle.HeadersBackgroundColor
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintFixedColsBackground", CLng(mGridReportStyle.PrintFixedColsBackground)
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintOtherBackgrounds", CLng(mGridReportStyle.PrintOtherBackgrounds)
        
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintOuterBorder", CLng(mGridReportStyle.PrintOuterBorder)
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_OuterBorderColor", mGridReportStyle.OuterBorderColor
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintHeadersBorder", CLng(mGridReportStyle.PrintHeadersBorder)
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_HeadersBorderColor", mGridReportStyle.HeadersBorderColor
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintColumnsDataLines", CLng(mGridReportStyle.PrintColumnsDataLines)
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_ColumnsDataLinesColor", mGridReportStyle.ColumnsDataLinesColor
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintColumnsHeadersLines", CLng(mGridReportStyle.PrintColumnsHeadersLines)
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_ColumnsHeadersLinesColor", mGridReportStyle.ColumnsHeadersLinesColor
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintRowsLines", CLng(mGridReportStyle.PrintRowsLines)
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_RowsLinesColor", mGridReportStyle.RowsLinesColor
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintHeadersSeparatorLine", CLng(mGridReportStyle.PrintHeadersSeparatorLine)
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_LineWidthHeadersSeparatorLine", mGridReportStyle.LineWidthHeadersSeparatorLine
        
    End If
    
    SaveSetting AppNameForRegistry, "PrintingSettings", Keyname & "_GridReportStyle", iStyleID
    
End Sub

Private Function GridReportStylesAreEqual(nGridReportStyle1 As GridReportStyle, nGridReportStyle2 As GridReportStyle) As Boolean
    Dim iChange As Boolean
    
    If nGridReportStyle1.LineWidth <> nGridReportStyle2.LineWidth Then
        iChange = True
    ElseIf nGridReportStyle1.PrintHeadersBackground <> nGridReportStyle2.PrintHeadersBackground Then
        iChange = True
    ElseIf nGridReportStyle1.HeadersBackgroundColor <> nGridReportStyle2.HeadersBackgroundColor Then
        iChange = True
    ElseIf nGridReportStyle1.PrintFixedColsBackground <> nGridReportStyle2.PrintFixedColsBackground Then
        iChange = True
    ElseIf nGridReportStyle1.PrintOtherBackgrounds <> nGridReportStyle2.PrintOtherBackgrounds Then
        iChange = True
    ElseIf nGridReportStyle1.PrintOuterBorder <> nGridReportStyle2.PrintOuterBorder Then
        iChange = True
    ElseIf nGridReportStyle1.OuterBorderColor <> nGridReportStyle2.OuterBorderColor Then
        iChange = True
    ElseIf nGridReportStyle1.PrintHeadersBorder <> nGridReportStyle2.PrintHeadersBorder Then
        iChange = True
    ElseIf nGridReportStyle1.HeadersBorderColor <> nGridReportStyle2.HeadersBorderColor Then
        iChange = True
    ElseIf nGridReportStyle1.PrintColumnsDataLines <> nGridReportStyle2.PrintColumnsDataLines Then
        iChange = True
    ElseIf nGridReportStyle1.ColumnsDataLinesColor <> nGridReportStyle2.ColumnsDataLinesColor Then
        iChange = True
    ElseIf nGridReportStyle1.PrintColumnsHeadersLines <> nGridReportStyle2.PrintColumnsHeadersLines Then
        iChange = True
    ElseIf nGridReportStyle1.ColumnsHeadersLinesColor <> nGridReportStyle2.ColumnsHeadersLinesColor Then
        iChange = True
    ElseIf nGridReportStyle1.PrintRowsLines <> nGridReportStyle2.PrintRowsLines Then
        iChange = True
    ElseIf nGridReportStyle1.RowsLinesColor <> nGridReportStyle2.RowsLinesColor Then
        iChange = True
    ElseIf nGridReportStyle1.PrintHeadersSeparatorLine <> nGridReportStyle2.PrintHeadersSeparatorLine Then
        iChange = True
    ElseIf nGridReportStyle1.LineWidthHeadersSeparatorLine <> nGridReportStyle2.LineWidthHeadersSeparatorLine Then
        iChange = True
    End If
    
    GridReportStylesAreEqual = Not iChange
End Function


Public Function GetStyleID(nGridReportStyle As GridReportStyle, nNumberNewCustomStyle As Long) As String
    Dim c As Long
    Dim iGridReportStyle As GridReportStyle
    Dim iStyleID As String
    
    c = 1
    Set iGridReportStyle = GetGridReportStyle("Default" & c)
    Do Until iGridReportStyle.Tag = ""
        If GridReportStylesAreEqual(iGridReportStyle, nGridReportStyle) Then
            iStyleID = iGridReportStyle.Tag
            Exit Do
        End If
        c = c + 1
        Set iGridReportStyle = GetGridReportStyle("Default" & c)
    Loop
    
    If iStyleID = "" Then
        c = 1
        Set iGridReportStyle = GetGridReportStyle("Custom" & c)
        Do Until iGridReportStyle.Tag = ""
            If GridReportStylesAreEqual(iGridReportStyle, nGridReportStyle) Then
                iStyleID = iGridReportStyle.Tag
                Exit Do
            End If
            c = c + 1
            Set iGridReportStyle = GetGridReportStyle("Custom" & c)
        Loop
    End If
    
    nNumberNewCustomStyle = c
    GetStyleID = iStyleID
End Function

Public Property Get DefaultFormatSettings() As PrintGridFormatSettings
    Set DefaultFormatSettings = mPrintGridFormatSettings
End Property


Public Property Let Orientation(nValue As gfnOrientation)
    If (nValue < gfnPrinterDefault) Or (nValue > gfnDecideOnReportWidth) Then Exit Property
    If nValue <> mOrientation Then
        mOrientation = nValue
        If (mOrientation > 0) And (mOrientation < 3) Then
            mPrintFnObject.Orientation = mOrientation
        End If
        mDefaultReportOrientation = mPrintFnObject.Orientation
        mEnableAutoOrientation = (mOrientation = gfnDecideOnReportWidth)
    End If
End Property

Public Property Get Orientation() As gfnOrientation
    Orientation = mOrientation
End Property


Public Property Let ScalePercent(nValue As Long)
    mPrintGridFormatSettings.ScalePercent = nValue
End Property

Public Property Get ScalePercent() As Long
    ScalePercent = mPrintGridFormatSettings.ScalePercent
End Property


Private Sub mPrintFnObject_OrientationChange(ByVal NewOrientation As Long)
    mOrientationChangedToLandscapeAutomatically = False
    mAuxAllowAutoOrientation = False
    RaiseEvent OrientationChange(NewOrientation)
End Sub

Private Sub mPrintGridFormatSettings_ScalePercentChange()
    mPrintFnObject.ScalePercent = mPrintGridFormatSettings.ScalePercent
End Sub


Public Property Let CopyToClipboardMode(nValue As gfnCopyToClipboardModeOptions)
    mCopyToClipboardMode = nValue
End Property

Public Property Get CopyToClipboardMode() As gfnCopyToClipboardModeOptions
    CopyToClipboardMode = mCopyToClipboardMode
End Property


Public Property Let SpecialSeparatorCharacters(nValue As String)
    mSpecialSeparatorCharacters = nValue
End Property

Public Property Get SpecialSeparatorCharacters() As String
    SpecialSeparatorCharacters = mSpecialSeparatorCharacters
End Property


Public Property Let ScrollWithMouseWheel(nValue As gfnScrollWithMouseWheelSettings)
    If nValue <> mScrollWithMouseWheel Then
        mScrollWithMouseWheel = nValue
        
        If mAmbientUserMode Then
            If mScrollWithMouseWheel <> 0 Then
                If Not mForm Is Nothing Then
                    Set mMouseWheelNotifierObject = New MouseWheelNotifierObject
                    mMouseWheelNotifierObject.SetForm mForm
                Else
                    If mGridExplicitelySet Then
                        If Not mGridParent Is Nothing Then
                            Set mMouseWheelNotifierObject = New MouseWheelNotifierObject
                            mMouseWheelNotifierObject.SetForm , GetParentFormHwnd(mGridParent.hWnd)
                        Else
                            If Not CurrentGrid Is Nothing Then
                                Set mMouseWheelNotifierObject = New MouseWheelNotifierObject
                                mMouseWheelNotifierObject.SetForm , GetParentFormHwnd(CurrentGrid.hWnd)
                            End If
                        End If
                    End If
                End If
            Else
                Set mMouseWheelNotifierObject = Nothing
            End If
        End If
    End If
End Property

Public Property Get ScrollWithMouseWheel() As gfnScrollWithMouseWheelSettings
    ScrollWithMouseWheel = mScrollWithMouseWheel
End Property


Public Property Let IgnoreEmptyRowsAtTheEnd(nValue As Boolean)
    If nValue <> mIgnoreEmptyRowsAtTheEnd Then
        mIgnoreEmptyRowsAtTheEnd = nValue
    End If
End Property

Public Property Get IgnoreEmptyRowsAtTheEnd() As Boolean
    IgnoreEmptyRowsAtTheEnd = mIgnoreEmptyRowsAtTheEnd
End Property


Private Sub mMouseWheelNotifierObject_MouseWheelRotation(Direction As Long, Handled As Boolean)
    Static sGrid As Object
    Dim iNewGrid As Object
    Static sScrollStep As Long
    Static sGridNamePrev As String
    Static sGridRowsPrev As Long
    Dim iGridIsActiveControl As Boolean
    Dim iAuxHeight As Long
    Dim r As Long
    Dim iHasScrollBar As Boolean
    
    If mScrollWithMouseWheel = 0 Then Exit Sub
    
On Error GoTo TheExit:
    
    If sGrid Is Nothing Then
        Set sGrid = mCurrentGrid
    End If
    If sGrid Is Nothing Then
        Set iNewGrid = GetGridControl2(True)
        If iNewGrid Is Nothing Then Exit Sub
        Set sGrid = iNewGrid
    End If
    On Error Resume Next
    iGridIsActiveControl = sGrid.Parent.ActiveControl Is sGrid
    On Error GoTo TheExit:
    If Not iGridIsActiveControl Then
        If mScrollWithMouseWheel = gfnScrollOnlyOnFocus Then
            Exit Sub
        Else
            Set iNewGrid = GetGridControl2(True)
            If iNewGrid Is Nothing Then Exit Sub
            Set sGrid = iNewGrid
        End If
    Else
        If Not IsWindowVisibleOnScreen(sGrid.hWnd, True) Then
            Set iNewGrid = GetGridControl2(True)
            If iNewGrid Is Nothing Then Exit Sub
            Set sGrid = iNewGrid
        End If
    End If
    
    iHasScrollBar = mGlobals.IsShowingVerticalScrollBar(sGrid) Or Not mGridExplicitelySetIsTypeSupported
    If iHasScrollBar Then
        If (sGridNamePrev <> sGrid.Name) Or (sGridRowsPrev <> sGrid.Rows) Then
            ' calculate the scroll step based on Grid variables
            iAuxHeight = 0
            For r = 0 To sGrid.Rows - 1
                iAuxHeight = iAuxHeight + sGrid.RowHeight(r)
                If iAuxHeight >= sGrid.Height Then
                    Exit For
                End If
            Next r
            sScrollStep = r / 4
            If sScrollStep = 0 Then sScrollStep = 1
        End If
        
        If Direction = 1 Then
            If Not sGrid.RowIsVisible(sGrid.Rows - 1) Then
                If (sGrid.TopRow + sScrollStep) > (sGrid.Rows - 1) Then
                    sGrid.TopRow = sGrid.Rows - 1
                Else
                    sGrid.TopRow = sGrid.TopRow + sScrollStep
                End If
            Else
                sGrid.TopRow = sGrid.Rows - 1
            End If
        Else
            If Not sGrid.RowIsVisible(sGrid.FixedRows) Then
                If (sGrid.TopRow - sScrollStep) < 1 Then
                    sGrid.TopRow = sGrid.FixedRows
                Else
                    sGrid.TopRow = sGrid.TopRow - sScrollStep
                End If
            Else
                sGrid.TopRow = sGrid.FixedRows
            End If
        End If
        Handled = True
    End If
    
    sGridNamePrev = sGrid.Name
    sGridRowsPrev = sGrid.Rows
    Exit Sub
    
TheExit:
    Handled = False
End Sub


Private Function WordWrappedInMax3Lines(nText As String) As String
    Dim iWordCount As Long
    Dim iLines() As String
    Dim iPos1 As Long
    Dim iPos2 As Long
    Dim iLen As Long
    Dim iPos1a As Long
    Dim iPos1b As Long
    Dim iPos2a As Long
    Dim iPos2b As Long
    Dim iPos1f As Long
    Dim iPos2f As Long
    Dim iText As String
    
    iText = Trim2(WithoutConsecutiveSpaces(nText))
    iLines = Split(iText, " ")
    iWordCount = UBound(iLines) + 1
    
    If iWordCount <= 3 Then
        WordWrappedInMax3Lines = Replace(iText, " ", vbCrLf)
    Else
        iLen = Len(nText)
        iPos1 = Round(iLen / 3)
        iPos2 = Round(iLen / 3 * 2)
        iPos1a = InStr(iPos1, nText, " ")
        iPos1b = InStrRev(nText, " ", iPos1)
        If iPos1b = 0 Then
            iPos1b = iPos1a
        End If
        If iPos1a = 0 Then
            iPos1a = iPos1b
        End If
        iPos2a = InStr(iPos2, nText, " ")
        iPos2b = InStrRev(nText, " ", iPos2)
        If iPos2b = 0 Then
            iPos2b = iPos2a
        End If
        If iPos2a = 0 Then
            iPos2a = iPos2b
        End If
        If Abs(iPos1a - iPos1) < Abs(iPos1b - iPos1) Then
           iPos1f = iPos1a
        Else
            iPos1f = iPos1b
        End If
        If Abs(iPos2a - iPos2) < Abs(iPos2b - iPos2) Then
           iPos2f = iPos2a
        Else
            If iPos2b > (iPos1f + 1) Then
                iPos2f = iPos2b
            Else
                iPos2f = iPos2a
            End If
        End If
        If Mid$(nText, iPos1f, 1) <> " " Then Stop
        If Mid$(nText, iPos2f, 1) <> " " Then Stop
        WordWrappedInMax3Lines = Left$(nText, iPos1f - 1) & vbCrLf & Mid$(nText, iPos1f + 1, iPos2f - iPos1f - 1) & vbCrLf & Right$(nText, iLen - iPos2f)
    End If
End Function

Private Sub ResetReportSpecificVariables()
    mPrintFormatSettingsLoaded = False
    mDefaultReportOrientation = mPrintFnObject.Orientation
    mUsingDefaultReportStyle = False
    mUserChangedScale = False
    mAuxAllowAutoOrientation = False
    mOrientationChangedToLandscapeAutomatically = False

    mSave_EnableAutoOrientation = False
    mSave_ScalePercent = False
    mSave_ColorMode = False
    mSave_GridAlign = False
    mSave_PrintPageNumbers = False
    mSave_PageNumbersPosition = False
    mSave_PageNumbersFormatLong = False
    mSave_PageNumbersFont = False
    mSave_PageNumbersForeColor = False
    mSave_HeadingFont = False
    mSave_HeadingFontColor = False
    mSave_SubheadingFont = False
    mSave_SubheadingFontColor = False
    mSave_OtherTextsFont = False
    mSave_OtherTextsFontColor = False
End Sub

Public Property Get ThereAreHiddenCols(Optional nGrid As Object) As Boolean
    Dim iGH As cGridHandler
    
    If nGrid Is Nothing Then
        If Not mCurrentGridHandler Is Nothing Then
            ThereAreHiddenCols = mCurrentGridHandler.ThereAreHiddenCols
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            ThereAreHiddenCols = iGH.ThereAreHiddenCols
        End If
    End If
End Property

Public Sub UnHideAllCols(Optional nGrid As Object)
    Dim iGH As cGridHandler
    
    If nGrid Is Nothing Then
        If Not mCurrentGridHandler Is Nothing Then
            mCurrentGridHandler.UnHideAllCols
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If iGH Is Nothing Then
            AddGridHandler nGrid
        End If
        iGH.UnHideAllCols
    End If
End Sub

Public Sub HideCol(nCol As Long, Optional nGrid As Object)
    Dim iGH As cGridHandler
    
    If nGrid Is Nothing Then
        If Not mCurrentGridHandler Is Nothing Then
            mCurrentGridHandler.HideCol nCol
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If iGH Is Nothing Then
            AddGridHandler nGrid
        End If
        iGH.HideCol nCol
    End If
End Sub

Private Property Get FileName2() As String
    FileName2 = FileName
    If FileName2 = "" Then
        FileName2 = mHeading
    End If
    If Right$(FileName2, 1) = ":" Then
        FileName2 = Left$(FileName2, Len(FileName2) - 1)
    End If
    FileName2 = ValidFileName(FileName2)
End Property
    
Private Function FirstWords(nText As String, ByVal nWords As Long)
    Dim iWords() As String
    Dim c As Long
    
    If nText = "" Then Exit Function
    nWords = nWords - 1
    iWords = Split(nText, " ")
    For c = 0 To UBound(iWords)
        FirstWords = FirstWords & iWords(c) & " "
        If c = nWords Then Exit For
    Next c
    FirstWords = Left$(FirstWords, Len(FirstWords) - 1)
End Function

Public Property Get PrintGridFormatSettings() As PrintGridFormatSettings
    Set PrintGridFormatSettings = mPrintGridFormatSettings
End Property


Public Property Let AfterSaveAction(nValue As gfnAfterSaveActionSettings)
    mAfterSaveAction = nValue
End Property

Public Property Get AfterSaveAction() As gfnAfterSaveActionSettings
    AfterSaveAction = mAfterSaveAction
End Property


Public Property Let MergeCellsExcel(nValue As Boolean)
    mMergeCellsExcel = nValue
End Property

Public Property Get MergeCellsExcel() As Boolean
    MergeCellsExcel = mMergeCellsExcel
End Property


Public Property Get DefaultReportStyle() As Long
    DefaultReportStyle = mDefaultReportStyle
End Property

Public Property Let DefaultReportStyle(nValue As Long)
    mDefaultReportStyle = nValue
End Property


Public Property Let ShowCopyConfirmationMessage(nValue As Boolean)
    mShowCopyConfirmationMessage = nValue
End Property

Public Property Get ShowCopyConfirmationMessage() As Boolean
    ShowCopyConfirmationMessage = mShowCopyConfirmationMessage
End Property


Public Property Let CopyConfirmationMessage(nValue As String)
    mCopyConfirmationMessage = nValue
End Property

Public Property Get CopyConfirmationMessage() As String
    CopyConfirmationMessage = mCopyConfirmationMessage
End Property


Public Property Let AutoSelect125PercentScaleOnSmallGrids(nValue As Boolean)
    mAutoSelect125PercentScaleOnSmallGrids = nValue
End Property

Public Property Get AutoSelect125PercentScaleOnSmallGrids() As Boolean
    AutoSelect125PercentScaleOnSmallGrids = mAutoSelect125PercentScaleOnSmallGrids
End Property


Private Sub DrawGrid(nCanceled As Boolean)
    Dim c As Long
    Dim iGridNormalWidth As Long
    Dim iGridWidthProportion As Single
    Dim iScaleWidth As Long
    Dim iScaleHeight As Long
    Dim iFontsProportion As Single
    Dim iLongerTexts() As Long
    Dim iLongerTextsRow() As Long
    Dim r As Long
    Dim iGridReDraw As Boolean
    Dim iLen As Long
    Dim iSng As Single
    Dim iCellMarginsX As Long
    Dim iCellMarginsY As Long
    Dim iGridCols As Long
    Dim iGridRows As Long
    Dim iAuxLng As Long
    Dim iAuxLng2 As Long
    Dim iAuxLng3 As Long
    Dim iAuxLng4 As Long
    Dim iGridStatus_Row As Long
    Dim iGridStatus_RowSel As Long
    Dim iGridStatus_Col As Long
    Dim iGridStatus_RowColSel As Long
    Dim iGridAlign As Long
    Dim iFixedRows As Long
    Dim iFixedCols As Long
    Dim iFontAttibutes As cFontAttributes
    Dim iDoNotTouchColumnsWidthIfNotNecessary As Boolean
    Dim iDefAuxFont As StdFont
    Dim iAuxCurFont As StdFont
    
    Dim iDone As Boolean
    Dim iIgnore As Boolean
    
    Dim iColumnHeadersTextWidths() As Long
    Dim iColumnHeadersTextWidthsWithoutReducingTheFont() As Long
    Dim iColumnDataWordWraps() As Boolean
    Dim iColumnTextWidthsGeneralFontNormal() As Long
    Dim iColumnTextWidthsGeneralFontReduced() As Long
    Dim iColumnIsFixed() As Boolean
    Dim iColumnDataTextWidths() As Long
    Dim iAuxColumnTextWidths() As Long
    Dim iTotalTextWidth As Long
    Dim iFixedWidth As Long
    Dim iAuxColumnsToWordBreak() As Boolean
    Dim iAuxWidthToApply() As Long
    
    Dim iWordWrapTheColumnHeaders As Boolean
    Dim iReduceTheColumnsHeadersFont As Boolean
    Dim iWordWrapTheColumnData As Boolean
    Dim iReduceTheColumnsDataFont As Boolean
    Dim iWordBreakTheColumnTexts As Boolean
    Dim iAuxBool As Boolean
    Dim iDoNotReduceFontSize As Boolean
    Dim iLastRowToPrint As Long
    Dim iProposedFontProportion As Single
    Dim iColWidthsPrev() As Long
    Dim iAuxSng As Single
    Dim iConsiderGridWordWrap As Boolean
    Dim iLines() As String
    
    Dim iGrid As Object   ' to have intellisense replace 'Object' with MsFlexGrid or MSHFlexGrid and add the proper reference...
    
    Set iGrid = mCurrentGrid
    
    Set iDefAuxFont = Printer2.Font
    
    Dim iT1 As Double
    iT1 = timer
    
    iGridStatus_Row = iGrid.Row
    iGridStatus_RowSel = iGrid.RowSel
    iGridStatus_Col = iGrid.col
    iGridStatus_RowColSel = iGrid.ColSel
    iGridReDraw = iGrid.Redraw
    iGrid.Redraw = False
    
    'iConsiderGridWordWrap = True '(esta la había sacado -comentado- no sé por qué, tal vez solo porque me parecía que estaba demás ya que de todas maneras aplica el WordWrap automáticamente cuando hace falta, pero la tuve que volver a poner porque no lo hacía cuando hay celdas merged con MergeRows)
Start:
    If mGridPrintingData Is Nothing Then
        Set mGridPrintingData = New cGridPrintingData
        mGridPrintingData.SetGrid mCurrentGrid, mGridReportStyle.HeadersBackgroundColor
        If mGridPrintingData.Canceled Then
            nCanceled = True
            iGrid.Redraw = iGridReDraw
            Exit Sub
        End If
    Else
        mGridPrintingData.ResetAddedData mGridReportStyle.HeadersBackgroundColor
    End If
    If mIgnoreEmptyRowsAtTheEnd Then
        iLastRowToPrint = mGridPrintingData.LastRowWithData
    Else
        iLastRowToPrint = iGridRows - 1
    End If
    
    iScaleWidth = Printer2.ScaleWidth
    iScaleHeight = Printer2.ScaleHeight
    iFontsProportion = 1 ' mPrintGridFormatSettings.ScalePercent / 100 * mScalePercent / 100
    mPrintFnObject.HandleMargins = True
  '  mPrintFnObject.ScalePercent = mPrintGridFormatSettings.ScalePercent
    iGridCols = iGrid.Cols
    iGridRows = iGrid.Rows
    iFixedRows = iGrid.FixedRows
    iFixedCols = iGrid.FixedCols
    
    iCellMarginsX = Printer2.ScaleX(1, vbMillimeters, vbTwips)
    iCellMarginsY = Printer2.ScaleY(1, vbMillimeters, vbTwips)
    
    For c = 0 To iGridCols - 1
        If mGridPrintingData.ColIsVisible(c) <> 0 Then
            iAuxLng = mGridPrintingData.ColLargerFontSizeFontIndex(c)
            Set iFontAttibutes = mGridPrintingData.CellFontAttibutesByFontAttibutesIndex(iAuxLng)
            
            Printer2.FontName = iFontAttibutes.Name
            Printer2.FontSize = iFontAttibutes.Size * iFontsProportion
            Printer2.FontBold = iFontAttibutes.Bold
            Printer2.FontItalic = iFontAttibutes.Italic
            Printer2.FontUnderLine = iFontAttibutes.Underline
            Printer2.FontStrikeThru = iFontAttibutes.Strikethrough
            PrinterEx.FontWidth = iFontAttibutes.Width
            
            iAuxLng = Printer2.TextWidth("W")
            mGridPrintingData.MinColWidth(c) = iAuxLng
            Printer2.FontSize = iFontAttibutes.Size * iFontsProportion * 0.75
            iAuxLng = Printer2.TextWidth("W")
            mGridPrintingData.MinColWidthFontReduced(c) = iAuxLng
        End If
    Next c
    
    iGridNormalWidth = 0
    For c = 0 To iGridCols - 1
        iGridNormalWidth = iGridNormalWidth + mGridPrintingData.ColOriginalWidth(c)
    Next c
    
    If (iGridNormalWidth * iFontsProportion) * 0.81 < iScaleWidth Then
        If (iGridNormalWidth * iFontsProportion) > iScaleWidth Then
            iProposedFontProportion = iFontsProportion / (iGridNormalWidth * iFontsProportion) * iScaleWidth
            If iProposedFontProportion < 0.9 Then
                iDoNotReduceFontSize = True
            End If
        End If
OptionA:
        If mPrintGridFormatSettings.GridAlign = 3 Then ' stretch
            iGridWidthProportion = iScaleWidth / iGridNormalWidth
            If iGridWidthProportion < 1.5 Then
                iFontsProportion = iFontsProportion * iGridWidthProportion
            Else
                iFontsProportion = iFontsProportion * 1.5
            End If
            If iFontsProportion > 2 Then
                iFontsProportion = 2
            End If
            iDoNotTouchColumnsWidthIfNotNecessary = True
            mGridPrintingData.SetColsWidthsToProportion iGridWidthProportion
            mGridPrintingData.SetRowsHeightsToProportion iFontsProportion
        Else ' whatever Grid align
            iGridWidthProportion = iFontsProportion
            If (iGridNormalWidth * iGridWidthProportion) > iScaleWidth Then
                iGridWidthProportion = iScaleWidth / iGridNormalWidth
            End If
            If mAutoSelect125PercentScaleOnSmallGrids Then
                If mPrintFnObject.ScalePercent = 100 Then
                    If (iGridNormalWidth * iFontsProportion) < iScaleWidth Then
                        If Not mUserChangedScale And (iFontsProportion = 1) Then
                            If (iGridNormalWidth / iScaleWidth) < (1 / 1.25) Then
                                If mMaxScalePercent >= 125 Then
                                    iGridWidthProportion = iGridWidthProportion * 1.25
                                    iFontsProportion = 1.25
                                    mPrintFnObject.ScalePercent = 125
                                    mPrintGridFormatSettings.ScalePercent = 125
                                    'Printer2.KillDoc
                                    mPrintFnObject.ReSetPrinterExSettings
                                    GoTo Start
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            iGridAlign = mPrintGridFormatSettings.GridAlign + 1
            mGridPrintingData.SetColsWidthsToProportion iGridWidthProportion
            mGridPrintingData.SetRowsHeightsToProportion iFontsProportion
            iDoNotTouchColumnsWidthIfNotNecessary = True
        End If
    Else
        If mAuxAllowAutoOrientation Then
            If mPrintFnObject.Orientation = vbPRORPortrait Then
                mPrintFnObject.Orientation = vbPRORLandscape
                mAuxAllowAutoOrientation = False
                mDefaultReportOrientation = vbPRORLandscape
                mOrientationChangedToLandscapeAutomatically = True
                mAuxAllowAutoOrientation = False
                iScaleWidth = Printer2.ScaleWidth
                iScaleHeight = Printer2.ScaleHeight
'                If (iGridNormalWidth * iFontsProportion) < iScaleWidth Then
'                    GoTo OptionA:
'                End If
                GoTo Start:
            End If
        End If
    End If
    
'        If (iGridNormalWidth * iFontsProportion) < iScaleWidth Then
'            iGridWidthProportion = iScaleWidth / iGridNormalWidth
'
'            ReDim iProposedColumnWidth(iGridCols - 1)
'            iAuxLng = 0
'            For c = 0 To iGridCols - 1
'                iProposedColumnWidth(c) = mGridPrintingData.ColOriginalWidth(c) * iGridWidthProportion
'                iAuxLng = iAuxLng + iProposedColumnWidth(c)
'            Next c
'            If iAuxLng <= iScaleWidth Then
'                mGridPrintingData.SetColsWidthsToProportion iGridWidthProportion
'                mGridPrintingData.SetRowsHeightsToProportion iFontsProportion
'                iDone = True
'            End If
'
'        End If
        
        'If Not iDone Then
        
    '    iConsiderGridWordWrap = False
        
Part2:

    ' find the 5 longer data strings in each column
    ReDim iLongerTexts(iGridCols - 1, 5)
    ReDim iLongerTextsRow(iGridCols - 1, 5)
    For c = 0 To iGridCols - 1
        For r = iFixedRows To iLastRowToPrint
            If mGridPrintingData.RowIsVisible(r) Then
                iIgnore = False
                If (iGrid.MergeCells = flexMergeFree) Or (iGrid.MergeCells = flexMergeRestrictRows) Then
                    If iGrid.MergeRow(r) = True Then
                        iConsiderGridWordWrap = True
                        If c > 0 Then
                            If iGrid.TextMatrix(r, c) = iGrid.TextMatrix(r, c - 1) Then
                                iIgnore = True
                            End If
                        End If
                        If Not iIgnore Then
                            If c < (iGridCols - 1) Then
                                If iGrid.TextMatrix(r, c) = iGrid.TextMatrix(r, c + 1) Then
                                    iIgnore = True
                                End If
                            End If
                        End If
                    Else
                        iConsiderGridWordWrap = False
                    End If
                End If
                
                If Not iIgnore Then
                    iLen = Len(iGrid.TextMatrix(r, c)) * mGridPrintingData.CellFontAttibutes(r, c).Size
                    If iLen > iLongerTexts(c, 5) Then
                        If iLen > iLongerTexts(c, 1) Then
                            iLongerTexts(c, 1) = iLen
                            iLongerTextsRow(c, 1) = r
                        Else
                            If iLen > iLongerTexts(c, 2) Then
                                iLongerTexts(c, 2) = iLen
                                iLongerTextsRow(c, 2) = r
                            Else
                                If iLen > iLongerTexts(c, 3) Then
                                    iLongerTexts(c, 3) = iLen
                                    iLongerTextsRow(c, 3) = r
                                Else
                                    If iLen > iLongerTexts(c, 4) Then
                                        iLongerTexts(c, 4) = iLen
                                        iLongerTextsRow(c, 4) = r
                                    Else
                                        If iLen > iLongerTexts(c, 5) Then
                                            iLongerTexts(c, 5) = iLen
                                            iLongerTextsRow(c, 5) = r
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next r
    Next c
    
    Set Printer2.Font = iGrid.Font
    On Error Resume Next
    PrinterEx.FontWidth = iGrid.FontWidth
    On Error GoTo 0
    ' columns headers widths
    ReDim iColumnHeadersTextWidths(iGridCols - 1)
    ReDim iColumnTextWidthsGeneralFontNormal(iGridCols - 1)
    ReDim iColumnTextWidthsGeneralFontReduced(iGridCols - 1)
    For c = 0 To iGridCols - 1
        If mGridPrintingData.ColIsVisible(c) Then
            iGrid.col = c
            For r = 0 To iFixedRows - 1
                If mGridPrintingData.RowIsVisible(r) Then
                    iGrid.Row = r
                    Printer2.FontName = iGrid.CellFontName
                    Printer2.FontSize = iGrid.CellFontSize * iFontsProportion
                    iSng = iFontsProportion
                    Do Until Printer2.FontSize <= (iGrid.CellFontSize * iSng * 1.001)
                        iFontsProportion = iFontsProportion * 0.995
                        Printer2.FontSize = iGrid.CellFontSize * iFontsProportion
                        If Printer2.FontSize < 2 Then Exit Do
                    Loop
                    Printer2.FontBold = iGrid.CellFontBold
                    Printer2.FontItalic = iGrid.CellFontItalic
                    On Error Resume Next
                    Printer2.FontStrikeThru = iGrid.CellFontStrikeThrough
                    Printer2.FontUnderLine = iGrid.CellFontUnderline
                    PrinterEx.FontWidth = iGrid.CellFontWidth
                    On Error GoTo 0
                    
                    iSng = Printer2.TextWidth(iGrid.Text) + iCellMarginsX * 2 + 15
                    If iSng > iColumnHeadersTextWidths(c) Then
                        iColumnHeadersTextWidths(c) = iSng
                    End If
                End If
            Next r
        End If
    Next c
    
    ' columns data widths
    ReDim iColumnDataTextWidths(iGridCols - 1)
    For c = 0 To iGridCols - 1
        If mGridPrintingData.ColIsVisible(c) Then
            iGrid.col = c
            For r = 1 To 5
                If mGridPrintingData.RowIsVisible(r) Then
                    If iLongerTexts(c, r) <> 0 Then
                        iGrid.Row = iLongerTextsRow(c, r)
                        Printer2.FontName = iGrid.CellFontName
                        Printer2.FontSize = iGrid.CellFontSize * iFontsProportion
                        iSng = iFontsProportion
                        Do Until Printer2.FontSize <= (iGrid.CellFontSize * iSng * 1.001)
                            iFontsProportion = iFontsProportion * 0.995
                            Printer2.FontSize = iGrid.CellFontSize * iFontsProportion
                            If Printer2.FontSize < 2 Then Exit Do
                        Loop
                        Printer2.FontBold = iGrid.CellFontBold
                        Printer2.FontItalic = iGrid.CellFontItalic
                        On Error Resume Next
                        Printer2.FontStrikeThru = iGrid.CellFontStrikeThrough
                        Printer2.FontUnderLine = iGrid.CellFontUnderline
                        PrinterEx.FontWidth = iGrid.CellFontWidth
                        On Error GoTo 0
                    
                        iSng = Printer2.TextWidth(iGrid.Text) + iCellMarginsX * 2 + 15
                        If iSng > iColumnDataTextWidths(c) Then
                            iColumnDataTextWidths(c) = iSng
                        End If
                    End If
                End If
            Next r
        End If
    Next c
    
    ' Column text width considering widtest texts
    iTotalTextWidth = 0
    For c = 0 To iGridCols - 1
        If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
            iTotalTextWidth = iTotalTextWidth + iColumnHeadersTextWidths(c)
            If iDoNotTouchColumnsWidthIfNotNecessary Then
                If iColumnHeadersTextWidths(c) > mGridPrintingData.ColWidth(c) Then
                    iDoNotTouchColumnsWidthIfNotNecessary = False
                End If
            End If
            iColumnTextWidthsGeneralFontNormal(c) = iColumnHeadersTextWidths(c)
        Else
            iTotalTextWidth = iTotalTextWidth + iColumnDataTextWidths(c)
            If iDoNotTouchColumnsWidthIfNotNecessary Then
                If iColumnDataTextWidths(c) > mGridPrintingData.ColWidth(c) Then
                    iDoNotTouchColumnsWidthIfNotNecessary = False
                End If
            End If
            iColumnTextWidthsGeneralFontNormal(c) = iColumnDataTextWidths(c)
        End If
    Next c
    
    If iTotalTextWidth * 1.15 <= iScaleWidth Then
        iProposedFontProportion = 0
    Else
        If iProposedFontProportion <> 0 Then
            iFontsProportion = iProposedFontProportion
            iProposedFontProportion = 0
            GoTo Part2:
        End If
    End If
    
    If iTotalTextWidth <= iScaleWidth Then ' if adjusting the width of the columns to the width of the actual text then all fits
        iDone = True
        If iGridAlign = 0 Then
            If Not iDoNotTouchColumnsWidthIfNotNecessary Then
                If Not mGridPrintingData.ColWidthsAreSet Then
                    mGridPrintingData.SetColsWidthsToProportion iFontsProportion
                End If
                ReDim iColWidthsPrev(iGridCols - 1)
                
                For c = 0 To iGridCols - 1
                    iColWidthsPrev(c) = mGridPrintingData.ColWidth(c)
                    
                    If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                        mGridPrintingData.ColWidth(c) = iColumnHeadersTextWidths(c) ' / iTotalTextWidth * iScaleWidth
                    Else
                        mGridPrintingData.ColWidth(c) = iColumnDataTextWidths(c) '/ iTotalTextWidth * iScaleWidth
                    End If
                Next c
                iAuxLng2 = 0
                iAuxLng = 0
                For c = 0 To iGridCols - 1
                    If mGridPrintingData.ColWidth(c) < (mGridPrintingData.ColOriginalWidth(c) * (iScaleWidth / iTotalTextWidth)) Then
                        iAuxLng2 = iAuxLng2 + mGridPrintingData.ColOriginalWidth(c) * (iScaleWidth / iTotalTextWidth)
                    Else
                        iAuxLng = iAuxLng + mGridPrintingData.ColWidth(c) * (iScaleWidth / iTotalTextWidth)
                    End If
                Next c
                iAuxBool = False
                On Error Resume Next
                For c = 0 To iGridCols - 1
                    If mGridPrintingData.ColWidth(c) * (iScaleWidth / iTotalTextWidth) < mGridPrintingData.ColOriginalWidth(c) Then
                        If (mGridPrintingData.ColOriginalWidth(c) * (iScaleWidth - iAuxLng) / iAuxLng2) < mGridPrintingData.ColWidth(c) Then
                            iAuxBool = True
                        End If
                    End If
                Next c
                If Not iAuxBool Then
                    For c = 0 To iGridCols - 1
                        If mGridPrintingData.ColWidth(c) < (mGridPrintingData.ColOriginalWidth(c) * (iScaleWidth / iTotalTextWidth)) Then
                            mGridPrintingData.ColWidth(c) = mGridPrintingData.ColOriginalWidth(c) * (iScaleWidth - iAuxLng) / iAuxLng2 * (iScaleWidth / iTotalTextWidth)
                        End If
                    Next c
                End If
                On Error GoTo 0
                
                iAuxLng = 0
                For c = 0 To iGridCols - 1
                    iAuxLng = iAuxLng + mGridPrintingData.ColWidth(c)
                Next c
                If iAuxLng < iScaleWidth Then
                    iAuxLng2 = 0
                    For c = 0 To iGridCols - 1
                        If (iColWidthsPrev(c) * iFontsProportion) > mGridPrintingData.ColWidth(c) Then
                            iAuxLng2 = iAuxLng2 + (iColWidthsPrev(c) * iFontsProportion - mGridPrintingData.ColWidth(c))
                        End If
                    Next c
                    If iAuxLng2 <= (iScaleWidth - iAuxLng) Then
                        For c = 0 To iGridCols - 1
                            If (iColWidthsPrev(c) * iFontsProportion) > mGridPrintingData.ColWidth(c) Then
                                mGridPrintingData.ColWidth(c) = iColWidthsPrev(c) * iFontsProportion
                            End If
                        Next c
                    Else
                        iAuxSng = (iScaleWidth - iAuxLng) / iAuxLng2
                        If iAuxSng > 1 Then iAuxSng = 1
                        For c = 0 To iGridCols - 1
                            If (iColWidthsPrev(c) * iFontsProportion) > mGridPrintingData.ColWidth(c) Then
                                mGridPrintingData.ColWidth(c) = mGridPrintingData.ColWidth(c) + (iColWidthsPrev(c) * iFontsProportion - mGridPrintingData.ColWidth(c)) * iAuxSng
                            End If
                        Next c
                    End If
                End If
            End If
        Else
            
            If Not iDoNotTouchColumnsWidthIfNotNecessary Then
                If Not mGridPrintingData.ColWidthsAreSet Then
                    mGridPrintingData.SetColsWidthsToProportion iFontsProportion
                End If
                ReDim iColWidthsPrev(iGridCols - 1)
                
                For c = 0 To iGridCols - 1
                    iColWidthsPrev(c) = mGridPrintingData.ColWidth(c)
                    
                    If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                        mGridPrintingData.ColWidth(c) = iColumnHeadersTextWidths(c) ' / iTotalTextWidth * iScaleWidth
                    Else
                        mGridPrintingData.ColWidth(c) = iColumnDataTextWidths(c) '/ iTotalTextWidth * iScaleWidth
                    End If
                Next c
'                iAuxLng2 = 0
'                iAuxLng = 0
'                For c = 0 To iGridCols - 1
                    'bbbbbbbbbbbbbbbbbbbbb
'                    If mGridPrintingData.ColWidth(c) < mGridPrintingData.ColOriginalWidth(c) * iFontsProportion Then
                        'mGridPrintingData.ColWidth(c) = mGridPrintingData.ColOriginalWidth(c) * iFontsProportion
'                    End If
'                Next c
            End If
            
            
        End If
        mGridPrintingData.SetRowsHeightsToProportion iFontsProportion
    Else
        
'                ReDim iColumnHeaderWordWraps(iGridCols - 1)
        ReDim iColumnDataWordWraps(iGridCols - 1)
        ReDim iColumnHeadersTextWidthsWithoutReducingTheFont(iGridCols - 1)
        For c = 0 To iGridCols - 1
            iColumnHeadersTextWidthsWithoutReducingTheFont(c) = iColumnHeadersTextWidths(c)
        Next c
        
        ' try reducing columns headers font size
        iTotalTextWidth = 0
        For c = 0 To iGridCols - 1
            If mGridPrintingData.ColIsVisible(c) Then
                If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                    iReduceTheColumnsHeadersFont = True And (Not iDoNotReduceFontSize)
                    iColumnHeadersTextWidths(c) = 0
                    For r = 0 To iFixedRows - 1
                        If mGridPrintingData.RowIsVisible(r) Then
                            iGrid.col = c
                            iGrid.Row = r
                            Printer2.FontName = iGrid.CellFontName
                            Printer2.FontSize = iGrid.CellFontSize * iFontsProportion * 0.75
                            Printer2.FontBold = iGrid.CellFontBold
                            Printer2.FontItalic = iGrid.CellFontItalic
                            On Error Resume Next
                            Printer2.FontStrikeThru = iGrid.CellFontStrikeThrough
                            Printer2.FontUnderLine = iGrid.CellFontUnderline
                            PrinterEx.FontWidth = iGrid.CellFontWidth
                            On Error GoTo 0
                            
                            iSng = Printer2.TextWidth(iGrid.Text) + iCellMarginsX * 2 + 15
                            If iSng > iColumnHeadersTextWidths(c) Then
                                iColumnHeadersTextWidths(c) = iSng
                            End If
                        End If
                    Next r
                End If
                
                If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                    iTotalTextWidth = iTotalTextWidth + iColumnHeadersTextWidths(c)
                Else
                    iTotalTextWidth = iTotalTextWidth + iColumnDataTextWidths(c)
                End If
            End If
        Next c
        
        If iReduceTheColumnsHeadersFont Then
            If iTotalTextWidth <= iScaleWidth Then
                iDone = True
                For c = 0 To iGridCols - 1
                    If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                        mGridPrintingData.ColWidth(c) = iColumnHeadersTextWidths(c) / iTotalTextWidth * iScaleWidth
                    Else
                        mGridPrintingData.ColWidth(c) = iColumnDataTextWidths(c) / iTotalTextWidth * iScaleWidth
                    End If
                Next c
                mGridPrintingData.SetRowsHeightsToProportion iFontsProportion
            End If
        End If
        
        If Not iDone Then
            ' try with columns headers wordwrap
            ReDim iAuxColumnTextWidths(iGridCols - 1)
            iTotalTextWidth = 0
            For c = 0 To iGridCols - 1
                If mGridPrintingData.ColIsVisible(c) Then
                    If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                        For r = 0 To iFixedRows - 1
                            If mGridPrintingData.RowIsVisible(r) Then
                                If InStr(iGrid.TextMatrix(r, c), " ") > 0 Then
                                    iWordWrapTheColumnHeaders = True
                                End If
                                iGrid.col = c
                                iGrid.Row = r
                                Printer2.FontName = iGrid.CellFontName
                                If iReduceTheColumnsHeadersFont Then
                                    Printer2.FontSize = iGrid.CellFontSize * iFontsProportion * 0.75
                                Else
                                    iSng = iFontsProportion
                                    Printer2.FontSize = iGrid.CellFontSize * iFontsProportion
                                    Do Until Printer2.FontSize <= (iGrid.CellFontSize * iSng * 1.001)
                                        iFontsProportion = iFontsProportion * 0.995
                                        Printer2.FontSize = iGrid.CellFontSize * iFontsProportion
                                        If Printer2.FontSize < 2 Then Exit Do
                                    Loop
                                End If
                                Printer2.FontBold = iGrid.CellFontBold
                                Printer2.FontItalic = iGrid.CellFontItalic
                                On Error Resume Next
                                Printer2.FontStrikeThru = iGrid.CellFontStrikeThrough
                                Printer2.FontUnderLine = iGrid.CellFontUnderline
                                PrinterEx.FontWidth = iGrid.CellFontWidth
                                On Error GoTo 0
                            
                                iSng = Printer2.TextWidth(WordWrappedInMax3Lines(FirstWords(iGrid.Text, 15))) + iCellMarginsX * 2 + 15
                                If iSng > iAuxColumnTextWidths(c) Then
                                    iAuxColumnTextWidths(c) = iSng
                                End If
                            End If
                        Next r
                        iColumnHeadersTextWidths(c) = iAuxColumnTextWidths(c)
                    End If
                    
                    If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                        iTotalTextWidth = iTotalTextWidth + iColumnHeadersTextWidths(c)
                    Else
                        iTotalTextWidth = iTotalTextWidth + iColumnDataTextWidths(c)
                    End If
                End If
            Next c
        End If
        
        If iWordWrapTheColumnHeaders Then
            If iTotalTextWidth <= iScaleWidth Then
                iDone = True
                For c = 0 To iGridCols - 1
                    If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                        mGridPrintingData.ColWidth(c) = iColumnHeadersTextWidths(c) / iTotalTextWidth * iScaleWidth
                    Else
                        mGridPrintingData.ColWidth(c) = iColumnDataTextWidths(c) / iTotalTextWidth * iScaleWidth
                    End If
                Next c
                mGridPrintingData.SetRowsHeightsToProportion iFontsProportion
            End If
        End If
        
        If Not iDone Then
            ' try reducing columns data font
            iTotalTextWidth = 0
            For c = 0 To iGridCols - 1
'                        If c = 2 Then Stop
                If mGridPrintingData.ColIsVisible(c) Then
                    If iColumnDataTextWidths(c) > iColumnHeadersTextWidths(c) Then
                        iReduceTheColumnsDataFont = True And (Not iDoNotReduceFontSize)
                        iColumnDataTextWidths(c) = 0
                        For r = 1 To 5
                            If mGridPrintingData.RowIsVisible(r) Then
                                If iLongerTexts(c, r) <> 0 Then
                                    iGrid.col = c
                                    iGrid.Row = iLongerTextsRow(c, r)
                                    Printer2.FontName = iGrid.CellFontName
                                    Printer2.FontSize = iGrid.CellFontSize * iFontsProportion * 0.75
                                    Printer2.FontBold = iGrid.CellFontBold
                                    Printer2.FontItalic = iGrid.CellFontItalic
                                    On Error Resume Next
                                    Printer2.FontStrikeThru = iGrid.CellFontStrikeThrough
                                    Printer2.FontUnderLine = iGrid.CellFontUnderline
                                    PrinterEx.FontWidth = iGrid.CellFontWidth
                                    On Error GoTo 0
                                    
                                    iSng = Printer2.TextWidth(iGrid.Text) + iCellMarginsX * 2 + 15
                                    If iSng > iColumnDataTextWidths(c) Then
                                        iColumnDataTextWidths(c) = iSng
                                    End If
                                End If
                            End If
                        Next r
                    End If
                    
                    If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                        iTotalTextWidth = iTotalTextWidth + iColumnHeadersTextWidths(c)
                        iColumnTextWidthsGeneralFontReduced(c) = iColumnHeadersTextWidths(c)
                    Else
                        iTotalTextWidth = iTotalTextWidth + iColumnDataTextWidths(c)
                        iColumnTextWidthsGeneralFontReduced(c) = iColumnDataTextWidths(c)
                    End If
                End If
            Next c
        
            If iReduceTheColumnsDataFont Then
                If iTotalTextWidth <= iScaleWidth Then
                    iDone = True
                    For c = 0 To iGridCols - 1
                        If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                            mGridPrintingData.ColWidth(c) = iColumnHeadersTextWidths(c) / iTotalTextWidth * iScaleWidth
                        Else
                            mGridPrintingData.ColWidth(c) = iColumnDataTextWidths(c) / iTotalTextWidth * iScaleWidth
                        End If
                    Next c
                    mGridPrintingData.SetRowsHeightsToProportion iFontsProportion
                End If
            End If
        End If
        
        If Not iDone Then
            ' try with columns data wordwrap
            ReDim iAuxColumnTextWidths(iGridCols - 1)
            iTotalTextWidth = 0
            For c = 0 To iGridCols - 1
                If mGridPrintingData.ColIsVisible(c) Then
                    If iColumnDataTextWidths(c) > iColumnHeadersTextWidths(c) Then
                        For r = 1 To 5
                            If mGridPrintingData.RowIsVisible(r) Then
                                If iLongerTexts(c, r) <> 0 Then
                                    If InStr(iGrid.TextMatrix(iLongerTextsRow(c, r), c), " ") > 0 Then
                                        iWordWrapTheColumnData = True
                                        iColumnDataWordWraps(c) = True
                                    End If
                                    iGrid.col = c
                                    iGrid.Row = iLongerTextsRow(c, r)
                                    Printer2.FontName = iGrid.CellFontName
                                    If iReduceTheColumnsDataFont Then
                                        Printer2.FontSize = iGrid.CellFontSize * iFontsProportion * 0.75
                                    Else
                                        iSng = iFontsProportion
                                        Printer2.FontSize = iGrid.CellFontSize * iFontsProportion
                                        Do Until Printer2.FontSize <= (iGrid.CellFontSize * iSng * 1.001)
                                            iFontsProportion = iFontsProportion * 0.995
                                            Printer2.FontSize = iGrid.CellFontSize * iFontsProportion
                                            If Printer2.FontSize < 2 Then Exit Do
                                        Loop
                                    End If
                                    Printer2.FontBold = iGrid.CellFontBold
                                    Printer2.FontItalic = iGrid.CellFontItalic
                                    On Error Resume Next
                                    Printer2.FontStrikeThru = iGrid.CellFontStrikeThrough
                                    Printer2.FontUnderLine = iGrid.CellFontUnderline
                                    PrinterEx.FontWidth = iGrid.CellFontWidth
                                    On Error GoTo 0
                                    
                                    iSng = Printer2.TextWidth(WordWrappedInMax3Lines(FirstWords(iGrid.Text, 15))) + iCellMarginsX * 2 + 15
                                    If iSng > iAuxColumnTextWidths(c) Then
                                        iAuxColumnTextWidths(c) = iSng
                                    End If
                                End If
                            End If
                        Next r
                        iColumnDataTextWidths(c) = iAuxColumnTextWidths(c)
                    End If
                    
                    If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                        iTotalTextWidth = iTotalTextWidth + iColumnHeadersTextWidths(c)
                    Else
                        iTotalTextWidth = iTotalTextWidth + iColumnDataTextWidths(c)
                    End If
                End If
            Next c
            
            If iWordWrapTheColumnData Then
                If iTotalTextWidth <= iScaleWidth Then
                    iDone = True
                    For c = 0 To iGridCols - 1
                        If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                            mGridPrintingData.ColWidth(c) = iColumnHeadersTextWidths(c) / iTotalTextWidth * iScaleWidth
                        Else
                            mGridPrintingData.ColWidth(c) = iColumnDataTextWidths(c) / iTotalTextWidth * iScaleWidth
                        End If
                    Next c
                    mGridPrintingData.SetRowsHeightsToProportion iFontsProportion
                End If
            End If
        End If
        
        If Not iDone Then
            If Not (iTotalTextWidth <= iScaleWidth) Then
                iWordBreakTheColumnTexts = True
            End If
            
            iDone = True
        End If
        
        If iWordBreakTheColumnTexts Then
            
            iAuxLng = 0
            For c = 0 To iGridCols - 1
                If Not (iReduceTheColumnsHeadersFont And iReduceTheColumnsDataFont) Then
                    iAuxLng = iAuxLng + mGridPrintingData.MinColWidth(c)
                Else
                    iAuxLng = iAuxLng + mGridPrintingData.MinColWidthFontReduced(c)
                End If
            Next c
            For c = 0 To iGridCols - 1
                If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                    mGridPrintingData.ColWidth(c) = iColumnHeadersTextWidths(c) / iTotalTextWidth * (iScaleWidth - iAuxLng)
                Else
                    mGridPrintingData.ColWidth(c) = iColumnDataTextWidths(c) / iTotalTextWidth * (iScaleWidth - iAuxLng)
                End If
                If Not (iReduceTheColumnsHeadersFont And iReduceTheColumnsDataFont) Then
                    mGridPrintingData.ColWidth(c) = mGridPrintingData.ColWidth(c) + mGridPrintingData.MinColWidth(c)
                Else
                    mGridPrintingData.ColWidth(c) = mGridPrintingData.ColWidth(c) + mGridPrintingData.MinColWidthFontReduced(c)
                End If
            Next c
            
            ReDim iAuxWidthToApply(iGridCols - 1)
            ' esto que sigue es para que no corte los textos de las columnas si es por poco
            ReDim iAuxColumnsToWordBreak(iGridCols - 1)
            For c = 0 To iGridCols - 1
                If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                    iAuxColumnsToWordBreak(c) = Abs(iColumnHeadersTextWidths(c) - mGridPrintingData.ColWidth(c)) > 80
                    If Not iAuxColumnsToWordBreak(c) Then
                        iAuxWidthToApply(c) = iColumnHeadersTextWidths(c)
                    Else
                        iAuxColumnsToWordBreak(c) = Abs(iColumnDataTextWidths(c) - mGridPrintingData.ColWidth(c)) > 80
                        If Not iAuxColumnsToWordBreak(c) Then
                            iAuxWidthToApply(c) = iColumnDataTextWidths(c)
                        End If
                    End If
                Else
                    iAuxColumnsToWordBreak(c) = Abs(iColumnDataTextWidths(c) - mGridPrintingData.ColWidth(c)) > 80
                    If Not iAuxColumnsToWordBreak(c) Then
                        iAuxWidthToApply(c) = iColumnDataTextWidths(c)
                    End If
                End If
            Next c
        
            iAuxLng = 0
            iAuxLng2 = 0
            For c = 0 To iGridCols - 1
                If iAuxColumnsToWordBreak(c) Then
                    If Not (iReduceTheColumnsHeadersFont And iReduceTheColumnsDataFont) Then
                        iAuxLng = iAuxLng + mGridPrintingData.MinColWidth(c)
                    Else
                        iAuxLng = iAuxLng + mGridPrintingData.MinColWidthFontReduced(c)
                    End If
                    If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                        iAuxLng2 = iAuxLng2 + iColumnHeadersTextWidths(c)
                    Else
                        iAuxLng2 = iAuxLng2 + iColumnDataTextWidths(c)
                    End If
                Else
                    iAuxLng = iAuxLng + iAuxWidthToApply(c)
                End If
            Next c
            If (iScaleWidth - iAuxLng) > 1000 Then
                For c = 0 To iGridCols - 1
                    If iAuxColumnsToWordBreak(c) Then
                        If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                            mGridPrintingData.ColWidth(c) = iColumnHeadersTextWidths(c) / iAuxLng2 * (iScaleWidth - iAuxLng)
                        Else
                            mGridPrintingData.ColWidth(c) = iColumnDataTextWidths(c) / iAuxLng2 * (iScaleWidth - iAuxLng)
                        End If
                        If Not (iReduceTheColumnsHeadersFont And iReduceTheColumnsDataFont) Then
                            mGridPrintingData.ColWidth(c) = mGridPrintingData.ColWidth(c) + mGridPrintingData.MinColWidth(c)
                        Else
                            mGridPrintingData.ColWidth(c) = mGridPrintingData.ColWidth(c) + mGridPrintingData.MinColWidthFontReduced(c)
                        End If
                    Else
                        mGridPrintingData.ColWidth(c) = iAuxWidthToApply(c)
                    End If
                Next c
            End If
        Else
            ReDim iColumnIsFixed(iGridCols - 1)
            iFixedWidth = 0
            iTotalTextWidth = 0
            For c = 0 To iGridCols - 1
                If iWordWrapTheColumnData Then
                    If Not iColumnDataWordWraps(c) Then
                        If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                            iFixedWidth = iFixedWidth + iColumnHeadersTextWidths(c)
                            mGridPrintingData.ColWidth(c) = iColumnHeadersTextWidths(c)
                        Else
                            iFixedWidth = iFixedWidth + iColumnDataTextWidths(c)
                            mGridPrintingData.ColWidth(c) = iColumnDataTextWidths(c)
                        End If
                        iColumnIsFixed(c) = True
                    End If
                End If
                If Not iColumnIsFixed(c) Then
                    If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                        iTotalTextWidth = iTotalTextWidth + iColumnHeadersTextWidths(c)
                    Else
                        iTotalTextWidth = iTotalTextWidth + iColumnDataTextWidths(c)
                    End If
                End If
            Next c
            
            For c = 0 To iGridCols - 1
                If mGridPrintingData.ColIsVisible(c) Then
                    If Not iColumnIsFixed(c) Then
                        iAuxLng = 0
                        If iReduceTheColumnsDataFont Then
                            If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                                If (iColumnHeadersTextWidths(c) / iTotalTextWidth * (iScaleWidth - iFixedWidth)) > iColumnTextWidthsGeneralFontReduced(c) Then
                                    mGridPrintingData.ColWidth(c) = iColumnTextWidthsGeneralFontReduced(c)
                                    iAuxLng = iColumnTextWidthsGeneralFontReduced(c)
                                    iFixedWidth = iFixedWidth + iAuxLng
                                    iTotalTextWidth = iTotalTextWidth - iColumnHeadersTextWidths(c)
                                    iColumnIsFixed(c) = True
                                End If
                            Else
                                If (iColumnDataTextWidths(c) / iTotalTextWidth * (iScaleWidth - iFixedWidth)) > iColumnTextWidthsGeneralFontReduced(c) Then
                                    mGridPrintingData.ColWidth(c) = iColumnTextWidthsGeneralFontReduced(c)
                                    iAuxLng = iColumnTextWidthsGeneralFontReduced(c)
                                    iFixedWidth = iFixedWidth + iAuxLng
                                    iTotalTextWidth = iTotalTextWidth - iColumnDataTextWidths(c)
                                    iColumnIsFixed(c) = True
                                End If
                            End If
                        Else
                            If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                                If (iColumnHeadersTextWidths(c) / iTotalTextWidth * (iScaleWidth - iFixedWidth)) > iColumnTextWidthsGeneralFontNormal(c) Then
                                    mGridPrintingData.ColWidth(c) = iColumnTextWidthsGeneralFontNormal(c)
                                    iAuxLng = iColumnTextWidthsGeneralFontNormal(c)
                                    iFixedWidth = iFixedWidth + iAuxLng
                                    iTotalTextWidth = iTotalTextWidth - iColumnHeadersTextWidths(c)
                                    iColumnIsFixed(c) = True
                                End If
                            Else
                                If (iColumnDataTextWidths(c) / iTotalTextWidth * (iScaleWidth - iFixedWidth)) > iColumnTextWidthsGeneralFontNormal(c) Then
                                    mGridPrintingData.ColWidth(c) = iColumnTextWidthsGeneralFontNormal(c)
                                    iAuxLng = iColumnTextWidthsGeneralFontNormal(c)
                                    iTotalTextWidth = iTotalTextWidth - iColumnDataTextWidths(c)
                                    iFixedWidth = iFixedWidth + iAuxLng
                                    iColumnIsFixed(c) = True
                                End If
                            End If
                        End If
                        
                    End If
                End If
            Next c
            
            For c = 0 To iGridCols - 1
                If mGridPrintingData.ColIsVisible(c) Then
                    If Not iColumnIsFixed(c) Then
                        If iColumnHeadersTextWidths(c) > iColumnDataTextWidths(c) Then
                            mGridPrintingData.ColWidth(c) = iColumnHeadersTextWidths(c) / iTotalTextWidth * (iScaleWidth - iFixedWidth)
                        Else
                            mGridPrintingData.ColWidth(c) = iColumnDataTextWidths(c) / iTotalTextWidth * (iScaleWidth - iFixedWidth)
                        End If
                    End If
                End If
            Next c
            
            iAuxLng = 0
            For c = 0 To iGridCols - 1
                iAuxLng = iAuxLng + mGridPrintingData.ColWidth(c)
            Next c
            If iAuxLng < iScaleWidth Then
                iAuxLng2 = iScaleWidth - iAuxLng
                iAuxLng3 = iAuxLng2 / mGridPrintingData.VisibleCols
            End If
            If iAuxLng3 > 0 Then
                For c = 0 To iGridCols - 1
                    If mGridPrintingData.ColIsVisible(c) Then
                        mGridPrintingData.ColWidth(c) = mGridPrintingData.ColWidth(c) + iAuxLng3
                    End If
                Next c
            End If
        End If
        
        mGridPrintingData.SetRowsHeightsToProportion iFontsProportion
        
        If iReduceTheColumnsHeadersFont Then
            ' check if it's really necessary at the end whether to reduce the columns headers font or not
            iReduceTheColumnsHeadersFont = False
            For c = 0 To iGridCols - 1
                If mGridPrintingData.ColIsVisible(c) Then
                    If iColumnHeadersTextWidthsWithoutReducingTheFont(c) > mGridPrintingData.ColWidth(c) Then
                        iReduceTheColumnsHeadersFont = True
                        Exit For
                    End If
                End If
            Next c
        End If
        
    End If
        'End If
    'End If
    mGridPrintingData.CalcColPos
    

    ' begin the printing process
    Set iAuxCurFont = Printer2.Font
    Set Printer2.Font = iDefAuxFont
    RaiseEvent StartDocument(iGrid.Name)
    RaiseEvent StartPage(iGrid.Name)
    Set Printer2.Font = iAuxCurFont
    mPrintFnObject.DocumentName = FileName2
    If mHeading <> "" Then
        Set Printer2.Font = mPrintGridFormatSettings.HeadingFont
'        Printer2.FontSize = Printer2.FontSize * mPrintGridFormatSettings.ScalePercent / 100
        Printer2.ForeColor = mPrintGridFormatSettings.HeadingFontColor
        PrinterEx.PrintEx mHeading, , , , , , flexAlignCenterCenter
        Printer2.Print
    End If
    If mSubheading <> "" Then
        Set Printer2.Font = mPrintGridFormatSettings.SubheadingFont
 '       Printer2.FontSize = Printer2.FontSize * mPrintGridFormatSettings.ScalePercent / 100
        Printer2.ForeColor = mPrintGridFormatSettings.SubheadingFontColor
        PrinterEx.PrintEx mSubheading
        Printer2.Print
    End If
    If mMiddleText <> "" Then
        Set Printer2.Font = mPrintGridFormatSettings.OtherTextsFont
  '      Printer2.FontSize = Printer2.FontSize * mPrintGridFormatSettings.ScalePercent / 100
        Printer2.ForeColor = mPrintGridFormatSettings.OtherTextsFontColor
        iLines = Split(mMiddleText, vbCrLf)
        For c = 0 To UBound(iLines)
            PrinterEx.PrintEx iLines(c)
        Next c
        Printer2.Print
    End If
    
    Set iAuxCurFont = Printer2.Font
    Set Printer2.Font = iDefAuxFont
    RaiseEvent BeforePrintGrid(iGrid.Name)
    Set Printer2.Font = iAuxCurFont
    
    'now we print the Grid
    Dim iMergeCols As Boolean
    Dim iMergeRows As Boolean
    Dim iMergeCol As Boolean
    Dim iMergeRow As Boolean
    Dim iRowHeight As Long
    Dim iPosCellLineLeft As Long
    Dim iPosCellLineRight As Long
    Dim iCellWidth As Long
    Dim iText As String
    Dim iAuxFontIndex As Long
    Dim iLastFontIndex As Long
    Dim iLastFontWasReduced As Boolean
    Dim iAuxTop As Long
    Dim c2 As Long
    Dim ci As Long
    Dim r2 As Long
    Dim iRow As Long
    Dim h As Long
    Dim iRh As Long
    Dim iAuxReduceFont As Boolean
    Dim iGridLeftPos As Long
    Dim iWW As Boolean
    Dim iCY As Long
    Dim iOuterBorderLeft As Long
    Dim iOuterBorderRight As Long
    Dim iOuterBorderTop As Long
    Dim iOuterBorderBottom As Long
    Dim iHeadersHeight As Long
    Dim iPrintCellBackColor As Boolean
    Dim iAuxColor As Long
    Dim iItsAFixedRow As Boolean
    Dim iThereAreVisibleFixedRows As Boolean
    Dim iGridReportStyle As GridReportStyle
    Dim iRh2 As Long
    Dim iAuxColIsMerged As Boolean
    Dim iRowInPage As Long
    Dim iAuxRowHeight As Long
    Dim iAuxMergedHeight As Long
'    Dim iAuxNewRowHeightForRows As Long
    Dim iLastMergedRow As Long
    Dim iThereWereColsBeingMerged As Boolean
    Dim iMergeUpToRow As Long
    Dim iCY2 As Long
    Dim iAuxColEndMerge As Boolean
    Dim iAuxStartsNewPage As Boolean
    Dim iFactor As Single
    Dim iColsAreMerged As Boolean
    Dim iAlign As Long
    Dim iCheckRowTextHeight As Boolean
    Dim iOldFontSize As Single
    Dim c3 As Long
    Dim iTextExceedsWidth As Boolean
    
    iColsAreMerged = False
    If ((mCurrentGrid.MergeCells = flexMergeFree) Or (mCurrentGrid.MergeCells = flexMergeRestrictColumns)) Then
        If mCurrentGrid.Cols > 1 Then
            If mCurrentGrid.MergeCol(1) Then
                iColsAreMerged = True
            End If
        End If
    End If
    
    Set iGridReportStyle = mGridReportStyle.Clone
    If iColsAreMerged Then
        If iGridReportStyle.PrintColumnsDataLines Then
            iGridReportStyle.PrintRowsLines = True
            iGridReportStyle.RowsLinesColor = iGridReportStyle.ColumnsDataLinesColor
        End If
    Else
        If iWordWrapTheColumnData Then
            If Not iGridReportStyle.PrintRowsLines Then
                If mUsingDefaultReportStyle Then
                    iGridReportStyle.PrintRowsLines = True
                    If iGridReportStyle.RowsLinesColor <> vbWhite Then
                        iGridReportStyle.RowsLinesColor = RGB(227, 227, 227)
                    End If
                End If
            End If
        End If
    End If
    
    If iConsiderGridWordWrap Then
        If iGrid.WordWrap Then
            iWordWrapTheColumnHeaders = True
            iWordWrapTheColumnData = True
        End If
    End If
    
    iThereAreVisibleFixedRows = mGridPrintingData.VisibleFixedRows <> 0
    Select Case iGridAlign
        Case 0, 2 ' stretch, Left$
            iGridLeftPos = 0
        Case 1 ' center
            iGridLeftPos = (iScaleWidth - mGridPrintingData.GridWidth) / 2
        Case 3 ' Right$
            iGridLeftPos = iScaleWidth - mGridPrintingData.GridWidth
    End Select
    
    iMergeCols = (iGrid.MergeCells = flexMergeFree) Or (iGrid.MergeCells = flexMergeRestrictColumns)
    iMergeRows = (iGrid.MergeCells = flexMergeFree) Or (iGrid.MergeCells = flexMergeRestrictRows)
    
    ' prepare object with preinting data
    iAuxTop = 0
    For r = 0 To iLastRowToPrint
        If mGridPrintingData.RowIsVisible(r) Then
            iMergeRow = iMergeRows And iGrid.MergeRow(r)
    
            ' calculate the height of the row
            iPosCellLineLeft = 0
            For c = 0 To iGridCols - 1
'                If mGridPrintingData.ColIsVisible(c) Then
                    ci = c
                    iText = iGrid.TextMatrix(r, c)
                    iLen = Len(iText)
                    iMergeCol = iMergeCols And iGrid.MergeCol(c)
                    iPosCellLineLeft = mGridPrintingData.ColPosLeft(c)
                    iAuxFontIndex = mGridPrintingData.CellFontAttibutesIndex(r, c)
                    If r < iFixedRows Then
                        iAuxReduceFont = iReduceTheColumnsHeadersFont
                    Else
                        iAuxReduceFont = iReduceTheColumnsDataFont
                    End If
                    If (iAuxFontIndex <> iLastFontIndex) Or (iAuxReduceFont <> iLastFontWasReduced) Then
                        iLastFontIndex = iAuxFontIndex
                        Set iFontAttibutes = mGridPrintingData.CellFontAttibutes(r, c)
                        Printer2.FontName = iFontAttibutes.Name
                        If iAuxReduceFont Then
                            Printer2.FontSize = iFontAttibutes.Size * iFontsProportion * 0.75
                            iLastFontWasReduced = True
                        Else
                            Printer2.FontSize = iFontAttibutes.Size * iFontsProportion
                            iLastFontWasReduced = False
                        End If
                        Printer2.FontBold = iFontAttibutes.Bold
                        Printer2.FontItalic = iFontAttibutes.Italic
                        Printer2.FontUnderLine = iFontAttibutes.Underline
                        Printer2.FontStrikeThru = iFontAttibutes.Strikethrough
                        PrinterEx.FontWidth = iFontAttibutes.Width
                    End If
                    
                    If iMergeRow And Not (iLen = 0) Then
                        If (c = (iGridCols - 1)) Or (c = (iFixedCols - 1)) Then
                            iPosCellLineRight = mGridPrintingData.ColPosRight(c)
                        Else
                            c2 = c + 1
                            Do Until iGrid.TextMatrix(r, c2) <> iText
                                c2 = c2 + 1
                                If c2 = iGridCols Then
                                    Exit Do
                                End If
                            Loop
                            c2 = c2 - 1
                            iPosCellLineRight = mGridPrintingData.ColPosRight(c2)
                            If c2 <> c Then
                                mGridPrintingData.LastColMergedInThisCellAtThisRow(r, c) = c2
                                c = c2
                            End If
                        End If
                    Else
                        iPosCellLineRight = mGridPrintingData.ColPosRight(c)
                    End If
                    iIgnore = False
                    iMergeUpToRow = -1
                    If iMergeCol And Not (iLen = 0) Then
                        If mGridPrintingData.RowWhereCellIsMerged(r, c) <> -1 Then
                            iIgnore = True
                        Else
                            If (r <> (iGridRows - 1)) And (r <> (iFixedRows - 1)) Then
                                r2 = r + 1
                                Do Until iGrid.TextMatrix(r2, c) <> iText
                                    r2 = r2 + 1
                                    If r2 = iGridRows Then
                                        Exit Do
                                    End If
                                Loop
                                r2 = r2 - 1
                                If r2 <> r Then
                                    iMergeUpToRow = r2
                                End If
                            End If
                        End If
                    End If
                    
                    If Not iIgnore Then
                        iCellWidth = iPosCellLineRight - iPosCellLineLeft
                        If iLen <> 0 Then
                            If r < iFixedRows Then
                                iWW = iWordWrapTheColumnHeaders Or iWordBreakTheColumnTexts
                            Else
                                iWW = iWordWrapTheColumnData Or iWordBreakTheColumnTexts
                            End If
                            iAuxLng = PrinterEx.TextHeightEx(iText, vxNoLineFeed, iPosCellLineLeft, , iCellWidth, mGridPrintingData.CellAlignment(r, c), iWW, iWordBreakTheColumnTexts, iCellMarginsX, iCellMarginsY)
                            If mGridPrintingData.LastRowMergedInThisCellAtThisCol(r, c) = 0 Then
                                If iMergeUpToRow = -1 Then
                                    If iAuxLng > mGridPrintingData.TextHeightRow(r) Then
                                        mGridPrintingData.TextHeightRow(r) = iAuxLng
                                    End If
                                End If
                            Else
                                mGridPrintingData.TextHeightCell(r, c) = iAuxLng
                            End If
                        End If
                        If iCellWidth <> 0 Then
                            mGridPrintingData.CellText(r, ci) = iText
                            If iMergeUpToRow = -1 Then
                                mGridPrintingData.PrintCell(r, ci) = True
                            End If
                        End If
                        mGridPrintingData.CellWidth(r, ci) = iCellWidth
                        mGridPrintingData.CellPrintRightLine(r, ci) = c <> (mGridPrintingData.LastVisibleCol)
                        If iMergeUpToRow <> -1 Then
                            mGridPrintingData.LastRowMergedInThisCellAtThisCol(r, c) = iMergeUpToRow
                        End If
                    End If
'                End If
            Next c
        End If
    Next r
    
    ' print the Grid
    ' print the first headers
    iRowInPage = 0
    Printer2.DrawWidth = iGridReportStyle.LineWidth
    iOuterBorderLeft = iGridLeftPos
    iOuterBorderTop = Printer2.CurrentY
    iOuterBorderRight = iOuterBorderLeft + mGridPrintingData.GridWidth
    
    iHeadersHeight = 0
    For h = 0 To iFixedRows - 1
        iRh = mGridPrintingData.GetRowHeightWithText(h)
        iHeadersHeight = iHeadersHeight + iRh
    Next h
    If iFixedRows > iGrid.Rows - 1 Then
        iRh = mGridPrintingData.GetRowHeightWithText(iFixedRows)
    End If
    If (Printer2.CurrentY + iHeadersHeight + iRh) > Printer2.ScaleHeight Then
        
        Printer2.CurrentX = 0
'        Printer2.CurrentY = Printer2.ScaleHeight - 1000
        
        Set iAuxCurFont = Printer2.Font
        Set Printer2.Font = iDefAuxFont
        PrinterEx.HandleMargins = False
        RaiseEvent EndPage(iGrid.Name)
        PrinterEx.HandleMargins = True
        Set Printer2.Font = iAuxCurFont
        
        Printer2.NewPage
        
        Set iAuxCurFont = Printer2.Font
        Set Printer2.Font = iDefAuxFont
        RaiseEvent StartPage(iGrid.Name)
        Set Printer2.Font = iAuxCurFont
        
        iOuterBorderTop = Printer2.CurrentY
    End If
    For h = 0 To iFixedRows - 1
        iRow = h
        GoSub PrintRow:
    Next h
    If iGridReportStyle.PrintHeadersBorder And iThereAreVisibleFixedRows Then
        iOuterBorderTop = iOuterBorderTop + iHeadersHeight
    End If
    If iGridReportStyle.PrintHeadersBorder And iThereAreVisibleFixedRows Then
        iCY = Printer2.CurrentY
        Printer2.Line (iOuterBorderLeft, iOuterBorderTop - iHeadersHeight)-(iOuterBorderRight, iOuterBorderTop - iHeadersHeight), iGridReportStyle.HeadersBorderColor
        Printer2.Line (iOuterBorderLeft, iOuterBorderTop - iHeadersHeight)-(iOuterBorderLeft, iOuterBorderTop - iHeadersHeight + iHeadersHeight), iGridReportStyle.HeadersBorderColor
        Printer2.Line (iOuterBorderRight, iOuterBorderTop - iHeadersHeight)-(iOuterBorderRight, iOuterBorderTop - iHeadersHeight + iHeadersHeight), iGridReportStyle.HeadersBorderColor
        Printer2.CurrentY = iCY
    End If
    For r = iFixedRows To iLastRowToPrint
        If mGridPrintingData.RowIsVisible(r) Then
            'aaaaaaaaaaaaaaaa
            If mGridPrintingData.mfrmSettingGridDataProgressShown Then
                If r Mod 350 = 0 Then
                    frmSettingGridDataProgress.CurrentPage = Printer2.Page
                    frmSettingGridDataProgress.ZOrder
                    DoEvents
                    If frmSettingGridDataProgress.Canceled Then
                        Unload frmSettingGridDataProgress
                        DoEvents
                        Set frmSettingGridDataProgress = Nothing
                        mGridPrintingData.mfrmSettingGridDataProgressShown = False
                        nCanceled = True
                        iGrid.Redraw = iGridReDraw
                        Exit Sub
                    End If
                    frmSettingGridDataProgress.pgb1.Value = frmSettingGridDataProgress.pgb1.Max / 3 + (r * iGridCols / 1000 * 2)
                End If
            End If
            iRowInPage = iRowInPage + 1
            
            If iMergeCols Then
                ' the row final height can change due to different fonts in the cells
                iAuxLng = 0
                For c2 = 0 To iGridCols - 1
                    If mGridPrintingData.PrintCell(r, c2) Or mGridPrintingData.RowWhereCellIsMerged(r, c2) <> -1 Then
                        
                        iAuxFontIndex = mGridPrintingData.CellFontAttibutesIndex(r, c2)
                        If r < iFixedRows Then
                            iAuxReduceFont = iReduceTheColumnsHeadersFont
                        Else
                            iAuxReduceFont = iReduceTheColumnsDataFont
                        End If
                        If (iAuxFontIndex <> iLastFontIndex) Or (iAuxReduceFont <> iLastFontWasReduced) Then
                            iLastFontIndex = iAuxFontIndex
                            Set iFontAttibutes = mGridPrintingData.CellFontAttibutes(r, c2)
                            Printer2.FontName = iFontAttibutes.Name
                            If iAuxReduceFont Then
                                Printer2.FontSize = iFontAttibutes.Size * iFontsProportion * 0.75
                                iLastFontWasReduced = True
                            Else
                                Printer2.FontSize = iFontAttibutes.Size * iFontsProportion
                                iLastFontWasReduced = False
                            End If
                            Printer2.FontBold = iFontAttibutes.Bold
                            Printer2.FontItalic = iFontAttibutes.Italic
                            Printer2.FontUnderLine = iFontAttibutes.Underline
                            Printer2.FontStrikeThru = iFontAttibutes.Strikethrough
                            PrinterEx.FontWidth = iFontAttibutes.Width
                        End If
                        
                        If r < iFixedRows Then
                            iWW = iWordWrapTheColumnHeaders Or iWordBreakTheColumnTexts
                        Else
                            iWW = iWordWrapTheColumnData Or iWordBreakTheColumnTexts
                        End If
                        iAuxLng2 = PrinterEx.TextHeightEx(mGridPrintingData.CellText(r, c2), vxNoLineFeed, mGridPrintingData.ColPosLeft(c2), , mGridPrintingData.CellWidth(r, c2), mGridPrintingData.CellAlignment(r, c2), iWW, iWordBreakTheColumnTexts, iCellMarginsX, iCellMarginsY)
                        
                        If mGridPrintingData.PrintCell(r, c2) Or mGridPrintingData.ColIsBeingMergedForPrintingInRow(c2) = r Then  ' no se si esto ultimo va (si sirve)
                            If iAuxLng2 > iAuxLng Then
                                iAuxLng = iAuxLng2
                            End If
                        End If
                        
                        ' put the text height
                        mGridPrintingData.TextHeightCell(r, c2) = iAuxLng2
                    End If
                Next c2
            '    If iAuxLng > 0 Then
            '        mGridPrintingData.TextHeightRow(r) = iAuxLng
            '    End If
                iRowHeight = mGridPrintingData.GetRowHeightWithText(r)
                If iAuxLng > iRowHeight Then
                    iRowHeight = iAuxLng
                End If
                
            Else
                iRowHeight = mGridPrintingData.GetRowHeightWithText(r)
            End If
            
            ' check new page
            If (Printer2.CurrentY + iRowHeight) > iScaleHeight Then
MakeNewPage:
                If iMergeCols Then
                    ' print pending cells that were merged
                    iThereWereColsBeingMerged = False
                    For c = 0 To iGridCols - 1
                        
                        If (mGridPrintingData.ColIsBeingMergedForPrintingInRow(c) <> -1) And (mGridPrintingData.ColMergeForPrintingBegunAtRow(c) <> r) Then
                            iThereWereColsBeingMerged = True
                            ' we need to end the merge of the col
                            mGridPrintingData.PrintCell(r - 1, c) = True
                            mGridPrintingData.CellTop(r - 1, c) = mGridPrintingData.RowFinalPosYForTop(mGridPrintingData.ColMergeForPrintingBegunAtRow(c))
                            
                            mGridPrintingData.EndMergingColForPrinting (c)
                            mGridPrintingData.StartMergingColForPrinting c, mGridPrintingData.RowWidthDataForColMergedForPrinting(c), r
                        
                            'if one merged cell has a larger font, it can take several other rows to fit its height, and if the space is not enought we need to increse the space
                            iAuxRowHeight = mGridPrintingData.TextHeightCell(r, c)
                            iAuxMergedHeight = 0
                            r2 = r
                            Do
                                If r2 > (iGridRows - 1) Then
                                    Exit Do
                                End If
                                If mGridPrintingData.CellIsColMergedInRow(r2, c) <> mGridPrintingData.RowWidthDataForColMergedForPrinting(c) Then
                                    Exit Do
                                End If
                                iAuxMergedHeight = iAuxMergedHeight + mGridPrintingData.GetRowHeightWithText(r2)
                                If iAuxMergedHeight >= iAuxRowHeight Then
                                    Exit Do
                                End If
                                r2 = r2 + 1
                            Loop
                            iLastMergedRow = r2 - 1
                            
                            If iAuxMergedHeight < iAuxRowHeight Then
                                If iAuxMergedHeight > 0 Then
                                    iFactor = iAuxRowHeight / iAuxMergedHeight
                                    For r2 = r To iLastMergedRow
                                        mGridPrintingData.TextHeightRow(r2) = mGridPrintingData.GetRowHeightWithText(r2) * iFactor
                                    Next r2
                                End If
                            End If
                        Else
                            mGridPrintingData.PrintCell(r - 1, c) = False
                        End If
                    Next c
                    If iThereWereColsBeingMerged Then
                        ' we need to print these cells
                        Printer2.CurrentY = Printer2.CurrentY - iRh ' we need to go back to the last row position
                        GoSub PrintRow:
                    End If
                End If
                
                ' make new page
                iOuterBorderBottom = Printer2.CurrentY
                If iGridReportStyle.PrintOuterBorder Then
                    If Not (iGridReportStyle.PrintHeadersBorder And iThereAreVisibleFixedRows) Then
                        Printer2.Line (iOuterBorderLeft, iOuterBorderTop)-(iOuterBorderRight, iOuterBorderTop), iGridReportStyle.OuterBorderColor
                    End If
                    Printer2.Line (iOuterBorderLeft, iOuterBorderTop)-(iOuterBorderLeft, iOuterBorderBottom), iGridReportStyle.OuterBorderColor
                    Printer2.Line (iOuterBorderRight, iOuterBorderTop)-(iOuterBorderRight, iOuterBorderBottom), iGridReportStyle.OuterBorderColor
                End If
    
                Printer2.CurrentX = 0
    '            Printer2.CurrentY = Printer2.ScaleHeight - 1000
            
                Set iAuxCurFont = Printer2.Font
                Set Printer2.Font = iDefAuxFont
                PrinterEx.HandleMargins = False
                RaiseEvent EndPage(iGrid.Name)
                PrinterEx.HandleMargins = True
                Set Printer2.Font = iAuxCurFont
                
                Printer2.NewPage
                
                Set iAuxCurFont = Printer2.Font
                Set Printer2.Font = iDefAuxFont
                RaiseEvent StartPage(iGrid.Name)
                Set Printer2.Font = iAuxCurFont
                
                iRowInPage = 0
                iOuterBorderTop = Printer2.CurrentY
                If iGridReportStyle.PrintHeadersBorder And iThereAreVisibleFixedRows Then
                    iOuterBorderTop = iOuterBorderTop + iHeadersHeight
                End If
                
                ' print headers
                If (iFontsProportion * 0.95) > 0.3 Then
'                    iRh = mGridPrintingData.GetRowHeightWithText(r)
'                    If (Printer2.CurrentY + iHeadersHeight + iRh) > Printer2.ScaleHeight Then
'                        Printer2.KillDoc
'                        mPrintGridFormatSettings.ScalePercent = 95
    '                    iFontsProportion = mPrintGridFormatSettings.ScalePercent / 100
    '                    mPrintFnObject.ScalePercent = mPrintGridFormatSettings.ScalePercent
    '                    iAuxFontIndex = 0
'                        DrawGrid nCanceled
'                        GoTo TheExit:
'                    End If
                Else
                    If (Printer2.CurrentY + iHeadersHeight + iRh) > Printer2.ScaleHeight Then
                        'Printer2.KillDoc
                        Printer2.Print "Paper too small, can't draw Grid."
                        GoTo TheExit:
                    End If
                End If
                
                For h = 0 To iFixedRows - 1
                    iRow = h
                    iRh = mGridPrintingData.GetRowHeightWithText(h)
                    GoSub PrintRow:
                Next h
                If iGridReportStyle.PrintHeadersBorder And iThereAreVisibleFixedRows Then
                    iCY = Printer2.CurrentY
                    Printer2.Line (iOuterBorderLeft, iOuterBorderTop - iHeadersHeight)-(iOuterBorderRight, iOuterBorderTop - iHeadersHeight), iGridReportStyle.HeadersBorderColor
                    Printer2.Line (iOuterBorderLeft, iOuterBorderTop - iHeadersHeight)-(iOuterBorderLeft, iOuterBorderTop - iHeadersHeight + iHeadersHeight), iGridReportStyle.HeadersBorderColor
                    Printer2.Line (iOuterBorderRight, iOuterBorderTop - iHeadersHeight)-(iOuterBorderRight, iOuterBorderTop - iHeadersHeight + iHeadersHeight), iGridReportStyle.HeadersBorderColor
                    Printer2.CurrentY = iCY
                End If
            End If
            
            If iMergeCols Then
                iAuxStartsNewPage = False
                For c = 0 To iGridCols - 1
                    If mGridPrintingData.ColIsVisible(c) Then
                        iAuxColEndMerge = False
                        If mGridPrintingData.ColIsBeingMergedForPrintingInRow(c) <> -1 Then
                            If r = iGridRows - 1 Then
                                iAuxColEndMerge = True
                            Else
                                iAuxColEndMerge = mGridPrintingData.CellIsColMergedInRow(r + 1, c) <> mGridPrintingData.RowWidthDataForColMergedForPrinting(c)
                            End If
                        End If
                        If iAuxColEndMerge Then
                            ' we need to end the merge of the col
                            mGridPrintingData.PrintCell(r, c) = True
                            mGridPrintingData.CellTop(r, c) = mGridPrintingData.RowFinalPosYForTop(mGridPrintingData.ColMergeForPrintingBegunAtRow(c))
                            mGridPrintingData.EndMergingColForPrinting (c)
                        Else
                            If mGridPrintingData.CellIsColMergedInRow(r, c) <> -1 Then
                                If mGridPrintingData.CellIsColMergedInRow(r, c) <> mGridPrintingData.RowWidthDataForColMergedForPrinting(c) Then
                                    
                                    ' start merge
                                    mGridPrintingData.StartMergingColForPrinting c, r, r
                                    'if one merged cell has a larger font, it can take several other rows to fit its height, and if the space is not enought we need to increse the space
                                    iAuxRowHeight = mGridPrintingData.TextHeightCell(r, c)
                                    
                                    If (Printer2.CurrentY + iAuxRowHeight) > iScaleHeight Then
                                        mGridPrintingData.CellText(r - 1, c) = ""
                                        iAuxStartsNewPage = True
                                    End If
                                    
                                    iAuxMergedHeight = 0
                                    r2 = r
                                    Do
                                        If r2 > (iGridRows - 1) Then
                                            Exit Do
                                        End If
                                        If mGridPrintingData.CellIsColMergedInRow(r2, c) <> mGridPrintingData.RowWidthDataForColMergedForPrinting(c) Then
                                            Exit Do
                                        End If
                                        If Not iAuxStartsNewPage Then
                                            If (Printer2.CurrentY + iAuxMergedHeight + mGridPrintingData.GetRowHeightWithText(r2)) > iScaleHeight Then
                                                Exit Do
                                            Else
                                                iAuxMergedHeight = iAuxMergedHeight + mGridPrintingData.GetRowHeightWithText(r2)
                                            End If
                                        Else
                                            iAuxMergedHeight = iAuxMergedHeight + mGridPrintingData.GetRowHeightWithText(r2)
                                        End If
                                        If iAuxMergedHeight >= iAuxRowHeight Then
                                            Exit Do
                                        End If
                                        r2 = r2 + 1
                                    Loop
                                    iLastMergedRow = r2 - 1
                                    If iAuxMergedHeight < iAuxRowHeight Then
                                        If iAuxMergedHeight > 0 Then
                                            iFactor = iAuxRowHeight / iAuxMergedHeight
                                            For r2 = r To iLastMergedRow
                                                mGridPrintingData.TextHeightRow(r2) = mGridPrintingData.GetRowHeightWithText(r2) * iFactor
                                            Next r2
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next c
                If iAuxStartsNewPage Then
                    GoTo MakeNewPage:
                End If
            End If
        
            ' print the row
            iRowHeight = mGridPrintingData.GetRowHeightWithText(r)
            iRow = r
            iRh = iRowHeight
            GoSub PrintRow:
        End If
    Next r
    If mGridPrintingData.mfrmSettingGridDataProgressShown Then
        Unload frmSettingGridDataProgress
        DoEvents
        Set frmSettingGridDataProgress = Nothing
        mGridPrintingData.mfrmSettingGridDataProgressShown = False
    End If
    
    iOuterBorderBottom = Printer2.CurrentY
    If iGridReportStyle.PrintOuterBorder Then
        Printer2.Line (iOuterBorderLeft, iOuterBorderTop)-(iOuterBorderRight, iOuterBorderTop), iGridReportStyle.OuterBorderColor
        Printer2.Line (iOuterBorderLeft, iOuterBorderTop)-(iOuterBorderLeft, iOuterBorderBottom), iGridReportStyle.OuterBorderColor
        Printer2.Line (iOuterBorderRight, iOuterBorderTop)-(iOuterBorderRight, iOuterBorderBottom), iGridReportStyle.OuterBorderColor
        Printer2.Line (iOuterBorderLeft, iOuterBorderBottom)-(iOuterBorderRight, iOuterBorderBottom), iGridReportStyle.OuterBorderColor
    End If
    
    If mFinalText <> "" Then
        Set Printer2.Font = mPrintGridFormatSettings.OtherTextsFont
'        Printer2.FontSize = Printer2.FontSize * mPrintGridFormatSettings.ScalePercent / 100
        Printer2.ForeColor = mPrintGridFormatSettings.OtherTextsFontColor
        Printer2.Print
        iLines = Split(mFinalText, vbCrLf)
        For c = 0 To UBound(iLines)
            PrinterEx.PrintEx iLines(c)
        Next c
        Printer2.Print
    End If
    
TheExit:
'    mAuxAllowAutoOrientation = False
    iGrid.Row = iGridStatus_Row
    iGrid.RowSel = iGridStatus_RowSel
    iGrid.col = iGridStatus_Col
    iGrid.ColSel = iGridStatus_RowColSel
    iGrid.Redraw = iGridReDraw
    
    Printer2.CurrentX = 0
    
    Set iAuxCurFont = Printer2.Font
    Set Printer2.Font = iDefAuxFont
    iCY = Printer2.CurrentY
    PrinterEx.HandleMargins = False
    RaiseEvent EndPage(iGrid.Name)
    PrinterEx.HandleMargins = True
    Printer2.CurrentY = iCY
    Printer2.CurrentX = 0
    Set Printer2.Font = iDefAuxFont
    RaiseEvent EndDocument(iGrid.Name)
    
Exit Sub


PrintRow:
    iConsiderGridWordWrap = False
    ' print one row
    iCY = Printer2.CurrentY
    mGridPrintingData.RowFinalPosYForTop(iRow) = iCY
    iCY2 = iCY
    
    iItsAFixedRow = iRow < iFixedRows
    If Not iThereAreVisibleFixedRows Then
        If iItsAFixedRow Then
            Return
        End If
    End If

    iCheckRowTextHeight = False
    If iRh > (Printer2.ScaleHeight - (iCY + Printer2.TwipsPerPixelY * Printer2.DrawWidth)) Then
        iRh = (Printer2.ScaleHeight - (iCY + Printer2.TwipsPerPixelY * Printer2.DrawWidth))
        iCheckRowTextHeight = True
    End If

    For c = 0 To iGridCols - 1
        If mGridPrintingData.PrintCell(iRow, c) Then
            If iMergeCols Then
                If mGridPrintingData.CellTop(iRow, c) = 0 Then
                    iRh2 = iRh
                    Printer2.CurrentY = iCY2
                    iCY = iCY2
                Else
                    iRh2 = iRh + iCY2 - mGridPrintingData.CellTop(iRow, c)
                    iCY = mGridPrintingData.CellTop(iRow, c)
                    Printer2.CurrentY = iCY
                End If
            Else
                iRh2 = iRh
            End If
            
            
            If (iGrid.MergeCells = flexMergeFree) Or (iGrid.MergeCells = flexMergeRestrictRows) Then
                If iGrid.MergeRow(iRow) = True Then
                    iAuxFontIndex = mGridPrintingData.CellFontAttibutesIndex(iRow, mGridPrintingData.FirstColMergedInThisCellAtThisRow(iRow, c))
                Else
                    iAuxFontIndex = mGridPrintingData.CellFontAttibutesIndex(iRow, c)
                End If
            Else
                iAuxFontIndex = mGridPrintingData.CellFontAttibutesIndex(iRow, c)
            End If
            
            If iItsAFixedRow Then
                iAuxReduceFont = iReduceTheColumnsHeadersFont
            Else
                iAuxReduceFont = iReduceTheColumnsDataFont
            End If
            
            If (iAuxFontIndex <> iLastFontIndex) Or (iAuxReduceFont <> iLastFontWasReduced) Then
                iLastFontIndex = iAuxFontIndex
                Set iFontAttibutes = mGridPrintingData.CellFontAttibutesByFontAttibutesIndex(iAuxFontIndex)
                Printer2.FontName = iFontAttibutes.Name
                If iAuxReduceFont Then
                    Printer2.FontSize = iFontAttibutes.Size * iFontsProportion * 0.75
                    iLastFontWasReduced = True
                Else
                    Printer2.FontSize = iFontAttibutes.Size * iFontsProportion
                    iLastFontWasReduced = False
                End If
                Printer2.FontBold = iFontAttibutes.Bold
                Printer2.FontItalic = iFontAttibutes.Italic
                Printer2.FontUnderLine = iFontAttibutes.Underline
                Printer2.FontStrikeThru = iFontAttibutes.Strikethrough
                PrinterEx.FontWidth = iFontAttibutes.Width
            End If
            
            iWW = False
            If iConsiderGridWordWrap Then
                If mGridPrintingData.GridType = efnGridTypeMSHFlex Then
                    iWW = (iGrid.ColWordWrapOption(c) And 1) = 1
                End If
            End If
            If iItsAFixedRow Then
                iWW = iWW Or iWordWrapTheColumnHeaders Or iWordBreakTheColumnTexts
                If iConsiderGridWordWrap Then
                    If mGridPrintingData.GridType = efnGridTypeMSHFlex Then
                        iWW = iWW Or (iGrid.ColWordWrapOptionFixed(c) And 1) = 1
                    End If
                End If
            Else
                iWW = iWW Or iWordWrapTheColumnData Or iWordBreakTheColumnTexts
                If iConsiderGridWordWrap Then
                    If c < iGrid.FixedCols Then
                        If mGridPrintingData.GridType = efnGridTypeMSHFlex Then
                            iWW = iWW Or (iGrid.ColWordWrapOptionFixed(c) And 1) = 1
                        End If
                    End If
                End If
            End If
            
            ' cell backcolor
            If iItsAFixedRow Then
                iPrintCellBackColor = iGridReportStyle.PrintHeadersBackground
            Else
                If (c < iFixedCols) Then
                    iPrintCellBackColor = iGridReportStyle.PrintFixedColsBackground
                Else
                    iPrintCellBackColor = iGridReportStyle.PrintOtherBackgrounds
                End If
            End If
            If iPrintCellBackColor Then
                If (iGrid.MergeCells = flexMergeFree) Or (iGrid.MergeCells = flexMergeRestrictRows) Then
                    If iGrid.MergeRow(iRow) = True Then
                        iAuxColor = mGridPrintingData.CellBackColor(iRow, mGridPrintingData.FirstColMergedInThisCellAtThisRow(iRow, c))
                    Else
                        iAuxColor = mGridPrintingData.CellBackColor(iRow, c)
                    End If
                Else
                    iAuxColor = mGridPrintingData.CellBackColor(iRow, c)
                End If
                If iAuxColor <> vbWhite Then
                    iAuxLng2 = Printer2.CurrentY
                    Printer2.Line (iGridLeftPos + mGridPrintingData.ColPosLeft(c) + Printer2.TwipsPerPixelX * Printer2.DrawWidth, iCY + Printer2.TwipsPerPixelY * Printer2.DrawWidth)-(iGridLeftPos + mGridPrintingData.ColPosLeft(c) + mGridPrintingData.CellWidth(iRow, c), iCY + iRh2), iAuxColor, BF
                    Printer2.CurrentY = iAuxLng2
                End If
            End If
            
            If Not iGridReportStyle.PrintRowsLines Then
                If iAuxColIsMerged Then
                    If (iRow <> 0) And (iRow <> mGridPrintingData.VisibleFixedRows) And (iRowInPage <> 0) Then
                        iAuxLng = iGridLeftPos + mGridPrintingData.ColPosLeft(c)
                        iAuxLng2 = Printer2.CurrentY
                        Printer2.Line (iAuxLng + Printer2.TwipsPerPixelX * Printer2.DrawWidth, iCY)-(iAuxLng + Printer2.TwipsPerPixelX * Printer2.DrawWidth + mGridPrintingData.CellWidth(iRow, c), iCY), iGridReportStyle.ColumnsDataLinesColor
                        Printer2.CurrentY = iAuxLng2
                    End If
                End If
            End If
            
            If (iGrid.MergeCells = flexMergeFree) Or (iGrid.MergeCells = flexMergeRestrictRows) Then
                If iGrid.MergeRow(iRow) = True Then
                    iAuxColor = mGridPrintingData.CellForeColor(iRow, mGridPrintingData.FirstColMergedInThisCellAtThisRow(iRow, c))
                Else
                    iAuxColor = mGridPrintingData.CellForeColor(iRow, c)
                End If
            Else
                iAuxColor = mGridPrintingData.CellForeColor(iRow, c)
            End If
            
            If iAuxColor <> Printer2.ForeColor Then
                Printer2.ForeColor = iAuxColor
            End If
            
            If (iGrid.MergeCells = flexMergeFree) Or (iGrid.MergeCells = flexMergeRestrictRows) Then
                If iGrid.MergeRow(iRow) = True Then
                    iAlign = mGridPrintingData.CellAlignment(iRow, mGridPrintingData.FirstColMergedInThisCellAtThisRow(iRow, c))
                Else
                    iAlign = mGridPrintingData.CellAlignment(iRow, c)
                End If
            Else
                iAlign = mGridPrintingData.CellAlignment(iRow, c)
            End If
            
            iOldFontSize = 0
            If iCheckRowTextHeight Then
                iAuxLng = PrinterEx.TextHeightEx(mGridPrintingData.CellText(iRow, c), vxNoLineFeed, iGridLeftPos + mGridPrintingData.ColPosLeft(c), , mGridPrintingData.CellWidth(iRow, c), mGridPrintingData.CellAlignment(iRow, c), iWW, iWordBreakTheColumnTexts, iCellMarginsX, iCellMarginsY)
                If iAuxLng > (Printer2.ScaleHeight - (iCY + Printer2.TwipsPerPixelY * Printer2.DrawWidth)) Then
                    iOldFontSize = Printer2.FontSize
                    Printer2.FontSize = Printer2.FontSize * 0.99
                    iAuxLng = PrinterEx.TextHeightEx(mGridPrintingData.CellText(iRow, c), vxNoLineFeed, iGridLeftPos + mGridPrintingData.ColPosLeft(c), , mGridPrintingData.CellWidth(iRow, c), mGridPrintingData.CellAlignment(iRow, c), iWW, iWordBreakTheColumnTexts, iCellMarginsX, iCellMarginsY)
                    Do Until iAuxLng <= (Printer2.ScaleHeight - (iCY + Printer2.TwipsPerPixelY * Printer2.DrawWidth))
                        Printer2.FontSize = Printer2.FontSize * 0.99
                        iAuxLng2 = iAuxLng
                        iAuxLng = PrinterEx.TextHeightEx(mGridPrintingData.CellText(iRow, c), vxNoLineFeed, iGridLeftPos + mGridPrintingData.ColPosLeft(c), , mGridPrintingData.CellWidth(iRow, c), mGridPrintingData.CellAlignment(iRow, c), iWW, iWordBreakTheColumnTexts, iCellMarginsX, iCellMarginsY)
                        If iAuxLng = iAuxLng2 Then
                            iAuxSng = Printer2.FontSize
                            c3 = 0
                            Do Until Printer2.FontSize <> iAuxSng
                                c3 = c3 + 1
                                If c > 90 Then Exit Do
                                Printer2.FontSize = Printer2.FontSize * (1 - c3 / 100)
                                If Printer2.FontSize < 1 Then Exit Do
                            Loop
                        End If
                        If Printer2.FontSize < 1 Then Exit Do
                    Loop
                End If
            End If
            
            iTextExceedsWidth = False ' ' adds the check above that is that no cell text exceeds the cell width (when there is no WordWrap and the text must be printed in one line). It happens just in some specific situations when the above algorytm (the one that calculates the columns widths, WordWrap's and font sizes) seems to have a bug. The correct font size to fit is assured now with this meassure (anyway perhaps one day I'll revisit the algorytm and correct it).
            PrinterEx.PrintEx mGridPrintingData.CellText(iRow, c), vxNoLineFeed, iGridLeftPos + mGridPrintingData.ColPosLeft(c), , mGridPrintingData.CellWidth(iRow, c), iRh2, iAlign, iWW, iWordBreakTheColumnTexts, iCellMarginsX, iCellMarginsY, , iTextExceedsWidth, True
            If (Not iWW) And iTextExceedsWidth Then
                iAuxLng = PrinterEx.TextWidthEx(mGridPrintingData.CellText(iRow, c), vxNoLineFeed, iGridLeftPos + mGridPrintingData.ColPosLeft(c), , iRh2, mGridPrintingData.CellAlignment(iRow, c), iWW, iWordBreakTheColumnTexts, iCellMarginsX, iCellMarginsY)
                If iAuxLng > mGridPrintingData.CellWidth(iRow, c) Then
                    If iOldFontSize = 0 Then
                        iOldFontSize = Printer2.FontSize
                    End If
                    Printer2.FontSize = Printer2.FontSize * 0.99
                    iAuxLng = PrinterEx.TextWidthEx(mGridPrintingData.CellText(iRow, c), vxNoLineFeed, iGridLeftPos + mGridPrintingData.ColPosLeft(c), , iRh2, mGridPrintingData.CellAlignment(iRow, c), iWW, iWordBreakTheColumnTexts, iCellMarginsX, iCellMarginsY)
                    Do Until iAuxLng <= mGridPrintingData.CellWidth(iRow, c)
                        Printer2.FontSize = Printer2.FontSize * 0.99
                        iAuxLng2 = iAuxLng
                        iAuxLng = PrinterEx.TextWidthEx(mGridPrintingData.CellText(iRow, c), vxNoLineFeed, iGridLeftPos + mGridPrintingData.ColPosLeft(c), , iRh2, mGridPrintingData.CellAlignment(iRow, c), iWW, iWordBreakTheColumnTexts, iCellMarginsX, iCellMarginsY)
                        If iAuxLng = iAuxLng2 Then
                            iAuxSng = Printer2.FontSize
                            c3 = 0
                            Do Until Printer2.FontSize <> iAuxSng
                                c3 = c3 + 1
                                If c > 90 Then Exit Do
                                Printer2.FontSize = Printer2.FontSize * (1 - c3 / 100)
                                If Printer2.FontSize < 1 Then Exit Do
                            Loop
                        End If
                        If Printer2.FontSize < 1 Then Exit Do
                    Loop
                End If
                PrinterEx.PrintEx mGridPrintingData.CellText(iRow, c), vxNoLineFeed, iGridLeftPos + mGridPrintingData.ColPosLeft(c), , mGridPrintingData.CellWidth(iRow, c), iRh2, iAlign, iWW, iWordBreakTheColumnTexts, iCellMarginsX, iCellMarginsY
            End If
            If iOldFontSize <> 0 Then
                Printer2.FontSize = iOldFontSize
            End If
            
            If iGridReportStyle.PrintRowsLines Then
                If iRow <> (iFixedRows - 1) Then
                    iAuxLng = iGridLeftPos + mGridPrintingData.ColPosLeft(c)
                    iAuxLng2 = Printer2.CurrentY
                    Printer2.Line (iAuxLng + Printer2.TwipsPerPixelX * Printer2.DrawWidth, iCY + iRh2)-(iAuxLng + Printer2.TwipsPerPixelX * Printer2.DrawWidth + mGridPrintingData.CellWidth(iRow, c), iCY + iRh2), iGridReportStyle.RowsLinesColor
                    Printer2.CurrentY = iAuxLng2
                End If
            Else
                ' if there are cells merged in a row, and the contiguous row is not merged or merged in another way, and the columns lines are drawn, it's necessary to draw the row line also.
                If iRow > 0 Then
                    If iRowInPage <> 0 Then
                        If mGridPrintingData.LastColMergedInThisCellAtThisRow(iRow, c) <> mGridPrintingData.LastColMergedInThisCellAtThisRow(iRow - 1, c) Then
                            If mGridPrintingData.LastColMergedInThisCellAtThisRow(iRow, c) > mGridPrintingData.LastColMergedInThisCellAtThisRow(iRow - 1, c) Then
                                iAuxLng4 = mGridPrintingData.CellWidth(iRow, c)
                            Else
                                iAuxLng4 = mGridPrintingData.CellWidth(iRow - 1, c)
                            End If
                            
                            If iRow <> (iFixedRows) Then
                                iAuxLng = iGridLeftPos + mGridPrintingData.ColPosLeft(c)
                                iAuxLng2 = Printer2.CurrentY
                                Printer2.Line (iAuxLng + Printer2.TwipsPerPixelX * Printer2.DrawWidth, iCY)-(iAuxLng + Printer2.TwipsPerPixelX * Printer2.DrawWidth + iAuxLng4, iCY), iGridReportStyle.ColumnsHeadersLinesColor
                                Printer2.CurrentY = iAuxLng2
                            End If
                        End If
                    End If
                End If
            End If
            
            If iItsAFixedRow Then
                If iGridReportStyle.PrintColumnsHeadersLines Then
                    If mGridPrintingData.CellPrintRightLine(iRow, c) Then
                        iAuxLng = iGridLeftPos + mGridPrintingData.ColPosLeft(c) + mGridPrintingData.CellWidth(iRow, c)
                        iAuxLng2 = Printer2.CurrentY
                        Printer2.Line (iAuxLng, iCY + Printer2.TwipsPerPixelY * Printer2.DrawWidth)-(iAuxLng, iCY + iRh2), iGridReportStyle.ColumnsHeadersLinesColor
                        Printer2.CurrentY = iAuxLng2
                    End If
                End If
            Else
                If iGridReportStyle.PrintColumnsDataLines Then
                    If mGridPrintingData.CellPrintRightLine(iRow, c) Then
                        iAuxLng = iGridLeftPos + mGridPrintingData.ColPosLeft(c) + mGridPrintingData.CellWidth(iRow, c)
                        iAuxLng2 = Printer2.CurrentY
                        Printer2.Line (iAuxLng, iCY + Printer2.TwipsPerPixelY * Printer2.DrawWidth)-(iAuxLng, iCY + iRh2), iGridReportStyle.ColumnsDataLinesColor
                        Printer2.CurrentY = iAuxLng2
                    End If
                End If
            End If
            If iThereAreVisibleFixedRows Then
                If iGridReportStyle.PrintHeadersSeparatorLine Then
                    If iRow = (iFixedRows - 1) Then
                        If iGridReportStyle.LineWidthHeadersSeparatorLine < 4 Then
                            Printer2.DrawWidth = iGridReportStyle.LineWidthHeadersSeparatorLine
                        Else
                            Printer2.DrawWidth = 1
                        End If
                        iAuxLng = iGridLeftPos + mGridPrintingData.ColPosLeft(c)
                        iAuxLng2 = Printer2.CurrentY
                        If iGridReportStyle.LineWidthHeadersSeparatorLine < 4 Then
                            Printer2.Line (iAuxLng, iCY + iRh2)-(iAuxLng + mGridPrintingData.CellWidth(iRow, c), iCY + iRh2), iGridReportStyle.HeadersBorderColor
                        Else
                            iAuxLng3 = (iGridReportStyle.LineWidthHeadersSeparatorLine - 1) * Printer2.TwipsPerPixelY
                            iAuxLng4 = iAuxLng3 / 2
                            Printer2.Line (iAuxLng, iCY + iRh2 - iAuxLng3 + iAuxLng4)-(iAuxLng + mGridPrintingData.CellWidth(iRow, c), iCY + iRh2 + iAuxLng4), iGridReportStyle.HeadersBorderColor, BF
                        End If
                        Printer2.CurrentY = iAuxLng2
                        Printer2.DrawWidth = iGridReportStyle.LineWidth
                    End If
                End If
            End If
        End If
    Next c
    Printer2.CurrentY = iCY2 + iRh
Return

End Sub

Public Sub EnableGridOrderByColumns(nEnabled As Boolean, Optional nGrid As Object, Optional nGridName)
    
    SavePreviousGeneralSettings
    SetCurrentGrid nGrid, nGridName
    If CurrentGrid Is Nothing Then
        MsgBox "Grid not found"
        RestorePreviousGeneralSettings
        Exit Sub
    End If
    
    mCurrentGridHandler.EnableOrderByColumns = nEnabled
    
    RestorePreviousGeneralSettings
End Sub

Public Property Let GridsFlatAppearance(Optional nGrid As Object, nValue As Boolean)
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        If nValue <> mGridsFlatAppearance Then
            mGridsFlatAppearance = nValue
            For Each iGH In mGridHandlersCollection
                iGH.GridsFlatAppearance = mGridsFlatAppearance
            Next
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.GridsFlatAppearance = nValue
        End If
    End If
End Property

Public Property Get GridsFlatAppearance(Optional nGrid As Object) As Boolean
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        GridsFlatAppearance = mGridsFlatAppearance
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
             mGridsFlatAppearance = iGH.GridsFlatAppearance
        End If
    End If
End Property


Public Property Let InitialOrderColumn(Optional nGrid As Object, nValue As Long)
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        If nValue <> mInitialOrderColumn Then
            mInitialOrderColumn = nValue
            For Each iGH In mGridHandlersCollection
                iGH.InitialOrderColumn = mInitialOrderColumn
            Next
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.InitialOrderColumn = nValue
        End If
    End If
End Property

Public Property Get InitialOrderColumn(Optional nGrid As Object) As Long
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        InitialOrderColumn = mInitialOrderColumn
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
             mInitialOrderColumn = iGH.InitialOrderColumn
        End If
    End If
End Property


Public Property Let InitialOrderDescending(Optional nGrid As Object, nValue As Boolean)
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        If nValue <> mInitialOrderDescending Then
            mInitialOrderDescending = nValue
            For Each iGH In mGridHandlersCollection
                iGH.InitialOrderDescending = mInitialOrderDescending
            Next
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.InitialOrderDescending = nValue
        End If
    End If
End Property

Public Property Get InitialOrderDescending(Optional nGrid As Object) As Boolean
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        InitialOrderDescending = mInitialOrderDescending
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
             mInitialOrderDescending = iGH.InitialOrderDescending
        End If
    End If
End Property


Public Property Let GridName(nValue As String)
    Dim iAuxGrid As Object
    Dim iGridHwnd As Long
    
    If nValue <> mGridName Then
        mGridName = nValue
        mGridExplicitelySet = mGridName <> ""
        If Not mForm Is Nothing Then
            SetCurrentGrid Nothing, mGridName
            If Not mCurrentGrid Is Nothing Then
                SetGridAsHandledIndividually
            End If
        Else
            On Error Resume Next
            Set iAuxGrid = mGridParent.Controls(mGridName)
            iGridHwnd = iAuxGrid.hWnd
            On Error GoTo 0
            If Not iAuxGrid Is Nothing Then
                If iGridHwnd <> 0 Then
                    Set CurrentGrid = iAuxGrid
                    SetGridAsHandledIndividually
                End If
            End If
        End If
    End If
End Property

Public Property Get GridName() As String
    GridName = mGridName
End Property


Public Property Set Form(nForm As Object)
    Dim iGridHwnd As Long
    
    If Not nForm Is mForm Then
        If Not mForm Is Nothing Then
            FinishForm
        End If
        Set mForm = nForm
        If Not mForm Is Nothing Then
            mGridControlsStored = False
            mFormHwnd = mForm.hWnd
            SubclassForm
            If mScrollWithMouseWheel <> 0 Then
                Set mMouseWheelNotifierObject = New MouseWheelNotifierObject
                mMouseWheelNotifierObject.SetForm mForm
            End If
            Set mTimerFormSet = New cTimer
            mTimerFormSet.Interval = 1
            If mGridName <> "" Then
                SetCurrentGrid Nothing, mGridName
                If Not mCurrentGrid Is Nothing Then
                    On Error Resume Next
                    iGridHwnd = mCurrentGrid.hWnd
                    On Error GoTo 0
                    If iGridHwnd <> 0 Then
                        SetGridAsHandledIndividually
                    End If
                End If
            End If
        End If
    End If
End Property

Public Property Get Form() As Object
    Set Form = mForm
End Property

Private Sub SetCurrentGrid(nGrid, nGridName)
    If Not nGrid Is Nothing Then
        Set CurrentGrid = nGrid
    Else
        If Not IsMissing(nGridName) Then
            If Not mForm Is Nothing Then
                On Error Resume Next
                Set CurrentGrid = mForm.Controls(nGridName)
                On Error GoTo 0
            End If
        End If
        If CurrentGrid Is Nothing Then
            If IsMissing(nGridName) And Not mGridExplicitelySet Then
                StoreGridControls
                Set CurrentGrid = GetGridControl
            End If
        End If
    End If
End Sub

Private Function GetGridControl(Optional nRequirePartiallyVisibleAtLeast As Boolean) As Object
    Dim iGH As cGridHandler
    Dim iGrid2 As Object
    Dim iActiveControl As Object
    Dim iConsider As Boolean
    Dim iActiveControlHwnd As Long
    
    On Error Resume Next
    Set iActiveControl = mForm.ActiveControl
    If Not iActiveControl Is Nothing Then
        iActiveControlHwnd = iActiveControl.hWnd
    End If
    On Error GoTo 0
    If Not iActiveControl Is Nothing Then
        If iActiveControlHwnd <> 0 Then
            If MyGetProp(iActiveControlHwnd, "HandledIndiv") = 1 Then
                Exit Function
            End If
        End If
    End If
    
    If mGridExplicitelySet Then
        Set GetGridControl = mCurrentGrid
        Exit Function
    End If
    
    If Not iActiveControl Is Nothing Then
        If IsTypeSupported(iActiveControl) Then
            If IsWindowVisibleOnScreen(iActiveControl.hWnd, nRequirePartiallyVisibleAtLeast) Then
                Set GetGridControl = iActiveControl
            End If
        End If
    End If
    
    If GetGridControl Is Nothing Then
        On Error Resume Next
        For Each iGH In mGridHandlersCollection
            If mForm Is Nothing Then
                iConsider = True
            Else
                iConsider = iGH.GridFormHwnd = mFormHwnd
            End If
            If iConsider Then
                If IsWindowVisibleOnScreen(iGH.GridHwnd, nRequirePartiallyVisibleAtLeast) Then
                    Set GetGridControl = iGH.Grid
                Else
                    If Not nRequirePartiallyVisibleAtLeast Then
                        Set iGrid2 = iGH.Grid
                    End If
                End If
            End If
        Next
        On Error GoTo 0
        If Not nRequirePartiallyVisibleAtLeast Then
            If GetGridControl Is Nothing Then
                Set GetGridControl = iGrid2
            End If
        End If
    End If
End Function

Private Function GetGridControl2(Optional nRequirePartiallyVisibleAtLeast As Boolean) As Object
    Dim iGH As cGridHandler
    Dim iGrid2 As Object
    Dim iActiveControl As Object
    Dim iConsider As Boolean
    Dim iActiveControlHwnd As Long
    Dim iAuxHwnd As Long
    
    On Error Resume Next
    Set iActiveControl = mForm.ActiveControl
    If Not iActiveControl Is Nothing Then
        iActiveControlHwnd = iActiveControl.hWnd
    End If
    On Error GoTo 0
    If Not iActiveControl Is Nothing Then
        If iActiveControlHwnd <> 0 Then
            If MyGetProp(iActiveControlHwnd, "HandledIndiv") = 1 Then
                Exit Function
            End If
        End If
    End If
    
    If mGridExplicitelySet Then
        Set GetGridControl2 = mCurrentGrid
        Exit Function
    End If
    
    iAuxHwnd = WindowUnderMouseHwnd
    If iAuxHwnd <> 0 Then
        For Each iGH In GridHandlersCollection
            If iGH.Grid.hWnd = iAuxHwnd Then
                Set GetGridControl2 = iGH.Grid
                Exit Function
            End If
        Next
    End If
    
    If Not iActiveControl Is Nothing Then
        If IsTypeSupported(iActiveControl) Then
            If IsWindowVisibleOnScreen(iActiveControl.hWnd, nRequirePartiallyVisibleAtLeast) Then
                Set GetGridControl2 = iActiveControl
            End If
        End If
    End If
    
    If GetGridControl2 Is Nothing Then
        On Error Resume Next
        For Each iGH In mGridHandlersCollection
            If mForm Is Nothing Then
                iConsider = True
            Else
                iConsider = iGH.GridFormHwnd = mFormHwnd
            End If
            If iConsider Then
                If IsWindowVisibleOnScreen(iGH.GridHwnd, nRequirePartiallyVisibleAtLeast) Then
                    Set GetGridControl2 = iGH.Grid
                Else
                    If Not nRequirePartiallyVisibleAtLeast Then
                        Set iGrid2 = iGH.Grid
                    End If
                End If
            End If
        Next
        On Error GoTo 0
        If Not nRequirePartiallyVisibleAtLeast Then
            If GetGridControl2 Is Nothing Then
                Set GetGridControl2 = iGrid2
            End If
        End If
    End If
End Function


Private Sub SavePreviousGeneralSettings()
    Set mGridPrev = CurrentGrid
    mReportIDPrev = ReportID
    mHeadingPrev = Heading
    mSubheadingPrev = Subheading
    mMiddleTextPrev = MiddleText
    mFinalTextPrev = FinalText
    mOrientationPrev = Orientation
    mScalePercentPrev = ScalePercent
    mFileNamePrev = FileName
    mDefaultFolderPathPrev = DefaultFolderPath
    mCopyToClipboardModePrev = CopyToClipboardMode
    mSpecialSeparatorCharactersPrev = SpecialSeparatorCharacters
    mFinalTextPrev = mFinalText
End Sub

Private Sub RestorePreviousGeneralSettings()
    Set CurrentGrid = mGridPrev
    ReportID = mReportIDPrev
    Heading = mHeadingPrev
    Subheading = mSubheadingPrev
    MiddleText = mMiddleTextPrev
    FinalText = mFinalTextPrev
    Orientation = mOrientationPrev
    ScalePercent = mScalePercentPrev
    FileName = mFileNamePrev
    DefaultFolderPath = mDefaultFolderPathPrev
    CopyToClipboardMode = mCopyToClipboardModePrev
    SpecialSeparatorCharacters = mSpecialSeparatorCharactersPrev
    mFinalText = mFinalTextPrev
End Sub
    

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrPostProcess
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Select Case iMsg
        Case WM_DESTROY
'            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            FinishForm
        Case WM_PARENTNOTIFY
'            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            If mGridControlsStored Then
                UpdateGridsList
            End If
        Case Else
'            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
    End Select
End Function

Private Sub SubclassForm()
    If Not mFormSubclassed Then
        AttachMessage Me, mFormHwnd, WM_PARENTNOTIFY
        AttachMessage Me, mFormHwnd, WM_DESTROY
        mFormSubclassed = True
    End If
End Sub

Private Sub UnsubclassForm()
    If Not mFormSubclassed Then Exit Sub
    DetachMessage Me, mFormHwnd, WM_PARENTNOTIFY
    DetachMessage Me, mFormHwnd, WM_DESTROY
    mFormSubclassed = False
End Sub

Private Sub mTimerFormSet_ThatTime()
    mTimerFormSet.Interval = 0
    Set mTimerFormSet = Nothing
    StoreGridControls
End Sub

Public Sub StoreGridControls()
    Dim iSuccess As Boolean
    
    If mGridControlsStored Then Exit Sub
    If mStoringGridControls Then Exit Sub
    If mForm Is Nothing Then Exit Sub
    mStoringGridControls = True
    UpdateGridsList iSuccess
    mStoringGridControls = False
    If iSuccess Then
        mGridControlsStored = True
    End If
End Sub

Private Sub UpdateGridsList(Optional nSuccess As Boolean = True)
    Dim iCtl As Control
    Dim iGridHwnd As Long
    
    If mGridName <> "" Then
        On Error Resume Next
        Set iCtl = mGridParent.Controls(mGridName)
        iGridHwnd = iCtl.hWnd
        On Error GoTo 0
        If Not iCtl Is Nothing Then
            If iGridHwnd <> 0 Then
                Set CurrentGrid = iCtl
                SetGridAsHandledIndividually
                If GetGridHandler(iCtl) Is Nothing Then
                    AddGridHandler iCtl
                End If
                Set mCurrentGridHandler = GetGridHandler(iCtl)
            Else
                nSuccess = False
            End If
        Else
            nSuccess = False
        End If
    Else
        
        On Error Resume Next
        Err.Clear
        If mLastControlsCount = mGridParent.Controls.Count Then Exit Sub
        If Err.Number <> 0 Then Exit Sub
        On Error GoTo 0
            
        For Each iCtl In mGridParent.Controls
            If IsTypeSupported(iCtl) Then
                If GetGridHandler(iCtl) Is Nothing Then
                    AddGridHandler iCtl
                End If
            End If
        Next
        mLastControlsCount = mGridParent.Controls.Count
    End If
End Sub

Private Function IsTypeSupported(nControl As Object) As Boolean
    Dim iTn As String
    
    iTn = LCase$(TypeName(nControl))
    'IsTypeSupported = (iTn = "msflexgrid") Or (iTn = "mshflexgrid")
    IsTypeSupported = InStr(iTn, "flexgrid") > 0
End Function

Private Function GetGridHandler(nGrid As Object) As cGridHandler
    StoreGridControls
    On Error Resume Next
    Set GetGridHandler = mGridHandlersCollection(CStr(nGrid.hWnd))
End Function

Private Sub AddGridHandler(nGrid As Object)
    Dim iGH As cGridHandler
    Dim iList() As Double
    Dim iList2() As Object
    Dim c As Long
    Dim iFlag As Boolean
    Dim iDbl As Double
    
    Set iGH = GetGridHandler(nGrid)
    
    If iGH Is Nothing Then
        Set iGH = New cGridHandler
        iGH.SetGrid nGrid, Me, mStretchColumnsWidthsToFill, mEnableOrderByColumns, mGridsFlatAppearance, mRememberUserPrintingPreferences, mReportID, IsTypeSupported(nGrid), mInitialOrderColumn, mInitialOrderDescending, mSameDataGroupedInColumns, mBorderColor, mBorderWidth, mShowToolTipsOnLongerCellTexts, mDoNotRememberOrder, mShowToolTipsForOrderColumns, mAllowTextEdition, mTextEditionLocked
        If iGH.Error Then Exit Sub
        mGridHandlersCollection.Add iGH, CStr(nGrid.hWnd)
         
         ' order by larger and second order by Top Left position
        ReDim iList(mGridHandlersCollection.Count)
        ReDim iList2(mGridHandlersCollection.Count)
        For c = 1 To mGridHandlersCollection.Count
            iList(c) = CDbl(mGridHandlersCollection.Item(c).Grid.Width * mGridHandlersCollection.Item(c).Grid.Height + (1000000000 - (mGridHandlersCollection.Item(c).Grid.Left * mGridHandlersCollection.Item(c).Grid.Top) * 1.4))
            Set iList2(c) = mGridHandlersCollection.Item(c)
        Next c
    
        Do
            iFlag = False
            For c = 1 To UBound(iList) - 1
                If iList(c) < iList(c + 1) Then
                    iDbl = iList(c + 1)
                    Set iGH = iList2(c + 1)
                    iList(c + 1) = iList(c)
                    Set iList2(c + 1) = iList2(c)
                    iList(c) = iDbl
                    Set iList2(c) = iGH
                    iFlag = True
                End If
            Next
        Loop While iFlag
        
        Set mGridHandlersCollection = New Collection
        For c = UBound(iList2) To 1 Step -1
            mGridHandlersCollection.Add iList2(c), CStr(iList2(c).Grid.hWnd)
        Next c
    End If
End Sub

Private Sub RemoveGridHandler(nGrid As Object)
    On Error Resume Next
    mGridHandlersCollection.Remove (CStr(nGrid.hWnd))
End Sub

Public Sub NotifyUIToBeUpdated()
Attribute NotifyUIToBeUpdated.VB_MemberFlags = "40"
    If Not TimerUpdateUI Is Nothing Then
        TimerUpdateUI.Interval = 0
    End If
    Set TimerUpdateUI = New cTimer
    TimerUpdateUI.Interval = 10
End Sub

Private Sub FinishForm()
    Dim iGH As cGridHandler
    
    If mFormHwnd <> 0 Then
        UnsubclassForm
        For Each iGH In mGridHandlersCollection
            If iGH.GridFormHwnd = mFormHwnd Then
                mGridHandlersCollection.Remove CStr(iGH.GridHwnd)
            End If
        Next
    End If
    If Not mTimerFormSet Is Nothing Then
        mTimerFormSet.Interval = 0
        Set mTimerFormSet = Nothing
    End If
    Set mMouseWheelNotifierObject = Nothing
    mFormHwnd = 0
    Set mForm = Nothing
End Sub

Private Sub TimerUpdateUI_ThatTime()
    Static sLastGrid As Object
    Dim iGrid As Object
    Dim iGridName As String
    
    TimerUpdateUI.Interval = 0
    Set TimerUpdateUI = Nothing
    If mFormHwnd <> 0 Then
        If IsWindow(mFormHwnd) <> 0 Then
            Set iGrid = GetGridControl
            If Not sLastGrid Is iGrid Then
                If Not iGrid Is Nothing Then
                    iGridName = iGrid.Name
                Else
                    If sLastGrid Is Nothing Then
                        iGridName = "-1"
                    Else
                        iGridName = ""
                    End If
                End If
                If iGridName <> "-1" Then
                    RaiseEvent WorkGridChange(iGridName)
                End If
                Set sLastGrid = iGrid
            End If
            RaiseEvent UpdateUI
        End If
    End If
End Sub

Public Function GetWorkGridControl(Optional nRequirePartiallyVisibleAtLeast As Boolean) As Object
    Set GetWorkGridControl = GetGridControl(nRequirePartiallyVisibleAtLeast)
End Function

Public Sub EventRaise(nEventName, nParameter1 As String, nParameter2 As Long, nParameter3 As Long, Optional nParameter4 As Long)
Attribute EventRaise.VB_MemberFlags = "40"
    Dim iCancel As Boolean
    Dim iAuxBool As Boolean
    
    Select Case nEventName
        Case "BeforeOrderingByColumn"
            iAuxBool = CBool(nParameter4)
            RaiseEvent BeforeOrderingByColumn(nParameter1, nParameter2, nParameter3, iAuxBool)
            nParameter4 = CLng(iAuxBool)
        Case "AfterOrderingByColumn"
            RaiseEvent AfterOrderingByColumn(nParameter1, nParameter2, nParameter3, CBool(nParameter4))
        Case "AfterSettingColumnsWidths"
            RaiseEvent AfterSettingColumnsWidths(nParameter1, CBool(nParameter2))
        Case "CellTextChange"
            RaiseEvent CellTextChange(nParameter1, nParameter2, nParameter3)
        Case "BeforeTextEdit"
            RaiseEvent BeforeTextEdit(nParameter1, iCancel)
            nParameter2 = Abs(CLng(iCancel))
    End Select
End Sub

Public Function IsShowingVerticalScrollBar(Optional nGrid As Object) As Boolean

    If nGrid Is Nothing Then
        If Not mCurrentGridHandler Is Nothing Then
            IsShowingVerticalScrollBar = mGlobals.IsShowingVerticalScrollBar(mCurrentGridHandler.Grid)
        End If
    Else
        IsShowingVerticalScrollBar = mGlobals.IsShowingVerticalScrollBar(nGrid)
    End If
    
End Function


Public Property Let SameDataGroupedInColumns(Optional nGrid As Object, nValue As Boolean)
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        If nValue <> mSameDataGroupedInColumns Then
            mSameDataGroupedInColumns = nValue
            For Each iGH In mGridHandlersCollection
                iGH.SameDataGroupedInColumns = mSameDataGroupedInColumns
            Next
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.SameDataGroupedInColumns = nValue
        End If
    End If
    
End Property

Public Property Get SameDataGroupedInColumns(Optional nGrid As Object) As Boolean
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        SameDataGroupedInColumns = mSameDataGroupedInColumns
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
             SameDataGroupedInColumns = iGH.SameDataGroupedInColumns
        End If
    End If
    
End Property


Public Property Let ShowToolTipsOnLongerCellTexts(Optional nGrid As Object, nValue As Boolean)
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        If nValue <> mShowToolTipsOnLongerCellTexts Then
            mShowToolTipsOnLongerCellTexts = nValue
            For Each iGH In mGridHandlersCollection
                iGH.ShowToolTipsOnLongerCellTexts = mShowToolTipsOnLongerCellTexts
            Next
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.ShowToolTipsOnLongerCellTexts = nValue
        End If
    End If
    
End Property

Public Property Get ShowToolTipsOnLongerCellTexts(Optional nGrid As Object) As Boolean
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        ShowToolTipsOnLongerCellTexts = mShowToolTipsOnLongerCellTexts
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
             mShowToolTipsOnLongerCellTexts = iGH.ShowToolTipsOnLongerCellTexts
        End If
    End If
    
End Property


Public Property Let AllowTextEdition(Optional nGrid As Object, nValue As Boolean)
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        If nValue <> mAllowTextEdition Then
            mAllowTextEdition = nValue
            For Each iGH In mGridHandlersCollection
                iGH.AllowTextEdition = mAllowTextEdition
            Next
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.AllowTextEdition = nValue
        End If
    End If
    
End Property

Public Property Get AllowTextEdition(Optional nGrid As Object) As Boolean
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        AllowTextEdition = mAllowTextEdition
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
             mAllowTextEdition = iGH.AllowTextEdition
        End If
    End If
    
End Property


Public Property Let TextEditionLocked(Optional nGrid As Object, nValue As Boolean)
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        If nValue <> mTextEditionLocked Then
            mTextEditionLocked = nValue
            For Each iGH In mGridHandlersCollection
                iGH.TextEditionLocked = mTextEditionLocked
            Next
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.TextEditionLocked = nValue
        End If
    End If
    
End Property

Public Property Get TextEditionLocked(Optional nGrid As Object) As Boolean
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        TextEditionLocked = mTextEditionLocked
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
             mTextEditionLocked = iGH.TextEditionLocked
        End If
    End If
    
End Property

Public Property Let DoNotRememberOrder(Optional nGrid As Object, nValue As Boolean)
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        If nValue <> mDoNotRememberOrder Then
            mDoNotRememberOrder = nValue
            For Each iGH In mGridHandlersCollection
                iGH.DoNotRememberOrder = mDoNotRememberOrder
            Next
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.DoNotRememberOrder = nValue
        End If
    End If
    
End Property

Public Property Get DoNotRememberOrder(Optional nGrid As Object) As Boolean
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        DoNotRememberOrder = mDoNotRememberOrder
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
             mDoNotRememberOrder = iGH.DoNotRememberOrder
        End If
    End If
    
End Property


Public Property Let ShowToolTipsForOrderColumns(Optional nGrid As Object, nValue As Boolean)
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        If nValue <> mShowToolTipsForOrderColumns Then
            mShowToolTipsForOrderColumns = nValue
            For Each iGH In mGridHandlersCollection
                iGH.ShowToolTipsForOrderColumns = mShowToolTipsForOrderColumns
            Next
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.ShowToolTipsForOrderColumns = nValue
        End If
    End If
    
End Property

Public Property Get ShowToolTipsForOrderColumns(Optional nGrid As Object) As Boolean
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        ShowToolTipsForOrderColumns = mShowToolTipsForOrderColumns
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
             mShowToolTipsForOrderColumns = iGH.ShowToolTipsForOrderColumns
        End If
    End If
    
End Property


Friend Property Get GridHandlersCollection() As Collection
    Set GridHandlersCollection = mGridHandlersCollection
End Property


Public Property Let EnableOrderByColumns(Optional nGrid As Object, nValue As Boolean)
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        If nValue <> mEnableOrderByColumns Then
            mEnableOrderByColumns = nValue
            For Each iGH In mGridHandlersCollection
                iGH.EnableOrderByColumns = mEnableOrderByColumns
            Next
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.EnableOrderByColumns = nValue
        End If
    End If
End Property

Public Property Get EnableOrderByColumns(Optional nGrid As Object) As Boolean
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        EnableOrderByColumns = mEnableOrderByColumns
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
             mEnableOrderByColumns = iGH.EnableOrderByColumns
        End If
    End If
End Property

Public Property Let StretchColumnsWidthsToFill(Optional nGrid As Object, nValue As Boolean)
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        If nValue <> mStretchColumnsWidthsToFill Then
            mStretchColumnsWidthsToFill = nValue
            For Each iGH In mGridHandlersCollection
                iGH.StretchColumnsWidthsToFill = mStretchColumnsWidthsToFill
            Next
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.StretchColumnsWidthsToFill = nValue
        End If
    End If
End Property

Public Property Get StretchColumnsWidthsToFill(Optional nGrid As Object) As Boolean
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        StretchColumnsWidthsToFill = mStretchColumnsWidthsToFill
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
             mStretchColumnsWidthsToFill = iGH.StretchColumnsWidthsToFill
        End If
    End If
End Property

Public Sub UpdateGridColumnsWidthsStretched(Optional nGrid As Object)
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        For Each iGH In mGridHandlersCollection
            iGH.UpdateGridColumnsWidthsStretched
        Next
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.UpdateGridColumnsWidthsStretched
        End If
    End If
End Sub

Public Property Let BorderColor(Optional nGrid As Object, nValue As Long)
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        If nValue <> mBorderColor Then
            mBorderColor = nValue
            For Each iGH In mGridHandlersCollection
                iGH.BorderColor = mBorderColor
            Next
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.BorderColor = nValue
        End If
    End If
End Property

Public Property Get BorderColor(Optional nGrid As Object) As Long
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        BorderColor = mBorderColor
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
             mBorderColor = iGH.BorderColor
        End If
    End If
End Property


Public Property Let BorderWidth(Optional nGrid As Object, nValue As Long)
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        If nValue <> mBorderWidth Then
            mBorderWidth = nValue
            For Each iGH In mGridHandlersCollection
                iGH.BorderWidth = mBorderWidth
            Next
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.BorderWidth = nValue
        End If
    End If
End Property

Public Property Get BorderWidth(Optional nGrid As Object) As Long
    Dim iGH As cGridHandler

    If nGrid Is Nothing Then
        BorderWidth = mBorderWidth
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
             mBorderWidth = iGH.BorderWidth
        End If
    End If
End Property

Public Property Set GridParent(nValue As Object)
Attribute GridParent.VB_MemberFlags = "40"
    Set mGridParent = nValue
End Property

Public Property Let AmbientUserMode(nValue As Boolean)
Attribute AmbientUserMode.VB_MemberFlags = "40"
    mAmbientUserMode = nValue
    mPrintFnObject.AmbientUserMode = nValue
End Property

Private Sub RemoveGridAsHandledIndividually()
    If mHwndGridHandledIndividually <> 0 Then
        MyRemoveProp mHwndGridHandledIndividually, "HandledIndiv"
    End If
End Sub

Private Sub SetGridAsHandledIndividually()
    RemoveGridAsHandledIndividually
    MySetProp mCurrentGrid.hWnd, "HandledIndiv", 1
    mHwndGridHandledIndividually = mCurrentGrid.hWnd
    mGridExplicitelySetIsTypeSupported = IsTypeSupported(mCurrentGrid)
End Sub

Public Property Get ReportStyle() As GridReportStyle
    Set ReportStyle = mGridReportStyle
End Property

Public Property Get PrintFnObject() As PrintFnObject
    Set PrintFnObject = mPrintFnObject
End Property

Public Function GetControls() As Object
    Dim iUCC As Object
    RaiseEvent GetUCControls(iUCC)
    Set GetControls = iUCC
End Function

Public Sub OrderGridByColumn(Optional nGrid As Object, Optional nColumn As Long = -1, Optional nDescending)
    Dim iGH As cGridHandler
    
    If nGrid Is Nothing Then
        If Not mCurrentGridHandler Is Nothing Then
            mCurrentGridHandler.OrderGridByColumn nColumn, nDescending
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.OrderGridByColumn nColumn, nDescending
        End If
    End If
    
End Sub

Public Sub UpdateOrderByColumn(Optional nGrid As Object)
    Dim iGH As cGridHandler
    
    If nGrid Is Nothing Then
        If Not mCurrentGridHandler Is Nothing Then
            mCurrentGridHandler.UpdateOrderByColumn
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.UpdateOrderByColumn
        End If
    End If
End Sub

Public Property Get GetOrderColumn(Optional nGrid As Object) As Long
    Dim iGH As cGridHandler
    
    If nGrid Is Nothing Then
        If Not mCurrentGridHandler Is Nothing Then
            GetOrderColumn = mCurrentGridHandler.GetOrderColumn
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            GetOrderColumn = iGH.GetOrderColumn
        End If
    End If
End Property

Public Property Get GetOrderColumnDescending(Optional nGrid As Object) As Boolean
    Dim iGH As cGridHandler
    
    If nGrid Is Nothing Then
        If Not mCurrentGridHandler Is Nothing Then
            GetOrderColumnDescending = mCurrentGridHandler.GetOrderColumnDescending
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            GetOrderColumnDescending = iGH.GetOrderColumnDescending
        End If
    End If
End Property

Public Property Get EditingCell(Optional nGrid As Object) As Boolean
    Dim iGH As cGridHandler
    
    If nGrid Is Nothing Then
        If Not mCurrentGridHandler Is Nothing Then
            EditingCell = mCurrentGridHandler.EditingCell
        Else
            For Each iGH In mGridHandlersCollection
                If iGH.EditingCell Then
                    EditingCell = True
                    Exit For
                End If
            Next
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            EditingCell = iGH.EditingCell
        End If
    End If
End Property

Public Sub CopyTextEditingCell(Optional nGrid As Object)
    Dim iGH As cGridHandler
    
    If nGrid Is Nothing Then
        If Not mCurrentGridHandler Is Nothing Then
            mCurrentGridHandler.CopyTextEditingCell
        Else
            For Each iGH In mGridHandlersCollection
                If iGH.EditingCell Then
                    iGH.CopyTextEditingCell
                    Exit For
                End If
            Next
        End If
    Else
        Set iGH = GetGridHandler(nGrid)
        If Not iGH Is Nothing Then
            iGH.CopyTextEditingCell
        End If
    End If
End Sub

Public Property Let MinScalePercent(nValue As Long)
    If nValue <> mMinScalePercent Then
        If nValue < mMaxScalePercent Then
            mMinScalePercent = nValue
            If mMinScalePercent < cPrintPreviewMinScale Then
                mMinScalePercent = cPrintPreviewMinScale
            End If
        End If
    End If
End Property

Public Property Get MinScalePercent() As Long
    MinScalePercent = mMinScalePercent
End Property


Public Property Let MaxScalePercent(nValue As Long)
    If nValue <> mMaxScalePercent Then
        If nValue > mMinScalePercent Then
            mMaxScalePercent = nValue
            If mMaxScalePercent > cPrintPreviewMaxScale Then
                mMaxScalePercent = cPrintPreviewMaxScale
            End If
        End If
    End If
End Property

Public Property Get MaxScalePercent() As Long
    MaxScalePercent = mMaxScalePercent
End Property


Public Property Let PrintPrevUseAltScaleIcons(nValue As Boolean)
    If nValue <> mPrintFnObject.PrintPrevUseAltScaleIcons Then
        mPrintFnObject.PrintPrevUseAltScaleIcons = nValue
    End If
End Property

Public Property Get PrintPrevUseAltScaleIcons() As Boolean
    PrintPrevUseAltScaleIcons = mPrintFnObject.PrintPrevUseAltScaleIcons
End Property

