VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrinterEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements VB.Printer
Implements IVBPrint
Implements IPrinterEx

'Private Const LF_FACESIZE = 32
Private Type LOGFONTW
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(0 To ((LF_FACESIZE * 2) - 1)) As Byte
End Type

Private Declare Function CreateFontIndirectW Lib "gdi32" (ByRef lpLogFont As LOGFONTW) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type Charrange
  cpMin As Long     ' First character of range (0 for start of doc)
  cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Private Type FormatRange
  hDC As Long       ' Actual DC to draw on
  hdcTarget As Long ' Target DC for determining text formatting
  RC As RECT        ' Region of the DC to draw to (in twips)
  rcPage As RECT    ' Region of the entire DC (page size) (in twips)
  chrg As Charrange ' Range of text to draw (see above declaration)
End Type

Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57

Private Declare Function CreateDCByLong Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long

Private Declare Function GetStretchBltMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Arc Lib "gdi32" (ByVal hDC As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long, ByVal nXStartArc As Long, ByVal nYStartArc As Long, ByVal nXEndArc As Long, ByVal nYEndArc As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Long

'Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDC As Long, lpInitData As Any) As Long

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Const DC_PAPERS = 2
'Const DC_PAPERNAMES = 16
Const DC_PAPERSIZE = 3

Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, lpDevMode As Any) As Long

Private Type DOCINFO
        cbSize As Long
        lpszDocName As String
        lpszOutput As String
        lpszDatatype As String
        fwType As Long
End Type

Private Declare Function StartDocAPI Lib "gdi32" Alias "StartDocA" (ByVal hDC As Long, lpdi As DOCINFO) As Long
Private Declare Function StartPage Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function EndPage Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function EndDocAPI Lib "gdi32" Alias "EndDoc" (ByVal hDC As Long) As Long

Private Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hDC As Long, ByVal hemf As Long, lpRect As RECT) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long

Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nDrawStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private Const BS_SOLID As Long = 0
'Private Const BS_NULL As Long = 1
'Private Const BS_HATCHED As Long = 2

Private Enum eLineFlags
    LF_STEP1 = 1        ' // First step was specified
    LF_COLOR = 2        ' // Color was specified
    LF_FIRSTCOORD = 4   ' // First coords were specified
    LF_STEP2 = 8        ' // Second step was specified
    LF_B = &H10         ' // B - flag
    LF_BF = &H20        ' // BF - flag
End Enum

Private Declare Function CreateBrushIndirect Lib "gdi32.dll" (lpLogBrush As LOGBRUSH) As Long

Private mPages() As cPrinterExPage
Private mCurrentXPixels As Single
Private mCurrentYPixels As Single
Private mCurrentPage As Long

Private mPrinterDefaultPageWidthInMillimeters As Single
Private mPrinterDefaultPageHeightInMillimeters As Single
Private mPrinterDefaultPageWidthInPixels As Long
Private mPrinterDefaultPageHeightInPixels As Long
Private mDPIX As Single
Private mDPIY As Single
Private mPaperWidthInMillimeters As Single
Private mPaperHeightInMillimeters As Single

Private mHdcReference As Long
Private mLogFont As LOGFONTW

Private Const cMaxPages As Long = 1500

Private mDrawMode As Long
Private mDrawStyle As Long
Private mDrawWidth As Long
Private mForeColor As Long
Private mFontTransparent As Boolean
Private mFillColor As Long
Private mFillStyle As Long
Private mDeviceName As String
Private mPort As String
Private mDriverName As String
Private mOrientation As Long
Private mPaperSize As Long
Private mPaperBin As Long
Private mPrintQuality As Long
Private mColorMode As Long
Private mDuplex As Long
Private mScaleMode As Long
Private mScaleLeftPixels As Single
Private mScaleTopPixels As Single
Private mScaleWidthUser As Single
Private mScaleHeightUser As Single
Private mLeftMarginInMillimeters As Single
Private mRightMarginInMillimeters As Single
Private mTopMarginInMillimeters As Single
Private mBottomMarginInMillimeters As Single
Private mFonts() As String
Private mFontCount As Long
Private mZoom As Long
Private mTrackDefault As Boolean
Private mFontName As String
Private mFontItalic As Boolean
Private mFontStrikethrough As Boolean
Private mFontUnderline As Boolean
Private mFontSize As Single
Private mFontWeight As Long
Private mFontWidth As Long
Private mFontCharset As Long
Private mFontBold As Boolean

Private mDevModePtr As Long

Private mHandleMargins As Boolean
Private mFromPage As Long
Private mToPage As Long
Private mDocumentName As String
Private mDocKey As String
Private mPrintPageNumbers As Boolean
Private WithEvents mPageNumbersFont As StdFont
Attribute mPageNumbersFont.VB_VarHelpID = -1
Private mPageNumbersForeColor As Long
Private mPageNumbersPosition As vbExPageNumbersPositionConstants
Private mPageNumbersFormat As String
Private mPrinted As Boolean

Private mNonPrintableAreaLeftInPixels As Long
Private mNonPrintableAreaTopInPixels As Long
Private mNonPrintableAreaRightInPixels As Long
Private mNonPrintableAreaBottomInPixels As Long
Private mNonPrintableAreaLeftInMillimeters As Single
Private mNonPrintableAreaTopInMillimeters As Single
Private mNonPrintableAreaRightInMillimeters As Single
Private mNonPrintableAreaBottomInMillimeters As Single
Private mPrintPrevFormatButtonVisible As Boolean
Private mPrintPrevFormatButtonToolTipText As String
Private mPrintPrevFormatButtonPicture(4) As Object
Private mPrintPrevPageSetupButtonVisible As Boolean
Private mPrintPrevMinScalePercent As Long
Private mPrintPrevMaxScalePercent As Long
Private mPrintPrevToolBarIconsSize As vbExPrintPrevToolBarIconsSizeConstants
Private mPrintPrevUseAltScaleIcons As Boolean
Private mPrinterFlags As cdeCommonDialogExPrinterFlagsConstants
Private mPageSetupFlags As cdeCommonDialogExPageSetupFlagsConstants

Private mFontName_Prev As String
Private mFontItalic_Prev As Boolean
Private mFontStrikethrough_Prev As Boolean
Private mFontUnderline_Prev As Boolean
Private mFontSize_Prev As Long
Private mFontWeight_Prev As Single
Private mFontWidth_Prev As Long
Private mFontCharset_Prev As Long

Private mFontsUsed_LogFont() As LOGFONTW
Private mhFontsUsed() As Long
Private mFontsUsed_HeightInPixels() As Long
Private mCurrentFontIndex As Long
Private mFontChanged As Boolean ' used for the Property Get Font

Private mDocumentCreated As Boolean
Private mPageCreated As Boolean
Private mSomethingPrintedOnDoc As Boolean
Private mNewPagePending As Boolean
Private mPaperSizeSet As Boolean
Private mDocumentTooLongWarningPrinted As Boolean
Private mPrintingDocumentTooLongWarning As Boolean
Private mForeColorTranslated As Long
Private mFillColorTranslated As Long
Private mPageNumbers() As Long
Private mIVBPrintColumn As Long
Private mNextPageNumberIsOne As Boolean
Private mLastPageAddedByEndDoc As Boolean
Private mDlg As CommonDialogExObject
Private mCanceled As Boolean
Private mRaisingEventsStartPage As Boolean
Private mPrinterExEvents As PrinterExEvents
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1
Private mDisableEvents As Boolean
Private WithEvents mPrintFnObject As PrintFnObject
Attribute mPrintFnObject.VB_VarHelpID = -1
Private mPrinterDefaultOrientation As Long
Private mFixedImageScale As Single
Private mTopMarginSet As Single
Private mBottomMarginSet As Single
Private mCheckingMarginsLimits As Boolean
Private mLastEndDoc_FirstPageIndex As Long
Private mLastEndDoc_LastPageIndex As Long
Private mDefaultColorMode As Long
Private mDontCheckError396 As Boolean
Private mShowingPrintPreview As Boolean
Private mPrintingPageNumbers As Boolean
Private mPageSetupMarginsDisabledAutomatically As Boolean

Private mSave_ScalePercent As Boolean
Private mSave_PrintPageNumbers  As Boolean
Private mSave_PageNumbersPosition  As Boolean
Private mSave_PageNumbersFormat  As Boolean
Private mSave_PageNumbersFont  As Boolean
Private mSave_PageNumbersForeColor  As Boolean
Private mSave_Orientation As Boolean
Private mSave_ColorMode As Boolean
Private mSave_PaperSize As Boolean
Private mSave_DeviceName As Boolean
Private mSave_PaperBin As Boolean
Private mSave_PrintQuality As Boolean
Private mSave_Collate As Boolean
Private mSave_Duplex As Boolean
Private mSave_LeftMargin As Boolean
Private mSave_TopMargin As Boolean
Private mSave_RightMargin As Boolean
Private mSave_BottomMargin As Boolean

Private mLoaded_ScalePercent As Boolean
Private mLoaded_PrintPageNumbers  As Boolean
Private mLoaded_PageNumbersPosition  As Boolean
Private mLoaded_PageNumbersFormat  As Boolean
Private mLoaded_PageNumbersFont  As Boolean
Private mLoaded_PageNumbersForeColor  As Boolean
Private mLoaded_Orientation As Boolean
Private mLoaded_ColorMode As Boolean
Private mLoaded_PaperSize As Boolean
Private mLoaded_DeviceName As Boolean
Private mLoaded_PaperBin As Boolean
Private mLoaded_PrintQuality As Boolean
Private mLoaded_Collate As Boolean
Private mLoaded_Duplex As Boolean
Private mLoaded_LeftMargin As Boolean
Private mLoaded_TopMargin As Boolean
Private mLoaded_RightMargin As Boolean
Private mLoaded_BottomMargin As Boolean

Private mOriginal_ScalePercent As Long
Private mOriginal_PrintPageNumbers  As Boolean
Private mOriginal_PageNumbersPosition  As Long
Private mOriginal_PageNumbersFormat  As String
Private mOriginal_PageNumbersFont  As Object
Private mOriginal_PageNumbersForeColor  As Long
Private mOriginal_Orientation As Long
Private mOriginal_ColorMode As Long
Private mOriginal_PaperSize As Long
Private mOriginal_DeviceName As String
Private mOriginal_PaperBin As Long
Private mOriginal_PrintQuality As Long
Private mOriginal_Collate As Boolean
Private mOriginal_Duplex As Long
Private mOriginal_LeftMargin As Single
Private mOriginal_TopMargin As Single
Private mOriginal_RightMargin As Single
Private mOriginal_BottomMargin As Single


Private Sub CreateDocument()
    Dim iHdc As Long
    
    Dim iDefaultPrintableAreaWidthInPixels As Long
    Dim iDefaultPrintableAreaHeightInPixels As Long
    
    mPrinterDefaultPageWidthInPixels = GetDeviceCaps(HdcReference, PHYSICALWIDTH)
    If mPrinterDefaultPageWidthInPixels = 0 Then
        DeleteDC mHdcReference
        mHdcReference = 0
        mPrinterDefaultPageWidthInPixels = GetDeviceCaps(HdcReference, PHYSICALWIDTH)
        If mPrinterDefaultPageWidthInPixels = 0 Then
            On Error Resume Next
        End If
    End If
    iHdc = HdcReference
    
    mPrinterDefaultPageHeightInPixels = GetDeviceCaps(iHdc, PHYSICALHEIGHT)
    
    mPrinterDefaultPageWidthInMillimeters = Round(mPrinterDefaultPageWidthInPixels / GetDeviceCaps(iHdc, LOGPIXELSX) * 25.4 * 4) / 4
    mPrinterDefaultPageHeightInMillimeters = Round(mPrinterDefaultPageHeightInPixels / GetDeviceCaps(iHdc, LOGPIXELSY) * 25.4 * 4) / 4
    
    mDPIX = Round(25.4 * mPrinterDefaultPageWidthInPixels / mPrinterDefaultPageWidthInMillimeters)
    mDPIY = Round(25.4 * mPrinterDefaultPageHeightInPixels / mPrinterDefaultPageHeightInMillimeters)
    
    iDefaultPrintableAreaWidthInPixels = GetDeviceCaps(iHdc, HORZRES)
    iDefaultPrintableAreaHeightInPixels = GetDeviceCaps(iHdc, VERTRES)
    
    mNonPrintableAreaLeftInPixels = GetDeviceCaps(iHdc, PHYSICALOFFSETX)
    mNonPrintableAreaTopInPixels = GetDeviceCaps(iHdc, PHYSICALOFFSETY)
    mNonPrintableAreaRightInPixels = (mPrinterDefaultPageWidthInPixels - iDefaultPrintableAreaWidthInPixels) - mNonPrintableAreaLeftInPixels
    mNonPrintableAreaBottomInPixels = (mPrinterDefaultPageHeightInPixels - iDefaultPrintableAreaHeightInPixels) - mNonPrintableAreaTopInPixels
    
    mNonPrintableAreaLeftInMillimeters = ScaleX(mNonPrintableAreaLeftInPixels, vbPixels, vbMillimeters)
    mNonPrintableAreaTopInMillimeters = ScaleY(mNonPrintableAreaTopInPixels, vbPixels, vbMillimeters)
    mNonPrintableAreaRightInMillimeters = ScaleX(mNonPrintableAreaRightInPixels, vbPixels, vbMillimeters)
    mNonPrintableAreaBottomInMillimeters = ScaleY(mNonPrintableAreaBottomInPixels, vbPixels, vbMillimeters)
    
    mDocumentCreated = True
    
    ReDim mPages(0)
    ReDim mPageNumbers(0)
    NewPage
    mSomethingPrintedOnDoc = False
End Sub

Public Sub PrintDocumentToPrinterDC(nHdc As Long)
    Dim iPC As Long
    Dim iFromPage As Long
    Dim iToPage As Long
    Dim c As Long
    
    iPC = PageCount
    If mFromPage = 0 Then
        iFromPage = 1
    Else
        iFromPage = mFromPage
    End If
    If iFromPage > iPC Then Exit Sub
    
    If mToPage = 0 Then
        iToPage = iPC
    Else
        If mToPage > iPC Then
            iToPage = iPC
        Else
            iToPage = mToPage
        End If
    End If
    
    For c = 2 To iToPage
        If mPageNumbers(c) <= mPageNumbers(c - 1) Then
            PrintDocumentToPrinterDC1 nHdc, iFromPage, c - 1
            iFromPage = c
        End If
    Next c
    PrintDocumentToPrinterDC1 nHdc, iFromPage, iToPage
    mPrinted = True
    mPrinterExEvents.RaiseEvent_DocPrinted mDocKey
End Sub

Private Sub PrintDocumentToPrinterDC1(nHdc As Long, nFromPage As Long, nToPage As Long)
    Dim iFromPage As Long
    Dim iToPage As Long
    Dim iPn As Long
    Dim iPC As Long
    Dim iDocumentName As String
    Dim iDI As DOCINFO
    Dim ihMetaFile As Long
    Dim iOrientationPrev As Long
    Dim iPaperSizePrev As Long
    Dim iPaperBinPrev As Long
    Dim iPrintQualityPrev As Long
    Dim iDuplexPrev As Long
    Dim iColorModePrev As Long
    Dim iResetDC As Boolean
    Dim iDevMode As DEVMODE
    Dim iDCWasReset As Boolean
    Dim iOriginalDevMode As DEVMODE
    Dim iRectPage As RECT
    
'    Exit Sub
    
    If mDevModePtr <> 0 Then
        CopyMemory iOriginalDevMode, ByVal mDevModePtr, Len(iOriginalDevMode)
        iOrientationPrev = iOriginalDevMode.dmOrientation
        iPaperBinPrev = iOriginalDevMode.dmPaperSize
        iPaperSizePrev = iOriginalDevMode.dmDefaultSource
        iPrintQualityPrev = iOriginalDevMode.dmPrintQuality
        iDuplexPrev = iOriginalDevMode.dmDuplex
        iColorModePrev = iOriginalDevMode.dmColor
    End If
    
    iPC = PageCount
    If nFromPage = 0 Then
        iFromPage = 1
    Else
        iFromPage = nFromPage
    End If
    If iFromPage > iPC Then Exit Sub
    
    If nToPage = 0 Then
        iToPage = iPC
    Else
        iToPage = nToPage
    End If
    
    If mPages(iToPage).DocumentName <> "" Then
        iDocumentName = ClientProductName & " - " & mPages(iToPage).DocumentName
    Else
        iDocumentName = ClientProductName
    End If
    
    iDI.cbSize = Len(iDI)
    iDI.lpszDocName = iDocumentName
    iDI.lpszOutput = vbNullString
    iDI.lpszDatatype = vbNullString
    Call StartDocAPI(nHdc, iDI)
    
    For iPn = iFromPage To iToPage Step 1
        If iPn <= iPC Then
            If Not mPages(iPn) Is Nothing Then
                
                iRectPage.Left = 0
                iRectPage.Top = 0
                iRectPage.Right = mPages(iPn).WidthInPixels / 100 * mZoom ' GetDeviceCaps(nHdc, PHYSICALWIDTH) ' iRectPage.Left + mPages(iPN).WidthInPixels
                iRectPage.Bottom = mPages(iPn).HeightInPixels / 100 * mZoom ' GetDeviceCaps(nHdc, PHYSICALHEIGHT) ' iRectPage.Top + mPages(iPN).HeightInPixels
                
                If mDevModePtr <> 0 Then
                    iResetDC = False
                    If mPages(iPn).Orientation <> iOrientationPrev Then iResetDC = True
                    If mPages(iPn).PaperSize <> iPaperBinPrev Then iResetDC = True
                    If mPages(iPn).PaperBin <> iPaperSizePrev Then iResetDC = True
                    If mPages(iPn).PrintQuality <> iPrintQualityPrev Then iResetDC = True
                    If mPages(iPn).Duplex <> iDuplexPrev Then iResetDC = True
                    If mPages(iPn).ColorMode <> iColorModePrev Then iResetDC = True
                    
                    If iResetDC Then
                        CopyMemory iDevMode, ByVal mDevModePtr, Len(iDevMode)
                        iDevMode.dmOrientation = mPages(iPn).Orientation
                        iDevMode.dmPaperSize = mPages(iPn).PaperSize
                        iDevMode.dmDefaultSource = mPages(iPn).PaperBin
                        iDevMode.dmDuplex = mPages(iPn).Duplex
                        iDevMode.dmPrintQuality = mPages(iPn).PrintQuality
                        iDevMode.dmColor = mPages(iPn).ColorMode
                        CopyMemory ByVal mDevModePtr, iDevMode, Len(iDevMode)
                        ResetDC nHdc, iDevMode
                        iDCWasReset = True
                        iRectPage.Right = GetDeviceCaps(nHdc, PHYSICALWIDTH) ' iRectPage.Left + mPages(iPN).WidthInPixels
                        iRectPage.Bottom = GetDeviceCaps(nHdc, PHYSICALHEIGHT) ' iRectPage.Top + mPages(iPN).HeightInPixels
                    End If
                End If
                
                iOrientationPrev = mPages(iPn).Orientation
                iPaperBinPrev = mPages(iPn).PaperSize
                iPaperSizePrev = mPages(iPn).PaperBin
                iPrintQualityPrev = mPages(iPn).PrintQuality
                iDuplexPrev = mPages(iPn).Duplex
                iColorModePrev = mPages(iPn).ColorMode
                
                ihMetaFile = mPages(iPn).hMetaFile
                If ihMetaFile <> 0 Then
                    Call StartPage(nHdc)
                    PlayEnhMetaFile nHdc, ihMetaFile, iRectPage
                    Call EndPage(nHdc)
                End If
            End If
        End If
    Next iPn
    Call EndDocAPI(nHdc)
    
    If iDCWasReset Then
        ResetDC nHdc, iOriginalDevMode
    End If
End Sub


Public Function PaintPageIntoPictureBox(nPageNumber As Long, nPic As PictureBox) As Boolean
    Dim iRect As RECT
    Dim iPn As Long
    Dim ihMetaFile As Long
    Dim iAuxSM As Long
    
    If nPageNumber = 0 Then
        iPn = mCurrentPage
    Else
        iPn = nPageNumber
    End If
    If (iPn < 1) Or (iPn > PageCount) Then Exit Function
    
    If Not mPages(iPn) Is Nothing Then
        ihMetaFile = mPages(iPn).hMetaFile
        If ihMetaFile <> 0 Then
            iAuxSM = nPic.ScaleMode
            nPic.ScaleMode = vbPixels
            iRect.Right = nPic.ScaleWidth
            iRect.Bottom = nPic.ScaleHeight
            nPic.ScaleMode = iAuxSM
            nPic.Cls
            nPic.AutoRedraw = True
            On Error Resume Next
            PaintPageIntoPictureBox = PlayEnhMetaFile(nPic.hDC, ihMetaFile, iRect) <> 0
        End If
    End If
End Function


Public Function NewPage() As Boolean
    If mRaisingEventsStartPage Then Exit Function
    
    If PageCount > cMaxPages Then
        mCurrentXPixels = -mScaleLeftPixels + PageLeftMarginInPixels
        mCurrentYPixels = -mScaleTopPixels + PageTopMarginInPixels
        Exit Function
    End If
    Err.Clear
    On Error Resume Next
    mCurrentPage = UBound(mPages) + 1
    If Err.Number = 9 Then
        EnsureDocumentCreated2
        mNewPagePending = True
        AddNewPage
        mCurrentPage = UBound(mPages) + 1
    ElseIf mNewPagePending Then
        AddNewPage
        mCurrentPage = UBound(mPages) + 1
    End If
    On Error GoTo 0
    
    mNewPagePending = True
    
    mCurrentXPixels = -mScaleLeftPixels + PageLeftMarginInPixels
    mCurrentYPixels = -mScaleTopPixels + PageTopMarginInPixels
    EnsureDocumentCreated2 'SetPixelScale
    NewPage = True
End Function

Private Sub AddNewPage()
    Dim iStartDocEvent As Boolean
    Static sLast_PaperWidthInMillimeters As Single
    Static sLast_PaperHeightInMillimeters As Single
    Static sLast_Orientation As Single
    Static sLast_LeftMarginInMillimeters As Single
    Static sLast_TopMarginInMillimeters As Single
    Static sLast_RightMarginInMillimeters As Single
    Static sLast_BottomMarginInMillimeters As Single
    
    If mRaisingEventsStartPage Then Exit Sub
    mCurrentPage = UBound(mPages) + 1
    ReDim Preserve mPages(mCurrentPage)
    ReDim Preserve mPageNumbers(mCurrentPage)
    If mNextPageNumberIsOne Then
        mPageNumbers(mCurrentPage) = 1
        mNextPageNumberIsOne = False
        mLastPageAddedByEndDoc = True
    Else
        mPageNumbers(mCurrentPage) = mPageNumbers(mCurrentPage - 1) + 1
        mLastPageAddedByEndDoc = False
    End If
    If mPageNumbers(mCurrentPage) = 1 Then
        iStartDocEvent = True
    End If
    Set mPages(mCurrentPage) = New cPrinterExPage
    mPages(mCurrentPage).CreatePage HdcReference, PaperWidthInMillimeters, PaperHeightInMillimeters, mOrientation, mhFontsUsed(mCurrentFontIndex), mDrawStyle, mDrawWidth, mForeColorTranslated, mDrawMode, mFontTransparent, mFillStyle, mFillColorTranslated, mLeftMarginInMillimeters, mTopMarginInMillimeters, mRightMarginInMillimeters, mBottomMarginInMillimeters, mPaperSize, mPaperBin, mPrintQuality, mDuplex, mColorMode, mNonPrintableAreaLeftInPixels, mNonPrintableAreaTopInPixels, mNonPrintableAreaRightInPixels, mNonPrintableAreaBottomInPixels, mDPIX, mDPIY, mZoom, mFixedImageScale, mDocumentName
    SetTextColor mPages(mCurrentPage).MetaDC, mForeColorTranslated
    mNewPagePending = False
    mPageCreated = True
    
    mCurrentXPixels = -mScaleLeftPixels + PageLeftMarginInPixels
    mCurrentYPixels = -mScaleTopPixels + PageTopMarginInPixels
    
    If Not mDisableEvents Then
        If (mCurrentPage = 1) Or (sLast_PaperWidthInMillimeters <> mPaperWidthInMillimeters) Or (sLast_PaperHeightInMillimeters <> mPaperHeightInMillimeters) Or (sLast_Orientation <> mOrientation) Or (sLast_LeftMarginInMillimeters <> mLeftMarginInMillimeters) Or (sLast_TopMarginInMillimeters <> mTopMarginInMillimeters) Or (sLast_RightMarginInMillimeters <> mRightMarginInMillimeters) Or (sLast_BottomMarginInMillimeters <> mBottomMarginInMillimeters) Then
            CheckMarginsLimits
            mPages(mCurrentPage).FixedImageScale = mFixedImageScale
            sLast_PaperWidthInMillimeters = mPaperWidthInMillimeters
            sLast_PaperHeightInMillimeters = mPaperHeightInMillimeters
            sLast_Orientation = mOrientation
            sLast_LeftMarginInMillimeters = mLeftMarginInMillimeters
            sLast_TopMarginInMillimeters = mTopMarginInMillimeters
            sLast_RightMarginInMillimeters = mRightMarginInMillimeters
            sLast_BottomMarginInMillimeters = mBottomMarginInMillimeters
        End If
        
        If iStartDocEvent Then
            mLastEndDoc_FirstPageIndex = 0
            mLastEndDoc_LastPageIndex = 0
            mPrinted = False
            mPrinterExEvents.RaiseEvent_StartDoc mDocKey
            mSomethingPrintedOnDoc = False
        End If
        
        ' raise event StartPage
        Dim iCurrentPage_Prev As Long
        Dim iFont_Prev As StdFont
        Dim iFontName_Prev As String
        Dim iFontSize_Prev As Single
        Dim iFontBold_prev As Boolean
        Dim iFontItalic_Prev As Boolean
        Dim iFontStrikethrough_Prev As Boolean
        Dim iFontUnderline_Prev As Boolean
        Dim iFontWidth_Prev As Long
        Dim iDrawMode_Prev As Long
        Dim iDrawStyle_Prev As Long
        Dim iDrawWidth_Prev As Long
        Dim iForeColor_Prev As Long
        Dim iFontTransparent_Prev As Boolean
        Dim iFillColor_Prev As Long
        Dim iFillStyle_Prev As Long
        Dim iHandleMargins_Prev As Boolean
        
        iCurrentPage_Prev = mCurrentPage
        Set iFont_Prev = Font
        iHandleMargins_Prev = mHandleMargins
        
        iFontName_Prev = mFontName
        iFontSize_Prev = mFontSize
        iFontBold_prev = mFontBold
        iFontItalic_Prev = mFontItalic
        iFontStrikethrough_Prev = mFontStrikethrough
        iFontUnderline_Prev = mFontUnderline
        iFontWidth_Prev = mFontWeight
        
        iDrawMode_Prev = mDrawMode
        iDrawStyle_Prev = mDrawStyle
        iDrawWidth_Prev = mDrawWidth
        iForeColor_Prev = mForeColor
        iFontTransparent_Prev = mFontTransparent
        iFillColor_Prev = mFillColor
        iFillStyle_Prev = mFillStyle
        
        mRaisingEventsStartPage = True
        mPrinterExEvents.RaiseEvent_StartPage mDocKey
        mRaisingEventsStartPage = False
    
        mCurrentPage = iCurrentPage_Prev
        If mHandleMargins <> iHandleMargins_Prev Then HandleMargins = iHandleMargins_Prev
        mCurrentXPixels = -mScaleLeftPixels + PageLeftMarginInPixels
        mCurrentYPixels = -mScaleTopPixels + PageTopMarginInPixels
        If Not mFont Is iFont_Prev Then Set Font = iFont_Prev
        
        If mFontName <> iFontName_Prev Then FontName = iFontName_Prev
        If mFontSize <> iFontSize_Prev Then FontSize = iFontSize_Prev
        If mFontBold <> iFontBold_prev Then FontBold = iFontBold_prev
        If mFontItalic <> iFontItalic_Prev Then FontItalic = iFontItalic_Prev
        If mFontStrikethrough <> iFontStrikethrough_Prev Then FontStrikeThru = iFontStrikethrough_Prev
        If mFontUnderline <> iFontUnderline_Prev Then FontUnderLine = iFontUnderline_Prev
        If mFontWeight <> iFontWidth_Prev Then mFont.Weight = iFontWidth_Prev
        
        If mDrawMode <> iDrawMode_Prev Then DrawMode = iDrawMode_Prev
        If mDrawStyle <> iDrawStyle_Prev Then DrawStyle = iDrawStyle_Prev
        If mDrawWidth <> iDrawWidth_Prev Then DrawWidth = iDrawWidth_Prev
        If mForeColor <> iForeColor_Prev Then ForeColor = iForeColor_Prev
        If mFontTransparent <> iFontTransparent_Prev Then FontTransparent = iFontTransparent_Prev
        If mFillColor <> iFillColor_Prev Then FillColor = iFillColor_Prev
        If mFillStyle <> iFillStyle_Prev Then FillStyle = iFillStyle_Prev
    
    End If
    
    SomethingPrintedOnPage = False
End Sub

Public Sub KillDoc()
    mNewPagePending = False
    mCurrentPage = 0
    ReDim mPages(0)
    ReDim mPageNumbers(0)
    mPageCreated = False
    mCurrentXPixels = -mScaleLeftPixels + PageLeftMarginInPixels
    mCurrentYPixels = -mScaleTopPixels + PageTopMarginInPixels
    mDocumentCreated = False
    mSomethingPrintedOnDoc = False
End Sub

Private Sub CheckNewPagePending()
    If mNewPagePending Then AddNewPage
End Sub

Public Sub GoToPage(nPageNumber As Long)
    Dim iMetaDC As Long
    
    If mRaisingEventsStartPage Then Exit Sub
    If nPageNumber = mCurrentPage Then Exit Sub
    mNewPagePending = False
    If (nPageNumber < 1) Or (nPageNumber > PageCount) Then
        If nPageNumber = (PageCount + 1) Then
            AddNewPage
        Else
            RaiseError 9, "PrinterEx"
            Exit Sub
        End If
    End If
    mCurrentPage = nPageNumber
    
    iMetaDC = mPages(mCurrentPage).MetaDC
    If iMetaDC <> 0 Then
        SetTextColor iMetaDC, mForeColorTranslated
        mPages(mCurrentPage).SelectFont mhFontsUsed(mCurrentFontIndex)
        SetPen
        SetBrush
        SetBackMode
    End If
    mCurrentXPixels = -mScaleLeftPixels + PageLeftMarginInPixels
    mCurrentYPixels = -mScaleTopPixels + PageTopMarginInPixels

End Sub


Public Property Get Page() As Long
    Page = mCurrentPage
End Property


Public Property Get CurrentX() As Single
    CurrentX = ScaleX(mCurrentXPixels + mScaleLeftPixels - PageLeftMarginInPixels, vbPixels, mScaleMode)
End Property

Public Property Let CurrentX(nValue As Single)
    mCurrentXPixels = ScaleX(nValue, mScaleMode, vbPixels) - mScaleLeftPixels + PageLeftMarginInPixels
End Property


Public Property Get CurrentY() As Single
    CurrentY = ScaleY(mCurrentYPixels + mScaleTopPixels - PageTopMarginInPixels, vbPixels, mScaleMode)
End Property

Public Property Let CurrentY(nValue As Single)
    mCurrentYPixels = ScaleY(nValue, mScaleMode, vbPixels) - mScaleTopPixels + PageTopMarginInPixels
End Property


Public Property Get FromPage() As Long
    FromPage = mFromPage
End Property

Public Property Let FromPage(nValue As Long)
    mFromPage = nValue
    If mFromPage < 0 Then mFromPage = 0
    If mFromPage > mToPage Then mToPage = mFromPage
End Property


Public Property Get ToPage() As Long
    ToPage = mToPage
End Property

Public Property Let ToPage(nValue As Long)
    mToPage = nValue
    If mToPage < 0 Then mToPage = 0
    If mToPage < mFromPage Then mFromPage = mToPage
End Property


Public Property Get HandleMargins() As Boolean
    HandleMargins = mHandleMargins
End Property

Public Property Let HandleMargins(nValue As Boolean)
    If nValue <> mHandleMargins Then
        mHandleMargins = nValue
        If mHandleMargins Then
            mCurrentXPixels = mCurrentXPixels - PageLeftMarginInPixels + mNonPrintableAreaLeftInPixels
            mCurrentYPixels = mCurrentYPixels - PageTopMarginInPixels + mNonPrintableAreaTopInPixels
            If mCurrentXPixels < (PageLeftMarginInPixels - mScaleLeftPixels) Then
                mCurrentXPixels = PageLeftMarginInPixels - mScaleLeftPixels
            End If
            If mCurrentYPixels < (PageTopMarginInPixels - mScaleTopPixels) Then
                mCurrentYPixels = PageTopMarginInPixels - mScaleTopPixels
            End If
            If mPageSetupMarginsDisabledAutomatically Then
                mPageSetupMarginsDisabledAutomatically = False
                mPageSetupFlags = mPageSetupFlags And Not cdePSDisableMargins
            End If
        Else
            mCurrentXPixels = mCurrentXPixels + PageLeftMarginInPixels - mNonPrintableAreaLeftInPixels
            mCurrentYPixels = mCurrentYPixels + PageTopMarginInPixels - mNonPrintableAreaTopInPixels
            If (mPageSetupFlags And cdePSDisableMargins) = 0 Then
                mPageSetupFlags = mPageSetupFlags Or cdePSDisableMargins
                mPageSetupMarginsDisabledAutomatically = True
            End If
        End If
    End If
End Property


Public Property Get DrawMode() As DrawModeConstants
    DrawMode = mDrawMode
End Property

Public Property Let DrawMode(nValue As DrawModeConstants)
    If nValue <> mDrawMode Then
        mDrawMode = nValue
        SetPen
    End If
End Property


Public Property Get DrawStyle() As DrawStyleConstants
    DrawStyle = mDrawStyle
End Property

Public Property Let DrawStyle(nValue As DrawStyleConstants)
    If mDrawStyle <> nValue Then
        mDrawStyle = nValue
        SetPen
    End If
End Property


Public Property Get DrawWidth() As Long
    DrawWidth = mDrawWidth
End Property

Public Property Let DrawWidth(nValue As Long)
    If nValue <> mDrawWidth Then
        mDrawWidth = nValue
        SetPen
    End If
End Property


Public Property Get ForeColor() As Long
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(nValue As Long)
    Dim iColor As Long
    Dim iMetaDC As Long
    Static sMDCWasZero As Boolean
    
    If nValue <> mForeColor Then
        mForeColor = nValue
        TranslateColor nValue, 0, iColor
        If mCurrentPage > 0 Then
            If mCurrentPage <= UBound(mPages) Then
                iMetaDC = mPages(mCurrentPage).MetaDC
                If iMetaDC <> 0 Then
                    If (iColor <> mForeColorTranslated) Or sMDCWasZero Then
                        mForeColorTranslated = iColor
                        SetTextColor iMetaDC, mForeColorTranslated
                        SetPen
                    End If
                    sMDCWasZero = False
                Else
                    sMDCWasZero = True
                    mForeColorTranslated = iColor
                End If
            Else
                sMDCWasZero = True
                mForeColorTranslated = iColor
            End If
        Else
            sMDCWasZero = True
            mForeColorTranslated = iColor
        End If
    End If
End Property


Public Property Get FontTransparent() As Boolean
    FontTransparent = mFontTransparent
End Property

Public Property Let FontTransparent(nValue As Boolean)
    If nValue <> mFontTransparent Then
        mFontTransparent = nValue
        SetBackMode
    End If
End Property


Public Property Get FillColor() As Long
    FillColor = mFillColor
End Property

Public Property Let FillColor(nValue As Long)
    Dim iColor As Long
    
    If nValue <> mFillColor Then
        mFillColor = nValue
        TranslateColor nValue, 0, iColor
        If iColor <> mFillColorTranslated Then
            mFillColorTranslated = iColor
            SetBrush
        End If
    End If
End Property


Public Property Get FillStyle() As FillStyleConstants
    FillStyle = mFillStyle
End Property

Public Property Let FillStyle(nValue As FillStyleConstants)
    If nValue <> mFillStyle Then
        mFillStyle = nValue
        SetBrush
    End If
End Property


Public Property Get DocumentName() As String
    DocumentName = mDocumentName
End Property

Public Property Let DocumentName(ByVal nValue As String)
    mDocumentName = nValue
    If mPageCreated Then
        mPages(mCurrentPage).DocumentName = mDocumentName
    End If
End Property


Public Property Get DeviceName() As String
    DeviceName = mDeviceName
End Property

Friend Property Let DeviceName(nValue As String)
    Dim iPrn As VB.Printer
    Dim iDlgHasDeviceName  As Boolean
    
    If nValue <> mDeviceName Then
        If Not mDlg Is Nothing Then
            If mDlg.DeviceName = nValue Then
                iDlgHasDeviceName = True
            Else
                mDlg.DoTerminate
                mDlg.ManualTermination = False
            End If
        End If
        If Not iDlgHasDeviceName Then
            Set mDlg = New CommonDialogExObject
            mDlg.ManualTermination = True
            mDlg.DeviceName = nValue
        End If
        mDeviceName = mDlg.DeviceName
        If mDeviceName <> "" Then
            For Each iPrn In VB.Printers
                If iPrn.DeviceName = mDeviceName Then
                    mDriverName = iPrn.DriverName
                    mPort = iPrn.Port
                    
                    If Not iDlgHasDeviceName Then
                        mPaperSize = iPrn.PaperSize
                        mPaperBin = iPrn.PaperBin
                        mDuplex = iPrn.Duplex
                        mOrientation = iPrn.Orientation
                        mPrinterDefaultOrientation = iPrn.Orientation
                        mPrintQuality = iPrn.PrintQuality
                        mColorMode = iPrn.ColorMode
                        mDefaultColorMode = mColorMode
                        
                        mDlg.PaperSize = mPaperSize
                        mDlg.PaperBin = mPaperBin
                        mDlg.Duplex = mDuplex
                        mDlg.Orientation = mOrientation
                        mDlg.PrintQuality = mPrintQuality
                        mDlg.ColorMode = mColorMode
                    Else
                        mPrinterDefaultOrientation = iPrn.Orientation
                        mDefaultColorMode = mColorMode
                        
                        mPaperSize = mDlg.PaperSize
                        mPaperBin = mDlg.PaperBin
                        mDuplex = mDlg.Duplex
                        mOrientation = mDlg.Orientation
                        mPrintQuality = mDlg.PrintQuality
                        mColorMode = mDlg.ColorMode
                    End If
                    
                    If mHdcReference <> 0 Then
                        DeleteDC mHdcReference
                    End If
                    mHdcReference = CreateCompatibleDC(iPrn.hDC)
                    
                    Exit For
                End If
            Next
            LoadPrinterFonts
        End If
    End If
End Property


Public Property Get DriverName() As String
    DriverName = mDriverName
End Property


Public Property Get Port() As String
    Port = mPort
End Property


Public Property Get Orientation() As Integer
    If mPageCreated Then
        Orientation = mPages(mCurrentPage).Orientation
    Else
        Orientation = mOrientation
    End If
End Property

Public Property Let Orientation(nValue As Integer)
    If mDocumentCreated And (Not mNewPagePending) Then
        If mPages(mCurrentPage).SomethingPrintedOnPage Then
            If Not mDontCheckError396 Then
                RaiseError 396, "PrinterEx", "Orientation property cannot be set within a page"
                Exit Property
            End If
        Else
            KillLastPage
        End If
    End If
    
    If nValue <> mOrientation Then
        If nValue = 0 Then
            mOrientation = mPrinterDefaultOrientation
        Else
            mOrientation = nValue
        End If
        If PageCount > 0 Then
            If Not SomethingPrintedOnPage Then
                RecreatePage
            End If
        End If
        mDlg.Orientation = mOrientation
        mPrinterExEvents.RaiseEvent_OrientationChange mOrientation
    End If
End Property

Private Sub RecreatePage()
    Dim iPn As Long
    
    If Not mNewPagePending Then
        iPn = mPageNumbers(mCurrentPage)
        KillLastPage
        NewPage
        CheckNewPagePending
        mPageNumbers(mCurrentPage) = iPn
    End If
End Sub

Public Property Get PaperSize() As Integer
    If mPageCreated Then
        PaperSize = mPages(mCurrentPage).PaperSize
    Else
        PaperSize = mPaperSize
    End If
End Property

Public Property Let PaperSize(nValue As Integer)
    If mDocumentCreated And (Not mNewPagePending) Then
        If mPages(mCurrentPage).SomethingPrintedOnPage Then
            If Not mDontCheckError396 Then
                RaiseError 396, "PrinterEx", "PaperSize property cannot be set within a page"
                Exit Property
            End If
        Else
            KillLastPage
        End If
    End If
    
    If (nValue <> mPaperSize) Or Not mPaperSizeSet Then
        mPaperSize = nValue
        mPaperSizeSet = True
        If PageCount > 0 Then
            If Not SomethingPrintedOnPage Then
                RecreatePage
            End If
        End If
        SetPageSize
        mDlg.PaperSize = mPaperSize
    End If
End Property


Public Property Get PrintQuality() As Integer
    If mPageCreated Then
        PrintQuality = mPages(mCurrentPage).PrintQuality
    Else
        PrintQuality = mPrintQuality
    End If
End Property

Public Property Let PrintQuality(nValue As Integer)
    If mDocumentCreated And (Not mNewPagePending) Then
        If mPages(mCurrentPage).SomethingPrintedOnPage Then
            If Not mDontCheckError396 Then
                RaiseError 396, "PrinterEx", "PrintQuality property cannot be set within a page"
                Exit Property
            End If
        Else
            KillLastPage
        End If
    End If
    
    If nValue <> mPrintQuality Then
        mPrintQuality = nValue
        If PageCount > 0 Then
            If Not SomethingPrintedOnPage Then
                RecreatePage
            End If
        End If
        mDlg.PrintQuality = mPrintQuality
    End If
End Property


Public Property Get ColorMode() As Integer
    If mPageCreated Then
        ColorMode = mPages(mCurrentPage).ColorMode
    Else
        ColorMode = mColorMode
    End If
End Property

Public Property Let ColorMode(nValue As Integer)
    If mDocumentCreated And (Not mNewPagePending) Then
        If mPages(mCurrentPage).SomethingPrintedOnPage Then
            If Not mDontCheckError396 Then
                RaiseError 396, "PrinterEx", "ColorMode property cannot be set within a page"
                Exit Property
            End If
        Else
            KillLastPage
        End If
    End If
    
    If nValue <> mColorMode Then
        If nValue = vbPRCMPrinterDefault Then
            mColorMode = mDefaultColorMode
        Else
            mColorMode = nValue
        End If
        If (mColorMode < vbPRCMPrinterDefault) Or (mColorMode > vbPRCMColor) Then
            mColorMode = vbPRCMColor
        End If
        mDlg.ColorMode = mColorMode
    End If
End Property


Public Property Get PaperBin() As Integer
    If mPageCreated Then
        PaperBin = mPages(mCurrentPage).PaperBin
    Else
        PaperBin = mPaperBin
    End If
End Property

Public Property Let PaperBin(nValue As Integer)
    If mDocumentCreated And (Not mNewPagePending) Then
        If mPages(mCurrentPage).SomethingPrintedOnPage Then
            If Not mDontCheckError396 Then
                RaiseError 396, "PrinterEx", "PaperBin property cannot be set within a page"
                Exit Property
            End If
        Else
            KillLastPage
        End If
    End If
    
    If nValue <> mPaperBin Then
        mPaperBin = nValue
        If PageCount > 0 Then
            If Not SomethingPrintedOnPage Then
                RecreatePage
            End If
        End If
        mDlg.PaperBin = mPaperBin
    End If
End Property


Public Property Get Duplex() As Integer
    If mPageCreated Then
        Duplex = mPages(mCurrentPage).Duplex
    Else
        Duplex = mDuplex
    End If
End Property

Public Property Let Duplex(nValue As Integer)
    If mDocumentCreated And (Not mNewPagePending) Then
        If mPages(mCurrentPage).SomethingPrintedOnPage Then
            If Not mDontCheckError396 Then
                RaiseError 396, "PrinterEx", "Duplex property cannot be set within a page"
                Exit Property
            End If
        Else
            KillLastPage
        End If
    End If
    
    If nValue <> mDuplex Then
        mDuplex = nValue
        If PageCount > 0 Then
            If Not SomethingPrintedOnPage Then
                RecreatePage
            End If
        End If
        mDlg.Duplex = mDuplex
    End If
End Property


Public Property Let DevModePtr(nValue As Long)
    mDevModePtr = nValue
End Property



Public Property Get ScaleMode() As Integer
    ScaleMode = mScaleMode
End Property

Public Property Let ScaleMode(nValue As Integer)
    If nValue <> mScaleMode Then
        If (nValue > 0) And (nValue < 10) Then
'            If nValue = vbHimetric Then Err.Raise 380
            mScaleLeftPixels = 0
            mScaleTopPixels = 0
            mScaleMode = nValue
        End If
    End If
End Property


Public Property Get ScaleLeft() As Single
    ScaleLeft = ScaleX(mScaleLeftPixels, vbPixels, ScaleMode)
End Property

Public Property Let ScaleLeft(nValue As Single)
    Scale1 nValue, ScaleTop, nValue + ScaleWidth, ScaleTop + ScaleHeight
End Property


Public Property Get ScaleTop() As Single
    ScaleTop = ScaleY(mScaleTopPixels, vbPixels, ScaleMode)
End Property

Public Property Let ScaleTop(nValue As Single)
    Scale1 ScaleLeft, nValue, ScaleLeft + ScaleWidth, nValue + ScaleHeight
End Property


Public Property Get ScaleWidth() As Single
    If ScaleMode = vbUser Then
        ScaleWidth = mScaleWidthUser
    Else
        If (mScaleMode = vbTwips) Or (mScaleMode = vbPixels) Then
            ScaleWidth = Round(ScaleX(CurrentPageScaleWidthInPixels, vbPixels, ScaleMode))
        Else
            ScaleWidth = ScaleX(CurrentPageScaleWidthInPixels, vbPixels, ScaleMode)
        End If
    End If
End Property

Public Property Let ScaleWidth(nValue As Single)
    Scale1 ScaleLeft, ScaleTop, ScaleLeft + nValue, ScaleTop + ScaleHeight
End Property


Public Property Get ScaleHeight() As Single
    If ScaleMode = vbUser Then
        ScaleHeight = mScaleHeightUser
    Else
        If (mScaleMode = vbTwips) Or (mScaleMode = vbPixels) Then
            ScaleHeight = Round(ScaleY(CurrentPageScaleHeightInPixels, vbPixels, ScaleMode))
        Else
            ScaleHeight = ScaleY(CurrentPageScaleHeightInPixels, vbPixels, ScaleMode)
        End If
    End If
End Property

Public Property Let ScaleHeight(nValue As Single)
    Scale1 ScaleLeft, ScaleTop, ScaleLeft + ScaleWidth, ScaleTop + nValue
End Property


Public Property Get LeftMargin() As Single ' millimeters
    LeftMargin = mLeftMarginInMillimeters
End Property

Public Property Get PageLeftMargin() As Single
    If mNewPagePending Then
        PageLeftMargin = ScaleX(mLeftMarginInMillimeters, vbMillimeters, mScaleMode) / mZoom * 100
    Else
        PageLeftMargin = ScaleX(mPages(mCurrentPage).LeftMarginInPixels, vbPixels, mScaleMode)
    End If
End Property

Public Property Let LeftMargin(ByVal nValue As Single)
    Dim iChange As Boolean
    
    If nValue < MinLeftMargin Then nValue = MinLeftMargin
    If nValue < mNonPrintableAreaLeftInMillimeters Then nValue = mNonPrintableAreaLeftInMillimeters
    iChange = nValue <> mLeftMarginInMillimeters
    If mDocumentCreated Then
        iChange = iChange Or (nValue <> mPages(mCurrentPage).LeftMarginInMillimeters)
    End If
    
    If iChange Then
        mLeftMarginInMillimeters = nValue
        If mDocumentCreated Then
            mPages(mCurrentPage).LeftMarginInMillimeters = mLeftMarginInMillimeters
            mCurrentXPixels = mCurrentXPixels + mPages(mCurrentPage).LeftMarginInPixels - mPages(mCurrentPage).NonPrintableAreaLeftInPixels - mScaleLeftPixels
        Else
            mCurrentXPixels = -mScaleLeftPixels
        End If
    End If
End Property


Public Property Get RightMargin() As Single ' millimeters
    RightMargin = mRightMarginInMillimeters
End Property

Public Property Get PageRightMargin() As Single
    If mNewPagePending Then
        PageRightMargin = ScaleX(mRightMarginInMillimeters, vbMillimeters, mScaleMode) / mZoom * 100
    Else
        PageRightMargin = ScaleX(mPages(mCurrentPage).RightMarginInPixels, vbPixels, mScaleMode)
    End If
End Property

Public Property Let RightMargin(ByVal nValue As Single)
    Dim iChange As Boolean
    
    If nValue < MinRightMargin Then nValue = MinRightMargin
    If nValue < mNonPrintableAreaRightInMillimeters Then nValue = mNonPrintableAreaRightInMillimeters
    iChange = nValue <> mRightMarginInMillimeters
    If mDocumentCreated Then
        iChange = iChange Or (nValue <> mPages(mCurrentPage).RightMarginInMillimeters)
    End If
    
    If iChange Then
        mRightMarginInMillimeters = nValue
        If mDocumentCreated Then
            mPages(mCurrentPage).RightMarginInMillimeters = mRightMarginInMillimeters
        End If
    End If
End Property


Public Property Get TopMargin() As Single ' millimeters
    TopMargin = mTopMarginInMillimeters
End Property

Public Property Get PageTopMargin() As Single
    If mNewPagePending Then
        PageTopMargin = ScaleY(mTopMarginInMillimeters, vbMillimeters, mScaleMode) / mZoom * 100
    Else
        PageTopMargin = ScaleY(mPages(mCurrentPage).TopMarginInPixels, vbPixels, mScaleMode)
    End If
End Property

Public Property Let TopMargin(ByVal nValue As Single)
    Dim iChange As Boolean
    
    If nValue < MinTopMargin Then nValue = MinTopMargin
    If nValue < mNonPrintableAreaTopInMillimeters Then nValue = mNonPrintableAreaTopInMillimeters
    iChange = nValue <> mTopMarginInMillimeters
    If mDocumentCreated Then
        iChange = iChange Or (nValue <> mPages(mCurrentPage).TopMarginInMillimeters)
    End If
    
    If iChange Then
        mTopMarginInMillimeters = nValue
        If Not mCheckingMarginsLimits Then mTopMarginSet = nValue
        If mDocumentCreated Then
            mPages(mCurrentPage).TopMarginInMillimeters = mTopMarginInMillimeters
            mCurrentYPixels = mCurrentYPixels + mPages(mCurrentPage).TopMarginInPixels - mPages(mCurrentPage).NonPrintableAreaTopInPixels - mScaleTopPixels
        Else
            mCurrentYPixels = -mScaleTopPixels
        End If
    End If
End Property


Public Property Get BottomMargin() As Single ' millimeters
    BottomMargin = mBottomMarginInMillimeters
End Property

Public Property Get PageBottomMargin() As Single
    If mNewPagePending Then
        PageBottomMargin = ScaleX(mBottomMarginInMillimeters, vbMillimeters, mScaleMode) / mZoom * 100
    Else
        PageBottomMargin = ScaleX(mPages(mCurrentPage).BottomMarginInPixels, vbPixels, mScaleMode)
    End If
End Property

Public Property Let BottomMargin(ByVal nValue As Single)
    Dim iChange As Boolean
    
    If nValue < MinBottomMargin Then nValue = MinBottomMargin
    If nValue < mNonPrintableAreaBottomInMillimeters Then nValue = mNonPrintableAreaBottomInMillimeters
'    nValue = nValue * 100 / mZoom
    iChange = nValue <> mBottomMarginInMillimeters
    If mDocumentCreated Then
        iChange = iChange Or (nValue <> mPages(mCurrentPage).BottomMarginInMillimeters)
    End If
    
    If iChange Then
        mBottomMarginInMillimeters = nValue
        If Not mCheckingMarginsLimits Then mBottomMarginSet = nValue
        If mDocumentCreated Then
            mPages(mCurrentPage).BottomMarginInMillimeters = mBottomMarginInMillimeters
        End If
    End If
End Property


Private Property Get PageLeftMarginInPixels() As Single
    If mHandleMargins Then
        If (Not mDocumentCreated) Or mNewPagePending Or (mCurrentPage < 1) Then
            PageLeftMarginInPixels = (ScaleX(mLeftMarginInMillimeters, vbMillimeters, vbPixels)) / mZoom * 100
        Else
            PageLeftMarginInPixels = mPages(mCurrentPage).LeftMarginInPixels
        End If
    Else
        If (Not mDocumentCreated) Or mNewPagePending Or (mCurrentPage < 1) Then
            PageLeftMarginInPixels = (ScaleX(mNonPrintableAreaLeftInPixels, vbPixels, vbPixels)) / mZoom * 100
        Else
            PageLeftMarginInPixels = mPages(mCurrentPage).NonPrintableAreaLeftInPixels
        End If
    End If
End Property

Private Property Get PageTopMarginInPixels() As Single
    If mHandleMargins Then
        If (Not mDocumentCreated) Or mNewPagePending Or (mCurrentPage < 1) Then
            PageTopMarginInPixels = (ScaleY(mTopMarginInMillimeters, vbMillimeters, vbPixels)) / mZoom * 100
        Else
            PageTopMarginInPixels = mPages(mCurrentPage).TopMarginInPixels
        End If
    Else
        If (Not mDocumentCreated) Or mNewPagePending Or (mCurrentPage < 1) Then
            PageTopMarginInPixels = (ScaleY(mNonPrintableAreaTopInPixels, vbPixels, vbPixels)) / mZoom * 100
        Else
            PageTopMarginInPixels = mPages(mCurrentPage).NonPrintableAreaTopInPixels
        End If
    End If
End Property

'Private Sub SetPixelScale()
'    mScaleWidthPixels = ScaleX(Width - LeftMargin - RightMargin, vbMillimeters, vbPixels)
'    mScaleHeightPixels = ScaleX(Height - TopMargin - BottomMargin, vbMillimeters, vbPixels)
'End Sub

Public Sub Scale1(ByVal X1 As Variant, ByVal Y1 As Variant, ByVal X2 As Variant, ByVal Y2 As Variant)
    Dim iScaleLeftPixelsAnt As Single
    Dim iScaleTopPixelsAnt As Single
    
    If IsMissing(X1) Then X1 = 0
    If IsMissing(Y1) Then Y1 = 0
    If IsMissing(X2) Then X2 = ScaleWidth
    If IsMissing(Y2) Then Y2 = ScaleHeight
    
    mScaleWidthUser = X2 - X1
    mScaleHeightUser = Y2 - Y1
    mScaleMode = vbUser
    
    iScaleLeftPixelsAnt = mScaleLeftPixels
    If X1 <> 0 Then
        mScaleLeftPixels = ScaleX(CSng(X1), ScaleMode, vbPixels)
    Else
        mScaleLeftPixels = 0
    End If
    mCurrentXPixels = mCurrentXPixels - mScaleLeftPixels + iScaleLeftPixelsAnt
    
    iScaleTopPixelsAnt = mScaleTopPixels
    If Y1 <> 0 Then
        mScaleTopPixels = ScaleY(CSng(Y1), ScaleMode, vbPixels)
    Else
        mScaleTopPixels = 0
    End If
    mCurrentYPixels = mCurrentYPixels - mScaleTopPixels + iScaleTopPixelsAnt
    
End Sub

    
Public Property Get PageCount() As Long
    On Error Resume Next
    PageCount = UBound(mPages)
    If PageCount = -1 Then PageCount = 0
End Property

Public Property Get hDC() As Long
    'If Not mDocumentCreated Then Exit Property
    EnsureDocumentCreated
    CheckNewPagePending
    SelectFont
    hDC = mPages(mCurrentPage).MetaDC
End Property


Private Sub SelectFont()
    Dim c As Long
    Dim iFound As Boolean
    Dim iFontChanged As Boolean
    
    If Not mDocumentCreated Then Exit Sub
    CheckNewPagePending
    
    If mFontName <> mFontName_Prev Then
        iFontChanged = True
    ElseIf mFontSize <> mFontSize_Prev Then
        iFontChanged = True
    ElseIf mFontWeight <> mFontWeight_Prev Then
        iFontChanged = True
    ElseIf mFontItalic <> mFontItalic_Prev Then
        iFontChanged = True
    ElseIf mFontUnderline <> mFontUnderline_Prev Then
        iFontChanged = True
    ElseIf mFontStrikethrough <> mFontStrikethrough_Prev Then
        iFontChanged = True
    ElseIf mFontWidth <> mFontWidth_Prev Then
        iFontChanged = True
    ElseIf mFontCharset <> mFontCharset_Prev Then
        iFontChanged = True
    End If
        
    If iFontChanged Then
        mFontName_Prev = mFontName
        mFontSize_Prev = mFontSize
        mFontItalic_Prev = mFontItalic
        mFontStrikethrough_Prev = mFontStrikethrough
        mFontUnderline_Prev = mFontUnderline
        mFontWeight_Prev = mFontWeight
        mFontCharset_Prev = mFontCharset
        mFontWidth_Prev = mFontWidth
        
        iFound = False
        For c = 1 To UBound(mFontsUsed_LogFont)
            If CStr(mLogFont.lfFaceName) = CStr(mFontsUsed_LogFont(c).lfFaceName) Then
                If mLogFont.lfHeight = mFontsUsed_LogFont(c).lfHeight Then
                    If mLogFont.lfItalic = mFontsUsed_LogFont(c).lfItalic Then
                        If mLogFont.lfStrikeOut = mFontsUsed_LogFont(c).lfStrikeOut Then
                            If mLogFont.lfUnderline = mFontsUsed_LogFont(c).lfUnderline Then
                                If mLogFont.lfWeight = mFontsUsed_LogFont(c).lfWeight Then
                                    If mLogFont.lfCharSet = mFontsUsed_LogFont(c).lfCharSet Then
                                        If mLogFont.lfWidth = mFontsUsed_LogFont(c).lfWidth Then
                                            iFound = True
                                            mCurrentFontIndex = c
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next c
        
        If Not iFound Then
            ReDim Preserve mFontsUsed_LogFont(UBound(mFontsUsed_LogFont) + 1)
            mCurrentFontIndex = UBound(mFontsUsed_LogFont)
            ReDim Preserve mhFontsUsed(mCurrentFontIndex)
            ReDim Preserve mFontsUsed_HeightInPixels(mCurrentFontIndex)
            
            mFontsUsed_LogFont(mCurrentFontIndex) = mLogFont
            mhFontsUsed(mCurrentFontIndex) = CreateFontIndirectW(mLogFont)
        End If
        
        mPages(mCurrentPage).SelectFont mhFontsUsed(mCurrentFontIndex)
    End If
End Sub


Public Property Get FontBold() As Boolean
    FontBold = mFontBold
End Property

Public Property Let FontBold(ByVal nValue As Boolean)
    If nValue <> mFontBold Then
        SetFontAttribute "Bold", nValue
        mFontChanged = True
    End If
End Property


Public Property Get FontItalic() As Boolean
    FontItalic = mFontItalic
End Property

Public Property Let FontItalic(ByVal nValue As Boolean)
    If nValue <> mFontItalic Then
        SetFontAttribute "Italic", nValue
        mFontChanged = True
    End If
End Property


Public Property Get FontName() As String
    FontName = mFontName
End Property

Public Property Let FontName(ByVal nValue As String)
    If nValue <> mFontName Then
'        If Not FontExistInPrinter(nValue) Then
'            SetFontAttribute "Name", "Arial"
'        Else
            SetFontAttribute "Name", nValue
'        End If
        mFontChanged = True
    End If
End Property


Public Property Get FontSize() As Single
    FontSize = mFontSize
End Property

Public Property Let FontSize(ByVal nValue As Single)
    If nValue <> mFontSize Then
        SetFontAttribute "Size", nValue
        mFontChanged = True
    End If
End Property


Public Property Get FontStrikeThru() As Boolean
    FontStrikeThru = mFontStrikethrough
End Property

Public Property Let FontStrikeThru(ByVal nValue As Boolean)
    If nValue <> mFontStrikethrough Then
        SetFontAttribute "Strikethrough", nValue
        mFontChanged = True
    End If
End Property


Public Property Get FontUnderLine() As Boolean
    FontUnderLine = mFontUnderline
End Property

Public Property Let FontUnderLine(ByVal nValue As Boolean)
    If nValue <> mFontUnderline Then
        SetFontAttribute "Underline", nValue
        mFontChanged = True
    End If
End Property


Public Property Get FontWidth() As Integer
    FontWidth = mFontWidth
End Property

Public Property Let FontWidth(ByVal nValue As Integer)
    If nValue <> mFontWidth Then
        SetFontAttribute "Width", nValue
    End If
End Property


Public Property Get Font() As StdFont
    If (mFont Is Nothing) Or mFontChanged Then
        Set mFont = LogFontToStdFont(mLogFont)
        mFontChanged = False
    End If
    Set Font = mFont
End Property

Public Property Set Font(ByVal nFont As StdFont)
    Dim iFontName As String
    Dim iNoChange As Boolean
    
    If nFont Is Nothing Then
        Exit Property
    End If
    iFontName = nFont.Name
'    If Not FontExistInPrinter(iFontName) Then
'        iFontName = "Arial"
'    End If
    
    If mFontName = iFontName Then
        If mFontSize = nFont.Size Then
            If mFontWeight = nFont.Weight Then
                If mFontItalic = nFont.Italic Then
                    If mFontUnderline = nFont.Underline Then
                        If mFontStrikethrough = nFont.Strikethrough Then
                            If mFontCharset = nFont.Charset Then
                                If mFontWidth = 0 Then
                                    iNoChange = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    If Not iNoChange Then
        mLogFont = StdFontToLogFont_Printer(nFont, iFontName)
        AssignFontVariables
        mFontChanged = True
    End If
End Property


Private Sub Class_Initialize()
    Dim iThereIsAPrinter As Boolean
    
    iThereIsAPrinter = VB.Printers.Count > 0
    If iThereIsAPrinter Then
        Set Font = CloneFont(VB.Printer.Font)
'    mLogFont = StdFontToLogFont_Printer(VB.Printer.Font)
        AssignFontVariables
    End If
    ReDim mhFontsUsed(0)
    ReDim mFontsUsed_LogFont(0)
    ReDim mPageNumbers(0)
    
    If gPrinterExPrinted Then
        mPrinted = True
        gPrinterExPrinted = False
    End If
    
    mPrintPrevPageSetupButtonVisible = True
    mPrintPrevFormatButtonVisible = False
    mPrintPrevMinScalePercent = cPrintPreviewDefaultMinScale
    mPrintPrevMaxScalePercent = cPrintPreviewDefaultMaxScale
    mPrintPrevToolBarIconsSize = vxPPTIconsAuto
    mPrintPrevUseAltScaleIcons = False
    
    mTrackDefault = Not gPrinterExNotTrackDefault
    mScaleMode = vbTwips
    mDrawStyle = 0
    mDrawWidth = 1
    ForeColor = vbWindowText
    mFontTransparent = True
    mFillColor = 0
    mFillColorTranslated = 0
    mFillStyle = vbFSTransparent
    mDrawMode = vbCopyPen
    mOrientation = vbPRORPortrait
    mDuplex = vbPRDPSimplex
    mColorMode = vbPRCMColor
    mPrintQuality = vbPRPQPrinterDefault
    mPaperSize = vbPRPSPrinterDefault
    mPaperBin = vbPRBNAuto
    mZoom = 100
    
    If gPrinterExDeviceName = "" Then
        If iThereIsAPrinter Then
            DeviceName = VB.Printer.DeviceName
        Else
            Set mDlg = New CommonDialogExObject
            mDlg.ManualTermination = True
        End If
    Else
        DeviceName = gPrinterExDeviceName
    End If
    
    If Not IsEmpty(gPrinterExZoom) Then mZoom = gPrinterExZoom
    If Not IsEmpty(gPrinterExHandleMargins) Then mHandleMargins = gPrinterExHandleMargins Else mHandleMargins = True
    If Not IsEmpty(gPrinterExPrintPageNumbers) Then PrintPageNumbers = gPrinterExPrintPageNumbers Else PrintPageNumbers = True
    If Not IsEmpty(gPrinterExPageNumbersFont) Then
        Set PageNumbersFont = gPrinterExPageNumbersFont
    Else
        Set mPageNumbersFont = New StdFont: mPageNumbersFont.Name = "Arial": mPageNumbersFont.Bold = True: mPageNumbersFont.Size = 12
    End If
    If Not IsEmpty(gPrinterExPageNumbersForeColor) Then mPageNumbersForeColor = gPrinterExPageNumbersForeColor Else mPageNumbersForeColor = vbBlack
    If Not IsEmpty(gPrinterExPageNumbersPosition) Then PageNumbersPosition = gPrinterExPageNumbersPosition Else PageNumbersPosition = vxPositionBottomRight
    If Not IsEmpty(gPrinterExPageNumbersFormat) Then mPageNumbersFormat = gPrinterExPageNumbersFormat Else mPageNumbersFormat = "Default"
    If Not IsEmpty(gPrinterExEvents) Then
        If Not gPrinterExEvents Is Nothing Then
            Set mPrinterExEvents = gPrinterExEvents
        Else
            Set mPrinterExEvents = New PrinterExEvents
        End If
    Else
        Set mPrinterExEvents = New PrinterExEvents
    End If
    
    ' page setup and printer settings
    If Not IsEmpty(gPrinterExLeftMargin) Then LeftMargin = gPrinterExLeftMargin Else LeftMargin = cLeftMarginDefault
    If Not IsEmpty(gPrinterExTopMargin) Then TopMargin = gPrinterExTopMargin Else TopMargin = cTopMarginDefault
    If Not IsEmpty(gPrinterExRightMargin) Then RightMargin = gPrinterExRightMargin Else RightMargin = cRightMarginDefault
    If Not IsEmpty(gPrinterExBottomMargin) Then BottomMargin = gPrinterExBottomMargin Else BottomMargin = cBottomMarginDefault

    If Not IsEmpty(gPrinterExMinLeftMargin) Then MinLeftMargin = gPrinterExMinLeftMargin
    If Not IsEmpty(gPrinterExMinTopMargin) Then MinTopMargin = gPrinterExMinTopMargin
    If Not IsEmpty(gPrinterExMinRightMargin) Then MinRightMargin = gPrinterExMinRightMargin
    If Not IsEmpty(gPrinterExMinBottomMargin) Then MinBottomMargin = gPrinterExMinBottomMargin
    If Not IsEmpty(gPrinterExUnits) Then Units = gPrinterExUnits
    If Not IsEmpty(gPrinterExUnitsForUser) Then UnitsForUser = gPrinterExUnitsForUser
    
    If Not IsEmpty(gPrinterExPaperSize) Then PaperSize = gPrinterExPaperSize
    If Not IsEmpty(gPrinterExPaperBin) Then PaperBin = gPrinterExPaperBin
    If Not IsEmpty(gPrinterExPrintQuality) Then PrintQuality = gPrinterExPrintQuality
    If Not IsEmpty(gPrinterExColorMode) Then ColorMode = gPrinterExColorMode
    If Not IsEmpty(gPrinterExCollate) Then Collate = gPrinterExCollate

    SelectFont
    
End Sub

Private Sub Class_Terminate()
    Dim c As Long
    
    If PageCount > 0 Then
        If mSomethingPrintedOnDoc Then
            If gPrinterExFromPrintFn = 0 Then
                EndDoc
            End If
        End If
    End If
    
    If Not mDlg Is Nothing Then
        mDlg.DoTerminate
        mDlg.ManualTermination = False
        Set mDlg = Nothing
    End If
    
    Erase mPages ' first erase the array of cPrPrintPages objects in order to make these objects to terminate and deselect the fonts
    
    For c = 1 To UBound(mhFontsUsed)
        DeleteObject mhFontsUsed(c)
    Next c
    If mHdcReference <> 0 Then
        DeleteDC mHdcReference
    End If
    gPrinterExPrinted = mPrinted

End Sub

Public Function PrintText(nText As String, Optional ByVal nLeft, Optional ByVal nTop, Optional ByVal nWidth, Optional ByVal nHeight, Optional ByVal nAlignment As vbExPrintExAlignConstants, Optional nAddLineFeed As Boolean = True, Optional nWordWrap As Boolean = True, Optional nWordBreak As Boolean = True, Optional nRectMarginsX As Single, Optional nRectMarginsY As Single, Optional nStartsNewPage As Boolean, Optional nCalcSize As Boolean, Optional ByRef nRequiredWidth As Single, Optional ByRef TextExceedsWidth As Boolean, Optional ByVal CancelIfTextExceedsWidth As Boolean) As Single
    Dim iEP As DRAWTEXTPARAMS
    Dim iRect As RECT
    Dim iX As Single
    Dim iY As Single
    Dim iW As Single
    Dim iH As Single
    Dim iRectCalc As RECT
    Dim iTextSize As POINTAPI
    Dim iAddLineFeed As Boolean
    Dim iWWFlag As Long
    Dim iWBFlag As Long
    Dim iYAlign As Long ' 0: nTop, 1: center, 2 = bottom
    Dim iTopPrev As Long
    Dim iXMargins As Long
    Dim iYMargins As Long
    Dim iMetaDC As Long
    Dim iDTAlign As Long
    Dim iAuxText As String
    Dim iWords() As String
    Dim c As Long
    Dim N As Long
    Dim iRectCalcAux As RECT
    Dim iHeightSpaceAvailable As Long
    Dim iPos As Long
    Dim iChar As String
    Dim iStr As String
    Dim iConcat As SmartConcat
    
'    If (mDocumentCreated And (Not mNewPagePending)) <> mPageCreated Then
'        Stop
'    End If
    
    If Not mPrintingDocumentTooLongWarning Then
        If mCurrentPage > cMaxPages Then
            If Not mDocumentTooLongWarningPrinted Then
                PrintDocumentTooLongWarning
            End If
            Exit Function
        End If
    End If
    
    'Debug.Print Printer_CurrentY, ScaleY(PageTopMargin, mScaleMode, vbPixels)
    If (nText = "") And Not nAddLineFeed Then Exit Function
    EnsureDocumentCreated
    iMetaDC = mPages(mCurrentPage).MetaDC
    If iMetaDC = 0 Then Exit Function
    
    iAddLineFeed = nAddLineFeed
    
    SelectFont
        
    iEP.cbSize = Len(iEP)
    
'    If Not nCalcSize Then Stop
    If IsMissing(nLeft) Then
        If mCurrentXPixels > CurrentPageScaleWidthInPixels Then
            If Not nCalcSize Then
                mCurrentXPixels = CurrentPageScaleWidthInPixels
            End If
        End If
        iX = mCurrentXPixels
    Else
        iX = ScaleX(CSng(nLeft), mScaleMode, vbPixels) - mScaleLeftPixels + PageLeftMarginInPixels
    End If
    If IsMissing(nTop) Then
        iY = mCurrentYPixels
    Else
        iY = ScaleY(CSng(nTop), mScaleMode, vbPixels) - mScaleTopPixels + PageTopMarginInPixels
    End If
    If IsMissing(nWidth) Then
        If Not IsMissing(nLeft) Then
            iW = CurrentPageScaleWidthInPixels - iX + PageLeftMarginInPixels - mScaleLeftPixels
        Else
            iW = CurrentPageScaleWidthInPixels - mCurrentXPixels + PageLeftMarginInPixels - mScaleLeftPixels
        End If
    Else
        iW = ScaleX(CSng(nWidth), mScaleMode, vbPixels)
    End If
    If IsMissing(nHeight) Then
        If Not IsMissing(nTop) Then
            iH = CurrentPageScaleHeightInPixels - iY + PageTopMarginInPixels - mScaleTopPixels
        Else
            iH = CurrentPageScaleHeightInPixels - mCurrentYPixels + PageTopMarginInPixels - mScaleTopPixels
        End If
    Else
        iH = ScaleY(CSng(nHeight), mScaleMode, vbPixels)
    End If
    
'    flexAlignLeftTop = 0
'    flexAlignLeftCenter = 1
'    flexAlignLeftBottom = 2
'    flexAlignCenterTop = 3
'    flexAlignCenterCenter = 4
'    flexAlignCenterBottom = 5
'    flexAlignRightTop = 6
'    flexAlignRightCenter = 7
'    flexAlignRightBottom = 8
'    flexAlignGeneral = 9
    
    If nWordWrap Then
        iWWFlag = DT_WORDBREAK
        If nWordBreak Then
            iWBFlag = DT_EDITCONTROL
        End If
    End If
    
    If nAlignment = flexAlignGeneral Then ' 9
        If IsNumeric(nText) Then
            nAlignment = flexAlignRightCenter
        Else
            nAlignment = flexAlignLeftCenter
        End If
    End If
    
    Select Case nAlignment
        Case flexAlignLeftTop To flexAlignLeftBottom ' 0 to 2, nLeft
            iDTAlign = DT_LEFT
        Case flexAlignRightTop To flexAlignRightBottom  ' 6 to 8, Right$
            iDTAlign = DT_RIGHT
        Case Else ' center
            iDTAlign = DT_CENTER
    End Select
    
    Select Case nAlignment
        Case flexAlignLeftTop, flexAlignCenterTop, flexAlignRightTop
            iYAlign = 0
        Case flexAlignLeftCenter, flexAlignCenterCenter, flexAlignRightCenter
            iYAlign = 1
        Case Else ' bottom
            iYAlign = 2
    End Select
    
    iRect.Left = iX
    iRect.Top = iY
    iRect.Right = iRect.Left + iW
    
    If nRectMarginsX > 0 Then
        iXMargins = ScaleX(nRectMarginsX, mScaleMode, vbPixels)
        'If Not nCalcSize Then
            iRect.Left = iRect.Left + iXMargins
            iRect.Right = iRect.Right - iXMargins '* 2
        'End If
    End If
    If nRectMarginsY > 0 Then
        iYMargins = ScaleY(nRectMarginsY, mScaleMode, vbPixels)
        If Not nCalcSize Then
            iRect.Top = iRect.Top + iYMargins
            iRect.Bottom = iRect.Bottom - iYMargins
        End If
    End If
    
    iRectCalc.Left = iRect.Left
    iRectCalc.Top = iRect.Top
    iRectCalc.Right = iRect.Right
    If iRectCalc.Bottom < iRectCalc.Top Then iRectCalc.Bottom = iRectCalc.Top
    
    iAuxText = nText
    If Len(iAuxText) = 0 Then iAuxText = iAuxText & " "
    Call DrawTextEx(iMetaDC, StrPtr(iAuxText), Len(iAuxText), iRectCalc, iWWFlag Or DT_EXPANDTABS Or DT_TABSTOP Or iWBFlag Or iDTAlign Or DT_NOPREFIX Or DT_CALCRECT, iEP)
    If iRectCalc.Bottom = iRectCalc.Top Then
        iRectCalc.Bottom = iRectCalc.Bottom + FontHeightInternal
    End If
    If Not nCalcSize And IsMissing(nWidth) Then
        If nWordWrap Then
            If iRectCalc.Right > CurrentPageScaleWidthInPixels + PageLeftMarginInPixels - mScaleLeftPixels Then
                If mCurrentXPixels <> (-mScaleLeftPixels + PageLeftMarginInPixels) Then
                    If mRaisingEventsStartPage Then Exit Function
                    mCurrentXPixels = -mScaleLeftPixels + PageLeftMarginInPixels
                    mCurrentYPixels = mCurrentYPixels + FontHeightInternal
                    PrintText = PrintText(nText, , , nWidth, nHeight, nAlignment, nAddLineFeed, nWordWrap, nWordBreak, nRectMarginsX, nRectMarginsY, , , nRequiredWidth)
                    Exit Function
                End If
            End If
            If IsMissing(nLeft) Then
                If (iRectCalc.Bottom - iRectCalc.Top) > FontHeightInternal Then
                    If mCurrentXPixels > (-mScaleLeftPixels + PageLeftMarginInPixels) Then
                        If mRaisingEventsStartPage Then Exit Function
                        mCurrentXPixels = -mScaleLeftPixels + PageLeftMarginInPixels
                        mCurrentYPixels = mCurrentYPixels + FontHeightInternal
                        PrintText = PrintText(nText, , , nWidth, nHeight, nAlignment, nAddLineFeed, nWordWrap, nWordBreak, nRectMarginsX, nRectMarginsY, , , nRequiredWidth)
                        Exit Function
                    End If
                End If
            End If
        End If
    ElseIf (Not nCalcSize) And (Not nWordBreak) Then
        If (iRectCalc.Right - iRectCalc.Left) > iW Then
            TextExceedsWidth = True
            If CancelIfTextExceedsWidth Then
                Exit Function
            End If
        End If
    End If
    
    If Not IsMissing(nHeight) Then
        If (iRect.Top + iH - iYMargins * 2) < (CurrentPageScaleHeightInPixels - mScaleTopPixels + PageTopMarginInPixels) Then
            iRect.Bottom = iRect.Top + iH - iYMargins * 2
        Else
            iRect.Bottom = (CurrentPageScaleHeightInPixels - mScaleTopPixels + PageTopMarginInPixels) - 1
        End If
    Else
        iRect.Bottom = iRectCalc.Bottom '- iRectCalc.Top + iRect.Top
    End If
    
    nRequiredWidth = iRectCalc.Right - iRectCalc.Left
    If IsMissing(nHeight) Then
        PrintText = (iRectCalc.Bottom - iRectCalc.Top)
    Else
        PrintText = (iRect.Bottom - iRect.Top)
    End If
    
    iTopPrev = iRect.Top
    Select Case iYAlign
        Case 0 ' nTop
            ' nothing to do
        Case 1 ' center
            iRect.Top = iRect.Top + (iRect.Bottom - iRect.Top - iRectCalc.Bottom + iRectCalc.Top) / 2
        Case 2 ' bottom
            iRect.Top = iRect.Bottom - iRectCalc.Bottom + iRectCalc.Top
    End Select
    If iRect.Top < iTopPrev Then
        If nRectMarginsY > 0 Then
            iYMargins = ScaleX(nRectMarginsY, mScaleMode, vbPixels)
            iRect.Top = iRect.Top + iYMargins
            If iRect.Top > iTopPrev Then
                iRect.Top = iTopPrev
            End If
        End If
        If iRect.Top < iTopPrev Then
            iRect.Bottom = iRect.Bottom - iRect.Top + iTopPrev
            iRect.Top = iTopPrev
        End If
    End If
    If iRect.Bottom > (CurrentPageScaleHeightInPixels - mScaleTopPixels + PageTopMarginInPixels) Then
        If IsMissing(nTop) Then
            If Round(mCurrentYPixels) >= Round((mScaleTopPixels + PageTopMarginInPixels)) Then
                If (iRect.Bottom - iRect.Top) < CurrentPageScaleHeightInPixels Then ' If (iRect.Bottom - iRect.Top) is bigger than the page height, do not start a new page (only if it is smaller)
                '    If Not nCalcSize Then
                    nStartsNewPage = True
                '    End If
                End If
            End If
        End If
    End If
    If nCalcSize Then
        If mRaisingEventsStartPage Then Exit Function
        nRequiredWidth = ScaleX(nRequiredWidth, vbPixels, mScaleMode) + nRectMarginsX * 2
        If ScaleMode = vbTwips Then
            PrintText = Round(ScaleY(PrintText, vbPixels, vbTwips)) + nRectMarginsY * 2 + 1.5
        Else
            PrintText = ScaleY(PrintText, vbPixels, mScaleMode) + nRectMarginsY * 2
        End If
        Exit Function
    End If
    
    If nStartsNewPage Then
        If ((iRect.Bottom - iRect.Top) / FontHeightInternal) > 3 And (InStr(nText, " ") > 0) Then 'if tit's more than three lines we will split in between pages
            iWords = Split(Replace(nText, vbCrLf, " " & vbCrLf), " ")
            iHeightSpaceAvailable = iH ' CurrentPageScaleHeightInPixels - mCurrentYPixels + PageTopMarginInPixels
            If iHeightSpaceAvailable > (iRect.Bottom - iRect.Top) Then
                iHeightSpaceAvailable = (iRect.Bottom - iRect.Top)
            End If
            N = UBound(iWords) / (iRect.Bottom - iRect.Top) * iHeightSpaceAvailable
            If N > UBound(iWords) Then N = UBound(iWords)
            iAuxText = ""
            
            Set iConcat = New SmartConcat
            For c = 0 To N
                'iAuxText = iAuxText & iWords(c) & " "
                iConcat.AddString iWords(c)
                iConcat.AddString " "
            Next c
            iAuxText = iConcat.GenerateCurrentString
            
            iRectCalcAux.Left = iRect.Left
            iRectCalcAux.Right = iRect.Right
            iRectCalcAux.Bottom = 0
            iStr = Replace(iAuxText, " " & vbCrLf, vbCrLf)
            Call DrawTextEx(iMetaDC, StrPtr(iStr), Len(iStr), iRectCalcAux, iWWFlag Or DT_EXPANDTABS Or DT_TABSTOP Or iWBFlag Or iDTAlign Or DT_NOPREFIX Or DT_CALCRECT, iEP)
            
            If iRectCalcAux.Bottom > iHeightSpaceAvailable Then
                Do Until iRectCalcAux.Bottom < iHeightSpaceAvailable
                    iPos = InStrRev(iAuxText, " ")
                    If iPos = 0 Then
                        N = -2
                        Exit Do
                    Else
                        iAuxText = Left$(iAuxText, iPos - 1)
                        N = N - 1
                        iRectCalcAux.Left = iRect.Left
                        iRectCalcAux.Right = iRect.Right
                        iRectCalcAux.Bottom = 0
                        iStr = Replace(iAuxText, " " & vbCrLf, vbCrLf)
                        Call DrawTextEx(iMetaDC, StrPtr(iStr), Len(iStr), iRectCalcAux, iWWFlag Or DT_EXPANDTABS Or DT_TABSTOP Or iWBFlag Or iDTAlign Or DT_NOPREFIX Or DT_CALCRECT, iEP)
                    End If
                Loop
                If Right$(iAuxText, 1) = " " Then
                    iAuxText = Left$(iAuxText, Len(iAuxText) - 1)
                End If
                N = N + 1
            Else
                Do Until iRectCalcAux.Bottom > iHeightSpaceAvailable
                    If (N + 1) < UBound(iWords) Then
                        N = N + 1
                        iAuxText = iAuxText & iWords(N) & " "
                        iRectCalcAux.Left = iRect.Left
                        iRectCalcAux.Right = iRect.Right
                        iRectCalcAux.Bottom = 0
                        iStr = Replace(iAuxText, " " & vbCrLf, vbCrLf)
                        Call DrawTextEx(iMetaDC, StrPtr(iStr), Len(iStr), iRectCalcAux, iWWFlag Or DT_EXPANDTABS Or DT_TABSTOP Or iWBFlag Or iDTAlign Or DT_NOPREFIX Or DT_CALCRECT, iEP)
                    Else
                        Exit Do
                    End If
                Loop
                If Right$(iAuxText, 1) = " " Then
                    iAuxText = Left$(iAuxText, Len(iAuxText) - 1)
                End If
                iPos = InStrRev(iAuxText, " ")
                If iPos = 0 Then
                    N = -1
                Else
                    If iRectCalcAux.Bottom > iHeightSpaceAvailable Then
                        iAuxText = Left$(iAuxText, iPos - 1)
                        N = N - 1
                    End If
                End If
            End If
            If N > -1 Then
                iRect.Bottom = iRect.Top + iH - iYMargins * 2
                iStr = Replace(iAuxText, " " & vbCrLf, vbCrLf)
                Call DrawTextEx(iMetaDC, StrPtr(iStr), Len(iStr), iRect, iWWFlag Or DT_EXPANDTABS Or DT_TABSTOP Or iWBFlag Or iDTAlign Or DT_NOPREFIX, iEP)
                
                nText = ""
                If (N + 1) < UBound(iWords) Then
                    iStr = ""
                    For c = 1 To Len(iWords(N + 1))
                        iChar = Mid$(iWords(N + 1), c, 1)
                        Select Case iChar
                            Case vbCr, vbLf
                            Case Else
                                iStr = Mid$(iWords(N + 1), c)
                                Exit For
                        End Select
                    Next c
                    If Len(iStr) <> 0 Then
                        iWords(N + 1) = iStr
                    Else
                        N = N + 1
                    End If
                End If
                
                Set iConcat = New SmartConcat
                For c = N + 1 To UBound(iWords)
                    'nText = nText & iWords(c) & " "
                    iConcat.AddString iWords(c)
                    iConcat.AddString " "
                Next c
                nText = iConcat.GenerateCurrentString
                Set iConcat = Nothing
                
                If Right$(nText, 1) = " " Then
                    nText = Left$(nText, Len(nText) - 1)
                End If
            End If
        End If
        
        If mRaisingEventsStartPage Then Exit Function
        If Not NewPage Then Exit Function
        nText = Replace(nText, " " & vbCrLf, vbCrLf)
        PrintText = PrintText(nText, nLeft, , nWidth, nHeight, nAlignment, nAddLineFeed, nWordWrap, nWordBreak, nRectMarginsX, nRectMarginsY, , , nRequiredWidth)
        Exit Function
    End If
    
    If Not IsMissing(nHeight) Then
        If iAddLineFeed Then
            mCurrentYPixels = mCurrentYPixels + ScaleY(CSng(nHeight), mScaleMode, vbPixels)
        End If
    Else
        mCurrentYPixels = mCurrentYPixels + ScaleY(Round(ScaleY(PrintText, vbPixels, vbTwips)), vbTwips, vbPixels)
        If Not iAddLineFeed Then
            mCurrentYPixels = mCurrentYPixels - FontHeightInternal
        End If
    End If
    If iAddLineFeed Then
        mCurrentXPixels = -mScaleLeftPixels + PageLeftMarginInPixels
    Else
        If Not IsMissing(nWidth) Then
            mCurrentXPixels = mCurrentXPixels + ScaleX(CSng(nWidth), mScaleMode, vbPixels)
        Else
            If nWordWrap Then
                mCurrentXPixels = mCurrentXPixels + iRectCalc.Right - iRectCalc.Left
            Else
                GetTextExtentPoint32 iMetaDC, nText, Len(nText), iTextSize
                mCurrentXPixels = mCurrentXPixels + iTextSize.x
            End If
        End If
    End If
    
    Call DrawTextEx(iMetaDC, StrPtr(nText), Len(nText), iRect, iWWFlag Or DT_EXPANDTABS Or DT_TABSTOP Or iWBFlag Or iDTAlign Or DT_NOPREFIX, iEP)
    SomethingPrintedOnPage = True
    
    nRequiredWidth = ScaleX(nRequiredWidth, vbPixels, mScaleMode) + nRectMarginsX * 2
    PrintText = ScaleY(PrintText, vbPixels, mScaleMode) + nRectMarginsY * 2
End Function

Private Function FontHeightInternal() As Long
    Dim iTextSize As POINTAPI
    Dim iMetaDC As Long
    
    iMetaDC = mPages(mCurrentPage).MetaDC
    If iMetaDC = 0 Then Exit Function
    
    FontHeightInternal = mFontsUsed_HeightInPixels(mCurrentFontIndex)
    If FontHeightInternal = 0 Then
        GetTextExtentPoint32 iMetaDC, "A", Len("A"), iTextSize
        mFontsUsed_HeightInPixels(mCurrentFontIndex) = (iTextSize.y)
        FontHeightInternal = mFontsUsed_HeightInPixels(mCurrentFontIndex)
    End If
End Function

'Private Property Get FontHeight() As Long
'    EnsureDocumentCreated
'    SelectFont
'    FontHeight = FontHeightInternal
'End Property

Public Function TextWidth(Text As String) As Single
    Dim iRequiredWidth As Single
    
    PrintText Text, , , , , , False, False, False, , , , True, iRequiredWidth
    TextWidth = iRequiredWidth
End Function


Public Property Get TwipsPerPixelX() As Single
    EnsureDocumentCreated
    TwipsPerPixelX = 1440 / mPages(mCurrentPage).MetaDPIX
End Property

Public Property Get TwipsPerPixelY() As Single
    EnsureDocumentCreated
    TwipsPerPixelY = 1440 / mPages(mCurrentPage).MetaDPIY
End Property

Public Function ScaleX(Width, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
    Dim iDPIX As Long
    
    If IsMissing(FromScale) Then FromScale = vbHimetric
    If IsMissing(ToScale) Then ToScale = vbTwips
    
    If Width = 0 Then
        ScaleX = 0
        Exit Function
    End If
    
'    If (FromScale = vbPixels) Or (ToScale = vbPixels) Then
 '       EnsureDocumentCreated
  '  End If
    
    If ScaleMode <> vbUser Then
        If (FromScale = vbUser) Or (ToScale = vbUser) Then
            RaiseError 5, "PrinterEx": Exit Function
        End If
    End If
    
    If FromScale = vbPixels Or ToScale = vbPixels Then
        If (Not mDocumentCreated) Or mNewPagePending Or (mCurrentPage < 1) Then
            iDPIX = mDPIX
        Else
            iDPIX = mPages(mCurrentPage).MetaDPIX
        End If
    End If
    
    Select Case FromScale
        Case vbCentimeters
            Select Case ToScale
                Case vbCentimeters
                    ScaleX = Width
                Case vbCharacters
                    ScaleX = Width * 1440 / 2.54 / 120
                Case vbHimetric
                    ScaleX = Width * 1000
                Case vbInches
                    ScaleX = Width / 2.54
                Case vbMillimeters
                    ScaleX = Width * 10
                Case vbPixels
                    ScaleX = Width * iDPIX / 2.54
                Case vbPoints
                    ScaleX = Width * 72 / 2.54
                Case vbTwips
                    ScaleX = Width * 1440 / 2.54
                Case vbUser
                    ScaleX = Width / CurrentPageScaleWidthInPixels * mScaleWidthUser * iDPIX / 2.54
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case vbCharacters
            Select Case ToScale
                Case vbCentimeters
                    ScaleX = Width / 1440 * 2.54 * 120
                Case vbCharacters
                    ScaleX = Width
                Case vbHimetric
                    ScaleX = Width * 2540 / 1440 * 120
                Case vbInches
                    ScaleX = Width / 1440 * 120
                Case vbMillimeters
                    ScaleX = Width / 1440 * 120 * 25.4
                Case vbPixels
                    ScaleX = Width / 1440 * iDPIX * 120
                Case vbPoints
                    ScaleX = Width / 20 * 120
                Case vbTwips
                    ScaleX = Width * 120
                Case vbUser
                    ScaleX = Width / CurrentPageScaleWidthInPixels * mScaleWidthUser / 1440 * iDPIX * 120
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case vbHimetric
            Select Case ToScale
                Case vbCentimeters
                    ScaleX = Width / 1000
                Case vbCharacters
                    ScaleX = Width / 2540 * 1440 / 120
                Case vbHimetric
                    ScaleX = Width
                Case vbInches
                    ScaleX = Width / 2540
                Case vbMillimeters
                    ScaleX = Width / 100
                Case vbPixels
                    ScaleX = Width * iDPIX / 2540
                Case vbPoints
                    ScaleX = Width / 2540 * 72
                Case vbTwips
                    ScaleX = Width / 2540 * 1440
                Case vbUser
                    ScaleX = Width / CurrentPageScaleWidthInPixels * mScaleWidthUser * iDPIX / 2540
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case vbInches
            Select Case ToScale
                Case vbCentimeters
                    ScaleX = Width * 2.54
                Case vbCharacters
                    ScaleX = Width * 1440 / 120
                Case vbHimetric
                    ScaleX = Width * 2540
                Case vbInches
                    ScaleX = Width
                Case vbMillimeters
                    ScaleX = Width * 25.4
                Case vbPixels
                    ScaleX = Width * iDPIX
                Case vbPoints
                    ScaleX = Width * 72
                Case vbTwips
                    ScaleX = Width * 1440
                Case vbUser
                    ScaleX = Width / CurrentPageScaleWidthInPixels * mScaleWidthUser * iDPIX
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case vbMillimeters
            Select Case ToScale
                Case vbCentimeters
                    ScaleX = Width / 10
                Case vbCharacters
                    ScaleX = Width * 1440 / 120 / 25.4
                Case vbHimetric
                    ScaleX = Width / 100
                Case vbInches
                    ScaleX = Width / 25.4
                Case vbMillimeters
                    ScaleX = Width
                Case vbPixels
                    ScaleX = Width * iDPIX / 25.4
                Case vbPoints
                    ScaleX = Width * 72 / 25.4
                Case vbTwips
                    ScaleX = Width * 1440 / 25.4
                Case vbUser
                    ScaleX = Width / CurrentPageScaleWidthInPixels * mScaleWidthUser * iDPIX / 25.4
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case vbPixels
            Select Case ToScale
                Case vbCentimeters
                    ScaleX = Width / iDPIX * 2.54
                Case vbCharacters
                    ScaleX = Width * 1440 / iDPIX / 120
                Case vbHimetric
                    ScaleX = Width / iDPIX * 2540
                Case vbInches
                    ScaleX = Width / iDPIX
                Case vbMillimeters
                    ScaleX = Width / iDPIX * 25.4
                Case vbPixels
                    ScaleX = Width
                Case vbPoints
                    ScaleX = Width * 1440 / iDPIX / 20
                Case vbTwips
                    ScaleX = Width * 1440 / iDPIX
                Case vbUser
                    ScaleX = Width / CurrentPageScaleWidthInPixels * mScaleWidthUser
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case vbPoints
            Select Case ToScale
                Case vbCentimeters
                    ScaleX = Width / 72 * 2.54
                Case vbCharacters
                    ScaleX = Width * 20 / 120
                Case vbHimetric
                    ScaleX = Width * 2540 / 72
                Case vbInches
                    ScaleX = Width / 72
                Case vbMillimeters
                    ScaleX = Width / 72 * 25.4
                Case vbPixels
                    ScaleX = Width / 1440 * iDPIX * 20
                Case vbPoints
                    ScaleX = Width
                Case vbTwips
                    ScaleX = Width * 20
                Case vbUser
                    ScaleX = Width / CurrentPageScaleWidthInPixels * mScaleWidthUser * iDPIX / 72
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case vbTwips
            Select Case ToScale
                Case vbCentimeters
                    ScaleX = Width / 1440 * 2.54
                Case vbCharacters
                    ScaleX = Width / 120
                Case vbHimetric
                    ScaleX = Width * 2540 / 1440
                Case vbInches
                    ScaleX = Width / 1440
                Case vbMillimeters
                    ScaleX = Width / 1440 * 25.4
                Case vbPixels
                    ScaleX = Width / 1440 * iDPIX
                Case vbPoints
                    ScaleX = Width / 20
                Case vbTwips
                    ScaleX = Width
                Case vbUser
                    ScaleX = Width / CurrentPageScaleWidthInPixels * mScaleWidthUser / 1440 * iDPIX
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case vbUser
            Select Case ToScale
                Case vbCentimeters
                    ScaleX = Width * CurrentPageScaleWidthInPixels / mScaleWidthUser / iDPIX * 2.54
                Case vbCharacters
                    ScaleX = Width * CurrentPageScaleWidthInPixels / mScaleWidthUser * 1440 / iDPIX / 120
                Case vbHimetric
                    ScaleX = Width * CurrentPageScaleWidthInPixels / mScaleWidthUser / iDPIX * 2540
                Case vbInches
                    ScaleX = Width * CurrentPageScaleWidthInPixels / mScaleWidthUser / iDPIX
                Case vbMillimeters
                    ScaleX = Width * CurrentPageScaleWidthInPixels / mScaleWidthUser / iDPIX * 25.4
                Case vbPixels
                    ScaleX = Width * CurrentPageScaleWidthInPixels / mScaleWidthUser
                Case vbPoints
                    ScaleX = Width * CurrentPageScaleWidthInPixels / mScaleWidthUser / iDPIX * 72
                Case vbTwips
                    ScaleX = Width * CurrentPageScaleWidthInPixels / mScaleWidthUser * 1440 / iDPIX
                Case vbUser
                    ScaleX = Width
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case Else
            RaiseError 5, "PrinterEx": Exit Function
    End Select
End Function

Public Function ScaleY(Height, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
    Dim iDPIY As Long
    
    If IsMissing(FromScale) Then FromScale = vbHimetric
    If IsMissing(ToScale) Then ToScale = vbTwips
    
    If Height = 0 Then
        ScaleY = 0
        Exit Function
    End If
    
'    If (FromScale = vbPixels) Or (ToScale = vbPixels) Then
 '       EnsureDocumentCreated
  '  End If
    
    If ScaleMode <> vbUser Then
        If (FromScale = vbUser) Or (ToScale = vbUser) Then
            RaiseError 5, "PrinterEx": Exit Function
        End If
    End If
    
    If FromScale = vbPixels Or ToScale = vbPixels Then
        If (Not mDocumentCreated) Or mNewPagePending Or (mCurrentPage < 1) Then
            iDPIY = mDPIY
        Else
            iDPIY = mPages(mCurrentPage).MetaDPIY
        End If
    End If
    
    Select Case FromScale
        Case vbCentimeters
            Select Case ToScale
                Case vbCentimeters
                    ScaleY = Height
                Case vbCharacters
                    ScaleY = Height * 1440 / 2.54 / 240
                Case vbHimetric
                    ScaleY = Height * 1000
                Case vbInches
                    ScaleY = Height / 2.54
                Case vbMillimeters
                    ScaleY = Height * 10
                Case vbPixels
                    ScaleY = Height * iDPIY / 2.54
                Case vbPoints
                    ScaleY = Height * 72 / 2.54
                Case vbTwips
                    ScaleY = Height * 1440 / 2.54
                Case vbUser
                    ScaleY = Height / CurrentPageScaleHeightInPixels * mScaleHeightUser * iDPIY / 2.54
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case vbCharacters
            Select Case ToScale
                Case vbCentimeters
                    ScaleY = Height / 1440 * 2.54 * 240
                Case vbCharacters
                    ScaleY = Height
                Case vbHimetric
                    ScaleY = Height * 2540 / 1440 * 240
                Case vbInches
                    ScaleY = Height / 1440 * 240
                Case vbMillimeters
                    ScaleY = Height / 1440 * 240 * 25.4
                Case vbPixels
                    ScaleY = Height / 1440 * iDPIY * 240
                Case vbPoints
                    ScaleY = Height / 20 * 240
                Case vbTwips
                    ScaleY = Height * 240
                Case vbUser
                    ScaleY = Height / CurrentPageScaleHeightInPixels * mScaleHeightUser / 1440 * iDPIY * 240
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case vbHimetric
            Select Case ToScale
                Case vbCentimeters
                    ScaleY = Height / 1000
                Case vbCharacters
                    ScaleY = Height / 2540 * 1440 / 240
                Case vbHimetric
                    ScaleY = Height
                Case vbInches
                    ScaleY = Height / 2540
                Case vbMillimeters
                    ScaleY = Height / 100
                Case vbPixels
                    ScaleY = Height * iDPIY / 2540
                Case vbPoints
                    ScaleY = Height / 2540 * 72
                Case vbTwips
                    ScaleY = Height / 2540 * 1440
                Case vbUser
                    ScaleY = Height / CurrentPageScaleHeightInPixels * mScaleHeightUser * iDPIY / 2540
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case vbInches
            Select Case ToScale
                Case vbCentimeters
                    ScaleY = Height * 2.54
                Case vbCharacters
                    ScaleY = Height * 1440 / 240
                Case vbHimetric
                    ScaleY = Height * 2540
                Case vbInches
                    ScaleY = Height
                Case vbMillimeters
                    ScaleY = Height * 25.4
                Case vbPixels
                    ScaleY = Height * iDPIY
                Case vbPoints
                    ScaleY = Height * 72
                Case vbTwips
                    ScaleY = Height * 1440
                Case vbUser
                    ScaleY = Height / CurrentPageScaleHeightInPixels * mScaleHeightUser * iDPIY
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case vbMillimeters
            Select Case ToScale
                Case vbCentimeters
                    ScaleY = Height / 10
                Case vbCharacters
                    ScaleY = Height * 1440 / 240 / 25.4
                Case vbHimetric
                    ScaleY = Height / 100
                Case vbInches
                    ScaleY = Height / 25.4
                Case vbMillimeters
                    ScaleY = Height
                Case vbPixels
                    ScaleY = Height * iDPIY / 25.4
                Case vbPoints
                    ScaleY = Height * 72 / 25.4
                Case vbTwips
                    ScaleY = Height * 1440 / 25.4
                Case vbUser
                    ScaleY = Height / CurrentPageScaleHeightInPixels * mScaleHeightUser * iDPIY / 25.4
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case vbPixels
            Select Case ToScale
                Case vbCentimeters
                    ScaleY = Height / iDPIY * 2.54
                Case vbCharacters
                    ScaleY = Height * 1440 / iDPIY / 240
                Case vbHimetric
                    ScaleY = Height / iDPIY * 2540
                Case vbInches
                    ScaleY = Height / iDPIY
                Case vbMillimeters
                    ScaleY = Height / iDPIY * 25.4
                Case vbPixels
                    ScaleY = Height
                Case vbPoints
                    ScaleY = Height * 1440 / iDPIY / 20
                Case vbTwips
                    ScaleY = Height * 1440 / iDPIY
                Case vbUser
                    ScaleY = Height / CurrentPageScaleHeightInPixels * mScaleHeightUser
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case vbPoints
            Select Case ToScale
                Case vbCentimeters
                    ScaleY = Height / 72 * 2.54
                Case vbCharacters
                    ScaleY = Height * 20 / 240
                Case vbHimetric
                    ScaleY = Height * 2540 / 72
                Case vbInches
                    ScaleY = Height / 72
                Case vbMillimeters
                    ScaleY = Height / 72 * 25.4
                Case vbPixels
                    ScaleY = Height / 1440 * iDPIY * 20
                Case vbPoints
                    ScaleY = Height
                Case vbTwips
                    ScaleY = Height * 20
                Case vbUser
                    ScaleY = Height / CurrentPageScaleHeightInPixels * mScaleHeightUser * iDPIY / 72
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case vbTwips
            Select Case ToScale
                Case vbCentimeters
                    ScaleY = Height / 1440 * 2.54
                Case vbCharacters
                    ScaleY = Height / 240
                Case vbHimetric
                    ScaleY = Height * 2540 / 1440
                Case vbInches
                    ScaleY = Height / 1440
                Case vbMillimeters
                    ScaleY = Height / 1440 * 25.4
                Case vbPixels
                    ScaleY = Height / 1440 * iDPIY
                Case vbPoints
                    ScaleY = Height / 20
                Case vbTwips
                    ScaleY = Height
                Case vbUser
                    ScaleY = Height / CurrentPageScaleHeightInPixels * mScaleHeightUser / 1440 * iDPIY
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case vbUser
            Select Case ToScale
                Case vbCentimeters
                    ScaleY = Height * CurrentPageScaleHeightInPixels / mScaleHeightUser / iDPIY * 2.54
                Case vbCharacters
                    ScaleY = Height * CurrentPageScaleHeightInPixels / mScaleHeightUser * 1440 / iDPIY / 240
                Case vbHimetric
                    ScaleY = Height * CurrentPageScaleHeightInPixels / mScaleHeightUser / iDPIY * 2540
                Case vbInches
                    ScaleY = Height * CurrentPageScaleHeightInPixels / mScaleHeightUser / iDPIY
                Case vbMillimeters
                    ScaleY = Height * CurrentPageScaleHeightInPixels / mScaleHeightUser / iDPIY * 25.4
                Case vbPixels
                    ScaleY = Height * CurrentPageScaleHeightInPixels / mScaleHeightUser
                Case vbPoints
                    ScaleY = Height * CurrentPageScaleHeightInPixels / mScaleHeightUser / iDPIY * 72
                Case vbTwips
                    ScaleY = Height * CurrentPageScaleHeightInPixels / mScaleHeightUser * 1440 / iDPIY
                Case vbUser
                    ScaleY = Height
                Case Else
                    RaiseError 5, "PrinterEx": Exit Function
            End Select
        Case Else
            RaiseError 5, "PrinterEx": Exit Function
    End Select
End Function

Private Sub EnsureDocumentCreated()
    If Not mDocumentCreated Then CreateDocument
    CheckNewPagePending
End Sub

Private Sub EnsureDocumentCreated2()
    If Not mDocumentCreated Then CreateDocument
End Sub

Public Property Get HdcReference() As Long
    Dim iDlgAux As New CommonDialogExObject
    
    If mHdcReference <> 0 Then
        HdcReference = mHdcReference
    Else
        If Printers.Count = 0 Then
            HdcReference = CreateCompatibleDC(0)
        Else
            iDlgAux.ShowPrinter cdePDReturnIC Or cdePDReturnDefault
            mHdcReference = CreateCompatibleDC(iDlgAux.hDC) ' this should be updated with the ResetDC API if the PrintQuality, PaperSize or DeviceName change (*), or create a new reference DC if it can't be updated (I don't know if the IC can be updated after being used as a reference -to create the metafiles-), also doing so it's necessary to change the way CreatePage in the cPrinterExPage class module sets the paper dimensions. It's not implemented at this time. (*) DeviceName can't be updated with ResetDC, in that case it's necessary to create a new DC with the defaults of that printer with the common dialog API setting the device name in the DEVNAMES (that's not implemented in the CommonDialogExObject class module at this time -to be able to set the DeviceName property before calling ShowPrinter with cdePDReturnIC Or cdePDReturnDefault-).
            HdcReference = mHdcReference
        End If
    End If
End Property

Private Sub SetPen()
    If Not mDocumentCreated Then Exit Sub
    CheckNewPagePending
    mPages(mCurrentPage).SetPen mDrawStyle, mDrawWidth, mForeColorTranslated, mDrawMode
End Sub

Private Sub SetBrush()
    If Not mDocumentCreated Then Exit Sub
    CheckNewPagePending
    mPages(mCurrentPage).SetBrush mFillStyle, mFillColorTranslated
End Sub

Private Sub SetBackMode()
    If Not mDocumentCreated Then Exit Sub
    CheckNewPagePending
    mPages(mCurrentPage).FontTransparent = mFontTransparent
End Sub

Public Sub DrawLine(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, Optional Color, Optional B As Boolean, Optional f As Boolean)
    Dim iMetaDC As Long
    Dim iPenPrev As Long
    Dim iPen As Long
    Dim iColor As Long
    Dim iPoints(4) As POINTAPI
    Dim iLb As LOGBRUSH
    Dim iBrushPrev As Long
    Dim iBrush As Long
    Dim iFT As Boolean
    Dim iX1 As Single
    Dim iY1 As Single
    Dim iX2 As Single
    Dim iY2 As Single
    
    EnsureDocumentCreated
    'Debug.Print Printer_CurrentY, ScaleY(PageTopMargin, mScaleMode, vbPixels)
    iMetaDC = mPages(mCurrentPage).MetaDC
    If iMetaDC = 0 Then Exit Sub
    If DocumentTooLong Then Exit Sub
    
    iX1 = ScaleX(X1, mScaleMode, vbPixels) - mScaleLeftPixels + PageLeftMarginInPixels
    iY1 = ScaleY(Y1, mScaleMode, vbPixels) - mScaleTopPixels + PageTopMarginInPixels
    iX2 = ScaleX(X2, mScaleMode, vbPixels) - mScaleLeftPixels + PageLeftMarginInPixels
    iY2 = ScaleY(Y2, mScaleMode, vbPixels) - mScaleTopPixels + PageTopMarginInPixels
    
    If Not IsMissing(Color) Then
        TranslateColor CLng(Color), 0, iColor
        If iColor <> mForeColorTranslated Then
            iPen = CreatePen(mDrawStyle, mDrawWidth, iColor)
            iPenPrev = SelectObject(iMetaDC, iPen)
        End If
    Else
        iColor = mForeColorTranslated
    End If
    
    If Not B Then
        ' draw line
        iPoints(0).x = iX1
        iPoints(0).y = iY1
        iPoints(1).x = iX2
        iPoints(1).y = iY2
        Polyline iMetaDC, iPoints(0), 2
    Else
        ' draw rectangle
        iPoints(0).x = iX1
        iPoints(0).y = iY1
        iPoints(1).x = iX2
        iPoints(1).y = iY1
        iPoints(2).x = iX2
        iPoints(2).y = iY2
        iPoints(3).x = iX1
        iPoints(3).y = iY2
        iPoints(4).x = iX1
        iPoints(4).y = iY1
        If f Then
            If (iColor <> mFillColorTranslated) Or (mFillStyle <> vbFSSolid) Then
                iLb.lbStyle = BS_SOLID ' 0
                iLb.lbColor = iColor
                
                iBrush = CreateBrushIndirect(iLb)
                iBrushPrev = SelectObject(iMetaDC, iBrush)
            End If
            
            Polygon iMetaDC, iPoints(0), 5
            
        Else
            If mFillStyle <> vbFSTransparent Then
                If mFillStyle <> vbFSSolid Then
                    If Not mFontTransparent Then
                        mFontTransparent = True
                        SetBackMode
                        iFT = True
                    End If
                End If
                Polygon iMetaDC, iPoints(0), 5
                If iFT Then
                    mFontTransparent = False
                    SetBackMode
                End If
            End If
        End If
        Polyline iMetaDC, iPoints(0), 5
    End If
    
    If iPenPrev <> 0 Then
        Call SelectObject(iMetaDC, iPenPrev)
    End If
    If iPen <> 0 Then
        DeleteObject iPen
    End If
    If iBrushPrev <> 0 Then
        Call SelectObject(iMetaDC, iBrushPrev)
    End If
    If iBrush <> 0 Then
        DeleteObject iBrush
    End If
    mCurrentXPixels = iX2
    mCurrentYPixels = iY2
    SomethingPrintedOnPage = True
End Sub


Public Property Let Width(nValue As Long)
    If mDocumentCreated And (Not mNewPagePending) Then
        If mPages(mCurrentPage).SomethingPrintedOnPage Then
            RaiseError 396, "PrinterEx", "Width property cannot be set within a page"
            Exit Property
        Else
            KillLastPage
        End If
    End If
    If nValue = 0 Then
        RaiseError 380, "PrinterEx"
        Exit Property
    End If
    
    PaperSize = vbPRPSUser
    
    If mOrientation = vbPRORPortrait Then
        mPaperWidthInMillimeters = ScaleX(nValue, vbTwips, vbMillimeters)
    Else
        mPaperHeightInMillimeters = ScaleX(nValue, vbTwips, vbMillimeters)
    End If
End Property

Public Property Get Width() As Long
    Dim iPw As Single
    
    If mDocumentCreated Then
        EnsureDocumentCreated
        Width = Round(ScaleX(mPages(mCurrentPage).WidthInMillimeters, vbMillimeters, vbTwips))
    Else
        If mOrientation = vbPRORPortrait Then
            iPw = mPaperWidthInMillimeters
        Else
            iPw = mPaperHeightInMillimeters
        End If
        If iPw <> 0 Then
            Width = ScaleX(iPw, vbMillimeters, vbTwips)
        Else
            Width = ScaleX(GetDeviceCaps(HdcReference, HORZRES) / GetDeviceCaps(HdcReference, LOGPIXELSX), vbInches, vbTwips)
        End If
    End If
End Property


Public Property Let Height(nValue As Long)
    If mDocumentCreated And (Not mNewPagePending) Then
        If mPages(mCurrentPage).SomethingPrintedOnPage Then
            RaiseError 396, "PrinterEx", "Height property cannot be set within a page"
            Exit Property
        Else
            KillLastPage
        End If
    End If
    If nValue = 0 Then
        RaiseError 380, "PrinterEx"
        Exit Property
    End If
    
    PaperSize = vbPRPSUser
    
    If mOrientation = vbPRORPortrait Then
        mPaperHeightInMillimeters = ScaleY(nValue, vbTwips, vbMillimeters)
    Else
        mPaperWidthInMillimeters = ScaleY(nValue, vbTwips, vbMillimeters)
    End If
End Property


Public Property Get Height() As Long
    Dim iPh As Single
    
    If mDocumentCreated Then
        EnsureDocumentCreated
        Height = Round(ScaleY(mPages(mCurrentPage).HeightInMillimeters, vbMillimeters, vbTwips))
    Else
        If mOrientation = vbPRORPortrait Then
            iPh = mPaperHeightInMillimeters
        Else
            iPh = mPaperWidthInMillimeters
        End If
        If iPh <> 0 Then
            Height = ScaleY(iPh, vbMillimeters, vbTwips)
        Else
            Height = ScaleY(GetDeviceCaps(HdcReference, VERTRES) / GetDeviceCaps(HdcReference, LOGPIXELSY), vbInches, vbTwips)
        End If
    End If
End Property


' IPrinterEx Interface

' Kept for keeping binary compatibility, it will be removed in the next version
Private Property Let IPrinterEx_CommonDialogFlags(RHS As Long)
    
End Property

Private Property Get IPrinterEx_CommonDialogFlags() As Long
    
End Property


Private Property Let IPrinterEx_DocKey(RHS As String)
    DocKey = RHS
End Property

Private Property Get IPrinterEx_DocKey() As String
    IPrinterEx_DocKey = DocKey
End Property

Private Property Let IPrinterEx_HandleMargins(ByVal RHS As Boolean)
    HandleMargins = RHS
    gPrinterExHandleMargins = mHandleMargins
End Property

Private Property Get IPrinterEx_HandleMargins() As Boolean
    IPrinterEx_HandleMargins = HandleMargins
End Property


Private Property Get IPrinterEx_PageHeight() As Single
    IPrinterEx_PageHeight = PageHeight
End Property


Private Property Get IPrinterEx_PageNonPrintableAreaBottom() As Single
    IPrinterEx_PageNonPrintableAreaBottom = PageNonPrintableAreaBottom
End Property


Private Property Get IPrinterEx_PageNonPrintableAreaLeft() As Single
    IPrinterEx_PageNonPrintableAreaLeft = PageNonPrintableAreaLeft
End Property


Private Property Get IPrinterEx_PageNonPrintableAreaRight() As Single
    IPrinterEx_PageNonPrintableAreaRight = PageNonPrintableAreaRight
End Property


Private Property Get IPrinterEx_PageNonPrintableAreaTop() As Single
    IPrinterEx_PageNonPrintableAreaTop = PageNonPrintableAreaTop
End Property


Private Property Let IPrinterEx_PageSetupFlags(RHS As cdeCommonDialogExPageSetupFlagsConstants)
    If RHS <> mPageSetupFlags Then
        If mPageSetupMarginsDisabledAutomatically Then
            If (RHS And cdePSDisableMargins) <> 0 Then
                mPageSetupMarginsDisabledAutomatically = False
            End If
        End If
        mPageSetupFlags = RHS
        If Not mPrintFnObject Is Nothing Then
            mPrintFnObject.PageSetupFlags = mPageSetupFlags
        End If
    End If
End Property

Private Property Get IPrinterEx_PageSetupFlags() As cdeCommonDialogExPageSetupFlagsConstants
    IPrinterEx_PageSetupFlags = mPageSetupFlags
End Property


Private Property Let IPrinterEx_PrinterFlags(RHS As cdeCommonDialogExPrinterFlagsConstants)
    If RHS <> mPrinterFlags Then
        mPrinterFlags = RHS
        If Not mPrintFnObject Is Nothing Then
            mPrintFnObject.PrinterFlags = mPrinterFlags
        End If
    End If
End Property

Private Property Get IPrinterEx_PrinterFlags() As cdeCommonDialogExPrinterFlagsConstants
    IPrinterEx_PrinterFlags = mPrinterFlags
End Property


Private Property Set IPrinterEx_PrintPrevFormatButtonPicture(nSizeIdentifier As VBExToobarDAButtonIconSizeConstants, RHS As stdole.Picture)
    If (nSizeIdentifier < vxIZ16) Or (nSizeIdentifier > vxIZ36) Then
        RaiseError 381, "PrinterEx" ' Invalid property array index
        Exit Property
    End If
    Set mPrintPrevFormatButtonPicture(nSizeIdentifier) = RHS
End Property

Private Property Get IPrinterEx_PrintPrevFormatButtonPicture(nSizeIdentifier As VBExToobarDAButtonIconSizeConstants) As stdole.Picture
    Dim iFrm As frmPrintPreview
    
    If (nSizeIdentifier < vxIZ16) Or (nSizeIdentifier > vxIZ36) Then
        RaiseError 381, "PrinterEx" ' Invalid property array index
        Exit Property
    End If
    
    If mPrintPrevFormatButtonPicture(nSizeIdentifier) Is Nothing Then
        Set iFrm = New frmPrintPreview
        Set IPrinterEx_PrintPrevFormatButtonPicture = iFrm.FormatButtonPicture(nSizeIdentifier)
        Unload iFrm
        Set iFrm = Nothing
    Else
        Set IPrinterEx_PrintPrevFormatButtonPicture = mPrintPrevFormatButtonPicture(nSizeIdentifier)
    End If
End Property


Private Property Let IPrinterEx_PrintPrevFormatButtonToolTipText(RHS As String)
    mPrintPrevFormatButtonToolTipText = RHS
End Property

Private Property Get IPrinterEx_PrintPrevFormatButtonToolTipText() As String
    IPrinterEx_PrintPrevFormatButtonToolTipText = mPrintPrevFormatButtonToolTipText
End Property


Private Property Let IPrinterEx_PrintPrevMaxScalePercent(RHS As Long)
    mPrintPrevMinScalePercent = RHS
End Property

Private Property Get IPrinterEx_PrintPrevMaxScalePercent() As Long
    IPrinterEx_PrintPrevMaxScalePercent = mPrintPrevMinScalePercent
End Property


Private Property Let IPrinterEx_PrintPrevMinScalePercent(RHS As Long)
    mPrintPrevMinScalePercent = RHS
End Property

Private Property Get IPrinterEx_PrintPrevMinScalePercent() As Long
    IPrinterEx_PrintPrevMinScalePercent = mPrintPrevMinScalePercent
End Property


Private Property Let IPrinterEx_PrintPrevFormatButtonVisible(RHS As Boolean)
    mPrintPrevFormatButtonVisible = RHS
    If Not mPrintFnObject Is Nothing Then
        mPrintFnObject.FormatButtonVisible = mPrintPrevFormatButtonVisible
    End If
End Property

Private Property Get IPrinterEx_PrintPrevFormatButtonVisible() As Boolean
    IPrinterEx_PrintPrevFormatButtonVisible = mPrintPrevFormatButtonVisible
End Property


Private Property Let IPrinterEx_PrintPrevPageSetupButtonVisible(RHS As Boolean)
    mPrintPrevPageSetupButtonVisible = RHS
    If Not mPrintFnObject Is Nothing Then
        mPrintFnObject.PageSetupButtonVisible = mPrintPrevPageSetupButtonVisible
    End If
End Property

Private Property Get IPrinterEx_PrintPrevPageSetupButtonVisible() As Boolean
    IPrinterEx_PrintPrevPageSetupButtonVisible = mPrintPrevPageSetupButtonVisible
End Property



Private Property Get IPrinterEx_PageWidth() As Single
    IPrinterEx_PageWidth = PageWidth
End Property


Private Property Let IPrinterEx_PrintPageNumbers(ByVal RHS As Boolean)
    If mLoaded_PrintPageNumbers Then
        mOriginal_PrintPageNumbers = RHS
    Else
        PrintPageNumbers = RHS
    End If
    gPrinterExPrintPageNumbers = RHS
End Property

Private Property Get IPrinterEx_PrintPageNumbers() As Boolean
    IPrinterEx_PrintPageNumbers = PrintPageNumbers
End Property


Private Property Get IPrinterEx_Canceled() As Boolean
    IPrinterEx_Canceled = mCanceled
End Property


Private Property Let IPrinterEx_Collate(ByVal RHS As Boolean)
    If mLoaded_Collate Then
        mOriginal_Collate = RHS
    Else
        Collate = RHS
    End If
End Property

Private Property Get IPrinterEx_Collate() As Boolean
    IPrinterEx_Collate = Collate
End Property


Private Property Get IPrinterEx_CurrentPageNumber() As Long
    IPrinterEx_CurrentPageNumber = CurrentPageNumber
End Property


Private Property Let IPrinterEx_DocumentName(ByVal RHS As String)
    DocumentName = RHS
End Property

Private Property Get IPrinterEx_DocumentName() As String
    IPrinterEx_DocumentName = DocumentName
End Property


Private Property Get IPrinterEx_Events() As Object
    Set IPrinterEx_Events = Events
End Property

Private Property Let IPrinterEx_FontWidth(ByVal RHS As Integer)
    FontWidth = RHS
End Property

Private Property Get IPrinterEx_FontWidth() As Integer
    IPrinterEx_FontWidth = FontWidth
End Property


Private Property Let IPrinterEx_FromPage(ByVal RHS As Long)
    FromPage = RHS
End Property

Private Property Get IPrinterEx_FromPage() As Long
    IPrinterEx_FromPage = FromPage
End Property


Private Function IPrinterEx_GetDocFirstPageIndex(Optional ByVal nPageIndex As Long) As Long
    IPrinterEx_GetDocFirstPageIndex = (GetDocFirstPageIndex(nPageIndex))
End Function


Private Function IPrinterEx_GetDocLastPageIndex(Optional ByVal nPageIndex As Long) As Long
    IPrinterEx_GetDocLastPageIndex = GetDocLastPageIndex(nPageIndex)
End Function


Private Function IPrinterEx_GetDocPageCount(Optional ByVal nPageIndex As Long) As Long
    If nPageIndex < 0 Then Exit Function
    If nPageIndex = 0 Then nPageIndex = mCurrentPage
    IPrinterEx_GetDocPageCount = GetDocPageCount
End Function


Private Function IPrinterEx_GetPageNumber(ByVal nPageIndex As Long) As Long
    IPrinterEx_GetPageNumber = GetPageNumber(nPageIndex)
End Function


Private Sub IPrinterEx_GoToPage(ByVal nPageNumber As Long)
    GoToPage nPageNumber
End Sub


Private Property Set IPrinterEx_PageNumbersFont(ByVal RHS As stdole.Font)
    If mLoaded_PageNumbersFont Then
        Set mOriginal_PageNumbersFont = RHS
    Else
        Set PageNumbersFont = RHS
    End If
    Set gPrinterExPageNumbersFont = RHS
End Property

Private Property Get IPrinterEx_PageNumbersFont() As stdole.Font
    Set IPrinterEx_PageNumbersFont = PageNumbersFont
End Property


Private Property Let IPrinterEx_PageNumbersForeColor(ByVal RHS As OLE_COLOR)
    If mLoaded_PageNumbersForeColor Then
        mOriginal_PageNumbersForeColor = RHS
    Else
        PageNumbersForeColor = RHS
    End If
    gPrinterExPageNumbersForeColor = RHS
End Property

Private Property Get IPrinterEx_PageNumbersForeColor() As OLE_COLOR
    IPrinterEx_PageNumbersForeColor = PageNumbersForeColor
End Property


Private Property Let IPrinterEx_PageNumbersFormat(ByVal RHS As String)
    If mLoaded_PageNumbersFormat Then
        mOriginal_PageNumbersFormat = RHS
    Else
        PageNumbersFormat = RHS
    End If
    If Not mPrintFnObject Is Nothing Then
        mPrintFnObject.EnsurePageNumberFormatInList RHS
    End If
End Property

Private Property Get IPrinterEx_PageNumbersFormat() As String
    IPrinterEx_PageNumbersFormat = PageNumbersFormat
End Property


Private Property Let IPrinterEx_PageNumbersPosition(ByVal RHS As vbExPageNumbersPositionConstants)
    If mLoaded_PageNumbersPosition Then
        mOriginal_PageNumbersPosition = RHS
    Else
        PageNumbersPosition = RHS
    End If
    gPrinterExPageNumbersPosition = RHS
End Property

Private Property Get IPrinterEx_PageNumbersPosition() As vbExPageNumbersPositionConstants
    IPrinterEx_PageNumbersPosition = PageNumbersPosition
End Property


Private Property Get IPrinterEx_Printed() As Boolean
    IPrinterEx_Printed = Printed
End Property


Private Property Let IPrinterEx_PrintPrevToolBarIconsSize(RHS As vbExPrintPrevToolBarIconsSizeConstants)
    mPrintPrevToolBarIconsSize = RHS
End Property

Private Property Get IPrinterEx_PrintPrevToolBarIconsSize() As vbExPrintPrevToolBarIconsSizeConstants
    IPrinterEx_PrintPrevToolBarIconsSize = mPrintPrevToolBarIconsSize
End Property


Private Property Let IPrinterEx_PrintPrevUseAltScaleIcons(RHS As Boolean)
    mPrintPrevUseAltScaleIcons = RHS
End Property

Private Property Get IPrinterEx_PrintPrevUseAltScaleIcons() As Boolean
    IPrinterEx_PrintPrevUseAltScaleIcons = mPrintPrevUseAltScaleIcons
End Property


Private Sub IPrinterEx_PrintRichTextBox(nRTB As Object)
    PrintRichTextBox nRTB
End Sub

Private Sub IPrinterEx_SetPageFixedElement(Optional ByVal nPicture As Variant, Optional ByVal nText As Variant, Optional ByVal nPositionTop As Variant)
    SetPageFixedElement nPicture, nText, nPositionTop
End Sub


Private Sub IPrinterEx_ShowPageSetup(Optional ByVal nFlags As cdeCommonDialogExPageSetupFlagsConstants = -1)
    Dim iUnitsMultiplier As Single
    Dim iLeftMargin_Prev As Single
    Dim iTopMargin_Prev As Single
    Dim iRightMargin_Prev As Single
    Dim iBottomMargin_Prev As Single
    
    Dim iDeviceName_Prev As String
    Dim iPaperSize_Prev As Long
    Dim iPaperBin_Prev As Long
    Dim iPrintQuality_Prev As Long
    Dim iColorMode_Prev As Long
    Dim iFlags_Prev As Long
    
    If gPrinterExFromPrintFn > 0 Then Exit Sub
    
    If mShowingPrintPreview Then
        RaiseError 1214, "PrinterEx", "Invalid operation when inside ShowPrintPreview"
        Exit Sub
    End If
    
    mDlg.PaperSize = mPaperSize
    mDlg.PaperBin = mPaperBin
    mDlg.Duplex = mDuplex
    mDlg.Orientation = mOrientation
    mDlg.PrintQuality = mPrintQuality
    mDlg.ColorMode = mColorMode
    mDlg.UpdateDevModeWithCurrentSettings
    
    iFlags_Prev = mDlg.Flags
    mDlg.Flags = mPageSetupFlags
    If nFlags <> -1 Then
        mDlg.Flags = mDlg.Flags Or nFlags
    End If
    
    If (LeftMargin > 0) Or (TopMargin > 0) Or (RightMargin > 0) Or (BottomMargin > 0) Then
        mDlg.LeftMargin = LeftMargin
        mDlg.TopMargin = TopMargin
        mDlg.RightMargin = RightMargin
        mDlg.BottomMargin = BottomMargin
    End If
    
    iLeftMargin_Prev = mDlg.LeftMargin
    iTopMargin_Prev = mDlg.TopMargin
    iRightMargin_Prev = mDlg.RightMargin
    iBottomMargin_Prev = mDlg.BottomMargin
    
    iDeviceName_Prev = mDlg.DeviceName
    iPaperSize_Prev = mDlg.PaperBin
    iPaperBin_Prev = mDlg.PaperSize
    iPrintQuality_Prev = mDlg.PrintQuality
    iColorMode_Prev = mDlg.ColorMode
    
    mDlg.ShowPageSetup
    mCanceled = mDlg.Canceled
    If Not mCanceled Then
        mDontCheckError396 = True
        DeviceName = mDlg.DeviceName
        Orientation = mDlg.Orientation
        PaperSize = mDlg.PaperSize
        PaperBin = mDlg.PaperBin
        PrintQuality = mDlg.PrintQuality
        ColorMode = mDlg.ColorMode
        DevModePtr = mDlg.DevModePtr
        mDontCheckError396 = False
        
        If mDlg.Units = vbInches Then
            iUnitsMultiplier = 25.4
        Else
            iUnitsMultiplier = 1
        End If
        LeftMargin = mDlg.LeftMargin * iUnitsMultiplier
        TopMargin = mDlg.TopMargin * iUnitsMultiplier
        RightMargin = mDlg.RightMargin * iUnitsMultiplier
        BottomMargin = mDlg.BottomMargin * iUnitsMultiplier
        If (iLeftMargin_Prev <> mDlg.LeftMargin) Or (iTopMargin_Prev <> mDlg.TopMargin) Or (iRightMargin_Prev <> mDlg.RightMargin) Or (iBottomMargin_Prev <> mDlg.BottomMargin) Then
            gPrinterExLeftMargin = mDlg.LeftMargin
            gPrinterExTopMargin = mDlg.TopMargin
            gPrinterExRightMargin = mDlg.RightMargin
            gPrinterExBottomMargin = mDlg.BottomMargin
        End If
        
        If mDlg.DeviceName <> "" Then
            If mDlg.DeviceName <> iDeviceName_Prev Then gPrinterExDeviceName = mDlg.DeviceName
        End If
        If mDlg.PaperBin <> iPaperSize_Prev Then gPrinterExPaperSize = mDlg.PaperBin
        If mDlg.PaperSize <> iPaperBin_Prev Then gPrinterExPaperSize = mDlg.PaperSize
        If mDlg.PrintQuality <> iPrintQuality_Prev Then gPrinterExPrintQuality = mDlg.PrintQuality
        If mDlg.ColorMode <> iColorMode_Prev Then gPrinterExColorMode = mDlg.ColorMode
        
        CheckMarginsLimits
        
    End If
    mDlg.Flags = iFlags_Prev

End Sub


Private Sub IPrinterEx_ShowPrint(Optional ByVal nFlags As cdeCommonDialogExPrinterFlagsConstants = -1, Optional nDocKey As String, Optional ByVal CausesEndDoc As Boolean = True)
    Dim iDeviceName_Prev As String
    Dim iPaperSize_Prev As Long
    Dim iPaperBin_Prev As Long
    Dim iPrintQuality_Prev As Long
    Dim iColorMode_Prev As Long
    Dim iCollate_Prev As Boolean
    Dim iOrientation_Prev As Long
    Dim iDuplex_Prev As Long
    Dim iFlags_Prev As Long
    Dim iDocKey_Prev As String
    
    If gPrinterExFromPrintFn > 0 Then Exit Sub
    
    If mShowingPrintPreview Then
        RaiseError 1214, "PrinterEx", "Invalid operation when inside ShowPrintPreview"
        Exit Sub
    End If
    
    iDocKey_Prev = mDocKey
    If nDocKey <> "" Then DocKey = nDocKey

    mDlg.PaperSize = mPaperSize
    mDlg.PaperBin = mPaperBin
    mDlg.Duplex = mDuplex
    mDlg.Orientation = mOrientation
    mDlg.PrintQuality = mPrintQuality
    mDlg.ColorMode = mColorMode
    mDlg.UpdateDevModeWithCurrentSettings
    
    iFlags_Prev = mDlg.Flags
    mDlg.Flags = mPrinterFlags
    If nFlags <> -1 Then
        mDlg.Flags = mDlg.Flags Or nFlags
    End If
    
    mDlg.Max = 9999 ' we still don't know how many pages will be printed.
    If mFromPage <> 0 Then
        mDlg.FromPage = mFromPage
    End If
    If mToPage <> 0 Then
        mDlg.ToPage = mToPage
    End If
    
    iDeviceName_Prev = mDlg.DeviceName
    iPaperSize_Prev = mDlg.PaperBin
    iPaperBin_Prev = mDlg.PaperSize
    iPrintQuality_Prev = mDlg.PrintQuality
    iColorMode_Prev = mDlg.ColorMode
    iCollate_Prev = mDlg.Collate
    iOrientation_Prev = mDlg.Orientation
    iDuplex_Prev = mDlg.Duplex
    
    If (mFromPage <> 0) And (mToPage <> 0) Then
        mDlg.ShowPrinter cdePDPageNums
    Else
        mDlg.ShowPrinter
    End If
    mPrinted = False
    mCanceled = mDlg.Canceled
    
    If Not mCanceled Then
        mDontCheckError396 = True
        DeviceName = mDlg.DeviceName
        Orientation = mDlg.Orientation
        PaperSize = mDlg.PaperSize
        PaperBin = mDlg.PaperBin
        PrintQuality = mDlg.PrintQuality
        Duplex = mDlg.Duplex
        ColorMode = mDlg.ColorMode
        DevModePtr = mDlg.DevModePtr
        mDontCheckError396 = False
        
        If mDlg.DeviceName <> "" Then
            If mDlg.DeviceName <> iDeviceName_Prev Then
                gPrinterExDeviceName = mDlg.DeviceName
                mSave_DeviceName = True
            End If
        End If
        If mDlg.PaperBin <> iPaperBin_Prev Then
            gPrinterExPaperBin = mDlg.PaperBin
            mSave_PaperBin = True
        End If
        If mDlg.PaperSize <> iPaperSize_Prev Then
            gPrinterExPaperSize = mDlg.PaperSize
            mSave_PaperSize = True
        End If
        If mDlg.PrintQuality <> iPrintQuality_Prev Then
            gPrinterExPrintQuality = mDlg.PrintQuality
            mSave_PrintQuality = True
        End If
        If mDlg.ColorMode <> iColorMode_Prev Then
            gPrinterExColorMode = mDlg.ColorMode
            mSave_ColorMode = True
        End If
        If mDlg.Collate <> iCollate_Prev Then
            gPrinterExCollate = mDlg.Collate
            mSave_Collate = True
        End If
        If mDlg.Orientation <> iOrientation_Prev Then
            mSave_Orientation = True
        End If
        If mDlg.Duplex <> iDuplex_Prev Then
            mSave_Duplex = True
        End If
        
        SaveUserDocPreferences
        
        mPrinterExEvents.RaiseEvent_PrepareDoc mDocKey
        
        If CausesEndDoc Then
            EndDoc
        End If
    Else
        If CausesEndDoc Then
            KillDoc
        End If
    End If
    
    DocKey = iDocKey_Prev
End Sub


Private Sub IPrinterEx_ShowPrintPreview(Optional nObject As Object, Optional nProcedureName As String, Optional nDocKey As String)
    Dim c As Long
    Dim iDocKey_Prev As String
    
    Dim iDeviceName_Prev As String
    Dim iPaperSize_Prev As Long
    Dim iPaperBin_Prev As Long
    Dim iPrintQuality_Prev As Long
    Dim iColorMode_Prev As Long
    Dim iCollate_Prev As Boolean
    Dim iOrientation_Prev As Long
    Dim iDuplex_Prev As Long
    
    Dim iScalePercent_Prev As Long
    Dim iPrintPageNumbers_Prev As Boolean
    Dim iPageNumbersPosition_Prev As Long
    Dim iPageNumbersFormat_Prev As String
    Dim iPageNumbersFont_Prev As Object
    Dim iPageNumbersForeColor_Prev As Long
    
    Dim iLeftMargin_Prev As Boolean
    Dim iTopMargin_Prev As Boolean
    Dim iRightMargin_Prev As Boolean
    Dim iBottomMargin_Prev As Boolean
    
    If mShowingPrintPreview Then
        RaiseError 1213, "PrinterEx", "Already inside ShowPrintPreview"
        Exit Sub
    End If
    
    mPrinted = False
    mShowingPrintPreview = True
    gPrinterExDoNotMakeNewPrinterExObject = True
    
    iDocKey_Prev = mDocKey
    If nDocKey <> "" Then DocKey = nDocKey
    
    If mPrintFnObject Is Nothing Then
        Set mPrintFnObject = New PrintFnObject
    End If
    Set mPrintFnObject.Parent = nObject
    mPrintFnObject.ProcedureName = nProcedureName
    mPrintFnObject.FormatButtonVisible = mPrintPrevFormatButtonVisible
    If mPrintPrevFormatButtonVisible Then
        mPrintFnObject.FormatButtonToolTipText = mPrintPrevFormatButtonToolTipText
        For c = 0 To 4
            If Not mPrintPrevFormatButtonPicture(c) Is Nothing Then
                Set mPrintFnObject.FormatButtonPicture(c) = mPrintPrevFormatButtonPicture(c)
            End If
        Next c
    End If
    mPrintFnObject.PageSetupButtonVisible = mPrintPrevPageSetupButtonVisible
    mPrintFnObject.MinScalePercent = mPrintPrevMinScalePercent
    mPrintFnObject.MaxScalePercent = mPrintPrevMaxScalePercent
    
    mPrintFnObject.PrintPrevToolBarIconsSize = mPrintPrevToolBarIconsSize
    mPrintFnObject.PrintPrevUseAltScaleIcons = mPrintPrevUseAltScaleIcons
    mPrintFnObject.HandleMargins = mHandleMargins
    mPrintFnObject.PrinterFlags = mPrinterFlags
    mPrintFnObject.PageSetupFlags = mPageSetupFlags
    mPrintFnObject.FromPage = mFromPage
    mPrintFnObject.ToPage = mToPage
    
    mPrintFnObject.CommonDialogExObject.DeviceName = mDeviceName
    mPrintFnObject.MinLeftMargin = MinLeftMargin
    mPrintFnObject.MinRightMargin = MinRightMargin
    mPrintFnObject.MinTopMargin = MinTopMargin
    mPrintFnObject.MinBottomMargin = MinBottomMargin
    mPrintFnObject.UnitsForUser = UnitsForUser
    
    mPrintFnObject.LeftMargin = LeftMargin
    mPrintFnObject.TopMargin = TopMargin
    mPrintFnObject.RightMargin = RightMargin
    mPrintFnObject.BottomMargin = BottomMargin
    
    mPrintFnObject.PaperSize = mPaperSize
    mPrintFnObject.PaperBin = mPaperBin
    mPrintFnObject.Duplex = mDuplex
    mPrintFnObject.Orientation = mOrientation
    mPrintFnObject.PrintQuality = mPrintQuality
    mPrintFnObject.ColorMode = mColorMode
    mPrintFnObject.CommonDialogExObject.UpdateDevModeWithCurrentSettings
    mPrintFnObject.PrintPageNumbers = mPrintPageNumbers
    mPrintFnObject.PageNumbersPosition = mPageNumbersPosition
    mPrintFnObject.PageNumbersFormat = mPageNumbersFormat
    Set mPrintFnObject.PageNumbersFont = mPageNumbersFont
    mPrintFnObject.PageNumbersForeColor = mPageNumbersForeColor
    
    iDeviceName_Prev = mPrintFnObject.CommonDialogExObject.DeviceName
    iPaperSize_Prev = mPrintFnObject.CommonDialogExObject.PaperBin
    iPaperBin_Prev = mPrintFnObject.CommonDialogExObject.PaperSize
    iPrintQuality_Prev = mPrintFnObject.CommonDialogExObject.PrintQuality
    iColorMode_Prev = mPrintFnObject.CommonDialogExObject.ColorMode
    iCollate_Prev = mPrintFnObject.CommonDialogExObject.Collate
    iOrientation_Prev = mPrintFnObject.CommonDialogExObject.Orientation
    iDuplex_Prev = mPrintFnObject.CommonDialogExObject.Duplex
    
    iScalePercent_Prev = mPrintFnObject.ScalePercent
    iPrintPageNumbers_Prev = mPrintFnObject.PrintPageNumbers
    iPageNumbersPosition_Prev = mPrintFnObject.PageNumbersPosition
    iPageNumbersFormat_Prev = mPrintFnObject.PageNumbersFormat
    Set iPageNumbersFont_Prev = mPrintFnObject.PageNumbersFont
    iPageNumbersForeColor_Prev = mPrintFnObject.PageNumbersForeColor
    
    iLeftMargin_Prev = mPrintFnObject.LeftMargin
    iTopMargin_Prev = mPrintFnObject.TopMargin
    iRightMargin_Prev = mPrintFnObject.RightMargin
    iBottomMargin_Prev = mPrintFnObject.BottomMargin
    
    mPrintFnObject.ShowPrintPreview mDocKey
    
    gPrinterExDoNotMakeNewPrinterExObject = False
    mShowingPrintPreview = False
    
    If mDlg.DeviceName <> "" Then
        If mDlg.DeviceName <> iDeviceName_Prev Then mSave_DeviceName = True
    End If
    If mDlg.PaperBin <> iPaperBin_Prev Then mSave_PaperBin = True
    If mDlg.PaperSize <> iPaperSize_Prev Then mSave_PaperSize = True
    If mDlg.PrintQuality <> iPrintQuality_Prev Then mSave_PrintQuality = True
    If mDlg.ColorMode <> iColorMode_Prev Then mSave_ColorMode = True
    If mDlg.Collate <> iCollate_Prev Then mSave_Collate = True
    If mDlg.Orientation <> iOrientation_Prev Then mSave_Orientation = True
    If mDlg.Duplex <> iDuplex_Prev Then mSave_Duplex = True
    
    If mPrintFnObject.ScalePercent <> iScalePercent_Prev Then mSave_ScalePercent = True
    If mPrintFnObject.PrintPageNumbers <> iPrintPageNumbers_Prev Then mSave_PrintPageNumbers = True
    If mPrintFnObject.PrintPageNumbers Then
        If mPrintFnObject.PageNumbersPosition <> iPageNumbersPosition_Prev Then mSave_PageNumbersPosition = True
        If mPrintFnObject.PageNumbersFormat <> iPageNumbersFormat_Prev Then mSave_PageNumbersFormat = True
        If Not FontsAreEqual(mPrintFnObject.PageNumbersFont, iPageNumbersFont_Prev) Then mSave_PageNumbersFont = True
        If mPrintFnObject.PageNumbersForeColor <> iPageNumbersForeColor_Prev Then mSave_PageNumbersForeColor = True
    End If
    
    If mPrintFnObject.LeftMargin <> iLeftMargin_Prev Then mSave_LeftMargin = True
    If mPrintFnObject.TopMargin <> iTopMargin_Prev Then mSave_TopMargin = True
    If mPrintFnObject.RightMargin <> iRightMargin_Prev Then mSave_RightMargin = True
    If mPrintFnObject.BottomMargin <> iBottomMargin_Prev Then mSave_BottomMargin = True
    
    If mPrinted Then
        SaveUserDocPreferences
    End If
    
    DocKey = iDocKey_Prev
End Sub


Private Property Get IPrinterEx_SomethingPrintedOnPage() As Boolean
    IPrinterEx_SomethingPrintedOnPage = SomethingPrintedOnPage
End Property


Private Property Let IPrinterEx_ToPage(ByVal RHS As Long)
    ToPage = RHS
End Property

Private Property Get IPrinterEx_ToPage() As Long
    IPrinterEx_ToPage = ToPage
End Property


Private Property Let IPrinterEx_LeftMargin(ByVal RHS As Single)
    If mLoaded_LeftMargin Then
        mOriginal_LeftMargin = RHS
    Else
        LeftMargin = RHS
    End If
    gPrinterExLeftMargin = RHS
End Property

Private Property Get IPrinterEx_LeftMargin() As Single
    IPrinterEx_LeftMargin = LeftMargin
End Property

Private Property Get IPrinterEx_PageLeftMargin() As Single
    IPrinterEx_PageLeftMargin = PageLeftMargin
End Property


Private Property Let IPrinterEx_TopMargin(ByVal RHS As Single)
    If mLoaded_TopMargin Then
        mOriginal_TopMargin = RHS
    Else
        TopMargin = RHS
    End If
    gPrinterExTopMargin = RHS
End Property

Private Property Get IPrinterEx_TopMargin() As Single
    IPrinterEx_TopMargin = TopMargin
End Property

Private Property Get IPrinterEx_PageTopMargin() As Single
    IPrinterEx_PageTopMargin = PageTopMargin
End Property


Private Property Let IPrinterEx_RightMargin(ByVal RHS As Single)
    If mLoaded_RightMargin Then
        mOriginal_RightMargin = RHS
    Else
        RightMargin = RHS
    End If
    gPrinterExRightMargin = RHS
End Property

Private Property Get IPrinterEx_RightMargin() As Single
    IPrinterEx_RightMargin = RightMargin
End Property

Private Property Get IPrinterEx_PageRightMargin() As Single
    IPrinterEx_PageRightMargin = PageRightMargin
End Property


Private Property Let IPrinterEx_BottomMargin(ByVal RHS As Single)
    If mLoaded_BottomMargin Then
        mOriginal_BottomMargin = RHS
    Else
        BottomMargin = RHS
    End If
    gPrinterExBottomMargin = RHS
End Property

Private Property Get IPrinterEx_BottomMargin() As Single
    IPrinterEx_BottomMargin = BottomMargin
End Property

Private Property Get IPrinterEx_PageBottomMargin() As Single
    IPrinterEx_PageBottomMargin = PageBottomMargin
End Property


Private Property Let IPrinterEx_MinLeftMargin(ByVal RHS As Single)
    MinLeftMargin = RHS
    gPrinterExMinLeftMargin = RHS
End Property

Private Property Get IPrinterEx_MinLeftMargin() As Single
    IPrinterEx_MinLeftMargin = MinLeftMargin
End Property


Private Property Let IPrinterEx_MinRightMargin(ByVal RHS As Single)
    MinRightMargin = RHS
    gPrinterExMinRightMargin = RHS
End Property

Private Property Get IPrinterEx_MinRightMargin() As Single
    IPrinterEx_MinRightMargin = MinRightMargin
End Property


Private Property Let IPrinterEx_MinTopMargin(ByVal RHS As Single)
    MinTopMargin = RHS
    gPrinterExMinTopMargin = RHS
End Property

Private Property Get IPrinterEx_MinTopMargin() As Single
    IPrinterEx_MinTopMargin = MinTopMargin
End Property


Private Property Let IPrinterEx_MinBottomMargin(ByVal RHS As Single)
    MinBottomMargin = RHS
    gPrinterExMinBottomMargin = RHS
End Property

Private Property Get IPrinterEx_MinBottomMargin() As Single
    IPrinterEx_MinBottomMargin = MinBottomMargin
End Property


Private Property Get IPrinterEx_PageCount() As Long
    IPrinterEx_PageCount = PageCount
End Property


Private Sub IPrinterEx_PrintEx(Optional Text As Variant, Optional ByVal Options As vbExPrintExOptionsConstants, Optional ByVal Left As Variant, Optional ByVal Top As Variant, Optional ByVal Width As Variant, Optional ByVal Height As Variant, Optional ByVal Alignment As vbExPrintExAlignConstants, Optional ByVal WordWrap As Boolean = True, Optional ByVal WordBreak As Boolean = True, Optional ByVal RectMarginsX As Single, Optional ByVal RectMarginsY As Single, Optional ByRef StartsNewPage As Boolean, Optional ByRef TextExceedsWidth As Boolean, Optional ByVal CancelIfTextExceedsWidth As Boolean)
    Dim iAddLineFeed As Boolean
    Dim iAddTab As Boolean
    Dim iCellMarginsX As Single
    Dim iCellMarginsY As Single
    Dim iText As String
    
    If (Options And vxNoLineFeed) = 0 Then
        iAddLineFeed = True
    End If
    If (Options And vxAddTab) <> 0 Then
        iAddTab = True
        iAddLineFeed = False
    End If
    
    iCellMarginsX = RectMarginsX
    iCellMarginsY = RectMarginsY
    
    If Not IsMissing(Text) Then
        iText = CStr(Text)
    End If
    
    If iAddTab Then
        ' the tabs are not added properly (not added at all in fact), from http://msdn.microsoft.com/en-us/library/aa227486%28v=vs.60%29.aspx  The Tab function will position the output cursor at a particular column in the output field. The width of a column is determined by the average width of all of the characters in the current font.
        PrintText iText & vbTab, Left, Top, Width, Height, Alignment, iAddLineFeed, WordWrap, WordBreak, iCellMarginsX, iCellMarginsY, StartsNewPage, , , TextExceedsWidth, CancelIfTextExceedsWidth
    Else
        PrintText iText, Left, Top, Width, Height, Alignment, iAddLineFeed, WordWrap, WordBreak, iCellMarginsX, iCellMarginsY, StartsNewPage, , , TextExceedsWidth, CancelIfTextExceedsWidth
    End If
End Sub


Private Function IPrinterEx_TextHeightEx(ByVal Str As String, Optional ByVal Options As vbExPrintExOptionsConstants = 0&, Optional ByVal Left As Variant, Optional ByVal Top As Variant, Optional ByVal Width As Variant, Optional ByVal Alignment As vbExPrintExAlignConstants = 0&, Optional ByVal WordWrap As Boolean = True, Optional ByVal WordBreak As Boolean = True, Optional ByVal CellMarginsX As Single, Optional ByVal CellMarginsY As Single, Optional ByRef StartsNewPage As Boolean) As Single
    Dim iAddLineFeed As Boolean
    Dim iAddTab As Boolean
    Dim iCellMarginsX As Single
    Dim iCellMarginsY As Single

    If (Options And vxNoLineFeed) = 0 Then
        iAddLineFeed = True
    End If
    If (Options And vxAddTab) <> 0 Then
        iAddTab = True
    End If

    iCellMarginsX = CellMarginsX
    iCellMarginsY = CellMarginsY

    If iAddTab Then
        IPrinterEx_TextHeightEx = PrintText(Str & vbTab, Left, Top, Width, , Alignment, iAddLineFeed, WordWrap, WordBreak, iCellMarginsX, iCellMarginsY, StartsNewPage, True)
    Else
        IPrinterEx_TextHeightEx = PrintText(Str, Left, Top, Width, , Alignment, iAddLineFeed, WordWrap, WordBreak, iCellMarginsX, iCellMarginsY, StartsNewPage, True)
    End If
End Function


Private Function IPrinterEx_TextWidthEx(ByVal Str As String, Optional ByVal Options As vbExPrintExOptionsConstants = 0&, Optional ByVal Left As Variant, Optional ByVal Top As Variant, Optional ByVal Height As Variant, Optional ByVal Alignment As vbExPrintExAlignConstants = 0&, Optional ByVal WordWrap As Boolean = True, Optional ByVal WordBreak As Boolean = True, Optional ByVal CellMarginsX As Single, Optional ByVal CellMarginsY As Single, Optional ByRef StartsNewPage As Boolean) As Single
    Dim iAddLineFeed As Boolean
    Dim iAddTab As Boolean
    Dim iCellMarginsX As Single
    Dim iCellMarginsY As Single
    Dim iRequiredWidth As Single

    If (Options And vxNoLineFeed) = 0 Then
        iAddLineFeed = True
    End If
    If (Options And vxAddTab) <> 0 Then
        iAddTab = True
    End If

    iCellMarginsX = CellMarginsX
    iCellMarginsY = CellMarginsY

    If iAddTab Then
        PrintText Str & vbTab, Left, Top, , Height, Alignment, iAddLineFeed, WordWrap, WordBreak, iCellMarginsX, iCellMarginsY, StartsNewPage, True, iRequiredWidth
    Else
        PrintText Str, Left, Top, , Height, Alignment, iAddLineFeed, WordWrap, WordBreak, iCellMarginsX, iCellMarginsY, StartsNewPage, True, iRequiredWidth
    End If
    IPrinterEx_TextWidthEx = iRequiredWidth
End Function


Private Property Let IPrinterEx_Units(ByVal RHS As cdeUnits)
    Units = RHS
    gPrinterExUnits = RHS
End Property

Private Property Get IPrinterEx_Units() As cdeUnits
    IPrinterEx_Units = Units
End Property


Private Property Let IPrinterEx_UnitsForUser(ByVal RHS As cdeUnitsForUser)
    UnitsForUser = RHS
    gPrinterExUnitsForUser = RHS
End Property

Private Property Get IPrinterEx_UnitsForUser() As cdeUnitsForUser
    IPrinterEx_UnitsForUser = UnitsForUser
End Property

Private Sub mPageNumbersFont_FontChanged(ByVal PropertyName As String)
    If mPrintPageNumbers Then
        CheckMarginsLimits
    End If
End Sub

Private Sub mPrintFnObject_AfterShowingPageSetupDialog()
    mPrinterExEvents.RaiseEvent_AfterShowingPageSetupDialog
End Sub

Private Sub mPrintFnObject_AfterShowingPrinterDialog()
    mPrinterExEvents.RaiseEvent_AfterShowingPrinterDialog
End Sub

Private Sub mPrintFnObject_BeforeShowingPageSetupDialog(CancelDisplay As Boolean)
    mPrinterExEvents.RaiseEvent_BeforeShowingPageSetupDialog CancelDisplay
End Sub

Private Sub mPrintFnObject_BeforeShowingPrinterDialog(CancelPrint As Boolean)
    mPrinterExEvents.RaiseEvent_BeforeShowingPrinterDialog CancelPrint
End Sub

Private Sub mPrintFnObject_FormatOptionsClick(Canceled As Boolean)
    mPrinterExEvents.RaiseEvent_FormatOptionsClick Canceled
End Sub

Private Sub mPrintFnObject_PrepareDoc(Cancel As Boolean, ByVal DocKey As String)
    mPrinterExEvents.RaiseEvent_PrepareDoc DocKey
End Sub

Private Sub mPrintFnObject_ScaleChange(NewScalePercent As Integer)
    mPrinterExEvents.RaiseEvent_ScaleChange NewScalePercent
End Sub

' End IPrinterEx Interface

' Printer Interface

'.Circle [Step] (x, y), radius, [color, start, end, aspect]
Private Sub Printer_Circle(ByVal Step As Integer, ByVal x As Single, ByVal y As Single, ByVal Radius As Single, ByVal Color As Long, ByVal StartArc As Single, ByVal EndArc As Single, ByVal Aspect As Single)
    Dim iStep As Boolean
    Dim iColor As Long
    Dim iEndArcIsMissing As Boolean
    Dim iStartArcIsMissing As Boolean
    
    iStep = (Step And 1) = 1
    If (Step And 2) = 2 Then
        iColor = Color
    Else
        iColor = -1
    End If
    iStartArcIsMissing = (Step And 64) = 0
    iEndArcIsMissing = (Step And 128) = 0
    If Aspect = 0 Then Aspect = 1
    
    If iStartArcIsMissing And iEndArcIsMissing Then
        DrawCircle x, y, Radius, Aspect, iColor, , , iStep
    ElseIf iStartArcIsMissing Then
        DrawCircle x, y, Radius, Aspect, iColor, , EndArc, iStep
    ElseIf iEndArcIsMissing Then
        DrawCircle x, y, Radius, Aspect, iColor, StartArc, , iStep
    Else
        DrawCircle x, y, Radius, Aspect, iColor, StartArc, EndArc, iStep
    End If

End Sub


Private Property Let Printer_ColorMode(ByVal RHS As Integer)
    If mLoaded_ColorMode Then
        mOriginal_ColorMode = RHS
    Else
        ColorMode = RHS
    End If
End Property

Private Property Get Printer_ColorMode() As Integer
    Printer_ColorMode = ColorMode
End Property


Private Property Let Printer_Copies(ByVal RHS As Integer)
    Copies = RHS
End Property

Private Property Get Printer_Copies() As Integer
    Printer_Copies = Copies
End Property


Private Property Let Printer_CurrentX(ByVal RHS As Single)
    CurrentX = RHS
    If (mScaleMode = vbTwips) Or (mScaleMode = vbPixels) Then
        CurrentX = Round(RHS)
    Else
        CurrentX = RHS
    End If
End Property

Private Property Get Printer_CurrentX() As Single
    Printer_CurrentX = CurrentX
    If (mScaleMode = vbTwips) Or (mScaleMode = vbPixels) Then
        Printer_CurrentX = Round(Printer_CurrentX)
    End If
End Property


Private Property Let Printer_CurrentY(ByVal RHS As Single)
    If (mScaleMode = vbTwips) Or (mScaleMode = vbPixels) Then
        CurrentY = Round(RHS)
    Else
        CurrentY = RHS
    End If
End Property

Private Property Get Printer_CurrentY() As Single
    Printer_CurrentY = CurrentY
    If (mScaleMode = vbTwips) Or (mScaleMode = vbPixels) Then
        Printer_CurrentY = Round(Printer_CurrentY)
    End If
End Property


Private Property Get Printer_DeviceName() As String
    Printer_DeviceName = DeviceName
End Property


Private Property Let Printer_DrawMode(ByVal RHS As Integer)
    DrawMode = RHS
End Property

Private Property Get Printer_DrawMode() As Integer
    Printer_DrawMode = DrawMode
End Property


Private Property Let Printer_DrawStyle(ByVal RHS As Integer)
    DrawStyle = RHS
End Property

Private Property Get Printer_DrawStyle() As Integer
    Printer_DrawStyle = DrawStyle
End Property


Private Property Let Printer_DrawWidth(ByVal RHS As Integer)
    DrawWidth = RHS
End Property

Private Property Get Printer_DrawWidth() As Integer
    Printer_DrawWidth = DrawWidth
End Property


Private Property Get Printer_DriverName() As String
    Printer_DriverName = DriverName
End Property


Private Property Get Printer_Duplex() As Integer
    Printer_Duplex = Duplex
End Property

Private Property Let Printer_Duplex(ByVal RHS As Integer)
    If mLoaded_Duplex Then
        mOriginal_Duplex = RHS
    Else
        Duplex = RHS
    End If
End Property


Private Sub Printer_EndDoc()
    EndDoc
End Sub


Private Property Let Printer_FillColor(ByVal RHS As Long)
    FillColor = RHS
End Property

Private Property Get Printer_FillColor() As Long
    Printer_FillColor = FillColor
End Property


Private Property Let Printer_FillStyle(ByVal RHS As Integer)
    FillStyle = RHS
End Property

Private Property Get Printer_FillStyle() As Integer
    Printer_FillStyle = FillStyle
End Property


Private Property Set Printer_Font(ByVal RHS As stdole.StdFont)
    Set Font = RHS
End Property

Private Property Get Printer_Font() As stdole.StdFont
    Set Printer_Font = Font
End Property


Private Property Let Printer_FontBold(ByVal RHS As Boolean)
    FontBold = RHS
End Property

Private Property Get Printer_FontBold() As Boolean
    Printer_FontBold = FontBold
End Property


Private Property Get Printer_FontCount() As Integer
    Printer_FontCount = FontCount
End Property


Private Property Let Printer_FontItalic(ByVal RHS As Boolean)
    FontItalic = RHS
End Property

Private Property Get Printer_FontItalic() As Boolean
    Printer_FontItalic = FontItalic
End Property


Private Property Let Printer_FontName(ByVal RHS As String)
    FontName = RHS
End Property

Private Property Get Printer_FontName() As String
    Printer_FontName = FontName
End Property


Private Property Get Printer_Fonts(ByVal Index As Integer) As String
    Printer_Fonts = Fonts(Index)
End Property


Private Property Let Printer_FontSize(ByVal RHS As Single)
    FontSize = RHS
End Property

Private Property Get Printer_FontSize() As Single
    Printer_FontSize = FontSize
End Property


Private Property Let Printer_FontStrikethru(ByVal RHS As Boolean)
    FontStrikeThru = RHS
End Property

Private Property Get Printer_FontStrikethru() As Boolean
    Printer_FontStrikethru = FontStrikeThru
End Property


Private Property Let Printer_FontTransparent(ByVal RHS As Boolean)
    FontTransparent = RHS
End Property

Private Property Get Printer_FontTransparent() As Boolean
    Printer_FontTransparent = FontTransparent
End Property


Private Property Let Printer_FontUnderline(ByVal RHS As Boolean)
    FontUnderLine = RHS
End Property

Private Property Get Printer_FontUnderline() As Boolean
    Printer_FontUnderline = FontUnderLine
End Property


Private Property Let Printer_ForeColor(ByVal RHS As Long)
    ForeColor = RHS
End Property

Private Property Get Printer_ForeColor() As Long
    Printer_ForeColor = ForeColor
End Property


Private Property Get Printer_hDC() As Long
    Printer_hDC = hDC
End Property


Private Property Let Printer_Height(ByVal RHS As Long)
    Height = RHS
End Property

Private Property Get Printer_Height() As Long
    Printer_Height = Height
End Property


Private Sub Printer_KillDoc()
    KillDoc
End Sub


Private Sub Printer_Line(ByVal Flags As Integer, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal Color As Long)
    Dim B As Boolean
    Dim f As Boolean
    Dim c As Boolean
    Dim iStep1 As Boolean
    Dim iStep2 As Boolean
    Dim iFirstCoord As Boolean
    
    iFirstCoord = (Flags And LF_FIRSTCOORD)
    iStep1 = (Flags And LF_STEP1)
    iStep2 = (Flags And LF_STEP2)
    c = (Flags And LF_COLOR)
    
    B = (Flags And LF_B)
    f = (Flags And LF_BF)
    If f Then B = True
    
    If iFirstCoord Then
        If iStep1 Then
            X1 = X1 + CurrentX
            Y1 = Y1 + CurrentY
        End If
    Else
        X1 = CurrentX
        Y1 = CurrentY
    End If
    If iStep2 Then
        X2 = X2 + X1
        Y2 = Y2 + Y1
    End If
    
    If c Then
        DrawLine X1, Y1, X2, Y2, Color, B, f
    Else
        DrawLine X1, Y1, X2, Y2, , B, f
    End If
    
End Sub


Private Sub Printer_NewPage()
    NewPage
End Sub


Private Property Let Printer_Orientation(ByVal RHS As Integer)
    If mLoaded_Orientation Then
        mOriginal_Orientation = RHS
    Else
        Orientation = RHS
    End If
End Property

Private Property Get Printer_Orientation() As Integer
    Printer_Orientation = Orientation
End Property


Private Property Get Printer_Page() As Integer
    Printer_Page = Page
End Property


Private Sub Printer_PaintPicture(ByVal Picture As stdole.IPictureDisp, ByVal X1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal X2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
    PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
End Sub


Private Property Let Printer_PaperBin(ByVal RHS As Integer)
    If mLoaded_PaperBin Then
        mOriginal_PaperBin = RHS
    Else
        PaperBin = RHS
    End If
End Property
    
Private Property Get Printer_PaperBin() As Integer
    Printer_PaperBin = PaperBin
End Property


Private Property Let Printer_PaperSize(ByVal RHS As Integer)
    If mDeviceName = "" Then
        mDeviceName = VB.Printer.DeviceName
        mPort = VB.Printer.Port
    End If
    If mLoaded_PaperSize Then
        mOriginal_PaperSize = RHS
    Else
        PaperSize = RHS
    End If
End Property

Private Property Get Printer_PaperSize() As Integer
    Printer_PaperSize = PaperSize
End Property


Private Property Get Printer_Port() As String
    Printer_Port = Port
End Property


Private Sub Printer_PrintEx(Optional Text, Optional Options As vbExPrintExOptionsConstants)
    Dim iAddLineFeed As Boolean
    Dim iAddTab As Boolean
    Dim iText As String
    
    If (Options And vxNoLineFeed) = 0 Then
        iAddLineFeed = True
    End If
    If (Options And vxAddTab) <> 0 Then
        iAddTab = True
        iAddLineFeed = False
    End If
    
    If Not IsMissing(Text) Then
        iText = CStr(Text)
    End If

    Select Case VarType(Text)
        Case vbByte, vbCurrency, vbDate, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
            iText = " " & iText
    End Select
    
    If iAddTab Then
        ' the tabs are not added properly (not added at all in fact), from http://msdn.microsoft.com/en-us/library/aa227486%28v=vs.60%29.aspx  The Tab function will position the output cursor at a particular column in the output field. The width of a column is determined by the average width of all of the characters in the current font.
        PrintText iText & vbTab, , , , , , iAddLineFeed, False, False
    Else
        PrintText iText, , , , , , iAddLineFeed, False, False
    End If
End Sub


Private Property Let Printer_PrintQuality(ByVal RHS As Integer)
    If mLoaded_PrintQuality Then
        mOriginal_PrintQuality = RHS
    Else
        PrintQuality = RHS
    End If
End Property

Private Property Get Printer_PrintQuality() As Integer
    Printer_PrintQuality = PrintQuality
End Property


Private Sub Printer_PSet(ByVal Step As Integer, ByVal x As Single, ByVal y As Single, ByVal Color As Long)
    Dim iStep As Boolean
    Dim iColor As Long
    
    iStep = (Step And 1) = 1
    If (Step And 2) = 2 Then
        iColor = Color
    Else
        iColor = -1
    End If
    DrawPoint x, y, iColor, iStep
End Sub


Private Property Get Printer_RightToLeft() As Boolean
    ' not implemented
End Property

Private Property Let Printer_RightToLeft(ByVal RHS As Boolean)
    ' not implemented
End Property


Private Sub Printer_Scale(ByVal Flags As Integer, Optional ByVal X1 As Variant, Optional ByVal Y1 As Variant, Optional ByVal X2 As Variant, Optional ByVal Y2 As Variant)
    Scale1 X1, Y1, X2, Y2
End Sub


Private Property Let Printer_ScaleHeight(ByVal RHS As Single)
    ScaleHeight = RHS
End Property

Private Property Get Printer_ScaleHeight() As Single
    Printer_ScaleHeight = ScaleHeight
End Property


Private Property Let Printer_ScaleLeft(ByVal RHS As Single)
    ScaleLeft = RHS
End Property

Private Property Get Printer_ScaleLeft() As Single
    Printer_ScaleLeft = ScaleLeft
End Property


Private Property Let Printer_ScaleMode(ByVal RHS As Integer)
    ScaleMode = RHS
End Property

Private Property Get Printer_ScaleMode() As Integer
    Printer_ScaleMode = ScaleMode
End Property


Private Property Let Printer_ScaleTop(ByVal RHS As Single)
    ScaleTop = RHS
End Property

Private Property Get Printer_ScaleTop() As Single
    Printer_ScaleTop = ScaleTop
End Property


Private Property Let Printer_ScaleWidth(ByVal RHS As Single)
    ScaleWidth = RHS
End Property

Private Property Get Printer_ScaleWidth() As Single
    Printer_ScaleWidth = ScaleWidth
End Property


Private Function Printer_ScaleX(ByVal Width As Single, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
    Printer_ScaleX = ScaleX(Width, FromScale, ToScale)
End Function


Private Function Printer_ScaleY(ByVal Height As Single, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
    Printer_ScaleY = ScaleY(Height, FromScale, ToScale)
End Function


Private Function Printer_TextHeight(ByVal Str As String) As Single
    Printer_TextHeight = IPrinterEx_TextHeightEx(Str, , , , , , , False)
    If (mScaleMode = vbTwips) Or (mScaleMode = vbPixels) Then
        Printer_TextHeight = Round(Printer_TextHeight)
    End If
End Function


Private Function Printer_TextWidth(ByVal Str As String) As Single
    Printer_TextWidth = TextWidth(Str)
    If (mScaleMode = vbTwips) Or (mScaleMode = vbPixels) Then
        Printer_TextWidth = Round(Printer_TextWidth)
    End If
End Function


Private Property Get Printer_TrackDefault() As Boolean
    Printer_TrackDefault = TrackDefault
End Property

Private Property Let Printer_TrackDefault(ByVal RHS As Boolean)
    TrackDefault = RHS
End Property


Private Property Get Printer_TwipsPerPixelX() As Single
    Printer_TwipsPerPixelX = TwipsPerPixelX
End Property


Private Property Get Printer_TwipsPerPixelY() As Single
    Printer_TwipsPerPixelY = TwipsPerPixelY
End Property


Private Property Let Printer_Zoom(ByVal RHS As Long)
    Zoom = RHS
    If mZoom = RHS Then
        gPrinterExZoom = mZoom
    End If
End Property

Private Property Get Printer_Zoom() As Long
    Printer_Zoom = Zoom
End Property


Private Property Let Printer_Width(ByVal RHS As Long)
    Width = RHS
End Property

Private Property Get Printer_Width() As Long
    Printer_Width = Width
End Property

' End Printer interface


Private Sub SetPageSize()
    Dim iPs As POINTAPI
    
    iPs = GetPaperSize(mPaperSize)
    If iPs.x > 0 Then
        mPaperWidthInMillimeters = iPs.x / 10
        mPaperHeightInMillimeters = iPs.y / 10
    Else
        If mDlg.Units = vbInches Then
            mPaperWidthInMillimeters = ScaleX(mDlg.PaperWidth, vbInches, vbMillimeters)
            mPaperHeightInMillimeters = ScaleX(mDlg.PaperHeight, vbInches, vbMillimeters)
        Else
            mPaperWidthInMillimeters = mDlg.PaperWidth
            mPaperHeightInMillimeters = mDlg.PaperHeight
        End If
    End If
End Sub

Private Property Get PaperWidthInMillimeters()
    If mPaperWidthInMillimeters = 0 Then
        PaperWidthInMillimeters = mPrinterDefaultPageWidthInMillimeters
    Else
        PaperWidthInMillimeters = mPaperWidthInMillimeters
    End If
End Property

Private Property Get PaperHeightInMillimeters()
    If mPaperHeightInMillimeters = 0 Then
        PaperHeightInMillimeters = mPrinterDefaultPageHeightInMillimeters
    Else
        PaperHeightInMillimeters = mPaperHeightInMillimeters
    End If
End Property

Private Property Get CurrentPageScaleWidthInPixels() As Long
    EnsureDocumentCreated
    If mHandleMargins Then
        CurrentPageScaleWidthInPixels = mPages(mCurrentPage).ScaleWidthInPixels
    Else
        CurrentPageScaleWidthInPixels = mPages(mCurrentPage).ScaleWidthInPixelsNoMargins
    End If
End Property

Private Property Get CurrentPageScaleHeightInPixels() As Long
    EnsureDocumentCreated
    If mHandleMargins Then
        CurrentPageScaleHeightInPixels = mPages(mCurrentPage).ScaleHeightInPixels
    Else
        CurrentPageScaleHeightInPixels = mPages(mCurrentPage).ScaleHeightInPixelsNoMargins
    End If
End Property


Private Function GetPaperSize(ByVal nPaperSizeNumber As Long) As POINTAPI
    Dim ret As Long
    Dim iPaperSizesNumbers() As Integer
    Dim iPaperSizes() As POINTAPI
    Dim c As Long
    Dim iLng As Long
    Dim iPrn As VB.Printer
    
    If nPaperSizeNumber = vbPRPSPrinterDefault Then
        For Each iPrn In VB.Printers
            If iPrn.DeviceName = mDeviceName Then
                nPaperSizeNumber = iPrn.PaperSize
                Exit For
            End If
        Next
    End If
    If nPaperSizeNumber = vbPRPSPrinterDefault Then
        nPaperSizeNumber = vbPRPSLetter
    End If
    
    ret = DeviceCapabilities(mDeviceName, mPort, DC_PAPERS, ByVal 0&, ByVal 0&)
    If ret = -1 Then Exit Function
    
    ReDim iPaperSizesNumbers(1 To ret)
    ReDim iPaperSizes(1 To ret)
    
    Call DeviceCapabilities(mDeviceName, mPort, DC_PAPERS, iPaperSizesNumbers(1), ByVal 0&)
    Call DeviceCapabilities(mDeviceName, mPort, DC_PAPERSIZE, iPaperSizes(1), ByVal 0&)
    
    For c = 1 To UBound(iPaperSizesNumbers)
        If iPaperSizesNumbers(c) = nPaperSizeNumber Then
            GetPaperSize.x = iPaperSizes(c).x
            GetPaperSize.y = iPaperSizes(c).y
            If GetPaperSize.x > GetPaperSize.y Then
                iLng = GetPaperSize.x
                GetPaperSize.x = GetPaperSize.y
                GetPaperSize.y = iLng
            End If
            Exit Function
        End If
    Next c
End Function

Private Sub LoadPrinterFonts()
    Dim iPrn As VB.Printer
    Dim c As Long
    Dim iDeviceNamePrev As String
    
    If VB.Printer.DeviceName <> mDeviceName Then
        iDeviceNamePrev = VB.Printer.DeviceName
        For Each iPrn In Printers
            If iPrn.DeviceName = mDeviceName Then
                Set VB.Printer = iPrn
                Exit For
            End If
        Next
    End If
    
    ReDim mFonts(VB.Printer.FontCount - 1)
    For c = 0 To VB.Printer.FontCount - 1
        mFonts(c) = VB.Printer.Fonts(c)
    Next c
    mFontCount = VB.Printer.FontCount
    
    If iDeviceNamePrev <> "" Then
        For Each iPrn In Printers
            If iPrn.DeviceName = iDeviceNamePrev Then
                Set VB.Printer = iPrn
                Exit For
            End If
        Next
    End If
End Sub

Public Property Get FontCount() As Integer
    FontCount = mFontCount
End Property

Public Property Get Fonts(Index As Integer) As String
    Fonts = mFonts(Index)
End Property

'Private Function FontExistInPrinter(nFontName As String) As Boolean
'    Dim c As Long
'    Dim iFn As String
'
'    iFn = Trim$(LCase$(nFontName))
'    For c = 0 To mFontCount - 1
'        If LCase$(mFonts(c)) = iFn Then
'            FontExistInPrinter = True
'            Exit For
'        End If
'    Next c
'End Function

Private Sub SetFontRatioForPrinterStdFont(nFont As iFont)
    If mDocumentCreated Then
        EnsureDocumentCreated
        Call nFont.SetRatio(mPages(mCurrentPage).MetaDPIY, 2540)
    Else
        Call nFont.SetRatio(GetDeviceCaps(VB.Printer.hDC, LOGPIXELSY), 2540)
    End If
End Sub

Public Sub KillLastPage()
    On Error Resume Next
    If mCurrentPage < 0 Then mCurrentPage = 0
    If Not mNewPagePending Then
        ReDim Preserve mPages(UBound(mPages) - 1)
        ReDim Preserve mPageNumbers(UBound(mPages))
    End If
    mNewPagePending = False
    If mCurrentPage > UBound(mPages) Then
        mCurrentPage = UBound(mPages)
    End If
    If mCurrentPage = 0 Then
        mPageCreated = False
        KillDoc
    End If
    mCurrentXPixels = -mScaleLeftPixels + PageLeftMarginInPixels
    mCurrentYPixels = -mScaleTopPixels + PageTopMarginInPixels
End Sub

Private Function StdFontToLogFont_Printer(nFont As StdFont, Optional nFontName As String) As LOGFONTW
    Dim iFontName As String
    Dim iDPIY As Single
    Dim c As Long
    Dim iBytes() As Byte
    
    If nFontName = "" Then
        iFontName = nFont.Name
    Else
        iFontName = nFontName
    End If
    
    If mDocumentCreated Then
        CheckNewPagePending
        iDPIY = mPages(mCurrentPage).MetaDPIY
    Else
        iDPIY = GetDeviceCaps(VB.Printer.hDC, LOGPIXELSY)
    End If
    
    iBytes = iFontName
    For c = 0 To 31
        If c < Len(iFontName) Then
            StdFontToLogFont_Printer.lfFaceName(c * 2) = iBytes(c * 2) 'Asc(Mid$(iFontName, c + 1, 1))
            StdFontToLogFont_Printer.lfFaceName(c * 2 + 1) = iBytes(c * 2 + 1)
        Else
            StdFontToLogFont_Printer.lfFaceName(c * 2) = 0
            StdFontToLogFont_Printer.lfFaceName(c * 2 + 1) = 0
        End If
    Next
     
    StdFontToLogFont_Printer.lfHeight = -Round(nFont.Size * iDPIY / 72)
     
    If nFont.Italic Then
        StdFontToLogFont_Printer.lfItalic = 1
    End If
    StdFontToLogFont_Printer.lfQuality = CLEARTYPE_QUALITY
    StdFontToLogFont_Printer.lfOutPrecision = OUT_TT_ONLY_PRECIS
    If nFont.Strikethrough Then
        StdFontToLogFont_Printer.lfStrikeOut = 1
    End If
    If nFont.Underline Then
        StdFontToLogFont_Printer.lfUnderline = 1
    End If
    StdFontToLogFont_Printer.lfWeight = nFont.Weight
    StdFontToLogFont_Printer.lfCharSet = nFont.Charset
    
End Function

Private Function LogFontToStdFont(lF As LOGFONTW) As iFont
    Dim iFontName As String
    Dim iDPIY As Single
    
    Set LogFontToStdFont = New StdFont
    SetFontRatioForPrinterStdFont LogFontToStdFont
    
    If mDocumentCreated Then
        CheckNewPagePending
        iDPIY = mPages(mCurrentPage).MetaDPIY
    Else
        iDPIY = GetDeviceCaps(VB.Printer.hDC, LOGPIXELSY)
    End If
    
    
    iFontName = CStr(lF.lfFaceName)
    If InStr(iFontName, Chr(0)) > 0 Then
        iFontName = Left$(iFontName, InStr(iFontName, Chr(0)) - 1)
    End If
    
    LogFontToStdFont.Name = iFontName
    LogFontToStdFont.Size = -lF.lfHeight / iDPIY * 72
    LogFontToStdFont.Weight = lF.lfWeight
    LogFontToStdFont.Italic = lF.lfItalic
    LogFontToStdFont.Strikethrough = lF.lfStrikeOut
    LogFontToStdFont.Underline = lF.lfUnderline
    LogFontToStdFont.Charset = lF.lfCharSet
End Function

Private Function GetFontAttribute(nAttribute As String) As Variant
    Dim iFontName As String
    Dim iDPIY As Single
    Dim iSng As Single
    
    Select Case nAttribute
        Case "Bold"
            GetFontAttribute = (mLogFont.lfWeight > 550)
        Case "Italic"
            GetFontAttribute = CBool(mLogFont.lfItalic)
        Case "Name"
            iFontName = CStr(mLogFont.lfFaceName)
            If InStr(iFontName, Chr(0)) > 0 Then
                iFontName = Left$(iFontName, InStr(iFontName, Chr(0)) - 1)
            End If
            GetFontAttribute = iFontName
        Case "Size"
            If mDocumentCreated Then
                CheckNewPagePending
                iDPIY = mPages(mCurrentPage).MetaDPIY
            Else
                iDPIY = GetDeviceCaps(VB.Printer.hDC, LOGPIXELSY)
            End If
            iSng = -mLogFont.lfHeight / iDPIY * 72
            GetFontAttribute = iSng
        Case "Strikethrough", "StrikeOut", "Strikethru"
            GetFontAttribute = CBool(mLogFont.lfStrikeOut)
        Case "Underline"
            GetFontAttribute = CBool(mLogFont.lfUnderline)
        Case "Weight"
            GetFontAttribute = mLogFont.lfWeight
        Case "Width"
            If mDocumentCreated Then
                CheckNewPagePending
                iDPIY = mPages(mCurrentPage).MetaDPIY
            Else
                iDPIY = GetDeviceCaps(VB.Printer.hDC, LOGPIXELSY)
            End If
            iSng = mLogFont.lfWidth / iDPIY * 72
            GetFontAttribute = iSng
        Case "Charset"
            GetFontAttribute = mLogFont.lfCharSet
    End Select
    
End Function

Private Sub SetFontAttribute(nAttribute As String, nValue As Variant)
    Dim iFontName As String
    Dim iDPIY As Single
    Dim c As Long
    
    Select Case nAttribute
        Case "Bold"
            If nValue Then
                mLogFont.lfWeight = FW_BOLD
            Else
                mLogFont.lfWeight = FW_NORMAL
            End If
            mFontWeight = mLogFont.lfWeight
            mFontBold = nValue
        Case "Italic"
            If nValue Then
                mLogFont.lfItalic = 1
            Else
                mLogFont.lfItalic = 0
            End If
            mFontItalic = nValue
        Case "Name"
            Dim iBytes() As Byte
            
            If VarType(nValue) <> vbString Then Exit Sub
            iFontName = nValue
            iBytes = iFontName
            For c = 0 To 31
                If c < Len(iFontName) Then
                    mLogFont.lfFaceName(c * 2) = iBytes(c * 2) 'Asc(Mid$(iFontName, c + 1, 1))
                    mLogFont.lfFaceName(c * 2 + 1) = iBytes(c * 2 + 1)
                Else
                    mLogFont.lfFaceName(c * 2) = 0
                    mLogFont.lfFaceName(c * 2 + 1) = 0
                End If
            Next
            mFontName = nValue
            mFontCharset = DefaultCharsetForFontName(mFontName)
            mLogFont.lfCharSet = mFontCharset
        Case "Size"
            If mDocumentCreated Then
                CheckNewPagePending
                iDPIY = mPages(mCurrentPage).MetaDPIY
            Else
                iDPIY = GetDeviceCaps(VB.Printer.hDC, LOGPIXELSY)
            End If
            If nValue < 1 Then nValue = 1
            mLogFont.lfHeight = -Round(nValue * iDPIY / 72)
            mFontSize = -mLogFont.lfHeight * 72 / iDPIY
        Case "Strikethrough", "StrikeOut", "Strikethru"
            If nValue Then
                mLogFont.lfStrikeOut = 1
            Else
                mLogFont.lfStrikeOut = 0
            End If
            mFontStrikethrough = nValue
        Case "Underline"
            If nValue Then
                mLogFont.lfUnderline = 1
            Else
                mLogFont.lfUnderline = 0
            End If
            mFontUnderline = nValue
        Case "Weight"
            mLogFont.lfWeight = nValue
            mFontWeight = nValue
        Case "Width"
            If mDocumentCreated Then
                CheckNewPagePending
                iDPIY = mPages(mCurrentPage).MetaDPIY
            Else
                iDPIY = GetDeviceCaps(VB.Printer.hDC, LOGPIXELSY)
            End If
            mLogFont.lfWidth = Round(nValue * iDPIY / 72)
            mFontWidth = nValue
        Case "Charset"
            mLogFont.lfCharSet = nValue
            mFontCharset = nValue
    End Select
End Sub

Private Sub AssignFontVariables()
    mFontItalic = GetFontAttribute("Italic")
    mFontName = GetFontAttribute("Name")
    mFontSize = GetFontAttribute("Size")
    mFontStrikethrough = GetFontAttribute("Strikethrough")
    mFontUnderline = GetFontAttribute("Underline")
    mFontWeight = GetFontAttribute("Weight")
    mFontWidth = GetFontAttribute("Width")
    mFontCharset = GetFontAttribute("Charset")
    mFontBold = GetFontAttribute("Bold")
End Sub

Private Function DefaultCharsetForFontName(nFontName As String) As Long
    Static sCalled As Boolean
    Static sFontNames() As String
    Static sCharsets() As Long
    Dim i As Long
    Dim iFont As New StdFont
    
    If Not sCalled Then
        ReDim sFontNames(0)
        ReDim sCharsets(0)
        sCalled = True
    End If
    i = IndexInList(sFontNames, nFontName)
    If i < 1 Then
        ReDim Preserve sFontNames(UBound(sFontNames) + 1)
        ReDim Preserve sCharsets(UBound(sFontNames))
        sFontNames(UBound(sFontNames)) = nFontName
        iFont.Name = nFontName
        sCharsets(UBound(sFontNames)) = iFont.Charset
        i = UBound(sFontNames)
    End If
    DefaultCharsetForFontName = sCharsets(i)
End Function

Private Sub PrintDocumentTooLongWarning()
    Dim iFontSize As Single
    Dim iFontName As String
    Dim iFontBold As Boolean
    
    mDocumentTooLongWarningPrinted = True
    mPrintingDocumentTooLongWarning = True
    iFontSize = FontSize
    iFontName = FontName
    iFontBold = FontBold
    CurrentX = 0
    CurrentY = 0
    FontName = "Arial"
    FontSize = 12
    FontBold = True
    PrintText "Documento demasiado extenso para imprimir completo, hecho hasta la página " & cMaxPages & ".", Me.LeftMargin, Me.TopMargin, Me.ScaleWidth - Me.LeftMargin - Me.RightMargin
    FontSize = iFontSize
    FontName = iFontName
    FontBold = iFontBold
    mPrintingDocumentTooLongWarning = False
    MsgBox GetLocalizedString(efnGUIStr_cPrinterEx_PrintDocumentTooLongWarning_MsgBoxWarning), vbExclamation, ClientProductName
End Sub

Private Sub mFont_FontChanged(ByVal PropertyName As String)
    Set Font = mFont
End Sub

Public Sub DrawPoint(x As Single, y As Single, Optional nColor As Long = -1, Optional nStep As Boolean)
    Dim iMetaDC As Long
    Dim iPenPrev As Long
    Dim iPen As Long
    Dim iColor As Long
    Dim iX As Single
    Dim iY As Single
    
    EnsureDocumentCreated
    iMetaDC = mPages(mCurrentPage).MetaDC
    If iMetaDC = 0 Then Exit Sub
    
    If DocumentTooLong Then Exit Sub
    
    If nStep Then
        iX = mCurrentXPixels + ScaleX(x, mScaleMode, vbPixels) - mScaleLeftPixels + PageLeftMarginInPixels
        iY = mCurrentYPixels + ScaleY(y, mScaleMode, vbPixels) - mScaleTopPixels + PageTopMarginInPixels
    Else
        iX = ScaleX(x, mScaleMode, vbPixels) - mScaleLeftPixels + PageLeftMarginInPixels
        iY = ScaleY(y, mScaleMode, vbPixels) - mScaleTopPixels + PageTopMarginInPixels
    End If
    
    If nColor <> -1 Then
        TranslateColor CLng(nColor), 0, iColor
        If iColor <> mForeColorTranslated Then
            iPen = CreatePen(mDrawStyle, mDrawWidth, iColor)
            iPenPrev = SelectObject(iMetaDC, iPen)
        End If
    End If
    
    MoveToEx iMetaDC, iX - 1, iY - 1, ByVal 0&
    LineTo iMetaDC, iX, iY
    
    If iPenPrev <> 0 Then
        Call SelectObject(iMetaDC, iPenPrev)
    End If
    If iPen <> 0 Then
        DeleteObject iPen
    End If
    mCurrentXPixels = iX
    mCurrentYPixels = iY
    SomethingPrintedOnPage = True
End Sub

Public Function GetPageNumber(ByVal nPageIndex As Long) As Long
    If nPageIndex < 1 Or nPageIndex > PageCount Then
        RaiseError 9, "PrinterEx"
        Exit Function
    End If
    GetPageNumber = mPageNumbers(nPageIndex)
End Function

Public Property Get CurrentPageNumber() As Long
    CurrentPageNumber = GetPageNumber(mCurrentPage)
End Property

Public Function GetDocPageCount(Optional ByVal nPageIndex As Long) As Long
    Dim iTo As Long
    
    If nPageIndex < 0 Then Exit Function
    If nPageIndex = 0 Then nPageIndex = mCurrentPage
    If nPageIndex > PageCount Then nPageIndex = PageCount
    
    iTo = mPageNumbers(nPageIndex)
    If nPageIndex < PageCount Then
        Do Until nPageIndex + 1 > PageCount
            nPageIndex = nPageIndex + 1
            If mPageNumbers(nPageIndex) <= iTo Then Exit Do
            iTo = mPageNumbers(nPageIndex)
        Loop
    End If
    
    GetDocPageCount = iTo
End Function

Public Function GetDocFirstPageIndex(Optional ByVal nPageIndex As Long) As Long
    Dim iPageNumber As Long
    
    If nPageIndex < 0 Then Exit Function
    If mCurrentPage < 0 Then Exit Function
    If nPageIndex = 0 Then nPageIndex = mCurrentPage
    If nPageIndex > PageCount Then nPageIndex = PageCount
    
    GetDocFirstPageIndex = nPageIndex
    iPageNumber = mPageNumbers(nPageIndex)
    Do
        nPageIndex = nPageIndex - 1
        If nPageIndex < 1 Then Exit Do
        If mPageNumbers(nPageIndex) >= iPageNumber Then Exit Do
        iPageNumber = mPageNumbers(nPageIndex)
        GetDocFirstPageIndex = nPageIndex
    Loop
End Function

Public Function GetDocLastPageIndex(Optional ByVal nPageIndex As Long) As Long
    Dim iPageNumber As Long
    
    If nPageIndex < 0 Then Exit Function
    If mCurrentPage < 0 Then Exit Function
    If nPageIndex = 0 Then nPageIndex = mCurrentPage
    If nPageIndex > PageCount Then nPageIndex = PageCount
    
    GetDocLastPageIndex = nPageIndex
    iPageNumber = mPageNumbers(nPageIndex)
    Do
        nPageIndex = nPageIndex + 1
        If nPageIndex = PageCount + 1 Then Exit Do
        If mPageNumbers(nPageIndex) <= iPageNumber Then Exit Do
        iPageNumber = mPageNumbers(nPageIndex)
        GetDocLastPageIndex = nPageIndex
    Loop
End Function

Private Sub EndDoc()
    Dim iHdc As Long
    
    If gPrinterExFromPrintFn > 0 Then
        
        If PageCount > 0 Then
            If Not SomethingPrintedOnPage Then KillLastPage
            If PageCount > 0 Then
                RaiseEvent_EndDoc1 GetDocFirstPageIndex, GetDocLastPageIndex
                NewPage
                mNextPageNumberIsOne = True
            End If
        End If
    Else
        If PageCount > 0 Then
            If Not SomethingPrintedOnPage Then KillLastPage
            RaiseEvent_EndDoc1 1, PageCount
            
            mDlg.PaperSize = mPaperSize
            mDlg.PaperBin = mPaperBin
            mDlg.Duplex = mDuplex
            mDlg.Orientation = mOrientation
            mDlg.PrintQuality = mPrintQuality
            mDlg.ColorMode = mColorMode
            mDlg.UpdateDevModeWithCurrentSettings
            
            DevModePtr = mDlg.DevModePtr
            If mDlg.hDC <> 0 Then
                PrintDocumentToPrinterDC mDlg.hDC
            Else
                iHdc = CreateDCByLong("WINSPOOL", mDeviceName, vbNullString, ByVal mDevModePtr)
                If iHdc <> 0 Then
                    PrintDocumentToPrinterDC iHdc
                End If
                DeleteDC iHdc
            End If
            KillDoc
            If gPrinterExDeviceName = "" Then
                DeviceName = VB.Printer.DeviceName
            Else
                DeviceName = gPrinterExDeviceName
            End If
        End If
    End If
End Sub

Public Sub DrawCircle(x As Single, y As Single, Radius As Single, Optional Aspect As Single, Optional Color As Long = -1, Optional StartArc, Optional EndArc, Optional Step As Boolean)
    Dim iXStartArc As Long
    Dim iYStartArc As Long
    Dim iXEndArc As Long
    Dim iYEndArc As Long
    Dim iAspectX As Single
    Dim iAspectY As Single
    Dim iStartArc As Single
    Dim iEndArc As Single
    Dim iDontDraw As Boolean
    Dim iFillShape As Boolean
    Dim iMetaDC As Long
    Dim iX As Single
    Dim iY As Single
    Dim iFT As Boolean
    Dim iColor As Long
    Dim iPen As Long
    Dim iPenPrev As Long
    Dim iRadius As Single
    
    EnsureDocumentCreated
    iMetaDC = mPages(mCurrentPage).MetaDC
    If iMetaDC = 0 Then Exit Sub
    If DocumentTooLong Then Exit Sub
    
    If Step Then
        iX = mCurrentXPixels + ScaleX(x, mScaleMode, vbPixels) - mScaleLeftPixels + PageLeftMarginInPixels
        iY = mCurrentYPixels + ScaleY(y, mScaleMode, vbPixels) - mScaleTopPixels + PageTopMarginInPixels
    Else
        iX = ScaleX(x, mScaleMode, vbPixels) - mScaleLeftPixels + PageLeftMarginInPixels
        iY = ScaleY(y, mScaleMode, vbPixels) - mScaleTopPixels + PageTopMarginInPixels
    End If
    
    iRadius = ScaleX(Radius, mScaleMode, vbPixels)
    
    If Color <> -1 Then
        TranslateColor CLng(Color), 0, iColor
        If iColor <> mForeColorTranslated Then
            iPen = CreatePen(mDrawStyle, mDrawWidth, iColor)
            iPenPrev = SelectObject(iMetaDC, iPen)
        End If
    End If
    
    If Aspect = 0 Then Aspect = 1
    
    If IsMissing(StartArc) And IsMissing(EndArc) Then
        If mFillStyle = vbSolid Then
            iFillShape = True
        End If
    End If

    If Aspect > 1 Then
        iAspectX = 1 / Aspect
        iAspectY = 1
    Else
        iAspectX = 1
        iAspectY = 1 * Aspect
    End If

    If IsMissing(StartArc) Then
        iStartArc = 0
    Else
        iStartArc = StartArc
    End If
    If IsMissing(EndArc) Then
        iEndArc = 0 ' 2 * pi
    Else
        iEndArc = EndArc
    End If

    If Not IsMissing(EndArc) Then
        If iStartArc = 0 And iEndArc >= 6.2768 Then iDontDraw = True
    End If
    
    If Not iDontDraw Then
        iXStartArc = iRadius * iAspectX * Cos(iStartArc) + iX
        iYStartArc = iRadius * iAspectY * sIn(iStartArc) * -1 + iY
        iXEndArc = iRadius * iAspectX * Cos(iEndArc) + iX
        iYEndArc = iRadius * iAspectY * sIn(iEndArc) * -1 + iY
        
        If iFillShape Then
            If mFillStyle <> vbFSTransparent Then
                If mFillStyle <> vbFSSolid Then
                    If Not mFontTransparent Then
                        mFontTransparent = True
                        SetBackMode
                        iFT = True
                    End If
                End If
            End If
            Ellipse iMetaDC, iX - iRadius * iAspectX, iY - iRadius * iAspectY, iX + iRadius * iAspectX, iY + iRadius * iAspectY
            If iFT Then
                mFontTransparent = False
                SetBackMode
            End If
        Else
            Arc iMetaDC, iX - iRadius * iAspectX, iY - iRadius * iAspectY, iX + iRadius * iAspectX, iY + iRadius * iAspectY, iXStartArc, iYStartArc, iXEndArc, iYEndArc
        End If
    End If

    If iPenPrev <> 0 Then
        Call SelectObject(iMetaDC, iPenPrev)
    End If
    If iPen <> 0 Then
        DeleteObject iPen
    End If

    mCurrentXPixels = iX
    mCurrentYPixels = iY
    SomethingPrintedOnPage = True
End Sub

Private Function DocumentTooLong() As Boolean
    If mCurrentPage > cMaxPages Then
        If Not mDocumentTooLongWarningPrinted Then
            PrintDocumentTooLongWarning
        End If
        DocumentTooLong = True
    End If
End Function

Public Sub PaintPicture(Picture As stdole.IPictureDisp, X1 As Single, Y1 As Single, Optional Width1 As Variant, Optional Height1 As Variant, Optional X2 As Variant, Optional Y2 As Variant, Optional Width2 As Variant, Optional Height2 As Variant, Optional Opcode As Variant)
    Dim iWidth1 As Long
    Dim iHeight1 As Long
    Dim iWidth2 As Long
    Dim iHeight2 As Long
    Dim iOpcode As Long
    Dim iTransparentColor As Long
    Dim iDrawTransparent As Boolean
    Dim iStretchMode As Long
    Dim iDrawStretched As Boolean
    Dim iX1 As Long
    Dim iY1 As Long
    Dim iX2 As Long
    Dim iY2 As Long
    Dim iMetaDC As Long
    Dim iTempDC As Long
    Dim iOldObject As Long
    Dim iPic As IPicture
    Dim iHigherOpcodeFlags As Long
    Dim iQualityStretch As Boolean
    Dim iPic2 As StdPicture
    Const cHalfTone = 4
    Dim iMirroring As Boolean
    Dim iX2ScrDC As Long ' in screen pixels
    Dim iY2ScrDC As Long ' in screen pixels
    Dim iWidth2ScrDC As Long ' in screen pixels
    Dim iHeight2ScrDC As Long ' in screen pixels
    
    ' error checking
    If Not IsObject(Picture) Then
        RaiseError 424, "PrinterEx" ', "Object required."
        Exit Sub
    ElseIf Picture Is Nothing Then
        RaiseError 91, "PrinterEx" ', "Object variable not set."
        Exit Sub
    ElseIf TypeName(Picture) <> "Picture" Then
        RaiseError 13, "PrinterEx" ', "Type mismatch."
        Exit Sub
    ElseIf Picture.Handle = 0 Then
        RaiseError 481, "PrinterEx" ', "Invalid picture."
        Exit Sub
    End If
    
    On Error GoTo Err_Exit:
    
    EnsureDocumentCreated
    iMetaDC = mPages(mCurrentPage).MetaDC
    If iMetaDC = 0 Then Exit Sub
    If DocumentTooLong Then Exit Sub
    
    Set iPic = Picture
    
    ' prepare parameters
    iX1 = ScaleX(X1, mScaleMode, vbPixels) - mScaleLeftPixels + PageLeftMarginInPixels
    iY1 = ScaleY(Y1, mScaleMode, vbPixels) - mScaleTopPixels + PageTopMarginInPixels
    If IsMissing(X2) Then
        iX2 = 0
        iX2ScrDC = 0
    Else
        iX2 = ScaleX(CSng(X2), ScaleMode, vbPixels)
        iX2ScrDC = ScaleX(CSng(X2), ScaleMode, vbTwips) / Screen.TwipsPerPixelX
    End If
    If IsMissing(Y2) Then
        iY2 = 0
        iY2ScrDC = 0
    Else
        iY2 = ScaleY(CSng(Y2), ScaleMode, vbPixels)
        iY2ScrDC = ScaleY(CSng(Y2), ScaleMode, vbTwips) / Screen.TwipsPerPixelY
    End If
    If IsMissing(Width2) Then
        iWidth2 = ScaleX(iPic.Width, vbHimetric, vbPixels) - iX2
        iWidth2ScrDC = iPic.Width / 2540 * 1440 / Screen.TwipsPerPixelX - iX2ScrDC
    Else
        iWidth2 = ScaleX(CSng(Width2), Me.ScaleMode, vbPixels) - iX2
        iWidth2ScrDC = ScaleX(CSng(Width2), ScaleMode, vbTwips) / Screen.TwipsPerPixelX - iX2ScrDC
    End If
    If IsMissing(Height2) Then
        iHeight2 = ScaleY(iPic.Height, vbHimetric, vbPixels) - iY2
        iHeight2ScrDC = iPic.Height / 2540 * 1440 / Screen.TwipsPerPixelY - iY2ScrDC
    Else
        iHeight2 = ScaleY(CSng(Height2), Me.ScaleMode, vbPixels) - iY2
        iHeight2ScrDC = ScaleY(CSng(Height2), ScaleMode, vbTwips) / Screen.TwipsPerPixelY - iY2ScrDC
    End If
    If (iX1 + iWidth2) > CurrentPageScaleWidthInPixels Then
        iWidth2 = CurrentPageScaleWidthInPixels - iX1
        iWidth2ScrDC = ScaleX(CSng(iWidth2), vbPixels, vbTwips) / Screen.TwipsPerPixelX
    End If
    If (iY1 + iHeight2) > CurrentPageScaleHeightInPixels Then
        Height2 = CurrentPageScaleHeightInPixels - iY1
        iHeight2ScrDC = ScaleY(CSng(iHeight2), vbPixels, vbTwips) / Screen.TwipsPerPixelY
    End If
    If IsMissing(Width1) Then
        iWidth1 = iWidth2
    Else
        iWidth1 = ScaleX(CSng(Width1), Me.ScaleMode, vbPixels)
    End If
    If IsMissing(Height1) Then
        iHeight1 = iHeight2
    Else
        iHeight1 = ScaleY(CSng(Height1), Me.ScaleMode, vbPixels)
    End If
    If IsMissing(Opcode) Then
        iOpcode = vbSrcCopy
    Else
        iOpcode = CLng(Opcode)
    End If
    iMirroring = (iWidth1 < 0) Or (iHeight1 < 0) Or (iWidth2 < 0) Or (iHeight2 < 0)
    If (iWidth1 = 0) Or (iHeight1 = 0) Or (iWidth2 = 0) Or (iHeight2 = 0) Then Exit Sub
    
    iHigherOpcodeFlags = iOpcode And &HF000000
    iOpcode = iOpcode And Not &HF000000
    If ((iHigherOpcodeFlags And vbTransparentColor) <> 0) And Not iMirroring Then ' not transparent color
        iDrawTransparent = True
    End If
    iHigherOpcodeFlags = iHigherOpcodeFlags And Not vbTransparentColor
    iDrawStretched = Not ((iWidth2 = iWidth1) And (iHeight2 = iHeight1)) Or iMirroring
    If iDrawStretched Or iDrawTransparent Then
        iStretchMode = (iHigherOpcodeFlags And &HF000000) / &H1000000
        If iStretchMode = 0 Then iStretchMode = cHalfTone
        If iStretchMode = 5 Then
            iQualityStretch = True
        End If
        If iStretchMode < 0 Or iStretchMode > cHalfTone Then
            iStretchMode = cHalfTone
        End If
    End If
    If iOpcode = 0 Then iOpcode = vbSrcCopy
    
    ' so we have: iPic.Type, iDrawStretched, iStretchMode, iQualityStretch and iDrawTransparent
    
    If (iPic.Type = vbPicTypeBitmap) And iQualityStretch And (Not iDrawTransparent) Then
        iPic.SelectPicture 0&, iOldObject, 0&
        Set iPic2 = ResampleBMP(iPic, ScaleX(CSng(iWidth1), vbPixels, vbTwips) / Screen.TwipsPerPixelX, ScaleY(CSng(iHeight1), vbPixels, vbTwips) / Screen.TwipsPerPixelY)
        iX2ScrDC = iX2
        iY2ScrDC = iY2
        iWidth2ScrDC = ScaleX(CSng(iWidth1), vbPixels, vbTwips) / Screen.TwipsPerPixelX
        iHeight2ScrDC = ScaleY(CSng(iHeight1), vbPixels, vbTwips) / Screen.TwipsPerPixelY
        iPic.SelectPicture iOldObject, 0&, 0&
        iOldObject = 0
        Set iPic = iPic2
    End If
    If (iPic.Type = vbPicTypeBitmap) And (iDrawStretched And (iOpcode <> vbSrcCopy)) Or iDrawTransparent Then
        If GetStretchBltMode(iTempDC) <> iStretchMode Then
            SetStretchBltMode iTempDC, iStretchMode
        End If
    End If
    If (iPic.Type = vbPicTypeBitmap) And iDrawTransparent Then
        iTempDC = CreateCompatibleDC(iMetaDC)
        iPic.SelectPicture iTempDC, iOldObject, 0&
        iTransparentColor = GetPixel(iTempDC, 0, 0)
        TransparentBlt iMetaDC, iX1, iY1, iWidth1, iHeight1, iTempDC, iX2ScrDC, iY2ScrDC, iWidth2ScrDC, iHeight2ScrDC, iTransparentColor
    ElseIf (iPic.Type = vbPicTypeBitmap) And (iOpcode <> vbSrcCopy) Or iQualityStretch Then
        iTempDC = CreateCompatibleDC(iMetaDC)
        iPic.SelectPicture iTempDC, iOldObject, 0&
        StretchBlt iMetaDC, iX1, iY1, iWidth1, iHeight1, iTempDC, iX2ScrDC, iY2ScrDC, iWidth2ScrDC, iHeight2ScrDC, iOpcode
    Else
        iPic.Render iMetaDC, iX1, iY1, iWidth1, iHeight1, ScaleX(CSng(iX2), vbPixels, vbHimetric), ScaleY(CSng(iHeight2 - iY2), vbPixels, vbHimetric), ScaleX(CSng(iWidth2), vbPixels, vbHimetric), ScaleY(-CSng(iHeight2), vbPixels, vbHimetric), 0&
    End If
    
    If iTempDC <> 0 Then
        iPic.SelectPicture iOldObject, iOldObject, 0&
        DeleteDC iTempDC
        iTempDC = 0
    End If
    Exit Sub
    
Err_Exit:
    On Error GoTo -1
    On Error Resume Next
    If iTempDC <> 0 Then
        iPic.SelectPicture iOldObject, iOldObject, 0&
        DeleteDC iTempDC
    End If
End Sub

' IVBPrint interface

Private Property Let IVBPrint_Column(ByVal RHS As Long)
    If RHS > mIVBPrintColumn Then PrintText Space$(RHS - mIVBPrintColumn), , , , , , False
    mIVBPrintColumn = RHS
End Property

Private Property Get IVBPrint_Column() As Long
    IVBPrint_Column = mIVBPrintColumn
End Property

Private Sub IVBPrint_WriteText(ByVal strText As String)
    Dim iCrLfPos As Long
    Dim iStrs() As String
    Dim c As Long
    Dim iCount As Long
    
    iCrLfPos = InStrRev(strText, vbCrLf, , vbBinaryCompare)
    If (iCrLfPos > 0) And (strText <> vbCrLf) Then
        iStrs = Split(strText, vbCrLf)
        iCount = UBound(iStrs)
        For c = 0 To iCount
            If c = iCount Then
                PrintText iStrs(c), , , , , , False
            Else
                PrintText iStrs(c)
            End If
        Next c
    Else
        If (strText <> vbCrLf) Then
            PrintText strText, , , , , , False, False, False
        End If
    End If
    
    If iCrLfPos Then
        mIVBPrintColumn = Len(strText) - iCrLfPos - 1
        If strText = vbCrLf Then
            PrintText vbNullString
        End If
    Else
        mIVBPrintColumn = mIVBPrintColumn + Len(strText)
    End If
End Sub

' End IVBPrint interface

Public Property Let Zoom(nValue As Long)
    If mDocumentCreated And (Not mNewPagePending) Then
        If mPages(mCurrentPage).SomethingPrintedOnPage Then
            RaiseError 396, "PrinterEx", "Zoom property cannot be set within a page"
            Exit Property
        Else
            KillLastPage
        End If
    End If
    If nValue = 0 Then
        RaiseError 380, "PrinterEx"
        Exit Property
    ElseIf nValue > 1000 Then
        Exit Property
    End If
    If nValue <> mZoom Then
'        mLeftMarginInMillimeters = mLeftMarginInMillimeters * mZoom / nValue
 '       mTopMarginInMillimeters = mTopMarginInMillimeters * mZoom / nValue
  '      mRightMarginInMillimeters = mRightMarginInMillimeters * mZoom / nValue
   '     mBottomMarginInMillimeters = mBottomMarginInMillimeters * mZoom / nValue
        mZoom = nValue
    End If
End Property

Public Property Get Zoom() As Long
    Zoom = mZoom
End Property


Private Property Get TrackDefault() As Boolean
    TrackDefault = mTrackDefault
End Property

Private Property Let TrackDefault(ByVal nValue As Boolean)
    If nValue <> mTrackDefault Then
        mTrackDefault = nValue
        If mTrackDefault Then
            gPrinterExDeviceName = ""
        Else
            gPrinterExDeviceName = mDeviceName
        End If
        gPrinterExNotTrackDefault = Not mTrackDefault
    End If
End Property

Public Sub RaiseEvent_EndDoc(FirstPageIndex As Long, LastPageIndex As Long)
    If Not SomethingPrintedOnPage Then KillLastPage
    RaiseEvent_EndDoc1 FirstPageIndex, LastPageIndex
End Sub

Private Sub RaiseEvent_EndDoc1(FirstPageIndex As Long, LastPageIndex As Long)
    Dim iCurrentPage_Prev As Long
    Dim iFont_Prev As StdFont
    Dim iFontName_Prev As String
    Dim iFontSize_Prev As Single
    Dim iFontBold_prev As Boolean
    Dim iFontItalic_Prev As Boolean
    Dim iFontStrikethrough_Prev As Boolean
    Dim iFontUnderline_Prev As Boolean
    Dim iFontWidth_Prev As Long
    Dim iDrawMode_Prev As Long
    Dim iDrawStyle_Prev As Long
    Dim iDrawWidth_Prev As Long
    Dim iForeColor_Prev As Long
    Dim iFontTransparent_Prev As Boolean
    Dim iFillColor_Prev As Long
    Dim iFillStyle_Prev As Long
    Dim iHandleMargins_Prev As Boolean
    Dim iCurrentXPixels_Prev As Long
    Dim iCurrentYPixels_Prev As Long
    
    If (FirstPageIndex = mLastEndDoc_FirstPageIndex) And (LastPageIndex = mLastEndDoc_LastPageIndex) Then Exit Sub
    
    ' raise event EndDoc
    iCurrentPage_Prev = mCurrentPage
    Set iFont_Prev = Font
    iHandleMargins_Prev = mHandleMargins
    
    iFontName_Prev = mFontName
    iFontSize_Prev = mFontSize
    iFontBold_prev = mFontBold
    iFontItalic_Prev = mFontItalic
    iFontStrikethrough_Prev = mFontStrikethrough
    iFontUnderline_Prev = mFontUnderline
    iFontWidth_Prev = mFontWeight
    
    iDrawMode_Prev = mDrawMode
    iDrawStyle_Prev = mDrawStyle
    iDrawWidth_Prev = mDrawWidth
    iForeColor_Prev = mForeColor
    iFontTransparent_Prev = mFontTransparent
    iFillColor_Prev = mFillColor
    iFillStyle_Prev = mFillStyle
    
    iCurrentXPixels_Prev = mCurrentXPixels
    iCurrentYPixels_Prev = mCurrentYPixels
    
    If Not mDisableEvents Then
        mPrinterExEvents.RaiseEvent_EndDoc FirstPageIndex, LastPageIndex, mDocKey
    End If
    
    If mPrintPageNumbers Or (gPrinterExPageFixedText <> "") Or Not (gPrinterExPageFixedImage Is Nothing) Then PrintPageNumbers1 FirstPageIndex, LastPageIndex
    
'    mCurrentPage = iCurrentPage_Prev
'    If mCurrentPage < PageCount Then mCurrentPage = PageCount
    If iCurrentPage_Prev > PageCount Then iCurrentPage_Prev = PageCount
    GoToPage iCurrentPage_Prev
    
    If mHandleMargins <> iHandleMargins_Prev Then HandleMargins = iHandleMargins_Prev
    mCurrentXPixels = iCurrentXPixels_Prev
    mCurrentYPixels = iCurrentYPixels_Prev
    If Not mFont Is iFont_Prev Then Set Font = iFont_Prev
    
    If mFontName <> iFontName_Prev Then FontName = iFontName_Prev
    If mFontSize <> iFontSize_Prev Then FontSize = iFontSize_Prev
    If mFontBold <> iFontBold_prev Then FontBold = iFontBold_prev
    If mFontItalic <> iFontItalic_Prev Then FontItalic = iFontItalic_Prev
    If mFontStrikethrough <> iFontStrikethrough_Prev Then FontStrikeThru = iFontStrikethrough_Prev
    If mFontUnderline <> iFontUnderline_Prev Then FontUnderLine = iFontUnderline_Prev
    If mFontWeight <> iFontWidth_Prev Then mFont.Weight = iFontWidth_Prev
    
    If mDrawMode <> iDrawMode_Prev Then DrawMode = iDrawMode_Prev
    If mDrawStyle <> iDrawStyle_Prev Then DrawStyle = iDrawStyle_Prev
    If mDrawWidth <> iDrawWidth_Prev Then DrawWidth = iDrawWidth_Prev
    If mForeColor <> iForeColor_Prev Then ForeColor = iForeColor_Prev
    If mFontTransparent <> iFontTransparent_Prev Then FontTransparent = iFontTransparent_Prev
    If mFillColor <> iFillColor_Prev Then FillColor = iFillColor_Prev
    If mFillStyle <> iFillStyle_Prev Then FillStyle = iFillStyle_Prev
    
    mLastEndDoc_FirstPageIndex = FirstPageIndex
    mLastEndDoc_LastPageIndex = LastPageIndex
End Sub


Public Property Let MinLeftMargin(ByVal nValue As Single)
    If nValue < mNonPrintableAreaLeftInMillimeters Then nValue = mNonPrintableAreaLeftInMillimeters
    If nValue <> mDlg.MinLeftMargin Then
        mDlg.MinLeftMargin = nValue
        If LeftMargin < nValue Then
            LeftMargin = nValue
        End If
    End If
End Property

Public Property Get MinLeftMargin() As Single
    MinLeftMargin = mDlg.MinLeftMargin
End Property


Public Property Let MinTopMargin(ByVal nValue As Single)
    If nValue < mNonPrintableAreaTopInMillimeters Then nValue = mNonPrintableAreaTopInMillimeters
    If nValue <> mDlg.MinTopMargin Then
        mDlg.MinTopMargin = nValue
        If TopMargin < nValue Then
            TopMargin = nValue
        End If
    End If
End Property

Public Property Get MinTopMargin() As Single
    MinTopMargin = mDlg.MinTopMargin
End Property


Public Property Let MinRightMargin(ByVal nValue As Single)
    If nValue < mNonPrintableAreaRightInMillimeters Then nValue = mNonPrintableAreaRightInMillimeters
    If nValue <> mDlg.MinRightMargin Then
        mDlg.MinRightMargin = nValue
        If RightMargin < nValue Then
            RightMargin = nValue
        End If
    End If
End Property

Public Property Get MinRightMargin() As Single
    MinRightMargin = mDlg.MinRightMargin
End Property


Public Property Let MinBottomMargin(ByVal nValue As Single)
    If nValue < mNonPrintableAreaBottomInMillimeters Then nValue = mNonPrintableAreaBottomInMillimeters
    If nValue <> mDlg.MinBottomMargin Then
        mDlg.MinBottomMargin = nValue
        If BottomMargin < nValue Then
            BottomMargin = nValue
        End If
    End If
End Property

Public Property Get MinBottomMargin() As Single
    MinBottomMargin = mDlg.MinBottomMargin
End Property


Public Property Let Units(nValue As cdeUnits)
    If nValue <> mDlg.Units Then
        mDlg.Units = nValue
    End If
End Property

Public Property Get Units() As cdeUnits
    Units = mDlg.Units
End Property


Public Property Let UnitsForUser(nValue As cdeUnitsForUser)
    If nValue <> mDlg.UnitsForUser Then
        mDlg.UnitsForUser = nValue
    End If
End Property

Public Property Get UnitsForUser() As cdeUnitsForUser
    UnitsForUser = mDlg.UnitsForUser
End Property


Public Property Let PrintPageNumbers(nValue As Boolean)
    If nValue <> mPrintPageNumbers Then
        mPrintPageNumbers = nValue
    End If
End Property

Public Property Get PrintPageNumbers() As Boolean
    PrintPageNumbers = mPrintPageNumbers
End Property


Public Property Set PageNumbersFont(ByVal nFont As StdFont)
    If Not nFont Is mPageNumbersFont Then
        Set mPageNumbersFont = nFont
    End If
End Property

Public Property Get PageNumbersFont() As StdFont
    Set PageNumbersFont = mPageNumbersFont
End Property


Public Property Let PageNumbersForeColor(nValue As OLE_COLOR)
    If nValue <> mPageNumbersForeColor Then
        mPageNumbersForeColor = nValue
    End If
End Property

Public Property Get PageNumbersForeColor() As OLE_COLOR
    PageNumbersForeColor = mPageNumbersForeColor
End Property


Public Property Let PageNumbersPosition(nValue As vbExPageNumbersPositionConstants)
    If nValue <> mPageNumbersPosition Then
        mPageNumbersPosition = nValue
    End If
End Property

Public Property Get PageNumbersPosition() As vbExPageNumbersPositionConstants
    PageNumbersPosition = mPageNumbersPosition
End Property


Public Property Let PageNumbersFormat(nValue As String)
    If nValue <> mPageNumbersFormat Then
        mPageNumbersFormat = nValue
    End If
End Property

Public Property Get PageNumbersFormat() As String
    PageNumbersFormat = mPageNumbersFormat
End Property


Public Property Let DisableEvents(nValue As Boolean)
    mDisableEvents = nValue
End Property

Public Property Get DisableEvents() As Boolean
    DisableEvents = mDisableEvents
End Property


Public Property Let Printed(nValue As Boolean)
    mPrinted = nValue
End Property

Public Property Get Printed() As Boolean
    Printed = mPrinted
End Property


Public Property Let Copies(nValue As Long)
    If nValue <> mDlg.Copies Then
        mDlg.Copies = nValue
    End If
End Property

Public Property Get Copies() As Long
    Copies = mDlg.Copies
End Property


Public Property Let Collate(nValue As Boolean)
    If nValue <> mDlg.Collate Then
        mDlg.Collate = nValue
    End If
End Property

Public Property Get Collate() As Boolean
    Collate = mDlg.Collate
End Property
    
    
Public Property Get SomethingPrintedOnPage() As Boolean
    On Error Resume Next
    SomethingPrintedOnPage = mPages(mCurrentPage).SomethingPrintedOnPage
End Property

Private Property Let SomethingPrintedOnPage(nValue As Boolean)
    If Not mPrintingPageNumbers Then
        On Error Resume Next
        mPages(mCurrentPage).SomethingPrintedOnPage = nValue
        If Err.Number = 0 Then
            mSomethingPrintedOnDoc = True
        End If
    End If
End Property
    
    
Public Sub SetPageFixedElement(Optional nPicture, Optional nText, Optional nPositionTop As Variant)
    Dim iChanged As Boolean
    
    If Not IsMissing(nPicture) Then
        If Not nPicture Is gPrinterExPageFixedImage Then
            Set gPrinterExPageFixedImage = nPicture
            iChanged = True
        End If
    End If
    If Not IsMissing(nText) Then
        If nText <> gPrinterExPageFixedText Then
            gPrinterExPageFixedText = nText
            iChanged = True
        End If
    End If
    If Not IsMissing(nPositionTop) Then
        If nPositionTop <> gPrinterExPageFixedElementPositionTop Then
            gPrinterExPageFixedElementPositionTop = nPositionTop
            iChanged = True
        End If
    End If
    If iChanged Then
        CheckMarginsLimits
    End If
End Sub


Private Sub CheckMarginsLimits()
    CheckTopMarginLimit
    CheckBottomMarginLimit
End Sub

Private Sub CheckBottomMarginLimit()
    Dim iFixedElement As Boolean
    Dim iAuxDocCreated As Boolean
    Dim iAuxFont As StdFont
    Dim iFixedImage As Boolean
    Dim iPageNumbers As Boolean
    Dim iAuxMinBottomMarginMillimeters As Single
    Dim iFixedTextHeight As Single
    Dim iRequiredHeight As Single
    Dim iHM As Boolean
    Dim iSM As Long
    Dim iFixedText As Boolean
    Dim iPageNumbersWidth As Single
    Dim iFixedAvailableWidth As Single
    Dim iImageHeight As Single
    Dim iImageWidth As Single
    
    If mCheckingMarginsLimits Then Exit Sub
    mCheckingMarginsLimits = True
    If Not gPrinterExPageFixedElementPositionTop Then
        iFixedImage = (Not gPrinterExPageFixedImage Is Nothing)
        If iFixedImage Then iFixedImage = gPrinterExPageFixedImage.Handle <> 0
        If Not iFixedImage Then
            iFixedText = gPrinterExPageFixedText <> ""
        End If
        iFixedElement = iFixedImage Or iFixedText
    End If
    
    iPageNumbers = mPrintPageNumbers And ((mPageNumbersPosition = vxPositionBottomLeft) Or (mPageNumbersPosition = vxPositionBottomCenter) Or (mPageNumbersPosition = vxPositionBottomRight))
    
    If iPageNumbers Or iFixedElement Then
        iHM = mHandleMargins
        iSM = mScaleMode
        ScaleMode = vbMillimeters
        HandleMargins = False
        If mNonPrintableAreaBottomInMillimeters < 6 Then
            iAuxMinBottomMarginMillimeters = 12 * 100 / mZoom
        Else
            iAuxMinBottomMarginMillimeters = (6 + mNonPrintableAreaBottomInMillimeters) * 100 / mZoom
        End If
        If iPageNumbers Or iFixedElement Then
            iAuxDocCreated = mDocumentCreated
            mDisableEvents = True
            EnsureDocumentCreated
            Set iAuxFont = Font
            Set Font = mPageNumbersFont
            FontSize = mPageNumbersFont.Size * 100 / mZoom
            iRequiredHeight = 0
            If iPageNumbers Then
                iRequiredHeight = Printer_TextHeight("Atj")
                iPageNumbersWidth = Printer_TextWidth(GetFormattedPageNumberString(mPageNumbersFormat, 10, 100))
            End If
            iFixedAvailableWidth = PageWidth - PageLeftMargin - PageRightMargin - 10 * 100 / mZoom - iPageNumbersWidth
            If iFixedImage Then
                iImageWidth = ScaleX(gPrinterExPageFixedImage.Width, vbHimetric, mScaleMode) * 100 / mZoom
                iImageHeight = ScaleY(gPrinterExPageFixedImage.Height, vbHimetric, mScaleMode) * 100 / mZoom
                If iImageWidth > iFixedAvailableWidth Then
                    mFixedImageScale = (iFixedAvailableWidth / iImageWidth)
                    iImageHeight = iImageHeight * mFixedImageScale
                    iImageWidth = iImageWidth * mFixedImageScale
                Else
                    mFixedImageScale = 1
                End If
                If iImageHeight > (PageHeight / 5) - 10 / mZoom * 100 Then
                    mFixedImageScale = mFixedImageScale * ((PageHeight / 5) - 10 / mZoom * 100) / iImageHeight
                    iImageHeight = iImageHeight * mFixedImageScale
                    iImageWidth = iImageWidth * mFixedImageScale
                End If
                If iImageHeight > iRequiredHeight Then iRequiredHeight = iImageHeight
            ElseIf iFixedText Then
                iFixedTextHeight = PrintText(gPrinterExPageFixedText, 0, 0, iFixedAvailableWidth, , , , , , , , , True)
                If iFixedTextHeight > iRequiredHeight Then iRequiredHeight = iFixedTextHeight
            End If
            iRequiredHeight = iRequiredHeight + 10 / mZoom * 100
            Set Font = iAuxFont
            If iRequiredHeight > iAuxMinBottomMarginMillimeters Then
                iAuxMinBottomMarginMillimeters = iRequiredHeight
                If iAuxMinBottomMarginMillimeters > PageHeight / 5 Then
                    iAuxMinBottomMarginMillimeters = PageHeight / 5
                End If
            End If
            If Not iAuxDocCreated Then KillDoc
            mDisableEvents = False
        End If
        iAuxMinBottomMarginMillimeters = iAuxMinBottomMarginMillimeters / 100 * mZoom
        If mBottomMarginSet + 0.01 < iAuxMinBottomMarginMillimeters Then
            BottomMargin = iAuxMinBottomMarginMillimeters
        ElseIf TopMargin <> mTopMarginSet Then
            BottomMargin = mBottomMarginSet
        End If
        mScaleMode = iSM
        HandleMargins = iHM
    End If
    mCheckingMarginsLimits = False
End Sub

Private Sub CheckTopMarginLimit()
    Dim iFixedElement As Boolean
    Dim iAuxDocCreated As Boolean
    Dim iAuxFont As StdFont
    Dim iFixedImage As Boolean
    Dim iPageNumbers As Boolean
    Dim iAuxMinTopMarginMillimeters As Single
    Dim iFixedTextHeight As Single
    Dim iRequiredHeight As Single
    Dim iHM As Boolean
    Dim iSM As Long
    Dim iFixedText As Boolean
    Dim iPageNumbersWidth As Single
    Dim iFixedAvailableWidth As Single
    Dim iImageHeight As Single
    Dim iImageWidth As Single
    
    If mCheckingMarginsLimits Then Exit Sub
    mCheckingMarginsLimits = True
    If gPrinterExPageFixedElementPositionTop Then
        iFixedImage = (Not gPrinterExPageFixedImage Is Nothing)
        If iFixedImage Then iFixedImage = gPrinterExPageFixedImage.Handle <> 0
        If Not iFixedImage Then
            iFixedText = gPrinterExPageFixedText <> ""
        End If
        iFixedElement = iFixedImage Or iFixedText
    End If
    
    iPageNumbers = mPrintPageNumbers And ((mPageNumbersPosition = vxPositionTopLeft) Or (mPageNumbersPosition = vxPositionTopCenter) Or (mPageNumbersPosition = vxPositionTopRight))
    
    If iPageNumbers Or iFixedElement Then
        iHM = mHandleMargins
        iSM = mScaleMode
        ScaleMode = vbMillimeters
        HandleMargins = False
        If mNonPrintableAreaTopInMillimeters < 6 Then
            iAuxMinTopMarginMillimeters = 12 * 100 / mZoom
        Else
            iAuxMinTopMarginMillimeters = (6 + mNonPrintableAreaTopInMillimeters) * 100 / mZoom
        End If
        If iPageNumbers Or iFixedElement Then
            iAuxDocCreated = mDocumentCreated
            mDisableEvents = True
            EnsureDocumentCreated
            Set iAuxFont = Font
            Set Font = mPageNumbersFont
            
            FontSize = mPageNumbersFont.Size * 100 / mZoom
            iRequiredHeight = 0
            If iPageNumbers Then
                iRequiredHeight = Printer_TextHeight("Atj")
                iPageNumbersWidth = Printer_TextWidth(GetFormattedPageNumberString(mPageNumbersFormat, 10, 100))
            End If
            iFixedAvailableWidth = PageWidth - PageLeftMargin - PageRightMargin - 10 * 100 / mZoom - iPageNumbersWidth
            If iFixedImage Then
                iImageWidth = ScaleX(gPrinterExPageFixedImage.Width, vbHimetric, mScaleMode) * 100 / mZoom
                iImageHeight = ScaleY(gPrinterExPageFixedImage.Height, vbHimetric, mScaleMode) * 100 / mZoom
                If iImageWidth > iFixedAvailableWidth Then
                    mFixedImageScale = (iFixedAvailableWidth / iImageWidth)
                    iImageHeight = iImageHeight * mFixedImageScale
                    iImageWidth = iImageWidth * mFixedImageScale
                Else
                    mFixedImageScale = 1
                End If
                If iImageHeight > (PageHeight / 5) - 10 / mZoom * 100 Then
                    mFixedImageScale = mFixedImageScale * ((PageHeight / 5) - 10 / mZoom * 100) / iImageHeight
                    iImageHeight = iImageHeight * mFixedImageScale
                    iImageWidth = iImageWidth * mFixedImageScale
                End If
                If iImageHeight > iRequiredHeight Then iRequiredHeight = iImageHeight
            ElseIf iFixedText Then
                iFixedTextHeight = PrintText(gPrinterExPageFixedText, 0, 0, iFixedAvailableWidth, , , , , , , , , True)
                If iFixedTextHeight > iRequiredHeight Then iRequiredHeight = iFixedTextHeight
            End If
            iRequiredHeight = iRequiredHeight + 12 / mZoom * 100
            Set Font = iAuxFont
            If iRequiredHeight > iAuxMinTopMarginMillimeters Then
                iAuxMinTopMarginMillimeters = iRequiredHeight
                If iAuxMinTopMarginMillimeters > PageHeight / 5 Then
                    iAuxMinTopMarginMillimeters = PageHeight / 5
                End If
            End If
            If Not iAuxDocCreated Then KillDoc
            mDisableEvents = False
        End If
        iAuxMinTopMarginMillimeters = iAuxMinTopMarginMillimeters / 100 * mZoom
        If mTopMarginSet + 0.01 < iAuxMinTopMarginMillimeters Then
            TopMargin = iAuxMinTopMarginMillimeters
        ElseIf TopMargin <> mTopMarginSet Then
            TopMargin = mTopMarginSet
        End If
        mScaleMode = iSM
        HandleMargins = iHM
    End If
    mCheckingMarginsLimits = False
End Sub

Private Sub PrintPageNumbers1(nFirstPageIndex As Long, nLastPageIndex As Long)
    Dim c As Long
    Dim iSM As Long
    Dim iHandleMargins As Boolean
    Dim iBottomMarginMillimeters As Single
    Dim iTopMarginMillimeters As Single
    Dim iAuxFont As StdFont
    Dim iPageNumbersFontHeight As Single
    Dim iPnf As String
    Dim iAuxWidth As String
    Dim iM As Single
    Dim iAuxY As Single
    Dim i5Millimeters As Single
    Dim iEndFreeSpace As Single
    Dim i1Millimeter As Single
    Dim iForeColor_Prev As Long
    Dim iAvailableWidth As Single
    Dim iAvailableHeight As Single
    Dim iOneLine As Boolean
    Dim iAuxHeight As Single
    Dim iPageNumbersPositionTop As Boolean
    Dim iPageNumbersPositionRight As Boolean
    Dim iTop As Single
    Dim iPageNumbersWidth As Single
    Dim iFixedImage As Boolean
    Dim iPageCount As Long
    Dim iCurrentPage As Long
    
    mPrintingPageNumbers = True
    iPageCount = nLastPageIndex - nFirstPageIndex + 1
    i5Millimeters = 5 * 100 / mZoom
    iEndFreeSpace = 10 * 100 / mZoom
    i1Millimeter = 1 * 100 / mZoom
    
    iSM = mScaleMode
    iHandleMargins = mHandleMargins
    mScaleMode = vbMillimeters
    HandleMargins = False
    
    Set iAuxFont = Font
    Set Font = mPageNumbersFont
    FontSize = mPageNumbersFont.Size * 100 / mZoom
    iPageNumbersFontHeight = ScaleY(Printer_TextHeight("AAA"), mScaleMode, vbMillimeters)
    iCurrentPage = mCurrentPage
    iForeColor_Prev = ForeColor
    ForeColor = mPageNumbersForeColor
    
    For c = nFirstPageIndex To nLastPageIndex
        GoToPage c
        
        Select Case mPageNumbersPosition
            Case vxPositionTopRight, vxPositionTopLeft, vxPositionTopCenter
                If iEndFreeSpace < PageNonPrintableAreaTop Then
                    iEndFreeSpace = PageNonPrintableAreaTop
                End If
                iPageNumbersPositionTop = True
                iTopMarginMillimeters = PageTopMargin
                CurrentY = iEndFreeSpace - PageNonPrintableAreaTop + 0.5
            Case Else ' Bottom
                If iEndFreeSpace < PageNonPrintableAreaBottom Then
                    iEndFreeSpace = PageNonPrintableAreaBottom
                End If
                iBottomMarginMillimeters = PageBottomMargin
                CurrentY = PageHeight - iPageNumbersFontHeight - iEndFreeSpace - PageNonPrintableAreaTop - 0.5
        End Select
        
        iPnf = GetFormattedPageNumberString(mPageNumbersFormat, GetPageNumber(c), iPageCount) 'GetDocPageCount(c))
        iPageNumbersWidth = TextWidth(iPnf)
        
        Select Case mPageNumbersPosition
            Case vxPositionBottomRight, vxPositionTopRight
                iPageNumbersPositionRight = True
                iM = PageRightMargin
                If iM < i5Millimeters Then iM = i5Millimeters
                If PageWidth - iM - iPageNumbersWidth - i5Millimeters > ScaleWidth - iPageNumbersWidth Then
                    CurrentX = ScaleWidth - iPageNumbersWidth - i1Millimeter
                Else
                    CurrentX = PageWidth - iM - iPageNumbersWidth - i1Millimeter - PageNonPrintableAreaLeft
                End If
            Case vxPositionBottomLeft, vxPositionTopLeft
                iM = PageLeftMargin
                If iM < i5Millimeters Then iM = i5Millimeters
                CurrentX = iM
            Case Else ' vxPositionBottomCenter, vxPositionTopCenter
                CurrentX = (PageWidth - iPageNumbersWidth) / 2
        End Select
        iAuxY = CurrentY
        If mPrintPageNumbers Then
            PrintText iPnf
        End If
        
        iFixedImage = Not (gPrinterExPageFixedImage Is Nothing)
        If iFixedImage Then iFixedImage = gPrinterExPageFixedImage.Handle <> 0
        If iFixedImage Then
            Dim iXPos As Single
            Dim iYPos As Single
            Dim iW As Long
            Dim iH As Long
            
            iW = ScaleX(gPrinterExPageFixedImage.Width, vbHimetric, mScaleMode) * 100 / mZoom
            iH = ScaleY(gPrinterExPageFixedImage.Height, vbHimetric, mScaleMode) * 100 / mZoom
            
            If mPrintPageNumbers And (((mPageNumbersPosition = vxPositionBottomLeft) And (Not gPrinterExPageFixedElementPositionTop)) Or ((mPageNumbersPosition = vxPositionTopLeft)) And gPrinterExPageFixedElementPositionTop) Then
                iXPos = PageWidth - PageRightMargin - iW
            Else
                If iPageNumbersPositionRight Or (gPrinterExPageFixedElementPositionTop <> iPageNumbersPositionTop) Or (Not mPrintPageNumbers) Then
                    iXPos = PageLeftMargin
                Else
                    iXPos = PageWidth - PageRightMargin - iW - 3 * 100 / mZoom
                End If
            End If
            
            If gPrinterExPageFixedElementPositionTop Then
                iYPos = 6 * 100 / mZoom
            Else
                iYPos = PageHeight - PageBottomMargin + 3 * 100 / mZoom
            End If
            PaintPicture gPrinterExPageFixedImage, iXPos, iYPos, iW * mPages(c).FixedImageScale, iH * mPages(c).FixedImageScale
        ElseIf gPrinterExPageFixedText <> "" Then
            FontBold = False
            FontItalic = True
            FontSize = 12 * 100 / mZoom
            If (gPrinterExPageFixedElementPositionTop = iPageNumbersPositionTop) And mPrintPageNumbers Then
                iAvailableWidth = PageWidth - iPageNumbersWidth - 5 * 100 / mZoom - PageLeftMargin - PageRightMargin
            Else
                iAvailableWidth = PageWidth - PageLeftMargin - PageRightMargin
            End If
            iAuxWidth = Printer_TextWidth(gPrinterExPageFixedText)
            If iAuxWidth > iAvailableWidth Then
                iAuxWidth = iAvailableWidth
            Else
                iOneLine = True
            End If
            If iPageNumbersPositionRight Or (gPrinterExPageFixedElementPositionTop <> iPageNumbersPositionTop) Or (Not mPrintPageNumbers) Then
                CurrentX = PageLeftMargin
            Else
                If iOneLine Then
                    CurrentX = PageWidth - PageRightMargin - iAuxWidth - 3 * 100 / mZoom
                Else
                    CurrentX = PageLeftMargin + iPageNumbersWidth + 5 * 100 / mZoom
                End If
            End If
            If iOneLine And (gPrinterExPageFixedElementPositionTop = iPageNumbersPositionTop) Then
                CurrentY = iAuxY
                PrintText gPrinterExPageFixedText
            Else
                If iOneLine Then
                    If gPrinterExPageFixedElementPositionTop Then
                        CurrentY = iEndFreeSpace - PageNonPrintableAreaTop
                    Else
                        CurrentY = PageHeight - iPageNumbersFontHeight - iEndFreeSpace - PageNonPrintableAreaTop
                    End If
                    PrintText gPrinterExPageFixedText
                Else
                    If gPrinterExPageFixedElementPositionTop Then
                        iAvailableHeight = (PageTopMargin - 12 * 100 / mZoom)
                    Else
                        iAvailableHeight = (PageBottomMargin - 12 * 100 / mZoom)
                    End If
                    iAuxHeight = PrintText(gPrinterExPageFixedText, PageLeftMargin, , iAvailableWidth, , , , , , , , , True)
                    Do Until iAuxHeight <= iAvailableHeight
                        FontSize = FontSize * 0.9
                        If FontSize < 1 Then Exit Do
                        iAuxHeight = PrintText(gPrinterExPageFixedText, PageLeftMargin, , iAvailableWidth, , , , , , , , , True)
                    Loop
                    If gPrinterExPageFixedElementPositionTop Then
                        iTop = 6 * 100 / mZoom
                    Else
                        iTop = PageHeight - PageBottomMargin + 3 * 100 / mZoom
                    End If
                    PrintText gPrinterExPageFixedText, CurrentX, iTop, iAvailableWidth, iAvailableHeight + 10 * 100 / mZoom, vxAlignLeftCenter, False
                End If
            End If
            Set Font = PageNumbersFont
            FontSize = PageNumbersFont.Size * 100 / mZoom
        End If
    Next c
    If mCurrentPage <> iCurrentPage Then
        GoToPage iCurrentPage
    End If
    ForeColor = iForeColor_Prev
    Set Font = iAuxFont
    ScaleMode = iSM
    HandleMargins = iHandleMargins
    mPrintingPageNumbers = False
    
End Sub

Public Function GetFormattedPageNumberString(ByVal nPageNumbersFormat As String, nPageNumber As Long, nPageCount As Long) As String
    GetFormattedPageNumberString = nPageNumbersFormat
    If GetFormattedPageNumberString = "Default" Then
        GetFormattedPageNumberString = GetLocalizedString(efnGUIStr_General_PageNumbersFormatString_Default)
    End If
    GetFormattedPageNumberString = GetFormattedPageNumberString & " "
    
    GetFormattedPageNumberString = Replace(GetFormattedPageNumberString, "#", CStr(nPageNumber))
    GetFormattedPageNumberString = Replace(GetFormattedPageNumberString, " N ", " " & nPageCount & " ")
    GetFormattedPageNumberString = Replace(GetFormattedPageNumberString, "/N ", "/" & nPageCount & " ")
    
    If Right$(GetFormattedPageNumberString, 1) = " " Then
        GetFormattedPageNumberString = Left$(GetFormattedPageNumberString, Len(GetFormattedPageNumberString) - 1)
    End If
    GetFormattedPageNumberString = Trim(GetFormattedPageNumberString)
End Function

Private Property Get PageNonPrintableAreaLeft() As Single
    If mNewPagePending Then
        PageNonPrintableAreaLeft = ScaleX(IIf(mOrientation <> vbPRORLandscape, mNonPrintableAreaLeftInPixels, mNonPrintableAreaTopInPixels), vbPixels, mScaleMode) / mZoom * 100
    Else
        PageNonPrintableAreaLeft = ScaleX(mPages(mCurrentPage).NonPrintableAreaLeftInPixels, vbPixels, mScaleMode)
    End If
End Property

Private Property Get PageNonPrintableAreaTop() As Single
    If mNewPagePending Then
        PageNonPrintableAreaTop = ScaleY(IIf(mOrientation <> vbPRORLandscape, mNonPrintableAreaTopInPixels, mNonPrintableAreaLeftInPixels), vbPixels, mScaleMode) / mZoom * 100
    Else
        PageNonPrintableAreaTop = ScaleY(mPages(mCurrentPage).NonPrintableAreaTopInPixels, vbPixels, mScaleMode)
    End If
End Property

Private Property Get PageNonPrintableAreaRight() As Single
    If mNewPagePending Then
        PageNonPrintableAreaRight = ScaleX(IIf(mOrientation <> vbPRORLandscape, mNonPrintableAreaRightInPixels, mNonPrintableAreaBottomInPixels), vbPixels, mScaleMode) / mZoom * 100
    Else
        PageNonPrintableAreaRight = ScaleX(mPages(mCurrentPage).NonPrintableAreaRightInPixels, vbPixels, mScaleMode)
    End If
End Property

Private Property Get PageNonPrintableAreaBottom() As Single
    If mNewPagePending Then
        PageNonPrintableAreaBottom = ScaleY(IIf(mOrientation <> vbPRORLandscape, mNonPrintableAreaBottomInPixels, mNonPrintableAreaRightInPixels), vbPixels, mScaleMode) / mZoom * 100
    Else
        PageNonPrintableAreaBottom = ScaleY(mPages(mCurrentPage).NonPrintableAreaBottomInPixels, vbPixels, mScaleMode)
    End If
End Property

Public Property Get PageWidth() As Single
    If mNewPagePending Then
        PageWidth = ScaleX(IIf(mOrientation <> vbPRORLandscape, PaperWidthInMillimeters, PaperHeightInMillimeters), vbMillimeters, mScaleMode) / mZoom * 100
    Else
        PageWidth = ScaleX(mPages(mCurrentPage).WidthInPixels, vbPixels, mScaleMode)
    End If
End Property

Public Property Get PageHeight() As Single
    If mNewPagePending Then
        PageHeight = ScaleY(IIf(mOrientation <> vbPRORLandscape, PaperHeightInMillimeters, PaperWidthInMillimeters), vbMillimeters, mScaleMode) / mZoom * 100
    Else
        PageHeight = ScaleY(mPages(mCurrentPage).HeightInPixels, vbPixels, mScaleMode)
    End If
End Property

Public Property Get Events() As PrinterExEvents
    Set Events = mPrinterExEvents
    Set gPrinterExEvents = mPrinterExEvents
End Property

Public Sub PrintRichTextBox(nRTB As Object)
    Dim iLeftMargin As Long
    Dim iTopMargin As Long
    Dim iRightMargin As Long
    Dim iBottomMargin As Long
    Dim fr As FormatRange
    Dim rcDrawTo As RECT
    Dim rcMeasure As RECT
    Dim rcPage As RECT
    Dim iTextLength As Long
    Dim iNextCharPosition As Long
    Dim R As Long
    Dim iPPScaleMode As Long
    Dim p As Long
    Dim iCx As Long
    
    iPPScaleMode = Printer_ScaleMode
    Printer_ScaleMode = vbPixels
    If Printer_CurrentX = 0 Then
        ' Start a PrintEx job to get a valid Printer_hDC
        PrintText Space(1), , , , , , False
        Printer_CurrentX = 0
    End If
    Printer_ScaleMode = vbTwips
    
    ' Calculate the Left$, Top, Right$, and Bottom margins
    iCx = Printer_CurrentX
    iLeftMargin = Me.PageLeftMargin + iCx
    iTopMargin = Me.PageTopMargin + Printer_CurrentY
    iRightMargin = Me.PageWidth - Me.PageRightMargin
    iBottomMargin = Me.PageHeight - Me.PageBottomMargin
    
    ' Set printable area rect
    rcPage.Left = 0
    rcPage.Top = 0
    rcPage.Right = Me.PageWidth
    rcPage.Bottom = Me.PageHeight
    
    ' Set rect in which to PrintEx (relative to printable area)
    rcDrawTo.Left = iLeftMargin
    rcDrawTo.Top = iTopMargin
    rcDrawTo.Right = iRightMargin
    rcDrawTo.Bottom = iBottomMargin
    
    rcMeasure.Left = iLeftMargin
    rcMeasure.Top = iTopMargin
    rcMeasure.Right = iRightMargin
    rcMeasure.Bottom = iBottomMargin
    
    ' Set up the PrintEx instructions
    fr.hDC = Printer_hDC   ' Use the same DC for measuring and rendering
    fr.hdcTarget = Printer_hDC  ' Point at Printer2 hDC
    fr.RC = rcDrawTo            ' Indicate the area on page to draw to
    fr.rcPage = rcPage          ' Indicate entire size of page
    fr.chrg.cpMin = 0           ' Indicate start of text through
    fr.chrg.cpMax = -1          ' end of the text
    
    ' Get length of text in nRTB
    iTextLength = Len(nRTB.Text)
    
    ' Loop printing each page until done
    p = 0
    Do
         p = p + 1
         fr.RC = rcMeasure
         Call SendMessage(nRTB.hWnd, EM_FORMATRANGE, 0&, fr)
         rcMeasure = fr.RC
         fr.RC = rcDrawTo
         If p > 1 Then
             fr.RC.Top = Me.PageTopMargin
         End If
         ' PrintEx the page by sending EM_FORMATRANGE message
         iNextCharPosition = SendMessage(nRTB.hWnd, EM_FORMATRANGE, True, fr)
         If iNextCharPosition >= iTextLength Then Exit Do  'If done then exit
         fr.chrg.cpMin = iNextCharPosition ' Starting position for next page
         Printer_NewPage                  ' Move on to next page
         PrintText Space(1), , , , , , False  ' Re-initialize hDC
         Printer_CurrentX = 0
         fr.hDC = Printer_hDC
         fr.hdcTarget = Printer_hDC
         rcDrawTo.Left = Me.PageLeftMargin + iCx
         rcDrawTo.Top = Me.PageTopMargin
    Loop
    
    ' Allow the nRTB to free up memory
    R = SendMessage(nRTB.hWnd, EM_FORMATRANGE, False, ByVal CLng(0))
    Printer_CurrentY = Printer_CurrentY + rcMeasure.Bottom - rcMeasure.Top - Printer_TextHeight("A")
    Printer_ScaleMode = iPPScaleMode
End Sub

Public Property Get DefaultColorMode() As cdeColorModeConstants
    DefaultColorMode = mDefaultColorMode
End Property

Friend Property Let DontCheckError396(nValue As Boolean)
    mDontCheckError396 = nValue
End Property

Friend Property Get DontCheckError396() As Boolean
    DontCheckError396 = mDontCheckError396
End Property


Friend Property Let DocKey(ByVal nDocKey As String)
    If nDocKey <> mDocKey Then
        mDocKey = nDocKey
        LoadUserDocPreferences
    End If
End Property

Friend Property Get DocKey() As String
    DocKey = mDocKey
End Property

Private Sub SaveUserDocPreferences()
    Dim iContextByDevice As String
    Dim iContextGeneral As String
    
    If mDocKey <> "" Then
        iContextByDevice = Base64Encode("Prn" & mDeviceName & "_" & mDocKey)
        iContextGeneral = Base64Encode("Prn" & "_" & mDocKey)
    
        If mSave_DeviceName Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextGeneral & "_DeviceName", mPrintFnObject.CommonDialogExObject.DeviceName
        End If
    
        If mSave_PrintPageNumbers Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextGeneral & "_PrintPageNumbers", CLng(mPrintFnObject.PrintPageNumbers)
        End If
        If mSave_PageNumbersPosition Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextGeneral & "_PageNumbersPosition", mPrintFnObject.PageNumbersPosition
        End If
        If mSave_PageNumbersFormat Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextGeneral & "_PageNumbersFormat", mPrintFnObject.PageNumbersFormat
        End If
        If mSave_PageNumbersFont Then
            SaveSettingFont AppNameForRegistry, "PrintingSettings", iContextGeneral & "_PageNumbersFont", mPrintFnObject.PageNumbersFont
        End If
        If mSave_PageNumbersForeColor Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextGeneral & "_PageNumbersForeColor", mPrintFnObject.PageNumbersForeColor
        End If
    
        If mSave_ColorMode Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextGeneral & "_ColorMode", mPrintFnObject.CommonDialogExObject.ColorMode
        End If
    
        If mSave_LeftMargin Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextGeneral & "_LeftMargin", Str$(mPrintFnObject.CommonDialogExObject.LeftMargin)
        End If
        If mSave_RightMargin Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextGeneral & "_RightMargin", Str$(mPrintFnObject.CommonDialogExObject.RightMargin)
        End If
        If mSave_TopMargin Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextGeneral & "_TopMargin", Str$(mPrintFnObject.CommonDialogExObject.TopMargin)
        End If
        If mSave_BottomMargin Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextGeneral & "_BottomMargin", Str$(mPrintFnObject.CommonDialogExObject.BottomMargin)
        End If
    
        If mSave_ScalePercent Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextByDevice & "_ScalePercent", mPrintFnObject.ScalePercent
        End If
    
        If mSave_Orientation Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextByDevice & "_Orientation", mPrintFnObject.CommonDialogExObject.Orientation
        End If
        If mSave_PaperSize Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextByDevice & "_PaperSize", mPrintFnObject.CommonDialogExObject.PaperSize
        End If
        If mSave_PaperBin Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextByDevice & "_PaperBin", mPrintFnObject.CommonDialogExObject.PaperBin
        End If
        If mSave_PrintQuality Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextByDevice & "_PrintQuality", mPrintFnObject.CommonDialogExObject.PrintQuality
        End If
        If mSave_Collate Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextByDevice & "_Collate", CLng(mPrintFnObject.CommonDialogExObject.Collate)
        End If
        If mSave_Duplex Then
            SaveSetting AppNameForRegistry, "PrintingSettings", iContextByDevice & "_Duplex", mPrintFnObject.CommonDialogExObject.Duplex
        End If
    End If
    
    mSave_DeviceName = False
    mSave_PrintPageNumbers = False
    mSave_PageNumbersPosition = False
    mSave_PageNumbersFormat = False
    mSave_PageNumbersFont = True
    mSave_PageNumbersForeColor = False
    mSave_ScalePercent = False
    mSave_Orientation = False
    mSave_ColorMode = False
    mSave_PaperSize = False
    mSave_PaperBin = False
    mSave_PrintQuality = False
    mSave_Collate = False
    mSave_Duplex = False
    mSave_LeftMargin = False
    mSave_RightMargin = False
    mSave_TopMargin = False
    mSave_BottomMargin = False
    
End Sub

Private Sub LoadUserDocPreferences()
    Dim iContextByDevice As String
    Dim iContextGeneral As String
    Dim iStr As String
    Static sLastDocKey As String
    Dim iDn As String
    
    If (sLastDocKey <> mDocKey) Then
        ' restore originals
        If mLoaded_Orientation Then
            Orientation = mOriginal_Orientation
            mLoaded_Orientation = False
        End If
        If mLoaded_PaperSize Then
            PaperSize = mOriginal_PaperSize
            mLoaded_PaperSize = False
        End If
        If mLoaded_PaperBin Then
            PaperBin = mOriginal_PaperBin
            mLoaded_PaperBin = False
        End If
        If mLoaded_PrintQuality Then
            PrintQuality = mOriginal_PrintQuality
            mLoaded_PrintQuality = False
        End If
        If mLoaded_Collate Then
            Collate = mOriginal_Collate
            mLoaded_Collate = False
        End If
        If mLoaded_Duplex Then
            Duplex = mOriginal_Duplex
            mLoaded_Duplex = False
        End If
        
        If mLoaded_ScalePercent Then
            mPrintFnObject.ScalePercent = mOriginal_ScalePercent
            mLoaded_ScalePercent = False
        End If
        
        If mLoaded_LeftMargin Then
            LeftMargin = mOriginal_LeftMargin
            mLoaded_LeftMargin = False
        End If
        If mLoaded_TopMargin Then
            TopMargin = mOriginal_TopMargin
            mLoaded_TopMargin = False
        End If
        If mLoaded_RightMargin Then
            RightMargin = mOriginal_RightMargin
            mLoaded_RightMargin = False
        End If
        If mLoaded_BottomMargin Then
            BottomMargin = mOriginal_BottomMargin
            mLoaded_BottomMargin = False
        End If
        
        If mLoaded_ColorMode Then
            ColorMode = mOriginal_ColorMode
            mLoaded_ColorMode = False
        End If
        
        If mLoaded_PrintPageNumbers Then
            PrintPageNumbers = mOriginal_PrintPageNumbers
            mLoaded_PrintPageNumbers = False
        End If
        If mLoaded_PageNumbersPosition Then
            PageNumbersPosition = mOriginal_PageNumbersPosition
            mLoaded_PageNumbersPosition = False
        End If
        If mLoaded_PageNumbersFormat Then
            PageNumbersFormat = mOriginal_PageNumbersFormat
            mLoaded_PageNumbersFormat = False
        End If
        If mLoaded_PageNumbersFont Then
            Set PageNumbersFont = mOriginal_PageNumbersFont
            mLoaded_PageNumbersFont = False
        End If
        If mLoaded_PageNumbersForeColor Then
            PageNumbersForeColor = mOriginal_PageNumbersForeColor
            mLoaded_PageNumbersForeColor = False
        End If
        
        If mLoaded_DeviceName Then
            DeviceName = mOriginal_DeviceName
            mLoaded_DeviceName = False
        End If
        
        
        
        If (mDocKey <> "") Then
            iContextByDevice = Base64Encode("Prn" & mDeviceName & "_" & mDocKey)
            iContextGeneral = Base64Encode("Prn" & "_" & mDocKey)
            
            ' DeviceName
            iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextGeneral & "_DeviceName", "-")
            If (iStr <> "") And (iStr <> "-") Then
                iDn = DeviceName
                DeviceName = iStr
                If DeviceName <> iStr Then
                    DeviceName = iDn
                    Exit Sub
                End If
                If Not mLoaded_DeviceName Then
                    mOriginal_DeviceName = iDn
                    mLoaded_DeviceName = True
                End If
            End If
            
            ' Page numbers
            If mPrintPageNumbers Then
                If mPrintFnObject Is Nothing Then
                    Set mPrintFnObject = New PrintFnObject
                End If
                mPrintFnObject.PageNumbersButtonVisible = True
            End If
            iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextGeneral & "_PrintPageNumbers", CStr(CLng(mPrintPageNumbers)))
            If (iStr <> "") And (iStr <> "-") Then
                If IsNumeric(iStr) Then
                    If Not mLoaded_PrintPageNumbers Then
                        mOriginal_PrintPageNumbers = PrintPageNumbers
                        mLoaded_PrintPageNumbers = True
                    End If
                    PrintPageNumbers = CBool(Val(iStr))
                    If PrintPageNumbers Then
                        ' position
                        iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextGeneral & "_PageNumbersPosition", "-")
                        If (iStr <> "") And (iStr <> "-") Then
                            If IsNumeric(iStr) Then
                                If Not mLoaded_PageNumbersPosition Then
                                    mOriginal_PageNumbersPosition = PageNumbersPosition
                                    mLoaded_PageNumbersPosition = True
                                End If
                                PageNumbersPosition = Val(iStr)
                            End If
                        End If
                        ' format
                        iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextGeneral & "_PageNumbersFormat", "-")
                        If (iStr <> "") And (iStr <> "-") Then
                            If Not mLoaded_PageNumbersFormat Then
                                mOriginal_PageNumbersFormat = PageNumbersFormat
                                mLoaded_PageNumbersFormat = True
                            End If
                            If mPrintFnObject Is Nothing Then
                                Set mPrintFnObject = New PrintFnObject
                            End If
                            mPrintFnObject.EnsurePageNumberFormatInList mPageNumbersFormat
                            PageNumbersFormat = iStr
                        End If
                        ' Font
                        Set PageNumbersFont = GetSettingFont(AppNameForRegistry, "PrintingSettings", iContextGeneral & "_PageNumbersFont", mPageNumbersFont)
                        ' ForeColor
                        iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextGeneral & "_PageNumbersForeColor", "-")
                        If (iStr <> "") And (iStr <> "-") Then
                            If IsNumeric(iStr) Then
                                If Not mLoaded_PageNumbersForeColor Then
                                    mOriginal_PageNumbersForeColor = PageNumbersForeColor
                                    mLoaded_PageNumbersForeColor = True
                                End If
                                PageNumbersForeColor = Val(iStr)
                            End If
                        End If
                    End If
                End If
            End If
            
            ' ColorMode
            iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextGeneral & "_ColorMode", "-")
            If (iStr <> "") And (iStr <> "-") Then
                If IsNumeric(iStr) Then
                    If Not mLoaded_ColorMode Then
                        mOriginal_ColorMode = ColorMode
                        mLoaded_ColorMode = True
                    End If
                    ColorMode = Val(iStr)
                End If
            End If
            
            ' LeftMargin
            iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextGeneral & "_LeftMargin", "-")
            If (iStr <> "") And (iStr <> "-") Then
                If IsNumeric(iStr) Then
                    If Not mLoaded_LeftMargin Then
                        mOriginal_LeftMargin = LeftMargin
                        mLoaded_LeftMargin = True
                    End If
                    LeftMargin = Val(iStr)
                End If
            End If
            ' TopMargin
            iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextGeneral & "_TopMargin", "-")
            If (iStr <> "") And (iStr <> "-") Then
                If IsNumeric(iStr) Then
                    If Not mLoaded_TopMargin Then
                        mOriginal_TopMargin = TopMargin
                        mLoaded_TopMargin = True
                    End If
                    TopMargin = Val(iStr)
                End If
            End If
            ' RightMargin
            iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextGeneral & "_RightMargin", "-")
            If (iStr <> "") And (iStr <> "-") Then
                If IsNumeric(iStr) Then
                    If Not mLoaded_RightMargin Then
                        mOriginal_RightMargin = RightMargin
                        mLoaded_RightMargin = True
                    End If
                    RightMargin = Val(iStr)
                End If
            End If
            ' BottomMargin
            iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextGeneral & "_BottomMargin", "-")
            If (iStr <> "") And (iStr <> "-") Then
                If IsNumeric(iStr) Then
                    If Not mLoaded_BottomMargin Then
                        mOriginal_BottomMargin = BottomMargin
                        mLoaded_BottomMargin = True
                    End If
                    BottomMargin = Val(iStr)
                End If
            End If
            
            ' ScalePercent
            iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextByDevice & "_ScalePercent", "-")
            If (iStr <> "") And (iStr <> "-") Then
                If IsNumeric(iStr) Then
                    If Not mPrintFnObject Is Nothing Then
                        If Not mLoaded_ScalePercent Then
                            mOriginal_ScalePercent = mPrintFnObject.ScalePercent
                            mLoaded_ScalePercent = True
                        End If
                        mPrintFnObject.ScalePercent = Val(iStr)
                    End If
                End If
            End If
            
            ' Orientation
            iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextByDevice & "_Orientation", "-")
            If (iStr <> "") And (iStr <> "-") Then
                If IsNumeric(iStr) Then
                    If Not mLoaded_Orientation Then
                        mOriginal_Orientation = Orientation
                        mLoaded_Orientation = True
                    End If
                    Orientation = Val(iStr)
                End If
            End If
            ' PaperSize
            iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextByDevice & "_PaperSize", "-")
            If (iStr <> "") And (iStr <> "-") Then
                If IsNumeric(iStr) Then
                    If Not mLoaded_PaperSize Then
                        mOriginal_PaperSize = PaperSize
                        mLoaded_PaperSize = True
                    End If
                    PaperSize = Val(iStr)
                End If
            End If
            ' PaperBin
            iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextByDevice & "_PaperBin", "-")
            If (iStr <> "") And (iStr <> "-") Then
                If IsNumeric(iStr) Then
                    If Not mLoaded_PaperBin Then
                        mOriginal_PaperBin = PaperBin
                        mLoaded_PaperBin = True
                    End If
                    PaperBin = Val(iStr)
                End If
            End If
            ' PrintQuality
            iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextByDevice & "_PrintQuality", "-")
            If (iStr <> "") And (iStr <> "-") Then
                If IsNumeric(iStr) Then
                    If Not mLoaded_PrintQuality Then
                        mOriginal_PrintQuality = PrintQuality
                        mLoaded_PrintQuality = True
                    End If
                    PrintQuality = Val(iStr)
                End If
            End If
            ' Collate
            iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextByDevice & "_Collate", "-")
            If (iStr <> "") And (iStr <> "-") Then
                If Not mLoaded_Collate Then
                    mOriginal_Collate = Collate
                    mLoaded_Collate = True
                End If
                Collate = CBool(Val(iStr))
            End If
            ' Duplex
            iStr = GetSetting(AppNameForRegistry, "PrintingSettings", iContextByDevice & "_Duplex", "-")
            If (iStr <> "") And (iStr <> "-") Then
                If IsNumeric(iStr) Then
                    If Not mLoaded_Duplex Then
                        mOriginal_Duplex = Duplex
                        mLoaded_Duplex = True
                    End If
                    Duplex = Val(iStr)
                End If
            End If
        End If
        sLastDocKey = mDocKey
    End If
End Sub
