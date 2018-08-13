VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrePrintDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IvbExtra.IPrePrinterObj

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
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Declare Function BitBlt Lib "gdi32" _
  (ByVal hdcDest As Long, ByVal XDest As Long, _
   ByVal YDest As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hdcSrc As Long, _
   ByVal XSrc As Long, ByVal YSrc As Long, _
   ByVal dwRop As Long) As Long
   
'Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hdc As Long, lpInitData As Any) As Long

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

Private Declare Function StartDoc Lib "gdi32" Alias "StartDocA" (ByVal hdc As Long, lpdi As DOCINFO) As Long
Private Declare Function StartPage Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPage Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndDoc Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hemf As Long, lpRect As RECT) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nDrawStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private Const BS_SOLID As Long = 0
'Private Const BS_NULL As Long = 1
'Private Const BS_HATCHED As Long = 2

Private Declare Function CreateBrushIndirect Lib "gdi32.dll" (lpLogBrush As LOGBRUSH) As Long


Private mPages() As cPrePrintPage
Private mCurrentXPixels As Single
Private mCurrentYPixels As Single
Private mCurrentPageNumber As Long

Private mPrinterDefaultPageWidthInMillimeters As Single
Private mPrinterDefaultPageHeightInMillimeters As Single
Private mPrinterDefaultPageWidthInPixels As Long
Private mPrinterDefaultPageHeightInPixels As Long
Private mDPIX As Single
Private mDPIY As Single
Private mPaperWidthSettingInMillimeters As Single
Private mPaperHeightSettingInMillimeters As Single

Private mHdcReference As Long

Private mFromPage As Long
Private mToPage As Long
Private mPrintFromBackToFront As Boolean
Private mDocumentName As String
Private mLogFont As LOGFONT

'Private mAutoLineFeed As Boolean
Private mAutoHandleMargins As Boolean
Private mDrawMode As Long
Private mDrawStyle As Long
Private mDrawWidth As Long
Private mForeColor As Long
Private mFontTransparent As Boolean
Private mFillColor As Long
Private mFillStyle As Long
Private mDeviceName As String
Private mPort As String
Private mOrientation As Long
Private mPaperSize As Long
Private mPaperBin As Long
Private mPrintQuality As Long
Private mColorMode As Long
Private mDuplex As Long
Private mUserSelectedOrientation As Long
Private mUserSelectedPaperSize As Long
Private mUserSelectedPaperBin As Long
Private mUserSelectedPrintQuality As Long
Private mUserSelectedColorMode As Long
Private mUserSelectedDuplex As Long
Private mDevModePtr As Long
Private mScaleMode As Long
Private mScaleLeftPixels As Single
Private mScaleTopPixels As Single
Private mScaleWidthUser As Single
'Private mScaleHeightPixels As Single
'Private mScaleWidthPixels As Single
Private mScaleHeightUser As Single
Private mLeftMarginInMillimeters As Single
Private mRightMarginInMillimeters As Single
Private mTopMarginInMillimeters As Single
Private mBottomMarginInMillimeters As Single
Private mFonts() As String
Private mFontCount As Long

Private mFontName As String
Private mFontItalic As Boolean
Private mFontStrikethrough As Boolean
Private mFontUnderline As Boolean
Private mFontSize As Single
Private mFontWeight As Long
Private mFontWidth As Long
Private mFontCharset As Long
Private mFontBold As Boolean

Private mFontName_Prev As String
Private mFontItalic_Prev As Boolean
Private mFontStrikethrough_Prev As Boolean
Private mFontUnderline_Prev As Boolean
Private mFontSize_Prev As Long
Private mFontWeight_Prev As Single
Private mFontWidth_Prev As Long
Private mFontCharset_Prev As Long

Private mFontsUsed_LogFont() As LOGFONT
Private mhFontsUsed() As Long
Private mFontsUsed_HeightInPixels() As Long
Private mCurrentFontIndex As Long
Private mFontChanged As Boolean ' used for the Property Get Font

Private mDocumentCreated As Boolean
Private mPageCreated As Boolean
Private mSomethingPrintedOnPage As Boolean
Private mNewPagePending As Boolean
Private mDeleteReferenceDC As Boolean
Private mPaperSizeSet As Boolean
Private mDocumentTooLongWarningPrinted As Boolean
Private mPrintingDocumentTooLongWarning As Boolean
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

Private Const cMaxPages As Long = 1500

Public Sub CreateDocument(nHdcReference As Long, Optional nDocumentName As String)
    mHdcReference = nHdcReference
    mDocumentName = nDocumentName
    
    mPrinterDefaultPageWidthInPixels = GetDeviceCaps(mHdcReference, HORZRES)
    mPrinterDefaultPageHeightInPixels = GetDeviceCaps(mHdcReference, VERTRES)
    
    mPrinterDefaultPageWidthInMillimeters = mPrinterDefaultPageWidthInPixels / GetDeviceCaps(nHdcReference, LOGPIXELSX) * 25.4
    mPrinterDefaultPageHeightInMillimeters = mPrinterDefaultPageHeightInPixels / GetDeviceCaps(nHdcReference, LOGPIXELSY) * 25.4
    
    mDPIX = 25.4 * mPrinterDefaultPageWidthInPixels / mPrinterDefaultPageWidthInMillimeters
    mDPIY = 25.4 * mPrinterDefaultPageHeightInPixels / mPrinterDefaultPageHeightInMillimeters
    
    mPrinterDefaultPageWidthInMillimeters = 210
    mPrinterDefaultPageHeightInMillimeters = 297
    mPrinterDefaultPageWidthInPixels = mDPIX * mPrinterDefaultPageWidthInMillimeters / 25.4
    mPrinterDefaultPageHeightInPixels = mDPIY * mPrinterDefaultPageHeightInMillimeters / 25.4
    
    mDocumentCreated = True
    
    ReDim mPages(0)
    NewPage
End Sub


Public Sub PrintDocumentToPrinterDC(nHdc As Long)
    Dim iFromPage As Long
    Dim iToPage As Long
    Dim iAuxLng As Long
    Dim iStep As Long
    Dim iPN As Long
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
    If mFromPage = 0 Then
        iFromPage = 1
    Else
        iFromPage = mFromPage
    End If
    If iFromPage > iPC Then Exit Sub
    
    If mToPage = 0 Then
        iToPage = iPC
    Else
        iToPage = mToPage
    End If
    
    If mPrintFromBackToFront Then
        iAuxLng = iToPage
        iToPage = iFromPage
        iFromPage = iAuxLng
        iStep = -1
    Else
        iStep = 1
    End If
    
    If mDocumentName <> "" Then
        iDocumentName = ClientProductName & " - " & mDocumentName
    Else
        iDocumentName = ClientProductName
    End If
    
    iDI.cbSize = Len(iDI)
    iDI.lpszDocName = iDocumentName
    iDI.lpszOutput = vbNullString
    iDI.lpszDatatype = vbNullString
    Call StartDoc(nHdc, iDI)
    
    For iPN = iFromPage To iToPage Step iStep
        If iPN <= iPC Then
            If Not mPages(iPN) Is Nothing Then
                
                iRectPage.Left = 0
                iRectPage.Top = 0
                iRectPage.Right = GetDeviceCaps(nHdc, PHYSICALWIDTH) ' iRectPage.Left + mPages(iPN).WidthInPixels
                iRectPage.Bottom = GetDeviceCaps(nHdc, PHYSICALHEIGHT) ' iRectPage.Top + mPages(iPN).HeightInPixels
                
                If mDevModePtr <> 0 Then
                    iResetDC = False
                    If mPages(iPN).Orientation <> iOrientationPrev Then iResetDC = True
                    If mPages(iPN).PaperSize <> iPaperBinPrev Then iResetDC = True
                    If mPages(iPN).PaperBin <> iPaperSizePrev Then iResetDC = True
                    If mPages(iPN).PrintQuality <> iPrintQualityPrev Then iResetDC = True
                    If mPages(iPN).Duplex <> iDuplexPrev Then iResetDC = True
                    If mPages(iPN).ColorMode <> iColorModePrev Then iResetDC = True
                    
                    If iResetDC Then
                        CopyMemory iDevMode, ByVal mDevModePtr, Len(iDevMode)
                        iDevMode.dmOrientation = mPages(iPN).Orientation
                        iDevMode.dmPaperSize = mPages(iPN).PaperSize
                        iDevMode.dmDefaultSource = mPages(iPN).PaperBin
                        iDevMode.dmDuplex = mPages(iPN).Duplex
                        iDevMode.dmPrintQuality = mPages(iPN).PrintQuality
                        iDevMode.dmColor = mPages(iPN).ColorMode
                        CopyMemory ByVal mDevModePtr, iDevMode, Len(iDevMode)
                        ResetDC nHdc, iDevMode
                        iDCWasReset = True
                        iRectPage.Right = GetDeviceCaps(nHdc, PHYSICALWIDTH) ' iRectPage.Left + mPages(iPN).WidthInPixels
                        iRectPage.Bottom = GetDeviceCaps(nHdc, PHYSICALHEIGHT) ' iRectPage.Top + mPages(iPN).HeightInPixels
                    End If
                End If
                
                iOrientationPrev = mPages(iPN).Orientation
                iPaperBinPrev = mPages(iPN).PaperSize
                iPaperSizePrev = mPages(iPN).PaperBin
                iPrintQualityPrev = mPages(iPN).PrintQuality
                iDuplexPrev = mPages(iPN).Duplex
                iColorModePrev = mPages(iPN).ColorMode
                
                ihMetaFile = mPages(iPN).hMetaFile
                If ihMetaFile <> 0 Then
                    Call StartPage(nHdc)
                    PlayEnhMetaFile nHdc, ihMetaFile, iRectPage
                    Call EndPage(nHdc)
                End If
            End If
        End If
    Next iPN
    Call EndDoc(nHdc)
    
    If iDCWasReset Then
        ResetDC nHdc, iOriginalDevMode
    End If
End Sub


Public Sub PaintPageIntoPictureBox(nPageNumber As Long, nPic As PictureBox)
    Dim iRect As RECT
    Dim iPN As Long
    Dim ihMetaFile As Long
    Dim iAuxSM As Long
    
    If nPageNumber = 0 Then
        iPN = mCurrentPageNumber
    Else
        iPN = nPageNumber
    End If
    If (iPN < 1) Or (iPN > PageCount) Then Exit Sub
    
    If Not mPages(iPN) Is Nothing Then
        ihMetaFile = mPages(iPN).hMetaFile
        If ihMetaFile <> 0 Then
            iAuxSM = nPic.ScaleMode
            nPic.ScaleMode = vbPixels
            iRect.Right = nPic.ScaleWidth
            iRect.Bottom = nPic.ScaleHeight
            nPic.ScaleMode = iAuxSM
            nPic.Cls
            nPic.AutoRedraw = True
            On Error Resume Next
            PlayEnhMetaFile nPic.hdc, ihMetaFile, iRect
        End If
    End If
End Sub


Public Function NewPage() As Boolean
    If PageCount > cMaxPages Then
        mCurrentXPixels = 0
        mCurrentYPixels = 0
        mSomethingPrintedOnPage = False
        Exit Function
    End If
    Err.Clear
    On Error Resume Next
    mCurrentPageNumber = UBound(mPages) + 1
    If Err.Number = 9 Then
        On Error GoTo 0
        Err.Raise 3435, "PrePrinter1 not initialized, use PrePrinter1 from a PrintFn object"
    End If
    
    mNewPagePending = True
    
    mCurrentXPixels = 0
    mCurrentYPixels = 0
    mSomethingPrintedOnPage = False
    EnsureDocumentCreated2 'SetPixelScale
    NewPage = True
End Function

Private Sub AddNewPage()
    mCurrentPageNumber = UBound(mPages) + 1
    ReDim Preserve mPages(mCurrentPageNumber)
    Set mPages(mCurrentPageNumber) = New cPrePrintPage
    If mOrientation = vbPRORLandscape Then
        mPages(mCurrentPageNumber).CreatePage mHdcReference, PaperHeightSettingInMillimeters, PaperWidthSettingInMillimeters, mhFontsUsed(mCurrentFontIndex), mDrawStyle, mDrawWidth, mForeColor, mDrawMode, mFontTransparent, mFillStyle, mFillColor, mLeftMarginInMillimeters, mTopMarginInMillimeters, mRightMarginInMillimeters, mBottomMarginInMillimeters, mPaperSize, mPaperBin, mPrintQuality, mDuplex, mColorMode
    Else
        mPages(mCurrentPageNumber).CreatePage mHdcReference, PaperWidthSettingInMillimeters, PaperHeightSettingInMillimeters, mhFontsUsed(mCurrentFontIndex), mDrawStyle, mDrawWidth, mForeColor, mDrawMode, mFontTransparent, mFillStyle, mFillColor, mLeftMarginInMillimeters, mTopMarginInMillimeters, mRightMarginInMillimeters, mBottomMarginInMillimeters, mPaperSize, mPaperBin, mPrintQuality, mDuplex, mColorMode
    End If
    SetTextColor mPages(mCurrentPageNumber).MetaDC, mForeColor
    mNewPagePending = False
    mPageCreated = True
    
    If Not gPageFixedImage Is Nothing Then
        If Not gPageFixedImage.Handle = 0 Then
            Dim iDC As Long
            Dim iOO As Long
            Dim iXPos As Long
            Dim iYPos As Long
            Dim iW As Long
            Dim iH As Long
            
            iDC = CreateCompatibleDC(mPages(mCurrentPageNumber).MetaDC)
            iOO = SelectObject(iDC, gPageFixedImage.Handle)
            
            If iOO = 0 Then
                Dim iBytesPerLine As Long
                Dim iPicBits() As Byte
                Dim iTempBmp As Long
                Dim iPicInfo As BITMAP
                
                GetObjectAPI gPageFixedImage.Handle, Len(iPicInfo), iPicInfo
                iBytesPerLine = (iPicInfo.bmWidth * 3 + 3) And &HFFFFFFFC
                ReDim iPicBits(1 To iBytesPerLine * iPicInfo.bmHeight * 3) As Byte
                GetBitmapBits gPageFixedImage.Handle, UBound(iPicBits), iPicBits(1)
                iTempBmp = CreateCompatibleBitmap(iDC, iPicInfo.bmWidth, iPicInfo.bmHeight)
                SetBitmapBits iTempBmp, UBound(iPicBits), iPicBits(1)
                
                iOO = SelectObject(iDC, iTempBmp)
            End If
            
            iW = ScaleX(gPageFixedImage.Width, vbHimetric, vbTwips) / Screen.TwipsPerPixelX
            iH = ScaleY(gPageFixedImage.Height, vbHimetric, vbTwips) / Screen.TwipsPerPixelY
            
            If gPageFixedImage_PrintAtRight Then
                iXPos = mPages(mCurrentPageNumber).ScaleWidthInPixelsNoMargins - mPages(mCurrentPageNumber).RightMarginInPixels - iW
            Else
                iXPos = ScaleX(LeftMargin, vbMillimeters, vbPixels)
            End If
            iYPos = ScaleY(mPages(mCurrentPageNumber).HeightInMillimeters / 10 - 0.95, vbCentimeters, vbPixels) - iH
            
            BitBlt mPages(mCurrentPageNumber).MetaDC, iXPos, iYPos, iW, iH, iDC, 0, 0, vbSrcCopy
            SelectObject iDC, iOO
            If iTempBmp <> 0 Then
                DeleteObject iTempBmp
            End If
            DeleteDC iDC
        End If
    End If
    
    mCurrentXPixels = -mScaleLeftPixels + LeftMarginInPixels
    mCurrentYPixels = -mScaleTopPixels + TopMarginInPixels
End Sub

Public Sub KillDoc()
    mNewPagePending = False
    mCurrentPageNumber = 0
    ReDim mPages(0)
    mPageCreated = False
    mSomethingPrintedOnPage = False
    mCurrentXPixels = 0
    mCurrentYPixels = 0
    mDocumentCreated = False
End Sub

Private Sub CheckNewPagePending()
    If mNewPagePending Then AddNewPage
End Sub

Public Sub GoToPage(nPageNumber As Long)
    Dim iMetaDC As Long
    
    If (nPageNumber < 1) Or (nPageNumber > PageCount) Then
        Err.Raise 5787, "Bad page number"
        Exit Sub
    End If
    mCurrentPageNumber = nPageNumber
    
    iMetaDC = mPages(mCurrentPageNumber).MetaDC
    If iMetaDC <> 0 Then
        SetTextColor iMetaDC, mForeColor
        mPages(mCurrentPageNumber).SelectFont mhFontsUsed(mCurrentFontIndex)
        SetPen
        SetBrush
        SetBackMode
    End If
    mCurrentXPixels = -mScaleLeftPixels + LeftMarginInPixels
    mCurrentYPixels = -mScaleTopPixels + TopMarginInPixels

End Sub


Public Property Get Page() As Long
    Page = mCurrentPageNumber
End Property


Public Property Get CurrentX() As Single
    CurrentX = ScaleX(mCurrentXPixels + mScaleLeftPixels - LeftMarginInPixels, vbPixels, mScaleMode)
End Property

Public Property Let CurrentX(nValue As Single)
    mCurrentXPixels = ScaleX(nValue, mScaleMode, vbPixels) - mScaleLeftPixels + LeftMarginInPixels
End Property


Public Property Get CurrentY() As Single
    CurrentY = ScaleY(mCurrentYPixels + mScaleTopPixels - TopMarginInPixels, vbPixels, mScaleMode)
End Property

Public Property Let CurrentY(nValue As Single)
    mCurrentYPixels = ScaleY(nValue, mScaleMode, vbPixels) - mScaleTopPixels + TopMarginInPixels
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


Public Property Get PrintFromBackToFront() As Boolean
    PrintFromBackToFront = mPrintFromBackToFront
End Property

Public Property Let PrintFromBackToFront(nValue As Boolean)
    mPrintFromBackToFront = nValue
End Property


'Public Property Get AutoLineFeed() As Boolean
'    AutoLineFeed = mAutoLineFeed
'End Property
'
'Public Property Let AutoLineFeed(nValue As Boolean)
'    mAutoLineFeed = nValue
'End Property


Public Property Get AutoHandleMargins() As Boolean
    AutoHandleMargins = mAutoHandleMargins
End Property

Public Property Let AutoHandleMargins(nValue As Boolean)
    mAutoHandleMargins = nValue
    If mAutoHandleMargins Then
        If mCurrentXPixels < (LeftMarginInPixels - mScaleLeftPixels) Then
            mCurrentXPixels = LeftMarginInPixels - mScaleLeftPixels
        End If
    Else
        If mCurrentXPixels = (LeftMarginInPixels - mScaleLeftPixels) Then
            mCurrentXPixels = -mScaleLeftPixels
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
    
    TranslateColor nValue, 0, iColor
    If mCurrentPageNumber > 0 Then
        If mCurrentPageNumber <= UBound(mPages) Then
            iMetaDC = mPages(mCurrentPageNumber).MetaDC
            If iMetaDC <> 0 Then
                If (iColor <> mForeColor) Or sMDCWasZero Then
                    mForeColor = iColor
                    SetTextColor iMetaDC, mForeColor
                    SetPen
                End If
                sMDCWasZero = False
            Else
                sMDCWasZero = True
                mForeColor = iColor
            End If
        Else
            sMDCWasZero = True
            mForeColor = iColor
        End If
    Else
        sMDCWasZero = True
        mForeColor = iColor
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
    
    TranslateColor nValue, 0, iColor
    If iColor <> mFillColor Then
        mFillColor = iColor
        SetBrush
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
End Property


Public Property Get DeviceName() As String
    DeviceName = mDeviceName
End Property

Public Property Let DeviceName(nValue As String)
    If nValue <> mDeviceName Then
        If mSomethingPrintedOnPage Then
            'Err.Raise 396
        End If
        mDeviceName = nValue
        LoadPrinterFonts
    End If
End Property


Public Property Get Port() As String
    Port = mPort
End Property

Public Property Let Port(nValue As String)
    If nValue <> mPort Then
        If mSomethingPrintedOnPage Then
            'Err.Raise 396
        End If
        mPort = nValue
    End If
End Property


Public Property Get Orientation() As IvbExtra.cefPageOrientationConstants
    If mPageCreated Then
        Orientation = mPages(mCurrentPageNumber).Orientation
    Else
        Orientation = mOrientation
    End If
End Property

Public Property Let Orientation(nValue As IvbExtra.cefPageOrientationConstants)
    If nValue <> mOrientation Then
        If mSomethingPrintedOnPage Then
            'Err.Raise 396
        End If
        mOrientation = nValue
        If PageCount > 0 Then
            If Not mSomethingPrintedOnPage Then
                KillLastPage
                NewPage
            End If
        End If
    End If
End Property


Public Property Get PaperSize() As IvbExtra.cefPaperSizeConstants
    If mPageCreated Then
        PaperSize = mPages(mCurrentPageNumber).PaperSize
    Else
        PaperSize = mPaperSize
    End If
End Property

Public Property Let PaperSize(nValue As IvbExtra.cefPaperSizeConstants)
    If (nValue <> mPaperSize) Or Not mPaperSizeSet Then
        If mSomethingPrintedOnPage Then
            'Err.Raise 396
        End If
        mPaperSize = nValue
        mPaperSizeSet = True
        If mDeviceName = "" Then Err.Raise 4568, "DeviceName and Port must be set first"
        If mPort = "" Then Err.Raise 4568, "DeviceName and Port must be set first"
        If PageCount > 0 Then
            If Not mSomethingPrintedOnPage Then
                KillLastPage
                NewPage
            End If
        End If
        SetPageSize
    End If
End Property


Public Property Get PrintQuality() As IvbExtra.cefPrintQualityConstants
    If mPageCreated Then
        PrintQuality = mPages(mCurrentPageNumber).PrintQuality
    Else
        PrintQuality = mPrintQuality
    End If
End Property

Public Property Let PrintQuality(nValue As IvbExtra.cefPrintQualityConstants)
    If nValue <> mPrintQuality Then
        If mSomethingPrintedOnPage Then
            'Err.Raise 396
        End If
        mPrintQuality = nValue
        If PageCount > 0 Then
            If Not mSomethingPrintedOnPage Then
                KillLastPage
                NewPage
            End If
        End If
    End If
End Property


Public Property Get ColorMode() As IvbExtra.cefColorModeConstants
    If mPageCreated Then
        ColorMode = mPages(mCurrentPageNumber).ColorMode
    Else
        ColorMode = mColorMode
    End If
End Property

Public Property Let ColorMode(nValue As IvbExtra.cefColorModeConstants)
    If nValue <> mColorMode Then
        If mSomethingPrintedOnPage Then
            'Err.Raise 396
        End If
        mColorMode = nValue
    End If
End Property


Public Property Get PaperBin() As IvbExtra.cefPaperBinConstants
    If mPageCreated Then
        PaperBin = mPages(mCurrentPageNumber).PaperBin
    Else
        PaperBin = mPaperBin
    End If
End Property

Public Property Let PaperBin(nValue As IvbExtra.cefPaperBinConstants)
    If nValue <> mPaperBin Then
        If mSomethingPrintedOnPage Then
            'Err.Raise 396
        End If
        mPaperBin = nValue
        If PageCount > 0 Then
            If Not mSomethingPrintedOnPage Then
                KillLastPage
                NewPage
            End If
        End If
    End If
End Property


Public Property Get Duplex() As IvbExtra.cefDuplexConstants
    If mPageCreated Then
        Duplex = mPages(mCurrentPageNumber).Duplex
    Else
        Duplex = mDuplex
    End If
End Property

Public Property Let Duplex(nValue As IvbExtra.cefDuplexConstants)
    If nValue <> mDuplex Then
        If mSomethingPrintedOnPage Then
            'Err.Raise 396
        End If
        mDuplex = nValue
        If PageCount > 0 Then
            If Not mSomethingPrintedOnPage Then
                KillLastPage
                NewPage
            End If
        End If
    End If
End Property


Public Property Let UserSelectedOrientation(nValue As IvbExtra.cefPageOrientationConstants)
    mUserSelectedOrientation = nValue
    Orientation = mUserSelectedOrientation
End Property


Public Property Let UserSelectedPaperSize(nValue As IvbExtra.cefPaperSizeConstants)
    mUserSelectedPaperSize = nValue
    PaperSize = mUserSelectedPaperSize
End Property


Public Property Let UserSelectedPrintQuality(nValue As IvbExtra.cefPrintQualityConstants)
    mUserSelectedPrintQuality = nValue
    PrintQuality = mUserSelectedPrintQuality
End Property


Public Property Let UserSelectedColorMode(nValue As IvbExtra.cefColorModeConstants)
    mUserSelectedColorMode = nValue
    ColorMode = mUserSelectedColorMode
End Property


Public Property Let UserSelectedPaperBin(nValue As IvbExtra.cefPaperBinConstants)
    mUserSelectedPaperBin = nValue
    PaperBin = mUserSelectedPaperBin
End Property


Public Property Let UserSelectedDuplex(nValue As IvbExtra.cefDuplexConstants)
    mUserSelectedDuplex = nValue
    Duplex = mUserSelectedDuplex
End Property

Public Property Let DevModePtr(nValue As Long)
    mDevModePtr = nValue
End Property



Public Property Get ScaleMode() As IvbExtra.cefScaleModeConstants
    ScaleMode = mScaleMode
End Property

Public Property Let ScaleMode(nValue As IvbExtra.cefScaleModeConstants)
    If nValue <> mScaleMode Then
        If (nValue > 0) And (nValue < 10) Then
            If nValue = vbHimetric Then Err.Raise 380
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
    Scale_ nValue, ScaleTop, nValue + ScaleWidth, ScaleTop + ScaleHeight
End Property


Public Property Get ScaleTop() As Single
    ScaleTop = ScaleY(mScaleTopPixels, vbPixels, ScaleMode)
End Property

Public Property Let ScaleTop(nValue As Single)
    Scale_ ScaleLeft, nValue, ScaleLeft + ScaleWidth, nValue + ScaleHeight
End Property


Public Property Get ScaleWidth() As Single
    If ScaleMode = vbUser Then
        ScaleWidth = mScaleWidthUser
    Else
        If mScaleMode = vbTwips Then
            ScaleWidth = Round(ScaleX(CurrentPageScaleWidthInPixels, vbPixels, ScaleMode))
        Else
            ScaleWidth = ScaleX(CurrentPageScaleWidthInPixels, vbPixels, ScaleMode)
        End If
    End If
End Property

Public Property Let ScaleWidth(nValue As Single)
    Scale_ ScaleLeft, ScaleTop, ScaleLeft + nValue, ScaleTop + ScaleHeight
End Property


Public Property Get ScaleHeight() As Single
    If ScaleMode = vbUser Then
        ScaleHeight = mScaleHeightUser
    Else
        If mScaleMode = vbTwips Then
            ScaleHeight = Round(ScaleY(CurrentPageScaleHeightInPixels, vbPixels, ScaleMode))
        Else
            ScaleHeight = ScaleY(CurrentPageScaleHeightInPixels, vbPixels, ScaleMode)
        End If
    End If
End Property

Public Property Let ScaleHeight(nValue As Single)
    Scale_ ScaleLeft, ScaleTop, ScaleLeft + ScaleWidth, ScaleTop + nValue
End Property


Public Property Get LeftMargin() As Single ' millimeters
    LeftMargin = mLeftMarginInMillimeters
End Property

Public Property Let LeftMargin(nValue As Single)
    Dim iChange As Boolean
    
    iChange = nValue <> mLeftMarginInMillimeters
    If mDocumentCreated Then
        iChange = iChange Or (nValue <> mPages(mCurrentPageNumber).LeftMarginInMillimeters)
    End If
    
    If iChange Then
        mLeftMarginInMillimeters = nValue
        If mDocumentCreated Then
            mPages(mCurrentPageNumber).LeftMarginInMillimeters = mLeftMarginInMillimeters
            mCurrentXPixels = mCurrentXPixels - mPages(mCurrentPageNumber).LeftMarginInPixels
            mCurrentXPixels = -mScaleLeftPixels + LeftMarginInPixels
        Else
            mCurrentXPixels = -mScaleLeftPixels
        End If
    End If
End Property


Public Property Get RightMargin() As Single ' millimeters
    RightMargin = mRightMarginInMillimeters
End Property

Public Property Let RightMargin(nValue As Single)
    Dim iChange As Boolean
    
    iChange = nValue <> mRightMarginInMillimeters
    If mDocumentCreated Then
        iChange = iChange Or (nValue <> mPages(mCurrentPageNumber).RightMarginInMillimeters)
    End If
    
    If iChange Then
        mRightMarginInMillimeters = nValue
        If mDocumentCreated Then
            mPages(mCurrentPageNumber).RightMarginInMillimeters = mRightMarginInMillimeters
        End If
    End If
End Property


Public Property Get TopMargin() As Single ' millimeters
    TopMargin = mTopMarginInMillimeters
End Property

Public Property Let TopMargin(nValue As Single)
    Dim iChange As Boolean
    
    iChange = nValue <> mTopMarginInMillimeters
    If mDocumentCreated Then
        iChange = iChange Or (nValue <> mPages(mCurrentPageNumber).TopMarginInMillimeters)
    End If
    
    If iChange Then
        mTopMarginInMillimeters = nValue
        If mDocumentCreated Then
            mPages(mCurrentPageNumber).TopMarginInMillimeters = mTopMarginInMillimeters
            mCurrentYPixels = mCurrentYPixels - mPages(mCurrentPageNumber).TopMarginInPixels
            mCurrentYPixels = -mScaleTopPixels + TopMarginInPixels
        Else
            mCurrentYPixels = -mScaleTopPixels
        End If
    End If
End Property


Public Property Get BottomMargin() As Single ' millimeters
    BottomMargin = mBottomMarginInMillimeters
End Property

Public Property Let BottomMargin(nValue As Single)
    Dim iChange As Boolean
    
    iChange = nValue <> mBottomMarginInMillimeters
    If mDocumentCreated Then
        iChange = iChange Or (nValue <> mPages(mCurrentPageNumber).BottomMarginInMillimeters)
    End If
    
    If iChange Then
        mBottomMarginInMillimeters = nValue
        If mDocumentCreated Then
            mPages(mCurrentPageNumber).BottomMarginInMillimeters = mBottomMarginInMillimeters
        End If
    End If
End Property


Private Property Get LeftMarginInPixels() As Single
    If Not mDocumentCreated Then Exit Property
    If Not mAutoHandleMargins Then Exit Property
    If mNewPagePending Then Exit Property
    LeftMarginInPixels = mPages(mCurrentPageNumber).LeftMarginInPixels
End Property

Private Property Get TopMarginInPixels() As Single
    If Not mDocumentCreated Then Exit Property
    If Not mAutoHandleMargins Then Exit Property
    If mNewPagePending Then Exit Property
    TopMarginInPixels = mPages(mCurrentPageNumber).TopMarginInPixels
End Property

Private Property Get RightMarginInPixels() As Single
    If Not mDocumentCreated Then Exit Property
    If Not mAutoHandleMargins Then Exit Property
    If mNewPagePending Then Exit Property
    RightMarginInPixels = mPages(mCurrentPageNumber).RightMarginInPixels
End Property

Private Property Get BottomMarginInPixels() As Single
    If Not mDocumentCreated Then Exit Property
    If Not mAutoHandleMargins Then Exit Property
    If mNewPagePending Then Exit Property
    BottomMarginInPixels = mPages(mCurrentPageNumber).BottomMarginInPixels
End Property

'Private Sub SetPixelScale()
'    mScaleWidthPixels = ScaleX(Width - LeftMargin - RightMargin, vbMillimeters, vbPixels)
'    mScaleHeightPixels = ScaleX(Height - TopMargin - BottomMargin, vbMillimeters, vbPixels)
'End Sub

Public Sub Scale_(X1 As Variant, Y1 As Variant, x2 As Variant, Y2 As Variant)
    Dim iScaleLeftPixelsAnt As Single
    Dim iScaleTopPixelsAnt As Single
    
    mScaleWidthUser = x2 - X1
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
End Property


Public Property Get hdc() As Long
    If Not mDocumentCreated Then Exit Property
    
    CheckNewPagePending
    SelectFont
    hdc = mPages(mCurrentPageNumber).MetaDC
End Property


Private Sub SelectFont()
    Dim c As Long
    Dim iFound As Boolean
    Dim iFontChanged As Boolean
    
    If Not mDocumentCreated Then Exit Sub
    
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
            If StrConv(mLogFont.lfFaceName, vbUnicode) = StrConv(mFontsUsed_LogFont(c).lfFaceName, vbUnicode) Then
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
            mhFontsUsed(mCurrentFontIndex) = CreateFontIndirect(mLogFont)
        End If
        
        mPages(mCurrentPageNumber).SelectFont mhFontsUsed(mCurrentFontIndex)
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

Public Property Set Font(nFont As StdFont)
    Dim iFontName As String
    Dim iNoChange As Boolean
    
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
    mLogFont = StdFontToLogFont_Printer(Printer.Font)
    AssignFontVariables
    
    ReDim mhFontsUsed(0)
    ReDim mFontsUsed_LogFont(0)
    
    mScaleMode = vbTwips
'    mAutoLineFeed = True
    mAutoHandleMargins = True
    mDrawStyle = 0
    mDrawWidth = 1
    mForeColor = 0
    mFontTransparent = True
    mFillColor = 0
    mFillStyle = vbFSTransparent
    mDrawMode = vbCopyPen
    mOrientation = vbPRORPortrait
    mDuplex = vbPRDPSimplex
    mColorMode = vbPRCMColor
    mPrintQuality = vbPRPQPrinterDefault
    mPaperSize = vbPRPSA4
    mPaperBin = vbPRBNAuto
    SelectFont
End Sub

Private Sub Class_Terminate()
    Dim c As Long
    
    Erase mPages ' first erase the array of cPrPrintPages objects in order to make these objects to terminate and deselect the fonts
    
    For c = 1 To UBound(mhFontsUsed)
        DeleteObject mhFontsUsed(c)
    Next c
    If mDeleteReferenceDC Then
        DeleteDC mHdcReference
    End If
End Sub

Public Function PrintText(nText As String, Optional ByVal nLeft, Optional ByVal nTop, Optional ByVal nWidth, Optional ByVal nHeight, Optional ByVal nAlignment As cefAligmentConstants, Optional nAddLineFeed As Boolean = True, Optional nWordWrap As Boolean = True, Optional nWordBreak As Boolean = True, Optional nRectMarginsX As Single, Optional nRectMarginsY As Single, Optional nStartsNewPage As Boolean, Optional nCalcSize As Boolean, Optional ByRef nRequiredWidth As Single) As Single
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
    
    If Not mPrintingDocumentTooLongWarning Then
        If mCurrentPageNumber > cMaxPages Then
            If Not mDocumentTooLongWarningPrinted Then
                PrintDocumentTooLongWarning
            End If
            Exit Function
        End If
    End If
    
    If (nText = "") And Not nAddLineFeed Then Exit Function
    EnsureDocumentCreated
    iMetaDC = mPages(mCurrentPageNumber).MetaDC
    If iMetaDC = 0 Then Exit Function
    
    iAddLineFeed = nAddLineFeed 'And mAutoLineFeed
    
    SelectFont
        
    iEP.cbSize = Len(iEP)
    
'    If Not nCalcSize Then Stop
    If IsMissing(nLeft) Then
        iX = mCurrentXPixels
    Else
        iX = ScaleX(CSng(nLeft), mScaleMode, vbPixels) - mScaleLeftPixels + LeftMarginInPixels
    End If
    If IsMissing(nTop) Then
        iY = mCurrentYPixels
    Else
        iY = ScaleY(CSng(nTop), mScaleMode, vbPixels) - mScaleTopPixels + TopMarginInPixels
    End If
    If IsMissing(nWidth) Then
        If Not IsMissing(nLeft) Then
            iW = CurrentPageScaleWidthInPixels - iX + LeftMarginInPixels - mScaleLeftPixels
        Else
            iW = CurrentPageScaleWidthInPixels - mCurrentXPixels + LeftMarginInPixels - mScaleLeftPixels
        End If
    Else
        iW = ScaleX(CSng(nWidth), mScaleMode, vbPixels)
    End If
    If IsMissing(nHeight) Then
        If Not IsMissing(nTop) Then
            iH = CurrentPageScaleHeightInPixels - iY + TopMarginInPixels - mScaleTopPixels
        Else
            iH = CurrentPageScaleHeightInPixels - mCurrentYPixels + TopMarginInPixels - mScaleTopPixels
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
    iAuxText = nText
    If Len(iAuxText) = 0 Then iAuxText = iAuxText & " "
    Call DrawTextEx(iMetaDC, StrPtr(iAuxText), Len(iAuxText), iRectCalc, iWWFlag Or DT_EXPANDTABS Or DT_TABSTOP Or iWBFlag Or iDTAlign Or DT_NOPREFIX Or DT_CALCRECT, iEP)
    If Not nCalcSize And IsMissing(nWidth) Then
        If nWordWrap Then
            If iRectCalc.Right > CurrentPageScaleWidthInPixels + LeftMarginInPixels - mScaleLeftPixels Then
                If mCurrentXPixels <> (-mScaleLeftPixels + LeftMarginInPixels) Then
                    mCurrentXPixels = -mScaleLeftPixels + LeftMarginInPixels
                    mCurrentYPixels = mCurrentYPixels + FontHeightInternal
                    PrintText = PrintText(nText, , , nWidth, nHeight, nAlignment, nAddLineFeed, nWordWrap, nWordBreak, nRectMarginsX, nRectMarginsY, , , nRequiredWidth)
                    Exit Function
                End If
            End If
            If IsMissing(nLeft) Then
                If (iRectCalc.Bottom - iRectCalc.Top) > FontHeightInternal Then
                    If mCurrentXPixels > (-mScaleLeftPixels + LeftMarginInPixels) Then
                        mCurrentXPixels = -mScaleLeftPixels + LeftMarginInPixels
                        mCurrentYPixels = mCurrentYPixels + FontHeightInternal
                        PrintText = PrintText(nText, , , nWidth, nHeight, nAlignment, nAddLineFeed, nWordWrap, nWordBreak, nRectMarginsX, nRectMarginsY, , , nRequiredWidth)
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    If Not IsMissing(nHeight) Then
        iRect.Bottom = iRect.Top + iH - iYMargins * 2
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
    
    If iRect.Bottom > (CurrentPageScaleHeightInPixels - mScaleTopPixels + TopMarginInPixels) Then
        nStartsNewPage = True
    End If
    If nCalcSize Then
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
            iHeightSpaceAvailable = iH ' CurrentPageScaleHeightInPixels - mCurrentYPixels + TopMarginInPixels
            N = UBound(iWords) / (iRect.Bottom - iRect.Top) * iHeightSpaceAvailable
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
        mCurrentXPixels = -mScaleLeftPixels + LeftMarginInPixels
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
    mSomethingPrintedOnPage = True
    
    nRequiredWidth = ScaleX(nRequiredWidth, vbPixels, mScaleMode) + nRectMarginsX * 2
    PrintText = ScaleY(PrintText, vbPixels, mScaleMode) + nRectMarginsY * 2
End Function

Private Function FontHeightInternal() As Long
    Dim iTextSize As POINTAPI
    Dim iMetaDC As Long
    
    iMetaDC = mPages(mCurrentPageNumber).MetaDC
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
    TwipsPerPixelX = 1440 / mPages(mCurrentPageNumber).MetaDPIX
End Property

Public Property Get TwipsPerPixelY() As Single
    EnsureDocumentCreated
    TwipsPerPixelY = 1440 / mPages(mCurrentPageNumber).MetaDPIY
End Property

Public Function ScaleX(Width As Single, Optional ByVal FromScale As ScaleModeConstants = -1, Optional ByVal ToScale As ScaleModeConstants = -1) As Single
    If FromScale = -1 Then FromScale = vbHimetric
    If ToScale = -1 Then ToScale = vbTwips
    
    If (FromScale = vbPixels) Or (ToScale = vbPixels) Then
        EnsureDocumentCreated
    End If
    
    If ScaleMode <> vbUser Then
        If (FromScale = vbUser) Or (ToScale = vbUser) Then
            Err.Raise 5
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
                    ScaleX = Width * mPages(mCurrentPageNumber).MetaDPIX / 2.54
                Case vbPoints
                    ScaleX = Width * 72 / 2.54
                Case vbTwips
                    ScaleX = Width * 1440 / 2.54
                Case vbUser
                    ScaleX = Width / CurrentPageScaleWidthInPixels * mScaleWidthUser * mPages(mCurrentPageNumber).MetaDPIX / 2.54
                Case Else
                    Err.Raise 5
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
                    ScaleX = Width / 1440 * mPages(mCurrentPageNumber).MetaDPIX * 120
                Case vbPoints
                    ScaleX = Width / 20 * 120
                Case vbTwips
                    ScaleX = Width * 120
                Case vbUser
                    ScaleX = Width / CurrentPageScaleWidthInPixels * mScaleWidthUser / 1440 * mPages(mCurrentPageNumber).MetaDPIX * 120
                Case Else
                    Err.Raise 5
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
                    ScaleX = Width * mPages(mCurrentPageNumber).MetaDPIX / 2540
                Case vbPoints
                    ScaleX = Width / 2540 * 72
                Case vbTwips
                    ScaleX = Width / 2540 * 1440
                Case vbUser
                    ScaleX = Width / CurrentPageScaleWidthInPixels * mScaleWidthUser * mPages(mCurrentPageNumber).MetaDPIX / 2540
                Case Else
                    Err.Raise 5
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
                    ScaleX = Width * mPages(mCurrentPageNumber).MetaDPIX
                Case vbPoints
                    ScaleX = Width * 72
                Case vbTwips
                    ScaleX = Width * 1440
                Case vbUser
                    ScaleX = Width / CurrentPageScaleWidthInPixels * mScaleWidthUser * mPages(mCurrentPageNumber).MetaDPIX
                Case Else
                    Err.Raise 5
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
                    ScaleX = Width * mPages(mCurrentPageNumber).MetaDPIX / 25.4
                Case vbPoints
                    ScaleX = Width * 72 / 25.4
                Case vbTwips
                    ScaleX = Width * 1440 / 25.4
                Case vbUser
                    ScaleX = Width / CurrentPageScaleWidthInPixels * mScaleWidthUser * mPages(mCurrentPageNumber).MetaDPIX / 25.4
                Case Else
                    Err.Raise 5
            End Select
        Case vbPixels
            Select Case ToScale
                Case vbCentimeters
                    ScaleX = Width / mPages(mCurrentPageNumber).MetaDPIX * 2.54
                Case vbCharacters
                    ScaleX = Width * 1440 / mPages(mCurrentPageNumber).MetaDPIX / 120
                Case vbHimetric
                    ScaleX = Width / mPages(mCurrentPageNumber).MetaDPIX * 2540
                Case vbInches
                    ScaleX = Width / mPages(mCurrentPageNumber).MetaDPIX
                Case vbMillimeters
                    ScaleX = Width / mPages(mCurrentPageNumber).MetaDPIX * 25.4
                Case vbPixels
                    ScaleX = Width
                Case vbPoints
                    ScaleX = Width * 1440 / mPages(mCurrentPageNumber).MetaDPIX / 20
                Case vbTwips
                    ScaleX = Width * 1440 / mPages(mCurrentPageNumber).MetaDPIX
                Case vbUser
                    ScaleX = Width / CurrentPageScaleWidthInPixels * mScaleWidthUser
                Case Else
                    Err.Raise 5
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
                    ScaleX = Width / 1440 * mPages(mCurrentPageNumber).MetaDPIX * 20
                Case vbPoints
                    ScaleX = Width
                Case vbTwips
                    ScaleX = Width * 20
                Case vbUser
                    ScaleX = Width / CurrentPageScaleWidthInPixels * mScaleWidthUser * mPages(mCurrentPageNumber).MetaDPIX / 72
                Case Else
                    Err.Raise 5
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
                    ScaleX = Width / 1440 * mPages(mCurrentPageNumber).MetaDPIX
                Case vbPoints
                    ScaleX = Width / 20
                Case vbTwips
                    ScaleX = Width
                Case vbUser
                    ScaleX = Width / CurrentPageScaleWidthInPixels * mScaleWidthUser / 1440 * mPages(mCurrentPageNumber).MetaDPIX
                Case Else
                    Err.Raise 5
            End Select
        Case vbUser
            Select Case ToScale
                Case vbCentimeters
                    ScaleX = Width * CurrentPageScaleWidthInPixels / mScaleWidthUser / mPages(mCurrentPageNumber).MetaDPIX * 2.54
                Case vbCharacters
                    ScaleX = Width * CurrentPageScaleWidthInPixels / mScaleWidthUser * 1440 / mPages(mCurrentPageNumber).MetaDPIX / 120
                Case vbHimetric
                    ScaleX = Width * CurrentPageScaleWidthInPixels / mScaleWidthUser / mPages(mCurrentPageNumber).MetaDPIX * 2540
                Case vbInches
                    ScaleX = Width * CurrentPageScaleWidthInPixels / mScaleWidthUser / mPages(mCurrentPageNumber).MetaDPIX
                Case vbMillimeters
                    ScaleX = Width * CurrentPageScaleWidthInPixels / mScaleWidthUser / mPages(mCurrentPageNumber).MetaDPIX * 25.4
                Case vbPixels
                    ScaleX = Width * CurrentPageScaleWidthInPixels / mScaleWidthUser
                Case vbPoints
                    ScaleX = Width * CurrentPageScaleWidthInPixels / mScaleWidthUser / mPages(mCurrentPageNumber).MetaDPIX * 72
                Case vbTwips
                    ScaleX = Width * CurrentPageScaleWidthInPixels / mScaleWidthUser * 1440 / mPages(mCurrentPageNumber).MetaDPIX
                Case vbUser
                    ScaleX = Width
                Case Else
                    Err.Raise 5
            End Select
        Case Else
            Err.Raise 5
    End Select
End Function

Public Function ScaleY(Height As Single, Optional ByVal FromScale As ScaleModeConstants = -1, Optional ByVal ToScale As ScaleModeConstants = -1) As Single
    If FromScale = -1 Then FromScale = vbHimetric
    If ToScale = -1 Then ToScale = vbTwips
    
    If (FromScale = vbPixels) Or (ToScale = vbPixels) Then
        EnsureDocumentCreated
    End If
    
    If ScaleMode <> vbUser Then
        If (FromScale = vbUser) Or (ToScale = vbUser) Then
            Err.Raise 5
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
                    ScaleY = Height * mPages(mCurrentPageNumber).MetaDPIY / 2.54
                Case vbPoints
                    ScaleY = Height * 72 / 2.54
                Case vbTwips
                    ScaleY = Height * 1440 / 2.54
                Case vbUser
                    ScaleY = Height / CurrentPageScaleHeightInPixels * mScaleHeightUser * mPages(mCurrentPageNumber).MetaDPIY / 2.54
                Case Else
                    Err.Raise 5
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
                    ScaleY = Height / 1440 * mPages(mCurrentPageNumber).MetaDPIY * 240
                Case vbPoints
                    ScaleY = Height / 20 * 240
                Case vbTwips
                    ScaleY = Height * 240
                Case vbUser
                    ScaleY = Height / CurrentPageScaleHeightInPixels * mScaleHeightUser / 1440 * mPages(mCurrentPageNumber).MetaDPIY * 240
                Case Else
                    Err.Raise 5
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
                    ScaleY = Height * mPages(mCurrentPageNumber).MetaDPIY / 2540
                Case vbPoints
                    ScaleY = Height / 2540 * 72
                Case vbTwips
                    ScaleY = Height / 2540 * 1440
                Case vbUser
                    ScaleY = Height / CurrentPageScaleHeightInPixels * mScaleHeightUser * mPages(mCurrentPageNumber).MetaDPIY / 2540
                Case Else
                    Err.Raise 5
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
                    ScaleY = Height * mPages(mCurrentPageNumber).MetaDPIY
                Case vbPoints
                    ScaleY = Height * 72
                Case vbTwips
                    ScaleY = Height * 1440
                Case vbUser
                    ScaleY = Height / CurrentPageScaleHeightInPixels * mScaleHeightUser * mPages(mCurrentPageNumber).MetaDPIY
                Case Else
                    Err.Raise 5
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
                    ScaleY = Height * mPages(mCurrentPageNumber).MetaDPIY / 25.4
                Case vbPoints
                    ScaleY = Height * 72 / 25.4
                Case vbTwips
                    ScaleY = Height * 1440 / 25.4
                Case vbUser
                    ScaleY = Height / CurrentPageScaleHeightInPixels * mScaleHeightUser * mPages(mCurrentPageNumber).MetaDPIY / 25.4
                Case Else
                    Err.Raise 5
            End Select
        Case vbPixels
            Select Case ToScale
                Case vbCentimeters
                    ScaleY = Height / mPages(mCurrentPageNumber).MetaDPIY * 2.54
                Case vbCharacters
                    ScaleY = Height * 1440 / mPages(mCurrentPageNumber).MetaDPIY / 240
                Case vbHimetric
                    ScaleY = Height / mPages(mCurrentPageNumber).MetaDPIY * 2540
                Case vbInches
                    ScaleY = Height / mPages(mCurrentPageNumber).MetaDPIY
                Case vbMillimeters
                    ScaleY = Height / mPages(mCurrentPageNumber).MetaDPIY * 25.4
                Case vbPixels
                    ScaleY = Height
                Case vbPoints
                    ScaleY = Height * 1440 / mPages(mCurrentPageNumber).MetaDPIY / 20
                Case vbTwips
                    ScaleY = Height * 1440 / mPages(mCurrentPageNumber).MetaDPIY
                Case vbUser
                    ScaleY = Height / CurrentPageScaleHeightInPixels * mScaleHeightUser
                Case Else
                    Err.Raise 5
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
                    ScaleY = Height / 1440 * mPages(mCurrentPageNumber).MetaDPIY * 20
                Case vbPoints
                    ScaleY = Height
                Case vbTwips
                    ScaleY = Height * 20
                Case vbUser
                    ScaleY = Height / CurrentPageScaleHeightInPixels * mScaleHeightUser * mPages(mCurrentPageNumber).MetaDPIY / 72
                Case Else
                    Err.Raise 5
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
                    ScaleY = Height / 1440 * mPages(mCurrentPageNumber).MetaDPIY
                Case vbPoints
                    ScaleY = Height / 20
                Case vbTwips
                    ScaleY = Height
                Case vbUser
                    ScaleY = Height / CurrentPageScaleHeightInPixels * mScaleHeightUser / 1440 * mPages(mCurrentPageNumber).MetaDPIY
                Case Else
                    Err.Raise 5
            End Select
        Case vbUser
            Select Case ToScale
                Case vbCentimeters
                    ScaleY = Height * CurrentPageScaleHeightInPixels / mScaleHeightUser / mPages(mCurrentPageNumber).MetaDPIY * 2.54
                Case vbCharacters
                    ScaleY = Height * CurrentPageScaleHeightInPixels / mScaleHeightUser * 1440 / mPages(mCurrentPageNumber).MetaDPIY / 240
                Case vbHimetric
                    ScaleY = Height * CurrentPageScaleHeightInPixels / mScaleHeightUser / mPages(mCurrentPageNumber).MetaDPIY * 2540
                Case vbInches
                    ScaleY = Height * CurrentPageScaleHeightInPixels / mScaleHeightUser / mPages(mCurrentPageNumber).MetaDPIY
                Case vbMillimeters
                    ScaleY = Height * CurrentPageScaleHeightInPixels / mScaleHeightUser / mPages(mCurrentPageNumber).MetaDPIY * 25.4
                Case vbPixels
                    ScaleY = Height * CurrentPageScaleHeightInPixels / mScaleHeightUser
                Case vbPoints
                    ScaleY = Height * CurrentPageScaleHeightInPixels / mScaleHeightUser / mPages(mCurrentPageNumber).MetaDPIY * 72
                Case vbTwips
                    ScaleY = Height * CurrentPageScaleHeightInPixels / mScaleHeightUser * 1440 / mPages(mCurrentPageNumber).MetaDPIY
                Case vbUser
                    ScaleY = Height
                Case Else
                    Err.Raise 5
            End Select
        Case Else
            Err.Raise 5
    End Select
End Function

Private Sub EnsureDocumentCreated()
    If Not mDocumentCreated Then CreateDocument HdcReference, mDocumentName
    CheckNewPagePending
End Sub

Private Sub EnsureDocumentCreated2()
    If Not mDocumentCreated Then CreateDocument HdcReference, mDocumentName
End Sub

'Public Property Let HdcReference(nValue As Long)
'    mHdcReference = nValue
'End Property

Public Property Get HdcReference() As Long
    Dim iDlgAux As New CommonDialogExObject
    
    If mHdcReference <> 0 Then
        HdcReference = mHdcReference
    Else
        If Printers.Count = 0 Then
            HdcReference = CreateCompatibleDC(0)
        Else
            iDlgAux.ShowPrinter cdePDReturnIC Or cdePDReturnDefault
            mHdcReference = CreateCompatibleDC(iDlgAux.hdc) ' this should be updated with the ResetDC API if the PrintQuality, PaperSize or DeviceName change (*), or create a new reference DC if it can't be updated (I don't know if the IC can be updated after being used as a reference -to create the metafiles-), also doing so it's necessary to change the way CreatePage in the cPrePrintPage class module sets the paper dimensions. It's not implemented at this time. (*) DeviceName can't be updated with ResetDC, in that case it's necessary to create a new DC with the defaults of that printer with the common dialog API setting the devide name in the DEVNAMES (that's not implemented in the CommonDialogExObject class module at this time -to be able to set the DeviceName property before calling ShowPrinter with cdePDReturnIC Or cdePDReturnDefault-).
            HdcReference = mHdcReference
        End If
        mDeleteReferenceDC = True
    End If
End Property

Private Sub SetPen()
    If Not mDocumentCreated Then Exit Sub
    mPages(mCurrentPageNumber).SetPen mDrawStyle, mDrawWidth, mForeColor, mDrawMode
End Sub

Private Sub SetBrush()
    If Not mDocumentCreated Then Exit Sub
    mPages(mCurrentPageNumber).SetBrush mFillStyle, mFillColor
End Sub

Private Sub SetBackMode()
    If Not mDocumentCreated Then Exit Sub
    mPages(mCurrentPageNumber).FontTransparent = mFontTransparent
End Sub

Public Sub DrawLine(X1 As Single, Y1 As Single, x2 As Single, Y2 As Single, Optional Color, Optional b As Boolean, Optional F As Boolean)
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
    iMetaDC = mPages(mCurrentPageNumber).MetaDC
    If iMetaDC = 0 Then Exit Sub
    
    If mCurrentPageNumber > cMaxPages Then
        If Not mDocumentTooLongWarningPrinted Then
            PrintDocumentTooLongWarning
        End If
        Exit Sub
    End If
    
    iX1 = ScaleX(X1, mScaleMode, vbPixels) - mScaleLeftPixels + LeftMarginInPixels
    iY1 = ScaleY(Y1, mScaleMode, vbPixels) - mScaleTopPixels + TopMarginInPixels
    iX2 = ScaleX(x2, mScaleMode, vbPixels) - mScaleLeftPixels + LeftMarginInPixels
    iY2 = ScaleY(Y2, mScaleMode, vbPixels) - mScaleTopPixels + TopMarginInPixels
    
    If Not IsMissing(Color) Then
        TranslateColor CLng(Color), 0, iColor
        If iColor <> mForeColor Then
            iPen = CreatePen(mDrawStyle, mDrawWidth, iColor)
            iPenPrev = SelectObject(iMetaDC, iPen)
        End If
    Else
        iColor = mForeColor
    End If
    
    If Not b Then
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
        If F Then
            If (iColor <> mFillColor) Or (mFillStyle <> vbFSSolid) Then
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
    mCurrentXPixels = iX2 '- mScaleLeftPixels + LeftMarginInPixels
    mCurrentYPixels = iY2 '- mScaleTopPixels + TopMarginInPixels
    mSomethingPrintedOnPage = True
End Sub


Public Property Get Width() As Single ' millimeters
    Dim iPw As Single
    
    If mDocumentCreated Then
        EnsureDocumentCreated
        Width = Round(ScaleY(mPages(mCurrentPageNumber).PageWidthInPixels, vbPixels, vbTwips))
    Else
        If mOrientation = vbPRORPortrait Then
            iPw = mPaperWidthSettingInMillimeters
        Else
            iPw = mPaperHeightSettingInMillimeters
        End If
        If iPw <> 0 Then
            Width = ScaleX(iPw, vbMillimeters, vbTwips)
        Else
            Width = ScaleX(GetDeviceCaps(HdcReference, HORZRES) / GetDeviceCaps(HdcReference, LOGPIXELSX), vbInches, vbTwips)
        End If
    End If
End Property


Public Property Get Height() As Single ' millimeters
    Dim iPh As Single
    
    If mDocumentCreated Then
        EnsureDocumentCreated
        Height = Round(ScaleX(mPages(mCurrentPageNumber).PageHeightInPixels, vbPixels, vbTwips))
    Else
        If mOrientation = vbPRORPortrait Then
            iPh = mPaperHeightSettingInMillimeters
        Else
            iPh = mPaperWidthSettingInMillimeters
        End If
        If iPh <> 0 Then
            Height = ScaleY(iPh, vbMillimeters, vbTwips)
        Else
            Height = ScaleY(GetDeviceCaps(HdcReference, VERTRES) / GetDeviceCaps(HdcReference, LOGPIXELSY), vbInches, vbTwips)
        End If
    End If
End Property





' Interface

Private Property Let IPrePrinterObj_AutoHandleMargins(ByVal RHS As Long)
    AutoHandleMargins = RHS
End Property

Private Property Get IPrePrinterObj_AutoHandleMargins() As Long
    IPrePrinterObj_AutoHandleMargins = AutoHandleMargins
End Property


Private Property Get IPrePrinterObj_BottomMargin() As Long
    IPrePrinterObj_BottomMargin = ScaleX(BottomMarginInPixels, vbPixels, ScaleMode)
End Property

Private Sub IPrePrinterObj_Circle(Step As Integer, x As Single, y As Single, Radius As Single, Color As Long, Start As Single, EndArc As Single, Aspect As Single)
    ' need implementation
End Sub


Private Property Let IPrePrinterObj_ColorMode(ByVal RHS As IvbExtra.cefColorModeConstants)
    ColorMode = RHS
End Property

Private Property Get IPrePrinterObj_ColorMode() As IvbExtra.cefColorModeConstants
    IPrePrinterObj_ColorMode = ColorMode
End Property


Private Property Let IPrePrinterObj_CurrentX(ByVal RHS As Single)
    CurrentX = RHS
    If mScaleMode = vbTwips Then
        CurrentX = Round(RHS)
    Else
        CurrentX = RHS
    End If
End Property

Private Property Get IPrePrinterObj_CurrentX() As Single
    IPrePrinterObj_CurrentX = CurrentX
    If mScaleMode = vbTwips Then
        IPrePrinterObj_CurrentX = Round(IPrePrinterObj_CurrentX)
    End If
End Property


Private Property Let IPrePrinterObj_CurrentY(ByVal RHS As Single)
    If mScaleMode = vbTwips Then
        CurrentY = Round(RHS)
    Else
        CurrentY = RHS
    End If
End Property

Private Property Get IPrePrinterObj_CurrentY() As Single
    IPrePrinterObj_CurrentY = CurrentY
    If mScaleMode = vbTwips Then
        IPrePrinterObj_CurrentY = Round(IPrePrinterObj_CurrentY)
    End If
End Property


Private Property Let IPrePrinterObj_DrawMode(ByVal RHS As IvbExtra.cefDrawModeConstants)
    DrawMode = RHS
End Property

Private Property Get IPrePrinterObj_DrawMode() As IvbExtra.cefDrawModeConstants
    IPrePrinterObj_DrawMode = DrawMode
End Property


Private Property Let IPrePrinterObj_DrawStyle(ByVal RHS As IvbExtra.cefDrawStyleConstants)
    DrawStyle = RHS
End Property

Private Property Get IPrePrinterObj_DrawStyle() As IvbExtra.cefDrawStyleConstants
    IPrePrinterObj_DrawStyle = DrawStyle
End Property


Private Property Let IPrePrinterObj_DrawWidth(ByVal RHS As Integer)
    DrawWidth = RHS
End Property

Private Property Get IPrePrinterObj_DrawWidth() As Integer
    IPrePrinterObj_DrawWidth = DrawWidth
End Property


Private Property Get IPrePrinterObj_Duplex() As IvbExtra.cefDuplexConstants
    IPrePrinterObj_Duplex = Duplex
End Property

Private Property Let IPrePrinterObj_Duplex(ByVal RHS As IvbExtra.cefDuplexConstants)
    Duplex = RHS
End Property


Private Sub IPrePrinterObj_EndDoc()
    ' nothing, this is ignored, added just for compatibility
End Sub

Private Property Let IPrePrinterObj_FillColor(ByVal RHS As Long)
    FillColor = RHS
End Property

Private Property Get IPrePrinterObj_FillColor() As Long
    IPrePrinterObj_FillColor = FillColor
End Property


Private Property Let IPrePrinterObj_FillStyle(ByVal RHS As IvbExtra.cefFillStyleConstants)
    FillStyle = RHS
End Property

Private Property Get IPrePrinterObj_FillStyle() As IvbExtra.cefFillStyleConstants
    IPrePrinterObj_FillStyle = FillStyle
End Property


Private Property Set IPrePrinterObj_Font(ByVal RHS As stdole.StdFont)
    Set Font = RHS
End Property

Private Property Get IPrePrinterObj_Font() As stdole.StdFont
    Set IPrePrinterObj_Font = Font
End Property


Private Property Let IPrePrinterObj_FontBold(ByVal RHS As Boolean)
    FontBold = RHS
End Property

Private Property Get IPrePrinterObj_FontBold() As Boolean
    IPrePrinterObj_FontBold = FontBold
End Property


Private Property Get IPrePrinterObj_FontCount() As Integer
    IPrePrinterObj_FontCount = FontCount
End Property


Private Property Let IPrePrinterObj_FontItalic(ByVal RHS As Boolean)
    FontItalic = RHS
End Property

Private Property Get IPrePrinterObj_FontItalic() As Boolean
    IPrePrinterObj_FontItalic = FontItalic
End Property


Private Property Let IPrePrinterObj_FontName(ByVal RHS As String)
    FontName = RHS
End Property

Private Property Get IPrePrinterObj_FontName() As String
    IPrePrinterObj_FontName = FontName
End Property


Private Property Get IPrePrinterObj_Fonts(ByVal Index As Integer) As String
    IPrePrinterObj_Fonts = Fonts(Index)
End Property


Private Property Let IPrePrinterObj_FontSize(ByVal RHS As Single)
    FontSize = RHS
End Property

Private Property Get IPrePrinterObj_FontSize() As Single
    IPrePrinterObj_FontSize = FontSize
End Property


Private Property Let IPrePrinterObj_FontStrikethru(ByVal RHS As Boolean)
    FontStrikeThru = RHS
End Property

Private Property Get IPrePrinterObj_FontStrikethru() As Boolean
    IPrePrinterObj_FontStrikethru = FontStrikeThru
End Property


Private Property Let IPrePrinterObj_FontTransparent(ByVal RHS As Boolean)
    FontTransparent = RHS
End Property

Private Property Get IPrePrinterObj_FontTransparent() As Boolean
    IPrePrinterObj_FontTransparent = FontTransparent
End Property


Private Property Let IPrePrinterObj_FontUnderline(ByVal RHS As Boolean)
    FontUnderLine = RHS
End Property

Private Property Get IPrePrinterObj_FontUnderline() As Boolean
    IPrePrinterObj_FontUnderline = FontUnderLine
End Property


Private Property Let IPrePrinterObj_FontWidth(ByVal RHS As Integer)
    FontWidth = RHS
End Property

Private Property Get IPrePrinterObj_FontWidth() As Integer
    IPrePrinterObj_FontWidth = FontWidth
End Property


Private Property Let IPrePrinterObj_ForeColor(ByVal RHS As Long)
    ForeColor = RHS
End Property

Private Property Get IPrePrinterObj_ForeColor() As Long
    IPrePrinterObj_ForeColor = ForeColor
End Property


Private Sub IPrePrinterObj_GoToPage(ByVal PageNumber As Long)
    GoToPage PageNumber
End Sub

Private Property Get IPrePrinterObj_hDC() As Long
    IPrePrinterObj_hDC = hdc
End Property


Private Property Get IPrePrinterObj_Height() As Single
    IPrePrinterObj_Height = Height
End Property


Private Sub IPrePrinterObj_KillDoc()
    KillDoc
End Sub


Private Property Get IPrePrinterObj_LeftMargin() As Long
    IPrePrinterObj_LeftMargin = ScaleX(LeftMarginInPixels, vbPixels, ScaleMode)
End Property

Private Sub IPrePrinterObj_Line(ByVal flags As Integer, ByVal X1 As Single, ByVal Y1 As Single, ByVal x2 As Single, ByVal Y2 As Single, ByVal Color As Long)
    Dim b As Boolean
    Dim F As Boolean
    Dim c As Boolean
    Dim iStep1 As Boolean
    Dim iStep2 As Boolean
    
    iStep1 = (flags And 1) <> 0
    iStep2 = (flags And 8) <> 0
    c = (flags And 2) <> 0
    
    b = (flags And 16) <> 0
    F = (flags And 32) <> 0
    If F Then b = True
    
    If iStep1 Then
        X1 = X1 + CurrentX
        Y1 = Y1 + CurrentY
    End If
    If iStep2 Then
        x2 = x2 + X1
        Y2 = Y2 + Y1
    End If
    
    If c Then
        DrawLine X1, Y1, x2, Y2, Color, b, F
    Else
        DrawLine X1, Y1, x2, Y2, , b, F
    End If
    
End Sub


Private Sub IPrePrinterObj_NewPage()
    NewPage
End Sub


Private Property Let IPrePrinterObj_Orientation(ByVal RHS As IvbExtra.cefPageOrientationConstants)
    Orientation = RHS
End Property

Private Property Get IPrePrinterObj_Orientation() As IvbExtra.cefPageOrientationConstants
    IPrePrinterObj_Orientation = Orientation
End Property


Private Property Get IPrePrinterObj_Page() As Long
    IPrePrinterObj_Page = Page
End Property


Private Property Get IPrePrinterObj_PageCount() As Long
    IPrePrinterObj_PageCount = PageCount
End Property


Private Sub IPrePrinterObj_PaintPicture(ByVal Picture As stdole.Picture, ByVal X1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal x2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
    ' need implementation
End Sub


Private Property Let IPrePrinterObj_PaperBin(ByVal RHS As IvbExtra.cefPaperBinConstants)
    PaperBin = RHS
End Property
    
Private Property Get IPrePrinterObj_PaperBin() As IvbExtra.cefPaperBinConstants
    IPrePrinterObj_PaperBin = PaperBin
End Property


Private Property Let IPrePrinterObj_PaperSize(ByVal RHS As IvbExtra.cefPaperSizeConstants)
    If mDeviceName = "" Then
        mDeviceName = Printer.DeviceName
        mPort = Printer.Port
    End If
    PaperSize = RHS
End Property

Private Property Get IPrePrinterObj_PaperSize() As IvbExtra.cefPaperSizeConstants
    IPrePrinterObj_PaperSize = PaperSize
End Property


Private Sub IPrePrinterObj_Preprint(Optional Text, Optional Options As IvbExtra.cefPrePrintOptionsConstants)
    Dim iAddLineFeed As Boolean
    Dim iAddTab As Boolean
    Dim iText As String
    
    If (Options And ppNoLineFeed) = 0 Then
        iAddLineFeed = True
    End If
    If (Options And ppAddTab) <> 0 Then
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


Private Sub IPrePrinterObj_PrePrintEx(Optional Text, Optional ByVal Options As IvbExtra.cefPrePrintOptionsConstants, Optional ByVal Left As Variant, Optional ByVal Top As Variant, Optional ByVal Width As Variant, Optional ByVal Height As Variant, Optional ByVal Alignment As IvbExtra.cefAligmentConstants, Optional ByVal WordWrap As Boolean = True, Optional ByVal WordBreak As Boolean = True, Optional ByVal RectMarginsX As Single, Optional ByVal RectMarginsY As Single, Optional StartsNewPage As Boolean)
    Dim iAddLineFeed As Boolean
    Dim iAddTab As Boolean
    Dim iCellMarginsX As Single
    Dim iCellMarginsY As Single
    Dim iText As String
    
    If (Options And ppNoLineFeed) = 0 Then
        iAddLineFeed = True
    End If
    If (Options And ppAddTab) <> 0 Then
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
        PrintText iText & vbTab, Left, Top, Width, Height, Alignment, iAddLineFeed, WordWrap, WordBreak, iCellMarginsX, iCellMarginsY, StartsNewPage
    Else
        PrintText iText, Left, Top, Width, Height, Alignment, iAddLineFeed, WordWrap, WordBreak, iCellMarginsX, iCellMarginsY, StartsNewPage
    End If
End Sub


Private Property Let IPrePrinterObj_PrintQuality(ByVal RHS As IvbExtra.cefPrintQualityConstants)
    PrintQuality = RHS
End Property

Private Property Get IPrePrinterObj_PrintQuality() As IvbExtra.cefPrintQualityConstants
    IPrePrinterObj_PrintQuality = PrintQuality
End Property


Private Sub IPrePrinterObj_PSet(ByVal Step As Integer, ByVal x As Single, ByVal y As Single, ByVal Color As Long)
    ' need implementation
End Sub


Private Property Get IPrePrinterObj_RightMargin() As Long
    IPrePrinterObj_RightMargin = ScaleX(RightMarginInPixels, vbPixels, ScaleMode)
End Property

Private Sub IPrePrinterObj_Scale(ByVal flags As Integer, ByVal X1 As Variant, ByVal Y1 As Variant, ByVal x2 As Variant, ByVal Y2 As Variant)
    Scale_ X1, Y1, x2, Y2
End Sub


Private Property Let IPrePrinterObj_ScaleHeight(ByVal RHS As Single)
    ScaleHeight = RHS
End Property

Private Property Get IPrePrinterObj_ScaleHeight() As Single
    IPrePrinterObj_ScaleHeight = ScaleHeight
End Property


Private Property Let IPrePrinterObj_ScaleLeft(ByVal RHS As Single)
    ScaleLeft = RHS
End Property

Private Property Get IPrePrinterObj_ScaleLeft() As Single
    IPrePrinterObj_ScaleLeft = ScaleLeft
End Property


Private Property Let IPrePrinterObj_ScaleMode(ByVal RHS As IvbExtra.cefScaleModeConstants)
    ScaleMode = RHS
End Property

Private Property Get IPrePrinterObj_ScaleMode() As IvbExtra.cefScaleModeConstants
    IPrePrinterObj_ScaleMode = ScaleMode
End Property


Private Property Let IPrePrinterObj_ScaleTop(ByVal RHS As Single)
    ScaleTop = RHS
End Property

Private Property Get IPrePrinterObj_ScaleTop() As Single
    IPrePrinterObj_ScaleTop = ScaleTop
End Property


Private Property Let IPrePrinterObj_ScaleWidth(ByVal RHS As Single)
    ScaleWidth = RHS
End Property

Private Property Get IPrePrinterObj_ScaleWidth() As Single
    IPrePrinterObj_ScaleWidth = ScaleWidth
End Property


Private Function IPrePrinterObj_ScaleX(ByVal Width As Single, Optional ByVal FromScale As IvbExtra.cefScaleConversionConstantsPrinter = -1&, Optional ByVal ToScale As IvbExtra.cefScaleConversionConstantsPrinter = -1&) As Single
    IPrePrinterObj_ScaleX = ScaleX(Width, FromScale, ToScale)
End Function


Private Function IPrePrinterObj_ScaleY(ByVal Height As Single, Optional ByVal FromScale As IvbExtra.cefScaleConversionConstantsPrinter = -1&, Optional ByVal ToScale As IvbExtra.cefScaleConversionConstantsPrinter = -1&) As Single
    IPrePrinterObj_ScaleY = ScaleY(Height, FromScale, ToScale)
End Function


Private Function IPrePrinterObj_TextHeight(Str As String) As Single
    IPrePrinterObj_TextHeight = IPrePrinterObj_TextHeightEx(Str, , , , , , , False)
    If mScaleMode = vbTwips Then
        IPrePrinterObj_TextHeight = Round(IPrePrinterObj_TextHeight)
    End If
End Function


Private Function IPrePrinterObj_TextHeightEx(Str As String, Optional ByVal Options As IvbExtra.cefPrePrintOptionsConstants, Optional ByVal Left As Variant, Optional ByVal Top As Variant, Optional ByVal Width As Variant, Optional ByVal Alignment As IvbExtra.cefAligmentConstants, Optional ByVal WordWrap As Boolean = True, Optional ByVal WordBreak As Boolean = True, Optional ByVal CellMarginsX As Single, Optional ByVal CellMarginsY As Single, Optional StartsNewPage As Boolean) As Single
    Dim iAddLineFeed As Boolean
    Dim iAddTab As Boolean
    Dim iCellMarginsX As Single
    Dim iCellMarginsY As Single
    
    If (Options And ppNoLineFeed) = 0 Then
        iAddLineFeed = True
    End If
    If (Options And ppAddTab) <> 0 Then
        iAddTab = True
    End If
    
    iCellMarginsX = CellMarginsX
    iCellMarginsY = CellMarginsY
    
    If iAddTab Then
        IPrePrinterObj_TextHeightEx = PrintText(Str & vbTab, Left, Top, Width, , Alignment, iAddLineFeed, WordWrap, WordBreak, iCellMarginsX, iCellMarginsY, StartsNewPage, True)
    Else
        IPrePrinterObj_TextHeightEx = PrintText(Str, Left, Top, Width, , Alignment, iAddLineFeed, WordWrap, WordBreak, iCellMarginsX, iCellMarginsY, StartsNewPage, True)
    End If
    
End Function


Private Function IPrePrinterObj_TextWidth(Str As String) As Single
    IPrePrinterObj_TextWidth = TextWidth(Str)
    If mScaleMode = vbTwips Then
        IPrePrinterObj_TextWidth = Round(IPrePrinterObj_TextWidth)
    End If
End Function


Private Function IPrePrinterObj_TextWidthEx(Str As String, Optional ByVal Options As IvbExtra.cefPrePrintOptionsConstants, Optional ByVal Left As Variant, Optional ByVal Top As Variant, Optional ByVal Height As Variant, Optional ByVal Alignment As IvbExtra.cefAligmentConstants, Optional ByVal WordWrap As Boolean = True, Optional ByVal WordBreak As Boolean = True, Optional ByVal CellMarginsX As Single, Optional ByVal CellMarginsY As Single, Optional StartsNewPage As Boolean) As Single
    Dim iAddLineFeed As Boolean
    Dim iAddTab As Boolean
    Dim iCellMarginsX As Single
    Dim iCellMarginsY As Single
    Dim iRequiredWidth As Single
    
    If (Options And ppNoLineFeed) = 0 Then
        iAddLineFeed = True
    End If
    If (Options And ppAddTab) <> 0 Then
        iAddTab = True
    End If
    
    iCellMarginsX = CellMarginsX
    iCellMarginsY = CellMarginsY
    
    If iAddTab Then
        PrintText Str & vbTab, Left, Top, , Height, Alignment, iAddLineFeed, WordWrap, WordBreak, iCellMarginsX, iCellMarginsY, StartsNewPage, True, iRequiredWidth
    Else
        PrintText Str, Left, Top, , Height, Alignment, iAddLineFeed, WordWrap, WordBreak, iCellMarginsX, iCellMarginsY, StartsNewPage, True, iRequiredWidth
    End If
    IPrePrinterObj_TextWidthEx = iRequiredWidth
End Function


Private Property Get IPrePrinterObj_TopMargin() As Long
    IPrePrinterObj_TopMargin = ScaleX(TopMarginInPixels, vbPixels, ScaleMode)
End Property

Private Property Get IPrePrinterObj_TwipsPerPixelX() As Single
    IPrePrinterObj_TwipsPerPixelX = TwipsPerPixelX
End Property


Private Property Get IPrePrinterObj_TwipsPerPixelY() As Single
    IPrePrinterObj_TwipsPerPixelY = TwipsPerPixelY
End Property


Private Property Get IPrePrinterObj_Width() As Single
    IPrePrinterObj_Width = Width
End Property

Private Sub SetPageSize()
    Dim iPs As POINTAPI
    
    iPs = GetPaperSize(mPaperSize)
    mPaperWidthSettingInMillimeters = iPs.x / 10
    mPaperHeightSettingInMillimeters = iPs.y / 10
End Sub

Private Property Get PaperWidthSettingInMillimeters()
    If mPaperWidthSettingInMillimeters = 0 Then
        PaperWidthSettingInMillimeters = mPrinterDefaultPageWidthInMillimeters
    Else
        PaperWidthSettingInMillimeters = mPaperWidthSettingInMillimeters
    End If
End Property

Private Property Get PaperHeightSettingInMillimeters()
    If mPaperHeightSettingInMillimeters = 0 Then
        PaperHeightSettingInMillimeters = mPrinterDefaultPageHeightInMillimeters
    Else
        PaperHeightSettingInMillimeters = mPaperHeightSettingInMillimeters
    End If
End Property

Private Property Get CurrentPageScaleWidthInPixels() As Long
    EnsureDocumentCreated
    If mAutoHandleMargins Then
        CurrentPageScaleWidthInPixels = mPages(mCurrentPageNumber).ScaleWidthInPixels
    Else
        CurrentPageScaleWidthInPixels = mPages(mCurrentPageNumber).ScaleWidthInPixelsNoMargins
    End If
End Property

Private Property Get CurrentPageScaleHeightInPixels() As Long
    EnsureDocumentCreated
    If mAutoHandleMargins Then
        CurrentPageScaleHeightInPixels = mPages(mCurrentPageNumber).ScaleHeightInPixels
    Else
        CurrentPageScaleHeightInPixels = mPages(mCurrentPageNumber).ScaleHeightInPixelsNoMargins
    End If
End Property


Private Function GetPaperSize(nPaperSizeNumber As Long) As POINTAPI
    Dim ret As Long
    Dim iPaperSizesNumbers() As Integer
    Dim iPaperSizes() As POINTAPI
    Dim c As Long
    Dim iLng As Long
    
    ret = DeviceCapabilities(mDeviceName, mPort, DC_PAPERS, ByVal 0&, ByVal 0&)
    
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
    Dim iPrn As Printer
    Dim c As Long
    Dim iDeviceNamePrev As String
    
    If Printer.DeviceName <> mDeviceName Then
        iDeviceNamePrev = Printer.DeviceName
        For Each iPrn In Printers
            If iPrn.DeviceName = mDeviceName Then
                Set Printer = iPrn
                Exit For
            End If
        Next
    End If
    
    ReDim mFonts(Printer.FontCount - 1)
    For c = 0 To Printer.FontCount - 1
        mFonts(c) = Printer.Fonts(c)
    Next c
    
    If iDeviceNamePrev <> "" Then
        For Each iPrn In Printers
            If iPrn.DeviceName = iDeviceNamePrev Then
                Set Printer = iPrn
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
        Call nFont.SetRatio(mPages(mCurrentPageNumber).MetaDPIY, 2540)
    Else
        Call nFont.SetRatio(GetDeviceCaps(Printer.hdc, LOGPIXELSY), 2540)
    End If
End Sub

Public Sub KillLastPage()
    mNewPagePending = False
    mCurrentPageNumber = mCurrentPageNumber - 1
    If mCurrentPageNumber < 0 Then mCurrentPageNumber = 0
    ReDim Preserve mPages(mCurrentPageNumber)
    If mCurrentPageNumber = 0 Then
        mPageCreated = False
    End If
    mSomethingPrintedOnPage = False
    mCurrentXPixels = 0
    mCurrentYPixels = 0
End Sub

Private Function StdFontToLogFont_Printer(nFont As StdFont, Optional nFontName As String) As LOGFONT
    Dim iFontName As String
    Dim iDPIY As Single
    Dim c As Long
    
    If nFontName = "" Then
        iFontName = nFont.Name
    Else
        iFontName = nFontName
    End If
    
    If mDocumentCreated Then
        CheckNewPagePending
        iDPIY = mPages(mCurrentPageNumber).MetaDPIY
    Else
        iDPIY = GetDeviceCaps(Printer.hdc, LOGPIXELSY)
    End If
    
    For c = 0 To 31
        If c < Len(iFontName) Then
            StdFontToLogFont_Printer.lfFaceName(c) = Asc(Mid$(iFontName, c + 1, 1))
        Else
            StdFontToLogFont_Printer.lfFaceName(c) = 0
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

Private Function LogFontToStdFont(lF As LOGFONT) As iFont
    Dim iFontName As String
    Dim iDPIY As Single
    
    Set LogFontToStdFont = New StdFont
    SetFontRatioForPrinterStdFont LogFontToStdFont
    
    If mDocumentCreated Then
        CheckNewPagePending
        iDPIY = mPages(mCurrentPageNumber).MetaDPIY
    Else
        iDPIY = GetDeviceCaps(Printer.hdc, LOGPIXELSY)
    End If
    
    
    iFontName = StrConv(lF.lfFaceName, vbUnicode)
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
            iFontName = StrConv(mLogFont.lfFaceName, vbUnicode)
            If InStr(iFontName, Chr(0)) > 0 Then
                iFontName = Left$(iFontName, InStr(iFontName, Chr(0)) - 1)
            End If
            GetFontAttribute = iFontName
        Case "Size"
            If mDocumentCreated Then
                CheckNewPagePending
                iDPIY = mPages(mCurrentPageNumber).MetaDPIY
            Else
                iDPIY = GetDeviceCaps(Printer.hdc, LOGPIXELSY)
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
                iDPIY = mPages(mCurrentPageNumber).MetaDPIY
            Else
                iDPIY = GetDeviceCaps(Printer.hdc, LOGPIXELSY)
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
            iFontName = nValue
            For c = 0 To 31
                If c < Len(iFontName) Then
                    mLogFont.lfFaceName(c) = Asc(Mid$(iFontName, c + 1, 1))
                Else
                    mLogFont.lfFaceName(c) = 0
                End If
            Next
            mFontName = nValue
            mFontCharset = DefaultCharsetForFontName(mFontName)
            mLogFont.lfCharSet = mFontCharset
        Case "Size"
            If mDocumentCreated Then
                CheckNewPagePending
                iDPIY = mPages(mCurrentPageNumber).MetaDPIY
            Else
                iDPIY = GetDeviceCaps(Printer.hdc, LOGPIXELSY)
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
                iDPIY = mPages(mCurrentPageNumber).MetaDPIY
            Else
                iDPIY = GetDeviceCaps(Printer.hdc, LOGPIXELSY)
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
    MsgBox GetLocalizedString(efnGUIStr_cPrePrintDocument_PrintDocumentTooLongWarning_MsgBoxWarning), vbExclamation, ClientProductName
End Sub

Private Sub mFont_FontChanged(ByVal PropertyName As String)
    Set Font = mFont
End Sub
