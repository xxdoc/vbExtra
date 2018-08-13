VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrePrintPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As Long, ByVal lpFileName As String, lpRect As RECT, ByVal lpDescription As String) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hemf As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Const OPAQUE As Long = 2
Private Const TRANSPARENT As Long = 1

Private Declare Function CreatePen Lib "gdi32" (ByVal nDrawStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'Private Const PS_SOLID As Long = 0

Private Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long

Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private Const BS_SOLID As Long = 0
Private Const BS_NULL As Long = 1
Private Const BS_HATCHED As Long = 2

Private Declare Function CreateBrushIndirect Lib "gdi32.dll" (lpLogBrush As LOGBRUSH) As Long

Private mOldFont As Long
Private mOldPen As Long
Private mPen As Long
Private mOldBrush As Long
Private mBrush As Long
Private mFontTransparent As Boolean

Private mMetaDC As Long
Private mhMetaFile As Long
Private mMetaDPIX As Long
Private mMetaDPIY As Long

Private mLeftMarginInPixels As Single
Private mTopMarginInPixels As Single
Private mRightMarginInPixels As Single
Private mBottomMarginInPixels As Single

Private mWidthInMillimeters As Single
Private mHeightInMillimeters As Single
Private mPageWidthInPixels As Long
Private mPageHeightInPixels As Long
Private mPrintableAreaWidthInPixels As Long
Private mPrintableAreaHeightInPixels As Long
Private mScaleWidthInPixels As Long
Private mScaleHeightInPixels As Long
Private mScaleWidthInPixelsNoMargins As Long
Private mScaleHeightInPixelsNoMargins As Long
Private mLeftMarginInMillimeters As Single
Private mTopMarginInMillimeters As Single
Private mRightMarginInMillimeters As Single
Private mBottomMarginInMillimeters As Single
Private mPaperSize As Long
Private mPaperBin As Long
Private mPrintQuality As Long
Private mDuplex As Long
Private mColorMode As Long

Public Sub CreatePage(nHdcReference As Long, nWidthInMillimeters, nHeightInMillimeters, nhFont As Long, nDrawStyle As Long, nDrawWidth As Long, nForeColor As Long, nDrawMode As Long, nFontTransparent As Boolean, nFillStyle As Long, nFillColor As Long, nLeftMarginInMillimeters, nTopMarginInMillimeters, nRightMarginInMillimeters, nBottomMarginInMillimeters, nPaperSize As Long, nPaperBin As Long, nPrintQuality As Long, nDuplex As Long, nColorMode As Long)
    Dim iRect As RECT
    Dim iLB As LOGBRUSH
    Dim iDefaultPageWidthInPixels As Long
    Dim iDefaultPageHeightInPixels As Long
    Dim iDefaultPrintableAreaWidthInPixels As Long
    Dim iDefaultPrintableAreaHeightInPixels As Long
    Dim iNonPrintableAreaAtLeftInPixels As Long
    Dim iNonPrintableAreaAtTopInPixels As Long
    Dim iNonPrintableAreaAtRightInPixels As Long
    Dim iNonPrintableAreaAtBottomInPixels As Long
    Dim iAuxMetaDPIX As Single
    Dim iAuxMetaDPIY As Single
    
    iRect.Right = nWidthInMillimeters * 100
    iRect.Bottom = nHeightInMillimeters * 100
    
    mWidthInMillimeters = nWidthInMillimeters
    mHeightInMillimeters = nHeightInMillimeters
    
    mMetaDC = CreateEnhMetaFile(nHdcReference, vbNullString, iRect, vbNullString)
    
    If Orientation = vbPRORPortrait Then
        mMetaDPIX = 25.4 * (GetDeviceCaps(mMetaDC, HORZRES)) / GetDeviceCaps(mMetaDC, HORZSIZE)
        mMetaDPIY = 25.4 * (GetDeviceCaps(mMetaDC, VERTRES)) / GetDeviceCaps(mMetaDC, VERTSIZE)
        iAuxMetaDPIX = 25.4 * (GetDeviceCaps(mMetaDC, PHYSICALWIDTH)) / GetDeviceCaps(mMetaDC, HORZSIZE)
        iAuxMetaDPIY = 25.4 * (GetDeviceCaps(mMetaDC, PHYSICALHEIGHT)) / GetDeviceCaps(mMetaDC, VERTSIZE)
        
        iDefaultPageWidthInPixels = GetDeviceCaps(nHdcReference, PHYSICALWIDTH)
        iDefaultPageHeightInPixels = GetDeviceCaps(nHdcReference, PHYSICALHEIGHT)
        iDefaultPrintableAreaWidthInPixels = GetDeviceCaps(nHdcReference, HORZRES)
        iDefaultPrintableAreaHeightInPixels = GetDeviceCaps(nHdcReference, VERTRES)
        
        iNonPrintableAreaAtTopInPixels = GetDeviceCaps(nHdcReference, PHYSICALOFFSETY)
        iNonPrintableAreaAtBottomInPixels = iDefaultPageHeightInPixels - iDefaultPrintableAreaHeightInPixels - iNonPrintableAreaAtTopInPixels
        iNonPrintableAreaAtLeftInPixels = GetDeviceCaps(nHdcReference, PHYSICALOFFSETX)
        iNonPrintableAreaAtRightInPixels = iDefaultPageWidthInPixels - iDefaultPrintableAreaWidthInPixels - iNonPrintableAreaAtLeftInPixels
    
        mPageWidthInPixels = iAuxMetaDPIX * mWidthInMillimeters / 25.4
        mPageHeightInPixels = iAuxMetaDPIY * mHeightInMillimeters / 25.4
        mPrintableAreaWidthInPixels = mPageWidthInPixels - iNonPrintableAreaAtLeftInPixels - iNonPrintableAreaAtRightInPixels
        mPrintableAreaHeightInPixels = mPageHeightInPixels - iNonPrintableAreaAtTopInPixels - iNonPrintableAreaAtBottomInPixels
        
    Else
        mMetaDPIY = 25.4 * (GetDeviceCaps(mMetaDC, HORZRES)) / GetDeviceCaps(mMetaDC, HORZSIZE)
        mMetaDPIX = 25.4 * (GetDeviceCaps(mMetaDC, VERTRES)) / GetDeviceCaps(mMetaDC, VERTSIZE)
        iAuxMetaDPIY = 25.4 * (GetDeviceCaps(mMetaDC, PHYSICALWIDTH)) / GetDeviceCaps(mMetaDC, HORZSIZE)
        iAuxMetaDPIX = 25.4 * (GetDeviceCaps(mMetaDC, PHYSICALHEIGHT)) / GetDeviceCaps(mMetaDC, VERTSIZE)
        
        iDefaultPageHeightInPixels = GetDeviceCaps(nHdcReference, PHYSICALWIDTH)
        iDefaultPageWidthInPixels = GetDeviceCaps(nHdcReference, PHYSICALHEIGHT)
        iDefaultPrintableAreaHeightInPixels = GetDeviceCaps(nHdcReference, HORZRES)
        iDefaultPrintableAreaWidthInPixels = GetDeviceCaps(nHdcReference, VERTRES)
        
        iNonPrintableAreaAtLeftInPixels = GetDeviceCaps(nHdcReference, PHYSICALOFFSETY)
        iNonPrintableAreaAtTopInPixels = GetDeviceCaps(nHdcReference, PHYSICALOFFSETX)
        iNonPrintableAreaAtBottomInPixels = iDefaultPageHeightInPixels - iDefaultPrintableAreaHeightInPixels - iNonPrintableAreaAtTopInPixels
        iNonPrintableAreaAtRightInPixels = iDefaultPageWidthInPixels - iDefaultPrintableAreaWidthInPixels - iNonPrintableAreaAtLeftInPixels
        
        mPageWidthInPixels = iAuxMetaDPIX * mWidthInMillimeters / 25.4
        mPageHeightInPixels = iAuxMetaDPIY * mHeightInMillimeters / 25.4
        mPrintableAreaWidthInPixels = mPageWidthInPixels - iNonPrintableAreaAtLeftInPixels - iNonPrintableAreaAtRightInPixels
        mPrintableAreaHeightInPixels = mPageHeightInPixels - iNonPrintableAreaAtTopInPixels - iNonPrintableAreaAtBottomInPixels
    End If
    
    mOldFont = SelectObject(mMetaDC, nhFont)
    
    mFontTransparent = nFontTransparent
    SetBackMode
    
    mPen = CreatePen(nDrawStyle, nDrawWidth, nForeColor)
    mOldPen = SelectObject(mMetaDC, mPen)
    SetROP2 mMetaDC, nDrawMode
    
    Select Case nFillStyle
        Case vbFSSolid ' 0
            iLB.lbStyle = BS_SOLID ' 0
        Case vbFSTransparent ' 1
            iLB.lbStyle = BS_NULL ' 1
        Case Else
            iLB.lbStyle = BS_HATCHED ' 2
            iLB.lbHatch = nFillStyle - 2
    End Select
    iLB.lbColor = nFillColor
    
    mBrush = CreateBrushIndirect(iLB)
    mOldBrush = SelectObject(mMetaDC, mBrush)

    mLeftMarginInMillimeters = nLeftMarginInMillimeters
    mTopMarginInMillimeters = nTopMarginInMillimeters
    mRightMarginInMillimeters = nRightMarginInMillimeters
    mBottomMarginInMillimeters = nBottomMarginInMillimeters
    
'    If (mLeftMarginInMillimeters * iAuxMetaDPIX / 25.4) < iNonPrintableAreaAtLeftInPixels Then
'        mLeftMarginInMillimeters = iNonPrintableAreaAtLeftInPixels / iAuxMetaDPIX * 25.4
'    End If
'    If (mRightMarginInMillimeters * iAuxMetaDPIX / 25.4) < iNonPrintableAreaAtRightInPixels Then
'        mRightMarginInMillimeters = iNonPrintableAreaAtRightInPixels / iAuxMetaDPIX * 25.4
'    End If
'    If (mTopMarginInMillimeters * iAuxMetaDPIY / 25.4) < iNonPrintableAreaAtTopInPixels Then
'        mTopMarginInMillimeters = iNonPrintableAreaAtTopInPixels / iAuxMetaDPIY * 25.4
'    End If
'    If (mBottomMarginInMillimeters * iAuxMetaDPIY / 25.4) < iNonPrintableAreaAtBottomInPixels Then
'        mBottomMarginInMillimeters = iNonPrintableAreaAtBottomInPixels / iAuxMetaDPIY * 25.4
'    End If
    
    mLeftMarginInPixels = mLeftMarginInMillimeters * mMetaDPIX / 25.4
    mTopMarginInPixels = mTopMarginInMillimeters * mMetaDPIY / 25.4
    mRightMarginInPixels = mRightMarginInMillimeters * mMetaDPIX / 25.4
    mBottomMarginInPixels = mBottomMarginInMillimeters * mMetaDPIY / 25.4
    
    SetScale
    
    mPaperSize = nPaperSize
    mPaperBin = nPaperBin
    mPrintQuality = nPrintQuality
    mDuplex = nDuplex
    mColorMode = nColorMode
    
End Sub

Private Sub Class_Terminate()
    If mMetaDC <> 0 Then
        Call SelectObject(mMetaDC, mOldFont)
        Call SelectObject(mMetaDC, mOldPen)
        DeleteObject mPen
        mhMetaFile = CloseEnhMetaFile(mMetaDC)
    End If
    DeleteEnhMetaFile mhMetaFile
    
End Sub

Public Property Get hMetaFile() As Long
    If mhMetaFile = 0 Then
        If mMetaDC <> 0 Then
            Call SelectObject(mMetaDC, mOldFont)
            mhMetaFile = CloseEnhMetaFile(mMetaDC)
            mMetaDC = 0
        End If
    End If
    hMetaFile = mhMetaFile
End Property

Public Property Get MetaDPIX() As Long
    MetaDPIX = mMetaDPIX
End Property

Public Property Get MetaDPIY() As Long
    MetaDPIY = mMetaDPIY
End Property

Public Sub SelectFont(nhFont As Long)
    If mMetaDC <> 0 Then
        Call SelectObject(mMetaDC, nhFont)
    End If
End Sub

Public Property Get MetaDC() As Long
    MetaDC = mMetaDC
End Property

Public Sub SetPen(nDrawStyle As Long, nDrawWidth As Long, nForeColor As Long, nDrawMode As Long)
    Dim iPenPrev As Long
    
    If mMetaDC = 0 Then Exit Sub
    
    iPenPrev = mPen
    mPen = CreatePen(nDrawStyle, nDrawWidth, nForeColor)
    Call SelectObject(mMetaDC, mPen)
    DeleteObject iPenPrev
    
    SetROP2 mMetaDC, nDrawMode
End Sub

Public Sub SetBrush(nFillStyle As Long, nFillColor As Long)
    Dim iLB As LOGBRUSH
    Dim iBrushPrev As Long
    
    If mMetaDC = 0 Then Exit Sub
    
    iBrushPrev = mBrush
    Select Case nFillStyle
        Case vbFSSolid ' 0
            iLB.lbStyle = BS_SOLID ' 0
        Case vbFSTransparent ' 1
            iLB.lbStyle = BS_NULL ' 1
        Case Else
            iLB.lbStyle = BS_HATCHED ' 2
            iLB.lbHatch = nFillStyle - 2
    End Select
    iLB.lbColor = nFillColor
    
    mBrush = CreateBrushIndirect(iLB)
    Call SelectObject(mMetaDC, mBrush)
    DeleteObject iBrushPrev
End Sub

Public Property Get FontTransparent() As Boolean
    FontTransparent = mFontTransparent
End Property

Public Property Let FontTransparent(nValue As Boolean)
    mFontTransparent = nValue
    SetBackMode
End Property

Private Sub SetBackMode()
    If mFontTransparent Then
        SetBkMode mMetaDC, TRANSPARENT
    Else
        SetBkMode mMetaDC, OPAQUE
    End If
End Sub


Public Property Get LeftMarginInPixels() As Single
    LeftMarginInPixels = mLeftMarginInPixels
End Property


Public Property Get TopMarginInPixels() As Single
    TopMarginInPixels = mTopMarginInPixels
End Property


Public Property Get RightMarginInPixels() As Single
    RightMarginInPixels = mRightMarginInPixels
End Property


Public Property Get BottomMarginInPixels() As Single
    BottomMarginInPixels = mBottomMarginInPixels
End Property


Public Property Get WidthInMillimeters() As Single
    WidthInMillimeters = mWidthInMillimeters
End Property


Public Property Get HeightInMillimeters() As Single
    HeightInMillimeters = mHeightInMillimeters
End Property


Public Property Get PageWidthInPixels() As Long
    PageWidthInPixels = mPageWidthInPixels
End Property


Public Property Get PageHeightInPixels() As Long
    PageHeightInPixels = mPageHeightInPixels
End Property


Public Property Get ScaleWidthInPixels() As Long
    ScaleWidthInPixels = mScaleWidthInPixels
End Property


Public Property Get ScaleHeightInPixels() As Long
    ScaleHeightInPixels = mScaleHeightInPixels
End Property


Public Property Get ScaleWidthInPixelsNoMargins() As Long
    ScaleWidthInPixelsNoMargins = mScaleWidthInPixelsNoMargins
End Property


Public Property Get ScaleHeightInPixelsNoMargins() As Long
    ScaleHeightInPixelsNoMargins = mScaleHeightInPixelsNoMargins
End Property


Public Property Let LeftMarginInMillimeters(nValue As Single)
    mLeftMarginInMillimeters = nValue
    mLeftMarginInPixels = mLeftMarginInMillimeters * mMetaDPIX / 25.4
    SetScale
End Property

Public Property Get LeftMarginInMillimeters() As Single
    LeftMarginInMillimeters = mLeftMarginInMillimeters
End Property


Public Property Let TopMarginInMillimeters(nValue As Single)
    mTopMarginInMillimeters = nValue
    mTopMarginInPixels = mTopMarginInMillimeters * mMetaDPIY / 25.4
    SetScale
End Property

Public Property Get TopMarginInMillimeters() As Single
    TopMarginInMillimeters = mTopMarginInMillimeters
End Property


Public Property Let RightMarginInMillimeters(nValue As Single)
    mRightMarginInMillimeters = nValue
    mRightMarginInPixels = mRightMarginInMillimeters * mMetaDPIX / 25.4
    SetScale
End Property

Public Property Get RightMarginInMillimeters() As Single
    RightMarginInMillimeters = mRightMarginInMillimeters
End Property


Public Property Let BottomMarginInMillimeters(nValue As Single)
    mBottomMarginInMillimeters = nValue
    mBottomMarginInPixels = mBottomMarginInMillimeters * mMetaDPIY / 25.4
    SetScale
End Property

Public Property Get BottomMarginInMillimeters() As Single
    BottomMarginInMillimeters = mBottomMarginInMillimeters
End Property


Private Sub SetScale()
    mScaleWidthInPixels = mPrintableAreaWidthInPixels - mLeftMarginInPixels - mRightMarginInPixels
    mScaleHeightInPixels = mPrintableAreaHeightInPixels - mTopMarginInPixels - mBottomMarginInPixels
    mScaleWidthInPixelsNoMargins = mPrintableAreaWidthInPixels
    mScaleHeightInPixelsNoMargins = mPrintableAreaHeightInPixels
End Sub


Public Property Get Orientation() As cefPageOrientationConstants
    If mHeightInMillimeters < mWidthInMillimeters Then
        Orientation = vbPRORLandscape
    Else
        Orientation = vbPRORPortrait
    End If
End Property


Public Property Get PaperSize() As cefPaperSizeConstants
    PaperSize = mPaperSize
End Property

Public Property Get PaperBin() As cefPaperBinConstants
    PaperBin = mPaperBin
End Property

Public Property Get PrintQuality() As cefPrintQualityConstants
    PrintQuality = mPrintQuality
End Property

Public Property Get Duplex() As cefDuplexConstants
    Duplex = mDuplex
End Property

Public Property Get ColorMode() As cefColorModeConstants
    ColorMode = mColorMode
End Property
