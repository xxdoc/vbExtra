VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPrinterExPage"
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
Private mWidthInPixels As Long
Private mHeightInPixels As Long
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
Private mOrientation As Long
Private mSomethingPrintedOnPage As Boolean
Private mNonPrintableAreaLeftInPixels As Long
Private mNonPrintableAreaTopInPixels As Long
Private mNonPrintableAreaRightInPixels As Long
Private mNonPrintableAreaBottomInPixels As Long
Private mZoom As Long
Private mFixedImageScale As Single
Private mDocumentName As String

Public Sub CreatePage(nHdcReference As Long, nWidthInMillimeters, nHeightInMillimeters, nOrientation As Long, nhFont As Long, nDrawStyle As Long, nDrawWidth As Long, nForeColor As Long, nDrawMode As Long, nFontTransparent As Boolean, nFillStyle As Long, nFillColor As Long, nLeftMarginInMillimeters, nTopMarginInMillimeters, nRightMarginInMillimeters, nBottomMarginInMillimeters, nPaperSize As Long, nPaperBin As Long, nPrintQuality As Long, nDuplex As Long, nColorMode As Long, nNonPrintableAreaLeftInPixels, nNonPrintableAreaTopInPixels, nNonPrintableAreaRightInPixels, nNonPrintableAreaBottomInPixels, nDPIX, nDPIY, nZoom As Long, nFixedImageScale As Single, nDocumentName As String)
    Dim iRect As RECT
    Dim iLb As LOGBRUSH
    
    mZoom = nZoom
    mOrientation = nOrientation
    If mOrientation <> vbPRORLandscape Then
        iRect.Right = nWidthInMillimeters * 100 / mZoom * 100
        iRect.Bottom = nHeightInMillimeters * 100 / mZoom * 100
        
        mWidthInMillimeters = nWidthInMillimeters
        mHeightInMillimeters = nHeightInMillimeters
        
        mMetaDPIX = nDPIX
        mMetaDPIY = nDPIY
        
        mNonPrintableAreaLeftInPixels = nNonPrintableAreaLeftInPixels / mZoom * 100
        mNonPrintableAreaTopInPixels = nNonPrintableAreaTopInPixels / mZoom * 100
        mNonPrintableAreaRightInPixels = nNonPrintableAreaRightInPixels / mZoom * 100
        mNonPrintableAreaBottomInPixels = nNonPrintableAreaBottomInPixels / mZoom * 100
    Else
        iRect.Right = nHeightInMillimeters * 100 / mZoom * 100
        iRect.Bottom = nWidthInMillimeters * 100 / mZoom * 100
    
        mWidthInMillimeters = nHeightInMillimeters
        mHeightInMillimeters = nWidthInMillimeters
    
        mMetaDPIX = nDPIY
        mMetaDPIY = nDPIX
    
        mNonPrintableAreaLeftInPixels = nNonPrintableAreaTopInPixels
        mNonPrintableAreaTopInPixels = nNonPrintableAreaLeftInPixels
        mNonPrintableAreaRightInPixels = nNonPrintableAreaBottomInPixels
        mNonPrintableAreaBottomInPixels = nNonPrintableAreaRightInPixels
    End If
    
    mDocumentName = nDocumentName
    mMetaDC = CreateEnhMetaFile(nHdcReference, vbNullString, iRect, vbNullString)
    
    mWidthInPixels = mMetaDPIX * mWidthInMillimeters / 25.4 / mZoom * 100
    mHeightInPixels = mMetaDPIY * mHeightInMillimeters / 25.4 / mZoom * 100
    mPrintableAreaWidthInPixels = mWidthInPixels - mNonPrintableAreaLeftInPixels - mNonPrintableAreaRightInPixels
    mPrintableAreaHeightInPixels = mHeightInPixels - mNonPrintableAreaTopInPixels - mNonPrintableAreaBottomInPixels
    
    mOldFont = SelectObject(mMetaDC, nhFont)
    
    mFontTransparent = nFontTransparent
    SetBackMode
    
    mPen = CreatePen(nDrawStyle, nDrawWidth, nForeColor)
    mOldPen = SelectObject(mMetaDC, mPen)
    SetROP2 mMetaDC, nDrawMode
    
    Select Case nFillStyle
        Case vbFSSolid ' 0
            iLb.lbStyle = BS_SOLID ' 0
        Case vbFSTransparent ' 1
            iLb.lbStyle = BS_NULL ' 1
        Case Else
            iLb.lbStyle = BS_HATCHED ' 2
            iLb.lbHatch = nFillStyle - 2
    End Select
    iLb.lbColor = nFillColor
    
    mBrush = CreateBrushIndirect(iLb)
    mOldBrush = SelectObject(mMetaDC, mBrush)
    
    mLeftMarginInMillimeters = nLeftMarginInMillimeters
    mTopMarginInMillimeters = nTopMarginInMillimeters
    mRightMarginInMillimeters = nRightMarginInMillimeters
    mBottomMarginInMillimeters = nBottomMarginInMillimeters
    
    mLeftMarginInPixels = mLeftMarginInMillimeters * mMetaDPIX / 25.4 / mZoom * 100
    mTopMarginInPixels = mTopMarginInMillimeters * mMetaDPIY / 25.4 / mZoom * 100
    mRightMarginInPixels = mRightMarginInMillimeters * mMetaDPIX / 25.4 / mZoom * 100
    mBottomMarginInPixels = mBottomMarginInMillimeters * mMetaDPIY / 25.4 / mZoom * 100
    
    SetScale
    
    mPaperSize = nPaperSize
    mPaperBin = nPaperBin
    mPrintQuality = nPrintQuality
    mDuplex = nDuplex
    mColorMode = nColorMode
    mFixedImageScale = nFixedImageScale
    
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
    Dim iLb As LOGBRUSH
    Dim iBrushPrev As Long
    
    If mMetaDC = 0 Then Exit Sub
    
    iBrushPrev = mBrush
    Select Case nFillStyle
        Case vbFSSolid ' 0
            iLb.lbStyle = BS_SOLID ' 0
        Case vbFSTransparent ' 1
            iLb.lbStyle = BS_NULL ' 1
        Case Else
            iLb.lbStyle = BS_HATCHED ' 2
            iLb.lbHatch = nFillStyle - 2
    End Select
    iLb.lbColor = nFillColor
    
    mBrush = CreateBrushIndirect(iLb)
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


Public Property Get WidthInPixels() As Long
    WidthInPixels = mWidthInPixels
End Property


Public Property Get HeightInPixels() As Long
    HeightInPixels = mHeightInPixels
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
    mLeftMarginInPixels = mLeftMarginInMillimeters * mMetaDPIX / 25.4 / mZoom * 100
    SetScale
End Property

Public Property Get LeftMarginInMillimeters() As Single
    LeftMarginInMillimeters = mLeftMarginInMillimeters
End Property


Public Property Let TopMarginInMillimeters(nValue As Single)
    mTopMarginInMillimeters = nValue
    mTopMarginInPixels = mTopMarginInMillimeters * mMetaDPIY / 25.4 / mZoom * 100
    SetScale
End Property

Public Property Get TopMarginInMillimeters() As Single
    TopMarginInMillimeters = mTopMarginInMillimeters
End Property


Public Property Let RightMarginInMillimeters(nValue As Single)
    mRightMarginInMillimeters = nValue
    mRightMarginInPixels = mRightMarginInMillimeters * mMetaDPIX / 25.4 / mZoom * 100
    SetScale
End Property

Public Property Get RightMarginInMillimeters() As Single
    RightMarginInMillimeters = mRightMarginInMillimeters
End Property


Public Property Let BottomMarginInMillimeters(nValue As Single)
    mBottomMarginInMillimeters = nValue
    mBottomMarginInPixels = mBottomMarginInMillimeters * mMetaDPIY / 25.4 / mZoom * 100
    SetScale
End Property

Public Property Get BottomMarginInMillimeters() As Single
    BottomMarginInMillimeters = mBottomMarginInMillimeters
End Property


Private Sub SetScale()
    mScaleWidthInPixels = mWidthInPixels - mLeftMarginInPixels - mRightMarginInPixels
    mScaleHeightInPixels = mHeightInPixels - mTopMarginInPixels - mBottomMarginInPixels
    mScaleWidthInPixelsNoMargins = mPrintableAreaWidthInPixels
    mScaleHeightInPixelsNoMargins = mPrintableAreaHeightInPixels
End Sub


Public Property Get Orientation() As Integer
    Orientation = mOrientation
End Property


Public Property Get PaperSize() As Integer
    PaperSize = mPaperSize
End Property

Public Property Get PaperBin() As Integer
    PaperBin = mPaperBin
End Property

Public Property Get PrintQuality() As Integer
    PrintQuality = mPrintQuality
End Property

Public Property Get Duplex() As Integer
    Duplex = mDuplex
End Property

Public Property Get ColorMode() As Integer
    ColorMode = mColorMode
End Property


Public Property Let SomethingPrintedOnPage(nValue As Boolean)
    mSomethingPrintedOnPage = nValue
End Property

Public Property Get SomethingPrintedOnPage() As Boolean
    SomethingPrintedOnPage = mSomethingPrintedOnPage
End Property

Public Property Get Zoom() As Long
    Zoom = mZoom
End Property

Public Property Get NonPrintableAreaLeftInPixels() As Long
    NonPrintableAreaLeftInPixels = mNonPrintableAreaLeftInPixels
End Property
    
Public Property Get NonPrintableAreaTopInPixels() As Long
    NonPrintableAreaTopInPixels = mNonPrintableAreaTopInPixels
End Property
    
Public Property Get NonPrintableAreaRightInPixels() As Long
    NonPrintableAreaRightInPixels = mNonPrintableAreaRightInPixels
End Property
    
Public Property Get NonPrintableAreaBottomInPixels() As Long
    NonPrintableAreaBottomInPixels = mNonPrintableAreaBottomInPixels
End Property


Public Property Let FixedImageScale(nValue As Single)
    mFixedImageScale = nValue
End Property

Public Property Get FixedImageScale() As Single
    FixedImageScale = mFixedImageScale
End Property


Public Property Let DocumentName(nValue As String)
    mDocumentName = nValue
End Property

Public Property Get DocumentName() As String
    DocumentName = mDocumentName
End Property

