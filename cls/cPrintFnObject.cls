VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintFnObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event PrepareDoc(Cancel As Boolean, ByVal DocKey As String)
Public Event BeforeShowingPageSetupDialog(ByRef CancelDisplay As Boolean)
Public Event AfterShowingPageSetupDialog()
Public Event BeforeShowingPrinterDialog(ByRef CancelPrint As Boolean)
Public Event AfterShowingPrinterDialog()
Public Event FormatOptionsClick(ByRef Canceled As Boolean)
Public Event ScaleChange(NewScalePercent As Integer)
Public Event OrientationChange(ByVal NewOrientation As Long)
Public Event DocPrinted(ByVal DocKey As String)

Public Event StartPage(ByVal DocKey As String)
Public Event StartDoc(ByVal DocKey As String)
Public Event EndDoc(ByVal FirstPageIndex As Long, ByVal LastPageIndex As Long, ByVal DocKey As String)

Private WithEvents mfrmPrintPreview As frmPrintPreview
Attribute mfrmPrintPreview.VB_VarHelpID = -1
Private mDlg As New CommonDialogExObject

Private mCanceled As Boolean

Public Enum vbExPageNumbersPositionConstants
    vxPositionBottomRight = 0&
    vxPositionBottomLeft = 1&
    vxPositionBottomCenter = 2&
    vxPositionTopRight = 3&
    vxPositionTopLeft = 4&
    vxPositionTopCenter = 5&
End Enum

Public Enum vbExPrintPrevToolBarIconsSizeConstants
    vxPPTIAppDefault = 0
    vxPPTIconsSmall = 1
    vxPPTIconsMedium = 2
    vxPPTIconsBig = 3
    vxPPTIconsAuto = 4
    vxPPTIconsSmallAndUserCanChange = 5
    vxPPTIconsMediumAndUserCanChange = 6
    vxPPTIconsBigAndUserCanChange = 7
End Enum

Private mHandleMargins As Boolean
Private mPrintPageNumbers As Boolean
Private mPageNumbersFont As StdFont
Private mPageNumbersPosition As Long
Private mPageNumbersFormat As String
Private mPageNumbersForeColor As Long
Private mDocumentName As String
Private mFormatButtonVisible As Boolean
Private mFormatButtonToolTipText As String
Private mFormatButtonPicture(4) As Object
Private mAllowUserChangeScale As Boolean
Private mScalePercent As Long
Private mMinScalePercent As Long
Private mMaxScalePercent As Long
Private mPrintWithDefaultSettings As Boolean
Private mDocKey As String
Private mProcedureName As String
Private mPageSetupButtonVisible As Boolean
Private mPrintPrevToolBarIconsSize As vbExPrintPrevToolBarIconsSizeConstants
Private mPrintPrevUseAltScaleIcons As Boolean
Private mDocPrepared As Boolean
Private mParent As Object
Private mPrinted As Boolean
Private WithEvents mPrinterExEvents As PrinterExEvents
Attribute mPrinterExEvents.VB_VarHelpID = -1
Private mPredefinedPageNumbersFormatStrings()  As String

Public Sub ShowPageSetup()
    Dim iCancelDisplay  As Boolean
    Dim iOrientationPrev As Long
    
    mCanceled = False
    RaiseEvent BeforeShowingPageSetupDialog(iCancelDisplay)
    If Not iCancelDisplay Then
        iOrientationPrev = mDlg.Orientation
        mDlg.ShowPageSetup
        mCanceled = mDlg.Canceled
        If Not mCanceled Then
            If iOrientationPrev <> mDlg.Orientation Then
                ResetPrinter2
                Printer2.Orientation = mDlg.Orientation
                RaiseEvent OrientationChange(mDlg.Orientation)
            End If
            RaiseEvent AfterShowingPageSetupDialog
        End If
    End If
End Sub

Public Sub PrintNow(Optional DocKey As String)
    mPrintWithDefaultSettings = True
    ShowPrint DocKey
    mPrintWithDefaultSettings = False
End Sub

Public Sub ShowPrint(Optional DocKey As String)
    Dim iCancelPrint  As Boolean
    Dim iChanged As Boolean
    Dim iPrinterNotSet As Boolean
    
    Dim iDeviceName As String
    Dim iDevModePtr As Long
    Dim iPaperSize As Long
    Dim iOrientation As Long
    Dim iPaperBin As Long
    Dim iPrintQuality As Long
    Dim iColorMode As Long
    Dim iDuplex As Long
    Dim iAuxOrientation As Long
    Dim iAuxPaperSize As Long
    Dim iAuxColorMode As Long
    Dim iAuxPaperBin As Long
    Dim iAuxPrintQuality As Long
    
    mDocKey = DocKey
    mCanceled = False
    If Not PrinterExCurrentDocument Is Nothing Then
        PrinterExCurrentDocument.Printed = False
    End If
    mPrinted = False
    RaiseEvent BeforeShowingPrinterDialog(iCancelPrint)
    If Not iCancelPrint Then
        mDlg.Min = 1
        mDlg.Max = 10000
        
        iDeviceName = mDlg.DeviceName
        iDevModePtr = mDlg.DevModePtr
        iPaperSize = mDlg.PaperSize
        iOrientation = mDlg.Orientation
        iPaperBin = mDlg.PaperBin
        iPrintQuality = mDlg.PrintQuality
        iColorMode = mDlg.ColorMode
        iDuplex = mDlg.Duplex
        
        If mPrintWithDefaultSettings Then
            iPrinterNotSet = mDlg Is Nothing
            If Not iPrinterNotSet Then
                iPrinterNotSet = mDlg.DeviceName = ""
            End If
            If iPrinterNotSet Then
                
                iAuxOrientation = mDlg.Orientation
                iAuxPaperSize = mDlg.PaperSize
                iAuxColorMode = mDlg.ColorMode
                iAuxPaperBin = mDlg.PaperBin
                iAuxPrintQuality = mDlg.PrintQuality
                
                mDlg.ShowPrinter cdePDReturnDefault
                
                mDlg.Orientation = iAuxOrientation
                mDlg.PaperSize = iAuxPaperSize
                mDlg.ColorMode = iAuxColorMode
                mDlg.PaperBin = iAuxPaperBin
                mDlg.PrintQuality = iAuxPrintQuality
                
            End If
            mCanceled = False
        Else
            mDlg.ShowPrinter cdlPDNoSelection
            mCanceled = mDlg.Canceled
        End If
        If Not mCanceled Then
            If mDlg.Orientation <> iOrientation Then
                ResetPrinter2
                Printer2.Orientation = mDlg.Orientation
                RaiseEvent OrientationChange(mDlg.Orientation)
            End If
            RaiseEvent AfterShowingPrinterDialog
            If PrinterExCurrentDocument Is Nothing Then
                RaiseEventPrepareDoc (iCancelPrint)
                If iCancelPrint Then Exit Sub
            Else
                iChanged = False
                
                If mDlg.DeviceName <> iDeviceName Then iChanged = True
                If mDlg.PaperSize <> iPaperSize Then iChanged = True
                If mDlg.DevModePtr <> iDevModePtr Then iChanged = True
                If mDlg.Orientation <> iOrientation Then iChanged = True
                If mDlg.PaperBin <> iPaperBin Then iChanged = True
                If mDlg.PrintQuality <> iPrintQuality Then iChanged = True
                If mDlg.ColorMode <> iColorMode Then iChanged = True
                If mDlg.Duplex <> iDuplex Then iChanged = True
                
                If iChanged Or PrinterExCurrentDocument.PageCount = 0 Then
                    ResetPrinter2
                    RaiseEventPrepareDoc (iCancelPrint)
                    If iCancelPrint Then Exit Sub
                End If
            End If
            If Not PrinterExCurrentDocument Is Nothing Then
                If PrinterExCurrentDocument.PageCount = 0 Then
                    PrinterExCurrentDocument.KillDoc
                    ResetPrinter2
                Else
                    If (mDlg.Flags And cdlPDPageNums) <> 0 Then
                        PrinterExCurrentDocument.FromPage = mDlg.FromPage
                        PrinterExCurrentDocument.ToPage = mDlg.ToPage
                    Else
                        PrinterExCurrentDocument.FromPage = 0
                        PrinterExCurrentDocument.ToPage = 0
                    End If
                    PrinterExCurrentDocument.DocKey = mDocKey
                    PrinterExCurrentDocument.PrintDocumentToPrinterDC mDlg.hDC
                    RaiseEvent DocPrinted(mDocKey)
                    PrinterExCurrentDocument.DocKey = ""
                End If
            End If
            If Not mfrmPrintPreview Is Nothing Then
                Unload mfrmPrintPreview
            End If
            mPrinted = True
        End If
    End If
    
End Sub

Public Sub ShowPrintPreview(Optional DocKey As String)
    Dim c As Long
    
    mPrinted = False
    If Printers.Count = 0 Then
        mDlg.ShowPageSetup cdePSReturnDefault  ' Windows will show the message and allow to add a printer
        If Printers.Count > 0 Then
            ResetPrinter2
            If gPrinterExDoNotMakeNewPrinterExObject Then
                Set PrinterExCurrentDocument.Font = CloneFont(VB.Printer.Font)
                PrinterExCurrentDocument.DeviceName = VB.Printer.DeviceName
            End If
        End If
    End If
    If Printers.Count > 0 Then
        mDocKey = DocKey
        Set mfrmPrintPreview = New frmPrintPreview
        Set mfrmPrintPreview.PrePrintCtrlObj = Me
        mfrmPrintPreview.MinScalePercent = mMinScalePercent
        mfrmPrintPreview.MaxScalePercent = mMaxScalePercent
        mfrmPrintPreview.FormatButtonVisible = mFormatButtonVisible
        mfrmPrintPreview.UseAltScaleIcons = mPrintPrevUseAltScaleIcons
        mfrmPrintPreview.ToolBarIconsSize = mPrintPrevToolBarIconsSize
        If mFormatButtonVisible Then
            If mFormatButtonToolTipText <> "" Then
                mfrmPrintPreview.FormatButtonToolTipText = mFormatButtonToolTipText
            End If
            For c = 0 To 4
                If Not mFormatButtonPicture(c) Is Nothing Then
                    Set mfrmPrintPreview.FormatButtonPicture(c) = mFormatButtonPicture(c)
                End If
            Next c
        End If
        mfrmPrintPreview.PageSetupButtonVisible = mPageSetupButtonVisible
        mfrmPrintPreview.AllowUserChangeScale = mAllowUserChangeScale And Not (mMinScalePercent = mMaxScalePercent)
        
        mfrmPrintPreview.ScalePercent = mScalePercent
        
        ShowModal mfrmPrintPreview
        Set mfrmPrintPreview.PrePrintCtrlObj = Nothing
        Set mfrmPrintPreview = Nothing
        mPrinted = PrinterEx.Printed
    End If
End Sub

Public Property Get Canceled() As Boolean
    Canceled = mCanceled
End Property

'Public Property Let Canceled(ByVal nValue As Boolean)
'    If mCanceled <> nValue Then
'        mCanceled = nValue
'    End If
'End Property

Private Sub Class_Initialize()
    LoadPredefinedPageNumbersFormatStrings
    mHandleMargins = True
    mPrintPageNumbers = True
    Set mPageNumbersFont = New StdFont
    mPageNumbersFont.Name = "Arial"
    mPageNumbersFont.Bold = True
    mPageNumbersFont.Size = 12
    mPageNumbersForeColor = vbWindowText
    mPageNumbersPosition = vxPositionBottomRight
    mPageNumbersFormat = "Default"
    mFormatButtonVisible = False
    mPageSetupButtonVisible = True
    mFormatButtonToolTipText = ""
    mAllowUserChangeScale = True
    mMinScalePercent = cPrintPreviewDefaultMinScale
    mMaxScalePercent = cPrintPreviewDefaultMaxScale
    mScalePercent = 100
    mPrintPrevToolBarIconsSize = vxPPTIconsAuto
    mPrintPrevUseAltScaleIcons = False
End Sub

Private Sub Class_Terminate()
    If mDocPrepared Then ResetPrinter2
End Sub

Private Sub mfrmPrintPreview_PrepareDoc(Cancel As Boolean)
    ResetPrinter2
    RaiseEventPrepareDoc Cancel
End Sub

Public Sub RefreshPreview()
    If Not mfrmPrintPreview Is Nothing Then
        mfrmPrintPreview.RefreshPreview
    End If
End Sub

Private Sub mfrmPrintPreview_ScaleChange(NewScalePercent As Integer)
    mScalePercent = NewScalePercent
    RaiseEvent ScaleChange(NewScalePercent)
End Sub

Private Sub mfrmPrintPreview_FormatOptionsClick(ByRef Canceled As Boolean)
    RaiseEvent FormatOptionsClick(Canceled)
End Sub


Public Property Let PaperSize(ByVal nValue As cdePaperSizeConstants)
    If mDlg.PaperSize <> nValue Then
        mDlg.PaperSize = nValue
        If Not PrinterExCurrentDocument Is Nothing Then
            PrinterExCurrentDocument.DontCheckError396 = True
            PrinterExCurrentDocument.PaperSize = nValue
            PrinterExCurrentDocument.DontCheckError396 = False
        End If
    End If
End Property

Public Property Get PaperSize() As cdePaperSizeConstants
    PaperSize = mDlg.PaperSize
End Property


Public Property Let PaperBin(ByVal nValue As cdePaperBinConstants)
    If mDlg.PaperBin <> nValue Then
        mDlg.PaperBin = nValue
        If Not PrinterExCurrentDocument Is Nothing Then
            PrinterExCurrentDocument.DontCheckError396 = True
            PrinterExCurrentDocument.PaperBin = nValue
            PrinterExCurrentDocument.DontCheckError396 = False
        End If
    End If
End Property

Public Property Get PaperBin() As cdePaperBinConstants
    PaperBin = mDlg.PaperBin
End Property


Public Property Let PrintQuality(ByVal nValue As cdePrintQualityConstants)
    If mDlg.PrintQuality <> nValue Then
        mDlg.PrintQuality = nValue
        If Not PrinterExCurrentDocument Is Nothing Then
            PrinterExCurrentDocument.DontCheckError396 = True
            PrinterExCurrentDocument.PrintQuality = nValue
            PrinterExCurrentDocument.DontCheckError396 = False
        End If
    End If
End Property

Public Property Get PrintQuality() As cdePrintQualityConstants
    PrintQuality = mDlg.PrintQuality
End Property


Public Property Let ColorMode(ByVal nValue As cdeColorModeConstants)
    If mDlg.ColorMode <> nValue Then
        mDlg.ColorMode = nValue
        If Not PrinterExCurrentDocument Is Nothing Then
            PrinterExCurrentDocument.DontCheckError396 = True
            If nValue = vbPRCMPrinterDefault Then
                PrinterExCurrentDocument.ColorMode = PrinterExCurrentDocument.DefaultColorMode
            Else
                PrinterExCurrentDocument.ColorMode = nValue
            End If
            PrinterExCurrentDocument.DontCheckError396 = False
        End If
    End If
End Property

Public Property Get ColorMode() As cdeColorModeConstants
    ColorMode = mDlg.ColorMode
End Property


Public Property Let Duplex(ByVal nValue As cdeDuplexConstants)
    If mDlg.Duplex <> nValue Then
        mDlg.Duplex = nValue
        If Not PrinterExCurrentDocument Is Nothing Then
            PrinterExCurrentDocument.DontCheckError396 = True
            PrinterExCurrentDocument.Duplex = nValue
            PrinterExCurrentDocument.DontCheckError396 = False
        End If
    End If
End Property

Public Property Get Duplex() As cdeDuplexConstants
    Duplex = mDlg.Duplex
End Property


Public Property Let Orientation(ByVal nValue As cdePageOrientationConstants)
    If mDlg.Orientation <> nValue Then
        mDlg.Orientation = nValue
        If Not PrinterExCurrentDocument Is Nothing Then
            PrinterExCurrentDocument.DontCheckError396 = True
            PrinterExCurrentDocument.Orientation = nValue
            PrinterExCurrentDocument.DontCheckError396 = False
        End If
        If Not mfrmPrintPreview Is Nothing Then
            mfrmPrintPreview.Orientation = nValue
        End If
        RaiseEvent OrientationChange(mDlg.Orientation)
    End If
End Property

Public Property Get Orientation() As cdePageOrientationConstants
    Orientation = mDlg.Orientation
End Property


Public Property Let LeftMargin(ByVal nValue As Single)
    If mDlg.LeftMargin <> nValue Then
        mDlg.LeftMargin = nValue
        If Not PrinterExCurrentDocument Is Nothing Then
            Dim iUnits As Long
            
            iUnits = mDlg.Units
            mDlg.Units = vbMillimeters
            PrinterExCurrentDocument.LeftMargin = nValue
            mDlg.Units = iUnits
        End If
    End If
End Property

Public Property Get LeftMargin() As Single
    LeftMargin = mDlg.LeftMargin
End Property


Public Property Let RightMargin(ByVal nValue As Single)
    If mDlg.RightMargin <> nValue Then
        mDlg.RightMargin = nValue
        If Not PrinterExCurrentDocument Is Nothing Then
            Dim iUnits As Long
            
            iUnits = mDlg.Units
            mDlg.Units = vbMillimeters
            PrinterExCurrentDocument.RightMargin = nValue
            mDlg.Units = iUnits
        End If
    End If
End Property

Public Property Get RightMargin() As Single
    RightMargin = mDlg.RightMargin
End Property


Public Property Let TopMargin(ByVal nValue As Single)
    If mDlg.TopMargin <> nValue Then
        mDlg.TopMargin = nValue
        If Not PrinterExCurrentDocument Is Nothing Then
            Dim iUnits As Long
            
            iUnits = mDlg.Units
            mDlg.Units = vbMillimeters
            PrinterExCurrentDocument.TopMargin = nValue
            mDlg.Units = iUnits
        End If
    End If
End Property

Public Property Get TopMargin() As Single
    TopMargin = mDlg.TopMargin
End Property


Public Property Let BottomMargin(ByVal nValue As Single)
    If mDlg.BottomMargin <> nValue Then
        mDlg.BottomMargin = nValue
        If Not PrinterExCurrentDocument Is Nothing Then
            Dim iUnits As Long
            
            iUnits = mDlg.Units
            mDlg.Units = vbMillimeters
            PrinterExCurrentDocument.BottomMargin = nValue
            mDlg.Units = iUnits
        End If
    End If
End Property

Public Property Get BottomMargin() As Single
    BottomMargin = mDlg.BottomMargin
End Property


Public Property Let MinLeftMargin(ByVal nValue As Single)
    If mDlg.MinLeftMargin <> nValue Then
        mDlg.MinLeftMargin = nValue
    End If
End Property

Public Property Get MinLeftMargin() As Single
    MinLeftMargin = mDlg.MinLeftMargin
End Property


Public Property Let MinRightMargin(ByVal nValue As Single)
    If mDlg.MinRightMargin <> nValue Then
        mDlg.MinRightMargin = nValue
    End If
End Property

Public Property Get MinRightMargin() As Single
    MinRightMargin = mDlg.MinRightMargin
End Property


Public Property Let MinTopMargin(ByVal nValue As Single)
    If mDlg.MinTopMargin <> nValue Then
        mDlg.MinTopMargin = nValue
    End If
End Property

Public Property Get MinTopMargin() As Single
    MinTopMargin = mDlg.MinTopMargin
End Property


Public Property Let MinBottomMargin(ByVal nValue As Single)
    If mDlg.MinBottomMargin <> nValue Then
        mDlg.MinBottomMargin = nValue
    End If
End Property

Public Property Get MinBottomMargin() As Single
    MinBottomMargin = mDlg.MinBottomMargin
End Property


Public Property Let AmbientUserMode(ByVal nValue As Boolean)
Attribute AmbientUserMode.VB_MemberFlags = "40"
    If mDlg.AmbientUserMode <> nValue Then
        mDlg.AmbientUserMode = nValue
    End If
End Property


Public Property Get Units() As cdeUnits
    Units = mDlg.Units
End Property

Public Property Let Units(ByVal nValue As cdeUnits)
    If mDlg.Units <> nValue Then
        mDlg.Units = nValue
    End If
End Property


Public Property Get UnitsForUser() As cdeUnitsForUser
    UnitsForUser = mDlg.UnitsForUser
End Property

Public Property Let UnitsForUser(ByVal nValue As cdeUnitsForUser)
    If mDlg.UnitsForUser <> nValue Then
        mDlg.UnitsForUser = nValue
    End If
End Property


Public Property Get HandleMargins() As Boolean
    HandleMargins = mHandleMargins
End Property

Public Property Let HandleMargins(ByVal nValue As Boolean)
    If mHandleMargins <> nValue Then
        mHandleMargins = nValue
        If Not PrinterExCurrentDocument Is Nothing Then
            PrinterExCurrentDocument.HandleMargins = mHandleMargins
        End If
    End If
End Property


Public Property Get PrintPageNumbers() As Boolean
    PrintPageNumbers = mPrintPageNumbers
End Property

Public Property Let PrintPageNumbers(ByVal nValue As Boolean)
    If mPrintPageNumbers <> nValue Then
        mPrintPageNumbers = nValue
    End If
End Property


Public Property Set PageNumbersFont(ByVal nFont As StdFont)
    Set mPageNumbersFont = nFont
End Property

Public Property Get PageNumbersFont() As StdFont
    Set PageNumbersFont = mPageNumbersFont
End Property


Public Property Get PageNumbersForeColor() As OLE_COLOR
    PageNumbersForeColor = mPageNumbersForeColor
End Property

Public Property Let PageNumbersForeColor(ByVal nValue As OLE_COLOR)
    If mPageNumbersForeColor <> nValue Then
        mPageNumbersForeColor = nValue
    End If
End Property


Public Property Get PageNumbersPosition() As vbExPageNumbersPositionConstants
    PageNumbersPosition = mPageNumbersPosition
End Property

Public Property Let PageNumbersPosition(ByVal nValue As vbExPageNumbersPositionConstants)
    If mPageNumbersPosition <> nValue Then
        mPageNumbersPosition = nValue
    End If
End Property


Public Property Get PageNumbersFormat() As String
    If mPageNumbersFormat = "Default" Then
        PageNumbersFormat = GetLocalizedString(efnGUIStr_General_PageNumbersFormatString_Default)
    Else
        PageNumbersFormat = mPageNumbersFormat
    End If
End Property

Public Property Let PageNumbersFormat(ByVal nValue As String)
    If nValue = GetLocalizedString(efnGUIStr_General_PageNumbersFormatString_Default) Then
        mPageNumbersFormat = "Default"
    Else
        mPageNumbersFormat = nValue
    End If
End Property


Public Property Get DocumentName() As String
    DocumentName = mDocumentName
End Property

Public Property Let DocumentName(ByVal nValue As String)
    If mDocumentName <> nValue Then
        mDocumentName = nValue
        If Not PrinterExCurrentDocument Is Nothing Then
            PrinterExCurrentDocument.DocumentName = mDocumentName
        End If
    End If
End Property


Public Property Get FormatButtonVisible() As Boolean
    FormatButtonVisible = mFormatButtonVisible
End Property

Public Property Let FormatButtonVisible(ByVal nValue As Boolean)
    If mFormatButtonVisible <> nValue Then
        mFormatButtonVisible = nValue
        If Not mfrmPrintPreview Is Nothing Then
            mfrmPrintPreview.FormatButtonVisible = mFormatButtonVisible
        End If
    End If
End Property


Public Property Get FormatButtonToolTipText() As String
    FormatButtonToolTipText = mFormatButtonToolTipText
    If FormatButtonToolTipText = "" Then
        FormatButtonToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Format)
    End If
End Property

Public Property Let FormatButtonToolTipText(ByVal nValue As String)
    If mFormatButtonToolTipText <> nValue Then
        mFormatButtonToolTipText = nValue
        If Not mfrmPrintPreview Is Nothing Then
            mfrmPrintPreview.FormatButtonToolTipText = mFormatButtonToolTipText
        End If
    End If
End Property


Public Property Get FormatButtonPicture(ByVal nSizeIdentifier As VBExToobarDAButtonIconSizeConstants) As StdPicture
    Dim iFrm As frmPrintPreview
    Dim iWritingProps As Boolean
    Static sDefaultFormatButtonPicture(vxIZ36) As Object
    Dim c As Long
    
    iWritingProps = (nSizeIdentifier And 64) <> 0
    nSizeIdentifier = nSizeIdentifier And Not 64
    
    If (nSizeIdentifier < vxIZ16) Or (nSizeIdentifier > vxIZ36) Then
        RaiseError 381, TypeName(Me) ' Invalid property array index
        Exit Property
    End If
    
    If mFormatButtonPicture(nSizeIdentifier) Is Nothing Then
        If Not iWritingProps Then
            If sDefaultFormatButtonPicture(vxIZ16) Is Nothing Then
                Set iFrm = New frmPrintPreview
                For c = vxIZ16 To vxIZ36
                    Set sDefaultFormatButtonPicture(c) = iFrm.FormatButtonPicture(c)
                Next
                Unload iFrm
                Set iFrm = Nothing
            End If
            Set FormatButtonPicture = sDefaultFormatButtonPicture(nSizeIdentifier)
        End If
    Else
        Set FormatButtonPicture = mFormatButtonPicture(nSizeIdentifier)
    End If
End Property

Public Property Set FormatButtonPicture(ByVal nSizeIdentifier As VBExToobarDAButtonIconSizeConstants, nPic As StdPicture)
    If (nSizeIdentifier < vxIZ16) Or (nSizeIdentifier > vxIZ36) Then
        RaiseError 381, TypeName(Me) ' Invalid property array index
        Exit Property
    End If
    Set mFormatButtonPicture(nSizeIdentifier) = nPic
    If Not mfrmPrintPreview Is Nothing Then
        Set mfrmPrintPreview.FormatButtonPicture(nSizeIdentifier) = nPic
    End If
End Property


Public Property Get AllowUserChangeScale() As Boolean
    AllowUserChangeScale = mAllowUserChangeScale
End Property

Public Property Let AllowUserChangeScale(ByVal nValue As Boolean)
    If mAllowUserChangeScale <> nValue Then
        mAllowUserChangeScale = nValue
    End If
End Property


Public Sub ReSetPrinterExSettings()
    SetPrinterExSettings
End Sub

Private Sub SetPrinterExSettings()
    Dim iAuxPrinterEx As Object
    Dim iUnits As Long
    
    iUnits = mDlg.Units
    mDlg.Units = vbMillimeters
    
    If PrinterExCurrentDocument Is Nothing Then
        Set iAuxPrinterEx = Printer2
    End If
    PrinterExCurrentDocument.DisableEvents = True
    'PrinterExCurrentDocument.HdcReference = mDlg.hDC
    PrinterExCurrentDocument.DeviceName = mDlg.DeviceName
    PrinterExCurrentDocument.Orientation = mDlg.Orientation
    PrinterExCurrentDocument.Zoom = mScalePercent
    PrinterExCurrentDocument.PaperSize = mDlg.PaperSize
    PrinterExCurrentDocument.PaperBin = mDlg.PaperBin
    PrinterExCurrentDocument.PrintQuality = mDlg.PrintQuality
    PrinterExCurrentDocument.Duplex = mDlg.Duplex
    PrinterExCurrentDocument.ColorMode = mDlg.ColorMode
    PrinterExCurrentDocument.DevModePtr = mDlg.DevModePtr
    PrinterExCurrentDocument.LeftMargin = mDlg.LeftMargin
    PrinterExCurrentDocument.TopMargin = mDlg.TopMargin
    PrinterExCurrentDocument.RightMargin = mDlg.RightMargin
    PrinterExCurrentDocument.BottomMargin = mDlg.BottomMargin
    PrinterExCurrentDocument.HandleMargins = mHandleMargins
    PrinterExCurrentDocument.PrintPageNumbers = mPrintPageNumbers
    If mPrintPageNumbers Then
        Set PrinterExCurrentDocument.PageNumbersFont = mPageNumbersFont
        PrinterExCurrentDocument.PageNumbersForeColor = mPageNumbersForeColor
        PrinterExCurrentDocument.PageNumbersFormat = mPageNumbersFormat
        PrinterExCurrentDocument.PageNumbersPosition = mPageNumbersPosition
    End If
    PrinterExCurrentDocument.DocumentName = mDocumentName
    PrinterExCurrentDocument.DisableEvents = False
    
    mDlg.Units = iUnits
End Sub

Public Property Get PaperWidth() As Single
    PaperWidth = mDlg.PaperWidth
End Property

Public Property Get PaperHeight() As Single
    PaperHeight = mDlg.PaperHeight
End Property


Public Property Let FromPage(ByVal nValue As Long)
    If nValue <> mDlg.FromPage Then
        mDlg.FromPage = nValue
        If (mDlg.FromPage <> 0) And (mDlg.ToPage <> 0) Then
            mDlg.Flags = mDlg.Flags Or cdlPDPageNums
        End If
    End If
End Property

Public Property Get FromPage() As Long
    FromPage = mDlg.FromPage
End Property


Public Property Let ToPage(ByVal nValue As Long)
    If nValue <> mDlg.ToPage Then
        mDlg.ToPage = nValue
        If (mDlg.FromPage <> 0) And (mDlg.ToPage <> 0) Then
            mDlg.Flags = mDlg.Flags Or cdlPDPageNums
        End If
    End If
End Property

Public Property Get ToPage() As Long
    ToPage = mDlg.ToPage
End Property


Public Property Let Copies(ByVal nValue As Long)
    If nValue <> mDlg.Copies Then
        mDlg.Copies = nValue
    End If
End Property

Public Property Get Copies() As Long
    Copies = mDlg.Copies
End Property


Private Sub RaiseEventPrepareDoc(nCancel As Boolean)
    Dim iCancel As Boolean
    
    gPrinterExFromPrintFn = gPrinterExFromPrintFn + 1
    
    SetPrinterExSettings
    
    Set mPrinterExEvents = PrinterExCurrentDocument.Events
    PrinterExCurrentDocument.DocKey = mDocKey
    
    If Not (mParent Is Nothing) And (mProcedureName <> "") Then
        CallByName mParent, mProcedureName, VbMethod
    Else
        RaiseEvent PrepareDoc(iCancel, mDocKey)
    End If
    
    If iCancel Then
        On Error Resume Next
        nCancel = True
        Set mPrinterExEvents = Nothing
        Exit Sub
    End If
    
    If Not PrinterExCurrentDocument Is Nothing Then
        PrinterExCurrentDocument.RaiseEvent_EndDoc PrinterExCurrentDocument.GetDocFirstPageIndex(PrinterExCurrentDocument.PageCount), PrinterExCurrentDocument.GetDocLastPageIndex(PrinterExCurrentDocument.PageCount)
        PrinterExCurrentDocument.DocKey = ""
    End If
    Set mPrinterExEvents = Nothing
    
    gPrinterExFromPrintFn = gPrinterExFromPrintFn - 1
    mDocPrepared = True
End Sub

Public Property Let ScalePercent(ByVal nValue As Long)
    If nValue <> mScalePercent Then
        mScalePercent = nValue
        If Not mfrmPrintPreview Is Nothing Then
            mfrmPrintPreview.ScalePercent = nValue
        End If
    End If
End Property

Public Property Get ScalePercent() As Long
    ScalePercent = mScalePercent
End Property

Public Property Get PrintWithDefaultSettings() As Boolean
    PrintWithDefaultSettings = mPrintWithDefaultSettings
End Property

Public Property Let PrintWithDefaultSettings(ByVal nValue As Boolean)
    If mPrintWithDefaultSettings <> nValue Then
        mPrintWithDefaultSettings = nValue
    End If
End Property

Public Property Get CommonDialogExObject() As CommonDialogExObject
    Set CommonDialogExObject = mDlg
End Property

Public Property Get DocKey() As String
    DocKey = mDocKey
End Property

Public Property Let PrintPrevToolBarIconsSize(ByVal nValue As vbExPrintPrevToolBarIconsSizeConstants)
    If nValue <> mPrintPrevToolBarIconsSize Then
        mPrintPrevToolBarIconsSize = nValue
    End If
End Property

Public Property Get PrintPrevToolBarIconsSize() As vbExPrintPrevToolBarIconsSizeConstants
    PrintPrevToolBarIconsSize = mPrintPrevToolBarIconsSize
End Property

Public Property Let PrintPrevUseAltScaleIcons(ByVal nValue As Boolean)
    If nValue <> mPrintPrevUseAltScaleIcons Then
        mPrintPrevUseAltScaleIcons = nValue
    End If
End Property

Public Property Get PrintPrevUseAltScaleIcons() As Boolean
    PrintPrevUseAltScaleIcons = mPrintPrevUseAltScaleIcons
End Property


Public Property Let MinScalePercent(ByVal nValue As Long)
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


Public Property Let MaxScalePercent(ByVal nValue As Long)
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


Public Property Set Parent(nForm As Object)
    Set mParent = nForm
End Property

Public Property Get Parent() As Object
    Set Parent = mParent
End Property


Public Property Let ProcedureName(ByVal NameOfPublicSubOnForm As String)
    If mProcedureName <> NameOfPublicSubOnForm Then
        mProcedureName = NameOfPublicSubOnForm
    End If
End Property

Public Property Get ProcedureName() As String
    ProcedureName = mProcedureName
End Property


Public Property Get Printed() As Boolean
    If Not PrinterExCurrentDocument Is Nothing Then
        Printed = mPrinted
    End If
End Property


Friend Property Get Changed() As Boolean
    Changed = mDlg.Changed
End Property


Public Property Let PageSetupButtonVisible(ByVal nValue As Boolean)
    If nValue <> mPageSetupButtonVisible Then
        mPageSetupButtonVisible = nValue
        If Not mfrmPrintPreview Is Nothing Then
            mfrmPrintPreview.PageSetupButtonVisible = mPageSetupButtonVisible
        End If
    End If
End Property

Public Property Get PageSetupButtonVisible() As Boolean
    PageSetupButtonVisible = mPageSetupButtonVisible
End Property


Public Property Let CommonDialogFlags(ByVal nValue As Long)
    If mDlg.Flags <> nValue Then
        mDlg.Flags = nValue
    End If
End Property

Public Property Get CommonDialogFlags() As Long
    CommonDialogFlags = mDlg.Flags
End Property

Private Sub mPrinterExEvents_EndDoc(ByVal FirstPageIndex As Long, ByVal LastPageIndex As Long, ByVal DocKey As String)
    RaiseEvent EndDoc(FirstPageIndex, LastPageIndex, DocKey)
End Sub

Private Sub mPrinterExEvents_StartDoc(ByVal DocKey As String)
    RaiseEvent StartDoc(DocKey)
End Sub

Private Sub mPrinterExEvents_StartPage(ByVal DocKey As String)
    RaiseEvent StartPage(DocKey)
End Sub

Public Property Get DefaultColorMode() As cdeColorModeConstants
    If Not PrinterExCurrentDocument Is Nothing Then
        DefaultColorMode = PrinterExCurrentDocument.DefaultColorMode
    End If
End Property

Private Sub LoadPredefinedPageNumbersFormatStrings()
    Dim c As Long
    Dim iCount As Long
    
    iCount = Val(GetLocalizedString(efnGUIStr_FlexFnObject_PageNumbersFormatStrings_Count))
    ReDim mPredefinedPageNumbersFormatStrings(iCount)
    mPredefinedPageNumbersFormatStrings(0) = GetLocalizedString(efnGUIStr_General_PageNumbersFormatString_Default)
    For c = 1 To iCount
        mPredefinedPageNumbersFormatStrings(c) = GetLocalizedString(efnGUIStr_FlexFnObject_PageNumbersFormatStrings, c)
    Next c
End Sub

Public Function GetPredefinedPageNumbersFormatString(nIndex As Long) As String
    GetPredefinedPageNumbersFormatString = mPredefinedPageNumbersFormatStrings(nIndex)
End Function

Public Property Get GetPredefinedPageNumbersFormatStringsCount() As Long
    GetPredefinedPageNumbersFormatStringsCount = UBound(mPredefinedPageNumbersFormatStrings) + 1
End Property

Public Function GetPageNumbersFormatStringsIndex(nFormatString As String) As Long
    Dim c As Long
    
    GetPageNumbersFormatStringsIndex = -1
    For c = 0 To GetPredefinedPageNumbersFormatStringsCount - 1
        If GetPredefinedPageNumbersFormatString(c) = nFormatString Then
            GetPageNumbersFormatStringsIndex = c
            Exit Function
        End If
    Next c
End Function
