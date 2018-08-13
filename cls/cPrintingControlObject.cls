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

Public Event DocumentPrePrinting(ByVal DocScale As Single, Cancel As Boolean, ByVal DocKey As String)
Public Event BeforeShowingPageSetupDialog(ByRef CancelDisplay As Boolean)
Public Event AfterShowingPageSetupDialog()
Public Event BeforeShowingPrinterDialog(ByRef CancelPrint As Boolean)
Public Event AfterShowingPrinterDialog()
Public Event FormatOptionsClick(ByRef Cancelled As Boolean)
Public Event FontSizeChange(NewPercentValue As Integer)
Public Event UserChangedOrientation(NewOrientation As Long)

Private WithEvents mfrmPrintPreview As frmPrintPreview
Attribute mfrmPrintPreview.VB_VarHelpID = -1
Private mDlg As New CommonDialogExObject

Private mCancelled As Boolean

Public Enum vbExPageNumbersPositionConstants
    vxPositionRight = 0&
    vxPositionLeft = 1&
    vxPositionCenter = 2&
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

Private mAutoHandleMarginsOnPrePrinter As Boolean
Private mAutoPrintPageNumbers As Boolean
Private mPageNumbersFont As StdFont
Private mPageNumbersPosition As Long
Private mPageNumbersFormat As String
Private mPageNumbersFontColor As Long
Private mDocumentName As String
Private mFormatButtonVisible As Boolean
Private mAllowUserChangeScale As Boolean
Private mScalePercent As Long
Private mPrintWithDefaultSettings As Boolean
Private mDocKey As String
Private mPrintPrevToolBarIconsSize As vbExPrintPrevToolBarIconsSizeConstants
Private mPrintPrevUseAltScaleIcons As Boolean

Public Sub ShowPageSetup()
    Dim iCancelDisplay  As Boolean
    Dim iOrientationPrev As Long
    
    mCancelled = False
    RaiseEvent BeforeShowingPageSetupDialog(iCancelDisplay)
    If Not iCancelDisplay Then
        iOrientationPrev = mDlg.Orientation
        mDlg.ShowPageSetup
        mCancelled = mDlg.Cancelled
        If Not mCancelled Then
            If iOrientationPrev <> mDlg.Orientation Then
                RaiseEvent UserChangedOrientation(mDlg.Orientation)
            End If
            RaiseEvent AfterShowingPageSetupDialog
        End If
    End If
End Sub

Public Sub PrintNow()
    mPrintWithDefaultSettings = True
    ShowPrint
    mPrintWithDefaultSettings = False
End Sub

Public Sub ShowPrint(Optional DocKey As String)
    Dim iCancelPrint  As Boolean
    Dim iChanged As Boolean

    Dim iDeviceName As String
    Dim iDevModePtr As Long
    Dim iPaperSize As Long
    Dim iOrientation As Long
    Dim iPaperBin As Long
    Dim iPrintQuality As Long
    Dim iColorMode As Long
    Dim iDuplex As Long
    
    mDocKey = DocKey
    mCancelled = False
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
            mCancelled = False
        Else
            mDlg.ShowPrinter
            mCancelled = mDlg.Cancelled
        End If
        If Not mCancelled Then
            If mDlg.Orientation <> iOrientation Then
                RaiseEvent UserChangedOrientation(mDlg.Orientation)
            End If
            RaiseEvent AfterShowingPrinterDialog
            If PrePrinterCurrentDocument Is Nothing Then
                RaiseEventDocumentPrePrinting (iCancelPrint)
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
                
                If iChanged Then
                    ResetPrePrinter1
                    RaiseEventDocumentPrePrinting (iCancelPrint)
                    If iCancelPrint Then Exit Sub
                End If
            End If
            If Not PrePrinterCurrentDocument Is Nothing Then
                If PrePrinterCurrentDocument.PageCount = 0 Then
                    PrePrinterCurrentDocument.KillDoc
                    ResetPrePrinter1
                Else
                    If (mDlg.flags And cdlPDPageNums) <> 0 Then
                        PrePrinterCurrentDocument.FromPage = mDlg.FromPage
                        PrePrinterCurrentDocument.ToPage = mDlg.ToPage
                    Else
                        PrePrinterCurrentDocument.FromPage = 0
                        PrePrinterCurrentDocument.ToPage = 0
                    End If
                    PrePrinterCurrentDocument.PrintDocumentToPrinterDC mDlg.hdc
                End If
            End If
            If Not mfrmPrintPreview Is Nothing Then
                Unload mfrmPrintPreview
            End If
        End If
    End If
End Sub

Public Sub ShowPrintPreview(Optional DocKey As String)
    If Printers.Count = 0 Then
        mDlg.ShowPageSetup cdePSReturnDefault  ' Windows will show the message and allow to add a printer
    End If
    If Printers.Count <> 0 Then
        mDocKey = DocKey
        Set mfrmPrintPreview = New frmPrintPreview
        Set mfrmPrintPreview.PrePrintCtrlObj = Me
        mfrmPrintPreview.FormatButtonVisible = mFormatButtonVisible
        mfrmPrintPreview.AllowUserChangeScale = mAllowUserChangeScale
        mfrmPrintPreview.ScalePercent = mScalePercent
        
        ShowModal mfrmPrintPreview
        Set mfrmPrintPreview.PrePrintCtrlObj = Nothing
        Set mfrmPrintPreview = Nothing
    End If
End Sub

Public Property Get Cancelled() As Boolean
    Cancelled = mCancelled
End Property

Public Property Let Cancelled(nValue As Boolean)
    mCancelled = nValue
End Property

Private Sub Class_Initialize()
    mAutoHandleMarginsOnPrePrinter = True
    mAutoPrintPageNumbers = True
    Set mPageNumbersFont = New StdFont
    mPageNumbersFont.Name = "Arial"
    mPageNumbersFont.Bold = True
    mPageNumbersFont.Size = 12
    mPageNumbersFontColor = 0
    mPageNumbersPosition = vxPositionRight
    mPageNumbersFormat = "Página # de N"
    mFormatButtonVisible = False
    mAllowUserChangeScale = True
    mScalePercent = 100
    mPrintPrevToolBarIconsSize = vxPPTIconsAuto
    mPrintPrevUseAltScaleIcons = False
End Sub

Private Sub Class_Terminate()
    ResetPrePrinter1
End Sub

Private Sub mfrmPrintPreview_DocumentPrePrinting(Cancel As Boolean)
    ResetPrePrinter1
    RaiseEventDocumentPrePrinting Cancel
End Sub

Public Sub RefreshPreview()
    If Not mfrmPrintPreview Is Nothing Then
        mfrmPrintPreview.RefreshPreview
    End If
End Sub

Private Sub mfrmPrintPreview_FontSizeChange(NewPercentValue As Integer)
    mScalePercent = NewPercentValue
    RaiseEvent FontSizeChange(NewPercentValue)
End Sub

Private Sub mfrmPrintPreview_FormatOptionsClick(ByRef Cancelled As Boolean)
    RaiseEvent FormatOptionsClick(Cancelled)
End Sub


Public Property Let PaperSize(nValue As cdePaperSizeConstants)
    mDlg.PaperSize = nValue
    If Not PrePrinterCurrentDocument Is Nothing Then
        PrePrinterCurrentDocument.PaperSize = nValue
    End If
End Property

Public Property Get PaperSize() As cdePaperSizeConstants
    PaperSize = mDlg.PaperSize
End Property


Public Property Let PaperBin(nValue As cdePaperBinConstants)
    mDlg.PaperBin = nValue
    If Not PrePrinterCurrentDocument Is Nothing Then
        PrePrinterCurrentDocument.PaperBin = nValue
    End If
End Property

Public Property Get PaperBin() As cdePaperBinConstants
    PaperBin = mDlg.PaperBin
End Property


Public Property Let PrintQuality(nValue As cdePrintQualityConstants)
    mDlg.PrintQuality = nValue
    If Not PrePrinterCurrentDocument Is Nothing Then
        PrePrinterCurrentDocument.PrintQuality = nValue
    End If
End Property

Public Property Get PrintQuality() As cdePrintQualityConstants
    PrintQuality = mDlg.PrintQuality
End Property


Public Property Let ColorMode(nValue As cdeColorModeConstants)
    mDlg.ColorMode = nValue
    If Not PrePrinterCurrentDocument Is Nothing Then
        PrePrinterCurrentDocument.ColorMode = nValue
    End If
End Property

Public Property Get ColorMode() As cdeColorModeConstants
    ColorMode = mDlg.ColorMode
End Property


Public Property Let Duplex(nValue As cdeDuplexConstants)
    mDlg.Duplex = nValue
    If Not PrePrinterCurrentDocument Is Nothing Then
        PrePrinterCurrentDocument.Duplex = nValue
    End If
End Property

Public Property Get Duplex() As cdeDuplexConstants
    Duplex = mDlg.Duplex
End Property


Public Property Let Orientation(nValue As cdePageOrientationConstants)
    mDlg.Orientation = nValue
    If Not PrePrinterCurrentDocument Is Nothing Then
        PrePrinterCurrentDocument.Orientation = nValue
    End If
    If Not mfrmPrintPreview Is Nothing Then
        mfrmPrintPreview.Orientation = nValue
    End If
End Property

Public Property Get Orientation() As cdePageOrientationConstants
    Orientation = mDlg.Orientation
End Property


Public Property Let LeftMargin(nValue As Single)
    mDlg.LeftMargin = nValue
    If Not PrePrinterCurrentDocument Is Nothing Then
        PrePrinterCurrentDocument.LeftMargin = nValue
    End If
End Property

Public Property Get LeftMargin() As Single
    LeftMargin = mDlg.LeftMargin
End Property


Public Property Let RightMargin(nValue As Single)
    mDlg.RightMargin = nValue
    If Not PrePrinterCurrentDocument Is Nothing Then
        PrePrinterCurrentDocument.RightMargin = nValue
    End If
End Property

Public Property Get RightMargin() As Single
    RightMargin = mDlg.RightMargin
End Property


Public Property Let TopMargin(nValue As Single)
    mDlg.TopMargin = nValue
    If Not PrePrinterCurrentDocument Is Nothing Then
        PrePrinterCurrentDocument.TopMargin = nValue
    End If
End Property

Public Property Get TopMargin() As Single
    TopMargin = mDlg.TopMargin
End Property


Public Property Let BottomMargin(nValue As Single)
    mDlg.BottomMargin = nValue
    If Not PrePrinterCurrentDocument Is Nothing Then
        PrePrinterCurrentDocument.BottomMargin = nValue
    End If
End Property

Public Property Get BottomMargin() As Single
    BottomMargin = mDlg.BottomMargin
End Property


Public Property Let MinLeftMargin(nValue As Single)
    mDlg.MinLeftMargin = nValue
End Property

Public Property Get MinLeftMargin() As Single
    MinLeftMargin = mDlg.MinLeftMargin
End Property


Public Property Let MinRightMargin(nValue As Single)
    mDlg.MinRightMargin = nValue
End Property

Public Property Get MinRightMargin() As Single
    MinRightMargin = mDlg.MinRightMargin
End Property


Public Property Let MinTopMargin(nValue As Single)
    mDlg.MinTopMargin = nValue
End Property

Public Property Get MinTopMargin() As Single
    MinTopMargin = mDlg.MinTopMargin
End Property


Public Property Let MinBottomMargin(nValue As Single)
    mDlg.MinBottomMargin = nValue
End Property

Public Property Get MinBottomMargin() As Single
    MinBottomMargin = mDlg.MinBottomMargin
End Property


Public Property Let AmbientUserMode(nValue As Boolean)
Attribute AmbientUserMode.VB_MemberFlags = "40"
    mDlg.AmbientUserMode = nValue
End Property


Public Property Get Units() As cdeUnits
    Units = mDlg.Units
End Property

Public Property Let Units(nValue As cdeUnits)
    mDlg.Units = nValue
End Property


Public Property Get UnitsForUser() As cdeUnitsForUser
    UnitsForUser = mDlg.UnitsForUser
End Property

Public Property Let UnitsForUser(nValue As cdeUnitsForUser)
    mDlg.UnitsForUser = nValue
End Property


Public Property Get AutoHandleMarginsOnPrePrinter() As Boolean
    AutoHandleMarginsOnPrePrinter = mAutoHandleMarginsOnPrePrinter
End Property

Public Property Let AutoHandleMarginsOnPrePrinter(nValue As Boolean)
    mAutoHandleMarginsOnPrePrinter = nValue
    If Not PrePrinterCurrentDocument Is Nothing Then
        PrePrinterCurrentDocument.AutoHandleMargins = mAutoHandleMarginsOnPrePrinter
    End If
End Property


Public Property Get AutoPrintPageNumbers() As Boolean
    AutoPrintPageNumbers = mAutoPrintPageNumbers
End Property

Public Property Let AutoPrintPageNumbers(nValue As Boolean)
    mAutoPrintPageNumbers = nValue
End Property


Public Property Set PageNumbersFont(nFont As StdFont)
    Set mPageNumbersFont = nFont
End Property

Public Property Get PageNumbersFont() As StdFont
    Set PageNumbersFont = mPageNumbersFont
End Property


Public Property Get PageNumbersFontColor() As Long
    PageNumbersFontColor = mPageNumbersFontColor
End Property

Public Property Let PageNumbersFontColor(nValue As Long)
    mPageNumbersFontColor = nValue
End Property


Public Property Get PageNumbersPosition() As vbExPageNumbersPositionConstants
    PageNumbersPosition = mPageNumbersPosition
End Property

Public Property Let PageNumbersPosition(nValue As vbExPageNumbersPositionConstants)
    mPageNumbersPosition = nValue
End Property


Public Property Get PageNumbersFormat() As String
    PageNumbersFormat = mPageNumbersFormat
End Property

Public Property Let PageNumbersFormat(nValue As String)
    mPageNumbersFormat = nValue
End Property


Public Property Get DocumentName() As String
    DocumentName = mDocumentName
End Property

Public Property Let DocumentName(nValue As String)
    mDocumentName = nValue
    If Not PrePrinterCurrentDocument Is Nothing Then
        PrePrinterCurrentDocument.DocumentName = mDocumentName
    End If
End Property


Public Property Get FormatButtonVisible() As Boolean
    FormatButtonVisible = mFormatButtonVisible
End Property

Public Property Let FormatButtonVisible(nValue As Boolean)
    mFormatButtonVisible = nValue
End Property


Public Property Get AllowUserChangeScale() As Boolean
    AllowUserChangeScale = mAllowUserChangeScale
End Property

Public Property Let AllowUserChangeScale(nValue As Boolean)
    mAllowUserChangeScale = nValue
End Property


Private Sub SetPrePrinterSettings()
    Dim iUnitsMultiplier As Single
    Dim iPrePrinter As Object
    Dim iAuxFont As StdFont
    Dim iAuxMinBottomMarginMillimeters As Single
    Dim iPageNumbersFontHeight As Single
    
    If mDlg.Units = vbInches Then
        iUnitsMultiplier = 25.4
    Else
        iUnitsMultiplier = 1
    End If
    
    If PrePrinterCurrentDocument Is Nothing Then
        Set iPrePrinter = PrePrinter1
    End If
    'PrePrinterCurrentDocument.HdcReference = mDlg.hDC
    PrePrinterCurrentDocument.DeviceName = mDlg.DeviceName
    PrePrinterCurrentDocument.Port = mDlg.Port
    PrePrinterCurrentDocument.UserSelectedOrientation = mDlg.Orientation
    PrePrinterCurrentDocument.UserSelectedPaperSize = mDlg.PaperSize
    PrePrinterCurrentDocument.UserSelectedPaperBin = mDlg.PaperBin
    PrePrinterCurrentDocument.UserSelectedPrintQuality = mDlg.PrintQuality
    PrePrinterCurrentDocument.UserSelectedDuplex = mDlg.Duplex
    PrePrinterCurrentDocument.UserSelectedColorMode = mDlg.ColorMode
    PrePrinterCurrentDocument.DevModePtr = mDlg.DevModePtr
    If mAutoHandleMarginsOnPrePrinter Then
        PrePrinterCurrentDocument.LeftMargin = mDlg.LeftMargin * iUnitsMultiplier
        PrePrinterCurrentDocument.TopMargin = mDlg.TopMargin * iUnitsMultiplier
        PrePrinterCurrentDocument.RightMargin = mDlg.RightMargin * iUnitsMultiplier
        If mAutoPrintPageNumbers Then
            iAuxMinBottomMarginMillimeters = 12
            Set iAuxFont = PrePrinter1.Font
            Set PrePrinter1.Font = mPageNumbersFont
            gPageFixedImage_PrintAtRight = (mPageNumbersPosition = vxPositionLeft)
            iPageNumbersFontHeight = PrePrinter1.ScaleY(PrePrinter1.TextHeight("AAA"), PrePrinter1.ScaleMode, vbMillimeters)
            Set PrePrinter1.Font = iAuxFont
            If (iPageNumbersFontHeight + 6) > iAuxMinBottomMarginMillimeters Then
                iAuxMinBottomMarginMillimeters = iPageNumbersFontHeight + 6
            End If
            If PrePrinterCurrentDocument.ScaleY(mDlg.BottomMargin, mDlg.Units, vbMillimeters) < iAuxMinBottomMarginMillimeters Then
                PrePrinterCurrentDocument.BottomMargin = iAuxMinBottomMarginMillimeters
            Else
                PrePrinterCurrentDocument.BottomMargin = mDlg.BottomMargin * iUnitsMultiplier
            End If
            PrePrinterCurrentDocument.KillDoc
        Else
            PrePrinterCurrentDocument.BottomMargin = mDlg.BottomMargin * iUnitsMultiplier
        End If
    Else
        PrePrinterCurrentDocument.LeftMargin = 0
        PrePrinterCurrentDocument.TopMargin = 0
        PrePrinterCurrentDocument.RightMargin = 0
        PrePrinterCurrentDocument.BottomMargin = 0
    End If
    PrePrinterCurrentDocument.DocumentName = mDocumentName
End Sub

Public Property Get PaperWidth() As Single
    PaperWidth = mDlg.PaperWidth
End Property

Public Property Get PaperHeight() As Single
    PaperHeight = mDlg.PaperHeight
End Property

Public Property Get FromPage() As Long
    FromPage = mDlg.FromPage
End Property

Private Sub RaiseEventDocumentPrePrinting(nCancel As Boolean)
    Dim iCancel As Boolean
    
    SetPrePrinterSettings
    RaiseEvent DocumentPrePrinting(mScalePercent / 100, iCancel, mDocKey)
    If iCancel Then
        On Error Resume Next
        nCancel = True
        Exit Sub
    End If
    If mAutoPrintPageNumbers Or (gPageFixedText <> "") Then
        If PrePrinter1.PageCount > 0 Then
            PrintPageNumbers
        End If
    End If
End Sub

Private Sub PrintPageNumbers()
    Dim c As Long
    Dim iSM As Long
    Dim iAHM As Boolean
    Dim iBottomMarginMillimeters As Single
    Dim iAuxFont As StdFont
    Dim iPageNumbersFontHeight As Single
    Dim iPNF As String
    Dim iLen As String
    Dim iM As Single
    Dim iAuxY As Single
    
    iSM = PrePrinter1.ScaleMode
    iAHM = PrePrinter1.AutoHandleMargins
    PrePrinter1.ScaleMode = vbMillimeters
    PrePrinter1.AutoHandleMargins = False
    
    Set iAuxFont = PrePrinter1.Font
    Set PrePrinter1.Font = mPageNumbersFont
    iPageNumbersFontHeight = PrePrinter1.ScaleY(PrePrinter1.TextHeight("AAA"), PrePrinter1.ScaleMode, vbMillimeters)
    
    PrePrinter1.ForeColor = PageNumbersFontColor
    For c = 1 To PrePrinter1.PageCount
        PrePrinter1.GoToPage c
        iBottomMarginMillimeters = PrePrinterCurrentDocument.BottomMargin
        
        If (iBottomMarginMillimeters - iPageNumbersFontHeight - 5) > 10 Then
            PrePrinter1.CurrentY = PrePrinter1.ScaleHeight - iPageNumbersFontHeight - 10
        Else
            PrePrinter1.CurrentY = PrePrinter1.ScaleHeight - iBottomMarginMillimeters + 5
        End If
        iPNF = GetFormattedPageNumberString(mPageNumbersFormat, c, PrePrinter1.PageCount)
        iLen = PrePrinter1.TextWidth(iPNF)
        
        If mAutoHandleMarginsOnPrePrinter Then
            Select Case mPageNumbersPosition
                Case vxPositionRight
                    iM = PrePrinterCurrentDocument.RightMargin
                    If iM < 5 Then iM = 5
                    PrePrinter1.CurrentX = PrePrinter1.ScaleWidth - iM - iLen - 1
                Case vxPositionLeft
                    iM = PrePrinterCurrentDocument.LeftMargin
                    If iM < 5 Then iM = 5
                    PrePrinter1.CurrentX = iM
                Case vxPositionCenter
                    PrePrinter1.CurrentX = (PrePrinter1.ScaleWidth - iLen) / 2
            End Select
        Else
            Select Case mPageNumbersPosition
                Case vxPositionRight
                    iM = PrePrinterCurrentDocument.RightMargin
                    If iM < 5 Then iM = 5
                    PrePrinter1.CurrentX = PrePrinter1.ScaleWidth - iM - iLen - 1 - 10
                Case vxPositionLeft
                    iM = PrePrinterCurrentDocument.LeftMargin
                    If iM < 5 Then iM = 5
                    PrePrinter1.CurrentX = iM + 10
                Case vxPositionCenter
                    PrePrinter1.CurrentX = (PrePrinter1.ScaleWidth - iLen) / 2
            End Select
        End If
        iAuxY = PrePrinter1.CurrentY
        If mAutoPrintPageNumbers Then
            PrePrinter1.Preprint iPNF
        End If
        If gPageFixedText <> "" Then
            If gPageFixedImage Is Nothing Then
                PrePrinter1.CurrentY = iAuxY
                If mPageNumbersPosition = vxPositionLeft Then
                    iM = PrePrinterCurrentDocument.RightMargin
                    If iM < 5 Then iM = 5
                    iLen = PrePrinter1.TextWidth(gPageFixedText)
                    PrePrinter1.CurrentX = PrePrinter1.ScaleWidth - iM - iLen - 1 - 10
                Else
                    PrePrinter1.CurrentX = PrePrinterCurrentDocument.LeftMargin
                End If
                PrePrinter1.FontBold = False
                PrePrinter1.FontItalic = True
                PrePrinter1.FontSize = 12
                PrePrinter1.Preprint gPageFixedText
                Set PrePrinter1.Font = mPageNumbersFont
            End If
        End If
        
    Next c
    Set PrePrinter1.Font = iAuxFont
    PrePrinter1.ScaleMode = iSM
    PrePrinter1.AutoHandleMargins = iAHM

End Sub

Public Function GetFormattedPageNumberString(nPageNumbersFormat As String, nPageNumber As Long, nPageCount As Long) As String
    GetFormattedPageNumberString = nPageNumbersFormat
    GetFormattedPageNumberString = GetFormattedPageNumberString & " "
    
    GetFormattedPageNumberString = Replace(GetFormattedPageNumberString, "#", CStr(nPageNumber))
    GetFormattedPageNumberString = Replace(GetFormattedPageNumberString, " N ", " " & nPageCount & " ")
    
    If Right$(GetFormattedPageNumberString, 1) = " " Then
        GetFormattedPageNumberString = Left$(GetFormattedPageNumberString, Len(GetFormattedPageNumberString) - 1)
    End If
End Function


Public Property Let ScalePercent(nValue As Long)
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

Private Sub mfrmPrintPreview_OrientationClicked()
    RaiseEvent UserChangedOrientation(Me.Orientation)
End Sub


Public Property Get PrintWithDefaultSettings() As Boolean
    PrintWithDefaultSettings = mPrintWithDefaultSettings
End Property

Public Property Let PrintWithDefaultSettings(nValue As Boolean)
    mPrintWithDefaultSettings = nValue
End Property

Public Property Get CommonDialogExObject() As CommonDialogExObject
    Set CommonDialogExObject = mDlg
End Property

Public Property Get DocKey() As String
    DocKey = mDocKey
End Property

Public Property Let PrintPrevToolBarIconsSize(nValue As vbExPrintPrevToolBarIconsSizeConstants)
    If nValue <> mPrintPrevToolBarIconsSize Then
        mPrintPrevToolBarIconsSize = nValue
    End If
End Property

Public Property Get PrintPrevToolBarIconsSize() As vbExPrintPrevToolBarIconsSizeConstants
    PrintPrevToolBarIconsSize = mPrintPrevToolBarIconsSize
End Property

Public Property Let PrintPrevUseAltScaleIcons(nValue As Boolean)
    If nValue <> mPrintPrevUseAltScaleIcons Then
        mPrintPrevUseAltScaleIcons = nValue
    End If
End Property

Public Property Get PrintPrevUseAltScaleIcons() As Boolean
    PrintPrevUseAltScaleIcons = mPrintPrevUseAltScaleIcons
End Property


