VERSION 5.00
Begin VB.UserControl PrintFn 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00C0FFFF&
   MaskPicture     =   "ctlPrintFn.ctx":0000
   Picture         =   "ctlPrintFn.ctx":0E12
   PropertyPages   =   "ctlPrintFn.ctx":1C26
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ctlPrintFn.ctx":1CAA
End
Attribute VB_Name = "PrintFn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event PrepareDoc(ByVal DocKey As String)
Public Event BeforeShowingPageSetupDialog(ByRef CancelDisplay As Boolean)
Public Event AfterShowingPageSetupDialog()
Public Event BeforeShowingPrinterDialog(ByRef CancelPrint As Boolean)
'Public Event AfterShowingPrinterDialog()
Public Event FormatOptionsClick(ByRef Canceled As Boolean)
Public Event ScaleChange(NewScalePercent As Integer)
Public Event OrientationChange(ByVal NewOrientation As Long)
Public Event DocPrinted(ByVal DocKey As String)

Public Event StartPage(ByVal DocKey As String)
Public Event StartDoc(ByVal DocKey As String)
Public Event EndDoc(ByVal FirstPageIndex As Long, ByVal LastPageIndex As Long, ByVal DocKey As String)

Private WithEvents mPrintFnObject As PrintFnObject
Attribute mPrintFnObject.VB_VarHelpID = -1

Public Sub ShowPageSetup()
    mPrintFnObject.ShowPageSetup
End Sub

Public Sub ShowPrint(Optional DocKey As String)
    mPrintFnObject.ShowPrint DocKey
End Sub

Public Sub ShowPrintPreview(Optional DocKey As String)
    mPrintFnObject.ShowPrintPreview DocKey
End Sub

Public Property Get Canceled() As Boolean
Attribute Canceled.VB_MemberFlags = "400"
    Canceled = mPrintFnObject.Canceled
End Property

'Public Property Let Canceled(nValue As Boolean)
'    mPrintFnObject.Canceled = nValue
'End Property

Private Sub mPrintFnObject_AfterShowingPageSetupDialog()
    RaiseEvent AfterShowingPageSetupDialog
End Sub

'Private Sub mPrintFnObject_AfterShowingPrinterDialog()
'    RaiseEvent AfterShowingPrinterDialog
'End Sub
'
Private Sub mPrintFnObject_BeforeShowingPageSetupDialog(CancelDisplay As Boolean)
    RaiseEvent BeforeShowingPageSetupDialog(CancelDisplay)
End Sub

Private Sub mPrintFnObject_BeforeShowingPrinterDialog(CancelPrint As Boolean)
    RaiseEvent BeforeShowingPrinterDialog(CancelPrint)
End Sub

Private Sub mPrintFnObject_DocPrinted(ByVal DocKey As String)
    RaiseEvent DocPrinted(DocKey)
End Sub

Private Sub mPrintFnObject_EndDoc(ByVal FirstPageIndex As Long, ByVal LastPageIndex As Long, ByVal DocKey As String)
    RaiseEvent EndDoc(FirstPageIndex, LastPageIndex, DocKey)
End Sub

Private Sub mPrintFnObject_PrepareDoc(Cancel As Boolean, ByVal DocKey As String)
    RaiseEvent PrepareDoc(DocKey)
End Sub

Private Sub mPrintFnObject_ScaleChange(NewScalePercent As Integer)
    RaiseEvent ScaleChange(NewScalePercent)
End Sub

Private Sub mPrintFnObject_FormatOptionsClick(ByRef Canceled As Boolean)
    RaiseEvent FormatOptionsClick(Canceled)
End Sub

Private Sub mPrintFnObject_StartDoc(ByVal DocKey As String)
    RaiseEvent StartDoc(DocKey)
End Sub

Private Sub mPrintFnObject_StartPage(ByVal DocKey As String)
    RaiseEvent StartPage(DocKey)
End Sub

Private Sub mPrintFnObject_OrientationChange(ByVal NewOrientation As Long)
    RaiseEvent OrientationChange(NewOrientation)
End Sub

Private Sub UserControl_Initialize()
    Set mPrintFnObject = New PrintFnObject
End Sub

Private Sub UserControl_InitProperties()
    mPrintFnObject.AmbientUserMode = Ambient.UserMode
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim c As Long
    Dim iFont As StdFont
    Dim iLng As Long
    Dim iBool As Boolean
    
    mPrintFnObject.AmbientUserMode = Ambient.UserMode
    
    mPrintFnObject.HandleMargins = PropBag.ReadProperty("HandleMargins", True)
    mPrintFnObject.PrintPageNumbers = PropBag.ReadProperty("PrintPageNumbers", True)
    mPrintFnObject.ColorMode = PropBag.ReadProperty("ColorMode", vbPRCMPrinterDefault)
    mPrintFnObject.Duplex = PropBag.ReadProperty("Duplex", vbPRDPPrinterDefault)
    mPrintFnObject.DocumentName = PropBag.ReadProperty("DocumentName", "")
    mPrintFnObject.AllowUserChangeScale = PropBag.ReadProperty("AllowUserChangeScale", True)
    mPrintFnObject.AllowUserChangeOrientation = PropBag.ReadProperty("AllowUserChangeOrientation", True)
    mPrintFnObject.AllowUserChangePaper = PropBag.ReadProperty("AllowUserChangePaper", True)
    mPrintFnObject.MinScalePercent = PropBag.ReadProperty("MinScalePercent", cPrintPreviewDefaultMinScale)
    mPrintFnObject.MaxScalePercent = PropBag.ReadProperty("MaxScalePercent", cPrintPreviewDefaultMaxScale)
    mPrintFnObject.FormatButtonVisible = PropBag.ReadProperty("FormatButtonVisible", False)
    iBool = PropBag.ReadProperty("PageNumbersButtonVisible", True)
    If iBool <> mPrintFnObject.PageNumbersButtonVisible Then
        mPrintFnObject.PageNumbersButtonVisible = iBool
    End If
    mPrintFnObject.FormatButtonToolTipText = PropBag.ReadProperty("FormatButtonToolTipText", "")
    For c = 0 To 4
        Set mPrintFnObject.FormatButtonPicture(c) = PropBag.ReadProperty("FormatButtonPicture_" & CStr(c), Nothing)
    Next c
    mPrintFnObject.PageSetupButtonVisible = PropBag.ReadProperty("PageSetupButtonVisible", True)
    mPrintFnObject.Orientation = PropBag.ReadProperty("Orientation", vbPRORPrinterDefault)
    iLng = PropBag.ReadProperty("PageNumbersFormatIndex", -1)
    If iLng > -1 Then
        mPrintFnObject.PageNumbersFormat = mPrintFnObject.GetPredefinedPageNumbersFormatString(iLng)
    Else
        mPrintFnObject.PageNumbersFormat = PropBag.ReadProperty("PageNumbersFormat", "Default")
    End If
    mPrintFnObject.PageNumbersPosition = PropBag.ReadProperty("PageNumbersPosition", vxPositionBottomRight)
    mPrintFnObject.PaperBin = PropBag.ReadProperty("PaperBin", vbPRBNPrinterDefault)
    mPrintFnObject.PaperSize = PropBag.ReadProperty("PaperSize", vbPRPSPrinterDefault)
    mPrintFnObject.PrintQuality = PropBag.ReadProperty("PrintQuality", vbPRPQPrinterDefault)
    mPrintFnObject.ScalePercent = PropBag.ReadProperty("ScalePercent", 100)
    mPrintFnObject.Units = vbMillimeters
    mPrintFnObject.MinLeftMargin = PropBag.ReadProperty("MinLeftMargin", 0)
    mPrintFnObject.MinRightMargin = PropBag.ReadProperty("MinRightMargin", 0)
    mPrintFnObject.MinTopMargin = PropBag.ReadProperty("MinTopMargin", 0)
    mPrintFnObject.MinBottomMargin = PropBag.ReadProperty("MinBottomMargin", 0)
    mPrintFnObject.LeftMargin = PropBag.ReadProperty("LeftMargin", cLeftMarginDefault)
    mPrintFnObject.RightMargin = PropBag.ReadProperty("RightMargin", cRightMarginDefault)
    mPrintFnObject.TopMargin = PropBag.ReadProperty("TopMargin", cTopMarginDefault)
    mPrintFnObject.BottomMargin = PropBag.ReadProperty("BottomMargin", cBottomMarginDefault)
    mPrintFnObject.Units = PropBag.ReadProperty("Units", vbMillimeters)
    mPrintFnObject.UnitsForUser = PropBag.ReadProperty("UnitsForUser", cdeMUUserLocale)
    mPrintFnObject.PrintPrevToolBarIconsSize = PropBag.ReadProperty("PrintPrevToolBarIconsSize", vxPPTIconsAuto)
    mPrintFnObject.PrintPrevUseAltScaleIcons = PropBag.ReadProperty("PrintPrevUseAltScaleIcons", False)
    ProcedureName = PropBag.ReadProperty("ProcedureName", "")
'    mPrintFnObject.FromPage = PropBag.ReadProperty("FromPage", 0)
'    mPrintFnObject.ToPage = PropBag.ReadProperty("ToPage", 0)
    mPrintFnObject.Copies = PropBag.ReadProperty("Copies", 1)
    Set iFont = PropBag.ReadProperty("PageNumbersFont", Nothing)
    If Not iFont Is Nothing Then
        Set mPrintFnObject.PageNumbersFont = iFont
    End If
    mPrintFnObject.PageNumbersForeColor = PropBag.ReadProperty("PageNumbersForeColor", vbWindowText)
    mPrintFnObject.PrinterFlags = PropBag.ReadProperty("PrinterFlags", 0&)
    mPrintFnObject.PageSetupFlags = PropBag.ReadProperty("PageSetupFlags", 0&)
End Sub

Private Sub UserControl_Terminate()
    Set mPrintFnObject = Nothing
End Sub

Private Sub UserControl_Resize()
    Dim iH As Long
    Dim iW As Long
    
    If Not Ambient.UserMode Then
        iH = UserControl.ScaleY(UserControl.ScaleHeight, vbTwips, vbPixels)
        iW = UserControl.ScaleX(UserControl.ScaleWidth, vbTwips, vbPixels)
        
        If (iH <> 34) Or (iW <> 34) Then
            If (iH <> 34) Then
                iH = 34
            End If
            If (iW <> 34) Then
                iW = 34
            End If
            UserControl.Size UserControl.ScaleX(iW, vbPixels, vbTwips), UserControl.ScaleY(iH, vbPixels, vbTwips)
        End If
    End If
End Sub

Public Sub RefreshPreview()
    mPrintFnObject.RefreshPreview
End Sub


Public Property Let PaperSize(nValue As cdePaperSizeConstants)
    If nValue <> mPrintFnObject.PaperSize Then
        mPrintFnObject.PaperSize = nValue
        PropertyChanged "PaperSize"
    End If
End Property

Public Property Get PaperSize() As cdePaperSizeConstants
    PaperSize = mPrintFnObject.PaperSize
End Property


Public Property Let PaperBin(nValue As cdePaperBinConstants)
    If nValue <> mPrintFnObject.PaperBin Then
        mPrintFnObject.PaperBin = nValue
        PropertyChanged "PaperBin"
    End If
End Property

Public Property Get PaperBin() As cdePaperBinConstants
    PaperBin = mPrintFnObject.PaperBin
End Property


Public Property Let PrintQuality(nValue As cdePrintQualityConstants)
    If nValue <> mPrintFnObject.PrintQuality Then
        mPrintFnObject.PrintQuality = nValue
        PropertyChanged "PrintQuality"
    End If
End Property

Public Property Get PrintQuality() As cdePrintQualityConstants
    PrintQuality = mPrintFnObject.PrintQuality
End Property


Public Property Let ColorMode(nValue As cdeColorModeConstants)
    If nValue <> mPrintFnObject.ColorMode Then
        mPrintFnObject.ColorMode = nValue
        PropertyChanged "ColorMode"
    End If
End Property

Public Property Get ColorMode() As cdeColorModeConstants
    ColorMode = mPrintFnObject.ColorMode
End Property


Public Property Let Duplex(nValue As cdeDuplexConstants)
    If nValue <> mPrintFnObject.Duplex Then
        mPrintFnObject.Duplex = nValue
        PropertyChanged "Duplex"
    End If
End Property

Public Property Get Duplex() As cdeDuplexConstants
    Duplex = mPrintFnObject.Duplex
End Property


Public Property Let Orientation(nValue As cdePageOrientationConstants)
    If nValue <> mPrintFnObject.Orientation Then
        mPrintFnObject.Orientation = nValue
        PropertyChanged "Orientation"
    End If
End Property

Public Property Get Orientation() As cdePageOrientationConstants
    Orientation = mPrintFnObject.Orientation
End Property


Public Property Let LeftMargin(nValue As Single)
    If nValue <> mPrintFnObject.LeftMargin Then
        mPrintFnObject.LeftMargin = nValue
        PropertyChanged "LeftMargin"
    End If
End Property

Public Property Get LeftMargin() As Single
    LeftMargin = mPrintFnObject.LeftMargin
End Property


Public Property Let RightMargin(nValue As Single)
    If nValue <> mPrintFnObject.RightMargin Then
        mPrintFnObject.RightMargin = nValue
        PropertyChanged "RightMargin"
    End If
End Property

Public Property Get RightMargin() As Single
    RightMargin = mPrintFnObject.RightMargin
End Property


Public Property Let TopMargin(nValue As Single)
    If nValue <> mPrintFnObject.TopMargin Then
        mPrintFnObject.TopMargin = nValue
        PropertyChanged "TopMargin"
    End If
End Property

Public Property Get TopMargin() As Single
    TopMargin = mPrintFnObject.TopMargin
End Property


Public Property Let BottomMargin(nValue As Single)
    If nValue <> mPrintFnObject.BottomMargin Then
        mPrintFnObject.BottomMargin = nValue
        PropertyChanged "BottomMargin"
    End If
End Property

Public Property Get BottomMargin() As Single
    BottomMargin = mPrintFnObject.BottomMargin
End Property


Public Property Let MinLeftMargin(nValue As Single)
    If nValue <> mPrintFnObject.MinLeftMargin Then
        mPrintFnObject.MinLeftMargin = nValue
        PropertyChanged "MinLeftMargin"
    End If
End Property

Public Property Get MinLeftMargin() As Single
    MinLeftMargin = mPrintFnObject.MinLeftMargin
End Property


Public Property Let MinRightMargin(nValue As Single)
    If nValue <> mPrintFnObject.MinRightMargin Then
        mPrintFnObject.MinRightMargin = nValue
        PropertyChanged "MinRightMargin"
    End If
End Property

Public Property Get MinRightMargin() As Single
    MinRightMargin = mPrintFnObject.MinRightMargin
End Property


Public Property Let MinTopMargin(nValue As Single)
    If nValue <> mPrintFnObject.MinTopMargin Then
        mPrintFnObject.MinTopMargin = nValue
        PropertyChanged "MinTopMargin"
    End If
End Property

Public Property Get MinTopMargin() As Single
    MinTopMargin = mPrintFnObject.MinTopMargin
End Property


Public Property Let MinBottomMargin(nValue As Single)
    If nValue <> mPrintFnObject.MinBottomMargin Then
        mPrintFnObject.MinBottomMargin = nValue
        PropertyChanged "MinBottomMargin"
    End If
End Property

Public Property Get MinBottomMargin() As Single
    MinBottomMargin = mPrintFnObject.MinBottomMargin
End Property


Public Property Get Units() As cdeUnits
    Units = mPrintFnObject.Units
End Property

Public Property Let Units(nValue As cdeUnits)
    If nValue <> mPrintFnObject.Units Then
        mPrintFnObject.Units = nValue
        PropertyChanged "Units"
    End If
End Property


Public Property Get UnitsForUser() As cdeUnitsForUser
    UnitsForUser = mPrintFnObject.UnitsForUser
End Property

Public Property Let UnitsForUser(nValue As cdeUnitsForUser)
    If nValue <> mPrintFnObject.UnitsForUser Then
        mPrintFnObject.UnitsForUser = nValue
        PropertyChanged "UnitsForUser"
    End If
End Property


Public Property Get HandleMargins() As Boolean
    HandleMargins = mPrintFnObject.HandleMargins
End Property

Public Property Let HandleMargins(nValue As Boolean)
    If nValue <> mPrintFnObject.HandleMargins Then
        mPrintFnObject.HandleMargins = nValue
        PropertyChanged "HandleMargins"
    End If
End Property


Public Property Get PrintPageNumbers() As Boolean
    PrintPageNumbers = mPrintFnObject.PrintPageNumbers
End Property

Public Property Let PrintPageNumbers(nValue As Boolean)
    If nValue <> mPrintFnObject.PrintPageNumbers Then
        mPrintFnObject.PrintPageNumbers = nValue
        PropertyChanged "PrintPageNumbers"
    End If
End Property


Public Property Get PageNumbersPosition() As vbExPageNumbersPositionConstants
    PageNumbersPosition = mPrintFnObject.PageNumbersPosition
End Property

Public Property Let PageNumbersPosition(nValue As vbExPageNumbersPositionConstants)
    If nValue <> mPrintFnObject.PageNumbersPosition Then
        mPrintFnObject.PageNumbersPosition = nValue
        PropertyChanged "PageNumbersPosition"
    End If
End Property


Public Property Get PageNumbersFormat() As String
Attribute PageNumbersFormat.VB_Description = "# and N are keywords. #: current page number; N: total number of pages"
    PageNumbersFormat = mPrintFnObject.PageNumbersFormat
End Property

Public Property Let PageNumbersFormat(nValue As String)
    If nValue <> mPrintFnObject.PageNumbersFormat Then
        mPrintFnObject.PageNumbersFormat = nValue
        PropertyChanged "PageNumbersFormat"
    End If
End Property


Public Property Get FormatButtonVisible() As Boolean
    FormatButtonVisible = mPrintFnObject.FormatButtonVisible
End Property

Public Property Let FormatButtonVisible(nValue As Boolean)
    If nValue <> mPrintFnObject.FormatButtonVisible Then
        mPrintFnObject.FormatButtonVisible = nValue
        PropertyChanged "FormatButtonVisible"
    End If
End Property


Public Property Get PageNumbersButtonVisible() As Boolean
    PageNumbersButtonVisible = mPrintFnObject.PageNumbersButtonVisible
End Property

Public Property Let PageNumbersButtonVisible(nValue As Boolean)
    If nValue <> mPrintFnObject.PageNumbersButtonVisible Then
        mPrintFnObject.PageNumbersButtonVisible = nValue
        PropertyChanged "PageNumbersButtonVisible"
    End If
End Property


Public Property Get PageSetupButtonVisible() As Boolean
    PageSetupButtonVisible = mPrintFnObject.PageSetupButtonVisible
End Property

Public Property Let PageSetupButtonVisible(nValue As Boolean)
    If nValue <> mPrintFnObject.PageSetupButtonVisible Then
        mPrintFnObject.PageSetupButtonVisible = nValue
        PropertyChanged "PageSetupButtonVisible"
    End If
End Property


Public Property Get FormatButtonToolTipText() As String
    FormatButtonToolTipText = mPrintFnObject.FormatButtonToolTipText
End Property

Public Property Let FormatButtonToolTipText(nValue As String)
    If nValue <> mPrintFnObject.FormatButtonToolTipText Then
        mPrintFnObject.FormatButtonToolTipText = nValue
        PropertyChanged "FormatButtonToolTipText"
    End If
End Property


Public Property Get FormatButtonPicture(nSizeIdentifier As VBExToobarDAButtonIconSizeConstants) As StdPicture
    Set FormatButtonPicture = mPrintFnObject.FormatButtonPicture(nSizeIdentifier)
End Property

Public Property Set FormatButtonPicture(nSizeIdentifier As VBExToobarDAButtonIconSizeConstants, nPic As StdPicture)
    Set mPrintFnObject.FormatButtonPicture(nSizeIdentifier) = nPic
    PropertyChanged "FormatButtonPicture"
End Property

Public Property Let FormatButtonPicture(nSizeIdentifier As VBExToobarDAButtonIconSizeConstants, nPic As StdPicture)
    Set FormatButtonPicture(nSizeIdentifier) = nPic
End Property


Public Property Get AllowUserChangeScale() As Boolean
    AllowUserChangeScale = mPrintFnObject.AllowUserChangeScale
End Property

Public Property Let AllowUserChangeScale(nValue As Boolean)
    If nValue <> mPrintFnObject.AllowUserChangeScale Then
        mPrintFnObject.AllowUserChangeScale = nValue
        PropertyChanged "AllowUserChangeScale"
    End If
End Property


Public Property Get AllowUserChangeOrientation() As Boolean
    AllowUserChangeOrientation = mPrintFnObject.AllowUserChangeOrientation
End Property

Public Property Let AllowUserChangeOrientation(nValue As Boolean)
    If nValue <> mPrintFnObject.AllowUserChangeOrientation Then
        mPrintFnObject.AllowUserChangeOrientation = nValue
        PropertyChanged "AllowUserChangeOrientation"
    End If
End Property


Public Property Get AllowUserChangePaper() As Boolean
    AllowUserChangePaper = mPrintFnObject.AllowUserChangePaper
End Property

Public Property Let AllowUserChangePaper(nValue As Boolean)
    If nValue <> mPrintFnObject.AllowUserChangePaper Then
        mPrintFnObject.AllowUserChangePaper = nValue
        PropertyChanged "AllowUserChangePaper"
    End If
End Property


Public Property Get MinScalePercent() As Long
    MinScalePercent = mPrintFnObject.MinScalePercent
End Property

Public Property Let MinScalePercent(nValue As Long)
    If nValue <> mPrintFnObject.MinScalePercent Then
        mPrintFnObject.MinScalePercent = nValue
        PropertyChanged "MinScalePercent"
    End If
End Property


Public Property Get MaxScalePercent() As Long
    MaxScalePercent = mPrintFnObject.MaxScalePercent
End Property

Public Property Let MaxScalePercent(nValue As Long)
    If nValue <> mPrintFnObject.MaxScalePercent Then
        mPrintFnObject.MaxScalePercent = nValue
        PropertyChanged "MaxScalePercent"
    End If
End Property


Public Property Get ScalePercent() As Long
    ScalePercent = mPrintFnObject.ScalePercent
End Property

Public Property Let ScalePercent(nValue As Long)
    If nValue <> mPrintFnObject.ScalePercent Then
        mPrintFnObject.ScalePercent = nValue
        PropertyChanged "ScalePercent"
    End If
End Property


Public Sub PrintNow()
    mPrintFnObject.PrintNow
End Sub

Public Property Get CommonDialogExObject() As CommonDialogExObject
    Set CommonDialogExObject = mPrintFnObject.CommonDialogExObject
End Property

Public Property Let DocumentName(nDocName As String)
    If nDocName <> mPrintFnObject.DocumentName Then
        mPrintFnObject.DocumentName = nDocName
        PropertyChanged "DocumentName"
    End If
End Property

Public Property Get DocumentName() As String
    DocumentName = mPrintFnObject.DocumentName
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim c As Long
    Dim iStr As String
    Dim iLng As Long
    Dim iUnits As Long
    
    PropBag.WriteProperty "HandleMargins", mPrintFnObject.HandleMargins, True
    PropBag.WriteProperty "PrintPageNumbers", mPrintFnObject.PrintPageNumbers, True
    PropBag.WriteProperty "ColorMode", mPrintFnObject.ColorMode, vbPRCMPrinterDefault
    PropBag.WriteProperty "DocumentName", mPrintFnObject.DocumentName, ""
    PropBag.WriteProperty "Duplex", mPrintFnObject.Duplex, vbPRDPPrinterDefault
    PropBag.WriteProperty "AllowUserChangeScale", mPrintFnObject.AllowUserChangeScale, True
    PropBag.WriteProperty "AllowUserChangeOrientation", mPrintFnObject.AllowUserChangeOrientation, True
    PropBag.WriteProperty "AllowUserChangePaper", mPrintFnObject.AllowUserChangePaper, True
    PropBag.WriteProperty "MinScalePercent", mPrintFnObject.MinScalePercent, cPrintPreviewDefaultMinScale
    PropBag.WriteProperty "MaxScalePercent", mPrintFnObject.MaxScalePercent, cPrintPreviewDefaultMaxScale
    PropBag.WriteProperty "FormatButtonVisible", mPrintFnObject.FormatButtonVisible, False
    PropBag.WriteProperty "PageNumbersButtonVisible", mPrintFnObject.PageNumbersButtonVisible, True
    iStr = mPrintFnObject.FormatButtonToolTipText
    If iStr = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Format) Then
        iStr = ""
    End If
    PropBag.WriteProperty "FormatButtonToolTipText", iStr, ""
    For c = 0 To 4
        PropBag.WriteProperty "FormatButtonPicture_" & CStr(c), mPrintFnObject.FormatButtonPicture(c Or 64), Nothing
    Next c
    PropBag.WriteProperty "PageSetupButtonVisible", mPrintFnObject.PageSetupButtonVisible, True
    PropBag.WriteProperty "Orientation", mPrintFnObject.Orientation, vbPRORPrinterDefault
    iLng = mPrintFnObject.GetPageNumbersFormatStringsIndex(mPrintFnObject.PageNumbersFormat)
    If iLng > -1 Then
        PropBag.WriteProperty "PageNumbersFormat", "", "Default"
        PropBag.WriteProperty "PageNumbersFormatIndex", iLng, -1
    Else
        PropBag.WriteProperty "PageNumbersFormat", mPrintFnObject.PageNumbersFormat, "Default"
        PropBag.WriteProperty "PageNumbersFormatIndex", -1, -1
    End If
    PropBag.WriteProperty "PageNumbersPosition", mPrintFnObject.PageNumbersPosition, vxPositionBottomRight
    PropBag.WriteProperty "PaperBin", mPrintFnObject.PaperBin, vbPRBNPrinterDefault
    PropBag.WriteProperty "PaperSize", mPrintFnObject.PaperSize, vbPRPSPrinterDefault
    PropBag.WriteProperty "PrintQuality", mPrintFnObject.PrintQuality, vbPRPQPrinterDefault
    PropBag.WriteProperty "ScalePercent", mPrintFnObject.ScalePercent, 100
    PropBag.WriteProperty "Units", mPrintFnObject.Units, vbMillimeters
    PropBag.WriteProperty "UnitsForUser", mPrintFnObject.UnitsForUser, cdeMUUserLocale
    iUnits = mPrintFnObject.Units
    mPrintFnObject.Units = vbMillimeters
    PropBag.WriteProperty "MinLeftMargin", mPrintFnObject.MinLeftMargin, 0
    PropBag.WriteProperty "MinRightMargin", mPrintFnObject.MinRightMargin, 0
    PropBag.WriteProperty "MinTopMargin", mPrintFnObject.MinTopMargin, 0
    PropBag.WriteProperty "MinBottomMargin", mPrintFnObject.MinBottomMargin, 0
    PropBag.WriteProperty "LeftMargin", mPrintFnObject.LeftMargin, cLeftMarginDefault
    PropBag.WriteProperty "RightMargin", mPrintFnObject.RightMargin, cRightMarginDefault
    PropBag.WriteProperty "TopMargin", mPrintFnObject.TopMargin, cTopMarginDefault
    PropBag.WriteProperty "BottomMargin", mPrintFnObject.BottomMargin, cBottomMarginDefault
    mPrintFnObject.Units = iUnits
    PropBag.WriteProperty "PrintPrevToolBarIconsSize", mPrintFnObject.PrintPrevToolBarIconsSize, vxPPTIconsAuto
    PropBag.WriteProperty "PrintPrevUseAltScaleIcons", mPrintFnObject.PrintPrevUseAltScaleIcons, False
    PropBag.WriteProperty "ProcedureName", mPrintFnObject.ProcedureName, ""
'    PropBag.WriteProperty "FromPage", mPrintFnObject.FromPage, 0
'    PropBag.WriteProperty "ToPage", mPrintFnObject.ToPage, 0
    PropBag.WriteProperty "Copies", mPrintFnObject.Copies, 1
    PropBag.WriteProperty "PageNumbersFont", mPrintFnObject.PageNumbersFont, Nothing
    PropBag.WriteProperty "PageNumbersForeColor", mPrintFnObject.PageNumbersForeColor, vbWindowText
    PropBag.WriteProperty "PrinterFlags", mPrintFnObject.PrinterFlags, 0&
    PropBag.WriteProperty "PageSetupFlags", mPrintFnObject.PageSetupFlags, 0&
End Sub

Public Property Let PrintPrevToolBarIconsSize(nValue As vbExPrintPrevToolBarIconsSizeConstants)
    If nValue <> mPrintFnObject.PrintPrevToolBarIconsSize Then
        mPrintFnObject.PrintPrevToolBarIconsSize = nValue
        PropertyChanged "PrintPrevToolBarIconsSize"
    End If
End Property

Public Property Get PrintPrevToolBarIconsSize() As vbExPrintPrevToolBarIconsSizeConstants
    PrintPrevToolBarIconsSize = mPrintFnObject.PrintPrevToolBarIconsSize
End Property

Public Property Let PrintPrevUseAltScaleIcons(nValue As Boolean)
    If nValue <> mPrintFnObject.PrintPrevUseAltScaleIcons Then
        mPrintFnObject.PrintPrevUseAltScaleIcons = nValue
        PropertyChanged "PrintPrevUseAltScaleIcons"
    End If
End Property

Public Property Get PrintPrevUseAltScaleIcons() As Boolean
    PrintPrevUseAltScaleIcons = mPrintFnObject.PrintPrevUseAltScaleIcons
End Property


Public Property Let ProcedureName(nNameOfPublicSubOnForm As String)
    If nNameOfPublicSubOnForm <> mPrintFnObject.ProcedureName Then
        mPrintFnObject.ProcedureName = nNameOfPublicSubOnForm
        PropertyChanged "ProcedureName"
        If Ambient.UserMode Then
            Set mPrintFnObject.Parent = UserControl.Parent
        End If
    End If
End Property

Public Property Get ProcedureName() As String
Attribute ProcedureName.VB_MemberFlags = "200"
    ProcedureName = mPrintFnObject.ProcedureName
End Property


Public Property Let ParametersArray(ByVal nParametersArray As Variant)
    mPrintFnObject.ParametersArray = nParametersArray
End Property

Public Property Get ParametersArray() As Variant
    ParametersArray = mPrintFnObject.ParametersArray
End Property


Public Property Get Printed() As Boolean
    Printed = mPrintFnObject.Printed
End Property


Public Property Let FromPage(nValue As Long)
    If nValue <> mPrintFnObject.FromPage Then
        mPrintFnObject.FromPage = nValue
'        PropertyChanged "FromPage"
    End If
End Property

Public Property Get FromPage() As Long
Attribute FromPage.VB_MemberFlags = "400"
    FromPage = mPrintFnObject.FromPage
End Property


Public Property Let ToPage(nValue As Long)
    If nValue <> mPrintFnObject.ToPage Then
        mPrintFnObject.ToPage = nValue
'        PropertyChanged "ToPage"
    End If
End Property

Public Property Get ToPage() As Long
Attribute ToPage.VB_MemberFlags = "400"
    ToPage = mPrintFnObject.ToPage
End Property


Public Property Let Copies(nValue As Long)
    If nValue <> mPrintFnObject.Copies Then
        mPrintFnObject.Copies = nValue
        PropertyChanged "Copies"
    End If
End Property

Public Property Get Copies() As Long
    Copies = mPrintFnObject.Copies
End Property


Public Property Let PrinterFlags(nValue As cdeCommonDialogExPrinterFlagsConstants)
    If nValue <> mPrintFnObject.PrinterFlags Then
        mPrintFnObject.PrinterFlags = nValue
        PropertyChanged "PrinterFlags"
    End If
End Property

Public Property Get PrinterFlags() As cdeCommonDialogExPrinterFlagsConstants
    PrinterFlags = mPrintFnObject.PrinterFlags
End Property


Public Property Let PageSetupFlags(nValue As cdeCommonDialogExPageSetupFlagsConstants)
    If nValue <> mPrintFnObject.PageSetupFlags Then
        mPrintFnObject.PageSetupFlags = nValue
        PropertyChanged "PageSetupFlags"
    End If
End Property

Public Property Get PageSetupFlags() As cdeCommonDialogExPageSetupFlagsConstants
    PageSetupFlags = mPrintFnObject.PageSetupFlags
End Property


Public Function GetPredefinedPageNumbersFormatString(nIndex As Long) As String
    GetPredefinedPageNumbersFormatString = mPrintFnObject.GetPredefinedPageNumbersFormatString(nIndex)
End Function

Public Property Get GetPredefinedPageNumbersFormatStringsCount() As Long
    GetPredefinedPageNumbersFormatStringsCount = mPrintFnObject.GetPredefinedPageNumbersFormatStringsCount
End Property


Public Property Set PageNumbersFont(ByVal nFont As StdFont)
    If Not nFont Is mPrintFnObject.PageNumbersFont Then
        Set mPrintFnObject.PageNumbersFont = nFont
        PropertyChanged "PageNumbersFont"
    End If
End Property

Public Property Let PageNumbersFont(ByVal nFont As StdFont)
Attribute PageNumbersFont.VB_ProcData.VB_Invoke_PropertyPut = "StandardFont"
    Set PageNumbersFont = nFont
End Property

Public Property Get PageNumbersFont() As StdFont
    Set PageNumbersFont = mPrintFnObject.PageNumbersFont
End Property


Public Property Get PageNumbersForeColor() As OLE_COLOR
    PageNumbersForeColor = mPrintFnObject.PageNumbersForeColor
End Property

Public Property Let PageNumbersForeColor(ByVal nValue As OLE_COLOR)
    If mPrintFnObject.PageNumbersForeColor <> nValue Then
        mPrintFnObject.PageNumbersForeColor = nValue
        PropertyChanged "PageNumbersForeColor"
    End If
End Property


Public Property Get PrintFnObject() As PrintFnObject
    Set PrintFnObject = mPrintFnObject
End Property
