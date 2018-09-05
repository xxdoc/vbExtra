VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPrinterEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get HandleMargins() As Boolean
    
End Property

Public Property Let HandleMargins(ByVal nValue As Boolean)
    
End Property


Public Property Let LeftMargin(ByVal nValue As Single)

End Property

Public Property Get LeftMargin() As Single

End Property

Public Property Get PageLeftMargin() As Single

End Property


Public Property Let TopMargin(ByVal nValue As Single)

End Property

Public Property Get TopMargin() As Single

End Property

Public Property Get PageTopMargin() As Single

End Property


Public Property Let RightMargin(ByVal nValue As Single)

End Property

Public Property Get RightMargin() As Single

End Property

Public Property Get PageRightMargin() As Single

End Property


Public Property Let BottomMargin(ByVal nValue As Single)

End Property

Public Property Get BottomMargin() As Single

End Property

Public Property Get PageBottomMargin() As Single

End Property


Public Sub GoToPage(ByVal nPageNumber As Long)

End Sub


Public Property Get PageCount() As Long

End Property


Public Sub PrintEx(Optional Text, Optional ByVal Options As vbExPrintExOptionsConstants, Optional ByVal Left As Variant, Optional ByVal Top As Variant, Optional ByVal Width As Variant, Optional ByVal Height As Variant, Optional ByVal Alignment As vbExPrintExAlignConstants, Optional ByVal WordWrap As Boolean = True, Optional ByVal WordBreak As Boolean = True, Optional ByVal RectMarginsX As Single, Optional ByVal RectMarginsY As Single, Optional ByRef StartsNewPage As Boolean, Optional ByRef TextExceedsWidth As Boolean, Optional ByVal CancelIfTextExceedsWidth As Boolean)

End Sub


Public Function TextHeightEx(ByVal Str As String, Optional ByVal Options As vbExPrintExOptionsConstants = vxLineFeed, Optional ByVal Left As Variant, Optional ByVal Top As Variant, Optional ByVal Width As Variant, Optional ByVal Alignment As vbExPrintExAlignConstants = vxAlignLeftTop, Optional ByVal WordWrap As Boolean = True, Optional ByVal WordBreak As Boolean = True, Optional ByVal CellMarginsX As Single, Optional ByVal CellMarginsY As Single, Optional ByRef StartsNewPage As Boolean) As Single

End Function


Public Function TextWidthEx(ByVal Str As String, Optional ByVal Options As vbExPrintExOptionsConstants = vxLineFeed, Optional ByVal Left As Variant, Optional ByVal Top As Variant, Optional ByVal Height As Variant, Optional ByVal Alignment As vbExPrintExAlignConstants = vxAlignLeftTop, Optional ByVal WordWrap As Boolean = True, Optional ByVal WordBreak As Boolean = True, Optional ByVal CellMarginsX As Single, Optional ByVal CellMarginsY As Single, Optional ByRef StartsNewPage As Boolean) As Single

End Function


Public Property Get FontWidth() As Integer
Attribute FontWidth.VB_MemberFlags = "40"

End Property

Public Property Let FontWidth(ByVal nValue As Integer)
    
End Property


Public Sub ShowPrint(Optional ByVal nFlags As cdeCommonDialogExPrinterFlagsConstants = -1, Optional DocKey As String, Optional ByVal CausesEndDoc As Boolean = True)

End Sub


Public Sub ShowPageSetup(Optional ByVal nFlags As cdeCommonDialogExPageSetupFlagsConstants = -1)

End Sub


Public Property Get Canceled() As Boolean

End Property


Public Function GetPageNumber(ByVal nPageIndex As Long) As Long
    
End Function


Public Function GetDocPageCount(Optional ByVal nPageIndex As Long) As Long

End Function


Public Property Get CurrentPageNumber() As Long
    
End Property


Public Function GetDocFirstPageIndex(Optional ByVal nPageIndex As Long) As Long

End Function


Public Function GetDocLastPageIndex(Optional ByVal nPageIndex As Long) As Long

End Function


Public Property Let DocumentName(ByVal nValue As String)

End Property

Public Property Get DocumentName() As String
    
End Property


Public Property Let MinLeftMargin(ByVal nValue As Single)
    
End Property

Public Property Get MinLeftMargin() As Single
    
End Property


Public Property Let MinTopMargin(ByVal nValue As Single)
    
End Property

Public Property Get MinTopMargin() As Single
    
End Property


Public Property Let MinRightMargin(ByVal nValue As Single)
    
End Property

Public Property Get MinRightMargin() As Single
    
End Property


Public Property Let MinBottomMargin(ByVal nValue As Single)
    
End Property

Public Property Get MinBottomMargin() As Single
    
End Property


Public Property Let Units(ByVal nValue As cdeUnits)
    
End Property

Public Property Get Units() As cdeUnits
    
End Property


Public Property Let UnitsForUser(ByVal nValue As cdeUnitsForUser)
    
End Property

Public Property Get UnitsForUser() As cdeUnitsForUser
    
End Property


Public Property Let PrintPageNumbers(ByVal nValue As Boolean)
    
End Property

Public Property Get PrintPageNumbers() As Boolean
    
End Property


Public Property Set PageNumbersFont(ByVal nFont As StdFont)
    
End Property

Public Property Get PageNumbersFont() As StdFont
    
End Property


Public Property Let PageNumbersForeColor(ByVal nValue As OLE_COLOR)
    
End Property

Public Property Get PageNumbersForeColor() As OLE_COLOR
    
End Property


Public Property Let PageNumbersPosition(ByVal nValue As vbExPageNumbersPositionConstants)
    
End Property

Public Property Get PageNumbersPosition() As vbExPageNumbersPositionConstants
    
End Property


Public Property Let PageNumbersFormat(ByVal nValue As String)
    
End Property

Public Property Get PageNumbersFormat() As String
    
End Property


Public Property Let FromPage(ByVal nValue As Long)

End Property

Public Property Get FromPage() As Long
    
End Property


Public Property Let ToPage(ByVal nValue As Long)

End Property

Public Property Get ToPage() As Long
    
End Property


Public Property Get Printed() As Boolean
    
End Property


Public Property Get Events() As Object
    
End Property


Public Property Let Collate(ByVal nValue As Boolean)
    
End Property

Public Property Get Collate() As Boolean

End Property


Public Sub SetPageFixedElement(Optional ByVal nPicture, Optional ByVal nText, Optional ByVal nPositionTop As Variant)

End Sub

Public Property Get PageNonPrintableAreaLeft() As Single

End Property

Public Property Get PageNonPrintableAreaTop() As Single

End Property

Public Property Get PageNonPrintableAreaRight() As Single

End Property

Public Property Get PageNonPrintableAreaBottom() As Single

End Property

Public Property Get PageWidth() As Single

End Property

Public Property Get PageHeight() As Single

End Property


Public Property Let PrintPrevFormatButtonVisible(nValue As Boolean)

End Property

Public Property Get PrintPrevFormatButtonVisible() As Boolean

End Property


Public Property Get PrintPrevFormatButtonToolTipText() As String
    
End Property

Public Property Let PrintPrevFormatButtonToolTipText(nValue As String)
    
End Property


Public Property Get PrintPrevFormatButtonPicture(nSizeIdentifier As VBExToobarDAButtonIconSizeConstants) As StdPicture
    
End Property

Public Property Set PrintPrevFormatButtonPicture(nSizeIdentifier As VBExToobarDAButtonIconSizeConstants, nPic As StdPicture)
    
End Property


Public Property Let PrintPrevPageSetupButtonVisible(nValue As Boolean)

End Property

Public Property Get PrintPrevPageSetupButtonVisible() As Boolean

End Property


Public Property Let PrintPrevToolBarIconsSize(nValue As vbExPrintPrevToolBarIconsSizeConstants)

End Property

Public Property Get PrintPrevToolBarIconsSize() As vbExPrintPrevToolBarIconsSizeConstants

End Property


Public Property Let PrintPrevUseAltScaleIcons(nValue As Boolean)

End Property

Public Property Get PrintPrevUseAltScaleIcons() As Boolean

End Property



Public Sub ShowPrintPreview(Optional nObject As Object, Optional nProcedureName As String, Optional DocKey As String)
    
End Sub

' Kept for keeping binary compatibility, it will be removed in the next version
Public Property Let CommonDialogFlags(nValue As Long)
    
End Property

Public Property Get CommonDialogFlags() As Long
Attribute CommonDialogFlags.VB_MemberFlags = "40"
    
End Property


Public Property Let PrinterFlags(nValue As cdeCommonDialogExPrinterFlagsConstants)
    
End Property

Public Property Get PrinterFlags() As cdeCommonDialogExPrinterFlagsConstants
    
End Property


Public Property Let PageSetupFlags(nValue As cdeCommonDialogExPageSetupFlagsConstants)
    
End Property

Public Property Get PageSetupFlags() As cdeCommonDialogExPageSetupFlagsConstants
    
End Property


Public Property Let PrintPrevMinScalePercent(nValue As Long)
    
End Property

Public Property Get PrintPrevMinScalePercent() As Long
    
End Property


Public Property Let PrintPrevMaxScalePercent(nValue As Long)
    
End Property

Public Property Get PrintPrevMaxScalePercent() As Long
    
End Property


Public Sub PrintRichTextBox(nRTB As Object)

End Sub

Public Property Get SomethingPrintedOnPage() As Boolean

End Property


Public Property Get DocKey() As String
    
End Property

Public Property Let DocKey(nValue As String)
    
End Property


Public Sub DeleteSavedUserPreferences(Optional nDocKey As String)
    
End Sub
