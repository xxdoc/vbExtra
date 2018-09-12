Attribute VB_Name = "mPrinterEx"
Option Explicit

Private mPrinterExCurrentDocument As cPrinterEx
Public gPrinterExFromPrintFn As Long

' Fixed element (image or text)
Public gPrinterExPageFixedImage As StdPicture
Public gPrinterExPageFixedText As String
Public gPrinterExPageFixedElementPositionTop As Boolean

' Properties to remember
' general
Public gPrinterExDeviceName As String
Public gPrinterExNotTrackDefault As Boolean
Public gPrinterExPrinted As Boolean
Public gPrinterExHandleMargins As Variant
Public gPrinterExPrintPageNumbers As Variant
Public gPrinterExPageNumbersFont As Variant
Public gPrinterExPageNumbersForeColor As Variant
Public gPrinterExPageNumbersPosition As Variant
Public gPrinterExPageNumbersFormat As Variant
Public gPrinterExAllowUserChangeScale As Variant
Public gPrinterExEvents As Variant
' Page setup and printer options
Public gPrinterExLeftMargin As Variant
Public gPrinterExRightMargin As Variant
Public gPrinterExTopMargin As Variant
Public gPrinterExBottomMargin As Variant
Public gPrinterExMinLeftMargin As Variant
Public gPrinterExMinRightMargin As Variant
Public gPrinterExMinTopMargin As Variant
Public gPrinterExMinBottomMargin As Variant
Public gPrinterExUnits As Variant
Public gPrinterExUnitsForUser As Variant
Public gPrinterExDoNotMakeNewPrinterExObject As Boolean
' to remember but reset on setting new printer
Public gPrinterExZoom As Variant
Public gPrinterExPaperSize As Variant
Public gPrinterExPaperBin As Variant
Public gPrinterExPrintQuality As Variant
Public gPrinterExColorMode As Variant
Public gPrinterExCollate As Variant

Public Const cLeftMarginDefault As Single = 20
Public Const cRightMarginDefault As Single = 15
Public Const cTopMarginDefault As Single = 20
Public Const cBottomMarginDefault As Single = 20

Public Property Get Printer2() As Printer
    If mPrinterExCurrentDocument Is Nothing Then
        Set mPrinterExCurrentDocument = New cPrinterEx
    End If
    Set Printer2 = mPrinterExCurrentDocument
End Property

Public Property Let Printer2(nValue As Printer)
    ' the original VB.Printer doesn't remember these things when setting a Printer.
    gPrinterExPaperSize = Empty
    gPrinterExPaperSize = Empty
    gPrinterExPrintQuality = Empty
    gPrinterExColorMode = Empty
    gPrinterExCollate = Empty
    gPrinterExZoom = Empty
        
    If mPrinterExCurrentDocument Is Nothing Then
        Set mPrinterExCurrentDocument = New cPrinterEx
    End If
    mPrinterExCurrentDocument.DeviceName = nValue.DeviceName
    If mPrinterExCurrentDocument.DeviceName = nValue.DeviceName Then
        gPrinterExDeviceName = mPrinterExCurrentDocument.DeviceName
    End If
End Property

Public Property Get PrinterEx() As IPrinterEx
    If mPrinterExCurrentDocument Is Nothing Then
        Set mPrinterExCurrentDocument = New cPrinterEx
    End If
    Set PrinterEx = mPrinterExCurrentDocument
End Property

Public Sub ResetPrinter2()
    If Not mPrinterExCurrentDocument Is Nothing Then
        If mPrinterExCurrentDocument.PageCount > 0 Then
            mPrinterExCurrentDocument.KillDoc
        End If
    End If
    If Not gPrinterExDoNotMakeNewPrinterExObject Then Set mPrinterExCurrentDocument = Nothing
End Sub

Public Property Get PrinterExCurrentDocument() As cPrinterEx
    Set PrinterExCurrentDocument = mPrinterExCurrentDocument
End Property

