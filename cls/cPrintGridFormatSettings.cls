VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintGridFormatSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event ScalePercentChange()

Private mScalePercent As Long

Public ColorMode As cdeColorModeConstants
Public GridAlign As Long
Public PrintPageNumbers As Boolean
Public PageNumbersPosition As Long
Public PageNumbersFormatLong As Long
Public PageNumbersFont As StdFont
Public PageNumbersForeColor As Long
Public HeadingFont As StdFont
Public HeadingFontColor As Long
Public SubheadingFont As StdFont
Public SubheadingFontColor As Long
Public OtherTextsFont As StdFont
Public OtherTextsFontColor As Long


Private Sub Class_Initialize()
    ColorMode = vbPRCMPrinterDefault
    GridAlign = 1
    PrintPageNumbers = True
    PageNumbersPosition = 0
    PageNumbersFormatLong = 0
    
    Set PageNumbersFont = New StdFont
    PageNumbersFont.Name = "Arial"
    PageNumbersFont.Bold = True
    PageNumbersFont.Size = 12
    
    PageNumbersForeColor = 0
    
    Set HeadingFont = New StdFont
    HeadingFont.Name = "Arial"
    HeadingFont.Underline = True
    HeadingFont.Size = 14
    
    HeadingFontColor = 0
    
    Set SubheadingFont = New StdFont
    SubheadingFont.Name = "Arial"
    SubheadingFont.Bold = True
    SubheadingFont.Size = 12
    
    SubheadingFontColor = 0
    
    Set OtherTextsFont = New StdFont
    OtherTextsFont.Name = "Arial"
    OtherTextsFont.Size = 11
    
    OtherTextsFontColor = 0
    
End Sub


Public Property Let ScalePercent(nValue As Long)
    Dim iValue As Long
    
    iValue = nValue
    If (iValue < cPrintPreviewMinScale) Then
        iValue = cPrintPreviewMinScale
    End If
    If (iValue > cPrintPreviewMaxScale) Then
        iValue = cPrintPreviewMaxScale
    End If
    If iValue <> mScalePercent Then
        mScalePercent = iValue
        RaiseEvent ScalePercentChange
    End If
End Property

Public Property Get ScalePercent() As Long
    ScalePercent = mScalePercent
End Property

