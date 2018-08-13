Attribute VB_Name = "mPrePrinter"
Option Explicit

Private mPrePrintDocument As cPrePrintDocument
Public gPageFixedImage As StdPicture
Public gPageFixedImage_PrintAtRight As Boolean
Public gPageFixedText As String

Public Property Get PrePrinter1() As IvbExtra.IPrePrinterObj
    If mPrePrintDocument Is Nothing Then
        Set mPrePrintDocument = New cPrePrintDocument
    End If
    Set PrePrinter1 = mPrePrintDocument
End Property

Public Sub ResetPrePrinter1()
    Set mPrePrintDocument = Nothing
End Sub

Public Property Get PrePrinterCurrentDocument() As cPrePrintDocument
    Set PrePrinterCurrentDocument = mPrePrintDocument
End Property
