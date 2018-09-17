Attribute VB_Name = "mPrinterReplacement"
Option Explicit

Public Property Get Printer() As Printer
    Set Printer = vbExtra.Printer2
End Property

Public Property Set Printer(nPrinter As Printer)
    Set vbExtra.Printer2 = nPrinter
End Property
