VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrinterExEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event StartPage(ByVal DocKey As String)
Public Event StartDoc(ByVal DocKey As String)
Public Event EndDoc(ByVal FirstPageIndex As Long, ByVal LastPageIndex As Long, ByVal DocKey As String)

Public Event PrepareDoc(ByVal DocKey As String)
Public Event BeforeShowingPageSetupDialog(ByRef CancelDisplay As Boolean)
Public Event AfterShowingPageSetupDialog()
Public Event BeforeShowingPrinterDialog(ByRef CancelPrint As Boolean)
Public Event AfterShowingPrinterDialog()
Public Event FormatOptionsClick(ByRef Canceled As Boolean)
Public Event ScaleChange(NewScalePercent As Integer)
Public Event OrientationChange(ByVal NewOrientation As Long)
Public Event DocPrinted(ByVal DocKey As String)

Friend Sub RaiseEvent_StartPage(ByVal DocKey As String)
    RaiseEvent StartPage(DocKey)
End Sub

Friend Sub RaiseEvent_StartDoc(ByVal DocKey As String)
    RaiseEvent StartDoc(DocKey)
End Sub

Friend Sub RaiseEvent_EndDoc(ByVal FirstPageIndex As Long, ByVal LastPageIndex As Long, ByVal DocKey As String)
    RaiseEvent EndDoc(FirstPageIndex, LastPageIndex, DocKey)
End Sub


Friend Sub RaiseEvent_PrepareDoc(ByVal DocKey As String)
    RaiseEvent PrepareDoc(DocKey)
End Sub

Friend Sub RaiseEvent_BeforeShowingPageSetupDialog(ByRef CancelDisplay As Boolean)
    RaiseEvent BeforeShowingPageSetupDialog(CancelDisplay)
End Sub

Friend Sub RaiseEvent_AfterShowingPageSetupDialog()
    RaiseEvent AfterShowingPageSetupDialog
End Sub

Friend Sub RaiseEvent_BeforeShowingPrinterDialog(ByRef CancelPrint As Boolean)
    RaiseEvent BeforeShowingPrinterDialog(CancelPrint)
End Sub

Friend Sub RaiseEvent_AfterShowingPrinterDialog()
    RaiseEvent AfterShowingPrinterDialog
End Sub

Friend Sub RaiseEvent_FormatOptionsClick(ByRef Canceled As Boolean)
    RaiseEvent FormatOptionsClick(Canceled)
End Sub

Friend Sub RaiseEvent_ScaleChange(NewScalePercent As Integer)
    RaiseEvent ScaleChange(NewScalePercent)
End Sub

Friend Sub RaiseEvent_OrientationChange(ByVal NewOrientation As Long)
    RaiseEvent OrientationChange(NewOrientation)
End Sub

Friend Sub RaiseEvent_DocPrinted(ByVal DocKey As String)
    RaiseEvent DocPrinted(DocKey)
End Sub


