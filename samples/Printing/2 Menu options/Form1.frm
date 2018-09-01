VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3276
   ClientLeft      =   2268
   ClientTop       =   2364
   ClientWidth     =   5448
   LinkTopic       =   "Form1"
   ScaleHeight     =   3276
   ScaleWidth      =   5448
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageSetup 
         Caption         =   "Page setup"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuPrintPreview 
         Caption         =   "Print preview"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub MyPrintingRoutine()
    Printer.FontName = "Arial"
    Printer.FontSize = 12
    Printer.FontUnderline = True
    Printer.Print "This is a sample"
    Printer.Print
    Printer.FontSize = 16
    Printer.FontUnderline = False
    Printer.Print "Some other text bigger..."
    Printer.NewPage
    Printer.Print "a second page"
    Printer.DrawWidth = 30
    Printer.Circle (3000, 3000), 2000, vbRed
    Printer.EndDoc
End Sub

Private Sub Form_Load()
    ' This project shows how to use this Printer object as a replacement of the original VB's Printer object in an existent project
    ' But if you are programming a new project, then you can not put the following two lines
'    PrinterEx.PrintPrevPageSetupButtonVisible = False
'    PrinterEx.HandleMargins = False ' in existing projects it is necessary to change this property to False because it defaults to True. Existing projects must be already handling the margins with their code.
End Sub

Private Sub mnuEdit_Click()
    MsgBox "Menu Edit... (it does nothing)"
End Sub

Private Sub mnuOpen_Click()
    MsgBox "Menu Open... (it does nothing)"
End Sub

Private Sub mnuPageSetup_Click()
    PrinterEx.ShowPageSetup
End Sub

Private Sub mnuPrint_Click()
    PrinterEx.ShowPrint
    If Not PrinterEx.Canceled Then
        MyPrintingRoutine
    End If
End Sub

Private Sub mnuPrintPreview_Click()
    PrinterEx.ShowPrintPreview Me, "MyPrintingRoutine"
End Sub
