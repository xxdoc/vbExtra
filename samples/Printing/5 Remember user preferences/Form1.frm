VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3276
   ClientLeft      =   2268
   ClientTop       =   2364
   ClientWidth     =   5448
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3276
   ScaleWidth      =   5448
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1884
      Left            =   576
      TabIndex        =   0
      Top             =   612
      Width           =   3612
   End
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
    PrinterEx.DocKey = Me.Name & "_MyReport_01"
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
