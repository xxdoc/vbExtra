VERSION 5.00
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#1.7#0"; "vbExtra1.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8076
   ClientLeft      =   2268
   ClientTop       =   2016
   ClientWidth     =   8124
   LinkTopic       =   "Form1"
   ScaleHeight     =   8076
   ScaleWidth      =   8124
   Begin vbExtra.SizeGrip SizeGrip1 
      Height          =   228
      Left            =   7896
      Top             =   7848
      Width           =   228
      _ExtentX        =   402
      _ExtentY        =   402
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   408
      Left            =   5796
      TabIndex        =   2
      Top             =   7488
      Width           =   1596
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5124
      Left            =   108
      TabIndex        =   1
      Top             =   108
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   9038
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      FileName        =   "D:\Programas\vbExtra\samples\Printing\3 Print RichTextBox\RTF_Sample.rtf"
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   408
      Left            =   540
      TabIndex        =   0
      Top             =   7488
      Width           =   1596
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    PrinterEx.ShowPrintPreview Me, "MyPrintingRoutine"
End Sub

Public Sub MyPrintingRoutine()
    Printer.FontName = "Arial"
    Printer.FontSize = 12
    Printer.FontUnderline = True
    Printer.Print "This is a sample"
    Printer.Print
    Printer.FontSize = 16
    Printer.FontUnderline = False
    Printer.Print "Some other text bigger..."
    Printer.Print
    PrinterEx.PrintRichTextBox RichTextBox1
    Printer.NewPage
    Printer.Print "a last page"
    Printer.DrawWidth = 30
    Printer.Circle (3000, 3000), 2000, vbRed
    Printer.EndDoc
End Sub

Private Sub Form_Resize()
    RichTextBox1.Move 60, 60, Me.ScaleWidth - 120, Me.ScaleHeight - 850
    cmdPrint.Top = Me.ScaleHeight - 600
    cmdClose.Top = cmdPrint.Top
    cmdClose.Left = Me.ScaleWidth - cmdClose.Width - 600
End Sub
