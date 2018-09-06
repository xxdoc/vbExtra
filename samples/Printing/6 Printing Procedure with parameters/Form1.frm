VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3276
   ClientLeft      =   2268
   ClientTop       =   2016
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
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   660
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   984
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "This is Sample 1 modified. MyPrintingRoutine now has parameters (also called arguments)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1128
      Left            =   2124
      TabIndex        =   1
      Top             =   828
      Width           =   2604
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    ' This project shows how to use this Printer object as a replacement of the original VB's Printer object in an existent project
    ' But if you are programming a new project, then you may want to remove or comment the following two lines and let the object to automatically handle the margins and page numbers
    PrinterEx.PrintPageNumbers = False
    PrinterEx.HandleMargins = False
    PrinterEx.ShowPrintPreview Me, "MyPrintingRoutine", , Array("This is the title", vbRed) ' here are the parameters to send to MyPrintingRoutine
End Sub

Public Sub MyPrintingRoutine(nTitle As String, nTitleColor As Long)
    Printer.FontName = "Arial"
    Printer.FontSize = 16
    Printer.FontUnderline = True
    Printer.ForeColor = nTitleColor
    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(nTitle)) / 2
    Printer.Print nTitle
    Printer.Print
    Printer.ForeColor = vbBlack
    Printer.FontSize = 12
    Printer.FontUnderline = False
    Printer.Print "Some other text bigger..."
    Printer.NewPage
    Printer.Print "a second page"
    Printer.DrawWidth = 30
    Printer.Circle (3000, 3000), 2000, vbRed
    Printer.EndDoc
End Sub
