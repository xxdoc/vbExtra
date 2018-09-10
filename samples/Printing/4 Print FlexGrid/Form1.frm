VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.0#0"; "vbExtra2.ocx"
Begin VB.Form Form1 
   Caption         =   "World populaton data from UN"
   ClientHeight    =   6468
   ClientLeft      =   2268
   ClientTop       =   2016
   ClientWidth     =   9804
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
   ScaleHeight     =   6468
   ScaleWidth      =   9804
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   408
      Left            =   504
      TabIndex        =   3
      Top             =   5868
      Width           =   1488
   End
   Begin vbExtra.SizeGrip SizeGrip1 
      Height          =   228
      Left            =   9576
      Top             =   6240
      Width           =   228
      _ExtentX        =   402
      _ExtentY        =   402
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   408
      Left            =   7668
      TabIndex        =   1
      Top             =   5832
      Width           =   1488
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5052
      Left            =   144
      TabIndex        =   0
      Top             =   108
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   8911
      _Version        =   393216
      HighLight       =   2
   End
   Begin VB.Label lblBottomNote 
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   516
      Left            =   180
      TabIndex        =   2
      Top             =   5292
      Width           =   7140
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Note: Almost all the code in the form's code module is to fill the grid and display the data.
' The code for printing the FlexGrid is at the end

Private Function LoadTextFile(nFilePath As String)
    Dim iMP1 As Long
    Dim iFile As Long
    
    If Dir(nFilePath) <> "" Then
        iFile = FreeFile
        Open nFilePath For Input Access Read As #iFile
        If LOF(iFile) > 0 Then
            LoadTextFile = Input(LOF(iFile), iFile)
        End If
        Close #iFile
    End If
End Function

Private Function ChrCount(nText As String, ByVal nCharW As Long) As Long
    Dim iStrs() As String
    
    iStrs = Split(nText, ChrW(nCharW))
    ChrCount = UBound(iStrs)
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    SetMinMax Me, 7500, 4000
    PersistForm Me, Forms
    LoadGrid
End Sub

Private Sub LoadGrid()
    MSFlexGrid1.Redraw = False
    LoadSampleData
    FormatGrid
    AdjustGridColumnsWidths MSFlexGrid1
    MSFlexGrid1.Redraw = True
End Sub

Private Sub AdjustGridColumnsWidths(nGrid As Object)
    Dim iPic As Control
    Dim c As Long
    Dim r As Long
    Dim iWidest As Long
    Dim iLng As Long
    Dim iSpace As Long
    Dim iRedraw As Boolean
    
    On Error Resume Next
    Call nGrid.Parent.Controls.Add("VB.PictureBox", "Aux_Pic_For_Masuring")
    Set iPic = nGrid.Parent.Controls("Aux_Pic_For_Masuring")
    On Error GoTo 0
    If iPic Is Nothing Then Exit Sub
    
    iSpace = Screen.TwipsPerPixelX * 20
    
    iRedraw = nGrid.Redraw
    nGrid.Redraw = False
    For c = 0 To nGrid.Cols - 1
        iWidest = 0
        nGrid.Col = c
        For r = 0 To nGrid.Rows - 1
            nGrid.Row = r
            iPic.Font.Name = nGrid.CellFontName
            iPic.Font.Size = nGrid.CellFontSize
            iPic.Font.Bold = nGrid.CellFontBold
            iPic.Font.Italic = nGrid.CellFontItalic
            iPic.Font.Underline = nGrid.CellFontUnderline
            iPic.Font.Strikethrough = nGrid.CellFontStrikeThrough
            iLng = iPic.TextWidth(nGrid.TextMatrix(r, c)) + iSpace
            If r = 0 Then
                iLng = iLng + 250 ' for the arrows to order columns
            End If
            If iLng > iWidest Then
                iWidest = iLng
            End If
        Next r
        nGrid.ColWidth(c) = iWidest
    Next c
    nGrid.Redraw = iRedraw
    nGrid.Parent.Controls.Remove ("Aux_Pic_For_Masuring")
End Sub
    

Private Sub LoadSampleData()
    Dim iText As String
    Dim iLines() As String
    Dim c As Long
    Dim iCols As Long
    Dim iRedraw As Boolean
    
    iRedraw = MSFlexGrid1.Redraw
    MSFlexGrid1.Redraw = False
    iText = LoadTextFile(App.Path & "\" & "UNdata.txt")
    iText = Replace(iText, """", "")
    iText = Replace(iText, vbCrLf, vbLf)
    iText = Replace(iText, vbLf & vbCr, vbLf)
    iLines = Split(iText, vbLf)
    MSFlexGrid1.Clear
    MSFlexGrid1.Rows = MSFlexGrid1.FixedRows
    For c = 0 To UBound(iLines)
        If Trim$(iLines(c)) <> "" Then
            iCols = ChrCount(iLines(c), AscW(vbTab)) + 1
            If iCols > MSFlexGrid1.Cols Then
                MSFlexGrid1.Cols = iCols
            End If
            MSFlexGrid1.AddItem iLines(c)
        End If
    Next c
    MSFlexGrid1.FixedRows = 0
    MSFlexGrid1.RemoveItem (0)
    MSFlexGrid1.FixedRows = 1
    MSFlexGrid1.Redraw = iRedraw
End Sub

Private Sub FormatGrid()
    Dim c As Long
    
    MSFlexGrid1.Row = 0
    For c = 0 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.Col = c
        MSFlexGrid1.CellFontBold = True
        MSFlexGrid1.FixedAlignment(c) = flexAlignCenterCenter
    Next c

    MSFlexGrid1.Col = 0
    For c = 0 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Row = c
        MSFlexGrid1.CellFontBold = True
    Next c

End Sub

Private Sub Form_Resize()
    MSFlexGrid1.Move MSFlexGrid1.Left, MSFlexGrid1.Top, Me.ScaleWidth - MSFlexGrid1.Left * 2, Me.ScaleHeight - MSFlexGrid1.Top - 1250
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 600, Me.ScaleHeight - cmdClose.Height - 150
    cmdPrint.Top = cmdClose.Top
    lblBottomNote.Top = MSFlexGrid1.Top + MSFlexGrid1.Height + 100
End Sub


' Here starts the code for printing the FlexGrid

Private Sub cmdPrint_Click()
    PrinterEx.DocKey = Me.Name & "_My_Report_2"
    PrinterEx.ShowPrintPreview Me, "MyPrintingRoutine"
End Sub

Public Sub MyPrintingRoutine()
    Printer.FontSize = 14
    Printer.FontName = "Arial"
    Printer.FontUnderline = True
    Printer.Print "Title"
    Printer.Print
    
    PrinterEx.PrintFlexGrid MSFlexGrid1
    
    Printer.Print
    Printer.FontSize = 12
    Printer.FontUnderline = False
    Printer.Print "Final text"
End Sub

