VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#1.7#0"; "vbExtra1.ocx"
Begin VB.Form frmFlexFnTest 
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
   Begin VB.CheckBox chkEnableOrderByColumn 
      Caption         =   "Enable to order the data by Columns by clicking on the grid's columns headers"
      Height          =   228
      Left            =   180
      TabIndex        =   5
      Top             =   6120
      Width           =   5808
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
      Top             =   5472
      Width           =   1488
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4620
      Left            =   144
      TabIndex        =   0
      Top             =   720
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   8149
      _Version        =   393216
      HighLight       =   2
   End
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   396
      Left            =   7236
      TabIndex        =   2
      Top             =   180
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   699
      PrintButtonVisible=   -1  'True
      TextEditionLocked=   0   'False
      PageNumbersFormat=   ""
      PageNumbersFormatIndex=   0
      BeginProperty PageNumbersFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   $"frmFlexFnTest.frx":0000
      Height          =   588
      Left            =   216
      TabIndex        =   4
      Top             =   72
      Width           =   6852
   End
   Begin VB.Label lblBottomNote 
      Caption         =   "Note: Almost all the code in the form's code module is to fill the grid and display the data."
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
      Height          =   228
      Left            =   180
      TabIndex        =   3
      Top             =   5688
      Width           =   7140
   End
End
Attribute VB_Name = "frmFlexFnTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub chkEnableOrderByColumn_Click()
    FlexFn1.EnableOrderByColumns = CBool(chkEnableOrderByColumn.Value)
    If FlexFn1.EnableOrderByColumns Then
        MSFlexGrid1.HighLight = flexHighlightNever
    Else
        MSFlexGrid1.HighLight = flexHighlightWithFocus
        LoadGrid
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub FlexFn1_BeforeAction(Action As String, ByVal GridName As String, ByVal ExtraParam As Variant, Cancel As Boolean)
    FlexFn1.Heading = Me.Caption
End Sub

Private Sub Form_Load()
    SetMinMax Me, 4000, 4000
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
    MSFlexGrid1.Move MSFlexGrid1.Left, MSFlexGrid1.Top, Me.ScaleWidth - MSFlexGrid1.Left * 2, Me.ScaleHeight - MSFlexGrid1.Top - 800
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 600, MSFlexGrid1.Top + MSFlexGrid1.Height + 200
    FlexFn1.Move Me.ScaleWidth - FlexFn1.Width - 1000
    lblBottomNote.Top = cmdClose.Top - 90
    chkEnableOrderByColumn.Top = lblBottomNote.Top + lblBottomNote.Height + 75
End Sub
