VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGridPrintingData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long

Public Enum efnGridType
    efnGridTypeMSFlex = 1
    efnGridTypeMSHFlex = 2
    efnGridTypeOtherFlex = 3
End Enum

' intrinsic data
Private mCellAlignment() As Long
Private mCellBackColor() As Long
Private mCellForeColor() As Long
Private mCellFontIndex() As Long
Private mColLargerFontSize() As Single
Private mColLargerFontSizeFontIndex() As Single
Private mColIsVisible() As Boolean
Private mRowIsVisible() As Boolean
Private mVisibleCols As Long
Private mFontsAttributes() As cFontAttributes
Private mGrid As Object
Private mVisibleFixedRows As Long
Private mLastVisibleCol As Long
Private mGridType As efnGridType
Private mCellBackColorIsExplicitelySet() As Boolean
Private mColOriginalWidth() As Long
Private mLastRowWithData As Long

' added data
Private mCellText() As String
Private mPrintCell() As Boolean
Private mCellPrintRightLine() As Boolean
Private mCellWidth() As Long
Private mLastRowMergedInThisCellAtThisCol() As Long
Private mLastColMergedInThisCellAtThisRow() As Long
Private mFirstColMergedInThisCellAtThisRow() As Long
Private mRowWhereCellIsMerged() As Long
Private mTextHeightCell() As Long
Private mCellTop() As Long
Private mColWidth() As Long
Private mColPosLeft() As Long
Private mColPosRight() As Long
Private mMinColWidth() As Long
Private mMinColWidthFontReduced() As Long
Private mColMergeForPrintingBegunAtRow() As Long
Private mRowWidthDataForColMergedForPrinting() As Long
Private mRowHeight() As Long
Private mTextHeightRow() As Long
Private mRowFinalPosYForTop() As Long
Private mColWidthsAreSet As Boolean
Public Canceled As Boolean
Public mfrmSettingGridDataProgressShown As Boolean

Public Sub SetGrid(nGrid As Object, Optional nBackColorFixed As Long = -1)
    Dim c As Long
    Dim R As Long
    Dim iFontChanged As Boolean
    Dim iCurrentFontIndex As Long
    Dim f As Long
    Dim iFontFound As Boolean
    
    Dim iFontName_Prev As String
    Dim iFontSize_Prev As Single
    Dim iFontBold_prev As Boolean
    Dim iFontItalic_Prev As Boolean
    Dim iFontStrikethrough_Prev As Boolean
    Dim iFontUnderline_Prev As Boolean
    Dim iFontWidth_Prev As Long
    Dim iLng As Long
    Dim iFixedRows As Long
    Dim iFixedCols As Long
    Set mGrid = nGrid
    Dim iBackColor As Long
    Dim iForeColorFixed As Long
    Dim iForeColor As Long
    Dim iTx As Long
    Dim iGridBackColorFixed As Long
    Dim iGridCellBackColor As Long
    Dim iAuxMergeRow0 As Boolean
    Dim iBackColorFixed As Long
    Dim iGrid_CellFontStrikeThrough As Boolean
    Dim iLng2 As Long
    
    ReDim mFontsAttributes(0)
    ReDim mCellAlignment(mGrid.Rows - 1, mGrid.Cols - 1)
    ReDim mCellBackColor(mGrid.Rows - 1, mGrid.Cols - 1)
    ReDim mCellBackColorIsExplicitelySet(mGrid.Rows - 1, mGrid.Cols - 1)
    ReDim mCellForeColor(mGrid.Rows - 1, mGrid.Cols - 1)
    ReDim mCellFontIndex(mGrid.Rows - 1, mGrid.Cols - 1)
    ReDim mColLargerFontSize(mGrid.Cols - 1)
    ReDim mColLargerFontSizeFontIndex(mGrid.Cols - 1)
    ReDim mColIsVisible(mGrid.Cols - 1)
    ReDim mColOriginalWidth(mGrid.Cols - 1)
    ReDim mRowIsVisible(mGrid.Rows - 1)
    
    Canceled = False
    iTx = Screen.TwipsPerPixelX
    
    Select Case LCase$(TypeName(mGrid))
        Case "mshflexgrid"
            mGridType = efnGridTypeMSHFlex
        Case "msflexgrid"
            mGridType = efnGridTypeMSFlex
        Case Else
            If InStr(LCase$(TypeName(mGrid)), "flexgrid") > 0 Then
                mGridType = efnGridTypeOtherFlex
            End If
    End Select
    
    iFixedRows = mGrid.FixedRows
    iFixedCols = mGrid.FixedCols
    
    mVisibleFixedRows = 0
    For R = 0 To iFixedRows - 1
        If mGrid.RowHeight(R) <> 0 Then
            mVisibleFixedRows = mVisibleFixedRows + 1
        End If
    Next R
    
    For R = 0 To mGrid.Rows - 1
        If mGrid.RowHeight(R) = -1 Then
            mRowIsVisible(R) = True
        Else
            mRowIsVisible(R) = mGrid.RowHeight(R) > 35
        End If
    Next R
    
    iAuxMergeRow0 = mGrid.MergeRow(0)
    mGrid.MergeRow(0) = False
    For c = 0 To mGrid.Cols - 1
        If mGrid.ColWidth(c) <> 0 Then
            mLastVisibleCol = c
        End If
        
        If mGridType = efnGridTypeMSFlex Then
            mColOriginalWidth(c) = mGrid.ColWidth(c)
        Else
            If mGrid.ColWidth(c) <> -1 Then
                mColOriginalWidth(c) = mGrid.ColWidth(c)
            Else
                If mGrid.Rows > 0 Then
                    mGrid.Row = 0
                    mGrid.col = c
                    mColOriginalWidth(c) = mGrid.CellWidth + iTx
                Else
                    mColOriginalWidth(c) = 960
                End If
            End If
        End If
    Next c
    mGrid.MergeRow(0) = iAuxMergeRow0
    
    mLastRowWithData = -1
    For R = mGrid.Rows - 1 To mGrid.FixedRows Step -1
        For c = 0 To mGrid.Cols - 1
            If mGrid.TextMatrix(R, c) <> "" Then
                mLastRowWithData = R
                Exit For
            End If
        Next c
        If mLastRowWithData <> -1 Then Exit For
    Next R
    
    iBackColorFixed = mGrid.BackColorFixed
    
    iGridBackColorFixed = mGrid.BackColorFixed
    TranslateColor mGrid.BackColor, 0, iBackColor
    TranslateColor mGrid.ForeColor, 0, iForeColor
    TranslateColor iForeColorFixed, 0, iForeColorFixed
    
    If mGrid.Rows > 3000 Then
        frmSettingGridDataProgress.Show
        frmSettingGridDataProgress.pgb1.Max = mGrid.Rows * mGrid.Cols / 1000 * 3
        mfrmSettingGridDataProgressShown = True
        DoEvents
    End If
    
    For c = 0 To mGrid.Cols - 1
        mGrid.col = c
        mColLargerFontSize(c) = iFontSize_Prev
        mColLargerFontSizeFontIndex(c) = iCurrentFontIndex
        mColIsVisible(c) = mColOriginalWidth(c) > 35
        If mColIsVisible(c) Then
            mVisibleCols = mVisibleCols + 1
        End If
        
        For R = 0 To mGrid.Rows - 1
            If mfrmSettingGridDataProgressShown Then
                If R Mod 1000 = 0 Then
                    frmSettingGridDataProgress.pgb1.Value = (R + c * mGrid.Rows) / 1000
                    frmSettingGridDataProgress.ZOrder
                    DoEvents
                    If frmSettingGridDataProgress.Canceled Then
                        Unload frmSettingGridDataProgress
                        Set frmSettingGridDataProgress = Nothing
                        Canceled = True
                        Exit Sub
                    End If
                End If
            End If
            mGrid.Row = R
            If mGridType = efnGridTypeMSFlex Then
                ' CellAlignment
                iLng = mGrid.CellAlignment
                If iLng = 0 Then
                    If R < iFixedRows Then
                        mCellAlignment(R, c) = mGrid.FixedAlignment(c)
                    Else
                        mCellAlignment(R, c) = mGrid.ColAlignment(c)
                    End If
                Else
                    mCellAlignment(R, c) = iLng
                End If
                
                'CellBackColor
                iLng = mGrid.CellBackColor
                If iLng = iBackColorFixed Then
                    If nBackColorFixed <> -1 Then
                        iLng = nBackColorFixed
                    End If
                End If
                TranslateColor iLng, 0, iLng
                If iLng = 0 Then
                    If Not ((R < iFixedRows) Or (c < iFixedCols)) Then
                        mCellBackColor(R, c) = iBackColor
                    End If
                Else
                    mCellBackColor(R, c) = iLng
                    mCellBackColorIsExplicitelySet(R, c) = True
                End If
                
                TranslateColor mGrid.CellForeColor, 0, iLng
                If iLng = 0 Then
                    If (R < iFixedRows) Or (c < iFixedCols) Then
                        mCellForeColor(R, c) = iForeColorFixed
                    Else
                        mCellForeColor(R, c) = iForeColor
                    End If
                Else
                    mCellForeColor(R, c) = iLng
                End If
                
            Else
                mCellAlignment(R, c) = mGrid.CellAlignment
                If (R < iFixedRows) Or (c < iFixedCols) Then
                    iGridCellBackColor = mGrid.CellBackColor
                    If iGridCellBackColor = 0 Then
                        TranslateColor mGrid.CellForeColor, 0, iLng2
                        If iLng2 = 0 Then
                            iGridCellBackColor = vbWhite
                        End If
                    End If
                    If iGridCellBackColor <> iGridBackColorFixed Then
                        TranslateColor mGrid.CellBackColor, 0, iLng
                        mCellBackColor(R, c) = iLng
                        mCellBackColorIsExplicitelySet(R, c) = True
                    End If
                Else
                    iLng = mGrid.CellBackColor
                    If iLng = 0 Then
                        TranslateColor mGrid.CellForeColor, 0, iLng2
                        If iLng2 = 0 Then
                            iLng = vbWhite
                        End If
                    End If
                    If iLng = iBackColorFixed Then
                        If nBackColorFixed <> -1 Then
                            iLng = nBackColorFixed
                        End If
                    End If
                    TranslateColor iLng, 0, iLng
                    mCellBackColor(R, c) = iLng
                End If
            
                If (R < iFixedRows) Or (c < iFixedCols) Then
                    TranslateColor mGrid.CellForeColor, 0, iLng
                    mCellForeColor(R, c) = iLng
                Else
                    TranslateColor mGrid.CellForeColor, 0, iLng
                    mCellForeColor(R, c) = iLng
                End If
            
            End If
            
            On Error Resume Next
            Err.Clear
            iGrid_CellFontStrikeThrough = mGrid.CellFontStrikeThrough
            If Err.Number Then
                iGrid_CellFontStrikeThrough = iFontStrikethrough_Prev
            End If
            On Error GoTo 0
            
            iFontChanged = False
            If mGrid.CellFontName <> iFontName_Prev Then
                iFontChanged = True
            ElseIf mGrid.CellFontSize <> iFontSize_Prev Then
                iFontChanged = True
            ElseIf mGrid.CellFontBold <> iFontBold_prev Then
                iFontChanged = True
            ElseIf mGrid.CellFontItalic <> iFontItalic_Prev Then
                iFontChanged = True
            ElseIf mGrid.CellFontUnderline <> iFontUnderline_Prev Then
                iFontChanged = True
            ElseIf iGrid_CellFontStrikeThrough <> iFontStrikethrough_Prev Then
                iFontChanged = True
            Else
                Err.Clear
                On Error Resume Next
                If mGrid.CellFontWidth <> iFontWidth_Prev Then
                    If Err.Number = 0 Then
                        iFontChanged = True
                    End If
                End If
                On Error GoTo 0
            End If
            
            If iFontChanged Then
                iFontName_Prev = mGrid.CellFontName
                iFontSize_Prev = mGrid.CellFontSize
                iFontBold_prev = mGrid.CellFontBold
                iFontItalic_Prev = mGrid.CellFontItalic
                iFontUnderline_Prev = mGrid.CellFontUnderline
                On Error Resume Next
                iFontStrikethrough_Prev = mGrid.CellFontStrikeThrough
                iFontWidth_Prev = mGrid.CellFontWidth
                On Error GoTo 0
                
                iFontFound = False
                For f = 1 To UBound(mFontsAttributes)
                    If mFontsAttributes(f).Name = iFontName_Prev Then
                        If mFontsAttributes(f).Size = iFontSize_Prev Then
                            If mFontsAttributes(f).Bold = iFontBold_prev Then
                                If mFontsAttributes(f).Italic = iFontItalic_Prev Then
                                    If mFontsAttributes(f).Underline = iFontUnderline_Prev Then
                                        If mFontsAttributes(f).Strikethrough = iFontStrikethrough_Prev Then
                                            If mFontsAttributes(f).Width = iFontWidth_Prev Then
                                                iFontFound = True
                                                iCurrentFontIndex = f
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next f
                
                If Not iFontFound Then
                    iCurrentFontIndex = UBound(mFontsAttributes) + 1
                    ReDim Preserve mFontsAttributes(iCurrentFontIndex)
                    Set mFontsAttributes(iCurrentFontIndex) = New cFontAttributes
                    
                    mFontsAttributes(iCurrentFontIndex).Name = iFontName_Prev
                    mFontsAttributes(iCurrentFontIndex).Size = iFontSize_Prev
                    mFontsAttributes(iCurrentFontIndex).Bold = iFontBold_prev
                    mFontsAttributes(iCurrentFontIndex).Italic = iFontItalic_Prev
                    mFontsAttributes(iCurrentFontIndex).Underline = iFontUnderline_Prev
                    mFontsAttributes(iCurrentFontIndex).Strikethrough = iFontStrikethrough_Prev
                    mFontsAttributes(iCurrentFontIndex).Width = iFontWidth_Prev
                End If
                
                If iFontSize_Prev > mColLargerFontSize(c) Then
                    mColLargerFontSize(c) = iFontSize_Prev
                    mColLargerFontSizeFontIndex(c) = iCurrentFontIndex
                End If
            
            End If
            
            mCellFontIndex(R, c) = iCurrentFontIndex
            
        Next R
    Next c
    
    ResetAddedData nBackColorFixed
End Sub


Public Property Get CellFontAttibutes(Row As Long, col As Long) As cFontAttributes
    Set CellFontAttibutes = mFontsAttributes(mCellFontIndex(Row, col))
End Property

Public Property Get CellFontAttibutesByFontAttibutesIndex(FontAttibutesIndex As Long) As cFontAttributes
    Set CellFontAttibutesByFontAttibutesIndex = mFontsAttributes(FontAttibutesIndex)
End Property

Public Property Get CellFontAttibutesIndex(Row As Long, col As Long) As Long
    CellFontAttibutesIndex = mCellFontIndex(Row, col)
End Property

Public Property Get RowHeight(Row As Long) As Long
    RowHeight = mRowHeight(Row)
End Property

Public Property Get ColPosLeft(col As Long) As Long
    ColPosLeft = mColPosLeft(col)
End Property

Public Property Get ColPosRight(col As Long) As Long
    ColPosRight = mColPosRight(col)
End Property

Public Property Get CellAlignment(Row As Long, col As Long) As Long
    CellAlignment = mCellAlignment(Row, col)
End Property

Public Property Get CellBackColor(Row As Long, col As Long) As Long
    CellBackColor = mCellBackColor(Row, col)
End Property

Public Property Get CellForeColor(Row As Long, col As Long) As Long
    CellForeColor = mCellForeColor(Row, col)
End Property

Public Property Get GridWidth() As Long
    Dim c As Long
    
    For c = 0 To mGrid.Cols - 1
        GridWidth = GridWidth + mColWidth(c)
    Next c
End Property

Public Property Get VisibleFixedRows() As Long
    VisibleFixedRows = mVisibleFixedRows
End Property

Public Property Get ColLargerFontSizeFontIndex(col As Long) As Long
    ColLargerFontSizeFontIndex = mColLargerFontSizeFontIndex(col)
End Property

Public Property Get GridType() As efnGridType
    GridType = mGridType
End Property

Public Property Get LastVisibleCol() As Long
    LastVisibleCol = mLastVisibleCol
End Property

Public Property Get LastRowWithData() As Long
    LastRowWithData = mLastRowWithData
End Property

Public Property Let ColWidth(col As Long, nValue As Long)
    mColWidth(col) = nValue
End Property

Public Property Get ColWidth(col As Long) As Long
    ColWidth = mColWidth(col)
End Property


Public Property Let CellText(Row As Long, col As Long, nText As String)
    mCellText(Row, col) = nText
End Property

Public Property Get CellText(Row As Long, col As Long) As String
    CellText = mCellText(Row, col)
End Property


Public Property Let CellWidth(Row As Long, col As Long, nWidth As Long)
    mCellWidth(Row, col) = nWidth
End Property

Public Property Get CellWidth(Row As Long, col As Long) As Long
    CellWidth = mCellWidth(Row, col)
End Property


Public Property Let CellPrintRightLine(Row As Long, col As Long, nPrintRightLine As Boolean)
    mCellPrintRightLine(Row, col) = nPrintRightLine
End Property

Public Property Get CellPrintRightLine(Row As Long, col As Long) As Boolean
    CellPrintRightLine = mCellPrintRightLine(Row, col)
End Property


Public Property Let LastRowMergedInThisCellAtThisCol(Row As Long, col As Long, nLastRow As Long)
    Dim R As Long
    mLastRowMergedInThisCellAtThisCol(Row, col) = nLastRow
    For R = Row To nLastRow
        mRowWhereCellIsMerged(R, col) = Row
        mCellForeColor(R, col) = mCellForeColor(Row, col)
        mCellBackColor(R, col) = mCellBackColor(Row, col)
        mCellFontIndex(R, col) = mCellFontIndex(Row, col)
        mCellText(R, col) = mCellText(Row, col)
        mCellWidth(R, col) = mCellWidth(Row, col)
        mCellPrintRightLine(R, col) = mCellPrintRightLine(Row, col)
        mCellAlignment(R, col) = mCellAlignment(Row, col)
    Next R
End Property

Public Property Get LastRowMergedInThisCellAtThisCol(Row As Long, col As Long) As Long
    LastRowMergedInThisCellAtThisCol = mLastRowMergedInThisCellAtThisCol(Row, col)
End Property


Public Property Let LastColMergedInThisCellAtThisRow(Row As Long, col As Long, nLastCol As Long)
    Dim c As Long
    
    mLastColMergedInThisCellAtThisRow(Row, col) = nLastCol
    
    For c = col + 1 To nLastCol - 1
        mLastColMergedInThisCellAtThisRow(Row, c) = nLastCol
    Next c
    For c = col To nLastCol
        If col < mFirstColMergedInThisCellAtThisRow(Row, c) Then
            mFirstColMergedInThisCellAtThisRow(Row, c) = col
        End If
    Next c
End Property

Public Property Get LastColMergedInThisCellAtThisRow(Row As Long, col As Long) As Long
    LastColMergedInThisCellAtThisRow = mLastColMergedInThisCellAtThisRow(Row, col)
End Property

Public Property Get FirstColMergedInThisCellAtThisRow(Row As Long, col As Long) As Long
    FirstColMergedInThisCellAtThisRow = mFirstColMergedInThisCellAtThisRow(Row, col)
    If FirstColMergedInThisCellAtThisRow = mGrid.Cols Then
        FirstColMergedInThisCellAtThisRow = col
    End If
End Property


Public Property Let RowWhereCellIsMerged(Row As Long, col As Long, nValue As Long)
    mRowWhereCellIsMerged(Row, col) = nValue
End Property

Public Property Get RowWhereCellIsMerged(Row As Long, col As Long) As Long
    RowWhereCellIsMerged = mRowWhereCellIsMerged(Row, col)
End Property


Public Property Let TextHeightRow(Row As Long, nValue As Long)
    mTextHeightRow(Row) = nValue
End Property

Public Property Get TextHeightRow(Row As Long) As Long
    TextHeightRow = mTextHeightRow(Row)
End Property


Public Property Let RowFinalPosYForTop(Row As Long, nValue As Long)
    mRowFinalPosYForTop(Row) = nValue
End Property

Public Property Get RowFinalPosYForTop(Row As Long) As Long
    RowFinalPosYForTop = mRowFinalPosYForTop(Row)
End Property


Public Property Let TextHeightCell(Row As Long, col As Long, nValue As Long)
    mTextHeightCell(Row, col) = nValue
End Property

Public Property Get TextHeightCell(Row As Long, col As Long) As Long
    TextHeightCell = mTextHeightCell(Row, col)
End Property


Public Property Let PrintCell(Row As Long, col As Long, nValue As Boolean)
    mPrintCell(Row, col) = nValue
End Property

Public Property Get PrintCell(Row As Long, col As Long) As Boolean
    PrintCell = mPrintCell(Row, col)
End Property


Public Property Let MinColWidth(col As Long, nValue As Long)
    mMinColWidth(col) = nValue
End Property

Public Property Get MinColWidth(col As Long) As Long
    MinColWidth = mMinColWidth(col)
End Property


Public Property Let MinColWidthFontReduced(col As Long, nValue As Long)
    mMinColWidthFontReduced(col) = nValue
End Property

Public Property Get MinColWidthFontReduced(col As Long) As Long
    MinColWidthFontReduced = mMinColWidthFontReduced(col)
End Property


Public Sub SetColsWidthsToProportion(nProportion As Single)
    Dim c As Long
    
    For c = 0 To mGrid.Cols - 1
        mColWidth(c) = mColOriginalWidth(c) * nProportion
    Next c
    mColWidthsAreSet = True
End Sub


Public Sub SetRowsHeightsToProportion(nProportion As Single)
    Dim R As Long
    
    For R = 0 To mGrid.Rows - 1
        mRowHeight(R) = mGrid.RowHeight(R) * nProportion
    Next R
End Sub

Public Sub CalcColPos()
    Dim c As Long
    Dim iLng As Long
    
    For c = 0 To UBound(mColWidth)
        mColPosLeft(c) = iLng
        iLng = iLng + mColWidth(c)
        mColPosRight(c) = iLng
    Next c
End Sub


Public Function GetRowHeightWithText(Row As Long) As Long
    GetRowHeightWithText = mRowHeight(Row)
    If mTextHeightRow(Row) > GetRowHeightWithText Then
        GetRowHeightWithText = mTextHeightRow(Row)
    Else
        If mTextHeightRow(Row) > 0 Then
            If GetRowHeightWithText > (mTextHeightRow(Row) * 1.1) Then
                GetRowHeightWithText = mTextHeightRow(Row) * 1.1
            End If
        End If
    End If
End Function


Public Sub StartMergingColForPrinting(col As Long, nRowWithData As Long, nStartRowForPrinting As Long)
    mColMergeForPrintingBegunAtRow(col) = nStartRowForPrinting
    mRowWidthDataForColMergedForPrinting(col) = nRowWithData
End Sub

Public Property Get ColMergeForPrintingBegunAtRow(col As Long) As Long
    ColMergeForPrintingBegunAtRow = mColMergeForPrintingBegunAtRow(col)
End Property

Public Property Get ColIsBeingMergedForPrintingInRow(col As Long) As Long
    ColIsBeingMergedForPrintingInRow = mColMergeForPrintingBegunAtRow(col)
End Property

Public Sub EndMergingColForPrinting(col As Long)
    mColMergeForPrintingBegunAtRow(col) = -1
End Sub

Public Property Get CellIsColMergedInRow(Row, col) As Long
    CellIsColMergedInRow = mRowWhereCellIsMerged(Row, col)
End Property


Public Property Let CellTop(Row As Long, col As Long, nValue As Long)
    mCellTop(Row, col) = nValue
End Property

Public Property Get CellTop(Row As Long, col As Long) As Long
    CellTop = mCellTop(Row, col)
End Property

Public Property Get ColIsVisible(col As Long) As Boolean
    ColIsVisible = mColIsVisible(col)
End Property

Public Property Get RowIsVisible(Row As Long) As Boolean
    If Row <= UBound(mRowIsVisible) Then
        RowIsVisible = mRowIsVisible(Row)
    End If
End Property

Public Property Get RowWidthDataForColMergedForPrinting(col As Long) As Long
    RowWidthDataForColMergedForPrinting = mRowWidthDataForColMergedForPrinting(col)
End Property

Public Sub ResetAddedData(nBackColorFixed As Long)
    Dim c As Long
    Dim R As Long
    Dim iBackColorFixed As Long
    Dim iFixedRows As Long
    Dim iFixedCols As Long
    Dim iCols As Long
    
    iFixedRows = mGrid.FixedRows
    iFixedCols = mGrid.FixedCols
    iCols = mGrid.Cols
    
    ReDim mCellText(mGrid.Rows - 1, mGrid.Cols - 1)
    ReDim mPrintCell(mGrid.Rows - 1, mGrid.Cols - 1)
    ReDim mCellPrintRightLine(mGrid.Rows - 1, mGrid.Cols - 1)
    ReDim mCellWidth(mGrid.Rows - 1, mGrid.Cols - 1)
    ReDim mLastRowMergedInThisCellAtThisCol(mGrid.Rows - 1, mGrid.Cols - 1)
    ReDim mLastColMergedInThisCellAtThisRow(mGrid.Rows - 1, mGrid.Cols - 1)
    ReDim mFirstColMergedInThisCellAtThisRow(mGrid.Rows - 1, mGrid.Cols - 1)
    ReDim mRowWhereCellIsMerged(mGrid.Rows - 1, mGrid.Cols - 1)
    ReDim mTextHeightCell(mGrid.Rows - 1, mGrid.Cols - 1)
    ReDim mCellTop(mGrid.Rows - 1, mGrid.Cols - 1)
    ReDim mColWidth(mGrid.Cols - 1)
    ReDim mRowHeight(mGrid.Rows - 1)
    ReDim mTextHeightRow(mGrid.Rows - 1)
    ReDim mRowFinalPosYForTop(mGrid.Rows - 1)
    ReDim mColPosLeft(mGrid.Cols - 1)
    ReDim mColPosRight(mGrid.Cols - 1)
    ReDim mMinColWidth(mGrid.Cols - 1)
    ReDim mMinColWidthFontReduced(mGrid.Cols - 1)
    ReDim mColMergeForPrintingBegunAtRow(mGrid.Cols - 1)
    ReDim mRowWidthDataForColMergedForPrinting(mGrid.Cols - 1)

    If nBackColorFixed = -1 Then
        iBackColorFixed = mGrid.BackColorFixed
    Else
        iBackColorFixed = nBackColorFixed
    End If
    TranslateColor iBackColorFixed, 0, iBackColorFixed

    For c = 0 To mGrid.Cols - 1
        mColMergeForPrintingBegunAtRow(c) = -1
        mRowWidthDataForColMergedForPrinting(c) = -1
        For R = 0 To mGrid.Rows - 1
            mRowWhereCellIsMerged(R, c) = -1
            mFirstColMergedInThisCellAtThisRow(R, c) = iCols
            
            If (R < iFixedRows) Or (c < iFixedCols) Then
                If Not mCellBackColorIsExplicitelySet Then
                    mCellBackColor(R, c) = iBackColorFixed
                End If
            End If
            
        Next R
    Next c
    
    mColWidthsAreSet = False
End Sub

Public Property Get ColOriginalWidth(col As Long) As Long
    ColOriginalWidth = mColOriginalWidth(col)
End Property

Public Property Get VisibleCols() As Long
    VisibleCols = mVisibleCols
End Property

Public Property Get ColWidthsAreSet() As Boolean
    ColWidthsAreSet = mColWidthsAreSet
End Property
    
