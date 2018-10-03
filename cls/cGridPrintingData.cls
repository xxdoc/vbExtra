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

Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long

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
Private mColLargerFontSizeFontIndex() As Long
Private mColIsVisible() As Boolean
Private mRowIsVisible() As Boolean
Private mVisibleCols As Long
Private mFontsAttributes() As cFontAttributes
Private mGrid As Object
Private mVisibleFixedRows As Long
Private mLastVisibleCol As Long
Private mFirstVisibleCol As Long
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
Private mSizeStrStoredGrid As String
Private mFixedRows As Long
Private mFixedCols As Long
Private mBackColor As Long
Private mBackColorFixed As Long
Private mTextMatrix() As String

Public Canceled As Boolean
Public mfrmSettingGridDataProgressShown As Boolean

Public Sub SetGrid(nGrid As Object, Optional nBackColorFixed As Long = -1, Optional nPrintCellsFormatting As vbExPrintCellsFormatting = vxPCFPrintAllFormatting)
    Dim c1 As Long
    Dim r1 As Long
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
    Dim iPrintCellsColors As Boolean
    Dim iPrintCellsFormat As Boolean
    Dim iPrintFontFormat As Boolean
    
    Set mGrid = nGrid
    
    iPrintCellsFormat = nPrintCellsFormatting <> vxPCFPrintNone
    iPrintCellsColors = (nPrintCellsFormatting = vxPCFPrintAllFormatting) Or (nPrintCellsFormatting = vxPCFIgnoreFont)
    iPrintFontFormat = (nPrintCellsFormatting = vxPCFPrintAllFormatting) Or (nPrintCellsFormatting = vxPCFIgnoreColors)
    
    'ReDim mFontsAttributes(0)
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
    ReDim mTextMatrix(mGrid.Rows - 1, mGrid.Cols - 1)
    
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
    
    mFixedRows = mGrid.FixedRows
    mFixedCols = mGrid.FixedCols
    
    mVisibleFixedRows = 0
    For r1 = 0 To mFixedRows - 1
        If mGrid.RowHeight(r1) <> 0 Then
            mVisibleFixedRows = mVisibleFixedRows + 1
        End If
    Next r1
    
    For r1 = 0 To mGrid.Rows - 1
        If mGrid.RowHeight(r1) = -1 Then
            mRowIsVisible(r1) = True
        Else
            mRowIsVisible(r1) = mGrid.RowHeight(r1) > 35
        End If
    Next r1
    
    iAuxMergeRow0 = mGrid.MergeRow(0)
    mGrid.MergeRow(0) = False
    mFirstVisibleCol = -1
    For c1 = 0 To mGrid.Cols - 1
        If mGrid.ColWidth(c1) <> 0 Then
            mLastVisibleCol = c1
            If mFirstVisibleCol = -1 Then
                mFirstVisibleCol = c1
            End If
        End If
        
        If mGridType = efnGridTypeMSFlex Then
            mColOriginalWidth(c1) = mGrid.ColWidth(c1)
        Else
            If mGrid.ColWidth(c1) <> -1 Then
                mColOriginalWidth(c1) = mGrid.ColWidth(c1)
            Else
                If mGrid.Rows > 0 Then
                    mGrid.Row = 0
                    mGrid.Col = c1
                    mColOriginalWidth(c1) = mGrid.CellWidth + iTx
                Else
                    mColOriginalWidth(c1) = 960
                End If
            End If
        End If
    Next c1
    mGrid.MergeRow(0) = iAuxMergeRow0
    
    mLastRowWithData = -1
    For r1 = mGrid.Rows - 1 To mGrid.FixedRows Step -1
        For c1 = 0 To mGrid.Cols - 1
            If mGrid.TextMatrix(r1, c1) <> "" Then
                mLastRowWithData = r1
                Exit For
            End If
        Next c1
        If mLastRowWithData <> -1 Then Exit For
    Next r1
    
    If mLastRowWithData = -1 Then
        mLastRowWithData = mFixedRows
        If mLastRowWithData > (mGrid.Rows - 1) Then
            mLastRowWithData = mGrid.Rows - 1
        End If
    End If
    
    iBackColorFixed = mGrid.BackColorFixed
    
    iGridBackColorFixed = mGrid.BackColorFixed
    TranslateColor mGrid.BackColor, 0, iBackColor
    TranslateColor mGrid.ForeColor, 0, iForeColor
    TranslateColor iForeColorFixed, 0, iForeColorFixed
    
    If (mGrid.Rows > 3000) And iPrintCellsFormat Then
        frmSettingGridDataProgress.Show
        frmSettingGridDataProgress.pgb1.Max = mGrid.Rows * mGrid.Cols / 1000 * 3
        mfrmSettingGridDataProgressShown = True
        DoEvents
    End If
    
    iFontName_Prev = mGrid.Font.Name
    iFontSize_Prev = mGrid.Font.Size
    iFontBold_prev = mGrid.Font.Bold
    iFontItalic_Prev = mGrid.Font.Italic
    iFontUnderline_Prev = mGrid.Font.Underline
    On Error Resume Next
    iFontStrikethrough_Prev = mGrid.Font.Strikethrough
    On Error GoTo 0
    
    ReDim mFontsAttributes(1)
    Set mFontsAttributes(1) = New cFontAttributes
    
    mFontsAttributes(1).Name = iFontName_Prev
    mFontsAttributes(1).Size = iFontSize_Prev
    mFontsAttributes(1).Bold = iFontBold_prev
    mFontsAttributes(1).Italic = iFontItalic_Prev
    mFontsAttributes(1).Underline = iFontUnderline_Prev
    mFontsAttributes(1).Strikethrough = iFontStrikethrough_Prev
    mFontsAttributes(1).Width = iFontWidth_Prev
    iCurrentFontIndex = 1
    
    For c1 = 0 To mGrid.Cols - 1
'        If iPrintCellsFormat Or (c1 < mFixedCols) Then
        mGrid.Col = c1
'        End If
        mColLargerFontSize(c1) = iFontSize_Prev
        mColLargerFontSizeFontIndex(c1) = iCurrentFontIndex
        mColIsVisible(c1) = mColOriginalWidth(c1) > 35
        If mColIsVisible(c1) Then
            mVisibleCols = mVisibleCols + 1
        End If
        
        For r1 = 0 To mLastRowWithData
            mTextMatrix(r1, c1) = mGrid.TextMatrix(r1, c1)
            If mfrmSettingGridDataProgressShown Then
                If r1 Mod 1000 = 0 Then
                    frmSettingGridDataProgress.pgb1.Value = (r1 + c1 * mLastRowWithData) / 1000
                    frmSettingGridDataProgress.ZOrder
                    DoEvents
                    If frmSettingGridDataProgress.Canceled Then
                        Unload frmSettingGridDataProgress
                        Set frmSettingGridDataProgress = Nothing
                        mfrmSettingGridDataProgressShown = False
                        Canceled = True
                        Exit Sub
                    End If
                End If
            End If
            If iPrintCellsFormat Or ((r1 < mFixedRows) Or (c1 < mFixedCols)) Then
                mGrid.Row = r1
                
                If mGridType = efnGridTypeMSFlex Then ' Grid type MSFlexGrid
                    ' CellAlignment
                    iLng = mGrid.CellAlignment
                    If iLng = 0 Then
                        If r1 < mFixedRows Then
                            mCellAlignment(r1, c1) = mGrid.FixedAlignment(c1)
                        Else
                            mCellAlignment(r1, c1) = mGrid.ColAlignment(c1)
                        End If
                    Else
                        mCellAlignment(r1, c1) = iLng
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
                        If Not ((r1 < mFixedRows) Or (c1 < mFixedCols)) Then
                            mCellBackColor(r1, c1) = iBackColor
                        End If
                    Else
                        If iPrintCellsColors Then
                            mCellBackColor(r1, c1) = iLng
                            mCellBackColorIsExplicitelySet(r1, c1) = True
                        Else
                            mCellBackColor(r1, c1) = iBackColor
                        End If
                    End If
                    
                    'CellForeColor
                    TranslateColor mGrid.CellForeColor, 0, iLng
                    If iLng = 0 Then
                        If (r1 < mFixedRows) Or (c1 < mFixedCols) Then
                            mCellForeColor(r1, c1) = iForeColorFixed
                        Else
                            mCellForeColor(r1, c1) = iForeColor
                        End If
                    Else
                        If iPrintCellsColors Then
                            mCellForeColor(r1, c1) = iLng
                        Else
                            mCellForeColor(r1, c1) = iForeColor
                        End If
                    End If
                    
                Else ' Grid type other than MSFlexGrid (MSHFlexGrid, Etc.)
                    ' CellAlignment
                    mCellAlignment(r1, c1) = mGrid.CellAlignment
                    ' CellBackColor
                    If (r1 < mFixedRows) Or (c1 < mFixedCols) Then
                        If iPrintCellsColors Then
                            iGridCellBackColor = mGrid.CellBackColor
                            If iGridCellBackColor = 0 Then
                                TranslateColor mGrid.CellForeColor, 0, iLng2
                                If iLng2 = 0 Then
                                    iGridCellBackColor = vbWhite
                                End If
                            End If
                            If iGridCellBackColor <> iGridBackColorFixed Then
                                TranslateColor mGrid.CellBackColor, 0, iLng
                                mCellBackColor(r1, c1) = iLng
                                mCellBackColorIsExplicitelySet(r1, c1) = True
                            End If
                        Else
                            If nBackColorFixed <> -1 Then
                                iLng = nBackColorFixed
                            Else
                                iLng = iBackColorFixed
                            End If
                            mCellBackColor(r1, c1) = iLng
                        End If
                    Else
                        If iPrintCellsColors Then
                            iLng = mGrid.CellBackColor
                            If iLng = 0 Then
                                TranslateColor mGrid.CellForeColor, 0, iLng2
                                If iLng2 = 0 Then
                                    iLng = vbWhite
                                End If
                            End If
                        Else
                            iLng = iBackColor
                        End If
                        If iLng = iBackColorFixed Then
                            If nBackColorFixed <> -1 Then
                                iLng = nBackColorFixed
                            End If
                        End If
                        TranslateColor iLng, 0, iLng
                        mCellBackColor(r1, c1) = iLng
                    End If
                    
                    ' CellForeColor
                    If (r1 < mFixedRows) Or (c1 < mFixedCols) Then
                        If iPrintCellsColors Then
                            TranslateColor mGrid.CellForeColor, 0, iLng
                        Else
                            iLng = iForeColor
                        End If
                        mCellForeColor(r1, c1) = iLng
                    Else
                        If iPrintCellsColors Then
                            TranslateColor mGrid.CellForeColor, 0, iLng
                        Else
                            iLng = iForeColor
                        End If
                        mCellForeColor(r1, c1) = iLng
                    End If
                End If
                
                If iPrintFontFormat Or ((r1 < mFixedRows) Or (c1 < mFixedCols)) Then
                    ' CellFont...
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
                        
                        If iFontSize_Prev > mColLargerFontSize(c1) Then
                            mColLargerFontSize(c1) = iFontSize_Prev
                            mColLargerFontSizeFontIndex(c1) = iCurrentFontIndex
                        End If
                    
                    End If
                    ' Font
                    mCellFontIndex(r1, c1) = iCurrentFontIndex
                Else
                    mCellFontIndex(r1, c1) = 1
                End If
            Else
                ' CellAlignment
                mCellAlignment(r1, c1) = mGrid.ColAlignment(c1)
                ' CellBackColor
                If (r1 < mFixedRows) Or (c1 < mFixedCols) Then
                    If nBackColorFixed <> -1 Then
                        iLng = nBackColorFixed
                    Else
                        iLng = iBackColorFixed
                    End If
                Else
                    iLng = iBackColor
                End If
                mCellBackColor(r1, c1) = iLng
                ' CellForeColor
                If (r1 < mFixedRows) Or (c1 < mFixedCols) Then
                     iLng = iForeColorFixed
                Else
                    iLng = iForeColor
                End If
                mCellForeColor(r1, c1) = iLng
                ' Font
                mCellFontIndex(r1, c1) = 1
            End If
        Next r1
    Next c1
    
    ResetAddedData nBackColorFixed
    mSizeStrStoredGrid = GetGridSizeStr(mGrid)

    mBackColor = iBackColor
    If nBackColorFixed <> -1 Then
        mBackColorFixed = nBackColorFixed
    Else
        mBackColorFixed = mGrid.BackColor
    End If
    
End Sub


Public Property Get CellFontAttibutes(Row As Long, Col As Long) As cFontAttributes
    Set CellFontAttibutes = mFontsAttributes(mCellFontIndex(Row, Col))
    If CellFontAttibutes Is Nothing Then
        Set CellFontAttibutes = mFontsAttributes(1)
    End If
End Property

Public Property Get CellFontAttibutesByFontAttibutesIndex(ByVal FontAttibutesIndex As Long) As cFontAttributes
    If FontAttibutesIndex = 0 Then FontAttibutesIndex = 1
    Set CellFontAttibutesByFontAttibutesIndex = mFontsAttributes(FontAttibutesIndex)
End Property

Public Property Get CellFontAttibutesIndex(Row As Long, Col As Long) As Long
    CellFontAttibutesIndex = mCellFontIndex(Row, Col)
End Property

Public Property Get RowHeight(Row As Long) As Long
    RowHeight = mRowHeight(Row)
End Property

Public Property Get ColPosLeft(Col As Long) As Long
    ColPosLeft = mColPosLeft(Col)
End Property

Public Property Get ColPosRight(Col As Long) As Long
    ColPosRight = mColPosRight(Col)
End Property

Public Property Get CellAlignment(Row As Long, Col As Long) As Long
    CellAlignment = mCellAlignment(Row, Col)
End Property

Public Property Get CellBackColor(Row As Long, Col As Long) As Long
    If Row <= mLastRowWithData Then
        CellBackColor = mCellBackColor(Row, Col)
    Else
        If Col < mFixedCols Then
            CellBackColor = mBackColorFixed
        Else
            CellBackColor = mBackColor
        End If
    End If
End Property

Public Property Get CellForeColor(Row As Long, Col As Long) As Long
    CellForeColor = mCellForeColor(Row, Col)
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

Public Property Get ColLargerFontSizeFontIndex(Col As Long) As Long
    ColLargerFontSizeFontIndex = mColLargerFontSizeFontIndex(Col)
End Property

Public Property Get GridType() As efnGridType
    GridType = mGridType
End Property

Public Property Get LastVisibleCol() As Long
    LastVisibleCol = mLastVisibleCol
End Property

Public Property Get FirstVisibleCol() As Long
    FirstVisibleCol = mFirstVisibleCol
End Property

Public Property Get LastRowWithData() As Long
    LastRowWithData = mLastRowWithData
End Property

Public Property Let ColWidth(Col As Long, nValue As Long)
    mColWidth(Col) = nValue
End Property

Public Property Get ColWidth(Col As Long) As Long
    ColWidth = mColWidth(Col)
End Property


Public Property Let CellText(Row As Long, Col As Long, nText As String)
    mCellText(Row, Col) = nText
End Property

Public Property Get CellText(Row As Long, Col As Long) As String
    CellText = mCellText(Row, Col)
End Property


Public Property Let CellWidth(Row As Long, Col As Long, nWidth As Long)
    mCellWidth(Row, Col) = nWidth
End Property

Public Property Get CellWidth(Row As Long, Col As Long) As Long
    CellWidth = mCellWidth(Row, Col)
End Property


Public Property Let CellPrintRightLine(Row As Long, Col As Long, nPrintRightLine As Boolean)
    mCellPrintRightLine(Row, Col) = nPrintRightLine
End Property

Public Property Get CellPrintRightLine(Row As Long, Col As Long) As Boolean
    CellPrintRightLine = mCellPrintRightLine(Row, Col)
End Property


Public Property Let LastRowMergedInThisCellAtThisCol(Row As Long, Col As Long, nLastRow As Long)
    Dim R As Long
    mLastRowMergedInThisCellAtThisCol(Row, Col) = nLastRow
    For R = Row To nLastRow
        mRowWhereCellIsMerged(R, Col) = Row
        mCellForeColor(R, Col) = mCellForeColor(Row, Col)
        mCellBackColor(R, Col) = mCellBackColor(Row, Col)
        mCellFontIndex(R, Col) = mCellFontIndex(Row, Col)
        mCellText(R, Col) = mCellText(Row, Col)
        mCellWidth(R, Col) = mCellWidth(Row, Col)
        mCellPrintRightLine(R, Col) = mCellPrintRightLine(Row, Col)
        mCellAlignment(R, Col) = mCellAlignment(Row, Col)
    Next R
End Property

Public Property Get LastRowMergedInThisCellAtThisCol(Row As Long, Col As Long) As Long
    LastRowMergedInThisCellAtThisCol = mLastRowMergedInThisCellAtThisCol(Row, Col)
End Property


Public Property Let LastColMergedInThisCellAtThisRow(Row As Long, Col As Long, nLastCol As Long)
    Dim c As Long
    
    mLastColMergedInThisCellAtThisRow(Row, Col) = nLastCol
    
    For c = Col + 1 To nLastCol - 1
        mLastColMergedInThisCellAtThisRow(Row, c) = nLastCol
    Next c
    For c = Col To nLastCol
        If Col < mFirstColMergedInThisCellAtThisRow(Row, c) Then
            mFirstColMergedInThisCellAtThisRow(Row, c) = Col
        End If
    Next c
End Property

Public Property Get LastColMergedInThisCellAtThisRow(Row As Long, Col As Long) As Long
    LastColMergedInThisCellAtThisRow = mLastColMergedInThisCellAtThisRow(Row, Col)
End Property

Public Property Get FirstColMergedInThisCellAtThisRow(Row As Long, Col As Long) As Long
    FirstColMergedInThisCellAtThisRow = mFirstColMergedInThisCellAtThisRow(Row, Col)
    If FirstColMergedInThisCellAtThisRow = mGrid.Cols Then
        FirstColMergedInThisCellAtThisRow = Col
    End If
End Property


Public Property Let RowWhereCellIsMerged(Row As Long, Col As Long, nValue As Long)
    mRowWhereCellIsMerged(Row, Col) = nValue
End Property

Public Property Get RowWhereCellIsMerged(Row As Long, Col As Long) As Long
    RowWhereCellIsMerged = mRowWhereCellIsMerged(Row, Col)
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


Public Property Let TextHeightCell(Row As Long, Col As Long, nValue As Long)
    mTextHeightCell(Row, Col) = nValue
End Property

Public Property Get TextHeightCell(Row As Long, Col As Long) As Long
    TextHeightCell = mTextHeightCell(Row, Col)
End Property


Public Property Let PrintCell(Row As Long, Col As Long, nValue As Boolean)
    mPrintCell(Row, Col) = nValue
End Property

Public Property Get PrintCell(Row As Long, Col As Long) As Boolean
    PrintCell = mPrintCell(Row, Col)
End Property


Public Property Let MinColWidth(Col As Long, nValue As Long)
    mMinColWidth(Col) = nValue
End Property

Public Property Get MinColWidth(Col As Long) As Long
    MinColWidth = mMinColWidth(Col)
End Property


Public Property Let MinColWidthFontReduced(Col As Long, nValue As Long)
    mMinColWidthFontReduced(Col) = nValue
End Property

Public Property Get MinColWidthFontReduced(Col As Long) As Long
    MinColWidthFontReduced = mMinColWidthFontReduced(Col)
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


Public Sub StartMergingColForPrinting(Col As Long, nRowWithData As Long, nStartRowForPrinting As Long)
    mColMergeForPrintingBegunAtRow(Col) = nStartRowForPrinting
    mRowWidthDataForColMergedForPrinting(Col) = nRowWithData
End Sub

Public Property Get ColMergeForPrintingBegunAtRow(Col As Long) As Long
    ColMergeForPrintingBegunAtRow = mColMergeForPrintingBegunAtRow(Col)
End Property

Public Property Get ColIsBeingMergedForPrintingInRow(Col As Long) As Long
    ColIsBeingMergedForPrintingInRow = mColMergeForPrintingBegunAtRow(Col)
End Property

Public Sub EndMergingColForPrinting(Col As Long)
    mColMergeForPrintingBegunAtRow(Col) = -1
End Sub

Public Property Get CellIsColMergedInRow(Row, Col) As Long
    CellIsColMergedInRow = mRowWhereCellIsMerged(Row, Col)
End Property


Public Property Let CellTop(Row As Long, Col As Long, nValue As Long)
    mCellTop(Row, Col) = nValue
End Property

Public Property Get CellTop(Row As Long, Col As Long) As Long
    CellTop = mCellTop(Row, Col)
End Property

Public Property Get ColIsVisible(Col As Long) As Boolean
    ColIsVisible = mColIsVisible(Col)
End Property

Public Property Get RowIsVisible(Row As Long) As Boolean
    If Row <= UBound(mRowIsVisible) Then
        RowIsVisible = mRowIsVisible(Row)
    End If
End Property

Public Property Get RowWidthDataForColMergedForPrinting(Col As Long) As Long
    RowWidthDataForColMergedForPrinting = mRowWidthDataForColMergedForPrinting(Col)
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

Public Property Get ColOriginalWidth(Col As Long) As Long
    ColOriginalWidth = mColOriginalWidth(Col)
End Property

Public Property Get VisibleCols() As Long
    VisibleCols = mVisibleCols
End Property

Public Property Get ColWidthsAreSet() As Boolean
    ColWidthsAreSet = mColWidthsAreSet
End Property
    
Public Function GetGridSizeStr(nGrid As Object) As String
    On Error Resume Next
    GetGridSizeStr = nGrid.Rows & "-" & nGrid.Cols & "-" & nGrid.FixedRows & "-" & nGrid.FixedCols
End Function

Public Property Get SizeStrStoredGrid() As String
    SizeStrStoredGrid = mSizeStrStoredGrid
End Property

Public Property Get TextMatrix(Row As Long, Col As Long) As String
    TextMatrix = mTextMatrix(Row, Col)
End Property
