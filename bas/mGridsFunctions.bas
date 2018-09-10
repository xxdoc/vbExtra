Attribute VB_Name = "mGridsFunctions"
Option Explicit

Public Enum efnFlexSortSettings
    flexSortNone = 0
    flexSortGenericAscending = 1
    flexSortGenericDescending = 2
    flexSortNumericAscending = 3
    flexSortNumericDescending = 4
    flexSortStringNoCaseAscending = 5
    flexSortStringNoCaseDescending = 6
    flexSortStringAscending = 7
    flexSortStringDescending = 8
End Enum

Private mGridsArrowUpImageCollection As New Collection
Private mGridsArrowDownImageCollection As New Collection

Public Function GetGridArrowUpImage(nGrid As Object) As StdPicture
    On Error Resume Next
    Set GetGridArrowUpImage = mGridsArrowUpImageCollection(CStr(nGrid.BackColorFixed))
    On Error GoTo 0
    If GetGridArrowUpImage Is Nothing Then
        StoreGridArrowImages nGrid
    End If
    On Error Resume Next
    Set GetGridArrowUpImage = mGridsArrowUpImageCollection(CStr(nGrid.BackColorFixed))
    On Error GoTo 0
End Function

Public Function GetGridArrowDownImage(nGrid As Object) As StdPicture
    On Error Resume Next
    Set GetGridArrowDownImage = mGridsArrowDownImageCollection(CStr(nGrid.BackColorFixed))
    On Error GoTo 0
    If GetGridArrowDownImage Is Nothing Then
        StoreGridArrowImages nGrid
    End If
    On Error Resume Next
    Set GetGridArrowDownImage = mGridsArrowDownImageCollection(CStr(nGrid.BackColorFixed))
    On Error GoTo 0
End Function

Private Sub StoreGridArrowImages(nGrid As Object)
    Dim iPic As Control
    Dim iImage As StdPicture
        
    On Error Resume Next
    nGrid.Parent.Controls.Add "VB.Picturebox", "picAux1x"
    Set iPic = nGrid.Parent.Controls("picAux1x")
    On Error GoTo TheExit:
    
    iPic.Width = nGrid.Parent.ScaleX(17 + nGrid.GridLineWidth, vbPixels, nGrid.Parent.ScaleMode)
    iPic.Height = nGrid.Parent.ScaleY(11, vbPixels, nGrid.Parent.ScaleMode)
    iPic.BorderStyle = 0
    iPic.AutoRedraw = True
    iPic.BackColor = nGrid.BackColorFixed
    iPic.ForeColor = nGrid.GridColorFixed
    iPic.ScaleMode = vbPixels
    
    iPic.Line (1, 8)-(15, 8)
    iPic.Line (1, 8)-(8, 1)
    iPic.Line (2, 8)-(8, 2)
    iPic.Line (8, 2)-(14, 8)
    iPic.Line (7, 2)-(13, 8)
    Set iImage = iPic.Image
    
    mGridsArrowUpImageCollection.Add iImage, CStr(nGrid.BackColorFixed)

    iPic.Cls
    Set iPic.Picture = Nothing
    iPic.PaintPicture iImage, 0, 0, iPic.ScaleWidth, iPic.ScaleHeight, 0, iPic.ScaleHeight, iPic.ScaleWidth, -iPic.ScaleHeight, vbSrcCopy
    Set iImage = iPic.Image
    mGridsArrowDownImageCollection.Add iImage, CStr(nGrid.BackColorFixed)
    
    On Error Resume Next
    nGrid.Parent.Controls.Remove "picAux1x"

TheExit:
End Sub

Public Function GetGridReportStyleID(nGridReportStyle As GridReportStyle, nNumberNewCustomStyle As Long) As String
    Dim c As Long
    Dim iGridReportStyle As GridReportStyle
    Dim iStyleID As String
    
    c = 1
    Set iGridReportStyle = GetGridReportStyle("GRStyle" & c)
    Do Until iGridReportStyle.Tag = ""
        If GridReportStylesAreEqual(iGridReportStyle, nGridReportStyle) Then
            iStyleID = iGridReportStyle.Tag
            Exit Do
        End If
        c = c + 1
        Set iGridReportStyle = GetGridReportStyle("GRStyle" & c)
    Loop
    
    If iStyleID = "" Then
        c = 1
        Set iGridReportStyle = GetGridReportStyle("Custom" & c)
        Do Until iGridReportStyle.Tag = ""
            If GridReportStylesAreEqual(iGridReportStyle, nGridReportStyle) Then
                iStyleID = iGridReportStyle.Tag
                Exit Do
            End If
            c = c + 1
            Set iGridReportStyle = GetGridReportStyle("Custom" & c)
        Loop
    End If
    
    nNumberNewCustomStyle = c
    GetGridReportStyleID = iStyleID
End Function

Public Function GridReportStylesAreEqual(nGridReportStyle1 As GridReportStyle, nGridReportStyle2 As GridReportStyle) As Boolean
    Dim iChange As Boolean
    
    If nGridReportStyle1.LineWidth <> nGridReportStyle2.LineWidth Then
        iChange = True
    ElseIf nGridReportStyle1.PrintHeadersBackground <> nGridReportStyle2.PrintHeadersBackground Then
        iChange = True
    ElseIf nGridReportStyle1.HeadersBackgroundColor <> nGridReportStyle2.HeadersBackgroundColor Then
        iChange = True
    ElseIf nGridReportStyle1.PrintFixedColsBackground <> nGridReportStyle2.PrintFixedColsBackground Then
        iChange = True
    ElseIf nGridReportStyle1.PrintOtherBackgrounds <> nGridReportStyle2.PrintOtherBackgrounds Then
        iChange = True
    ElseIf nGridReportStyle1.PrintOuterBorder <> nGridReportStyle2.PrintOuterBorder Then
        iChange = True
    ElseIf nGridReportStyle1.OuterBorderColor <> nGridReportStyle2.OuterBorderColor Then
        iChange = True
    ElseIf nGridReportStyle1.PrintHeadersBorder <> nGridReportStyle2.PrintHeadersBorder Then
        iChange = True
    ElseIf nGridReportStyle1.HeadersBorderColor <> nGridReportStyle2.HeadersBorderColor Then
        iChange = True
    ElseIf nGridReportStyle1.PrintColumnsDataLines <> nGridReportStyle2.PrintColumnsDataLines Then
        iChange = True
    ElseIf nGridReportStyle1.ColumnsDataLinesColor <> nGridReportStyle2.ColumnsDataLinesColor Then
        iChange = True
    ElseIf nGridReportStyle1.PrintColumnsHeadersLines <> nGridReportStyle2.PrintColumnsHeadersLines Then
        iChange = True
    ElseIf nGridReportStyle1.ColumnsHeadersLinesColor <> nGridReportStyle2.ColumnsHeadersLinesColor Then
        iChange = True
    ElseIf nGridReportStyle1.PrintRowsLines <> nGridReportStyle2.PrintRowsLines Then
        iChange = True
    ElseIf nGridReportStyle1.RowsLinesColor <> nGridReportStyle2.RowsLinesColor Then
        iChange = True
    ElseIf nGridReportStyle1.PrintHeadersSeparatorLine <> nGridReportStyle2.PrintHeadersSeparatorLine Then
        iChange = True
    ElseIf nGridReportStyle1.LineWidthHeadersSeparatorLine <> nGridReportStyle2.LineWidthHeadersSeparatorLine Then
        iChange = True
    End If
    
    GridReportStylesAreEqual = Not iChange
End Function

Public Function GetGridReportStyle(nStyleID As String) As GridReportStyle
    Set GetGridReportStyle = New GridReportStyle
    
    Select Case nStyleID
        Case "GRStyle1"
            GetGridReportStyle.Tag = "GRStyle1"
            
            GetGridReportStyle.LineWidth = 3
            GetGridReportStyle.PrintHeadersBackground = True
            GetGridReportStyle.HeadersBackgroundColor = RGB(239, 239, 239)
            GetGridReportStyle.PrintFixedColsBackground = False
            GetGridReportStyle.PrintOtherBackgrounds = True
            
            GetGridReportStyle.PrintOuterBorder = True
            GetGridReportStyle.OuterBorderColor = RGB(31, 31, 31)
            GetGridReportStyle.PrintHeadersBorder = False
            GetGridReportStyle.HeadersBorderColor = RGB(207, 207, 207)
            GetGridReportStyle.PrintColumnsDataLines = True
            GetGridReportStyle.ColumnsDataLinesColor = RGB(207, 207, 207)
            GetGridReportStyle.PrintColumnsHeadersLines = True
            GetGridReportStyle.ColumnsHeadersLinesColor = RGB(207, 207, 207)
            GetGridReportStyle.PrintRowsLines = True
            GetGridReportStyle.RowsLinesColor = RGB(207, 207, 207)
            GetGridReportStyle.PrintHeadersSeparatorLine = True
            GetGridReportStyle.LineWidthHeadersSeparatorLine = 3
        
        Case "GRStyle2"
            GetGridReportStyle.Tag = "GRStyle2"
            
            GetGridReportStyle.LineWidth = 3
            GetGridReportStyle.PrintHeadersBackground = True
            GetGridReportStyle.HeadersBackgroundColor = RGB(244, 249, 255)
            GetGridReportStyle.PrintFixedColsBackground = False
            GetGridReportStyle.PrintOtherBackgrounds = True
            
            GetGridReportStyle.PrintOuterBorder = False
            GetGridReportStyle.OuterBorderColor = RGB(40, 5, 207)
            GetGridReportStyle.PrintHeadersBorder = True
            GetGridReportStyle.HeadersBorderColor = RGB(29, 4, 145)
            GetGridReportStyle.PrintColumnsDataLines = True
            GetGridReportStyle.ColumnsDataLinesColor = RGB(177, 162, 253)
            GetGridReportStyle.PrintColumnsHeadersLines = True
            GetGridReportStyle.ColumnsHeadersLinesColor = RGB(177, 162, 253)
            GetGridReportStyle.PrintRowsLines = False
            GetGridReportStyle.RowsLinesColor = RGB(230, 224, 254)
            GetGridReportStyle.PrintHeadersSeparatorLine = True
            GetGridReportStyle.LineWidthHeadersSeparatorLine = 3
        
        Case "GRStyle3"
            GetGridReportStyle.Tag = "GRStyle3"
            
            GetGridReportStyle.LineWidth = 3
            GetGridReportStyle.PrintHeadersBackground = True
            GetGridReportStyle.HeadersBackgroundColor = RGB(244, 249, 255)
            GetGridReportStyle.PrintFixedColsBackground = False
            GetGridReportStyle.PrintOtherBackgrounds = True
            
            GetGridReportStyle.PrintOuterBorder = False
            GetGridReportStyle.OuterBorderColor = RGB(40, 5, 207)
            GetGridReportStyle.PrintHeadersBorder = False
            GetGridReportStyle.HeadersBorderColor = RGB(29, 4, 145)
            GetGridReportStyle.PrintColumnsDataLines = True
            GetGridReportStyle.ColumnsDataLinesColor = RGB(177, 162, 253)
            GetGridReportStyle.PrintColumnsHeadersLines = True
            GetGridReportStyle.ColumnsHeadersLinesColor = RGB(177, 162, 253)
            GetGridReportStyle.PrintRowsLines = False
            GetGridReportStyle.RowsLinesColor = RGB(230, 224, 254)
            GetGridReportStyle.PrintHeadersSeparatorLine = True
            GetGridReportStyle.LineWidthHeadersSeparatorLine = 10
        
        Case "GRStyle4"
            GetGridReportStyle.Tag = "GRStyle4"
            
            GetGridReportStyle.LineWidth = 3
            GetGridReportStyle.PrintHeadersBackground = True
            GetGridReportStyle.HeadersBackgroundColor = RGB(252, 243, 243)
            GetGridReportStyle.PrintFixedColsBackground = False
            GetGridReportStyle.PrintOtherBackgrounds = True
            
            GetGridReportStyle.PrintOuterBorder = False
            GetGridReportStyle.OuterBorderColor = RGB(139, 40, 35)
            GetGridReportStyle.PrintHeadersBorder = True
            GetGridReportStyle.HeadersBorderColor = RGB(139, 40, 35)
            GetGridReportStyle.PrintColumnsDataLines = True
            GetGridReportStyle.ColumnsDataLinesColor = RGB(237, 197, 194)
            GetGridReportStyle.PrintColumnsHeadersLines = True
            GetGridReportStyle.ColumnsHeadersLinesColor = RGB(237, 197, 194)
            GetGridReportStyle.PrintRowsLines = False
            GetGridReportStyle.RowsLinesColor = RGB(247, 230, 227)
            GetGridReportStyle.PrintHeadersSeparatorLine = True
            GetGridReportStyle.LineWidthHeadersSeparatorLine = 3
        
        Case "GRStyle5"
            GetGridReportStyle.Tag = "GRStyle5"
            
            GetGridReportStyle.LineWidth = 3
            GetGridReportStyle.PrintHeadersBackground = True
            GetGridReportStyle.HeadersBackgroundColor = RGB(244, 249, 255)
            GetGridReportStyle.PrintFixedColsBackground = False
            GetGridReportStyle.PrintOtherBackgrounds = True
            
            GetGridReportStyle.PrintOuterBorder = True
            GetGridReportStyle.OuterBorderColor = RGB(40, 5, 207)
            GetGridReportStyle.PrintHeadersBorder = False
            GetGridReportStyle.HeadersBorderColor = RGB(40, 5, 207)
            GetGridReportStyle.PrintColumnsDataLines = True
            GetGridReportStyle.ColumnsDataLinesColor = RGB(177, 162, 253)
            GetGridReportStyle.PrintColumnsHeadersLines = True
            GetGridReportStyle.ColumnsHeadersLinesColor = RGB(40, 5, 207)
            GetGridReportStyle.PrintRowsLines = False
            GetGridReportStyle.RowsLinesColor = RGB(230, 224, 254)
            GetGridReportStyle.PrintHeadersSeparatorLine = True
            GetGridReportStyle.LineWidthHeadersSeparatorLine = 3
        
        Case "GRStyle6"
            GetGridReportStyle.Tag = "GRStyle6"
            
            GetGridReportStyle.LineWidth = 3
            GetGridReportStyle.PrintHeadersBackground = True
            GetGridReportStyle.HeadersBackgroundColor = RGB(244, 249, 255)
            GetGridReportStyle.PrintFixedColsBackground = False
            GetGridReportStyle.PrintOtherBackgrounds = True
            
            GetGridReportStyle.PrintOuterBorder = True
            GetGridReportStyle.OuterBorderColor = RGB(40, 5, 207)
            GetGridReportStyle.PrintHeadersBorder = False
            GetGridReportStyle.HeadersBorderColor = RGB(40, 5, 207)
            GetGridReportStyle.PrintColumnsDataLines = True
            GetGridReportStyle.ColumnsDataLinesColor = RGB(177, 162, 253)
            GetGridReportStyle.PrintColumnsHeadersLines = True
            GetGridReportStyle.ColumnsHeadersLinesColor = RGB(177, 162, 253)
            GetGridReportStyle.PrintRowsLines = True
            GetGridReportStyle.RowsLinesColor = RGB(230, 224, 254)
            GetGridReportStyle.PrintHeadersSeparatorLine = True
            GetGridReportStyle.LineWidthHeadersSeparatorLine = 3
            
        Case "GRStyle7"
            GetGridReportStyle.Tag = "GRStyle7"
            
            GetGridReportStyle.LineWidth = 3
            GetGridReportStyle.PrintHeadersBackground = True
            GetGridReportStyle.HeadersBackgroundColor = RGB(252, 243, 243)
            GetGridReportStyle.PrintFixedColsBackground = False
            GetGridReportStyle.PrintOtherBackgrounds = True
            
            GetGridReportStyle.PrintOuterBorder = True
            GetGridReportStyle.OuterBorderColor = RGB(139, 40, 35)
            GetGridReportStyle.PrintHeadersBorder = False
            GetGridReportStyle.HeadersBorderColor = RGB(139, 40, 35)
            GetGridReportStyle.PrintColumnsDataLines = True
            GetGridReportStyle.ColumnsDataLinesColor = RGB(237, 197, 194)
            GetGridReportStyle.PrintColumnsHeadersLines = True
            GetGridReportStyle.ColumnsHeadersLinesColor = RGB(237, 197, 194)
            GetGridReportStyle.PrintRowsLines = True
            GetGridReportStyle.RowsLinesColor = RGB(247, 230, 227)
            GetGridReportStyle.PrintHeadersSeparatorLine = True
            GetGridReportStyle.LineWidthHeadersSeparatorLine = 3
        
        Case Else
            If Val(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintOuterBorder", -44)) <> -44 Then
                GetGridReportStyle.Tag = nStyleID
                
                GetGridReportStyle.PrintOuterBorder = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintOuterBorder", -1))
                GetGridReportStyle.PrintHeadersBorder = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintHeadersBorder", -1))
                GetGridReportStyle.PrintColumnsDataLines = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintColumnsDataLines", -1))
                GetGridReportStyle.PrintColumnsHeadersLines = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintColumnsHeadersLines", -1))
                GetGridReportStyle.PrintRowsLines = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintRowsLines", -1))
                GetGridReportStyle.PrintHeadersSeparatorLine = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintHeadersSeparatorLine", -1))
                GetGridReportStyle.PrintHeadersBackground = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintHeadersBackground", -1))
                GetGridReportStyle.PrintFixedColsBackground = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintFixedColsBackground", 0))
                GetGridReportStyle.PrintOtherBackgrounds = CBool(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_PrintOtherBackgrounds", -1))
                GetGridReportStyle.LineWidth = GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_LineWidth", 3)
                GetGridReportStyle.LineWidthHeadersSeparatorLine = GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_LineWidthHeadersSeparatorLine", 3)
            
                GetGridReportStyle.OuterBorderColor = Val(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_OuterBorderColor", RGB(31, 31, 31)))
                GetGridReportStyle.ColumnsDataLinesColor = Val(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_ColumnsDataLinesColor", RGB(207, 207, 207)))
                GetGridReportStyle.ColumnsHeadersLinesColor = Val(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_ColumnsHeadersLinesColor", RGB(207, 207, 207)))
                GetGridReportStyle.RowsLinesColor = Val(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_RowsLinesColor", RGB(207, 207, 207)))
                GetGridReportStyle.HeadersBorderColor = Val(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_HeadersBorderColor", RGB(207, 207, 207)))
                GetGridReportStyle.HeadersBackgroundColor = Val(GetSetting(AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & nStyleID & "_HeadersBackgroundColor", RGB(239, 239, 239)))
            Else
                GetGridReportStyle.Tag = ""
            End If
    End Select
End Function


Public Function GetIDGridReportStyleSaved(nGridReportStyle As GridReportStyle) As String
    Dim iStyleID As String
    Dim iNumberNewCustomStyle As Long
    
    iStyleID = GetGridReportStyleID(nGridReportStyle, iNumberNewCustomStyle)
    
    If iStyleID = "" Then
        If iNumberNewCustomStyle > 10 Then
            iNumberNewCustomStyle = 1
            iNumberNewCustomStyle = Val(GetSetting(AppNameForRegistry, "PrintingSettings", "LastCustomGridReportStyleNumber", 0)) + 1
            If iNumberNewCustomStyle > 10 Then
                iNumberNewCustomStyle = 1
            End If
            SaveSetting AppNameForRegistry, "PrintingSettings", "LastCustomGridReportStyleNumber", iNumberNewCustomStyle
        End If
        iStyleID = "Custom" & iNumberNewCustomStyle
        
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_LineWidth", nGridReportStyle.LineWidth
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintHeadersBackground", CLng(nGridReportStyle.PrintHeadersBackground)
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_HeadersBackgroundColor", nGridReportStyle.HeadersBackgroundColor
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintFixedColsBackground", CLng(nGridReportStyle.PrintFixedColsBackground)
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintOtherBackgrounds", CLng(nGridReportStyle.PrintOtherBackgrounds)
        
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintOuterBorder", CLng(nGridReportStyle.PrintOuterBorder)
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_OuterBorderColor", nGridReportStyle.OuterBorderColor
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintHeadersBorder", CLng(nGridReportStyle.PrintHeadersBorder)
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_HeadersBorderColor", nGridReportStyle.HeadersBorderColor
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintColumnsDataLines", CLng(nGridReportStyle.PrintColumnsDataLines)
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_ColumnsDataLinesColor", nGridReportStyle.ColumnsDataLinesColor
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintColumnsHeadersLines", CLng(nGridReportStyle.PrintColumnsHeadersLines)
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_ColumnsHeadersLinesColor", nGridReportStyle.ColumnsHeadersLinesColor
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintRowsLines", CLng(nGridReportStyle.PrintRowsLines)
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_RowsLinesColor", nGridReportStyle.RowsLinesColor
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_PrintHeadersSeparatorLine", CLng(nGridReportStyle.PrintHeadersSeparatorLine)
        SaveSetting AppNameForRegistry, "PrintingSettings", "GridReportStyle:" & iStyleID & "_LineWidthHeadersSeparatorLine", nGridReportStyle.LineWidthHeadersSeparatorLine
        
    End If
    
    GetIDGridReportStyleSaved = iStyleID
    
End Function

