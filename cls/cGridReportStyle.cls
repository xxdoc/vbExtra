VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GridReportStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public PrintOuterBorder As Boolean
Public PrintHeadersBorder As Boolean
Public PrintColumnsDataLines As Boolean
Public PrintColumnsHeadersLines As Boolean
Public PrintRowsLines As Boolean
Public PrintHeadersSeparatorLine As Boolean
Public PrintHeadersBackground As Boolean
Public PrintFixedColsBackground As Boolean
Public PrintOtherBackgrounds As Boolean
Public LineWidth As Long
Public LineWidthHeadersSeparatorLine As Long

Public OuterBorderColor As Long
Public ColumnsDataLinesColor As Long
Public ColumnsHeadersLinesColor As Long
Public RowsLinesColor As Long
Public HeadersBorderColor As Long
Public HeadersBackgroundColor As Long

Public Tag As String

Public Function Clone() As GridReportStyle
    Set Clone = New GridReportStyle
    
    Clone.PrintOuterBorder = PrintOuterBorder
    Clone.PrintHeadersBorder = PrintHeadersBorder
    Clone.PrintColumnsDataLines = PrintColumnsDataLines
    Clone.PrintColumnsHeadersLines = PrintColumnsHeadersLines
    Clone.PrintRowsLines = PrintRowsLines
    Clone.PrintHeadersSeparatorLine = PrintHeadersSeparatorLine
    Clone.PrintHeadersBackground = PrintHeadersBackground
    Clone.PrintFixedColsBackground = PrintFixedColsBackground
    Clone.PrintOtherBackgrounds = PrintOtherBackgrounds
    Clone.LineWidth = LineWidth
    Clone.LineWidthHeadersSeparatorLine = LineWidthHeadersSeparatorLine
    Clone.OuterBorderColor = OuterBorderColor
    Clone.ColumnsDataLinesColor = ColumnsDataLinesColor
    Clone.ColumnsHeadersLinesColor = ColumnsHeadersLinesColor
    Clone.RowsLinesColor = RowsLinesColor
    Clone.HeadersBorderColor = HeadersBorderColor
    Clone.HeadersBackgroundColor = HeadersBackgroundColor
    Clone.Tag = Tag
    
End Function
