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
Attribute VB_Creatable = True
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
Private mLineWidth As Long
Private mLineWidthHeadersSeparatorLine As Long

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

Private Sub Class_Initialize()
'    Tag = "GRStyle2"
    
    LineWidth = 3
    PrintHeadersBackground = True
    HeadersBackgroundColor = RGB(244, 249, 255)
    PrintFixedColsBackground = False
    PrintOtherBackgrounds = True
    
    PrintOuterBorder = False
    OuterBorderColor = RGB(40, 5, 207)
    PrintHeadersBorder = True
    HeadersBorderColor = RGB(29, 4, 145)
    PrintColumnsDataLines = True
    ColumnsDataLinesColor = RGB(177, 162, 253)
    PrintColumnsHeadersLines = True
    ColumnsHeadersLinesColor = RGB(177, 162, 253)
    PrintRowsLines = False
    RowsLinesColor = RGB(230, 224, 254)
    PrintHeadersSeparatorLine = True
    LineWidthHeadersSeparatorLine = 3

End Sub


Public Property Let LineWidth(ByVal nValue As Long)
    If (nValue < 1) Or (nValue > 80) Then
        RaiseError 380, TypeName(Me), "The thickness value of the lines must be between 1 and 80"
        Exit Property
    End If
    mLineWidth = nValue
End Property

Public Property Get LineWidth() As Long
    LineWidth = mLineWidth
End Property


Public Property Let LineWidthHeadersSeparatorLine(ByVal nValue As Long)
    If (nValue < 1) Or (nValue > 80) Then
        RaiseError 380, TypeName(Me), "The thickness value of the lines must be between 1 and 80"
        Exit Property
    End If
    mLineWidthHeadersSeparatorLine = nValue
End Property

Public Property Get LineWidthHeadersSeparatorLine() As Long
    LineWidthHeadersSeparatorLine = mLineWidthHeadersSeparatorLine
End Property

Public Sub Load(nStyleID)
    Dim iStr As String
    Dim iSt As GridReportStyle
    
    If IsNumeric(nStyleID) Then
        iStr = "GRStyle" & Val(nStyleID)
    Else
        If VarType(nStyleID) = vbString Then
            iStr = CStr(nStyleID)
        ElseIf VarType(nStyleID) = vbObject Then
            If TypeOf nStyleID Is GridReportStyle Then
                Set iSt = nStyleID
            Else
                RaiseError 380, TypeName(Me)
            End If
        Else
            RaiseError 380, TypeName(Me)
            Exit Sub
        End If
    End If
    
    If iSt Is Nothing Then
        Set iSt = GetGridReportStyle(iStr)
    End If
    
    PrintOuterBorder = iSt.PrintOuterBorder
    PrintHeadersBorder = iSt.PrintHeadersBorder
    PrintColumnsDataLines = iSt.PrintColumnsDataLines
    PrintColumnsHeadersLines = iSt.PrintColumnsHeadersLines
    PrintRowsLines = iSt.PrintRowsLines
    PrintHeadersSeparatorLine = iSt.PrintHeadersSeparatorLine
    PrintHeadersBackground = iSt.PrintHeadersBackground
    PrintFixedColsBackground = iSt.PrintFixedColsBackground
    PrintOtherBackgrounds = iSt.PrintOtherBackgrounds
    LineWidth = iSt.LineWidth
    LineWidthHeadersSeparatorLine = iSt.LineWidthHeadersSeparatorLine
    OuterBorderColor = iSt.OuterBorderColor
    ColumnsDataLinesColor = iSt.ColumnsDataLinesColor
    ColumnsHeadersLinesColor = iSt.ColumnsHeadersLinesColor
    RowsLinesColor = iSt.RowsLinesColor
    HeadersBorderColor = iSt.HeadersBorderColor
    HeadersBackgroundColor = iSt.HeadersBackgroundColor
    Tag = iSt.Tag
    
End Sub
