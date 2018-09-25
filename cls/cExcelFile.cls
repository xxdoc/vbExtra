VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExcelFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Class file for writing Microsoft Excel BIFF 2.1 files.

'This class is intended for users who do not want to use the huge
'Jet or ADO providers if they only want to export their data to
'an Excel compatible file.

'Newer versions of Excel use the OLE Structure Storage methods
'which are quite complicated.

'Paul Squires, April 3, 2001
'rambo2000@canada.com

'Added default-cellformats: Dieter Hauk January 8, 2001 dieter.hauk@epost.de

'the memory copy API is used in the MKI$ function which converts an integer
'value to a 2-byte string value to write to the file. (used by the Horizontal
'Page Break function).
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

'enum to handle the various types of values that can be written
'to the excel file.
Public Enum XLSValueTypes
    xlsInteger = 0
    xlsNumber = 1
    xlsText = 2
End Enum

'enum to hold cell alignment
Public Enum XLSCellAlignment
    xlsGeneralAlign = 0
    xlsLeftAlign = 1
    xlsCentreAlign = 2
    xlsrightAlign = 3
    xlsFillCell = 4
    xlsLeftBorder = 8
    xlsRightBorder = 16
    xlsTopBorder = 32
    xlsBottomBorder = 64
    xlsShaded = 128
End Enum

'enum to handle selecting the font for the cell
Public Enum XLSCellFont
    'used by rgbAttr2
    'bits 0-5 handle the *picture* formatting, not bold/underline etc...
    'bits 6-7 handle the font number
    xlsFont0 = 0
    xlsFont1 = 64
    xlsFont2 = 128
    xlsFont3 = 192
End Enum

Public Enum XLSCellHiddenLocked
    'used by rgbAttr1
    'bits 0-5 must be zero
    'bit 6 locked/unlocked
    'bit 7 hidden/not hidden
    xlsNormal = 0
    xlsLocked = 64
    xlsHidden = 128
End Enum

'set up variables to hold the spreadsheet's layout
Public Enum XLSMarginTypes
    xlsLeftMargin = 38
    xlsRightMargin = 39
    xlsTopMargin = 40
    xlsBottomMargin = 41
End Enum

Public Enum XLSFontFormatting
    'add these enums together. For example: xlsBold + xlsUnderline
    xlsNoFormat = 0
    xlsBold = 1
    xlsItalic = 2
    xlsUnderline = 4
    xlsStrikeout = 8
End Enum

Private Type FONT_RECORD
    Opcode As Integer  '49
    Length As Integer  '5+len(fontname)
    FontHeight As Integer

    'bit0 bold, bit1 italic, bit2 underline, bit3 strikeout, bit4-7 reserved
    FontAttributes1 As Byte

    FontAttributes2 As Byte  'reserved - always 0

    FontNameLength As Byte
End Type

Private Type PASSWORD_RECORD
    Opcode As Integer  '47
    Length As Integer  'len(password)
End Type

Private Type HEADER_FOOTER_RECORD
    Opcode As Integer  '20 Header, 21 Footer
    Length As Integer  '1+len(text)
    TextLength As Byte
End Type

Private Type PROTECT_SPREADSHEET_RECORD
    Opcode As Integer  '18
    Length As Integer  '2
    Protect As Integer
End Type

Private Type FORMAT_COUNT_RECORD
    Opcode As Integer  '1f
    Length As Integer '2
    Count As Integer
End Type

Private Type FORMAT_RECORD
    Opcode As Integer  '1e
    Length As Integer  '1+len(format)
    FormatLenght As Byte 'len(format)
End Type '+ followed by the Format-Picture


Private Type COLWIDTH_RECORD
    Opcode As Integer  '36
    Length As Integer  '4
    col1 As Byte       'first column
    col2 As Byte       'last column
    ColumnWidth As Integer   'at 1/256th of a character
End Type

'Beginning Of File record
Private Type BEG_FILE_RECORD
    Opcode As Integer
    Length As Integer
    Version As Integer
    ftype As Integer
End Type

'End Of File record
Private Type END_FILE_RECORD
    Opcode As Integer
    Length As Integer
End Type

'true/false to print gridlines
Private Type PRINT_GRIDLINES_RECORD
    Opcode As Integer
    Length As Integer
    PrintFlag As Integer
End Type

'Integer record
Private Type tInteger
    Opcode As Integer
    Length As Integer
    Row As Integer     'unsigned integer
    Col As Integer

    'rgbAttr1 handles whether cell is hidden and/or locked
    rgbAttr1 As Byte

    'rgbAttr2 handles the Font# and Formatting assigned to this cell
    rgbAttr2 As Byte

    'rgbAttr3 handles the Cell Alignment/borders/shading
    rgbAttr3 As Byte

    intValue As Integer  'the actual integer value
End Type

'Number record
Private Type tNumber
    Opcode As Integer
    Length As Integer
    Row As Integer
    Col As Integer
    rgbAttr1 As Byte
    rgbAttr2 As Byte
    rgbAttr3 As Byte
    NumberValue As Double  '8 Bytes
End Type

'Label (Text) record
Private Type tText
    Opcode As Integer
    Length As Integer
    Row As Integer
    Col As Integer
    rgbAttr1 As Byte
    rgbAttr2 As Byte
    rgbAttr3 As Byte
    TextLength As Byte
End Type

Private Type MARGIN_RECORD_LAYOUT
    Opcode As Integer
    Length As Integer
    MarginValue As Double  '8 bytes
End Type

Private Type HPAGE_BREAK_RECORD
    Opcode As Integer
    Length As Integer
    NumPageBreaks As Integer
End Type

Private FileNumber As Integer
Private BEG_FILE_MARKER As BEG_FILE_RECORD
Private END_FILE_MARKER As END_FILE_RECORD
Private HORIZ_PAGE_BREAK As HPAGE_BREAK_RECORD

'create an array that will hold the rows where a horizontal page
'break will be inserted just before.
Private HorizPageBreakRows() As Integer
Private NumHorizPageBreaks As Integer



Public Function CreateFile(ByVal FileName As String) As Integer

    On Error GoTo Write_Error

    If Dir$(FileName) > "" Then
        Kill FileName
    End If

    FileNumber = FreeFile
    Open FileName For Binary As #FileNumber
    Put #FileNumber, , BEG_FILE_MARKER  'must always be written first

    Call WriteDefaultFormats

    'create the Horizontal Page Break array
    ReDim HorizPageBreakRows(0)
    NumHorizPageBreaks = 0

    CreateFile = 0  'return with no error

    Exit Function

Write_Error:
    CreateFile = Err.Number
    Exit Function

End Function

Public Function CloseFile() As Integer
    Dim x%

    On Error GoTo Write_Error

    If FileNumber = 0 Then Exit Function

    'write the horizontal page breaks if necessary
    If NumHorizPageBreaks > 0 Then
        'the Horizontal Page Break array must be in sorted order.
        'Use a simple Bubble sort because the size of this array would
        'be pretty small most of the time. A QuickSort would probably
        'be overkill.
        Dim lLoop1 As Long
        Dim lLoop2 As Long
        Dim lTemp As Long
        For lLoop1 = UBound(HorizPageBreakRows) To LBound(HorizPageBreakRows) Step -1
            For lLoop2 = LBound(HorizPageBreakRows) + 1 To lLoop1
                If HorizPageBreakRows(lLoop2 - 1) > HorizPageBreakRows(lLoop2) Then
                    lTemp = HorizPageBreakRows(lLoop2 - 1)
                    HorizPageBreakRows(lLoop2 - 1) = HorizPageBreakRows(lLoop2)
                    HorizPageBreakRows(lLoop2) = lTemp
                End If
            Next lLoop2
        Next lLoop1

        'write the Horizontal Page Break Record
        With HORIZ_PAGE_BREAK
            .Opcode = 27
            .Length = 2 + (NumHorizPageBreaks * 2)
            .NumPageBreaks = NumHorizPageBreaks
        End With
        Put #FileNumber, , HORIZ_PAGE_BREAK

        'now write the actual page break values
        'the MKI$ function is standard in other versions of BASIC but
        'VisualBasic does not have it. A KnowledgeBase article explains
        'how to recreate it (albeit using 16-bit API, I switched it
        'to 32-bit).
        For x% = 1 To UBound(HorizPageBreakRows)
            Put #FileNumber, , MKI$(HorizPageBreakRows(x%))
        Next
    End If

    Put #FileNumber, , END_FILE_MARKER
    Close #FileNumber

    CloseFile = 0  'return with no error code

    Exit Function

Write_Error:
    CloseFile = Err.Number
    Exit Function

End Function

Private Sub Class_Initialize()

    'Set up default values for records
    'These should be the values that are the same for every record of these types

    With BEG_FILE_MARKER  'beginning of file
        .Opcode = 9
        .Length = 4
        .Version = 2
        .ftype = 10
    End With

    With END_FILE_MARKER  'end of file marker
        .Opcode = 10
    End With

End Sub

Public Function InsertHorizPageBreak(lrow As Long) As Integer
    Dim Row%

    On Error GoTo Page_Break_Error

    'the row and column values are written to the excel file as
    'unsigned integers. Therefore, must convert the longs to integer.
    If lrow > 32767 Then
        Row% = CInt(lrow - 65536)
    Else
        Row% = CInt(lrow) - 1    'rows/cols in Excel binary file are zero based
    End If

    NumHorizPageBreaks = NumHorizPageBreaks + 1
    ReDim Preserve HorizPageBreakRows(NumHorizPageBreaks)

    HorizPageBreakRows(NumHorizPageBreaks) = Row%

    Exit Function

Page_Break_Error:
    InsertHorizPageBreak = Err.Number
    Exit Function

End Function


Public Function WriteValue(ValueType As XLSValueTypes, CellFontUsed As XLSCellFont, Alignment As XLSCellAlignment, HiddenLocked As XLSCellHiddenLocked, lrow As Long, lcol As Long, Value As Variant, Optional CellFormat As Long = 0) As Integer
    Dim Row%
    Dim Col%
    Dim ST$
    Dim l%
    Dim a As Integer

    On Error GoTo Write_Error

    'the row and column values are written to the excel file as
    'unsigned integers. Therefore, must convert the longs to integer.

    If lrow > 32767 Then
        Row% = CInt(lrow - 65536)
    Else
        Row% = CInt(lrow) - 1    'rows/cols in Excel binary file are zero based
    End If

    If lcol > 32767 Then
        Col% = CInt(lcol - 65536)
    Else
        Col% = CInt(lcol) - 1    'rows/cols in Excel binary file are zero based
    End If

    Select Case ValueType
        Case XLSValueTypes.xlsInteger
            Dim INTEGER_RECORD As tInteger
            With INTEGER_RECORD
                .Opcode = 2
                .Length = 9
                .Row = Row%
                .Col = Col%
                .rgbAttr1 = CByte(HiddenLocked)
                .rgbAttr2 = CByte(CellFontUsed + CellFormat)
                .rgbAttr3 = CByte(Alignment)
                .intValue = CInt(Value)
            End With
            Put #FileNumber, , INTEGER_RECORD

        Case XLSValueTypes.xlsNumber
            Dim NUMBER_RECORD As tNumber
            With NUMBER_RECORD
                .Opcode = 3
                .Length = 15
                .Row = Row%
                .Col = Col%
                .rgbAttr1 = CByte(HiddenLocked)
                .rgbAttr2 = CByte(CellFontUsed + CellFormat)
                .rgbAttr3 = CByte(Alignment)
                .NumberValue = CDbl(Value)
            End With
            Put #FileNumber, , NUMBER_RECORD

        Case XLSValueTypes.xlsText
            Dim B As Byte
            ST$ = CStr(Value)
            If Len(ST$) > 255 Then
                ST$ = Left$(ST$, 252) & "..."
            End If
            l% = Len(ST$)

            Dim TEXT_RECORD As tText
            With TEXT_RECORD
                .Opcode = 4
                .Length = 10
                'Length of the text portion of the record
                .TextLength = l%

                'Total length of the record
                .Length = 8 + l

                .Row = Row%
                .Col = Col%

                .rgbAttr1 = CByte(HiddenLocked)
                .rgbAttr2 = CByte(CellFontUsed + CellFormat)
                .rgbAttr3 = CByte(Alignment)

                'Put record header
                Put #FileNumber, , TEXT_RECORD

                'Then the actual string data
                For a = 1 To l%
                    B = Asc(Mid$(ST$, a, 1))
                    Put #FileNumber, , B
                Next
            End With

    End Select

    WriteValue = 0   'return with no error

    Exit Function

Write_Error:
    WriteValue = Err.Number
    Exit Function

End Function

Public Function SetMargin(Margin As XLSMarginTypes, MarginValue As Double) As Integer

    On Error GoTo Write_Error

    'write the spreadsheet's layout information (in inches)
    Dim MarginRecord As MARGIN_RECORD_LAYOUT

    With MarginRecord
        .Opcode = Margin
        .Length = 8
        .MarginValue = MarginValue 'in inches
    End With
    Put #FileNumber, , MarginRecord

    SetMargin = 0

    Exit Function

Write_Error:
    SetMargin = Err.Number
    Exit Function

End Function

Public Function SetColumnWidth(FirstColumn As Byte, LastColumn As Byte, WidthValue As Integer)

    On Error GoTo Write_Error

    Dim ColWidth As COLWIDTH_RECORD

    With ColWidth
        .Opcode = 36
        .Length = 4
        .col1 = FirstColumn - 1
        .col2 = LastColumn - 1
        .ColumnWidth = WidthValue * 256  'values are specified as 1/256 of a character
    End With
    Put #FileNumber, , ColWidth

    SetColumnWidth = 0

    Exit Function

Write_Error:
    SetColumnWidth = Err.Number
    Exit Function

End Function

Public Function SetFont(FontName As String, FontHeight As Integer, FontFormat As XLSFontFormatting) As Integer
    Dim l%
    Dim a As Integer

    On Error GoTo Write_Error

    'you can set up to 4 fonts in the spreadsheet file. When writing a value such
    'as a Text or Number you can specify one of the 4 fonts (numbered 0 to 3)

    Dim FONTNAME_RECORD As FONT_RECORD

    l% = Len(FontName)

    With FONTNAME_RECORD
        .Opcode = 49
        .Length = 5 + l%
        .FontHeight = FontHeight * 20
        .FontAttributes1 = CByte(FontFormat)  'bold/underline etc...
        .FontAttributes2 = CByte(0) 'reserved-always zero!!
        .FontNameLength = CByte(Len(FontName))
    End With
    Put #FileNumber, , FONTNAME_RECORD

    'Then the actual font name data
    Dim B As Byte
    For a = 1 To l%
        B = Asc(Mid$(FontName, a, 1))
        Put #FileNumber, , B
    Next

    SetFont = 0

    Exit Function

Write_Error:
    SetFont = Err.Number
    Exit Function

End Function

Public Function SetHeader(HeaderText As String) As Integer
    Dim l%
    Dim a As Integer

    On Error GoTo Write_Error

    Dim HEADER_RECORD As HEADER_FOOTER_RECORD

    l% = Len(HeaderText)

    With HEADER_RECORD
        .Opcode = 20
        .Length = 1 + l%
        .TextLength = CByte(Len(HeaderText))
    End With
    Put #FileNumber, , HEADER_RECORD

    'Then the actual Header text
    Dim B As Byte
    For a = 1 To l%
        B = Asc(Mid$(HeaderText, a, 1))
        Put #FileNumber, , B
    Next

    SetHeader = 0

    Exit Function

Write_Error:
    SetHeader = Err.Number
    Exit Function

End Function


Public Function SetFooter(FooterText As String) As Integer
    Dim l%
    Dim a As Integer

    On Error GoTo Write_Error

    Dim FOOTER_RECORD As HEADER_FOOTER_RECORD

    l% = Len(FooterText)

    With FOOTER_RECORD
        .Opcode = 21
        .Length = 1 + l%
        .TextLength = CByte(Len(FooterText))
    End With
    Put #FileNumber, , FOOTER_RECORD

    'Then the actual Header text
    Dim B As Byte
    For a = 1 To l%
        B = Asc(Mid$(FooterText, a, 1))
        Put #FileNumber, , B
    Next

    SetFooter = 0

    Exit Function

Write_Error:
    SetFooter = Err.Number
    Exit Function

End Function


Public Function SetFilePassword(PasswordText As String) As Integer
    Dim l%
    Dim a As Integer

    On Error GoTo Write_Error

    Dim FILE_PASSWORD_RECORD As PASSWORD_RECORD

    l% = Len(PasswordText)

    With FILE_PASSWORD_RECORD
        .Opcode = 47
        .Length = l%
    End With
    Put #FileNumber, , FILE_PASSWORD_RECORD

    'Then the actual Password text
    Dim B As Byte
    For a = 1 To l%
        B = Asc(Mid$(PasswordText, a, 1))
        Put #FileNumber, , B
    Next

    SetFilePassword = 0

    Exit Function

Write_Error:
    SetFilePassword = Err.Number
    Exit Function

End Function



Public Property Let PrintGridLines(ByVal NewValue As Boolean)

    On Error GoTo Write_Error

    Dim GRIDLINES_RECORD As PRINT_GRIDLINES_RECORD

    With GRIDLINES_RECORD
        .Opcode = 43
        .Length = 2
        If NewValue = True Then
            .PrintFlag = 1
        Else
            .PrintFlag = 0
        End If

    End With
    Put #FileNumber, , GRIDLINES_RECORD

    Exit Property

Write_Error:
    Exit Property

End Property



Public Property Let ProtectSpreadsheet(ByVal NewValue As Boolean)

    On Error GoTo Write_Error

    Dim PROTECT_RECORD As PROTECT_SPREADSHEET_RECORD

    With PROTECT_RECORD
        .Opcode = 18
        .Length = 2
        If NewValue = True Then
            .Protect = 1
        Else
            .Protect = 0
        End If

    End With
    Put #FileNumber, , PROTECT_RECORD

    Exit Property

Write_Error:
    Exit Property

End Property

Public Function WriteDefaultFormats() As Integer

    Dim cFORMAT_COUNT_RECORD As FORMAT_COUNT_RECORD
    Dim cFORMAT_RECORD As FORMAT_RECORD
    Dim lIndex As Long
    Dim aFormat(0 To 23) As String
    Dim l As Long
    Dim Q As String
    Q = Chr$(34)

    aFormat(0) = "General"
    aFormat(1) = "0"
    aFormat(2) = "0.00"
    aFormat(3) = "#,##0"
    aFormat(4) = "#,##0.00"
    aFormat(5) = "#,##0\ " & Q & "$" & Q & ";\-#,##0\ " & Q & "$" & Q
    aFormat(6) = "#,##0\ " & Q & "$" & Q & ";[Red]\-#,##0\ " & Q & "$" & Q
    aFormat(7) = "#,##0.00\ " & Q & "$" & Q & ";\-#,##0.00\ " & Q & "$" & Q
    aFormat(8) = "#,##0.00\ " & Q & "$" & Q & ";[Red]\-#,##0.00\ " & Q & "$" & Q
    aFormat(9) = "0%"
    aFormat(10) = "0.00%"
    aFormat(11) = "0.00E+00"
    aFormat(12) = "dd/mm/yy"
    aFormat(13) = "dd/\ mmm\ yy"
    aFormat(14) = "dd/\ mmm"
    aFormat(15) = "mmm\ yy"
    aFormat(16) = "h:mm\ AM/PM"
    aFormat(17) = "h:mm:ss\ AM/PM"
    aFormat(18) = "hh:mm"
    aFormat(19) = "hh:mm:ss"
    aFormat(20) = "dd/mm/yy\ hh:mm"
    aFormat(21) = "##0.0E+0"
    aFormat(22) = "mm:ss"
    aFormat(23) = "@"

    With cFORMAT_COUNT_RECORD
        .Opcode = &H1F
        .Length = &H2
        .Count = CInt(UBound(aFormat))
    End With
    Put #FileNumber, , cFORMAT_COUNT_RECORD

    For lIndex = LBound(aFormat) To UBound(aFormat)
        l = Len(aFormat(lIndex))
        With cFORMAT_RECORD
            .Opcode = &H1E
            .Length = CInt(l + 1)
            .FormatLenght = CInt(l)
        End With
        Put #FileNumber, , cFORMAT_RECORD

        'Then the actual format
        Dim B As Byte, a As Long
        For a = 1 To l
            B = Asc(Mid$(aFormat(lIndex), a, 1))
            Put #FileNumber, , B
        Next
    Next lIndex

    Exit Function

End Function

Function MKI$(x As Integer)
    Dim temp$

    'used for writing integer array values to the disk file
    temp$ = Space$(2)
    CopyMemory ByVal temp$, x%, 2
    MKI$ = temp$
End Function
