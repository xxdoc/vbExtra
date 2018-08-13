VERSION 5.00
Begin VB.UserControl CommonDialogEx 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00D8E9EC&
   MaskPicture     =   "ctlCommonDialogEx.ctx":0000
   Picture         =   "ctlCommonDialogEx.ctx":0E12
   PropertyPages   =   "ctlCommonDialogEx.ctx":1C24
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ctlCommonDialogEx.ctx":1CD9
End
Attribute VB_Name = "CommonDialogEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mDlg As CommonDialogExObject

' Properties

' Returns/sets the path and filename of a selected file.
Public Property Get FileName() As String
Attribute FileName.VB_MemberFlags = "200"
    FileName = mDlg.FileName
End Property

Public Property Let FileName(nValue As String)
    If mDlg.FileName <> nValue Then
        mDlg.FileName = nValue
        PropertyChanged "FileName"
    End If
End Property


' Sets the string displayed in the title bar of the dialog box.
Public Property Get DialogTitle() As String
    DialogTitle = mDlg.DialogTitle
End Property

Public Property Let DialogTitle(nValue As String)
    If mDlg.DialogTitle <> nValue Then
        mDlg.DialogTitle = nValue
        PropertyChanged "DialogTitle"
    End If
End Property


' Returns/sets the filters that are displayed in the Type list box of a dialog box.
Public Property Get Filter() As String
    Filter = mDlg.Filter
End Property

Public Property Let Filter(nValue As String)
    If mDlg.Filter <> nValue Then
        mDlg.Filter = nValue
        PropertyChanged "Filter"
    End If
End Property


' Returns/sets the default filename extension for the dialog box.
Public Property Get DefaultExt() As String
    DefaultExt = mDlg.DefaultExt
End Property

Public Property Let DefaultExt(nValue As String)
    If mDlg.DefaultExt <> nValue Then
        mDlg.DefaultExt = nValue
        PropertyChanged "DefaultExt"
    End If
End Property


' Returns/sets the initial file directory.
Public Property Get InitDir() As String
    InitDir = mDlg.InitDir
End Property

Public Property Let InitDir(nValue As String)
    If mDlg.InitDir <> nValue Then
        mDlg.InitDir = nValue
        PropertyChanged "InitDir"
    End If
End Property


' Returns/sets the selected color.
Public Property Get Color() As OLE_COLOR
    Color = mDlg.Color
End Property

Public Property Let Color(nValue As OLE_COLOR)
    If mDlg.Color <> nValue Then
        mDlg.Color = nValue
        PropertyChanged "Color"
    End If
End Property


' Sets the options for a dialog box.
Public Property Get Flags() As Long
    Flags = mDlg.Flags
End Property

Public Property Let Flags(nValue As Long)
    If mDlg.Flags <> nValue Then
        mDlg.Flags = nValue
        PropertyChanged "Flags"
    End If
End Property


' Specifies the name of the font that appears in each row for the given level.
Public Property Get FontName() As String
    FontName = mDlg.FontName
End Property

Public Property Let FontName(nValue As String)
    If mDlg.FontName <> nValue Then
        mDlg.FontName = nValue
        PropertyChanged "FontName"
    End If
End Property


' Returns/sets bold font styles.
Public Property Get FontBold() As Boolean
    FontBold = mDlg.FontBold
End Property

Public Property Let FontBold(nValue As Boolean)
    If mDlg.FontBold <> nValue Then
        mDlg.FontBold = nValue
        PropertyChanged "FontBold"
    End If
End Property


' Returns/sets italic font styles.
Public Property Get FontItalic() As Boolean
    FontItalic = mDlg.FontItalic
End Property

Public Property Let FontItalic(nValue As Boolean)
    If mDlg.FontItalic <> nValue Then
        mDlg.FontItalic = nValue
        PropertyChanged "FontItalic"
    End If
End Property


' Returns/sets strikethrough font styles.
Public Property Get FontStrikeThru() As Boolean
    FontStrikeThru = mDlg.FontStrikeThru
End Property

Public Property Let FontStrikeThru(nValue As Boolean)
    If mDlg.FontStrikeThru <> nValue Then
        mDlg.FontStrikeThru = nValue
        PropertyChanged "FontStrikeThru"
    End If
End Property


' Returns/sets underline font styles.
Public Property Get FontUnderLine() As Boolean
    FontUnderLine = mDlg.FontUnderLine
End Property

Public Property Let FontUnderLine(nValue As Boolean)
    If mDlg.FontUnderLine <> nValue Then
        mDlg.FontUnderLine = nValue
        PropertyChanged "FontUnderLine"
    End If
End Property


' Returns/sets the value for the first page to be printed.
Public Property Get FromPage() As Integer
    FromPage = mDlg.FromPage
End Property

Public Property Let FromPage(nValue As Integer)
    If mDlg.FromPage <> nValue Then
        mDlg.FromPage = nValue
        PropertyChanged "FromPage"
    End If
End Property


' Returns/sets the value for the first page to be printed.
Public Property Get ToPage() As Integer
    ToPage = mDlg.ToPage
End Property

Public Property Let ToPage(nValue As Integer)
    If mDlg.ToPage <> nValue Then
        mDlg.ToPage = nValue
        PropertyChanged "ToPage"
    End If
End Property


' Sets the smallest allowable font size (Font dialog) or print range (Print dialog).
Public Property Get Min() As Integer
    Min = mDlg.Min
End Property

Public Property Let Min(nValue As Integer)
    If mDlg.Min <> nValue Then
        mDlg.Min = nValue
        PropertyChanged "Min"
    End If
End Property


' Returns/sets the maximum font size (Font dialog) or print range (Print dialog).
Public Property Get Max() As Integer
    Max = mDlg.Max
End Property

Public Property Let Max(nValue As Integer)
    If mDlg.Max <> nValue Then
        mDlg.Max = nValue
        PropertyChanged "Max"
    End If
End Property


' Returns/sets a value that determines the number of copies to be printed.
Public Property Get Copies() As Integer
    Copies = mDlg.Copies
End Property

Public Property Let Copies(nValue As Integer)
    If mDlg.Copies <> nValue Then
        mDlg.Copies = nValue
        PropertyChanged "Copies"
    End If
End Property


' Indicates whether an error is generated when the user chooses the Cancel button.
Public Property Get CancelError() As Boolean
    CancelError = mDlg.CancelError
End Property

Public Property Let CancelError(nValue As Boolean)
    If mDlg.CancelError <> nValue Then
        mDlg.CancelError = nValue
        PropertyChanged "CancelError"
    End If
End Property


' Returns/sets the name of the Help file associated with the project.
Public Property Get HelpFile() As String
    HelpFile = mDlg.HelpFile
End Property

Public Property Let HelpFile(nValue As String)
    If mDlg.HelpFile <> nValue Then
        mDlg.HelpFile = nValue
        PropertyChanged "HelpFile"
    End If
End Property


' Returns/sets the type of online Help requested.
Public Property Get HelpCommand() As Integer
    HelpCommand = mDlg.HelpCommand
End Property

Public Property Let HelpCommand(nValue As Integer)
    If mDlg.HelpCommand <> nValue Then
        mDlg.HelpCommand = nValue
        PropertyChanged "HelpCommand"
    End If
End Property


' Returns/sets the keyword that identifies the requested Help topic.
Public Property Get HelpKey() As String
    HelpKey = mDlg.HelpKey
End Property

Public Property Let HelpKey(nValue As String)
    If mDlg.HelpKey <> nValue Then
        mDlg.HelpKey = nValue
        PropertyChanged "HelpKey"
    End If
End Property


' Determines if user selections in the Print dialog box are used to change the default printer settings.
Public Property Get PrinterDefault() As Boolean  ' added just for compatibility

End Property

Public Property Let PrinterDefault(ByVal nValue As Boolean) ' added just for compatibility

End Property


' Returns/sets a default filter for an Open or Save As dialog box.
Public Property Get FilterIndex() As Integer
    FilterIndex = mDlg.FilterIndex
End Property

Public Property Let FilterIndex(nValue As Integer)
    If mDlg.FilterIndex <> nValue Then
        mDlg.FilterIndex = nValue
        PropertyChanged "FilterIndex"
    End If
End Property


' Returns/sets the context ID of the requested Help topic.
Public Property Get HelpContext() As Long
    HelpContext = mDlg.HelpContext
End Property

Public Property Let HelpContext(nValue As Long)
    If mDlg.HelpContext <> nValue Then
        mDlg.HelpContext = nValue
        PropertyChanged "HelpContext"
    End If
End Property


' Specifies the size (in points) of the font that appears in each row for the given level.
Public Property Get FontSize() As Single
    FontSize = mDlg.FontSize
End Property

Public Property Let FontSize(nValue As Single)
    If mDlg.FontSize <> nValue Then
        mDlg.FontSize = nValue
        PropertyChanged "FontSize"
    End If
End Property


' Sets the type of dialog box to be displayed.
Public Property Let Action(nValue As Integer)
    mDlg.Action = nValue
End Property


' Returns/sets the maximum size of the filename opened using the CommonDialog control.
Public Property Get MaxFileSize() As Integer
    MaxFileSize = mDlg.MaxFileSize
End Property

Public Property Let MaxFileSize(nValue As Integer)
    If mDlg.MaxFileSize <> nValue Then
        mDlg.MaxFileSize = nValue
        PropertyChanged "MaxFileSize"
    End If
End Property


' Returns a handle (from Microsoft Windows) to the object's device context.
Public Property Get hDC() As Long
    hDC = mDlg.hDC
End Property


' Returns/sets the name (without the path) of the file to open or save at run time.
Public Property Get FileTitle() As String
    FileTitle = mDlg.FileTitle
End Property

Public Property Let FileTitle(nValue As String)
    If mDlg.FileTitle <> nValue Then
        mDlg.FileTitle = nValue
        PropertyChanged "FileTitle"
    End If
End Property


' Returns/sets printer paper orientation
Public Property Get Orientation() As cdePageOrientationConstants
    Orientation = mDlg.Orientation
End Property

Public Property Let Orientation(nValue As cdePageOrientationConstants)
    If mDlg.Orientation <> nValue Then
        mDlg.Orientation = nValue
        PropertyChanged "Orientation"
    End If
End Property


' Devuelve el objeto en el que se encuentra este objeto.
Public Property Get Parent() As Object
    Parent = UserControl.Parent
End Property


' Devuelve un objeto de un control.
Public Property Get Object() As Object
    Set Object = Me
End Property



' Methods

' Displays the CommonDialog control's Open dialog box.
Public Sub ShowOpen(Optional ByVal nFlags As cdeCommonDialogExFileFlagsConstants = -1)
    mDlg.ShowOpen (nFlags)
End Sub


' Displays the CommonDialog control's Save As dialog box.
Public Sub ShowSave(Optional ByVal nFlags As cdeCommonDialogExFileFlagsConstants = -1)
    mDlg.ShowSave (nFlags)
End Sub


' Displays the CommonDialog control's Color dialog box.
Public Sub ShowColor(Optional ByVal nFlags As cdeCommonDialogExColorFlagsConstants = -1)
    mDlg.ShowColor (nFlags)
End Sub


' Displays the CommonDialog control's Font dialog box
Public Sub ShowFont(Optional ByVal nFlags As cdeCommonDialogExFontFlagsConstants = -1)
    mDlg.ShowFont (nFlags)
End Sub


' Displays the CommonDialog control's Printer dialog box.
Public Sub ShowPrinter(Optional ByVal nFlags As cdeCommonDialogExPrinterFlagsConstants = -1)
    mDlg.ShowPrinter (nFlags)
End Sub


' Runs Winhelp.EXE and displays the Help file you specify.
Public Sub ShowHelp()
    mDlg.ShowHelp
End Sub


Private Sub UserControl_Initialize()
    Set mDlg = New CommonDialogExObject
End Sub

Private Sub UserControl_InitProperties()
    mDlg.AmbientUserMode = Ambient.UserMode
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mDlg.AmbientUserMode = Ambient.UserMode
    
    ' Used by more than one dialog
    mDlg.Flags = PropBag.ReadProperty("Flags", 0)
    mDlg.Min = PropBag.ReadProperty("Min", 0)
    mDlg.Max = PropBag.ReadProperty("Max", 0)
    mDlg.CancelError = PropBag.ReadProperty("CancelError", False)
    mDlg.DialogTitle = PropBag.ReadProperty("DialogTitle", "")
    mDlg.InitDir = PropBag.ReadProperty("InitDir", "")
    mDlg.Orientation = PropBag.ReadProperty("Orientation", vbPRORPrinterDefault)
    
    ' Color dialog
    mDlg.Color = PropBag.ReadProperty("Color", 0)
    
    ' File dialog (Open / Save as)
    mDlg.FileName = PropBag.ReadProperty("FileName", "")
    mDlg.Filter = PropBag.ReadProperty("Filter", "")
    mDlg.DefaultExt = PropBag.ReadProperty("DefaultExt", "")
    mDlg.MaxFileSize = PropBag.ReadProperty("MaxFileSize", 0)
    mDlg.FilterIndex = PropBag.ReadProperty("FilterIndex", 0)
    
    ' Folder dialog
    mDlg.FolderDialogHeader = PropBag.ReadProperty("FolderDialogHeader", "")
    mDlg.FolderName = PropBag.ReadProperty("FolderName", "")
    mDlg.RootFolder = PropBag.ReadProperty("RootFolder", "")
    
    ' Font dialog
    mDlg.FontName = PropBag.ReadProperty("FontName", "MS Sans Serif")
    mDlg.FontSize = PropBag.ReadProperty("FontSize", 8)
    mDlg.FontBold = PropBag.ReadProperty("FontBold", False)
    mDlg.FontItalic = PropBag.ReadProperty("FontItalic", False)
    mDlg.FontUnderLine = PropBag.ReadProperty("FontUnderLine", False)
    mDlg.FontStrikeThru = PropBag.ReadProperty("FontStrikeThru", False)
    
    ' Help dialog
    mDlg.HelpContext = PropBag.ReadProperty("HelpContext", 0)
    mDlg.HelpCommand = PropBag.ReadProperty("HelpCommand", 0)
    mDlg.HelpKey = PropBag.ReadProperty("HelpKey", "")
    mDlg.HelpFile = PropBag.ReadProperty("HelpFile", "")
    
    ' Page setup dialog
    mDlg.Units = PropBag.ReadProperty("Units", vbMillimeters)
    mDlg.UnitsForUser = PropBag.ReadProperty("Units", cdeMUUserLocale)
    ' Margins
    mDlg.LeftMargin = PropBag.ReadProperty("LeftMargin", cLeftMarginDefault)
    mDlg.RightMargin = PropBag.ReadProperty("RightMargin", cRightMarginDefault)
    mDlg.TopMargin = PropBag.ReadProperty("TopMargin", cTopMarginDefault)
    mDlg.BottomMargin = PropBag.ReadProperty("BottomMargin", cBottomMarginDefault)
    ' Margin limits
    mDlg.MinLeftMargin = PropBag.ReadProperty("MinLeftMargin", 0)
    mDlg.MinRightMargin = PropBag.ReadProperty("MinRightMargin", 0)
    mDlg.MinTopMargin = PropBag.ReadProperty("MinTopMargin", 0)
    mDlg.MinBottomMargin = PropBag.ReadProperty("MinBottomMargin", 0)
    
    ' Print dialog
    mDlg.Copies = PropBag.ReadProperty("Copies", 1)
    mDlg.FromPage = PropBag.ReadProperty("FromPage", 0)
    mDlg.ToPage = PropBag.ReadProperty("ToPage", 0)
    
End Sub

Private Sub UserControl_Resize()
    Dim iH As Long
    Dim iW As Long
    
    If Not Ambient.UserMode Then
        iH = UserControl.ScaleY(UserControl.ScaleHeight, vbTwips, vbPixels)
        iW = UserControl.ScaleX(UserControl.ScaleWidth, vbTwips, vbPixels)
        
        If (iH <> 34) Or (iW <> 34) Then
            If (iH <> 34) Then
                iH = 34
            End If
            If (iW <> 34) Then
                iW = 34
            End If
            UserControl.Size UserControl.ScaleX(iW, vbPixels, vbTwips), UserControl.ScaleY(iH, vbPixels, vbTwips)
        End If
    End If
End Sub

Private Sub UserControl_Terminate()
    Set mDlg = Nothing
End Sub

' Added properties
Public Property Get Canceled() As Boolean
    Canceled = mDlg.Canceled
End Property

Public Sub ShowPageSetup(Optional ByVal nFlags As cdeCommonDialogExPageSetupFlagsConstants = -1)
    mDlg.ShowPageSetup (nFlags)
End Sub


Public Property Get PaperSize() As cdePaperSizeConstants
    PaperSize = mDlg.PaperSize
End Property

Public Property Let PaperSize(nValue As cdePaperSizeConstants)
    If mDlg.PaperSize <> nValue Then
        mDlg.PaperSize = nValue
        PropertyChanged "PaperSize"
    End If
End Property


Public Property Get PrintQuality() As cdePrintQualityConstants
    PrintQuality = mDlg.PrintQuality
End Property

Public Property Let PrintQuality(nValue As cdePrintQualityConstants)
    If mDlg.PrintQuality <> nValue Then
        mDlg.PrintQuality = nValue
        PropertyChanged "PrintQuality"
    End If
End Property


Public Property Get ColorMode() As cdeColorModeConstants
    ColorMode = mDlg.ColorMode
End Property

Public Property Let ColorMode(nValue As cdeColorModeConstants)
    If mDlg.ColorMode <> nValue Then
        mDlg.ColorMode = nValue
        PropertyChanged "ColorMode"
    End If
End Property


Public Property Get DriverName() As String
    DriverName = mDlg.DriverName
End Property


Public Property Get Duplex() As cdeDuplexConstants
    Duplex = mDlg.Duplex
End Property

Public Property Let Duplex(nValue As cdeDuplexConstants)
    If mDlg.Duplex <> nValue Then
        mDlg.Duplex = nValue
        PropertyChanged "Duplex"
    End If
End Property


Public Property Get PaperBin() As cdePaperBinConstants
    PaperBin = mDlg.PaperBin
End Property

Public Property Let PaperBin(nValue As cdePaperBinConstants)
    If mDlg.PaperBin <> nValue Then
        mDlg.PaperBin = nValue
        PropertyChanged "PaperBin"
    End If
End Property


Public Property Get Port() As String
    Port = mDlg.Port
End Property


Public Property Get DeviceName() As String
    DeviceName = mDlg.DeviceName
End Property


Public Property Get LeftMargin() As Single
    LeftMargin = mDlg.LeftMargin
End Property

Public Property Let LeftMargin(nValue As Single)
    If mDlg.LeftMargin <> nValue Then
        mDlg.LeftMargin = nValue
        PropertyChanged "LeftMargin"
    End If
End Property


Public Property Get MinLeftMargin() As Single
    MinLeftMargin = mDlg.MinLeftMargin
End Property

Public Property Let MinLeftMargin(nValue As Single)
    If mDlg.MinLeftMargin <> nValue Then
        mDlg.MinLeftMargin = nValue
        PropertyChanged "MinLeftMargin"
    End If
End Property


Public Property Get RightMargin() As Single
    RightMargin = mDlg.RightMargin
End Property

Public Property Let RightMargin(nValue As Single)
    If mDlg.RightMargin <> nValue Then
        mDlg.RightMargin = nValue
        PropertyChanged "RightMargin"
    End If
End Property


Public Property Get MinRightMargin() As Single
    MinRightMargin = mDlg.MinRightMargin
End Property

Public Property Let MinRightMargin(nValue As Single)
    If mDlg.MinRightMargin <> nValue Then
        mDlg.MinRightMargin = nValue
        PropertyChanged "MinRightMargin"
    End If
End Property


Public Property Get TopMargin() As Single
    TopMargin = mDlg.TopMargin
End Property

Public Property Let TopMargin(nValue As Single)
    If mDlg.TopMargin <> nValue Then
        mDlg.TopMargin = nValue
        PropertyChanged "TopMargin"
    End If
End Property


Public Property Get MinTopMargin() As Single
    MinTopMargin = mDlg.MinTopMargin
End Property

Public Property Let MinTopMargin(nValue As Single)
    If mDlg.MinTopMargin <> nValue Then
        mDlg.MinTopMargin = nValue
        PropertyChanged "MinTopMargin"
    End If
End Property


Public Property Get BottomMargin() As Single
    BottomMargin = mDlg.BottomMargin
End Property

Public Property Let BottomMargin(nValue As Single)
    If mDlg.BottomMargin <> nValue Then
        mDlg.BottomMargin = nValue
        PropertyChanged "BottomMargin"
    End If
End Property


Public Property Get MinBottomMargin() As Single
    MinBottomMargin = mDlg.MinBottomMargin
End Property

Public Property Let MinBottomMargin(nValue As Single)
    If mDlg.MinBottomMargin <> nValue Then
        mDlg.MinBottomMargin = nValue
        PropertyChanged "MinBottomMargin"
    End If
End Property


Public Property Get Units() As cdeUnits
    Units = mDlg.Units
End Property

Public Property Let Units(nValue As cdeUnits)
    If mDlg.Units <> nValue Then
        mDlg.Units = nValue
        PropertyChanged "Units"
    End If
End Property


Public Property Get UnitsForUser() As cdeUnitsForUser
    UnitsForUser = mDlg.UnitsForUser
End Property

Public Property Let UnitsForUser(nValue As cdeUnitsForUser)
    If mDlg.UnitsForUser <> nValue Then
        mDlg.UnitsForUser = nValue
        PropertyChanged "UnitsForUser"
    End If
End Property


Public Sub Reset()
    Set mDlg = New CommonDialogExObject
End Sub


Public Property Get FolderName() As String
    FolderName = mDlg.FolderName
End Property

Public Property Let FolderName(nValue As String)
    If mDlg.FolderName <> nValue Then
        mDlg.FolderName = nValue
        PropertyChanged "FolderName"
    End If
End Property


Public Property Get FolderDisplayName() As String
    FolderDisplayName = mDlg.FolderDisplayName
End Property


Public Property Get RootFolder() As String
    RootFolder = mDlg.RootFolder
End Property

Public Property Let RootFolder(nValue As String)
    If mDlg.RootFolder <> nValue Then
        mDlg.RootFolder = nValue
        PropertyChanged "RootFolder"
    End If
End Property


Public Property Get FolderDialogHeader() As String
    FolderDialogHeader = mDlg.FolderDialogHeader
End Property

Public Property Let FolderDialogHeader(nValue As String)
    If mDlg.FolderDialogHeader <> nValue Then
        mDlg.FolderDialogHeader = nValue
        PropertyChanged "FolderDialogHeader"
    End If
End Property

Public Sub ShowFolder(Optional nFlags As cdeCommonDialogExFolderFlagsConstants = -1)
    mDlg.ShowFolder nFlags
End Sub


Public Property Set Font(ByVal nFont As StdFont)
    Set mDlg.Font = nFont
    PropertyChanged "FontName"
    PropertyChanged "FontSize"
    PropertyChanged "FontBold"
    PropertyChanged "FontItalic"
    PropertyChanged "FontStrikeThru"
    PropertyChanged "FontUnderLine"
End Property

Public Property Get Font() As StdFont
    Set Font = mDlg.Font
End Property


Public Property Get PaperWidth() As Single
    PaperWidth = mDlg.PaperWidth
End Property


Public Property Get PaperHeight() As Single
    PaperHeight = mDlg.PaperHeight
End Property


Public Property Get Changed() As Boolean
    Changed = mDlg.Changed
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    ' Used by more than one dialog
    PropBag.WriteProperty "Flags", mDlg.Flags, 0
    PropBag.WriteProperty "Min", mDlg.Min, 0
    PropBag.WriteProperty "Max", mDlg.Max, 0
    PropBag.WriteProperty "CancelError", mDlg.CancelError, False
    PropBag.WriteProperty "DialogTitle", mDlg.DialogTitle, ""
    PropBag.WriteProperty "InitDir", mDlg.InitDir, ""
    PropBag.WriteProperty "Orientation", mDlg.Orientation, vbPRORPrinterDefault
    
    ' Color dialog
    PropBag.WriteProperty "Color", mDlg.Color, 0
    
    ' File dialog (Open / Save as)
    PropBag.WriteProperty "FileName", mDlg.FileName, ""
    PropBag.WriteProperty "Filter", mDlg.Filter, ""
    PropBag.WriteProperty "DefaultExt", mDlg.DefaultExt, ""
    PropBag.WriteProperty "MaxFileSize", mDlg.MaxFileSize, 0
    PropBag.WriteProperty "FilterIndex", mDlg.FilterIndex, 0
    
    ' Folder dialog
    PropBag.WriteProperty "FolderDialogHeader", mDlg.FolderDialogHeader, ""
    PropBag.WriteProperty "FolderName", mDlg.FolderName, ""
    PropBag.WriteProperty "RootFolder", mDlg.RootFolder, ""
    
    ' Font dialog
    PropBag.WriteProperty "FontName", mDlg.FontName, "MS Sans Serif"
    PropBag.WriteProperty "FontSize", mDlg.FontSize, 8
    PropBag.WriteProperty "FontBold", mDlg.FontBold, False
    PropBag.WriteProperty "FontItalic", mDlg.FontItalic, False
    PropBag.WriteProperty "FontUnderLine", mDlg.FontUnderLine, False
    PropBag.WriteProperty "FontStrikeThru", mDlg.FontStrikeThru, False
    
    ' Help dialog
    PropBag.WriteProperty "HelpContext", mDlg.HelpContext, 0
    PropBag.WriteProperty "HelpCommand", mDlg.HelpCommand, 0
    PropBag.WriteProperty "HelpKey", mDlg.HelpKey, ""
    PropBag.WriteProperty "HelpFile", mDlg.HelpFile, ""
    
    ' Page setup dialog
    PropBag.WriteProperty "Units", mDlg.Units, vbMillimeters
    PropBag.WriteProperty "Units", mDlg.UnitsForUser, cdeMUUserLocale
    ' Margins
    PropBag.WriteProperty "LeftMargin", mDlg.LeftMargin, cLeftMarginDefault
    PropBag.WriteProperty "RightMargin", mDlg.RightMargin, cRightMarginDefault
    PropBag.WriteProperty "TopMargin", mDlg.TopMargin, cTopMarginDefault
    PropBag.WriteProperty "BottomMargin", mDlg.BottomMargin, cBottomMarginDefault
    ' Margin limits
    PropBag.WriteProperty "MinLeftMargin", mDlg.MinLeftMargin, 0
    PropBag.WriteProperty "MinRightMargin", mDlg.MinRightMargin, 0
    PropBag.WriteProperty "MinTopMargin", mDlg.MinTopMargin, 0
    PropBag.WriteProperty "MinBottomMargin", mDlg.MinBottomMargin, 0
    
    ' Print dialog
    PropBag.WriteProperty "Copies", mDlg.Copies, 1
    PropBag.WriteProperty "FromPage", mDlg.FromPage, 0
    PropBag.WriteProperty "ToPage", mDlg.ToPage, 0
    
End Sub

