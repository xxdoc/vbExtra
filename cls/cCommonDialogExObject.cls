VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonDialogExObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function lstrlen Lib "Kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long

Private Type PRINTER_DEFAULTS
    pDatatype As Long
    pDevmode As Long
    DesiredAccess As Long
End Type

'Private Const DM_IN_BUFFER = 8
Private Const DM_OUT_BUFFER = 2
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, ByVal pDevModeOutput As Long, ByVal pDevModeInput As Long, ByVal fMode As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Private Declare Sub CopyMemoryAnyLong Lib "Kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal cbLength As Long)

Const DC_PAPERS = 2
'Const DC_PAPERNAMES = 16
Const DC_PAPERSIZE = 3

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long

'Private Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" (ByVal hDC As Long, lpInitData As Any) As Long
Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, lpDevMode As Any) As Long

Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Private Declare Function GetLocaleInfo Lib "Kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Const LOCALE_USER_DEFAULT As Long = &H400&
Private Const LOCALE_IMEASURE As Long = &HD&

'API memory functions
Private Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "Kernel32" (ByVal hMem As Long) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function lstrlenPtr Lib "Kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SHGetFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, pidl As ITEMIDLIST) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function SHParseDisplayName Lib "shell32.dll" (ByVal pszName As Long, ByVal pbc As Long, ByRef ppidl As Long, ByVal sfgaoIn As Long, ByRef psfgaoOut As Long) As Long

Private Const S_OK As Long = 0&
Private Const CSIDL_PERSONAL As Long = &H5
Private Const CSIDL_DRIVES As Long = &H11
Private Const CSIDL_PRINTERS As Long = &H4

'constants for API memory functions
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

'data buffer for the ChooseColor function
Private Type T_CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    RGBResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'constants for LOGFONT
Private Const LF_FACESIZE As Long = 32
'Private Const LF_FULLFACESIZE As Long = 64

'Private Const FW_THIN As Long = 100
'Private Const FW_EXTRALIGHT As Long = 200
'Private Const FW_LIGHT As Long = 300
Private Const FW_NORMAL As Long = 400
'Private Const FW_MEDIUM As Long = 500
'Private Const FW_SEMIBOLD As Long = 600
Private Const FW_BOLD As Long = 700
'Private Const FW_EXTRABOLD As Long = 800
'Private Const FW_HEAVY As Long = 900

'data buffer for the ChooseFont function
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(0 To LF_FACESIZE - 1) As Byte
End Type

'data buffer for the ChooseFont function
Private Type T_CHOOSEFONT
    lStructSize As Long
    hwndOwner As Long
    hDC As Long
    lpLogFont As Long
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

'data buffer for the GetOpenFileName and GetSaveFileName functions
Private Type T_OPENFILENAME
'    lStructSize As Long
'    hWndOwner As Long
'    hInstance As Long
'    lpstrFilter As String
'    lpstrCustomFilter As String
'    nMaxCustFilter As Long
'    mFilterIndex As Long
'    lpstrFile As String
'    nMaxFile As Long
'    lpstrFileTitle As String
'    nMaxFileTitle As Long
'    lpstrInitialDir As String
'    lpstrTitle As String
'    Flags As Long
'    nFileOffset As Integer
'    nFileExtension As Integer
'    lpstrDefExt As String
'    lCustData As Long
'    lpfnHook As Long
'    lpTemplateName As String
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As Long
    lpstrCustomFilter As Long
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As Long
    nMaxFile As Long
    lpstrFileTitle As Long
    nMaxFileTitle As Long
    lpstrInitialDir As Long
    lpstrTitle As Long
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
    pvReserved As Long
    dwReserved As Long
    FlagsEx As Long
End Type


'data buffer for the PrintDlg function
Private Type T_PrintDlg
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    Flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type


Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As T_CHOOSECOLOR) As Long
'Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As T_CHOOSEFONT) As Long
Private Declare Function ChooseFont Lib "comdlg32" Alias "ChooseFontW" (ByRef pChooseFont As T_CHOOSEFONT) As Long
'Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Private Declare Function WinHelp Lib "user32" Alias "WinHelpW" (ByVal hWnd As Long, ByVal lpHelpFile As Long, ByVal wCommand As Long, ByVal dwData As Long) As Long
'Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As T_OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameW" (pOpenfilename As T_OPENFILENAME) As Long
'Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As T_OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameW" (pOpenfilename As T_OPENFILENAME) As Long
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As T_PrintDlg) As Long
'Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgW" (pPrintdlg As T_PrintDlg) As Long ' it seems to work fine, but not sure what the advantage is using it

'API function to retrieve extended error information
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

Private Type T_PAGESETUPDLG
    lStructSize                 As Long
    hwndOwner                   As Long
    hDevMode                    As Long
    hDevNames                   As Long
    Flags                       As Long
    ptPaperSize                 As POINTAPI
    rtMinMargin                 As RECT
    rtMargin                    As RECT
    hInstance                   As Long
    lCustData                   As Long
    lpfnPageSetupHook           As Long
    lpfnPagePaintHook           As Long
    lpPageSetupTemplateName     As Long
    hPageSetupTemplate          As Long
End Type

Private Declare Function PageSetupDlg Lib "comdlg32" Alias "PageSetupDlgA" (lppage As T_PAGESETUPDLG) As Boolean
    
Private Const CC_ENABLEHOOK = &H10&
Private Const CC_ENABLETEMPLATE = &H20&
Private Const CC_RGBINIT = &H1&

Private Const CF_PRINTERFONTS As Long = &H2
Private Const CF_SCREENFONTS As Long = &H1
Private Const CF_INITTOLOGFONTSTRUCT As Long = &H40&
'Private Const CF_BOTH As Long = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_EFFECTS As Long = &H100
Private Const CF_APPLY As Long = &H200
Private Const CF_ENABLEHOOK As Long = &H8
Private Const CF_ENABLETEMPLATE As Long = &H10
Private Const CF_LIMITSIZE As Long = &H2000

Private Const PD_ENABLEPRINTHOOK = &H1000&
Private Const PD_ENABLESETUPHOOK = &H2000&
Private Const PD_ENABLEPRINTTEMPLATE = &H4000&
Private Const PD_ENABLESETUPTEMPLATE = &H8000&

Private Const PSD_ENABLEPAGEPAINTHOOK As Long = &H40000
Private Const PSD_ENABLEPAGESETUPHOOK As Long = &H2000&
Private Const PSD_ENABLEPAGESETUPTEMPLATE As Long = &H8000&


'constants for color dialog

'Private Const CDERR_DIALOGFAILURE = &HFFFF
'Private Const CDERR_FINDRESFAILURE = &H6
'Private Const CDERR_GENERALCODES = &H0
'Private Const CDERR_INITIALIZATION = &H2
'Private Const CDERR_LOADRESFAILURE = &H7
'Private Const CDERR_LOADSTRFAILURE = &H5
'Private Const CDERR_LOCKRESFAILURE = &H8
'Private Const CDERR_MEMALLOCFAILURE = &H9
'Private Const CDERR_MEMLOCKFAILURE = &HA
'Private Const CDERR_NOHINSTANCE = &H4
'Private Const CDERR_NOHOOK = &HB
'Private Const CDERR_NOTEMPLATE = &H3
'Private Const CDERR_REGISTERMSGFAIL = &HC
'Private Const CDERR_STRUCTSIZE = &H1


'constants for file dialog

'Private Const FNERR_BUFFERTOOSMALL As Long = &H3003
'Private Const FNERR_FILENAMECODES As Long = &H3000
'Private Const FNERR_INVALIDFILENAME As Long = &H3002
'Private Const FNERR_SUBCLASSFAILURE As Long = &H3001

' other constants
Private Const CLEARTYPE_QUALITY As Byte = 6
Private Const DEFAULT_CHARSET As Byte = 1

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Const PSD_MARGINS As Long = &H2&
Private Const PSD_INHUNDREDTHSOFMILLIMETERS = &H8&
Private Const PSD_INTHOUSANDTHSOFINCHES = &H4&

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
End Type


Public Enum CommonDialogsFlags

' File Open, Save as
    cdlOFNAllowMultiselect = &H200&
    cdlOFNCreatePrompt = &H2000&
    cdlOFNExplorer = &H80000
    cdlOFNExtensionDifferent = &H400&
    cdlOFNFileMustExist = &H1000&
    cdlOFNHelpButton = &H10&
    cdlOFNHideReadOnly = &H4&
    cdlOFNLongNames = &H200000
    cdlOFNNoChangeDir = &H8&
    cdlOFNNoDereferenceLinks = &H100000
    cdlOFNNoLongNames = &H40000
    cdlOFNNoReadOnlyReturn = &H8000&
    cdlOFNNoValidate = &H100&
    cdlOFNOverwritePrompt = &H2&
    cdlOFNPathMustExist = &H800&
    cdlOFNReadOnly = &H1&
    cdlOFNShareAware = &H4000&
 
'PrintFlags
    cdlPDAllPages = &H0&
    cdlPDSelection = &H1&
    cdlPDPageNums = &H2&
    cdlPDNoSelection = &H4&
    cdlPDNoPageNums = &H8&
    cdlPDCollate = &H10&
    cdlPDPrintToFile = &H20&
    cdlPDPrintSetup = &H40&
    cdlPDNoWarning = &H80&
    cdlPDReturnDC = &H100&
    cdlPDReturnIC = &H200&
    cdlPDReturnDefault = &H400&
    cdlPDHelpButton = &H800&
    cdlPDUseDevModeCopies = &H40000
    cdlPDDisablePrintToFile = &H80000
    cdlPDHidePrintToFile = &H100000
 
'ColorFlags
    cdlCCFullOpen = &H2&
    cdlCCShowHelpButton = &H8&
    cdlCCPreventFullOpen = &H4&
    cdlCCRGBInit = &H1&
 
'FontFlags
    cdlCFANSIOnly = &H400&
    cdlCFApply = &H200&
    cdlCFBoth = &H3&
    cdlCFEffects = &H100&
    cdlCFFixedPitchOnly = &H4000&
    cdlCFForceFontExist = &H10000
    cdlCFHelpButton = &H4&
    cdlCFLimitSize = &H2000&
    cdlCFNoFaceSel = &H80000
    cdlCFNoSimulations = &H1000&
    cdlCFNoSizeSel = &H200000
    cdlCFNoStyleSel = &H100000
    cdlCFNoVectorFonts = &H800&
    cdlCFPrinterFonts = &H2&
    cdlCFScalableOnly = &H20000
    cdlCFScreenFonts = &H1&
    cdlCFTTOnly = &H40000
    cdlCFWYSIWYG = &H8000&
    
' Page setup Flags
    cdlPSDefaultMinMargins = &H0&
    cdlPSDisableMargins = &H10&
    cdlPSDisableOrientation = &H100&
    cdlPSDisablePagePainting = &H80000
    cdlPSDisablePaper = &H200&
    cdlPSDisablePrinter = &H20&
    'cdlPSMargins = &H2&
    cdlPSMinMargins = &H1&
    cdlPSNoNetworkButton = &H200000
    cdlPSNoWarning = &H80&
    cdlPSReturnDefault = &H400&
    cdlPSShowHelp = &H800&
    
' Show Folder
    cdlSFReturnOnlyFSDirs = &H1&
    cdlSFDontGoBelowDomain = &H2&
    cdlSFStatusText = &H4&
    cdlSFReturnFSAncestors = &H8&
    cdlSFEditBox = &H10&
    cdlSFValidate = &H20&
    cdlSFNewDialogStyle = &H40&
    cdlSFUseNewUI = (cdlSFNewDialogStyle Or cdlSFEditBox)
    cdlSFBrowseIncludeURLs = &H80&
    cdlSFUAHint = &H100&
    cdlSFNoNewFolderButton = &H200&
    cdlSFNoTranslateTargets = &H400&
    cdlSFNoReturnOnlyFSDirs = &H800&
    cdlSFBrowseForComputer = &H1000&
    cdlSFBrowseForPrinter = &H2000&
    cdlSFBrowseIncludeFiles = &H4000&
    cdlSFShareable = &H8000&
End Enum

Public Enum cdeCommonDialogExFileFlagsConstants
    cdeOFNAllowMultiselect = CommonDialogsFlags.cdlOFNAllowMultiselect
    cdeOFNCreatePrompt = CommonDialogsFlags.cdlOFNCreatePrompt
    cdeOFNExplorer = CommonDialogsFlags.cdlOFNExplorer
    cdeOFNExtensionDifferent = CommonDialogsFlags.cdlOFNExtensionDifferent
    cdeOFNFileMustExist = CommonDialogsFlags.cdlOFNFileMustExist
    cdeOFNHelpButton = CommonDialogsFlags.cdlOFNHelpButton
    cdeOFNHideReadOnly = CommonDialogsFlags.cdlOFNHideReadOnly
    cdeOFNLongNames = CommonDialogsFlags.cdlOFNLongNames
    cdeOFNNoChangeDir = CommonDialogsFlags.cdlOFNNoChangeDir
    cdeOFNNoDereferenceLinks = CommonDialogsFlags.cdlOFNNoDereferenceLinks
    cdeOFNNoLongNames = CommonDialogsFlags.cdlOFNNoLongNames
    cdeOFNNoReadOnlyReturn = CommonDialogsFlags.cdlOFNNoReadOnlyReturn
    cdeOFNNoValidate = CommonDialogsFlags.cdlOFNNoValidate
    cdeOFNOverwritePrompt = CommonDialogsFlags.cdlOFNOverwritePrompt
    cdeOFNPathMustExist = CommonDialogsFlags.cdlOFNPathMustExist
    cdeOFNReadOnly = CommonDialogsFlags.cdlOFNReadOnly
    cdeOFNShareAware = CommonDialogsFlags.cdlOFNShareAware
End Enum

Public Enum cdeCommonDialogExPrinterFlagsConstants
    cdePDAllPages = CommonDialogsFlags.cdlPDAllPages
    cdePDSelection = CommonDialogsFlags.cdlPDSelection
    cdePDPageNums = CommonDialogsFlags.cdlPDPageNums
    cdePDNoSelection = CommonDialogsFlags.cdlPDNoSelection
    cdePDNoPageNums = CommonDialogsFlags.cdlPDNoPageNums
    cdePDCollate = CommonDialogsFlags.cdlPDCollate
    cdePDPrintToFile = CommonDialogsFlags.cdlPDPrintToFile
    cdePDPrintSetup = CommonDialogsFlags.cdlPDPrintSetup
    cdePDNoWarning = CommonDialogsFlags.cdlPDNoWarning
    cdePDReturnDC = CommonDialogsFlags.cdlPDReturnDC
    cdePDReturnIC = CommonDialogsFlags.cdlPDReturnIC
    cdePDReturnDefault = CommonDialogsFlags.cdlPDReturnDefault
    cdePDHelpButton = CommonDialogsFlags.cdlPDHelpButton
    cdePDUseDevModeCopies = CommonDialogsFlags.cdlPDUseDevModeCopies
    cdePDDisablePrintToFile = CommonDialogsFlags.cdlPDDisablePrintToFile
    cdePDHidePrintToFile = CommonDialogsFlags.cdlPDHidePrintToFile
End Enum

Public Enum cdeCommonDialogExColorFlagsConstants
    cdeCCFullOpen = CommonDialogsFlags.cdlCCFullOpen
    cdeCCShowHelpButton = CommonDialogsFlags.cdlCCShowHelpButton
    cdeCCPreventFullOpen = CommonDialogsFlags.cdlCCPreventFullOpen
'    cdeCCRGBInit = CommonDialogsFlags.cdlCCRGBInit
End Enum

Public Enum cdeCommonDialogExFontFlagsConstants
    cdeCFANSIOnly = CommonDialogsFlags.cdlCFANSIOnly
    cdeCFApply = CommonDialogsFlags.cdlCFApply
    cdeCFBoth = CommonDialogsFlags.cdlCFBoth
    cdeCFEffects = CommonDialogsFlags.cdlCFEffects
    cdeCFFixedPitchOnly = CommonDialogsFlags.cdlCFFixedPitchOnly
    cdeCFForceFontExist = CommonDialogsFlags.cdlCFForceFontExist
    cdeCFHelpButton = CommonDialogsFlags.cdlCFHelpButton
    cdeCFLimitSize = CommonDialogsFlags.cdlCFLimitSize
    cdeCFNoFaceSel = CommonDialogsFlags.cdlCFNoFaceSel
    cdeCFNoSimulations = CommonDialogsFlags.cdlCFNoSimulations
    cdeCFNoSizeSel = CommonDialogsFlags.cdlCFNoSizeSel
    cdeCFNoStyleSel = CommonDialogsFlags.cdlCFNoStyleSel
    cdeCFNoVectorFonts = CommonDialogsFlags.cdlCFNoVectorFonts
    cdeCFPrinterFonts = CommonDialogsFlags.cdlCFPrinterFonts
    cdeCFScalableOnly = CommonDialogsFlags.cdlCFScalableOnly
    cdeCFScreenFonts = CommonDialogsFlags.cdlCFScreenFonts
    cdeCFTTOnly = CommonDialogsFlags.cdlCFTTOnly
    cdeCFWYSIWYG = CommonDialogsFlags.cdlCFWYSIWYG
End Enum

Public Enum cdeCommonDialogExPageSetupFlagsConstants
    cdePSDefaultMinMargins = CommonDialogsFlags.cdlPSDefaultMinMargins
    cdePSDisableMargins = CommonDialogsFlags.cdlPSDisableMargins
    cdePSDisableOrientation = CommonDialogsFlags.cdlPSDisableOrientation
    cdePSDisablePagePainting = CommonDialogsFlags.cdlPSDisablePagePainting
    cdePSDisablePaper = CommonDialogsFlags.cdlPSDisablePaper
    cdePSDisablePrinter = CommonDialogsFlags.cdlPSDisablePrinter
'    cdePSMargins = CommonDialogsFlags.cdlPSMargins
    cdePSMinMargins = CommonDialogsFlags.cdlPSMinMargins
    cdePSNoNetworkButton = CommonDialogsFlags.cdlPSNoNetworkButton
    cdePSNoWarning = CommonDialogsFlags.cdlPSNoWarning
    cdePSReturnDefault = CommonDialogsFlags.cdlPSReturnDefault
    cdePSShowHelp = CommonDialogsFlags.cdlPSShowHelp
End Enum

Public Enum cdeCommonDialogExFolderFlagsConstants
    cdeSFReturnOnlyFSDirs = CommonDialogsFlags.cdlSFReturnOnlyFSDirs
    cdeSFDontGoBelowDomain = CommonDialogsFlags.cdlSFDontGoBelowDomain
    cdeSFStatusText = CommonDialogsFlags.cdlSFStatusText
    cdeSFReturnFSAncestors = CommonDialogsFlags.cdlSFReturnFSAncestors
    cdeSFEditBox = CommonDialogsFlags.cdlSFEditBox
    cdeSFValidate = CommonDialogsFlags.cdlSFValidate
    cdeSFNewDialogStyle = CommonDialogsFlags.cdlSFNewDialogStyle
    cdeSFUseNewUI = CommonDialogsFlags.cdlSFUseNewUI
    cdeSFBrowseIncludeURLs = CommonDialogsFlags.cdlSFBrowseIncludeURLs
    cdeSFUAHint = CommonDialogsFlags.cdlSFUAHint
    cdeSFNoNewFolderButton = CommonDialogsFlags.cdlSFNoNewFolderButton
    cdeSFNoTranslateTargets = CommonDialogsFlags.cdlSFNoTranslateTargets
    cdeSFNoReturnOnlyFSDirs = CommonDialogsFlags.cdlSFNoReturnOnlyFSDirs
    cdeSFBrowseForComputer = CommonDialogsFlags.cdlSFBrowseForComputer
    cdeSFBrowseForPrinter = CommonDialogsFlags.cdlSFBrowseForPrinter
    cdeSFBrowseIncludeFiles = CommonDialogsFlags.cdlSFBrowseIncludeFiles
    cdeSFShareable = CommonDialogsFlags.cdlSFShareable
    
End Enum


' Other Public enumerations
Public Enum cdeColorModeConstants
    vbPRCMPrinterDefault = 0&
    vbPRCMMonochrome = 1&
    vbPRCMColor = 2&
End Enum

Public Enum cdePaperBinConstants
    vbPRBNPrinterDefault = 0&
    vbPRBNUpper = 1&
    vbPRBNLower = 2&
    vbPRBNMiddle = 3&
    vbPRBNManual = 4&
    vbPRBNEnvelope = 5&
    vbPRBNEnvManual = 6&
    vbPRBNAuto = 7&
    vbPRBNTractor = 8&
    vbPRBNSmallFmt = 9&
    vbPRBNLargeFmt = 10&
    vbPRBNLargeCapacity = 11&
    vbPRBNCassette = 14&
End Enum

Public Enum cdePaperSizeConstants
    vbPRPSPrinterDefault = 0&
    vbPRPSLetter = 1&
    vbPRPSLetterSmall = 2&
    vbPRPSTabloid = 3&
    vbPRPSLedger = 4&
    vbPRPSLegal = 5&
    vbPRPSStatement = 6&
    vbPRPSExecutive = 7&
    vbPRPSA3 = 8&
    vbPRPSA4 = 9&
    vbPRPSA4Small = 10&
    vbPRPSA5 = 11&
    vbPRPSB4 = 12&
    vbPRPSB5 = 13&
    vbPRPSFolio = 14&
    vbPRPSQuarto = 15&
    vbPRPS10x14 = 16&
    vbPRPS11x17 = 17&
    vbPRPSNote = 18&
    vbPRPSEnv9 = 19&
    vbPRPSEnv10 = 20&
    vbPRPSEnv11 = 21&
    vbPRPSEnv12 = 22&
    vbPRPSEnv14 = 23&
    vbPRPSCSheet = 24&
    vbPRPSDSheet = 25&
    vbPRPSESheet = 26&
    vbPRPSEnvDL = 27&
    vbPRPSEnvC3 = 29&
    vbPRPSEnvC4 = 30&
    vbPRPSEnvC5 = 28&
    vbPRPSEnvC6 = 31&
    vbPRPSEnvC65 = 32&
    vbPRPSEnvB4 = 33&
    vbPRPSEnvB5 = 34&
    vbPRPSEnvB6 = 35&
    vbPRPSEnvItaly = 36&
    vbPRPSEnvMonarch = 37&
    vbPRPSEnvPersonal = 38&
    vbPRPSFanfoldUS = 39&
    vbPRPSFanfoldStdGerman = 40&
    vbPRPSFanfoldLglGerman = 41&
    vbPRPSUser = 256&
End Enum

Public Enum cdePrintQualityConstants
    vbPRPQPrinterDefault = 0&
    vbPRPQDraft = -1&
    vbPRPQLow = -2&
    vbPRPQMedium = -3&
    vbPRPQHigh = -4&
End Enum

Public Enum cdeDuplexConstants
    vbPRDPPrinterDefault = 0&
    vbPRDPSimplex = 1&
    vbPRDPHorizontal = 2&
    vbPRDPVertical = 3&
End Enum

Public Enum cdePageOrientationConstants
    vbPRORPrinterDefault = 0&
    vbPRORPortrait = 1&
    vbPRORLandscape = 2&
End Enum

Public Enum cdeUnits
    vbMillimeters = 6&
    vbInches = 5&
End Enum

Public Enum cdeUnitsForUser
    cdeMUUserLocale = 0&
    cdeMUMillimeters = 6&
    cdeMUInches = 5&
End Enum

'Properties

Private mAction As Integer
Private mCancelError As Boolean
Private mColor As Long
Private mCopies As Long
Private mCollate As Boolean
Private mDefaultExt As String
Private mDialogTitle As String
Private mFileName As String
Private mFileTitle As String
Private mFilter As String
Private mFilterIndex As Integer
Private mFlags As Long
Private mFont As StdFont
Private mFontBold As Boolean
Private mFontItalic As Boolean
Private mFontName As String
Private mFontNameSet As Boolean
Private mFontSize As Long
Private mFontStrikethrough As Boolean
Private mFontUnderline As Boolean
Private mFromPage As Long
Private mhDc As Long
Private mHelpCommand As Long
Private mHelpContext As Long
Private mHelpFile As String
Private mHelpKey As String
Private mInitDir As String
Private mMax As Long
Private mMaxFileSize As Long
Private mMin As Long
'Private mPrinterDefault As Boolean
Private mToPage As Long
Private mOrientation As Long
Private mCustomColors(0 To 15) As Long
Private mAutoSaveCustomColors As Boolean
Private mCustomColorsLoaded As Boolean

Private mApiReturn As Long
Private mExtendedError As Long
Private mCanceled As Boolean

' Added properties special for printer and page setup
Private mPaperSize As Long
Private mPaperBin As Long
Private mPrintQuality As Long
Private mColorMode As Long
Private mDuplex As Long
' read only
Private mDeviceName As String
Private mDriverName As String
Private mPort As String
Private mPaperWidth As Long
Private mPaperHeight As Long
Private mDefaultPaperWidth As Long
Private mDefaultPaperHeight As Long

' added properties special for page setup
Private mLeftMargin As Single
Private mRightMargin As Single
Private mTopMargin As Single
Private mBottomMargin As Single
Private mMinLeftMargin As Single
Private mMinRightMargin As Single
Private mMinTopMargin As Single
Private mMinBottomMargin As Single
Private mUnits As Long
Private mUnitsForUser As Long
Private mChanged As Boolean

' added property for show folder
Private mFolderName As String
Private mFolderDisplayName As String
Private mRootFolder As String
Private mFolderDialogHeader As String

' auxiliary variables
Private mDevMode As DEVMODE
Private mDevNames As DEVNAMES
Private mDevModePtr As Long
Private mDevNamesPtr As Long
Private mhDevNames As Long
Private mhDevMode As Long
Private mhDevModeToFree As Long
Private mPageSet As Boolean
Private mManualTermination As Boolean
Private mStandardPaperSizesLoaded As Boolean
Private mStandardPaperSizesNumbers() As Long
Private mStandardPaperSizes() As POINTAPI

Private mAmbientUserMode As Boolean

Public Sub ShowColor(Optional ByVal nFlags As cdeCommonDialogExColorFlagsConstants = -1)
    'display the color dialog box
    
    Dim tChooseColor As T_CHOOSECOLOR
    Dim lCustomColorSize As Long
    Dim lCustomColorAddress As Long
    Dim lMemHandle As Long
    Dim iHwndOwner As Long
    Dim iFlags As Long
    Dim iColor_Prev As Long
    
    On Error GoTo ShowColorError
    
    iHwndOwner = GetActiveWindowHwnd
    iFlags = mFlags
    If nFlags <> -1 Then
        iFlags = iFlags Or nFlags
    End If
    
    mAction = 3  'Action property - ShowColor
    mApiReturn = 0  'APIReturn property
    mExtendedError = 0  'ExtendedError property
    
    If mAutoSaveCustomColors Then
        If Not mCustomColorsLoaded Then LoadCustomColors
    End If
    
    tChooseColor.lStructSize = Len(tChooseColor)
    
    tChooseColor.hwndOwner = iHwndOwner
    
    ' Get size of memory needed for custom colors
    lCustomColorSize = Len(mCustomColors(0)) * 16
    ' Get a global memory block to hold a copy of the custom colors
    lMemHandle = GlobalAlloc(GHND, lCustomColorSize)
    
    If lMemHandle = 0 Then
        Exit Sub
    End If
    ' Lock the custom color's global memory block
    lCustomColorAddress = GlobalLock(lMemHandle)
    If lCustomColorAddress = 0 Then
        Exit Sub
    End If
    ' Copy custom colors to the global memory block
    CopyMemory ByVal lCustomColorAddress, mCustomColors(0), lCustomColorSize
 
    tChooseColor.lpCustColors = lCustomColorAddress
    
    tChooseColor.Flags = iFlags And Not (CC_ENABLEHOOK Or CC_ENABLETEMPLATE) Or CC_RGBINIT
        
    tChooseColor.RGBResult = TranslateAColor(mColor)
    
    iColor_Prev = tChooseColor.RGBResult
    mChanged = False
    
    '***    call the ChooseColor API function
    mApiReturn = ChooseColor(tChooseColor)
    
    '***    handle return from ChooseColor API function
    mCanceled = False
    Select Case mApiReturn
        
        Case 0  'user canceled
        mCanceled = True
        If mCancelError Then
            'generate an error
            On Error GoTo 0
            Err.Raise 32755, TypeName(Me), "Cancel Pressed"
            Exit Sub
        End If
        
        Case 1  'user selected a color
            'update property buffer
            mColor = tChooseColor.RGBResult
            mChanged = (tChooseColor.RGBResult <> iColor_Prev)
            CopyMemory mCustomColors(0), ByVal lCustomColorAddress, lCustomColorSize
            If mAutoSaveCustomColors Then
                SaveCustomColors
            End If
        
        Case Else   'an error occured
            'call CommDlgExtendedError
            mExtendedError = CommDlgExtendedError
        
    End Select
    
    GlobalFree lMemHandle
    
Exit Sub

ShowColorError:
    Exit Sub
End Sub

Public Sub ShowFont(Optional ByVal nFlags As cdeCommonDialogExFontFlagsConstants = -1)
    'display the font dialog box
    
    Dim tLogFont As LOGFONTW
    Dim iCF As T_CHOOSEFONT
    Dim lLogFontSize As Long
    Dim lLogFontAddress As Long
    Dim lMemHandle As Long
    Dim i As Long
    Dim iHwndOwner As Long
    Dim iFlags As Long
    Dim iBytes() As Byte
    Dim iStr As String
    Dim iLng As Long
    
    Dim iFontName_Prev As String
    Dim iFontSize_Prev As Single
    Dim iFontBold_prev As Boolean
    Dim iFontItalic_Prev As Boolean
    Dim iFontStrikethrough_Prev As Boolean
    Dim iFontUnderline_Prev As Boolean
    Dim iFontWidth_Prev As Long
    
    mChanged = False
    iFontName_Prev = mFont.Name
    iFontSize_Prev = mFont.Size
    iFontBold_prev = mFont.Bold
    iFontItalic_Prev = mFont.Italic
    iFontStrikethrough_Prev = mFont.Strikethrough
    iFontUnderline_Prev = mFont.Underline
    iFontWidth_Prev = mFont.Weight
    
    On Error GoTo ShowFontError
    
    iHwndOwner = GetActiveWindowHwnd
    iFlags = mFlags
    If nFlags <> -1 Then
        iFlags = iFlags Or nFlags
    End If
    
    '***    init property buffers
    
    mAction = 4  'Action property - ShowFont
    mApiReturn = 0  'APIReturn property
    mExtendedError = 0  'ExtendedError property

    
    '***    prepare T_CHOOSEFONT data
    
    tLogFont.lfHeight = mFont.Size * 20 / Screen.TwipsPerPixelY * -1 ' PointsPerTwip = 1440 / 72 = 20
    'tLogFont.lfWidth As Long
    'tLogFont.lfEscapement As Long
    'tLogFont.lfOrientation As Long
    
    'tLogFont.lfWeight As Long - init from FontBold property
    If mFont.Bold Then
        tLogFont.lfWeight = FW_BOLD
    Else
        tLogFont.lfWeight = FW_NORMAL
    End If
    
    'tLogFont.lfItalic As Byte - init from FontItalic property
    If mFont.Italic Then
        tLogFont.lfItalic = 1
    End If
    
    'tLogFont.lfUnderline As Byte - init from FontUnderline property
    If mFont.Underline Then
        tLogFont.lfUnderline = 1
    End If

    'tLogFont.lfStrikeOut As Byte - init from FontStrikethru property
    If mFont.Strikethrough Then
        tLogFont.lfStrikeOut = 1
    End If


    tLogFont.lfCharSet = DEFAULT_CHARSET
    'tLogFont.lfOutPrecision As Byte
    'tLogFont.lfClipPrecision As Byte
    tLogFont.lfQuality = CLEARTYPE_QUALITY
    'tLogFont.lfPitchAndFamily As Byte
    
    iBytes = mFont.Name
    For i = 0 To Len(mFont.Name) - 1
        If i > 31 Then Exit For
        tLogFont.lfFaceName(i * 2) = iBytes(i * 2) 'Asc(Mid$(mFont.Name, i + 1, 1))
        tLogFont.lfFaceName(i * 2 + 1) = iBytes(i * 2 + 1)
    Next
    
    lLogFontSize = Len(tLogFont)
    
    
    iCF.lStructSize = Len(iCF)
    iCF.hwndOwner = iHwndOwner
    iCF.rgbColors = mColor
'    iCF.lCustData = 0
'    iCF.lpfnHook = 0
    iCF.lpTemplateName = Space$(2048)
    'iCF.hInstance As Long
    
    'iCF.lpszStyle As String
    'sFont = Chr$(0) & Space$(20) & Chr$(0)
    'iCF.lpszStyle = sFont
    
 '   iCF.nFontType = Screen.FontCount
    'iCF.MISSING_ALIGNMENT As Integer
    iCF.nSizeMin = mMin
    iCF.nSizeMax = mMax
                    
    
    'iCF.lpLogFont As Long
    
    ' Get a global memory block to hold a copy of tLogFont - exit on failure
    lMemHandle = GlobalAlloc(GHND, lLogFontSize)
    If lMemHandle = 0 Then
        Exit Sub
    End If
    
    ' Lock tLogFont's global memory block - exit on failure
    lLogFontAddress = GlobalLock(lMemHandle)
    If lLogFontAddress = 0 Then
        Exit Sub
    End If
    
    ' Copy tLogFont to the global memory block
    Call CopyMemory(ByVal lLogFontAddress, tLogFont, lLogFontSize)
 
    iCF.lpLogFont = lLogFontAddress
    
    'iCF.iPointSize As Long - init from FontSize property
    iCF.iPointSize = mFont.Size * 10
    
    ' Flags can get reference variable or constant with bit Flags
    iCF.Flags = iFlags Or CF_INITTOLOGFONTSTRUCT And Not (CF_APPLY Or CF_ENABLEHOOK Or CF_ENABLETEMPLATE)
    
    If mhDc = 0 Then
        If Not Printer Is Nothing Then
            If ((iCF.Flags And CF_PRINTERFONTS) = CF_PRINTERFONTS) Then iCF.hDC = Printer.hDC
        End If
    Else
        iCF.Flags = iCF.Flags Or CF_PRINTERFONTS
        iCF.hDC = mhDc
    End If
    ' Must have some fonts
    If (iCF.Flags And CF_PRINTERFONTS) = 0 Then iCF.Flags = iCF.Flags Or CF_SCREENFONTS
    ' Color can take initial color, receive chosen color
    If mColor <> 0 Then iCF.Flags = iCF.Flags Or CF_EFFECTS
    ' MinSize can be minimum size accepted
    If (mMin <> 0) Then iCF.Flags = iCF.Flags Or CF_LIMITSIZE
    ' MaxSize can be maximum size accepted
    If (mMax <> 0) Then iCF.Flags = iCF.Flags Or CF_LIMITSIZE
    
    'iCF.Flags = iCF.Flags And Not CF_LIMITSIZE
    
    '***    call the CHOOSEFONT API function
    mApiReturn = ChooseFont(iCF)      'store to APIReturn property
    
    
    '***    handle return from CHOOSEFONT API function
    mCanceled = False
    Select Case mApiReturn
        
        Case 0  'user canceled
        mCanceled = True
        If mCancelError Then
            'generate an error
            Err.Raise 32755, TypeName(Me), "Cancel Pressed"
            Exit Sub
        End If
        
        Case 1  'user selected a font
            ' Copy global memory block to tLogFont
            Call CopyMemory(tLogFont, ByVal lLogFontAddress, lLogFontSize)
            
            'tLogFont.lfWeight As Long  - store to FontBold property
            If tLogFont.lfWeight >= FW_BOLD Then
                mFont.Bold = True
            Else
                mFont.Bold = False
            End If
                        
            'tLogFont.lfItalic As Byte - store to FontItalic property
            If tLogFont.lfItalic <> 0 Then
                mFont.Italic = True
            Else
                mFont.Italic = False
            End If
            
            'tLogFont.lfUnderline As Byte - store to FontUnderline property
            If tLogFont.lfUnderline <> 0 Then
                mFont.Underline = True
            Else
                mFont.Underline = False
            End If
        
            'tLogFont.lfStrikeOut As Byte - store to FontStrikethru property
            If tLogFont.lfStrikeOut <> 0 Then
                mFont.Strikethrough = True
            Else
                mFont.Strikethrough = False
            End If
            
            mColor = iCF.rgbColors
            
            'tLogFont.lfFaceName(LF_FACESIZE) As Byte - store to FontName property
            
            iStr = CStr(tLogFont.lfFaceName)
            iLng = InStr(iStr, Chr(0))
            If iLng > 0 Then
                iStr = Left$(iStr, iLng)
            End If
            FontName = iStr
            
            'iCF.iPointSize As Long - store to FontSize property
            mFont.Size = CLng(iCF.iPointSize / 10)
        
            If mFont.Name <> iFontName_Prev Then mChanged = True
            If mFont.Size <> iFontSize_Prev Then mChanged = True
            If mFont.Bold <> iFontBold_prev Then mChanged = True
            If mFont.Italic <> iFontItalic_Prev Then mChanged = True
            If mFont.Strikethrough <> iFontStrikethrough_Prev Then mChanged = True
            If mFont.Underline <> iFontUnderline_Prev Then mChanged = True
            If mFont.Weight <> iFontWidth_Prev Then mChanged = True
        
        Case Else   'an error occured
            'call CommDlgExtendedError
            mExtendedError = CommDlgExtendedError   'store to ExtendedError property
        
    End Select
    If lMemHandle <> 0 Then GlobalFree lMemHandle
Exit Sub

ShowFontError:
    Exit Sub
End Sub

Public Sub ShowHelp()
    'run winhelp.exe with the specified help file
    Dim sHelpFileBuff As String
    Dim lData As Long
    
    mChanged = False
    On Error GoTo ShowHelpError
    
    '***    init Private properties
    mAction = 6  'Action property - ShowHelp
    mApiReturn = 0  'APIReturn property
    mExtendedError = 0  'ExtendedError property

    '***    prepare the buffers and parameters for the API function
    'mHelpFile is a null terminated string
    sHelpFileBuff = mHelpFile & Chr$(0)
    
    'sData is dependent on mHelpCommand
    Select Case mHelpCommand
        Case 0
            lData = 0
        Case Else
            lData = 0
    End Select
    
    '***    call the API function
    mApiReturn = WinHelp(0&, StrPtr(mHelpFile), mHelpCommand, lData)      ' - Store to APIReturn property
    
    Select Case mApiReturn
        
        Case 0  '
            'call CommDlgExtendedError
            mExtendedError = CommDlgExtendedError   ' - store to ExtendedError property
            
        Case Else   '
            'call CommDlgExtendedError
            mExtendedError = CommDlgExtendedError
        
    End Select
        
Exit Sub

ShowHelpError:
    Exit Sub
End Sub


Public Sub ShowOpen(Optional ByVal nFlags As cdeCommonDialogExFileFlagsConstants = -1)
    Dim iHwndOwner As Long
    Dim iFlags As Long
    
    'display the file open dialog box
    iHwndOwner = GetActiveWindowHwnd
    iFlags = mFlags
    If nFlags <> -1 Then
        iFlags = iFlags Or nFlags
    End If
    ShowFileDialog 1, iFlags, iHwndOwner    'Action property - ShowOpen
    
End Sub

Public Sub ShowPrinter(Optional ByVal nFlags As cdeCommonDialogExPrinterFlagsConstants = -1)
    Dim iHwndOwner As Long
    Dim iFlags As Long
    
    Dim iDeviceName_Prev  As String
    Dim iOrientation_Prev As Long
    Dim iPaperSize_Prev As Long
    Dim iPaperBin_Prev As Long
    Dim iPrintQuality_Prev As Long
    Dim iColorMode_Prev As Long
    Dim iDuplex_Prev As Long
    Dim iCopies_Prev As Long
    Dim iCollate_Prev As Long
    Dim tPrintDlg As T_PrintDlg
    
    iDeviceName_Prev = mDeviceName
    iOrientation_Prev = mOrientation
    iPaperSize_Prev = mPaperSize
    iPaperBin_Prev = mPaperBin
    iPrintQuality_Prev = mPrintQuality
    iColorMode_Prev = mColorMode
    iDuplex_Prev = mDuplex
    iCopies_Prev = mCopies
    iCollate_Prev = mCollate
    
    mChanged = False
    
    On Error GoTo ShowPrinterError
    
    iHwndOwner = GetActiveWindowHwnd
    iFlags = mFlags
    If nFlags <> -1 Then
        iFlags = iFlags Or nFlags
    End If
    
    iFlags = iFlags And Not PD_ENABLEPRINTHOOK
    iFlags = iFlags And Not PD_ENABLEPRINTTEMPLATE
    iFlags = iFlags And Not PD_ENABLESETUPHOOK
    iFlags = iFlags And Not PD_ENABLESETUPTEMPLATE
    
    mAction = 5  'Action property - ShowPrint
    mApiReturn = 0  'APIReturn property
    mExtendedError = 0  'ExtendedError property
    
    tPrintDlg.lStructSize = Len(tPrintDlg)
    tPrintDlg.hwndOwner = iHwndOwner
    
    If (iFlags And cdlPDReturnDefault) <> 0 Then
        tPrintDlg.hDevMode = 0
        tPrintDlg.hDevNames = 0
        If mDevModePtr <> 0 Then
            mDevModePtr = 0
            mDevNamesPtr = 0
            GlobalUnlock mhDevMode
            GlobalUnlock mhDevNames
            mhDevMode = 0
            mhDevNames = 0
        End If
    Else
        If mDevModePtr <> 0 Then
            tPrintDlg.hDevMode = mhDevMode
            UpdateDevModeWithCurrentSettings
            tPrintDlg.hDevNames = mhDevNames
            GlobalUnlock mhDevMode
            GlobalUnlock mhDevNames
            mDevModePtr = 0
            mDevNamesPtr = 0
        End If
    End If
    
    tPrintDlg.Flags = iFlags Or cdlPDUseDevModeCopies Or cdlPDReturnDC ' it requires these two Flags to return the number of copies properly
    tPrintDlg.nFromPage = mFromPage
    tPrintDlg.nToPage = mToPage
    tPrintDlg.nMinPage = mMin
    tPrintDlg.nMaxPage = mMax
    
    If tPrintDlg.nMinPage < 1 Then tPrintDlg.nMinPage = 1
    If tPrintDlg.nMaxPage < 1 Then tPrintDlg.nMaxPage = 10000
    If (tPrintDlg.Flags And cdlPDPageNums) <> 0 Then
        If tPrintDlg.nFromPage < tPrintDlg.nMinPage Then tPrintDlg.nFromPage = tPrintDlg.nMinPage
        If (tPrintDlg.nToPage < tPrintDlg.nFromPage) Then tPrintDlg.nToPage = tPrintDlg.nFromPage
        If (tPrintDlg.nToPage > tPrintDlg.nMaxPage) Then tPrintDlg.nToPage = tPrintDlg.nMaxPage
    End If
    
    tPrintDlg.nCopies = mCopies
    
    '***    Call the PrintDlg API function
    mApiReturn = PrintDlg(tPrintDlg)
    
    '***    handle return from PrintDlg API function
    mCanceled = False
'    If (tPrintDlg.Flags And cdlPDReturnDefault) <> 0 Then
'        mApiReturn = 1
'    End If
    Select Case mApiReturn
        
        Case 0  'user canceled
            If tPrintDlg.hDevMode <> 0 Then
                mDevModePtr = GlobalLock(tPrintDlg.hDevMode)
                mhDevMode = tPrintDlg.hDevMode
                mDevNamesPtr = GlobalLock(tPrintDlg.hDevNames)
                mhDevNames = tPrintDlg.hDevNames
            End If
            
            mCanceled = True
            If mCancelError Then
                'generate an error
                Err.Raise 32755, TypeName(Me), "Cancel Pressed"
                Exit Sub
            End If
        
        Case 1  'user selected OK
            
            mFlags = mFlags And Not (cdlPDSelection Or cdlPDPageNums)
            mFlags = mFlags Or (tPrintDlg.Flags And (cdlPDSelection Or cdlPDPageNums))

            mFromPage = tPrintDlg.nFromPage
            mToPage = tPrintDlg.nToPage
            mMin = tPrintDlg.nMinPage
            mMax = tPrintDlg.nMaxPage
            mhDc = tPrintDlg.hDC
            
            If mDevModePtr = 0 Then
                mDevModePtr = GlobalLock(tPrintDlg.hDevMode)
                mhDevMode = tPrintDlg.hDevMode
            Else
                If mhDevMode <> tPrintDlg.hDevMode Then
                    GlobalUnlock mhDevMode
                    mDevModePtr = GlobalLock(tPrintDlg.hDevMode)
                    mhDevMode = tPrintDlg.hDevMode
                End If
            End If
            CopyMemory mDevMode, ByVal mDevModePtr, Len(mDevMode)
            mPaperSize = mDevMode.dmPaperSize
            mPaperBin = mDevMode.dmDefaultSource
            mDuplex = mDevMode.dmDuplex
            mOrientation = mDevMode.dmOrientation
            mPrintQuality = mDevMode.dmPrintQuality
            mColorMode = mDevMode.dmColor
            mDefaultPaperWidth = mDevMode.dmPaperWidth
            mDefaultPaperHeight = mDevMode.dmPaperLength
            mCopies = mDevMode.dmCopies
            If mCopies < 1 Then
                mCopies = 1
            End If
            mCollate = CBool(mDevMode.dmCollate)
            
            If mDevNamesPtr = 0 Then
                mDevNamesPtr = GlobalLock(tPrintDlg.hDevNames)
                mhDevNames = tPrintDlg.hDevNames
            Else
                If mhDevNames <> tPrintDlg.hDevNames Then
                    GlobalUnlock mhDevNames
                    mDevNamesPtr = GlobalLock(tPrintDlg.hDevNames)
                    mhDevNames = tPrintDlg.hDevNames
                End If
            End If
            CopyMemory mDevNames, ByVal mDevNamesPtr, Len(mDevNames)
            mDriverName = GetDevNameString(mDevNamesPtr, mDevNames.wDriverOffset)
            mDeviceName = GetDevNameString(mDevNamesPtr, mDevNames.wDeviceOffset)
            mPort = GetDevNameString(mDevNamesPtr, mDevNames.wOutputOffset)
            
'            Debug.Print mDevMode.dmDeviceName
            
            PutPaperSize
            mPageSet = True
            
            If mDeviceName <> iDeviceName_Prev Then mChanged = True
            If mOrientation <> iOrientation_Prev Then mChanged = True
            If mPaperSize <> iPaperSize_Prev Then mChanged = True
            If mPaperBin <> iPaperBin_Prev Then mChanged = True
            If mPrintQuality <> iPrintQuality_Prev Then mChanged = True
            If mColorMode <> iColorMode_Prev Then mChanged = True
            If mDuplex <> iDuplex_Prev Then mChanged = True
            If mCopies <> iCopies_Prev Then mChanged = True
            If mCollate <> iCollate_Prev Then mChanged = True
            
        Case Else   'an error occured
            'call CommDlgExtendedError
            mExtendedError = CommDlgExtendedError   'store to ExtendedError property
    
    End Select

Exit Sub

ShowPrinterError:
    
    Exit Sub
    
End Sub


Public Sub ShowSave(Optional ByVal nFlags As cdeCommonDialogExFileFlagsConstants = -1)
    Dim iHwndOwner As Long
    Dim iFlags As Long
    
    'display the file save dialog box
    iHwndOwner = GetActiveWindowHwnd
    iFlags = mFlags
    If nFlags <> -1 Then
        iFlags = iFlags Or nFlags
    End If
    
    ShowFileDialog 2, iFlags, iHwndOwner   'Action property - ShowSave

End Sub


Public Property Get FileName() As String
    FileName = mFileName
End Property

Public Property Let FileName(ByVal nValue As String)
    If mFileName <> nValue Then
        mFileName = nValue
    End If
End Property


Public Property Get Filter() As String
    Filter = mFilter
End Property

Public Property Let Filter(ByVal nValue As String)
    If mFilter <> nValue Then
        mFilter = nValue
    End If
End Property


Private Function sLeftOfNull(ByVal sIn As String)
    'returns the part of sIn preceding Chr$(0)
    Dim lNullPos As Long
    
    'init output
    sLeftOfNull = sIn
    
    'get position of first Chr$(0) in sIn
    lNullPos = InStr(sIn, Chr$(0))
    
    'return part of sIn to Left of first Chr$(0) if found
    If lNullPos > 0 Then
        sLeftOfNull = Mid$(sIn, 1, lNullPos - 1)
    End If
    
End Function

Private Function sLeftOfLastNull(ByVal sIn As String)
    'returns the part of sIn preceding Chr$(0)
    Dim iNullPos As Long
    Dim iChr As String
    
    'init output
    sLeftOfLastNull = sIn
    
    'get position of first Chr$(0) in sIn
    iNullPos = InStrRev(sIn, Chr$(0))
    
    'return part of sIn to Left of first Chr$(0) if found
    If iNullPos > 0 Then
        iChr = Mid$(sIn, iNullPos, 1)
        Do Until (iChr <> Chr$(0)) And (iChr <> " ")
            iNullPos = iNullPos - 1
            If iNullPos = 0 Then
                iNullPos = InStrRev(sIn, Chr$(0))
                Exit Do
            End If
            iChr = Mid$(sIn, iNullPos, 1)
        Loop
        iNullPos = iNullPos + 1
        sLeftOfLastNull = Mid$(sIn, 1, iNullPos - 1)
    End If
    
End Function


Private Function sAPIFilter(sIn)
    'prepares sIn for use as a filter string in API common dialog functions
    Dim lChrNdx As Long
    Dim sOneChr As String
    Dim sOutStr As String
    
    'convert any | characters to nulls
    For lChrNdx = 1 To Len(sIn)
        sOneChr = Mid$(sIn, lChrNdx, 1)
        If sOneChr = "|" Then
            sOutStr = sOutStr & Chr$(0)
        Else
            sOutStr = sOutStr & sOneChr
        End If
    Next
    
    'add a null to the end
    sOutStr = sOutStr & Chr$(0)
    
    'return sOutStr
    sAPIFilter = sOutStr
    
End Function

Public Property Get FilterIndex() As Integer
    FilterIndex = mFilterIndex
End Property

Public Property Let FilterIndex(ByVal nValue As Integer)
    If mFilterIndex <> nValue Then
        mFilterIndex = nValue
    End If
End Property

Public Property Get CancelError() As Boolean
    CancelError = mCancelError
End Property

Public Property Let CancelError(ByVal nValue As Boolean)
    If mCancelError <> nValue Then
        mCancelError = nValue
    End If
End Property

Public Property Get Color() As Long
    Color = mColor
End Property

Public Property Let Color(ByVal nValue As Long)
    If mColor <> nValue Then
        mColor = nValue
    End If
End Property

Public Property Get Copies() As Long
    Copies = mCopies
End Property

Public Property Let Copies(ByVal nValue As Long)
    If mCopies <> nValue Then
        mCopies = nValue
    End If
End Property

Public Property Get Collate() As Boolean
    Collate = mCollate
End Property

Public Property Let Collate(ByVal nValue As Boolean)
    If mCollate <> nValue Then
        mCollate = nValue
    End If
End Property

Public Property Get DefaultExt() As String
    DefaultExt = mDefaultExt
End Property

Public Property Let DefaultExt(ByVal nValue As String)
    If mDefaultExt <> nValue Then
        mDefaultExt = nValue
    End If
End Property

Public Property Get DialogTitle() As String
    DialogTitle = mDialogTitle
End Property

Public Property Let DialogTitle(ByVal nValue As String)
    If mDialogTitle <> nValue Then
        mDialogTitle = nValue
    End If
End Property

Public Property Get Flags() As CommonDialogsFlags
    Flags = mFlags
End Property

Public Property Let Flags(ByVal nValue As CommonDialogsFlags)
    If mFlags <> nValue Then
        mFlags = nValue
    End If
End Property

Public Property Get FontBold() As Boolean
    FontBold = mFontBold
End Property

Public Property Let FontBold(ByVal nValue As Boolean)
    If mFontBold <> nValue Then
        mFont.Bold = nValue
        mFontBold = mFont.Bold
    End If
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = mFontItalic
End Property

Public Property Let FontItalic(ByVal nValue As Boolean)
    If mFontItalic <> nValue Then
        mFont.Italic = nValue
        mFontItalic = mFont.Italic
    End If
End Property

Public Property Get FontName() As String
    If mFontNameSet Then
        FontName = mFontName
    End If
End Property

Public Property Let FontName(ByVal nValue As String)
    nValue = Trim$(nValue)
    If mFontName <> nValue Then
        mFont.Name = nValue
        If mAmbientUserMode Then
            mFontName = mFont.Name
        Else
            mFontName = nValue
        End If
        mFontNameSet = (nValue <> "")
    End If
End Property

Public Property Get FontSize() As Single
    FontSize = mFontSize
End Property

Public Property Let FontSize(ByVal nValue As Single)
    If mFontSize <> nValue Then
        mFont.Size = nValue
        mFontSize = nValue
    End If
End Property

Public Property Get FontStrikeThru() As Boolean
    FontStrikeThru = mFontStrikethrough
End Property

Public Property Let FontStrikeThru(ByVal nValue As Boolean)
    If mFontStrikethrough <> nValue Then
        mFont.Strikethrough = nValue
        mFontStrikethrough = mFont.Strikethrough
    End If
End Property

Public Property Get FontUnderLine() As Boolean
    FontUnderLine = mFontUnderline
End Property

Public Property Let FontUnderLine(ByVal nValue As Boolean)
    If mFontUnderline <> nValue Then
        mFont.Underline = nValue
        mFontUnderline = mFont.Underline
    End If
End Property

Public Property Get FromPage() As Long
Attribute FromPage.VB_MemberFlags = "40"
    FromPage = mFromPage
End Property

Public Property Let FromPage(ByVal nValue As Long)
    If mFromPage <> nValue Then
        mFromPage = nValue
    End If
End Property

Public Property Get hDC() As Long
    EnsurePageSet
    hDC = mhDc
End Property

'Public Property Let hDC(ByVal nValue As Long)
'    If mhDc <> nValue Then
'        mhDc = nValue
'    End If
'End Property

Public Property Get HelpCommand() As Long
    HelpCommand = mHelpCommand
End Property

Public Property Let HelpCommand(ByVal nValue As Long)
    If mHelpCommand <> nValue Then
        mHelpCommand = nValue
    End If
End Property

Public Property Get HelpContext() As Long
    HelpContext = mHelpContext
End Property

Public Property Let HelpContext(ByVal nValue As Long)
    If mHelpContext <> nValue Then
        mHelpContext = nValue
    End If
End Property

Public Property Get HelpFile() As String
    HelpFile = mHelpFile
End Property

Public Property Let HelpFile(ByVal nValue As String)
    If mHelpFile <> nValue Then
        mHelpFile = nValue
    End If
End Property

Public Property Get HelpKey() As String
    HelpKey = mHelpKey
End Property

Public Property Let HelpKey(ByVal nValue As String)
    If mHelpKey <> nValue Then
        mHelpKey = nValue
    End If
End Property

Public Property Get InitDir() As String
    InitDir = mInitDir
End Property

Public Property Let InitDir(ByVal nValue As String)
    nValue = Trim$(nValue)
    If mInitDir <> nValue Then
        mInitDir = nValue
    End If
End Property

Public Property Get Max() As Long
    Max = mMax
End Property

Public Property Let Max(ByVal nValue As Long)
    If mMax <> nValue Then
        mMax = nValue
        If mMin > mMax Then
            mMin = mMax
        End If
    End If
End Property

Public Property Get MaxFileSize() As Long
    MaxFileSize = mMaxFileSize
End Property

Public Property Let MaxFileSize(ByVal nValue As Long)
    If mMaxFileSize <> nValue Then
        mMaxFileSize = nValue
    End If
End Property

Public Property Get Min() As Long
    Min = mMin
End Property

Public Property Let Min(ByVal nValue As Long)
    If mMin <> nValue Then
        mMin = nValue
        If mMax < mMin Then
            mMax = mMin
        End If
    End If
End Property

Public Property Get Object() As Object
    Set Object = Me
End Property

Public Property Get PrinterDefault() As Boolean  ' added just for compatibility
Attribute PrinterDefault.VB_MemberFlags = "40"

End Property

Public Property Let PrinterDefault(ByVal nValue As Boolean) ' added just for compatibility

End Property

Public Property Get ToPage() As Long
    ToPage = mToPage
End Property

Public Property Let ToPage(ByVal nValue As Long)
    If mToPage <> nValue Then
        mToPage = nValue
    End If
End Property

Public Property Get FileTitle() As String
    FileTitle = mFileTitle
End Property

Public Property Let FileTitle(ByVal nValue As String)
    If mFileTitle <> nValue Then
        mFileTitle = nValue
    End If
End Property

'Private Property Get APIReturn() As Long
'    APIReturn = mApiReturn
'End Property
'
'Private Property Get ExtendedError() As Long
'    ExtendedError = mExtendedError
'End Property

Private Sub ShowFileDialog(ByVal mAction As Integer, nFlags As Long, Optional ByVal nHwndOwner As Long)
    
    'display the file dialog for ShowOpen or ShowSave
    
    Dim tOpenFile As T_OPENFILENAME
    Dim lMaxSize As Long
    Dim sFileNameBuff As String
    Dim sFileTitleBuff As String
    Dim iPos As Long
    Dim iFileName_Prev As String
    Dim iStr As String
    Dim iBuffer As String
    Dim iLen As Long
    
    On Error GoTo ShowFileDialogError
    
    mCanceled = False
    mChanged = False
    iFileName_Prev = mFileName
    '***    init property buffers
    
    mAction = mAction  'Action property
    mApiReturn = 0  'APIReturn property
    mExtendedError = 0  'ExtendedError property
        
    
    '***    prepare tOpenFile data
    
    tOpenFile.lStructSize = Len(tOpenFile)
    tOpenFile.hwndOwner = nHwndOwner
    iStr = sAPIFilter(mFilter)
    tOpenFile.lpstrFilter = StrPtr(iStr)
    tOpenFile.nFilterIndex = mFilterIndex
    
    'determine size of buffer from MaxFileSize property
    If mMaxFileSize > 0 Then
        lMaxSize = mMaxFileSize
    Else
        lMaxSize = 255
    End If
    
    'prepare sFileNameBuff
    sFileNameBuff = mFileName
    If InStr(sFileNameBuff, "//") > 0 Then
        Do Until InStr(sFileNameBuff, "//") = 0
            sFileNameBuff = Replace(sFileNameBuff, "//", "/")
        Loop
    End If
    If InStr(sFileNameBuff, "\\") > 0 Then
        Do Until InStr(sFileNameBuff, "\\") = 0
            sFileNameBuff = Replace(sFileNameBuff, "\\", "\")
        Loop
    End If
    If InStr(sFileNameBuff, "\\") > 0 Then
        Do Until InStr(sFileNameBuff, "\\") = 0
            sFileNameBuff = Replace(sFileNameBuff, "\\", "\")
        Loop
    End If
    iPos = InStr(sFileNameBuff, "\")
    If InStr(sFileNameBuff, "/") < iPos Then
        iPos = InStr(sFileNameBuff, "\")
    End If
    If iPos = 0 Then
        iPos = Len(sFileNameBuff)
    End If
    
    If InStr(iPos, sFileNameBuff, ":") > 0 Then
        sFileNameBuff = Replace(sFileNameBuff, ":", "")
    End If
    If InStr(iPos, sFileNameBuff, "?") > 0 Then
        sFileNameBuff = Replace(sFileNameBuff, "?", "")
    End If
    If InStr(iPos, sFileNameBuff, "*") > 0 Then
        sFileNameBuff = Replace(sFileNameBuff, "*", "")
    End If
    If InStr(iPos, sFileNameBuff, """") > 0 Then
        sFileNameBuff = Replace(sFileNameBuff, """", "")
    End If
    If InStr(iPos, sFileNameBuff, ">") > 0 Then
        sFileNameBuff = Replace(sFileNameBuff, ">", "")
    End If
    If InStr(iPos, sFileNameBuff, "<") > 0 Then
        sFileNameBuff = Replace(sFileNameBuff, "<", "")
    End If
    If InStr(iPos, sFileNameBuff, "|") > 0 Then
        sFileNameBuff = Replace(sFileNameBuff, "|", "")
    End If
    
    'pad with spaces
    If (nFlags And cdlOFNAllowMultiselect) <> 0 Then
        sFileNameBuff = sFileNameBuff & String$(500 - Len(sFileNameBuff) - 1, Chr(0))
        sFileNameBuff = Mid$(sFileNameBuff, 1, 499)
    Else
        While Len(sFileNameBuff) < lMaxSize - 1
            sFileNameBuff = sFileNameBuff & " "
        Wend
    'Trim$ to length of mMaxFileSize - 1
        sFileNameBuff = Mid$(sFileNameBuff, 1, lMaxSize - 1)
    End If
    'null terminate
    sFileNameBuff = sFileNameBuff & Chr$(0)
    tOpenFile.lpstrFile = StrPtr(sFileNameBuff)
    
    'nMaxFile As Long - init from MaxFileSize property
    'If mMaxFileSize <> 255 Then  'default is 255
    If (nFlags And cdlOFNAllowMultiselect) <> 0 Then
        tOpenFile.nMaxFile = 499
    Else
        tOpenFile.nMaxFile = lMaxSize
    End If
    'End If
            
    'prepare sFileTitleBuff
    sFileTitleBuff = mFileTitle
    'pad with spaces
    While Len(sFileTitleBuff) < lMaxSize - 1
        sFileTitleBuff = sFileTitleBuff & " "
    Wend
    'Trim$ to length of mMaxFileSize - 1
    sFileTitleBuff = Mid$(sFileTitleBuff, 1, lMaxSize - 1)
    'null terminate
    sFileTitleBuff = sFileTitleBuff & Chr$(0)
    tOpenFile.lpstrFileTitle = StrPtr(sFileTitleBuff)
        
    tOpenFile.lpstrInitialDir = StrPtr(mInitDir)
    tOpenFile.lpstrTitle = StrPtr(mDialogTitle)
    tOpenFile.Flags = nFlags Or cdlOFNNoChangeDir
    tOpenFile.lpstrDefExt = StrPtr(mDefaultExt)
    
    
    '***    call the GetOpenFileName API function
    Select Case mAction
        Case 1  'ShowOpen
            mApiReturn = GetOpenFileName(tOpenFile)
        Case 2  'ShowSave
            mApiReturn = GetSaveFileName(tOpenFile)
        Case Else   'unknown action
            Exit Sub
    End Select
    
    
    '***    handle return from GetOpenFileName API function
    mCanceled = False
    Select Case mApiReturn
        
        Case 0  'user canceled
            mCanceled = True
            mFileName = ""
            mFileTitle = ""
            If mCancelError Then
                'generate an error
                On Error Resume Next
                Err.Raise 32755, TypeName(Me), "Cancel Pressed"
                Exit Sub
            End If
        
        Case 1  'user selected or entered a file
            'mFileName gets part of tOpenFile.lpstrFile to the Left of first Chr$(0)
            
            iBuffer = String(lMaxSize, vbNullChar)
            iLen = lstrlen(tOpenFile.lpstrFile)
            If iLen > lMaxSize Then iLen = lMaxSize
            CopyMemory ByVal StrPtr(iBuffer), ByVal tOpenFile.lpstrFile, iLen * 2
            
            If (tOpenFile.Flags And cdlOFNAllowMultiselect) <> 0 Then
                mFileName = sLeftOfLastNull(iBuffer)
            Else
                mFileName = sLeftOfNull(iBuffer)
            End If
            mFileTitle = sLeftOfNull(tOpenFile.lpstrFileTitle)
            mChanged = mFileName <> iFileName_Prev
        Case Else   'an error occured
            'call CommDlgExtendedError
            mExtendedError = CommDlgExtendedError
        
    End Select
    

Exit Sub

ShowFileDialogError:
    
    Exit Sub

End Sub


Public Property Get Canceled() As Boolean
    Canceled = mCanceled
End Property

Private Function TranslateAColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
   If OleTranslateColor(clr, hPal, TranslateAColor) Then
      TranslateAColor = CLR_INVALID
   End If
End Function


Private Sub Class_Initialize()
    mMaxFileSize = 255
    mUnits = vbMillimeters
    mUnitsForUser = cdeMUUserLocale
    mAmbientUserMode = True
    mCopies = 1
    mLeftMargin = cLeftMarginDefault
    mRightMargin = cRightMarginDefault
    mTopMargin = cTopMarginDefault
    mBottomMargin = cBottomMarginDefault
    mAutoSaveCustomColors = True
    InitCustomColors
    
    Set mFont = New StdFont
    mFontName = "MS Sans Serif"
    mFontSize = 8
    mFontBold = False
    mFontItalic = False
    mFontStrikethrough = False
    mFontUnderline = False
    
    SetFontObject
    
End Sub

Private Sub SetFontObject()
    mFont.Name = mFontName: mFontName = mFont.Name
    mFont.Size = mFontSize
    mFont.Bold = mFontBold: mFontBold = mFont.Bold
    mFont.Italic = mFontItalic: mFontItalic = mFont.Italic
    mFont.Strikethrough = mFontStrikethrough: mFontStrikethrough = mFont.Strikethrough
    mFont.Underline = mFontUnderline: mFontUnderline = mFont.Underline
End Sub

Public Property Let Action(ByVal nValue As Integer)
    Select Case nValue
        Case 0
            '
        Case 1
            ShowOpen
        Case 2
            ShowSave
        Case 3
            ShowColor
        Case 4
            ShowFont
        Case 5
            ShowPrinter
        Case 6
            ShowHelp
    End Select
End Property


Public Property Get Orientation() As cdePageOrientationConstants
    EnsurePageSet
    Orientation = mOrientation
End Property

Public Property Let Orientation(ByVal nValue As cdePageOrientationConstants)
    If mOrientation <> nValue Then
        If mAmbientUserMode Then
            If (nValue = 0) Then Exit Property
            EnsurePageSet
        End If
        mOrientation = nValue
    End If
End Property


Public Property Get PaperSize() As cdePaperSizeConstants
    EnsurePageSet
    PaperSize = mPaperSize
End Property

Public Property Let PaperSize(ByVal nValue As cdePaperSizeConstants)
    If mPaperSize <> nValue Then
        If mAmbientUserMode Then
            If (nValue = 0) Then Exit Property
            EnsurePageSet
        End If
        mPaperSize = nValue
        If mAmbientUserMode Then
            If Not PaperSizeExists(mPaperSize) Then
                mPaperSize = FindMostSimilarPaperSize(mPaperSize)
            End If
            PutPaperSize
        End If
    End If
End Property


Public Property Get PrintQuality() As cdePrintQualityConstants
    EnsurePageSet
    PrintQuality = mPrintQuality
End Property

Public Property Let PrintQuality(ByVal nValue As cdePrintQualityConstants)
    If mPrintQuality <> nValue Then
        If mAmbientUserMode Then
            If (nValue = 0) Then Exit Property
            EnsurePageSet
        End If
        mPrintQuality = nValue
    End If
End Property


Public Property Get ColorMode() As cdeColorModeConstants
    EnsurePageSet
    ColorMode = mColorMode
End Property

Public Property Let ColorMode(ByVal nValue As cdeColorModeConstants)
    If mColorMode <> nValue Then
        If mAmbientUserMode Then
            If nValue = vbPRCMPrinterDefault Then
                If Not PrinterExCurrentDocument Is Nothing Then
                    nValue = PrinterExCurrentDocument.DefaultColorMode
                End If
            End If
            If (nValue = 0) Then Exit Property
            If (nValue < vbPRCMMonochrome) Or (nValue > vbPRCMColor) Then RaiseError 380, TypeName(Me) ', "Invalid property value"
            EnsurePageSet
        End If
        mColorMode = nValue
    End If
End Property


Public Property Get DriverName() As String
    EnsurePageSet
    DriverName = mDriverName
End Property


Public Property Get Duplex() As cdeDuplexConstants
    EnsurePageSet
    Duplex = mDuplex
End Property

Public Property Let Duplex(ByVal nValue As cdeDuplexConstants)
    If mDuplex <> nValue Then
        If mAmbientUserMode Then
            If (nValue = 0) Then Exit Property
            EnsurePageSet
        End If
        mDuplex = nValue
    End If
End Property


Public Property Get PaperBin() As cdePaperBinConstants
    EnsurePageSet
    PaperBin = mPaperBin
End Property

Public Property Let PaperBin(ByVal nValue As cdePaperBinConstants)
    If mPaperBin <> nValue Then
        If mAmbientUserMode Then
            If (nValue = 0) Then Exit Property
            EnsurePageSet
        End If
        mPaperBin = nValue
    End If
End Property


Public Property Get Port() As String
    EnsurePageSet
    Port = mPort
End Property


Public Property Get DeviceName() As String
    If mDeviceName = "" Then
        EnsurePageSet
    End If
    DeviceName = mDeviceName
End Property

Public Property Let DeviceName(ByVal nValue As String)
    Dim iDm As DEVMODE
    
    If nValue <> mDeviceName Then
        If GetDevModeByDeviceName(nValue, iDm) Then
            If mDevModePtr <> 0 Then
                GlobalUnlock mhDevMode
                If mhDevModeToFree <> 0 Then
                    GlobalFree mhDevModeToFree
                    mhDevModeToFree = 0
                End If
                GlobalUnlock mhDevNames
            End If
            mhDevMode = GlobalAlloc(GHND, Len(iDm))
            mhDevModeToFree = mhDevMode
            mDevModePtr = GlobalLock(mhDevMode)
            CopyMemory ByVal mDevModePtr, iDm, Len(iDm)
            CopyMemory mDevMode, ByVal mDevModePtr, Len(mDevMode)
            mhDevNames = 0
            
            mDeviceName = Left(mDevMode.dmDeviceName, InStr(mDevMode.dmDeviceName, Chr(0)) - 1)
        
            mPaperSize = mDevMode.dmPaperSize
            mPaperBin = mDevMode.dmDefaultSource
            mDuplex = mDevMode.dmDuplex
            mOrientation = mDevMode.dmOrientation
            mPrintQuality = mDevMode.dmPrintQuality
            mColorMode = mDevMode.dmColor
            mDefaultPaperWidth = mDevMode.dmPaperWidth
            mDefaultPaperHeight = mDevMode.dmPaperLength
            mPageSet = True
            PutPaperSize
        End If
    End If
End Property

Public Property Get PaperWidth() As Single
    EnsurePageSet
    If Units = vbInches Then
        PaperWidth = mPaperWidth / 254
    Else
        PaperWidth = mPaperWidth / 10
    End If
End Property


Public Property Get PaperHeight() As Single
    EnsurePageSet
    If Units = vbInches Then
        PaperHeight = mPaperHeight / 254
    Else
        PaperHeight = mPaperHeight / 10
    End If
End Property


Public Property Get LeftMargin() As Single
    EnsurePageSet
    LeftMargin = mLeftMargin
End Property

Public Property Let LeftMargin(ByVal nValue As Single)
    If mLeftMargin <> nValue Then
        EnsurePageSet
        mLeftMargin = nValue
        If mLeftMargin < mMinLeftMargin Then
            mLeftMargin = mMinLeftMargin
        End If
    End If
End Property


Public Property Get MinLeftMargin() As Single
    EnsurePageSet
    MinLeftMargin = mMinLeftMargin
End Property

Public Property Let MinLeftMargin(ByVal nValue As Single)
    If mMinLeftMargin <> nValue Then
        EnsurePageSet
        mMinLeftMargin = nValue
        If mLeftMargin < mMinLeftMargin Then
            mLeftMargin = mMinLeftMargin
        End If
    End If
End Property


Public Property Get RightMargin() As Single
    EnsurePageSet
    RightMargin = mRightMargin
End Property

Public Property Let RightMargin(ByVal nValue As Single)
    If mRightMargin <> nValue Then
        EnsurePageSet
        mRightMargin = nValue
        If mRightMargin < mMinRightMargin Then
            mRightMargin = mMinRightMargin
        End If
    End If
End Property


Public Property Get MinRightMargin() As Single
    EnsurePageSet
    MinRightMargin = mMinRightMargin
End Property

Public Property Let MinRightMargin(ByVal nValue As Single)
    If mMinRightMargin <> nValue Then
        EnsurePageSet
        mMinRightMargin = nValue
        If mRightMargin < mMinRightMargin Then
            mRightMargin = mMinRightMargin
        End If
    End If
End Property


Public Property Get TopMargin() As Single
    EnsurePageSet
    TopMargin = mTopMargin
End Property

Public Property Let TopMargin(ByVal nValue As Single)
    If mTopMargin <> nValue Then
        EnsurePageSet
        mTopMargin = nValue
        If mTopMargin < mMinTopMargin Then
            mTopMargin = mMinTopMargin
        End If
    End If
End Property


Public Property Get MinTopMargin() As Single
    EnsurePageSet
    MinTopMargin = mMinTopMargin
End Property

Public Property Let MinTopMargin(ByVal nValue As Single)
    If mMinTopMargin <> nValue Then
        EnsurePageSet
        mMinTopMargin = nValue
        If mTopMargin < mMinTopMargin Then
            mTopMargin = mMinTopMargin
        End If
    End If
End Property


Public Property Get BottomMargin() As Single
    EnsurePageSet
    BottomMargin = mBottomMargin
End Property

Public Property Let BottomMargin(ByVal nValue As Single)
    If mBottomMargin <> nValue Then
        EnsurePageSet
        mBottomMargin = nValue
        If mBottomMargin < mMinBottomMargin Then
            mBottomMargin = mMinBottomMargin
        End If
    End If
End Property


Public Property Get MinBottomMargin() As Single
    EnsurePageSet
    MinBottomMargin = mMinBottomMargin
End Property

Public Property Let MinBottomMargin(ByVal nValue As Single)
    If mMinBottomMargin <> nValue Then
        EnsurePageSet
        mMinBottomMargin = nValue
        If mBottomMargin < mMinBottomMargin Then
            mBottomMargin = mMinBottomMargin
        End If
    End If
End Property


Public Property Get Units() As cdeUnits
    Units = mUnits
End Property

Public Property Let Units(ByVal nValue As cdeUnits)
    If (nValue <> vbInches) And (nValue <> vbMillimeters) Then Exit Property
    If nValue <> mUnits Then
        mUnits = nValue
        ConvertMarginValues
    End If
End Property


Public Property Get UnitsForUser() As cdeUnitsForUser
    UnitsForUser = mUnitsForUser
End Property

Public Property Let UnitsForUser(ByVal nValue As cdeUnitsForUser)
    If (nValue <> cdeMUInches) And (nValue <> cdeMUMillimeters) And (nValue <> cdeMUUserLocale) Then Exit Property
    If nValue <> mUnitsForUser Then
        mUnitsForUser = nValue
    End If
End Property


Public Sub ShowPageSetup(Optional ByVal nFlags As cdeCommonDialogExPageSetupFlagsConstants = -1)
    Dim iHwndOwner As Long
    Dim iPsd As T_PAGESETUPDLG
    Dim iUnitsMultiplier As Single
    Dim iFlags As Long
    Dim iUnits_Ant  As Long
    
    Dim iLeftMargin_Prev As Single
    Dim iTopMargin_Prev As Single
    Dim iRightMargin_Prev As Single
    Dim iBottomMargin_Prev As Single
    Dim iDeviceName_Prev  As String
    Dim iOrientation_Prev As Long
    Dim iPaperSize_Prev As Long
    Dim iPaperBin_Prev As Long
    Dim iPrintQuality_Prev As Long
    Dim iColorMode_Prev As Long
    Dim iDuplex_Prev As Long
    Dim iCopies_Prev As Long
    Dim iCollate_Prev As Long
    
'    On Error GoTo ShowPageSetupError
    mChanged = False
    
    iHwndOwner = GetActiveWindowHwnd
    
    iFlags = mFlags
    If nFlags <> -1 Then
        iFlags = iFlags Or nFlags
    End If
    
    iUnits_Ant = mUnits
    If mUnits <> vbMillimeters Then
        Units = vbMillimeters
    End If
    
    iPsd.lStructSize = Len(iPsd)
    
    iPsd.Flags = iFlags And Not (PSD_ENABLEPAGEPAINTHOOK Or PSD_ENABLEPAGESETUPHOOK Or PSD_ENABLEPAGESETUPTEMPLATE)
    Select Case mUnitsForUser
        Case cdeMUInches
            iPsd.Flags = iPsd.Flags Or PSD_INTHOUSANDTHSOFINCHES
        Case cdeMUMillimeters
            iPsd.Flags = iPsd.Flags Or PSD_INHUNDREDTHSOFMILLIMETERS
        Case Else
            If GetLocaleMeasureSystem = 1 Then ' EEUU system, Inches
                iPsd.Flags = iPsd.Flags Or PSD_INTHOUSANDTHSOFINCHES
            Else
                iPsd.Flags = iPsd.Flags Or PSD_INHUNDREDTHSOFMILLIMETERS
            End If
    End Select
    
    If (iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) <> 0 Then
        iUnitsMultiplier = 1000
    Else
        iUnitsMultiplier = 100
    End If
    
    Select Case True
        Case (Units = vbInches) And ((iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) = 0)
            iUnitsMultiplier = iUnitsMultiplier * 25.4
        Case (Units = vbInches) And ((iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) <> 0)
            ' OK
        Case (Units = vbMillimeters) And ((iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) = 0)
            ' OK
        Case (Units = vbMillimeters) And ((iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) <> 0)
            iUnitsMultiplier = iUnitsMultiplier / 25.4
    End Select
    
    iPsd.Flags = iPsd.Flags Or PSD_MARGINS
    If (mMinTopMargin <> 0) Or (mMinLeftMargin <> 0) Or (mMinBottomMargin <> 0) Or (mMinRightMargin <> 0) Then
        iPsd.Flags = iPsd.Flags Or cdlPSMinMargins
    End If
    
    iPsd.hwndOwner = iHwndOwner
    iPsd.rtMargin.Top = mTopMargin * iUnitsMultiplier
    iPsd.rtMargin.Left = mLeftMargin * iUnitsMultiplier
    iPsd.rtMargin.Bottom = mBottomMargin * iUnitsMultiplier
    iPsd.rtMargin.Right = mRightMargin * iUnitsMultiplier
    iPsd.rtMinMargin.Top = mMinTopMargin * iUnitsMultiplier
    iPsd.rtMinMargin.Left = mMinLeftMargin * iUnitsMultiplier
    iPsd.rtMinMargin.Bottom = mMinBottomMargin * iUnitsMultiplier
    iPsd.rtMinMargin.Right = mMinRightMargin * iUnitsMultiplier
    
    If mDevModePtr <> 0 Then
        iPsd.hDevMode = mhDevMode
        UpdateDevModeWithCurrentSettings
        iPsd.hDevNames = mhDevNames
        GlobalUnlock mhDevMode
        GlobalUnlock mhDevNames
        mDevModePtr = 0
        mDevNamesPtr = 0
    End If
    
    iLeftMargin_Prev = mLeftMargin
    iTopMargin_Prev = mTopMargin
    iRightMargin_Prev = mRightMargin
    iBottomMargin_Prev = mBottomMargin
    iDeviceName_Prev = mDeviceName
    iOrientation_Prev = mOrientation
    iPaperSize_Prev = mPaperSize
    iPaperBin_Prev = mPaperBin
    iPrintQuality_Prev = mPrintQuality
    iColorMode_Prev = mColorMode
    iDuplex_Prev = mDuplex
    iCopies_Prev = mCopies
    iCollate_Prev = mCollate
    
    mApiReturn = PageSetupDlg(iPsd)
    mCanceled = False
    
    Select Case mApiReturn
        
        Case 0  'user canceled
        
        If iPsd.hDevMode <> 0 Then
            mDevModePtr = GlobalLock(iPsd.hDevMode)
            mhDevMode = iPsd.hDevMode
            mDevNamesPtr = GlobalLock(iPsd.hDevNames)
            mhDevNames = iPsd.hDevNames
        End If
        
        mCanceled = True
        If mCancelError Then
            'generate an error
            On Error GoTo 0
            Err.Raise 32755, TypeName(Me), "Cancel Pressed"
            Exit Sub
        End If
        
        Case 1  'user clicked OK
            If (iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) <> 0 Then
                iUnitsMultiplier = 1000
            Else
                iUnitsMultiplier = 100
            End If
            Select Case True
                Case (Units = vbInches) And ((iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) = 0)
                    iUnitsMultiplier = iUnitsMultiplier * 25.4
                Case (Units = vbInches) And ((iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) <> 0)
                    ' OK
                Case (Units = vbMillimeters) And ((iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) = 0)
                    ' OK
                Case (Units = vbMillimeters) And ((iPsd.Flags And PSD_INTHOUSANDTHSOFINCHES) <> 0)
                    iUnitsMultiplier = iUnitsMultiplier / 25.4
            End Select
            
            mTopMargin = Int((iPsd.rtMargin.Top + 0.49) / iUnitsMultiplier * 100) / 100
            mLeftMargin = Int((iPsd.rtMargin.Left + 0.49) / iUnitsMultiplier * 100) / 100
            mBottomMargin = Int((iPsd.rtMargin.Bottom + 0.49) / iUnitsMultiplier * 100) / 100
            mRightMargin = Int((iPsd.rtMargin.Right + 0.49) / iUnitsMultiplier * 100) / 100
            mMinTopMargin = Int((iPsd.rtMinMargin.Top + 0.49) / iUnitsMultiplier * 100) / 100
            mMinLeftMargin = Int((iPsd.rtMinMargin.Left + 0.49) / iUnitsMultiplier * 100) / 100
            mMinBottomMargin = Int((iPsd.rtMinMargin.Bottom + 0.49) / iUnitsMultiplier * 100) / 100
            mMinRightMargin = Int((iPsd.rtMinMargin.Right + 0.49) / iUnitsMultiplier * 100) / 100
            
            If mDevModePtr = 0 Then
                mDevModePtr = GlobalLock(iPsd.hDevMode)
                mhDevMode = iPsd.hDevMode
            End If
            CopyMemory mDevMode, ByVal mDevModePtr, Len(mDevMode)
            mPaperSize = mDevMode.dmPaperSize
            mPaperBin = mDevMode.dmDefaultSource
            mDuplex = mDevMode.dmDuplex
            mOrientation = mDevMode.dmOrientation
            mPrintQuality = mDevMode.dmPrintQuality
            mColorMode = mDevMode.dmColor
            mDefaultPaperWidth = mDevMode.dmPaperWidth
            mDefaultPaperHeight = mDevMode.dmPaperLength
            PutPaperSize
            
            If mDevNamesPtr = 0 Then
                mDevNamesPtr = GlobalLock(iPsd.hDevNames)
                mhDevNames = iPsd.hDevNames
            End If
            If mDevNamesPtr <> 0 Then
                CopyMemory mDevNames, ByVal mDevNamesPtr, Len(mDevNames)
                mDriverName = GetDevNameString(mDevNamesPtr, mDevNames.wDriverOffset)
                mDeviceName = GetDevNameString(mDevNamesPtr, mDevNames.wDeviceOffset)
                mPort = GetDevNameString(mDevNamesPtr, mDevNames.wOutputOffset)
            End If
            mPageSet = True
        
            If mLeftMargin <> iLeftMargin_Prev Then mChanged = True
            If mTopMargin <> iTopMargin_Prev Then mChanged = True
            If mRightMargin <> iRightMargin_Prev Then mChanged = True
            If mBottomMargin <> iBottomMargin_Prev Then mChanged = True
            If mDeviceName <> iDeviceName_Prev Then mChanged = True
            If mOrientation <> iOrientation_Prev Then mChanged = True
            If mPaperSize <> iPaperSize_Prev Then mChanged = True
            If mPaperBin <> iPaperBin_Prev Then mChanged = True
            If mPrintQuality <> iPrintQuality_Prev Then mChanged = True
            If mColorMode <> iColorMode_Prev Then mChanged = True
            If mDuplex <> iDuplex_Prev Then mChanged = True
            If mCopies <> iCopies_Prev Then mChanged = True
            If mCollate <> iCollate_Prev Then mChanged = True
        
        Case Else   'an error occured
            'call CommDlgExtendedError
            mExtendedError = CommDlgExtendedError
        
    End Select
    
    If mUnits <> iUnits_Ant Then
        Units = iUnits_Ant
    End If
    
'    If Not mCanceled Then
'        If mhDc <> 0 Then
'            Dim iDevMode As DEVMODE
'
'            CopyMemory iDevMode, ByVal mDevModePtr, Len(iDevMode)
'            ShowPrinter cdlPDReturnDefault Or cdePDReturnDC
'            ResetDC mhDc, iDevMode
'            mPaperSize = iDevMode.dmPaperSize
'            mPaperBin = iDevMode.dmDefaultSource
'            mDuplex = iDevMode.dmDuplex
'            mOrientation = iDevMode.dmOrientation
'            mPrintQuality = iDevMode.dmPrintQuality
'            mColorMode = iDevMode.dmColor
'            mDefaultPaperWidth = iDevMode.dmPaperWidth
'            mDefaultPaperHeight = iDevMode.dmPaperLength
'            PutPaperSize
'            UpdateDevModeWithCurrentSettings
'        End If
'    End If
    
'ShowPageSetupError:
End Sub

Private Function GetDevNameString( _
      ByVal ptrDevNames As Long, _
      ByVal ptrOffset As Long _
   )
   Dim Ptr As Long
   Dim lSize As Long
   Dim B() As Byte
      
   Ptr = UnsignedAdd(ptrDevNames, ptrOffset)
   lSize = lstrlenPtr(Ptr)
   If (lSize > 0) Then
      ReDim B(0 To lSize - 1) As Byte
      CopyMemory B(0), ByVal Ptr, lSize
      GetDevNameString = StrConv(B, vbUnicode)
   End If
End Function

Private Function UnsignedAdd(Start As Long, Incr As Long) As Long
' This function is useful when doing pointer arithmetic,
' but note it only works for positive values of Incr

   If Start And &H80000000 Then 'Start < 0
      UnsignedAdd = Start + Incr
   ElseIf (Start Or &H80000000) < -Incr Then
      UnsignedAdd = Start + Incr
   Else
      UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)
   End If
   
End Function

Private Sub Class_Terminate()
    If Not mManualTermination Then
        DoTerminate
    End If
End Sub

Friend Property Let ManualTermination(ByVal nValue As Boolean)
    mManualTermination = nValue
End Property

Friend Sub DoTerminate()
    If (mDevModePtr <> 0) Then
        GlobalUnlock mhDevMode
        If mhDevModeToFree <> 0 Then
            GlobalFree mhDevModeToFree
            mhDevModeToFree = 0
        End If
        GlobalUnlock mhDevNames
        mDevModePtr = 0
    End If
End Sub

Private Function GetLocaleMeasureSystem() As Long
    Dim Buffer As String * 100
    Dim nullpos&
    Dim dl&
    
    dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IMEASURE, Buffer, 99)
    nullpos& = InStr(Buffer, Chr$(0))
    GetLocaleMeasureSystem = Val(Left$(Buffer, nullpos - 1))
End Function

Private Sub EnsurePageSet()
    If mAmbientUserMode Then
        If Not mPageSet Then
            If mhDevModeToFree <> 0 Then
                GlobalFree mhDevModeToFree
                mhDevModeToFree = 0
            End If
            ShowPrinter cdlPDReturnDefault Or cdePDReturnDC
        End If
    End If
End Sub


Public Property Get FolderName() As String
    FolderName = mFolderName
End Property

Public Property Let FolderName(ByVal nValue As String)
    If mFolderName <> nValue Then
        mFolderName = nValue
    End If
End Property


Public Property Get FolderDisplayName() As String
    FolderDisplayName = mFolderDisplayName
End Property

Public Property Get RootFolder() As String
    RootFolder = mRootFolder
End Property

Public Property Let RootFolder(ByVal nValue As String)
    nValue = Trim$(nValue)
    If mRootFolder <> nValue Then
        mRootFolder = nValue
    End If
End Property


Public Property Get FolderDialogHeader() As String
    FolderDialogHeader = mFolderDialogHeader
End Property

Public Property Let FolderDialogHeader(ByVal nValue As String)
    If mFolderDialogHeader <> nValue Then
        mFolderDialogHeader = nValue
    End If
End Property


Public Sub ShowFolder(Optional nFlags As cdeCommonDialogExFolderFlagsConstants = -1)
    Dim bInf As BrowseInfo
    Dim RetVal As Long
    Dim PathID As Long
    Dim RetPath As String
    Dim Offset As Integer
    Dim iHwndOwner As Long
    Dim iRootFolder As String
    Dim iCSIDL As Long
    Dim iPidlRoot As ITEMIDLIST
    Dim ipIDList2 As Long
    Dim iFlags As Long
    Dim iAuxFolderPath As String
    Dim iFolder_Prev As String
    
    iFlags = mFlags
    If nFlags <> -1 Then
        iFlags = iFlags Or nFlags
    End If
    If (iFlags And Not cdeSFValidate And Not cdeSFUseNewUI And Not cdeSFUAHint And Not cdeSFStatusText And Not cdeSFNoNewFolderButton And Not cdeSFEditBox And Not cdeSFNewDialogStyle) = 0 Then
        iFlags = iFlags Or cdeSFReturnOnlyFSDirs
    End If
    iFlags = iFlags And Not cdeSFNoReturnOnlyFSDirs
    
    gWindowTitle = mDialogTitle
    iHwndOwner = GetActiveWindowHwnd
    'Set the properties of the folder dialog
    bInf.hwndOwner = iHwndOwner
    bInf.pszDisplayName = String(260, 32)
    
    iRootFolder = Replace(LCase$(mRootFolder), " ", "")
    If (iRootFolder = "") Or iRootFolder = "desktop" Then
        bInf.pIDLRoot = 0
    Else
        Select Case iRootFolder
            Case "mycomputer"
                iCSIDL = CSIDL_DRIVES
            Case "mydocuments"
                iCSIDL = CSIDL_PERSONAL
            Case Else
                ' nothing
        End Select
        If iCSIDL <> 0 Then
            If SHGetFolderLocation(0&, iCSIDL, 0&, 0&, iPidlRoot) = S_OK Then
                bInf.pIDLRoot = iPidlRoot.mkid.cb
            End If
        Else
            If PathIsDirectory(mRootFolder) Then
                SHParseDisplayName StrPtr(mRootFolder), ByVal 0&, ipIDList2, ByVal 0&, ByVal 0&
                bInf.pIDLRoot = ipIDList2
            End If
        End If
    End If
    bInf.lpszTitle = mFolderDialogHeader
    bInf.ulFlags = iFlags
    If bInf.ulFlags = 0 Then
        bInf.ulFlags = cdlSFReturnOnlyFSDirs Or cdlSFStatusText Or cdlSFUseNewUI Or cdlSFValidate
    End If
    
    If (bInf.ulFlags And cdlSFBrowseIncludeURLs) <> 0 Then
        bInf.ulFlags = bInf.ulFlags Or cdlSFUseNewUI Or cdlSFBrowseIncludeFiles
    End If
    If (bInf.ulFlags And cdlSFBrowseForPrinter) <> 0 Then
        bInf.ulFlags = bInf.ulFlags And Not cdlSFBrowseForPrinter
        If bInf.ulFlags = 0 Then
            bInf.ulFlags = cdlSFBrowseForPrinter Or cdlSFNewDialogStyle Or cdlSFNoNewFolderButton Or cdlSFShareable
        Else
            bInf.ulFlags = bInf.ulFlags Or cdlSFBrowseForPrinter Or cdlSFNoNewFolderButton
        End If
        iCSIDL = CSIDL_PRINTERS
        If SHGetFolderLocation(0&, iCSIDL, 0&, 0&, iPidlRoot) = S_OK Then
            bInf.pIDLRoot = iPidlRoot.mkid.cb
        End If
    End If
    If (bInf.ulFlags And cdlSFBrowseForComputer) <> 0 Then
        bInf.ulFlags = bInf.ulFlags And Not cdlSFBrowseForComputer
        If bInf.ulFlags = 0 Then
            bInf.ulFlags = cdlSFBrowseForComputer Or cdlSFNoNewFolderButton Or cdlSFShareable
        Else
            bInf.ulFlags = bInf.ulFlags Or cdlSFBrowseForComputer Or cdlSFNoNewFolderButton
        End If
    End If
    
    If mInitDir <> "" Then
       gCommonDialogEx_ShowFolder_StartFolder = mInitDir & vbNullChar
    Else
        If mFolderName <> "" Then
            iAuxFolderPath = mFolderName
            gCommonDialogEx_ShowFolder_StartFolder = iAuxFolderPath
        End If
   End If
    bInf.lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc) 'get address of function.
    
    mCanceled = False
    mChanged = False
    iFolder_Prev = mFolderName
    
    'Show the Browse For Folder dialog
    PathID = SHBrowseForFolder(bInf)
    If PathID <> 0 Then
        RetPath = Space$(512)
        RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
        
        If RetVal Then
             'Trim$ off the null chars ending the path
             'and display the returned folder
             Offset = InStr(RetPath, Chr$(0))
             mFolderName = Left$(RetPath, Offset - 1)
             mFolderDisplayName = Left$(bInf.pszDisplayName, InStr(bInf.pszDisplayName, Chr(0)) - 1)
             'Free memory allocated for PIDL
             CoTaskMemFree PathID
        Else
            mFolderDisplayName = Left$(bInf.pszDisplayName, InStr(bInf.pszDisplayName, Chr(0)) - 1)
            If (bInf.ulFlags And cdlSFBrowseForPrinter) <> 0 Then
                If Not IsPrinter(mFolderDisplayName) Then
                    mFolderDisplayName = ""
                End If
            End If
            mFolderName = mFolderDisplayName
        End If
    Else
        mCanceled = True
        If mCancelError Then
            'generate an error
            On Error GoTo 0
            Err.Raise 32755, TypeName(Me), "Cancel Pressed"
            Exit Sub
        End If
    End If
    mChanged = mFolderName <> iFolder_Prev
End Sub

Private Function GetAddressofFunction(Add As Long) As Long
    GetAddressofFunction = Add
End Function


Private Function IsPrinter(nName As String) As Boolean
    Dim iPrn As Printer
    
    For Each iPrn In Printers
        If LCase$(iPrn.DeviceName) = LCase$(nName) Then
            IsPrinter = True
            Exit For
        End If
    Next iPrn
    
End Function

Public Property Let AmbientUserMode(ByVal nValue As Boolean)
Attribute AmbientUserMode.VB_MemberFlags = "40"
    mAmbientUserMode = nValue
End Property

Public Property Get AmbientUserMode() As Boolean
    AmbientUserMode = mAmbientUserMode
End Property

Friend Sub UpdateDevModeWithCurrentSettings()
    
'    Debug.Print mDevMode.dmDeviceName, mDeviceName & Chr(0)
    mDevMode.dmPaperSize = mPaperSize
    mDevMode.dmDefaultSource = mPaperBin
    mDevMode.dmDuplex = mDuplex
    mDevMode.dmOrientation = mOrientation
    mDevMode.dmPrintQuality = mPrintQuality
    mDevMode.dmColor = mColorMode
    mDevMode.dmCopies = mCopies
    mDevMode.dmCollate = Abs(CLng(mCollate))
    
    CopyMemory ByVal mDevModePtr, mDevMode, Len(mDevMode)
End Sub

Private Sub ConvertMarginValues()
    Dim iMultiplier As Single
    
    If mUnits = vbMillimeters Then
        iMultiplier = 25.4
    Else
        iMultiplier = 1 / 25.4
    End If
    mLeftMargin = mLeftMargin * iMultiplier
    mRightMargin = mRightMargin * iMultiplier
    mTopMargin = mTopMargin * iMultiplier
    mBottomMargin = mBottomMargin * iMultiplier
    mMinLeftMargin = mMinLeftMargin * iMultiplier
    mMinRightMargin = mMinRightMargin * iMultiplier
    mMinTopMargin = mMinTopMargin * iMultiplier
    mMinBottomMargin = mMinBottomMargin * iMultiplier
    
End Sub


Public Property Set Font(ByVal nFont As StdFont)
    If nFont Is Nothing Then Exit Property
    Set mFont = CloneFont(nFont)
    SetFontProperties
End Property

Public Property Get Font() As StdFont
    Set Font = CloneFont(mFont)
End Property

Private Sub SetFontProperties()
    mFontName = mFont.Name
    mFontSize = mFont.Size
    mFontBold = mFont.Bold
    mFontItalic = mFont.Italic
    mFontStrikethrough = mFont.Strikethrough
    mFontUnderline = mFont.Underline
End Sub

Private Sub PutPaperSize()
    Dim iPs As POINTAPI
    
    iPs = GetPaperSize(mPaperSize)
    
    If iPs.x = 0 Then
        mPaperWidth = mDefaultPaperWidth
        mPaperHeight = mDefaultPaperHeight
    Else
        mPaperWidth = iPs.x
        mPaperHeight = iPs.y
    End If
    
End Sub

Private Function GetPaperSize(nPaperSizeNumber As Long) As POINTAPI
    Dim ret As Long
    Dim iPaperSizesNumbers() As Integer
    Dim iPaperSizes() As POINTAPI
    Dim c As Long
    Dim iLng As Long
    
    ret = DeviceCapabilities(mDeviceName, mPort, DC_PAPERS, ByVal 0&, ByVal 0&)
    If ret = -1 Then Exit Function
    
    ReDim iPaperSizesNumbers(1 To ret)
    ReDim iPaperSizes(1 To ret)
    
    Call DeviceCapabilities(mDeviceName, mPort, DC_PAPERS, iPaperSizesNumbers(1), ByVal 0&)
    Call DeviceCapabilities(mDeviceName, mPort, DC_PAPERSIZE, iPaperSizes(1), ByVal 0&)
    
    For c = 1 To UBound(iPaperSizesNumbers)
        If iPaperSizesNumbers(c) = nPaperSizeNumber Then
            GetPaperSize.x = iPaperSizes(c).x
            GetPaperSize.y = iPaperSizes(c).y
            If GetPaperSize.x > GetPaperSize.y Then
                iLng = GetPaperSize.x
                GetPaperSize.x = GetPaperSize.y
                GetPaperSize.y = iLng
            End If
            Exit Function
        End If
    Next c
End Function

Private Function PaperSizeExists(nPaperSizeNumber As Long) As Boolean
    Dim ret As Long
    Dim iPaperSizesNumbers() As Integer
    Dim c As Long
    
    ret = DeviceCapabilities(mDeviceName, mPort, DC_PAPERS, ByVal 0&, ByVal 0&)
    If ret = -1 Then
        PaperSizeExists = True 'won't tell that it doesn't exists because we don't know
        Exit Function
    End If
    
    ReDim iPaperSizesNumbers(1 To ret)
    
    Call DeviceCapabilities(mDeviceName, mPort, DC_PAPERS, iPaperSizesNumbers(1), ByVal 0&)
    
    For c = 1 To UBound(iPaperSizesNumbers)
        If iPaperSizesNumbers(c) = nPaperSizeNumber Then
            PaperSizeExists = True
            Exit Function
        End If
    Next c
End Function

Private Function FindMostSimilarPaperSize(nPaperSize) As Long
    Dim iPaperWidth As Long
    Dim iPaperHeight As Long
    Dim iAuxPrinterEx As Object
    Dim iAuxReset As Boolean
    
    If Not mStandardPaperSizesLoaded Then
        LoadStandadPaperSizes
    End If
    
    If (nPaperSize >= LBound(mStandardPaperSizesNumbers)) And (nPaperSize <= UBound(mStandardPaperSizesNumbers)) Then
        iPaperWidth = mStandardPaperSizes(nPaperSize).x
        iPaperHeight = mStandardPaperSizes(nPaperSize).y
        
        Dim ret As Long
        Dim iPaperSizesNumbers() As Integer
        Dim iPaperSizes() As POINTAPI
        Dim c As Long
        
        ret = DeviceCapabilities(mDeviceName, mPort, DC_PAPERS, ByVal 0&, ByVal 0&)
        If ret = -1 Then Exit Function
        
        ReDim iPaperSizesNumbers(1 To ret)
        ReDim iPaperSizes(1 To ret)
        
        Call DeviceCapabilities(mDeviceName, mPort, DC_PAPERS, iPaperSizesNumbers(1), ByVal 0&)
        Call DeviceCapabilities(mDeviceName, mPort, DC_PAPERSIZE, iPaperSizes(1), ByVal 0&)
        
        For c = 1 To UBound(iPaperSizesNumbers)
            If (iPaperSizes(c).x = iPaperWidth) And (iPaperSizes(c).y = iPaperHeight) Then
                FindMostSimilarPaperSize = iPaperSizesNumbers(c)
                Exit For
            End If
        Next c
        
        If FindMostSimilarPaperSize = 0 Then
            For c = 1 To UBound(iPaperSizesNumbers)
                If (Abs(iPaperSizes(c).x - iPaperWidth) < 2) And (Abs(iPaperSizes(c).y - iPaperHeight) < 2) Then
                    FindMostSimilarPaperSize = iPaperSizesNumbers(c)
                    Exit For
                End If
            Next c
        End If
        
        If FindMostSimilarPaperSize = 0 Then
            For c = 1 To UBound(iPaperSizesNumbers)
                If (Abs(iPaperSizes(c).x - iPaperWidth) < 3) And (Abs(iPaperSizes(c).y - iPaperHeight) < 3) Then
                    FindMostSimilarPaperSize = iPaperSizesNumbers(c)
                    Exit For
                End If
            Next c
        End If
        
        If FindMostSimilarPaperSize = 0 Then
            For c = 1 To UBound(iPaperSizesNumbers)
                If (Abs(iPaperSizes(c).x - iPaperWidth) < 4) And (Abs(iPaperSizes(c).y - iPaperHeight) < 4) Then
                    FindMostSimilarPaperSize = iPaperSizesNumbers(c)
                    Exit For
                End If
            Next c
        End If
        
        If FindMostSimilarPaperSize = 0 Then
            For c = 1 To UBound(iPaperSizesNumbers)
                If (Abs(iPaperSizes(c).x - iPaperWidth) < 12) And (Abs(iPaperSizes(c).y - iPaperHeight) < 12) Then
                    FindMostSimilarPaperSize = iPaperSizesNumbers(c)
                    Exit For
                End If
            Next c
        End If
        
        If FindMostSimilarPaperSize = 0 Then
            For c = 1 To UBound(iPaperSizesNumbers)
                If (Abs(iPaperSizes(c).x - iPaperWidth) < 30) And (Abs(iPaperSizes(c).y - iPaperHeight) < 30) Then
                    FindMostSimilarPaperSize = iPaperSizesNumbers(c)
                    Exit For
                End If
            Next c
        End If
        
        If FindMostSimilarPaperSize = 0 Then
            For c = 1 To UBound(iPaperSizesNumbers)
                If (Abs(iPaperSizes(c).x - iPaperWidth) < 100) And (Abs(iPaperSizes(c).y - iPaperHeight) < 100) Then
                    FindMostSimilarPaperSize = iPaperSizesNumbers(c)
                    Exit For
                End If
            Next c
        End If
        
        If FindMostSimilarPaperSize = 0 Then
            If PrinterExCurrentDocument Is Nothing Then
                Set iAuxPrinterEx = Printer2
                iAuxReset = True
            End If
            FindMostSimilarPaperSize = PrinterExCurrentDocument.PaperSize
            If iAuxReset Then
                ResetPrinter2
            End If
        End If
    Else
        If PrinterExCurrentDocument Is Nothing Then
            Set iAuxPrinterEx = Printer2
            iAuxReset = True
        End If
        FindMostSimilarPaperSize = PrinterExCurrentDocument.PaperSize
        If iAuxReset Then
            ResetPrinter2
        End If
    End If
End Function

Private Sub LoadStandadPaperSizes()
    If mStandardPaperSizesLoaded Then Exit Sub
    
    ReDim mStandardPaperSizesNumbers(1 To 41)
    ReDim mStandardPaperSizes(1 To 41)
    
    mStandardPaperSizesNumbers(1) = vbPRPSLetter: mStandardPaperSizes(1).x = 2159: mStandardPaperSizes(1).y = 2794
    mStandardPaperSizesNumbers(2) = vbPRPSLetterSmall: mStandardPaperSizes(2).x = 2159: mStandardPaperSizes(2).y = 2794
    mStandardPaperSizesNumbers(3) = vbPRPSTabloid: mStandardPaperSizes(3).x = 2794: mStandardPaperSizes(3).y = 4318
    mStandardPaperSizesNumbers(4) = vbPRPSLedger: mStandardPaperSizes(4).x = 4318: mStandardPaperSizes(4).y = 2794
    mStandardPaperSizesNumbers(5) = vbPRPSLegal: mStandardPaperSizes(5).x = 2159: mStandardPaperSizes(5).y = 3556
    mStandardPaperSizesNumbers(6) = vbPRPSStatement: mStandardPaperSizes(6).x = 1397: mStandardPaperSizes(6).y = 2159
    mStandardPaperSizesNumbers(7) = vbPRPSExecutive: mStandardPaperSizes(7).x = 1841: mStandardPaperSizes(7).y = 2667
    mStandardPaperSizesNumbers(8) = vbPRPSA3: mStandardPaperSizes(8).x = 2970: mStandardPaperSizes(8).y = 4200
    mStandardPaperSizesNumbers(9) = vbPRPSA4: mStandardPaperSizes(9).x = 2100: mStandardPaperSizes(9).y = 2970
    mStandardPaperSizesNumbers(10) = vbPRPSA4Small: mStandardPaperSizes(10).x = 2100: mStandardPaperSizes(10).y = 2970
    mStandardPaperSizesNumbers(11) = vbPRPSA5: mStandardPaperSizes(11).x = 1480: mStandardPaperSizes(11).y = 2100
    mStandardPaperSizesNumbers(12) = vbPRPSB4: mStandardPaperSizes(12).x = 2570: mStandardPaperSizes(12).y = 3640
    mStandardPaperSizesNumbers(13) = vbPRPSB5: mStandardPaperSizes(13).x = 1820: mStandardPaperSizes(13).y = 2570
    mStandardPaperSizesNumbers(14) = vbPRPSFolio: mStandardPaperSizes(14).x = 2159: mStandardPaperSizes(14).y = 3302
    mStandardPaperSizesNumbers(15) = vbPRPSQuarto: mStandardPaperSizes(15).x = 2150: mStandardPaperSizes(15).y = 2750
    mStandardPaperSizesNumbers(16) = vbPRPS10x14: mStandardPaperSizes(16).x = 2540: mStandardPaperSizes(16).y = 3556
    mStandardPaperSizesNumbers(17) = vbPRPS11x17: mStandardPaperSizes(17).x = 2794: mStandardPaperSizes(17).y = 4318
    mStandardPaperSizesNumbers(18) = vbPRPSNote: mStandardPaperSizes(18).x = 2159: mStandardPaperSizes(18).y = 2794
    mStandardPaperSizesNumbers(19) = vbPRPSEnv9: mStandardPaperSizes(19).x = 984: mStandardPaperSizes(19).y = 2254
    mStandardPaperSizesNumbers(20) = vbPRPSEnv10: mStandardPaperSizes(20).x = 1047: mStandardPaperSizes(20).y = 2413
    mStandardPaperSizesNumbers(21) = vbPRPSEnv11: mStandardPaperSizes(21).x = 1143: mStandardPaperSizes(21).y = 2635
    mStandardPaperSizesNumbers(22) = vbPRPSEnv12: mStandardPaperSizes(22).x = 1206: mStandardPaperSizes(22).y = 2794
    mStandardPaperSizesNumbers(23) = vbPRPSEnv14: mStandardPaperSizes(23).x = 1270: mStandardPaperSizes(23).y = 2921
    mStandardPaperSizesNumbers(24) = vbPRPSCSheet: mStandardPaperSizes(24).x = 4318: mStandardPaperSizes(24).y = 5588
    mStandardPaperSizesNumbers(25) = vbPRPSDSheet: mStandardPaperSizes(25).x = 5588: mStandardPaperSizes(25).y = 8636
    mStandardPaperSizesNumbers(26) = vbPRPSESheet: mStandardPaperSizes(26).x = 8636: mStandardPaperSizes(26).y = 11176
    mStandardPaperSizesNumbers(27) = vbPRPSEnvDL: mStandardPaperSizes(27).x = 1100: mStandardPaperSizes(27).y = 2200
    mStandardPaperSizesNumbers(28) = vbPRPSEnvC3: mStandardPaperSizes(28).x = 1620: mStandardPaperSizes(28).y = 2290
    mStandardPaperSizesNumbers(29) = vbPRPSEnvC4: mStandardPaperSizes(29).x = 3240: mStandardPaperSizes(29).y = 4580
    mStandardPaperSizesNumbers(30) = vbPRPSEnvC5: mStandardPaperSizes(30).x = 2290: mStandardPaperSizes(30).y = 3240
    mStandardPaperSizesNumbers(31) = vbPRPSEnvC6: mStandardPaperSizes(31).x = 1140: mStandardPaperSizes(31).y = 1620
    mStandardPaperSizesNumbers(32) = vbPRPSEnvC65: mStandardPaperSizes(32).x = 1140: mStandardPaperSizes(32).y = 2290
    mStandardPaperSizesNumbers(33) = vbPRPSEnvB4: mStandardPaperSizes(33).x = 2500: mStandardPaperSizes(33).y = 3530
    mStandardPaperSizesNumbers(34) = vbPRPSEnvB5: mStandardPaperSizes(34).x = 1760: mStandardPaperSizes(34).y = 2500
    mStandardPaperSizesNumbers(35) = vbPRPSEnvB6: mStandardPaperSizes(35).x = 1760: mStandardPaperSizes(35).y = 1250
    mStandardPaperSizesNumbers(36) = vbPRPSEnvItaly: mStandardPaperSizes(36).x = 1100: mStandardPaperSizes(36).y = 2300
    mStandardPaperSizesNumbers(37) = vbPRPSEnvMonarch: mStandardPaperSizes(37).x = 984: mStandardPaperSizes(37).y = 1905
    mStandardPaperSizesNumbers(38) = vbPRPSEnvPersonal: mStandardPaperSizes(38).x = 920: mStandardPaperSizes(38).y = 1651
    mStandardPaperSizesNumbers(39) = vbPRPSFanfoldUS: mStandardPaperSizes(39).x = 3778: mStandardPaperSizes(39).y = 2794
    mStandardPaperSizesNumbers(40) = vbPRPSFanfoldStdGerman: mStandardPaperSizes(40).x = 2159: mStandardPaperSizes(40).y = 3048
    mStandardPaperSizesNumbers(41) = vbPRPSFanfoldLglGerman: mStandardPaperSizes(41).x = 2159: mStandardPaperSizes(41).y = 3302
    
    mStandardPaperSizesLoaded = True
End Sub

Public Property Get DevModePtr() As Long
    DevModePtr = mDevModePtr
End Property


Public Property Let CustomColors(Index As Integer, nValue As Long)
    If (Index < 0) Or (Index > 15) Then Exit Property
    mCustomColors(Index) = nValue
End Property

Public Property Get CustomColors(Index As Integer) As Long
    CustomColors = mCustomColors(Index)
End Property


Public Property Let AutoSaveCustomColors(ByVal nValue As Boolean)
    If mAutoSaveCustomColors <> nValue Then
        mAutoSaveCustomColors = nValue
    End If
End Property

Public Property Get AutoSaveCustomColors() As Boolean
    AutoSaveCustomColors = mAutoSaveCustomColors
End Property

Private Sub InitCustomColors()
    Dim c As Long
    Dim iByte As Byte
    
    For c = 0 To 15
        iByte = 255 - c * 16
        mCustomColors(c) = RGB(iByte, iByte, iByte)
    Next c
End Sub

Private Sub LoadCustomColors()
    Dim c As Long
    
    For c = 0 To 15
        mCustomColors(c) = Val(GetSetting(AppNameForRegistry, "CommonDialogEx", "CustomColors" & CStr(c), mCustomColors(c)))
    Next c
    mCustomColorsLoaded = True
End Sub

Private Sub SaveCustomColors()
    Dim c As Long
    
    For c = 0 To 15
        SaveSetting AppNameForRegistry, "CommonDialogEx", "CustomColors" & CStr(c), CStr(mCustomColors(c))
    Next c
End Sub


Private Function GetActiveWindowHwnd() As Long
    GetActiveWindowHwnd = GetActiveFormHwnd
    If GetActiveWindowHwnd = 0 Then
        GetActiveWindowHwnd = GetForegroundWindow
        If GetWindowThreadProcessId(GetActiveWindowHwnd, 0&) <> App.ThreadID Then
            GetActiveWindowHwnd = 0
        End If
    End If
End Function

Private Function GetDevModeByDeviceName(ByVal nDeviceName As String, ByRef nDevMode As DEVMODE) As Boolean
    Dim hPrinter As Long
    Dim pd As PRINTER_DEFAULTS
    Dim yDevModeData() As Byte
    Dim nRet As Long
    
    On Error GoTo cleanup
    
    pd.DesiredAccess = PRINTER_ALL_ACCESS
    nRet = OpenPrinter(nDeviceName, hPrinter, pd)
    If (nRet = 0) Or (hPrinter = 0) Then
       If Err.LastDllError = 5 Then
          ' Access denied.
       Else
          ' Cannot open the printer specified. The printer name may be incorrect.
       End If
       Exit Function
    End If
    
    nRet = DocumentProperties(0, hPrinter, nDeviceName, 0, 0, 0)
    If (nRet < 0) Then
       ' Cannot get the size of the DEVMODE structure.
       GoTo cleanup
    End If
    
    ReDim yDevModeData(nRet + 100) As Byte
    nRet = DocumentProperties(0, hPrinter, nDeviceName, VarPtr(yDevModeData(0)), 0, DM_OUT_BUFFER)
    If (nRet < 0) Then
       ' Cannot get the DEVMODE structure.
       GoTo cleanup
    End If
    
    Call CopyMemoryAnyLong(nDevMode, yDevModeData(0), Len(nDevMode))
    GetDevModeByDeviceName = True
    
cleanup:
    If (hPrinter <> 0) Then Call ClosePrinter(hPrinter)

End Function

Public Property Get Changed() As Boolean
    Changed = mChanged
End Property

