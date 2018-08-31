VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Global"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

'[Enumerations]

Public Enum vbExCompareMethod
    vbBinaryCompare = 0
    vbTextCompare = 1
    vbDatabaseCompare = 2
End Enum

Public Enum vbExGUILanguageConstants
    vxLangAUTO = 0
    vxLangENGLISH = &H9
    vxLangSPANISH = &HA
    vxLangCHINESE_SIMPLIFIED = &H4

'    vxLangAFRIKAANS = &H36
'    vxLangALBANIAN = &H1C
'    vxLangALSATIAN = &H84
'    vxLangAMHARIC = &H5E
'    vxLangARABIC = &H1
'    vxLangARMENIAN = &H2B
'    vxLangASSAMESE = &H4D
'    vxLangAZERI = &H2C
'    vxLangBASHKIR = &H6D
'    vxLangBASQUE = &H2D
'    vxLangBELARUSIAN = &H23
'    vxLangBENGALI = &H45
'    vxLangBRETON = &H7E
'    vxLangBOSNIAN = &H1A
'    vxLangBOSNIAN_NEUTRAL = &H781A
'    vxLangBULGARIAN = &H2
'    vxLangCATALAN = &H3
'    vxLangCHINESE_TRADITIONAL = &H7C04
'    vxLangCORSICAN = &H83
'    vxLangCROATIAN = &H1A
'    vxLangCZECH = &H5
'    vxLangDANISH = &H6
'    vxLangDARI = &H8C
'    vxLangDIVEHI = &H65
'    vxLangDUTCH = &H13
'    vxLangESTONIAN = &H25
'    vxLangFAEROESE = &H38
'    vxLangFARSI = &H29
'    vxLangFILIPINO = &H64
'    vxLangFINNISH = &HB
'    vxLangFRENCH = &HC
'    vxLangFRISIAN = &H62
'    vxLangGALICIAN = &H56
'    vxLangGEORGIAN = &H37
'    vxLangGERMAN = &H7
'    vxLangGREEK = &H8
'    vxLangGREENLANDIC = &H6F
'    vxLangGUJARATI = &H47
'    vxLangHAUSA = &H68
'    vxLangHEBREW = &HD
'    vxLangHINDI = &H39
'    vxLangHUNGARIAN = &HE
'    vxLangICELANDIC = &HF
'    vxLangIGBO = &H70
'    vxLangINDONESIAN = &H21
'    vxLangINUKTITUT = &H5D
'    vxLangIRISH = &H3C
'    vxLangITALIAN = &H10
'    vxLangJAPANESE = &H11
'    vxLangKANNADA = &H4B
'    vxLangKASHMIRI = &H60
'    vxLangKAZAK = &H3F
'    vxLangKHMER = &H53
'    vxLangKICHE = &H86
'    vxLangKINYARWANDA = &H87
'    vxLangKONKANI = &H57
'    vxLangKOREAN = &H12
'    vxLangKYRGYZ = &H40
'    vxLangLAO = &H54
'    vxLangLATVIAN = &H26
'    vxLangLITHUANIAN = &H27
'    vxLangLOWER_SORBIAN = &H2E
'    vxLangLUXEMBOURGISH = &H6E
'    vxLangMACEDONIAN = &H2F
'    vxLangMALAY = &H3E
'    vxLangMALAYALAM = &H4C
'    vxLangMALTESE = &H3A
'    vxLangMANIPURI = &H58
'    vxLangMAORI = &H81
'    vxLangMAPUDUNGUN = &H7A
'    vxLangMARATHI = &H4E
'    vxLangMOHAWK = &H7C
'    vxLangMONGOLIAN = &H50
'    vxLangNEPALI = &H61
'    vxLangNORWEGIAN = &H14
'    vxLangOCCITAN = &H82
'    vxLangORIYA = &H48
'    vxLangPASHTO = &H63
'    vxLangPERSIAN = &H29
'    vxLangPOLISH = &H15
'    vxLangPORTUGUESE = &H16
'    vxLangPUNJABI = &H46
'    vxLangQUECHUA = &H6B
'    vxLangROMANIAN = &H18
'    vxLangROMANSH = &H17
'    vxLangRUSSIAN = &H19
'    vxLangSAMI = &H3B
'    vxLangSANSKRIT = &H4F
'    vxLangSCOTTISH_GAELIC = &H91
'    vxLangSERBIAN = &H1A
'    vxLangSERBIAN_NEUTRAL = &H7C1A
'    vxLangSINDHI = &H59
'    vxLangSINHALESE = &H5B
'    vxLangSLOVAK = &H1B
'    vxLangSLOVENIAN = &H24
'    vxLangSOTHO = &H6C
'    vxLangSWAHILI = &H41
'    vxLangSWEDISH = &H1D
'    vxLangSYRIAC = &H5A
'    vxLangTAJIK = &H28
'    vxLangTAMAZIGHT = &H5F
'    vxLangTAMIL = &H49
'    vxLangTATAR = &H44
'    vxLangTELUGU = &H4A
'    vxLangTHAI = &H1E
'    vxLangTIBETAN = &H51
'    vxLangTIGRIGNA = &H73
'    vxLangTSWANA = &H32
'    vxLangTURKISH = &H1F
'    vxLangTURKMEN = &H42
'    vxLangUIGHUR = &H80
'    vxLangUKRAINIAN = &H22
'    vxLangUPPER_SORBIAN = &H2E
'    vxLangURDU = &H20
'    vxLangUZBEK = &H43
'    vxLangVIETNAMESE = &H2A
'    vxLangWELSH = &H52
'    vxLangWOLOF = &H88
'    vxLangXHOSA = &H34
'    vxLangYAKUT = &H85
'    vxLangYI = &H78
'    vxLangYORUBA = &H6A
'    vxLangZULU = &H35
End Enum

Public Enum vbExSystemFontConstants
   vxIconFont = 1
   vxCaptionFont = 2
   vxSmallCaptionFont = 3
   vxMenuFont = 4
   vxStatusAndTooltipFont = 5
   vxMsgBoxFont = 6
End Enum

Public Enum vbExDataTypeConstants
    vxText = 0
    vxBinary = 1
End Enum

Public Enum vbExMouseButtonsConstants
    vxMBLeft = 1
    vxMBRight = 2
End Enum

Public Enum vbExToolbarDAIconsSizeConstants
    vxIconsAppDefault = 0
    vxIconsSmall = 1
    vxIconsMedium = 2
    vxIconsBig = 3
End Enum

Public Enum vbExPPOpCodeConstants
    vbDstInvert = &H550009
    vbMergeCopy = &HC000CA
    vbMergePaint = &HBB0226
    vbNotSrcCopy = &H330008
    vbNotSrcErase = &H1100A6
    vbPatCopy = &HF00021
    vbPatInvert = &H5A0049
    vbPatPaint = &HFB0A09
    vbSrcAnd = &H8800C6
    vbSrcCopy = &HCC0020
    vbSrcErase = &H440328
    vbSrcInvert = &H660046
    vbSrcPaint = &HEE0086
    
    ' extended codes (not raster)
    vbSMHalfTone = 0 ' now default
    vbSMBlackOnWhite = &H1000000
    vbSMVB6Default = vbSMBlackOnWhite
    vbSMWhiteOnBlack = &H2000000
    vbSMColorOnColor = &H3000000
    vbSMQualityStretch = &H5000000
    vbTransparentColor = &H8000000
End Enum

Public Enum vbExPrintExOptionsConstants
    vxLineFeed = 0
    vxNoLineFeed = 1
    vxAddTab = 2
End Enum

Public Enum vbExPrintExAlignConstants
    vxAlignLeftTop = 0
    vxAlignLeftCenter = 1
    vxAlignLeftBottom = 2
    vxAlignCenterTop = 3
    vxAlignCenterCenter = 4
    vxAlignCenterBottom = 5
    vxAlignRightTop = 6
    vxAlignRightCenter = 7
    vxAlignRightBottom = 8
End Enum


Private Sub Class_Initialize()
    InitGlobal
End Sub


Public Sub PersistForm(nForm As Object, nForms As Object, Optional nInitialCentered As Boolean = True, Optional nInitialLeft, Optional nInitialTop, Optional nInitialWidth, Optional nInitialHeight, Optional nPersistLeft As Boolean = True, Optional nPersistTop As Boolean = True, Optional nPersistWidth As Boolean = True, Optional nPersistHeight As Boolean = True, Optional nMaxTop, Optional nPersistMinimizedState As Boolean, Optional nContext As String)
    mGlobals.PersistForm nForm, nForms, nInitialCentered, nInitialLeft, nInitialTop, nInitialWidth, nInitialHeight, nPersistLeft, nPersistTop, nPersistWidth, nPersistHeight, nMaxTop, nPersistMinimizedState, nContext
End Sub

Public Sub UnpersistForm(nForm As Object)
    mGlobals.UnpersistForm nForm
End Sub

Public Sub SaveFormPersistance(nForm As Object)
    mGlobals.SaveFormPersistance nForm
End Sub

Public Function GetFormPersistedWindowState(nForm As Object) As Long
    GetFormPersistedWindowState = mGlobals.GetFormPersistedWindowState(nForm)
End Function

' Kept for keeping binary compatibility, it will be removed in the next version
Public Function GetSavedFormWindowState(nForm As Object, Optional nContext As String, Optional nDefaultValue As Long = -1) As FormWindowStateConstants
    
End Function

Public Sub ShowModal(nForm As Object, Optional nWaitWithDoevents As Boolean = True, Optional nSetIcon As Boolean = True, Optional nFormsHwndToKeepEnabled As Variant, Optional nKeepEnabledTaskBarWindows As Boolean = True, Optional nNoOwner As Boolean)
    mGlobals.ShowModal nForm, nWaitWithDoevents, nSetIcon, nFormsHwndToKeepEnabled, nKeepEnabledTaskBarWindows, nNoOwner
End Sub

Public Sub ShowNoActivate(nForm As Object, Optional nOwnerForm, Optional nSetIcon As Boolean = True, Optional nSetActiveFormAsOwner As Boolean)
    mGlobals.ShowNoActivate nForm, nOwnerForm, nSetIcon, nSetActiveFormAsOwner
End Sub

Public Function SetMinMax(nForm As Object, Optional nMinWidth, Optional nMinHeight, Optional nMaxWidth, Optional nMaxHeight, Optional ScaleMode As Integer = vbTwips) As FormMinMax
    Set SetMinMax = mGlobals.SetMinMax(nForm, nMinWidth, nMinHeight, nMaxWidth, nMaxHeight, ScaleMode)
End Function

Public Property Get Printer2() As Object
    Set Printer2 = mPrinterEx.Printer2
End Property

Public Property Set Printer2(nPrinter As Object)
    mPrinterEx.Printer2 = nPrinter
End Property

Public Property Get PrinterEx() As IPrinterEx
    Set PrinterEx = mPrinterEx.PrinterEx
End Property

' This method is hidden
Public Sub ResetPrinterEx()
Attribute ResetPrinterEx.VB_MemberFlags = "40"
    ResetPrinter2
End Sub

Public Property Get ScreenActiveForm(nForms As Object, Optional nNotOwned As Boolean) As Object
    Set ScreenActiveForm = mGlobals.ScreenActiveForm(nForms, nNotOwned)
End Property
    
Public Function ActivatePrevInstance(Optional nTimeoutCheckHangedSeconds As Long = 25, Optional nTerminateHanged As Boolean = True, Optional nWaitForHavingWindowsTimeoutSeconds As Long = 25, Optional nTerminateWindowless As Boolean = True) As Boolean
    Dim iAP As New cActivatePrevInstance
    
    If Not ClientAppIsCompiled Then Exit Function
    
    If iAP.PrevInstance Then
        ActivatePrevInstance = iAP.PrevInstanceActivated(nTimeoutCheckHangedSeconds, nTerminateHanged, nWaitForHavingWindowsTimeoutSeconds, nTerminateWindowless)
    End If
End Function

Public Function AppPrevInstance() As Boolean
    Dim iAP As New cActivatePrevInstance
    
    If Not ClientAppIsCompiled Then Exit Function
    
    AppPrevInstance = iAP.PrevInstance
End Function

Public Function IsAppRunning(nAppPath As String) As Boolean
    Dim iAP As New cActivatePrevInstance
    
    IsAppRunning = iAP.IsAppRunning(nAppPath)
End Function

' Kept for keeping binary compatibility, it will be removed in the next version
Public Function IsProgramRunning(nAppPath As String) As Boolean

End Function

Public Function MinimizeApp() As AppMinimizer
    Set MinimizeApp = New AppMinimizer
End Function

Public Property Get ScreenUsableHeight() As Single
    ScreenUsableHeight = mGlobals.ScreenUsableHeight
End Property

Public Function IsShowingVerticalScrollBar(nControl As Object) As Boolean
    IsShowingVerticalScrollBar = mGlobals.IsShowingVerticalScrollBar(nControl)
End Function

Public Function IsShowingHorizontalScrollBar(nControl As Object) As Boolean
    IsShowingHorizontalScrollBar = mGlobals.IsShowingHorizontalScrollBar(nControl)
End Function

Public Function FileExists(ByVal nPathName As String) As Boolean
    FileExists = mGlobals.FileExists(nPathName)
End Function

Public Function GetTempFolder() As String
    GetTempFolder = mGlobals.GetTempFolder
End Function

' Kept for keeping binary compatibility, it will be removed in the next version
Public Property Get TempFolder() As String
    
End Property

' Kept for keeping binary compatibility, it will be removed in the next version
Public Function GetTempDir() As String
    
End Function

Public Function GetTempFileFullPath() As String
    GetTempFileFullPath = mGlobals.GetTempFileFullPath
End Function

Public Function ControlNameWithParent(nControl As Object) As String
    ControlNameWithParent = mGlobals.ControlNameWithParent(nControl)
End Function

' adds an String to the Clipboard keeping the text that it already has
Public Sub ToClipboard(nText As String)
    mGlobals.ToClipboard nText
End Sub

Public Function IsWindowVisibleOnScreen(nHwnd As Long, Optional AtLeastPartially As Boolean) As Boolean
    IsWindowVisibleOnScreen = mGlobals.IsWindowVisibleOnScreen(nHwnd, AtLeastPartially)
End Function

Public Sub SetWindowRedraw(nHwnd As Long, nRedraw As Boolean, Optional nForce As Boolean)
    mGlobals.SetWindowRedraw nHwnd, nRedraw, nForce
End Sub

Public Function IsThemed() As Boolean
    IsThemed = mGlobals.IsThemed
End Function
    
Public Function IsWindowsThemed() As Boolean
    IsWindowsThemed = mGlobals.IsWindowsThemed
End Function

Public Function CommonControlsVersionLoaded() As Long
    CommonControlsVersionLoaded = mGlobals.CommonControlsVersionLoaded
End Function

Public Function IsWindowsNT() As Boolean
    IsWindowsNT = mGlobals.IsWindowsNT
End Function

Public Function IsWindows98OrMore() As Boolean
    IsWindows98OrMore = mGlobals.IsWindows98OrMore
End Function

Public Function IsWindows2000OrMore() As Boolean
    IsWindows2000OrMore = mGlobals.IsWindows2000OrMore
End Function

Public Function IsWindowsXPOrMore() As Boolean
    IsWindowsXPOrMore = mGlobals.IsWindowsXPOrMore
End Function

Public Function IsWindowsXP() As Boolean
    IsWindowsXP = mGlobals.IsWindowsXP
End Function

Public Function IsWindowsVistaOrMore() As Boolean
    IsWindowsVistaOrMore = mGlobals.IsWindowsVistaOrMore
End Function
    
Public Function IsWindows7OrMore() As Boolean
    IsWindows7OrMore = mGlobals.IsWindows7OrMore
End Function

Public Function IsWindows8OrMore() As Boolean
    IsWindows8OrMore = mGlobals.IsWindows8OrMore
End Function
    
Public Function IsWindowsServer() As Boolean
    IsWindowsServer = mGlobals.IsWindowsServer
End Function
    
Public Function IsWindows64Bit() As Boolean
    IsWindows64Bit = mGlobals.IsWindows64Bit
End Function

Public Function ThemeColor(nThemeID As String) As Long
    ThemeColor = mGlobals.ThemeColor(nThemeID)
End Function

Public Function ColorsBlended(ByVal nColor1 As Long, ByVal nColor2 As Long, ByVal nPercentColor2 As Long)
    ColorsBlended = mGlobals.ColorsBlended(nColor1, nColor2, nPercentColor2)
End Function

Public Function GetSystemFont(nSystemFont As vbExSystemFontConstants) As StdFont
    Set GetSystemFont = mGlobals.GetSystemFont(nSystemFont)
End Function

Public Sub SetTextBoxNumeric(nTextBox As Object)
    mGlobals.SetTextBoxNumeric nTextBox
End Sub

Public Function GetComboListHwnd(nCombo As Object) As Long
    GetComboListHwnd = mGlobals.GetComboListHwnd(nCombo)
End Function

Public Function GetComboEditHwnd(nCombo As Object) As Long
    GetComboEditHwnd = mGlobals.GetComboEditHwnd(nCombo)
End Function

Public Function HiByte(ByVal wParam As Integer) As Integer
    HiByte = (Abs(wParam) \ &H100) And &HFF&
End Function

Public Function LoByte(ByVal wParam As Integer) As Integer
    LoByte = Abs(wParam) And &HFF&
End Function

Public Function MakeWord(ByVal wLow As Integer, ByVal wHigh As Integer) As Integer

    If wHigh And &H80 Then
        MakeWord = (((wHigh And &H7F) * 256) + wLow) Or &H8000
    Else
        MakeWord = (wHigh * 256) + wLow
    End If
    
End Function

Public Function HiWord(ByVal dwValue As Long) As Long
    Call CopyMemory(HiWord, ByVal VarPtr(dwValue) + 2, 2)
End Function
  
Public Function LoWord(ByVal dwValue As Long) As Long
    Call CopyMemory(LoWord, dwValue, 2)
End Function

Public Function MakeLong(ByVal wLow As Long, ByVal wHi As Long) As Long

    If (wHi And &H8000&) Then
        MakeLong = (((wHi And &H7FFF&) * 65536) Or (wLow And &HFFFF&)) Or &H80000000
    Else
        MakeLong = LoWord(wLow) Or (&H10000 * LoWord(wHi))
    End If

End Function

Public Function GetFolder(nFilePath As String) As String
    GetFolder = mGlobals.GetFolder(nFilePath)
End Function

' Kept for keeping binary compatibility, it will be removed in the next version
Public Function FolderName(nFilePath As String) As String
    
End Function

' Kept for keeping binary compatibility, it will be removed in the next version
Public Function GetFileFolderPath(nFileFullPath As String) As String
    
End Function

Public Function FolderExists(nFolderPath As String) As Boolean
    FolderExists = mGlobals.FolderExists(nFolderPath)
End Function

' Kept for keeping binary compatibility, it will be removed in the next version
Public Function DirExists(strDirName As String) As Boolean
    
End Function

Public Sub EnsureFocusRect(nForm As Object)
    mGlobals.EnsureFocusRect nForm
End Sub

Public Function GetActiveFormHwnd(Optional nNotOwned As Boolean) As Long
    GetActiveFormHwnd = mGlobals.GetActiveFormHwnd(nNotOwned)
End Function

Public Function GetFormUnderMouseHwnd() As Long
    GetFormUnderMouseHwnd = mGlobals.GetFormUnderMouseHwnd
End Function

Public Function ShowToolTipEx(nTipText As String, Optional nTitle As String, Optional nStyle As vbExBalloonTooltipStyleConstants = vxTTBalloon, Optional nCloseButton As Boolean, Optional nIcon As vbExBalloonTooltipIconConstants = vxTTNoIcon, Optional nDelayTimeSeconds, Optional nVisibleTimeSeconds, Optional nPositionX, Optional nPositionY, Optional nPositionIsRelative As Boolean, Optional nWidth, Optional nBackColor, Optional nForeColor, Optional nClosePrevious As Boolean = True, Optional nRestrictMouseMoveToTwips As Long = 1000) As ToolTipEx
    'If False Then
    If CommonControlsVersionLoaded < 6 Then
        If nCloseButton Then
            If nTipText <> "" Then
                Set ShowToolTipEx = frmShowLittleHelp.ShowIt(nTipText)
            Else
                Unload frmShowLittleHelp
            End If
        End If
    Else
        Set ShowToolTipEx = mGlobals.ShowToolTipEx(nTipText, nTitle, nStyle, nCloseButton, nIcon, nDelayTimeSeconds, nVisibleTimeSeconds, nPositionX, nPositionY, nPositionIsRelative, nWidth, nBackColor, nForeColor, nClosePrevious, nRestrictMouseMoveToTwips)
    End If
End Function

Public Function GetProgramDocumentsFolder() As String
    GetProgramDocumentsFolder = mGlobals.GetProgramDocumentsFolder
End Function

Public Sub SaveProgramDocumentsFolder(nPath As String)
    mGlobals.SaveProgramDocumentsFolder nPath
End Sub

Public Function CheckFreeDiskSpace(nPathToTest As String, nRequiredFreeSpaceInBytes As Double) As Boolean
    CheckFreeDiskSpace = mGlobals.CheckFreeDiskSpace(nPathToTest, nRequiredFreeSpaceInBytes)
End Function

' Kept for keeping binary compatibility, it will be removed in the next version
Public Function CheckDiskFreeSpace(nPathToTest As String, nRequiredFreeSpaceInBytes As Double) As Boolean
    
End Function

' Kept for keeping binary compatibility, it will be removed in the next version
Public Function CheckDiskSpace(nPathToTest As String, nRequiredFreeSpaceInBytes As Double) As Boolean
    
End Function

Public Function GMTDateTime(nLocalDateTime As Date) As Date
    GMTDateTime = mGlobals.GMTDateTime(nLocalDateTime)
End Function

Public Function LocalDateTime(nGMTDateTime As Date) As Date
    LocalDateTime = mGlobals.LocalDateTime(nGMTDateTime)
End Function

Public Function GetLocalToGMTDifference() As Long
    GetLocalToGMTDifference = mGlobals.GetLocalToGMTDifference
End Function

Public Function Replace(Text As String, sOld As String, sNew As String, Optional ByVal Start As Long = 1, Optional ByVal Count As Long = 2147483647, Optional ByVal Compare As vbExCompareMethod = vbBinaryCompare) As String
    Replace = mVBFunctionsReplacement.Replace(Text, sOld, sNew, Start, Count, Compare)
End Function

Public Function InStrCount(ByRef sCheck As String, ByRef sMatch As String, Optional ByVal Start As Long = 1, Optional ByVal Compare As vbExCompareMethod = vbBinaryCompare) As Long
    InStrCount = mVBFunctionsReplacement.InStrCount(sCheck, sMatch, Start, Compare)
End Function

Public Function InStrRev(ByRef sCheck As String, ByRef sMatch As String, Optional ByVal Start As Long, Optional ByVal Compare As vbExCompareMethod = vbBinaryCompare) As Long
    InStrRev = mVBFunctionsReplacement.InStrRev(sCheck, sMatch, Start, Compare)
End Function

Public Function WordCount(s$) As Long
    WordCount = mVBFunctionsReplacement.WordCount(s$)
End Function

Public Function CloneFont(nOrigFont As Object) As Object
    Set CloneFont = mGlobals.CloneFont(nOrigFont)
End Function

Public Sub StartLogging(ByVal nPath As String)
    mGlobals.StartLogging nPath
End Sub

Public Sub WriteTraceLog(nText As String, Optional nStep As Long)
    mGlobals.WriteTraceLog nText, nStep
End Sub

Public Sub StopLogging()
    mGlobals.StopLogging
End Sub

Public Sub ContinueLogging()
    mGlobals.ContinueLogging
End Sub

Public Function GetParentFormHwnd(nControlHwnd As Long) As Long
    GetParentFormHwnd = mGlobals.GetParentFormHwnd(nControlHwnd)
End Function

Public Function WindowUnderMouseHwnd() As Long
    WindowUnderMouseHwnd = mGlobals.WindowUnderMouseHwnd
End Function

Public Function Trim2(nText As String) As String
    Trim2 = mGlobals.Trim2(nText)
End Function

Public Sub CenterForm(frm As Object, Optional TrueCenter As Boolean = False)
    mGlobals.CenterForm frm, TrueCenter
End Sub

Public Function LoadTextFile(nFilePath As String, Optional nDataType As vbExDataTypeConstants = vxText)
    LoadTextFile = mGlobals.LoadTextFile(nFilePath, nDataType)
End Function

Public Function GetFileName(nFileFullPath As String) As String
    GetFileName = mGlobals.GetFileName(nFileFullPath)
End Function

' Kept for keeping binary compatibility, it will be removed in the next version
Public Function BaseName(nFullFilePath As String) As String
    
End Function

' Kept for keeping binary compatibility, it will be removed in the next version
Public Function GetFileNameFromFullPath(nFileFullPath As String) As String
    
End Function

Public Function FileInUse(ByVal nFilePath As String) As Boolean
    FileInUse = mGlobals.FileInUse(nFilePath)
End Function

Public Sub SendKeysAPI(ByVal sKeys As String, Optional ByVal Wait As Boolean)
    mGlobals.SendKeysAPI sKeys, Wait
End Sub

Public Function GetWindowClassName(nHwnd As Long) As String
    GetWindowClassName = mGlobals.GetWindowClassName(nHwnd)
End Function

Public Function WindowIsForm(nHwnd As Long) As Boolean
    WindowIsForm = mGlobals.WindowIsForm(nHwnd)
End Function

Public Function GetSystemColorDepth() As Long
    GetSystemColorDepth = mGlobals.GetSystemColorDepth
End Function

' it is for the MS' original SSTab, not the SSTabEx that already has a TabBackColor property
Public Sub SetSSTabBackColor(nSSTab As Object)
    mGlobals.SetSSTabBackColor nSSTab
End Sub


Public Property Let ButtonsStyle(nValue As vbExButtonStyleConstants)
    gButtonsStyle = nValue
    
    If gButtonsStyle = vxInstallShield2 Then
        gToolbarsButtonsStyle = vxInstallShieldToolbar
    End If
End Property

Public Property Get ButtonsStyle() As vbExButtonStyleConstants
    ButtonsStyle = gButtonsStyle
End Property


Public Property Let ToolbarsButtonsStyle(nValue As vbExButtonStyleConstants)
    gToolbarsButtonsStyle = nValue
End Property

Public Property Get ToolbarsButtonsStyle() As vbExButtonStyleConstants
    ToolbarsButtonsStyle = gToolbarsButtonsStyle
End Property


Public Property Let ToolbarsDefaultIconsSize(nValue As vbExToolbarDAIconsSizeConstants)
    gToolbarsDefaultIconsSize = nValue
End Property

Public Property Get ToolbarsDefaultIconsSize() As vbExToolbarDAIconsSizeConstants
    ToolbarsDefaultIconsSize = gToolbarsDefaultIconsSize
End Property


Public Function PictureOfWindowSection(nHwnd As Long, Optional nSectionLeft As Long = 0, Optional nSectionTop As Long = 0, Optional nSectionWidth As Long = -1, Optional nSectionHeight As Long = -1) As Object
    Dim iSection As RECT
    Dim iRect As RECT
    
    If nHwnd = 0 Then Exit Function
    
    If (nSectionWidth = -1) Or (nSectionHeight = -1) Then
        GetWindowRect nHwnd, iRect
        If nSectionWidth = 1 Then
            nSectionWidth = iRect.Right - iRect.Left
        End If
        If nSectionHeight = -1 Then
            nSectionHeight = iRect.Bottom - iRect.Top
        End If
    End If
    
    iSection.Left = nSectionLeft
    iSection.Top = nSectionTop
    iSection.Right = nSectionLeft + nSectionWidth
    iSection.Bottom = nSectionTop + nSectionHeight
    
    Set PictureOfWindowSection = mGlobals.PictureOfWindowSection(nHwnd, iSection)
End Function

Public Function PictureOfWindow(nHwnd As Long) As Object
    Set PictureOfWindow = mGlobals.PictureOfWindow(nHwnd)
End Function

Public Function IsMouseButtonPressed(nButton As vbExMouseButtonsConstants) As Boolean
    IsMouseButtonPressed = mGlobals.IsMouseButtonPressed(nButton)
End Function

Public Function ResampleBMP(nPicSrc As StdPicture, nWidth As Long, nHeight As Long) As StdPicture
    Set ResampleBMP = MPicStretch.ResampleBMP(nPicSrc, nWidth, nHeight)
End Function

Public Sub ChDirEx(nPath As String)
    mGlobals.ChDirEx nPath
End Sub

Public Function GetFileDateTimesUnsigned(nPath As String, nCreatedLow As Long, nCreatedHigh As Long, nLastAccessLow As Long, nLastAccessHigh As Long, nModifiedLow As Long, nModifiedHigh As Long) As Boolean
    GetFileDateTimesUnsigned = mGlobals.GetFileDateTimesUnsigned(nPath, nCreatedLow, nCreatedHigh, nLastAccessLow, nLastAccessHigh, nModifiedLow, nModifiedHigh)
End Function

Public Function SetFileDateTimesUnsigned(nPath As String, nCreatedLow As Long, nCreatedHigh As Long, nLastAccessLow As Long, nLastAccessHigh As Long, nModifiedLow As Long, nModifiedHigh As Long) As Boolean
    SetFileDateTimesUnsigned = mGlobals.SetFileDateTimesUnsigned(nPath, nCreatedLow, nCreatedHigh, nLastAccessLow, nLastAccessHigh, nModifiedLow, nModifiedHigh)
End Function

Public Sub DateTimeToUnsigned(ByVal nDateTime As Date, nLow As Long, nHigh As Long)
    mGlobals.DateTimeToUnsigned nDateTime, nLow, nHigh
End Sub

' Kept for keeping binary compatibility, it will be removed in the next version
Public Sub CVrStr(nString As String)
Attribute CVrStr.VB_MemberFlags = "40"
    
End Sub

' Kept for keeping binary compatibility, it will be removed in the next version
Public Sub CVrBA(nBA() As Byte)
Attribute CVrBA.VB_MemberFlags = "40"
    
End Sub

Public Sub SaveTextFile(nPath As String, nText As String)
    mGlobals.SaveTextFile nPath, nText
End Sub

Public Function GetAllPrevInstancesMainWindowsHwnds() As Long()
    Dim iAP As New cActivatePrevInstance
    
    GetAllPrevInstancesMainWindowsHwnds = iAP.GetAllPrevInstancesMainWindowsHwnds
End Function

Public Function GetAllPrevInstancesWindowsHwnds() As Long()
    Dim iAP As New cActivatePrevInstance
    
    GetAllPrevInstancesWindowsHwnds = iAP.GetAllPrevInstancesWindowsHwnds
End Function

Public Property Get UniClip() As UniClipboard
    Set UniClip = New UniClipboard
End Property

Public Function GetTopZOrderForm(nForms As Object, Optional nFromOwnerForm, Optional nVisible As Boolean = True) As Object
    Set GetTopZOrderForm = mGlobals.GetTopZOrderForm(nForms, nFromOwnerForm, nVisible)
End Function
    
Public Property Get ExFnGUILanguage() As vbExGUILanguageConstants
    ExFnGUILanguage = mGUIStrings.GUILanguage
End Property

Public Property Let ExFnGUILanguage(nLang As vbExGUILanguageConstants)
    mGUIStrings.GUILanguage = nLang
End Property

Public Function AddToList(nList, nValue, Optional nOnlyIfMissing As Boolean, Optional nFirstElement As Long = 0) As Boolean
    AddToList = mGlobals.AddToList(nList, nValue, nOnlyIfMissing, nFirstElement)
End Function

Public Function IsInList(nList, nValue, Optional nFirstElement As Long = 0, Optional nLastElement As Long = -1) As Boolean
    IsInList = mGlobals.IsInList(nList, nValue, nFirstElement, nLastElement)
End Function

Public Function IndexInList(nList, nValue) As Long
    IndexInList = mGlobals.IndexInList(nList, nValue)
End Function

Public Sub OrderVector(ByRef nMainVector As Variant, Optional ByRef nSecondaryVector1 As Variant, Optional ByRef nSecondaryVector2 As Variant, Optional ByRef nSecondaryVector3 As Variant, Optional ByRef nSecondaryVector4 As Variant, Optional ByRef nSecondaryVector5 As Variant, Optional ByRef nSecondaryVector6 As Variant, Optional nBynaryCompare As Boolean, Optional nOrderDescending As Boolean)
    mOrderVector.OrderVector nMainVector, nSecondaryVector1, nSecondaryVector2, nSecondaryVector3, nSecondaryVector4, nSecondaryVector5, nSecondaryVector6, nBynaryCompare, nOrderDescending
End Sub

Public Sub AssignAccelerators(nObject As Object, Optional nAsignToLabels As Boolean)
    mGlobals.AssignAccelerators nObject, nAsignToLabels
End Sub

Public Function AssignAcceleratorToControl(nControl As Object, Optional nUsedLetters) As String
    AssignAcceleratorToControl = mGlobals.AssignAcceleratorToControl(nControl, nUsedLetters)
End Function

Public Function ValidFileName(ByVal ProposedFileName As String, Optional ReplacementChar As String = "", Optional DefaultFileName As String = "[Untitled]", Optional ForOldFileFormat_8Dot3 As Boolean = False, Optional AllowExtension As Boolean = True, Optional ByRef HasExtension As Boolean) As String
    ValidFileName = mGlobals.ValidFileName(ProposedFileName, ReplacementChar, DefaultFileName, ForOldFileFormat_8Dot3, AllowExtension, HasExtension)
End Function

Public Function ChrCount(nText As String, ByVal nCharW As Long) As Long
    ChrCount = mGlobals.ChrCount(nText, nCharW)
End Function

Public Sub SaveBinaryFile(nFilePath As String, nBytes() As Byte, Optional nDateTime)
    mGlobals.SaveBinaryFile nFilePath, nBytes, nDateTime
End Sub

Public Function AdjustPictureWithHLS(nSourcePic As StdPicture, Optional HAddition As Long, Optional LAddition As Long, Optional SAddition As Long, Optional LFactor As Single = 1, Optional SFactor As Single = 1, Optional ColorToPreserve As Long = -1) As StdPicture
    Set AdjustPictureWithHLS = MPicStretch.AdjustPictureWithHLS(nSourcePic, HAddition, LAddition, SAddition, LFactor, SFactor, ColorToPreserve)
End Function

Public Function AdjustColorWithHLS(nColor As Long, Optional HAddition As Long, Optional LAddition As Long, Optional SAddition As Long, Optional LFactor As Single = 1, Optional SFactor As Single = 1) As Long
    AdjustColorWithHLS = MPicStretch.AdjustColorWithHLS(nColor, HAddition, LAddition, SAddition, LFactor, SFactor)
End Function

Public Function SetColorToSameHue(nColor As Long, nReferenceColor As Long) As Long
    SetColorToSameHue = MPicStretch.SetColorToSameHue(nColor, nReferenceColor)
End Function
