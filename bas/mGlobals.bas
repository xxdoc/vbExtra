Attribute VB_Name = "mGlobals"
Option Explicit

Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function rtcCallByName Lib "msvbvm60" (ByRef vRet As Variant, ByVal cObj As Object, ByVal sMethod As Long, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal LCID As Long) As Long
Private Declare Function rtcCallByNameIDE Lib "vba6" Alias "rtcCallByName" (ByRef vRet As Variant, ByVal cObj As Object, ByVal sMethod As Long, ByVal eCallType As VbCallType, ByRef pArgs() As Variant, ByVal LCID As Long) As Long

Public Const LF_FACESIZE = 32

Public Type LOGFONTW
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
    lfFaceName(0 To ((LF_FACESIZE * 2) - 1)) As Byte
End Type

Public Declare Function CreateFontIndirectW Lib "gdi32" (ByRef lpLogFont As LOGFONTW) As Long

Private Declare Function GetLocaleInfo Lib "Kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Const LOCALE_USER_DEFAULT = &H400
Private Const LOCALE_SDECIMAL = &HE

Private Declare Function PathGetCharType Lib "shlwapi.dll" Alias "PathGetCharTypeW" (ByVal ch As Long) As Long

Private Const GCT_LFNCHAR = &H1
Private Const GCT_SHORTCHAR = &H2

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Const GW_HWNDFIRST As Long = 0
Private Const GW_HWNDNEXT As Long = 2
Private Const GW_CHILD As Long = 5

Private Enum efnEnumWindowsMode
    efnEWM_GetActiveFormHwnd = 1
    efnEWM_GetAppFormsHwnds = 2
    efnEWM_NumberOfOwnedForms = 3
    efnEWM_BroadcastUILanguageChange = 4
End Enum

Public Enum gfnMergeCellsSettings
    flexMergeNever = 0
    flexMergeFree = 1
    flexMergeRestrictRows = 2
    flexMergeRestrictColumns = 3
    flexMergeRestrictAll = 4
End Enum

Private Const OFS_MAXPATHNAME = 128
Private Const OF_READ = &H0

Private Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Declare Function LocalFileTimeToFileTime Lib "Kernel32" (lpFileTime As FileTime, lpLocalFileTime As FileTime) As Long
Private Declare Function VariantTimeToSystemTime Lib "oleaut32.dll" (ByVal vtime As Date, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function OpenFile Lib "Kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function GetFileTime Lib "Kernel32" (ByVal hFile As Long, lpCreationTime As FileTime, lpLastAccessTime As FileTime, lpLastWriteTime As FileTime) As Long
Private Declare Function SetFileTime Lib "Kernel32" (ByVal hFile As Long, lpCreationTime As FileTime, lpLastAccessTime As FileTime, lpLastWriteTime As FileTime) As Long
Private Declare Function SystemTimeToFileTime Lib "Kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FileTime) As Long
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function SetCurrentDirectory Lib "Kernel32" Alias "SetCurrentDirectoryA" (ByVal PathName As String) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
Private Const SM_SWAPBUTTON = 23&

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PicBmp
    Size As Long
    Type As PictureTypeConstants
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle&, iPic As IPicture) As Long

Private Const Planes& = 14
Private Const BITSPIXEL& = 12

Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function ReleaseDC& Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long)

Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

Private Const FILE_FLAG_WRITE_THROUGH As Long = &H80000000
Private Const ERROR_SHARING_VIOLATION As Long = 32

Private Const clOneMask = 16515072          '000000 111111 111111 111111
Private Const clTwoMask = 258048            '111111 000000 111111 111111
Private Const clThreeMask = 4032            '111111 111111 000000 111111
Private Const clFourMask = 63               '111111 111111 111111 000000

Private Const clHighMask = 16711680         '11111111 00000000 00000000
Private Const clMidMask = 65280             '00000000 11111111 00000000
Private Const clLowMask = 255               '00000000 00000000 11111111

Private Const cl2Exp18 = 262144             '2 to the 18th power
Private Const cl2Exp12 = 4096               '2 to the 12th
Private Const cl2Exp6 = 64                  '2 to the 6th
Private Const cl2Exp8 = 256                 '2 to the 8th
Private Const cl2Exp16 = 65536              '2 to the 16th

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private Declare Function GetTimeZoneInformation Lib "Kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const OPEN_ALWAYS As Long = 4 'Create file if it does NOT exist
Private Const FILE_BEGIN As Long = 0
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
'Private Const CREATE_NEW = 1
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Declare Function CreateFile Lib "Kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFilePointer Lib "Kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "Kernel32" (ByVal hFile As Long) As Long

'* These are used in every form to ensure focus rectangle visibility
Private Const WM_CHANGEUISTATE As Long = &H127
Private Const UIS_CLEAR As Integer = &H2
Private Const UISF_HIDEFOCUS As Integer = &H1

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Const CB_GETMINVISIBLE As Long = &H1702&

Private Const NOERROR = 0
Private Const gintMAX_PATH_LEN = 260                    ' Maximum allowed path length including path, filename,

Public Enum efnSpecialFolderIDs
    CSIDL_DESKTOP = &H0
    CSIDL_PROGRAMS = &H2
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_STARTMENU = &HB
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_PROFILE = &H28
    CSIDL_FLAG_CREATE = &H8000&
End Enum

'Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As efnSpecialFolderIDs, ByRef pidl As Long) As Long
Private Declare Function SHGetFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, pidl As Long) As Long
Private Declare Function SHGetPathFromIDListA Lib "shell32" (ByVal pidl As Long, ByVal pszPath As String) As Long
'Private Declare Function SHGetMalloc Lib "shell32" (ByRef pMalloc As IVBMalloc) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pvoid As Long)

Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Function GetMessageExtraInfo Lib "user32" () As Long

Public Const MOUSEEVENTF_LEFTDOWN = &H2 ' Left button down
Public Const MOUSEEVENTF_LEFTUP = &H4 ' Left button up

Public Const WS_EX_TOOLWINDOW = &H80
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Public Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long

Private Const WS_CAPTION = &HC00000
Private Const SM_CMONITORS As Long = 80
Private Const SM_XVIRTUALSCREEN = 76
Private Const SM_YVIRTUALSCREEN = 77
Private Const SM_CXVIRTUALSCREEN = 78
Private Const SM_CYVIRTUALSCREEN = 79

'Private Const MONITORINFOF_PRIMARY = &H1
'Public Const MONITOR_DEFAULTTONEAREST = &H2
Public Const MONITOR_DEFAULTTONULL = &H0
Public Const MONITOR_DEFAULTTOPRIMARY = &H1

Public Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
Public Declare Function MonitorFromPoint Lib "user32.dll" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As Long) As Long

Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function RealChildWindowFromPoint Lib "user32" (ByVal hWndParent As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Private Const GW_OWNER = &H4
'Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long

Private Const WM_SETREDRAW As Long = &HB&

Public Const CB_GETDROPPEDSTATE As Long = &H157

Public Type COMBOBOXINFO
   cbSize As Long
   rcItem As RECT
   rcButton As RECT
   stateButton As Long
   hWndCombo As Long
   hWndEdit As Long
   hWndList As Long
End Type

Public Declare Function GetComboBoxInfo Lib "user32" (ByVal hWndCombo As Long, CBInfo As COMBOBOXINFO) As Long

Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Const GWL_EXSTYLE = (-20)

'Public Type LOGFONT
'    lfHeight As Long
'    lfWidth As Long
'    lfEscapement As Long
'    lfOrientation As Long
'    lfWeight As Long
'    lfItalic As Byte
'    lfUnderline As Byte
'    lfStrikeOut As Byte
'    lfCharSet As Byte
'    lfOutPrecision As Byte
'    lfClipPrecision As Byte
'    lfQuality As Byte
'    lfPitchAndFamily As Byte
'    lfFaceName(0 To LF_FACESIZE - 1) As Byte
'End Type

Private Type NONCLIENTMETRICSW
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONTW
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONTW
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONTW
    lfStatusFont As LOGFONTW
    lfMessageFont As LOGFONTW
End Type

Public Const CLEARTYPE_QUALITY As Byte = 6
Public Const OUT_TT_ONLY_PRECIS As Long = 7

Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700

Private Declare Sub CopyMemoryAny1 Lib "Kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Const SPI_GETNONCLIENTMETRICS = 41
Private Const SPI_GETICONTITLELOGFONT = 31

Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long

Public Const CB_SETDROPPEDWIDTH = &H160
Public Const CB_GETDROPPEDWIDTH = &H15F

Public Const CBS_DROPDOWN = &H2&
Public Const CBS_DROPDOWNLIST = &H3&

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Public Const HORZRES As Long = 8
Public Const VERTRES As Long = 10
Public Const LOGPIXELSX As Long = 88
Public Const LOGPIXELSY As Long = 90
Public Const PHYSICALWIDTH As Long = 110
Public Const PHYSICALHEIGHT As Long = 111
Public Const PHYSICALOFFSETX As Long = 112
Public Const PHYSICALOFFSETY As Long = 113

Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type DRAWTEXTPARAMS
  cbSize As Long
  iTabLength As Long
  iLeftMargin As Long
  iRightMargin As Long
  uiLengthDrawn As Long
End Type

Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExW" (ByVal hDC As Long, ByVal lpsz As Long, ByVal N As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long

Public Const DT_EDITCONTROL As Long = &H2000&
Public Const DT_EXPANDTABS As Long = &H40
Public Const DT_LEFT As Long = &H0
Public Const DT_RIGHT As Long = &H2
Public Const DT_CENTER As Long = &H1
Public Const DT_NOPREFIX As Long = &H800
Public Const DT_TABSTOP As Long = &H80
Public Const DT_WORDBREAK As Long = &H10
Public Const DT_CALCRECT As Long = &H400

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Public Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long

Private Const MAX_PATH = 260

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST + TH32CS_SNAPPROCESS + TH32CS_SNAPTHREAD + TH32CS_SNAPMODULE)
'Private Const TH32CS_INHERIT = &H80000000

Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
'Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
'Private Const SYNCHRONIZE = &H100000

Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "Kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "Kernel32" (ByVal hSnapshot As Long, lppe As Any) As Long
Private Declare Function Process32Next Lib "Kernel32" (ByVal hSnapshot As Long, lppe As Any) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcessId Lib "Kernel32" () As Long

Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long

Private Declare Function GlobalAlloc Lib "Kernel32" (ByVal Flags As Long, ByVal lengh As Long) As Long
Private Declare Function GlobalLock Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Sub RtlMoveMemory Lib "Kernel32" (ByVal pDest As Long, ByVal pSource As Long, ByVal lengh As Long)

Private Const CF_UNICODETEXT = &HD&
Private Const GMEM_MOVEABLE = &O2&
Private Const GMEM_ZEROINIT = &O40&

Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal Format As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long

Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type

Private Declare Function DllGetVersion Lib "comctl32" (pdvi As DLLVERSIONINFO) As Long

Private Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFOEX
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformID As Long
        szCSDVersion As String * 128        ' Maintenance string for PSS usage
        wSPMajor As Integer                 ' Service Pack Major Version
        wSPMinor As Integer                 ' Service Pack Minor Version
        wSuiteMask As Integer               ' Suite Identifier
        bProductType As Byte                ' Server / Workstation / Domain Controller ?
        bReserved As Byte                   ' Reserved
End Type

Private Declare Function GetOSVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long

'Private Const VER_NT_WORKSTATION As Long = &H1
Private Const VER_NT_DOMAIN_CONTROLLER As Long = &H2
Private Const VER_NT_SERVER As Long = &H3

Private Declare Function GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function IsWow64Process Lib "Kernel32" (ByVal hProc As Long, ByRef bWow64Process As Boolean) As Long
Private Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcess Lib "Kernel32" () As Long

Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
Private Declare Function IsThemeActive Lib "uxtheme" () As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function GetThemeColor Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, pColor As Long) As Long

Private Const EP_EDITBORDER_NOSCROLL As Long = 6
Private Const EPSN_NORMAL As Long = 1
Private Const TMT_BORDERCOLOR As Long = &HED9

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_ASYNCWINDOWPOS As Long = &H4000
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOACTIVATE = &H10&
Public Const SWP_NOMOVE As Long = &H2
Public Const HWND_TOP As Long = 0
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
' Redraw window:
Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_ERASE = &H4
Public Const RDW_INTERNALPAINT = &H2
Public Const RDW_INVALIDATE = &H1
Public Const RDW_UPDATENOW = &H100

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const ES_NUMBER As Long = &H2000&
Public Const GWL_STYLE = (-16)

Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Const SM_CXVSCROLL As Long = 2
Public Const SM_CYVSCROLL As Long = 20
Public Const SM_CXEDGE As Long = 45
Public Const SM_CYEDGE As Long = 46
Public Const SM_CXBORDER  As Long = 5
Public Const SM_CYBORDER  As Long = 6
Public Const SM_CXMINTRACK = 34
Public Const SM_CYMINTRACK = 35

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Const WM_RBUTTONDOWN As Long = &H204&
Public Const WM_LBUTTONDOWN As Long = &H201&
Public Const WM_DESTROY As Long = &H2&
Public Const WM_NCACTIVATE As Long = &H86&
Public Const WM_SIZE As Long = &H5&
Public Const WM_MOVE As Long = &H3&
Public Const WM_ERASEBKGND  As Long = &H14&
Public Const WM_MOUSEMOVE As Long = &H200&
Public Const WM_GETMINMAXINFO As Long = &H24&
Public Const WM_WINDOWPOSCHANGING As Long = &H46&
Public Const WM_WINDOWPOSCHANGED As Long = &H47&
Public Const WM_PAINT As Long = &HF&
Public Const WM_MOVING As Long = &H216&
Public Const WM_PARENTNOTIFY As Long = &H210&

Private Declare Function GetTempPath Lib "Kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "Kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function SystemParametersInfoW Lib "user32" (ByVal uAction As Long, ByVal uParam As Long, ByRef pvParam As Any, ByVal fWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Const GWL_HWNDPARENT As Long = (-8)
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000

Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal wndenmprc As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)
Private Const BFFM_VALIDATEFAILED = 3

Private Const WM_UILANGCHANGED As Long = WM_USER + 12

'Private Declare Function WriteFile Lib "Kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
'Private Declare Function FlushFileBuffers Lib "Kernel32" (ByVal hFile As Long) As Long

Private Const gstrSEP_DIR$ = "\"                         ' Directory separator character
'Private Const gstrAT$ = "@"
Private Const gstrSEP_DRIVE$ = ":"                       ' Driver separater character, e.g., C:\
Private Const gstrSEP_DIRALT$ = "/"                      ' Alternate directory separator character
Private Const gstrSEP_EXT$ = "."                         ' Filename extension separator character
Private Const gstrSEP_URLDIR$ = "/"                      ' Separator for dividing directories in URL addresses.

Public Const flexSelectionFree = 0
Public Const flexHighlightNever = 0
Public Const flexSelectionByRow = 1
Public Const flexGridFlat = 1
Public Const flexSelectionByColumn = 2
Public Const flexAlignLeftTop = 0
Public Const flexAlignLeftCenter = 1
Public Const flexAlignLeftBottom = 2
Public Const flexAlignCenterTop = 3
Public Const flexAlignCenterCenter = 4
Public Const flexAlignCenterBottom = 5
Public Const flexAlignRightTop = 6
Public Const flexAlignRightCenter = 7
Public Const flexAlignRightBottom = 8
Public Const flexAlignGeneral = 9

Private mGetActiveFormHwnd As Long
Private mNotOwned As Boolean
Private mStartAtHwnd As Long

Public gWindowTitle As String
Public gCommonDialogEx_ShowFolder_StartFolder As String

Private mEnumMode As efnEnumWindowsMode
Private mUILangPrev As Long
Private mFormsHwnds() As Long
Private mOnlyVisibleWindows As Boolean
Private mNumberOfOwnedForms As Long
Private mHwndOwner_ForNumberOfOwnedForms As Long

Private mFormShownNACollection As New cObjectHandlersCollection
Private mFormMinMaxCollection As New cObjectHandlersCollection
Private mFormPersistCollection As New cObjectHandlersCollection
Private mToolTipExCollection As New cToolTipExCollection
Private mSSTabColorsHandlersCollection As New cObjectHandlersCollection
Public gButtonsStyle As vbExButtonStyleConstants
Public gToolbarsButtonsStyle As vbExButtonStyleConstants
Public gToolbarsDefaultIconsSize As vbExToolbarDAIconsSizeConstants
Public gToolBarDAButtonCopied As ToolBarDAButton

Public Const cPrintPreviewMinScale = 1
Public Const cPrintPreviewMaxScale = 1000
Public Const cPrintPreviewDefaultMinScale = 20
Public Const cPrintPreviewDefaultMaxScale = 500

Private mLogFilePath As String
Private mLogging As Boolean
Private mModalFormsHwnd() As Long
Private mCommonButtonsAccelerators As String
Private mFormsTracker As New cFormsTracker

Public Function GetActiveFormHwnd(Optional nNotOwned As Boolean, Optional nStartAtHwnd As Long) As Long
    Dim iDo As Boolean
    
    On Error GoTo TheExit:
    
    mStartAtHwnd = nStartAtHwnd
    If mStartAtHwnd = 0 Then
        GetActiveFormHwnd = GetActiveWindow
        If Not WindowIsForm(GetActiveFormHwnd) Then
            iDo = True
        End If
        If Not iDo Then
            If nNotOwned Then
                If FormIsOwned(GetActiveFormHwnd) Then
                    If GetProp(GetActiveFormHwnd, "ShownModal") <> 1 Then
                        iDo = True
                    End If
                End If
            End If
        End If
    Else
        iDo = True
    End If
    If iDo Then
        mNotOwned = nNotOwned
        mEnumMode = efnEWM_GetActiveFormHwnd
        mGetActiveFormHwnd = 0
        EnumWindows AddressOf EnumCallback, 0
        GetActiveFormHwnd = mGetActiveFormHwnd
    End If
    
TheExit:
End Function

Public Function GetFormUnderMouseHwnd() As Long
    GetFormUnderMouseHwnd = WindowUnderMouseHwnd
    If Not WindowIsForm(GetFormUnderMouseHwnd) Then
        GetFormUnderMouseHwnd = GetParentFormHwnd(GetFormUnderMouseHwnd)
        If Not WindowIsForm(GetFormUnderMouseHwnd) Then
            GetFormUnderMouseHwnd = 0
        End If
    End If
End Function

Public Sub GetAppFormsHwnds(nFormsHwnds() As Long, nTopFormHwnd As Long, Optional nOnlyVisibleWindows As Boolean)
    Dim c As Long
    
    On Error GoTo TheExit:
    mEnumMode = efnEWM_GetAppFormsHwnds
    mOnlyVisibleWindows = nOnlyVisibleWindows
    mGetActiveFormHwnd = 0
    ReDim mFormsHwnds(0)
    EnumWindows AddressOf EnumCallback, 0
    nTopFormHwnd = mGetActiveFormHwnd
    ReDim nFormsHwnds(UBound(mFormsHwnds))
    For c = 1 To UBound(nFormsHwnds)
        nFormsHwnds(c) = mFormsHwnds(c)
    Next c

TheExit:
End Sub

Public Function WindowIsForm(nHwnd As Long) As Boolean
    Dim iClassName As String
    
    If nHwnd = 0 Then Exit Function
    
    iClassName = GetWindowClassName(nHwnd)
    WindowIsForm = (iClassName = "ThunderRT6FormDC") Or (iClassName = "ThunderFormDC") Or (iClassName = "ThunderForm")
    
End Function

Public Function GetWindowClassName(nHwnd As Long) As String
    Dim iClassName As String
    Dim iSize As Long
    
    If nHwnd = 0 Then Exit Function
    
    iClassName = Space(64)
    iSize = GetClassName(nHwnd, iClassName, Len(iClassName))
    GetWindowClassName = Left$(iClassName, iSize)
    
End Function

Public Function GetOwnerHwnd(nHwnd As Long) As Long
    GetOwnerHwnd = GetWindowLong(nHwnd, GWL_HWNDPARENT)
End Function

Public Function FormIsOwned(nHwndForm As Long) As Boolean
    Dim iHwndOwner As Long
    Dim iOwnerClassName As String
    
    iHwndOwner = GetOwnerHwnd(nHwndForm)
    iOwnerClassName = GetWindowClassName(iHwndOwner)
    FormIsOwned = (iOwnerClassName = "ThunderFormDC") Or (iOwnerClassName = "ThunderRT6FormDC")
    
End Function
    
Public Function NumberOfOwnedForms(nHwndOwner As Long) As Long
     mEnumMode = efnEWM_NumberOfOwnedForms
     mNumberOfOwnedForms = 0
     mHwndOwner_ForNumberOfOwnedForms = nHwndOwner
     EnumWindows AddressOf EnumCallback, 0
     NumberOfOwnedForms = mNumberOfOwnedForms
End Function
    
    
Private Function EnumCallback(ByVal nEnumHwnd As Long, ByVal param As Long) As Long
    Dim iFound As Boolean
    Dim iIgnore As Boolean
    
    Select Case mEnumMode
        Case efnEWM_GetActiveFormHwnd
            On Error GoTo TheExit:
            If GetParent(nEnumHwnd) = 0 Then
                If IsWindowLocal(nEnumHwnd) Then
                    If IsWindowVisible(nEnumHwnd) <> 0 Then
                        If WindowIsForm(nEnumHwnd) Then
                            If mStartAtHwnd <> 0 Then
                                iIgnore = True
                            End If
                            If mNotOwned Then
                                If Not FormIsOwned(nEnumHwnd) Then
                                    iFound = True
                                Else
                                    If GetProp(nEnumHwnd, "ShownModal") = 1 Then
                                        iFound = True
                                    End If
                                End If
                            Else
                                iFound = True
                            End If
                            If iFound Then
                                If mStartAtHwnd <> 0 Then
                                    If nEnumHwnd = mStartAtHwnd Then
                                        mStartAtHwnd = 0
                                    End If
                                End If
                                If Not iIgnore Then
                                    mGetActiveFormHwnd = nEnumHwnd
                                    EnumCallback = 0
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            EnumCallback = 1
        Case efnEWM_GetAppFormsHwnds
            On Error GoTo TheExit:
            iIgnore = False
            If mOnlyVisibleWindows Then
                If IsWindowVisible(nEnumHwnd) = 0 Then
                    iIgnore = True
                End If
            End If
            If Not iIgnore Then
                If GetParent(nEnumHwnd) = 0 Then
                    If IsWindowLocal(nEnumHwnd) Then
                        If WindowIsForm(nEnumHwnd) Then
                            If mGetActiveFormHwnd = 0 Then
                                If IsWindowVisible(nEnumHwnd) <> 0 Then
                                    mGetActiveFormHwnd = nEnumHwnd
                                End If
                            End If
                            ReDim Preserve mFormsHwnds(UBound(mFormsHwnds) + 1)
                            mFormsHwnds(UBound(mFormsHwnds)) = nEnumHwnd
                        End If
                    End If
                End If
            End If
            EnumCallback = 1
        Case efnEWM_NumberOfOwnedForms
            On Error GoTo TheExit:
            If GetParent(nEnumHwnd) = 0 Then
                If IsWindowLocal(nEnumHwnd) Then
                    If WindowIsForm(nEnumHwnd) Then
                        If IsWindowVisible(nEnumHwnd) <> 0 Then
                            If FormIsOwned(nEnumHwnd) Then
                                'Debug.Print nEnumHwnd, GetOwnerHwnd(nEnumHwnd), mHwndOwner_ForNumberOfOwnedForms
                                If GetOwnerHwnd(nEnumHwnd) = mHwndOwner_ForNumberOfOwnedForms Then
                                    mNumberOfOwnedForms = mNumberOfOwnedForms + 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            EnumCallback = 1
        Case efnEWM_BroadcastUILanguageChange
            If IsWindowLocal(nEnumHwnd) Then
                If WindowIsForm(nEnumHwnd) Then
                    BroadcastUILanguageChangeToChildControls nEnumHwnd
                End If
            End If
            EnumCallback = 1
    End Select
    
Exit Function
TheExit:
    EnumCallback = 0
End Function

Private Sub BroadcastUILanguageChangeToChildControls(nHwnd As Long)
    Dim iHwnd As Long
    
    iHwnd = GetWindow(nHwnd, GW_CHILD)
    Do Until iHwnd = 0
        If GetProp(iHwnd, "FnExUI") <> 0 Then
            PostMessage iHwnd, WM_UILANGCHANGED, mUILangPrev, 0
        End If
        BroadcastUILanguageChangeToChildControls iHwnd
        iHwnd = GetWindow(iHwnd, GW_HWNDNEXT)
    Loop
    
End Sub

Public Function IsWindowLocal(ByVal hWnd As Long) As Boolean
    Dim idWnd As Long
    Call GetWindowThreadProcessId(hWnd, idWnd)
    IsWindowLocal = (idWnd = GetCurrentProcessId())
End Function


Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
   On Error Resume Next
   Dim ret As Long
   Dim sBuffer As String
   Dim iEh As Long
   
   Select Case uMsg
       Case BFFM_INITIALIZED
           Call SendMessageString(hWnd, BFFM_SETSELECTION, 1, gCommonDialogEx_ShowFolder_StartFolder)
           SetWindowText hWnd, gWindowTitle
            iEh = FindWindowEx(hWnd, 0, "Edit", "")
            SetWindowText iEh, gCommonDialogEx_ShowFolder_StartFolder
       Case BFFM_SELCHANGED
           sBuffer = Space(MAX_PATH)
           ret = SHGetPathFromIDList(lp, sBuffer)
           If ret = 1 Then
               Call SendMessageString(hWnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
               iEh = FindWindowEx(hWnd, 0, "Edit", "")
               SetWindowText iEh, sBuffer
           End If
        Case BFFM_VALIDATEFAILED
            Call SendMessageString(hWnd, BFFM_SETSELECTION, 1, "")
   End Select
   BrowseCallbackProc = 0
End Function


Public Function CloneFont(nOrigFont As Object) As Object
    Dim iFont As New StdFont
    
    If nOrigFont Is Nothing Then Exit Function
    If Not TypeOf nOrigFont Is StdFont Then Exit Function
    
    iFont.Name = nOrigFont.Name
    iFont.Size = nOrigFont.Size
    iFont.Bold = nOrigFont.Bold
    iFont.Italic = nOrigFont.Italic
    iFont.Strikethrough = nOrigFont.Strikethrough
    iFont.Underline = nOrigFont.Underline
    iFont.Weight = nOrigFont.Weight
    iFont.Charset = nOrigFont.Charset
    
    Set CloneFont = iFont
End Function

Public Function ScreenActiveForm(nForms As Object, Optional nNotOwned As Boolean) As Object
    Dim iForm As Object
    Dim iAFHwnd As Long
    
    iAFHwnd = mGlobals.GetActiveFormHwnd(nNotOwned)
    
    Do Until Not ScreenActiveForm Is Nothing
        For Each iForm In nForms
            If iForm.hWnd = iAFHwnd Then
                Set ScreenActiveForm = iForm
                Exit For
            End If
        Next
        iAFHwnd = mGlobals.GetActiveFormHwnd(nNotOwned, iAFHwnd)
        If iAFHwnd = 0 Then Exit Do
    Loop
    
End Function

Public Function GetParentFormHwnd(nControlHwnd As Long) As Long
    Dim lPar As Long
    Dim iHwnd As Long
    
    iHwnd = nControlHwnd
    lPar = GetParent(iHwnd)
    While lPar <> 0
        
        If WindowIsForm(lPar) Then
            iHwnd = lPar
        End If
        lPar = GetParent(lPar)
    Wend
    GetParentFormHwnd = iHwnd
'    Debug.Print GetParentFormHwnd
End Function

Public Function MinimizeApp() As AppMinimizer
    Set MinimizeApp = New AppMinimizer
End Function

Public Function ScreenUsableHeight()
    Dim iRect As RECT
    Static sValue As Long
    
    If sValue = 0 Then
        Call SystemParametersInfo(SPI_GETWORKAREA, 0, iRect, 0)
        sValue = (iRect.Bottom - iRect.Top) * Screen.TwipsPerPixelY
    End If
    ScreenUsableHeight = sValue
End Function

Public Function IsShowingVerticalScrollBar(nControl As Object) As Boolean
    On Error Resume Next
    IsShowingVerticalScrollBar = (GetWindowLong(nControl.hWnd, GWL_STYLE) And WS_VSCROLL) = WS_VSCROLL
End Function

Public Function IsShowingHorizontalScrollBar(nControl As Object) As Boolean
    On Error Resume Next
    IsShowingHorizontalScrollBar = (GetWindowLong(nControl.hWnd, GWL_STYLE) And WS_HSCROLL) = WS_HSCROLL
End Function

Public Function FileExists(ByVal strPathName As String) As Boolean
    Dim intFileNum As Integer

    On Error Resume Next

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    'Open strPathName For Input As intFileNum
    Open strPathName For Input As intFileNum
'    Debug.Print Err.Number, Err.Description
    FileExists = (Err.Number = 0) Or (Err.Number = 70) Or (Err.Number = 55)
    
    Close intFileNum

    Err.Clear
'    If Not FileExists Then
'        If Dir(strPathName) <> "" Then
'            FileExists = True
'        End If
'    End If
End Function

Public Function GetTempFolder() As String
    Dim lChar As Long
    
    GetTempFolder = String$(255, 0)
    lChar = GetTempPath(255, GetTempFolder)
    GetTempFolder = Left$(GetTempFolder, lChar)
    AddDirSep GetTempFolder
End Function

Public Sub AddDirSep(ByRef strPathName As String)
    strPathName = RTrim$(strPathName)
    If Right$(strPathName, Len(gstrSEP_URLDIR)) <> gstrSEP_URLDIR Then
        If Right$(strPathName, Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
            strPathName = strPathName & gstrSEP_DIR
        End If
    End If
End Sub

Public Function GetTempFileFullPath() As String
    Dim iTemp As String
    
    iTemp = String(260, 0)
    'Get a temporary filename
    GetTempFileName GetTempFolder, "", 0, iTemp
    'Remove all the unnecessary chr$(0)'s
    iTemp = Left$(iTemp, InStr(1, iTemp, Chr$(0)) - 1)
    GetTempFileFullPath = iTemp
End Function

'Given a fully qualified filename, returns the path portion and the file
'   portion.
Public Sub SeparatePathAndFileName(FullPath As String, _
    Optional ByRef Path As String, _
    Optional ByRef FileName As String)

    Dim nSepPos As Long
    Dim nSepPos2 As Long
    Dim fUsingDriveSep As Boolean

    nSepPos = InStrRev(FullPath, gstrSEP_DIR)
    nSepPos2 = InStrRev(FullPath, gstrSEP_DIRALT)
    If nSepPos2 > nSepPos Then
        nSepPos = nSepPos2
    End If
    nSepPos2 = InStrRev(FullPath, gstrSEP_DRIVE)
    If nSepPos2 > nSepPos Then
        nSepPos = nSepPos2
        fUsingDriveSep = True
    End If

    If nSepPos = 0 Then
        'Separator was not found.
        Path = CurDir$
        FileName = FullPath
    Else
        If fUsingDriveSep Then
            Path = Left$(FullPath, nSepPos)
        Else
            Path = Left$(FullPath, nSepPos - 1)
        End If
        FileName = Mid$(FullPath, nSepPos + 1)
    End If
End Sub

Public Function IsFullPath(ByVal nFileName As String) As Boolean
    Dim iFolderPath As String
    Dim iOnlyFile As String
    Dim nSepPos As Long
    Dim nSepPos2 As Long
    Dim fUsingDriveSep As Boolean

    nSepPos = InStrRev(nFileName, gstrSEP_DIR)
    nSepPos2 = InStrRev(nFileName, gstrSEP_DIRALT)
    If nSepPos2 > nSepPos Then
        nSepPos = nSepPos2
    End If
    nSepPos2 = InStrRev(nFileName, gstrSEP_DRIVE)
    
    If (nSepPos > 0) And (nSepPos2 > 0) Then
        SeparatePathAndFileName nFileName, iFolderPath, iOnlyFile
        IsFullPath = (iFolderPath <> "") And (iOnlyFile <> "")
    End If
End Function

Public Function StripExtension(ByVal nFileName As String) As String
    Dim nSepPos As Long
    Dim nSepPos2 As Long

    nSepPos = InStrRev(nFileName, gstrSEP_DIR)
    nSepPos2 = InStrRev(nFileName, gstrSEP_DIRALT)
    If nSepPos2 > nSepPos Then
        nSepPos = nSepPos2
    End If
    nSepPos2 = InStrRev(nFileName, gstrSEP_EXT)
    If (nSepPos2 > nSepPos) Or (nSepPos = 0) Then
        If nSepPos2 > 1 Then
            StripExtension = Left$(nFileName, nSepPos2 - 1)
        End If
    End If
    If StripExtension = "" Then StripExtension = nFileName
End Function

Public Function FileNameHasExtension(ByVal nFileName As String) As Boolean
    Dim nSepPos As Long
    Dim nSepPos2 As Long

    nSepPos = InStrRev(nFileName, gstrSEP_DIR)
    nSepPos2 = InStrRev(nFileName, gstrSEP_DIRALT)
    If nSepPos2 > nSepPos Then
        nSepPos = nSepPos2
    End If
    nSepPos2 = InStrRev(nFileName, gstrSEP_EXT)
    If (nSepPos2 > nSepPos) Or (nSepPos = 0) Then
        If nSepPos2 > 1 Then
            FileNameHasExtension = True
        End If
    End If
End Function

Public Function ControlNameWithParent(nControl As Object) As String
    Dim iContainer As Object
    Dim iAuxStr As String
    Dim iAuxStr2 As String
    
On Error GoTo TheExit:
    If nControl Is Nothing Then
        ControlNameWithParent = ""
    Else
        iAuxStr = nControl.Name
        On Error Resume Next
        iAuxStr = iAuxStr & "_" & nControl.Index
        On Error GoTo TheExit:
        Set iContainer = nControl.Container
        Do Until iContainer Is nControl.Parent
            iAuxStr2 = iContainer.Name
            On Error Resume Next
            iAuxStr2 = iAuxStr2 & "_" & iContainer.Index
            On Error GoTo TheExit:
            iAuxStr = iAuxStr2 & "." & iAuxStr
            Set iContainer = iContainer.Container
        Loop
        iAuxStr = nControl.Parent.Name & "." & iAuxStr
        ControlNameWithParent = iAuxStr
    End If

    Exit Function

TheExit:
    On Error GoTo -1
    On Error Resume Next
    ControlNameWithParent = nControl.Parent.Name
    ControlNameWithParent = ControlNameWithParent & nControl.Name
    ControlNameWithParent = ControlNameWithParent & "_" & nControl.Index
End Function


Public Sub SetTextBoxNumeric(nTxt As Control)
    SetWindowLong nTxt.hWnd, GWL_STYLE, GetWindowLong(nTxt.hWnd, GWL_STYLE) Or ES_NUMBER
End Sub


Public Sub ToClipboard(nText As String)
    Dim a As String
    Static sYa As Boolean
    
    If Not sYa Then
        sYa = True
    Else
        a = Clipboard.GetText
    End If
    
    Clipboard.Clear
    Clipboard.SetText a & vbCrLf & nText
End Sub


Public Function IsThemed() As Boolean
    Dim osVer As OSVERSIONINFO
    Static sValue As Long
    
    If sValue = 0 Then
        sValue = 1
        'Set size of structure.
        osVer.dwOSVersionInfoSize = Len(osVer)
        
        'Fill structure with data.
        GetVersionEx osVer
        
        'Evaluate return. If greater than or equal to 5.1 then running
        'WindowsXP or newer.
        If osVer.dwMajorVersion + osVer.dwMinorVersion / 10 >= 5.1 Then
            'Check for Active Visual Style(modified as per paravoid's suggestion).
            If IsWindowsThemed Then
                'Double Check by assessing DLL version loaded
                If (CommonControlsVersionLoaded >= 6) Then
                    sValue = 2
                End If
            End If
        End If
    End If
    IsThemed = sValue = 2
End Function

Public Function IsWindowsThemed() As Boolean
    Static sValue As Long
    
    If sValue = 0 Then
        If (IsAppThemed <> 0) And (IsThemeActive <> 0) Then
            sValue = 2
        Else
            sValue = 1
        End If
    End If
    IsWindowsThemed = (sValue = 2)
End Function

Public Function CommonControlsVersionLoaded() As Long
    Dim dllVer As DLLVERSIONINFO
    Static sValue As Long
    
    If sValue = 0 Then
        dllVer.cbSize = Len(dllVer)
        DllGetVersion dllVer
        sValue = dllVer.dwMajor
        If sValue = 0 Then sValue = -1
    End If
    CommonControlsVersionLoaded = sValue
End Function

Public Function GetTextBoxBorderColorThemed() As Long
    Dim iTheme As Long
    Dim iClass As String
    Dim iColor As Long
    Static sValue As Long
    
    If sValue = 0 Then
        iClass = "Edit"
        iTheme = OpenThemeData(0&, StrPtr(iClass))
        
        If iTheme = 0 Then
            sValue = &H8000000A
        Else
            Call GetThemeColor(iTheme, EP_EDITBORDER_NOSCROLL, EPSN_NORMAL, TMT_BORDERCOLOR, iColor)
            sValue = iColor
            CloseThemeData iTheme
        End If
        
    End If
    GetTextBoxBorderColorThemed = sValue
End Function

Public Property Get ClientExeFile() As String
    Static sAlreadySet As Boolean
    Static sValue As String
    
    If Not sAlreadySet Then
        sValue = GetClientExe
        sAlreadySet = True
    End If
    ClientExeFile = sValue
End Property

Public Property Get ClientProductName() As String
    Static sAlreadySet As Boolean
    Static sValue As String
    
    If Not sAlreadySet Then
        sValue = GetProductNameFromExeFile(ClientExeFile)
        sAlreadySet = True
    End If
    ClientProductName = sValue
End Property


Public Property Get AppNameForRegistry()
    Static sAlreadySet As Boolean
    Static sValue As String
    
    If Not sAlreadySet Then
        sValue = App.Title & "\" & Base64Encode(ClientProductName)
        sAlreadySet = True
    End If
    AppNameForRegistry = sValue
End Property

Public Sub ClipboardCopyUnicode(nText As String)
    Dim hMem As Long
    Dim pMem As Long
    
    If nText <> "" Then
        hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, LenB(nText) + 2)
        pMem = GlobalLock(hMem)
        RtlMoveMemory pMem, StrPtr(nText), LenB(nText) + 2
        If GlobalUnlock(hMem) = 0 Then
            If OpenClipboard(0) <> 0 Then
                EmptyClipboard
                SetClipboardData CF_UNICODETEXT, hMem
                CloseClipboard
                DoEvents
            End If
        End If
    End If
End Sub
    

Private Function GetClientExe() As String
    Dim cbNeeded2 As Long
    Dim Modules(1 To 200) As Long
    Dim ModuleName As String
    Dim nSize As Long
    Dim hProcess As Long
    Dim hSnapshot As Long, LRet As Long, p As PROCESSENTRY32
    Dim iProcessID As Long
    
    iProcessID = GetCurrentProcessId
    p.dwSize = Len(p)
    
    If IsWindowsNT Then
        ' NT
        'Get a handle to the Process
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, iProcessID)
        'Got a Process handle
        If hProcess <> 0 Then
            'Get an array of the module handles for the specified
            'process
            LRet = EnumProcessModules(hProcess, Modules(1), 200, _
                                         cbNeeded2)
            'If the Module Array is retrieved, Get the ModuleFileName
            If LRet <> 0 Then
               ModuleName = Space(MAX_PATH)
               nSize = 500
               LRet = GetModuleFileNameExA(hProcess, Modules(1), _
                               ModuleName, nSize)
               
               GetClientExe = Left$(ModuleName, LRet)
            End If
            CloseHandle hProcess
        End If
    Else
        'Windows 95/98
        hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, ByVal 0)
        If hSnapshot Then
            LRet = Process32First(hSnapshot, p)
            Do While LRet
                If p.th32ProcessID = iProcessID Then
                    GetClientExe = Left$(p.szExeFile, InStr(p.szExeFile, Chr$(0)) - 1)
                    Exit Do
                End If
                LRet = Process32Next(hSnapshot, p)
            Loop
            LRet = CloseHandle(hSnapshot)
        End If
    End If
End Function

Public Function GetProductNameFromExeFile(strFileName As String) As String
    Dim sInfo As String, lSizeof As Long
    Dim lResult As Long, intDel As Integer
    Dim lHandle As Long
    Dim intStrip As Integer
    Dim iIsNT As Boolean
    Dim iGetProductNameFromExeFile As String
    Dim iAuxStr As String
    
    If strFileName <> "" Then
        lHandle = 0
        lSizeof = GetFileVersionInfoSize(strFileName, lHandle)
        If lSizeof > 0 Then
            sInfo = String$(lSizeof, 0)
            lResult = GetFileVersionInfo(ByVal strFileName, 0&, ByVal lSizeof, ByVal sInfo)
            If lResult Then
                iIsNT = IsWindowsNT
                If iIsNT Then
                    sInfo = StrConv(sInfo, vbFromUnicode)
                End If
                intDel = InStr(sInfo, "ProductName")
                If intDel > 0 Then
                    If iIsNT Then
                        intDel = intDel + 13
                    Else
                        intDel = intDel + 12
                    End If
                    intStrip = InStr(intDel, sInfo, vbNullChar)
                    iGetProductNameFromExeFile = Trim$(Mid$(sInfo, intDel, intStrip - intDel))
                End If
                If Len(iGetProductNameFromExeFile) > 30 Or iGetProductNameFromExeFile = "" Then
                    intDel = InStr(sInfo, "Description")
                    If intDel > 0 Then
                        If iIsNT Then
                            intDel = intDel + 13
                        Else
                            intDel = intDel + 12
                        End If
                        intStrip = InStr(intDel, sInfo, vbNullChar)
                        iAuxStr = Trim$(Mid$(sInfo, intDel, intStrip - intDel))
                        If iAuxStr <> "" Then
                            iGetProductNameFromExeFile = iAuxStr
                        End If
                    End If
                End If
                If Len(iGetProductNameFromExeFile) > 40 Then
                    iGetProductNameFromExeFile = Left$(iGetProductNameFromExeFile, 40) & "..."
                End If
            End If
        End If
    End If
    iGetProductNameFromExeFile = Trim$(iGetProductNameFromExeFile)
    If iGetProductNameFromExeFile = "" Then
        iGetProductNameFromExeFile = GetFileName(strFileName)
    End If
    GetProductNameFromExeFile = iGetProductNameFromExeFile
End Function

Public Function GetFolder(nFileFullPath As String) As String
    Dim iFolderPath As String
    
    SeparatePathAndFileName nFileFullPath, iFolderPath
    GetFolder = iFolderPath
    AddDirSep GetFolder
End Function

Public Function ClientAppIsCompiled() As Boolean
    Static sValue As Long
    
    If sValue = 0 Then
        If ClientProductName = "Visual Basic" Then
            sValue = 1
        Else
            sValue = 2
        End If
    End If
    ClientAppIsCompiled = (sValue = 2)
End Function

Public Sub ShowNoActivate(nForm As Object, Optional nOwnerForm, Optional nSetIcon As Boolean = True, Optional nSetActiveFormAsOwner As Boolean)
    Dim iFormShownNA As cFormShownNA
    
    HandleMonitor nForm
    AddFormToTracker nForm
    
    Set iFormShownNA = mFormShownNACollection.GetInstance(nForm)
    If iFormShownNA Is Nothing Then
        Set iFormShownNA = New cFormShownNA
        iFormShownNA.ShowForm nForm, nOwnerForm, mFormShownNACollection, nSetIcon, nSetActiveFormAsOwner
        mFormShownNACollection.Add iFormShownNA, nForm.hWnd
    Else
        iFormShownNA.ShowForm nForm, nOwnerForm
    End If
End Sub

Public Function IsWindowsNT() As Boolean
    Static sValue As Long
    
    If sValue = 0 Then
        Dim osinfo As OSVERSIONINFO
        Dim retvalue As Integer
        
        osinfo.dwOSVersionInfoSize = 148
        osinfo.szCSDVersion = Space$(128)
        retvalue = GetVersionEx(osinfo)
        If osinfo.dwPlatformID = 2 Then
            sValue = 2
        Else
            sValue = 1
        End If
    End If
    
    IsWindowsNT = (sValue = 2)
End Function

Public Function IsWindows98OrMore() As Boolean
    Static sValue As Long
    
    If sValue = 0 Then
        Dim osinfo As OSVERSIONINFO
        Dim retvalue As Integer
        
        osinfo.dwOSVersionInfoSize = 148
        osinfo.szCSDVersion = Space$(128)
        retvalue = GetVersionEx(osinfo)
        sValue = 1
        If osinfo.dwPlatformID = 2 Then ' if it is NT
            If osinfo.dwMajorVersion > 4 Then ' más que NT4 (o sea win 2000, 2003, XP o Vista, etc)
                sValue = 2
            End If
        Else ' si no es NT
            If (osinfo.dwMajorVersion >= 4) And (osinfo.dwMinorVersion >= 10) Then  ' Si es 98 o ME, o bien
                sValue = 2
            End If
        End If
    End If
    
    IsWindows98OrMore = (sValue = 2)
End Function

Public Function IsWindows2000OrMore() As Boolean
    Static sValue As Long

    If sValue = 0 Then
        Dim osinfo As OSVERSIONINFO
        Dim retvalue As Integer
        
        osinfo.dwOSVersionInfoSize = 148
        osinfo.szCSDVersion = Space$(128)
        retvalue = GetVersionEx(osinfo)
        If osinfo.dwMajorVersion >= 5 Then
            sValue = 2
        Else
            sValue = 1
        End If
    End If
    
    IsWindows2000OrMore = (sValue = 2)
End Function

Public Function IsWindowsXPOrMore() As Boolean
    Static sValue As Long

    If sValue = 0 Then
        Dim osinfo As OSVERSIONINFO
        Dim retvalue As Integer
        
        osinfo.dwOSVersionInfoSize = 148
        osinfo.szCSDVersion = Space$(128)
        retvalue = GetVersionEx(osinfo)
        sValue = 1
        If osinfo.dwPlatformID = 2 Then ' if it is NT
            If (osinfo.dwMajorVersion = 5) And (osinfo.dwMinorVersion >= 1) Or (osinfo.dwMajorVersion > 5) Then
                sValue = 2
            End If
        End If
    End If
    
    IsWindowsXPOrMore = (sValue = 2)
End Function

Public Function IsWindowsXP() As Boolean
    Static sValue As Long

    If sValue = 0 Then
        Dim osinfo As OSVERSIONINFO
        Dim retvalue As Integer
        
        osinfo.dwOSVersionInfoSize = 148
        osinfo.szCSDVersion = Space$(128)
        retvalue = GetVersionEx(osinfo)
        sValue = 1
        If osinfo.dwPlatformID = 2 Then ' if it is NT
            If (osinfo.dwMajorVersion = 5) And (osinfo.dwMinorVersion = 1) Then
                sValue = 2
            End If
        End If
    End If
    
    IsWindowsXP = (sValue = 2)
End Function

Public Function IsWindowsVistaOrMore() As Boolean
    Static sValue As Long

    If sValue = 0 Then
        Dim osinfo As OSVERSIONINFO
        Dim retvalue As Integer
        
        osinfo.dwOSVersionInfoSize = 148
        osinfo.szCSDVersion = Space$(128)
        retvalue = GetVersionEx(osinfo)
        sValue = 1
        If osinfo.dwPlatformID = 2 Then ' if it is NT
            If osinfo.dwMajorVersion >= 6 Then ' Vista is 6
                sValue = 2
            End If
        End If
    End If
    
    IsWindowsVistaOrMore = (sValue = 2)
End Function

Public Function IsWindows7OrMore() As Boolean
    Static sValue As Long

    If sValue = 0 Then
        Dim osinfo As OSVERSIONINFO
        Dim retvalue As Integer
        
        osinfo.dwOSVersionInfoSize = 148
        osinfo.szCSDVersion = Space$(128)
        retvalue = GetVersionEx(osinfo)
        sValue = 1
        If osinfo.dwPlatformID = 2 Then ' if it is NT
            If osinfo.dwMajorVersion > 6 Then
                sValue = 2
            Else
                If osinfo.dwMajorVersion = 6 Then
                    If osinfo.dwMinorVersion >= 1 Then
                        sValue = 2
                    End If
                End If
            End If
        End If
    End If
    
    IsWindows7OrMore = (sValue = 2)
End Function

Public Function IsWindows8OrMore() As Boolean
    Static sValue As Long

    If sValue = 0 Then
        Dim osinfo As OSVERSIONINFO
        Dim retvalue As Integer
        
        osinfo.dwOSVersionInfoSize = 148
        osinfo.szCSDVersion = Space$(128)
        retvalue = GetVersionEx(osinfo)
        sValue = 1
        If osinfo.dwPlatformID = 2 Then ' if it is NT
            If osinfo.dwMajorVersion > 6 Then
                sValue = 2
            Else
                If osinfo.dwMajorVersion = 6 Then
                    If osinfo.dwMinorVersion >= 2 Then
                        sValue = 2
                    End If
                End If
            End If
        End If
    End If
    
    IsWindows8OrMore = (sValue = 2)
End Function

Public Function IsWindowsServer() As Boolean
    Static sValue As Long

    If sValue = 0 Then
        Dim osinfo As OSVERSIONINFOEX
        Dim retvalue As Integer
        
        osinfo.dwOSVersionInfoSize = Len(osinfo)
        osinfo.szCSDVersion = Space$(128)
        retvalue = GetOSVersionEx(osinfo)
        sValue = 1
        If osinfo.dwPlatformID = 2 Then ' if it is NT
            If (osinfo.bProductType = VER_NT_SERVER) Or (osinfo.bProductType = VER_NT_DOMAIN_CONTROLLER) Then
                sValue = 2
            End If
        End If
    End If
    IsWindowsServer = (sValue = 2)
End Function

Public Function IsWindows64Bits() As Boolean
    Dim iHandle As Long
    Dim iIs64Bits As Boolean
    Static sValue As Long
    
    If sValue = 0 Then
        ' Assume initially that this is not a WOW64 process
        iIs64Bits = False
    
        ' Then try to prove that wrong by attempting to load the
        ' IsWow64Process function dynamically
        iHandle = GetProcAddress(GetModuleHandle("Kernel32"), "IsWow64Process")
    
        ' The function exists, so call it
        If iHandle <> 0 Then
            IsWow64Process GetCurrentProcess(), iIs64Bits
        End If
    
        ' Return the value
        If iIs64Bits Then
            sValue = 2
        Else
            sValue = 1
        End If
    End If
    
    IsWindows64Bits = (sValue = 2)
End Function

Public Function WindowUnderMouseHwnd() As Long
    Dim iP As POINTAPI
    
    GetCursorPos iP
    WindowUnderMouseHwnd = WindowFromPoint(iP.x, iP.y)
    
End Function

Public Function MouseIsOverControl(nControl As Control) As Boolean
    Dim iP As POINTAPI
    Dim iR As RECT
    Dim iHwndWindowUnderMouse As Long
    
    GetCursorPos iP
    GetWindowRect nControl.hWnd, iR
    
    iHwndWindowUnderMouse = WindowFromPoint(iP.x, iP.y)
    
    If (iHwndWindowUnderMouse = nControl.hWnd) Or (MyGetProp(iHwndWindowUnderMouse, "TTPic") <> 0) Then
        If iP.y > iR.Top Then
            If iP.y < iR.Bottom Then
                If iP.x > iR.Left Then
                    
                    iR.Right = iR.Right - GetSystemMetrics(SM_CXEDGE) * 2
                    If TypeName(nControl) = "ComboBox" Then
                        If ComboHasDropDownButton(nControl) Then
                            iR.Right = iR.Right - GetSystemMetrics(SM_CXVSCROLL)
                        End If
                    End If
                    
                    If iP.x < iR.Right Then
                        MouseIsOverControl = True
                    End If
                End If
            End If
        End If
    End If
End Function

Public Function ComboHasDropDownButton(nCombo As Control) As Boolean
    Dim lStyle As Long
    
    lStyle = GetWindowLong(nCombo.hWnd, GWL_STYLE)
    If ((lStyle And CBS_DROPDOWN) = CBS_DROPDOWN) Or ((lStyle And CBS_DROPDOWNLIST) = CBS_DROPDOWNLIST) Then
        ComboHasDropDownButton = True
    End If
End Function

Public Function ControlTextWidth(nControl As Control, Optional ByVal nText As String) As Long
    Dim iP As POINTAPI
    Dim iDC As Long
    Dim iLOGFONT As LOGFONTW
    Dim iFontHandle As Long
    Dim iOldFont As Long
    Dim iDCCtl As Long
    
    If nText = "" Then
        nText = nControl.Text
    End If
    
    iDCCtl = GetDC(nControl.hWnd)
    iDC = CreateCompatibleDC(iDCCtl)
    ReleaseDC nControl.hWnd, iDCCtl
    If iDC = 0 Then Exit Function
    
    iLOGFONT = StdFontToLogFont_Screen(iDC, nControl.Font)
    iFontHandle = CreateFontIndirectW(iLOGFONT)
    iOldFont = SelectObject(iDC, iFontHandle)
    
    GetTextExtentPoint32 iDC, nText, Len(nText), iP
    ControlTextWidth = iP.x
    
    If iOldFont <> 0 Then
        SelectObject iDC, iOldFont
    End If
    DeleteDC iDC
    DeleteObject iFontHandle
End Function

Public Function ControlWidth(nControl As Control) As Long
    Dim iR As RECT
    Dim iTextWidth As Long
    
    GetWindowRect nControl.hWnd, iR
    iTextWidth = iR.Right - iR.Left - GetSystemMetrics(SM_CXEDGE) * 2
    If ComboHasDropDownButton(nControl) Then
        iTextWidth = iTextWidth - GetSystemMetrics(SM_CXVSCROLL) - GetSystemMetrics(SM_CXEDGE)
    End If
    ControlWidth = iTextWidth
End Function

Public Function StdFontToLogFont_Screen(nHdc As Long, nFont As StdFont) As LOGFONTW
    Dim iFontName As String
    Dim iDPIY As Single
    Dim c As Long
    Dim iBytes() As Byte
    
    iFontName = nFont.Name
    
    iDPIY = GetDeviceCaps(nHdc, LOGPIXELSY)
    
    iBytes = iFontName
    For c = 0 To 31
        If c < Len(iFontName) Then
            StdFontToLogFont_Screen.lfFaceName(c * 2) = iBytes(c * 2) '                 .lfFaceName(c) = Asc(Mid$(iFontName, c + 1, 1))
            StdFontToLogFont_Screen.lfFaceName(c * 2 + 1) = iBytes(c * 2 + 1)
        Else
            StdFontToLogFont_Screen.lfFaceName(c * 2) = 0
            StdFontToLogFont_Screen.lfFaceName(c * 2 + 1) = 0
        End If
    Next
     
    StdFontToLogFont_Screen.lfHeight = -Round(nFont.Size * iDPIY / 72)
     
    If nFont.Italic Then
        StdFontToLogFont_Screen.lfItalic = 1
    End If
    StdFontToLogFont_Screen.lfQuality = CLEARTYPE_QUALITY
    StdFontToLogFont_Screen.lfOutPrecision = 0 ' default
    If nFont.Strikethrough Then
        StdFontToLogFont_Screen.lfStrikeOut = 1
    End If
    If nFont.Underline Then
        StdFontToLogFont_Screen.lfUnderline = 1
    End If
    StdFontToLogFont_Screen.lfWeight = nFont.Weight
    StdFontToLogFont_Screen.lfCharSet = nFont.Charset
    
End Function

Public Sub SetWindowRedraw(nHwnd As Long, nRedraw As Boolean, Optional nForce As Boolean)
    
    If Not nRedraw Then
        If IsWindowVisible(nHwnd) = 0 Then Exit Sub
    End If
    
    Static sHwnds() As Long
    Static sCalls() As Long
    Dim c As Long
    Dim t As Long
    Dim i As Long
   
    i = 0
    On Error Resume Next
    Err.Clear
    t = UBound(sHwnds)
    If Err.Number = 9 Then
        ReDim sHwnds(0)
        ReDim sCalls(0)
        t = 0
    Else
        For c = 1 To t
            If sHwnds(c) = nHwnd Then
                i = c
                Exit For
            End If
        Next c
    End If
    On Error GoTo 0
    If (i = 0) Then
        If nRedraw Then Exit Sub
        ReDim Preserve sHwnds(t + 1)
        sHwnds(t + 1) = nHwnd
        ReDim Preserve sCalls(t + 1)
        sCalls(t + 1) = 1
        i = 1
    Else
        If nRedraw Then
            sCalls(i) = sCalls(i) - 1
            If sCalls(i) < 0 Then sCalls(i) = 0
        Else
            sCalls(i) = sCalls(i) + 1
        End If
    End If
    If nRedraw And nForce Then
        SendMessageLong nHwnd, WM_SETREDRAW, True, 0&
        RedrawWindow nHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
        sCalls(i) = 0
    Else
        Select Case sCalls(i)
            Case 1
                SendMessageLong nHwnd, WM_SETREDRAW, False, 0&
            Case 0
                SendMessageLong nHwnd, WM_SETREDRAW, True, 0&
                RedrawWindow nHwnd, ByVal 0&, 0&, RDW_INVALIDATE Or RDW_ALLCHILDREN
        End Select
    End If
End Sub

Public Function IsWindowVisibleOnScreen(nHwnd As Long, Optional AtLeastPartially As Boolean) As Boolean
    Dim iPt As POINTAPI
    Dim iRect As RECT
    Dim iHwnd As Long
    Dim ihMonitor As Long
    Dim iMi As MONITORINFO
    
    GetWindowRect nHwnd, iRect
    
    iHwnd = WindowFromPoint((iRect.Left + iRect.Right) / 2, (iRect.Top + iRect.Bottom) / 2)
    IsWindowVisibleOnScreen = nHwnd = iHwnd
    If Not IsWindowVisibleOnScreen Then
        Do Until iHwnd = 0
            iHwnd = GetParent(iHwnd)
            IsWindowVisibleOnScreen = nHwnd = iHwnd
            If IsWindowVisibleOnScreen Then Exit Do
        Loop
    End If
    
    If (Not IsWindowVisibleOnScreen) And AtLeastPartially Then
        iHwnd = WindowFromPoint(iRect.Right - 1, iRect.Bottom - 1)
        IsWindowVisibleOnScreen = nHwnd = iHwnd
        If Not IsWindowVisibleOnScreen Then
            iHwnd = GetParent(iHwnd)
            IsWindowVisibleOnScreen = nHwnd = iHwnd
        End If
        If Not IsWindowVisibleOnScreen Then
            iHwnd = WindowFromPoint(CLng((iRect.Left + iRect.Right) / 2), CLng((iRect.Top + iRect.Bottom) / 2))
            IsWindowVisibleOnScreen = nHwnd = iHwnd
            If Not IsWindowVisibleOnScreen Then
                iHwnd = GetParent(iHwnd)
                IsWindowVisibleOnScreen = nHwnd = iHwnd
            End If
        End If
'        If Not IsWindowVisibleOnScreen Then
'            ihMonitor = MonitorFromWindow(nHwnd, MONITOR_DEFAULTTONULL)
'            If ihMonitor <> 0 Then
'                iMI.cbSize = Len(iMI)
'                GetMonitorInfo ihMonitor, iMI
'                If (iRect.Right > iMI.rcWork.Left) And (iRect.Bottom > iMI.rcWork.Top) And (iRect.Left < iMI.rcWork.Right) And (iRect.Top < iMI.rcWork.Bottom) Then
'                    IsWindowVisibleOnScreen = True
'                End If
'            End If
'        End If
        If Not IsWindowVisibleOnScreen Then
            iPt.x = (iRect.Left + iRect.Right) / 2
            iPt.y = (iRect.Top + iRect.Bottom) / 2
            ScreenToClient GetParent(nHwnd), iPt
            iHwnd = RealChildWindowFromPoint(GetParent(nHwnd), iPt.x, iPt.y)
            If nHwnd = iHwnd Then
                ihMonitor = MonitorFromWindow(nHwnd, MONITOR_DEFAULTTONULL)
                If ihMonitor <> 0 Then
                    iMi.cbSize = Len(iMi)
                    GetMonitorInfo ihMonitor, iMi
                    If (iRect.Right > iMi.rcWork.Left) And (iRect.Bottom > iMi.rcWork.Top) And (iRect.Left < iMi.rcWork.Right) And (iRect.Top < iMi.rcWork.Bottom) Then
                        IsWindowVisibleOnScreen = True
                    End If
                End If
            End If
        End If
    End If
End Function

Public Sub PersistForm(nForm As Object, nForms As Object, Optional nInitialCentered As Boolean = True, Optional nInitialLeft, Optional nInitialTop, Optional nInitialWidth, Optional nInitialHeight, Optional nPersistLeft As Boolean = True, Optional nPersistTop As Boolean = True, Optional nPersistWidth As Boolean = True, Optional nPersistHeight As Boolean = True, Optional nMaxTop, Optional nPersistMinimizedState As Boolean, Optional nContext As String)
    Dim iFormPersist As cFormPersist
    
    Set iFormPersist = mFormPersistCollection.GetInstance(nForm)
    If iFormPersist Is Nothing Then
        Set iFormPersist = New cFormPersist
        iFormPersist.SetForm nForm, nForms, , nInitialCentered, nInitialLeft, nInitialTop, nInitialWidth, nInitialHeight, nPersistLeft, nPersistTop, nPersistWidth, nPersistHeight, nMaxTop, nPersistMinimizedState, nContext, mFormPersistCollection
        If mFormPersistCollection.GetInstance(nForm) Is Nothing Then
            mFormPersistCollection.Add iFormPersist, nForm.hWnd
        End If
    End If
    
End Sub

Public Sub UnpersistForm(nForm As Object)
    Dim iFormPersist As cFormPersist
    
    Set iFormPersist = mFormPersistCollection.GetInstance(nForm)
    If Not iFormPersist Is Nothing Then
        iFormPersist.Unpersist
    End If
End Sub

Public Sub SaveFormPersistence(nForm As Object)
    Dim iFormPersist As cFormPersist
    
    Set iFormPersist = mFormPersistCollection.GetInstance(nForm)
    If Not iFormPersist Is Nothing Then
        iFormPersist.SaveFormPersistence
    End If
End Sub

Public Function GetFormPersistedWindowState(nForm As Object) As Long
    Dim iFormPersist As cFormPersist
    
    Set iFormPersist = mFormPersistCollection.GetInstance(nForm)
    If Not iFormPersist Is Nothing Then
        GetFormPersistedWindowState = iFormPersist.GetFormPersistedWindowState
    End If
End Function

Public Sub ShowModal(nForm As Object, Optional nWaitWithDoevents As Boolean = True, Optional nSetIcon As Boolean = True, Optional nFormsHwndToKeepEnabled As Variant, Optional nKeepEnabledTaskBarWindows As Boolean = True, Optional nNoOwner As Boolean)
    Dim iMF As New cFormModal
    
    HandleMonitor nForm
    AddFormToTracker nForm
    iMF.Show nForm, nWaitWithDoevents, nSetIcon, nFormsHwndToKeepEnabled, nKeepEnabledTaskBarWindows, nNoOwner
End Sub

Private Sub AddFormToTracker(nForm As Object)
    If WindowHasCaption(nForm.hWnd) Then
        mFormsTracker.AddForm nForm
    Else
        mFormsTracker.Update  ' to ensure the monitor set with mouse location with the first form
    End If
End Sub

Private Sub HandleMonitor(nForm As Form)
    Dim iMonitorForm As Long
    Dim iMICurrent As MONITORINFO
    Dim iMIForm As MONITORINFO
    Dim iLng As Long
    
    If (MonitorCount > 1) And GetSetting(AppNameForRegistry, "MInfo", Base64Encode(nForm.Name) & ".MI", "0") = "0" Then
        iMonitorForm = MonitorFromWindow(nForm.hWnd, MONITOR_DEFAULTTOPRIMARY)
        If mFormsTracker.CurrentMonitor <> iMonitorForm Then
            iMICurrent.cbSize = Len(iMICurrent)
            iMIForm.cbSize = Len(iMIForm)
            GetMonitorInfo mFormsTracker.CurrentMonitor, iMICurrent
            GetMonitorInfo iMonitorForm, iMIForm
            If ((iMICurrent.rcWork.Bottom - iMICurrent.rcWork.Top) <> 0) And ((iMIForm.rcWork.Bottom - iMIForm.rcWork.Top) <> 0) Then
                nForm.Move nForm.Left + (iMICurrent.rcWork.Left - iMIForm.rcWork.Left) * Screen.TwipsPerPixelX, nForm.Top + (iMICurrent.rcWork.Top - iMIForm.rcWork.Top) * Screen.TwipsPerPixelY
                If nForm.Left < (iMICurrent.rcWork.Left * Screen.TwipsPerPixelX) Then
                    nForm.Left = iMICurrent.rcWork.Left * Screen.TwipsPerPixelX
                End If
                If nForm.Top < (iMICurrent.rcWork.Top * Screen.TwipsPerPixelY) Then
                    nForm.Top = iMICurrent.rcWork.Top * Screen.TwipsPerPixelY
                End If
                If nForm.BorderStyle = vbSizable Then
                    If nForm.Height > (iMICurrent.rcWork.Bottom - iMICurrent.rcWork.Top) * Screen.TwipsPerPixelY Then
                        nForm.Height = (iMICurrent.rcWork.Bottom - iMICurrent.rcWork.Top) * Screen.TwipsPerPixelY
                    End If
                    If nForm.Width > (iMICurrent.rcWork.Right - iMICurrent.rcWork.Left) * Screen.TwipsPerPixelX Then
                        nForm.Width = (iMICurrent.rcWork.Right - iMICurrent.rcWork.Left) * Screen.TwipsPerPixelX
                    End If
                    If (nForm.Left + nForm.Width) / Screen.TwipsPerPixelX > VirtualScreenRight Then
                        iLng = VirtualScreenRight - nForm.Width / Screen.TwipsPerPixelX
                        nForm.Left = iLng * Screen.TwipsPerPixelX
                    End If
                    If (nForm.Top + nForm.Height) / Screen.TwipsPerPixelY > VirtualScreenBottom Then
                        iLng = VirtualScreenBottom - nForm.Height / Screen.TwipsPerPixelY
                        nForm.Top = iLng * Screen.TwipsPerPixelY
                    End If
                    iLng = iMICurrent.rcWork.Right - nForm.Width / Screen.TwipsPerPixelX
                    If (nForm.Left / Screen.TwipsPerPixelX) > iLng Then
                        If MonitorFromPoint((nForm.Left + nForm.Width) / Screen.TwipsPerPixelX, (nForm.Top + nForm.Height) / Screen.TwipsPerPixelY, MONITOR_DEFAULTTONULL) = 0 Then ' if there is no monitor covering that point
                            nForm.Left = iLng * Screen.TwipsPerPixelX
                        End If
                    End If
                    iLng = iMICurrent.rcWork.Bottom - nForm.Height / Screen.TwipsPerPixelY
                    If (nForm.Top / Screen.TwipsPerPixelY) > iLng Then
                        If MonitorFromPoint((nForm.Left + nForm.Width) / Screen.TwipsPerPixelX, (nForm.Top + nForm.Height) / Screen.TwipsPerPixelY, MONITOR_DEFAULTTONULL) = 0 Then ' if there is no monitor covering that point
                            nForm.Top = iLng * Screen.TwipsPerPixelY
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Function SetMinMax(nForm As Object, Optional nMinWidth, Optional nMinHeight, Optional nMaxWidth, Optional nMaxHeight, Optional ScaleMode As Integer = vbTwips) As FormMinMax
    Dim iFormMinMax As FormMinMax
    
    Set iFormMinMax = mFormMinMaxCollection.GetInstance(nForm)
    If iFormMinMax Is Nothing Then
        Set iFormMinMax = New FormMinMax
        iFormMinMax.SetMinMax nForm, nMinWidth, nMinHeight, nMaxWidth, nMaxHeight, ScaleMode, mFormMinMaxCollection
        If mFormMinMaxCollection.GetInstance(nForm) Is Nothing Then
            mFormMinMaxCollection.Add iFormMinMax, nForm.hWnd
        End If
    Else
        iFormMinMax.SetMinMax nForm, nMinWidth, nMinHeight, nMaxWidth, nMaxHeight, ScaleMode
    End If
    Set SetMinMax = iFormMinMax
End Function

Public Sub SetSSTabBackColor(nSSTab As Object)
    Dim iSSTabColorsHandler As cSSTabColorsHandler
    Dim iHwnd As Long
    
    If LCase$(TypeName(nSSTab)) <> "sstab" Then Exit Sub
    On Error Resume Next
    iHwnd = nSSTab.hWnd
    On Error GoTo 0
    If iHwnd = 0 Then Exit Sub
    
    Set iSSTabColorsHandler = mSSTabColorsHandlersCollection.GetInstance(nSSTab)
    
    If iSSTabColorsHandler Is Nothing Then
        Set iSSTabColorsHandler = New cSSTabColorsHandler
        iSSTabColorsHandler.SetSSTab nSSTab, mSSTabColorsHandlersCollection
        If mSSTabColorsHandlersCollection.GetInstance(nSSTab) Is Nothing Then
            mSSTabColorsHandlersCollection.Add iSSTabColorsHandler, iHwnd
        End If
    End If
End Sub

Public Function GetSpecialfolder(nFolder As efnSpecialFolderIDs) As String
'    Dim oMalloc As IVBMalloc
    Dim sPath   As String
    Dim IDL     As Long
    
'    If SHGetSpecialFolderLocation(0, nFolder, IDL) = NOERROR Then
    If SHGetFolderLocation(0&, nFolder, 0&, 0&, IDL) = NOERROR Then
        sPath = String$(gintMAX_PATH_LEN, 0)
        SHGetPathFromIDListA IDL, sPath
'        SHGetMalloc oMalloc
'        oMalloc.Free IDL
        CoTaskMemFree IDL
        GetSpecialfolder = StringFromBuffer(sPath)
        AddDirSep GetSpecialfolder
    End If
End Function

Public Function StringFromBuffer(Buffer As String) As String
    Dim iPos As Long

    iPos = InStr(Buffer, vbNullChar)
    If iPos > 0 Then
        StringFromBuffer = Left$(Buffer, iPos - 1)
    Else
        StringFromBuffer = Buffer
    End If
End Function

Public Function GetSettingFont(AppName As String, Section As String, Key As String, DefaultFont As StdFont) As StdFont
    Set GetSettingFont = CloneFont(DefaultFont)
    
    GetSettingFont.Name = GetSetting(AppName, Section, Key & "_Name", GetSettingFont.Name)
    GetSettingFont.Size = GetSetting(AppName, Section, Key & "_Size", GetSettingFont.Size)
    GetSettingFont.Bold = CBool(GetSetting(AppName, Section, Key & "_Bold", GetSettingFont.Bold))
    GetSettingFont.Italic = CBool(GetSetting(AppName, Section, Key & "_Italic", GetSettingFont.Italic))
    GetSettingFont.Strikethrough = CBool(GetSetting(AppName, Section, Key & "_Strikethrough", GetSettingFont.Strikethrough))
    GetSettingFont.Underline = CBool(GetSetting(AppName, Section, Key & "_Underline", GetSettingFont.Underline))
    GetSettingFont.Weight = GetSetting(AppName, Section, Key & "_Weight", GetSettingFont.Weight)
    GetSettingFont.Charset = GetSetting(AppName, Section, Key & "_Charset", GetSettingFont.Charset)
    
End Function

Public Sub SaveSettingFont(AppName As String, Section As String, Key As String, Font As StdFont)
    SaveSetting AppName, Section, Key & "_Name", Font.Name
    SaveSetting AppName, Section, Key & "_Size", Font.Size
    SaveSetting AppName, Section, Key & "_Bold", CLng(Font.Bold)
    SaveSetting AppName, Section, Key & "_Italic", CLng(Font.Italic)
    SaveSetting AppName, Section, Key & "_Strikethrough", CLng(Font.Strikethrough)
    SaveSetting AppName, Section, Key & "_Underline", CLng(Font.Underline)
    SaveSetting AppName, Section, Key & "_Weight", Font.Weight
    SaveSetting AppName, Section, Key & "_Charset", Font.Charset
End Sub

Public Function FontsAreEqual(nFont1 As StdFont, nFont2 As StdFont) As Boolean
    If nFont1 Is Nothing Or nFont2 Is Nothing Then Exit Function
    
    If nFont1.Name = nFont2.Name Then
        If nFont1.Size = nFont2.Size Then
            If nFont1.Bold = nFont2.Bold Then
                If nFont1.Italic = nFont2.Italic Then
                    If nFont1.Strikethrough = nFont2.Strikethrough Then
                        If nFont1.Underline = nFont2.Underline Then
                            If nFont1.Weight = nFont2.Weight Then
                                If nFont1.Charset = nFont2.Charset Then
                                    FontsAreEqual = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
End Function

Public Function AutoSizeDropDownWidth(Combo As Object, Optional ObjectForHdc) As Long
    '**************************************************************
    'PURPOSE: Automatically size the combo box drop down width
    '         based on the width of the longest item in the combo box
    
    'PARAMETERS: Combo - ComboBox to size
    
    'RETURNS: True if successful, false otherwise
    
    'ASSUMPTIONS: 1. Form's Scale Mode is vbTwips, which is why
    '                conversion from twips to pixels are made.
    '                API functions require units in pixels
    '
    '             2. Combo Box's parent is a form or other
    '                container that support the hDC property
    
    'EXAMPLE: AutoSizeDropDownWidth Combo1
    '****************************************************************
    Dim LRet As Long
    Dim lCurrentWidth As Single
    Dim rectCboText As RECT
    Dim lParentHDC As Long
    Dim lListCount As Long
    Dim lCtr As Long
    Dim lTempWidth As Long
    Dim lWidth As Long
    Dim sSavedFont As String
    Dim sngSavedSize As Single
    Dim bSavedBold As Boolean
    Dim bSavedItalic As Boolean
    Dim bSavedUnderline As Boolean
    Dim bFontSaved As Boolean
    Dim iRc As RECT
    Dim iObjectForHdc As Object
    Dim iMaxItemsWithoutScrollBar As Long
    
    On Error GoTo errorHandler
    
    If Not TypeOf Combo Is ComboBox Then Exit Function
    
    If Not IsMissing(ObjectForHdc) Then
        Set iObjectForHdc = ObjectForHdc
    Else
        Set iObjectForHdc = Combo.Parent
    End If
    
    lParentHDC = iObjectForHdc.hDC
    If lParentHDC = 0 Then Exit Function
    lListCount = Combo.ListCount
    If lListCount = 0 Then Exit Function
    
    'Change font of parent to combo box's font
    'Save first so it can be reverted when finished
    'this is necessary for drawtext API Function
    'which is used to determine longest string in combo box
    With iObjectForHdc
        sSavedFont = .FontName
        sngSavedSize = .FontSize
        bSavedBold = .FontBold
        bSavedItalic = .FontItalic
        bSavedUnderline = .FontUnderLine
        
        .FontName = Combo.FontName
        .FontSize = Combo.FontSize
        .FontBold = Combo.FontBold
        .FontItalic = Combo.FontItalic
        .FontUnderLine = Combo.FontItalic
    End With
    
    bFontSaved = True
    
    'Get the width of the largest item
    For lCtr = 0 To lListCount
       DrawText lParentHDC, Combo.List(lCtr), -1, rectCboText, DT_CALCRECT
       'adjust the number added (20 in this case to
       'achieve desired right margin
       lTempWidth = rectCboText.Right - rectCboText.Left + GetSystemMetrics(SM_CXEDGE) * 2
    
       If (lTempWidth > lWidth) Then
          lWidth = lTempWidth
       End If
    Next
     
    If IsWindows7OrMore Then
        If IsThemed Then
            iMaxItemsWithoutScrollBar = SendMessageLong(Combo.hWnd, CB_GETMINVISIBLE, 0&, 0&)
        Else
            iMaxItemsWithoutScrollBar = 30
        End If
    Else
        iMaxItemsWithoutScrollBar = 8
    End If
    
    If Combo.ListCount > iMaxItemsWithoutScrollBar Then
         lTempWidth = lTempWidth + GetSystemMetrics(SM_CXVSCROLL)
    End If
     
     
    GetWindowRect Combo.hWnd, iRc
    LRet = SendMessageLong(Combo.hWnd, CB_SETDROPPEDWIDTH, iRc.Right - iRc.Left, 0)
    
    lCurrentWidth = SendMessageLong(Combo.hWnd, CB_GETDROPPEDWIDTH, 0, 0)
    
    If lCurrentWidth > lWidth Then 'current drop-down width is
    '                               sufficient
'        AutoSizeDropDownWidth = True
        AutoSizeDropDownWidth = lCurrentWidth
        GoTo errorHandler
        Exit Function
    End If
     
    'don't allow drop-down width to
    'exceed screen.width
    If lWidth > Screen.Width \ Screen.TwipsPerPixelX - 20 Then lWidth = Screen.Width \ Screen.TwipsPerPixelX - 20
    
    LRet = SendMessageLong(Combo.hWnd, CB_SETDROPPEDWIDTH, lWidth, 0)
    AutoSizeDropDownWidth = lWidth
'    AutoSizeDropDownWidth = LRet > 0

errorHandler:
    On Error Resume Next
    If bFontSaved Then
    'restore parent's font settings
      With iObjectForHdc
        .FontName = sSavedFont
        .FontSize = sngSavedSize
        .FontUnderLine = bSavedUnderline
        .FontBold = bSavedBold
        .FontItalic = bSavedItalic
     End With
    End If
End Function

Public Function ThemeColor(ByVal nThemeID As String) As Long
    nThemeID = UCase$(nThemeID)
    Select Case nThemeID
        Case "TEXTBOXBORDER"
            ThemeColor = GetTextBoxBorderColorThemed
    End Select
End Function

Public Function ColorsBlended(ByVal nColor1 As Long, ByVal nColor2 As Long, ByVal nPercentColor2 As Long)
    If nPercentColor2 > 100 Then nPercentColor2 = 100
    If nPercentColor2 < 0 Then nPercentColor2 = 0
    If nPercentColor2 = 0 Then
        ColorsBlended = nColor1
        Exit Function
    End If
    If nPercentColor2 = 1000 Then
        ColorsBlended = nColor2
        Exit Function
    End If
    
    Dim iR1 As Long
    Dim iG1 As Long
    Dim iB1 As Long
    Dim iR2 As Long
    Dim iG2 As Long
    Dim iB2 As Long
    Dim iR As Long
    Dim iG As Long
    Dim iB As Long

    iR1 = nColor1 And 255
    iG1 = (nColor1 \ 256) And 255
    iB1 = (nColor1 \ 65536) And 255
    iR2 = nColor2 And 255
    iG2 = (nColor2 \ 256) And 255
    iB2 = (nColor2 \ 65536) And 255
    
    iR = (iR1 * (100 - nPercentColor2) + iR2 * nPercentColor2) / 100
    iG = (iG1 * (100 - nPercentColor2) + iG2 * nPercentColor2) / 100
    iB = (iB1 * (100 - nPercentColor2) + iB2 * nPercentColor2) / 100
    If iR > 255 Then iR = 255
    If iG > 255 Then iG = 255
    If iB > 255 Then iB = 255
    
    ColorsBlended = RGB(iR, iG, iB)
End Function

Public Function GetSystemFont(nSystemFont As vbExSystemFontConstants) As StdFont
    Dim iLF As LOGFONTW
    Dim iNcm As NONCLIENTMETRICSW
    Dim iILf As LOGFONTW
    Dim iRet As Long
    
'    iNcm.cbSize = 340
'    iNcm.cbSize = 500
    iNcm.cbSize = LenB(iNcm)
    iRet = SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, iNcm.cbSize, iNcm, 0)
    If (iRet = 0) Then Exit Function
    
    Select Case nSystemFont
        Case vxCaptionFont
            CopyMemoryAny1 iLF, iNcm.lfCaptionFont, LenB(iNcm.lfCaptionFont)
        Case vxIconFont
            iRet = SystemParametersInfoW(SPI_GETICONTITLELOGFONT, LenB(iILf), iILf, 0)
            If (iRet <> 0) Then
                CopyMemoryAny1 iLF, iILf, LenB(iILf)
            End If
        Case vxMenuFont
            CopyMemoryAny1 iLF, iNcm.lfMenuFont, LenB(iNcm.lfMenuFont)
        Case vxMsgBoxFont
            CopyMemoryAny1 iLF, iNcm.lfMessageFont, LenB(iNcm.lfMessageFont)
        Case vxSmallCaptionFont
            CopyMemoryAny1 iLF, iNcm.lfSMCaptionFont, LenB(iNcm.lfSMCaptionFont)
        Case vxStatusAndTooltipFont
            CopyMemoryAny1 iLF, iNcm.lfStatusFont, LenB(iNcm.lfStatusFont)
        Case Else
            Exit Function
    End Select
    
    Set GetSystemFont = LogFontToStdFont(iLF)
End Function

Private Function LogFontToStdFont(lF As LOGFONTW, Optional nPrinterFont As Boolean) As iFont
    Dim iFontName As String
    Dim iDPIY As Single
    Dim iDC As Long
    
    Set LogFontToStdFont = New StdFont
    
    If lF.lfHeight = 0 Then Exit Function
    
    If nPrinterFont Then
        iDPIY = GetDeviceCaps(Printer.hDC, LOGPIXELSY)
    Else
        iDC = GetDC(0)
        iDPIY = GetDeviceCaps(iDC, LOGPIXELSY)
        ReleaseDC 0, iDC
    End If
    
    iFontName = lF.lfFaceName
    If Len(iFontName) > 0 Then
        If InStr(iFontName, Chr(0)) > 0 Then
            iFontName = Left$(iFontName, InStr(iFontName, Chr(0)) - 1)
        End If
    End If
    
    If iFontName <> "" Then
        LogFontToStdFont.Name = iFontName
    Else
        LogFontToStdFont.Name = "Arial"
    End If
    
    Select Case lF.lfHeight
        Case Is < 0
            LogFontToStdFont.Size = -lF.lfHeight / iDPIY * 72
        Case Is > 0
            LogFontToStdFont.Size = lF.lfHeight / iDPIY * 72 * 0.8 ' lF.lfHeight / Screen.TwipsPerPixelY / 2.777777777
        Case Else
            LogFontToStdFont.Size = 12
    End Select
    
    If lF.lfWeight > 1000 Then
        LogFontToStdFont.Weight = 400
        If LogFontToStdFont.Size > 20 Then LogFontToStdFont.Size = 12
    Else
        LogFontToStdFont.Weight = lF.lfWeight
    End If
    
    LogFontToStdFont.Italic = lF.lfItalic
    LogFontToStdFont.Strikethrough = lF.lfStrikeOut
    LogFontToStdFont.Underline = lF.lfUnderline
    LogFontToStdFont.Charset = lF.lfCharSet
End Function


Public Function GetComboListHwnd(nCombo As Object) As Long
    Dim iCboInf As COMBOBOXINFO
    
    iCboInf.cbSize = Len(iCboInf)
    GetComboBoxInfo nCombo.hWnd, iCboInf
    GetComboListHwnd = iCboInf.hWndList

End Function

Public Function GetComboEditHwnd(nCombo As Object) As Long
    Dim iCboInf As COMBOBOXINFO
    
    iCboInf.cbSize = Len(iCboInf)
    GetComboBoxInfo nCombo.hWnd, iCboInf
    GetComboEditHwnd = iCboInf.hWndEdit

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

Public Function FolderExists(ByVal nFolderPath As String) As Boolean
    On Error Resume Next

    FolderExists = (GetAttr(nFolderPath) And vbDirectory) = vbDirectory

    Err.Clear
End Function

Public Sub EnsureFocusRect(nForm As Object)
    Dim iHwnd As Long
    
    On Error Resume Next
    iHwnd = nForm.hWnd
    On Error GoTo 0
    If iHwnd <> 0 Then
        SendMessageLong iHwnd, WM_CHANGEUISTATE, MakeLong(UIS_CLEAR, UISF_HIDEFOCUS), ByVal 0&
    End If
End Sub

Public Function ShowToolTipEx(nTipText As String, Optional nTitle As String, Optional nStyle As vbExBalloonTooltipStyleConstants = vxTTBalloon, Optional nCloseButton As Boolean, Optional nIcon As vbExBalloonTooltipIconConstants = vxTTNoIcon, Optional nDelayTimeSeconds, Optional nVisibleTimeSeconds, Optional nPositionX, Optional nPositionY, Optional nPositionIsRelative As Boolean, Optional nWidth, Optional nBackColor, Optional nForeColor, Optional nClosePrevious As Boolean = True, Optional nRestrictMouseMoveToTwips As Long = 1000) As ToolTipEx
    Dim iPt As POINTAPI
    Dim iDelayTimeSeconds As Variant
    Dim iVisibleTimeSeconds As Variant
    Dim iPositionX As Variant
    Dim iPositionY As Variant
    Dim iWidth As Variant
    Dim iBackColor As Variant
    Dim iForeColor As Variant
    Dim iCBT As ToolTipEx
    Dim iParentHwnd As Long
    
    iParentHwnd = GetFormUnderMouseHwnd
    If (iParentHwnd = 0) Then
        iParentHwnd = GetActiveFormHwnd
    ElseIf Not IsWindowLocal(iParentHwnd) Then
        iParentHwnd = GetActiveFormHwnd
    End If
    If iParentHwnd = 0 Then Exit Function
    
    If nPositionIsRelative Then
        GetCursorPos iPt
        ScreenToClient iParentHwnd, iPt
        
        If Not IsMissing(nPositionX) Then
            iPositionX = (iPt.x * Screen.TwipsPerPixelX) + nPositionX
        End If
        If Not IsMissing(nPositionY) Then
            iPositionY = (iPt.y * Screen.TwipsPerPixelY) + nPositionY
        End If
    Else
        If Not IsMissing(nPositionX) Then
            iPositionX = nPositionX
        End If
        If Not IsMissing(nPositionY) Then
            iPositionY = nPositionY
        End If
    End If
    If Not IsMissing(nDelayTimeSeconds) Then
        iDelayTimeSeconds = nDelayTimeSeconds
    End If
    If Not IsMissing(nVisibleTimeSeconds) Then
        iVisibleTimeSeconds = nVisibleTimeSeconds
    End If
    If Not IsMissing(nWidth) Then
        iWidth = nWidth
    End If
    If Not IsMissing(nBackColor) Then
        iBackColor = nBackColor
    End If
    If Not IsMissing(nForeColor) Then
        iForeColor = nForeColor
    End If
    
    For Each iCBT In mToolTipExCollection.GetCollection
        If iCBT.ParentHwnd = iParentHwnd Then
            If iCBT.TipText = nTipText Then
                If iCBT.Title = nTitle Then
                    If iCBT.BackColor = iBackColor Then
                        If iCBT.ForeColor = iForeColor Then
                            If iCBT.CloseButton = nCloseButton Then
                                If iCBT.DelayTimeSeconds = iDelayTimeSeconds Then
                                    If iCBT.VisibleTimeSeconds = iVisibleTimeSeconds Then
                                        If iCBT.Icon = nIcon Then
                                            If iCBT.PositionX = iPositionX Then
                                                If iCBT.PositionY = iPositionY Then
                                                    If iCBT.Style = nStyle Then
                                                        If iCBT.Width = iWidth Then
                                                            If iCBT.RestrictMouseMoveToTwips = nRestrictMouseMoveToTwips Then
                                                                iCBT.Reset
                                                                Set ShowToolTipEx = iCBT
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If ShowToolTipEx Is Nothing Then
        If nClosePrevious Then
            For Each iCBT In mToolTipExCollection.GetCollection
                If iCBT.ParentHwnd = iParentHwnd Then
                    iCBT.CloseTip
                End If
            Next
        End If
        Set iCBT = New ToolTipEx
        iCBT.TipText = nTipText
        iCBT.Title = nTitle
        iCBT.BackColor = iBackColor
        iCBT.ForeColor = iForeColor
        iCBT.CloseButton = nCloseButton
        iCBT.DelayTimeSeconds = iDelayTimeSeconds
        iCBT.VisibleTimeSeconds = iVisibleTimeSeconds
        iCBT.Icon = nIcon
        iCBT.PositionX = iPositionX
        iCBT.PositionY = iPositionY
        iCBT.Style = nStyle
        iCBT.Width = iWidth
        iCBT.RestrictMouseMoveToTwips = nRestrictMouseMoveToTwips
        Set iCBT.TTCollection = mToolTipExCollection
        iCBT.Create iParentHwnd
        
        mToolTipExCollection.Add iCBT, iCBT.ToolTipHwnd
        
        Set ShowToolTipEx = iCBT
    End If
End Function

Public Function GetProgramDocumentsFolder() As String
    Dim iStr As String
    
    iStr = GetSetting(AppNameForRegistry, "Preferences", "DocsFolder", "")
    On Error Resume Next
    If iStr = "" Then
        iStr = GetSpecialfolder(CSIDL_PERSONAL Or CSIDL_FLAG_CREATE)
    Else
        If Not FolderExists(iStr) Then
            iStr = GetSpecialfolder(CSIDL_PERSONAL Or CSIDL_FLAG_CREATE)
        End If
    End If
    On Error GoTo 0
    If iStr = "" Then
        iStr = "C:\"
    Else
        If Right$(iStr, 1) <> "\" Then
            iStr = iStr & "\"
        End If
    End If
    
    GetProgramDocumentsFolder = iStr
End Function

Public Sub SaveProgramDocumentsFolder(ByVal nPath As String)
    If nPath = "" Then Exit Sub
    
    If Right$(nPath, 1) <> "\" Then
        nPath = nPath & "\"
    End If
    
    SaveSetting AppNameForRegistry, "Preferences", "DocsFolder", nPath
End Sub

Public Function GetFileName(nFileFullPath As String) As String
    Dim iFileName As String
    
    SeparatePathAndFileName nFileFullPath, , iFileName
    GetFileName = iFileName
End Function

Private Sub DoubleTo2Longs(ByVal dbl As Double, nLongLOW As Long, nLongHigh As Long)
    'convert a double -> 2 longs
    
    'Note: a double is stored in 64 bits, of which
    '52 bits are available for the mantissa. So
    'don't use more than 5 hex digits for the
    'upper DWord, or you'll loose precision.
    
    Dim temp As Long
    Dim dblTemp As Double
    Dim IsNegative As Boolean
    
    'use a constant to avoid the
    'slow "2^31" below (2147483648)
    Const g2 = 2# * &H40000000
    
    If dbl < 0 Then
        IsNegative = True
        dblTemp# = -dbl - 1
    Else
        dblTemp# = dbl
    End If
    
    '(2 ^ 31)) = 2147483648
    temp = Int(dblTemp# / g2)
    
    '(2 ^ 31))
    nLongLOW = dblTemp# - (temp * g2)
    
    '(2 ^ 31)
    If temp And 1 Then nLongLOW = nLongLOW Or -g2
    
    nLongHigh = Int(temp / 2)
    
    If IsNegative Then
        nLongLOW = Not nLongLOW
        nLongHigh = Not nLongHigh
    End If

End Sub

Public Function CheckFreeDiskSpace(nPathToTest As String, nRequiredFreeSpaceInBytes As Double) As Boolean
    Dim iFileHandle As Long
    Dim iFilePath As String
    Dim iFileSizeBytes As Double
    Dim iFileSizeLow As Long
    Dim iFileSizeHigh As Long
    
    If nRequiredFreeSpaceInBytes < 1 Then
        CheckFreeDiskSpace = True
        Exit Function
    End If
    
    iFilePath = nPathToTest ' "\\Ix2-200-thubce\general\Informática\CARPETA DE TRABAJO DE RELEVAMIENTO\Relevamiento-Programa\hugefile1.dat"
    If Right$(iFilePath, 1) <> "\" Then
        iFilePath = iFilePath & "\"
        iFilePath = iFilePath & "temp_xfst_1.tmp"
        If Dir(iFilePath) <> "" Then
            On Error Resume Next
            Kill iFilePath
            On Error GoTo 0
        End If
        If Dir(iFilePath) <> "" Then
'            CheckFreeDiskSpace = True
            Exit Function
        End If
    Else
        iFilePath = iFilePath & "temp_xfst_1.tmp"
    End If
    iFileSizeBytes = nRequiredFreeSpaceInBytes
    
    iFileHandle = CreateFile(iFilePath, GENERIC_READ Or GENERIC_WRITE, 0&, ByVal 0&, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0&)
    
    If iFileHandle <> INVALID_HANDLE_VALUE Then
        Call DoubleTo2Longs(ByVal iFileSizeBytes, iFileSizeLow, iFileSizeHigh)
        If SetFilePointer(iFileHandle, iFileSizeLow, iFileSizeHigh, FILE_BEGIN) <> 0 Then
            If SetEndOfFile(iFileHandle) = 1 Then
                CheckFreeDiskSpace = True
            End If 'SetEndOfFile
        End If 'SetFilePointer
    End If 'iFileHandle
    CloseHandle iFileHandle
    
    On Error Resume Next
    Kill iFilePath
    On Error GoTo 0
End Function

'Purpose     :  Converts local time to GMT.
'Inputs      :  nLocalDateTime                 The local data time to return as GMT.
'Outputs     :  Returns the local time in GMT.
'Author      :  Andrew Baker
'Date        :  13/11/2002 10:16
'Notes       :
'Revisions   :
Public Function GMTDateTime(nLocalDateTime As Date) As Date
    Dim lSecsDiff As Long
    
    'Get the GMT time diff
    lSecsDiff = GetLocalToGMTDifference()

    'Return the time in GMT
    GMTDateTime = DateAdd("s", -lSecsDiff, nLocalDateTime)
End Function


'Purpose     :  Converts GMT time to local time.
'Inputs      :  dtLocalDate                 The GMT data time to return as local time.
'Outputs     :  Returns GMT as local time.
'Author      :  Andrew Baker
'Date        :  13/11/2002 10:16
'Notes       :
'Revisions   :

Public Function LocalDateTime(nGMTDateTime As Date) As Date
    Dim Differerence As Long
    
    Differerence = GetLocalToGMTDifference()
    LocalDateTime = DateAdd("s", Differerence, nGMTDateTime)
End Function



'Purpose     :  Returns the time lDiff between local and GMT (secs).
'Inputs      :  dtLocalDate                 The local data time to return as GMT.
'Outputs     :  Returns the local time in GMT.
'Author      :  Andrew Baker
'Date        :  13/11/2002 10:16
'Notes       :  A positive number indicates your ahead of GMT.
'Revisions   :

Public Function GetLocalToGMTDifference() As Long
'    Const TIME_ZONE_ID_INVALID& = &HFFFFFFFF
'    Const TIME_ZONE_ID_STANDARD& = 1
'    Const TIME_ZONE_ID_UNKNOWN& = 0
    Const TIME_ZONE_ID_DAYLIGHT& = 2
    
    Dim tTimeZoneInf As TIME_ZONE_INFORMATION
    Dim LRet As Long
    Dim lDiff As Long
    
    'Get time zone info
    LRet = GetTimeZoneInformation(tTimeZoneInf)
    
    'Convert diff to secs
    lDiff = -tTimeZoneInf.Bias * 60
    GetLocalToGMTDifference = lDiff
    
    'Check if we are in daylight saving time.
    If LRet = TIME_ZONE_ID_DAYLIGHT& Then
        'In daylight savings, apply the bias
        If tTimeZoneInf.DaylightDate.wMonth <> 0 Then
            'if tTimeZoneInf.DaylightDate.wMonth = 0 then the daylight
            'saving time change doesn't occur
            GetLocalToGMTDifference = lDiff - tTimeZoneInf.DaylightBias * 60
        End If
    End If
End Function

'Public Function IsInControlList(nControl As Control, nItem As String) As Boolean
'    Dim c As Long
'
'    For c = 0 To nControl.ListCount - 1
'        If nControl.List(c) = nItem Then
'            IsInControlList = True
'            Exit For
'        End If
'    Next c
'End Function

Public Function IsControlArray(nControl As Control) As Boolean
    Dim iIndex As Long
    
    iIndex = -1
    On Error Resume Next
    iIndex = nControl.Index
    IsControlArray = iIndex <> -1
End Function

Public Sub StartLogging(Optional ByVal nPath As String)
    If nPath <> "" Then
        mLogFilePath = nPath
    Else
        mLogFilePath = ClientExeFile
        If mLogFilePath <> "" Then
            mLogFilePath = Left$(mLogFilePath, Len(mLogFilePath) - 4) & ".log"
        End If
    End If
    If Dir(mLogFilePath) <> "" Then
        On Error Resume Next
        Kill mLogFilePath
        On Error GoTo 0
    End If
    mLogging = True
End Sub

Public Sub WriteTraceLog(nText As String, Optional nStep As Long)
    Dim iStr As String
    Dim hFile As Long
    Static sLevel As Long
    
    If Not mLogging Then Exit Sub
    'If Not mLogging Then StartLogging
    
    If nStep = 1 Then
        iStr = "Enter " & nText
        sLevel = sLevel + 1
    ElseIf nStep = 2 Then
        iStr = "Leave " & nText
    Else
        iStr = nText & IIf(nStep = 0, "", " step " & CStr(nStep))
    End If
    If sLevel < 1 Then sLevel = 1
    
    hFile = FreeFile
    Open mLogFilePath For Append As #hFile
    Print #hFile, Space$((sLevel - 1) * 4) & iStr
    Close #hFile
    If nStep = 2 Then
        sLevel = sLevel - 1
    End If
    
End Sub

Public Sub StopLogging()
    mLogging = False
End Sub

Public Sub ContinueLogging()
    If mLogFilePath = "" Then
        RaiseError 1001, App.ProductName, "Logging not started."
    Else
        mLogging = True
    End If
End Sub

Public Function IsInVector(nVector, nValor, Optional nBaseElement As Long = 0) As Boolean
    Dim c As Long
    
    For c = nBaseElement To UBound(nVector)
        If nVector(c) = nValor Then
            IsInVector = True
            Exit For
        End If
    Next c
End Function

Public Function Trim2(nText As String) As String
    Dim iChar As String
    
    Trim2 = nText
    iChar = Left$(Trim2, 1)
    Do While (iChar = " ") Or (iChar = vbTab) Or (iChar = vbCr) Or (iChar = vbLf) Or (iChar = Chr(160))
        Trim2 = Mid$(Trim2, 2)
        iChar = Left$(Trim2, 1)
    Loop
    iChar = Right$(Trim2, 1)
    Do While (iChar = " ") Or (iChar = vbTab) Or (iChar = vbCr) Or (iChar = vbLf) Or (iChar = Chr(160))
        Trim2 = Left$(Trim2, Len(Trim2) - 1)
        iChar = Right$(Trim2, 1)
    Loop
End Function

Public Function WithoutConsecutiveSpaces(nText As String) As String
    WithoutConsecutiveSpaces = nText
    Do Until InStr(WithoutConsecutiveSpaces, "  ") = 0
        WithoutConsecutiveSpaces = Replace(WithoutConsecutiveSpaces, "  ", " ")
    Loop
End Function

Public Sub CenterForm(nForm As Object, Optional TrueCenter As Boolean = False)
    Dim iTop As Long
    Dim iLeft As Long
    Dim iInPrimaryMonitor As Boolean
    Dim iMonitor As Long
    Dim iMi As MONITORINFO
    
    If nForm.WindowState <> vbNormal Then Exit Sub
    
    iInPrimaryMonitor = True
    
    If MonitorCount > 1 Then
        If WindowHasCaption(nForm.hWnd) Then
            iMi.cbSize = Len(iMi)
            GetMonitorInfo mFormsTracker.CurrentMonitor, iMi
            If (iMi.rcWork.Bottom - iMi.rcWork.Top) <> 0 Then
                If (iMi.rcWork.Left <> 0) Or (iMi.rcWork.Top <> 0) Then
                    iInPrimaryMonitor = False
                End If
            End If
        End If
    End If
    
    If iInPrimaryMonitor Then
        If TrueCenter Then
            iTop = Screen.Height \ 2 - nForm.Height \ 2
            iLeft = Screen.Width \ 2 - nForm.Width \ 2
        Else
            iTop = (Screen.Height * 0.9) \ 2 - nForm.Height \ 2
            iLeft = Screen.Width \ 2 - nForm.Width \ 2
            If iTop < 0 Then
                iTop = Screen.Height \ 2 - nForm.Height \ 2
            End If
        End If
    
        If iLeft < 0 Then iLeft = 0
        If iTop < 0 Then iTop = 0
    
        If nForm.BorderStyle = vbSizable Then
            If nForm.Height > ScreenUsableHeight Then
                nForm.Height = ScreenUsableHeight
            End If
            If nForm.Width > Screen.Width Then
                nForm.Width = Screen.Width
            End If
        End If
    
    Else
        If TrueCenter Then
            iTop = (iMi.rcWork.Bottom + iMi.rcWork.Top) / 2 * Screen.TwipsPerPixelY - nForm.Height \ 2
            iLeft = (iMi.rcWork.Right + iMi.rcWork.Left) / 2 * Screen.TwipsPerPixelX - nForm.Width \ 2
        Else
            iTop = (iMi.rcWork.Bottom + iMi.rcWork.Top) * 0.9 / 2 * Screen.TwipsPerPixelY - nForm.Height \ 2
            iLeft = (iMi.rcWork.Right + iMi.rcWork.Left) / 2 * Screen.TwipsPerPixelX - nForm.Width \ 2
            If iTop < 0 Then
                iTop = (iMi.rcWork.Bottom + iMi.rcWork.Top) / 2 * Screen.TwipsPerPixelY - nForm.Height \ 2
            End If
        End If
        
        If iLeft < VirtualScreenLeft Then iLeft = VirtualScreenLeft
        If iTop < VirtualScreenTop Then iTop = VirtualScreenTop
               
        If nForm.BorderStyle = vbSizable Then
            If nForm.Height > VirtualScreenHeight Then
                nForm.Height = VirtualScreenHeight
            End If
            If nForm.Width > VirtualScreenWidth Then
                nForm.Width = VirtualScreenWidth
            End If
        End If
               
    End If
    
    nForm.Top = iTop
    nForm.Left = iLeft
    
End Sub

Public Function MonitorCount() As Long
    MonitorCount = GetSystemMetrics(SM_CMONITORS)
End Function

Public Function WindowHasCaption(nHwnd As Long) As Boolean
    WindowHasCaption = (GetWindowLong(nHwnd, GWL_STYLE) And WS_CAPTION) <> 0
End Function

Public Function IsFormLoaded(nFormNameOrObject As Object) As Boolean
    Dim frm As Form
    
    For Each frm In Forms
        If frm Is nFormNameOrObject Then
'            If frm.Visible Then
                IsFormLoaded = True
                Exit For
'            End If
        End If
    Next
End Function

Public Function LoadTextFile(nFilePath As String, Optional nDataType As vbExDataTypeConstants = vxText)
    Dim iMP1 As Long
    Dim iFile As Long
    
    If FileExists(nFilePath) Then
        iMP1 = Screen.MousePointer
        Screen.MousePointer = vbHourglass
        iFile = FreeFile
        Open nFilePath For Input Access Read As #iFile
        If LOF(iFile) > 0 Then
            If nDataType = vxText Then
                LoadTextFile = Input(LOF(iFile), iFile)
            Else
                LoadTextFile = InputB(LOF(iFile), iFile)
            End If
        End If
        Close #iFile
        Screen.MousePointer = iMP1
    End If
End Function

Public Function Base64Encode(ByVal sString As String, Optional nTextHasUnicodeCharacters As Boolean) As String

    Dim bTrans(63) As Byte, lPowers8(255) As Long, lPowers16(255) As Long, bOut() As Byte, bIn() As Byte
    Dim lChar As Long, lTrip As Long, iPad As Integer, lLen As Long, lTemp As Long, lPos As Long, lOutSize As Long
    
    For lTemp = 0 To 63                                 'Fill the translation table.
        Select Case lTemp
            Case 0 To 25
                bTrans(lTemp) = 65 + lTemp              'A - Z
            Case 26 To 51
                bTrans(lTemp) = 71 + lTemp              'a - z
            Case 52 To 61
                bTrans(lTemp) = lTemp - 4               '1 - 0
            Case 62
                bTrans(lTemp) = 43                      'Chr(43) = "+"
            Case 63
                bTrans(lTemp) = 47                      'Chr(47) = "/"
        End Select
    Next lTemp

    For lTemp = 0 To 255                                'Fill the 2^8 and 2^16 lookup tables.
        lPowers8(lTemp) = lTemp * cl2Exp8
        lPowers16(lTemp) = lTemp * cl2Exp16
    Next lTemp
    
    If nTextHasUnicodeCharacters Then
        sString = StrConv(sString, vbUnicode)
    End If
    
    iPad = Len(sString) Mod 3                           'See if the length is divisible by 3
    If iPad Then                                        'If not, figure out the end pad and resize the input.
        iPad = 3 - iPad
        sString = sString & String(iPad, Chr(0))
    End If

    bIn = StrConv(sString, vbFromUnicode)               'Load the input string.
    lLen = ((UBound(bIn) + 1) \ 3) * 4                  'Length of resulting string.
    lTemp = lLen \ 72                                   'Added space for vbCrLfs.
    lOutSize = ((lTemp * 2) + lLen) - 1                 'Calculate the size of the output buffer.
    ReDim bOut(lOutSize)                                'Make the output buffer.
    
    lLen = 0                                            'Reusing this one, so reset it.
    
    For lChar = LBound(bIn) To UBound(bIn) Step 3
        lTrip = lPowers16(bIn(lChar)) + lPowers8(bIn(lChar + 1)) + bIn(lChar + 2)    'Combine the 3 bytes
        lTemp = lTrip And clOneMask                     'Mask for the first 6 bits
        bOut(lPos) = bTrans(lTemp \ cl2Exp18)           'Shift it down to the low 6 bits and get the value
        lTemp = lTrip And clTwoMask                     'Mask for the second set.
        bOut(lPos + 1) = bTrans(lTemp \ cl2Exp12)       'Shift it down and translate.
        lTemp = lTrip And clThreeMask                   'Mask for the third set.
        bOut(lPos + 2) = bTrans(lTemp \ cl2Exp6)        'Shift it down and translate.
        bOut(lPos + 3) = bTrans(lTrip And clFourMask)   'Mask for the low set.
        If lLen = 68 Then                               'Ready for a newline
            bOut(lPos + 4) = 13                         'Chr(13) = vbCr
            bOut(lPos + 5) = 10                         'Chr(10) = vbLf
            lLen = 0                                    'Reset the counter
            lPos = lPos + 6
        Else
            lLen = lLen + 4
            lPos = lPos + 4
        End If
    Next lChar
    
    If bOut(lOutSize) = 10 Then lOutSize = lOutSize - 2 'Shift the padding chars down if it ends with CrLf.
    
    If iPad = 1 Then                                    'Add the padding chars if any.
        bOut(lOutSize) = 61                             'Chr(61) = "="
    ElseIf iPad = 2 Then
        bOut(lOutSize) = 61
        bOut(lOutSize - 1) = 61
    End If
    
    Base64Encode = StrConv(bOut, vbUnicode)                 'Convert back to a string and return it.
    
End Function

Public Function Base64Decode(sString As String, Optional nTextHasUnicodeCharacters As Boolean) As String

    Dim bOut() As Byte, bIn() As Byte, bTrans(255) As Byte, lPowers6(63) As Long, lPowers12(63) As Long
    Dim lPowers18(63) As Long, lQuad As Long, iPad As Integer, lChar As Long, lPos As Long, sOut As String
    Dim lTemp As Long
    
    If Len(sString) = 0 Then Exit Function
    On Error GoTo TheExit:
    
    sString = Replace(sString, vbCr, vbNullString)      'Get rid of the vbCrLfs.  These could be in...
    sString = Replace(sString, vbLf, vbNullString)      'either order.

    lTemp = Len(sString) Mod 4                          'Test for valid input.
    If lTemp Then
        Call Err.Raise(vbObjectError, "MyDecode", "Input string is not valid Base64.")
    End If
    
    If InStrRev(sString, "==") Then                     'InStrRev is faster when you know it's at the end.
        iPad = 2                                        'Note:  These translate to 0, so you can leave them...
    ElseIf InStrRev(sString, "=") Then                  'in the string and just resize the output.
        iPad = 1
    End If
     
    For lTemp = 0 To 255                                'Fill the translation table.
        Select Case lTemp
            Case 65 To 90
                bTrans(lTemp) = lTemp - 65              'A - Z
            Case 97 To 122
                bTrans(lTemp) = lTemp - 71              'a - z
            Case 48 To 57
                bTrans(lTemp) = lTemp + 4               '1 - 0
            Case 43
                bTrans(lTemp) = 62                      'Chr(43) = "+"
            Case 47
                bTrans(lTemp) = 63                      'Chr(47) = "/"
        End Select
    Next lTemp

    For lTemp = 0 To 63                                 'Fill the 2^6, 2^12, and 2^18 lookup tables.
        lPowers6(lTemp) = lTemp * cl2Exp6
        lPowers12(lTemp) = lTemp * cl2Exp12
        lPowers18(lTemp) = lTemp * cl2Exp18
    Next lTemp

    bIn = StrConv(sString, vbFromUnicode)               'Load the input byte array.
    ReDim bOut((((UBound(bIn) + 1) \ 4) * 3) - 1)       'Prepare the output buffer.
    
    For lChar = 0 To UBound(bIn) Step 4
        lQuad = lPowers18(bTrans(bIn(lChar))) + lPowers12(bTrans(bIn(lChar + 1))) + _
                lPowers6(bTrans(bIn(lChar + 2))) + bTrans(bIn(lChar + 3))           'Rebuild the bits.
        lTemp = lQuad And clHighMask                    'Mask for the first byte
        bOut(lPos) = lTemp \ cl2Exp16                   'Shift it down
        lTemp = lQuad And clMidMask                     'Mask for the second byte
        bOut(lPos + 1) = lTemp \ cl2Exp8                'Shift it down
        bOut(lPos + 2) = lQuad And clLowMask            'Mask for the third byte
        lPos = lPos + 3
    Next lChar

    sOut = StrConv(bOut, vbUnicode)                     'Convert back to a string.
    If iPad Then sOut = Left$(sOut, Len(sOut) - iPad)   'Chop off any extra bytes.
    
    If nTextHasUnicodeCharacters Then
        Base64Decode = StrConv(sOut, vbFromUnicode)
    Else
        Base64Decode = sOut
    End If
    
    Exit Function

TheExit:
    Base64Decode = ""

End Function

Public Function FileInUse(ByVal strPathName As String) As Boolean
    Dim hFile As Long
    
    On Error Resume Next
    '
    'Remove any trailing directory separator character
    If Right$(strPathName, 1) = gstrSEP_DIR Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    hFile = CreateFile(strPathName, GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL Or FILE_FLAG_WRITE_THROUGH, 0)
    
    If hFile = INVALID_HANDLE_VALUE Then
        If (Err.LastDllError = ERROR_SHARING_VIOLATION) Then
            FileInUse = True
        ElseIf (Err.LastDllError <> 0) Then
            Sleep 100
            Err.Clear
            hFile = CreateFile(strPathName, GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL Or FILE_FLAG_WRITE_THROUGH, 0)
            If hFile = INVALID_HANDLE_VALUE Then
                FileInUse = (Err.LastDllError = ERROR_SHARING_VIOLATION)
            Else
                CloseHandle hFile
            End If
        End If
    Else
        CloseHandle hFile
    End If
    Err.Clear
End Function

Public Sub SetFocusTo(nControl As Variant)
    
    On Error Resume Next
    If GetActiveFormHwnd = GetFormHwnd(nControl.hWnd) Then
        If VarType(nControl) = vbLong Then
            SetFocusAPI nControl
        Else
            nControl.SetFocus
        End If
    End If
    On Error GoTo 0
End Sub

Private Function GetFormHwnd(nControlHwnd As Long)
    Dim lPar As Long
    Dim iHwnd As Long
    
    iHwnd = nControlHwnd
    lPar = GetParent(iHwnd)
    While lPar <> 0
        iHwnd = lPar
        lPar = GetParent(lPar)
    Wend
    GetFormHwnd = iHwnd
End Function

Public Sub SendKeysAPI(ByVal sKeys As String, Optional ByVal Wait As Boolean)
    Dim cSK As New cSendKeys
    
    cSK.SendKeys sKeys, Wait
End Sub

Public Function GetSystemColorDepth() As Long
    Dim nPlanes As Long, BitsPerPixel As Long, dc As Long
    
    dc = GetDC(0)
    nPlanes = GetDeviceCaps(dc, Planes)
    BitsPerPixel = GetDeviceCaps(dc, BITSPIXEL)
    ReleaseDC 0, dc
    GetSystemColorDepth = nPlanes * BitsPerPixel
End Function

Public Function PictureOfWindowSection(nHwnd As Long, nSection As RECT) As StdPicture
    Dim hdcSrc As Long
    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim Pic As PicBmp
    Dim iPic As IPicture
    Dim IID_IDispatch As GUID
    Dim iWidth As Long
    Dim iHeight As Long

    iWidth = nSection.Right - nSection.Left
    iHeight = nSection.Bottom - nSection.Top

    hdcSrc = GetWindowDC(nHwnd)
    hDCMemory = CreateCompatibleDC(hdcSrc)
    hBmp = CreateCompatibleBitmap(hdcSrc, iWidth, iHeight)
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    Call BitBlt(hDCMemory, 0, 0, iWidth, iHeight, hdcSrc, nSection.Left, _
        nSection.Top, vbSrcCopy)
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    Call DeleteDC(hDCMemory)
    Call ReleaseDC(nHwnd, hdcSrc)

    'fill in OLE IDispatch Interface ID
    With IID_IDispatch
       .Data1 = &H20400
       .Data4(0) = &HC0
       .Data4(7) = &H46
     End With

    'fill Pic with necessary parts
    With Pic
       .Size = Len(Pic)         'Length of structure
       .Type = vbPicTypeBitmap  'Type of Picture (bitmap)
       .hBmp = hBmp             'Handle to bitmap
       .hPal = 0&               'Handle to palette (may be null)
     End With

    Call OleCreatePictureIndirect(Pic, IID_IDispatch, 1, iPic)

    'return the new Picture object
    Set PictureOfWindowSection = iPic
End Function

Public Function PictureOfWindow(nHwnd As Long) As StdPicture
    Dim iRect As RECT
    
    GetWindowRect nHwnd, iRect
    iRect.Right = iRect.Right - iRect.Left
    iRect.Bottom = iRect.Bottom - iRect.Top
    iRect.Left = 0
    iRect.Top = 0
    Set PictureOfWindow = PictureOfWindowSection(nHwnd, iRect)
End Function

Public Function IsMouseButtonPressed(nButton As vbExMouseButtonsConstants) As Boolean
    Dim iButton As Long
    
    iButton = nButton
    If GetSystemMetrics(SM_SWAPBUTTON) <> 0 Then
        If nButton = vxMBLeft Then
            iButton = VK_RBUTTON
        ElseIf nButton = vxMBRight Then
            iButton = VK_LBUTTON
        End If
    End If
    IsMouseButtonPressed = GetAsyncKeyState(iButton) <> 0
End Function
    
Public Sub InitGlobal()
    Static sInitialized As Boolean
    
    If Not sInitialized Then
        gButtonsStyle = -1
        gToolbarsButtonsStyle = vxInstallShieldToolbar
        gToolbarsDefaultIconsSize = vxIconsMedium
        sInitialized = True
        mFormsTracker.Update  ' to set the current monitor with mouse location as soon as possible when the program starts
    End If
End Sub

Public Function ColorBrightNess(nColor As Long) As Long
    Dim iR As Integer
    Dim iG As Integer
    Dim iB As Integer
    
    iR = nColor And 255
    iG = (nColor \ 256) And 255
    iB = (nColor \ 65536) And 255
    
    ColorBrightNess = Int((iR + iG + iB) / 3) * 100 / 255
    
End Function

Public Sub ChDirEx(nPath As String)
    SetCurrentDirectory nPath
End Sub

Public Function GetFileDateTimesUnsigned(nPath As String, nCreatedLow As Long, nCreatedHigh As Long, nLastAccessLow As Long, nLastAccessHigh As Long, nModifiedLow As Long, nModifiedHigh As Long) As Boolean
    Dim hFile As Long, rval As Long
    Dim buff As OFSTRUCT
    Dim ctime As FileTime, atime As FileTime, wtime As FileTime
    
    'Open the File for Reading
    hFile = OpenFile(nPath, buff, OF_READ)
    If hFile Then
        'Get File time
        rval = GetFileTime(hFile, ctime, atime, wtime)
        
        nCreatedLow = ctime.dwLowDateTime
        nCreatedHigh = ctime.dwHighDateTime
        
        nLastAccessLow = atime.dwLowDateTime
        nLastAccessHigh = atime.dwHighDateTime
    
    
        nModifiedLow = wtime.dwLowDateTime
        nModifiedHigh = wtime.dwHighDateTime
        
        GetFileDateTimesUnsigned = True
    End If
    rval = CloseHandle(hFile)
    
End Function

Public Function SetFileDateTimesUnsigned(nPath As String, nCreatedLow As Long, nCreatedHigh As Long, nLastAccessLow As Long, nLastAccessHigh As Long, nModifiedLow As Long, nModifiedHigh As Long) As Boolean
    Dim iFileHandle As Long
    Dim iFTCreated As FileTime
    Dim iFTAccess As FileTime
    Dim iFTModified As FileTime

    SetFileDateTimesUnsigned = False

    ' Open the file.
    iFileHandle = CreateFile(nPath, GENERIC_WRITE, _
        FILE_SHARE_READ Or FILE_SHARE_WRITE, _
        0&, OPEN_EXISTING, 0&, 0&)
    If iFileHandle = 0 Then Exit Function

    iFTCreated.dwLowDateTime = nCreatedLow
    iFTCreated.dwHighDateTime = nCreatedHigh
    iFTAccess.dwLowDateTime = nLastAccessLow
    iFTAccess.dwHighDateTime = nLastAccessHigh
    iFTModified.dwLowDateTime = nModifiedLow
    iFTModified.dwHighDateTime = nModifiedHigh

    ' Set the times.
    If SetFileTime(iFileHandle, iFTCreated, iFTAccess, iFTModified) = 0 Then
        CloseHandle iFileHandle
        Exit Function
    End If

    ' Close the file.
    If CloseHandle(iFileHandle) = 0 Then Exit Function

    SetFileDateTimesUnsigned = True
End Function

Public Sub DateTimeToUnsigned(ByVal nDateTime As Date, nLow As Long, nHigh As Long)
    Dim iSt As SYSTEMTIME
    Dim iFT As FileTime
    
    ' Convert the Date into a SYSTEMTIME.
    iSt = DateToSystemTime(nDateTime)
    
    ' Convert the SYSTEMTIME into a FILETIME.
    SystemTimeToFileTime iSt, iFT
    nLow = iFT.dwLowDateTime
    nHigh = iFT.dwHighDateTime
End Sub

Private Function DateToSystemTime(ByVal the_date As Date) As SYSTEMTIME
    With DateToSystemTime
        .wYear = Year(the_date)
        .wMonth = Month(the_date)
        .wDay = Day(the_date)
        .wHour = Hour(the_date)
        .wMinute = Minute(the_date)
        .wSecond = Second(the_date)
    End With
End Function


Public Function FontExists(FontName As String) As Boolean
    Dim oFont As New StdFont
    Dim bAns As Boolean
    
    oFont.Name = FontName
    bAns = StrComp(FontName, oFont.Name, vbTextCompare) = 0
    FontExists = bAns
End Function

Public Sub SaveTextFile(nPath As String, nText As String)
    Dim iFreeFile
    
    If FileExists(nPath) Then
        Err.Raise 867, , "File already exists"
    Else
        On Error Resume Next
        iFreeFile = FreeFile
        Open nPath For Output As #iFreeFile
        Print #iFreeFile, nText
        Close #iFreeFile
    End If
End Sub

Public Function GetOwnerForm(nForm As Object, nForms As Object) As Object
    Dim iHwndOwner As Long
    Dim iFrm As Form
    
    If nForm Is Nothing Then Exit Function
    If nForms Is Nothing Then Exit Function
    
    iHwndOwner = GetWindowLong(nForm.hWnd, GWL_HWNDPARENT)
    
    For Each iFrm In nForms
        If iFrm.hWnd = iHwndOwner Then
            If iFrm.Enabled Then
                Set GetOwnerForm = iFrm
            End If
            Exit For
        End If
    Next
End Function

Public Function GetOwnerForm2(nForm As Object, nForms As Object) As Object
    Dim iForm As Object
    
    Set iForm = GetOwnerForm(nForm, nForms)
    Do Until iForm Is Nothing
        Set GetOwnerForm2 = iForm
        Set iForm = GetOwnerForm(iForm, nForms)
    Loop
    If GetOwnerForm2 Is Nothing Then
        Set GetOwnerForm2 = nForm
    End If
End Function

Public Function GetTopZOrderForm(nForms As Object, Optional nFromOwnerForm, Optional nVisible As Boolean = True) As Object
    Dim iHwnd As Long
    Dim iFrm As Object
    
    If IsMissing(nFromOwnerForm) Then
        If nForms.Count > 0 Then
            iHwnd = GetWindow(nForms(0).hWnd, GW_HWNDFIRST)
            Do Until iHwnd = 0
                For Each iFrm In nForms
                    If iFrm.hWnd = iHwnd Then
                        If nVisible Then
                            If IsWindowVisible(iFrm.hWnd) <> 0 Then
                                Set GetTopZOrderForm = iFrm
                                Exit Function
                            End If
                        Else
                            Set GetTopZOrderForm = iFrm
                            Exit Function
                        End If
                    End If
                Next
                iHwnd = GetWindow(iHwnd, GW_HWNDNEXT)
            Loop
        End If
    Else
        If nForms.Count > 0 Then
            iHwnd = GetWindow(nFromOwnerForm.hWnd, GW_HWNDFIRST)
            Do Until iHwnd = 0
                For Each iFrm In nForms
                    If iFrm.hWnd = iHwnd Then
                        If iFrm.hWnd <> nFromOwnerForm.hWnd Then
                            If GetOwnerForm2(iFrm, nForms) Is nFromOwnerForm Then
                                If nVisible Then
                                    If IsWindowVisible(iFrm.hWnd) <> 0 Then
                                        Set GetTopZOrderForm = iFrm
                                        Exit Function
                                    End If
                                Else
                                    Set GetTopZOrderForm = iFrm
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next
                iHwnd = GetWindow(iHwnd, GW_HWNDNEXT)
            Loop
        End If
    End If
End Function

Public Function GetTopZOrderFormHwnd(nForms As Object, Optional nFromOwnerFormHwnd As Long) As Long
    Dim iHwnd As Long
    Dim iFrm As Object
    
    If nFromOwnerFormHwnd = 0 Then
        If nForms.Count > 0 Then
            iHwnd = GetWindow(nForms(0).hWnd, GW_HWNDFIRST)
            Do Until iHwnd = 0
                For Each iFrm In nForms
                    If iFrm.hWnd = iHwnd Then
                        GetTopZOrderFormHwnd = iHwnd
                        Exit Function
                    End If
                Next
                iHwnd = GetWindow(iHwnd, GW_HWNDNEXT)
            Loop
        End If
    Else
        If nForms.Count > 0 Then
            iHwnd = GetWindow(nFromOwnerFormHwnd, GW_HWNDFIRST)
            Do Until iHwnd = 0
                For Each iFrm In nForms
                    If iFrm.hWnd = iHwnd Then
                        If iFrm.hWnd <> nFromOwnerFormHwnd Then
                            If GetOwnerForm2Hwnd(iFrm.hWnd, nForms) = nFromOwnerFormHwnd Then
                                GetTopZOrderFormHwnd = iFrm.hWnd
                                Exit Function
                            End If
                        End If
                    End If
                Next
                iHwnd = GetWindow(iHwnd, GW_HWNDNEXT)
            Loop
        End If
    End If
End Function

Public Function GetOwner(nHwnd As Long)
    GetOwner = GetWindowLong(nHwnd, GWL_HWNDPARENT)
End Function

Public Function GetOwnerFormHwnd(nFormHwnd As Long, nForms As Object) As Long
    Dim iHwndOwner As Long
    Dim iFrm As Object
    
    If nFormHwnd = 0 Then Exit Function
    
    iHwndOwner = GetWindowLong(nFormHwnd, GWL_HWNDPARENT)
    
    If iHwndOwner <> 0 Then
        For Each iFrm In nForms
            If iFrm.hWnd = iHwndOwner Then
                GetOwnerFormHwnd = iHwndOwner
                Exit For
            End If
        Next
    End If
End Function

Public Function GetOwnerForm2Hwnd(nFormHwnd As Long, nForms As Object) As Long
    Dim iFormHwnd As Long
    
    iFormHwnd = GetOwnerFormHwnd(nFormHwnd, nForms)
    Do Until iFormHwnd = 0
        GetOwnerForm2Hwnd = iFormHwnd
        iFormHwnd = GetOwnerFormHwnd(iFormHwnd, nForms)
    Loop
    If GetOwnerForm2Hwnd = 0 Then
        GetOwnerForm2Hwnd = nFormHwnd
    End If
End Function

Public Sub RaiseError(ByVal Number As Long, Optional ByVal Source, Optional ByVal Description, Optional ByVal HelpFile, Optional ByVal HelpContext)
    If InIDE Then
        On Error Resume Next
        Err.Raise Number, Source, Description, HelpFile, HelpContext
        MsgBox "Error " & Err.Number & ". " & Err.Description, vbCritical
    Else
        Err.Raise Number, Source, Description, HelpFile, HelpContext
    End If
End Sub

Public Function InIDE() As Boolean
    Static sValue As Long
    
    If sValue = 0 Then
        Err.Clear
        On Error Resume Next
        Debug.Print 1 / 0
        If Err.Number Then
            sValue = 1
        Else
            sValue = 2
        End If
        Err.Clear
    End If
    InIDE = (sValue = 1)
End Function

Public Sub ShowComponentHelp(nItem As String)
    Dim iFrm As frmComponentHelp
    
    Set iFrm = New frmComponentHelp
    iFrm.ShowItem nItem
    ShowModal iFrm
    Set iFrm = Nothing
End Sub

Public Sub AddModalForm(nHwnd As Long)
    Dim c As Long
    Dim iLngs() As Long
    Dim iUb As Long
    Dim c2 As Long
    
    If Not IsArrayDimmed(mModalFormsHwnd) Then
        ReDim mModalFormsHwnd(0)
    End If
    iUb = UBound(mModalFormsHwnd)
    If ExistsModalForm(nHwnd) Then
        If mModalFormsHwnd(iUb) <> nHwnd Then
            ReDim iLngs(iUb)
            c2 = 0
            For c = 1 To iUb
                If mModalFormsHwnd(c) <> nHwnd Then
                    c2 = c2 + 1
                    iLngs(c2) = mModalFormsHwnd(c)
                End If
            Next c
            iLngs(iUb) = nHwnd
            ReDim mModalFormsHwnd(iUb)
            For c = 1 To iUb
                mModalFormsHwnd(c) = iLngs(c)
            Next c
        End If
    Else
        ReDim Preserve mModalFormsHwnd(iUb + 1)
        mModalFormsHwnd(iUb + 1) = nHwnd
    End If
End Sub

Public Sub RemoveModalForm(nHwnd As Long)
    Dim iUb As Long
    Dim c As Long
    Dim c2 As Long
    
    iUb = UBound(mModalFormsHwnd)
    If mModalFormsHwnd(iUb) = nHwnd Then
        ReDim Preserve mModalFormsHwnd(iUb - 1)
    Else
        If ExistsModalForm(nHwnd) Then
            ReDim iLngs(iUb)
            c2 = 0
            For c = 1 To iUb
                If mModalFormsHwnd(c) <> nHwnd Then
                    c2 = c2 + 1
                    iLngs(c2) = mModalFormsHwnd(c)
                End If
            Next c
            iUb = UBound(iLngs)
            ReDim mModalFormsHwnd(iUb)
            For c = 1 To iUb
                mModalFormsHwnd(c) = iLngs(c)
            Next c
        End If
    End If
End Sub

Public Function TopModalFormHwnd() As Long
    If Not IsArrayDimmed(mModalFormsHwnd) Then
        ReDim mModalFormsHwnd(0)
    End If
    TopModalFormHwnd = mModalFormsHwnd(UBound(mModalFormsHwnd))
End Function

Public Function ExistsModalForm(nHwnd As Long) As Boolean
    Dim c As Long
    
    If Not IsArrayDimmed(mModalFormsHwnd) Then
        ReDim mModalFormsHwnd(0)
    End If
    
    For c = 1 To UBound(mModalFormsHwnd)
        If mModalFormsHwnd(c) = nHwnd Then
            ExistsModalForm = True
            Exit For
        End If
    Next c
End Function

Private Function IsArrayDimmed(nArray) As Boolean
    Dim c As Long
    
    On Error GoTo TheExit:
    c = UBound(nArray)
    IsArrayDimmed = True
    
    Exit Function

TheExit:
    
End Function
    
Public Function GetTopOwnerFormHwnd(nFormHwnd As Long) As Long
    Dim iHwnd As Long
    
    iHwnd = GetOwnerHwnd(nFormHwnd)
    Do Until iHwnd = 0
        If Not WindowIsForm(iHwnd) Then Exit Do
        GetTopOwnerFormHwnd = iHwnd
        iHwnd = GetOwnerHwnd(iHwnd)
    Loop
End Function

Public Function AddToList(nList, nValue, Optional nOnlyIfMissing As Boolean, Optional nFirstElement As Long = 0) As Boolean
    Dim i As Long
    Dim iAdd As Boolean
    
    If Not nOnlyIfMissing Then
        iAdd = True
    Else
        iAdd = Not IsInList(nList, nValue, nFirstElement)
    End If
    If iAdd Then
        i = UBound(nList) + 1
        ReDim Preserve nList(LBound(nList) To i)
        nList(i) = nValue
        AddToList = True
    End If
End Function

Public Function IsInList(nList, nValue, Optional nFirstElement As Long = 0, Optional nLastElement As Long = -1) As Boolean
    Dim c As Long
    
    If nLastElement = -1 Then
        nLastElement = UBound(nList)
    Else
        If nLastElement > UBound(nList) Then
            nLastElement = UBound(nList)
        End If
    End If
    
    For c = nFirstElement To nLastElement
        If nList(c) = nValue Then
            IsInList = True
            Exit For
        End If
    Next c
End Function

Public Function IndexInList(nList, nValue) As Long
    Dim c As Long
    
    IndexInList = LBound(nList) - 1
    For c = LBound(nList) To UBound(nList)
        If nList(c) = nValue Then
            IndexInList = c
            Exit For
        End If
    Next c
End Function

Public Sub AssignAccelerators(nObject As Object, Optional nAsignToLabels As Boolean)
    Dim iCtl As Control
    Dim c As Long
    Dim iAUsed As String
    Dim iControlsNames() As String
    Dim iSSTab_Index() As Long
    Dim iCaptions() As String
    Dim iDiffCharCount() As Long
    Dim iAlreadyAssigned() As Boolean
    Dim iTypeName As String
    Dim iAuxAssigned As Boolean
    Dim iAuxLetterAssigned As String
    Dim iSkip As Boolean
    Dim iCaption As String
    Dim iUseMnemonic As Boolean
    Dim iTabStop As Boolean
    
    ReDim iControlsNames(0)
    ReDim iCaptions(0)
    ReDim iSSTab_Index(0)
    ReDim iDiffCharCount(0)
    ReDim iAlreadyAssigned(0)
    
    For Each iCtl In nObject.Controls
        iTypeName = TypeName(iCtl)
        Select Case iTypeName
            Case "CommandButton", "CheckBox", "OptionButton", "ButtonEx", "Label"
                If Not ((iTypeName = "Label") And (Not nAsignToLabels)) Then
                    iCaption = Replace(iCtl.Caption, "&&", "")
                    If iCaption <> "" Then
                        If HasLettersOrNumbers(iCaption) Then
                            If InStr(iCaption, "&") = 0 Then
                                iSkip = iCtl.Tag = "na"
                                If Not iSkip Then
                                    iUseMnemonic = True
                                    On Error Resume Next
                                    iUseMnemonic = iCtl.UseMnemonic
                                    iSkip = Not iUseMnemonic
                                    If Not iSkip Then
                                        iTabStop = True
                                        iTabStop = iCtl.TabStop
                                        iSkip = Not iTabStop
                                    End If
                                    If Not iSkip Then
                                        iSkip = IsHiddenAtLeft(iCtl)
                                    End If
                                    On Error GoTo 0
                                End If
                                If Not iSkip Then
                                    AddToList iControlsNames, iCtl.Name
                                    AddToList iCaptions, iCaption
                                    AddToList iSSTab_Index, -1
                                    AddToList iDiffCharCount, -1
                                    AddToList iAlreadyAssigned, False
                                End If
                            End If
                        End If
                    End If
                End If
            Case "SSTab", "SSTabEx"
                For c = 0 To iCtl.Tabs - 1
                    If iCtl.TabCaption(c) <> "" Then
                        If InStr(iCtl.TabCaption(c), "&") = 0 Then
                            AddToList iControlsNames, iCtl.Name
                            AddToList iCaptions, iCtl.TabCaption(c)
                            AddToList iSSTab_Index, c
                            AddToList iDiffCharCount, -1
                            AddToList iAlreadyAssigned, False
                        End If
                    End If
                Next c
            Case Else
        End Select
    Next iCtl
    
    iAuxAssigned = True
    Do Until Not iAuxAssigned
        iAuxAssigned = False
        
        If UBound(iControlsNames) > 0 Then
            iAUsed = GetUsedAccelerators(nObject.Controls(iControlsNames(1)).Parent)
            For c = 1 To UBound(iControlsNames)
                iDiffCharCount(c) = DifferentCharactersCount((iCaptions(c)), iAUsed)
            Next c
        End If
        OrderVector iDiffCharCount, iControlsNames, iCaptions, iSSTab_Index, iAlreadyAssigned
        
        ' First the ones of the SSTabs
        For c = 1 To UBound(iControlsNames)
            If Not iAlreadyAssigned(c) Then
                If iSSTab_Index(c) <> -1 Then
                    nObject.Controls(iControlsNames(c)).TabCaption(iSSTab_Index(c)) = AssignAcceleratorToCaption(nObject.Controls(iControlsNames(c)).TabCaption(iSSTab_Index(c)), GetNotToUseAccelerators & iAUsed, False, iAuxLetterAssigned)
                    iAUsed = iAUsed & iAuxLetterAssigned
                    If InStr(nObject.Controls(iControlsNames(c)).TabCaption(iSSTab_Index(c)), "&") = 0 Then
                        nObject.Controls(iControlsNames(c)).TabCaption(iSSTab_Index(c)) = AssignAcceleratorToCaption(nObject.Controls(iControlsNames(c)).TabCaption(iSSTab_Index(c)), CommonButtonsAccelerators & iAUsed, True, iAuxLetterAssigned)
                        iAUsed = iAUsed & iAuxLetterAssigned
                    End If
                    iAlreadyAssigned(c) = True
                    iAuxAssigned = True
                    Exit For
                End If
            End If
        Next c
        
        ' Then the other controls
        If Not iAuxAssigned Then
            For c = 1 To UBound(iControlsNames)
                If Not iAlreadyAssigned(c) Then
                    If iSSTab_Index(c) = -1 Then
                        AssignAcceleratorToControl nObject.Controls(iControlsNames(c)), iAUsed
                    End If
                    iAlreadyAssigned(c) = True
                    iAuxAssigned = True
                    Exit For
                End If
            Next c
        End If
    Loop
    
End Sub

Private Function HasLettersOrNumbers(nTexto As String) As Boolean
    Dim iChr As String
    Dim c As Long
    
    For c = 1 To Len(nTexto)
        iChr = Mid$(nTexto, c, 1)
        If IsAlphaNumeric(iChr) Then
            HasLettersOrNumbers = True
            Exit Function
        End If
    Next c
    
End Function

Private Function DifferentCharactersCount(ByVal nText As String, Optional ByVal nUsed As String) As Long
    Dim c As Long
    Dim iChr As String
    
    nText = UCase$(nText)
    nUsed = UCase$(nUsed)
    For c = 1 To Len(nText)
        iChr = Mid$(nText, c, 1)
        If Not InStr(nUsed, iChr) Then
            If IsAlphaNumeric(iChr) Then
                DifferentCharactersCount = DifferentCharactersCount + 1
                nUsed = nUsed & iChr
            End If
        End If
    Next c
End Function

Public Function AssignAcceleratorToControl(nControl As Object, Optional nUsedLetters) As String
    Dim iAUsed As String
    Dim iLetterAssigned As String
    
    If InStr(nControl.Caption, "&") > 0 Then Exit Function
    If IsMissing(nUsedLetters) Then
        iAUsed = GetUsedAccelerators(nControl.Parent)
    Else
        iAUsed = nUsedLetters
    End If
    
    nControl.Caption = AssignAcceleratorToCaption(Replace(nControl.Caption, "&", ""), GetNotToUseAccelerators & "I" & iAUsed, , iLetterAssigned)
    If iLetterAssigned = "" Then
        nControl.Caption = AssignAcceleratorToCaption(Replace(nControl.Caption, "&", ""), GetNotToUseAccelerators & iAUsed, , iLetterAssigned)
    End If
    If Not IsMissing(nUsedLetters) Then
        nUsedLetters = nUsedLetters & iLetterAssigned
    End If
End Function

Private Function GetUsedAccelerators(nObject As Object) As String
    Dim iCtl As Control
    Dim iLng As Long
    Dim iStr As String
    Dim iAUsed As String
    Dim c As Long
    Dim iStr2 As String
    Dim iConsider As Boolean
    
    On Error Resume Next
    For Each iCtl In nObject.Controls
        If (TypeName(iCtl) = "SSTab") Or (TypeName(iCtl) = "SSTabEx") Then
            For c = 0 To iCtl.Tabs - 1
                iStr = iCtl.TabCaption(c)
                If iStr <> "" Then
                    iLng = InStr(iStr, "&")
                    If iLng > 0 Then
                        If Len(iStr) > iLng Then
                            iStr2 = UCase(Mid$(iStr, iLng + 1, 1))
                            If InStr(iAUsed, iStr2) = 0 Then
                                iAUsed = iAUsed & iStr2
                            End If
                        End If
                    End If
                End If
            Next c
        Else
            iStr = ""
            iStr = iCtl.Caption
            iConsider = (iStr <> "")
            If iConsider Then
                iConsider = iCtl.UseMnemonic
                If iConsider Then
                    iConsider = iCtl.TabStop
                    If iConsider Then
                        iConsider = Not IsHiddenAtLeft(iCtl)
                        If iConsider Then
                            iLng = InStr(iStr, "&")
                            If iLng > 0 Then
                                If Len(iStr) > iLng Then
                                    iStr2 = UCase(Mid$(iStr, iLng + 1, 1))
                                    If InStr(iAUsed, iStr2) = 0 Then
                                        iAUsed = iAUsed & iStr2
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next iCtl
    
    GetUsedAccelerators = iAUsed
End Function

Private Function IsHiddenAtLeft(nControl As Object) As Boolean
    Dim iParent As Object
    Dim iContainer As Object
    Dim iObj As Object
    
    On Error Resume Next
    IsHiddenAtLeft = nControl.Left <= -15000
    If Not IsHiddenAtLeft Then
        Set iParent = nControl.Parent
        Set iContainer = nControl.Container
        Do Until (iContainer Is iParent) Or (iContainer Is Nothing)
            IsHiddenAtLeft = iContainer.Left <= -15000
            If IsHiddenAtLeft Then Exit Do
            Set iObj = iContainer
            Set iContainer = Nothing
            Set iContainer = iObj.Container
        Loop
    End If
End Function

Private Function AssignAcceleratorToCaption(nCaption As String, nUsedLetters As String, Optional nAssignRepeatedIfNeccesary As Boolean = True, Optional nLetterAssigned As String) As String
    Dim c As Long
    Dim iLU As String
    Dim iCap As String
    Dim iChar As String
    
    If HasLetters(nCaption) Then
        AssignAcceleratorToCaption = nCaption
        iCap = LCase$(nCaption)
        iLU = LCase$(nUsedLetters)
        For c = 1 To Len(iCap)
            iChar = Mid$(iCap, c, 1)
            If IsLetter(iChar) Then
                If InStr(iLU, iChar) = 0 Then
                    AssignAcceleratorToCaption = ""
                    If c > 1 Then
                        AssignAcceleratorToCaption = Left$(nCaption, c - 1)
                    End If
                    AssignAcceleratorToCaption = AssignAcceleratorToCaption & "&"
                    nLetterAssigned = UCase(Left$(Right$(nCaption, Len(nCaption) - c + 1), 1))
                    AssignAcceleratorToCaption = AssignAcceleratorToCaption & Right$(nCaption, Len(nCaption) - c + 1)
                    Exit Function
                End If
            End If
        Next c
        If nAssignRepeatedIfNeccesary Then
            AssignAcceleratorToCaption = AssignAcceleratorToCaption(nCaption, CommonButtonsAccelerators)
        End If
    End If
End Function

Private Function CommonButtonsAccelerators() As String
    Dim iStr As String
    
    If mCommonButtonsAccelerators = "" Then
        mCommonButtonsAccelerators = UCase(GetAccelerator(GetLocalizedString(efnGUIStr_General_CloseButton_Caption)))
        iStr = UCase(GetAccelerator(GetLocalizedString(efnGUIStr_General_OKButton_Caption)))
        If InStr(mCommonButtonsAccelerators, iStr) = 0 Then
            mCommonButtonsAccelerators = mCommonButtonsAccelerators & iStr
        End If
        iStr = UCase(GetAccelerator(GetLocalizedString(efnGUIStr_General_CancelButton_Caption)))
        If InStr(mCommonButtonsAccelerators, iStr) = 0 Then
            mCommonButtonsAccelerators = mCommonButtonsAccelerators & iStr
        End If
    End If
    CommonButtonsAccelerators = mCommonButtonsAccelerators
End Function

Private Function GetNotToUseAccelerators()
    GetNotToUseAccelerators = "PGYJ" & CommonButtonsAccelerators
End Function

Private Function GetAccelerator(nCaption As String) As String
    Dim iPos As Long
    
    iPos = InStr(Replace(nCaption, "&&", ""), "&")
    If iPos > 0 Then
        If Len(nCaption) > iPos Then
            GetAccelerator = Mid(nCaption, iPos + 1, 1)
        End If
    End If
End Function

Private Function HasLetters(nTexto As String) As Boolean
    Dim c As Long
    Dim iChar As String
    
    For c = 1 To Len(nTexto)
        iChar = Mid$(nTexto, c, 1)
        If IsLetter(iChar) Then
            HasLetters = True
            Exit Function
        End If
    Next c
End Function

Private Function IsLetter(nCharacter As String) As Boolean
    Dim iAsc As Long
    
    If nCharacter = "" Then Exit Function
    
    iAsc = Asc(UCase$(nCharacter))
    
    If (iAsc >= 65) And (iAsc <= 90) Then
        IsLetter = True
    End If
End Function

Private Function IsAlphaNumeric(sChr As String) As Boolean
    IsAlphaNumeric = sChr Like "[0-9A-Za-z]"
End Function

Public Sub ResetCommonButtonsAccelerators()
    mCommonButtonsAccelerators = ""
End Sub

Public Sub BroadcastUILanguageChange(nLanguagePrev As Long)
    mUILangPrev = nLanguagePrev
    mEnumMode = efnEWM_BroadcastUILanguageChange
    EnumWindows AddressOf EnumCallback, 0
End Sub


Public Sub SelectInComboByItemData(nCombo As Control, nItemData As Long)
    Dim c As Long
    
    For c = 0 To nCombo.ListCount - 1
        If nCombo.ItemData(c) = nItemData Then
            nCombo.ListIndex = c
            Exit Sub
        End If
    Next c
End Sub

'*****************************************
' ValidFileName: returns a valid file name String

' The cases where an invalid file name can be returned are:
' 1) The DefaultFileName parameter is a null String ("") and ProposedFileName was completely invalid or also a null string.
' 2) The DefaultFileName parameter is a null String ("") and ProposedFileName was only an extension (".something").
' In both above cases it will return a null string (""). So if you set DefaultFileName to "" you should check that the returned value is not "". In all other cases some valid file name will be returned.
' Also consider that this function only returns a valid filename but it doesn't check if the name is available, because it doesn't deal with paths but just filenames, so you still need to check if the files already exists and do something to handle that.

' Parameters
' ProposedFileName: the String that is proposed to use as file name
' ReplacementChar (Optional): a character to be used as a replacement for invalid characters, the default is nothing (a null string)
' DefaultFileName (Optional): the file name that will be used when the is not file name
' ForOldFileFormat_8Dot3 (Optional): set it to True in the case that you need to support the old file format convention that was used in the D.O.S. or if for some reason you want to restrict the file name to that format
' AllowExtension (Optional): specifies if the string can supply not only the file name but also the file extension. The default is True.
' [out] HasExtension (Optional): It is a return value. It is passed ByRef, it returns if the file name included an Extension.
'*****************************************
Public Function ValidFileName(ByVal ProposedFileName As String, Optional ByVal ReplacementChar As String = "", Optional DefaultFileName As String = "[Untitled]", Optional ForOldFileFormat_8Dot3 As Boolean = False, Optional AllowExtension As Boolean = True, Optional ByRef HasExtension As Boolean) As String
    Dim iChar As String
    Dim c  As Long
    Dim iFlag As Long
    Dim iName As String
    Dim iExt As String
    Dim iDotPos As Long
    Dim iNameLen As Long
    Dim iExtLen As Long
    Dim iFileName As String
    
    If ForOldFileFormat_8Dot3 Then
        iFlag = GCT_SHORTCHAR
    Else
        iFlag = GCT_LFNCHAR
    End If
     
     ' Validate the replacement char (thanks Lavolpe)
    If ReplacementChar <> "" Then
        If Not ((PathGetCharType(AscW(ReplacementChar)) And iFlag) = iFlag) Then
            RaiseError 2069, App.Title & " ValidFileName function", "ReplacementChar is not valid."
            Exit Function
        End If
    End If
   
    ProposedFileName = Trim$(ProposedFileName)
    If InStr(ProposedFileName, "/") Then ProposedFileName = Replace(ProposedFileName, "/", "-")  ' to preserve a date formatting
    If InStr(ProposedFileName, """") Then ProposedFileName = Replace(ProposedFileName, """", "'") ' convert double quotes to single quotes to preserve quotation
    
    'strip out not allowed characters in all the file name:
    iFileName = ""
    For c = 1 To Len(ProposedFileName)
        iChar = Mid$(ProposedFileName, c, 1)
        If (PathGetCharType(AscW(iChar)) And iFlag) = iFlag Then
            iFileName = iFileName & iChar
        Else
            iFileName = iFileName & ReplacementChar
        End If
    Next c
    
    ' strip out illegal characters at the end:
    Do
        iChar = Right$(iFileName, 1)
        Select Case iChar
            Case " ", "."
                iFileName = Left(iFileName, Len(iFileName) - 1)
            Case Else
                Exit Do
        End Select
    Loop
    
    ' separate name and (optional) extension
    iDotPos = InStrRev(iFileName, ".")
    If iDotPos > 0 Then
        iName = Left(iFileName, iDotPos - 1)
        iExt = Mid(iFileName, iDotPos + 1)
    Else
        iName = iFileName
        iExt = ""
    End If
    
    ' strip out illegal characters at the beginning of the name
    For c = 1 To Len(iName)
        iChar = Left(iName, 1)
        Select Case iChar
            Case " ", "."
                iName = Mid(iName, 2)
            Case Else
                Exit For
        End Select
    Next c
    
    ' don't permit too long file names (or estensions in the case of ForOldFileFormat_8Dot3)
    iNameLen = Len(iName)
    iExtLen = Len(iExt)
    If ForOldFileFormat_8Dot3 Then
        If iNameLen > 8 Then
            iName = Left(iName, 8)
        End If
        If iExtLen > 3 Then
            iExt = Left(iExt, 3)
        End If
    Else
        If iExtLen > 0 Then
            If iNameLen > 258 Then
                iName = Left(iName, 258)
                iNameLen = 258
            End If
        End If
        If (iNameLen + iExtLen + 1) > 260 Then
            iExt = Left(iExt, 260 - iNameLen - 1)
        End If
    End If
    
    ' don't permit forbidden file names
    Select Case UCase(iName)
        Case "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            If ReplacementChar <> "" Then
                iName = iName & ReplacementChar
            Else
                iName = iName & "_"
            End If
    End Select
    If iName = "" Then
        'if there is not a valid file name, then use the default
        If DefaultFileName = "[Untitled]" Then
            iName = GetLocalizedString(efnGUIStr_mGlobals_ValidFileName_DefaultFileName)
        Else
            iName = DefaultFileName
        End If
    End If
    
    ' compose file name again
    ValidFileName = iName
    If iDotPos > 0 Then
        If AllowExtension Then
            If ValidFileName <> "" Then
                ValidFileName = ValidFileName & "." & iExt
                HasExtension = True
            End If
        End If
    End If

End Function

Public Sub SaveBinaryFile(nFilePath As String, nBytes() As Byte, Optional nDateTime)
    Dim iFreeFile As Long
    
    iFreeFile = FreeFile
    Open nFilePath For Binary Access Write As #iFreeFile
    Put #iFreeFile, , nBytes
    Close #iFreeFile
    
    If Not IsMissing(nDateTime) Then
        SetFileDateTime nFilePath, CDate(nDateTime)
    End If
End Sub

Private Sub SetFileDateTime(nFilePath As String, nDateTime As Date)
    Dim lTime As FileTime
    Dim hFile As Long
            
    lTime = GetFileDateTime(nDateTime)
    hFile = CreateFile(nFilePath, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    SetFileTime hFile, lTime, lTime, lTime
    CloseHandle hFile
End Sub

Private Function GetFileDateTime(ByVal aDate As Date) As FileTime
    Dim lTemp As SYSTEMTIME
    Dim lTime As FileTime
    
    VariantTimeToSystemTime aDate, lTemp
    SystemTimeToFileTime lTemp, lTime
    LocalFileTimeToFileTime lTime, GetFileDateTime
End Function

Public Sub SelectTxtOnGotFocus(nTextBox As Control)
    If nTextBox.SelStart = 0 Then
        If nTextBox.SelLength = 0 Then
            nTextBox.SelLength = Len(nTextBox.Text)
        End If
    End If
End Sub

Public Function ChrCount(nText As String, ByVal nCharW As Long) As Long
    Dim iStrs() As String
    
    iStrs = Split(nText, ChrW(nCharW))
    ChrCount = UBound(iStrs)
End Function

Public Function GetTwoDigitYearCenturyChange() As Long
    Dim y As Long
    Dim iYear As Long
    Dim iLast As Long
    Static sValue As Long
    
    If sValue = 0 Then
        iLast = 2028
        For y = 29 To 99
            iYear = Year(DateSerial(y, 1, 1))
            If iYear < iLast Then
                sValue = y
                Exit For
            End If
            iLast = iYear
        Next
    End If
    GetTwoDigitYearCenturyChange = sValue
End Function


Public Function DecimalSignAsc() As Long
    Dim iBuff As String * 100
    Dim iPos As Long
    Dim iRet As Long
    Static sValue As Long
    
    If sValue = 0 Then
        iRet = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, iBuff, 99)
        iPos = InStr(iBuff, Chr$(0))
        sValue = Asc(Left$(iBuff, iPos - 1))
    End If
    DecimalSignAsc = sValue
End Function

Public Function CallByNameEx(nObject As Object, nProcedureName As String, nCallType As VbCallType, nParametersArray As Variant)
    If IsArray(nParametersArray) Then
        CallByNameEx = CallByName2(nObject, nProcedureName, nCallType, nParametersArray)
    Else
        CallByNameEx = CallByName(nObject, nProcedureName, nCallType)
    End If
End Function

' Author: The trick: http://www.vbforums.com/showthread.php?866039&p=5315395&viewfull=1#post5315395
Private Function CallByName2(ByVal cObject As Object, ByRef sProcName As String, ByVal eCallType As VbCallType, vArgs As Variant) As Variant
    Dim hr      As Long
    Dim vLoc()
    
    vLoc = vArgs
    
    If InIDE Then
        hr = rtcCallByNameIDE(CallByName2, cObject, StrPtr(sProcName), eCallType, vLoc, &H409)
    Else
        hr = rtcCallByName(CallByName2, cObject, StrPtr(sProcName), eCallType, vLoc, &H409)
    End If
    
    
    If hr < 0 Then
        Err.Raise hr
    End If
    
End Function

Public Function IsWindowsVersionOrMore(nRequiredVersion As vbExWindowsVersion) As Boolean
    Static sPlatformID As Long
    Static sMajorVersion As Long
    Static sMinorVersion As Long
    
    If sMajorVersion = 0 Then
        Dim osinfo As OSVERSIONINFO
        Dim retvalue As Integer
        
        osinfo.dwOSVersionInfoSize = 148
        osinfo.szCSDVersion = Space$(128)
        retvalue = GetVersionEx(osinfo)
        
        sPlatformID = osinfo.dwPlatformID
        sMajorVersion = osinfo.dwMajorVersion
        sMinorVersion = osinfo.dwMinorVersion
    End If
    
    If nRequiredVersion = vx98 Then
        If sPlatformID = 2 Then ' if it is NT
            If sMajorVersion > 4 Then ' more than NT4 (win 2000, 2003, XP or Vista, etc)
                IsWindowsVersionOrMore = True
            End If
        Else ' if isn't NT
            If (sMajorVersion >= 4) And (sMinorVersion >= 10) Then  ' If it is 98, ME...
                IsWindowsVersionOrMore = True
            End If
        End If
    ElseIf nRequiredVersion = vx2000 Then
        If sPlatformID = 2 Then ' if it is NT
            If sMajorVersion >= 5 Then
                IsWindowsVersionOrMore = True
            End If
        End If
    ElseIf nRequiredVersion = vxXP Then
        If osinfo.dwPlatformID = 2 Then ' if it is NT
            If (sMajorVersion = 5) And (sMinorVersion >= 1) Or (sMajorVersion > 5) Then
                IsWindowsVersionOrMore = True
            End If
        End If
    ElseIf nRequiredVersion = vxVista Then
        If osinfo.dwPlatformID = 2 Then ' if it is NT
            If sMajorVersion >= 6 Then
                IsWindowsVersionOrMore = True
            End If
        End If
    ElseIf nRequiredVersion = vx7 Then
        If osinfo.dwPlatformID = 2 Then ' if it is NT
            If sMajorVersion > 6 Then
                IsWindowsVersionOrMore = True
            Else
                If sMajorVersion = 6 Then
                    If sMinorVersion >= 1 Then
                        IsWindowsVersionOrMore = True
                    End If
                End If
            End If
        End If
    ElseIf nRequiredVersion = vx8 Then
        If osinfo.dwPlatformID = 2 Then ' if it is NT
            If sMajorVersion > 6 Then
                IsWindowsVersionOrMore = True
            Else
                If sMajorVersion = 6 Then
                    If sMinorVersion >= 2 Then
                        IsWindowsVersionOrMore = True
                    End If
                End If
            End If
        End If
    ElseIf nRequiredVersion = vx81 Then
        If osinfo.dwPlatformID = 2 Then ' if it is NT
            If sMajorVersion > 6 Then
                IsWindowsVersionOrMore = True
            Else
                If sMajorVersion = 6 Then
                    If sMinorVersion >= 3 Then
                        IsWindowsVersionOrMore = True
                    End If
                End If
            End If
        End If
    ElseIf nRequiredVersion = vx10 Then
        If osinfo.dwPlatformID = 2 Then ' if it is NT
            If sMajorVersion > 10 Then
                IsWindowsVersionOrMore = True
            End If
        End If
    End If
End Function

Public Sub ShowForm(nForm As Object, Optional Modal As vbExShowFormConstants = vbModeless, Optional OwnerForm)
    Dim iHwndOwnerForm As Long
    Dim iMonitorOwner As Long
    Dim iMonitorForm As Long
    Dim iMICurrent As MONITORINFO
    Dim iMIForm As MONITORINFO
    Dim iLng As Long
    
    If Not IsMissing(OwnerForm) Then
        If Not OwnerForm Is Nothing Then
            On Error Resume Next
            iHwndOwnerForm = OwnerForm.hWnd
            On Error GoTo 0
        End If
    End If
    
    If iHwndOwnerForm <> 0 Then
        iMonitorOwner = MonitorFromWindow(iHwndOwnerForm, MONITOR_DEFAULTTOPRIMARY)
        If iMonitorOwner <> 0 Then
            iMonitorForm = MonitorFromWindow(nForm.hWnd, MONITOR_DEFAULTTOPRIMARY)
            If iMonitorForm <> iMonitorOwner Then
                iMICurrent.cbSize = Len(iMICurrent)
                iMIForm.cbSize = Len(iMIForm)
                GetMonitorInfo iMonitorOwner, iMICurrent
                GetMonitorInfo iMonitorForm, iMIForm
                If ((iMICurrent.rcWork.Bottom - iMICurrent.rcWork.Top) <> 0) And ((iMIForm.rcWork.Bottom - iMIForm.rcWork.Top) <> 0) Then
                    nForm.Move nForm.Left + (iMICurrent.rcWork.Left - iMIForm.rcWork.Left) * Screen.TwipsPerPixelX, nForm.Top + (iMICurrent.rcWork.Top - iMIForm.rcWork.Top) * Screen.TwipsPerPixelY
                    If nForm.Left < (iMICurrent.rcWork.Left * Screen.TwipsPerPixelX) Then
                        nForm.Left = iMICurrent.rcWork.Left * Screen.TwipsPerPixelX
                    End If
                    If nForm.Top < (iMICurrent.rcWork.Top * Screen.TwipsPerPixelY) Then
                        nForm.Top = iMICurrent.rcWork.Top * Screen.TwipsPerPixelY
                    End If
                    If nForm.BorderStyle = vbSizable Then
                        If nForm.Height > (iMICurrent.rcWork.Bottom - iMICurrent.rcWork.Top) * Screen.TwipsPerPixelY Then
                            nForm.Height = (iMICurrent.rcWork.Bottom - iMICurrent.rcWork.Top) * Screen.TwipsPerPixelY
                        End If
                        If nForm.Width > (iMICurrent.rcWork.Right - iMICurrent.rcWork.Left) * Screen.TwipsPerPixelX Then
                            nForm.Width = (iMICurrent.rcWork.Right - iMICurrent.rcWork.Left) * Screen.TwipsPerPixelX
                        End If
                        
                        iLng = iMICurrent.rcWork.Right - nForm.Width / Screen.TwipsPerPixelX
                        If (nForm.Left / Screen.TwipsPerPixelX) > iLng Then
                            If MonitorFromPoint((nForm.Left + nForm.Width) / Screen.TwipsPerPixelX, (nForm.Top + nForm.Height) / Screen.TwipsPerPixelY, MONITOR_DEFAULTTONULL) = 0 Then ' if there is no monitor covering that point
                                nForm.Left = iLng * Screen.TwipsPerPixelX
                            End If
                        End If
                        iLng = iMICurrent.rcWork.Bottom - nForm.Height / Screen.TwipsPerPixelY
                        If (nForm.Top / Screen.TwipsPerPixelY) > iLng Then
                            If MonitorFromPoint((nForm.Left + nForm.Width) / Screen.TwipsPerPixelX, (nForm.Top + nForm.Height) / Screen.TwipsPerPixelY, MONITOR_DEFAULTTONULL) = 0 Then ' if there is no monitor covering that point
                                nForm.Top = iLng * Screen.TwipsPerPixelY
                            End If
                        End If
                    
                    End If
                End If
            End If
        End If
    End If
    
    If Modal = vbModalEx Then
        ShowModal nForm
    Else
        If MonitorCount > 1 And GetSetting(AppNameForRegistry, "MInfo", Base64Encode(nForm.Name) & ".MI", "0") = "0" Then
            iMonitorForm = MonitorFromWindow(nForm.hWnd, MONITOR_DEFAULTTONULL)
            If iMonitorForm <> 0 Then
                If mFormsTracker.CurrentMonitor <> iMonitorForm Then
                    iMICurrent.cbSize = Len(iMICurrent)
                    iMIForm.cbSize = Len(iMIForm)
                    GetMonitorInfo mFormsTracker.CurrentMonitor, iMICurrent
                    GetMonitorInfo iMonitorForm, iMIForm
                    If ((iMICurrent.rcWork.Bottom - iMICurrent.rcWork.Top) <> 0) And ((iMIForm.rcWork.Bottom - iMIForm.rcWork.Top) <> 0) Then
                        nForm.Move nForm.Left + (iMICurrent.rcWork.Left - iMIForm.rcWork.Left) * Screen.TwipsPerPixelX, nForm.Top + (iMICurrent.rcWork.Top - iMIForm.rcWork.Top) * Screen.TwipsPerPixelY
                        If nForm.Left < (iMICurrent.rcWork.Left * Screen.TwipsPerPixelX) Then
                            nForm.Left = iMICurrent.rcWork.Left * Screen.TwipsPerPixelX
                        End If
                        If nForm.Top < (iMICurrent.rcWork.Top * Screen.TwipsPerPixelY) Then
                            nForm.Top = iMICurrent.rcWork.Top * Screen.TwipsPerPixelY
                        End If
                        If nForm.BorderStyle = vbSizable Then
                            If nForm.Height > (iMICurrent.rcWork.Bottom - iMICurrent.rcWork.Top) * Screen.TwipsPerPixelY Then
                                nForm.Height = (iMICurrent.rcWork.Bottom - iMICurrent.rcWork.Top) * Screen.TwipsPerPixelY
                            End If
                            If nForm.Width > (iMICurrent.rcWork.Right - iMICurrent.rcWork.Left) * Screen.TwipsPerPixelX Then
                                nForm.Width = (iMICurrent.rcWork.Right - iMICurrent.rcWork.Left) * Screen.TwipsPerPixelX
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If WindowHasCaption(nForm.hWnd) Then
            mFormsTracker.AddForm nForm
        Else
            mFormsTracker.Update  ' to ensure the monitor set with mouse location with the first form
        End If
        If Modal = vbModeless Then
            If IsMissing(OwnerForm) Then
                nForm.Show
            Else
                nForm.Show , OwnerForm
            End If
        Else
            If IsMissing(OwnerForm) Then
                nForm.Show vbModal
            Else
                nForm.Show vbModal, OwnerForm
            End If
        End If
    End If
    
End Sub

Public Function GetActiveWindowHwnd() As Long
    GetActiveWindowHwnd = GetActiveFormHwnd
    If GetActiveWindowHwnd = 0 Then
        GetActiveWindowHwnd = GetForegroundWindow
        If GetWindowThreadProcessId(GetActiveWindowHwnd, 0&) <> App.ThreadID Then
            GetActiveWindowHwnd = 0
        End If
    End If
End Function

Public Function VirtualScreenLeft() As Long
    VirtualScreenLeft = GetSystemMetrics(SM_XVIRTUALSCREEN)
End Function

Public Function VirtualScreenTop() As Long
    VirtualScreenTop = GetSystemMetrics(SM_YVIRTUALSCREEN)
End Function

Public Function VirtualScreenWidth() As Long
    VirtualScreenWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
End Function

Public Function VirtualScreenHeight() As Long
    VirtualScreenHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
End Function

Public Function VirtualScreenRight() As Long
    VirtualScreenRight = GetSystemMetrics(SM_XVIRTUALSCREEN) + GetSystemMetrics(SM_CXVIRTUALSCREEN)
End Function

Public Function VirtualScreenBottom() As Long
    VirtualScreenBottom = GetSystemMetrics(SM_YVIRTUALSCREEN) + GetSystemMetrics(SM_CYVIRTUALSCREEN)
End Function
