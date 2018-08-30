VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cActivatePrevInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Const WS_EX_APPWINDOW = &H40000

Private Type FileTime
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Private Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FileTime, lpExitTime As FileTime, lpKernelTime As FileTime, lpUserTime As FileTime) As Long

Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST + TH32CS_SNAPPROCESS + TH32CS_SNAPTHREAD + TH32CS_SNAPMODULE)
'Private Const TH32CS_INHERIT = &H80000000

Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformID As Long           '1 = Windows 95.
                                  '2 = Windows NT
   szCSDVersion As String * 128
End Type

Const MAX_PATH = 260

Private Type WINDOWPLACEMENT
        Length As Long
        Flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type

Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Private Const SW_SHOWMINIMIZED = 2
Private Const SW_RESTORE = 9

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function SetForegroundWindow& Lib "user32" (ByVal hWnd&)
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach&, ByVal idAttachTo&, ByVal bAttach&) As Long

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
    szExeFile As String * 260
End Type

Private Const PROCESS_VM_READ = 16
Private Const hNull = 0

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Private Const GW_HWNDNEXT = 2

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024

Const SMTO_BLOCK = &H1
Const SMTO_ABORTIFHUNG = &H2
Const WM_NULL = &H0
'Const WM_CLOSE = &H10
Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private mWM_RESTORE_FROM_SYSTEM_TRAY As Long

Private mCurrentProcessExe As String

Private Function BringAppOnFront(Optional nExePath As String, Optional nProcessID As Long, Optional nCaptionPart As String, Optional nTimeoutCheckHangedSeconds As Long, Optional nTerminateHanged As Boolean, Optional nWaitForHavingWindowsTimeoutSeconds As Long, Optional nTerminateWidndowless As Boolean, Optional nOnlyProcessesStartedEarlier As Boolean) As Boolean
    Dim WinEst As WINDOWPLACEMENT
    Dim rtn As Long
    Dim iHwnd As Long
    Dim iCaption As String
    Dim iC As Long
    Dim iAllCaptions() As String
    Dim iAllHwnds() As Long
    Dim iOtherThreadID As Long
    Dim iWindowExistent As Boolean
    Dim iT1
    Dim iT2
    Dim iTimes As Long
    Dim iLng As Long
    Dim hProcess As Long
    
    iT2 = timer
    
StartAgain:
    BringAppOnFront = True
    iHwnd = GetAppWindowHwnd(nExePath, nProcessID, iCaption, iAllCaptions, iAllHwnds, , nOnlyProcessesStartedEarlier)
    If UBound(iAllHwnds) = 0 Then
        Call GetAppWindowHwnd(nExePath, nProcessID, iCaption, iAllCaptions, iAllHwnds, , nOnlyProcessesStartedEarlier, True)
    End If
    If nCaptionPart <> "" Then
        For iC = 2 To UBound(iAllCaptions)
            If iAllCaptions(iC) <> "" Then
                If InStr(LCase$(iAllCaptions(iC)), LCase$(nCaptionPart)) Then
                    If iAllHwnds(iC) <> 0 Then
                        iHwnd = iAllHwnds(iC)
                        iCaption = iAllCaptions(iC)
                        Exit For
                    End If
                End If
            End If
        Next iC
    End If
    If UBound(iAllHwnds) = 0 Then
        If ((timer - iT2) > nWaitForHavingWindowsTimeoutSeconds) Or (nWaitForHavingWindowsTimeoutSeconds = 0) Then
            If nTerminateWidndowless Then
                iLng = GetProcessID(nExePath, , , True)
                Do While iLng <> 0
                    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0&, iLng)
                    TerminateProcess hProcess, 0&
                    CloseHandle hProcess
                    Sleep 500
                    iLng = GetProcessID(nExePath, , , True)
                Loop
                BringAppOnFront = False
                Exit Function
            Else
                BringAppOnFront = True
                Exit Function
            End If
        Else
            Sleep 500
            If Not PrevInstance Then
                BringAppOnFront = False
                Exit Function
            Else
                GoTo StartAgain:
            End If
        End If
    End If
    If mWM_RESTORE_FROM_SYSTEM_TRAY = 0 Then mWM_RESTORE_FROM_SYSTEM_TRAY = RegisterWindowMessage("WM_RESTORE_FROM_SYSTEM_TRAY")
    For iC = UBound(iAllHwnds) To 1 Step -1
        If iAllHwnds(iC) <> 0 Then
            PostMessage iAllHwnds(iC), mWM_RESTORE_FROM_SYSTEM_TRAY, 0, 0
            WinEst.Length = Len(WinEst)
            rtn = GetWindowPlacement(iAllHwnds(iC), WinEst)
'            If WinEst.showCmd = SW_SHOWMINIMIZED Then
            If nTimeoutCheckHangedSeconds > 0 Then
                iT1 = timer
                If Not WindowsIsResponding(iAllHwnds(iC), nTimeoutCheckHangedSeconds) Then
                    If nTerminateHanged Then
                        TerminateProcessHangedWindow iAllHwnds(iC)
                        Sleep 100
                        GoTo StartAgain:
                    Else
                        BringAppOnFront = False
                        Exit Function
                    End If
                End If
                If iTimes < 5 Then
                    If (timer - iT1) > 2 Then
                        iTimes = iTimes + 1
                        GoTo StartAgain:
                    End If
                End If
            End If
'            End If
            'Call SetWindowPos(iAllHwnds(iC), HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            If IsWindowVisible(iAllHwnds(iC)) = 0 Then
                WinEst.Length = Len(WinEst)
                rtn = GetWindowPlacement(iAllHwnds(iC), WinEst)
                If WinEst.showCmd = SW_SHOWMINIMIZED Then
                    ShowWindow iAllHwnds(iC), SW_SHOW
                    ShowWindow iAllHwnds(iC), SW_RESTORE
                End If
            End If
            Call SetWindowPos(iAllHwnds(iC), HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_ASYNCWINDOWPOS)
            iOtherThreadID = GetWindowThreadProcessId(iAllHwnds(iC), ByVal 0&)
            If iOtherThreadID <> App.ThreadID Then
                AttachThreadInput App.ThreadID, iOtherThreadID, 1
            End If
            SetForegroundWindow iAllHwnds(iC)
            If iOtherThreadID <> App.ThreadID Then
                AttachThreadInput App.ThreadID, iOtherThreadID, 0
            End If
        End If
    Next iC
    iWindowExistent = False
    For iC = 0 To UBound(iAllHwnds)
        If IsWindow(iAllHwnds(iC)) <> 0 Then
            iWindowExistent = True
            Exit For
        End If
    Next iC
    If Not iWindowExistent Then
        BringAppOnFront = False
        Exit Function
    End If
    If iHwnd <> 0 Then
        On Error Resume Next
        Do
            If IsIconic(iHwnd) Then
                ShowWindow iHwnd, SW_RESTORE
            End If
            If iCaption <> "" Then
                Err.Clear
                AppActivate iCaption
                If Err.Number <> 0 Then
                    Sleep 500
                    iC = iC + 1
                    If iC >= 3 Then
                        Exit Do
                    End If
                End If
            End If
        Loop While Err.Number <> 0
        
        If Err.Number <> 0 Then
            For iC = 2 To UBound(iAllCaptions)
                If iAllCaptions(iC) <> "" Then
                    iHwnd = iAllHwnds(iC)
                    WinEst.Length = Len(WinEst)
                    rtn = GetWindowPlacement(iHwnd, WinEst)
                    If WinEst.showCmd = SW_SHOWMINIMIZED Then
                        ShowWindow iHwnd, SW_RESTORE
                    End If
                    Call SetWindowPos(iAllHwnds(iC), HWND_TOPMOST, 0, 0, 0, 0, &H3)
                    Call SetWindowPos(iAllHwnds(iC), HWND_NOTOPMOST, 0, 0, 0, 0, &H3)
                    Err.Clear
                    AppActivate iAllCaptions(iC)
                    If Err.Number = 0 Then Exit Function
                    Sleep 500
                End If
            Next iC
        End If
    End If
End Function

Private Function GetAppWindowHwnd(Optional nExePath As String, Optional nProcessID As Long, Optional ByRef nCaption As String, Optional ByRef nAllCaptions As Variant, Optional ByRef nAllHwnds As Variant, Optional nAll As Boolean, Optional nOnlyProcessesStartedEarlier As Boolean, Optional nGetNotVisibleWindows As Boolean) As Long
    Dim iHwnds() As Long
    Dim iC As Long
    Dim iAuxCaption As String
    Dim iHwnd As Long
    Dim iAllCaptions() As String
    Dim iSendAllCaptions As Boolean
    Dim iAllHwnds() As Long
    Dim iSendAllHwnds As Boolean
    Dim iGetAppWindowHwnd As Long
    Dim iCaption As String
    Dim iC2 As Long
    Dim iProcessIDList() As Long
    Dim iAuxProcessID As Long
    Dim iDo As Boolean
    
    If Not IsMissing(nAllCaptions) Then
        iSendAllCaptions = True
        ReDim iAllCaptions(0)
    End If
    If Not IsMissing(nAllHwnds) Then
        iSendAllHwnds = True
        ReDim iAllHwnds(0)
    End If
    
    If nProcessID <> 0 Then
        ReDim iProcessIDList(1)
        iProcessIDList(1) = nProcessID
    Else
        If Trim$(nExePath) = "" Then Exit Function
        ReDim iProcessIDList(0)
        iC2 = 1
        Do
            iAuxProcessID = GetProcessID(Trim$(nExePath), , iC2, nOnlyProcessesStartedEarlier)
            If iAuxProcessID <> 0 Then
                ReDim Preserve iProcessIDList(iC2)
                iProcessIDList(iC2) = iAuxProcessID
            Else
                Exit Do
            End If
            iC2 = iC2 + 1
        Loop
    End If
    
    iGetAppWindowHwnd = 0
    iCaption = ""
    For iC2 = 1 To UBound(iProcessIDList)
        iHwnds = InstanceToWnd2(iProcessIDList(iC2), nAll)
        For iC = 1 To UBound(iHwnds)
            iDo = (IsWindowVisible(iHwnds(iC)))
            If Not iDo Then
                If nGetNotVisibleWindows Then
                    If WindowShowsInTaskbar(iHwnds(iC)) Then
                        iDo = True
                    ElseIf WindowIsForm(iHwnds(iC)) Then
                        iDo = True
                    End If
                End If
            End If
            If iDo Then
                If iGetAppWindowHwnd = 0 Then ' si pasa por primera vez
                    iHwnd = iHwnds(iC)
                    iAuxCaption = String(255, Chr$(0))
                    GetWindowText iHwnd, iAuxCaption, 255&
                    iAuxCaption = Left$(iAuxCaption, InStr(iAuxCaption, Chr(0)) - 1)
                    iAuxCaption = Trim$(iAuxCaption)
                    iCaption = iAuxCaption
                    iGetAppWindowHwnd = iHwnds(iC)
                End If
                If iSendAllCaptions Or iSendAllHwnds Then
                    iHwnd = iHwnds(iC)
                    If iSendAllCaptions Then
                        iAuxCaption = String(255, Chr$(0))
                        GetWindowText iHwnd, iAuxCaption, 255&
                        iAuxCaption = Left$(iAuxCaption, InStr(iAuxCaption, Chr(0)) - 1)
                        iAuxCaption = Trim$(iAuxCaption)
                        ReDim Preserve iAllCaptions(UBound(iAllCaptions) + 1)
                        iAllCaptions(UBound(iAllCaptions)) = iAuxCaption
                    End If
                    If iSendAllHwnds Then
                        ReDim Preserve iAllHwnds(UBound(iAllHwnds) + 1)
                        iAllHwnds(UBound(iAllHwnds)) = iHwnd
                    End If
                Else
                    Exit For
                End If
            End If
        Next iC
    Next iC2
    If iSendAllCaptions Then
        nAllCaptions = iAllCaptions
    End If
    If iSendAllHwnds Then
        nAllHwnds = iAllHwnds
    End If
    nCaption = iCaption
    GetAppWindowHwnd = iGetAppWindowHwnd
End Function

Private Function GetProcessID(strProcess As String, Optional nIgnoreCurrentProcess As Boolean = True, Optional nCheckExecutionInstance As Long = 1, Optional nOnlyProcessesStartedEarlier As Boolean) As Long
    Dim iCurrentProcessId As Long
    Dim iDo As Boolean
    Dim iExeInstanceCount As Long
    Dim iFT0 As FileTime
    Dim iFT1 As FileTime
    Dim iFTCP1 As FileTime
    
    iCurrentProcessId = GetCurrentProcessId
    If IsWindowsNT Then
        Dim cb As Long
        Dim cbNeeded As Long
        Dim NumElements As Long
        Dim ProcessIDs() As Long
        Dim cbNeeded2 As Long
        Dim Modules(1 To 200) As Long
        Dim LRet As Long
        Dim ModuleName As String
        Dim nSize As Long
        Dim hProcess As Long
        Dim i As Long
        
        If nOnlyProcessesStartedEarlier Then
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, iCurrentProcessId)
            GetProcessTimes hProcess, iFTCP1, iFT0, iFT0, iFT0
            CloseHandle (hProcess)
        End If
        
        'Get the array containing the process id's for each process object
        cb = 8
        cbNeeded = 96
        Do While cb <= cbNeeded
            cb = cb * 2
            ReDim ProcessIDs(cb / 4) As Long
            LRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
        Loop
        NumElements = cbNeeded / 4
        
        If nCheckExecutionInstance < 1 Then nCheckExecutionInstance = 1
        iExeInstanceCount = 0
        For i = 1 To NumElements
            iDo = True
            If nIgnoreCurrentProcess Then
                If ProcessIDs(i) = iCurrentProcessId Then
                    iDo = False
                End If
            End If
            If iDo Then
                'Get a handle to the Process
                hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
                Or PROCESS_VM_READ, 0, ProcessIDs(i))
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
                        ModuleName = Left$(ModuleName, InStr(ModuleName, Chr(0)) - 1)
                        If LCase$(ModuleName) = LCase$(strProcess) Then
                            iDo = True
                            If nOnlyProcessesStartedEarlier Then
                                GetProcessTimes hProcess, iFT1, iFT0, iFT0, iFT0
                                If SignedLongToUnsigned(iFT1.dwHighDateTime) > SignedLongToUnsigned(iFTCP1.dwHighDateTime) Then
                                    iDo = False
                                ElseIf SignedLongToUnsigned(iFT1.dwHighDateTime) = SignedLongToUnsigned(iFTCP1.dwHighDateTime) Then
                                    If SignedLongToUnsigned(iFT1.dwLowDateTime) > SignedLongToUnsigned(iFTCP1.dwLowDateTime) Then
                                        iDo = False
                                    End If
                                End If
                            End If
                            If iDo Then
                                iExeInstanceCount = iExeInstanceCount + 1
                                If iExeInstanceCount = nCheckExecutionInstance Then
                                    GetProcessID = ProcessIDs(i)
                                    'Close the handle to the process
                                    LRet = CloseHandle(hProcess)
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
                'Close the handle to the process
                LRet = CloseHandle(hProcess)
            End If
        Next
    Else
        Dim f As Long, sName As String
        Dim hSnap As Long, proc As PROCESSENTRY32
        
        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        If hSnap = hNull Then Exit Function
        proc.dwSize = Len(proc)
        ' Iterate through the processes
        f = Process32First(hSnap, proc)
        Do While f
            sName = StrZToStr(proc.szExeFile)
            If InStr(sName, Chr(0)) Then
                sName = Left$(sName, InStr(sName, Chr(0)) - 1)
            End If
            If LCase$(sName) = LCase$(strProcess) Then
                iDo = True
                If nIgnoreCurrentProcess Then
                    If proc.th32ProcessID = iCurrentProcessId Then
                        iDo = False
                    End If
                End If
                If iDo Then
                    GetProcessID = proc.th32ProcessID
                    Exit Function
                End If
            End If
            f = Process32Next(hSnap, proc)
        Loop
    End If
End Function


Private Function InstanceToWnd2(ByVal nProcessID As Long, Optional nAll As Boolean, Optional nIgnoreVB6MainWindows As Boolean = True) As Variant
    Dim iHwnds() As Long
    Dim test_hWnd As Long
    Dim test_pid As Long
    Dim test_thread_id As Long
    Dim iDo As Boolean
    
    ReDim iHwnds(0)
    'Find the first window
    test_hWnd = FindWindow(ByVal 0&, ByVal 0&)
    Do While test_hWnd <> 0
        'Check if the window isn't a child
        If (GetParent(test_hWnd) = 0) Or nAll Then
            'Get the window's thread
            test_thread_id = GetWindowThreadProcessId(test_hWnd, test_pid)
            If test_pid = nProcessID Then
                iDo = True
                If nIgnoreVB6MainWindows Then
                    Select Case GetWindowClassName(test_hWnd)
                        Case "ThunderRT6Main", "VBMsoStdCompMgr"
                            iDo = False
                    End Select
                End If
                If iDo Then
                    ReDim Preserve iHwnds(UBound(iHwnds) + 1)
                    iHwnds(UBound(iHwnds)) = test_hWnd
                End If
            End If
        End If
        'retrieve the next window
        test_hWnd = GetWindow(test_hWnd, GW_HWNDNEXT)
    Loop
    InstanceToWnd2 = iHwnds
End Function


Private Function IsWindowsNT() As Boolean
    Const dwMaskNT = &H2&
    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function

' FUNCTION: GetWinPlatform
' Get the current windows platform.
' ---------------------------------------------------------
Private Function GetWinPlatform() As Long
    Dim osvi As OSVERSIONINFO

    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
        Exit Function
    End If
    GetWinPlatform = osvi.dwPlatformID
End Function

Private Function StrZToStr(s As String) As String
    StrZToStr = Left$(s, Len(s) - 1)
End Function

Private Sub TerminateProcessHangedWindow(nHwndHangedWindow As Long)
    Dim lngProcessID As Long
    Dim lngReturnValue As Long
    Dim lngProcess As Long
    
    lngReturnValue = GetWindowThreadProcessId(nHwndHangedWindow, lngProcessID)
    lngProcess = OpenProcess(PROCESS_ALL_ACCESS, 0&, lngProcessID)
    lngReturnValue = TerminateProcess(lngProcess, 0&)
    CloseHandle lngProcess
End Sub

Private Function WindowsIsResponding(nHwnd As Long, nTimeoutSeconds As Long) As Boolean
    Dim lngResult As Long
    
    WindowsIsResponding = SendMessageTimeout(nHwnd, WM_NULL, 0&, 0&, SMTO_ABORTIFHUNG And SMTO_BLOCK, nTimeoutSeconds * 1000, lngResult)
End Function

Public Function PrevInstanceActivated(nTimeoutCheckHangedSeconds As Long, nTerminateHanged As Boolean, nWaitForHavingWindowsTimeoutSeconds As Long, nTerminateWindowless As Boolean) As Boolean
    PrevInstanceActivated = BringAppOnFront(CurrentProcessExe, , , nTimeoutCheckHangedSeconds, nTerminateHanged, nWaitForHavingWindowsTimeoutSeconds, nTerminateWindowless, True)
End Function


Private Function ExeByProcessId(nThreadId As Long) As String
    Dim cbNeeded2 As Long
    Dim Modules(1 To 200) As Long
    Dim ModuleName As String
    Dim nSize As Long
    Dim hProcess As Long
    Dim hSnapshot As Long, LRet As Long, p As PROCESSENTRY32
    p.dwSize = Len(p)
    
    If IsWindowsNT Then
        ' NT
        'Get a handle to the Process
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, nThreadId)
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
               ExeByProcessId = Left$(ModuleName, LRet)
            End If
            CloseHandle hProcess
        End If
    Else
        'Windows 95/98
        hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, ByVal 0)
        If hSnapshot Then
            LRet = Process32First(hSnapshot, p)
            Do While LRet
                If p.th32ProcessID = nThreadId Then
                    ExeByProcessId = Left$(p.szExeFile, InStr(p.szExeFile, Chr$(0)) - 1)
                    Exit Do
                End If
                LRet = Process32Next(hSnapshot, p)
            Loop
            LRet = CloseHandle(hSnapshot)
        End If
    End If
End Function

Public Property Get CurrentProcessExe() As String
'    If mCurrentProcessExe = "" Then
'        If InIDE Then
'            mCurrentProcessExe = "D:\Programas\vbExtra, sampes que no andan\App\ActivatePrevInstance\Project1.exe"
'        Else
            mCurrentProcessExe = ExeByProcessId(GetCurrentProcessId)
'        End If
'    End If
    CurrentProcessExe = mCurrentProcessExe
End Property

Public Property Get PrevInstance() As Boolean
        PrevInstance = Not (GetProcessID(CurrentProcessExe, , , True) = 0)
End Property

Public Function IsAppRunning(nAppPath As String) As Boolean
    IsAppRunning = Not (GetProcessID(nAppPath) = 0)
End Function

Private Function SignedLongToUnsigned(nValue As Long) As Double
    Const MaxUnsignedLongAnd1 As Double = 4294967296#
    
    ' function converts a signed long to an unsigned Double
    If (nValue And &H80000000) Then
        SignedLongToUnsigned = nValue + MaxUnsignedLongAnd1
    Else
        SignedLongToUnsigned = nValue
    End If
End Function

Public Function GetAllPrevInstancesMainWindowsHwnds() As Long()
    Dim c As Long
    Dim iLng As Long
    Dim iHwnds() As Long
    Dim iLng2 As Long
    
    ReDim iHwnds(0)
    c = 1
    iLng = GetProcessID(CurrentProcessExe, , c, False)
    Do Until iLng = 0
        iLng2 = 0
        iLng2 = GetAppWindowHwnd(, iLng)
        If iLng2 > 0 Then
            ReDim Preserve iHwnds(UBound(iHwnds) + 1)
            iHwnds(UBound(iHwnds)) = iLng2
        End If
        c = c + 1
        iLng = 0
        iLng = GetProcessID(CurrentProcessExe, , c, False)
    Loop
    GetAllPrevInstancesMainWindowsHwnds = iHwnds
End Function

Public Function GetAllPrevInstancesWindowsHwnds() As Long()
    Dim c As Long
    Dim iLng As Long
    Dim iHwnds() As Long
    Dim iLngs() As Long
    Dim c2 As Long
    
    ReDim iHwnds(0)
    c = 1
    iLng = GetProcessID(CurrentProcessExe, , c, False)
    Do Until iLng = 0
        ReDim iLngs(0)
        Call GetAppWindowHwnd(, iLng, , , iLngs)
        If UBound(iLngs) > 0 Then
            For c2 = 0 To UBound(iLngs)
                If iLngs(c2) <> 0 Then
                    ReDim Preserve iHwnds(UBound(iHwnds) + 1)
                    iHwnds(UBound(iHwnds)) = iLngs(c2)
                End If
            Next c2
        End If
        c = c + 1
        iLng = 0
        iLng = GetProcessID(CurrentProcessExe, , c, False)
    Loop
    GetAllPrevInstancesWindowsHwnds = iHwnds
End Function

Private Function WindowShowsInTaskbar(nHwnd As Long) As Boolean
    WindowShowsInTaskbar = ((GetWindowLong(nHwnd, GWL_EXSTYLE) And WS_EX_APPWINDOW) = WS_EX_APPWINDOW)
End Function

