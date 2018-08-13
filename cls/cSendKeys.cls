VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSendKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Public Enum KeyConstantsEx
'   VK_LWIN = &H5B 'Left Windows key (Microsoft® Natural® keyboard)
'   VK_RWIN = &H5C 'Right Windows key (Natural keyboard)
'   VK_APPS = &H5D 'Applications key (Natural keyboard)
'   VK_SLEEP = &H5F 'Computer Sleep key
'
'   VK_RMENU = &HA5 ' Right MENU key
'   VK_BROWSER_BACK = &HA6 'Windows 2000/XP: Browser Back key
'   VK_BROWSER_FORWARD = &HA7 'Windows 2000/XP: Browser Forward key
'   VK_BROWSER_REFRESH = &HA8 'Windows 2000/XP: Browser Refresh key
'   VK_BROWSER_STOP = &HA9 'Windows 2000/XP: Browser Stop key
'   VK_BROWSER_SEARCH = &HAA 'Windows 2000/XP: Browser Search key
'   VK_BROWSER_FAVORITES = &HAB 'Windows 2000/XP: Browser Favorites key
'   VK_BROWSER_HOME = &HAC 'Windows 2000/XP: Browser Start and Home key
'   VK_VOLUME_MUTE = &HAD 'Windows 2000/XP: Volume Mute key
'   VK_VOLUME_DOWN = &HAE  'Windows 2000/XP: Volume Down key
'   VK_VOLUME_UP = &HAF  'Windows 2000/XP: Volume Up key
'   VK_MEDIA_NEXT_TRACK = &HB0  'Windows 2000/XP: Next Track key
'   VK_MEDIA_PREV_TRACK = &HB1  'Windows 2000/XP: Previous Track key
'   VK_MEDIA_STOP = &HB2  'Windows 2000/XP: Stop Media key
'   VK_MEDIA_PLAY_PAUSE = &HB3  'Windows 2000/XP: Play/Pause Media key
'   VK_LAUNCH_MAIL = &HB4  'Windows 2000/XP: Start Mail key
'   VK_LAUNCH_MEDIA_SELECT = &HB5  'Windows 2000/XP: Select Media key
'   VK_LAUNCH_APP1 = &HB6  'Windows 2000/XP: Start Application 1 key
'   VK_LAUNCH_APP2 = &HB7  'Windows 2000/XP: Start Application 2 key
'   VK_OEM_1 = &HBA 'Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the ';:' key
'
'   VK_OEM_PLUS = &HBB 'Windows 2000/XP: For any country/region, the '+' key
'   VK_OEM_COMMA = &HBC 'Windows 2000/XP: For any country/region, the ',' key
'   VK_OEM_MINUS = &HBD 'Windows 2000/XP: For any country/region, the '-' key
'   VK_OEM_PERIOD = &HBE 'Windows 2000/XP: For any country/region, the '.' key
'   VK_OEM_2 = &HBF 'Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the '/?' key
'   VK_OEM_3 = &HC0 'Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the '`~' key
'
''—  C1–D7 Reserved
''—  D8–DA Unassigned
'   VK_OEM_4 = &HDB 'Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the '[{' key
'   VK_OEM_5 = &HDC 'Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the '\|' key
'   VK_OEM_6 = &HDD 'Used for miscellaneous characters; it can vary by keyboard Windows 2000/XP: For the US standard keyboard, the ']}' key
'   VK_OEM_7 = &HDE ' Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the 'single-quote/double-quote' key
'   VK_OEM_8 = &HDF 'Used for miscellaneous characters; it can vary by keyboard. —  E0 Reserved
''- E1 OEM specific
'   VK_OEM_102 = &HE2 'Windows 2000/XP: Either the angle bracket key or the backslash key on the RT 102-key keyboard
'' E3–E4 OEM specific
'   VK_PROCESSKEY = &HE5 'Windows 95/98/Me, Windows NT 4.0, Windows 2000/XP: IME PROCESS key
'' E6 OEM specific
'   VK_PACKET = &HE7 'Windows 2000/XP: Used to pass Unicode characters as if they were keystrokes. The VK_PACKET key is the low word of a 32-bit Virtual Key value used for non-keyboard input methods. For more information, see Remark in KEYBDINPUT, SendInput, WM_KEYDOWN, and WM_KEYUP
''—  E8 Unassigned
'' E9–F5 OEM specific
'   VK_ATTN = &HF6 'Attn key
'   VK_CRSEL = &HF7 'CrSel key
'   VK_EXSEL = &HF8 'ExSel key
'   VK_EREOF = &HF9 'Erase EOF key
'   VK_PLAY = &HFA 'Play key
'   VK_ZOOM = &HFB 'Zoom key
'   VK_NONAME = &HFC 'Reserved for future use
'   VK_PA1 = &HFD 'PA1 key
'   VK_OEM_CLEAR = &HFE 'Clear key
'End Enum

Private m_colKeyMap As New Collection

Private Declare Sub keybd_event Lib "user32" ( _
   ByVal bVk As Byte, ByVal bScan As Byte, _
   ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" ( _
   ByVal cChar As Byte) As Integer
Private Declare Function VkKeyScanW Lib "user32" ( _
   ByVal cChar As Integer) As Integer

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Function nextChar(ByRef sString As String, ByVal iPos As Long, Optional ByVal lLen As Long = 0) As String
   If (lLen = 0) Then lLen = Len(sString)
   If (iPos + 1 <= lLen) Then
      nextChar = Mid$(sString, iPos + 1, 1)
   End If
End Function

Public Sub SendKeys(ByVal sKeys As String, Optional ByVal Wait As Boolean)

   ' The plus sign (+), caret (^), percent sign (%),
   ' tilde (~), and parentheses ( ) have special
   ' meanings to SendKeys
   ' Brackets ([ ]) have no special meaning to SendKeys,
   ' but you must enclose them in braces.
   ' To specify brace characters, use {{} and {}}.
   ' Repeating keys: {left 42} do left 42 times.
   
   ' + = Shift
   ' ^ = Ctrl
   ' % = Alt
   ' ~ = enter
   ' ( = start sub expression. +(EC) = Shift then E then C
      
On Error GoTo errorHandler

   Dim sMsg As String
   Dim lErr As Long
   Dim iPos As Long
   Dim iNextPos As Long
   Dim iLen As Long
   Dim sChar As String
   Dim colBrace As New Collection
   Dim sContent As String
   Dim sKey As String
   Dim sCount As String
   Dim lCount As Long

   iPos = 1
   iLen = Len(sKeys)
   Do While iPos <= iLen
   
      sChar = Mid$(sKeys, iPos, 1)
      Select Case sChar
      Case "+", "~", "%"
         If nextChar(sKeys, iPos, iLen) = "(" Then
            ' Add to brace stack:
            colBrace.Add sChar
            ' send key down
            Select Case sChar
            Case "+"
               KeyDown vbKeyShift
            Case "!"
                KeyDown vbAltMask
            Case "~"
               KeyDown vbKeyReturn
            Case "^"
               KeyDown vbKeyControl
              ' Debug.Print "Control1"
            Case "%"
               KeyDown vbKeyMenu
            End Select
            iPos = iPos + 2
         Else
            ' Key press the key (probably not what you wanted)
            Select Case sChar
            Case "+"
               KeyDown vbKeyShift
               KeyUp vbKeyShift
            Case "!"
                KeyDown vbAltMask
                KeyUp vbAltMask
            Case "~"
               KeyDown vbKeyReturn
               KeyUp vbKeyReturn
            Case "^"
               KeyDown vbKeyControl
               KeyUp vbKeyControl
'               Debug.Print "Control2"
            Case "%"
               KeyDown vbKeyMenu
               KeyUp vbKeyMenu
            End Select
            iPos = iPos + 1
         End If
      
      Case "~"
         ' Enter key:
         KeyDown vbKeyReturn
         KeyUp vbKeyReturn
         iPos = iPos + 1
         
      Case ")"
         If (colBrace.Count > 0) Then
            sChar = colBrace(colBrace.Count)
            ' send key up
            Select Case sChar
            Case "+"
               KeyUp vbKeyShift
            Case "!"
                KeyUp vbAltMask
            Case "~"
               KeyUp vbKeyReturn
            Case "^"
               KeyUp vbKeyControl
'               Debug.Print "Control1"
            Case "%"
               KeyUp vbKeyMenu
            End Select
            colBrace.Remove colBrace.Count
            iPos = iPos + 1
         Else
            ' Invalid sendkeys command:
            sMsg = "Invalid sendkeys command: unmatched ) at position " & iPos
            GoTo errorHandler
         End If
      
      Case "{"
         ' special key
         If (iPos + 2 > iLen) Then
            sMsg = "Invalid sendkeys command; opening { without content or closing } at position " & iPos
            GoTo errorHandler
         Else
            iNextPos = InStr(iPos + 2, sKeys, "}")
            If (iNextPos = 0) Then
               sMsg = "Invalid sendkeys command; opening { without closing } at position " & iPos
               GoTo errorHandler
            Else
               sContent = Mid$(sKeys, iPos + 1, iNextPos - iPos - 1)
               iPos = iNextPos + 1
               ' is this a key/presses pair?
               iNextPos = InStr(sContent, " ")
               If (iNextPos > 0) Then
                  sKey = Left$(sContent, iNextPos - 1)
                  sCount = Mid$(sContent, iNextPos + 1)
                  If Not (IsNumeric(sCount)) Then
                     sMsg = "Invalid sendkeys command; key repetitions '" & sCount & "' is invalid near position " & iPos
                     lCount = CLng(sCount)
                  End If
               Else
                  sKey = sContent
                  lCount = 1
               End If
               KeyPress sKey, lCount
               
            End If
         End If
         
      Case Else
         ' send the key as is
        KeyPress sChar, 1
        iPos = iPos + 1
         
      End Select
   
   Loop
   
   If (colBrace.Count > 0) Then
      sMsg = "Invalid sendkeys command: more open brackets than close brackets."
      GoTo errorHandler
   End If
   
   Exit Sub
   
errorHandler:
   If Len(sMsg) = 0 Then
      sMsg = Err.Description
      lErr = Err.Number
   End If
   
   ' If we don't clear up the shift/control/alt keys,
   ' then you might find other apps on the system are hard to
   ' use.
   ' Make sure you have Break on Unhandled Errors switched
   ' on.
   Do While colBrace.Count > 0
      sChar = colBrace(colBrace.Count)
      ' send key up
      Select Case sChar
            Case "+"
               KeyUp vbKeyShift
            Case "!"
                KeyUp vbAltMask
            Case "~"
               KeyUp vbKeyReturn
            Case "^"
               KeyUp vbKeyControl
'               Debug.Print "Control1"
            Case "%"
               KeyUp vbKeyMenu
      End Select
      colBrace.Remove colBrace.Count
   Loop
   
   On Error GoTo 0
   Err.Raise lErr, App.EXEName & ".cSendKeys", sMsg
   
   Exit Sub

End Sub

Public Sub KeyPress(ByVal sKey As String, Optional ByVal lCount = 1)
Dim vKey As KeyCodeConstants
Dim l As Long

   On Error Resume Next
   vKey = m_colKeyMap(sKey)
   On Error GoTo 0
   
   If (vKey = 0) Then
      ' translate string into v key code
      vKey = KeyCode(sKey)
   End If
   
   If (vKey <> 0) Then
      For l = 1 To lCount
         If vKey <> (vKey And &HFF&) Then KeyDown vbKeyShift
         KeyDown vKey And &HFF&
         KeyUp vKey And &HFF&
         If vKey <> (vKey And &HFF&) Then KeyUp vbKeyShift
      Next l
   Else
      Err.Raise 9, , "Key " & sKey & " could not be interpreted."
   End If
   
End Sub

Public Sub KeyDown(ByVal vKey As KeyCodeConstants)
   keybd_event vKey, 0, KEYEVENTF_EXTENDEDKEY, 0
End Sub

Public Sub KeyUp(ByVal vKey As KeyCodeConstants)
   keybd_event vKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
End Sub

Public Function KeyCode(ByVal sChar As String) As KeyCodeConstants
Dim bNt As Boolean
Dim iKeyCode As Integer
Dim b() As Byte
Dim iKey As Integer
'Dim vKey As KeyCodeConstants
'Dim iShift As ShiftConstants

   ' Determine if we have Unicode support or not:
   bNt = ((GetVersion() And &H80000000) = 0)
   
   ' Get the keyboard scan code for the character:
   If (bNt) Then
      b = sChar
      CopyMemory iKey, b(0), 2
      iKeyCode = VkKeyScanW(iKey)
   Else
      b = StrConv(sChar, vbFromUnicode)
      iKeyCode = VkKeyScan(b(0))
   End If
   
'   KeyCode = (iKeyCode And &HFF&)
   KeyCode = iKeyCode 'And &HFF&)

End Function


Private Sub Class_Initialize()

   m_colKeyMap.Add vbKeyBack, "BACKSPACE"
   m_colKeyMap.Add vbKeyBack, "BS"
   m_colKeyMap.Add vbKeyBack, "BKSP"
   m_colKeyMap.Add vbKeyPause, "BREAK"
   m_colKeyMap.Add vbKeyCapital, "CAPSLOCK"
   m_colKeyMap.Add vbKeyDelete, "DELETE"
   m_colKeyMap.Add vbKeyDelete, "DEL"
   m_colKeyMap.Add vbKeyDown, "DOWN"
   m_colKeyMap.Add vbKeyEnd, "END"
   m_colKeyMap.Add vbKeyReturn, "ENTER"
   m_colKeyMap.Add vbKeyReturn, "~"
   m_colKeyMap.Add vbKeyEscape, "ESC"
   m_colKeyMap.Add vbKeyHelp, "HELP"
   m_colKeyMap.Add vbKeyHome, "HOME"
   m_colKeyMap.Add vbKeyInsert, "INS"
   m_colKeyMap.Add vbKeyInsert, "INSERT"
   m_colKeyMap.Add vbKeyLeft, "Left$"
   m_colKeyMap.Add vbKeyNumlock, "NUMLOCK"
   m_colKeyMap.Add vbKeyPageDown, "PGDN"
   m_colKeyMap.Add vbKeyPageUp, "PGUP"
   m_colKeyMap.Add vbKeyPrint, "PRTSC"
   m_colKeyMap.Add vbKeyRight, "Right$"
   m_colKeyMap.Add vbKeyScrollLock, "SCROLLLOCK"
   m_colKeyMap.Add vbKeyTab, "TAB"
   m_colKeyMap.Add vbKeyUp, "UP"
   m_colKeyMap.Add vbKeyF1, "F1"
   m_colKeyMap.Add vbKeyF2, "F2"
   m_colKeyMap.Add vbKeyF3, "F3"
   m_colKeyMap.Add vbKeyF4, "F4"
   m_colKeyMap.Add vbKeyF5, "F5"
   m_colKeyMap.Add vbKeyF6, "F6"
   m_colKeyMap.Add vbKeyF7, "F7"
   m_colKeyMap.Add vbKeyF8, "F8"
   m_colKeyMap.Add vbKeyF9, "F9"
   m_colKeyMap.Add vbKeyF10, "F10"
   m_colKeyMap.Add vbKeyF11, "F11"
   m_colKeyMap.Add vbKeyF12, "F12"
   m_colKeyMap.Add vbKeyF13, "F13"
   m_colKeyMap.Add vbKeyF14, "F14"
   m_colKeyMap.Add vbKeyF15, "F15"
   m_colKeyMap.Add vbKeyF16, "F16"

End Sub


