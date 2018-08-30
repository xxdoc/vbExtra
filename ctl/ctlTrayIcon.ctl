VERSION 5.00
Begin VB.UserControl TrayIcon 
   CanGetFocus     =   0   'False
   ClientHeight    =   1224
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1164
   ClipBehavior    =   0  'None
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ctlTrayIcon.ctx":0000
   PropertyPages   =   "ctlTrayIcon.ctx":0E12
   ScaleHeight     =   1224
   ScaleWidth      =   1164
   ToolboxBitmap   =   "ctlTrayIcon.ctx":0E28
   Begin VB.Timer tmrRestoreWindow 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   72
      Top             =   540
   End
   Begin VB.Frame hWndHolder 
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
End
Attribute VB_Name = "TrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_BALLOONLCLK = &H405
Private Const WM_BALLOONRCLK = &H404
Private Const WM_BALLOONXCLK = WM_BALLOONRCLK

Private Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128
   dwState As Long
   dwStateMask As Long
   szInfo As String * 256
   uTimeout As Long
   szInfoTitle As String * 64
   dwInfoFlags As Long
End Type

'Private Const NOTIFYICON_VERSION = 3
'Private Const NOTIFYICON_OLDVERSION = 0
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
'Private Const NIM_SETFOCUS = &H3
'Private Const NIM_SETVERSION = &H4
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
'Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
'Private Const NIS_HIDDEN = &H1
'Private Const NIS_SHAREDICON = &H2
Private Const NIIF_NONE = &H0
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_INFO = &H1
'Private Const NIIF_GUID = &H4

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

Public Enum vbExTrayIconBalloonTipStyleConstants
    vxBTSNoIcon = NIIF_NONE
    vxBTSWarning = NIIF_WARNING
    vxBTSError = NIIF_ERROR
    vxBTSInfo = NIIF_INFO
End Enum

Public Enum vbExTrayIconMouseEventConstants
    vxMELeftButtonDown = WM_LBUTTONDOWN
    vxMELeftButtonUp = WM_LBUTTONUP
    vxMELeftButtonDoubleClick = WM_LBUTTONDBLCLK
    vxMELeftButtonClick = WM_LBUTTONUP
    vxMERightButtonDown = WM_RBUTTONDOWN
    vxMERightButtonUp = WM_RBUTTONUP
    vxMERightButtonDoubleClick = WM_RBUTTONDBLCLK
    vxMERightButtonClick = WM_RBUTTONUP
End Enum

Public Enum vbExTrayIconClickTypeConstants
    vxCTLeftClick = WM_BALLOONLCLK
    vxCTRightClick = WM_BALLOONRCLK
    vxCTXClick = WM_BALLOONXCLK
End Enum

Public Event TrayClick(Button As vbExTrayIconMouseEventConstants)
Public Event BalloonClick(ClickType As vbExTrayIconClickTypeConstants)
Public Event Restored(ByRef RemoveTrayIcon As Boolean)

Private mActive As Boolean
Private m_TrayIcon As StdPicture
Private m_IconData As NOTIFYICONDATA
Private mWM_RESTORE_FROM_SYSTEM_TRAY As Long
Private mParentHwnd As Long

Private Sub hWndHolder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim msg As Long
    msg = x / Screen.TwipsPerPixelX
    If msg >= WM_LBUTTONDOWN And msg <= WM_RBUTTONDBLCLK Then
        RaiseEvent TrayClick(msg)
    ElseIf msg >= WM_BALLOONXCLK And msg <= WM_BALLOONLCLK Then
        RaiseEvent BalloonClick(msg)
    End If
End Sub

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrPostProcess
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, wParam As Long, lParam As Long, bConsume As Boolean) As Long
    If iMsg = mWM_RESTORE_FROM_SYSTEM_TRAY Then
        tmrRestoreWindow.Enabled = True
    End If
End Function

Private Sub tmrRestoreWindow_Timer()
    Dim iBool As Boolean
    
    tmrRestoreWindow.Enabled = False
    ShowWindow mParentHwnd, SW_SHOW
    iBool = True
    RaiseEvent Restored(iBool)
    If iBool Then Remove
End Sub

Private Sub UserControl_Initialize()
    mWM_RESTORE_FROM_SYSTEM_TRAY = RegisterWindowMessage("WM_RESTORE_FROM_SYSTEM_TRAY")
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set m_TrayIcon = PropBag.ReadProperty("Icon", Nothing)
End Sub

Private Sub UserControl_Resize()
    UserControl.Size ScaleX(34, vbPixels, vbTwips), ScaleY(34, vbPixels, vbTwips)
End Sub

Public Property Let ToolTip(ByVal Caption As String)
    With m_IconData
        .szTip = Caption & vbNullChar
        .szInfo = "" & Chr(0)
        .szInfoTitle = "" & Chr(0)
        .dwInfoFlags = NIIF_NONE
        .uTimeout = 0
    End With
    Shell_NotifyIcon NIM_MODIFY, m_IconData
End Property

Public Property Get ToolTip() As String
    ToolTip = m_IconData.szTip
End Property

Public Sub Create(nToolTipText, Optional nIcon As StdPicture)
    If mActive Then Remove
    If Not nIcon Is Nothing Then Set m_TrayIcon = nIcon
    With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = hWndHolder.hWnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        If Not m_TrayIcon Is Nothing Then .hIcon = m_TrayIcon
        If Not IsMissing(nToolTipText) Then .szTip = nToolTipText & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = "" & Chr(0)
        .szInfoTitle = "" & Chr(0)
        .dwInfoFlags = NIIF_NONE
    End With
    Shell_NotifyIcon NIM_ADD, m_IconData
    mActive = True
    
    mParentHwnd = 0
    On Error Resume Next
    mParentHwnd = Parent.hWnd
    
    If mParentHwnd <> 0 Then
        AttachMessage Me, Parent.hWnd, mWM_RESTORE_FROM_SYSTEM_TRAY
    End If
End Sub

Public Sub Remove()
    Shell_NotifyIcon NIM_DELETE, m_IconData
    mActive = False
    If mParentHwnd <> 0 Then
        DetachMessage Me, Parent.hWnd, mWM_RESTORE_FROM_SYSTEM_TRAY
        mParentHwnd = 0
    End If
End Sub

Public Sub BalloonTip(Prompt As String, Optional Style As vbExTrayIconBalloonTipStyleConstants = vxBTSNoIcon, Optional Title As String, Optional Timeout As Long = 2000)
    If Title = Empty Then Title = App.Title
    If Prompt = Empty Then Prompt = " "
    With m_IconData
        .szInfo = Prompt & Chr(0)
        .szInfoTitle = Title & Chr(0)
        .dwInfoFlags = Style
        .uTimeout = Timeout
    End With
    Shell_NotifyIcon NIM_MODIFY, m_IconData
End Sub

Public Sub PopupMenu(Menu As Object, Optional Flags, Optional DefaultMenu)
    SetForegroundWindow Menu.Parent.hWnd
    If IsMissing(Flags) And IsMissing(DefaultMenu) Then
        Menu.Parent.PopupMenu Menu
    ElseIf IsMissing(Flags) Then
        Menu.Parent.PopupMenu Menu, , , , DefaultMenu
    Else
        Menu.Parent.PopupMenu Menu, Flags, , , DefaultMenu
    End If
End Sub

Property Get Icon() As IPictureDisp
    Set Icon = m_TrayIcon
End Property

Property Set Icon(ByVal nIcon As IPictureDisp)
Attribute Icon.VB_MemberFlags = "200"
    If nIcon Is Nothing Then
        Set m_TrayIcon = Nothing
    ElseIf nIcon.Handle = 0 Then
        Set m_TrayIcon = Nothing
    Else
        Set m_TrayIcon = nIcon
    End If
    
    PropertyChanged "nIcon"
    With m_IconData
        .hIcon = m_TrayIcon
        .szInfo = "" & Chr(0)
        .szInfoTitle = "" & Chr(0)
        .dwInfoFlags = NIIF_NONE
        .uTimeout = 0
    End With
    Shell_NotifyIcon NIM_MODIFY, m_IconData
End Property

Property Let Icon(nIcon As IPictureDisp)
    Set Icon = nIcon
End Property

Private Sub UserControl_Terminate()
    If mActive Then Remove
End Sub

Public Property Get Active() As Boolean
    Active = mActive
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Icon", m_TrayIcon, Nothing
End Sub
