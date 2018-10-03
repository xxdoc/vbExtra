VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFormPersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type WINDOWPLACEMENT
        Length As Long
        Flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type

Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetForegroundWindow& Lib "user32" (ByVal hWnd&)

'Private Const SW_MAXIMIZE As Long = 3
'Private Const WPF_RESTORETOMAXIMIZED As Long = &H2

Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private WithEvents mTimerPosition As cTimer
Attribute mTimerPosition.VB_VarHelpID = -1
Private mFormHwnd As Long

Private mName As String
Private mMe As Object
Private mNameAndContext As String
Private mCollection As cObjectHandlersCollection
Private mPersistMinimized As Boolean

Public Sub SetForm(nForm As Form, nForms As Object, Optional nConsiderContext As Boolean = True, Optional nCenterFormTheFirstTime As Boolean = True, Optional nInitialLeft, Optional nInitialTop, Optional nInitialWidth, Optional nInitialHeight, Optional nPersistLeft As Boolean = True, Optional nPersistTop As Boolean = True, Optional nPersistWidth As Boolean = True, Optional nPersistHeight As Boolean = True, Optional nMaxTop, Optional nPersistMinimizedState As Boolean, Optional nContext As String, Optional nParent As cObjectHandlersCollection)
    Dim iWP As WINDOWPLACEMENT
    Dim iRect1 As RECT
    Dim iRect2 As RECT
    Dim iWidth As Long
    Dim iHeight As Long
    Dim iWindowState As Long
    Dim ScrW As Long
    Dim ScrH As Long
    Dim iScrTPPX As Single
    Dim iScrTPPY As Single
    Dim iActiveForm As Object
    Dim iLng As Long
    
    If nForm Is Nothing Then Exit Sub
        
    Set mMe = Me
    
    Set mForm = nForm
    mFormHwnd = mForm.hWnd
    mName = mForm.Name
    If Not nParent Is Nothing Then
        Set mCollection = nParent
    End If
    
    If nConsiderContext Then
        Set iActiveForm = ScreenActiveForm(nForms, True)
        If Not iActiveForm Is Nothing Then
            If Not iActiveForm Is nForm Then
                Set iActiveForm = GetOwnerForm2(iActiveForm, nForms)
                If Not iActiveForm Is Nothing Then
                    mName = mName & "_" & iActiveForm.Name
                End If
            End If
        End If
    End If
    mNameAndContext = Base64Encode(mName & nContext)

    If GetSetting(AppNameForRegistry, "WindowsPos", mNameAndContext & ".ScrW", -1) = -1 Then
        If Not IsMissing(nInitialWidth) Then
            mForm.Width = nInitialWidth
        End If
        If Not IsMissing(nInitialHeight) Then
            mForm.Height = nInitialHeight
        End If
        If nCenterFormTheFirstTime Then
            CenterForm nForm
        End If
        If Not IsMissing(nInitialLeft) Then
            mForm.Left = nInitialLeft
        End If
        If Not IsMissing(nInitialTop) Then
            mForm.Top = nInitialTop
        End If
    End If
    
    ScrW = GetSetting(AppNameForRegistry, "WindowsPos", mNameAndContext & ".ScrW", Screen.Width)
    ScrH = GetSetting(AppNameForRegistry, "WindowsPos", mNameAndContext & ".ScrH", Screen.Height)
    
    If nPersistLeft Then
        iRect2.Left = GetSetting(AppNameForRegistry, "WindowsPos", mNameAndContext & ".L", mForm.Left)
    Else
        iRect2.Left = mForm.Left
    End If
    If nPersistTop Then
        iRect2.Top = GetSetting(AppNameForRegistry, "WindowsPos", mNameAndContext & ".T", mForm.Top)
    Else
        iRect2.Top = mForm.Top
    End If
    If nPersistWidth Then
        iWidth = GetSetting(AppNameForRegistry, "WindowsPos", mNameAndContext & ".W", mForm.Width)
    Else
        iWidth = mForm.Width
    End If
    If nPersistHeight Then
        iHeight = GetSetting(AppNameForRegistry, "WindowsPos", mNameAndContext & ".H", mForm.Height)
    Else
        iHeight = mForm.Height
    End If
    If Not IsMissing(nMaxTop) Then
        If nMaxTop = -1 Then
            nMaxTop = ScreenUsableHeight - iHeight
        End If
        If iRect2.Top > nMaxTop Then
            iRect2.Top = nMaxTop
        End If
    End If
    iScrTPPX = GetSetting(AppNameForRegistry, "WindowsPos", mNameAndContext & ".ScrTPPX", Screen.TwipsPerPixelX)
    iScrTPPY = GetSetting(AppNameForRegistry, "WindowsPos", mNameAndContext & ".ScrTPPX", Screen.TwipsPerPixelY)
    
    If (iWidth < ScrW) And (iWidth < Screen.Width) Then
        iRect2.Left = iRect2.Left / (ScrW - iWidth) * (Screen.Width - iWidth)
    Else
        iRect2.Left = 0
    End If
    If (iHeight < ScrH) And (iHeight < Screen.Height) Then
        iRect2.Top = iRect2.Top / (ScrH - iHeight) * (Screen.Height - iHeight)
    Else
        iRect2.Top = 0
    End If
    iRect2.Right = iRect2.Left + iWidth
    iRect2.Bottom = iRect2.Top + iHeight
    
    iRect1.Left = iRect2.Left / iScrTPPX
    iRect1.Right = iRect2.Right / iScrTPPX
    iRect1.Top = iRect2.Top / iScrTPPY
    iRect1.Bottom = iRect2.Bottom / iScrTPPY
    
    iWindowState = GetSetting(AppNameForRegistry, "WindowsPos", mNameAndContext & ".WS", 1000)
    iWidth = iRect1.Right - iRect1.Left
    iHeight = iRect1.Bottom - iRect1.Top
    
    If iWidth < 130 * 15 / iScrTPPX Then
        iWidth = 130 * 15 / iScrTPPX
        If (iWidth * iScrTPPX) > Screen.Width Then
            iWidth = Screen.Width / iScrTPPX
        End If
        iRect1.Right = iRect1.Left + iWidth
    End If
    If iHeight < 60 * 15 / iScrTPPY Then
        iHeight = 60 * 15 / iScrTPPY
        If iHeight * iScrTPPY > ScreenUsableHeight Then
            iHeight = ScreenUsableHeight / iScrTPPY
        End If
        iRect1.Bottom = iRect1.Top + iHeight
    End If
    If iRect1.Left < 0 Then
        iRect1.Left = 0
        iRect1.Right = iRect1.Left + iWidth
    End If
    If iRect1.Top < 0 Then
        iRect1.Top = 0
        iRect1.Bottom = iRect1.Top + iHeight
    End If
    
    mPersistMinimized = nPersistMinimizedState
    
    iLng = GetSetting(AppNameForRegistry, "WindowsPos", mNameAndContext & ".FL", 0)
    If (iLng > -1) And (iLng < 8) Then
        iWP.Flags = iLng
    End If
    iWP.Length = Len(iWP)
    iWP.rcNormalPosition.Left = iRect1.Left
    iWP.rcNormalPosition.Top = iRect1.Top
    If (mForm.BorderStyle = 2) Or (mForm.BorderStyle = 5) Then ' sizable
        iWP.rcNormalPosition.Right = iRect1.Right
        iWP.rcNormalPosition.Bottom = iRect1.Bottom
    Else
        iWP.rcNormalPosition.Right = iWP.rcNormalPosition.Left + mForm.Width / Screen.TwipsPerPixelX
        iWP.rcNormalPosition.Bottom = iRect1.Top + mForm.Height / Screen.TwipsPerPixelY
    End If
    If (iWindowState = vbMinimized) And nPersistMinimizedState Then
        iWP.showCmd = SW_MINIMIZE
    End If
    SetWindowPlacement mForm.hWnd, iWP
    
    Select Case iWindowState
        Case vbMaximized
            mForm.WindowState = vbMaximized
        Case 1000 ' no se sabe
        Case Else
            If (iWindowState = vbMinimized) And nPersistMinimizedState Then
                'mForm.WindowState = vbMinimized
            Else
                If (iWP.Flags And 2) = 2 Then
                    mForm.WindowState = vbMaximized
                Else
                    mForm.WindowState = vbNormal
                End If
            End If
            If mForm.WindowState = vbNormal Then
                If (mForm.BorderStyle = 2) Or (mForm.BorderStyle = 5) Then ' sizable
                    mForm.Move iRect1.Left * iScrTPPX, iRect1.Top * iScrTPPY, iWidth * iScrTPPX, iHeight * iScrTPPY
                Else
                    mForm.Move iRect1.Left * iScrTPPX, iRect1.Top * iScrTPPY, mForm.Width, mForm.Height
                End If
            End If
    End Select

End Sub

Private Sub CenterForm(frm As Form, Optional TrueCenter As Boolean = False)
    Dim iTop As Long
    Dim iLeft As Long
    
    If frm.WindowState <> vbNormal Then Exit Sub
    If TrueCenter Then
        iTop = Screen.Height \ 2 - frm.Height \ 2
        iLeft = Screen.Width \ 2 - frm.Width \ 2
    Else
        iTop = (Screen.Height * 0.9) \ 2 - frm.Height \ 2
        iLeft = Screen.Width \ 2 - frm.Width \ 2
    End If
    
    If iLeft < 0 Then iLeft = 0
    If iTop < 0 Then iTop = 0
    
    frm.Top = iTop
    frm.Left = iLeft
    
    If frm.BorderStyle = vbSizable Then
        If frm.Height > ScreenUsableHeight Then
            frm.Height = ScreenUsableHeight
        End If
        If frm.Width > Screen.Width Then
            frm.Width = Screen.Width
        End If
    End If
End Sub

Private Sub mForm_Activate()
    Static sFirst As Boolean
    
    If Not sFirst Then
        sFirst = True
        Set mTimerPosition = New cTimer
        mTimerPosition.Interval = 1
    End If
End Sub

Private Sub mForm_Unload(Cancel As Integer)
    Dim iHwndOwner As Long
    
    If Cancel = 0 Then
        SaveFormPersistence
        
        If Not mCollection Is Nothing Then
            mCollection.Remove mFormHwnd
            Set mCollection = Nothing
        End If
        
        If Not mTimerPosition Is Nothing Then
            mTimerPosition.Interval = 0
            Set mTimerPosition = Nothing
        End If
        
        Set mMe = Nothing
        Set mForm = Nothing
    
        ' All the following is to fix a VB6 bug, that set the focus to the first unowned form when the last owned form by a form is closed and not to itself
        If FormIsOwned(mFormHwnd) Then
            iHwndOwner = GetOwnerHwnd(mFormHwnd)
            If iHwndOwner <> 0 Then
                If NumberOfOwnedForms(iHwndOwner) = 1 Then ' (it is the last one)
                    SetForegroundWindow iHwndOwner
                End If
            End If
        End If
    End If
    
End Sub

Public Sub Unpersist()
    If Not mCollection Is Nothing Then
        mCollection.Remove mFormHwnd
        Set mCollection = Nothing
    End If
    Set mForm = Nothing
    Set mMe = Nothing
End Sub

Public Sub SaveFormPersistence()
    Dim iWP As WINDOWPLACEMENT
    Dim iWindowState As Long
    
    If IsWindow(mFormHwnd) <> 0 Then
        If GetWindowPlacement(mForm.hWnd, iWP) Then
            SaveSetting AppNameForRegistry, "WindowsPos", mNameAndContext & ".L", iWP.rcNormalPosition.Left * Screen.TwipsPerPixelX
            SaveSetting AppNameForRegistry, "WindowsPos", mNameAndContext & ".T", iWP.rcNormalPosition.Top * Screen.TwipsPerPixelY
            SaveSetting AppNameForRegistry, "WindowsPos", mNameAndContext & ".W", (iWP.rcNormalPosition.Right - iWP.rcNormalPosition.Left) * Screen.TwipsPerPixelX
            SaveSetting AppNameForRegistry, "WindowsPos", mNameAndContext & ".H", (iWP.rcNormalPosition.Bottom - iWP.rcNormalPosition.Top) * Screen.TwipsPerPixelY
            iWindowState = mForm.WindowState
            If iWindowState = vbMinimized Then
                If Not mPersistMinimized Then
                    If (iWP.Flags And 2) = 2 Then
                        iWindowState = vbMaximized
                    Else
                        iWindowState = vbNormal
                    End If
                End If
            End If
            SaveSetting AppNameForRegistry, "WindowsPos", mNameAndContext & ".WS", iWindowState
            SaveSetting AppNameForRegistry, "WindowsPos", mNameAndContext & ".FL", iWP.Flags
            SaveSetting AppNameForRegistry, "WindowsPos", mNameAndContext & ".ScrW", Screen.Width
            SaveSetting AppNameForRegistry, "WindowsPos", mNameAndContext & ".ScrH", Screen.Height
            SaveSetting AppNameForRegistry, "WindowsPos", mNameAndContext & ".ScrTPPX", Screen.TwipsPerPixelX
            SaveSetting AppNameForRegistry, "WindowsPos", mNameAndContext & ".ScrTPPY", Screen.TwipsPerPixelY
        End If
    End If
    
End Sub

Private Sub mTimerPosition_ThatTime()
    Dim iRect As RECT
    
    mTimerPosition.Interval = 0
    Set mTimerPosition = Nothing
    If Not mForm Is Nothing Then
        If IsWindow(mFormHwnd) <> 0 Then
            If IsWindowVisible(mFormHwnd) <> 0 Then
                GetWindowRect mFormHwnd, iRect
                If ((iRect.Left + 50) > (Screen.Width / Screen.TwipsPerPixelX)) Or ((iRect.Top + 70) > (Screen.Height / Screen.TwipsPerPixelY)) Then
                    CenterForm mForm
                End If
                If ((iRect.Right - 50) < 0) Or ((iRect.Bottom - 70) < 0) Then
                    CenterForm mForm
                End If
            End If
        End If
    End If
End Sub

Public Function GetFormPersistedWindowState() As Long
    GetFormPersistedWindowState = GetSetting(AppNameForRegistry, "WindowsPos", mNameAndContext & ".WS", 0)
End Function
