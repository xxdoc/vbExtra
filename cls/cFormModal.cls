VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFormModal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISubclass

Private Const WM_ACTIVATEAPP As Long = &H1C&
Private Declare Function SetForegroundWindow& Lib "user32" (ByVal hWnd&)
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function SetActiveWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Const ICON_SMALL As Long = 0&
'Private Const ICON_BIG As Long = 1&
Private Const WM_SETICON As Long = &H80&
Private Const WM_GETICON As Long = &H7F&

Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Private Const GWL_HWNDPARENT As Long = (-8)

Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

Private Const WS_EX_APPWINDOW = &H40000

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Private Const WM_SYSCOMMAND As Long = &H112&
Private Const SC_CLOSE As Long = &HF060&

Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private WithEvents mMinimizeApp As AppMinimizer
Attribute mMinimizeApp.VB_VarHelpID = -1
Private WithEvents mSetActiveWindow As cTimer
Attribute mSetActiveWindow.VB_VarHelpID = -1
'Private WithEvents mOwnerForm As Form

Private mEnabled() As Long
Private mHwnds() As Long
Private mMe As Object

Private mRgn As Long
Private WithEvents mTimerShow As cTimer
Attribute mTimerShow.VB_VarHelpID = -1
Private mOldOwner As Long
Private mOwnerFormHwnd As Long
Private mFormHwnd As Long
Private mFormWindowState As Long
Private mAttached As Boolean
Private mFormWasTheActiveWindow As Boolean
Private mFormToSetActiveHwnd As Long


Public Sub Show(nForm As Form, Optional nWaitWithDoevents As Boolean, Optional nSetIcon As Boolean = True, Optional nFormsHwndToKeepEnabled As Variant, Optional nKeepEnabledTaskBarWindows As Boolean = True, Optional nNoOwner As Boolean)
    Dim c As Long
    Dim iWindowStyle As Long
    Dim iFormHwnd As Long
    Dim iStyle As Long
    Dim iDoNotChangeStyle As Boolean
    Dim iFormsHwndToKeepEnabled() As Long
    Dim iKeepEnabled As Boolean
    Dim c2 As Long
    Dim iAuxFormsHwnds() As Long
    
    If ExistsModalForm(nForm.hWnd) Then
        AddModalForm nForm.hWnd
        nForm.ZOrder
        On Error Resume Next
        nForm.SetFocus
        Exit Sub
    End If
    
    If IsMissing(nFormsHwndToKeepEnabled) Then
        ReDim iFormsHwndToKeepEnabled(0)
    Else
        If (VarType(nFormsHwndToKeepEnabled) And vbArray) = vbArray Then
            ReDim iFormsHwndToKeepEnabled(UBound(nFormsHwndToKeepEnabled) + 1)
            For c2 = 0 To UBound(nFormsHwndToKeepEnabled)
                 iFormsHwndToKeepEnabled(c2 + 1) = Val(nFormsHwndToKeepEnabled(c2))
            Next c2
        Else
            ReDim iFormsHwndToKeepEnabled(1)
            iFormsHwndToKeepEnabled(1) = Val(nFormsHwndToKeepEnabled)
        End If
    End If
    
    If Not nNoOwner Then
        mOwnerFormHwnd = GetActiveFormHwnd ' mOwnerForm.hWnd
    End If
    If mOwnerFormHwnd <> 0 Then
        AttachMessage Me, mOwnerFormHwnd, WM_DESTROY
        AttachMessage Me, mOwnerFormHwnd, WM_NCACTIVATE
        
        If nSetIcon Then
            If MyGetProp(nForm.hWnd, "IconChanged") = 0 Then
                MySetProp nForm.hWnd, "IconChanged", 1
                SendMessageA nForm.hWnd, WM_SETICON, ICON_SMALL, SendMessageA(mOwnerFormHwnd, WM_GETICON, ICON_SMALL, 0&)
            End If
        End If
    End If
    
    EnsureFocusRect nForm ' there is a bug in VB (it has nothing to do with modal forms, but it is a good place to fix this bug)
    
    Set mForm = nForm
    mFormHwnd = mForm.hWnd
    AddModalForm mFormHwnd
    If mOwnerFormHwnd <> 0 Then
        AttachMessage Me, mOwnerFormHwnd, WM_ACTIVATEAPP
    Else
        AttachMessage Me, mFormHwnd, WM_ACTIVATEAPP
    End If
    AttachMessage Me, mFormHwnd, WM_NCACTIVATE
    mAttached = True
    
    mRgn = CreateRectRgn(0, 0, 0, 0)
    Call SetWindowRgn(mForm.hWnd, mRgn, True)
    SetProp mForm.hWnd, "NotVisible", 1
    SetProp mForm.hWnd, "ShownModal", 1
    Set mTimerShow = New cTimer
    mTimerShow.Interval = 10
    
    GetAppFormsHwnds mHwnds(), 0&
    ReDim mEnabled(UBound(mHwnds))
    ReDim mWindowPlacements(UBound(mHwnds))
    
    For c = 1 To UBound(mHwnds)
        mEnabled(c) = IsWindowEnabled(mHwnds(c))
        If mHwnds(c) <> nForm.hWnd Then
            iKeepEnabled = False
            For c2 = 1 To UBound(iFormsHwndToKeepEnabled)
                If mHwnds(c) = iFormsHwndToKeepEnabled(c2) Then
                    iKeepEnabled = True
                    Exit For
                End If
            Next c2
            If nKeepEnabledTaskBarWindows Then
                ' que no deshabilite las pantallas que se muestran en el taskbar y que no tienen owner (al menos esa es la idea)
                If Not iKeepEnabled Then
                    If mHwnds(c) <> mOwnerFormHwnd Then
                        iStyle = GetWindowLong(mHwnds(c), GWL_EXSTYLE)
                        If (iStyle And WS_EX_APPWINDOW) <> 0 Then
                            iKeepEnabled = True
                        End If
                    End If
                End If
            End If
            If Not iKeepEnabled Then
                EnableWindow mHwnds(c), 0&
            End If
        Else
            EnableWindow nForm.hWnd, 1&
        End If
    Next c
    
    If mOwnerFormHwnd <> 0 Then
        iStyle = GetWindowLong(mOwnerFormHwnd, GWL_EXSTYLE)
        If (iStyle And WS_EX_APPWINDOW) = 0 Then
            If GetProp(mOwnerFormHwnd, "ShownModal") = 0 Then
                iDoNotChangeStyle = True
            End If
        End If
    Else
        iDoNotChangeStyle = True
    End If
    If Not iDoNotChangeStyle Then
        iWindowStyle = GetWindowLong(nForm.hWnd, GWL_EXSTYLE)
        iWindowStyle = iWindowStyle And Not WS_EX_APPWINDOW
        SetWindowLong nForm.hWnd, GWL_EXSTYLE, iWindowStyle
    End If
    
    iFormHwnd = nForm.hWnd
    If mOwnerFormHwnd = 0 Then
        On Error Resume Next
        Err.Clear
        nForm.Show
        If Err.Number = 401 Then
            EnableForms
            On Error GoTo 0
            nForm.Show 1
            Exit Sub
        Else
            nForm.Visible = True
            nForm.ZOrder
            nForm.SetFocus
        End If
        On Error GoTo 0
    Else
        On Error Resume Next
        Err.Clear
        nForm.Show
        If Err.Number = 401 Then
            EnableForms
            On Error GoTo 0
            nForm.Show 1
            Exit Sub
        End If
        nForm.SetFocus
'        nForm.Visible = True
'        mForm.ZOrder
        On Error GoTo 0
        mOldOwner = SetOwner(nForm.hWnd, mOwnerFormHwnd)
        nForm.ZOrder
    End If
    
    If nWaitWithDoevents Then
        Do While IsWindow(iFormHwnd) > 0
            DoEvents
            Sleep 1
            If (IsWindowVisible(iFormHwnd) = 0) And Not (IsIconic(iFormHwnd) <> 0) Then
                
                ' close any non modal form opened after this one
                GetAppFormsHwnds iAuxFormsHwnds, 0&
                For c = 1 To UBound(iAuxFormsHwnds)
                    If Not IsInVector(mHwnds, iAuxFormsHwnds(c)) Then
                        If IsWindow(iAuxFormsHwnds(c)) <> 0 Then
                            iStyle = GetWindowLong(iAuxFormsHwnds(c), GWL_EXSTYLE)
                            If (iStyle And WS_EX_APPWINDOW) = 0 Then
                                Call SendMessage(iAuxFormsHwnds(c), WM_SYSCOMMAND, SC_CLOSE, 0&)
                            End If
                        End If
                    End If
                Next c
                
                EnableForms
                Exit Do
            End If
        Loop
    Else
        Set mMe = Me
    End If
    RemoveModalForm mFormHwnd
    
    RestoreOnExit
    Set mForm = Nothing
End Sub

Private Sub Class_Terminate()
    Finish
End Sub

Private Sub Finish()
    If mAttached <> 0 Then
        If mOwnerFormHwnd <> 0 Then
            If IsWindow(mOwnerFormHwnd) <> 0 Then
                DetachMessage Me, mOwnerFormHwnd, WM_DESTROY
                DetachMessage Me, mOwnerFormHwnd, WM_NCACTIVATE
                DetachMessage Me, mOwnerFormHwnd, WM_ACTIVATEAPP
            End If
        End If
        If IsWindow(mFormHwnd) <> 0 Then
            DetachMessage Me, mFormHwnd, WM_NCACTIVATE
            If mOwnerFormHwnd = 0 Then
                DetachMessage Me, mFormHwnd, WM_ACTIVATEAPP
            End If
        End If
        mAttached = False
    End If
    If Not mTimerShow Is Nothing Then
        mTimerShow.Interval = 0
        Set mTimerShow = Nothing
    End If
    If Not mSetActiveWindow Is Nothing Then
        mSetActiveWindow.Interval = 0
        Set mSetActiveWindow = Nothing
    End If
    Set mForm = Nothing
End Sub

Private Sub mForm_Resize()
'    If mRgn <> 0 Then
'        Set mTimerShow = New cTimer
'        mTimerShow.Interval = 10
'    End If
    If mForm.WindowState = vbMinimized Then
        If mOwnerFormHwnd <> 0 Then
            Set mMinimizeApp = MinimizeApp
        End If
    Else
        mFormWindowState = mForm.WindowState
    End If
End Sub

Private Sub mForm_Unload(Cancel As Integer)
    If Cancel = 0 Then
        If MyGetProp(mFormHwnd, "IconChanged") <> 0 Then
            If mForm.Visible Then
                mForm.Hide
            End If
            'SendMessageA mFormHwnd, WM_SETICON, ICON_SMALL, 0&
            MyRemoveProp mFormHwnd, "IconChanged"
        End If
        Finish
    End If
End Sub

Private Sub RestoreOnExit()
    If mFormHwnd = 0 Then Exit Sub
    
    If GetProp(mFormHwnd, "ShownModal") <> 0 Then
        RemoveProp mFormHwnd, "ShownModal"
    End If
    EnableForms
    RestoreOldOwner
    If MyGetProp(mFormHwnd, "IconChanged") <> 0 Then
      '  If IsWindow(mFormHwnd) <> 0 Then
      '      SendMessageA mFormHwnd, WM_SETICON, ICON_SMALL, 0&
      '  End If
        MyRemoveProp mFormHwnd, "IconChanged"
    End If
    If mOwnerFormHwnd <> 0 Then
        Finish
        SetFocusAPI mOwnerFormHwnd
        mOwnerFormHwnd = 0
    End If
    mFormHwnd = 0
End Sub

Private Sub EnableForms()
    Dim c As Long
    
    For c = 0 To UBound(mHwnds)
        If IsWindow(mHwnds(c)) > 0 Then
            If IsWindow(mHwnds(c)) <> 0 Then
                EnableWindow mHwnds(c), mEnabled(c)
            End If
            mHwnds(c) = 0
        End If
    Next c
    Set mMe = Nothing
End Sub

Private Sub mSetActiveWindow_ThatTime()
    Dim iAf As Long
    
    mSetActiveWindow.Interval = 0
    Set mSetActiveWindow = Nothing
    
    If (IsWindowVisible(mFormToSetActiveHwnd) <> 0) And (IsWindowEnabled(mFormToSetActiveHwnd) <> 0) Then
        If (TopModalFormHwnd = mFormToSetActiveHwnd) Then
            iAf = GetActiveFormHwnd(True)
            If iAf = GetTopOwnerFormHwnd(mFormToSetActiveHwnd) Then
                SetForegroundWindow mFormToSetActiveHwnd
                SetActiveWindow mFormToSetActiveHwnd
            End If
        End If
    End If
End Sub

Private Sub mTimerShow_ThatTime()
    If mRgn <> 0 Then
        Call SetWindowRgn(mForm.hWnd, 0, True)
        DeleteObject mRgn
        mRgn = 0
        RemoveProp mForm.hWnd, "NotVisible"
    End If
    mTimerShow.Interval = 0
    Set mTimerShow = Nothing
End Sub

Private Function SetOwner(ByVal HwndWindow As Long, ByVal hwndOwner As Long) As Long
    SetOwner = SetWindowLong(HwndWindow, GWL_HWNDPARENT, hwndOwner)
End Function

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrConsume
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Select Case iMsg
        Case WM_NCACTIVATE
            If hWnd = mOwnerFormHwnd Then
                If wParam = 1 Then
                    If IsWindowVisible(mFormHwnd) Then
'                        wParam = 0
                        ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
                        
                       ' SetForegroundWindow mFormHwnd
                        Exit Function
                    End If
                End If
            Else
'                Debug.Print "wParam: " & CStr(wParam), , "IsIconic: " & CStr(CBool(IsIconic(mFormHwnd) <> 0)), , "IsWindowVisible: " & CStr(CBool(IsWindowVisible(mFormHwnd) <> 0))
                If wParam = 1 Then
                    If IsIconic(mFormHwnd) <> 0 Then
                        If IsWindowVisible(mFormHwnd) <> 0 Then
                            ShowWindow mFormHwnd, SW_RESTORE
                        End If
                    End If
                    If TopModalFormHwnd <> mFormHwnd Then
                        If GetTopOwnerFormHwnd(TopModalFormHwnd) = mFormHwnd Then
                            mFormToSetActiveHwnd = TopModalFormHwnd
                            If Not mSetActiveWindow Is Nothing Then
                                mSetActiveWindow.Interval = 0
                                Set mSetActiveWindow = Nothing
                            End If
                            Set mSetActiveWindow = New cTimer
                            mSetActiveWindow.Interval = 300
                        End If
                    End If
                End If
            End If
        Case WM_ACTIVATEAPP
            If Not mSetActiveWindow Is Nothing Then
                mSetActiveWindow.Interval = 0
                Set mSetActiveWindow = Nothing
            End If
            'WriteTraceLog wParam & "    " & lParam & "    " & App.ThreadID
            If wParam = 1 Then
                ' there is a VB6 bug that happens compiled that changes the active window when the program is restored or activated from the taskbar
                ' it happens when there are more than one owned window. It changes the ZOrder of the owned windows.
                ' Then the modal window can be left below other windows also owned by the same owner
                ' this is to fix it.
                If mFormWasTheActiveWindow Then
                    Set mSetActiveWindow = New cTimer
                    mSetActiveWindow.Interval = 300
                    mFormToSetActiveHwnd = mFormHwnd
                End If
            Else
                mFormWasTheActiveWindow = (GetActiveFormHwnd = mFormHwnd)
                'Debug.Print "mFormHwnd: " & mFormHwnd, "Owner: " & GetOwnerHwnd(mFormHwnd), "TopModalFormHwnd: " & TopModalFormHwnd, "GetActiveWindow: " & GetActiveWindow, "GetActiveFormHwnd: " & GetActiveFormHwnd
            End If
        Case WM_DESTROY
            RestoreOldOwner
            Finish
    End Select
    
    ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
End Function

Private Sub RestoreOldOwner()
    If mOldOwner <> 0 Then
        If (IsWindow(mFormHwnd) <> 0) And (IsWindow(mOldOwner) <> 0) Then
            SetOwner mFormHwnd, mOldOwner
        End If
        mOldOwner = 0
'        Set mOwnerForm = Nothing
    End If
End Sub

'Private Function IsWindowTaskbarVisible(nHwnd As Long) As Boolean
'    IsWindowTaskbarVisible = ((GetWindowLong(nHwnd, GWL_EXSTYLE) And WS_EX_APPWINDOW) = WS_EX_APPWINDOW)
'End Function
