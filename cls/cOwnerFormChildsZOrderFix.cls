VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOwnerFormChildsZOrderFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const SWP_ASYNCWINDOWPOS As Long = &H4000
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOACTIVATE = &H10&
Private Const SWP_NOMOVE As Long = &H2
Private Const HWND_TOP As Long = 0
Private Const SWP_NOOWNERZORDER As Long = 512
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_SHOWWINDOW As Long = &H40&
Private Const hWnd_BOTTOM As Long = 1
Private Const SWP_NOREDRAW As Long = 8
Private Const SWP_DEFERERASE As Long = &H2000&
Private Const SWP_NOSENDCHANGING As Long = &H400&
Private Const SWP_NOZORDER As Long = &H4&

Private Const WM_WINDOWPOSCHANGING As Long = &H46&


Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private mFormHwnd As Long
Private mForms As Object
Private WithEvents mTimerFixZOrder As cTimer
Attribute mTimerFixZOrder.VB_VarHelpID = -1

Private mInfocus As Boolean
Private mLastTopZOrderOwnedFormHwnd As Long
Private mFixFocus As Boolean
Private mOwnedFormActiveHwnd As Long
Private mOwnedFormActiveHwnd_Previous As Long
Private mAttached As Boolean

Public Sub SetForm(nForm As Object, nForms As Object)
    If Not TypeOf nForm Is Form Then Exit Sub
    
    Set mForm = nForm
    mFormHwnd = mForm.hWnd
    Set mForms = nForms
    AttachMessage Me, mFormHwnd, WM_WINDOWPOSCHANGING
    AttachMessage Me, mFormHwnd, WM_DESTROY
    mAttached = True
End Sub

Private Sub mForm_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Detach
    End If
End Sub

Private Sub Detach()
    If mAttached Then
        DetachMessage Me, mFormHwnd, WM_WINDOWPOSCHANGING
        DetachMessage Me, mFormHwnd, WM_DESTROY
        If Not mTimerFixZOrder Is Nothing Then
            mTimerFixZOrder.Interval = 0
            Set mTimerFixZOrder = Nothing
        End If
        Set mForm = Nothing
        Set mForms = Nothing
        mAttached = False
    End If
End Sub

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrPreprocess
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Dim iFormOwnerActive As Object
    Dim iTopZOrderOwnedFormHwnd As Long
    Dim iFocusNow As Boolean
    
    If iMsg = WM_WINDOWPOSCHANGING Then
        If IsWindow(mFormHwnd) = 0 Then
            Detach
            Exit Function
        End If
        
        Set iFormOwnerActive = GetOwnerForm2(ScreenActiveForm(mForms), mForms)
        If iFormOwnerActive Is mForm Then
            If mInfocus Then
                iTopZOrderOwnedFormHwnd = GetTopZOrderFormHwnd(mForms, mFormHwnd)
                mLastTopZOrderOwnedFormHwnd = iTopZOrderOwnedFormHwnd
                If mInfocus Then
                    If Not mFixFocus Then
                        mOwnedFormActiveHwnd_Previous = mOwnedFormActiveHwnd
                    End If
                End If
                mOwnedFormActiveHwnd = ScreenActiveForm(mForms).hWnd
            End If
            iFocusNow = True
        Else
            iFocusNow = False
        End If
        If iFocusNow <> mInfocus Then
            If iFocusNow Then
                ' Main (owner) form receive focus
                If mFixFocus Then
                    Set mTimerFixZOrder = New cTimer
                    mTimerFixZOrder.Interval = 10
                    mTimerFixZOrder.Tag = mLastTopZOrderOwnedFormHwnd
                    ' Need to set Zorder to mLastTopZOrderOwnedFormHwnd
                End If
            Else
                ' Main form loses focus
                iTopZOrderOwnedFormHwnd = GetTopZOrderFormHwnd(mForms, mFormHwnd)
                If iTopZOrderOwnedFormHwnd <> 0 Then
                    If iTopZOrderOwnedFormHwnd <> mLastTopZOrderOwnedFormHwnd Then
                        ' Need to fix Zorder of mLastTopZOrderOwnedFormHwnd
                        mFixFocus = True
                    End If
                End If
            End If
        End If
        mInfocus = iFocusNow
    ElseIf iMsg = WM_DESTROY Then
        Detach
    End If
    
End Function

Private Sub mTimerFixZOrder_ThatTime()
    Dim iHwnd As Long
    
    iHwnd = Val(mTimerFixZOrder.Tag)
    If IsWindow(iHwnd) <> 0 Then
        SetWindowPos iHwnd, HWND_TOP, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_ASYNCWINDOWPOS     'Or SWP_NOREDRAW Or SWP_ASYNCWINDOWPOS Or SWP_DEFERERASE Or SWP_NOSENDCHANGING    'Or SWP_SHOWWINDOW
        If mOwnedFormActiveHwnd_Previous <> 0 Then
            If IsWindow(mOwnedFormActiveHwnd_Previous) <> 0 Then
                SetWindowPos mOwnedFormActiveHwnd_Previous, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_ASYNCWINDOWPOS      'Or SWP_NOREDRAW Or SWP_ASYNCWINDOWPOS Or SWP_DEFERERASE Or SWP_NOSENDCHANGING    'Or SWP_SHOWWINDOW
            End If
        End If
    End If
    mTimerFixZOrder.Interval = 0
    Set mTimerFixZOrder = Nothing
End Sub



