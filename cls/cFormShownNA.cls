VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFormShownNA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISubclass

Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
      
Private Const ICON_SMALL As Long = 0&
'Private Const ICON_BIG As Long = 1&
Private Const WM_SETICON As Long = &H80&
Private Const WM_GETICON As Long = &H7F&

Private Const GWL_HWNDPARENT As Long = (-8)

Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private WithEvents mOwnerForm As Form
Attribute mOwnerForm.VB_VarHelpID = -1
Private mCollection As cObjectHandlersCollection
Private mFormHwnd As Long
Private mOldOwner As Long
Private mOwnerFormHwnd As Long
Private mAttached As Boolean

Public Sub ShowForm(nForm As Object, Optional nOwnerForm, Optional nParent As cObjectHandlersCollection, Optional nSetIcon As Boolean = True, Optional nSetActiveFormAsOwner As Boolean)
    Dim iAuxAttached As Boolean
    
    If nForm Is Nothing Then Exit Sub
    
    Set mForm = nForm
    If Not nParent Is Nothing Then
        Set mCollection = nParent
    End If
    mFormHwnd = mForm.hWnd
    
    iAuxAttached = mAttached
    If Not iAuxAttached Then
        AttachMessage Me, mFormHwnd, WM_DESTROY
        mAttached = True
        iAuxAttached = True
    End If
    ShowWindow mForm.hWnd, SW_SHOWNA
    If Not IsMissing(nOwnerForm) Then
        If Not nOwnerForm Is Nothing Then
            mOwnerFormHwnd = nOwnerForm.hWnd
        Else
            If nSetActiveFormAsOwner Then
                mOwnerFormHwnd = GetActiveFormHwnd
            End If
        End If
    Else
        If nSetActiveFormAsOwner Then
            mOwnerFormHwnd = GetActiveFormHwnd
        End If
    End If
    If mOwnerFormHwnd <> 0 Then
        If GetWindowLong(mForm.hWnd, GWL_HWNDPARENT) <> mOwnerFormHwnd Then
            If Not IsMissing(nOwnerForm) Then
                Set mOwnerForm = nOwnerForm
            End If
            mOldOwner = SetOwner(mForm.hWnd, mOwnerFormHwnd)
        End If
        If Not iAuxAttached Then AttachMessage Me, mOwnerFormHwnd, WM_DESTROY
        If nSetIcon Then
            If MyGetProp(nForm.hWnd, "IconChanged") = 0 Then
                MySetProp nForm.hWnd, "IconChanged", 1
                SendMessageA nForm.hWnd, WM_SETICON, ICON_SMALL, SendMessageA(mOwnerFormHwnd, WM_GETICON, ICON_SMALL, 0&)
            End If
        End If
    End If
End Sub

Private Sub mForm_Activate()
    If IsIconic(mOwnerFormHwnd) <> 0 Then
        ShowWindow mOwnerFormHwnd, SW_RESTORE
    End If
End Sub

Private Sub mForm_Resize()
    If mForm.WindowState <> vbMinimized Then
        If IsIconic(mOwnerFormHwnd) <> 0 Then
            ShowWindow mOwnerFormHwnd, SW_SHOWNA
        End If
    End If
End Sub

Private Sub mForm_Unload(Cancel As Integer)
    Finish
End Sub

Private Sub Finish()
    Detach
    If MyGetProp(mFormHwnd, "IconChanged") <> 0 Then
        'SendMessageA mFormHwnd, WM_SETICON, ICON_SMALL, 0&
        MyRemoveProp mFormHwnd, "IconChanged"
    End If
    RestoreOldOwner
    If Not mCollection Is Nothing Then
        mCollection.Remove mFormHwnd
        Set mCollection = Nothing
    End If
    Set mForm = Nothing
    Set mOwnerForm = Nothing
End Sub

Private Sub Detach()
    If mAttached Then
        mAttached = False
        DetachMessage Me, mFormHwnd, WM_DESTROY
        If mOwnerFormHwnd <> 0 Then
            DetachMessage Me, mOwnerFormHwnd, WM_DESTROY
        End If
    End If
End Sub

Private Function SetOwner(ByVal HwndWindow, ByVal HwndofOwner) As Long
    On Error Resume Next
    SetOwner = SetWindowLong(HwndWindow, GWL_HWNDPARENT, HwndofOwner)
End Function

Private Sub RestoreOldOwner()
    If mOldOwner <> 0 Then
        SetOwner mForm.hWnd, mOldOwner
        mOldOwner = 0
        Set mOwnerForm = Nothing
    End If
End Sub

Private Sub mOwnerForm_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Unload mForm
    End If
End Sub

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrPreprocess
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    
    Select Case iMsg
        Case WM_DESTROY
            On Error Resume Next
            If IsWindow(mFormHwnd) <> 0 Then
                Unload mForm
            End If
            Finish
    End Select
'    ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
End Function

