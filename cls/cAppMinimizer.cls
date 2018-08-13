VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppMinimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Private Type WINDOWPLACEMENT
        Length As Long
        Flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type

Private Declare Function SetActiveWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Const SW_MINIMIZE As Long = 6&
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SetForegroundWindow& Lib "user32" (ByVal hWnd&)

Public Event Restored()

Private mHwnds() As Long
Private mWindowPlacements() As WINDOWPLACEMENT
Private mTopFormHwnd As Long
Private mActiveWindowHwnd As Long
Private mMe As Object

Private Sub Class_Initialize()
    MinimizeForms
    Set mMe = Me
End Sub

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrConsume
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    
    Select Case iMsg
        Case WM_NCACTIVATE
            'GetWindowPlacement hWnd, iwp
            'CopyMemory ByVal lParam, iwp, Len(iwp)
            If IsWindowEnabled(hWnd) = 0 Then
                wParam = 0
            End If
            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            RestoreForms
        Case WM_DESTROY
            DetachMessage Me, hWnd, WM_NCACTIVATE
            DetachMessage Me, hWnd, WM_DESTROY
            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
    End Select
End Function


Private Sub MinimizeForms()
    Dim iWP As WINDOWPLACEMENT
    Dim c As Long
    
    GetAppFormsHwnds mHwnds(), mTopFormHwnd, True
    mActiveWindowHwnd = GetActiveWindow
    If mActiveWindowHwnd <> 0 Then
        If Not IsWindowLocal(mActiveWindowHwnd) Then
            mActiveWindowHwnd = 0
        End If
    End If
    ReDim mWindowPlacements(UBound(mHwnds))
    
    For c = 1 To UBound(mHwnds)
        GetWindowPlacement mHwnds(c), iWP
        mWindowPlacements(c) = iWP
    Next c
    
    For c = 1 To UBound(mHwnds)
        iWP = mWindowPlacements(c)
        iWP.showCmd = SW_MINIMIZE
        SetWindowPlacement mHwnds(c), iWP
        AttachMessage Me, mHwnds(c), WM_NCACTIVATE
        AttachMessage Me, mHwnds(c), WM_DESTROY
    Next c
End Sub

Private Sub RestoreForms()
    Dim iWP As WINDOWPLACEMENT
    Dim c As Long
    
    For c = 1 To UBound(mHwnds)
        If IsWindow(mHwnds(c)) <> 0 Then
            iWP = mWindowPlacements(c)
            SetWindowPlacement mHwnds(c), iWP
            DetachMessage Me, mHwnds(c), WM_NCACTIVATE
            DetachMessage Me, mHwnds(c), WM_DESTROY
        End If
    Next c
    
    If mTopFormHwnd <> 0 Then
        ShowWindow mTopFormHwnd, SW_RESTORE
        SetForegroundWindow mTopFormHwnd
    End If
    
    If mActiveWindowHwnd <> 0 Then
        SetActiveWindow mActiveWindowHwnd
    End If

    RaiseEvent Restored
    Set mMe = Nothing
End Sub

