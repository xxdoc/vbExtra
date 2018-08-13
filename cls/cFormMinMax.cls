VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormMinMax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Public Event BeforeMove(ByRef NewLeft As Single, ByRef NewTop As Single)
Public Event Move(Initialized As Boolean)
Public Event Size(Initialized As Boolean)
Public Event Unload()
Public Event MovedByUser()

Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
          
Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
      
Private Type WINDOWPOS
   hWnd As Long
   hWndInsertAfter As Long
   x As Long
   y As Long
   cx As Long
   cy As Long
   Flags As Long
End Type
      
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private mMinWidth As Long
Private mMinHeight As Long
Private mMaxWidth As Long
Private mMaxHeight As Long
Private mFormHwnd As Long
Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private mCollection As cObjectHandlersCollection
Private mInitialized As Boolean

Public Sub SetMinMax(nForm As Object, Optional nMinWidth, Optional nMinHeight, Optional nMaxWidth, Optional nMaxHeight, Optional ScaleMode As Integer = vbTwips, Optional nParent As Object)
    Dim iLng As Long
    Dim iFormWidth As Long
    Dim iFormHeight As Long
    
    If Not nParent Is Nothing Then
        Set mCollection = nParent
    End If
    
    If nForm.WindowState = vbNormal Then
        If Not IsMissing(nMinWidth) Then
            If nMinWidth <> 0 Then
                mMinWidth = nForm.ScaleX(nMinWidth, ScaleMode, vbPixels)
                iLng = GetSystemMetrics(SM_CXMINTRACK)
                If mMinWidth < iLng Then
                    mMinWidth = iLng
                End If
                iFormWidth = nForm.ScaleX(nForm.Width, vbTwips, vbPixels)
                If iFormWidth < mMinWidth Then
                    nForm.Width = nForm.ScaleX(mMinWidth, vbPixels, vbTwips)
                End If
            End If
        End If
        If Not IsMissing(nMinHeight) Then
            If nMinHeight <> 0 Then
                mMinHeight = nForm.ScaleX(nMinHeight, ScaleMode, vbPixels)
                iLng = GetSystemMetrics(SM_CYMINTRACK)
                If mMinHeight < iLng Then
                    mMinHeight = iLng
                End If
                iFormHeight = nForm.ScaleX(nForm.Height, vbTwips, vbPixels)
                If iFormHeight < mMinHeight Then
                    nForm.Height = nForm.ScaleX(mMinHeight, vbPixels, vbTwips)
                End If
            End If
        End If
        If Not IsMissing(nMaxWidth) Then
            If nMaxWidth <> 0 Then
                mMaxWidth = nForm.ScaleX(nMaxWidth, ScaleMode, vbPixels)
            End If
            If iFormWidth = 0 Then
                iFormWidth = nForm.ScaleX(nForm.Width, vbTwips, vbPixels)
            End If
            If iFormWidth > mMaxWidth Then
                nForm.Width = nForm.ScaleX(mMaxWidth, vbPixels, vbTwips)
            End If
        End If
        If Not IsMissing(nMaxHeight) Then
            If nMaxHeight <> 0 Then
                mMaxHeight = nForm.ScaleX(nMaxHeight, ScaleMode, vbPixels)
            End If
            If iFormHeight = 0 Then
                iFormHeight = nForm.ScaleX(nForm.Height, vbTwips, vbPixels)
            End If
            If iFormHeight > mMaxHeight Then
                nForm.Height = nForm.ScaleX(mMaxHeight, vbPixels, vbTwips)
            End If
        End If
    End If
    
    If mFormHwnd = 0 Then
        Set mForm = nForm
        mFormHwnd = nForm.hWnd
        AttachMessage Me, mFormHwnd, WM_GETMINMAXINFO
        AttachMessage Me, mFormHwnd, WM_MOVE
        AttachMessage Me, mFormHwnd, WM_SIZE
        AttachMessage Me, mFormHwnd, WM_DESTROY
        AttachMessage Me, mFormHwnd, WM_WINDOWPOSCHANGING
        AttachMessage Me, mFormHwnd, WM_PAINT
        AttachMessage Me, mFormHwnd, WM_MOVING
    End If
End Sub

Private Sub Class_Terminate()
    Detach
End Sub

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    Select Case iMsg
        Case WM_GETMINMAXINFO, WM_WINDOWPOSCHANGING
            ISubclass_MsgResponse = emrConsume
        Case WM_MOVE, WM_SIZE, WM_MOVING
            ISubclass_MsgResponse = emrPostProcess
        Case WM_DESTROY, WM_PAINT
            ISubclass_MsgResponse = emrPreprocess
    End Select
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Dim iMinMax As MINMAXINFO
    Dim iWindowPos As WINDOWPOS
    Dim iRect As RECT
    Dim iNewLeft As Single
    Dim iNewTop As Single
    Dim iNewLeft_Prev As Single
    Dim iNewTop_Prev As Single
    Dim iChanged As Boolean
    
    Select Case iMsg
        Case WM_GETMINMAXINFO
            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)

            CopyMemory iMinMax, ByVal lParam, Len(iMinMax)
            
            If mMinWidth <> 0 Then
                iMinMax.ptMinTrackSize.x = mMinWidth
            End If
            If mMinHeight <> 0 Then
                iMinMax.ptMinTrackSize.y = mMinHeight
            End If
            If mMaxWidth <> 0 Then
                iMinMax.ptMaxTrackSize.x = mMaxWidth
            End If
            If mMaxHeight <> 0 Then
                iMinMax.ptMaxTrackSize.y = mMaxHeight
            End If
            CopyMemory ByVal lParam, iMinMax, Len(iMinMax)
            
            ISubclass_WindowProc = DefWindowProc(hWnd, iMsg, wParam, lParam)
        Case WM_MOVE
'            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            RaiseEvent Move(mInitialized)
        Case WM_SIZE
'            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            RaiseEvent Size(mInitialized)
        Case WM_DESTROY
            Finish
'            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
        Case WM_WINDOWPOSCHANGING
            CopyMemory iWindowPos, ByVal lParam, Len(iWindowPos)
            GetWindowRect hWnd, iRect
            If (iWindowPos.x <> iRect.Left) Or (iWindowPos.y <> iRect.Top) Then
                iNewLeft = iWindowPos.x * Screen.TwipsPerPixelX
                iNewTop = iWindowPos.y * Screen.TwipsPerPixelY
                iNewLeft_Prev = iNewLeft
                iNewTop_Prev = iNewTop
                RaiseEvent BeforeMove(iNewLeft, iNewTop)
                If iNewLeft_Prev <> iNewLeft Then
                    iWindowPos.x = iNewLeft / Screen.TwipsPerPixelX
                    iChanged = True
                End If
                If iNewTop_Prev <> iNewTop Then
                    iWindowPos.y = iNewTop / Screen.TwipsPerPixelX
                    iChanged = True
                End If
                If iChanged Then
                    CopyMemory ByVal lParam, iWindowPos, Len(iWindowPos)
                End If
            End If
            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
        Case WM_PAINT
            mInitialized = True
'            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
        Case WM_MOVING
'            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
            RaiseEvent MovedByUser
    End Select
End Function

Private Sub Detach()
    If mFormHwnd <> 0 Then
        DetachMessage Me, mFormHwnd, WM_GETMINMAXINFO
        DetachMessage Me, mFormHwnd, WM_MOVE
        DetachMessage Me, mFormHwnd, WM_SIZE
        DetachMessage Me, mFormHwnd, WM_DESTROY
        DetachMessage Me, mFormHwnd, WM_WINDOWPOSCHANGING
        DetachMessage Me, mFormHwnd, WM_PAINT
        DetachMessage Me, mFormHwnd, WM_MOVING
        mFormHwnd = 0
    End If
End Sub

Private Sub mForm_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Finish
    End If
    RaiseEvent Unload
End Sub

Private Sub Finish()
    If Not mCollection Is Nothing Then
        mCollection.Remove mFormHwnd
        Set mCollection = Nothing
    End If
    Detach
'    If Not mCollection Is Nothing Then
'        mCollection.Remove mFormHwnd
'        Set mCollection = Nothing
'    End If
    Set mForm = Nothing
End Sub
