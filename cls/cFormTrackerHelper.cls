VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFormTrackerHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISubclass

Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private mParent As cFormsTracker
Private mFormHwnd As Long
Private WithEvents mTimerSetStartMonitor As cTimer
Attribute mTimerSetStartMonitor.VB_VarHelpID = -1
Private mStartMonitor As Long

Public Sub SetForm(nForm As Object, nParent As cFormsTracker)
    If nForm Is Nothing Then Exit Sub
    Set mForm = nForm
    Set mParent = nParent
    Set mTimerSetStartMonitor = New cTimer
    mTimerSetStartMonitor.Interval = 100
    mFormHwnd = mForm.hWnd
    AttachMessage Me, mFormHwnd, WM_MOVE
    AttachMessage Me, mFormHwnd, WM_DESTROY
    AttachMessage Me, mFormHwnd, WM_NCACTIVATE
End Sub

Private Sub Class_Terminate()
    If Not mTimerSetStartMonitor Is Nothing Then
        mTimerSetStartMonitor.Interval = 0
        Set mTimerSetStartMonitor = Nothing
    End If
End Sub

Private Sub mForm_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Finish
    End If
End Sub

Private Sub Finish()
    Dim iMonitor As Long
    
    If mFormHwnd <> 0 Then
        If Not mTimerSetStartMonitor Is Nothing Then
            mTimerSetStartMonitor.Interval = 0
            Set mTimerSetStartMonitor = Nothing
        End If
        If mStartMonitor <> 0 Then
            iMonitor = MonitorFromWindow(mFormHwnd, MONITOR_DEFAULTTONULL)
            If iMonitor <> mStartMonitor Then
                mParent.ChangedMonitor mForm, mStartMonitor, iMonitor
            End If
        End If
        
        If Not mParent Is Nothing Then
            mParent.RemoveForm mFormHwnd
            Set mParent = Nothing
        End If
        DetachMessage Me, mFormHwnd, WM_MOVE
        DetachMessage Me, mFormHwnd, WM_DESTROY
        DetachMessage Me, mFormHwnd, WM_NCACTIVATE
        mFormHwnd = 0
    End If
End Sub

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    If iMsg = WM_DESTROY Then
        ISubclass_MsgResponse = emrPreprocess
    Else
        ISubclass_MsgResponse = emrPostProcess
    End If
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, wParam As Long, lParam As Long, bConsume As Boolean) As Long
    If iMsg = WM_MOVE Then
        mParent.FormAction mFormHwnd
    ElseIf iMsg = WM_NCACTIVATE Then
        If wParam = 1 Then ' activate
            mParent.FormAction mFormHwnd
        End If
    ElseIf iMsg = WM_DESTROY Then
        Finish
    End If
End Function

Private Sub mTimerSetStartMonitor_ThatTime()
    mTimerSetStartMonitor.Interval = 0
    Set mTimerSetStartMonitor = Nothing
    
    mStartMonitor = MonitorFromWindow(mFormHwnd, MONITOR_DEFAULTTONULL)
End Sub
