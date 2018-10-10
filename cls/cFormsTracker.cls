VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFormsTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mTrackedFormsCollection As New cObjectHandlersCollection
Private WithEvents mTimer As cTimer
Attribute mTimer.VB_VarHelpID = -1
Private mLastFormHwnd As Long
Private mCurrentMonitor As Long

Public Sub AddForm(nForm As Object)
    Dim iFormTrackerHelper As cFormTrackerHelper
    
    Set iFormTrackerHelper = mTrackedFormsCollection.GetInstance(nForm)
    If iFormTrackerHelper Is Nothing Then
        Set iFormTrackerHelper = New cFormTrackerHelper
        iFormTrackerHelper.SetForm nForm, Me
        If mTrackedFormsCollection.GetInstance(nForm) Is Nothing Then
            mTrackedFormsCollection.Add iFormTrackerHelper, nForm.hWnd
        End If
    End If
    SetCurrentMonitorFromWindow nForm.hWnd
End Sub

Public Sub RemoveForm(nFormHwnd As Long)
       On Error Resume Next
       mTrackedFormsCollection.Remove nFormHwnd
End Sub

Private Sub Class_Terminate()
    Set mTrackedFormsCollection = Nothing
End Sub

Public Sub FormAction(nHwndForm As Long)
    If mTimer Is Nothing Then
        Set mTimer = New cTimer
        mTimer.Interval = 1000
        mLastFormHwnd = nHwndForm
    End If
End Sub

Private Sub mTimer_ThatTime()
    mTimer.Interval = 0
    Set mTimer = Nothing
    SetCurrentMonitorFromWindow mLastFormHwnd
End Sub

Private Sub SetCurrentMonitorFromWindow(nWindowHwnd As Long)
    Dim iLng As Long
    
    iLng = MonitorFromWindow(nWindowHwnd, MONITOR_DEFAULTTONULL)
    If iLng <> 0 Then
        mCurrentMonitor = iLng
    End If
End Sub

Private Sub SetCurrentMonitorFromMouse()
    Dim iCP As POINTAPI
    
    GetCursorPos iCP
    mCurrentMonitor = MonitorFromPoint(iCP.x, iCP.y, MONITOR_DEFAULTTOPRIMARY)
End Sub

Public Property Get CurrentMonitor() As Long
    Dim iActiveFormHwnd As Long
    
    Update
    iActiveFormHwnd = GetActiveFormHwnd
    If iActiveFormHwnd <> 0 Then
        SetCurrentMonitorFromWindow iActiveFormHwnd
    End If
    CurrentMonitor = mCurrentMonitor
End Property

Public Sub Update()
    If mCurrentMonitor = 0 Then
        SetCurrentMonitorFromMouse
    End If
End Sub

Public Sub ChangedMonitor(nForm As Form, nStartMonitor As Long, nEndMonitor As Long)
'    MsgBox nForm.Name & " changed from  " & nStartMonitor & " to " & nEndMonitor
    SaveSetting AppNameForRegistry, "MInfo", Base64Encode(nForm.Name) & ".MI", "-1"
End Sub
