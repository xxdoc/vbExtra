VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimerEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Timer()

Private WithEvents mTimer As cTimer
Attribute mTimer.VB_VarHelpID = -1

Private mEnabled As Boolean
Private mInterval As Long
Private mTag As Variant
Private mSetForOnce As Boolean

Public Property Get Interval() As Long
    Interval = mInterval
End Property

Public Property Let Interval(nValue As Long)
    If nValue <> mInterval Then
        mInterval = nValue
        SetTimer
    End If
End Property

Public Property Get Tag() As Variant
    Tag = mTag
End Property

Public Property Let Tag(nValue As Variant)
    mTag = nValue
End Property

Public Sub Reset()
    If mEnabled And (mInterval > 0) Then
        mTimer.Interval = 0
        mTimer.Interval = mInterval
    End If
End Sub

Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(nValue As Boolean)
    If nValue <> mEnabled Then
        mEnabled = nValue
        SetTimer
    End If
End Property

Private Sub SetTimer(Optional ByVal nInterval As Long)
    If nInterval <= 0 Then
        nInterval = mInterval
    End If
    If (mEnabled Or mSetForOnce) And (nInterval > 0) Then
        If mTimer Is Nothing Then
            Set mTimer = New cTimer
        End If
        If mTimer.Interval <> nInterval Then
            mTimer.Interval = nInterval
        End If
    Else
        If Not mTimer Is Nothing Then
            mTimer.Interval = 0
            Set mTimer = Nothing
        End If
    End If
End Sub

Private Sub Class_Terminate()
    If Not mTimer Is Nothing Then
        mTimer.Interval = 0
        Set mTimer = Nothing
    End If
End Sub

Private Sub mTimer_ThatTime()
    If mSetForOnce Then
        mSetForOnce = False
        mTimer.Interval = 0
        Set mTimer = Nothing
    End If
    RaiseEvent Timer
End Sub

Public Sub SetForOnce(Optional ByVal nInterval As Long)
    
    If nInterval <= 0 Then
        nInterval = mInterval
    End If
    If (nInterval > 0) Then
        mSetForOnce = True
        mEnabled = False
        If Not mTimer Is Nothing Then
            mTimer.Interval = 0
            Set mTimer = Nothing
        End If
        SetTimer nInterval
    End If
End Sub
