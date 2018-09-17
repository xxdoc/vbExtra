VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MouseWheelNotifierObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Public Event MouseWheelRotation(Direction As Long, Handled As Boolean)
Public Event Message(ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean)

Private Const WM_MOUSEWHEEL As Long = &H20A

Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private mFormHwnd As Long
Private mAttached As Boolean

Public Sub SetForm(Optional nForm, Optional nFormHwnd)
    If Not InIDE Then
    'If True Then
        On Error Resume Next
        If Not IsMissing(nForm) Then
            Set mForm = nForm
            mFormHwnd = mForm.hWnd
        Else
            mFormHwnd = nFormHwnd
        End If
        On Error GoTo 0
        If mFormHwnd <> 0 Then
            AttachMessage Me, mFormHwnd, WM_MOUSEWHEEL
            AttachMessage Me, mFormHwnd, WM_DESTROY
            mAttached = True
        End If
    End If
End Sub
    
Private Sub Class_Terminate()
    Detach
End Sub

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrPreprocess
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Dim iDelta As Long
    Dim iDirection As Long
    Dim iHandled As Boolean
    
    Select Case iMsg
        Case WM_MOUSEWHEEL
            If (wParam And 128) = 0 Then ' if not already handled
                iHandled = False
                RaiseEvent Message(iMsg, wParam, lParam, iHandled)
                If iHandled Then
                    wParam = wParam Or 128
                Else
                    iDelta = WordHi(wParam)
                    If iDelta > 1 Then
                        iDirection = -1
                    Else
                        iDirection = 1
                    End If
                    RaiseEvent MouseWheelRotation(iDirection, iHandled)
                    If iHandled Then
                        wParam = wParam Or 128
                    End If
                End If
            End If
        Case WM_DESTROY
            Set mForm = Nothing
            Detach
        Case Else
    End Select
'    ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
End Function

Private Sub mForm_Unload(Cancel As Integer)
    Set mForm = Nothing
    Detach
End Sub

Private Sub Detach()
    If Not mAttached Then Exit Sub
    DetachMessage Me, mFormHwnd, WM_MOUSEWHEEL
    DetachMessage Me, mFormHwnd, WM_DESTROY
    mAttached = False
End Sub

Private Function WordHi(ByVal LongIn As Long) As Integer
    ' Mask off low word then do integer divide to
    ' shift right by 16.
    
    WordHi = (LongIn And &HFFFF0000) \ &H10000
End Function

'Private Function WordLo(ByVal LongIn As Long) As Integer
'    ' Low word retrieved by masking off high word.
'    ' If low word is too large, twiddle sign bit.
'
'    If (LongIn And &HFFFF&) > &H7FFF Then
'        WordLo = (LongIn And &HFFFF&) - &H10000
'    Else
'        WordLo = LongIn And &HFFFF&
'    End If
'End Function

Private Function InIDE() As Boolean
    Static sValue As Long
    
    If sValue = 0 Then
        Err.Clear
        On Error Resume Next
        Debug.Print 1 / 0
        If Err.Number Then
            sValue = 1
        Else
            sValue = 2
        End If
        Err.Clear
    End If
    InIDE = (sValue = 1)
End Function

Public Property Get Form() As Object
    Set Form = mForm
End Property
