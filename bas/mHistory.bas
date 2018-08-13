Attribute VB_Name = "mHistory"
Option Explicit

Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Public gHistoryControlsCollection As New Collection

Private mWM_HISTORYERASED As Long

Public Property Get WM_HISTORYERASED() As Long
    
    If mWM_HISTORYERASED = 0 Then
        mWM_HISTORYERASED = RegisterWindowMessage(App.Title & "_WM_HISTORYERASED")
    End If
    WM_HISTORYERASED = mWM_HISTORYERASED
End Property

Public Sub AddHistoryControl(nHwnd As Long)
    On Error Resume Next
    gHistoryControlsCollection.Add CVar(nHwnd), CStr(nHwnd)
End Sub
    
Public Sub RemoveHistoryControl(nHwnd As Long)
    On Error Resume Next
    gHistoryControlsCollection.Remove (CStr(nHwnd))
End Sub

Public Sub EraseAllHistories()
    Dim iVar
    
    For Each iVar In gHistoryControlsCollection
        SendMessage CLng(iVar), WM_HISTORYERASED, 0, ByVal CLng(0)
    Next
End Sub
