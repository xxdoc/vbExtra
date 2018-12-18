VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cObjectHandlersCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCollection As New Collection

Public Sub Remove(nHwnd As Long)
    On Error Resume Next
    mCollection.Remove CStr(nHwnd)
End Sub

Public Sub Add(nObject As Object, nHwnd As Long)
    mCollection.Add nObject, CStr(nHwnd)
End Sub

Public Function GetInstance(nObjectWithHwnd As Object) As Object
    On Error Resume Next
    Set GetInstance = mCollection(CStr(nObjectWithHwnd.hWnd))
End Function

Public Property Get Count() As Long
    Count = mCollection.Count
End Property
