VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolBarDAButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Local objects
Private mItems As Collection

Friend Sub AddButtonObject(ByVal nNew As ToolBarDAButton, Optional ByVal nKey As Variant, Optional nBefore, Optional nAfter)
    mItems.Add nNew, nKey, nBefore, nAfter
End Sub

Public Function Count() As Long
Attribute Count.VB_UserMemId = -531
    Count = mItems.Count
End Function

Public Function Item(ByVal IndexOrKey As Variant) As ToolBarDAButton
Attribute Item.VB_UserMemId = 0
    On Error Resume Next
    Set Item = mItems(IndexOrKey)
    If Err.Number Then
        Err.Clear
        Err.Raise 5 ' Invalid procedure call or argument
        Dim iStr As String
        iStr = Err.Description
        On Error GoTo 0
        RaiseError 5, TypeName(Me), iStr & ". Button doesn't exist."
    End If
End Function

Friend Sub Remove(ByVal IndexOrKey As Variant)
    mItems.Remove IndexOrKey
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mItems.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set mItems = New Collection
End Sub

