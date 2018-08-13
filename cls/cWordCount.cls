VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWordCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' by Olaf Schmidt, os@datenhaus.de, 20011128
Option Explicit

Private Declare Function ArrPtr& Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any)
Private Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal nBytes&)

Private saSrcL&(5), SrcL&()

Private Sub Class_Initialize()
  saSrcL(0) = 1: saSrcL(1) = 4: saSrcL(4) = &H7FFFFFFF
  RtlMoveMemory ByVal ArrPtr(SrcL), VarPtr(saSrcL(0)), 4
End Sub


Friend Function WordCount(s$) As Long
Dim i&, SLen&, ub&: Static C32&, C32L&
  SLen = Len(s): If SLen = 0 Then Exit Function
  ub = (SLen - 1) \ 2
  saSrcL(3) = StrPtr(s): C32& = &H20: C32L& = &H20FFFF
  For i = 0 To ub
    If (SrcL(i) And &HFF&) > C32 Then
      WordCount = WordCount + 1
      If SrcL(i) <= C32L Then GoTo m2
      i = i + 1: If (SrcL(i) And &HFF&) <= C32 Then GoTo m1
      If SrcL(i) <= C32L Then GoTo m2
      i = i + 1: If (SrcL(i) And &HFF&) <= C32 Then GoTo m1
      If SrcL(i) <= C32L Then GoTo m2
      i = i + 1: If (SrcL(i) And &HFF&) <= C32 Then GoTo m1
      If SrcL(i) <= C32L Then GoTo m2
      i = i + 1: If (SrcL(i) And &HFF&) <= C32 Then GoTo m1
      Do
        If SrcL(i) <= C32L Then GoTo m2
        i = i + 1: If (SrcL(i) And &HFF&) <= C32 Then GoTo m1
      Loop
    End If
m1: If SrcL(i) > C32L And i <= ub Then
      WordCount = WordCount + 1
      i = i + 1: If (SrcL(i) And &HFF&) <= C32 Then GoTo m1
      If SrcL(i) <= C32L Then GoTo m2
      i = i + 1: If (SrcL(i) And &HFF&) <= C32 Then GoTo m1
      If SrcL(i) <= C32L Then GoTo m2
      i = i + 1: If (SrcL(i) And &HFF&) <= C32 Then GoTo m1
      If SrcL(i) <= C32L Then GoTo m2
      i = i + 1: If (SrcL(i) And &HFF&) <= C32 Then GoTo m1
      If SrcL(i) <= C32L Then GoTo m2
      Do
        i = i + 1: If (SrcL(i) And &HFF&) <= C32 Then GoTo m1
        If SrcL(i) <= C32L Then GoTo m2
      Loop
    End If
m2: Next i
End Function


Private Sub Class_Terminate()
  RtlMoveMemory ByVal ArrPtr(SrcL), 0&, 4 'important
End Sub
