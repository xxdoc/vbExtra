Attribute VB_Name = "mVBFunctionsReplacement"
Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Private mInstrEx As New cInsStrEx
Private mWordCount As New cWordCount

Public Function Replace(Text As String, sOld As String, sNew As String, Optional ByVal Start As Long = 1, Optional ByVal Count As Long = 2147483647, Optional ByVal Compare As vbExCompareMethod = vbBinaryCompare) As String
    If InStr(Start, Text, sOld, Compare) = 0 Then
        Replace = Text
    Else
        Replace = Replace09(Text, sOld, sNew, Start, Count, Compare)
    End If
End Function

Public Function InStrCount(ByRef sCheck As String, ByRef sMatch As String, Optional ByVal Start As Long = 1, Optional ByVal Compare As vbExCompareMethod = vbBinaryCompare) As Long
    InStrCount = mInstrEx.InStrCount(sCheck, sMatch, Start, Compare)
End Function

Public Function InStrRev(ByRef sCheck As String, ByRef sMatch As String, Optional ByVal Start As Long, Optional ByVal Compare As vbExCompareMethod = vbBinaryCompare) As Long
'    If InIDE Then
'        If Start = 0 Then
'            InStrRev = VBA.InStrRev(sCheck, sMatch, , Compare)
'        Else
'            InStrRev = VBA.InStrRev(sCheck, sMatch, Start, Compare)
'        End If
'    Else
    InStrRev = mInstrEx.InStrRevEx(sCheck, sMatch, Start, Compare)
'    End If
End Function

Public Function WordCount(s$) As Long
    WordCount = mWordCount.WordCount(s$)
End Function

' by Jost Schwider, jost@schwider.de, 20001218
Private Function Replace09(ByRef Text As String, ByRef sOld As String, ByRef sNew As String, Optional ByVal Start As Long = 1, Optional ByVal Count As Long = 2147483647, Optional ByVal Compare As vbExCompareMethod = vbBinaryCompare) As String
    Dim iAuxLCaseText As String
    
    If LenB(sOld) Then
        If Compare = vbBinaryCompare Then
            Replace09Bin Replace09, Text, Text, sOld, sNew, Start, Count
        Else
             iAuxLCaseText = LCase$(Text)
             Replace09Bin Replace09, Text, iAuxLCaseText, LCase$(sOld), sNew, Start, Count
             ZeroStr iAuxLCaseText
        End If
    Else 'Suchstring ist leer:
        Replace09 = Text
    End If
End Function

Private Sub ZeroStr(nString As String)
    If Len(nString) > 0 Then
        ZeroMemory ByVal StrPtr(nString), LenB(nString)
    End If
End Sub

' by Jost Schwider, jost@schwider.de, 20001218
Private Static Sub Replace09Bin(ByRef Result As String, ByRef Text As String, ByRef Search As String, ByRef sOld As String, ByRef sNew As String, ByVal Start As Long, ByVal Count As Long)
    Dim TextLen As Long
    Dim OldLen As Long
    Dim NewLen As Long
    Dim ReadPos As Long
    Dim WritePos As Long
    Dim CopyLen As Long
    Dim Buffer As String
    Dim BufferLen As Long
    Dim BufferPosNew As Long
    Dim BufferPosNext As Long
    
    'Ersten Treffer bestimmen:
    If Start < 2 Then
        Start = InStrB(Search, sOld)
    Else
        Start = InStrB(Start + Start - 1, Search, sOld)
    End If
    If Start Then
  
    OldLen = LenB(sOld)
    NewLen = LenB(sNew)
    Select Case NewLen
        Case OldLen 'einfaches Überschreiben:
        
            Result = Text
            For Count = 1 To Count
                MidB$(Result, Start) = sNew
                Start = InStrB(Start + OldLen, Search, sOld)
                If Start = 0 Then Exit Sub
            Next Count
            Exit Sub
        
        Case Is < OldLen 'Ergebnis wird kürzer:
        
            'Buffer initialisieren:
            TextLen = LenB(Text)
            If TextLen > BufferLen Then
                Buffer = Text
                BufferLen = TextLen
            End If
          
            'Ersetzen:
            ReadPos = 1
            WritePos = 1
            If NewLen Then
          
            'Einzufügenden Text beachten:
            For Count = 1 To Count
                CopyLen = Start - ReadPos
                If CopyLen Then
                    BufferPosNew = WritePos + CopyLen
                    MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen)
                    MidB$(Buffer, BufferPosNew) = sNew
                    WritePos = BufferPosNew + NewLen
                Else
                    MidB$(Buffer, WritePos) = sNew
                    WritePos = WritePos + NewLen
                End If
                ReadPos = Start + OldLen
                Start = InStrB(ReadPos, Search, sOld)
                If Start = 0 Then Exit For
            Next Count
          
          Else
          
            'Einzufügenden Text ignorieren (weil leer):
            For Count = 1 To Count
                CopyLen = Start - ReadPos
                If CopyLen Then
                    MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen)
                    WritePos = WritePos + CopyLen
                End If
                ReadPos = Start + OldLen
                Start = InStrB(ReadPos, Search, sOld)
                If Start = 0 Then Exit For
            Next Count
          
            End If
          
            'Ergebnis zusammenbauen:
            If ReadPos > TextLen Then
                Result = LeftB$(Buffer, WritePos - 1)
            Else
                MidB$(Buffer, WritePos) = MidB$(Text, ReadPos)
                Result = LeftB$(Buffer, WritePos + LenB(Text) - ReadPos)
            End If
            Exit Sub
        
        Case Else 'Ergebnis wird länger:
        
            'Buffer initialisieren:
            TextLen = LenB(Text)
            BufferPosNew = TextLen + NewLen
            If BufferPosNew > BufferLen Then
                Buffer = Space$(BufferPosNew)
                BufferLen = LenB(Buffer)
            End If
          
            'Ersetzung:
            ReadPos = 1
            WritePos = 1
            For Count = 1 To Count
                CopyLen = Start - ReadPos
                If CopyLen Then
                    'Positionen berechnen:
                    BufferPosNew = WritePos + CopyLen
                    BufferPosNext = BufferPosNew + NewLen
                    
                    'Ggf. Buffer vergrößern:
                    If BufferPosNext > BufferLen Then
                        Buffer = Buffer & Space$(BufferPosNext)
                        BufferLen = LenB(Buffer)
                    End If
                    
                    'String "patchen":
                    MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen)
                    MidB$(Buffer, BufferPosNew) = sNew
                Else
                    'Position bestimmen:
                    BufferPosNext = WritePos + NewLen
                    
                    'Ggf. Buffer vergrößern:
                    If BufferPosNext > BufferLen Then
                        Buffer = Buffer & Space$(BufferPosNext)
                        BufferLen = LenB(Buffer)
                    End If
                    
                    'String "patchen":
                    MidB$(Buffer, WritePos) = sNew
                End If
                WritePos = BufferPosNext
                ReadPos = Start + OldLen
                Start = InStrB(ReadPos, Search, sOld)
                If Start = 0 Then Exit For
            Next Count
          
            'Ergebnis zusammenbauen:
            If ReadPos > TextLen Then
                Result = LeftB$(Buffer, WritePos - 1)
            Else
                BufferPosNext = WritePos + TextLen - ReadPos
                If BufferPosNext < BufferLen Then
                    MidB$(Buffer, WritePos) = MidB$(Text, ReadPos)
                    Result = LeftB$(Buffer, BufferPosNext)
                Else
                    Result = LeftB$(Buffer, WritePos - 1) & MidB$(Text, ReadPos)
                End If
            End If
            Exit Sub
    
    End Select
  
    Else 'Kein Treffer:
        Result = Text
    End If
    
    ZeroStr Buffer
End Sub

