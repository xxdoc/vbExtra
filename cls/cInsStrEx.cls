VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInsStrEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ArrPtr& Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any)
Private Declare Sub RtlMoveMemory Lib "Kernel32" (dst As Any, src As Any, ByVal nBytes&)

Private Header1(5) As Long
Private Header2(5) As Long
Private SafeArray1() As Integer
Private SafeArray2() As Integer

Private Declare Function CharUpperBuffW& Lib "user32" (lpsz As Any, ByVal cchLength&)
Private Declare Function CharUpperBuffA& Lib "user32" (lpsz As Any, ByVal cchLength&)

Private aUChars(&H8000 To &H7FFF) As Integer


Private Sub Class_Initialize()
    ' Set up our template for looking at strings
    Header1(0) = 1              ' Number of dimensions
    Header1(1) = 2              ' Bytes per element (integer = 2)
    Header1(4) = &H7FFFFFFF     ' Array size

    ' Force SafeArray1 to use Header1 as its own header
    RtlMoveMemory ByVal ArrPtr(SafeArray1), VarPtr(Header1(0)), 4

    ' Set up our template for look at search text
    Header2(0) = 1                 ' Number of dimensions
    Header2(1) = 2                 ' Bytes per element (integer = 2)
    Header2(4) = &H7FFFFFFF        ' Array size

    ' Force SafeArray1 to use Header1 as its own header
    RtlMoveMemory ByVal ArrPtr(SafeArray2), VarPtr(Header2(0)), 4

    Dim c As Long
    Dim ret As Long

    'here, we borrow the LUT from UCase04/UCase05:
    For c = &H8000 To &H7FFF: aUChars(c) = c: Next
    If CharUpperBuffW(aUChars(-32768), &H10000) = 0 Then
        ' for 0 to 255 CharUpperBuffA <=> Asc(UCase$(Chr$(c)))
        ret = CharUpperBuffA(aUChars(0), 256 * 2)  '2 bytes/char
    End If
    ' patch the stooges
    ' š 154/353   Š 138/352
    ' œ 156/339   Œ 140/338
    ' ž 158/382   Ž 142/381
    ' ÿ 255/255   Ÿ 159/376
    aUChars(353) = 352
    aUChars(339) = 338
    aUChars(382) = 381
    aUChars(255) = 376
End Sub


Private Sub Class_Terminate()
    ' Make SafeArray1 once again use its own header
    ' If this code doesn't run the IDE will crash
    RtlMoveMemory ByVal ArrPtr(SafeArray1), 0&, 4
    RtlMoveMemory ByVal ArrPtr(SafeArray2), 0&, 4
End Sub


    'by Marzo Sette Torres Junior, marzojr@taskmail.com.br, 20021006
Friend Function InStrCount(ByRef sCheck As String, ByRef sMatch As String, Optional ByVal Start As Long = 1, Optional ByVal Compare As vbExCompareMethod = vbBinaryCompare) As Long
    Dim lLenCheck As Long, lLenMatch As Long
    Dim J As Long, iFirstMatchChar As Integer
    
    'store the length of both strings:
    lLenCheck = Len(sCheck)
    lLenMatch = Len(sMatch)
    If lLenCheck Then
        If lLenMatch Then
            If Start < 1 Then
                'silently correct invalid value:
                Start = 1
            ElseIf Start > lLenCheck Then
                'return zero
                Exit Function
            End If
            'point the arrays to our strings:
            Header1(3) = StrPtr(sCheck)
            Header2(3) = StrPtr(sMatch)
            'here we optimize slightly for Len(sMatch) = 1:
            If lLenMatch > 1 Then
                Dim lDelta As Long
                lDelta = lLenMatch - 1
                If Compare = vbBinaryCompare Then
                    'Store the first char of sMatch in a buffer:
                    iFirstMatchChar = SafeArray2(0)
                    'Here is where things get weird; we compare the first chars
                    'of both strings. But then we continue from the *end* of sMatch
                    'to the start (avoiding recheking the first chars).
                    For Start = Start - 1 To lLenCheck - lDelta
                        If SafeArray1(Start) = iFirstMatchChar Then
                            For J = lDelta To 1 Step -1
                                If SafeArray1(Start + J) <> SafeArray2(J) Then GoTo NotEqual
                            Next
                            InStrCount = InStrCount + 1
                            Start = Start + lDelta
                        End If
NotEqual:
                    Next
                Else
                    'change sMatch to uppercase
                    For J = lDelta To 1 Step -1
                        SafeArray2(J) = aUChars(SafeArray2(J))
                    Next
                    'Store the first char of sMatch in a buffer:
                    iFirstMatchChar = aUChars(SafeArray2(0))
                    'Here is where things get weird; we compare the ucases of
                    'the first chars of both strings. But then we continue
                    'from the *end* of sMatch to the start (avoiding recheking
                    'the first chars).
                    For Start = Start - 1 To lLenCheck - lDelta
                        If aUChars(SafeArray1(Start)) = iFirstMatchChar Then
                            For J = lDelta To 1 Step -1
                                If aUChars(SafeArray1(Start + J)) <> SafeArray2(J) Then GoTo NotEqual2
                            Next
                            InStrCount = InStrCount + 1
                            Start = Start + lDelta
                        End If
NotEqual2:
                    Next
                End If
            Else
                If Compare = vbBinaryCompare Then
                    'Store the only char of sMatch in a buffer:
                    iFirstMatchChar = SafeArray2(0)
                    'here we compare each char with iFirstMatchChar.
                    For Start = Start - 1 To lLenCheck
                        If SafeArray1(Start) = iFirstMatchChar Then
                            InStrCount = InStrCount + 1
                        End If
                    Next
                Else
                    'Store the (UCase$ of the) only char of sMatch in a buffer:
                    iFirstMatchChar = aUChars(SafeArray2(0))
                    'here we compare the UCase$ of each char with iFirstMatchChar.
                    For Start = Start - 1 To lLenCheck
                        If aUChars(SafeArray1(Start)) = iFirstMatchChar Then
                            InStrCount = InStrCount + 1
                        End If
                    Next
                End If
            End If
        End If
    End If
End Function


    'by Marzo Sette Torres Junior, marzojr@taskmail.com.br, 20020625
Friend Function InStrRevEx(ByRef sCheck As String, ByRef sMatch As String, Optional ByVal Start As Long, Optional ByVal Compare As vbExCompareMethod = vbBinaryCompare) As Long
    Dim lLenCheck As Long, lLenMatch As Long
    Dim i As Long, J As Long, iLastMatchChar As Integer
    
    'store the length of both strings:
    lLenCheck = Len(sCheck)
    lLenMatch = Len(sMatch)
    If lLenCheck Then
        If lLenMatch Then
            If Start <= 0 Then
                'no Start was specified, so we start at the end:
                Start = lLenCheck
            ElseIf Start < lLenMatch Then
                'return zero
                Exit Function
            End If
            'point the arrays to our strings:
            Header1(3) = StrPtr(sCheck)
            Header2(3) = StrPtr(sMatch)
            'here we optimize slightly for Len(sMatch) = 1:
            If lLenMatch > 1 Then
                If Compare = vbBinaryCompare Then
                    'Store the last char of sMatch in a buffer:
                    iLastMatchChar = SafeArray2(lLenMatch - 1)
                    'here we start at the *end* of sCheck and compare each char
                    'with iLastMatchChar. If a match is found, we compare from
                    'the next-to-last char of sMatch to the start of sMatch.
                    Do
                        Start = Start - 1
                        If SafeArray1(Start) = iLastMatchChar Then
                            J = lLenMatch - 1
                            i = Start - J
                            Do
                                J = J - 1
                                If SafeArray1(i + J) <> SafeArray2(J) Then GoTo NotEqual
                            Loop While J
                            InStrRevEx = i + 1
                            Exit Function
                        End If
NotEqual:
                    Loop Until Start < lLenMatch
                Else
                    'Store the (UCase$ of) last char of sMatch in a buffer:
                    iLastMatchChar = aUChars(SafeArray2(lLenMatch - 1))
                    'here we start at the *end* of sCheck and compare the UCase$
                    'of each char with iLastMatchChar. If a match is found, we
                    'compare ucases from the next-to-last char of sMatch to the
                    'start of sMatch.
                    Do
                        Start = Start - 1
                        If aUChars(SafeArray1(Start)) = iLastMatchChar Then
                            J = lLenMatch - 1
                            i = Start - J
                            Do
                                J = J - 1
                                If aUChars(SafeArray1(i + J)) <> aUChars(SafeArray2(J)) Then GoTo NotEqual2
                            Loop While J
                            InStrRevEx = i + 1
                            Exit Function
                        End If
NotEqual2:
                    Loop Until Start < lLenMatch
                End If
            Else
                If Compare = vbBinaryCompare Then
                    'Store the only char of sMatch in a buffer:
                    iLastMatchChar = SafeArray2(0)
                    'here we start at the *end* of sCheck and compare each char
                    'with iLastMatchChar.
                    While Start
                        Start = Start - 1
                        If SafeArray1(Start) = iLastMatchChar Then
                            InStrRevEx = Start + 1
                            Exit Function
                        End If
                    Wend
                Else
                    'Store the (UCase$ of the) only char of sMatch in a buffer:
                    iLastMatchChar = aUChars(SafeArray2(0))
                    'here we start at the *end* of sCheck and compare the UCase$
                    'of each char with iLastMatchChar.
                    While Start
                        Start = Start - 1
                        If aUChars(SafeArray1(Start)) = iLastMatchChar Then
                            InStrRevEx = Start + 1
                            Exit Function
                        End If
                    Wend
                End If
            End If
        Else
            If Start <= lLenCheck Then InStrRevEx = Start
        End If
    End If
End Function
