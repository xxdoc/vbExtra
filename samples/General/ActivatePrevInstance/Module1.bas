Attribute VB_Name = "Module1"
Option Explicit

Public Sub Main()
    If ActivatePrevInstance Then
        Exit Sub
    End If
    Form1.Show 1
End Sub
