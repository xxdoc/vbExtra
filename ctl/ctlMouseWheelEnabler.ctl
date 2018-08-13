VERSION 5.00
Begin VB.UserControl MouseWheelEnabler 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00D8E9EC&
   MaskPicture     =   "ctlMouseWheelEnabler.ctx":0000
   Picture         =   "ctlMouseWheelEnabler.ctx":0E12
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ctlMouseWheelEnabler.ctx":1C26
End
Attribute VB_Name = "MouseWheelEnabler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event MouseWheelRotation(Direction As Long)
Public Event Message(ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean)

Private WithEvents mMouseWheelNotifierObject As MouseWheelNotifierObject
Attribute mMouseWheelNotifierObject.VB_VarHelpID = -1

Private mAutoScrollControls As Boolean
Private mControlToScroll As Object

Private Sub mMouseWheelNotifierObject_Message(ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean)
     RaiseEvent Message(iMsg, wParam, lParam, Handled)
End Sub

Private Sub UserControl_InitProperties()
    mAutoScrollControls = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mAutoScrollControls = PropBag.ReadProperty("AutoScrollControls", True)
    
    If Ambient.UserMode Then
        If mAutoScrollControls Then
            Set mMouseWheelNotifierObject = New MouseWheelNotifierObject
            If TypeOf Parent Is Form Then
                mMouseWheelNotifierObject.SetForm Parent
            Else
                mMouseWheelNotifierObject.SetForm , GetParentFormHwnd(UserControl.hWnd)
            End If
        End If
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "AutoScrollControls", mAutoScrollControls, True
End Sub

Private Sub UserControl_Resize()
    Dim iH As Long
    Dim iW As Long
    
    If Not Ambient.UserMode Then
        iH = UserControl.ScaleY(UserControl.ScaleHeight, vbTwips, vbPixels)
        iW = UserControl.ScaleX(UserControl.ScaleWidth, vbTwips, vbPixels)
        
        If (iH <> 34) Or (iW <> 34) Then
            If (iH <> 34) Then
                iH = 34
            End If
            If (iW <> 34) Then
                iW = 34
            End If
            UserControl.Size UserControl.ScaleX(iW, vbPixels, vbTwips), UserControl.ScaleY(iH, vbPixels, vbTwips)
        End If
    End If
End Sub


Public Property Get AutoScrollControls() As Boolean
Attribute AutoScrollControls.VB_MemberFlags = "200"
    AutoScrollControls = mAutoScrollControls
End Property

Public Property Let AutoScrollControls(ByVal nValue As Boolean)
    If nValue <> mAutoScrollControls Then
        mAutoScrollControls = nValue
        PropertyChanged "AutoScrollControls"
        If mAutoScrollControls Then
            Set mMouseWheelNotifierObject = New MouseWheelNotifierObject
            mMouseWheelNotifierObject.SetForm Parent
        Else
            Set mMouseWheelNotifierObject = Nothing
        End If
    End If
End Property


Public Property Get ControlToScroll() As Object
    Set ControlToScroll = mControlToScroll
End Property

Public Property Set ControlToScroll(nControl As Object)
    Set mControlToScroll = nControl
End Property

Private Function GetScrollableControl(Optional RequirePartiallyVisibleAtLeast As Boolean) As Object
    Dim iCtl As Control
    Dim iControl2 As Object
    Dim iActiveControl As Object
    
    If Not Ambient.UserMode Then Exit Function
    
    On Error Resume Next
    Set iActiveControl = Parent.ActiveControl
    On Error GoTo 0
    If Not iActiveControl Is Nothing Then
        If IsTypeSupported(iActiveControl) Then
            If IsWindowVisibleOnScreen(iActiveControl.hWnd, RequirePartiallyVisibleAtLeast) Then
                Set GetScrollableControl = iActiveControl
            End If
        End If
    End If
    
    If GetScrollableControl Is Nothing Then
        For Each iCtl In Parent.Controls
            If IsTypeSupported(iCtl) Then
                If IsWindowVisibleOnScreen(iCtl.hWnd, RequirePartiallyVisibleAtLeast) Then
                    Set GetScrollableControl = iCtl
                Else
                    If Not RequirePartiallyVisibleAtLeast Then
                        Set iControl2 = iCtl
                    End If
                End If
            End If
        Next
        If Not RequirePartiallyVisibleAtLeast Then
            If GetScrollableControl Is Nothing Then
                Set GetScrollableControl = iControl2
            End If
        End If
    End If
End Function

Private Function IsTypeSupported(nControl As Object) As Boolean
    Dim iTn As String
    
    iTn = LCase$(TypeName(nControl))
    IsTypeSupported = (iTn = "msflexgrid") Or (iTn = "mshflexgrid") Or (iTn = "vscrollbar") Or (iTn = "richtextbox")
    
End Function

Private Function ControlIsGrid(nControl As Object) As Boolean
    Dim iTn As String
    
    iTn = LCase$(TypeName(nControl))
    ControlIsGrid = (iTn = "msflexgrid") Or (iTn = "mshflexgrid")
End Function

Private Sub mMouseWheelNotifierObject_MouseWheelRotation(Direction As Long, Handled As Boolean)
    RaiseEvent MouseWheelRotation(Direction)
    If Not mAutoScrollControls Then Exit Sub
    
    Static sControlToScroll As Object
    Dim iNewControl As Object
    Static sScrollStep As Long
    Static sGridNamePrev As String
    Static sGridRowsPrev As Long
    Dim iControlIsActiveControl As Boolean
    Dim iAuxHeight As Long
    Dim r As Long
    Dim iControlIsGrid As Boolean
    
    On Error GoTo TheExit:
    
    If Not mControlToScroll Is Nothing Then
        Set sControlToScroll = mControlToScroll
        If Not IsWindowVisibleOnScreen(sControlToScroll.hWnd, True) Then Exit Sub
    Else
        If sControlToScroll Is Nothing Then
            Set iNewControl = GetScrollableControl(True)
            If iNewControl Is Nothing Then Exit Sub
            Set sControlToScroll = iNewControl
        End If
        On Error Resume Next
        iControlIsActiveControl = sControlToScroll.Parent.ActiveControl Is sControlToScroll
        On Error GoTo TheExit:
        If Not iControlIsActiveControl Then
            Set iNewControl = GetScrollableControl(True)
            If iNewControl Is Nothing Then Exit Sub
            Set sControlToScroll = iNewControl
        Else
            If Not IsWindowVisibleOnScreen(sControlToScroll.hWnd, True) Then
                Set iNewControl = GetScrollableControl(True)
                If iNewControl Is Nothing Then Exit Sub
                Set sControlToScroll = iNewControl
            End If
        End If
    End If
    iControlIsGrid = ControlIsGrid(sControlToScroll)
    
    If iControlIsGrid Then
        If mGlobals.IsShowingVerticalScrollBar(sControlToScroll) Then
            If (sGridNamePrev <> sControlToScroll.Name) Or (sGridRowsPrev <> sControlToScroll.Rows) Then
                ' calculate the scroll step based on grid variables
                iAuxHeight = 0
                For r = 0 To sControlToScroll.Rows - 1
                    iAuxHeight = iAuxHeight + sControlToScroll.RowHeight(r)
                    If iAuxHeight >= sControlToScroll.Height Then
                        Exit For
                    End If
                Next r
                sScrollStep = r / 4
                If sScrollStep = 0 Then sScrollStep = 1
            End If
            
            If Direction = 1 Then
                If Not sControlToScroll.RowIsVisible(sControlToScroll.Rows - 1) Then
                    If (sControlToScroll.TopRow + sScrollStep) > (sControlToScroll.Rows - 1) Then
                        sControlToScroll.TopRow = sControlToScroll.Rows - 1
                    Else
                        sControlToScroll.TopRow = sControlToScroll.TopRow + sScrollStep
                    End If
                Else
                    sControlToScroll.TopRow = sControlToScroll.Rows - 1
                End If
            Else
                If Not sControlToScroll.RowIsVisible(sControlToScroll.FixedRows) Then
                    If (sControlToScroll.TopRow - sScrollStep) < 1 Then
                        sControlToScroll.TopRow = sControlToScroll.FixedRows
                    Else
                        sControlToScroll.TopRow = sControlToScroll.TopRow - sScrollStep
                    End If
                Else
                    sControlToScroll.TopRow = sControlToScroll.FixedRows
                End If
            End If
        
            sGridNamePrev = sControlToScroll.Name
            sGridRowsPrev = sControlToScroll.Rows
            Handled = True
        End If
    Else
        If Direction = 1 Then
            If (sControlToScroll.Value + sControlToScroll.SmallChange) <= sControlToScroll.Max Then
                sControlToScroll.Value = sControlToScroll.Value + sControlToScroll.SmallChange
            Else
                sControlToScroll.Value = sControlToScroll.Max
            End If
        Else
            If (sControlToScroll.Value - sControlToScroll.SmallChange) >= sControlToScroll.Min Then
                sControlToScroll.Value = sControlToScroll.Value - sControlToScroll.SmallChange
            Else
                sControlToScroll.Value = sControlToScroll.Min
            End If
        End If
        Handled = True
    End If
    Exit Sub
    
TheExit:
    Handled = False
End Sub



