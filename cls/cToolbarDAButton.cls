VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolBarDAButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Enum vbExToolBarDAButtonStyleConstants
    vxTBDefault = 0
    vxTBCheck = 1
    vxTBButtonGroup = 2
    vxTBSeparator = 3
    vxTBPlaceholder = 4
End Enum

Public Enum VBExToobarDAButtonIconSizeConstants
    vxIZ16 = 0
    vxIZ20 = 1
    vxIZ24 = 2
    vxIZ30 = 3
    vxIZ36 = 4
End Enum

Private mIndex As Long
Private mStyle As vbExToolBarDAButtonStyleConstants
Private mKey As String
Private mTag As String
Private mToolTipText As String
Private mWidth As Long
Private mEnabled As Boolean
Private mVisible As Boolean
Private mChecked As Boolean
Private mPic16 As StdPicture
Private mPic20 As StdPicture
Private mPic24 As StdPicture
Private mPic30 As StdPicture
Private mPic36 As StdPicture
Private mPic16Alt As StdPicture
Private mPic20Alt As StdPicture
Private mPic24Alt As StdPicture
Private mPic30Alt As StdPicture
Private mPic36Alt As StdPicture
Private mUseAltPic As Boolean
Private mIconSize As vbExToolbarDAIconsSizeConstants
Private mOrderToHide As Long
Private mHidden As Boolean

Private mButtonControlPtr As Long
Private mParentTBPtr As Long
Private mButtonControlHwnd As Long
Private mParentTBHwnd As Long

Private Sub Class_Initialize()
    mEnabled = True
    mVisible = True
    mWidth = 360
End Sub

Private Sub Class_Terminate()
    Set mPic16 = Nothing
    Set mPic20 = Nothing
    Set mPic24 = Nothing
    Set mPic30 = Nothing
    Set mPic36 = Nothing
    Set mPic16Alt = Nothing
    Set mPic20Alt = Nothing
    Set mPic24Alt = Nothing
    Set mPic30Alt = Nothing
    Set mPic36Alt = Nothing
End Sub


Public Property Get Index() As Long
    Index = mIndex
End Property

Public Property Let Index(ByVal Index As Long)
    mIndex = Index
End Property


Public Property Get Style() As vbExToolBarDAButtonStyleConstants
    Style = mStyle
End Property

Public Property Let Style(ByVal eStyle As vbExToolBarDAButtonStyleConstants)
    Dim iTb As ToolBarDA
    
    If eStyle <> mStyle Then
        mStyle = eStyle
        If mParentTBHwnd <> 0 Then
            If IsWindow(mParentTBHwnd) <> 0 Then
                Set iTb = ParentTB
                If Not iTb Is Nothing Then
                    iTb.DoRefresh
                End If
            End If
        End If
    End If
End Property


Public Property Get OrderToHide() As Long
    OrderToHide = mOrderToHide
End Property

Public Property Let OrderToHide(ByVal lOrderToHide As Long)
    Dim iTb As ToolBarDA
    
    If lOrderToHide <> mOrderToHide Then
        mOrderToHide = lOrderToHide
        If mParentTBHwnd <> 0 Then
            If IsWindow(mParentTBHwnd) <> 0 Then
                Set iTb = ParentTB
                If Not iTb Is Nothing Then
                    iTb.DoRefresh
                End If
            End If
        End If
    End If
End Property


Public Property Get Key() As String
    Key = mKey
End Property

Public Property Let Key(ByVal sKey As String)
    mKey = sKey
End Property


Public Property Get Tag() As String
    Tag = mTag
End Property

Public Property Let Tag(ByVal sTag As String)
    mTag = sTag
End Property


Public Property Get ToolTipText() As String
    ToolTipText = mToolTipText
End Property

Public Property Let ToolTipText(ByVal sToolTipText As String)
    Dim iControl As Object
    
    If sToolTipText <> mToolTipText Then
        mToolTipText = sToolTipText
        If mButtonControlHwnd <> 0 Then
            If IsWindow(mButtonControlHwnd) <> 0 Then
                Set iControl = ButtonControl
                If Not iControl Is Nothing Then
                    If iControl.ToolTipText <> mToolTipText Then
                        iControl.ToolTipText = mToolTipText
                    End If
                End If
            End If
        End If
    End If
End Property


Public Property Get Width() As Long
    Dim iTb As ToolBarDA
    
    If mParentTBHwnd <> 0 Then
        If IsWindow(mParentTBHwnd) <> 0 Then
            Set iTb = ParentTB
            If Not iTb Is Nothing Then
                iTb.EnsureDrawn
            End If
        End If
    End If
    
    Width = mWidth
End Property

Public Property Let Width(ByVal lWidth As Long)
    Dim iTb As ToolBarDA
    
    If lWidth <> mWidth Then
        mWidth = lWidth
        If mParentTBHwnd <> 0 Then
            If IsWindow(mParentTBHwnd) <> 0 Then
                Set iTb = ParentTB
                If Not iTb Is Nothing Then
                    iTb.DoRefresh
                End If
            End If
        End If
    End If
End Property


Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal bEnabled As Boolean)
    Dim iControl As Object
    
    If bEnabled <> mEnabled Then
        mEnabled = bEnabled
        If mButtonControlHwnd <> 0 Then
            If IsWindow(mButtonControlHwnd) <> 0 Then
                Set iControl = ButtonControl
                If Not iControl Is Nothing Then
                    If iControl.Enabled <> mEnabled Then
                        iControl.Enabled = mEnabled
                    End If
                End If
            End If
        End If
    End If
End Property


Public Property Get Visible() As Boolean
    Visible = mVisible
End Property

Public Property Let Visible(ByVal bVisible As Boolean)
    Dim iTb As ToolBarDA
    
    If bVisible <> mVisible Then
        mVisible = bVisible
        If mParentTBHwnd <> 0 Then
            If IsWindow(mParentTBHwnd) <> 0 Then
                Set iTb = ParentTB
                If Not iTb Is Nothing Then
                    iTb.DoRefresh
                End If
            End If
        End If
    End If
End Property


Public Property Get Checked() As Boolean
    Checked = mChecked
End Property

Public Property Let Checked(ByVal bChecked As Boolean)
    Dim iControl As Object
    Dim iTb As ToolBarDA
    
    If bChecked <> mChecked Then
        mChecked = bChecked
        If mButtonControlHwnd <> 0 Then
            If IsWindow(mButtonControlHwnd) <> 0 Then
                Set iControl = ButtonControl
                If Not iControl Is Nothing Then
                    If iControl.Value <> mChecked Then
                        iControl.Value = mChecked
                    End If
                End If
            End If
        End If
        If iControl Is Nothing Then
            Set iTb = ParentTB
            If Not iTb Is Nothing Then
                iTb.ClickButton mIndex
            End If
        End If
    End If
End Property


Public Property Get UseAltPic() As Boolean
    UseAltPic = mUseAltPic
End Property

Public Property Let UseAltPic(ByVal bUseAltPic As Boolean)
    If bUseAltPic <> mUseAltPic Then
        mUseAltPic = bUseAltPic
        SetPics
    End If
End Property


Friend Property Let IconSize(nValue As vbExToolbarDAIconsSizeConstants)
    If nValue <> mIconSize Then
        mIconSize = nValue
        SetPics
    End If
End Property

Public Property Get IconSize() As vbExToolbarDAIconsSizeConstants
    IconSize = mIconSize
End Property


Public Property Get Hidden() As Boolean
    Hidden = mHidden
End Property

Public Property Let Hidden(ByVal bHidden As Boolean)
    mHidden = bHidden
End Property


Public Property Get Pic16() As StdPicture
    Set Pic16 = mPic16
End Property

Public Property Set Pic16(nPic As StdPicture)
    Dim iControl As Object
    
    Set mPic16 = nPic
    If mButtonControlHwnd <> 0 Then
        If Not mUseAltPic Then
            If IsWindow(mButtonControlHwnd) <> 0 Then
                Set iControl = ButtonControl
                If Not iControl Is Nothing Then
                    If mIconSize = vxIconsSmall Then
                        Set iControl.Pic16 = mPic16
                    End If
                End If
            End If
        End If
    End If
End Property


Public Property Get Pic16Alt() As StdPicture
    Set Pic16Alt = mPic16Alt
End Property

Public Property Set Pic16Alt(nPic As StdPicture)
    Dim iControl As Object
    
    Set mPic16Alt = nPic
    If mButtonControlHwnd <> 0 Then
        If mUseAltPic Then
            If IsWindow(mButtonControlHwnd) <> 0 Then
                Set iControl = ButtonControl
                If Not iControl Is Nothing Then
                    If mIconSize = vxIconsSmall Then
                        Set iControl.Pic16 = mPic16Alt
                    End If
                End If
            End If
        End If
    End If
End Property


Public Property Get Pic20() As StdPicture
    Set Pic20 = mPic20
End Property

Public Property Set Pic20(nPic As StdPicture)
    Dim iControl As Object
    
    Set mPic20 = nPic
    If mButtonControlHwnd <> 0 Then
        If Not mUseAltPic Then
            If IsWindow(mButtonControlHwnd) <> 0 Then
                Set iControl = ButtonControl
                If Not iControl Is Nothing Then
                    If mIconSize = vxIconsSmall Then
                        Set iControl.Pic20 = mPic20
                    ElseIf mIconSize = vxIconsMedium Then
                        Set iControl.Pic16 = mPic20
                    End If
                End If
            End If
        End If
    End If
End Property


Public Property Get Pic20Alt() As StdPicture
    Set Pic20Alt = mPic20Alt
End Property

Public Property Set Pic20Alt(nPic As StdPicture)
    Dim iControl As Object
    
    Set mPic20Alt = nPic
    If mButtonControlHwnd <> 0 Then
        If mUseAltPic Then
            If IsWindow(mButtonControlHwnd) <> 0 Then
                Set iControl = ButtonControl
                If Not iControl Is Nothing Then
                    If mIconSize = vxIconsSmall Then
                        Set iControl.Pic20 = mPic20Alt
                    ElseIf mIconSize = vxIconsMedium Then
                        Set iControl.Pic16 = mPic20Alt
                    End If
                End If
            End If
        End If
    End If
End Property


Public Property Get Pic24() As StdPicture
    Set Pic24 = mPic24
End Property

Public Property Set Pic24(nPic As StdPicture)
    Dim iControl As Object
    
    Set mPic24 = nPic
    If mButtonControlHwnd <> 0 Then
        If Not mUseAltPic Then
            If IsWindow(mButtonControlHwnd) <> 0 Then
                Set iControl = ButtonControl
                If Not iControl Is Nothing Then
                    If mIconSize = vxIconsSmall Then
                        Set iControl.Pic24 = mPic24
                    ElseIf mIconSize = vxIconsMedium Then
                        Set iControl.Pic20 = mPic24
                    ElseIf mIconSize = vxIconsBig Then
                        Set iControl.Pic16 = mPic24
                    End If
                End If
            End If
        End If
    End If
End Property


Public Property Get Pic24Alt() As StdPicture
    Set Pic24Alt = mPic24Alt
End Property

Public Property Set Pic24Alt(nPic As StdPicture)
    Dim iControl As Object
    
    Set mPic24Alt = nPic
    If mButtonControlHwnd <> 0 Then
        If mUseAltPic Then
            If IsWindow(mButtonControlHwnd) <> 0 Then
                Set iControl = ButtonControl
                If Not iControl Is Nothing Then
                    If mIconSize = vxIconsSmall Then
                        Set iControl.Pic24 = mPic24Alt
                    ElseIf mIconSize = vxIconsMedium Then
                        Set iControl.Pic20 = mPic24Alt
                    ElseIf mIconSize = vxIconsBig Then
                        Set iControl.Pic16 = mPic24Alt
                    End If
                End If
            End If
        End If
    End If
End Property


Public Property Get Pic30() As StdPicture
    Set Pic30 = mPic30
End Property

Public Property Set Pic30(nPic As StdPicture)
    Dim iControl As Object
    
    Set mPic30 = nPic
    If mButtonControlHwnd <> 0 Then
        If Not mUseAltPic Then
            If IsWindow(mButtonControlHwnd) <> 0 Then
                Set iControl = ButtonControl
                If Not iControl Is Nothing Then
                    If mIconSize = vxIconsMedium Then
                        Set iControl.Pic24 = mPic30
                    ElseIf mIconSize = vxIconsBig Then
                        Set iControl.Pic20 = mPic30
                    End If
                End If
            End If
        End If
    End If
End Property


Public Property Get Pic30Alt() As StdPicture
    Set Pic30Alt = mPic30Alt
End Property

Public Property Set Pic30Alt(nPic As StdPicture)
    Dim iControl As Object
    
    Set mPic30Alt = nPic
    If mButtonControlHwnd <> 0 Then
        If mUseAltPic Then
            If IsWindow(mButtonControlHwnd) <> 0 Then
                Set iControl = ButtonControl
                If Not iControl Is Nothing Then
                    If mIconSize = vxIconsMedium Then
                        Set iControl.Pic24 = mPic30Alt
                    ElseIf mIconSize = vxIconsBig Then
                        Set iControl.Pic20 = mPic30Alt
                    End If
                End If
            End If
        End If
    End If
End Property


Public Property Get Pic36() As StdPicture
    Set Pic36 = mPic36
End Property

Public Property Set Pic36(nPic As StdPicture)
    Dim iControl As Object
    
    Set mPic36 = nPic
    If mButtonControlHwnd <> 0 Then
        If Not mUseAltPic Then
            If IsWindow(mButtonControlHwnd) <> 0 Then
                Set iControl = ButtonControl
                If Not iControl Is Nothing Then
                    If mIconSize = vxIconsBig Then
                        Set iControl.Pic24 = mPic36
                    End If
                End If
            End If
        End If
    End If
End Property


Public Property Get Pic36Alt() As StdPicture
    Set Pic36Alt = mPic36Alt
End Property

Public Property Set Pic36Alt(nPic As StdPicture)
    Dim iControl As Object
    
    Set mPic36Alt = nPic
    If mButtonControlHwnd <> 0 Then
        If mUseAltPic Then
            If IsWindow(mButtonControlHwnd) <> 0 Then
                Set iControl = ButtonControl
                If Not iControl Is Nothing Then
                    If mIconSize = vxIconsBig Then
                        Set iControl.Pic24 = mPic36Alt
                    End If
                End If
            End If
        End If
    End If
End Property


Friend Function Clone() As ToolBarDAButton
    Set Clone = New ToolBarDAButton
    Clone.Checked = mChecked
    Clone.Enabled = mEnabled
    Clone.Index = mIndex
    Clone.Key = mKey
    Set Clone.Pic16 = mPic16
    Set Clone.Pic20 = mPic20
    Set Clone.Pic24 = mPic24
    Set Clone.Pic30 = mPic30
    Set Clone.Pic36 = mPic36
    Set Clone.Pic16Alt = mPic16Alt
    Set Clone.Pic20Alt = mPic20Alt
    Set Clone.Pic24Alt = mPic24Alt
    Set Clone.Pic30Alt = mPic30Alt
    Set Clone.Pic36Alt = mPic36Alt
    Clone.Width = mWidth
    Clone.UseAltPic = mUseAltPic
    Clone.Style = mStyle
    Clone.ToolTipText = mToolTipText
    Clone.Visible = mVisible
    Clone.OrderToHide = mOrderToHide
    Clone.Tag = mTag
End Function


Friend Sub SetParentToolBarDAAndButtonControl(nParentTB As Object, nControl As Object)
    If nParentTB Is Nothing Then
        mParentTBPtr = 0
        mParentTBHwnd = 0
    Else
        mParentTBPtr = ObjPtr(nParentTB)
        mParentTBHwnd = nParentTB.hWnd
    End If
    If nControl Is Nothing Then
        mButtonControlPtr = 0
        mButtonControlHwnd = 0
    Else
        mButtonControlPtr = ObjPtr(nControl)
        mButtonControlHwnd = nControl.hWnd
    End If
End Sub

Private Property Get ButtonControl() As Object
    Dim iTempObj As Object
    
    If mButtonControlPtr = 0 Then Exit Property
    CopyMemory iTempObj, mButtonControlPtr, 4
    Set ButtonControl = iTempObj
    CopyMemory iTempObj, 0&, 4
End Property

Private Property Get ParentTB() As Object
    Dim iTempObj As Object
    
    If mParentTBPtr = 0 Then Exit Property
    CopyMemory iTempObj, mParentTBPtr, 4
    Set ParentTB = iTempObj
    CopyMemory iTempObj, 0&, 4
End Property

Public Sub Click()
    Dim iControl As Object
    Dim iTb As ToolBarDA
    
    If mButtonControlHwnd <> 0 Then
        If IsWindow(mButtonControlHwnd) <> 0 Then
            Set iControl = ButtonControl
            If Not iControl Is Nothing Then
                iControl.Value = True
            End If
        End If
    End If
    If iControl Is Nothing Then
        Set iTb = ParentTB
        If Not iTb Is Nothing Then
            iTb.ClickButton mIndex
        End If
    End If

End Sub

Private Sub SetPics()
    Dim iControl As Object
    
    If mButtonControlHwnd <> 0 Then
        If IsWindow(mButtonControlHwnd) <> 0 Then
            Set iControl = ButtonControl
            If Not iControl Is Nothing Then
                If mUseAltPic Then
                    If mIconSize = vxIconsSmall Then
                        Set iControl.Pic16 = mPic16Alt
                        Set iControl.Pic20 = mPic20Alt
                        Set iControl.Pic24 = mPic24Alt
                    ElseIf mIconSize = vxIconsBig Then
                        Set iControl.Pic16 = mPic24Alt
                        Set iControl.Pic20 = mPic30Alt
                        Set iControl.Pic24 = mPic36Alt
                    Else
                        Set iControl.Pic16 = mPic20Alt
                        Set iControl.Pic20 = mPic24Alt
                        Set iControl.Pic24 = mPic30Alt
                    End If
                Else
                    If mIconSize = vxIconsSmall Then
                        Set iControl.Pic16 = mPic16
                        Set iControl.Pic20 = mPic20
                        Set iControl.Pic24 = mPic24
                    ElseIf mIconSize = vxIconsBig Then
                        Set iControl.Pic16 = mPic24
                        Set iControl.Pic20 = mPic30
                        Set iControl.Pic24 = mPic36
                    Else
                        Set iControl.Pic16 = mPic20
                        Set iControl.Pic20 = mPic24
                        Set iControl.Pic24 = mPic30
                    End If
                End If
            End If
        End If
    End If

End Sub

Public Property Get Left() As Single
    Dim iControl As Object
    Dim iTb As ToolBarDA
    
    If mParentTBHwnd <> 0 Then
        If IsWindow(mParentTBHwnd) <> 0 Then
            Set iTb = ParentTB
            If Not iTb Is Nothing Then
                iTb.EnsureDrawn
            End If
        End If
    End If
    
    If mButtonControlHwnd <> 0 Then
        If IsWindow(mButtonControlHwnd) <> 0 Then
            Set iControl = ButtonControl
            If Not iControl Is Nothing Then
                Left = iControl.Left
            End If
        End If
    End If
End Property


Public Property Get Picture(nSizeIdentifier As VBExToobarDAButtonIconSizeConstants) As StdPicture
    If (nSizeIdentifier < vxIZ16) Or (nSizeIdentifier > vxIZ36) Then
        RaiseError 381, TypeName(Me) ' Invalid property array index
        Exit Property
    End If
    
    Select Case nSizeIdentifier
        Case vxIZ16
            Set Picture = mPic16
        Case vxIZ20
            Set Picture = mPic20
        Case vxIZ24
            Set Picture = mPic24
        Case vxIZ30
            Set Picture = mPic30
        Case vxIZ36
            Set Picture = mPic36
    End Select
End Property

Public Property Set Picture(nSizeIdentifier As VBExToobarDAButtonIconSizeConstants, nPic As StdPicture)
    If (nSizeIdentifier < vxIZ16) Or (nSizeIdentifier > vxIZ36) Then
        RaiseError 381, TypeName(Me) ' Invalid property array index
        Exit Property
    End If
    
    Select Case nSizeIdentifier
        Case vxIZ16
            Set Pic16 = nPic
        Case vxIZ20
            Set Pic20 = nPic
        Case vxIZ24
            Set Pic24 = nPic
        Case vxIZ30
            Set Pic30 = nPic
        Case vxIZ36
            Set Pic36 = nPic
    End Select
End Property

Public Property Let Picture(nSizeIdentifier As VBExToobarDAButtonIconSizeConstants, nPic As StdPicture)
    If (nSizeIdentifier < vxIZ16) Or (nSizeIdentifier > vxIZ36) Then
        RaiseError 381, TypeName(Me) ' Invalid property array index
        Exit Property
    End If
    Set Picture(nSizeIdentifier) = nPic
End Property


Public Property Get PictureAlt(nSizeIdentifier As VBExToobarDAButtonIconSizeConstants) As StdPicture
    If (nSizeIdentifier < vxIZ16) Or (nSizeIdentifier > vxIZ36) Then
        RaiseError 381, TypeName(Me) ' Invalid property array index
        Exit Property
    End If
    
    Select Case nSizeIdentifier
        Case vxIZ16
            Set PictureAlt = mPic16Alt
        Case vxIZ20
            Set PictureAlt = mPic20Alt
        Case vxIZ24
            Set PictureAlt = mPic24Alt
        Case vxIZ30
            Set PictureAlt = mPic30Alt
        Case vxIZ36
            Set PictureAlt = mPic36Alt
    End Select
End Property

Public Property Set PictureAlt(nSizeIdentifier As VBExToobarDAButtonIconSizeConstants, nPic As StdPicture)
    If (nSizeIdentifier < vxIZ16) Or (nSizeIdentifier > vxIZ36) Then
        RaiseError 381, TypeName(Me) ' Invalid property array index
        Exit Property
    End If
    
    Select Case nSizeIdentifier
        Case vxIZ16
            Set Pic16Alt = nPic
        Case vxIZ20
            Set Pic20Alt = nPic
        Case vxIZ24
            Set Pic24Alt = nPic
        Case vxIZ30
            Set Pic30Alt = nPic
        Case vxIZ36
            Set Pic36Alt = nPic
    End Select
End Property

Public Property Let PictureAlt(nSizeIdentifier As VBExToobarDAButtonIconSizeConstants, nPic As StdPicture)
    If (nSizeIdentifier < vxIZ16) Or (nSizeIdentifier > vxIZ36) Then
        RaiseError 381, TypeName(Me) ' Invalid property array index
        Exit Property
    End If
    Set PictureAlt(nSizeIdentifier) = nPic
End Property

