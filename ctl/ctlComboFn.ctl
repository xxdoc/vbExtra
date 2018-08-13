VERSION 5.00
Begin VB.UserControl ComboFn 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MaskColor       =   &H00D8E9EC&
   MaskPicture     =   "ctlComboFn.ctx":0000
   Picture         =   "ctlComboFn.ctx":0E12
   PropertyPages   =   "ctlComboFn.ctx":1C24
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ctlComboFn.ctx":1C35
   Begin VB.Label Label1 
      Caption         =   $"ctlComboFn.ctx":1F47
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   1020
      Width           =   3015
   End
End
Attribute VB_Name = "ComboFn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event AboutToDropDown(ComboName As String, ByRef Cancel As Boolean)

' properties
Private mAutoSizeList As Boolean
Private mShowFullTextOnMouseOver As Boolean
Private mComboBoxName As String

Private mComboHandlersCollection As Collection

Private Sub UserControl_InitProperties()
    mAutoSizeList = True
    mShowFullTextOnMouseOver = True
    mComboBoxName = ""
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mAutoSizeList = PropBag.ReadProperty("AutoSizeList", True)
    mShowFullTextOnMouseOver = PropBag.ReadProperty("ShowFullTextOnMouseOver", True)
    mComboBoxName = PropBag.ReadProperty("ComboBoxName", "")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AutoSizeList", mAutoSizeList, True)
    Call PropBag.WriteProperty("ShowFullTextOnMouseOver", mShowFullTextOnMouseOver, True)
    Call PropBag.WriteProperty("ComboBoxName", mComboBoxName, "")
End Sub

Private Sub UserControl_Resize()
    Dim iH As Long
    Dim iW As Long
    
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
End Sub

Private Sub UserControl_Show()
    If Ambient.UserMode Then
        ShowWindow UserControl.hWnd, SW_HIDE
        StoreCombos
    End If
End Sub

Private Sub StoreCombos()
    Dim iCtl As Control
    Dim iComboHandler As cComboHandler
    
    If Ambient.UserMode Then
        Set mComboHandlersCollection = New Collection
        If mComboBoxName = "" Then
            On Error GoTo TheExit:
            For Each iCtl In Parent.Controls
'                Debug.Print iCtl.Name
                If TypeName(iCtl) = "ComboBox" Then
                    Set iComboHandler = New cComboHandler
                    iComboHandler.SetCombo iCtl, Me, ObjPtr(UserControl.Controls), mAutoSizeList, mShowFullTextOnMouseOver
                    mComboHandlersCollection.Add iComboHandler, CStr(iCtl.hWnd)
                End If
            Next
        Else
            On Error Resume Next
            Set iCtl = Parent.Controls(mComboBoxName)
            On Error GoTo 0
            If Not iCtl Is Nothing Then
                Set iComboHandler = New cComboHandler
                iComboHandler.SetCombo iCtl, Me, ObjPtr(UserControl.Controls), mAutoSizeList, mShowFullTextOnMouseOver
                mComboHandlersCollection.Add iComboHandler, CStr(iCtl.hWnd)
            End If
        End If
    End If
    
TheExit:
End Sub

Private Sub UserControl_Terminate()
    Set mComboHandlersCollection = Nothing
End Sub


Public Property Let AutoSizeList(Optional nCombo As Object, nValue As Boolean)
    Dim iCH As cComboHandler
    
    If nCombo Is Nothing Then
        If nValue <> mAutoSizeList Then
            mAutoSizeList = nValue
            PropertyChanged "AutoSizeList"
            If Not mComboHandlersCollection Is Nothing Then
                For Each iCH In mComboHandlersCollection
                    iCH.AutoSizeList = mAutoSizeList
                Next
            End If
        End If
    Else
        Set iCH = GetComboHandler(nCombo)
        If Not iCH Is Nothing Then
            iCH.AutoSizeList = nValue
        End If
    End If
End Property

Public Property Get AutoSizeList(Optional nCombo As Object) As Boolean
    Dim iCH As cComboHandler
    
    If Not nCombo Is Nothing Then
        Set iCH = GetComboHandler(nCombo)
        If Not iCH Is Nothing Then
            AutoSizeList = iCH.AutoSizeList
        End If
    Else
        AutoSizeList = mAutoSizeList
    End If
End Property


Public Property Let ShowFullTextOnMouseOver(Optional nCombo As Object, nValue As Boolean)
    Dim iCH As cComboHandler
    
    If nCombo Is Nothing Then
        If nValue <> mShowFullTextOnMouseOver Then
            mShowFullTextOnMouseOver = nValue
            PropertyChanged "ShowFullTextOnMouseOver"
            If Not mComboHandlersCollection Is Nothing Then
                For Each iCH In mComboHandlersCollection
                    iCH.ShowFullTextOnMouseOver = mShowFullTextOnMouseOver
                Next
            End If
        End If
    Else
        Set iCH = GetComboHandler(nCombo)
        If Not iCH Is Nothing Then
            iCH.ShowFullTextOnMouseOver = nValue
        End If
    End If
End Property

Public Property Get ShowFullTextOnMouseOver(Optional nCombo As Object) As Boolean
    Dim iCH As cComboHandler
    
    If Not nCombo Is Nothing Then
        Set iCH = GetComboHandler(nCombo)
        If Not iCH Is Nothing Then
            AutoSizeList = iCH.ShowFullTextOnMouseOver
        End If
    Else
        ShowFullTextOnMouseOver = mShowFullTextOnMouseOver
    End If
End Property

Public Property Let ComboBoxName(nValue As String)
    If nValue <> mComboBoxName Then
        mComboBoxName = nValue
        PropertyChanged "ComboBoxName"
    End If
End Property

Public Property Get ComboBoxName() As String
    ComboBoxName = mComboBoxName
End Property


Private Function GetComboHandler(nCombo As Object) As cComboHandler
    On Error Resume Next
    Set GetComboHandler = mComboHandlersCollection(CStr(nCombo.hWnd))
End Function


Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

Public Function IsDropped(Optional nCombo As Object) As Boolean
    Dim iCH As cComboHandler
    
    If nCombo Is Nothing Then
        If Not mComboHandlersCollection Is Nothing Then
            For Each iCH In mComboHandlersCollection
                If iCH.IsDropped Then
                    IsDropped = True
                    Exit For
                End If
            Next
        End If
    Else
        Set iCH = GetComboHandler(nCombo)
        If Not iCH Is Nothing Then
            IsDropped = iCH.IsDropped
        End If
    End If

End Function

Public Sub ProperSizeDropDownWidth(Optional nCombo As Object)
    Dim iCH As cComboHandler
    
    If nCombo Is Nothing Then
        If Not mComboHandlersCollection Is Nothing Then
            For Each iCH In mComboHandlersCollection
                iCH.ProperSizeDropDownWidth
                Exit For
            Next
        End If
    Else
        Set iCH = GetComboHandler(nCombo)
        If Not iCH Is Nothing Then
            iCH.ProperSizeDropDownWidth
        End If
    End If

End Sub

Public Sub RaiseEvent_AboutToDropDown(nComboName As String, ByRef nCancel As Boolean)
Attribute RaiseEvent_AboutToDropDown.VB_MemberFlags = "40"
    RaiseEvent AboutToDropDown(nComboName, nCancel)
End Sub
    
Public Sub UpdateHookedCombos()
    StoreCombos
End Sub

Public Function GetNeededWidth(Optional nCombo As Object) As Long
    Dim iCH As cComboHandler
    
    If nCombo Is Nothing Then
        If mComboHandlersCollection Is Nothing Then
            UpdateHookedCombos
        End If
        If Not mComboHandlersCollection Is Nothing Then
            For Each iCH In mComboHandlersCollection
                GetNeededWidth = iCH.GetNeededWidth
                Exit For
            Next
        End If
    Else
        Set iCH = GetComboHandler(nCombo)
        If iCH Is Nothing Then
            UpdateHookedCombos
            Set iCH = GetComboHandler(nCombo)
        End If
        If Not iCH Is Nothing Then
            GetNeededWidth = iCH.GetNeededWidth
        End If
    End If
    
End Function
