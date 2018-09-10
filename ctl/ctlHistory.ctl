VERSION 5.00
Begin VB.UserControl History 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "ctlHistory.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ctlHistory.ctx":0040
   Begin VB.Timer tmrTextChanged2 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   690
      Top             =   960
   End
   Begin vbExtra.ComboFn ComboFn1 
      Height          =   510
      Left            =   1020
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2670
      Width           =   510
      _ExtentX        =   720
      _ExtentY        =   720
      ShowFullTextOnMouseOver=   0   'False
      ComboBoxName    =   "cboHistory"
   End
   Begin VB.Timer tmrCurrentText 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   180
      Top             =   1500
   End
   Begin VB.Timer tmrTextChanged 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   180
      Top             =   960
   End
   Begin vbExtra.ButtonEx cmdHistoryBack 
      Height          =   372
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Ir a ítem anterior (o click con el botón derecho para seleccionar)"
      Top             =   0
      Width           =   336
      _ExtentX        =   593
      _ExtentY        =   656
      ButtonStyle     =   8
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Pic16           =   "ctlHistory.ctx":0352
      Pic24           =   "ctlHistory.ctx":052E
      Pic20           =   "ctlHistory.ctx":094A
      PictureAlign    =   4
      UseMaskCOlor    =   -1  'True
   End
   Begin vbExtra.ButtonEx cmdHistoryForward 
      Height          =   240
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "Ir a ítem siguiente (o click con el botón derecho para seleccionar)"
      Top             =   36
      Width           =   216
      _ExtentX        =   381
      _ExtentY        =   423
      ButtonStyle     =   8
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Pic16           =   "ctlHistory.ctx":0C26
      Pic24           =   "ctlHistory.ctx":0E02
      Pic20           =   "ctlHistory.ctx":121E
      PictureAlign    =   4
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.PictureBox picCoveringCombo 
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   120
      ScaleHeight     =   348
      ScaleWidth      =   3012
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      Width           =   3015
   End
   Begin VB.ComboBox cboHistory 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   750
      Style           =   2  'Dropdown List
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   270
      Width           =   2115
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "# Delete"
      End
   End
End
Attribute VB_Name = "History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Private Const WM_UILANGCHANGED As Long = WM_USER + 12

Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Const WM_RBUTTONDOWN As Long = &H204&
Private Const LB_ITEMFROMPOINT = &H1A9
Private Const CB_SHOWDROPDOWN = &H14F

Public Event ConfigClick(ByRef nCancel As Boolean)
Public Event Click(nText As String)
Public Event BeforeClick()
Public Event Updated()
Public Event BeforeAddItem(ByRef nCancel As Boolean)
Public Event GetTextToDisplay(ByVal nTextInControl As String, ByRef nTextToDisplay As String)
Attribute GetTextToDisplay.VB_MemberFlags = "200"

Private mHistoryItems As Variant
Private mTextsToDisplay As Variant
Private mItemsTags As Variant
Private mPosition As Long
Public mCurrentText As String
Private mAutoToolTipText As Boolean
Private mAutoShowConfig As Boolean
Private mToolTipTextStart As String
Private mToolTipTextEnd As String
Private mToolTipTextSelect As String
Private mEnableToConfigure As Boolean
Private mEnabled As Boolean
Private mContext As String
Private mBoundControlName As String
Private mBoundProperty As String
Private mBoundControlTag As String
Private mHistoriesCollection As Collection
Private mAutoAddItemEnabled As Boolean
Private mBackColor As Long
Private mShowHistoryMenu As Boolean

Private mAmbientUserMode As Boolean
Private mUserControlHwnd As Long
Private mcboHistoryListHwnd As Long
Private mSelectedItemToDelete As Long
Private mPopupShown As Boolean
Private mHistoryLoaded As Boolean
Private mAmbientDesignModeParent As Boolean
Private mButtonStyle As vbExButtonStyleConstants

Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Private Sub cboHistory_DropDown()
    On Error Resume Next
    ComboFn1.ProperSizeDropDownWidth
End Sub

Private Sub mForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If tmrTextChanged.Enabled Then
        If mHistoryLoaded Then
            tmrTextChanged_Timer
        Else
            tmrTextChanged.Enabled = False
        End If
    End If
End Sub

Private Sub mnuDelete_Click()
    On Error Resume Next
    RemoveFromHistory IndexInList(mTextsToDisplay, cboHistory.List(mSelectedItemToDelete))
End Sub

Private Sub tmrCurrentText_Timer()
    If (Not Ambient.UserMode) Or mAmbientDesignModeParent Then
        tmrCurrentText.Enabled = False
        Exit Sub
    End If
    
    If mAutoAddItemEnabled Then
        UpdateCurrentText
    End If
End Sub

Private Sub UpdateCurrentText()
    Static sAnt As String
    Dim iText As String
    Dim iCtl As Control
    
    On Error Resume Next
    Set iCtl = Parent.Controls(mBoundControlName)
    On Error GoTo 0
    If Not iCtl Is Nothing Then
        iText = CallByName(iCtl, mBoundProperty, VbGet)
    End If
    iText = Trim$(iText)
    
    If iText <> sAnt Then
        CurrentText = iText
    End If
    sAnt = mCurrentText

End Sub

Private Sub tmrTextChanged_Timer()
    tmrTextChanged.Enabled = False
    If mCurrentText <> "" Then
        If mAutoAddItemEnabled Then
            AddItem mCurrentText
        End If
    End If
End Sub

Private Sub tmrTextChanged2_Timer()
    Static sVeces As Long
    
    If Not mAutoAddItemEnabled Then Exit Sub
    sVeces = sVeces + 1
    If sVeces >= 2 Then ' si queda 16 segundos sin cambiar, el ítem actual pasa a ser el último
        sVeces = 0
        tmrTextChanged2.Enabled = False
        If mCurrentText <> "" Then
            mPosition = 0
            AddItem mCurrentText
        End If
    End If
End Sub

Private Sub UserControl_Hide()
    'If tmrTextChanged.Enabled Then tmrTextChanged_Timer
    tmrTextChanged.Enabled = False
    tmrCurrentText.Enabled = False
End Sub

Private Sub UserControl_Initialize()
    
    ReDim mHistoryItems(0)
    ReDim mTextsToDisplay(0)
    ReDim mItemsTags(0)
    mPosition = 0
    
    mAutoToolTipText = True
    mAutoShowConfig = True
    mEnableToConfigure = True
    mToolTipTextSelect = GetLocalizedString(efnGUIStr_History_ToolTipTextSelect_Default)
    mToolTipTextStart = GetLocalizedString(efnGUIStr_History_ToolTipTextStart_Default)
    mToolTipTextEnd = ""
    mAutoAddItemEnabled = True
    mShowHistoryMenu = True
    
    Set mHistoriesCollection = New Collection
End Sub

Private Sub UserControl_InitProperties()
    On Error Resume Next
    mContext = UserControl.Parent.Name
    mContext = mContext & "_" & Ambient.DisplayName
    mAmbientUserMode = Ambient.UserMode
    If mAmbientUserMode Then
        mUserControlHwnd = UserControl.hWnd
        AddHistoryControl mUserControlHwnd
        AttachMessage Me, mUserControlHwnd, WM_HISTORYERASED
        mcboHistoryListHwnd = GetComboListHwnd(cboHistory)
        If mcboHistoryListHwnd <> 0 Then
            AttachMessage Me, mcboHistoryListHwnd, WM_RBUTTONDOWN
        End If
    End If
    mBackColor = vbButtonFace
    mButtonStyle = vxInstallShieldToolbar
    On Error GoTo 0

    mAmbientUserMode = Ambient.UserMode
    If mAmbientUserMode Then
        mUserControlHwnd = UserControl.hWnd
        AddHistoryControl mUserControlHwnd
        On Error Resume Next
        AttachMessage Me, mUserControlHwnd, WM_HISTORYERASED
        mcboHistoryListHwnd = GetComboListHwnd(cboHistory)
        If mcboHistoryListHwnd <> 0 Then
            AttachMessage Me, mcboHistoryListHwnd, WM_RBUTTONDOWN
        End If
        On Error GoTo 0
        If TypeOf Parent Is Form Then Set mForm = Parent
        AttachMessage Me, mUserControlHwnd, WM_UILANGCHANGED
        SetProp mUserControlHwnd, "FnExUI", 1
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mAutoToolTipText = PropBag.ReadProperty("AutoToolTipText", True)
    mAutoShowConfig = PropBag.ReadProperty("AutoShowConfig", True)
    mEnableToConfigure = PropBag.ReadProperty("EnableToConfigure", True)
    Enabled = PropBag.ReadProperty("Enabled", True)
    mToolTipTextStart = PropBag.ReadProperty("ToolTipTextStart", GetLocalizedString(efnGUIStr_History_ToolTipTextStart_Default))
    mToolTipTextEnd = PropBag.ReadProperty("ToolTipTextEnd", "")
    mToolTipTextSelect = PropBag.ReadProperty("ToolTipTextSelect", GetLocalizedString(efnGUIStr_History_ToolTipTextSelect_Default))
    mAutoAddItemEnabled = PropBag.ReadProperty("AutoAddItemEnabled", True)
    
    mContext = PropBag.ReadProperty("Context", "")
    If mContext = "" Then
        On Error Resume Next
        mContext = UserControl.Parent.Name
        On Error GoTo 0
    End If
    mBoundControlName = PropBag.ReadProperty("BoundControlName", "")
    mBoundProperty = PropBag.ReadProperty("BoundProperty", "")
    mBoundControlTag = PropBag.ReadProperty("BoundControlTag", "")
    tmrCurrentText.Enabled = (mBoundControlName <> "") And (mBoundProperty <> "")
    BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    mShowHistoryMenu = PropBag.ReadProperty("ShowHistoryMenu", True)
    cmdHistoryBack.ToolTipText = PropBag.ReadProperty("BackButtonToolTipText", GetLocalizedString(efnGUIStr_History_BackButtonToolTipText_Default))
    cmdHistoryForward.ToolTipText = PropBag.ReadProperty("ForwardButtonToolTipText", GetLocalizedString(efnGUIStr_History_ForwardButtonToolTipText_Default))
    ButtonStyle = PropBag.ReadProperty("ButtonStyle", vxInstallShieldToolbar)
    
    mAmbientUserMode = Ambient.UserMode
    If mAmbientUserMode Then
        mUserControlHwnd = UserControl.hWnd
        AddHistoryControl mUserControlHwnd
        On Error Resume Next
        AttachMessage Me, mUserControlHwnd, WM_HISTORYERASED
        mcboHistoryListHwnd = GetComboListHwnd(cboHistory)
        If mcboHistoryListHwnd <> 0 Then
            AttachMessage Me, mcboHistoryListHwnd, WM_RBUTTONDOWN
        End If
        On Error GoTo 0
        If TypeOf Parent Is Form Then Set mForm = Parent
        AttachMessage Me, mUserControlHwnd, WM_UILANGCHANGED
        SetProp mUserControlHwnd, "FnExUI", 1
    End If
    
End Sub

Public Sub AddItem(nItem As String, Optional nAtTop As Boolean)
    AddToHistory nItem, nAtTop
    EnableDisableButtons
End Sub

Public Sub AddCurrentTextToHistory(Optional nAtTop As Boolean)
    tmrCurrentText_Timer
    AddItem mCurrentText, nAtTop
End Sub

Public Sub ClearCurrentHistory()
    ReDim mHistoryItems(0)
    ReDim mTextsToDisplay(0)
    ReDim mItemsTags(0)
    mPosition = 0
    mCurrentText = ""
    tmrTextChanged.Enabled = False
    tmrTextChanged2.Enabled = False
    EnableDisableButtons
End Sub

Public Property Let Position(nValue As Long)
    If nValue <> mPosition Then
        If nValue < 0 Then Exit Property
        If nValue > UBound(mHistoryItems) Then Exit Property
        mPosition = nValue
        EnableDisableButtons
    End If
End Property

Public Property Get Position() As Long
Attribute Position.VB_MemberFlags = "400"
    Position = mPosition
End Property

Public Property Get ItemCount() As Long
    ItemCount = UBound(mHistoryItems)
End Property

Public Property Get Item(nIndex As Long)
    Item = mHistoryItems(nIndex)
End Property

Public Property Get ItemsCollection() As Collection
    Dim iCol As New Collection
    Dim c As Long
    
    If tmrTextChanged.Enabled Then tmrTextChanged_Timer
    
    For c = 1 To UBound(mHistoryItems)
        iCol.Add CVar(mHistoryItems(c))
    Next c
    
    Set ItemsCollection = iCol
End Property

Public Property Set ItemsCollection(nItemsCollection As Collection)
    Dim c As Long
    Dim iTextToDisplay As String
    Dim iText As String
    
    ClearCurrentHistory
    If nItemsCollection Is Nothing Then
        Exit Property
    End If
    
    ReDim mHistoryItems(nItemsCollection.Count)
    ReDim mTextsToDisplay(nItemsCollection.Count)
    ReDim mItemsTags(nItemsCollection.Count)
    
    For c = 1 To nItemsCollection.Count
        iText = CStr(nItemsCollection(c))
        mHistoryItems(c) = iText
'        iTextToDisplay = Chr(34) & iText & Chr(34)
        iTextToDisplay = iText
        RaiseEvent GetTextToDisplay(iText, iTextToDisplay)
        mTextsToDisplay(c) = iTextToDisplay
    Next
    
    Position = UBound(mHistoryItems)
End Property

Private Sub cboHistory_Click()
    Dim iCancel As Boolean
    
    If mPopupShown Then Exit Sub
    
    On Error GoTo TheExit:
    
    If mEnableToConfigure Then
        Select Case cboHistory.ListIndex
            Case cboHistory.ListCount - 1
                RaiseEvent ConfigClick(iCancel)
                If mAutoShowConfig Then
                    If Not iCancel Then
                        
                        frmConfigHistory.Context = mContext
                        frmConfigHistory.Show 1
                        If frmConfigHistory.HistoryErased Then
                            ClearCurrentHistory
                            Set mHistoriesCollection = New Collection
                        End If
                        Set frmConfigHistory = Nothing
                    
                    End If
                End If
            Case cboHistory.ListCount - 2
            Case Else
                If tmrTextChanged.Enabled Then tmrTextChanged_Timer
                'AddToHistory CStr(mHistoryItems(cboHistory.ItemData(cboHistory.ListIndex)))
                RaiseEvent BeforeClick
                mPosition = cboHistory.ItemData(cboHistory.ListIndex) 'UBound(mHistoryItems)
                If mPosition <= UBound(mHistoryItems) Then
                    RaiseEventClick (CStr(mHistoryItems(mPosition)))
                    UpdateTT
                    EnableDisableButtons
                    tmrTextChanged2.Enabled = False
                    tmrTextChanged2.Enabled = True
                End If
        End Select
    Else
        If tmrTextChanged.Enabled Then tmrTextChanged_Timer
        RaiseEvent BeforeClick
        mPosition = cboHistory.ItemData(cboHistory.ListIndex)
        RaiseEventClick (CStr(mHistoryItems(mPosition)))
        UpdateTT
        EnableDisableButtons
    End If
    SendKeysAPI "{TAB}"
    
TheExit:
End Sub

Private Sub cmdHistoryBack_Click()
    Dim iAuxAutoAddItemEnabled As Boolean
    
    iAuxAutoAddItemEnabled = mAutoAddItemEnabled
'    mAutoAddItemEnabled = True
    If Not iAuxAutoAddItemEnabled Then
        UpdateCurrentText
    End If
    If tmrTextChanged.Enabled Or Not iAuxAutoAddItemEnabled Then
        tmrTextChanged_Timer
    End If
    mAutoAddItemEnabled = iAuxAutoAddItemEnabled
    If mPosition = 0 Then Exit Sub
    RaiseEvent BeforeClick
    If mCurrentText = mHistoryItems(mPosition) Then
        mPosition = mPosition - 1
    End If

    RaiseEventClick (CStr(mHistoryItems(mPosition)))
    UpdateTT
    EnableDisableButtons
    tmrTextChanged2.Enabled = False
    tmrTextChanged2.Enabled = True
End Sub

Private Sub cmdHistoryBack_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        LoadHistoryCombo
        SendMessage cboHistory.hWnd, CB_SHOWDROPDOWN, True, ByVal CLng(0)
    End If
End Sub

Private Sub cmdHistoryForward_Click()
    If tmrTextChanged.Enabled Then tmrTextChanged_Timer
    RaiseEvent BeforeClick
    mPosition = mPosition + 1
    If mPosition > UBound(mHistoryItems) Then
        mPosition = UBound(mHistoryItems)
    End If
    RaiseEventClick (CStr(mHistoryItems(mPosition)))
    UpdateTT
    EnableDisableButtons
    tmrTextChanged2.Enabled = False
    tmrTextChanged2.Enabled = True
End Sub

Private Sub cmdHistoryForward_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        LoadHistoryCombo
        SendMessage cboHistory.hWnd, CB_SHOWDROPDOWN, True, ByVal CLng(0)
    End If
End Sub

Private Sub AddToHistory(nText As String, Optional nAtTop As Boolean)
    Dim iIndexFound As Long
    Dim c As Long
    Dim iCancel As Boolean
    Dim iTextToDisplay As String
    Dim iDo As Boolean
    Dim iAuxTag As String
    
    If Not mHistoryLoaded Then LoadHistory
    
    RaiseEvent BeforeAddItem(iCancel)
    If iCancel Then Exit Sub
'    iTextToDisplay = Chr(34) & nText & Chr(34)
    iTextToDisplay = nText
    RaiseEvent GetTextToDisplay(nText, iTextToDisplay)
    
    If (mHistoryItems(mPosition) <> nText) Then
        iDo = True
    ElseIf nAtTop Then
        If (mHistoryItems(UBound(mHistoryItems)) <> nText) Then
            iDo = True
        End If
    End If
    If iDo Then
        For c = 1 To UBound(mHistoryItems)
            If mHistoryItems(c) = nText Then
                iIndexFound = c
                Exit For
            End If
        Next c
        If iIndexFound <> 0 Then
            iAuxTag = mItemsTags(iIndexFound)
            For c = 1 + iIndexFound To UBound(mHistoryItems)
                mHistoryItems(c - 1) = mHistoryItems(c)
                mTextsToDisplay(c - 1) = mTextsToDisplay(c)
                mItemsTags(c - 1) = mItemsTags(c)
            Next c
        Else
            ReDim Preserve mHistoryItems(UBound(mHistoryItems) + 1)
            ReDim Preserve mTextsToDisplay(UBound(mHistoryItems))
            ReDim Preserve mItemsTags(UBound(mHistoryItems))
        End If
        mHistoryItems(UBound(mHistoryItems)) = nText
        mTextsToDisplay(UBound(mHistoryItems)) = iTextToDisplay
        mPosition = UBound(mHistoryItems)
        mItemsTags(mPosition) = iAuxTag
        UpdateTT
    End If
End Sub

Private Sub EnableDisableButtons()
    
    If UBound(mHistoryItems) = 0 Then
        cmdHistoryBack.Enabled = False
        cmdHistoryForward.Enabled = False
    Else
        cmdHistoryBack.Enabled = mPosition > 1 Or mCurrentText <> mHistoryItems(mPosition)
        cmdHistoryForward.Enabled = mPosition < UBound(mHistoryItems)
    End If
    
End Sub

Private Sub LoadHistoryCombo()
    Dim c As Long
    
    cboHistory.Clear
    If mShowHistoryMenu Then
        For c = UBound(mHistoryItems) To 1 Step -1
            If mHistoryItems(c) <> mCurrentText Then
                cboHistory.AddItem mTextsToDisplay(c)
                cboHistory.ItemData(cboHistory.NewIndex) = c
            End If
        Next c
    End If
    If mEnableToConfigure Then
        If mShowHistoryMenu Then
            cboHistory.AddItem "--   --   --"
        End If
        cboHistory.AddItem "Configurar"
    End If
End Sub

Public Property Let CurrentText(ByVal nText As String)
Attribute CurrentText.VB_MemberFlags = "400"
    nText = Trim$(nText)
    If nText <> mCurrentText Then
        mCurrentText = nText
        tmrTextChanged.Enabled = False
        tmrTextChanged2.Enabled = False
        tmrTextChanged.Enabled = True
        If CStr(mHistoryItems(mPosition)) <> "" Then
            If mCurrentText <> mHistoryItems(mPosition) Then
                cmdHistoryBack.Enabled = True
            End If
        End If
    End If
End Property

Private Sub UserControl_Resize()
    Static sInside As Boolean
    
    If sInside Then Exit Sub
    sInside = True
    PositionControls
    sInside = False
End Sub

Private Sub UserControl_Show()
    If Ambient.UserMode Then
        If Not mHistoryLoaded Then LoadHistory
        tmrCurrentText.Enabled = (mBoundControlName <> "") And (mBoundProperty <> "")
        If tmrCurrentText.Enabled Then
            tmrCurrentText_Timer
            If tmrTextChanged.Enabled Then
                tmrTextChanged_Timer
            End If
        End If
    End If
End Sub

Private Sub UserControl_Terminate()
    Dim c As Long
    Dim iAuxCN As String
    Dim iAuxCT As String
    Dim iLastCN As String
    Dim iLastCT As String
    
    If mAmbientUserMode Then
        If mUserControlHwnd <> 0 Then
            RemoveHistoryControl mUserControlHwnd
            On Error Resume Next
            DetachMessage Me, mUserControlHwnd, WM_HISTORYERASED
            If mcboHistoryListHwnd <> 0 Then
                DetachMessage Me, mcboHistoryListHwnd, WM_RBUTTONDOWN
            End If
            On Error GoTo 0
        End If
        tmrCurrentText.Enabled = False
        SaveHistory
        iLastCN = mBoundControlName
        iLastCT = mBoundControlTag
        If mHistoriesCollection.Count > 0 Then
            For c = 1 To mHistoriesCollection.Count Step 6
                iAuxCN = CStr(mHistoriesCollection(c + 2))
                iAuxCT = CStr(mHistoriesCollection(c + 3))
                If (iAuxCN <> iLastCN) Or (iAuxCT <> iLastCT) Then
                    mHistoryItems = mHistoriesCollection(c)
                    mTextsToDisplay = mHistoriesCollection(c + 5)
                    mBoundControlName = iAuxCN
                    mBoundControlTag = iAuxCT
                    SaveHistory
                End If
            Next
        End If
    End If
    If tmrTextChanged.Enabled Then tmrTextChanged.Enabled = False
    If tmrTextChanged2.Enabled Then tmrTextChanged2.Enabled = False
    
    If mUserControlHwnd <> 0 Then
        DetachMessage Me, mUserControlHwnd, WM_UILANGCHANGED
        RemoveProp mUserControlHwnd, "FnExUI"
    End If
End Sub


Public Property Let ForwardButtonToolTipText(nText As String)
    If nText <> cmdHistoryForward.ToolTipText Then
        cmdHistoryForward.ToolTipText = nText
        PropertyChanged "ForwardButtonToolTipText"
    End If
End Property

Public Property Get ForwardButtonToolTipText() As String
    ForwardButtonToolTipText = cmdHistoryForward.ToolTipText
End Property


Public Property Let BackButtonToolTipText(nText As String)
    If nText <> cmdHistoryBack.ToolTipText Then
        cmdHistoryBack.ToolTipText = nText
        PropertyChanged "BackButtonToolTipText"
    End If
End Property

Public Property Get BackButtonToolTipText() As String
    BackButtonToolTipText = cmdHistoryBack.ToolTipText
End Property


Public Property Let AutoToolTipText(nValue As Boolean)
    If nValue <> mAutoToolTipText Then
        mAutoToolTipText = nValue
        PropertyChanged "AutoToolTipText"
    End If
End Property

Public Property Get AutoToolTipText() As Boolean
    AutoToolTipText = mAutoToolTipText
End Property


Public Property Let AutoShowConfig(nValue As Boolean)
    If nValue <> mAutoShowConfig Then
        mAutoShowConfig = nValue
        PropertyChanged "AutoShowConfig"
    End If
End Property

Public Property Get AutoShowConfig() As Boolean
    AutoShowConfig = mAutoShowConfig
End Property


Public Property Let ToolTipTextStart(nValue As String)
    If nValue <> mToolTipTextStart Then
        mToolTipTextStart = nValue
        PropertyChanged "ToolTipTextStart"
        If Ambient.UserMode Then UpdateTT
    End If
End Property

Public Property Get ToolTipTextStart() As String
    ToolTipTextStart = mToolTipTextStart
End Property


Public Property Let ToolTipTextEnd(nValue As String)
    If nValue <> mToolTipTextEnd Then
        mToolTipTextEnd = nValue
        PropertyChanged "ToolTipTextEnd"
        If Ambient.UserMode Then UpdateTT
    End If
End Property

Public Property Get ToolTipTextEnd() As String
    ToolTipTextEnd = mToolTipTextEnd
End Property


Public Property Let ToolTipTextSelect(nValue As String)
    If nValue <> mToolTipTextSelect Then
        mToolTipTextSelect = nValue
        PropertyChanged "ToolTipTextSelect"
        If Ambient.UserMode Then UpdateTT
    End If
End Property

Public Property Get ToolTipTextSelect() As String
    ToolTipTextSelect = mToolTipTextSelect
End Property


Public Property Let EnableToConfigure(nValue As Boolean)
    If nValue <> mEnableToConfigure Then
        mEnableToConfigure = nValue
        PropertyChanged "EnableToConfigure"
    End If
End Property

Public Property Get EnableToConfigure() As Boolean
    EnableToConfigure = mEnableToConfigure
End Property


Public Property Let ShowHistoryMenu(nValue As Boolean)
    If nValue <> mShowHistoryMenu Then
        mShowHistoryMenu = nValue
        PropertyChanged "ShowHistoryMenu"
    End If
End Property

Public Property Get ShowHistoryMenu() As Boolean
    ShowHistoryMenu = mShowHistoryMenu
End Property


Public Property Let AutoAddItemEnabled(nValue As Boolean)
    If nValue <> mAutoAddItemEnabled Then
        mAutoAddItemEnabled = nValue
        PropertyChanged "AutoAddItemEnabled"
    End If
    If Ambient.UserMode Then
        If mAutoAddItemEnabled Then
            tmrCurrentText_Timer
'            UpdateTT
            tmrTextChanged.Enabled = True
'            tmrTextChanged_Timer
            tmrTextChanged2.Enabled = False
        End If
    End If
End Property

Public Property Get AutoAddItemEnabled() As Boolean
    AutoAddItemEnabled = mAutoAddItemEnabled
End Property


Public Property Let Enabled(nValue As Boolean)
    If nValue <> mEnabled Then
        mEnabled = nValue
        If mEnabled Then
            UserControl.Enabled = True
            EnableDisableButtons
        Else
            UserControl.Enabled = True
            cmdHistoryBack.Enabled = False
            cmdHistoryForward.Enabled = False
        End If
        PropertyChanged "Enabled"
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property


Public Property Let Context(nValue As String)
    If nValue <> mContext Then
        mContext = nValue
        PropertyChanged "Context"
    End If
End Property

Public Property Get Context() As String
    Context = mContext
End Property


Public Property Let BackColor(nValue As OLE_COLOR)
    If nValue <> mBackColor Then
        mBackColor = nValue
        PropertyChanged "BackColor"
        UserControl.BackColor = mBackColor
        cmdHistoryBack.BackColor = mBackColor
        cmdHistoryForward.BackColor = mBackColor
        picCoveringCombo.BackColor = mBackColor
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property


Public Property Let BoundControlName(nValue As String)
    If mBoundControlName <> nValue Then
        If Ambient.UserMode Then
            Err.Raise 5676, "At run time use ChangeBoundControl Method"
        End If
        mBoundControlName = nValue
        PropertyChanged "BoundControlName"
    End If
End Property

Public Property Get BoundControlName() As String
Attribute BoundControlName.VB_MemberFlags = "200"
    BoundControlName = mBoundControlName
End Property


Public Property Let BoundProperty(nValue As String)
    If Ambient.UserMode Then
        Err.Raise 5676, "At run time use ChangeBoundControl Method"
    End If
    mBoundProperty = nValue
    PropertyChanged "BoundProperty"
End Property

Public Property Get BoundProperty() As String
    BoundProperty = mBoundProperty
End Property


Public Property Let BoundControlTag(nValue As String)
    If Ambient.UserMode Then
        Err.Raise 5676, "At run time use ChangeBoundControl Method"
    End If
    mBoundControlTag = nValue
    PropertyChanged "BoundControlTag"
End Property

Public Property Get BoundControlTag() As String
    BoundControlTag = mBoundControlTag
End Property


Public Sub ChangeBoundControl(nControlName As String, nPropertyName As String, Optional nTag As String, Optional nDoNotRestorePreviousText As Boolean)
    Dim iH As Variant
    Dim iP As Variant
    Dim iAux As Variant
    Dim iCtl As Control
    Dim iStr As String
    
    If mBoundControlName <> "" Then
        If tmrTextChanged.Enabled Then tmrTextChanged_Timer
        If UBound(mHistoryItems) > 0 Then
            On Error Resume Next
            iAux = mHistoriesCollection(mBoundControlName & mBoundControlTag)
            On Error GoTo 0
            If Not IsEmpty(iAux) Then
                mHistoriesCollection.Remove mBoundControlName & mBoundControlTag
                mHistoriesCollection.Remove mBoundControlName & mBoundControlTag & "_Pos"
                mHistoriesCollection.Remove mBoundControlName & mBoundControlTag & "_BCN"
                mHistoriesCollection.Remove mBoundControlName & mBoundControlTag & "_BCT"
                mHistoriesCollection.Remove mBoundControlName & mBoundControlTag & "_CT"
                mHistoriesCollection.Remove mBoundControlName & mBoundControlTag & "_TD"
            End If
            If UBound(mHistoryItems) > 0 Then
                mHistoriesCollection.Add mHistoryItems, mBoundControlName & mBoundControlTag
                mHistoriesCollection.Add CVar(mPosition), mBoundControlName & mBoundControlTag & "_Pos"
                mHistoriesCollection.Add CVar(mBoundControlName), mBoundControlName & mBoundControlTag & "_BCN"
                mHistoriesCollection.Add CVar(mBoundControlTag), mBoundControlName & mBoundControlTag & "_BCT"
                If mAutoAddItemEnabled Then
                    mHistoriesCollection.Add CVar(mCurrentText), mBoundControlName & mBoundControlTag & "_CT"
                Else
                    mHistoriesCollection.Add CVar(""), mBoundControlName & mBoundControlTag & "_CT"
                End If
                mHistoriesCollection.Add mTextsToDisplay, mBoundControlName & mBoundControlTag & "_TD"
            End If
        End If
    End If
    mBoundControlName = nControlName
    mBoundProperty = nPropertyName
    mBoundControlTag = nTag
    tmrCurrentText.Enabled = (mBoundControlName <> "") And (mBoundProperty <> "")
    
    ClearCurrentHistory
    On Error Resume Next
    iH = mHistoriesCollection(mBoundControlName & mBoundControlTag)
    On Error GoTo 0
    If Not IsEmpty(iH) Then
        iP = mHistoriesCollection(mBoundControlName & mBoundControlTag & "_Pos")
        mHistoryItems = iH
        mPosition = CLng(iP)
        iStr = CStr(mHistoriesCollection(mBoundControlName & mBoundControlTag & "_CT"))
'        If mAutoAddItemEnabled Then
            mCurrentText = iStr
'        End If
        mTextsToDisplay = mHistoriesCollection(mBoundControlName & mBoundControlTag & "_TD")
        ReDim mItemsTags(UBound(mTextsToDisplay))
        
        EnableDisableButtons
    Else
        LoadHistory
        If mAutoAddItemEnabled Then
            mCurrentText = ""
        End If
    End If
    UpdateTT
    
    If Not nDoNotRestorePreviousText Then
        'If mAutoAddItemEnabled And (mBoundControlTag <> "") Then
        If (mBoundControlTag <> "") Then
            Set iCtl = Parent.Controls(mBoundControlName)
            If Not iCtl Is Nothing Then
                CallByName iCtl, mBoundProperty, VbLet, mCurrentText
                On Error Resume Next
                iCtl.SelStart = 0
                iCtl.SelLength = Len(mCurrentText)
                iCtl.Refresh
                SetFocusTo iCtl
                On Error GoTo 0
            End If
        End If
    End If
End Sub

Private Sub UpdateTT()
    Dim iAuxTTEnd As String
    Dim iAuxTTStart As String
    Dim iAuxTTConf As String
    Dim iAuxItemText As String
    
    If mAutoToolTipText Then
        iAuxTTEnd = Trim$(mToolTipTextEnd)
        If iAuxTTEnd <> "" Then
            iAuxTTEnd = " " & iAuxTTEnd
        End If
        iAuxTTStart = RTrim(mToolTipTextStart)
        If iAuxTTStart <> "" Then
            iAuxTTStart = iAuxTTStart & " "
        End If
        iAuxTTConf = LTrim(mToolTipTextSelect)
        If iAuxTTConf <> "" Then
            iAuxTTConf = " " & iAuxTTConf
        End If
        
        iAuxItemText = ""
        If mPosition > 1 Then
            If (mCurrentText <> mHistoryItems(mPosition)) Then
                iAuxItemText = mTextsToDisplay(mPosition)
            Else
                iAuxItemText = mTextsToDisplay(mPosition - 1)
            End If
        Else
            If UBound(mHistoryItems) > 0 Then
                If mPosition = UBound(mHistoryItems) Then
                    If mCurrentText <> mHistoryItems(mPosition) Then
                        iAuxItemText = mTextsToDisplay(mPosition)
                    End If
                End If
            End If
        End If
        If iAuxItemText <> "" Then
            cmdHistoryBack.ToolTipText = iAuxTTStart & iAuxItemText & iAuxTTEnd
            If mEnableToConfigure Then
                If mToolTipTextSelect <> "" Then
                    If UBound(mHistoryItems) > 2 Then
                        cmdHistoryBack.ToolTipText = cmdHistoryBack.ToolTipText & iAuxTTConf
                    End If
                End If
            End If
        Else
            cmdHistoryBack.ToolTipText = ""
        End If
        If mPosition < (UBound(mHistoryItems)) Then
            cmdHistoryForward.ToolTipText = iAuxTTStart & mTextsToDisplay(mPosition + 1) & iAuxTTEnd
            If mEnableToConfigure Then
                If mToolTipTextSelect <> "" Then
                    If UBound(mHistoryItems) > 2 Then
                        cmdHistoryForward.ToolTipText = cmdHistoryForward.ToolTipText & iAuxTTConf
                    End If
                End If
            End If
        Else
            cmdHistoryForward.ToolTipText = ""
        End If
    End If
    RaiseEvent Updated
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "AutoToolTipText", mAutoToolTipText, True
    PropBag.WriteProperty "AutoShowConfig", mAutoShowConfig, True
    PropBag.WriteProperty "EnableToConfigure", mEnableToConfigure, True
    PropBag.WriteProperty "Enabled", mEnabled, True
    PropBag.WriteProperty "ToolTipTextStart", mToolTipTextStart, GetLocalizedString(efnGUIStr_History_ToolTipTextStart_Default)
    PropBag.WriteProperty "ToolTipTextEnd", mToolTipTextEnd, ""
    PropBag.WriteProperty "ToolTipTextSelect", mToolTipTextSelect, GetLocalizedString(efnGUIStr_History_ToolTipTextSelect_Default)
    PropBag.WriteProperty "AutoAddItemEnabled", mAutoAddItemEnabled, True
    PropBag.WriteProperty "Context", mContext, ""
    PropBag.WriteProperty "BoundControlName", mBoundControlName, ""
    PropBag.WriteProperty "BoundProperty", mBoundProperty, ""
    PropBag.WriteProperty "BoundControlTag", mBoundControlTag, ""
    PropBag.WriteProperty "BackColor", mBackColor, vbButtonFace
    PropBag.WriteProperty "ShowHistoryMenu", mShowHistoryMenu, True
    PropBag.WriteProperty "BackButtonToolTipText", cmdHistoryBack.ToolTipText, GetLocalizedString(efnGUIStr_History_BackButtonToolTipText_Default)
    PropBag.WriteProperty "ForwardButtonToolTipText", cmdHistoryForward.ToolTipText, GetLocalizedString(efnGUIStr_History_ForwardButtonToolTipText_Default)
    PropBag.WriteProperty "ButtonStyle", mButtonStyle, vxInstallShieldToolbar
End Sub

Private Sub SaveHistory()
    Dim iCH As Long
    Dim c As Long
    Dim iStr As String
    Dim iINi As Long
    
    If Not mHistoryLoaded Then Exit Sub
    
    If Not CBool(Val(GetSetting(AppNameForRegistry, "History", "Record", "-1"))) Then Exit Sub
    
    If tmrTextChanged.Enabled Then tmrTextChanged_Timer
    
    On Error Resume Next
    iCH = UBound(mHistoryItems)
    On Error GoTo 0
    
    If iCH > 0 Then
        iINi = 1
        If iCH > 30 Then
            iINi = iCH - 29
        End If
        For c = iINi To iCH
            iStr = iStr & Chr(124) & Replace(CStr(mHistoryItems(c)), Chr(124), "{_:_}") & Chr(124) & Replace(CStr(mTextsToDisplay(c)), Chr(124), "{_:_}")
        Next c
        SaveSetting AppNameForRegistry, "History", Base64Encode(mContext & Trim$(mBoundControlName) & Trim$(mBoundControlTag)), Base64Encode(iStr)
    Else
        On Error Resume Next
        DeleteSetting AppNameForRegistry, "History", Base64Encode(mContext & Trim$(mBoundControlName) & Trim$(mBoundControlTag))
    End If
End Sub

Private Sub LoadHistory()
    Dim c As Long
    Dim iStr As String
    Dim iAuxItems() As String
    Dim iLng As Long
    
    On Error GoTo TheExit:
    mHistoryLoaded = True
    iStr = Base64Decode(GetSetting(AppNameForRegistry, "History", Base64Encode(mContext & Trim$(mBoundControlName) & Trim$(mBoundControlTag)), ""))
    If InStr(iStr, Chr(124)) = 0 Then Exit Sub
    
    iAuxItems = Split(iStr, Chr(124))
    ReDim mHistoryItems(UBound(iAuxItems) / 2)
    ReDim mTextsToDisplay(UBound(mHistoryItems))
    ReDim mItemsTags(UBound(mHistoryItems))
    
    For c = 1 To UBound(iAuxItems) Step 2
        iLng = Round((c + 0.1) / 2)
        mHistoryItems(iLng) = Replace(iAuxItems(c), "{_:_}", Chr(124))
        mTextsToDisplay(iLng) = Replace(iAuxItems(c + 1), "{_:_}", Chr(124))
    Next c
    mPosition = UBound(mHistoryItems)
    EnableDisableButtons
    UpdateTT
    Exit Sub
    
TheExit:
    ReDim mHistoryItems(0)
    ReDim mTextsToDisplay(0)
    ReDim mItemsTags(0)
    
    mPosition = 0
    EnableDisableButtons
    UpdateTT
End Sub

Private Sub RaiseEventClick(nText As String)
    Dim iCtl As Control
    Dim iAuxHwnd As Long
    
    If (mBoundControlName <> "") And (mBoundProperty <> "") Then
        Set iCtl = Parent.Controls(mBoundControlName)
        If Not iCtl Is Nothing Then
            On Error Resume Next
            iAuxHwnd = iCtl.hWnd
            On Error GoTo 0
            CallByName iCtl, mBoundProperty, VbLet, nText
            If iAuxHwnd <> 0 Then
                If IsWindow(iAuxHwnd) = 0 Then
                    Exit Sub
                End If
            End If
            On Error Resume Next
            iCtl.SelStart = 0
            iCtl.SelLength = Len(nText)
            On Error GoTo 0
            SetFocusTo iCtl
            mCurrentText = nText
        End If
    End If
    RaiseEvent Click(nText)
End Sub


Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrPreprocess
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Dim iIndex As Long
    Dim iProcess As Boolean
    Dim iP1 As POINTAPI
    
    Select Case iMsg
        Case WM_HISTORYERASED
            ClearCurrentHistory
            Set mHistoriesCollection = New Collection
        Case WM_RBUTTONDOWN
            GetCursorPos iP1
            If (IsWindowVisible(mcboHistoryListHwnd) <> 0) And (WindowFromPoint(iP1.x, iP1.y) = mcboHistoryListHwnd) Then
                iIndex = SendMessage(hWnd, LB_ITEMFROMPOINT, 0&, ByVal lParam)
                iIndex = iIndex And &HFF
                If iIndex > -1 And iIndex < cboHistory.ListCount Then
                    mSelectedItemToDelete = iIndex
                    iProcess = True
                    If AutoShowConfig Then
                        If iIndex >= (cboHistory.ListCount - 2) Then
                            iProcess = False
                        End If
                    End If
                    If iProcess Then
                        On Error Resume Next
                        mnuDelete.Caption = GetLocalizedString(efnGUIStr_History_mnuDelete_Caption1) & " '" & cboHistory.List(iIndex) & "' " & GetLocalizedString(efnGUIStr_History_mnuDelete_Caption2)
                        On Error GoTo 0
                        mPopupShown = True
                        PopupMenu mnuPopup
                        mPopupShown = False
                    End If
                End If
            End If
        Case WM_UILANGCHANGED
            UILangChange wParam
        Case Else
    End Select
End Function

Private Sub UILangChange(nPrevLang As Long)
    If mToolTipTextSelect = GetLocalizedString(efnGUIStr_History_ToolTipTextSelect_Default, , nPrevLang) Then ToolTipTextSelect = GetLocalizedString(efnGUIStr_History_ToolTipTextSelect_Default)
    If mToolTipTextStart = GetLocalizedString(efnGUIStr_History_ToolTipTextStart_Default, , nPrevLang) Then ToolTipTextStart = GetLocalizedString(efnGUIStr_History_ToolTipTextStart_Default)
    If BackButtonToolTipText = GetLocalizedString(efnGUIStr_History_BackButtonToolTipText_Default, , nPrevLang) Then BackButtonToolTipText = GetLocalizedString(efnGUIStr_History_BackButtonToolTipText_Default)
    If ForwardButtonToolTipText = GetLocalizedString(efnGUIStr_History_ForwardButtonToolTipText_Default, , nPrevLang) Then ForwardButtonToolTipText = GetLocalizedString(efnGUIStr_History_ForwardButtonToolTipText_Default)
End Sub

Private Sub RemoveFromHistory(nItem As Long)
    Dim c As Long
    
    If (nItem < 1) Or (nItem > UBound(mHistoryItems)) Then Exit Sub
    
    For c = nItem + 1 To UBound(mHistoryItems)
        If c > 1 Then
            mHistoryItems(c - 1) = mHistoryItems(c)
            mTextsToDisplay(c - 1) = mTextsToDisplay(c)
            mItemsTags(c - 1) = mItemsTags(c)
        End If
    Next c
    ReDim Preserve mTextsToDisplay(UBound(mHistoryItems) - 1)
    ReDim Preserve mItemsTags(UBound(mHistoryItems) - 1)
    ReDim Preserve mHistoryItems(UBound(mHistoryItems) - 1)
    
    If mPosition > nItem Then
        mPosition = mPosition - 1
    End If
    If mPosition < 0 Then mPosition = 0
    If mPosition > UBound(mHistoryItems) Then mPosition = UBound(mHistoryItems)
    EnableDisableButtons

End Sub

Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Let AmbientDesignModeParent(nValue As Boolean)
    mAmbientDesignModeParent = nValue
End Property


Public Property Let ItemTag(nIndex As Long, nTag As String)
    mItemsTags(nIndex) = nTag
End Property

Public Property Get ItemTag(nIndex As Long) As String
    ItemTag = mItemsTags(nIndex)
End Property


Public Property Let ButtonStyle(nValue As vbExButtonStyleConstants)
    If nValue <> mButtonStyle Then
        mButtonStyle = nValue
        PropertyChanged "ButtonStyle"
        cmdHistoryBack.ButtonStyle = mButtonStyle
        cmdHistoryForward.ButtonStyle = mButtonStyle
    End If
End Property

Public Property Get ButtonStyle() As vbExButtonStyleConstants
    ButtonStyle = mButtonStyle
End Property

Private Sub PositionControls()
    cboHistory.Move 0, 0
    picCoveringCombo.Move 0, 0
    cmdHistoryBack.Move 0, 0, 270, 345
    UserControl.Height = cmdHistoryBack.Height
    cmdHistoryForward.Move cmdHistoryBack.Width + 15, cmdHistoryBack.Top, cmdHistoryBack.Width, cmdHistoryBack.Height
    UserControl.Width = cmdHistoryForward.Left + cmdHistoryForward.Width
    picCoveringCombo.Height = UserControl.ScaleHeight
End Sub

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
