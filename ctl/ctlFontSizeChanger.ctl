VERSION 5.00
Begin VB.UserControl FontSizeChanger 
   ClientHeight    =   408
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1068
   LockControls    =   -1  'True
   PropertyPages   =   "ctlFontSizeChanger.ctx":0000
   ScaleHeight     =   408
   ScaleWidth      =   1068
   ToolboxBitmap   =   "ctlFontSizeChanger.ctx":0029
   Begin vbExtra.ButtonEx btnMinus 
      Height          =   285
      Left            =   60
      TabIndex        =   4
      ToolTipText     =   "# Decrease font size"
      Top             =   30
      Width           =   225
      _ExtentX        =   402
      _ExtentY        =   508
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "-"
   End
   Begin vbExtra.ButtonEx btnPlus 
      Height          =   285
      Left            =   735
      TabIndex        =   5
      ToolTipText     =   "# Increase font size"
      Top             =   30
      Width           =   225
      _ExtentX        =   402
      _ExtentY        =   508
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "+"
   End
   Begin VB.Label lblA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E35500&
      Height          =   285
      Left            =   300
      TabIndex        =   0
      Top             =   60
      Width           =   165
   End
   Begin VB.Label lblA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E35500&
      Height          =   360
      Left            =   480
      TabIndex        =   2
      Top             =   0
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   360
      Index           =   0
      Left            =   525
      TabIndex        =   3
      Top             =   15
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   345
      TabIndex        =   1
      Top             =   75
      Width           =   165
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuDefaultFontSize 
         Caption         =   "# Set default value"
      End
   End
End
Attribute VB_Name = "FontSizeChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Private Const WM_UILANGCHANGED As Long = WM_USER + 12

Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Private mTTT As String
Private mMaxFontSize As Single
Private mMinFontSize As Single
Private mStep As Single
Private mFontSize As Single
Private mForeColor As Long
Private mBackColor As Long
Private mDefaultFontSize As Single
Private mBoundControlName As String

Public Event Change()
Public Event Click()

Private Const cMinFontSize_Default As Single = 8
Private Const cMaxFontSize_Default As Single = 16
Private Const cStep_Default As Single = 1
Private Const cFontSize_Default As Single = 10
Private Const cForeColor_Default As Long = &HE35500
Private Const cBackColor_Default As Long = vbButtonFace
Private mButtonStyle As vbExButtonStyleConstants
Private mUserControlHwnd As Long


Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Private Sub btnMinus_Click()
    mFontSize = mFontSize - mStep
    AdjustValues
    RaiseEvent Click
    RaiseEvent_Change
    PutToolTip
End Sub

Private Sub btnPlus_Click()
    mFontSize = mFontSize + mStep
    AdjustValues
    RaiseEvent Click
    RaiseEvent_Change
    PutToolTip
End Sub

Private Sub lblA2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    PopupMenu mnuPopup
End Sub

Private Sub mnuDefaultFontSize_Click()
    FontSize = mDefaultFontSize
End Sub

Private Sub UserControl_Initialize()
    mForeColor = cForeColor_Default
    mBackColor = cBackColor_Default
End Sub

Private Sub UserControl_InitProperties()
    LoadGUICaptions
    mMinFontSize = cMinFontSize_Default
    mMaxFontSize = cMaxFontSize_Default
    mStep = cStep_Default
    mFontSize = cFontSize_Default
    mForeColor = cForeColor_Default
    mBackColor = cBackColor_Default
    mButtonStyle = vxInstallShieldToolbar
    If Ambient.UserMode Then
        mUserControlHwnd = UserControl.hWnd
        AttachMessage Me, mUserControlHwnd, WM_UILANGCHANGED
        SetProp mUserControlHwnd, "FnExUI", 1
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    LoadGUICaptions
    mBoundControlName = PropBag.ReadProperty("BoundControlName", "")
    mMinFontSize = PropBag.ReadProperty("MinFontSize", cMinFontSize_Default)
    mMaxFontSize = PropBag.ReadProperty("MaxFontSize", cMaxFontSize_Default)
    mStep = PropBag.ReadProperty("Step", cStep_Default)
    mFontSize = PropBag.ReadProperty("FontSize", cFontSize_Default)
    mDefaultFontSize = mFontSize
    ForeColor = PropBag.ReadProperty("ForeColor", cForeColor_Default)
    BackColor = PropBag.ReadProperty("BackColor", cBackColor_Default)
    ButtonStyle = PropBag.ReadProperty("ButtonStyle", vxInstallShieldToolbar)
    AdjustValues
    If Ambient.UserMode Then
        mUserControlHwnd = UserControl.hWnd
        AttachMessage Me, mUserControlHwnd, WM_UILANGCHANGED
        SetProp mUserControlHwnd, "FnExUI", 1
    End If
End Sub


Public Property Let MaxFontSize(nValue As Single)
    If nValue <> mMaxFontSize Then
        mMaxFontSize = nValue
        PropertyChanged "MaxFontSize"
        AdjustValues
    End If
End Property

Public Property Get MaxFontSize() As Single
    MaxFontSize = mMaxFontSize
End Property


Public Property Let MinFontSize(nValue As Single)
    If nValue <> mMinFontSize Then
        mMinFontSize = nValue
        PropertyChanged "MinFontSize"
        AdjustValues
    End If
End Property

Public Property Get MinFontSize() As Single
    MinFontSize = mMinFontSize
End Property


Public Property Let Step(nValue As Single)
    If (nValue <= 0.01) Or (nValue > 10) Then Exit Property
    If nValue <> mStep Then
        mStep = nValue
        PropertyChanged "Step"
        AdjustValues
    End If
End Property

Public Property Get Step() As Single
    Step = mStep
End Property


Public Property Let FontSize(nValue As Single)
    Dim iSng As Single
    
    If nValue <> mFontSize Then
        iSng = nValue
        If iSng < mMinFontSize Then iSng = mMinFontSize
        If iSng > mMaxFontSize Then iSng = mMaxFontSize
        mFontSize = iSng
        PropertyChanged "FontSize"
        AdjustValues
        RaiseEvent_Change
        PutToolTip
    End If
End Property

Public Property Get FontSize() As Single
Attribute FontSize.VB_MemberFlags = "200"
    FontSize = mFontSize
End Property


Public Property Let ForeColor(nValue As OLE_COLOR)
    If nValue <> mForeColor Then
        mForeColor = nValue
        PropertyChanged "ForeColor"
        lblA1.ForeColor = mForeColor
        lblA2.ForeColor = mForeColor
    End If
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property


Public Property Let BackColor(nValue As OLE_COLOR)
    If nValue <> mBackColor Then
        mBackColor = nValue
        PropertyChanged "BackColor"
        UserControl.BackColor = mBackColor
        btnMinus.BackColor = mBackColor
        btnPlus.BackColor = mBackColor
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property


Private Sub UserControl_Resize()
    Dim iH As Long
    Dim iW As Long
    
    iH = UserControl.ScaleHeight
    iW = UserControl.ScaleWidth
    
    If (iH <> 370) Or (iW <> 1000) Then
        If (iH <> 370) Then
            iH = 370
        End If
        If (iW <> 1000) Then
            iW = 1000
        End If
        UserControl.Size iW, iH
    End If
End Sub

Private Sub UserControl_Show()
    PutToolTip
    If Ambient.UserMode Then
        If mBoundControlName <> "" Then
            SetBoundControlFont
        End If
    End If
End Sub

Private Sub PutToolTip()
    Dim iCtl As Control
    Static sLast As String
    Dim iFs As String
    
    If Ambient.UserMode Then
        If (Extender.ToolTipText <> "") And (Extender.ToolTipText <> sLast) Then
            mTTT = Extender.ToolTipText
        Else
            iFs = Format(mFontSize)
            mTTT = GetLocalizedString(efnGUIStr_FontSizeChanger_Extender_ToolTipText) & iFs & ")"
            Extender.ToolTipText = mTTT
        End If
    
        For Each iCtl In UserControl.Controls
            If TypeOf iCtl Is Label Then
                iCtl.ToolTipText = mTTT
            End If
        Next
        sLast = mTTT
    End If
End Sub

Private Sub UserControl_Terminate()
    If mUserControlHwnd <> 0 Then
        DetachMessage Me, mUserControlHwnd, WM_UILANGCHANGED
        RemoveProp mUserControlHwnd, "FnExUI"
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BoundControlName", mBoundControlName, ""
    PropBag.WriteProperty "MinFontSize", mMinFontSize, cMinFontSize_Default
    PropBag.WriteProperty "MaxFontSize", mMaxFontSize, cMaxFontSize_Default
    PropBag.WriteProperty "Step", mStep, cStep_Default
    PropBag.WriteProperty "FontSize", mFontSize, cFontSize_Default
    PropBag.WriteProperty "ForeColor", mForeColor, cForeColor_Default
    PropBag.WriteProperty "BackColor", mBackColor, cBackColor_Default
    PropBag.WriteProperty "ButtonStyle", mButtonStyle, vxInstallShieldToolbar
End Sub

Private Sub AdjustValues()
    If mMinFontSize < 3 Then
        mMinFontSize = 3
    End If
    If mMaxFontSize < (mMinFontSize + mStep * 3) Then
        mMaxFontSize = mMinFontSize + mStep * 3
    End If
    If mMaxFontSize > 100 Then
        mMaxFontSize = 100
    End If
    If mFontSize < mMinFontSize Then
        mFontSize = mMinFontSize
    End If
    If mFontSize > mMaxFontSize Then
        mFontSize = mMaxFontSize
    End If
    btnMinus.Enabled = mFontSize > (mMinFontSize + mStep * 0.99)
    btnPlus.Enabled = mFontSize < (mMaxFontSize - mStep * 0.99)
End Sub

Public Sub IncreaseFontSize()
    If btnPlus.Enabled Then
        btnPlus_Click
    End If
End Sub

Public Sub DecreaseFontSize()
    If btnMinus.Enabled Then
        btnMinus_Click
    End If
End Sub

Public Sub DecreaseFont()
    If btnMinus.Enabled Then
        btnMinus_Click
    End If
End Sub

Public Sub IncreaseFont()
    If btnPlus.Enabled Then
        btnPlus_Click
    End If
End Sub

Public Property Get CanIncrease() As Boolean
    CanIncrease = btnPlus.Enabled
End Property

Public Property Get CanDecrease() As Boolean
    CanDecrease = btnMinus.Enabled
End Property


Public Property Let ButtonStyle(nValue As vbExButtonStyleConstants)
    If nValue <> mButtonStyle Then
        mButtonStyle = nValue
        PropertyChanged "ButtonStyle"
        btnMinus.ButtonStyle = mButtonStyle
        btnPlus.ButtonStyle = mButtonStyle
    End If
End Property

Public Property Get ButtonStyle() As vbExButtonStyleConstants
    ButtonStyle = mButtonStyle
End Property

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrPostProcess
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Select Case iMsg
        Case WM_UILANGCHANGED
            UILangChange wParam
    End Select
End Function

Private Sub UILangChange(nPrevLang As Long)
    If mnuDefaultFontSize.Caption = GetLocalizedString(efnGUIStr_FontSizeChanger_mnuDefaultFontSize_Caption, , nPrevLang) Then mnuDefaultFontSize.Caption = GetLocalizedString(efnGUIStr_FontSizeChanger_mnuDefaultFontSize_Caption)
    If btnMinus.ToolTipText = GetLocalizedString(efnGUIStr_FontSizeChanger_btnMinus_ToolTipText, , nPrevLang) Then btnMinus.ToolTipText = GetLocalizedString(efnGUIStr_FontSizeChanger_btnMinus_ToolTipText)
    If btnPlus.ToolTipText = GetLocalizedString(efnGUIStr_FontSizeChanger_btnPlus_ToolTipText, , nPrevLang) Then btnPlus.ToolTipText = GetLocalizedString(efnGUIStr_FontSizeChanger_btnPlus_ToolTipText)
End Sub

Private Sub LoadGUICaptions()
    mnuDefaultFontSize.Caption = GetLocalizedString(efnGUIStr_FontSizeChanger_mnuDefaultFontSize_Caption)
    btnMinus.ToolTipText = GetLocalizedString(efnGUIStr_FontSizeChanger_btnMinus_ToolTipText)
    btnPlus.ToolTipText = GetLocalizedString(efnGUIStr_FontSizeChanger_btnPlus_ToolTipText)
End Sub

Public Property Let BoundControlName(ByVal nValue As String)
    nValue = Trim$(nValue)
    If mBoundControlName <> nValue Then
        mBoundControlName = nValue
        If Ambient.UserMode Then
            If mBoundControlName <> "" Then
                SetBoundControlFont
            End If
        End If
        PropertyChanged "BoundControlName"
    End If
End Property

Public Property Get BoundControlName() As String
    BoundControlName = mBoundControlName
End Property

Private Sub RaiseEvent_Change()
    Static sLast As Single
    
    If Ambient.UserMode Then
        If mFontSize <> sLast Then
            If mBoundControlName <> "" Then
                SetBoundControlFont
            End If
            RaiseEvent Change
            sLast = mFontSize
        End If
    End If
End Sub

Private Sub SetBoundControlFont()
    Dim iCtl As Object
    'Debug.Print mFontSize
    On Error Resume Next
    Set iCtl = Parent.Controls(mBoundControlName)
    If Not iCtl Is Nothing Then
        iCtl.Font.Size = mFontSize
    End If
End Sub

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

