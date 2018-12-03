VERSION 5.00
Begin VB.UserControl DateEnter 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "ctlDateEnter.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ctlDateEnter.ctx":0044
   Begin VB.Timer tmrSetFocusDTPickerDroppedDown 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   144
      Top             =   2556
   End
   Begin VB.Timer tmrValidate 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   144
      Top             =   2196
   End
   Begin VB.Timer tmrSetFocus2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   144
      Top             =   1836
   End
   Begin VB.Timer tmrSetFocusToMasked 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   144
      Top             =   1476
   End
   Begin VB.TextBox txtMasked 
      BorderStyle     =   0  'None
      Height          =   408
      Left            =   2340
      TabIndex        =   2
      Top             =   936
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   390
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   990
      Visible         =   0   'False
      Width           =   1485
   End
   Begin vbExtra.DTPickerEx DTPicker1 
      Height          =   480
      Left            =   1800
      TabIndex        =   3
      Top             =   144
      Width           =   2316
      _ExtentX        =   4085
      _ExtentY        =   847
      CheckBox        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   15
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1380
   End
End
Attribute VB_Name = "DateEnter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Private Type DATETIMEPICKERINFO
    cbSize As Long
    rcCheck As RECT
    stateCheck As Long
    rcButton As RECT
    stateButton As Long
    hwndEdit As Long
    hwndUD As Long
    hwndDropDown As Long
End Type

Private Const DTM_FIRST As Long = &H1000
Private Const DTM_GETDATETIMEPICKERINFO As Long = DTM_FIRST + 14

Private Const EM_REPLACESEL As Long = &HC2
Private Const WM_UILANGCHANGED As Long = WM_USER + 12
Private Const WM_KILLFOCUS As Long = &H8&
Private Const WM_SETFOCUS As Long = &H7

Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetLocaleInfo Lib "Kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetUserDefaultLCID% Lib "Kernel32" ()
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Const SM_CXVSCROLL = 2
Private Const LOCALE_SSHORTDATE As Long = &H1F

Public Event Change()
Public Event TextChange()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event DropDown()
Public Event CloseUp()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event Click()
Public Event DblClick()

Private mRemoveBorder As New cRemoveBorder

Public Enum vbExDateSeparatorConstants
    [Use from system] = 0
    ["/"] = 1
    ["-"] = 2
    ["."] = 3
    [" " (space)] = 4
End Enum

Public Enum vbExAppearanceConstants
    ccFlat = 0
    cc3d = 1
End Enum

Public Enum vbExDateFormatConstants
    [System date format] = 0
    [System but 4 digits year] = 1
    [System but 2 digits year] = 2
    [dd/MM/yyyy] = 3
    [dd/MM/yy] = 4
    [MM/dd/yyyy] = 5
    [MM/dd/yy] = 6
    [yyyy/MM/dd] = 7
    [yy/MM/dd] = 8
End Enum

Private Const cDTPickerMinDate                    As Date = "01/01/1601"
Private Const cDTPickerMaxDate                    As Date = "31/12/9999"

Private mDateFormat As vbExDateFormatConstants
Private mDateFormatStr As String
Private mEmptyDate As String

Private mValue As Variant
Private mFont As StdFont
Private mAutoValidate As Boolean
Private mDateSeparator As vbExDateSeparatorConstants

Private mEnabled As Boolean
Private mHelpContextID As Integer
Private mMouseIcon As StdPicture
Private mMousePointer As Integer
Private mToolTipTextStart As String
Private mToolTipTextEnd As String
Private mWhatsThisHelpID As Integer
Private mAppearance As Long

Private mTextBackColor As Long
Private mCalendarBackColor As Long
Private mCalendarForeColor As Long
Private mCalendarTitleBackColor As Long
Private mCalendarTitleForeColor As Long
Private mCalendarTrailingForeColor As Long

Private mMinDate As Date
Private mMaxDate As Date

Private mFlatPending As Boolean
Private mUserControlHwnd As Long
Private mMaskedHwnd As Long

Private mVerticalScrollbarWidth As Long
Private mDontDoGotFocus As Boolean
Private mValue_Prev As Variant

Private mMask As String
Private mEmptyMask As String
Private mTextValue As String
Private mInsideKeyPress As Boolean
Private mLenMask As Long
Private mValidationErrors As Long
Private mDateSeparatorChar As String
Private mNeedValidation As Boolean
Private mSetFocus2Control As Variant
Private mOnFocus As Boolean
Private mInsideValidate1 As Boolean
Private mSettingIncompleteValue As Boolean
Private mSettingNullDateFromDTPicker1ChangeEvent As Boolean

Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1

Public Property Let DateSeparator(nValue As vbExDateSeparatorConstants)
    If nValue <> mDateSeparator Then
        mDateSeparator = nValue
        PropertyChanged "DateSeparator"
        SetDateFormat
    End If
End Property

Public Property Get DateSeparator() As vbExDateSeparatorConstants
    DateSeparator = mDateSeparator
End Property

Private Property Let Mask(nValue As String)
    If nValue <> mMask Then
        mMask = nValue
        mEmptyMask = Replace(mMask, "#", "_")
        If txtMasked.Text = "" Then
            txtMasked.Text = mEmptyMask
        End If
        mLenMask = Len(mEmptyMask)
    End If
End Property

Private Property Get Mask() As String
    Mask = mMask
End Property

Private Sub Form_Load()
    Mask = "##/##/####"
End Sub

Private Sub DTPicker1_CalendarClick()
    Value = DTPicker1.Value
End Sub

Private Sub DTPicker1_Click()
    RaiseEvent Click
End Sub

Private Sub DTPicker1_CloseUp()
    RaiseEvent CloseUp
    If txtMasked.Visible Then
        If mOnFocus Then tmrSetFocusToMasked.Enabled = True
    Else
        SetFocusTo2 DTPicker1
    End If
End Sub

Private Sub DTPicker1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub DTPicker1_DropDown()
    tmrSetFocusToMasked.Enabled = False
    tmrSetFocus2.Enabled = False
    tmrValidate.Enabled = False
'    If mNeedValidation Then
'        If Not Validate Then
'            DTPicker1.DroppedDown = False
'            Exit Sub
'        End If
'    End If
    RaiseEvent DropDown
End Sub

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub DTPicker1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrSetFocusToMasked.Enabled = False
    tmrSetFocus2.Enabled = False
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub DTPicker1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub DTPicker1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub mForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mAutoValidate Then
        If txtMasked.Text <> mEmptyDate Then
            If mNeedValidation Then
                Dim iUserReEntering As Boolean
                Validate1 True, , , iUserReEntering
                If iUserReEntering Then
                    Cancel = 1
                End If
            End If
        End If
    End If
End Sub

Private Sub tmrSetFocus2_Timer()
    tmrSetFocus2.Enabled = False
    If Not mOnFocus Then Exit Sub
    If DTPicker1.DroppedDown Then Exit Sub
    SetFocusTo mSetFocus2Control
    Set mSetFocus2Control = Nothing
End Sub

Private Sub tmrSetFocusDTPickerDroppedDown_Timer()
    tmrSetFocusDTPickerDroppedDown.Enabled = False
    SetFocusAPI DTPicker1.hWnd
End Sub

Private Sub tmrSetFocusToMasked_Timer()
    tmrSetFocusToMasked.Enabled = False
    'If Not mOnFocus Then Exit Sub
    If DTPicker1.DroppedDown Then Exit Sub
    If IsWindowVisibleOnScreen(UserControl.hWnd, True) And UserControl.Enabled Then
        'Debug.Print "tmrSetFocusToMasked_Timer " & Ambient.DisplayName
        SetFocusTo txtMasked
    End If
End Sub

Private Sub tmrValidate_Timer()
    tmrValidate.Enabled = False
    If DTPicker1.DroppedDown Then Exit Sub
    If mNeedValidation Then
        If IsWindowVisibleOnScreen(UserControl.hWnd, True) And UserControl.Enabled Then
            SetFocusTo txtMasked
            Validate1 True
        End If
    End If
'    Debug.Print "tmrValidate " & Ambient.DisplayName
End Sub

Private Sub txtMasked_Change()
    Dim iStr As String
    Dim iStr2 As String
    Dim c As Long
    Dim iChr As String
    Static sInside As Boolean
    Dim iSS As Long
    
    Static sLastDate As Variant
    Dim iDateAnt As Variant
    
    If mSettingNullDateFromDTPicker1ChangeEvent Then Exit Sub
    If mInsideKeyPress Then Exit Sub
    If sInside Then Exit Sub
    sInside = True
    
    If txtMasked.SelStart > 0 Then
        If Mid(txtMasked.Text, txtMasked.SelStart + 1, 1) = mDateSeparatorChar Then
            txtMasked.SelStart = txtMasked.SelStart + 1
        End If
    End If
    iStr = txtMasked.Text
    iStr = Replace(iStr, "_", "")
    iStr = Replace(iStr, mDateSeparatorChar, "")
    If (Not IsNumeric(iStr)) Then
        For c = 1 To Len(iStr)
            iChr = Mid(iStr, c, 1)
            If IsNumeric(iChr) Then
                iStr2 = iStr2 & iChr
            End If
        Next c
        If iStr2 = "" Then
            txtMasked.Text = mEmptyMask
            txtMasked.SelStart = 0
        Else
            SaveTextValue iStr2
            txtMasked.Text = GetTextValue
            SetCursorToLastNumber
        End If
    ElseIf Len(txtMasked.Text) > mLenMask Then
        iSS = txtMasked.SelStart
        For c = 1 To Len(iStr)
            iChr = Mid(iStr, c, 1)
            If IsNumeric(iChr) Then
                iStr2 = iStr2 & iChr
            End If
        Next c
        If iStr2 = "" Then
            txtMasked.Text = mEmptyMask
            txtMasked.SelStart = 0
        Else
            SaveTextValue iStr2
            txtMasked.Text = GetTextValue
            txtMasked.SelStart = iSS
        End If
    End If

    iDateAnt = sLastDate
    If IsDate(txtMasked.Text) Then
        
        If IsValidDate(Replace(txtMasked.Text, mDateSeparatorChar, "/")) Then
            On Error Resume Next
            Err.Clear
            DTPicker1.Value = CDate(Replace(txtMasked.Text, mDateSeparatorChar, "/"))
            If Err.Number = 0 Then
                SetFocusTo2 DTPicker1
                txtMasked.Visible = False
                DTPicker1.TabStop = True
            End If
            sLastDate = CDate(Replace(txtMasked.Text, mDateSeparatorChar, "/"))
        End If
    Else
        SetIncompleteDateInDTPicker1
        sLastDate = Empty
    End If
    If Ambient.UserMode Then
        RaiseEvent_TextChange
        If iDateAnt <> sLastDate Then
            RaiseEvent Change
        End If
    End If
    sInside = False
End Sub

Private Sub SetCursorToLastNumber()
    Dim c As Long
    Dim iChr As String
    
    For c = 1 To Len(txtMasked.Text)
        iChr = Mid(txtMasked.Text, c, 1)
        If Not (IsNumeric(iChr) Or (iChr = mDateSeparatorChar)) Then
            txtMasked.SelStart = c - 1
            Exit Sub
        End If
    Next c
    txtMasked.SelStart = Len(txtMasked.Text)
End Sub

Private Sub txtMasked_Click()
    RaiseEvent Click
End Sub

Private Sub txtMasked_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub txtMasked_GotFocus()
    mOnFocus = True
    If mDontDoGotFocus Then
        mDontDoGotFocus = False
        Exit Sub
    End If
    tmrValidate.Enabled = False
    txtMasked.ToolTipText = ""
    If txtMasked.Text <> mEmptyMask Then
        SetCursorToLastNumber
    Else
        txtMasked.SelStart = 0
    End If
    txtMasked.SelLength = 0
    'SetFocusTo2 txtMasked
End Sub

Private Sub txtMasked_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iStr As String
    Dim iSS As Long
    
    If KeyCode = vbKeyDelete Then
        If txtMasked.SelLength = 0 Then
            txtMasked.SelLength = 1
        End If
        iSS = txtMasked.SelStart
        If iSS < mLenMask Then
            iStr = txtMasked.Text
            Mid$(iStr, iSS + 1, txtMasked.SelLength) = Mid$(mEmptyMask, iSS + 1, txtMasked.SelLength)
            SaveTextValue iStr
            txtMasked.Text = GetTextValue
            txtMasked.SelStart = iSS
            KeyCode = 0
        End If
    ElseIf KeyCode = vbKeyDown Then
        tmrSetFocusDTPickerDroppedDown.Enabled = True
        SetIncompleteDateInDTPicker1
        DTPicker1.DroppedDown = True
        KeyCode = 0
    End If
    
    If KeyCode <> 0 Then
        If GetFocus = txtMasked.hWnd Then
            If (Shift And vbCtrlMask) = vbCtrlMask Then
                If KeyCode = vbKeyDelete Then
                        Value = Null
                End If
            End If
            RaiseEvent KeyDown(KeyCode, Shift)
        End If
    End If
End Sub

Private Sub SaveTextValue(Optional nText As String)
    If nText = "" Then
        mTextValue = txtMasked.Text
    Else
        mTextValue = nText
    End If
    mTextValue = Replace(mTextValue, "_", "")
    mTextValue = Replace(mTextValue, mDateSeparatorChar, "")
End Sub

Private Function GetTextValue() As String
    Dim iChr As String
    Dim c As Long
    Dim c2 As Long
    
    GetTextValue = mEmptyMask

    c2 = 0
    For c = 1 To Len(mTextValue)
        iChr = Mid(mTextValue, c, 1)
        c2 = c2 + 1
        If Mid(GetTextValue, c2, 1) = mDateSeparatorChar Then
            c2 = c2 + 1
        End If
        If c2 > mLenMask Then Exit For
        Mid(GetTextValue, c2, 1) = iChr
    Next c

End Function

Private Sub txtMasked_KeyPress(KeyAscii As Integer)
    Dim iSS As Long
    Dim iStr As String
    
    mInsideKeyPress = True
    Select Case KeyAscii
        Case vbKeyReturn, vbKeyEscape
            '
        Case 48 To 57
            If txtMasked.SelStart = mLenMask Then
                SaveTextValue txtMasked.Text
                txtMasked.Text = GetTextValue
                SetCursorToLastNumber
                If txtMasked.SelStart = mLenMask Then
                    KeyAscii = 0
                End If
            ElseIf txtMasked.SelLength = 0 Then
                SaveTextValue
                If Len(mTextValue) + 2 >= mLenMask Then
                    KeyAscii = 0
                Else
                    txtMasked.SelLength = 1
                    If txtMasked.SelText = "_" Then
                        txtMasked.SelText = ""
                    ElseIf txtMasked.SelText = mDateSeparatorChar Then
                        txtMasked.SelStart = txtMasked.SelStart + 1
                        Exit Sub
                    Else
                        iSS = txtMasked.SelStart
                        iStr = txtMasked.Text
                        SaveTextValue iStr
                        txtMasked.Text = GetTextValue
                        txtMasked.SelStart = iSS
                    End If
                End If
            Else
                If txtMasked.SelText <> "" Then
                    iSS = txtMasked.SelStart
                    iStr = txtMasked.Text
                    Mid$(iStr, iSS + 1, txtMasked.SelLength) = Mid$(mEmptyMask, iSS + 1, txtMasked.SelLength)
                    If Mid(iStr, iSS + 2, 1) = mDateSeparatorChar Then
                        iStr = Left(iStr, iSS) & mDateSeparatorChar & Mid(iStr, iSS + 3)
                        iSS = iSS
                    Else
                        iStr = Left(iStr, iSS + 1) & Mid(iStr, iSS + 3)
                    End If
                    txtMasked.Text = iStr
                    txtMasked.SelStart = iSS
                    txtMasked.SelStart = iSS
                End If
            End If
        Case vbKeyBack
            If txtMasked.SelStart > 0 Then
                If txtMasked.SelLength = 0 Then
                    If txtMasked.SelStart = mLenMask Then
                        txtMasked.SelStart = txtMasked.SelStart - 1
                        iSS = txtMasked.SelStart
                    Else
                        iSS = txtMasked.SelStart - 1
                    End If
                    txtMasked.SelLength = 1
                Else
                    iSS = txtMasked.SelStart
                End If
                iStr = txtMasked.Text
                Mid$(iStr, iSS + 1, txtMasked.SelLength) = Mid$(mEmptyMask, iSS + 1, txtMasked.SelLength)
                SaveTextValue iStr
                txtMasked.Text = GetTextValue
                iStr = txtMasked.Text
                If iSS > 0 Then
                    If Mid(iStr, iSS, 1) = mDateSeparatorChar Then
                        iSS = iSS - 1
                    End If
                End If
                txtMasked.SelStart = iSS
                KeyAscii = 0
            End If
        Case 22, 3, 24
            ' Ctrl+C, Ctrl+V, Ctrl+X
        Case Else
            KeyAscii = 0
    End Select
    
    mInsideKeyPress = False

    If KeyAscii <> 0 Then
        RaiseEvent KeyPress(KeyAscii)
        If KeyAscii = vbKeyReturn Then
            Validate1 True, True
            KeyAscii = 0
        ElseIf KeyAscii = vbKeyEscape Then
            If Not IsNull(mValue_Prev) Then
                Value = mValue_Prev
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub DTPicker1_Change()
    If mSettingIncompleteValue Then Exit Sub
    If IsDate(DTPicker1.Value) Then
        txtMasked.Text = Format(CDate(DTPicker1.Value), mDateFormatStr)
        If (mValue <> DTPicker1.Value) Or (IsNull(mValue) <> IsNull(DTPicker1.Value)) Then
            Value = DTPicker1.Value
        End If
    Else
        mSettingNullDateFromDTPicker1ChangeEvent = True
        Value = Null
        mSettingNullDateFromDTPicker1ChangeEvent = False
    End If
End Sub

Private Sub DTPicker1_GotFocus()
    mOnFocus = True
    If txtMasked.Visible Then
        SetFocusTo2 txtMasked
    End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtMasked_KeyUp(KeyCode As Integer, Shift As Integer)
    mInsideKeyPress = False
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtMasked_LostFocus()
    txtMasked.ToolTipText = mToolTipTextStart & " " & LCase$(VBA.Replace(mDateFormatStr, "y", "a")) & " " & mToolTipTextEnd
End Sub

Private Sub txtMasked_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub txtMasked_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GetFocus <> txtMasked.hWnd Then
        If txtMasked.ToolTipText = "" Then
            On Error Resume Next
            txtMasked.ToolTipText = mToolTipTextStart & " " & LCase$(VBA.Replace(mDateFormatStr, "y", "a")) & " " & mToolTipTextEnd
            On Error GoTo 0
        End If
    Else
        txtMasked.ToolTipText = ""
    End If
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub txtMasked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_EnterFocus()
    mOnFocus = True
    If txtMasked.Visible Then
        SetFocusTo2 txtMasked
    Else
        SetFocusTo2 DTPicker1
    End If
End Sub

Private Sub UserControl_ExitFocus()
    If mAutoValidate Then
        If txtMasked.Text <> mEmptyDate Then
            If mNeedValidation Then
                Validate1 True
            End If
        End If
    End If
    mOnFocus = False
End Sub

Private Sub UserControl_GotFocus()
    mOnFocus = True
    If Not IsNull(Value) Then
        SetFocusTo2 DTPicker1
    Else
        SetFocusTo2 txtMasked
    End If
End Sub

Private Sub UserControl_Initialize()
    mVerticalScrollbarWidth = GetSystemMetrics(SM_CXVSCROLL)
End Sub

Private Sub UserControl_InitProperties()
    mTextBackColor = vbWindowBackground
    mCalendarBackColor = vbWindowBackground
    mCalendarForeColor = vbButtonText
    mCalendarTitleBackColor = vbActiveTitleBar
    mCalendarTitleForeColor = vbActiveTitleBarText
    mCalendarTrailingForeColor = vbGrayText
    mAppearance = cc3d
    
    mMinDate = cDTPickerMinDate
    mMaxDate = cDTPickerMaxDate
    
    mMousePointer = cc2Default
    mEnabled = True
    
    Set Font = Ambient.Font
    Value = Null
    SetDateFormat
    mAutoValidate = True
    mDateSeparator = [Use from system]
    mToolTipTextStart = GetLocalizedString(efnGUIStr_DateEnter_ToolTipTextStart_Default)
    mToolTipTextEnd = GetLocalizedString(efnGUIStr_DateEnter_ToolTipTextEnd_Default)
    UserControl.Size 1200, 300
    
    SetProperties
    
    If Ambient.UserMode Then
        mUserControlHwnd = UserControl.hWnd
        mMaskedHwnd = txtMasked.hWnd
        AttachMessage Me, mUserControlHwnd, WM_UILANGCHANGED
        AttachMessage Me, mMaskedHwnd, WM_KILLFOCUS
        AttachMessage Me, mMaskedHwnd, WM_SETFOCUS
        SetProp mUserControlHwnd, "FnExUI", 1
        If TypeOf Parent Is Form Then
            Set mForm = Parent
        End If
    End If
    
End Sub

Private Sub UserControl_Resize()
    Dim iAuxSng As Single
    Dim iDTPInfo As DATETIMEPICKERINFO
    Dim iTop As Long
    Dim iLeft As Long
    Dim iHeight As Long
    
    If UserControl.Width < 800 Then UserControl.Width = 800
    
    On Error Resume Next
    iAuxSng = Combo1.FontSize
    Combo1.FontSize = iAuxSng + 2
    Combo1.FontSize = iAuxSng
    
    If Abs(UserControl.Height - Combo1.Height) > (Screen.TwipsPerPixelY / 2 + 1) Then
        UserControl.Height = Combo1.Height
        Exit Sub
    End If
    DTPicker1.Left = 0
    DTPicker1.Top = 0
    DTPicker1.Width = UserControl.ScaleWidth
    DTPicker1.Height = UserControl.ScaleHeight
    iDTPInfo.cbSize = Len(iDTPInfo)
    SendMessage DTPicker1.hWnd, DTM_GETDATETIMEPICKERINFO, 0&, VarPtr(iDTPInfo)
    iTop = 3
    iLeft = 3
    
    If IsWindowsVistaOrMore Then
        If iDTPInfo.rcCheck.Left < iLeft Then iLeft = iDTPInfo.rcCheck.Left
        If iDTPInfo.rcCheck.Top < iTop Then iTop = iDTPInfo.rcCheck.Top
    End If
    
    txtMasked.Left = iLeft * Screen.TwipsPerPixelX
    txtMasked.Top = iTop * Screen.TwipsPerPixelY
    
    iHeight = UserControl.ScaleHeight - 5 * Screen.TwipsPerPixelY
    If IsWindowsVistaOrMore Then
        If iHeight < (iDTPInfo.rcCheck.Bottom * Screen.TwipsPerPixelY - txtMasked.Top) Then iHeight = iDTPInfo.rcCheck.Bottom * Screen.TwipsPerPixelY - txtMasked.Top
    End If
    
    txtMasked.Height = iHeight
    If IsWindowsVistaOrMore Then
        txtMasked.Width = UserControl.ScaleWidth - (iDTPInfo.rcButton.Right - iDTPInfo.rcButton.Left + 2) * Screen.TwipsPerPixelX
    Else
        txtMasked.Width = UserControl.ScaleWidth - (6.49 + GetSystemMetrics(SM_CXVSCROLL)) * Screen.TwipsPerPixelX
    End If
    
    lblBorder.Move Screen.TwipsPerPixelX, 0, UserControl.ScaleWidth - (3 + mVerticalScrollbarWidth) * Screen.TwipsPerPixelX, UserControl.ScaleHeight - Screen.TwipsPerPixelY
End Sub


Public Property Let Value(nDate As Variant)
    
    If IsDate(nDate) Then
        mValue = CDate(CLng(CDate(nDate)))
    ElseIf IsNumeric(nDate) Then
        mValue = CDate(CLng(nDate))
    Else
        mValue = Null
    End If
    If IsNull(mValue) Or (Not IsValidDate(mValue)) Then
        If mSettingNullDateFromDTPicker1ChangeEvent Then
            txtMasked.SelStart = 0
            txtMasked.SelLength = Len(txtMasked.Text)
            SendMessage txtMasked.hWnd, EM_REPLACESEL, ByVal 1&, ByVal mEmptyDate
        Else
            txtMasked.Text = mEmptyDate
        End If
        txtMasked.Visible = True
        DTPicker1.TabStop = False
        txtMasked.Refresh
        If mOnFocus Then tmrSetFocusToMasked.Enabled = True
    Else
        If txtMasked.Visible Then
            txtMasked.Visible = False
            DTPicker1.TabStop = True
        End If
        If IsNull(DTPicker1.Value) Or (DTPicker1.Value <> CDate(mValue)) Then
            DTPicker1.Value = CDate(mValue)
        End If
        txtMasked.Text = Format(CDate(DTPicker1.Value), mDateFormatStr)
        SetFocusTo2 DTPicker1
        If (mValue <> mValue_Prev) Or (IsNull(mValue) <> IsNull(mValue_Prev)) Then
            mValue_Prev = mValue
        End If
    End If
    PropertyChanged "Value"

End Property

Public Property Get Value() As Variant
Attribute Value.VB_MemberFlags = "200"
    Value = mValue
End Property


Public Property Get AutoValidate() As Boolean
    AutoValidate = mAutoValidate
End Property

Public Property Let AutoValidate(nValue As Boolean)
    If nValue <> mAutoValidate Then
        mAutoValidate = nValue
        PropertyChanged ("AutoValidate")
    End If
End Property
    
    
Private Sub SetDateFormat()
    Dim iRet1 As Long
    Dim lpLCDataVar As String
    Dim Locale As Long
    Dim iAuxSelStart As Long
    Dim c As Long
    Dim iYCount As Long
    Dim iUseTwoDigitsDateFormat As Boolean
    Dim iDateFormatPattern As String
    
    If mDateFormat = [dd/MM/yy] Then
        mDateFormatStr = "dd/MM/yy"
    ElseIf mDateFormat = [dd/MM/yyyy] Then
        mDateFormatStr = "dd/MM/yyyy"
    ElseIf mDateFormat = [MM/dd/yy] Then
        mDateFormatStr = "MM/dd/yy"
    ElseIf mDateFormat = [MM/dd/yyyy] Then
        mDateFormatStr = "MM/dd/yyyy"
    ElseIf mDateFormat = [yy/MM/dd] Then
        mDateFormatStr = "yy/MM/dd"
    ElseIf mDateFormat = [yyyy/MM/dd] Then
        mDateFormatStr = "yyyy/MM/dd"
    Else 'If (mDateFormat = [System date format]) Or (mDateFormat = [System but 4 digits year]) Or (mDateFormat = [System but 2 digits year]) Then
        Locale = GetUserDefaultLCID
        iRet1 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, lpLCDataVar, 0)
        mDateFormatStr = String$(iRet1, 0)
        GetLocaleInfo Locale, LOCALE_SSHORTDATE, mDateFormatStr, iRet1
        mDateFormatStr = Left$(mDateFormatStr, InStr(mDateFormatStr, Chr(0)) - 1)
        mDateFormatStr = mDateFormatStr
    End If
    
    If mDateSeparator = ["/"] Then
        mDateSeparatorChar = "/"
    ElseIf mDateSeparator = ["-"] Then
        mDateSeparatorChar = "-"
    ElseIf mDateSeparator = [" " (space)] Then
        mDateSeparatorChar = " "
    ElseIf mDateSeparator = ["."] Then
        mDateSeparatorChar = "."
    ElseIf mDateSeparator = [Use from system] Then
        If InStr(mDateFormatStr, "/") > 0 Then
            mDateSeparatorChar = "/"
        ElseIf InStr(mDateFormatStr, "-") > 0 Then
            mDateSeparatorChar = "-"
        Else
            mDateSeparatorChar = ""
            For c = 1 To Len(mDateFormatStr)
                Select Case Mid(mDateFormatStr, c, 1)
                    Case "d", "M", "y"
                        '
                    Case Else
                        mDateSeparatorChar = Mid(mDateFormatStr, c, 1)
                        Exit For
                End Select
            Next c
            If mDateSeparatorChar = "" Then
                mDateSeparatorChar = "/"
            End If
        End If
    End If
    mDateFormatStr = Replace(mDateFormatStr, "m", "M")
    For c = 1 To Len(mDateFormatStr)
        Select Case Mid(mDateFormatStr, c, 1)
            Case "d", "M", "y"
                '
            Case Else
                Mid(mDateFormatStr, c, 1) = mDateSeparatorChar
        End Select
    Next c
    
    If mDateFormat = [System but 2 digits year] Then
        iUseTwoDigitsDateFormat = True
    ElseIf mDateFormat = [System but 4 digits year] Then
        iUseTwoDigitsDateFormat = False
    Else ' mDigitsYear = [Use system configuration]
        iYCount = ChrCount(mDateFormatStr, AscW("y"))
        If iYCount > 2 Then
            iUseTwoDigitsDateFormat = False
        Else
            iUseTwoDigitsDateFormat = True
        End If
    End If
    Do Until InStr(mDateFormatStr, "dd") = 0
        mDateFormatStr = Replace(mDateFormatStr, "dd", "d")
    Loop
    If InStr(mDateFormatStr, "d") = 0 Then
        mDateFormatStr = "d" & mDateSeparatorChar & mDateFormatStr
    End If
    If InStr(mDateFormatStr, "M") = 0 Then
        mDateFormatStr = "M" & mDateSeparatorChar & mDateFormatStr
    End If
    If InStr(mDateFormatStr, "d") = 0 Then
        mDateFormatStr = "d" & mDateSeparatorChar & mDateFormatStr
    End If
    Do Until InStr(mDateFormatStr, mDateSeparatorChar & mDateSeparatorChar) = 0
        mDateFormatStr = Replace(mDateFormatStr, mDateSeparatorChar & mDateSeparatorChar, mDateSeparatorChar)
    Loop
    Do Until InStr(mDateFormatStr, "MM") = 0
        mDateFormatStr = Replace(mDateFormatStr, "MM", "M")
    Loop
    Do Until InStr(mDateFormatStr, "yy") = 0
        mDateFormatStr = Replace(mDateFormatStr, "yy", "y")
    Loop
    Do Until Left(mDateFormatStr, 1) <> mDateSeparatorChar
        mDateFormatStr = Mid(mDateFormatStr, 2)
    Loop
    Do Until Right(mDateFormatStr, 1) <> mDateSeparatorChar
        mDateFormatStr = Left(mDateSeparatorChar, Len(mDateSeparatorChar) - 1)
    Loop
    mDateFormatStr = Replace(mDateFormatStr, "d", "dd")
    mDateFormatStr = Replace(mDateFormatStr, "M", "MM")
    mDateFormatStr = Replace(mDateFormatStr, "y", IIf(iUseTwoDigitsDateFormat, "yy", "yyyy"))
    
    DTPicker1.Format = dtpCustom
    DTPicker1.CustomFormat = mDateFormatStr
    
    iDateFormatPattern = mDateFormatStr
    iDateFormatPattern = VBA.Replace(iDateFormatPattern, "d", "#")
    iDateFormatPattern = VBA.Replace(iDateFormatPattern, "M", "#")
    iDateFormatPattern = VBA.Replace(iDateFormatPattern, "y", "#")
    
    mEmptyDate = VBA.Replace(iDateFormatPattern, "#", "_")
    Mask = iDateFormatPattern
    
    On Error Resume Next
    txtMasked.ToolTipText = mToolTipTextStart & " " & LCase$(VBA.Replace(mDateFormatStr, "y", "a")) & " " & mToolTipTextEnd
    On Error GoTo 0
    
    iAuxSelStart = txtMasked.SelStart
    If Not IsNull(mValue) Then
        txtMasked.Text = Format(CDate(mValue), mDateFormatStr)
    Else
        txtMasked.Text = mEmptyMask
    End If

End Sub


Public Property Get Font() As StdFont
Attribute Font.VB_UserMemId = -512
    Set Font = mFont
End Property

Public Property Set Font(ByVal nFont As StdFont)
    Set mFont = nFont
    Set txtMasked.Font = mFont
    Set DTPicker1.Font = mFont
    Set Combo1.Font = mFont
    UserControl_Resize
    PropertyChanged "Font"
End Property

Public Property Let Font(ByVal nFont As StdFont)
    Set Font = nFont
End Property


Private Sub UserControl_Show()
    If mFlatPending Then
        mRemoveBorder.SetControl DTPicker1
        lblBorder.Visible = True
        'lblBorder.BackColor = vbBlue
        mFlatPending = False
    End If
    DTPicker1.Visible = True
    DTPicker1.TabStop = (Not txtMasked.Visible)
    UserControl.Refresh
    If txtMasked.Visible Then txtMasked.Refresh
End Sub

Private Sub UserControl_Terminate()
    tmrSetFocus2.Enabled = False
    tmrSetFocusToMasked.Enabled = False
    tmrValidate.Enabled = False
    Set mSetFocus2Control = Nothing
    Set mMouseIcon = Nothing
    Set mForm = Nothing
    
    Set mFont = Nothing
    Detach
End Sub

Private Sub Detach()
    If mUserControlHwnd <> 0 Then
        DetachMessage Me, mUserControlHwnd, WM_UILANGCHANGED
        DetachMessage Me, mMaskedHwnd, WM_KILLFOCUS
        DetachMessage Me, mMaskedHwnd, WM_SETFOCUS
        RemoveProp mUserControlHwnd, "FnExUI"
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", mEnabled, True
    PropBag.WriteProperty "Font", mFont, Ambient.Font
    PropBag.WriteProperty "HelpContextID", mHelpContextID, 0
    PropBag.WriteProperty "MouseIcon", mMouseIcon, Nothing
    PropBag.WriteProperty "MousePointer", mMousePointer, cc2Default
    
    PropBag.WriteProperty "TextBackColor", mTextBackColor, vbWindowBackground
    PropBag.WriteProperty "CalendarBackColor", mCalendarBackColor, vbWindowBackground
    PropBag.WriteProperty "CalendarForeColor", mCalendarForeColor, vbButtonText
    PropBag.WriteProperty "CalendarTitleBackColor", mCalendarTitleBackColor, vbActiveTitleBar
    PropBag.WriteProperty "CalendarTitleForeColor", mCalendarTitleForeColor, vbActiveTitleBarText
    PropBag.WriteProperty "CalendarTrailingForeColor", mCalendarTrailingForeColor, vbGrayText
    
    PropBag.WriteProperty "MinDate", mMinDate, cDTPickerMinDate
    PropBag.WriteProperty "MaxDate", mMaxDate, cDTPickerMaxDate
    
    PropBag.WriteProperty "ToolTipTextStart", mToolTipTextStart, GetLocalizedString(efnGUIStr_DateEnter_ToolTipTextStart_Default)
    PropBag.WriteProperty "ToolTipTextEnd", mToolTipTextEnd, GetLocalizedString(efnGUIStr_DateEnter_ToolTipTextEnd_Default)
    PropBag.WriteProperty "WhatsThisHelpID", mWhatsThisHelpID, 0
    PropBag.WriteProperty "Text", txtMasked.Text, ""
    PropBag.WriteProperty "DateFormat", mDateFormat, [System date format]
    PropBag.WriteProperty "AutoValidate", mAutoValidate, True
    PropBag.WriteProperty "DateSeparator", mDateSeparator, [Use from system]
    PropBag.WriteProperty "Value", IIf(IsNull(mValue), Empty, mValue), Empty
    PropBag.WriteProperty "Appearance", mAppearance, cc3d
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim iVar As Variant
    Dim iFont As StdFont
    
    mEnabled = PropBag.ReadProperty("Enabled", True)
    
    mTextBackColor = PropBag.ReadProperty("TextBackColor", vbWindowBackground)
    mCalendarBackColor = PropBag.ReadProperty("CalendarBackColor", vbWindowBackground)
    mCalendarForeColor = PropBag.ReadProperty("CalendarForeColor", vbButtonText)
    mCalendarTitleBackColor = PropBag.ReadProperty("CalendarTitleBackColor", vbActiveTitleBar)
    mCalendarTitleForeColor = PropBag.ReadProperty("CalendarTitleForeColor", vbActiveTitleBarText)
    mCalendarTrailingForeColor = PropBag.ReadProperty("CalendarTrailingForeColor", vbGrayText)
    
    mMinDate = PropBag.ReadProperty("MinDate", cDTPickerMinDate)
    mMaxDate = PropBag.ReadProperty("MaxDate", cDTPickerMaxDate)
    
    On Error Resume Next
    Set iFont = PropBag.ReadProperty("Font", Ambient.Font)
    If Err.Number Or (iFont Is Nothing) Then
        Set iFont = Ambient.Font
    End If
    Set Font = iFont
    On Error GoTo 0
    mHelpContextID = PropBag.ReadProperty("HelpContextID", 0)
    Set mMouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    mMousePointer = PropBag.ReadProperty("MousePointer", cc2Default)
    mToolTipTextStart = PropBag.ReadProperty("ToolTipTextStart", GetLocalizedString(efnGUIStr_DateEnter_ToolTipTextStart_Default))
    mToolTipTextEnd = PropBag.ReadProperty("ToolTipTextEnd", GetLocalizedString(efnGUIStr_DateEnter_ToolTipTextEnd_Default))
    mWhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", 0)
    mAppearance = PropBag.ReadProperty("Appearance", cc3d)
    If mAppearance = ccFlat Then mFlatPending = True
    On Error Resume Next
    txtMasked.Text = PropBag.ReadProperty("Text", "")
    On Error GoTo 0
    mDateFormat = PropBag.ReadProperty("DateFormat", [System date format])
    mAutoValidate = PropBag.ReadProperty("AutoValidate", True)
    mDateSeparator = PropBag.ReadProperty("DateSeparator", [Use from system])
    
    iVar = PropBag.ReadProperty("Value", Empty)
    Value = IIf(IsEmpty(iVar), Null, iVar)
    
    SetDateFormat
    SetProperties
    
    If Ambient.UserMode Then
        mUserControlHwnd = UserControl.hWnd
        mMaskedHwnd = txtMasked.hWnd
        AttachMessage Me, mUserControlHwnd, WM_UILANGCHANGED
        AttachMessage Me, mMaskedHwnd, WM_KILLFOCUS
        AttachMessage Me, mMaskedHwnd, WM_SETFOCUS
        SetProp mUserControlHwnd, "FnExUI", 1
        If TypeOf Parent Is Form Then
            Set mForm = Parent
        End If
    End If
    
End Sub

Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal nValue As Boolean)
    If mEnabled <> nValue Then
        mEnabled = nValue
        UserControl.Enabled = mEnabled
        DTPicker1.Enabled = mEnabled
        txtMasked.Enabled = mEnabled
        PropertyChanged "Enabled"
    End If
End Property

Public Property Get FontStrikeThrough() As Boolean
Attribute FontStrikeThrough.VB_MemberFlags = "400"
    FontStrikeThrough = mFont.Strikethrough
End Property

Public Property Let FontStrikeThrough(ByVal nValue As Boolean)
    Dim iDo As Boolean
    Dim iFont As New StdFont
    
    If mFont Is Nothing Then
        iDo = True
        Set Font = New StdFont
    End If
    If Not iDo Then
        iDo = nValue <> mFont.Strikethrough
    End If
    
    If iDo Then
        Set iFont = CloneFont(mFont)
        iFont.Strikethrough = nValue
        SetFont iFont
        PropertyChanged "Font"
    End If
End Property

Public Property Get FontSize() As Single
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = mFont.Size
End Property

Public Property Let FontSize(ByVal nValue As Single)
    Dim iDo As Boolean
    Dim iFont As New StdFont
    
    If mFont Is Nothing Then
        iDo = True
        Set Font = New StdFont
    End If
    If Not iDo Then
        iDo = nValue <> mFont.Size
    End If
    
    If iDo Then
        Set iFont = CloneFont(mFont)
        iFont.Size = nValue
        SetFont iFont
        PropertyChanged "Font"
    End If
End Property

Public Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
    FontName = mFont.Name
End Property

Public Property Let FontName(ByVal nValue As String)
    Dim iDo As Boolean
    Dim iFont As New StdFont
    
    If mFont Is Nothing Then
        iDo = True
        Set Font = New StdFont
    End If
    If Not iDo Then
        iDo = nValue <> mFont.Name
    End If
    
    If iDo Then
        Set iFont = CloneFont(mFont)
        iFont.Name = nValue
        SetFont iFont
        PropertyChanged "Font"
    End If
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = mFont.Italic
End Property

Public Property Let FontItalic(ByVal nValue As Boolean)
    Dim iDo As Boolean
    Dim iFont As New StdFont
    
    If mFont Is Nothing Then
        iDo = True
        Set Font = New StdFont
    End If
    If Not iDo Then
        iDo = nValue <> mFont.Italic
    End If
    
    If iDo Then
        Set iFont = CloneFont(mFont)
        iFont.Italic = nValue
        SetFont iFont
        PropertyChanged "Font"
    End If
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = mFont.Bold
End Property

Public Property Let FontBold(ByVal nValue As Boolean)
    Dim iDo As Boolean
    Dim iFont As New StdFont
    
    If mFont Is Nothing Then
        iDo = True
        Set Font = New StdFont
    End If
    If Not iDo Then
        iDo = nValue <> mFont.Bold
    End If
    
    If iDo Then
        Set iFont = CloneFont(mFont)
        iFont.Bold = nValue
        SetFont iFont
        PropertyChanged "Font"
    End If
End Property

Public Property Get HelpContextID() As Long
    HelpContextID = mHelpContextID
End Property

Public Property Let HelpContextID(ByVal nHelpContextID As Long)
    On Error Resume Next
    mHelpContextID = nHelpContextID
    DTPicker1.HelpContextID = nHelpContextID
    txtMasked.HelpContextID = nHelpContextID
    PropertyChanged "HelpContextID"
End Property

Public Property Get MousePointer() As vbExCC2MousePointerConstants
    MousePointer = mMousePointer
End Property

Public Property Let MousePointer(ByVal nMousePointer As vbExCC2MousePointerConstants)
    If mMousePointer <> nMousePointer Then
        On Error Resume Next
        mMousePointer = nMousePointer
        DTPicker1.MousePointer = nMousePointer
        txtMasked.MousePointer = nMousePointer
        PropertyChanged "MousePointer"
    End If
End Property


Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = mMouseIcon
End Property

Public Property Set MouseIcon(ByVal nMouseIcon As IPictureDisp)
    Set mMouseIcon = nMouseIcon
    Set DTPicker1.MouseIcon = nMouseIcon
    Set txtMasked.MouseIcon = nMouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Let MouseIcon(ByVal nMouseIcon As IPictureDisp)
    Set MouseIcon = nMouseIcon
End Property


Public Property Get ToolTipTextStart() As String
    ToolTipTextStart = mToolTipTextStart
End Property

Public Property Let ToolTipTextStart(ByVal nToolTipText As String)
    On Error Resume Next
    mToolTipTextStart = nToolTipText
    txtMasked.ToolTipText = mToolTipTextStart & " " & LCase$(VBA.Replace(mDateFormatStr, "y", "a")) & " " & mToolTipTextEnd
    PropertyChanged "ToolTipTextStart"
End Property

Public Property Get ToolTipTextEnd() As String
    ToolTipTextEnd = mToolTipTextEnd
End Property

Public Property Let ToolTipTextEnd(ByVal nToolTipText As String)
    On Error Resume Next
    mToolTipTextEnd = nToolTipText
    txtMasked.ToolTipText = mToolTipTextStart & " " & LCase$(VBA.Replace(mDateFormatStr, "y", "a")) & " " & mToolTipTextEnd
    PropertyChanged "ToolTipTextEnd"
End Property

Private Sub SetFont(nFont As StdFont)
    Set mFont = nFont
    Set DTPicker1.Font = nFont
    Set txtMasked.Font = nFont
    Set Combo1.Font = nFont
    UserControl_Resize
End Sub

Private Sub SetProperties()
    Set DTPicker1.MouseIcon = mMouseIcon
    Set txtMasked.MouseIcon = mMouseIcon
    DTPicker1.MousePointer = mMousePointer
    txtMasked.MousePointer = mMousePointer
    DTPicker1.HelpContextID = mHelpContextID
    txtMasked.HelpContextID = mHelpContextID
    txtMasked.ToolTipText = mToolTipTextStart & " " & LCase$(VBA.Replace(mDateFormatStr, "y", "a")) & " " & mToolTipTextEnd
    DTPicker1.WhatsThisHelpID = mWhatsThisHelpID
    txtMasked.WhatsThisHelpID = mWhatsThisHelpID
    
    
    DTPicker1.TextBackColor = mTextBackColor
    DTPicker1.CalendarBackColor = mCalendarBackColor
    DTPicker1.CalendarForeColor = mCalendarForeColor
    DTPicker1.CalendarTitleBackColor = mCalendarTitleBackColor
    DTPicker1.CalendarTitleForeColor = mCalendarTitleForeColor
    DTPicker1.CalendarTrailingForeColor = mCalendarTrailingForeColor
    
    DTPicker1.MinDate = mMinDate
    DTPicker1.MaxDate = mMaxDate
    
    DTPicker1.Enabled = Enabled
    txtMasked.Enabled = Enabled
    UserControl.Enabled = Enabled
End Sub

Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "400"
    Text = txtMasked.Text
End Property

Public Property Let Text(nText As String)
    If IsDate(nText) Then
        If IsValidDate(CDate(nText)) Then
            txtMasked.Text = Format(nText, mDateFormatStr)
        Else
            txtMasked.Text = mEmptyDate
        End If
    Else
        txtMasked.Text = mEmptyDate
    End If
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Parent()
    If Ambient.UserMode Then
        On Error Resume Next
        Set Parent = UserControl.Parent
    End If
End Property

Private Sub SetFocusTo(nControl As Variant)
'    Debug.Print "GetActiveFormHwnd = GetFormHwnd(nControl.hWnd): " & CBool(GetActiveFormHwnd = GetFormHwnd(nControl.hWnd))
    If GetActiveFormHwnd = GetFormHwnd(nControl.hWnd) Then
        If VarType(nControl) = vbLong Then
            SetFocusAPI nControl
        Else
            On Error Resume Next
            nControl.SetFocus
            On Error GoTo 0
        End If
    End If
End Sub

Private Function GetFormHwnd(nControlHwnd As Long)
    Dim lPar As Long
    Dim iHwnd As Long
    
    iHwnd = nControlHwnd
    lPar = GetParent(iHwnd)
    While lPar <> 0
        iHwnd = lPar
        lPar = GetParent(lPar)
    Wend
    GetFormHwnd = iHwnd
End Function

Public Function CloneFont(nOrigFont) As StdFont
    Dim iFont As New StdFont
    
    If nOrigFont Is Nothing Then Exit Function
    If Not TypeOf nOrigFont Is StdFont Then Exit Function
    
    iFont.Name = nOrigFont.Name
    iFont.Size = nOrigFont.Size
    iFont.Bold = nOrigFont.Bold
    iFont.Italic = nOrigFont.Italic
    iFont.Strikethrough = nOrigFont.Strikethrough
    iFont.Underline = nOrigFont.Underline
    iFont.Weight = nOrigFont.Weight
    iFont.Charset = nOrigFont.Charset
    
    Set CloneFont = iFont
End Function

Public Function Validate(Optional nAllowEmpty As Boolean = True) As Boolean
    Dim iIsValid As Boolean
    
    Validate1 nAllowEmpty, , iIsValid
    Validate = iIsValid
End Function

Private Sub Validate1(nAllowEmpty As Boolean, Optional nKeyReturn As Boolean, Optional ByRef nIsValid As Boolean, Optional ByRef nUserReEnteringDate As Boolean)
    Dim iStrDate As String
    Dim iStrDay As String
    Dim iStrMonth As String
    Dim iStrYear As String
    Dim iPosD As Long
    Dim iLenD  As Long
    Dim iPosM As Long
    Dim iLenM  As Long
    Dim iPosY As Long
    Dim iLenY  As Long
    Dim iDayValue As Long
    Dim iMonthValue As Long
    Dim iYearValue As Long
    Dim iValue As Variant
    
    nIsValid = True
    nUserReEnteringDate = False
    If mInsideValidate1 Then Exit Sub
    tmrValidate.Enabled = False
    
    iStrDate = txtMasked.Text
    If nAllowEmpty Then
        If iStrDate = mEmptyDate Then
            Exit Sub
        End If
    End If
    If tmrSetFocusToMasked.Enabled Then
        tmrSetFocusToMasked_Timer
        Exit Sub
    End If

    iPosD = InStr(mDateFormatStr, "d")
    If iPosD = 0 Then Exit Sub
    iLenD = VBA.InStrRev(mDateFormatStr, "d")
    iLenD = iLenD - iPosD + 1
    
    iPosM = InStr(mDateFormatStr, "M")
    If iPosM = 0 Then Exit Sub
    iLenM = VBA.InStrRev(mDateFormatStr, "M")
    iLenM = iLenM - iPosM + 1
    
    iPosY = InStr(mDateFormatStr, "y")
    If iPosY = 0 Then Exit Sub
    iLenY = VBA.InStrRev(mDateFormatStr, "y")
    iLenY = iLenY - iPosY + 1
    
    mInsideValidate1 = True
    
    iStrDay = Trim$(Replace(Mid(iStrDate, iPosD, iLenD), "_", " "))
    iStrMonth = Trim$(Replace(Mid(iStrDate, iPosM, iLenM), "_", " "))
    iStrYear = Trim$(Replace(Mid(iStrDate, iPosY, iLenY), "_", " "))
    If Not IsNumeric(iStrDay) Then
        nIsValid = False
        If Not IsWindowVisibleOnScreen(GetParentFormHwnd(UserControl.hWnd), True) Then
            txtMasked.Text = mEmptyMask
            GoTo TheExit
        End If
        If Not (IsWindowVisibleOnScreen(UserControl.hWnd, True) And UserControl.Enabled) Then
            txtMasked.Text = mEmptyMask
            GoTo TheExit
        End If
        If (mValidationErrors = 0) Or nKeyReturn And (mValidationErrors < 2) Then
            'efnGUIStr_DateEnter_Validate1_MsgBoxError1: "You did not enter the day in the date entry."
            MsgBox GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxError1), vbExclamation, GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxTitle)
            mValidationErrors = mValidationErrors + 1
        Else
            If MsgBox(GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxError1), vbExclamation + vbOKCancel, GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxTitle)) = vbCancel Then
                txtMasked.Text = mEmptyMask
                GoTo TheExit
            End If
        End If
        mDontDoGotFocus = True
        nUserReEnteringDate = True
        If mOnFocus Then tmrSetFocusToMasked.Enabled = True
        tmrValidate.Enabled = False
        txtMasked.SelStart = iPosD - 1
        txtMasked.SelLength = iLenD
        GoTo TheExit
    Else
        iDayValue = Val(iStrDay)
        If iDayValue < 1 Then
            nIsValid = False
            If Not IsWindowVisibleOnScreen(GetParentFormHwnd(UserControl.hWnd), True) Then
                txtMasked.Text = mEmptyMask
                GoTo TheExit
            End If
            If Not (IsWindowVisibleOnScreen(UserControl.hWnd, True) And UserControl.Enabled) Then
                txtMasked.Text = mEmptyMask
                GoTo TheExit
            End If
            If (mValidationErrors = 0) Or nKeyReturn And (mValidationErrors < 2) Then
                ' efnGUIStr_DateEnter_Validate1_MsgBoxError2: "The day can't be zero."
                MsgBox GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxError2), vbExclamation, GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxTitle)
                mValidationErrors = mValidationErrors + 1
            Else
                If MsgBox(GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxError2), vbExclamation + vbOKCancel, GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxTitle)) = vbCancel Then
                    txtMasked.Text = mEmptyMask
                    GoTo TheExit
                End If
            End If
            mDontDoGotFocus = True
            nUserReEnteringDate = True
            If mOnFocus Then tmrSetFocusToMasked.Enabled = True
            tmrValidate.Enabled = False
            txtMasked.SelStart = iPosD - 1
            txtMasked.SelLength = iLenD
            GoTo TheExit
        ElseIf iDayValue > 31 Then
            nIsValid = False
            If Not IsWindowVisibleOnScreen(GetParentFormHwnd(UserControl.hWnd), True) Then
                txtMasked.Text = mEmptyMask
                GoTo TheExit
            End If
            If Not (IsWindowVisibleOnScreen(UserControl.hWnd, True) And UserControl.Enabled) Then
                txtMasked.Text = mEmptyMask
                GoTo TheExit
            End If
            If (mValidationErrors = 0) Or nKeyReturn And (mValidationErrors < 2) Then
                ' efnGUIStr_DateEnter_Validate1_MsgBoxError3: "The day can't be greater than 31."
                MsgBox GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxError3), vbExclamation, GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxTitle)
                mValidationErrors = mValidationErrors + 1
            Else
                If MsgBox(GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxError3), vbExclamation + vbOKCancel, GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxTitle)) = vbCancel Then
                    txtMasked.Text = mEmptyMask
                    GoTo TheExit
                End If
            End If
            mDontDoGotFocus = True
            nUserReEnteringDate = True
            If mOnFocus Then tmrSetFocusToMasked.Enabled = True
            tmrValidate.Enabled = False
            txtMasked.SelStart = iPosD - 1
            txtMasked.SelLength = iLenD
            GoTo TheExit
        End If
    End If
    If Not IsNumeric(iStrMonth) Then
        nIsValid = False
        If Not IsWindowVisibleOnScreen(GetParentFormHwnd(UserControl.hWnd), True) Then
            txtMasked.Text = mEmptyMask
            GoTo TheExit
        End If
        If Not (IsWindowVisibleOnScreen(UserControl.hWnd, True) And UserControl.Enabled) Then
            txtMasked.Text = mEmptyMask
            GoTo TheExit
        End If
        If (mValidationErrors = 0) Or nKeyReturn And (mValidationErrors < 2) Then
            ' efnGUIStr_DateEnter_Validate1_MsgBoxError4: "You did not enter the month in the date entry."
            MsgBox GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxError4), vbExclamation, GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxTitle)
            mValidationErrors = mValidationErrors + 1
        Else
            If MsgBox(GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxError4), vbExclamation + vbOKCancel, GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxTitle)) = vbCancel Then
                txtMasked.Text = mEmptyMask
                GoTo TheExit
            End If
        End If
        mDontDoGotFocus = True
        nUserReEnteringDate = True
        If mOnFocus Then tmrSetFocusToMasked.Enabled = True
        tmrValidate.Enabled = False
        txtMasked.SelStart = iPosM - 1
        txtMasked.SelLength = iLenM
        GoTo TheExit
    Else
        iMonthValue = Val(iStrMonth)
        If iMonthValue < 1 Or iMonthValue > 12 Then
            nIsValid = False
            If Not IsWindowVisibleOnScreen(GetParentFormHwnd(UserControl.hWnd), True) Then
                txtMasked.Text = mEmptyMask
                GoTo TheExit
            End If
            If Not (IsWindowVisibleOnScreen(UserControl.hWnd, True) And UserControl.Enabled) Then
                txtMasked.Text = mEmptyMask
                GoTo TheExit
            End If
            If (mValidationErrors = 0) Or nKeyReturn And (mValidationErrors < 2) Then
                ' efnGUIStr_DateEnter_Validate1_MsgBoxError5: "The value of the month must be between 1 y 12."
                MsgBox GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxError5), vbExclamation, GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxTitle)
                mValidationErrors = mValidationErrors + 1
            Else
                If MsgBox(GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxError5), vbExclamation + vbOKCancel, GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxTitle)) = vbCancel Then
                    txtMasked.Text = mEmptyMask
                    GoTo TheExit
                End If
            End If
            mDontDoGotFocus = True
            nUserReEnteringDate = True
            If mOnFocus Then tmrSetFocusToMasked.Enabled = True
            tmrValidate.Enabled = False
            txtMasked.SelStart = iPosM - 1
            txtMasked.SelLength = iLenM
            GoTo TheExit
        End If
    End If
    If Not IsNumeric(iStrYear) Then
        nIsValid = False
        If Not IsWindowVisibleOnScreen(GetParentFormHwnd(UserControl.hWnd), True) Then
            txtMasked.Text = mEmptyMask
            GoTo TheExit
        End If
        If Not (IsWindowVisibleOnScreen(UserControl.hWnd, True) And UserControl.Enabled) Then
            txtMasked.Text = mEmptyMask
            GoTo TheExit
        End If
        If (mValidationErrors = 0) Or nKeyReturn And (mValidationErrors < 2) Then
            ' efnGUIStr_DateEnter_Validate1_MsgBoxError6: "You did not enter the year in the date entry."
            MsgBox GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxError6), vbExclamation, GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxTitle)
            mValidationErrors = mValidationErrors + 1
        Else
            If MsgBox(GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxError6), vbExclamation + vbOKCancel, GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxTitle)) = vbCancel Then
                txtMasked.Text = mEmptyMask
                GoTo TheExit
            End If
        End If
        mDontDoGotFocus = True
        nUserReEnteringDate = True
        If mOnFocus Then tmrSetFocusToMasked.Enabled = True
        tmrValidate.Enabled = False
        txtMasked.SelStart = iPosY - 1
        txtMasked.SelLength = iLenY
        GoTo TheExit
    Else
        iYearValue = Val(iStrYear)
    End If
    iValue = DateSerial(iYearValue, iMonthValue, iDayValue)
    
    If iValue < mMinDate Then
        nIsValid = False
        If Not IsWindowVisibleOnScreen(GetParentFormHwnd(UserControl.hWnd), True) Then
            'txtMasked.Text = Format(CDate(mMinDate), mDateFormatStr)
            txtMasked.Text = mEmptyMask
            GoTo TheExit
        End If
        If Not (IsWindowVisibleOnScreen(UserControl.hWnd, True) And UserControl.Enabled) Then
            'txtMasked.Text = Format(CDate(mMinDate), mDateFormatStr)
            txtMasked.Text = mEmptyMask
            GoTo TheExit
        End If
        If (mValidationErrors = 0) Or nKeyReturn And (mValidationErrors < 2) Then
            ' efnGUIStr_DateEnter_Validate1_MsgBoxError7:
            MsgBox GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxError7) & " " & mMinDate, vbExclamation, GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxTitle)
            mValidationErrors = mValidationErrors + 1
        Else
            If MsgBox(GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxError7) & " " & mMinDate, vbExclamation + vbOKCancel, GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxTitle)) = vbCancel Then
                'txtMasked.Text = Format(CDate(mMinDate), mDateFormatStr)
                txtMasked.Text = mEmptyMask
                GoTo TheExit
            End If
        End If
        mDontDoGotFocus = True
        nUserReEnteringDate = True
        If mOnFocus Then tmrSetFocusToMasked.Enabled = True
        tmrValidate.Enabled = False
        txtMasked.SelStart = iPosY - 1
        txtMasked.SelLength = iLenY
        GoTo TheExit
    End If
    
    If iValue > mMaxDate Then
        nIsValid = False
        If Not IsWindowVisibleOnScreen(GetParentFormHwnd(UserControl.hWnd), True) Then
            'txtMasked.Text = Format(CDate(mMaxDate), mDateFormatStr)
            txtMasked.Text = mEmptyMask
            GoTo TheExit
        End If
        If Not (IsWindowVisibleOnScreen(UserControl.hWnd, True) And UserControl.Enabled) Then
            'txtMasked.Text = Format(CDate(mMaxDate), mDateFormatStr)
            txtMasked.Text = mEmptyMask
            GoTo TheExit
        End If
        If (mValidationErrors = 0) Or nKeyReturn And (mValidationErrors < 2) Then
            ' efnGUIStr_DateEnter_Validate1_MsgBoxError8:
            MsgBox GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxError8) & " " & mMaxDate, vbExclamation, GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxTitle)
            mValidationErrors = mValidationErrors + 1
        Else
            If MsgBox(GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxError8) & " " & mMaxDate, vbExclamation + vbOKCancel, GetLocalizedString(efnGUIStr_DateEnter_Validate1_MsgBoxTitle)) = vbCancel Then
                'txtMasked.Text = Format(CDate(mMaxDate), mDateFormatStr)
                txtMasked.Text = mEmptyMask
                GoTo TheExit
            End If
        End If
        mDontDoGotFocus = True
        nUserReEnteringDate = True
        If mOnFocus Then tmrSetFocusToMasked.Enabled = True
        tmrValidate.Enabled = False
        txtMasked.SelStart = iPosY - 1
        txtMasked.SelLength = iLenY
        GoTo TheExit
    End If
    
    Value = iValue
    mValidationErrors = 0
    mNeedValidation = False
    
TheExit:
    mInsideValidate1 = False
End Sub

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrPostProcess
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Select Case iMsg
        Case WM_UILANGCHANGED
            UILangChange wParam
        Case WM_KILLFOCUS
            If mAutoValidate Then
                On Error Resume Next
                If txtMasked.Text <> mEmptyDate Then
                    If mNeedValidation Then
                        tmrValidate.Enabled = False
                        tmrValidate.Enabled = True
                    End If
                End If
            End If
            mOnFocus = False
        Case WM_SETFOCUS
            On Error Resume Next
            tmrValidate.Enabled = False
            tmrSetFocusToMasked.Enabled = False
            mOnFocus = True
    End Select
End Function

Private Sub UILangChange(nPrevLang As Long)
    If mToolTipTextStart = GetLocalizedString(efnGUIStr_DateEnter_ToolTipTextStart_Default, , nPrevLang) Then ToolTipTextStart = GetLocalizedString(efnGUIStr_DateEnter_ToolTipTextStart_Default)
    If mToolTipTextEnd = GetLocalizedString(efnGUIStr_DateEnter_ToolTipTextEnd_Default, , nPrevLang) Then ToolTipTextEnd = GetLocalizedString(efnGUIStr_DateEnter_ToolTipTextEnd_Default)
End Sub

Private Sub RaiseEvent_TextChange()
    Static sLast As String
    Static sLastWasNull As Boolean
    
    If (txtMasked.Text <> sLast) Or (sLastWasNull <> IsNull(mValue)) Then
        sLast = txtMasked.Text
        RaiseEvent TextChange
        mNeedValidation = True
        sLastWasNull = IsNull(mValue)
    End If
End Sub

Public Property Get NeedValidation() As Boolean
    NeedValidation = mNeedValidation
End Property

Private Sub SetFocusTo2(nControl As Variant)
'    Debug.Print Ambient.DisplayName
    tmrSetFocus2.Enabled = False
    If Not mOnFocus Then Exit Sub
    If mOnFocus Then tmrSetFocus2.Enabled = True
    Set mSetFocus2Control = nControl
End Sub


Public Property Get DroppedDown() As Boolean
    DroppedDown = DTPicker1.DroppedDown
End Property

Public Property Let DroppedDown(ByVal nValue As Boolean)
    If nValue <> DTPicker1.DroppedDown Then
        DTPicker1.DroppedDown = nValue
    End If
End Property

Public Property Get InTextMode() As Boolean
    If Ambient.UserMode Then
        InTextMode = txtMasked.Visible
    End If
End Property


Public Property Get GetIdealHeight() As Long
    GetIdealHeight = DTPicker1.GetIdealHeight
End Property

Public Property Get GetIdealWith() As Long
    GetIdealWith = DTPicker1.GetIdealWith
End Property


Public Property Get TextBackColor() As OLE_COLOR
    TextBackColor = mTextBackColor
End Property

Public Property Let TextBackColor(ByVal nValue As OLE_COLOR)
    If mTextBackColor <> nValue Then
        mTextBackColor = nValue
        DTPicker1.TextBackColor = mTextBackColor
        PropertyChanged "TextBackColor"
    End If
End Property


Public Property Get CalendarBackColor() As OLE_COLOR
    CalendarBackColor = mCalendarBackColor
End Property

Public Property Let CalendarBackColor(ByVal nValue As OLE_COLOR)
    If mCalendarBackColor <> nValue Then
        mCalendarBackColor = nValue
        DTPicker1.CalendarBackColor = mCalendarBackColor
        PropertyChanged "CalendarBackColor"
    End If
End Property


Public Property Get CalendarForeColor() As OLE_COLOR
    CalendarForeColor = mCalendarForeColor
End Property

Public Property Let CalendarForeColor(ByVal nValue As OLE_COLOR)
    If mCalendarForeColor <> nValue Then
        mCalendarForeColor = nValue
        DTPicker1.CalendarForeColor = mCalendarForeColor
        PropertyChanged "CalendarForeColor"
    End If
End Property


Public Property Get CalendarTitleBackColor() As OLE_COLOR
    CalendarTitleBackColor = mCalendarTitleBackColor
End Property

Public Property Let CalendarTitleBackColor(ByVal nValue As OLE_COLOR)
    If mCalendarTitleBackColor <> nValue Then
        mCalendarTitleBackColor = nValue
        DTPicker1.CalendarTitleBackColor = mCalendarTitleBackColor
        PropertyChanged "CalendarTitleBackColor"
    End If
End Property


Public Property Get CalendarTitleForeColor() As OLE_COLOR
    CalendarTitleForeColor = mCalendarTitleForeColor
End Property

Public Property Let CalendarTitleForeColor(ByVal nValue As OLE_COLOR)
    If mCalendarTitleForeColor <> nValue Then
        mCalendarTitleForeColor = nValue
        DTPicker1.CalendarTitleForeColor = mCalendarTitleForeColor
        PropertyChanged "CalendarTitleForeColor"
    End If
End Property


Public Property Get CalendarTrailingForeColor() As OLE_COLOR
    CalendarTrailingForeColor = mCalendarTrailingForeColor
End Property

Public Property Let CalendarTrailingForeColor(ByVal nValue As OLE_COLOR)
    If mCalendarTrailingForeColor <> nValue Then
        mCalendarTrailingForeColor = nValue
        DTPicker1.CalendarTrailingForeColor = mCalendarTrailingForeColor
        PropertyChanged "CalendarTrailingForeColor"
    End If
End Property


Public Property Get MinDate() As Date
    MinDate = mMinDate
End Property

Public Property Let MinDate(ByVal nValue As Date)
    If nValue <> mMinDate Then
        If nValue >= cDTPickerMinDate Then
            mMinDate = nValue
            DTPicker1.MinDate = mMinDate
            PropertyChanged "MinDate"
        Else
            RaiseError 380, TypeName(Me), "The minimum date must be equal or greater than " & cDTPickerMinDate
        End If
    End If
End Property


Public Property Get MaxDate() As Date
    MaxDate = mMaxDate
End Property

Public Property Let MaxDate(ByVal nValue As Date)
    If nValue <> mMaxDate Then
        If nValue <= cDTPickerMaxDate Then
            mMaxDate = nValue
            DTPicker1.MaxDate = mMaxDate
            PropertyChanged "MaxDate"
        Else
            RaiseError 380, TypeName(Me), "The maximum date must be equal or lower than " & cDTPickerMaxDate
        End If
    End If
End Property


Private Function IsValidDate(nDate As Variant) As Boolean
    If IsDate(nDate) Then
        If (DateValue(nDate) >= mMinDate) And (DateValue(nDate) <= mMaxDate) Then
            IsValidDate = True
        End If
    End If
End Function

Public Sub Refresh()
'    UserControl.Refresh
'    RedrawWindow UserControl.hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN
    Call SetWindowPos(UserControl.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_FRAMECHANGED)
End Sub


Private Sub SetIncompleteDateInDTPicker1()
    Dim iStrDate As String
    Dim iStrDay As String
    Dim iStrMonth As String
    Dim iStrYear As String
    Dim iPosD As Long
    Dim iLenD  As Long
    Dim iPosM As Long
    Dim iLenM  As Long
    Dim iPosY As Long
    Dim iLenY  As Long
    Dim iDayValue As Long
    Dim iMonthValue As Long
    Dim iYearValue As Long
    Dim iValue As Date
    
    iStrDate = txtMasked.Text
    
    iPosD = InStr(mDateFormatStr, "d")
    If iPosD = 0 Then Exit Sub
    iLenD = VBA.InStrRev(mDateFormatStr, "d")
    iLenD = iLenD - iPosD + 1
    
    iPosM = InStr(mDateFormatStr, "M")
    If iPosM = 0 Then Exit Sub
    iLenM = VBA.InStrRev(mDateFormatStr, "M")
    iLenM = iLenM - iPosM + 1
    
    iPosY = InStr(mDateFormatStr, "y")
    If iPosY = 0 Then Exit Sub
    iLenY = VBA.InStrRev(mDateFormatStr, "y")
    iLenY = iLenY - iPosY + 1
    
    iStrDay = Trim$(Replace(Mid(iStrDate, iPosD, iLenD), "_", " "))
    iStrMonth = Trim$(Replace(Mid(iStrDate, iPosM, iLenM), "_", " "))
    iStrYear = Trim$(Replace(Mid(iStrDate, iPosY, iLenY), "_", " "))
    
    If IsNumeric(iStrDay) Then
        iDayValue = Val(iStrDay)
    End If
    If IsNumeric(iStrMonth) Then
        iMonthValue = Val(iStrMonth)
    End If
    If IsNumeric(iStrYear) Then
        iYearValue = Val(iStrYear)
    End If

    If (iDayValue < 1) Or (iDayValue > 31) Then
        iDayValue = Day(Date)
    End If
    If (iMonthValue < 1) Or (iMonthValue > 12) Then
        iMonthValue = Month(Date)
    End If
    If (iYearValue > 0) And (iYearValue < 100) Then
        If iYearValue < GetTwoDigitYearCenturyChange Then
            iYearValue = iYearValue + 2000
        Else
            iYearValue = iYearValue + 1900
        End If
    End If
    If (iYearValue < Year(mMinDate)) Or (iYearValue > Year(mMaxDate)) Then
        iYearValue = Year(Date)
    End If
    
    iValue = DateSerial(iYearValue, iMonthValue, iDayValue)
    If iValue < mMinDate Then
        iValue = mMinDate
    End If
    If iValue > mMaxDate Then
        iValue = mMaxDate
    End If
    
    mSettingIncompleteValue = True
    DTPicker1.Value = iValue
    mSettingIncompleteValue = False
End Sub


Public Property Get Appearance() As vbExAppearanceConstants
    Appearance = mAppearance
End Property

Public Property Let Appearance(nValue As vbExAppearanceConstants)
    If nValue <> mAppearance Then
        mAppearance = nValue
        If mAppearance = ccFlat Then
            If IsWindowVisible(DTPicker1.hWnd) Then
                mRemoveBorder.SetControl DTPicker1
                lblBorder.Visible = True
            Else
                mFlatPending = True
            End If
        Else
            Set mRemoveBorder = Nothing
            lblBorder.Visible = False
        End If
        PropertyChanged "Appearance"
    End If
End Property


Public Property Get DateFormat() As vbExDateFormatConstants
    DateFormat = mDateFormat
End Property

Public Property Let DateFormat(nValue As vbExDateFormatConstants)
    If (nValue < 0) Or (nValue > 8) Then
        RaiseError 380, TypeName(Me)
        Exit Property
    End If
    If nValue <> mDateFormat Then
        mDateFormat = nValue
        SetDateFormat
        PropertyChanged "DateFormat"
    End If
End Property

Public Property Get hWndTextBox() As Long
    hWndTextBox = txtMasked.hWnd
End Property

Public Property Get hWndDTPicker() As Long
    hWndDTPicker = DTPicker1.hWnd
End Property
