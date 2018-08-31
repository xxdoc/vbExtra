VERSION 5.00
Begin VB.UserControl FontPicker 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "ctlFontPicker.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ctlFontPicker.ctx":0033
   Begin vbExtra.ButtonEx cmdSelectFont 
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      ToolTipText     =   "Seleccionar fuente"
      Top             =   150
      Width           =   300
      _ExtentX        =   402
      _ExtentY        =   402
      ButtonStyle     =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "..."
   End
   Begin VB.PictureBox picSample 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   300
      ScaleHeight     =   588
      ScaleWidth      =   1992
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   1995
   End
End
Attribute VB_Name = "FontPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Private Const WM_UILANGCHANGED As Long = WM_USER + 12

Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long

Public Event Change()

Public Enum vbExTextEndingStyleConstants
    vxTECut = 0
    vxTEVanish = 1
    vxTEAddElipsis = 2
End Enum

Private mSampleText As String
Private Const cDefaultSampleText = "AaBbYyZz"
Private mForeColor As Long
Private Const cDefaultForeColor = vbButtonText
Private mBackColor As Long
Private Const cDefaultBackColor = vbWindowBackground
Private mChooseForeColor As Boolean
Private mBorderColor As Long
Private Const cDefaultBorderColor = vbWindowFrame
Private mVisualStyles As Boolean
Private mButtonToolTipText As String
Private mButtonToolTipText_Default As String
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1
Private mSampleTextEnding As vbExTextEndingStyleConstants
Private Const cDefaultSampleTextEnding = vxTEVanish
Private mEnabled As Boolean

Private mUserControlHwnd As Long

Public Property Get SampleText() As String
    SampleText = mSampleText
End Property

Public Property Let SampleText(nValue As String)
    If nValue <> mSampleText Then
        mSampleText = nValue
        PropertyChanged "SampleText"
        DrawSample
    End If
End Property


Public Property Get Font() As StdFont
Attribute Font.VB_MemberFlags = "200"
    Set Font = mFont
End Property

Public Property Set Font(ByVal nValue As StdFont)
    Set mFont = nValue
    PropertyChanged "Font"
    DrawSample
End Property

Public Property Let Font(ByVal nValue As StdFont)
    Set Font = nValue
End Property


Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(nValue As OLE_COLOR)
    If nValue <> mForeColor Then
        mForeColor = nValue
        PropertyChanged "ForeColor"
        DrawSample
    End If
End Property


Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let BackColor(nValue As OLE_COLOR)
    If nValue <> mBackColor Then
        mBackColor = nValue
        PropertyChanged "BackColor"
        DrawSample
    End If
End Property


Public Property Get ChooseForeColor() As Boolean
    ChooseForeColor = mChooseForeColor
End Property

Public Property Let ChooseForeColor(nValue As Boolean)
    If nValue <> mChooseForeColor Then
        mChooseForeColor = nValue
        PropertyChanged "ChooseForeColor"
    End If
End Property


Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mBorderColor
End Property

Public Property Let BorderColor(nValue As OLE_COLOR)
    If nValue <> mBorderColor Then
        mBorderColor = nValue
        PropertyChanged "BorderColor"
        DrawSample
    End If
End Property


Public Property Get VisualStyles() As Boolean
    VisualStyles = mVisualStyles
End Property

Public Property Let VisualStyles(nValue As Boolean)
    If nValue <> mVisualStyles Then
        mVisualStyles = nValue
        PropertyChanged "VisualStyles"
        DrawSample
    End If
End Property


Public Property Get ButtonToolTipText() As String
    ButtonToolTipText = mButtonToolTipText
End Property

Public Property Let ButtonToolTipText(nValue As String)
    If nValue <> mButtonToolTipText Then
        mButtonToolTipText = nValue
        PropertyChanged "ButtonToolTipText"
        cmdSelectFont.ToolTipText = mButtonToolTipText
    End If
End Property


Public Property Get SampleTextEnding() As vbExTextEndingStyleConstants
    SampleTextEnding = mSampleTextEnding
End Property

Public Property Let SampleTextEnding(nValue As vbExTextEndingStyleConstants)
    If nValue <> mSampleTextEnding Then
        mSampleTextEnding = nValue
        PropertyChanged "SampleTextEnding"
        DrawSample
    End If
End Property


Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(nValue As Boolean)
    If nValue <> mEnabled Then
        mEnabled = nValue
        PropertyChanged "Enabled"
        cmdSelectFont.Enabled = mEnabled
        UserControl.Enabled = mEnabled
        picSample.Enabled = mEnabled
        DrawSample
    End If
End Property


Private Sub cmdSelectFont_Click()
    Dim iDlg As New CommonDialogExObject
    
    Set iDlg.Font = mFont
    
    If mChooseForeColor Then
        iDlg.Color = mForeColor
        iDlg.ShowFont cdeCFEffects
    Else
        iDlg.ShowFont
    End If
    If Not iDlg.Canceled Then
        Set mFont = iDlg.Font
        If mChooseForeColor Then
            mForeColor = iDlg.Color
        End If
        DrawSample
        RaiseEvent Change
    End If

End Sub

Private Sub mFont_FontChanged(ByVal PropertyName As String)
    DrawSample
End Sub

Private Sub picSample_Click()
    cmdSelectFont_Click
End Sub

Private Sub picSample_DblClick()
    cmdSelectFont_Click
End Sub

Private Sub picSample_GotFocus()
    cmdSelectFont.SetFocus
End Sub

Private Sub UserControl_Click()
    cmdSelectFont_Click
End Sub

Private Sub UserControl_DblClick()
    cmdSelectFont_Click
End Sub

Private Sub UserControl_EnterFocus()
    cmdSelectFont.SetFocus
End Sub

Private Sub UserControl_GotFocus()
    cmdSelectFont.SetFocus
End Sub

Private Sub UserControl_Initialize()
    mButtonToolTipText_Default = GetLocalizedString(efnGUIStr_FontPicker_ButtonToolTipTextDefault)
End Sub

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    mSampleText = cDefaultSampleText
    Set mFont = UserControl.Font
    mForeColor = cDefaultForeColor
    mBackColor = cDefaultBackColor
    mChooseForeColor = False
    mBorderColor = cDefaultBorderColor
    mVisualStyles = True
    mButtonToolTipText = mButtonToolTipText_Default
    mSampleTextEnding = cDefaultSampleTextEnding
    mEnabled = True
    UserControl.Size 1600, 400
    If Ambient.UserMode Then
        mUserControlHwnd = UserControl.hWnd
        AttachMessage Me, mUserControlHwnd, WM_UILANGCHANGED
        SetProp mUserControlHwnd, "FnExUI", 1
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mSampleText = PropBag.ReadProperty("SampleText", cDefaultSampleText)
    Set mFont = PropBag.ReadProperty("Font", UserControl.Font)
    mForeColor = PropBag.ReadProperty("ForeColor", cDefaultForeColor)
    mBackColor = PropBag.ReadProperty("BackColor", cDefaultBackColor)
    mChooseForeColor = PropBag.ReadProperty("ChooseForeColor", False)
    mBorderColor = PropBag.ReadProperty("BorderColor", cDefaultBorderColor)
    mVisualStyles = PropBag.ReadProperty("VisualStyles", True)
    mButtonToolTipText = PropBag.ReadProperty("ButtonToolTipText", mButtonToolTipText_Default)
    mSampleTextEnding = PropBag.ReadProperty("SampleTextEnding", cDefaultSampleTextEnding)
    mEnabled = PropBag.ReadProperty("Enabled", True)
    
    cmdSelectFont.Enabled = mEnabled
    UserControl.Enabled = mEnabled
    picSample.Enabled = mEnabled
    
    If Ambient.UserMode Then
        mUserControlHwnd = UserControl.hWnd
        AttachMessage Me, mUserControlHwnd, WM_UILANGCHANGED
        SetProp mUserControlHwnd, "FnExUI", 1
    End If
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width < 800 Then UserControl.Width = 800
    If UserControl.Height < 280 Then UserControl.Height = 280
    DrawSample
    cmdSelectFont.Left = UserControl.ScaleWidth - cmdSelectFont.Width - 60
    cmdSelectFont.Top = UserControl.ScaleHeight / 2 - cmdSelectFont.Height / 2
    cmdSelectFont.ToolTipText = mButtonToolTipText
End Sub

Private Sub UserControl_Terminate()
    If mUserControlHwnd <> 0 Then
        DetachMessage Me, mUserControlHwnd, WM_UILANGCHANGED
        RemoveProp mUserControlHwnd, "FnExUI"
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "SampleText", mSampleText, cDefaultSampleText
    PropBag.WriteProperty "Font", mFont, UserControl.Font
    PropBag.WriteProperty "ForeColor", mForeColor, cDefaultForeColor
    PropBag.WriteProperty "BackColor", mBackColor, cDefaultBackColor
    PropBag.WriteProperty "ChooseForeColor", mChooseForeColor, False
    PropBag.WriteProperty "BorderColor", mBorderColor, cDefaultBorderColor
    PropBag.WriteProperty "VisualStyles", mVisualStyles, True
    PropBag.WriteProperty "ButtonToolTipText", mButtonToolTipText, mButtonToolTipText_Default
    PropBag.WriteProperty "SampleTextEnding", mSampleTextEnding, cDefaultSampleTextEnding
    PropBag.WriteProperty "Enabled", mEnabled, True
End Sub

Private Sub DrawSample()
    Dim iBorderColor As Long
    Dim iX As Long
    Dim iY As Long
    Dim iColor As Long
    Dim iBackColor As Long
    Dim p As Long
    Dim iSampleText As String
    
    If mFont Is Nothing Then Exit Sub
    UserControl.Cls
    picSample.Cls
    Set picSample.Font = mFont
    picSample.Height = picSample.TextHeight(mSampleText)
    If picSample.Height > (UserControl.ScaleHeight - 2 * Screen.TwipsPerPixelY) Then
        picSample.Height = UserControl.ScaleHeight - 2 * Screen.TwipsPerPixelY
    End If
    picSample.Top = ScaleHeight / 2 - picSample.Height / 2
    If picSample.Top > (ScaleHeight - picSample.Height) Then
        picSample.Top = ScaleHeight - picSample.Height
    End If
    picSample.Left = 180
    picSample.Width = ScaleWidth - picSample.Left - cmdSelectFont.Width - 120 - Screen.TwipsPerPixelX
    picSample.CurrentY = 0
    picSample.CurrentX = 0
    If mEnabled Then
        UserControl.BackColor = mBackColor
        picSample.BackColor = mBackColor
        cmdSelectFont.BackColor = vbButtonFace
        picSample.ForeColor = mForeColor
    Else
        UserControl.BackColor = &HDCE4E7
        picSample.BackColor = &HDCE4E7
        cmdSelectFont.BackColor = &HDCE4E7
        picSample.ForeColor = vbGrayText
    End If
    
    If cmdSelectFont.Height > (UserControl.Height - Screen.TwipsPerPixelY * 8) Then
        cmdSelectFont.Height = UserControl.Height - Screen.TwipsPerPixelY * 8
        cmdSelectFont.Width = cmdSelectFont.Height
    Else
        If cmdSelectFont.Height < 300 Then
            If (UserControl.Height - Screen.TwipsPerPixelY * 8) > 300 Then
                cmdSelectFont.Height = 300
            Else
                cmdSelectFont.Height = UserControl.Height - Screen.TwipsPerPixelY * 8
            End If
            cmdSelectFont.Width = cmdSelectFont.Height
        End If
    End If
    
    If mSampleTextEnding = vxTEAddElipsis Then
        iSampleText = mSampleText
        If picSample.TextWidth(iSampleText) > picSample.ScaleWidth Then
            iSampleText = Left$(iSampleText, Len(iSampleText) - 1)
            Do Until picSample.TextWidth(iSampleText & "...") <= picSample.ScaleWidth
                iSampleText = Left$(iSampleText, Len(iSampleText) - 1)
            Loop
            iSampleText = iSampleText & "..."
        End If
        picSample.Print iSampleText
    Else
        picSample.Print mSampleText
    End If
    
    picSample.Visible = True
    
    If mVisualStyles Then
        If IsThemed Then
            iBorderColor = ThemeColor("TextBoxBorder")
        Else
            iBorderColor = mBorderColor
        End If
    Else
        iBorderColor = mBorderColor
    End If
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - Screen.TwipsPerPixelX, UserControl.ScaleHeight - Screen.TwipsPerPixelY), iBorderColor, B
    
    If mSampleTextEnding = vxTEVanish Then
        If picSample.TextWidth(mSampleText) > picSample.ScaleWidth Then
            picSample.ScaleMode = vbPixels
            TranslateColor mBackColor, 0&, iBackColor
            For iX = picSample.ScaleWidth - 26 To picSample.ScaleWidth - 1
                p = p + 4
                For iY = 0 To picSample.ScaleHeight - 1
                    iColor = GetPixel(picSample.hDC, iX, iY)
                    iColor = ColorsBlended(iColor, iBackColor, p)
                    SetPixel picSample.hDC, iX, iY, iColor
                Next iY
            Next iX
            picSample.ScaleMode = vbTwips
        End If
    End If
    
    cmdSelectFont.Enabled = mEnabled
    UserControl.Enabled = mEnabled
    picSample.Enabled = mEnabled
    
    picSample.ToolTipText = mFont.Name & " " & Round(mFont.Size) & IIf(mFont.Bold, " " & GetLocalizedString(efnGUIStr_FontPicker_DrawSample_Bold), "") & IIf(mFont.Italic, " " & GetLocalizedString(efnGUIStr_FontPicker_DrawSample_Italic), "")
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
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
    If mButtonToolTipText = GetLocalizedString(efnGUIStr_FontPicker_ButtonToolTipTextDefault, , nPrevLang) Then ButtonToolTipText = GetLocalizedString(efnGUIStr_FontPicker_ButtonToolTipTextDefault)
    DrawSample
End Sub
