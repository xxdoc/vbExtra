VERSION 5.00
Begin VB.Form frmPageNumbersOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "# Printing format"
   ClientHeight    =   3036
   ClientLeft      =   6288
   ClientTop       =   4752
   ClientWidth     =   5508
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPageNumbersOptions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3036
   ScaleWidth      =   5508
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboPageNumbersPosition 
      Height          =   288
      ItemData        =   "frmPageNumbersOptions.frx":000C
      Left            =   2760
      List            =   "frmPageNumbersOptions.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   2550
   End
   Begin VB.ComboBox cboPageNumbersFormat 
      Height          =   288
      ItemData        =   "frmPageNumbersOptions.frx":005E
      Left            =   2760
      List            =   "frmPageNumbersOptions.frx":0060
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   840
      Width           =   2550
   End
   Begin VB.Timer tmrInit 
      Interval        =   1
      Left            =   1440
      Top             =   2412
   End
   Begin VB.Timer tmrUnload 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1080
      Top             =   2412
   End
   Begin vbExtra.ButtonEx cmdOK_2 
      Height          =   372
      Left            =   240
      TabIndex        =   2
      Top             =   2424
      Visible         =   0   'False
      Width           =   348
      _ExtentX        =   614
      _ExtentY        =   656
      ShowFocusRect   =   -1  'True
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
      UseMaskCOlor    =   -1  'True
   End
   Begin vbExtra.ButtonEx cmdCancel_2 
      Height          =   372
      Left            =   636
      TabIndex        =   3
      Top             =   2424
      Visible         =   0   'False
      Width           =   348
      _ExtentX        =   614
      _ExtentY        =   656
      ShowFocusRect   =   -1  'True
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
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "# Cancel"
      Height          =   435
      Left            =   3690
      TabIndex        =   0
      Top             =   2424
      Width           =   1515
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "# OK"
      Height          =   435
      Left            =   1920
      TabIndex        =   1
      Top             =   2424
      Width           =   1515
   End
   Begin vbExtra.FontPicker fpcPageNumbers 
      Height          =   432
      Left            =   2772
      TabIndex        =   9
      Top             =   1452
      Width           =   2556
      _ExtentX        =   4509
      _ExtentY        =   762
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ChooseForeColor =   -1  'True
   End
   Begin VB.Label lblPageNumbersPosition 
      Alignment       =   1  'Right Justify
      Caption         =   "# Page numbers position:"
      Height          =   444
      Left            =   108
      TabIndex        =   8
      Top             =   396
      Width           =   2604
   End
   Begin VB.Label lblPageNumbersFormat 
      Alignment       =   1  'Right Justify
      Caption         =   "# Page numbers format:"
      Height          =   444
      Left            =   108
      TabIndex        =   7
      Top             =   900
      Width           =   2604
   End
   Begin VB.Label lblPageNumbersFont 
      Alignment       =   1  'Right Justify
      Caption         =   "# Page numbers font:"
      Height          =   372
      Left            =   108
      TabIndex        =   6
      Top             =   1476
      Width           =   2604
   End
End
Attribute VB_Name = "frmPageNumbersOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long

Private Enum eGreyMethods
    gITU = 0
    gNTSCPAL = 1
    gAverage = 2
    gWeighted = 3
    gVector = 4
    gEye = 5
    gAclarado = 6
End Enum

Private mPrintFnObject As PrintFnObject
Private mOKPressed As Boolean
Private mChanged As Boolean
Private mPrintPageNumbers As Boolean
Private mPageNumbersPosition As Long
Private mPageNumbersFormatLong As Long
Private mPageNumbersFont As StdFont
Private mPageNumbersForeColor As Long

Private mLoading As Boolean
Private mRgn As Long

Private Sub cboPageNumbersFormat_Click()
    mPageNumbersFormatLong = cboPageNumbersFormat.ListIndex
    fpcPageNumbers.SampleText = Replace(cboPageNumbersFormat.Text, " ", Chr(&HA0&))
    mChanged = True
End Sub

Private Sub cboPageNumbersPosition_Click()
    If cboPageNumbersPosition.ListIndex = cboPageNumbersPosition.ListCount - 1 Then
        mPrintPageNumbers = False
    Else
        mPrintPageNumbers = True
        mPageNumbersPosition = cboPageNumbersPosition.ItemData(cboPageNumbersPosition.ListIndex)
    End If
    mChanged = True
End Sub
Private Sub cmdCancel_2_Click()
    cmdCancel_Click
End Sub

Private Sub cmdOK_2_Click()
    cmdOK_Click
End Sub

Private Sub cmdOK_Click()
    mOKPressed = True
    mRgn = CreateRectRgn(0, 0, 0, 0)
    Call SetWindowRgn(Me.hWnd, mRgn, True)
    tmrUnload.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    mRgn = CreateRectRgn(0, 0, 0, 0)
    Call SetWindowRgn(Me.hWnd, mRgn, True)
    tmrUnload.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mRgn <> 0 Then
'        Call SetWindowRgn(Me.hWnd, 0, True)
        DeleteObject mRgn
        mRgn = 0
    End If
End Sub

Private Sub fpcPageNumbers_Change()
    Set mPageNumbersFont = fpcPageNumbers.Font
    mPageNumbersForeColor = fpcPageNumbers.ForeColor
    mChanged = True
End Sub

Private Sub Form_Load()
    Dim iPt As POINTAPI
    
    mLoading = True
    LoadGUICaptions
    
    GetCursorPos iPt
    iPt.x = iPt.x - 15
    If iPt.x < 10 Then iPt.x = 10
    iPt.y = iPt.y + 20
    Me.Move ScaleX(iPt.x, vbPixels, ScaleMode), ScaleY(iPt.y, vbPixels, ScaleMode)
    PersistForm Me, Forms, False
    Set mPageNumbersFont = New StdFont
    LoadPageNumbersFormatStrings
    
    If gButtonsStyle <> -1 Then
        cmdOK_2.Move cmdOK.Left, cmdOK.Top, cmdOK.Width, cmdOK.Height
        cmdOK_2.Caption = cmdOK.Caption
        cmdOK.Visible = False
        cmdOK_2.Default = cmdOK.Default
        cmdOK_2.Cancel = cmdOK.Cancel
        cmdOK_2.Visible = True
        cmdOK_2.TabIndex = cmdOK.TabIndex
        cmdOK_2.ButtonStyle = gButtonsStyle
    
        cmdCancel_2.Move cmdCancel.Left, cmdCancel.Top, cmdCancel.Width, cmdCancel.Height
        cmdCancel_2.Caption = cmdCancel.Caption
        cmdCancel.Visible = False
        cmdCancel_2.Default = cmdCancel.Default
        cmdCancel_2.Cancel = cmdCancel.Cancel
        cmdCancel_2.Visible = True
        cmdCancel_2.TabIndex = cmdCancel.TabIndex
        cmdCancel_2.ButtonStyle = gButtonsStyle
        
    End If
End Sub

Public Property Get OKPressed() As Boolean
    OKPressed = mOKPressed
End Property

Public Property Get Changed() As Boolean
    Changed = mChanged
End Property


Public Property Let PrintPageNumbers(nValue As Boolean)
    mPrintPageNumbers = nValue
    
    If mPrintPageNumbers Then
        If cboPageNumbersPosition.ListIndex = cboPageNumbersPosition.ListCount - 1 Then
            SelectInComboByItemData cboPageNumbersPosition, Val(cboPageNumbersPosition.Tag)
        End If
    Else
        cboPageNumbersPosition.ListIndex = cboPageNumbersPosition.ListCount - 1
    End If
End Property

Public Property Get PrintPageNumbers() As Boolean
    PrintPageNumbers = mPrintPageNumbers
End Property


Public Property Let PageNumbersPosition(nValue As Long)
    mPageNumbersPosition = nValue
    
    If mPrintPageNumbers Then
        SelectInComboByItemData cboPageNumbersPosition, mPageNumbersPosition
        cboPageNumbersPosition.Tag = mPageNumbersPosition
    End If

End Property

Public Property Get PageNumbersPosition() As Long
    PageNumbersPosition = mPageNumbersPosition
End Property


Public Property Let PageNumbersFormatLong(nValue As Long)
    mPageNumbersFormatLong = nValue
    
    If (mPageNumbersFormatLong >= 0) And (mPageNumbersFormatLong < cboPageNumbersFormat.ListCount) Then
        cboPageNumbersFormat.ListIndex = mPageNumbersFormatLong
    End If
    
End Property

Public Property Get PageNumbersFormatLong() As Long
    PageNumbersFormatLong = mPageNumbersFormatLong
End Property


Public Property Set PageNumbersFont(ByVal nValue As StdFont)
    Set mPageNumbersFont = nValue
    Set fpcPageNumbers.Font = nValue
End Property

Public Property Get PageNumbersFont() As StdFont
    Set PageNumbersFont = mPageNumbersFont
End Property


Public Property Let PageNumbersForeColor(nValue As OLE_COLOR)
    mPageNumbersForeColor = nValue
    fpcPageNumbers.ForeColor = mPageNumbersForeColor
End Property

Public Property Get PageNumbersForeColor() As OLE_COLOR
    PageNumbersForeColor = mPageNumbersForeColor
End Property


Public Property Set PrintFnObject(nPrintFnObject As Object)
    Set mPrintFnObject = nPrintFnObject
End Property


Private Sub LoadPageNumbersFormatStrings()
    Dim c As Long
    
    cboPageNumbersFormat.Clear
    For c = 0 To mPrintFnObject.GetPredefinedPageNumbersFormatStringsCount - 1
        cboPageNumbersFormat.AddItem PrinterExCurrentDocument.GetFormattedPageNumberString(mPrintFnObject.GetPredefinedPageNumbersFormatString(c), 10, 30)
    Next c
End Sub

Private Sub tmrInit_Timer()
    tmrInit.Enabled = False
    mChanged = False
End Sub

Private Sub tmrUnload_Timer()
    tmrUnload.Enabled = False
    Unload Me
End Sub

Private Sub LoadGUICaptions()
    Dim c As Long
    Dim iSkip As Boolean
    
    Me.Caption = GetLocalizedString(efnGUIStr_frmPageNumbersOptions_Caption)
    cmdOK.Caption = GetLocalizedString(efnGUIStr_General_OKButton_Caption)
    cmdCancel.Caption = GetLocalizedString(efnGUIStr_General_CancelButton_Caption)
    
    lblPageNumbersFont.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersFont_Caption)
    lblPageNumbersFormat.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersFormat_Caption)
    lblPageNumbersPosition.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersPosition_Caption)
    
    cboPageNumbersPosition.Clear
    For c = 0 To 6
        iSkip = False
        If c = 2 Then ' bottom centered
            If Not gPrinterExPageFixedElementPositionTop Then
                If Not gPrinterExPageFixedImage Is Nothing Or (gPrinterExPageFixedText <> "") Then
                    iSkip = True
                End If
            End If
        ElseIf c = 5 Then ' top centered
            If gPrinterExPageFixedElementPositionTop Then
                If Not gPrinterExPageFixedImage Is Nothing Or (gPrinterExPageFixedText <> "") Then
                    iSkip = True
                End If
            End If
        End If
        If Not iSkip Then
            cboPageNumbersPosition.AddItem GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_cboPageNumbersPosition_List, c)
            cboPageNumbersPosition.ItemData(cboPageNumbersPosition.NewIndex) = c
        End If
    Next c
    
End Sub

