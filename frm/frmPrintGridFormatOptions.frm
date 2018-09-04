VERSION 5.00
Begin VB.Form frmPrintGridFormatOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "# Printing format"
   ClientHeight    =   6300
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
   Icon            =   "frmPrintGridFormatOptions.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   5508
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrInit 
      Interval        =   1
      Left            =   1440
      Top             =   5688
   End
   Begin VB.Timer tmrUnload 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1080
      Top             =   5688
   End
   Begin vbExtra.ButtonEx cmdOK_2 
      Height          =   375
      Left            =   240
      TabIndex        =   44
      Top             =   5700
      Visible         =   0   'False
      Width           =   345
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
      Height          =   375
      Left            =   630
      TabIndex        =   45
      Top             =   5700
      Visible         =   0   'False
      Width           =   345
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
   Begin vbExtra.SSTabEx sst1 
      Height          =   5445
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   5415
      _ExtentX        =   9546
      _ExtentY        =   9610
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   4
      Tab             =   2
      TabHeight       =   520
      Themed          =   -1  'True
      TabCaption(0)   =   "# Options"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "picOptionsContainer"
      TabCaption(1)   =   "# Style"
      Tab(1).ControlCount=   1
      Tab(1).Control(0)=   "picStyleContainer"
      TabCaption(2)   =   "# More"
      Tab(2).ControlCount=   1
      Tab(2).Control(0)=   "chkEnableAutoOrientation"
      Begin VB.CheckBox chkEnableAutoOrientation 
         Caption         =   "# Automatically change the page orientation to horizontal if the report is wider than the paper."
         Height          =   984
         Left            =   504
         TabIndex        =   41
         Top             =   648
         Value           =   1  'Checked
         Width           =   4416
      End
      Begin VB.PictureBox picStyleContainer 
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   -74820
         ScaleHeight     =   4932
         ScaleWidth      =   5100
         TabIndex        =   43
         Top             =   390
         Width           =   5100
         Begin VB.PictureBox picSample 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   2700
            Left            =   0
            ScaleHeight     =   2700
            ScaleWidth      =   4704
            TabIndex        =   48
            Top             =   2160
            Width           =   4700
         End
         Begin VB.PictureBox picPrintHeadersSeparatorLine 
            BorderStyle     =   0  'None
            Height          =   372
            Left            =   2592
            ScaleHeight     =   372
            ScaleWidth      =   1776
            TabIndex        =   46
            Top             =   1620
            Width           =   1776
            Begin VB.CheckBox chkPrintHeadersSeparatorLine 
               Caption         =   "# Headers Sep."
               Height          =   375
               Left            =   96
               TabIndex        =   47
               Top             =   0
               Width           =   1620
            End
         End
         Begin VB.TextBox txtLineWidthHeadersSeparatorLine 
            Height          =   300
            Left            =   4400
            TabIndex        =   39
            Text            =   "1"
            ToolTipText     =   "# Headers separator line thickness"
            Top             =   1650
            Width           =   315
         End
         Begin VB.TextBox txtLineWidth 
            Height          =   300
            Left            =   1980
            TabIndex        =   24
            Text            =   "1"
            ToolTipText     =   "# Change line thickness (general)"
            Top             =   660
            Width           =   315
         End
         Begin vbExtra.ButtonEx cmdHeadersBorderColor2 
            Height          =   300
            Left            =   4760
            TabIndex        =   40
            ToolTipText     =   "# Change color"
            Top             =   1650
            Width           =   300
            _ExtentX        =   402
            _ExtentY        =   402
            ButtonStyle     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "..."
         End
         Begin vbExtra.ButtonEx cmdHeadersBackgroundColor 
            Height          =   300
            Left            =   1980
            TabIndex        =   27
            ToolTipText     =   "# Change background color of headers (and / or fixed columns)"
            Top             =   1128
            Width           =   300
            _ExtentX        =   402
            _ExtentY        =   402
            ButtonStyle     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "..."
         End
         Begin vbExtra.ButtonEx cmdColumnsHeadersLinesColor 
            Height          =   300
            Left            =   4760
            TabIndex        =   36
            ToolTipText     =   "# Change color"
            Top             =   990
            Width           =   300
            _ExtentX        =   402
            _ExtentY        =   402
            ButtonStyle     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "..."
         End
         Begin VB.CheckBox chkPrintColumnsHeadersLines 
            Caption         =   "# Col. headers lines"
            Height          =   375
            Left            =   2670
            TabIndex        =   35
            Top             =   990
            Width           =   1980
         End
         Begin VB.CheckBox chkPrintHeadersBorder 
            Caption         =   "# Headers borders"
            Height          =   375
            Left            =   2670
            TabIndex        =   31
            Top             =   330
            Width           =   1980
         End
         Begin VB.ComboBox cboStyle 
            Height          =   300
            ItemData        =   "frmPrintGridFormatOptions.frx":000C
            Left            =   480
            List            =   "frmPrintGridFormatOptions.frx":0019
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   120
            Width           =   1812
         End
         Begin VB.CheckBox chkPrintFixedColsBackground 
            Caption         =   "# Fixed columns Bk."
            Height          =   375
            Left            =   0
            TabIndex        =   26
            Top             =   1320
            Width           =   1908
         End
         Begin vbExtra.ButtonEx cmdHeadersBorderColor 
            Height          =   300
            Left            =   4760
            TabIndex        =   32
            ToolTipText     =   "# Change color"
            Top             =   330
            Width           =   300
            _ExtentX        =   402
            _ExtentY        =   402
            ButtonStyle     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "..."
         End
         Begin vbExtra.ButtonEx cmdRowsLinesColor 
            Height          =   300
            Left            =   4760
            TabIndex        =   38
            ToolTipText     =   "# Change color"
            Top             =   1320
            Width           =   300
            _ExtentX        =   402
            _ExtentY        =   402
            ButtonStyle     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "..."
         End
         Begin vbExtra.ButtonEx cmdColumnsDataLinesColor 
            Height          =   300
            Left            =   4760
            TabIndex        =   34
            ToolTipText     =   "# Change color"
            Top             =   660
            Width           =   300
            _ExtentX        =   402
            _ExtentY        =   402
            ButtonStyle     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "..."
         End
         Begin vbExtra.ButtonEx cmdOuterBorderColor 
            Height          =   300
            Left            =   4760
            TabIndex        =   30
            ToolTipText     =   "# Change color"
            Top             =   0
            Width           =   300
            _ExtentX        =   402
            _ExtentY        =   402
            ButtonStyle     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "..."
         End
         Begin VB.CheckBox chkPrintHeadersBackground 
            Caption         =   "# Headers Backgr."
            Height          =   375
            Left            =   0
            TabIndex        =   25
            Top             =   990
            Width           =   1980
         End
         Begin VB.CheckBox chkPrintOtherBackgrounds 
            Caption         =   "# Other Backgr. colors"
            Height          =   375
            Left            =   0
            TabIndex        =   28
            Top             =   1640
            Width           =   2316
         End
         Begin VB.CheckBox chkPrintRowsLines 
            Caption         =   "# Row lines"
            Height          =   375
            Left            =   2670
            TabIndex        =   37
            Top             =   1320
            Width           =   1980
         End
         Begin VB.CheckBox chkPrintColumnsDataLines 
            Caption         =   "# Columns data Lin."
            Height          =   375
            Left            =   2670
            TabIndex        =   33
            Top             =   660
            Width           =   1980
         End
         Begin VB.CheckBox chkPrintOuterBorder 
            Caption         =   "# Outer edge"
            Height          =   375
            Left            =   2670
            TabIndex        =   29
            Top             =   30
            Width           =   1980
         End
         Begin VB.Label lblSample 
            Caption         =   "Ejemplo:"
            Height          =   408
            Left            =   720
            TabIndex        =   49
            Tag             =   "na"
            Top             =   1620
            Visible         =   0   'False
            Width           =   1416
         End
         Begin VB.Label lblLineWidth 
            Caption         =   "# Lines width:"
            Height          =   270
            Left            =   0
            TabIndex        =   23
            Top             =   700
            Width           =   1416
         End
         Begin VB.Label lblStyle 
            Caption         =   "#Style"
            Height          =   195
            Left            =   0
            TabIndex        =   21
            Top             =   180
            Width           =   495
         End
      End
      Begin VB.PictureBox picOptionsContainer 
         BorderStyle     =   0  'None
         Height          =   4932
         Left            =   -74940
         ScaleHeight     =   4932
         ScaleWidth      =   5244
         TabIndex        =   42
         Top             =   420
         Width           =   5240
         Begin vbExtra.FontPicker fpcHeading 
            Height          =   432
            Left            =   2650
            TabIndex        =   16
            Top             =   3216
            Width           =   2550
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
         Begin VB.ComboBox cboPageNumbersFormat 
            Height          =   300
            ItemData        =   "frmPrintGridFormatOptions.frx":003F
            Left            =   2650
            List            =   "frmPrintGridFormatOptions.frx":0041
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1560
            Width           =   2550
         End
         Begin VB.ComboBox cboPageNumbersPosition 
            Height          =   300
            ItemData        =   "frmPrintGridFormatOptions.frx":0043
            Left            =   2650
            List            =   "frmPrintGridFormatOptions.frx":0053
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1080
            Width           =   2550
         End
         Begin VB.ComboBox cboGridAlign 
            Height          =   300
            ItemData        =   "frmPrintGridFormatOptions.frx":0095
            Left            =   2650
            List            =   "frmPrintGridFormatOptions.frx":00A5
            Style           =   2  'Dropdown List
            TabIndex        =   14
            ToolTipText     =   "# It only has effect when the data grid is narrower than the page"
            Top             =   2700
            Width           =   2550
         End
         Begin VB.ComboBox cboColor 
            Height          =   300
            ItemData        =   "frmPrintGridFormatOptions.frx":00D1
            Left            =   2650
            List            =   "frmPrintGridFormatOptions.frx":00DE
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   600
            Width           =   2550
         End
         Begin VB.ComboBox cboScalePercent 
            Height          =   300
            ItemData        =   "frmPrintGridFormatOptions.frx":010C
            Left            =   2650
            List            =   "frmPrintGridFormatOptions.frx":0125
            TabIndex        =   4
            Text            =   "cboScalePercent"
            Top             =   108
            Width           =   2550
         End
         Begin vbExtra.FontPicker fpcPageNumbers 
            Height          =   432
            Left            =   2650
            TabIndex        =   12
            Top             =   2076
            Width           =   2550
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
         Begin vbExtra.FontPicker fpcSubheading 
            Height          =   432
            Left            =   2650
            TabIndex        =   18
            Top             =   3840
            Width           =   2550
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
         Begin vbExtra.FontPicker fpcOtherTexts 
            Height          =   432
            Left            =   2650
            TabIndex        =   20
            Top             =   4476
            Width           =   2550
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
         Begin VB.Label lblOtherTextsFont 
            Alignment       =   1  'Right Justify
            Caption         =   "# Other texts font:"
            Height          =   336
            Left            =   0
            TabIndex        =   19
            Top             =   4608
            Width           =   2600
         End
         Begin VB.Label lblSubheadingFont 
            Alignment       =   1  'Right Justify
            Caption         =   "# Sub-heading font:"
            Height          =   336
            Left            =   0
            TabIndex        =   17
            Top             =   3960
            Width           =   2600
         End
         Begin VB.Label lblHeadingFont 
            Alignment       =   1  'Right Justify
            Caption         =   "# Heading font:"
            Height          =   336
            Left            =   0
            TabIndex        =   15
            Top             =   3348
            Width           =   2600
         End
         Begin VB.Label lblPageNumbersFont 
            Alignment       =   1  'Right Justify
            Caption         =   "# Page numbers font:"
            Height          =   372
            Left            =   0
            TabIndex        =   11
            Top             =   2196
            Width           =   2600
         End
         Begin VB.Label lblPageNumbersFormat 
            Alignment       =   1  'Right Justify
            Caption         =   "# Page numbers format:"
            Height          =   444
            Left            =   0
            TabIndex        =   9
            Top             =   1620
            Width           =   2600
         End
         Begin VB.Label lblPageNumbersPosition 
            Alignment       =   1  'Right Justify
            Caption         =   "# Page numbers position:"
            Height          =   444
            Left            =   0
            TabIndex        =   7
            Top             =   1116
            Width           =   2600
         End
         Begin VB.Label lblGridAlign 
            Alignment       =   1  'Right Justify
            Caption         =   "# Grid alignment:"
            Height          =   372
            Left            =   0
            TabIndex        =   13
            Top             =   2736
            Width           =   2600
         End
         Begin VB.Label lblColor 
            Alignment       =   1  'Right Justify
            Caption         =   "# Color:"
            Height          =   444
            Left            =   0
            TabIndex        =   5
            Top             =   648
            Width           =   2600
         End
         Begin VB.Label lblScalePercent 
            Alignment       =   1  'Right Justify
            Caption         =   "# Scale:"
            Height          =   444
            Left            =   0
            TabIndex        =   3
            Top             =   180
            Width           =   2600
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "# Cancel"
      Height          =   435
      Left            =   3690
      TabIndex        =   0
      Top             =   5700
      Width           =   1515
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "# OK"
      Height          =   435
      Left            =   1920
      TabIndex        =   1
      Top             =   5700
      Width           =   1515
   End
End
Attribute VB_Name = "frmPrintGridFormatOptions"
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

Private mOKPressed As Boolean
Private mChanged As Boolean
Private mGrid As Control
Private mPrintFnObject As PrintFnObject
Attribute mPrintFnObject.VB_VarHelpID = -1
Private mFlexFnObject As FlexFnObject

Private mScalePercent As Long
Private mMinScalePercent As Long
Private mMaxScalePercent As Long
Private mColorMode As cdeColorModeConstants
Private mGridAlign As Long
Private mPrintPageNumbers As Boolean
Private mPageNumbersPosition As Long
Private mPageNumbersFormatLong As Long
Private mPageNumbersFont As StdFont
Private mPageNumbersForeColor As Long
Private mHeadingFont As StdFont
Private mHeadingFontColor As Long
Private mSubheadingFont As StdFont
Private mSubheadingFontColor As Long
Private mOtherTextsFont As StdFont
Private mOtherTextsFontColor As Long
Private mHeadingSampleText As String
Private mSubheadingSampleText As String
Private mOtherTextsSampleText As String
Private mDisplayHeadingFont As Boolean
Private mDisplaySubheadingFont As Boolean
Private mDisplayOtherTextsFont As Boolean

Private mGridReportStyle As New GridReportStyle

Private mEnableAutoOrientation As Boolean

Private mLoading As Boolean
Private mSelectingProperStyle As Boolean
Private mNewCustomStyle As Boolean
Private mPuttingControlsToStyle As Boolean
Private mStylesIDs() As String
Private mSampleTop As Long
Private mStyleChanged As Boolean
Private mRgn As Long

Private Sub cboColor_Click()
    mColorMode = cboColor.ListIndex
    DrawSample
    mChanged = True
End Sub

Private Sub cboScalePercent_Change()
    mChanged = True
End Sub

Private Sub cboScalePercent_Click()
    mScalePercent = Val(cboScalePercent.List(cboScalePercent.ListIndex))
    mChanged = True
End Sub

Private Sub cboScalePercent_KeyPress(KeyAscii As Integer)
    Dim iVal As Long
    
    If KeyAscii = 13 Then
        iVal = Val(cboScalePercent.Text)
        If (iVal < mMinScalePercent) Then
            cboScalePercent.Text = mMinScalePercent
        End If
        If (iVal > mMaxScalePercent) Then
            cboScalePercent.Text = mMaxScalePercent
        End If
        cboScalePercent.Refresh
        cboScalePercent_Click
    End If
End Sub

Private Sub cboGridAlign_Click()
    mGridAlign = cboGridAlign.ListIndex
    mChanged = True
End Sub

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

Private Sub cboStyle_Click()
    If cboStyle.ListIndex < cboStyle.ListCount - 1 Then
        SetContainerElementsVisibility picStyleContainer, False
        picSample.Visible = True
        lblSample.Visible = True
        lblStyle.Visible = True
        cboStyle.Visible = True
        mSampleTop = 1350
        lblSample.Top = mSampleTop - lblSample.Height
        lblStyle.Move 480, 280
        lblSample.Left = lblStyle.Left
        cboStyle.Move 960, 220
    Else
        SetContainerElementsVisibility picStyleContainer, True
        lblSample.Visible = False
        mSampleTop = 2150
        lblStyle.Move 0, 180
        cboStyle.Move 480, 120
    End If
    If Not mPuttingControlsToStyle And Not mSelectingProperStyle Then
        If cboStyle.ListIndex < cboStyle.ListCount - 1 Then
            Set mGridReportStyle = mFlexFnObject.GetGridReportStyle(mStylesIDs(cboStyle.ListIndex))
            StyleChanged = True
            PutControlsToStyle
        End If
    End If
    DrawSample
    mChanged = True
End Sub

Private Sub chkEnableAutoOrientation_Click()
    mEnableAutoOrientation = (chkEnableAutoOrientation.Value = 1)
    mChanged = True
End Sub

Private Sub chkPrintFixedColsBackground_Click()
    mGridReportStyle.PrintFixedColsBackground = (chkPrintFixedColsBackground = 1)
    DrawSample
    StyleChanged = True
    mChanged = True
End Sub

Private Sub chkPrintHeadersBackground_Click()
    mGridReportStyle.PrintHeadersBackground = (chkPrintHeadersBackground.Value = 1)
    DrawSample
    StyleChanged = True
    mChanged = True
End Sub

Private Sub chkPrintHeadersBorder_Click()
    mGridReportStyle.PrintHeadersBorder = (chkPrintHeadersBorder.Value = 1)
    DrawSample
    StyleChanged = True
    picPrintHeadersSeparatorLine.Enabled = (chkPrintRowsLines.Value = 0) And ((chkPrintHeadersBorder.Value = 0) Or Not chkPrintHeadersBorder.Enabled)
    If Not picPrintHeadersSeparatorLine.Enabled Then
        chkPrintHeadersSeparatorLine.Tag = chkPrintHeadersSeparatorLine.Value
        chkPrintHeadersSeparatorLine.Value = 2
    Else
        If chkPrintHeadersSeparatorLine.Value = 2 Then
            chkPrintHeadersSeparatorLine.Value = Val(chkPrintHeadersSeparatorLine.Tag)
        End If
    End If
    mChanged = True
End Sub

Private Sub chkPrintColumnsHeadersLines_Click()
    mGridReportStyle.PrintColumnsHeadersLines = (chkPrintColumnsHeadersLines.Value = 1)
    DrawSample
    StyleChanged = True
    mChanged = True
End Sub

Private Sub chkPrintOtherBackgrounds_Click()
    mGridReportStyle.PrintOtherBackgrounds = (chkPrintOtherBackgrounds.Value = 1)
    DrawSample
    StyleChanged = True
    mChanged = True
End Sub

Private Sub chkPrintOuterBorder_Click()
    mGridReportStyle.PrintOuterBorder = (chkPrintOuterBorder.Value = 1)
    DrawSample
    StyleChanged = True
    mChanged = True
End Sub

Private Sub chkPrintHeadersSeparatorLine_Click()
    mGridReportStyle.PrintHeadersSeparatorLine = (chkPrintHeadersSeparatorLine.Value <> 0)
    DrawSample
    StyleChanged = True
    mChanged = True
End Sub

Private Sub chkPrintColumnsDataLines_Click()
    mGridReportStyle.PrintColumnsDataLines = (chkPrintColumnsDataLines.Value = 1)
    DrawSample
    StyleChanged = True
    mChanged = True
End Sub

Private Sub chkPrintRowsLines_Click()
    mGridReportStyle.PrintRowsLines = (chkPrintRowsLines.Value = 1)
    DrawSample
    StyleChanged = True
    picPrintHeadersSeparatorLine.Enabled = (chkPrintRowsLines.Value = 0) And ((chkPrintHeadersBorder.Value = 0) Or Not chkPrintHeadersBorder.Enabled)
    If Not picPrintHeadersSeparatorLine.Enabled Then
        chkPrintHeadersSeparatorLine.Tag = chkPrintHeadersSeparatorLine.Value
        chkPrintHeadersSeparatorLine.Value = 2
    Else
        If chkPrintHeadersSeparatorLine.Value = 2 Then
            chkPrintHeadersSeparatorLine.Value = Val(chkPrintHeadersSeparatorLine.Tag)
        End If
    End If
    mChanged = True
End Sub

Private Sub cmdCancel_2_Click()
    cmdCancel_Click
End Sub

Private Sub cmdColumnsDataLinesColor_Click()
    Dim iDlg As New CommonDialogExObject
    
    iDlg.Color = mGridReportStyle.ColumnsDataLinesColor
    iDlg.ShowColor cdeCCFullOpen
    If Not iDlg.Canceled Then
        mGridReportStyle.ColumnsDataLinesColor = iDlg.Color
        DrawSample
        StyleChanged = True
        mChanged = True
    End If
End Sub

Private Sub cmdHeadersBorderColor2_Click()
    cmdHeadersBorderColor_Click
End Sub

Private Sub cmdColumnsHeadersLinesColor_Click()
    Dim iDlg As New CommonDialogExObject
    
    iDlg.Color = mGridReportStyle.ColumnsHeadersLinesColor
    iDlg.ShowColor cdeCCFullOpen
    If Not iDlg.Canceled Then
        mGridReportStyle.ColumnsHeadersLinesColor = iDlg.Color
        DrawSample
        StyleChanged = True
        mChanged = True
    End If
End Sub

Private Sub cmdHeadersBackgroundColor_Click()
    Dim iDlg As New CommonDialogExObject
    
    iDlg.Color = mGridReportStyle.HeadersBackgroundColor
    iDlg.ShowColor cdeCCFullOpen
    If Not iDlg.Canceled Then
        mGridReportStyle.HeadersBackgroundColor = iDlg.Color
        DrawSample
        StyleChanged = True
        mChanged = True
    End If
End Sub

Private Sub cmdHeadersBorderColor_Click()
    Dim iDlg As New CommonDialogExObject
    
    iDlg.Color = mGridReportStyle.HeadersBorderColor
    iDlg.ShowColor cdeCCFullOpen
    If Not iDlg.Canceled Then
        mGridReportStyle.HeadersBorderColor = iDlg.Color
        DrawSample
        StyleChanged = True
        mChanged = True
    End If
End Sub

Private Sub cmdOK_2_Click()
    cmdOK_Click
End Sub

Private Sub cmdOK_Click()
    Dim iVal As Long
    
    iVal = Val(cboScalePercent.Text)
    
    If (iVal > 30) And (iVal < 300) Then
        mScalePercent = iVal
    End If
    
    mOKPressed = True
    If StyleChanged Then
        mGridReportStyle.Tag = "Save"
    End If
    
    mRgn = CreateRectRgn(0, 0, 0, 0)
    Call SetWindowRgn(Me.hWnd, mRgn, True)
    tmrUnload.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    mRgn = CreateRectRgn(0, 0, 0, 0)
    Call SetWindowRgn(Me.hWnd, mRgn, True)
    tmrUnload.Enabled = True
End Sub

Private Sub cmdOuterBorderColor_Click()
    Dim iDlg As New CommonDialogExObject
    
    iDlg.Color = mGridReportStyle.OuterBorderColor
    iDlg.ShowColor cdeCCFullOpen
    If Not iDlg.Canceled Then
        mGridReportStyle.OuterBorderColor = iDlg.Color
        DrawSample
        StyleChanged = True
        mChanged = True
    End If
End Sub


Private Sub cmdRowsLinesColor_Click()
    Dim iDlg As New CommonDialogExObject
    
    iDlg.Color = mGridReportStyle.RowsLinesColor
    iDlg.ShowColor cdeCCFullOpen
    If Not iDlg.Canceled Then
        mGridReportStyle.RowsLinesColor = iDlg.Color
        DrawSample
        StyleChanged = True
        mChanged = True
    End If
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
    
    GetCursorPos iPt
    iPt.x = iPt.x - 15
    If iPt.x < 10 Then iPt.x = 10
    iPt.y = iPt.y + 20
    Me.Move ScaleX(iPt.x, vbPixels, ScaleMode), ScaleY(iPt.y, vbPixels, ScaleMode)
    
    mLoading = True
    LoadGUICaptions
    
    mSampleTop = 2150
    
    sst1.TabSel = 0
    
    SetTextBoxNumeric txtLineWidth
    SetTextBoxNumeric txtLineWidthHeadersSeparatorLine
    
    PersistForm Me, Forms
    CreateFonts
    LoadPageNumbersFormatStrings
    LoadDefaultSettings
    LoadStyles
    UpdatecboScalePercentList
    
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
        
        sst1.VisualStyles = False
    End If
    sst1.TabVisible(2) = False
End Sub

Private Sub DrawSample()
    Dim c As Long
    Dim R As Long
    Dim d As Long
    Dim iHeadersBackgroundColor As Long
    Dim iOtherBackgroundColor As Long
    Dim iOuterBorderColor As Long
    Dim iColumnsDataLinesColor As Long
    Dim iColumnsHeadersLinesColor As Long
    Dim iRowsLinesColor As Long
    Dim iHeadersBorderColor As Long
    Dim iLng As Long
    Dim iLng2 As Long
    Dim iPic As StdPicture
    Dim iColorMode As Boolean
    
    If mLoading Or mPuttingControlsToStyle Then Exit Sub
    
    If mColorMode > 0 Then
        iColorMode = (mColorMode = vbPRCMColor)
    Else
        If Not PrinterExCurrentDocument Is Nothing Then
            iColorMode = PrinterExCurrentDocument.DefaultColorMode = vbPRCMColor
        Else
            iColorMode = True
        End If
    End If
    
'    TranslateColor mGrid.BackColorFixed, 0, iHeadersBackgroundColor
    iOtherBackgroundColor = &HFFFFE1
    If Not iColorMode Then
        iHeadersBackgroundColor = ToGrey(iHeadersBackgroundColor)
        iOtherBackgroundColor = ToGrey(iOtherBackgroundColor)
        iOuterBorderColor = ToGrey(mGridReportStyle.OuterBorderColor)
        iColumnsDataLinesColor = ToGrey(mGridReportStyle.ColumnsDataLinesColor)
        iColumnsHeadersLinesColor = ToGrey(mGridReportStyle.ColumnsHeadersLinesColor)
        iRowsLinesColor = ToGrey(mGridReportStyle.RowsLinesColor)
        iHeadersBorderColor = ToGrey(mGridReportStyle.HeadersBorderColor)
        iHeadersBackgroundColor = ToGrey(mGridReportStyle.HeadersBackgroundColor)
    Else
        iOuterBorderColor = mGridReportStyle.OuterBorderColor
        iColumnsDataLinesColor = mGridReportStyle.ColumnsDataLinesColor
        iColumnsHeadersLinesColor = mGridReportStyle.ColumnsHeadersLinesColor
        iRowsLinesColor = mGridReportStyle.RowsLinesColor
        iHeadersBorderColor = mGridReportStyle.HeadersBorderColor
        iHeadersBackgroundColor = mGridReportStyle.HeadersBackgroundColor
    End If
    
  '  SetWindowRedraw picSample.hWnd, False
    picSample.Cls
    picSample.Line (0, 0)-(picSample.ScaleWidth - 1, picSample.ScaleHeight - 1), vbButtonFace, BF
    iLng = Round(mGridReportStyle.LineWidth / 4)
    If iLng = 0 Then iLng = 1
    picSample.DrawWidth = iLng
    
    picSample.Top = mSampleTop
    
    ' paper shadow
    picSample.ForeColor = RGB(166, 166, 166)
    picSample.Line (4615, 70)-(4670, 2700), , BF
    
    ' paper
    picSample.ForeColor = vbWhite
    picSample.Line (500, 0)-(4600, 2700), , BF
    
    ' header background
    If mGridReportStyle.PrintHeadersBackground Then
        picSample.Line (700, 250)-(4400, 800), iHeadersBackgroundColor, BF
    End If
    
    ' fixed rows background
    If mGridReportStyle.PrintFixedColsBackground Then
        picSample.Line (700, 800)-(1900, 2450), iHeadersBackgroundColor, BF
    End If
    
    ' other backgrounds
    If mGridReportStyle.PrintOtherBackgrounds Then
        picSample.Line (1900, 1900)-(3100, 2450), iOtherBackgroundColor, BF
    End If
    
    ' lines
    If mGridReportStyle.PrintColumnsDataLines Then
        picSample.Line (1900, 800)-(1900, 2700), iColumnsDataLinesColor
        picSample.Line (3100, 800)-(3100, 2700), iColumnsDataLinesColor
    End If
    
    If mGridReportStyle.PrintColumnsHeadersLines Then
        picSample.Line (1900, 250)-(1900, 800), iColumnsHeadersLinesColor
        picSample.Line (3100, 250)-(3100, 800), iColumnsHeadersLinesColor
    End If
    
    If mGridReportStyle.PrintRowsLines Then
        picSample.Line (700, 1350)-(4400, 1350), iRowsLinesColor
        picSample.Line (700, 1900)-(4400, 1900), iRowsLinesColor
        picSample.Line (700, 2450)-(4400, 2450), iRowsLinesColor
    End If
    If mGridReportStyle.PrintHeadersSeparatorLine Then
        iLng = Round(mGridReportStyle.LineWidthHeadersSeparatorLine / 4)
        If iLng = 0 Then iLng = 1
        If iLng < 3 Then
            picSample.DrawWidth = iLng
            picSample.Line (700, 800)-(4400, 800), iHeadersBorderColor
        Else
            picSample.DrawWidth = 1
            iLng2 = (iLng - 1) * Screen.TwipsPerPixelY / 2
            picSample.Line (700, 800 - iLng2)-(4400, 800 + iLng2), iHeadersBorderColor, BF
        End If
        iLng = Round(mGridReportStyle.LineWidth / 4)
        If iLng = 0 Then iLng = 1
        picSample.DrawWidth = iLng
    End If
    
    If mGridReportStyle.PrintHeadersBorder Then
        picSample.Line (700, 250)-(4400, 800), iHeadersBorderColor, B
    End If
    
    If mGridReportStyle.PrintOuterBorder Then
        If mGridReportStyle.PrintHeadersBorder Then
            picSample.Line (700, 800)-(4400, 2700), iOuterBorderColor, B
            picSample.Line (700, 800)-(4400, 800), iHeadersBorderColor
        Else
            picSample.Line (700, 250)-(4400, 2700), iOuterBorderColor, B
        End If
    End If
    
    picSample.FillColor = RGB(216, 216, 216)
    picSample.ForeColor = vbWhite
    picSample.FillStyle = 7
    picSample.Line (500, 2650)-(4600, 2700), , B
    picSample.FillStyle = 1
    
    picSample.ForeColor = vbBlack
    picSample.FontBold = True
    For c = 1 To 3
        picSample.CurrentX = 880 + 1200 * (c - 1)
        picSample.CurrentY = 430
        picSample.Print GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_DrawSample_Column) & " " & c
    Next c
    
    d = 0
    picSample.FontBold = False
    For c = 1 To 3
        For R = 1 To 3
            d = d + 1
            If iColorMode Then
                picSample.ForeColor = vbBlack
                If c = 3 Then
                    If R = 2 Then
                        picSample.ForeColor = vbRed
                    ElseIf R = 3 Then
                        picSample.ForeColor = vbBlue
                    End If
                End If
            End If
            picSample.CurrentY = 430 + 550 * R
            If c = 1 Then
                picSample.CurrentX = 880 + 1200 * (c - 1)
                picSample.Print GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_DrawSample_Element) & " " & R
            Else
                picSample.CurrentX = 1080 + 1200 * (c - 1)
                picSample.Print GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_DrawSample_Data) & " " & c - 1
            End If
        Next R
    Next c
    
    Set iPic = picSample.Image
    picSample.Cls
    Set picSample.Picture = iPic
  '  SetWindowRedraw picSample.hWnd, True
    
End Sub

Private Function ToGrey(ByVal Color As Long, Optional ByVal Method As eGreyMethods = gITU) As Long
    Dim iR As Integer
    Dim iG As Integer
    Dim iB As Integer
    Dim iC As Integer
    
    iR = Color And 255
    iG = (Color \ 256) And 255
    iB = (Color \ 65536) And 255
    Select Case Method
        Case gITU ' International Telecommunications Union standard - recommended
            iC = (0.2125 * iR + 0.7154 * iG + 0.0721 * iB)
        Case gNTSCPAL ' NTSC and PAL
            iC = (0.299 * iR + 0.587 * iG + 0.114 * iB)
        Case gAverage ' Simple average
            iC = (iR + iG + iB) / 3
        Case gWeighted ' Weighted average - common
            iC = (3 * iR + 4 * iG + 2 * iB) / 9
        Case gVector ' Distance of color vector in color cube - not recommended
            iC = Sqr(iR ^ 2 + iG ^ 2 + iB ^ 2)
        Case gEye ' Human eye responsive - not recommended (ignores red & blue)
            iC = iG
        Case gAclarado
            iC = (0.2125 * iR + 0.7154 * iG + 0.0721 * iB) + 50
    End Select
    ToGrey = RGB(iC, iC, iC)
End Function
    
    
Public Property Get OKPressed() As Boolean
    OKPressed = mOKPressed
End Property

Public Property Get Changed() As Boolean
    Changed = mChanged
End Property

Public Property Set Grid(nGrid As Control)
    Set mGrid = nGrid
End Property


Public Property Set PrintFnObject(nPrintFnObject As Object)
    Set mPrintFnObject = nPrintFnObject
End Property


Public Property Set FlexFnObject(nFlexFnObject As FlexFnObject)
    Set mFlexFnObject = nFlexFnObject
End Property


Public Property Let ScalePercent(nValue As Long)
    mScalePercent = nValue
    cboScalePercent.Text = mScalePercent & "%"
End Property

Public Property Get ScalePercent() As Long
    ScalePercent = mScalePercent
End Property
    
    
Public Property Let MinScalePercent(nValue As Long)
    mMinScalePercent = nValue
End Property

Public Property Get MinScalePercent() As Long
    MinScalePercent = mMinScalePercent
End Property
    
    
Public Property Let MaxScalePercent(nValue As Long)
    mMaxScalePercent = nValue
End Property

Public Property Get MaxScalePercent() As Long
    MaxScalePercent = mMaxScalePercent
End Property
    
    
Public Property Let ColorMode(nValue As cdeColorModeConstants)
    If (nValue < vbPRCMPrinterDefault) Or (nValue > vbPRCMColor) Then Exit Property
    
    mColorMode = nValue
    cboColor.ListIndex = mColorMode
End Property

Public Property Get ColorMode() As cdeColorModeConstants
    ColorMode = mColorMode
End Property


Public Property Let GridAlign(nValue As Long)
    mGridAlign = nValue
    
    If (mGridAlign >= 0) And (mGridAlign < cboGridAlign.ListCount) Then
        cboGridAlign.ListIndex = mGridAlign
    End If
    
End Property

Public Property Get GridAlign() As Long
    GridAlign = mGridAlign
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


Public Property Set HeadingFont(ByVal nValue As StdFont)
    Set mHeadingFont = nValue
    Set fpcHeading.Font = nValue
End Property

Public Property Get HeadingFont() As StdFont
    Set HeadingFont = mHeadingFont
End Property


Public Property Let HeadingFontColor(nValue As Long)
    mHeadingFontColor = nValue
    fpcHeading.ForeColor = mHeadingFontColor
End Property

Public Property Get HeadingFontColor() As Long
    HeadingFontColor = mHeadingFontColor
End Property


Public Property Let HeadingSampleText(nValue As String)
    Dim iTextLines() As String
    Dim c As Long
    
    iTextLines = Split(nValue, vbCrLf)
    
    For c = 0 To UBound(iTextLines)
        If Trim$(iTextLines(c)) <> "" Then
            mHeadingSampleText = Trim$(iTextLines(c))
            fpcHeading.SampleText = mHeadingSampleText
            Exit For
        End If
    Next c
End Property

Public Property Get HeadingSampleText() As String
    HeadingSampleText = mHeadingSampleText
End Property


Public Property Set SubheadingFont(ByVal nValue As StdFont)
    Set mSubheadingFont = nValue
    Set fpcSubheading.Font = nValue
End Property

Public Property Get SubheadingFont() As StdFont
    Set SubheadingFont = mSubheadingFont
End Property


Public Property Let SubheadingFontColor(nValue As Long)
    mSubheadingFontColor = nValue
    fpcSubheading.ForeColor = mSubheadingFontColor
End Property

Public Property Get SubheadingFontColor() As Long
    SubheadingFontColor = mSubheadingFontColor
End Property


Public Property Let SubheadingSampleText(nValue As String)
    Dim iTextLines() As String
    Dim c As Long
    
    iTextLines = Split(nValue, vbCrLf)
    
    For c = 0 To UBound(iTextLines)
        If Trim$(iTextLines(c)) <> "" Then
            mSubheadingSampleText = Trim$(iTextLines(c))
            fpcSubheading.SampleText = mSubheadingSampleText
            Exit For
        End If
    Next c
End Property

Public Property Get SubheadingSampleText() As String
    SubheadingSampleText = mSubheadingSampleText
End Property


Public Property Set OtherTextsFont(ByVal nValue As StdFont)
    Set mOtherTextsFont = nValue
    Set fpcOtherTexts.Font = nValue
End Property

Public Property Get OtherTextsFont() As StdFont
    Set OtherTextsFont = mOtherTextsFont
End Property


Public Property Let OtherTextsFontColor(nValue As Long)
    mOtherTextsFontColor = nValue
    fpcOtherTexts.ForeColor = mOtherTextsFontColor
End Property

Public Property Get OtherTextsFontColor() As Long
    OtherTextsFontColor = mOtherTextsFontColor
End Property


Public Property Let OtherTextsSampleText(nValue As String)
    Dim iTextLines() As String
    Dim c As Long
    
    iTextLines = Split(nValue, vbCrLf)
    
    For c = 0 To UBound(iTextLines)
        If Trim$(iTextLines(c)) <> "" Then
            mOtherTextsSampleText = Trim$(iTextLines(c))
            fpcOtherTexts.SampleText = mOtherTextsSampleText
            Exit For
        End If
    Next c
End Property

Public Property Get OtherTextsSampleText() As String
    OtherTextsSampleText = mOtherTextsSampleText
End Property


Public Property Let DisplayHeadingFont(nValue As Boolean)
    mDisplayHeadingFont = nValue
    lblHeadingFont.Visible = mDisplayHeadingFont
    fpcHeading.Visible = mDisplayHeadingFont
End Property

Public Property Get DisplayHeadingFont() As Boolean
     DisplayHeadingFont = mDisplayHeadingFont
End Property


Public Property Let DisplaySubheadingFont(nValue As Boolean)
    mDisplaySubheadingFont = nValue
    lblSubheadingFont.Visible = mDisplaySubheadingFont
    fpcSubheading.Visible = mDisplaySubheadingFont
End Property

Public Property Get DisplaySubheadingFont() As Boolean
     DisplaySubheadingFont = mDisplaySubheadingFont
End Property


Public Property Let DisplayOtherTextsFont(nValue As Boolean)
    mDisplayOtherTextsFont = nValue
    lblOtherTextsFont.Visible = mDisplayOtherTextsFont
    fpcOtherTexts.Visible = mDisplayOtherTextsFont
End Property

Public Property Get DisplayOtherTextsFont() As Boolean
     DisplayOtherTextsFont = mDisplayOtherTextsFont
End Property


Public Property Set GridReportStyle(nGridReportStyle As GridReportStyle)
    Set mGridReportStyle = nGridReportStyle.Clone
    SelectProperStyle
    PutControlsToStyle
    DrawSample
End Property

Public Property Get GridReportStyle() As GridReportStyle
    Set GridReportStyle = mGridReportStyle
End Property


Public Property Let EnableAutoOrientation(nValue As Boolean)
    mEnableAutoOrientation = nValue
    chkEnableAutoOrientation.Value = Abs(CLng(mEnableAutoOrientation))
End Property

Public Property Get EnableAutoOrientation() As Boolean
    EnableAutoOrientation = mEnableAutoOrientation
End Property


Private Sub CreateFonts()
    Set mPageNumbersFont = New StdFont
    Set mHeadingFont = New StdFont
    Set mSubheadingFont = New StdFont
    Set mOtherTextsFont = New StdFont
End Sub

Private Sub LoadDefaultSettings()
    cboScalePercent.ListIndex = 1
    cboColor.ListIndex = 0
    cboGridAlign.ListIndex = 0
    cboPageNumbersPosition.ListIndex = 0
    cboPageNumbersFormat.ListIndex = 0
    
    chkEnableAutoOrientation.Value = 1
    
End Sub

Private Sub Form_Resize()
    mLoading = False
    DrawSample

    If Not mDisplayHeadingFont Then
        lblSubheadingFont.Top = lblHeadingFont.Top
        lblSubheadingFont.Caption = lblHeadingFont.Caption
        fpcSubheading.Top = fpcHeading.Top
        If Not mDisplaySubheadingFont Then
            lblOtherTextsFont.Top = lblSubheadingFont.Top
            fpcOtherTexts.Top = fpcSubheading.Top
        Else
            lblOtherTextsFont.Top = lblSubheadingFont.Top + 630
            fpcOtherTexts.Top = fpcSubheading.Top + 630
        End If
    Else
        If Not mDisplaySubheadingFont Then
            lblOtherTextsFont.Top = lblSubheadingFont.Top
            fpcOtherTexts.Top = fpcSubheading.Top
        End If
    End If

End Sub

Private Sub fpcHeading_Change()
    Set mHeadingFont = fpcHeading.Font
    mHeadingFontColor = fpcHeading.ForeColor
    mChanged = True
End Sub

Private Sub fpcSubheading_Change()
    Set mSubheadingFont = fpcSubheading.Font
    mSubheadingFontColor = fpcSubheading.ForeColor
    mChanged = True
End Sub

Private Sub fpcOtherTexts_Change()
    Set mOtherTextsFont = fpcOtherTexts.Font
    mOtherTextsFontColor = fpcOtherTexts.ForeColor
    mChanged = True
End Sub

Private Sub LoadPageNumbersFormatStrings()
    Dim c As Long
    
    cboPageNumbersFormat.Clear
    For c = 0 To mFlexFnObject.GetPredefinedPageNumbersFormatStringsCount - 1
        cboPageNumbersFormat.AddItem PrinterExCurrentDocument.GetFormattedPageNumberString(mFlexFnObject.GetPredefinedPageNumbersFormatString(c), 10, 30)
    Next c
End Sub

Private Sub SelectProperStyle()
    Dim iStyleID As String
    Dim c As Long

    iStyleID = mFlexFnObject.GetStyleID(mGridReportStyle, 0)
    For c = 0 To UBound(mStylesIDs)
        If mStylesIDs(c) = iStyleID Then
            cboStyle.ListIndex = c
            Exit Sub
        End If
    Next c
    cboStyle.ListIndex = cboStyle.ListCount - 1
End Sub

Private Sub SetContainerElementsVisibility(nContainer As Control, nVisible As Boolean)
    Dim iCtl As Control
    
    For Each iCtl In Me.Controls
        If TypeName(iCtl) <> "Timer" Then
            If iCtl.Container Is nContainer Then
                If iCtl.Name <> "cboStyle" Then
                    iCtl.Visible = nVisible
                End If
            End If
        End If
    Next iCtl
End Sub

Private Sub sst1_ChangeControlBackColor(ControlName As String, ControlTypeName As String, Cancel As Boolean)
    If ControlTypeName = "ButtonEx" Then Cancel = True
End Sub

Private Sub sst1_TabSelChange()
    AssignAccelerators Me, True
End Sub

Private Sub tmrInit_Timer()
    tmrInit.Enabled = False
    mChanged = False
End Sub

Private Sub tmrUnload_Timer()
    tmrUnload.Enabled = False
    Unload Me
End Sub

Private Sub txtLineWidth_Change()
    If txtLineWidth.Text <> "" Then
        ValidateLineWidth
        DrawSample
        StyleChanged = True
    End If
    mChanged = True
End Sub

Public Property Get NewCustomStyle() As Boolean
    NewCustomStyle = mNewCustomStyle
End Property

Private Sub LoadStyles()
    Dim c As Long
    Dim iGridReportStyle As GridReportStyle
    Dim iCount As Long
    
    ReDim mStylesIDs(0)
    cboStyle.Clear
    iCount = -1
    c = 1
    Set iGridReportStyle = mFlexFnObject.GetGridReportStyle("Default" & c)
    Do Until iGridReportStyle.Tag = ""
        iCount = iCount + 1
        ReDim Preserve mStylesIDs(iCount)
        mStylesIDs(iCount) = "Default" & c
        cboStyle.AddItem GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_Style) & " " & c
        c = c + 1
        Set iGridReportStyle = mFlexFnObject.GetGridReportStyle("Default" & c)
    Loop
    c = 1
    Set iGridReportStyle = mFlexFnObject.GetGridReportStyle("Custom" & c)
    Do Until iGridReportStyle.Tag = ""
        iCount = iCount + 1
        ReDim Preserve mStylesIDs(iCount)
        mStylesIDs(iCount) = "Custom" & c
        cboStyle.AddItem GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_CustomStyle) & " " & c
        c = c + 1
        Set iGridReportStyle = mFlexFnObject.GetGridReportStyle("Custom" & c)
    Loop
    
    cboStyle.AddItem GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_cboStyle_List_Customize)
End Sub

Private Sub txtLineWidth_LostFocus()
    ValidateLineWidth
End Sub

Private Sub txtLineWidthHeadersSeparatorLine_Change()
    If txtLineWidthHeadersSeparatorLine.Text <> "" Then
        ValidateLineWidthHeadersSeparatorLine
        DrawSample
        StyleChanged = True
        mChanged = True
    End If
End Sub

Private Sub PutControlsToStyle()
    mPuttingControlsToStyle = True
    chkPrintOuterBorder.Value = Abs(CLng(mGridReportStyle.PrintOuterBorder))
    chkPrintHeadersBorder.Value = Abs(CLng(mGridReportStyle.PrintHeadersBorder))
    chkPrintColumnsDataLines.Value = Abs(CLng(mGridReportStyle.PrintColumnsDataLines))
    chkPrintColumnsHeadersLines.Value = Abs(CLng(mGridReportStyle.PrintColumnsHeadersLines))
    chkPrintRowsLines.Value = Abs(CLng(mGridReportStyle.PrintRowsLines))
    chkPrintHeadersSeparatorLine.Value = Abs(CLng(mGridReportStyle.PrintHeadersSeparatorLine))
    If Not picPrintHeadersSeparatorLine.Enabled Then
        chkPrintHeadersSeparatorLine.Tag = chkPrintHeadersSeparatorLine.Value
        chkPrintHeadersSeparatorLine.Value = 2
    End If
    chkPrintHeadersBackground.Value = Abs(CLng(mGridReportStyle.PrintHeadersBackground))
    chkPrintFixedColsBackground.Value = Abs(CLng(mGridReportStyle.PrintFixedColsBackground))
    chkPrintOtherBackgrounds.Value = Abs(CLng(mGridReportStyle.PrintOtherBackgrounds))
    txtLineWidth.Text = mGridReportStyle.LineWidth
    txtLineWidthHeadersSeparatorLine.Text = mGridReportStyle.LineWidthHeadersSeparatorLine
    mPuttingControlsToStyle = False
    DrawSample
End Sub

Private Sub ValidateLineWidth()
    Dim iVal As Long
    Dim iSM As Boolean
    
    iSM = False
    iVal = Val(txtLineWidth.Text)
    If iVal < 1 Then
        iSM = True
        txtLineWidth.Text = "1"
    Else
        If iVal > 10 Then
            iSM = True
            txtLineWidth.Text = "10"
        End If
    End If
    mGridReportStyle.LineWidth = Val(txtLineWidth.Text)
    If iSM Then
        MsgBox GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_ValidateLineWidth_Message), vbExclamation, ClientProductName
        txtLineWidth.SelStart = 0
        txtLineWidth.SelLength = Len(txtLineWidth.Text)
        txtLineWidth.SetFocus
    End If
    
End Sub

Private Sub ValidateLineWidthHeadersSeparatorLine()
    Dim iVal As Long
    Dim iSM As Boolean
    
    iSM = False
    iVal = Val(txtLineWidthHeadersSeparatorLine.Text)
    If iVal < 1 Then
        iSM = True
        txtLineWidthHeadersSeparatorLine.Text = "1"
    Else
        If iVal > 20 Then
            iSM = True
            txtLineWidthHeadersSeparatorLine.Text = "20"
        End If
    End If
    If iSM Then
        MsgBox GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_ValidateLineWidthHeadersSeparatorLine_Message), vbExclamation, ClientProductName
        txtLineWidthHeadersSeparatorLine.SelStart = 0
        txtLineWidthHeadersSeparatorLine.SelLength = Len(txtLineWidthHeadersSeparatorLine.Text)
        txtLineWidthHeadersSeparatorLine.SetFocus
    End If
    mGridReportStyle.LineWidthHeadersSeparatorLine = Val(txtLineWidthHeadersSeparatorLine.Text)

End Sub

Private Sub txtLineWidthHeadersSeparatorLine_LostFocus()
    ValidateLineWidthHeadersSeparatorLine
End Sub

Private Property Get StyleChanged() As Boolean
    StyleChanged = mStyleChanged
End Property

Private Property Let StyleChanged(nValue As Boolean)
    If mLoading Then Exit Property
    mStyleChanged = nValue
End Property

Private Sub LoadGUICaptions()
    Dim c As Long
    Dim iSkip As Boolean
    
    Me.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_Caption)
    cmdOK.Caption = GetLocalizedString(efnGUIStr_General_OKButton_Caption)
    cmdCancel.Caption = GetLocalizedString(efnGUIStr_General_CancelButton_Caption)
    
    sst1.TabCaption(0) = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_0)
    sst1.TabCaption(1) = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_1)
    sst1.TabCaption(2) = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_sst1_TabCaption_2)
    
    chkEnableAutoOrientation.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_chkEnableAutoOrientation_Caption)
    chkPrintHeadersSeparatorLine.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersSeparatorLine_Caption)
    chkPrintColumnsHeadersLines.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_chkPrintColumnsHeadersLines_Caption)
    chkPrintHeadersBorder.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersBorder_Caption)
    chkPrintFixedColsBackground.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_chkPrintFixedColsBackground_Caption)
    chkPrintHeadersBackground.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_chkPrintHeadersBackground_Caption)
    chkPrintOtherBackgrounds.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_chkPrintOtherBackgrounds_Caption)
    chkPrintRowsLines.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_chkPrintRowsLines_Caption)
    chkPrintColumnsDataLines.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_chkPrintColumnsDataLines_Caption)
    chkPrintOuterBorder.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_chkPrintOuterBorder_Caption)
    lblLineWidth.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_lblLineWidth_Caption)
    lblStyle.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_lblStyle_Caption)
    lblOtherTextsFont.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_lblOtherTextsFont_Caption)
    lblSubheadingFont.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_lblSubheadingFont_Caption)
    lblHeadingFont.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_lblHeadingFont_Caption)
    lblPageNumbersFont.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersFont_Caption)
    lblPageNumbersFormat.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersFormat_Caption)
    lblPageNumbersPosition.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_frmPageNumbersOptions_lblPageNumbersPosition_Caption)
    lblGridAlign.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_lblGridAlign_Caption)
    lblColor.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_lblColor_Caption)
    lblScalePercent.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_lblScalePercent_Caption)
    
    lblSample.Caption = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_lblSample_Caption)
    
    cboGridAlign.ToolTipText = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_cboGridAlign_ToolTipText)
    txtLineWidth.ToolTipText = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_txtLineWidth_ToolTipText)
    cmdHeadersBackgroundColor.ToolTipText = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_cmdHeadersBackgroundColor_ToolTipText)
    txtLineWidthHeadersSeparatorLine.ToolTipText = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_txtLineWidthHeadersSeparatorLine_ToolTipText)
    cmdOuterBorderColor.ToolTipText = GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_VariousChangeColorCommandButtons_ToolTipText)
    cmdHeadersBorderColor.ToolTipText = cmdOuterBorderColor.ToolTipText
    cmdColumnsDataLinesColor.ToolTipText = cmdOuterBorderColor.ToolTipText
    cmdColumnsHeadersLinesColor.ToolTipText = cmdOuterBorderColor.ToolTipText
    cmdRowsLinesColor.ToolTipText = cmdOuterBorderColor.ToolTipText
    cmdHeadersBorderColor2.ToolTipText = cmdOuterBorderColor.ToolTipText
    
    cboColor.Clear
    For c = 0 To 2
        cboColor.AddItem GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_cboColor_List, c)
    Next c
    
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
    
    cboGridAlign.Clear
    For c = 0 To 3
        cboGridAlign.AddItem GetLocalizedString(efnGUIStr_frmPrintGridFormatOptions_cboGridAlign_List, c)
    Next c
    
    
End Sub

Private Sub UpdatecboScalePercentList()
    Dim c As Long
    Dim iLng As Long
    
    For c = cboScalePercent.ListCount - 1 To 0 Step -1
        iLng = Val(cboScalePercent.List(c))
        If (iLng < mMinScalePercent) Or (iLng > mMaxScalePercent) Then
            cboScalePercent.RemoveItem (c)
        End If
    Next
    If cboScalePercent.ListCount < 3 Then
        cboScalePercent.Clear
        cboScalePercent.AddItem CStr(mMinScalePercent) & "%"
        If (mMaxScalePercent - mMinScalePercent + 1) > 2 Then
            If (mMinScalePercent < 100) And (mMaxScalePercent > 100) Then
                cboScalePercent.AddItem "100%"
            Else
                cboScalePercent.AddItem CStr(Round((mMinScalePercent + mMaxScalePercent) / 2)) & "%"
            End If
        End If
        cboScalePercent.AddItem CStr(mMaxScalePercent) & "%"
    End If
    
End Sub
