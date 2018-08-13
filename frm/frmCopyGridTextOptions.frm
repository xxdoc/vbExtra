VERSION 5.00
Begin VB.Form frmCopyGridTextOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "# Copy text options"
   ClientHeight    =   3972
   ClientLeft      =   5772
   ClientTop       =   3720
   ClientWidth     =   7236
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCopyGridTextOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3972
   ScaleWidth      =   7236
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstColumns 
      Height          =   1872
      Left            =   132
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   1260
      Width           =   6975
   End
   Begin VB.ComboBox cboMode 
      Height          =   300
      ItemData        =   "frmCopyGridTextOptions.frx":000C
      Left            =   132
      List            =   "frmCopyGridTextOptions.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   6975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "# OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   3864
      TabIndex        =   0
      Top             =   3384
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "# Cancel"
      Height          =   435
      Left            =   5460
      TabIndex        =   5
      Top             =   3384
      Width           =   1515
   End
   Begin vbExtra.ButtonEx cmdOK_2 
      Height          =   375
      Left            =   330
      TabIndex        =   6
      Top             =   3510
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   614
      _ExtentY        =   656
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8
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
      Left            =   720
      TabIndex        =   7
      Top             =   3510
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   614
      _ExtentY        =   656
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Label lblSelectComunsToInclude 
      Caption         =   "# Select columns to include:"
      Height          =   312
      Left            =   132
      TabIndex        =   3
      Top             =   924
      Width           =   5232
   End
   Begin VB.Label lblColumnsSeparationMode 
      Caption         =   "# Separation of the columns:"
      Height          =   348
      Left            =   132
      TabIndex        =   1
      Top             =   204
      Width           =   4668
   End
End
Attribute VB_Name = "frmCopyGridTextOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mOKPressed As Boolean
Private mCopyToClipboardMode As Long
Private mSpecialSeparatorCharacters As String
Private mGrid As Object
Private mKeyname As String
Private mColumnsHeaders() As String
Private mConsiderColumn() As Boolean
Private mCopyColumn() As Boolean
Private mFont As New StdFont

Private mLoading As Boolean

Private mSave_CopyToClipboardMode As Boolean


Public Property Get OKPressed() As Boolean
    OKPressed = mOKPressed
End Property


Public Property Let SpecialSeparatorCharacters(nValue As String)
    mSpecialSeparatorCharacters = nValue
End Property

Public Property Get SpecialSeparatorCharacters() As String
    SpecialSeparatorCharacters = mSpecialSeparatorCharacters
End Property

'Public Property Let CopyToClipboardMode(nValue As Long)
'    If nValue <> 0 Then
'        mCopyToClipboardMode = nValue
'        cboMode.ListIndex = mCopyToClipboardMode
'    End If
'End Property

Public Sub SetGrid(nGrid As Object, nKeyname As String)
    Set mGrid = nGrid
    mKeyname = nKeyname
    ShowColumns
End Sub

Public Property Get CopyToClipboardMode() As Long
    CopyToClipboardMode = mCopyToClipboardMode
End Property

Private Sub cboMode_Click()
    If mLoading Then Exit Sub
    
    mCopyToClipboardMode = cboMode.ListIndex + 1
    mSave_CopyToClipboardMode = True
End Sub

Private Sub cmdCancel_2_Click()
    cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_2_Click()
    cmdOK_Click
End Sub

Private Sub cmdOK_Click()
    Dim c As Long
    Dim x As Long
    Dim iStr As String
    Dim iDlg As New CommonDialogExObject
    
    If mCopyToClipboardMode = 4 Then
        iStr = InputBox(GetLocalizedString(efnGUIStr_EnterColumnSeparatorMessage), GetLocalizedString(efnGUIStr_EnterColumnSeparatorMessageTitle), mSpecialSeparatorCharacters)
        If iStr = "" Then
            Exit Sub
        End If
        mSpecialSeparatorCharacters = iStr
        SaveSetting AppNameForRegistry, "Preferences", "CopyGridText_SpecialSeparatorCharacters", mSpecialSeparatorCharacters
    Else
        If mCopyToClipboardMode = 3 Then
            MsgBox GetLocalizedString(efnGUIStr_SelectFontMessage), , ClientProductName
            mFont.Name = GetSetting(AppNameForRegistry, "Preferences", "CopyGridText_FontName", "Arial")
            mFont.Size = CSng(GetSetting(AppNameForRegistry, "Preferences", "CopyGridText_FontSize", "12"))
            mFont.Bold = CBool(GetSetting(AppNameForRegistry, "Preferences", "CopyGridText_FontBold", "0"))
            mFont.Italic = CBool(GetSetting(AppNameForRegistry, "Preferences", "CopyGridText_FontItalic", "0"))
            mFont.Underline = CBool(GetSetting(AppNameForRegistry, "Preferences", "CopyGridText_FontUnderline", "0"))
            mFont.Strikethrough = CBool(GetSetting(AppNameForRegistry, "Preferences", "CopyGridText_Strikethrough", "0"))
            Set iDlg.Font = mFont
            iDlg.ShowFont
            If iDlg.Canceled Then
                Exit Sub
            End If
            Set mFont = iDlg.Font
            SaveSetting AppNameForRegistry, "Preferences", "CopyGridText_FontName", mFont.Name
            SaveSetting AppNameForRegistry, "Preferences", "CopyGridText_FontSize", mFont.Size
            SaveSetting AppNameForRegistry, "Preferences", "CopyGridText_FontBold", CLng(mFont.Bold)
            SaveSetting AppNameForRegistry, "Preferences", "CopyGridText_FontItalic", CLng(mFont.Italic)
            SaveSetting AppNameForRegistry, "Preferences", "CopyGridText_FontUnderline", CLng(mFont.Underline)
            SaveSetting AppNameForRegistry, "Preferences", "CopyGridText_Strikethrough", CLng(mFont.Strikethrough)
        End If
    End If
    
    If mSave_CopyToClipboardMode Then
        SaveSetting AppNameForRegistry, "Preferences", "CopyGridText_" & mKeyname & "_CopyToClipboardMode", mCopyToClipboardMode
        mSave_CopyToClipboardMode = False
    End If
    mOKPressed = True
    
    ReDim mCopyColumn(mGrid.Cols - 1)
    
    For c = 0 To mGrid.Cols - 1
        If mConsiderColumn(c) Then
            For x = 0 To lstColumns.ListCount - 1
                If lstColumns.ItemData(x) = c Then
                    SaveSetting AppNameForRegistry, "Preferences", "CopyGridText_" & mKeyname & "_Col" & mColumnsHeaders(c), CLng(lstColumns.Selected(x))
                    If lstColumns.Selected(x) Then
                        mCopyColumn(c) = True
                    End If
                    Exit For
                End If
            Next x
        End If
    Next c
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim iLng As Long
    
    PersistForm Me, Forms
    AutoSizeDropDownWidth cboMode
    LoadGUICaptions
    AssignAccelerators Me, True
    
    iLng = CLng(Val(GetSetting(AppNameForRegistry, "Preferences", "CopyGridText_" & mKeyname & "_CopyToClipboardMode", 2)))
    If (iLng > 0) And (iLng < 5) Then
        cboMode.ListIndex = iLng - 1
    End If
    
    mSpecialSeparatorCharacters = GetSetting(AppNameForRegistry, "Preferences", "CopyGridText_SpecialSeparatorCharacters", mSpecialSeparatorCharacters)

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

Private Sub ShowColumns()
    Dim c As Long
    Dim r As Long
    Dim iStr As String
    Dim x As Long
    
    ReDim mColumnsHeaders(mGrid.Cols - 1)
    ReDim mConsiderColumn(mGrid.Cols - 1)
    
    For c = 0 To mGrid.Cols - 1
        iStr = ""
        For r = 0 To mGrid.FixedRows - 1
            iStr = iStr & mGrid.TextMatrix(r, c) & " "
        Next r
        If Len(iStr) > 0 Then
            iStr = Left$(iStr, Len(iStr) - 1)
        End If
        If Trim$(iStr) <> "" Then
            mColumnsHeaders(c) = iStr
        Else
            mColumnsHeaders(c) = "Col. " & c + 1
        End If
        If (mGrid.ColWidth(c) = -1) Or (mGrid.ColWidth(c) > (Screen.TwipsPerPixelX * 2 + 5)) Then
            mConsiderColumn(c) = True
        End If
    Next c
    
    For c = 0 To mGrid.Cols - 1
        If mConsiderColumn(c) Then
            lstColumns.AddItem mColumnsHeaders(c)
            lstColumns.Selected(lstColumns.NewIndex) = True
            lstColumns.ItemData(lstColumns.NewIndex) = c
        End If
    Next c
    
    On Error Resume Next
    For c = 0 To mGrid.Cols - 1
        If mConsiderColumn(c) Then
            For x = 0 To lstColumns.ListCount - 1
                If lstColumns.ItemData(x) = c Then
                    lstColumns.Selected(x) = GetSetting(AppNameForRegistry, "Preferences", "CopyGridText_" & mKeyname & "_Col" & mColumnsHeaders(c), -1)
                    Exit For
                End If
            Next x
        End If
    Next c
    
    lstColumns.ListIndex = -1
End Sub

Public Property Get CopyColumn(nIndex As Long) As Boolean
    CopyColumn = mCopyColumn(nIndex)
End Property

Public Property Get TargetFont() As StdFont
    Set TargetFont = mFont
End Property

Private Sub LoadGUICaptions()
    Dim c As Long
    
    Me.Caption = GetLocalizedString(efnGUIStr_frmCopyGridTextOptions_Caption)
    cmdOK.Caption = GetLocalizedString(efnGUIStr_General_OKButton_Caption)
    cmdCancel.Caption = GetLocalizedString(efnGUIStr_General_CancelButton_Caption)
    lblColumnsSeparationMode.Caption = GetLocalizedString(efnGUIStr_frmCopyGridTextOptions_lblColumnsSeparationMode_Caption)
    cboMode.Clear
    For c = 0 To 3
        cboMode.AddItem GetLocalizedString(efnGUIStr_frmCopyGridTextOptions_cboMode_List, c)
    Next c
    lblSelectComunsToInclude.Caption = GetLocalizedString(efnGUIStr_lblSelectComunsToInclude_Caption)
End Sub
