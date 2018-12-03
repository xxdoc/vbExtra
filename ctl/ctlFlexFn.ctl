VERSION 5.00
Begin VB.UserControl FlexFn 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LockControls    =   -1  'True
   PropertyPages   =   "ctlFlexFn.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ctlFlexFn.ctx":00C0
   Begin VB.TextBox txtAux 
      BorderStyle     =   0  'None
      Height          =   588
      Left            =   1224
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2808
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Timer tmrFirstResize 
      Interval        =   1
      Left            =   60
      Top             =   2160
   End
   Begin vbExtra.ToolBarDA tbrButtons 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   2595
      _extentx        =   4577
      _extenty        =   714
      buttonscount    =   6
      buttonkey1      =   "Print"
      buttonpic161    =   "ctlFlexFn.ctx":03D2
      buttonpic201    =   "ctlFlexFn.ctx":0726
      buttonpic241    =   "ctlFlexFn.ctx":0C2A
      buttonpic301    =   "ctlFlexFn.ctx":133E
      buttonpic361    =   "ctlFlexFn.ctx":1E5A
      buttonwidth1    =   438
      buttonkey2      =   "Copy"
      buttonpic162    =   "ctlFlexFn.ctx":2DDE
      buttonpic202    =   "ctlFlexFn.ctx":3132
      buttonpic242    =   "ctlFlexFn.ctx":3636
      buttonpic302    =   "ctlFlexFn.ctx":3D4A
      buttonpic362    =   "ctlFlexFn.ctx":4866
      buttonwidth2    =   438
      buttonkey3      =   "Save"
      buttonpic163    =   "ctlFlexFn.ctx":57EA
      buttonpic203    =   "ctlFlexFn.ctx":5B3E
      buttonpic243    =   "ctlFlexFn.ctx":6042
      buttonpic303    =   "ctlFlexFn.ctx":6756
      buttonpic363    =   "ctlFlexFn.ctx":7272
      buttonwidth3    =   438
      buttonkey4      =   "Find"
      buttonpic164    =   "ctlFlexFn.ctx":81F6
      buttonpic204    =   "ctlFlexFn.ctx":854A
      buttonpic244    =   "ctlFlexFn.ctx":8A4E
      buttonpic304    =   "ctlFlexFn.ctx":9162
      buttonpic364    =   "ctlFlexFn.ctx":9C7E
      buttonwidth4    =   438
      buttonkey5      =   "GroupData"
      buttonpic165    =   "ctlFlexFn.ctx":AC02
      buttonpic205    =   "ctlFlexFn.ctx":AF56
      buttonpic245    =   "ctlFlexFn.ctx":B45A
      buttonpic305    =   "ctlFlexFn.ctx":BB6E
      buttonpic365    =   "ctlFlexFn.ctx":C68A
      buttonwidth5    =   438
      buttonkey6      =   "ConfigColumns"
      buttonpic166    =   "ctlFlexFn.ctx":D60E
      buttonpic206    =   "ctlFlexFn.ctx":D962
      buttonpic246    =   "ctlFlexFn.ctx":DE66
      buttonpic306    =   "ctlFlexFn.ctx":E57A
      buttonpic366    =   "ctlFlexFn.ctx":F096
      buttonpic16alt6 =   "ctlFlexFn.ctx":1001A
      buttonpic20alt6 =   "ctlFlexFn.ctx":1036C
      buttonpic24alt6 =   "ctlFlexFn.ctx":1086E
      buttonpic30alt6 =   "ctlFlexFn.ctx":10F80
      buttonpic36alt6 =   "ctlFlexFn.ctx":11A9A
      buttonwidth6    =   438
   End
   Begin VB.Timer tmrShow 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   60
      Top             =   1710
   End
   Begin VB.Image imgIcon 
      Height          =   510
      Left            =   690
      Picture         =   "ctlFlexFn.ctx":12A1C
      Top             =   1680
      Width           =   510
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Begin VB.Menu mnuCustomItemBefore 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSepCustomBefore 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCopyParent 
         Caption         =   "# Copy..."
         Begin VB.Menu mnuCopyCell 
            Caption         =   "# Cell"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCopyRow 
            Caption         =   "# Row"
         End
         Begin VB.Menu mnuCopyColumn 
            Caption         =   "# Column"
         End
         Begin VB.Menu mnuCopyAll 
            Caption         =   "# All"
         End
         Begin VB.Menu mnuCopySelection 
            Caption         =   "# Selection"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "# Print"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "# Save to a file"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "# Find"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGroupData 
         Caption         =   "# Group texts that are the same in columns"
      End
      Begin VB.Menu mnuConfigColumns 
         Caption         =   "# Configure what columns to show in this report"
      End
      Begin VB.Menu mnuSepCustomAfter 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCustomItemAfter 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "FlexFn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Private Const WM_UILANGCHANGED As Long = WM_USER + 12

Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long

' Events
Public Event ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonName As String, ByRef Cancel As Boolean, ByVal ButtonStatePressed As Boolean)
Public Event ButtonClicked(ByVal ButtonIndex As Integer, ByVal ButtonName As String, ByVal ButtonStatePressed As Boolean)
Public Event BeforeShowingPopupMenu(ByVal GridName As String, ByVal GridHasData As Boolean, ByRef Cancel As Boolean)
Public Event TextFound(ByVal Row As Long, ByVal Col As Long, ByVal Text As String)
Public Event BeforeAction(Action As String, ByVal GridName As String, ByVal ExtraParam As Variant, ByRef Cancel As Boolean)
Attribute BeforeAction.VB_MemberFlags = "200"
Public Event AfterAction(Action As String, ByVal GridName As String, ByVal ExtraParam As Variant)
Public Event BeforeOrderingByColumn(ByVal GridName As String, ByVal ColumnClicked As Long, ByRef OrderByThisColumn As Long, Descending As Boolean)
Public Event AfterOrderingByColumn(ByVal GridName As String, ByVal ColumnClicked As Long, ByVal OrderedByCol As Long, ByVal Descending As Boolean)
Public Event BeforePrintGrid(ByVal GridName As String)
Public Event StartPage(ByVal GridName As String)
Public Event EndPage(ByVal GridName As String)
Public Event StartDocument(ByVal GridName As String)
Public Event EndDocument(ByVal GridName As String)
Public Event WorkGridChange(GridName As String)
Public Event PersonalizeDefaultReportStyle(GridName As String, nRS As GridReportStyle)
Public Event PersonalizeDefaultPrintGridFormatSettings(nRF As PrintGridFormatSettings)
Public Event BeforeCopyingToClipboard(ByVal GridName As String, ByRef TextBefore As String, ByRef TextAfter As String)
Public Event AfterSettingColumnsWidths(ByVal GridName As String, WidthsWereAdjusted As Boolean)
Public Event CellTextChange(ByVal GridName As String, Row As Long, Col As Long)
Public Event BeforeTextEdit(ByVal GridName As String, ByRef Cancel As Boolean)
Public Event OrientationChange(ByVal NewOrientation As Long)
Public Event GridHasDataCheck(ByVal GridName As String, ByRef GridHasData As Boolean)
Public Event EnabledFunctionsUpdated()
Public Event CustomPopupMenuItemClick(ItemName As String)
Public Event DocPrinted()

Private WithEvents mFlexFnObject As FlexFnObject
Attribute mFlexFnObject.VB_VarHelpID = -1
'Private WithEvents mTimer As cTimer

' Buttons
Private mPrintButtonVisible As Boolean
Private mFindButtonVisible As Boolean
Private mCopyButtonVisible As Boolean
Private mSaveButtonVisible As Boolean
Private mGroupDataButtonVisible As Boolean
Private mConfigColumnsButtonVisible As Boolean

Private mPrintButtonEnabled As Boolean
Private mFindButtonEnabled As Boolean
Private mCopyButtonEnabled As Boolean
Private mSaveButtonEnabled As Boolean
Private mGroupDataButtonEnabled As Boolean
Private mConfigColumnsButtonEnabled As Boolean

Private mPrintButtonActionPrintPreview As Boolean

' Tooltips
Private mPrintButton_ToolTipText As String
Private mFindButton_ToolTipText As String
Private mCopyButton_ToolTipText As String
Private mSaveButton_ToolTipText As String
Private mGroupDataButton_ToolTipText As String
Private mGroupDataButtonPressed_ToolTipText As String
Private mConfigColumnsButton_ToolTipText As String
Private mConfigColumnsButtonColsHidden_ToolTipText As String
Private mCopyCellMenuCaption As String
Private mCopyRowMenuCaption As String
Private mCopyColumnMenuCaption As String
Private mCopyAllMenuCaption As String
Private mCopySelectionMenuCaption As String

' Default tooltip variables
Private mPrintButton_ToolTipText_Default As String
Private mFindButton_ToolTipText_Default As String
Private mCopyButton_ToolTipText_Default As String
Private mSaveButton_ToolTipText_Default As String
Private mGroupDataButton_ToolTipText_Default As String
Private mGroupDataButtonPressed_ToolTipText_Default As String
Private mConfigColumnsButton_ToolTipText_Default As String
Private mConfigColumnsButtonColsHidden_ToolTipText_Default As String
Private mCopyCellMenuCaption_Default As String
Private mCopyRowMenuCaption_Default As String
Private mCopyColumnMenuCaption_Default As String
Private mCopyAllMenuCaption_Default As String
Private mCopySelectionMenuCaption_Default As String

Public Enum gfnFlexFnStyles
    gfnShowToolbar = 0&
    gfnNoToolbar = 1&
End Enum

Private mStyle As Long

Private mGridExplicitelySet As Boolean

Private mAutoDisplayContextMenu As Boolean
Private mAutoHandleEnabledButtons As Boolean

Private mSubclassedHwnds() As Long
Private mGridPopup As Object
Private mGUIEDisabled As Boolean
Private mShown As Boolean
'Private mGridLast As Object

Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private mParentFormHwnd As Long
'Private mUpdated As Boolean
'Private mLastControlsCount As Long
Private mResizing As Boolean
Private mFormLoaded As Boolean
'Private mGroupDataButtonPressed As Boolean
Private mUserControlHwnd As Long
Private mCellTextToCopy As String
Private mRowTextToCopy As String
Private mColumnTextToCopy As String
Private mIconsSize As vbExToolbarDAIconsSizeConstants

Private mCustomPopupMenuItems_Names() As String
Private mCustomPopupMenuItems_Captions() As String
Private mCustomPopupMenuItems_Enabled() As Boolean
Private mCustomPopupMenuItems_Checked() As Boolean
Private mCustomPopupMenuItems_Before() As Boolean
Private mGridMouseRowAtPopupMenuPoint As Long
Private mGridMouseColAtPopupMenuPoint As Long

Public Property Get Style() As gfnFlexFnStyles
Attribute Style.VB_MemberFlags = "200"
    Style = mStyle
End Property

Public Property Let Style(nValue As gfnFlexFnStyles)
    If mStyle <> nValue Then
        mStyle = nValue
        PropertyChanged "Style"
        SetStyle
    End If
End Property


Public Property Let PrintButtonVisible(nValue As Boolean)
    If nValue <> mPrintButtonVisible Then
        mPrintButtonVisible = nValue
        PropertyChanged "PrintButtonVisible"
        LoadButtons
    End If
End Property

Public Property Get PrintButtonVisible() As Boolean
Attribute PrintButtonVisible.VB_MemberFlags = "400"
    PrintButtonVisible = mPrintButtonVisible
End Property


Public Property Let FindButtonVisible(nValue As Boolean)
    If nValue <> mFindButtonVisible Then
        mFindButtonVisible = nValue
        PropertyChanged "FindButtonVisible"
        LoadButtons
    End If
End Property

Public Property Get FindButtonVisible() As Boolean
Attribute FindButtonVisible.VB_MemberFlags = "400"
    FindButtonVisible = mFindButtonVisible
End Property


Public Property Let CopyButtonVisible(nValue As Boolean)
    If nValue <> mCopyButtonVisible Then
        mCopyButtonVisible = nValue
        PropertyChanged "CopyButtonVisible"
        LoadButtons
    End If
End Property

Public Property Get CopyButtonVisible() As Boolean
Attribute CopyButtonVisible.VB_MemberFlags = "400"
    CopyButtonVisible = mCopyButtonVisible
End Property


Public Property Let SaveButtonVisible(nValue As Boolean)
    If nValue <> mSaveButtonVisible Then
        mSaveButtonVisible = nValue
        PropertyChanged "SaveButtonVisible"
        LoadButtons
    End If
End Property

Public Property Get SaveButtonVisible() As Boolean
Attribute SaveButtonVisible.VB_MemberFlags = "400"
    SaveButtonVisible = mSaveButtonVisible
End Property


Public Property Let GroupDataButtonVisible(nValue As Boolean)
    If nValue <> mGroupDataButtonVisible Then
        mGroupDataButtonVisible = nValue
        PropertyChanged "GroupDataButtonVisible"
        LoadButtons
    End If
End Property

Public Property Get GroupDataButtonVisible() As Boolean
Attribute GroupDataButtonVisible.VB_MemberFlags = "400"
    GroupDataButtonVisible = mGroupDataButtonVisible
End Property


Public Property Let ConfigColumnsButtonVisible(nValue As Boolean)
    If nValue <> mConfigColumnsButtonVisible Then
        mConfigColumnsButtonVisible = nValue
        PropertyChanged "ConfigColumnsButtonVisible"
        LoadButtons
    End If
End Property

Public Property Get ConfigColumnsButtonVisible() As Boolean
Attribute ConfigColumnsButtonVisible.VB_MemberFlags = "400"
    ConfigColumnsButtonVisible = mConfigColumnsButtonVisible
End Property


Public Property Let PrintButtonEnabled(nValue As Boolean)
    If nValue <> mPrintButtonEnabled Then
        mPrintButtonEnabled = nValue
        PropertyChanged "PrintButtonEnabled"
        On Error Resume Next
        tbrButtons.Buttons("Print").Enabled = nValue
    End If
End Property

Public Property Get PrintButtonEnabled() As Boolean
Attribute PrintButtonEnabled.VB_MemberFlags = "400"
    PrintButtonEnabled = mPrintButtonEnabled
End Property


Public Property Let FindButtonEnabled(nValue As Boolean)
    If nValue <> mFindButtonEnabled Then
        mFindButtonEnabled = nValue
        PropertyChanged "FindButtonEnabled"
        On Error Resume Next
        tbrButtons.Buttons("Find").Enabled = nValue
    End If
End Property

Public Property Get FindButtonEnabled() As Boolean
Attribute FindButtonEnabled.VB_MemberFlags = "400"
    FindButtonEnabled = mFindButtonEnabled
End Property


Public Property Let CopyButtonEnabled(nValue As Boolean)
    If nValue <> mCopyButtonEnabled Then
        mCopyButtonEnabled = nValue
        PropertyChanged "CopyButtonEnabled"
        On Error Resume Next
        tbrButtons.Buttons("Copy").Enabled = nValue
    End If
End Property

Public Property Get CopyButtonEnabled() As Boolean
Attribute CopyButtonEnabled.VB_MemberFlags = "400"
    CopyButtonEnabled = mCopyButtonEnabled
End Property


Public Property Let SaveButtonEnabled(nValue As Boolean)
    If nValue <> mSaveButtonEnabled Then
        mSaveButtonEnabled = nValue
        PropertyChanged "SaveButtonEnabled"
        On Error Resume Next
        tbrButtons.Buttons("Save").Enabled = nValue
    End If
End Property

Public Property Get SaveButtonEnabled() As Boolean
Attribute SaveButtonEnabled.VB_MemberFlags = "400"
    SaveButtonEnabled = mSaveButtonEnabled
End Property


Public Property Let GroupDataButtonEnabled(nValue As Boolean)
    If nValue <> mGroupDataButtonEnabled Then
        mGroupDataButtonEnabled = nValue
        PropertyChanged "GroupDataButtonEnabled"
        On Error Resume Next
        tbrButtons.Buttons("GroupData").Enabled = nValue
        If tbrButtons.Buttons("GroupData").Enabled Then
            mFlexFnObject.SameDataGroupedInColumns = (tbrButtons.Buttons("GroupData").Checked = True)
        End If
'        If Not mGroupDataButtonEnabled Then
'            mFlexFnObject.SameDataGroupedInColumns = False
'        Else
'            Action "GroupData", , mGroupDataButtonPressed
'        End If
    End If
End Property

Public Property Get GroupDataButtonEnabled() As Boolean
Attribute GroupDataButtonEnabled.VB_MemberFlags = "400"
    GroupDataButtonEnabled = mGroupDataButtonEnabled
End Property


Public Property Let ConfigColumnsButtonEnabled(nValue As Boolean)
    If nValue <> mConfigColumnsButtonEnabled Then
        mConfigColumnsButtonEnabled = nValue
        PropertyChanged "ConfigColumnsButtonEnabled"
        On Error Resume Next
        tbrButtons.Buttons("ConfigColumns").Enabled = nValue
    End If
End Property

Public Property Get ConfigColumnsButtonEnabled() As Boolean
Attribute ConfigColumnsButtonEnabled.VB_MemberFlags = "400"
    ConfigColumnsButtonEnabled = mConfigColumnsButtonEnabled
End Property


Public Property Let PrintButtonActionPrintPreview(nValue As Boolean)
    If nValue <> mPrintButtonActionPrintPreview Then
        mPrintButtonActionPrintPreview = nValue
        PropertyChanged "PrintButtonActionPrintPreview"
    End If
End Property

Public Property Get PrintButtonActionPrintPreview() As Boolean
    PrintButtonActionPrintPreview = mPrintButtonActionPrintPreview
End Property


Public Property Let PrintButton_ToolTipText(nValue As String)
Attribute PrintButton_ToolTipText.VB_MemberFlags = "400"
    If nValue <> mPrintButton_ToolTipText Then
        mPrintButton_ToolTipText = nValue
        PropertyChanged "PrintButton_ToolTipText"
        On Error Resume Next
        tbrButtons.Buttons("Print").ToolTipText = nValue
    End If
End Property

Public Property Get PrintButton_ToolTipText() As String
Attribute PrintButton_ToolTipText.VB_MemberFlags = "400"
    PrintButton_ToolTipText = mPrintButton_ToolTipText
End Property


Public Property Let FindButton_ToolTipText(nValue As String)
    If nValue <> mFindButton_ToolTipText Then
        mFindButton_ToolTipText = nValue
        PropertyChanged "FindButton_ToolTipText"
        On Error Resume Next
        tbrButtons.Buttons("Find").ToolTipText = nValue
    End If
End Property

Public Property Get FindButton_ToolTipText() As String
Attribute FindButton_ToolTipText.VB_MemberFlags = "400"
    FindButton_ToolTipText = mFindButton_ToolTipText
End Property


Public Property Let CopyButton_ToolTipText(nValue As String)
    If nValue <> mCopyButton_ToolTipText Then
        mCopyButton_ToolTipText = nValue
        PropertyChanged "CopyButton_ToolTipText"
        On Error Resume Next
        tbrButtons.Buttons("Copy").ToolTipText = nValue
    End If
End Property

Public Property Get CopyButton_ToolTipText() As String
Attribute CopyButton_ToolTipText.VB_MemberFlags = "400"
    CopyButton_ToolTipText = mCopyButton_ToolTipText
End Property


Public Property Let SaveButton_ToolTipText(nValue As String)
    If nValue <> mSaveButton_ToolTipText Then
        mSaveButton_ToolTipText = nValue
        PropertyChanged "SaveButton_ToolTipText"
        On Error Resume Next
        tbrButtons.Buttons("Save").ToolTipText = nValue
    End If
End Property

Public Property Get SaveButton_ToolTipText() As String
Attribute SaveButton_ToolTipText.VB_MemberFlags = "400"
    SaveButton_ToolTipText = mSaveButton_ToolTipText
End Property


Public Property Let GroupDataButton_ToolTipText(nValue As String)
    If nValue <> mGroupDataButton_ToolTipText Then
        mGroupDataButton_ToolTipText = nValue
        PropertyChanged "GroupDataButton_ToolTipText"
        On Error Resume Next
        If Not tbrButtons.Buttons("GroupData").Checked Then
            tbrButtons.Buttons("GroupData").ToolTipText = nValue
        End If
    End If
End Property

Public Property Get GroupDataButton_ToolTipText() As String
Attribute GroupDataButton_ToolTipText.VB_MemberFlags = "400"
    GroupDataButton_ToolTipText = mGroupDataButton_ToolTipText
End Property

Public Property Let GroupDataButtonPressed_ToolTipText(nValue As String)
    If nValue <> mGroupDataButtonPressed_ToolTipText Then
        mGroupDataButtonPressed_ToolTipText = nValue
        PropertyChanged "GroupDataButtonPressed_ToolTipText"
        On Error Resume Next
        If tbrButtons.Buttons("GroupData").Checked Then
            tbrButtons.Buttons("GroupData").ToolTipText = nValue
        End If
    End If
End Property

Public Property Get GroupDataButtonPressed_ToolTipText() As String
Attribute GroupDataButtonPressed_ToolTipText.VB_MemberFlags = "400"
    GroupDataButtonPressed_ToolTipText = mGroupDataButtonPressed_ToolTipText
End Property


Public Property Let ConfigColumnsButton_ToolTipText(nValue As String)
    If nValue <> mConfigColumnsButton_ToolTipText Then
        mConfigColumnsButton_ToolTipText = nValue
        PropertyChanged "ConfigColumnsButton_ToolTipText"
        On Error Resume Next
        tbrButtons.Buttons("ConfigColumns").ToolTipText = nValue
    End If
End Property

Public Property Get ConfigColumnsButton_ToolTipText() As String
Attribute ConfigColumnsButton_ToolTipText.VB_MemberFlags = "400"
    ConfigColumnsButton_ToolTipText = mConfigColumnsButton_ToolTipText
End Property


Public Property Let ConfigColumnsButtonColsHidden_ToolTipText(nValue As String)
    If nValue <> mConfigColumnsButtonColsHidden_ToolTipText Then
        mConfigColumnsButtonColsHidden_ToolTipText = nValue
        PropertyChanged "ConfigColumnsButtonColsHidden_ToolTipText"
        On Error Resume Next
        tbrButtons.Buttons("ConfigColumns").ToolTipText = nValue
    End If
End Property

Public Property Get ConfigColumnsButtonColsHidden_ToolTipText() As String
Attribute ConfigColumnsButtonColsHidden_ToolTipText.VB_MemberFlags = "400"
    ConfigColumnsButtonColsHidden_ToolTipText = mConfigColumnsButtonColsHidden_ToolTipText
End Property


Public Property Let CopyCellMenuCaption(nValue As String)
    If nValue <> mCopyCellMenuCaption Then
        mCopyCellMenuCaption = nValue
        PropertyChanged "CopyCellMenuCaption"
    End If
End Property

Public Property Get CopyCellMenuCaption() As String
    CopyCellMenuCaption = mCopyCellMenuCaption
End Property


Public Property Let CopyRowMenuCaption(nValue As String)
    If nValue <> mCopyRowMenuCaption Then
        mCopyRowMenuCaption = nValue
        PropertyChanged "CopyRowMenuCaption"
        mnuCopyRow.Caption = nValue
    End If
End Property

Public Property Get CopyRowMenuCaption() As String
    CopyRowMenuCaption = mCopyRowMenuCaption
End Property


Public Property Let CopyColumnMenuCaption(nValue As String)
    If nValue <> mCopyColumnMenuCaption Then
        mCopyColumnMenuCaption = nValue
        PropertyChanged "CopyColumnMenuCaption"
        mnuCopyColumn.Caption = nValue
    End If
End Property

Public Property Get CopyColumnMenuCaption() As String
    CopyColumnMenuCaption = mCopyColumnMenuCaption
End Property


Public Property Let CopyAllMenuCaption(nValue As String)
    If nValue <> mCopyAllMenuCaption Then
        mCopyAllMenuCaption = nValue
        PropertyChanged "CopyAllMenuCaption"
        mnuCopyAll.Caption = nValue
    End If
End Property

Public Property Get CopyAllMenuCaption() As String
    CopyAllMenuCaption = mCopyAllMenuCaption
End Property


Public Property Let CopySelectionMenuCaption(nValue As String)
    If nValue <> mCopySelectionMenuCaption Then
        mCopySelectionMenuCaption = nValue
        PropertyChanged "CopySelectionMenuCaption"
        mnuCopySelection.Caption = nValue
    End If
End Property

Public Property Get CopySelectionMenuCaption() As String
    CopySelectionMenuCaption = mCopySelectionMenuCaption
End Property


Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrPreprocess
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Dim c As Long
    Dim iGrid As Object
    Dim iCancel As Boolean
    Dim iGridHasData As Boolean
    
    Select Case iMsg
        Case WM_RBUTTONDOWN
            If Not mGUIEDisabled Then
                Set iGrid = GetGridByHwnd(hWnd)
                If Not iGrid Is Nothing Then
                    On Error Resume Next
                    mGridMouseRowAtPopupMenuPoint = iGrid.MouseRow
                    mGridMouseColAtPopupMenuPoint = iGrid.MouseCol
                    On Error GoTo 0
                    ResetCustomPopupMenuItems
                    CheckWhatFunctionsToMakeAvailable iGridHasData
                    RaiseEvent BeforeShowingPopupMenu(iGrid.Name, iGridHasData, iCancel)
                    If BuildPopupMenu(iGrid) Then
                        If Not iCancel Then
                            Set mGridPopup = iGrid
                            PopupMenu mnuPopup
                        End If
                    End If
                End If
            End If
        Case WM_DESTROY
            DetachMessage Me, hWnd, WM_RBUTTONDOWN
            DetachMessage Me, hWnd, WM_DESTROY
            For c = 1 To UBound(mSubclassedHwnds)
                If mSubclassedHwnds(c) = hWnd Then
                    mSubclassedHwnds(c) = 0
                    Exit For
                End If
            Next c
        Case WM_UILANGCHANGED
            UILangChange wParam
        Case Else
    End Select
'    ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, wParam, lParam)
End Function

Private Sub UILangChange(nPrevLang As Long)
    If mPrintButton_ToolTipText = GetLocalizedString(efnGUIStr_FlexFn_PrintButton_ToolTipText_Default, , nPrevLang) Then PrintButton_ToolTipText = GetLocalizedString(efnGUIStr_FlexFn_PrintButton_ToolTipText_Default)
    If mFindButton_ToolTipText = GetLocalizedString(efnGUIStr_FlexFn_FindButton_ToolTipText_Default, , nPrevLang) Then FindButton_ToolTipText = GetLocalizedString(efnGUIStr_FlexFn_FindButton_ToolTipText_Default)
    If mCopyButton_ToolTipText = GetLocalizedString(efnGUIStr_FlexFn_CopyButton_ToolTipText_Default, , nPrevLang) Then CopyButton_ToolTipText = GetLocalizedString(efnGUIStr_FlexFn_CopyButton_ToolTipText_Default)
    If mSaveButton_ToolTipText = GetLocalizedString(efnGUIStr_FlexFn_SaveButton_ToolTipText_Default, , nPrevLang) Then SaveButton_ToolTipText = GetLocalizedString(efnGUIStr_FlexFn_SaveButton_ToolTipText_Default)
    If mGroupDataButton_ToolTipText = GetLocalizedString(efnGUIStr_FlexFn_GroupDataButton_ToolTipText_Default, , nPrevLang) Then GroupDataButton_ToolTipText = GetLocalizedString(efnGUIStr_FlexFn_GroupDataButton_ToolTipText_Default)
    If mGroupDataButtonPressed_ToolTipText = GetLocalizedString(efnGUIStr_FlexFn_GroupDataButtonPressed_ToolTipText_Default, , nPrevLang) Then GroupDataButtonPressed_ToolTipText = GetLocalizedString(efnGUIStr_FlexFn_GroupDataButtonPressed_ToolTipText_Default)
    If mConfigColumnsButton_ToolTipText = GetLocalizedString(efnGUIStr_FlexFn_ConfigColumnsButton_ToolTipText_Default, , nPrevLang) Then ConfigColumnsButton_ToolTipText = GetLocalizedString(efnGUIStr_FlexFn_ConfigColumnsButton_ToolTipText_Default)
    If mConfigColumnsButtonColsHidden_ToolTipText = GetLocalizedString(efnGUIStr_FlexFn_ConfigColumnsButtonColsHidden_ToolTipText_Default, , nPrevLang) Then ConfigColumnsButtonColsHidden_ToolTipText = GetLocalizedString(efnGUIStr_FlexFn_ConfigColumnsButtonColsHidden_ToolTipText_Default)

    If mCopyCellMenuCaption = GetLocalizedString(efnGUIStr_FlexFn_CopyCellMenuCaption_Default, , nPrevLang) Then CopyCellMenuCaption = GetLocalizedString(efnGUIStr_FlexFn_CopyCellMenuCaption_Default)
    If mCopyRowMenuCaption = GetLocalizedString(efnGUIStr_FlexFn_CopyRowMenuCaption_Default, , nPrevLang) Then CopyRowMenuCaption = GetLocalizedString(efnGUIStr_FlexFn_CopyRowMenuCaption_Default)
    If mCopyColumnMenuCaption = GetLocalizedString(efnGUIStr_FlexFn_CopyColumnMenuCaption_Default, , nPrevLang) Then CopyColumnMenuCaption = GetLocalizedString(efnGUIStr_FlexFn_CopyColumnMenuCaption_Default)
    If mCopyAllMenuCaption = GetLocalizedString(efnGUIStr_FlexFn_CopyAllMenuCaption_Default, , nPrevLang) Then CopyAllMenuCaption = GetLocalizedString(efnGUIStr_FlexFn_CopyAllMenuCaption_Default)
    If mCopySelectionMenuCaption = GetLocalizedString(efnGUIStr_FlexFn_CopySelectionMenuCaption_Default, , nPrevLang) Then CopySelectionMenuCaption = GetLocalizedString(efnGUIStr_FlexFn_CopySelectionMenuCaption_Default)
    If mnuCopyParent.Caption = GetLocalizedString(efnGUIStr_FlexFn_mnuCopyParent_Caption, , nPrevLang) Then mnuCopyParent.Caption = GetLocalizedString(efnGUIStr_FlexFn_mnuCopyParent_Caption)
End Sub

Private Sub mFlexFnObject_DocPrinted()
    RaiseEvent DocPrinted
End Sub

Private Sub mFlexFnObject_GetAuxTextBox(nTB As Object)
    Set nTB = txtAux
End Sub

Private Sub mFlexFnObject_OrientationChange(ByVal NewOrientation As Long)
    RaiseEvent OrientationChange(NewOrientation)
End Sub

Private Sub mForm_Load()
    If tmrShow.Enabled Then tmrShow.Enabled = False
    If Not mShown Then
        UserControl_Show
    End If
    
    mFormLoaded = True
End Sub

Private Sub mFlexFnObject_AfterOrderingByColumn(ByVal GridName As String, ByVal ColumnClicked As Long, ByVal OrderedByCol As Long, ByVal Descending As Boolean)
    RaiseEvent AfterOrderingByColumn(GridName, ColumnClicked, OrderedByCol, Descending)
End Sub

Private Sub mFlexFnObject_AfterSettingColumnsWidths(ByVal GridName As String, WidthsWereAdjusted As Boolean)
    RaiseEvent AfterSettingColumnsWidths(GridName, WidthsWereAdjusted)
End Sub

Private Sub mFlexFnObject_BeforeCopyingToClipboard(ByVal GridName As String, TextBefore As String, TextAfter As String)
    RaiseEvent BeforeCopyingToClipboard(GridName, TextBefore, TextAfter)
End Sub

Private Sub mFlexFnObject_BeforeOrderingByColumn(ByVal GridName As String, ByVal ColumnClicked As Long, OrderByThisColumn As Long, Descending As Boolean)
    RaiseEvent BeforeOrderingByColumn(GridName, ColumnClicked, OrderByThisColumn, Descending)
End Sub

Private Sub mFlexFnObject_BeforePrintGrid(ByVal GridName As String)
    RaiseEvent BeforePrintGrid(GridName)
End Sub

Private Sub mFlexFnObject_BeforeTextEdit(ByVal GridName As String, ByRef Cancel As Boolean)
    RaiseEvent BeforeTextEdit(GridName, Cancel)
End Sub

Private Sub mFlexFnObject_CellTextChange(ByVal GridName As String, Row As Long, Col As Long)
    RaiseEvent CellTextChange(GridName, Row, Col)
End Sub

Private Sub mFlexFnObject_PersonalizeDefaultReportStyle(GridName As String, nRS As GridReportStyle)
    RaiseEvent PersonalizeDefaultReportStyle(GridName, nRS)
End Sub

Private Sub mFlexFnObject_StartPage(ByVal GridName As String)
    RaiseEvent StartPage(GridName)
End Sub

Private Sub mFlexFnObject_EndPage(ByVal GridName As String)
    RaiseEvent EndPage(GridName)
End Sub

Private Sub mFlexFnObject_StartDocument(ByVal GridName As String)
    RaiseEvent StartDocument(GridName)
End Sub

Private Sub mFlexFnObject_EndDocument(ByVal GridName As String)
    RaiseEvent EndDocument(GridName)
End Sub

Private Sub mFlexFnObject_TextFound(ByVal Row As Long, ByVal Col As Long, ByVal Text As String)
    RaiseEvent TextFound(Row, Col, Text)
End Sub

Private Sub mFlexFnObject_UpdateUI()
    If mAutoHandleEnabledButtons Then
        If IsWindow(mUserControlHwnd) <> 0 Then
            CheckWhatFunctionsToMakeAvailable
        End If
    End If
End Sub

Private Sub mFlexFnObject_WorkGridChange(GridName As String)
    If IsWindow(mUserControlHwnd) <> 0 Then
        RaiseEvent WorkGridChange(GridName)
    End If
End Sub

Private Sub mnuConfigColumns_Click()
    If Not mGridPopup Is Nothing Then
        Action "ConfigColumns", mGridPopup
    End If
End Sub

Private Sub mnuCopyAll_Click()
    If Not mGridPopup Is Nothing Then
        Action "Copy", mGridPopup
    End If
End Sub

Private Sub mnuCopyCell_Click()
    ClipboardCopyUnicode mCellTextToCopy
End Sub

Private Sub mnuCopyColumn_Click()
    ClipboardCopyUnicode mColumnTextToCopy
End Sub

Private Sub mnuCopyRow_Click()
    ClipboardCopyUnicode mRowTextToCopy
End Sub

Private Sub mnuCopySelection_Click()
    mnuCopyAll_Click
End Sub

Private Sub mnuCustomItemAfter_Click(Index As Integer)
    RaiseEvent CustomPopupMenuItemClick(mnuCustomItemAfter(Index).Tag)
End Sub

Private Sub mnuCustomItemBefore_Click(Index As Integer)
    RaiseEvent CustomPopupMenuItemClick(mnuCustomItemBefore(Index).Tag)
End Sub

Private Sub mnuSave_Click()
    If Not mGridPopup Is Nothing Then
        Action "Save", mGridPopup
    End If
End Sub

Private Sub mnuFind_Click()
    If Not mGridPopup Is Nothing Then
        Action "Find", mGridPopup
    End If
End Sub

Private Sub mnuGroupData_Click()
    If Not mGridPopup Is Nothing Then
        Action "GroupData", mGridPopup, Not (tbrButtons.Buttons.Item("GroupData").Checked = True)
    End If
End Sub

Private Sub mnuPopup_Click()
    mnuPrint.Caption = mPrintButton_ToolTipText
    mnuCopyAll.Caption = mCopyAllMenuCaption
    mnuCopySelection.Caption = mCopySelectionMenuCaption
    mnuSave.Caption = mSaveButton_ToolTipText
    mnuFind.Caption = mFindButton_ToolTipText
    
    If mGroupDataButtonVisible Then
        If mFlexFnObject.ThereAreHiddenCols(Grid) Then
            mnuConfigColumns.Caption = mConfigColumnsButtonColsHidden_ToolTipText
        Else
            mnuConfigColumns.Caption = mConfigColumnsButton_ToolTipText
        End If
    End If
    
    If mConfigColumnsButtonVisible Then
        If SameDataGroupedInColumns(Grid) Then
            mnuGroupData.Caption = mGroupDataButtonPressed_ToolTipText
        Else
            mnuGroupData.Caption = mGroupDataButton_ToolTipText
        End If
    End If
End Sub

Private Sub mnuPrint_Click()
    If Not mGridPopup Is Nothing Then
        If mPrintButtonActionPrintPreview Then
            Action "PrintPreview", mGridPopup
        Else
            Action "Print", mGridPopup
        End If
    End If
End Sub

Private Sub tbrButtons_ButtonClick(Button As ToolBarDAButton)
    Dim iCancel As Boolean
    Dim iGroupData As Boolean
    Dim iButtonIndex As Integer
    Dim iButtonKey As String
    
    iButtonIndex = Button.Index
    iButtonKey = Button.Key
    
    RaiseEvent ButtonClick(iButtonIndex, iButtonKey, iCancel, tbrButtons.Buttons.Item(iButtonIndex).Checked = True)
    If iCancel Then Exit Sub
    Select Case iButtonKey
        Case "Print"
            If mPrintButtonActionPrintPreview Then
                Action "PrintPreview"
            Else
                Action "Print"
            End If
        Case "PrintPreview"
            Action "PrintPreview"
        Case "Find"
            Action "Find"
        Case "Copy"
            Action "Copy"
        Case "Save"
            Action "Save"
        Case "GroupData"
            If tbrButtons.Buttons.Item("GroupData").Checked Then
                tbrButtons.Buttons.Item("GroupData").ToolTipText = mGroupDataButtonPressed_ToolTipText
                iGroupData = True
            Else
                tbrButtons.Buttons.Item("GroupData").ToolTipText = mGroupDataButton_ToolTipText
                iGroupData = False
            End If
            Action "GroupData", , iGroupData
'            mGroupDataButtonPressed = iGroupData
        Case "ConfigColumns"
            Action "ConfigColumns"
    End Select
    RaiseEvent ButtonClicked(iButtonIndex, iButtonKey, tbrButtons.Buttons.Item(iButtonIndex).Checked)
End Sub

Private Sub tmrFirstResize_Timer()
    If Not tbrButtons.Sized Then Exit Sub
    tmrFirstResize.Enabled = False
    UserControl_Resize
End Sub

Private Sub tmrShow_Timer()
    tmrShow.Enabled = False
    If Not mShown Then
        UserControl_Show
    End If
End Sub

Private Sub UserControl_Initialize()
    Set mFlexFnObject = New FlexFnObject
    ReDim mSubclassedHwnds(0)
    
    mPrintButton_ToolTipText_Default = GetLocalizedString(efnGUIStr_FlexFn_PrintButton_ToolTipText_Default)
    mFindButton_ToolTipText_Default = GetLocalizedString(efnGUIStr_FlexFn_FindButton_ToolTipText_Default)
    mCopyButton_ToolTipText_Default = GetLocalizedString(efnGUIStr_FlexFn_CopyButton_ToolTipText_Default)
    mSaveButton_ToolTipText_Default = GetLocalizedString(efnGUIStr_FlexFn_SaveButton_ToolTipText_Default)
    mGroupDataButton_ToolTipText_Default = GetLocalizedString(efnGUIStr_FlexFn_GroupDataButton_ToolTipText_Default)
    mGroupDataButtonPressed_ToolTipText_Default = GetLocalizedString(efnGUIStr_FlexFn_GroupDataButtonPressed_ToolTipText_Default)
    mConfigColumnsButton_ToolTipText_Default = GetLocalizedString(efnGUIStr_FlexFn_ConfigColumnsButton_ToolTipText_Default)
    mConfigColumnsButtonColsHidden_ToolTipText_Default = GetLocalizedString(efnGUIStr_FlexFn_ConfigColumnsButtonColsHidden_ToolTipText_Default)

    mCopyCellMenuCaption_Default = GetLocalizedString(efnGUIStr_FlexFn_CopyCellMenuCaption_Default)
    mCopyRowMenuCaption_Default = GetLocalizedString(efnGUIStr_FlexFn_CopyRowMenuCaption_Default)
    mCopyColumnMenuCaption_Default = GetLocalizedString(efnGUIStr_FlexFn_CopyColumnMenuCaption_Default)
    mCopyAllMenuCaption_Default = GetLocalizedString(efnGUIStr_FlexFn_CopyAllMenuCaption_Default)
    mCopySelectionMenuCaption_Default = GetLocalizedString(efnGUIStr_FlexFn_CopySelectionMenuCaption_Default)
    
    ResetCustomPopupMenuItems
End Sub

Private Sub UserControl_InitProperties()
    mFlexFnObject.PrintFnObject.AmbientUserMode = Ambient.UserMode
    mIconsSize = vxIconsAppDefault
    
    mPrintButtonVisible = True
    mFindButtonVisible = True
    mCopyButtonVisible = True
    mSaveButtonVisible = True
    mGroupDataButtonVisible = False
    mConfigColumnsButtonVisible = False
    
    mPrintButtonActionPrintPreview = True
    
    mPrintButtonEnabled = True
    mFindButtonEnabled = True
    mCopyButtonEnabled = True
    mSaveButtonEnabled = True
    mGroupDataButtonEnabled = True
    mConfigColumnsButtonEnabled = True
    
    mPrintButton_ToolTipText = mPrintButton_ToolTipText_Default
    mFindButton_ToolTipText = mFindButton_ToolTipText_Default
    mCopyButton_ToolTipText = mCopyButton_ToolTipText_Default
    mSaveButton_ToolTipText = mSaveButton_ToolTipText_Default
    mGroupDataButton_ToolTipText = mGroupDataButton_ToolTipText_Default
    mGroupDataButtonPressed_ToolTipText = mGroupDataButtonPressed_ToolTipText_Default
    mConfigColumnsButton_ToolTipText = mConfigColumnsButton_ToolTipText_Default
    mConfigColumnsButtonColsHidden_ToolTipText = mConfigColumnsButtonColsHidden_ToolTipText_Default
    
    mCopyCellMenuCaption = mCopyCellMenuCaption_Default
    mCopyRowMenuCaption = mCopyRowMenuCaption_Default
    mCopyColumnMenuCaption = mCopyColumnMenuCaption_Default
    mCopyAllMenuCaption = mCopyAllMenuCaption_Default
    mCopySelectionMenuCaption = mCopySelectionMenuCaption_Default
    
    mStyle = 0
    tbrButtons.IconsSize = mIconsSize
    
    mFlexFnObject.AmbientUserMode = Ambient.UserMode
    
    If Ambient.UserMode Then
        mUserControlHwnd = UserControl.hWnd
        If TypeOf UserControl.Parent Is Form Then
            Set mForm = UserControl.Parent
        End If
    End If
    
    LoadButtons
    
    mnuCopyRow.Caption = mCopyRowMenuCaption
    mnuCopyColumn.Caption = mCopyColumnMenuCaption
    mnuCopyAll.Caption = mCopyAllMenuCaption
    mnuCopySelection.Caption = mCopySelectionMenuCaption
    mnuCopyParent.Caption = GetLocalizedString(efnGUIStr_FlexFn_mnuCopyParent_Caption)
    
    SetStyle
    
    mAutoDisplayContextMenu = True
    mAutoHandleEnabledButtons = True
    
    If mUserControlHwnd <> 0 Then
        AttachMessage Me, mUserControlHwnd, WM_UILANGCHANGED
        SetProp mUserControlHwnd, "FnExUI", 1
    End If
    
    mFlexFnObject.AutoSelect125PercentScaleOnSmallGrids = True ' on this property, unlike the others, the default setting is neccesary because in the mFlexFnObject its default is False, unlike the control where it is True
    mFlexFnObject.PrintFnObject.FormatButtonVisible = True
'    mFlexFnObject.EnableOrderByColumns = False
'    mFlexFnObject.StretchColumnsWidthsToFill = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim iLng As Long
    Dim iFont As StdFont
    
    mFlexFnObject.AmbientUserMode = Ambient.UserMode
    
    If Ambient.UserMode Then
        mUserControlHwnd = UserControl.hWnd
        If TypeOf UserControl.Parent Is Form Then
            Set mForm = UserControl.Parent
        End If
    End If
    
    mIconsSize = PropBag.ReadProperty("IconsSize", vxIconsAppDefault)
    mPrintButtonVisible = PropBag.ReadProperty("PrintButtonVisible", True)
    mFindButtonVisible = PropBag.ReadProperty("FindButtonVisible", True)
    mCopyButtonVisible = PropBag.ReadProperty("CopyButtonVisible", True)
    mSaveButtonVisible = PropBag.ReadProperty("SaveButtonVisible", True)
    mGroupDataButtonVisible = PropBag.ReadProperty("GroupDataButtonVisible", False)
    mConfigColumnsButtonVisible = PropBag.ReadProperty("ConfigColumnsButtonVisible", False)
    
    mPrintButtonEnabled = PropBag.ReadProperty("PrintButtonEnabled", True)
    mFindButtonEnabled = PropBag.ReadProperty("FindButtonEnabled", True)
    mCopyButtonEnabled = PropBag.ReadProperty("CopyButtonEnabled", True)
    mSaveButtonEnabled = PropBag.ReadProperty("SaveButtonEnabled", True)
    mGroupDataButtonEnabled = PropBag.ReadProperty("GroupDataButtonEnabled", True)
    mConfigColumnsButtonEnabled = PropBag.ReadProperty("ConfigColumnsButtonEnabled", True)
    
    mPrintButtonActionPrintPreview = PropBag.ReadProperty("PrintButtonActionPrintPreview", True)
    
    mPrintButton_ToolTipText = PropBag.ReadProperty("PrintButton_ToolTipText", mPrintButton_ToolTipText_Default)
    mFindButton_ToolTipText = PropBag.ReadProperty("FindButton_ToolTipText", mFindButton_ToolTipText_Default)
    mCopyButton_ToolTipText = PropBag.ReadProperty("CopyButton_ToolTipText", mCopyButton_ToolTipText_Default)
    mSaveButton_ToolTipText = PropBag.ReadProperty("SaveButton_ToolTipText", mSaveButton_ToolTipText_Default)
    mGroupDataButton_ToolTipText = PropBag.ReadProperty("GroupDataButton_ToolTipText", mGroupDataButton_ToolTipText_Default)
    mGroupDataButtonPressed_ToolTipText = PropBag.ReadProperty("GroupDataButtonPressed_ToolTipText", mGroupDataButtonPressed_ToolTipText_Default)
    mConfigColumnsButton_ToolTipText = PropBag.ReadProperty("ConfigColumnsButton_ToolTipText", mConfigColumnsButton_ToolTipText_Default)
    mConfigColumnsButtonColsHidden_ToolTipText = PropBag.ReadProperty("ConfigColumnsButtonColsHidden_ToolTipText", mConfigColumnsButtonColsHidden_ToolTipText_Default)
    
    mCopyCellMenuCaption = PropBag.ReadProperty("CopyCellMenuCaption", mCopyCellMenuCaption_Default)
    mCopyRowMenuCaption = PropBag.ReadProperty("CopyRowMenuCaption", mCopyRowMenuCaption_Default)
    mCopyColumnMenuCaption = PropBag.ReadProperty("CopyColumnMenuCaption", mCopyColumnMenuCaption_Default)
    mCopyAllMenuCaption = PropBag.ReadProperty("CopyAllMenuCaption", mCopyAllMenuCaption_Default)
    mCopySelectionMenuCaption = PropBag.ReadProperty("CopySelectionMenuCaption", mCopySelectionMenuCaption_Default)
    
    mStyle = PropBag.ReadProperty("Style", gfnShowToolbar)
    If Ambient.UserMode Then
        Set mFlexFnObject.GridParent = UserControl.Parent
    End If
    mFlexFnObject.GridName = PropBag.ReadProperty("GridName", "")
    mFlexFnObject.ReportID = PropBag.ReadProperty("ReportID", "")
    mFlexFnObject.FileName = PropBag.ReadProperty("FileName", "")
    mFlexFnObject.Heading = PropBag.ReadProperty("Heading", "")
    mFlexFnObject.Subheading = PropBag.ReadProperty("Subheading", "")
    mFlexFnObject.MiddleText = PropBag.ReadProperty("FinalText", "")
    mFlexFnObject.FinalText = PropBag.ReadProperty("FinalText", "")
    mFlexFnObject.DefaultFolderPath = PropBag.ReadProperty("DefaultFolderPath", "")
    
    mAutoHandleEnabledButtons = PropBag.ReadProperty("AutoHandleEnabledButtons", True)
    mFlexFnObject.GridsFlatAppearance = PropBag.ReadProperty("GridsFlatAppearance", False)
    mFlexFnObject.EnableOrderByColumns = PropBag.ReadProperty("EnableOrderByColumns", False)
    mFlexFnObject.StretchColumnsWidthsToFill = PropBag.ReadProperty("StretchColumnsWidthsToFill", True)
    mFlexFnObject.InitialOrderColumn = PropBag.ReadProperty("InitialOrderColumn", -1)
    mFlexFnObject.InitialOrderDescending = PropBag.ReadProperty("InitialOrderDescending", False)
    mFlexFnObject.SameDataGroupedInColumns = PropBag.ReadProperty("SameDataGroupedInColumns", False)
    mFlexFnObject.BorderColor = PropBag.ReadProperty("BorderColor", &HC0C0C0)
    mFlexFnObject.BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    mFlexFnObject.ShowToolTipsOnLongerCellTexts = PropBag.ReadProperty("ShowToolTipsOnLongerCellTexts", False)
    mFlexFnObject.DoNotRememberOrder = PropBag.ReadProperty("DoNotRememberOrder", False)
    mFlexFnObject.ShowToolTipsForOrderColumns = PropBag.ReadProperty("ShowToolTipsForOrderColumns", True)
    mFlexFnObject.AllowTextEdition = PropBag.ReadProperty("AllowTextEdition", False)
    mFlexFnObject.TextEditionLocked = PropBag.ReadProperty("TextEditionLocked", False)
    mFlexFnObject.PrintPrevUseAltScaleIcons = PropBag.ReadProperty("PrintPrevUseAltScaleIcons", True)
    mFlexFnObject.PrintCellsFormatting = PropBag.ReadProperty("PrintCellsFormatting", vxPCFPrintAllFormatting)
    
    tbrButtons.IconsSize = mIconsSize
    LoadButtons
    
    mnuCopyRow.Caption = mCopyRowMenuCaption
    mnuCopyColumn.Caption = mCopyColumnMenuCaption
    mnuCopyAll.Caption = mCopyAllMenuCaption
    mnuCopySelection.Caption = mCopySelectionMenuCaption
    mnuCopyParent.Caption = GetLocalizedString(efnGUIStr_FlexFn_mnuCopyParent_Caption)
    
    SetStyle
    
    mFlexFnObject.RememberUserPrintingPreferences = PropBag.ReadProperty("RememberUserPrintingPreferences", gfnRememberByReportIDIfAvailable)
    mFlexFnObject.Orientation = PropBag.ReadProperty("Orientation", gfnDecideOnReportWidth)
    mFlexFnObject.MinScalePercent = PropBag.ReadProperty("MinScalePercent", cPrintPreviewDefaultMinScale)
    mFlexFnObject.MaxScalePercent = PropBag.ReadProperty("MaxScalePercent", cPrintPreviewDefaultMaxScale)
    mFlexFnObject.DefaultFormatSettings.ScalePercent = PropBag.ReadProperty("ScalePercent", 100)
    mFlexFnObject.DefaultReportStyle = PropBag.ReadProperty("DefaultReportStyle", 2)
    
    mFlexFnObject.CopyToClipboardMode = PropBag.ReadProperty("CopyToClipboardMode", gfnSeparateWithTabs)
    mFlexFnObject.SpecialSeparatorCharacters = PropBag.ReadProperty("SpecialSeparatorCharacters", "|")
    
    mFlexFnObject.ScrollWithMouseWheel = PropBag.ReadProperty("ScrollWithMouseWheel", gfnScrollEnabled)
    mFlexFnObject.IgnoreEmptyRowsAtTheEnd = PropBag.ReadProperty("IgnoreEmptyRowsAtTheEnd", True)
    mFlexFnObject.AfterSaveAction = PropBag.ReadProperty("AfterSaveAction", 2)
    mFlexFnObject.MergeCellsExcel = PropBag.ReadProperty("MergeCellsExcel", True)
    mFlexFnObject.ShowCopyConfirmationMessage = PropBag.ReadProperty("ShowCopyConfirmationMessage", True)
    mFlexFnObject.AutoSelect125PercentScaleOnSmallGrids = PropBag.ReadProperty("AutoSelect125PercentScaleOnSmallGrids", True)
    mFlexFnObject.CopyConfirmationMessage = PropBag.ReadProperty("CopyConfirmationMessage", "Se copió el texto")
    
    mAutoDisplayContextMenu = PropBag.ReadProperty("AutoDisplayContextMenu", True)
    
    ' PrintFnObject properties
    mFlexFnObject.PrintFnObject.PaperSize = PropBag.ReadProperty("PaperSize", vbPRPSPrinterDefault)
    mFlexFnObject.PrintFnObject.PaperBin = PropBag.ReadProperty("PaperBin", vbPRBNPrinterDefault)
    mFlexFnObject.PrintFnObject.PrintQuality = PropBag.ReadProperty("PrintQuality", vbPRPQPrinterDefault)
    mFlexFnObject.PrintFnObject.ColorMode = PropBag.ReadProperty("ColorMode", vbPRCMPrinterDefault)
    mFlexFnObject.PrintFnObject.Duplex = PropBag.ReadProperty("Duplex", vbPRDPPrinterDefault)
    mFlexFnObject.PrintFnObject.Units = vbMillimeters
    mFlexFnObject.PrintFnObject.MinLeftMargin = PropBag.ReadProperty("MinLeftMargin", 0)
    mFlexFnObject.PrintFnObject.MinRightMargin = PropBag.ReadProperty("MinRightMargin", 0)
    mFlexFnObject.PrintFnObject.MinTopMargin = PropBag.ReadProperty("MinTopMargin", 0)
    mFlexFnObject.PrintFnObject.MinBottomMargin = PropBag.ReadProperty("MinBottomMargin", 0)
    mFlexFnObject.PrintFnObject.LeftMargin = PropBag.ReadProperty("LeftMargin", cLeftMarginDefault)
    mFlexFnObject.PrintFnObject.RightMargin = PropBag.ReadProperty("RightMargin", cRightMarginDefault)
    mFlexFnObject.PrintFnObject.TopMargin = PropBag.ReadProperty("TopMargin", cTopMarginDefault)
    mFlexFnObject.PrintFnObject.BottomMargin = PropBag.ReadProperty("BottomMargin", cBottomMarginDefault)
    mFlexFnObject.PrintFnObject.Units = PropBag.ReadProperty("Units", vbMillimeters)
    mFlexFnObject.PrintFnObject.UnitsForUser = PropBag.ReadProperty("UnitsForUser", cdeMUUserLocale)
    mFlexFnObject.PrintFnObject.PrintPageNumbers = PropBag.ReadProperty("PrintPageNumbers", True)
    mFlexFnObject.PrintFnObject.PageNumbersPosition = PropBag.ReadProperty("PageNumbersPosition", vxPositionBottomRight)
    iLng = PropBag.ReadProperty("PageNumbersFormatIndex", -1)
    If iLng > -1 Then
        mFlexFnObject.PrintFnObject.PageNumbersFormat = mFlexFnObject.PrintFnObject.GetPredefinedPageNumbersFormatString(iLng)
    Else
        mFlexFnObject.PrintFnObject.PageNumbersFormat = PropBag.ReadProperty("PageNumbersFormat", "Default")
    End If
    Set iFont = PropBag.ReadProperty("PageNumbersFont", Nothing)
    If Not iFont Is Nothing Then
        Set mFlexFnObject.PrintFnObject.PageNumbersFont = iFont
    End If
    mFlexFnObject.PrintFnObject.PageNumbersForeColor = PropBag.ReadProperty("PageNumbersForeColor", vbWindowText)
    mFlexFnObject.PrintFnObject.AllowUserChangeScale = PropBag.ReadProperty("AllowUserChangeScale", True)
    mFlexFnObject.PrintFnObject.AllowUserChangeOrientation = PropBag.ReadProperty("AllowUserChangeOrientation", True)
    mFlexFnObject.PrintFnObject.AllowUserChangePaper = PropBag.ReadProperty("AllowUserChangePaper", True)
    mFlexFnObject.PrintFnObject.PrintPrevUseOneToolBar = PropBag.ReadProperty("PrintPrevUseOneToolBar", False)
    mFlexFnObject.PrintFnObject.PrintPrevToolBarIconsSize = PropBag.ReadProperty("PrintPrevToolBarIconsSize", vxPPTIconsAuto)
    mFlexFnObject.PrintFnObject.PageSetupButtonVisible = PropBag.ReadProperty("PageSetupButtonVisible", True)
    mFlexFnObject.PrintFnObject.FormatButtonVisible = PropBag.ReadProperty("FormatButtonVisible", True)
    mFlexFnObject.PrintFnObject.DocumentName = PropBag.ReadProperty("DocumentName", "")
    ' End PrintFnObject properties
    
    If Ambient.UserMode Then
        tmrShow.Enabled = True
        mParentFormHwnd = GetParentFormHwnd(UserControl.Parent.hWnd)
        If TypeOf UserControl.Parent Is Form Then
            Set mFlexFnObject.Form = UserControl.Parent
        End If
        If mUserControlHwnd <> 0 Then
            AttachMessage Me, mUserControlHwnd, WM_UILANGCHANGED
            SetProp mUserControlHwnd, "FnExUI", 1
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    Dim iH As Long
    Dim iW As Long
    
    If mResizing Then Exit Sub
    mResizing = True
    If mStyle = 0 Then
        UserControl.Width = tbrButtons.Width
        UserControl.Height = tbrButtons.Height
    Else
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
    mResizing = False
End Sub

Private Sub UserControl_Show()
    If Ambient.UserMode Then
        mShown = True

        If Not TypeOf Parent Is Form Then
            mFlexFnObject.StoreGridControls
        End If
        If mStyle <> 0 Then
            ShowWindow UserControl.hWnd, SW_HIDE
        End If
        
        If mAutoDisplayContextMenu Then
            SetContextMenu
        End If
        
        If mGroupDataButtonVisible Then
            mFlexFnObject.SameDataGroupedInColumns = CBool(GetSetting(AppNameForRegistry, "Preferences", mFlexFnObject.Context & "_DataGrouped", CLng(mFlexFnObject.SameDataGroupedInColumns)))
        End If

        RaiseEvent PersonalizeDefaultPrintGridFormatSettings(mFlexFnObject.PrintGridFormatSettings)
    End If
End Sub

Private Sub UserControl_Terminate()
    Dim c As Long
    
    For c = 1 To UBound(mSubclassedHwnds)
        If mSubclassedHwnds(c) <> 0 Then
            DetachMessage Me, mSubclassedHwnds(c), WM_RBUTTONDOWN
            DetachMessage Me, mSubclassedHwnds(c), WM_DESTROY
        End If
    Next c
    
    If mUserControlHwnd <> 0 Then
        DetachMessage Me, mUserControlHwnd, WM_UILANGCHANGED
        RemoveProp mUserControlHwnd, "FnExUI"
    End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim iLng As Long
    Dim iUnits As Long
    
    Call PropBag.WriteProperty("IconsSize", mIconsSize, vxIconsAppDefault)
    Call PropBag.WriteProperty("PrintButtonVisible", mPrintButtonVisible, True)
    Call PropBag.WriteProperty("FindButtonVisible", mFindButtonVisible, True)
    Call PropBag.WriteProperty("CopyButtonVisible", mCopyButtonVisible, True)
    Call PropBag.WriteProperty("SaveButtonVisible", mSaveButtonVisible, True)
    Call PropBag.WriteProperty("GroupDataButtonVisible", mGroupDataButtonVisible, False)
    Call PropBag.WriteProperty("ConfigColumnsButtonVisible", mConfigColumnsButtonVisible, False)

    Call PropBag.WriteProperty("PrintButtonEnabled", mPrintButtonEnabled, True)
    Call PropBag.WriteProperty("FindButtonEnabled", mFindButtonEnabled, True)
    Call PropBag.WriteProperty("CopyButtonEnabled", mCopyButtonEnabled, True)
    Call PropBag.WriteProperty("SaveButtonEnabled", mSaveButtonEnabled, True)
    Call PropBag.WriteProperty("GroupDataButtonEnabled", mGroupDataButtonEnabled, True)
    Call PropBag.WriteProperty("ConfigColumnsButtonEnabled", mConfigColumnsButtonEnabled, True)
    
    Call PropBag.WriteProperty("PrintButtonActionPrintPreview", mPrintButtonActionPrintPreview, True)
    
    Call PropBag.WriteProperty("PrintButton_ToolTipText", mPrintButton_ToolTipText, mPrintButton_ToolTipText_Default)
    Call PropBag.WriteProperty("FindButton_ToolTipText", mFindButton_ToolTipText, mFindButton_ToolTipText_Default)
    Call PropBag.WriteProperty("CopyButton_ToolTipText", mCopyButton_ToolTipText, mCopyButton_ToolTipText_Default)
    Call PropBag.WriteProperty("SaveButton_ToolTipText", mSaveButton_ToolTipText, mSaveButton_ToolTipText_Default)
    Call PropBag.WriteProperty("GroupDataButton_ToolTipText", mGroupDataButton_ToolTipText, mGroupDataButton_ToolTipText_Default)
    Call PropBag.WriteProperty("GroupDataButtonPressed_ToolTipText", mGroupDataButtonPressed_ToolTipText, mGroupDataButtonPressed_ToolTipText_Default)
    Call PropBag.WriteProperty("ConfigColumnsButton_ToolTipText", mConfigColumnsButton_ToolTipText, mConfigColumnsButton_ToolTipText_Default)
    Call PropBag.WriteProperty("ConfigColumnsButtonColsHidden_ToolTipText", mConfigColumnsButtonColsHidden_ToolTipText, mConfigColumnsButtonColsHidden_ToolTipText_Default)
    
    Call PropBag.WriteProperty("CopyCellMenuCaption", mCopyCellMenuCaption, mCopyCellMenuCaption_Default)
    Call PropBag.WriteProperty("CopyRowMenuCaption", mCopyRowMenuCaption, mCopyRowMenuCaption_Default)
    Call PropBag.WriteProperty("CopyColumnMenuCaption", mCopyColumnMenuCaption, mCopyColumnMenuCaption_Default)
    Call PropBag.WriteProperty("CopyAllMenuCaption", mCopyAllMenuCaption, mCopyAllMenuCaption_Default)
    Call PropBag.WriteProperty("CopySelectionMenuCaption", mCopySelectionMenuCaption, mCopySelectionMenuCaption_Default)
    
    Call PropBag.WriteProperty("Style", mStyle, gfnShowToolbar)
    Call PropBag.WriteProperty("GridName", mFlexFnObject.GridName, "")
    Call PropBag.WriteProperty("ReportID", mFlexFnObject.ReportID, "")
    Call PropBag.WriteProperty("FileName", mFlexFnObject.FileName, "")
    Call PropBag.WriteProperty("Heading", mFlexFnObject.Heading, "")
    Call PropBag.WriteProperty("Subheading", mFlexFnObject.Subheading, "")
    Call PropBag.WriteProperty("MiddleText", mFlexFnObject.FinalText, "")
    Call PropBag.WriteProperty("FinalText", mFlexFnObject.FinalText, "")
    Call PropBag.WriteProperty("DefaultFolderPath", mFlexFnObject.DefaultFolderPath, "")
    
    Call PropBag.WriteProperty("RememberUserPrintingPreferences", mFlexFnObject.RememberUserPrintingPreferences, gfnRememberByReportIDIfAvailable)
    Call PropBag.WriteProperty("Orientation", mFlexFnObject.Orientation, gfnDecideOnReportWidth)
    Call PropBag.WriteProperty("MinScalePercent", mFlexFnObject.MinScalePercent, cPrintPreviewDefaultMinScale)
    Call PropBag.WriteProperty("MaxScalePercent", mFlexFnObject.MaxScalePercent, cPrintPreviewDefaultMaxScale)
    Call PropBag.WriteProperty("ScalePercent", mFlexFnObject.DefaultFormatSettings.ScalePercent, 100)
    Call PropBag.WriteProperty("DefaultReportStyle", mFlexFnObject.DefaultReportStyle, 2)
    
    Call PropBag.WriteProperty("CopyToClipboardMode", mFlexFnObject.CopyToClipboardMode, gfnSeparateWithTabs)
    Call PropBag.WriteProperty("SpecialSeparatorCharacters", mFlexFnObject.SpecialSeparatorCharacters, "|")
    
    Call PropBag.WriteProperty("ScrollWithMouseWheel", mFlexFnObject.ScrollWithMouseWheel, gfnScrollEnabled)
    Call PropBag.WriteProperty("IgnoreEmptyRowsAtTheEnd", mFlexFnObject.IgnoreEmptyRowsAtTheEnd, True)
    Call PropBag.WriteProperty("AfterSaveAction", mFlexFnObject.AfterSaveAction, 2)
    Call PropBag.WriteProperty("MergeCellsExcel", mFlexFnObject.MergeCellsExcel, True)
    Call PropBag.WriteProperty("ShowCopyConfirmationMessage", mFlexFnObject.ShowCopyConfirmationMessage, True)
    Call PropBag.WriteProperty("AutoSelect125PercentScaleOnSmallGrids", mFlexFnObject.AutoSelect125PercentScaleOnSmallGrids, True)
    Call PropBag.WriteProperty("CopyConfirmationMessage", mFlexFnObject.CopyConfirmationMessage, "Se copió el texto")
    
    Call PropBag.WriteProperty("AutoDisplayContextMenu", mAutoDisplayContextMenu, True)
    Call PropBag.WriteProperty("AutoHandleEnabledButtons", mAutoHandleEnabledButtons, True)
    Call PropBag.WriteProperty("GridsFlatAppearance", mFlexFnObject.GridsFlatAppearance, False)
    Call PropBag.WriteProperty("EnableOrderByColumns", mFlexFnObject.EnableOrderByColumns, False)
    Call PropBag.WriteProperty("StretchColumnsWidthsToFill", mFlexFnObject.StretchColumnsWidthsToFill, True)
    Call PropBag.WriteProperty("InitialOrderColumn", mFlexFnObject.InitialOrderColumn, -1)
    Call PropBag.WriteProperty("InitialOrderDescending", mFlexFnObject.InitialOrderDescending, False)
    Call PropBag.WriteProperty("SameDataGroupedInColumns", mFlexFnObject.SameDataGroupedInColumns, False)
    Call PropBag.WriteProperty("BorderColor", mFlexFnObject.BorderColor, &HC0C0C0)
    Call PropBag.WriteProperty("BorderWidth", mFlexFnObject.BorderWidth, 1)
    Call PropBag.WriteProperty("ShowToolTipsOnLongerCellTexts", mFlexFnObject.ShowToolTipsOnLongerCellTexts, False)
    Call PropBag.WriteProperty("DoNotRememberOrder", mFlexFnObject.DoNotRememberOrder, False)
    Call PropBag.WriteProperty("ShowToolTipsForOrderColumns", mFlexFnObject.ShowToolTipsForOrderColumns, True)
    Call PropBag.WriteProperty("AllowTextEdition", mFlexFnObject.AllowTextEdition, False)
    Call PropBag.WriteProperty("TextEditionLocked", mFlexFnObject.TextEditionLocked, False)
    Call PropBag.WriteProperty("PrintPrevUseAltScaleIcons", mFlexFnObject.PrintPrevUseAltScaleIcons, True)
    Call PropBag.WriteProperty("PrintCellsFormatting", mFlexFnObject.PrintCellsFormatting, vxPCFPrintAllFormatting)
    
    ' PrintFnObject properties
    Call PropBag.WriteProperty("PaperSize", mFlexFnObject.PrintFnObject.PaperSize, vbPRPSPrinterDefault)
    Call PropBag.WriteProperty("PaperBin", mFlexFnObject.PrintFnObject.PaperBin, vbPRBNPrinterDefault)
    Call PropBag.WriteProperty("PrintQuality", mFlexFnObject.PrintFnObject.PrintQuality, vbPRPQPrinterDefault)
    Call PropBag.WriteProperty("ColorMode", mFlexFnObject.PrintFnObject.ColorMode, vbPRCMPrinterDefault)
    Call PropBag.WriteProperty("Duplex", mFlexFnObject.PrintFnObject.Duplex, vbPRDPPrinterDefault)
    Call PropBag.WriteProperty("Units", mFlexFnObject.PrintFnObject.Units, vbMillimeters)
    Call PropBag.WriteProperty("UnitsForUser", mFlexFnObject.PrintFnObject.UnitsForUser, cdeMUUserLocale)
    iUnits = mFlexFnObject.PrintFnObject.Units
    mFlexFnObject.PrintFnObject.Units = vbMillimeters
    Call PropBag.WriteProperty("MinLeftMargin", mFlexFnObject.PrintFnObject.MinLeftMargin, 0)
    Call PropBag.WriteProperty("MinRightMargin", mFlexFnObject.PrintFnObject.MinRightMargin, 0)
    Call PropBag.WriteProperty("MinTopMargin", mFlexFnObject.PrintFnObject.MinTopMargin, 0)
    Call PropBag.WriteProperty("MinBottomMargin", mFlexFnObject.PrintFnObject.MinBottomMargin, 0)
    Call PropBag.WriteProperty("LeftMargin", mFlexFnObject.PrintFnObject.LeftMargin, cLeftMarginDefault)
    Call PropBag.WriteProperty("RightMargin", mFlexFnObject.PrintFnObject.RightMargin, cRightMarginDefault)
    Call PropBag.WriteProperty("TopMargin", mFlexFnObject.PrintFnObject.TopMargin, cTopMarginDefault)
    Call PropBag.WriteProperty("BottomMargin", mFlexFnObject.PrintFnObject.BottomMargin, cBottomMarginDefault)
    mFlexFnObject.PrintFnObject.Units = iUnits
    Call PropBag.WriteProperty("PrintPageNumbers", mFlexFnObject.PrintFnObject.PrintPageNumbers, True)
    Call PropBag.WriteProperty("PageNumbersPosition", mFlexFnObject.PrintFnObject.PageNumbersPosition, vxPositionBottomRight)
    iLng = mFlexFnObject.PrintFnObject.GetPageNumbersFormatStringsIndex(mFlexFnObject.PrintFnObject.PageNumbersFormat)
    If iLng > -1 Then
        PropBag.WriteProperty "PageNumbersFormat", "", "Default"
        PropBag.WriteProperty "PageNumbersFormatIndex", iLng, -1
    Else
        PropBag.WriteProperty "PageNumbersFormat", mFlexFnObject.PrintFnObject.PageNumbersFormat, "Default"
        PropBag.WriteProperty "PageNumbersFormatIndex", -1, -1
    End If
    PropBag.WriteProperty "PageNumbersFont", mFlexFnObject.PrintFnObject.PageNumbersFont, Nothing
    Call PropBag.WriteProperty("PageNumbersForeColor", mFlexFnObject.PrintFnObject.PageNumbersForeColor, vbWindowText)
    Call PropBag.WriteProperty("AllowUserChangeScale", mFlexFnObject.PrintFnObject.AllowUserChangeScale, True)
    Call PropBag.WriteProperty("AllowUserChangeOrientation", mFlexFnObject.PrintFnObject.AllowUserChangeOrientation, True)
    Call PropBag.WriteProperty("AllowUserChangePaper", mFlexFnObject.PrintFnObject.AllowUserChangePaper, True)
    Call PropBag.WriteProperty("PrintPrevUseOneToolBar", mFlexFnObject.PrintFnObject.PrintPrevUseOneToolBar, False)
    Call PropBag.WriteProperty("PrintPrevToolBarIconsSize", mFlexFnObject.PrintFnObject.PrintPrevToolBarIconsSize, vxPPTIconsAuto)
    Call PropBag.WriteProperty("PageSetupButtonVisible", mFlexFnObject.PrintFnObject.PageSetupButtonVisible, True)
    Call PropBag.WriteProperty("FormatButtonVisible", mFlexFnObject.PrintFnObject.FormatButtonVisible, True)
    Call PropBag.WriteProperty("DocumentName", mFlexFnObject.PrintFnObject.DocumentName, "")
    ' End PrintFnObject properties
    
End Sub

Private Sub LoadButtons()
    Dim iDataGrouped As Boolean
    Dim iToolbarWidth As Long
    Dim iButton As ToolBarDAButton
    
    tbrButtons.Redraw = False
    On Error Resume Next
    iDataGrouped = tbrButtons.Buttons("GroupData").Checked
    On Error GoTo 0
    
    iToolbarWidth = tbrButtons.Width
'    Do Until tbrButtons.Buttons.Count = 0
'        tbrButtons.Buttons.Remove tbrButtons.Buttons.Count
'    Loop
    
    Set iButton = tbrButtons.GetButtonByKey("Print")
    If Not iButton Is Nothing Then
        If mPrintButtonVisible Then
            iButton.Visible = True
            iButton.Enabled = mPrintButtonEnabled
            iButton.ToolTipText = mPrintButton_ToolTipText
        Else
            iButton.Visible = False
        End If
    End If
    
    Set iButton = tbrButtons.GetButtonByKey("Copy")
    If Not iButton Is Nothing Then
        If mCopyButtonVisible Then
            iButton.Visible = True
            iButton.Enabled = mCopyButtonEnabled
            iButton.ToolTipText = mCopyButton_ToolTipText
        Else
            iButton.Visible = False
        End If
    End If
    
    Set iButton = tbrButtons.GetButtonByKey("Save")
    If Not iButton Is Nothing Then
        If mSaveButtonVisible Then
            iButton.Visible = True
            iButton.Enabled = mSaveButtonEnabled
            iButton.ToolTipText = mSaveButton_ToolTipText
        Else
            iButton.Visible = False
        End If
    End If
    
    Set iButton = tbrButtons.GetButtonByKey("Find")
    If Not iButton Is Nothing Then
        If mFindButtonVisible Then
            iButton.Visible = True
            iButton.Enabled = mFindButtonEnabled
            iButton.ToolTipText = mFindButton_ToolTipText
        Else
            iButton.Visible = False
        End If
    End If
    
    Set iButton = tbrButtons.GetButtonByKey("GroupData")
    If Not iButton Is Nothing Then
        If mGroupDataButtonVisible Then
            iButton.Visible = True
            iButton.Enabled = mGroupDataButtonEnabled
            iButton.ToolTipText = mGroupDataButton_ToolTipText
            iButton.Checked = iDataGrouped
        Else
            iButton.Visible = False
        End If
    End If
    
    Set iButton = tbrButtons.GetButtonByKey("ConfigColumns")
    If Not iButton Is Nothing Then
        If mConfigColumnsButtonVisible Then
            iButton.Visible = True
            iButton.Enabled = mConfigColumnsButtonEnabled
            iButton.ToolTipText = mConfigColumnsButton_ToolTipText
        Else
            iButton.Visible = False
        End If
    End If
    
    tbrButtons.Redraw = True
    If mStyle = 0 Then
'        tbrButtons.Refresh
        If iToolbarWidth <> tbrButtons.Width Then
            UserControl_Resize
        End If
    End If
End Sub

Private Sub SetStyle()
    If mStyle = 0 Then
     '   tbrButtons.Refresh
        tbrButtons.Visible = True
        Set UserControl.Picture = Nothing
        UserControl_Resize
        If Ambient.UserMode Then
            ShowWindow UserControl.hWnd, SW_SHOW
        End If
    Else
        tbrButtons.Visible = False
        Set UserControl.Picture = imgIcon.Picture
        Set UserControl.MaskPicture = imgIcon.Picture
        UserControl.MaskColor = 14215660
        UserControl.BackStyle = 0
        
        Width = ScaleX(34, vbPixels, vbTwips)
        Height = ScaleY(34, vbPixels, vbTwips)
        If Ambient.UserMode Then
            ShowWindow UserControl.hWnd, SW_HIDE
'            If Not mTimer Is Nothing Then
'                mTimer.Interval = 0
'                Set mTimer = Nothing
'            End If
        End If
    End If
End Sub


Public Property Set Grid(nGrid As Object)
    Set mFlexFnObject.Grid = nGrid
    If Not nGrid Is Nothing Then
        mGridExplicitelySet = True
    Else
        mGridExplicitelySet = False
    End If
End Property

Public Property Get Grid() As Object
    Set Grid = mFlexFnObject.Grid
End Property

Public Property Let Grid(nGrid As Object)
    Set Grid = nGrid
End Property


Public Property Get GridName() As String
    GridName = mFlexFnObject.GridName
End Property

Public Property Let GridName(nValue As String)
    If nValue <> mFlexFnObject.GridName Then
        mFlexFnObject.GridName = nValue
        PropertyChanged "GridName"
    End If
End Property


Public Property Get ReportID() As String
    ReportID = mFlexFnObject.ReportID
End Property

Public Property Let ReportID(nValue As String)
    If nValue <> mFlexFnObject.ReportID Then
        mFlexFnObject.ReportID = nValue
        PropertyChanged "ReportID"
    End If
End Property


Public Property Get FileName() As String
    FileName = mFlexFnObject.FileName
End Property

Public Property Let FileName(nValue As String)
    If nValue <> mFlexFnObject.FileName Then
        mFlexFnObject.FileName = nValue
        PropertyChanged "FileName"
    End If
End Property


Public Property Get Heading() As String
    Heading = mFlexFnObject.Heading
End Property

Public Property Let Heading(nValue As String)
    If nValue <> mFlexFnObject.Heading Then
        mFlexFnObject.Heading = nValue
        PropertyChanged "Heading"
    End If
End Property


Public Property Get Subheading() As String
    Subheading = mFlexFnObject.Subheading
End Property

Public Property Let Subheading(nValue As String)
    If nValue <> mFlexFnObject.Subheading Then
        mFlexFnObject.Subheading = nValue
        PropertyChanged "Subheading"
    End If
End Property


Public Property Get MiddleText() As String
    MiddleText = mFlexFnObject.MiddleText
End Property

Public Property Let MiddleText(nValue As String)
    If nValue <> mFlexFnObject.MiddleText Then
        mFlexFnObject.MiddleText = nValue
        PropertyChanged "MiddleText"
    End If
End Property


Public Property Get FinalText() As String
    FinalText = mFlexFnObject.FinalText
End Property

Public Property Let FinalText(nValue As String)
    If nValue <> mFlexFnObject.FinalText Then
        mFlexFnObject.FinalText = nValue
        PropertyChanged "FinalText"
    End If
End Property


Public Property Get DefaultFolderPath() As String
Attribute DefaultFolderPath.VB_MemberFlags = "400"
    DefaultFolderPath = mFlexFnObject.DefaultFolderPath
End Property

Public Property Let DefaultFolderPath(nValue As String)
    If nValue <> mFlexFnObject.DefaultFolderPath Then
        mFlexFnObject.DefaultFolderPath = nValue
        PropertyChanged "DefaultFolderPath"
    End If
End Property

Public Sub PrintNow(Optional nGrid As Object, Optional nGridName, Optional nReportID, Optional nHeading, Optional nSubheading, Optional nMiddleText, Optional nFinalText, Optional nOrientation As gfnOrientation = -1, Optional nScalePercent)
    mFlexFnObject.PrintNow nGrid, nGridName, nReportID, nHeading, nSubheading, nMiddleText, nFinalText, nOrientation, nScalePercent
End Sub

Public Sub ShowPrint(Optional nGrid As Object, Optional nGridName, Optional nReportID, Optional nHeading, Optional nSubheading, Optional nMiddleText, Optional nFinalText, Optional nOrientation As gfnOrientation = -1, Optional nScalePercent, Optional nPrintWithDefaultSettings As Boolean)
    Dim iCancel As Boolean
    Dim iLng As Long
    Dim iGridName As String
    Dim iGrid As Object
    
    If Not nGrid Is Nothing Then
        Set iGrid = nGrid
    Else
        If Not IsMissing(nGridName) Then
            If Not IsEmpty(nGridName) Then
                iGridName = CStr(nGridName)
            End If
        End If
    End If
    If iGridName = "" Then
        If iGrid Is Nothing Then
            Set iGrid = mFlexFnObject.Grid
            If iGrid Is Nothing Then
                Set iGrid = mFlexFnObject.GetWorkGridControl
            End If
        End If
    End If
    If Not iGrid Is Nothing Then
        iGridName = iGrid.Name
    End If
    
    iLng = CLng(nPrintWithDefaultSettings)
    RaiseEvent BeforeAction("Print", iGridName, iLng, iCancel)
    If Not iCancel Then
        mFlexFnObject.ShowPrint iGrid, nGridName, nReportID, nHeading, nSubheading, nMiddleText, nFinalText, nOrientation, nScalePercent, CBool(iLng)
        RaiseEvent AfterAction("Print", iGridName, iLng)
    End If
    
End Sub

Public Sub ShowPrintPreview(Optional nGrid As Object, Optional nGridName, Optional nReportID, Optional nHeading, Optional nSubheading, Optional nMiddleText, Optional nFinalText, Optional nOrientation As gfnOrientation = -1, Optional nScalePercent)
    Dim iCancel As Boolean
    Dim iGridName As String
    Dim iGrid As Object
    
    If Not nGrid Is Nothing Then
        Set iGrid = nGrid
    Else
        If Not IsMissing(nGridName) Then
            If Not IsEmpty(nGridName) Then
                iGridName = CStr(nGridName)
            End If
        End If
    End If
    If iGridName = "" Then
        If iGrid Is Nothing Then
            Set iGrid = mFlexFnObject.Grid
            If iGrid Is Nothing Then
                Set iGrid = mFlexFnObject.GetWorkGridControl
            End If
        End If
    End If
    If Not iGrid Is Nothing Then
        iGridName = iGrid.Name
    End If
    
    RaiseEvent BeforeAction("PrintPreview", iGridName, 0, iCancel)
    If Not iCancel Then
        mFlexFnObject.ShowPrintPreview iGrid, nGridName, nReportID, nHeading, nSubheading, nMiddleText, nFinalText, nOrientation, nScalePercent
        RaiseEvent AfterAction("PrintPreview", iGridName, 0)
    End If
    
End Sub

Public Sub SaveAsExcelFile(Optional nGrid As Object, Optional nGridName, Optional nReportID, Optional nFileName, Optional nHeading, Optional nSubheading, Optional nMiddleText, Optional nFinalText, Optional nDefaultFolderPath)
    Dim iCancel As Boolean
    Dim iGridName As String
    Dim iGrid As Object
    
    If Not nGrid Is Nothing Then
        Set iGrid = nGrid
    Else
        If Not IsMissing(nGridName) Then
            If Not IsEmpty(nGridName) Then
                iGridName = CStr(nGridName)
            End If
        End If
    End If
    If iGridName = "" Then
        If iGrid Is Nothing Then
            Set iGrid = mFlexFnObject.Grid
            If iGrid Is Nothing Then
                Set iGrid = mFlexFnObject.GetWorkGridControl
            End If
        End If
    End If
    If Not iGrid Is Nothing Then
        iGridName = iGrid.Name
    End If
    
    RaiseEvent BeforeAction("Save", iGridName, 0, iCancel)
    If Not iCancel Then
        mFlexFnObject.SaveAsExcelFile iGrid, nGridName, nReportID, nFileName, nHeading, nSubheading, nMiddleText, nFinalText, nDefaultFolderPath
        RaiseEvent AfterAction("Save", iGridName, 0)
    End If
    
End Sub

Public Sub CopyToClipboard(Optional nGrid As Object, Optional nGridName, Optional nCopyToClipboardMode As gfnCopyToClipboardModeOptions = -1, Optional nSpecialSeparatorCharacters)
    Dim iCancel As Boolean
    Dim iGridName As String
    Dim iGrid As Object
    
    If Not nGrid Is Nothing Then
        Set iGrid = nGrid
    Else
        If Not IsMissing(nGridName) Then
            If Not IsEmpty(nGridName) Then
                iGridName = CStr(nGridName)
            End If
        End If
    End If
    If iGridName = "" Then
        If iGrid Is Nothing Then
            Set iGrid = mFlexFnObject.Grid
            If iGrid Is Nothing Then
                Set iGrid = mFlexFnObject.GetWorkGridControl
            End If
        End If
    End If
    If Not iGrid Is Nothing Then
        iGridName = iGrid.Name
    End If
    
    RaiseEvent BeforeAction("Copy", iGridName, 0, iCancel)
    If Not iCancel Then
        mFlexFnObject.CopyToClipboard iGrid, nGridName, nCopyToClipboardMode, nSpecialSeparatorCharacters
        RaiseEvent AfterAction("Copy", iGridName, 0)
    End If
    
End Sub

Public Sub FindText(Optional nGrid As Object, Optional nGridName, Optional TextToFind, Optional nFindNext As Boolean)
    Dim iCancel As Boolean
    Dim iGridName As String
    Dim iGrid As Object
    
    If Not nGrid Is Nothing Then
        Set iGrid = nGrid
    Else
        If Not IsMissing(nGridName) Then
            If Not IsEmpty(nGridName) Then
                iGridName = CStr(nGridName)
            End If
        End If
    End If
    If iGridName = "" Then
        If iGrid Is Nothing Then
            Set iGrid = mFlexFnObject.Grid
            If iGrid Is Nothing Then
                Set iGrid = mFlexFnObject.GetWorkGridControl
            End If
        End If
    End If
    If Not iGrid Is Nothing Then
        iGridName = iGrid.Name
    End If
    
    RaiseEvent BeforeAction("Find", iGridName, 0, iCancel)
    If Not iCancel Then
        mFlexFnObject.FindText iGrid, nGridName, TextToFind, nFindNext
        RaiseEvent AfterAction("Find", iGridName, 0)
    End If
    
End Sub

Public Sub ShowConfigColumns(Optional nGrid As Object, Optional nGridName)
    Dim iCancel As Boolean
    Dim iGridName As String
    Dim iGrid As Object
    
    If Not nGrid Is Nothing Then
        Set iGrid = nGrid
    Else
        If Not IsMissing(nGridName) Then
            If Not IsEmpty(nGridName) Then
                iGridName = CStr(nGridName)
            End If
        End If
    End If
    If iGridName = "" Then
        If iGrid Is Nothing Then
            Set iGrid = mFlexFnObject.Grid
            If iGrid Is Nothing Then
                Set iGrid = mFlexFnObject.GetWorkGridControl
            End If
        End If
    End If
    If Not iGrid Is Nothing Then
        iGridName = iGrid.Name
    End If
    
    RaiseEvent BeforeAction("ConfigColumns", iGridName, 0, iCancel)
    If Not iCancel Then
        mFlexFnObject.ShowConfigColumns iGrid, nGridName
    '    Set mGridLast = Nothing
        CheckWhatFunctionsToMakeAvailable
        RaiseEvent AfterAction("ConfigColumns", iGridName, 0)
    End If
    
End Sub

Private Sub UpdateConfigColsIcon(nGrid As Object)
    If mFlexFnObject.ThereAreHiddenCols(nGrid) Then
        tbrButtons.Buttons("ConfigColumns").UseAltPic = True
        tbrButtons.Buttons("ConfigColumns").ToolTipText = mConfigColumnsButtonColsHidden_ToolTipText
    Else
        tbrButtons.Buttons("ConfigColumns").UseAltPic = False
        tbrButtons.Buttons("ConfigColumns").ToolTipText = mConfigColumnsButton_ToolTipText
    End If
End Sub

Public Property Get DefaultFormatSettings() As PrintGridFormatSettings
    Set DefaultFormatSettings = mFlexFnObject.DefaultFormatSettings
End Property


Public Property Let Orientation(nValue As gfnOrientation)
    If nValue <> mFlexFnObject.Orientation Then
        mFlexFnObject.Orientation = nValue
        PropertyChanged "Orientation"
    End If
End Property

Public Property Get Orientation() As gfnOrientation
    Orientation = mFlexFnObject.Orientation
End Property


Public Property Let ScalePercent(nValue As Long)
    If nValue <> mFlexFnObject.DefaultFormatSettings.ScalePercent Then
        mFlexFnObject.DefaultFormatSettings.ScalePercent = nValue
        PropertyChanged "ScalePercent"
    End If
End Property

Public Property Get ScalePercent() As Long
      ScalePercent = mFlexFnObject.DefaultFormatSettings.ScalePercent
End Property


Public Property Let CopyToClipboardMode(nValue As gfnCopyToClipboardModeOptions)
    If nValue <> mFlexFnObject.CopyToClipboardMode Then
        mFlexFnObject.CopyToClipboardMode = nValue
        PropertyChanged "CopyToClipboardMode"
    End If
End Property

Public Property Get CopyToClipboardMode() As gfnCopyToClipboardModeOptions
    CopyToClipboardMode = mFlexFnObject.CopyToClipboardMode
End Property


Public Property Let SpecialSeparatorCharacters(nValue As String)
    If nValue <> mFlexFnObject.SpecialSeparatorCharacters Then
        mFlexFnObject.SpecialSeparatorCharacters = nValue
        PropertyChanged "SpecialSeparatorCharacters"
    End If
End Property

Public Property Get SpecialSeparatorCharacters() As String
    SpecialSeparatorCharacters = mFlexFnObject.SpecialSeparatorCharacters
End Property


Public Property Let ScrollWithMouseWheel(nValue As gfnScrollWithMouseWheelSettings)
    If nValue <> mFlexFnObject.ScrollWithMouseWheel Then
        mFlexFnObject.ScrollWithMouseWheel = nValue
        PropertyChanged "ScrollWithMouseWheel"
    End If
End Property

Public Property Get ScrollWithMouseWheel() As gfnScrollWithMouseWheelSettings
    ScrollWithMouseWheel = mFlexFnObject.ScrollWithMouseWheel
End Property


Public Property Let IgnoreEmptyRowsAtTheEnd(nValue As Boolean)
    If nValue <> mFlexFnObject.IgnoreEmptyRowsAtTheEnd Then
        mFlexFnObject.IgnoreEmptyRowsAtTheEnd = nValue
        PropertyChanged "IgnoreEmptyRowsAtTheEnd"
    End If
End Property

Public Property Get IgnoreEmptyRowsAtTheEnd() As Boolean
    IgnoreEmptyRowsAtTheEnd = mFlexFnObject.IgnoreEmptyRowsAtTheEnd
End Property


Public Property Let AutoDisplayContextMenu(nValue As Boolean)
    If nValue <> mAutoDisplayContextMenu Then
        mAutoDisplayContextMenu = nValue
        PropertyChanged "AutoDisplayContextMenu"
        SetContextMenu
    End If
End Property

Public Property Get AutoDisplayContextMenu() As Boolean
    AutoDisplayContextMenu = mAutoDisplayContextMenu
End Property


Public Property Let AutoHandleEnabledButtons(nValue As Boolean)
    If nValue <> mAutoHandleEnabledButtons Then
        mAutoHandleEnabledButtons = nValue
        PropertyChanged "AutoHandleEnabledButtons"
    End If
End Property

Public Property Get AutoHandleEnabledButtons() As Boolean
    AutoHandleEnabledButtons = mAutoHandleEnabledButtons
End Property

Private Sub SetContextMenu()
    Dim iAuxLng As Long
    Dim c As Long
    Dim iGH As cGridHandler
    
    If Not Ambient.UserMode Then Exit Sub
    If UBound(mSubclassedHwnds) > 0 Then Exit Sub
    
    If mAutoDisplayContextMenu Then
        ReDim mSubclassedHwnds(0)
        If mFlexFnObject.GridName = "" Then
            If mGridExplicitelySet And Not mFlexFnObject.Grid Is Nothing Then
                On Error Resume Next
                iAuxLng = UserControl.Parent.Controls(mFlexFnObject.Grid.Name).hWnd
                On Error GoTo 0
                If iAuxLng <> 0 Then
                    ReDim Preserve mSubclassedHwnds(UBound(mSubclassedHwnds) + 1)
                    mSubclassedHwnds(UBound(mSubclassedHwnds)) = iAuxLng
                End If
            Else
                On Error Resume Next
                For Each iGH In mFlexFnObject.GridHandlersCollection
                    ReDim Preserve mSubclassedHwnds(UBound(mSubclassedHwnds) + 1)
                    mSubclassedHwnds(UBound(mSubclassedHwnds)) = iGH.Grid.hWnd
                Next iGH
                On Error GoTo 0
            End If
        Else
            On Error Resume Next
            iAuxLng = UserControl.Parent.Controls(mFlexFnObject.GridName).hWnd
            On Error GoTo 0
            If iAuxLng <> 0 Then
                ReDim Preserve mSubclassedHwnds(UBound(mSubclassedHwnds) + 1)
                mSubclassedHwnds(UBound(mSubclassedHwnds)) = iAuxLng
            End If
        End If
        
        For c = 1 To UBound(mSubclassedHwnds)
            AttachMessage Me, mSubclassedHwnds(c), WM_RBUTTONDOWN
            AttachMessage Me, mSubclassedHwnds(c), WM_DESTROY
        Next c
    
    Else
        For c = 1 To UBound(mSubclassedHwnds)
            If mSubclassedHwnds(c) <> 0 Then
                DetachMessage Me, mSubclassedHwnds(c), WM_RBUTTONDOWN
                DetachMessage Me, mSubclassedHwnds(c), WM_DESTROY
            End If
        Next c
        ReDim mSubclassedHwnds(0)
    End If
End Sub

Private Function GetGridByHwnd(nHwnd As Long) As Object
    Dim iGH As cGridHandler
    
    For Each iGH In mFlexFnObject.GridHandlersCollection
        If iGH.Grid.hWnd = nHwnd Then
            Set GetGridByHwnd = iGH.Grid
            Exit Function
        End If
    Next
End Function

'Private Sub UpdateGUIEnabledFunctions()
'    CheckWhatFunctionsToMakeAvailable
'End Sub

Private Sub EnableDisableGUIFunctions(nEnabled As Boolean)
    Dim c As Long
    
    mGUIEDisabled = Not nEnabled
    
    For c = 1 To tbrButtons.Buttons.Count
        If (tbrButtons.Buttons(c).Key <> "ConfigColumns") Then
            tbrButtons.Buttons(c).Enabled = nEnabled
        End If
    Next c
End Sub

Private Sub CheckWhatFunctionsToMakeAvailable(Optional nGridHasData As Boolean)
    Static sGrid As Object
    Dim iGridIsActiveControl As Boolean
    Dim iGridHasData As Boolean
    Dim iColsAreMerged As Boolean
    Dim iGrid_Prev As Object
    Dim iStr As String
    Dim iButtonsEnabled() As Boolean
    Dim iChanged As Boolean
    Dim c As Long
    Static sFirst As Boolean
    
    If Not mAutoHandleEnabledButtons Then Exit Sub
    
    On Error GoTo TheExit:
    
    If Not IsWindowVisibleOnScreen(mParentFormHwnd, True) Then Exit Sub
    If IsWindowEnabled(mParentFormHwnd) = 0 Then Exit Sub
    
    mGUIEDisabled = False
    Set iGrid_Prev = sGrid
    If sGrid Is Nothing Then
        Set sGrid = mFlexFnObject.Grid
    End If
    If sGrid Is Nothing Then
        Set sGrid = mFlexFnObject.GetWorkGridControl(True)
        If sGrid Is Nothing Then
            EnableDisableGUIFunctions False
            Exit Sub
        End If
    End If
    On Error Resume Next
    iGridIsActiveControl = sGrid.Parent.ActiveControl Is sGrid
    On Error GoTo TheExit:
    If Not iGridIsActiveControl Then
        Set sGrid = mFlexFnObject.GetWorkGridControl(True)
        If sGrid Is Nothing Then
            On Error Resume Next
            Set sGrid = UserControl.Parent.Controls(mFlexFnObject.GridName)
            On Error GoTo 0
            If sGrid Is Nothing Then
                EnableDisableGUIFunctions False
                Exit Sub
            End If
        End If
    Else
        If Not IsWindowVisibleOnScreen(sGrid.hWnd, True) Then
            Set sGrid = mFlexFnObject.GetWorkGridControl(True)
            If sGrid Is Nothing Then
                EnableDisableGUIFunctions False
                Exit Sub
            End If
        End If
    End If
    
    ReDim iButtonsEnabled(tbrButtons.Buttons.Count)
    For c = 1 To tbrButtons.Buttons.Count
        iButtonsEnabled(c) = tbrButtons.Buttons(c).Enabled
    Next c
    
    iGridHasData = GridHasData(sGrid)
    If Not sGrid Is Nothing Then
        RaiseEvent GridHasDataCheck(sGrid.Name, iGridHasData)
    End If
    
    If mPrintButtonVisible Then
        tbrButtons.Buttons("Print").Enabled = iGridHasData And mPrintButtonEnabled
    End If
    If mCopyButtonVisible Then
        tbrButtons.Buttons("Copy").Enabled = iGridHasData And mCopyButtonEnabled
    End If
    If mSaveButtonVisible Then
        tbrButtons.Buttons("Save").Enabled = iGridHasData And mSaveButtonEnabled
    End If
    If mFindButtonVisible Then
        tbrButtons.Buttons("Find").Enabled = (sGrid.Rows - sGrid.FixedRows) > 5
    End If
    If mGroupDataButtonVisible Then
        If mGroupDataButtonEnabled Then
            tbrButtons.Buttons("GroupData").Enabled = iGridHasData And mGroupDataButtonEnabled
            iColsAreMerged = False
            If tbrButtons.Buttons("GroupData").Enabled Then
                If ((sGrid.MergeCells = flexMergeFree) Or (sGrid.MergeCells = flexMergeRestrictColumns)) Then
                    If sGrid.Cols > 1 Then
                        If sGrid.MergeCol(1) Then
                            iColsAreMerged = True
                        End If
                    End If
                End If
            End If
            If iColsAreMerged Then
                tbrButtons.Buttons("GroupData").Checked = True
                SameDataGroupedInColumns(sGrid) = True
            Else
                tbrButtons.Buttons("GroupData").Checked = False
                SameDataGroupedInColumns(sGrid) = False
            End If
        End If
    End If
    
    If mConfigColumnsButtonVisible Then
        UpdateConfigColsIcon sGrid
    End If
    
    iChanged = False
    For c = 1 To tbrButtons.Buttons.Count
        If tbrButtons.Buttons(c).Enabled <> iButtonsEnabled(c) Then
            iChanged = True
            Exit For
        End If
    Next c
    If iChanged Or (Not sFirst) Then RaiseEvent EnabledFunctionsUpdated
    sFirst = True
    
TheExit:
    nGridHasData = iGridHasData
    If Not sGrid Is iGrid_Prev Then
        If Not sGrid Is Nothing Then
            iStr = sGrid.Name
        End If
    End If
End Sub

Private Function GridHasData(nGrid As Object) As Boolean
    Dim R As Long
    Dim c As Long
    
    If nGrid.Rows = nGrid.FixedRows Then Exit Function
    
    R = nGrid.FixedRows
    If Trim$(nGrid.TextMatrix(R, 0)) <> "" Then
        GridHasData = True
    Else
        If Trim$(nGrid.TextMatrix(R, nGrid.Cols - 1)) <> "" Then
            GridHasData = True
        End If
    End If
    If Not GridHasData Then
        R = nGrid.Rows - 1
        If Trim$(nGrid.TextMatrix(R, 0)) <> "" Then
            GridHasData = True
        Else
            If Trim$(nGrid.TextMatrix(R, nGrid.Cols - 1)) <> "" Then
                GridHasData = True
            End If
        End If
    End If
    If Not GridHasData Then
        R = (nGrid.Rows - 1) / 2
        If R > nGrid.FixedRows Then
            If Trim$(nGrid.TextMatrix(R, 0)) <> "" Then
                GridHasData = True
            Else
                If Trim$(nGrid.TextMatrix(R, nGrid.Cols - 1)) <> "" Then
                    GridHasData = True
                End If
            End If
        End If
    End If
    If Not GridHasData Then
        R = nGrid.FixedRows
        For c = 1 To nGrid.Cols - 1
            If Trim$(nGrid.TextMatrix(R, c)) <> "" Then
                GridHasData = True
            End If
            If GridHasData Then Exit For
        Next c
        If Not GridHasData Then
            For c = 1 To nGrid.Cols - 1
                R = nGrid.Rows - 1
                If Trim$(nGrid.TextMatrix(R, c)) <> "" Then
                    GridHasData = True
                End If
                If GridHasData Then Exit For
            Next c
        End If
    End If
End Function

Private Function BuildPopupMenu(nGrid As Object) As Boolean
    Dim iCtl As Control
    Dim iLastSep As Long
    Dim iCellText As String
    Dim iPos As Long
    Dim iPos1 As Long
    Dim iPos2 As Long
    Dim iPos3 As Long
    Dim iPos4 As Long
    Dim iStr As String
    Dim c As Long
    Dim iMo As Long
    Dim cb As Long
    Dim ca As Long
    Dim iCopySelection As Boolean
    
    On Error Resume Next
    iCellText = Trim2(nGrid.TextMatrix(nGrid.MouseRow, nGrid.MouseCol))
    On Error GoTo 0
    If iCellText <> "" Then
        mCellTextToCopy = iCellText
        iCellText = Replace(iCellText, vbTab, " ")
        iCellText = Replace(iCellText, vbCrLf, " ")
        iCellText = Replace(iCellText, vbCr, " ")
        iCellText = Replace(iCellText, vbLf, " ")
        If Replace(StrConv(StrConv(iCellText, vbFromUnicode), vbUnicode), "?", "") <> "" Then
            If Len(iCellText) > 30 Then
                
                iPos1 = InStr(30, iCellText, " ")
                iPos2 = InStr(28, iCellText, ",")
                iPos3 = InStr(28, iCellText, ".")
                iPos4 = InStr(28, iCellText, ";")
                
                If iPos1 = 0 Then iPos1 = 100
                If iPos2 = 0 Then iPos2 = 100
                If iPos3 = 0 Then iPos3 = 100
                If iPos4 = 0 Then iPos4 = 100
                
                If (iPos1 < iPos2) And (iPos1 < iPos3) And (iPos1 < iPos4) Then
                    iPos = iPos1
                ElseIf (iPos2 < iPos3) And (iPos2 < iPos4) Then
                    iPos = iPos2
                ElseIf (iPos3 < iPos4) Then
                    iPos = iPos3
                Else
                    iPos = iPos4
                End If
                
                If (iPos > 50) Then
                    iPos = 40
                End If
                iCellText = Left$(iCellText, iPos - 1) & "..."
            End If
            mnuCopyCell.Caption = "'" & iCellText & "'"
        Else
            mnuCopyCell.Caption = mCopyCellMenuCaption
        End If
        mnuCopyCell.Visible = True
    Else
        mnuCopyCell.Visible = False
    End If
    
    
    If (nGrid.HighLight <> flexHighlightNever) Then
        If (nGrid.SelectionMode = flexSelectionByRow) And (nGrid.RowSel > nGrid.Row) Then
            iCopySelection = True
        ElseIf (nGrid.SelectionMode = flexSelectionByColumn) And (nGrid.ColSel > nGrid.Col) Then
            iCopySelection = True
        ElseIf (nGrid.RowSel > nGrid.Row) Or (nGrid.ColSel > nGrid.Col) Then  ' flexSelectionFree
            iCopySelection = True
        End If
    End If
    
    
    On Error Resume Next
    iMo = nGrid.MouseRow
    On Error GoTo 0
    iStr = ""
    For c = 0 To nGrid.Cols - 1
        If nGrid.ColIsVisible(c) Then
            If nGrid.ColWidth(c) <> 0 Then
                If iStr <> "" Then iStr = iStr & vbTab
                iStr = iStr & nGrid.TextMatrix(iMo, c)
            End If
        End If
    Next c
    If Trim2(iStr) <> "" Then
        mRowTextToCopy = iStr
        mnuCopyRow.Visible = True
    Else
        mnuCopyRow.Visible = False
    End If
    
    On Error Resume Next
    iMo = nGrid.MouseCol
    On Error GoTo 0
    iStr = ""
    For c = 0 To nGrid.Rows - 1
        iStr = iStr & nGrid.TextMatrix(c, iMo)
        If c < (nGrid.Rows - 1) Then
            iStr = iStr & vbCrLf
        End If
    Next c
    If Trim2(iStr) <> "" Then
        mColumnTextToCopy = iStr
        mnuCopyColumn.Visible = True
    Else
        mnuCopyColumn.Visible = False
    End If
    
    If mPrintButtonVisible Then
        mnuPrint.Enabled = tbrButtons.Buttons("Print").Enabled
    End If
    If mFindButtonVisible Then
        mnuFind.Enabled = tbrButtons.Buttons("Find").Enabled
    End If
    If mCopyButtonVisible Then
        mnuCopyAll.Enabled = tbrButtons.Buttons("Copy").Enabled
        mnuCopySelection.Enabled = tbrButtons.Buttons("Copy").Enabled
    End If
    If mSaveButtonVisible Then
        mnuSave.Enabled = tbrButtons.Buttons("Save").Enabled
    End If
    If mGroupDataButtonVisible Then
        mnuGroupData.Enabled = tbrButtons.Buttons("GroupData").Enabled
        If tbrButtons.Buttons.Item("GroupData").Checked Then
            mnuGroupData.Caption = mGroupDataButtonPressed_ToolTipText
        Else
            mnuGroupData.Caption = mGroupDataButton_ToolTipText
        End If
    End If
    If mConfigColumnsButtonVisible Then
        mnuConfigColumns.Enabled = tbrButtons.Buttons("ConfigColumns").Enabled
    End If
    
    mnuPrint.Visible = mPrintButtonVisible
    mnuFind.Visible = mFindButtonVisible
    mnuCopyAll.Visible = mCopyButtonVisible And Not iCopySelection
    mnuCopySelection.Visible = mCopyButtonVisible And iCopySelection
    mnuSave.Visible = mSaveButtonVisible
    mnuSep2.Visible = (mFindButtonVisible Or mCopyButtonVisible Or mSaveButtonVisible Or mPrintButtonVisible Or mCopyButtonVisible) And (mGroupDataButtonVisible Or mConfigColumnsButtonVisible)
    If mnuSep2.Visible Then
        iLastSep = 2
    End If
    
    mnuGroupData.Visible = mGroupDataButtonVisible
    mnuConfigColumns.Visible = mConfigColumnsButtonVisible
    
    For Each iCtl In Controls
        If TypeOf iCtl Is Menu Then
            If iCtl.Visible And iCtl.Enabled Then
                If (iCtl.Name <> "mnuPopup") And (Left$(iCtl.Name, 6) <> "mnuSep") Then
                    BuildPopupMenu = True
                    Exit For
                End If
            End If
        End If
    Next iCtl
    
    For c = 0 To mnuCustomItemBefore.UBound
        mnuCustomItemBefore(c).Visible = False
    Next c
    For c = 0 To mnuCustomItemAfter.UBound
        mnuCustomItemAfter(c).Visible = False
    Next c
    mnuSepCustomBefore.Visible = False
    mnuSepCustomAfter.Visible = False
    cb = -1
    ca = -1
    For c = 1 To UBound(mCustomPopupMenuItems_Names)
        If mCustomPopupMenuItems_Before(c) Then
            cb = cb + 1
            If (cb) > mnuCustomItemBefore.UBound Then
                Load mnuCustomItemBefore(cb)
            End If
            mnuCustomItemBefore(cb).Caption = mCustomPopupMenuItems_Captions(c)
            mnuCustomItemBefore(cb).Tag = mCustomPopupMenuItems_Names(c)
            mnuCustomItemBefore(cb).Enabled = mCustomPopupMenuItems_Enabled(c)
            mnuCustomItemBefore(cb).Checked = mCustomPopupMenuItems_Checked(c)
            mnuCustomItemBefore(cb).Visible = True
        Else
            ca = ca + 1
            If (ca) > mnuCustomItemAfter.UBound Then
                Load mnuCustomItemAfter(ca)
            End If
            mnuCustomItemAfter(ca).Caption = mCustomPopupMenuItems_Captions(c)
            mnuCustomItemAfter(ca).Tag = mCustomPopupMenuItems_Names(c)
            mnuCustomItemAfter(ca).Enabled = mCustomPopupMenuItems_Enabled(c)
            mnuCustomItemAfter(ca).Checked = mCustomPopupMenuItems_Checked(c)
            mnuCustomItemAfter(ca).Visible = True
        End If
    Next c
    If mnuCustomItemBefore(0).Visible Then mnuSepCustomBefore.Visible = True
    If mnuCustomItemAfter(0).Visible Then mnuSepCustomAfter.Visible = True
    
End Function


Public Sub Action(ActionName As String, Optional nGrid As Object, Optional nExtraParam As Variant)
    Dim iGrid As Object
    Dim iCancel As Boolean
    Dim iGridName As String
    Dim iActionName As String
    
    If Not nGrid Is Nothing Then
        Set iGrid = nGrid
    Else
        Set iGrid = mFlexFnObject.Grid
        If iGrid Is Nothing Then
            Set iGrid = mFlexFnObject.GetWorkGridControl
        End If
    End If
    If iGrid Is Nothing Then
        Exit Sub
    Else
        iGridName = iGrid.Name
    End If
    
    Select Case ActionName
        Case "Print"
            ShowPrint iGrid, , , , , , , , , CBool(nExtraParam)
        Case "PrintPreview"
            ShowPrintPreview iGrid
        Case "Find"
            FindText iGrid
        Case "FindNext"
            FindText iGrid, , , True
        Case "Copy"
            CopyToClipboard iGrid
        Case "Save"
            SaveAsExcelFile iGrid
        Case "GroupData"
            SameDataGroupedInColumns(iGrid) = CBool(nExtraParam)
        Case "ConfigColumns"
            ShowConfigColumns iGrid
        Case Else
            iActionName = ActionName
            RaiseEvent BeforeAction(iActionName, iGridName, nExtraParam, iCancel)
            If Not iCancel Then
                RaiseEvent AfterAction(iActionName, iGridName, nExtraParam)
            End If
    End Select
    
End Sub

Public Property Get Buttons()
Attribute Buttons.VB_ProcData.VB_Invoke_Property = "ptpGFButtons"
    Set Buttons = tbrButtons.Buttons
End Property

Public Property Get PrintGridFormatSettings() As PrintGridFormatSettings
    Set PrintGridFormatSettings = mFlexFnObject.PrintGridFormatSettings
End Property


Public Property Let AfterSaveAction(nValue As gfnAfterSaveActionSettings)
    mFlexFnObject.AfterSaveAction = nValue
    PropertyChanged "AfterSaveAction"
End Property

Public Property Get AfterSaveAction() As gfnAfterSaveActionSettings
    AfterSaveAction = mFlexFnObject.AfterSaveAction
End Property


Public Property Let MergeCellsExcel(nValue As Boolean)
    mFlexFnObject.MergeCellsExcel = nValue
    PropertyChanged "MergeCellsExcel"
End Property

Public Property Get MergeCellsExcel() As Boolean
    MergeCellsExcel = mFlexFnObject.MergeCellsExcel
End Property


Public Property Let ShowCopyConfirmationMessage(nValue As Boolean)
    mFlexFnObject.ShowCopyConfirmationMessage = nValue
    PropertyChanged "ShowCopyConfirmationMessage"
End Property

Public Property Get ShowCopyConfirmationMessage() As Boolean
    ShowCopyConfirmationMessage = mFlexFnObject.ShowCopyConfirmationMessage
End Property


Public Property Let CopyConfirmationMessage(nValue As String)
    mFlexFnObject.CopyConfirmationMessage = nValue
    PropertyChanged "CopyConfirmationMessage"
End Property

Public Property Get CopyConfirmationMessage() As String
    CopyConfirmationMessage = mFlexFnObject.CopyConfirmationMessage
End Property


Public Property Let AutoSelect125PercentScaleOnSmallGrids(nValue As Boolean)
    mFlexFnObject.AutoSelect125PercentScaleOnSmallGrids = nValue
    PropertyChanged "AutoSelect125PercentScaleOnSmallGrids"
End Property

Public Property Get AutoSelect125PercentScaleOnSmallGrids() As Boolean
    AutoSelect125PercentScaleOnSmallGrids = mFlexFnObject.AutoSelect125PercentScaleOnSmallGrids
End Property


Public Property Get DefaultReportStyle() As Long
    DefaultReportStyle = mFlexFnObject.DefaultReportStyle
End Property

Public Property Let DefaultReportStyle(nValue As Long)
    mFlexFnObject.DefaultReportStyle = nValue
    PropertyChanged "DefaultReportStyle"
End Property


Public Function IsShowingVerticalScrollBar(Optional nGrid As Object)
    IsShowingVerticalScrollBar = mFlexFnObject.IsShowingVerticalScrollBar(nGrid)
End Function

Public Sub EnableGridOrderByColumns(nEnabled As Boolean, Optional nGrid As Object, Optional nGridName)
    mFlexFnObject.EnableGridOrderByColumns nEnabled, nGrid, nGridName
End Sub

Public Property Get ThereAreHiddenCols(Optional nGrid As Object) As Boolean
    ThereAreHiddenCols = mFlexFnObject.ThereAreHiddenCols(nGrid)
End Property

Public Sub UnHideAllCols(Optional nGrid As Object)
    mFlexFnObject.UnHideAllCols (nGrid)
End Sub

Public Sub HideCol(nCol As Long, Optional nGrid As Object)
    mFlexFnObject.HideCol nCol, nGrid
End Sub


Public Sub DeleteOrderByColumnSaved(Optional nGrid As Object)
    'mFlexFnObject.DeleteOrderByColumnSaved (nGrid)
    mFlexFnObject.DeleteOrderByColumnSaved
End Sub


Public Property Let InitialOrderColumn(Optional nGrid As Object, nValue As Long)
    mFlexFnObject.InitialOrderColumn(nGrid) = nValue
    If nGrid Is Nothing Then
        PropertyChanged "InitialOrderColumn"
    End If
End Property

Public Property Get InitialOrderColumn(Optional nGrid As Object) As Long
    InitialOrderColumn = mFlexFnObject.InitialOrderColumn(nGrid)
End Property


Public Property Let InitialOrderDescending(Optional nGrid As Object, nValue As Boolean)
    If nValue <> mFlexFnObject.InitialOrderDescending(nGrid) Then
        mFlexFnObject.InitialOrderDescending(nGrid) = nValue
        If nGrid Is Nothing Then
            PropertyChanged "InitialOrderDescending"
        End If
    End If
End Property

Public Property Get InitialOrderDescending(Optional nGrid As Object) As Boolean
    InitialOrderDescending = mFlexFnObject.InitialOrderDescending(nGrid)
End Property


Public Property Let SameDataGroupedInColumns(Optional nGrid As Object, nValue As Boolean)
    Dim iCancel As Boolean
    Dim iLng As Long
    Dim iGridName As String
    
    If nValue <> mFlexFnObject.SameDataGroupedInColumns(nGrid) Then
        If Not nGrid Is Nothing Then
            iGridName = nGrid.Name
        End If
        
        iLng = CLng(nValue)
        RaiseEvent BeforeAction("GroupData", iGridName, iLng, iCancel)
        If Not iCancel Then
            mFlexFnObject.SameDataGroupedInColumns(nGrid) = CBool(iLng)
            SaveSetting AppNameForRegistry, "Preferences", mFlexFnObject.Context & "_DataGrouped", iLng
            RaiseEvent AfterAction("GroupData", iGridName, iLng)
        End If
        
        If nGrid Is Nothing Then
            PropertyChanged "SameDataGroupedInColumns"
        End If
    End If
End Property

Public Property Get SameDataGroupedInColumns(Optional nGrid As Object) As Boolean
    SameDataGroupedInColumns = mFlexFnObject.SameDataGroupedInColumns(nGrid)
End Property


Public Property Let ShowToolTipsOnLongerCellTexts(Optional nGrid As Object, nValue As Boolean)
    If nValue <> mFlexFnObject.ShowToolTipsOnLongerCellTexts(nGrid) Then
        mFlexFnObject.ShowToolTipsOnLongerCellTexts(nGrid) = nValue
        If nGrid Is Nothing Then
            PropertyChanged "ShowToolTipsOnLongerCellTexts"
        End If
    End If
End Property

Public Property Get ShowToolTipsOnLongerCellTexts(Optional nGrid As Object) As Boolean
    ShowToolTipsOnLongerCellTexts = mFlexFnObject.ShowToolTipsOnLongerCellTexts(nGrid)
End Property


Public Property Let AllowTextEdition(Optional nGrid As Object, nValue As Boolean)
    If nValue <> mFlexFnObject.AllowTextEdition(nGrid) Then
        mFlexFnObject.AllowTextEdition(nGrid) = nValue
        If nGrid Is Nothing Then
            PropertyChanged "AllowTextEdition"
        End If
    End If
End Property

Public Property Get AllowTextEdition(Optional nGrid As Object) As Boolean
    AllowTextEdition = mFlexFnObject.AllowTextEdition(nGrid)
End Property


Public Property Let TextEditionLocked(Optional nGrid As Object, nValue As Boolean)
    If nValue <> mFlexFnObject.TextEditionLocked(nGrid) Then
        mFlexFnObject.TextEditionLocked(nGrid) = nValue
        If nGrid Is Nothing Then
            PropertyChanged "TextEditionLocked"
        End If
    End If
End Property

Public Property Get TextEditionLocked(Optional nGrid As Object) As Boolean
    TextEditionLocked = mFlexFnObject.TextEditionLocked(nGrid)
End Property


Public Property Let DoNotRememberOrder(Optional nGrid As Object, nValue As Boolean)
    If nValue <> mFlexFnObject.DoNotRememberOrder(nGrid) Then
        mFlexFnObject.DoNotRememberOrder(nGrid) = nValue
        If nGrid Is Nothing Then
            PropertyChanged "DoNotRememberOrder"
        End If
    End If
End Property

Public Property Get DoNotRememberOrder(Optional nGrid As Object) As Boolean
    DoNotRememberOrder = mFlexFnObject.DoNotRememberOrder(nGrid)
End Property


Public Property Let ShowToolTipsForOrderColumns(Optional nGrid As Object, nValue As Boolean)
    If nValue <> mFlexFnObject.ShowToolTipsForOrderColumns(nGrid) Then
        mFlexFnObject.ShowToolTipsForOrderColumns(nGrid) = nValue
        If nGrid Is Nothing Then
            PropertyChanged "ShowToolTipsForOrderColumns"
        End If
    End If
End Property

Public Property Get ShowToolTipsForOrderColumns(Optional nGrid As Object) As Boolean
    ShowToolTipsForOrderColumns = mFlexFnObject.ShowToolTipsForOrderColumns(nGrid)
End Property


Public Property Let GridsFlatAppearance(nValue As Boolean)
    If nValue <> mFlexFnObject.GridsFlatAppearance Then
        mFlexFnObject.GridsFlatAppearance = nValue
        PropertyChanged "GridsFlatAppearance"
    End If
End Property

Public Property Get GridsFlatAppearance() As Boolean
    GridsFlatAppearance = mFlexFnObject.GridsFlatAppearance
End Property


Public Property Let EnableOrderByColumns(Optional nGrid As Object, nValue As Boolean)
    If nValue <> mFlexFnObject.EnableOrderByColumns(nGrid) Then
        mFlexFnObject.EnableOrderByColumns(nGrid) = nValue
        If nGrid Is Nothing Then
            PropertyChanged "EnableOrderByColumns"
        End If
    End If
End Property

Public Property Get EnableOrderByColumns(Optional nGrid As Object) As Boolean
    EnableOrderByColumns = mFlexFnObject.EnableOrderByColumns(nGrid)
End Property


Public Property Let StretchColumnsWidthsToFill(Optional nGrid As Object, nValue As Boolean)
    If nValue <> mFlexFnObject.StretchColumnsWidthsToFill(nGrid) Then
        mFlexFnObject.StretchColumnsWidthsToFill(nGrid) = nValue
        If nGrid Is Nothing Then
            PropertyChanged "StretchColumnsWidthsToFill"
        End If
    End If
End Property

Public Property Get StretchColumnsWidthsToFill(Optional nGrid As Object) As Boolean
    StretchColumnsWidthsToFill = mFlexFnObject.StretchColumnsWidthsToFill(nGrid)
End Property

Public Sub UpdateGridColumnsWidthsStretched(Optional nGrid As Object)
    mFlexFnObject.UpdateGridColumnsWidthsStretched nGrid
End Sub

Public Property Let BorderColor(Optional nGrid As Object, nValue As Long)
    If nValue <> mFlexFnObject.BorderColor(nGrid) Then
        mFlexFnObject.BorderColor(nGrid) = nValue
        If nGrid Is Nothing Then
            PropertyChanged "BorderColor"
        End If
    End If
End Property

Public Property Get BorderColor(Optional nGrid As Object) As Long
    BorderColor = mFlexFnObject.BorderColor(nGrid)
End Property


Public Property Let BorderWidth(Optional nGrid As Object, nValue As Long)
    If nValue <> mFlexFnObject.BorderWidth(nGrid) Then
        mFlexFnObject.BorderWidth(nGrid) = nValue
        If nGrid Is Nothing Then
            PropertyChanged "BorderWidth"
        End If
    End If
End Property

Public Property Get BorderWidth(Optional nGrid As Object) As Long
    BorderWidth = mFlexFnObject.BorderWidth(nGrid)
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

Public Property Get ReportStyle() As GridReportStyle
    Set ReportStyle = mFlexFnObject.ReportStyle
End Property

Public Property Get MenuCount() As Long
    MenuCount = 10
End Property

Public Property Get MenuCaption(nIndex As Long)
    Select Case nIndex
        Case 0 ' mnuPrint
            MenuCaption = mPrintButton_ToolTipText
        Case 1 ' mnuSep1
            MenuCaption = "-"
        Case 2 ' mnuCopyAll
            MenuCaption = mCopyButton_ToolTipText 'mCopyAllMenuCaption
        Case 3 ' mnuSave
            MenuCaption = mSaveButton_ToolTipText
        Case 4 ' mnuFind
            MenuCaption = mFindButton_ToolTipText
        Case 5 ' mnuSep2
            MenuCaption = "-"
        Case 6 ' mnuGroupData
            If mGroupDataButtonVisible Then
                If mFlexFnObject.ThereAreHiddenCols(Grid) Then
                    MenuCaption = mConfigColumnsButtonColsHidden_ToolTipText
                Else
                    MenuCaption = mConfigColumnsButton_ToolTipText
                End If
            Else
                MenuCaption = " "
            End If
        Case 7 ' mnuConfigColumns
            If mConfigColumnsButtonVisible Then
                If SameDataGroupedInColumns(Grid) Then
                    MenuCaption = mGroupDataButtonPressed_ToolTipText
                Else
                    MenuCaption = mGroupDataButton_ToolTipText
                End If
            Else
                MenuCaption = " "
            End If
        Case 8 ' mnuSep3
            MenuCaption = "-"
    End Select
End Property

Public Property Get MenuVisible(nIndex As Long) As Boolean
    On Error Resume Next
    MenuVisible = Controls(GetMenuNameByIndex(nIndex)).Visible
End Property

Public Property Get MenuEnabled(nIndex As Long) As Boolean
    On Error Resume Next
    MenuEnabled = Controls(GetMenuNameByIndex(nIndex)).Enabled
End Property

Private Function GetMenuNameByIndex(nIndex As Long) As String
    Select Case nIndex
        Case 0 ' mnuPrint
            GetMenuNameByIndex = "mnuPrint"
        Case 1 ' mnuSep1
            GetMenuNameByIndex = "mnuSep1"
        Case 2 ' mnuCopyAll
            GetMenuNameByIndex = "mnuCopyAll"
        Case 3 ' mnuSave
            GetMenuNameByIndex = "mnuSave"
        Case 4 ' mnuFind
            GetMenuNameByIndex = "mnuFind"
        Case 5 ' mnuSep2
            GetMenuNameByIndex = "mnuSep2"
        Case 6 ' mnuGroupData
            GetMenuNameByIndex = "mnuGroupData"
        Case 7 ' mnuConfigColumns
            GetMenuNameByIndex = "mnuConfigColumns"
        Case 8 ' mnuSep3
            GetMenuNameByIndex = "mnuSep3"
    End Select
End Function

Public Sub MenuClick(nIndex As Long)
    Select Case nIndex
        Case 0 ' mnuPrint
            mnuPrint_Click
        Case 1 ' mnuSep1
            '
        Case 2 ' mnuCopyAll
            mnuCopyAll_Click
        Case 3 ' mnuSave
            mnuSave_Click
        Case 4 ' mnuFind
            mnuFind_Click
        Case 5 ' mnuSep2
            '
        Case 6 ' mnuGroupData
            mnuGroupData_Click
        Case 7 ' mnuConfigColumns
            mnuConfigColumns_Click
        Case 8 ' mnuSep3
            '
    End Select
End Sub

Public Function BuildMenu() As Boolean
    Dim iGrid As Object
    Dim iGridHasData As Boolean
    Dim iCancel As Boolean
    
    Set iGrid = mFlexFnObject.GetWorkGridControl
    If Not iGrid Is Nothing Then
        On Error Resume Next
        mGridMouseRowAtPopupMenuPoint = iGrid.MouseRow
        mGridMouseColAtPopupMenuPoint = iGrid.MouseCol
        On Error GoTo 0
        ResetCustomPopupMenuItems
        CheckWhatFunctionsToMakeAvailable iGridHasData
        RaiseEvent BeforeShowingPopupMenu(iGrid.Name, iGridHasData, iCancel)
        If BuildPopupMenu(iGrid) Then
            If Not iCancel Then
                Set mGridPopup = iGrid
                BuildMenu = True
            End If
        End If
    End If
End Function

Public Property Get ButtonsCount() As Long
    ButtonsCount = MenuCount
End Property

Public Property Get ButtonText(nIndex As Long) As String
    ButtonText = MenuCaption(nIndex)
End Property

Public Property Get ButtonVisible(nIndex As Long) As Boolean
    Select Case nIndex
        Case 0 ' mnuPrint
            ButtonVisible = mPrintButtonVisible
        Case 1 ' mnuSep1
            ButtonVisible = mPrintButtonVisible
            If ButtonVisible Then
                If Not ButtonVisible(9) Then
                    If Not ButtonVisible(6) Then
                        ButtonVisible = False
                    End If
                End If
            End If
        Case 2 ' mnuCopyAll
            ButtonVisible = mCopyButtonVisible
        Case 3 ' mnuSave
            ButtonVisible = mSaveButtonVisible
        Case 4 ' mnuFind
            ButtonVisible = mFindButtonVisible
        Case 5 ' mnuSep2
            ButtonVisible = (mFindButtonVisible Or mCopyButtonVisible Or mSaveButtonVisible)
            If ButtonVisible Then
                If Not ButtonVisible(9) Then
                    ButtonVisible = False
                End If
            End If
        Case 6 ' mnuGroupData
            ButtonVisible = mGroupDataButtonVisible
        Case 7 ' mnuConfigColumns
            ButtonVisible = mConfigColumnsButtonVisible
        Case 8 ' mnuSep3
            ButtonVisible = (mGroupDataButtonVisible Or mConfigColumnsButtonVisible)
            If ButtonVisible Then
                ButtonVisible = False
            End If
    End Select
End Property

Public Property Get ButtonEnabled(nIndex As Long) As Boolean
    Select Case nIndex
        Case 0 ' mnuPrint
            ButtonEnabled = mPrintButtonEnabled
        Case 1 ' mnuSep1
            ButtonEnabled = True
        Case 2 ' mnuCopyAll
            ButtonEnabled = mCopyButtonEnabled
        Case 3 ' mnuSave
            ButtonEnabled = mSaveButtonEnabled
        Case 4 ' mnuFind
            ButtonEnabled = mFindButtonEnabled
        Case 5 ' mnuSep2
            ButtonEnabled = True
        Case 6 ' mnuGroupData
            ButtonEnabled = mGroupDataButtonEnabled
        Case 7 ' mnuConfigColumns
            ButtonEnabled = mConfigColumnsButtonEnabled
        Case 8 ' mnuSep3
            ButtonEnabled = True
    End Select
End Property

Public Sub ButtonClick(nIndex As Long)
    Dim iButton As ToolBarDAButton
    Dim iKey As String
    
    Select Case nIndex
        Case 0 ' mnuPrint
             iKey = "Print"
        Case 1 ' mnuSep1
            '
        Case 2 ' mnuCopyAll
            iKey = "Copy"
        Case 3 ' mnuSave
            iKey = "Save"
        Case 4 ' mnuFind
            iKey = "Find"
        Case 5 ' mnuSep2
            '
        Case 6 ' mnuGroupData
            iKey = "GroupData"
        Case 7 ' mnuConfigColumns
            iKey = "ConfigColumns"
        Case 8 ' mnuSep3
            '
    End Select
    Set iButton = tbrButtons.Buttons(iKey)
    tbrButtons_ButtonClick iButton
End Sub

Public Property Get PrintFnObject() As PrintFnObject
    Set PrintFnObject = mFlexFnObject.PrintFnObject
End Property

Public Sub OrderGridByColumn(Optional nGrid As Object, Optional nColumn As Long = -1, Optional nDescending)
    mFlexFnObject.OrderGridByColumn nGrid, nColumn, nDescending
End Sub

Public Sub UpdateOrderByColumn(Optional nGrid As Object)
    mFlexFnObject.UpdateOrderByColumn nGrid
End Sub

Public Property Get GetOrderColumn(Optional nGrid As Object) As Long
    GetOrderColumn = mFlexFnObject.GetOrderColumn(nGrid)
End Property

Public Property Get GetOrderColumnDescending(Optional nGrid As Object) As Boolean
    GetOrderColumnDescending = mFlexFnObject.GetOrderColumnDescending(nGrid)
End Property

Public Property Get EditingCell(Optional nGrid As Object) As Boolean
    EditingCell = mFlexFnObject.EditingCell(nGrid)
End Property

Public Sub CopyTextEditingCell(Optional nGrid As Object)
    mFlexFnObject.CopyTextEditingCell nGrid
End Sub

Public Property Let IconsSize(nValue As vbExToolbarDAIconsSizeConstants)
    If nValue <> mIconsSize Then
        PropertyChanged "IconsSize"
        mIconsSize = nValue
        tbrButtons.IconsSize = mIconsSize
        UserControl_Resize
    End If
End Property

Public Property Get IconsSize() As vbExToolbarDAIconsSizeConstants
    IconsSize = mIconsSize
End Property

Private Sub ResetCustomPopupMenuItems()
    ReDim mCustomPopupMenuItems_Names(0)
    ReDim mCustomPopupMenuItems_Captions(0)
    ReDim mCustomPopupMenuItems_Enabled(0)
    ReDim mCustomPopupMenuItems_Checked(0)
    ReDim mCustomPopupMenuItems_Before(0)
End Sub

Public Sub AddCustomPopupMenuItem(ItemName As String, Caption As String, Enabled As Boolean, PlaceBefore As Boolean, Optional Checked As Boolean)
    Dim i As Long
    
    i = UBound(mCustomPopupMenuItems_Names) + 1
    
    ReDim Preserve mCustomPopupMenuItems_Names(i)
    ReDim Preserve mCustomPopupMenuItems_Captions(i)
    ReDim Preserve mCustomPopupMenuItems_Enabled(i)
    ReDim Preserve mCustomPopupMenuItems_Checked(i)
    ReDim Preserve mCustomPopupMenuItems_Before(i)
    
    mCustomPopupMenuItems_Names(i) = ItemName
    mCustomPopupMenuItems_Captions(i) = Caption
    mCustomPopupMenuItems_Enabled(i) = Enabled
    mCustomPopupMenuItems_Checked(i) = Checked
    mCustomPopupMenuItems_Before(i) = PlaceBefore
    
End Sub

Public Property Get GridMouseRowAtPopupMenuPoint() As Long
    GridMouseRowAtPopupMenuPoint = mGridMouseRowAtPopupMenuPoint
End Property

Public Property Get GridMouseColAtPopupMenuPoint() As Long
    GridMouseColAtPopupMenuPoint = mGridMouseColAtPopupMenuPoint
End Property


' PrintFnObject properties

Public Property Let MinScalePercent(nValue As Long)
    If nValue <> mFlexFnObject.MinScalePercent Then
        mFlexFnObject.MinScalePercent = nValue
    End If
End Property

Public Property Get MinScalePercent() As Long
    MinScalePercent = mFlexFnObject.MinScalePercent
End Property


Public Property Let MaxScalePercent(nValue As Long)
    If nValue <> mFlexFnObject.MaxScalePercent Then
        mFlexFnObject.MaxScalePercent = nValue
    End If
End Property

Public Property Get MaxScalePercent() As Long
    MaxScalePercent = mFlexFnObject.MaxScalePercent
End Property


Public Property Let PrintPrevUseAltScaleIcons(nValue As Boolean)
    If nValue <> mFlexFnObject.PrintPrevUseAltScaleIcons Then
        mFlexFnObject.PrintPrevUseAltScaleIcons = nValue
        PropertyChanged "PrintPrevUseAltScaleIcons"
    End If
End Property

Public Property Get PrintPrevUseAltScaleIcons() As Boolean
    PrintPrevUseAltScaleIcons = mFlexFnObject.PrintPrevUseAltScaleIcons
End Property


Public Property Let PrintCellsFormatting(nValue As vbExPrintCellsFormatting)
    If nValue <> mFlexFnObject.PrintCellsFormatting Then
        mFlexFnObject.PrintCellsFormatting = nValue
        PropertyChanged "PrintCellsFormatting"
    End If
End Property

Public Property Get PrintCellsFormatting() As vbExPrintCellsFormatting
    PrintCellsFormatting = mFlexFnObject.PrintCellsFormatting
End Property


Public Property Let PaperSize(nValue As cdePaperSizeConstants)
    If nValue <> mFlexFnObject.PrintFnObject.PaperSize Then
        mFlexFnObject.PrintFnObject.PaperSize = nValue
        PropertyChanged "PaperSize"
    End If
End Property

Public Property Get PaperSize() As cdePaperSizeConstants
Attribute PaperSize.VB_MemberFlags = "400"
    PaperSize = mFlexFnObject.PrintFnObject.PaperSize
End Property


Public Property Let PaperBin(nValue As cdePaperBinConstants)
    If nValue <> mFlexFnObject.PrintFnObject.PaperBin Then
        mFlexFnObject.PrintFnObject.PaperBin = nValue
        PropertyChanged "PaperBin"
    End If
End Property

Public Property Get PaperBin() As cdePaperBinConstants
Attribute PaperBin.VB_MemberFlags = "400"
    PaperBin = mFlexFnObject.PrintFnObject.PaperBin
End Property


Public Property Let PrintQuality(nValue As cdePrintQualityConstants)
    If nValue <> mFlexFnObject.PrintFnObject.PrintQuality Then
        mFlexFnObject.PrintFnObject.PrintQuality = nValue
        PropertyChanged "PrintQuality"
    End If
End Property

Public Property Get PrintQuality() As cdePrintQualityConstants
Attribute PrintQuality.VB_MemberFlags = "400"
    PrintQuality = mFlexFnObject.PrintFnObject.PrintQuality
End Property


Public Property Let ColorMode(nValue As cdeColorModeConstants)
    If nValue <> mFlexFnObject.PrintFnObject.ColorMode Then
        mFlexFnObject.PrintFnObject.ColorMode = nValue
        PropertyChanged "ColorMode"
    End If
End Property

Public Property Get ColorMode() As cdeColorModeConstants
    ColorMode = mFlexFnObject.PrintFnObject.ColorMode
End Property


Public Property Let Duplex(nValue As cdeDuplexConstants)
    If nValue <> mFlexFnObject.PrintFnObject.Duplex Then
        mFlexFnObject.PrintFnObject.Duplex = nValue
        PropertyChanged "Duplex"
    End If
End Property

Public Property Get Duplex() As cdeDuplexConstants
Attribute Duplex.VB_MemberFlags = "400"
    Duplex = mFlexFnObject.PrintFnObject.Duplex
End Property


Public Property Let LeftMargin(nValue As Single)
    If nValue <> mFlexFnObject.PrintFnObject.LeftMargin Then
        mFlexFnObject.PrintFnObject.LeftMargin = nValue
        PropertyChanged "LeftMargin"
    End If
End Property

Public Property Get LeftMargin() As Single
    LeftMargin = mFlexFnObject.PrintFnObject.LeftMargin
End Property


Public Property Let RightMargin(nValue As Single)
    If nValue <> mFlexFnObject.PrintFnObject.RightMargin Then
        mFlexFnObject.PrintFnObject.RightMargin = nValue
        PropertyChanged "RightMargin"
    End If
End Property

Public Property Get RightMargin() As Single
    RightMargin = mFlexFnObject.PrintFnObject.RightMargin
End Property


Public Property Let TopMargin(nValue As Single)
    If nValue <> mFlexFnObject.PrintFnObject.TopMargin Then
        mFlexFnObject.PrintFnObject.TopMargin = nValue
        PropertyChanged "TopMargin"
    End If
End Property

Public Property Get TopMargin() As Single
    TopMargin = mFlexFnObject.PrintFnObject.TopMargin
End Property


Public Property Let BottomMargin(nValue As Single)
    If nValue <> mFlexFnObject.PrintFnObject.BottomMargin Then
        mFlexFnObject.PrintFnObject.BottomMargin = nValue
        PropertyChanged "BottomMargin"
    End If
End Property

Public Property Get BottomMargin() As Single
    BottomMargin = mFlexFnObject.PrintFnObject.BottomMargin
End Property


Public Property Let MinLeftMargin(nValue As Single)
    If nValue <> mFlexFnObject.PrintFnObject.MinLeftMargin Then
        mFlexFnObject.PrintFnObject.MinLeftMargin = nValue
        PropertyChanged "MinLeftMargin"
    End If
End Property

Public Property Get MinLeftMargin() As Single
    MinLeftMargin = mFlexFnObject.PrintFnObject.MinLeftMargin
End Property


Public Property Let MinRightMargin(nValue As Single)
    If nValue <> mFlexFnObject.PrintFnObject.MinRightMargin Then
        mFlexFnObject.PrintFnObject.MinRightMargin = nValue
        PropertyChanged "MinRightMargin"
    End If
End Property

Public Property Get MinRightMargin() As Single
    MinRightMargin = mFlexFnObject.PrintFnObject.MinRightMargin
End Property


Public Property Let MinTopMargin(nValue As Single)
    If nValue <> mFlexFnObject.PrintFnObject.MinTopMargin Then
        mFlexFnObject.PrintFnObject.MinTopMargin = nValue
        PropertyChanged "MinTopMargin"
    End If
End Property

Public Property Get MinTopMargin() As Single
    MinTopMargin = mFlexFnObject.PrintFnObject.MinTopMargin
End Property


Public Property Let MinBottomMargin(nValue As Single)
    If nValue <> mFlexFnObject.PrintFnObject.MinBottomMargin Then
        mFlexFnObject.PrintFnObject.MinBottomMargin = nValue
        PropertyChanged "MinBottomMargin"
    End If
End Property

Public Property Get MinBottomMargin() As Single
    MinBottomMargin = mFlexFnObject.PrintFnObject.MinBottomMargin
End Property


Public Property Get Units() As cdeUnits
    Units = mFlexFnObject.PrintFnObject.Units
End Property

Public Property Let Units(nValue As cdeUnits)
    If nValue <> mFlexFnObject.PrintFnObject.Units Then
        mFlexFnObject.PrintFnObject.Units = nValue
        PropertyChanged "Units"
    End If
End Property


Public Property Get UnitsForUser() As cdeUnitsForUser
    UnitsForUser = mFlexFnObject.PrintFnObject.UnitsForUser
End Property

Public Property Let UnitsForUser(nValue As cdeUnitsForUser)
    If nValue <> mFlexFnObject.PrintFnObject.UnitsForUser Then
        mFlexFnObject.PrintFnObject.UnitsForUser = nValue
        PropertyChanged "UnitsForUser"
    End If
End Property


Public Property Get PrintPageNumbers() As Boolean
    PrintPageNumbers = mFlexFnObject.PrintFnObject.PrintPageNumbers
End Property

Public Property Let PrintPageNumbers(nValue As Boolean)
    If nValue <> mFlexFnObject.PrintFnObject.PrintPageNumbers Then
        mFlexFnObject.PrintFnObject.PrintPageNumbers = nValue
        PropertyChanged "PrintPageNumbers"
    End If
End Property


Public Property Get PageNumbersPosition() As vbExPageNumbersPositionConstants
    PageNumbersPosition = mFlexFnObject.PrintFnObject.PageNumbersPosition
End Property

Public Property Let PageNumbersPosition(nValue As vbExPageNumbersPositionConstants)
    If nValue <> mFlexFnObject.PrintFnObject.PageNumbersPosition Then
        mFlexFnObject.PrintFnObject.PageNumbersPosition = nValue
        PropertyChanged "PageNumbersPosition"
    End If
End Property


Public Property Get PageNumbersFormat() As String
    PageNumbersFormat = mFlexFnObject.PrintFnObject.PageNumbersFormat
End Property

Public Property Let PageNumbersFormat(nValue As String)
    If nValue <> mFlexFnObject.PrintFnObject.PageNumbersFormat Then
        mFlexFnObject.PrintFnObject.PageNumbersFormat = nValue
        PropertyChanged "PageNumbersFormat"
    End If
End Property


Public Property Get AllowUserChangeScale() As Boolean
    AllowUserChangeScale = mFlexFnObject.PrintFnObject.AllowUserChangeScale
End Property

Public Property Let AllowUserChangeScale(nValue As Boolean)
    If nValue <> mFlexFnObject.PrintFnObject.AllowUserChangeScale Then
        mFlexFnObject.PrintFnObject.AllowUserChangeScale = nValue
        PropertyChanged "AllowUserChangeScale"
    End If
End Property


Public Property Get AllowUserChangeOrientation() As Boolean
    AllowUserChangeOrientation = mFlexFnObject.PrintFnObject.AllowUserChangeOrientation
End Property

Public Property Let AllowUserChangeOrientation(nValue As Boolean)
    If nValue <> mFlexFnObject.PrintFnObject.AllowUserChangeOrientation Then
        mFlexFnObject.PrintFnObject.AllowUserChangeOrientation = nValue
        PropertyChanged "AllowUserChangeOrientation"
    End If
End Property


Public Property Get AllowUserChangePaper() As Boolean
    AllowUserChangePaper = mFlexFnObject.PrintFnObject.AllowUserChangePaper
End Property

Public Property Let AllowUserChangePaper(nValue As Boolean)
    If nValue <> mFlexFnObject.PrintFnObject.AllowUserChangePaper Then
        mFlexFnObject.PrintFnObject.AllowUserChangePaper = nValue
        PropertyChanged "AllowUserChangePaper"
    End If
End Property


Public Property Get PrintPrevUseOneToolBar() As Boolean
    PrintPrevUseOneToolBar = mFlexFnObject.PrintFnObject.PrintPrevUseOneToolBar
End Property

Public Property Let PrintPrevUseOneToolBar(nValue As Boolean)
    If nValue <> mFlexFnObject.PrintFnObject.PrintPrevUseOneToolBar Then
        mFlexFnObject.PrintFnObject.PrintPrevUseOneToolBar = nValue
        PropertyChanged "PrintPrevUseOneToolBar"
    End If
End Property


Public Property Let DocumentName(nDocName As String)
    If nDocName <> mFlexFnObject.PrintFnObject.DocumentName Then
        mFlexFnObject.PrintFnObject.DocumentName = nDocName
        PropertyChanged "DocumentName"
    End If
End Property

Public Property Get DocumentName() As String
    DocumentName = mFlexFnObject.PrintFnObject.DocumentName
End Property


Public Property Get FormatButtonVisible() As Boolean
Attribute FormatButtonVisible.VB_MemberFlags = "400"
    FormatButtonVisible = mFlexFnObject.PrintFnObject.FormatButtonVisible
End Property

Public Property Let FormatButtonVisible(nValue As Boolean)
    If nValue <> mFlexFnObject.PrintFnObject.FormatButtonVisible Then
        mFlexFnObject.PrintFnObject.FormatButtonVisible = nValue
        PropertyChanged "FormatButtonVisible"
    End If
End Property


Public Property Get PageSetupButtonVisible() As Boolean
Attribute PageSetupButtonVisible.VB_MemberFlags = "400"
    PageSetupButtonVisible = mFlexFnObject.PrintFnObject.PageSetupButtonVisible
End Property

Public Property Let PageSetupButtonVisible(nValue As Boolean)
    If nValue <> mFlexFnObject.PrintFnObject.PageSetupButtonVisible Then
        mFlexFnObject.PrintFnObject.PageSetupButtonVisible = nValue
        PropertyChanged "PageSetupButtonVisible"
    End If
End Property


Public Property Let PrintPrevToolBarIconsSize(nValue As vbExPrintPrevToolBarIconsSizeConstants)
    If nValue <> mFlexFnObject.PrintFnObject.PrintPrevToolBarIconsSize Then
        mFlexFnObject.PrintFnObject.PrintPrevToolBarIconsSize = nValue
        PropertyChanged "PrintPrevToolBarIconsSize"
    End If
End Property

Public Property Get PrintPrevToolBarIconsSize() As vbExPrintPrevToolBarIconsSizeConstants
Attribute PrintPrevToolBarIconsSize.VB_MemberFlags = "400"
    PrintPrevToolBarIconsSize = mFlexFnObject.PrintFnObject.PrintPrevToolBarIconsSize
End Property


Public Property Set PageNumbersFont(ByVal nFont As StdFont)
    If Not nFont Is mFlexFnObject.PrintFnObject.PageNumbersFont Then
        Set mFlexFnObject.PrintFnObject.PageNumbersFont = nFont
        PropertyChanged "PageNumbersFont"
    End If
End Property

Public Property Let PageNumbersFont(ByVal nFont As StdFont)
    Set PageNumbersFont = nFont
End Property

Public Property Get PageNumbersFont() As StdFont
    Set PageNumbersFont = mFlexFnObject.PrintFnObject.PageNumbersFont
End Property


Public Property Get PageNumbersForeColor() As OLE_COLOR
    PageNumbersForeColor = mFlexFnObject.PrintFnObject.PageNumbersForeColor
End Property

Public Property Let PageNumbersForeColor(ByVal nValue As OLE_COLOR)
    If mFlexFnObject.PrintFnObject.PageNumbersForeColor <> nValue Then
        mFlexFnObject.PrintFnObject.PageNumbersForeColor = nValue
        PropertyChanged "PageNumbersForeColor"
    End If
End Property


Public Function GetPredefinedPageNumbersFormatString(nIndex As Long) As String
    GetPredefinedPageNumbersFormatString = mFlexFnObject.PrintFnObject.GetPredefinedPageNumbersFormatString(nIndex)
End Function

Public Property Get GetPredefinedPageNumbersFormatStringsCount() As Long
    GetPredefinedPageNumbersFormatStringsCount = mFlexFnObject.PrintFnObject.GetPredefinedPageNumbersFormatStringsCount
End Property


Public Property Let RememberUserPrintingPreferences(nValue As gfnRememberUserPrintingPreferences)
    If nValue <> mFlexFnObject.RememberUserPrintingPreferences Then
        mFlexFnObject.RememberUserPrintingPreferences = nValue
        PropertyChanged "RememberUserPrintingPreferences"
    End If
End Property

' End PrintFnObject properties


Public Property Get RememberUserPrintingPreferences() As gfnRememberUserPrintingPreferences
    RememberUserPrintingPreferences = mFlexFnObject.RememberUserPrintingPreferences
End Property


Public Property Get FlexFnObject() As FlexFnObject
    Set FlexFnObject = mFlexFnObject
End Property

