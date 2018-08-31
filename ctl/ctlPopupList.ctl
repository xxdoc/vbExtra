VERSION 5.00
Begin VB.UserControl PopupList 
   ClientHeight    =   8916
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6432
   PropertyPages   =   "ctlPopupList.ctx":0000
   ScaleHeight     =   8916
   ScaleWidth      =   6432
   ToolboxBitmap   =   "ctlPopupList.ctx":0023
   Begin VB.PictureBox picBackSelectedItem_Default 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   372
      Left            =   3600
      Picture         =   "ctlPopupList.ctx":0335
      ScaleHeight     =   372
      ScaleWidth      =   1812
      TabIndex        =   18
      Top             =   1440
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.PictureBox picBackground_Default 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   372
      Left            =   3564
      Picture         =   "ctlPopupList.ctx":50CF
      ScaleHeight     =   372
      ScaleWidth      =   1812
      TabIndex        =   17
      Top             =   504
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.PictureBox picAux_picMouseOver 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   372
      Left            =   4200
      Picture         =   "ctlPopupList.ctx":8951
      ScaleHeight     =   372
      ScaleWidth      =   1812
      TabIndex        =   16
      Top             =   7680
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.PictureBox picAux_picNormal 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   372
      Left            =   4200
      Picture         =   "ctlPopupList.ctx":D6EB
      ScaleHeight     =   372
      ScaleWidth      =   1812
      TabIndex        =   15
      Top             =   7080
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.PictureBox picBackSelectedItem 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   372
      Left            =   3600
      ScaleHeight     =   372
      ScaleWidth      =   1812
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   372
      Left            =   3564
      ScaleHeight     =   372
      ScaleWidth      =   1812
      TabIndex        =   13
      Top             =   36
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.Timer tmrTransparency 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   660
      Top             =   8340
   End
   Begin VB.PictureBox picRegion 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   3780
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   11
      Top             =   7140
      Width           =   270
   End
   Begin VB.Timer tmrMouseOverCheck 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   180
      Top             =   8340
   End
   Begin VB.PictureBox picNormal 
      Height          =   435
      Left            =   600
      ScaleHeight     =   384
      ScaleWidth      =   2952
      TabIndex        =   10
      Top             =   7110
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox picMouseOver 
      Height          =   465
      Left            =   600
      ScaleHeight     =   420
      ScaleWidth      =   2952
      TabIndex        =   9
      Top             =   7650
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox picPopupList 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   0
      ScaleHeight     =   6372
      ScaleWidth      =   3468
      TabIndex        =   2
      Top             =   500
      Width           =   3465
      Begin VB.VScrollBar VScroll1 
         Height          =   6360
         Left            =   3120
         TabIndex        =   8
         Top             =   300
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.PictureBox picContainer 
         BorderStyle     =   0  'None
         Height          =   6360
         Left            =   0
         ScaleHeight     =   6360
         ScaleWidth      =   3300
         TabIndex        =   3
         Top             =   0
         Width           =   3300
         Begin VB.PictureBox picSelectedItem 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   0
            ScaleHeight     =   492
            ScaleWidth      =   3156
            TabIndex        =   6
            Top             =   480
            Width           =   3150
            Begin VB.Shape shpSelectedItem 
               BorderColor     =   &H00C0C0C0&
               FillColor       =   &H00D9BB45&
               FillStyle       =   0  'Solid
               Height          =   135
               Left            =   75
               Shape           =   3  'Circle
               Top             =   150
               Visible         =   0   'False
               Width           =   135
            End
            Begin VB.Label lblSelectedItem 
               BackStyle       =   0  'Transparent
               Caption         =   "# Selected item"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   264
               Left            =   240
               TabIndex        =   7
               Top             =   120
               Width           =   2892
            End
         End
         Begin VB.PictureBox picItem 
            BackColor       =   &H0000FFFF&
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   0
            Left            =   90
            ScaleHeight     =   492
            ScaleWidth      =   3000
            TabIndex        =   4
            Top             =   0
            Width           =   3000
            Begin VB.Shape shpIem 
               BorderColor     =   &H00C0C0C0&
               FillColor       =   &H006B6B6B&
               FillStyle       =   0  'Solid
               Height          =   135
               Index           =   0
               Left            =   75
               Shape           =   3  'Circle
               Top             =   165
               Visible         =   0   'False
               Width           =   135
            End
            Begin VB.Label lblItem 
               BackStyle       =   0  'Transparent
               Caption         =   "# Item"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   240
               TabIndex        =   5
               Top             =   120
               Width           =   2595
            End
         End
      End
   End
   Begin VB.PictureBox picText 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   0
      Top             =   0
      Width           =   3030
      Begin vbExtra.ButtonEx btnDropDown 
         Height          =   255
         Left            =   2685
         TabIndex        =   12
         Top             =   45
         Width           =   285
         _ExtentX        =   508
         _ExtentY        =   445
         ButtonStyle     =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14932157
         Caption         =   "€"
         CheckBoxMode    =   -1  'True
         UseMaskCOlor    =   -1  'True
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "# Text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   240
         TabIndex        =   1
         Top             =   60
         Width           =   2892
      End
   End
End
Attribute VB_Name = "PopupList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

Private Declare Function ColorRGBToHLS Lib "shlwapi.dll" (ByVal clrRGB As Long, pwHue As Long, pwLuminance As Long, pwSaturation As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const LWA_ALPHA = &H2&
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_HWNDPARENT As Long = (-8)
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Const WM_WINDOWPOSCHANGED = &H47&
Private Const RGN_OR = 2&
Private Const RGN_DIFF = 4&

Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long

Public Event DropDown()
Public Event Click()
Attribute Click.VB_MemberFlags = "200"
Public Event ListHided()
Public Event ItemClick()

Private mText As String
Private mItems() As String
Private mItemData() As String
Private mNewIndex As Long
Private mListCount As Long
Private mListIndex As Long
Private mListDropped As Boolean
Private mParentFormHwnd As Long
Private mOldPopupOwnerHwnd As Long
Private mIndexBefore As Long
Private mMaxPopupItems As Long
Private mPopupListHeight As Long
Private mBackColor As Long
Private mForeColor As Long
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

Private mPopupWindowRgn As Long
Private mPopupLayered As Long
Private mEIV As Boolean
Private mMR As Boolean
Private mVScroll1Visible As Boolean
Private mLoadingFont As Boolean

Private Const cBackColor_Default As Long = 15853257

Private Sub btnDropDown_Click()
    If btnDropDown.Value Then
        If Not mListDropped Then
            RaiseEvent DropDown
            ShowList
        End If
    Else
        If mListDropped Then
            HideList
        End If
    End If
End Sub


Private Sub lblItem_Click(Index As Integer)
    ListIndex = Index
    BuildPopupList
    RaiseEvent ItemClick
    btnDropDown.Value = False
End Sub

Private Sub lblSelectedItem_Click()
    RaiseEvent ItemClick
    RaiseEvent Click
    btnDropDown.Value = False
End Sub

Private Sub lblText_Click()
    btnDropDown.Value = Not btnDropDown.Value
End Sub

Private Sub mFont_FontChanged(ByVal PropertyName As String)
    SetFont
End Sub

Private Sub picItem_Click(Index As Integer)
    ListIndex = Index
    BuildPopupList
    RaiseEvent ItemClick
    btnDropDown.Value = False
End Sub

Private Sub picPopupList_Resize()
    Dim iLng As Long
    
    iLng = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    VScroll1.Move picPopupList.ScaleWidth - iLng, 0, iLng, picPopupList.ScaleHeight
    VScroll1.ZOrder
End Sub

Private Sub picSelectedItem_Click()
    RaiseEvent ItemClick
    RaiseEvent Click
    btnDropDown.Value = False
End Sub

Private Sub picText_Click()
    btnDropDown.Value = Not btnDropDown.Value
End Sub

Private Sub tmrMouseOverCheck_Timer()
    Dim iHwnd As Long
    Dim iM As POINTAPI
    Dim iCtl As Object
    Static sLast As Long
    Static sMot As Long
    
    GetCursorPos iM
    
    iHwnd = WindowFromPoint(iM.x, iM.y)
    
    Set iCtl = GetControlByHwnd(iHwnd)
    If Not iCtl Is Nothing Then
'        Text = iCtl.Name
        If iCtl.Name = "picItem" Then
            If mPopupLayered Then
                If Not tmrTransparency.Enabled Then
                    sMot = sMot + 1
                Else
                    If tmrTransparency.Interval = 5000 Then
                        tmrTransparency.Enabled = False
                        tmrTransparency.Enabled = True
                    End If
                End If
                If sMot >= 17 Then
                    MakeFullyVisible
                    tmrTransparency.Enabled = False
                    tmrTransparency.Interval = 5000
                    tmrTransparency.Enabled = True
                    sMot = 0
                End If
            End If
            If (iCtl.Index + 1) <> sLast Then
                If sLast <> 0 Then
                    Set picItem(sLast - 1).Picture = picNormal.Picture
                    lblItem(sLast - 1).FontBold = False
                    sLast = 0
                End If
                Set iCtl.Picture = picMouseOver.Picture
                sLast = iCtl.Index + 1
                lblItem(sLast - 1).FontBold = True
            End If
        Else
            If sLast <> 0 Then
                Set picItem(sLast - 1).Picture = picNormal.Picture
                lblItem(sLast - 1).FontBold = False
                sLast = 0
                sMot = 0
            End If
        End If
    Else
        If sLast <> 0 Then
            Set picItem(sLast - 1).Picture = picNormal.Picture
            lblItem(sLast - 1).FontBold = False
            sLast = 0
            sMot = 0
        End If
    End If
End Sub

Private Sub tmrTransparency_Timer()
    Dim iM As POINTAPI
    Dim iRect As RECT
    
    GetWindowRect picPopupList.hWnd, iRect
    GetCursorPos iM
    If (iM.x >= iRect.Left) Then
        If (iM.x <= iRect.Right) Then
            If (iM.y >= iRect.Top) Then
                If (iM.y <= iRect.Bottom) Then
                    tmrTransparency.Interval = 5000
                    tmrTransparency.Enabled = False
                    tmrTransparency.Enabled = True
                    Exit Sub
                End If
            End If
        End If
    End If
    SetTransparency 160
    tmrTransparency.Enabled = False
End Sub

Private Sub UserControl_Hide()
    HideList False
End Sub

Private Sub UserControl_Initialize()
    ReDim mItems(-1 To -1)
    ReDim mItemData(-1 To -1)
    mListIndex = -1
End Sub

Private Sub UserControl_InitProperties()
    Text = Ambient.DisplayName
    mMaxPopupItems = 10
    mBackColor = cBackColor_Default
    mForeColor = vbWindowText
    mLoadingFont = True
    Set mFont = New StdFont
    mFont.Name = "Arial"
    mFont.Size = 10
    mLoadingFont = False
    SetFont
    PlaceBackPictures
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mText = PropBag.ReadProperty("Text", "")
    mMaxPopupItems = PropBag.ReadProperty("MaxPopupItems", 10)
    mBackColor = PropBag.ReadProperty("BackColor", cBackColor_Default)
    mForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
    mLoadingFont = True
    Set mFont = PropBag.ReadProperty("Font", Nothing)
    If mFont Is Nothing Then
        Set mFont = New StdFont
        mFont.Name = "Arial"
        mFont.Size = 10
    End If
    mLoadingFont = False
    If Ambient.UserMode Then
        mParentFormHwnd = GetParentFormHwnd(UserControl.Parent.hWnd)
    End If
    lblText.Caption = Text
    SetFont
    PlaceBackPictures
End Sub

Private Sub UserControl_Resize()
    Dim iH As Long
    Dim iW As Long
    
    iH = UserControl.ScaleHeight
    iW = UserControl.ScaleWidth
    
    If (iH <> picText.Height) Or (iW <> picText.Width) Then
        If (iH <> picText.Height) Then
            iH = picText.Height
        End If
        If (iW <> picText.Width) Then
            iW = picText.Width
        End If
        UserControl.Size iW, iH
    End If
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property


Public Property Let Text(nValue As String)
    If nValue <> mText Then
        mText = nValue
        PropertyChanged "Text"
        lblText.Caption = Text
    End If
End Property

Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "200"
    Text = mText
End Property

Private Sub UserControl_Show()
    If Not FontExists("Wingdings 3") Then
        If FontExists("Marlett") Then
            btnDropDown.FontName = "Marlett"
            btnDropDown.FontSize = 16
            btnDropDown.Caption = "6"
        ElseIf FontExists("Wingdings") Then
            btnDropDown.FontName = "Wingdings"
            btnDropDown.FontSize = 14
            btnDropDown.Caption = Chr(242)
        End If
    End If
End Sub

Private Sub UserControl_Terminate()
    HideList False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Text", mText, ""
    PropBag.WriteProperty "MaxPopupItems", mMaxPopupItems, 10
    PropBag.WriteProperty "BackColor", mBackColor, cBackColor_Default
    PropBag.WriteProperty "ForeColor", mForeColor, vbWindowText
    PropBag.WriteProperty "Font", mFont, Nothing
End Sub

Public Sub AddItem(nItem As String)
    mNewIndex = UBound(mItems) + 1
    If mNewIndex = 0 Then
        ReDim mItems(mNewIndex)
        ReDim mItemData(mNewIndex)
    Else
        ReDim Preserve mItems(mNewIndex)
        ReDim Preserve mItemData(mNewIndex)
    End If
    mItems(mNewIndex) = nItem
    mListCount = UBound(mItems) + 1
End Sub


Public Property Get List(ByVal Index As Integer) As String
    List = mItems(Index)
End Property

Public Property Let List(ByVal Index As Integer, ByVal nValue As String)
    mItems(Index) = nValue
    PropertyChanged "List"
End Property


Public Property Get NewIndex() As Long
    NewIndex = mNewIndex
End Property


Public Property Get ItemData(ByVal Index As Integer) As String
    ItemData = mItemData(Index)
End Property

Public Property Let ItemData(ByVal Index As Integer, ByVal nValue As String)
    mItemData(Index) = nValue
    PropertyChanged "ItemData"
End Property

Public Sub LoadFromSerializedString(nValue As String)
    Dim iStrs() As String
    Dim iStrs2() As String
    Dim c As Long
    
    iStrs = Split(nValue, "||")
    mListCount = UBound(iStrs)
    ReDim mItems(mListCount)
    ReDim mItemData(mListCount)
    For c = 0 To UBound(iStrs)
        iStrs2 = Split(iStrs(c), "|")
        If UBound(iStrs2) <> 1 Then
            ReDim mItems(-1 To -1)
            ReDim mItemData(-1, -1)
            mListCount = 0
            Exit Sub
        Else
            mItems(c) = iStrs2(0)
            mItemData(c) = iStrs2(1)
        End If
    Next c
End Sub


Public Property Let ListIndex(nValue As Long)
    
    If nValue > (mListCount - 1) Then Exit Property
    If nValue <> mListIndex Then
        mIndexBefore = mListIndex
        mListIndex = nValue
        If mListDropped Then
            UpdateList
        End If
        If mListIndex > -1 Then
            Text = mItems(mListIndex)
        Else
            Text = ""
        End If
        RaiseEvent Click
    End If
End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_MemberFlags = "400"
    ListIndex = mListIndex
End Property

Public Sub DropList()
    btnDropDown.Value = True
End Sub

Private Sub ShowList()
    Dim iWindowStyle As Long
    Dim iLng As Long
    Dim iPCT As Long
    
    If mParentFormHwnd = 0 Then
        mParentFormHwnd = GetParentFormHwnd(UserControl.Parent.hWnd)
        If mParentFormHwnd = 0 Then
            Parent.hWnd
        End If
    End If
    If mParentFormHwnd = 0 Then Exit Sub
    
    BuildPopupList
    
    iLng = VisibleItemsCount * picItem(0).Height
    mPopupListHeight = iLng
    picContainer.Height = picItem(0).Height * mListCount
    SetScroll
    
    mListDropped = True
    tmrMouseOverCheck.Enabled = True
    If mListIndex > 0 Then
        iPCT = picContainer.Top
        If iPCT > 0 Then
            EnsureItemVisible mListIndex
        End If
    End If
    PositionPopupList
    MakeRegion

    iWindowStyle = GetWindowLong(picPopupList.hWnd, GWL_EXSTYLE)
    iWindowStyle = iWindowStyle Or WS_EX_TOOLWINDOW
    SetWindowLong picPopupList.hWnd, GWL_EXSTYLE, iWindowStyle
    SetParent picPopupList.hWnd, 0
    mOldPopupOwnerHwnd = SetOwner(picPopupList.hWnd, mParentFormHwnd)

    AttachMessage Me, mParentFormHwnd, WM_WINDOWPOSCHANGED
'    AttachMessage Me, mParentFormHwnd, WM_SIZE
    AttachMessage Me, UserControl.hWnd, WM_MOVE

    If mListIndex > 0 Then
        If iPCT = 0 Then
            EnsureItemVisible mListIndex
            picPopupList.Visible = True
            picPopupList.Refresh
        End If
    End If
    
    picPopupList.Visible = True
    VScroll1.Visible = mVScroll1Visible
    picPopupList.ZOrder
    If IsWindows2000OrMore Then
        mPopupLayered = True
        iWindowStyle = GetWindowLong(picPopupList.hWnd, GWL_EXSTYLE)
        If (iWindowStyle And WS_EX_LAYERED) <> WS_EX_LAYERED Then
            SetWindowLong picPopupList.hWnd, GWL_EXSTYLE, iWindowStyle Or WS_EX_LAYERED
        End If
        MakeFullyVisible
    End If
    
    picContainer.Width = picPopupList.ScaleWidth - picContainer.Left - Screen.TwipsPerPixelX - 15
    picPopupList.Line (picPopupList.ScaleWidth - Screen.TwipsPerPixelX, 0)-(picPopupList.ScaleWidth - Screen.TwipsPerPixelX, picPopupList.ScaleHeight), &HC0C0C0
    
End Sub

Public Sub HideList(Optional nRaiseEvent As Boolean = True)
    If mListDropped Then
        picPopupList.Visible = False
        If mPopupWindowRgn <> 0 Then DeleteObject mPopupWindowRgn
        SetOwner picPopupList.hWnd, mOldPopupOwnerHwnd
        SetParent picPopupList, UserControl.hWnd
        DetachMessage Me, mParentFormHwnd, WM_WINDOWPOSCHANGED
'        DetachMessage Me, mParentFormHwnd, WM_SIZE
        DetachMessage Me, UserControl.hWnd, WM_MOVE
        tmrMouseOverCheck.Enabled = False
        tmrTransparency.Enabled = False
        mListDropped = False
        If nRaiseEvent Then
            RaiseEvent ListHided
        End If
    End If
    btnDropDown.Value = False
    btnDropDown.Refresh
End Sub

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrPostProcess
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Select Case iMsg
        Case WM_MOVE, WM_WINDOWPOSCHANGED
            If mListDropped Then
                PositionPopupList
                picPopupList.Visible = True
            End If
        Case Else
            '
    End Select
End Function

Private Sub PositionPopupList()
    Dim iRectControl As RECT
    Dim iRectPopup As RECT
    Dim iWidthPopup As Long
    Dim iHeightPopup As Long
    
    If IsIconic(mParentFormHwnd) <> 0 Then
        picPopupList.Visible = False
    Else
        GetWindowRect UserControl.hWnd, iRectControl
        GetWindowRect picPopupList.hWnd, iRectPopup
        iWidthPopup = UserControl.Width + picItem(0).Left + 4 * Screen.TwipsPerPixelX   ' (iRectPopup.Right - iRectPopup.Left)
        iHeightPopup = mPopupListHeight ' (iRectPopup.Bottom - iRectPopup.Top)
        MoveWindow picPopupList.hWnd, iRectControl.Right - iWidthPopup / Screen.TwipsPerPixelX, iRectControl.Bottom + 1, iWidthPopup / Screen.TwipsPerPixelX, iHeightPopup / Screen.TwipsPerPixelY, 1
    End If
End Sub

Private Sub BuildPopupList()
    Dim c As Long
    Dim iItemHeight As Long
    
    If picPopupList.Visible Then SetWindowRedraw picPopupList.hWnd, False
    For c = picItem.UBound + 1 To UBound(mItems)
        Load picItem(c)
        Load lblItem(c)
        Set lblItem(c).Container = picItem(c)
        lblItem(c).Move lblItem(0).Left, lblItem(0).Top
        lblItem(c).Visible = True
        Load shpIem(c)
        Set shpIem(c).Container = picItem(c)
        shpIem(c).Move shpIem(0).Left, shpIem(0).Top
'        picItem(c).Width =picPopupList.ScaleWidth -
'        If c = mListIndex Then
'            picSelectedItem.Top = picItem(mListIndex).Top
'            lblSelectedItem.Caption = lblItem(mListIndex).Caption
'            picSelectedItem.ZOrder
'            If picPopupList.Visible Then SetWindowRedraw picPopupList.hWnd, True
'            If picPopupList.Visible Then picPopupList.Refresh
'        End If
    Next c
    
    For c = (UBound(mItems) + 1) To picItem.UBound
        If c > 0 Then
            Unload lblItem(c)
            Unload shpIem(c)
            Unload picItem(c)
        End If
    Next c
    
    iItemHeight = picItem(0).Height
    For c = 0 To UBound(mItems)
        picItem(c).Top = iItemHeight * c
        lblItem(c).Caption = mItems(c)
        Set picItem(c).Picture = picNormal.Picture
        picItem(c).Visible = (c <> mListIndex)
        If c = mListIndex Then
            picSelectedItem.Top = picItem(mListIndex).Top
            lblSelectedItem.Caption = lblItem(mListIndex).Caption
            picSelectedItem.ZOrder
            If picPopupList.Visible Then SetWindowRedraw picPopupList.hWnd, True
            If picPopupList.Visible Then picPopupList.Refresh
        End If
    Next c
    
    If mListIndex > -1 Then
        picSelectedItem.Top = picItem(mListIndex).Top
        lblSelectedItem.Caption = lblItem(mListIndex).Caption
        picSelectedItem.ZOrder
        picSelectedItem.Visible = True
    Else
        picSelectedItem.Visible = False
    End If
    mListCount = UBound(mItems) + 1
    If picPopupList.Visible Then SetWindowRedraw picPopupList.hWnd, True
    If picPopupList.Visible Then picPopupList.Refresh
'    UserControl.Parent.Refresh
End Sub

Private Function GetControlByHwnd(nHwnd As Long) As Object
    Dim iCtl As Control
    Dim iHwnd As Long
    
    On Error Resume Next
    If nHwnd <> 0 Then
        For Each iCtl In UserControl.Controls
            iHwnd = iCtl.hWnd
            If iHwnd = nHwnd Then
                Set GetControlByHwnd = iCtl
                Exit Function
            End If
        Next
    End If
End Function

Private Sub UpdateList()
    If mIndexBefore > -1 Then
        picItem(mIndexBefore).Visible = True
    End If
    
    If mListIndex > -1 Then
        picSelectedItem.Top = picItem(mListIndex).Top
        lblSelectedItem.Caption = lblItem(mListIndex).Caption
        picSelectedItem.Visible = True
        EnsureItemVisible mListIndex
    Else
        picSelectedItem.Visible = False
    End If
    MakeRegion
End Sub

Private Sub SetScroll()
    If mListCount <= mMaxPopupItems Then
        VScroll1.Visible = False
        mVScroll1Visible = False
'        picContainer.Width = picPopupList.ScaleWidth - picContainer.Left - Screen.TwipsPerPixelX - 15
'        picPopupList.Line (picPopupList.ScaleWidth - Screen.TwipsPerPixelX, 0)-(picPopupList.ScaleWidth - Screen.TwipsPerPixelX, picPopupList.ScaleHeight), &HC0C0C0
    Else
        VScroll1.Max = mListCount - mMaxPopupItems
        VScroll1.Min = 0
        VScroll1.Value = 0
        mVScroll1Visible = True
        picContainer.Width = picPopupList.ScaleWidth - VScroll1.Width
        VScroll1.LargeChange = 3
    End If
End Sub

Private Sub VScroll1_Change()
     Dim iMR As Boolean
    
    If mListIndex > -1 Then
        If IsItemVisible(mListIndex) Then
            iMR = True
        End If
    End If
    picContainer.Top = VScroll1.Value * -1 * picItem(0).Height
    If mListIndex > -1 Then
'        If Not iMR Then
'            If IsItemVisible(mListIndex) Then
                iMR = True
'            End If
'        End If
    End If
    If iMR Then MakeRegion
End Sub

Private Sub VScroll1_GotFocus()
    MakeFullyVisible
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
    MakeFullyVisible
End Sub

Private Function IsItemVisible(nItem As Long) As Boolean
     IsItemVisible = (picItem(nItem).Top > picContainer.Top * -1) And ((picItem(nItem).Top + picContainer.Top) < picPopupList.ScaleHeight)
End Function

Private Sub EnsureItemVisible(nItem As Long)
    mEIV = True
    If (picItem(nItem).Top + picContainer.Top) < 0 Then
        Do Until ((picItem(nItem).Top + picContainer.Top) >= 0)
            VScroll1.Value = VScroll1.Value - 1
        Loop
        If VScroll1.Value < 4 Then
            VScroll1.Value = 0
        Else
            If VScroll1.Value > 0 Then VScroll1.Value = VScroll1.Value - 1
        End If
    ElseIf ((picItem(nItem).Top + picContainer.Top) >= picPopupList.ScaleHeight) Then
        Do Until ((picItem(nItem).Top + picContainer.Top) < picPopupList.ScaleHeight)
            VScroll1.Value = VScroll1.Value + 1
        Loop
        If (VScroll1.Max - VScroll1.Value) < 4 Then
            VScroll1.Value = VScroll1.Max
        Else
            If VScroll1.Value < VScroll1.Max Then VScroll1.Value = VScroll1.Value + 1
        End If
    End If
    mEIV = False
    If mMR Then
        MakeRegion
    End If
End Sub

Private Function RegionFromPic(nPic As PictureBox) As Long
    Dim iTmpRgn   As Long
    Dim iResultRgn  As Long
    Dim iStart    As Long
    Dim iY      As Long
    Dim iX      As Long
    Dim iHeight As Long
    Dim iWidth As Long
    Dim iTransparecyColor As Long
    
    iTransparecyColor = vbGreen
    ' Create a rectangular region.
    ' A region is a rectangle, polygon, or ellipse (or a combination
    ' of two or more of these shapes) that can be filled, painted,
    ' inverted, framed, and used to perform hit testing (testing
    ' for the cursor location).
    '
    iResultRgn = CreateRectRgn(0, 0, 0, 0)

    ' Get the dimensions of the bitmap.
    iHeight = UserControl.ScaleX(nPic.Height, vbTwips, vbPixels)
    iWidth = UserControl.ScaleY(nPic.Width, vbTwips, vbPixels)

    '
    ' Loop through the bitmap, row by row, examining each pixel.
    ' In each row, work from left to right comparing each pixel
    ' to the transparency color.
    '
    For iY = 0 To iHeight - 1
        iX = 0
        Do While iX < iWidth
            '
            ' Skip all pixels in a row with the same
            ' color as the transparency color.
            '
            Do While iX < iWidth And (GetPixel(nPic.hDC, iX, iY) <> iTransparecyColor)
                iX = iX + 1
            Loop

            If iX < iWidth Then
                '
                ' Get the start and end of the block of pixels in the
                ' row that are not the same color as the transparency.
                '
                iStart = iX
                Do While iX < iWidth And GetPixel(nPic.hDC, iX, iY) = iTransparecyColor
                    iX = iX + 1
                Loop
                If iX > iWidth Then iX = iWidth
                '
                ' Create a region equal in size to the line of pixels
                ' that don't match the transparency color. Combine this
                ' region with our final region.
                '
                iTmpRgn = CreateRectRgn(iStart, iY, iX, iY + 1)
                Call CombineRgn(iResultRgn, iResultRgn, iTmpRgn, RGN_OR)
                Call DeleteObject(iTmpRgn)
            End If
        Loop
    Next

    RegionFromPic = iResultRgn
End Function

Private Sub MakeRegion()
    Dim iRgn As Long
    Dim c As Long
    Dim iItemHeight As Long
    Dim iAuxLng As Long
    Dim iPicItemLeft As Long
    Dim iBb As Boolean
    Dim iRgnBox As Long
    Dim iRgnSelectedItem As Long
    
    If mEIV Then
        mMR = True
        Exit Sub
    End If
    iItemHeight = UserControl.ScaleY(picItem(0).Height, vbTwips, vbPixels)
    iPicItemLeft = UserControl.ScaleX(picItem(0).Left, vbTwips, vbPixels)
    iRgnBox = CreateRectRgn(UserControl.ScaleX(picItem(0).Left, vbTwips, vbPixels), 0, UserControl.ScaleX(picPopupList.Width, vbTwips, vbPixels), UserControl.ScaleY(picPopupList.Height, vbTwips, vbPixels))
    
    If mListIndex > -1 Then
        iAuxLng = mListIndex - VScroll1.Value
    Else
        iAuxLng = -1
    End If
    iRgn = RegionFromPic(picRegion)
    OffsetRgn iRgn, iPicItemLeft, 0
    For c = 0 To VisibleItemsCount - 1
        If c > 0 Then
            OffsetRgn iRgn, 0, iItemHeight
        End If
        iBb = False
        If c = iAuxLng Then ' selected item
            iRgnSelectedItem = CreateRectRgn(0, iItemHeight * c, iPicItemLeft, iItemHeight * (c + 1))
            CombineRgn iRgnBox, iRgnBox, iRgnSelectedItem, RGN_OR
            DeleteObject iRgnSelectedItem
            OffsetRgn iRgn, iPicItemLeft * -1, 0
            iBb = True
        End If
        CombineRgn iRgnBox, iRgnBox, iRgn, RGN_DIFF
        If iBb Then
            OffsetRgn iRgn, iPicItemLeft, 0
        End If
    Next c
    
    DeleteObject iRgn
    
    SetWindowRgn picPopupList.hWnd, iRgnBox, True
    If mPopupWindowRgn <> 0 Then DeleteObject mPopupWindowRgn
    mPopupWindowRgn = iRgnBox
    If picPopupList.Visible Then picPopupList.Refresh
'    UserControl.Parent.Refresh
End Sub


Private Function VisibleItemsCount() As Long
    If mListCount > mMaxPopupItems Then
        VisibleItemsCount = mMaxPopupItems
    Else
        VisibleItemsCount = mListCount
    End If
End Function

Private Sub SetTransparency(nValue As Long)
    Static sLast As Long
    
    If Not mPopupLayered Then Exit Sub
    If nValue <> sLast Then
        SetLayeredWindowAttributes picPopupList.hWnd, 0, nValue, LWA_ALPHA
        sLast = nValue
    End If
End Sub

Private Sub MakeFullyVisible()
    If Not mPopupLayered Then Exit Sub
    SetTransparency 255
    tmrTransparency.Interval = 20000
    tmrTransparency.Enabled = False
    tmrTransparency.Enabled = True
End Sub


Public Property Let MaxPopupItems(nValue As Long)
    If nValue <> mMaxPopupItems Then
        mMaxPopupItems = nValue
        If mListDropped Then
            HideList
            ShowList
        End If
    End If
End Property

Public Property Get MaxPopupItems() As Long
    MaxPopupItems = mMaxPopupItems
End Property

Public Sub Clear()
    HideList
    
    Text = ""
    ReDim mItems(-1 To -1)
    ReDim mItemData(-1 To -1)
    mListIndex = -1
    mNewIndex = -1
    mListCount = 0
End Sub

Public Property Get ListCount() As Long
    ListCount = mListCount
End Property

Private Function SetOwner(ByVal HwndWindow, ByVal HwndofOwner) As Long
    On Error Resume Next
    SetOwner = SetWindowLong(HwndWindow, GWL_HWNDPARENT, HwndofOwner)
End Function

Private Sub PlaceBackPictures()
    Dim iSng As Single
    Dim iHeight As Long
    Dim iBackColor As Long
    
    Dim R1 As Long
    Dim G1 As Long
    Dim B1 As Long
    Dim H1 As Long
    Dim L1 As Long
    Dim S1 As Long
    
    Dim R2 As Long
    Dim G2 As Long
    Dim B2 As Long
    Dim H2 As Long
    Dim L2 As Long
    Dim S2 As Long
    
    Dim H3 As Long
    Dim L3 As Long
    Dim S3 As Long
    
    TranslateColor mBackColor, 0, iBackColor
    
    R1 = cBackColor_Default And 255 ' R
    G1 = (cBackColor_Default \ 256) And 255 ' G
    B1 = (cBackColor_Default \ 65536) And 255 ' B
    
    ColorRGBToHLS RGB(R1, G1, B1), H1, L1, S1
    
    R2 = iBackColor And 255 ' R
    G2 = (iBackColor \ 256) And 255 ' G
    B2 = (iBackColor \ 65536) And 255 ' B
    
    ColorRGBToHLS RGB(R2, G2, B2), H2, L2, S2
    
    H3 = H2 - H1
    
    If H3 > 120 Then
        H3 = H3 - 240
    End If
    If H3 < -120 Then
        H3 = H3 + 240
    End If
    
    L3 = L2 - L1
    S3 = S2 - S1
   
    ' some limits:
    If L3 > 50 Then L3 = 50
    If S3 > 50 Then S3 = 50
    
    Set picBackground.Picture = AdjustPictureWithHLS(picBackground_Default.Picture, H3, L3, S3)
    Set picBackSelectedItem.Picture = AdjustPictureWithHLS(picBackSelectedItem_Default.Picture, H3, L3, S3)
    btnDropDown.BackColor = AdjustColorWithHLS(&HE3D8BD, H3, L3, S3)
    shpSelectedItem.BackColor = AdjustColorWithHLS(&HD9BB45, H3, L3, S3)
    
    picItem(0).Height = Int(picItem(0).Height / Screen.TwipsPerPixelY) * Screen.TwipsPerPixelY
    picItem(0).Width = picPopupList.ScaleWidth - picItem(0).Left
    picSelectedItem.Width = picPopupList.ScaleWidth - picSelectedItem.Left
    
    picBackground.Width = picText.Width
    picBackground.Height = picText.Height
    picBackground.PaintPicture picBackground.Picture, 0, 0, picText.Width, picText.Height
    Set picText.Picture = picBackground.Image
    picBackground.Cls
    
    iSng = picItem(0).Height / picAux_picNormal.Picture.Height
    iHeight = picAux_picNormal.Picture.Height * iSng
    
    picBackSelectedItem.Width = picSelectedItem.Width
    picBackSelectedItem.Height = iHeight
    picBackSelectedItem.PaintPicture picBackSelectedItem.Picture, 0, 0, , iHeight
    picBackSelectedItem.PaintPicture picBackSelectedItem.Picture, UserControl.ScaleX(15, vbPixels, vbTwips), 0, picItem(0).Width - UserControl.ScaleX(15, vbPixels, vbTwips), iHeight, UserControl.ScaleX(15, vbPixels, vbTwips)
    Set picSelectedItem.Picture = picBackSelectedItem.Image
    picBackSelectedItem.Cls
    
    picNormal.Width = picItem(0).Width
    picMouseOver.Width = picItem(0).Width
    
    picAux_picNormal.Width = picItem(0).Width
    picAux_picNormal.Height = iHeight
    picAux_picNormal.PaintPicture picAux_picNormal.Picture, 0, 0, , iHeight
    picAux_picNormal.PaintPicture picAux_picNormal.Picture, UserControl.ScaleX(15, vbPixels, vbTwips), 0, picItem(0).Width - UserControl.ScaleX(15, vbPixels, vbTwips), iHeight, UserControl.ScaleX(15, vbPixels, vbTwips)
    Set picNormal.Picture = picAux_picNormal.Image
    picAux_picNormal.Cls
    
    picAux_picMouseOver.Width = picItem(0).Width
    picAux_picMouseOver.Height = iHeight
    picAux_picMouseOver.PaintPicture picAux_picMouseOver.Picture, 0, 0, , iHeight
    picAux_picMouseOver.PaintPicture picAux_picMouseOver.Picture, UserControl.ScaleX(15, vbPixels, vbTwips), 0, picItem(0).Width - UserControl.ScaleX(15, vbPixels, vbTwips), iHeight, UserControl.ScaleX(15, vbPixels, vbTwips)
    Set picMouseOver.Picture = picAux_picMouseOver.Image
    picAux_picMouseOver.Cls
    
'    picAux_picRegion.Width = picRegion.Width
 '   picAux_picRegion.Height = picRegion.Height
  '  picAux_picRegion.PaintPicture picAux_picRegion.Picture, 0, 0, picRegion.Width, picRegion.Height
   ' Set picRegion.Picture = picAux_picRegion.Image
    'picAux_picRegion.Cls
    
    picRegion.PaintPicture picNormal.Picture, 0, 0, , , , , UserControl.ScaleX(15, vbPixels, vbTwips)
    
    lblText.ForeColor = mForeColor
    lblSelectedItem.ForeColor = mForeColor
    btnDropDown.ForeColor = mForeColor
End Sub

Public Property Let BackColor(nColor As OLE_COLOR)
    If nColor <> mBackColor Then
        mBackColor = nColor
        PropertyChanged "BackColor"
        PlaceBackPictures
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property


Public Property Let ForeColor(nColor As OLE_COLOR)
    If nColor <> mForeColor Then
        mForeColor = nColor
        PropertyChanged "ForeColor"
        lblText.ForeColor = mForeColor
        lblSelectedItem.ForeColor = mForeColor
        btnDropDown.ForeColor = mForeColor
    End If
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property


Public Property Get Font() As StdFont
    Set Font = mFont
End Property

Public Property Set Font(ByVal nFont As StdFont)
    If Not mFont Is nFont Then
        Set mFont = nFont
        SetFont
        PropertyChanged "Font"
    End If
End Property

Public Property Let Font(ByVal nFont As StdFont)
    Set Font = nFont
End Property

Private Sub SetFont()
    Dim iNormal As StdFont
    Dim iBold As StdFont
    Dim c As Long
    
    If mFont Is Nothing Then Exit Sub
    If mLoadingFont Then Exit Sub
    
    Set iNormal = CloneFont(mFont)
    Set iBold = CloneFont(mFont)
    
    iNormal.Bold = False
    iBold.Bold = True
    
    For c = lblItem.LBound To lblItem.UBound
        Set lblItem(c).Font = iNormal
    Next
    Set lblText.Font = iBold
    Set lblSelectedItem.Font = iBold
End Sub
