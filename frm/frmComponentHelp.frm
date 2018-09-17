VERSION 5.00
Begin VB.Form frmComponentHelp 
   Caption         =   "Help"
   ClientHeight    =   5184
   ClientLeft      =   6600
   ClientTop       =   4752
   ClientWidth     =   6768
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5184
   ScaleWidth      =   6768
   Begin vbExtra.ctlBuildHelp BuildHelp1 
      Left            =   504
      Top             =   4104
      _ExtentX        =   3048
      _ExtentY        =   1185
   End
   Begin vbExtra.SizeGrip SizeGrip1 
      Height          =   228
      Left            =   6540
      Top             =   4956
      Width           =   228
      _ExtentX        =   402
      _ExtentY        =   402
   End
   Begin vbExtra.SSTabEx sst1 
      Height          =   3180
      Left            =   108
      TabIndex        =   4
      Top             =   504
      Width           =   5664
      _ExtentX        =   9991
      _ExtentY        =   5609
      Tabs            =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   0   'False
      Style           =   1
      TabHeight       =   520
      TabSelExtraHeight=   71
      TabSelHighlight =   -1  'True
      TabSelFontBold  =   0
      TabBackColor    =   15987699
      TabWidthStyle   =   0
      TabAppearance   =   4
      TabCaption(0)   =   " ScrollableContainer control"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "txtTabText(0)"
      TabCaption(1)   =   " SSTabEx control"
      Tab(1).ControlCount=   1
      Tab(1).Control(0)=   "txtTabText(1)"
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   1
      Tab(2).Control(0)=   "txtTabText(2)"
      TabCaption(3)   =   "Tab 3"
      Tab(3).ControlCount=   1
      Tab(3).Control(0)=   "txtTabText(3)"
      Begin VB.TextBox txtTabText 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1692
         HideSelection   =   0   'False
         Index           =   3
         Left            =   -74676
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   696
         Width           =   2604
      End
      Begin VB.TextBox txtTabText 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1692
         HideSelection   =   0   'False
         Index           =   2
         Left            =   -74748
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   744
         Width           =   2604
      End
      Begin VB.TextBox txtTabText 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1692
         HideSelection   =   0   'False
         Index           =   1
         Left            =   -74676
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "frmComponentHelp.frx":0000
         Top             =   552
         Width           =   2604
      End
      Begin VB.TextBox txtTabText 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1692
         HideSelection   =   0   'False
         Index           =   0
         Left            =   144
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "frmComponentHelp.frx":5E7C
         Top             =   432
         Width           =   2604
      End
   End
   Begin vbExtra.ToolBarDA tbrActions 
      Height          =   396
      Left            =   4140
      Top             =   108
      Width           =   2532
      _ExtentX        =   4466
      _ExtentY        =   699
      ButtonsCount    =   6
      ButtonKey1      =   "DecreaseFont"
      ButtonPic161    =   "frmComponentHelp.frx":7118
      ButtonPic201    =   "frmComponentHelp.frx":746A
      ButtonPic241    =   "frmComponentHelp.frx":796C
      ButtonPic301    =   "frmComponentHelp.frx":807E
      ButtonPic361    =   "frmComponentHelp.frx":8B98
      ButtonToolTipText1=   "Decrease the font size"
      ButtonOrderToHide1=   5
      ButtonKey2      =   "IncreaseFont"
      ButtonPic162    =   "frmComponentHelp.frx":9B1A
      ButtonPic202    =   "frmComponentHelp.frx":9E6C
      ButtonPic242    =   "frmComponentHelp.frx":A36E
      ButtonPic302    =   "frmComponentHelp.frx":AA80
      ButtonPic362    =   "frmComponentHelp.frx":B59A
      ButtonToolTipText2=   "Increase the font size"
      ButtonOrderToHide2=   5
      ButtonKey3      =   "Print"
      ButtonPic163    =   "frmComponentHelp.frx":C51C
      ButtonPic203    =   "frmComponentHelp.frx":C86E
      ButtonPic243    =   "frmComponentHelp.frx":CD70
      ButtonPic303    =   "frmComponentHelp.frx":D482
      ButtonPic363    =   "frmComponentHelp.frx":DF9C
      ButtonToolTipText3=   "Print"
      ButtonOrderToHide3=   3
      ButtonKey4      =   "Copy"
      ButtonPic164    =   "frmComponentHelp.frx":EF1E
      ButtonPic204    =   "frmComponentHelp.frx":F270
      ButtonPic244    =   "frmComponentHelp.frx":F772
      ButtonPic304    =   "frmComponentHelp.frx":FE84
      ButtonPic364    =   "frmComponentHelp.frx":1099E
      ButtonToolTipText4=   "Copy"
      ButtonOrderToHide4=   2
      ButtonKey5      =   "Save"
      ButtonPic165    =   "frmComponentHelp.frx":11920
      ButtonPic205    =   "frmComponentHelp.frx":11C72
      ButtonPic245    =   "frmComponentHelp.frx":12174
      ButtonPic305    =   "frmComponentHelp.frx":12886
      ButtonPic365    =   "frmComponentHelp.frx":133A0
      ButtonToolTipText5=   "Save to a file"
      ButtonOrderToHide5=   1
      ButtonKey6      =   "Find"
      ButtonPic166    =   "frmComponentHelp.frx":14322
      ButtonPic206    =   "frmComponentHelp.frx":14674
      ButtonPic246    =   "frmComponentHelp.frx":14B76
      ButtonPic306    =   "frmComponentHelp.frx":15288
      ButtonPic366    =   "frmComponentHelp.frx":15DA2
      ButtonToolTipText6=   "Find text"
      ButtonOrderToHide6=   1
   End
   Begin vbExtra.PrintFn PrintFn1 
      Left            =   6048
      Top             =   600
      _ExtentX        =   720
      _ExtentY        =   720
      FormatButtonPicture_0=   "frmComponentHelp.frx":16D24
      FormatButtonPicture_1=   "frmComponentHelp.frx":17226
      FormatButtonPicture_2=   "frmComponentHelp.frx":17938
      FormatButtonPicture_3=   "frmComponentHelp.frx":18452
      FormatButtonPicture_4=   "frmComponentHelp.frx":193D4
      PageNumbersFormat=   ""
      PageNumbersFormatIndex=   3
      BeginProperty PageNumbersFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmComponentHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const cMinFontSize = 7
Private Const cMaxFontSize = 18
Private Const cDefaultFontSize = 11

Private mFontSize1 As Single

Private Function GetTextBoxOfTab(nTab As Integer) As Control
    Dim iCtl As Control
    
    For Each iCtl In sst1.TabControls(nTab)
        If TypeName(iCtl) = "TextBox" Then
            Set GetTextBoxOfTab = iCtl
            Exit Function
        End If
    Next
End Function

Private Sub Form_Load()
    Dim c As Long
    
    PersistForm Me, Forms
    SetMinMax Me, 3500, 3500
    Me.Caption = App.Title & " help"
    
    FontSize1 = Val(GetSetting(App.Title, "Design", "HelpFontSize", Trim$(Str$(cDefaultFontSize))))
    
    For c = 0 To sst1.Tabs - 1
        If Left(sst1.TabCaption(c), 4) = "Tab " Then
            sst1.TabVisible(c) = False
        End If
    Next
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    tbrActions.Move Me.ScaleWidth - tbrActions.Width - 480, 30
    tbrActions.ZOrder
    
    sst1.Left = 45
    sst1.Width = Me.ScaleWidth - 90
    If (sst1.Left + sst1.EndOfTabs) < tbrActions.Left Then
        sst1.Top = tbrActions.Top + tbrActions.Height + 45 - sst1.TabBodyTop
    Else
        sst1.Top = tbrActions.Top + tbrActions.Height + 45
    End If
    
    If Me.WindowState = vbMaximized Then
        sst1.Height = Me.ScaleHeight - sst1.Top - 30
    Else
        sst1.Height = Me.ScaleHeight - sst1.Top - 30 - SizeGrip1.Height
    End If
    
End Sub

Public Sub ShowItem(ByVal nItem As String)
    Dim c As Long
    
    nItem = LCase(nItem)
    For c = 0 To sst1.Tabs - 1
        If InStr(LCase(sst1.TabCaption(c)), nItem) > 0 Then
            sst1.TabSel = c
            Exit For
        End If
    Next
End Sub

Private Sub PrintFn1_PrepareDoc(ByVal DocKey As String)
    Dim iTb As TextBox
    
    Set iTb = GetTextBoxOfTab(sst1.TabSel)
    If Not iTb Is Nothing Then
        Printer2.Font.Name = iTb.Font.Name
        Printer2.Font.Size = iTb.Font.Size * 1.2
        Printer2.Font.Bold = True
        Printer2.Print sst1.Caption
        Printer2.Print
        Printer2.Font.Bold = False
        Printer2.Font.Size = iTb.Font.Size
        Printer2.Print iTb.Text
    End If
End Sub

Private Sub sst1_TabBodyResize()
    Dim t As Integer
    Dim iTextBox As Control
    
    For t = 0 To sst1.Tabs - 1
        If sst1.TabVisible(t) Then
            Set iTextBox = GetTextBoxOfTab(t)
            If Not iTextBox Is Nothing Then
                SetWindowRedraw iTextBox.hWnd, False
            End If
        End If
    Next t
    
    For t = 0 To sst1.Tabs - 1
        If sst1.TabVisible(t) Then
            Set iTextBox = GetTextBoxOfTab(t)
            If Not iTextBox Is Nothing Then
                iTextBox.Move sst1.TabBodyLeft, sst1.TabBodyTop, sst1.TabBodyWidth, sst1.TabBodyHeight
            End If
        End If
    Next t

    For t = 0 To sst1.Tabs - 1
        If sst1.TabVisible(t) Then
            Set iTextBox = GetTextBoxOfTab(t)
            If Not iTextBox Is Nothing Then
                SetWindowRedraw iTextBox.hWnd, True
            End If
        End If
    Next t

End Sub

Private Sub tbrActions_ButtonClick(Button As ToolBarDAButton)
    Dim iTextBox As TextBox

    If Button.Key = "DecreaseFont" Then
        DecreaseFontSize
    ElseIf Button.Key = "IncreaseFont" Then
        IncreaseFontSize
    ElseIf Button.Key = "Print" Then
        PrintFn1.ShowPrintPreview
    ElseIf Button.Key = "Copy" Then
        Dim iFrm As frmClipboardCopiedMessage
        
        Set iTextBox = GetTextBoxOfTab(sst1.TabSel)
        If Not iTextBox Is Nothing Then
            Clipboard.Clear
            Clipboard.SetText iTextBox.Text
            Set iFrm = New frmClipboardCopiedMessage
          '  iFrm.lblMessage.Caption = "Text copied"
            iFrm.ShowMessage
            If IsFormLoaded(iFrm) Then
                Unload iFrm
            End If
            Set iFrm = Nothing
        End If
    ElseIf Button.Key = "Save" Then
        Dim iDlg As New CommonDialogExObject
        
        Set iTextBox = GetTextBoxOfTab(sst1.TabSel)
        If Not iTextBox Is Nothing Then
            iDlg.FileName = sst1.Caption & ".txt"
            iDlg.Filter = "Text files (*.txt)|*.txt"
            iDlg.ShowSave
            If Not iDlg.CancelError Then
                On Error Resume Next
                SaveTextFile iDlg.FileName, iTextBox.Text
            End If
        End If
    ElseIf Button.Key = "Find" Then
        Static sSearch As String
        Dim iText As String
        Dim iStr As String
        Dim iPos As Long
        
        Set iTextBox = GetTextBoxOfTab(sst1.TabSel)
        If Not iTextBox Is Nothing Then
            iStr = InputBox("Enter the text to seach for:", "Find text", sSearch)
            If iStr = "" Then Exit Sub
            sSearch = LCase(iStr)
            iText = LCase(iTextBox.Text)
            iPos = InStr(iTextBox.SelStart + 2, iText, sSearch)
            If iPos = 0 Then
                iPos = InStr(1, iText, sSearch)
                If iPos = 0 Then
                    MsgBox "Text not found.", vbInformation
                Else
                    iTextBox.SelStart = iPos - 1
                    iTextBox.SelLength = Len(sSearch)
                End If
            Else
                iTextBox.SelStart = iPos - 1
                iTextBox.SelLength = Len(sSearch)
            End If
        End If
    End If
End Sub

Private Sub DecreaseFontSize()
    If FontSizeCanDecrease Then
        FontSize1 = FontSize1 - 1
    End If
End Sub

Private Sub IncreaseFontSize()
    If FontSizeCanIncrease Then
        FontSize1 = FontSize1 + 1
    End If
End Sub

Private Function FontSizeCanDecrease() As Boolean
    FontSizeCanDecrease = (FontSize1 - 1) >= cMinFontSize
End Function

Private Function FontSizeCanIncrease() As Boolean
    FontSizeCanIncrease = (FontSize1 + 1) <= cMaxFontSize
End Function

Private Property Get FontSize1() As Single
    FontSize1 = mFontSize1
End Property

Private Property Let FontSize1(nValor As Single)
    If nValor <> mFontSize1 Then
        mFontSize1 = nValor
        ChangeFontSize
    End If
End Property

Private Sub ChangeFontSize()
    Dim t As Integer
    Dim iTextBox As Control
    
    For t = 0 To sst1.Tabs - 1
        If sst1.TabVisible(t) Then
            Set iTextBox = GetTextBoxOfTab(t)
            If Not iTextBox Is Nothing Then
                iTextBox.Font.Size = mFontSize1
            End If
        End If
    Next t
        
    tbrActions.Buttons("DecreaseFont").Enabled = FontSizeCanDecrease
    tbrActions.Buttons("IncreaseFont").Enabled = FontSizeCanIncrease
    tbrActions.Buttons("DecreaseFont").ToolTipText = "Decrease the font size (current size is " & FontSize1 & ")"
    tbrActions.Buttons("IncreaseFont").ToolTipText = "Increase the font size (current size is " & FontSize1 & ")"
    
    SaveSetting App.Title, "Design", "HelpFontSize", Trim$(Str$(FontSize1))

End Sub
