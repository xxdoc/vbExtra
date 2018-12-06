VERSION 5.00
Begin VB.Form frmPrintPreview 
   BackColor       =   &H80000010&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "# Print preview"
   ClientHeight    =   6828
   ClientLeft      =   3600
   ClientTop       =   1248
   ClientWidth     =   9648
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrintPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6828
   ScaleWidth      =   9648
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picStatus 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   264
      Left            =   216
      ScaleHeight     =   264
      ScaleWidth      =   3648
      TabIndex        =   20
      Top             =   9900
      Width           =   3648
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   192
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   48
      End
   End
   Begin VB.Timer tmrCheckPopupMenu 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   504
      Top             =   3024
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default action"
      Default         =   -1  'True
      Height          =   405
      Left            =   17850
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   11940
      Width           =   945
   End
   Begin VB.Timer tmrPopupcboView 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   3024
   End
   Begin VB.PictureBox picMainContainer 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   13755
      Left            =   0
      ScaleHeight     =   13752
      ScaleWidth      =   16872
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   16872
      Begin VB.Timer tmrUnload 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1332
         Top             =   3492
      End
      Begin VB.Timer tmrIgnoreKeysUpDown 
         Enabled         =   0   'False
         Interval        =   400
         Left            =   936
         Top             =   3492
      End
      Begin VB.Timer tmrIgnoreMouseWheelEvents 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   540
         Top             =   3492
      End
      Begin VB.Timer tmrShowPage 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   108
         Top             =   3492
      End
      Begin vbExtra.ToolBarDA tbrTop 
         Height          =   516
         Left            =   0
         Top             =   648
         Width           =   16512
         _ExtentX        =   29125
         _ExtentY        =   910
         IconsSize       =   2
         LeaveBorderSpace=   -1  'True
         AutoSize        =   0   'False
         PicHeight16     =   20
         PicHeight20     =   24
         PicHeight24     =   30
         PicHeight30     =   36
         PicHeight36     =   44
         ButtonsCount    =   21
         ButtonWidth1    =   80
         ButtonStyle1    =   4
         ButtonKey2      =   "Print"
         ButtonPic162    =   "frmPrintPreview.frx":000C
         ButtonPic202    =   "frmPrintPreview.frx":050E
         ButtonPic242    =   "frmPrintPreview.frx":0C20
         ButtonPic302    =   "frmPrintPreview.frx":173A
         ButtonPic362    =   "frmPrintPreview.frx":26BC
         ButtonWidth2    =   512
         ButtonToolTipText2=   "# Print"
         ButtonKey3      =   "PageSetup"
         ButtonPic163    =   "frmPrintPreview.frx":3DBE
         ButtonPic203    =   "frmPrintPreview.frx":42C0
         ButtonPic243    =   "frmPrintPreview.frx":49D2
         ButtonPic303    =   "frmPrintPreview.frx":54EC
         ButtonPic363    =   "frmPrintPreview.frx":646E
         ButtonWidth3    =   512
         ButtonToolTipText3=   "# Page setup"
         ButtonKey4      =   "PageNumbers"
         ButtonPic164    =   "frmPrintPreview.frx":7B70
         ButtonPic204    =   "frmPrintPreview.frx":8072
         ButtonPic244    =   "frmPrintPreview.frx":8784
         ButtonPic304    =   "frmPrintPreview.frx":929E
         ButtonPic364    =   "frmPrintPreview.frx":A220
         ButtonToolTipText4=   "# Page numbers options"
         ButtonVisible4  =   0   'False
         ButtonKey5      =   "Format"
         ButtonPic165    =   "frmPrintPreview.frx":B922
         ButtonPic205    =   "frmPrintPreview.frx":BE24
         ButtonPic245    =   "frmPrintPreview.frx":C536
         ButtonPic305    =   "frmPrintPreview.frx":D050
         ButtonPic365    =   "frmPrintPreview.frx":DFD2
         ButtonWidth5    =   512
         ButtonToolTipText5=   "# Format"
         ButtonKey6      =   "OrientationLabelSpace"
         ButtonWidth6    =   1100
         ButtonStyle6    =   4
         ButtonChecked7  =   -1  'True
         ButtonKey7      =   "OrientationPortrait"
         ButtonPic167    =   "frmPrintPreview.frx":F6D4
         ButtonPic207    =   "frmPrintPreview.frx":FBD6
         ButtonPic247    =   "frmPrintPreview.frx":102E8
         ButtonPic307    =   "frmPrintPreview.frx":10E02
         ButtonPic367    =   "frmPrintPreview.frx":11D84
         ButtonWidth7    =   512
         ButtonStyle7    =   2
         ButtonToolTipText7=   "# Orientation portrait"
         ButtonKey8      =   "OrientationLandscape"
         ButtonPic168    =   "frmPrintPreview.frx":13486
         ButtonPic208    =   "frmPrintPreview.frx":13988
         ButtonPic248    =   "frmPrintPreview.frx":1409A
         ButtonPic308    =   "frmPrintPreview.frx":14BB4
         ButtonPic368    =   "frmPrintPreview.frx":15B36
         ButtonWidth8    =   512
         ButtonStyle8    =   2
         ButtonToolTipText8=   "# Orientation landscape"
         ButtonKey9      =   "ViewLabelSpace"
         ButtonWidth9    =   660
         ButtonStyle9    =   4
         ButtonKey10     =   "ViewNormalSize"
         ButtonTag10     =   "view"
         ButtonPic1610   =   "frmPrintPreview.frx":17238
         ButtonPic2010   =   "frmPrintPreview.frx":1773A
         ButtonPic2410   =   "frmPrintPreview.frx":17E4C
         ButtonPic3010   =   "frmPrintPreview.frx":18966
         ButtonPic3610   =   "frmPrintPreview.frx":198E8
         ButtonWidth10   =   512
         ButtonStyle10   =   2
         ButtonToolTipText10=   "# View normal page size"
         ButtonKey11     =   "ViewScreenWidth"
         ButtonTag11     =   "view"
         ButtonPic1611   =   "frmPrintPreview.frx":1AFEA
         ButtonPic2011   =   "frmPrintPreview.frx":1B4EC
         ButtonPic2411   =   "frmPrintPreview.frx":1BBFE
         ButtonPic3011   =   "frmPrintPreview.frx":1C718
         ButtonPic3611   =   "frmPrintPreview.frx":1D69A
         ButtonWidth11   =   512
         ButtonStyle11   =   2
         ButtonToolTipText11=   "# View page adjusted to the screen width"
         ButtonChecked12 =   -1  'True
         ButtonKey12     =   "ViewScreenHeight"
         ButtonTag12     =   "view"
         ButtonPic1612   =   "frmPrintPreview.frx":1ED9C
         ButtonPic2012   =   "frmPrintPreview.frx":1F29E
         ButtonPic2412   =   "frmPrintPreview.frx":1F9B0
         ButtonPic3012   =   "frmPrintPreview.frx":204CA
         ButtonPic3612   =   "frmPrintPreview.frx":2144C
         ButtonWidth12   =   512
         ButtonStyle12   =   2
         ButtonToolTipText12=   "# View page adjusted to the screen height"
         ButtonEnabled13 =   0   'False
         ButtonKey13     =   "ViewSeveralPages"
         ButtonTag13     =   "view"
         ButtonPic1613   =   "frmPrintPreview.frx":22B4E
         ButtonPic2013   =   "frmPrintPreview.frx":23050
         ButtonPic2413   =   "frmPrintPreview.frx":23762
         ButtonPic3013   =   "frmPrintPreview.frx":2427C
         ButtonPic3613   =   "frmPrintPreview.frx":251FE
         ButtonWidth13   =   512
         ButtonStyle13   =   2
         ButtonToolTipText13=   "# View several pages"
         ButtonKey14     =   "ScaleSpace"
         ButtonWidth14   =   2100
         ButtonStyle14   =   4
         ButtonKey15     =   "DecreaseScale"
         ButtonPic1615   =   "frmPrintPreview.frx":26900
         ButtonPic2015   =   "frmPrintPreview.frx":26E02
         ButtonPic2415   =   "frmPrintPreview.frx":27514
         ButtonPic3015   =   "frmPrintPreview.frx":2802E
         ButtonPic3615   =   "frmPrintPreview.frx":28FB0
         ButtonPic16Alt15=   "frmPrintPreview.frx":2A6B2
         ButtonPic20Alt15=   "frmPrintPreview.frx":2ABB4
         ButtonPic24Alt15=   "frmPrintPreview.frx":2B2C6
         ButtonPic30Alt15=   "frmPrintPreview.frx":2BDE0
         ButtonPic36Alt15=   "frmPrintPreview.frx":2CD62
         ButtonWidth15   =   512
         ButtonToolTipText15=   "# Decrease fonts and elements size"
         ButtonKey16     =   "IncreaseScale"
         ButtonPic1616   =   "frmPrintPreview.frx":2E464
         ButtonPic2016   =   "frmPrintPreview.frx":2E966
         ButtonPic2416   =   "frmPrintPreview.frx":2F078
         ButtonPic3016   =   "frmPrintPreview.frx":2FB92
         ButtonPic3616   =   "frmPrintPreview.frx":30B14
         ButtonPic16Alt16=   "frmPrintPreview.frx":32216
         ButtonPic20Alt16=   "frmPrintPreview.frx":32718
         ButtonPic24Alt16=   "frmPrintPreview.frx":32E2A
         ButtonPic30Alt16=   "frmPrintPreview.frx":33944
         ButtonPic36Alt16=   "frmPrintPreview.frx":348C6
         ButtonWidth16   =   512
         ButtonToolTipText16=   "# Increase fonts and elements size"
         ButtonStyle17   =   4
         ButtonKey18     =   "FirstPage"
         ButtonTag18     =   "GoToPage"
         ButtonPic1618   =   "frmPrintPreview.frx":35FC8
         ButtonPic2018   =   "frmPrintPreview.frx":364CA
         ButtonPic2418   =   "frmPrintPreview.frx":36BDC
         ButtonPic3018   =   "frmPrintPreview.frx":376F6
         ButtonPic3618   =   "frmPrintPreview.frx":38678
         ButtonVisible18 =   0   'False
         ButtonKey19     =   "PreviousPage"
         ButtonTag19     =   "GoToPage"
         ButtonPic1619   =   "frmPrintPreview.frx":39D7A
         ButtonPic2019   =   "frmPrintPreview.frx":3A27C
         ButtonPic2419   =   "frmPrintPreview.frx":3A98E
         ButtonPic3019   =   "frmPrintPreview.frx":3B4A8
         ButtonPic3619   =   "frmPrintPreview.frx":3C42A
         ButtonVisible19 =   0   'False
         ButtonKey20     =   "NextPage"
         ButtonTag20     =   "GoToPage"
         ButtonPic1620   =   "frmPrintPreview.frx":3DB2C
         ButtonPic2020   =   "frmPrintPreview.frx":3E02E
         ButtonPic2420   =   "frmPrintPreview.frx":3E740
         ButtonPic3020   =   "frmPrintPreview.frx":3F25A
         ButtonPic3620   =   "frmPrintPreview.frx":401DC
         ButtonVisible20 =   0   'False
         ButtonKey21     =   "LastPage"
         ButtonTag21     =   "GoToPage"
         ButtonPic1621   =   "frmPrintPreview.frx":418DE
         ButtonPic2021   =   "frmPrintPreview.frx":41DE0
         ButtonPic2421   =   "frmPrintPreview.frx":424F2
         ButtonPic3021   =   "frmPrintPreview.frx":4300C
         ButtonPic3621   =   "frmPrintPreview.frx":43F8E
         ButtonVisible21 =   0   'False
         Begin VB.Timer tmrcboScalePercentChange 
            Enabled         =   0   'False
            Interval        =   5000
            Left            =   9828
            Top             =   108
         End
         Begin VB.Timer tmrcboScalePercentUpdate 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   9540
            Top             =   108
         End
         Begin VB.CommandButton cmdClose 
            Cancel          =   -1  'True
            Caption         =   "# Close print preview"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   13680
            TabIndex        =   1
            Top             =   36
            Width           =   2232
         End
         Begin vbExtra.ButtonEx cmdClose_2 
            Height          =   372
            Left            =   10200
            TabIndex        =   2
            Top             =   96
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
         Begin VB.PictureBox picScalePercent 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6624
            ScaleHeight     =   492
            ScaleWidth      =   1932
            TabIndex        =   0
            Top             =   0
            Width           =   1932
            Begin VB.ComboBox cboScalePercent 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               ItemData        =   "frmPrintPreview.frx":45690
               Left            =   948
               List            =   "frmPrintPreview.frx":456AC
               TabIndex        =   7
               Text            =   "cboScalePercent"
               Top             =   60
               Width           =   975
            End
            Begin VB.Label lblScalePercent 
               Alignment       =   1  'Right Justify
               Caption         =   "# Scale:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   348
               Left            =   36
               TabIndex        =   6
               Top             =   96
               Width           =   804
            End
         End
         Begin VB.Label lblView 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "# View:"
            Height          =   192
            Left            =   3780
            TabIndex        =   5
            Top             =   108
            Width           =   624
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblPageOrientation 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "# Page orientation:"
            Height          =   384
            Left            =   1764
            TabIndex        =   4
            Top             =   36
            Width           =   876
            WordWrap        =   -1  'True
         End
      End
      Begin vbExtra.ToolBarDA tbrBottom 
         Height          =   564
         Left            =   576
         Top             =   11592
         Width           =   7392
         _ExtentX        =   13039
         _ExtentY        =   995
         IconsSize       =   1
         AutoSize        =   0   'False
         PicHeight16     =   32
         PicHeight20     =   40
         PicHeight24     =   48
         ButtonsCount    =   5
         ButtonWidth1    =   2360
         ButtonStyle1    =   4
         ButtonEnabled2  =   0   'False
         ButtonKey2      =   "FirstPage"
         ButtonPic162    =   "frmPrintPreview.frx":456DD
         ButtonPic202    =   "frmPrintPreview.frx":4632F
         ButtonPic242    =   "frmPrintPreview.frx":47641
         ButtonWidth2    =   626
         ButtonToolTipText2=   "# First page"
         ButtonEnabled3  =   0   'False
         ButtonKey3      =   "PreviousPage"
         ButtonPic163    =   "frmPrintPreview.frx":49193
         ButtonPic203    =   "frmPrintPreview.frx":49DE5
         ButtonPic243    =   "frmPrintPreview.frx":4B0F7
         ButtonWidth3    =   626
         ButtonToolTipText3=   "# Previous page"
         ButtonEnabled4  =   0   'False
         ButtonKey4      =   "NextPage"
         ButtonPic164    =   "frmPrintPreview.frx":4CC49
         ButtonPic204    =   "frmPrintPreview.frx":4D89B
         ButtonPic244    =   "frmPrintPreview.frx":4EBAD
         ButtonWidth4    =   626
         ButtonToolTipText4=   "# Next page"
         ButtonEnabled5  =   0   'False
         ButtonKey5      =   "LastPage"
         ButtonPic165    =   "frmPrintPreview.frx":506FF
         ButtonPic205    =   "frmPrintPreview.frx":51351
         ButtonPic245    =   "frmPrintPreview.frx":52663
         ButtonWidth5    =   626
         ButtonToolTipText5=   "# Last page"
         Begin VB.PictureBox picPageNumber 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   3440
            ScaleHeight     =   552
            ScaleWidth      =   2292
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   45
            Width           =   2295
            Begin VB.TextBox txtPage 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   840
               TabIndex        =   9
               Text            =   "1"
               Top             =   120
               Width           =   675
            End
            Begin VB.Label lblPageCount 
               AutoSize        =   -1  'True
               Caption         =   "# of..."
               Height          =   204
               Left            =   1620
               TabIndex        =   18
               Top             =   180
               Width           =   432
            End
            Begin VB.Label lblPage 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "# Page:"
               Height          =   304
               Left            =   144
               TabIndex        =   8
               Top             =   180
               Width           =   588
            End
         End
      End
      Begin VB.Timer tmrtxtPageChanged 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   4500
         Top             =   13380
      End
      Begin VB.PictureBox picAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   240
         ScaleHeight     =   828
         ScaleWidth      =   768
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   10155
         Left            =   15660
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   420
         Width           =   315
      End
      Begin VB.Timer tmrFirstDraw 
         Interval        =   1
         Left            =   120
         Top             =   2460
      End
      Begin VB.PictureBox picPagesContainer 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9915
         Left            =   2988
         ScaleHeight     =   9912
         ScaleWidth      =   9252
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   660
         Visible         =   0   'False
         Width           =   9255
         Begin VB.PictureBox picPage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   9795
            Index           =   0
            Left            =   0
            ScaleHeight     =   9768
            ScaleWidth      =   6888
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   6915
         End
         Begin VB.Shape shpPageShadow 
            BackColor       =   &H00585858&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00303030&
            Height          =   975
            Index           =   0
            Left            =   6480
            Top             =   8880
            Width           =   495
         End
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   10320
         Width           =   15435
      End
      Begin VB.ComboBox cboView 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmPrintPreview.frx":541B5
         Left            =   4980
         List            =   "frmPrintPreview.frx":541B7
         Style           =   2  'Dropdown List
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cboAlternativePopupMenu 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         ItemData        =   "frmPrintPreview.frx":541B9
         Left            =   9108
         List            =   "frmPrintPreview.frx":541BB
         Style           =   2  'Dropdown List
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   216
         Width           =   3240
      End
   End
   Begin VB.Menu mnuPopupView 
      Caption         =   "mnuPopupView"
      Visible         =   0   'False
      Begin VB.Menu mnuView2p 
         Caption         =   "# View 2 pages"
      End
      Begin VB.Menu mnuView3p 
         Caption         =   "# View 3 pages"
      End
      Begin VB.Menu mnuView6p 
         Caption         =   "# View 6 pages"
      End
      Begin VB.Menu mnuView12p 
         Caption         =   "# View 12 pages"
      End
   End
   Begin VB.Menu mnuPopupChangeIconsSize 
      Caption         =   "mnuPopupChangeIconsSize"
      Visible         =   0   'False
      Begin VB.Menu mnuIconsAuto 
         Caption         =   "# Automatic"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuIconsSmall 
         Caption         =   "# Small icons "
      End
      Begin VB.Menu mnuIconsMedium 
         Caption         =   "# Medium icons"
      End
      Begin VB.Menu mnuIconsBig 
         Caption         =   "# Large icons"
      End
      Begin VB.Menu mnuIconsSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIconsShowBottomToolBar 
         Caption         =   "# Show bottom bar"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmPrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Const CB_SHOWDROPDOWN = &H14F

Private Enum efmViewSelectedButtonConstants
    efnViewButtonNormal = 0
    efnViewButtonFitToWidth = 1
    efnViewButtonFitToHeight = 2
    efnViewButtonSevealPages = 3
End Enum

Private Enum efnViewPagesConstants
    efnViewNormal = 1
    efnViewFitWidth = 2
    efnViewFitHeight = 3
    efnViewTwoPages = 4
    efnViewThreePages = 5
    efnViewSixPages = 6
    efnViewTwelvePages = 7
End Enum

Private Enum efnUserToolbarIconsSizeConstants
    efnISAuto = 0
    efnISSmall = 1
    efnISMedium = 2
    efnISBig = 3
End Enum

Private Type COLORADJUSTMENT
    caSize As Integer
    caFlags As Integer
    caIlluminantIndex As Integer
    caRedGamma As Integer
    caGreenGamma As Integer
    caBlueGamma As Integer
    caReferenceBlack As Integer
    caReferenceWhite As Integer
    caContrast As Integer
    caBrightness As Integer
    caColorfulness As Integer
    caRedGreenTint As Integer
End Type

Private Const HALFTONE = 4

Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GetStretchBltMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetColorAdjustment Lib "gdi32" (ByVal hDC As Long, lpca As COLORADJUSTMENT) As Long
Private Declare Function SetColorAdjustment Lib "gdi32" (ByVal hDC As Long, lpca As COLORADJUSTMENT) As Long

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Private WithEvents mMouseWheel As MouseWheelNotifierObject
Attribute mMouseWheel.VB_VarHelpID = -1

Private mPrintFnObject As PrintFnObject
Private mCurrentPageNumber As Long
Private mAllowUserChangeScale As Boolean
Private mPageOrientationButtonsVisible As Boolean
Private mScalePercent As Long
Private mMinScalePercent As Long
Private mMaxScalePercent As Long
Private mUseOneToolBar As Boolean

Public Event PrepareDoc(ByRef Cancel As Boolean)
Public Event FormatOptionsClick(ByRef Canceled As Boolean)
Public Event PageNumbersOptionsClick(ByRef Canceled As Boolean)
Public Event ScaleChange(NewScalePercent As Integer)

Private mRefreshed As Boolean
Private mAvailableScreenHeightSpace As Long
Private mAvailableScreenWidthSpace As Long

Private mPageView As efnViewPagesConstants
Private mPagesOnScreen As Long
Private mRaiseFontEvent As Boolean
Private mSelectingView As Boolean
Private mUserCanChangeIconsSize As Boolean
Private mClickChangeIconsSize_X As Single
Private mFormatButtonToolTipText As String
Private mAddedToScaleForRounding As Long
Public DoNotLoad As Boolean
Private mCanceled As Boolean
Private mLoading As Boolean
Private mpicPageNumberVisible As Boolean
Private mtbrTopIconsSize As Long

Private Sub cboAlternativePopupMenu_Click()
    If cboAlternativePopupMenu.ListIndex < 4 Then
        SetToolBarIconsSize cboAlternativePopupMenu.ListIndex
    Else
        If cboAlternativePopupMenu.ListIndex = 5 Then
            mnuIconsShowBottomToolBar_Click
        End If
    End If
End Sub

Private Sub cboScalePercent_Change()
    Dim iVal As Long
    Static sInside As Boolean
    Static sLastScalePercent As Long
    
    If tmrcboScalePercentChange.Enabled Then Exit Sub
    If sInside Then Exit Sub
    
    sInside = True
    iVal = Val(cboScalePercent.Text)
    If iVal > mMaxScalePercent Then
        iVal = mMaxScalePercent
    End If
    If iVal = sLastScalePercent Then
        sInside = False
        Exit Sub
    End If
    If iVal >= mMinScalePercent Then
        mScalePercent = iVal
        sLastScalePercent = iVal
        cboScalePercent.Text = CStr(mScalePercent) & "%"
        tmrcboScalePercentUpdate.Enabled = False
        
        If mRaiseFontEvent Then
            RaiseEvent ScaleChange(CInt(mScalePercent))
            RaiseEventPrepareDoc
            If mCanceled Then Exit Sub
        End If
    End If

    tbrTop.Buttons("DecreaseScale").Enabled = Val(cboScalePercent.Text) > mMinScalePercent
    tbrTop.Buttons("IncreaseScale").Enabled = Val(cboScalePercent.Text) < mMaxScalePercent
    
    If Me.ActiveControl Is cboScalePercent Then
        cboScalePercent.SelStart = 0
        cboScalePercent.SelLength = Len(cboScalePercent.Text) - 1
    End If
    sInside = False
End Sub

Private Sub cboScalePercent_Click()
    tmrcboScalePercentUpdate.Enabled = False
    mAddedToScaleForRounding = 0
    If Not tmrFirstDraw.Enabled Then
        tmrcboScalePercentUpdate.Enabled = True
    End If
    
    tbrTop.Buttons("DecreaseScale").Enabled = Val(cboScalePercent.Text) >= mMinScalePercent
    tbrTop.Buttons("IncreaseScale").Enabled = Val(cboScalePercent.Text) <= mMaxScalePercent
End Sub

Private Sub cboScalePercent_DropDown()
    tmrcboScalePercentChange.Enabled = False
    tmrcboScalePercentChange.Enabled = True
End Sub

Private Sub cboScalePercent_GotFocus()
    cboScalePercent.SelStart = 0
    cboScalePercent.SelLength = Len(cboScalePercent.Text) - 1
End Sub

Private Sub cboScalePercent_KeyPress(KeyAscii As Integer)
    Dim iVal As Long
    
    mAddedToScaleForRounding = 0
    If KeyAscii = 13 Then
        iVal = Val(cboScalePercent.Text)
        If (iVal < mMinScalePercent) Then
            cboScalePercent.Text = mMinScalePercent & "%"
        End If
        If (iVal > mMaxScalePercent) Then
            cboScalePercent.Text = mMaxScalePercent & "%"
        End If
        tmrcboScalePercentChange.Enabled = False
        cboScalePercent_Change
        If mCanceled Then Exit Sub
    Else
        tmrcboScalePercentChange.Enabled = False
        tmrcboScalePercentChange.Enabled = True
    End If
End Sub

Private Sub cboScalePercent_LostFocus()
    cboScalePercent.SelStart = Len(cboScalePercent.Text)
    cboScalePercent.SelLength = 0
End Sub

Private Sub cboView_Click()
    SelectView efnViewButtonSevealPages, cboView.ListIndex
End Sub

Private Sub cmdClose_2_Click()
    cmdClose_Click
End Sub

Private Sub cmdClose_2_GotFocus()
    cmdClose_GotFocus
End Sub

Private Sub cmdClose_Click()
    If IsFormLoaded(frmSettingGridDataProgress) Then
        frmSettingGridDataProgress.Canceled = True
    Else
        Unload Me
    End If
End Sub

Private Sub cmdClose_GotFocus()
    If IsFormLoaded(frmSettingGridDataProgress) Then
        frmSettingGridDataProgress.Canceled = True
    End If
End Sub

Private Sub cmdDefault_Click()
    Select Case Me.ActiveControl
        Case cboScalePercent
            cboScalePercent_KeyPress 13
        Case txtPage
            txtPage_KeyPress 13
        Case Else
            mPrintFnObject.ShowPrint
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iHandled As Boolean
    Dim iActiveControlName As String
    
    Select Case KeyCode
        Case vbKeyPageDown
            If VScroll1.Visible And Not ((Shift And vbCtrlMask) = vbCtrlMask) Then
                If VScroll1.Value < VScroll1.Max Then
                    iHandled = True
                    If VScroll1.Value + VScroll1.LargeChange < VScroll1.Max Then
                        VScroll1.Value = VScroll1.Value + VScroll1.LargeChange
                    Else
                        VScroll1.Value = VScroll1.Max
                    End If
                End If
            End If
            If Not iHandled Then
                If tbrBottom.Buttons("NextPage").Enabled Then
                    tbrBottom_ButtonClick tbrBottom.Buttons("NextPage")
                End If
            End If
        Case vbKeyPageUp
            If VScroll1.Visible And Not ((Shift And vbCtrlMask) = vbCtrlMask) Then
                If VScroll1.Value > VScroll1.Min Then
                    iHandled = True
                    If VScroll1.Value - VScroll1.LargeChange > VScroll1.Min Then
                        VScroll1.Value = VScroll1.Value - VScroll1.LargeChange
                    Else
                        VScroll1.Value = VScroll1.Min
                    End If
                End If
            End If
            If Not iHandled Then
                If tbrBottom.Buttons("PreviousPage").Enabled Then
                    tbrBottom_ButtonClick tbrBottom.Buttons("PreviousPage")
                    If VScroll1.Visible And Not ((Shift And vbCtrlMask) = vbCtrlMask) Then
                        VScroll1.Value = VScroll1.Max
                    End If
                End If
            End If
        Case vbKeyHome
            On Error Resume Next
            iActiveControlName = Me.ActiveControl.Name
            On Error GoTo 0
            
            Select Case iActiveControlName
                Case "cboScalePercent", "txtPage"
                    '
                Case Else
                    If VScroll1.Visible And Not ((Shift And vbCtrlMask) = vbCtrlMask) Then
                        VScroll1.Value = VScroll1.Min
                    Else
                        If tbrBottom.Buttons("FirstPage").Enabled Then
                            tbrBottom_ButtonClick tbrBottom.Buttons("FirstPage")
                        End If
                    End If
            End Select
        Case vbKeyEnd
            On Error Resume Next
            iActiveControlName = Me.ActiveControl.Name
            On Error GoTo 0
            
            Select Case iActiveControlName
                Case "cboScalePercent", "txtPage"
                    '
                Case Else
                    If VScroll1.Visible And Not ((Shift And vbCtrlMask) = vbCtrlMask) Then
                        VScroll1.Value = VScroll1.Max
                    Else
                        If tbrBottom.Buttons("LastPage").Enabled Then
                            tbrBottom_ButtonClick tbrBottom.Buttons("LastPage")
                        End If
                    End If
                End Select
        Case vbKeyDown
            If Not Me.ActiveControl Is cboScalePercent Then
                If VScroll1.Visible Then
                    If VScroll1.Value + VScroll1.SmallChange <= VScroll1.Max Then
                        VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
                        tmrIgnoreKeysUpDown.Enabled = False
                        tmrIgnoreKeysUpDown.Enabled = True
                    Else
                        If VScroll1.Value <> VScroll1.Max Then
                            VScroll1.Value = VScroll1.Max
                        Else
                            If Not tmrIgnoreKeysUpDown.Enabled Then
                                If mCurrentPageNumber < PrinterExCurrentDocument.PageCount Then
                                    mCurrentPageNumber = mCurrentPageNumber + 1
                                    ShowPages
                                End If
                            End If
                        End If
                    End If
                Else
                    If mCurrentPageNumber < PrinterExCurrentDocument.PageCount Then
                        mCurrentPageNumber = mCurrentPageNumber + 1
                        ShowPages
                    End If
                End If
            End If
        Case vbKeyUp
            If Not Me.ActiveControl Is cboScalePercent Then
                If VScroll1.Visible Then
                    If VScroll1.Value - VScroll1.SmallChange >= VScroll1.Min Then
                        VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
                        tmrIgnoreKeysUpDown.Enabled = False
                        tmrIgnoreKeysUpDown.Enabled = True
                    Else
                        If VScroll1.Value <> VScroll1.Min Then
                            VScroll1.Value = VScroll1.Min
                        Else
                            If Not tmrIgnoreKeysUpDown.Enabled Then
                                If mCurrentPageNumber > 1 Then
                                    mCurrentPageNumber = mCurrentPageNumber - 1
                                    ShowPages
                                    If VScroll1.Visible And Not ((Shift And vbCtrlMask) = vbCtrlMask) Then
                                        VScroll1.Value = VScroll1.Max
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If mCurrentPageNumber > 1 Then
                        mCurrentPageNumber = mCurrentPageNumber - 1
                        ShowPages
                    End If
                End If
            Else
                If mCurrentPageNumber > 1 Then
                    mCurrentPageNumber = mCurrentPageNumber - 1
                    ShowPages
                End If
            End If
        Case vbKeyRight
            If Not Me.ActiveControl Is txtPage Then
                If Not Me.ActiveControl Is cboScalePercent Then
                    If mCurrentPageNumber < PrinterExCurrentDocument.PageCount Then
                        mCurrentPageNumber = mCurrentPageNumber + 1
                        ShowPages
                    End If
                End If
            End If
        Case vbKeyLeft
            If Not Me.ActiveControl Is txtPage Then
                If Not Me.ActiveControl Is cboScalePercent Then
                    If mCurrentPageNumber > 1 Then
                        mCurrentPageNumber = mCurrentPageNumber - 1
                        ShowPages
                    End If
                End If
            End If
    End Select
End Sub

Private Sub Form_Load()
    Dim c As Long
    Dim iLng As Long
    
    If DoNotLoad Then Exit Sub
    mCanceled = False
    
    LoadGUICaptions
    AssignAcceleratorToControl lblScalePercent
    AssignAcceleratorToControl cmdClose
    AssignAcceleratorToControl lblPage
    
    If gButtonsStyle <> -1 Then
        cmdClose_2.Move cmdClose.Left, cmdClose.Top, cmdClose.Width, cmdClose.Height
        cmdClose_2.Caption = cmdClose.Caption
        cmdClose.Visible = False
        cmdClose_2.Default = cmdClose.Default
        cmdClose_2.Cancel = cmdClose.Cancel
        cmdClose_2.Visible = True
        cmdClose_2.TabIndex = cmdClose.TabIndex
        cmdClose_2.ButtonStyle = gButtonsStyle
        Set cmdClose_2.Font = cmdClose.Font
    End If
    
    If mPrintFnObject Is Nothing Then Set mPrintFnObject = New PrintFnObject
    
    mPageOrientationButtonsVisible = True
    
    If (mPrintFnObject.PageSetupFlags And cdePSDisableOrientation) <> 0 Then
        PageOrientationButtonsVisible = False
    End If
    
'    picPageNumber.Visible = False
    
    Set mMouseWheel = New MouseWheelNotifierObject
    mMouseWheel.SetForm Me
    
    Set cboView.Font = GetSystemFont(vxMenuFont)
    
    cboView.Visible = False
    cboAlternativePopupMenu.Visible = False
    
    cboScalePercent.ListIndex = 2
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
    
    mPageView = efnViewFitHeight
    picPagesContainer.BackColor = Me.BackColor
    picMainContainer.BackColor = Me.BackColor
    VScroll1.Width = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    HScroll1.Height = VScroll1.Width
    VScroll1.Visible = False
    HScroll1.Visible = False
    SetTextBoxNumeric txtPage
    
    tbrBottom.Buttons(2).Enabled = False
    tbrBottom.Buttons(3).Enabled = False
    PositionForm
    
    mCurrentPageNumber = 1
    If mPrintFnObject.FromPage > 0 Then
        mCurrentPageNumber = mPrintFnObject.FromPage
    End If
    For c = 1 To 11
        Load picPage(c)
        Load shpPageShadow(c)
    Next c
    cboScalePercent.SelStart = Len(cboScalePercent.Text)
    cboScalePercent.SelLength = 0
    
    cboView.Top = tbrTop.Height - cboView.Height
    ShowPages
    mRaiseFontEvent = True
    PositionControls
End Sub

Private Sub Form_Resize()
    Dim iWidth As Long
    Dim iLng As Long
    
    If DoNotLoad Then Exit Sub
    
    picMainContainer.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    tbrTop.Move 0, 0, picMainContainer.ScaleWidth
    tbrBottom.Move 0, picMainContainer.ScaleHeight - tbrBottom.Height, picMainContainer.ScaleWidth
    cmdDefault.Top = Me.ScaleHeight + 1000
    
    iWidth = (picPageNumber.Width + tbrBottom.ButtonWidth * 4)
    tbrBottom.Buttons(1).Width = (Me.ScaleWidth - iWidth) / 2 + picPageNumber.Width
    
    VScroll1.Move Me.ScaleWidth - VScroll1.Width, tbrTop.Height, GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX, Me.ScaleHeight - tbrTop.Height - IIf(mUseOneToolBar, 0, tbrBottom.Height)
    
    If VScroll1.Visible Then
        HScroll1.Move 0, IIf(mUseOneToolBar, Me.ScaleHeight, tbrBottom.Top) - HScroll1.Height, Me.ScaleWidth - VScroll1.Width, GetSystemMetrics(SM_CXVSCROLL)
    Else
        HScroll1.Move 0, IIf(mUseOneToolBar, Me.ScaleHeight, tbrBottom.Top) - HScroll1.Height, Me.ScaleWidth, GetSystemMetrics(SM_CXVSCROLL)
    End If
    
    PositionControls
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not DoNotLoad Then
        tmrcboScalePercentChange.Enabled = False
        tmrcboScalePercentUpdate.Enabled = False
        ResetPrinter2
    End If
    DoNotLoad = False
End Sub

Private Sub lblPage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If picPageNumber.Container Is tbrTop Then
        tbrTop_MouseUp Button, Shift, x, y
    End If
End Sub

Private Sub lblPageCount_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If picPageNumber.Container Is tbrTop Then
        tbrTop_MouseUp Button, Shift, x, y
    End If
End Sub

Private Sub lblPageOrientation_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tbrTop_MouseUp Button, Shift, x, y
End Sub

Private Sub lblScalePercent_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tbrTop_MouseUp Button, Shift, x, y
End Sub

Private Sub lblView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tbrTop_MouseUp Button, Shift, x, y
End Sub

Private Sub mnuIconsAuto_Click()
    SetToolBarIconsSize efnISAuto
End Sub

Private Sub mnuIconsBig_Click()
    SetToolBarIconsSize efnISBig
End Sub

Private Sub mnuIconsMedium_Click()
    SetToolBarIconsSize efnISMedium
End Sub

Private Sub mnuIconsShowBottomToolBar_Click()
    mnuIconsShowBottomToolBar.Checked = Not mnuIconsShowBottomToolBar.Checked
    mUseOneToolBar = Not mnuIconsShowBottomToolBar.Checked
    Form_Resize
    tmrShowPage.Enabled = True
    SaveSetting AppNameForRegistry, "Preferences", "PrintPreview_ShowBottomToolBar", CStr(CLng(mUseOneToolBar))
End Sub

Private Sub mnuIconsSmall_Click()
    SetToolBarIconsSize efnISSmall
End Sub

Private Sub mnuView12p_Click()
    SelectView efnViewButtonSevealPages, 3
End Sub

Private Sub mnuView2p_Click()
    SelectView efnViewButtonSevealPages, 0
End Sub

Private Sub mnuView3p_Click()
    SelectView efnViewButtonSevealPages, 1
End Sub

Private Sub mnuView6p_Click()
    SelectView efnViewButtonSevealPages, 2
End Sub

Private Sub MouseWheelEnabler1_MouseWheelRotation(Direction As Long)

End Sub

Private Sub picPage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim iLng As Long
    
    If ((Shift And vbCtrlMask) = vbCtrlMask) Or ((Shift And vbShiftMask) = vbShiftMask) Or Button = vbRightButton Then
        If mPageView = efnViewFitHeight Then
            iLng = 0
            If Not PrinterExCurrentDocument Is Nothing Then
                If PrinterExCurrentDocument.PageCount > 6 Then
                    iLng = 3
                ElseIf PrinterExCurrentDocument.PageCount > 3 Then
                    iLng = 2
                ElseIf PrinterExCurrentDocument.PageCount > 2 Then
                    iLng = 1
                End If
            End If
            SelectView efnViewButtonSevealPages, iLng
        ElseIf mPageView = efnViewNormal Then
            SelectView efnViewButtonFitToHeight
        ElseIf mPageView = efnViewFitWidth Then
            SelectView efnViewButtonNormal
        Else
            ' nothing
        End If
    Else
        If mPageView = efnViewFitHeight Then
            iLng = Val(picPage(Index).Tag)
            If iLng > 0 Then
                mCurrentPageNumber = iLng
                SelectView efnViewButtonNormal
            End If
        ElseIf mPageView = efnViewNormal Then
            SelectView efnViewButtonFitToWidth
        ElseIf mPageView = efnViewFitWidth Then
            ' nothing
        Else
            iLng = Val(picPage(Index).Tag)
            If iLng > 0 Then
                mCurrentPageNumber = iLng
                SelectView efnViewButtonFitToHeight
            End If
        End If
    End If
End Sub

Private Sub picPageNumber_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If picPageNumber.Container Is tbrTop Then
        tbrTop_MouseUp Button, Shift, x, y
    End If
End Sub

Private Sub picScalePercent_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tbrTop_MouseUp Button, Shift, x, y
End Sub

Private Sub tbrBottom_ButtonClick(Button As ToolBarDAButton)
    Select Case Button.Key
        Case "FirstPage"
            mCurrentPageNumber = 1
        Case "PreviousPage"
            mCurrentPageNumber = mCurrentPageNumber - mPagesOnScreen
            If mCurrentPageNumber < 1 Then mCurrentPageNumber = 1
        Case "NextPage"
            mCurrentPageNumber = mCurrentPageNumber + mPagesOnScreen
            If mCurrentPageNumber > PrinterExCurrentDocument.PageCount Then mCurrentPageNumber = PrinterExCurrentDocument.PageCount
        Case "LastPage"
            mCurrentPageNumber = PrinterExCurrentDocument.PageCount
    End Select
    
    ShowPages
End Sub

Private Sub EnableDisablePageNavigation()
    If mCurrentPageNumber = 1 Then
        tbrBottom.Buttons("FirstPage").Enabled = False
        tbrBottom.Buttons("PreviousPage").Enabled = False
    Else
        tbrBottom.Buttons("FirstPage").Enabled = True
        tbrBottom.Buttons("PreviousPage").Enabled = True
    End If
    If Not PrinterExCurrentDocument Is Nothing Then
        If (mCurrentPageNumber + mPagesOnScreen - 1) >= PrinterExCurrentDocument.PageCount Then
            tbrBottom.Buttons("NextPage").Enabled = False
            tbrBottom.Buttons("LastPage").Enabled = False
        Else
            tbrBottom.Buttons("NextPage").Enabled = True
            tbrBottom.Buttons("LastPage").Enabled = True
        End If
    Else
        tbrBottom.Buttons("NextPage").Enabled = False
        tbrBottom.Buttons("LastPage").Enabled = False
    End If
    tbrTop.Buttons("FirstPage").Enabled = tbrBottom.Buttons("FirstPage").Enabled
    tbrTop.Buttons("PreviousPage").Enabled = tbrBottom.Buttons("PreviousPage").Enabled
    tbrTop.Buttons("NextPage").Enabled = tbrBottom.Buttons("NextPage").Enabled
    tbrTop.Buttons("LastPage").Enabled = tbrBottom.Buttons("LastPage").Enabled
    
    If mPagesOnScreen = 1 Then
        tbrBottom.Buttons("PreviousPage").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_PreviousPage_Singular)
        tbrBottom.Buttons("NextPage").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Singular)
    Else
        If mCurrentPageNumber = 2 Then
            tbrBottom.Buttons("PreviousPage").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_PreviousPage_Singular)
        Else
            tbrBottom.Buttons("PreviousPage").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_PreviousPage_Plural)
        End If
        If Not PrinterExCurrentDocument Is Nothing Then
            If (mCurrentPageNumber + mPagesOnScreen) = PrinterExCurrentDocument.PageCount Then
                tbrBottom.Buttons("NextPage").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Singular)
            Else
                tbrBottom.Buttons("NextPage").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Plural)
            End If
        Else
            tbrBottom.Buttons("NextPage").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Plural)
        End If
    End If
    tbrTop.Buttons("PreviousPage").ToolTipText = tbrBottom.Buttons("PreviousPage").ToolTipText
    tbrTop.Buttons("NextPage").ToolTipText = tbrBottom.Buttons("NextPage").ToolTipText
End Sub

Private Sub tbrTop_ButtonClick(Button As ToolBarDAButton)
    Dim iCanceled As Boolean
    Dim iSng As Single
    Dim iVal As Long
    Dim iStrScalePercent As String
    Const cScaleStepFactor = 1.12
    If mSelectingView Then Exit Sub
    
    If Button.Key = "ViewSeveralPages" Then
        mnuPopupView.Tag = "1"
        tmrPopupcboView.Enabled = True
        cboView.Left = Button.Left
        cboView.Top = tbrTop.Height - cboView.Height
        PopupMenu mnuPopupView, , Button.Left, tbrTop.Height
        mSelectingView = True
        PressProperViewButtonInToolbar
        mSelectingView = False
    Else
        If Button.Tag = "view" Then
            SelectView Button.Index - tbrTop.Buttons("ViewNormalSize").Index
        ElseIf Button.Tag = "GoToPage" Then
            tbrBottom_ButtonClick Button
        Else
            Select Case Button.Key
                Case "Print"
                    mPrintFnObject.ShowPrint mPrintFnObject.DocKey
                Case "PageSetup"
                    mPrintFnObject.ShowPageSetup
                    If mPrintFnObject.Changed Then
                        Orientation = mPrintFnObject.Orientation
                        RaiseEventPrepareDoc
                        If mCanceled Then Exit Sub
                    End If
                Case "OrientationPortrait", "OrientationLandscape"
                    If Button.Key = "OrientationPortrait" Then
                        If mPrintFnObject.Orientation <> vbPRORPortrait Then
                            ResetPrinter2
                            mPrintFnObject.Orientation = vbPRORPortrait
                            RaiseEventPrepareDoc
                            If mCanceled Then Exit Sub
                        End If
                    Else
                        If mPrintFnObject.Orientation <> vbPRORLandscape Then
                            ResetPrinter2
                            mPrintFnObject.Orientation = vbPRORLandscape
                            RaiseEventPrepareDoc
                            If mCanceled Then Exit Sub
                        End If
                    End If
                    Button.Checked = True
                Case "Format"
'                    If Not PrinterExCurrentDocument Is Nothing Then
'                        PrinterExCurrentDocument.DontCheckError396 = True
'                    End If
                    RaiseEvent FormatOptionsClick(iCanceled)
'                    If Not PrinterExCurrentDocument Is Nothing Then
'                        PrinterExCurrentDocument.DontCheckError396 = False
'                    End If
                    If Not iCanceled Then
                        RaiseEventPrepareDoc
                        If mCanceled Then Exit Sub
                    End If
                Case "PageNumbers"
                    RaiseEvent PageNumbersOptionsClick(iCanceled)
                    If Not iCanceled Then
                        RaiseEventPrepareDoc
                        If mCanceled Then Exit Sub
                    End If
                Case "DecreaseScale"
                    iStrScalePercent = cboScalePercent.Text
                    iSng = Val(iStrScalePercent) / cScaleStepFactor - mAddedToScaleForRounding
                    mAddedToScaleForRounding = 0
                    If iSng < mMinScalePercent Then
                        iSng = mMinScalePercent
                    End If
                    iVal = Val(iStrScalePercent)
                    If iSng > 15 Then
                        If iSng > 120 Then
                            iStrScalePercent = Round(iSng / 20) * 20 & "%"
                        Else
                            iStrScalePercent = Round(iSng / 10) * 10 & "%"
                        End If
                    Else
                        If Round(iSng) = iVal Then
                            iStrScalePercent = Round(iSng - 1) & "%"
                        Else
                            iStrScalePercent = Round(iSng) & "%"
                        End If
                    End If
                    If Val(iStrScalePercent) = iVal Then
                        If Val(Val(iStrScalePercent)) > (mMinScalePercent + 5) Then
                            iStrScalePercent = Val(iStrScalePercent) - 5 & "%"
                        End If
                    End If
                    cboScalePercent.Text = iStrScalePercent
                    cboScalePercent_KeyPress (13)
                Case "IncreaseScale"
                    iStrScalePercent = cboScalePercent.Text
                    iSng = Val(iStrScalePercent) * cScaleStepFactor
                    If Abs(iSng - mMaxScalePercent) <= mMaxScalePercent * 0.05 Then
                        mAddedToScaleForRounding = mMaxScalePercent - iSng
                        iSng = mMaxScalePercent
                    End If
                    iVal = Val(iStrScalePercent)
                    If iSng > 15 Then
                        If iSng > 120 Then
                            iStrScalePercent = Round(iSng / 20) * 20 & "%"
                        Else
                            iStrScalePercent = Round(iSng / 10) * 10 & "%"
                        End If
                    Else
                        If Round(iSng) = iVal Then
                            iStrScalePercent = Int(iSng + 1) & "%"
                        Else
                            iStrScalePercent = Round(iSng) & "%"
                        End If
                    End If
                    If Val(iStrScalePercent) = iVal Then
                        If Val(Val(iStrScalePercent)) < (mMaxScalePercent - 5) Then
                            iStrScalePercent = Val(iStrScalePercent) + 5 & "%"
                        End If
                    End If
                    cboScalePercent.Text = iStrScalePercent
                    cboScalePercent_KeyPress (13)
            End Select
        End If
    End If
End Sub

Private Sub SelectView(nIndex As efmViewSelectedButtonConstants, Optional nNumberOfPagesIndex As Long)
    Dim iP As Long
    Dim c As Long
    
    cboView.Visible = False
    mSelectingView = True
    
    iP = tbrTop.Buttons("ViewNormalSize").Index
    
    For c = 0 To 3
        If c = nIndex Then
            tbrTop.Buttons(c + iP).Checked = True
        Else
            tbrTop.Buttons(c + iP).Checked = False
        End If
    Next c
    mnuPopupView.Tag = ""
    
    Select Case nIndex
        Case 0
            mPageView = efnViewNormal
        Case 1
            mPageView = efnViewFitWidth
        Case 2
            mPageView = efnViewFitHeight
        Case 3
            Select Case nNumberOfPagesIndex
                Case 0
                    mPageView = efnViewTwoPages
                Case 1
                    mPageView = efnViewThreePages
                Case 2
                    mPageView = efnViewSixPages
                Case 3
                    mPageView = efnViewTwelvePages
            End Select
    End Select
    ShowPages
    mSelectingView = False
End Sub

Private Sub ShowPages()
    Dim c As Long
    Dim iUnits As Long
    Dim iPagesWidth(11) As Long
    Dim iPagesHeight(11) As Long
    Dim iTotalWidth As Long
    Dim iProportion As Single
    Dim iMaxHeight As Long
    Dim iMaxRowHeight(3) As Long
    Dim iTotalRowWidth(3) As Long
    Dim iTotalPagesRow(3) As Long
    Dim iAux As Long
    Dim iView As efnViewPagesConstants
    Dim iPageCount As String
    
    'Static sFirst As Boolean
    
    'If Not sFirst Then
        
    If mLoading Then
        If mUserCanChangeIconsSize Then
            iAux = CLng(GetSetting(AppNameForRegistry, "Preferences", "PrintPreview_ToolBarIconsSize", efnISAuto))
            If iAux <> efnISAuto Then
                SetToolBarIconsSize iAux
            End If
        End If
    End If
    
    If Not PrinterExCurrentDocument Is Nothing Then
        If PrinterExCurrentDocument.PageCount = 0 Then
            iPageCount = "..."
        Else
            iPageCount = " " & PrinterExCurrentDocument.PageCount
        End If
        lblPageCount.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_lblPageCount_Caption) & iPageCount
        If PrinterExCurrentDocument.PageCount = 1 Then
            tbrTop.Buttons("ViewSeveralPages").Enabled = False
        Else
            mnuView3p.Enabled = False
            mnuView6p.Enabled = False
            mnuView12p.Enabled = False
            tbrTop.Buttons("ViewSeveralPages").Enabled = True
            If PrinterExCurrentDocument.PageCount > 2 Then
                mnuView3p.Enabled = True
                If PrinterExCurrentDocument.PageCount > 3 Then
                    mnuView6p.Enabled = True
                    If PrinterExCurrentDocument.PageCount > 6 Then
                        mnuView12p.Enabled = True
                    End If
                End If
            End If
            LoadcboView
        End If
        
    End If
'    For c = 0 To picPage.UBound
'        picPage(c).Cls
'        picPage(c).Refresh
'    Next c
    SetWindowRedraw picMainContainer.hWnd, False
    
    For c = 1 To picPage.UBound
        picPage(c).Visible = False
        shpPageShadow(c).Visible = False
    Next c
    For c = 0 To picPage.UBound
        picPage(c).Tag = ""
    Next c
    
    iView = mPageView
    If iView = efnViewFitHeight Then
        If Not PrinterExCurrentDocument Is Nothing Then
            If PrinterExCurrentDocument.PageCount > 0 Then
    '            sFirst = True
                If PrinterExCurrentDocument.PageCount > 1 Then
                    If PrinterExCurrentDocument.Orientation = vbPRORPortrait Then
                        If Me.Width / Me.Height > 1.29 Then
                            'SelectView efnViewButtonSevealPages, 0
                            iView = efnViewTwoPages
                            'Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End If
    If iView >= efnViewTwoPages Then
        If Not PrinterExCurrentDocument Is Nothing Then
            If PrinterExCurrentDocument.PageCount < 2 Then
                iView = efnViewFitHeight
            End If
        End If
    End If
    
    Select Case iView
        Case efnViewTwoPages
            mPagesOnScreen = 2
        Case efnViewThreePages
            mPagesOnScreen = 3
        Case efnViewSixPages
            mPagesOnScreen = 6
        Case efnViewTwelvePages
            mPagesOnScreen = 12
        Case Else
            mPagesOnScreen = 1
    End Select
    If mPagesOnScreen = 1 Then
        lblPage.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_lblPage_Singular_Caption) ' "Page:"
    Else
        lblPage.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_lblPage_Plural_Caption)  ' "Pages:"
    End If
    AssignAcceleratorToControl lblPage
    lblPage.Left = txtPage.Left - lblPage.Width - 110
    
    picPagesContainer.Visible = False
    HScroll1.Visible = False
    VScroll1.Visible = False
    picPage(0).Move 300, 150
    iUnits = mPrintFnObject.Units
    mPrintFnObject.Units = vbMillimeters
    mAvailableScreenHeightSpace = Me.ScaleHeight - tbrTop.Height - IIf(mUseOneToolBar, 0, tbrBottom.Height)
    
    If Not PrinterExCurrentDocument Is Nothing Then
        If PrinterExCurrentDocument.PageCount = 0 Then
'            PrinterExCurrentDocument.DisableEvents = True
'            PrinterExCurrentDocument.PrintText " "
'            PrinterExCurrentDocument.DisableEvents = False
            mCurrentPageNumber = 1
        End If
        If (mCurrentPageNumber + mPagesOnScreen - 1) > PrinterExCurrentDocument.PageCount Then
            mCurrentPageNumber = PrinterExCurrentDocument.PageCount - mPagesOnScreen + 1
            If mCurrentPageNumber < 1 Then mCurrentPageNumber = 1
        End If
        PrinterExCurrentDocument.GoToPage mCurrentPageNumber
        iPagesWidth(0) = PrinterExCurrentDocument.Width
        iPagesHeight(0) = PrinterExCurrentDocument.Height
    Else
        iPagesWidth(0) = Me.ScaleX(mPrintFnObject.PaperWidth, vbMillimeters, vbTwips) + Screen.TwipsPerPixelX - 1
        iPagesHeight(0) = Me.ScaleY(mPrintFnObject.PaperHeight, vbMillimeters, vbTwips) + Screen.TwipsPerPixelY - 1
        If mPrintFnObject.Orientation = vbPRORLandscape Then
            iAux = iPagesWidth(0)
            iPagesWidth(0) = iPagesHeight(0)
            iPagesHeight(0) = iAux
            tbrTop.Buttons("OrientationLandscape").Checked = True
        End If
    End If
    Select Case iView
        Case efnViewNormal
            picPage(0).Width = iPagesWidth(0) '/ 100 * PrinterExCurrentDocument.Zoom
            picPage(0).Height = iPagesHeight(0) ' / 100 * PrinterExCurrentDocument.Zoom
            picPagesContainer.Move picPagesContainer.Left, picPagesContainer.Top, picPage(0).Width + 600, picPage(0).Height + 600
            If Not PrinterExCurrentDocument Is Nothing Then
                AuxPaintPage mCurrentPageNumber, iPagesWidth(0), iPagesHeight(0), picPage(0), PrinterExCurrentDocument.ColorMode = vbPRCMMonochrome
            End If
            txtPage_Text = mCurrentPageNumber
        Case efnViewFitWidth
            picPage(0).Width = Me.ScaleWidth - VScroll1.Width - 600
            iProportion = picPage(0).Width / iPagesWidth(0)
            picPage(0).Height = iPagesHeight(0) * picPage(0).Width / iPagesWidth(0)
            picPagesContainer.Move picPagesContainer.Left, picPagesContainer.Top, picPage(0).Width + 600, picPage(0).Height + 300
            If Not PrinterExCurrentDocument Is Nothing Then
                AuxPaintPage mCurrentPageNumber, iPagesWidth(0), iPagesHeight(0), picPage(0), PrinterExCurrentDocument.ColorMode = vbPRCMMonochrome
            End If
            txtPage_Text = mCurrentPageNumber
        Case efnViewFitHeight
            picPage(0).Height = mAvailableScreenHeightSpace - 300
            picPage(0).Width = iPagesWidth(0) * picPage(0).Height / iPagesHeight(0)
            If picPage(0).Width > (Me.ScaleWidth - 600) Then
                picPage(0).Width = Me.ScaleWidth - 600
                picPage(0).Height = iPagesHeight(0) * picPage(0).Width / iPagesWidth(0)
            End If
            picPagesContainer.Move picPagesContainer.Left, picPagesContainer.Top, picPage(0).Width + 600, picPage(0).Height + 300
            If Not PrinterExCurrentDocument Is Nothing Then
                AuxPaintPage mCurrentPageNumber, iPagesWidth(0), iPagesHeight(0), picPage(0), PrinterExCurrentDocument.ColorMode = vbPRCMMonochrome
            End If
            txtPage_Text = mCurrentPageNumber
        Case efnViewTwoPages
            picPagesContainer.Move 0, 0, Me.ScaleWidth, mAvailableScreenHeightSpace
            If Not PrinterExCurrentDocument Is Nothing Then
                If Not mCurrentPageNumber + 1 > PrinterExCurrentDocument.PageCount Then
                    PrinterExCurrentDocument.GoToPage mCurrentPageNumber + 1
                End If
                iPagesWidth(1) = PrinterExCurrentDocument.Width
                iPagesHeight(1) = PrinterExCurrentDocument.Height
            End If
            
            iMaxHeight = iPagesHeight(0)
            If iPagesHeight(1) > iMaxHeight Then
                iMaxHeight = iPagesHeight(1)
            End If
            
            iProportion = (mAvailableScreenHeightSpace - 300) / iMaxHeight
            
            iTotalWidth = iPagesWidth(0) + iPagesWidth(1)
            
            If iTotalWidth * iProportion > Me.ScaleWidth - 900 Then
                iProportion = (Me.ScaleWidth - 900) / iTotalWidth
            End If
            
            For c = 0 To 1
                picPage(c).Height = iPagesHeight(c) * iProportion
                picPage(c).Width = iPagesWidth(c) * iProportion
                If Not PrinterExCurrentDocument Is Nothing Then
                    If Not mCurrentPageNumber + c > PrinterExCurrentDocument.PageCount Then
                        PrinterExCurrentDocument.GoToPage mCurrentPageNumber + c
                        AuxPaintPage mCurrentPageNumber + c, iPagesWidth(c), iPagesHeight(c), picPage(c), PrinterExCurrentDocument.ColorMode = vbPRCMMonochrome
                    End If
                End If
                picPage(c).Top = (mAvailableScreenHeightSpace - iMaxHeight * iProportion) / 2
                If c = 0 Then
                    picPage(c).Left = (Me.ScaleWidth - (iTotalWidth * iProportion + 300)) / 2
                Else
                    picPage(c).Left = picPage(c - 1).Left + picPage(c - 1).Width + 300
                End If
                picPage(c).Visible = True
                shpPageShadow(c).Visible = True
                If Not PrinterExCurrentDocument Is Nothing Then
                    picPage(c).Tag = PrinterExCurrentDocument.Page
                End If
            Next c
            
            txtPage_Text = mCurrentPageNumber & "-" & mCurrentPageNumber + 1
        Case efnViewThreePages
            picPagesContainer.Move 0, 0, Me.ScaleWidth, mAvailableScreenHeightSpace
            
            If (mCurrentPageNumber + 1) <= PrinterExCurrentDocument.PageCount Then
                PrinterExCurrentDocument.GoToPage mCurrentPageNumber + 1
                iPagesWidth(1) = PrinterExCurrentDocument.Width
                iPagesHeight(1) = PrinterExCurrentDocument.Height
            End If
            
            If (mCurrentPageNumber + 2) <= PrinterExCurrentDocument.PageCount Then
                PrinterExCurrentDocument.GoToPage mCurrentPageNumber + 2
                iPagesWidth(2) = PrinterExCurrentDocument.Width
                iPagesHeight(2) = PrinterExCurrentDocument.Height
            End If
            
            iMaxHeight = iPagesHeight(0)
            If iPagesHeight(1) > iMaxHeight Then
                iMaxHeight = iPagesHeight(1)
            End If
            If iPagesHeight(2) > iMaxHeight Then
                iMaxHeight = iPagesHeight(2)
            End If
            
            iProportion = (mAvailableScreenHeightSpace - 300) / iMaxHeight
            
            iTotalWidth = iPagesWidth(0) + iPagesWidth(1) + iPagesWidth(2)
            
            If iTotalWidth * iProportion > (Me.ScaleWidth - 1200) Then
                iProportion = (Me.ScaleWidth - 1200) / iTotalWidth
            End If
            
            For c = 0 To 2
                picPage(c).Height = iPagesHeight(c) * iProportion
                picPage(c).Width = iPagesWidth(c) * iProportion
                If Not PrinterExCurrentDocument Is Nothing Then
                    If (mCurrentPageNumber + c) <= PrinterExCurrentDocument.PageCount Then
                        PrinterExCurrentDocument.GoToPage mCurrentPageNumber + c
                    End If
                    AuxPaintPage mCurrentPageNumber + c, iPagesWidth(c), iPagesHeight(c), picPage(c), PrinterExCurrentDocument.ColorMode = vbPRCMMonochrome
                End If
                picPage(c).Top = (mAvailableScreenHeightSpace - iMaxHeight * iProportion) / 2
                If c = 0 Then
                    picPage(c).Left = (Me.ScaleWidth - 1200 - (iTotalWidth * iProportion)) / 2 + 300
                Else
                    picPage(c).Left = picPage(c - 1).Left + picPage(c - 1).Width + 300
                End If
                picPage(c).Visible = True
                shpPageShadow(c).Visible = True
                picPage(c).Tag = PrinterExCurrentDocument.Page
            Next c
            
            txtPage_Text = mCurrentPageNumber & "-" & mCurrentPageNumber + 2
        
        Case efnViewSixPages
            picPagesContainer.Move 0, 0, Me.ScaleWidth, mAvailableScreenHeightSpace
            For c = 0 To 2
                If (mCurrentPageNumber + c) <= PrinterExCurrentDocument.PageCount Then
                    PrinterExCurrentDocument.GoToPage mCurrentPageNumber + c
                    iPagesWidth(c) = PrinterExCurrentDocument.Width
                    iPagesHeight(c) = PrinterExCurrentDocument.Height
                    iTotalRowWidth(0) = iTotalRowWidth(0) + iPagesWidth(c)
                    iTotalPagesRow(0) = iTotalPagesRow(0) + 1
                    If iPagesHeight(c) > iMaxRowHeight(0) Then
                        iMaxRowHeight(0) = iPagesHeight(c)
                    End If
                End If
            Next c
            For c = 3 To 5
                If (mCurrentPageNumber + c) <= PrinterExCurrentDocument.PageCount Then
                    PrinterExCurrentDocument.GoToPage mCurrentPageNumber + c
                    iPagesWidth(c) = PrinterExCurrentDocument.Width
                    iPagesHeight(c) = PrinterExCurrentDocument.Height
                    iTotalRowWidth(1) = iTotalRowWidth(1) + iPagesWidth(c)
                    iTotalPagesRow(1) = iTotalPagesRow(1) + 1
                    If iPagesHeight(c) > iMaxRowHeight(1) Then
                        iMaxRowHeight(1) = iPagesHeight(c)
                    End If
                End If
            Next c
            
            iProportion = (mAvailableScreenHeightSpace - 600) / (iMaxRowHeight(0) + iMaxRowHeight(1))
            
            If iTotalRowWidth(0) >= iTotalRowWidth(1) Then
                If iTotalRowWidth(0) * iProportion > (Me.ScaleWidth - 1200) Then
                    iProportion = (Me.ScaleWidth - 1200) / iTotalRowWidth(0)
                End If
            Else
                If iTotalRowWidth(1) * iProportion > (Me.ScaleWidth - 1200) Then
                    iProportion = (Me.ScaleWidth - 1200) / iTotalRowWidth(1)
                End If
            End If
            
            For c = 0 To 5
                If (mCurrentPageNumber + c) <= PrinterExCurrentDocument.PageCount Then
                    picPage(c).Height = iPagesHeight(c) * iProportion
                    picPage(c).Width = iPagesWidth(c) * iProportion
                    If Not PrinterExCurrentDocument Is Nothing Then
                        If mCurrentPageNumber + c <= PrinterExCurrentDocument.PageCount Then
                            PrinterExCurrentDocument.GoToPage mCurrentPageNumber + c
                            AuxPaintPage mCurrentPageNumber + c, iPagesWidth(c), iPagesHeight(c), picPage(c), PrinterExCurrentDocument.ColorMode = vbPRCMMonochrome
                            picPage(c).Visible = True
                            shpPageShadow(c).Visible = True
                        End If
                    End If
                    
                    If Int(c / 3) = 1 Then
                        picPage(c).Top = picPage(0).Top + iMaxRowHeight(0) * iProportion + 150
                    Else
                        picPage(c).Top = (mAvailableScreenHeightSpace - 150 - ((iMaxRowHeight(0) + iMaxRowHeight(1)) * iProportion)) / 2
                    End If
                    If (c = 0) Or (c = 3) Then
                        picPage(c).Left = (Me.ScaleWidth - 600 - (iTotalPagesRow(Int(c / 3)) - 1) * 300 - (iTotalRowWidth(Int(c / 3)) * iProportion)) / 2 + 300
                    Else
                        picPage(c).Left = picPage(c - 1).Left + picPage(c - 1).Width + 300
                    End If
                    picPage(c).Tag = PrinterExCurrentDocument.Page
                End If
            Next c
            
            If PrinterExCurrentDocument.PageCount < mCurrentPageNumber + 5 Then
                txtPage_Text = mCurrentPageNumber & "-" & PrinterExCurrentDocument.PageCount
            Else
                txtPage_Text = mCurrentPageNumber & "-" & mCurrentPageNumber + 5
            End If
            
        Case efnViewTwelvePages
            If (Me.Width / Me.Height > 1.7) And (PrinterExCurrentDocument.Width < PrinterExCurrentDocument.Height) Then
                
                picPagesContainer.Move 0, 0, Me.ScaleWidth, mAvailableScreenHeightSpace
                For c = 0 To 5
                    If (mCurrentPageNumber + c) <= PrinterExCurrentDocument.PageCount Then
                        PrinterExCurrentDocument.GoToPage mCurrentPageNumber + c
                        iPagesWidth(c) = PrinterExCurrentDocument.Width
                        iPagesHeight(c) = PrinterExCurrentDocument.Height
                        iTotalRowWidth(0) = iTotalRowWidth(0) + iPagesWidth(c)
                        iTotalPagesRow(0) = iTotalPagesRow(0) + 1
                        If iPagesHeight(c) > iMaxRowHeight(0) Then
                            iMaxRowHeight(0) = iPagesHeight(c)
                        End If
                    End If
                Next c
                For c = 6 To 11
                    If (mCurrentPageNumber + c) <= PrinterExCurrentDocument.PageCount Then
                        PrinterExCurrentDocument.GoToPage mCurrentPageNumber + c
                        iPagesWidth(c) = PrinterExCurrentDocument.Width
                        iPagesHeight(c) = PrinterExCurrentDocument.Height
                        iTotalRowWidth(1) = iTotalRowWidth(1) + iPagesWidth(c)
                        iTotalPagesRow(1) = iTotalPagesRow(1) + 1
                        If iPagesHeight(c) > iMaxRowHeight(1) Then
                            iMaxRowHeight(1) = iPagesHeight(c)
                        End If
                    End If
                Next c
                
                iProportion = (mAvailableScreenHeightSpace - 600) / (iMaxRowHeight(0) + iMaxRowHeight(1))
                If iTotalRowWidth(0) >= iTotalRowWidth(1) Then
                    If iTotalRowWidth(0) * iProportion > (Me.ScaleWidth - 1800) Then
                        iProportion = (Me.ScaleWidth - 1800) / iTotalRowWidth(0)
                    End If
                Else
                    If iTotalRowWidth(1) * iProportion > (Me.ScaleWidth - 1800) Then
                        iProportion = (Me.ScaleWidth - 1800) / iTotalRowWidth(1)
                    End If
                End If
                
                For c = 0 To 11
                    If (mCurrentPageNumber + c) <= PrinterExCurrentDocument.PageCount Then
                        picPage(c).Height = iPagesHeight(c) * iProportion
                        picPage(c).Width = iPagesWidth(c) * iProportion
                        If Not PrinterExCurrentDocument Is Nothing Then
                            If mCurrentPageNumber + c <= PrinterExCurrentDocument.PageCount Then
                                PrinterExCurrentDocument.GoToPage mCurrentPageNumber + c
                                AuxPaintPage mCurrentPageNumber + c, iPagesWidth(c), iPagesHeight(c), picPage(c), PrinterExCurrentDocument.ColorMode = vbPRCMMonochrome
                                picPage(c).Visible = True
                                shpPageShadow(c).Visible = True
                            End If
                        End If
                        
                        If Int(c / 6) = 1 Then
                            picPage(c).Top = picPage(0).Top + iMaxRowHeight(0) * iProportion + 150
                        Else
                            picPage(c).Top = (mAvailableScreenHeightSpace - 150 - ((iMaxRowHeight(0) + iMaxRowHeight(1)) * iProportion)) / 2
                        End If
                        If (c = 0) Or (c = 6) Then
                            picPage(c).Left = (Me.ScaleWidth - 600 - (iTotalPagesRow(Int(c / 6)) - 1) * 300 - (iTotalRowWidth(Int(c / 6)) * iProportion)) / 2 + 300
                        Else
                            picPage(c).Left = picPage(c - 1).Left + picPage(c - 1).Width + 300
                        End If
                        picPage(c).Tag = PrinterExCurrentDocument.Page
                    End If
                Next c
                
                
                If PrinterExCurrentDocument.PageCount < mCurrentPageNumber + 11 Then
                    txtPage_Text = mCurrentPageNumber & "-" & PrinterExCurrentDocument.PageCount
                Else
                    txtPage_Text = mCurrentPageNumber & "-" & mCurrentPageNumber + 11
                End If
            
            Else
                picPagesContainer.Move 0, 0, Me.ScaleWidth, mAvailableScreenHeightSpace
                For c = 0 To 3
                    If (mCurrentPageNumber + c) <= PrinterExCurrentDocument.PageCount Then
                        PrinterExCurrentDocument.GoToPage mCurrentPageNumber + c
                        iPagesWidth(c) = PrinterExCurrentDocument.Width
                        iPagesHeight(c) = PrinterExCurrentDocument.Height
                        iTotalRowWidth(0) = iTotalRowWidth(0) + iPagesWidth(c)
                        iTotalPagesRow(0) = iTotalPagesRow(0) + 1
                        If iPagesHeight(c) > iMaxRowHeight(0) Then
                            iMaxRowHeight(0) = iPagesHeight(c)
                        End If
                    End If
                Next c
                For c = 4 To 7
                    If (mCurrentPageNumber + c) <= PrinterExCurrentDocument.PageCount Then
                        PrinterExCurrentDocument.GoToPage mCurrentPageNumber + c
                        iPagesWidth(c) = PrinterExCurrentDocument.Width
                        iPagesHeight(c) = PrinterExCurrentDocument.Height
                        iTotalRowWidth(1) = iTotalRowWidth(1) + iPagesWidth(c)
                        iTotalPagesRow(1) = iTotalPagesRow(1) + 1
                        If iPagesHeight(c) > iMaxRowHeight(1) Then
                            iMaxRowHeight(1) = iPagesHeight(c)
                        End If
                    End If
                Next c
                For c = 8 To 11
                    If (mCurrentPageNumber + c) <= PrinterExCurrentDocument.PageCount Then
                        PrinterExCurrentDocument.GoToPage mCurrentPageNumber + c
                        iPagesWidth(c) = PrinterExCurrentDocument.Width
                        iPagesHeight(c) = PrinterExCurrentDocument.Height
                        iTotalRowWidth(2) = iTotalRowWidth(2) + iPagesWidth(c)
                        iTotalPagesRow(2) = iTotalPagesRow(2) + 1
                        If iPagesHeight(c) > iMaxRowHeight(2) Then
                            iMaxRowHeight(2) = iPagesHeight(c)
                        End If
                    End If
                Next c
                
                iProportion = (mAvailableScreenHeightSpace - 600) / (iMaxRowHeight(0) + iMaxRowHeight(1) + iMaxRowHeight(2))
                If (iTotalRowWidth(0) >= iTotalRowWidth(1)) And (iTotalRowWidth(0) >= iTotalRowWidth(2)) Then
                    If iTotalRowWidth(0) * iProportion > (Me.ScaleWidth - 1500) Then
                        iProportion = (Me.ScaleWidth - 1500) / iTotalRowWidth(0)
                    End If
                Else
                    If iTotalRowWidth(1) >= iTotalRowWidth(2) Then
                        If iTotalRowWidth(1) * iProportion > (Me.ScaleWidth - 1500) Then
                            iProportion = (Me.ScaleWidth - 1500) / iTotalRowWidth(1)
                        End If
                    Else
                        If iTotalRowWidth(2) * iProportion > (Me.ScaleWidth - 1500) Then
                            iProportion = (Me.ScaleWidth - 1500) / iTotalRowWidth(2)
                        End If
                    End If
                End If
                
                For c = 0 To 11
                    picPage(c).Height = iPagesHeight(c) * iProportion
                    picPage(c).Width = iPagesWidth(c) * iProportion
                    If Not PrinterExCurrentDocument Is Nothing Then
                        If mCurrentPageNumber + c <= PrinterExCurrentDocument.PageCount Then
                            PrinterExCurrentDocument.GoToPage mCurrentPageNumber + c
                            AuxPaintPage mCurrentPageNumber + c, iPagesWidth(c), iPagesHeight(c), picPage(c), PrinterExCurrentDocument.ColorMode = vbPRCMMonochrome
                            picPage(c).Visible = True
                            shpPageShadow(c).Visible = True
                        End If
                    End If
                    
                    If Int(c / 4) = 1 Then
                        picPage(c).Top = picPage(0).Top + iMaxRowHeight(0) * iProportion + 150
                    Else
                        If Int(c / 4) = 2 Then
                            picPage(c).Top = picPage(4).Top + iMaxRowHeight(1) * iProportion + 150
                        Else
                            picPage(c).Top = (mAvailableScreenHeightSpace - 300 - ((iMaxRowHeight(0) + iMaxRowHeight(1) + iMaxRowHeight(2)) * iProportion)) / 2
                        End If
                    End If
                    If (c = 0) Or (c = 4) Or (c = 8) Then
                        picPage(c).Left = (Me.ScaleWidth - 600 - (iTotalPagesRow(Int(c / 3)) - 1) * 300 - (iTotalRowWidth(Int(c / 4)) * iProportion)) / 2 + 300
                    Else
                        picPage(c).Left = picPage(c - 1).Left + picPage(c - 1).Width + 300
                    End If
                    picPage(c).Tag = PrinterExCurrentDocument.Page
                Next c
                
                If PrinterExCurrentDocument.PageCount < mCurrentPageNumber + 11 Then
                    txtPage_Text = mCurrentPageNumber & "-" & PrinterExCurrentDocument.PageCount
                Else
                    txtPage_Text = mCurrentPageNumber & "-" & mCurrentPageNumber + 11
                End If
            End If
    End Select
    
    For c = 0 To picPage.UBound
        shpPageShadow(c).Move picPage(c).Left + 3 * Screen.TwipsPerPixelX, picPage(c).Top + 3 * Screen.TwipsPerPixelY, picPage(c).Width, picPage(c).Height
    Next c
    
    If picPagesContainer.Height > mAvailableScreenHeightSpace Then
        picPagesContainer.Top = tbrTop.Height
        VScroll1.Max = ((picPagesContainer.Height - mAvailableScreenHeightSpace) / mAvailableScreenHeightSpace * 300 + 1)
        VScroll1.Min = 0
        VScroll1.Value = 0
        VScroll1.SmallChange = 15
        VScroll1.LargeChange = 200
        VScroll1.Visible = True
        VScroll1.ZOrder
        mAvailableScreenWidthSpace = Me.ScaleWidth - VScroll1.Width
    Else
        VScroll1.Visible = False
        picPagesContainer.Top = tbrTop.Height + (mAvailableScreenHeightSpace - picPagesContainer.Height) / 2
        mAvailableScreenWidthSpace = Me.ScaleWidth
    End If
    
    If picPagesContainer.Width > mAvailableScreenWidthSpace Then
        picPagesContainer.Left = 0
        If VScroll1.Visible Then
            HScroll1.Move 0, IIf(mUseOneToolBar, Me.ScaleHeight, tbrBottom.Top) - VScroll1.Width, Me.ScaleWidth - VScroll1.Width, VScroll1.Width
        Else
            HScroll1.Move 0, IIf(mUseOneToolBar, Me.ScaleHeight, tbrBottom.Top) - VScroll1.Width, Me.ScaleWidth, VScroll1.Width
        End If
        HScroll1.Max = (picPagesContainer.Width - mAvailableScreenWidthSpace) / mAvailableScreenWidthSpace * 30 + 1
        HScroll1.Min = 0
        HScroll1.Value = 0
        HScroll1.SmallChange = 15
        HScroll1.LargeChange = 20
        HScroll1.Visible = True
        HScroll1.ZOrder
    Else
        HScroll1.Visible = False
        picPagesContainer.Left = (mAvailableScreenWidthSpace - picPagesContainer.Width) / 2
    End If
    
    mPrintFnObject.Units = iUnits
    picPagesContainer.Visible = True
    SetWindowRedraw picMainContainer.hWnd, True
    'Debug.Print picPagesContainer.Left, picPagesContainer.Top, picPagesContainer.Width, picPagesContainer.Height
    EnableDisablePageNavigation
    
    If Not mLoading Then
        If mUseOneToolBar Then
            If (picPageNumber.Left + picPageNumber.Width) > Me.ScaleWidth Then
                picPageNumber.Visible = False
            Else
                picPageNumber.Visible = mpicPageNumberVisible
                picPageNumber.Refresh
            End If
        Else
            picPageNumber.Visible = mpicPageNumberVisible
        End If
    End If
    
End Sub


Private Sub tbrTop_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If mUserCanChangeIconsSize Then
            tmrCheckPopupMenu.Enabled = True
            mClickChangeIconsSize_X = x
            PopupMenu mnuPopupChangeIconsSize
        End If
    End If
End Sub

Private Sub tmrcboScalePercentChange_Timer()
    tmrcboScalePercentChange.Enabled = False
    cboScalePercent_Change
End Sub

Private Sub tmrcboScalePercentUpdate_Timer()
    If Not tmrFirstDraw.Enabled Then
        tmrcboScalePercentChange.Enabled = False
        cboScalePercent_Change
    End If
End Sub

Private Sub tmrFirstDraw_Timer()
    If Not mPrintFnObject Is Nothing Then
        Orientation = mPrintFnObject.Orientation
    End If
    
    If GetProp(Me.hWnd, "NotVisible") = 1 Then
        Exit Sub
    End If
    tmrFirstDraw.Enabled = False
    RaiseEventPrepareDoc
    If Not mCanceled Then SetFocusTo picScalePercent
End Sub

Public Property Get PrePrintCtrlObj() As PrintFnObject
    Set PrePrintCtrlObj = mPrintFnObject
End Property

Public Property Set PrePrintCtrlObj(nObj As PrintFnObject)
    Set mPrintFnObject = nObj
End Property

Public Sub RefreshPreview()
    ShowPages
    mRefreshed = True
End Sub

Private Sub tmrIgnoreKeysUpDown_Timer()
    tmrIgnoreKeysUpDown.Enabled = False
End Sub

Private Sub tmrIgnoreMouseWheelEvents_Timer()
    tmrIgnoreMouseWheelEvents.Enabled = False
End Sub

Private Sub tmrCheckPopupMenu_Timer()
    tmrCheckPopupMenu.Enabled = False
    
    If IsWindowVisible(FindWindowEx(GetParent(Me.hWnd), ByVal 0&, "#32768", vbNullString)) = 0 Then
        LoadcboAlternativePopupMenu
        cboAlternativePopupMenu.Left = mClickChangeIconsSize_X
        If cboAlternativePopupMenu.Left + cboAlternativePopupMenu.Width > Me.ScaleWidth Then
            cboAlternativePopupMenu.Left = Me.ScaleWidth - cboAlternativePopupMenu.Width
        End If
        cboAlternativePopupMenu.Top = tbrTop.Height - cboAlternativePopupMenu.Height
        cboAlternativePopupMenu.Visible = True
        SendMessage cboAlternativePopupMenu.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
    End If
End Sub

Private Sub tmrPopupcboView_Timer()
    tmrPopupcboView.Enabled = False
    
    If IsWindowVisible(FindWindowEx(GetParent(Me.hWnd), ByVal 0&, "#32768", vbNullString)) = 0 Then
        cboView.Visible = True
        SendMessage cboView.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
    End If
End Sub

Private Sub tmrShowPage_Timer()
    tmrShowPage.Enabled = False
    ShowPages
End Sub

Private Sub tmrtxtPageChanged_Timer()
    Dim iPageNumber As Long
    
    tmrtxtPageChanged.Enabled = False
    iPageNumber = Val(txtPage.Text)
    If iPageNumber > 0 Then
        mCurrentPageNumber = iPageNumber
        ShowPages
        txtPage.SelStart = 0
        txtPage.SelLength = Len(txtPage.Text)
    End If
End Sub

Private Sub tmrUnload_Timer()
    tmrUnload.Enabled = False
    Unload Me
End Sub

Private Sub txtPage_GotFocus()
    txtPage.SelStart = 0
    txtPage.SelLength = Len(txtPage.Text)
End Sub

Private Sub txtPage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        tmrtxtPageChanged_Timer
        KeyAscii = 0
    Else
        tmrtxtPageChanged.Enabled = False
        tmrtxtPageChanged.Enabled = True
    End If
End Sub

Private Sub VScroll1_Change()
    picPagesContainer.Top = tbrTop.Height - VScroll1.Value / 300 * mAvailableScreenHeightSpace
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
End Sub


Private Sub HScroll1_Change()
    picPagesContainer.Left = -HScroll1.Value / 30 * mAvailableScreenWidthSpace
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub AuxPaintPage(nPageNumber As Long, nWidth As Long, nHeight As Long, nPic As PictureBox, Optional nMonochrome As Boolean)
    Dim iCA As COLORADJUSTMENT
    Dim iSuccess As Boolean
    
    picAux.ScaleMode = vbPixels
    picAux.AutoRedraw = True
    nPic.ScaleMode = vbPixels
    
'    picAux.Width = Printer2.Width / Printer2.TwipsPerPixelX * Screen.TwipsPerPixelX
'    picAux.Height = Printer2.Height / Printer2.TwipsPerPixelY * Screen.TwipsPerPixelY
    
    If nPic.Width > Me.ScaleWidth / 2 Then
        picAux.Width = nPic.Width * 2
        picAux.Height = nPic.Height * 2
    Else
        picAux.Width = nPic.Width * 3
        picAux.Height = nPic.Height * 3
    End If
    
    picAux.BackColor = nPic.BackColor
    picAux.Cls
    nPic.Cls
    
    Err.Clear
    On Error Resume Next
    iSuccess = PrinterExCurrentDocument.PaintPageIntoPictureBox(nPageNumber, picAux)
    If (Err.Number <> 0) Or (Not iSuccess) Then
        picAux.Width = nWidth * 1.1
        picAux.Height = nHeight * 1.1
        Err.Clear
        iSuccess = PrinterExCurrentDocument.PaintPageIntoPictureBox(nPageNumber, picAux)
        If (Err.Number <> 0) Or (Not iSuccess) Then
            picAux.Width = nWidth * 1
            picAux.Height = nHeight * 1
            Err.Clear
            iSuccess = PrinterExCurrentDocument.PaintPageIntoPictureBox(nPageNumber, picAux)
            If (Err.Number <> 0) Or (Not iSuccess) Then
                picAux.Width = nWidth * 0.5
                picAux.Height = nHeight * 0.5
                Err.Clear
                iSuccess = PrinterExCurrentDocument.PaintPageIntoPictureBox(nPageNumber, picAux)
            End If
        End If
    End If
'    Debug.Print Err.Number
    On Error GoTo 0
    
    On Error GoTo TheExit:
    
    GetColorAdjustment nPic.hDC, iCA
    iCA.caSize = Len(iCA)
    If nMonochrome Then
        iCA.caColorfulness = -100
    Else
        iCA.caColorfulness = 0
    End If
    If GetStretchBltMode(nPic.hDC) <> HALFTONE Then
        SetStretchBltMode nPic.hDC, HALFTONE
    End If
    SetColorAdjustment nPic.hDC, iCA
    StretchBlt nPic.hDC, 0, 0, nPic.ScaleWidth, nPic.ScaleHeight, picAux.hDC, 0, 0, picAux.ScaleWidth, picAux.ScaleHeight, vbSrcCopy

TheExit:
End Sub

Private Sub RaiseEventPrepareDoc()
    Dim iDocumentPosition As Single
    Dim iOrientationPortrait As Boolean
    Dim iOrientationLandscape  As Boolean
    Dim iP As Long
    Dim c As Long
    Dim iCancel As Boolean
    
    Me.MousePointer = vbHourglass
    mRefreshed = False
    If Not PrinterExCurrentDocument Is Nothing Then
        iDocumentPosition = 0
        If PrinterExCurrentDocument.PageCount > 0 Then
            iDocumentPosition = (mCurrentPageNumber - 1) / PrinterExCurrentDocument.PageCount
        End If
    End If
    
    'picPageNumber.Visible = False
    
    lblStatus.Visible = False
    lblStatus.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_PreparingDoc_Caption)
    PositionPicStatus
    lblStatus.Visible = True
'    lblStatus.Top = tbrBottom.Height / 2 - lblStatus.Height / 2
    If lblStatus.Left + lblStatus.Width + 100 > tbrBottom.Buttons("FirstPage").Left Then
        lblStatus.Caption = ""
        picStatus.Visible = False
    Else
        picStatus.Visible = True
        lblStatus.Refresh
    End If
    RaiseEvent PrepareDoc(iCancel)
    lblStatus.Caption = ""
    picStatus.Visible = False
    
    If iCancel Then
        On Error Resume Next
        Err.Clear
        Unload Me
        If Err.Number Then
            tmrUnload.Enabled = True
        End If
        On Error GoTo 0
        mCanceled = True
        Exit Sub
    End If
    mCurrentPageNumber = iDocumentPosition * PrinterExCurrentDocument.PageCount + 1
    If mCurrentPageNumber < 1 Then mCurrentPageNumber = 1
    If mCurrentPageNumber > PrinterExCurrentDocument.PageCount Then mCurrentPageNumber = PrinterExCurrentDocument.PageCount
    
    If Not mRefreshed Then ShowPages
    Me.MousePointer = vbDefault
    
    If Not PrinterExCurrentDocument Is Nothing Then
        If PrinterExCurrentDocument.PageCount > 0 Then
            iDocumentPosition = mCurrentPageNumber / PrinterExCurrentDocument.PageCount
            iP = PrinterExCurrentDocument.Page
            For c = 1 To PrinterExCurrentDocument.PageCount
                PrinterExCurrentDocument.GoToPage c
                If PrinterExCurrentDocument.Orientation = vbPRORPortrait Then
                    iOrientationPortrait = True
                Else
                    iOrientationLandscape = True
                End If
            Next c
        End If
    End If
    Select Case True
        Case iOrientationPortrait And Not iOrientationLandscape
            mPrintFnObject.Orientation = vbPRORPortrait
        Case iOrientationLandscape And Not iOrientationPortrait
            mPrintFnObject.Orientation = vbPRORLandscape
        Case Else
            ' nothing
    End Select
    If iP > 0 Then
        PrinterExCurrentDocument.GoToPage iP
    End If
End Sub


Public Property Let FormatButtonVisible(nValue As Boolean)
    tbrTop.Buttons("Format").Visible = nValue
    PositionControls
End Property

Public Property Let PageSetupButtonVisible(nValue As Boolean)
    tbrTop.Buttons("PageSetup").Visible = nValue
    PositionControls
End Property

Public Property Let PageNumbersButtonVisible(nValue As Boolean)
    tbrTop.Buttons("PageNumbers").Visible = nValue
    PositionControls
End Property

Public Property Let FormatButtonToolTipText(nValue As String)
    mFormatButtonToolTipText = nValue
    tbrTop.Buttons("Format").ToolTipText = Trim(mFormatButtonToolTipText)
End Property

Public Property Set FormatButtonPicture(nSizeIdentifier As VBExToolbarDAButtonIconSizeConstants, nPic As StdPicture)
    If (nSizeIdentifier < vxIZ16) Or (nSizeIdentifier > vxIZ36) Then
        RaiseError 381, TypeName(Me) ' Invalid property array index
        Exit Property
    End If
    Set tbrTop.Buttons("Format").Picture(nSizeIdentifier) = nPic
End Property

Public Property Get FormatButtonPicture(nSizeIdentifier As VBExToolbarDAButtonIconSizeConstants) As StdPicture
    If (nSizeIdentifier < vxIZ16) Or (nSizeIdentifier > vxIZ36) Then
        RaiseError 381, TypeName(Me) ' Invalid property array index
        Exit Property
    End If
    Set FormatButtonPicture = tbrTop.Buttons("Format").Picture(nSizeIdentifier)
End Property

Public Property Get AllowUserChangeScale() As Boolean
    AllowUserChangeScale = mAllowUserChangeScale
End Property

Public Property Let AllowUserChangeScale(nValue As Boolean)
    If nValue <> mAllowUserChangeScale Then
        mAllowUserChangeScale = nValue
        PositionControls
'    picScalePercent.Visible = mAllowUserChangeScale
'    tbrTop.Buttons("DecreaseScale").Visible = mAllowUserChangeScale
'    tbrTop.Buttons("IncreaseScale").Visible = mAllowUserChangeScale
    End If
End Property

Public Property Get PageOrientationButtonsVisible() As Boolean
    PageOrientationButtonsVisible = mPageOrientationButtonsVisible
End Property

Public Property Let PageOrientationButtonsVisible(nValue As Boolean)
    If nValue <> mPageOrientationButtonsVisible Then
'        tbrTop.Visible = False
        mPageOrientationButtonsVisible = nValue
        tbrTop.Buttons("OrientationLabelSpace").Visible = mPageOrientationButtonsVisible
        lblPageOrientation.Visible = mPageOrientationButtonsVisible
        tbrTop.Buttons("OrientationPortrait").Visible = mPageOrientationButtonsVisible
        tbrTop.Buttons("OrientationLandscape").Visible = mPageOrientationButtonsVisible
        PositionControls
'        lblView.Left = tbrTop.Buttons("ViewNormalSize").Left - lblView.Width - 60
'        picScalePercent.Left = tbrTop.Buttons("DecreaseScale").Left - picScalePercent.Width - 30
'        tbrTop.Visible = True
'        tbrTop.Refresh
'        Me.Refresh
    End If
End Property


Private Sub LoadcboView()
    Dim iCtl As Control
    
    cboView.Clear
    For Each iCtl In Me.Controls
        If TypeOf iCtl Is Menu Then
            If Left(iCtl.Name, 7) = "mnuView" Then
                cboView.AddItem "    " & iCtl.Caption & "    "
            End If
        End If
    Next iCtl
    
    AutoSizeDropDownWidth cboView
    
    If PrinterExCurrentDocument.PageCount <= 6 Then
        cboView.RemoveItem (cboView.ListCount - 1)
        If PrinterExCurrentDocument.PageCount <= 3 Then
            cboView.RemoveItem (cboView.ListCount - 1)
            If PrinterExCurrentDocument.PageCount <= 2 Then
                cboView.RemoveItem (cboView.ListCount - 1)
            End If
        End If
    End If

End Sub

Private Sub LoadcboAlternativePopupMenu()
    Dim iCtl As Control
    Dim iStr As String
    Dim iPastSep As Boolean
    
    cboAlternativePopupMenu.Clear
    For Each iCtl In Me.Controls
        If TypeOf iCtl Is Menu Then
            If Left(iCtl.Name, 8) = "mnuIcons" Then
                iStr = iCtl.Caption
                If iStr = "-" Then
                    iStr = "---------------------"
                    iPastSep = True
                End If
                cboAlternativePopupMenu.AddItem "    " & iStr & IIf(iCtl.Checked, IIf(iPastSep, "  (" & GetLocalizedString(efnGUIStr_frmPrintPreview_CurrentlyShown) & ")", "  (" & GetLocalizedString(efnGUIStr_frmPrintPreview_CurrentlySelected) & ")"), "") & "    "
            End If
        End If
    Next iCtl
    
    cboAlternativePopupMenu.Width = AutoSizeDropDownWidth(cboAlternativePopupMenu) * Screen.TwipsPerPixelX
    
End Sub


Public Property Let Orientation(nValue As Integer)
    If nValue = vbPRORLandscape Then
        tbrTop.Buttons("OrientationLandscape").Checked = True
    Else
        tbrTop.Buttons("OrientationPortrait").Checked = True
    End If
End Property

Public Sub PressProperViewButtonInToolbar()
    Dim iP As Long
    Dim c As Long
    Dim iIndex As Long
    
    Select Case mPageView
        Case efnViewNormal
            iIndex = 0
        Case efnViewFitWidth
            iIndex = 1
        Case efnViewFitHeight
            iIndex = 2
        Case Else
            iIndex = 3
    End Select
    
    iP = tbrTop.Buttons("ViewNormalSize").Index
    
    For c = 0 To 3
        If c = iIndex Then
            tbrTop.Buttons(c + iP).Checked = True
        Else
            tbrTop.Buttons(c + iP).Checked = False
        End If
    Next c
    
End Sub

Public Property Let ScalePercent(nValue As Long)
    Dim c As Long
    
    If nValue <> mScalePercent Then
        mScalePercent = nValue
        mRaiseFontEvent = False
        
        c = 0
        Do While Val(cboScalePercent.List(c)) <= nValue
            c = c + 1
        Loop
        c = c - 1
        If c < 0 Then c = 0
        cboScalePercent.ListIndex = c
        
        cboScalePercent.Text = nValue & "%"
        
        mRaiseFontEvent = True
    End If
End Property

Public Property Get ScalePercent() As Long
    ScalePercent = mScalePercent
End Property

Private Sub mMouseWheel_MouseWheelRotation(Direction As Long, Handled As Boolean)
    If tmrIgnoreMouseWheelEvents.Enabled Then Exit Sub
    If VScroll1.Visible Then
        If Direction = 1 Then
            If (VScroll1.Value + VScroll1.SmallChange) <= VScroll1.Max Then
                VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
            Else
                If VScroll1.Value = VScroll1.Max Then
                    If mCurrentPageNumber < PrinterExCurrentDocument.PageCount Then
                        mCurrentPageNumber = mCurrentPageNumber + 1
                        ShowPages
                        VScroll1.Value = VScroll1.Min
                        tmrIgnoreMouseWheelEvents.Enabled = False
                        tmrIgnoreMouseWheelEvents.Enabled = True
                    End If
                Else
                    VScroll1.Value = VScroll1.Max
                End If
            End If
        Else
            If (VScroll1.Value - VScroll1.SmallChange) >= VScroll1.Min Then
                VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
            Else
                If VScroll1.Value = VScroll1.Min Then
                    If mCurrentPageNumber > 1 Then
                        mCurrentPageNumber = mCurrentPageNumber - 1
                        ShowPages
                        VScroll1.Value = VScroll1.Max
                        tmrIgnoreMouseWheelEvents.Enabled = False
                        tmrIgnoreMouseWheelEvents.Enabled = True
                    End If
                Else
                    VScroll1.Value = VScroll1.Min
                End If
            End If
        End If
        Handled = True
    Else
        If Not Handled Then
            If Direction = 1 Then
                If mCurrentPageNumber < PrinterExCurrentDocument.PageCount Then
                    mCurrentPageNumber = mCurrentPageNumber + 1
                    ShowPages
'                    tmrIgnoreMouseWheelEvents.Enabled = False
'                    tmrIgnoreMouseWheelEvents.Enabled = True
                End If
            Else
                If mCurrentPageNumber > 1 Then
                    mCurrentPageNumber = mCurrentPageNumber - 1
                    ShowPages
'                    tmrIgnoreMouseWheelEvents.Enabled = False
'                    tmrIgnoreMouseWheelEvents.Enabled = True
                End If
            End If
        End If

    End If
End Sub

Public Sub PositionControls()
    Dim iFontSize As Single
    Dim iScreenWidth As Long
    Dim iLastToolBarElementPos As Long
    Dim iAuxFont As StdFont
    Dim iLng As Long
    Dim iPosEndButtonViewSeveralPages  As Long
    Dim iPosEndButtonIncreaseScale  As Long
    Dim iPosEndGotoPageLastButton As Long
    Dim iPosEndpicPageNumber As Long
    Const cDefaultWidthCloseButton = 2265
    Dim iCloseButtonVisible As Boolean
    Dim iRedraw As Boolean
    
    If mLoading Then Exit Sub
    
    iRedraw = tbrTop.Redraw
    mnuIconsShowBottomToolBar.Checked = Not mUseOneToolBar
    tbrTop.Redraw = False
    picScalePercent.Visible = False
    iScreenWidth = Screen.Width / Screen.TwipsPerPixelX
    cboView.Top = tbrTop.Height - cboView.Height
    cboAlternativePopupMenu.Top = tbrTop.Height - cboAlternativePopupMenu.Height
    
    lblPageOrientation.AutoSize = True
    lblPageOrientation.WordWrap = False
    If lblPageOrientation.Width > 1000 Then
        lblPageOrientation.WordWrap = True
        lblPageOrientation.Width = 876
    End If
    tbrTop.Buttons("OrientationLabelSpace").Width = lblPageOrientation.Width + IIf(Not ((iScreenWidth < 1025) And (tbrTop.IconsSize = vxIconsBig)), 300, 100)
    lblPageOrientation.Left = tbrTop.Buttons("OrientationPortrait").Left - lblPageOrientation.Width - 60
    lblPageOrientation.Top = tbrTop.Height / 2 - lblPageOrientation.Height / 2
    
    lblView.AutoSize = True
    lblView.WordWrap = False
    If lblView.Width > 1000 Then
        lblView.WordWrap = True
        lblView.Width = 876
    End If
    lblView.Top = tbrTop.Height / 2 - lblView.Height / 2
    tbrTop.Buttons("ViewLabelSpace").Width = lblView.Width + IIf(Not ((iScreenWidth < 1025) And (tbrTop.IconsSize = vxIconsBig)), 300, 100)
    lblView.Left = tbrTop.Buttons("ViewNormalSize").Left - lblView.Width - 60
    
    tbrBottom.Visible = Not mUseOneToolBar
    
    iFontSize = 9
    If iFontSize < 7.5 Then iFontSize = 7.5
    
    lblScalePercent.FontSize = iFontSize
    cboScalePercent.FontSize = iFontSize
    lblPageOrientation.FontSize = iFontSize
    lblView.FontSize = iFontSize
    
    picScalePercent.Top = 0
    picScalePercent.Height = tbrTop.Height
    cboScalePercent.Top = picScalePercent.Height / 2 - cboScalePercent.Height / 2
    picScalePercent.FontName = lblScalePercent.FontName
    picScalePercent.FontSize = iFontSize
    lblScalePercent.Top = picScalePercent.Height / 2 - picScalePercent.TextHeight(lblScalePercent.Caption) / 2
    lblScalePercent.Left = cboScalePercent.Left - lblScalePercent.Width - 60
    
    lblStatus.FontSize = iFontSize
    picStatus.Font = lblStatus.Font
    picStatus.FontSize = lblStatus.FontSize
    
    If mUseOneToolBar Then
        If Not picPageNumber.Container Is tbrTop Then
            Set picPageNumber.Container = tbrTop
        End If
        picPageNumber.Top = 0
        picPageNumber.Height = tbrTop.Height
    Else
        If Not picPageNumber.Container Is tbrBottom Then
            Set picPageNumber.Container = tbrBottom
        End If
        picPageNumber.Top = 0
        picPageNumber.Height = tbrBottom.Height
    End If
    
    tbrTop.Buttons("ScaleSpace").Width = picScalePercent.Width
    picScalePercent.Left = tbrTop.Buttons("ScaleSpace").Left - 30
    picScalePercent.Visible = mAllowUserChangeScale
    tbrTop.Buttons("ScaleSpace").Visible = mAllowUserChangeScale
    tbrTop.Buttons("DecreaseScale").Visible = mAllowUserChangeScale
    tbrTop.Buttons("IncreaseScale").Visible = mAllowUserChangeScale
    
    iPosEndButtonViewSeveralPages = tbrTop.Buttons("ViewSeveralPages").Left + tbrTop.Buttons("ViewSeveralPages").Width
    If mAllowUserChangeScale Then
        iPosEndButtonIncreaseScale = tbrTop.Buttons("IncreaseScale").Left + tbrTop.Buttons("IncreaseScale").Width
        iLastToolBarElementPos = iPosEndButtonIncreaseScale
        picScalePercent.Refresh
    Else
        iLastToolBarElementPos = iPosEndButtonViewSeveralPages
    End If
    
    tbrTop.Buttons("FirstPage").Visible = mUseOneToolBar
    tbrTop.Buttons("PreviousPage").Visible = mUseOneToolBar
    tbrTop.Buttons("NextPage").Visible = mUseOneToolBar
    tbrTop.Buttons("LastPage").Visible = mUseOneToolBar
    
    If mUseOneToolBar Then
        iPosEndGotoPageLastButton = tbrTop.Buttons("LastPage").Left + tbrTop.Buttons("LastPage").Width
        iPosEndpicPageNumber = iPosEndGotoPageLastButton + 300 + picPageNumber.Width
        iLastToolBarElementPos = iPosEndpicPageNumber
    End If
    
    If tbrTop.IconsSize = vxIconsSmall Then
        cmdClose.Font.Size = 9
    Else
        cmdClose.Font.Size = 10
    End If
    
    cmdClose.Height = 405
    If (cmdClose.Height + 6 * Screen.TwipsPerPixelY) > tbrTop.Height Then
        cmdClose.Height = tbrTop.Height - 6 * Screen.TwipsPerPixelY
    End If
    cmdClose.Top = (tbrTop.Height - cmdClose.Height) / 2 - Screen.TwipsPerPixelY
    
    iCloseButtonVisible = True
    mpicPageNumberVisible = True
    If (Me.ScaleWidth - iLastToolBarElementPos - 120) > cDefaultWidthCloseButton Then
        cmdClose.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_cmdClose_Caption)
        cmdClose.Width = cDefaultWidthCloseButton
    Else
        cmdClose.Caption = GetLocalizedString(efnGUIStr_General_CloseButton_Caption)
        Set iAuxFont = Me.Font
        Set Me.Font = cmdClose.Font
        iLng = Me.TextWidth(cmdClose.Caption) + 120
        If iLng < 900 Then iLng = 900
        cmdClose.Width = iLng
        Set Me.Font = iAuxFont
        If (Me.ScaleWidth - iLastToolBarElementPos - 60) < cmdClose.Width Then
            If mUseOneToolBar Then
                mpicPageNumberVisible = False
                iLastToolBarElementPos = iPosEndGotoPageLastButton
                If (Me.ScaleWidth - iPosEndGotoPageLastButton - 60) < cmdClose.Width Then
                    iCloseButtonVisible = False
                End If
            Else
                iCloseButtonVisible = False
            End If
        End If
    End If
    
'    If iCloseButtonVisible Then mpicPageNumberVisible = mUseOneToolBar
        
    mpicPageNumberVisible = mpicPageNumberVisible And (Me.ScaleWidth - iPosEndpicPageNumber - 30) > 0
    
    If mUseOneToolBar Then
        If Not mpicPageNumberVisible Then
            ' check to hide more things
            If mAllowUserChangeScale Then
                If (Me.ScaleWidth - iPosEndButtonIncreaseScale - 30) < 0 Then
                    tbrTop.Buttons("DecreaseScale").Visible = False
                    tbrTop.Buttons("IncreaseScale").Visible = False
                    iLastToolBarElementPos = tbrTop.Buttons("ScaleSpace").Left + tbrTop.Buttons("ScaleSpace").Width
                End If
            End If
            
            If (Me.ScaleWidth - (tbrTop.Buttons("LastPage").Left + tbrTop.Buttons("LastPage").Width) - 30) < 0 Then
                tbrTop.Buttons("FirstPage").Visible = False
                tbrTop.Buttons("LastPage").Visible = False
                iLastToolBarElementPos = tbrTop.Buttons("NextPage").Left + tbrTop.Buttons("NextPage").Width
                'tbrTop.Buttons("ScaleSpace").Width = tbrTop.Buttons("ScaleSpace").Width - 250
            End If
            If (Me.ScaleWidth - (tbrTop.Buttons("NextPage").Left + tbrTop.Buttons("NextPage").Width) - 30) < 0 Then
                tbrTop.Buttons("PreviousPage").Visible = False
                tbrTop.Buttons("NextPage").Visible = False
                If mAllowUserChangeScale Then
                    iLastToolBarElementPos = tbrTop.Buttons("ScaleSpace").Left + tbrTop.Buttons("ScaleSpace").Width
                Else
                    iLastToolBarElementPos = iPosEndButtonViewSeveralPages
                End If
            End If
        End If
    Else
        If mAllowUserChangeScale Then
            If (Me.ScaleWidth - iPosEndButtonIncreaseScale - 30) < 0 Then
                tbrTop.Buttons("DecreaseScale").Visible = False
                tbrTop.Buttons("IncreaseScale").Visible = False
                iLastToolBarElementPos = tbrTop.Buttons("ScaleSpace").Left + tbrTop.Buttons("ScaleSpace").Width
            End If
        End If
    End If
    
    If iCloseButtonVisible Then
        'cmdClose.Left = iLastToolBarElementPos
        ' tbrTop.Buttons("IncreaseScale").Left + tbrTop.Buttons("IncreaseScale").Width
        cmdClose.Left = (Me.ScaleWidth + iLastToolBarElementPos - cmdClose.Width) / 2
        If gButtonsStyle <> -1 Then
            cmdClose_2.Font.Size = cmdClose.Font.Size
            cmdClose_2.Caption = cmdClose.Caption
            cmdClose_2.Move cmdClose.Left, cmdClose.Top, cmdClose.Width, cmdClose.Height
            cmdClose_2.Visible = True
            cmdClose.Visible = False
        Else
            cmdClose.Visible = True
        End If
    Else
        cmdClose.Visible = False
        cmdClose_2.Visible = False
    End If
    
    If mpicPageNumberVisible Then
        txtPage.FontSize = iFontSize
        txtPage.Height = cboScalePercent.Height
        lblPage.FontSize = iFontSize
        lblPage.Left = txtPage.Left - lblPage.Width - 110
        picPageNumber.FontName = txtPage.FontName
        picPageNumber.FontSize = iFontSize
        lblPageCount.Left = txtPage.Left + txtPage.Width + 100
        
        lblPageCount.FontSize = iFontSize
        txtPage.Top = (picPageNumber.ScaleHeight - txtPage.Height) / 2
        lblPage.Top = picPageNumber.ScaleHeight / 2 - picPageNumber.TextHeight("awqigWÑÓ") / 2 - Screen.TwipsPerPixelY '* 2
        lblPageCount.Top = lblPage.Top
        
        
        If mUseOneToolBar Then
            picPageNumber.Left = tbrTop.Buttons("LastPage").Left + tbrTop.Buttons("LastPage").Width + 300
        Else
            picPageNumber.Left = tbrBottom.Buttons(2).Left - tbrBottom.ButtonWidth - picPageNumber.Width
        End If
    End If
    
    picPageNumber.Visible = mpicPageNumberVisible
    
    If Not Me.ActiveControl Is cboScalePercent Then
        cboScalePercent.SelStart = Len(cboScalePercent.Text)
        cboScalePercent.SelLength = 0
    End If
    
    tbrTop.Redraw = iRedraw
    
End Sub

Private Function SetIconsAutoSize() As Boolean
    If Screen.Width / Screen.TwipsPerPixelX < 1100 Then
        If tbrTop.IconsSize <> vxIconsSmall Then
            tbrTop.IconsSize = vxIconsSmall
            SetIconsAutoSize = True
        End If
    ElseIf Screen.Width / Screen.TwipsPerPixelX > 2000 Then
        If tbrTop.IconsSize <> vxIconsBig Then
            tbrTop.IconsSize = vxIconsBig
            SetIconsAutoSize = True
        End If
    Else
        If tbrTop.IconsSize <> vxIconsMedium Then
            tbrTop.IconsSize = vxIconsMedium
            SetIconsAutoSize = True
        End If
    End If

End Function

Private Sub SetToolBarIconsSize(nSize As efnUserToolbarIconsSizeConstants)
    Dim iChanged As Boolean
    
    cboAlternativePopupMenu.Visible = False
    mnuIconsAuto.Checked = False
    mnuIconsSmall.Checked = False
    mnuIconsMedium.Checked = False
    mnuIconsBig.Checked = False
    tbrTop.Redraw = False
    If nSize = efnISAuto Then
        iChanged = SetIconsAutoSize
        mnuIconsAuto.Checked = True
    ElseIf nSize = efnISSmall Then
        If tbrTop.IconsSize <> vxIconsSmall Then
            tbrTop.IconsSize = vxIconsSmall
            iChanged = True
        End If
        mnuIconsSmall.Checked = True
    ElseIf nSize = efnISMedium Then
        If tbrTop.IconsSize <> vxIconsMedium Then
            tbrTop.IconsSize = vxIconsMedium
            iChanged = True
        End If
        mnuIconsMedium.Checked = True
    Else
        If tbrTop.IconsSize <> vxIconsBig Then
            tbrTop.IconsSize = vxIconsBig
            iChanged = True
        End If
        mnuIconsBig.Checked = True
    End If
    If Not tmrFirstDraw.Enabled Then
        SaveSetting AppNameForRegistry, "Preferences", "PrintPreview_ToolBarIconsSize", CStr(nSize)
    End If
    If iChanged Then
        Form_Resize
        tmrShowPage.Enabled = True
    End If
    tbrTop.Redraw = True
End Sub

Private Property Let txtPage_Text(nValue As String)
    Dim iWidth As Long
    
    txtPage.Text = nValue
    
    iWidth = picPageNumber.TextWidth(txtPage.Text) + 6 * Screen.TwipsPerPixelX
    If iWidth > txtPage.Width Then
        txtPage.Width = iWidth
        lblPageCount.Left = txtPage.Left + txtPage.Width + 100
        If picPageNumber.Width < (lblPageCount.Left + lblPageCount.Width) Then
            picPageNumber.Width = lblPageCount.Left + lblPageCount.Width
            If Not mUseOneToolBar Then
                If picPageNumber.Left + picPageNumber.Width + 30 > tbrBottom.Buttons("FirstPage").Left Then
                    picPageNumber.Left = tbrBottom.Buttons("FirstPage").Left - picPageNumber.Width - 30
                End If
            End If
        End If
    End If
End Property
    
Private Sub LoadGUICaptions()
    mnuView2p.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_mnuView2p_Caption)
    mnuView3p.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_mnuView3p_Caption)
    mnuView6p.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_mnuView6p_Caption)
    mnuView12p.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_mnuView12p_Caption)
    mnuIconsAuto.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_mnuIconsAuto_Caption)
    mnuIconsSmall.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_mnuIconsSmall_Caption)
    mnuIconsMedium.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_mnuIconsMedium_Caption)
    mnuIconsBig.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_mnuIconsBig_Caption)
    mnuIconsShowBottomToolBar.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_mnuIconsShowBottomToolBar_Caption)
    
    lblPageOrientation.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_lblPageOrientation_Caption)
    lblView.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_lblView_Caption)
    lblScalePercent.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_lblScalePercent_Caption)
    lblPage.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_lblPage_Singular_Caption)
    lblPageCount.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_lblPageCount_Caption) & "..."
    cmdClose.Caption = GetLocalizedString(efnGUIStr_frmPrintPreview_cmdClose_Caption)
    
    tbrTop.Buttons("Print").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Print)
    tbrTop.Buttons("PageSetup").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_PageSetup)
    tbrTop.Buttons("PageNumbers").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_PageNumbersOptions)
    If mFormatButtonToolTipText = "" Then
        tbrTop.Buttons("Format").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_Format)
    End If
    tbrTop.Buttons("OrientationPortrait").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_OrientationPortrait)
    tbrTop.Buttons("OrientationLandscape").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_OrientationLandscape)
    tbrTop.Buttons("ViewNormalSize").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewNormalSize)
    tbrTop.Buttons("ViewScreenWidth").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewScreenWidth)
    tbrTop.Buttons("ViewScreenHeight").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewScreenHeight)
    tbrTop.Buttons("ViewSeveralPages").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_ViewSeveralPages)
    tbrTop.Buttons("DecreaseScale").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_DecreaseScale)
    tbrTop.Buttons("IncreaseScale").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrTop_Buttons_ToolTipText_IncreaseScale)
    
    tbrBottom.Buttons("FirstPage").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_FirstPage)
    tbrBottom.Buttons("PreviousPage").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_PreviousPage_Singular)
    tbrBottom.Buttons("NextPage").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_NextPage_Singular)
    tbrBottom.Buttons("LastPage").ToolTipText = GetLocalizedString(efnGUIStr_frmPrintPreview_tbrBottom_Buttons_ToolTipText_LastPage)
    
    tbrTop.Buttons("FirstPage").ToolTipText = tbrBottom.Buttons("FirstPage").ToolTipText
    tbrTop.Buttons("PreviousPage").ToolTipText = tbrBottom.Buttons("PreviousPage").ToolTipText
    tbrTop.Buttons("NextPage").ToolTipText = tbrBottom.Buttons("NextPage").ToolTipText
    tbrTop.Buttons("LastPage").ToolTipText = tbrBottom.Buttons("LastPage").ToolTipText
    
End Sub

Public Property Let MinScalePercent(nValue As Long)
    If nValue <> mMinScalePercent Then
        mMinScalePercent = nValue
    End If
End Property

Public Property Let MaxScalePercent(nValue As Long)
    If nValue <> mMaxScalePercent Then
        mMaxScalePercent = nValue
    End If
End Property

Public Property Let UseAltScaleIcons(nValue As Boolean)
    If nValue Then
        tbrTop.Buttons("DecreaseScale").UseAltPic = True
        tbrTop.Buttons("IncreaseScale").UseAltPic = True
    Else
        tbrTop.Buttons("DecreaseScale").UseAltPic = False
        tbrTop.Buttons("IncreaseScale").UseAltPic = False
    End If
End Property

Public Property Let ToolBarIconsSize(nValue As Long)
    mUserCanChangeIconsSize = False
    Select Case nValue
        Case vxPPTIAppDefault
            mtbrTopIconsSize = gToolbarsDefaultIconsSize
        Case vxPPTIconsSmall
            mtbrTopIconsSize = vxIconsSmall
        Case vxPPTIconsMedium
            mtbrTopIconsSize = vxIconsMedium
        Case vxPPTIconsBig
            mtbrTopIconsSize = vxIconsBig
        Case vxPPTIconsAuto
'            SetIconsAutoSize
            mUserCanChangeIconsSize = True
        Case vxPPTIconsSmallAndUserCanChange
            mtbrTopIconsSize = vxIconsSmall
            mUserCanChangeIconsSize = True
            If Not tmrFirstDraw.Enabled Then
                If CLng(GetSetting(AppNameForRegistry, "Preferences", "PrintPreview_ToolBarIconsSize", efnISAuto)) = efnISAuto Then
                    SaveSetting AppNameForRegistry, "Preferences", "PrintPreview_ToolBarIconsSize", CStr(efnISSmall)
                End If
            End If
        Case vxPPTIconsMediumAndUserCanChange
            mtbrTopIconsSize = vxIconsMedium
            mUserCanChangeIconsSize = True
            If Not tmrFirstDraw.Enabled Then
                If CLng(GetSetting(AppNameForRegistry, "Preferences", "PrintPreview_ToolBarIconsSize", efnISAuto)) = efnISAuto Then
                    SaveSetting AppNameForRegistry, "Preferences", "PrintPreview_ToolBarIconsSize", CStr(efnISMedium)
                End If
            End If
        Case vxPPTIconsBigAndUserCanChange
            mtbrTopIconsSize = vxIconsBig
            mUserCanChangeIconsSize = True
            If Not tmrFirstDraw.Enabled Then
                If CLng(GetSetting(AppNameForRegistry, "Preferences", "PrintPreview_ToolBarIconsSize", efnISAuto)) = efnISAuto Then
                    SaveSetting AppNameForRegistry, "Preferences", "PrintPreview_ToolBarIconsSize", CStr(efnISBig)
                End If
            End If
    End Select
    
    If Not mLoading Then
        tbrTop.IconsSize = mtbrTopIconsSize
        If mUserCanChangeIconsSize Then
            Set cboAlternativePopupMenu.Font = GetSystemFont(vxMenuFont)
            SetToolBarIconsSize CLng(GetSetting(AppNameForRegistry, "Preferences", "PrintPreview_ToolBarIconsSize", efnISAuto))
        End If
    End If
End Property


Public Property Let UseOneToolBar(ByVal nValue As Boolean)
    Dim iStr As String
    
    If mLoading Then
        iStr = GetSetting(AppNameForRegistry, "Preferences", "PrintPreview_ShowBottomToolBar", "")
        If iStr <> "" Then
            nValue = CBool(Val(iStr))
        End If
    End If
    
    If nValue <> mUseOneToolBar Then
        mUseOneToolBar = nValue
        If Not mLoading Then
            Form_Resize
        End If
    End If
End Property

Public Property Get UseOneToolBar() As Boolean
    UseOneToolBar = mUseOneToolBar
End Property

Public Property Let Loading(nValue As Boolean)
    mLoading = nValue
End Property

Private Sub PositionPicStatus()
    Dim iTop As Long
    Dim iLng As Long
    Dim iLng2 As Long
    
    If Not mUseOneToolBar Then
        If HScroll1.Visible Then
            iTop = HScroll1.Top - picStatus.Height
        Else
            iTop = tbrBottom.Top - picStatus.Height
        End If
    Else
        If HScroll1.Visible Then
            iTop = HScroll1.Top - picStatus.Height
        Else
            iTop = Me.ScaleHeight - picStatus.Height
        End If
    End If
    
    picStatus.Move 200, iTop
    iLng = picPagesContainer.Left + picPage(0).Left - picStatus.Left - 2 * Screen.TwipsPerPixelX
    iLng2 = picStatus.TextWidth(lblStatus.Caption & " ")
    If iLng > 3648 Then
        iLng = 3648
    End If
    If iLng < iLng2 Then
        iLng = iLng2
    End If
    picStatus.Width = iLng
End Sub

Private Sub PositionForm()
    Dim iHwndActiveWindow As Long
    Dim iMonitor As Long
    Dim iMi As MONITORINFO
    
    iHwndActiveWindow = GetActiveWindowHwnd
    If iHwndActiveWindow <> 0 Then
        iMonitor = MonitorFromWindow(iHwndActiveWindow, MONITOR_DEFAULTTOPRIMARY)
    End If
    If iMonitor <> 0 Then
        iMi.cbSize = Len(iMi)
        GetMonitorInfo iMonitor, iMi
        If (iMi.rcWork.Bottom - iMi.rcWork.Top) <> 0 Then
            Me.Move iMi.rcWork.Left * Screen.TwipsPerPixelX, iMi.rcWork.Top * Screen.TwipsPerPixelY, (iMi.rcWork.Right - iMi.rcWork.Left) * Screen.TwipsPerPixelX, (iMi.rcWork.Bottom - iMi.rcWork.Top + 1) * Screen.TwipsPerPixelY
        Else
            Me.Move 0, 0, Screen.Width, ScreenUsableHeight
        End If
    Else
        Me.Move 0, 0, Screen.Width, ScreenUsableHeight
    End If
End Sub
