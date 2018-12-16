VERSION 5.00
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "vbExtra2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   2892
   ClientTop       =   2160
   ClientWidth     =   6228
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   6228
   Begin vbExtra.SizeGrip SizeGrip1 
      Height          =   228
      Left            =   6000
      Top             =   4872
      Width           =   228
      _ExtentX        =   402
      _ExtentY        =   402
   End
   Begin vbExtra.SSTabEx SSTabEx1 
      Height          =   3432
      Left            =   468
      TabIndex        =   0
      Top             =   864
      Width           =   5376
      _ExtentX        =   9483
      _ExtentY        =   6054
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      TabHeight       =   617
      Themed          =   -1  'True
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   2
      Tab(0).Control(0)=   "Text3"
      Tab(0).Control(1)=   "Text2"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   2
      Tab(1).Control(0)=   "Command2"
      Tab(1).Control(1)=   "Command1"
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   1
      Tab(2).Control(0)=   "ScrollableContainer1"
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Left            =   612
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   1548
         Width           =   1200
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Left            =   612
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   936
         Width           =   1200
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Left            =   -74424
         TabIndex        =   5
         Top             =   1656
         Width           =   1344
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   444
         Left            =   -74424
         TabIndex        =   4
         Top             =   1008
         Width           =   1344
      End
      Begin vbExtra.ScrollableContainer ScrollableContainer1 
         Height          =   2712
         Left            =   -74856
         TabIndex        =   1
         Top             =   540
         Width           =   4944
         _ExtentX        =   8721
         _ExtentY        =   4784
         SavedVScrollMax =   205
         SavedVirtualHeight=   2460
         SavedHScrollMax =   391
         SavedVirtualWidth=   4692
         BottomFreeSpace =   0
         BorderStyle     =   0
         BorderColor     =   -2147483638
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   0
            Left            =   2520
            TabIndex        =   3
            Text            =   "Text1(0)"
            Top             =   216
            Width           =   1452
         End
         Begin VB.Label Label1 
            Caption         =   "Label1(0)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   336
            Index           =   0
            Left            =   324
            TabIndex        =   2
            Top             =   252
            Width           =   1956
         End
      End
   End
   Begin VB.Label Label2 
      Caption         =   "This sample also uses the SizeGrip control and the SSTabEx control. The ScrollableContainer is in Tab2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   516
      Left            =   216
      TabIndex        =   8
      Top             =   180
      Width           =   5736
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim c As Long
    
    For c = 1 To 10
        Load Label1(c)
        Load Text1(c)
        Label1(c).Caption = "Label1(" & c & ")"
        Text1(c).Text = "Text1(" & c & ")"
        Label1(c).Top = Label1(c - 1).Top + 700
        Text1(c).Top = Text1(c - 1).Top + 700
        Label1(c).Visible = True
        Text1(c).Visible = True
    Next c
    
    ScrollableContainer1.BottomFreeSpace = 200
End Sub

Private Sub Form_Resize()
    SSTabEx1.Move 140, Label2.Top + Label2.Height + 140, Me.ScaleWidth - 280, Me.ScaleHeight - 280 - Label2.Top - Label2.Height
End Sub

Private Sub SSTabEx1_TabBodyResize()
    ScrollableContainer1.Move SSTabEx1.TabBodyLeft + 60, SSTabEx1.TabBodyTop + 60, SSTabEx1.TabBodyWidth - 120, SSTabEx1.TabBodyHeight - 120
End Sub
