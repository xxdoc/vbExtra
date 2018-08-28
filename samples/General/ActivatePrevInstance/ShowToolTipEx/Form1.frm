VERSION 5.00
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#1.3#0"; "vbExtra1.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4464
   ClientLeft      =   2736
   ClientTop       =   2352
   ClientWidth     =   5568
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
   ScaleHeight     =   4464
   ScaleWidth      =   5568
   Begin vbExtra.ButtonEx cmdHelpDocumentName 
      Height          =   228
      Left            =   2700
      TabIndex        =   2
      Top             =   2088
      Width           =   228
      _ExtentX        =   402
      _ExtentY        =   402
      ButtonStyle     =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16354918
      Caption         =   "?"
      ForeColor       =   16777215
      PictureAlign    =   4
   End
   Begin VB.Label Label4 
      Caption         =   "The manifest file is already addeed to the project, then please compile in order to test it"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   432
      TabIndex        =   4
      Top             =   828
      Width           =   4728
   End
   Begin VB.Label Label3 
      Caption         =   "Note: For the Balloon tooltips to work Visual Styles must be manifested and the program must be compiled"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   516
      Left            =   432
      TabIndex        =   3
      Top             =   324
      Width           =   4728
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Click on the '?' button"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   756
      TabIndex        =   1
      Top             =   1872
      Width           =   1812
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Position the mouse over here"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   756
      TabIndex        =   0
      Top             =   3348
      Width           =   1812
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHelpDocumentName_Click()
    ShowToolTipEx "ToolTipText that will stay displayed until the user closes it (or the mouse moves over another TooTipEx zone)", "Close me", vxTTBalloon, True, vxTTIconInfo
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowToolTipEx "ToolTipText that will auto hide when mouse goes away", "ToolTip", vxTTBalloon, , vxTTIconInfo, 2
End Sub
