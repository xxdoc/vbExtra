VERSION 5.00
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.0#0"; "vbExtra2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3996
   ClientLeft      =   1740
   ClientTop       =   2640
   ClientWidth     =   4164
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
   ScaleHeight     =   3996
   ScaleWidth      =   4164
   Begin VB.CommandButton Command2 
      Caption         =   "Minimize to the system tray"
      Height          =   372
      Left            =   324
      TabIndex        =   3
      Top             =   1980
      Width           =   2300
   End
   Begin vbExtra.TrayIcon TrayIcon1 
      Left            =   3168
      Top             =   2160
      _ExtentX        =   720
      _ExtentY        =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simulate a hang"
      Height          =   372
      Left            =   324
      TabIndex        =   2
      Top             =   2520
      Width           =   2300
   End
   Begin VB.Label Label4 
      Caption         =   $"Form1.frx":0000
      Height          =   876
      Left            =   324
      TabIndex        =   5
      Top             =   3060
      Width           =   3528
   End
   Begin VB.Label Label3 
      Caption         =   "The sample code is in Sub Main of Module1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   324
      TabIndex        =   4
      Top             =   1260
      Width           =   3388
   End
   Begin VB.Label Label2 
      Caption         =   "Because ActivatePrevInstance function only has effect when the program is compiled."
      Height          =   480
      Left            =   324
      TabIndex        =   1
      Top             =   756
      Width           =   3388
   End
   Begin VB.Label Label1 
      Caption         =   "This sample needs to run compiled"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   324
      TabIndex        =   0
      Top             =   288
      Width           =   3388
   End
   Begin VB.Menu mnuTayPopup 
      Caption         =   "mnuTayPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore window"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close program"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' All the code here has nothing to do with ActivatePrevInstance function, it has been added to provide some functionality and to check different situations and how ActivatePrevInstance behaves
' All the neccesary code is in Sub main of Module1

Private Sub Command1_Click()
    
    Do
    Loop
    
End Sub

Private Sub Command2_Click()
    TrayIcon1.Create App.Title, Me.Icon
    Me.Visible = False
    TrayIcon1.BalloonTip "The program is here", vxBTSInfo, App.Title
End Sub

Private Sub Form_Load()
    PersistForm Me, Forms
    Randomize
    Me.Caption = "ID: " & App.ThreadID
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuRestore_Click()
    Me.Visible = True
    TrayIcon1.Remove
End Sub

Private Sub TrayIcon1_TrayClick(Button As vbExtra.vbExTrayIconMouseEventConstants)
    TrayIcon1.PopupMenu mnuTayPopup
End Sub
