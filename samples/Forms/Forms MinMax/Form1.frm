VERSION 5.00
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#1.3#0"; "vbExtra1.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3132
   ClientLeft      =   5472
   ClientTop       =   4740
   ClientWidth     =   5376
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
   ScaleHeight     =   3132
   ScaleWidth      =   5376
   Begin vbExtra.SizeGrip SizeGrip1 
      Height          =   228
      Left            =   5148
      Top             =   2904
      Width           =   228
      _ExtentX        =   402
      _ExtentY        =   402
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Set the minimun and/or maximun size that a form can have"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1632
      Left            =   1080
      TabIndex        =   0
      Top             =   432
      Width           =   3036
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SetMinMax Me, 5000, 2000, 6000, 6000
End Sub
