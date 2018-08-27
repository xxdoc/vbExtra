VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3132
   ClientLeft      =   1704
   ClientTop       =   2280
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
   Begin VB.CommandButton Command1 
      Caption         =   "Show Form2 modally"
      Height          =   552
      Left            =   396
      TabIndex        =   2
      Top             =   2268
      Width           =   1704
   End
   Begin VB.Label Label2 
      Caption         =   "(That's not possible with VB's .Show vbModal)"
      Height          =   408
      Left            =   1080
      TabIndex        =   1
      Top             =   1584
      Width           =   3648
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "The forms are shown 'modally' but still can show a non modal form. "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
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

Private Sub Command1_Click()
    ShowModal Form2
End Sub
