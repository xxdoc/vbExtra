VERSION 5.00
Begin VB.Form frmFormPersistanceTest 
   Caption         =   "Form persistance test"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "To have 'persistance' means that the form will remember the size and position the next time"
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
Attribute VB_Name = "frmFormPersistanceTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    PersistForm Me, Forms
End Sub
