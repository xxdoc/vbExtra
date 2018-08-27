VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4740
   ClientLeft      =   7308
   ClientTop       =   2028
   ClientWidth     =   4836
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4740
   ScaleWidth      =   4836
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Minimizing this form will minimize all forms (in this case both forms)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   876
      Left            =   792
      TabIndex        =   0
      Top             =   864
      Width           =   3000
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        MinimizeApp
    End If
End Sub
