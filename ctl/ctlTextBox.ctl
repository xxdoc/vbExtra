VERSION 5.00
Begin VB.UserControl ctlTextBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LockControls    =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox txtMultiLine 
      Height          =   1005
      Left            =   300
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1500
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtNoMultiLine 
      Height          =   1005
      Left            =   300
      TabIndex        =   0
      Top             =   330
      Width           =   2295
   End
End
Attribute VB_Name = "ctlTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mMultiLine As Boolean

Private WithEvents mTextBox As TextBox
Attribute mTextBox.VB_VarHelpID = -1

Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Devuelve o establece un valor que determina si un control puede aceptar múltiples líneas de texto."
    MultiLine = mMultiLine
End Property

Public Property Let MultiLine(ByVal New_MultiLine As Boolean)
    If (New_MultiLine <> mMultiLine) Or mTextBox Is Nothing Then
        mMultiLine = New_MultiLine
        PropertyChanged "MultiLine"
        If mMultiLine Then
            txtMultiLine.Visible = True
            txtNoMultiLine.Visible = False
            Set mTextBox = txtMultiLine
        Else
            txtMultiLine.Visible = False
            txtNoMultiLine.Visible = True
            Set mTextBox = txtNoMultiLine
        End If
    End If
End Property

Private Sub UserControl_InitProperties()
    MultiLine = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    MultiLine = PropBag.ReadProperty("MultiLine", False)
End Sub

Private Sub UserControl_Resize()
    txtNoMultiLine.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    txtMultiLine.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("MultiLine", mMultiLine, False)
End Sub

Public Property Get TextBoxControl() As Object
    Set TextBoxControl = mTextBox
End Property
