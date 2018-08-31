VERSION 5.00
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#1.7#0"; "vbExtra1.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Number of fixed thelephone lines by country"
   ClientHeight    =   5964
   ClientLeft      =   2268
   ClientTop       =   2016
   ClientWidth     =   5712
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5964
   ScaleWidth      =   5712
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   408
      Left            =   3564
      TabIndex        =   2
      Top             =   5400
      Width           =   1524
   End
   Begin vbExtra.PopupList pplYear 
      Height          =   372
      Left            =   2520
      TabIndex        =   1
      Top             =   144
      Width           =   3036
      _ExtentX        =   5355
      _ExtentY        =   656
      Text            =   "pplYear"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4440
      Left            =   144
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   648
      Width           =   5412
   End
   Begin VB.Label Label1 
      Caption         =   "The PopupList control is similar to a ComboBox with its Style set to DropDownList, but with a different look"
      ForeColor       =   &H00FF0000&
      Height          =   624
      Left            =   180
      TabIndex        =   3
      Top             =   5256
      Width           =   3036
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim y As Long
    
    For y = 2000 To 2014
        pplYear.AddItem "data year " & y
        pplYear.ItemData(pplYear.NewIndex) = y
    Next y
    pplYear.ListIndex = pplYear.ListCount - 1
End Sub

Private Sub pplYear_Click()
    LoadData
End Sub

Private Sub LoadData()
    Dim iYear As String
    Dim iDataStr As String
    Dim iLines() As String
    Dim iParts() As String
    Dim iLines2() As String
    Dim l As Long
    Dim l2 As Long
    Dim s As String
    
    iYear = pplYear.ItemData(pplYear.ListIndex)

    iDataStr = LoadTextFile(App.Path & "\UNdata_Fixed_Telephone_Lines.txt")
    iLines = Split(iDataStr, vbCrLf)
    ReDim iLines2(UBound(iLines))
    
    Set Me.Font = txtData.Font
    
    l2 = 0
    For l = 0 To UBound(iLines)
        If iLines(l) <> "" Then
            iParts = Split(iLines(l), "|")
            If iParts(1) = iYear Then
                s = iParts(0) & ": "
                Do Until Me.TextWidth(s) > 2500
                    s = s & " "
                Loop
                iLines2(l2) = s & vbTab & iParts(2)
                l2 = l2 + 1
            End If
        End If
    Next
    ReDim Preserve iLines2(l2)
    txtData.Text = Join(iLines2, vbCrLf)
End Sub

