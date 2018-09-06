VERSION 5.00
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.0#0"; "vbExtra2.ocx"
Begin VB.Form frmTestHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Number of fixed thelephone lines by country"
   ClientHeight    =   5820
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
   ScaleHeight     =   5820
   ScaleWidth      =   5712
   Begin vbExtra.History History1 
      Height          =   348
      Left            =   144
      TabIndex        =   4
      Top             =   216
      Width           =   552
      _ExtentX        =   974
      _ExtentY        =   614
      Enabled         =   0   'False
      Context         =   "frmTestHistory_History1"
      BoundControlName=   "txtCountry"
      BoundProperty   =   "Text"
   End
   Begin VB.Timer tmrList 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4896
      Top             =   216
   End
   Begin VB.TextBox txtCountry 
      Height          =   372
      Left            =   3024
      TabIndex        =   1
      Top             =   216
      Width           =   1596
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   408
      Left            =   3528
      TabIndex        =   3
      Top             =   5256
      Width           =   1524
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
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   648
      Width           =   5412
   End
   Begin VB.Label Label1 
      Caption         =   "Show countries starting with:"
      Height          =   264
      Left            =   900
      TabIndex        =   0
      Top             =   288
      Width           =   2100
   End
End
Attribute VB_Name = "frmTestHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub LoadData()
    Dim iDataStr As String
    Dim iLines() As String
    Dim iParts() As String
    Dim iLines2() As String
    Dim l As Long
    Dim l2 As Long
    Dim s As String
    Dim iPut As Boolean
    
    iDataStr = LoadTextFile(App.Path & "\UNdata_Fixed_Telephone_Lines.txt")
    iLines = Split(iDataStr, vbCrLf)
    ReDim iLines2(UBound(iLines))
    
    Set Me.Font = txtData.Font
    
    iLines(0) = "Country|Year|Lines"
    l2 = 0
    For l = 0 To UBound(iLines)
        If iLines(l) <> "" Then
            iParts = Split(iLines(l), "|")
            iPut = LCase(iParts(0)) Like LCase(txtCountry.Text) & "*"
            If l = 0 Then iPut = True
            If iPut Then
                s = iParts(0) & ": "
                Do Until Me.TextWidth(s) > 2500
                    s = s & " "
                Loop
                iLines2(l2) = s & vbTab & iParts(1) & vbTab & iParts(2)
                l2 = l2 + 1
            End If
        End If
    Next
    ReDim Preserve iLines2(l2)
    txtData.Text = Join(iLines2, vbCrLf)
End Sub

Private Sub History1_Click(nText As String)
    tmrList_Timer
End Sub

Private Sub tmrList_Timer()
    tmrList.Enabled = False
    LoadData
End Sub

Private Sub txtCountry_Change()
    tmrList.Enabled = False
    tmrList.Enabled = True
End Sub

Private Sub txtCountry_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        tmrList_Timer
    End If
End Sub
