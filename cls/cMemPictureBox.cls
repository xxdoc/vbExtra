VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMemPictureBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This Class repricates the PictureBox interface and functions.
' The purpose is to be able to use existing code that need a PictureBox, but to use it within a Class module or in a place where there are no PictureBoxes.

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Type PicBmp
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

' Declarations
Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle&, iPic As IPicture) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDc As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GetStretchBltMode Lib "gdi32" (ByVal hDc As Long) As Long

Private Const HALFTONE = 4

' Variables
Private mDC As Long
Private mBitMap As Long
Private mWidth As Long
Private mHeight As Long
Private mPicture As StdPicture
Private mPrevObject As Long
Private mAutoSize As Boolean
Private mUpdateBitmapDimensionsPending As Boolean
Private mPaintThePicturePropertyPending As Boolean
Private mThereArePendingTasks As Boolean


' Class Events Procedures
Private Sub Class_Initialize()
    mAutoSize = True
    CreateObjects
End Sub

Private Sub Class_Terminate()
    DeleteObjects
End Sub


' Private Procedures
Private Sub CreateObjects()
    mDC = CreateCompatibleDC(0)
    
    If mWidth = 0 Then
        mWidth = 1
    End If
    If mHeight = 0 Then
        mHeight = 1
    End If
    mUpdateBitmapDimensionsPending = True
    mThereArePendingTasks = True
End Sub

Private Sub DeleteObjects()
    SelectObject mDC, mPrevObject
    DeleteObject mBitMap
    DeleteDC mDC
End Sub

Private Sub UpdateBitmapDimensions()
    Dim iBitMapAnt As Long
    Dim iDWHwnd As Long
    Dim iDWDC As Long
    
    iBitMapAnt = mBitMap
    iDWHwnd = GetDesktopWindow
    iDWDC = GetDC(iDWHwnd)
    mBitMap = CreateCompatibleBitmap(iDWDC, mWidth, mHeight)
    ReleaseDC iDWHwnd, iDWDC
    'DeleteDC iDWDC
    If iBitMapAnt = 0 Then
        mPrevObject = SelectObject(mDC, mBitMap)
    Else
        DeleteObject SelectObject(mDC, mBitMap)
    End If
End Sub

Private Sub PaintThePictureProperty()
    If Not mPicture Is Nothing Then
        Me.PaintPicture mPicture, 0, 0
    End If
End Sub

Private Sub AdjustSize()
    Dim iPicInfo As BITMAP
    
    If Not mPicture Is Nothing Then
        If mPicture.Handle <> 0 Then
            GetObjectAPI mPicture.Handle, Len(iPicInfo), iPicInfo
            mWidth = iPicInfo.bmWidth
            Height = iPicInfo.bmHeight
        End If
    End If
End Sub


' Public Methods
Public Sub PaintPicture(nPicture As StdPicture, ByVal X1 As Long, _
    ByVal Y1 As Long, Optional ByVal Width1 As Long, Optional ByVal _
    Height1 As Long, Optional ByVal X2 As Long, Optional ByVal Y2 As _
    Long, Optional ByVal Width2 As Long, Optional ByVal Height2 As _
    Long, Optional ByVal Opcode As Long, Optional nUseHalftone As Boolean = True)
    
    Dim iPicInfo As BITMAP
    Dim iPicWidth As Long
    Dim iPicHeight As Long
    Dim iTempDC As Long
    Dim iOldObject As Long
    
    Dim iPicBits() As Byte
    Dim iBytesPerLine As Long
    Dim iTempBmp As Long
    
    DoPendingTasks
    If Not nPicture Is Nothing Then
        If nPicture.Handle <> 0 Then
            GetObjectAPI nPicture.Handle, Len(iPicInfo), iPicInfo
            If iPicInfo.bmWidth = 0 Then Exit Sub
            iPicWidth = iPicInfo.bmWidth
            iPicHeight = iPicInfo.bmHeight
            
            iTempDC = CreateCompatibleDC(mDC)
            iOldObject = SelectObject(iTempDC, nPicture.Handle)
            
            If iOldObject = 0 Then
                iBytesPerLine = (iPicInfo.bmWidth * 3 + 3) And &HFFFFFFFC
                ReDim iPicBits(1 To iBytesPerLine * iPicInfo.bmHeight * 3) As Byte
                GetBitmapBits nPicture.Handle, UBound(iPicBits), iPicBits(1)
                iTempBmp = CreateCompatibleBitmap(mDC, iPicInfo.bmWidth, iPicInfo.bmHeight)
                SetBitmapBits iTempBmp, UBound(iPicBits), iPicBits(1)
                
                iOldObject = SelectObject(iTempDC, iTempBmp)
            End If
            
            If Width2 = 0 Then
                Width2 = iPicWidth - X2
            End If
            If Height2 = 0 Then
                Height2 = iPicHeight - Y2
            End If
            If (X1 + Width2) > mWidth Then
                Width2 = mWidth - X1
            End If
            If (Y1 + Height2) > mHeight Then
                Height2 = mHeight - Y1
            End If
            If Opcode = 0 Then
                Opcode = vbSrcCopy
            End If
            If Width1 = 0 Then
                Width1 = Width2
            End If
            If Height1 = 0 Then
                Height1 = Height2
            End If
            If (Width1 = Width2) And (Height1 = Height2) Then
                BitBlt mDC, X1, Y1, Width2, Height2, iTempDC, X2, Y2, Opcode
            Else
                
                If nUseHalftone Then
                    If GetStretchBltMode(mDC) <> HALFTONE Then
                        SetStretchBltMode mDC, HALFTONE
                    End If
                End If
                
                StretchBlt mDC, X1, Y1, Width1, Height1, iTempDC, X2, Y2, _
                    Width2, Height2, Opcode
            End If
        
            SelectObject iTempDC, iOldObject
            If iTempBmp <> 0 Then
                DeleteObject iTempBmp
            End If
            DeleteDC iTempDC
        End If
    End If
End Sub

Public Sub Cls()
    If mBitMap <> 0 Then
        SelectObject mDC, mPrevObject
        DeleteObject mBitMap
    End If
    mUpdateBitmapDimensionsPending = True
    mPaintThePicturePropertyPending = True
    mThereArePendingTasks = True
End Sub

Public Function Point(x As Long, y As Long) As Long
    Point = GetPixel(mDC, x, y)
End Function


' Properties

Public Property Let Picture(nPicture As StdPicture)
    Set Picture = nPicture
End Property

Public Property Set Picture(nPicture As StdPicture)
    Set mPicture = nPicture
    If mAutoSize Then
        AdjustSize
    End If
    Cls
    mPaintThePicturePropertyPending = True
End Property

Public Property Get Picture() As StdPicture
    Set Picture = mPicture
End Property


Public Property Get Image() As StdPicture
    Dim iPicBmp As PicBmp
    Dim iPic As IPicture
    Dim IID_IDispatch As GUID
    Dim iBmp As Long
    Dim iTempDC As Long
    Dim iOldObject As Long
    
    DoPendingTasks
    
    iTempDC = CreateCompatibleDC(mDC)
    iBmp = CreateCompatibleBitmap(mDC, mWidth, mHeight)
    iOldObject = SelectObject(iTempDC, iBmp)
    
    BitBlt iTempDC, 0, 0, mWidth, mHeight, mDC, 0, 0, vbSrcCopy
    SelectObject iTempDC, iOldObject
    DeleteDC iTempDC
    
    'fill in OLE IDispatch Interface ID
    With IID_IDispatch
       .Data1 = &H20400
       .Data4(0) = &HC0
       .Data4(7) = &H46
    End With
    
    'fill Pic with necessary parts
    With iPicBmp
       .Size = Len(iPicBmp)         'Length of structure
       .Type = vbPicTypeBitmap  'Type of Picture (bitmap)
       .hBmp = iBmp          'Handle to bitmap
       .hPal = 0&               'Handle to palette (may be null)
     End With
    
    'create OLE Picture object
    Call OleCreatePictureIndirect(iPicBmp, IID_IDispatch, 1, iPic)
    'return the new Picture object
    Set Image = iPic
End Property


Public Property Get hDc() As Long
    If mThereArePendingTasks Then
        DoPendingTasks
    End If
    hDc = mDC
End Property

Public Property Get AutoRedraw() As Boolean
    AutoRedraw = True
End Property


Public Property Get Width() As Long
    Width = mWidth
End Property

Public Property Let Width(nWidth As Long)
    mWidth = nWidth
    mUpdateBitmapDimensionsPending = True
    mPaintThePicturePropertyPending = True
    mThereArePendingTasks = True
End Property


Public Property Get Height() As Long
    Height = mHeight
End Property

Public Property Let Height(nHeight As Long)
    mHeight = nHeight
    mUpdateBitmapDimensionsPending = True
    mPaintThePicturePropertyPending = True
    mThereArePendingTasks = True
End Property


Public Property Get AutoSize() As Boolean
    AutoSize = mAutoSize
End Property

Public Property Let AutoSize(ByVal nValue As Boolean)
    mAutoSize = nValue
    AdjustSize
End Property


Public Property Get ScaleWidth() As Long
    ScaleWidth = mWidth
End Property

Public Property Let ScaleWidth(ByVal nWidth As Long)
    Width = nWidth
End Property


Public Property Get ScaleHeight() As Long
    ScaleHeight = mHeight
End Property

Public Property Let ScaleHeight(ByVal nHeight As Long)
    Height = nHeight
End Property


' Properties provided just for interface compatibility
Public Property Get ScaleLeft() As Long
    ScaleLeft = 0
End Property

Public Property Let ScaleLeft(nValue As Long)
    '
End Property


Public Property Get ScaleTop() As Long
    ScaleTop = 0
End Property

Public Property Let ScaleTop(nValue As Long)
    '
End Property


Public Property Let AutoRedraw(ByVal vNewValue As Boolean)
    '
End Property


Public Property Get ScaleMode() As Integer
    ScaleMode = vbPixels
End Property

Public Property Let ScaleMode(ByVal vNewValue As Integer)
    '
End Property


Public Property Get BorderStyle() As Integer
    BorderStyle = 0
End Property

Public Property Let BorderStyle(ByVal vNewValue As Integer)
    '
End Property


Public Property Get BackColor() As Long
    BackColor = vbBlack
End Property

Public Property Let BackColor(ByVal vNewValue As Long)
    '
End Property

Private Sub DoPendingTasks()
    If mThereArePendingTasks Then
        If mUpdateBitmapDimensionsPending Then
            mUpdateBitmapDimensionsPending = False
            UpdateBitmapDimensions
            If Not mPaintThePicturePropertyPending Then
                mThereArePendingTasks = False
            End If
        End If
        If mPaintThePicturePropertyPending Then
            mPaintThePicturePropertyPending = False
            PaintThePictureProperty
            If Not mUpdateBitmapDimensionsPending Then
                mThereArePendingTasks = False
            End If
        End If
    End If
End Sub

Public Sub Refresh()
    '
End Sub
