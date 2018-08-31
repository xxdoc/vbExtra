Attribute VB_Name = "MPicStretch"
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

' Representation of 32-bit RGBA color
Private Type RGBQUAD
    rgbRed As Byte
    rgbGreen As Byte
    rgbBlue As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
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

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

' VB's array header structure
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(1) As SAFEARRAYBOUND
End Type

Private Type PicBmp
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function ColorRGBToHLS Lib "shlwapi.dll" (ByVal clrRGB As Long, pwHue As Long, pwLuminance As Long, pwSaturation As Long) As Long
Private Declare Function ColorHLSToRGB Lib "shlwapi.dll" (ByVal wHue As Long, ByVal wLuminance As Long, ByVal wSaturation As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle&, iPic As IPicture) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC&, pBitmapInfo As BITMAPINFO, ByVal un&, lplpVoid&, ByVal Handle&, ByVal dw&) As Long
'Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

'Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC&) As Long
'Declare Function DeleteObject Lib "gdi32" (ByVal hObject&) As Long
'Declare Function SelectObject Lib "gdi32" (ByVal hDC&, ByVal hObject&) As Long
'Declare Function DeleteDC Lib "gdi32" (ByVal hDC&) As Long
'Declare Function BitBlt Lib "gdi32" (ByVal hDestDC&, ByVal X&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&) As Long

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy&)

Private Const pi = 3.14159265358979

Public Function ResampleBMP(nPicSrc As Picture, nWidth As Long, nHeight As Long) As StdPicture
    Dim bDibD() As Byte
    Dim bDibS() As Byte
    Dim bDibT() As Byte
    Dim tSAD As SAFEARRAY2D
    Dim tSAS As SAFEARRAY2D
    Dim tBMD As BITMAP
    Dim tBMS As BITMAP
    Dim hBitmapDst As StdPicture
    Dim tBM As BITMAP
    Dim sPic As StdPicture
    Dim CDC As Long
    Dim CDC1 As Long
    Dim iBicubic_B As Double
    Dim iBicubic_C As Double
    Dim iGaussianExtent As Long
    Dim iAuxDbl1 As Double
    Dim iAuxDbl2 As Double
    
    If nPicSrc = 0 Then Exit Function
    
    iBicubic_B = 0
    iBicubic_C = 0.32
    ' creates a new 24 bits bmp picture object that will be used as the destiny
    Set hBitmapDst = CreatePicture(nWidth, nHeight, 24)
    
    ' se ve si es origen es de 24 bits (la funcion solo funciona con bitmaps de 24 bits)
    GetObjectAPI nPicSrc, Len(tBM), tBM
    If tBM.bmBitsPixel <> 24 Then
        ' Create 24bpp empty (black) image
        Set sPic = CreatePicture(tBM.bmWidth, tBM.bmHeight, 24)
        CDC = CreateCompatibleDC(0) ' Temporary devices
        CDC1 = CreateCompatibleDC(0)
        DeleteObject SelectObject(CDC, nPicSrc) ' Source bitmap
        DeleteObject SelectObject(CDC1, sPic) ' Converted bitmap
        ' Copy between two different depths
        BitBlt CDC1, 0, 0, tBM.bmWidth, tBM.bmHeight, CDC, 0, 0, vbSrcCopy
        DeleteDC CDC: DeleteDC CDC1 ' Erase devices
        Set nPicSrc = sPic ' Set visible image
    End If
    
    
    GetObjectAPI hBitmapDst, Len(tBMD), tBMD
    With tSAD ' Array header structure
        .cbElements = 1
        .cDims = 2
        .Bounds(0).cElements = tBMD.bmHeight
        .Bounds(1).cElements = tBMD.bmWidthBytes ' (Width*3 aligned to 4)
        .pvData = tBMD.bmBits ' Pointer to array (bitmap)
    End With
    ' Associate header with array (no need of copying large blocks, direct access)
    CopyMemory ByVal VarPtrArray(bDibD), VarPtr(tSAD), 4
    GetObjectAPI nPicSrc, Len(tBMS), tBMS
    With tSAS
        .cbElements = 1
        .cDims = 2
        .Bounds(0).cElements = tBMS.bmHeight
        .Bounds(1).cElements = tBMS.bmWidthBytes
        .pvData = tBMS.bmBits
    End With
    CopyMemory ByVal VarPtrArray(bDibS), VarPtr(tSAS), 4
    
    iAuxDbl1 = tBM.bmWidth * tBM.bmHeight
    iAuxDbl2 = tBMD.bmWidth * tBMD.bmHeight
    
    If iAuxDbl2 >= iAuxDbl1 Then
        ' destination is bigger
        'metodo BicubicBCSpline
        ReDim bDibT(-3 To tBMS.bmWidth * 3 + 5, -1 To tBMS.bmHeight + 1)
        CopyImage24 bDibS, bDibT, (tBMS.bmWidth - 1) * 3 + 2
        Resample_BicubicBCSpline bDibD, tBMD.bmWidth, tBMD.bmHeight, bDibT, tBMS.bmWidth, tBMS.bmHeight, iBicubic_B, iBicubic_C
    Else
        ' Origin is bigger
        If (tBM.bmWidth / tBMD.bmWidth) > (tBM.bmHeight / tBMD.bmHeight) Then
            iGaussianExtent = (tBM.bmWidth / tBMD.bmWidth) + 1
        Else
            iGaussianExtent = (tBM.bmHeight / tBMD.bmHeight) + 1
        End If
        If iGaussianExtent < 2 Then iGaussianExtent = 2
        If iGaussianExtent > 20 Then iGaussianExtent = 20
        
        ReDim bDibT(-3 * (iGaussianExtent - 1) To (tBMS.bmWidth - 1 + iGaussianExtent) * 3 + 2, -(iGaussianExtent - 1) To tBMS.bmHeight - 1 + iGaussianExtent)
        CopyImage24 bDibS, bDibT, (tBMS.bmWidth - 1) * 3 + 2
        Resample_Gaussian bDibD, tBMD.bmWidth, tBMD.bmHeight, bDibT, tBMS.bmWidth, tBMS.bmHeight, iGaussianExtent
        
    End If
    CopyMemory ByVal VarPtrArray(bDibD), 0&, 4 ' Important under WinNT platform
    CopyMemory ByVal VarPtrArray(bDibS), 0&, 4 ' Free arrays
    
    Set ResampleBMP = hBitmapDst
End Function

Private Sub Resample_BicubicBCSpline(bDibDest() As Byte, ByVal dstWidth As Long, ByVal dstHeight As Long, bDibSource() As Byte, ByVal srcWidth As Long, ByVal srcHeight As Long, nBicubic_B As Double, nBicubic_C As Double)
    Dim x As Long
    Dim y As Long
    Dim X1 As Long
    Dim Y1 As Long
    Dim M As Long
    Dim N As Long
    Dim kX As Double
    Dim kY As Double
    Dim fX As Double
    Dim fY As Double
    Dim iR As Double
    Dim iG As Double
    Dim iB As Double
    Dim R1 As Double
    Dim R2 As Double
    
    kX = (srcWidth - 1) / (dstWidth - 1)
    kY = (srcHeight - 1) / (dstHeight - 1)
    For y = dstHeight - 1 To 0 Step -1
        fY = y * kY
        Y1 = Int(fY)
        fY = fY - Y1
        For x = 0 To dstWidth - 1
            fX = x * kX
            X1 = Int(fX)
            fX = fX - X1
            X1 = X1 * 3
            iR = 0: iG = 0: iB = 0
            For M = -1 To 2
                R1 = Cubic_BCSpline(M - fY, nBicubic_B, nBicubic_C)
                For N = -1 To 2
                    R2 = Cubic_BCSpline(fX - N, nBicubic_B, nBicubic_C)
                    iB = iB + bDibSource(X1 + N * 3, Y1 + M) * R1 * R2
                    iG = iG + bDibSource(X1 + N * 3 + 1, Y1 + M) * R1 * R2
                    iR = iR + bDibSource(X1 + N * 3 + 2, Y1 + M) * R1 * R2
                Next
            Next
            If iB < 0 Then iB = 0
            If iG < 0 Then iG = 0
            If iR < 0 Then iR = 0
            If iB > 255 Then iB = 255
            If iG > 255 Then iG = 255
            If iR > 255 Then iR = 255
            bDibDest(x * 3, y) = iB
            bDibDest(x * 3 + 1, y) = iG
            bDibDest(x * 3 + 2, y) = iR
        Next
    Next
End Sub

' Cubic BC-spline function
' Mitchell and Netravali derived a family of such cubic filters dependent on two variables: B, C
' Some of the values for B and C correspond to well-known filters,
' e.g., B=1 and C=0 corresponds to the cubic B-spline,
' and C=0 results in the family of Duff's tensioned B-splines.
' Setting B=0 and C=-a results in the family of the cardinal splines which were derived by Keys in 1981.
' Using Taylor series expansion they determined that, numerically, the filters for which B + 2 * C = 1 with 0 <= B <= 1
' are the most accurate within that class
' and that the reconstruction error for synthetic examples is proportional to the square of the sampling distance.
' Two new filters were proposed, the first with B=3/2 and C=1/3 suppresses post-aliasing but is unnecessarily blurring,
' the second with B=1/3 and C=1/3 turns out to be a satisfactory compromise between ringing, blurring, and anisotropy.
Private Function Cubic_BCSpline(ByVal x As Double, cubic_B As Double, cubic_c As Double) As Double
    x = Abs(x)
    If x < 1 Then
        Cubic_BCSpline = ((12 - 9 * cubic_B - 6 * cubic_c) * x * x * x + (-18 + 12 * cubic_B + 6 * cubic_c) * x * x + 6 - 2 * cubic_B) / 6
    ElseIf x < 2 Then
        Cubic_BCSpline = ((-cubic_B - 6 * cubic_c) * x * x * x + (6 * cubic_B + 30 * cubic_c) * x * x + (-12 * cubic_B - 48 * cubic_c) * x + (8 * cubic_B + 24 * cubic_c)) / 6
    End If
End Function

' Copies one memory bitmap into another with edge extending
Private Sub CopyImage24(InArray() As Byte, OutArray() As Byte, ByVal InUBound As Long, Optional ByVal DisablePad As Boolean)
    Dim i As Long
    Dim J As Long
    
    J = InUBound + 1
    ' Copy full content of input image into output
    For i = 0 To UBound(InArray, 2)
        CopyMemory OutArray(0, i), InArray(0, i), J
    Next
    ' Fill extended pad bytes with color of edges or not?
    If DisablePad Then Exit Sub
    ' Fill left and right edges
    For J = 0 To UBound(InArray, 2)
        For i = LBound(OutArray) To -3 Step 3
            OutArray(i, J) = OutArray(0, J) ' Blue
            OutArray(i + 1, J) = OutArray(1, J) ' Green
            OutArray(i + 2, J) = OutArray(2, J) ' Red
        Next
        For i = InUBound + 1 To UBound(OutArray) - 2 Step 3
            OutArray(i, J) = OutArray(InUBound - 2, J)
            OutArray(i + 1, J) = OutArray(InUBound - 1, J)
            OutArray(i + 2, J) = OutArray(InUBound, J)
        Next
    Next
    J = UBound(OutArray) - LBound(OutArray) + 1
    InUBound = LBound(OutArray)
    ' Fill top and bottom edges
    For i = LBound(OutArray, 2) To -1
        CopyMemory OutArray(InUBound, i), OutArray(InUBound, 0), J
    Next
    For i = UBound(InArray, 2) + 1 To UBound(OutArray, 2)
        CopyMemory OutArray(InUBound, i), OutArray(InUBound, UBound(InArray, 2)), J
    Next
End Sub

Private Function CreatePicture(ByVal nWidth&, ByVal nHeight&, ByVal nBPP&) As Picture
    Dim Pic As PicBmp, IID_IDispatch As GUID, BMI As BITMAPINFO
    With BMI.bmiHeader
        .biSize = Len(BMI.bmiHeader)
        .biWidth = nWidth
        .biHeight = nHeight
        .biPlanes = 1
        .biBitCount = nBPP
    End With
    Pic.hBmp = CreateDIBSection(0, BMI, 0, 0, 0, 0)
    With IID_IDispatch
        .Data1 = &H20400: .Data4(0) = &HC0: .Data4(7) = &H46
    End With
    Pic.Size = Len(Pic)
    Pic.Type = vbPicTypeBitmap
    OleCreatePictureIndirect Pic, IID_IDispatch, 1, CreatePicture
    If CreatePicture = 0 Then Set CreatePicture = Nothing
End Function

Sub Resample_Gaussian(bDibD() As Byte, ByVal dstWidth&, ByVal dstHeight&, bDibS() As Byte, ByVal srcWidth&, ByVal srcHeight&, nGaussianExtent As Long)
    Dim x&, y&, X1&, Y1&, M&, N&, kX#, kY#, fX#, fY#
    Dim iR#, iG#, iB#, R1#, R2#
    kX = (srcWidth - 1) / (dstWidth - 1)
    kY = (srcHeight - 1) / (dstHeight - 1)
    For y = dstHeight - 1 To 0 Step -1
        fY = y * kY
        Y1 = Int(fY)
        fY = fY - Y1
        For x = 0 To dstWidth - 1
            fX = x * kX
            X1 = Int(fX)
            fX = fX - X1
            X1 = X1 * 3
            iR = 0: iG = 0: iB = 0
            ' Uses various kernel size
            For M = -nGaussianExtent + 1 To nGaussianExtent
                R1 = Gaussian_Func(M - fY, nGaussianExtent)
                For N = -nGaussianExtent + 1 To nGaussianExtent
                    R2 = Gaussian_Func(fX - N, nGaussianExtent)
                    iB = iB + bDibS(X1 + N * 3, Y1 + M) * R1 * R2
                    iG = iG + bDibS(X1 + N * 3 + 1, Y1 + M) * R1 * R2
                    iR = iR + bDibS(X1 + N * 3 + 2, Y1 + M) * R1 * R2
                Next
            Next
            If iB < 0 Then iB = 0
            If iG < 0 Then iG = 0
            If iR < 0 Then iR = 0
            If iB > 255 Then iB = 255
            If iG > 255 Then iG = 255
            If iR > 255 Then iR = 255
            bDibD(x * 3, y) = iB
            bDibD(x * 3 + 1, y) = iG
            bDibD(x * 3 + 2, y) = iR
        Next
    Next
End Sub

' Gaussian function
' Could generate very blurry output.
' The wider the function is (higher standard deviation),
' more blurry the image is. This is useful for removing
' noise and aliasing but it won't preserve details.
Function Gaussian_Func(ByVal x#, nGaussianExtent As Long) As Double
    ' 0.398942280401433 = 1 / Sqr(2 * pi)
    Dim o#
    If Abs(x) < nGaussianExtent Then
        o = nGaussianExtent / pi ' standard deviation - could be changed
        Gaussian_Func = 0.398942280401433 / o * Exp(-x * x / (o * o * 2))
    End If
End Function

Public Function AdjustPictureWithHLS(nSourcePic As StdPicture, Optional HAddition As Long, Optional LAddition As Long, Optional SAddition As Long, Optional LFactor As Single = 1, Optional SFactor As Single = 1, Optional ColorToPreserve As Long = -1) As StdPicture
    Dim iBMP As BITMAP
    Dim iBMPiH As BITMAPINFOHEADER
    Dim iBits() As Byte
    
    Dim iTmpDC As Long
    Dim x As Long
    Dim iMax As Long
    Dim iColor As Long
    Dim iPicWidth As Long
    Dim iPicHeight As Long
    Dim iBitMap As Long
    Dim iOldObj As Long
    
    Dim iPicBmp As PicBmp
    Dim IID_IDispatch As GUID
    Dim iPic As IPicture
    Dim iBMPInfo As BITMAPINFO
    
    Dim R1 As Long
    Dim G1 As Long
    Dim B1 As Long
    Dim H1 As Long
    Dim L1 As Long
    Dim S1 As Long
    Dim iPreservePixel As Boolean
    Dim iPixelColor As Long
    Dim iLastH As Long
    
    If (HAddition < -120) Or (HAddition > 120) Then
        RaiseError 2195, "AdjustPictureWithHLS function", "HAddition must be between -120 and 120)"
        Exit Function
    End If
    If (LAddition < -240) Or (LAddition > 240) Then
        RaiseError 2195, "AdjustPictureWithHLS function", "LAddition and SAddition must be between -240 and 240)"
        Exit Function
    End If
    If (SAddition < -240) Or (SAddition > 240) Then
        RaiseError 2195, "AdjustPictureWithHLS function", "LAddition and SAddition must be between -240 and 240)"
        Exit Function
    End If
    If nSourcePic Is Nothing Then Exit Function
    If nSourcePic.Handle = 0 Then Exit Function
    
    GetObject nSourcePic.Handle, Len(iBMP), iBMP
    
    If iBMP.bmBitsPixel <> 24 Then
        RaiseError 2196, "AdjustPictureWithHLS function", "AdjustPictureWithHLS function only works with 24 bits bitmaps"
        Exit Function
    End If
    
    With iBMPiH
        .biSize = Len(iBMPiH) '40
        .biPlanes = 1
        .biBitCount = 24
        .biWidth = iBMP.bmWidth
        .biHeight = iBMP.bmHeight
        .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
        iPicWidth = .biWidth
        iPicHeight = .biHeight
    End With
    
    ReDim iBits(Len(iBMPiH) + iBMPiH.biSizeImage)
    
    iTmpDC = CreateCompatibleDC(0)
    
    GetDIBits iTmpDC, nSourcePic.Handle, 0, iBMP.bmHeight, iBits(0), iBMPiH, DIB_RGB_COLORS
    
    iMax = iBMPiH.biSizeImage - 1
    
    For x = 0 To iMax - 3 Step 3
        B1 = iBits(x)
        G1 = iBits(x + 1)
        R1 = iBits(x + 2)
        
        iPreservePixel = False
        If ColorToPreserve > -1 Then
            iPixelColor = RGB(R1, G1, B1)
            If iPixelColor = ColorToPreserve Then
                iPreservePixel = True
            End If
        End If
        
        If Not iPreservePixel Then
            ColorRGBToHLS RGB(R1, G1, B1), H1, L1, S1
            
            If (R1 = 255) And (G1 = 255) And (B1 = 255) Then
                If iLastH <> 0 Then
                    H1 = iLastH
                    S1 = 240
                End If
            Else
                iLastH = H1
            End If
            
            H1 = H1 + HAddition
            If H1 > 240 Then H1 = H1 - 240
            If H1 < 0 Then H1 = H1 + 240
            
            L1 = L1 * LFactor
            L1 = L1 + LAddition
            If L1 < 0 Then L1 = 0
            If L1 > 240 Then L1 = 240
            
            S1 = S1 * LFactor
            S1 = S1 + SAddition
            If S1 < 1 Then S1 = 1
            If S1 > 240 Then S1 = 240
            
            iColor = ColorHLSToRGB(H1, L1, S1)
            
            If ColorToPreserve > -1 Then ' we don't want to make a color equal to this one
                If iColor = ColorToPreserve Then
                    If iColor = &HFFFFFF Then
                        iColor = iColor - &H10101
                    Else
                        iColor = iColor + &H10101
                    End If
                End If
            End If
            
            iBits(x) = (iColor \ 65536) And 255 ' B
            iBits(x + 1) = (iColor \ 256) And 255 ' G
            iBits(x + 2) = iColor And 255 ' R
        End If
    Next x
    
    'DeleteDC iTmpDC
    
    'iTmpDC = CreateCompatibleDC(0)
    
    iBMPInfo.bmiHeader = iBMPiH
    iBitMap = CreateDIBSection(iTmpDC, iBMPInfo, 0, 0, 0, 0)   ' Create a temp blank image
    iOldObj = SelectObject(iTmpDC, iBitMap)
    SetDIBitsToDevice iTmpDC, 0, 0, iPicWidth, iPicHeight, 0, 0, 0, iBMP.bmHeight, iBits(0), iBMPiH, DIB_RGB_COLORS
    'StretchDIBits iTmpDC, 0, 0, iPicWidth, iPicHeight, 0, 0, iBMP.bmWidth, iBMP.bmHeight, iBits(0), iBMPiH, DIB_RGB_COLORS, vbSrcCopy
    
    SelectObject iTmpDC, iOldObj
'    DeleteObject iBitMap
    DeleteDC iTmpDC
    
    With IID_IDispatch
       .Data1 = &H20400
       .Data4(0) = &HC0
       .Data4(7) = &H46
    End With
    
    With iPicBmp
       .Size = Len(iPicBmp)         'Length of structure
       .Type = vbPicTypeBitmap  'Type of Picture (bitmap)
       .hBmp = iBitMap          'Handle to bitmap
       .hPal = 0&               'Handle to palette (may be null)
     End With
    
    Call OleCreatePictureIndirect(iPicBmp, IID_IDispatch, 1, iPic)
    
    Set AdjustPictureWithHLS = iPic

End Function

Public Function AdjustColorWithHLS(nColor As Long, Optional HAddition As Long, Optional LAddition As Long, Optional SAddition As Long, Optional LFactor As Single = 1, Optional SFactor As Single = 1) As Long
    Dim R1 As Long
    Dim G1 As Long
    Dim B1 As Long
    Dim H1 As Long
    Dim L1 As Long
    Dim S1 As Long
    
    R1 = nColor And 255 ' R
    G1 = (nColor \ 256) And 255 ' G
    B1 = (nColor \ 65536) And 255 ' B
    
    ColorRGBToHLS RGB(R1, G1, B1), H1, L1, S1
    
    H1 = H1 + HAddition
    If H1 > 240 Then H1 = H1 - 240
    If H1 < 0 Then H1 = H1 + 240
    
    L1 = L1 * LFactor
    L1 = L1 + LAddition
    If L1 < 0 Then L1 = 0
    If L1 > 240 Then L1 = 240
    
    S1 = S1 * LFactor
    S1 = S1 + SAddition
    If S1 < 1 Then S1 = 1
    If S1 > 240 Then S1 = 240
    
    AdjustColorWithHLS = ColorHLSToRGB(H1, L1, S1)
    
End Function

Public Function SetColorToSameHue(nColor As Long, nReferenceColor As Long) As Long
    Dim iColor As Long
    Dim iReferenceColor As Long
    
    Dim R1 As Long
    Dim G1 As Long
    Dim B1 As Long
    Dim H1 As Long
    Dim L1 As Long
    Dim S1 As Long
    
    Dim R2 As Long
    Dim G2 As Long
    Dim B2 As Long
    Dim H2 As Long
    Dim L2 As Long
    Dim S2 As Long
    
    TranslateColor nColor, 0, iColor
    TranslateColor nReferenceColor, 0, iReferenceColor

    R1 = iColor And 255 ' R
    G1 = (iColor \ 256) And 255 ' G
    B1 = (iColor \ 65536) And 255 ' B
    
    ColorRGBToHLS RGB(R1, G1, B1), H1, L1, S1
    
    R2 = iReferenceColor And 255 ' R
    G2 = (iReferenceColor \ 256) And 255 ' G
    B2 = (iReferenceColor \ 65536) And 255 ' B
    
    ColorRGBToHLS RGB(R2, G2, B2), H2, L2, S2
    
    SetColorToSameHue = ColorHLSToRGB(H2, L1, S1)
End Function
