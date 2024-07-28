Attribute VB_Name = "mGDIplus"
' From great stuff:
'
'   http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1
'   by Avery
'
'   Platform SDK Redistributable: GDI+ RTM
'   http://www.microsoft.com/downloads/release.asp?releaseid=32738

Option Explicit

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type RGBTRIPPLE
   rgbRed As Byte
   rgbGreen As Byte
   rgbBlue As Byte
End Type

Private Type RGBQUAD
     rgbBlue As Byte
     rgbGreen As Byte
     rgbRed As Byte
     rgbReserved As Byte
End Type

Public Type BITMAPINFOHEADER '40 bytes
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

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors() As RGBQUAD
End Type

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Public Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_REALSIZE = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_LOADTRANSPARENT = &H20
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Public Const BI_RGB = 0&
Public Const BI_RLE4 = 2&
Public Const BI_RLE8 = 1&
Public Const DIB_RGB_COLORS = 0
Public Type GDIPlusStartupInput
    GDIPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Public Enum GpStatus2
    [Ok1] = 0
    [GenericError] = 1
    [InvalidParameter] = 2
    [OutOfMemory] = 3
    [ObjectBusy] = 4
    [InsufficientBuffer] = 5
    [NotImplemented] = 6
    [Win32Error] = 7
    [WrongState] = 8
    [Aborted] = 9
    [FileNotFound] = 10
    [ValueOverflow ] = 11
    [AccessDenied] = 12
    [UnknownImageFormat] = 13
    [FontFamilyNotFound] = 14
    [FontStyleNotFound] = 15
    [NotTrueTypeFont] = 16
    [UnsupportedGdiplusVersion] = 17
    [GdiplusNotInitialized ] = 18
    [PropertyNotFound] = 19
    [PropertyNotSupported] = 20
End Enum

Public Enum GpUnit
    [UnitWorld]
    [UnitDisplay]
    [UnitPixel]
    [UnitPoint]
    [UnitInch]
    [UnitDocument]
    [UnitMillimeter]
End Enum

Public Enum InterpolationMode
    [InterpolationModeInvalid] = -1
    [InterpolationModeDefault]
    [InterpolationModeLowQuality]
    [InterpolationModeHighQuality]
    [InterpolationModeBilinear]
    [InterpolationModeBicubic]
    [InterpolationModeNearestNeighbor]
    [InterpolationModeHighQualityBilinear]
    [InterpolationModeHighQualityBicubic]
End Enum

Public Enum PixelOffsetMode
    [PixelOffsetModeInvalid] = -1
    [PixelOffsetModeDefault]
    [PixelOffsetModeHighSpeed]
    [PixelOffsetModeHighQuality]
    [PixelOffsetModeNone]
    [PixelOffsetModeHalf]
End Enum

Public Enum QualityMode
    [QualityModeInvalid] = -1
    [QualityModeDefault]
    [QualityModeLow]
    [QualityModeHigh]
End Enum

Public Enum ImageLockMode
    [ImageLockModeRead] = &H1
    [ImageLockModeWrite] = &H2
    [ImageLockModeUserInputBuf] = &H4
End Enum

Public Enum RotateFlipType
    [RotateNoneFlipNone] = 0
    [Rotate90FlipNone] = 1
    [Rotate180FlipNone] = 2
    [Rotate270FlipNone] = 3
    [RotateNoneFlipX] = 4
    [Rotate90FlipX] = 5
    [Rotate180FlipX] = 6
    [Rotate270FlipX] = 7
    [RotateNoneFlipY] = Rotate180FlipX
    [Rotate90FlipY] = Rotate270FlipX
    [Rotate180FlipY] = RotateNoneFlipX
    [Rotate270FlipY] = Rotate90FlipX
    [RotateNoneFlipXY] = Rotate180FlipNone
    [Rotate90FlipXY] = Rotate270FlipNone
    [Rotate180FlipXY] = RotateNoneFlipNone
    [Rotate270FlipXY] = Rotate90FlipNone
End Enum

'//

Public Const PixelFormat24bppRGB      As Long = &H21808

Public Const PropertyTagTypeByte      As Long = 1
Public Const PropertyTagTypeASCII     As Long = 2
Public Const PropertyTagTypeShort     As Long = 3
Public Const PropertyTagTypeLong      As Long = 4
Public Const PropertyTagTypeRational  As Long = 5
Public Const PropertyTagTypeUndefined As Long = 7
Public Const PropertyTagTypeSLONG     As Long = 9
Public Const PropertyTagTypeSRational As Long = 10

Public Const PropertyTagFrameDelay    As Long = &H5100
Public Const PropertyTagLoopCount     As Long = &H5101

Public Const FrameDimensionTime       As String = "{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"
Public Const FrameDimensionResolution As String = "{84236F7B-3BD3-428F-8DAB-4EA1439CA315}"
Public Const FrameDimensionPage       As String = "{7462DC86-6180-4C7E-8E3F-EE7333A7A483}"

'//

Private Type BitmapData
    Width       As Long
    Height      As Long
    Stride      As Long
    PixelFormat As Long
    Scan0       As Long
    Reserved    As Long
End Type

Private Type RECTL
    x As Long
    y As Long
    W As Long
    H As Long
End Type

Public Type CLSID
    Data1         As Long
    Data2         As Integer
    Data3         As Integer
    Data4(0 To 7) As Byte
End Type

Public Type PropertyItem
    propId As Long
    Length As Long
    Type   As Integer
    Value  As Long
End Type
    
Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, InputBuf As GDIPlusStartupInput, Optional ByVal OutputBuf As Long = 0) As GpStatus2
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As GpStatus2

Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal Filename As String, hImage As Long) As GpStatus2
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal hImage As Long, Width As Long) As GpStatus2
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal hImage As Long, Height As Long) As GpStatus2
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GpStatus2

Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDc As Long, hGraphics As Long) As GpStatus2
Public Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal PixelFormat As Long, Scan0 As Any, hBitmap As Long) As GpStatus2
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As GpStatus2

Public Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Mode As InterpolationMode) As GpStatus2
Public Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Mode As PixelOffsetMode) As GpStatus2

Public Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hBitmap As Long, RECT As RECTL, ByVal Flags As Long, ByVal PixelFormat As Long, LockedBitmapData As BitmapData) As GpStatus2
Public Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hBitmap As Long, LockedBitmapData As BitmapData) As GpStatus2

Public Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal hImage As Long, ByVal rfType As RotateFlipType) As GpStatus2
Public Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus2
Public Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As GpStatus2

Public Declare Function GdipImageGetFrameCount Lib "gdiplus" (ByVal Image As Long, dimensionID As CLSID, Count As Long) As GpStatus2
Public Declare Function GdipImageSelectActiveFrame Lib "gdiplus" (ByVal Image As Long, dimensionID As CLSID, ByVal frameIndex As Long) As GpStatus2
Public Declare Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long, Size As Long) As GpStatus2
Public Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long, ByVal propSize As Long, Buffer As PropertyItem) As GpStatus2

'//

Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As CLSID) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal psString As Any) As Long

Private Type ARGBQUAD
    B As Byte
    G As Byte
    R As Byte
    a As Byte
End Type


Public Function ColorARGB(ByVal Color As Long, ByVal Alpha As Byte) As Long
  
  Dim uARGB As ARGBQUAD
  Dim aSwap As Byte

   Call CopyMemory(uARGB, Color, 4)
   With uARGB
        .a = Alpha: aSwap = .R: .R = .B: .B = aSwap
   End With
   Call CopyMemory(ColorARGB, uARGB, 4)
End Function



'========================================================================================
' Helpers
'========================================================================================

Public Function GetPropertyValue(Item As PropertyItem) As Variant
   
    If (Item.Value = 0 Or Item.Length = 0) Then Call Err.Raise(5, "GetPropertyValue")

    '-- We'll make Undefined types a Btye array as it seems the safest choice...
    Select Case Item.Type
        
        Case PropertyTagTypeByte, PropertyTagTypeUndefined
         
            ReDim buffByte(1 To Item.Length) As Byte
            Call CopyMemory(buffByte(1), ByVal Item.Value, Item.Length)
            GetPropertyValue = buffByte()
            Erase buffByte()

        Case PropertyTagTypeASCII
         
            GetPropertyValue = PtrToStrA(Item.Value)
         
        Case PropertyTagTypeShort
         
            ReDim buffShort(1 To (Item.Length / 2)) As Integer
            Call CopyMemory(buffShort(1), ByVal Item.Value, Item.Length)
            GetPropertyValue = buffShort()
            Erase buffShort()
         
        Case PropertyTagTypeLong, PropertyTagTypeSLONG
         
            ReDim buffLong(1 To (Item.Length / 4)) As Long
            Call CopyMemory(buffLong(1), ByVal Item.Value, Item.Length)
            GetPropertyValue = buffLong()
            Erase buffLong()
         
        Case PropertyTagTypeRational, PropertyTagTypeSRational
         
            ReDim buffLongPair(1 To (Item.Length / 8), 1 To 2) As Long
            Call CopyMemory(buffLongPair(1, 1), ByVal Item.Value, Item.Length)
            GetPropertyValue = buffLongPair()
            Erase buffLongPair()

        Case Else
            
            Call Err.Raise(461, "GetPropertyValue")
    End Select
End Function

Public Sub DEFINE_GUID2(ByVal sGuid As String, uCLSID As CLSID)
    
    Call CLSIDFromString(StrPtr(sGuid), uCLSID)
End Sub

Public Function StretchDIB24Ex( _
                oDIB24 As cDIB, _
                ByVal hDc As Long, _
                ByVal x As Long, ByVal y As Long, _
                ByVal nWidth As Long, ByVal nHeight As Long, _
                Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, _
                Optional ByVal nSrcWidth As Long, Optional ByVal nSrcHeight As Long, _
                Optional ByVal Interpolate As Boolean = False _
                ) As Long

  Dim gplRet As Long
  
  Dim hGraphics As Long
  Dim hBitmap   As Long
  Dim bmpRect   As RECTL
  Dim bmpData   As BitmapData
  
    If (oDIB24.BPP = 24) Then
        
        If (nSrcWidth = 0) Then nSrcWidth = oDIB24.Width
        If (nSrcHeight = 0) Then nSrcHeight = oDIB24.Height
      
        '-- Prepare image info
        With bmpRect
            .W = oDIB24.Width
            .H = oDIB24.Height
        End With
        With bmpData
            .Width = oDIB24.Width
            .Height = oDIB24.Height
            .Stride = -oDIB24.BytesPerScanLine
            .PixelFormat = [PixelFormat24bppRGB]
            .Scan0 = oDIB24.lpBits - .Stride * (oDIB24.Height - 1)
        End With
        
        '-- Initialize Graphics object
        gplRet = GdipCreateFromHDC(hDc, hGraphics)
        
        '-- Initialize blank Bitmap and assign DIB data
        gplRet = GdipCreateBitmapFromScan0(oDIB24.Width, oDIB24.Height, 0, [PixelFormat24bppRGB], ByVal 0, hBitmap)
        gplRet = GdipBitmapLockBits(hBitmap, bmpRect, [ImageLockModeWrite] Or [ImageLockModeUserInputBuf], [PixelFormat24bppRGB], bmpData)
        gplRet = GdipBitmapUnlockBits(hBitmap, bmpData)

        '-- Render
        gplRet = GdipSetInterpolationMode(hGraphics, [InterpolationModeNearestNeighbor] + -(2 * Interpolate))
        gplRet = GdipSetPixelOffsetMode(hGraphics, [PixelOffsetModeHighQuality])
        gplRet = GdipDrawImageRectRectI(hGraphics, hBitmap, x, y, nWidth, nHeight, xSrc, ySrc, nSrcWidth, nSrcHeight, [UnitPixel], 0)
        
        '-- Clean up
        gplRet = GdipDeleteGraphics(hGraphics)
        gplRet = GdipDisposeImage(hBitmap)
        
        '-- Success
        StretchDIB24Ex = (gplRet = [Ok1])
    End If
End Function

'//

Private Function PtrToStrW(ByVal lpsz As Long) As String
  
  Dim sOut As String
  Dim lLen As Long

    lLen = lstrlenW(lpsz)

    If (lLen > 0) Then
        sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen * 2)
        PtrToStrW = StrConv(sOut, vbFromUnicode)
    End If
End Function

Private Function PtrToStrA(ByVal lpsz As Long) As String
  
  Dim sOut As String
  Dim lLen As Long

    lLen = lstrlenA(lpsz)

    If (lLen > 0) Then
        sOut = String$(lLen, vbNullChar)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen)
        PtrToStrA = sOut
    End If
End Function


Public Function GetTrueBits(pb As PictureBox, abPicture() As Byte, BI As BITMAPINFO) As Boolean
   Dim bmp As BITMAP
   Call GetObjectAPI(pb.Picture, Len(bmp), bmp)
   ReDim BI.bmiColors(0)
   With BI.bmiHeader
       .biSize = Len(BI.bmiHeader)
       .biWidth = bmp.bmWidth
       .biHeight = bmp.bmHeight
       .biPlanes = 1
       .biBitCount = 24
       .biCompression = BI_RGB
       .biSizeImage = BytesPerScanLine(.biWidth) * .biHeight
       ReDim abPicture(BytesPerScanLine(.biWidth) - 1, .biHeight - 1)
   End With
   GetTrueBits = GetDIBits(pb.hDc, pb.Picture, 0, BI.bmiHeader.biHeight, abPicture(0, 0), BI, DIB_RGB_COLORS)
End Function

Public Function GetBits(pb As PictureBox, abPicture() As Byte, BI As BITMAPINFO, Optional clrDepth As Long) As Boolean
   Dim BuffSize As Long
   Dim biArray() As Byte
   Dim bih As BITMAPINFOHEADER
   
   ReDim BI.bmiColors(0)
   BI.bmiHeader.biSize = Len(BI.bmiHeader)
   Call GetDIBits(pb.hDc, pb.Picture, 0, 0, ByVal 0, BI.bmiHeader, DIB_RGB_COLORS)
   If clrDepth > 0 Then
      If clrDepth < BI.bmiHeader.biBitCount Then
         BI.bmiHeader.biBitCount = clrDepth
      End If
   End If
   BI.bmiHeader.biCompression = BI_RGB
   BuffSize = BI.bmiHeader.biWidth
   Select Case BI.bmiHeader.biBitCount
       Case 1
            BuffSize = Int((BuffSize + 7) / 8)
            ReDim biArray(Len(bih) + 4 * 2 - 1)
       Case 4
            BuffSize = Int((BuffSize + 1) / 2)
            ReDim biArray(Len(bih) + 4 * 16 - 1)
       Case 8
            BuffSize = BuffSize
            ReDim biArray(Len(bih) + 4 * 256 - 1)
       Case 16
            BuffSize = BuffSize * 2
            ReDim biArray(Len(bih) + 4 - 1)
       Case 24
            BuffSize = BuffSize * 3
            ReDim biArray(Len(bih) + 4 - 1)
       Case 32
            BuffSize = BuffSize * 3
            ReDim biArray(Len(bih) + 4 * 3 - 1)
   End Select
   ReDim BI.bmiColors((UBound(biArray) + 1 - Len(bih)) \ 4 - 1)
   BuffSize = (Int((BuffSize + 3) / 4)) * 4
   ReDim abPicture(BuffSize - 1, BI.bmiHeader.biHeight - 1)
   BuffSize = BuffSize * BI.bmiHeader.biHeight
   BI.bmiHeader.biSizeImage = BuffSize
   CopyMemory biArray(0), BI, Len(BI.bmiHeader)
   GetBits = GetDIBits(pb.hDc, pb.Picture, 0, BI.bmiHeader.biHeight, abPicture(0, 0), biArray(0), DIB_RGB_COLORS)
   CopyMemory BI, biArray(0), Len(bih)
   CopyMemory BI.bmiColors(0), biArray(Len(bih) + 1), UBound(biArray) - Len(bih) + 1
End Function

Public Function SetBits(pb As PictureBox, abPicture() As Byte, BI As BITMAPINFO) As Boolean
   SetBits = SetDIBitsToDevice(pb.hDc, 0, 0, BI.bmiHeader.biWidth, BI.bmiHeader.biHeight, 0, 0, 0, BI.bmiHeader.biHeight, abPicture(0, 0), BI, DIB_RGB_COLORS)
   pb.Refresh
End Function

Public Function BytesPerScanLine(ByVal lWidth As Long) As Long
    BytesPerScanLine = (lWidth * 3 + 3) And &HFFFFFFFC
End Function
' Convert Automation color to Windows color
Public Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function



