VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Class:         cDIB.cls
' Author:        Carles P.V.
' Dependencies:
' Last revision: 2004.9.12
'================================================
'
' - 2004.08.18
'
'   * Fixed CreateFromBitmapFile() function:
'     m_uBIH structure was not initialized (filled) -> Save() function failed
'     Thanks to Paul Caton
'
'     Now, private variables 'm_Width, m_Height, m_BPP, m_BytesPerScanline and m_Size'
'     have been removed. Directly taken from m_uBIH structure.
'
'   * Also added CreateFromResourceBitmap() variation (by Paul Caton)
'
' - 2004.08.19
'
'   * Fixed Image() Property:
'     Going back to original routine. Probably problems for using a still selected DIB (?).
'     Anyway, old function worked well when setting/extracting Image to/from VB-Clipboard
'     using vbCFDIB format (but not using vbCFBitmap one) (?).
'     Also, curiously, Image was returned as 'down-top DIB'.
'
'   * Improved Save() function: Removed GlobalLock/GlobalUnlock. Redundant when allocating
'     fixed memory. The GlobalAlloc returns the memory pointer.
'
'   Thanks again, Paul.
'
' - 2004.09.12
'
'   * Fixed CreateFromStdPicture() function.
'   * Added StretchBltMode param. in Stretch() and Paint() functions (NT only).
'

Option Explicit

'-- API:

Private Type BITMAPFILEHEADER
    bfType      As Integer
    bfSize      As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits   As Long
End Type

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Type BITMAPINFO_001
    bmiHeader       As BITMAPINFOHEADER
    bmiColors(7)    As Byte
End Type

Private Type BITMAPINFO_004
    bmiHeader       As BITMAPINFOHEADER
    bmiColors(63)   As Byte
End Type

Private Type BITMAPINFO_008
    bmiHeader       As BITMAPINFOHEADER
    bmiColors(1023) As Byte
End Type

Private Type BITMAPINFO_RGB
    bmiHeader       As BITMAPINFOHEADER
End Type

Private Type GUID2
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PICTDESC
    Size     As Long
    Type     As Long
    hBmp     As Long
    hPal     As Long
    Reserved As Long
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Const DIB_RGB_COLORS      As Long = 0
Private Const COLORONCOLOR        As Long = 3
Private Const HALFTONE            As Long = 4
Private Const OBJ_BITMAP          As Long = 7
Private Const LR_LOADFROMFILE     As Long = &H10
Private Const IMAGE_BITMAP        As Long = 0
Private Const LR_CREATEDIBSECTION As Long = &H2000

Private Declare Function CreateDIBSection_001 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDc As Long, lpBitsInfo As BITMAPINFO_001, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection_004 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDc As Long, lpBitsInfo As BITMAPINFO_004, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection_008 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDc As Long, lpBitsInfo As BITMAPINFO_008, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateDIBSection_RGB Lib "gdi32" Alias "CreateDIBSection" (ByVal hDc As Long, lpBitsInfo As BITMAPINFO_RGB, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hDc As Long, ByVal un1 As Long, ByVal un2 As Long, lpRGBQuad As Any) As Long
Private Declare Function GetDIBColorTable Lib "gdi32" (ByVal hDc As Long, ByVal un1 As Long, ByVal un2 As Long, lpRGBQuad As Any) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDc As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long

'//

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)

'//

Private Const CF_DIB As Long = 8

Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long

'//

Private Const GMEM_FIXED    As Long = &H0
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_DDESHARE As Long = &H2000
Private Const GMEM_ZEROINIT As Long = &H40

Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long

'//

Private Const INVALID_HANDLE_VALUE  As Long = -1
Private Const CREATE_ALWAYS         As Long = 2
Private Const GENERIC_READ          As Long = &H80000000
Private Const GENERIC_WRITE         As Long = &H40000000
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

'//

'-- Public Enums.:
Public Enum eBPPCts
    [01_bpp] = 1
    [04_bpp] = 4
    [08_bpp] = 8
    [16_bpp] = 16
    [24_bpp] = 24
    [32_bpp] = 32
End Enum

Public Enum eStretchBltModeCts
    [sbmColorOnColor] = COLORONCOLOR
    [sbmHalftone] = HALFTONE
End Enum

'-- Private Variables:
Private m_uBIH    As BITMAPINFOHEADER
Private m_hDC     As Long
Private m_hDIb    As Long
Private m_hOldDIB As Long
Private m_lpBits  As Long

'//

'========================================================================================
' Class
'========================================================================================

Private Sub Class_Terminate()

    '-- Destroy current DIB
    Call Me.Destroy
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Function Create(ByVal NewWidth As Long, ByVal NewHeight As Long, ByVal NewBPP As eBPPCts) As Long
    
  Dim BI_001 As BITMAPINFO_001
  Dim BI_004 As BITMAPINFO_004
  Dim BI_008 As BITMAPINFO_008
  Dim BI_RGB As BITMAPINFO_RGB
  
    '-- Destroy previous
    Call Me.Destroy
    
    '-- Define DIB header
    With m_uBIH
        .biSize = Len(m_uBIH)
        .biPlanes = 1
        .biBitCount = NewBPP
        .biWidth = NewWidth
        .biHeight = NewHeight
        .biSizeImage = 4 * ((.biWidth * .biBitCount + 31) \ 32) * .biHeight
    End With
    Select Case NewBPP
        Case [01_bpp]: BI_001.bmiHeader = m_uBIH
        Case [04_bpp]: BI_004.bmiHeader = m_uBIH
        Case [08_bpp]: BI_008.bmiHeader = m_uBIH
        Case Else:     BI_RGB.bmiHeader = m_uBIH
    End Select
    
    '-- Create DIB and select into a DC
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
        Select Case NewBPP
            Case [01_bpp]: m_hDIb = CreateDIBSection_001(m_hDC, BI_001, DIB_RGB_COLORS, m_lpBits, 0, 0)
            Case [04_bpp]: m_hDIb = CreateDIBSection_004(m_hDC, BI_004, DIB_RGB_COLORS, m_lpBits, 0, 0)
            Case [08_bpp]: m_hDIb = CreateDIBSection_008(m_hDC, BI_008, DIB_RGB_COLORS, m_lpBits, 0, 0)
            Case Else:     m_hDIb = CreateDIBSection_RGB(m_hDC, BI_RGB, DIB_RGB_COLORS, m_lpBits, 0, 0)
        End Select
        If (m_hDIb <> 0) Then
            m_hOldDIB = SelectObject(m_hDC, m_hDIb)
          Else
            Call Me.Destroy
        End If
    End If
    
    '-- Success
    Create = (m_hDIb <> 0)
End Function

Public Function CreateFromBitmapFile(ByVal Filename As String) As Long
  
  Dim uBI As BITMAP
    
    '-- File exists [?]
    If (FileLen(Filename)) Then
    
        '-- Destroy previous
        Call Me.Destroy

        '-- Create DIB and select into a DC
        m_hDC = CreateCompatibleDC(0)
        If (m_hDC <> 0) Then
            m_hDIb = LoadImage(App.hInstance, Filename, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
            If (m_hDIb <> 0) Then
                m_hOldDIB = SelectObject(m_hDC, m_hDIb)
              Else
                Call Me.Destroy
            End If
        End If
        
        '-- Get DIB props.:
        If (m_hDIb <> 0) Then
            '-- Get object
            Call GetObject(m_hDIb, Len(uBI), uBI)
            '-- Get props.:
            With m_uBIH
                .biSize = Len(m_uBIH)
                .biPlanes = 1
                .biBitCount = uBI.bmBitsPixel
                .biWidth = uBI.bmWidth
                .biHeight = uBI.bmHeight
                .biSizeImage = 4 * ((.biWidth * .biBitCount + 31) \ 32) * .biHeight
            End With
            '-- Bits pointer
            m_lpBits = uBI.bmBits
            
            '-- Success
            CreateFromBitmapFile = uBI.bmBitsPixel
        End If
    End If
End Function

Public Function CreateFromResourceBitmap(ByVal Filename As String, ByVal ResID As Long) As Long
  
  Dim hInstance As Long
  Dim uBI       As BITMAP
    
    '-- File exists [?]
    If (FileLen(Filename)) Then
    
        '-- Destroy previous
        Call Me.Destroy
 
        '-- Create DIB and select into a DC
        m_hDC = CreateCompatibleDC(0)
        If (m_hDC <> 0) Then
          
            hInstance = LoadLibrary(Filename)
            If (hInstance) Then
                m_hDIb = LoadImageLong(hInstance, ResID, IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION)
                If (m_hDIb <> 0) Then
                    m_hOldDIB = SelectObject(m_hDC, m_hDIb)
                  Else
                    Call Me.Destroy
                End If
                Call FreeLibrary(hInstance)
            End If
        End If
        
        '-- Get DIB props.:
        If (m_hDIb <> 0) Then
            '-- Get object
            Call GetObject(m_hDIb, Len(uBI), uBI)
            '-- Get props.:
            With m_uBIH
                .biSize = Len(m_uBIH)
                .biPlanes = 1
                .biBitCount = uBI.bmBitsPixel
                .biWidth = uBI.bmWidth
                .biHeight = uBI.bmHeight
                .biSizeImage = 4 * ((.biWidth * .biBitCount + 31) \ 32) * .biHeight
            End With
            '-- Bits pointer
            m_lpBits = uBI.bmBits
            
            '-- Success
            CreateFromResourceBitmap = uBI.bmBitsPixel
        End If
    End If
End Function

Public Function CreateFromStdPicture(Image As StdPicture, Optional ByVal Force32bpp As Boolean = False) As Long

  Dim lhDC     As Long
  Dim lhOldBmp As Long
  Dim uBI      As BITMAP
  Dim aPal()   As Byte
  Dim aBI()    As Byte
    
    If (Not Image Is Nothing) Then
        
        '-- Check object type
        If (GetObjectType(Image.handle) = OBJ_BITMAP) Then
        
            '-- Get object info
            Call GetObject(Image.handle, Len(uBI), uBI)
            
            '-- Create DIB
            If (Create(uBI.bmWidth, uBI.bmHeight, IIf(Force32bpp, [32_bpp], uBI.bmBitsPixel))) Then
                lhDC = CreateCompatibleDC(m_hDC)
                If (lhDC <> 0) Then
                    lhOldBmp = SelectObject(lhDC, Image.handle)
                    
                    With m_uBIH
                        If (Force32bpp Or .biBitCount > 8) Then
                            Call BitBlt(m_hDC, 0, 0, .biWidth, .biHeight, lhDC, 0, 0, vbSrcCopy)
                          Else
                            ReDim aBI(1 To Len(m_uBIH) + 4 * 2 ^ .biBitCount)
                            Call CopyMemory(aBI(1), m_uBIH, Len(m_uBIH))
                            Call GetDIBColorTable(lhDC, 0, 2 ^ .biBitCount, aBI(Len(m_uBIH) + 1))
                            Call SetDIBColorTable(m_hDC, 0, 2 ^ .biBitCount, aBI(Len(m_uBIH) + 1))
                            Call GetDIBits(lhDC, Image.handle, 0, .biHeight, ByVal m_lpBits, aBI(1), DIB_RGB_COLORS)
                        End If
                    End With
                    
                    '-- Destroy temp. DC
                    Call SelectObject(lhDC, lhOldBmp)
                    Call DeleteDC(lhDC)
                    
                    '-- Success
                    CreateFromStdPicture = m_uBIH.biBitCount
                End If
            End If
        End If
    End If
End Function

Public Function CreateFromClipboard() As Long
    
  Dim hClipMem    As Long
  Dim lClipPtrHDR As Long
  Dim lClipPtrPAL As Long
  Dim lClipPtrBMP As Long
  Dim uBIH        As BITMAPINFOHEADER
  
    If (OpenClipboard(0)) Then
        
        '-- Get clipboard mem. handle
        hClipMem = GetClipboardData(CF_DIB)
        If (hClipMem <> 0) Then
            
            '-- Get pointer / access data
            lClipPtrHDR = GlobalLock(hClipMem)
            If (lClipPtrHDR <> 0) Then
                
                '-- Extract DIB info header
                Call CopyMemory(uBIH, ByVal lClipPtrHDR, Len(uBIH))
                
                '-- Create DIB / fill data
                If (Me.Create(uBIH.biWidth, uBIH.biHeight, uBIH.biBitCount)) Then
                    
                    With m_uBIH
                        If (.biBitCount <= 8) Then
                            lClipPtrPAL = lClipPtrHDR + Len(m_uBIH)
                            lClipPtrBMP = lClipPtrPAL + 4 * 2 ^ .biBitCount
                            Call SetDIBColorTable(m_hDC, 0, 2 ^ .biBitCount, ByVal lClipPtrPAL)
                            Call CopyMemory(ByVal m_lpBits, ByVal lClipPtrBMP, .biSizeImage)
                          Else
                            lClipPtrBMP = lClipPtrHDR + Len(m_uBIH)
                            Call CopyMemory(ByVal m_lpBits, ByVal lClipPtrBMP, .biSizeImage)
                        End If
                        
                        '-- Success
                        CreateFromClipboard = m_uBIH.biBitCount
                    End With
                End If
                Call GlobalUnlock(hClipMem)
            End If
        End If
        Call CloseClipboard
    End If
End Function

Public Function CopyToClipboard() As Long

  Dim hMem        As Long
  Dim lClipPtrHDR As Long
  Dim lClipPtrPAL As Long
  Dim lClipPtrBMP As Long
  Dim lSize       As Long
  Dim aPal()      As Byte
    
    If (m_hDIb <> 0) Then
    
        If (OpenClipboard(0)) Then
            
            '-- Size of packed DIB
            lSize = IIf(m_uBIH.biBitCount <= 8, 4 * 2 ^ m_uBIH.biBitCount, 0)
            lSize = lSize + Len(m_uBIH) + m_uBIH.biSizeImage
            
            '-- Get mem. handle and its pointer
            hMem = GlobalAlloc(GMEM_MOVEABLE, lSize)
            If (hMem <> 0) Then
                lClipPtrHDR = GlobalLock(hMem)
                If (lClipPtrHDR <> 0) Then
                
                    '-- Copy data
                    Call CopyMemory(ByVal lClipPtrHDR, m_uBIH, Len(m_uBIH))
                    
                    If (m_uBIH.biBitCount <= 8) Then
                        ReDim aPal(1 To 4 * 2 ^ m_uBIH.biBitCount)
                        lClipPtrPAL = lClipPtrHDR + Len(m_uBIH)
                        lClipPtrBMP = lClipPtrPAL + 4 * 2 ^ m_uBIH.biBitCount
                        Call GetDIBColorTable(m_hDC, 0, 2 ^ m_uBIH.biBitCount, aPal(1))
                        Call CopyMemory(ByVal lClipPtrPAL, aPal(1), UBound(aPal()))
                        Call CopyMemory(ByVal lClipPtrBMP, ByVal lpBits, m_uBIH.biSizeImage)
                      Else
                        lClipPtrBMP = lClipPtrHDR + Len(m_uBIH)
                        Call CopyMemory(ByVal lClipPtrBMP, ByVal lpBits, m_uBIH.biSizeImage)
                    End If
                    Call GlobalUnlock(hMem)
                    
                    '-- Paste to Clipboard
                    Call EmptyClipboard
                    CopyToClipboard = SetClipboardData(CF_DIB, hMem)
                    Call CloseClipboard
                End If
            End If
        End If
    End If
End Function

Public Function CloneTo(oDIB As cDib) As Long
    
  Dim aPal() As Byte
    
    If (m_hDIb <> 0) Then
    
        '-- Create dest. DIB
        Call oDIB.Create(m_uBIH.biWidth, m_uBIH.biHeight, m_uBIH.biBitCount)
        
        '-- Palette [?]
        If (m_uBIH.biBitCount <= 8) Then
            ReDim aPal(4 * 2 ^ m_uBIH.biBitCount - 1)
            Call Me.GetPalette(aPal())
            Call oDIB.SetPalette(aPal())
        End If
        
        '-- Copy color data
        Call CopyMemory(ByVal oDIB.lpBits, ByVal m_lpBits, m_uBIH.biSizeImage)
        
        '-- Success
        CloneTo = oDIB.BPP
    End If
End Function

Public Function Save(ByVal Filename As String) As Boolean

  Const BITMAPTYPE  As Integer = &H4D42
  
  Dim hFile         As Long
  Dim lret          As Long
  Dim lBytesWritten As Long

  Dim hMem          As Long
  Dim lptr          As Long
  Dim uBFH          As BITMAPFILEHEADER
  Dim aPal()        As Byte
  Dim lColors       As Long
  
    If (m_hDIb <> 0) Then
    
        '-- Get file handler
        hFile = CreateFile(Filename, GENERIC_READ Or GENERIC_WRITE, ByVal 0, ByVal 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
        
        If (hFile <> INVALID_HANDLE_VALUE) Then
            
            '-- Get palette entries count
            lColors = IIf(m_uBIH.biBitCount <= 8, 2 ^ m_uBIH.biBitCount, 0)
            
            '-- Prepare file header
            With uBFH
                .bfType = BITMAPTYPE
                .bfSize = Len(m_uBIH) + 4 * lColors + m_uBIH.biSizeImage
                .bfOffBits = Len(uBFH) + Len(m_uBIH) + 4 * lColors
            End With
            
            '-- Not long alignment...
            hMem = GlobalAlloc(GMEM_FIXED + GMEM_ZEROINIT, Len(uBFH))
            lptr = GlobalLock(hMem)
            Call CopyMemory(ByVal lptr, uBFH.bfType, 2)
            Call CopyMemory(ByVal lptr + 2, uBFH.bfSize, 4)
            Call CopyMemory(ByVal lptr + 6, 0, 4)
            Call CopyMemory(ByVal lptr + 10, uBFH.bfOffBits, 4)
            
            '-- Write file header
            lret = WriteFile(hFile, ByVal lptr, Len(uBFH), lBytesWritten, ByVal 0)
            Call GlobalUnlock(hMem)
            Call GlobalFree(hMem)
            
            If (lret <> 0) Then
                
                '-- Write bitmap info header
                lret = WriteFile(hFile, m_uBIH, Len(m_uBIH), lBytesWritten, ByVal 0)
                
                '-- Palette [?]
                If (lret <> 0 And m_uBIH.biBitCount <= 8) Then
                    ReDim aPal(1 To 4 * lColors)
                    Call GetDIBColorTable(m_hDC, 0, lColors, aPal(1))
                    '-- Write palette
                    lret = WriteFile(hFile, ByVal VarPtr(aPal(1)), 4 * lColors, lBytesWritten, ByVal 0)
                End If
                
                '-- Write bitmap data
                If (lret <> 0) Then
                    lret = WriteFile(hFile, ByVal m_lpBits, m_uBIH.biSizeImage, lBytesWritten, ByVal 0)
                End If
            End If
            Call CloseHandle(hFile)
            
            '-- Success
            Save = (lret <> 0)
        End If
    End If
End Function

Public Sub Destroy()

    '-- Destroy DIB
    If (m_hDC <> 0) Then
        If (m_hDIb <> 0) Then
            Call SelectObject(m_hDC, m_hOldDIB)
            Call DeleteObject(m_hDIb)
        End If
        Call DeleteDC(m_hDC)
    End If
    
    '-- Reset BIH structure
    Call ZeroMemory(m_uBIH, Len(m_uBIH))
    
    '-- Reset DIB vars.
    m_hDC = 0
    m_hDIb = 0
    m_hOldDIB = 0
    m_lpBits = 0
End Sub

'//

Public Function LoadBlt(ByVal hSrcDC As Long, _
                        Optional ByVal X As Long = 0, _
                        Optional ByVal Y As Long = 0, _
                        Optional ByVal nWidth As Long, _
                        Optional ByVal nHeight As Long) As Long
    
    If (m_hDIb <> 0) Then
        
        If (nWidth = 0) Then nWidth = m_uBIH.biWidth
        If (nHeight = 0) Then nHeight = m_uBIH.biHeight
        
        '-- Paint from source (*)
        LoadBlt = BitBlt(m_hDC, 0, 0, nWidth, nHeight, hSrcDC, X, Y, vbSrcCopy)
    End If
    
'(*) - Palette based bitmaps: indexes not preserved.
'    - ARGB bitmaps: alpha not preserved (->0)
End Function

Public Function Paint(ByVal hDc As Long, _
                      Optional ByVal X As Long = 0, Optional ByVal Y As Long = 0, _
                      Optional ByVal ROP As RasterOpConstants = vbSrcCopy, _
                      Optional ByVal StretchBltMode As eStretchBltModeCts = [sbmColorOnColor]) As Long
    
    Paint = Me.Stretch(hDc, X, Y, m_uBIH.biWidth, m_uBIH.biHeight, , , , , ROP, StretchBltMode)
End Function

Public Function Stretch(ByVal hDc As Long, _
                        ByVal X As Long, ByVal Y As Long, _
                        ByVal nWidth As Long, ByVal nHeight As Long, _
                        Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, _
                        Optional ByVal nSrcWidth As Long, Optional ByVal nSrcHeight As Long, _
                        Optional ByVal ROP As RasterOpConstants = vbSrcCopy, _
                        Optional ByVal StretchBltMode As eStretchBltModeCts = [sbmColorOnColor]) As Long
  
  Dim lOldMode As Long
    
    If (m_hDIb <> 0) Then
        
        If (nSrcWidth = 0) Then nSrcWidth = m_uBIH.biWidth
        If (nSrcHeight = 0) Then nSrcHeight = m_uBIH.biHeight
        
        lOldMode = SetStretchBltMode(hDc, StretchBltMode)
        Stretch = StretchBlt(hDc, X, Y, nWidth, nHeight, m_hDC, xSrc, ySrc, nSrcWidth, nSrcHeight, ROP)
        Call SetStretchBltMode(hDc, lOldMode)
    End If

' Note: StretchDIBits preserves indexes in palette based
'       bitmaps only under Windows XP.
End Function

Public Sub Cls(Optional ByVal ColorRef As Long = &H0)
  
  Dim uRect  As RECT2
  Dim hBrush As Long
   
    If (m_hDIb <> 0) Then
        Call SetRect(uRect, 0, 0, m_uBIH.biWidth, m_uBIH.biHeight)
        hBrush = CreateSolidBrush(ColorRef)
        Call FillRect(m_hDC, uRect, hBrush)
        Call DeleteObject(hBrush)
    End If
End Sub

Public Sub Reset()
    
    If (m_hDIb <> 0) Then
        Call ZeroMemory(ByVal m_lpBits, m_uBIH.biSizeImage)
    End If
End Sub

'//

Public Sub SetPalette(Palette() As Byte)
    
    If (m_hDIb <> 0) Then
        Call SetDIBColorTable(m_hDC, 0, (UBound(Palette()) - LBound(Palette()) + 1) \ 4, Palette(LBound(Palette())))
    End If
    
' Note: Palette array should be initialized before.
End Sub

Public Sub GetPalette(Palette() As Byte)
    
    If (m_hDIb <> 0) Then
        Call GetDIBColorTable(m_hDC, 0, 2 ^ m_uBIH.biBitCount, Palette(LBound(Palette())))
    End If
    
' Note: Palette array should be initialized before.
End Sub

'//

Public Sub GetBestFitInfo(ByVal SrcW As Long, ByVal SrcH As Long, _
                          ByVal DstW As Long, ByVal DstH As Long, _
                          bfx As Long, bfy As Long, _
                          bfW As Long, bfH As Long, _
                          Optional ByVal StretchFit As Boolean = False)
                          
  Dim cW As Single
  Dim Ch As Single
    
    If ((SrcW > DstW Or SrcH > DstH) Or StretchFit) Then
        cW = DstW / SrcW
        Ch = DstH / SrcH
        If (cW < Ch) Then
            bfW = DstW
            bfH = SrcH * cW
          Else
            bfH = DstH
            bfW = SrcW * Ch
        End If
      Else
        bfW = SrcW
        bfH = SrcH
    End If
    
    If (bfW < 1) Then bfW = 1
    If (bfH < 1) Then bfH = 1
    bfx = (DstW - bfW) \ 2
    bfy = (DstH - bfH) \ 2
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get hDc() As Long
    hDc = m_hDC
End Property

Public Property Get hDib() As Long
    hDib = m_hDIb
End Property

Public Property Get lpBits() As Long
    lpBits = m_lpBits
End Property

Public Property Get Width() As Long
    Width = m_uBIH.biWidth
End Property

Public Property Get Height() As Long
    Height = m_uBIH.biHeight
End Property

Public Property Get BPP() As eBPPCts
    BPP = m_uBIH.biBitCount
End Property

Public Property Get BytesPerScanLine() As Long
    BytesPerScanLine = ((m_uBIH.biWidth * m_uBIH.biBitCount + 31) \ 32) * 4
End Property

Public Property Get Size() As Long
    Size = m_uBIH.biSizeImage
End Property

Public Property Get Image() As StdPicture
    
  Dim lhSrcDC       As Long
  Dim lhDC          As Long
  Dim lhBmp         As Long
  Dim lhOldBmp      As Long
  Dim lpPictDesc    As PICTDESC
  Dim IID_IDispatch As GUID2

    If (m_hDIb <> 0) Then
    
        '-- Create screen compatible bitmap
        lhSrcDC = GetDC(0)
        lhDC = CreateCompatibleDC(lhSrcDC)
        lhBmp = CreateCompatibleBitmap(lhSrcDC, m_uBIH.biWidth, m_uBIH.biHeight)
        lhOldBmp = SelectObject(lhDC, lhBmp)
        Call ReleaseDC(0, lhSrcDC)
        
        '-- Paint from DIB
        Call Me.Paint(lhDC)
        '-- Clean up
        Call SelectObject(lhDC, lhOldBmp)
        Call DeleteDC(lhDC)
        
        '-- Fill image info
        With lpPictDesc
            .Size = Len(lpPictDesc)
            .Type = vbPicTypeBitmap
            .hBmp = lhBmp
            .hPal = 0
        End With
        '-- Fill GUID2 info
        With IID_IDispatch
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
        '-- Create picture from bitmap handle
        Call OleCreatePictureIndirect(lpPictDesc, IID_IDispatch, 1, Image)
    End If
End Property

'Public Property Get Image() As StdPicture
'
'  Dim lpPictDesc As PICTDESC
'  Dim lGuid(3)   As Long
'
'    '-- Fill image info
'    With lpPictDesc
'        .Size = Len(lpPictDesc)
'        .Type = vbPicTypeBitmap
'        .hBmp = m_hDIB
'        .hPal = 0
'    End With
'    '-- Fill GUID2 info
'    lGuid(0) = &H7BF80980
'    lGuid(1) = &H101ABF32
'    lGuid(2) = &HAA00BB8B
'    lGuid(3) = &HAB0C3000
'
'    '-- Create image from DIB handle
'    Call OleCreatePictureIndirect(lpPictDesc, lGuid(0), True, Image)
'End Property
