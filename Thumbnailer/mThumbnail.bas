Attribute VB_Name = "mThumbnail"
Option Explicit
Option Compare Text

Private Const MAX_PATH                   As Long = 260
Private Const INVALID_HANDLE_VALUE       As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY   As Long = &H10
Private Const FILE_ATTRIBUTE_READONLY    As Long = &H1
Private Const FILE_ATTRIBUTE_RODIRECTORY As Long = FILE_ATTRIBUTE_DIRECTORY + FILE_ATTRIBUTE_READONLY

Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved0      As Long
    dwReserved1      As Long
    cFileName        As String * MAX_PATH
    cAlternate       As String * 14
End Type
Public m_CRC As clsCRC
'Dim PicEx As cPictureEx

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private DIBbpp               As Byte         ' Current color depth

Public Const DRIVE_UNKNOWN As Long = 0
Public Const DRIVE_NO_ROOT_DIR As Long = 1
Public Const DRIVE_REMOVABLE As Long = 2
Public Const DRIVE_FIXED As Long = 3
Public Const DRIVE_REMOTE As Long = 4
Public Const DRIVE_CDROM As Long = 5
Public Const DRIVE_RAMDISK As Long = 6


Private Const LOCALE_USER_DEFAULT     As Long = &H400
Private Const LOCALE_NOUSEROVERRIDE   As Long = &H80000000
Private Const DATE_SHORTDATE          As Long = &H1

Private Type SYSTEMTIME
    wYear         As Integer
    wMonth        As Integer
    wDayOfWeek    As Integer
    wDay          As Integer
    wHour         As Integer
    wMinute       As Integer
    wSecond       As Integer
    wMilliseconds As Integer
End Type

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Private Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long

'//

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

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'//

Public Const IMAGETYPES_MASK As String = "|PNG|JPG|JPEG|BMP|GIF|WMF|EMF|ICO|TIF|TIFF|PDI|PCX|PCD|ORA|MBM|KOA|HDP|JXR|JLS|HGT|HEIF|HEIC|G3|DNG|CBZ|AVIF|PFM|PGM|PIC|PICT|APNG|PNM|PPM|PSP|QOI|RAW|RGB|BW|WEBP|XPM|JNG|KOALA|LBM|IFF|LBM|MNG|PBM|PBMRAW|PPMRAW|RAS|TARGA|WBMP|CUT|XBM|DDS|FAXG3|SGI|EXR|J2K|HDR|PDF|PSD|TGA|XCF|SVG|JP2|"


Public Type DATABASE_INFO
    Size    As Long
    Entries As Long
End Type


Private Type FILE_INFO
    Filename As String
    FileDate As String
    Filecreationdate As String
    FileSize As Long
    Filelastaccesstime As String
    Filelastwritetime As String
    Filecrc As String
    FilePath As String
    Filewidth As String
    Fileheight As String
    fileExtension As String
End Type

Public m_sDatabasePath As String
Private m_bThumbnailing As Boolean
Private m_oTile         As cTile
Private m_sFolder       As String



'========================================================================================
' Methods
'========================================================================================

Public Sub InitializeModule()
    
    '-- Initialize pattern brush
    Set m_oTile = New cTile
End Sub

Public Sub TerminateModule()
    Set m_oTile = Nothing
End Sub

Public Sub Cancel()

    '-- Cancel thumbnailing
    m_bThumbnailing = False
End Sub

Public Sub UpdateFolder(ByVal sFolder As String)
  
  Dim lItem   As Long
  Dim uFile() As FILE_INFO
  
  Dim aTBIH() As Byte, uTBIH As BITMAPINFOHEADER
  Dim aData() As Byte
    '-- Folder path
    Let m_sFolder = sFolder
    
    '-- Thumbnailing...
    m_bThumbnailing = True
    
    On Local Error GoTo errH
    With frm_Thumbnailer.ucThumbnailView

        '-- Get files list
        If (pvGetFiles(IMAGETYPES_MASK, uFile())) Then
            

            '-- Disable redraw and set items count
            Call .SetRedraw(bRedraw:=False)
            Call .ItemSetCount(UBound(uFile()) + 1)
               
            '-- Add items
            For lItem = 0 To UBound(uFile())
            
               Call .ItemAdd(lItem, uFile(lItem).Filename, uFile(lItem).FileSize, uFile(lItem).FileDate, uFile(lItem).Filecreationdate, uFile(lItem).Filelastaccesstime, uFile(lItem).Filelastwritetime, uFile(lItem).Filecrc, uFile(lItem).FilePath, uFile(lItem).Filewidth)
            Next lItem
            
            '-- Enable redraw and ensure visible first item
            Call .ItemEnsureVisible(0)
            Call .SetRedraw(bRedraw:=True)
           
          
  
                '-
            '-- Add/Get thumbnails
            frm_Thumbnailer.ucProgress.Max = UBound(uFile()) + 1
            
            For lItem = 0 To .Count - 1
                
                '-- Current progress
                frm_Thumbnailer.ucProgress.Value = lItem + 1
                 Call pvSetThumbnail(uFile(lItem), lItem)

                '-- Refresh added item
                Call .RefreshItems(lItem, lItem)
                
                Call VBA.DoEvents
                
            Next lItem
        End If
    End With

errH:
   

    m_bThumbnailing = False
    frm_Thumbnailer.ucProgress.Value = 0
End Sub







'========================================================================================
' Private
'========================================================================================

Private Sub pvSetThumbnail(ByRef uFile As FILE_INFO, ByVal lItem As Long)
            
  Dim oDIBThumb As cDIB
  Dim hImage    As Long
  Dim hGraphics As Long
  Dim bfx As Long, bfW As Long, W As Long
  Dim bfy As Long, bfH As Long, H As Long
  Dim aTBIH(39) As Byte, uTBIH As BITMAPINFOHEADER
  Dim aData()   As Byte
  Dim dstFilename As String '
  Dim sPath As String '
  Dim loadSuccessful As Boolean
  Dim pngpath As String
  Dim sext As String
  Dim TempBmp$  ' Temporary bmp file name
  Dim pnum As String
  Dim FileSpecPath$ '
  Dim IViewPath$
  Dim SavePath$
  Dim aString$
  Dim bString$
  Dim cString$
  Dim res&
  
TempBmp$ = "~!@temp.bmp"   ' NB UNIQUENESS NOT CHECKED!
sPath = ""
sPath = m_sFolder & uFile.Filename
sext = Mid$(sPath, InStrRev(sPath, ".") + 1)
pngpath = App.Path & "\" & "PhotoDemon_print.png"
If Len(sPath) = 0 Then
   Close
   Exit Sub
End If

  '----------------------------------------------------------------------
       Select Case sext
         Case "PNG", "JPG", "JPEG", "BMP", "GIF", "WMF", "EMF", "ICO", "TIF", "TIFF"
       dstFilename = m_sFolder & uFile.Filename '  Call SendMessage(m_hDlg, CDM_GETFILEPATH, MAX_PATH, ByVal sPath)
'======================================================================================================================================================

        Case "PDI", "PCX", "PCD", "ORA", "MBM", "KOA", "HDP", "JXR", "JLS", "HGT", "HEIF", "HEIC", "G3", "DNG", "CBZ", "AVIF", "PFM", "PGM", "PIC", "PICT", "APNG", "PNM", "PPM", "PSP", "QOI", "RAW", "RGB", "BW", "WEBP", "XPM", "JNG", "KOALA", "LBM", "IFF", "LBM", "MNG", "PBM", "PBMRAW", "PPMRAW", "RAS", "TARGA", "WBMP", "CUT", "XBM", "DDS", "FAXG3", "SGI", "EXR", "J2K", "HDR", "PDF", "PSD", "TGA", "XCF", "SVG", "JP2"
    Dim TmpDIB As pdDIB
            
 Set TmpDIB = New pdDIB
loadSuccessful = False
 Set TmpDIB = New pdDIB
      If (Len(sPath) <> 0) Then loadSuccessful = Loading.QuickLoadImageToDIB(sPath, TmpDIB, False, False, True)
    TmpDIB.CompositeBackgroundColor 255, 255, 255
    TmpDIB.SetInitialAlphaPremultiplicationState False
    Saving.QuickSaveDIBAsPNG pngpath, TmpDIB, True, True
    Set TmpDIB = Nothing
dstFilename = pngpath
'======================================================================================================================================================

        Case Else
       MsgBox sPath & "IMAGE FORMAT NOT SUPPORTED NOW"

        Exit Sub
      End Select
'====================================================================================================
    '-- Generate thumbnail...
    If (mGDIplus.GdipLoadImageFromFile(StrConv(dstFilename, vbUnicode), hImage) = [Ok1]) Then
        
        '-- Initialize DIB
        Set oDIBThumb = New cDIB

        '-- Image size
        Call mGDIplus.GdipGetImageWidth(hImage, W)
        Call mGDIplus.GdipGetImageHeight(hImage, H)
        
        '-- Best fit to current thumbnail max. size
        Call oDIBThumb.GetBestFitInfo(W, H, frm_Thumbnailer.ucThumbnailView.ThumbnailWidth, frm_Thumbnailer.ucThumbnailView.ThumbnailHeight, bfx, bfy, bfW, bfH)
        Call oDIBThumb.Create(bfW, bfH, [16_bpp])

        '-- Prepare target surface
        Call mGDIplus.GdipCreateFromHDC(oDIBThumb.hDc, hGraphics)
        
        '-- Tile 'transparent' layer and render thumbnail
        Call m_oTile.Tile(oDIBThumb.hDc, 0, 0, bfW, bfH)
        Call mGDIplus.GdipDrawImageRectI(hGraphics, hImage, 0, 0, bfW, bfH)
        
        '-- Clean up
        Call mGDIplus.GdipDeleteGraphics(hGraphics)
        Call mGDIplus.GdipDisposeImage(hImage)
        
        '-- Prepare bitmap header (thumbnail)
        With uTBIH
            .biSize = Len(uTBIH)
            .biBitCount = oDIBThumb.BPP
            .biWidth = oDIBThumb.Width
            .biHeight = oDIBThumb.Height
            .biSizeImage = oDIBThumb.Size
            .biPlanes = 1
        End With
        
        '-- Prepare data
        ReDim aData(oDIBThumb.Size - 1)
        Call CopyMemory(aTBIH(0), uTBIH, Len(uTBIH))
        Call CopyMemory(aData(0), ByVal oDIBThumb.lpBits, oDIBThumb.Size)
        
        '-- Transfer to database
      
     
      Else
        '-- *Null* transfer to database
        ReDim aData(0)
        Call ZeroMemory(aTBIH(0), Len(uTBIH))
    End If
    '-- Transfer to thumbnail viewer
    Call frm_Thumbnailer.ucThumbnailView.ThumbnailInfo_SetTBIH(lItem, aTBIH())
    Call frm_Thumbnailer.ucThumbnailView.ThumbnailInfo_SetData(lItem, aData())
    If FileExist(pngpath) Then Kill pngpath
End Sub

Private Function pvGetFiles(ByVal sMask As String, uFile() As FILE_INFO) As Boolean
  
  Dim uFileTmp()  As FILE_INFO
  Dim sext        As String
  Dim lExtSep     As Long
  Dim lCount      As Long
  Dim lc          As Long
  
  Dim uWFD        As WIN32_FIND_DATA
  Dim hSearch     As Long
  Dim hNext       As Long
  Dim strHex As String
  Dim strWidth As String
  Dim strHeight As String
  
      Set m_CRC = New clsCRC
    
       m_CRC.Algorithm = Crc32

    '-- Initial storage
    ReDim uFileTmp(100)

    '-- Start searching files (all)
    hNext = 1
    hSearch = FindFirstFile(m_sFolder & "*.*" & vbNullChar, uWFD)
    
    If (hSearch <> INVALID_HANDLE_VALUE) Then
        
        Do While hNext
            If (uWFD.dwFileAttributes <> FILE_ATTRIBUTE_DIRECTORY) Then
              If (FileLen(m_sFolder & pvStripNulls(uWFD.cFileName)) = 0) Then
              strHex = "000000"
               Else
              strHex = Hex(m_CRC.CalculateFile(m_sFolder & pvStripNulls(uWFD.cFileName)))
               End If

              
                '-- Get file name, date and size
                With uFileTmp(lCount)
                    .Filename = pvStripNulls(uWFD.cFileName)
                    .FileDate = pvGetFileDateTimeStr(uWFD.ftLastWriteTime)
                    .FileSize = uWFD.nFileSizeHigh * &HFFFF0000 + uWFD.nFileSizeLow
                    .Filecreationdate = pvGetFileDateTimeStr(uWFD.ftCreationTime)
                    .Filelastaccesstime = pvGetFileDateTimeStr(uWFD.ftLastAccessTime)
                    .Filelastwritetime = pvGetFileDateTimeStr(uWFD.ftLastWriteTime)
                    .Filecrc = strHex
                    .FilePath = m_sFolder & pvStripNulls(uWFD.cFileName)
                                    End With
                lCount = lCount + 1
                
                '-- Resize array [?]
                If ((lCount Mod 100) = 0) Then
                    ReDim Preserve uFileTmp(UBound(uFileTmp()) + 100)
                End If
            End If
            hNext = FindNextFile(hSearch, uWFD)
        Loop
        hNext = FindClose(hSearch)
    End If
    ReDim Preserve uFileTmp(lCount - -(lCount > 0))
    
    '-- Filter files
    If (lCount > 0) Then
        lCount = 0
        ReDim uFile(100)
        
        '-- Check all files
        For lc = 0 To UBound(uFileTmp())
        
            '-- Extension ?
            lExtSep = InStrRev(uFileTmp(lc).Filename, ".")
            If (lExtSep) Then
                
                '-- Get extension
                sext = "|" & Mid$(uFileTmp(lc).Filename, lExtSep + 1) & "|"
                
                '-- Supported file
                If (InStr(1, sMask, sext)) Then
                    
                    '-- Get this file
                    uFile(lCount) = uFileTmp(lc)
                    lCount = lCount + 1
                    
                    '-- Resize array [?]
                    If ((lCount Mod 100) = 0) Then
                        ReDim Preserve uFile(UBound(uFile()) + 100)
                    End If
                End If
            End If
        Next lc
        ReDim Preserve uFile(lCount - -(lCount > 0))
    End If
    
    '-- Success
    pvGetFiles = (lCount > 0)
End Function

Private Static Function pvGetFileDateTimeStr(uFileTime As FILETIME) As String
  
  Dim uFT As FILETIME
  Dim uST As SYSTEMTIME

    Call FileTimeToLocalFileTime(uFileTime, uFT)
    Call FileTimeToSystemTime(uFT, uST)
  
    pvGetFileDateTimeStr = pvGetFileDateStr(uST) & " " & pvGetFileTimeStr(uST)
End Function

Private Static Function pvGetFileDateStr(uSystemTime As SYSTEMTIME) As String
  
  Dim sDate As String * 32
  Dim lLen  As Long
  
    lLen = GetDateFormat(LOCALE_USER_DEFAULT, LOCALE_NOUSEROVERRIDE Or DATE_SHORTDATE, uSystemTime, vbNullString, sDate, 64)
    If (lLen) Then
        pvGetFileDateStr = Left$(sDate, lLen - 1)
    End If
End Function

Private Static Function pvGetFileTimeStr(uSystemTime As SYSTEMTIME) As String
  
  Dim sTime As String * 32
  Dim lLen  As Long
  
    lLen = GetTimeFormat(LOCALE_USER_DEFAULT, LOCALE_NOUSEROVERRIDE, uSystemTime, vbNullString, sTime, 64)
    If (lLen) Then
        pvGetFileTimeStr = Left$(sTime, lLen - 1)
    End If
End Function

Private Function pvStripNulls(ByVal sString As String) As String
    
  Dim lPos As Long

    lPos = InStr(sString, vbNullChar)
    
    If (lPos = 1) Then
        pvStripNulls = vbNullString
    ElseIf (lPos > 1) Then
        pvStripNulls = Left$(sString, lPos - 1)
        Exit Function
    End If
    
    pvStripNulls = sString
End Function
