VERSION 5.00
Begin VB.Form frm_Thumbnailer 
   Caption         =   "Thumbnailer for photodemon"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   10995
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   579
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   733
   Begin PhotoDemon.pdPictureBox ucplayer 
      Height          =   3255
      Left            =   240
      Top             =   4320
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5741
   End
   Begin VB.Timer tmrExploreFolder 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3780
      Top             =   855
   End
   Begin PhotoDemon.ucStatusbar ucStatusbar 
      Height          =   285
      Left            =   30
      Top             =   7680
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   503
   End
   Begin PhotoDemon.ucSplitter ucSplitterH 
      Height          =   6735
      Left            =   4320
      Top             =   840
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   11880
   End
   Begin PhotoDemon.ucSplitter ucSplitterV 
      Height          =   60
      Left            =   240
      Top             =   4080
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   106
   End
   Begin VB.ComboBox cbPath 
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "fMain1.frx":0000
      Left            =   240
      List            =   "fMain1.frx":0002
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin PhotoDemon.ucProgress ucProgress 
      Height          =   270
      Left            =   7755
      Top             =   7695
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   476
   End
   Begin PhotoDemon.ucFolderView ucFolderView 
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5530
   End
   Begin PhotoDemon.ucThumbnailView ucThumbnailView 
      Height          =   6735
      Left            =   5160
      TabIndex        =   2
      Top             =   480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   11880
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "create New Folders.."
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Exit"
         Index           =   1
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuGoTop 
      Caption         =   "&Go"
      Begin VB.Menu mnuGo 
         Caption         =   "&Back"
         Index           =   0
      End
      Begin VB.Menu mnuGo 
         Caption         =   "&Forward"
         Index           =   1
      End
      Begin VB.Menu mnuGo 
         Caption         =   "&Up"
         Index           =   2
      End
   End
   Begin VB.Menu mnuViewTop 
      Caption         =   "&View"
      Begin VB.Menu mnuView 
         Caption         =   "&Refresh"
         Index           =   0
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Thumbnails"
         Checked         =   -1  'True
         Index           =   2
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Details"
         Index           =   3
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnusysinfo 
         Caption         =   "Detailed System info"
      End
      Begin VB.Menu mnuhelp22 
         Caption         =   "Drive Space info"
      End
      Begin VB.Menu mnuok 
         Caption         =   "SYSTEM Required info"
      End
      Begin VB.Menu mnuhelp 
         Caption         =   "&About"
         Index           =   0
      End
   End
   Begin VB.Menu mnuViewModeTop 
      Caption         =   "View mode"
      Visible         =   0   'False
      Begin VB.Menu mnuViewMode 
         Caption         =   "View &thumbnails"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "View &details"
         Index           =   1
      End
   End
   Begin VB.Menu mnuContextThumbnailTop 
      Caption         =   "Context thumbnail"
      Visible         =   0   'False
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Properties"
         Index           =   0
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Shell open..."
         Index           =   1
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Shell edit"
         Index           =   2
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "File Information"
         Index           =   3
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "open into Photodemon"
         Index           =   4
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "open this folder"
         Index           =   5
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Print Image"
         Index           =   6
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "openwith..external program"
         Index           =   7
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Copy to..."
         Index           =   8
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Move to.."
         Index           =   9
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Change Folder"
         Index           =   10
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "convert File"
         Index           =   11
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Add New Folders"
         Index           =   12
      End
      Begin VB.Menu mnuContextThumbnail 
         Caption         =   "Delete"
         Index           =   13
      End
   End
End
Attribute VB_Name = "frm_Thumbnailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
' Application:   Thumbnailer.exe
' Version:       1.0.0
' Last revision: 2004.11.29
' Dependencies:  gdiplus.dll (place in application folder)
'
' Author:        Carles P.V. - ©2004
'========================================================================================



Option Explicit

'-- A little bit of API

Private Declare Sub InitCommonControls Lib "Comctl32" ()
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellExecuteEx Lib "shell32" (USEI As SHELLEXECUTEINFO) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Private Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long

Private Const SW_SHOWNORMAL = 1
'Dim PicEx As cPictureEx

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 
Private Const WM_SETICON              As Long = &H80
Private Const LR_SHARED               As Long = &H8000&
Private Const ICON_SMALL              As Long = 0
Private Const IMAGE_ICON              As Long = 1

Private Const CB_ERR                  As Long = (-1)
Private Const CB_GETCURSEL            As Long = &H147
Private Const CB_SETCURSEL            As Long = &H14E
Private Const CB_SHOWDROPDOWN         As Long = &H14F
Private Const CB_GETDROPPEDSTATE      As Long = &H157

Private Const SEM_NOGPFAULTERRORBOX   As Long = &H2&

Private Const SEE_MASK_INVOKEIDLIST   As Long = &HC
Private Const SEE_MASK_FLAG_NO_UI     As Long = &H400
Private Const SW_NORMAL               As Long = 1
'Public WithEvents DIBDither As cDIBDither   ' DIB Dither object  (1, 4, 8 bpp)

Private Type SHELLEXECUTEINFO
    cbSize       As Long
    fMask        As Long
    hWnd         As Long
    lpVerb       As String
    lpFile       As String
    lpParameters As String
    lpDirectory  As String
    nShow        As Long
    hInstApp     As Long
    lpIDList     As Long
    lpClass      As String
    hkeyClass    As Long
    dwHotKey     As Long
    hIcon        As Long
    hProcess     As Long
End Type
Private Const MAX_PATH                   As Long = 260
Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type
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
'-- Private variables

Private m_bInIDE           As Boolean
Private m_GDIPlusToken     As Long
Private m_bLoaded          As Boolean
Private m_bEnding          As Boolean
Private m_bComboHasFocus   As Boolean
Public m_LastFilename      As String       ' Current file
Public m_LastPath          As String       ' Current path
Private m_Temp              As String       ' Temporary folder
Private m_CRC As clsCRC

Private Const m_PathLevels As Long = 100
Private m_Paths()          As String
Private m_PathsPos         As Long
Private m_PathsMax         As Long
Private m_bSkipPath        As Boolean
Const PROCESS_ALL_ACCESS& = &H1F0FFF
Const STILL_ACTIVE& = &H103&
Const INFINITE& = &HFFFF
Public DIBbpp               As Byte         ' Current color depth
'Public DIBPal               As New cDIBPal  ' DIB Palette object (1, 4, 8 bpp)
Private m_FileExt           As String       ' Current file/ext
'Public DIBSave              As New cDIBSave ' Save object (BMP)  (1, 4, 8, 24 bpp)

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long



Private Sub cbPath_DblClick()
Dim a As String

        a = Files.PathBrowseDialog(frm_Thumbnailer.hWnd)

    '-- Path selected
         With ucFolderView
            If (.Path <> cbPath.Text) Then
                .Path = a
            End If
        End With
End Sub

'========================================================================================
' Initializing / Terminating
'========================================================================================

Private Sub Form_Initialize()

    If (App.PrevInstance) Then End
   
    '-- Initialize common controls
    Call InitCommonControls
    
    '-- Load the GDI+ library
    Dim uGpSI As mGDIplus.GDIPlusStartupInput
    Let uGpSI.GDIPlusVersion = 1
    If (mGDIplus.GdiplusStartup(m_GDIPlusToken, uGpSI) <> [Ok1]) Then
        Call MsgBox("Error initializing application!", vbCritical)
        End
    End If
End Sub
Private Function JustDoIt1() As Boolean
    On Error Resume Next
    Dim a As String
    Dim I As Integer
      Dim sfile As String
  Dim sPath As String
    Dim cFile As pdFSO
    Set cFile = New pdFSO

  
  Dim sext As String
  Dim bSuccess As Boolean
  
  sPath = ucFolderView.Path
  sfile = sPath & ucThumbnailView.ItemText(ucThumbnailView.ItemFindState(, [tvFocused]), [tvFileName]) & vbNullChar

     a = PathBrowseDialog(frm_Thumbnailer.hWnd)
     
    If Len(a) > 0 Then
           'FileCopy FindPath(ucFolderView.Path, sfile), FindPath(a, sfile)
                                    If cFile.FileCopyW(ucFolderView.Path & ucThumbnailView.ItemText(ucThumbnailView.ItemFindState(, [tvFocused]), [tvFileName]) & vbNullChar, a & ucThumbnailView.ItemText(ucThumbnailView.ItemFindState(, [tvFocused]), [tvFileName]) & vbNullChar) Then

                        End If

            DoEvents
        JustDoIt1 = True
    End If
End Function
Private Sub Form_Load()
Dim strExe As String
        strExe = App.Path
            If Right$(strExe, 1) <> "\" Then strExe = strExe & "\"
        strExe = strExe & App.EXEName & ".exe"
        If Not strExe = App.Path & "\" & "Photo" & "Demon" & ".exe" Then
        MsgBox App.Path & "\" & "Photo" & "Demon" & ".exe"
        End
        End If
    If (m_bLoaded = False) Then
        m_bLoaded = True
        
        '-- Small icon
        Call SendMessage(Me.hWnd, WM_SETICON, ICON_SMALL, ByVal LoadImageAsString(App.hInstance, ByVal "SMALL_ICON", IMAGE_ICON, 16, 16, LR_SHARED))
        
        '-- Initialize database-thumbnail module / Load settings
        Call mThumbnail.InitializeModule
        Call mSettings.LoadSettings

        '-- Modify some menus
        mnuGo(0).Caption = mnuGo(0).Caption & vbTab & "Alt+Left"
        mnuGo(1).Caption = mnuGo(1).Caption & vbTab & "Alt+Right"
        mnuGo(2).Caption = mnuGo(2).Caption & vbTab & "Alt+Up"
     
        
        '-- Initialize paths list
        Call pvChangeDropDownListHeight(cbPath, 400)

        '-- Initialize folder view
        With ucFolderView
            Call .Initialize
            .HasLines = False
        End With
        
        '-- Initialize thumbnail view
       '-- Initialize thumbnail view
        With ucThumbnailView
            Call .Initialize(IMAGETYPES_MASK, "|", _
                             uAPP_SETTINGS.ViewMode, _
                             uAPP_SETTINGS.ViewColumnWidth(0), _
                             uAPP_SETTINGS.ViewColumnWidth(1))
            Call .SetThumbnailSize(uAPP_SETTINGS.ThumbnailWidth, uAPP_SETTINGS.ThumbnailHeight)
        End With
        
        '-- Initialize player
      
        
        '-- Initialize status bar
        With ucStatusbar
            Call .Initialize(SizeGrip:=True)
            Call .AddPanel(, 150, , [sbSpring])
            Call .AddPanel(, 150)
            Call .AddPanel(, 150)
        End With
        
        '-- Initialize splitters
        Call ucSplitterH.Initialize(Me)
        Call ucSplitterV.Initialize(Me)
        
        '-- Show form
        Call Me.Show: Me.Refresh: Call VBA.DoEvents
        
        '-- Initialize Back/Forward paths list / Go to last recent path
        ReDim m_Paths(0 To m_PathLevels)
        If (cbPath.List(0) <> vbNullString) Then
            m_bSkipPath = True
            cbPath.ListIndex = 0
            m_Paths(1) = cbPath.List(0)
            m_PathsPos = 1
          Else
            Call pvCheckNavigationButtons
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If (m_bLoaded) Then
        m_bEnding = False
        
        '-- Save all settings
        Call mSettings.SaveSettings
        
        '-- Terminate all
        Call mThumbnail.Cancel 'Fix this termination! (-> independent thread: ActiveX EXE ?)
        Call mThumbnail.TerminateModule
        m_bLoaded = False
        '-- Shut down gdiplus session
    End If
End Sub

Private Sub Form_Terminate()

    If (Not inIDE()) Then
        Call SetErrorMode(SEM_NOGPFAULTERRORBOX) '(*)
    End If
    End
    
'(*) From vbAccelerator
'    http://www.vbaccelerator.com/home/VB/Code/Libraries/XP_Visual_Styles/Preventing_Crashes_at_Shutdown/article.asp
'    KBID 309366 (http://support.microsoft.com/default.aspx?scid=kb;en-us;309366)
End Sub




'========================================================================================
' Resizing
'========================================================================================

Private Sub Form_Resize()
  
  Const DXMIN As Long = 200
  Const DXMAX As Long = 225
  Const DYMIN As Long = 200
  Const DYMAX As Long = 200
  Const DSEP  As Long = 2
    
    On Error Resume Next
    
    '-- Resize splitters
    Call ucSplitterH.Move(ucSplitterH.Left, cbPath.Height + 2 * DSEP, ucSplitterH.Width, Me.ScaleHeight - cbPath.Height - ucStatusbar.Height - 3 * DSEP)
    Call ucSplitterV.Move(DSEP, ucSplitterV.Top, ucSplitterH.Left, ucSplitterV.Height)
    
    '-- Update their min/max pos.
    ucSplitterH.xMax = Me.ScaleWidth - DXMAX
    ucSplitterH.xMin = DXMIN
    ucSplitterV.yMax = Me.ScaleHeight - DYMAX
    ucSplitterV.yMin = DYMIN
    
    '-- Relocate splitters
    If (Me.WindowState = vbNormal) Then
        If (ucSplitterH.Left < ucSplitterH.xMin) Then ucSplitterH.Left = ucSplitterH.xMin
        If (ucSplitterV.Top < ucSplitterV.yMin) Then ucSplitterV.Top = ucSplitterV.yMin
        If (ucSplitterH.Left > ucSplitterH.xMax) Then ucSplitterH.Left = ucSplitterH.xMax
        If (ucSplitterV.Top > ucSplitterV.yMax) Then ucSplitterV.Top = ucSplitterV.yMax
    End If
    
    '-- Status bar size-grip?
    Call SetParent(ucProgress.hWnd, Me.hWnd)
    ucStatusbar.SizeGrip = Not (Me.WindowState = vbMaximized)
    Call SetParent(ucProgress.hWnd, ucStatusbar.hWnd)
    Call ucStatusbar_Resize
    
    '-- Relocate controls
    Call cbPath.Move(DSEP, DSEP, Me.ScaleWidth - 2 * DSEP)
    Call ucFolderView.Move(DSEP, cbPath.Height + 2 * DSEP, ucSplitterH.Left - DSEP, ucSplitterV.Top - cbPath.Height - 2 * DSEP)
    Call ucThumbnailView.Move(ucSplitterH.Left + ucSplitterH.Width, cbPath.Height + 2 * DSEP, Me.ScaleWidth - ucSplitterH.Left - ucSplitterH.Width - DSEP, Me.ScaleHeight - cbPath.Height - ucStatusbar.Height - 3 * DSEP)
    Call ucplayer.Move(DSEP, ucSplitterV.Top + ucSplitterV.Height, ucSplitterH.Left - DSEP, Me.ScaleHeight - cbPath.Height - ucStatusbar.Height - ucSplitterV.Height - ucFolderView.Height - 3 * DSEP)
 
 
 
 
 Dim srcImagePath As String
 Dim Item As Long
     srcImagePath = ucFolderView.Path & ucThumbnailView.ItemText(ucThumbnailView.ItemFindState(, [tvFocused]), [tvFileName]) & vbNullChar

  Dim loadSuccessful As Boolean
  
                       Dim TmpDIB As pdDIB: Set TmpDIB = New pdDIB
  
loadSuccessful = False


        'Use PD's central load function to load a copy of the requested image
        If (LenB(srcImagePath) <> 0) Then loadSuccessful = Loading.QuickLoadImageToDIB(srcImagePath, TmpDIB, False, False)
        
                'If the image load failed, display a placeholder message; otherwise, render the image to the picture box
        If loadSuccessful Then
            ucplayer.CopyDIB TmpDIB, True, True
        Else
            ucplayer.PaintText g_Language.TranslateMessage("previews disabled"), 10!, False, True
        End If
                 
                 
                 
                 
                 
                 

        
              
'
        
        
            ucStatusbar.PanelText(2) = ucplayer.GetWidth & "x" & ucplayer.GetHeight

       


    
    Screen.MousePointer = vbDefault

    On Error GoTo 0
End Sub

Private Sub mnusystem_Click()
End Sub

Private Sub MNUOOPEN_Click()
      
  

      



End Sub

Private Sub mnuDatabase_Click(Index As Integer)

End Sub

Private Sub mnuhelp22_Click()
frm_SystemInfo.Show
End Sub

Private Sub mnuok_Click()
   MsgBox "PhotoDemon thumbnailer unicode version" & App.Major & "." & App.Minor & vbCrLf & vbCrLf & "PhotoDemon- 2024 and system required info:" & vbCrLf & "tested ok upto windows  os 10 version 22h2 os build 19045.4598 " & vbCrLf & "not for windows  os 10 version 22h2 os build 19045.4651"

End Sub

Private Sub mnusysinfo_Click()
Dim OSVer As clsOSInfo

    Set OSVer = New clsOSInfo

    Dim s As String
    With OSVer
        s = s & vbCrLf & "OS Name: " & .OSName
        s = s & vbCrLf & "Service Pack ver.: " & .SPVer
        s = s & vbCrLf & "Is Server? " & .IsServer
        s = s & vbCrLf & "Bitness: " & .Bitness
        s = s & vbCrLf & "Is Win x64: " & .IsWin64
        s = s & vbCrLf & "Is Win x32: " & .IsWin32
        s = s & vbCrLf & "Edition: " & .Edition
        s = s & vbCrLf & "Suite mask: " & .SuiteMask
        s = s & vbCrLf & "ProductType: " & .ProductType
        s = s & vbCrLf & "PlatformID: " & .PlatformID & " (" & .Platform & ")"
        s = s & vbCrLf & "Is Domain controller: " & .IsDomainController
        s = s & vbCrLf & "Is Embedded: " & .IsEmbedded
        s = s & vbCrLf & "OS - XP/Server 2003(R2)? " & .IsWindowsXP
        s = s & vbCrLf & "OS - Vista/Server 2008? " & .IsWindowsVista
        s = s & vbCrLf & "OS - 7/Server 2008R2? " & .IsWindows7
        s = s & vbCrLf & "OS - 8/Server 2012? " & .IsWindows8
        s = s & vbCrLf & "OS - 8.1/Server 2012R2? " & .IsWindows8OrGreater
        s = s & vbCrLf & "OS - 10/Server 2016? " & .IsWindows10
        s = s & vbCrLf & "OS - XP or newer? " & .IsWindowsXPOrGreater
        s = s & vbCrLf & "OS - XP SP3 or newer? " & .IsWindowsXP_SP3OrGreater
        s = s & vbCrLf & "OS - Vista or newer? " & .IsWindowsVistaOrGreater
        s = s & vbCrLf & "OS - 7 or newer? " & .IsWindows7OrGreater
        s = s & vbCrLf & "OS - 8 or newer? " & .IsWindows8OrGreater
        s = s & vbCrLf & "OS - 8.1 or newer? " & .IsWindows8Point1OrGreater
        s = s & vbCrLf & "OS - 10 or newer? " & .IsWindows10OrGreater
        s = s & vbCrLf & "OS - 11 or newer? " & .IsWindows11OrGreater
        s = s & vbCrLf & "Major: " & .Major
        s = s & vbCrLf & "Minor: " & .Minor
        s = s & vbCrLf & "Major + Minor:         " & .MajorMinor
        s = s & vbCrLf & "Major + Minor (NtDll): " & .MajorMinorNTDLL
        s = s & vbCrLf & "Build: " & .Build
        s = s & vbCrLf & "NT Dll Major.Minor.Rev: " & .NtDllVersion
        s = s & vbCrLf & "Revision: " & .Revision
        s = s & vbCrLf & "ReleaseId: " & .ReleaseId
        s = s & vbCrLf & "DisplayVersion: " & .DisplayVersion
        s = s & vbCrLf & "Language in dialogues: " & .LangDisplayCode & " " & .LangDisplayName & " " & .LangDisplayNameFull
        s = s & vbCrLf & "Language of OS inslallation: " & .LangSystemCode & " " & .LangSystemName & " " & .LangSystemNameFull
        s = s & vbCrLf & "Language for non-Unicode programs: " & .LangNonUnicodeCode & " " & .LangNonUnicodeName & " " & .LangNonUnicodeNameFull
        s = s & vbCrLf & "ID of default locale: " & .LCID_UserDefault
        s = s & vbCrLf & "Process integrity level: " & .IntegrityLevel
        s = s & vbCrLf & "Elevated process? " & .IsElevated
        s = s & vbCrLf & "Is Local system context? " & .IsLocalSystemContext
        s = s & vbCrLf & "User name: " & .UserName
        s = s & vbCrLf & "User group: " & .UserType
        s = s & vbCrLf & "Is in Admin group? " & .IsAdminGroup
        s = s & vbCrLf & "User sid of current process owner: " & .SID_CurrentProcess
        s = s & vbCrLf & "Computer name: " & .ComputerName
        s = s & vbCrLf & "Safe boot? " & .IsSafeBoot & " (" & .SafeBootMode & ")"
        s = s & vbCrLf & "Secure Boot supported? " & .SecureBootSupported & " (Enabled? " & .SecureBoot & ")"
        s = s & vbCrLf & "TestSigning: " & .TestSigning
        s = s & vbCrLf & "DebugMode: " & .DebugMode
        s = s & vbCrLf & "CodeIntegrity: " & .CodeIntegrity
        s = s & vbCrLf & "File System Case sensitive? " & .IsFileSystemCaseSensitive
        s = s & vbCrLf & "OEM Codepage: " & .CodepageOEM & " (" & .CodepageOEM_File & ")"
        s = s & vbCrLf & "ANSI Codepage: " & .CodepageANSI & " (" & .CodepageANSI_File & ")"
        s = s & vbCrLf & "Memory MiB (Free/Total): " & .MemoryFree & "/" & .MemoryTotal & " (Loaded: " & .MemoryLoad & "%)"
        MsgBox s
    End With
        Set OSVer = Nothing

End Sub

Private Sub pdListBoxViewOD1_Click()

End Sub


Private Sub ucStatusbar_Resize()

  Dim x1 As Long, y1 As Long
  Dim x2 As Long, y2 As Long
    
    '-- Relocate progress bar
    If (ucStatusbar.hWnd) Then
        Call ucStatusbar.GetPanelRect(3, x1, y1, x2, y2)
        Call MoveWindow(ucProgress.hWnd, x1 + 2, y1 + 2, x2 - x1 - 4, y2 - y1 - 4, 0)
    End If
End Sub

Private Sub ucSplitterH_Release()
    Call Form_Resize
End Sub

Private Sub ucSplitterV_Release()
    Call Form_Resize
End Sub



'========================================================================================
' Menus
'========================================================================================

Private Sub MNUFILE_Click(Index As Integer)
     Select Case Index
        
        Case 0 '-- Back
            ShowPDDialog vbModal, frm_AddDir

            
        Case 1 '-- Forward
           '-- Exit
    Call Unload(Me)
     
    End Select
   
End Sub

Private Sub mnuGo_Click(Index As Integer)
    
    Select Case Index
        
        Case 0 '-- Back
            Call pvUndoPath
            
        Case 1 '-- Forward
            Call pvRedoPath
            
        Case 2 '-- Up
            Call ucFolderView.Go([fvGoUp])
            Call pvCheckNavigationButtons
    End Select
End Sub

Private Sub mnuView_Click(Index As Integer)
  
    Select Case Index
        
        Case 0    '-- Refresh
            
            If (Not ucFolderView.PathIsRoot) Then
                Call ucThumbnailView.Clear
                m_bSkipPath = True
                Call ucFolderView_ChangeAfter(vbNullString)
            End If
        
        Case Else '-- View mode changed
            
            Screen.MousePointer = vbArrowHourglass
            ucThumbnailView.Visible = False
            
            '-- Modify main menu and change view mode
            Select Case Index
                
                Case 2 '-- Thumbnails
                    mnuView(3).Checked = False
                    mnuView(2).Checked = True
                    mnuViewMode(1).Checked = False
                    mnuViewMode(0).Checked = True
                    ucThumbnailView.ViewMode = [tvThumbnail]
                
                Case 3 '-- Details
                    mnuView(2).Checked = False
                    mnuView(3).Checked = True
                    mnuViewMode(0).Checked = False
                    mnuViewMode(1).Checked = True
                    ucThumbnailView.ViewMode = [tvDetails]
            End Select
          
            
            '-- Store
            uAPP_SETTINGS.ViewMode = ucThumbnailView.ViewMode
            
            ucThumbnailView.Visible = True
            Screen.MousePointer = vbDefault
    End Select
End Sub

Private Sub mnuViewMode_Click(Index As Integer)
    
    Call mnuView_Click(Index + 2)
End Sub


Private Sub mnuHelp_Click(Index As Integer)
    
    Call MsgBox("PhotoDemon  v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & _
                "PhotoDemon  - 2024" & Space$(15), _
                vbInformation, "About")
End Sub

'//

Private Sub mnuContextPreview_Click(Index As Integer)

  Dim lColor As Long

    Select Case Index
            
        Case 0 '-- Background color...
            
            If (lColor <> -1) Then
                uAPP_SETTINGS.PreviewBackColor = lColor
            End If
            
        Case 2 '-- Pause/Resume
            
          
        
        Case 4 '-- Rotate +90º
         
        
        Case 5 '-- Rotate -90º
            
           
            
        Case 6 '-- Copy image
            
          
            Case 7 '-- Copy image
            MsgBox "mnu empty"
    End Select
    
End Sub

Private Sub mnuContextThumbnail_Click(Index As Integer)

  Dim lItm As Long
  Dim USEI As SHELLEXECUTEINFO
  Dim lRet As Long
  Dim sfile As String
  Dim sFileName As String
  Dim sPath As String
  Dim sTitle As String

  Dim sext As String
  Dim bSuccess As Boolean
  
  sPath = ucFolderView.Path
  sfile = sPath & ucThumbnailView.ItemText(ucThumbnailView.ItemFindState(, [tvFocused]), [tvFileName]) & vbNullChar
  sFileName = ucThumbnailView.ItemText(ucThumbnailView.ItemFindState(, [tvFocused]), [tvFileName]) & vbNullChar
  
Select Case Index
'=======================================================================================================================================================
Case 0  'properties
      With USEI
                .cbSize = Len(USEI)
                .fMask = SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
                .hWnd = Me.hWnd
                .lpParameters = vbNullChar
                .lpDirectory = vbNullChar
                .lpVerb = "properties"
                .lpFile = sfile
                .nShow = 0
      End With
                Call VBA.DoEvents
                lRet = ShellExecuteEx(USEI)
'=======================================================================================================================================================
Case 1 '-- Shell open
            Call ShellExecute(Me.hWnd, vbNullString, sfile, vbNullString, "C:\", SW_SHOWNORMAL)
'=======================================================================================================================================================
Case 2 '-- Shell edit
            Call ShellExecute(Me.hWnd, "edit", sfile, "", sPath, 1)
'=======================================================================================================================================================
Case 3 '-- File informations
       Dim s As String
       Dim R As String

       Dim sCRC As String
       Dim sexte As String

       Set m_CRC = New clsCRC
           m_CRC.Algorithm = Crc32
       Dim uWFD        As WIN32_FIND_DATA
    
           sCRC = Hex(m_CRC.CalculateFile(sfile))
           sexte = Files.FileGetExtension(sFileName)
                
                Dialoginfo.Show
                With Dialoginfo.Text1
.Text = "Pragram Name:- " & App.EXEName & vbCrLf & "Pragram Description:- " & App.FileDescription & vbCrLf & "Pragram Title:- " & App.Title & vbCrLf & "Pragram version:- " & Updates.GetPhotoDemonVersion() & vbCrLf & "Pragram maker:- " & App.CompanyName
End With
With Dialoginfo.Text2
.Text = "File Folder:- " & sPath & vbCrLf
.Text = .Text + "- " & "  File Fullname:- " & sfile
.SelStart = Len(.Text)
                End With
                With Dialoginfo.Text3
.Text = "File Name:- " & sFileName & vbCrLf & " Extension:- *." & sexte & vbCrLf
.Text = .Text + "  CRC Checksum:- " & sCRC
.SelStart = Len(.Text)
                End With

 '====================================
 Case 4 ' import image


  sTitle = Files.FileGetName(sPath, True)
                        
                       
                        If Files.FileGetExtension(sPath) <> "pdi" Then
                        Loading.LoadFileAsNewImage1 sfile, sTitle, False
                          frm_Thumbnailer.ZOrder

                        On Error GoTo gh
                    
                     End If
      

gh:
Loading.LoadFileAsNewImage sfile, sTitle, False
frm_Thumbnailer.ZOrder
Exit Sub
'=======================================================================================================================================================
Case 5 '-- explore this folder
                                Dim FilePath As String, shellCommand As String
            shellCommand = "explorer.exe /select,""" & sfile & """"
            Shell shellCommand, vbNormalFocus
'=======================================================================================================================================================
Case 6  '"print"
                   Call ShellExecute(Me.hWnd, "print", sfile, "", App.Path, 1)

Case 7 'openwith
Dim openDialog As pdOpenSaveDialog
        Set openDialog = New pdOpenSaveDialog
        
        Dim spfile As String
        
        Dim cdFilter As String
        cdFilter = g_Language.TranslateMessage("select program") & " (.exe)|*.exe|"
        cdFilter = cdFilter & g_Language.TranslateMessage("program files") & "|*.exe"
        
        Dim cdTitle As String
        cdTitle = g_Language.TranslateMessage("Load a program to edit this image")
                
        If openDialog.GetOpenFileName(spfile, vbNullString, True, False, cdFilter, 1, "C:\", cdTitle, , frm_Thumbnailer.hWnd) Then
ShellExecute Me.hWnd, "open", spfile, Chr$(34) & sfile & Chr$(34), vbNullString, SW_SHOWNORMAL
End If

'=======================================================================================================================================================
 Case 8 'Copy to..MNUCONTEXTTHUMBNAIL
              JustDoIt1

 Case 9 'move to
                           If MsgBox("Are you sure you want to move to this location ?", vbQuestion + vbYesNo) = vbYes Then
                           JustDoIt1
                    Call ucThumbnailView.Clear
                   ' Call mThumbnail.DeleteFolderThumbnails(sPath)
                 Files.FileDeleteIfExists sfile
                 Call ucFolderView_ChangeAfter(sPath)
                    Call mnuView_Click(0)

                   End If

  
'=======================================================================================================================================================

'=======================================================================================================================================================
Case 10 '-- browse path
                    
        
        ucProgress.Visible = True
        Screen.MousePointer = vbArrowHourglass
                            Dim a As String

        a = PathBrowseDialog(frm_Thumbnailer.hWnd)

                   With ucFolderView
            If (.Path <> cbPath.Text) Then
                .Path = a
            End If
        End With

        '-- Add to recent paths
        Call pvAddPath(a): m_bSkipPath = False

        '-- Add items from path
        Call mThumbnail.UpdateFolder(a)
        
        '-- Items ?
        If (ucThumbnailView.Count) Then
            
            '-- Select first by default
            If (ucThumbnailView.ItemFindState(, [tvSelected]) = -1) Then
                ucThumbnailView.ItemSelected(0) = True
            End If
            
          Else
            ucStatusbar.PanelText(1) = vbNullString
            ucStatusbar.PanelText(2) = vbNullString
            ucStatusbar.PanelText(3) = vbNullString
        End If
        
        '-- Show # of items found
        ucStatusbar.PanelText(3) = Format$(ucThumbnailView.Count, "#,#0 image/s found")
        
        ucProgress.Visible = False
        Screen.MousePointer = vbDefault
'=======================================================================================================================================================
Case 11 '-- convert file
Dim stitle1 As String
  stitle1 = Files.FileGetName(sPath, True)
                        
                       
                        If Files.FileGetExtension(sPath) <> "pdi" Then
                        Loading.LoadFileAsNewImage1 sfile, stitle1, False
                        FileMenu.MenuSaveAs PDImages.GetActiveImage()
                        CanvasManager.FullPDImageUnload PDImages.GetActiveImageID()
frm_Thumbnailer.ZOrder

                        On Error GoTo gh1
                    
                     End If
      



        
gh1:
Loading.LoadFileAsNewImage sfile, sTitle, False
FileMenu.MenuSaveAs PDImages.GetActiveImage()
CanvasManager.FullPDImageUnload PDImages.GetActiveImageID()
frm_Thumbnailer.ZOrder

Exit Sub
Case 12 ' EXIF TOOL
On Error GoTo e
    Dim d$
    d = InputBox("enter directory name: ", "New Directory")
    If Not d = "" Then
        If Right(cbPath.Text, 1) = "\" Then
            cbPath.Text = Left(cbPath.Text, Len(cbPath.Text) - 1)
        End If
        MkDir cbPath.Text & "\" & d
        ucFolderView.Path = cbPath.Text & "\" & d
    End If
    
     Call mThumbnail.UpdateFolder(frm_Thumbnailer.ucFolderView.Path)
        
        '-- Items ?
        If (frm_Thumbnailer.ucThumbnailView.Count) Then
            
            '-- Select first by default
            If (frm_Thumbnailer.ucThumbnailView.ItemFindState(, [tvSelected]) = -1) Then
                frm_Thumbnailer.ucThumbnailView.ItemSelected(0) = True
            End If
            
          Else
            frm_Thumbnailer.ucStatusbar.PanelText(1) = vbNullString
            frm_Thumbnailer.ucStatusbar.PanelText(2) = vbNullString
            frm_Thumbnailer.ucStatusbar.PanelText(3) = vbNullString
        End If
Exit Sub
e:
    MsgBox "Invalid filename !", vbExclamation, "Error": Exit Sub
    
    Case 13 ' EXIF TOOL
   ' If KeyCode = vbKeyDelete Then
    If MsgBox("Are you sure you want to move to this location ?", vbQuestion + vbYesNo) = vbYes Then
    Kill sfile
     If (Not ucFolderView.PathIsRoot) Then
               
                Call ucThumbnailView.Clear
                m_bSkipPath = True
                Call ucFolderView_ChangeAfter(vbNullString)
            End If
        ' Call mThumbnail.UpdateFolder(frm_Thumbnailer.ucFolderView.Path)
End If
'End If
    End Select

End Sub
Private Static Function pvGetFileDateTimeStr(uFileTime As FILETIME) As String
  
  Dim uFT As FILETIME
  Dim uST As SYSTEMTIME

    Call FileTimeToLocalFileTime(uFileTime, uFT)
    Call FileTimeToSystemTime(uFT, uST)
  
    pvGetFileDateTimeStr = pvGetFileDateStr(uST) & " " & pvGetFileTimeStr(uST)
End Function
Private Static Function pvGetFileTimeStr(uSystemTime As SYSTEMTIME) As String
  
  Dim sTime As String * 32
  Dim lLen  As Long
  
    lLen = GetTimeFormat(LOCALE_USER_DEFAULT, LOCALE_NOUSEROVERRIDE, uSystemTime, vbNullString, sTime, 64)
    If (lLen) Then
        pvGetFileTimeStr = Left$(sTime, lLen - 1)
    End If
End Function
Private Static Function pvGetFileDateStr(uSystemTime As SYSTEMTIME) As String
  
  Dim sDate As String * 32
  Dim lLen  As Long
  
    lLen = GetDateFormat(LOCALE_USER_DEFAULT, LOCALE_NOUSEROVERRIDE Or DATE_SHORTDATE, uSystemTime, vbNullString, sDate, 64)
    If (lLen) Then
        pvGetFileDateStr = Left$(sDate, lLen - 1)
    End If
End Function
Private Function FindPath(Parent As String, Optional Child As String, Optional Divider As String = "\") As String
    On Error Resume Next
    If Right$(Parent, 1) = Divider Then Parent = Left$(Parent, Len(Parent) - 1)
    If Left$(Child, 1) = Divider Then Child = Mid$(Child, 2)
    FindPath = Parent & Divider & Child
End Function

Private Function pvCorrectExt(sFileName As String)
    If (Right$(sFileName, 4) <> Right$(m_FileExt, 4)) Then
        sFileName = sFileName & Right$(m_FileExt, 4)
    End If
End Function
Public Function ShellFile(hWnd As Long, strOperation As String, ByVal File As String, WindowStyle As VbAppWinStyle) As Long
'"Open, Print, Explore, Find, Edit, Play, 0&"
    ShellFile = ShellExecute(hWnd, strOperation, File, vbNullString, App.Path, WindowStyle)
End Function
Function fWait1(ByVal lProgID As Long) As Long
    Dim lExitCode As Long, hdlProg As Long
    hdlProg = OpenProcess(PROCESS_ALL_ACCESS, False, lProgID)
    GetExitCodeProcess hdlProg, lExitCode

    Do While lExitCode = STILL_ACTIVE&
        DoEvents
        GetExitCodeProcess hdlProg, lExitCode
    Loop
    
    CloseHandle hdlProg
    fWait1 = lExitCode
End Function


'========================================================================================
' Toolbar
'========================================================================================

Private Sub ucToolbar_ButtonClick(ByVal Button As Long)
    
    Select Case Button
    
        Case 1  '-- Back
            Call mnuGo_Click(0)
      
        Case 2  '-- Forward
            Call mnuGo_Click(1)
      
        Case 3  '-- Up
            Call mnuGo_Click(2)
      
        Case 5  '-- Refresh
            Call mnuView_Click(0)
       
        Case 7  '-- View
            Select Case ucThumbnailView.ViewMode
                Case [tvThumbnail]
                    Call mnuView_Click(3)
                Case [tvDetails]
                    Call mnuView_Click(2)
            End Select
      
        Case 8  '-- Full screen
            
      
        Case 10 '-- Database
                   
            Case 12 '-- Database
    
             '
    

    End Select
End Sub

Private Sub ucToolbar_ButtonDropDown(ByVal Button As Long, ByVal x As Long, ByVal y As Long)
    
    '-- Drop-down menu (view mode)
    Call PopupMenu(mnuViewModeTop, , x, y)
End Sub



'========================================================================================
' Changing path
'========================================================================================

Private Sub ucFolderView_ChangeBefore(ByVal NewPath As String, Cancel As Boolean)

    If (Not m_bEnding And Not ucFolderView.PathIsValid(NewPath)) Then
            
        '-- Invalid path
        Call MsgBox("The specified path is invalid or does not exist.")
        Call SendMessage(cbPath.hWnd, CB_SETCURSEL, 0, ByVal 0)
        Cancel = True
        
      Else
        '-- Stop thumbnailing / Clear
        Call mThumbnail.Cancel
        Call ucThumbnailView.Clear
    End If
End Sub

Private Sub ucFolderView_ChangeAfter(ByVal OldPath As String)
    tmrExploreFolder.Enabled = False
    tmrExploreFolder.Enabled = True
End Sub

Private Sub tmrExploreFolder_Timer()

    tmrExploreFolder.Enabled = False
    
    If (Not m_bEnding) Then
        
        ucProgress.Visible = True
        Screen.MousePointer = vbArrowHourglass
        
        '-- Add to recent paths
        Call pvAddPath(ucFolderView.Path): m_bSkipPath = False

        '-- Add items from path
        Call mThumbnail.UpdateFolder(ucFolderView.Path)
        
        '-- Items ?
        If (ucThumbnailView.Count) Then
            
            '-- Select first by default
            If (ucThumbnailView.ItemFindState(, [tvSelected]) = -1) Then
                ucThumbnailView.ItemSelected(0) = True
            End If
            
          Else
            ucStatusbar.PanelText(1) = vbNullString
            ucStatusbar.PanelText(2) = vbNullString
            ucStatusbar.PanelText(3) = vbNullString
        End If
        
        '-- Show # of items found
        ucStatusbar.PanelText(3) = Format$(ucThumbnailView.Count, "#,#0 image/s found")
        
        ucProgress.Visible = False
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cbPath_GotFocus()
    m_bComboHasFocus = True
End Sub
Private Sub cbPath_LostFocus()
    m_bComboHasFocus = False
End Sub

Private Sub cbPath_Click()
    
    '-- Path selected
    If (SendMessage(cbPath.hWnd, CB_GETDROPPEDSTATE, 0, ByVal 0) = 0) Then
        
        With ucFolderView
            If (.Path <> cbPath.Text) Then
                .Path = cbPath.Text
            End If
        End With
    End If
End Sub

Private Sub cbPath_KeyDown(KeyCode As Integer, Shift As Integer)
    
  Dim lIdx As Long
  
    Select Case KeyCode
    
        '-- New path typed
        Case vbKeyReturn
            
            '-- Check combo's list state (visible)
            If (SendMessage(cbPath.hWnd, CB_GETDROPPEDSTATE, 0, ByVal 0) <> 0) Then
                '-- Get current list box selected (hot) item
                lIdx = SendMessage(cbPath.hWnd, CB_GETCURSEL, 0, ByVal 0)
                If (lIdx <> CB_ERR) Then
                    Call SendMessage(cbPath.hWnd, CB_SETCURSEL, lIdx, ByVal 0)
                End If
            End If
            
            '-- Hide combo's list and force combo click
            Call SendMessage(cbPath.hWnd, CB_SHOWDROPDOWN, 0, ByVal 0)
            Call cbPath_Click
      
        '-- Avoids navigation when list hidden (also avoids mouse-wheel navigation).
        Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown
            
            '-- Preserve manual drop-down
            If (Shift <> vbAltMask) Then
                If (SendMessage(cbPath.hWnd, CB_GETDROPPEDSTATE, 0, ByVal 0) = 0) Then
                    KeyCode = 0
                End If
            End If
    End Select
End Sub



'========================================================================================
' Displaying image / 'full screen' mode
'========================================================================================



Private Sub UpdatePreview(ByVal srcImagePath As String, Optional ByVal forceUpdate As Boolean = False)
    

End Sub






Private Sub ucThumbnailView_ItemClick(ByVal Item As Long)

     Dim srcImagePath As String
     srcImagePath = ucFolderView.Path & ucThumbnailView.ItemText(Item, [tvFileName])

  Dim loadSuccessful As Boolean
  
                       Dim TmpDIB As pdDIB: Set TmpDIB = New pdDIB
  
loadSuccessful = False


        'Use PD's central load function to load a copy of the requested image
        If (LenB(srcImagePath) <> 0) Then loadSuccessful = Loading.QuickLoadImageToDIB(srcImagePath, TmpDIB, False, False)
        
                'If the image load failed, display a placeholder message; otherwise, render the image to the picture box
        If loadSuccessful Then
            ucplayer.CopyDIB TmpDIB, True, True
        Else
            ucplayer.PaintText g_Language.TranslateMessage("previews disabled"), 10!, False, True
        End If
                 
                 
                 
                 
                 
                 

        
              
'
        
        
            ucStatusbar.PanelText(2) = ucplayer.GetWidth & "x" & ucplayer.GetHeight

       


    
    Screen.MousePointer = vbDefault
End Sub
'========================================================================================
' Context menus
'========================================================================================

Private Sub ucThumbnailView_ItemDblClick(ByVal Item As Long)
Dim sPath As String
Dim sTitle As String
Dim SEXT1 As String


                        sPath = ucFolderView.Path & ucThumbnailView.ItemText(ucThumbnailView.ItemFindState(, [tvFocused]), [tvFileName]) & vbNullChar
                        sTitle = Files.FileGetName(sPath, True)
                        SEXT1 = Files.FileGetExtension(sPath)
                       
                        If Files.FileGetExtension(sPath) <> "pdi" Then
                        Loading.LoadFileAsNewImage1 sPath, sTitle, False
                          frm_Thumbnailer.ZOrder

                        On Error GoTo gh
                    
                     End If
      

gh:
Loading.LoadFileAsNewImage sPath, sTitle, False
frm_Thumbnailer.ZOrder
Exit Sub
End Sub
Private Sub ucThumbnailView_ItemRightClick(ByVal Item As Long)
    
    '-- Thumbnail context menu
    Call Me.PopupMenu(mnuContextThumbnailTop, , , , mnuContextThumbnail(0))
End Sub



'========================================================================================
' Navigating
'========================================================================================

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  Const SCROLL_FACTOR As Long = 5
  Dim lFocused        As Long
  Dim bResize         As Boolean
    
    Select Case Shift
    
        Case vbAltMask
    
            If (Not m_bComboHasFocus) Then

                
                    Select Case KeyCode
                                    
                        Case vbKeyLeft  '-- Back
                            Call mnuGo_Click(0)
                        
                        Case vbKeyRight '-- Forward
                            Call mnuGo_Click(1)
                        
                        Case vbKeyUp    '-- Up
                            Call mnuGo_Click(2)
                    End Select
                    KeyCode = 0
                End If
         
      
        Case vbCtrlMask
       
            Select Case KeyCode
                
                Case vbKeyP        '-- Pause/Resume
                    Call mnuContextPreview_Click(2)
                
                Case vbKeyAdd      '-- Pause/Resume
                    Call mnuContextPreview_Click(4)
                
                Case vbKeySubtract '-- Pause/Resume
                    Call mnuContextPreview_Click(5)
                    
                Case vbKeyC        '-- Copy image
                    Call mnuContextPreview_Click(6)
            End Select
            KeyCode = 0
               
        Case Else
            
            Select Case KeyCode
                    
                '-- Navigating thumbnails (full-screen)
                Case vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd
                        
                    If (Not m_bComboHasFocus) Then
                        
                      
        
                            With ucThumbnailView
                                
                                '-- Currently selected
                                lFocused = .ItemFindState(, [tvFocused])
                                
                                Select Case KeyCode
                            
                                    Case vbKeyPageUp   '-- Previous
                                        .ItemSelected(lFocused + 1 * (lFocused > 0)) = True
                            
                                    Case vbKeyPageDown '-- Next
                                        .ItemSelected(lFocused - 1 * (lFocused < .Count - 1)) = True
                            
                                    Case vbKeyHome     '-- First
                                        .ItemSelected(0) = True
                            
                                    Case vbKeyEnd      '-- Last
                                        .ItemSelected(.Count - 1) = True
                                End Select
                                
                                Call .ItemEnsureVisible(.ItemFindState(, [tvFocused]))
                            End With
                            KeyCode = 0
                       
                    End If
                       
                '-- Best fit mode / zoom
                Case vbKeySpace, vbKeyAdd, vbKeySubtract
                        
                    If (Not m_bComboHasFocus) Then
                        
                        With ucplayer
                        
                            Select Case KeyCode
                                
                                Case vbKeySpace    '-- Best fit mode on/off
                                    
                                Case vbKeyAdd      '-- Zoom +
                                    
                                Case vbKeySubtract '-- Zoom -
                            End Select
                            
                            If (bResize) Then
                                
                                
                               
                               
                            End If
                        End With
                        KeyCode = 0
                    End If
                    
                '-- Scrolling preview
                Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
                        
                    If (Not m_bComboHasFocus) Then
                        
                        With ucplayer
                            
                            '-- Scroll <SCROLL_FACTOR> pixels
                            Select Case KeyCode
                                
                                Case vbKeyUp
                                    
                                Case vbKeyDown
                                    
                                Case vbKeyLeft
                                
                                Case vbKeyRight
                            End Select
                        End With
                        KeyCode = 0
                    End If
                         
                '-- Toggle 'full screen'
                Case vbKeyReturn
                    If (Not m_bComboHasFocus) Then
                         KeyCode = 0
                    End If
                    
                '-- Restore combo edit text
                Case vbKeyEscape
                    Call SendMessage(cbPath.hWnd, CB_SETCURSEL, 0, ByVal 0)
                    KeyCode = 0
                    
                '-- Avoid combo drop-down
                Case vbKeyF4
                    KeyCode = 0
            End Select
    End Select
End Sub

'========================================================================================
' Misc
'========================================================================================

Private Sub ucThumbnailView_ColumnResize(ByVal ColumnID As tvColumnIDConstants)
    
    With uAPP_SETTINGS
        .ViewColumnWidth(ColumnID) = ucThumbnailView.ColumnWidth(ColumnID)
    End With
End Sub



'========================================================================================
' Private
'========================================================================================

Private Sub pvUndoPath()

    If (m_PathsPos > 1) Then
        m_PathsPos = m_PathsPos - 1
        
        '-- Update path
        m_bSkipPath = True
        ucFolderView.Path = m_Paths(m_PathsPos)
        
        '-- Update buttons
        Call pvCheckNavigationButtons
    End If
End Sub

Private Sub pvRedoPath()
  
    If (m_PathsPos < m_PathsMax) Then
        m_PathsPos = m_PathsPos + 1
        
        '-- Update path
        m_bSkipPath = True
        ucFolderView.Path = m_Paths(m_PathsPos)
        
        '-- Update buttons
        Call pvCheckNavigationButtons
    End If
End Sub

Private Sub pvAddPath(ByVal sPath As String)
  
 Dim lc   As Long
 Dim lptr As Long
    
    With uAPP_SETTINGS
           
        '-- Add to recent paths list
        For lc = 0 To cbPath.ListCount - 1
            If (sPath = cbPath.List(lc)) Then
                Call cbPath.RemoveItem(lc)
                Exit For
            End If
        Next lc
        If (cbPath.ListCount = 25) Then
            Call cbPath.RemoveItem(cbPath.ListCount - 1)
        End If
        Call cbPath.AddItem(sPath, 0): cbPath.ListIndex = 0
        
        If (m_bSkipPath = False) Then
            
            If (m_PathsPos = m_PathLevels) Then
                '-- Move down items
                lptr = StrPtr(m_Paths(1))
                Call CopyMemory(ByVal VarPtr(m_Paths(1)), ByVal VarPtr(m_Paths(2)), (m_PathLevels - 1) * 4)
                Call CopyMemory(ByVal VarPtr(m_Paths(m_PathLevels)), lptr, 4)
              Else
                '-- One position up
                m_PathsPos = m_PathsPos + 1
                m_PathsMax = m_PathsPos
            End If
            
            '-- Store path
            m_Paths(m_PathsPos) = sPath
        End If
    End With
    
    '-- Update buttons
    Call pvCheckNavigationButtons
End Sub

Private Sub pvCheckNavigationButtons()
    
    '-- Menu buttons
    mnuGo(0).Enabled = (m_PathsPos > 1)
    mnuGo(1).Enabled = (m_PathsPos < m_PathsMax)
    mnuGo(2).Enabled = Not ucFolderView.PathParentIsRoot And Not ucFolderView.PathIsRoot
   
End Sub

Private Sub pvChangeDropDownListHeight(oCombo As ComboBox, ByVal lHeight As Long)
    
    With oCombo
        '-- Drop down list height
        Call MoveWindow(.hWnd, .Left \ Screen.TwipsPerPixelX, .Top \ Screen.TwipsPerPixelY, .Width \ Screen.TwipsPerPixelX, lHeight, 0)
    End With
End Sub

'//

Private Property Get inIDE() As Boolean
   Debug.Assert (IsInIDE())
   inIDE = m_bInIDE
End Property

Private Function IsInIDE() As Boolean
   m_bInIDE = True
   IsInIDE = m_bInIDE
End Function
