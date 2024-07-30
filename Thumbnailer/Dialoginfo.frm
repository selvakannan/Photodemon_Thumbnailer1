VERSION 5.00
Begin VB.Form frm_info 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Informations.."
   ClientHeight    =   8685
   ClientLeft      =   2760
   ClientTop       =   4050
   ClientWidth     =   19230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   19230
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAbout 
      Caption         =   "about"
      Height          =   375
      Left            =   17880
      TabIndex        =   61
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Height          =   1695
      Left            =   10920
      TabIndex        =   56
      Top             =   3840
      Width           =   8055
      Begin VB.TextBox txtDestination 
         Height          =   525
         Left            =   120
         TabIndex        =   58
         Top             =   1080
         Width           =   7815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   360
         Width           =   7695
      End
      Begin VB.Label Label13 
         Caption         =   "Destination File to operations:"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label12 
         Caption         =   "Source Folder path..."
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Flags"
      Height          =   1695
      Left            =   14880
      TabIndex        =   50
      Top             =   5760
      Width           =   2175
      Begin VB.CheckBox ChFlag 
         Caption         =   "Allow Undo"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   55
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox ChFlag 
         Caption         =   "Rename on Collision"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   54
         Top             =   600
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox ChFlag 
         Caption         =   "Silent"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   53
         Top             =   840
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox ChFlag 
         Caption         =   "Show Progress"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   52
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox ChFlag 
         Caption         =   "No Confirmation"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   51
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Action"
      Height          =   975
      Left            =   17040
      TabIndex        =   47
      Top             =   6480
      Width           =   1935
      Begin VB.OptionButton OptAction 
         Caption         =   "Rename"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   49
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton OptAction 
         Caption         =   "Delete"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   48
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   17760
      TabIndex        =   46
      Top             =   5640
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   3600
      Left            =   15000
      Pattern         =   "*.TXT;*.RTF;*.DOC;*.INI"
      TabIndex        =   44
      Top             =   120
      Width           =   3975
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   11400
      TabIndex        =   43
      Top             =   480
      Width           =   3375
   End
   Begin VB.ListBox filelist 
      Height          =   2205
      Left            =   11280
      TabIndex        =   42
      Top             =   5760
      Width           =   3375
   End
   Begin VB.TextBox txtdir 
      Height          =   405
      Left            =   5280
      TabIndex        =   41
      Top             =   7560
      Width           =   5895
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   11400
      TabIndex        =   40
      Top             =   120
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Computer information"
      Height          =   780
      Left            =   0
      TabIndex        =   37
      Top             =   7320
      Width           =   5025
      Begin VB.TextBox txt_Name 
         Height          =   330
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "Unknown"
         Top             =   270
         Width           =   3090
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cumputer Name:"
         Height          =   195
         Left            =   135
         TabIndex        =   39
         Top             =   315
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Drive information:  "
      Height          =   4110
      Left            =   7560
      TabIndex        =   15
      Top             =   3240
      Width           =   3210
      Begin VB.DriveListBox drv_Drive 
         Height          =   315
         Left            =   135
         TabIndex        =   26
         Top             =   270
         Width           =   2895
      End
      Begin VB.TextBox txt_VolumeName 
         Height          =   330
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   675
         Width           =   1275
      End
      Begin VB.TextBox txt_DriveType 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "ERROR"
         Top             =   1125
         Width           =   915
      End
      Begin VB.TextBox txt_FileSystem 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Unknown"
         Top             =   1440
         Width           =   1050
      End
      Begin VB.TextBox txt_ID 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Unknown"
         Top             =   1755
         Width           =   1005
      End
      Begin VB.TextBox txt_SectorsPerCluster 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "ERROR"
         Top             =   2070
         Width           =   1005
      End
      Begin VB.TextBox txt_BytesPerSector 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "ERROR"
         Top             =   2385
         Width           =   960
      End
      Begin VB.TextBox txt_TotalSpace 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "ERROR"
         Top             =   2700
         Width           =   915
      End
      Begin VB.TextBox txt_FreeSpace 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "ERROR"
         Top             =   3015
         Width           =   1005
      End
      Begin VB.TextBox txt_Clusters 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "ERROR"
         Top             =   3330
         Width           =   1005
      End
      Begin VB.TextBox txt_Sectors 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "ERROR"
         Top             =   3645
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   1035
         TabIndex        =   36
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Drive Type:"
         Height          =   195
         Left            =   675
         TabIndex        =   35
         Top             =   1125
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "File System:"
         Height          =   195
         Left            =   675
         TabIndex        =   34
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Serial ID:"
         Height          =   195
         Left            =   870
         TabIndex        =   33
         Top             =   1755
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sectors per cluster:"
         Height          =   195
         Left            =   150
         TabIndex        =   32
         Top             =   2070
         Width           =   1365
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bytes per sector:"
         Height          =   195
         Left            =   315
         TabIndex        =   31
         Top             =   2385
         Width           =   1185
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Space:"
         Height          =   195
         Left            =   570
         TabIndex        =   30
         Top             =   2700
         Width           =   915
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Free Space:"
         Height          =   195
         Left            =   615
         TabIndex        =   29
         Top             =   3015
         Width           =   870
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Clusters:"
         Height          =   195
         Left            =   495
         TabIndex        =   28
         Top             =   3330
         Width           =   1005
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Sectors:"
         Height          =   195
         Left            =   510
         TabIndex        =   27
         Top             =   3645
         Width           =   990
      End
   End
   Begin PhotoDemon.pdPictureBox ucplayer 
      Height          =   3015
      Left            =   7560
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5318
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Copy to clipboard"
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton Command13 
      Caption         =   "paste from clipboard"
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Edit with notepad"
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Replace Text"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Find  Text"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "make LIST"
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   5760
      TabIndex        =   8
      Text            =   "DIR C:"
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   120
      Width           =   7335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "os version"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "system info"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "system metrics"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Username:"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "VersionInfo"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NumberOrfProcessors"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "System Requirements"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label lblfile 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   0
      TabIndex        =   45
      Top             =   8160
      Width           =   19095
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufile1 
         Caption         =   "open"
         Index           =   0
      End
      Begin VB.Menu mnufile1 
         Caption         =   "close"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frm_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const ModuleName As String = "SystemInfo"

Private Type LARGE_INTEGER
    lowPart As Long
    highPart As Long
End Type
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Private Type SYSTEMINFO
    sysDriveName As String
    sysDriveType As String
    sysFileSystem As String
    sysSerialID As Long
    sysSectorsPerCluster As Double
    sysBytesPerSector As Double
    sysTotalSpace As Double
    sysFreeSpace As Double
    sysTotalClusters As Double
    sysTotalSectors As Double
    sysComputerName As String
    sysPsysicalMemory As Long
    sysVirtualMemory As Long
    sysMemoryLoad As Long
    sysPageFile As Long
End Type

Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32.dll" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, ByRef lpFreeBytesAvailableToCaller As LARGE_INTEGER, ByRef lpTotalNumberOfBytes As LARGE_INTEGER, ByRef lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Boolean
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExA" (ByVal lpFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Boolean
Private Enum EXTENDED_NAME_FORMAT
    NameUnknown = 0
    NameFullyQualifiedDN = 1
    NameSamCompatible = 2
    NameDisplay = 3
    NameUniqueId = 6
    NameCanonical = 7
    NameUserPrincipal = 8
    NameCanonicalEx = 9
    NameServicePrincipal = 10
End Enum
Dim m_Action As Long
Dim txtsource As String

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL As Long = 1

Private Declare Function GetUserNameEx Lib "secur32.dll" Alias "GetUserNameExA" (ByVal NameFormat As EXTENDED_NAME_FORMAT, ByVal lpNameBuffer As String, ByRef nSize As Long) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Const SYNCHRONIZE = &H100000
Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Sub GetNativeSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
'In general section
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Const SM_CXSCREEN = 0 'X Size of screen
Const SM_CYSCREEN = 1 'Y Size of Screen
Const SM_CXVSCROLL = 2 'X Size of arrow in vertical scroll bar.
Const SM_CYHSCROLL = 3 'Y Size of arrow in horizontal scroll bar
Const SM_CYCAPTION = 4 'Height of windows caption
Const SM_CXBORDER = 5 'Width of no-sizable borders
Const SM_CYBORDER = 6 'Height of non-sizable borders
Const SM_CXDLGFRAME = 7 'Width of dialog box borders
Const SM_CYDLGFRAME = 8 'Height of dialog box borders
Const SM_CYVTHUMB = 9 'Height of scroll box on horizontal scroll bar
Const SM_CXHTHUMB = 10 ' Width of scroll box on horizontal scroll bar
Const SM_CXICON = 11 'Width of standard icon
Const SM_CYICON = 12 'Height of standard icon
Const SM_CXCURSOR = 13 'Width of standard cursor
Const SM_CYCURSOR = 14 'Height of standard cursor
Const SM_CYMENU = 15 'Height of menu
Const SM_CXFULLSCREEN = 16 'Width of client area of maximized window
Const SM_CYFULLSCREEN = 17 'Height of client area of maximized window
Const SM_CYKANJIWINDOW = 18 'Height of Kanji window
Const SM_MOUSEPRESENT = 19 'True is a mouse is present
Const SM_CYVSCROLL = 20 'Height of arrow in vertical scroll bar
Const SM_CXHSCROLL = 21 'Width of arrow in vertical scroll bar
Const SM_DEBUG = 22 'True if deugging version of windows is running
Const SM_SWAPBUTTON = 23 'True if left and right buttons are swapped.
Const SM_CXMIN = 28 'Minimum width of window
Const SM_CYMIN = 29 'Minimum height of window
Const SM_CXSIZE = 30 'Width of title bar bitmaps
Const SM_CYSIZE = 31 'height of title bar bitmaps
Const SM_CXMINTRACK = 34 'Minimum tracking width of window
Const SM_CYMINTRACK = 35 'Minimum tracking height of window
Const SM_CXDOUBLECLK = 36 'double click width
Const SM_CYDOUBLECLK = 37 'double click height
Const SM_CXICONSPACING = 38 'width between desktop icons
Const SM_CYICONSPACING = 39 'height between desktop icons
Const SM_MENUDROPALIGNMENT = 40 'Zero if popup menus are aligned to the left of the memu bar item. True if it is aligned to the right.
Const SM_PENWINDOWS = 41 'The handle of the pen windows DLL if loaded.
Const SM_DBCSENABLED = 42 'True if double byte characteds are enabled
Const SM_CMOUSEBUTTONS = 43 'Number of mouse buttons.
Const SM_CMETRICS = 44 'Number of system metrics
Const SM_CLEANBOOT = 67 'Windows 95 boot mode. 0 = normal, 1 = safe, 2 = safe with network
Const SM_CXMAXIMIZED = 61 'default width of win95 maximised window
Const SM_CXMAXTRACK = 59 'maximum width when resizing win95 windows
Const SM_CXMENUCHECK = 71 'width of menu checkmark bitmap
Const SM_CXMENUSIZE = 54 'width of button on menu bar
Const SM_CXMINIMIZED = 57 'width of rectangle into which minimised windows must fit.
Const SM_CYMAXIMIZED = 62 'default height of win95 maximised window
Const SM_CYMAXTRACK = 60 'maximum width when resizing win95 windows
Const SM_CYMENUCHECK = 72 'height of menu checkmark bitmap
Const SM_CYMENUSIZE = 55 'height of button on menu bar
Const SM_CYMINIMIZED = 58 'height of rectangle into which minimised windows must fit.
Const SM_CYSMCAPTION = 51 'height of windows 95 small caption
Const SM_MIDEASTENABLED = 74 'Hebrw and Arabic enabled for windows 95
Const SM_NETWORK = 63 'bit o is set if a network is present. Const SM_SECURE = 44 'True if security is present on windows 95 system
Const SM_SLOWMACHINE = 73 'true if machine is too slow to run win95.
Private m_CRC As clsCRC
Private INIReadOnly As Boolean
Private INIFileName As String
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Sub OKButton_Click()
Unload Me
End Sub
Private Sub WriteINI(ByVal Section As String, ByVal Key As String, ByVal Text As String, Optional ByVal AlternateINIFile As String)
On Error GoTo Error:

If AlternateINIFile = "" Then AlternateINIFile = INIFileName

If INIReadOnly = True Then Exit Sub
Call writeprivateprofilestring(Section, Key, Text, ByVal AlternateINIFile)

Exit Sub
Error:
End Sub

Private Sub Command1_Click()
Text4.Text = ""
   Text4.Text = "PhotoDemon thumbnailer unicode version" & App.Major & "." & App.Minor & vbCrLf & vbCrLf & "PhotoDemon- 2024 and system required info:" & vbCrLf & "tested ok upto windows  os 10 version 22h2 os build 19045.4598 " & vbCrLf & "not for windows  os 10 version 22h2 os build 19045.4651"

End Sub

Private Sub Command10_Click()
Dim s As String
  If Text4.SelLength > 0 Then s = Text4.SelText Else s = "Find Text"
  ShowFind Me, Text4, FR_SHOWHELP, s
End Sub

Private Sub Command11_Click()
 Dim s As String
  If Text4.SelLength > 0 Then s = Text4.SelText Else s = "Find Text"
  ShowFind Me, Text4, FR_SHOWHELP, s, True, "Replace Text"


End Sub

Private Sub Command12_Click()
      FormOnTop Me, False

Dim fso As FileSystemObject

 Dim f As String
    On Error Resume Next
If Right(App.Path, 1) = "\" Then
    f = App.Path & "settings.txt"
Else
    f = App.Path & "\settings.txt"
End If
Open f For Output As #1
Print #1, Text4.Text
Close #1
    Shell "notepad.exe """ & f & """", vbNormalFocus
End Sub

Private Sub Command13_Click()
Dim nYesNo As Integer

If Not Clipboard.GetText = "" Then
    If Not Text4.Text = "" Then
        nYesNo = MsgBox("Clear paste area first?", vbYesNo + vbDefaultButton1 + vbQuestion, "Clear?")
    Else
        nYesNo = vbYes
    End If
    
    If nYesNo = vbYes Then
        Text4.Text = Clipboard.GetText
    Else
        Text4.Text = Clipboard.GetText & vbCrLf & vbCrLf & Text4.Text
    End If
End If
End Sub

Private Sub Command14_Click()
 With Clipboard
    .Clear
    .SetText Text4.Text, vbCFText
  End With
End Sub

Private Sub Command15_Click()

End Sub


Private Sub Command3_Click()
 Dim ret As Long
    IsWow64Process GetCurrentProcess, ret
    If ret = 0 Then
        MsgBox "This application is not running on an x86 emulator for a 64-bit computer!"
    Else
        Dim SysInfo64 As SYSTEM_INFO
        GetNativeSystemInfo SysInfo64
       Text4.Text = "Number of processors on your 64-bit system: " + CStr(SysInfo64.dwNumberOrfProcessors)
    End If
End Sub
Public Function GetWinVersion() As String
    Dim Ver As Long, WinVer As Long
    Ver = GetVersion()
    WinVer = Ver And &HFFFF&
    'retrieve the windows version
    GetWinVersion = Format((WinVer Mod 256) + ((WinVer \ 256) / 100), "Fixed")
End Function
Private Sub Command4_Click()
Dim OSInfo As OSVERSIONINFO, PID As String
Dim ret As String
    'Set the graphical mode to persistent
    Me.AutoRedraw = True
    'Set the structure size
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    'Get the Windows version
    ret = GetVersionEx(OSInfo)
    'Chack for errors
    If ret = 0 Then MsgBox "Error Getting Version Information": Exit Sub
    'Print the information to the form
    Select Case OSInfo.dwPlatformId
        Case 0
            PID = "Windows 32s "
        Case 1
            PID = "Windows 95/98"
        Case 2
            PID = "Windows NT "
    End Select
    Text4.Text = ""
    Text4.Text = "OS: " + PID & vbCrLf
    Text4.Text = Text4.Text + "Win version:" + Str$(OSInfo.dwMajorVersion) + "." + LTrim(Str(OSInfo.dwMinorVersion)) & vbCrLf
    Text4.Text = Text4.Text + "Build: " + Str(OSInfo.dwBuildNumber)
End Sub

Private Sub Command5_Click()
 Dim sBuffer As String, ret As Long
    sBuffer = String(256, 0)
    ret = Len(sBuffer)
    If GetUserNameEx(NameSamCompatible, sBuffer, ret) <> 0 Then
        Text4.Text = "Username: " + Left$(sBuffer, ret)
    Else
        Text4.Text = "Error while retrieving the username"
    End If
End Sub

Private Sub Command6_Click()
Text4.Text = ""
    Text4.Text = "Number of mouse buttons:" + Str$(GetSystemMetrics(SM_CMOUSEBUTTONS))
    Text4.Text = Text4.Text + "Screen X:" + Str$(GetSystemMetrics(SM_CXSCREEN)) & vbCrLf
    Text4.Text = Text4.Text + "Screen Y:" + Str$(GetSystemMetrics(SM_CYSCREEN)) & vbCrLf
    Text4.Text = Text4.Text + "Height of windows caption:" + Str$(GetSystemMetrics(SM_CYCAPTION)) & vbCrLf
    Text4.Text = Text4.Text + "Width between desktop icons:" + Str$(GetSystemMetrics(SM_CXICONSPACING)) & vbCrLf
    Text4.Text = Text4.Text + "Maximum width when resizing a window:" + Str$(GetSystemMetrics(SM_CYMAXTRACK)) & vbCrLf
    Text4.Text = Text4.Text + "Is machine is too slow to run windows?" + Str$(GetSystemMetrics(SM_SLOWMACHINE))
    
    
End Sub

Private Sub Command7_Click()
Dim sInfo As SYSTEM_INFO

    'Set the graphical mode to persistent
    Text4.Text = ""
    GetSystemInfo sInfo
    'Print it to the form
  Text4.Text = "Number of procesor:" + Str$(sInfo.dwNumberOrfProcessors)
   Text4.Text = Text4.Text + "Processor:" + Str$(sInfo.dwProcessorType) & vbCrLf
   Text4.Text = Text4.Text + "Low memory address:" + Str$(sInfo.lpMinimumApplicationAddress) & vbCrLf
  Text4.Text = Text4.Text + "High memory address:" + Str$(sInfo.lpMaximumApplicationAddress)
End Sub

Private Sub Command8_Click()
 Text4.Text = "Windows version: " + GetWinVersion
End Sub
Function SHELLGETTEXT(PROGRAM As String, Optional SHOCMD As Long = vbMinimizedNoFocus) As String
Dim sFile As String
Dim hFile As String
Dim ILENGTH As Long
Dim PID As Long
Dim hProcess As Long

sFile = Space(1024)
ILENGTH = GetTempFileName(Environ("TEMP"), "OUT", 0, sFile)
sFile = Left(sFile, ILENGTH)
PID = Shell(Environ("COMSPEC") & " /C" & PROGRAM & ">" & sFile, SHOCMD)
hProcess = OpenProcess(SYNCHRONIZE, True, PID)
WaitForSingleObject hProcess, -1
CloseHandle hProcess
hFile = FreeFile
Open sFile For Binary As #hFile
SHELLGETTEXT = Input$(LOF(hFile), hFile)
Close #hFile
Kill sFile
End Function

Private Sub Command9_Click()
Text4.Text = SHELLGETTEXT(Text5.Text)

End Sub

Private Sub File1_Click()
    txtsource = IIf(Right(File1.Path, 1) = "\", File1.Path, File1.Path & "\") & File1.Filename
        txtDestination.Text = IIf(Right(File1.Path, 1) = "\", File1.Path, File1.Path & "\") & File1.Filename
 Dim srcImagePath As String
     srcImagePath = IIf(Right(File1.Path, 1) = "\", File1.Path, File1.Path & "\") & File1.Filename
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
  

    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
Dir1.Path = "C:\"
   Me.Move Screen.Width - Me.Width, (Screen.Height - Me.Height) / 2
      FormOnTop Me, True

    m_Action = 1
    File1_Click

Caption = "Find/Replace dialogs"
   Command10.Caption = "Find dialog"
   Command11.Caption = "Replace dialog"
   Dim sFile3 As String, sText As String
   sFile3 = "c:\windows\win.ini"
   Open sFile3 For Binary As #1
   sText = Space$(LOF(1))
   Get #1, , sText
   Close #1
   Text4 = sText

  Dim sFile As String
        sFile = App.Path & "\" & App.EXEName & ".exe"
            Dim s As String
       Dim R As String

       Dim sCRC As String
       Dim sexte As String

       Set m_CRC = New clsCRC
           m_CRC.Algorithm = Crc32
       Dim uWFD        As WIN32_FIND_DATA
    
           sCRC = Hex(m_CRC.CalculateFile(sFile))
    Text4.Text = "1.Pragram Name:-" & App.EXEName & vbCrLf & "2.Pragram Description:-" & App.FileDescription & vbCrLf & "3.Pragram Title:-" & App.Title & vbCrLf & "4.Pragram version:- " & Updates.GetPhotoDemonVersion() & vbCrLf & "5.Pragram maker:- " & App.CompanyName
    Text4.Text = Text4.Text & vbCrLf & "6.File Folder:- " & App.Path & vbCrLf & "7.program path:- " & sFile
    Text4.Text = Text4.Text & vbCrLf & "8.CRC Checksum:- " & sCRC & vbCrLf & "9.File Name:- " & App.EXEName & ".exe" & vbCrLf & "10.Extension:- *.exe" & vbCrLf


End Sub
Private Sub FormOnTop(f_form As Form, i As Boolean)
   
   If i = True Then 'On
      SetWindowPos f_form.hWnd, -1, 0, 0, 0, 0, &H2 + &H1
   Else      'off
      SetWindowPos f_form.hWnd, -2, 0, 0, 0, 0, &H2 + &H1
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
      FormOnTop Me, False

End Sub

Private Sub Text4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        Text4.Enabled = False
        Text4.Enabled = True
        PopupMenu mnufile1
    End If
End Sub


Private Sub drv_Drive_Change()
On Error GoTo e
    Get_Info Left(drv_Drive, 2) & "\"
Exit Sub
e:
    Dim i
    i = MsgBox(Err.Description, vbCritical Or vbAbortRetryIgnore, "Error: " & Err)
    Select Case i
    Case vbAbort: Exit Sub
    Case vbIgnore: Resume Next
    Case vbRetry: Resume
    End Select
End Sub
Private Function win_SystemInfo(sInfo As SYSTEMINFO, sDrive As String)
'
' Simple system info
'
On Error GoTo e
    ' This will get all the info, just select the drive
    Dim dType&, SectorsPerCluster&, BytesPerSector&, TotalClusters&, FreeClusters&
    Dim FreeSpace#, TotalSpace#, TotalSectors&, FreeSectors&
    Dim FreeSpaceEx#, liFS As LARGE_INTEGER
    Dim TotalSpaceEx#, liTS As LARGE_INTEGER
    Dim TotalFreeSpaceEx#, liTFS As LARGE_INTEGER
    Dim vName$, vID&, vFile$, vFileFlags&, mcl&
    Dim cName$, mStatus As MEMORYSTATUS, mPageFile!, cPos!
    
    cPos = 1
    dType = GetDriveType(sDrive)
    Select Case dType
        Case DRIVE_UNKNOWN: sInfo.sysDriveType = "Unknown"
        Case DRIVE_NO_ROOT_DIR: sInfo.sysDriveType = "No Root Dir"
        Case DRIVE_REMOVABLE: sInfo.sysDriveType = "Removable"
        Case DRIVE_FIXED: sInfo.sysDriveType = "Fixed"
        Case DRIVE_REMOTE: sInfo.sysDriveType = "Remote"
        Case DRIVE_RAMDISK: sInfo.sysDriveType = "Ram Disk"
        Case DRIVE_CDROM: sInfo.sysDriveType = "CD Rom"
    End Select
    
    cPos = 2
    GetDiskFreeSpace sDrive, SectorsPerCluster, BytesPerSector, FreeClusters, TotalClusters
    ' Calculate the drive space
    TotalSectors = SectorsPerCluster * TotalClusters
    FreeSectors = SectorsPerCluster * FreeClusters
    FreeSpace = (FreeSectors * BytesPerSector) / 1048576
    TotalSpace = (TotalSectors * BytesPerSector) / 1048576
    ' if old win, then use the normal api
    If Not win_Function_Exist("kernel32.dll", "GetDiskFreeSpaceExA") Then
        cPos = 2.1
        sInfo.sysFreeSpace = FreeSpace
        sInfo.sysTotalClusters = TotalSpace
    ' Else, use the advanced stuff
    Else
        cPos = 2.2
        GetDiskFreeSpaceEx sDrive, liFS, liTS, liTFS
        ' You just need to convert the high and low values
        FreeSpaceEx = win_C32to64(liFS.lowPart, liFS.highPart)
        TotalSpaceEx = win_C32to64(liTS.lowPart, liTS.highPart)
        sInfo.sysFreeSpace = FreeSpaceEx
        sInfo.sysTotalSpace = TotalSpaceEx
    End If
    sInfo.sysTotalClusters = TotalClusters
    sInfo.sysTotalSectors = TotalSectors
    sInfo.sysBytesPerSector = BytesPerSector
    sInfo.sysSectorsPerCluster = SectorsPerCluster
    
    cPos = 3
    vName = String(256, 0) ' Fill it, if you will use Len
    vFile = String(256, 0)
    GetVolumeInformation sDrive, vName, Len(vName), vID, mcl, vFileFlags, vFile, Len(vFile)
    sInfo.sysDriveName = vName
    sInfo.sysFileSystem = vFile
    sInfo.sysSerialID = vID
    
    cPos = 4
    cName = Space(32)
    GetComputerName cName, 32
    sInfo.sysComputerName = cName
    
    ' The memo status may be false on some computers
    cPos = 5
    GlobalMemoryStatus mStatus
    sInfo.sysPsysicalMemory = (CDbl(mStatus.dwAvailPhys) * 100) / mStatus.dwTotalPhys
    sInfo.sysVirtualMemory = (CDbl(mStatus.dwAvailVirtual) * 100) / mStatus.dwTotalVirtual
    sInfo.sysPageFile = (CDbl(mStatus.dwAvailPageFile) * 100) / mStatus.dwTotalPageFile
    sInfo.sysMemoryLoad = mStatus.dwMemoryLoad
Exit Function
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "win_systeminfo: " & cPos
    Resume Next
End Function


Public Sub Log_Append(Text As String)
'
' Appends a text at "PROGRAM.LOG" file in appdir
' Use err_raise for errors
'
    
    ' Use this stuff for putting text in error log
    Open App.Path & "\PROGRAM.LOG" For Append As #2
        Print #2, "DATE: " & Date & ",TIME: " & time & " - " & Text
    Close #2
e:
    Exit Sub
End Sub

Public Sub Err_Raise(Err_Num As String, Err_Description As String, Err_Module As String, Err_Function As String)
'
' Writes a detailed error information to a log file
' Use err object for getting info
'
    Log_Append "ERROR " & Err_Num & " - " & Err_Module & "\" & Err_Function & " ::: " & Err_Description
End Sub

Public Function win_C32to64(ByVal lLo As Long, ByVal lHi As Long) As Double
'
' Gets a Double from LargeInt
'
    
    Dim dLo As Double
    Dim dHi As Double
    
    If lLo < 0 Then
        dLo = (2 ^ 32) + lLo
    Else
        dLo = lLo
    End If
    If lHi < 0 Then
        dHi = (2 ^ 32) + lHi
    Else
        dHi = lHi
    End If
    
    win_C32to64 = (dLo + (dHi * (2 ^ 32)))
End Function

Public Function win_Function_Exist(sModule As String, sFunction As String) As Boolean
'
' Checks if spec. function exists.
' be sure to add .dll at the end :)
'

    Dim hHandle As Long
    hHandle = GetModuleHandle(sModule)
    If hHandle = 0 Then
        hHandle = LoadLibraryEx(sModule, 0&, 0&)
        If GetProcAddress(hHandle, sFunction) = 0 Then
            win_Function_Exist = False
        Else
            win_Function_Exist = True
        End If
        FreeLibrary hHandle
    Else
        If GetProcAddress(hHandle, sFunction) <> 0 Then
            win_Function_Exist = True
        End If
    End If
End Function


Public Sub Get_Info(Drive As String)
    Dim sInfo As SYSTEMINFO
    win_SystemInfo sInfo, Drive ' Get all the stuffs
    
    txt_Name.Text = sInfo.sysComputerName
    txt_VolumeName.Text = sInfo.sysDriveName
    txt_FileSystem.Text = sInfo.sysFileSystem
    txt_ID.Text = sInfo.sysSerialID
    txt_SectorsPerCluster = sInfo.sysSectorsPerCluster
    txt_BytesPerSector.Text = sInfo.sysBytesPerSector
    txt_TotalSpace.Text = Round(sInfo.sysTotalSpace / 1048576, 1) & " MB"
    txt_FreeSpace.Text = Round(sInfo.sysFreeSpace / 1048576, 1) & " MB"
    txt_Clusters.Text = sInfo.sysTotalClusters
    txt_Sectors.Text = sInfo.sysTotalSectors
    txt_DriveType.Text = sInfo.sysDriveType
End Sub

Private Sub Form_Activate()
On Error GoTo e
    drv_Drive.Drive = "c:\"
    Get_Info Left(drv_Drive, 2) & "\"
Exit Sub
e:
    Dim i
    i = MsgBox(Err.Description, vbCritical Or vbAbortRetryIgnore, "Error: " & Err)
    Select Case i
    Case vbAbort: Unload Me
    Case vbIgnore: Resume Next
    Case vbRetry: Resume
    End Select
End Sub


Private Sub tmr_Refresh_Timer()
On Error GoTo e
    ' Refresh it every 2 seconds
    Get_Info Left(drv_Drive, 2) & "\"
Exit Sub
e:
    Dim i
    i = MsgBox(Err.Description, vbCritical Or vbAbortRetryIgnore, "Error: " & Err)
    Select Case i
    Case vbAbort: Exit Sub
    Case vbIgnore: Resume Next
    Case vbRetry: Resume
    End Select
End Sub


Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1_Click

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
On Error GoTo GH
            Call ShellExecute(Me.hWnd, vbNullString, File1.Filename, vbNullString, "C:\", SW_SHOWNORMAL)
GH:
Exit Sub

End Sub


Private Sub filelist_DblClick()
On Error GoTo GH
            Call ShellExecute(Me.hWnd, vbNullString, lblfile.Caption, vbNullString, "C:\", SW_SHOWNORMAL)
GH:
Exit Sub
End Sub

Private Sub cmdAbout_Click()
Dim temp As String
temp = "This is a simple demo of using the API to manage files. It is by no means"
temp = temp & " a 'File Manager' and Im sure if you fiddle with the paths enough, you'll"
temp = temp & " get around the simple error handling I have employed and cause some"
temp = temp & " errors. Choose some files that are dispensible in order to test this"
temp = temp & " demo. It was put together in order to answer questions from many coders"
temp = temp & " new to VB in the Discussion Forum. I hope it is of some help to those"
temp = temp & " folks."
MsgBox temp
End Sub

Private Sub CmdOK_Click()
    Dim z As Long, IsSpecial As Boolean, FolderDelete As Boolean
    'First some checks
    
    'Not deleting so need a destination
    If m_Action <> 3 Then
        If Len(txtDestination.Text) = 0 Then
            MsgBox "You need to specify a destination"
            Exit Sub
        End If
    Else
        If GetAttr(txtsource) = vbDirectory Then FolderDelete = True
    End If
    'If we're performing an action on a folder remove traling backslash
    If Right(txtsource, 1) = "\" Then txtsource = Left(txtsource, Len(txtsource) - 1)
    'Does the file we're acting on exist?
    If Not FileExists11(txtsource) Then
        MsgBox "File not found"
        Exit Sub
    End If
    'Dont mess with drives!
    If Len(txtsource) < 4 Then
        MsgBox "Not a good idea to perform actions on drives!"
        Exit Sub
    End If
   
    If IsSpecial Then
        MsgBox "Not a good idea to perform actions on Special Folders!"
        Exit Sub
    End If
    'Make sure paths match when renaming
    If m_Action = 4 Then
        If LCase(PathOnly(txtsource)) <> LCase(PathOnly(txtDestination.Text)) Then
            MsgBox "When renaming, the paths must be the same - only the name changes."
            txtDestination.SetFocus
            Exit Sub
        End If
    End If
    'Ok, do it!
    ShellAction txtsource, txtDestination.Text, m_Action, GetFlags
    'If we deleted a folder go up one level
    'Refresh the view
    File1.Refresh
    'If we didn't delete then txtDestination.Text should now exist
    If m_Action <> 3 Then
        If Not FileExists11(txtDestination.Text) Then GoTo woops
    Else
    'If we did delete then txtsource should not exist
        If FileExists11(txtsource) Then GoTo woops
    End If
    With frm_Thumbnailer
  If (Not .ucFolderView.PathIsRoot) Then
                Call frm_Thumbnailer.ucThumbnailView.Clear
                Call frm_Thumbnailer.ucFolderView_ChangeAfter(vbNullString)
            End If
   End With
    
    Exit Sub
woops:
    MsgBox "Woops! Something went wrong!"
End Sub



Private Function GetFlags() As Long
    'Create SHFileOperation flags according to checkboxes
    Dim mFlags As Long, z As Long
    For z = 0 To ChFlag.Count - 1
        If ChFlag(z).Value = 1 Then
            Select Case z
                Case 0
                    mFlags = mFlags Or FOF_ALLOWUNDO
                Case 1
                    mFlags = mFlags Or FOF_RENAMEONCOLLISION
                Case 2
                    mFlags = mFlags Or FOF_SILENT
                Case 3
                    mFlags = mFlags Or FOF_SIMPLEPROGRESS
                Case 4
                    mFlags = mFlags Or FOF_NOCONFIRMATION
            End Select
        End If
    Next
    GetFlags = mFlags
End Function


Private Sub OptAction_Click(Index As Integer)
    'Adjust the SHFileOperation action variable according to the Option selected
    m_Action = Index
End Sub





