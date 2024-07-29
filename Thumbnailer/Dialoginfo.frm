VERSION 5.00
Begin VB.Form Dialoginfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Informations.."
   ClientHeight    =   7890
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command11 
      Caption         =   "Command11"
      Height          =   375
      Left            =   10200
      TabIndex        =   12
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Find"
      Height          =   375
      Left            =   10200
      TabIndex        =   11
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      Caption         =   "make LIST"
      Height          =   495
      Left            =   8640
      TabIndex        =   10
      Top             =   7080
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   8640
      TabIndex        =   9
      Text            =   "DIR C:"
      Top             =   6600
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
      Height          =   6375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   120
      Width           =   11175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "os version"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "system info"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "system metrics"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Username:"
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "VersionInfo"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NumberOrfProcessors"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "System Memory informations"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "System Requirements"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6720
      Width           =   1935
   End
End
Attribute VB_Name = "Dialoginfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
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
Private Declare Function GetUserNameEx Lib "secur32.dll" Alias "GetUserNameExA" (ByVal NameFormat As EXTENDED_NAME_FORMAT, ByVal lpNameBuffer As String, ByRef nSize As Long) As Long
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
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

Private Declare Function IsWow64Process Lib "kernel32" (ByVal HPROCESS As Long, ByRef Wow64Process As Long) As Long
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

Private Sub OKButton_Click()
Unload Me
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

Private Sub Command3_Click()
 Dim Ret As Long
    IsWow64Process GetCurrentProcess, Ret
    If Ret = 0 Then
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
Dim Ret As String
    'Set the graphical mode to persistent
    Me.AutoRedraw = True
    'Set the structure size
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    'Get the Windows version
    Ret = GetVersionEx(OSInfo)
    'Chack for errors
    If Ret = 0 Then MsgBox "Error Getting Version Information": Exit Sub
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
 Dim sBuffer As String, Ret As Long
    sBuffer = String(256, 0)
    Ret = Len(sBuffer)
    If GetUserNameEx(NameSamCompatible, sBuffer, Ret) <> 0 Then
        Text4.Text = "Username: " + Left$(sBuffer, Ret)
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
Dim SInfo As SYSTEM_INFO

    'Set the graphical mode to persistent
    Text4.Text = ""
    GetSystemInfo SInfo
    'Print it to the form
  Text4.Text = "Number of procesor:" + Str$(SInfo.dwNumberOrfProcessors)
   Text4.Text = Text4.Text + "Processor:" + Str$(SInfo.dwProcessorType) & vbCrLf
   Text4.Text = Text4.Text + "Low memory address:" + Str$(SInfo.lpMinimumApplicationAddress) & vbCrLf
  Text4.Text = Text4.Text + "High memory address:" + Str$(SInfo.lpMaximumApplicationAddress)
End Sub

Private Sub Command8_Click()
 Text4.Text = "Windows version: " + GetWinVersion
End Sub
Function SHELLGETTEXT(PROGRAM As String, Optional SHOCMD As Long = vbMinimizedNoFocus) As String
Dim sFile As String
Dim HFILE As String
Dim ILENGTH As Long
Dim PID As Long
Dim HPROCESS As Long

sFile = Space(1024)
ILENGTH = GetTempFileName(Environ("TEMP"), "OUT", 0, sFile)
sFile = Left(sFile, ILENGTH)
PID = Shell(Environ("COMSPEC") & " /C" & PROGRAM & ">" & sFile, SHOCMD)
HPROCESS = OpenProcess(SYNCHRONIZE, True, PID)
WaitForSingleObject HPROCESS, -1
CloseHandle HPROCESS
HFILE = FreeFile
Open sFile For Binary As #HFILE
SHELLGETTEXT = Input$(LOF(HFILE), HFILE)
Close #HFILE
Kill sFile
End Function

Private Sub Command9_Click()
Text4.Text = SHELLGETTEXT(Text5.Text)

End Sub

Private Sub Form_Load()

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
        sFile = App.path & "\" & App.EXEName & ".exe"
            Dim s As String
       Dim R As String

       Dim sCRC As String
       Dim sexte As String

       Set m_CRC = New clsCRC
           m_CRC.Algorithm = Crc32
       Dim uWFD        As WIN32_FIND_DATA
    
           sCRC = Hex(m_CRC.CalculateFile(sFile))
    Text4.Text = "1.Pragram Name:-" & App.EXEName & vbCrLf & "2.Pragram Description:-" & App.FileDescription & vbCrLf & "3.Pragram Title:-" & App.Title & vbCrLf & "4.Pragram version:- " & Updates.GetPhotoDemonVersion() & vbCrLf & "5.Pragram maker:- " & App.CompanyName
    Text4.Text = Text4.Text & vbCrLf & "6.File Folder:- " & App.path & vbCrLf & "7.program path:- " & sFile
    Text4.Text = Text4.Text & vbCrLf & "8.CRC Checksum:- " & sCRC & vbCrLf & "9.File Name:- " & App.EXEName & ".exe" & vbCrLf & "10.Extension:- *.exe" & vbCrLf


End Sub
