VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_SystemInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Information"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   Icon            =   "frm_SystemInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmr_Refresh 
      Interval        =   2000
      Left            =   2790
      Top             =   1890
   End
   Begin VB.Frame Frame2 
      Caption         =   "Computer information"
      Height          =   780
      Left            =   3375
      TabIndex        =   22
      Top             =   45
      Width           =   3825
      Begin VB.TextBox txt_Name 
         Height          =   330
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Unknown"
         Top             =   270
         Width           =   2010
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cumputer Name:"
         Height          =   195
         Left            =   135
         TabIndex        =   23
         Top             =   315
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Drive information:  "
      Height          =   4110
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   3210
      Begin VB.TextBox txt_Sectors 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "ERROR"
         Top             =   3645
         Width           =   1005
      End
      Begin VB.TextBox txt_Clusters 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "ERROR"
         Top             =   3330
         Width           =   1005
      End
      Begin VB.TextBox txt_FreeSpace 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "ERROR"
         Top             =   3015
         Width           =   1005
      End
      Begin VB.TextBox txt_TotalSpace 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "ERROR"
         Top             =   2700
         Width           =   915
      End
      Begin VB.TextBox txt_BytesPerSector 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "ERROR"
         Top             =   2385
         Width           =   960
      End
      Begin VB.TextBox txt_SectorsPerCluster 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "ERROR"
         Top             =   2070
         Width           =   1005
      End
      Begin VB.TextBox txt_ID 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Unknown"
         Top             =   1755
         Width           =   1005
      End
      Begin VB.TextBox txt_FileSystem 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Unknown"
         Top             =   1440
         Width           =   1050
      End
      Begin VB.TextBox txt_DriveType 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "ERROR"
         Top             =   1125
         Width           =   915
      End
      Begin VB.TextBox txt_VolumeName 
         Height          =   330
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   675
         Width           =   1275
      End
      Begin VB.DriveListBox drv_Drive 
         Height          =   315
         Left            =   135
         TabIndex        =   1
         Top             =   270
         Width           =   2895
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Sectors:"
         Height          =   195
         Left            =   510
         TabIndex        =   20
         Top             =   3645
         Width           =   990
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Clusters:"
         Height          =   195
         Left            =   495
         TabIndex        =   18
         Top             =   3330
         Width           =   1005
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Free Space:"
         Height          =   195
         Left            =   615
         TabIndex        =   16
         Top             =   3015
         Width           =   870
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Space:"
         Height          =   195
         Left            =   570
         TabIndex        =   14
         Top             =   2700
         Width           =   915
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bytes per sector:"
         Height          =   195
         Left            =   315
         TabIndex        =   12
         Top             =   2385
         Width           =   1185
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sectors per cluster:"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   2070
         Width           =   1365
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Serial ID:"
         Height          =   195
         Left            =   870
         TabIndex        =   8
         Top             =   1755
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "File System:"
         Height          =   195
         Left            =   675
         TabIndex        =   6
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Drive Type:"
         Height          =   195
         Left            =   675
         TabIndex        =   4
         Top             =   1125
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   1035
         TabIndex        =   2
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "System Status"
      Height          =   3255
      Left            =   3375
      TabIndex        =   25
      Top             =   900
      Width           =   3825
      Begin MSComctlLib.ProgressBar pb_Space 
         Height          =   285
         Left            =   135
         TabIndex        =   27
         ToolTipText     =   "Free Space"
         Top             =   495
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar pb_PMemory 
         Height          =   285
         Left            =   135
         TabIndex        =   29
         ToolTipText     =   "Free Space"
         Top             =   1080
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar pb_VMemory 
         Height          =   285
         Left            =   135
         TabIndex        =   31
         ToolTipText     =   "Free Space"
         Top             =   1665
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar pb_MemoryLoad 
         Height          =   285
         Left            =   135
         TabIndex        =   33
         ToolTipText     =   "Free Space"
         Top             =   2250
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar pb_Total 
         Height          =   285
         Left            =   135
         TabIndex        =   35
         ToolTipText     =   "Free Space"
         Top             =   2835
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lbl_Total 
         AutoSize        =   -1  'True
         Caption         =   "System Resourses"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   34
         Top             =   2610
         Width           =   1560
      End
      Begin VB.Label lbl_MemLoad 
         AutoSize        =   -1  'True
         Caption         =   "Memory Load"
         Height          =   195
         Left            =   135
         TabIndex        =   32
         Top             =   2025
         Width           =   960
      End
      Begin VB.Label lbl_VMem 
         AutoSize        =   -1  'True
         Caption         =   "Virtual Memory"
         Height          =   195
         Left            =   135
         TabIndex        =   30
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label lbl_PMem 
         AutoSize        =   -1  'True
         Caption         =   "Phsysical Memory"
         Height          =   195
         Left            =   135
         TabIndex        =   28
         Top             =   855
         Width           =   1260
      End
      Begin VB.Label lbl_Drive 
         AutoSize        =   -1  'True
         Caption         =   "Drive Free Space:"
         Height          =   195
         Left            =   135
         TabIndex        =   26
         Top             =   270
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frm_SystemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const ModuleName As String = "SystemInfo"

Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
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

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub drv_Drive_Change()
On Error GoTo e
    Get_Info Left(drv_Drive, 2) & "\"
Exit Sub
e:
    Dim I
    I = MsgBox(Err.Description, vbCritical Or vbAbortRetryIgnore, "Error: " & Err)
    Select Case I
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
        FreeSpaceEx = win_C32to64(liFS.LowPart, liFS.HighPart)
        TotalSpaceEx = win_C32to64(liTS.LowPart, liTS.HighPart)
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
        Print #2, "DATE: " & Date & ",TIME: " & Time & " - " & Text
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
    
    pb_Space.Max = CLng(sInfo.sysTotalSpace / 1048576)
    pb_Space.Value = CLng(sInfo.sysFreeSpace / 1048576)
    pb_PMemory.Value = sInfo.sysPsysicalMemory
    pb_VMemory.Value = sInfo.sysVirtualMemory
    pb_MemoryLoad = sInfo.sysMemoryLoad
    pb_Total.Value = (pb_PMemory.Value + pb_VMemory.Value + pb_MemoryLoad.Value + sInfo.sysPageFile) / 4
    lbl_PMem.Caption = "Phsysical Memory: " & CLng(pb_PMemory.Value) & " % Free"
    lbl_VMem.Caption = "Virtual Memory: " & CLng(pb_VMemory.Value) & " % Free"
    lbl_MemLoad.Caption = "Memory Load: " & CLng(pb_MemoryLoad.Value) & " %"
    lbl_Total.Caption = "Total: " & CLng(pb_Total.Value) & " % Free"
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
    Dim I
    I = MsgBox(Err.Description, vbCritical Or vbAbortRetryIgnore, "Error: " & Err)
    Select Case I
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
    Dim I
    I = MsgBox(Err.Description, vbCritical Or vbAbortRetryIgnore, "Error: " & Err)
    Select Case I
    Case vbAbort: Exit Sub
    Case vbIgnore: Resume Next
    Case vbRetry: Resume
    End Select
End Sub


