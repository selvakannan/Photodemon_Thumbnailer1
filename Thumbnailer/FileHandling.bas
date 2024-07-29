Attribute VB_Name = "FileHandling"
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - September 2002
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au
Option Explicit
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type
'Actions
Public Const FO_MOVE = &H1
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_RENAME = &H4&
'Flags
Public Const FOF_SILENT = &H4
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_SIMPLEPROGRESS = &H100
Public Const FOF_ALLOWUNDO = &H40
Public Sub ShellAction(mSource As String, mDestination As String, mAction As Long, mFlags As Long)
    Dim SHFileOp As SHFILEOPSTRUCT
    mSource = mSource & Chr$(0) & Chr$(0)
    With SHFileOp
        .wFunc = mAction
        .pFrom = mSource
        .pTo = mDestination
        .fFlags = mFlags
    End With
    SHFileOperation SHFileOp
End Sub


Public Function FileExists11(sSource As String) As Boolean
    'Thorogh FileExists function
    If Right(sSource, 2) = ":\" Then
        Dim allDrives As String
        allDrives = Space$(64)
        Call GetLogicalDriveStrings(Len(allDrives), allDrives)
        FileExists11 = InStr(1, allDrives, Left(sSource, 1), 1) > 0
        Exit Function
    Else
        If Not sSource = "" Then
            Dim WFD As WIN32_FIND_DATA
            Dim hFile As Long
            hFile = FindFirstFile(sSource, WFD)
            FileExists11 = hFile <> INVALID_HANDLE_VALUE
            Call FindClose(hFile)
        Else
            FileExists11 = False
        End If
    End If
End Function

Public Function PathOnly(ByVal filepath As String) As String
    Dim temp As String
    temp = Mid$(filepath, 1, InStrRev(filepath, "\"))
    If Right(temp, 1) = "\" And Len(temp) > 3 Then temp = Left(temp, Len(temp) - 1)
    PathOnly = temp
End Function

Public Function FileOnly(ByVal filepath As String) As String
    FileOnly = Mid$(filepath, InStrRev(filepath, "\") + 1)
End Function
Public Function SafeSave(Path As String, Optional ByRef safesavename As String) As String
    'Simple parsing routine to return a unique filename
    Dim mPath As String, mname As String, mTemp As String, mfile As String, mExt As String, m As Integer
    On Error Resume Next
    mPath = Mid$(Path, 1, InStrRev(Path, "\"))
    mname = Mid$(Path, InStrRev(Path, "\") + 1)
    mfile = Left(Mid$(mname, 1, InStrRev(mname, ".")), Len(Mid$(mname, 1, InStrRev(mname, "."))) - 1)
    If mfile = "" Then mfile = mname
    mExt = Mid$(mname, InStrRev(mname, "."))
    mTemp = ""
    Do
        If Not FileExists11(mPath + mfile + mTemp + mExt) Then
            SafeSave = mPath + mfile + mTemp + mExt
            safesavename = mfile + mTemp + mExt
            Exit Do
        End If
        m = m + 1
        mTemp = "(" & m & ")"
    Loop
End Function

