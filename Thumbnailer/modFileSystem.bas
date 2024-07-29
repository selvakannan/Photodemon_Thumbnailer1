Attribute VB_Name = "modFileSystem"
Option Explicit

Private Const SW_SHOWMAXIMIZED = 3
Private Const ArrGrow As Long = 5000
Private Const MaxLong As Long = 2147483647
Private Const MAX_PATH = 260
Private Const MAXDWORD = &HFFFF
Private Const INVALID_HANDLE_VALUE = -1
Private Const LB_SETHORIZONTALEXTENT = &H194
Private Const LB_ADDSTRING = &H180

Enum eFileAttribute
    ATTR_READONLY = &H1
    ATTR_HIDDEN = &H2
    ATTR_SYSTEM = &H4
    ATTR_DIRECTORY = &H10
    ATTR_ARCHIVE = &H20
    ATTR_NORMAL = &H80
    ATTR_TEMPORARY = &H100
End Enum

Enum eSortMethods
    SortNot = 0
    SortByNames = 1
End Enum

Enum eSizeConstants
    BIPerB = 8
    BPERKB = 1024
    KBPerMB = 1024
    MBPerGB = 1024
    GBPerTB = 1024
    TBPerPT = 1024
End Enum

Private Type TextSize
    Width As Long
    Height As Long
End Type

Type tFile
    Name As String
    path As String
    FullName As String
    CreationDate As String
    AccessDate As String
    WriteDate As String
    Size As Currency
    Attr As VbFileAttribute
End Type

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved As Long
    dwReserved1 As Long
    Filename As String * MAX_PATH
    cAlternateFileName As String * 14
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

'Window
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetFocus Lib "user32" () As Long

'Shell
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

'File Stuff
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Time
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

'Image Stuff
Private Declare Function ImageList_Draw Lib "Comctl32" (ByVal himl As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal Flags As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean

'Text Size
Private Declare Function GetTextExtentPoint32 Lib "gdi32" (ByVal hDC As Long, ByVal lpString As String, ByVal cbString As Long, lpSize As TextSize) As Boolean

'Memory stuff
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Public FileSearchCount As Long
Public FilesFound As Long
Public RecurseAmmount As Long
Public CurrentName As String
Public Abort As Boolean

Private Options_DisplayFullName As Boolean
Private Options_DisplayFiles As Boolean
Private Options_DisplayFolders As Boolean
Private Options_MinSize As Long
Private Options_MaxSize As Long
Private Options_DisplayHidden As Boolean
Private Options_DisplayArchive As Boolean
Private Options_DisplayReadOnly As Boolean
Private Options_DisplaySystem As Boolean

Private CURWFD As WIN32_FIND_DATA

Private Sub Compress_RLE(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim x As Long
    Dim RLE_Count As Long
    Dim OutPos As Long
    Dim FileLong As Long
    Dim Char As Long
    Dim OldChar As Long
    
    ReDim OutStream(LBound(ByteArray) To UBound(ByteArray) * 1.33) 'Worst case
    FileLong = UBound(ByteArray)
    OldChar = -1

    For x = LBound(ByteArray) To UBound(ByteArray)
        Char = ByteArray(x)
        
        If Char = OldChar Then
            RLE_Count = RLE_Count + 1
            
            If RLE_Count < 4 Then
                OutStream(OutPos) = Char
                OutPos = OutPos + 1
            End If
            If RLE_Count = 258 Then
                OutStream(OutPos) = RLE_Count - 3
                OutPos = OutPos + 1
                RLE_Count = 0
                OldChar = -1
            End If
        Else
            If RLE_Count > 2 Then
                OutStream(OutPos) = RLE_Count - 3
                OutPos = OutPos + 1
            End If
            
            OutStream(OutPos) = Char
            OutPos = OutPos + 1
            RLE_Count = 1
            OldChar = Char
        End If
    Next
    
    If RLE_Count > 2 Then
        OutStream(OutPos) = RLE_Count - 3
        OutPos = OutPos + 1
    End If
    
    ReDim ByteArray(OutPos + 3)
    CopyMemory ByteArray(OutPos), FileLong, 4
    CopyMemory ByteArray(0), OutStream(0), OutPos
End Sub

Private Sub DeCompress_RLE(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim FileLong As Long
    Dim x As Long
    Dim Char As Long
    Dim OldChar As Long
    Dim RLE_Count As Long
    Dim OutPos As Long
    Dim RRun1 As Boolean
    Dim RRun2 As Boolean
    CopyMemory FileLong, ByteArray(UBound(ByteArray) - 3), 4
    ReDim OutStream(LBound(ByteArray) To FileLong)
    OldChar = -1
    
    For x = LBound(ByteArray) To UBound(ByteArray) - 4
        If RRun1 Then
            If RRun2 Then
                RLE_Count = ByteArray(x)
                If RLE_Count Then
                    FillMemory OutStream(OutPos), RLE_Count, Char
                    OutPos = OutPos + RLE_Count
                End If
                RRun1 = False
                RRun2 = False
                OldChar = -1
            Else
                Char = ByteArray(x)

                OutStream(OutPos) = Char
                OutPos = OutPos + 1
                
                If Char = OldChar Then RRun2 = True Else RRun1 = False: OldChar = Char
            End If
        Else
            Char = ByteArray(x)
            OutStream(OutPos) = Char
            OutPos = OutPos + 1
            
            If Char = OldChar Then RRun1 = True Else OldChar = Char
        End If
    Next
    
    ReDim ByteArray(0 To OutPos - 1)
    CopyMemory ByteArray(0), OutStream(0), OutPos
End Sub

Function Base64Enc(s As String) As String
    Static enc() As Byte
    Dim B() As Byte, Out() As Byte, i As Long, j As Long, L As Long
    
    If (Not val(Not enc)) = 0 Then 'Null-Ptr = not initialized
        enc = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
    End If
    
    L = Len(s): B = StrConv(s, vbFromUnicode)
    ReDim Preserve B(0 To (UBound(B) \ 3) * 3 + 2)
    ReDim Preserve Out(0 To (UBound(B) \ 3) * 4 + 3)
    For i = 0 To UBound(B) - 1 Step 3
        Out(j) = enc(B(i) \ 4)
        j = j + 1
        Out(j) = enc((B(i + 1) \ 16) Or (B(i) And 3) * 16)
        j = j + 1
        Out(j) = enc((B(i + 2) \ 64) Or (B(i + 1) And 15) * 4)
        j = j + 1
        Out(j) = enc(B(i + 2) And 63)
        j = j + 1
    Next
    For i = 1 To i - L
        Out(UBound(Out) - i + 1) = 61
    Next
    Base64Enc = StrConv(Out, vbUnicode)
End Function

Sub NONAME_Encode(Data() As Byte, Key() As Byte)
    Const RANDOMSIZE As Long = 2047
    Dim RandomTable(RANDOMSIZE) As Byte
    Dim i As Long
    Dim KeyPos As Long, RandomSeed As Long
    Dim LKey As Long, UKey As Long, lData As Long, uData As Long
    Dim TotalAdd As Long, CurKey As Long
    
    LKey = LBound(Key)
    UKey = UBound(Key)
    lData = LBound(Data)
    uData = UBound(Data)
    
    For i = LKey To UKey
        RandomSeed = RandomSeed + Key(i)
    Next
    Randomize RandomSeed
    
    For i = LBound(RandomTable) To UBound(RandomTable)
        RandomTable(i) = Int(Rnd * 256)
    Next
    
    KeyPos = LKey
    For i = lData To uData
        CurKey = Key(KeyPos)
        Data(i) = Data(i) Xor CurKey Xor RandomTable(TotalAdd) Xor (TotalAdd And 255)
        TotalAdd = ((RandomTable(CurKey) + TotalAdd) Xor CurKey) And RANDOMSIZE
        If KeyPos >= UKey Then KeyPos = LKey Else KeyPos = KeyPos + 1
    Next
End Sub

Sub NONAME_Encrypt(Data() As Byte, Key() As Byte)
    Call NONAME_Encode(Data, Key)
End Sub

Sub NONAME_Decrypt(Data() As Byte, Key() As Byte)
    Call NONAME_Encode(Data, Key)
End Sub

Function FileGetFirst(path As String, Data As tFile) As Long
    FileGetFirst = FindFirstFile(path & "*", CURWFD)
    DataToFile path, CURWFD, Data
End Function

Function FileGetNext(path As String, hSearch As Long, Data As tFile) As Long
    FileGetNext = FindNextFile(hSearch, CURWFD)
    DataToFile path, CURWFD, Data
End Function

Sub DataToFile(path As String, WFD As WIN32_FIND_DATA, Data As tFile)
    With Data
        'Strings need to be converted
        .Name = StripNulls(WFD.Filename)
        .path = path
        .Size = (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
        .Attr = 0
        If WFD.dwFileAttributes And ATTR_ARCHIVE Then .Attr = .Attr Or vbArchive
        If WFD.dwFileAttributes And ATTR_DIRECTORY Then .Attr = .Attr Or vbDirectory
        If WFD.dwFileAttributes And ATTR_HIDDEN Then .Attr = .Attr Or vbHidden
        If WFD.dwFileAttributes And ATTR_NORMAL Then .Attr = .Attr Or vbNormal
        If WFD.dwFileAttributes And ATTR_READONLY Then .Attr = .Attr Or vbReadOnly
        If WFD.dwFileAttributes And ATTR_SYSTEM Then .Attr = .Attr Or vbSystem
    End With
End Sub

Sub AddItem(TheListbox As ListBox, TheText As String)
    On Error Resume Next
    
    Call SendMessageAny(TheListbox.hwnd, LB_ADDSTRING, 0, ByVal TheText)
    
    Dim TextWidth As Long
    TextWidth = TheListbox.Parent.TextWidth(TheText) + 10
    If TextWidth > CLng(TheListbox.Tag) Then
        TheListbox.Tag = TextWidth
        Call AddHorizontalScrollBar(TheListbox, TextWidth)
    End If
End Sub

Private Function StripNulls(Str As String) As String
    Dim pos As Long
    pos = InStr(1, Str, vbNullChar)
    If pos Then StripNulls = Left$(Str, pos - 1) Else StripNulls = Str
End Function

Function OpenBinaryFile(FilePath As String, Optional bWrite As Boolean) As Integer
    OpenBinaryFile = FreeFile
    
    If bWrite Then
        Open FilePath For Binary Access Write As #OpenBinaryFile
    Else
        Open FilePath For Binary Access Read As #OpenBinaryFile
    End If
End Function

Function OpenRandomFile(FilePath As String, Optional bWrite As Boolean) As Integer
    OpenRandomFile = FreeFile
    
    If bWrite Then
        Open FilePath For Random Access Write As #OpenRandomFile
    Else
        Open FilePath For Random Access Read As #OpenRandomFile
    End If
End Function

Function OpenTextFile(FilePath As String, Optional bWrite As Boolean) As Integer
    OpenTextFile = FreeFile
    
    If bWrite Then
        Open FilePath For Output As #OpenTextFile
    Else
        Open FilePath For Input As #OpenTextFile
    End If
End Function

Sub CloseFile(FileNumber As Integer)
    Close #FileNumber
End Sub

Function StartDoc(DocName As String) As Long
    StartDoc = ShellExecute(0, "Open", DocName, vbNullString, vbNullString, 1)
End Function

Function BrowseWebPage(PageName As String) As Long
    BrowseWebPage = ShellExecute(0, "Open", PageName, vbNullString, vbNullString, vbNullString)
End Function

Function Execute(Filename As String, Optional Windowstate As Long = vbMinimizedFocus) As Boolean
    On Error GoTo Handler
    Call Shell(Filename, Windowstate)
    Execute = True
Handler:
End Function

Function SafeDelete(FilePath As String) As Long
    Dim fileNum As Integer
    Dim CurNum As Long
    'Resize the byte array
    
    For CurNum = 0 To 5
        'Generate a random byte array
        fileNum = OpenBinaryFile(FilePath, True)
            Do Until EOF(fileNum)
                Put #fileNum, , CByte(Rnd * 255)
            Loop
        CloseFile fileNum
    Next
    
    Kill (FilePath)
End Function

Sub SaveBytes(FilePath As String, Bytes() As Byte)
    FileClear FilePath
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open FilePath For Binary Access Write As #fileNum
        Put #fileNum, , Bytes()
    Close #fileNum
End Sub

Sub OpenBytes(FilePath As String, Bytes() As Byte)
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open FilePath For Binary Access Read As #fileNum
        ReDim Bytes(0 To LOF(fileNum) - 1)
        Get #fileNum, , Bytes()
    Close #fileNum
End Sub

Function LoadTextFile(FilePath As String) As String
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open FilePath For Input As #fileNum
        LoadTextFile = Input(LOF(fileNum), #fileNum)
    Close #fileNum
End Function

Sub SaveTextFile(TheString As String, FilePath As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open FilePath For Output As #fileNum
        Print #fileNum, TheString
    Close #fileNum
End Sub


'Folder
Function MakeFolder(FolderName As String) As Boolean
    On Error GoTo Handler
    Call MkDir(FolderName)
    MakeFolder = True
Handler:
End Function

Function DeleteFolder(FolderName As String) As Boolean
    'On Error GoTo Handler

    Kill Right$(FolderName, Len(FolderName) - 1) & "?"
    RmDir FolderName
    DeleteFolder = True
Handler:
End Function

Function DeleteFile(Filename As String) As Boolean
    On Error GoTo Handler
    Kill Filename
    DeleteFile = True
Handler:
End Function

Function FolderExists(FolderName As String) As Boolean
    On Error GoTo Handler
    FolderExists = Dir$(FolderName, vbDirectory) <> vbNullString
Handler:
End Function

Function FileMove(src As String, dest As String) As Boolean
    On Error GoTo Handler
    FileCopy src, dest
    FileMove = True
Handler:
End Function

Function FileMake(Filename As String) As Boolean
    On Error GoTo Handler
    Dim fileNum As Integer
    fileNum = OpenBinaryFile(Filename, True)
    CloseFile fileNum
    FileMake = True
Handler:
End Function

Function FileExists(Filename As String) As Boolean
    On Error GoTo Handler
    FileExists = Dir$(Filename) <> vbNullString
Handler:
End Function

Function FileClear(FilePath As String)
    On Error GoTo Handler
    
    Dim fileNum As Integer
    fileNum = OpenTextFile(FilePath, True)
        Print #fileNum, vbNullString
    CloseFile fileNum
    
    FileClear = True
Handler:
End Function

Function GetDirectoryFolders(Directory As String) As String()
    Dim TheName As String
    Dim Count As Long
    Dim Names() As String
    ReDim Names(ArrGrow)
    
    TheName = Dir$(Directory, vbDirectory)
    
    Do While TheName <> vbNullString
        If TheName <> "." And TheName <> ".." Then
            If (GetAttr(Directory & TheName) And vbDirectory) Then
                If Count > UBound(Names) Then ReDim Preserve Names(Count + ArrGrow)
                Names(Count) = TheName
                Count = Count + 1
            End If
        End If
        TheName = Dir$
    Loop
    ReDim Preserve Names(0 To Count - 1)
    
    GetDirectoryFolders = Names
End Function

Function GetDirectoryFiles(Directory As String) As String()
    Dim TheName As String
    Dim Count As Long
    Dim Names() As String
    ReDim Names(ArrGrow)
    
    TheName = Dir$(Directory, vbNormal)
    
    Do While TheName <> vbNullString
        If TheName <> "." And TheName <> ".." Then
            If Count > UBound(Names) Then ReDim Preserve Names(Count + ArrGrow)
            Names(Count) = TheName
            Count = Count + 1
        End If
        TheName = Dir$
    Loop
    ReDim Preserve Names(0 To Count - 1)
    
    GetDirectoryFiles = Names
End Function

Function GetDirectoryFoldersAndFiles(Directory As String) As tFile()
    Dim Count As Long
    Dim Files() As tFile
    ReDim Files(0 To ArrGrow)
    FileSearchCount = 0
    
    AddFoldersAndFiles Directory, Count, Files
    ReDim Preserve Files(Count - 1)
    GetDirectoryFoldersAndFiles = Files
End Function

Function AddFile(Files() As tFile, Count As Long, Name As String, path As String)
    If Count > UBound(Files) Then ReDim Files(0 To Count + ArrGrow)
    With Files(Count)
        .Name = Name
        .path = path
        .Attr = GetAttr(.path & .Name)
        If (.Attr And vbDirectory) = 0 Then .Size = FileLen(.path & .Name)
    End With
    Count = Count + 1
End Function

Sub AddFilesToListBox(TheListbox As ListBox, Files() As tFile)
    TheListbox.Clear
    
    LockWindowUpdate TheListbox.hwnd
    Dim i As Long
    For i = LBound(Files) To UBound(Files)
        Call TheListbox.AddItem(Files(i).path & Files(i).Name)
    Next
    LockWindowUpdate 0
End Sub

Function FolderFind(Directory As String, Optional Filter As String, Optional eSortMethods As eSortMethods = SortNot, Optional MinSize As Long = 0, Optional MaxSize As Long = MaxLong) As tFile()
    Dim i As Long
    Dim Files() As tFile
    Dim Count As Long
    Dim Count2 As Long
    Dim PrevCount As Long
    Dim StartCount As Long
    Dim Added As Boolean
    Dim FilteredFiles() As tFile
    ReDim Files(ArrGrow)
    ReDim FilteredFiles(ArrGrow)
    If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    Added = True
    FileSearchCount = 0
    Call AddFoldersAndFiles(Directory, Count, Files)
    
    If Count Then
        Do
            StartCount = Count
            For i = PrevCount To Count - 1
                If Files(i).Attr And vbDirectory Then
                    AddFoldersAndFiles Files(i).path & Files(i).Name, Count, Files
                End If
            Next
            
            If PrevCount = Count Then Exit Do
            PrevCount = StartCount
        Loop
    End If
    
    For i = LBound(Files) To Count - 1
        With Files(i)
            If InStr(1, .Name, Filter, vbTextCompare) <> 0 Then
                If .Size >= MinSize And .Size <= MaxSize Then
                    If Count2 > UBound(FilteredFiles) Then ReDim Preserve FilteredFiles(0 To Count2 + ArrGrow)
                    
                    FilteredFiles(Count2) = Files(i)
                    Count2 = Count2 + 1
                End If
            End If
        End With
    Next
    
    If Count2 Then ReDim Preserve FilteredFiles(Count2 - 1)
    
    Select Case eSortMethods
    Case SortByNames
        Call FileSortName(FilteredFiles, LBound(FilteredFiles), UBound(FilteredFiles), -1)
    End Select
    
    If Count2 Then FolderFind = FilteredFiles
End Function

Function AddFoldersAndFiles(Directory As String, Count As Long, Files() As tFile) As Long
    Dim File As tFile
    Dim hSearch As Long

    hSearch = FindFirstFile(Directory & "*", CURWFD) 'FileGetFirst(Directory, File)
    If hSearch = INVALID_HANDLE_VALUE Then Exit Function

    Do
        If File.Name <> "." And File.Name <> ".." And Len(File.Name) <> 0 Then
            DoEvents    'Translate messages
            If Count > UBound(Files) Then ReDim Preserve Files(Count + ArrGrow)
            With Files(Count)
                .path = Directory
                .Attr = File.Attr
                If .Attr And vbDirectory Then
                    .Name = File.Name & "\"
                    CurrentName = .path & .Name
                    .FullName = CurrentName
                Else
                    .Name = File.Name
                    .Size = File.Size
                    .FullName = File.path & File.Name
                End If
    
                Count = Count + 1
            End With
            FileSearchCount = FileSearchCount + 1
        End If
    Loop While FileGetNext(Directory, hSearch, File)
    
    FindClose hSearch
End Function

Function GetRecurseFolders(ByVal Directory As String, Count As Long, Files() As tFile) As Long
    Dim File As tFile
    Dim StartCount As Long, i As Long, hSearch As Long
    StartCount = Count
    
    hSearch = FindFirstFile(Directory & "*", CURWFD)
    If hSearch = INVALID_HANDLE_VALUE Then Exit Function

    Do
        If File.Name <> "." And File.Name <> ".." And File.Name <> vbNullString Then
            DoEvents    'Translate messages
            If Count > UBound(Files) Then ReDim Preserve Files(Count + ArrGrow)
            With Files(Count)
                .path = Directory
                .Attr = File.Attr
                If .Attr And vbDirectory Then
                    .Name = File.Name & "\"
                    CurrentName = .path & .Name
                    .FullName = CurrentName
                Else
                    .Name = File.Name
                    .Size = File.Size
                    .FullName = File.path & File.Name
                End If
            End With
            
            Count = Count + 1
            
            FileSearchCount = FileSearchCount + 1
        End If
    Loop While FileGetNext(Directory, hSearch, File)
    
    
    For i = StartCount To Count - 1
        If Files(i).Attr And vbDirectory Then Call GetRecurseFolders(Files(i).FullName, Count, Files)
    Next
    FindClose hSearch
End Function

Function GetRecurseFoldersListBox(TheListbox As ListBox, ByVal Directory As String, Filter As String, Count As Long, Files() As tFile) As Long
    Dim File As tFile, StartCount As Long, i As Long, hSearch As Long
    StartCount = Count
    
    hSearch = FindFirstFile(Directory & "*", CURWFD)
    If hSearch = INVALID_HANDLE_VALUE Then Exit Function

    Do
        If File.Name <> "." And File.Name <> ".." And File.Name <> vbNullString Then
            DoEvents    'Translate messages
            If Count > UBound(Files) Then ReDim Preserve Files(Count + ArrGrow)
            With Files(Count)
                .path = Directory
                .Attr = File.Attr
                If .Attr And vbDirectory Then
                    .Name = File.Name & "\"
                    CurrentName = .path & .Name
                    .FullName = CurrentName
                Else
                    .Name = File.Name
                    .Size = File.Size
                    .FullName = File.path & File.Name
                End If
            End With
            
            Count = Count + 1
            FileSearchCount = FileSearchCount + 1
        End If
    Loop While FileGetNext(Directory, hSearch, File) <> 0 And (Abort = False)
    FindClose hSearch
    
    'IF THE FILE IS A DIRECTORY THEN ONLY DISPLAY THE FILE IF SHOWDIRECTORY = TRUE
    'IF THE FILE IS A FILE THEN ONLY DISPLAY THE FILE IF SHOWFILE = TRUE
    'IF THE FILE.HIDDEN THEN ONLY DISPLAY THE FILE IF SHOWHIDDEN = TRUE
    'IF THE FILE.READONLY THEN ONLY DISPLAY THE FILE IF SHOWREADONLY = TRUE
    'IF THE FILE.ARCHIVE THEN ONLY DISPLAY THE FILE IF SHOWARCHIVE = TRUE
    
    For i = StartCount To Count - 1
        If (Files(i).Size >= Options_MinSize Or Files(i).Size <= Options_MaxSize) And _
        ((Files(i).Attr And vbDirectory) = 0 Or Options_DisplayFiles) And _
        ((Files(i).Attr And vbDirectory) <> 0 Or Options_DisplayFolders) And _
        ((Files(i).Attr And vbReadOnly) <> 0 Or Options_DisplayReadOnly) And _
        ((Files(i).Attr And vbArchive) <> 0 Or Options_DisplayArchive) And _
        ((Files(i).Attr And vbHidden) <> 0 Or Options_DisplayHidden) And _
        ((Files(i).Attr And vbSystem) <> 0 Or Options_DisplaySystem) And _
        InStr(1, Files(i).Name, Filter, vbTextCompare) <> 0 Then
            Call AddItem(TheListbox, Files(i).FullName)
            FilesFound = FilesFound + 1
        End If
        If Files(i).Attr And vbDirectory Then GetRecurseFoldersListBox TheListbox, Files(i).FullName, Filter, Count, Files
NextItem:
    Next
End Function

Function FileSearch(ListBox As ListBox, Directory As String, Filter As String, Optional MinSize As Long = 0, Optional MaxSize As Long = -1, _
Optional ShowFiles As Boolean = True, Optional ShowFolders As Boolean = True, _
Optional ShowReadOnly As Boolean = True, Optional ShowArchive As Boolean = True, Optional ShowHidden As Boolean = True, _
Optional ShowSystem As Boolean = True _
) As tFile()
    'Our variables
    Dim Files() As tFile
    Dim Count As Long
    
    'Start the search
    Call searchStart(Files)
    
    'Clear the list box
    ListBox.Clear
    
    'Make sure the Directory is right
    If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    
    'Set the module level variables for no OUT OF STACK SPACE ERRORS
    Options_MinSize = MinSize
    If MaxSize = -1 Then Options_MaxSize = MaxLong Else Options_MaxSize = MaxSize
    Options_DisplayFiles = Not ShowFiles
    Options_DisplayFolders = Not ShowFolders
    Options_DisplayReadOnly = Not ShowReadOnly
    Options_DisplayHidden = Not ShowHidden
    Options_DisplayArchive = Not ShowArchive
    Options_DisplaySystem = Not ShowSystem
    
    'Recursivly get folders and files
    Call GetRecurseFoldersListBox(ListBox, Directory, Filter, Count, Files)
    
    'Resize the files to only how much we found, remove the padding
    ReDim Preserve Files(0 To Count - 1)
    
    'Return the files we found
    FileSearch = Files
End Function

Function UpOne(Dir As String) As String
    Dim pos As Long
    pos = InStrRev(Dir, "\", Len(Dir) - 1)
    If pos Then UpOne = Left$(Dir, pos) Else UpOne = Dir
End Function

Sub ResetOptions()
    Options_DisplayFiles = True
    Options_DisplayFolders = True
    Options_DisplayFullName = True
End Sub

Private Sub searchStart(Files() As tFile)
    ReDim Files(ArrGrow)
    Abort = False
    FileSearchCount = 0
    FilesFound = 0
End Sub

Function GetFoldersAndFiles(ByVal Directory As String) As tFile()
    Dim Files() As tFile
    Dim Count As Long
    ReDim Files(0)
    FileSearchCount = 0
    If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    Call GetRecurseFolders(Directory, Count, Files)
    ReDim Preserve Files(0 To Count - 1)
    GetFoldersAndFiles = Files
End Function

Function AddHorizontalScrollBar(TheListbox As ListBox, Pixels As Long) As Long
    AddHorizontalScrollBar = SendMessage(TheListbox.hwnd, LB_SETHORIZONTALEXTENT, Pixels, 0&)
End Function

Private Sub FileSortName(arr() As tFile, lLbound As Long, lUbound As Long, Direction As Long)
    If lUbound <= lLbound Then Exit Sub
    
    Static Buffer As tFile
    Dim Compare As String
    Dim CurHigh As Long
    Dim CurLow As Long

    CurLow = lLbound
    CurHigh = lUbound
    Compare = arr((lLbound + lUbound) \ 2).FullName

    Do While CurLow <= CurHigh
        Do While StrComp(arr(CurLow).FullName, Compare, vbTextCompare) = Direction And CurLow <> lUbound: CurLow = CurLow + 1: Loop
        Do While StrComp(Compare, arr(CurHigh).FullName, vbTextCompare) = Direction And CurHigh <> lLbound: CurHigh = CurHigh - 1: Loop

        If CurLow <= CurHigh Then
            Buffer = arr(CurLow)
            arr(CurLow) = arr(CurHigh)
            arr(CurHigh) = Buffer
            CurLow = CurLow + 1
            CurHigh = CurHigh - 1
        End If
    Loop
    
    If lLbound < CurHigh Then FileSortName arr(), lLbound, CurHigh, Direction
    If CurLow < lUbound Then FileSortName arr(), CurLow, lUbound, Direction
End Sub


