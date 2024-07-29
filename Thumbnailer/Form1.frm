VERSION 5.00
Begin VB.Form HELPLIST1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Manager"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   11355
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   6735
   End
   Begin VB.TextBox txtdir 
      Height          =   405
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3855
   End
   Begin VB.ListBox filelist 
      Height          =   7665
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3855
   End
   Begin VB.DirListBox Dir1 
      Height          =   4140
      Left            =   4080
      TabIndex        =   1
      Top             =   600
      Width           =   6855
   End
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   4080
      Pattern         =   "*.TXT;*.RTF;*.DOC;*.INI"
      TabIndex        =   0
      Top             =   4800
      Width           =   6855
   End
   Begin VB.Label lblfile 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   8400
      Width           =   10935
   End
End
Attribute VB_Name = "HELPLIST1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Function FillFileList()
filelist.Clear
If Dir1.path = "C:\" Then
 For x = 0 To Dir1.ListCount - 1
 newdir = "<" + Right(Dir1.List(x), Len(Dir1.List(x)) - Len(Dir1.path)) + ">"
 filelist.AddItem newdir
 Next x
Else
 filelist.AddItem "<..>"
 For x = 0 To Dir1.ListCount - 1
 newdir = "<" + Right(Dir1.List(x), Len(Dir1.List(x)) - Len(Dir1.path) - 1) + ">"
 filelist.AddItem newdir
 Next x
End If
For x = 1 To File1.ListCount - 1
filelist.AddItem File1.List(x)
Next x
txtdir.Text = Dir1.path
End Function


Private Sub Dir1_Change()
File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
End Sub

Private Sub File1_DblClick()
On Error GoTo GH
            Call ShellExecute(Me.hwnd, vbNullString, File1.Filename, vbNullString, "C:\", SW_SHOWNORMAL)
GH:
Exit Sub

End Sub

Private Sub filelist_Click()
If filelist.Text = "<..>" Then
 lblfile.Caption = ""
 Exit Sub
End If
x = InStr(1, filelist.Text, "<")
If x = 1 Then
 y = InStr(1, filelist.Text, ">")
 z = Mid(filelist.Text, x + 1, y - 2)
 If Dir1.path = "C:\" Then
  lblfile.Caption = Dir1.path + z + "\"
 Else
  lblfile.Caption = Dir1.path + "\" + z + "\"
 End If
Else
 If Dir1.path = "C:\" Then
  lblfile.Caption = Dir1.path + filelist.Text
 Else
  lblfile.Caption = Dir1.path + "\" + filelist.Text
 End If
End If
End Sub

Private Sub filelist_DblClick()
x = InStr(1, filelist.Text, "<")
If x = 1 Then
 y = InStr(1, filelist.Text, ">")
 z = Mid(filelist.Text, x + 1, y - 2)
 If Dir1.path = "C:\" Then
  Dir1.path = Dir1.path + z
 Else
  Dir1.path = Dir1.path + "\" + z
 End If
 FillFileList
End If
On Error GoTo GH
            Call ShellExecute(Me.hwnd, vbNullString, lblfile.Caption, vbNullString, "C:\", SW_SHOWNORMAL)
GH:
Exit Sub
End Sub

Private Sub Form_Load()
Dir1.path = "C:\"
FillFileList
End Sub

