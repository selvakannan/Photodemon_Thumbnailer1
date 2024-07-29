VERSION 5.00
Begin VB.Form frmSource 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SHFileOperation Demo"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   3615
      Left            =   4440
      TabIndex        =   14
      Top             =   0
      Width           =   6255
      Begin VB.FileListBox File1 
         Height          =   3210
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   9360
      TabIndex        =   13
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   9360
      TabIndex        =   12
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Action"
      Height          =   975
      Left            =   2400
      TabIndex        =   9
      Top             =   6480
      Width           =   1935
      Begin VB.OptionButton OptAction 
         Caption         =   "Delete"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton OptAction 
         Caption         =   "Rename"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Flags"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   2175
      Begin VB.CheckBox ChFlag 
         Caption         =   "No Confirmation"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox ChFlag 
         Caption         =   "Show Progress"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox ChFlag 
         Caption         =   "Silent"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox ChFlag 
         Caption         =   "Rename on Collision"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox ChFlag 
         Caption         =   "Allow Undo"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   10575
      Begin VB.TextBox Text1 
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   480
         Width           =   10215
      End
      Begin VB.TextBox txtDestination 
         Height          =   525
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   10215
      End
      Begin VB.Label Label1 
         Caption         =   "Source Folder path..."
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Destination File to operations:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
      End
   End
   Begin PhotoDemon.pdPictureBox ucplayer 
      Height          =   3255
      Left            =   240
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5741
   End
End
Attribute VB_Name = "frmSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - September 2002
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au
Option Explicit
'Action to perform
Dim m_Action As Long
Dim txtsource As String
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
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

Private Sub Dir1_Change()
    File1.Path = App.Path
    File1_Click
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
   Me.Move Screen.Width - Me.Width, (Screen.Height - Me.Height) / 2
      FormOnTop Me, True

    m_Action = 1
    File1_Click
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
Private Sub FormOnTop(f_form As Form, i As Boolean)
   
   If i = True Then 'On
      SetWindowPos f_form.hwnd, -1, 0, 0, 0, 0, &H2 + &H1
   Else      'off
      SetWindowPos f_form.hwnd, -2, 0, 0, 0, 0, &H2 + &H1
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
      FormOnTop Me, False

End Sub

Private Sub OptAction_Click(Index As Integer)
    'Adjust the SHFileOperation action variable according to the Option selected
    m_Action = Index
End Sub
