VERSION 5.00
Begin VB.Form frm_AddDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Directories"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "frm_AddDir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   2745
      TabIndex        =   5
      Top             =   1530
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Done"
      Height          =   420
      Left            =   2745
      TabIndex        =   4
      Top             =   1035
      Width           =   1500
   End
   Begin VB.TextBox txt_Name 
      Height          =   330
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   2625
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Remove"
      Height          =   420
      Left            =   2745
      TabIndex        =   2
      Top             =   540
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   420
      Left            =   2745
      TabIndex        =   1
      Top             =   45
      Width           =   1500
   End
   Begin VB.ListBox lst_Dirs 
      Height          =   2205
      Left            =   45
      TabIndex        =   0
      Top             =   450
      Width           =   2625
   End
End
Attribute VB_Name = "frm_AddDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const ModuleName As String = "frm_adddir"
Public ifSelected           As Integer


Private Sub Command1_Click()
On Error GoTo e
    Dim I&
    With txt_Name
        If Not .Text = "" Then
            For I = 0 To lst_Dirs.ListCount - 1
                If lst_Dirs.List(I) = .Text Then
                    MsgBox "That directory is allready listed !", vbExclamation, "Error"
                    Exit Sub
                End If
            Next I
            lst_Dirs.AddItem .Text
        End If
    End With
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "command1": Resume Next
End Sub
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

Private Sub Command2_Click()
On Error Resume Next
    lst_Dirs.RemoveItem lst_Dirs.ListIndex
End Sub

Private Sub Command3_Click()
On Error GoTo e
    Dim I&
    For I = 0 To lst_Dirs.ListCount - 1
        MkDir frm_Thumbnailer.ucFolderView.Path & "\" & lst_Dirs.List(I)
    Next I
  '-- Add items from path
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
    Unload Me
Exit Sub
e:
    MsgBox "Invalid directory name or directory exists ! Check all directories for possible mistakes.", vbCritical, "Error"
    Exit Sub
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

