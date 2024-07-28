VERSION 5.00
Begin VB.Form Dialoginfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Informations.."
   ClientHeight    =   7665
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "System Memory informations"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   2655
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   4560
      Width           =   10215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "System Requirements"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2880
      Width           =   11295
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   11295
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   11295
   End
End
Attribute VB_Name = "Dialoginfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Text4.Text = ""
   Text4.Text = "PhotoDemon thumbnailer unicode version" & App.Major & "." & App.Minor & vbCrLf & vbCrLf & "PhotoDemon- 2024 and system required info:" & vbCrLf & "tested ok upto windows  os 10 version 22h2 os build 19045.4598 " & vbCrLf & "not for windows  os 10 version 22h2 os build 19045.4651"

End Sub

Private Sub Command2_Click()
frm_SystemInfo.Show
End Sub
