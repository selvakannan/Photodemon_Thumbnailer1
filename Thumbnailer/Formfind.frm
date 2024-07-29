VERSION 5.00
Begin VB.Form Formfind 
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   240
      Width           =   6975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   3360
      Width           =   1815
   End
End
Attribute VB_Name = "Formfind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim s As String
  If Text1.SelLength > 0 Then s = Text1.SelText Else s = "Find Text"
  ShowFind Me, Text1, FR_SHOWHELP, s
End Sub

Private Sub Command2_Click()
  Dim s As String
  If Text1.SelLength > 0 Then s = Text1.SelText Else s = "Find Text"
  ShowFind Me, Text1, FR_SHOWHELP, s, True, "Replace Text"
End Sub

Private Sub Form_Load()
   Caption = "Find/Replace dialogs"
   Command1.Caption = "Find dialog"
   Command2.Caption = "Replace dialog"
   Dim sFile As String, sText As String
   sFile = "c:\windows\win.ini"
   Open sFile For Binary As #1
   sText = Space$(LOF(1))
   Get #1, , sText
   Close #1
   Text1 = sText
End Sub
