VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   16800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Height          =   8535
      Left            =   9600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   6975
   End
   Begin PhotoDemon.pdPictureBox ucplayer 
      Height          =   8895
      Left            =   0
      Top             =   240
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   15690
   End
   Begin VB.Menu MNUFILE1 
      Caption         =   "FILE"
      Visible         =   0   'False
      Begin VB.Menu MNUFILE 
         Caption         =   "PRINT"
         Index           =   0
      End
      Begin VB.Menu MNUFILE 
         Caption         =   "FILE"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
        Text4.Enabled = False
        Text4.Enabled = True
        PopupMenu MNUFILE1
    End If
End Sub
