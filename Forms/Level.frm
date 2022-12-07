VERSION 5.00
Begin VB.Form Level 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   ClientHeight    =   10650
   ClientLeft      =   -15
   ClientTop       =   -105
   ClientWidth     =   19320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   19320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Back"
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7200
      TabIndex        =   2
      Top             =   6240
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Medium"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7200
      TabIndex        =   1
      Top             =   4320
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Easy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7200
      TabIndex        =   0
      Top             =   2400
      Width           =   4575
   End
End
Attribute VB_Name = "Level"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Easy.Show
Unload Me
End Sub

Private Sub Command2_Click()
Medium.Show
Unload Me
End Sub

Private Sub Command3_Click()
Hard.Show
Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BackColor = &HC000&
End Sub
