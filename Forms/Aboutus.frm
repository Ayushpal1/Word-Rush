VERSION 5.00
Begin VB.Form AboutUs 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10650
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   19320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   19320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Ayush : Programmer and Designer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   750
      Left            =   3495
      TabIndex        =   8
      Top             =   8760
      Width           =   5265
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Prasad : UI and Debugging"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   750
      Left            =   3495
      TabIndex        =   7
      Top             =   7800
      Width           =   3945
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "About creators"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   750
      Left            =   2500
      TabIndex        =   6
      Top             =   6960
      Width           =   2500
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Version : 1.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   750
      Left            =   3500
      TabIndex        =   5
      Top             =   6000
      Width           =   2500
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Genre : Puzzle, Trivia and Educational"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   750
      Left            =   3495
      TabIndex        =   4
      Top             =   4995
      Width           =   6105
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mode : Single Player"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   750
      Left            =   3495
      TabIndex        =   3
      Top             =   3840
      Width           =   3105
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Title : Word Rush"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   750
      Left            =   3495
      TabIndex        =   2
      Top             =   2400
      Width           =   2625
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "About Game"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   750
      Left            =   2500
      TabIndex        =   1
      Top             =   1320
      Width           =   2500
   End
End
Attribute VB_Name = "AboutUs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
