VERSION 5.00
Begin VB.Form Howtoplay 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10680
   ScaleWidth      =   19320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "4. Select a Level to Continue."
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
      Height          =   735
      Left            =   600
      TabIndex        =   4
      Top             =   6480
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   1095
      Left            =   3840
      Picture         =   "Howtoplay.frx":0000
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "3. After creating account select player and click on play."
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
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   4800
      Width           =   10335
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   3960
      Picture         =   "Howtoplay.frx":08A9
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2. If a player does not exist in the list create a Account by clicking on this icon."
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
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   3000
      Width           =   13095
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   2760
      Picture         =   "Howtoplay.frx":0C2C
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Select a player from the list."
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
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Image Image4 
      Height          =   4335
      Left            =   3480
      Picture         =   "Howtoplay.frx":1FE8
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   3855
   End
End
Attribute VB_Name = "Howtoplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Main.Show
Unload Me
End Sub

