VERSION 5.00
Begin VB.Form Scorecard 
   BackColor       =   &H80000008&
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
   Begin VB.ListBox List4 
      Height          =   4935
      Left            =   14040
      TabIndex        =   4
      Top             =   2760
      Width           =   1000
   End
   Begin VB.ListBox List3 
      Height          =   4935
      Left            =   9600
      TabIndex        =   3
      Top             =   2760
      Width           =   1000
   End
   Begin VB.ListBox List2 
      Height          =   4935
      Left            =   5280
      TabIndex        =   2
      Top             =   2760
      Width           =   1000
   End
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   840
      TabIndex        =   1
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton Back 
      Caption         =   "Back"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Scorecard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private con As New ADODB.Connection
Public rs As New ADODB.Recordset
Private Sub Back_Click()
Unload Me
Main.Show
End Sub

Private Sub Form_Load()
Dim con As New ADODB.Connection
con.ConnectionString = "Dim con As New ADODB.Connection"
con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Ayush\Account.accdb;Persist Security Info=False"
con.Open
Set rs = New ADODB.Recordset
rs.Open "SELECT * from Player", con, adOpenStatic, adLockOptimistic
Do While Not rs.EOF = True
List1.AddItem rs.Fields("Playername")
List2.AddItem rs.Fields("EScore")
List3.AddItem rs.Fields("MScore")
List4.AddItem rs.Fields("HScore")
rs.MoveNext
Loop
End Sub
