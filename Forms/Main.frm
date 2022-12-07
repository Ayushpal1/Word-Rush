VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Main 
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   2880
      Top             =   5280
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Ayush\Account.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Ayush\Account.accdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Player"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   13800
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton Scorecardbtn 
      Appearance      =   0  'Flat
      Caption         =   "Scorecard"
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   9960
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "Playername"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7800
      TabIndex        =   0
      Top             =   100
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   18360
      Picture         =   "Main.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "How To Play"
      Top             =   9600
      Width           =   795
   End
   Begin VB.Image About 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   120
      Picture         =   "Main.frx":027D
      Stretch         =   -1  'True
      ToolTipText     =   "About Us"
      Top             =   9600
      Width           =   795
   End
   Begin VB.Image Exit 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   18360
      Picture         =   "Main.frx":06BD
      Stretch         =   -1  'True
      ToolTipText     =   "Exit"
      Top             =   100
      Width           =   795
   End
   Begin VB.Image Accounticon 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   120
      Picture         =   "Main.frx":0A8F
      Stretch         =   -1  'True
      ToolTipText     =   "Account"
      Top             =   100
      Width           =   795
   End
   Begin VB.Image Play 
      Appearance      =   0  'Flat
      Height          =   1755
      Left            =   9120
      Picture         =   "Main.frx":0E12
      Stretch         =   -1  'True
      ToolTipText     =   "Play"
      Top             =   4440
      Width           =   1755
   End
   Begin VB.Image Background 
      Height          =   10680
      Left            =   0
      Picture         =   "Main.frx":16BB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19950
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private con As New ADODB.Connection
Public rs As New ADODB.Recordset
Private Sub About_Click()
AboutUs.Show
End Sub
Private Sub Accounticon_Click()
Account.Show
Unload Me
End Sub

Private Sub Exit_Click()
Dim t As String
t = MsgBox("Are you sure you want to Quit?", vbYesNo, "Question")
If t = vbYes Then
Unload Me
End If
End Sub
Private Sub Form_Load()
Dim con As New ADODB.Connection
'con.ConnectionString = "Dim con As New ADODB.Connection"
con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Ayush\Account.accdb;Persist Security Info=False"
con.Open
Set rs = New ADODB.Recordset
rs.Open "SELECT * from Player", con, adOpenStatic, adLockOptimistic
Do While Not rs.EOF = True
Combo1.AddItem rs.Fields("Playername")
rs.MoveNext
Loop
End Sub

Private Sub Image1_Click()
Howtoplay.Show
Unload Me
End Sub

Private Sub Play_Click()
Text1.Text = Combo1.List(Combo1.ListIndex)
If Text1.Text = "" Then
 MsgBox ("Please Select a account to play")
Else
Level.Show
End If
End Sub
Private Sub Scorecardbtn_Click()
DataReport1.Show

End Sub
