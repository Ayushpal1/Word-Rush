VERSION 5.00
Begin VB.Form Medium 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10650
   ClientLeft      =   0
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
      Caption         =   "back"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   8320
      TabIndex        =   0
      Top             =   2900
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   4920
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   0
      Left            =   18240
      Picture         =   "Medium.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   1
      Left            =   17640
      Picture         =   "Medium.frx":0D37
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   2
      Left            =   17040
      Picture         =   "Medium.frx":1A6E
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei UI"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   17640
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Playername 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   9000
      TabIndex        =   8
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   5
      Left            =   12480
      TabIndex        =   7
      Top             =   6960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   4
      Left            =   12480
      TabIndex        =   6
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   3
      Left            =   8640
      TabIndex        =   5
      Top             =   6960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   2
      Left            =   4800
      TabIndex        =   4
      Top             =   6960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   1
      Left            =   8640
      TabIndex        =   3
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   600
      Index           =   0
      Left            =   4800
      TabIndex        =   2
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   18615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      Height          =   375
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   18615
   End
End
Attribute VB_Name = "Medium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private con As New ADODB.Connection
Public rs, rsscore As New ADODB.Recordset
Dim random, wcount, Score, heart As Integer 'random is the id of randomness, wcount stores the count of the word, score is current score which will be compared.
Dim Pscore As Integer 'Stores previous score
Dim a, b, c, d, e, f, word, org, pname, com As String 'letters are used to store letters, word is used to store the word which arrived from the database, org is used to store the original word, pname stores the name of player which is used to display and update the score.
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
random = 0
wcount = 0
Score = 0
heart = 2
Playername.Caption = Main.Text1.Text
pname = Playername.Caption
Dim con As New ADODB.Connection
'con.ConnectionString = "Dim con As New ADODB.Connection"
con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Ayush\Account.accdb;Persist Security Info=False"
con.Open
Set rs = New ADODB.Recordset
Set rsscore = New ADODB.Recordset
rs.Open "SELECT top 10 * FROM Dictionary ORDER BY Rnd(INT(NOW*id)-NOW*id)", con, adOpenStatic, adLockOptimistic
rsscore.Open "SELECT * from Player WHERE Playername = '" & pname & "'", con, adOpenStatic, adLockOptimistic
word = rs.Fields("Six")
Pscore = rsscore.Fields("MScore")
org = word
'Text3.Text = org
Text1.Text = ""
Call Divide
End Sub

Private Sub Label1_Click(Index As Integer)
If Label1(Index).BackColor = &H80000005 Then
    Label1(Index).BackColor = &HC000&
    Text1.Text = Text1.Text + Label1(Index).Caption
    wcount = wcount + 1
    com = Left$(org, wcount)
    'Text1.Text = com
    Call Life(Index)
    If wcount = 6 Then
        'Text3.Text = Text1.Text
        If Text1.Text = org Then
            Call savescore
            If Not rs.EOF = True Then
                rs.MoveNext
            End If
            If rs.EOF = False Then
                Call karna
            Else
                MsgBox ("You won")
                'con.Close
                Unload Me
            End If
        Else
            wcount = 0
            Text1.Text = ""
            MsgBox ("Wrong!")
            Call Reset
        End If
    End If
'Else
    'Label1(Index).BackColor = &H80000005
    'Text1.Text = Replace(Text1.Text, Label1(Index).Caption, "", 1, 1)
    'wcount = wcount - 1
End If
'Label1(Index).BackColor = &HC000&
'Text1.Text = Text1.Text + Label1(Index).Caption
'cnt = cnt + 1
End Sub

Private Sub Timer1_Timer()
             Shape2.Width = Shape2.Width - 18
             If Shape2.Width <= 15 Then
             Timer1.Enabled = False
             'con.Close
             MsgBox ("Game over")
             Unload Me
             End If
End Sub
Private Sub karna()
Shape2.Width = 18615
wcount = 0
Text1.Text = ""
'Call savescore
Call Reset
If Not rs.EOF = True Then
word = rs.Fields("Six")
org = word
'Text3.Text = org
Call Divide
End If
End Sub
Private Sub Divide()
a = Right$(word, 1)
word = Left$(word, Len(word) - 1)
b = Right$(word, 1)
word = Left$(word, Len(word) - 1)
c = Right$(word, 1)
word = Left$(word, Len(word) - 1)
d = Right$(word, 1)
word = Left$(word, Len(word) - 1)
e = Right$(word, 1)
word = Left$(word, Len(word) - 1)
f = Right$(word, 1)
word = Left$(word, Len(word) - 1)
Call allocsix
End Sub
Private Sub allocsix()
If random = 0 Then
Label1(0).Caption = b
Label1(1).Caption = a
Label1(2).Caption = d
Label1(3).Caption = c
Label1(4).Caption = f
Label1(5).Caption = e
random = random + 1
ElseIf random = 1 Then
Label1(0).Caption = e
Label1(1).Caption = a
Label1(2).Caption = b
Label1(3).Caption = c
Label1(4).Caption = d
Label1(5).Caption = f
random = random + 1
ElseIf random = 2 Then
Label1(0).Caption = a
Label1(1).Caption = c
Label1(2).Caption = e
Label1(3).Caption = f
Label1(4).Caption = d
Label1(5).Caption = b
random = 0
End If
End Sub
Private Sub savescore()
Pscore = rsscore.Fields("MScore")
    Score = Score + 10
If Pscore < Score Then
    rsscore.Fields("MScore") = Score
    rsscore.Update
    Label2.Caption = Score
Else
    Label2.Caption = Score
End If
End Sub
Private Sub Reset()
Dim i As Integer
For i = 0 To Len(org) - 1
Label1(i).BackColor = &H80000005
Next
End Sub
Private Sub Life(n)
If StrComp(Text1.Text, com) = -1 Then
    Image1(heart).Visible = False
    heart = heart - 1
    Label1(n).BackColor = &H80000005
    wcount = wcount - 1
    'If cnt > -1 Then
        Text1.Text = Left$(Text1.Text, wcount)
    'End If
End If
If StrComp(Text1.Text, com) = 1 Then
    Image1(heart).Visible = False
    heart = heart - 1
    Label1(n).BackColor = &H80000005
    wcount = wcount - 1
    'If cnt > -1 Then
        Text1.Text = Left$(Text1.Text, wcount)
    'End If
End If
If heart = -1 Then
    Timer1.Enabled = False
    MsgBox ("Game over")
    Unload Me
End If
End Sub
