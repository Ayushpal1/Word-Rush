VERSION 5.00
Begin VB.Form Hard 
   Appearance      =   0  'Flat
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   ClientHeight    =   10650
   ClientLeft      =   -15
   ClientTop       =   -105
   ClientWidth     =   19320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10680
   ScaleMode       =   0  'User
   ScaleWidth      =   19320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "back"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   480
      Top             =   4920
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   1095
      Left            =   7080
      TabIndex        =   11
      Top             =   3360
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   0
      Left            =   18000
      Picture         =   "Hard.frx":0000
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   1
      Left            =   17400
      Picture         =   "Hard.frx":0D37
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   2
      Left            =   16800
      Picture         =   "Hard.frx":1A6E
      Stretch         =   -1  'True
      Top             =   2040
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
      TabIndex        =   10
      Top             =   480
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
      TabIndex        =   9
      Top             =   480
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   375
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   18615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   18615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   7
      Left            =   12960
      TabIndex        =   8
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   6
      Left            =   9840
      TabIndex        =   7
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   5
      Left            =   6600
      TabIndex        =   6
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   4
      Left            =   3360
      TabIndex        =   5
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   3
      Left            =   12960
      TabIndex        =   4
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   2
      Left            =   9840
      TabIndex        =   3
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   1
      Left            =   6600
      TabIndex        =   2
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   5160
      Width           =   2535
   End
End
Attribute VB_Name = "Hard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private con As New ADODB.Connection
Public rs, rsscore As New ADODB.Recordset
Dim rand, cnt, Score, heart As Integer
Dim Pscore As Integer
Dim a, b, c, d, e, f, g, h, shabad, org, pname, com As String
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
rand = 0
cnt = 0
Score = 0
com = ""
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
shabad = rs.Fields("Eight")
org = shabad
'Text3.Text = org
Label3.Caption = ""
Call shuffle
End Sub
Private Sub karna()
Shape2.Width = 18615
cnt = 0
Label3.Caption = ""
'Call savescore
Call thikkar
If Not rs.EOF = True Then
shabad = rs.Fields("Eight")
org = shabad
'Text3.Text = org
Call shuffle
End If
End Sub
Private Sub shuffle()
a = Right$(shabad, 1)
shabad = Left$(shabad, Len(shabad) - 1)
b = Right$(shabad, 1)
shabad = Left$(shabad, Len(shabad) - 1)
c = Right$(shabad, 1)
shabad = Left$(shabad, Len(shabad) - 1)
d = Right$(shabad, 1)
shabad = Left$(shabad, Len(shabad) - 1)
e = Right$(shabad, 1)
shabad = Left$(shabad, Len(shabad) - 1)
f = Right$(shabad, 1)
shabad = Left$(shabad, Len(shabad) - 1)
g = Right$(shabad, 1)
shabad = Left$(shabad, Len(shabad) - 1)
h = Right$(shabad, 1)
shabad = Left$(shabad, Len(shabad) - 1)
Call alloceight
End Sub
Private Sub alloceight()
If rand = 0 Then
Label1(0).Caption = d
Label1(1).Caption = b
Label1(2).Caption = a
Label1(3).Caption = e
Label1(4).Caption = c
Label1(5).Caption = h
Label1(6).Caption = f
Label1(7).Caption = g
rand = rand + 1
ElseIf rand = 1 Then
Label1(0).Caption = c
Label1(1).Caption = a
Label1(2).Caption = e
Label1(3).Caption = d
Label1(4).Caption = g
Label1(5).Caption = b
Label1(6).Caption = f
Label1(7).Caption = h
rand = rand + 1
ElseIf rand = 2 Then
Label1(0).Caption = b
Label1(1).Caption = e
Label1(2).Caption = d
Label1(3).Caption = g
Label1(4).Caption = c
Label1(5).Caption = f
Label1(6).Caption = a
Label1(7).Caption = h
rand = rand + 1
ElseIf rand = 3 Then
Label1(0).Caption = a
Label1(1).Caption = d
Label1(2).Caption = f
Label1(3).Caption = c
Label1(4).Caption = h
Label1(5).Caption = e
Label1(6).Caption = b
Label1(7).Caption = g
rand = rand + 1
ElseIf rand = 4 Then
Label1(0).Caption = c
Label1(1).Caption = b
Label1(2).Caption = g
Label1(3).Caption = a
Label1(4).Caption = e
Label1(5).Caption = f
Label1(6).Caption = d
Label1(7).Caption = h
rand = rand + 1
ElseIf rand = 5 Then
Label1(0).Caption = e
Label1(1).Caption = c
Label1(2).Caption = b
Label1(3).Caption = f
Label1(4).Caption = d
Label1(5).Caption = g
Label1(6).Caption = h
Label1(7).Caption = a
rand = 0
End If
End Sub
Private Sub savescore()
Pscore = rsscore.Fields("HScore")
    Score = Score + 15
If Pscore < Score Then
    rsscore.Fields("HScore") = Score
    rsscore.Update
    Label2.Caption = Score
Else
    Label2.Caption = Score
End If
End Sub
Private Sub thikkar()
Dim i As Integer
For i = 0 To Len(org) - 1
Label1(i).BackColor = &H80000005
Next
End Sub

Private Sub Label1_Click(Index As Integer)
If Label1(Index).BackColor = &H80000005 Then
 Label1(Index).BackColor = &HC000&
     Label3.Caption = Label3.Caption + Label1(Index).Caption
 cnt = cnt + 1
    com = Left$(org, cnt)
    'Text1.Text = com
    Call Life(Index)
    If cnt = 8 Then
    'Text3.Text = Text1.Text
        If Label3.Caption = org Then
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
            cnt = 0
            Label3.Caption = ""
            MsgBox ("Wrong!")
            Call thikkar
        End If
    End If
Else
    Label1(Index).BackColor = &H80000005
    Label3.Caption = Replace(Label3.Caption, Label1(Index).Caption, "")
    cnt = cnt - 1
End If
End Sub

Private Sub Timer1_Timer()
             Shape2.Width = Shape2.Width - 22
             If Shape2.Width <= 15 Then
             Timer1.Enabled = False
             'con.Close
             MsgBox ("Game over")
             Unload Me
             End If
End Sub
Private Sub Life(n)
If StrComp(Label3.Caption, com) = -1 Then
    Image1(heart).Visible = False
    heart = heart - 1
    Label1(n).BackColor = &H80000005
    cnt = cnt - 1
    'If cnt > -1 Then
        Label3.Caption = Left$(Label3.Caption, cnt)
    'End If
End If
If StrComp(Label3.Caption, com) = 1 Then
    Image1(heart).Visible = False
    heart = heart - 1
    Label1(n).BackColor = &H80000005
    cnt = cnt - 1
    'If cnt > -1 Then
        Label3.Caption = Left$(Label3.Caption, cnt)
    'End If
End If
If heart = -1 Then
    Timer1.Enabled = False
    MsgBox ("Game over")
    Unload Me
End If
End Sub
