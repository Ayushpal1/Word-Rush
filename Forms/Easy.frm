VERSION 5.00
Begin VB.Form Easy 
   Appearance      =   0  'Flat
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
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   5040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "back"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   2
      Left            =   17040
      Picture         =   "Easy.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   1
      Left            =   17640
      Picture         =   "Easy.frx":0D37
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   750
      Index           =   0
      Left            =   18240
      Picture         =   "Easy.frx":1A6E
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   7920
      TabIndex        =   7
      Top             =   2400
      Width           =   3255
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
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Playername 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9000
      TabIndex        =   5
      Top             =   480
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   18615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      Height          =   375
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   18615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   600
      Index           =   3
      Left            =   10680
      TabIndex        =   4
      Top             =   7140
      Width           =   2655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   600
      Index           =   2
      Left            =   5880
      TabIndex        =   3
      Top             =   7140
      Width           =   2655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   600
      Index           =   1
      Left            =   10560
      TabIndex        =   2
      Top             =   4380
      Width           =   2655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "word"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   600
      Index           =   0
      Left            =   5760
      TabIndex        =   1
      Top             =   4380
      Width           =   2655
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Easy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private con As New ADODB.Connection
Public rs, rsscore As New ADODB.Recordset
Dim rand, cnt, Pscore, Score, heart As Integer
Dim a, b, c, d, word, org, pname, com As String
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Load()
rand = 0
cnt = 0
Score = 0
'com = ""
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
word = rs.Fields("Four") 'word stores the original word from the database
org = word               'org stores original word for comparison and other purpose
Call Divide
End Sub
Private Sub Divide()   'Sub procedure to divide words based upon size
a = Right$(word, 1)
word = Left$(word, Len(word) - 1)
b = Right$(word, 1)
word = Left$(word, Len(word) - 1)
c = Right$(word, 1)
word = Left$(word, Len(word) - 1)
d = Right$(word, 1)
word = Left$(word, Len(word) - 1)
Call Alloc
End Sub
Private Sub Alloc()    'Sub procedure to Allocate words to labels
If rand = 0 Then           'This is not actual random just imitation of random
    Label1(0).Caption = b      'a b c d are used to store the splited words
    Label1(1).Caption = a
    Label1(2).Caption = d
    Label1(3).Caption = c
    rand = rand + 1
ElseIf rand = 1 Then
    Label1(0).Caption = d
    Label1(1).Caption = a
    Label1(2).Caption = b
    Label1(3).Caption = c
    rand = rand + 1
ElseIf rand = 2 Then
    Label1(0).Caption = a
    Label1(1).Caption = d
    Label1(2).Caption = b
    Label1(3).Caption = c
    rand = 0 'resets the rand variable to 0 so it keeps on repeating this sequence until total words are completed
End If
End Sub

Private Sub Label1_Click(Index As Integer)
If Label1(Index).BackColor = &H80000005 Then
    Label1(Index).BackColor = &HC000&
    cnt = cnt + 1
    Label3.Caption = Label3.Caption + Label1(Index).Caption
    com = Left$(org, cnt)
    'Text1.Text = com
    Call Life(Index)
    If cnt = 4 Then
        If Label3.Caption = org Then
            Call savescore
            If Not rs.EOF = True Then
                rs.MoveNext
            End If
            If rs.EOF = False Then
                Call Subreset
            Else
                MsgBox ("You won")
                'con.Close
                Unload Me
            End If
        Else
            cnt = 0
            Label3.Caption = ""
            MsgBox ("Wrong!")
            Call Reset
        End If
    End If
'Else
    'Label1(Index).BackColor = &H80000005
    'Label3.Caption = Replace(Label3.Caption, Label1(Index).Caption, "")
    'cnt = cnt - 1
End If
End Sub

Private Sub Subreset()  'A Sub procedure that cordinates with reset to all the changes made during game
Shape2.Width = 18615
cnt = 0
Label3.Caption = ""
Call Reset
If Not rs.EOF = True Then
    word = rs.Fields("Four")
    org = word
    Call Divide
End If
End Sub

Private Sub savescore()  'saves score of the player if the previous score is less then the score is updated if not then it is kept as it is.
Pscore = rsscore.Fields("EScore")
    Score = Score + 10
If Pscore < Score Then
    rsscore.Fields("EScore") = Score
    rsscore.Update
    Label2.Caption = Score
Else
    Label2.Caption = Score
End If
End Sub

Private Sub Timer1_Timer()   'Timer which acts as the progress bar sets the timer of accurate 15 secs
Shape2.Width = Shape2.Width - 22
    If Shape2.Width <= 15 Then
        Timer1.Enabled = False
        'con.Close
        MsgBox ("Game over")
        Unload Me
    End If
End Sub
Private Sub Reset()  'Sub procedure that resets all the value of the form whenever a word is complete
Dim i As Integer
For i = 0 To Len(org) - 1
    Label1(i).BackColor = &H80000005
Next
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
