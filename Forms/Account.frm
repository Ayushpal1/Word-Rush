VERSION 5.00
Begin VB.Form Account 
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
   Begin VB.CommandButton Command5 
      Caption         =   "Update"
      Height          =   975
      Left            =   14160
      TabIndex        =   15
      Top             =   6840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton updatebtn 
      Caption         =   "update"
      Height          =   975
      Left            =   6840
      TabIndex        =   14
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete Account"
      Height          =   975
      Left            =   6840
      TabIndex        =   10
      Top             =   6600
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   14160
      TabIndex        =   9
      Top             =   4320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   14160
      TabIndex        =   8
      Top             =   5640
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   14160
      TabIndex        =   7
      Top             =   4320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create new account"
      Height          =   855
      Left            =   6840
      TabIndex        =   6
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   14160
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete"
      Height          =   1095
      Left            =   14280
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add"
      Height          =   855
      Left            =   14160
      TabIndex        =   2
      Top             =   6960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7230
      Left            =   600
      TabIndex        =   1
      ToolTipText     =   "List of Players"
      Top             =   2400
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "New password :"
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   11520
      TabIndex        =   16
      Top             =   4440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm password :"
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   10800
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   12000
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Microsoft YaHei"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   12600
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Players"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   25.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   600
      TabIndex        =   5
      Top             =   1680
      Width           =   4935
   End
End
Attribute VB_Name = "Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private con As New ADODB.Connection
Public rs, rspass, rsup As New ADODB.Recordset

Private Sub Add_Click()
Dim pname, pass As String
pname = Text1.Text
pass = Text2.Text
If Len(pname) < 1 Then
    MsgBox ("Playername can't be empty")
Else
    If Len(pname) > 2 Then
        If Len(Text2.Text) < 8 Then
            MsgBox ("Password should be atleast 8 Characters")
        Else
            If StrComp(Text2.Text, Text3.Text) = 0 Then
                rs.AddNew
                rs.Fields("Password") = pass
                rs.Fields("Score") = 0
                rs.Fields("EScore") = 0
                rs.Fields("MScore") = 0
                rs.Fields("HScore") = 0
                rs.Fields("Playername") = pname
                rs.Update
                Unload Me
                Account.Show
                MsgBox "Account created successfully"
            Else
                MsgBox ("Passwords does not match")
            End If
        End If
    Else
        MsgBox ("Playername too short Enter again")
    End If
End If
End Sub

Private Sub Command1_Click()
Main.Show
Unload Me
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text1.Enabled = True
Delete.Visible = False
Text4.Visible = False
Label5.Visible = False
Command5.Visible = False
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Add.Visible = True
End Sub

Private Sub Command3_Click()
Dim name As String
name = List1.List(List1.ListIndex)
If Len(name) < 1 Then
    MsgBox ("Please select a Player from the list to update and then again click on Update")
Else
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Text1.Visible = False
    Label5.Visible = False
    Text2.Visible = False
    Text3.Visible = False
    Command5.Visible = False
    Add.Visible = False
    'MsgBox ("Please select a Account from the List And enter the password")
    Delete.Visible = True
    Text4.Visible = True
    Label3.Visible = True
    Label2.Visible = True
    Text1.Visible = True
    Text1.Enabled = False
    Text1.Text = List1.List(List1.ListIndex)
End If
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
Dim pname, pass As String
Dim name As String
Dim con As New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Ayush\Account.accdb;Persist Security Info=False"
con.Open
name = List1.List(List1.ListIndex)
Set rsup = New ADODB.Recordset
rsup.Open "SELECT * from Player WHERE Playername = '" & name & "'", con, adOpenStatic, adLockOptimistic
pname = Text1.Text
pass = Text2.Text
If Len(pname) < 1 Then
    MsgBox ("Playername can't be empty")
Else
    If Len(pname) > 2 Then
        If Len(Text2.Text) < 8 Then
            MsgBox ("Password should be atleast 8 Characters")
        Else
            If StrComp(Text2.Text, Text3.Text) = 0 Then
                rsup.Fields("Playername") = pname
                rsup.Fields("Password") = pass
                rsup.Fields("Score") = 0
                rsup.Fields("EScore") = 0
                rsup.Fields("MScore") = 0
                rsup.Fields("HScore") = 0
                rsup.Fields("Playername") = pname
                rsup.Update
                Unload Me
                Account.Show
                MsgBox "Account updated successfully"
            Else
                MsgBox ("Passwords does not match")
            End If
        End If
    Else
        MsgBox ("Playername too short Enter again")
    End If
End If
End Sub

Private Sub Delete_Click()
Dim name, pass, dpass As String
name = List1.List(List1.ListIndex)
'Text1.Text = name
If Len(name) < 1 Then
    MsgBox ("Please select a Account")
Else
    'Text1.Enabled = True
    'Text1.Text = name
    'Text1.Enabled = False
    pass = Text4.Text
    con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Ayush\Account.accdb;Persist Security Info=False"
    con.Open
    Set rspass = New ADODB.Recordset
    rspass.Open "SELECT * from Player WHERE Playername = '" & name & "'", con, adOpenStatic, adLockOptimistic
    dpass = rspass.Fields("Password")
    If pass = dpass Then
        con.BeginTrans
        con.Execute "DELETE FROM Player WHERE Playername = '" & name & "' AND Password = '" & pass & "'"
        con.CommitTrans
        con.Close
        MsgBox ("Record Deleted Succesfully!")
        Unload Me
        Account.Show
    Else
        MsgBox ("Wrong password!")
        con.Close
    End If
End If
End Sub

Private Sub Form_Load()
Dim con As New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Ayush\Account.accdb;Persist Security Info=False"
con.Open
Set rs = New ADODB.Recordset
rs.Open "SELECT * from Player", con, adOpenStatic, adLockOptimistic
Do While Not rs.EOF = True
    List1.AddItem rs.Fields("Playername")
    rs.MoveNext
Loop
End Sub

Private Sub updatebtn_Click()
Dim name As String
name = List1.List(List1.ListIndex)
If Len(name) < 1 Then
    MsgBox ("Please select a Player from the list to update and then again click on Update")
Else
    Add.Visible = False
    Delete.Visible = False
    Text1.Text = name
    Text1.Enabled = True
    Delete.Visible = False
    Text4.Visible = False
    Label2.Visible = True
    Label3.Visible = False
    Label5.Visible = True
    Label4.Visible = True
    Text1.Visible = True
    Text2.Visible = True
    Text3.Visible = True
    Command5.Visible = True
End If
End Sub
