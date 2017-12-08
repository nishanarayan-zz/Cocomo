VERSION 5.00
Begin VB.Form PAMplatfrom 
   Caption         =   "Post architectural methods"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5730
   LinkTopic       =   "Form7"
   ScaleHeight     =   4185
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Platform factors and Project factors "
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5175
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3480
         TabIndex        =   15
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3480
         TabIndex        =   14
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3480
         TabIndex        =   13
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3480
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3480
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   2760
         Width           =   1095
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "SITE"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "TOOL"
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "PVOL"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "STOR"
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "TIME"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "PAMplatfrom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim id As Integer
Private Sub Command1_Click()
    PAMplatfrom.Hide
    PAMpersonnel.Show
End Sub
Private Sub Combo1_Click()
rs.Open "tim", con, adOpenDynamic, adLockOptimistic
id = (Combo1.ListIndex)
rs.MoveFirst
Do While id > 0
 rs.MoveNext
 id = id - 1
Loop
Text1.Text = rs("value")
rs.Close
End Sub

Private Sub Combo2_Click()
rs.Open "stor", con, adOpenDynamic, adLockOptimistic
id = (Combo2.ListIndex)
rs.MoveFirst
Do While id > 0
 rs.MoveNext
 id = id - 1
Loop
Text2.Text = rs("value")
rs.Close
End Sub

Private Sub Combo3_Click()
rs.Open "pvol", con, adOpenDynamic, adLockOptimistic
id = (Combo3.ListIndex)
rs.MoveFirst
Do While id > 0
 rs.MoveNext
 id = id - 1
Loop
Text3.Text = rs("value")
rs.Close
End Sub

Private Sub Combo4_Click()
rs.Open "tool", con, adOpenDynamic, adLockOptimistic
id = (Combo4.ListIndex)
rs.MoveFirst
Do While id > 0
 rs.MoveNext
 id = id - 1
Loop
Text4.Text = rs("value")
rs.Close
End Sub

Private Sub Combo5_Click()
rs.Open "site", con, adOpenDynamic, adLockOptimistic
id = (Combo5.ListIndex)
rs.MoveFirst
Do While id > 0
 rs.MoveNext
 id = id - 1
Loop
Text5.Text = rs("value")
rs.Close
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\cocomo\post.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open "stor", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo2.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "tim", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo1.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "pvol", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo3.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "tool", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo4.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "site", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo5.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
End Sub


