VERSION 5.00
Begin VB.Form PAMproduct 
   Caption         =   " Post Architectural methods - Cost Drivers"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4395
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Product Factor"
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin VB.Frame Frame2 
         Caption         =   "Product Factor"
         Height          =   3735
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   5415
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   3840
            TabIndex        =   22
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   3840
            TabIndex        =   21
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   3840
            TabIndex        =   20
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   3840
            TabIndex        =   19
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   3840
            TabIndex        =   18
            Top             =   2880
            Width           =   735
         End
         Begin VB.ComboBox Combo3 
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   1920
            TabIndex        =   17
            Top             =   960
            Width           =   1095
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   1920
            TabIndex        =   16
            Top             =   1440
            Width           =   1095
         End
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   1920
            TabIndex        =   15
            Top             =   1920
            Width           =   1095
         End
         Begin VB.ComboBox Combo6 
            Height          =   315
            Left            =   1920
            TabIndex        =   14
            Top             =   2400
            Width           =   1095
         End
         Begin VB.ComboBox Combo7 
            Height          =   315
            Left            =   1920
            TabIndex        =   13
            Top             =   2880
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "OK"
            Height          =   375
            Left            =   3000
            TabIndex        =   12
            Top             =   3240
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "DOCU"
            Height          =   375
            Left            =   360
            TabIndex        =   30
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "RUSE"
            Height          =   495
            Left            =   360
            TabIndex        =   29
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "CPLX"
            Height          =   375
            Left            =   360
            TabIndex        =   28
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "DATA"
            Height          =   255
            Left            =   360
            TabIndex        =   27
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "RELY"
            Height          =   375
            Left            =   360
            TabIndex        =   26
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Value"
            Height          =   375
            Left            =   3960
            TabIndex        =   25
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Ranking"
            Height          =   375
            Left            =   2160
            TabIndex        =   24
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Description"
            Height          =   375
            Left            =   360
            TabIndex        =   23
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "DOCU"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "RUSE"
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "CPLX"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "DATA"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "RELY"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Value"
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Ranking"
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Description"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "PAMproduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim id As Integer
Private Sub Combo3_Click()
rs.Open "rely", con, adOpenDynamic, adLockOptimistic
id = (Combo3.ListIndex)
rs.MoveFirst
Do While id > 0
 rs.MoveNext
 id = id - 1
Loop
Text1.Text = rs("value")
rs.Close
End Sub
Private Sub Command1_Click()
    PAMproduct.Hide
    PAMplatfrom.Show
End Sub
Private Sub Combo4_Click()
rs.Open "data", con, adOpenDynamic, adLockOptimistic
id = (Combo4.ListIndex)
rs.MoveFirst
Do While id > 0
 rs.MoveNext
 id = id - 1
Loop
Text2.Text = rs("value")
rs.Close
End Sub

Private Sub Combo5_Click()
rs.Open "cplx", con, adOpenDynamic, adLockOptimistic
id = (Combo5.ListIndex)
rs.MoveFirst
Do While id > 0
 rs.MoveNext
 id = id - 1
Loop
Text3.Text = rs("value")
rs.Close
End Sub

Private Sub Combo6_Click()
rs.Open "ruse", con, adOpenDynamic, adLockOptimistic
id = (Combo6.ListIndex)
rs.MoveFirst
Do While id > 0
 rs.MoveNext
 id = id - 1
Loop
Text4.Text = rs("value")
rs.Close
End Sub

Private Sub Combo7_Click()
rs.Open "docu", con, adOpenDynamic, adLockOptimistic
id = (Combo7.ListIndex)
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
rs.Open "rely", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo3.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "data", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo4.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "cplx", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo5.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "ruse", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo6.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "docu", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo7.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
End Sub

