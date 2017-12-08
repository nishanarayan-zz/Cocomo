VERSION 5.00
Begin VB.Form PAMpersonnel 
   Caption         =   "Post Architectural Methods"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6120
   LinkTopic       =   "Form8"
   ScaleHeight     =   4680
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Personnel factors"
      Height          =   4095
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      Begin VB.CommandButton Command1 
         Caption         =   "Estimated Effort"
         Height          =   375
         Left            =   1920
         TabIndex        =   19
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3600
         TabIndex        =   18
         Top             =   3120
         Width           =   975
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   1800
         TabIndex        =   17
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3600
         TabIndex        =   16
         Top             =   2640
         Width           =   975
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3600
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1800
         TabIndex        =   13
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3600
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3600
         TabIndex        =   10
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3600
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "PCON"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "LTEX"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "PEXP"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "AEXP"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "PCAP"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "ACAP"
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "PAMpersonnel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim id As Integer
'Public total As Integer
Public prod As Long
Private Sub Command1_Click()
    'total = Val(PAMproduct.Text1.Text) + Val(PAMproduct.Text2.Text) + Val(PAMproduct.Text3.Text) + Val(PAMproduct.Text4.Text) + Val(PAMproduct.Text5.Text) + Val(PAMplatfrom.Text1.Text) + Val(PAMplatfrom.Text2.Text) + Val(PAMplatfrom.Text3.Text) + Val(PAMplatfrom.Text4.Text) + Val(PAMplatfrom.Text5.Text) + Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text) + Val(Text6.Text)
    'Set con = New ADODB.Connection
    'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\data.mdb;Persist Security Info = false"
    'con.Execute ("UPDATE " & opening.comboval & " set PAM = " & total)
    PAMpersonnel.Hide
    prod = Val(PAMproduct.Text1.Text) * Val(PAMproduct.Text2.Text) * Val(PAMproduct.Text3.Text) * Val(PAMproduct.Text4.Text) * Val(PAMproduct.Text5.Text) * Val(PAMplatfrom.Text1.Text) * Val(PAMplatfrom.Text2.Text) * Val(PAMplatfrom.Text3.Text) * Val(PAMplatfrom.Text4.Text) * Val(PAMplatfrom.Text5.Text) * Val(Text1.Text) * Val(Text2.Text) * Val(Text3.Text) * Val(Text4.Text) * Val(Text5.Text) * Val(Text6.Text)
    'SF.Form_Load
    SF.Show
End Sub
Private Sub Combo1_Click()
rs.Open "acap", con, adOpenDynamic, adLockOptimistic
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
rs.Open "pcap", con, adOpenDynamic, adLockOptimistic
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
rs.Open "aexp", con, adOpenDynamic, adLockOptimistic
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
rs.Open "pexp", con, adOpenDynamic, adLockOptimistic
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
rs.Open "ltex", con, adOpenDynamic, adLockOptimistic
id = (Combo5.ListIndex)
rs.MoveFirst
Do While id > 0
 rs.MoveNext
 id = id - 1
Loop
Text5.Text = rs("value")
rs.Close
End Sub
Private Sub Combo6_Click()
rs.Open "pcon", con, adOpenDynamic, adLockOptimistic
id = (Combo6.ListIndex)
rs.MoveFirst
Do While id > 0
 rs.MoveNext
 id = id - 1
Loop
Text6.Text = rs("value")
rs.Close
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\cocomo\post.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open "acap", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo1.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "pcap", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo2.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "aexp", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo3.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "pexp", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo4.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "ltex", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo5.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "pcon", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo6.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
End Sub



