VERSION 5.00
Begin VB.Form EDM 
   Caption         =   "Early Design Models"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   375
      Left            =   1560
      TabIndex        =   22
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scaling factor"
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton exit 
         Caption         =   "Estimated Effort"
         Height          =   375
         Left            =   2880
         TabIndex        =   21
         Top             =   4080
         Width           =   1695
      End
      Begin VB.ComboBox Combo5 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2280
         TabIndex        =   20
         Top             =   3000
         Width           =   1095
      End
      Begin VB.ComboBox Combo4 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2280
         TabIndex        =   19
         Top             =   2520
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2280
         TabIndex        =   18
         Top             =   2040
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2280
         TabIndex        =   17
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   4560
         TabIndex        =   16
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   4560
         TabIndex        =   14
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   4560
         TabIndex        =   13
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4560
         TabIndex        =   12
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4560
         TabIndex        =   11
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text1 
         DataField       =   "value"
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   4560
         TabIndex        =   10
         Top             =   1080
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2280
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Total"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "PMAT"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "TEAM"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "RESL"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "FLEX"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "PREC"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Value"
         Height          =   495
         Left            =   4800
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Rating"
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Description"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "EDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim con2 As New ADODB.Connection
Dim rs2 As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim id As Integer
Public pnom As Integer
Public b As Long
Public prpo As Integer
Private Sub Combo1_Click()
rs.Open "prec", con, adOpenDynamic, adLockOptimistic
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
rs.Open "flex", con, adOpenDynamic, adLockOptimistic
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
rs.Open "resl", con, adOpenDynamic, adLockOptimistic
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
rs.Open "team", con, adOpenDynamic, adLockOptimistic
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
rs.Open "pmat", con, adOpenDynamic, adLockOptimistic
id = (Combo5.ListIndex)
rs.MoveFirst
Do While id > 0
 rs.MoveNext
 id = id - 1
Loop
Text5.Text = rs("value")
rs.Close
End Sub

Private Sub exit_Click()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\data.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open "pam", con, adOpenDynamic, adLockOptimistic
rs.AddNew
rs("codename") = opening.comboval

'rs("codename") = existing.Combo1.List(Combo1.ListIndex)
rs("pnom") = pnom
rs("b") = b
rs.Update
con.Close
EDM.Hide
SF.Show
End Sub

Private Sub Command2_Click()
Text6.Text = Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text)
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\data.mdb;Persist Security Info = false"
'con.Execute ("UPDATE " & opening.comboval & " set EDM = " & Val(Text6.Text))
Set rs = New ADODB.Recordset
rs.Open opening.comboval, con, adOpenDynamic, adLockOptimistic
b = 0.91 + 0.01 * (Val(Text6.Text))
pnom = 2.5 * (rs("SIZ")) ^ b
prpo = Val(Text1.Text) * Val(Text2.Text) * Val(Text3.Text) * Val(Text4.Text) * Val(Text5.Text)
rs.Close
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\cocomo\early.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open "prec", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo1.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "flex", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo2.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "resl", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo3.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "team", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo4.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
rs.Open "pmat", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
    Combo5.AddItem (rs("complexity"))
    rs.MoveNext
Loop
rs.Close
End Sub

