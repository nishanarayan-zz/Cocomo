VERSION 5.00
Begin VB.Form extinp 
   Caption         =   "Function Count"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5580
   LinkTopic       =   "Form2"
   ScaleHeight     =   3930
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Function Count"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "extinp.frx":0000
         Left            =   480
         List            =   "extinp.frx":0016
         TabIndex        =   5
         Text            =   "Requirements Input"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   495
         Left            =   3720
         TabIndex        =   4
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   495
         Left            =   2280
         TabIndex        =   3
         Top             =   2520
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "extinp.frx":009F
         Left            =   3360
         List            =   "extinp.frx":00AC
         TabIndex        =   2
         Text            =   "Simple"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "External Input"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
   End
End
Attribute VB_Name = "extinp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset



Private Sub Command1_Click()
Text1.Text = ""
rs.AddNew
rs("requirement") = Combo2.List(Combo2.ListIndex)
rs("complexity") = Combo1.List(Combo1.ListIndex)
rs.Update
extinp.Hide
extinp.Show
End Sub

Private Sub Command2_Click()
    rs.AddNew
    rs("requirement") = Combo2.List(Combo2.ListIndex)
    rs("complexity") = Combo1.List(Combo1.ListIndex)
    rs.Update
    extinp.Hide
    extout.Show
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = c:\nivi\softest.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open "extinp", con, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Text1_Click()
      Text1.Text = ""
End Sub

