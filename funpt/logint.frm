VERSION 5.00
Begin VB.Form logint 
   Caption         =   "Function Count"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5190
   LinkTopic       =   "Form4"
   ScaleHeight     =   3645
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Function Count"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "logint.frx":0000
         Left            =   240
         List            =   "logint.frx":0016
         TabIndex        =   5
         Text            =   "Option Settings"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   495
         Left            =   3600
         TabIndex        =   4
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   495
         Left            =   2280
         TabIndex        =   3
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "logint.frx":009B
         Left            =   3120
         List            =   "logint.frx":00A8
         TabIndex        =   2
         Text            =   "Simple"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Logical Internal File"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "logint"
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
logint.Hide
logint.Show
End Sub

Private Sub Command2_Click()
    rs.AddNew
    rs("requirement") = Combo2.List(Combo2.ListIndex)
    rs("complexity") = Combo1.List(Combo1.ListIndex)
    rs.Update
    logint.Hide
    extint.Show
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = c:\nivi\softest.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open "logint", con, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Text1_Click()
      Text1.Text = ""
End Sub

