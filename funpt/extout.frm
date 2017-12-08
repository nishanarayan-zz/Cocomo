VERSION 5.00
Begin VB.Form extout 
   Caption         =   "Function Count"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5250
   LinkTopic       =   "Form3"
   ScaleHeight     =   3900
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Function Count"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "extout.frx":0000
         Left            =   240
         List            =   "extout.frx":0013
         TabIndex        =   5
         Text            =   "Plug ins"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   495
         Left            =   2160
         TabIndex        =   4
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   495
         Left            =   3600
         TabIndex        =   3
         Top             =   2400
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "extout.frx":00AE
         Left            =   3120
         List            =   "extout.frx":00BB
         TabIndex        =   2
         Text            =   "Simple"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "External Output"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "extout"
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
extout.Hide
extout.Show
End Sub

Private Sub Command2_Click()
    rs.AddNew
    rs("requirement") = Combo2.List(Combo2.ListIndex)
    rs("complexity") = Combo1.List(Combo1.ListIndex)
    rs.Update
    extout.Hide
    logint.Show
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = c:\nivi\softest.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open "extout", con, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Text1_Click()
      Text1.Text = ""
End Sub

