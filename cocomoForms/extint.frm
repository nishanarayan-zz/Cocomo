VERSION 5.00
Begin VB.Form extint 
   Caption         =   "Function Count"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   LinkTopic       =   "Form5"
   ScaleHeight     =   3270
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Function Count"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "extint.frx":0000
         Left            =   240
         List            =   "extint.frx":000A
         TabIndex        =   5
         Text            =   "Plugin XML file"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   495
         Left            =   3240
         TabIndex        =   4
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   495
         Left            =   1920
         TabIndex        =   3
         Top             =   1920
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "extint.frx":0046
         Left            =   3120
         List            =   "extint.frx":0053
         TabIndex        =   2
         Text            =   "Simple"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "External Interface File"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "extint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()

rs.AddNew
rs("requirement") = Combo2.List(Combo2.ListIndex)
rs("complexity") = Combo1.List(Combo1.ListIndex)
rs.Update
extint.Hide
extint.Show
End Sub

Private Sub Command2_Click()
    rs.AddNew
    rs("requirement") = Combo2.List(Combo2.ListIndex)
    rs("complexity") = Combo1.List(Combo1.ListIndex)
    rs.Update
    extint.Hide
    extinq.Show
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\funpt\softest.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open "extint", con, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Text1_Click()
      Text1.Text = ""
End Sub

