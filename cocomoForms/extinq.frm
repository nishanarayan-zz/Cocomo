VERSION 5.00
Begin VB.Form extinq 
   Caption         =   "Function Count"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5190
   LinkTopic       =   "Form6"
   ScaleHeight     =   3315
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Function Count"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "extinq.frx":0000
         Left            =   240
         List            =   "extinq.frx":000D
         TabIndex        =   5
         Text            =   "Inquiry to run plugins"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   495
         Left            =   2160
         TabIndex        =   4
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   495
         Left            =   3600
         TabIndex        =   3
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "extinq.frx":0064
         Left            =   3360
         List            =   "extinq.frx":0071
         TabIndex        =   2
         Text            =   "Simple"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "External Inquiry"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "extinq"
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
extinq.Hide
extinq.Show
End Sub

Private Sub Command2_Click()
    rs.AddNew
    rs("requirement") = Combo2.List(Combo2.ListIndex)
    rs("complexity") = Combo1.List(Combo1.ListIndex)
    rs.Update
    extinq.Hide
    total.Show
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\funpt\softest.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open "extinq", con, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Text1_Click()
      Text1.Text = ""
End Sub

