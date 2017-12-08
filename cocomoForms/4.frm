VERSION 5.00
Begin VB.Form Main 
   Caption         =   "COCOMO II"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6315
   LinkTopic       =   "Form4"
   ScaleHeight     =   4560
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "COCOMO II"
      Height          =   3975
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5655
      Begin VB.CommandButton Command3 
         Caption         =   "Post Architecture Model"
         Height          =   615
         Left            =   1200
         TabIndex        =   3
         Top             =   3120
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Early Design Model"
         Height          =   615
         Left            =   1200
         TabIndex        =   2
         Top             =   1860
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Application Composition Model"
         Height          =   615
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Private Sub Command1_Click()
   Main.Hide
   ACM.Show
End Sub

Private Sub Command2_Click()
   Main.Hide
   EDM.Show
End Sub

Private Sub Command3_Click()
    Main.Hide
    PAMproduct.Show
End Sub

Public Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\data.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open opening.comboval, con, adOpenDynamic, adLockOptimistic
If (rs("ACM") = -1) Then
    Command2.Enabled = False
Else
    Command2.Enabled = True
End If
If (rs("EDM") = -1) Then
    Command3.Enabled = False
Else
    Command3.Enabled = True
End If
End Sub
