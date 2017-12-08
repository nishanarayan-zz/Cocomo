VERSION 5.00
Begin VB.Form final 
   Caption         =   "Functional Points"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5955
   LinkTopic       =   "Form12"
   ScaleHeight     =   4515
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit and Generate report for Size Estimation"
      Height          =   735
      Left            =   4080
      TabIndex        =   6
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Moving on to Effort Estimation using COCOMOII Model"
      Height          =   735
      Left            =   1560
      TabIndex        =   5
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Total number of SLOC's"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Enter SLOC's per Function Point"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Function Points Measure      FC * PCA "
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Processing Complexity Adjustment      0.65 + (0.01 * PC)"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "final"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim con As ADODB.Connection
Dim intval As Integer

Private Sub Command1_Click()
  con.Execute ("UPDATE " & opening.comboval & " set SIZ = " & Val(Text4.Text))
  final.Hide
  Main.Form_Load
  Main.Show
End Sub

Private Sub Command2_Click()
    intval = Val(Text2.Text) * Val(Text3.Text)
    Text4.Text = intval
End Sub

Private Sub Command3_Click()
     con.Execute ("UPDATE " & opening.comboval & " set SIZ = " & Val(Text4.Text))
    final.Hide
    reportsiz.Show
End Sub

Private Sub Form_Load()
Text1.Text = 0.65 + 0.01 * Val(di3.Text5.Text)
Text2.Text = Val(Text1.Text) * Val(total.Text21.Text)
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\data.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open opening.comboval, con, adOpenDynamic, adLockOptimistic
End Sub
