VERSION 5.00
Begin VB.Form ACM 
   Caption         =   "Application Composition Model"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6135
   LinkTopic       =   "Form5"
   ScaleHeight     =   4605
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Application Composition Model"
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5655
      Begin VB.CommandButton Command2 
         Caption         =   "Update"
         Height          =   735
         Left            =   720
         TabIndex        =   9
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   2640
         TabIndex        =   8
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4320
         TabIndex        =   6
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit and Generate Report on size and effort estimation"
         Height          =   735
         Left            =   2640
         TabIndex        =   5
         Top             =   3000
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "5.frx":0000
         Left            =   2640
         List            =   "5.frx":0013
         TabIndex        =   4
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Effort = NOP/PROD"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Productivity (PROD)"
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Number of Object Points"
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
   End
End
Attribute VB_Name = "ACM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sum As Long
Dim x As Long
Dim i As Integer








Private Sub Combo1_Click()
If (Combo1.ListIndex = 4) Then
Text2.Text = (x + (x * 0.1))
End If
If (Combo1.ListIndex = 3) Then
Text2.Text = (x + (x * 0.2))
End If
If (Combo1.ListIndex = 2) Then
Text2.Text = x
End If
If (Combo1.ListIndex = 1) Then
Text2.Text = (x - (x * 0.2))
End If
If (Combo1.ListIndex = 0) Then
Text2.Text = (x - (x * 0.1))
End If
End Sub

Private Sub Command1_Click()
    con.Execute ("UPDATE " & opening.comboval & " set ACM = " & Val(Text3.Text) & " ,ESTEFF = " & Val(Text3.Text))
    ACM.Hide
reporteff.Show
    'opening.Form_Load
    'opening.Show
End Sub
Private Sub Command2_Click()
    Text3.Text = Val(Text1.Text) / Val(Text2.Text)
    con.Execute ("UPDATE " & opening.comboval & " set ACM = " & Val(Text3.Text) & " ,ESTEFF = " & Val(Text3.Text))
End Sub
Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\data.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open "dom", con, adOpenDynamic, adLockOptimistic


sum = 0
i = 0

Do Until (rs.EOF)

    If (rs("domain") = opening.domname) Then
        sum = sum + rs("mean")
        i = i + 1
        rs.MoveNext
    Else
        rs.MoveNext
    End If
Loop
If (sum = 0) Then
  x = 75
Else
  x = sum / i
End If

End Sub

