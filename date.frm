VERSION 5.00
Begin VB.Form date 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton caleff 
      Caption         =   "Actual Effort Taken"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton subeffort 
      Caption         =   "Submit Effort"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Enter actual Project starting date"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Enter actual  Finishing Date of  the project"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "date"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim diff As Integer
Dim mean As Long
Dim esteff As Long

Private Sub caleff_Click()
diff = DateDiff("d", Text2.Text, Text1.Text)
Text3.Text = diff
con.Execute ("UPDATE " & opening.comboval & " set CALEFF = " & diff)
mean = (esteff + diff) / 2
con.Close
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\data.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open opening.comboval, con, adOpenDynamic, adLockOptimistic
esteff = rs("ESTEFF")
rs.Close
End Sub

Private Sub subeffort_Click()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\data.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset

     rs.Open "dom", con, adOpenDynamic, adLockOptimistic
'rs.MoveFirst
'If (rs("code name") = opening.comboval) Then
     'con.Execute ("UPDATE " & dom & " set  esteff=" & esteff & " , calceff = " & diff & ",mean =" & mean)
    rs.AddNew
      rs("domain") = opening.domname
      rs("code name") = opening.comboval
      rs("esteff") = esteff
      rs("calceff") = diff
      rs("mean") = mean
      rs.Update
     rs.Close
con.Close
date.Hide

MsgBox ("Effort submitted for historical analysis ")
reporteff.Show
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Not IsDate(Text1.Text) Then
MsgBox ("Invalid Date format ")
End If
End Sub

