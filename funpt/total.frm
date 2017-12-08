VERSION 5.00
Begin VB.Form total 
   Caption         =   "Function Count"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7560
   LinkTopic       =   "Form7"
   ScaleHeight     =   6300
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Total Unadjusted Function Points"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7215
      Begin VB.CommandButton Update 
         Caption         =   "Update"
         Height          =   495
         Left            =   2520
         TabIndex        =   34
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   495
         Left            =   3960
         TabIndex        =   33
         Top             =   5160
         Width           =   975
      End
      Begin VB.TextBox Text21 
         Height          =   375
         Left            =   6000
         TabIndex        =   32
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   6000
         TabIndex        =   30
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   4680
         TabIndex        =   29
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   3360
         TabIndex        =   28
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   2160
         TabIndex        =   27
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   6000
         TabIndex        =   26
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   4680
         TabIndex        =   25
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   3360
         TabIndex        =   24
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   6000
         TabIndex        =   22
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   4680
         TabIndex        =   21
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   3360
         TabIndex        =   20
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   2160
         TabIndex        =   19
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   6000
         TabIndex        =   18
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   4680
         TabIndex        =   17
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3360
         TabIndex        =   16
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2160
         TabIndex        =   15
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   6000
         TabIndex        =   14
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4680
         TabIndex        =   13
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3360
         TabIndex        =   12
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   11
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "TOTAL (FC)"
         Height          =   375
         Left            =   360
         TabIndex        =   31
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "External Inquiry"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "External Interface file"
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Logical Internal File"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "External Output"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Total"
         Height          =   375
         Left            =   6120
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Complex"
         Height          =   495
         Left            =   4680
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Average"
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Simple"
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Description"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "External Input"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   1320
         Width           =   1335
      End
   End
End
Attribute VB_Name = "total"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmp As Integer
Dim sim As Integer
Dim avg As Integer
Public Sub dosome()
    cmp = 0
    sim = 0
    avg = 0
    rs.MoveFirst
    Do Until rs.EOF
        If (rs("complexity") = "Simple") Then
            sim = sim + 1
        End If
        If (rs("complexity") = "Average") Then
            avg = avg + 1
        End If
        If (rs("complexity") = "Complex") Then
            cmp = cmp + 1
        End If
        rs.MoveNext
    Loop
End Sub

Private Sub Command1_Click()
       rs.Open "extinp", con, adOpenDynamic, adLockOptimistic
       rs.MoveFirst
       Do Until rs.EOF
       rs("requirement") = Null
       rs("complexity") = Null
       rs.MoveNext
       Loop
       rs.Close
       
       rs.Open "extout", con, adOpenDynamic, adLockOptimistic
       rs.MoveFirst
       Do Until rs.EOF
       rs("requirement") = Null
       rs("complexity") = Null
       rs.MoveNext
       Loop
       rs.Close
       
       rs.Open "logint", con, adOpenDynamic, adLockOptimistic
       rs.MoveFirst
       Do Until rs.EOF
       rs("requirement") = Null
       rs("complexity") = Null
       rs.MoveNext
       Loop
         rs.Close
       
       rs.Open "extint", con, adOpenDynamic, adLockOptimistic
       rs.MoveFirst
       Do Until rs.EOF
       rs("requirement") = Null
       rs("complexity") = Null
       rs.MoveNext
       Loop
         rs.Close
       
       rs.Open "extinq", con, adOpenDynamic, adLockOptimistic
       rs.MoveFirst
       Do Until rs.EOF
       rs("requirement") = Null
       rs("complexity") = Null
       rs.MoveNext
       Loop
         rs.Close
       
       
       total.Hide
       di1.Show
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\funpt\softest.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
End Sub

Private Sub Update_Click()
    rs.Open "extinp", con, adOpenDynamic, adLockOptimistic
    dosome
    Text1.Text = (sim * 3)
    Text2.Text = (avg * 4)
    Text3.Text = (cmp * 6)
    Text4.Text = Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text)
    rs.Close
    rs.Open "extout", con, adOpenDynamic, adLockOptimistic
    dosome
    Text5.Text = (sim * 4)
    Text6.Text = (avg * 5)
    Text7.Text = (cmp * 7)
    Text8.Text = Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text)
    rs.Close
    rs.Open "logint", con, adOpenDynamic, adLockOptimistic
    dosome
    Text9.Text = (sim * 7)
    Text10.Text = (avg * 10)
    Text11.Text = (cmp * 15)
    Text12.Text = Val(Text9.Text) + Val(Text10.Text) + Val(Text11.Text)
    rs.Close
    rs.Open "extint", con, adOpenDynamic, adLockOptimistic
    dosome
    Text13.Text = (sim * 5)
    Text14.Text = (avg * 7)
    Text15.Text = (cmp * 10)
    Text16.Text = Val(Text13.Text) + Val(Text14.Text) + Val(Text15.Text)
    rs.Close
    rs.Open "extinq", con, adOpenDynamic, adLockOptimistic
    dosome
    Text17.Text = (sim * 3)
    Text18.Text = (avg * 4)
    Text19.Text = (cmp * 6)
    Text20.Text = Val(Text17.Text) + Val(Text18.Text) + Val(Text19.Text)
    rs.Close
    Text21.Text = Val(Text4.Text) + Val(Text8.Text) + Val(Text12.Text) + Val(Text16.Text) + Val(Text20.Text)


End Sub
