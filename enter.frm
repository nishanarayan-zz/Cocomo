VERSION 5.00
Begin VB.Form enter 
   Caption         =   "New Project"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2760
      TabIndex        =   6
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Enter Project Code Name"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Domain"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Project Name"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "enter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim tabnam As String
Dim flag As String



Private Sub Command1_Click()


If Text1.Text = "" Or Text3.Text = "" Or Combo1.ListIndex < 0 Then
    MsgBox "Enter the Required Values", vbCritical
Else
    flag = "n"
    Set rs = con.OpenSchema(adSchemaTables)
    While Not rs.EOF
     tabnam = rs("TABLE_NAME")
     If (InStr(tabnam, "MSys") = 0) Then
         If (Text3.Text = tabnam) Then
             flag = "y"
         End If
     End If
    rs.MoveNext
    Wend
    If flag = "n" Then
     con.Execute ("CREATE TABLE " & Text3.Text & "(CALEFF NUMBER NOT NULL,ACM NUMBER NOT NULL,EDM NUMBER NOT NULL,PAM NUMBER NOT NULL,ESTEFF NUMBER NOT NULL,SIZ NUMBER NOT NULL)")
     Set rs = New ADODB.Recordset
     rs.Open Text3.Text, con, adOpenDynamic, adLockOptimistic
     rs.AddNew
     
     
     rs("CALEFF") = -1
     rs("ESTEFF") = -1
     rs("SIZ") = -1
     rs("ACM") = -1
     rs("EDM") = -1
     rs("PAM") = -1
     rs.Update
     rs.Close
     con.Close
     'rs.Open "dom", con, adOpenDynamic, adLockOptimistic
     'rs.AddNew
     'rs("domain") = Combo1.List(Combo1.ListIndex)
     'rs("code name") = Text3.Text
     'rs.Update
      'rs.Close
     opening.comboval = Text3.Text
     opening.domname = Combo1.List(Combo1.ListIndex)
     enter.Hide
     first.Form_Load
     first.Show
     Else
     Text3.Text = "Enter New Code Name"
     MsgBox "ERROR", vbCritical
    End If
End If


End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = c:\8th sem project\data.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open "domin", con, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    Combo1.AddItem (rs("domain"))
    rs.MoveNext
Wend
rs.Close

End Sub



Private Sub Label1_Click()

End Sub
