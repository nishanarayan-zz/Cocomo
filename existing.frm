VERSION 5.00
Begin VB.Form existing 
   Caption         =   "Select Existing Project"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "domain"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Select Project"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "existing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim tabnam As String
Private Sub Command1_Click()
    opening.domname = Combo2.List(Combo2.ListIndex)
    opening.comboval = Combo1.List(Combo1.ListIndex)
    existing.Hide
    first.Form_Load
    first.Show
End Sub

Public Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = c:\8th sem project\data.mdb;Persist Security Info = false"
Set rs = con.OpenSchema(adSchemaTables)
While Not rs.EOF
    tabnam = rs("TABLE_NAME")
    If ((InStr(tabnam, "MSys") = 0) And (InStr(tabnam, "dom") = 0) And (InStr(tabnam, "domin") = 0) And (InStr(tabnam, "pam") = 0)) Then
        Combo1.AddItem (tabnam)
    End If
    rs.MoveNext
Wend
rs.Close

rs.Open "domin", con, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    Combo2.AddItem (rs("domain"))
    rs.MoveNext
Wend
rs.Close


End Sub
