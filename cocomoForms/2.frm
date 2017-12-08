VERSION 5.00
Begin VB.Form SF 
   Caption         =   "Effort"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3585
   LinkTopic       =   "Form2"
   ScaleHeight     =   3165
   ScaleWidth      =   3585
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame SFa 
      Caption         =   "Effort"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      Begin VB.CommandButton Command1 
         Caption         =   "Exit and Generate Report on Size and Effort Estimation"
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "TDEV"
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Height          =   495
         Left            =   960
         TabIndex        =   1
         Top             =   1680
         Width           =   1335
      End
   End
End
Attribute VB_Name = "SF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim pmadj As Long
Dim tdev As Long
Dim prpo As Integer
Dim prod As Long
Dim b As Long
Dim pnom As Integer

Private Sub Command1_Click()
    SF.Hide
    reporteff.Show
    ' opening.Show
End Sub
Public Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\data.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
 rs.Open "pam", con, adOpenDynamic, adLockOptimistic
   b = rs("b")
   pnom = rs("pnom")
   rs.Close

rs.Open opening.comboval, con, adOpenDynamic, adLockOptimistic
If (rs("EDM") = -1) Then
   pmadj = Val(EDM.pnom) * Val(EDM.prpo)
   tdev = 3.67 * pmadj * (0.28 + 0.2 * (b - 0.91))
   con.Execute ("UPDATE " & opening.comboval & " set EDM= " & tdev & " ,ESTEFF = " & tdev)
   
Else
   pmadj = pnom * Val(PAMpersonnel.prod)
   tdev = 3.67 * pmadj * (0.28 + 0.2 * (b - 0.91))
   'rs.Open opening.comboval, con, adOpenDynamic, adLockOptimistic
   con.Execute ("UPDATE " & opening.comboval & " set PAM= " & tdev & " ,ESTEFF = " & tdev)
End If

Text1.Text = tdev
rs.Close
End Sub
