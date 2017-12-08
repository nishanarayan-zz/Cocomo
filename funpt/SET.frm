VERSION 5.00
Begin VB.Form first 
   Caption         =   "Software Estimation Tool"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4155
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cocomo 
      Caption         =   "COCOMO II"
      Enabled         =   0   'False
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "This model carries out the effort estimation"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton fp 
      Caption         =   "FUNCTIONAL POINT"
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "This model carries out size estimation"
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "first"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset





Private Sub cocomo_Click()
first.Hide
    Main.Show
End Sub



Public Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\data.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open opening.comboval, con, adOpenDynamic, adLockOptimistic
If (rs("SIZ") = -1) Then
    cocomo.Enabled = False
Else
    cocomo.Enabled = True
End If
'If (rs("ESTEFF") = -1) Then
    'Command1.Enabled = False
'Else
    'Command1.Enabled = True
'End If
End Sub

Private Sub fp_Click()
    extinp.Show
    first.Hide
End Sub
