VERSION 5.00
Begin VB.Form opening 
   Caption         =   "Estimata "
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "ESTIMATE EXISTING PROJECT"
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ESTIMATE A NEW PROJECT"
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   Estimata:   A tool for Software Project Estimation "
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "opening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public domname As String
Public comboval As String

Private Sub Command1_Click()
    enter.Show
    opening.Hide
End Sub

Private Sub Command2_Click()
opening.Hide
existing.Show
End Sub

Public Sub Form_Load()

End Sub
