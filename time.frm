VERSION 5.00
Begin VB.Form time 
   Caption         =   "Form1"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   2040
      Top             =   1800
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "     THANKYOU FOR USING ESTIMATA        "
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "time"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()
End
End Sub
