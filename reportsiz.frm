VERSION 5.00
Begin VB.Form reportsiz 
   Caption         =   "Generate Report"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Report for Size  Estimates"
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton ok 
         Caption         =   "Exit"
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label esize 
         BackColor       =   &H8000000A&
         Height          =   495
         Left            =   3000
         TabIndex        =   7
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label domain 
         BackColor       =   &H8000000A&
         Height          =   495
         Left            =   3000
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Domain"
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Project Code"
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label pname 
         BackColor       =   &H8000000A&
         Height          =   495
         Left            =   3000
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Estimated Size using Functional Point Analysis"
         Height          =   615
         Left            =   600
         TabIndex        =   1
         Top             =   1920
         Width           =   1575
      End
   End
End
Attribute VB_Name = "reportsiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
pname.Caption = opening.comboval
domain.Caption = opening.domname
esize.Caption = final.Text4
End Sub

Private Sub ok_Click()
reportsiz.Hide
time.Show
End Sub
