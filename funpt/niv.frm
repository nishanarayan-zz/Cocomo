VERSION 5.00
Begin VB.Form di3 
   Caption         =   "Processing Complexity"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5730
   LinkTopic       =   "Form11"
   ScaleHeight     =   4800
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Processing Complexity"
      Height          =   4335
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton Command2 
         Caption         =   "Update"
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3120
         TabIndex        =   12
         Top             =   3000
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3120
         TabIndex        =   4
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3120
         TabIndex        =   1
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Total Degree of Influence (PC)"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Facilitate Change"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Multiple Sites"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Operational Ease"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Installation ease"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   $"niv.frx":0000
         Height          =   1095
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   4575
      End
   End
End
Attribute VB_Name = "di3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
      di3.Hide
      final.Show
End Sub

Private Sub Command2_Click()
Text5.Text = Val(Form9.Text1.Text) + Val(Form9.Text2.Text) + Val(Form9.Text3.Text) + Val(Form9.Text4.Text) + Val(Form9.Text5.Text) + Val(Form10.Text1.Text) + Val(Form10.Text2.Text) + Val(Form10.Text3.Text) + Val(Form10.Text4.Text) + Val(Form10.Text5.Text) + Val(Form11.Text1.Text) + Val(Form11.Text2.Text) + Val(Form11.Text3.Text) + Val(Form11.Text4.Text)
End Sub


