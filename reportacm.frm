VERSION 5.00
Begin VB.Form reportacm 
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Report for Size and Effort  Estimates"
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton ok 
         Caption         =   "OK"
         Height          =   495
         Left            =   1800
         TabIndex        =   1
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label eeffort 
         BackColor       =   &H8000000B&
         Height          =   495
         Left            =   3000
         TabIndex        =   9
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Estimated Effort using COCOMO II Estimation"
         Height          =   735
         Left            =   600
         TabIndex        =   8
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Estimated Size using Functional Point Analysis"
         Height          =   615
         Left            =   600
         TabIndex        =   7
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label pname 
         BackColor       =   &H8000000A&
         Height          =   495
         Left            =   3000
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Project Name"
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Domain"
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label domain 
         BackColor       =   &H8000000A&
         Height          =   495
         Left            =   3000
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label esize 
         BackColor       =   &H8000000A&
         Height          =   495
         Left            =   3000
         TabIndex        =   2
         Top             =   1920
         Width           =   1215
      End
   End
End
Attribute VB_Name = "reportacm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Load()
pname.Caption = opening.comboval
domain.Caption = opening.domname
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\data.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open opening.comboval, con, adOpenDynamic, adLockOptimistic
esize.Caption = rs("SIZ")
eeffort.Caption = rs("ESTEFF")
rs.Close
End Sub
