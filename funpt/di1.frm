VERSION 5.00
Begin VB.Form di1 
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5850
   LinkTopic       =   "Form9"
   ScaleHeight     =   4590
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Processing Complexity"
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3120
         TabIndex        =   13
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3120
         TabIndex        =   11
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3120
         TabIndex        =   9
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3120
         TabIndex        =   8
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   7
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Transaction Rate"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Heavily Used Configuration"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Performance"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Distributed Functions"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Data Communications"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   $"di1.frx":0000
         Height          =   1095
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "di1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()



     di1.Hide
     di2.Show
End Sub


Private Sub Text1_Validate(Cancel As Boolean)

If val(Text1.Text) < 0 Or val(Text1.Text) > 5 Then
MsgBox ("value out of range!")
Cancel = True
Text1.Text = ""

ElseIf Text1.Text = "" Then
MsgBox ("No value entered")
Cancel = True
Text1.Text = ""
ElseIf Not IsNumeric(Text1.Text) Then
MsgBox ("Invalid data type ")
Cancel = True
Text1.Text = ""
ElseIf Len(Text1.Text) > 1 Then
MsgBox ("Invalid data")
Cancel = True
Text1.Text = ""
End If
End Sub


Private Sub Text2_Validate(Cancel As Boolean)
If val(Text2.Text) < 0 Or val(Text2.Text) > 5 Then
MsgBox ("value out of range!")
Cancel = True
Text2.Text = ""
ElseIf Text2.Text = "" Then
MsgBox ("No value entered")
Cancel = True
Text2.Text = ""
ElseIf Not IsNumeric(Text2.Text) Then
MsgBox ("Invalid data type ")
Cancel = True
Text2.Text = ""
ElseIf Len(Text2.Text) > 1 Then
MsgBox ("Invalid data")
Cancel = True
Text2.Text = ""
End If
End Sub


Private Sub Text3_Validate(Cancel As Boolean)
If val(Text3.Text) < 0 Or val(Text3.Text) > 5 Then
MsgBox ("value out of range!")
Cancel = True
Text3.Text = ""
ElseIf Text3.Text = "" Then
MsgBox ("No value entered")
Cancel = True
Text3.Text = ""
ElseIf Not IsNumeric(Text3.Text) Then
MsgBox ("Invalid data type ")
Cancel = True
Text3.Text = ""
ElseIf Len(Text3.Text) > 1 Then
MsgBox ("Invalid data")
Cancel = True
Text3.Text = ""
End If
End Sub


Private Sub Text4_Validate(Cancel As Boolean)
If val(Text4.Text) < 0 Or val(Text4.Text) > 5 Then
MsgBox ("value out of range!")
Cancel = True
Text4.Text = ""
ElseIf Text4.Text = "" Then
MsgBox ("No value entered")
Cancel = True
Text4.Text = ""
ElseIf Not IsNumeric(Text4.Text) Then
MsgBox ("Invalid data type ")
Cancel = True
Text4.Text = ""
ElseIf Len(Text4.Text) > 1 Then
MsgBox ("Invalid data")
Cancel = True
Text4.Text = ""
End If
End Sub


Private Sub Text5_Validate(Cancel As Boolean)
If val(Text5.Text) < 0 Or val(Text5.Text) > 5 Then
MsgBox ("value out of range!")
Cancel = True
Text5.Text = ""
ElseIf Text5.Text = "" Then
MsgBox ("No value entered")
Cancel = True
Text5.Text = ""
ElseIf Not IsNumeric(Text5.Text) Then
MsgBox ("Invalid data type ")
Cancel = True
Text5.Text = ""
ElseIf Len(Text5.Text) > 1 Then
MsgBox ("Invalid data")
Cancel = True
Text5.Text = ""
End If
End Sub
