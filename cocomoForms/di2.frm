VERSION 5.00
Begin VB.Form di2 
   Caption         =   "Processing Complexity"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5205
   LinkTopic       =   "Form10"
   ScaleHeight     =   4470
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Processing Complexity"
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3120
         TabIndex        =   1
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3120
         TabIndex        =   4
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3120
         TabIndex        =   5
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   $"di2.frx":0000
         Height          =   1095
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label2 
         Caption         =   "Online Data Entry"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "End User Efficiency"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Online Update"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Complex Processing"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Reuseability"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Width           =   2055
      End
   End
End
Attribute VB_Name = "di2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
     di2.Hide
     di3.Show
End Sub

Private Sub Text1_Validate(Cancel As Boolean)

If Val(Text1.Text) < 0 Or Val(Text1.Text) > 5 Then
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
If Val(Text2.Text) < 0 Or Val(Text2.Text) > 5 Then
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
If Val(Text3.Text) < 0 Or Val(Text3.Text) > 5 Then
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
If Val(Text4.Text) < 0 Or Val(Text4.Text) > 5 Then
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
If Val(Text5.Text) < 0 Or Val(Text5.Text) > 5 Then
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

