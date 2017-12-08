VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form extinp 
   Caption         =   "Function Count"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form2"
   ScaleHeight     =   3405
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Function Count"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   360
         Top             =   1920
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1085
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "extinp.frx":0000
         Left            =   480
         List            =   "extinp.frx":000D
         TabIndex        =   5
         Text            =   "Requirements Input"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   495
         Left            =   3720
         TabIndex        =   4
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   1920
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "extinp.frx":0077
         Left            =   3480
         List            =   "extinp.frx":0084
         TabIndex        =   2
         Text            =   "Simple"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "External Input"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "extinp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
rs.AddNew
rs("requirement") = Combo2.List(Combo2.ListIndex)
rs("complexity") = Combo1.List(Combo1.ListIndex)
rs.Update
extinp.Hide
extinp.Show
End Sub

Private Sub Command2_Click()
    rs.AddNew
    rs("requirement") = Combo2.List(Combo2.ListIndex)
    rs("complexity") = Combo1.List(Combo1.ListIndex)
    rs.Update
    extinp.Hide
    extout.Show
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\8th sem project\funpt\softest.mdb;Persist Security Info = false"
Set rs = New ADODB.Recordset
rs.Open "extinp", con, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Text1_Click()
      Text1.Text = ""
End Sub

