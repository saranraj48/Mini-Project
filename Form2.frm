VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFF00&
   Caption         =   "Customer Details1"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form2"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text5 
      DataField       =   "Complaint Number"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   8640
      TabIndex        =   14
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\SARAN\Documents\BPO mini project\DB1.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   11760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Mytable"
      Top             =   7920
      Width           =   1860
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Back"
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataField       =   "City"
      DataSource      =   "Data1"
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   8640
      TabIndex        =   10
      Top             =   6240
      Width           =   3855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   495
      Left            =   8880
      TabIndex        =   8
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Submit"
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   7680
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      DataField       =   "Address"
      DataSource      =   "Data1"
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   8640
      TabIndex        =   2
      Top             =   5040
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      DataField       =   "PhoneNo"
      DataSource      =   "Data1"
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   8640
      TabIndex        =   1
      Top             =   3720
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "Data1"
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   8640
      TabIndex        =   0
      Top             =   2880
      Width           =   3855
   End
   Begin VB.Label Label6 
      Caption         =   "Compalint Number"
      Height          =   615
      Left            =   2280
      TabIndex        =   13
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF00FF&
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE DETAILS"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   615
      Left            =   5520
      TabIndex        =   12
      Top             =   600
      Width           =   6855
   End
   Begin VB.Label Label4 
      Caption         =   "City"
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   6120
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Phone No"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2880
      Width           =   3135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
End Sub

Private Sub Command2_Click()
Form6.Show
Me.Hide
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Command4_Click()
Form6.Show
End Sub

