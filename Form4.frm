VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFF00&
   Caption         =   "Customer Details3"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form4"
   ScaleHeight     =   7515
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text5 
      DataField       =   "Complaint Number"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   8640
      TabIndex        =   14
      Text            =   "7893435533"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\SARAN\Documents\BPO mini project\DB1.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   11520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Mytable"
      Top             =   7680
      Width           =   2340
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Back"
      Height          =   495
      Left            =   1200
      TabIndex        =   11
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   495
      Left            =   8880
      TabIndex        =   10
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Submit"
      Height          =   495
      Left            =   5880
      TabIndex        =   9
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New"
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   7680
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      DataField       =   "City"
      DataSource      =   "Data1"
      Height          =   525
      Left            =   8760
      TabIndex        =   7
      Text            =   "vellore"
      Top             =   6120
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      DataField       =   "Address"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   8760
      TabIndex        =   6
      Text            =   "villupuram"
      Top             =   4560
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      DataField       =   "PhoneNo"
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   8760
      TabIndex        =   5
      Text            =   "9442567415"
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   8760
      TabIndex        =   0
      Text            =   "saran"
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Label Label6 
      Caption         =   "Complaint Number"
      Height          =   615
      Left            =   2640
      TabIndex        =   13
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "customer details"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   615
      Left            =   5640
      TabIndex        =   12
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label4 
      Caption         =   "City"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   6120
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Phone No"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   2160
      Width           =   3135
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Data1.Recordset.AddNew
End Sub

Private Sub Command2_Click()
Form7.Show
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
