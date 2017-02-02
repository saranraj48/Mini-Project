VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form7"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form7"
   ScaleHeight     =   6540
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text5 
      DataField       =   "Complaint Number"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   8760
      TabIndex        =   11
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\SARAN\Documents\BPO mini project\DB1.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Mytable"
      Top             =   7440
      Width           =   1980
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   495
      Left            =   9960
      TabIndex        =   8
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataField       =   "City"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   8760
      TabIndex        =   3
      Top             =   6120
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      DataField       =   "Address"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   8760
      TabIndex        =   2
      Top             =   5040
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      DataField       =   "PhoneNo"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   8760
      TabIndex        =   1
      Top             =   4080
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   8880
      TabIndex        =   0
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "Complaint Number"
      Height          =   615
      Left            =   3360
      TabIndex        =   10
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label Label5 
      Caption         =   "CUSTOMER DETAILS"
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
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label4 
      Caption         =   "City"
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   6360
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Phone No"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   4080
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   2760
      Width           =   3255
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form6.Show
Me.Hide

End Sub
