VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form6"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form6"
   ScaleHeight     =   6300
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Complaint Number"
      Height          =   495
      Left            =   7440
      TabIndex        =   3
      Top             =   3600
      Width           =   4695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   615
      Left            =   7320
      TabIndex        =   2
      Top             =   6000
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update Details"
      Height          =   615
      Left            =   7440
      TabIndex        =   1
      Top             =   4680
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Customer details"
      Height          =   615
      Left            =   7320
      TabIndex        =   0
      Top             =   2160
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BPO Management System"
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
      Height          =   975
      Left            =   6120
      TabIndex        =   4
      Top             =   600
      Width           =   8655
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form7.Show
Me.Hide

End Sub

Private Sub Command2_Click()
Form2.Show
Me.Hide

End Sub

Private Sub Command3_Click()
Form5.Show
Me.Hide
End Sub

Private Sub Command4_Click()
Form3.Show
Me.Hide
End Sub
