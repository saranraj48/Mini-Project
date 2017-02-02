VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN FORM"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   15930
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Login"
      Height          =   615
      Left            =   6480
      TabIndex        =   4
      Top             =   7080
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   8400
      PasswordChar    =   "*"
      TabIndex        =   3
      Tag             =   "**"
      Top             =   4920
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   8400
      TabIndex        =   2
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "BPO Management System"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Password"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "UserName"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   2520
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "saran" And Text2.Text = "saran" Then
Form6.Show
Me.Hide
End If

End Sub

