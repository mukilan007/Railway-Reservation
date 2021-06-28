VERSION 5.00
Begin VB.Form admin 
   BackColor       =   &H00FFC0FF&
   Caption         =   "ADMINISTRATION"
   ClientHeight    =   10605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17955
   LinkTopic       =   "Form1"
   ScaleHeight     =   10605
   ScaleWidth      =   17955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Delect Train"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   10320
      TabIndex        =   3
      Top             =   4560
      Width           =   4800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   10320
      TabIndex        =   2
      Top             =   6840
      Width           =   4800
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   2880
      Picture         =   "admin.frx":0000
      ScaleHeight     =   3075
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   3000
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Train"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   10320
      TabIndex        =   0
      Top             =   2040
      Width           =   4800
   End
End
Attribute VB_Name = "admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
addtrain.Show
login.Hide
register.Hide
admin.Hide
End Sub

Private Sub Command2_Click()
delectTrain.Show
login.Hide
register.Hide
admin.Hide
End Sub

Private Sub Command3_Click()
login.Show
admin.Hide
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End Sub
