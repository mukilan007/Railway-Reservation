VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form splash 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Welcome to Indian Railway"
   ClientHeight    =   10845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18435
   LinkTopic       =   "form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10845
   ScaleWidth      =   18435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   18435
      _ExtentX        =   32517
      _ExtentY        =   741
      Appearance      =   1
      _Version        =   327682
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   15360
      Top             =   1800
   End
   Begin VB.PictureBox Picture1 
      Height          =   3135
      Left            =   6600
      Picture         =   "splash.frx":0000
      ScaleHeight     =   3075
      ScaleWidth      =   3555
      TabIndex        =   3
      Top             =   1680
      Width           =   3615
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   4920
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   1085
      _Version        =   327682
      Appearance      =   1
      Max             =   1000
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   10110
      Width           =   18435
      _ExtentX        =   32517
      _ExtentY        =   1296
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label percent 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   14640
      TabIndex        =   4
      Top             =   4440
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   6015
      Left            =   1440
      Top             =   2160
      Width           =   14535
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   7920
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = True
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1 + 10
percent.Caption = ProgressBar1.Value / 10 & "%"
If ProgressBar1.Value = ProgressBar1.Max Then
Timer1.Enabled = False
Unload Me
login.Show
End If
End Sub
