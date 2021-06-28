VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LOG IN PAGE"
   ClientHeight    =   10215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   17175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSAdodcLib.Adodc loginado 
      Height          =   330
      Left            =   15120
      Top             =   9000
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\access\registerform.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\access\registerform.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from register"
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
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   6360
      Picture         =   "login.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   3315
      TabIndex        =   7
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtpassword 
      BeginProperty Font 
         Name            =   "Monospac821 BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      ToolTipText     =   "Enter your Password"
      Top             =   5880
      Width           =   6855
   End
   Begin VB.TextBox txtmail 
      BeginProperty Font 
         Name            =   "Monospac821 BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      ToolTipText     =   "Enter your E-Mail ID"
      Top             =   4320
      Width           =   6855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Admin"
      Height          =   375
      Left            =   14760
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton rsgister 
      Caption         =   "Register now?"
      BeginProperty Font 
         Name            =   "Harlow Solid Italic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   8040
      Width           =   3015
   End
   Begin VB.CommandButton loginbut 
      BackColor       =   &H00C0C0FF&
      Caption         =   "LOG IN"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   0
      Top             =   7080
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Monospac821 BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   5280
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail ID"
      BeginProperty Font 
         Name            =   "Monospac821 BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   3720
      Width           =   3855
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
AdminLogin.Show
login.Hide
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End Sub
Private Sub loginbut_Click()
loginado.RecordSource = "select * from register where mailid='" + txtmail.Text + "' and password='" + txtpassword.Text + "' "
loginado.Refresh
If loginado.Recordset.EOF Then
MsgBox "Entered E-Mail or Password is worng, Try Again", vbCritical, "Login field"
login.Show
Else
txtmail.Text = ""
txtpassword.Text = ""
trainDetail.Show
MsgBox "Welcome Again", vbInformation, "Greeting"
login.Hide
End If
End Sub



Private Sub rsgister_Click()
register.Show
login.Hide
End Sub
