VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form register 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Register Here!"
   ClientHeight    =   10260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17775
   LinkTopic       =   "Form1"
   ScaleHeight     =   10260
   ScaleWidth      =   17775
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      TabIndex        =   21
      Top             =   6600
      Width           =   4575
   End
   Begin MSAdodcLib.Adodc registerado 
      Height          =   495
      Left            =   15960
      Top             =   9360
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "register"
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
   Begin VB.TextBox txtcontact 
      DataField       =   "phoneno"
      DataSource      =   "registerado"
      Height          =   495
      Left            =   4320
      TabIndex        =   20
      ToolTipText     =   "Enter your Contact Number"
      Top             =   8520
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
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
      Height          =   735
      Left            =   10920
      TabIndex        =   18
      Top             =   5160
      Width           =   4575
   End
   Begin VB.TextBox txtprof 
      DataField       =   "profid"
      DataSource      =   "registerado"
      Height          =   495
      Left            =   4320
      TabIndex        =   17
      ToolTipText     =   "Enter your Prof ID"
      Top             =   7680
      Width           =   3495
   End
   Begin VB.TextBox txtmail 
      DataField       =   "mailid"
      DataSource      =   "registerado"
      Height          =   495
      Left            =   4320
      TabIndex        =   16
      ToolTipText     =   "Enter your Personal Mail ID"
      Top             =   6840
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      TabIndex        =   13
      Top             =   3600
      Width           =   4575
   End
   Begin VB.TextBox txtaddress 
      DataField       =   "address"
      DataSource      =   "registerado"
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      ToolTipText     =   "Enter your Address"
      Top             =   6000
      Width           =   3495
   End
   Begin VB.TextBox txtdob 
      DataField       =   "dob"
      DataSource      =   "registerado"
      Height          =   495
      Left            =   4320
      TabIndex        =   11
      ToolTipText     =   "Enter your Date of Birth"
      Top             =   5160
      Width           =   3495
   End
   Begin VB.TextBox txtconfirmpassword 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   10
      ToolTipText     =   "Enter the Password Agin"
      Top             =   4320
      Width           =   3495
   End
   Begin VB.TextBox txtpassword 
      DataField       =   "password"
      DataSource      =   "registerado"
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   9
      ToolTipText     =   "Enter the Password"
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox txtlastname 
      DataField       =   "lastname"
      DataSource      =   "registerado"
      Height          =   495
      Left            =   4320
      TabIndex        =   8
      ToolTipText     =   "Enter your Second Name"
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox txtfirstname 
      DataField       =   "firstname"
      DataSource      =   "registerado"
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      ToolTipText     =   "Enter your First Name"
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   19
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   15
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   14
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Prof.Id Number"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   7680
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Mail Id"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Address"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCOUNT REGISTRATION"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   0
      Top             =   480
      Width           =   6495
   End
End
Attribute VB_Name = "register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
registerado.Recordset.Fields("firstname") = txtfirstname.Text
registerado.Recordset.Fields("lastname") = txtlastname.Text
registerado.Recordset.Fields("password") = txtpassword.Text
registerado.Recordset.Fields("dob") = txtdob.Text
registerado.Recordset.Fields("address") = txtaddress.Text
registerado.Recordset.Fields("mailid") = txtmail.Text
registerado.Recordset.Fields("profid") = txtprof.Text
registerado.Recordset.Fields("phoneno") = txtcontact.Text
registerado.Recordset.Update
MsgBox ("Account Created")
txtfirstname.Text = ""
txtlastname.Text = ""
txtpassword.Text = ""
txtdob.Text = ""
txtaddress.Text = ""
txtmail.Text = ""
txtprof.Text = ""
txtcontact.Text = ""
txtconfirmpassword = ""
login.Show
register.Hide
End Sub

Private Sub Command2_Click()
login.Show
register.Hide
End Sub

Private Sub Command3_Click()
txtfirstname.Text = ""
txtlastname.Text = ""
txtpassword.Text = ""
txtdob.Text = ""
txtaddress.Text = ""
txtmail.Text = ""
txtprof.Text = ""
txtcontact.Text = ""
txtconfirmpassword = ""
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
registerado.Recordset.AddNew
End Sub
