VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form addTrain 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADD_TRAIN"
   ClientHeight    =   10350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18555
   BeginProperty Font 
      Name            =   "Monospac821 BT"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10350
   ScaleWidth      =   18555
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc adddetailado 
      Height          =   615
      Left            =   14280
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
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
      RecordSource    =   "addtraindetails"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Monospac821 BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton back 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   4800
      TabIndex        =   19
      Top             =   8640
      Width           =   3015
   End
   Begin VB.TextBox txtfare 
      DataField       =   "fare"
      DataSource      =   "adddetailado"
      Height          =   495
      Left            =   4800
      TabIndex        =   18
      ToolTipText     =   "Enter tne Fare"
      Top             =   6720
      Width           =   4935
   End
   Begin VB.TextBox txtseat 
      DataField       =   "noofseat"
      DataSource      =   "adddetailado"
      Height          =   500
      Left            =   4800
      TabIndex        =   17
      ToolTipText     =   "Enter No of Seat"
      Top             =   7560
      Width           =   5000
   End
   Begin VB.TextBox txtarrtime 
      DataField       =   "arrvialtime"
      DataSource      =   "adddetailado"
      Height          =   495
      Left            =   4800
      TabIndex        =   16
      ToolTipText     =   "Enter the Arrival Time"
      Top             =   5040
      Width           =   5055
   End
   Begin VB.TextBox txtdistime 
      DataField       =   "dispatchtime"
      DataSource      =   "adddetailado"
      Height          =   500
      Left            =   4800
      TabIndex        =   13
      ToolTipText     =   "Enter the Dispatch Time"
      Top             =   5880
      Width           =   5000
   End
   Begin VB.TextBox Text5 
      Height          =   500
      Left            =   4800
      TabIndex        =   12
      ToolTipText     =   "Train Arrival Time"
      Top             =   5040
      Width           =   5000
   End
   Begin VB.TextBox txtplatform 
      DataField       =   "platformno"
      DataSource      =   "adddetailado"
      Height          =   500
      Left            =   4800
      TabIndex        =   11
      ToolTipText     =   "Enter the Platform Number"
      Top             =   4200
      Width           =   5000
   End
   Begin VB.TextBox txtdate 
      DataField       =   "date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   3
      EndProperty
      DataSource      =   "adddetailado"
      Height          =   500
      Left            =   4800
      TabIndex        =   10
      ToolTipText     =   "Enter the Date"
      Top             =   3360
      Width           =   5000
   End
   Begin VB.TextBox txtname 
      DataField       =   "trainname"
      DataSource      =   "adddetailado"
      Height          =   500
      Left            =   4800
      TabIndex        =   9
      ToolTipText     =   "Enter the Train Name"
      Top             =   2520
      Width           =   5000
   End
   Begin VB.TextBox txtnumber 
      DataField       =   "trainno"
      DataSource      =   "adddetailado"
      Height          =   500
      Left            =   4800
      TabIndex        =   8
      ToolTipText     =   "Enter the Train Number"
      Top             =   1680
      Width           =   5000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   720
      TabIndex        =   6
      Top             =   8640
      Width           =   3015
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "no of seat"
      Height          =   495
      Left            =   840
      TabIndex        =   15
      Top             =   7560
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Fare"
      Height          =   495
      Left            =   840
      TabIndex        =   14
      Top             =   6720
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Train Details"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   6120
      TabIndex        =   7
      Top             =   240
      Width           =   6615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Dispatch Timing"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   5880
      Width           =   3000
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Arrival Timing"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   5040
      Width           =   3000
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Platform number"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   4200
      Width           =   3000
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   3360
      Width           =   3000
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Train Name"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   2520
      Width           =   3000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Train Number"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   1680
      Width           =   3000
   End
End
Attribute VB_Name = "addtrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub back_Click()
txtnumber.Text = ""
txtname.Text = ""
txtdate.Text = ""
txtplatform.Text = ""
txtarrtime.Text = ""
txtdistime.Text = ""
txtfare.Text = ""
txtseat.Text = ""
admin.Show
addtrain.Hide
End Sub

Private Sub Command1_Click()
adddetailado.Recordset.Fields("trainno") = txtnumber.Text
adddetailado.Recordset.Fields("trainname") = txtname.Text
adddetailado.Recordset.Fields("date") = txtdate.Text
adddetailado.Recordset.Fields("platformno") = txtplatform.Text
adddetailado.Recordset.Fields("arrvialtime") = txtarrtime.Text
adddetailado.Recordset.Fields("dispatchtime") = txtdistime.Text
adddetailado.Recordset.Fields("fare") = txtfare.Text
adddetailado.Recordset.Fields("noofseat") = txtseat.Text
adddetailado.Recordset.Update
If MsgBox("Train Detail Added Successfully...Need to add more", vbYesNo) = vbYes Then
addtrain.Show
txtnumber.Text = ""
txtname.Text = ""
txtdate.Text = ""
txtplatform.Text = ""
txtarrtime.Text = ""
txtdistime.Text = ""
txtfare.Text = ""
txtseat.Text = ""
Else
txtnumber.Text = ""
txtname.Text = ""
txtdate.Text = ""
txtplatform.Text = ""
txtarrtime.Text = ""
txtdistime.Text = ""
txtfare.Text = ""
txtseat.Text = ""
admin.Show
addtrain.Hide
End If
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
adddetailado.Recordset.AddNew
End Sub
