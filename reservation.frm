VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form reservation 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RESERVATION"
   ClientHeight    =   10140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17235
   BeginProperty Font 
      Name            =   "Monospac821 BT"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   17235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   25
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13800
      TabIndex        =   22
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4320
      TabIndex        =   20
      Top             =   6840
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "TRAIN DETAIL"
      Height          =   3495
      Left            =   840
      TabIndex        =   7
      Top             =   960
      Width           =   14415
      Begin VB.CommandButton autofill 
         Caption         =   "Auto fill"
         BeginProperty Font 
            Name            =   "Snap ITC"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12000
         TabIndex        =   21
         Top             =   2760
         Width           =   2175
      End
      Begin VB.ComboBox codate 
         DataField       =   "date"
         DataSource      =   "Adodc1"
         Height          =   405
         Left            =   9240
         TabIndex        =   15
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         DataField       =   "fare"
         DataSource      =   "Adodc1"
         Height          =   615
         Left            =   2400
         TabIndex        =   24
         Top             =   2280
         Width           =   3615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Fare"
         Height          =   615
         Left            =   240
         TabIndex        =   23
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lalseat 
         BackStyle       =   0  'Transparent
         DataField       =   "noofseat"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   9240
         TabIndex        =   18
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label lalname 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   2400
         TabIndex        =   16
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   495
         Left            =   6960
         TabIndex        =   14
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label lalamount 
         BackStyle       =   0  'Transparent
         DataField       =   "fare"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   9240
         TabIndex        =   13
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   495
         Left            =   6960
         TabIndex        =   12
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "No of Seat"
         Height          =   495
         Left            =   6960
         TabIndex        =   11
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Train Name"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lalnumber 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   2400
         TabIndex        =   9
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Train Number"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   15480
      Top             =   9600
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
      RecordSource    =   "select *  from addtraindetails"
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
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      ToolTipText     =   "Enter get down station"
      Top             =   6000
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      ToolTipText     =   "Enter get in Station"
      Top             =   5160
      Width           =   3735
   End
   Begin VB.CommandButton back 
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
      Height          =   615
      Left            =   11400
      TabIndex        =   4
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton reset 
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
      Height          =   615
      Left            =   8400
      TabIndex        =   3
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton bill 
      Caption         =   "Bill"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   2
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Seat to be Booked"
      Height          =   495
      Left            =   1320
      TabIndex        =   19
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RESERVATION"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   17
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Station Down"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Station Up"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   5160
      Width           =   2055
   End
End
Attribute VB_Name = "reservation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Integer
Private Sub autofill_Click()
lalnumber.Caption = trainDetail.Text1.Text
lalname.Caption = trainDetail.Text2.Text
Adodc1.RecordSource = "select * from addtraindetails where trainno ='" + lalnumber.Caption + "' and trainname='" + lalname.Caption + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
Else
Adodc1.Caption = Adodc1.RecordSource
flag = 1
End If
End Sub

Private Sub back_Click()
lalnumber.Caption = ""
lalname.Caption = ""
lalamount.Caption = ""
lalseat.Caption = ""
Text4.Text = ""
Text5.Text = ""
Text1.Text = ""
codate.Text = ""
flag = 0
reservation.Hide
trainDetail.Show
End Sub

Private Sub bill_Click()
If flag = 1 Then
If Text1.Text < lalseat.Caption Then
Text1.Text = Text1.Text
billing.Show
reservation.Hide
Else
MsgBox ("Number of Seat available is less " & lalseat), vbCritical + vbExclamation, "Seat avaiable"
Text1.Text = ""
End If
Else
MsgBox "Use Auto fill", vbInformation
End If
End Sub
Private Sub Command2_Click()
Text1.Text = ""
End Sub


Public Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
Adodc1.Recordset.AddNew
flag = 0
End Sub

Private Sub reset_Click()
Text4.Text = ""
Text5.Text = ""
Text1.Text = ""
codate.Text = ""
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
MsgBox "Please enter number only...", vbCritical
End If
End Sub
