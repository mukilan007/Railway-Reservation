VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form billing 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BILL"
   ClientHeight    =   10950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18480
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
   ScaleHeight     =   10950
   ScaleWidth      =   18480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   240
      TabIndex        =   18
      Top             =   840
      Width           =   17535
      Begin VB.CommandButton Command1 
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
         Left            =   15000
         TabIndex        =   27
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   10200
         TabIndex        =   26
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "seat Choosed"
         Height          =   495
         Left            =   7680
         TabIndex        =   25
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lalname 
         BackStyle       =   0  'Transparent
         DataField       =   "trainname"
         DataSource      =   "Adodc1"
         Height          =   615
         Left            =   10200
         TabIndex        =   24
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Train Name"
         Height          =   615
         Left            =   7680
         TabIndex        =   23
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lalfare 
         BackStyle       =   0  'Transparent
         DataField       =   "fare"
         DataSource      =   "Adodc1"
         Height          =   615
         Left            =   2760
         TabIndex        =   22
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label lalnumber 
         BackStyle       =   0  'Transparent
         DataField       =   "trainno"
         DataSource      =   "Adodc1"
         Height          =   615
         Left            =   2760
         TabIndex        =   21
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fare"
         Height          =   615
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Train Number"
         Height          =   615
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   17
      Top             =   8400
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   14400
      Top             =   10080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   794
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
   Begin VB.CheckBox Ch2 
      Caption         =   "Check2"
      Height          =   285
      Left            =   600
      TabIndex        =   15
      Top             =   10200
      Width           =   2535
   End
   Begin VB.CheckBox Ch1 
      Caption         =   "Check1"
      Height          =   285
      Left            =   600
      TabIndex        =   14
      Top             =   9600
      Width           =   2535
   End
   Begin VB.CommandButton print 
      Caption         =   "Confirm && Print"
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
      Left            =   14160
      TabIndex        =   13
      Top             =   7800
      Width           =   3255
   End
   Begin VB.CommandButton exit 
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
      Height          =   735
      Left            =   14160
      TabIndex        =   12
      Top             =   8880
      Width           =   3255
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
      Height          =   735
      Left            =   14160
      TabIndex        =   11
      Top             =   6720
      Width           =   3255
   End
   Begin VB.CommandButton clear 
      Caption         =   "Clear"
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
      Left            =   14160
      TabIndex        =   9
      Top             =   5640
      Width           =   3255
   End
   Begin VB.CommandButton delect 
      Caption         =   "Delect"
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
      Left            =   14160
      TabIndex        =   8
      Top             =   4560
      Width           =   3255
   End
   Begin VB.CommandButton add 
      Caption         =   "Add"
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
      Left            =   14160
      TabIndex        =   7
      Top             =   3480
      Width           =   3255
   End
   Begin VB.ListBox List 
      Height          =   2340
      Left            =   4200
      TabIndex        =   5
      Top             =   4560
      Width           =   3975
   End
   Begin VB.TextBox txtname 
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   3600
      Width           =   3975
   End
   Begin VB.Label lalamount 
      Caption         =   "0"
      Height          =   495
      Left            =   4200
      TabIndex        =   16
      Top             =   8400
      Width           =   3975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   8520
      Width           =   3135
   End
   Begin VB.Label lalcount 
      Caption         =   "0"
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   7440
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Total No.of Passenger"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   7440
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Passenger List"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Passenger Name"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3600
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BILLING"
      BeginProperty Font 
         Name            =   "Monospac821 BT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "billing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total_amount As Integer
Dim flag As Integer

Private Sub add_Click()
If lalcount.Caption < reservation.Text1.Text Then
List.AddItem txtname.Text
txtname.Text = ""
txtname.SetFocus
lalcount.Caption = List.ListCount
total_amount = lalfare * lalcount
lalamount.Caption = total_amount
If lalcount.Caption = reservation.Text1.Text Then
MsgBox "No of Passenger completed", vbInformation, "Information"
add.Enabled = False
txtname.Enabled = False
End If
Else
add.Enabled = False
End If
End Sub

Private Sub back_Click()
reservation.Show
billing.Hide
End Sub

Private Sub clear_Click()
List.clear
lalcount.Caption = 0
lalamount.Caption = " "
End Sub
Private Sub Command1_Click()
Adodc1.RecordSource = "select * from addtraindetails where trainno ='" + reservation.lalnumber.Caption + "' and trainname='" + reservation.lalname.Caption + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
Else
Adodc1.Caption = Adodc1.RecordSource
flag = 1
End If
End Sub

Private Sub delect_Click()
Dim x As Integer
x = List.ListIndex
If x >= 0 Then
List.RemoveItem x
lalcount.Caption = List.ListCount
total_amount = lalfare * lalcount
lalamount.Caption = total_amount
txtname.Enabled = True
End If
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
add.Enabled = False
flag = 0
End Sub

Private Sub print_Click()
If flag = 1 Then
If lalcount = reservation.Text1.Text Then
If lalcount > 0 Then
MsgBox "Reservation Conformed", vbInformation, "Happy Journey"
Else
MsgBox "List is empty...", vbCritical, "Information"
End If
Else
MsgBox "Uneven Traveller", vbCritical, "Message"
End If
Else
MsgBox "use Auto Fill", vbInformation
End If
End Sub

Private Sub txtname_Change()
add.Enabled = (Len(txtname.Text) > 0)
End Sub

