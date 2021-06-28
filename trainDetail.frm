VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form trainDetail 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TRAIN DETAILS"
   ClientHeight    =   9990
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17655
   BeginProperty Font 
      Name            =   "Monospac821 BT"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   17655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Height          =   495
      Left            =   8760
      TabIndex        =   11
      Top             =   8040
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      DataField       =   "trainname"
      DataSource      =   "traindetailado"
      Height          =   495
      Left            =   8160
      TabIndex        =   9
      Top             =   1680
      Width           =   3855
   End
   Begin VB.CommandButton view 
      Caption         =   "View All"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   8040
      Width           =   2895
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
      Height          =   495
      Left            =   12600
      TabIndex        =   7
      Top             =   8040
      Width           =   2415
   End
   Begin VB.CommandButton reservate 
      Caption         =   "Reservate"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1200
      TabIndex        =   6
      Top             =   8040
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc traindetailado 
      Height          =   375
      Left            =   16200
      Top             =   9120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Enabled         =   0
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\access\registerform.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\access\registerform.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from addtraindetails"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "trainDetail.frx":0000
      Height          =   4335
      Left            =   480
      TabIndex        =   5
      Top             =   3120
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Monospac821 BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Monospac821 BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Train Search Board"
      BeginProperty Font 
         Name            =   "Monospac821 BT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   14775
      Begin VB.CommandButton search 
         Caption         =   "Search"
         Height          =   495
         Left            =   11880
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         DataField       =   "trainno"
         DataSource      =   "traindetailado"
         Height          =   495
         Left            =   2280
         TabIndex        =   3
         ToolTipText     =   "Enter the Train Number OR Name"
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "Train Name"
         Height          =   375
         Left            =   6000
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Train Number"
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TRAIN DETAILS"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      TabIndex        =   0
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "trainDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub exit_Click()
login.Show
trainDetail.Hide
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
traindetailado.Recordset.AddNew
End Sub

Private Sub reservate_Click()
reservation.Refresh
reservation.Show
trainDetail.Hide
End Sub

Private Sub reset_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub search_Click()
traindetailado.RecordSource = "select * from addtraindetails where trainno ='" + Text1.Text + "' or trainname ='" + Text2.Text + "' "
traindetailado.Refresh
If traindetailado.Recordset.EOF Then
MsgBox "Data is not found, Enter the other Detail", vbCritical, "Invalid Date"
trainDetail.Show
traindetailado.RecordSource = "select * from addtraindetails"
traindetailado.Refresh
Text1.Text = ""
Text2.Text = ""
Else
traindetailado.Caption = traindetailado.RecordSource
End If
End Sub

Private Sub view_Click()
traindetailado.RecordSource = "select * from addtraindetails"
traindetailado.Refresh
traindetailado.Caption = traindetailado.RecordSource
traindetailado.Recordset.AddNew
End Sub
