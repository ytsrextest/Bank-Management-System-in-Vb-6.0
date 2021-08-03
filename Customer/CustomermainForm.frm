VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form CustomermainForm 
   Caption         =   "Customer Dashboars | YT Bank Of India"
   ClientHeight    =   6851
   ClientLeft      =   195
   ClientTop       =   793
   ClientWidth     =   11778
   Icon            =   "CustomermainForm.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "CustomermainForm.frx":1084A
   ScaleHeight     =   6851
   ScaleWidth      =   11778
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Transaction History"
      ForeColor       =   &H000000FF&
      Height          =   1651
      Left            =   11349
      TabIndex        =   4
      Top             =   117
      Width           =   5746
      Begin VB.Image Image4 
         Height          =   715
         Left            =   3042
         Picture         =   "CustomermainForm.frx":24110
         Stretch         =   -1  'True
         Top             =   468
         Width           =   2353
      End
      Begin VB.Image Image3 
         Height          =   715
         Left            =   234
         Picture         =   "CustomermainForm.frx":25813
         Stretch         =   -1  'True
         Top             =   468
         Width           =   2353
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9126
      Top             =   3627
   End
   Begin MSAdodcLib.Adodc fecthcustomerdetails 
      Height          =   299
      Left            =   5733
      Top             =   3627
      Visible         =   0   'False
      Width           =   1417
      _ExtentX        =   2612
      _ExtentY        =   551
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
      Connect         =   "Provider=MSDAORA.1;Password=9122335311;User ID=bankadmin;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=9122335311;User ID=bankadmin;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "bankadmin"
      Password        =   "9122335311"
      RecordSource    =   "select * from CUSTOMER_DETAILS"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Active Services"
      ForeColor       =   &H000000FF&
      Height          =   1651
      Left            =   585
      TabIndex        =   6
      Top             =   117
      Width           =   5746
      Begin VB.Image Image1 
         Height          =   715
         Left            =   2925
         Picture         =   "CustomermainForm.frx":26E03
         Stretch         =   -1  'True
         Top             =   468
         Width           =   2353
      End
      Begin VB.Image Image5 
         Height          =   715
         Left            =   234
         Picture         =   "CustomermainForm.frx":28244
         Stretch         =   -1  'True
         Top             =   468
         Width           =   2353
      End
   End
   Begin VB.Image Image6 
      Height          =   4459
      Left            =   -351
      Picture         =   "CustomermainForm.frx":29582
      Stretch         =   -1  'True
      Top             =   5031
      Width           =   4810
   End
   Begin VB.Image Image2 
      Height          =   4693
      Left            =   12870
      Picture         =   "CustomermainForm.frx":8DD8B
      Stretch         =   -1  'True
      Top             =   4914
      Width           =   5161
   End
   Begin VB.Label LabelAccount 
      Caption         =   "AccountNo"
      DataField       =   "ACCOUNT_NO"
      DataSource      =   "fecthcustomerdetails"
      Height          =   247
      Left            =   8073
      TabIndex        =   5
      Top             =   3276
      Visible         =   0   'False
      Width           =   1066
   End
   Begin VB.Label availablefundinaccount 
      Caption         =   "Available Fund"
      DataField       =   "ACCOUNT_BALLANCE"
      DataSource      =   "fecthcustomerdetails"
      Height          =   247
      Left            =   9945
      TabIndex        =   3
      Top             =   3627
      Visible         =   0   'False
      Width           =   1651
   End
   Begin VB.Label Label2 
      DataField       =   "NAME"
      DataSource      =   "fecthcustomerdetails"
      Height          =   247
      Left            =   7371
      TabIndex        =   2
      Top             =   3627
      Visible         =   0   'False
      Width           =   1534
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   1534
      Left            =   5382
      TabIndex        =   1
      Top             =   5499
      Width           =   7150
   End
   Begin VB.Label Username 
      Height          =   247
      Left            =   3861
      TabIndex        =   0
      Top             =   3627
      Visible         =   0   'False
      Width           =   1534
   End
   Begin VB.Menu home 
      Caption         =   "Home"
      Begin VB.Menu customersupport 
         Caption         =   "Customer Support"
      End
   End
   Begin VB.Menu profile 
      Caption         =   "Customer Profile "
   End
   Begin VB.Menu editrofile 
      Caption         =   "Edit Customer Profile"
   End
   Begin VB.Menu logout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "CustomermainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub backtologin_Click()

End Sub

Private Sub customersupport_Click()
contactfrm.Show
End Sub

Private Sub editrofile_Click()
UdateCustomerDialog.Show

End Sub

Private Sub Form_Load()
    Username.Caption = CustomerfrmLogin.customertxtUserName
    fecthcustomerdetails.RecordSource = "select * from CUSTOMER_DETAILS where USERNAME='" + Username.Caption + "'"
    fecthcustomerdetails.Refresh
    
    DE1Deposite.rsShowdeositehistory.Open "select * from DEPOSITE where ACCOUNT_NO= '" + LabelAccount + "'"
    DataReport_DepositeHistory.Refresh
    
    DE2Withdraw.rswidhdrawfetch.Open "select * from WITHDRAW where ACCOUNT_NO= '" + LabelAccount + "'"
    DataReport_WithdrawHistory.Refresh
   

End Sub

Private Sub Image1_Click()
WithdrawDialog.Show
End Sub

Private Sub Image3_Click()
DataReport_DepositeHistory.Show
DataReport_WithdrawHistory.Hide

End Sub

Private Sub Image4_Click()
DataReport_WithdrawHistory.Show
DataReport_DepositeHistory.Hide

End Sub

Private Sub Image5_Click()
DepositeDialog.Show
End Sub

Private Sub logout_Click()
    End

End Sub

Private Sub profile_Click()
Customerprofile.Show

End Sub

Private Sub Timer1_Timer()
fecthcustomerdetails.Refresh

Label1.Caption = "Welcome " + Label2 + " in Customer Dashboard" & vbNewLine & "Your Account Balance Is: " + availablefundinaccount + " Rs." & vbNewLine & " Date: " & Date & ", Time: " & Time


End Sub
