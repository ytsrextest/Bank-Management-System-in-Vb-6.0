VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Customerprofile 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Customer Profile"
   ClientHeight    =   6188
   ClientLeft      =   104
   ClientTop       =   416
   ClientWidth     =   6253
   Icon            =   "Customerprofile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6188
   ScaleWidth      =   6253
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc fetchcustomerdetailsado 
      Height          =   364
      Left            =   4212
      Top             =   3276
      Visible         =   0   'False
      Width           =   1768
      _ExtentX        =   3259
      _ExtentY        =   671
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
      UserName        =   ""
      Password        =   ""
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
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "A/C Reg. Date:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   4446
      TabIndex        =   26
      Top             =   2457
      Width           =   1417
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "DOB:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   234
      TabIndex        =   25
      Top             =   3276
      Width           =   949
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "DOB"
      DataField       =   "DOB"
      DataSource      =   "fetchcustomerdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1638
      TabIndex        =   24
      Top             =   3276
      Width           =   2236
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Gvt Id no"
      DataField       =   "GVTID_NO"
      DataSource      =   "fetchcustomerdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1638
      TabIndex        =   23
      Top             =   5148
      Width           =   2470
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Gvt. ID No. :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   234
      TabIndex        =   22
      Top             =   5148
      Width           =   1183
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Gvt Id Tye"
      DataField       =   "GVTID_TYPE"
      DataSource      =   "fetchcustomerdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1638
      TabIndex        =   21
      Top             =   4680
      Width           =   2470
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Gvt. ID Type:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   234
      TabIndex        =   20
      Top             =   4680
      Width           =   1183
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      DataField       =   "PHONENO"
      DataSource      =   "fetchcustomerdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1638
      TabIndex        =   19
      Top             =   4212
      Width           =   2470
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No.:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   234
      TabIndex        =   18
      Top             =   4212
      Width           =   1183
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1768
      Left            =   4212
      Stretch         =   -1  'True
      Top             =   585
      Width           =   1885
   End
   Begin VB.Label showimglink 
      DataField       =   "IMAGE"
      DataSource      =   "fetchcustomerdetailsado"
      Height          =   247
      Left            =   4212
      TabIndex        =   17
      Top             =   3627
      Visible         =   0   'False
      Width           =   1885
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      DataField       =   "GENDER"
      DataSource      =   "fetchcustomerdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1638
      TabIndex        =   16
      Top             =   3744
      Width           =   2236
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   234
      TabIndex        =   15
      Top             =   3744
      Width           =   949
   End
   Begin VB.Image Image2 
      Height          =   715
      Left            =   234
      Picture         =   "Customerprofile.frx":1084A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3406
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      DataField       =   "REG_DATE"
      DataSource      =   "fetchcustomerdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   247
      Left            =   4446
      TabIndex        =   14
      Top             =   2808
      Width           =   1417
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "A/C No. :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   234
      TabIndex        =   13
      Top             =   2340
      Width           =   1066
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   234
      TabIndex        =   12
      Top             =   5616
      Width           =   1066
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   234
      TabIndex        =   11
      Top             =   2808
      Width           =   1066
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Type:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   247
      TabIndex        =   10
      Top             =   1872
      Width           =   1287
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   247
      Left            =   4212
      TabIndex        =   9
      Top             =   234
      Width           =   1300
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   234
      TabIndex        =   8
      Top             =   936
      Width           =   1066
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   247
      TabIndex        =   7
      Top             =   1404
      Width           =   1066
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      DataField       =   "ACCOUNT_NO"
      DataSource      =   "fetchcustomerdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   286
      Left            =   1638
      TabIndex        =   6
      Top             =   2457
      Width           =   1989
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      DataField       =   "ADDRESS"
      DataSource      =   "fetchcustomerdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   481
      Left            =   1638
      TabIndex        =   5
      Top             =   5616
      Width           =   4342
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      DataField       =   "EMAIL"
      DataSource      =   "fetchcustomerdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   286
      Left            =   1638
      TabIndex        =   4
      Top             =   2808
      Width           =   3159
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      DataField       =   "ACCOUNT_TYPE"
      DataSource      =   "fetchcustomerdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1638
      TabIndex        =   3
      Top             =   1872
      Width           =   2002
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      DataField       =   "ID"
      DataSource      =   "fetchcustomerdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   247
      Left            =   5382
      TabIndex        =   2
      Top             =   234
      Width           =   598
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1638
      TabIndex        =   1
      Top             =   936
      Width           =   2002
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      DataField       =   "NAME"
      DataSource      =   "fetchcustomerdetailsado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1638
      TabIndex        =   0
      Top             =   1404
      Width           =   2002
   End
End
Attribute VB_Name = "Customerprofile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Label2.Caption = CustomerfrmLogin.customertxtUserName
fetchcustomerdetailsado.RecordSource = "select * from CUSTOMER_DETAILS where USERNAME='" + Label2.Caption + "'"
fetchcustomerdetailsado.Refresh

End Sub


Private Sub Label2_Click()
Label2.Caption = CustomerfrmLogin.customertxtUserName
End Sub




Private Sub showimglink_Change()
Image1.Picture = LoadPicture(showimglink.Caption)
End Sub

