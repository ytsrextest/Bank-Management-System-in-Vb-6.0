VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form staffrofileForm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Staff Profile"
   ClientHeight    =   4953
   ClientLeft      =   104
   ClientTop       =   416
   ClientWidth     =   6266
   Icon            =   "staffprofile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4953
   ScaleWidth      =   6266
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc fetchadmindetailsado 
      Height          =   364
      Left            =   4212
      Top             =   2808
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
      RecordSource    =   "select * from ADMIN_DETAILS"
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
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      DataField       =   "DOJ"
      DataSource      =   "fetchadmindetailsado"
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
      TabIndex        =   18
      Top             =   3861
      Width           =   2236
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "DOJ:"
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
      TabIndex        =   17
      Top             =   3861
      Width           =   949
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      DataField       =   "NAME"
      DataSource      =   "fetchadmindetailsado"
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
      Top             =   1521
      Width           =   2002
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
      TabIndex        =   15
      Top             =   1053
      Width           =   2002
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      DataField       =   "ID"
      DataSource      =   "fetchadmindetailsado"
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
      Left            =   5499
      TabIndex        =   14
      Top             =   234
      Width           =   598
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      DataField       =   "EMAIL"
      DataSource      =   "fetchadmindetailsado"
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
      TabIndex        =   13
      Top             =   1989
      Width           =   2002
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      DataField       =   "PHONENO"
      DataSource      =   "fetchadmindetailsado"
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
      TabIndex        =   12
      Top             =   2925
      Width           =   3159
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      DataField       =   "ADDRESS"
      DataSource      =   "fetchadmindetailsado"
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
      TabIndex        =   11
      Top             =   4329
      Width           =   4342
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      DataField       =   "PIN"
      DataSource      =   "fetchadmindetailsado"
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
      TabIndex        =   10
      Top             =   2457
      Width           =   1989
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
      TabIndex        =   9
      Top             =   1521
      Width           =   1066
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
      Left            =   247
      TabIndex        =   8
      Top             =   1053
      Width           =   1066
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Staff ID:"
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
      Left            =   3978
      TabIndex        =   7
      Top             =   234
      Width           =   1534
   End
   Begin VB.Label Label7 
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
      TabIndex        =   6
      Top             =   1989
      Width           =   1066
   End
   Begin VB.Label Label9 
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
      TabIndex        =   5
      Top             =   2925
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
      Left            =   247
      TabIndex        =   4
      Top             =   4329
      Width           =   1066
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Pin:"
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
      TabIndex        =   3
      Top             =   2457
      Width           =   1066
   End
   Begin VB.Image Image2 
      Height          =   715
      Left            =   234
      Picture         =   "staffprofile.frx":1084A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3406
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
      TabIndex        =   2
      Top             =   3393
      Width           =   949
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      DataField       =   "GENDER"
      DataSource      =   "fetchadmindetailsado"
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
      Top             =   3393
      Width           =   2236
   End
   Begin VB.Label showimglink 
      DataField       =   "IMAGE"
      DataSource      =   "fetchadmindetailsado"
      Height          =   247
      Left            =   4212
      TabIndex        =   0
      Top             =   3159
      Visible         =   0   'False
      Width           =   1885
   End
   Begin VB.Image Image1 
      Height          =   1768
      Left            =   4212
      Stretch         =   -1  'True
      Top             =   585
      Width           =   1885
   End
End
Attribute VB_Name = "staffrofileForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label2.Caption = admintfrmLogin.admintxtUserName
fetchadmindetailsado.RecordSource = "select * from ADMIN_DETAILS where USERNAME='" + Label2.Caption + "'"
fetchadmindetailsado.Refresh
End Sub



Private Sub showimglink_Change()
Image1.Picture = LoadPicture(showimglink.Caption)
End Sub


