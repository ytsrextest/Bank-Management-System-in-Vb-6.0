VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form registerDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Register Now | Customer"
   ClientHeight    =   7839
   ClientLeft      =   2756
   ClientTop       =   3744
   ClientWidth     =   6123
   Icon            =   "registerDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7839
   ScaleWidth      =   6123
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox dob 
      DataField       =   "DOB"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   117
      TabIndex        =   27
      Text            =   "DD/MM/YYYY"
      Top             =   6669
      Width           =   2353
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Accept Term and Condition."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   3510
      TabIndex        =   26
      Top             =   6552
      Width           =   2353
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2691
      Top             =   3978
      _ExtentX        =   839
      _ExtentY        =   839
      _Version        =   393216
   End
   Begin VB.TextBox pwText 
      DataField       =   "PASSWORD"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      IMEMode         =   3  'DISABLE
      Left            =   3510
      PasswordChar    =   "*"
      TabIndex        =   22
      Top             =   5967
      Width           =   2470
   End
   Begin MSAdodcLib.Adodc regado 
      Height          =   299
      Left            =   2340
      Top             =   4329
      Visible         =   0   'False
      Width           =   1300
      _ExtentX        =   2396
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
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from NEW_CUSTOMER_REGISTRATION"
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
   Begin VB.CommandButton uploadCommand 
      BackColor       =   &H8000000D&
      Caption         =   "Upload Image"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   481
      Left            =   4095
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3042
      Width           =   1300
   End
   Begin VB.TextBox addressText 
      DataField       =   "ADDRESS"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   832
      Left            =   117
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   5382
      Width           =   2353
   End
   Begin VB.ComboBox gvtidCombo 
      DataField       =   "GVTID_TYPE"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   117
      TabIndex        =   18
      Text            =   "Select ID"
      Top             =   4680
      Width           =   2353
   End
   Begin VB.TextBox phoneText 
      DataField       =   "PHONENO"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   117
      TabIndex        =   16
      Top             =   3861
      Width           =   2353
   End
   Begin VB.ComboBox genderCombo 
      DataField       =   "GENDER"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   3510
      TabIndex        =   15
      Text            =   "Choose Gender"
      Top             =   5382
      Width           =   2470
   End
   Begin VB.TextBox emailText 
      DataField       =   "EMAIL"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   3510
      TabIndex        =   14
      Top             =   3978
      Width           =   2470
   End
   Begin VB.ComboBox AccuuntCombo 
      DataField       =   "ACCOUNT_TYPE"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   117
      TabIndex        =   13
      Text            =   "Select Type"
      Top             =   3159
      Width           =   2353
   End
   Begin VB.TextBox govtidnoText 
      DataField       =   "GVTID_NO"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   3510
      TabIndex        =   12
      Top             =   4680
      Width           =   2470
   End
   Begin VB.TextBox nameText 
      DataField       =   "NAME"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   117
      TabIndex        =   11
      Top             =   2457
      Width           =   2353
   End
   Begin VB.TextBox usernameText 
      DataField       =   "USERNAMEE"
      DataSource      =   "regado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   117
      TabIndex        =   9
      Top             =   1755
      Width           =   2353
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H8000000D&
      Caption         =   "SUBMIT FOR REVIEW"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   598
      Left            =   1989
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7137
      Width           =   2002
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   585
      TabIndex        =   25
      Top             =   6318
      Width           =   1417
   End
   Begin VB.Label showimglink 
      DataField       =   "IMAGE"
      DataSource      =   "regado"
      Height          =   247
      Left            =   2340
      TabIndex        =   24
      Top             =   1404
      Visible         =   0   'False
      Width           =   1183
   End
   Begin VB.Label datelabel 
      DataField       =   "REG_DATE"
      DataSource      =   "regado"
      Height          =   247
      Left            =   2340
      TabIndex        =   23
      Top             =   5031
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Left            =   4329
      TabIndex        =   21
      Top             =   5733
      Width           =   949
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1534
      Left            =   3861
      Stretch         =   -1  'True
      Top             =   1404
      Width           =   1768
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Govt ID:"
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
      Left            =   585
      TabIndex        =   17
      Top             =   4329
      Width           =   1300
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "New Customer Registration Form"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.23
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   364
      Left            =   1521
      TabIndex        =   10
      Top             =   1053
      Width           =   3406
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
      Left            =   585
      TabIndex        =   8
      Top             =   2223
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
      Left            =   585
      TabIndex        =   7
      Top             =   1404
      Width           =   1066
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Govt. ID No.:"
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
      Left            =   4329
      TabIndex        =   6
      Top             =   4446
      Width           =   1300
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
      Left            =   4329
      TabIndex        =   5
      Top             =   3627
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
      Left            =   585
      TabIndex        =   4
      Top             =   5031
      Width           =   1066
   End
   Begin VB.Label Label13 
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
      Left            =   585
      TabIndex        =   3
      Top             =   2925
      Width           =   1417
   End
   Begin VB.Label Label152 
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
      Left            =   4329
      TabIndex        =   2
      Top             =   5148
      Width           =   949
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
      Left            =   585
      TabIndex        =   1
      Top             =   3510
      Width           =   1183
   End
   Begin VB.Image Image2 
      Height          =   1183
      Left            =   0
      Picture         =   "registerDialog.frx":1084A
      Stretch         =   -1  'True
      Top             =   -117
      Width           =   6097
   End
End
Attribute VB_Name = "registerDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub datelabel_Click()
datelabel.Caption = Date
End Sub

Private Sub dob_Click()
dob.Text = Empty
End Sub

Private Sub Form_Load()
AccuuntCombo.AddItem "Saving"
AccuuntCombo.AddItem "Current"


gvtidCombo.AddItem "Adhar Card"
gvtidCombo.AddItem "Pan Card"
gvtidCombo.AddItem "Passport"

genderCombo.AddItem "Male"
genderCombo.AddItem "Female"

regado.Recordset.AddNew
datelabel.Caption = Date
dob.Text = "DD/MM/YYYY"
End Sub



Private Sub OKButton_Click()
If usernameText = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
usernameText.SetFocus
SendKeys "{Home}+{End}"

ElseIf nameText = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
nameText.SetFocus
SendKeys "{Home}+{End}"

ElseIf AccuuntCombo = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
AccuuntCombo.SetFocus
SendKeys "{Home}+{End}"

ElseIf phoneText = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
phoneText.SetFocus
SendKeys "{Home}+{End}"

ElseIf govtidnoText = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
govtidnoText.SetFocus
SendKeys "{Home}+{End}"

ElseIf addressText = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
addressText.SetFocus
SendKeys "{Home}+{End}"

ElseIf emailText = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
emailText.SetFocus
SendKeys "{Home}+{End}"

ElseIf govtidnoText = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
govtidnoText.SetFocus
SendKeys "{Home}+{End}"

ElseIf genderCombo = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
genderCombo.SetFocus
SendKeys "{Home}+{End}"

ElseIf pwText = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
pwText.SetFocus
SendKeys "{Home}+{End}"

ElseIf showimglink.Caption = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
uploadCommand.SetFocus
SendKeys "{Home}+{End}"

ElseIf dob.Text = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
dob.SetFocus
SendKeys "{Home}+{End}"

ElseIf Check1.Value = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
Check1.SetFocus
SendKeys "{Home}+{End}"

Else
regado.Recordset.AddNew
MsgBox "Successfully Submited! (Aur Team Will Send You Mail Under 24 hrs.)", vbInformation + vbOKOnly, "Your Application is Under Review."
Form2_1Mainpage.Show
Me.Hide


End If
End Sub

Private Sub showimglink_Change()
Image1.Picture = LoadPicture(showimglink.Caption)
End Sub




Private Sub uploadCommand_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "jpeg|*.jpg"
showimglink.Caption = CommonDialog1.FileName
Image1.Picture = LoadPicture(showimglink.Caption)

End Sub
