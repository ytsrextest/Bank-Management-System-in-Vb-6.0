VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form changestaffdetailsDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Staff Profile"
   ClientHeight    =   4043
   ClientLeft      =   2756
   ClientTop       =   3744
   ClientWidth     =   6175
   Icon            =   "changestaffdetailsDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4043
   ScaleWidth      =   6175
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      BackColor       =   &H8000000D&
      Caption         =   "UPDATE DETAILS"
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
      Left            =   2223
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3276
      Width           =   1534
   End
   Begin VB.TextBox UpdateemailText 
      DataField       =   "EMAIL"
      DataSource      =   "updatestaffdetailsado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   1989
      TabIndex        =   5
      Top             =   1521
      Width           =   2119
   End
   Begin VB.TextBox UpdatePinText 
      DataField       =   "PIN"
      DataSource      =   "updatestaffdetailsado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   1989
      TabIndex        =   4
      Top             =   2106
      Width           =   2119
   End
   Begin VB.TextBox UpdatephoneText 
      DataField       =   "PHONENO"
      DataSource      =   "updatestaffdetailsado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   1989
      TabIndex        =   3
      Top             =   2691
      Width           =   2119
   End
   Begin VB.CommandButton ImageCommand 
      BackColor       =   &H008080FF&
      Caption         =   "Change Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   481
      Left            =   4563
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3276
      Width           =   1300
   End
   Begin MSAdodcLib.Adodc updatestaffdetailsado 
      Height          =   364
      Left            =   234
      Top             =   3159
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3861
      Top             =   3393
      _ExtentX        =   839
      _ExtentY        =   839
      _Version        =   393216
   End
   Begin VB.Label showimglink 
      DataField       =   "IMAGE"
      DataSource      =   "updatestaffdetailsado"
      Height          =   247
      Left            =   117
      TabIndex        =   10
      Top             =   3510
      Visible         =   0   'False
      Width           =   1885
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Pin No.:"
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
      Left            =   117
      TabIndex        =   9
      Top             =   2223
      Width           =   1768
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Phone No.:"
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
      Left            =   117
      TabIndex        =   8
      Top             =   2808
      Width           =   1755
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Email:"
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
      Left            =   117
      TabIndex        =   7
      Top             =   1638
      Width           =   1768
   End
   Begin VB.Image Image1 
      Height          =   1768
      Left            =   4212
      Stretch         =   -1  'True
      Top             =   1404
      Width           =   1885
   End
   Begin VB.Label Label2 
      Height          =   364
      Left            =   117
      TabIndex        =   1
      Top             =   1053
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE STAFF DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.87
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   364
      Left            =   1521
      TabIndex        =   0
      Top             =   1053
      Width           =   3523
   End
   Begin VB.Image Image2 
      Height          =   1300
      Left            =   0
      Picture         =   "changestaffdetailsDialog.frx":1084A
      Stretch         =   -1  'True
      Top             =   -117
      Width           =   6097
   End
End
Attribute VB_Name = "changestaffdetailsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Label2.Caption = admintfrmLogin.admintxtUserName
updatestaffdetailsado.RecordSource = "select * from ADMIN_DETAILS where USERNAME='" + Label2.Caption + "'"
updatestaffdetailsado.Refresh
End Sub

Private Sub ImageCommand_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "jpeg|*.jpg"
showimglink.Caption = CommonDialog1.FileName
Image1.Picture = LoadPicture(showimglink.Caption)
MsgBox "Image Selected, Now please click on Update Details button to save!", vbOKOnly + vbInformation, "Success"
End Sub

Private Sub Label2_Click()
Label2.Caption = admintfrmLogin.admintxtUserName
End Sub

Private Sub OKButton_Click()
updatestaffdetailsado.Recordset.Fields("EMAIL") = UpdateemailText.Text
updatestaffdetailsado.Recordset.Fields("PIN") = UpdatePinText.Text
updatestaffdetailsado.Recordset.Fields("PHONENO") = UpdatephoneText.Text
updatestaffdetailsado.Recordset.Fields("IMAGE") = showimglink.Caption
updatestaffdetailsado.Recordset.Update
MsgBox "Successfully Updated!", vbOKOnly + vbInformation, "Success"
Unload Me
End Sub

Private Sub showimglink_Change()
Image1.Picture = LoadPicture(showimglink.Caption)
End Sub


