VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Forgetpwadmin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Forget Password- Bank Staff"
   ClientHeight    =   3198
   ClientLeft      =   2756
   ClientTop       =   3744
   ClientWidth     =   6032
   Icon            =   "Forgetpwadmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3198
   ScaleWidth      =   6032
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc forgetadminpwado 
      Height          =   364
      Left            =   3510
      Top             =   2691
      Visible         =   0   'False
      Width           =   1885
      _ExtentX        =   3475
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
   Begin VB.TextBox Forgetadminemail 
      Height          =   364
      Left            =   234
      TabIndex        =   0
      Top             =   351
      Width           =   2587
   End
   Begin VB.TextBox forgetadminpincode 
      Height          =   364
      Left            =   234
      TabIndex        =   1
      Top             =   1170
      Width           =   2587
   End
   Begin VB.TextBox adminnewpw 
      Height          =   364
      Left            =   234
      TabIndex        =   2
      Top             =   1989
      Width           =   2587
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1755
      TabIndex        =   4
      Top             =   2691
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   117
      TabIndex        =   3
      Top             =   2691
      Width           =   1215
   End
   Begin VB.Label studentForgetpw 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Email:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1053
      TabIndex        =   7
      Top             =   117
      Width           =   1417
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Pin Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   1053
      TabIndex        =   6
      Top             =   819
      Width           =   1417
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   247
      Left            =   819
      TabIndex        =   5
      Top             =   1638
      Width           =   1651
   End
   Begin VB.Image Image1 
      Height          =   3172
      Left            =   3042
      Picture         =   "Forgetpwadmin.frx":1084A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3055
   End
End
Attribute VB_Name = "Forgetpwadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
admintfrmLogin.Show
Me.Hide

End Sub



Private Sub forgetadminpincode_Change()

End Sub

Private Sub OKButton_Click()
forgetadminpwado.RecordSource = "select * from ADMIN_DETAILS where EMAIL='" + Forgetadminemail.Text + "' and PIN='" + forgetadminpincode.Text + "'"
forgetadminpwado.Refresh

If forgetadminpwado.Recordset.EOF Then
    
    MsgBox "Invalid Details, try again!", vbCritical, "Oops! Wrong Detail"
    forgetadminpincode.SetFocus
    SendKeys "{Home}+{End}"
    Else
    
       forgetadminpwado.Recordset.Fields("PASSWORD") = adminnewpw.Text
       forgetadminpwado.Recordset.Update
       Me.Hide
       MsgBox "Password Changed Successfully.", vbInformation, "Password Changed"
       admintfrmLogin.Show
    
       
       
    End If
End Sub
