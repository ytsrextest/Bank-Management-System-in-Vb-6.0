VERSION 5.00
Begin VB.Form adminsendmailDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send Mail"
   ClientHeight    =   4043
   ClientLeft      =   2756
   ClientTop       =   3744
   ClientWidth     =   5707
   Icon            =   "adminsendmailDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4043
   ScaleWidth      =   5707
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1638
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3510
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FF8080&
      Caption         =   "Send"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   234
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3510
      Width           =   1140
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   117
      TabIndex        =   0
      Top             =   351
      Width           =   2325
   End
   Begin VB.TextBox txtMassage 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1508
      Left            =   117
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1755
      Width           =   2912
   End
   Begin VB.TextBox txtSubject 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   344
      Left            =   117
      TabIndex        =   2
      Top             =   1053
      Width           =   2327
   End
   Begin VB.TextBox txtEmail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   344
      Left            =   3159
      TabIndex        =   1
      Top             =   351
      Width           =   2324
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Email:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   273
      Index           =   1
      Left            =   3978
      TabIndex        =   9
      Top             =   117
      Width           =   1079
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   273
      Index           =   0
      Left            =   936
      TabIndex        =   8
      Top             =   117
      Width           =   1079
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   273
      Left            =   936
      TabIndex        =   7
      Top             =   819
      Width           =   1079
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Massage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   585
      TabIndex        =   6
      Top             =   1404
      Width           =   1534
   End
   Begin VB.Image Image1 
      Height          =   3991
      Left            =   2106
      Picture         =   "adminsendmailDialog.frx":1084A
      Stretch         =   -1  'True
      Top             =   234
      Width           =   3757
   End
End
Attribute VB_Name = "adminsendmailDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
Me.Hide

End Sub

Private Sub cmdOK_Click()

If txtName = Empty Then
    MsgBox "Please fill all details", vbCritical, "Massage"
        txtName.SetFocus
        SendKeys "{Home}+{End}"
        
ElseIf txtEmail = Empty Then
    MsgBox "Please fill all details", vbCritical, "Massage"
        txtEmail.SetFocus
        SendKeys "{Home}+{End}"
        
        
ElseIf txtSubject = Empty Then
    MsgBox "Please fill all details", vbCritical, "Massage"
        txtSubject.SetFocus
        SendKeys "{Home}+{End}"

        
ElseIf txtMassage = Empty Then
    MsgBox "Please fill all details", vbCritical, "Massage"
        txtMassage.SetFocus
        SendKeys "{Home}+{End}"
Else
   Dim oSmtp As New EASendMailObjLib.Mail
    oSmtp.LicenseCode = "TryIt"
    
    ' Set your Gmail email address
    oSmtp.FromAddr = "YT Bank Of India <hello@newsoflix.com>"  'Enter your Email ID here
    
    ' Add recipient email address
    oSmtp.AddRecipientEx txtEmail, 0   'Enter Reciver Email ID here
    
    ' Set email subject
    oSmtp.Subject = txtSubject
    
    ' Set email body
    oSmtp.BodyText = "Hello " & txtName & ", A very important massage for you from YT Bank Of India:  " & vbNewLine & txtMassage
    
      
    ' Gmail SMTP server address
    oSmtp.ServerAddr = "smtp.hostinger.com"
    
    ' set direct SSL 465 port,
    oSmtp.ServerPort = 465
    
    ' detect SSL/TLS automatically
    oSmtp.SSL_init

    ' Gmail user authentication should use your
    ' Gmail email address as the user name.
    ' For example: your email is "gmailid@gmail.com", then the user should be "gmailid@gmail.com"
    oSmtp.UserName = "hello@newsoflix.com" 'Enter your Email ID here again
    oSmtp.Password = "Vscode@123"    'Enter Your Mail Password
    
    'MsgBox "Wait! Start to send email ..."

    If oSmtp.SendMail() = 0 Then
        MsgBox "Email was sent successfully!", vbOKOnly + vbInformation, "Sent Successfully"
        txtName.Text = Empty
        txtEmail.Text = Empty
        txtSubject.Text = Empty
        txtMassage.Text = Empty
    Else
        MsgBox "Failed to send email with the following error: " & oSmtp.GetLastErrDescription(), vbOKOnly + vbCritical, "Opps! Sorry"
    End If
    End If
End Sub

