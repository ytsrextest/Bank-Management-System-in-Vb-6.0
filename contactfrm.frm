VERSION 5.00
Begin VB.Form contactfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contact Us | YT Bank Of India"
   ClientHeight    =   4108
   ClientLeft      =   2834
   ClientTop       =   3484
   ClientWidth     =   7644
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000B&
   Icon            =   "contactfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2424.004
   ScaleMode       =   0  'User
   ScaleWidth      =   7163.929
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1287
      TabIndex        =   0
      Top             =   117
      Width           =   2325
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
      Left            =   5148
      TabIndex        =   2
      Top             =   117
      Width           =   2324
   End
   Begin VB.ComboBox acctype 
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
      Left            =   5148
      TabIndex        =   3
      Text            =   "Select Account Type"
      Top             =   702
      Width           =   2327
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
      Left            =   1287
      TabIndex        =   12
      Top             =   1170
      Width           =   2327
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
      Height          =   1391
      Left            =   117
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1989
      Width           =   3731
   End
   Begin VB.TextBox txtphoneno 
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
      Left            =   1287
      TabIndex        =   1
      Top             =   702
      Width           =   2324
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
      Left            =   702
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3510
      Width           =   1140
   End
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
      Left            =   2106
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3510
      Width           =   1140
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
      Left            =   1170
      TabIndex        =   13
      Top             =   1638
      Width           =   1534
   End
   Begin VB.Image Image1 
      Height          =   3289
      Left            =   3978
      Picture         =   "contactfrm.frx":1084A
      Stretch         =   -1  'True
      Top             =   1053
      Width           =   3718
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
      Left            =   117
      TabIndex        =   11
      Top             =   1170
      Width           =   1079
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Acc. Type"
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
      Left            =   4095
      TabIndex        =   10
      Top             =   702
      Width           =   1079
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No. :"
      Height          =   273
      Left            =   117
      TabIndex        =   9
      Top             =   702
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
      Left            =   117
      TabIndex        =   5
      Top             =   117
      Width           =   1079
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
      Left            =   4095
      TabIndex        =   6
      Top             =   195
      Width           =   1079
   End
End
Attribute VB_Name = "contactfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSmtp As New EASendMailObjLib.Mail
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
        
ElseIf txtphoneno = Empty Then
    MsgBox "Please fill all details", vbCritical, "Massage"
        txtphoneno.SetFocus
        SendKeys "{Home}+{End}"
        
ElseIf txtSubject = Empty Then
    MsgBox "Please fill all details", vbCritical, "Massage"
        txtSubject.SetFocus
        SendKeys "{Home}+{End}"

        
ElseIf acctype = Empty Then
    MsgBox "Please fill all details", vbCritical, "Massage"
        acctype.SetFocus
        SendKeys "{Home}+{End}"
        
ElseIf txtMassage = Empty Then
    MsgBox "Please fill all details", vbCritical, "Massage"
        txtMassage.SetFocus
        SendKeys "{Home}+{End}"
Else
   Dim oSmtp As New EASendMailObjLib.Mail
    oSmtp.LicenseCode = "TryIt"
    
    ' Set your Gmail email address
    oSmtp.FromAddr = "Contact- YTS Bank od India <hello@newsoflix.com>"  'Enter your Email ID here
    
    ' Add recipient email address
    oSmtp.AddRecipientEx "ytsrex2@gmail.com", 0   'Enter Reciver Email ID here
    
    ' Set email subject
    oSmtp.Subject = txtSubject
    
    ' Set email body
    oSmtp.BodyText = "Name- " & txtName & vbNewLine & "Customer Email- " & txtEmail & vbNewLine & "Phone No- " & txtphoneno & vbNewLine & "Account Type- " & acctype & vbNewLine & "Massage- " & vbNewLine & txtMassage
    
      
    ' Gmail SMTP server address
    oSmtp.ServerAddr = "smtp.hostinger.com"
    
    ' set direct SSL 465 port,
    oSmtp.ServerPort = 465
    
    ' detect SSL/TLS automatically
    oSmtp.SSL_init

    ' Gmail user authentication should use your
    ' Gmail email address as the user name.
    ' For example: your email is "gmailid@gmail.com", then the user should be "gmailid@gmail.com"
    oSmtp.Username = "hello@newsoflix.com" 'Enter your Email ID here again
    oSmtp.Password = "Vscode@123"    'Enter Your Mail Password
    
    'MsgBox "Wait! Start to send email ..."

    If oSmtp.SendMail() = 0 Then
        MsgBox "Email was sent successfully!", vbOKOnly + vbInformation, "Sent Successfully"
        txtName.Text = Empty
        txtEmail.Text = Empty
        txtSubject.Text = Empty
        txtphoneno.Text = Empty
        acctype.Text = Empty
        txtMassage.Text = Empty
    Else
        MsgBox "Failed to send email with the following error: " & oSmtp.GetLastErrDescription(), vbOKOnly + vbCritical, "Opps! Sorry"
    End If
    End If
End Sub



Private Sub Form_Load()
acctype.AddItem "Current"
acctype.AddItem "Saving"

End Sub

Private Sub Text1_Change()

End Sub

