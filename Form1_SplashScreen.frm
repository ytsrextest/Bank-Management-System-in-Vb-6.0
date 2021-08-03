VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1_SplashScreen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Welcome to YT Bank | Please Wait..."
   ClientHeight    =   4147
   ClientLeft      =   39
   ClientTop       =   299
   ClientWidth     =   6721
   Icon            =   "Form1_SplashScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4147
   ScaleWidth      =   6721
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   364
      Left            =   1482
      TabIndex        =   3
      Top             =   3042
      Width           =   3757
      _ExtentX        =   6925
      _ExtentY        =   671
      _Version        =   327682
      Appearance      =   1
      Max             =   105
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   5265
      Top             =   1404
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "We Understand your World"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   247
      Left            =   3861
      TabIndex        =   0
      Top             =   702
      Width           =   2587
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   247
      Left            =   2808
      TabIndex        =   1
      Top             =   3510
      Width           =   949
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   247
      Left            =   3627
      TabIndex        =   5
      Top             =   2691
      Width           =   832
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   364
      Left            =   2691
      TabIndex        =   4
      Top             =   2691
      Width           =   1066
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "© all rights reserved by alok kumar (ytsrex media)"
      BeginProperty Font 
         Name            =   "Technic"
         Size            =   8.83
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   1053
      TabIndex        =   2
      Top             =   3861
      Width           =   4810
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H008080FF&
      Height          =   1651
      Left            =   5265
      Shape           =   3  'Circle
      Top             =   2925
      Width           =   2236
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FF80&
      FillColor       =   &H0080FF80&
      Height          =   1651
      Left            =   -585
      Shape           =   3  'Circle
      Top             =   2925
      Width           =   2236
   End
   Begin VB.Image Image5 
      Height          =   1183
      Left            =   0
      Picture         =   "Form1_SplashScreen.frx":1084A
      Stretch         =   -1  'True
      Top             =   -234
      Width           =   6682
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0080FFFF&
      BorderColor     =   &H0080FFFF&
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   1651
      Left            =   234
      Shape           =   2  'Oval
      Top             =   3627
      Width           =   1768
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00C0FFFF&
      BorderColor     =   &H0080FFFF&
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   1651
      Left            =   5148
      Shape           =   2  'Oval
      Top             =   3510
      Width           =   1651
   End
   Begin VB.Image Image2 
      Height          =   2236
      Left            =   2223
      Picture         =   "Form1_SplashScreen.frx":23312
      Stretch         =   -1  'True
      Top             =   702
      Width           =   2353
   End
End
Attribute VB_Name = "Form1_SplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Timer1.Enabled = True
Label3.Visible = False
Label5.Visible = False

End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
Label3.Visible = True
Label5.Visible = True
Label5.Caption = ProgressBar1.Value & "%"


If (ProgressBar1.Value = ProgressBar1.Max) Then
Timer1.Enabled = False
Unload Me
Form2_1Mainpage.Show

End If



End Sub
