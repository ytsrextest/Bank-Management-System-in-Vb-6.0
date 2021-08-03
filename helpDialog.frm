VERSION 5.00
Begin VB.Form helpDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   3744
   ClientLeft      =   2756
   ClientTop       =   3744
   ClientWidth     =   6045
   FillColor       =   &H00FFFFFF&
   Icon            =   "helpDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3744
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H008080FF&
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
      Height          =   481
      Left            =   585
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2691
      Width           =   1417
   End
   Begin VB.Label reportbug 
      BackStyle       =   0  'Transparent
      Caption         =   "4. Report bug !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   481
      Left            =   117
      TabIndex        =   5
      Top             =   2106
      Width           =   3289
   End
   Begin VB.Label maagement 
      BackStyle       =   0  'Transparent
      Caption         =   "3. How admin can manage DB?"
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
      TabIndex        =   4
      Top             =   1755
      Width           =   3289
   End
   Begin VB.Label student 
      BackStyle       =   0  'Transparent
      Caption         =   "2. How new customers can login?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   481
      Left            =   117
      TabIndex        =   3
      Top             =   1404
      Width           =   3289
   End
   Begin VB.Label howto 
      BackStyle       =   0  'Transparent
      Caption         =   "1. How to use this application?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   481
      Left            =   117
      TabIndex        =   2
      Top             =   1053
      Width           =   3289
   End
   Begin VB.Shape Shape1 
      Height          =   13
      Left            =   1378
      Top             =   585
      Width           =   3289
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Help Center"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18.34
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   598
      Left            =   1085
      TabIndex        =   1
      Top             =   0
      Width           =   3874
   End
   Begin VB.Image Image1 
      Height          =   3198
      Left            =   2574
      Picture         =   "helpDialog.frx":1084A
      Stretch         =   -1  'True
      Top             =   585
      Width           =   3640
   End
End
Attribute VB_Name = "helpDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Label3_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub CancelButton_Click()
Form2_1Mainpage.Show
Me.Hide

End Sub

Private Sub howto_Click()
howtourl.Show

End Sub

Private Sub maagement_Click()
MsgBox "Bank Staff have to login & There is several option for Database management, Also Admin visit developer tab For any Help.", vbOKOnly + vbInformation, "Help Center"
End Sub

Private Sub reportbug_Click()
MsgBox "Please visit (www.ytsrex.com/ytsbank) to report bug.", vbOKOnly + vbInformation, "Help Center"
End Sub

Private Sub student_Click()
MsgBox "* GoTo Customer Login > Click On Register Now > Fill Form & Submit Form." & vbCrLf & vbCrLf & "*After Submiton We Review And Approve it Under 24 Hrs.", vbOKOnly + vbInformation, "Help Center"

End Sub
