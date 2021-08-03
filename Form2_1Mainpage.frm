VERSION 5.00
Begin VB.Form Form2_1Mainpage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "YT BANK OF INDIA | MANAGEMENT SYSTEM VERSION 1.0"
   ClientHeight    =   6045
   ClientLeft      =   39
   ClientTop       =   299
   ClientWidth     =   7995
   Icon            =   "Form2_1Mainpage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8841.94
   ScaleMode       =   0  'User
   ScaleWidth      =   10188.75
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image9 
      Height          =   884
      Left            =   234
      Picture         =   "Form2_1Mainpage.frx":1084A
      Top             =   3861
      Width           =   1872
   End
   Begin VB.Image Image8 
      Height          =   884
      Left            =   5850
      Picture         =   "Form2_1Mainpage.frx":1139E
      Top             =   3861
      Width           =   1755
   End
   Begin VB.Image Image7 
      Height          =   949
      Left            =   2340
      Picture         =   "Form2_1Mainpage.frx":11E89
      Stretch         =   -1  'True
      Top             =   3861
      Width           =   3172
   End
   Begin VB.Image Image6 
      Height          =   949
      Left            =   2340
      Picture         =   "Form2_1Mainpage.frx":12EDC
      Stretch         =   -1  'True
      Top             =   2808
      Width           =   3172
   End
   Begin VB.Image Image5 
      Height          =   949
      Left            =   2340
      Picture         =   "Form2_1Mainpage.frx":143AC
      Stretch         =   -1  'True
      Top             =   1755
      Width           =   3172
   End
   Begin VB.Image Image4 
      Height          =   2938
      Left            =   5616
      Picture         =   "Form2_1Mainpage.frx":157E4
      Stretch         =   -1  'True
      Top             =   1170
      Width           =   2236
   End
   Begin VB.Image Image3 
      Height          =   2938
      Left            =   117
      Picture         =   "Form2_1Mainpage.frx":3B7CDB
      Stretch         =   -1  'True
      Top             =   1170
      Width           =   2236
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "© all rights reserved by alok kumar (ytsrex media)"
      BeginProperty Font 
         Name            =   "Technic"
         Size            =   10.87
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   364
      Left            =   1053
      TabIndex        =   1
      Top             =   5733
      Width           =   5863
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN AREA"
      DragIcon        =   "Form2_1Mainpage.frx":75A1D2
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   14.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   364
      Left            =   2873
      TabIndex        =   0
      Top             =   1287
      Width           =   2470
   End
   Begin VB.Image Image2 
      Height          =   559
      Left            =   2639
      Picture         =   "Form2_1Mainpage.frx":76AA1C
      Stretch         =   -1  'True
      Top             =   1170
      Width           =   546
   End
   Begin VB.Image Image1 
      Height          =   1534
      Left            =   0
      Picture         =   "Form2_1Mainpage.frx":772D72
      Stretch         =   -1  'True
      Top             =   -234
      Width           =   7969
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFC0&
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   2353
      Left            =   -936
      Shape           =   2  'Oval
      Top             =   4914
      Width           =   2704
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF8080&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   2353
      Left            =   6318
      Shape           =   2  'Oval
      Top             =   4797
      Width           =   2704
   End
   Begin VB.Shape Shape3 
      DrawMode        =   8  'Xor Pen
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   1248
      Left            =   1053
      Top             =   5616
      Width           =   5863
   End
End
Attribute VB_Name = "Form2_1Mainpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image5_Click()
CustomerfrmLogin.Show
Form2_1Mainpage.Hide

End Sub

Private Sub Image6_Click()
admintfrmLogin.Show
Form2_1Mainpage.Hide
End Sub

Private Sub Image7_Click()
contactfrm.Show

End Sub

Private Sub Image8_Click()
End
End Sub

Private Sub Image9_Click()
helpDialog.Show
Me.Hide

End Sub
