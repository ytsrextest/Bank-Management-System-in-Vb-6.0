VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form howtourl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "How to use ?"
   ClientHeight    =   4394
   ClientLeft      =   2756
   ClientTop       =   3744
   ClientWidth     =   5824
   Icon            =   "howtourl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4394
   ScaleWidth      =   5824
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowserhowto 
      Height          =   4576
      Left            =   0
      TabIndex        =   0
      Top             =   -117
      Width           =   5863
      ExtentX         =   10342
      ExtentY         =   8072
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "howtourl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Me.Hide

End Sub

Private Sub Form_Load()
WebBrowserhowto.Navigate "google.com"

End Sub

