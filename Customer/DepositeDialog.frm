VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form DepositeDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Deposite Balance | Yt Bank Of India"
   ClientHeight    =   3419
   ClientLeft      =   2756
   ClientTop       =   3744
   ClientWidth     =   4901
   Icon            =   "DepositeDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3419
   ScaleWidth      =   4901
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc udatecustomerado 
      Height          =   299
      Left            =   3393
      Top             =   2808
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
      UserName        =   "bankadmin"
      Password        =   "9122335311"
      RecordSource    =   "select *From CUSTOMER_DETAILS"
      Caption         =   "udatecustomer"
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
   Begin VB.TextBox deosittxt 
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
      TabIndex        =   0
      Top             =   1638
      Width           =   2119
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3406
      Left            =   2925
      TabIndex        =   3
      Top             =   234
      Visible         =   0   'False
      Width           =   1885
      Begin MSAdodcLib.Adodc deositefillado 
         Height          =   299
         Left            =   0
         Top             =   1404
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
         UserName        =   "bankadmin"
         Password        =   "9122335311"
         RecordSource    =   "select * from DEPOSITE"
         Caption         =   "Fill Diosite"
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
      Begin MSAdodcLib.Adodc balanceado 
         Height          =   299
         Left            =   117
         Top             =   819
         Width           =   1417
         _ExtentX        =   2612
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
         UserName        =   "bankadmin"
         Password        =   "9122335311"
         RecordSource    =   "select * from CUSTOMER_DETAILS"
         Caption         =   "before deposite"
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
      Begin VB.Label LabelTime 
         Caption         =   "Time"
         DataField       =   "DPOSITE_TIME"
         DataSource      =   "deositefillado"
         Height          =   247
         Left            =   1053
         TabIndex        =   16
         Top             =   3042
         Width           =   598
      End
      Begin VB.Label newaccountno 
         Caption         =   "new account no"
         DataField       =   "ACCOUNT_NO"
         DataSource      =   "deositefillado"
         Height          =   247
         Left            =   234
         TabIndex        =   14
         Top             =   3159
         Width           =   1417
      End
      Begin VB.Label accountno 
         Caption         =   "Account no"
         DataField       =   "ACCOUNT_NO"
         DataSource      =   "balanceado"
         Height          =   364
         Left            =   234
         TabIndex        =   13
         Top             =   2925
         Width           =   1066
      End
      Begin VB.Label macid 
         Caption         =   "Mac Id"
         DataField       =   "MAC_ID"
         DataSource      =   "deositefillado"
         Height          =   247
         Left            =   117
         TabIndex        =   12
         Top             =   2691
         Width           =   1417
      End
      Begin VB.Label transactionid 
         Caption         =   "Transaction ID"
         DataField       =   "TRANSACTION_ID"
         DataSource      =   "deositefillado"
         Height          =   247
         Left            =   234
         TabIndex        =   11
         Top             =   2457
         Width           =   1183
      End
      Begin VB.Label TotalBalancecustomer 
         Caption         =   "New total balance customer"
         DataField       =   "ACCOUNT_BALLANCE"
         DataSource      =   "udatecustomerado"
         Height          =   364
         Left            =   117
         TabIndex        =   10
         Top             =   468
         Width           =   1651
      End
      Begin VB.Label fillbalancebeforedeosite 
         Caption         =   "fill balance before deosite"
         DataField       =   "BALLANCE_BEFORE_DEPOSITE"
         DataSource      =   "deositefillado"
         Height          =   247
         Left            =   117
         TabIndex        =   9
         Top             =   1989
         Width           =   1768
      End
      Begin VB.Label Username 
         Caption         =   "Username"
         Height          =   247
         Left            =   0
         TabIndex        =   7
         Top             =   1170
         Width           =   949
      End
      Begin VB.Label newtotalballance 
         Caption         =   "New Total Ballance"
         DataField       =   "NEW_TOTAL_BALLANCE"
         DataSource      =   "deositefillado"
         Height          =   247
         Left            =   117
         TabIndex        =   6
         Top             =   1755
         Width           =   1534
      End
      Begin VB.Label depositeballance 
         Caption         =   "Deposite Ballance"
         DataField       =   "DEPOSITE_BALLANCE"
         DataSource      =   "deositefillado"
         Height          =   247
         Left            =   117
         TabIndex        =   5
         Top             =   2223
         Width           =   1534
      End
      Begin VB.Label balancebeforedeposite 
         Caption         =   "Balance Before Deposite"
         DataField       =   "ACCOUNT_BALLANCE"
         DataSource      =   "balanceado"
         Height          =   247
         Left            =   117
         TabIndex        =   4
         Top             =   234
         Width           =   1534
      End
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Deposite Now"
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
      Left            =   468
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2106
      Width           =   1417
   End
   Begin VB.Label Storetransid 
      Caption         =   "Store Transaction ID"
      Height          =   247
      Left            =   351
      TabIndex        =   17
      Top             =   3042
      Visible         =   0   'False
      Width           =   1651
   End
   Begin VB.Label datecap 
      DataField       =   "DEPOSITE_DATE"
      DataSource      =   "deositefillado"
      Height          =   247
      Left            =   3744
      TabIndex        =   15
      Top             =   702
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Amount:"
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
      Top             =   1287
      Width           =   1417
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit Money"
      DataField       =   "ACCOUNT_NO"
      DataSource      =   "deositefillado"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   364
      Left            =   351
      TabIndex        =   1
      Top             =   702
      Width           =   1768
   End
   Begin VB.Image Image1 
      Height          =   832
      Left            =   0
      Picture         =   "DepositeDialog.frx":1084A
      Stretch         =   -1  'True
      Top             =   -117
      Width           =   4810
   End
   Begin VB.Image Image2 
      Height          =   3172
      Left            =   1638
      Picture         =   "DepositeDialog.frx":23312
      Stretch         =   -1  'True
      Top             =   351
      Width           =   3406
   End
End
Attribute VB_Name = "DepositeDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'------------ MAC ID---------------
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetAdaptersInfo Lib "iphlpapi" (lpAdapterInfo As Any, lpSize As Long) As Long
 
Public Function GetMacAddress() As String
    Const OFFSET_LENGTH As Long = 400
    Dim lSize           As Long
    Dim baBuffer()      As Byte
    Dim lIdx            As Long
    Dim sRetVal         As String
    
    Call GetAdaptersInfo(ByVal 0, lSize)
    If lSize <> 0 Then
        ReDim baBuffer(0 To lSize - 1) As Byte
        Call GetAdaptersInfo(baBuffer(0), lSize)
        Call CopyMemory(lSize, baBuffer(OFFSET_LENGTH), 4)
        For lIdx = OFFSET_LENGTH + 4 To OFFSET_LENGTH + 4 + lSize - 1
            sRetVal = IIf(LenB(sRetVal) <> 0, sRetVal & ":", vbNullString) & Right$("0" & Hex$(baBuffer(lIdx)), 2)
        Next
    End If
    GetMacAddress = sRetVal
End Function
'------------ END MAC ID ---------------



Private Sub deosittxt_Change()
depositeballance.Caption = deosittxt


'Fill Balance Before Deosite
fillbalancebeforedeosite.Caption = balancebeforedeposite

' Total New Balance
Dim a, b, X As Integer
    a = Val(balancebeforedeposite.Caption)
    b = Val(depositeballance.Caption)
    X = a + b
newtotalballance.Caption = X
TotalBalancecustomer.Caption = newtotalballance
End Sub
Public Function RandomString(Length As Integer) As String
    Dim i As Integer
    Do While i < Length
        Randomize
        Select Case IIf(i = 0, Int(1 * Rnd + 1), Int(2 * Rnd))
            Case 0: RandomString = RandomString & Chr(Int(10 * Rnd + 48))
            Case 1: RandomString = RandomString & Chr(Int(26 * Rnd + 65))
        End Select
        i = i + 1
        transactionid.Caption = RandomString
        Storetransid.Caption = transactionid
    Loop
    
End Function



Private Sub Form_Load()
deositefillado.Recordset.AddNew
udatecustomerado.Recordset.AddNew
    Username.Caption = CustomerfrmLogin.customertxtUserName
    balanceado.RecordSource = "select * from CUSTOMER_DETAILS where USERNAME='" + Username.Caption + "'"
    balanceado.Refresh
    
    udatecustomerado.RecordSource = "select * from CUSTOMER_DETAILS where USERNAME='" + Username.Caption + "'"
    udatecustomerado.Refresh
    
    
macid.Caption = GetMacAddress
newaccountno.Caption = accountno
datecap.Caption = Date
LabelTime.Caption = Time


' Random Transaction ID
    Debug.Print RandomString(9)
    Debug.Print RandomString(9)
End Sub

Private Sub OKButton_Click()

Dim b As Integer
    b = Val(depositeballance.Caption)
    
If deosittxt = Empty Then
MsgBox "Please fill all details", vbCritical, "Massage"
    deosittxt.SetFocus
    SendKeys "{Home}+{End}"
    
ElseIf b < 1 Then
MsgBox "Please Deposite More than ""1"" Rs.", vbCritical, "Massage"
    deosittxt.SetFocus
    SendKeys "{Home}+{End}"

Else

udatecustomerado.Recordset.AddNew
deositefillado.Recordset.AddNew
MsgBox "Successfully Deposited!" & vbCrLf & "Transaction ID: " & Storetransid, vbOKOnly + vbInformation, "Congratulations"

'AGAIN CALL
deosittxt.Text = Empty
macid.Caption = GetMacAddress
newaccountno.Caption = accountno
datecap.Caption = Date
LabelTime.Caption = Time
'transactionid.Caption = RandomString
CustomermainForm.Show
Unload Me

End If

End Sub
