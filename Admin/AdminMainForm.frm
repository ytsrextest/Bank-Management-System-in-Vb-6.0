VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form AdminMainForm 
   Caption         =   "Staff Dashboard | YT Bank Of India"
   ClientHeight    =   6669
   ClientLeft      =   195
   ClientTop       =   793
   ClientWidth     =   11596
   Icon            =   "AdminMainForm.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "AdminMainForm.frx":1084A
   ScaleHeight     =   6669
   ScaleWidth      =   11596
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid Withdraw_DATAGRID 
      Bindings        =   "AdminMainForm.frx":24110
      Height          =   4342
      Left            =   936
      TabIndex        =   20
      Top             =   2223
      Width           =   15691
      _ExtentX        =   28922
      _ExtentY        =   8003
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483642
      ForeColor       =   16776960
      HeadLines       =   2
      RowHeight       =   15
      RowDividerStyle =   1
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "WITHDRAW TRANSACTION DATABASE"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid Deposite_DATAGRID 
      Bindings        =   "AdminMainForm.frx":2412D
      Height          =   4342
      Left            =   936
      TabIndex        =   19
      Top             =   2223
      Width           =   15691
      _ExtentX        =   28922
      _ExtentY        =   8003
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483642
      ForeColor       =   16776960
      HeadLines       =   2
      RowHeight       =   15
      RowDividerStyle =   1
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "DEPOSITE TRANSACTION DATABASE"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid NewCustomerreg_DataGrid 
      Bindings        =   "AdminMainForm.frx":2414A
      Height          =   4342
      Left            =   936
      TabIndex        =   18
      Top             =   2223
      Width           =   15691
      _ExtentX        =   28922
      _ExtentY        =   8003
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483642
      ForeColor       =   16776960
      HeadLines       =   2
      RowHeight       =   15
      RowDividerStyle =   1
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NEW CUSTOMER REGISTRATION DATABASE"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H008080FF&
      Caption         =   "Select Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1534
      Left            =   351
      TabIndex        =   10
      Top             =   351
      Width           =   7969
      Begin MSAdodcLib.Adodc customerdetailsfetchado 
         Height          =   299
         Left            =   3159
         Top             =   234
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
         RecordSource    =   "select * from CUSTOMER_DETAILS"
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
      Begin MSAdodcLib.Adodc Admindatafetchado 
         Height          =   299
         Left            =   1404
         Top             =   234
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
         RecordSource    =   "select * from ADMIN_DETAILS"
         Caption         =   "Admin Data"
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
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Admin Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   832
         Left            =   234
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   351
         Width           =   2353
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080FFFF&
         Caption         =   "Customer Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   832
         Left            =   2808
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   351
         Width           =   2353
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H0080FFFF&
         Caption         =   "New Customers Registration Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   832
         Left            =   5382
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   351
         Width           =   2353
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Search By Name/ Username"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   247
         Left            =   5616
         TabIndex        =   23
         Top             =   1287
         Width           =   2119
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Search By Account Number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   247
         Left            =   3042
         TabIndex        =   22
         Top             =   1287
         Width           =   2119
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Search By PIN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   247
         Left            =   819
         TabIndex        =   21
         Top             =   1287
         Width           =   1066
      End
   End
   Begin MSDataGridLib.DataGrid CustomerDatabase_DataGrid 
      Bindings        =   "AdminMainForm.frx":2416A
      Height          =   4342
      Left            =   936
      TabIndex        =   17
      Top             =   2223
      Width           =   15691
      _ExtentX        =   28922
      _ExtentY        =   8003
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483642
      ForeColor       =   16776960
      HeadLines       =   2
      RowHeight       =   15
      RowDividerStyle =   1
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CUSTOMER DATABASE"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid AdminDatabase_DATAGRID 
      Bindings        =   "AdminMainForm.frx":24190
      Height          =   4342
      Left            =   936
      TabIndex        =   16
      Top             =   2223
      Width           =   15691
      _ExtentX        =   28922
      _ExtentY        =   8003
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483642
      ForeColor       =   16776960
      HeadLines       =   2
      RowHeight       =   15
      RowDividerStyle =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ADMIN DATABASE"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton closemanagementdb 
      BackColor       =   &H8000000D&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   598
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6903
      Width           =   1885
   End
   Begin VB.CommandButton resetbtn 
      BackColor       =   &H8000000D&
      Caption         =   "Reset All"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   598
      Left            =   7488
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6903
      Width           =   1651
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Transaction Detals"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1534
      Left            =   8658
      TabIndex        =   7
      Top             =   351
      Width           =   5746
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Deposite Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   832
         Left            =   585
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   351
         Width           =   1885
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H008080FF&
         Caption         =   "Withdraw Project"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   832
         Left            =   3159
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   351
         Width           =   1885
      End
      Begin MSAdodcLib.Adodc Deposite_Dbado 
         Height          =   299
         Left            =   936
         Top             =   117
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
         RecordSource    =   "select * from DEPOSITE"
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
      Begin MSAdodcLib.Adodc Withdraw_Dbado 
         Height          =   299
         Left            =   3744
         Top             =   117
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
         RecordSource    =   "select * from WITHDRAW"
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
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Search By Transaction ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   247
         Left            =   3276
         TabIndex        =   26
         Top             =   1287
         Width           =   1768
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Search By Transaction ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   247
         Left            =   702
         TabIndex        =   24
         Top             =   1287
         Width           =   1768
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H008080FF&
      Caption         =   "Search Box"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1534
      Left            =   14742
      TabIndex        =   4
      Top             =   351
      Width           =   2353
      Begin VB.TextBox searchtxt 
         Height          =   364
         Left            =   117
         TabIndex        =   6
         Top             =   351
         Width           =   2119
      End
      Begin VB.CommandButton seaechcmd 
         BackColor       =   &H0080FFFF&
         Caption         =   "Search"
         Height          =   364
         Left            =   702
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   819
         Width           =   1066
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3276
      Top             =   6084
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   949
      Left            =   351
      TabIndex        =   0
      Top             =   5850
      Visible         =   0   'False
      Width           =   3640
      Begin MSAdodcLib.Adodc fetchadmindetailsado 
         Height          =   364
         Left            =   1053
         Top             =   234
         Width           =   1092
         _ExtentX        =   2013
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
      Begin VB.Label Label1 
         Caption         =   "Label1"
         DataField       =   "NAME"
         DataSource      =   "fetchadmindetailsado"
         Height          =   247
         Left            =   2457
         TabIndex        =   2
         Top             =   234
         Width           =   949
      End
      Begin VB.Label Username 
         Caption         =   "Username"
         Height          =   247
         Left            =   117
         TabIndex        =   1
         Top             =   234
         Width           =   832
      End
   End
   Begin MSAdodcLib.Adodc NewRegCustomerado 
      Height          =   299
      Left            =   6318
      Top             =   1755
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Search By Transaction ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   247
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   1768
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0FFFF&
      Height          =   2236
      Left            =   0
      Top             =   0
      Width           =   17797
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12.23
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   1300
      Left            =   5616
      TabIndex        =   3
      Top             =   5499
      Width           =   6799
   End
   Begin VB.Menu profile 
      Caption         =   "Profile"
      Begin VB.Menu viewprofile 
         Caption         =   "View Profile"
      End
      Begin VB.Menu Changeprofiledetails 
         Caption         =   "Change Profile Details"
      End
   End
   Begin VB.Menu BackupDatabase 
      Caption         =   "Backup Database"
   End
   Begin VB.Menu sendMail 
      Caption         =   "Send Mail"
   End
   Begin VB.Menu contactdeveloper 
      Caption         =   "Contact Developer"
   End
   Begin VB.Menu logout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "AdminMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BackupDatabase_Click()
DBbackupDialog.Show

End Sub

Private Sub changeprofiledetails_Click()
changestaffdetailsDialog.Show

End Sub

Private Sub closemanagementdb_Click()
AdminDatabase_DATAGRID.Visible = False
CustomerDatabase_DataGrid.Visible = False
NewCustomerreg_DataGrid.Visible = False
Deposite_DATAGRID.Visible = False
Withdraw_DATAGRID.Visible = False

closemanagementdb.Visible = False
End Sub

Private Sub Command1_Click()
Deposite_DATAGRID.Visible = True
closemanagementdb.Visible = True
AdminDatabase_DATAGRID.Visible = False
CustomerDatabase_DataGrid.Visible = False
NewCustomerreg_DataGrid.Visible = False
Withdraw_DATAGRID.Visible = False
End Sub

Private Sub Command2_Click()
Withdraw_DATAGRID.Visible = True
closemanagementdb.Visible = True
AdminDatabase_DATAGRID.Visible = False
CustomerDatabase_DataGrid.Visible = False
NewCustomerreg_DataGrid.Visible = False
Deposite_DATAGRID.Visible = False
End Sub

Private Sub Command4_Click()
AdminDatabase_DATAGRID.Visible = True
closemanagementdb.Visible = True
CustomerDatabase_DataGrid.Visible = False
NewCustomerreg_DataGrid.Visible = False
Deposite_DATAGRID.Visible = False
Withdraw_DATAGRID.Visible = False
End Sub

Private Sub Command5_Click()
CustomerDatabase_DataGrid.Visible = True
closemanagementdb.Visible = True
AdminDatabase_DATAGRID.Visible = False
NewCustomerreg_DataGrid.Visible = False
Deposite_DATAGRID.Visible = False
Withdraw_DATAGRID.Visible = False
End Sub

Private Sub Command6_Click()
NewCustomerreg_DataGrid.Visible = True
closemanagementdb.Visible = True
AdminDatabase_DATAGRID.Visible = False
CustomerDatabase_DataGrid.Visible = False
Deposite_DATAGRID.Visible = False
Withdraw_DATAGRID.Visible = False
End Sub

Private Sub contactdeveloper_Click()
developercontactDialog.Show
End Sub

Private Sub Form_Load()
Username.Caption = admintfrmLogin.admintxtUserName
fetchadmindetailsado.RecordSource = "select * from ADMIN_DETAILS where USERNAME='" + Username.Caption + "'"
fetchadmindetailsado.Refresh

'BTN Visible
resetbtn.Visible = False
closemanagementdb.Visible = False
'Datagrid Visibal
AdminDatabase_DATAGRID.Visible = False
CustomerDatabase_DataGrid.Visible = False
NewCustomerreg_DataGrid.Visible = False
Deposite_DATAGRID.Visible = False
Withdraw_DATAGRID.Visible = False
End Sub

Private Sub Label2_Change()
Label2.Caption = "Welcome " + Label1 + " in Staff Dashboard" & vbNewLine & " Date: " & Date & ", Time: " & Time
End Sub


Private Sub logout_Click()
End
End Sub

Private Sub resetbtn_Click()
'------RESET SEARCH-----
'Admin DB
If AdminDatabase_DATAGRID.Visible = True Then

Admindatafetchado.RecordSource = " select * from ADMIN_DETAILS "
Admindatafetchado.Refresh
Admindatafetchado.Caption = Admindatafetchado.RecordSource
resetbtn.Visible = False

'CUSTOMER DB
ElseIf CustomerDatabase_DataGrid.Visible = True Then
customerdetailsfetchado.RecordSource = " select * from CUSTOMER_DETAILS "
customerdetailsfetchado.Refresh
customerdetailsfetchado.Caption = customerdetailsfetchado.RecordSource
resetbtn.Visible = False


'NEW CUSTOMER REGISTRATION
ElseIf NewCustomerreg_DataGrid.Visible = True Then
NewRegCustomerado.RecordSource = " select * from NEW_CUSTOMER_REGISTRATION"
NewRegCustomerado.Refresh
NewRegCustomerado.Caption = NewRegCustomerado.RecordSource
resetbtn.Visible = False


'DEPOSITE DB
ElseIf Deposite_DATAGRID.Visible = True Then
Deposite_Dbado.RecordSource = " select * from DEPOSITE "
Deposite_Dbado.Refresh
Deposite_Dbado.Caption = Deposite_Dbado.RecordSource
resetbtn.Visible = False

'WITHDRAW DB
ElseIf Withdraw_DATAGRID.Visible = True Then
Withdraw_Dbado.RecordSource = " select * from WITHDRAW"
Withdraw_Dbado.Refresh
Withdraw_Dbado.Caption = Withdraw_Dbado.RecordSource
resetbtn.Visible = False

'Main Else
Else
MsgBox "Please Select anything then Seaech", vbCritical, "Search Record!"
searchtxt.Text = Empty

End If

'-----------------------
End Sub

Private Sub seaechcmd_Click()
'------SEARCH FROM DB------

'ADMINDB
If AdminDatabase_DATAGRID.Visible = True Then

Admindatafetchado.RecordSource = " select * from ADMIN_DETAILS where PIN='" + searchtxt.Text + "'"
Admindatafetchado.Refresh

If Admindatafetchado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Search Record!"
resetbtn.Visible = True
searchtxt.Text = Empty


Else
Admindatafetchado.Caption = Admindatafetchado.RecordSource
searchtxt.Text = Empty
resetbtn.Visible = True
End If

'CUSTOMER DB
ElseIf CustomerDatabase_DataGrid.Visible = True Then
customerdetailsfetchado.RecordSource = " select * from CUSTOMER_DETAILS where ACCOUNT_NO='" + searchtxt.Text + "'"
customerdetailsfetchado.Refresh

If customerdetailsfetchado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Search Record!"
resetbtn.Visible = True
searchtxt.Text = Empty
Else
customerdetailsfetchado.Caption = customerdetailsfetchado.RecordSource
searchtxt.Text = Empty
resetbtn.Visible = True
End If

'NEW REGISTER CUSTOMER
ElseIf NewCustomerreg_DataGrid.Visible = True Then
NewRegCustomerado.RecordSource = " select * from NEW_CUSTOMER_REGISTRATION where NAME='" + searchtxt.Text + "' or USERNAMEE='" + searchtxt.Text + "'"
NewRegCustomerado.Refresh

If NewRegCustomerado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Search Record!"
resetbtn.Visible = True
searchtxt.Text = Empty
Else
NewRegCustomerado.Caption = NewRegCustomerado.RecordSource
searchtxt.Text = Empty
resetbtn.Visible = True
End If

'DEPOSITE
ElseIf Deposite_DATAGRID.Visible = True Then
Deposite_Dbado.RecordSource = " select * from DEPOSITE where TRANSACTION_ID='" + searchtxt.Text + "'"
Deposite_Dbado.Refresh

If Deposite_Dbado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Search Record!"
resetbtn.Visible = True
searchtxt.Text = Empty
Else
Deposite_Dbado.Caption = Deposite_Dbado.RecordSource
searchtxt.Text = Empty
resetbtn.Visible = True
End If


'WITHDRAW
ElseIf Withdraw_DATAGRID.Visible = True Then
Withdraw_Dbado.RecordSource = " select * from WITHDRAW where TRANSACTION_ID='" + searchtxt.Text + "'"
Withdraw_Dbado.Refresh

If Withdraw_Dbado.Recordset.EOF Then
MsgBox "Data Not Found", vbCritical, "Search Record!"
resetbtn.Visible = True
searchtxt.Text = Empty
Else
Withdraw_Dbado.Caption = Withdraw_Dbado.RecordSource
searchtxt.Text = Empty
resetbtn.Visible = True
End If


'Main Else
Else
MsgBox "Please Select anything then Seaech", vbCritical, "Search Record!"
searchtxt.Text = Empty
End If
'--------------------------
End Sub

Private Sub sendmail_Click()
adminsendmailDialog.Show

End Sub

Private Sub Timer1_Timer()
Label2.Caption = Date & Time
End Sub

Private Sub viewprofile_Click()
staffrofileForm.Show

End Sub

