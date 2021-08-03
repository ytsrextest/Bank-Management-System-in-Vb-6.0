VERSION 5.00
Begin VB.Form DBbackupDialog 
   BackColor       =   &H00400040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Backup "
   ClientHeight    =   3718
   ClientLeft      =   2756
   ClientTop       =   3744
   ClientWidth     =   4212
   Icon            =   "DBbackupDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3718
   ScaleWidth      =   4212
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   364
      Left            =   234
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2106
      Visible         =   0   'False
      Width           =   481
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Clean Now"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   598
      Left            =   1268
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   702
      Width           =   1534
   End
   Begin VB.CommandButton TakeBackup 
      BackColor       =   &H008080FF&
      Caption         =   "Take Backup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   598
      Left            =   1287
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2106
      Width           =   1534
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Step2: Click On Take Backup, And Wait For Success Massage. ""It Will Take Some Time"""
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   481
      Left            =   98
      TabIndex        =   4
      Top             =   1521
      Width           =   3991
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: Backup is Saved on ""app\oracle\oradata"" Path Where You Installed Oracle ."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   481
      Left            =   7
      TabIndex        =   3
      Top             =   3042
      Width           =   4199
   End
   Begin VB.Label filenamelabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Step1: Clear Space, Previous Backup. After Success  Massage Step 2 Will Be Unlock"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.19
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   598
      Left            =   234
      TabIndex        =   2
      Top             =   117
      Width           =   3731
   End
End
Attribute VB_Name = "DBbackupDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function GetNamedPipeInfo Lib "kernel32" (ByVal hNamedPipe As Long, lType As Long, lLenOutBuf As Long, lLenInBuf As Long, lMaxInstances As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As Any, lpProcessInformation As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


'Purpose     :  Synchronously runs a DOS command line and returns the captured screen output.
'Inputs      :  sCommandLine                The DOS command line to run.
'               [bShowWindow]               If True displays the DOS output window.
'Outputs     :  Returns the screen output
'Notes       :  This routine will work only with those program that send their output to
'               the standard output device (stdout).
'               Windows NT ONLY.
'Revisions   :

Function ShellExecuteCapture(sCommandLine As String, Optional bShowWindow As Boolean = False) As String
    Const clReadBytes As Long = 256, INFINITE As Long = &HFFFFFFFF
    Const STARTF_USESHOWWINDOW = &H1, STARTF_USESTDHANDLES = &H100&
    Const SW_HIDE = 0, SW_NORMAL = 1
    Const NORMAL_PRIORITY_CLASS = &H20&

    Const PIPE_CLIENT_END = &H0     'The handle refers to the client end of a named pipe instance. This is the default.
    Const PIPE_SERVER_END = &H1     'The handle refers to the server end of a named pipe instance. If this value is not specified, the handle refers to the client end of a named pipe instance.
    Const PIPE_TYPE_BYTE = &H0      'The named pipe is a byte pipe. This is the default.
    Const PIPE_TYPE_MESSAGE = &H4   'The named pipe is a message pipe. If this value is not specified, the pipe is a byte pipe


    Dim tProcInfo As PROCESS_INFORMATION, lRetVal As Long, lSuccess As Long
    Dim tStartupInf As STARTUPINFO
    Dim tSecurAttrib As SECURITY_ATTRIBUTES, lhwndReadPipe As Long, lhwndWritePipe As Long
    Dim lBytesRead As Long, sBuffer As String
    Dim lPipeOutLen As Long, lPipeInLen As Long, lMaxInst As Long

    tSecurAttrib.nLength = Len(tSecurAttrib)
    tSecurAttrib.bInheritHandle = 1&
    tSecurAttrib.lpSecurityDescriptor = 0&

    lRetVal = CreatePipe(lhwndReadPipe, lhwndWritePipe, tSecurAttrib, 0)
    If lRetVal = 0 Then
        'CreatePipe failed
        Exit Function
    End If

    tStartupInf.cb = Len(tStartupInf)
    tStartupInf.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    tStartupInf.hStdOutput = lhwndWritePipe
    If bShowWindow Then
        'Show the DOS window
        tStartupInf.wShowWindow = SW_NORMAL
    Else
        'Hide the DOS window
        tStartupInf.wShowWindow = SW_HIDE
    End If

    lRetVal = CreateProcessA(0&, sCommandLine, tSecurAttrib, tSecurAttrib, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, tStartupInf, tProcInfo)
    If lRetVal <> 1 Then
        'CreateProcess failed
        Exit Function
    End If

    'Process created, wait for completion. Note, this will cause your application
    'to hang indefinately until this process completes.
    WaitForSingleObject tProcInfo.hProcess, INFINITE

    'Determine pipes contents
    lSuccess = GetNamedPipeInfo(lhwndReadPipe, PIPE_TYPE_BYTE, lPipeOutLen, lPipeInLen, lMaxInst)
    If lSuccess Then
        'Got pipe info, create buffer
        sBuffer = String(lPipeOutLen, 0)
        'Read Output Pipe
        lSuccess = ReadFile(lhwndReadPipe, sBuffer, lPipeOutLen, lBytesRead, 0&)
        If lSuccess = 1 Then
            'Pipe read successfully
            ShellExecuteCapture = Left$(sBuffer, lBytesRead)
        End If
    End If

    'Close handles
    Call CloseHandle(tProcInfo.hProcess)
    Call CloseHandle(tProcInfo.hThread)
    Call CloseHandle(lhwndReadPipe)
    Call CloseHandle(lhwndWritePipe)
End Function

Sub Clear()
    'Debug.Print ShellExecuteCapture("C:\Users\Ytsrex\Desktop\Bank Management System\Admin\grantpermissionBAT\deleteFile.bat", False)
    Text1.Text = ShellExecuteCapture("C:\Users\Ytsrex\Desktop\Bank Management System\Admin\grantpermissionBAT\deleteFile.bat", False)
End Sub
Sub Backup()
    'Debug.Print ShellExecuteCapture("C:\Users\Ytsrex\Desktop\Bank Management System\Admin\grantpermissionBAT\ExportFile.bat", False)
    Text1.Text = ShellExecuteCapture("C:\Users\Ytsrex\Desktop\Bank Management System\Admin\grantpermissionBAT\ExportFile.bat", False)
End Sub

Private Sub Command1_Click()
    Call Clear
    MsgBox "Space Clear Successfully!, Now Click On Take Backup.", vbOKOnly + vbInformation, "Success ! "
    TakeBackup.Visible = True
    
End Sub

Private Sub Form_Load()
TakeBackup.Visible = False

End Sub

Private Sub TakeBackup_Click()
Call Backup
 MsgBox "Backup Taken Successfully!", vbOKOnly + vbInformation, "Success ! "
 Unload Me
End Sub


