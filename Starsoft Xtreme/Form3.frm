VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3960
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   3960
   StartUpPosition =   1  'CenterOwner
   Begin StarsoftXtremeOS.Windowtwo Windowtwo1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5318
      Begin VB.CommandButton cmdReboot 
         Caption         =   "Start Action"
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optShut 
         Caption         =   "Change OS"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Value           =   -1  'True
         Width           =   1500
      End
      Begin VB.OptionButton optShut 
         Caption         =   "Reboot"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton optShut 
         Caption         =   "Shut Down"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   2
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optShut 
         Caption         =   "Force"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   1
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Shutdown Options"
         Height          =   2175
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   3495
         Begin VB.CommandButton Exit 
            Caption         =   "Exit OS"
            Height          =   375
            Left            =   1680
            TabIndex        =   9
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Log Off"
            Height          =   375
            Left            =   1680
            TabIndex        =   8
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Shutdown Options"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   0
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bWindowsNT As Boolean
'
' Operating System Constants, Types and Declares
'
Const VER_PLATFORM_WIN32s = 0
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'
' Shutdown/change resolution Constants, Types and Declares
'
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
'
' Using this option to shutdown windows does not send
' the WM_QUERYENDSESSION and WM_ENDSESSION messages to
' the open applications. Thus, those apps may loose
' any unsaved data.
'
Const EWX_FORCE = 4
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_BITSPERPEL = &H40000
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const CDS_UPDATEREGISTRY = &H1
Const CDS_TEST = &H2
Const CDS_FULLSCREEN = &H4
Const DISP_CHANGE_SUCCESSFUL = 0
Const DISP_CHANGE_RESTART = 1
 
Const HWND_BROADCAST = &HFFFF&
Const WM_DISPLAYCHANGE = &H7E&
Const SPI_SETNONCLIENTMETRICS = 42

Private Type DEVMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'
' The following are required to shutdown NT.
'
Const ERROR_NOT_ALL_ASSIGNED = 1300
Const SE_PRIVILEGE_ENABLED = 2
Const TOKEN_QUERY = &H8
Const TOKEN_ADJUST_PRIVILEGES = &H20

Private Type LUID
    lowpart As Long
    highpart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges As LUID_AND_ATTRIBUTES
End Type

Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpUid As LUID) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Sub Command1_Click()
frmlogin.Show
Me.Hide
frmMain.Hide
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
frmlogin.Show
Me.Hide
frmMain.Hide
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub optClose_Click()
frmMain.Show
End Sub

Private Sub Timer1_Timer()


End Sub

Private Sub cmdReboot_Click()
Dim tLuid          As LUID
Dim tTokenPriv     As TOKEN_PRIVILEGES
Dim tPrevTokenPriv As TOKEN_PRIVILEGES
Dim lResult        As Long
Dim lToken         As Long
Dim lLenBuffer     As Long
Dim lMode As Long
'
' Determine the shutdown mode.
'
' EWX_LOGOFF
'   Shuts down all processes running and
'   logs off the user.
'
' EWX_REBOOT
'   Shuts down and restarts the system.
'
' EWX_SHUTDOWN
'   Shuts down the system to a point where
'   it is safe to turn off the system.
'
' EWX_POWEROFF
'   Shuts down the system and turns off power.
'   The system must support this feature.
'
' EWX_FORCE
'   Forcibly shuts down the system. Files are not closed,...
'   data may be lost.
'
If optShut(0) Then
    lMode = EWX_LOGOFF
ElseIf optShut(1) Then
    lMode = EWX_REBOOT
ElseIf optShut(2) Then
    lMode = EWX_SHUTDOWN
Else: lMode = EWX_FORCE
End If

If Not bWindowsNT Then
    Call ExitWindowsEx(lMode, 0)
Else
    '
    ' Get the access token of the current process.  Get it
    ' with the privileges of querying the access token and
    ' adjusting its privileges.
    '
    lResult = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, lToken)
    If lResult = 0 Then
        Exit Sub 'Failed
    End If
    '
    ' Get the locally unique identifier (LUID) which
    ' represents the shutdown privilege.
    '
    lResult = LookupPrivilegeValue(0&, "SeShutdownPrivilege", tLuid)
    If lResult = 0 Then Exit Sub 'Failed
    '
    ' Populate the new TOKEN_PRIVILEGES values with the LUID
    ' and allow your current process to shutdown the computer.
    '
    With tTokenPriv
        .PrivilegeCount = 1
        .Privileges.Attributes = SE_PRIVILEGE_ENABLED
        .Privileges.pLuid = tLuid
    lResult = AdjustTokenPrivileges(lToken, False, tTokenPriv, Len(tPrevTokenPriv), tPrevTokenPriv, lLenBuffer)
    End With
    
    If lResult = 0 Then
        Exit Sub 'Failed
    Else
        If Err.LastDllError = ERROR_NOT_ALL_ASSIGNED Then Exit Sub 'Failed
    End If
    '
    '  Shutdown Windows.
    '
    Call ExitWindowsEx(lMode, 0)
End If
End Sub

Private Sub Exit_Click()
End
 
End Sub

Private Sub Form_Load()


Dim OSInfo As OSVERSIONINFO
'
' See if we are running Windows 9x or NT.
'
OSInfo.dwOSVersionInfoSize = Len(OSInfo)
Call GetVersionEx(OSInfo)
bWindowsNT = (OSInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)

End Sub

Private Sub Form_Resize()
Windowtwo1.Top = 0
Windowtwo1.Left = 0
Windowtwo1.Height = Me.ScaleHeight
Windowtwo1.Width = Me.ScaleWidth
End Sub

Private Sub y_Click(Index As Integer)
frmlogin.Show
Me.Hide
frmMain.Hide
End Sub

Private Sub log_Click()

End Sub
