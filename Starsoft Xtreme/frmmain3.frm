VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmdis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Display Settings"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   Icon            =   "frmmain3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10365
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Wallpaper &Screensaver"
      TabPicture(0)   =   "frmmain3.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgscreen"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imgpic"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblpreview"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "udInt"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdextview"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkcyclic"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmddeselectall"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdselectall"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txttimer"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "File2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "framode"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdfull"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdexit"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "comPattern"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkpic"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "File1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Dir1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Drive1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Command4"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Screen Resolution"
      TabPicture(1)   =   "frmmain3.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image1"
      Tab(1).Control(1)=   "Image2"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(3)=   "optRes(2)"
      Tab(1).Control(4)=   "optRes(1)"
      Tab(1).Control(5)=   "optRes(0)"
      Tab(1).Control(6)=   "cmdChange"
      Tab(1).Control(7)=   "Command1"
      Tab(1).Control(8)=   "Command2"
      Tab(1).Control(9)=   "Command3"
      Tab(1).ControlCount=   10
      Begin VB.CommandButton Command4 
         Caption         =   "Set As Wallpaper"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   28
         Top             =   6000
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Preview"
         Height          =   375
         Left            =   -74640
         TabIndex        =   26
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Preview"
         Height          =   375
         Left            =   -74640
         TabIndex        =   25
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Preview"
         Height          =   375
         Left            =   -74640
         TabIndex        =   24
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "&Change Resolution"
         Height          =   375
         Left            =   -74280
         TabIndex        =   23
         Top             =   4440
         Width           =   2175
      End
      Begin VB.OptionButton optRes 
         Caption         =   "640 x 480"
         Height          =   195
         Index           =   0
         Left            =   -73080
         TabIndex        =   22
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton optRes 
         Caption         =   "800 x 600"
         Height          =   195
         Index           =   1
         Left            =   -73080
         TabIndex        =   21
         Top             =   2400
         Width           =   1335
      End
      Begin VB.OptionButton optRes 
         Caption         =   "1024 x 768"
         Height          =   195
         Index           =   2
         Left            =   -73080
         TabIndex        =   20
         Top             =   3240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   0
         TabIndex        =   17
         Top             =   480
         Width           =   3135
      End
      Begin VB.DirListBox Dir1 
         Height          =   3915
         Left            =   0
         TabIndex        =   16
         Top             =   960
         Width           =   3135
      End
      Begin VB.FileListBox File1 
         Height          =   3990
         Left            =   3120
         Pattern         =   "*.jpg"
         TabIndex        =   15
         Top             =   960
         Width           =   3135
      End
      Begin VB.CheckBox chkpic 
         Caption         =   " Show Wallpaper Preview"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         TabIndex        =   14
         Top             =   4080
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.ComboBox comPattern 
         Height          =   315
         ItemData        =   "frmmain3.frx":0044
         Left            =   3120
         List            =   "frmmain3.frx":0054
         TabIndex        =   13
         Text            =   "*.jpg"
         Top             =   480
         Width           =   3135
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   12
         Top             =   6000
         Width           =   1335
      End
      Begin VB.CommandButton cmdfull 
         Caption         =   "&Screensaver"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   11
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Frame framode 
         Caption         =   "Mode"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   7320
         TabIndex        =   8
         Top             =   4560
         Width           =   1815
         Begin VB.OptionButton optsingle 
            Caption         =   "Wallpaper"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optmulti 
            Caption         =   "Screensaver"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.FileListBox File2 
         Height          =   3990
         Left            =   3120
         MultiSelect     =   1  'Simple
         Pattern         =   "*.jpg"
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox txttimer 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Text            =   "1000"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton cmdselectall 
         Caption         =   "Select &All"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   6000
         Width           =   1575
      End
      Begin VB.CommandButton cmddeselectall 
         Caption         =   "&DeSelect All"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   6000
         Width           =   1575
      End
      Begin VB.CheckBox chkcyclic 
         Caption         =   "&Cyclic Screensaver"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   3
         Top             =   5280
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdextview 
         Caption         =   "&External Viewer"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   6000
         Width           =   1575
      End
      Begin MSComCtl2.UpDown udInt 
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   5400
         Width           =   195
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Value           =   1000
         Increment       =   100
         Max             =   10000
         Min             =   100
         Enabled         =   0   'False
      End
      Begin VB.Frame Frame1 
         Caption         =   "Resolutions"
         Height          =   2895
         Left            =   -74760
         TabIndex        =   27
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Image Image2 
         Height          =   2775
         Left            =   -70080
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Image Image1 
         Height          =   4860
         Left            =   -71040
         Picture         =   "frmmain3.frx":0074
         Stretch         =   -1  'True
         Top             =   600
         Width           =   5655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interval (milli Seonds)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   19
         Top             =   5040
         Width           =   2205
      End
      Begin VB.Label lblpreview 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Preview"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1485
         Left            =   7560
         TabIndex        =   18
         Top             =   1320
         Width           =   1665
      End
      Begin VB.Image imgpic 
         BorderStyle     =   1  'Fixed Single
         Height          =   2655
         Left            =   6720
         Stretch         =   -1  'True
         Top             =   720
         Width           =   3255
      End
      Begin VB.Image imgscreen 
         Height          =   3900
         Left            =   6120
         Picture         =   "frmmain3.frx":06CB
         Stretch         =   -1  'True
         Top             =   360
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmdis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim s As Integer
Dim path As String
Dim i As Integer
Dim ti As Long

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



Private Sub chkpic_Click()
    On Error Resume Next
    If chkpic.Value = Unchecked Then
        imgpic.Picture = LoadPicture()
        lblpreview.Visible = True
    ElseIf chkpic.Value = Checked Then
        If File1.ListIndex > 0 Then
            imgpic.Picture = LoadPicture(path)
            lblpreview.Visible = False
        Else
            lblpreview.Visible = True
        End If
    End If
End Sub

Private Sub cmdChange_Click()
Dim DevM    As DEVMODE
Dim lResult As Long
Dim iAns    As Integer
'
' Retrieve info about the current graphics mode
' on the current display device.
'
lResult = EnumDisplaySettings(0, 0, DevM)
'
' Set the new resolution. Don't change the color
' depth so a restart is not necessary.
'
With DevM
    .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Or DM_BITSPERPEL
    If optRes(0) Then
        .dmPelsWidth = 640  'ScreenWidth
        .dmPelsHeight = 480 'ScreenHeight
    ElseIf optRes(1) Then
        .dmPelsWidth = 800
        .dmPelsHeight = 600
    Else
        .dmPelsWidth = 1024
        .dmPelsHeight = 768
    End If
    '.dmBitsPerPel = 32 (could be 8, 16, 32 or even 4)
End With
'
' Change the display settings to the specified graphics mode.
'
lResult = ChangeDisplaySettings(DevM, CDS_FULLSCREEN)
Select Case lResult
    Case DISP_CHANGE_RESTART
        iAns = MsgBox("You must restart your computer to apply these changes." & _
            vbCrLf & vbCrLf & "Do you want to restart now?", _
            vbYesNo + vbSystemModal, "Screen Resolution")
        If iAns = vbYes Then Call ExitWindowsEx(EWX_REBOOT, 0)
    Case DISP_CHANGE_SUCCESSFUL
        Call ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
        Call SendMessage(HWND_BROADCAST, WM_DISPLAYCHANGE, SPI_SETNONCLIENTMETRICS, ByVal 0&)
        MsgBox "Screen resolution changed", vbInformation, "Resolution Changed"
    Case Else
        MsgBox "Mode not supported", vbSystemModal, "Error"
End Select


End Sub

Private Sub cmddeselectall_Click()
  On Error Resume Next
  For i = 0 To File2.ListCount - 1
        File2.Selected(i) = False
  Next
  cmdselectall.Enabled = True
  cmddeselectall.Enabled = False
End Sub

Private Sub cmdexit_Click()
    On Error Resume Next
    Unload Me
    Unload frmfullscreen
    Unload frmslideshow
    Unload frmextview
    End
End Sub

Private Sub cmdextview_Click()
   On Error Resume Next
   If optsingle.Value = True Then
   If File1.ListIndex >= 0 Then
        images(0) = path
        Load frmextview
    If frmdis.optsingle.Value = True Then
        frmextview.cmdnext.Visible = False
        frmextview.cmdprevious.Visible = False
        frmextview.cmdslide.Visible = False
        frmextview.cmdviewline2.Visible = False
        tis = 1
    ElseIf frmdis.optmulti.Value = True Then
        frmextview.cmdnext.Visible = True
        frmextview.cmdprevious.Visible = True
        frmextview.cmdslide.Visible = True
        frmextview.cmdviewline2.Visible = True
    End If
        frmextview.Show 1
   Else
        s = MsgBox("No Image Selected For External Viewer", vbInformation, "External Viewer Error")
   End If
   ElseIf optmulti.Value = True Then
       For i = 0 To 9999
            images(i) = " "
        Next
        If Val(txttimer.Text) < 100 Then
            txttimer.Text = "1000"
            s = MsgBox("Minimum Value is 100 Millseconds", vbCritical + vbOKOnly, "Interval Underflow Error")
            Exit Sub
        ElseIf Val(txttimer.Text) > 10000 Then
            txttimer.Text = "1000"
            s = MsgBox("Maximum Value is 10000 Millseconds", vbCritical + vbOKOnly, "Interval Overflow Error")
            Exit Sub
        End If
        tis = 0
            For i = 0 To File2.ListCount - 1
                If File2.Selected(i) = True Then
                    images(tis) = Dir1.path & "\" & File2.List(i)
                    tis = tis + 1
                End If
            Next
        If tis > 0 Then
            Load frmextview
            frmextview.Show 1
        End If
   End If
End Sub

Private Sub cmdfull_Click()
    On Error Resume Next
    If cmdfull.Caption = "&Full Screen" Then
        If File1.ListIndex >= 0 Then
            frmfullscreen.Image1.Picture = LoadPicture(path)
            frmfullscreen.Show 1
        Else
            s = MsgBox("No Image Selected For Full Screen View", vbInformation, "Full Screen View Error")
        End If
    ElseIf cmdfull.Caption = "&Screensaver" Then
        For i = 0 To 9999
            images(i) = " "
        Next
        If File2.ListIndex <> -1 Then
        If IsNumeric(txttimer.Text) Then
            If Val(txttimer.Text) < 100 Then
                s = MsgBox("Minimum Value is 100 Millseconds", vbCritical + vbOKOnly, "Interval Underflow Error")
                txttimer.Text = "1000"
                Exit Sub
            ElseIf Val(txttimer.Text) > 10000 Then
                s = MsgBox("Maximum Value is 10000 Millseconds", vbCritical + vbOKOnly, "Interval Overflow Error")
                txttimer.Text = "1000"
            Exit Sub
            End If
        tis = 0
            For i = 0 To File2.ListCount - 1
                If File2.Selected(i) = True Then
                    images(tis) = Dir1.path & "\" & File2.List(i)
                    tis = tis + 1
                End If
            Next
        If tis >= 0 Then
            frmslideshow.Show 1
        Else
            s = MsgBox("No Image Selected For Screen Saver", vbInformation, "External Viewer Error")
        End If
        Else
            s = MsgBox("Interval Cannot Be Non-Numeric", vbCritical + vbOKOnly, "Timer Value Error")
            txttimer.Text = "1000"
        End If
        Else
            s = MsgBox("No Image Selected For Screen Saver", vbInformation, "Full Screen View Error")
        End If
    End If
End Sub

Private Sub cmdselectall_Click()
  On Error Resume Next
  For i = 0 To File2.ListCount - 1
        File2.Selected(i) = True
  Next
  cmddeselectall.Enabled = True
  cmdselectall.Enabled = False
End Sub

Private Sub Combo1_Change()
  On Error Resume Next
    File1.Pattern = comPattern.Text
    File2.Pattern = comPattern.Text
End Sub

Private Sub Command1_Click()
Image2.Picture = LoadPicture("min.jpg")
End Sub

Private Sub Command2_Click()
Image2.Picture = LoadPicture("med.jpg")

End Sub

Private Sub Command3_Click()
Image2.Picture = LoadPicture("Max.jpg")

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command4_Click()
If File1.ListIndex >= 0 Then
            frmMain.Picture = LoadPicture(path)
        Else
            s = MsgBox("No Image Selected ", vbInformation, "Wallpaper")
        End If
End Sub

Private Sub comPattern_Change()
    On Error Resume Next
    File1.Pattern = comPattern.Text
    File2.Pattern = comPattern.Text
End Sub

Private Sub comPattern_Click()
    On Error Resume Next
    File1.Pattern = comPattern.Text
    File2.Pattern = comPattern.Text
End Sub

Private Sub Dir1_Change()
    On Error Resume Next
    File1.path = Dir1.path
    File2.path = Dir1.path
End Sub

Private Sub Drive1_Change()
    On Error GoTo errorhandel
    Dir1.path = Drive1.Drive
errorhandel:
    If Err.Number = 68 Then
        s = MsgBox("Drive " & UCase$(Drive1.Drive) & " Not Ready", vbCritical + vbOKOnly, " Drive Select Error")
        Drive1.Drive = "C:\"
    End If
End Sub

Private Sub Drive2_Change()
 
End Sub

Private Sub File1_Click()
    On Error Resume Next
    If chkpic.Value = Checked Then
        path = Dir1.path & "\" & File1.List(File1.ListIndex)
        imgpic.Picture = LoadPicture(path)
        lblpreview.Visible = False
    ElseIf chkpic.Value = Unchecked Then
        path = Dir1.path & "\" & File1.List(File1.ListIndex)
    End If
End Sub

Private Sub Form_Load()
    Drive1.Drive = "C:\"
    cmdfull.Caption = "&Full Screen"
End Sub

Private Sub Label2_Click()
End Sub

Private Sub Label3_Click()
End Sub

Private Sub Label4_Click()

End Sub

Private Sub optmulti_Click()
    On Error Resume Next
    cmdfull.Caption = "&Screensaver"
    For i = 0 To File2.ListCount - 1
        File2.Selected(i) = False
    Next
    chkcyclic.Enabled = True
    cmdselectall.Enabled = True
    cmddeselectall.Enabled = False
    File1.Visible = False
    File2.Visible = True
    txttimer.Enabled = True
    udInt.Enabled = True
    Label1.Enabled = True
    lblpreview.Caption = " Preview Not Available In Screensaver Mode"
    chkpic.Enabled = False
    imgpic.Picture = LoadPicture()
    lblpreview.Visible = True
End Sub

Private Sub optsingle_Click()
    On Error Resume Next
    cmdfull.Caption = "&Full Screen"
    chkcyclic.Enabled = False
    cmdselectall.Enabled = False
    cmddeselectall.Enabled = False
    File1.Visible = True
    File2.Visible = False
    lblpreview.Caption = "Preview"
    txttimer.Enabled = False
    udInt.Enabled = False
    Label1.Enabled = False
    chkpic.Enabled = True
    If File1.ListIndex >= 0 Then
        lblpreview.Visible = False
        imgpic.Picture = LoadPicture(path)
    End If
End Sub

Private Sub udInt_Change()
    On Error Resume Next
    txttimer.Text = udInt.Value
End Sub
