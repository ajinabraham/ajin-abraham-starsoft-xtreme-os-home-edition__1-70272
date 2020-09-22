VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "starsoft Xtreme ********"
   ClientHeight    =   8940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   Picture         =   "frmMain.frx":2055D
   ScaleHeight     =   8940
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11400
      Top             =   6000
   End
   Begin VB.CommandButton cmoodContn 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      Caption         =   "Display Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   10920
      MaskColor       =   &H0000C000&
      Picture         =   "frmMain.frx":26C8B
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2640
      Width           =   975
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   8925
      Left            =   6720
      TabIndex        =   7
      Top             =   -100
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton Command1yuu 
         Caption         =   "File Explorer"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Picture         =   "frmMain.frx":3F1AD
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   6720
         Width           =   2175
      End
      Begin VB.CommandButton Calculator 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "Calculator"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         MaskColor       =   &H0080C0FF&
         Picture         =   "frmMain.frx":401ED
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CommandButton Command1f 
         Caption         =   " IExplorer"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Picture         =   "frmMain.frx":4122D
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3360
         Width           =   2175
      End
      Begin VB.CommandButton MP3 
         Caption         =   "MP3 Player"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Picture         =   "frmMain.frx":4226D
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4200
         Width           =   2175
      End
      Begin VB.CommandButton Pain 
         Caption         =   "Glaze Paint"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Picture         =   "frmMain.frx":432AD
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CommandButton web 
         Caption         =   "WEB Editor"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Picture         =   "frmMain.frx":442ED
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5040
         Width           =   2175
      End
      Begin VB.CommandButton Commandff2 
         Caption         =   "Application Sheduler"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Picture         =   "frmMain.frx":4532D
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5880
         Width           =   2175
      End
      Begin VB.Label Labeiiil1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "All Applications"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.Image hh 
         Height          =   510
         Index           =   1
         Left            =   480
         Picture         =   "frmMain.frx":4636D
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1995
      End
      Begin VB.Image Imagde4 
         Height          =   390
         Index           =   1
         Left            =   3360
         Picture         =   "frmMain.frx":47C1F
         Top             =   360
         Width           =   435
      End
      Begin VB.Image Imarrge4 
         Height          =   795
         Left            =   3000
         Picture         =   "frmMain.frx":48551
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   795
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000005&
         BorderWidth     =   12
         X1              =   2520
         X2              =   2520
         Y1              =   7560
         Y2              =   8040
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         BorderWidth     =   8
         Index           =   0
         X1              =   360
         X2              =   360
         Y1              =   7560
         Y2              =   8040
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000E&
         BorderWidth     =   5
         X1              =   360
         X2              =   2520
         Y1              =   8040
         Y2              =   8040
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         BorderWidth     =   6
         X1              =   360
         X2              =   2520
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "All Application"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   19
         Top             =   7680
         Width           =   1245
      End
      Begin VB.Image hh 
         Height          =   510
         Index           =   0
         Left            =   360
         Picture         =   "frmMain.frx":87D97
         Stretch         =   -1  'True
         Top             =   7560
         Width           =   2115
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Run"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3240
         TabIndex        =   20
         Top             =   7920
         Width           =   330
      End
      Begin VB.Image Image4i 
         Height          =   510
         Left            =   2880
         Picture         =   "frmMain.frx":89649
         Stretch         =   -1  'True
         Top             =   7800
         Width           =   1035
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         DrawMode        =   5  'Not Copy Pen
         X1              =   2640
         X2              =   3960
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Image Imagfe7 
         Height          =   600
         Left            =   3000
         Picture         =   "frmMain.frx":8AEFB
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   705
      End
      Begin VB.Image Imagef6 
         Height          =   480
         Left            =   3120
         Picture         =   "frmMain.frx":9174D
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   480
      End
      Begin VB.Image Imagef4 
         Height          =   720
         Left            =   3000
         Picture         =   "frmMain.frx":98D5F
         Top             =   4080
         Width           =   720
      End
      Begin VB.Image Imagef3 
         Height          =   720
         Left            =   3000
         Picture         =   "frmMain.frx":99C29
         Top             =   3240
         Width           =   720
      End
      Begin VB.Image Imagfe2 
         Height          =   720
         Left            =   3000
         Picture         =   "frmMain.frx":B42DA
         Top             =   2400
         Width           =   720
      End
      Begin VB.Image Imafge5 
         Height          =   720
         Left            =   3000
         Picture         =   "frmMain.frx":BAB2C
         Top             =   1560
         Width           =   720
      End
      Begin VB.Image Image1ff 
         Height          =   8880
         Left            =   0
         Picture         =   "frmMain.frx":DABAE
         Stretch         =   -1  'True
         Top             =   120
         Width           =   4095
      End
   End
   Begin VB.Frame fraTime 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   3
      Top             =   9000
      Width           =   4560
      Begin VB.CommandButton optStore 
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         Picture         =   "frmMain.frx":1856A4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1800
         Width           =   615
      End
      Begin VB.Timer timTime 
         Interval        =   500
         Left            =   4080
         Top             =   120
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   15
         TabIndex        =   4
         Top             =   0
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   0
         BackColor       =   -2147483636
         Appearance      =   1
         MonthBackColor  =   -2147483636
         StartOfWeek     =   48627713
         TitleBackColor  =   8421504
         TitleForeColor  =   0
         TrailingForeColor=   14737632
         CurrentDate     =   37824
      End
      Begin VB.Label lbltheDate 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.Shape Shape2 
         Height          =   2415
         Left            =   0
         Top             =   0
         Width           =   4545
      End
      Begin VB.Label lblTheTime 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Timer timDown 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   11520
      Top             =   5520
   End
   Begin VB.Timer timUP 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   11520
      Top             =   5520
   End
   Begin VB.Label lbtime 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   10850
      TabIndex        =   24
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Image Imagde4 
      Height          =   390
      Index           =   0
      Left            =   11160
      Picture         =   "frmMain.frx":185A1F
      Top             =   960
      Width           =   435
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Run"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   360
      Picture         =   "frmMain.frx":186351
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   720
   End
   Begin VB.Label lblTrash 
      BackStyle       =   0  'Transparent
      Caption         =   "   Trash"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Lab 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11040
      TabIndex        =   10
      ToolTipText     =   "Begin here"
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblInternet 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Internet"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4560
      Width           =   735
   End
   Begin VB.Image imgInternet 
      Height          =   720
      Left            =   240
      Picture         =   "frmMain.frx":186CEF
      Top             =   3720
      Width           =   720
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   735
   End
   Begin VB.Image ImgTime 
      Height          =   720
      Left            =   240
      Picture         =   "frmMain.frx":1A13A0
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   720
   End
   Begin VB.Label lblAtom 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "System"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Image imgAtom 
      Height          =   720
      Left            =   240
      Picture         =   "frmMain.frx":1A1D98
      Top             =   240
      Width           =   720
   End
   Begin VB.Label lblGO 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11040
      TabIndex        =   0
      ToolTipText     =   "Begin here"
      Top             =   240
      Width           =   735
   End
   Begin VB.Image imgGO 
      Height          =   720
      Left            =   11040
      Picture         =   "frmMain.frx":1B3461
      Top             =   240
      Width           =   720
   End
   Begin VB.Image imgTrash 
      Height          =   480
      Left            =   360
      OLEDropMode     =   2  'Automatic
      Picture         =   "frmMain.frx":1B8C43
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1395
      Left            =   10905
      Shape           =   4  'Rounded Rectangle
      Top             =   180
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   10800
      Left            =   10800
      Picture         =   "frmMain.frx":1BF495
      Stretch         =   -1  'True
      Top             =   -360
      Width           =   2250
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngFromX As Long, lngFromY As Long, bolMoving As Boolean
Dim intMoveIndex As Integer
Private Sub Calculator_Click()
Shell ("Executables\Calculator.exe"), vbNormalFocus
fraMenu.Visible = False
End Sub

Private Sub Command1_Click()
Shell ("Executables\Starsoft Xtreme IExplorer.exe"), vbNormalFocus

fraMenu.Visible = False
End Sub

Private Sub Command2_Click()
Shell ("Executables\ScheduleWizard.exe"), vbNormalFocus

fraMenu.Visible = False
End Sub

Private Sub Check1_Click()
End Sub

Private Sub cmoodContn_Click(Index As Integer)
frmdis.Show
End Sub




Private Sub Command1yuu_Click()
Shell ("Executables\SX Xplorert.exe"), vbNormalFocus
fraMenu.Visible = False
End Sub

Private Sub Commpand1_Click()
myapps.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Apps.Visible = True Then
Apps.Hide
Else
End If
End Sub

Private Sub hh_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Apps.Show
End Sub

Private Sub Imagde4_Click(Index As Integer)
Form3.Show
If fraMenu.Visible = True Then
fraMenu.Visible = False
Else
End If
End Sub

Private Sub Image1_DblClick()
frmRun.Show
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 lngFromX = X
    lngFromY = Y
    bolMoving = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bolMoving = True Then
        Image1.Top = Image1.Top - (lngFromY - Y)
        Image1.Left = Image1.Left - (lngFromX - X)
        Label2.Top = Label2.Top - (lngFromY - Y)
        Label2.Left = Label2.Left - (lngFromX - X)
    End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  bolMoving = False
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Apps.Visible = True Then
Apps.Hide
Else
End If
End Sub

Private Sub Image4i_Click()
frmRun.Show
End Sub

Private Sub Image5_Click()
 Form3.Show
 fraMenu.Visible = False

End Sub

Private Sub Image6_Click()
Shell ("Executables\Starsoft Xtreme  Web Editor.exe"), vbNormalFocus
fraMenu.Visible = False
End Sub

Private Sub Imagee7_Click()
Shell ("Executables\Scheduler.exe"), vbNormalFocus

fraMenu.Visible = False
End Sub

Private Sub Command1f_Click()
Shell ("Executables\Starsoft Xtreme IExplorer.exe"), vbNormalFocus
fraMenu.Visible = False
End Sub

Private Sub Commandff2_Click()
Shell ("Executables\Sheduler.exe"), vbNormalFocus

fraMenu.Visible = False
End Sub

Private Sub Imafge5_Click()
Paint.Show
fraMenu.Visible = False
End Sub

Private Sub Imagef3_Click()
Shell ("Executables\Starsoft Xtreme IExplorer.exe"), vbNormalFocus
fraMenu.Visible = False
End Sub

Private Sub Imagef4_Click()
Shell ("Executables\Starsoft Xtreme MP3 Player.exe"), vbNormalFocus

fraMenu.Visible = False
End Sub

Private Sub Imagef6_Click()
Shell ("Executables\Starsoft Xtreme  Web Editor.exe"), vbNormalFocus
fraMenu.Visible = False
End Sub

Private Sub Imagfe2_Click()
Shell ("Executables\Calculator.exe"), vbNormalFocus
fraMenu.Visible = False
End Sub

Private Sub Imagfe7_Click()
Shell ("Executables\Sheduler.exe"), vbNormalFocus
fraMenu.Visible = False
End Sub

Private Sub Imarrge4_Click()
Shell ("Executables\SX Xplorert.exe"), vbNormalFocus
fraMenu.Visible = False
End Sub

Private Sub Labeiiil1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Apps.Show
End Sub

Private Sub Label2_DblClick()
 frmRun.Show
End Sub


Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Apps.Show
End Sub

Private Sub Label4_Click()
frmRun.Show

End Sub

Private Sub MP3_Click()
Shell ("Executables\Starsoft Xtreme MP3 Player.exe"), vbNormalFocus
fraMenu.Visible = False
End Sub

Private Sub Pain_Click()
Paint.Show
fraMenu.Visible = False

End Sub

Private Sub Timer1_Timer()
lbtime = Time
End Sub

Private Sub web_Click()
Shell ("Executables\Starsoft Xtreme  Web Editor.exe"), vbNormalFocus
fraMenu.Visible = False
End Sub



Private Sub Image4_Click()
Form3.Show
End Sub

Private Sub imgAtom_DblClick()
    lblAtom_Click
End Sub

Private Sub imgAtom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngFromX = X
    lngFromY = Y
    bolMoving = True
End Sub

Private Sub imgAtom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bolMoving = True Then
        imgAtom.Top = imgAtom.Top - (lngFromY - Y)
        imgAtom.Left = imgAtom.Left - (lngFromX - X)
        lblAtom.Top = lblAtom.Top - (lngFromY - Y)
        lblAtom.Left = lblAtom.Left - (lngFromX - X)
    End If
End Sub

Private Sub imgAtom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bolMoving = False
End Sub

Private Sub imgFiles_Click()
    lblFiles_Click
End Sub

Private Sub imgInternet_DblClick()
   Shell ("Executables\Starsoft Xtreme IExplorer.exe"), vbNormalFocus
End Sub

Private Sub imgInternet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngFromX = X
    lngFromY = Y
    bolMoving = True
End Sub

Private Sub imgInternet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bolMoving = True Then
        imgInternet.Top = imgInternet.Top - (lngFromY - Y)
        imgInternet.Left = imgInternet.Left - (lngFromX - X)
        lblInternet.Top = lblInternet.Top - (lngFromY - Y)
        lblInternet.Left = lblInternet.Left - (lngFromX - X)
    End If
End Sub

Private Sub imgInternet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bolMoving = False
End Sub

Private Sub imgRun_Click()
    frmRun.Show
End Sub

Private Sub imgShutDown_Click()
    
End Sub

Private Sub ImgTime_dblClick()
    intMoveIndex = 2
    If fraTime.Top >= 8000 Then
        timDown.Enabled = True
    Else
        timUP.Enabled = True
    End If
End Sub

Private Sub ImgTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngFromX = X
    lngFromY = Y
    bolMoving = True
End Sub

Private Sub ImgTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bolMoving = True Then
        ImgTime.Top = ImgTime.Top - (lngFromY - Y)
        ImgTime.Left = ImgTime.Left - (lngFromX - X)
        lblTime.Top = lblTime.Top - (lngFromY - Y)
        lblTime.Left = lblTime.Left - (lngFromX - X)
    End If
End Sub

Private Sub ImgTime_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bolMoving = False
End Sub

Private Sub imgTrash_DblClick()
    frmDelete.Show
End Sub

Private Sub imgTrash_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lngFromX = X
    lngFromY = Y
    bolMoving = True
End Sub

Private Sub imgTrash_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bolMoving = True Then
        imgTrash.Top = imgTrash.Top - (lngFromY - Y)
        imgTrash.Left = imgTrash.Left - (lngFromX - X)
        lblTrash.Top = lblTrash.Top - (lngFromY - Y)
        lblTrash.Left = lblTrash.Left - (lngFromX - X)
    End If
End Sub

Private Sub imgTrash_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bolMoving = False
End Sub

Private Sub Lab_Click()
fraMenu.Visible = False
lblGO.Visible = True
Lab.Visible = False

End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


End Sub

Private Sub lblAtom_Click()
    frmControls.Show
End Sub

Private Sub lblFiles_Click()
  
End Sub

Private Sub lblFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
End Sub

Private Sub lblFiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub lblGO_Click()
 fraMenu.Visible = True
 lblGO.Visible = False
 Lab.Visible = True
 
End Sub

Private Sub lblRun_Click()
    frmRun.Show
End Sub


   



Private Sub lblShutdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 
End Sub

Private Sub lblShutdown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
End Sub



Private Sub optStore_Click()
    intMoveIndex = 2
    timUP.Enabled = True
End Sub

Private Sub timTime_Timer()
    lblTheTime.Caption = Time
    lbltheDate.Caption = Date
End Sub

Private Sub timUp_Timer()
    If intMoveIndex = 1 Then
        fraMenu.Top = fraMenu.Top + 200
        If fraMenu.Top >= -90 Then
            timUP.Enabled = False
            fraMenu.Top = -90
        End If
    ElseIf intMoveIndex = 2 Then
        fraTime.Top = fraTime.Top + 200
        If fraTime.Top >= 9000 Then
            timUP.Enabled = False
            optStore.Value = False
        End If
    End If
End Sub

Private Sub timDown_Timer()
    If intMoveIndex = 1 Then
        fraMenu.Top = fraMenu.Top - 200
            If fraMenu.Top <= -4070 Then
                timDown.Enabled = False
                fraMenu.Top = -4070
            End If
    ElseIf intMoveIndex = 2 Then
        fraTime.Top = fraTime.Top - 200
        If fraTime.Top <= 6700 Then
            timDown.Enabled = False
            optStore.Value = False
        End If
    End If
End Sub


