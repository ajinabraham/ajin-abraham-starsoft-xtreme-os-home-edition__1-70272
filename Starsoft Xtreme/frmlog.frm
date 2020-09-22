VERSION 5.00
Begin VB.Form frmlogin 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   Icon            =   "frmlog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmlog.frx":2055D
   ScaleHeight     =   420
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   544
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   5040
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Type Your Password Here"
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox txtUser 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Starsoft Xtreme OS"
      ToolTipText     =   "Username"
      Top             =   4200
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      DragIcon        =   "frmlog.frx":3E059F
      Height          =   2175
      Left            =   3720
      Picture         =   "frmlog.frx":3E08A9
      ScaleHeight     =   2115
      ScaleWidth      =   4275
      TabIndex        =   2
      ToolTipText     =   "Login User Interface"
      Top             =   3480
      Width           =   4335
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   3600
         Picture         =   "frmlog.frx":6208EB
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "GO"
         Top             =   840
         Width           =   495
      End
      Begin VB.Image Imagde4 
         Height          =   390
         Index           =   1
         Left            =   3600
         Picture         =   "frmlog.frx":6211B5
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Password"
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Username"
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Xtreme OS"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   4560
      TabIndex        =   6
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Starsoft"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   1560
      TabIndex        =   5
      Top             =   1560
      Width           =   2775
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

frmMain.Show

End Sub

Private Sub Form_Load()
Me.Height = Screen.Height
Me.Width = Screen.Width
Me.Top = 0
Me.Left = 0
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Imagde4_Click(Index As Integer)
Form3.Show
End Sub

Private Sub Timer1_Timer()

End Sub
