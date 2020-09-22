VERSION 5.00
Begin VB.Form Apps 
   BorderStyle     =   0  'None
   Caption         =   "apps"
   ClientHeight    =   5655
   ClientLeft      =   3345
   ClientTop       =   450
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   Picture         =   "Apps.frx":0000
   ScaleHeight     =   5655
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Multimedia"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Apps.frx":3C0042
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Utlities"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "Apps.frx":3C1082
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Office"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "Apps.frx":3C20C2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Games"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "Apps.frx":3C3102
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Developer"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "Apps.frx":3C4142
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Security"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "Apps.frx":3C5182
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000E&
      Caption         =   "Applications"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "Apps.frx":3C61C2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.Shape Shape13 
      BorderColor     =   &H8000000B&
      Height          =   615
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   615
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H8000000B&
      Height          =   615
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H8000000B&
      Height          =   615
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   615
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H8000000B&
      Height          =   615
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H8000000B&
      Height          =   615
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   615
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H8000000B&
      Height          =   615
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000B&
      Height          =   615
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   615
   End
   Begin VB.Image g 
      Height          =   375
      Left            =   2640
      Picture         =   "Apps.frx":3C7202
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   375
   End
   Begin VB.Image f 
      Height          =   375
      Left            =   2640
      Picture         =   "Apps.frx":3C7ACC
      Stretch         =   -1  'True
      Top             =   600
      Width           =   375
   End
   Begin VB.Image e 
      Height          =   375
      Left            =   2640
      Picture         =   "Apps.frx":3E8029
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   375
   End
   Begin VB.Image d 
      Height          =   480
      Left            =   2640
      Picture         =   "Apps.frx":3E88F3
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image c 
      Height          =   495
      Left            =   2640
      Picture         =   "Apps.frx":3EF145
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image b 
      Height          =   345
      Left            =   2640
      Picture         =   "Apps.frx":3EFD07
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image a 
      Height          =   525
      Left            =   2520
      Picture         =   "Apps.frx":3F7319
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   570
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000A&
      BorderWidth     =   3
      FillColor       =   &H00808080&
      Height          =   6015
      Left            =   0
      Top             =   -360
      Width           =   3375
   End
   Begin VB.Shape shp1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      FillColor       =   &H00808080&
      Height          =   5295
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Apps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Picture1_Click()

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Font = "&H8000000E&"
Mult.Show
Gam.Hide
Offi.Hide
allapps.Hide
Dev.Hide

Secur.Hide
Utli.Hide
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.Font = "&H8000000E&"
allapps.Show
Dev.Hide
Gam.Hide
Mult.Hide
Offi.Hide
Secur.Hide
Utli.Hide
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command3.Font = "&H8000000E&"
Secur.Show
Gam.Hide
Offi.Hide
allapps.Hide
Dev.Hide
Mult.Hide
Utli.Hide
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command4.Font = "&H8000000E&"
Dev.Show
Gam.Hide
Offi.Hide
allapps.Hide
Mult.Hide
Secur.Hide
Utli.Hide
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command5.Font = "&H8000000E&"
Gam.Show
Offi.Hide
allapps.Hide
Dev.Hide
Mult.Hide
Secur.Hide
Utli.Hide
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command6.Font = "&H8000000E&"
Offi.Show
allapps.Hide
Dev.Hide
Gam.Hide
Mult.Hide
Secur.Hide
Utli.Hide
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command7.Font = "&H8000000E&"
Utli.Show
Gam.Hide
Offi.Hide
allapps.Hide
Dev.Hide
Mult.Hide
Secur.Hide
End Sub

Private Sub Form_Load()
Shape2.Height = Apps.Height
Shape2.Width = Apps.Width
Shape2.Left = 0
Shape2.Top = 0
End Sub
