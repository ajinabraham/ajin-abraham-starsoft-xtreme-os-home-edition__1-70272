VERSION 5.00
Begin VB.Form commander 
   BorderStyle     =   0  'None
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Appearance      =   0  'Flat
      Caption         =   "CD Tray"
      Height          =   975
      Left            =   1440
      Picture         =   "commander.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Appearance      =   0  'Flat
      Caption         =   "CD Tray"
      Height          =   975
      Left            =   2640
      Picture         =   "commander.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Appearance      =   0  'Flat
      Caption         =   "CD Tray"
      Height          =   975
      Left            =   5280
      Picture         =   "commander.frx":3994
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      Caption         =   "CD Tray"
      Height          =   975
      Left            =   3960
      Picture         =   "commander.frx":565E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      Caption         =   "CD Tray"
      Height          =   975
      Left            =   5280
      Picture         =   "commander.frx":7328
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      Caption         =   "CD Tray"
      Height          =   975
      Left            =   6600
      Picture         =   "commander.frx":8FF2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "CD Tray"
      Height          =   975
      Left            =   120
      Picture         =   "commander.frx":ACBC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "CD Tray"
      Height          =   975
      Left            =   1320
      Picture         =   "commander.frx":C986
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "CD Tray"
      Height          =   975
      Left            =   2640
      Picture         =   "commander.frx":E650
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "CD Tray"
      Height          =   975
      Left            =   3960
      Picture         =   "commander.frx":1031A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton sas 
      Appearance      =   0  'Flat
      Caption         =   "CD Tray"
      Height          =   975
      Left            =   240
      Picture         =   "commander.frx":11FE4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin StarsoftXtremeOS.Window Window1 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9340
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         Caption         =   "CD Tray"
         Height          =   975
         Left            =   6600
         Picture         =   "commander.frx":13CAE
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Commander"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         Top             =   0
         Width           =   1455
      End
   End
End
Attribute VB_Name = "commander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
Window1.Top = 0
Window1.Left = 0
Window1.Height = Me.ScaleHeight
Window1.Width = Me.ScaleWidth


End Sub

