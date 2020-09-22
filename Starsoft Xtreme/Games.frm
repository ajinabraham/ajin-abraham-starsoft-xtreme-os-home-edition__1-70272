VERSION 5.00
Begin VB.Form Gam 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   Picture         =   "Games.frx":0000
   ScaleHeight     =   3720
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Navel Battle"
      Height          =   1215
      Left            =   1200
      Picture         =   "Games.frx":851F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Coloroids 2"
      Height          =   1215
      Left            =   0
      Picture         =   "Games.frx":91E9
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape bo 
      BorderColor     =   &H80000000&
      BorderWidth     =   3
      Height          =   4575
      Left            =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Gam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
bo.Height = Me.Height
bo.Width = Me.Width
bo.Left = 0
bo.Top = 0
End Sub
