VERSION 5.00
Begin VB.Form Secur 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Security.frx":0000
   ScaleHeight     =   2430
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "ULTRAVIRES Safe Note"
      Height          =   1215
      Left            =   1200
      Picture         =   "Security.frx":851F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Folder Security"
      Height          =   1215
      Left            =   1200
      Picture         =   "Security.frx":32DD1
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ULTRAVIRES File Encrypter"
      Height          =   1215
      Left            =   0
      Picture         =   "Security.frx":33C9B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Vitus Buster2008"
      Height          =   1215
      Left            =   0
      Picture         =   "Security.frx":5E54D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   1215
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape bo 
      BorderColor     =   &H80000000&
      BorderWidth     =   3
      Height          =   4575
      Left            =   -1200
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Secur"
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
