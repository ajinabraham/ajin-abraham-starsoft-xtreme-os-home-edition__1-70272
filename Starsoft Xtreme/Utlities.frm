VERSION 5.00
Begin VB.Form Utli 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   Picture         =   "Utlities.frx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Get Icon"
      Height          =   1215
      Left            =   1200
      Picture         =   "Utlities.frx":851F
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Directory Backup Utlity"
      Height          =   1215
      Left            =   0
      Picture         =   "Utlities.frx":ED71
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sheduler"
      Height          =   1215
      Left            =   1200
      Picture         =   "Utlities.frx":F07B
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "HTML Song Attacher"
      Height          =   1215
      Left            =   0
      Picture         =   "Utlities.frx":1CA8C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Easy Web Page Brander"
      Height          =   1215
      Left            =   1200
      Picture         =   "Utlities.frx":1D756
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cab Explorer"
      Height          =   1215
      Left            =   0
      Picture         =   "Utlities.frx":1DB98
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape bo 
      BorderColor     =   &H80000000&
      BorderWidth     =   3
      Height          =   4575
      Left            =   1200
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "Utli"
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
