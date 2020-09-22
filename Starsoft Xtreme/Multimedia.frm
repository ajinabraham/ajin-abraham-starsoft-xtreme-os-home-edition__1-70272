VERSION 5.00
Begin VB.Form Mult 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   Picture         =   "Multimedia.frx":0000
   ScaleHeight     =   3930
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Multimedia Player"
      Height          =   975
      Left            =   0
      Picture         =   "Multimedia.frx":851F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   " MP3 Player"
      Height          =   975
      Left            =   1200
      Picture         =   "Multimedia.frx":301B1
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Thunder MP3 Player"
      Height          =   975
      Left            =   0
      Picture         =   "Multimedia.frx":30A7B
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
Attribute VB_Name = "Mult"
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
