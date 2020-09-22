VERSION 5.00
Begin VB.Form Offi 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   Picture         =   "Office.frx":0000
   ScaleHeight     =   5415
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "3D Motion + Morphing"
      Height          =   1095
      Left            =   1200
      Picture         =   "Office.frx":851F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Calculator"
      Height          =   1095
      Left            =   0
      Picture         =   "Office.frx":285A1
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Word Processor 2008"
      Height          =   1095
      Left            =   1200
      Picture         =   "Office.frx":2EDF3
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Advanced Paint"
      Height          =   1095
      Left            =   0
      Picture         =   "Office.frx":596A5
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Glaze Paint"
      Height          =   1095
      Left            =   1200
      Picture         =   "Office.frx":59F6F
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Wonder HTML"
      Height          =   1095
      Left            =   0
      Picture         =   "Office.frx":79FF1
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Web Editor"
      Height          =   1095
      Left            =   1200
      Picture         =   "Office.frx":7F7D3
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Digital Book 2008"
      Height          =   1095
      Left            =   1200
      Picture         =   "Office.frx":86DE5
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Quick Mailer"
      Height          =   1095
      Left            =   0
      Picture         =   "Office.frx":AEA77
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Digital Book 2000"
      Height          =   1095
      Left            =   0
      Picture         =   "Office.frx":D6709
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape bo 
      BorderColor     =   &H80000000&
      BorderWidth     =   3
      Height          =   4575
      Left            =   -600
      Top             =   -120
      Width           =   2535
   End
End
Attribute VB_Name = "Offi"
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
