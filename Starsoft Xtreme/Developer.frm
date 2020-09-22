VERSION 5.00
Begin VB.Form Dev 
   BorderStyle     =   0  'None
   Caption         =   "Devloper"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   Picture         =   "Developer.frx":0000
   ScaleHeight     =   2415
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape bo 
      BorderColor     =   &H80000000&
      BorderWidth     =   3
      Height          =   255
      Left            =   1680
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Dev"
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
