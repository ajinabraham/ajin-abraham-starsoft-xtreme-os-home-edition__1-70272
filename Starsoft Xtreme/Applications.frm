VERSION 5.00
Begin VB.Form allapps 
   BorderStyle     =   0  'None
   Caption         =   "apps"
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   Picture         =   "Applications.frx":0000
   ScaleHeight     =   5130
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape bo 
      BorderColor     =   &H80000000&
      BorderWidth     =   3
      Height          =   5175
      Left            =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "allapps"
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
