VERSION 5.00
Begin VB.Form Fo1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Information"
   ClientHeight    =   7155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   7095
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Boot installer.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   7215
   End
End
Attribute VB_Name = "Fo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Height = Screen.Height
Me.Width = Screen.Width
Me.Top = 0
Me.Left = 0
Text1.Height = Screen.Height
Text1.Width = Screen.Width
Text1.Left = 0
Text1.Top = 0

End Sub

