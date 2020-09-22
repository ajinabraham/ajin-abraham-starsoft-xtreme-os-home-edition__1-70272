VERSION 5.00
Begin VB.Form frmfullscreen 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Full Screen"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1800
      ToolTipText     =   "Press Any Key Or Click to Exit"
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmfullscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    Image1_Click
End Sub

Private Sub Form_Load()
    On Error Resume Next
    frmfullscreen.WindowState = 2
    Image1.Top = 0
    Image1.Left = 0
    Image1.Height = Screen.Height
    Image1.Width = Screen.Width
    Image1.Stretch = True
    Image1.Refresh
End Sub

Private Sub Image1_Click()
    On Error Resume Next
    Unload Me
End Sub
