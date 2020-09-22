VERSION 5.00
Begin VB.Form frmslideshow 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Screensaver"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   2760
   End
   Begin VB.Image imgslide 
      Height          =   495
      Left            =   0
      Stretch         =   -1  'True
      ToolTipText     =   "Press any Key Or Click Mouse For Exiting"
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmslideshow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim t As Integer
Dim running As Boolean

Private Sub Form_Activate()
    If frmdis.File2.ListIndex < 0 Then
        imgslide_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    imgslide_Click
End Sub

Private Sub Form_Load()
    On Error Resume Next
    t = 0
    frmslideshow.Left = 0
    frmslideshow.Top = 0
    frmslideshow.Height = Screen.Height
    frmslideshow.Width = Screen.Width
    imgslide.Top = 0
    imgslide.Left = 0
    imgslide.Height = Screen.Height
    imgslide.Width = Screen.Width
    Timer1.Enabled = True
    running = False
    t = 0
End Sub

Private Sub imgslide_Click()
    On Error Resume Next
    Timer1.Enabled = False
    Unload Me
   Unload frmdis
    
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    If running = False Then
        running = True
        frmslideshow.Timer1.Interval = Val(frmdis.txttimer.Text)
    End If
    If t = tis Then
        If frmdis.chkcyclic.Value = Checked And tis > 0 Then
            t = 0
            imgslide.Picture = LoadPicture(images(t))
            t = t + 1
        Else
            Timer1.Enabled = False
            Unload Me
        End If
    Else
    imgslide.Picture = LoadPicture(images(t))
    t = t + 1
    End If
End Sub
