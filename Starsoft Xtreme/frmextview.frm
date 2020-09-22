VERSION 5.00
Begin VB.Form frmextview 
   BackColor       =   &H80000007&
   Caption         =   " Extrenal Image Viewer"
   ClientHeight    =   4800
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   6600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picbox 
      AutoSize        =   -1  'True
      Height          =   2175
      Left            =   2760
      ScaleHeight     =   2115
      ScaleWidth      =   1995
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2040
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu cmdcount 
         Caption         =   "&Count Images"
         Shortcut        =   ^I
      End
      Begin VB.Menu cmdchange 
         Caption         =   "Change &Images"
         Shortcut        =   ^C
      End
      Begin VB.Menu cmdexitview 
         Caption         =   "E&xit Viewer"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu cmdfull 
         Caption         =   "Full &Screen"
         Shortcut        =   ^F
      End
      Begin VB.Menu cmdslide 
         Caption         =   "Screensaver"
         Shortcut        =   ^S
      End
      Begin VB.Menu cmdviewline1 
         Caption         =   "-"
      End
      Begin VB.Menu cmdnext 
         Caption         =   "&Next Picture"
         Shortcut        =   ^N
      End
      Begin VB.Menu cmdprevious 
         Caption         =   "&Previous Picture"
         Shortcut        =   ^P
      End
      Begin VB.Menu cmdviewline2 
         Caption         =   "-"
      End
      Begin VB.Menu cmdzoomin 
         Caption         =   "Zoom In"
         Shortcut        =   {F11}
      End
      Begin VB.Menu cmdzoomout 
         Caption         =   "Zoom Out"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "frmextview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim current As Integer
Dim s As Integer

Private Sub cmdChange_Click()
    On Error Resume Next
    frmdis.cmdextview.Caption = "&Images Selected"
    frmdis.cmdexit.Enabled = False
    frmdis.cmdfull.Enabled = False
    Unload Me
    frmdis.Show 1
End Sub

Private Sub cmdcount_Click()
    On Error Resume Next
        s = MsgBox("Total Images = " & tis, vbInformation + vbOKOnly, "External Viewer Image Counter")
End Sub

Private Sub cmdexitview_Click()
    On Error Resume Next
    frmdis.cmdextview.Caption = "&External Viewer"
    frmdis.cmdexit.Enabled = True
    frmdis.cmdfull.Enabled = True
    Unload Me
    frmdis.Show
End Sub

Private Sub cmdfull_Click()
    On Error Resume Next
    frmfullscreen.Image1.Picture = LoadPicture(images(current))
    frmfullscreen.Show 1
End Sub

Private Sub cmdnext_Click()
    On Error Resume Next
    If current < tis - 1 Then
            current = current + 1
            Image1.Visible = False
            Image1.Stretch = False
            picbox.Picture = LoadPicture(images(current))
            Image1.Height = picbox.Height
            Image1.Width = picbox.Width
            Image1.Top = (Screen.Height - Image1.Height) / 2
            Image1.Left = (Screen.Width - Image1.Width) / 2
            Image1.Picture = LoadPicture(images(current))
            Me.Caption = " External Viewer [ " & images(current) & " ]" '& " Resolution:  [ " & Int(ScaleX(picbox.Width - 4, vbTwips, vbPixels)) & " * " & Int(ScaleY(picbox.Height - 4, vbTwips, vbPixels)) & " ]"
            Image1.Stretch = True
            Image1.Refresh
            Image1.Visible = True
    Else
        s = MsgBox("This is Last Image", vbInformation, "External Viewer Error")
    End If
End Sub

Private Sub cmdprevious_Click()
    On Error Resume Next
    If current > 0 Then
            current = current - 1
            Image1.Visible = False
            Image1.Stretch = False
            picbox.Picture = LoadPicture(images(current))
            Image1.Height = picbox.Height
            Image1.Width = picbox.Width
            Image1.Top = (Screen.Height - Image1.Height) / 2
            Image1.Left = (Screen.Width - Image1.Width) / 2
            Image1.Picture = LoadPicture(images(current))
            Me.Caption = " External Viewer [ " & images(current) & " ]" '& " Resolution:  [ " & Int(ScaleX(picbox.Width - 4, vbTwips, vbPixels)) & " * " & Int(ScaleY(picbox.Height - 4, vbTwips, vbPixels)) & " ]"
            Image1.Stretch = True
            Image1.Refresh
            Image1.Visible = True
    Else
        s = MsgBox("This is First Image", vbInformation, "External Viewer Error")
    End If
End Sub

Private Sub cmdslide_Click()
    On Error Resume Next
    frmslideshow.Show 1
End Sub

Private Sub cmdzoomin_Click()
    On Error Resume Next
    Image1.Visible = False
    Image1.Height = Image1.Height + (Image1.Height / 5)
    Image1.Width = Image1.Width + (Image1.Width / 5)
    Image1.Top = (Screen.Height - Image1.Height) / 2
    Image1.Left = (Screen.Width - Image1.Width) / 2
    Image1.Refresh
    Image1.Visible = True
End Sub

Private Sub cmdzoomout_Click()
    On Error Resume Next
    Image1.Visible = False
    Image1.Height = Image1.Height - (Image1.Height / 5)
    Image1.Width = Image1.Width - (Image1.Width / 5)
    Image1.Top = (Screen.Height - Image1.Height) / 2
    Image1.Left = (Screen.Width - Image1.Width) / 2
    Image1.Refresh
    Image1.Visible = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
If KeyCode = 37 Then
    If Image1.Width > Screen.Width And Image1.Left > Screen.Width - Image1.Width Then
        Image1.Left = Image1.Left - 200
    End If
ElseIf KeyCode = 38 Then
    If Image1.Height > Screen.Height And Image1.Top > Screen.Height - Image1.Height Then
        Image1.Top = Image1.Top - 200
    End If
ElseIf KeyCode = 39 Then
    If Image1.Left < 0 Then
        Image1.Left = Image1.Left + 200
    End If
ElseIf KeyCode = 40 Then
    If Image1.Top < 0 Then
        Image1.Top = Image1.Top + 200
    End If
End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    current = 0
    Image1.Stretch = False
    picbox.Picture = LoadPicture(images(current))
    Image1.Height = picbox.Height
    Image1.Width = picbox.Width
    Image1.Top = (Screen.Height - Image1.Height) / 2
    Image1.Left = (Screen.Width - Image1.Width) / 2
    Image1.Picture = LoadPicture(images(current))
    Me.Caption = " External Viewer [ " & images(current) & " ]" '& " Resolution:  [ " & Int(ScaleX(picbox.Width - 4, vbTwips, vbPixels)) & " * " & Int(ScaleY(picbox.Height, vbTwips, vbPixels)) & " ]"
    Image1.Stretch = True
    Image1.Refresh
End Sub

Private Sub Form_Resize()
    Me.WindowState = vbMaximized
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Image1.Visible = False
    If Button = vbLeftButton Then
        Image1.Height = Image1.Height + (Image1.Height / 5)
        Image1.Width = Image1.Width + (Image1.Width / 5)
    ElseIf Button = vbRightButton Then
        Image1.Height = Image1.Height - (Image1.Height / 5)
        Image1.Width = Image1.Width - (Image1.Width / 5)
    End If
    Image1.Top = (Screen.Height - Image1.Height) / 2
    Image1.Left = (Screen.Width - Image1.Width) / 2
    Image1.Refresh
    Image1.Visible = True
End Sub
