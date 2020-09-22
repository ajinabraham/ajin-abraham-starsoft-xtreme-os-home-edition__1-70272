VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDelete 
   BorderStyle     =   0  'None
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtDelete 
      Height          =   285
      Left            =   230
      TabIndex        =   2
      Top             =   600
      Width           =   3855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin StarsoftXtremeOS.Windowtwo Windowtwo1 
      Height          =   1695
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2990
      Begin VB.Label lblDelete 
         BackStyle       =   0  'Transparent
         Caption         =   "Delete -enter a file name and extention"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   3495
      End
   End
   Begin MSComDlg.CommonDialog sx4 
      Left            =   3720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
    sx4.ShowOpen
    txtDelete.Text = sx4.FileName
End Sub

Private Sub cmdClose_Click()
    Unload frmDelete
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo Error
    Dim Confirm As Integer
    If txtDelete.Text = "" Then Exit Sub
    Confirm = MsgBox("Are you sure you want this item to be premently deleted", vbYesNo, "Confirm")
    If Confirm = vbNo Then
        frmMain.SetFocus
        Exit Sub
    Else
        Kill txtDelete.Text
        txtDelete.Text = ""
    End If
    frmMain.SetFocus
    Exit Sub
Error:
    MsgBox Err.Description & "  -Folders can not be deleted"
    frmMain.SetFocus
End Sub

Private Sub Form_Resize()
  Windowtwo1.Top = 0
Windowtwo1.Left = 0
Windowtwo1.Height = Me.ScaleHeight
Windowtwo1.Width = Me.ScaleWidth

End Sub

