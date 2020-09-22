VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRun 
   BorderStyle     =   0  'None
   ClientHeight    =   1590
   ClientLeft      =   75
   ClientTop       =   75
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   4215
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtRun 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3855
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin StarsoftXtremeOS.Windowtwo Windowtwo1 
      Height          =   1575
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2778
      Begin VB.Label lblRun 
         BackStyle       =   0  'Transparent
         Caption         =   "Run -enter a file name and extention"
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
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3255
      End
   End
   Begin MSComDlg.CommonDialog CD2 
      Left            =   4200
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
    CD2.ShowOpen
    txtRun.Text = CD2.FileName
End Sub

Private Sub cmdClose_Click()
    Unload frmRun
End Sub

Private Sub cmdRun_Click()
    On Error GoTo Error
    Shell (txtRun.Text)
    Unload frmRun
Error: MsgBox "Select any executable file to run", vbOKOnly, "Starsoft Xtreme"
End Sub

Private Sub Form_Resize()
  Windowtwo1.Top = 0
Windowtwo1.Left = 0
Windowtwo1.Height = Me.ScaleHeight
Windowtwo1.Width = Me.ScaleWidth
  
End Sub

