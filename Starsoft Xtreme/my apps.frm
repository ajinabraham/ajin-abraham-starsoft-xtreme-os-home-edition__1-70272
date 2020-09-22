VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form myapps 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6405
   ClientLeft      =   1005
   ClientTop       =   3495
   ClientWidth     =   7740
   LinkTopic       =   "Form4"
   ScaleHeight     =   6405
   ScaleWidth      =   7740
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox p 
      BackColor       =   &H8000000D&
      Height          =   5895
      Left            =   120
      Picture         =   "my apps.frx":0000
      ScaleHeight     =   5835
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   240
      Width           =   7215
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000E&
         Caption         =   "More Applications"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -120
         TabIndex        =   17
         Top             =   4560
         Width           =   7455
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Browse"
            Height          =   375
            Left            =   1680
            TabIndex        =   22
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtRun 
            Height          =   285
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   3855
         End
         Begin VB.CommandButton cmdRun 
            Caption         =   "Run"
            Height          =   375
            Left            =   3120
            TabIndex        =   20
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton Command16 
            BackColor       =   &H8000000E&
            Caption         =   "How To Run More Apps"
            Height          =   1095
            Left            =   5760
            Picture         =   "my apps.frx":3C0042
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   120
            Width           =   1455
         End
         Begin MSComDlg.CommonDialog CD2 
            Left            =   4200
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H8000000E&
         Caption         =   "How To Install Your Apps"
         Height          =   1815
         Left            =   6360
         Picture         =   "my apps.frx":3C08EE
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmd9 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         Picture         =   "my apps.frx":3C119A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmd5 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Picture         =   "my apps.frx":3C739C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmd12 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         Picture         =   "my apps.frx":3CD59E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmd13 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         Picture         =   "my apps.frx":3D37A0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton cmd10 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         Picture         =   "my apps.frx":3D99A2
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton cmd8 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         Picture         =   "my apps.frx":3DFBA4
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmd11 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         Picture         =   "my apps.frx":3E5DA6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton cmd4 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Picture         =   "my apps.frx":3EBFA8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton cmd2 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Picture         =   "my apps.frx":3F21AA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmd3 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Picture         =   "my apps.frx":3F83AC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton cmd6 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Picture         =   "my apps.frx":3FE5AE
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton cmd7 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Picture         =   "my apps.frx":4047B0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3840
         Width           =   1695
      End
      Begin VB.CommandButton cmd14 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         Picture         =   "my apps.frx":40A9B2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3840
         Width           =   1695
      End
      Begin VB.CommandButton cmd1 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Picture         =   "my apps.frx":410BB4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.Image Image21 
         Height          =   615
         Left            =   5520
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   735
      End
      Begin VB.Image Image7 
         Height          =   615
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   720
         Width           =   735
      End
      Begin VB.Image Image6 
         Height          =   615
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   735
      End
      Begin VB.Image Image5 
         Height          =   615
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   735
      End
      Begin VB.Image Image4 
         Height          =   615
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image Image3 
         Height          =   615
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   735
      End
      Begin VB.Image Image2 
         Height          =   615
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
      Begin VB.Image Image20 
         Height          =   615
         Left            =   5520
         Stretch         =   -1  'True
         Top             =   240
         Width           =   735
      End
      Begin VB.Image Image19 
         Height          =   615
         Left            =   5520
         Stretch         =   -1  'True
         Top             =   840
         Width           =   735
      End
      Begin VB.Image Image18 
         Height          =   615
         Left            =   5520
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   735
      End
      Begin VB.Image Image17 
         Height          =   615
         Left            =   5520
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   735
      End
      Begin VB.Image Image16 
         Height          =   615
         Left            =   5520
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image Image15 
         Height          =   615
         Left            =   5520
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "My Applications"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6375
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "myapps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso1 As New FileSystemObject
Dim fso2 As New FileSystemObject
Dim fso3 As New FileSystemObject
Dim fso4 As New FileSystemObject
Dim fso5 As New FileSystemObject
Dim fso6 As New FileSystemObject
Dim fso7 As New FileSystemObject
Dim fso8 As New FileSystemObject
Dim fso9 As New FileSystemObject
Dim fso10 As New FileSystemObject
Dim fso11 As New FileSystemObject
Dim fso12 As New FileSystemObject
Dim fso13 As New FileSystemObject
Dim fso14 As New FileSystemObject
Dim strm1 As TextStream
Dim strm2 As TextStream
Dim strm3 As TextStream
Dim strm4 As TextStream
Dim strm5 As TextStream
Dim strm6 As TextStream
Dim strm7 As TextStream
Dim strm8 As TextStream
Dim strm9 As TextStream
Dim strm10 As TextStream
Dim strm11 As TextStream
Dim strm12 As TextStream
Dim strm13 As TextStream
Dim strm14 As TextStream
Dim name1 As String
Dim name2 As String
Dim name3 As String
Dim name4 As String
Dim name5 As String
Dim name6 As String
Dim name7 As String
Dim name8 As String
Dim name9 As String
Dim name10 As String
Dim name11 As String
Dim name12 As String
Dim name13 As String
Dim name14 As String


Private Sub cmdBrowse_Click()
CD2.ShowOpen
    txtRun.Text = CD2.FileName
End Sub

Private Sub cmdClose_Click()
Me.Hide
End Sub

Private Sub cmdRun_Click()
 On Error GoTo Error
    Shell (txtRun.Text)
    Unload frmRun
Error: MsgBox "Select any executable file to run", vbOKOnly, "Starsoft Xtreme"
End Sub

Private Sub Command15_Click()
hwtoinstall.Show
End Sub

Private Sub Command16_Click()
MsgBox "Clck on browse button,select the executable file you want to run and click on Run button.", vbDefaultButton1, "Starsoft Xtreme"
End Sub

Private Sub Form_Load()
 
    Set strm1 = fso1.OpenTextFile(App.path + "\MyApps\1\1.txt")
        name1 = strm1.ReadLine
        cmd1.Caption = name1
    strm1.Close
      Set strm2 = fso2.OpenTextFile(App.path + "\MyApps\2\2.txt")
      name2 = strm2.ReadLine
      cmd2.Caption = name2
   strm2.Close

Set strm3 = fso3.OpenTextFile(App.path + "\MyApps\3\3.txt")
      name3 = strm3.ReadLine
      cmd3.Caption = name3
   strm3.Close
   Set strm4 = fso4.OpenTextFile(App.path + "\MyApps\4\4.txt")
      name4 = strm4.ReadLine
      cmd4.Caption = name4
   strm4.Close
   Set strm5 = fso5.OpenTextFile(App.path + "\MyApps\5\5.txt")
      name5 = strm5.ReadLine
      cmd5.Caption = name5
   strm5.Close
   Set strm6 = fso6.OpenTextFile(App.path + "\MyApps\6\6.txt")
      name6 = strm6.ReadLine
      cmd6.Caption = name6
   strm6.Close
   Set strm7 = fso7.OpenTextFile(App.path + "\MyApps\7\7.txt")
      name7 = strm7.ReadLine
      cmd7.Caption = name7
   strm7.Close
   Set strm8 = fso8.OpenTextFile(App.path + "\MyApps\8\8.txt")
      name8 = strm8.ReadLine
      cmd8.Caption = name8
   strm8.Close
   Set strm9 = fso9.OpenTextFile(App.path + "\MyApps\9\9.txt")
      name9 = strm9.ReadLine
      cmd9.Caption = name9
   strm9.Close
   Set strm10 = fso10.OpenTextFile(App.path + "\MyApps\10\10.txt")
      name10 = strm10.ReadLine
      cmd10.Caption = name10
   strm10.Close
   Set strm11 = fso11.OpenTextFile(App.path + "\MyApps\11\11.txt")
      name11 = strm11.ReadLine
      cmd11.Caption = name11
   strm11.Close
   Set strm12 = fso12.OpenTextFile(App.path + "\MyApps\12\12.txt")
      name12 = strm12.ReadLine
      cmd12.Caption = name12
   strm12.Close
   Set strm13 = fso13.OpenTextFile(App.path + "\MyApps\13\13.txt")
      name13 = strm13.ReadLine
      cmd13.Caption = name13
   strm13.Close
Set strm14 = fso14.OpenTextFile(App.path + "\MyApps\14\14.txt")
      name14 = strm14.ReadLine
      cmd14.Caption = name14
   strm14.Close


End Sub

