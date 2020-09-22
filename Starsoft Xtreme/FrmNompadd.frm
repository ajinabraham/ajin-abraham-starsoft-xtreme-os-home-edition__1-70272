VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmNompadd 
   BorderStyle     =   0  'None
   ClientHeight    =   8880
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "FrmNompadd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox RTBmain 
      Height          =   8295
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   14631
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"FrmNompadd.frx":2A8B2
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   3  'Align Left
      Height          =   8880
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   15663
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageKey        =   "Align Left"
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageKey        =   "Align Right"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNompadd.frx":2A934
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNompadd.frx":2AA46
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNompadd.frx":2AB58
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNompadd.frx":2AC6A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNompadd.frx":2AD7C
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNompadd.frx":2AE8E
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNompadd.frx":2AFA0
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNompadd.frx":2B0B2
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNompadd.frx":2B1C4
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNompadd.frx":2B2D6
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNompadd.frx":2B3E8
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNompadd.frx":2B4FA
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNompadd.frx":2B60C
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNompadd.frx":2B71E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmNompadd.frx":2B832
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin StarsoftXtremeOS.Window Window1 
      Height          =   9255
      Left            =   -240
      TabIndex        =   2
      Top             =   -840
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   16325
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   525
      Left            =   0
      Top             =   0
      Width           =   600
   End
   Begin VB.Image imgRestore 
      Height          =   450
      Left            =   120
      Picture         =   "FrmNompadd.frx":2BD76
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
End
Attribute VB_Name = "frmNompadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub Form_Resize()
Window1.Top = 0
Window1.Left = 0
Window1.Height = Me.ScaleHeight
Window1.Width = Me.ScaleWidth

    On Error Resume Next
    tbToolBar.Visible = True
   
End Sub

Private Sub imgRestore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.Hide: Exit Sub
        Me.WindowState = 2
        Me.BackColor = vbButtonFace
        Me.tbToolBar.Visible = True
        imgRestore.Visible = False
End Sub

Private Sub optClose_Click()
    Unload frmNompadd
End Sub

Private Sub optStore_Click()
    Me.WindowState = 0
    Me.Height = 500
    Me.Width = 1000
    Me.BackColor = &HC00000
    tbToolBar.Visible = False
    imgRestore.Visible = True
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Dim sFile As String
    Select Case Button.Key
        Case "New"
            Dim SaveCheck As Integer
            If RTBmain.Text <> "" Then SaveCheck = MsgBox("Do you want to quit without saving?", vbYesNo, "Nompadd")
                If SaveCheck = vbYes Then
                    RTBmain.Text = ""
                Else
                    Exit Sub
                End If
        Case "Font"
            With dlgCommonDialog
                .Flags = 3
                .DialogTitle = "Fonts"
                .CancelError = True
                .ShowFont
            End With
            RTBmain.SelFontName = dlgCommonDialog.FontName
            RTBmain.SelFontSize = dlgCommonDialog.FontSize
            RTBmain.SelBold = dlgCommonDialog.FontBold
            RTBmain.SelItalic = dlgCommonDialog.FontItalic
            RTBmain.SelUnderline = dlgCommonDialog.FontUnderline
        Case "Open"
            With dlgCommonDialog
                .DialogTitle = "Open"
                .CancelError = False
                .Filter = "Text File|*.txt|DAT file|*.dat|All Files|*.*"
                .ShowOpen
                If Len(.FileName) = 0 Then
                    Exit Sub
                End If
            sFile = .FileName
            End With
            RTBmain.LoadFile sFile
        Case "Save"
            With dlgCommonDialog
                .DialogTitle = "Save"
                .CancelError = False
                .Filter = "Text File|*.txt|DAT file|*.dat"
                .ShowSave
                If Len(.FileName) = 0 Then
                    Exit Sub
                End If
                sFile = .FileName
            End With
                RTBmain.SaveFile sFile
                RTBmain.SaveFile sFile
        Case "Print"
            With dlgCommonDialog
                .DialogTitle = "Print"
                .CancelError = True
                .Flags = cdlPDReturnDC + cdlPDNoPageNums
                .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            RTBmain.SelPrint .hDC
        End If
    End With
        Case "Cut"
            Clipboard.SetText RTBmain.SelRTF
            RTBmain.SelText = vbNullString
        Case "Copy"
           Clipboard.SetText RTBmain.SelRTF
        Case "Paste"
            RTBmain.SelRTF = Clipboard.GetText
        Case "Bold"
            RTBmain.SelBold = Not RTBmain.SelBold
            Button.Value = IIf(RTBmain.SelBold, tbrPressed, tbrUnpressed)
        Case "Italic"
            RTBmain.SelItalic = Not RTBmain.SelItalic
            Button.Value = IIf(RTBmain.SelItalic, tbrPressed, tbrUnpressed)
        Case "Underline"
            RTBmain.SelUnderline = Not RTBmain.SelUnderline
            Button.Value = IIf(RTBmain.SelUnderline, tbrPressed, tbrUnpressed)
        Case "Center"
            RTBmain.SelAlignment = rtfCenter
            Button.Value = IIf(RTBmain.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
        Case "Align Left"
            RTBmain.SelAlignment = rtfLeft
        Case "Align Right"
            RTBmain.SelAlignment = rtfRight
    End Select
End Sub

