VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00AFAFAF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Green Screen Shop"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6615
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture7 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   420
      Left            =   4200
      Picture         =   "frmMain.frx":0CCE
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   27
      Top             =   3960
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   7200
      Width           =   5535
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00AFAFAF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7200
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   7680
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00AFAFAF&
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4200
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2280
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00AFAFAF&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   405
      TabIndex        =   20
      Top             =   4680
      Width           =   6135
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00AFAFAF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   15
         Left            =   0
         ScaleHeight     =   1
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   21
         Top             =   0
         Width           =   15
         Begin VB.Label Liney 
            BackColor       =   &H00000000&
            Height          =   15
            Left            =   960
            TabIndex        =   26
            Top             =   600
            Visible         =   0   'False
            Width           =   1215
         End
      End
   End
   Begin VB.HScrollBar HScroll3 
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   6840
      Width           =   6135
   End
   Begin VB.VScrollBar VScroll3 
      Enabled         =   0   'False
      Height          =   2175
      Left            =   6240
      TabIndex        =   18
      Top             =   4680
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00AFAFAF&
      Caption         =   "Set Broadness"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4200
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll2 
      Enabled         =   0   'False
      Height          =   2775
      Left            =   6240
      TabIndex        =   11
      Top             =   240
      Width           =   255
   End
   Begin VB.HScrollBar HScroll2 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   3000
      Width           =   2895
   End
   Begin VB.VScrollBar VScroll1 
      Enabled         =   0   'False
      Height          =   2775
      Left            =   3000
      TabIndex        =   9
      Top             =   240
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00AFAFAF&
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3360
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00AFAFAF&
      Height          =   2775
      Left            =   3360
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   3
      Top             =   240
      Width           =   2895
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00AFAFAF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   0
         ScaleHeight     =   1
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   5
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00AFAFAF&
      Height          =   2775
      Left            =   120
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   2
      Top             =   240
      Width           =   2895
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00AFAFAF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   0
         ScaleHeight     =   1
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   4
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00AFAFAF&
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Color"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4440
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Background"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Picture"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
On Error GoTo 3
CD1.CancelError = True
CD1.Filter = "Picture Files|*.bmp;*.gif;*.jpg;*.jpeg;*.png"
CD1.ShowOpen
Picture3.Picture = LoadPicture(CD1.FileName)
Picture6.Picture = LoadPicture(CD1.FileName)
Text1.Text = CD1.FileName
HScroll1.Max = -(Picture1.ScaleWidth - (Picture3.ScaleWidth))
VScroll1.Max = -(Picture1.ScaleHeight - (Picture3.ScaleHeight))
HScroll1.LargeChange = HScroll1.Max / (Picture1.Width / Picture3.Width)
VScroll1.LargeChange = VScroll1.Max / (Picture1.Height / Picture3.Height)
HScroll1.SmallChange = Picture3.Width / HScroll1.LargeChange
VScroll1.SmallChange = Picture3.Height / VScroll1.LargeChange
'''
HScroll3.Max = -(Picture5.ScaleWidth - (Picture6.ScaleWidth))
VScroll3.Max = -(Picture5.ScaleHeight - (Picture6.ScaleHeight))
HScroll3.LargeChange = HScroll3.Max / (Picture5.Width / Picture6.Width)
VScroll3.LargeChange = VScroll3.Max / (Picture5.Height / Picture6.Height)
HScroll3.SmallChange = Picture6.Width / HScroll3.LargeChange
VScroll3.SmallChange = Picture6.Height / VScroll3.LargeChange
3 End Sub
Private Sub Command3_Click()
Form2.Show
End Sub
Private Sub Command4_Click()
Dim xi As Long
Dim xU As Long
PB1.Value = 0
PB1.Max = Picture3.ScaleHeight
Liney.Visible = True
Liney.Top = 0
Liney.Left = 0
Liney.Width = Picture3.ScaleWidth
SetControls False
If VScroll3.Enabled = True Then VScroll3.Value = 0
For xi = 1 To Picture3.ScaleHeight
Liney.Top = xi
On Error Resume Next
If VScroll3.Enabled = True Then VScroll3.Value = xi - Picture3.ScaleHeight / 6
For xU = 1 To Picture3.ScaleWidth
If ColorIsClose(Picture3.Point(xU - 1, xi), Label4.BackColor, Form2.Slider1 * 700, Form2.Slider2.Value * 700) = True Then Picture6.Line (xU - 1, xi)-(xU, xi - 2), Picture4.Point(xU, xi - 1): GoTo SkipStep
Picture6.Line (xU - 1, xi)-(xU, xi - 1), Picture3.Point(xU - 1, xi)
SkipStep:
Next xU
PB1.Value = PB1.Value + 1
DoEvents
Next xi
PB1.Value = 0
Liney.Visible = False
DoEvents
If Reg.IsRegistered = False Then Picture6.PaintPicture Picture7.Image, 0, 0, Picture6.Width / 4.5, Picture7.Height / 15
SetControls True
End Sub
Private Sub Command5_Click()
On Error GoTo 3
CD1.CancelError = True
CD1.Filter = "Bitmap Files|*.bmp"
CD1.ShowSave
SavePicture Picture6.Image, CD1.FileName
Text3.Text = CD1.FileName
3 End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
If Button = 1 Then Label4.BackColor = Picture3.Point(x, Y)
End Sub
Private Sub Command1_Click()
On Error GoTo 3
CD1.CancelError = True
CD1.Filter = "Picture Files|*.bmp;*.gif;*.jpg;*.jpeg;*.png"
CD1.ShowOpen
Picture4.Picture = LoadPicture(CD1.FileName)
Text2.Text = CD1.FileName
HScroll2.Max = -(Picture2.ScaleWidth - (Picture4.ScaleWidth))
VScroll2.Max = -(Picture2.ScaleHeight - (Picture4.ScaleHeight))
HScroll2.LargeChange = HScroll2.Max / (Picture2.Width / Picture4.Width)
VScroll2.LargeChange = VScroll2.Max / (Picture2.Height / Picture4.Height)
HScroll2.SmallChange = Picture4.Width / HScroll2.LargeChange
VScroll2.SmallChange = Picture4.Height / VScroll2.LargeChange
3 End Sub
Function ColorIsClose(WhatColor As Long, ToWhatColor As Long, Subtract As Long, Add As Long) As Boolean
On Error Resume Next
If WhatColor - Subtract <= ToWhatColor Then If WhatColor + Add >= ToWhatColor Then ColorIsClose = True
End Function
Private Sub Picture3_Change()
If Picture3.ScaleHeight <= Picture1.ScaleHeight Then VScroll1.Enabled = False
If Picture3.ScaleWidth <= Picture1.ScaleWidth Then HScroll1.Enabled = False
If Picture3.ScaleHeight >= Picture1.ScaleHeight Then VScroll1.Enabled = True
If Picture3.ScaleWidth >= Picture1.ScaleWidth Then HScroll1.Enabled = True
End Sub
Private Sub Picture4_Change()
If Picture4.ScaleHeight <= Picture2.ScaleHeight Then VScroll2.Enabled = False
If Picture4.ScaleWidth <= Picture2.ScaleWidth Then HScroll2.Enabled = False
If Picture4.ScaleHeight >= Picture2.ScaleHeight Then VScroll2.Enabled = True
If Picture4.ScaleWidth >= Picture2.ScaleWidth Then HScroll2.Enabled = True
End Sub
Private Sub Picture6_Change()
If Picture6.ScaleHeight <= Picture5.ScaleHeight Then VScroll3.Enabled = False
If Picture6.ScaleWidth <= Picture5.ScaleWidth Then HScroll3.Enabled = False
If Picture6.ScaleHeight >= Picture5.ScaleHeight Then VScroll3.Enabled = True
If Picture6.ScaleWidth >= Picture5.ScaleWidth Then HScroll3.Enabled = True
End Sub
Private Sub HScroll1_Change()
On Error Resume Next
Picture3.Left = -HScroll1.Value
End Sub
Private Sub HScroll1_Scroll()
On Error Resume Next
Picture3.Left = -HScroll1.Value
End Sub
Private Sub VScroll1_Change()
On Error Resume Next
Picture3.Top = -VScroll1.Value
End Sub
Private Sub VScroll1_Scroll()
On Error Resume Next
Picture3.Top = -VScroll1.Value
End Sub
Private Sub HScroll3_Change()
On Error Resume Next
Picture6.Left = -HScroll3.Value
End Sub
Private Sub HScroll3_Scroll()
On Error Resume Next
Picture6.Left = -HScroll3.Value
End Sub
Private Sub VScroll3_Change()
On Error Resume Next
Picture6.Top = -VScroll3.Value
End Sub
Private Sub VScroll3_Scroll()
On Error Resume Next
Picture6.Top = -VScroll3.Value
End Sub
Private Sub hscroll2_Change()
On Error Resume Next
Picture4.Left = -HScroll2.Value
End Sub
Private Sub hscroll2_Scroll()
On Error Resume Next
Picture4.Left = -HScroll2.Value
End Sub
Private Sub vscroll2_Change()
On Error Resume Next
Picture4.Top = -VScroll2.Value
End Sub
Private Sub vscroll2_Scroll()
On Error Resume Next
Picture4.Top = -VScroll2.Value
End Sub
Function NegNumber(WhatNumber As Long)
NegNumber = -WhatNumber
End Function
Sub SetControls(toWhat As Boolean)
Command1.Enabled = toWhat
Command2.Enabled = toWhat
Command3.Enabled = toWhat
Command4.Enabled = toWhat
Command5.Enabled = toWhat
Text1.Enabled = toWhat
Text2.Enabled = toWhat
Text3.Enabled = toWhat
End Sub
