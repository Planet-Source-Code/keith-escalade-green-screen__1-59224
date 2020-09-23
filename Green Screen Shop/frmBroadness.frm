VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00AFAFAF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Broadness"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3855
   Icon            =   "frmBroadness.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00AFAFAF&
      Caption         =   "OK"
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00AFAFAF&
      Caption         =   "Subtract Color Value"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3615
      Begin MSComctlLib.Slider Slider1 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         LargeChange     =   500
         Max             =   8000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00AFAFAF&
      Caption         =   "Add Color Value"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin MSComctlLib.Slider Slider2 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         LargeChange     =   500
         Max             =   8000
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Me.Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Me.Hide
End Sub
