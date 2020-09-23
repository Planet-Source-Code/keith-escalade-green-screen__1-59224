VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00AFAFAF&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3345
   ClientLeft      =   270
   ClientTop       =   1425
   ClientWidth     =   6975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00AFAFAF&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00AFAFAF&
      Caption         =   "OK"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3120
      Left            =   120
      Picture         =   "frmSplash.frx":0CCE
      ScaleHeight     =   3060
      ScaleWidth      =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   4620
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Code"
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
      Left            =   4800
      TabIndex        =   6
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "www.keithware.com"
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
      Left            =   4800
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Garren Fitzenreiter"
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
      Left            =   4800
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2005"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Registration Code     110+110+98
Private Sub Command1_Click()
If DecryptDDIO2(Text2.Text, 27, "+") = "GSS" Then MsgBox "Registration code accepted.", vbInformation: SaveSetting "GSS", "Registration", "Number", Text2: Reg.IsRegistered = True
Form1.Show
Me.Hide
If Reg.IsRegistered = False Then MsgBox "You have not purchased the full version of Green Screen Shop", vbInformation
Unload Me
End Sub
Private Sub Command2_Click()
Unload Me
End
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim GetReg As String
GetReg = GetSetting("GSS", "Registration", "Number")
If DecryptDDIO2(GetReg, 27, "+") = "GSS" Then Reg.IsRegistered = True: Text2.Visible = False: Label4.Caption = "Registered"
End Sub
