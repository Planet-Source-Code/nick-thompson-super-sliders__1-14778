VERSION 5.00
Begin VB.Form frmcontrols 
   BorderStyle     =   0  'None
   Caption         =   "Controls"
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4755
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   1620
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   4020
      Picture         =   "frmcontrols.frx":0000
      ScaleHeight     =   690
      ScaleWidth      =   690
      TabIndex        =   2
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label2 
      Caption         =   $"frmcontrols.frx":04D2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   3915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2100
      TabIndex        =   0
      Top             =   0
      Width           =   840
   End
End
Attribute VB_Name = "frmcontrols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
 'close window
  frmcontrols.Visible = False
  frmmain.Enabled = True
  frmmain.SetFocus
  Unload frmcontrols

End Sub

Private Sub Form_Load()
  'centre form in main form
  frmcontrols.Left = frmmain.Left + ((frmmain.Width - frmcontrols.Width) / 2)
  frmcontrols.Top = frmmain.Top + ((frmmain.Height - frmcontrols.Height) / 2)

End Sub
