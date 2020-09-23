VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   0  'None
   Caption         =   "About Super Sliders"
   ClientHeight    =   2025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4755
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
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
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "TheVBGod@Hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   60
      TabIndex        =   3
      Top             =   1260
      Width           =   4665
   End
   Begin VB.Label Label2 
      Caption         =   $"frmabout.frx":0000
      Height          =   915
      Left            =   60
      TabIndex        =   2
      Top             =   420
      Width           =   4695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "About Super Sliders"
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
      Left            =   660
      TabIndex        =   0
      Top             =   0
      Width           =   3525
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdok_Click()
  'close window
  frmabout.Visible = False
  frmmain.Enabled = True
  frmmain.SetFocus
  Unload frmabout
End Sub

Private Sub Form_Load()
  'centre form in main form
  frmabout.Left = frmmain.Left + ((frmmain.Width - frmabout.Width) / 2)
  frmabout.Top = frmmain.Top + ((frmmain.Height - frmabout.Height) / 2)
End Sub
