VERSION 5.00
Begin VB.Form frmoptions 
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   2115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll4 
      Height          =   255
      LargeChange     =   3
      Left            =   1980
      Max             =   60
      TabIndex        =   10
      Top             =   1440
      Width           =   2055
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      LargeChange     =   50
      Left            =   1980
      Max             =   30000
      Min             =   10
      TabIndex        =   9
      Top             =   1140
      Value           =   10
      Width           =   2055
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   3
      Left            =   1980
      Max             =   6
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   3
      Left            =   1980
      Max             =   6
      TabIndex        =   5
      Top             =   540
      Width           =   2055
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
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
      Left            =   3480
      TabIndex        =   1
      Top             =   1740
      Width           =   1155
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
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
      Left            =   120
      TabIndex        =   0
      Top             =   1740
      Width           =   1155
   End
   Begin VB.Label lbltime 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "None"
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label lbldiff 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   1140
      Width           =   555
   End
   Begin VB.Label lblheight 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   840
      Width           =   555
   End
   Begin VB.Label lblwidth 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   540
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Time Limit (mins)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1770
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Scramble Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   1140
      Width           =   1635
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Height"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Width"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Options"
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
      Left            =   1800
      TabIndex        =   2
      Top             =   0
      Width           =   1380
   End
End
Attribute VB_Name = "frmoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Temporary variables, there are the ones changed by the
'scroll bars, this allows the user to cancel, getting
'back their original settings
Dim tempwidth As Integer
Dim tempheight As Integer
Dim tempdiff As Integer
Dim temptimelimit As Integer

Private Sub setwidth(sqrwidth As Integer)
  'set up the width from the value on the scroll bar
  If sqrwidth = 0 Then
    tempwidth = 2
  ElseIf sqrwidth = 1 Then tempwidth = 4
  ElseIf sqrwidth = 2 Then tempwidth = 5
  ElseIf sqrwidth = 3 Then tempwidth = 8
  ElseIf sqrwidth = 4 Then tempwidth = 10
  ElseIf sqrwidth = 5 Then tempwidth = 16
  ElseIf sqrwidth = 6 Then tempwidth = 20
  End If
  lblwidth.Caption = tempwidth

End Sub
Private Sub setheight(sqrheight As Integer)
  'set up the height from the value on the scroll bar
  If sqrheight = 0 Then
    tempheight = 2
  ElseIf sqrheight = 1 Then tempheight = 4
  ElseIf sqrheight = 2 Then tempheight = 5
  ElseIf sqrheight = 3 Then tempheight = 8
  ElseIf sqrheight = 4 Then tempheight = 10
  ElseIf sqrheight = 5 Then tempheight = 16
  ElseIf sqrheight = 6 Then tempheight = 20
  End If
  lblheight.Caption = tempheight
End Sub




Private Sub cmdcancel_Click()
  'since the cancel the diff, tileheight etc have not been changed
  'only the temp ones which are not saved
  'close the window
  frmoptions.Visible = False
  frmmain.Enabled = True
  frmmain.SetFocus
  Unload frmoptions
End Sub

Private Sub cmdsave_Click()
  'now they choose to save make the probwer variables equal
  'to the new values
  tilewidth = tempwidth
  tileheight = tempheight
  diff = tempdiff
  timelimit = temptimelimit
  'save information to file
  Open App.Path & "\gameinfo.dat" For Output As #1
    Write #1, tilewidth;
    Write #1, tileheight;
    Write #1, diff;
    Write #1, timelimit
  Close #1
  'close this window
  frmmain.Enabled = True
  frmoptions.Visible = False
  frmmain.SetFocus
  Unload frmoptions
End Sub

Private Sub Form_Load()
  'centre form in main form
  frmoptions.Left = frmmain.Left + ((frmmain.Width - frmoptions.Width) / 2)
  frmoptions.Top = frmmain.Top + ((frmmain.Height - frmoptions.Height) / 2)
  'set all the scroll bars to their correct positions
  'the label boxes are done automatically by the change event
  If tilewidth = 2 Then HScroll1.Value = 0
  If tilewidth = 4 Then HScroll1.Value = 1
  If tilewidth = 5 Then HScroll1.Value = 2
  If tilewidth = 8 Then HScroll1.Value = 3
  If tilewidth = 10 Then HScroll1.Value = 4
  If tilewidth = 16 Then HScroll1.Value = 5
  If tilewidth = 20 Then HScroll1.Value = 6
  If tileheight = 2 Then HScroll2.Value = 0
  If tileheight = 4 Then HScroll2.Value = 1
  If tileheight = 5 Then HScroll2.Value = 2
  If tileheight = 8 Then HScroll2.Value = 3
  If tileheight = 10 Then HScroll2.Value = 4
  If tileheight = 16 Then HScroll2.Value = 5
  If tileheight = 20 Then HScroll2.Value = 6
  HScroll3.Value = diff
  HScroll4.Value = (timelimit / 60) 'scroll bar uses minutes
  tempwidth = tilewidth
  tempheight = tileheight
  tempdiff = diff
  temptimelimit = timelimit
End Sub


Private Sub HScroll1_Change()
  Call setwidth(HScroll1.Value) 'call setwidth with the scrollbar value
End Sub

Private Sub HScroll1_Scroll()
  Call setwidth(HScroll1.Value)
End Sub

Private Sub HScroll2_Change()
  Call setheight(HScroll2.Value) 'call setheight with the scroll bar value
End Sub

Private Sub HScroll2_Scroll()
  Call setheight(HScroll2.Value)
End Sub

Private Sub HScroll3_Change()
  tempdiff = HScroll3.Value 'set the difficulty (scramble level)
  lbldiff.Caption = tempdiff
End Sub

Private Sub HScroll3_Scroll()
  tempdiff = HScroll3.Value
  lbldiff.Caption = tempdiff
End Sub

Private Sub HScroll4_Change()
  temptimelimit = (HScroll4.Value * 60) 'set the timelimit (converts it to seconds for the variable
  If temptimelimit = 0 Then
    lbltime.Caption = "None"
  Else
    lbltime.Caption = HScroll4.Value
  End If
End Sub

Private Sub HScroll4_Scroll()
  temptimelimit = (HScroll4.Value * 60)
  If temptimelimit = 0 Then
    lbltime.Caption = "None"
  Else
    lbltime.Caption = HScroll4.Value
  End If
End Sub
