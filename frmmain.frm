VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Super Sliders"
   ClientHeight    =   7425
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9825
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timerstart 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   60
      Top             =   60
   End
   Begin VB.PictureBox picstore 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00808080&
      Height          =   7200
      Left            =   9060
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   5
      Top             =   6600
      Visible         =   0   'False
      Width           =   9600
   End
   Begin VB.PictureBox picsource 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   7500
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   540
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.PictureBox pictest 
      Height          =   675
      Left            =   8820
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.FileListBox fileload 
      Height          =   7110
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   5955
   End
   Begin VB.PictureBox picback 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00808080&
      Height          =   7200
      Left            =   120
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   9600
      Begin VB.Timer timertime 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   6840
         Top             =   180
      End
   End
   Begin VB.Label lblinfo 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   6180
      TabIndex        =   2
      Top             =   3300
      Width           =   3435
   End
   Begin VB.Image imgpreview 
      Appearance      =   0  'Flat
      Height          =   2955
      Left            =   6180
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Menu menugame 
      Caption         =   "&Game"
      Begin VB.Menu menunew 
         Caption         =   "&New Game"
      End
      Begin VB.Menu menuload 
         Caption         =   "&Load Picture"
      End
      Begin VB.Menu menuscramble 
         Caption         =   "&Scramble"
      End
      Begin VB.Menu menusolve 
         Caption         =   "Sol&ve"
      End
      Begin VB.Menu menuoptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu menuexit 
         Caption         =   "E&xit Game"
      End
   End
   Begin VB.Menu menuhelp 
      Caption         =   "&Help"
      Begin VB.Menu menucontrols 
         Caption         =   "&Help"
      End
      Begin VB.Menu menuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub setupgame()
  ReDim tileinfo(tilewidth - 1, tileheight - 1) 'redim for correct size
  Dim cpos As Integer
  cpos = 1
  For i = 0 To tileheight - 1 'start off with correct order
    For j = 0 To tilewidth - 1
      tileinfo(j, i) = cpos
      cpos = cpos + 1
    Next j
  Next i
  timeleft = timelimit 'timeleft = the timelimit
  If timelimit = 0 Then 'if zero time is infinate (shown by timeleft of -1)
    timeleft = -1
  End If
  Call scramble 'scramble the tiles up
End Sub

Private Sub scramble()
Dim starttilei As Integer 'stores position in tileinfo(this one,not this) of selected tile
Dim starttilej As Integer 'stores position in tileinfo(not this,this one) of selected tile
Dim starttile As Integer 'the value of tileinfo(starttilei,starttilej)
Dim endtilegval As Byte 'whether to swap it with the one up, down etc)
Dim endtilei As Integer 'same as above but for end tile
Dim endtilej As Integer
Dim endtile As Integer
  
For k = 1 To diff 'loop for number of selected scrambles (10 to 30000)
  
  Randomize   ' Initialize random-number generator.
  'get starttile to change
  starttilei = Int((tilewidth) * Rnd)       ' Generate random value between 0 and tilewidth - 1.
  starttilej = Int((tileheight) * Rnd)       ' Generate random value between 0 and tileheight - 1.
  starttile = tileinfo(starttilei, starttilej)
  'get endtile to change
  Dim gotval As Byte
  gotval = 0
  Do
  endtilegval = Int((4 * Rnd) + 1)    ' Generate random value between 1 and 4
    If endtilegval = 1 Then 'swap it with tile to the right
      If starttilei < tilewidth - 1 Then
        endtilei = starttilei + 1
        endtilej = starttilej
        gotval = 1
      End If
    ElseIf endtilegval = 2 Then 'swap it with tile below
      If starttilej < tileheight - 1 Then
        endtilei = starttilei
        endtilej = starttilej + 1
        gotval = 1
      End If
    ElseIf endtilegval = 3 Then 'swap it with tile to left
      If starttilei > 0 Then
        endtilei = starttilei - 1
        endtilej = starttilej
        gotval = 1
      End If
    ElseIf endtilegval = 4 Then 'swap it with tile above
      If starttilej > 0 Then
        endtilei = starttilei
        endtilej = starttilej - 1
        gotval = 1
      End If
    End If
  Loop Until gotval = 1 'go back if this tile does not exist
  endtile = tileinfo(endtilei, endtilej) 'swap tiles
  tileinfo(starttilei, starttilej) = endtile
  tileinfo(endtilei, endtilej) = starttile
Next k
Dim cpos As Integer
cpos = 1
For i = 0 To tileheight - 1 'check it is not the same as it was at beginning. this is possible e.g. in a 2x2 square
  For j = 0 To tilewidth - 1
    If tileinfo(j, i) <> cpos Then
      Exit For
    End If
    If cpos = tileheight * tilewidth Then
      Call setupgame
      Exit Sub
    End If
    cpos = cpos + 1
  Next j
Next i
timerstart.Enabled = True 'start the starttimer
timertime.Enabled = True 'start the time left timer
Call drawtiles 'draw tiles on screen

End Sub

Private Sub drawtiles()
'this commented out code is for checking it has created the
'order correctly
'  Open App.Path & "\test.dat" For Output As #1
'  For i = 0 To tileheight - 1
'    For j = 0 To tilewidth - 1
'      If j = tilewidth - 1 Then
'      Write #1, tileinfo(j, i)
'
'      Else
'      Write #1, tileinfo(j, i);
'     End If
'    Next j
'  Next i
'  Close #1
  Dim numdown As Integer 'number of tiles down on storeage picture to bitblt from
  Dim numacross As Integer 'see above but across
  For i = 0 To tileheight - 1 'go through each tile
    For j = 0 To tilewidth - 1
      numdown = ((tileinfo(j, i) - 1) \ tilewidth) 'using values in tileinfo decide which tile it is
      numacross = tileinfo(j, i)
      If numacross > tilewidth Then
        Do
          numacross = numacross - tilewidth
        Loop Until numacross <= tilewidth
      End If
      numacross = numacross - 1
      success = BitBlt(picback.hdc, j * (640 / tilewidth), i * (480 / tileheight), 640 / tilewidth, 480 / tileheight, picstore.hdc, (640 / tilewidth) * numacross, (480 / tileheight) * numdown, SRCCOPY) 'copy that tile
    Next j
  Next i
  For i = 0 To tilewidth 'draw on lines
    picback.Line (i * (640 / tilewidth), 0)-(i * (640 / tilewidth), 480)
  Next i
  For i = 0 To tileheight
    picback.Line (0, i * (480 / tileheight))-(640, i * (480 / tileheight))
  Next i
  picback.Refresh 'refresh (autoredraw is on)
  If running = 0 Then Exit Sub 'this will exit if they have clicked solve
  Dim cpos As Integer 'check to see if they have completed it
  cpos = 1
  For i = 0 To tileheight - 1
    For j = 0 To tilewidth - 1
      If tileinfo(j, i) <> cpos Then
        Exit For
      End If
      If cpos = tileheight * tilewidth Then
        picback.Enabled = False
        running = 0
        timertime.Enabled = False
        MsgBox "Congratulations. You have Won.", vbOKOnly
        Exit Sub
      End If
      cpos = cpos + 1
    Next j
  Next i

End Sub

Private Sub fileload_Click()
  lblinfo.Caption = "" 'clear info box
  On Error GoTo errorinload 'if not supported or corrupted goto errorinload
    If fileload.filename <> "" Then 'if a file is selected
      imgpreview.Picture = LoadPicture(fileload.Path & "\" & fileload.filename) 'load it into image for preview
      pictest.Picture = LoadPicture(fileload.Path & "\" & fileload.filename) 'load it in to picture box for compatibility test
    End If
  Exit Sub
errorinload:
  lblinfo.Caption = "Not a valid picture or it is corrupted"
  imgpreview.Picture = LoadPicture("")
End Sub

Private Sub fileload_DblClick()
  On Error GoTo errorinmainload
  picsource.Picture = LoadPicture(fileload.Path & "\" & fileload.filename) 'load file into the source picture
  fileload.Visible = False
  imgpreview.Visible = False
  lblinfo.Visible = False
  success = StretchBlt(picback.hdc, 0, 0, 640, 480, picsource.hdc, 0, 0, picsource.Width / 15, picsource.Height / 15, SRCCOPY) 'stretch it into 2 pictures, one visible, one storage
  success = StretchBlt(picstore.hdc, 0, 0, 640, 480, picsource.hdc, 0, 0, picsource.Width / 15, picsource.Height / 15, SRCCOPY)
  picback.Enabled = True 'make all the stuff visible
  picback.Visible = True
  menunew.Enabled = True
  menusolve.Enabled = True
  menuscramble.Enabled = True
  Call setupgame 'set it up
  Exit Sub

errorinmainload:
  lblinfo.Caption = "Not a valid picture or it is corrupted"
  imgpreview.Picture = LoadPicture("")
End Sub

Private Sub Form_Load()
  'load information from file and assign to variables
  Open App.Path & "\gameinfo.dat" For Input As #1
    Input #1, tilewidth
    Input #1, tileheight
    Input #1, diff
    Input #1, timelimit
  Close #1
  timeleft = 0
  running = 0 'indicates no game is running
  'start of game, so no picture loaded, therefore they cant start (new)
  'auto sole or scramble.
  menunew.Enabled = False
  menusolve.Enabled = False
  menuscramble.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End 'end game when cross is pressed
End Sub

Private Sub menuabout_Click()
  'This code move the main form if it is in a position where the
  'about form will not be visible
  If frmmain.Top + (frmmain.Height / 2) + 1500 > Screen.Height Then
    frmmain.Top = Screen.Height - ((frmmain.Height / 2) + 1500)
  End If
  If frmmain.Left < 0 Then frmmain.Left = 0
  If frmmain.Left + (frmmain.Width / 2) + 2500 > Screen.Width Then
    frmmain.Left = Screen.Width - ((frmmain.Width / 2) + 2500)
  End If
  'make about form visible
  frmabout.Visible = True
  frmmain.Enabled = False
End Sub

Private Sub menucontrols_Click()
  'This code move the main form if it is in a position where the
  'about form will not be visible
  If frmmain.Top + (frmmain.Height / 2) + 1500 > Screen.Height Then
    frmmain.Top = Screen.Height - ((frmmain.Height / 2) + 1500)
  End If
  If frmmain.Left < 0 Then frmmain.Left = 0
  If frmmain.Left + (frmmain.Width / 2) + 2500 > Screen.Width Then
    frmmain.Left = Screen.Width - ((frmmain.Width / 2) + 2500)
  End If
  'make about form visible
  frmcontrols.Visible = True
  frmmain.Enabled = False

End Sub

Private Sub menuexit_Click()
  If running = 1 Then 'if game is running check this is what the really want to do
    res = MsgBox("Are you sure, this will end the current game.", vbQuestion Or vbYesNo)
    If res = vbNo Then Exit Sub
  End If
  End 'exit game
End Sub

Private Sub menuload_Click()
  If running = 1 Then 'if game is running check this is what they really want to do
    res = MsgBox("Are you sure, this will end the current game.", vbQuestion Or vbYesNo)
    If res = vbNo Then Exit Sub
  End If
  picback.Visible = False
  running = 0  'indicates end of game
  fileload.Path = App.Path & "\pics\"
  fileload.Visible = True
  imgpreview.Picture = LoadPicture("")
  imgpreview.Visible = True
  timertime.Enabled = False
  timerstart.Enabled = False
  menusolve.Enabled = False
  menuscramble.Enabled = False
  menunew.Enabled = False
  frmmain.Caption = "Super Sliders"
End Sub

Private Sub menunew_Click()
  If running = 1 Then 'if game is running check this is what the really want to do
    res = MsgBox("Are you sure, this will end the current game.", vbQuestion Or vbYesNo)
    If res = vbNo Then Exit Sub
  End If
  picback.Enabled = False
  running = 0
  menuscramble.Enabled = True
  menusolve.Enabled = True
  Call setupgame
  picback.Enabled = True
End Sub

Private Sub menuoptions_Click()
  If running = 1 Then 'if game is running check this is what they really want to do
    res = MsgBox("Are you sure, this will end the current game.", vbQuestion Or vbYesNo)
    If res = vbNo Then Exit Sub
  End If
  picback.Visible = False
  timertime.Enabled = False
  menunew.Enabled = False
  menusolve.Enabled = False
  menuscramble.Enabled = False
  running = 0
  frmmain.Caption = "Super Sliders"
  'This code move the main form if it is in a position where the
  'options form will not be visible
  If frmmain.Top + (frmmain.Height / 2) + 1500 > Screen.Height Then
    frmmain.Top = Screen.Height - ((frmmain.Height / 2) + 1500)
  End If
  If frmmain.Left < 0 Then frmmain.Left = 0
  If frmmain.Left + (frmmain.Width / 2) + 2500 > Screen.Width Then
    frmmain.Left = Screen.Width - ((frmmain.Width / 2) + 2500)
  End If
  'make the options form visible
  frmoptions.Visible = True
  frmmain.Enabled = False
End Sub

Private Sub menuscramble_Click()
  If running = 1 Then 'if game is running check this is what the really want to do
    res = MsgBox("Are you sure, this will scamble the image from its current position.", vbQuestion Or vbYesNo)
    If res = vbNo Then Exit Sub
  End If
  running = 0
  picback.Enabled = False
  Call scramble
  picback.Enabled = True
End Sub

Private Sub menusolve_Click()
  If running = 1 Then 'if game is running check this is what the really want to do
    res = MsgBox("Are you sure, this will end the current game.", vbQuestion Or vbYesNo)
    If res = vbNo Then Exit Sub
  End If
  Dim cpos As Integer
  cpos = 1
  For i = 0 To tileheight - 1
    For j = 0 To tilewidth - 1
      tileinfo(j, i) = cpos
      cpos = cpos + 1
    Next j
  Next i
  running = 0
  timertime.Enabled = False
  frmmain.Caption = "Super Sliders"
  timeleft = 0
  menuscramble.Enabled = False
  menusolve.Enabled = False
  Call drawtiles
  
End Sub

Private Sub picback_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If running = 0 Then Exit Sub
Dim starti As Integer
Dim startj As Integer
Dim inboxposi As Integer
Dim inboxposj As Integer
Dim starttile As Integer
Dim endtile As Integer
Dim endi As Integer
Dim endj As Integer
Dim tempnum As Integer
endi = 0
endj = 0
starti = X \ (640 / tilewidth)
startj = Y \ (480 / tileheight)
starttile = tileinfo(starti, startj)
inboxposi = X - ((starti) * (640 / tilewidth))
inboxposj = Y - ((startj) * (480 / tileheight))

If inboxposj < (((480 / tileheight) * inboxposi) / (640 / tilewidth)) And inboxposj < (480 / tileheight) - (((480 / tileheight) * inboxposi) / (640 / tilewidth)) Then
  If startj = 0 Then Exit Sub
  tempnum = tileinfo(starti, startj)
  tileinfo(starti, startj) = tileinfo(starti, startj - 1)
  tileinfo(starti, startj - 1) = tempnum
ElseIf inboxposj < (((480 / tileheight) * inboxposi) / (640 / tilewidth)) And inboxposj > (480 / tileheight) - (((480 / tileheight) * inboxposi) / (640 / tilewidth)) Then
  If starti = tilewidth - 1 Then Exit Sub
  tempnum = tileinfo(starti, startj)
  tileinfo(starti, startj) = tileinfo(starti + 1, startj)
  tileinfo(starti + 1, startj) = tempnum
ElseIf inboxposj > (((480 / tileheight) * inboxposi) / (640 / tilewidth)) And inboxposj > (480 / tileheight) - (((480 / tileheight) * inboxposi) / (640 / tilewidth)) Then
  If startj = tileheight - 1 Then Exit Sub
  tempnum = tileinfo(starti, startj)
  tileinfo(starti, startj) = tileinfo(starti, startj + 1)
  tileinfo(starti, startj + 1) = tempnum
ElseIf inboxposj > (((480 / tileheight) * inboxposi) / (640 / tilewidth)) And inboxposj < (480 / tileheight) - (((480 / tileheight) * inboxposi) / (640 / tilewidth)) Then
  If starti = 0 Then Exit Sub
  tempnum = tileinfo(starti, startj)
  tileinfo(starti, startj) = tileinfo(starti - 1, startj)
  tileinfo(starti - 1, startj) = tempnum
End If
Call drawtiles
End Sub

Private Sub timerstart_Timer()
  running = 1
  timerstart.Enabled = False
End Sub

Private Sub timertime_Timer()
  If frmmain.Enabled = False Then Exit Sub
  If timeleft < 0 Then
    timertime.Enabled = False
    Exit Sub
  End If
  If timeleft = 0 Then
    running = 0
    timertime.Enabled = False
    MsgBox "You have run out of time.", vbOKOnly
  End If
  timeleft = timeleft - 1
  If timeleft >= 0 Then frmmain.Caption = "Super Sliders " & timeleft & " Seconds Left"
End Sub
