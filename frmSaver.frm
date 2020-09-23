VERSION 5.00
Begin VB.Form frmSaver 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Water Screen Saver"
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   DrawWidth       =   2
   ForeColor       =   &H00FF0000&
   Icon            =   "frmSaver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmSaver.frx":0442
   MousePointer    =   99  'Custom
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   359
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      Picture         =   "frmSaver.frx":074C
      ScaleHeight     =   24
      ScaleMode       =   2  'Point
      ScaleWidth      =   51.75
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox picFish 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      Picture         =   "frmSaver.frx":218E
      ScaleHeight     =   24
      ScaleMode       =   2  'Point
      ScaleWidth      =   51.75
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picWater 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   7500
      Left            =   1560
      Picture         =   "frmSaver.frx":3BD0
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   820
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   12300
   End
End
Attribute VB_Name = "frmSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WaterFlow As Integer
Dim WaterHeight As Single
Dim i As Integer
Dim RainInterval As Integer
Const YVel = 4
Dim X_old As Integer, Y_old As Integer
Dim HaveValues As Boolean
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Unload Me: End
End Sub

'I cheated to make the icon disappear (made it a transparent .ico file.
Private Sub Form_Load()
Dim XVel As Integer
Randomize

WaterFlow = 0
WaterHeight = 75    'Start a little from the bottom
RainInterval = 0
HaveValues = False
For i = 1 To 300    'Set random positions for the rain
    RainDrop(i).XVal = Rnd * 800
    RainDrop(i).YVal = Rnd * 525
Next i
For i = 1 To 5      'Make five random fish.
    Fishy(i).XVal = Rnd * 800
    Fishy(i).YVal = 550                     'Must start at bottom of screen
    Fishy(i).DescentSpeed = Rnd * 6 - 3     'Number between -3 and 3
    Fishy(i).XSpeed = Rnd * 6 - 3
Next i

'Get picture of desktop:
GetWindowRect GetDesktopWindow(), DispRec ' Get DeskTop Rectangle dimentions
InitDeskDC DeskDC, DeskBmp, DispRec ' Initialize desktop image information.

frmSaver.Show
Do
    WaterFlow = WaterFlow + 1   'Move the water to the right
    
    If WaterFlow >= 20 Then WaterFlow = 0
    '...The water doesn't keep moving right. It moves right the length of a wave(20)
    'then returns to the original posistion. This saves wrapping the image, and
    'is why picWater is 820 width (instead of 20)
    
    If WaterHeight < 500 Then WaterHeight = WaterHeight + 0.1      'Water rises (but not to the top)
    RainInterval = RainInterval + 1
    If RainInterval > 6280 Then RainInterval = 0    'Otherwise it will crash due to overflow at 32000
    
    XVel = Sin(RainInterval / 100) * 5
    'Sine means that it will go strongly one way then reverse and go the other way. (ie follow sine curve).
   
    Call BitBlt(frmSaver.hDC, 0, 0, 800, 600, DeskDC, 0, 0, SRCCOPY)
    'Re-paint the desktop screen

    For i = 1 To 300 'Move the rain
        RainDrop(i).XVal = RainDrop(i).XVal + XVel
        RainDrop(i).YVal = RainDrop(i).YVal + YVel
        If RainDrop(i).XVal > 800 Then RainDrop(i).XVal = 0
        If RainDrop(i).XVal < 0 Then RainDrop(i).XVal = 800
        If RainDrop(i).YVal > (600 - WaterHeight) Then RainDrop(i).YVal = 0
        ''Faster - but only 1 pixel (to thin): Call SetPixel(frmSaver.hDC, RainDrop(I).XVal, RainDrop(I).YVal, RGB(150, 150, 250))
        frmSaver.PSet (RainDrop(i).XVal, RainDrop(i).YVal)
    Next i
    'Paint Water (fishes should be on to of water (looks better)
    Call BitBlt(frmSaver.hDC, 0, 600 - WaterHeight, 800, WaterHeight, picWater.hDC, WaterFlow, 0, SRCAND)

    'Move the little fishies
    For i = 1 To 5
        Fishy(i).XVal = Fishy(i).XVal + Fishy(i).XSpeed
        Fishy(i).YVal = Fishy(i).YVal + Fishy(i).DescentSpeed
        If Rnd * 40 < 1 Then Fishy(i).DescentSpeed = Rnd * 6 - 3    'Occasionally change the descent speed
        If Rnd * 40 < 1 Then Fishy(i).XSpeed = Rnd * 6 - 3    'Occasionally change the horizontal speed
        If Fishy(i).XVal > 730 Then Fishy(i).XSpeed = Rnd * -3
        If Fishy(i).XVal < 0 Then Fishy(i).XSpeed = Rnd * 3
        If Fishy(i).YVal < (605 - WaterHeight) Then Fishy(i).DescentSpeed = Rnd * 3
        If Fishy(i).YVal > 570 Then Fishy(i).DescentSpeed = Rnd * -3
        'Draw it depending on what direction the fish is moving. BitBlt draws normally
        'StretchBlt draws it with negative width therefore reversed.
        If Fishy(i).XSpeed < 0 Then
            Call StretchBlt(frmSaver.hDC, Fishy(i).XVal + 69, Fishy(i).YVal, -69, 32, picMask.hDC, 0, 0, 69, 32, SRCPAINT)
            Call StretchBlt(frmSaver.hDC, Fishy(i).XVal + 69, Fishy(i).YVal, -69, 32, picFish.hDC, 0, 0, 69, 32, SRCAND)
        Else
            Call BitBlt(frmSaver.hDC, Fishy(i).XVal, Fishy(i).YVal, 69, 32, picMask.hDC, 0, 0, SRCPAINT)
            Call BitBlt(frmSaver.hDC, Fishy(i).XVal, Fishy(i).YVal, 69, 32, picFish.hDC, 0, 0, SRCAND)
        End If
    Next i
    frmSaver.Refresh    'Paint it all on the screen
    DoEvents
Loop

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This code means that moving the mouse slowly won't cause it to end.
    If HaveValues Then
        If Abs(X - X_old) > 5 Or Abs(Y - Y_old) > 5 Then    'Only if we move it a lot - not just bump the mouse.
            Unload Me: End
        End If
    End If
    X_old = X
    Y_old = Y
    HaveValues = True
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me: End
End Sub
