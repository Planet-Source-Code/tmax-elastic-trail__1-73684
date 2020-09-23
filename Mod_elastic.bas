Attribute VB_Name = "Mod_elastic"
Option Explicit
' Elastic Trail with Png Images
' Vb version from tmax_visiber@yahoo.com

' LayeredWindow and FrmPng from "PngMania" by Agustin Rodriguez
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=63880&lngWId=1

' Autoit version
' http://www.autoitscript.com/forum/topic/89544-elastic-images-with-a-happy-valentines-example/
' Original JavaScript by Philip Winston - pwinston@yahoo.com
' Adapted for AutoIT by AndyBiochem,used with GDI+ for PNG's by yehia

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Type Ball
  x As Long
  y As Long
  XVel As Single
  YVel As Single
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Const iBalls = 8            'Number of balls
Const iDeltaT = 0.01        'fixed time step, no relation to real time
Const iSegLen = 20          'size of one spring in pixels
Const iSpringK = 12         'spring constant, stiffness of springs
Const iMass = 1             'Object mass
Const iXGravity = 0         'Positive XGRAVITY pulls right, negative pulls left
Const iYGravity = 50        'Positive YGRAVITY pulls down, negative up
Const iRes = 10             'RESISTANCE determines a slowing force proportional to velocity
Const iStopVel = 0.1        'stopping criterea to prevent endless jittering doesn't work when sitting on bottom since floor
Const iStopAcc = 0.1        'doesn't push back so acceleration always as big as gravity
Const iBounce = 0.75        'BOUNCE is percent of velocity retained when bouncing off a wall
Const iBallsize = 100

Dim MouseObj(0 To iBalls) As Ball
Dim Spring As POINTAPI
Dim Resist As POINTAPI
Dim Accel As POINTAPI
Dim cursorPoint As POINTAPI
Dim iHeight As Integer
Dim iWidth As Integer
Dim iXpos As Long
Dim iYpos As Long
Dim f(1 To iBalls) As New FrmPng

Function Animate()
    Dim i%
    'MouseObj(0) follows the mouse, though nothing is drawn there
    MouseObj(0).x = iXpos
    MouseObj(0).y = iYpos
    For i% = 1 To (iBalls - 1)
      Spring.x = 0
      Spring.y = 0
      Spring_Force i% - 1, i%
      If i% < (iBalls - 1) Then Spring_Force i% + 1, i%
      'air resisitance/friction
      Resist.x = -MouseObj(i%).XVel * iRes
      Resist.y = -MouseObj(i%).YVel * iRes
      'compute new accel, including gravity
      Accel.x = (Spring.x + Resist.x) / iMass + iXGravity
      Accel.y = (Spring.y + Resist.y) / iMass + iYGravity
      'compute new velocity
      MouseObj(i%).XVel = MouseObj(i%).XVel + (iDeltaT * Accel.x)
      MouseObj(i%).YVel = MouseObj(i%).YVel + (iDeltaT * Accel.y)
      'stop dead so it doesn't jitter when nearly still
      If Abs(MouseObj(i%).XVel) < iStopVel And Abs(MouseObj(i%).YVel) < iStopVel And Abs(Accel.x) < iStopAcc And Abs(Accel.y) < iStopAcc Then
        MouseObj(i%).XVel = 0
        MouseObj(i%).YVel = 0
      End If
      ' move to new position
      MouseObj(i%).x = MouseObj(i%).x + MouseObj(i%).XVel + f(i%).ScaleWidth / 15 / Screen.TwipsPerPixelX
      MouseObj(i%).y = MouseObj(i%).y + MouseObj(i%).YVel
      'bounce off 3 walls (leave ceiling open)
      If (MouseObj(i%).y < 0) Then
        If (MouseObj(i%).YVel < 0) Then MouseObj(i%).YVel = iBounce * -MouseObj(i%).YVel
        MouseObj(i%).y = 0
      End If
      If (MouseObj(i%).y >= iHeight - iBallsize - 1) Then
        If (MouseObj(i%).YVel > 0) Then MouseObj(i%).YVel = iBounce * -MouseObj(i%).YVel
        MouseObj(i%).y = iHeight - iBallsize - 1
      End If
      If (MouseObj(i%).x >= iWidth - iBallsize) Then
        If (MouseObj(i%).XVel > 0) Then MouseObj(i%).XVel = iBounce * -MouseObj(i%).XVel
        MouseObj(i%).x = iWidth - iBallsize - 1
      End If
      If (MouseObj(i%).x < 0) Then
        If (MouseObj(i%).XVel < 0) Then MouseObj(i%).XVel = iBounce * -MouseObj(i%).XVel
        MouseObj(i%).x = 0
      End If
      'move img to new position
      f(i%).Left = MouseObj(i%).x * Screen.TwipsPerPixelX
      f(i%).Top = MouseObj(i%).y * Screen.TwipsPerPixelY
    Next
End Function

'Adds force in X and Y to spring for MouseObj(i%) on MouseObj(j%)
Function Spring_Force(i%, j%)
  On Error Resume Next
  Dim Dx As Long, Dy As Long
  Dim Leni As Long
  Dim springF%
  Dx = MouseObj(i%).x - MouseObj(j%).x
  Dy = MouseObj(i%).y - MouseObj(j%).y
  Leni = Sqr(Dx * Dx + Dy * Dy)
  If Leni > iSegLen Then
    springF = iSpringK * (Leni - iSegLen)
    Spring.x = Spring.x + (Dx / Leni) * springF
    Spring.y = Spring.y + (Dy / Leni) * springF
  End If
End Function

Private Sub Main()
  Dim i%
  iWidth = Screen.Width
  iHeight = Screen.Height
  For i% = 1 To iBalls
    f(i%).Show
    f(i%).LoadPng (i%)
  Next
  Randomize
  Do While DoEvents()
    Sleep (Rnd(10) * 10 + 20)
    'Get mouse position for animation
    GetCursorPos cursorPoint
    iXpos = cursorPoint.x
    iYpos = cursorPoint.y
    If GetAsyncKeyState(16) And GetAsyncKeyState(27) Then UnloadFrm
    Animate
  Loop
End Sub

Public Sub UnloadFrm()
  Dim i%
  For i% = 1 To iBalls
    Unload f(i%)
  Next
  End
End Sub

