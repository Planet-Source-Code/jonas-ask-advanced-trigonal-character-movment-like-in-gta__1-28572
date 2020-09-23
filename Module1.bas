Attribute VB_Name = "Module1"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const SRCCOPY = &HCC0020


Public Type Character
 X As Currency 'X coordiante of the character
 Y As Currency 'guess
 VSpeed As Currency 'Speed at which the character is moving verticaly
 HSpeed As Currency 'Speed at which the character is moving horizontaly
 Angle As Currency 'The angle the character is facing
 Dire As Byte 'Direction image to paint
End Type

Public P1 As Character 'the character of this sample

Public Sub MoveMan()
    P1.X = P1.X + P1.HSpeed 'move him
    P1.Y = P1.Y + P1.VSpeed
    
    'now, slow down a bit
    P1.HSpeed = P1.HSpeed * 0.1
    P1.VSpeed = P1.VSpeed * 0.1
    
    
End Sub

Public Sub PaintBoard()
    Main.Pic.Cls
    BitBlt Main.Pic.hdc, P1.X, P1.Y, 15, 15, Main.PicManM.hdc, 15, (P1.Dire - 1) * 15, SRCAND
    BitBlt Main.Pic.hdc, P1.X, P1.Y, 15, 15, Main.PicMan.hdc, 15, (P1.Dire - 1) * 15, SRCPAINT
End Sub

Public Sub DoKeys()
Dim Left As Boolean
Dim Down As Boolean
Dim Right As Boolean
Dim Up As Boolean
Const Pi = 3.14159265358979 'define PI
    Left = GetAsyncKeyState(vbKeyLeft)
    Down = GetAsyncKeyState(vbKeyDown)
    Right = GetAsyncKeyState(vbKeyRight)
    Up = GetAsyncKeyState(vbKeyUp)
    
    If Left Then 'turing
        P1.Angle = P1.Angle + 10
    End If
    If Right Then 'turing
        P1.Angle = P1.Angle - 10
    End If
    
    'adjust angles
    If P1.Angle = -10 Then P1.Angle = 350
    If P1.Angle = 360 Then P1.Angle = 0
    
    If Up Then 'going forward
        P1.VSpeed = P1.VSpeed - Sin(P1.Angle / 180 * Pi) * 1.3
        P1.HSpeed = P1.HSpeed + Cos(P1.Angle / 180 * Pi) * 1.3
    End If
    
    If Down Then 'going backward
        P1.VSpeed = P1.VSpeed + Sin(P1.Angle / 180 * Pi) * 1.3
        P1.HSpeed = P1.HSpeed - Cos(P1.Angle / 180 * Pi) * 1.3
    End If
    
    'Now set the right Dire for the angle, so we paint the correct picture
    P1.Dire = (P1.Angle / (10 * 2)) + 1
    For a = 1 To 9 'adjust the dire so that it fitts the images
        P1.Dire = P1.Dire + 1
        If P1.Dire = 19 Then P1.Dire = 1
    Next a
End Sub

