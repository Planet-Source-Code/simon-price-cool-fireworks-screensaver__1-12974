VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFirework"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Const GRAVITY = 20
Private Const PI = 3.1415
Private Const PIby2 = PI * 2
Private Const PIdiv180 = PI / 180

Private Type tParticle
    x As Single
    y As Single
    xm As Single
    ym As Single
    colorref As Byte
End Type

Private Particle() As tParticle
Private ExplosionTime As Long
Private Exploded As Boolean
Private EndTime As Long
Private Color1 As Long
Private Color2 As Long
Private Brightness As Single
Private ExplosionPower As Single

Sub Fire(x As Single, y As Single, Angle As Single, Speed As Single, ExplodePower As Single, NumParticles As Long, ExplodeTime As Long, FadeTime As Long)
On Error Resume Next
Dim l As Long
NumParticles = NumParticles - 1
ReDim Particle(0 To NumParticles)
For l = 0 To NumParticles
    With Particle(l)
        .colorref = Int(Rnd * 2) + 1
        .x = x + Rnd * 2 - 1
        .y = y + Rnd * 2 - 1
        .xm = Sin(Angle) * Speed
        .ym = Cos(Angle) * Speed
    End With
Next
Dim r As Byte, g As Byte, b As Byte
r = Rnd * 35 + 220
g = Rnd * 155 + 100
b = Rnd * 155 + 100
Color1 = RGB(r, g, b)
r = Rnd * 35 + 220
g = Rnd * 155 + 100
b = Rnd * 155 + 100
Color2 = RGB(r, g, b)
ExplosionTime = timeGetTime() + ExplodeTime
EndTime = ExplosionTime + FadeTime
ExplosionPower = ExplodePower
Exploded = False
Brightness = 1
l = Int(Rnd * 3) + 1
RocketSound(l).Play DSBPLAY_DEFAULT
End Sub

Sub Move(TimeElapsed As Long)
Dim l As Long
Dim t As Single
t = TimeElapsed / 1000
On Error Resume Next
For l = 0 To UBound(Particle)
    With Particle(l)
        .x = .x + .xm * t
        .ym = .ym + GRAVITY * t
        .y = .y + .ym * t
    End With
Next
Dim Now As Long
Now = timeGetTime()
If Exploded Then
    Brightness = 1 - (Now - ExplosionTime) / (EndTime - ExplosionTime)
Else
    If Now > ExplosionTime Then Explode
End If
If Now > EndTime Then
    Dim Spd As Single
    Spd = Rnd * 120 + 50
    Fire Rnd * 640, 480, PI + (Rnd - 0.5) / 2, Spd, Rnd * 30 + 30, Rnd * 100 + 50, 300000 / Spd, Rnd * 5000 + 5000
End If
End Sub

Sub Draw()
Dim l As Long
Dim x As Integer
Dim y As Integer
Dim BackColor As Long
Dim ForeColor As Long
Dim r As Byte, g As Byte, b As Byte
Dim r2 As Byte, g2 As Byte, b2 As Byte
Dim Dark As Single
Dim Bright As Single
On Error Resume Next
If Exploded Then
    For l = 0 To UBound(Particle)
        With Particle(l)
            Select Case .x
                Case 0 To 640
                    Select Case .y
                        Case 0 To 480
                            'If Int(Rnd * 2) Then ForeColor = Color1 Else ForeColor = Color2
                            If .colorref = 1 Then ForeColor = Color1 Else ForeColor = Color2
                            r = ForeColor And 255
                            g = (ForeColor And 65280) \ 256&
                            b = (ForeColor And 16711680) \ 65535
                            BackColor = Backbuffer.GetLockedPixel(.x, .y)
                            r2 = BackColor And 255
                            g2 = (BackColor And 65280) \ 256&
                            b2 = (BackColor And 16711680) \ 65535
                            Dark = 1 - Brightness
                            r = r * Brightness + r2 * Dark
                            g = g * Brightness + g2 * Dark
                            b = b * Brightness + b2 * Dark
                            Backbuffer.SetLockedPixel .x, .y, RGB(r, g, b)
                            Bright = Brightness * 0.5
                            Dark = 1 - Bright
                            r = r * Bright + r2 * Dark
                            g = g * Bright + g2 * Dark
                            b = b * Bright + b2 * Dark
                            Backbuffer.SetLockedPixel .x + 1, .y, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 1, .y, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x, .y + 1, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x, .y - 1, RGB(r, g, b)
                            Bright = Brightness * 0.25
                            Dark = 1 - Bright
                            r = r * Bright + r2 * Dark
                            g = g * Bright + g2 * Dark
                            b = b * Bright + b2 * Dark
                            Backbuffer.SetLockedPixel .x + 1, .y + 1, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x + 1, .y - 1, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 1, .y + 1, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 1, .y - 1, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x + 2, .y, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 2, .y, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x, .y + 2, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x, .y - 2, RGB(r, g, b)
                            Bright = Brightness * 0.125
                            Dark = 1 - Bright
                            r = r * Bright + r2 * Dark
                            g = g * Bright + g2 * Dark
                            b = b * Bright + b2 * Dark
                            Backbuffer.SetLockedPixel .x + 2, .y + 2, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x + 2, .y - 2, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 2, .y + 2, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 2, .y - 2, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x + 2, .y + 1, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x + 2, .y - 1, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 2, .y + 1, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 2, .y - 1, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x + 1, .y + 2, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x + 1, .y - 2, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 1, .y + 2, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 1, .y - 2, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x + 3, .y, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 3, .y, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x, .y + 3, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x, .y - 3, RGB(r, g, b)
                    
                            Bright = Brightness * 0.612
                            Dark = 1 - Bright
                            r = r * Bright + r2 * Dark
                            g = g * Bright + g2 * Dark
                            b = b * Bright + b2 * Dark
                            Backbuffer.SetLockedPixel .x + 2, .y + 3, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x + 2, .y - 3, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 2, .y + 3, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 2, .y - 3, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x + 3, .y + 2, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x + 3, .y - 2, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 3, .y + 2, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 3, .y - 2, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x + 4, .y + 3, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x + 4, .y - 3, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 4, .y + 3, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 4, .y - 3, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x + 3, .y + 4, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x + 3, .y - 4, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 3, .y + 4, RGB(r, g, b)
                            Backbuffer.SetLockedPixel .x - 3, .y - 4, RGB(r, g, b)
                    End Select
            End Select
        End With
    Next
Else
    For l = 0 To UBound(Particle)
        With Particle(l)
            Select Case .x
                Case 0 To 640
                    Select Case .y
                        Case 0 To 480
                            If Int(Rnd * 2) Then ForeColor = Color1 Else ForeColor = Color2
                            Backbuffer.SetLockedPixel .x, .y, ForeColor
                    End Select
            End Select
        End With
    Next
End If
End Sub

Private Sub Explode()
On Error Resume Next
Dim l As Long
Dim Angle As Single
If Exploded Then Exit Sub
l = Int(Rnd * 3) + 1
ExplodeSound(l).Play DSBPLAY_DEFAULT
For l = 0 To UBound(Particle)
    With Particle(l)
        Angle = Rnd * PIby2
        .xm = .xm + (Sin(Angle) * Rnd * ExplosionPower)
        .ym = .ym + (Cos(Angle) * Rnd * ExplosionPower)
    End With
Next
Exploded = True
End Sub

Private Sub Class_Initialize()
ReDim Particle(0)
End Sub
