VERSION 5.00
Begin VB.Form MainForm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2496
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2496
   ScaleWidth      =   3744
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Made by Simon Price on 19/11/00
' Visit www.VBgames.co.uk for more cool VB programs
' Send feedback to Si@VBgames.co.uk

Private Const NUM_FIREWORKS = 20

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Firework(0 To NUM_FIREWORKS - 1) As New cFirework

Private EndNow As Boolean

Private Const PI = 3.1415

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
EndNow = True
End Sub

Private Sub Form_Load()
Randomize Timer
If modDX.Init(hwnd) = False Then
    modDX.CleanUp hwnd
    MsgBox "FATAL ERROR! in DirectX initiation... closing down now!", vbCritical, "ERROR:"
    Unload Me
End If
MainLoop
End Sub

Private Sub MainLoop()
On Error GoTo EndHere
Dim rec As RECT
Dim SurfDesc As DDSURFACEDESC2
Dim TimeElapsed As Long
Dim LastTime As Long
Dim ThisTime As Long
Dim Speed As Single
Dim l As Long
Dim Cur As Long
Cur = ShowCursor(0)
For l = LBound(Firework) To UBound(Firework)
    Speed = Rnd * 120 + 50
    Firework(l).Fire Rnd * 640, 480, PI + (Rnd - 0.5) / 2, Speed, Rnd * 30 + 30, Rnd * 100 + 50, 300000 / Speed, Rnd * 5000 + 5000
Next
LastTime = timeGetTime
Do
    DoEvents
    Backbuffer.BltFast 0, 0, Background, rec, DDBLTFAST_WAIT
    ThisTime = timeGetTime
    For l = LBound(Firework) To UBound(Firework)
        Firework(l).Move (ThisTime - LastTime)
    Next
    Backbuffer.Lock rec, SurfDesc, DDLOCK_WAIT, 0
    For l = LBound(Firework) To UBound(Firework)
        Firework(l).Draw
    Next
    Backbuffer.Unlock rec
    LastTime = ThisTime
    Primary.Flip Nothing, DDFLIP_WAIT
Loop Until EndNow
EndHere:
On Error Resume Next
ShowCursor (Cur)
Unload Me
End Sub

Private Sub Form_LostFocus()
EndNow = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
EndNow = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim l As Long
For l = LBound(Firework) To UBound(Firework)
    Set Firework(l) = Nothing
Next
modDX.CleanUp hwnd
End Sub
