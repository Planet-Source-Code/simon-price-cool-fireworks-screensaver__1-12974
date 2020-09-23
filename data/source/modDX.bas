Attribute VB_Name = "modDX"
Public DX As New DirectX7
Public DDRAW As DirectDraw7
Public DSOUND As DirectSound
Public Primary As DirectDrawSurface7
Public Backbuffer As DirectDrawSurface7
Public Background As DirectDrawSurface7
Public ExplodeSound(1 To 3) As DirectSoundBuffer
Public RocketSound(1 To 3) As DirectSoundBuffer

' initiate directx objects
Function Init(hwnd As Long) As Boolean
On Error Resume Next
Dim SurfDesc As DDSURFACEDESC2
Dim Caps As DDSCAPS2
Dim i As Integer
Dim BufferDesc As DSBUFFERDESC
Dim WavFormat As WAVEFORMATEX
' create directdraw
Set DDRAW = DX.DirectDrawCreate("")
' set coop level and screen size
DDRAW.SetCooperativeLevel hwnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE Or DDSCL_ALLOWREBOOT
DDRAW.SetDisplayMode 640, 480, 24, 0, DDSDM_DEFAULT
' create primary surface
With SurfDesc
    .lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    .lBackBufferCount = 1
End With
Set Primary = DDRAW.CreateSurface(SurfDesc)
' attach a backbuffer
Caps.lCaps = DDSCAPS_BACKBUFFER
Set Backbuffer = Primary.GetAttachedSurface(Caps)
' load background
SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
SurfDesc.lFlags = DDSD_CAPS
Dim temppic As IPictureDisp
Set temppic = LoadPicture(App.Path & "\data\graphics\background.jpg")
SavePicture temppic, App.Path & "\data\graphics\background.bmp"
Set Background = DDRAW.CreateSurfaceFromFile(App.Path & "\data\graphics\background.bmp", SurfDesc)
' create directsound
Set DSOUND = DX.DirectSoundCreate("")
' set coop level
DSOUND.SetCooperativeLevel hwnd, DSSCL_PRIORITY
' load sounds
For i = 1 To 3
    Set ExplodeSound(i) = DSOUND.CreateSoundBufferFromFile(App.Path & "\data\sound\explode" & i & ".wav", BufferDesc, WavFormat)
    Set RocketSound(i) = DSOUND.CreateSoundBufferFromFile(App.Path & "\data\sound\rocket" & i & ".wav", BufferDesc, WavFormat)
Next
' report success/failure
If Err.Number = DD_OK Then Init = True
End Function

' clean up directx objects
Function CleanUp(hwnd As Long) As Boolean
On Error Resume Next
' change screen res back and set coop level to normal
DDRAW.RestoreDisplayMode
DDRAW.SetCooperativeLevel hwnd, DDSCL_NORMAL
' set surfaces to nothing
Set Backbuffer = Nothing
Set Primary = Nothing
' set sound buffers to nothing
Dim i As Integer
For i = 1 To 3
    Set ExplodeSound(i) = Nothing
    Set RocketSound(i) = Nothing
Next
' set directx objects to nothing
Set DSOUND = Nothing
Set DDRAW = Nothing
Set DX = Nothing
' report success/failure
If Err.Number = DD_OK Then CleanUp = True
End Function
