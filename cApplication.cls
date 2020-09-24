VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OPTION SETTINGS ...

    ' Enforce variable declaration
    Option Explicit

' PRIVATE TYPES ...

    ' Input state type
    Private Type tInputState
         MousePX       As Long            ' Mouse screen position
         MousePY       As Long            '
         MouseDX       As Long            ' Mouse position delta
         MouseDY       As Long            '
         MouseB1       As Boolean         ' Mouse button 1
         MouseB2       As Boolean         ' Mouse button 2
         KeyEscape     As Boolean         ' Keyboard key
         KeyArrowU     As Boolean         '
         KeyArrowD     As Boolean         '
         KeyArrowL     As Boolean         '
         KeyArrowR     As Boolean         '
         KeySpace      As Boolean         '
         KeyCtrl       As Boolean         '
         KeyPlus       As Boolean         '
         KeyMinus      As Boolean         '
         KeyP          As Boolean         '
    End Type

' PROPERTY VARIABLES ...

    Public DXInstance           As DirectX7                  ' Instance of DirectX
    Public DDInstance           As DirectDraw4               ' Instance of DirectDraw
    Public D3Instance           As Direct3DRM3               ' Instance of Direct3DRM
    Public DSInstance           As DirectSound               ' Instance of DirectSound

    Public DDPrimary            As DirectDrawSurface4        ' Primary display device
    Public DDSurface            As DirectDrawSurface4        ' Backbuffer for rendering

    Public D3Frame              As Direct3DRMFrame3          ' Root frame of D3DRM hierarchy
    Public D3Camera             As Direct3DRMFrame3          ' Camera frame
    Public D3Device             As Direct3DRMDevice3         ' D3DRM Device
    Public D3Viewport           As Direct3DRMViewport2       ' D3DRM Viewport
    
    Public DSBPrimary          As DirectSoundBuffer          ' Primary sound buffer
    Public DSListener          As DirectSound3DListener      ' Listener to the primary buffer
    
    Public FrameCount           As Long
    Public FrameRate            As Single
    Public FrameLimit           As Long
    
    Public CursorX              As Long
    Public CursorY              As Long
    
    Public Terminating          As Boolean
    
' SUBCLASS INSTANCES ...

    Public Shots            As cShots                    ' Shots in game
    Public Effects          As cEffects                  ' Effects in game
    Public Enemies          As cEnemies                  ' Enemies in game
    Public Particles        As cParticles                ' Particles in game
    Public Player           As cPlayer                   ' Player in game
    Public Environment      As cEnvironment              ' Game environment
    Public Interface        As cInterface                ' User interface display
    
    Public WaterMessage     As Boolean
    Public FuelMessage      As Boolean
    
    Public SoundEnabled     As Boolean
    

    '
    ' INITIALIZE: Initializes application
    '
    Public Sub Initialize()
    
        ' Declare local variables ...
        
            Dim L_dDDSD         As DDSURFACEDESC2           ' Surface descriptor for creating primary and bbuffer

        ' Code ...
            
            ' Prepare windows environment ...
                
                ' Hide cursor
                ShowCursor 0
                
            ' Create DirectX instances ...
                
                ' DirectX
                On Error Resume Next
                Set DXInstance = New DirectX7
                If Err.Number <> 0 Then
                    MsgBox "This program requires Microsoft DirectX7.", vbOKOnly + vbCritical, "ERacer"
                    End
                End If
                On Error GoTo 0
                
                ' DirectDraw
                Set DDInstance = DXInstance.DirectDraw4Create("")
    
                ' DirectSound (optional)
                On Error Resume Next
                Set DSInstance = DXInstance.DirectSoundCreate("")
                On Error GoTo 0
                 
            ' Prepare DirectX states ...
                
                ' Cooperative level (draw)
                DDInstance.SetCooperativeLevel fER.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
                
                ' Display size and color depth
                DDInstance.SetDisplayMode 640, 480, 16, 0&, DDSDM_DEFAULT
                
                ' Cooperative level (sound)
                If Not DSInstance Is Nothing Then DSInstance.SetCooperativeLevel fER.hWnd, DSSCL_PRIORITY
                
            ' Create primary ...
                            
                With L_dDDSD
                    .lFlags = DDSD_CAPS
                    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_VIDEOMEMORY
                End With
                Set DDPrimary = DDInstance.CreateSurface(L_dDDSD)
    
            ' Create backbuffer ...
                            
                ' Create
                With L_dDDSD
                    .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
                    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE Or DDSCAPS_VIDEOMEMORY
                    .lWidth = 640
                    .lHeight = 480
                End With
                Set DDSurface = DDInstance.CreateSurface(L_dDDSD)
                    
                ' Set properties
                With DDSurface
                    .SetFont fER.Font
                    .SetFontTransparency True
                    .SetForeColor RGB(128, 128, 128)
                End With
                    
            ' Initialize Direct3DRM ...
            
                ' Create and initalize D3DRM objects ...

                    ' Create D3DRM instance
                    Set D3Instance = DXInstance.Direct3DRMCreate
                    
                    ' Create D3DRM device
                     Set D3Device = D3Instance.CreateDeviceFromSurface("IID_IDirect3DHALDevice", DDInstance, DDSurface, D3DRMDEVICE_DEFAULT)
    
                    ' Set texture search path
                    D3Instance.SetSearchPath App.Path + "\gfx"
                    
                    ' Create and initialize D3DRM frame for world
                    Set D3Frame = D3Instance.CreateFrame(Nothing)
                    D3Frame.SetSceneBackgroundRGB 0.1, 0.2, 0.4
                    
                    ' Create D3DRM frame for camera: Moves with racer (smooth)
                    Set D3Camera = D3Instance.CreateFrame(D3Frame)
                    
                    ' Create and initialize D3DRM viewport
                    Set D3Viewport = D3Instance.CreateViewport(D3Device, D3Camera, 0, 0, 640, 480)

                ' Initialize D3DRM states ...
                
                    ' Set alphatranslucency enable
                    D3Device.SetRenderMode D3DRMRENDERMODE_BLENDEDTRANSPARENCY Or D3DRMRENDERMODE_SORTEDTRANSPARENCY
                
                    ' Set gouraud shading, texture mapping and lighting enable
                    D3Device.SetQuality D3DRMRENDER_GOURAUD
                                     
                    ' Set texture quality to full
                    D3Device.SetTextureQuality D3DRMTEXTURE_LINEARMIPLINEAR
                    
                    ' Set fog off (not supported, looks funny on many cards that support it)
                    D3Frame.SetSceneFogEnable D_FALSE
                                                    
            ' Initialize sound system ...
                
                If Not DSInstance Is Nothing Then
                    
                    Dim L_dWF As WAVEFORMATEX
                    Dim L_ddsbd As DSBUFFERDESC
                    
                    DSInstance.SetSpeakerConfig DSSPEAKER_HEADPHONE Or DSSPEAKER_STEREO
                    L_dWF = GetWaveFileFormat(App.Path + "\snd\fmt.wav")
                    With L_ddsbd
                        .lBufferBytes = 0
                        .lFlags = DSBCAPS_CTRL3D Or DSBCAPS_PRIMARYBUFFER
                    End With
                    Set DSBPrimary = Application.DSInstance.CreateSoundBuffer(L_ddsbd, L_dWF)
                    Set DSListener = Application.DSBPrimary.GetDirectSound3DListener
                                                            
                    With DSListener
                        .SetDistanceFactor 1, DS3D_IMMEDIATE
                        .SetDopplerFactor 1, DS3D_IMMEDIATE
                        .SetRolloffFactor 1, DS3D_IMMEDIATE
                        .SetVelocity 0, 0, 0, DS3D_IMMEDIATE
                    End With
                    
                End If
                
                   
            ' Initialize general data ...
            
                FrameCount = 0
                FrameRate = 0
                FrameLimit = 50
                
            ' Initialize subclasses ...
            
                Set Interface = New cInterface
                Interface.Initialize
                Interface.Render True
                
                Set Environment = New cEnvironment
                Environment.Initialize
                
                Set Shots = New cShots
                Shots.Initialize
                
                Set Effects = New cEffects
                Effects.Initialize
                
                Set Enemies = New cEnemies
                Enemies.Initialize
                
                Set Particles = New cParticles
                Particles.Initialize
                
                Set Player = New cPlayer
                Player.Initialize

    End Sub

    '
    ' EXECUTE: Application main loop
    '
    Public Sub Execute()
    
        ' Declare local variables ...
        
            Dim L_nNextFrame    As Long     ' Time at which to advance to next frame
            Dim L_nNextRate     As Long     ' Time at which to calculate next frame rate
            Dim L_nFrameRate    As Long     ' Local frame rate counter
            
        ' Code ...
        
            ' Master loop
            Do
                
                ' Execute frame timing calculations
                
                    L_nNextFrame = timeGetTime + (1000 \ FrameLimit)
                    FrameCount = FrameCount + 1
                    
                    L_nFrameRate = L_nFrameRate + 1
                    If timeGetTime > L_nNextRate Then
                        FrameRate = L_nFrameRate
                        L_nFrameRate = 0
                        L_nNextRate = timeGetTime + 1000
                    End If
                
                ' React to user input
                    Interact
                    
                ' Update state
                    Update
                    
                ' Render next frame
                    Render
                
                ' React to windows events
                    If FrameCount Mod 5 = 0 Then DoEvents
                    
                ' Wait until frame time is over
                Do
                Loop Until timeGetTime >= L_nNextFrame
                
            Loop Until Application.Terminating
    
    End Sub

    '
    ' TERMINATE: Cleans up application
    '
    Public Sub Terminate()
    
        ' Restore DirectDraw windows states ...
        
            DDInstance.RestoreDisplayMode
            DDInstance.SetCooperativeLevel fER.hWnd, DDSCL_NORMAL
            
        ' Terminate Direct3DRM ...
            
            ' Terminate viewport
            Set D3Viewport = Nothing
            
            ' Terminate device
            Set D3Device = Nothing
            
            ' Terminate instance
            Set D3Instance = Nothing
            
        ' Terminate DirectDraw ...
    
            Set DDSurface = Nothing
            Set DDPrimary = Nothing
            
        ' Terminate DirectSound ...
        
            Set DSInstance = Nothing
            
        ' Terminate DirectX ...
          
            Set DXInstance = Nothing
            
        ' Restore windows environment ...
        
            ' Show cursor
            ShowCursor -1
                
    End Sub

    '
    ' UPDATE: Update gamestate
    '
    Public Sub Update()
    
            
        ' Update enemies
        
            Enemies.Update
            
        ' Update player
        
            Player.Update
            
        ' Update environment
        
            Environment.Update
                 
        ' Update shots
        
            Shots.Update
            
        ' Update effects
        
            Effects.Update
            
        ' Update particles
        
            Particles.Update
            
        ' Update initial messages
        
            If Application.FrameCount = 1 Then Application.Interface.DisplayMessage "welcome to ERACER ..."
            If Application.FrameCount = 300 And (Not Application.Enemies.Active) Then Application.Interface.DisplayMessage "demonstrating industrial quality 3D using visual basic ..."
            If Application.FrameCount = 800 And (Not Application.Enemies.Active) Then Application.Interface.DisplayMessage "use arrow keys to steer, space to jump, ctrl to fire ..."
            If Application.FrameCount = 1300 And (Not Application.Enemies.Active) Then Application.Interface.DisplayMessage "use menu buttons to change environment and start combat ..."
         
    End Sub

    '
    ' INTERACT: Gather and process user input
    '
    Public Sub Interact()
    
        ' Declare local variables ...
        
            Static S_nKeyDelay          As Long         ' Delay for keys
            Static S_dMousePos          As POINTAPI     ' Mouse position now
            Static S_dMouseOld          As POINTAPI     ' Old mouse position
            Static S_dInputState        As tInputState  ' State of input controls
            Dim L_dJoy                  As JOYINFOEX
            Dim L_bJoy                  As Boolean
            
        ' Code ...
        
            ' Gather input ...
            
                ' Joy
                    
                    L_dJoy.dwSize = Len(L_dJoy)
                    L_dJoy.dwFlags = 131
                    L_bJoy = (joyGetPosEx(0, L_dJoy) = 0)
                    
                ' Gather mouse input...
                        
                    ' Get position
                    GetCursorPos S_dMousePos
                    
                    ' Calculate derived values
                    S_dInputState.MouseDX = S_dMousePos.X - 320
                    S_dInputState.MouseDY = S_dMousePos.Y - 240
                    S_dInputState.MousePX = S_dMousePos.X
                    S_dInputState.MousePY = S_dMousePos.Y
                                                    
                    ' Get button state
                    S_dInputState.MouseB1 = (GetAsyncKeyState(&H1) <= -32767)
                    S_dInputState.MouseB2 = (GetAsyncKeyState(&H2) <= -32767)
                                                    
                ' Gather keyboard input (continuous keys)
                          
                    ' ARROW UP
                    If (GetAsyncKeyState(&H26) <= -32767) Then
                        S_dInputState.KeyArrowU = True
                    Else
                        S_dInputState.KeyArrowU = False
                    End If
                                
                    ' ARROW DOWN
                    If (GetAsyncKeyState(&H28) <= -32767) Then
                        S_dInputState.KeyArrowD = True
                    Else
                        S_dInputState.KeyArrowD = False
                    End If
                    
                    ' ARROW LEFT
                    If (GetAsyncKeyState(&H25) <= -32767) Then
                        S_dInputState.KeyArrowL = True
                    Else
                        S_dInputState.KeyArrowL = False
                    End If
                                
                    ' ARROW RIGHT
                    If (GetAsyncKeyState(&H27) <= -32767) Then
                        S_dInputState.KeyArrowR = True
                    Else
                        S_dInputState.KeyArrowR = False
                    End If
                    
                    ' SPACE
                    If (GetAsyncKeyState(&H20) <= -32767) Or (L_bJoy And (L_dJoy.dwButtonNumber = 2)) Then
                        S_dInputState.KeySpace = True
                    Else
                        S_dInputState.KeySpace = False
                    End If
                    
                ' Gather keyboard input (event keys)
                
                    If S_nKeyDelay = 0 Then
                                        
                         ' ESCAPE
                        If (GetAsyncKeyState(&H1B) <= -32767) Then
                            S_dInputState.KeyEscape = True
                            S_nKeyDelay = 10
                        Else
                            S_dInputState.KeyEscape = False
                        End If
                        
                        ' PLUS
                        If (GetAsyncKeyState(&H6B) <= -32767) Then
                            S_dInputState.KeyPlus = True
                            S_nKeyDelay = 10
                        Else
                            S_dInputState.KeyPlus = False
                        End If
                        
                        ' MINUS
                        If (GetAsyncKeyState(&H6D) <= -32767) Then
                            S_dInputState.KeyMinus = True
                            S_nKeyDelay = 10
                        Else
                            S_dInputState.KeyMinus = False
                        End If
                        
                       
                        ' CTRL
                        If (GetAsyncKeyState(&H11) <= -32767) Or (L_bJoy And (L_dJoy.dwButtonNumber = 1)) Then
                            S_dInputState.KeyCtrl = True
                            S_nKeyDelay = 10
                        Else
                            S_dInputState.KeyCtrl = False
                        End If
                        
                    Else
                        
                        ' Reset states ...
                        
                            S_dInputState.KeyEscape = False
                            S_dInputState.KeyPlus = False
                            S_dInputState.KeyMinus = False
                            S_dInputState.KeyCtrl = False
                            S_nKeyDelay = S_nKeyDelay - 1
                        
                    End If
                    If L_bJoy Then
                        S_dInputState.KeyArrowL = S_dInputState.KeyArrowL Or (L_dJoy.dwXpos < 32767& - 15000)
                        S_dInputState.KeyArrowR = S_dInputState.KeyArrowR Or (L_dJoy.dwXpos > 32767& + 15000)
                        S_dInputState.KeyArrowU = S_dInputState.KeyArrowU Or (L_dJoy.dwYpos < 32767& - 15000)
                        S_dInputState.KeyArrowD = S_dInputState.KeyArrowD Or (L_dJoy.dwYpos > 32767& + 15000)
                    End If
                    
            ' Process input ...
                
                ' General functions ...
                
                    ' ESCAPE: Leave game
                    If S_dInputState.KeyEscape Then
                        Terminating = True
                    End If
                    
                   
                    ' PLUS: Increase viewport size
                    If S_dInputState.KeyPlus Then
                        If Interface.ViewportSize > 0 Then
                            Interface.ViewportSize = Interface.ViewportSize - 1
                        End If
                    End If
                         
                    ' MINUS: Decrease viewport size
                    If S_dInputState.KeyMinus Then
                        If Interface.ViewportSize < 20 Then
                            Interface.ViewportSize = Interface.ViewportSize + 1
                        End If
                    End If
                    
                    ' Mouse on menu ...
                    
                        ' Only when left button click
                        If S_dInputState.MouseB1 Then
                        
                            
                            ' Change flyer
                            If S_dInputState.MousePX > 480 + 57 And S_dInputState.MousePX < 480 + 57 + 22 Then
                                If S_dInputState.MousePY > 390 + 9 And S_dInputState.MousePY < 390 + 9 + 22 Then
                                    If S_nKeyDelay = 0 Then
                                            S_nKeyDelay = 10
                                            Player.ChangeRacer
                                            Application.Enemies.Active = False
                                            Application.Reset
                                            Application.Interface.DisplayMessage "new fighter arrived ... " & IIf(Application.Player.Model = rmJaguar, "Jaguar", "Warthog")
                                    End If
                                End If
                            End If
                            
                            ' Change daytime
                            If S_dInputState.MousePX > 480 + 32 And S_dInputState.MousePX < 480 + 32 + 22 Then
                                If S_dInputState.MousePY > 390 + 34 And S_dInputState.MousePY < 390 + 34 + 22 Then
                                    If S_nKeyDelay = 0 Then
                                            S_nKeyDelay = 10
                                            Environment.ChangeDaytime
                                            Application.Enemies.Active = False
                                            Application.Reset
                                            Application.Interface.DisplayMessage "daytime changed"
                                    End If
                                End If
                            End If
                        
                            ' Change island
                            If S_dInputState.MousePX > 480 + 7 And S_dInputState.MousePX < 480 + 7 + 22 Then
                                If S_dInputState.MousePY > 390 + 59 And S_dInputState.MousePY < 390 + 59 + 22 Then
                                    If S_nKeyDelay = 0 Then
                                            S_nKeyDelay = 10
                                            Interface.Render True
                                            If (Not Application.DSInstance Is Nothing) And Application.SoundEnabled Then
                                                DSListener.SetPosition 0, 0, -1000, DS3D_IMMEDIATE
                                            End If
                                            Environment.ChangeMap
                                            Application.Enemies.Active = False
                                            Application.Reset
                                            Application.Interface.DisplayMessage "new island selected ... " & Application.Environment.IslandName
                                    End If
                                End If
                            End If
                            
                            ' Change screen size (+)
                            If S_dInputState.MousePX > 480 + 100 And S_dInputState.MousePX < 480 + 100 + 22 Then
                                If S_dInputState.MousePY > 390 + 9 And S_dInputState.MousePY < 390 + 9 + 22 Then
                                    If Interface.ViewportSize < 20 Then
                                        Interface.ViewportSize = Interface.ViewportSize + 1
                                    End If
                                End If
                            End If
                            
                            ' Change screen size (-)
                            If S_dInputState.MousePX > 480 + 130 And S_dInputState.MousePX < 480 + 130 + 22 Then
                                If S_dInputState.MousePY > 390 + 9 And S_dInputState.MousePY < 390 + 9 + 22 Then
                                    If Interface.ViewportSize > 0 Then
                                        Interface.ViewportSize = Interface.ViewportSize - 1
                                    End If
                                End If
                            End If
                            
                            ' Exit
                            If S_dInputState.MousePX > 480 + 130 And S_dInputState.MousePX < 480 + 130 + 22 Then
                                If S_dInputState.MousePY > 390 + 60 And S_dInputState.MousePY < 390 + 60 + 22 Then
                                    Terminating = True
                                End If
                            End If
                            
                            ' Combat
                            If S_dInputState.MousePX > 480 + 40 And S_dInputState.MousePX < 480 + 70 + 40 Then
                                If S_dInputState.MousePY > 390 + 40 And S_dInputState.MousePY < 390 + 43 + 40 Then
                                    If S_nKeyDelay = 0 Then
                                        Enemies.Active = Not Enemies.Active
                                        Application.Interface.DisplayMessage IIf(Enemies.Active, "incoming fighters detected, protect station ...", "sector clear, no enemy activity reported ...")
                                        Application.Reset
                                        S_nKeyDelay = 10
                                    End If
                                End If
                            End If
                            
                            ' Sound
                            If S_dInputState.MousePX > 480 + 114 And S_dInputState.MousePX < 480 + 118 + 17 Then
                                If S_dInputState.MousePY > 390 + 33 And S_dInputState.MousePY < 390 + 33 + 17 Then
                                    If S_nKeyDelay = 0 Then
                                        If Not (Application.DSInstance Is Nothing) Then
                                            Application.SoundEnabled = Not Application.SoundEnabled
                                            Application.Enemies.ToggleSound
                                            Application.Player.ToggleSound
                                        End If
                                        S_nKeyDelay = 10
                                    End If
                                End If
                            End If
                            
                        End If
                        
                ' Racer control ...
                    
                    ' UP: Accellerate
                    If S_dInputState.KeyArrowU Then Player.Command pcAccellerate
                    
                    ' DOWN: Decellerate
                    If S_dInputState.KeyArrowD Then Player.Command pcDecellerate
                           
                    ' LEFT: Bank left
                    If S_dInputState.KeyArrowR Then Player.Command pcBankLeft
                    
                    ' RIGHT: Bank right
                    If S_dInputState.KeyArrowL Then Player.Command pcBankRight
                    
                    ' SPACE: Jump
                    If S_dInputState.KeySpace Then
                        Player.Command pcJump
                        If Not Me.FuelMessage Then
                            If Not Application.Enemies.Active Then Application.Interface.DisplayMessage "jumps use up lots of fuel, watch the blue fuel bar ..."
                            Me.FuelMessage = True
                        End If
                    End If
                    
                    ' CTRL: Shoot
                    If S_dInputState.KeyCtrl Then Player.Command pcShoot
            
            ' Limit cursor to menu, remember cursor position ...
            
                If S_dInputState.MousePX < 470 Then S_dInputState.MousePX = 470
                If S_dInputState.MousePY < 380 Then S_dInputState.MousePY = 380
                If S_dInputState.MousePX > 624 Then S_dInputState.MousePX = 624
                If S_dInputState.MousePY > 464 Then S_dInputState.MousePY = 464
                SetCursorPos S_dInputState.MousePX, S_dInputState.MousePY
                
                CursorX = S_dInputState.MousePX
                CursorY = S_dInputState.MousePY
            
    End Sub

    '
    ' RENDER: Render frame
    '
    Public Sub Render()
    
        ' Declare local variables ...
        
            Dim L_dArea         As RECT                 ' Area from/to which to render
            Dim L_dDDSD         As DDSURFACEDESC2       ' Descriptor for loading static surfaces
            Dim L_dDDCK         As DDCOLORKEY           ' Colorkey for making static surfaces transparent
            Static S_oDDSCursor As DirectDrawSurface4   ' Surface holding cursor
            Static S_oDDSLogo   As DirectDrawSurface4   ' Surface holding logo
            Static S_oDDSEracer As DirectDrawSurface4   ' Surface holding game logo
            
        ' Code ...
        
            ' Load static surfaces ...
            
                ' Cursor ...
                                        
                    If S_oDDSCursor Is Nothing Then
                        
                        ' Prepare surface descriptor
                        With L_dDDSD
                            .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
                            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
                            .lWidth = 16
                            .lHeight = 16
                        End With
                        
                        ' Load
                        Set S_oDDSCursor = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\gfx\cursor.bmp", L_dDDSD)
                
                        ' Make transparent
                        With L_dDDCK
                            .low = 0
                            .high = 0
                        End With
                        S_oDDSCursor.SetColorKey DDCKEY_SRCBLT, L_dDDCK
                        
                    End If
                    
                ' Logo ...
                
                    If S_oDDSLogo Is Nothing Then
                        
                        ' Prepare surface descriptor
                        With L_dDDSD
                            .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
                            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
                            .lWidth = 359
                            .lHeight = 59
                        End With
                        
                        ' Load
                        Set S_oDDSLogo = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\gfx\nls.bmp", L_dDDSD)
                
                    End If
        
                ' Eracer Logo ...
                
                    If S_oDDSEracer Is Nothing Then
                        
                        ' Prepare surface descriptor
                        With L_dDDSD
                            .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
                            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
                            .lWidth = 240
                            .lHeight = 60
                        End With
                        
                        ' Load
                        Set S_oDDSEracer = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\gfx\logo.bmp", L_dDDSD)
                
                    End If
        
            ' Clear cursor area ...
            
                With L_dArea
                    .Left = 470
                    .Top = 380
                    .Right = 640
                    .Bottom = 480
                End With
                Application.DDSurface.BltColorFill L_dArea, 0
               
            ' Render background if necessary
            If Interface.ViewportSizeChanged Then
            
                ' Reset
                Interface.ViewportSizeChanged = False
                
                ' Logo ...
                
                    With L_dArea
                        .Left = 0
                        .Top = 0
                        .Right = 359
                        .Bottom = 59
                    End With
                    Application.DDSurface.BltFast 5, 415, S_oDDSLogo, L_dArea, DDBLTFAST_NOCOLORKEY Or DDBLTFAST_WAIT
                   
                ' Sign ...
                
                    With L_dArea
                        .Left = 0
                        .Top = 0
                        .Right = 240
                        .Bottom = 60
                    End With
                    Application.DDSurface.BltFast 360, 40, S_oDDSEracer, L_dArea, DDBLTFAST_NOCOLORKEY Or DDBLTFAST_WAIT
                
                End If
                
            ' Render 3D ...

                ' Clear
                D3Viewport.Clear D3DRMCLEAR_TARGET Or D3DRMCLEAR_ZBUFFER

                ' Render
                D3Instance.Tick 1
            
            ' Render 2D interface ...
                
                ' Render
                Interface.Render
               
            ' Render Cursor ...
            
                With L_dArea
                    .Left = 0
                    .Top = 0
                    .Right = 16
                    .Bottom = 16
                End With
                DDSurface.BltFast CursorX, CursorY, S_oDDSCursor, L_dArea, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
               
            ' Render backbuffer onto primary ...
            
                ' Set source area
                With L_dArea
                    .Left = 0
                    .Top = 0
                    .Right = 640
                    .Bottom = 480
                End With
                
                ' Render backbuffer onto primary
                DDPrimary.BltFast 0, 0, DDSurface, L_dArea, DDBLTFAST_NOCOLORKEY Or DDBLTFAST_WAIT
                
    End Sub
    
    '
    ' RESET: Reset states on all subclasses
    '
    Public Sub Reset()
        Player.Reset
        Environment.Reset
        Effects.Reset
        Particles.Reset
        Enemies.Reset
        Shots.Reset
    End Sub

