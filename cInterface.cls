VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OPTION SETTINGS ...

    ' Enforce variable declarations
    Option Explicit

' PRIVATE VARIABLES ...

    Private I_oDDSBar           As DirectDrawSurface4        ' Surface holding bar display
    Private I_oDDSBar2          As DirectDrawSurface4        ' Surface holding 2nd bar display
    Private I_oDDSDamage        As DirectDrawSurface4        ' Surface holding bar display element
    Private I_oDDSFuel          As DirectDrawSurface4        ' Surface holding bar display element
    Private I_oDDSSpeed         As DirectDrawSurface4        ' Surface holding bar display element
    Private I_oDDSAgrav         As DirectDrawSurface4        ' Surface holding bar display element
    Private I_oDDSCursor        As DirectDrawSurface4        ' Surface holding cursor
    Private I_oDDSCombat        As DirectDrawSurface4        ' Surface holding combat sign
    Private I_oDDSText          As DirectDrawSurface4        ' Surface holding scrolling text
    Private I_oDDSSnd           As DirectDrawSurface4        ' Surface holding sound sign
    
    Private I_nTextPos          As Single                    ' Position of scrolling text
    Private I_nViewportSize     As Long                      ' Size of 3D viewport
    Public MessageInProgress    As Boolean
    
    Public ViewportSizeChanged  As Boolean
    
    
' CODE ...

    '
    ' DISPLAYMESSAGE: Displays a message in the interface
    '
    Public Sub DisplayMessage(ByVal P_sMessage As String)
    
        Dim L_dArea As RECT
        Static S_sMessage As String
        
        If P_sMessage = S_sMessage Then Exit Sub
        S_sMessage = P_sMessage
        
        With L_dArea
            .Left = 0
            .Top = 0
            .Right = 640
            .Bottom = 15
        End With
        
        I_oDDSText.BltColorFill L_dArea, 0
        I_oDDSText.SetForeColor RGB(0, 255, 0)
        I_oDDSText.SetFont fER.Font
        
        I_oDDSText.DrawText 122, 1, P_sMessage, False
        
        I_nTextPos = 640
            
        MessageInProgress = True
        
    End Sub
    
    
    '
    ' INITIALIZE: Initialize interface
    '
    Public Sub Initialize()
            
        ' Declare local variables ...
        
            Dim L_dDDSD         As DDSURFACEDESC2           ' Surface descriptor for creating primary and bbuffer
            Dim L_dDDCK         As DDCOLORKEY               ' Color key for making surfaces transparent

        ' Code ...
        
            ' Prepare surface description
            With L_dDDSD
                .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
                .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
                .lWidth = 0
                .lHeight = 0
            End With
            
            ' Prepare color key
            With L_dDDCK
                .low = 0
                .high = 0
            End With
            
            ' Load and prepare bar display ...
                
                L_dDDSD.lWidth = 640
                L_dDDSD.lHeight = 90
                Set I_oDDSBar = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\gfx\bar.bmp", L_dDDSD)
                I_oDDSBar.SetColorKey DDCKEY_SRCBLT, L_dDDCK
                
                L_dDDSD.lWidth = 160
                L_dDDSD.lHeight = 90
                Set I_oDDSBar2 = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\gfx\bar2.bmp", L_dDDSD)
                I_oDDSBar2.SetColorKey DDCKEY_SRCBLT, L_dDDCK
                
                L_dDDSD.lWidth = 94
                L_dDDSD.lHeight = 20
                Set I_oDDSFuel = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\gfx\fuel.bmp", L_dDDSD)
                I_oDDSFuel.SetColorKey DDCKEY_SRCBLT, L_dDDCK
                
                L_dDDSD.lWidth = 106
                L_dDDSD.lHeight = 13
                Set I_oDDSDamage = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\gfx\damage.bmp", L_dDDSD)
                I_oDDSDamage.SetColorKey DDCKEY_SRCBLT, L_dDDCK
                
                L_dDDSD.lWidth = 52
                L_dDDSD.lHeight = 25
                Set I_oDDSSpeed = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\gfx\speed.bmp", L_dDDSD)
                I_oDDSSpeed.SetColorKey DDCKEY_SRCBLT, L_dDDCK
                
                L_dDDSD.lWidth = 18
                L_dDDSD.lHeight = 25
                Set I_oDDSAgrav = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\gfx\agrav.bmp", L_dDDSD)
                I_oDDSAgrav.SetColorKey DDCKEY_SRCBLT, L_dDDCK
                
                L_dDDSD.lWidth = 16
                L_dDDSD.lHeight = 16
                Set I_oDDSCursor = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\gfx\cursor.bmp", L_dDDSD)
                I_oDDSCursor.SetColorKey DDCKEY_SRCBLT, L_dDDCK
                
                L_dDDSD.lWidth = 34
                L_dDDSD.lHeight = 33
                Set I_oDDSCombat = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\gfx\combat.bmp", L_dDDSD)
                I_oDDSCombat.SetColorKey DDCKEY_SRCBLT, L_dDDCK
                        
                L_dDDSD.lWidth = 17
                L_dDDSD.lHeight = 17
                Set I_oDDSSnd = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\gfx\snd.bmp", L_dDDSD)
                I_oDDSSnd.SetColorKey DDCKEY_SRCBLT, L_dDDCK
                        
            ' Prepare scrolling text surface
                L_dDDSD.lWidth = 640
                L_dDDSD.lHeight = 15
                Set I_oDDSText = Application.DDInstance.CreateSurface(L_dDDSD)
                
        ' Initialize general data ...
            
            ' Set initial viewport size
            ViewportSize = 11
                
    End Sub

    '
    ' RENDER: Render 2D interface to backbuffer
    '
    Public Sub Render(Optional ByVal P_bRenderLoading As Boolean)

        ' Declare local variables ...
        
            Dim L_dArea         As RECT                 ' Area from/to which to render
            Dim L_dDDSD         As DDSURFACEDESC2       ' Descriptor for loading static surfaces
            
            Static S_oDDSLogo   As DirectDrawSurface4   ' Surface holding logo
            Static S_oDDSLoad   As DirectDrawSurface4   ' Surface holding loading sign
            Static S_oDDSEracer As DirectDrawSurface4   ' Surface holding game logo
            
        ' Code ...
        
            ' Load static surfaces ...
                                        
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
                    
                ' Loading sign ...
                
                    If S_oDDSLoad Is Nothing Then
                        
                        ' Prepare surface descriptor
                        With L_dDDSD
                            .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
                            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
                            .lWidth = 178
                            .lHeight = 26
                        End With
                        
                        ' Load
                        Set S_oDDSLoad = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\gfx\loading.bmp", L_dDDSD)
                
                    End If
                    
            ' Loading display ...
            
                If P_bRenderLoading Then
                                    
                    ' Clear background ...
                    
                        With L_dArea
                            .Left = 0
                            .Top = 0
                            .Right = 640
                            .Bottom = 480
                        End With
                        Application.DDSurface.BltColorFill L_dArea, 0
                    
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
                       
                    ' Loading sign ...
                    
                        With L_dArea
                            .Left = 0
                            .Top = 0
                            .Right = 178
                            .Bottom = 26
                        End With
                        Application.DDSurface.BltFast 231, 232, S_oDDSLoad, L_dArea, DDBLTFAST_NOCOLORKEY Or DDBLTFAST_WAIT
                
                    ' Instantly update primary
                    
                        With L_dArea
                            .Left = 0
                            .Top = 0
                            .Right = 640
                            .Bottom = 480
                        End With
                        Application.DDPrimary.BltFast 0, 0, Application.DDSurface, L_dArea, DDBLTFAST_NOCOLORKEY Or DDBLTFAST_WAIT
                        
                    ' Finished
                    Exit Sub
                
                End If
                
            ' Menu display ...
            
                ' Default configuration ...
            
                    With L_dArea
                        .Left = 0
                        .Top = 0
                        .Right = 160
                        .Bottom = 90
                    End With
                    Application.DDSurface.BltFast 480, 390, I_oDDSBar2, L_dArea, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                
                ' Blinking combat sign  ...
                
                    If Application.Enemies.Active Then
                        If Application.FrameCount Mod 30 < 15 Then
                        
                            With L_dArea
                                .Left = 0
                                .Top = 0
                                .Right = 34
                                .Bottom = 33
                            End With
                            Application.DDSurface.BltFast 480 + 74, 390 + 46, I_oDDSCombat, L_dArea, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                            
                        End If
                    End If
            
                ' Sound enabled sign ...
                
                    If Application.SoundEnabled Then
                        With L_dArea
                            .Left = 0
                            .Top = 0
                            .Right = 17
                            .Bottom = 17
                        End With
                        Application.DDSurface.BltFast 480 + 118, 390 + 33, I_oDDSSnd, L_dArea, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                    End If
                    
            ' Bar display ...
            
                ' Background ...
                
                    ' Set area
                    With L_dArea
                        .Left = 0
                        .Top = 0
                        .Right = 640
                        .Bottom = 90
                    End With
                    
                    ' Render
                    Application.DDSurface.BltFast 0, 0, I_oDDSBar, L_dArea, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                
                ' FPS display ...
                                    
                    Application.DDSurface.SetForeColor RGB(32, 32, 32)
                    Application.DDSurface.DrawLine 573, 24, 630, 24
                    Application.DDSurface.SetForeColor RGB((55 - Application.FrameRate) * 5, Application.FrameRate * 5, 0)
                    Application.DDSurface.DrawLine 573, 24, 575 + Application.FrameRate, 24
                                    
                ' Game related...
                
                    ' Map on bar display ...
                    
                        ' Set area
                        Application.Environment.DDSMap.GetSurfaceDesc L_dDDSD
                        With L_dArea
                            .Left = 0
                            .Top = 0
                            .Right = L_dDDSD.lWidth
                            .Bottom = L_dDDSD.lHeight
                        End With
                        
                        ' Render
                        Application.DDSurface.BltFast 45 - L_dDDSD.lWidth \ 2, 45 - L_dDDSD.lHeight \ 2, Application.Environment.DDSMap, L_dArea, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                    
                        ' Set player
                        Application.DDSurface.SetForeColor RGB(0, 128, 255)
                        Application.DDSurface.DrawCircle 45 - L_dDDSD.lWidth \ 2 + Application.Player.X, 45 - L_dDDSD.lHeight \ 2 + Application.Player.Z, 1 + (Application.FrameCount Mod 24) \ 6
                        Application.DDSurface.DrawLine 45 - L_dDDSD.lWidth \ 2 + Application.Player.X, 45 - L_dDDSD.lHeight \ 2 + Application.Player.Z, 45 - L_dDDSD.lWidth \ 2 + Application.Player.X + Cos(Application.Player.Heading) * 4, 45 - L_dDDSD.lHeight \ 2 + Application.Player.Z + Sin(Application.Player.Heading) * 4
                        
                        ' Set enemies
                        Application.Enemies.Render
                    
                    ' Info display ...
                    
                        ' Damage
                        With L_dArea
                            .Left = 0
                            .Top = 0
                            .Right = 106 * (Application.Player.Health / IIf(Application.Player.Model = rmWarthog, 1000, 700))
                            .Bottom = 13
                        End With
                        Application.DDSurface.BltFast 92, 15, I_oDDSDamage, L_dArea, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                                            
                        ' Fuel
                        With L_dArea
                            .Left = 0
                            .Top = 0
                            .Right = 94 * (Application.Player.Fuel / IIf(Application.Player.Model = rmJaguar, 10500, 16500))
                            .Bottom = 20
                        End With
                        Application.DDSurface.BltFast 92, 25, I_oDDSFuel, L_dArea, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                        
                        ' Speed
                        With L_dArea
                            .Left = 0
                            .Top = 0
                            .Right = 52 * IIf(Application.Player.Velocity > 0, (Application.Player.Velocity / IIf(Application.Player.Model = rmJaguar, 0.125, 0.1)), 0)
                            .Bottom = 25
                        End With
                        Application.DDSurface.BltFast 117, 50, I_oDDSSpeed, L_dArea, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                        
                        ' Agrav
                        With L_dArea
                            .Left = 0
                            .Top = 25 - Application.Player.Hoover * 200
                            .Right = 18
                            .Bottom = 25
                        End With
                        Application.DDSurface.BltFast 92, 75 - Application.Player.Hoover * 200, I_oDDSAgrav, L_dArea, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            
                    ' Text display
                    
                        If I_nTextPos > 122 Then
                            
                            I_nTextPos = I_nTextPos - 1
                            With L_dArea
                                .Left = (640 - I_nTextPos)
                                .Top = 0
                                .Right = .Left + 122
                                .Bottom = 15
                            End With
                            Application.DDSurface.BltFast 219, 10, I_oDDSText, L_dArea, DDBLTFAST_NOCOLORKEY Or DDBLTFAST_WAIT
                            
                        Else
                            MessageInProgress = False
                        End If
    End Sub

    '
    ' VIEWPORTSIZE GET: Return size of viewport
    '
    Public Property Get ViewportSize() As Long
        ViewportSize = I_nViewportSize
    End Property
    
    '
    ' VIEWPORTSIZE LET: Set size of viewport, create accordingly
    '
    Public Property Let ViewportSize(ByVal P_nViewportSize As Long)
    
        ' Declare local variables ...
            
            Dim L_dArea As RECT ' Area for clearing
            
        ' Code ...
            
            ' Clear background ...
            
                With L_dArea
                    .Left = 0
                    .Top = 0
                    .Right = 640
                    .Bottom = 480
                End With
                Application.DDSurface.BltColorFill L_dArea, 0
        
            ' Create and configure new viewport ...
            
                Set Application.D3Viewport = Application.D3Instance.CreateViewport(Application.D3Device, Application.D3Camera, P_nViewportSize * 9, P_nViewportSize * 6, 640 - P_nViewportSize * 18, 480 - P_nViewportSize * 12)
                With Application.D3Viewport
                    .SetFront 0.25
                    .SetBack 20
                    .SetField 0.3
                End With
        
            ' Remember size ...
            
                I_nViewportSize = P_nViewportSize
                
            ' Flag change ...
            
                ViewportSizeChanged = True
                
        
    End Property
