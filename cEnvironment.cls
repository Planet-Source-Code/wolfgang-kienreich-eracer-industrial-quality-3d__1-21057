VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public D3Frame              As Direct3DRMFrame3
Public D3FrameStation       As Direct3DRMFrame3
Public StationHealth        As Single
Public StationDeactivating  As Single

Public DDSMap               As DirectDrawSurface4

Public IslandSize           As Long                      ' Size of current map
Public IslandName           As String                    ' Name of current map
Public DayTime              As eDayTime                  ' Current time of day in environment

Private I_oD3FrameWater     As Direct3DRMFrame3          ' D3DRM Water frame
Private I_oD3FrameSky       As Direct3DRMFrame3          ' D3DRM Sky frame
Private I_oD3FrameFloor     As Direct3DRMFrame3          ' D3DRM Floor frame

Private I_oD3FrameLight     As Direct3DRMFrame3          ' D3DRM Light frame

Private I_oD3FrameFlare()   As Direct3DRMFrame3          ' D3DRM lensflare frames

Private I_oD3MeshWater      As Direct3DRMMeshBuilder3    ' D3DRM Mesh holding water
Private I_oD3MeshSky        As Direct3DRMMeshBuilder3    ' D3DRM Mesh holding sky

Private I_oD3MeshFloor      As Direct3DRMMeshBuilder3    ' D3DRM Mesh holding floor

Private I_oD3TFloor()       As Direct3DRMTexture3         ' D3DRM Textures for floor

Private I_oD3LightAmb       As Direct3DRMLight           ' D3DRM Light, environment
Private I_oD3LightDir       As Direct3DRMLight           ' D3DRM Light, directional

Private I_oDSBDie           As DirectSoundBuffer
Private I_oD3BDie           As DirectSound3DBuffer

Public Sub Initialize()

    Dim L_oD3Texture    As Direct3DRMTexture3       ' Texture for local generation
    Dim L_dD3MatOvr     As D3DRMMATERIALOVERRIDE    ' Override for alpha
    Dim L_nRunT         As Long
    
    ' Create base frame
    Set D3Frame = Application.D3Instance.CreateFrame(Application.D3Frame)
    
    ' Create D3DRM frame for light: Still frame
    Set I_oD3FrameLight = Application.D3Instance.CreateFrame(D3Frame)
       
    ' Initialize light frame
    I_oD3FrameLight.SetOrientation Nothing, 0, -1, -1, 0, 1, 0
    
    ' Create and add ambient light
    Set I_oD3LightAmb = Application.D3Instance.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.5, 0.5, 0.5)
    I_oD3FrameLight.AddLight I_oD3LightAmb
    
    ' Create and add directional light
    Set I_oD3LightDir = Application.D3Instance.CreateLightRGB(D3DRMLIGHT_DIRECTIONAL, 0.75, 0.75, 0.75)
    I_oD3FrameLight.AddLight I_oD3LightDir
       
    ' Create D3DRM frame for Sky: Moves with racer (smooth)
    Set I_oD3FrameSky = Application.D3Instance.CreateFrame(Application.D3Frame)
       
    ' Create D3DRM frame for water: Still frame
    Set I_oD3FrameWater = Application.D3Instance.CreateFrame(D3Frame)
                                             
    ' Create D3DRM frame for Station: Players camp
    Set D3FrameStation = Application.D3Instance.CreateFrame(D3Frame)
    
    ' Water
    Set I_oD3MeshWater = Application.D3Instance.CreateMeshBuilder
    I_oD3MeshWater.LoadFromFile App.Path + "\mdl\w.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    I_oD3MeshWater.Optimize
    I_oD3MeshWater.SetAppData 1 ' Remember for identification on fly-over

    I_oD3FrameWater.AddVisual I_oD3MeshWater
    
    ' Sky
    Set I_oD3MeshSky = Application.D3Instance.CreateMeshBuilder
    I_oD3MeshSky.LoadFromFile App.Path + "\mdl\s.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    I_oD3MeshSky.Optimize
    I_oD3FrameSky.AddVisual I_oD3MeshSky

    ' Floor
    Set I_oD3FrameFloor = Application.D3Instance.CreateFrame(D3Frame)
    Set I_oD3MeshFloor = Application.D3Instance.CreateMeshBuilder
    I_oD3FrameFloor.AddVisual I_oD3MeshFloor
    I_oD3MeshFloor.SetAppData 2 ' Remember for identification on fly-over
    
    ' Initialize lensflare system ...
    
        ' Dimension flare frames
        ReDim I_oD3FrameFlare(4)
        
        ' Element #1 ...
        
            ' Prepare texture
            Set L_oD3Texture = Application.D3Instance.LoadTexture(App.Path + "\gfx\f01.bmp")
            L_oD3Texture.SetDecalTransparency D_TRUE
            L_oD3Texture.SetDecalTransparentColor 0
            
            ' Create and add geometry
            Set I_oD3FrameFlare(0) = Application.D3Instance.CreateFrame(Application.D3Frame)
            I_oD3FrameFlare(0).AddVisual GenerateBaseMesh(0.75, L_oD3Texture)
        
        ' Element #2 ...
        
            ' Prepare texture
            Set L_oD3Texture = Application.D3Instance.LoadTexture(App.Path + "\gfx\f02.bmp")
            L_oD3Texture.SetDecalTransparency D_TRUE
            L_oD3Texture.SetDecalTransparentColor 0
        
            ' Create and add geometry
            Set I_oD3FrameFlare(1) = Application.D3Instance.CreateFrame(Application.D3Frame)
            I_oD3FrameFlare(1).AddVisual GenerateBaseMesh(0.3, L_oD3Texture)
        
        ' Element #3 ...
        
            ' Prepare texture
            Set L_oD3Texture = Application.D3Instance.LoadTexture(App.Path + "\gfx\f02.bmp")
            L_oD3Texture.SetDecalTransparency D_TRUE
            L_oD3Texture.SetDecalTransparentColor 0
        
            ' Create and add geometry
            Set I_oD3FrameFlare(2) = Application.D3Instance.CreateFrame(Application.D3Frame)
            I_oD3FrameFlare(2).AddVisual GenerateBaseMesh(0.05, L_oD3Texture)
         
        ' Element #4 ...
        
            ' Prepare texture
            Set L_oD3Texture = Application.D3Instance.LoadTexture(App.Path + "\gfx\f03.bmp")
            L_oD3Texture.SetDecalTransparency D_TRUE
            L_oD3Texture.SetDecalTransparentColor 0
        
            ' Create and add geometry
            Set I_oD3FrameFlare(3) = Application.D3Instance.CreateFrame(Application.D3Frame)
            I_oD3FrameFlare(3).AddVisual GenerateBaseMesh(0.75, L_oD3Texture)
            
        ' Load floor textures ...
        
            ' Reserve space
            ReDim I_oD3TFloor(4)
            
            ' Load from file
            Set I_oD3TFloor(0) = Application.D3Instance.LoadTexture(App.Path + "\gfx\s2.bmp")
            For L_nRunT = 1 To 4
                Set I_oD3TFloor(L_nRunT) = Application.D3Instance.LoadTexture(App.Path + "\gfx\s2" & L_nRunT & ".bmp")
                I_oD3TFloor(L_nRunT).GenerateMIPMap
            Next
        
        ' Set general data
            DayTime = dtDay
            
        If (Not Application.DSInstance Is Nothing) Then
            Dim L_ddsbd As DSBUFFERDESC
            With L_ddsbd
                .lBufferBytes = 0
                .lFlags = DSBCAPS_CTRL3D Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
            End With
            Set I_oDSBDie = Application.DSInstance.CreateSoundBufferFromFile(App.Path + "\snd\explobig.wav", L_ddsbd, GetWaveFileFormat(App.Path + "\snd\explobig.wav"))
            Set I_oD3BDie = I_oDSBDie.GetDirectSound3DBuffer
            With I_oD3BDie
                .SetConeAngles DS3D_MINCONEANGLE, DS3D_MAXCONEANGLE, DS3D_IMMEDIATE
                .SetConeOutsideVolume 0, DS3D_IMMEDIATE
                .SetMinDistance 1, DS3D_IMMEDIATE
                .SetMaxDistance 15, DS3D_IMMEDIATE
                .SetVelocity 0, 0, 0, DS3D_IMMEDIATE
                .SetMode DS3DMODE_NORMAL, DS3D_IMMEDIATE
            End With
        End If
        
        ' Finish ...
        
            ChangeMap
            Reset
        
End Sub

Public Sub ChangeMap()

    Dim L_sMapFile      As String                   ' Map file to load
    Dim L_nMapCount     As Long                     ' Counter over maps
    Static S_nMapNumber As Long
    
    S_nMapNumber = S_nMapNumber + 1
    L_sMapFile = Dir(App.Path + "\map\" & S_nMapNumber & "*.map")
    If L_sMapFile = "" Then
        S_nMapNumber = 1
        L_sMapFile = Dir(App.Path + "\map\" & S_nMapNumber & "*.map")
    End If
    
    IslandName = Left(L_sMapFile, Len(L_sMapFile) - 4)
    
    LoadMap

End Sub

Public Sub ChangeDaytime()
    DayTime = IIf(DayTime = dtDay, dtNight, dtDay)
End Sub

Public Sub Reset()

    Dim L_nIndex As Long
    Dim L_nHeight As Single
    Dim L_dD3Pos As D3DVECTOR
    
    
    
    ' Day
    If DayTime = dtDay Then
    
        ' Sky
        I_oD3MeshSky.SetTexture Application.D3Instance.LoadTexture(App.Path + "\gfx\s5.bmp")
        
        ' Water
        I_oD3MeshWater.SetTexture Application.D3Instance.LoadTexture(App.Path + "\gfx\s1.bmp")
        
        ' Flares
        For L_nIndex = 0 To 3
            I_oD3FrameFlare(L_nIndex).SetTraversalOptions D3DRMFRAME_RENDERENABLE
        Next
        
    End If
    
    ' Night
    If DayTime = dtNight Then
        
        ' Sky
        I_oD3MeshSky.SetTexture Application.D3Instance.LoadTexture(App.Path + "\gfx\s6.bmp")
        
        ' Water
        I_oD3MeshWater.SetTexture Application.D3Instance.LoadTexture(App.Path + "\gfx\s7.bmp")
        
        ' Flares
        For L_nIndex = 0 To 3
            I_oD3FrameFlare(L_nIndex).SetTraversalOptions 0
        Next
        
    End If
    
' Set lightstates according to daytime ...

    ' Day
    If DayTime = dtDay Then
        I_oD3LightAmb.SetColorRGB 0.3, 0.3, 0.3
        I_oD3LightDir.SetColorRGB 0.6, 0.6, 0.6
    End If
    
    ' Night
    If DayTime = dtNight Then
        I_oD3LightAmb.SetColorRGB 0.2, 0.15, 0.2
        I_oD3LightDir.SetColorRGB 0, 0, 0
    End If
    
    ' Station
    StationHealth = 2000
    StationDeactivating = 0
    D3FrameStation.SetRotation Nothing, 0, 0, 0, 0
    D3FrameStation.SetOrientation Nothing, 0, 0, 1, 0, 1, 0
    I_oD3MeshFloor.GetVertex (IslandSize \ 2) * IslandSize + IslandSize \ 2, L_dD3Pos
    L_nHeight = L_dD3Pos.Y
    D3FrameStation.SetPosition Nothing, (IslandSize \ 2), L_nHeight - 0.2, (IslandSize \ 2)
    
End Sub

Public Sub Update()

        Dim L_dD3Pos        As D3DVECTOR
        Dim L_dD3Ori        As D3DVECTOR
        Dim L_dD3Nor        As D3DVECTOR
        Dim L_nIdx          As Long
        Dim L_dD3MatOvr     As D3DRMMATERIALOVERRIDE
        
            ' Update water animation ...
            
                ' Raise/lower
                I_oD3FrameWater.GetPosition Nothing, L_dD3Pos
                L_dD3Pos.Y = Sin(Application.FrameCount / 10) * 0.05
                I_oD3FrameWater.SetPosition Nothing, L_dD3Pos.X, L_dD3Pos.Y, L_dD3Pos.Z
                   
                 ' Animate texture
                I_oD3MeshWater.SetTextureCoordinates 0, Sin(Application.FrameCount / 50), Cos(Application.FrameCount / 50)
                I_oD3MeshWater.SetTextureCoordinates 1, 200 + Sin(Application.FrameCount / 50), Cos(Application.FrameCount / 50)
                I_oD3MeshWater.SetTextureCoordinates 2, 200 + Sin(Application.FrameCount / 50), 200 + Cos(Application.FrameCount / 50)
                I_oD3MeshWater.SetTextureCoordinates 3, Sin(Application.FrameCount / 50), 200 + Cos(Application.FrameCount / 50)
                
            ' Update lensflare system ...
                     
                ' Get camera data
                Application.D3Camera.GetPosition Nothing, L_dD3Pos
                Application.D3Camera.GetOrientation Nothing, L_dD3Ori, L_dD3Nor
                
                ' Set flare frames position accordingly
                I_oD3FrameFlare(0).SetPosition Nothing, L_dD3Pos.X + L_dD3Ori.X * 6, L_dD3Pos.Y + 0.2 + (L_dD3Ori.Y) * 3, L_dD3Pos.Z + 0.5
                I_oD3FrameFlare(1).SetPosition Nothing, L_dD3Pos.X + L_dD3Ori.X * 4, L_dD3Pos.Y + 0.2 + (L_dD3Ori.Y) * 2, L_dD3Pos.Z + 0.65
                I_oD3FrameFlare(2).SetPosition Nothing, L_dD3Pos.X + L_dD3Ori.X * 2, L_dD3Pos.Y + 0.2 + (L_dD3Ori.Y) * 1, L_dD3Pos.Z + 0.75
                I_oD3FrameFlare(3).SetPosition Nothing, L_dD3Pos.X, 5, L_dD3Pos.Z + 10
    
                ' Make flare frames face camera
               I_oD3FrameFlare(0).SetOrientation Nothing, L_dD3Ori.X, L_dD3Ori.Y, L_dD3Ori.Z, L_dD3Nor.X, L_dD3Nor.Y, L_dD3Nor.Z
               I_oD3FrameFlare(1).SetOrientation Nothing, L_dD3Ori.X, L_dD3Ori.Y, L_dD3Ori.Z, L_dD3Nor.X, L_dD3Nor.Y, L_dD3Nor.Z
               I_oD3FrameFlare(2).SetOrientation Nothing, L_dD3Ori.X, L_dD3Ori.Y, L_dD3Ori.Z, L_dD3Nor.X, L_dD3Nor.Y, L_dD3Nor.Z
               I_oD3FrameFlare(3).SetOrientation Nothing, L_dD3Ori.X, L_dD3Ori.Y, L_dD3Ori.Z, L_dD3Nor.X, L_dD3Nor.Y, L_dD3Nor.Z

                ' Set flare brightness
                L_dD3MatOvr.lFlags = D3DRMMATERIALOVERRIDE_DIFFUSE_ALPHAONLY Or D3DRMMATERIALOVERRIDE_DIFFUSE_ALPHAMULTIPLY
                L_dD3MatOvr.dcDiffuse.a = 0.1 * (1 - Abs(L_dD3Ori.X) * 1.5)
                I_oD3FrameFlare(0).SetMaterialOverride L_dD3MatOvr
                L_dD3MatOvr.dcDiffuse.a = 0.15 * (1 - Abs(L_dD3Ori.X) * 1.5)
                I_oD3FrameFlare(1).SetMaterialOverride L_dD3MatOvr
                I_oD3FrameFlare(2).SetMaterialOverride L_dD3MatOvr

            ' Update according to racer position ...
                    
                ' Camera and sky
                If Application.Environment.StationDeactivating = 0 Then
                    If Application.Player.Deactivating > 0 Then Application.Player.Heading = Application.Player.Heading + 0.05
                    Application.D3Camera.SetPosition Nothing, Application.Player.X - 2 * Cos(Application.Player.Heading), 2 + 1.5 + -Application.Player.Y * 0.25 + IIf(Application.Player.Deactivating > 0, (100 - Application.Player.Deactivating) * 0.025, 0), Application.Player.Z - 2 * Sin(Application.Player.Heading)
                    Application.D3Camera.LookAt Application.Player.D3Frame, Nothing, D3DRMCONSTRAIN_Z
                    I_oD3FrameSky.SetPosition Nothing, Application.Player.X, 0, Application.Player.Z
                Else
                    Application.Environment.D3FrameStation.GetPosition Nothing, L_dD3Pos
                    Application.D3Camera.SetPosition Nothing, L_dD3Pos.X + 3 * Cos(StationDeactivating / 50), 5, L_dD3Pos.Z + 3 * Sin(StationDeactivating / 50)
                    Application.D3Camera.LookAt Application.Environment.D3FrameStation, Nothing, D3DRMCONSTRAIN_Z
                    I_oD3FrameSky.SetPosition Nothing, L_dD3Pos.X, 0, L_dD3Pos.Z
                End If
                
                ' Listener
                If (Not Application.DSInstance Is Nothing) And Application.SoundEnabled Then
                    With Application.DSListener
                        .SetPosition L_dD3Pos.X, L_dD3Pos.Y, L_dD3Pos.Z, DS3D_IMMEDIATE
                        .SetOrientation L_dD3Ori.X, L_dD3Ori.Y, L_dD3Ori.Z, L_dD3Nor.X, L_dD3Nor.Y, L_dD3Nor.Z, DS3D_IMMEDIATE
                    End With
                End If
                
            ' Deactivating procedures ...
                If StationDeactivating > 0 Then
                
                    ' Play explo
                    If StationDeactivating Mod 40 = 0 Then
                        If (Not Application.DSInstance Is Nothing) And Application.SoundEnabled Then
                            D3FrameStation.GetPosition Nothing, L_dD3Pos
                            I_oD3BDie.SetPosition L_dD3Pos.X, L_dD3Pos.Y, L_dD3Pos.Z, DS3D_IMMEDIATE
                            I_oD3BDie.SetConeOrientation 0, 1, 0, DS3D_IMMEDIATE
                            I_oDSBDie.Play DSBPLAY_DEFAULT
                        End If
                    End If
                    
                    ' Rock station
                    If StationDeactivating Mod 10 = 0 Then
                        D3FrameStation.GetPosition Nothing, L_dD3Pos
                        L_dD3Pos.Y = L_dD3Pos.Y - 0.005
                        D3FrameStation.SetPosition Nothing, L_dD3Pos.X, L_dD3Pos.Y, L_dD3Pos.Z
                        D3FrameStation.SetRotation Nothing, Rnd, Rnd, Rnd, Rnd * 0.005
                    End If
                                    
                    ' Place particles
                    If StationDeactivating Mod 5 = 0 Then
                        D3FrameStation.GetPosition Nothing, L_dD3Pos
                        With L_dD3Ori
                            .X = Rnd * 0.01
                            .Y = 0.075
                            .Z = Rnd * 0.01
                        End With
                        With L_dD3Pos
                            .X = .X + Rnd * 2 - 1
                            .Y = .Y + 3
                            .Z = .Z + Rnd * 2 - 1
                        End With
                        Application.Particles.Add ptDustDay, L_dD3Pos, L_dD3Ori
                    End If
                
                    ' Place explosions
                    If StationDeactivating Mod 5 = 0 Then
                        D3FrameStation.GetPosition Nothing, L_dD3Pos
                        With L_dD3Ori
                            .X = 0
                            .Y = 0
                            .Z = 0
                        End With
                        With L_dD3Pos
                            .X = .X + Rnd * 3 - 1.5
                            .Y = .Y + Rnd * 2 + 2
                            .Z = .Z + Rnd * 3 - 1.5
                        End With
                        Application.Effects.Add etExplo, L_dD3Pos, L_dD3Ori
                    End If
                    
                    StationDeactivating = StationDeactivating - 1
                    If StationDeactivating = 0 Then
                        Application.Enemies.Active = False
                        Application.Reset
                    End If
                                        
                    Exit Sub
                    
                Else
                    StationHealth = StationHealth + 1
                    If StationHealth > 2000 Then StationHealth = 2000
                End If
                
            ' Check for station damage
                If StationHealth < 0 Then
                    StationDeactivating = 200
                    Application.Interface.DisplayMessage "station destroyed ..."
                End If

End Sub

Private Sub LoadMap()

' Declare local variables ...

    Dim L_oD3Face       As Direct3DRMFace2          ' Face, temporary holds data for creation of floor
    Dim L_oD3Material   As Direct3DRMMaterial2      ' Material for faces
    Dim L_dD3MatOvr     As D3DRMMATERIALOVERRIDE    ' MAterial override for map display
    
    Dim L_dD3Vertex(3)  As D3DVECTOR                ' Vertex array for terrain creation
    Dim L_dD3Normal     As D3DVECTOR                ' Normal array for terrain creation
    Dim L_nRunX         As Long                     ' Runs over X coords during ground mesh generation
    Dim L_nRunZ         As Long                     ' Runs over Z coords during ground mesh generation
    Dim L_nRunI         As Long                     ' Runs over various arrays
    Dim L_nHeight       As Single                   ' Height of point being processed
    
    Dim L_bMap()        As Byte                     ' Holds loaded map data
    Dim L_dDDSD         As DDSURFACEDESC2           ' Surface description for map surface
    Dim L_dArea         As RECT                     ' Rectangle for locking
    Dim L_dDDCK         As DDCOLORKEY               ' Color key for making map transparent
    
' Code ...
    ' Generate floor data from file ...
                            
        ' Load map into surface
        With L_dDDSD
            .lFlags = DDSD_CAPS
            .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        End With
        Set DDSMap = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\map\" + IslandName + ".map", L_dDDSD)
        
        ' Get map properties
        DDSMap.GetSurfaceDesc L_dDDSD
        
        ' Lock surface, retrieve data
        With L_dArea
            .Right = L_dDDSD.lWidth
            .Bottom = L_dDDSD.lHeight
        End With
        DDSMap.Lock L_dArea, L_dDDSD, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, ByVal 0&
        DDSMap.GetLockedArray L_bMap
        Application.Environment.IslandSize = L_dArea.Right
        
    ' Rebuild floor mesh ...
    
        ' Create floor material
        Set L_oD3Material = Application.D3Instance.CreateMaterial(5)
        L_oD3Material.SetAmbient 1, 1, 1
        
        ' Empty mesh
        I_oD3MeshFloor.Empty
        
        ' Add vertices ...
        
            L_nRunI = 0
            For L_nRunX = 0 To L_dArea.Right - 1
                For L_nRunZ = 0 To L_dArea.Bottom - 1
                    L_nHeight = ((L_bMap(L_nRunX * 2, L_nRunZ) And 31) / 32)
                    I_oD3MeshFloor.AddVertex L_nRunX, L_nHeight, L_nRunZ
                Next
            Next
         
        ' Add dummy normal
        I_oD3MeshFloor.AddNormal 0, 1, 0
        
        ' Add faces ...
        
            For L_nRunX = 0 To L_dArea.Right - 2
                For L_nRunZ = 0 To L_dArea.Bottom - 2
                    
                    ' Check if vertices demand face
                    I_oD3MeshFloor.GetVertex L_nRunX + L_nRunZ * L_dArea.Right, L_dD3Vertex(0)
                    I_oD3MeshFloor.GetVertex L_nRunX + 1 + L_nRunZ * L_dArea.Right, L_dD3Vertex(1)
                    I_oD3MeshFloor.GetVertex L_nRunX + 1 + (L_nRunZ + 1) * L_dArea.Right, L_dD3Vertex(2)
                    I_oD3MeshFloor.GetVertex L_nRunX + (L_nRunZ + 1) * L_dArea.Right, L_dD3Vertex(3)
                    
                    ' Yes: Create face
                    If (L_dD3Vertex(0).Y + L_dD3Vertex(1).Y + L_dD3Vertex(2).Y + L_dD3Vertex(3).Y) > 0 Then
                        
                        ' Set geometry
                        Set L_oD3Face = I_oD3MeshFloor.CreateFace
                        L_oD3Face.AddVertexAndNormalIndexed L_nRunX + L_nRunZ * L_dArea.Right, 0
                        L_oD3Face.AddVertexAndNormalIndexed L_nRunX + 1 + L_nRunZ * L_dArea.Right, 0
                        L_oD3Face.AddVertexAndNormalIndexed L_nRunX + 1 + (L_nRunZ + 1) * L_dArea.Right, 0
                        L_oD3Face.AddVertexAndNormalIndexed L_nRunX + (L_nRunZ + 1) * L_dArea.Right, 0
                        
                        ' Set material
                        L_oD3Face.SetMaterial L_oD3Material

                    End If
                    
                Next
            Next
                        
        ' Adjust vertices and waterline ...
        
            For L_nRunI = 0 To I_oD3MeshFloor.GetVertexCount - 1
                I_oD3MeshFloor.GetVertex L_nRunI, L_dD3Vertex(0)
                If L_dD3Vertex(0).Y < 0.15 Then
                    L_dD3Vertex(0).Y = -0.15
                Else
                    L_dD3Vertex(0).Y = L_dD3Vertex(0).Y * 2
                End If
                I_oD3MeshFloor.SetVertex L_nRunI, L_dD3Vertex(0).X, L_dD3Vertex(0).Y, L_dD3Vertex(0).Z
            Next
        
        ' Generate normals
        I_oD3MeshFloor.GenerateNormals 1, D3DRMGENERATENORMALS_PRECOMPACT Or D3DRMGENERATENORMALS_USECREASEANGLE
                            
        ' NOTE: Only now can we apply texture coordinates and mappings,
        ' as applying different coords and mappings to faces sharing
        ' vertices duplicates that vertices, rendering normal calculation
        ' impossible. So the steps are: Build mesh using shared vertices,
        ' let normals be calculated, texture mesh (now that the normals
        ' are known it doesnt matter if the vertices are duplicated!)
            
        ' Set surfaces ...
        For L_nRunI = 0 To I_oD3MeshFloor.GetFaceCount - 1
            With I_oD3MeshFloor.GetFace(L_nRunI)
                                                
                ' Acquire position
                .GetVertex 0, L_dD3Vertex(0), L_dD3Normal
                L_nRunX = L_dD3Vertex(0).X
                L_nRunZ = L_dD3Vertex(0).Z
                
                'Set texture coordinates
                .SetTextureCoordinates 0, L_nRunX, L_nRunZ
                .SetTextureCoordinates 1, L_nRunX + 1, L_nRunZ
                .SetTextureCoordinates 2, L_nRunX + 1, L_nRunZ + 1
                .SetTextureCoordinates 3, L_nRunX, L_nRunZ + 1

                ' Set text ure
                 If L_dD3Vertex(0).Y < 0.25 * 2 Or Rnd > 0.2 Then
                    .SetTexture I_oD3TFloor(0)
                Else
                    .SetTexture I_oD3TFloor(Int(Rnd * 4) + 1)
                End If
            
            End With
        Next
        
    ' Finish ...
    
        ' Remember height at center
        I_oD3MeshFloor.GetVertex (L_dArea.Bottom \ 2) * L_dArea.Right + L_dArea.Right \ 2, L_dD3Normal
        L_nHeight = L_dD3Normal.Y

        ' Unlock surface
        DDSMap.Unlock L_dArea

        ' Make surface transparent
        With L_dDDCK
            .low = 0
            .high = 0
        End With
        DDSMap.SetColorKey DDCKEY_SRCBLT, L_dDDCK
          
        ' Set Station pos
        D3Frame.DeleteChild D3FrameStation
        Set D3FrameStation = Application.D3Instance.CreateFrame(D3Frame)
        D3FrameStation.SetPosition Nothing, 0, 0, 0
        D3FrameStation.LoadFromFile App.Path + "\mdl\st1.x", "", D3DRMLOAD_FROMFILE Or D3DRMLOAD_FIRST, Nothing, Nothing
        D3FrameStation.LoadFromFile App.Path + "\mdl\st2.x", "", D3DRMLOAD_FROMFILE Or D3DRMLOAD_FIRST, Nothing, Nothing
        D3FrameStation.GetChildren.GetElement(1).SetRotation Nothing, 0, 1, 0, 0.02
        D3FrameStation.SetPosition Nothing, L_dArea.Right \ 2, L_nHeight - 0.2, L_dArea.Bottom \ 2

End Sub
