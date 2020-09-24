VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OPTION SETTINGS ...

    ' Enforce variable declarations
    Option Explicit

' PUBLIC PROPERTIES ...

    Public D3Frame              As Direct3DRMFrame3         ' Master frame representing instance
    
    Public Deactivating         As Single                   ' Deactivation counter
    
    Public Heading              As Single                   ' Current heading in radians
    Public Velocity             As Single                   ' Current velocity
    Public Health               As Single                   ' Health
    Public Hoover               As Single                   ' Amount of hoover power applied
    Public Fuel                 As Single                   ' Fuel left
    Public Model                As eRacerModel              ' Model of racer in use

    Public X                    As Single                   ' Position in game world
    Public Y                    As Single                   '
    Public Z                    As Single                   '

' PRIVATE VARIABLES ...

    Private I_nFallspeed        As Single                   ' Speed of vertical position change
    Private I_nBanking          As Single                   ' Speed of heading change
    Private I_nAccelleration    As Single                   ' Accelleration
    Private I_oD3MeshRacer1     As Direct3DRMMeshBuilder3   ' Resource mesh: Racer #1
    Private I_oD3MeshRacer2     As Direct3DRMMeshBuilder3   ' Resource mesh: Racer #2
    Private I_oD3MeshLamps1     As Direct3DRMMeshBuilder3   ' Resource mesh: Racer lamp #1
    Private I_oD3MeshLamps2     As Direct3DRMMeshBuilder3   ' Resource mesh: Racer lamp #2
    Private I_oD3Light          As Direct3DRMLight          ' Light for racer lamp

    Private I_oDSB              As DirectSoundBuffer
    Private I_oD3B              As DirectSound3DBuffer

    Private I_oDSBShot          As DirectSoundBuffer
    Private I_oD3BShot          As DirectSound3DBuffer
    
    Private I_oDSBDie           As DirectSoundBuffer
    Private I_oD3BDie           As DirectSound3DBuffer
    
    
' CODE ...

    Public Sub ToggleSound()
        If Application.SoundEnabled Then
            I_oDSB.Play DSBPLAY_LOOPING
        Else
            I_oDSB.Stop
        End If
    End Sub

    Public Sub Initialize()
        
        Set D3Frame = Application.D3Instance.CreateFrame(Application.D3Frame)
        
        Set I_oD3MeshRacer1 = Application.D3Instance.CreateMeshBuilder
        I_oD3MeshRacer1.LoadFromFile App.Path + "\mdl\r1.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        I_oD3MeshRacer1.Optimize
                        
        Set I_oD3MeshRacer2 = Application.D3Instance.CreateMeshBuilder
        I_oD3MeshRacer2.LoadFromFile App.Path + "\mdl\r2.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        I_oD3MeshRacer2.Optimize
                        
        Set I_oD3MeshLamps1 = Application.D3Instance.CreateMeshBuilder
        I_oD3MeshLamps1.LoadFromFile App.Path + "\mdl\l1.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        I_oD3MeshLamps1.Optimize
                        
        Set I_oD3MeshLamps2 = Application.D3Instance.CreateMeshBuilder
        I_oD3MeshLamps2.LoadFromFile App.Path + "\mdl\l2.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        I_oD3MeshLamps2.Optimize
                     
        Model = rmJaguar
        
    If Not Application.DSInstance Is Nothing Then
        
        Dim L_ddsbd As DSBUFFERDESC
        With L_ddsbd
            .lBufferBytes = 0
            .lFlags = DSBCAPS_CTRL3D Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
        End With
        
        Set I_oDSB = Application.DSInstance.CreateSoundBufferFromFile(App.Path + "\snd\enginep.wav", L_ddsbd, GetWaveFileFormat(App.Path + "\snd\enginep.wav"))
        Set I_oD3B = I_oDSB.GetDirectSound3DBuffer
        With I_oD3B
            .SetConeAngles DS3D_MINCONEANGLE, DS3D_MAXCONEANGLE, DS3D_IMMEDIATE
            .SetConeOutsideVolume 0, DS3D_IMMEDIATE
            .SetMinDistance 1, DS3D_IMMEDIATE
            .SetMaxDistance 15, DS3D_IMMEDIATE
            .SetVelocity 0, 0, 0, DS3D_IMMEDIATE
            .SetMode DS3DMODE_NORMAL, DS3D_IMMEDIATE
        End With
        
        Set I_oDSBShot = Application.DSInstance.CreateSoundBufferFromFile(App.Path + "\snd\shotp.wav", L_ddsbd, GetWaveFileFormat(App.Path + "\snd\shotp.wav"))
        Set I_oD3BShot = I_oDSBShot.GetDirectSound3DBuffer
        With I_oD3BShot
            .SetConeAngles DS3D_MINCONEANGLE, DS3D_MAXCONEANGLE, DS3D_IMMEDIATE
            .SetConeOutsideVolume 0, DS3D_IMMEDIATE
            .SetMinDistance 1, DS3D_IMMEDIATE
            .SetMaxDistance 10, DS3D_IMMEDIATE
            .SetVelocity 0, 0, 0, DS3D_IMMEDIATE
            .SetMode DS3DMODE_NORMAL, DS3D_IMMEDIATE
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
        
                                
        Reset
        
    End Sub

    Public Sub ChangeRacer()
        Model = IIf(Model = rmJaguar, rmWarthog, rmJaguar)
    End Sub
    
    Public Sub Reset()
    
        Heading = 1.7 * Atn(1)
        Velocity = 0
        Hoover = 0
        I_nBanking = 0
        I_nAccelleration = 0
        Fuel = IIf(Model = rmWarthog, 9000, 5000)
        Health = IIf(Model = rmWarthog, 1000, 700)
        Deactivating = 0
        
        Application.D3Frame.DeleteChild D3Frame
        Set D3Frame = Application.D3Instance.CreateFrame(Application.D3Frame)
        D3Frame.SetPosition Nothing, Application.Environment.IslandSize / 2 - 3, 7, Application.Environment.IslandSize / 2
        
        D3Frame.AddVisual IIf(Model = rmWarthog, I_oD3MeshRacer1, I_oD3MeshRacer2)
        
        If Application.Environment.DayTime = dtNight Then
            Set I_oD3Light = Application.D3Instance.CreateLightRGB(D3DRMLIGHT_SPOT, 1, 1, 1)
            With I_oD3Light
                .SetLinearAttenuation 0.1
                .SetRange 10
                .SetUmbra 0.6
                .SetPenumbra 0.8
            End With
            D3Frame.AddLight I_oD3Light
            D3Frame.AddVisual IIf(Model = rmWarthog, I_oD3MeshLamps1, I_oD3MeshLamps2)
        End If
        
        If (Not Application.DSInstance Is Nothing) And Application.SoundEnabled Then
            I_oDSB.Play DSBPLAY_LOOPING
        End If
        
    End Sub
    
    Public Sub Update()
    
        ' Declare local variables ...
        
            Dim L_dD3Pos              As D3DVECTOR            ' Holds position of racer
            Dim L_dD3Ori              As D3DVECTOR            ' Holds orientation of racer
            Dim L_dD3Nor              As D3DVECTOR            ' Holds normal of racer
            
            Dim L_dD3PPos             As D3DVECTOR            ' Holds temporary calculation vector
            Dim L_dD3PVel             As D3DVECTOR            ' Holds temporary calculation vector
            Dim L_dD3Tmp              As D3DVECTOR            ' Holds temporary calculation vector
                        
            Dim L_nIdx              As Long                 ' Runs over various arrays
            Dim L_nIdx2             As Long                 ' Runs over various arrays
            Dim L_nTmp              As Single               ' Helper variable for calculations
            
            Dim L_nGroundType       As Long                 ' Type of ground racer is on
            
            Dim L_dD3Ray            As D3DRMRAY             ' Ray for picking
            Dim L_oD3PDA            As Direct3DRMPick2Array ' Result of picking
            Dim L_dD3PD             As D3DRMPICKDESC2       ' Result of picking
            Dim L_oD3Visual         As Direct3DRMVisual     ' Visual picked
            Dim L_dD3MatOvr         As D3DRMMATERIALOVERRIDE ' Material override for flare blending
            Dim L_oD3Texture        As Direct3DRMTexture3   ' Texture for animation management
            Dim L_dArea             As RECT                 ' Blitting area
            Dim L_dDummy(0)         As RECT                 ' Dummy rects for changed texture
            
            Static S_nShotDelay     As Long                 ' Delay until next shot to be fired
            
            Dim L_bColliding        As Boolean
            Dim L_nAltitude         As Single
            
        ' Code ...
        
            ' Gather racer data ...
            
                ' Gather basic data ...
                    
                    D3Frame.GetPosition Nothing, L_dD3Pos
                    D3Frame.GetOrientation Nothing, L_dD3Ori, L_dD3Nor
                
                    X = L_dD3Pos.X
                    Y = L_dD3Pos.Y
                    Z = L_dD3Pos.Z
                                                        
                ' Deactivation procedures ...
   
                    If Deactivating > 0 Then
                    
                        ' Place explosions
                        If Deactivating Mod 5 = 0 Then
                            With L_dD3Ori
                                .X = 0
                                .Y = 0.05
                                .Z = 0
                            End With
                            With L_dD3Pos
                                .X = .X + Rnd * 0.8 - 0.4
                                .Y = .Y + Rnd * 0.8 - 0.4
                                .Z = .Z + Rnd * 0.8 - 0.4
                            End With
                            Application.Effects.Add etExplo, L_dD3Pos, L_dD3Ori
                        End If
                        
                        ' Place dust
                        If Deactivating Mod 5 = 0 Then
                            With L_dD3Ori
                                .X = Rnd * 0.01
                                .Y = 0.09
                                .Z = Rnd * 0.01
                            End With
                            With L_dD3Pos
                                .X = .X
                                .Y = .Y
                                .Z = .Z
                            End With
                            Application.Particles.Add ptDustDay, L_dD3Pos, L_dD3Ori
                        End If
                        
                        ' Place rotation
                        If Deactivating Mod 10 = 0 Then
                            D3Frame.SetRotation Nothing, Rnd, Rnd, Rnd, Rnd * 0.1
                        End If
                        
                        ' Count towards total deactivation
                        Deactivating = Deactivating - 1
                        If Deactivating = 0 Then
                            Application.Enemies.Active = False
                            Application.Reset
                        End If
                        Exit Sub
                        
                    End If
        
                ' Gather racer collission ...
                    
                    ' Reset colission
                    L_bColliding = False
                    
                    ' Prepare ray
                    Application.DXInstance.VectorCopy L_dD3Ray.pos, L_dD3Pos
                    L_dD3Ray.Dir.X = L_dD3Ori.X
                    L_dD3Ray.Dir.Y = 0.25
                    L_dD3Ray.Dir.Z = L_dD3Ori.Z
                        
                    ' Cast ray
                    Set L_oD3PDA = Application.Environment.D3Frame.RayPick(Nothing, L_dD3Ray, D3DRMRAYPICK_IGNOREFURTHERPRIMITIVES)
                         
                    ' Retrieve results
                    If Not (L_oD3PDA.GetSize = 0) Then
                        If L_oD3PDA.GetPickFrame(0, L_dD3PD).GetSize > 0 Then
                            Set L_oD3Visual = L_oD3PDA.GetPickVisual(0, L_dD3PD)
                            Application.DXInstance.VectorSubtract L_dD3Tmp, L_dD3Ray.pos, L_dD3PD.vPostion
                            L_bColliding = (Application.DXInstance.VectorModulus(L_dD3Tmp) < 1.5)
                        End If
                    End If
                        
                ' Gather floor below racer ...
                
                    ' Prepare ray
                    Application.DXInstance.VectorCopy L_dD3Ray.pos, L_dD3Pos
                    L_dD3Ray.pos.Y = 10
                    L_dD3Ray.Dir.X = 0
                    L_dD3Ray.Dir.Y = -1
                    L_dD3Ray.Dir.Z = 0
                     
                    ' Cast ray
                    Set L_oD3PDA = Application.Environment.D3Frame.RayPick(Nothing, L_dD3Ray, D3DRMRAYPICK_IGNOREFURTHERPRIMITIVES Or D3DRMRAYPICK_INTERPOLATENORMAL)
    
                    ' Retrieve results
                    If Not (L_oD3PDA.GetSize = 0) Then
                        If L_oD3PDA.GetPickFrame(0, L_dD3PD).GetSize > 0 Then
                            Set L_oD3Visual = L_oD3PDA.GetPickVisual(0, L_dD3PD)
                            L_nGroundType = L_oD3Visual.GetAppData
                            
                            If L_nGroundType = 1 Then
                                If Not Application.WaterMessage Then
                                    If Not Application.Enemies.Active Then Application.Interface.DisplayMessage "your fighter doesn't handle well on water, watch the green speed bar ..."
                                    Application.WaterMessage = True
                                End If
                            
                            End If
                            
                            L_nAltitude = L_dD3Pos.Y - L_dD3PD.vPostion.Y
                        Else
                            L_nGroundType = 1
                            L_nAltitude = L_dD3Pos.Y
                        End If
                    Else
                         L_nAltitude = L_dD3Pos.Y
                    End If
                
            ' Cast particles ...
                 
                If Rnd > 0.5 And (L_nAltitude < (IIf(L_nGroundType = 1, 0.65, 0.5)) Or L_bColliding) Then
                            
                    With L_dD3PPos
                        .X = L_dD3Pos.X - Cos(Heading) * (Rnd * 0.8 - 0.4)
                        .Y = L_dD3Pos.Y + Rnd * 0.2
                        .Z = L_dD3Pos.Z - Sin(Heading) * (Rnd * 0.8 - 0.4)
                    End With
                    
                    With L_dD3PVel
                        .X = -Cos(Heading) * 0.01
                        .Y = 0.05 + Velocity * 0.5
                        .Z = -Sin(Heading) * 0.01
                    End With
                    
                    If L_nGroundType = 1 Then
                        Application.Particles.Add IIf(Application.Environment.DayTime = dtDay, ptWaterDay, ptWaterNight), L_dD3PPos, L_dD3PVel
                    Else
                        Application.Particles.Add IIf(Application.Environment.DayTime = dtDay, ptDustDay, ptDustNight), L_dD3PPos, L_dD3PVel
                    End If
                    
                End If
                
                   
            ' Update racer data ...
            
                ' Minimum L_nAltitude / collission
                If L_bColliding Then
                    Health = Health - Abs(Velocity * 200)
                    If Health < 10 Then Health = 10
                    Velocity = Velocity - 0.05
                End If
                
                If L_nAltitude < 0.25 Then
                    L_nAltitude = 0.25
                    Hoover = Hoover + 0.005
                End If
                                                                                                       
                ' AGrav dekey
                Hoover = Hoover * 0.95
                
                 ' AGrav dekey above L_nAltitude
                If L_nAltitude > 2 Then Hoover = Hoover * 0.9
                   
                ' Gravity
                 I_nFallspeed = I_nFallspeed - 0.01
                 
                 ' Hover
                I_nFallspeed = I_nFallspeed + (1 / L_nAltitude) * 0.006
                
                 ' AGrav
                I_nFallspeed = I_nFallspeed + Hoover * 0.1
                
                ' Air friction levels falling speed
                I_nFallspeed = I_nFallspeed * 0.95
                
               ' Banking
                Heading = Heading + I_nBanking * IIf(Model = rmJaguar, 1.4, 1.2)
                
                ' Banking dekeys
                I_nBanking = I_nBanking * 0.95
                 
                ' Slowdown over water
                If L_nGroundType = 1 And L_nAltitude < 1.5 Then Velocity = Velocity * IIf(Model = rmJaguar, 0.9, 0.95)
                
                ' Limit Velocity
                Velocity = Velocity * IIf(Velocity > 0, 0.99, 0.7) * IIf(Fuel = 10, 0.75, 1)
                If Velocity > 0.1 + IIf(Model = rmJaguar, 0.025, 0) Then Velocity = 0.1 + IIf(Model = rmJaguar, 0.025, 0)
                
                ' Limit accelleration
                I_nAccelleration = I_nAccelleration * IIf(I_nAccelleration > 0, 0.95, 0.75)
                
                ' Refuel when near station ...
                
                    ' Get station position
                    With L_dD3Tmp
                        .X = Application.Environment.IslandSize \ 2
                        .Y = L_dD3Pos.Y
                        .Z = Application.Environment.IslandSize \ 2
                    End With
                    
                    ' Calculate distance
                    Application.DXInstance.VectorSubtract L_dD3Tmp, L_dD3Tmp, L_dD3Pos
                    
                    ' Refuel when near
                    If Application.DXInstance.VectorModulus(L_dD3Tmp) < 4 Then
                        If Fuel < IIf(Model = rmJaguar, 10000, 16000) - 25 Then
                            Fuel = Fuel + 25
                            If Fuel >= IIf(Model = rmJaguar, 10000, 16000) - 25 Then
                                Application.Interface.DisplayMessage "fuel tanks filled ... "
                            End If
                        End If
                    End If
                    
                ' Die if too much damage
                
                    If Health < 20 Then
                        Deactivating = 100
                        If (Not Application.DSInstance Is Nothing) And Application.SoundEnabled Then
                            I_oD3BDie.SetPosition L_dD3Pos.X, L_dD3Pos.Y, L_dD3Pos.Z, DS3D_IMMEDIATE
                            I_oD3BDie.SetConeOrientation L_dD3Ori.X, L_dD3Ori.Y, L_dD3Ori.Z, DS3D_IMMEDIATE
                            I_oDSBDie.Play DSBPLAY_DEFAULT
                        End If
                        Application.Interface.DisplayMessage "fighter destroyed ... "
                        Exit Sub
                    End If
                    
                ' Die if out of map
                
                    If L_dD3Pos.X > Application.Environment.IslandSize + 10 Or L_dD3Pos.Z > Application.Environment.IslandSize + 10 Or L_dD3Pos.X < -10 Or L_dD3Pos.Z < -10 Then
                        Application.Reset
                        Exit Sub
                    End If
                
            ' Set racer data ...
                            
                ' Calculate Position ...
                
                    L_dD3Pos.X = L_dD3Pos.X + Cos(Heading) * Velocity
                    L_dD3Pos.Y = L_dD3Pos.Y + I_nFallspeed
                    If L_dD3Pos.Y < 0 Then L_dD3Pos.Y = 0
                    L_dD3Pos.Z = L_dD3Pos.Z + Sin(Heading) * Velocity
                         
                ' Calculate Orientation ...
                                
                    ' Orientation
                    L_dD3Ori.Y = I_nFallspeed * 5 - I_nAccelleration * 2.5
                    L_dD3Ori.X = Cos(Heading)
                    L_dD3Ori.Z = Sin(Heading)
                    
                    ' Normal
                    L_dD3Nor.X = I_nBanking * 50 * -Sin(Heading)
                    L_dD3Nor.Y = 1 - I_nFallspeed * 5 + I_nAccelleration * 2.5
                    L_dD3Nor.Z = I_nBanking * 50 * Cos(Heading)
                    
                    ' Normalize data
                    Application.DXInstance.VectorNormalize L_dD3Ori
                    Application.DXInstance.VectorNormalize L_dD3Nor
                
                ' Set values ...
                
                    D3Frame.SetPosition Nothing, L_dD3Pos.X, L_dD3Pos.Y, L_dD3Pos.Z
                    D3Frame.SetOrientation Nothing, L_dD3Ori.X, L_dD3Ori.Y, L_dD3Ori.Z, L_dD3Nor.X, L_dD3Nor.Y, L_dD3Nor.Z
                                     
                                     
                ' Set sound
                If (Not Application.DSInstance Is Nothing) And Application.SoundEnabled Then
                    If Deactivating > 0 Then
                        I_oDSB.Stop
                    Else
                        I_oD3B.SetPosition L_dD3Pos.X, L_dD3Pos.Y, L_dD3Pos.Z, DS3D_IMMEDIATE
                        I_oD3B.SetConeOrientation L_dD3Ori.X, L_dD3Ori.Y, L_dD3Ori.Z, DS3D_IMMEDIATE
                        I_oDSB.SetFrequency IIf(Model = rmJaguar, 12000, 10000) + Int(Velocity * 24000) + Int(Abs(Hoover) * 36000)
                    End If
                End If
    
    End Sub
    
    Public Sub Command(P_nCommand As ePlayerCommand)
    
        Dim L_dD3Pos As D3DVECTOR
        Dim L_dD3Ori As D3DVECTOR
        Dim L_dD3Nor As D3DVECTOR
        
        Dim L_nOldFuel As Single
        
        L_nOldFuel = Fuel
        
        Select Case P_nCommand
        
            Case pcAccellerate
                Velocity = Velocity + IIf(Model = rmWarthog, 0.0025, 0.005)
                I_nAccelleration = I_nAccelleration + IIf(Model = rmWarthog, 0.0025, 0.005)
                Fuel = Fuel - 2
                If Fuel < 10 Then Fuel = 10
                
            Case pcDecellerate
                Velocity = Velocity - 0.01
                I_nAccelleration = I_nAccelleration - 0.01
                Fuel = Fuel - 1
                If Fuel < 10 Then Fuel = 10
                
            Case pcBankLeft
                I_nBanking = I_nBanking - IIf(Model = rmWarthog, 0.001, 0.0015)
            
            Case pcBankRight
                I_nBanking = I_nBanking + IIf(Model = rmWarthog, 0.001, 0.0015)
            
            Case pcJump
                If Fuel > 10 Then
                    Hoover = Hoover + IIf(Model = rmWarthog, 0.005, 0.007)
                    Fuel = Fuel - IIf(Model = rmWarthog, 10, 15)
                    If Fuel < 10 Then Fuel = 10
                End If
            
            Case pcShoot
                    
                D3Frame.GetPosition Nothing, L_dD3Pos
                D3Frame.GetOrientation Nothing, L_dD3Ori, L_dD3Nor
                Application.Shots.Add stPlayer, L_dD3Pos, L_dD3Ori
                If (Not Application.DSInstance Is Nothing) And Application.SoundEnabled Then
                    I_oD3BShot.SetPosition L_dD3Pos.X, L_dD3Pos.Y, L_dD3Pos.Z, DS3D_IMMEDIATE
                    I_oD3BShot.SetConeOrientation L_dD3Ori.X, L_dD3Ori.Y, L_dD3Ori.Z, DS3D_IMMEDIATE
                    I_oDSBShot.Play DSBPLAY_DEFAULT
                End If
        End Select
    
    ' Set messages ...
    
        If L_nOldFuel > 10 And Fuel = 10 Then
            Application.Interface.DisplayMessage "out of fuel, return to station immediately ..."
        End If
        
    End Sub
    
    
    Public Function HitTest(ByRef P_dD3Pos As D3DVECTOR) As Boolean
    
        Dim L_dD3Pos As D3DVECTOR
        D3Frame.GetPosition Nothing, L_dD3Pos
        Application.DXInstance.VectorSubtract L_dD3Pos, L_dD3Pos, P_dD3Pos
        
        If Application.DXInstance.VectorModulus(L_dD3Pos) < 0.5 Then
            
            HitTest = True
        
        End If
        
    End Function
