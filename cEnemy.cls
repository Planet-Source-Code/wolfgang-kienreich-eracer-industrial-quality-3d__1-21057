VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum eEnemyState
    esAttackstation = 0
    esAttackPlayer = 1
    esInitEvade = 2
    esEvade = 3
    esInitRetreat = 4
    esRetreat = 5
End Enum

Public Parent               As cEnemies
Public D3Frame              As Direct3DRMFrame3

Public Health               As Single
Public Hit                  As Boolean
Public Deactivating         As Integer

Private I_nState            As eEnemyState
Private I_nVelocity         As Single
Private I_nBanking          As Single
Private I_nFallspeed        As Single
Private I_nHoover           As Single
Private I_nHeading          As Single

Public DSB                As DirectSoundBuffer
Private I_oD3B              As DirectSound3DBuffer

Private I_oDSBDie           As DirectSoundBuffer
Private I_oD3BDie           As DirectSound3DBuffer

Private I_oDSBShot          As DirectSoundBuffer
Private I_oD3BShot          As DirectSound3DBuffer

Public Sub Initialize()

    Set D3Frame = Application.D3Instance.CreateFrame(Application.Enemies.D3Frame)
    D3Frame.SetTraversalOptions 0
    
    If (Not Application.DSInstance Is Nothing) Then
        
        Dim L_ddsbd As DSBUFFERDESC
        With L_ddsbd
            .lBufferBytes = 0
            .lFlags = DSBCAPS_CTRL3D Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
        End With
        Set DSB = Application.DSInstance.CreateSoundBufferFromFile(App.Path + "\snd\enginee.wav", L_ddsbd, GetWaveFileFormat(App.Path + "\snd\enginee.wav"))
        Set I_oD3B = DSB.GetDirectSound3DBuffer
        With I_oD3B
            .SetConeAngles DS3D_MINCONEANGLE, DS3D_MAXCONEANGLE, DS3D_IMMEDIATE
            .SetConeOutsideVolume 0, DS3D_IMMEDIATE
            .SetMinDistance 1, DS3D_IMMEDIATE
            .SetMaxDistance 15, DS3D_IMMEDIATE
            .SetVelocity 0, 0, 0, DS3D_IMMEDIATE
            .SetMode DS3DMODE_NORMAL, DS3D_IMMEDIATE
        End With
        
        Set I_oDSBShot = Application.DSInstance.CreateSoundBufferFromFile(App.Path + "\snd\shote.wav", L_ddsbd, GetWaveFileFormat(App.Path + "\snd\shote.wav"))
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
    
End Sub

Public Sub Update()

    Dim L_dD3Pos            As D3DVECTOR
    Dim L_dD3Ori            As D3DVECTOR
    Dim L_dD3Nor            As D3DVECTOR
    Dim L_dD3Tmp            As D3DVECTOR

    Static S_dD3Destination As D3DVECTOR
    Static S_nShotDelay     As Integer
    Static S_nHooverDelay   As Integer
    Static S_bHasAttacked   As Boolean
    
    Dim L_dD3Ray            As D3DRMRAY             ' Ray for picking
    Dim L_oD3PDA            As Direct3DRMPick2Array ' Result of picking
    Dim L_dD3PD             As D3DRMPICKDESC2       ' Result of picking
    Dim L_oD3Visual         As Direct3DRMVisual     ' Visual picked
    
    Dim L_bColliding        As Boolean
    Dim L_nAltitude         As Single
    Dim L_nGroundType       As Long
    
    Dim L_nAngle            As Single
    Dim L_nDistance         As Single
    Dim L_nAngleDistance    As Single
    
    
    ' Dont process inactive enemies ...
        
        If D3Frame.GetTraversalOptions = 0 Then Exit Sub
    
    ' Gather enemy position and orientation ...
    
        D3Frame.GetPosition Nothing, L_dD3Pos
        D3Frame.GetOrientation Nothing, L_dD3Ori, L_dD3Nor
        
    ' Deactivating procedures ...
    
        If Deactivating > 0 Then
        
            ' Place explosions
            If Deactivating Mod 6 = 0 Then
                With L_dD3Ori
                    .X = 0
                    .Y = 0
                    .Z = 0
                End With
                With L_dD3Pos
                    .X = .X + Rnd * 0.8 - 0.4
                    .Y = .Y + Rnd * 0.8 - 0.4
                    .Z = .Z + Rnd * 0.8 - 0.4
                End With
                Application.Effects.Add etExplo, L_dD3Pos, L_dD3Ori
            End If
            
            ' Place rotation
            If Deactivating Mod 10 = 0 Then
                D3Frame.SetRotation Nothing, Rnd, Rnd, Rnd, Rnd * 0.25
            End If
            
            ' Count towards total deactivation
            Deactivating = Deactivating - 1
            If Deactivating = 0 Then
                L_dD3Ori.X = 0
                L_dD3Ori.Y = 0
                L_dD3Ori.Z = 0
                Application.Effects.Add etExplo, L_dD3Pos, L_dD3Ori
                Deactivate
            End If
            
            ' Dont process any further
            Exit Sub
            
        End If
        
    ' Absolute termination conditions ...
    
        ' Exit if out of map
        If L_dD3Pos.X > Application.Environment.IslandSize + 10 Or L_dD3Pos.Z > Application.Environment.IslandSize + 10 Or L_dD3Pos.X < -10 Or L_dD3Pos.Z < -10 Then
            Deactivate
            Exit Sub
        End If
    
        ' Die
        If Health < 10 Then
            Deactivating = 50
            If (Not Application.DSInstance Is Nothing) And Application.SoundEnabled Then
                I_oD3BDie.SetPosition L_dD3Pos.X, L_dD3Pos.Y, L_dD3Pos.Z, DS3D_IMMEDIATE
                I_oD3BDie.SetConeOrientation L_dD3Ori.X, L_dD3Ori.Y, L_dD3Ori.Z, DS3D_IMMEDIATE
                I_oDSBDie.Play DSBPLAY_DEFAULT
            End If
            Exit Sub
        End If
                    
    ' Gather enemy collission and floor ...
    
        ' Prepare ray
        Application.DXInstance.VectorCopy L_dD3Ray.pos, L_dD3Pos
        L_dD3Ray.pos.Y = 10
        L_dD3Ray.Dir.X = 0
        L_dD3Ray.Dir.Y = -1
        L_dD3Ray.Dir.Z = 0
    
        ' Cast ray
        Set L_oD3PDA = Application.Environment.D3Frame.RayPick(Nothing, L_dD3Ray, D3DRMRAYPICK_IGNOREFURTHERPRIMITIVES)
                         
        ' Retrieve results
        If Not (L_oD3PDA.GetSize = 0) Then
            If L_oD3PDA.GetPickFrame(0, L_dD3PD).GetSize > 0 Then
                Set L_oD3Visual = L_oD3PDA.GetPickVisual(0, L_dD3PD)
                L_nGroundType = L_oD3Visual.GetAppData
                L_nAltitude = L_dD3Pos.Y - L_dD3PD.vPostion.Y
            End If
        End If

        ' Prepare ray
        Application.DXInstance.VectorCopy L_dD3Ray.pos, L_dD3Pos
        L_dD3Ray.Dir.X = L_dD3Ori.X
        L_dD3Ray.Dir.Y = 0.2
        L_dD3Ray.Dir.Z = L_dD3Ori.Z
        Application.DXInstance.VectorNormalize L_dD3Ray.Dir
        
        ' Cast ray
        Set L_oD3PDA = Application.Environment.D3FrameStation.RayPick(Nothing, L_dD3Ray, D3DRMRAYPICK_IGNOREFURTHERPRIMITIVES)
                         
        ' Retrieve results
        If Not (L_oD3PDA.GetSize = 0) Then
            If L_oD3PDA.GetPickFrame(0, L_dD3PD).GetSize > 0 Then
                Set L_oD3Visual = L_oD3PDA.GetPickVisual(0, L_dD3PD)
                Application.DXInstance.VectorSubtract L_dD3Tmp, L_dD3Ray.pos, L_dD3PD.vPostion
                L_bColliding = (Application.DXInstance.VectorModulus(L_dD3Tmp) < 3)
            End If
        End If

    ' Decide if shoot at station...
    
        If S_nShotDelay = 0 Then
        
            ' At player
            Application.Player.D3Frame.GetPosition Nothing, L_dD3Tmp
            Application.DXInstance.VectorSubtract L_dD3Tmp, L_dD3Pos, L_dD3Tmp
            L_nDistance = Application.DXInstance.VectorModulus(L_dD3Tmp)
            Application.DXInstance.VectorNormalize L_dD3Tmp
            L_nAngle = DTA(-L_dD3Tmp.Z, -L_dD3Tmp.X)
            L_nAngleDistance = RTD(I_nHeading, L_nAngle)
        
            If L_nDistance < 9 Then
                If RTD(I_nHeading, L_nAngle) < PIVALUE * 0.2 Then
                    S_bHasAttacked = True
                    Application.Shots.Add stEnemy, L_dD3Pos, L_dD3Ori
                    If (Not Application.DSInstance Is Nothing) And Application.SoundEnabled Then
                        I_oD3BShot.SetPosition L_dD3Pos.X, L_dD3Pos.Y, L_dD3Pos.Z, DS3D_IMMEDIATE
                        I_oD3BShot.SetConeOrientation L_dD3Ori.X, L_dD3Ori.Y, L_dD3Ori.Z, DS3D_IMMEDIATE
                        I_oDSBShot.Play DSBPLAY_DEFAULT
                    End If
                    S_nShotDelay = IIf(Parent.Level < 4, 10, 8)
                End If
            End If

            ' At station
            If I_nState = esAttackstation Then
            
                Application.Environment.D3FrameStation.GetPosition Nothing, L_dD3Tmp
                Application.DXInstance.VectorSubtract L_dD3Tmp, L_dD3Pos, L_dD3Tmp
                L_nDistance = Application.DXInstance.VectorModulus(L_dD3Tmp)
                Application.DXInstance.VectorNormalize L_dD3Tmp
                L_nAngle = DTA(-L_dD3Tmp.Z, -L_dD3Tmp.X)
                L_nAngleDistance = RTD(I_nHeading, L_nAngle)
    
                If L_nDistance < 12 Then
                    If RTD(I_nHeading, L_nAngle) < PIVALUE * 0.2 Then
                        Application.Shots.Add stEnemy, L_dD3Pos, L_dD3Ori
                        If (Not Application.DSInstance Is Nothing) And Application.SoundEnabled Then
                            I_oD3BShot.SetPosition L_dD3Pos.X, L_dD3Pos.Y, L_dD3Pos.Z, DS3D_IMMEDIATE
                            I_oD3BShot.SetConeOrientation L_dD3Ori.X, L_dD3Ori.Y, L_dD3Ori.Z, DS3D_IMMEDIATE
                            I_oDSBShot.Play DSBPLAY_DEFAULT
                        End If
                        S_nShotDelay = IIf(Parent.Level < 4, 10, 8)
                    End If
                End If
                
            End If
            
        Else
            S_nShotDelay = S_nShotDelay - 1
        End If
        
    ' Gather destination properties...
        
        Application.DXInstance.VectorSubtract L_dD3Tmp, L_dD3Pos, S_dD3Destination
        L_nDistance = Application.DXInstance.VectorModulus(L_dD3Tmp)
        Application.DXInstance.VectorNormalize L_dD3Tmp
        L_nAngle = DTA(-L_dD3Tmp.Z, -L_dD3Tmp.X)
        L_nAngleDistance = RTD(I_nHeading, L_nAngle)
        
    ' Process state engine ...
    
        ' Default state processing
        Select Case I_nState
        
            Case esAttackstation
                                                                
                Application.Environment.D3FrameStation.GetPosition Nothing, S_dD3Destination
                
                If L_nDistance < 8 Then
                    I_nState = esInitEvade
                End If
            
            Case esAttackPlayer
                                        
                Application.Player.D3Frame.GetPosition Nothing, S_dD3Destination
                
                If L_nDistance < 3 And S_bHasAttacked Then
                    I_nState = esInitEvade
                End If
                
                If L_nDistance < 12 And L_nDistance > 6 And Rnd > 0.95 And S_nHooverDelay = 0 Then
                    S_nHooverDelay = 10
                End If
                
                If L_nDistance < 5 Then
                    If Rnd > 0.98 Then
                        I_nState = esInitEvade
                    End If
                End If
                
            Case esInitEvade
                            
                L_nAngle = PIVALUE * 2 * Rnd
                L_nDistance = Application.Environment.IslandSize \ 3
                
                S_dD3Destination.X = L_nDistance * Cos(L_nAngle) + L_dD3Pos.X
                S_dD3Destination.Z = L_nDistance * Sin(L_nAngle) + L_dD3Pos.Z
                
                I_nState = esEvade
                
                S_bHasAttacked = False
                
            Case esEvade
            
                If L_nDistance < 3 Then
 
                    I_nState = IIf(Rnd > 0.5, esAttackstation, esAttackPlayer)
                    If I_nState = esAttackstation Then
                        Application.Environment.D3FrameStation.GetPosition Nothing, S_dD3Destination
                    End If
                    If I_nState = esAttackPlayer Then
                        Application.Player.D3Frame.GetPosition Nothing, S_dD3Destination
                    End If
                End If
                
            Case esInitRetreat
            
                L_nAngle = PIVALUE * 2 * Rnd
                L_nDistance = Application.Environment.IslandSize * 8
                
                Application.Environment.D3FrameStation.GetPosition Nothing, L_dD3Tmp
                S_dD3Destination.X = L_nDistance * Cos(L_nAngle) + L_dD3Tmp.X
                S_dD3Destination.Z = L_nDistance * Sin(L_nAngle) + L_dD3Tmp.Z
                
                I_nState = esRetreat
            
            Case esRetreat
            
        End Select
        
        ' Reaction to being shot at
        If Hit Then
            I_nState = IIf(Rnd > 0.5, esAttackPlayer, esInitEvade)
            If Rnd > 0.75 Then S_nHooverDelay = 15
            If Rnd > 0.75 Then I_nVelocity = I_nVelocity * 0.1
        End If
    
        ' Reaction to too much damage
        If Health < 50 Then I_nState = esRetreat
       
    ' Control racer (set heading, velocity, hoover) ...
        
        ' Bank to new heading
        I_nBanking = I_nBanking + RTS(L_nAngle, I_nHeading) * 0.002
        
        ' Adjust velocity, incorporate banking
        I_nVelocity = I_nVelocity + (0.004 - I_nBanking * 0.1)
        
        ' Hoover if commanded
        If S_nHooverDelay > 0 Then
            S_nHooverDelay = S_nHooverDelay - 1
            I_nHoover = I_nHoover + 0.003
        End If
        
    ' Update racer data ...
    
        ' Minimum L_nAltitude / collission
        If L_nAltitude < 0.25 Then
            L_nAltitude = 0.25
            I_nHoover = I_nHoover + 0.005
        End If
        If L_bColliding Then
            I_nState = esInitEvade
            I_nHoover = I_nHoover + 0.005
            I_nVelocity = I_nVelocity - 0.006
        End If
                                                                                               
        ' AGrav dekey
        I_nHoover = I_nHoover * 0.95
        
         ' AGrav dekey above Altitude
        If L_nAltitude > 2 Then I_nHoover = I_nHoover * 0.9
           
        ' Gravity
         I_nFallspeed = I_nFallspeed - 0.01
         
         ' Hover
        I_nFallspeed = I_nFallspeed + (1 / L_nAltitude) * 0.006
        
         ' AGrav
        I_nFallspeed = I_nFallspeed + I_nHoover * 0.1
        
        ' Air friction levels falling speed
        I_nFallspeed = I_nFallspeed * 0.95
        
       ' Banking
        I_nHeading = I_nHeading + I_nBanking
        If I_nHeading > PIVALUE * 2 Then I_nHeading = I_nHeading - PIVALUE * 2
        If I_nHeading < 0 Then I_nHeading = I_nHeading + PIVALUE * 2
        
        ' Banking dekeys
        I_nBanking = I_nBanking * 0.95
         
        ' Limit Velocity
        I_nVelocity = I_nVelocity * IIf(I_nVelocity > 0, 0.99, 0.7)
        If I_nVelocity > 0.095 Then I_nVelocity = 0.095
        
    ' Set enemy data ...
                    
        ' Calculate Position ...
        
            L_dD3Pos.X = L_dD3Pos.X - Cos(I_nHeading) * I_nVelocity
            L_dD3Pos.Y = L_dD3Pos.Y + I_nFallspeed
            If L_dD3Pos.Y < 0 Then L_dD3Pos.Y = 0
            L_dD3Pos.Z = L_dD3Pos.Z + Sin(I_nHeading) * I_nVelocity
                 
        ' Calculate Orientation ...
                        
            ' Orientation
            L_dD3Ori.Y = I_nFallspeed * 5
            L_dD3Ori.X = -Cos(I_nHeading)
            L_dD3Ori.Z = Sin(I_nHeading)
            
            ' Normal
            L_dD3Nor.X = I_nBanking * 50 * Sin(I_nHeading)
            L_dD3Nor.Y = 1 - I_nFallspeed * 5
            L_dD3Nor.Z = I_nBanking * 50 * Cos(I_nHeading)
            
            ' Normalize data
            Application.DXInstance.VectorNormalize L_dD3Ori
            Application.DXInstance.VectorNormalize L_dD3Nor
        
        ' Set values ...
        
            D3Frame.SetPosition Nothing, L_dD3Pos.X, L_dD3Pos.Y, L_dD3Pos.Z
            D3Frame.SetOrientation Nothing, L_dD3Ori.X, L_dD3Ori.Y, L_dD3Ori.Z, L_dD3Nor.X, L_dD3Nor.Y, L_dD3Nor.Z
        
        ' Reset hit state ...
        
            Hit = False
            
        ' Set sound ...
                    
            If (Not Application.DSInstance Is Nothing) And Application.SoundEnabled Then
                With I_oD3B
                    .SetPosition L_dD3Pos.X, L_dD3Pos.Y, L_dD3Pos.Z, DS3D_IMMEDIATE
                    .SetConeOrientation L_dD3Ori.X, L_dD3Ori.Y, L_dD3Ori.Z, DS3D_IMMEDIATE
                    DSB.SetFrequency 12000 - Int(Abs(I_nBanking) * 24000)
                End With
            End If
    
End Sub

Public Sub Activate(ByRef P_dD3Pos As D3DVECTOR)

    D3Frame.AddVisual Parent.D3Mesh
    
    I_nState = IIf(Rnd > 0.5, esAttackPlayer, esAttackstation)
    Hit = False
    Health = 400 + Application.Enemies.Level * 20
    Parent.ActiveCount = Parent.ActiveCount + 1
    D3Frame.SetTraversalOptions D3DRMFRAME_RENDERANDPICKENABLE
    D3Frame.SetPosition Nothing, P_dD3Pos.X, P_dD3Pos.Y, P_dD3Pos.Z
    D3Frame.SetOrientation Nothing, 0, 0, 1, 0, 1, 0
    
    If (Not Application.DSInstance Is Nothing) And Application.SoundEnabled Then
        DSB.Play DSBPLAY_LOOPING
    End If
    
    Update
    
End Sub

Public Sub Deactivate()

    If (Not Application.DSInstance Is Nothing) And Application.SoundEnabled Then
        DSB.Stop
    End If
    
    If D3Frame.GetVisualCount > 0 Then D3Frame.DeleteVisual D3Frame.GetVisual(0)
    D3Frame.SetTraversalOptions 0
    Parent.ActiveCount = Parent.ActiveCount - 1
    
End Sub

Public Function HitTest(ByRef P_dD3Pos As D3DVECTOR) As cEnemy

    Dim L_dD3Pos As D3DVECTOR
    D3Frame.GetPosition Nothing, L_dD3Pos
    Application.DXInstance.VectorSubtract L_dD3Pos, L_dD3Pos, P_dD3Pos
    
    If Application.DXInstance.VectorModulus(L_dD3Pos) < 0.4 Then
        
        Set HitTest = Me
    
    End If
    
End Function
