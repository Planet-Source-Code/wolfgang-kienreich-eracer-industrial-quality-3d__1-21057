VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEffect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent           As cEffects
Public D3Frame          As Direct3DRMFrame3

Private I_oDDSTexture   As DirectDrawSurface4
Private I_oDDSSource    As DirectDrawSurface4
Private I_oD3Texture    As Direct3DRMTexture3

Private I_nType         As eEffectType
Private I_nScale        As Single
Private I_nAlpha        As Single
Private I_nCount        As Long

Private I_oDSB              As DirectSoundBuffer
Private I_oD3B              As DirectSound3DBuffer

Public Sub Initialize()

    Dim L_dDDSD         As DDSURFACEDESC2
    Dim L_oD3Mesh       As Direct3DRMMeshBuilder3
    Dim L_oD3Material   As Direct3DRMMaterial2
    
    Set D3Frame = Application.D3Instance.CreateFrame(Application.D3Frame)
    D3Frame.SetTraversalOptions 0
    
    With L_dDDSD
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .lHeight = 32
        .lWidth = 32
        .ddsCaps.lCaps = DDSCAPS_TEXTURE
    End With
                        
    Set I_oDDSTexture = Application.DDInstance.CreateSurface(L_dDDSD)
    
    Set I_oD3Texture = Application.D3Instance.CreateTextureFromSurface(I_oDDSTexture)
    With I_oD3Texture
        .SetDecalTransparency D_TRUE
        .SetDecalTransparentColor 0
    End With
    
    D3Frame.AddVisual GenerateBaseMesh(1, I_oD3Texture)
    
    If (Not Application.DSInstance Is Nothing) Then
        
        Dim L_ddsbd As DSBUFFERDESC
        With L_ddsbd
            .lBufferBytes = 0
            .lFlags = DSBCAPS_CTRL3D Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
        End With
        
        Set I_oDSB = Application.DSInstance.CreateSoundBufferFromFile(App.Path + "\snd\explosmall.wav", L_ddsbd, GetWaveFileFormat(App.Path + "\snd\explosmall.wav"))
        Set I_oD3B = I_oDSB.GetDirectSound3DBuffer
        With I_oD3B
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
            
    Dim L_dArea     As RECT
    Dim L_dAreas(0) As RECT
    Dim L_dD3Ori    As D3DVECTOR
    Dim L_dD3Nor    As D3DVECTOR
    Dim L_oD3MO     As D3DRMMATERIALOVERRIDE
    
    If D3Frame.GetTraversalOptions = 0 Then Exit Sub
    
    I_nCount = I_nCount - 1
    
    With L_dArea
        .Left = 32 * (((36 - I_nCount) \ 3) Mod 4)
        .Top = 32 * (((36 - I_nCount) \ 3) \ 4)
        .Right = .Left + 32
        .Bottom = .Top + 32
    End With
    
    I_oDDSTexture.BltFast 0, 0, I_oDDSSource, L_dArea, DDBLTFAST_NOCOLORKEY Or DDBLTFAST_WAIT
    I_oD3Texture.Changed D3DRMTEXTURE_CHANGEDPIXELS, 0, L_dAreas
    
    Application.D3Camera.GetOrientation Nothing, L_dD3Ori, L_dD3Nor
    D3Frame.SetOrientation Nothing, L_dD3Ori.X, L_dD3Ori.Y, L_dD3Ori.Z, L_dD3Nor.X, L_dD3Nor.Y, L_dD3Nor.Z
                                        
    With L_oD3MO
        .lFlags = D3DRMMATERIALOVERRIDE_DIFFUSE_ALPHAONLY Or D3DRMMATERIALOVERRIDE_DIFFUSE_ALPHAMULTIPLY
        .dcDiffuse.a = I_nAlpha * (I_nCount / 36)
    End With
    D3Frame.SetMaterialOverride L_oD3MO
    
    If I_nCount = 0 Then Deactivate

End Sub

Public Sub Activate(ByVal P_nType As eEffectType, ByRef P_dD3Pos As D3DVECTOR, ByRef P_dD3Vel As D3DVECTOR)

    Dim L_oD3Mesh As Direct3DRMMeshBuilder3
    
    Select Case P_nType
        
        Case etShotPlayer
            
            Set I_oDDSSource = Parent.DDSurfaceShotPlayer
            D3Frame.SetVelocity Nothing, 0, 0, 0, D_FALSE
            I_nType = etShotPlayer
            I_nCount = 36
            I_nScale = 0.25
            I_nAlpha = 1
            If (Not Application.DSInstance Is Nothing) And Application.SoundEnabled Then
                I_oD3B.SetPosition P_dD3Pos.X, P_dD3Pos.Y, P_dD3Pos.Z, DS3D_IMMEDIATE
                I_oD3B.SetConeOrientation 0, 1, 0, DS3D_IMMEDIATE
                I_oDSB.Play DSBPLAY_DEFAULT
            End If
            
        Case etShotEnemy
        
            Set I_oDDSSource = Parent.DDSurfaceShotEnemy
            D3Frame.SetVelocity Nothing, 0, 0, 0, D_FALSE
            I_nType = etShotEnemy
            I_nCount = 36
            I_nScale = 0.2
            I_nAlpha = 1
            If (Not Application.DSInstance Is Nothing) And Application.SoundEnabled Then
                I_oD3B.SetPosition P_dD3Pos.X, P_dD3Pos.Y, P_dD3Pos.Z, DS3D_IMMEDIATE
                I_oD3B.SetConeOrientation 0, 1, 0, DS3D_IMMEDIATE
                I_oDSB.Play DSBPLAY_DEFAULT
            End If
            
        Case etExplo
        
            Set I_oDDSSource = Parent.DDSurfaceExplo
            D3Frame.SetVelocity Nothing, 0, 0, 0, D_FALSE
            I_nType = etExplo
            I_nCount = 30
            I_nScale = 0.3
            I_nAlpha = 1
            
    End Select
    
    D3Frame.SetTraversalOptions D3DRMFRAME_RENDERENABLE
    D3Frame.SetPosition Nothing, P_dD3Pos.X, P_dD3Pos.Y, P_dD3Pos.Z
    D3Frame.SetVelocity Nothing, P_dD3Vel.X, P_dD3Vel.Y, P_dD3Vel.Z, D_FALSE
    
    Set L_oD3Mesh = D3Frame.GetVisual(0)
    L_oD3Mesh.ScaleMesh I_nScale, I_nScale, I_nScale
    
    Update
    
End Sub

Public Sub Deactivate()

    Dim L_oD3Mesh As Direct3DRMMeshBuilder3
    
    If D3Frame.GetTraversalOptions = 0 Then Exit Sub
    
    D3Frame.SetTraversalOptions 0
    D3Frame.SetVelocity Nothing, 0, 0, 0, D_FALSE
    
    Set L_oD3Mesh = D3Frame.GetVisual(0)
    L_oD3Mesh.ScaleMesh 1 / I_nScale, 1 / I_nScale, 1 / I_nScale

    
End Sub

