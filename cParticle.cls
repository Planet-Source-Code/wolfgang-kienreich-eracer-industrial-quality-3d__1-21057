VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cParticle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent           As cParticles
Public D3Frame          As Direct3DRMFrame3

Private I_nType         As eParticleType
Private I_nAlphaStart   As Single
Private I_nAlphaFinal   As Single
Private I_nCount        As Long
Private I_nScale        As Single

Public Sub Initialize()

    
    Set D3Frame = Application.D3Instance.CreateFrame(Application.D3Frame)
    D3Frame.SetTraversalOptions 0
    
End Sub

Public Sub Update()
            
    Dim L_dArea     As RECT
    Dim L_dD3Ori    As D3DVECTOR
    Dim L_dD3Nor    As D3DVECTOR
    Dim L_oD3MO     As D3DRMMATERIALOVERRIDE
    
    If D3Frame.GetTraversalOptions = 0 Then Exit Sub
    
    I_nCount = I_nCount - 1
    
    Application.D3Camera.GetOrientation Nothing, L_dD3Ori, L_dD3Nor
    D3Frame.SetOrientation Nothing, L_dD3Ori.X, L_dD3Ori.Y, L_dD3Ori.Z, L_dD3Nor.X, L_dD3Nor.Y, L_dD3Nor.Z
                                        
    With L_oD3MO
        .lFlags = D3DRMMATERIALOVERRIDE_DIFFUSE_ALPHAONLY Or D3DRMMATERIALOVERRIDE_DIFFUSE_ALPHAMULTIPLY
        .dcDiffuse.a = I_nAlphaFinal + (I_nAlphaStart - I_nAlphaFinal) * (I_nCount / 20)
    End With
    D3Frame.SetMaterialOverride L_oD3MO
    If I_nCount = 0 Then Deactivate

End Sub

Public Sub Activate(ByVal P_nType As eParticleType, ByRef P_dD3Pos As D3DVECTOR, ByRef P_dD3Vel As D3DVECTOR)

    Dim L_oD3Mesh As Direct3DRMMeshBuilder3
    
    Select Case P_nType
        
        Case ptDustDay
            
            D3Frame.AddVisual Parent.D3MeshP1.CloneObject
            D3Frame.SetVelocity Nothing, 0, 0, 0, D_FALSE
            I_nType = ptDustDay
            I_nCount = 20
            I_nScale = 0.35
            I_nAlphaStart = 0.15
            I_nAlphaFinal = 0
            
        Case ptDustNight
            
            D3Frame.AddVisual Parent.D3MeshP1.CloneObject
            D3Frame.SetVelocity Nothing, 0, 0, 0, D_FALSE
            I_nType = ptDustNight
            I_nCount = 15
            I_nScale = 0.35
            I_nAlphaStart = 0.1
            I_nAlphaFinal = 0
            
        Case ptWaterDay
            
            D3Frame.AddVisual Parent.D3MeshP2.CloneObject
            D3Frame.SetVelocity Nothing, 0, 0, 0, D_FALSE
            I_nType = ptWaterDay
            I_nCount = 20
            I_nScale = 0.2
            I_nAlphaStart = 0.25
            I_nAlphaFinal = 0
            
        Case ptWaterNight
            
            D3Frame.AddVisual Parent.D3MeshP2.CloneObject
            D3Frame.SetVelocity Nothing, 0, 0, 0, D_FALSE
            I_nType = ptWaterNight
            I_nCount = 15
            I_nScale = 0.2
            I_nAlphaStart = 0.15
            I_nAlphaFinal = 0
            
    End Select
    
    D3Frame.SetTraversalOptions D3DRMFRAME_RENDERENABLE
    D3Frame.SetPosition Nothing, P_dD3Pos.X, P_dD3Pos.Y, P_dD3Pos.Z
    D3Frame.SetVelocity Nothing, P_dD3Vel.X, P_dD3Vel.Y, P_dD3Vel.Z, D_FALSE
    
    Set L_oD3Mesh = D3Frame.GetVisual(0)
    L_oD3Mesh.ScaleMesh I_nScale, I_nScale, I_nScale
    
    Update
    
End Sub

Public Sub Deactivate()

    If D3Frame.GetTraversalOptions = 0 Then Exit Sub
    
    D3Frame.SetTraversalOptions 0
    D3Frame.SetVelocity Nothing, 0, 0, 0, D_FALSE
    D3Frame.DeleteVisual D3Frame.GetVisual(0)
    
End Sub


