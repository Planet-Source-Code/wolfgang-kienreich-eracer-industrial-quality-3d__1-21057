VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cParticles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public D3MeshP1              As Direct3DRMMeshBuilder3
Public D3MeshP2              As Direct3DRMMeshBuilder3

Private I_oParticle()           As cParticle

Public Sub Initialize()

    Dim L_nIndex As Long
    Dim L_oD3Texture As Direct3DRMTexture3
    
    Set L_oD3Texture = Application.D3Instance.LoadTexture(App.Path + "\gfx\p01.bmp")
    With L_oD3Texture
        .SetDecalTransparency D_TRUE
        .SetDecalTransparentColor 0
    End With
    Set D3MeshP1 = GenerateBaseMesh(1, L_oD3Texture)
    
    Set L_oD3Texture = Application.D3Instance.LoadTexture(App.Path + "\gfx\p02.bmp")
    With L_oD3Texture
        .SetDecalTransparency D_TRUE
        .SetDecalTransparentColor 0
    End With
    Set D3MeshP2 = GenerateBaseMesh(1, L_oD3Texture)
    
    ReDim I_oParticle(31) As cParticle
    
    For L_nIndex = 0 To 31
    
        Set I_oParticle(L_nIndex) = New cParticle
        Set I_oParticle(L_nIndex).Parent = Me
        I_oParticle(L_nIndex).Initialize
    
    Next
    
End Sub

Public Function Add(ByVal P_nType As eParticleType, ByRef P_dD3Pos As D3DVECTOR, ByRef P_dD3Vel As D3DVECTOR) As cParticle

    Dim L_nIndex As Long
    
    For L_nIndex = 0 To 31
        If I_oParticle(L_nIndex).D3Frame.GetTraversalOptions = 0 Then Exit For
    Next
    
    If L_nIndex < 32 Then
        I_oParticle(L_nIndex).Activate P_nType, P_dD3Pos, P_dD3Vel
        Set Add = I_oParticle(L_nIndex)
    Else
        Set Add = Nothing
    End If
    
End Function

Public Sub Update()

    Dim L_nIndex As Long
    
    For L_nIndex = 0 To 31
        If I_oParticle(L_nIndex).D3Frame.GetTraversalOptions <> 0 Then I_oParticle(L_nIndex).Update
    Next
    
End Sub

Public Sub Reset()

    Dim L_nIndex As Long
    For L_nIndex = 0 To 31
        I_oParticle(L_nIndex).Deactivate
    Next
    
End Sub

