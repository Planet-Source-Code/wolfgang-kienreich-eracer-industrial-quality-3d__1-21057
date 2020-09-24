VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cShots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private I_oShot()           As cShot
Public D3TexturePlayer      As Direct3DRMTexture3
Public D3TextureEnemy       As Direct3DRMTexture3

Public Sub Initialize()

    Dim L_nIndex As Long
    
    ReDim I_oShot(31) As cShot
    
    For L_nIndex = 0 To 31
    
        Set I_oShot(L_nIndex) = New cShot
        Set I_oShot(L_nIndex).Parent = Me
        I_oShot(L_nIndex).Initialize
    
    Next
    
    Set D3TexturePlayer = Application.D3Instance.LoadTexture(App.Path + "\gfx\shotp.bmp")
    With D3TexturePlayer
        .SetDecalOrigin 8, 8
        .SetDecalScale D_TRUE
        .SetDecalSize 0.1, 0.1
        .SetDecalTransparency D_TRUE
        .SetDecalTransparentColor Application.DXInstance.CreateColorRGBA(0, 0, 0, 0)
    End With
    
    Set D3TextureEnemy = Application.D3Instance.LoadTexture(App.Path + "\gfx\shote.bmp")
    With D3TextureEnemy
        .SetDecalOrigin 8, 8
        .SetDecalScale D_TRUE
        .SetDecalSize 0.1, 0.1
        .SetDecalTransparency D_TRUE
        .SetDecalTransparentColor Application.DXInstance.CreateColorRGBA(0, 0, 0, 0)
    End With
    
End Sub

Public Function Add(ByVal P_nType As eShotType, ByRef P_dD3Pos As D3DVECTOR, ByRef P_dDir As D3DVECTOR) As cShot

    Dim L_nIndex As Long
    
    For L_nIndex = 0 To 31
        If I_oShot(L_nIndex).D3Frame.GetTraversalOptions = 0 Then Exit For
    Next
    
    If L_nIndex < 32 Then
        I_oShot(L_nIndex).Activate P_nType, P_dD3Pos, P_dDir
        Set Add = I_oShot(L_nIndex)
    Else
        Set Add = Nothing
    End If
    
End Function

Public Sub Update()

    Dim L_nIndex As Long
    
    For L_nIndex = 0 To 31
        If I_oShot(L_nIndex).D3Frame.GetTraversalOptions <> 0 Then I_oShot(L_nIndex).Update
    Next
    
End Sub

Public Sub Reset()

    Dim L_nIndex As Long
    For L_nIndex = 0 To 31
        I_oShot(L_nIndex).Deactivate
    Next
    
End Sub

