VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEffects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private I_oEffect()             As cEffect
Public DDSurfaceShotPlayer      As DirectDrawSurface4
Public DDSurfaceShotEnemy       As DirectDrawSurface4
Public DDSurfaceExplo           As DirectDrawSurface4

Public Sub Initialize()

    Dim L_nIndex As Long
    Dim L_dDDSD  As DDSURFACEDESC2
    
    ReDim I_oEffect(31) As cEffect
    
    For L_nIndex = 0 To 31
    
        Set I_oEffect(L_nIndex) = New cEffect
        Set I_oEffect(L_nIndex).Parent = Me
        I_oEffect(L_nIndex).Initialize
    
    Next
    
    With L_dDDSD
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lWidth = 128
        .lHeight = 96
    End With
    
    Set DDSurfaceShotPlayer = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\gfx\explo1.bmp", L_dDDSD)
    Set DDSurfaceShotEnemy = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\gfx\explo2.bmp", L_dDDSD)
    Set DDSurfaceExplo = Application.DDInstance.CreateSurfaceFromFile(App.Path + "\gfx\explo3.bmp", L_dDDSD)
    
End Sub

Public Function Add(ByVal P_nType As eEffectType, ByRef P_dD3Pos As D3DVECTOR, ByRef P_dD3Vel As D3DVECTOR) As cEffect

    Dim L_nIndex As Long
    
    For L_nIndex = 0 To 31
        If I_oEffect(L_nIndex).D3Frame.GetTraversalOptions = 0 Then Exit For
    Next
    
    If L_nIndex < 32 Then
        I_oEffect(L_nIndex).Activate P_nType, P_dD3Pos, P_dD3Vel
        Set Add = I_oEffect(L_nIndex)
    Else
        Set Add = Nothing
    End If
    
End Function

Public Sub Update()

    Dim L_nIndex As Long
    
    For L_nIndex = 0 To 31
        If I_oEffect(L_nIndex).D3Frame.GetTraversalOptions <> 0 Then I_oEffect(L_nIndex).Update
    Next
    
End Sub

Public Sub Reset()

    Dim L_nIndex As Long
    For L_nIndex = 0 To 31
        I_oEffect(L_nIndex).Deactivate
    Next
    
End Sub
