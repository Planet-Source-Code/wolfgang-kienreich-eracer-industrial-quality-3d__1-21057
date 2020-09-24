VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnemies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Active                As Boolean
Public D3Frame               As Direct3DRMFrame3
Private I_oEnemy()           As cEnemy
Public Level                 As Single

Public ActiveCount           As Long

Public D3Mesh                As Direct3DRMMeshBuilder3

Public Sub ToggleSound()
    Dim L_nIndex As Integer
    For L_nIndex = 0 To 3
        If I_oEnemy(L_nIndex).D3Frame.GetTraversalOptions <> 0 Then
            If Application.SoundEnabled Then
                I_oEnemy(L_nIndex).DSB.Play DSBPLAY_LOOPING
            Else
                I_oEnemy(L_nIndex).DSB.Stop
            End If
        End If
    Next
End Sub

Public Sub Initialize()

    Dim L_nIndex    As Long
    
    Set D3Frame = Application.D3Instance.CreateFrame(Application.D3Frame)
    
    ReDim I_oEnemy(3) As cEnemy
    
    For L_nIndex = 0 To 3
    
        Set I_oEnemy(L_nIndex) = New cEnemy
        Set I_oEnemy(L_nIndex).Parent = Me
        I_oEnemy(L_nIndex).Initialize
    
    Next
    
    Set D3Mesh = Application.D3Instance.CreateMeshBuilder
    D3Mesh.LoadFromFile App.Path + "\mdl\a01.x", "", D3DRMLOAD_FROMFILE Or D3DRMLOAD_FIRST, Nothing, Nothing
    
    ActiveCount = 0
    Level = 0
    
End Sub

Public Function Add(ByRef P_dD3Pos As D3DVECTOR) As cEnemy

    Dim L_nIndex As Long
                        
    For L_nIndex = 0 To 3
        If I_oEnemy(L_nIndex).D3Frame.GetTraversalOptions = 0 Then Exit For
    Next
    
    If L_nIndex < 4 Then
        I_oEnemy(L_nIndex).Activate P_dD3Pos
        Set Add = I_oEnemy(L_nIndex)
    Else
        Set Add = Nothing
    End If
    
End Function

Public Sub Update()

    Dim L_nIndex    As Long
    Dim L_dD3Pos    As D3DVECTOR
    Dim L_nAngle    As Single
    
    For L_nIndex = 0 To 3
        If I_oEnemy(L_nIndex).D3Frame.GetTraversalOptions <> 0 Then I_oEnemy(L_nIndex).Update
    Next
    
    If Active Then
        If ActiveCount = 0 Then
            
            Level = Level + 1
            
            For L_nIndex = 0 To IIf(Level \ 4 > 3, 3, Level \ 3)
            
                L_nAngle = Rnd * PIVALUE * 2
                
                With L_dD3Pos
                    .X = Application.Environment.IslandSize \ 2 + Sin(L_nAngle) * (Application.Environment.IslandSize / 1.5)
                    .Y = -1
                    .Z = Application.Environment.IslandSize \ 2 + Cos(L_nAngle) * (Application.Environment.IslandSize / 1.5)
                End With
                
                Add L_dD3Pos
            
            Next
            
        End If
    End If
    
End Sub

Public Sub Reset()

    Dim L_nIndex As Long
    For L_nIndex = 0 To 3
        I_oEnemy(L_nIndex).Deactivate
    Next
    Level = 0
    ActiveCount = 0
    
End Sub

Public Sub Render()

    Dim L_nIndex    As Long
    Dim L_dD3Pos    As D3DVECTOR
    
    Application.DDSurface.SetForeColor RGB(255, 32, 0)
    
    For L_nIndex = 0 To 3
        If I_oEnemy(L_nIndex).D3Frame.GetTraversalOptions <> 0 Then
            I_oEnemy(L_nIndex).D3Frame.GetPosition Nothing, L_dD3Pos
            Application.DDSurface.DrawCircle 45 - Application.Environment.IslandSize \ 2 + Int(L_dD3Pos.X), 45 - Application.Environment.IslandSize \ 2 + Int(L_dD3Pos.Z), 1 + (Application.FrameCount Mod 24) \ 6
        End If
    Next

End Sub

Public Function HitTest(ByRef P_dD3Pos As D3DVECTOR) As cEnemy

    Dim L_nIndex As Long
    
    For L_nIndex = 0 To 3
        If I_oEnemy(L_nIndex).D3Frame.GetTraversalOptions <> 0 Then
            Set HitTest = I_oEnemy(L_nIndex).HitTest(P_dD3Pos)
            If Not (HitTest Is Nothing) Then Exit Function
        End If
    Next
    
End Function
