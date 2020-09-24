VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent           As cShots
Public D3Frame          As Direct3DRMFrame3

Private I_nCount        As Long
Private I_nType         As eShotType

Public Sub Initialize()

    Set D3Frame = Application.D3Instance.CreateFrame(Application.D3Frame)
    D3Frame.SetTraversalOptions 0
    
End Sub

Public Sub Update()

                                
    Dim L_dD3Ray    As D3DRMRAY
    Dim L_dD3PD     As D3DRMPICKDESC2
    Dim L_dD3Pos    As D3DVECTOR
    Dim L_dD3Nor    As D3DVECTOR
    Dim L_dD3Vel    As D3DVECTOR
    Dim L_oD3PDA    As Direct3DRMPick2Array
    Dim L_oD3Visual As Direct3DRMVisual
    Dim L_oEnemy    As cEnemy
    
    If D3Frame.GetTraversalOptions = 0 Then Exit Sub
    I_nCount = I_nCount - 1
    
    D3Frame.GetPosition Nothing, L_dD3Ray.pos
    D3Frame.GetOrientation Nothing, L_dD3Ray.Dir, L_dD3Nor
                             
    Set L_oD3PDA = Application.Environment.D3Frame.RayPick(Nothing, L_dD3Ray, D3DRMRAYPICK_IGNOREFURTHERPRIMITIVES)
            
    If Not (L_oD3PDA.GetSize = 0) Then
        If L_oD3PDA.GetPickFrame(0, L_dD3PD).GetSize > 0 Then
        
            Set L_oD3Visual = L_oD3PDA.GetPickVisual(0, L_dD3PD)
            
            Application.DXInstance.VectorSubtract L_dD3Pos, L_dD3PD.vPostion, L_dD3Ray.pos
            If Application.DXInstance.VectorModulus(L_dD3Pos) < 0.4 Then
                
                ' Station hit?
                If L_oD3Visual.GetAppData = 0 Then
                    
                    Application.Environment.StationHealth = Application.Environment.StationHealth - 90
                    
                    If Rnd > 0.75 And Application.Environment.StationHealth < 1500 And Application.Environment.StationHealth > 0 Then
                        If I_nType = stEnemy Then
                            Application.Interface.DisplayMessage " station under attack ... damage " & Format(((2000 - Application.Environment.StationHealth) / 2000) * 100, "0") & "% ..."
                        Else
                            Application.Interface.DisplayMessage " cease fire, watch your aim ... "
                        End If
                    End If
                End If

                Application.Effects.Add I_nType, L_dD3Ray.pos, L_dD3Vel
                
                I_nCount = 0
                
            End If
        End If
    End If
                                        
    If I_nType = stEnemy Then
    
        D3Frame.GetPosition Nothing, L_dD3Pos
        If Application.Player.HitTest(L_dD3Pos) Then
            Application.Effects.Add I_nType, L_dD3Ray.pos, L_dD3Vel
            Application.Player.Health = Application.Player.Health - 120
            I_nCount = 0
        End If
        
    End If
    
    If I_nType = stPlayer Then
    
        D3Frame.GetPosition Nothing, L_dD3Pos
        Set L_oEnemy = Application.Enemies.HitTest(L_dD3Pos)
        If Not (L_oEnemy Is Nothing) Then
            L_oEnemy.Health = L_oEnemy.Health - 90
            L_oEnemy.Hit = True
            Application.Effects.Add I_nType, L_dD3Ray.pos, L_dD3Vel
            I_nCount = 0
        End If
        
    End If
    
    If I_nCount = 0 Then Deactivate

End Sub

Public Sub Activate(ByVal P_nType As eShotType, ByRef P_dD3Pos As D3DVECTOR, ByRef P_dDir As D3DVECTOR)

    Select Case P_nType
        
        Case stPlayer
            D3Frame.AddVisual Parent.D3TexturePlayer
            D3Frame.SetVelocity Nothing, P_dDir.X * 0.4, P_dDir.Y * 0.4, P_dDir.Z * 0.4, D_FALSE
            I_nType = stPlayer
            I_nCount = 30
            
        Case stEnemy
            D3Frame.AddVisual Parent.D3TextureEnemy
            D3Frame.SetVelocity Nothing, P_dDir.X * 0.4, P_dDir.Y * 0.4, P_dDir.Z * 0.4, D_FALSE
            I_nType = stEnemy
            I_nCount = 30
            
    End Select
    
    D3Frame.SetTraversalOptions D3DRMFRAME_RENDERENABLE
    D3Frame.SetPosition Nothing, P_dD3Pos.X, P_dD3Pos.Y, P_dD3Pos.Z
    D3Frame.SetOrientation Nothing, P_dDir.X, P_dDir.Y, P_dDir.Z, 0, 1, 0
    
    Update
    
End Sub

Public Sub Deactivate()

    If D3Frame.GetVisualCount > 0 Then D3Frame.DeleteVisual D3Frame.GetVisual(0)
    D3Frame.SetTraversalOptions 0
    D3Frame.SetVelocity Nothing, 0, 0, 0, D_FALSE
    
End Sub
