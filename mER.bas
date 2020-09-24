Attribute VB_Name = "mER"
'
' ERACER - Industrial quality 3D
'
' (C) 2001 NLS - Nonlinear Solutions
' Wolfgang Kienreich - Graz - Austria
'
' MEMBERS.AON.AT/NLS   NONLINEAR@AON.AT
'
'
'
' THIS PROGRAM IS FREEWARE. ANY USE IN COMMERCIAL
' APPLICATIONS WITHOUT WRITTEN PERMISSION BY THE
' AUTHOR IS PROHIBITED. ENJOY!
'



' OPTION SETTINGS ...
    
    Option Explicit
    
' CONSTANTS ...

    Public Const PIFACTOR = 1.74532925199433E-02
    Public Const PIVALUE = 3.14159265358979
    

' ENUMERATIONS ...
    
    ' Model of racer in use
    Public Enum eRacerModel
        rmWarthog = 1
        rmJaguar = 2
    End Enum
    
    ' Commands to player
    Public Enum ePlayerCommand
        pcAccellerate = 1
        pcDecellerate = 2
        pcBankLeft = 3
        pcBankRight = 4
        pcJump = 5
        pcShoot = 6
    End Enum
    
    ' Time of day in game
    Public Enum eDayTime
        dtDay = 1
        dtNight = 2
    End Enum
    
    ' Type of shot
    Public Enum eShotType
        stPlayer = 1
        stEnemy = 2
    End Enum
    
    ' Type of effect
    Public Enum eEffectType
        etShotPlayer = 1
        etShotEnemy = 2
        etExplo = 3
    End Enum
    
    ' Type of enemy
    Public Enum eEnemyType
        etSeeker = 1
        etHunter = 2
    End Enum
    
    ' Type of particle
    Public Enum eParticleType
        ptDustDay = 1
        ptDustNight = 2
        ptWaterDay = 3
        ptWaterNight = 4
    End Enum
    
    ' Type of sound
    Public Enum eSoundType
        stShotPlayer = 0
        stShotEnemy = 1
        stEnginePlayer = 2
        stEngineEnemy = 3
        stShotHits = 4
        stExploFighter = 5
        stExploStation = 6
    End Enum
    
' API FUNCTIONS ...

    ' Pointapi: Represents 2D-Point for api calls
    Public Type POINTAPI
        X As Long
        Y As Long
    End Type
    
    Public Type JOYINFOEX
        dwSize As Long
        dwFlags As Long
        dwXpos As Long
        dwYpos As Long
        dwZpos As Long
        dwRpos As Long
        dwUpos As Long
        dwVpos As Long
        dwButtons As Long
        dwButtonNumber As Long
        dwPOV As Long
        dwReserved1 As Long
        dwReserved2 As Long
    End Type
    
    ' TimeGetTime: Retrieves exact system time
    Public Declare Function timeGetTime Lib "winmm.dll" () As Long
    ' GetCursorPos: Retrieves cursorpos without using events
    Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    ' SetCursorPos: Sets cursor position
    Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
    ' Showcursor: Used to hide the cursor, in this case
    Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
    ' GetAsyncKeyState: Retrieves key presses without using events
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    ' CopyMemory: For fast memory transfer
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    ' JoyGetPosEx: Query joystick info
    Public Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
    
' VARIABLES...

    Public Application As cApplication      ' Application instance
    

' CODE ...

    '
    ' GENERATEBASEMESH: Generates a square mesh consisting of two triangles
    '
    Public Function GenerateBaseMesh(ByVal P_nScaleFactor As Single, Optional ByRef P_oD3Texture As Direct3DRMTexture3) As Direct3DRMMeshBuilder3
    
        ' Declare local variables ...
        
            Dim L_oD3Material   As Direct3DRMMaterial2      ' Holds material for faces
            Dim L_oD3Face       As Direct3DRMFace2          ' Holds faces to be added
        
        ' Code ...
        
            ' Create mesh
            Set GenerateBaseMesh = Application.D3Instance.CreateMeshBuilder
            
            With GenerateBaseMesh
            
                ' Add face #1
                With .CreateFace
                    .AddVertex -1, -1, 0
                    .AddVertex -1, 1, 0
                    .AddVertex 1, 1, 0
                End With
                
                ' Add face #2
                With .CreateFace
                    .AddVertex -1, -1, 0
                    .AddVertex 1, 1, 0
                    .AddVertex 1, -1, 0
                End With
            
                ' Set texture coordinates
                .SetTextureCoordinates 0, 0, 0
                .SetTextureCoordinates 1, 1, 0
                .SetTextureCoordinates 2, 1, 1
                .SetTextureCoordinates 3, 0, 0
                .SetTextureCoordinates 4, 1, 1
                .SetTextureCoordinates 5, 0, 1
                
                ' Generate normals
                .GenerateNormals 0, D3DRMGENERATENORMALS_PRECOMPACT
                    
                ' Set texture if provided
                If Not P_oD3Texture Is Nothing Then .SetTexture P_oD3Texture
                    
                ' Create and apply material
                Set L_oD3Material = Application.D3Instance.CreateMaterial(5)
                L_oD3Material.SetAmbient 1, 1, 1
                L_oD3Material.SetEmissive 1, 1, 1
                .SetMaterial L_oD3Material
                        
                ' Scale mesh
                .ScaleMesh P_nScaleFactor, P_nScaleFactor, P_nScaleFactor
                
            End With
        
    End Function
    
    Public Function DTA(ByVal P_nX As Single, ByVal P_nY As Single) As Single
    
        Dim L_bMirrored As Boolean
        
        L_bMirrored = P_nX < 0
            
        If L_bMirrored Then P_nX = -P_nX
        
        If P_nY = 0 Then
            If L_bMirrored Then
                DTA = PIVALUE * 1.5
            Else
                DTA = PIVALUE * 0.5
            End If
        ElseIf P_nY > 0 Then
            If L_bMirrored Then
               DTA = PIVALUE + Atn(P_nX / P_nY)
            Else
               DTA = PIVALUE - Atn(P_nX / P_nY)
            End If
        ElseIf P_nY < 0 Then
            If L_bMirrored Then
               DTA = PIVALUE * 2 + Atn(P_nX / P_nY)
            Else
               DTA = Abs(Atn(P_nX / P_nY))
            End If
        End If
            
    End Function
    
    Public Function RTS(ByVal P_nHeading1 As Single, ByVal P_nHeading2 As Single) As Single
            
        If (P_nHeading2 < P_nHeading1) Then
            If ((P_nHeading1 - P_nHeading2) > PIVALUE) Then
                RTS = -1
            Else
                RTS = 1
            End If
        End If
        If (P_nHeading2 > P_nHeading1) Then
            If ((P_nHeading2 - P_nHeading1) > PIVALUE) Then
                RTS = 1
            Else
                RTS = -1
            End If
        End If
                
    End Function

    Public Function RTD(ByVal P_nHeading1 As Single, ByVal P_nHeading2 As Single) As Single
            
        If P_nHeading2 > PIVALUE Then
            If P_nHeading1 < PIVALUE Then
                RTD = IIf(P_nHeading2 - P_nHeading1 > PIVALUE, P_nHeading1 + (PIVALUE * 2 - P_nHeading2), P_nHeading2 - P_nHeading1)
            Else
                RTD = IIf(P_nHeading2 > P_nHeading1, P_nHeading2 - P_nHeading1, P_nHeading1 - P_nHeading2)
            End If
        Else
            If P_nHeading1 > PIVALUE Then
                RTD = IIf(P_nHeading1 - P_nHeading2 > PIVALUE, P_nHeading2 + (PIVALUE * 2 - P_nHeading1), P_nHeading1 - P_nHeading2)
            Else
                RTD = IIf(P_nHeading2 > P_nHeading1, P_nHeading2 - P_nHeading1, P_nHeading1 - P_nHeading2)
            End If
        End If
                
    End Function

    Public Function GetWaveFileFormat(ByVal sFileName As String) As WAVEFORMATEX
        Dim L_dWFX As WAVEFORMATEX
        Dim L_nPosition As Long
        Dim L_nWaveBytes() As Byte
        ReDim L_nWaveBytes(1 To FileLen(sFileName))
        Open sFileName For Binary As #1
        Get #1, , L_nWaveBytes
        Close #1
        L_nPosition = 1
        Do While Not (Chr(L_nWaveBytes(L_nPosition)) + Chr(L_nWaveBytes(L_nPosition + 1)) + Chr(L_nWaveBytes(L_nPosition + 2)) = "fmt")
            L_nPosition = L_nPosition + 1
        Loop
        CopyMemory VarPtr(L_dWFX), VarPtr(L_nWaveBytes(L_nPosition + 8)), Len(L_dWFX)
    End Function
    
    


    

    



