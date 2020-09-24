VERSION 5.00
Begin VB.Form fER 
   Appearance      =   0  '2D
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'Kein
   Caption         =   "ERacer"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   ControlBox      =   0   'False
   FillStyle       =   0  'Ausgefüllt
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00008080&
   Icon            =   "fER.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'Bildschirmmitte
   WindowState     =   2  'Maximiert
End
Attribute VB_Name = "fER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' OPTION SETTING ...
                          
    ' Enforce variable declaration
    Option Explicit

' CODE ...

    '
    ' FORM_ACTIVATE: Initialize, run and terminate application
    '
    Private Sub Form_Activate()
    
        ' Initialize
        Set Application = New cApplication
        Application.Initialize
            
        ' Execute main loop
        Application.Execute
        
        ' End program
        Unload Me

        
    End Sub
    
    '
    ' FORM_UNLOAD: Cleanup and terminate application
    '
    Private Sub Form_Unload(Cancel As Integer)
    
            ' Cleanup DirectX, application, everything
            Application.Terminate
            
            ' End program
            End
    
    End Sub


