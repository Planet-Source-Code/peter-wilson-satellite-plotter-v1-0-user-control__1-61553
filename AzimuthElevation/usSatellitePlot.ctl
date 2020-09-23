VERSION 5.00
Begin VB.UserControl usSatellitePlot 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "usSatellitePlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Reset()

    Dim sngRatio    As Single
    
    sngRatio = UserControl.Width / UserControl.Height

    UserControl.ScaleLeft = -1.2 * sngRatio
    UserControl.ScaleWidth = 2.4 * sngRatio
    UserControl.ScaleTop = -1.2
    UserControl.ScaleHeight = 2.4
    
    UserControl.Cls
    UserControl.DrawWidth = 1
    
    ' Draw Circles.
    UserControl.FontBold = False
    Call DrawCircle(5, 1, RGB(0, 0, 255))
    Call DrawCircle(10, 0.6666, RGB(192, 192, 255))
    Call DrawCircle(10, 0.5, RGB(192, 192, 255))
    Call DrawCircle(10, 0.3333, RGB(192, 192, 255))
    Call DrawCircle(10, 0, RGB(192, 192, 255))
    
    ' Draw Cross-Hairs.
    UserControl.FontBold = True
    UserControl.ForeColor = RGB(255, 192, 192)
    Call DrawCrossHairs(45, 1.1)
    
End Sub

Private Sub DrawCircle(StepSize As Single, Radius As Single, Color As OLE_COLOR)

    Dim sngDegrees  As Single
    Dim sngRadians  As Single
    Dim sngX        As Single
    Dim sngY        As Single
    
    For sngDegrees = 45 To 360 + 45 Step StepSize
    
        sngRadians = sngDegrees * 1.74532925199433E-02
        
        sngX = Radius * Sin(sngRadians)
        sngY = Radius * -Cos(sngRadians)
        
        If sngDegrees = 45 Then
            UserControl.CurrentX = sngX
            UserControl.CurrentY = sngY
        Else
            UserControl.Line -(sngX, sngY), Color
        End If
        
    Next sngDegrees
    
    UserControl.ForeColor = vbBlack
    UserControl.Print Round(90 - (90 * Radius))
    
End Sub

Private Sub DrawCrossHairs(StepSize As Single, Radius As Single)

    Dim sngDegrees  As Single
    Dim sngRadians  As Single
    Dim sngX        As Single
    Dim sngY        As Single
    
    For sngDegrees = 0 To (360 - StepSize) Step StepSize
    
        sngRadians = sngDegrees * 1.74532925199433E-02
        
        sngX = Radius * Sin(sngRadians)
        sngY = Radius * -Cos(sngRadians)
        
        UserControl.Line (0, 0)-(sngX, sngY), RGB(255, 192, 192)
        UserControl.ForeColor = vbBlack
        UserControl.Print sngDegrees
        
    Next sngDegrees

End Sub

Public Sub PlotSatellite(Azimuth As Single, Elevation As Single)
    
    ' =======================================================
    ' Convert Az/El to values suitable for plotting on graph.
    ' =======================================================
    Dim sngRadians  As Single
    Dim sngRadius   As Single
    Dim sngX        As Single
    Dim sngY        As Single
    sngRadians = Azimuth * 1.74532925199433E-02
    sngRadius = 1 - (1 / (90 / Elevation))
    sngX = sngRadius * Sin(sngRadians)
    sngY = sngRadius * -Cos(sngRadians)
    UserControl.DrawWidth = 6
    UserControl.PSet (sngX, sngY), RGB(255, 0, 0)
    
End Sub

Private Sub UserControl_Resize()

    Call Reset

End Sub

