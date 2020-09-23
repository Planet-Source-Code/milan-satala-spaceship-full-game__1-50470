Attribute VB_Name = "modGrafika"
Option Explicit

Public DirX As DirectX7
Public DDraw As DirectDraw7

Public sPrimary As DirectDrawSurface7
Public PrimaryRect As RECT
Public sBack As DirectDrawSurface7
Public BackRect As RECT

Public DefaultDesc As DDSURFACEDESC2

Public sTerrain() As DirectDrawSurface7
Public TerrainRect As RECT

Public sTerrainDirt As DirectDrawSurface7, sPlants As DirectDrawSurface7, sCloud As DirectDrawSurface7
Public rTerrainDirt As RECT, rPlants As RECT, rCloud As RECT

Public sSoldier As DirectDrawSurface7, sTank As DirectDrawSurface7, sBFRocket As DirectDrawSurface7, sSAM As DirectDrawSurface7, sRocket As DirectDrawSurface7
Public rSoldier As RECT, rTank As RECT, rBFRocket As RECT, rSAM As RECT, rRocket As RECT

Public sShip As DirectDrawSurface7, ShipDesc As DDSURFACEDESC2
Public ShipRect As RECT

Public Size As Long

Public gMoveX As Single, gCamX As Single, OnScrX As Single

Public ColorOfSky As Long, ColorOfGround As Long, ColorOfLandzone As Long, ColorOfGrass As Long, PurpleColor As Long, DarkRed As Long

Private Declare Function AlphaBlend Lib "msimg32" ( _
ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, _
ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean

' API DECLARATIONS [ COPY MEMORY FUNCTION ]
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
Destination As Any, Source As Any, ByVal Length As Long)

' TYPE STRUCTURES
Private Type typeBlendProperties
    tBlendOp As Byte
    tBlendOptions As Byte
    tBlendAmount As Byte
    tAlphaType As Byte
End Type

Public Sub Render()
  Dim TRect1 As RECT, TRect2 As RECT
  Dim SRect As RECT
  Dim CRect As RECT
  Dim X As Integer, Y As Integer, Angle As Single
  Dim T1 As Byte, T2 As Byte
  Dim a As Integer, b As Integer
  
  sBack.SetForeColor ColorOfSky
  sBack.SetFillColor ColorOfSky
  sBack.DrawBox 0, 0, 640, 480
  
  TRect1.Bottom = 300
  TRect2.Bottom = 300
  Y = Int(Spaceship.RealAngle / 90)
  SRect.Left = Spaceship.RealAngle * 3.5 - Y * 315
  SRect.Top = Y * 35
  
  SRect.Bottom = SRect.Top + 35
  SRect.Right = SRect.Left + 35
  T1 = GetTerrainNum(gCamX)
  T2 = GetTerrainNum(gCamX + 640)
  
  TRect1.Left = gCamX - (T1 - 1) * 640
  TRect1.Right = 640
  TRect2.Left = 0
  TRect2.Right = TRect1.Left
  
  For a = 1 To AliveEnemy.Count
   With Enemy(AliveEnemy(a))
    If .What = cSAM Then
     BltFastAndCull .X - gCamX - 12, 480 - .Y - 23, sSAM, rSAM
     If .TimeToNextFire < 2 Then
      rRocket.Left = 775
      rRocket.Left = Int(.WeaponAngle / 10) * 25
      rRocket.Right = rRocket.Left + 25
      rRocket.Bottom = rRocket.Top + 26
      BltFastAndCull .X - gCamX - 13, 480 - .Y - 34, sRocket, rRocket
     End If
    End If
   End With
  Next a
  
  sBack.BltFast 0, 180, sTerrain(T1), TRect1, DDBLTFAST_SRCCOLORKEY
  sBack.BltFast 640 - TRect1.Left, 180, sTerrain(T2), TRect2, DDBLTFAST_SRCCOLORKEY
    
    
  For a = 0 To 319
   b = a * 2
   sBack.SetForeColor PurpleColor
   sBack.setDrawWidth 1
   If RadarScanTimeLeft > 0 Then sBack.DrawLine b, 480 - Radar(gCamX + b), b + 2, 480 - Radar(gCamX + b + 2)
   
  Next a
  
  rTerrainDirt.Top = 100
  rTerrainDirt.Bottom = 120
  For a = CurrentLandZone - 1 To CurrentLandZone + 1
   If a >= 1 And a <= LevelCount Then
    If a = CurrentLandZone And ShipCanLand = True Then
     rTerrainDirt.Left = 10
     rTerrainDirt.Right = 20
    Else
     rTerrainDirt.Left = 0
     rTerrainDirt.Right = 10
    End If
    BltFastAndCull LandZone(a) - gCamX + 40, 480 - Terrain(LandZone(a)) - 19, sTerrainDirt, rTerrainDirt
   End If
  Next a
  
  For a = 1 To AliveEnemy.Count
   If AliveEnemy(a) > EnemiesInLevel(CurrentLandZone + 1) Then Exit For
   With Enemy(AliveEnemy(a))
    If .X + 20 > gCamX And .X < gCamX + 660 Then
    Select Case .What
     Case cSoldier
      .WeaponY = -13
      
      Select Case .WeaponAngle
       Case 0 To 30
        rSoldier.Top = 0
        rSoldier.Left = 20
        .WeaponX = 1
       Case 30 To 60
        rSoldier.Top = 0
        rSoldier.Left = 0
        .WeaponX = 1
       Case 60 To 180
        rSoldier.Top = 0
        rSoldier.Left = 40
        .WeaponX = 1
       Case 330 To 360
        rSoldier.Top = 20
        rSoldier.Left = 20
        .WeaponX = -1
       Case 300 To 330
        rSoldier.Top = 20
        rSoldier.Left = 40
        .WeaponX = -1
       Case 180 To 300
        rSoldier.Top = 20
        rSoldier.Left = 0
        .WeaponX = -1
      End Select
      
      rSoldier.Right = rSoldier.Left + 20
      rSoldier.Bottom = rSoldier.Top + 20
      BltFastAndCull .X - gCamX - 10, 480 - .Y - 20, sSoldier, rSoldier
      sBack.setDrawWidth 2
      sBack.SetForeColor 0
      
      sBack.DrawLine .X - gCamX + .WeaponX, 480 - .Y + .WeaponY, .X - gCamX + .WeaponX + Sin(.WeaponAngle * DegToRad) * 11, 480 - .Y + .WeaponY - Cos(.WeaponAngle * DegToRad) * 11

     Case cTank
      Select Case .Angle
       Case 0 To 90
        rTank.Left = Int((.Angle - 10) / 5) * 40
        rTank.Top = 42
       Case 90 To 180
        rTank.Left = Int((.Angle - 90) / 5) * 40
        rTank.Top = 0
       Case 180 To 270
        rTank.Left = Int((.Angle - 190) / 5) * 40
        rTank.Top = 42
       Case 270 To 360
        rTank.Left = Int((.Angle - 280) / 5) * 40
        rTank.Top = 0
       End Select
       
      rTank.Right = rTank.Left + 40
      
      rTank.Bottom = rTank.Top + 42
      sBack.setDrawWidth 2
      sBack.SetForeColor 0
      BltFastAndCull .X - gCamX - 20, 480 - .Y - 19, sTank, rTank
      
      If rTank.Top = 0 Then
       .WeaponX = 2 + 11 * Sin((.Angle - 90) * DegToRad)
       .WeaponY = -11 * Cos((.Angle - 90) * DegToRad)
      Else
       .WeaponX = -3 + 11 * Sin((.Angle - 90) * DegToRad)
       .WeaponY = -11 * Cos((.Angle - 90) * DegToRad)
      End If
      
      sBack.DrawLine .X - gCamX + .WeaponX, 480 - .Y + .WeaponY, .X - gCamX + .WeaponX + Sin(.WeaponAngle * DegToRad) * 10, 480 - .Y + .WeaponY - Cos(.WeaponAngle * DegToRad) * 10
            
    End Select
    
    End If
   End With
  Next a
    
  With Spaceship
   OnScrX = .X - gCamX
   If .Shield > 0 Then
    If .Thrust Then
     Angle = FixAngle(180 + .RealAngle)
     X = OnScrX + 17 + Posun_X(Angle, 15)
     Y = 480 - (.Y - 17 - Posun_Y(Angle, 15))
     If Rnd * 100 > 50 Then
      sBack.SetForeColor vbRed
     Else
      sBack.SetForeColor vbYellow
     End If
     sBack.setDrawWidth 7
     sBack.DrawLine X, Y, X + Posun_X(Angle, 6), Y + Posun_Y(Angle, 6)
    End If
    sBack.BltFast OnScrX, 480 - Spaceship.Y, sShip, SRect, DDBLTFAST_SRCCOLORKEY
   End If
  End With
    
  sBack.setDrawWidth 1
  For a = 1 To UsedEffect.Count
   With Effect(UsedEffect(a))
    Select Case .Effect
     Case cMyShell
      sBack.setDrawWidth 1
      sBack.SetForeColor 0
      sBack.SetFillColor PurpleColor
      sBack.DrawCircle .X - gCamX, 480 - .Y, 4
     
     Case cTankShell
      sBack.setDrawWidth 1
      sBack.SetForeColor 0
      sBack.SetFillColor vbRed
      sBack.DrawCircle .X - gCamX, 480 - .Y, 4
     
     Case cSoldierShell
      sBack.setDrawWidth 1
      sBack.SetForeColor 0
      sBack.SetFillColor vbRed
      sBack.DrawCircle .X - gCamX, 480 - .Y, 3
     
     Case cBFRocket
     
      Y = Int(.Angle / 90)
      rBFRocket.Left = Int(.Angle / 10) * 42 - Y * 378
      rBFRocket.Top = Y * 42
      rBFRocket.Right = rBFRocket.Left + 42
      rBFRocket.Bottom = rBFRocket.Top + 42
      sBack.setDrawWidth 5
      If Rnd * 100 > 50 Then
       sBack.SetForeColor vbRed
      Else
       sBack.SetForeColor vbYellow
      End If
      sBack.DrawLine .X - gCamX + 20 - Sin(Int(.Angle / 10) * 10 * DegToRad) * 19, 480 - .Y + 21 + Cos(Int(.Angle / 10) * 10 * DegToRad) * 19, _
      .X - gCamX + 20 - Sin(Int(.Angle / 10) * 10 * DegToRad) * 25, 480 - .Y + 21 + Cos(Int(.Angle / 10) * 10 * DegToRad) * 25
      BltFastAndCull .X - gCamX, 480 - .Y, sBFRocket, rBFRocket
      
     Case cRocket
      rRocket.Left = 775
      rRocket.Left = Int(.Angle / 10) * 25
      rRocket.Right = rRocket.Left + 25
      rRocket.Bottom = rRocket.Top + 26
      If Rnd * 100 > 50 Then
       sBack.SetForeColor vbRed
      Else
       sBack.SetForeColor vbYellow
      End If
      sBack.setDrawWidth 5
      If .Progress < 10 Then sBack.DrawLine .X - gCamX - 1 - Sin(Int(.Angle / 10) * 10 * DegToRad) * 12, 480 - .Y + Cos(Int(.Angle / 10) * 10 * DegToRad) * 12, _
      .X - gCamX - 1 - Sin(Int(.Angle / 10) * 10 * DegToRad) * 15, 480 - .Y + Cos(Int(.Angle / 10) * 10 * DegToRad) * 15
      
      BltFastAndCull .X - gCamX - 14, 480 - .Y - 14, sRocket, rRocket
           
     Case cFlesh
      If .X - 5 > gCamX And .X < gCamX + 635 Then
       sBack.setDrawWidth 1
       sBack.SetForeColor DarkRed
       sBack.SetFillColor vbRed
       sBack.DrawCircle .X - gCamX, 480 - .Y, 2
      End If
            
     Case cExplosion
     sBack.setDrawWidth 2
      If Rnd * 10 > 5 Then
       sBack.SetForeColor vbRed
       sBack.SetFillColor vbYellow
      Else
       sBack.SetForeColor vbYellow
       sBack.SetFillColor vbRed
      End If
      sBack.DrawCircle .X - gCamX, 480 - .Y, .Progress
      
    End Select
   End With
  Next a
  
  For a = 1 To UBound(Cloud)
   If Cloud(a).X + rCloud.Right > gCamX And Cloud(a).X < gCamX + 640 Then
    BltFastAndCull Cloud(a).X - gCamX, Cloud(a).Y, sCloud, rCloud
   End If
  Next a
      
  sBack.setDrawWidth 3
  
  sBack.SetFont frmGame.lblFont.Font
  sBack.SetForeColor PurpleColor
  sBack.DrawText 40, 40, "Fuel: ", False
  sBack.DrawText 520, 40, "Shield: ", False
  sBack.SetFont frmGame.lblFont2.Font
  
  sBack.DrawText 40, 60, Int(Spaceship.Fuel), False
  If Spaceship.Shield = 100 Then
   sBack.DrawText 520, 60, Int(Spaceship.Shield), False
  Else
   sBack.DrawText 537, 60, Int(Spaceship.Shield), False
  End If
  
  If RadarScanTimeLeft > 0 Then
   sBack.SetFont frmGame.lblFontRadar.Font
   sBack.SetForeColor vbRed
   sBack.DrawText 250, 40, "Radar scan !!!", False
   sBack.DrawText 190, 60, "Time till targeted: " & Int(TimeTillTargeted) & " seconds", False
  End If
  
  DirX.GetWindowRect frmGame.hWnd, PrimaryRect

  PrimaryRect.Top = PrimaryRect.Top + 22
  PrimaryRect.Left = PrimaryRect.Left + 4
  PrimaryRect.Right = PrimaryRect.Right - 4
  PrimaryRect.Bottom = PrimaryRect.Bottom - 4
 
  sPrimary.Blt PrimaryRect, sBack, BackRect, DDBLT_WAIT
  
End Sub

Sub BltFastAndCull(X As Long, Y As Long, S As DirectDrawSurface7, R As RECT)
  Dim CulledRect As RECT
  CulledRect = R
  If X < 0 Then
   CulledRect.Left = R.Left - X
  ElseIf X + R.Right - R.Left > 640 Then
   CulledRect.Right = R.Left + 640 - X
  End If
  If Y + R.Bottom - R.Top > 480 Then CulledRect.Bottom = R.Top + 480 - Y
  sBack.BltFast X - R.Left + CulledRect.Left, Y, S, CulledRect, DDBLTFAST_SRCCOLORKEY
End Sub
