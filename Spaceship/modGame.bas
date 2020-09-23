Attribute VB_Name = "modGame"
Option Explicit

Public Sub GameTick()
  Dim Time As Long
  Dim Index As Integer
  Dim Y As Long
  Dim Speed As Integer
  Dim Angle As Single
  
  Time = GetTickCount - OldTime
  Dim a, b
  b = GetTickCount
  OldTime = GetTickCount
  
  If Time > 1000 Then Exit Sub
  
  GameSpeed = Time / 1000
  
  If KeyState(vbKeyEscape) = 1 Then End
  If KeyState(vbKeyP) = 1 Then
   If Paused = True Then Paused = False Else Paused = True
  End If
  If Paused = True Then DoEvents: Exit Sub
  
  If GetTickCount - DrawTime > 10 Then
   DrawTime = GetTickCount
   DoEvents
   Render
  End If
    
  GameTime = GameTime + GS(1)
  
  If EndIn > -1 Then
   EndIn = EndIn - GS(1)
   If EndIn <= 0 Then
    EndGame = True
    frmGame.fraStatus.Visible = True
    frmGame.lblStatus = "Failure !!!"
    frmGame.lblEnemies = UBound(Enemy) - AliveEnemy.Count
    frmGame.lblLevel = CurrentLandZone - 1
    frmGame.lblTime = Int(GameTime / 60) & " min, " & GameTime Mod 60 & " s"
   End If
  Else
   CalculateShip
  End If
    
  If OnScrX > 300 Then
   gMoveX = 100
   gCamX = Spaceship.X - 300
  ElseIf OnScrX < 200 Then
   gMoveX = -100
   gCamX = Spaceship.X - 200
  End If
  
  If Spaceship.X < gCamX Then
   
  ElseIf Spaceship.X > gCamX + 600 Then
   gCamX = Spaceship.X - 600
  End If
  
  If gMoveX > 0 Then
   gMoveX = gMoveX - GS(50)
   If gMoveX < 0 Then gMoveX = 0
  ElseIf gMoveX < 0 Then
   gMoveX = gMoveX + GS(50)
   If gMoveX > 0 Then gMoveX = 0
  End If
  gCamX = gCamX + GS(gMoveX)
  
  If gCamX + 640 > MapSize Then
   gCamX = MapSize - 640
   gMoveX = 0
  ElseIf gCamX < 0 Then
   gCamX = 0
   
   gMoveX = 0
  End If
  
  a = 0
  
  Dim HighestTerrainLeft As Integer, HighestTerrainRight As Integer
  
  Do Until a = UsedEffect.Count
   a = a + 1
   
   With Effect(UsedEffect(a))
    Select Case .Effect
     Case cMyShell
      .X = .X + GS(400 * Sin(.Angle * DegToRad))
      .Y = .Y + GS(400 * Cos(.Angle * DegToRad))
      
      If Terrain(.X) >= .Y Then
       .Effect = cExplosion
       .Size = 10
       .Progress = 0
      End If
      If .Y > 480 Or Abs(.X - Spaceship.X) > 1000 Or .X < 10 Or .X > MapSize - 10 Then
       UnusedEffect.Add UsedEffect(a)
       UsedEffect.Remove a
       a = a - 1
      End If
      If Terrain(.X) > .Y - 50 Then
      For b = 1 To AliveEnemy.Count
       If AliveEnemy(b) > EnemiesInLevel(CurrentLandZone + 1) Then Exit For
       Select Case Enemy(AliveEnemy(b)).What
        Case cTank
         If Abs(Enemy(AliveEnemy(b)).X - .X) < 20 And Abs(Enemy(AliveEnemy(b)).Y - .Y) < 10 Then
          Enemy(AliveEnemy(b)).HitPoints = Enemy(AliveEnemy(b)).HitPoints - 1
          .Effect = cExplosion
          .Size = 7
          .Progress = 0
         End If
        Case cSoldier
         If Abs(Enemy(AliveEnemy(b)).X + 2 - .X) < 12 And Abs(Enemy(AliveEnemy(b)).Y - .Y) < 20 Then
          Enemy(AliveEnemy(b)).HitPoints = Enemy(AliveEnemy(b)).HitPoints - 1
          .Effect = cExplosion
          .Size = 7
          .Progress = 0
         End If
        Case cSAM
         If Abs(Enemy(AliveEnemy(b)).X + 2 - .X) < 12 And .Y < Enemy(AliveEnemy(b)).Y + 15 Then
          Enemy(AliveEnemy(b)).HitPoints = Enemy(AliveEnemy(b)).HitPoints - 1
          .Effect = cExplosion
          .Size = 7
          .Progress = 0
         End If
        
       End Select
      Next b
      End If
     Case cTankShell To cSoldierShell
      .X = .X + GS(170 * Sin(.Angle * DegToRad))
      .Y = .Y + GS(170 * Cos(.Angle * DegToRad))
      
      If Terrain(.X) >= .Y Then
       .Effect = cExplosion
       .Size = 7
       .Progress = 0
      End If
      
      If .X - Spaceship.X <= 35 And Spaceship.Y - .Y <= 35 And .Y < Spaceship.Y And .X > Spaceship.X Then

        Y = Int(Spaceship.RealAngle / 90)
        If sShip.GetLockedPixel((Spaceship.RealAngle * 3.5) - Y * 315 + .X - Spaceship.X, Y * 35 + Spaceship.Y - .Y) <> 0 Then
         If .Effect = cTankShell Then
          Spaceship.Shield = Spaceship.Shield - 4
         Else
          Spaceship.Shield = Spaceship.Shield - 2
         End If
         Effect(UsedEffect(a)).Effect = cExplosion
         Effect(UsedEffect(a)).Size = 7
         Effect(UsedEffect(a)).Progress = 0
        End If
        
      End If
      
      If .Y > 480 Or Abs(.X - Spaceship.X) > 1000 Or .X < 10 Or .X > MapSize - 10 Then
       UnusedEffect.Add UsedEffect(a)
       UsedEffect.Remove a
       a = a - 1
      End If
     
     
     Case cExplosion
      .Progress = .Progress + GS(50)
      If .Progress >= .Size Then
       UnusedEffect.Add UsedEffect(a)
       UsedEffect.Remove a
       a = a - 1
      End If
      
     Case cFlesh
      If .Y <= Terrain(.X) Then
       .Y = Terrain(.X)
       HighestTerrainLeft = .X - 1
       HighestTerrainRight = .X + 1
       
       For b = 1 To 5
         If GetDist(0, Terrain(.X), b, Terrain(.X - b)) <= 5 Then
          HighestTerrainLeft = .X - b
         End If
         If GetDist(0, Terrain(.X), b, Terrain(.X + b)) <= 5 Then
          HighestTerrainRight = .X + b
         End If
       Next b
       
       If HighestTerrainLeft < LandZone(CurrentLandZone) Or HighestTerrainRight > LandZone(CurrentLandZone) + 50 Then
        For b = HighestTerrainLeft To HighestTerrainRight
         Terrain(b) = Terrain(b) + 1
        Next b
        b = GetTerrainNum(HighestTerrainLeft)
       
        sTerrain(b).SetForeColor RGB(230, 0, 0)
        sTerrain(b).setDrawWidth 3
        sTerrain(b).DrawLine HighestTerrainLeft - (b - 1) * 640, 302 - Terrain(HighestTerrainLeft), HighestTerrainRight - (b - 1) * 640, 302 - Terrain(HighestTerrainRight)
       End If
       
       UnusedEffect.Add UsedEffect(a)
       UsedEffect.Remove a
       a = a - 1
       
      Else
       .X = .X + GS(.MoveX)
       .Y = .Y + GS(.MoveY)
       .MoveY = .MoveY - GS(Gravity * 2)
       If .MoveX > 0 Then .Angle = FixAngle(.Angle + GS(10)) Else .Angle = FixAngle(.Angle - GS(10))
      End If
     
     Case cBFRocket
            
      Angle = GetAngle(.X, -.Y, Spaceship.X, -Spaceship.Y)
      If Spaceship.Shield <= 0 Then Angle = 180
      
      If Angle < 230 Or .Y <> 400 Then
       If SubtractAngles(.Angle, Angle) > 0 Then
        .Angle = FixAngle(.Angle + GS(120))
       Else
        .Angle = FixAngle(.Angle - GS(120))
       End If
      End If
      .X = .X + GS(Sin(.Angle * DegToRad) * 150)
      .Y = .Y + GS(Cos(.Angle * DegToRad) * 150)
      .MoveX = .X + 22 + Sin(Int(.Angle / 10) * 10 * DegToRad) * 17
      .MoveY = -21 + .Y + Cos(Int(.Angle / 10) * 10 * DegToRad) * 17
      Y = Int(Spaceship.RealAngle / 90)
      If .MoveX - Spaceship.X <= 35 And Spaceship.Y - .MoveY <= 35 And .MoveY < Spaceship.Y And .MoveX > Spaceship.X Then
       If sShip.GetLockedPixel((Spaceship.RealAngle * 3.5) - Y * 315 + .MoveX - Spaceship.X, Y * 35 + Spaceship.Y - .MoveY) <> 0 Then
        .Effect = cExplosion
        .Size = 10
        .Progress = 0
        .X = .MoveX
        .Y = .MoveY
        Spaceship.Shield = Spaceship.Shield - 50
       End If
      End If
     
     Case cRocket
      If Spaceship.Shield <= 0 Then .Progress = 10
      If .Progress < 10 Then
       Angle = GetAngle(.X, -.Y, Spaceship.X + 17, -Spaceship.Y + 17)
      Else
       Angle = 180
      End If
       Angle = SubtractAngles(.Angle, Angle)
       If Abs(Angle) < 1 Then
        .Angle = .Angle - Angle
       ElseIf Angle > 0 Then
        .Angle = FixAngle(.Angle + GS(35))
       ElseIf Angle < 0 Then
        .Angle = FixAngle(.Angle - GS(35))
       End If
       .Progress = .Progress + GS(1)
      
       .X = .X + GS(Sin(.Angle * DegToRad) * 140)
       .Y = .Y + GS(Cos(.Angle * DegToRad) * 140)
      
      
      .MoveX = .X + Sin(Int(.Angle / 10) * 10 * DegToRad) * 10
      .MoveY = .Y + Cos(Int(.Angle / 10) * 10 * DegToRad) * 10
      If Terrain(.MoveX) > .MoveY Then
       .Effect = cExplosion
       .Size = 15
       .Progress = 0
       .X = .MoveX
       .Y = .MoveY
      End If
      Y = Int(Spaceship.RealAngle / 90)
      If .MoveX - Spaceship.X <= 35 And Spaceship.Y - .MoveY <= 35 And .MoveY < Spaceship.Y And .MoveX > Spaceship.X Then
       If sShip.GetLockedPixel((Spaceship.RealAngle * 3.5) - Y * 315 + .MoveX - Spaceship.X, Y * 35 + Spaceship.Y - .MoveY) <> 0 Then
        .Effect = cExplosion
        .Size = 10
        .Progress = 0
        .X = .MoveX
        .Y = .MoveY
        Spaceship.Shield = Spaceship.Shield - 15
       End If
      End If
      
    End Select
   End With
  Loop
  
  a = 0
  
  Dim InLOS As Boolean
    
  If AliveEnemy.Count = 0 Then
   ShipCanLand = True
  Else
   If AliveEnemy(1) > EnemiesInLevel(CurrentLandZone) Then
    ShipCanLand = True
   End If
  End If
      
  If Spaceship.Shield <= 0 Then Exit Sub
      
  Do Until a = AliveEnemy.Count
   a = a + 1
   If AliveEnemy(a) > EnemiesInLevel(CurrentLandZone + 1) Then Exit Do
   With Enemy(AliveEnemy(a))
    Select Case .What
     Case cSoldier
      .Y = Terrain(.X)
            
      If .X > gCamX And .X < gCamX + 640 Then
       .TimeToNextFire = .TimeToNextFire - GS(1)
       .WeaponAngle = GetAngle(.X + .WeaponX, -.Y - .WeaponY, Spaceship.X + 17, -Spaceship.Y + 30)
      End If
      If .TimeToNextFire < 0 Then
       .TimeToNextFire = 3
       
       Index = CreateEffect
       Effect(Index).Effect = cSoldierShell
       Effect(Index).X = .X + .WeaponX + 10 * Sin(.WeaponAngle * DegToRad)
       Effect(Index).Y = .Y - .WeaponY + 10 * Cos(.WeaponAngle * DegToRad)
       Effect(Index).Angle = .WeaponAngle
      End If
      
      If .HitPoints <= 0 Then
       AliveEnemy.Remove a
       a = a - 1
       .What = 0
       
       For b = 1 To 10
        Index = CreateEffect
        Effect(Index).Effect = cFlesh
        Effect(Index).X = .X
        Effect(Index).Y = .Y + 5
        Effect(Index).Angle = FixAngle(-40 + Rnd * 80)
        Speed = 10 + Rnd * 30
        Effect(Index).MoveX = Sin(Effect(Index).Angle * DegToRad) * Speed
        Effect(Index).MoveY = Cos(Effect(Index).Angle * DegToRad) * Speed
       
        Effect(Index).Progress = 0
       Next b
      End If
      
     Case cTank
      InLOS = True
             
       If .X > gCamX And .X < gCamX + 640 Then
        .TimeToNextFire = .TimeToNextFire - GS(1)
        .WeaponAngle = GetAngle(.X + .WeaponX, -.Y - .WeaponY, Spaceship.X + 17, -Spaceship.Y + 30)
       
        If .WeaponAngle > .Angle And .WeaponAngle < 180 Then
         .WeaponAngle = .Angle
         InLOS = False
         .TimeToNextFire = Rnd * 2
        ElseIf .WeaponAngle < .Angle + 180 And .WeaponAngle > 180 Then
         .WeaponAngle = .Angle + 180
         InLOS = False
         .TimeToNextFire = Rnd * 2
        End If
       End If
       
       If .TimeToNextFire < 0 And InLOS = True Then
        .TimeToNextFire = 3
        
        Index = CreateEffect
        Effect(Index).Effect = cTankShell
        Effect(Index).X = .X + .WeaponX + 10 * Sin(.WeaponAngle * DegToRad)
        Effect(Index).Y = .Y - .WeaponY + 10 * Cos(.WeaponAngle * DegToRad)
        Effect(Index).Angle = .WeaponAngle
       End If
       
       If .HitPoints <= 0 Then
        AliveEnemy.Remove a
        a = a - 1
        .What = 0
        
        Index = CreateEffect
        Effect(Index).Effect = cExplosion
        Effect(Index).X = .X
        Effect(Index).Y = .Y
        Effect(Index).Size = 15
        Effect(Index).Progress = 0
              
       End If
       
      Case cSAM
              
       If .X > gCamX And .X < gCamX + 640 Then
        .TimeToNextFire = .TimeToNextFire - GS(1)
        .WeaponAngle = GetAngle(.X + .WeaponX, -.Y - .WeaponY, Spaceship.X + 17, -Spaceship.Y + 30)
        
        If .WeaponAngle > 50 And .WeaponAngle < 180 Then
         .WeaponAngle = 50
        ElseIf .WeaponAngle < 310 And .WeaponAngle > 180 Then
         .WeaponAngle = 310
        End If
       End If
       
       If .TimeToNextFire < 0 Then
        .TimeToNextFire = 11
        
        Index = CreateEffect
        Effect(Index).Effect = cRocket
        Effect(Index).X = .X
        Effect(Index).Y = .Y + 20
        Effect(Index).Angle = .WeaponAngle
       End If
       
       If .HitPoints <= 0 Then
        AliveEnemy.Remove a
        a = a - 1
        .What = 0
        
        Index = CreateEffect
        Effect(Index).Effect = cExplosion
        Effect(Index).X = .X
        Effect(Index).Y = .Y
        Effect(Index).Size = 25
        Effect(Index).Progress = 0
              
       End If
       
    End Select
   End With
  Loop
  
End Sub
