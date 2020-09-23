Attribute VB_Name = "modShip"
Option Explicit

Private Type typLanderColl
  Angle As Single
  Dist As Single
End Type

Private Type typLanderLeg
  Angle As Single
  Coll As typLanderColl
End Type

Public Type tSpaceship
  X As Single
  Y As Single
  MoveX As Single
  MoveY As Single
  Angle As Single
  RealAngle As Single
  Thrust As Single
  Speed As Single
  HalfX As Single
  HalfY As Single
  UpCollision(0 To 35, 0 To 34) As Byte
  DownCollision(0 To 35, 0 To 34) As Byte
  LegDist As Single
  Fuel As Single
  Shield As Single
  Landed As Boolean
End Type

Public Spaceship As tSpaceship
Public Const Gravity As Single = 18

Dim SpacePressed As Boolean
Public StickLeft As Boolean, StickRight As Boolean, StickUp As Boolean, FirePressed As Boolean

Public Sub CalculateShip()
 With Spaceship
  Dim XUPD As Single, a As Byte
  Dim Index As Integer
  Dim Angle As Integer, BestAngle As Integer
    
  StickRight = False
  StickLeft = False
  StickUp = False
   
  If .Landed = False Then
   If KeyState(vbKeyRight) = 2 Then
    StickRight = True
   End If
   If KeyState(vbKeyLeft) = 2 Then
    StickLeft = True
   End If
  End If
  
  If KeyState(vbKeyUp) = 2 Then
   StickUp = True
  End If
  
  If KeyState(vbKeySpace) = 2 Then
   
   If SpacePressed = False Then
    SpacePressed = True
    
    Index = CreateEffect
    
    Effect(Index).Effect = cMyShell
    Effect(Index).X = Spaceship.X + 17 + Posun_X(Spaceship.RealAngle, 21)
    Effect(Index).Y = Spaceship.Y - 15 - Posun_Y(Spaceship.RealAngle, 21) - 2
    .Fuel = .Fuel - 0.085
    Effect(Index).Angle = Spaceship.RealAngle
        
    For a = 1 To AliveEnemy.Count
     If Enemy(AliveEnemy(a)).X > gCamX - 10 And Enemy(AliveEnemy(a)).X < gCamX + 650 Then
      Angle = GetAngle(Effect(Index).X, 480 - Effect(Index).Y, Enemy(AliveEnemy(a)).X, 480 - Enemy(AliveEnemy(a)).Y)
      If Abs(Angle - Spaceship.RealAngle) < Abs(BestAngle - Spaceship.Angle) Then BestAngle = Angle
     End If
    Next a
    If Abs(BestAngle - Spaceship.Angle) < 5 Then Effect(Index).Angle = BestAngle
    
   End If
  Else
   SpacePressed = False
  End If
  
  If StickRight Then
   Spaceship.Angle = FixAngle(Spaceship.Angle + GS(190))
  End If
  If StickLeft = True Then
   Spaceship.Angle = FixAngle(Spaceship.Angle - GS(190))
  End If
  
  If StickUp = True Then
   Spaceship.Thrust = 120
   Spaceship.Fuel = Spaceship.Fuel - GS(1)
   If .Landed = True Then
    .Landed = False
    ShipCanLand = False
    CurrentLandZone = CurrentLandZone + 1
   End If
  Else
   Spaceship.Thrust = 0
  End If
  
  If Spaceship.Fuel <= 0 Then
   Spaceship.Fuel = 0
   Spaceship.Thrust = 0
  End If
  
  .RealAngle = Int(.Angle / 10) * 10
  
  XUPD = (.Thrust * Sin(.RealAngle * DegToRad))
  If Abs(.MoveX) > Abs(XUPD) Then
   .MoveX = .MoveX + GS((XUPD - .MoveX) / 5)
  Else
   .MoveX = .MoveX + GS((XUPD - .MoveX))
  End If
  
  .MoveY = .MoveY + GS(.Thrust) * Sin((90 - .RealAngle) * DegToRad)
  If .MoveY < -120 Then
   .MoveY = -120
  ElseIf .MoveY > 80 Then
   .MoveY = 80
  End If
  
  If .Landed = False Then .MoveY = .MoveY - GS(Gravity)
  
  If .MoveX > 0 Then
   .MoveX = .MoveX - GS(2)
  ElseIf .MoveX < 0 Then
   .MoveX = .MoveX + GS(2)
  End If
  
  If .Y > 480 Then .MoveY = -.MoveY / 2: .Y = 480
  
  .Speed = GetDist(0, 0, .MoveX, .MoveY)
  
  If .X + 50 > MapSize Then
   .MoveX = -Abs(.MoveX)
  ElseIf .X - 1 < 0 Then
   .MoveX = Abs(.MoveX)
  End If
  
  .X = .X + GS(.MoveX)
  .Y = .Y + GS(.MoveY)
  
  If .Shield <= 0 Then
   EndIn = 2
   .Shield = 0
   Index = CreateEffect
   Effect(Index).Effect = cExplosion
   Effect(Index).X = .X + 17
   Effect(Index).Y = .Y - 17
   Effect(Index).Size = 40
   Effect(Index).Progress = 0
   Spaceship.Y = 500
  End If
  
  If .Landed = True Then
   If .Fuel < 100 Then
    .Fuel = .Fuel + GS(15)
   Else
    .Fuel = 100
   End If
   
   If CurrentLandZone = LevelCount Then
    EndGame = True
    frmGame.fraStatus.Visible = True
    frmGame.lblStatus = "Success !!!"
    frmGame.lblEnemies = UBound(Enemy) - AliveEnemy.Count
    frmGame.lblLevel = CurrentLandZone
    frmGame.lblTime = Int(GameTime / 60) & " min, " & GameTime Mod 60 & " s"
   End If
   
  Else
  
   If .X + 7 > LandZone(CurrentLandZone) And .X + 27 < LandZone(CurrentLandZone) + 50 And .Y - 31 < Terrain(.X) And .RealAngle = 0 And ShipCanLand = True Then
    If .RealAngle = 0 And Abs(.MoveX) < 3 And .MoveY > -20 Then
     .MoveY = 0
     .Y = Terrain(.X) + 31
     .MoveX = 0
     .Landed = True
     .Shield = .Shield + 20
     If .Shield > 100 Then .Shield = 100
    Else
     .MoveY = -.MoveY / 2
     .MoveX = .MoveX / 2
     .Shield = .Shield - .MoveY / 2
     .Y = Terrain(.X) + 31
     .MoveX = 0
    End If
   
   Else
   
    For a = 0 To 34
   
     If .Y - .DownCollision(.RealAngle / 10, a) + 3 < Terrain(.X + a) Then
      If .Fuel > 0 Then
       .MoveY = 50
       .Shield = .Shield - 30
       .X = .X + -.MoveX / 10
       .Y = .Y + .MoveY / 10
       .MoveX = 0
      Else
       .Shield = 0
       .MoveY = 0
       .MoveX = 0
      End If
     End If
   
    Next a
    Dim BeingTargeted As Boolean
    BeingTargeted = False
    
    If .X > RadarScan(CurrentRadarScan) Then
     CurrentRadarScan = CurrentRadarScan + 1
     RadarScanTimeLeft = 6
     TimeTillTargeted = 15
    End If
    
    If RadarScanTimeLeft > 0 Then
    For a = 0 To 34
   
     If .Y - .UpCollision(.RealAngle / 10, a) > Radar(.X + a) And .UpCollision(.RealAngle / 10, a) < 255 Then
      BeingTargeted = True
      TimeTillTargeted = TimeTillTargeted - GS(1)
      If TimeTillTargeted <= 0 Then
       RadarScanTimeLeft = 0
       Index = CreateEffect
       
       Effect(Index).Effect = cBFRocket
       Effect(Index).X = .X + 600
       Effect(Index).Y = 400
       Effect(Index).Angle = 270
      End If
      Exit For
     End If
   
    Next a
    If BeingTargeted = False Then
     RadarScanTimeLeft = RadarScanTimeLeft - GS(1)
    End If
    
    End If
   End If
  End If
 End With
End Sub

Function Posun_X(ByVal Uhol As Integer, ByVal Rychlost As Integer) As Double
  Posun_X = Rychlost * Sin(Uhol * DegToRad)
End Function

Function Posun_Y(ByVal Uhol As Integer, ByVal Rychlost As Integer) As Double
  Posun_Y = -Rychlost * Sin((90 - Uhol) * DegToRad)
End Function

