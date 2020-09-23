Attribute VB_Name = "modTerrain"
Option Explicit

Public Terrain() As Integer
Public Cloud() As tXY
Public Radar() As Integer
Public LandZone() As Integer
Public CurrentLandZone As Integer
Public ShipCanLand As Boolean
Public Const LevelCount As Byte = 5
Public EnemiesInLevel(0 To LevelCount + 1) As Integer

Public RadarScan() As Integer, RadarScanTimeLeft As Single, CurrentRadarScan As Integer
Public TimeTillTargeted As Single

Public EnemiesDestroyed As Integer

Public Sub GenerateTerrain()
  Dim a As Integer, b As Integer, T As Byte, X As Long
  Dim c As Integer
  Dim SizeLeft As Long, Size As Long, Tall As Single, NextTall As Long
  Dim CurX As Long, CurTall As Single, Move As Single
  Dim Grass1 As Long, Grass2 As Long
  Dim Grass As Long, NajGrass As Long
  Dim NextLandZone As Long, IsLandzone As Boolean
  
  SizeLeft = MapSize
  NextTall = Rnd * 300
  CurTall = Rnd * 300
  NextLandZone = 4 * 640
  ReDim LandZone(0)
  
  Do Until SizeLeft <= 0
   Size = 20 + Rnd * 80
       
   Tall = NextTall '/ 2
   NextTall = 0
   If IsLandzone = False Then
    Do Until NextTall > 10 And NextTall < 300
     NextTall = CurTall - 150 + Rnd * 300
    Loop
   Else
    Tall = Tall / 2 - 100 + Rnd * 200
    If Tall < 10 Then Tall = 10
    If Tall > 300 Then Tall = 300
    Do Until NextTall > 10 And NextTall < 300
     NextTall = Tall - 100 + Rnd * 200
    Loop
   End If
   
   If CurX > NextLandZone Then
    Tall = 20 + Rnd * 130
    NextTall = 2 * Tall
    Size = 50
    NextLandZone = NextLandZone + 3 * 640
    IsLandzone = True
    ReDim Preserve LandZone(UBound(LandZone) + 1)
    LandZone(UBound(LandZone)) = CurX + Size
    
   Else
    IsLandzone = False
   End If
   
   If SizeLeft - 2 * Size < 0 Then Size = SizeLeft / 2
   SizeLeft = SizeLeft - 2 * Size
   
   Move = (Tall - CurTall) / Size
   
   Grass1 = Sqr((Tall - CurTall) ^ 2 + Size ^ 2) * 7 / Size
   Grass2 = Sqr((NextTall / 2 - Tall) ^ 2 + Size ^ 2) * 8 / Size
   If Move > 0 Then NajGrass = CurTall - 7 Else NajGrass = Tall - 7
   
   For a = 1 To Size
    CurX = CurX + 1
    T = GetTerrainNum(CurX)
    
    CurTall = CurTall + Move
    X = CurX - (T - 1) * 640
    Grass = CurTall - Grass1
    If Grass < NajGrass Then Grass = NajGrass
    
    sTerrain(T).SetForeColor ColorOfGround
    sTerrain(T).DrawLine X, 300, X, 300 - Grass
    sTerrain(T).SetForeColor vbGreen
    sTerrain(T).DrawLine X, 300 - Grass, X, 300 - CurTall
    
    Terrain(CurX) = CurTall
   Next a
   
   If IsLandzone = False Then
    Move = ((NextTall / 2) - CurTall) / Size
    If Move > 0 Then NajGrass = CurTall - 8 Else NajGrass = NextTall / 2 - 8
   Else
    Move = 0
    Grass2 = 9
   End If
   
   
   For a = 1 To Size
    CurX = CurX + 1
    T = GetTerrainNum(CurX)
    
    Grass = CurTall - Grass2
    If Grass < NajGrass And IsLandzone = False Then Grass = NajGrass
    
    CurTall = CurTall + Move
    X = CurX - (T - 1) * 640
    
    sTerrain(T).SetForeColor ColorOfGround
    sTerrain(T).DrawLine X, 300, X, 300 - Grass
    
    If IsLandzone = False Then sTerrain(T).SetForeColor vbGreen Else sTerrain(T).SetForeColor ColorOfLandzone
    
    sTerrain(T).DrawLine X, 300 - Grass, X, 300 - CurTall
    
    Terrain(CurX) = CurTall
   Next a
  Loop
  
  CurX = Rnd * 50
  
  Do Until CurX >= MapSize
   T = GetTerrainNum(CurX)
   X = CurX - (T - 1) * 640
   
   rTerrainDirt.Left = 0
   rTerrainDirt.Right = 20
   rTerrainDirt.Top = Int(Rnd * 4) * 20
   If rTerrainDirt.Top > 80 Then rTerrainDirt.Top = 60
   rTerrainDirt.Bottom = rTerrainDirt.Top + 20
   
   sTerrain(T).BltFast X - 10, 280 - Rnd * (Terrain(CurX) - 40), sTerrainDirt, rTerrainDirt, DDBLTFAST_SRCCOLORKEY
   
   CurX = CurX + Rnd * 100
  Loop
  
  CurX = Rnd * 50
  a = 1
  
  Do Until CurX >= MapSize
   T = GetTerrainNum(CurX)
   X = CurX - (T - 1) * 640
   
   rPlants.Left = 0
   rPlants.Right = 10
   rPlants.Top = Int(Rnd * 12) * 10
   If rPlants.Top > 60 Then rPlants.Top = 0
   rPlants.Bottom = rPlants.Top + 10
   
   If Not (CurX > LandZone(a) And CurX < LandZone(a) + 50) Then
     sTerrain(T).BltFast X - 5, 300 - Terrain(CurX) - 5, sPlants, rPlants, DDBLTFAST_SRCCOLORKEY
   End If
   If CurX > LandZone(a) + 50 And a < UBound(LandZone) Then
     a = a + 1
   End If
   
   CurX = CurX + 10 + Rnd * 30
  Loop
  
  CurX = Rnd * 100
  
  ReDim Cloud(0)
  
  Do Until CurX >= MapSize
   ReDim Preserve Cloud(UBound(Cloud) + 1)
   Cloud(UBound(Cloud)).X = CurX
   Cloud(UBound(Cloud)).Y = Rnd * 10
   
   CurX = CurX + Rnd * 200
  Loop
  
  ReDim Radar(MapSize)
  Radar(MapSize) = Terrain(MapSize)
  Tall = Terrain(MapSize)
  For a = 1 To MapSize
   b = MapSize - a
   
   If Terrain(b) + 30 > Tall Then
    Tall = Terrain(b) + 30
   Else
    Tall = Tall - 0.7
   End If
   Radar(b) = Tall
  Next a
  
  ReDim RadarScan(0)
  
  Dim SoldierCount As Byte, TankCount As Byte, SAMCount As Byte, BFRocket As Boolean
  'ReDim Enemy(80)
  Dim Lowest As Integer
  
  Dim EnemySize(1 To 4) As Byte
  Dim TooClose As Boolean
  Dim HighestTerrainLeft As Integer, HighestTerrainRight As Integer
  
  EnemySize(cTank) = 10
  EnemySize(cSAM) = 20
  EnemySize(cSoldier) = 5
  c = 0
  
  For a = 1 To LevelCount
   Select Case a
    Case 1 ' Level 1
     SoldierCount = 40
     TankCount = 5
     SAMCount = 0
     BFRocket = False
    Case 2 ' Level 2
     SoldierCount = 30
     TankCount = 10
     SAMCount = 0 '
     BFRocket = True
    Case 3 ' Level 3
     SoldierCount = 30
     TankCount = 15
     SAMCount = 1 '
     BFRocket = True
    Case 4 ' Level 4
     SoldierCount = 30
     TankCount = 15
     SAMCount = 3 '
     BFRocket = False
    Case 5 ' Level 5
     SoldierCount = 20
     TankCount = 20
     SAMCount = 3 '
     BFRocket = True
   
   End Select
    
   If BFRocket = True Then
    ReDim Preserve RadarScan(UBound(RadarScan) + 1)
    RadarScan(UBound(RadarScan)) = (1 + (a - 1) * 3 + Rnd * 2) * 640
   End If
   
   EnemiesInLevel(a) = EnemiesInLevel(a - 1) + SoldierCount + TankCount + SAMCount
   c = EnemiesInLevel(a - 1)
   ReDim Preserve Enemy(EnemiesInLevel(a))
   
   For b = 1 To SAMCount
    c = c + 1
    Enemy(c).What = cSAM
    Enemy(c).HitPoints = 12
    Enemy(c).TimeToNextFire = Rnd * 10
    AliveEnemy.Add c
   Next b
   
   For b = 1 To TankCount
    c = c + 1
    Enemy(c).What = cTank
    Enemy(c).HitPoints = 8
    Enemy(c).TimeToNextFire = Rnd * 7
    AliveEnemy.Add c
   Next b
   
   For b = 1 To SoldierCount
    c = c + 1
    Enemy(c).What = cSoldier
    Enemy(c).HitPoints = 2
    Enemy(c).TimeToNextFire = Rnd * 7
    AliveEnemy.Add c
   Next b
             
   For b = EnemiesInLevel(a - 1) + 1 To EnemiesInLevel(a)
    TooClose = True
     
    Do Until TooClose = False
     Enemy(b).X = (1 + (a - 1) * 3 + Rnd * 3) * 640
     TooClose = False
     For c = EnemiesInLevel(a - 1) + 1 To b - 1
      If Abs(Enemy(c).X - Enemy(b).X) < EnemySize(Enemy(b).What) + EnemySize(Enemy(c).What) And c < b Then
       TooClose = True
       Exit For
      End If
     Next c
    
     If Enemy(b).What = cSAM And TooClose = False Then
      Lowest = 300
      Enemy(b).Y = 0
      For c = Enemy(b).X - 13 To Enemy(b).X + 13
       If Terrain(c) > Enemy(b).Y Then Enemy(b).Y = Terrain(c)
       If Terrain(c) < Lowest Then Lowest = Terrain(c)
      Next c
      If Abs(Lowest - Terrain(Enemy(b).X)) > 20 Then TooClose = True
      
     ElseIf Enemy(b).What = cTank And TooClose = False Then
      HighestTerrainLeft = 0
      HighestTerrainRight = 0
      
       For c = 1 To 15
        
        If Terrain(Enemy(b).X + c) > Terrain(HighestTerrainRight) Then
         HighestTerrainRight = Enemy(b).X + c
        End If
        If Terrain(Enemy(b).X - c) > Terrain(HighestTerrainLeft) Then
         HighestTerrainLeft = Enemy(b).X - c
        End If
       
       Next c
       Enemy(b).Angle = GetAngle(0, -Terrain(HighestTerrainLeft), HighestTerrainRight - HighestTerrainLeft, -Terrain(HighestTerrainRight))
         
       If Terrain(Enemy(b).X) < Terrain(HighestTerrainLeft) And Terrain(Enemy(b).X) < Terrain(HighestTerrainRight) Then
        Enemy(b).Y = (Terrain(HighestTerrainLeft) + Terrain(HighestTerrainRight)) / 2
       Else
        Enemy(b).Y = Terrain(Enemy(b).X)
       End If
     End If
    
    Loop
    
   Next b
  Next a
  
  EnemiesInLevel(LevelCount + 1) = EnemiesInLevel(LevelCount)
  
  ReDim Preserve RadarScan(UBound(RadarScan) + 1)
  RadarScan(UBound(RadarScan)) = MapSize
End Sub

Public Function GetTerrainNum(ByVal Pos As Long) As Byte
  GetTerrainNum = 1 + Int(Pos / 640)
  If GetTerrainNum > Size Then GetTerrainNum = Size
End Function
