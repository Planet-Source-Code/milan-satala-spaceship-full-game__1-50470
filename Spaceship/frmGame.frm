VERSION 5.00
Begin VB.Form frmGame 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spaceship by Milan Satala"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   455
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   634
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStatus 
      BackColor       =   &H00FF8080&
      Caption         =   "Staus"
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   3240
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton cmdEnd 
         Caption         =   "End"
         Height          =   495
         Left            =   1560
         TabIndex        =   3
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New game"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF8080&
         Caption         =   "Enemies destroyed:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblEnemies 
         BackColor       =   &H00FF8080&
         Caption         =   "200"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblLevel 
         BackColor       =   &H00FF8080&
         Caption         =   "10"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF8080&
         Caption         =   "Levels completed:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00FF8080&
         Caption         =   "5 minutes"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         Caption         =   "Time:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Success !!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2820
      End
   End
   Begin VB.Label lblFontRadar 
      Caption         =   "lblFontRadar"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   12
      Top             =   5520
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblFont2 
      Caption         =   "lblFont2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblFont 
      Caption         =   "lblFont"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   10
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblLoad 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Loading ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   9255
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnd_Click()
  End
End Sub

Private Sub cmdNew_Click()
  fraStatus.Visible = False
  Cls
  lblLoad.Visible = True
  StartGame
End Sub

Private Sub Form_Load()

  ColorOfSky = RGB(0, 255, 255)
  ColorOfGround = RGB(160, 80, 0)
  DarkRed = RGB(200, 0, 0)
  PurpleColor = RGB(255, 0, 255)
  ColorOfLandzone = RGB(180, 180, 180)
  
  Show
  DoEvents
  Randomize Timer
  
  Dim TMPDesc As DDSURFACEDESC2
  Dim DDColor As DDCOLORKEY
  Dim PrimaryDesc As DDSURFACEDESC2
  Dim BackDesc As DDSURFACEDESC2
  
  Set DirX = New DirectX7
  
  Set DDraw = DirX.DirectDrawCreate("")
  DDraw.SetCooperativeLevel frmGame.hWnd, DDSCL_NORMAL
  
  DirX.GetWindowRect frmGame.hWnd, PrimaryRect
  
  DDColor.low = 0
  DDColor.high = 0
  
  PrimaryDesc.lFlags = DDSD_CAPS
  PrimaryDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
  PrimaryRect.Top = PrimaryRect.Top + 22
  PrimaryRect.Left = PrimaryRect.Left + 4
  PrimaryRect.Right = PrimaryRect.Right - 4
  PrimaryRect.Bottom = PrimaryRect.Bottom - 4
  PrimaryDesc.lWidth = PrimaryRect.Right - PrimaryRect.Left
  PrimaryDesc.lHeight = PrimaryRect.Bottom - PrimaryRect.Top
    
  Set sPrimary = DDraw.CreateSurface(PrimaryDesc)
    
  BackDesc.lFlags = 7
  BackDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
  BackDesc.lWidth = 640
  BackDesc.lHeight = 480
  
  BackRect.Right = 640
  BackRect.Bottom = 480
  
  Set sBack = DDraw.CreateSurface(BackDesc)
  Set sShip = DDraw.CreateSurfaceFromFile(App.Path & "\Data\Spaceship.bmp", ShipDesc)
  sShip.SetColorKey DDCKEY_SRCBLT, DDColor
  ShipRect.Right = ShipDesc.lWidth
  ShipRect.Bottom = ShipDesc.lHeight
  
  TMPDesc = DefaultDesc
  Set sTerrainDirt = DDraw.CreateSurfaceFromFile(App.Path & "\Data\TerrainDirt.bmp", TMPDesc)
  sTerrainDirt.SetColorKey DDCKEY_SRCBLT, DDColor
  
  TMPDesc = DefaultDesc
  Set sPlants = DDraw.CreateSurfaceFromFile(App.Path & "\Data\Plants.bmp", TMPDesc)
  sPlants.SetColorKey DDCKEY_SRCBLT, DDColor
  
  TMPDesc = DefaultDesc
  Set sSoldier = DDraw.CreateSurfaceFromFile(App.Path & "\Data\Soldier.bmp", TMPDesc)
  sSoldier.SetColorKey DDCKEY_SRCBLT, DDColor
  
  TMPDesc = DefaultDesc
  Set sTank = DDraw.CreateSurfaceFromFile(App.Path & "\Data\Tank.bmp", TMPDesc)
  sTank.SetColorKey DDCKEY_SRCBLT, DDColor
  
  TMPDesc = DefaultDesc
  Set sSAM = DDraw.CreateSurfaceFromFile(App.Path & "\Data\SAM.bmp", TMPDesc)
  sSAM.SetColorKey DDCKEY_SRCBLT, DDColor
  rSAM.Right = TMPDesc.lWidth
  rSAM.Bottom = TMPDesc.lHeight
  
  TMPDesc = DefaultDesc
  Set sBFRocket = DDraw.CreateSurfaceFromFile(App.Path & "\Data\BigFuckingRocket.bmp", TMPDesc)
  sBFRocket.SetColorKey DDCKEY_SRCBLT, DDColor
  
  TMPDesc = DefaultDesc
  Set sRocket = DDraw.CreateSurfaceFromFile(App.Path & "\Data\Rocket.bmp", TMPDesc)
  sRocket.SetColorKey DDCKEY_SRCBLT, DDColor
  
  TMPDesc = DefaultDesc
  Set sCloud = DDraw.CreateSurfaceFromFile(App.Path & "\Data\Cloud.bmp", TMPDesc)
  sCloud.SetColorKey DDCKEY_SRCBLT, DDColor
  rCloud.Right = TMPDesc.lWidth
  rCloud.Bottom = TMPDesc.lHeight
  
  GetShip
  
  StartGame
End Sub

Sub StartGame()
  Dim TMPDesc As DDSURFACEDESC2
  Dim DDColor As DDCOLORKEY

  lblLoad = "Generating Terrain ..."
  DoEvents
  
  TMPDesc = DefaultDesc
  TMPDesc.lFlags = 7
  TMPDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
  
  Size = LevelCount * 3 + 2
  MapSize = Size * 640
  TMPDesc.lWidth = 640
  ReDim Terrain(MapSize)
  ReDim sTerrain(Size)
  TMPDesc.lHeight = 300
  TerrainRect.Right = 640
  TerrainRect.Bottom = 300
  
  For a = 1 To Size
   Set sTerrain(a) = DDraw.CreateSurface(TMPDesc)
   sTerrain(a).setDrawWidth 1
   sTerrain(a).BltColorFill TerrainRect, 0
   sTerrain(a).SetColorKey DDCKEY_SRCBLT, DDColor
  Next a
  
  SetDefault
  GenerateTerrain
  
  lblLoad.Visible = False
  OldTime = GetTickCount
  
  Do Until EndGame
   GameTick
  Loop
  EndGame = False
End Sub

Sub SetDefault()
  Spaceship.Angle = 0
  Spaceship.MoveX = 0
  Spaceship.MoveY = 0
  Spaceship.Thrust = 0
  Spaceship.Y = 350
  Spaceship.X = 350
  Spaceship.Fuel = 100
  Spaceship.Shield = 100
  GameTime = 0
  CurrentLandZone = 1
  CurrentRadarScan = 1
  gCamX = Spaceship.X - 300
  gMoveX = 0
  ShipCanLand = False
  ReDim Effect(0)
  Set UnusedEffect = New Collection
  Set UsedEffect = New Collection
  Set AliveEnemy = New Collection
  EnemiesDestroyed = 0
  EndIn = -1
  RadarScanTimeLeft = 0
End Sub

Private Sub GetShip()
  Dim a As Integer, b As Integer, c As Integer, d As Integer
  Dim IsCorner As Boolean, Leg As Byte
  Dim LegPos(1 To 2) As POINTAPI
  Dim SRect As RECT
  Dim Angle As Byte
  
  sShip.Lock ShipRect, ShipDesc, DDLOCK_WAIT, 0
  sShip.Unlock ShipRect
  
  With Spaceship

   For Angle = 0 To 35
     Y = Int(Angle / 9)
    
     SRect.Left = Angle * 35 - Y * 315
     SRect.Top = Y * 35
     SRect.Bottom = SRect.Top + 35
     SRect.Right = SRect.Left + 35
    
     For a = 0 To 34
       X = SRect.Left + a
       .UpCollision(Angle, a) = 255
       
       For b = 0 To 34
         Y = SRect.Top + b
         If sShip.GetLockedPixel(X, Y) <> 0 Then
          If .DownCollision(Angle, a) = 0 Then .UpCollision(Angle, a) = b
          .DownCollision(Angle, a) = b
         End If
       Next b
       
     Next a
     
   Next Angle
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub ScrSize_Change()
  lblSize = ScrSize.Value
End Sub
