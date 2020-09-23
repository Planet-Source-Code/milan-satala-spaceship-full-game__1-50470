Attribute VB_Name = "modMain"
Option Explicit

Public MapSize As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long

Private Const KEY_TOGGLED As Integer = &H1
Private Const KEY_DOWN As Integer = &H1000

Dim LastState(0 To 255) As Byte

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type tXY
  X As Long
  Y As Long
End Type

Public OldTime As Long, DrawTime As Long, GameSpeed As Single

Public EndGame As Boolean, GameTime As Single

Dim OldTick As Long, OldFPS As Integer, CurrentFPS As Integer

Public Const cMyShell As Byte = 1
Public Const cExplosion As Byte = 2
Public Const cTankShell As Byte = 3
Public Const cSoldierShell As Byte = 4
Public Const cFlesh As Byte = 5
Public Const cRocket As Byte = 6
Public Const cBFRocket As Byte = 7

Public Type tEffect
  Effect As Byte
  X As Single
  Y As Single
  Angle As Single
  MoveX As Single
  MoveY As Single
  Size As Single
  Progress As Single
End Type

Public UsedEffect As New Collection
Public UnusedEffect As New Collection

Public Effect() As tEffect

Public Const cSoldier As Byte = 1
Public Const cTank As Byte = 2
Public Const cSAM As Byte = 3

Public AliveEnemy As New Collection

Public Type tEnemy
  What As Byte
  X As Single
  Y As Single
  Angle As Single
  WeaponAngle As Integer
  WeaponX As Integer
  WeaponY As Integer
  TimeToNextFire As Single
  HitPoints As Integer
End Type

Public Enemy() As tEnemy

Public EndIn As Single, Paused As Boolean

Public Const RadToDeg As Single = 180 / 3.14159
Public Const DegToRad As Single = 3.14159 / 180

Public Function GS(ByVal Var As Single) As Single
  GS = Var * GameSpeed
End Function

Public Function KeyState(ByVal m_Key As Byte) As Byte
 KeyState = 0
 If (GetKeyState(m_Key) And KEY_DOWN) Then KeyState = 1
 If LastState(m_Key) > 0 And KeyState = 1 Then KeyState = 2
 If LastState(m_Key) = 3 Then LastState(m_Key) = 0
 If LastState(m_Key) > 0 And KeyState = 0 Then KeyState = 3
 LastState(m_Key) = KeyState
End Function

Public Function FixAngle(Angle As Single) As Single
  FixAngle = Angle
  If Angle > 359 Then FixAngle = Angle - 359
  If Angle < 0 Then
  FixAngle = 359 + Angle
  End If
End Function

Function GetAngle(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single
  Dim Cislo1 As Single
  Dim Cislo2 As Single
  Dim Uhol As Single
  Dim Poloha As Integer
  
  If X1 = X2 And Y1 < Y2 Then
   Cislo2 = 0
   Poloha = 180
  
  
  ElseIf X1 = X2 And Y1 > Y2 Then
   Cislo2 = 0
   Poloha = 0
  ElseIf X1 < X2 And Y1 = Y2 Then
   Cislo2 = 0
   Poloha = 90
  ElseIf X1 > X2 And Y1 = Y2 Then
   Cislo2 = 0
   Poloha = 270
  ElseIf X1 < X2 And Y1 > Y2 Then
   Cislo1 = Abs(X2 - X1)
   Cislo2 = Abs(Y2 - Y1)
   Poloha = 0
  ElseIf X1 < X2 And Y1 < Y2 Then
   Cislo1 = Abs(Y1 - Y2)
   Cislo2 = Abs(X2 - X1)
   Poloha = 90
  ElseIf X1 > X2 And Y1 < Y2 Then
   Cislo1 = Abs(X1 - X2)
   Cislo2 = Abs(Y1 - Y2)
   Poloha = 180
  ElseIf X1 > X2 And Y1 > Y2 Then
   Cislo1 = Abs(Y2 - Y1)
   Cislo2 = Abs(X1 - X2)
   Poloha = 270
  End If
  
On Error GoTo Chyba
  Uhol = Atn(Cislo1 / Cislo2) * RadToDeg
Chyba:

  GetAngle = Uhol + Poloha
End Function

Function GetDist(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single
  GetDist = Sqr((X1 - X2) * (X1 - X2) + (Y1 - Y2) * (Y1 - Y2))
End Function

Function SubtractAngles(ByVal Angle1 As Integer, ByVal Angle2 As Integer) As Integer
  Dim Plus As Integer
  Dim Minus As Integer
  Dim PlusKurz As Integer
  Dim MinusKurz As Integer
  
  PlusKurz = Angle1
  MinusKurz = Angle1
  
  If Angle2 < Angle1 Then Plus = 360 - Angle1: PlusKurz = 0
  If Angle2 > Angle1 Then Minus = -Angle1: MinusKurz = 359
  
  Plus = Plus + (Angle2 - PlusKurz)
  Minus = Minus + (Angle2 - MinusKurz)
  
  If Plus > Abs(Minus) Then SubtractAngles = Minus Else SubtractAngles = Plus
  
End Function

Public Function GetFPS() As Integer
  Dim Rozdiel As Long
  
  Rozdiel = GetTickCount - OldTick
  CurrentFPS = CurrentFPS + 1
  
  If Rozdiel >= 1000 Then
   OldFPS = CurrentFPS
   OldTick = GetTickCount
   CurrentFPS = 0
  End If
  GetFPS = OldFPS
End Function

Public Function CreateEffect() As Integer
  If UnusedEffect.Count = 0 Then
   ReDim Preserve Effect(UBound(Effect) + 1)
   CreateEffect = UBound(Effect)
  Else
   CreateEffect = UnusedEffect(1)
   UnusedEffect.Remove 1
   Dim Default As tEffect
   Effect(CreateEffect) = Default
  End If
        
  UsedEffect.Add CreateEffect
End Function
