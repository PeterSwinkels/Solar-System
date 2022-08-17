Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Base 0
Option Compare Binary
Option Explicit

Private Const PI As Double = 3.14159265358979           'Defines the value of PI.
Private Const DEGREES_PER_RADIAN As Double = 180 / PI   'Defines the number of degrees per radian.
   
'This structure defines a planet.
Private Type PlanetStr
   Name As String       'Defines a planet's name.
   Radius As Single     'Defines a planet's radius.
   Color As Long        'Defines a planet's color.
   Distance As Long     'Defines a planet's distance from the sun.
   Position As Double   'Defines a planet's position in its orbit.
   Velocity As Long     'Defines a planet's velocity.
End Type

'This structure defines a star.
Private Type StarStr
   x As Long              'Defines a star's horizontal position.
   y As Long              'Defines a star's vertical position.
   Twinkling As Boolean   'Defines whether a star is twinkling.
End Type

'This structure defines the sun.
Private Type SunStr
   Name As String         'Defines the sun's name.
   Diameter As Single     'Defines the sun's diameter.
   Color As Long          'Defines the sun's color.
End Type

'This structure defines the solar system.
Public Type SolarSystemStr
   Planets() As PlanetStr   'Defines the solar system's planets.
   Sun As SunStr            'Defines the solar system's sun.
   x As Long                'Defines the solar system's horizontal position.
   y As Long                'Defines the solar system's vertical position.
   ScaleV As Long           'Defines the scale used to draw the solar system.
   Tilt As Long             'Defines the solar system's tilt.
End Type

Public SolarSystem As SolarSystemStr   'Contains the solar system.
Private Stars() As StarStr   'Contains the stars.

'This procedure displays the specified text at the specified position using the specified color on the specified window.
Private Sub DisplayText(Window As Form, Text As String, x As Long, y As Long, TextColor As Long, EraseV As Boolean)
On Error GoTo ErrorTrap
   With Window
      .CurrentX = x
      .CurrentY = y
      If EraseV Then
         Window.Line -Step(.TextWidth(Text), .TextHeight(Text)), vbBlack, BF
      Else
         .ForeColor = TextColor
         Window.Print Text;
      End If
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure draws the specified planets' orbits on the specified window.
Private Sub DrawOrbits(Window As Form, SolarSystem As SolarSystemStr, OrbitStart As Long, OrbitEnd As Long, OrbitColor As Long)
On Error GoTo ErrorTrap
Dim Degree As Long
Dim Planet As Long
Dim x As Long
Dim y As Long

   With SolarSystem
      For Planet = LBound(.Planets()) To UBound(.Planets())
         With .Planets(Planet)
            For Degree = OrbitStart To OrbitEnd
               x = (Cos(Degree / DEGREES_PER_RADIAN) * (.Distance * SolarSystem.ScaleV)) + SolarSystem.x
               y = (Sin(Degree / DEGREES_PER_RADIAN) * ((.Distance * SolarSystem.ScaleV) * (SolarSystem.Tilt / 10))) + SolarSystem.y
               If Degree = OrbitStart Then
                  Window.CurrentX = x
                  Window.CurrentY = y
               Else
                  Window.Line -(x, y), OrbitColor
               End If
            Next Degree
         End With
      Next Planet
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure draws the specified planets on the specified window.
Private Sub DrawPlanets(Window As Form, SolarSystem As SolarSystemStr, OrbitStart As Long, OrbitEnd As Long, Direction As Long, EraseV As Boolean)
On Error GoTo ErrorTrap
Dim First As Long
Dim Last As Long
Dim Planet As Long

   With SolarSystem
      Select Case Direction
         Case -1
            First = UBound(.Planets())
            Last = LBound(.Planets())
         Case 1
            First = LBound(.Planets())
            Last = UBound(.Planets())
      End Select
      
      For Planet = First To Last Step Direction
         If .Planets(Planet).Position >= OrbitStart And .Planets(Planet).Position <= OrbitEnd Then
            DrawPlanet Window, .Planets(Planet), EraseV:=EraseV
         End If
      Next Planet
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure draws the specified planet on the specified window.
Private Sub DrawPlanet(Window As Form, Planet As PlanetStr, EraseV As Boolean)
On Error GoTo ErrorTrap
Dim x As Long
Dim y As Long

   With Planet
      x = (Sin((.Position - 90) / DEGREES_PER_RADIAN) * (.Distance * SolarSystem.ScaleV)) + SolarSystem.x
      y = (Cos((.Position - 90) / DEGREES_PER_RADIAN) * ((.Distance * SolarSystem.ScaleV) * (SolarSystem.Tilt / 10))) + SolarSystem.y
      If EraseV Then
         Window.FillColor = vbBlack
      Else
         Window.FillColor = .Color
      End If
      Window.Circle (x, y), .Radius * (SolarSystem.ScaleV / 10), vbBlack
      DisplayText Window, .Name, x, y, TextColor:=&HFFFFFF, EraseV:=EraseV
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure draws the solar system on the specified window.
Public Sub DrawSolarSystem(Window As Form, MoveV As Boolean)
On Error GoTo ErrorTrap
   DrawStars Window, Stars, EraseV:=True
   DrawPlanets Window, SolarSystem, OrbitStart:=0, OrbitEnd:=180, Direction:=1, EraseV:=True
   DrawPlanets Window, SolarSystem, OrbitStart:=180, OrbitEnd:=360, Direction:=-1, EraseV:=True
   DrawStars Window, Stars, EraseV:=False

   If MoveV Then MovePlanets SolarSystem
   
   DrawOrbits Window, SolarSystem, OrbitStart:=0, OrbitEnd:=180, OrbitColor:=&H808080
   DrawPlanets Window, SolarSystem, OrbitStart:=0, OrbitEnd:=180, Direction:=1, EraseV:=False

   DrawSun Window, SolarSystem, DiameterScale:=0.001

   DrawOrbits Window, SolarSystem, OrbitStart:=180, OrbitEnd:=360, OrbitColor:=&H808080
   DrawPlanets Window, SolarSystem, OrbitStart:=180, OrbitEnd:=360, Direction:=-1, EraseV:=False
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub



'This procedure draws the specified stars on the specified window.
Private Sub DrawStars(Window As Form, Stars() As StarStr, EraseV As Boolean)
On Error GoTo ErrorTrap
Dim ColorV As Long
Dim Star As Long
   
   TwinkleStars Stars
      
   If EraseV Then
      ColorV = vbBlack
   Else
      ColorV = vbWhite
   End If
      
   Window.FillColor = ColorV
   For Star = LBound(Stars()) To UBound(Stars())
      If Stars(Star).Twinkling Then
         Window.Circle (Stars(Star).x, Stars(Star).y), 1, ColorV
      Else
         Window.PSet (Stars(Star).x, Stars(Star).y), ColorV
         Window.PSet (Stars(Star).x + 1, Stars(Star).y), ColorV
      End If
   Next Star
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure draws the specified sun on the specified window.
Private Sub DrawSun(Window As Form, SolarSystem As SolarSystemStr, DiameterScale As Single)
On Error GoTo ErrorTrap
   With SolarSystem
      Window.FillColor = .Sun.Color
      Window.Circle (.x, .y), .Sun.Diameter * DiameterScale, vbBlack
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure generates the stars using the specified window's dimensions.
Private Function GenerateStars(Window As Form) As StarStr()
On Error GoTo ErrorTrap
Dim NewStars() As StarStr
Dim Star As Long

   ReDim NewStars(0 To Int(Rnd * 100) + 99) As StarStr
   
   For Star = LBound(NewStars()) To UBound(NewStars())
      NewStars(Star).x = Int(Rnd() * Window.ScaleWidth)
      NewStars(Star).y = Int(Rnd() * Window.ScaleHeight)
      NewStars(Star).Twinkling = (Rnd() > 0.5)
   Next Star
   
EndRoutine:
   GenerateStars = NewStars()
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function
'This procedure handles any errors that occur.
Public Sub HandleError()
Dim Choice As Integer
Dim ErrorCode As Long
Dim Message As String

   ErrorCode = Err.Number
   Message = Err.Description
   
   On Error Resume Next
   Message = Message & vbCr & "Error code: " & CStr(ErrorCode)
   Choice = MsgBox(Message, vbExclamation Or vbOKCancel)
   If Choice = vbCancel Then End
End Sub




'This procedure moves the specified planets.
Private Sub MovePlanets(SolarSystem As SolarSystemStr)
On Error GoTo ErrorTrap
Dim Planet As Long

   With SolarSystem
      For Planet = LBound(.Planets()) To UBound(.Planets())
         .Planets(Planet).Position = NewPosition(.Planets(Planet).Position, .Planets(Planet).Velocity)
      Next Planet
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure increments or resets the specified position and returns the result.
Private Function NewPosition(Position As Double, Velocity As Long) As Long
On Error GoTo ErrorTrap
   If Position + Velocity >= 360 Then Position = 0 Else Position = Position + Velocity
EndRoutine:
   NewPosition = Position
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure loads and returns the solar system's data.
Private Function LoadSolarSystem(DataFile As String) As SolarSystemStr
On Error GoTo ErrorTrap
Dim FileH As Integer
Dim HexadecimalColor As String
Dim NewSolarSystem As SolarSystemStr
Dim Planet As Long
Dim PlanetCount As Long

   FileH = FreeFile()
   Open DataFile For Input Lock Read Write As FileH
      With NewSolarSystem
         Input #FileH, .Sun.Name, HexadecimalColor, .Sun.Diameter
         .Sun.Color = CLng(Val("&H" & HexadecimalColor & "&"))
         
         Input #FileH, PlanetCount
         ReDim .Planets(0 To PlanetCount - 1) As PlanetStr
         For Planet = LBound(.Planets()) To UBound(.Planets())
            With .Planets(Planet)
               Input #FileH, .Name, HexadecimalColor, .Radius, .Distance, .Velocity, .Position
               .Color = CLng(Val("&H" & HexadecimalColor & "&"))
            End With
         Next Planet
      End With
   Close FileH

EndRoutine:
   LoadSolarSystem = NewSolarSystem
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure is executed when this program is started.
Public Sub Main()
On Error GoTo ErrorTrap
   Randomize
    
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   SolarSystem = LoadSolarSystem("Solar System.txt")
   SolarSystem.ScaleV = 1
   SolarSystem.Tilt = 2
   
   Stars = GenerateStars(InterfaceWindow)
   
   InterfaceWindow.Show
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure controls the specified stars' twinkling.
Private Sub TwinkleStars(Stars() As StarStr)
On Error GoTo ErrorTrap
Dim SelectedStar As Long
Dim Star As Long

   For Star = 0 To Int(Rnd * Abs(UBound(Stars()) - LBound(Stars()))) \ 2
      SelectedStar = Int(Rnd * Abs(UBound(Stars()) - LBound(Stars()))) + LBound(Stars())
      Stars(SelectedStar).Twinkling = (Rnd() > 0.5)
   Next Star
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub



