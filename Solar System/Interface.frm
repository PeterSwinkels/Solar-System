VERSION 5.00
Begin VB.Form InterfaceWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Interface.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer MotionGenerator 
      Interval        =   55
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's main interface window.
Option Base 0
Option Compare Binary
Option Explicit


'This procedure initializes this window.
Private Sub Form_Activate()
On Error GoTo ErrorTrap
   With SolarSystem
      .x = Me.ScaleWidth / 2
      .y = Me.ScaleHeight / 2
   End With
     
   DrawSolarSystem Me, MoveV:=False
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure handles the user's key strokes.
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap
   With SolarSystem
      Select Case KeyCode
         Case vbKeyAdd
            If MotionGenerator.Interval > 55 Then MotionGenerator.Interval = MotionGenerator.Interval - 100
         Case vbKeyDown
            .y = .y - (1 * .ScaleV)
         Case vbKeyEnd
            If .ScaleV > 1 Then .ScaleV = .ScaleV - 1
         Case vbKeyF1
            MsgBox App.Comments, vbInformation
         Case vbKeyHome
            If .ScaleV < 10 Then .ScaleV = .ScaleV + 1
         Case vbKeyLeft
            .x = .x - (1 * .ScaleV)
         Case vbKeyPageDown
            If .Tilt < 10 Then .Tilt = .Tilt + 1
         Case vbKeyPageUp
            If .Tilt > 0 Then .Tilt = .Tilt - 1
         Case vbKeyPause
            MotionGenerator.Enabled = Not MotionGenerator.Enabled
         Case vbKeyRight
            .x = .x + (1 * .ScaleV)
         Case vbKeySubtract
            If MotionGenerator.Interval < 1055 Then MotionGenerator.Interval = MotionGenerator.Interval + 100
         Case vbKeyUp
            .y = .y + (1 * .ScaleV)
      End Select
   End With
      
EndRoutine:
   Me.Cls
   DrawSolarSystem Me, MoveV:=False
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   Me.Width = Screen.Width / 1.1
   Me.Height = Screen.Height / 1.1
   
   With App
      Me.Caption = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName & ", ***2022*** - See Help.txt for usage."
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure handles the user's mouse clicks.
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo ErrorTrap
   If Button = vbLeftButton Then
      SolarSystem.x = x
      SolarSystem.y = y
      Me.Cls
      DrawSolarSystem Me, MoveV:=False
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure closes this window.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorTrap
   Unload Me
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure gives the command to draw the solar system at specific intervals.
Private Sub MotionGenerator_Timer()
On Error GoTo ErrorTrap
   DrawSolarSystem Me, MoveV:=True
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


