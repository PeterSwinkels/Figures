VERSION 5.00
Begin VB.Form FiguresWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Animator 
      Enabled         =   0   'False
      Interval        =   125
      Left            =   240
      Top             =   240
   End
End
Attribute VB_Name = "FiguresWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's main interface window.
Option Explicit

'The Microsoft Windows API functions and subroutines used by this program.
Private Declare Function SafeArrayGetDim Lib "Oleaut32.dll" (ByRef saArray() As Long) As Long
Private Declare Sub RtlMoveMemory Lib "Kernel32.dll" (Destination As Long, Source As Long, ByVal Length As Long)

'The constants used by this program:
Private Const PI As Double = 3.14159265358979           'Defines the value of PI.
Private Const DEGREES_PER_RADIAN As Double = 180 / PI   'Defines the number of degrees per radian.
Private Const NO_COLOR As Long = -1                     'Indicates that no color is to be used.

'This procedure draws the specified polygon at the specified angle and position.
Private Sub DrawFigure(x As Long, y As Long, Radii() As Long, Optional Angle As Double = 0, Optional EdgeColor As Long = vbBlack, Optional LineWidth As Long = 1, Optional DrawRadii As Boolean = False)
Dim Degrees As Double
Dim Increment As Double
Dim NextRadianTipX As Long
Dim NextRadianTipY As Long
Dim Radian As Long
Dim RadianTipX As Long
Dim RadianTipY As Long

Degrees = Angle
Increment = 360 / (Abs(UBound(Radii()) - LBound(Radii())) + 1)
Me.DrawWidth = LineWidth
For Radian = LBound(Radii()) To UBound(Radii())
   RadianTipX = (Cos(Degrees / DEGREES_PER_RADIAN) * Radii(Radian)) + x
   RadianTipY = (Sin(Degrees / DEGREES_PER_RADIAN) * Radii(Radian)) + y
   If Radian = UBound(Radii()) Then
      NextRadianTipX = (Cos((Degrees + Increment) / DEGREES_PER_RADIAN) * Radii(LBound(Radii()))) + x
      NextRadianTipY = (Sin((Degrees + Increment) / DEGREES_PER_RADIAN) * Radii(LBound(Radii()))) + y
   Else
      NextRadianTipX = (Cos((Degrees + Increment) / DEGREES_PER_RADIAN) * Radii(Radian + 1)) + x
      NextRadianTipY = (Sin((Degrees + Increment) / DEGREES_PER_RADIAN) * Radii(Radian + 1)) + y
   End If

   If DrawRadii Then
      Me.DrawWidth = 1
      Me.Line (x, y)-(NextRadianTipX, NextRadianTipY), vbWhite
      Me.DrawWidth = LineWidth
   End If

   Me.Line (RadianTipX, RadianTipY)-(NextRadianTipX, NextRadianTipY), EdgeColor
   Degrees = Degrees + Increment
Next Radian
End Sub

'This procedure gives the command to generate and draw several different polygons.
Private Sub DrawFigures(Optional Angle As Double = 0, Optional NewDrawRadii As Variant)
Static CurrentDrawRadii As Boolean

If Not IsMissing(NewDrawRadii) Then
   CurrentDrawRadii = CBool(NewDrawRadii)
   If Animator.Enabled Then Exit Sub
End If

Me.Cls
DrawFigure 110, 110, GenerateRadii(SeedRadii:="100", Repeat:=360), Angle, vbGreen, LineWidth:=2, DrawRadii:=CurrentDrawRadii
DrawFigure 320, 110, GenerateRadii(SeedRadii:="100", Repeat:=8), Angle, vbGreen, LineWidth:=2, DrawRadii:=CurrentDrawRadii
DrawFigure 530, 110, GenerateRadii(SeedRadii:="100", Repeat:=6), Angle, vbGreen, LineWidth:=2, DrawRadii:=CurrentDrawRadii
DrawFigure 740, 110, GenerateRadii(SeedRadii:="100", Repeat:=4), Angle, vbGreen, LineWidth:=2, DrawRadii:=CurrentDrawRadii
DrawFigure 950, 110, GenerateRadii(SeedRadii:="100", Repeat:=3), Angle, vbGreen, LineWidth:=2, DrawRadii:=CurrentDrawRadii
DrawFigure 110, 320, GenerateRadii(SeedRadii:="100", Repeat:=2), Angle, vbGreen, LineWidth:=2, DrawRadii:=CurrentDrawRadii
DrawFigure 320, 320, GenerateRadii(SeedRadii:="50,100", Repeat:=5), Angle, vbRed, LineWidth:=2, DrawRadii:=CurrentDrawRadii
DrawFigure 530, 320, GenerateRadii(SeedRadii:="100,100,80,80", Repeat:=15), Angle, vbRed, LineWidth:=2, DrawRadii:=CurrentDrawRadii
DrawFigure 850, 430, ExtendFigure(GenerateRadii(SeedRadii:="100", Repeat:=3), GenerateRadii(SeedRadii:="200", Repeat:=1)), Angle, vbBlue, 3, DrawRadii:=CurrentDrawRadii
End Sub


'This procedure extends one polygon by extending it with another polygon and returns the result.
Private Function ExtendFigure(Figure() As Long, Extension() As Long) As Long()
Dim Extended() As Long
Dim ExtensionSize As Long
Dim UnextendedSize As Long

Extended() = Figure()

ExtensionSize = (UBound(Extension()) - LBound(Extension())) + 1
UnextendedSize = UBound(Extended()) - LBound(Extended())

ReDim Preserve Extended(LBound(Extended()) To UBound(Extended()) + ExtensionSize) As Long

RtlMoveMemory Extended(UnextendedSize + 1), Extension(LBound(Extension())), (ExtensionSize * 4)

ExtendFigure = Extended()
End Function

'This procedure generates a set of radii using the specified seed radii.
Private Function GenerateRadii(Optional SeedRadii As String = Empty, Optional Repeat As Long = 1) As Long()
Dim CommaPosition As Long
Dim NewRadian As Long
Dim Radii() As Long
Dim RemainingRadii As String
Dim Repetition As Long

If Not SeedRadii = Empty Then
   For Repetition = 1 To Repeat
      RemainingRadii = SeedRadii
      If Not Right$(Trim$(RemainingRadii), 1) = "," Then RemainingRadii = RemainingRadii & ","
      Do Until RemainingRadii = Empty
         CommaPosition = InStr(LTrim$(RemainingRadii), ",")
   
         NewRadian = CLng(Val(LTrim$(Left$(LTrim$(RemainingRadii), CommaPosition - 1))))
         If SafeArrayGetDim(Radii()) = 0 Then
            ReDim Radii(0 To 0) As Long
         Else
            ReDim Preserve Radii(LBound(Radii()) To UBound(Radii()) + 1) As Long
         End If
         Radii(UBound(Radii())) = NewRadian
   
         RemainingRadii = LTrim$(Mid$(LTrim$(RemainingRadii), CommaPosition + 1))
      Loop
   Next Repetition
End If

GenerateRadii = Radii()
End Function


'This procedure animates the various figures to be displayed by changing their angle.
Private Sub Animator_Timer()
Static Angle As Double

DrawFigures Angle

If Angle >= 360 Then Angle = 0 Else Angle = Angle + 8
End Sub

'This procedure generates and displays several figures.
Private Sub Form_Activate()
DrawFigures Angle:=0
End Sub
'This procedure handles the user's double clicks.
Private Sub Form_DblClick()
Animator.Enabled = Not Animator.Enabled
End Sub

'This procedure handles the user's key strokes.
Private Sub Form_KeyPress(KeyAscii As Integer)
Static CurrentDrawRadii As Boolean

CurrentDrawRadii = Not CurrentDrawRadii

DrawFigures , CurrentDrawRadii
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
ChDrive Left$(App.Path, InStr(App.Path, ":"))
ChDir App.Path

Me.Width = Screen.Width / 1.1
Me.Height = Screen.Height / 1.1

With App
   Me.Caption = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName
End With
End Sub


'This procedure closes this program when this window is closed.
Private Sub Form_Unload(Cancel As Integer)
End
End Sub


