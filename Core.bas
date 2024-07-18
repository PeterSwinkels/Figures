Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API functions and subroutines used by this program.
Private Declare Function SafeArrayGetDim Lib "Oleaut32.dll" (ByRef saArray() As Long) As Long
Private Declare Sub RtlMoveMemory Lib "Kernel32.dll" (Destination As Long, Source As Long, ByVal Length As Long)

'The constants used by this program:
Private Const PI As Double = 3.14159265358979           'Defines the value of PI.
Private Const DEGREES_PER_RADIAN As Double = 180 / PI   'Defines the number of degrees per radian.
Private Const NO_COLOR As Long = -1                     'Indicates that no color is to be used.

'This procedure draws the specified polygon at the specified angle and position.
Private Sub DrawFigure(x As Long, y As Long, Radii() As Long, Canvas As Object, Optional Angle As Double = 0, Optional EdgeColor As Long = vbBlack, Optional LineWidth As Long = 1, Optional DrawRadii As Boolean = False)
On Error GoTo ErrorTrap
Dim Degrees As Double
Dim Increment As Double
Dim NextRadianVertexX As Long
Dim NextRadianVertexY As Long
Dim Radian As Long
Dim RadianVertexX As Long
Dim RadianVertexY As Long

   If Not SafeArrayGetDim(Radii()) = 0 Then
      Degrees = Angle
      Increment = 360 / (Abs(UBound(Radii()) - LBound(Radii())) + 1)
      Canvas.DrawWidth = LineWidth
      For Radian = LBound(Radii()) To UBound(Radii())
         RadianVertexX = (Cos(Degrees / DEGREES_PER_RADIAN) * Radii(Radian)) + x
         RadianVertexY = (Sin(Degrees / DEGREES_PER_RADIAN) * Radii(Radian)) + y
         If Radian = UBound(Radii()) Then
            NextRadianVertexX = (Cos((Degrees + Increment) / DEGREES_PER_RADIAN) * Radii(LBound(Radii()))) + x
            NextRadianVertexY = (Sin((Degrees + Increment) / DEGREES_PER_RADIAN) * Radii(LBound(Radii()))) + y
         Else
            NextRadianVertexX = (Cos((Degrees + Increment) / DEGREES_PER_RADIAN) * Radii(Radian + 1)) + x
            NextRadianVertexY = (Sin((Degrees + Increment) / DEGREES_PER_RADIAN) * Radii(Radian + 1)) + y
         End If
      
         If DrawRadii Then
            Canvas.DrawWidth = 1
            Canvas.Line (x, y)-(NextRadianVertexX, NextRadianVertexY), vbWhite
            Canvas.DrawWidth = LineWidth
         End If
      
         Canvas.Line (RadianVertexX, RadianVertexY)-(NextRadianVertexX, NextRadianVertexY), EdgeColor
         Degrees = Degrees + Increment
      Next Radian
   End If

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Sub

'This procedure gives the command to generate and draw several different polygons.
Public Sub DrawFigures(Canvas As Object, AnimatorActive As Boolean, Optional Angle As Double = 0)
On Error GoTo ErrorTrap
   
   Canvas.Cls
   DrawFigure 110, 110, GenerateRadii(SeedRadii:="100", Count:=360), Canvas, Angle, vbGreen, LineWidth:=2, DrawRadii:=DrawRadii()
   DrawFigure 320, 110, GenerateRadii(SeedRadii:="100", Count:=8), Canvas, Angle, vbGreen, LineWidth:=2, DrawRadii:=DrawRadii()
   DrawFigure 530, 110, GenerateRadii(SeedRadii:="100", Count:=6), Canvas, Angle, vbGreen, LineWidth:=2, DrawRadii:=DrawRadii()
   DrawFigure 740, 110, GenerateRadii(SeedRadii:="100", Count:=4), Canvas, Angle, vbGreen, LineWidth:=2, DrawRadii:=DrawRadii()
   DrawFigure 950, 110, GenerateRadii(SeedRadii:="100", Count:=3), Canvas, Angle, vbGreen, LineWidth:=2, DrawRadii:=DrawRadii()
   DrawFigure 110, 320, GenerateRadii(SeedRadii:="100", Count:=2), Canvas, Angle, vbGreen, LineWidth:=2, DrawRadii:=DrawRadii()
   DrawFigure 320, 320, GenerateRadii(SeedRadii:="50,100", Count:=5), Canvas, Angle, vbRed, LineWidth:=2, DrawRadii:=DrawRadii()
   DrawFigure 530, 320, GenerateRadii(SeedRadii:="100,100,80,80", Count:=15), Canvas, Angle, vbRed, LineWidth:=2, DrawRadii:=DrawRadii()
   DrawFigure 850, 430, ExtendFigure(GenerateRadii(SeedRadii:="100", Count:=3), GenerateRadii(SeedRadii:="200", Count:=1)), Canvas, Angle, vbBlue, 3, DrawRadii:=DrawRadii()

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Sub

'This procedure manages the option that indicates whether or not radians will be drawn.
Public Function DrawRadii(Optional NewDrawRadii As Variant) As Boolean
On Error GoTo ErrorTrap
Static CurrentDrawRadii As Boolean

   If Not IsMissing(NewDrawRadii) Then
      CurrentDrawRadii = CBool(NewDrawRadii)
   End If

EndProcedure:
   DrawRadii = CurrentDrawRadii
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function

'This procedure extends one polygon by appending another polygon and returns the result.
Private Function ExtendFigure(Figure() As Long, Extension() As Long) As Long()
On Error GoTo ErrorTrap
Dim Extended() As Long
Dim ExtensionSize As Long
Dim UnextendedSize As Long

   Extended() = Figure()
   
   ExtensionSize = (UBound(Extension()) - LBound(Extension())) + 1
   UnextendedSize = UBound(Extended()) - LBound(Extended())
   
   ReDim Preserve Extended(LBound(Extended()) To UBound(Extended()) + ExtensionSize) As Long
   
   RtlMoveMemory Extended(UnextendedSize + 1), Extension(LBound(Extension())), (ExtensionSize * 4)
   
EndProcedure:
   ExtendFigure = Extended()
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function

'This procedure generates a set of radii using the specified seed radii and returns the result.
Private Function GenerateRadii(Optional SeedRadii As String = vbNullString, Optional Count As Long = 1) As Long()
On Error GoTo ErrorTrap
Dim CommaPosition As Long
Dim NewRadian As Long
Dim Radian As Long
Dim Radii() As Long
Dim RemainingRadii As String
   
   If Not SeedRadii = vbNullString Then
      For Radian = 1 To Count
         RemainingRadii = SeedRadii
         If Not Right$(Trim$(RemainingRadii), 1) = "," Then RemainingRadii = RemainingRadii & ","
         Do Until RemainingRadii = vbNullString Or DoEvents() = 0
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
      Next Radian
   End If
   
EndProcedure:
   GenerateRadii = Radii()
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function

'This procedure handles any errors that occur.
Public Function HandleError(Optional ReturnPreviousChoice As Boolean = False) As Long
Dim Description As String
Dim ErrorCode As Long
Static Choice As Long

   Description = Err.Description
   ErrorCode = Err.Number
   On Error Resume Next
   If Not ReturnPreviousChoice Then
      Choice = MsgBox(Description & "." & vbCr & "Error code: " & CStr(ErrorCode), vbAbortRetryIgnore Or vbDefaultButton2 Or vbExclamation)
   End If
   
   If Choice = vbAbort Then End
   
   HandleError = Choice
End Function

'This procedure is executed when this program is started.
Public Sub Main()
On Error GoTo ErrorTrap
   
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path

   InterfaceWindow.Show

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Sub

'This procedure returns information about this program.
Public Function ProgramInformation() As String
On Error GoTo ErrorTrap
Dim Information As String

   With App
      Information = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName
   End With

EndProcedure:
   ProgramInformation = Information
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function


