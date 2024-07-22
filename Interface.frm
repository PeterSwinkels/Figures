VERSION 5.00
Begin VB.Form InterfaceWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   645
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
   Begin VB.Menu OptionsMainMenu 
      Caption         =   "&Options"
      Begin VB.Menu AnimateFiguresMenu 
         Caption         =   "&Animate Figures"
         Shortcut        =   ^A
      End
      Begin VB.Menu DrawRadiansMenu 
         Caption         =   "&Draw Radians"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu InformationMainMenu 
      Caption         =   "&Information"
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's main interface window.
Option Explicit

'This procedure toggles drawing radians on/off.
Private Sub ToggleDrawRadians()
On Error GoTo ErrorTrap
Dim CurrentDrawRadii As Boolean

   CurrentDrawRadii = DrawRadii()
   DrawRadii NewDrawRadii:=Not CurrentDrawRadii
   DrawRadiansMenu.Checked = DrawRadii()
   
   DrawFigures Me, AnimatorActive:=Animator.Enabled

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub

'This procedure toggles the animator on/off.
Private Sub AnimateFiguresMenu_Click()
On Error GoTo ErrorTrap
  
   Animator.Enabled = Not Animator.Enabled
   AnimateFiguresMenu.Checked = Animator.Enabled

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub

'This procedure animates the various figures to be displayed by changing their angle.
Private Sub Animator_Timer()
On Error GoTo ErrorTrap
Static Angle As Double

   DrawFigures Me, AnimatorActive:=Animator.Enabled, Angle:=Angle
   
   If Angle >= 360 Then Angle = 0 Else Angle = Angle + 8

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub

'This procedure gives the command to toggle drawing radians on/off.
Private Sub DrawRadiansMenu_Click()
On Error GoTo ErrorTrap
   
   ToggleDrawRadians

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub

'This procedure generates and displays several figures.
Private Sub Form_Activate()
On Error GoTo ErrorTrap
   
   DrawFigures Me, AnimatorActive:=Animator.Enabled, Angle:=0

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   
   Me.Width = Screen.Width / 1.1
   Me.Height = Screen.Height / 1.1
   
   Me.Caption = App.Title

   AnimateFiguresMenu.Checked = Animator.Enabled
   DrawRadiansMenu.Checked = DrawRadii()

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub

'This procedure closes this program when this window is closed.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorTrap
   
   Unload Me
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub

'This procedure displays this program's information.
Private Sub InformationMainMenu_Click()
On Error GoTo ErrorTrap

   MsgBox ProgramInformation(), vbInformation

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbRetry Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub


