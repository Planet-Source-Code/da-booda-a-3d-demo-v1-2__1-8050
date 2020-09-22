VERSION 5.00
Begin VB.Form View 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   -14940
   ClientTop       =   240
   ClientWidth     =   12000
   DrawStyle       =   5  'Transparent
   DrawWidth       =   4
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Display 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   9000
      Left            =   0
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin VB.Timer Anim 
         Interval        =   10
         Left            =   2640
         Top             =   2280
      End
   End
End
Attribute VB_Name = "View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem ***  This is where th object is displayed ***

Private Sub Anim_Timer()
'this is the animation routine for the clock
'which is set at a duration of tem
'if xa,ya,or za is set to 1 it adds 5
'to the complimenting rotation value
If Xa = 1 Then RotX = RotX + 5
If Ya = 1 Then RotY = RotY + 5
If Za = 1 Then RotZ = RotZ + 5
'Makes sure the rotation don't go out of bounds
If RotX > 360 Then RotX = -355
If RotY > 360 Then RotY = -355
If RotZ > 360 Then RotZ = -355
'change labels
If Xa = 1 Then Remote.LblXRot.Caption = "[X] :" + Str$(RotX)
If Ya = 1 Then Remote.LblYRot.Caption = "[Y] :" + Str$(RotY)
If Za = 1 Then Remote.LblZRot.Caption = "[Z] :" + Str$(RotZ)
'changes and displays object if animation is on
If Xa = 1 Or Ya = 1 Or Za = 1 Then Change
End Sub


Private Sub Form_Load()
'Put view form in top left-hand corner
Me.Top = 0
Me.Left = 0
End Sub
