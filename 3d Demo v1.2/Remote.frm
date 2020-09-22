VERSION 5.00
Begin VB.Form Remote 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote"
   ClientHeight    =   7470
   ClientLeft      =   9450
   ClientTop       =   375
   ClientWidth     =   2550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   2550
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00808080&
      Caption         =   "E X I T"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7185
      Left            =   2040
      MaskColor       =   &H000000FF&
      TabIndex        =   20
      Top             =   120
      Width           =   375
   End
   Begin VB.HScrollBar ScrZoom 
      Height          =   255
      LargeChange     =   10
      Left            =   240
      Max             =   2000
      TabIndex        =   18
      Top             =   6720
      Value           =   900
      Width           =   1575
   End
   Begin VB.HScrollBar ScrColor 
      Height          =   255
      LargeChange     =   5
      Left            =   240
      Max             =   100
      TabIndex        =   15
      Top             =   5520
      Value           =   50
      Width           =   1575
   End
   Begin VB.HScrollBar ScrZTrans 
      Height          =   255
      LargeChange     =   4
      Left            =   240
      Max             =   9999
      Min             =   -9999
      TabIndex        =   10
      Top             =   4320
      Width           =   1575
   End
   Begin VB.HScrollBar ScrYTrans 
      Height          =   255
      LargeChange     =   4
      Left            =   240
      Max             =   9999
      Min             =   -9999
      TabIndex        =   9
      Top             =   3840
      Width           =   1575
   End
   Begin VB.HScrollBar ScrXTrans 
      Height          =   255
      LargeChange     =   4
      Left            =   240
      Max             =   9999
      Min             =   -9999
      TabIndex        =   8
      Top             =   3360
      Width           =   1575
   End
   Begin VB.HScrollBar ScrZRot 
      Height          =   255
      LargeChange     =   5
      Left            =   240
      Max             =   360
      Min             =   -360
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.HScrollBar ScrYRot 
      Height          =   255
      LargeChange     =   5
      Left            =   240
      Max             =   360
      Min             =   -360
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.HScrollBar ScrXRot 
      Height          =   255
      LargeChange     =   5
      Left            =   240
      Max             =   360
      Min             =   -360
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Line Line28 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   1920
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line27 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   120
      Y1              =   600
      Y2              =   120
   End
   Begin VB.Line Line26 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   120
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line25 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label lblDrawMode 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Mode:Polygon"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   240
      TabIndex        =   21
      Top             =   240
      Width           =   1635
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   840
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00FF8080&
      X1              =   840
      X2              =   840
      Y1              =   6600
      Y2              =   6240
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   120
      Y1              =   6600
      Y2              =   7320
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   1920
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   1920
      Y1              =   7320
      Y2              =   6240
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FF8080&
      X1              =   840
      X2              =   1920
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label LblZm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom : 900"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label LblZoom 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   6240
      Width           =   645
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   1920
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   1920
      Y1              =   6120
      Y2              =   5040
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   120
      Y1              =   5400
      Y2              =   6120
   End
   Begin VB.Label LblSca 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Color :50%"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   1080
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   1080
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FF8080&
      X1              =   1080
      X2              =   1080
      Y1              =   5040
      Y2              =   5400
   End
   Begin VB.Label LblScale 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Shade%"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   5040
      Width           =   900
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   1200
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF8080&
      X1              =   1200
      X2              =   1200
      Y1              =   3240
      Y2              =   2880
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF8080&
      X1              =   1200
      X2              =   1920
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   1920
      Y1              =   2880
      Y2              =   4920
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   120
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   120
      Y1              =   3240
      Y2              =   4920
   End
   Begin VB.Label LblZTrans 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Z : 0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label LblYTrans 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y : 0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label LblXTrans 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X : 0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label LblTranslate 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Translate"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1020
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   120
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   1920
      Y1              =   2760
      Y2              =   720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   960
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF8080&
      X1              =   960
      X2              =   960
      Y1              =   720
      Y2              =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   960
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   120
      Y1              =   1080
      Y2              =   2760
   End
   Begin VB.Label LblZRot 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Z : 0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label LblYRot 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y : 0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label LblXRot 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "X : 0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label LblRotate 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Rotate"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   720
   End
End
Attribute VB_Name = "Remote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem *** the remote form, the gui to control the object ***

Private Sub CmdExit_Click()
End 'Ends the program
End Sub

Private Sub lblDrawMode_Click()
'This routine just changes the Drawmode
'when the label is clicked
If Dm = 1 Then
lblDrawMode.Caption = "Mode:Line"
Dm = 2
Change
GoTo EndDrawMode
End If
If Dm = 2 Then
lblDrawMode.Caption = "Mode:Line2"
Dm = 3
Change
GoTo EndDrawMode
End If
If Dm = 3 Then
lblDrawMode.Caption = "Mode:Vertex"
Dm = 4
Change
GoTo EndDrawMode
End If
If Dm = 4 Then
lblDrawMode.Caption = "Mode:Polygon"
Dm = 1
Change
GoTo EndDrawMode
End If
EndDrawMode: 'This is here so it won't keep changing the draw mode
End Sub

Private Sub LblXRot_Click() 'Turns the XAnimation on or off
'if Xa = 0 then the xanimation is off else it's on
If Xa = 0 Then
Xa = 1
Else
Xa = 0
RotX = ScrXRot.Value
LblXRot.Caption = "X :" + Str$(RotX)
Change
End If
End Sub

Private Sub LblXTrans_Click()
TransX = 0
ScrXTrans.Value = 0
LblXTrans.Caption = "X : 0"
End Sub

Private Sub LblYRot_Click() 'Same as Xanimation
If Ya = 0 Then
Ya = 1
Else
Ya = 0
RotY = ScrYRot.Value
LblYRot.Caption = "Y :" + Str$(RotY)
Change
End If

End Sub

Private Sub LblYTrans_Click()
TransY = 0
ScrYTrans.Value = 0
LblYTrans.Caption = "Y : 0"
End Sub

Private Sub LblZRot_Click() 'Same as Xanimation
If Za = 0 Then
Za = 1
Else
Za = 0
RotZ = ScrZRot.Value
LblZRot.Caption = "Z :" + Str$(RotZ)
Change
End If

End Sub


Private Sub LblZTrans_Click()
TransZ = 0
ScrZTrans.Value = 0
LblZTrans.Caption = "Z : 0"
End Sub

Private Sub ScrColor_Change() 'change color shading percentage
LblSca.Caption = "Color :" + Str$(ScrColor.Value) + "%"
Sa = ScrColor.Value * 0.01
Change
End Sub

Private Sub ScrColor_Scroll()
LblSca.Caption = "Color :" + Str$(ScrColor.Value) + "%"
Sa = ScrColor.Value * 0.01
Change
End Sub

Private Sub ScrXRot_Change()
LblXRot.Caption = "X :" + Str$(ScrXRot.Value)
RotX = ScrXRot.Value
Change
End Sub

Private Sub ScrXRot_Scroll()
LblXRot.Caption = "X :" + Str$(ScrXRot.Value)
RotX = ScrXRot.Value
Change
End Sub

Private Sub ScrXTrans_Change()
LblXTrans.Caption = "X :" + Str$(ScrXTrans.Value)
TransX = ScrXTrans.Value
Change
End Sub

Private Sub ScrXTrans_Scroll()
LblXTrans.Caption = "X :" + Str$(ScrXTrans.Value)
TransX = ScrXTrans.Value
Change
End Sub

Private Sub ScrYRot_Change()
LblYRot.Caption = "Y :" + Str$(ScrYRot.Value)
RotY = ScrYRot.Value
Change
End Sub

Private Sub ScrYRot_Scroll()
LblYRot.Caption = "Y :" + Str$(ScrYRot.Value)
RotY = ScrYRot.Value
Change
End Sub

Private Sub ScrYTrans_Change()
LblYTrans.Caption = "Y :" + Str$(ScrYTrans.Value)
TransY = ScrYTrans.Value
Change
End Sub

Private Sub ScrYTrans_Scroll()
LblYTrans.Caption = "Y :" + Str$(ScrYTrans.Value)
TransY = ScrYTrans.Value
Change
End Sub

Private Sub ScrZoom_Change() 'change Zoom value
LblZm.Caption = "Zoom :" + Str$(ScrZoom.Value)
Zm = ScrZoom.Value
Change
End Sub

Private Sub ScrZoom_Scroll()
LblZm.Caption = "Zoom :" + Str$(ScrZoom.Value)
Zm = ScrZoom.Value
Change
End Sub

Private Sub ScrZRot_Change()
LblZRot.Caption = "Z :" + Str$(ScrZRot.Value)
RotZ = ScrZRot.Value
Change
End Sub

Private Sub ScrZRot_Scroll()
LblZRot.Caption = "Z :" + Str$(ScrZRot.Value)
RotZ = ScrZRot.Value
Change
End Sub

Private Sub ScrZTrans_Change()
LblZTrans.Caption = "Z :" + Str$(ScrZTrans.Value)
TransZ = ScrZTrans.Value
Change
End Sub

Private Sub ScrZTrans_Scroll()
LblZTrans.Caption = "Z :" + Str$(ScrZTrans.Value)
TransZ = ScrZTrans.Value
Change
End Sub
