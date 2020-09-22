VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Form1"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   1740
   DrawStyle       =   3  'Dash-Dot
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   1740
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************
'*DA'                            *
'* +-+    +-+   +-+  +-+    +-+  *
'* |  )  |   | |   | |  |  |   | *
'* +-+   |___| |___| |  |  +---+ *
'* |  )  |  *| |  *| |  |  |   | *
'* +-+    +-+   +-+  +-+   |   | *
'*                               *
'*********************************
'
' 3D Demo v1.2
' May 2000
' Feel free to use this code, just remember to give credit
' where credit is due.....I would do the same for you.
'
' Da Booda
' Any comments or questions...
' Email vpbooda@hotmail.com
'
' P.S.
' This is the updated version of 3d demo, now with
' polygon fill routines...although still a little touchy.
' If anybody knows of a better way to SortZ or to speed
' the process up, Please email me the changes.
' I've now added comment lines to everything, it should
' help in understanding the code better.
' I've also added a few new commands to the remote
' They don't do much just there for fun.

' Instructions on the Remote:
'   The Mode Box:
'       Tells the user what draw mode the display is in.
'       Click on the label and it will change modes.
'       The 4 modes are;
'       Polygon - fills the object with color
'       Line1 - Shows just lines with backface removal
'       Line2 - Shows lines without backface removal
'       Vertex - Just shows vertexes as circles
'
'   The Rotate Box:
'       Pretty simple really, just use the scroll bars
'       to rotate the object on the 3 axis, which rotate
'       at point 0,0,0
'       However there is a new feature, just click on
'       the labels to animate the object.
'
'   The Translate Box:
'       Again just use the scroll bars to move the object.
'       Click on the labels to take back to zero.
'
'   The Shade Box:
'       This is the percentage to shade the Object when
'       it turns toward the horizon.
'
'   The Zoom Box:
'       Just used to zoom the camera.
'End instructions.
'
'Instructions on how to make an object:
'
'   The vertex points are based on the Cartesian
'   Coordinate system.
'   Left = -x
'   Right = +x
'   Up = -y
'   Down = +y
'   Away = -z
'   Towards = +z
'   apply points as such Ver(0,0)=x,Ver(0,1)=y,Ver(0,2)=z
'
'   The lin property is base on this
'   lin(0,0)=How many points the polygon consists of
'   lin(0,1-20) is the verteces the point apply to
'   lin(0,21-23) is the color as rgb
'   Note:   When Building a polygon, you have to go in
'           a counter-clockwise pattern, or the
'           Back face removal routine won't work.
'           Form it as if you are looking at it.
'           See examples below.


Private Sub Form_Load()
View.Show 'Show view form
Remote.Show 'Show remote form
AlwaysOnTop Remote, True 'put remote on top

Sa = 0.5 'Set Shade percentage to 50%, higher for darker, less for lighter

TransX = 0: TransY = 0: TransZ = 0 'set translate to 0

RotX = 0: RotY = 0: RotZ = 0 'set rotation values to 0

Zm = 900 'set zoom to 0

'The ViewDis variable tells to program not to draw
'a polygon when the z center is past this point.
'Keeps the screen from filling with color
ViewDis = 999

XOrigin = 400 'These two are to set the center of the view screen
YOrigin = 300

Dm = 1 'Sets the draw mode to Polygon

a = 4 'Sets what object is to be made
On a GoTo Box, Pyramid, Cone, Cylinder

'The following are premade objects, fool around and make your own.
Box:
VerNum = 7
Rem vertices
Ver(0, 0) = -5: Ver(0, 1) = -5: Ver(0, 2) = -5
Ver(1, 0) = 5: Ver(1, 1) = -5: Ver(1, 2) = -5
Ver(2, 0) = 5: Ver(2, 1) = 5: Ver(2, 2) = -5
Ver(3, 0) = -5: Ver(3, 1) = 5: Ver(3, 2) = -5

Ver(4, 0) = -5: Ver(4, 1) = -5: Ver(4, 2) = 5
Ver(5, 0) = 5: Ver(5, 1) = -5: Ver(5, 2) = 5
Ver(6, 0) = 5: Ver(6, 1) = 5: Ver(6, 2) = 5
Ver(7, 0) = -5: Ver(7, 1) = 5: Ver(7, 2) = 5

LineNum = 5
Rem Polygons
Lin(0, 1) = 4: Lin(0, 2) = 7: Lin(0, 3) = 6: Lin(0, 4) = 5
Lin(1, 1) = 0: Lin(1, 2) = 1: Lin(1, 3) = 2: Lin(1, 4) = 3
Lin(2, 1) = 0: Lin(2, 2) = 4: Lin(2, 3) = 5: Lin(2, 4) = 1
Lin(3, 1) = 7: Lin(3, 2) = 3: Lin(3, 3) = 2: Lin(3, 4) = 6
Lin(4, 1) = 0: Lin(4, 2) = 3: Lin(4, 3) = 7: Lin(4, 4) = 4
Lin(5, 1) = 5: Lin(5, 2) = 6: Lin(5, 3) = 2: Lin(5, 4) = 1
For a = 0 To 5
Lin(a, 0) = 4
Next a
Rem colors
For a = 0 To LineNum
    Lin(a, 23) = 255
Next a
GoTo EndMake

Pyramid:
VerNum = 4
Rem Vertices
Ver(0, 0) = 0: Ver(0, 1) = 0: Ver(0, 2) = 5
Ver(1, 0) = -5: Ver(1, 1) = -5: Ver(1, 2) = -5
Ver(2, 0) = 5: Ver(2, 1) = -5: Ver(2, 2) = -5
Ver(3, 0) = 5: Ver(3, 1) = 5: Ver(3, 2) = -5
Ver(4, 0) = -5: Ver(4, 1) = 5: Ver(4, 2) = -5
LineNum = 4
Rem Lines
Lin(0, 0) = 3: Lin(0, 1) = 0: Lin(0, 2) = 4: Lin(0, 3) = 3
Lin(1, 0) = 3: Lin(1, 1) = 0: Lin(1, 2) = 3: Lin(1, 3) = 2
Lin(2, 0) = 3: Lin(2, 1) = 0: Lin(2, 2) = 2: Lin(2, 3) = 1
Lin(3, 0) = 3: Lin(3, 1) = 0: Lin(3, 2) = 1: Lin(3, 3) = 4
Lin(4, 0) = 4: Lin(4, 1) = 1: Lin(4, 2) = 2: Lin(4, 3) = 3: Lin(4, 4) = 4
Rem Colors
For a = 0 To 4
    Lin(a, 21) = 255
    Lin(a, 22) = 255
Next a
GoTo EndMake

Cone:
VerNum = 12
Rem Vertices
Ver(0, 0) = 0: Ver(0, 1) = 0: Ver(0, 2) = 10
For t = 1 To 12
    Zn = ((t - 1) * 30) * (3.141592654 / 180)
    Ver(t, 0) = 0 * Cos(Zn) - -5 * Sin(Zn)
    Ver(t, 1) = -5 * Cos(Zn) + 0 * Sin(Zn)
    Ver(t, 2) = -10
Next t
LineNum = 12
Rem Polygons
For t = 0 To 10
    Lin(t, 0) = 3
    Lin(t, 1) = 0
    Lin(t, 2) = t + 2
    Lin(t, 3) = t + 1
Next t
Lin(11, 0) = 3
Lin(11, 1) = 0: Lin(11, 2) = 1: Lin(11, 3) = 12
Lin(12, 0) = 12
For t = 1 To 12
    Lin(12, t) = t
Next t
Rem colors
For t = 0 To 12
    Lin(t, 22) = 255
    Lin(t, 23) = 100
Next t
GoTo EndMake

Cylinder:
VerNum = 23
Rem Vertices
For t = 0 To 11
    Zn = ((t - 1) * 30) * (3.141592654 / 180)
    Ver(t, 0) = 0 * Cos(Zn) - -5 * Sin(Zn)
    Ver(t, 1) = -5 * Cos(Zn) + 0 * Sin(Zn)
    Ver(t, 2) = -10
Next t
For t = 12 To 23
    Ver(t, 0) = Ver(t - 12, 0)
    Ver(t, 1) = Ver(t - 12, 1)
    Ver(t, 2) = 10
Next t
LineNum = 13
Rem Polygons
For t = 0 To 10
    Lin(t, 0) = 4
    Lin(t, 1) = t + 1
    Lin(t, 2) = t
    Lin(t, 3) = t + 12
    Lin(t, 4) = t + 13
Next t
Lin(11, 0) = 4: Lin(11, 1) = 0: Lin(11, 2) = 11: Lin(11, 3) = 23: Lin(11, 4) = 12
Lin(12, 0) = 12
Lin(13, 0) = 12
For t = 1 To 12
    Lin(12, t) = t - 1
    Lin(13, t) = 24 - t
Next t
Rem Colors
For t = 0 To 13
    Lin(t, 21) = 0
    Lin(t, 22) = 100
    Lin(t, 23) = 200
Next t
GoTo EndMake

EndMake:
FindZcol 'find the maxes of each polygon for the color shading routine
Change 'This runs all the change funtions to the object and displays it
End Sub
