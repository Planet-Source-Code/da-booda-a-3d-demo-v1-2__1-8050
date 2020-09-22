Attribute VB_Name = "Engine"
Rem *** This is where it all happens
'       The object is manipulate and displayed
'       in this module***

Public Ver(5000, 2) As Single 'Vertices
Public TempV(5000, 2) As Single 'Temp Verteces
Public TempZ(5000) 'used to keep zcenter
Public TempLin(5000) 'Keeps the polygons to show
Public Lin(5000, 23) As Single 'Polygons
Public Zcol(5000) 'Keeps maxis of each polygon for color shading
Public Zcent1 As Double, Zcent2 As Double 'dummys for z center
'Scale,XTranslate,YTranslate
Public Sa As Double, TransX As Double, TransY As Double
'Ztranslate,XRotate,YRotate,ZRotate
Public TransZ As Double, RotX As Double, RotY As Double, RotZ As Double
'Zoom,Number of vertices,Number of polygons
Public Zm As Double, VerNum As Double, LineNum As Double
Public ViewDis As Double 'View distance tells where behind the camera is
'Dummys for the minimums and maxes of polygons
Public Xmin1 As Double, Xmin2 As Double, Xmax1 As Double, Xmax2 As Double
Public Ymin1 As Double, Ymin2 As Double, Ymax1 As Double, Ymax2 As Double
Public Zmin1 As Double, Zmin2 As Double, Zmax1 As Double, Zmax2 As Double
'dummys for calculating
Public X1 As Double, X2 As Double, X3 As Double
Public Y1 As Double, Y2 As Double, Y3 As Double
Public Z1 As Double, Z2 As Double, Z3 As Double
'More dummys
Public Zmin As Double, Zmax As Double
'the center of the screen
Public XOrigin As Double, YOrigin As Double
'Yet More Dummys
Public a1 As Double, b1 As Double, c1 As Double, d1 As Double
'DotProduct tells which way the polygon is facing
Public DotProduct As Double
'How many polygons to Draw
Public TempNum As Integer
'Draw Mode
Public Dm As Integer
'animation toggles
Public Xa, Ya, Za
'Sets the type for the polygon draw function
Public Type POINTAPI
  x As Long
  y As Long
End Type
'the maximum points to each polygon
Public Pol(20) As POINTAPI
'the polygon draw function
Private Declare Function Polygon Lib "gdi32" _
  (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function Polyline Lib "gdi32" _
  (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
'The Set form on top Variables and functions
Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
    'Sets the specified form on top, or bottom
    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hwnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub
Public Sub FindZcol()
'This routine Finds the maxes of all the polygons
'thus when the max is equal to zcol it shows
'the full color
For a = 0 To LineNum 'Go through all the polygons
    'Set variables to first point
    Xmin1 = Ver(Lin(a, 1), 0)
    Xmax1 = Xmin1
    Ymin1 = Ver(Lin(a, 1), 1)
    Ymax1 = Ymin1
    Zmin1 = Ver(Lin(a, 1), 2)
    Zmax1 = Zmin1
        For b = 1 To Lin(a, 0)
        'Set the comparisons to the first point
        Xmin2 = Ver(Lin(a, b), 0)
        Ymin2 = Ver(Lin(a, b), 1)
        Zmin2 = Ver(Lin(a, b), 2)
        If Xmin1 > Xmin2 Then Xmin1 = Xmin2
        If Xmax1 < Xmin2 Then Xmax1 = Xmin2
        If Ymin1 > Ymin2 Then Ymin1 = Ymin2
        If Ymax1 < Ymin2 Then Ymax1 = Ymin2
        If Zmin1 > Zmin2 Then Zmin1 = Zmin2
        If Zmax1 < Zmin2 Then Zmax1 = Zmin2
        Next b
    'finds the length
    X1 = Xmax1 - Xmin1
    Y1 = Ymax1 - Ymin1
    Z1 = Zmax1 - Zmin1
    'Decides the largest and it becomes Zcol
    If X1 >= Y1 And X1 >= Z1 Then Zcol(a) = X1
    If Y1 >= X1 And Y1 >= Z1 Then Zcol(a) = Y1
    If Z1 >= X1 And Z1 >= Y1 Then Zcol(a) = Z1
Next a
End Sub
Public Sub Change()
'This routine is where the object gets changed according
'to drawmode
Convert 'changes the Vertices over to temp vertices
Translate (TransX), (TransY), (TransZ) 'Translates vertices
Rotate (RotX), (RotY), (RotZ) 'Rotates vertices
Zoom (Zm) 'adds zoom value to vertices
Transform 'Creates the perspective
If Dm < 3 Then SortZ 'if backface removal is required the do this routine
If Dm = 3 Then 'Since ther is no sortz, turns every polygon to temp
For a = 0 To LineNum 'go through all polygons
    TempLin(a) = a 'turn to temp
Next a
TempNum = LineNum 'Change tempnum to all
End If
If Dm < 4 Then ImageView 'if polygon or line then display
If Dm = 4 Then 'if vertice then draw circles
View.Display.DrawStyle = 0 'Set drawstyle to solid
View.Display.FillStyle = 1 'set fill style to transparent
View.Display.Cls 'clear screen
For a = 0 To VerNum 'go through all vertices
    'Draw white circle to represent vertex
    View.Display.Circle (TempV(a, 0), TempV(a, 1)), 4, RGB(255, 255, 255)
Next a
View.Display.DrawStyle = 5 'Set drawstyle back to transparent
View.Display.FillStyle = 0 'set fillstyle back to solid
End If
End Sub
Public Sub Convert()
'Turns all vertices into temp vertices
'That way to information doesn't get changed
Dim a, b
For a = 0 To VerNum
    For b = 0 To 2
        TempV(a, b) = Ver(a, b)
Next b, a
End Sub
Public Sub Sca(S As Double)
'Multiplys all the temp vertices by Scale
Dim a, b
For a = 0 To VerNum
    For b = 0 To 2
        TempV(a, b) = TempV(a, b) * S
Next b, a
End Sub
Public Sub Translate(x As Double, y As Double, z As Double)
'Adds translate values to all the temp vertices
Dim a
For a = 0 To VerNum
    TempV(a, 0) = TempV(a, 0) + x
    TempV(a, 1) = TempV(a, 1) + y
    TempV(a, 2) = TempV(a, 2) + z
Next a
End Sub
Public Sub Rotate(Xx As Double, Yy As Double, z As Double)
'Rotates all the tempvertices
Dim Xn As Double, Yn As Double, Zn As Double
'turns the angle degrees into radians
Xn = Xx * (3.141592654 / 180)
Yn = Yy * (3.141592654 / 180)
Zn = z * (3.141592654 / 180)
For a = 0 To VerNum

'Rotate on z axis
Xx = TempV(a, 0): Yy = TempV(a, 1): z = TempV(a, 2)
X1 = Xx * Cos(Zn) - Yy * Sin(Zn)
Y1 = Yy * Cos(Zn) + Xx * Sin(Zn)
Z1 = z
TempV(a, 0) = X1: TempV(a, 1) = Y1: TempV(a, 2) = Z1

'Rotate on y axis
Xx = TempV(a, 0): Yy = TempV(a, 1): z = TempV(a, 2)
Z1 = z * Cos(Yn) - Xx * Sin(Yn)
X1 = z * Sin(Yn) + Xx * Cos(Yn)
Y1 = Yy
TempV(a, 0) = X1: TempV(a, 1) = Y1: TempV(a, 2) = Z1

'rotate on x axis
Xx = TempV(a, 0): Yy = TempV(a, 1): z = TempV(a, 2)
Y1 = Yy * Cos(Xn) - z * Sin(Xn)
Z1 = Yy * Sin(Xn) + z * Cos(Xn)
X1 = Xx
TempV(a, 0) = X1: TempV(a, 1) = Y1: TempV(a, 2) = Z1

Next a
End Sub
Public Sub Zoom(z As Double)
'Adds zoom value to all temp vertices
Dim a
For a = 0 To VerNum
    TempV(a, 2) = TempV(a, 2) + z
Next a
End Sub
Public Sub Transform()
'adds perspective to all vertices
Dim a
For a = 0 To VerNum
'Set up dummy variables
Xx = TempV(a, 0): Yy = TempV(a, 1): z = TempV(a, 2)
'sets the vertice in bounds
If z < -999 Then z = -999
If z > 999 Then z = 999
'Creates a difference between vertices x,y
'Divides by the z value and multiplys by 1000
'In essence the further away a vertice is the
'closer to center it is
z = z + 1000
z = 2000 - z
Xx = (Xx / z) * 1000
Yy = (Yy / z) * 1000
'centers vertice
TempV(a, 0) = XOrigin + Xx
TempV(a, 1) = YOrigin + Yy
Next a
End Sub
Public Sub SortZ()
'This is where backfaces are removed
'and polygons that are seen are then sorted according
'to there z center
Dim a, Dummy As Double, Sw
If LineNum = 0 Then GoTo Zend 'if there is one polygon, no need to sort
TempNum = -1 'When a polygon is said to be seen it is added to templin and tempnum is added by one
For a = 0 To LineNum 'Go throug all the lines
'these next few lines are finding the min's and maxes of all the polygons
Xmin1 = TempV(Lin(a, 1), 0)
Xmax1 = Xmin1
Ymin1 = TempV(Lin(a, 1), 1)
Ymax1 = Ymin1
Zmin1 = TempV(Lin(a, 1), 2)
Zmax1 = Zmin1
For b = 1 To Lin(a, 0)
X1 = TempV(Lin(a, b), 0)
Y1 = TempV(Lin(a, b), 1)
Z1 = TempV(Lin(a, b), 2)
If Xmin1 > X1 Then Xmin1 = X1
If Xmax1 < X1 Then Xmax1 = X1
If Ymin1 > Y1 Then Ymin1 = Y1
If Ymax1 < Y1 Then Ymax1 = Y1
If Zmin1 > Z1 Then Zmin1 = Z1
If Zmax1 < Z1 Then Zmax1 = Z1
Next b
'Finds the dot product of the polygon, according to
'the first 3 points of the polygon
'if these points are going in a counterclockwise
'pattern then it is seen
'if clockwise remove it
b = 0 'sets remove toggle to zero
'the dot product formula
X1 = TempV(Lin(a, 1), 0): X2 = TempV(Lin(a, 2), 0): X3 = TempV(Lin(a, 3), 0)
Y1 = TempV(Lin(a, 1), 1): Y2 = TempV(Lin(a, 2), 1): Y3 = TempV(Lin(a, 3), 1)
Z1 = TempV(Lin(a, 1), 2): Z2 = TempV(Lin(a, 2), 2): Z3 = TempV(Lin(a, 3), 2)
DotProduct = (X3 * ((Z1 * Y2) - (Y1 * Z2))) + (Y3 * ((X1 * Z2) - (Z1 * X2))) + (Z3 * ((Y1 * X2) - (X1 * Y2)))
'if the polygon is out of the viewing area, then
'b=b+1 and it isn't added to the templin
If Xmin1 < 0 And Xmax1 < 0 Then b = b + 1
If Xmin1 > 800 And Xmax1 > 800 Then b = b + 1
If Ymin1 < 0 And Ymax1 < 0 Then b = b + 1
If Ymin1 > 600 And Ymax1 > 600 Then b = b + 1
If Zmin1 > ViewDis Then b = b + 1
'if b doesn't equal zero then it isn't seen
'therefore not added
If b <> 0 Then GoTo ReSkip
'if dotproduct is counterclockwise then it is seen and added
If DotProduct > 0 Then
TempNum = TempNum + 1
TempLin(TempNum) = a
'Finds the z center of the polygon, used to sort
Zcent1 = Zmax1 - Zmin1: Zcent1 = Zcent1 / 2 + Zmin1
TempZ(a) = Zcent1
End If
ReSkip:
Next a
'this routine is also called the painters algorithim
'If the polygon is behind another then it is drawn first
If TempNum = 0 Then GoTo Zend 'If nothing is seen, there's nothing to sort
ZStart: 'Start the zstart routine
Sw = 0 'Set the swap toggle
For a = 0 To TempNum - 1
Zcent1 = TempZ(TempLin(a)) 'get zcenter of first polygon
Zcent2 = TempZ(TempLin(a + 1)) 'get the polygon to compare
'if the first polygon is further away
'then it is in the right order
'so no need to swap it
If Zcent1 <= Zcent2 Then GoTo Pass
'Swap the two around
Dummy = TempLin(a): TempLin(a) = TempLin(a + 1): TempLin(a + 1) = Dummy
Sw = 1 'tells the routine that a swap did occur, so redo
Pass:
Next a
If Sw <> 0 Then GoTo ZStart 'redo sort routine, if polygons were swapped
Zend:
End Sub
Public Sub ImageView()
'This routine draws the viewable polygons in order
'from back to start
Dim r1, r2, r3
Dim a, z As Double, t1 As Double, t2 As Double, t3 As Double
View.Display.Cls 'clear screen
For a = 0 To TempNum 'go through viewable polygons
'if drawmode isn't polygon the skip color shading process
If Dm <> 1 Then GoTo SkipColor
'find the zmin's and maxes for the color shade formula
Zmin = TempV(Lin(TempLin(a), 1), 2)
Zmax = TempV(Lin(TempLin(a), 1), 2)
For t = 1 To Lin(TempLin(a), 0)
    If Zmin > TempV(Lin(TempLin(a), t), 2) Then Zmin = TempV(Lin(TempLin(a), t), 2)
    If Zmax < TempV(Lin(TempLin(a), t), 2) Then Zmax = TempV(Lin(TempLin(a), t), 2)
Next t
'Start the color shade routine
'this in essence shades the polygon darker
'ass it turns away from the camera
r1 = Lin(TempLin(a), 21): r2 = Lin(TempLin(a), 22): r3 = Lin(TempLin(a), 23)
z = Zmax - Zmin 'find length
'if polygon is turned then z equals the center
If z > Zcol(a) Then z = Zcol(a)
'finds the value of the turn
'in essence the more it's turned
'the more it shades
t1 = (r1 * Sa) / Zcol(a)
t2 = (r2 * Sa) / Zcol(a)
t3 = (r3 * Sa) / Zcol(a)
r1 = Int(r1 - (t1 * z))
r2 = Int(r2 - (t2 * z))
r3 = Int(r3 - (t3 * z))

SkipColor:
If Dm = 2 Then 'sets up values to draw black polygines with white borders
r1 = 0: r2 = 0: r3 = 0
View.Display.DrawStyle = 0
View.Display.ForeColor = RGB(255, 255, 255)
End If
If Dm = 3 Then 'sets up values to draw transparent polygons with white borders
View.Display.DrawStyle = 0
View.Display.ForeColor = RGB(255, 255, 255)
View.Display.FillStyle = 1
End If
For t = 1 To Lin(TempLin(a), 0) 'got through all vertices for polygons
'insert values into polygon draw function
Pol(t - 1).x = TempV(Lin(TempLin(a), t), 0)
Pol(t - 1).y = TempV(Lin(TempLin(a), t), 1)
Next t
View.Display.FillColor = RGB(r1, r2, r3) 'set color to polygon
Polygon View.Display.hdc, Pol(0), Lin(TempLin(a), 0) 'draw it
Next a
'restore values to polygon mode
View.Display.DrawStyle = 5
View.Display.FillStyle = 0
End Sub
