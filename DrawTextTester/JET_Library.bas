Attribute VB_Name = "JET_Library"
Option Explicit

Const PI As Single = 3.14159265358979






Sub DrawGrid(PicBoxObject As Object)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   GRID DRAWING SUB
    '   ''''''''''''''''
    '
    '                         09/03/2007
    '      Programmed by: JAMES E. TRACY
    '        Sacramento, California, USA
    '          JamesTracy95820@gmail.com
    '                    Copyright, 2007
    '                    ^^^^^^^^^^^^^^^
    '
    '   DESCRIPTION
    '   ^^^^^^^^^^^
    '
    '   Draws a grid on any object that can accept the LINE method.  Usually a picture box.
    '
    '   EXAMPLE USAGE
    '   ^^^^^^^^^^^^^
    '
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   FUNCTION PARAMETERS
    '   ^^^^^^^^^^^^^^^^^^^
    '
    '   DrawGrid Me.Picture1            '   Draw a grid on the picture box.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim L1 As Single
    For L1 = 0# To PicBoxObject.ScaleWidth Step 10#
        PicBoxObject.Line (L1, 0)-(L1, Abs(PicBoxObject.ScaleHeight)), RGB(75, 75, 75)
    Next L1
    For L1 = 0# To Abs(PicBoxObject.ScaleHeight) Step 10#
        PicBoxObject.Line (0, L1)-(PicBoxObject.ScaleWidth, L1), RGB(75, 75, 75)
    Next L1
End Sub


Sub FlipScale(Obj As Object)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   FLIP SCALE SUB
    '   ''''''''''''''
    '
    '                         09/01/2007
    '      Programmed by: JAMES E. TRACY
    '        Sacramento, California, USA
    '          JamesTracy95820@gmail.com
    '                    Copyright, 2007
    '                    ^^^^^^^^^^^^^^^
    '
    '   DESCRIPTION
    '   ^^^^^^^^^^^
    '
    '   When you create a picture box, it's coordinate system originates in the upper
    '   left-hand corner instead of the lower left-hand corner that we're all used
    '   to.  This program flips the coordinates by modifying the Scale: Height, Left,
    '   Top and Width properties.
    '
    '   It's only necessary to call this subroutine one time to set those values.  Calling
    '   it a second time will cause a scale other than what is intended.
    '
    '   EXAMPLE USAGE
    '   ^^^^^^^^^^^^^
    '
    '   FlipScale Form2.Picture1
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   FUNCTION PARAMETERS
    '   ^^^^^^^^^^^^^^^^^^^
    '
    '   Obj is any object with scale parameters that need to be flipped.  Usually
    '   this would be a picture box.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Obj.ScaleLeft = 0#
    Obj.ScaleTop = Obj.ScaleHeight
    Obj.ScaleWidth = Obj.ScaleWidth
    Obj.ScaleHeight = -Obj.ScaleHeight
End Sub


Function iMin(ByVal iIn1 As Integer, ByVal iIn2 As Integer) As Integer
    Dim iReturn As Integer
    If iIn1 < iIn2 Then
        iReturn = iIn1
    Else
        iReturn = iIn2
    End If
    iMin = iReturn
End Function

Function iMax(ByVal iIn1 As Integer, ByVal iIn2 As Integer) As Integer
    Dim iReturn As Integer
    If iIn1 > iIn2 Then
        iReturn = iIn1
    Else
        iReturn = iIn2
    End If
    iMax = iReturn
End Function



Function FileExists(ByVal Fname As String) As Boolean
    '
    '
    '   Returns True  if file exists
    '           False if file does not exist.
    '
    If Fname = "" Or Right(Fname, 1) = "\" Then
        FileExists = False: Exit Function
    End If
    FileExists = (Dir(Fname) <> "")
End Function



Sub sRotateZ(ByVal AngleIn As Single, _
    ByVal XIn As Single, _
    ByVal YIn As Single, _
    ByVal XOrigin As Single, _
    ByVal YOrigin As Single, _
    ByRef XRotated As Single, _
    ByRef YRotated As Single)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   SINGLE PRECISION POINT ROTATION
    '   '''''''''''''''''''''''''''''''
    '
    '                         09/02/2007
    '      Programmed by: JAMES E. TRACY
    '        Sacramento, California, USA
    '          JamesTracy95820@gmail.com
    '                    Copyright, 2007
    '                    ^^^^^^^^^^^^^^^
    '
    '   DESCRIPTION
    '   ^^^^^^^^^^^
    '
    '   Rotates an X/Y point around an X/Y point at an angle.  Return new X/Y point.
    '
    '   EXAMPLE USAGE
    '   ^^^^^^^^^^^^^
    '   Dim X11 As Single
    '   Dim Y11 As Single
    '   sRotateZ sdtParms(5), X1, Y1, sdtParms(0), sdtParms(1), X11, Y11
    '   X1 = X11
    '   Y1 = Y11
    '
    '   I had trouble passing the same variable as what I got back (X1,Y1), and
    '   found that the compiler liked it when I handed it off as the example shows.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   FUNCTION PARAMETERS
    '   ^^^^^^^^^^^^^^^^^^^
    '
    '   AngleIn As Single   =   The angle amount to rotate.
    '   XIn As Single       =   X point to rotate.
    '   YIn As Single       =   Y point to rotate.
    '   XOrigin As Single   =   X origin around which to rotate.
    '   YOrigin As Single   =   Y origin around which to rotate.
    '   XRotated As Single  =   X resultant rotated point.
    '   YRotated As Single  =   Y resultant rotated point.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    XRotated = XOrigin + (XIn - XOrigin) * Cos(AngleIn * (PI / 180#)) + (YIn - YOrigin) * Sin(AngleIn * (PI / 180#))
    YRotated = YOrigin + (YIn - YOrigin) * Cos(AngleIn * (PI / 180#)) - (XIn - XOrigin) * Sin(AngleIn * (PI / 180#))
End Sub


