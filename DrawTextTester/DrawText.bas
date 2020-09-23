Attribute VB_Name = "DrawTextMod"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DrawText() DECLARATIONS
'   ^^^^^^^^^^^^^^^^^^^^^^^
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const TOTFONTS As Byte = 3 '''''''''''''''''''''''''''''' Total fonts available (0-TOTFONTS).
Const MAXVECTORS As Integer = 5000 '''''''''''''''''''''' Maximum number of lines per font.
Dim CharSet As String ''''''''''''''''''''''''''''''''''' The full characters set (94 characters).
Dim FontFileNames(0 To TOTFONTS) As String '''''''''''''' Font file names.
Dim VectorCount(TOTFONTS) As Integer '''''''''''''''''''' Number of vectors in each array.
Dim VectorBounds(0 To TOTFONTS, 0 To 93, 0 To 1) As Integer 'Lower/Upper bounds of each character.
Dim CarriageLocation As Single '''''''''''''''''''''''''' Next position for character.
Dim CharNum As Integer '''''''''''''''''''''''''''''''''' Current character number.
Dim X1, Y1, X2, Y2 As Single '''''''''''''''''''''''''''' Current vector being worked on.
Dim Vectors(0 To TOTFONTS, 0 To MAXVECTORS, 0 To 3) As Single
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'        Vectors(TOTFONT): Rec#  Col0  Col1  Col2  Col3
'                          ----  ----  ----  ----  ----
'                           1     X1    Y1    X2    Y2
'                           2     X1    Y1    X2    Y2
'                        ...n     X1    Y1    X2    Y2
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub DrawText(PicObj As Object, sdtParms() As Single, idtParms() As Integer, TextString As String)

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   DrawText() FUNCTION - TEXT TO LINES PROCEDURE
    '   '''''''''''''''''''''''''''''''''''''''''''''
    '
    '                Created: 08/28/2007
    '           Last Updated: 09/02/2007
    '      Programmed by: JAMES E. TRACY
    '        Sacramento, California, USA
    '          JamesTracy95820@gmail.com
    '                    Copyright, 2007
    '                    ^^^^^^^^^^^^^^^
    '
    '   DESCRIPTION
    '   ^^^^^^^^^^^
    '
    '   Occasionally, the ability to draw text on a picture box using the LINE method
    '   is needed, as it would allow complete control over a it's scale, location and
    '   other factors.  This function accomplishes that by taking a string of characters
    '   and drawing them as lines, using the LINE method.  Color, scale, rotation and
    '   other factors can be completely controlled.
    '
    '   The function is scale independent, so it doesn't matter if you're using twips,
    '   pixels, or any other measurement system.  The letters are one square unit in size,
    '   by default.  If you are using pixels, that means each character is one square
    '   pixel in size.  If scaled up by a factor of 20, then one character would fit within
    '   exactly 20 pixels.  You can control the scaling on both the X and Y-axis.
    '
    '   There are currently three fonts available, but fonts can be designed fairly easily.
    '   You can even design a font of your own handwriting.  If you'd like to design your
    '   own font, just contact me at JamesTracy95820@gmail.com.
    '
    '   I wrote a little program that compiles the fonts from DXF files.  I used AutoCAD
    '   release 14 to design the three fonts used by DrawText(), but any drawing program
    '   capable of outputting DXF files in AutoCAD release 14 format would work fine.  I'd
    '   be happy to accept new fonts!  The more the merrier.  Be sure to contact me first,
    '   because there is an order to how the characters have to be arranged within a drawing
    '   file.  As fonts are added, I'll either try to publish them somewhere, or you can
    '   just email me and ask if any new ones have been added - I'll send you the updated
    '   function if there are any.
    '
    '   ENJOY!
    '
    '   James Tracy
    '   Sacramento , CA
    '   JamesTracy95820@gmail.com
    '
    '   EXAMPLE USAGE
    '   ^^^^^^^^^^^^^
    '
    '   Dim sdtParms(0 To 10) As Single
    '   Dim idtParms(0 To 4) As Integer
    '
    '   Call the DrawTextInit() function.  Only call this once in your program!
    '   This will call will automatically initialize your sdtParms and idtParms
    '   variables to default values.
    '
    '   DrawTextInit sdtParms, idtParms '   Call only once in your program.
                                        '   Or - you can call it again to re-initialize
                                        '   your sdtParms and idtParms arrays to default
                                        '   values.
    '
    '   Set variables: Short Draw Text Parameters
    '
    '   sdtParms(0) = 20#   '   X location of text.
    '   sdtParms(1) = 280#  '   Y location of text.
    '   sdtParms(2) = 20#   '   X scale of text.
    '   sdtParms(3) = 20#   '   Y scale of text.
    '   sdtParms(4) = 0#    '   Inidividual letter rotation.
    '   sdtParms(5) = 0#    '   Text rotation.
    '   sdtParms(6) = 1#    '   Character spacing.
    '   sdtParms(7) = 1#    '   X overstrike amount (0=none)
    '   sdtParms(8) = 0.1   '   X overstrike increment amount.
    '   sdtParms(9) = 0#    '   Y overstrike amount (0=none)
    '   sdtParms(10) = 0#   '   Y overstrike increment amount.
    '
    '   Set variables: Integer Draw Text Parameters
    '
    '   idtParms(0) = 255   '   RED color of text.
    '   idtParms(1) = 0     '   GREEN color of text.
    '   idtParms(2) = 0     '   BLUE color of text.
    '   idtParms(3) = 0     '   FONT number:
                            '       0   =   Courier
                            '       1   =   JET, my own handwriting.
                            '       2   =   AutoCAD TXT simple font.
    ''''''''''''''''''''''''''''''''''
    '   Do the DrawText() function   '
    ''''''''''''''''''''''''''''''''''
    '   DrawText Me.Picture1, sdtParms, idtParms, "AaBbGgJj 123 !@#$%^&*()_;<>,./?"
    '
    '   idtParms(0) = 0     '   RED color of text.
    '   idtParms(1) = 255   '   GREEN color of text.
    '   idtParms(2) = 0     '   BLUE color of text.
    '   idtParms(3) = 0     '   FONT number
    '   sdtParms(4) = 45#   '   Inidividual letter rotation.
    '   sdtParms(5) = -45#  '   Text rotation.
    '   DrawText Me.Picture1, sdtParms, idtParms, "Testing one two."
    '
    '   idtParms(0) = 0     '   RED color of text.
    '   idtParms(1) = 0     '   GREEN color of text.
    '   idtParms(2) = 255   '   BLUE color of text.
    '   idtParms(3) = 1     '   FONT number
    '   sdtParms(4) = -45#  '   Inidividual letter rotation.
    '   sdtParms(5) = 45#   '   Text rotation.
    '   DrawText Me.Picture1, sdtParms, idtParms, "~`!@#$%^&*()_-+={[}]"
    '
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   FUNCTION PARAMETERS
    '   ^^^^^^^^^^^^^^^^^^^
    '
    '   See EXAMPLE USAGE for all the array parameters.
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim SingleChar As String
    Dim L1 As Integer
    'DumpArray sdtParms(), idtParms(), TextString
    CarriageLocation = 0#
    If Len(TextString) > 0 Then
        For L1 = 1 To Len(TextString)
            SingleChar = Mid(TextString, L1, 1)
            CharNum = dtGetCharNum(SingleChar)
            'm "Processing: " & SingleChar & "(" & Str(CharNum) & ")"
            If CharNum = 255 Then
                MsgBox "Invalid character.  Program aborted", , "PROGRAM ABORTED"
                End
            End If
            If CharNum <> 94 Then       '   NOT a SPACE
                dtProcessChar PicObj, sdtParms, idtParms
            Else
                dtIncrementCarriage sdtParms     '   We have a SPACE.
            End If
        Next L1
    End If
End Sub
Sub DrawTextInit(sdtParms() As Single, idtParms() As Integer)
    Dim Looper1 As Integer
    Dim Looper2 As Integer
    Dim Looper3 As Integer
    FontFileNames(0) = "FONT_COURIER.txt"
    FontFileNames(1) = "FONT_JET.txt"
    FontFileNames(2) = "FONT_TXT.txt"
    '
    '
    '
    sdtParms(0) = 0#        '   X location for text.
    sdtParms(1) = 0#        '   Y location for text.
    sdtParms(2) = 1#        '   X scale.
    sdtParms(3) = 1#        '   Y scale.
    sdtParms(4) = 0#        '   Letter rotation.
    sdtParms(5) = 0#        '   Text rotation.
    sdtParms(6) = 1#        '   Character spacing.
    sdtParms(7) = 0#        '   X overstrike amount.
    sdtParms(8) = 0#        '   X overstrike increment amount.
    sdtParms(9) = 0#        '   Y overstrike amount.
    sdtParms(10) = 0#       '   Y overstrike increment amount.
    '
    idtParms(0) = 255       '   RED color.
    idtParms(1) = 0         '   GREEN color.
    idtParms(2) = 0         '   BLUE color.
    idtParms(3) = 0         '   FONT number.
    '
    '   CHECK THAT FILES EXIST, AND INITIALIZE ARRAY
    '
    For Looper1 = 0 To TOTFONTS - 1
        If Not FileExists(FontFileNames(Looper1)) Then
            MsgBox "The file: " & FontFileNames(Looper1) & " does not exist.  Program aborted.", , "PROGRAM ABORTED"
            End
        Else
            VectorCount(Looper1) = 0#
            For Looper2 = 0 To MAXVECTORS - 1
                For Looper3 = 0 To 3
                    Vectors(Looper1, Looper2, Looper3) = 0#
                Next Looper3
            Next Looper2
        End If
    Next Looper1
    For Looper1 = 0 To TOTFONTS - 1
        For Looper2 = 0 To 93
            VectorBounds(Looper1, Looper2, 0) = 32000       '   LOWER bound
            VectorBounds(Looper1, Looper2, 1) = -32000      '   UPPER BOUND
        Next Looper2
    Next Looper1
    CharSet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890~`!@#$%^&*()_-+={[}]|\:;" & Chr(34) & Chr(39) & "<,>.?/"
    'm "CharSet length (space not in set)= " & Str(Len(CharSet))
    If Len(CharSet) <> 94 Then
        MsgBox "The character set does not contain exactly 94 characters.  Program aborted.", , "PROGRAM ABORTED"
        End
    End If
    dtReadVectors
    dtGetBounds
End Sub
Sub dtIncrementCarriage(sdtParms() As Single)
    CarriageLocation = CarriageLocation + (sdtParms(6) * sdtParms(2))
End Sub
Sub dtDoOffset(sdtParms() As Single)
    X1 = X1 + sdtParms(0)
    Y1 = Y1 + sdtParms(1)
    X2 = X2 + sdtParms(0)
    Y2 = Y2 + sdtParms(1)
End Sub
Sub dtProcessCarriageLocation()
    X1 = X1 + CarriageLocation
    X2 = X2 + CarriageLocation
End Sub
Sub dtMoveHome()
    X1 = X1 - CharNum
    X2 = X2 - CharNum
End Sub
Sub dtScale(sdtParms() As Single)
    X1 = X1 * sdtParms(2)
    Y1 = Y1 * sdtParms(3)
    X2 = X2 * sdtParms(2)
    Y2 = Y2 * sdtParms(3)
    
End Sub
Sub dtRotateLetters(sdtParms() As Single)
    Dim X11 As Single
    Dim Y11 As Single
    sRotateZ sdtParms(4), X1, Y1, 0#, 0#, X11, Y11
    X1 = X11
    Y1 = Y11
    sRotateZ sdtParms(4), X2, Y2, 0#, 0#, X11, Y11
    X2 = X11
    Y2 = Y11
End Sub
Sub dtRotateText(sdtParms() As Single)
    Dim X11 As Single
    Dim Y11 As Single
    sRotateZ sdtParms(5), X1, Y1, sdtParms(0), sdtParms(1), X11, Y11
    X1 = X11
    Y1 = Y11
    sRotateZ sdtParms(5), X2, Y2, sdtParms(0), sdtParms(1), X11, Y11
    X2 = X11
    Y2 = Y11
End Sub
Sub dtProcessChar(PicObj As Object, sdtParms() As Single, idtParms() As Integer)
    Dim LowerBound As Integer
    Dim UpperBound As Integer
    Dim L1  As Integer
    Dim OverStrikeLoop As Single
    If CharNum < 0 Or CharNum > 93 Then
        MsgBox "Character out of range error.  Program aborted!", , "PROGRAM ABORTED"
        End
    End If
    LowerBound = VectorBounds(idtParms(3), CharNum, 0)
    UpperBound = VectorBounds(idtParms(3), CharNum, 1)
    'm "Bounds from " & Str(LowerBound) & " to " & Str(UpperBound)
    For L1 = LowerBound To UpperBound
        X1 = Vectors(idtParms(3), L1, 0)
        Y1 = Vectors(idtParms(3), L1, 1)
        X2 = Vectors(idtParms(3), L1, 2)
        Y2 = Vectors(idtParms(3), L1, 3)
        dtMoveHome
        If sdtParms(4) <> 0# Then dtRotateLetters sdtParms
        dtScale sdtParms
        dtProcessCarriageLocation
        dtDoOffset sdtParms
        If sdtParms(5) <> 0# Then dtRotateText sdtParms
        PicObj.Line (X1, Y1)-(X2, Y2), RGB(idtParms(0), idtParms(1), idtParms(2))
        
        
        
        If sdtParms(7) <> 0# Then       '   X overstrike.
            If sdtParms(8) < sdtParms(7) Then
                For OverStrikeLoop = sdtParms(8) To sdtParms(7) Step sdtParms(8)
                    PicObj.Line (X1 + OverStrikeLoop, Y1)-(X2 + OverStrikeLoop, Y2), RGB(idtParms(0), idtParms(1), idtParms(2))
                Next OverStrikeLoop
            End If
        End If
        
        If sdtParms(9) <> 0# Then       '   Y overstrike.
            If sdtParms(10) < sdtParms(9) Then
                For OverStrikeLoop = sdtParms(10) To sdtParms(9) Step sdtParms(10)
                    PicObj.Line (X1, Y1 + OverStrikeLoop)-(X2, Y2 + OverStrikeLoop), RGB(idtParms(0), idtParms(1), idtParms(2))
                Next OverStrikeLoop
            End If
        End If
        
        
    Next L1
    dtIncrementCarriage sdtParms
End Sub
Sub DumpArray(sdtParms() As Single, idtParms() As Integer, TextString As String)
    Dim OutString As String
    Dim L1, L2, L3, L4 As Integer
    Dim Hand1 As Integer
    Hand1 = FreeFile
    Open "killme1.txt" For Output As #Hand1
    Print #Hand1, "VECTOR COUNT"
    Print #Hand1, "^^^^^^^^^^^^"
    For L1 = 0 To TOTFONTS - 1
        OutString = "Font: " & Str(L1) & " Vector count: " & Str(VectorCount(L1))
        Print #Hand1, OutString
    Next L1
    Print #Hand1, ""
    Print #Hand1, "FONT FILE NAMES"
    Print #Hand1, "^^^^^^^^^^^^^^^"
    For L1 = 0 To TOTFONTS - 1
        Print #Hand1, "FontFileNames(" & Str(L1) & ")=" & FontFileNames(L1)
    Next L1
    Print #Hand1, "VECTOR BOUNDS"
    Print #Hand1, "^^^^^^^^^^^^^"
    For L1 = 0 To TOTFONTS - 1
        For L2 = 0 To 93
                OutString = "Font: " & Str(L1) & " Char: " & Str(L2) & " Start: " & _
                            Str(VectorBounds(L1, L2, 0))
            Print #Hand1, OutString
                OutString = Space(18) & "Stop: " & Str(VectorBounds(L1, L2, 1))
            Print #Hand1, OutString
        Next L2
    Next L1
    Print #Hand1, "VECTORS"
    Print #Hand1, "^^^^^^^"
    For L1 = 0 To TOTFONTS - 1
        Print #Hand1, "FONT # " & Str(L1) & ">--> " & FontFileNames(L1)
        For L2 = 0 To 93
            Print #Hand1, "    CHARACTER # " & Str(L2) & ">--> " & Mid(CharSet, L2 + 1, 1)
            For L3 = VectorBounds(L1, L2, 0) To VectorBounds(L1, L2, 1)
                Print #Hand1, "        Font: " & Str(L1) & "|" & _
                    Vectors(L1, L3, 0) & "|" & _
                    Vectors(L1, L3, 1) & "|" & _
                    Vectors(L1, L3, 2) & "|" & _
                    Vectors(L1, L3, 3)
            Next L3
        Next L2
    Next L1
    '
    '   PARAMETERS PASSED TO DrawText()
    '
    Print #Hand1, ""
    Print #Hand1, "DrawText() Parameters"
    Print #Hand1, "^^^^^^^^^^^^^^^^^^^^^"
    Print #Hand1, ""
    For L1 = 0 To 1
        Print #Hand1, "sdtParms(" & Str(L1) & ")=" & Str(sdtParms(L1))
    Next L1
    Print #Hand1, ""
    For L1 = 0 To 3
        Print #Hand1, "idtParms(" & Str(L1) & ")=" & Str(idtParms(L1))
    Next L1
    Print #Hand1, ""
    Print #Hand1, "TextString=" & TextString
    Close #Hand1
End Sub
Sub dtGetBounds()
    Dim L1, L2  As Integer
    Dim CharNumber As Integer
    For L1 = 0 To TOTFONTS - 1
        For L2 = 0 To VectorCount(L1) - 1
            CharNumber = Int(Vectors(L1, L2, 0))
            VectorBounds(L1, CharNumber, 0) = iMin(L2, VectorBounds(L1, CharNumber, 0))  'LOWER bounds
            VectorBounds(L1, CharNumber, 1) = iMax(L2, VectorBounds(L1, CharNumber, 1))  'UPPER bounds
        Next L2
    Next L1
    '
    '   Make sure there aren't any -1s in the array.
    '   or less than zero, or greater than 93.
    '
    'For L1 = 0 To TOTFONTS - 1
        'For L2 = 0 To 93
            'm "Font # " & Str(L1) & " char: " & Str(L2) & " Starts: " & Str(VectorBounds(L1, L2, 0))
            'm "Font # " & Str(L1) & " char: " & Str(L2) & " Stops:  " & Str(VectorBounds(L1, L2, 1))
        'Next L2
    'Next L1
End Sub
Sub dtReadVectors()
    Dim InString As String
    Dim Hand1 As Integer
    Dim Hand2 As Integer
    Dim Looper As Integer
    Dim Fields() As String
    For Looper = 0 To TOTFONTS - 1
        Hand1 = FreeFile
        Open FontFileNames(Looper) For Input As #Hand1
        Do While Not EOF(Hand1)
            Line Input #Hand1, InString
            Fields = Split(InString, "|")
            VectorCount(Looper) = VectorCount(Looper) + 1
            If VectorCount(Looper) >= MAXVECTORS Then
                MsgBox "Maximum number of vectors has been exceeded.  Program aborted", , "ABNORMAL END"
                End
            End If
            Vectors(Looper, (VectorCount(Looper) - 1), 0) = CSng(Fields(0))
            Vectors(Looper, (VectorCount(Looper) - 1), 1) = CSng(Fields(1))
            Vectors(Looper, (VectorCount(Looper) - 1), 2) = CSng(Fields(2))
            Vectors(Looper, (VectorCount(Looper) - 1), 3) = CSng(Fields(3))
        Loop
        Close #Hand1
        'm "For font: " & FontFileNames(Looper) & " there are: " & Str(VectorCount(Looper))
    Next Looper
End Sub
Function dtGetCharNum(CharIn As String) As Integer
    Dim ReturnVal As Integer
    Dim Loop1 As Integer
    ReturnVal = 255  '   Invalid character!
    If CharIn >= "a" And CharIn <= "z" Then
        ReturnVal = Asc(CharIn) - 70 - 1
    Else
        If CharIn >= "A" And CharIn <= "Z" Then
            ReturnVal = Asc(CharIn) - 64 - 1
        Else
            If CharIn >= "1" And CharIn <= "9" Then
                ReturnVal = Asc(CharIn) + 4 - 1
            Else
                If CharIn = " " Then
                    ReturnVal = 94
                Else
                    For Loop1 = 62 To 94
                        If CharIn = Mid(CharSet, Loop1, 1) Then
                            ReturnVal = Loop1 - 1
                        End If
                    Next Loop1
                End If
            End If
        End If
    End If
    dtGetCharNum = ReturnVal
End Function





