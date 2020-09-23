VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DrawText() Function Tester"
   ClientHeight    =   9465
   ClientLeft      =   2370
   ClientTop       =   810
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   11730
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   3480
      TabIndex        =   3
      Top             =   0
      Width           =   8175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DRAW TEXT"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   8535
      Left            =   120
      ScaleHeight     =   565
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   765
      TabIndex        =   0
      Top             =   840
      Width           =   11535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command2_Click()
    m "DrawText() function called"
    '
    '   Define DrawText() function variables
    '
    Dim sdtParms(0 To 10) As Single
    Dim idtParms(0 To 4) As Integer
    
    '
    '   Call the DrawTextInit() function.  Only call this once in your program!
    '   This will call will automatically initialize your sdtParms and idtParms
    '   variables to default values.
    '
    DrawTextInit sdtParms, idtParms     '   Call only once in your program.
                                        '   Or - you can call it again to re-initialize
                                        '   your sdtParms and idtParms arrays to default
                                        '   values.
    '
    '   Set variables Short Draw Text Parameters
    '
    sdtParms(0) = 20#   '   X location of text.
    sdtParms(1) = 280#  '   Y location of text.
    sdtParms(2) = 20#   '   X scale of text.
    sdtParms(3) = 20#   '   Y scale of text.
    sdtParms(4) = 0#    '   Inidividual letter rotation.
    sdtParms(5) = 0#    '   Text rotation.
    sdtParms(6) = 1#    '   Character spacing.
    sdtParms(7) = 1#    '   X overstrike amount (0=none)
    sdtParms(8) = 0.1   '   X overstrike increment amount.
    sdtParms(9) = 0#    '   Y overstrike amount (0=none)
    sdtParms(10) = 0#   '   Y overstrike increment amount.
    '
    '   Set variables Integer Draw Text Parameters
    '
    idtParms(0) = 255   '   RED color of text.
    idtParms(1) = 0     '   GREEN color of text.
    idtParms(2) = 0     '   BLUE color of text.
    idtParms(3) = 0     '   FONT number:
                        '       0   =   Courier
                        '       1   =   JET, my own handwriting.
                        '       2   =   AutoCAD TXT simple font.
                                  
                                        
    ''''''''''''''''''''''''''''''''''
    '   Do the DrawText() function   '
    ''''''''''''''''''''''''''''''''''
    DrawText Me.Picture1, sdtParms, idtParms, "AaBbGgJj 123 !@#$%^&*()_;<>,./?"
    idtParms(0) = 0     '   RED color of text.
    idtParms(1) = 255   '   GREEN color of text.
    idtParms(2) = 0     '   BLUE color of text.
    idtParms(3) = 0     '   FONT number
    sdtParms(4) = 45#   '   Inidividual letter rotation.
    sdtParms(5) = -45#  '   Text rotation.
    '
    DrawText Me.Picture1, sdtParms, idtParms, "Testing one two."
    idtParms(0) = 0     '   RED color of text.
    idtParms(1) = 0     '   GREEN color of text.
    idtParms(2) = 255   '   BLUE color of text.
    idtParms(3) = 1     '   FONT number
    sdtParms(4) = -45#  '   Inidividual letter rotation.
    sdtParms(5) = 45#   '   Text rotation.
    DrawText Me.Picture1, sdtParms, idtParms, "~`!@#$%^&*()_-+={[}]"
    '
    m "DrawFunction() returned"
    m "PROGRAM ENDED..."
End Sub
Private Sub Form_Load()
    m "PROGRAM STARTED..."
    Me.Picture1.AutoRedraw = True   '   Automatically update the picture box if changed.
    FlipScale Me.Picture1           '   Flip the scale so the coordinates are in the lower left corner.
    DrawGrid Me.Picture1            '   Draw a grid on the picture box.
    '
    '   Sent the line type, if needed
    '
    Me.Picture1.DrawStyle = vbSolid     '   Set the line type:
                                        '       vbSolid
                                        '       vbDash
                                        '       vbDot
                                        '       vbDashDot
                                        '       vbDashDotDot
                                        '       vbInvisible
                                        '       vbInsideSolid
End Sub
Private Sub Command1_Click()
    End
End Sub
Sub m(Message As String)
    '
    '   Requires the form name (Form1 or other name)
    '   and the list box name (List1 or other name)
    '
    Form1.List1.AddItem (Time() & ">-->" & Message)
    Form1.List1.Refresh
End Sub

