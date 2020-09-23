VERSION 5.00
Begin VB.Form frmLife 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artificial Life"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   DrawWidth       =   2
   FillColor       =   &H008080FF&
   ForeColor       =   &H000000FF&
   Icon            =   "frmLife.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   406
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblEvolve 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   17
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label lblAverageEvolveLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "Average Ev. Level :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label lblMinEv 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3495
      TabIndex        =   15
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label lblMinEvolveLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "Min Evolve Level :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label lblFood 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   13
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lblAFood 
      BackStyle       =   0  'Transparent
      Caption         =   "Food Amount :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label lblEaterNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Ant Eater Number : 0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label lblAProp 
      BackStyle       =   0  'Transparent
      Caption         =   "Average Pop. Level :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label lblALife 
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Average Life Span :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label lblASpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "Average Speed :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Line lnBreak2 
      BorderColor     =   &H00FFFFFF&
      X1              =   264
      X2              =   264
      Y1              =   408
      Y2              =   480
   End
   Begin VB.Label lblHProp 
      BackStyle       =   0  'Transparent
      Caption         =   "Lowest Prop. Level :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label lblLLife 
      BackStyle       =   0  'Transparent
      Caption         =   "Longest Life :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label lblFast 
      BackStyle       =   0  'Transparent
      Caption         =   "Fastest :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Line lnBreak1 
      BorderColor     =   &H00FFFFFF&
      X1              =   136
      X2              =   136
      Y1              =   408
      Y2              =   480
   End
   Begin VB.Label lblTick 
      BackStyle       =   0  'Transparent
      Caption         =   "Ticks : "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label lblProp 
      BackStyle       =   0  'Transparent
      Caption         =   "Propegation Rate :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label lblDeath 
      BackStyle       =   0  'Transparent
      Caption         =   "Death Rate : "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label lblGen 
      BackStyle       =   0  'Transparent
      Caption         =   "Generations : "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "Number :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6120
      Width           =   1935
   End
End
Attribute VB_Name = "frmLife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program was created simply because I wanted to see what
'happened when random choices were taken within ceratin rules.
'The averages are taken because I wanted to see how they changed
'overall after a certain period of time. No two runnings of this
'program are exactly the same. Each time you run this program, it will
'be slightly different.
'On a technical note, this program contains a nice 2D array searching
'technique, even if i do say so myself :)
'
'This program was updated 16/11/2001 with the (very considreable)
'help of Pei Pei. If not for her valuable advice and insightfull questions
'and suggestions, this update would not have happened.
'Thank you.
'
'Eric
'DiskJunky@hotmail.com

Private Declare Function GetTickCount Lib "kernel32" () As Long 'time returned is in nanoseconds (sec/1000)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)    'time passed is in milliseconds (sec/100)

Private Enum AverageEnum
    AvSpeed = 0
    AvLifeSpan = 1
    AvPopLevel = 2
    AvEvolve = 3
End Enum

Public Sub DrawBlank()
'this draws a blank square

Dim X As Integer
Dim Y As Integer

For X = 0 To GridSize - 1
    For Y = 0 To GridSize - 1
        Grid(X, Y) = HereEmpty
    Next Y
Next X

FoodAmount = 0
End Sub

Public Sub CreateFood(Optional Multiply As Long, Optional Start As Boolean = False)
'This randomly places food in different places in the grid

Dim Upperbound As Integer
Const Lowerbound = 0

Dim Counter As Integer
Dim X As Integer
Dim Y As Integer
Dim MyFoodRate As Long
Dim MyUpperbound As Integer
Dim MyLowerbound As Integer

Upperbound = GridSize - 1

If Multiply <> 0 Then
    MyFoodRate = FoodRate * Multiply
Else
    MyUpperbound = FoodRate + 3
    MyLowerbound = 0
    MyFoodRate = GetRndInt(MyLowerbound, MyUpperbound) '((MyUpperbound - MyLowerbound + 1) * Rnd + MyLowerbound)
End If

'if the amount of food in grid is greater than 3/4 then, don't create more
If (FoodAmount + MyFoodRate) > ((GridSize ^ 2) * 0.75) Then '- ((GridSize ^ 2) / 4)
    Exit Sub
Else
    FoodAmount = FoodAmount + MyFoodRate
End If

'create "FoodRate" bits of food
For Counter = 1 To MyFoodRate
    Do
        'randomize
        X = GetRndInt(Lowerbound, Upperbound) '((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
        Y = GetRndInt(Lowerbound, Upperbound) '((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
    Loop Until (Grid(X, Y) <> WorldDetails.Ant) And (Grid(X, Y) <> WorldDetails.Eater)
    
    If Not Start Then
        'draw the dot if not creating food for the start.
        Call frmLife.DrawDot(X, Y, vbGreen)
    End If
    
    Grid(X, Y) = Food
Next Counter

If Start Then
    Call DrawFrame
End If
End Sub

Public Sub FlipVal(ByRef Val1 As Integer, ByRef Val2 As Integer)
'This function swaps the two value

Dim Temp As Integer

Temp = Val1
Val1 = Val2
Val2 = Temp
End Sub

Public Function RndBool() As Integer
'this function returns either 1 or -1

Dim RndVal As Integer

'randomize
'RndVal = Int(Rnd)

If Rnd < 0.5 Then
    RndVal = -1
Else
    RndVal = 1
End If

RndBool = RndVal
End Function

Public Function SearchForFood(XCo As Integer, YCo As Integer, Speed As Integer, MyDirection As Direction, Optional FoodType As Integer) As Boolean
'This will look for a piece of food within half the speed (distance)
'radius from the co-ordinates given starting from the point outwards,
'in a random direction.

Const X = 0
Const Y = 1

Dim CounterX As Integer
Dim CounterY As Integer
Dim Counter As Integer
Dim Low(2) As Integer
Dim High(2) As Integer
Dim Found(2) As Integer
Dim TheFood As WorldDetails
Dim Offset(2) As Integer
Dim Step(2) As Integer
Dim Start(2) As Integer
Dim Finish(2) As Integer

If FoodType <> 0 Then
    'look for ants
    TheFood = WorldDetails.Ant
Else
    TheFood = WorldDetails.Food
End If

SearchForFood = False
If Speed < 1 Then
    Exit Function
End If

'set the range in which to look for food
High(X) = XCo + Speed
Low(X) = XCo - Speed
High(Y) = YCo + Speed
Low(Y) = YCo - Speed

'half the search range depending on the direction
'RndBool is used to start the search from a random direction in the
'given search area
MyDirection = MyDirection Mod 4
Select Case MyDirection
Case Up
    'don't search the lower half
    Low(Y) = YCo - 1
    
    Step(X) = RndBool
    Step(Y) = -1
Case Right
    'don't search the left half
    Low(X) = XCo - 1
    
    Step(X) = 1
    Step(Y) = RndBool
Case Down
    'don't search the top half
    High(Y) = YCo + 1
    
    Step(X) = RndBool '-1
    Step(Y) = 1 'RndBool
Case Direction.Left
    'don't search the right half
    High(X) = XCo + 1
    
    Step(X) = -1 'RndBool
    Step(Y) = RndBool '1
End Select

'start searching the area directly ahead
Offset(X) = ((High(Y) - Low(Y)) / 2) * Step(X)
Offset(Y) = ((High(X) - Low(X)) / 2) * Step(Y)

'randomize search direction and set search values
For Counter = 0 To 1
    If Step(Counter) = -1 Then
        Call FlipVal(Low(Counter), High(Counter))
        'Stop
    End If
    
    Start(Counter) = (Low(Counter) + Offset(Counter)) ' * Step(Counter))
    Finish(Counter) = (High(Counter) + Offset(Counter))
Next Counter

For CounterX = Start(X) To Finish(X) Step Step(X)
    For CounterY = Start(Y) To Finish(Y) Step Step(Y)
        'check values to see if they go past the edge and change them
        'accordingly.
        Found(X) = CheckRange(0, GridSize, OffsetRange(High(X), Low(X), CounterX, Offset(X)))
        Found(Y) = CheckRange(0, GridSize, OffsetRange(High(Y), Low(Y), CounterY, Offset(Y)))
        
        'look for food (1 = AntFood, 2 = Anteaters food (ants))
        If Grid(Found(X), Found(Y)) = TheFood Then
            XCo = Found(X)
            YCo = Found(Y)
            SearchForFood = True
            Exit Function
        End If
    Next CounterY
Next CounterX
End Function

Public Function OffsetRange(Max As Integer, Min As Integer, Value As Integer, Offset As Integer) As Integer
'this function will add Offset to Value. Should the result exceed
'Max, then the result continues from Min.
'eg, assume the function was called with;
'OffsetRange(50,10,45,10)
'the Result = 15 because 45+10=55 but Max = 50. The excess 5 is
'added to Min (10) to give the result.

Dim Result As Integer

If Min > Max Then
    'swap values
    Call FlipVal(Max, Min)
End If

'if a positive number was passed, then
Result = (Value + Offset) Mod (Max + 1)
If Result < Value Then
    'add to min
    Result = CheckRange(Min, Max, Result) '((Min - Result) + Min)
End If

OffsetRange = Result
End Function

Public Sub DrawDot(ByVal X As Integer, ByVal Y As Integer, ByVal Colour As Long)
'this draws a dot at the specified coordinates in the given colour

Static LastWidth As Byte

Dim Offset As Integer

'if the display is pause, then exit
If PauseResult > 0 Then
    Exit Sub
End If

Select Case Colour
Case vbGreen
    If LastWidth <> 1 Then
        frmLife.DrawWidth = 1
        LastWidth = 1
        frmLife.ForeColor = Colour
    End If
Case Else
    If LastWidth <> 2 Then
        frmLife.DrawWidth = 2
        LastWidth = 2
    End If
    frmLife.ForeColor = Colour
End Select

'center the display horizontally
Offset = (frmLife.ScaleWidth / 2) - GridSize

'Select Case Colour
'Case vbWhite
'    frmLife.Circle ((X + 0) * 2 * Screen.TwipsPerPixelX, (Y + 0) * 2 * Screen.TwipsPerPixelY), 0, Colour
'Case Else '* Screen.TwipsPerPixelX
    frmLife.PSet (Offset + (X * 2), Y * 2)
'End Select
DoEvents
End Sub

Private Sub Form_Activate()
'This is where all the work is done. The program really runs from here.
'I put this as a While loop instead of a timer because it's quicker :)

Const MilliSecDelay = 1 'the delay to pause execution of code for in milliseconds (sec/100)

Static TickCount As Double
Static StartOnly As Boolean
Static LastEaterNum As Integer

Dim StartTick As Long
Dim Counter As Integer
Dim SleepDelay As Long
'Dim TotalDead As Long
'Dim AntsLastTick As Long
'Dim TotalBorn As Long

While True 'NumOfAnts > 0
    If Not StartOnly Then
        Call DrawBlank
        Call StartingAnt
        Call DrawDot(Ants(1).XPos, Ants(1).YPos, vbWhite)
        StartOnly = True
    End If
    
    Call frmLife.CreateFood
    
    'move each ant
    For Counter = 1 To NumOfAnts
        StartTick = GetTickCount
        If Counter > NumOfAnts Then
            Exit For
        End If
        Call MoveAnt(Counter)
        DoEvents
                    
        If Not PauseResult Then
            'if the display is not paused, then limit the framerate
            SleepDelay = (MilliSecDelay - ((GetTickCount - StartTick) / 10))
            If SleepDelay < 0 Then
                SleepDelay = 0
            End If
            Call Sleep(SleepDelay)
        End If
    Next Counter
    
    'if there are no eaters and the number of ants is available, then
    'create the ant eaters
    'If (NumOfEaters = 0) And (NumOfAnts > IntroduceEaterAt) Then
    '    Call StartingEater
    'Else
    '    If ((NumOfAnts * 3) < NumOfEaters) And ((TickNum Mod 10) = 0) Then 'And (NumOfEaters > 1)
    '        'reduce the number of eaters if there are more eaters than
    '        'ants, every tehnth tick
    '        Call KillEater(1)
    '    End If
    '
        'move the eaters
        For Counter = 1 To NumOfEaters
            StartTick = GetTickCount
            If Counter > NumOfEaters Then
                Exit For
            End If
            Call MoveEater(Counter)
            DoEvents
            
            If Not PauseResult Then
                'if the display is not paused, then limit the framerate
                'Debug.Print (MilliSecDelay - ((GetTickCount - StartTick) / 10))
                SleepDelay = (MilliSecDelay - ((GetTickCount - StartTick) / 10))
                If SleepDelay < 0 Then
                    SleepDelay = 0
                End If
                Call Sleep(SleepDelay)
            End If
        Next Counter
    'End If
    
    'calculate averages
    If (NumOfAnts > 0) And (PauseResult = 0) Then
        lblASpeed.Caption = "Average Speed : " & Format(GetAverage(AvSpeed), "0.0") '(TotalSpeed / NumOfAnts)
        lblALife.Caption = "Average Life Span : " & Format(GetAverage(AvLifeSpan), "0.0") '(TotalLife / NumOfAnts)
        lblAProp.Caption = "Average Prop. Rate : " & Format(GetAverage(AvPopLevel), "0.0") '(TotalProp / NumOfAnts)
        lblEvolve.Caption = Format(GetAverage(AvEvolve), "0.0")
    End If
    
    'show food amount
    If FoodAmount <> Val(lblFood.Caption) Then
        lblFood.Caption = GetFoodAmount 'FoodAmount
    End If
    
    'if no creatures left then,
    If (NumOfAnts = 0) And (NumOfEaters = 0) Then
        'if no ants, then create a new one
        If (NumOfAnts = 0) And (NumOfEaters > 1) Then
            Call StartingAnt
        Else
            'no creatures left - start again
            StartOnly = False
        End If
    End If
    
    'update the eater number display
    If (LastEaterNum <> NumOfEaters) Then 'And (PauseResult = 0)
        'update the display
        LastEaterNum = NumOfEaters
        lblEaterNum.Caption = "Ant Eater Number : " & NumOfEaters
    End If
    
    'update display if able
    lblCount = "Population : " & NumOfAnts
    
    If PauseResult <> 0 Then
        'one less tick to wait
        PauseResult = PauseResult - 1
        If PauseResult = 0 Then
            'redraw the entire frame
            Call DrawFrame
        End If
    End If
    
    TickCount = TickCount + 1
    lblTick.Caption = "Ticks : " & Format(TickCount, "0")
Wend

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'if the key pressed was between 1 and 9, then pause exexution
'of the display for key * 100 ticks

If (KeyAscii >= vbKey0) And (KeyAscii <= vbKey9) Then
    'if a number key was hit, then pause the display for that number
    'times 100 ticks
    PauseResult = Val(Chr(KeyAscii)) * 100
End If

If KeyAscii = 27 Then
    'if the escape key was pressed then create more food
    Call CreateFood((GridSize ^ 2) * 0.75)
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If NumOfAnts = 0 Then
        Call DrawBlank
        Call StartingAnt
        Call frmLife.DrawDot(Ants(1).XPos, Ants(1).YPos, vbWhite)
        StartOnly = True
    Else
        'create one third of food
        Call frmLife.CreateFood(((GridSize ^ 2) * 0.33) / FoodRate)
    End If
Else
    'vbRightButton - stop drawing results for a certain number of
    'ticks
    PauseResult = 100
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Public Sub DrawFrame()
'This will draw everything in the entire grid

Dim X As Integer
Dim Y As Integer
Dim Colour As Long

'draw a blank square
frmLife.Cls

For X = 1 To GridSize
    For Y = 1 To GridSize
        'draw the appropiate dot
        Select Case Grid(X, Y)
        Case WorldDetails.HereEmpty
            Colour = vbBlack
        Case WorldDetails.Food
            Colour = vbGreen
        Case WorldDetails.Ant
            Colour = vbWhite
        Case WorldDetails.Eater
            Colour = vbCyan
        End Select
        
        If Colour <> vbBlack Then
            Call DrawDot(X, Y, Colour)
        End If
    Next Y
Next X
End Sub

Private Function GetAverage(Value As AverageEnum) As Single
'This will calculate the average speed of all the ants

Dim Counter As Integer
Dim Total As Long
Dim Result As Single
Dim Num As Integer

For Counter = 0 To NumOfAnts
    'select the value to get average of
    Select Case Value
    Case AvSpeed
        Num = Ants(Counter).Speed
    Case AvLifeSpan
        Num = Ants(Counter).KillLevel
    Case AvPopLevel
        Num = Ants(Counter).FoodToSplit
    Case AvEvolve
        Num = Ants(Counter).EvolveLevel
    End Select
    
    'add to total
    Total = Total + Num
Next Counter

'find average
Result = Total / NumOfAnts

'return result
GetAverage = Result
End Function

Private Function GetFoodAmount() As Long
'This function returns the amount of food in the grid

Dim Row As Integer
Dim Col As Integer
Dim Num As Long

For Row = 0 To GridSize
    For Col = 0 To GridSize
        If Grid(Row, Col) = Food Then
            'found base food
            Num = Num + 1
        End If
    Next Col
Next Row

'this is an accurate result, so set the FoodAmount variable
FoodAmount = Result

GetFoodAmount = Num
End Function
