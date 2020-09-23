VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   DrawMode        =   5  'Not Copy Pen
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type V
'timer variables
'the interval of the timer
    Interval As Double
'the last tick the timer had
    LastTick As Double
    
'Player Point Variables
'the X position of the specifyed point
    X As Double
'the Y position of the specifyed point
    Y As Double
'the Height of the point
    Height As Double
'the width of the point
    Width As Double
    
'Bullet Variables
'The direction the bullet is going
    Direction As Double
'If the bullet has been shot or not
    Shot As Boolean
'Randomly chose the color of the bullet using RGB colors
    R As Double
    G As Double
    B As Double
'if the bullet is overlapping the opponents body point
    Overlap As Boolean
End Type


Private Type Points
'Player Points
'The Main Body Point
    Bodypoint As V
'The Point that Declares Direction
    DirectionPoint As V
'The Points that the bullets are
    BulletPoint(100) As V
    
'Variables that have to do with all points
'The speed that the Points are moveing
    Speed As Double
'the size of all the points
    Size As Double
'How much health you have
    Health As Double
'The rotation value of sin and cos
    Rotate As Double

'Bullet array
'The Counter for the bullet array
    Count As Double
'The number of bullets allowed in the game
    NumberOfBullets As Double
End Type

'The timer that controls speed of movement, movement, size, detection, drawings _
weapon power.
Dim Timer1 As V

'The timer that controls the fire rate of the wepon and marks which direction the _
bullet is going at what time.
Dim Timer2 As V

'Player1's Points
Dim P1 As Points

'Player2's Points
Dim P2 As Points

'The ticker for my timer
Private Declare Function GetTickCount Lib "kernel32" () As Long

'Calls GetAsyncKeyState to make keys easer to use
Private Declare Function GetAsyncKeyState Lib "User32" (ByVal vkey As Long) As Long

'this sub activates the timers as soon as the game starts
Private Sub Form_Activate()
'Sets timer1's interval as 100 ticks per second
    Timer1.Interval = 10
'Sets tiemr2's interval as 10 ticks per second
    Timer2.Interval = 100
    
'Starts a do doevents loop
    Do
    'Makes sure that the program does every thing befor it loops with out crashing the _
    program
        DoEvents
    'everything that starts with a . has to do with timer1
        With Timer1
        'Creates the timer Interval by subtracting the current tick by the last _
        tick to make it = to the interval
            If GetTickCount - .LastTick >= .Interval Then
            'last tick becomes curent so that the it can be subtracted by _
            the current tick to = the interval
                .LastTick = GetTickCount
            'if you press the escape key then the program ends
                If GetAsyncKeyState(vbKeyEscape) Then End
            'calls the sub timer1_timer.  the sub timer1_timer holds the main _
            controls for the players
                Timer1_Timer
            'constantly calls a variable so that that variable cannot = anything _
            else other that what it =s
                Constants
            'Draws the players
                Draw
            'Detects if the bullet is overlaping a body point
                Detection
            End If
        End With
        
        'same idea as timer1
        With Timer2
            If GetTickCount - .LastTick >= .Interval Then
                .LastTick = GetTickCount
            'Controls the speed of the bullet. Since the Interval is 100 the _
            bullets will only shoot 10 times every secont
                Timer2_Timer
            End If
        End With
    Loop
End Sub

'this sub has all of the things that are allways true
Sub Constants()
'everything in this with statement that starts with a . has to do with Player1
    With P1
    'States that the maximum number of bullets that can be shown on the screen _
    is 100 for player 1
        .NumberOfBullets = 100
    'Everything in this with statement that starts with a . has to do with _
    player1's bodypoint
        With .Bodypoint
        'states that player1's bodypoint's widht is 17 pixels * its size
            .Width = (17 * P1.Size)
        'states that player1's bodypoint's height is 17 pixels * its size
            .Height = (17 * P1.Size)
        End With

    'Everything in this with statement that starts with a . has to do with _
    player1's DirectionPoint
        With .DirectionPoint
        'states that player1's Directionpoint's width is 9 pixels * its size
            .Width = (9 * P1.Size)
        'states that player1's Directionpoint's height is 9 pixels * its size
            .Height = (9 * P1.Size)
        End With
        
    'A For Next loop that calls all bullets as X
        For X = 0 To .NumberOfBullets
        'Everything in this with statement that starts with a . has to do with _
        player1's bulletpoint's 0 to 100
            With .BulletPoint(X)
            'states that player1's bulletpoint's 0 to 100 width is _
            9 pixels * its size
                .Width = (9 * P1.Size)
            'states that player1's bulletpoint's 0 to 100 height is _
            9 pixels * its size
                .Height = (9 * P1.Size)
            End With
        Next
    End With
    
'Everything stated in this with statement is the same as player1's statement _
except it is for player2
    With P2
        .NumberOfBullets = P1.NumberOfBullets

        With .Bodypoint
            .Width = (17 * P2.Size)
            .Height = (17 * P2.Size)
        End With
        
        With .DirectionPoint
            .Width = (9 * P2.Size)
            .Height = (9 * P2.Size)
        End With
                
        For X = 0 To .NumberOfBullets
            With .BulletPoint(X)
                .Width = (9 * P2.Size)
                .Height = (9 * P2.Size)
            End With
        Next
    End With
End Sub

'this sub creates the controls
Sub Timer1_Timer()
'Every thing in this with statement that starts with . has something to do _
with player1
    With P1
    'A For Next loop that calls all bullets as X
        For X = 0 To P1.NumberOfBullets
        'Everything has to do with p1.bulletpoint(x) if it has a . infront of it
            With .BulletPoint(X)
            'if the bullet is visable make it go the direction that it has been _
            assigned
                If .Shot = True Then
                'make the .x postion of bulletpoint(x) go the direction it was _
                given by sin * the declared speed * 3
                    .X = (.X + Sin(.Direction) * P1.Speed * 3)
                'make the .y postion of the bulletpoint(x) go the direction it _
                was given by cos * the declared speed * 3
                    .Y = (.Y + Cos(.Direction) * P1.Speed * 3)
                End If
            End With
        Next
        
    'if you press the a button then rotate player1 counter clock wise
        If GetAsyncKeyState(vbKeyA) Then P1.Rotate = P1.Rotate + 44 * .Speed * 2
        
    'if you press the d button then rotate player1 clock wise
        If GetAsyncKeyState(vbKeyD) Then P1.Rotate = P1.Rotate - 44 * .Speed * 2
        
    'everything within the next with statement that starts with a . has to _
    do with p1.bodypoint
        With .Bodypoint
    'if you press the w button then the following happens
        If GetAsyncKeyState(vbKeyW) Then
        'bodypoint goes the direction of x that sin has given it * the speed that _
        has been declared
            .X = (.X + (Sin(P1.Rotate)) * P1.Speed)
        'bodypoint goes the direction of Y that sin has given it * the speed that _
        has been declared
            .Y = (.Y + (Cos(P1.Rotate)) * P1.Speed)
        End If
        If GetAsyncKeyState(vbKeyS) Then
        'bodypoint goes the opposite direction of x that sin has given it * the _
        speed that has been declared
            .X = (.X - (Sin(P1.Rotate)) * P1.Speed)
        'bodypoint goes the opposite direction of y that sin has given it * the _
        speed that has been declared
            .Y = (.Y - (Cos(P1.Rotate)) * P1.Speed)
        End If
    
        'this is where the directionpoint finds its direction and tells _
        the bullet and body points what direction they are pointing
        '\/'
        'p1.directionpoint's x postion = sin(multiple of 44)* player1's _
        bodypoint's width * the distance that it is from the body point + the _
        bodypoint's x's midpoint - the direcionpoint's x's midpoint
            P1.DirectionPoint.X = ((Sin(P1.Rotate) * .Width * 1 / 3) _
            + .X + .Width / 2 - P1.DirectionPoint.Width / 2)
        'the same thing except it uses p1's y's and height's and cos
            P1.DirectionPoint.Y = ((Cos(P1.Rotate) * .Height * 1 / 3) _
            + .Y + .Height / 2 - P1.DirectionPoint.Height / 2)
        End With
    End With
    
'the same thing as above except for player2
    With P2
        For X = 0 To P1.NumberOfBullets
            With .BulletPoint(X)
                If .Shot = True Then
                    .Y = .Y + Cos(.Direction) * P2.Speed * 3
                    .X = .X + Sin(.Direction) * P2.Speed * 3
                End If
            End With
        Next
        
        If GetAsyncKeyState(vbKeyNumpad4) Then P2.Rotate = P2.Rotate + 44 * _
        .Speed * 2
            
        If GetAsyncKeyState(vbKeyNumpad6) Then P2.Rotate = P2.Rotate - 44 * _
        .Speed * 2

        If GetAsyncKeyState(vbKeyNumpad8) Then
          .Bodypoint.X = .Bodypoint.X + (Sin(P2.Rotate)) * .Speed
          .Bodypoint.Y = .Bodypoint.Y + (Cos(P2.Rotate)) * .Speed
        End If
        If GetAsyncKeyState(vbKeyNumpad5) Then
          .Bodypoint.X = .Bodypoint.X - (Sin(P2.Rotate)) * .Speed
          .Bodypoint.Y = .Bodypoint.Y - (Cos(P2.Rotate)) * .Speed
        End If
    
        With .Bodypoint
            P2.DirectionPoint.X = (Sin(P2.Rotate) * .Width * 1 / 3) + .X + .Width / 2 - P2.DirectionPoint.Width / 2
            P2.DirectionPoint.Y = (Cos(P2.Rotate) * .Height * 1 / 3) + .Y + .Height / 2 - P2.DirectionPoint.Height / 2
        End With
    End With
End Sub

'this sub draws the lines and text directly on to the form with _
out any flickering. I didn 't feel like using bitblt this time
Sub Draw()
'everything in in this with statement has to do with form1
    With Form1
    'states that autoredraw is true
        .AutoRedraw = True
    'clears the screen so that only the shape you want is on the screen _
    if you want to see something cool disable this line of code
        .Cls
    'states that the drawstyle is solid
        .DrawStyle = 0
    'states that the drawmode is copy pen
        .DrawMode = 13
    'the color for the instruction text
        .ForeColor = RGB(255, 255, 255)
    'the instructions for how to play the game
        Form1.Print "To control player1(blue cannon) you use W, S, A, and D to go forward, backward, counter clockwise,"
        Form1.Print "or clockwise respectively. To control player2 use the numpad 8 ,5 ,4 ,and 6 to go forward, backward,"
        Form1.Print "counter clockwise, or clockwise respectively.  To fire with player1 use the space bar.  To fire with"
        Form1.Print "player2 use enter.  Press Escape key to quit"
    'states that the forecolor for the next line is blue
        .ForeColor = RGB(0, 0, 255)
    'prints the health of player1 directly on to the form in blue text
        Form1.Print P1.Health
    'states that the forecolor for the next line is red
        .ForeColor = RGB(255, 0, 0)
    'prints the health of player2 directly on to the form in red text
        Form1.Print P2.Health
    End With
    
'Everything in this with statement that starts with . has to do with p1
    With P1
    'Everything in this with statement that starts with . _
    has to do with p1.bodypoint
        With .Bodypoint
        'states that the drawwith is 17 pixels in diameter * player1's current size
            Form1.DrawWidth = 17 * P1.Size
        'states that the color for the next line will be white
            Form1.ForeColor = RGB(255, 255, 255)
        'creats a dot to be recognized  as the body point that is white
            Form1.Line (.X + .Width / 2, .Y + .Height / 2)-(.X + .Width / 2, .Y + .Height / 2)
        End With
        
    'states that the drawwidth will be 9 pixels in diameter * player1's _
    current size
        Form1.DrawWidth = 9 * P1.Size
    'calls all of the bullets (x)
        For X = 0 To .NumberOfBullets
        'everything that starts with a . in this with statement will be a part of _
        player1.bulletpoint(x)
            With .BulletPoint(X)
            'if the bullet has been fired then draw it
                If .Shot = True Then
                'make the color of the bullet ramdom depending on the current _
                color chosen for the current bullet
                    Form1.ForeColor = RGB(.R, .G, .B)
                'Draws the bullets
                    Form1.Line (.X + .Height / 2, .Y + .Height / 2)-(.X + .Height / 2, .Y + .Height / 2)
                End If
            End With
        Next
        
    'everything in this with statement that starts with a . has to do with _
    p1.directionpoint
        With .DirectionPoint
        'sets the color for the next line to blue
            Form1.ForeColor = RGB(0, 0, 255)
        'creates the line for the direction point
            Form1.Line (P1.Bodypoint.X + P1.Bodypoint.Height / 2, P1.Bodypoint.Y + P1.Bodypoint.Height / 2)-(.X + .Height / 2, .Y + .Height / 2)
        End With
    End With
    
'everything in this with statement has alredy been done in the last statment _
except that this is for player2
    With P2
        With .Bodypoint
            Form1.DrawWidth = 17 * P2.Size
            Form1.ForeColor = RGB(255, 255, 255)
            Form1.Line (.X + .Width / 2, .Y + .Height / 2)-(.X + .Width / 2, .Y + .Height / 2)
        End With
        Form1.DrawWidth = 9 * P2.Size
        For X = 0 To .NumberOfBullets
            With .BulletPoint(X)
                If .Shot = True Then
                    Form1.ForeColor = RGB(.R, .G, .B)
                    Form1.Line (.X + .Height / 2, .Y + .Height / 2)-(.X + .Height / 2, .Y + .Height / 2)
                End If
            End With
        Next
        With .DirectionPoint
            Form1.ForeColor = RGB(255, 0, 0)
            Form1.Line (P2.Bodypoint.X + P2.Bodypoint.Height / 2, P2.Bodypoint.Y + P2.Bodypoint.Height / 2)-(.X + .Height / 2, .Y + .Height / 2)
        End With
    End With
    Form1.Refresh
End Sub

'things that need to be loaded in startup
Private Sub Form_Load()
'Clears all of the keys
    For X = 0 To 255
        If GetAsyncKeyState(X) Then
        End If
    Next
'sets player1's speed
    P1.Speed = 3
'sets player2's speed
    P2.Speed = 3
'sets player1's size
    P1.Size = 5
'sets player2's size
    P2.Size = 5
'sets player1's health to its size * x
    P1.Health = P1.Size * 10
'sets player2's health to its size * x
    P2.Health = P2.Size * 10
'sets player2's cannon in the direction of player1
    P2.Rotate = 44 * 177
'this initially loads the constants that have already been stated in _
the sub "Constants"
    With P1
        .NumberOfBullets = 100
        With .Bodypoint
            .Width = 17 * P1.Size
            .Height = 17 * P1.Size
        End With
        
        With .DirectionPoint
            .Width = 9 * P1.Size
            .Height = 9 * P1.Size
        End With
        
        For X = 0 To .NumberOfBullets
            With .BulletPoint(X)
                .Width = 9 * P1.Size
                .Height = 9 * P1.Size
            End With
        Next
    End With
    With P2
        .NumberOfBullets = P1.NumberOfBullets

        With .Bodypoint
            .Width = 17 * P2.Size
            .Height = 17 * P2.Size
        End With
        
        With .DirectionPoint
            .Width = 9 * P2.Size
            .Height = 9 * P2.Size
        End With
                
        For X = 0 To .NumberOfBullets
            With .BulletPoint(X)
                .Width = 9 * P2.Size
                .Height = 9 * P2.Size
            End With
        Next
    End With

'this places the x and y of each player's bodypoints
    P1.Bodypoint.X = Form1.ScaleLeft + Form1.ScaleWidth / 2 - P1.Bodypoint.Width / 2
    P1.Bodypoint.Y = Form1.ScaleTop + Form1.ScaleHeight * 1 / 4 - P1.Bodypoint.Height / 2
    P2.Bodypoint.X = Form1.ScaleLeft + Form1.ScaleWidth / 2 - P2.Bodypoint.Width / 2
    P2.Bodypoint.Y = Form1.ScaleTop + Form1.ScaleHeight * 3 / 4 - P2.Bodypoint.Height / 2
End Sub

'This timer sets the direction of the bullet and where the bullet comes from _
and what color the bullet is
Sub Timer2_Timer()
'everything in this with statement is going to have something to do with p1 _
if it starts with a> .
    With P1
    'if you press space then the following will happen
        If GetAsyncKeyState(vbKeySpace) Then
        'everything in this sub is going to have something to do _
        with p1.bulletpoint(.count) if it starts with a .
            With .BulletPoint(.Count)
            'this starts a randomization of colors
                Randomize
            'do not include red
                .R = (Int(Rnd * 0))
            'randomly go through all of the colors of green
                .G = (Int(Rnd * 255))
            'set blue to as blue as it can get
                .B = (Int(Rnd * 255) + 255)
            'set the direction of the current bullet
                .Direction = P1.Rotate
            'set where the x and y of the bullet point come from
                .X = P1.DirectionPoint.X + P1.DirectionPoint.Width / 2 - .Width / 2
                .Y = P1.DirectionPoint.Y + P1.DirectionPoint.Height / 2 - .Height / 2
            'the bullet has been shot
                .Shot = True
            End With
            
            'if the number of bullets is exceded then go to the first bullet
            If .Count >= P1.NumberOfBullets Then .Count = -1
            
            'the number of bullets adds up every 100th of a secont that the _
            space bar is held
            .Count = .Count + 1
        End If
    End With
    
'every thing that happend in the player1 with statement will happen here too _
the only thing that changes is that this statement uses player2
    With P2
        If GetAsyncKeyState(vbKeyReturn) Then
            With .BulletPoint(.Count)
                Randomize
                .R = (Int(Rnd * 255 + 255))
                .G = (Int(Rnd * 255))
                .B = (Int(Rnd * 0))
                .Direction = P2.Rotate
                .X = P2.DirectionPoint.X + P2.DirectionPoint.Width / 2 - .Width / 2
                .Y = P2.DirectionPoint.Y + P2.DirectionPoint.Height / 2 - .Height / 2
                .Shot = True
            End With
            If .Count >= P2.NumberOfBullets Then .Count = -1

            .Count = .Count + 1
        End If
    End With
End Sub

'cheacks if the bullets have overlaped any of its opponets bodypoints
Sub Detection()
'use all of the bullets in this array by calling x
    For X = 0 To P1.NumberOfBullets
    'cheacks if the bullets have overlaped any of its opponets bodypoints
        If P1.BulletPoint(X).Shot = True And P1.BulletPoint(X).X + P1.BulletPoint(X).Width > P2.Bodypoint.X And P1.BulletPoint(X).X < P2.Bodypoint.X + P2.Bodypoint.Width And P1.BulletPoint(X).Y + P1.BulletPoint(X).Height > P2.Bodypoint.Y And P1.BulletPoint(X).Y < P2.Bodypoint.Y + P2.Bodypoint.Height Then P1.BulletPoint(X).Overlap = True Else P1.BulletPoint(X).Overlap = False
    
    'cheacks if the bullets have overlaped any of its opponets bodypoints
        If P2.BulletPoint(X).Shot = True And P2.BulletPoint(X).X + P2.BulletPoint(X).Width > P1.Bodypoint.X And P2.BulletPoint(X).X < P1.Bodypoint.X + P1.Bodypoint.Width And P2.BulletPoint(X).Y + P2.BulletPoint(X).Height > P1.Bodypoint.Y And P2.BulletPoint(X).Y < P1.Bodypoint.Y + P1.Bodypoint.Height Then P2.BulletPoint(X).Overlap = True Else P2.BulletPoint(X).Overlap = False
                    
    'if the bullet is overlaping then the following happens
        If P1.BulletPoint(X).Overlap = True Then
        'the player that fired the bullet gets bigger
            P1.Size = P1.Size + 0.02 * P1.Size
        'and its heath get bigger
            P1.Health = P1.Health + 0.2 * P1.Size
        'and he stays centered because of these lines of code
            P1.Bodypoint.X = P1.Bodypoint.X - 0.1779 * P1.Size
            P1.Bodypoint.Y = P1.Bodypoint.Y - 0.1779 * P1.Size
        'the one that got hit gets less health
            P2.Health = P2.Health - 0.2 * P1.Size
        'and shrinks
            P2.Size = P2.Size - 0.02 * P1.Size
        'but still stays centered because of the following lines of code
            P2.Bodypoint.X = P2.Bodypoint.X + 0.1779 * P1.Size
            P2.Bodypoint.Y = P2.Bodypoint.Y + 0.1779 * P1.Size
        'and the bullet that was fired disappears
            P1.BulletPoint(X).Shot = False
        End If
        
    'same thing except player2 fired the bullet
        If P2.BulletPoint(X).Overlap = True Then
            P2.Size = P2.Size + 0.02 * P2.Size
            P2.Health = P2.Health + 0.2 * P2.Size
            P2.Bodypoint.X = P2.Bodypoint.X - 0.1779 * P2.Size
            P2.Bodypoint.Y = P2.Bodypoint.Y - 0.1779 * P2.Size
            P1.Health = P1.Health - 0.2 * P2.Size
            P1.Size = P1.Size - 0.02 * P2.Size
            P1.Bodypoint.X = P1.Bodypoint.X + 0.1779 * P2.Size
            P1.Bodypoint.Y = P1.Bodypoint.Y + 0.1779 * P2.Size
            P2.BulletPoint(X).Shot = False
        End If
        
    'if your health get as low as it can go then
        If P1.Health <= 4 Then
        'a magbox congadulates the the winner
            MsgBox "Red Tank Wins"
        'and maks all bullets invisable
            For Y = 0 To P1.NumberOfBullets
              P2.BulletPoint(Y).Shot = False
              P1.BulletPoint(Y).Shot = False
            Next
        'then the size gets set for player1 and 2
            P1.Size = 5
            P2.Size = 5
        'then their health gets set
            P1.Health = P1.Size * 10
            P2.Health = P2.Size * 10
        'then their height and widht is set
            P2.Bodypoint.Width = 17 * P2.Size
            P2.Bodypoint.Height = 17 * P2.Size
            P1.Bodypoint.Width = 17 * P1.Size
            P1.Bodypoint.Height = 17 * P1.Size
        'then their location is set
            P1.Bodypoint.X = Form1.ScaleLeft + Form1.ScaleWidth / 2 - P1.Bodypoint.Width / 2
            P1.Bodypoint.Y = Form1.ScaleTop + Form1.ScaleHeight * 1 / 4 - P1.Bodypoint.Height / 2
            P2.Bodypoint.X = Form1.ScaleLeft + Form1.ScaleWidth / 2 - P2.Bodypoint.Width / 2
            P2.Bodypoint.Y = Form1.ScaleTop + Form1.ScaleHeight * 3 / 4 - P2.Bodypoint.Height / 2
        'sets player2's cannon in the direction of player1
            P2.Rotate = 44 * 177
        End If
        
        'same thing except player1 wins
        If P2.Health <= 4 Then
            MsgBox "Blue Tank Wins"
            For Y = 0 To P1.NumberOfBullets
              P2.BulletPoint(Y).Shot = False
              P1.BulletPoint(Y).Shot = False
            Next
            P1.Size = 5
            P2.Size = 5
            P1.Health = P1.Size * 10
            P2.Health = P2.Size * 10
            P2.Bodypoint.Width = 17 * P2.Size
            P2.Bodypoint.Height = 17 * P2.Size
            P1.Bodypoint.Width = 17 * P1.Size
            P1.Bodypoint.Height = 17 * P1.Size
            P1.Bodypoint.X = Form1.ScaleLeft + Form1.ScaleWidth / 2 - P1.Bodypoint.Width / 2
            P1.Bodypoint.Y = Form1.ScaleTop + Form1.ScaleHeight * 1 / 4 - P1.Bodypoint.Height / 2
            P2.Bodypoint.X = Form1.ScaleLeft + Form1.ScaleWidth / 2 - P2.Bodypoint.Width / 2
            P2.Bodypoint.Y = Form1.ScaleTop + Form1.ScaleHeight * 3 / 4 - P2.Bodypoint.Height / 2
            P2.Rotate = 44 * 177
        End If
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
'if unload form1 is called then end the game
    End
End Sub
