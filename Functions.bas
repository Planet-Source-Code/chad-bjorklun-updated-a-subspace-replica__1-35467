Attribute VB_Name = "Functions"
Public Function Initializer()
    'Randomizes the random number table
    Randomize
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'THESE VARIABLES ARE EASY TO CHANGE TO ADJUST GAMEPLAY, YOU SHOULDN'T NEED TO CHANGE'
    'ANY OF THE OTHER CODE WHEN YOU CHANGE ONE OF THESE                                 '
    '                                                                                   '
    'Map size                                                                           '
    Map_Height = frmMain.Height * 10                                                    '
    Map_Width = frmMain.Width * 10                                                      '
    'Initial position of the visible part of the map. Upper left corner                 '
    X_Pos = Map_Width / 2 - frmMain.Width * 2                                           '
    Y_Pos = Map_Height / 2 - frmMain.Height * 2                                         '
                                                                                        '
    'This is the size, in twips, of the height and width of the radar                   '
    Radar_Size = 2535                                                                   '
                                                                                        '
    'Max number of stars (actually double this since there are 500 bright and 500 dim)  '
    Num_Stars = 500                                                                     '
                                                                                        '
    'The fastest the ship can go                                                        '
    Max_Set_Vel = 300                                                                   '
    'Max_Set_Vel / Acceleration = The amount the ships velocity will increase or        '
    Acceleration = 50 '           decrease when the up or down arrow is hit             '
                                                                                        '
    'The maximum amount of health a ship can have                                       '
    Max_Health = 2000                                                                   '
    Health = Max_Health                                                                 '
                                                                                        '
    'TurnCounter is used to set the angle that the ship is pointing at.                 '
    'I needed this at first because I used a line instead of a picture                  '
    'as the ship and this variable let me draw it at the right angle but                '
    'now it's just easier to use instead of changing all the code                       '
    TurnCounter = 30                                                                    '
    'Sets the direction the ship is drawn pointing at first                             '
    DC_Y_Pos = 576 + 36                                                                 '
    
    'This is the number of times the object turns in one rotation
    'This should only be changed it you used a different sprite
    'This is used with TurnCounter
    TurnInc = 40
    
    'Initial position of the visible part of the map. Upper left corner
    X_Pos = Map_Width / 2 - frmMain.Width * 2
    Y_Pos = Map_Height / 2 - frmMain.Height * 2
    
    KQuit = vbKeyEscape         'Change these to customize keys
    KUp = vbKeyUp               '
    KDown = vbKeyDown           '
    KLeft = vbKeyLeft           '
    KRight = vbKeyRight         '
    KAfterBurners = vbKeySpace  '
    KBullet = vbKeyControl      '
    KBomb = vbKeyShift          '
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'EVERYTHING BELOW IS STUFF THAT YOU CAN STILL CUSTOMIZE BUT YOU HAVE TO HAVE A      '
    'BETTER UNDERSTANDING OF THE CODE FIRST BEFORE CHANGING THESE                       '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    V_Wall.Sprite = DC_Tiles
    V_Wall.Origin_X = 144
    V_Wall.Origin_Y = 128
    V_Wall.Width = 16
    V_Wall.Height = 16
    
    H_Wall.Sprite = DC_Tiles
    H_Wall.Origin_X = 206
    H_Wall.Origin_Y = 128
    H_Wall.Width = 16
    H_Wall.Height = 16
    
    V_Walls(0).DC_Pic = V_Wall
    V_Walls(0).Lower_Y = 0
    V_Walls(0).Upper_Y = Map_Height + (Map_Height Mod V_Walls(0).DC_Pic.Height) ' - V_Walls(0).DC_Pic.Height
    V_Walls(0).x = 0
    V_Walls(1).DC_Pic = V_Wall
    V_Walls(1).Lower_Y = 0
    V_Walls(1).Upper_Y = Map_Height ' + (Map_Height Mod V_Walls(1).DC_Pic.Height) ' - V_Walls(1).DC_Pic.Height
    H_Walls(0).DC_Pic = H_Wall
    H_Walls(0).Lower_X = 0
    H_Walls(0).Upper_X = Map_Width ' + (Map_Width Mod H_Walls(0).DC_Pic.Width) '- H_Walls(0).DC_Pic.Width
    H_Walls(0).y = 0
    H_Walls(1).DC_Pic = H_Wall
    H_Walls(1).Lower_X = 0
    H_Walls(1).Upper_X = Map_Width ' + (Map_Width Mod H_Walls(1).DC_Pic.Width) '- H_Walls(1).DC_Pic.Width
    H_Walls(1).y = Map_Height - H_Walls(1).DC_Pic.Height * Screen.TwipsPerPixelY ' - H_Walls(1).DC_Pic.Height ' - V_Walls(1).DC_Pic.Height
    V_Walls(1).x = Map_Width - V_Walls(1).DC_Pic.Width * Screen.TwipsPerPixelX ' + (Map_Width Mod H_Walls(1).DC_Pic.Width) - V_Walls(1).DC_Pic.Width ' - H_Walls(1).DC_Pic.Wid   th
    
    'Hmmm, what could this be
    pi = 3.14159265358979
    
    'This makes a bunch of random stars...Right now it seems to be fairly
    'inefficient but I just wanted to see what it looked like. Should be
    'optimized.
    For i = 0 To Num_Stars
        Bright_Star(i).x = Int(Rnd * Map_Width)     'X value within the limits of the map
        Bright_Star(i).y = Int(Rnd * Map_Height)    'Y value within the limits of the map
        Bright_Star(i).Color = &HFFFFFF             'White Color
        Bright_Star(i).Size = 3                     'Form draw width
        Dim_Star(i).x = Int(Rnd * Map_Width)        '
        Dim_Star(i).y = Int(Rnd * Map_Height)       '
        Dim_Star(i).Color = &HFF&      '&HC0&       'Red Color
        Dim_Star(i).Size = 2                        '
    Next i
    
    'This picks random places for the 10 background
    'pictures to be drawn
    For i = 0 To 50
        With BG_Pics(i)
            .Index = Int(Rnd * 9) + 1           'This picks which picture will be drawn
            .x = Int(Rnd * (Map_Width - 66))    'X value within the limits of the map
            .y = Int(Rnd * (Map_Height - 66))   'Y value within the limits of the map
            If i > 0 Then
                For j = 0 To i - 1
                    If BG_Pics(j).x < .x + Map_Height / 18 And BG_Pics(j).x > .x - Map_Height / 18 Then
                        If BG_Pics(j).y < .y + Map_Height / 18 And BG_Pics(j).y > .y - Map_Height / 18 Then
                            j = i       'This just makes sure the background pictures don't
                            i = i - 1   'get drawn too close
                        End If
                    End If
                Next j
            End If
        End With
    Next i
    
    BombTimer = Timer - 10
    BulletTimer = Timer - 10
    For N = 0 To 200
        Explosions(N).Origin_X = -1
    Next N
    
    ''''''''''''''''''''''''''''''''''''
    For i = 0 To 50
        NME_LookAngleInc(i) = 0
        NME_Max_Set_Vel(i) = 300
        NME_Max_Vel(i) = 200
        NME_X_Pos(i) = Rnd * Map_Width
        NME_Y_Pos(i) = Rnd * Map_Height
        NME_Ship_Type(i) = Int(Rnd * 7)
        
        'Sets the direction the ship is draw pointing at first
        NME_DC_Y_Pos(i) = 576 + 36
    Next i
    
    Turrets(0).Dest_Y = Map_Height / 5
    For i = 0 To 15
        Turrets(i).Dest_X = Map_Width / 5 * ((i Mod 4) + 1)
        If i > 0 Then
            Turrets(i).Dest_Y = Turrets(i - 1).Dest_Y
            If i Mod 4 = 0 Then Turrets(i).Dest_Y = Turrets(i).Dest_Y + Turrets(i - 1).Dest_Y
        End If
        Turrets(i).Sprite = DC_NME_Turret
        Turrets(i).Columns = 5
        Turrets(i).Rows = 2
        Turrets(i).Width = 96
        Turrets(i).Height = 96
    Next i
    
    Call Main_Loop
End Function

'This function contains the loop that
Public Function Main_Loop()
    frmMain.Visible = True
    frmMain.Refresh
    Do
        'When KQuit(currently Escape) is pressed, this will call the Good_Bye function
        If GetKey(KQuit) <> 0 Then Call Good_Bye

'Okay, currently the movement code sucks. While it does stay pretty true to frictionless motion, it does
'not mimick Subspace's motion. I just haven't been able to figure out the algorithm yet. Soon though.
'If you plan on making a true game out of this, you should definately change this.

        If GetKey(KAfterBurners) <> 0 And Health > 15 Then
            Max_Vel = Max_Set_Vel + 200
        Else
            Max_Vel = Max_Set_Vel
        End If
        'If GetKey(KUp) = -32767 Then
        If GetKey(KUp) <> 0 Then
            'Here's the pseudo-equation
            'If [Current_Vector] isn't = to [Max_Vel_Allowed_in_the_direction_you_are_pointing] then
            '[Current_Vec] = [Current_Vec] plus [a_50th_of_the_speed_in_the_direction_you_are_pointing]
            If H_Vel <> Cos(Look_Angle) * Max_Vel Then H_Vel = H_Vel + Cos(Look_Angle) * (Max_Vel / 50)
            If V_Vel <> Sin(Look_Angle) * Max_Vel Then V_Vel = V_Vel + Sin(Look_Angle) * (Max_Vel / 50)
            If GetKey(KAfterBurners) <> 0 And Health > 15 Then Health = Health - 15
            Call Start_Exhaust
        End If
        'If GetKey(KDown) = -32767 Then
        If GetKey(KDown) <> 0 Then
            If H_Vel <> Cos(Look_Angle) * Max_Vel * -1 Then H_Vel = H_Vel - Cos(Look_Angle) * (Max_Vel / 50)
            If V_Vel <> Sin(Look_Angle) * Max_Vel * -1 Then V_Vel = V_Vel - Sin(Look_Angle) * (Max_Vel / 50)
            If GetKey(KAfterBurners) <> 0 And Health > 15 Then Health = Health - 15
            Call Start_Exhaust
        End If
        
        'This keeps the velocity in check, in both positive and negative directions
        If Abs(H_Vel) > Max_Vel Then H_Vel = Max_Vel * H_Vel / Abs(H_Vel)
        If Abs(V_Vel) > Max_Vel Then V_Vel = Max_Vel * V_Vel / Abs(V_Vel)
        
        
        'If GetKey(KLeft) = -32767 Then Call Turn_Left
        'If GetKey(KRight) = -32767 Then Call Turn_Right
        If GetKey(KLeft) <> 0 Then Call Turn_Left
        If GetKey(KRight) <> 0 Then Call Turn_Right
        
        'If Roll_Pos is negative this adds one to it, if it is positive then
        'it subracts one from it. If you couldn't tell, this is part of the
        'rolling ship effect
        If GetKey(KLeft) = 0 And GetKey(KRight) = 0 And Roll_Pos <> 0 Then
            Roll_Pos = Roll_Pos + -1 * Abs(Roll_Pos) / Roll_Pos
        End If
        
        'This just makes sure the variable stays positive and won't overflow
        If LookAngleInc < 0 Then LookAngleInc = TurnInc + LookAngleInc
        
        'This also makes sure the variable won't overflow
        LookAngleInc = LookAngleInc Mod TurnInc
        
        If GetKey(KBomb) <> 0 Then Call Fire_Bomb
        If GetKey(KBullet) <> 0 Then Call Fire_Bullet
        
'You are probably wondering what the '= -32767' stuff is. It is the value that GetAsyncKeyState returns
'when it registers a key first being pressed. I used it when the program ran a lot faster because it helped
'slow things down. When it is used in a loop, GetAsyncKeyState returns a value not equal to zero every
'millasecond while the key being watched is pressed. It is pretty much impossible to press a key for less
'than a millesecond so when you just tap the key, it will record ~10 hits. However, it will only return
'-32767 every ~10 milleseconds. That way it will only read 1 hit. Anyway, I left those in there in case
'someone optimizes the code or has a really fast computer and it's running too fast.
        
        'Calculates the angle that the ship is pointing in
        Look_Angle = pi * LookAngleInc / (TurnInc / 2)
        
        'Increments the position
        X_Pos = X_Pos + H_Vel
        Y_Pos = Y_Pos + V_Vel
        
        Call Check_Wall_Collision
        
        'This just draws a big black square of the entire screen. You could also use
        'frmMain.Cls but that make the screen flash too much. If you have a decent understanding
        'of BitBlt, it would be fairly easy to make your own background and replace the DC_Clear
        'picture with that of your backgro und.   You would have to code for the scrolling effect but
        'that wouldn't be very hard at a ll
        BitBlt frmMain.HDC, 0, 0, frmMain.Width / Screen.TwipsPerPixelX, frmMain.Height / Screen.TwipsPerPixelY, DC_Clear, 0, 0, vbSrcCopy
        
        Call Draw_Background
        Call Draw_Exhaust
        Call Draw_Walls
        Call Draw_Projectile
        
        Call NME(0)
        Call NME(1)
        Call NME(2)
        Call NME(3)
         
        Call Radar
        Call Health_Bar
         Call Display
        'Draws the ship in the center of the screen
        BitBlt frmMain.HDC, (frmMain.Width / 2) / Screen.TwipsPerPixelX - 18, (frmMain.Height / 2) / Screen.TwipsPerPixelY - 18, 36, 36, DC_Ships_Mask(Ship_Type), DC_X_Pos, DC_Y_Pos + 144 * Int(Roll_Pos / 3), vbSrcAnd
        BitBlt frmMain.HDC, (frmMain.Width / 2) / Screen.TwipsPerPixelX - 18, (frmMain.Height / 2) / Screen.TwipsPerPixelY - 18, 36, 36, DC_Ships_Sprite(Ship_Type), DC_X_Pos, DC_Y_Pos + 144 * Int(Roll_Pos / 3), vbSrcPaint
         frmMain.Refresh
    Loop
End Function

Public Function Turn_Right()
    'LookAngleInc is used to find the angle that the ship is pointing in
    LookAngleInc = LookAngleInc + 1
    
    'This code makes it so the ship rolls to the opposite direction when
    'you are thrusting in reverse
    If GetKey(KUp) <> 0 Then                            'There are really only 4 different roll
        If Roll_Pos < 12 Then Roll_Pos = Roll_Pos + 1   'positions but I set the max to 12 because
    ElseIf GetKey(KDown) <> 0 Then                      'it was rolling too fast. In the BitBlt code
        If Roll_Pos > -12 Then Roll_Pos = Roll_Pos - 1  'that draws the ship, I have the Roll_Pos
    End If                                              'divided by 3 then rounded to draw it right.
    
    DC_X_Pos = DC_X_Pos + 36        'Look at Turn_Left for description
    If DC_X_Pos = 36 * 10 Then      '
        DC_X_Pos = 0                '
        DC_Y_Pos = DC_Y_Pos + 36    '
    End If                          '
    If DC_Y_Pos = 576 + 144 Then DC_Y_Pos = 576
End Function

Public Function Turn_Left()
    'LookAngleInc is used to find the angle that the ship is pointing in
    LookAngleInc = LookAngleInc - 1
    
    'This code makes it so the ship rolls to the opposite direction when
    'you are thrusting in reverse
    If GetKey(KUp) <> 0 Then                            'Look at Turn_Right for description
        If Roll_Pos > -12 Then Roll_Pos = Roll_Pos - 1  '
    ElseIf GetKey(KDown) <> 0 Then                      '
        If Roll_Pos < 12 Then Roll_Pos = Roll_Pos + 1   '
    End If                                              '
    
    DC_X_Pos = DC_X_Pos - 36        'The sprite used for the ships has 32 rows and 10 columns. Each cell
    If DC_X_Pos = -36 Then          'is a picture of a ship pointing in a different direction. 4 rows
        DC_X_Pos = 36 * 9           'and 10 columns are devoted to each ship. This code finds the upper
        DC_Y_Pos = DC_Y_Pos - 36    'left corner of the cell that will be drawn
    End If
    If DC_Y_Pos = 576 - 36 Then DC_Y_Pos = 576 + 144 - 36
End Function

'The only reason for this function is that I didn't want to write out
'GetAsyncKeyState everytime. To minimize the code, you could simply
'replace it in the main function
Public Function GetKey(key As Long) As Integer
     GetKey = GetAsyncKeyState(key)
End Function
                       
'Blts a square then draws a dot on top, pretty basic.  You can set the max radar size to about 4500
'Thats the size of the gray picture loaded into memory.
Public Function Radar()
    frmMain.DrawWidth = 3
    frmMain.ForeColor = &HFFFF&
    BitBlt frmMain.HDC, (frmMain.Width - Radar_Size) / Screen.TwipsPerPixelX, (frmMain.Height - Radar_Size) / Screen.TwipsPerPixelY, Radar_Size, Radar_Size, DC_Radar, 0, 0, vbSrcCopy
    frmMain.PSet (((X_Pos + frmMain.Width / 2) / Map_Width * Radar_Size) + frmMain.Width - Radar_Size, ((Y_Pos + frmMain.Height / 2) / Map_Height * Radar_Size) + frmMain.Height - Radar_Size)
    frmMain.ForeColor = &HFF0000
    For r = 0 To 3 ''''
        frmMain.PSet (((NME_X_Pos(r) + frmMain.Width / 2) / Map_Width * Radar_Size) + frmMain.Width - Radar_Size, ((NME_Y_Pos(r) + frmMain.Height / 2) / Map_Height * Radar_Size) + frmMain.Height - Radar_Size)
    Next r
    For p = 0 To 200
        If Explosions(p).Origin_X <> -1 Then
            frmMain.ForeColor = &HC0&
            frmMain.PSet (((Explosions(p).Dest_X) / Map_Width * Radar_Size) + frmMain.Width - Radar_Size, ((Explosions(p).Dest_Y) / Map_Height * Radar_Size) + frmMain.Height - Radar_Size)
        End If
    Next p
End Function

Public Function Fire_Bomb()
    If Timer - 1 > BombTimer Then
        If BombCounter < 200 And Health > 300 Then
            Health = Health - 300
            BombCounter = BombCounter + 1
            For f = 0 To 200
                With Bombs(f)
                    If .Active = False Then
                        .X_Vel = Cos(Look_Angle) * Max_Vel * 1.5 + H_Vel
                        .Y_Vel = Sin(Look_Angle) * Max_Vel * 1.5 + V_Vel
                        .Strength = 3   'CHANGE THIS
                        .Type = 1       'CHANGE THIS
                        .Timer = 10000
                        .DC_X_Pos = 0
                        .Active = True
                        f = 200
                    End If
                End With
            Next f
        End If
        BombTimer = Timer
    End If
End Function

Public Function Fire_Bullet()
    If Timer - 0.2 > BulletTimer Then
        If BulletCounter < 200 And Health > 20 Then
            Health = Health - 20
            BulletCounter = BulletCounter + 1
            For g = 0 To 200
                With Bullets(g)
                    If .Active = False Then
                        .X_Vel = Cos(Look_Angle) * Max_Vel * 1 + H_Vel
                        .Y_Vel = Sin(Look_Angle) * Max_Vel * 1 + V_Vel
                        .Strength = 2   'CHANGE THIS
                        .Type = 0       'CHANGE THIS
                        .Timer = 10000
                        .DC_X_Pos = 0
                        .Active = True
                        g = 200
                    End If
                End With
            Next g
        End If
        BulletTimer = Timer
    End If
End Function
  
Public Function Draw_Projectile()
    For d = 0 To 200
        If Bombs(d).Active = True Then
            If Bombs(d).Timer = 10000 Then
                Bombs(d).x = X_Pos + frmMain.Width / 2 + (Cos(Look_Angle) * (16) - 8) * Screen.TwipsPerPixelX
                Bombs(d).y = Y_Pos + frmMain.Height / 2 + (Sin(Look_Angle) * (16) - 8) * Screen.TwipsPerPixelY
            End If
            For l = 0 To 1
                If (Bombs(d).x + 16 * Screen.TwipsPerPixelX >= V_Walls(l).x And Bombs(d).x < V_Walls(l).x + V_Walls(l).DC_Pic.Width * Screen.TwipsPerPixelX And Bombs(d).y + 16 * Screen.TwipsPerPixelX > V_Walls(l).Lower_Y And Bombs(d).y < V_Walls(l).Upper_Y + V_Walls(l).DC_Pic.Width * Screen.TwipsPerPixelX) And Explosions(d).Origin_X = -1 Then
                    Explosions(d) = Explosions(201 + 2) 'Bombs(d).Strength)
                    Explosions(d).Dest_X = V_Walls(l).x - (Explosions(d).Width / 2 + 8) * Screen.TwipsPerPixelX
                    Explosions(d).Dest_Y = Bombs(d).y - (Explosions(d).Height / 2 + 8) * Screen.TwipsPerPixelY
                    If l = 0 Then Explosions(d).Dest_X = Explosions(d).Dest_X + V_Walls(l).DC_Pic.Width * Screen.TwipsPerPixelX
                    Bombs(d).X_Vel = 0
                    Bombs(d).Y_Vel = 0
                    Bombs(d).x = Explosions(d).Dest_X
                    Bombs(d).y = Explosions(d).Dest_Y
                    Explosions(d).Origin_X = 0
                    Explosions(d).Origin_Y = 0
                    Bombs(d).Timer = 100
                End If
                If (Bombs(d).y + 16 * Screen.TwipsPerPixelX >= H_Walls(l).y And Bombs(d).y < H_Walls(l).y + H_Walls(l).DC_Pic.Height * Screen.TwipsPerPixelX And Bombs(d).x + 16 * Screen.TwipsPerPixelX > H_Walls(l).Lower_X And Bombs(d).x < H_Walls(l).Upper_X + H_Walls(l).DC_Pic.Height * Screen.TwipsPerPixelX) And Explosions(d).Origin_X = -1 Then
                    Explosions(d) = Explosions(201 + 2) ' Bombs(d).Strength)
                    Explosions(d).Dest_X = Bombs(d).x - (Explosions(d).Width / 2 + 8) * Screen.TwipsPerPixelX
                    Explosions(d).Dest_Y = H_Walls(l).y - (Explosions(d).Height / 2 + 8) * Screen.TwipsPerPixelY
                    If l = 0 Then Explosions(d).Dest_Y = Explosions(d).Dest_Y + H_Walls(l).DC_Pic.Height * Screen.TwipsPerPixelY
                    Bombs(d).X_Vel = 0
                    Bombs(d).Y_Vel = 0
                    Bombs(d).x = Explosions(d).Dest_X
                    Bombs(d).y = Explosions(d).Dest_Y
                    Explosions(d).Origin_X = 0
                    Explosions(d).Origin_Y = 0
                    Bombs(d).Timer = 100
                End If
            Next l
            For l = 0 To 3
                If Bombs(d).x + 16 * Screen.TwipsPerPixelX >= NME_X_Pos(l) + frmMain.Width / 2 And Bombs(d).x <= NME_X_Pos(l) + frmMain.Width / 2 + 32 * Screen.TwipsPerPixelX And Explosions(d).Origin_X = -1 Then
                    If Bombs(d).y + 16 * Screen.TwipsPerPixelY >= NME_Y_Pos(l) + frmMain.Height / 2 And Bombs(d).y <= NME_Y_Pos(l) + frmMain.Height / 2 + 32 * Screen.TwipsPerPixelY Then
                        Explosions(d) = Explosions(201 + 1) ' Bombs(d).Strength)
                        Explosions(d).Dest_X = Bombs(d).x - (Explosions(d).Width / 2 + 8) * Screen.TwipsPerPixelX
                        Explosions(d).Dest_Y = Bombs(d).y - (Explosions(d).Height / 2 + 8) * Screen.TwipsPerPixelY
                        Bombs(d).X_Vel = 0
                        Bombs(d).Y_Vel = 0
                        Bombs(d).x = Explosions(d).Dest_X
                        Bombs(d).y = Explosions(d).Dest_Y
                        Explosions(d).Origin_X = 0
                        Explosions(d).Origin_Y = 0
                        Bombs(d).Timer = 100
                    End If
                End If
            Next l
            If Bombs(d).x > -1000 And Bombs(d).y > -1000 And Bombs(d).x < Map_Width + 1000 And Bombs(d).y < Map_Height + 1000 Then
                '''''''''''''
                If Bombs(d).x - X_Pos > 0 And Bombs(d).x - X_Pos < frmMain.Width And Bombs(d).y - Y_Pos > 0 And Bombs(d).y - Y_Pos < frmMain.Height Then
                    If Explosions(d).Origin_X = -1 Then
                        BitBlt frmMain.HDC, (Bombs(d).x - X_Pos) / Screen.TwipsPerPixelX, (Bombs(d).y - Y_Pos) / Screen.TwipsPerPixelY, 16, 16, DC_Bombs_Mask, Bombs(d).DC_X_Pos, 16 * Bombs(d).Type + 16 * 4 * Bombs(d).Type, vbSrcAnd
                        BitBlt frmMain.HDC, (Bombs(d).x - X_Pos) / Screen.TwipsPerPixelX, (Bombs(d).y - Y_Pos) / Screen.TwipsPerPixelY, 16, 16, DC_Bombs_Sprite, Bombs(d).DC_X_Pos, 16 * Bombs(d).Type + 16 * 4 * Bombs(d).Type, vbSrcPaint
                    Else
                        BitBlt frmMain.HDC, (Explosions(d).Dest_X - X_Pos) / Screen.TwipsPerPixelX, (Explosions(d).Dest_Y - Y_Pos) / Screen.TwipsPerPixelY, Explosions(d).Width, Explosions(d).Height, Explosions(d).Mask, Explosions(d).Origin_X, Explosions(d).Origin_Y, vbSrcAnd
                        BitBlt frmMain.HDC, (Explosions(d).Dest_X - X_Pos) / Screen.TwipsPerPixelX, (Explosions(d).Dest_Y - Y_Pos) / Screen.TwipsPerPixelY, Explosions(d).Width, Explosions(d).Height, Explosions(d).Sprite, Explosions(d).Origin_X, Explosions(d).Origin_Y, vbSrcPaint
                    End If
                End If
                If Explosions(d).Origin_X <> -1 Then
                    Explosions(d).Origin_X = Explosions(d).Origin_X + Explosions(d).Width
                    If Explosions(d).Origin_X = Explosions(d).Width * Explosions(d).Columns Then
                        Explosions(d).Origin_X = 0
                        Explosions(d).Origin_Y = Explosions(d).Origin_Y + Explosions(d).Height
                        If Explosions(d).Origin_Y = Explosions(d).Height * Explosions(d).Rows Then
                            Explosions(d).Origin_X = -1
                            Bombs(d).Timer = 0
                        End If
                    End If
                End If
                
                
                
                Bombs(d).x = Bombs(d).x + Bombs(d).X_Vel
                Bombs(d).y = Bombs(d).y + Bombs(d).Y_Vel
                Bombs(d).DC_X_Pos = Bombs(d).DC_X_Pos + 16
                If Bombs(d).DC_X_Pos = 160 Then Bombs(d).DC_X_Pos = 0
                Bombs(d).Timer = Bombs(d).Timer - 1
                If Bombs(d).Timer <= 0 Then Bombs(d).x = -9999
            Else
                Bombs(d).Active = False
                Explosions(d).Origin_X = -1
                BombCounter = BombCounter - 1
            End If
        End If
        If Bullets(d).Active = True Then
            If Bullets(d).Timer = 10000 Then
                Bullets(d).x = X_Pos + frmMain.Width / 2 + Cos(Look_Angle) * (18 + 2) * Screen.TwipsPerPixelX
                Bullets(d).y = Y_Pos + frmMain.Height / 2 + Sin(Look_Angle) * (18 + 2) * Screen.TwipsPerPixelY
            End If
            If Bullets(d).x > 0 And Bullets(d).y > 0 And Bullets(d).x < Map_Width And Bullets(d).y < Map_Height Then
                If Bullets(d).x - X_Pos > 0 And Bullets(d).x - X_Pos < frmMain.Width And Bullets(d).y - Y_Pos > 0 And Bullets(d).y - Y_Pos < frmMain.Height Then
                    BitBlt frmMain.HDC, (Bullets(d).x - X_Pos) / Screen.TwipsPerPixelX, (Bullets(d).y - Y_Pos) / Screen.TwipsPerPixelY, 5, 5, DC_Bullets_Mask, 0, 20 * Bullets(d).Type + 5 * Bullets(d).Strength, vbSrcAnd
                    BitBlt frmMain.HDC, (Bullets(d).x - X_Pos) / Screen.TwipsPerPixelX, (Bullets(d).y - Y_Pos) / Screen.TwipsPerPixelY, 5, 5, DC_Bullets_Sprite, 0, 20 * Bullets(d).Type + 5 * Bullets(d).Strength, vbSrcPaint
                    'For e = 1 To 3
                    '    BitBlt frmMain.HDC, (Bullets(d).X - Bullets(d).X_Vel * e / 15 - X_Pos) / Screen.TwipsPerPixelX, (Bullets(d).Y - Bullets(d).Y_Vel * e / 15 - Y_Pos) / Screen.TwipsPerPixelY, 5, 5, DC_Bullets_Mask, 5 * e, 20 * Bullets(d).Type + 5 * Bullets(d).Strength, vbSrcAnd
                    '    BitBlt frmMain.HDC, (Bullets(d).X - Bullets(d).X_Vel * e / 15 - X_Pos) / Screen.TwipsPerPixelX, (Bullets(d).Y - Bullets(d).Y_Vel * e / 15 - Y_Pos) / Screen.TwipsPerPixelY, 5, 5, DC_Bullets_Sprite, 5 * e, 20 * Bullets(d).Type + 5 * Bullets(d).Strength, vbSrcPaint
                    'Next e
                End If
                Bullets(d).x = Bullets(d).x + Bullets(d).X_Vel
                Bullets(d).y = Bullets(d).y + Bullets(d).Y_Vel
                Bullets(d).DC_X_Pos = Bullets(d).DC_X_Pos + 4
                If Bullets(d).DC_X_Pos = 16 Then Bullets(d).DC_X_Pos = 0
                Bullets(d).Timer = Bullets(d).Timer - 1
                If Bullets(d).Timer <= 0 Then Bullets(d).x = -99
            Else
                Bullets(d).Active = False
                BulletCounter = BulletCounter - 1
            End If
        End If
    Next d
End Function

Public Function Draw_Walls()
    Dim Draw_X As Single
    Dim Draw_Y As Single
    For k = 0 To 1
        If Y_Pos + frmMain.Height > H_Walls(k).y And Y_Pos < H_Walls(k).y - H_Walls(k).DC_Pic.Height Then
            If X_Pos + frmMain.Width > H_Walls(k).Lower_X And X_Pos < H_Walls(k).Upper_X - H_Walls(k).DC_Pic.Width Then
                If X_Pos > H_Walls(k).Lower_X Then
                    Draw_X = H_Walls(k).Lower_X Mod X_Pos
                Else
                    Draw_X = H_Walls(k).Lower_X
                End If
                While Draw_X < H_Walls(k).Upper_X And Draw_X < X_Pos + frmMain.Width
                    BitBlt frmMain.HDC, (Draw_X - X_Pos) / Screen.TwipsPerPixelX, (H_Walls(k).y - Y_Pos) / Screen.TwipsPerPixelY, H_Walls(k).DC_Pic.Width, H_Walls(k).DC_Pic.Height, H_Walls(k).DC_Pic.Sprite, H_Walls(k).DC_Pic.Origin_X, H_Walls(k).DC_Pic.Origin_Y, vbSrcCopy
                    Draw_X = Draw_X + H_Walls(k).DC_Pic.Width * Screen.TwipsPerPixelX
                Wend
            End If
        End If
        If X_Pos + frmMain.Width > V_Walls(k).x And X_Pos < V_Walls(k).x - V_Walls(k).DC_Pic.Width Then
            If Y_Pos + frmMain.Height > V_Walls(k).Lower_Y And Y_Pos < V_Walls(k).Upper_Y - V_Walls(k).DC_Pic.Height Then
                If Y_Pos > V_Walls(k).Lower_Y Then
                    Draw_Y = V_Walls(k).Lower_Y Mod Y_Pos
                Else
                    Draw_Y = V_Walls(k).Lower_Y
                End If
                While Draw_Y < V_Walls(k).Upper_Y And Draw_Y < Y_Pos + frmMain.Height
                    BitBlt frmMain.HDC, (V_Walls(k).x - X_Pos) / Screen.TwipsPerPixelX, (Draw_Y - Y_Pos) / Screen.TwipsPerPixelY, V_Walls(k).DC_Pic.Height, V_Walls(k).DC_Pic.Width, V_Walls(k).DC_Pic.Sprite, V_Walls(k).DC_Pic.Origin_X, H_Walls(k).DC_Pic.Origin_Y, vbSrcCopy
                    Draw_Y = Draw_Y + V_Walls(k).DC_Pic.Height * Screen.TwipsPerPixelY
                Wend
            End If
        End If
    Next k
End Function

Public Function Start_Exhaust()
    If ExhaustCounter = 12 Then ExhaustCounter = 0
    Exhaust(ExhaustCounter).DC_X_Pos = 0
    ExhaustCounter = ExhaustCounter + 1
End Function

Public Function Draw_Exhaust()
    For B = 0 To 11
        If Exhaust(B).DC_X_Pos = 0 Then
            Exhaust(B).x = X_Pos + frmMain.Width / 2 - (Cos(Look_Angle) * (16) + 8) * Screen.TwipsPerPixelX
            Exhaust(B).y = Y_Pos + frmMain.Height / 2 - (Sin(Look_Angle) * (16) + 8) * Screen.TwipsPerPixelY
        End If
        If Exhaust(B).DC_X_Pos >= 0 Then
            BitBlt frmMain.HDC, (Exhaust(B).x - X_Pos) / Screen.TwipsPerPixelX, (Exhaust(B).y - Y_Pos) / Screen.TwipsPerPixelY, 16, 16, DC_Exhaust_Mask, Exhaust(B).DC_X_Pos, 0, vbSrcAnd
            BitBlt frmMain.HDC, (Exhaust(B).x - X_Pos) / Screen.TwipsPerPixelX, (Exhaust(B).y - Y_Pos) / Screen.TwipsPerPixelY, 16, 16, DC_Exhaust_Sprite, Exhaust(B).DC_X_Pos, 0, vbSrcPaint
            Exhaust(B).DC_X_Pos = Exhaust(B).DC_X_Pos + 16
        End If
        If Exhaust(B).DC_X_Pos = 304 Then
            Exhaust(B).DC_X_Pos = -5
            B = 50
        End If
    Next B
End Function

'I don't really like this function.  It just seems really inefficient. I've left it though because
'it gives you a much better sense of movement.  If it is too slow, just lower the max number of stars
'If you plan to modify this, one recomendation that I would make is to use quadrants. For example,
'say I want 8000 stars, I could spread it out between four quadrants so I would only have to use a
'loop to search for the stars in that quadrant. You would have to deal with overlapping somehow but
'it would definately speed things up. Oh yeah, for a cool effect, make the smaller stars move slower
'thatn the bigger ones. You will have to revamp the code a bit but it looks really good.
Public Function Draw_Background()
    For C = 0 To Num_Stars
        'The if statements just check if the star is within the viewable area of the screen, if it is
        'then it will draw it, otherwise it just skips it
        If Bright_Star(C).x > X_Pos And Bright_Star(C).x < X_Pos + frmMain.Width Then
            If Bright_Star(C).y > Y_Pos And Bright_Star(C).y < Y_Pos + frmMain.Height Then
                frmMain.DrawWidth = Bright_Star(C).Size
                frmMain.ForeColor = Bright_Star(C).Color
                frmMain.PSet (Bright_Star(C).x - X_Pos, Bright_Star(C).y - Y_Pos)
            End If
        End If
        If Dim_Star(C).x > X_Pos And Dim_Star(C).x < X_Pos + frmMain.Width Then
            If Dim_Star(C).y > Y_Pos And Dim_Star(C).y < Y_Pos + frmMain.Height Then
                frmMain.DrawWidth = Dim_Star(C).Size
                frmMain.ForeColor = Dim_Star(C).Color
                frmMain.PSet (Dim_Star(C).x - X_Pos, Dim_Star(C).y - Y_Pos)
                End If
         End If
        If C <= 50 Then
            If BG_Pics(C).x > X_Pos - 64 * Screen.TwipsPerPixelX And BG_Pics(C).x < X_Pos + frmMain.Width Then
                If BG_Pics(C).y > Y_Pos - 64 * Screen.TwipsPerPixelY And BG_Pics(C).y < Y_Pos + frmMain.Height Then
                    BitBlt frmMain.HDC, (BG_Pics(C).x - X_Pos) / Screen.TwipsPerPixelX, (BG_Pics(C).y - Y_Pos) / Screen.TwipsPerPixelY, 64, 64, DC_BGs(BG_Pics(C).Index), 0, 0, vbSrcCopy
                End If
            End If
        End If
    Next C
End Function

Public Function Health_Bar()
    If Health <> Max_Health Then Health = Health + 5
    BitBlt frmMain.HDC, (((2000 - Health) / 2000 * 171) / 2) + ((frmMain.Width / 2) / Screen.TwipsPerPixelX - 85 - 13), 17, Health / 2000 * 85.5, 9, DC_Bar, ((2000 - Health) / 2000 * 171) / 2, 0, vbSrcCopy
    BitBlt frmMain.HDC, ((frmMain.Width / 2) / Screen.TwipsPerPixelX + 14), 17, Health / 2000 * 85.5, 9, DC_Bar, 85, 0, vbSrcCopy
    BitBlt frmMain.HDC, (frmMain.Width / 2) / Screen.TwipsPerPixelX - 86, 0, 172, 42, DC_Health_Bar, 174, 0, vbSrcAnd
    BitBlt frmMain.HDC, (frmMain.Width / 2) / Screen.TwipsPerPixelX - 86, 0, 172, 42, DC_Health_Bar, 0, 0, vbSrcPaint
End Function

Public Function Display()
    '84 x 68
    BitBlt frmMain.HDC, frmMain.Width / Screen.TwipsPerPixelX - 84, 4, 84, 68, DC_Display_Mask, 0, 0, vbSrcAnd
    BitBlt frmMain.HDC, frmMain.Width / Screen.TwipsPerPixelX - 84, 4, 84, 68, DC_Display_Sprite, 0, 0, vbSrcPaint
     Health_Temp = Health
    For q = 0 To 3
        BitBlt frmMain.HDC, frmMain.Width / Screen.TwipsPerPixelX - 84 + 10 + 16 * q, 0, 16, 22, DC_NRG_Font, (Int(Health_Temp / (10 ^ (3 - q)))) * 16, 0, vbSrcCopy
        Health_Temp = Health_Temp - (Int(Health_Temp / (10 ^ (3 - q)))) * (10 ^ (3 - q))
    Next q
End Function

Public Function Check_Wall_Collision()
    If X_Pos + frmMain.Width / 2 + 18 * Screen.TwipsPerPixelX >= V_Walls(1).x And H_Vel > 0 Then
        Call Collision("Right")
    ElseIf Y_Pos + frmMain.Height / 2 + 18 * Screen.TwipsPerPixelY >= H_Walls(1).y And V_Vel > 0 Then
        Call Collision("Bottom")
    ElseIf Y_Pos + frmMain.Height / 2 - 18 * Screen.TwipsPerPixelY <= H_Walls(0).DC_Pic.Height * Screen.TwipsPerPixelY And V_Vel < 0 Then
         Call Collision("Top")
    ElseIf X_Pos + frmMain.Width / 2 - 18 * Screen.TwipsPerPixelX <= V_Walls(0).DC_Pic.Width * Screen.TwipsPerPixelX And H_Vel < 0 Then
        Call Collision("Left")
    End If
End Function

Public Function Collision(Position As String)
    If Position = "Top" Then
        Y_Pos = H_Walls(0).DC_Pic.Height * Screen.TwipsPerPixelY - frmMain.Height / 2 + 18 * Screen.TwipsPerPixelY
        V_Vel = V_Vel * -0.5
    ElseIf Position = "Bottom" Then
        Y_Pos = H_Walls(1).y - frmMain.Height / 2 - 18 * Screen.TwipsPerPixelY
        V_Vel = V_Vel * -0.5
    ElseIf Position = "Left" Then
         X_Pos = V_Walls(0).DC_Pic.Width * Screen.TwipsPerPixelX - frmMain.Width / 2 + 18 * Screen.TwipsPerPixelX
        H_Vel = H_Vel * -0.5
    ElseIf Position = "Right" Then
        X_Pos = V_Walls(1).x - frmMain.Width / 2 - 18 * Screen.TwipsPerPixelX
        H_Vel = H_Vel * -0.5
    End If
End Function


'This just loads all of the pictures I'm going to use into memory. If you use BitBlt, I HIGHLY
'recommend you do this. Be careful though, if you don't delete them before the program quits,
'that memory is gone until you restart your computer.
Public Function Load_Pics()
    'The while loops, are necessary for slower computers that can't
    'always load the picture on the first try.
    While DC_Ships_Mask(0) = 0
        DC_Ships_Mask(0) = GenerateDC(App.Path + "\Graphics\Wbroll B&W.bmp")
    Wend
    While DC_Ships_Sprite(0) = 0
        DC_Ships_Sprite(0) = GenerateDC(App.Path + "\Graphics\Wbroll.bmp")
    Wend
    While DC_Ships_Mask(1) = 0
        DC_Ships_Mask(1) = GenerateDC(App.Path + "\Graphics\Jvroll B&W.bmp")
    Wend
    While DC_Ships_Sprite(1) = 0
        DC_Ships_Sprite(1) = GenerateDC(App.Path + "\Graphics\Jvroll.bmp")
    Wend
    While DC_Ships_Mask(2) = 0
        DC_Ships_Mask(2) = GenerateDC(App.Path + "\Graphics\Sproll B&W.bmp")
    Wend
    While DC_Ships_Sprite(2) = 0
        DC_Ships_Sprite(2) = GenerateDC(App.Path + "\Graphics\Sproll.bmp")
    Wend
    While DC_Ships_Mask(3) = 0
        DC_Ships_Mask(3) = GenerateDC(App.Path + "\Graphics\Lvroll B&W.bmp")
    Wend
    While DC_Ships_Sprite(3) = 0
        DC_Ships_Sprite(3) = GenerateDC(App.Path + "\Graphics\Lvroll.bmp")
    Wend
    While DC_Ships_Mask(4) = 0
        DC_Ships_Mask(4) = GenerateDC(App.Path + "\Graphics\Teroll B&W.bmp")
    Wend
    While DC_Ships_Sprite(4) = 0
        DC_Ships_Sprite(4) = GenerateDC(App.Path + "\Graphics\Teroll.bmp")
    Wend
    While DC_Ships_Mask(5) = 0
        DC_Ships_Mask(5) = GenerateDC(App.Path + "\Graphics\Weroll B&W.bmp")
    Wend
    While DC_Ships_Sprite(5) = 0
        DC_Ships_Sprite(5) = GenerateDC(App.Path + "\Graphics\Weroll.bmp")
    Wend
    While DC_Ships_Mask(6) = 0
        DC_Ships_Mask(6) = GenerateDC(App.Path + "\Graphics\Nwroll B&W.bmp")
    Wend
    While DC_Ships_Sprite(6) = 0
        DC_Ships_Sprite(6) = GenerateDC(App.Path + "\Graphics\Nwroll.bmp")
    Wend
    While DC_Ships_Mask(7) = 0
        DC_Ships_Mask(7) = GenerateDC(App.Path + "\Graphics\Shroll B&W.bmp")
    Wend
    While DC_Ships_Sprite(7) = 0
        DC_Ships_Sprite(7) = GenerateDC(App.Path + "\Graphics\Shroll.bmp")
    Wend
    While DC_Bombs_Mask = 0
        DC_Bombs_Mask = GenerateDC(App.Path + "\Graphics\Bombs B&W.bmp")
    Wend
    While DC_Bombs_Sprite = 0
        DC_Bombs_Sprite = GenerateDC(App.Path + "\Graphics\Bombs.bmp")
    Wend
    While DC_Bullets_Mask = 0
        DC_Bullets_Mask = GenerateDC(App.Path + "\Graphics\Bullets B&W.bmp")
    Wend
    While DC_Bullets_Sprite = 0
        DC_Bullets_Sprite = GenerateDC(App.Path + "\Graphics\Bullets.bmp")
    Wend
    While DC_Exhaust_Mask = 0
        DC_Exhaust_Mask = GenerateDC(App.Path + "\Graphics\Exhaust B&W.bmp")
    Wend
    While DC_Exhaust_Sprite = 0
        DC_Exhaust_Sprite = GenerateDC(App.Path + "\Graphics\Exhaust.bmp")
    Wend
    While DC_Health_Bar = 0
        DC_Health_Bar = GenerateDC(App.Path + "\Graphics\HealthBar.bmp")
    Wend
    While DC_Bar = 0
        DC_Bar = GenerateDC(App.Path + "\Graphics\Temp_Bar1.bmp")
    Wend
    While DC_Display_Mask = 0
        DC_Display_Mask = GenerateDC(App.Path + "\Graphics\Disp B&W.bmp")
    Wend
    While DC_Display_Sprite = 0
        DC_Display_Sprite = GenerateDC(App.Path + "\Graphics\Disp.bmp")
    Wend
    While DC_NRG_Font = 0
        DC_NRG_Font = GenerateDC(App.Path + "\Graphics\Engyfont.bmp")
    Wend
    While DC_NME_Turret = 0
        DC_NME_Turret = GenerateDC(App.Path + "\Graphics\Station.bmp")
    Wend
    'While DC_Trail_Mask = 0
    '    DC_Trail_Mask = GenerateDC(App.Path + "\Graphics\Trail B&W.bmp")
    'Wend
    'While DC_Trail_Sprite = 0
    '    DC_Trail_Sprite = GenerateDC(App.Path + "\Graphics\Trail.bmp")
    'Wend
    'Okay..Explosions() is a variable that will be used to generate the explosion animation. Since
    'there can be 201 bombs, the most explosions there can be are 201. Below I am making the last 4
    'indexes...or is it indicies...anyway, I make them equal to the pictures that I'll be using. It
    'would probably be clearer if these were just seperate variables but since I have so gosh darn
    'many already, I'm doing it this way.
    For M = 0 To 3
        Explosions(M + 201).Origin_X = -1
        If Explosions(M + 201).Mask = 0 Then Explosions(M + 201).Mask = GenerateDC(App.Path + "\Graphics\Explode" & M & " B&W.bmp")
        If Explosions(M + 201).Sprite = 0 Then Explosions(M + 201).Sprite = GenerateDC(App.Path + "\Graphics\Explode" & M & ".bmp")
        If Explosions(M + 201).Mask = 0 Or Explosions(M + 201).Sprite = 0 Then M = M - 1
    Next M
    'I came very close to just saying "screw it" and add a couple of API's to do this all for me but
    'it ended up just seeming like that would be even more work and just make things more complicated
    'So I did it the old fashion way. It's ugly but it works.
    Explosions(201).Rows = 1
    Explosions(201).Columns = 6
    Explosions(201).Width = 8
    Explosions(201).Height = 8
    Explosions(202).Rows = 6
    Explosions(202).Columns = 6
    Explosions(202).Width = 48
    Explosions(202).Height = 48
    Explosions(203).Rows = 11
    Explosions(203).Columns = 4
    Explosions(203).Width = 80
    Explosions(203).Height = 80
    Explosions(204).Rows = 1
    Explosions(204).Columns = 7
    Explosions(204).Width = 16
    Explosions(204).Height = 16
    
    While DC_Clear = 0
        DC_Clear = GenerateDC(App.Path + "\Graphics\Clear.bmp")
    Wend
    While DC_Radar = 0
         DC_Radar = GenerateDC(App.Path + "\Graphics\Radar.bmp")
    Wend
    While DC_BGs(1) = 0
        DC_BGs(1) = GenerateDC(App.Path + "\Graphics\Bg01.bmp")
    Wend
    While DC_BGs(2) = 0
        DC_BGs(2) = GenerateDC(App.Path + "\Graphics\Bg02.bmp")
    Wend
    While DC_BGs(3) = 0
        DC_BGs(3) = GenerateDC(App.Path + "\Graphics\Bg03.bmp")
    Wend
    While DC_BGs(4) = 0
        DC_BGs(4) = GenerateDC(App.Path + "\Graphics\Bg04.bmp")
    Wend
    While DC_BGs(5) = 0
        DC_BGs(5) = GenerateDC(App.Path + "\Graphics\Bg05.bmp")
    Wend
    While DC_BGs(6) = 0
        DC_BGs(6) = GenerateDC(App.Path + "\Graphics\Bg06.bmp")
    Wend
    While DC_BGs(7) = 0
        DC_BGs(7) = GenerateDC(App.Path + "\Graphics\Bg07.bmp")
    Wend
    While DC_BGs(8) = 0
        DC_BGs(8) = GenerateDC(App.Path + "\Graphics\Bg08.bmp")
    Wend
    While DC_BGs(9) = 0
        DC_BGs(9) = GenerateDC(App.Path + "\Graphics\Bg09.bmp")
    Wend
    While DC_BGs(10) = 0
        DC_BGs(10) = GenerateDC(App.Path + "\Graphics\Bg10.bmp")
    Wend
    While DC_Tiles = 0
        DC_Tiles = GenerateDC(App.Path + "\Graphics\Tiles.bmp")
    Wend
End Function

'Clears the memory of all the memory DC's that were created
Public Function Good_Bye()
    For i = 0 To 204
        If i < 8 Then
            DeleteGeneratedDC DC_Ships_Mask(i)
            DeleteGeneratedDC DC_Ships_Sprite(i)
        End If
        If i < 4 Then
            DeleteGeneratedDC DC_Explode_Mask(i)
            DeleteGeneratedDC DC_Explode_Sprite(i)
        End If
        If i < 10 Then DeleteGeneratedDC DC_BGs(i + 1)
        DeleteGeneratedDC Explosions(i).Mask
        DeleteGeneratedDC Explosions(i).Sprite
    Next i
    DeleteGeneratedDC DC_Bombs_Mask
    DeleteGeneratedDC DC_Bombs_Sprite
    DeleteGeneratedDC DC_Bullets_Mask
    DeleteGeneratedDC DC_Bullets_Sprite
    DeleteGeneratedDC DC_Clear
    DeleteGeneratedDC DC_Radar
    DeleteGeneratedDC DC_Exhaust_Mask
    DeleteGeneratedDC DC_Exhaust_Sprite
    DeleteGeneratedDC DC_Health_Bar
    DeleteGeneratedDC DC_Bar
    DeleteGeneratedDC DC_Display_Mask
    DeleteGeneratedDC DC_Display_Sprite
    DeleteGeneratedDC DC_NRG_Font
    End
End Function

'Did not make, I borrowed it from code a REALLY long time ago, I'm sorry I
'can't accredit the writer.
Public Function GenerateDC(ByVal Filename As String) As Long
    Dim DC As Long
    Dim hBitmap As Long
    'Create a Device Context, compatible with the screen
    DC = CreateCompatibleDC(0)
    If DC < 1 Then
        GenerateDC = 0
        Exit Function
    End If
    
    'Load the image....BIG NOTE: This function is not supported under NT, there you
    'can not specify the LR_LOADFROMFILE flag.
'''''The creator wrote the line above but if you put the GenerateDC function in a loop when
'''''you are creating it, it will load. Just look at my code (A counter can be added for
'''''error checking, I usually have 1000 attempts before I exit the loop and report an error)
    Const IMAGE_BITMAP As Long = 0
    Const LR_LOADFROMFILE As Long = &H10
    Const LR_CREATEDIBSECTION As Long = &H2000
    Const LR_DEFAULTSIZE As Long = &H40
    hBitmap = LoadImage(0, Filename, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

    If hBitmap = 0 Then 'Failure In loading bitmap
        DeleteDC DC
        GenerateDC = 0
        Exit Function
    End If
    'Throw the Bitmap into the Device Context

    SelectObject DC, hBitmap
    'Return the device context
    GenerateDC = DC
    'Delte the bitmap handle object
    DeleteObject hBitmap
End Function

'Just clears the memory
Public Function DeleteGeneratedDC(ByVal DC As Long) As Long
    If DC > 0 Then
        DeleteGeneratedDC = DeleteDC(DC)
    Else
        DeleteGeneratedDC = 0
    End If
End Function

