Attribute VB_Name = "Enemy"
'''''''''''MOVEMENT''''''''''''''''''
Public NME_Max_Set_Vel(50) As Single    '
Public NME_Max_Vel(50) As Single        ' These are the same as for
Public NME_Roll_Pos(50) As Integer      ' the regular ship
Public NME_H_Vel(50) As Single          '
Public NME_V_Vel(50) As Single          '
Public NME_Look_Angle(50) As Double     '
'Public NME_Move_Angle As Double     '
Public NME_X_Pos(50) As Single          '
Public NME_Y_Pos(50) As Single          '
Public NME_LookAngleInc(50) As Integer  '
'Public NME_MoveAngleInc As Integer  '
'''''''''''''''''''''''''''''''''''''

Public Rand_Num As Integer
Public NME_Forward(50) As Boolean
Public NME_Reverse(50) As Boolean
Public NME_Left(50) As Integer
Public NME_Right(50) As Integer
Public NME_Ship_Type(50) As Integer
Public NME_DC_X_Pos(50) As Integer
Public NME_DC_Y_Pos(50) As Integer

Public NME_BombCounter As Integer
Public NME_BombTimer As Variant
Public NME_BulletTimer As Variant
Public NME_BulletCounter As Integer
Public NME_Bombs(200) As Bomb
Public NME_Bullets(200) As Bomb

Public Turrets(15) As DC_Picture



Public Function NME(Index As Integer)
    NME_Forward(Index) = True
    NME_Reverse(Index) = False
    
    If NME_Forward(Index) = True Then
        If NME_H_Vel(Index) <> Cos(NME_Look_Angle(Index)) * NME_Max_Vel(Index) Then NME_H_Vel(Index) = NME_H_Vel(Index) + Cos(NME_Look_Angle(Index)) * (NME_Max_Vel(Index) / 50)
        If NME_V_Vel(Index) <> Sin(NME_Look_Angle(Index)) * NME_Max_Vel(Index) Then NME_V_Vel(Index) = NME_V_Vel(Index) + Sin(NME_Look_Angle(Index)) * (NME_Max_Vel(Index) / 50)
     End If
    
    If NME_Reverse(Index) = True Then
        If NME_H_Vel(Index) <> Cos(NME_Look_Angle(Index)) * NME_Max_Vel(Index) * -1 Then NME_H_Vel(Index) = NME_H_Vel(Index) - Cos(NME_Look_Angle(Index)) * (NME_Max_Vel(Index) / 50)
        If NME_V_Vel(Index) <> Sin(NME_Look_Angle(Index)) * NME_Max_Vel(Index) * -1 Then NME_V_Vel(Index) = NME_V_Vel(Index) - Sin(NME_Look_Angle(Index)) * (NME_Max_Vel(Index) / 50)
    End If
    
    If NME_Right(Index) <> 0 Then
        NME_Right(Index) = NME_Right(Index) - 1
        Call NME_Turn_Right(Index)
    ElseIf NME_Left(Index) <> 0 Then
        NME_Left(Index) = NME_Left(Index) - 1
        Call NME_Turn_Left(Index)
    Else
        Rand_Num = Int(Rnd * 30)
        If Rand_Num = 5 Then
            Rand_Num = Int(Rnd * 5)
            If Rand_Num = 2 Then
                NME_Right(Index) = Int(Rnd * 20)
            ElseIf Rand_Num = 3 Then
                NME_Left(Index) = Int(Rnd * 20)
            End If
        End If
    End If
        
    If Abs(NME_H_Vel(Index)) > NME_Max_Vel(Index) Then NME_H_Vel(Index) = NME_Max_Vel(Index) * NME_H_Vel(Index) / Abs(NME_H_Vel(Index))
    If Abs(NME_V_Vel(Index)) > NME_Max_Vel(Index) Then NME_V_Vel(Index) = NME_Max_Vel(Index) * NME_V_Vel(Index) / Abs(NME_V_Vel(Index))
    
    If NME_Left(Index) = 0 And NME_Right(Index) = 0 And NME_Roll_Pos(Index) <> 0 Then
        NME_Roll_Pos(Index) = NME_Roll_Pos(Index) + -1 * Abs(NME_Roll_Pos(Index)) / NME_Roll_Pos(Index)
    End If
    
    'This just makes sure the variable stays positive and won't overflow
    If NME_LookAngleInc(Index) < 0 Then NME_LookAngleInc(Index) = TurnInc + NME_LookAngleInc(Index)
    
    'This also makes sure the variable won't overflow
    NME_LookAngleInc(Index) = NME_LookAngleInc(Index) Mod TurnInc
    
    'Calculates the angle that the ship is pointing in
    NME_Look_Angle(Index) = pi * NME_LookAngleInc(Index) / (TurnInc / 2)
    
    'Increments the position
    NME_X_Pos(Index) = NME_X_Pos(Index) + NME_H_Vel(Index)
    NME_Y_Pos(Index) = NME_Y_Pos(Index) + NME_V_Vel(Index)
    
    Call NME_Check_Wall_Collision(Index)
    Call NME_Draw_Turrets
    
    If (NME_X_Pos(Index) - X_Pos + frmMain.Width / 2) + 36 * Screen.TwipsPerPixelX > 0 And (NME_X_Pos(Index) - X_Pos + frmMain.Width / 2) < frmMain.Width And (NME_Y_Pos(Index) - Y_Pos + frmMain.Height / 2) + 36 * Screen.TwipsPerPixelY > 0 And (NME_Y_Pos(Index) - Y_Pos + frmMain.Height / 2) < frmMain.Height Then
        BitBlt frmMain.HDC, (NME_X_Pos(Index) - X_Pos + frmMain.Width / 2) / Screen.TwipsPerPixelX, (NME_Y_Pos(Index) - Y_Pos + frmMain.Height / 2) / Screen.TwipsPerPixelY, 36, 36, DC_Ships_Mask(NME_Ship_Type(Index)), NME_DC_X_Pos(Index), NME_DC_Y_Pos(Index) + 144 * Int(NME_Roll_Pos(Index) / 3), vbSrcAnd
        BitBlt frmMain.HDC, (NME_X_Pos(Index) - X_Pos + frmMain.Width / 2) / Screen.TwipsPerPixelX, (NME_Y_Pos(Index) - Y_Pos + frmMain.Height / 2) / Screen.TwipsPerPixelY, 36, 36, DC_Ships_Sprite(NME_Ship_Type(Index)), NME_DC_X_Pos(Index), NME_DC_Y_Pos(Index) + 144 * Int(NME_Roll_Pos(Index) / 3), vbSrcPaint
    End If
End Function
 
Public Function NME_Turn_Right(Index As Integer)
    'LookAngleInc is used to find the angle that the ship is pointing in
    NME_LookAngleInc(Index) = NME_LookAngleInc(Index) + 1
    
    'This code makes it so the ship rolls to the opposite direction when
    'you are thrusting in reverse
    If NME_Forward(Index) = True Then                            'There are really only 4 different roll
        If NME_Roll_Pos(Index) < 12 Then NME_Roll_Pos(Index) = NME_Roll_Pos(Index) + 1   'positions but I set the max to 12 because
    ElseIf NME_Reverse(Index) = True Then                      'it was rolling too fast. In the BitBlt code
        If NME_Roll_Pos(Index) > -12 Then NME_Roll_Pos(Index) = NME_Roll_Pos(Index) - 1  'that draws the ship, I have the Roll_Pos
    End If                                              'divided by 3 then rounded to draw it right.
    
    NME_DC_X_Pos(Index) = NME_DC_X_Pos(Index) + 36        'Look at Turn_Left for description
    If NME_DC_X_Pos(Index) = 36 * 10 Then      '
        NME_DC_X_Pos(Index) = 0                '
        NME_DC_Y_Pos(Index) = NME_DC_Y_Pos(Index) + 36    '
    End If                          '
    If NME_DC_Y_Pos(Index) = 576 + 144 Then NME_DC_Y_Pos(Index) = 576
End Function

Public Function NME_Turn_Left(Index As Integer)
    'LookAngleInc is used to find the angle that the ship is pointing in
    NME_LookAngleInc(Index) = NME_LookAngleInc(Index) - 1
    
    'This code makes it so the ship rolls to the opposite direction when
    'you are thrusting in reverse
    If NME_Forward(Index) = True Then                            'Look at Turn_Right for description
        If NME_Roll_Pos(Index) > -12 Then NME_Roll_Pos(Index) = NME_Roll_Pos(Index) - 1  '
    ElseIf NME_Reverse(Index) = True Then                      '
        If NME_Roll_Pos(Index) < 12 Then NME_Roll_Pos(Index) = NME_Roll_Pos(Index) + 1   '
    End If                                              '
    
    NME_DC_X_Pos(Index) = NME_DC_X_Pos(Index) - 36        'The sprite used for the ships has 32 rows and 10 columns. Each cell
    If NME_DC_X_Pos(Index) = -36 Then          'is a picture of a ship pointing in a different direction. 4 rows
        NME_DC_X_Pos(Index) = 36 * 9           'and 10 columns are devoted to each ship. This code finds the upper
        NME_DC_Y_Pos(Index) = NME_DC_Y_Pos(Index) - 36    'left corner of the cell that will be drawn
    End If
    If NME_DC_Y_Pos(Index) = 576 - 36 Then NME_DC_Y_Pos(Index) = 576 + 144 - 36
End Function

Public Function NME_Check_Wall_Collision(Index As Integer)
    If NME_X_Pos(Index) + frmMain.Width / 2 + 18 * Screen.TwipsPerPixelX >= V_Walls(1).x And NME_H_Vel(Index) > 0 Then
        Call NME_Collision("Right", Index)
    ElseIf NME_Y_Pos(Index) + frmMain.Height / 2 + 18 * Screen.TwipsPerPixelY >= H_Walls(1).y And NME_V_Vel(Index) > 0 Then
        Call NME_Collision("Bottom", Index)
    ElseIf NME_Y_Pos(Index) + frmMain.Height / 2 - 18 * Screen.TwipsPerPixelY <= H_Walls(0).DC_Pic.Height * Screen.TwipsPerPixelY And NME_V_Vel(Index) < 0 Then
         Call NME_Collision("Top", Index)
    ElseIf NME_X_Pos(Index) + frmMain.Width / 2 - 18 * Screen.TwipsPerPixelX <= V_Walls(0).DC_Pic.Width * Screen.TwipsPerPixelX And NME_H_Vel(Index) < 0 Then
        Call NME_Collision("Left", Index)
    End If
End Function

Public Function NME_Collision(Position As String, Index As Integer)
    If Position = "Top" Then
        NME_Y_Pos(Index) = H_Walls(0).DC_Pic.Height * Screen.TwipsPerPixelY - frmMain.Height / 2 + 18 * Screen.TwipsPerPixelY
        NME_V_Vel(Index) = NME_V_Vel(Index) * -0.5
    ElseIf Position = "Bottom" Then
        NME_Y_Pos(Index) = H_Walls(1).y - frmMain.Height / 2 - 18 * Screen.TwipsPerPixelY
        NME_V_Vel(Index) = NME_V_Vel(Index) * -0.5
    ElseIf Position = "Left" Then
        NME_X_Pos(Index) = V_Walls(0).DC_Pic.Width * Screen.TwipsPerPixelX - frmMain.Width / 2 + 18 * Screen.TwipsPerPixelX
        NME_H_Vel(Index) = NME_H_Vel(Index) * -0.5
    ElseIf Position = "Right" Then
        NME_X_Pos(Index) = V_Walls(1).x - frmMain.Width / 2 - 18 * Screen.TwipsPerPixelX
        NME_H_Vel(Index) = NME_H_Vel(Index) * -0.5
    End If
End Function


Public Function NME_Fire_Bomb(Index As Integer)
    If Timer - 1 > BombTimer Then
        If BombCounter < 200 And Health > 300 Then
            Health = Health - 300
            NME_BombCounter = NME_BombCounter + 1
            For f = 0 To 200
                With NME_Bombs(f)
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
        NME_BombTimer = Timer
    End If
    
    
End Function


Public Function Draw_Projectile()
    For a = 0 To 200
        If NME_Bombs(a).Active = True Then
            If NME_Bombs(a).Timer = 10000 Then
                NME_Bombs(a).x = X_Pos + frmMain.Width / 2 + (Cos(Look_Angle) * (16) - 8) * Screen.TwipsPerPixelX
                NME_Bombs(a).y = Y_Pos + frmMain.Height / 2 + (Sin(Look_Angle) * (16) - 8) * Screen.TwipsPerPixelY
            End If
            For l = 0 To 1
                If (NME_Bombs(a).x + 16 * Screen.TwipsPerPixelX >= V_Walls(l).x And NME_Bombs(a).x < V_Walls(l).x + V_Walls(l).DC_Pic.Width * Screen.TwipsPerPixelX And NME_Bombs(a).y + 16 * Screen.TwipsPerPixelX > V_Walls(l).Lower_Y And NME_Bombs(a).y < V_Walls(l).Upper_Y + V_Walls(l).DC_Pic.Width * Screen.TwipsPerPixelX) And NME_Explosions(a).Origin_X = -1 Then
                    NME_Explosions(a) = NME_Explosions(201 + 2) 'Bombs(a).Strength)
                    NME_Explosions(a).Dest_X = V_Walls(l).x - (NME_Explosions(a).Width / 2 + 8) * Screen.TwipsPerPixelX
                    NME_Explosions(a).Dest_Y = NME_Bombs(a).y - (NME_Explosions(a).Height / 2 + 8) * Screen.TwipsPerPixelY
                    If l = 0 Then NME_Explosions(a).Dest_X = NME_Explosions(a).Dest_X + V_Walls(l).DC_Pic.Width * Screen.TwipsPerPixelX
                    NME_Bombs(a).X_Vel = 0
                    NME_Bombs(a).Y_Vel = 0
                    NME_Bombs(a).x = NME_Explosions(a).Dest_X
                    NME_Bombs(a).y = NME_Explosions(a).Dest_Y
                    NME_Explosions(a).Origin_X = 0
                    NME_Explosions(a).Origin_Y = 0
                    NME_Bombs(a).Timer = 100
                End If
                If (NME_Bombs(a).y + 16 * Screen.TwipsPerPixelX >= H_Walls(l).y And NME_Bombs(a).y < H_Walls(l).y + H_Walls(l).DC_Pic.Height * Screen.TwipsPerPixelX And NME_Bombs(a).x + 16 * Screen.TwipsPerPixelX > H_Walls(l).Lower_X And NME_Bombs(a).x < H_Walls(l).Upper_X + H_Walls(l).DC_Pic.Height * Screen.TwipsPerPixelX) And NME_Explosions(a).Origin_X = -1 Then
                    NME_Explosions(a) = NME_Explosions(201 + 2) ' Bombs(a).Strength)
                    NME_Explosions(a).Dest_X = NME_Bombs(a).x - (NME_Explosions(a).Width / 2 + 8) * Screen.TwipsPerPixelX
                    NME_Explosions(a).Dest_Y = H_Walls(l).y - (NME_Explosions(a).Height / 2 + 8) * Screen.TwipsPerPixelY
                    If l = 0 Then NME_Explosions(a).Dest_Y = NME_Explosions(a).Dest_Y + H_Walls(l).DC_Pic.Height * Screen.TwipsPerPixelY
                    NME_Bombs(a).X_Vel = 0
                    NME_Bombs(a).Y_Vel = 0
                    NME_Bombs(a).x = NME_Explosions(a).Dest_X
                    NME_Bombs(a).y = NME_Explosions(a).Dest_Y
                    NME_Explosions(a).Origin_X = 0
                    NME_Explosions(a).Origin_Y = 0
                    NME_Bombs(a).Timer = 100
                End If
            Next l
            For l = 0 To 3
                If NME_Bombs(a).x + 16 * Screen.TwipsPerPixelX >= X_Pos(l) + frmMain.Width / 2 And NME_Bombs(a).x <= X_Pos(l) + frmMain.Width / 2 + 32 * Screen.TwipsPerPixelX And NME_Explosions(a).Origin_X = -1 Then
                    If NME_Bombs(a).y + 16 * Screen.TwipsPerPixelY >= Y_Pos(l) + frmMain.Height / 2 And NME_Bombs(a).y <= Y_Pos(l) + frmMain.Height / 2 + 32 * Screen.TwipsPerPixelY Then
                        NME_Explosions(a) = NME_Explosions(201 + 1) ' Bombs(a).Strength)
                        NME_Explosions(a).Dest_X = NME_Bombs(a).x - (Explosions(a).Width / 2 + 8) * Screen.TwipsPerPixelX
                        NME_Explosions(a).Dest_Y = NME_Bombs(a).y - (Explosions(a).Height / 2 + 8) * Screen.TwipsPerPixelY
                        NME_Bombs(a).X_Vel = 0
                        NME_Bombs(a).Y_Vel = 0
                        NME_Bombs(a).x = NME_Explosions(a).Dest_X
                        NME_Bombs(a).y = NME_Explosions(a).Dest_Y
                        NME_Explosions(a).Origin_X = 0
                        NME_Explosions(a).Origin_Y = 0
                        NME_Bombs(a).Timer = 100
                    End If
                End If
            Next l
            If NME_Bombs(a).x > -1000 And NME_Bombs(a).y > -1000 And NME_Bombs(a).x < Map_Width + 1000 And NME_Bombs(a).y < Map_Height + 1000 Then
                '''''''''''''
                If Bombs(a).x - X_Pos > 0 And Bombs(a).x - X_Pos < frmMain.Width And Bombs(a).y - Y_Pos > 0 And Bombs(a).y - Y_Pos < frmMain.Height Then
                    If Explosions(a).Origin_X = -1 Then
                        BitBlt frmMain.HDC, (NME_Bombs(a).x - X_Pos) / Screen.TwipsPerPixelX, (NME_Bombs(a).y - Y_Pos) / Screen.TwipsPerPixelY, 16, 16, DC_Bombs_Mask, NME_Bombs(a).DC_X_Pos, 16 * NME_Bombs(a).Type + 16 * 4 * NME_Bombs(a).Type, vbSrcAnd
                        BitBlt frmMain.HDC, (NME_Bombs(a).x - X_Pos) / Screen.TwipsPerPixelX, (NME_Bombs(a).y - Y_Pos) / Screen.TwipsPerPixelY, 16, 16, DC_Bombs_Sprite, NME_Bombs(a).DC_X_Pos, 16 * NME_Bombs(a).Type + 16 * 4 * NME_Bombs(a).Type, vbSrcPaint
                    Else
                        BitBlt frmMain.HDC, (NME_Explosions(a).Dest_X - X_Pos) / Screen.TwipsPerPixelX, (NME_Explosions(a).Dest_Y - Y_Pos) / Screen.TwipsPerPixelY, NME_Explosions(a).Width, NME_Explosions(a).Height, NME_Explosions(a).Mask, NME_Explosions(a).Origin_X, NME_Explosions(a).Origin_Y, vbSrcAnd
                        BitBlt frmMain.HDC, (NME_Explosions(a).Dest_X - X_Pos) / Screen.TwipsPerPixelX, (NME_Explosions(a).Dest_Y - Y_Pos) / Screen.TwipsPerPixelY, NME_Explosions(a).Width, NME_Explosions(a).Height, NME_Explosions(a).Sprite, NME_Explosions(a).Origin_X, NME_Explosions(a).Origin_Y, vbSrcPaint
                    End If
                End If
                If NME_Explosions(a).Origin_X <> -1 Then
                    NME_Explosions(a).Origin_X = NME_Explosions(a).Origin_X + NME_Explosions(a).Width
                    If NME_Explosions(a).Origin_X = NME_Explosions(a).Width * NME_Explosions(a).Columns Then
                        NME_Explosions(a).Origin_X = 0
                        NME_Explosions(a).Origin_Y = NME_Explosions(a).Origin_Y + NME_Explosions(a).Height
                        If NME_Explosions(a).Origin_Y = NME_Explosions(a).Height * NME_Explosions(a).Rows Then
                            NME_Explosions(a).Origin_X = -1
                            NME_Bombs(a).Timer = 0
                        End If
                    End If
                End If
                
                
                
                NME_Bombs(a).x = NME_Bombs(a).x + NME_Bombs(a).X_Vel
                NME_Bombs(a).y = NME_Bombs(a).y + NME_Bombs(a).Y_Vel
                NME_Bombs(a).DC_X_Pos = NME_Bombs(a).DC_X_Pos + 16
                If NME_Bombs(a).DC_X_Pos = 160 Then NME_Bombs(a).DC_X_Pos = 0
                NME_Bombs(a).Timer = NME_Bombs(a).Timer - 1
                If NME_Bombs(a).Timer <= 0 Then NME_Bombs(a).x = -9999
            Else
                NME_Bombs(a).Active = False
                NME_Explosions(a).Origin_X = -1
                NME_BombCounter = NME_BombCounter - 1
            End If
        End If
        If NME_Bullets(a).Active = True Then
            If NME_Bullets(a).Timer = 10000 Then
                NME_Bullets(a).x = X_Pos + frmMain.Width / 2 + Cos(Look_Angle) * (18 + 2) * Screen.TwipsPerPixelX
                NME_Bullets(a).y = Y_Pos + frmMain.Height / 2 + Sin(Look_Angle) * (18 + 2) * Screen.TwipsPerPixelY
            End If
            If NME_Bullets(a).x > 0 And NME_Bullets(a).y > 0 And NME_Bullets(a).x < Map_Width And NME_Bullets(a).y < Map_Height Then
                If NME_Bullets(a).x - X_Pos > 0 And NME_Bullets(a).x - X_Pos < frmMain.Width And NME_Bullets(a).y - Y_Pos > 0 And NME_Bullets(a).y - Y_Pos < frmMain.Height Then
                    BitBlt frmMain.HDC, (NME_Bullets(a).x - X_Pos) / Screen.TwipsPerPixelX, (NME_Bullets(a).y - Y_Pos) / Screen.TwipsPerPixelY, 5, 5, DC_Bullets_Mask, 0, 20 * NME_Bullets(a).Type + 5 * NME_Bullets(a).Strength, vbSrcAnd
                    BitBlt frmMain.HDC, (NME_Bullets(a).x - X_Pos) / Screen.TwipsPerPixelX, (NME_Bullets(a).y - Y_Pos) / Screen.TwipsPerPixelY, 5, 5, DC_Bullets_Sprite, 0, 20 * NME_Bullets(a).Type + 5 * NME_Bullets(a).Strength, vbSrcPaint
                    'For e = 1 To 3
                    '    BitBlt frmMain.HDC, (Bullets(a).X - Bullets(a).X_Vel * e / 15 - X_Pos) / Screen.TwipsPerPixelX, (Bullets(a).Y - Bullets(a).Y_Vel * e / 15 - Y_Pos) / Screen.TwipsPerPixelY, 5, 5, DC_Bullets_Mask, 5 * e, 20 * Bullets(a).Type + 5 * Bullets(a).Strength, vbSrcAnd
                    '    BitBlt frmMain.HDC, (Bullets(a).X - Bullets(a).X_Vel * e / 15 - X_Pos) / Screen.TwipsPerPixelX, (Bullets(a).Y - Bullets(a).Y_Vel * e / 15 - Y_Pos) / Screen.TwipsPerPixelY, 5, 5, DC_Bullets_Sprite, 5 * e, 20 * Bullets(a).Type + 5 * Bullets(a).Strength, vbSrcPaint
                    'Next e
                End If
                NME_Bullets(a).x = NME_Bullets(a).x + NME_Bullets(a).X_Vel
                NME_Bullets(a).y = NME_Bullets(a).y + NME_Bullets(a).Y_Vel
                NME_Bullets(a).DC_X_Pos = NME_Bullets(a).DC_X_Pos + 4
                If NME_Bullets(a).DC_X_Pos = 16 Then NME_Bullets(a).DC_X_Pos = 0
                NME_Bullets(a).Timer = NME_Bullets(a).Timer - 1
                If NME_Bullets(a).Timer <= 0 Then NME_Bullets(a).x = -99
            Else
                NME_Bullets(a).Active = False
                NME_BulletCounter = NME_BulletCounter - 1
            End If
        End If
    Next a
End Function

Public Function NME_Draw_Turrets()
    For o = 0 To 15
        If Turrets(o).Dest_X - X_Pos + Turrets(o).Width * Screen.TwipsPerPixelX > 0 And Turrets(o).Dest_Y - Y_Pos + Turrets(o).Height * Screen.TwipsPerPixelY > 0 And Turrets(o).Dest_X - X_Pos < frmMain.Width And Turrets(o).Dest_Y - Y_Pos < frmMain.Height Then
            BitBlt frmMain.HDC, (Turrets(o).Dest_X - X_Pos) / Screen.TwipsPerPixelX, (Turrets(o).Dest_Y - Y_Pos) / Screen.TwipsPerPixelY, Turrets(o).Width, Turrets(o).Height, Turrets(o).Sprite, Turrets(o).Origin_X, Turrets(o).Origin_Y, vbSrcCopy
        End If
         Turrets(o).Origin_X = Turrets(o).Origin_X + Turrets(o).Width
        If Turrets(o).Origin_X = Turrets(o).Width * Turrets(o).Columns Then
            Turrets(o).Origin_X = 0
            Turrets(o).Origin_Y = Turrets(o).Origin_Y + Turrets(o).Height
            If Turrets(o).Origin_Y = Turrets(o).Height * Turrets(o).Rows Then
                Turrets(o).Origin_Y = 0
            End If
        End If
    Next o
End Function
