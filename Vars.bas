Attribute VB_Name = "Vars"
'API Declerations
Public Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function LoadImage Lib "User32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long

'''''''''''''''''''''''''
Public Type DC_Picture  ' I made this type after I had coded a bunch of other
    Sprite As Long      ' DC pictures into the program. That is why sometimes
    Mask As Long        ' I use it and other times I don't
    Dest_X As Single    '
    Dest_Y As Single    '
    Origin_X As Integer '
    Origin_Y As Integer '
    Width As Integer    '
    Height As Integer   '
    Rows As Integer     '
    Columns As Integer  '
End Type                '
'''''''''''''''''''''''''

'''''''KEYS''''''''''''''''''''''
Public KQuit As Long            '
Public KUp As Long              '
Public KDown As Long            '
Public KLeft As Long            '
Public KRight As Long           '
Public KAfterBurners As Long    '
Public KBullet As Long          '
Public KBomb As Long            '
'''''''''''''''''''''''''''''''''

'''''''MOVEMENT''''''''''''''''''
Public Max_Set_Vel As Single    ' This is the max set
Public Max_Vel As Single        ' This is the max used in calculations
Public Acceleration As Single   ' This is used to determine how fast a ship can reach Max Vel
Public TurnInc As Integer       ' A num, 0-39, representing the position the ship is facing
Public Roll_Pos As Integer      ' This is used for the roll effect
Public H_Vel As Single          ' The # of twips the background will move on the Y axis
Public V_Vel As Single          ' The # of twips the background will move on the V axis
Public Look_Angle As Double     ' The angle that the ship is looking at (not moving at)
Public Move_Angle As Double     ' I don't think I'm currently using this
Public X_Pos As Single          ' These last 4 are necessary because of
Public Y_Pos As Single          ' the inaccuracies that rounding causes.
Public LookAngleInc As Integer  ' These insure that things stay relatively
Public MoveAngleInc As Integer  ' close to the exact value, no matter how
''''''''''''''''''''''''''''''''' much you change them

''''''''''''''MAP''''''''''''
Public Map_Height As Long   ' These are the size of the background map
Public Map_Width As Long    '
'''''''''''''''''''''''''''''

''''''''Weapons''''''''''''''''''''''
Public Type Bomb                    '
    x As Single                      ' Position with respect to entire map
    y As Single                     '
    X_Vel As Single                 ' # of twips to move in x-axis
    Y_Vel As Single                 ' # of twips to move in y-axis
    Type As Integer                 ' 0 - 3: 0=regular 1=EMP 2=Bouncey 3=Thor's Hammer
    Strength As Integer             ' 0 - 3: 0=red 1=yellow 2=blue 3=purple
    Timer As Long                   ' How many iterations of the bomb will be drawn
    DC_X_Pos As Integer             ' For BitBlt animation effect
    Active As Boolean               ' Has it been fired or no
End Type                            '
                                    '
Public Type Bullet                  '
    x As Single                     '
    y As Single                     '
    X_Vel As Single                 '
    Y_Vel As Single                 '
    Type As Integer                 ' 0 - 1: 0=Regular 1=Bouncey
    Strength As Integer             ' 0 - 3: Same as bombs
    Timer As Long                   '
    DC_X_Pos As Integer             '
    Active As Boolean               '
End Type                            '
                                    '
'Public Type Explosion               '
'    X As Long                       '
'    Y As Long                       '
'    Exp_Type As DC_Picture          '
'End Type                            '
Public Explosions(204) As DC_Picture '
                                    '
Public BombCounter As Integer       ' Makes sure you dont go over the max number of bombs
Public BombTimer As Variant         ' Stores time bomb was fired so rate can be controlled
Public BulletTimer As Variant       ' Same
Public BulletCounter As Integer     ' Same
Public Bombs(200) As Bomb           ' Stores all the bombs
Public Bullets(200) As Bomb         ' Stores all the bullets
'''''''''''''''''''''''''''''''''''''
    
''''''''BACKGROUND'''''''''''''''''''
Public Type Star                    ' With respect to entire background
    x As Long                       '
    y As Long                       '
    Color As Long                   ' Hmmm
    Size As Integer                 ' Form draw size
End Type                            '
Public Num_Stars As Integer         ' This is the max number of start
Public Bright_Star(5000) As Star    ' All of the bright stars are stored here
Public Dim_Star(5000) As Star       ' All dim stars are here
                                    '
Public Type BackGrnd_Pic            ' This type is kinda dumb, I think it made things
    x As Long                       ' harder. Its only good for the background pictures
    y As Long                       '
    Index As Long                   '
End Type                            '
Public BG_Pics(50) As BackGrnd_Pic  '
                                    '
Public V_Wall As DC_Picture         '
Public H_Wall As DC_Picture         '
                                    '
Public Type Vertical_Wall           ' This could be used for vertical walls. Right now I'm
    x As Single                     ' just seeing how this works out, it could either suck
    Lower_Y As Single               ' or work really well, I have no idea yet
    Upper_Y As Single               '
    DC_Pic As DC_Picture            '
End Type                            '
Public V_Walls(1) As Vertical_Wall  '
                                    '
Public Type Horizontal_Wall         ' Same as above
    y As Single                     '
    Lower_X As Single               '
    Upper_X As Single               '
    DC_Pic As DC_Picture            '
End Type                            '
Public H_Walls(1) As Horizontal_Wall '
'''''''''''''''''''''''''''''''''''''

'''''''''IMAGES''''''''''''''''''''''
Public DC_Ships_Mask(7) As Long     ' All the ships B & W
Public DC_Ships_Sprite(7) As Long   ' All the ships Color
Public DC_Bombs_Mask As Long        '
Public DC_Bombs_Sprite As Long      '
Public DC_Bullets_Mask As Long      '
Public DC_Bullets_Sprite As Long    '
Public DC_X_Pos As Integer          '
Public DC_Y_Pos As Integer          '
Public Ship_Type As Integer         '
Public DC_BGs(1 To 10) As Long      '
Public DC_Clear As Long             '
Public DC_Radar As Long             '
Public DC_Exhaust_Mask As Long      '
Public DC_Exhaust_Sprite As Long    '
Public DC_Trail_Mask As Long        '
Public DC_Trail_Sprite As Long      '
Public DC_Tiles As Long             '
Public DC_Explode_Mask(3) As Long   '
Public DC_Explode_Sprite(3) As Long '
Public DC_Health_Bar As Long        '
Public DC_Bar As Long               '
Public DC_Display_Sprite As Long    '
Public DC_Display_Mask As Long      '
Public DC_NRG_Font As Long          '
Public DC_NME_Turret As Long        '
'Public DC_BG01 As Long             ' Burning World
'Public DC_BG02 As Long             ' Sun Behind World
'Public DC_BG03 As Long             ' Lighter, sun behind world
'Public DC_BG04 As Long             ' Half a blue world
'Public DC_BG05 As Long             ' Small, sun behind world
'Public DC_BG06 As Long             ' Red Flare
'Public DC_BG07 As Long             ' White Flare
'Public DC_BG08 As Long             ' Barely Lit Red Planet
'Public DC_BG09 As Long             ' World lit from bottom
'Public DC_BG10 As Long             ' Red Comet
'''''''''''''''''''''''''''''''''''''

Public Type Exhaust_Trail
    x As Long
    y As Long
    DC_X_Pos As Integer
End Type
Public ExhaustCounter As Integer
Public Exhaust(50) As Exhaust_Trail

Public Max_Health
Public Health As Integer    'Your Health
Public Health_Temp As Integer 'I use this in the Display function

Public Radar_Size As Integer    ' Guess what this is for

Public pi As Double ' Here's another brain-buster
