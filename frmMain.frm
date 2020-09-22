VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   12960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   3
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   12960
   ScaleWidth      =   17280
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrEnemy 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2520
      Top             =   5760
   End
   Begin VB.Timer tmrShips 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3840
      Top             =   9360
   End
   Begin VB.PictureBox picShip 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   7
      Left            =   6960
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   7
      Top             =   8760
      Width           =   615
   End
   Begin VB.PictureBox picShip 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   6
      Left            =   6960
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   6
      Top             =   8040
      Width           =   615
   End
   Begin VB.PictureBox picShip 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   5
      Left            =   6960
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   5
      Top             =   7320
      Width           =   615
   End
   Begin VB.PictureBox picShip 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   4
      Left            =   6960
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   4
      Top             =   6600
      Width           =   615
   End
   Begin VB.PictureBox picShip 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   3
      Left            =   6960
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   3
      Top             =   5880
      Width           =   615
   End
   Begin VB.PictureBox picShip 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   2
      Left            =   6960
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   2
      Top             =   5160
      Width           =   615
   End
   Begin VB.PictureBox picShip 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   1
      Left            =   6960
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   1
      Top             =   4440
      Width           =   615
   End
   Begin VB.PictureBox picShip 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   6960
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   0
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblShip 
      BackStyle       =   0  'Transparent
      Caption         =   "Shark"
      BeginProperty Font 
         Name            =   "Metrostyle Extended"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Index           =   7
      Left            =   8280
      TabIndex        =   15
      Top             =   8760
      Width           =   1575
   End
   Begin VB.Label lblShip 
      BackStyle       =   0  'Transparent
      Caption         =   "Lancaster"
      BeginProperty Font 
         Name            =   "Metrostyle Extended"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Index           =   6
      Left            =   8280
      TabIndex        =   14
      Top             =   8040
      Width           =   2775
   End
   Begin VB.Label lblShip 
      BackStyle       =   0  'Transparent
      Caption         =   "Weasel"
      BeginProperty Font 
         Name            =   "Metrostyle Extended"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Index           =   5
      Left            =   8280
      TabIndex        =   13
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label lblShip 
      BackStyle       =   0  'Transparent
      Caption         =   "Terrier"
      BeginProperty Font 
         Name            =   "Metrostyle Extended"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Index           =   4
      Left            =   8280
      TabIndex        =   12
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label lblShip 
      BackStyle       =   0  'Transparent
      Caption         =   "Levathian"
      BeginProperty Font 
         Name            =   "Metrostyle Extended"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Index           =   3
      Left            =   8280
      TabIndex        =   11
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Label lblShip 
      BackStyle       =   0  'Transparent
      Caption         =   "Spider"
      BeginProperty Font 
         Name            =   "Metrostyle Extended"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Index           =   2
      Left            =   8280
      TabIndex        =   10
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label lblShip 
      BackStyle       =   0  'Transparent
      Caption         =   "Javelin"
      BeginProperty Font 
         Name            =   "Metrostyle Extended"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Index           =   1
      Left            =   8280
      TabIndex        =   9
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lblShip 
      BackStyle       =   0  'Transparent
      Caption         =   "Warbird"
      BeginProperty Font 
         Name            =   "Metrostyle Extended"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Index           =   0
      Left            =   8280
      TabIndex        =   8
      Tag             =   "&H00808080&"
      Top             =   3720
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Okay, the reason I started this program was because I wanted to mimick a game called Subspace. Well
'it is actually called Continuum now, but that's besides the point. Anyway, I've reached a point that
'I believe would be a good example for others trying to make their own game. Currently, the only
'things in here that are really worth value are the trig algorithms, which aren't difficult but
'can sometimes be tricky, and the BitBlt examples. I will hopefully continue working on this program
'and imporving it but for now, I hope that it can help people out in some way or another. Oh yeah,
'if you have any questions, just email me at bjorklun@purdue.edu

Private Ship_X As Integer
Private Ship_Y As Integer

Private Sub Form_Load()
    'Loads all of the graphics into memory so they can be used on the fly
    Call Load_Pics
    Ship_Y = 576
     tmrShips.Enabled = True
End Sub

'All the code below is just the first screen you see with the ships. It doesn't contain
'any of the in-game code
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 0 To 7
        lblShip(i).ForeColor = &H808080
    Next i
End Sub

Private Sub lblShip_Click(Index As Integer)
    Ship_Type = Index
    
    For i = 0 To 7
        picShip(i).Visible = False
        lblShip(i).Visible = False
    Next i
    
''''This calls the function that starts the game
    Call Initializer
End Sub

Private Sub lblShip_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For i = 0 To 7
        If i = Index Then
            lblShip(i).ForeColor = &H8000000F
        Else
            lblShip(i).ForeColor = &H808080
        End If
    Next i
End Sub

Private Sub picShip_Click(Index As Integer)
    lblShip_Click (Index)
End Sub

Private Sub tmrShips_Timer()
    For i = 0 To 7
        BitBlt picShip(i).HDC, 0, 0, 36, 36, DC_Ships_Sprite(i), Ship_X, Ship_Y, vbSrcCopy
    Next i
    Ship_X = Ship_X + 36
    If Ship_X = 360 Then
        Ship_X = 0
        Ship_Y = Ship_Y + 36
        If Ship_Y = 720 Then Ship_Y = 576
    End If
    
    If GetKey(vbKeyEscape) <> 0 Then Call Good_Bye
End Sub
