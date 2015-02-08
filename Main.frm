VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   2640
   ClientLeft      =   4485
   ClientTop       =   2460
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   2655
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CommandButton cmdHighscores 
      Caption         =   "&High Scores"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdNewgame 
      Caption         =   "&New Game"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   3  'Dot
      Height          =   495
      Left            =   840
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "....  D ots...."
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   2295
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Program: Final Assignment - Dots Game
'By: Jordan Sherman
'Date written: June 9 2005
'Purpose: Lets the users play a game of Dots,
'for up to 4 users.
'==============================================

'==============================================
'Form: Main
'By: Jordan Sherman
'Date written: June 9 2005
'Purpose: Lets the user make a new game, see
'the high scores, or exit.
'==============================================

'Make sure all variables are declared
Option Explicit
Private Sub cmdExit_Click()
    '=============================================
    'Sub program: cmdExit_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Ends the program.
    'Input: None.
    'Output: None.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'End the program
    End
End Sub
Private Sub cmdHighscores_Click()
    '=============================================
    'Sub program: cmdHighscores_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Displays the high score form.
    'Input: None.
    'Output: High score form.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'Make the high scores form visible
    frmHighscores.Visible = True
    
    'Make the main form invisible
    frmMain.Visible = False
End Sub
Private Sub cmdNewgame_Click()
    '=============================================
    'Sub program: cmdNewgame_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Starts a new game, asking the user
    'if they want to use their previous settings,
    'or make new settings.
    'Input: None.
    'Output: A new game.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'Declare variables
    Dim iResponse As Integer    'Stores the response of the user to a question
    Dim iCount As Integer       'Counter variable for counted loops
    
    'Ask the user if they want to use their previously chosen settings
    iResponse = MsgBox("Would you like to revert to your previous settings?", vbYesNo, "New Game")
    
    'If their response is yes
    If iResponse = vbYes Then
        'Open the settings file for input
        Open App.Path & "\settings.txt" For Input As #1
        
        'Input the number of players
        Input #1, iNumberofplayers
        
        'Input the number of boxes
        Input #1, iNumberofboxes
        
        For iCount = 1 To iNumberofplayers
            Input #1, stPlayername(iCount - 1) 'Input player name
            Input #1, vPiececolour(iCount - 1) 'Input player colour
        Next iCount
        
        'Close the file
        Close #1
        
        'Display the corresponding grid according to the number
        'of boxes chosen
        If iNumberofboxes = 16 Then
            frmDotsgame.fraFourbyfour.Visible = True    'Make the 16 box grid visible
            frmDotsgame.FraEightbyeight.Visible = False 'Make the 36 box grid invisible
            frmDotsgame.FraSixbysix.Visible = False     'Make the 64 box grid invisible
        ElseIf iNumberofboxes = 36 Then
            frmDotsgame.fraFourbyfour.Visible = True    'Make the 16 box grid visible
            frmDotsgame.FraSixbysix.Visible = True      'Make the 36 box grid visible
            frmDotsgame.FraEightbyeight.Visible = False 'Make the 64 box grid invisible
        ElseIf iNumberofboxes = 64 Then
            frmDotsgame.fraFourbyfour.Visible = True    'Make the 16 box grid visible
            frmDotsgame.FraEightbyeight.Visible = True  'Make the 36 box grid visible
            frmDotsgame.FraSixbysix.Visible = True      'Make the 64 box grid visible
        End If
        
        frmMain.Visible = False 'Make the main form invisible
        frmDotsgame.Visible = True 'Make the game form visible
        
        'Clear all the lines
        For iCount = 0 To 79
            If iCount <> 8 And iCount <> 17 And iCount <> 26 And iCount <> 35 And iCount <> 44 And iCount <> 53 And iCount <> 62 And iCount <> 71 Then
                frmDotsgame.cmdHorizontal(iCount).BackColor = &H8000000F
                frmDotsgame.cmdHorizontal(iCount).Enabled = True
            End If
            If iCount < 72 Then
                frmDotsgame.cmdVertical(iCount).BackColor = &H8000000F
                frmDotsgame.cmdVertical(iCount).Enabled = True
            End If
                iHorizontal(iCount) = 0
                iVertical(iCount) = 0
        Next iCount
        
        'Clear all the labels inside the boxes
        For iCount = 0 To 70
            If iCount <> 8 And iCount <> 17 And iCount <> 26 And iCount <> 35 And iCount <> 44 And iCount <> 53 And iCount <> 62 Then
                frmDotsgame.lblTag(iCount).Caption = ""
            End If
        Next iCount
        
        'Reset the score
        For iCount = 0 To 3
            iScore(iCount) = 0
        Next iCount
    
        'Reset all the labels for the players
        For iCount = 0 To 3
            frmDotsgame.lblPlayer(iCount).Caption = ""
            frmDotsgame.lblPlayer(iCount).BorderStyle = 0
        Next iCount
        
        'Change the labels of the player identifiers to player names
        For iCount = 0 To iNumberofplayers - 1
            frmDotsgame.lblPlayer(iCount).ForeColor = vPiececolour(iCount) 'Set the labels colour to the player's colour
            frmDotsgame.lblPlayer(iCount) = stPlayername(iCount) 'Set the label caption to the player name
        Next iCount
        
        'Reset other variables
        bNextplayer = True
        iPlayerturn = 0
        frmDotsgame.lblPlayer(0).BorderStyle = 1
        
    'If the answer is no
    ElseIf iResponse = vbNo Then
        frmOptions.Visible = True 'Make the options form visible
        frmMain.Visible = False   'Make the main form invisible
    End If
End Sub
Private Sub mnuAbout_Click()
    '=============================================
    'Sub program: mnuAbout_Click
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Displays the games creator, date
    'written, and it's purpose.
    'Input: None.
    'Output: Game's creator, date finished, and
    'purpose.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'Display information about the game
    Call MsgBox("Dots Game" + vbNewLine + "By: Jordan Sherman" + vbNewLine + "Date written: June 11 2005" + vbNewLine + "Purpose: Lets users play a game of Dots for up to 4 players.", vbInformation, "About Dots")
End Sub
Private Sub mnuExit_Click()
    '=============================
    'Sub program: mnuExit_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Ends the program.
    'Input: None.
    'Output: None.
    'Sub program input: None.
    'Sub program output: None.
    '=============================
    
    'End the program
    End
End Sub
