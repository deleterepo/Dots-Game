VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHighscores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Scores"
   ClientHeight    =   3420
   ClientLeft      =   3825
   ClientTop       =   2235
   ClientWidth     =   4005
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4005
   Visible         =   0   'False
   Begin VB.CommandButton cmdBack 
      Caption         =   "Main"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2880
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid fgHighscores 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   11
      Cols            =   3
      FixedCols       =   0
      GridLinesFixed  =   1
      Appearance      =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuMain 
         Caption         =   "&Main"
      End
   End
End
Attribute VB_Name = "frmHighscores"
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
'Form: Highscores
'By: Jordan Sherman
'Date written: June 9 2005
'Purpose: Shows the user the high scores from
'the game.
'==============================================

'Make sure all variables are declared
Option Explicit
Private Sub cmdBack_Click()
    '=============================================
    'Sub program: cmdBack_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Allows the user to go back to the
    'main menu after they view the high scores.
    'Input: None.
    'Output: None.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'Make the high scores form invisible
    frmHighscores.Visible = False
    
    'Make the main menu visible
    frmMain.Visible = True
End Sub
Private Sub Form_Load()
    '=============================================
    'Sub program: cmdBack_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Loads and displays the high score
    'data, such as the player names, scores, and
    'the dates that the scores were achieved.
    'Input: None.
    'Output: Player names, high scores, dates.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'Open the high scores file for input
    Open App.Path & "\highscores.txt" For Input As #1
    
    'Display title for player names
    fgHighscores.Col = 0
    fgHighscores.Row = 0
    fgHighscores.ColWidth(0) = 1800
    fgHighscores.Text = "Player name/Place"
    
    'Display title for scores
    fgHighscores.Col = 1
    fgHighscores.Row = 0
    fgHighscores.ColWidth(1) = 1000
    fgHighscores.Text = "Score"
    
    'Display title for dates
    fgHighscores.Col = 2
    fgHighscores.Row = 0
    fgHighscores.ColWidth(2) = 1000
    fgHighscores.Text = "Date"
    
    'Reset the counter variable
    iCount = 0
    
    'Input all the high score data
    Do Until EOF(1)
        Input #1, stNames(iCount)
        Input #1, iScores(iCount)
        Input #1, stDates(iCount)
        iCount = iCount + 1
    Loop
    
    'Close the file
    Close #1
    
    'Display the high score data
    For iCount = 0 To 2
        fgHighscores.Col = iCount 'Set the column
        'Display player names and places
        If fgHighscores.Col = 0 Then
            For iCount2 = 1 To 10
                fgHighscores.Row = iCount2 'Set the row
                fgHighscores.CellAlignment = 0
                fgHighscores.Text = CStr(iCount2) + ". " + stNames(iCount2 - 1) 'Display name
            Next iCount2
        'Display scores
        ElseIf fgHighscores.Col = 1 Then
            For iCount2 = 1 To 10
                fgHighscores.Row = iCount2 'Set the row
                fgHighscores.Text = CStr(iScores(iCount2 - 1)) 'Display score
            Next iCount2
        'Display dates
        ElseIf fgHighscores.Col = 2 Then
            For iCount2 = 1 To 10
                fgHighscores.Row = iCount2 'Set the row
                fgHighscores.Text = stDates(iCount2 - 1) 'Display date
            Next iCount2
        End If
    Next iCount
End Sub
Private Sub mnuMain_Click()
    '=============================================
    'Sub program: mnuMain_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Goes back to the main menu.
    'Input: None.
    'Output: Main menu.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'Go back to the main menu
    Call cmdBack_Click
End Sub
