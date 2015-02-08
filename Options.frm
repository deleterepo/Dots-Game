VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game options"
   ClientHeight    =   2790
   ClientLeft      =   3600
   ClientTop       =   1950
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4485
   Visible         =   0   'False
   Begin VB.CommandButton cmdStartgame 
      Caption         =   "Start Game"
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      TabIndex        =   20
      Top             =   2880
      Width           =   2535
   End
   Begin VB.CommandButton cmdConfigureplayeroptions 
      Caption         =   "Configure Player Options"
      Height          =   495
      Left            =   960
      TabIndex        =   19
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Frame fraNumberofboxes 
      Caption         =   "Game size"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   4215
      Begin VB.OptionButton opt64boxes 
         Caption         =   "64 Boxes"
         Height          =   375
         Left            =   2880
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opt36boxes 
         Caption         =   "36 Boxes"
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opt16boxes 
         Caption         =   "16 Boxes"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame fraNumberofplayers 
      Caption         =   "Number of players"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton opt4players 
         Caption         =   "4"
         Height          =   495
         Left            =   2876
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton opt3players 
         Caption         =   "3"
         Height          =   495
         Left            =   1976
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton opt2players 
         Caption         =   "2"
         Height          =   495
         Left            =   1113
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame fraPlayeroptions 
      Caption         =   "Player 1"
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.OptionButton optColour 
         BackColor       =   &H00FF8080&
         Height          =   255
         Index           =   9
         Left            =   3120
         TabIndex        =   24
         Top             =   1680
         Width           =   615
      End
      Begin VB.OptionButton optColour 
         BackColor       =   &H00FFFF00&
         Height          =   255
         Index           =   8
         Left            =   2400
         TabIndex        =   23
         Top             =   1680
         Width           =   615
      End
      Begin VB.OptionButton optColour 
         BackColor       =   &H000080FF&
         Height          =   255
         Index           =   7
         Left            =   1680
         TabIndex        =   22
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdDone 
         Caption         =   "Done"
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
      End
      Begin VB.OptionButton optColour 
         BackColor       =   &H00FF00FF&
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   13
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtPlayername 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Text            =   "Player 1"
         Top             =   600
         Width           =   3135
      End
      Begin VB.OptionButton optColour 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   960
         TabIndex        =   11
         Top             =   1680
         Width           =   615
      End
      Begin VB.OptionButton optColour 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   615
      End
      Begin VB.OptionButton optColour 
         BackColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   9
         Top             =   1320
         Width           =   615
      End
      Begin VB.OptionButton optColour 
         BackColor       =   &H0000FF00&
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   8
         Top             =   1320
         Width           =   615
      End
      Begin VB.OptionButton optColour 
         BackColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   1320
         Width           =   615
      End
      Begin VB.OptionButton optColour 
         BackColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Name"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Piece colour"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
   End
   Begin VB.Frame fraInstructions 
      Caption         =   "Instructions"
      Height          =   2655
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.Label Label3 
         Caption         =   $"Options.frx":0000
         Height          =   2295
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmOptions"
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
'Form: Options
'By: Jordan Sherman
'Date written: June 9 2005
'Purpose: Lets the user choose how many players
'there are, how many boxes there are, the
'player's names, and their colours.
'==============================================

'Make sure all variables are declared
Option Explicit

'Declare variables
Dim iDefaultcount As Integer        'Counter for the default colour
Dim iLastcolourclicked As Integer   'Stores the last colour that the user chose
Dim iPlayercount As Integer         'Counter for players
Private Sub cmdConfigureplayeroptions_Click()
    '=============================================
    'Sub program: cmdConfigureplayeroptions_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Lets the user configure options
    'such as their name, and their colour.
    'Input: None.
    'Output: The player options form.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'Declare variables
    Dim iCount As Integer   'Counter for counted loops
    fraNumberofplayers.Visible = False 'Make the number of players frame invisible
    fraNumberofboxes.Visible = False   'Make the number of boxes frame invisible
    fraPlayeroptions.Visible = True    'Make the player options frame visible
    iLastcolourclicked = 0 'Change the last colour clicked to the first
    iDefaultcount = 0 'Make the first colour the default one
    iPlayercount = 0 'Reset the player counter variable
        
    'Change the caption of the player options frame
    fraPlayeroptions.Caption = "Player " + CStr(iPlayercount + 1)
    
    'Change the text boxes' caption to the next player
    txtPlayername.Text = "Player " + CStr(iPlayercount + 1)

    cmdConfigureplayeroptions.Visible = False 'Make the player options button invisible
    frmOptions.Height = 3990 'Change the height of the form
    
    For iCount = 0 To 9
        optColour(iCount).Enabled = True 'Enable all the colour options
    Next iCount
    
    optColour(0).SetFocus 'Set the focus on the first colour
End Sub
Private Sub cmdDone_Click()
    '=============================================
    'Sub program: cmdDone_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Finishes the players options.
    'Input: None.
    'Output: Next player's options, or a new game.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    stPlayername(iPlayercount) = txtPlayername.Text 'Store the player's name in a variable
    
    'If the last player has entered their options
    If iPlayercount + 1 = iNumberofplayers Then
        cmdStartgame.Enabled = True 'Enable the start game button
        frmOptions.Height = 3990 'Change the height of the form
        fraPlayeroptions.Visible = False 'Make the player options frame invisible
        fraInstructions.Visible = True 'Show the user instructions
    'Otherwise if there are still players who need to input their options
    ElseIf iPlayercount <> 4 Then
        optColour(iLastcolourclicked).Enabled = False 'Disable the user's chose colour
        iPlayercount = iPlayercount + 1 'Add one to the player counter
        iDefaultcount = iDefaultcount + 1 'Change the default colour
        optColour(iDefaultcount).Value = True 'Set the new default colour
    Else
        optColour(iLastcolourclicked).Enabled = False 'Otherwise disable the last clicked colour
    End If
    
    'Change the caption of the player options frame
    fraPlayeroptions.Caption = "Player " + CStr(iPlayercount + 1)
    
    'Change the text boxes' caption to the next player
    txtPlayername.Text = "Player " + CStr(iPlayercount + 1)

End Sub
Private Sub cmdStartgame_Click()
    '=============================================
    'Sub program: cmdStartgame_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Starts the new game.
    'Input: None.
    'Output: Starts the new game.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'Declare variables
    Dim iCount As Integer   'Counter for counted loops

    frmOptions.Visible = False 'Make the options form invisible
    frmDotsgame.Visible = True 'Make the game form visible
    
    'If any players have the same name
    If stPlayername(0) = stPlayername(1) Or stPlayername(1) = stPlayername(2) Or stPlayername(0) = stPlayername(2) Then
            For iCount = 0 To 2
                'Number each player
                stPlayername(iCount) = stPlayername(iCount) + CStr(iCount + 1)
            Next iCount
    End If
    
    'Open the settings file for output
    Open App.Path & "\settings.txt" For Output As #1
    
    'Write the number of players
    Write #1, iNumberofplayers
    
    'Write the number of boxes
    Write #1, iNumberofboxes
        
    For iCount = 1 To iNumberofplayers
        Write #1, stPlayername(iCount - 1) 'Write player's name
        Write #1, vPiececolour(iCount - 1) 'Write player's colour
    Next iCount
    
    'Close the file
    Close #1
    
    'Change the labels of the player identifiers to player names
    For iCount = 0 To iNumberofplayers - 1
        frmDotsgame.lblPlayer(iCount).ForeColor = vPiececolour(iCount) 'Set the labels colour to the player's colour
        frmDotsgame.lblPlayer(iCount) = stPlayername(iCount) 'Set the label caption to the player name
    Next iCount
End Sub
Private Sub Form_Load()
    '=============================================
    'Sub program: Form_Load
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Resets variables and displays the
    'smallest grid.
    'Input: None.
    'Output: Show the smallest grid.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'Set the default number of players and boxes
    iNumberofplayers = 2
    iNumberofboxes = 16
    iPlayercount = 0
    iLastcolourclicked = 0
    iDefaultcount = 0
    
    'Make the 16 box grid visible
    frmDotsgame.fraFourbyfour.Visible = True
    frmDotsgame.FraEightbyeight.Visible = False
    frmDotsgame.FraSixbysix.Visible = False
End Sub
Private Sub opt16boxes_Click()
    '=============================================
    'Sub program: opt16boxes_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Displays the 16 box grid.
    'Input: None.
    'Output: Displays the 16 box grid.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'Display the sixteen box grid
    frmDotsgame.fraFourbyfour.Visible = True
    frmDotsgame.FraEightbyeight.Visible = False
    frmDotsgame.FraSixbysix.Visible = False
    
    'Store the number of boxes
    iNumberofboxes = 16
End Sub
Private Sub opt36boxes_Click()
    '=============================================
    'Sub program: opt16boxes_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Displays the 36 box grid.
    'Input: None.
    'Output: Displays the 36 box grid.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'Display the 36 box grid
    frmDotsgame.fraFourbyfour.Visible = True
    frmDotsgame.FraEightbyeight.Visible = False
    frmDotsgame.FraSixbysix.Visible = True
    
    'Store the number of boxes
    iNumberofboxes = 36
End Sub
Private Sub opt3players_Click()
    '=============================================
    'Sub program: opt3players_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Stores the number of players in
    'a variable.
    'Input: None.
    'Output: None.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================

    'Store the number of players
    iNumberofplayers = 3
End Sub
Private Sub opt4players_Click()
    '=============================================
    'Sub program: opt3players_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Stores the number of players in
    'a variable.
    'Input: None.
    'Output: None.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================

    'Store the number of players
    iNumberofplayers = 4
End Sub
Private Sub opt64boxes_Click()
    '=============================================
    'Sub program: opt16boxes_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Displays the 64 box grid.
    'Input: None.
    'Output: Displays the 64 box grid.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'Display the 64 box grid
    frmDotsgame.fraFourbyfour.Visible = True
    frmDotsgame.FraEightbyeight.Visible = True
    frmDotsgame.FraSixbysix.Visible = True
    
    'Store the number of boxes
    iNumberofboxes = 64
End Sub
Private Sub optColour_Click(Index As Integer)
    '=============================================
    'Sub program: optColour_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Stores the number of players in
    'a variable.
    'Input: None.
    'Output: None.
    'Sub program input: The index of the control.
    'Sub program output: None.
    '=============================================
    
    'Store the index in a variable
    iLastcolourclicked = Index
    
    'Set the users colour according to what they
    'clicked on
    Select Case Index
        Case 0
            vPiececolour(iPlayercount) = vbRed
        Case 1
            vPiececolour(iPlayercount) = vbBlue
        Case 2
            vPiececolour(iPlayercount) = vbGreen
        Case 3
            vPiececolour(iPlayercount) = vbYellow
        Case 4
            vPiececolour(iPlayercount) = vbMagenta
        Case 5
            vPiececolour(iPlayercount) = vbBlack
        Case 6
            vPiececolour(iPlayercount) = vbWhite
        Case 7
            vPiececolour(iPlayercount) = &H80FF&  'Orange
        Case 8
            vPiececolour(iPlayercount) = &HFFFF00 'Teal
        Case 9
            vPiececolour(iPlayercount) = &HFF8080 'Purple
        End Select
End Sub
Private Sub opt2players_Click()
    '=============================================
    'Sub program: opt3players_Click
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Stores the number of players in
    'a variable.
    'Input: None.
    'Output: None.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================

    'Store the number of players
    iNumberofplayers = 2
End Sub

Private Sub txtPlayername_Change()
    '=============================================
    'Sub program: txtPlayername_Change
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Makes sure there are no quotes in
    'a text box.
    'Input: None.
    'Output: A string with no quotes.
    'Sub program input: None.
    'Sub program output: The text in the player
    'name text box.
    '=============================================
    
    'Make sure there are no quotes in the string
    Call DisallowQuotes(txtPlayername)
End Sub

Private Sub txtPlayername_GotFocus()
    '=============================================
    'Sub program: txtPlayername_GotFocus
    'Written by: Jordan Sherman
    'Date written: June 10 2005
    'Purpose: Clears the text box once it it's
    'clicked on.
    'Input: None.
    'Output: A cleared text box.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================

    txtPlayername.Text = "" 'Clear the player name textbox
End Sub

Private Sub DisallowQuotes(ByRef txtTextBox As TextBox, Optional boConvertToSingleQuote As Boolean = False)
    '--------------------------------------
    'Sub program:        DisallowQuotes
    'Written by:         Seth Climans
    'Date Written:       June 11, 2005
    'Purpose:            Make sure they don't enter Quotation Marks,
    'since they mess w/
    '                    the saveing and loading due to quotes in the file
    'Input:              Text from the textbox we're working with
    'Output:             Inability to enter quotes
    'Sub program Input : Textbox to work with
    'Sub program Output: none
    '--------------------------------------
    
    'If boConvertToSingleQuote is true, the doublequote is converted to a single
    'If it's false, it is simply deleted
    
    Dim iCount As Integer
    Dim iCursorPosition As Integer 'Where the cursor/blinking line/ is
    Dim stInput As String 'The string in the textbox
    
    'Assign the cursor, and the input text
    iCursorPosition = txtTextBox.SelStart
    stInput = txtTextBox.Text
    
    'Go through each letter and look for a quote
    For iCount = 1 To Len(stInput)
        'For some reason iCount doesn't stop at Len(stInput) if it changes and becomes less
        'so we must include this if statement
        If iCount <= Len(stInput) Then
            If Asc(Mid(stInput, iCount, 1)) = 34 Then 'Ascii 34 is adouble quote
                If boConvertToSingleQuote Then
                    'Remove the double quote and replace it with a single
                    stInput = Mid(stInput, 1, iCount - 1) & "'" & Mid(stInput, iCount + 1, Len(stInput) - 1)
                Else
                    'Trim out the invalid character using mid commands
                    stInput = Mid(stInput, 1, iCount - 1) & Mid(stInput, iCount + 1, Len(stInput) - 1)
                    'We've removed a character, so we must move the cursor left
                    iCursorPosition = iCursorPosition - 1
                End If
            End If
        End If
    Next iCount
    
    'Assign the changed text
    txtTextBox.Text = stInput
    
    'To ensure no errors, make sure the cursor position is valid, then assign it
    If iCursorPosition >= 0 Then
        txtTextBox.SelStart = iCursorPosition
    End If
End Sub


