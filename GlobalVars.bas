Attribute VB_Name = "GlobalVars"
'Declare global variables
Global iNumberofplayers As Integer  'Stores the number of players
Global vPiececolour(3) As Variant   'Stores each player's colour
Global stPlayername(3) As String    'Stores each player's name
Global iNumberofboxes As Integer    'Stores the number of boxes
Global stDates(9) As String         'Stores the date a high score was achieved
Global stNames(9) As String         'Stores high score names
Global iScores(9) As Integer        'Stores high scores
Global iCount As Integer            'Counter for counted loops
Global iCount2 As Integer           'Counter for counted loops
Global iCount3 As Integer           'Counter for counted loops
Global iHorizontal(88) As Integer   'Stores if a horizontal line has been clicked on
Global iVertical(88) As Integer     'Stores if a vertical line has been clicked on
Global iPlayerturn As Integer       'Stores which player's turn it is
Global iScore(3) As Integer         'Stores each player's score
Global bNextplayer As Boolean       'Stores if it is the next player's turn or not
