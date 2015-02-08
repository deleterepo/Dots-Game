VERSION 5.00
Begin VB.Form frmDotsgame 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dots"
   ClientHeight    =   6795
   ClientLeft      =   930
   ClientTop       =   1080
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9525
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   360
      TabIndex        =   216
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdNewgame 
      Caption         =   "&New Game"
      Height          =   495
      Left            =   360
      TabIndex        =   215
      Top             =   480
      Width           =   2055
   End
   Begin VB.Frame fraFourbyfour 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3015
      Left            =   2760
      TabIndex        =   5
      Top             =   480
      Width           =   3015
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   39
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   38
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   37
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   36
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   30
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   29
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   28
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   27
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   21
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   20
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   19
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   18
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   12
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   11
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   10
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   9
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   1
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   2
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   3
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   1
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   2
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   3
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   4
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   9
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   840
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   10
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   840
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   11
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   840
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   12
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   840
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   13
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   840
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   18
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1560
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   19
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1560
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   20
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1560
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   21
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1560
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   22
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1560
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   27
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2280
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   28
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2280
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   29
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2280
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   30
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2280
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   31
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   30
         Left            =   2280
         TabIndex        =   62
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   29
         Left            =   1560
         TabIndex        =   61
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   28
         Left            =   840
         TabIndex        =   60
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   27
         Left            =   120
         TabIndex        =   59
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   21
         Left            =   2280
         TabIndex        =   58
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   20
         Left            =   1560
         TabIndex        =   57
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   19
         Left            =   840
         TabIndex        =   56
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   18
         Left            =   120
         TabIndex        =   55
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   12
         Left            =   2280
         TabIndex        =   54
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   11
         Left            =   1560
         TabIndex        =   53
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   840
         TabIndex        =   52
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   120
         TabIndex        =   51
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   840
         TabIndex        =   29
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   1560
         TabIndex        =   28
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   2280
         TabIndex        =   27
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame FraSixbysix 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   2760
      TabIndex        =   4
      Top             =   480
      Width           =   4455
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   59
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   4320
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   58
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   4320
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   57
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   4320
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   56
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   4320
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   55
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   4320
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   54
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   4320
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   50
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   49
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   48
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   47
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   46
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   45
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   41
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   40
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   32
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   31
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   23
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   22
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   14
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   13
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   4
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   5
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   6
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   120
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   5
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   120
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   14
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   840
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   15
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   840
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   23
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   1560
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   24
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   1560
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   32
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   2280
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   33
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   2280
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   36
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   3000
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   37
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   3000
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   38
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   3000
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   39
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   3000
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   40
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   3000
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   41
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   3000
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   42
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   3000
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   45
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   3720
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   46
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   3720
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   47
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   3720
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   48
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   3720
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   49
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   3720
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   50
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   3720
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   51
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   3720
         Width           =   135
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   50
         Left            =   3720
         TabIndex        =   126
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   49
         Left            =   3000
         TabIndex        =   125
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   48
         Left            =   2280
         TabIndex        =   124
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   47
         Left            =   1560
         TabIndex        =   123
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   46
         Left            =   840
         TabIndex        =   122
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   45
         Left            =   120
         TabIndex        =   121
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   41
         Left            =   3720
         TabIndex        =   120
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   40
         Left            =   3000
         TabIndex        =   119
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   39
         Left            =   2280
         TabIndex        =   118
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   38
         Left            =   1560
         TabIndex        =   117
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   37
         Left            =   840
         TabIndex        =   116
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   36
         Left            =   120
         TabIndex        =   115
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   32
         Left            =   3720
         TabIndex        =   114
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   31
         Left            =   3000
         TabIndex        =   113
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   23
         Left            =   3720
         TabIndex        =   112
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   22
         Left            =   3000
         TabIndex        =   111
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   14
         Left            =   3720
         TabIndex        =   110
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   13
         Left            =   3000
         TabIndex        =   109
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   3720
         TabIndex        =   86
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   3000
         TabIndex        =   85
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame FraEightbyeight 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   2760
      TabIndex        =   6
      Top             =   480
      Width           =   5895
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   79
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   188
         Top             =   5760
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   78
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   187
         Top             =   5760
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   77
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   186
         Top             =   5760
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   76
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   5760
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   75
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   184
         Top             =   5760
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   74
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   183
         Top             =   5760
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   73
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   182
         Top             =   5760
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   72
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   181
         Top             =   5760
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   70
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   180
         Top             =   5040
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   69
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   179
         Top             =   5040
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   68
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   5040
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   67
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   5040
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   66
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   176
         Top             =   5040
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   65
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   175
         Top             =   5040
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   64
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   174
         Top             =   5040
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   63
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   173
         Top             =   5040
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   61
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   172
         Top             =   4320
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   60
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   171
         Top             =   4320
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   52
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   170
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   51
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   169
         Top             =   3600
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   43
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   168
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   42
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   167
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   34
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   33
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   2160
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   25
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   24
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   16
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   162
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   15
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   6
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdHorizontal 
         Height          =   135
         Index           =   7
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   8
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   156
         Top             =   120
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   7
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   155
         Top             =   120
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   16
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   840
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   17
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   840
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   25
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   1560
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   26
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   1560
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   34
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   150
         Top             =   2280
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   35
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   2280
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   43
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   148
         Top             =   3000
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   44
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   3000
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   52
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   146
         Top             =   3720
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   53
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   145
         Top             =   3720
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   54
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   4440
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   55
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   4440
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   56
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   4440
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   57
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   4440
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   58
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   4440
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   59
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   4440
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   60
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   4440
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   61
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   4440
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   62
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   4440
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   63
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   135
         Top             =   5160
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   64
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   5160
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   65
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   5160
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   66
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   5160
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   67
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   5160
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   68
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   5160
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   69
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   5160
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   70
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   128
         Top             =   5160
         Width           =   135
      End
      Begin VB.CommandButton cmdVertical 
         Height          =   615
         Index           =   71
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   5160
         Width           =   135
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   70
         Left            =   5160
         TabIndex        =   214
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   69
         Left            =   4440
         TabIndex        =   213
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   68
         Left            =   3720
         TabIndex        =   212
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   67
         Left            =   3000
         TabIndex        =   211
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   66
         Left            =   2280
         TabIndex        =   210
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   65
         Left            =   1560
         TabIndex        =   209
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   64
         Left            =   840
         TabIndex        =   208
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   63
         Left            =   120
         TabIndex        =   207
         Top             =   5160
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   61
         Left            =   5160
         TabIndex        =   206
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   60
         Left            =   4440
         TabIndex        =   205
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   59
         Left            =   3720
         TabIndex        =   204
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   58
         Left            =   3000
         TabIndex        =   203
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   57
         Left            =   2280
         TabIndex        =   202
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   56
         Left            =   1560
         TabIndex        =   201
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   55
         Left            =   840
         TabIndex        =   200
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   54
         Left            =   120
         TabIndex        =   199
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   52
         Left            =   5160
         TabIndex        =   198
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   51
         Left            =   4440
         TabIndex        =   197
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   43
         Left            =   5160
         TabIndex        =   196
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   42
         Left            =   4440
         TabIndex        =   195
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   34
         Left            =   5160
         TabIndex        =   194
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   33
         Left            =   4440
         TabIndex        =   193
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   25
         Left            =   5160
         TabIndex        =   192
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   24
         Left            =   4440
         TabIndex        =   191
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   16
         Left            =   5160
         TabIndex        =   190
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   15
         Left            =   4440
         TabIndex        =   189
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   5160
         TabIndex        =   158
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblTag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   4440
         TabIndex        =   157
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Label lblPlayer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblPlayer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label lblPlayer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   480
      TabIndex        =   1
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblPlayer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewgame 
         Caption         =   "&New Game"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuInstructions 
         Caption         =   "&Instructions"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmDotsgame"
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
'Form: Dotsgame
'By: Jordan Sherman
'Date written: June 9 2005
'Purpose: The main game's form.
'==============================================

'Make sure all variables are declared
Option Explicit

'Declare variables
Dim iWinner As Integer            'Stores which player has won the game
Dim iTotalscore As Integer        'Stores the cumulated score of all players
Dim iLastline As Integer          'Stores the index of the last line that was hovered over
Dim iResponse As Integer          'Stores the response from a user to a message box
Dim stDate As String              'Stores the date that the current game was won
Dim bHorizontal As Boolean        'Stores whether the last line hovered over was horizontal or not (vertical)
Dim bWingame As Boolean           'Stores if there was a winner in the last game or not
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
Private Sub cmdHorizontal_Click(Index As Integer)
    '=============================================
    'Sub program: cmdHorizontal_Click
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Changes the colour and disables
    'the horizontal line clicked on, then
    'checks if the line completes a box.
    'Input: None.
    'Output: None.
    'Sub program input: None.
    'Sub program output: Index of control array
    'to the checking sub program.
    '=============================================
    
    'Store that the line was clicked on
    iHorizontal(Index) = 1
    
    'Change the colour of the line to the user's desired colour
    cmdHorizontal(Index).BackColor = vPiececolour(iPlayerturn)
    
    'Check if the line completes any boxes
    Call Check_Horizontal(Index)
    
    'Check the scores to see if the game is over
    Call Check_Scores
    
    'Decide who plays next
    Call Decide_Player_Turn

    'Next player's turn
    bNextplayer = True

    'Disable the button
    cmdHorizontal(Index).Enabled = False
End Sub
Private Sub cmdHorizontal_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '=============================================
    'Sub program: cmdHorizontal_MouseMove
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Changes the colour of the line once
    'it is hovered over.
    'Input: None.
    'Output: The line is coloured.
    'Sub program input: Index, X and Y values.
    'Sub program output: None.
    '=============================================
    
    'Clear the previous line
    Call Clear_Lines
    
    'Change the colour of the line
    If iHorizontal(Index) = 0 And cmdHorizontal(Index).Enabled = True Then
        cmdHorizontal(Index).BackColor = vPiececolour(iPlayerturn)
    End If
    
    'The last line was horizontal
    bHorizontal = True
    
    'Store the last lines' index value
    iLastline = Index
End Sub
Private Sub cmdNewgame_Click()
    '=============================================
    'Sub program: cmdNewgame_Click
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Makes a new grid and
    'clears all the variables for a new game.
    'Input: None.
    'Output: A clear grid.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================

    'Ask the user if they would like to go back to the
    'options screen before starting a new grid
    iResponse = MsgBox("Would you like to reconfigure options?", vbYesNo, "New Game")
    
    'If the response is yes, then restart the game
    'by clearing all the variables and making
    'the options screen visible again
    If iResponse = vbYes Then
        frmDotsgame.Visible = False
        frmOptions.Visible = True
        iNumberofplayers = 2
        iNumberofboxes = 16
        frmDotsgame.fraFourbyfour.Visible = True
        frmDotsgame.FraEightbyeight.Visible = False
        frmDotsgame.FraSixbysix.Visible = False
        frmOptions.fraPlayeroptions.Visible = False
        frmOptions.fraNumberofboxes.Visible = True
        frmOptions.fraNumberofplayers.Visible = True
        frmOptions.cmdStartgame.Enabled = False
        frmOptions.cmdConfigureplayeroptions.Visible = True
        frmOptions.fraInstructions.Visible = False
        frmOptions.Height = 3165
        frmOptions.opt2players.SetFocus
        frmOptions.opt16boxes.SetFocus
        
        'Clear all the lines
        For iCount = 0 To 79
            If iCount <> 8 And iCount <> 17 And iCount <> 26 And iCount <> 35 And iCount <> 44 And iCount <> 53 And iCount <> 62 And iCount <> 71 Then
                cmdHorizontal(iCount).BackColor = &H8000000F
                cmdHorizontal(iCount).Enabled = True
            End If
            If iCount < 72 Then
                cmdVertical(iCount).BackColor = &H8000000F
                cmdVertical(iCount).Enabled = True
            End If
                iHorizontal(iCount) = 0
                iVertical(iCount) = 0
        Next iCount
        
        'Clear all the labels inside the boxes
        For iCount = 0 To 70
            If iCount <> 8 And iCount <> 17 And iCount <> 26 And iCount <> 35 And iCount <> 44 And iCount <> 53 And iCount <> 62 Then
                lblTag(iCount).Caption = ""
            End If
        Next iCount
        
        'Reset the score
        For iCount = 0 To 3
            iScore(iCount) = 0
        Next iCount
    
        'Reset all the labels for the players
        For iCount = 0 To 3
            lblPlayer(iCount).Caption = ""
            lblPlayer(iCount).BorderStyle = 0
        Next iCount
        
        'Reset other variables
        bNextplayer = True
        iPlayerturn = 0
        lblPlayer(0).BorderStyle = 1
    'If the response is no
    ElseIf iResponse = vbNo Then
        'Clear all the lines
        For iCount = 0 To 79
            If iCount <> 8 And iCount <> 17 And iCount <> 26 And iCount <> 35 And iCount <> 44 And iCount <> 53 And iCount <> 62 And iCount <> 71 Then
                cmdHorizontal(iCount).BackColor = &H8000000F
                cmdHorizontal(iCount).Enabled = True
            End If
            If iCount < 72 Then
                cmdVertical(iCount).BackColor = &H8000000F
                cmdVertical(iCount).Enabled = True
            End If
                iHorizontal(iCount) = 0
                iVertical(iCount) = 0
        Next iCount
        
        'Clear all the tags inside the game
        For iCount = 0 To 70
            If iCount <> 8 And iCount <> 17 And iCount <> 26 And iCount <> 35 And iCount <> 44 And iCount <> 53 And iCount <> 62 Then
                lblTag(iCount).Caption = ""
            End If
        Next iCount
        
        'Clear the scores
        For iCount = 0 To 3
            iScore(iCount) = 0
        Next iCount
        
        'Clear the player identifier labels
        For iCount = 0 To 3
                lblPlayer(iCount).BorderStyle = 0
        Next iCount
        
        'Reset other variables
        bNextplayer = True
        iPlayerturn = 0
        lblPlayer(0).BorderStyle = 1
    End If
End Sub
Private Sub cmdVertical_Click(Index As Integer)
    '=============================================
    'Sub program: cmdVertical_Click
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Changes the colour and disables
    'the vertical line clicked on, then
    'checks if the line completes a box.
    'Input: None.
    'Output: None.
    'Sub program input: None.
    'Sub program output: Index of control array
    'to the checking sub program.
    '=============================================
    
    'Store that the line was clicked on
    iVertical(Index) = 1
    
    'Change the colour of the line to the user's desired colour
    cmdVertical(Index).BackColor = vPiececolour(iPlayerturn)
    
    'Check if the line completes any boxes
    Call Check_Vertical(Index)

    'Check the scores to see if the game is over
    Call Check_Scores
    
    'Decide who plays next
    Call Decide_Player_Turn

    'Next player's turn
    bNextplayer = True

    'Disable the button
    cmdVertical(Index).Enabled = False
End Sub
Private Sub cmdVertical_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '=============================================
    'Sub program: cmdVertical_MouseMove
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Changes the colour of the line once
    'it is hovered over.
    'Input: None.
    'Output: The line is coloured.
    'Sub program input: Index, X and Y values.
    'Sub program output: None.
    '=============================================
    
    'Clear the previous line
    Call Clear_Lines
    
    'Change the colour of the line
    If iVertical(Index) = 0 And cmdVertical(Index).Enabled = True Then
        cmdVertical(Index).BackColor = vPiececolour(iPlayerturn)
    End If

    'The last line was vertical
    bHorizontal = False
    
    'Store the last lines index value
    iLastline = Index
End Sub
Private Sub Form_Load()
    '=============================================
    'Sub program: Form_Load
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Sets certain variables on load.
    'Input: None.
    'Output: None.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'Set certain variables
    bNextplayer = True
    iPlayerturn = 0
    
    'Make the first person's turn
    lblPlayer(0).BorderStyle = 1
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '=============================================
    'Sub program: Form_MouseMove
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Clears the last line hovered over.
    'Input: None.
    'Output: Clears the last line hovered.
    'Sub program input: Index, X and Y values.
    'Sub program output: None.
    '=============================================
    
    'Clear the last line
    Call Clear_Lines
End Sub
Private Sub FraEightbyeight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '=============================================
    'Sub program: FraEightbyeight_MouseMove
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Clears the last line hovered over.
    'Input: None.
    'Output: Clears the last line hovered.
    'Sub program input: Index, X and Y values.
    'Sub program output: None.
    '=============================================
    
    'Clear the last line
    Call Clear_Lines
End Sub
Private Sub fraFourbyfour_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '=============================================
    'Sub program: FraFourbyfour_MouseMove
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Clears the last line hovered over.
    'Input: None.
    'Output: Clears the last line hovered.
    'Sub program input: Index, X and Y values.
    'Sub program output: None.
    '=============================================
    
    'Clear the last line
    Call Clear_Lines
End Sub
Private Sub FraSixbysix_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '=============================================
    'Sub program: FraSixbysix_MouseMove
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Clears the last line hovered over.
    'Input: None.
    'Output: Clears the last line hovered.
    'Sub program input: Index, X and Y values.
    'Sub program output: None.
    '=============================================
    
    'Clear the last line
    Call Clear_Lines
End Sub
Private Sub lblTag_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '=============================================
    'Sub program: lblTag_MouseMove
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Clears the last line hovered over.
    'Input: None.
    'Output: Clears the last line hovered.
    'Sub program input: Index, X and Y values.
    'Sub program output: None.
    '=============================================
    
    'Clear the last line
    Call Clear_Lines
End Sub
Private Sub Check_Horizontal(Index As Integer)
    '=============================================
    'Sub program: Check_Horizontal
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Checks if the last horizontal line
    'clicked on completes any boxes.
    'Input: None.
    'Output: Puts the players initial in any
    'completed boxes.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================

    'If the lines' index is larger or equal to 9
    If Index >= 9 Then
        'If all lines below the line add up then make a box
        If iHorizontal(Index) + iHorizontal(Index - 9) + iVertical(Index - 9) + iVertical(Index - 8) = 4 Then
            bNextplayer = False 'Still the players turn
            iScore(iPlayerturn) = iScore(iPlayerturn) + 1 'Add one to the score
            
            'Change the caption inside the label of the box
            lblTag(Index - 9).ForeColor = vPiececolour(iPlayerturn)
            lblTag(Index - 9) = " " + Mid(stPlayername(iPlayerturn), 1, 1)
        Else
            bNextplayer = True 'Next player's turn
        End If
    End If
    
    'If the lines above the line add up then make a box
    If iHorizontal(Index) + iHorizontal(Index + 9) + iVertical(Index) + iVertical(Index + 1) = 4 Then
        bNextplayer = False 'Still player's turn
        iScore(iPlayerturn) = iScore(iPlayerturn) + 1 'Add one to the score
        
        'Change the caption inside the label of the box
        lblTag(Index).ForeColor = vPiececolour(iPlayerturn)
        lblTag(Index) = " " + Mid(stPlayername(iPlayerturn), 1, 1)
    End If

End Sub
Private Sub Check_Vertical(Index As Integer)
    '=============================================
    'Sub program: Check_Vertical
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Checks if the last horizontal line
    'clicked on completes any boxes.
    'Input: None.
    'Output: Puts the players initial in any
    'completed boxes.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================

     'If the line is not on the left side
     If Index Mod 9 <> 0 Then
        'If the lines to the right of it add up then make a box
        If iHorizontal(Index + 8) + iHorizontal(Index - 1) + iVertical(Index) + iVertical(Index - 1) = 4 Then
            bNextplayer = False 'Still player's turn
            iScore(iPlayerturn) = iScore(iPlayerturn) + 1 'Add one to the score

            'Change the caption inside the label of the box
            lblTag(Index - 1).ForeColor = vPiececolour(iPlayerturn)
            lblTag(Index - 1) = " " + Mid(stPlayername(iPlayerturn), 1, 1)
        Else
            bNextplayer = True 'Next player's turn
        End If
     End If

    'If the lines to the left and around it add up then make a box
     If iHorizontal(Index) + iHorizontal(Index + 9) + iVertical(Index) + iVertical(Index + 1) = 4 Then
        bNextplayer = False 'Still players' turn
        iScore(iPlayerturn) = iScore(iPlayerturn) + 1 'Add one to the score

        'Change the caption inside the label of the box
        lblTag(Index).ForeColor = vPiececolour(iPlayerturn)
        lblTag(Index) = " " + Mid(stPlayername(iPlayerturn), 1, 1)
     End If
End Sub
Private Sub Check_Scores()
    '=============================================
    'Sub program: Check_Scores
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Checks if the game is over or not,
    'and if so check if the winner's score is a
    'high score.
    'Input: None.
    'Output: Ask if the player wants to view the
    'high scores.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'Declare variables
    Dim iCount As Integer   'Counter variable for counted loops
    
    'Reset the cumulated score
    iTotalscore = 0
    
    'Calculate and store the cumulated score
    For iCount2 = 0 To 3
        iTotalscore = iTotalscore + iScore(iCount2)
    Next iCount2

    'If the cumulated score is equal to the number of boxes
    'in the grid then
    If iTotalscore = iNumberofboxes Then
        'Find the winner
        For iCount2 = iNumberofplayers - 2 To 0 Step -1
            For iCount3 = 0 To iCount2
                If iScore(iCount3) > iScore(iCount3 + 1) Then
                    iWinner = iCount3
                    bWingame = True
                ElseIf iScore(iCount3) <= iScore(iCount3 + 1) Then
                    bWingame = False
                End If
            Next iCount3
        Next iCount2
        
        If bWingame = True Then 'If there is a winner
            'Reset the counter variable
            iCount = 0
            
            'Open the high scores file for input
            Open App.Path & "\highscores.txt" For Input As #1
            
            'Input other high scores
            Do Until EOF(1)
                Input #1, stNames(iCount)
                Input #1, iScores(iCount)
                Input #1, stDates(iCount)
                iCount = iCount + 1
            Loop
            
            'Close the file
            Close #1
            
            'See if the winner's score is higher or equal
            'to any other high scores
            For iCount3 = 0 To 9
                'If the score is higher than any of the high scores
                If iScore(iWinner) >= iScores(iCount3) Then
                    iScores(iCount3) = iScore(iWinner)
                    stDates(iCount3) = Date$
                    stNames(iCount3) = stPlayername(iWinner)
                    iCount3 = 9
                End If
            Next iCount3
    
            'Open the high scores file for output
            Open App.Path & "\highscores.txt" For Output As #1
            
            'Store all the high scores in the file
            For iCount = 0 To 9
                Write #1, stNames(iCount)
                Write #1, iScores(iCount)
                Write #1, stDates(iCount)
            Next iCount
            
            Close #1 'Close the file
        
            'Depending on the number of players, display who has won
            If iNumberofplayers = 4 Then
                iResponse = MsgBox(CStr(stPlayername(iWinner) + " wins the game!" + vbNewLine + "Scores:" + vbNewLine + stPlayername(0) + " - " + CStr(iScore(0)) + vbNewLine + stPlayername(1) + " - " + CStr(iScore(1)) + vbNewLine + stPlayername(2) + " - " + CStr(iScore(2)) + vbNewLine + stPlayername(3) + " - " + CStr(iScore(3)) + vbNewLine + "View high scores?"), vbYesNo, "Game over")
            ElseIf iNumberofplayers = 3 Then
                iResponse = MsgBox(CStr(stPlayername(iWinner) + " wins the game!" + vbNewLine + "Scores:" + vbNewLine + stPlayername(0) + " - " + CStr(iScore(0)) + vbNewLine + stPlayername(1) + " - " + CStr(iScore(1)) + vbNewLine + stPlayername(2) + " - " + CStr(iScore(2)) + vbNewLine + "View high scores?"), vbYesNo, "Game over")
            ElseIf iNumberofplayers = 2 Then
                iResponse = MsgBox(CStr(stPlayername(iWinner) + " wins the game!" + vbNewLine + "Scores:" + vbNewLine + stPlayername(0) + " - " + CStr(iScore(0)) + vbNewLine + stPlayername(1) + " - " + CStr(iScore(1)) + vbNewLine + "View high scores?"), vbYesNo, "Game over")
            End If
        
            'If the player answers yes to viewing high scores
            'then show them the high scores
            If iResponse = vbYes Then
                frmDotsgame.Visible = False
                frmHighscores.Visible = True
            End If
        ElseIf bWingame = False Then 'If the game is a tie
            'Display that the game was a tie and player scores
            If iNumberofplayers = 2 And iScore(0) = iScore(1) Then
                iResponse = MsgBox(("Tie game!" + vbNewLine + "Scores:" + vbNewLine + stPlayername(0) + " - " + CStr(iScore(0)) + vbNewLine + stPlayername(1) + " - " + CStr(iScore(1)) + vbNewLine + "View high scores?"), vbYesNo, "Game over")
            ElseIf iNumberofplayers = 3 And iScore(0) = iScore(1) And iScore(1) = iScore(2) Then
                iResponse = MsgBox(("Tie game!" + vbNewLine + "Scores:" + vbNewLine + stPlayername(0) + " - " + CStr(iScore(0)) + vbNewLine + stPlayername(1) + " - " + CStr(iScore(1)) + vbNewLine + stPlayername(2) + " - " + CStr(iScore(2)) + vbNewLine + "View high scores?"), vbYesNo, "Game over")
            ElseIf iNumberofplayers = 4 And iScore(0) = iScore(1) And iScore(0) = iScore(2) And iScore(0) = iScore(3) Then
                iResponse = MsgBox(("Tie game!" + vbNewLine + "Scores:" + vbNewLine + stPlayername(0) + " - " + CStr(iScore(0)) + vbNewLine + stPlayername(1) + " - " + CStr(iScore(1)) + vbNewLine + stPlayername(2) + " - " + CStr(iScore(2)) + vbNewLine + stPlayername(3) + " - " + CStr(iScore(3)) + vbNewLine + "View high scores?"), vbYesNo, "Game over")
            End If
                
            'If the player answers yes to viewing high scores
            'then show them the high scores
            If iResponse = vbYes Then
                frmDotsgame.Visible = False
                frmHighscores.Visible = True
            End If
        End If
    End If
End Sub
Private Sub Decide_Player_Turn()
    '=============================================
    'Sub program: Decide_Player_Turn
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Decides if the player goes next or
    'not.
    'Input: None.
    'Output: None.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'If a box was not completed then go to the next player
    If bNextplayer = True Then
        If iPlayerturn = iNumberofplayers - 1 Then
            iPlayerturn = 0
        Else
            iPlayerturn = iPlayerturn + 1
        End If
    End If
    
    'Clear all the player identifier labels
    For iCount = 0 To 3
            lblPlayer(iCount).BorderStyle = 0
    Next iCount
    
    'Show the current player's turn
    lblPlayer(iPlayerturn).BorderStyle = 1
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
Private Sub mnuInstructions_Click()
    '=============================================
    'Sub program: mnuInstructions_Click
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Starts a new game.
    'Input: None.
    'Output: Instructions on the game.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'Display instructions for the game
    Call MsgBox("-The object of the game is to use four lines to complete as many boxes as possible without giving your opponent/s chances to complete boxes" + vbNewLine + "-Once a box is formed, the initial of the player's name who had completed the box will show up in the middle of that box" + vbNewLine + "-One point is awarded every time a box is completed" + vbNewLine + "-If a box is completed, the player who completed the box gets to go again" + vbNewLine + "-If a box is not completed, the next player gets to make their turn", vbInformation, "Instructions")
End Sub
Private Sub mnuNewgame_Click()
    '=============================================
    'Sub program: mnuNewgame_Click
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Starts a new game.
    'Input: None.
    'Output: A new game.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'Start a new game.
    Call cmdNewgame_Click
End Sub
Private Sub Clear_Lines()
    '=============================================
    'Sub program: Clear_Lines
    'Written by: Jordan Sherman
    'Date written: June 9 2005
    'Purpose: Clears the previous line that was
    'hovered over.
    'Input: None.
    'Output: Greys out the last line that was
    'moused over.
    'Sub program input: None.
    'Sub program output: None.
    '=============================================
    
    'If the last line was horizontal then
    If bHorizontal = True Then
        'If the last line was enabled then clear the line
        If cmdHorizontal(iLastline).Enabled = True Then
            cmdHorizontal(iLastline).BackColor = &H8000000F
        End If
    'If the last line was vertical then
    ElseIf bHorizontal = False Then
        'If the last line was enabled then clear the line
        If cmdVertical(iLastline).Enabled = True Then
            cmdVertical(iLastline).BackColor = &H8000000F
        End If
    End If
End Sub
