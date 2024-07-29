VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " Fuzzy Calculator : Derajat Keanggotaan"
   ClientHeight    =   12045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23430
   Icon            =   "ADK.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12045
   ScaleWidth      =   23430
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text61 
      Height          =   375
      Left            =   20280
      TabIndex        =   116
      Top             =   9360
      Width           =   615
   End
   Begin VB.TextBox Text60 
      Height          =   375
      Left            =   19680
      TabIndex        =   115
      Top             =   9360
      Width           =   615
   End
   Begin VB.TextBox Text59 
      Height          =   375
      Left            =   19080
      TabIndex        =   114
      Top             =   9360
      Width           =   615
   End
   Begin VB.TextBox Text58 
      Height          =   375
      Left            =   18480
      TabIndex        =   113
      Top             =   9360
      Width           =   615
   End
   Begin VB.TextBox Text57 
      Height          =   375
      Left            =   17880
      TabIndex        =   112
      Top             =   9360
      Width           =   615
   End
   Begin VB.TextBox Text56 
      Height          =   375
      Left            =   15600
      TabIndex        =   111
      Top             =   9360
      Width           =   2055
   End
   Begin VB.TextBox Text55 
      Height          =   375
      Left            =   20280
      TabIndex        =   110
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text54 
      Height          =   375
      Left            =   19680
      TabIndex        =   109
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text53 
      Height          =   375
      Left            =   19080
      TabIndex        =   108
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text52 
      Height          =   375
      Left            =   18480
      TabIndex        =   107
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text51 
      Height          =   375
      Left            =   17880
      TabIndex        =   106
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text50 
      Height          =   375
      Left            =   15600
      TabIndex        =   105
      Top             =   7440
      Width           =   2055
   End
   Begin VB.TextBox Text49 
      Height          =   375
      Left            =   20280
      TabIndex        =   104
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text48 
      Height          =   375
      Left            =   19680
      TabIndex        =   103
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text47 
      Height          =   375
      Left            =   19080
      TabIndex        =   102
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text46 
      Height          =   375
      Left            =   18480
      TabIndex        =   101
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text45 
      Height          =   375
      Left            =   17880
      TabIndex        =   100
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text44 
      Height          =   375
      Left            =   15600
      TabIndex        =   99
      Top             =   5520
      Width           =   2055
   End
   Begin VB.TextBox Text43 
      Height          =   375
      Left            =   18120
      TabIndex        =   98
      Top             =   3600
      Width           =   4215
   End
   Begin VB.TextBox Text42 
      Height          =   375
      Left            =   17400
      TabIndex        =   97
      Top             =   3000
      Width           =   4935
   End
   Begin VB.TextBox Text41 
      Height          =   375
      Left            =   12720
      TabIndex        =   96
      Top             =   9360
      Width           =   615
   End
   Begin VB.TextBox Text40 
      Height          =   375
      Left            =   12120
      TabIndex        =   95
      Top             =   9360
      Width           =   615
   End
   Begin VB.TextBox Text39 
      Height          =   375
      Left            =   11520
      TabIndex        =   94
      Top             =   9360
      Width           =   615
   End
   Begin VB.TextBox Text38 
      Height          =   375
      Left            =   10920
      TabIndex        =   93
      Top             =   9360
      Width           =   615
   End
   Begin VB.TextBox Text37 
      Height          =   375
      Left            =   10320
      TabIndex        =   92
      Top             =   9360
      Width           =   615
   End
   Begin VB.TextBox Text36 
      Height          =   375
      Left            =   8040
      TabIndex        =   91
      Top             =   9360
      Width           =   2055
   End
   Begin VB.TextBox Text35 
      Height          =   375
      Left            =   12720
      TabIndex        =   90
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text34 
      Height          =   375
      Left            =   12120
      TabIndex        =   89
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text33 
      Height          =   375
      Left            =   11520
      TabIndex        =   88
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text32 
      Height          =   375
      Left            =   10920
      TabIndex        =   87
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text31 
      Height          =   375
      Left            =   10320
      TabIndex        =   86
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text30 
      Height          =   375
      Left            =   8040
      TabIndex        =   85
      Top             =   7440
      Width           =   2055
   End
   Begin VB.TextBox Text29 
      Height          =   375
      Left            =   12720
      TabIndex        =   84
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text28 
      Height          =   375
      Left            =   12120
      TabIndex        =   83
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text27 
      Height          =   405
      Left            =   11520
      TabIndex        =   82
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text26 
      Height          =   375
      Left            =   10920
      TabIndex        =   81
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text25 
      Height          =   375
      Left            =   10320
      TabIndex        =   79
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text24 
      Height          =   375
      Left            =   8040
      TabIndex        =   78
      Top             =   5520
      Width           =   2055
   End
   Begin VB.TextBox Text23 
      Height          =   375
      Left            =   10680
      TabIndex        =   77
      Top             =   3600
      Width           =   4215
   End
   Begin VB.TextBox Text22 
      Height          =   375
      Left            =   9960
      TabIndex        =   76
      Top             =   3000
      Width           =   4935
   End
   Begin VB.TextBox Text21 
      Height          =   375
      Left            =   5160
      TabIndex        =   75
      Top             =   9360
      Width           =   615
   End
   Begin VB.TextBox Text20 
      Height          =   375
      Left            =   4560
      TabIndex        =   74
      Top             =   9360
      Width           =   615
   End
   Begin VB.OptionButton Option36 
      Caption         =   "Trapesium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   21120
      TabIndex        =   73
      Top             =   9840
      Width           =   1575
   End
   Begin VB.OptionButton Option35 
      Caption         =   "Segitiga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   21120
      TabIndex        =   72
      Top             =   9480
      Width           =   1575
   End
   Begin VB.OptionButton Option34 
      Caption         =   "Linear Turun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   21120
      TabIndex        =   71
      Top             =   9120
      Width           =   1575
   End
   Begin VB.OptionButton Option33 
      Caption         =   "Linear Naik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   21120
      TabIndex        =   70
      Top             =   8760
      Width           =   1575
   End
   Begin VB.OptionButton Option32 
      Caption         =   "Trapesium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   21120
      TabIndex        =   69
      Top             =   7920
      Width           =   1575
   End
   Begin VB.OptionButton Option31 
      Caption         =   "Segitiga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   21120
      TabIndex        =   68
      Top             =   7560
      Width           =   1575
   End
   Begin VB.OptionButton Option30 
      Caption         =   "Linear Turun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   21120
      TabIndex        =   67
      Top             =   7200
      Width           =   1575
   End
   Begin VB.OptionButton Option29 
      Caption         =   "Linear Naik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   21120
      TabIndex        =   66
      Top             =   6840
      Width           =   1575
   End
   Begin VB.OptionButton Option28 
      Caption         =   "Trapesium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   21120
      TabIndex        =   65
      Top             =   6000
      Width           =   1575
   End
   Begin VB.OptionButton Option27 
      Caption         =   "Segitiga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   21120
      TabIndex        =   64
      Top             =   5640
      Width           =   1575
   End
   Begin VB.OptionButton Option26 
      Caption         =   "Linear Turun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   21120
      TabIndex        =   63
      Top             =   5280
      Width           =   1575
   End
   Begin VB.OptionButton Option25 
      Caption         =   "Linear Naik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   21120
      TabIndex        =   62
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   3960
      TabIndex        =   58
      Top             =   9360
      Width           =   615
   End
   Begin VB.TextBox Text18 
      Height          =   375
      Left            =   3360
      TabIndex        =   57
      Top             =   9360
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   20400
      TabIndex        =   54
      Top             =   11040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   21840
      TabIndex        =   53
      Top             =   11040
      Width           =   1215
   End
   Begin VB.OptionButton Option24 
      Caption         =   "Trapesium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      TabIndex        =   52
      Top             =   9840
      Width           =   1575
   End
   Begin VB.OptionButton Option23 
      Caption         =   "Segitga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      TabIndex        =   51
      Top             =   9480
      Width           =   1575
   End
   Begin VB.OptionButton Option22 
      Caption         =   "Linear Turun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      TabIndex        =   50
      Top             =   9120
      Width           =   1575
   End
   Begin VB.OptionButton Option21 
      Caption         =   "Linear Naik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      TabIndex        =   49
      Top             =   8760
      Width           =   1575
   End
   Begin VB.OptionButton Option20 
      Caption         =   "Trapesium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      TabIndex        =   48
      Top             =   7920
      Width           =   1575
   End
   Begin VB.OptionButton Option19 
      Caption         =   "Segitiga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      TabIndex        =   47
      Top             =   7560
      Width           =   1575
   End
   Begin VB.OptionButton Option18 
      Caption         =   "Linear Turun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      TabIndex        =   46
      Top             =   7200
      Width           =   1575
   End
   Begin VB.OptionButton Option17 
      Caption         =   "Linear Naik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      TabIndex        =   45
      Top             =   6840
      Width           =   1575
   End
   Begin VB.OptionButton Option16 
      Caption         =   "Trapesium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      TabIndex        =   44
      Top             =   6000
      Width           =   1575
   End
   Begin VB.OptionButton Option15 
      Caption         =   "Segitiga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      TabIndex        =   43
      Top             =   5640
      Width           =   1575
   End
   Begin VB.OptionButton Option14 
      Caption         =   "Linear Turun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      TabIndex        =   42
      Top             =   5280
      Width           =   1575
   End
   Begin VB.OptionButton Option13 
      Caption         =   "Linear Naik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      TabIndex        =   41
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   2760
      TabIndex        =   40
      Top             =   9360
      Width           =   615
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   480
      TabIndex        =   39
      Top             =   9360
      Width           =   2055
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   5160
      TabIndex        =   38
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   4560
      TabIndex        =   37
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   3960
      TabIndex        =   36
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   3360
      TabIndex        =   35
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   2760
      TabIndex        =   31
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   480
      TabIndex        =   29
      Top             =   7440
      Width           =   2055
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   5160
      TabIndex        =   27
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   4560
      TabIndex        =   26
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   3960
      TabIndex        =   25
      Top             =   5520
      Width           =   615
   End
   Begin VB.OptionButton Option12 
      Caption         =   "Trapesium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   24
      Top             =   9840
      Width           =   1575
   End
   Begin VB.OptionButton Option11 
      Caption         =   "Segitiga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   23
      Top             =   9480
      Width           =   1575
   End
   Begin VB.OptionButton Option10 
      Caption         =   "Linear Turun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   22
      Top             =   9120
      Width           =   1575
   End
   Begin VB.OptionButton Option9 
      Caption         =   "Linear Naik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   21
      Top             =   8760
      Width           =   1575
   End
   Begin VB.OptionButton Option8 
      Caption         =   "Trapesium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   20
      Top             =   7920
      Width           =   1575
   End
   Begin VB.OptionButton Option7 
      Caption         =   "Segitiga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   19
      Top             =   7560
      Width           =   1575
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Linear Turun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   18
      Top             =   7200
      Width           =   1575
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Linear Naik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   17
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3360
      TabIndex        =   16
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2760
      TabIndex        =   15
      Top             =   5520
      Width           =   615
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Trapesium"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   14
      Top             =   6000
      Width           =   1575
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Segitiga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   13
      Top             =   5640
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Linear Turun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   12
      Top             =   5280
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Linear Naik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   5520
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   3120
      TabIndex        =   6
      Top             =   3600
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   3000
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   2040
      Width           =   8175
   End
   Begin VB.Label Label18 
      Caption         =   "®vtw"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15240
      TabIndex        =   80
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label17 
      Caption         =   "Fungsi Keanggotaan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   21240
      TabIndex        =   61
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Domain"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   19200
      TabIndex        =   60
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "Himpunan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   16320
      TabIndex        =   59
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "Semesta Pembicaraan :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15840
      TabIndex        =   56
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Line Line5 
      X1              =   360
      X2              =   23040
      Y1              =   10680
      Y2              =   10680
   End
   Begin VB.Label Label13 
      Caption         =   "Nama Variabel :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15840
      TabIndex        =   55
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Fungsi Keanggotaan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      TabIndex        =   34
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Domain"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   33
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Himpunan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   32
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Semesta Pembicaraan :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   30
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Nama Variabel :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   28
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Fungsi Keanggotaan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Domain"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Himpunan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Semesta Pembicaraan :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Line Line4 
      X1              =   23040
      X2              =   23040
      Y1              =   2880
      Y2              =   10680
   End
   Begin VB.Line Line3 
      X1              =   15480
      X2              =   15480
      Y1              =   2760
      Y2              =   10680
   End
   Begin VB.Line Line2 
      X1              =   7920
      X2              =   7920
      Y1              =   2760
      Y2              =   10680
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   360
      Y1              =   2760
      Y2              =   10680
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Variabel :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Kasus :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "FUZZY CALCULATOR"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      Top             =   840
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If MsgBox("Apakah Anda Yakin Ingin Keluar?", vbExclamation + vbYesNo + vbDefaultButton2, "Konfirmasi") = vbYes Then End
End Sub

Sub Reset()
    Text1.Text = Empty
    Text2.Text = Empty
    Text3.Text = Empty
    Text4.Text = Empty
    Text5.Text = Empty
    Text6.Text = Empty
    Text7.Text = Empty
    Text8.Text = Empty
    Text9.Text = Empty
    Text10.Text = Empty
    Text11.Text = Empty
    Text12.Text = Empty
    Text13.Text = Empty
    Text14.Text = Empty
    Text15.Text = Empty
    Text16.Text = Empty
    Text17.Text = Empty
    Text18.Text = Empty
    Text19.Text = Empty
    Text20.Text = Empty
    Text21.Text = Empty
    Text22.Text = Empty
    Text23.Text = Empty
    Text24.Text = Empty
    Text25.Text = Empty
    Text26.Text = Empty
    Text27.Text = Empty
    Text28.Text = Empty
    Text29.Text = Empty
    Text30.Text = Empty
    Text31.Text = Empty
    Text32.Text = Empty
    Text33.Text = Empty
    Text34.Text = Empty
    Text35.Text = Empty
    Text36.Text = Empty
    Text37.Text = Empty
    Text38.Text = Empty
    Text39.Text = Empty
    Text40.Text = Empty
    Text41.Text = Empty
    Text42.Text = Empty
    Text43.Text = Empty
    Text44.Text = Empty
    Text45.Text = Empty
    Text46.Text = Empty
    Text47.Text = Empty
    Text48.Text = Empty
    Text49.Text = Empty
    Text50.Text = Empty
    Text51.Text = Empty
    Text52.Text = Empty
    Text53.Text = Empty
    Text54.Text = Empty
    Text55.Text = Empty
    Text56.Text = Empty
    Text57.Text = Empty
    Text58.Text = Empty
    Text59.Text = Empty
    Text60.Text = Empty
    Text61.Text = Empty
End Sub

Private Sub Command2_Click()
    Reset
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        If Text9.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text9.SetFocus
        Else
        Form2.Show
        Form2.Label8 = Text4.Text
        End If
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        If Text5.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text5.SetFocus
        Else
        Form3.Show
        Form3.Label8 = Text4.Text
        End If
    End If
End Sub
Private Sub Option3_Click()
    If Option3.Value = True Then
        If Text5.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text5.SetFocus
        Else
        Form4.Show
        Form4.Label8 = Text4.Text
        End If
    End If
End Sub
Private Sub Option4_Click()
    If Option4.Value = True Then
        If Text5.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text5.SetFocus
        Else
        Form5.Show
        Form5.Label8 = Text4.Text
        End If
    End If
End Sub
Private Sub Option5_Click()
    If Option5.Value = True Then
        If Text15.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text15.SetFocus
        Else
        Form6.Show
        Form6.Label8 = Text10.Text
        End If
    End If
End Sub
Private Sub Option6_Click()
    If Option6.Value = True Then
        If Text15.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text15.SetFocus
        Else
        Form7.Show
        Form7.Label8 = Text10.Text
        End If
    End If
End Sub
Private Sub Option7_Click()
    If Option7.Value = True Then
        If Text15.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text15.SetFocus
        Else
        Form8.Show
        Form8.Label8 = Text10.Text
        End If
    End If
End Sub
Private Sub Option8_Click()
    If Option8.Value = True Then
If Text15.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text15.SetFocus
        Else
        Form9.Show
        Form9.Label8 = Text10.Text
        End If
    End If
End Sub
Private Sub Option9_Click()
    If Option9.Value = True Then
        If Text21.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text21.SetFocus
        Else
        Form10.Show
        Form10.Label8 = Text16.Text
        End If
    End If
End Sub
Private Sub Option10_Click()
    If Option10.Value = True Then
        If Text21.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text21.SetFocus
        Else
        Form11.Show
        Form11.Label8 = Text16.Text
        End If
    End If
End Sub
Private Sub Option11_Click()
    If Option11.Value = True Then
        If Text21.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text21.SetFocus
        Else
        Form12.Show
        Form12.Label8 = Text16.Text
        End If
    End If
End Sub
Private Sub Option12_Click()
    If Option12.Value = True Then
        If Text21.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text21.SetFocus
        Else
        Form13.Show
        Form13.Label8 = Text16.Text
        End If
    End If
End Sub
Private Sub Option13_Click()
    If Option13.Value = True Then
        If Text29.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text29.SetFocus
        Else
        Form14.Show
        Form14.Label8 = Text24.Text
        End If
    End If
End Sub
Private Sub Option14_Click()
    If Option14.Value = True Then
        If Text29.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text29.SetFocus
        Else
        Form15.Show
        Form15.Label8 = Text24.Text
        End If
    End If
End Sub
Private Sub Option15_Click()
    If Option15.Value = True Then
        If Text29.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text29.SetFocus
        Else
        Form16.Show
        Form16.Label8 = Text24.Text
        End If
    End If
End Sub
Private Sub Option16_Click()
    If Option16.Value = True Then
        If Text29.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text29.SetFocus
        Else
        Form17.Show
        Form17.Label8 = Text24.Text
        End If
    End If
End Sub
Private Sub Option17_Click()
    If Option17.Value = True Then
        If Text35.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text35.SetFocus
        Else
        Form18.Show
        Form18.Label8 = Text30.Text
        End If
    End If
End Sub
Private Sub Option18_Click()
    If Option18.Value = True Then
        If Text35.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text35.SetFocus
        Else
        Form19.Show
        Form19.Label8 = Text30.Text
        End If
    End If
End Sub
Private Sub Option19_Click()
    If Option19.Value = True Then
        If Text35.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text35.SetFocus
        Else
        Form20.Show
        Form20.Label8 = Text30.Text
        End If
    End If
End Sub
Private Sub Option20_Click()
    If Option20.Value = True Then
        If Text35.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text35.SetFocus
        Else
        Form21.Show
        Form21.Label8 = Text30.Text
        End If
    End If
End Sub
Private Sub Option21_Click()
    If Option21.Value = True Then
        If Text41.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text41.SetFocus
        Else
        Form22.Show
        Form22.Label8 = Text36.Text
        End If
    End If
End Sub
Private Sub Option22_Click()
    If Option22.Value = True Then
        If Text41.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text41.SetFocus
        Else
        Form23.Show
        Form23.Label8 = Text36.Text
        End If
    End If
End Sub
Private Sub Option23_Click()
    If Option23.Value = True Then
        If Text41.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text41.SetFocus
        Else
        Form24.Show
        Form24.Label8 = Text36.Text
        End If
    End If
End Sub
Private Sub Option24_Click()
    If Option24.Value = True Then
        If Text41.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text41.SetFocus
        Else
        Form25.Show
        Form25.Label8 = Text36.Text
        End If
    End If
End Sub
Private Sub Option25_Click()
    If Option25.Value = True Then
        If Text49.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text49.SetFocus
        Else
        Form26.Show
        Form26.Label8 = Text44.Text
        End If
    End If
End Sub
Private Sub Option26_Click()
    If Option26.Value = True Then
        If Text49.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text49.SetFocus
        Else
        Form27.Show
        Form27.Label8 = Text44.Text
        End If
    End If
End Sub
Private Sub Option27_Click()
    If Option27.Value = True Then
        If Text49.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text49.SetFocus
        Else
        Form28.Show
        Form28.Label8 = Text44.Text
        End If
    End If
End Sub
Private Sub Option28_Click()
    If Option28.Value = True Then
        If Text49.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text49.SetFocus
        Else
        Form29.Show
        Form29.Label8 = Text44.Text
        End If
    End If
End Sub
Private Sub Option29_Click()
    If Option29.Value = True Then
        If Text55.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text55.SetFocus
        Else
        Form30.Show
        Form30.Label8 = Text50.Text
        End If
    End If
End Sub
Private Sub Option30_Click()
    If Option30.Value = True Then
        If Text55.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text55.SetFocus
        Else
        Form31.Show
        Form31.Label8 = Text50.Text
        End If
    End If
End Sub
Private Sub Option31_Click()
    If Option31.Value = True Then
        If Text55.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text55.SetFocus
        Else
        Form32.Show
        Form32.Label8 = Text50.Text
        End If
    End If
End Sub
Private Sub Option32_Click()
    If Option32.Value = True Then
        If Text55.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text55.SetFocus
        Else
        Form33.Show
        Form33.Label8 = Text50.Text
        End If
    End If
End Sub
Private Sub Option33_Click()
    If Option33.Value = True Then
        If Text61.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text61.SetFocus
        Else
        Form34.Show
        Form34.Label8 = Text56.Text
        End If
    End If
End Sub

Private Sub Option34_Click()
    If Option34.Value = True Then
        If Text61.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text61.SetFocus
        Else
        Form35.Show
        Form35.Label8 = Text56.Text
        End If
    End If
End Sub

Private Sub Option35_Click()
    If Option35.Value = True Then
        If Text61.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text61.SetFocus
        Else
        Form36.Show
        Form36.Label8 = Text56.Text
        End If
    End If
End Sub

Private Sub Option36_Click()
    If Option36.Value = True Then
        If Text61.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text61.SetFocus
        Else
        Form37.Show
        Form37.Label8 = Text56.Text
        End If
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2.SetFocus
    End If
End Sub

Private Sub Text2_GotFocus()
    If Text1.Text = Empty Then
        MsgBox "Mohon Inputkan Nama Kasus"
        Text1.SetFocus
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text3.SetFocus
    End If
End Sub

Private Sub Text3_GotFocus()
    If Text2.Text = Empty Then
        MsgBox "Mohon Inputkan Nama Variabel"
        Text2.SetFocus
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text4.SetFocus
    End If
End Sub
Private Sub Text4_GotFocus()
    If Text3.Text = Empty Then
        MsgBox "Mohon Inputkan Semesta Pembicaraan"
        Text3.SetFocus
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text5.SetFocus
    End If
End Sub

Private Sub Text5_GotFocus()
    If Text4.Text = Empty Then
        MsgBox "Mohon Inputkan Nama Himpunan"
        Text4.SetFocus
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text6.SetFocus
    End If
End Sub
Private Sub Text6_GotFocus()
    If Text5.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text5.SetFocus
    End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text7.SetFocus
    End If
End Sub
Private Sub Text7_GotFocus()
    If Text6.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text6.SetFocus
    End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text8.SetFocus
    End If
End Sub
Private Sub Text8_GotFocus()
    If Text7.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text7.SetFocus
    End If
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text9.SetFocus
    End If
End Sub

Private Sub Text9_GotFocus()
    If Text8.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text8.SetFocus
    End If
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text9.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text9.SetFocus
        Else
        Text10.SetFocus
        End If
    End If
End Sub

Private Sub Text10_GotFocus()
    If Text9.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text9.SetFocus
    End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text11.SetFocus
    End If
End Sub

Private Sub Text11_GotFocus()
    If Text10.Text = Empty Then
        MsgBox "Mohon Inputkan Nama Himpunan"
        Text10.SetFocus
    End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text12.SetFocus
    End If
End Sub

Private Sub Text12_GotFocus()
    If Text11.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text11.SetFocus
    End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text13.SetFocus
    End If
End Sub

Private Sub Text13_GotFocus()
    If Text12.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text12.SetFocus
    End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text14.SetFocus
    End If
End Sub
Private Sub Text14_GotFocus()
    If Text13.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text13.SetFocus
    End If
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text15.SetFocus
    End If
End Sub
Private Sub Text15_GotFocus()
    If Text14.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text14.SetFocus
    End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text15.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text15.SetFocus
        Else
        Text16.SetFocus
        End If
    End If
End Sub
Private Sub Text16_GotFocus()
    If Text15.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text15.SetFocus
    End If
End Sub
Private Sub Text16_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text17.SetFocus
    End If
End Sub

Private Sub Text17_GotFocus()
    If Text16.Text = Empty Then
        MsgBox "Mohon Inputkan Nama Himpunan"
        Text16.SetFocus
    End If
End Sub
Private Sub Text17_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text18.SetFocus
    End If
End Sub
Private Sub Text18_GotFocus()
    If Text17.Text = Empty Then
        MsgBox "Mohon Inputkan Domain"
        Text17.SetFocus
    End If
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text19.SetFocus
    End If
End Sub

Private Sub Text19_GotFocus()
    If Text18.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text18.SetFocus
    End If
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text20.SetFocus
    End If
End Sub

Private Sub Text20_GotFocus()
    If Text19.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text19.SetFocus
    End If
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text21.SetFocus
    End If
End Sub

Private Sub Text21_GotFocus()
    If Text20.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text20.SetFocus
    End If
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text21.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text21.SetFocus
        End If
    End If
End Sub
Private Sub Text22_GotFocus()
    If Text1.Text = Empty Then
        MsgBox "Mohon Inputkan Nama Kasus"
        Text1.SetFocus
    End If
End Sub
Private Sub Text22_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text23.SetFocus
    End If
End Sub

Private Sub Text23_GotFocus()
    If Text22.Text = Empty Then
        MsgBox "Mohon Inputkan Nama Variabel"
        Text22.SetFocus
    End If
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text24.SetFocus
    End If
End Sub

Private Sub Text24_GotFocus()
    If Text23.Text = Empty Then
        MsgBox "Mohon Inputkan Semesta Pembicaraan"
        Text23.SetFocus
    End If
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text25.SetFocus
    End If
End Sub

Private Sub Text25_GotFocus()
    If Text24.Text = Empty Then
        MsgBox "Mohon Inputkan Nama Himpunan"
        Text24.SetFocus
    End If
End Sub

Private Sub Text25_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text26.SetFocus
    End If
End Sub
Private Sub Text26_GotFocus()
    If Text25.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text25.SetFocus
    End If
End Sub
Private Sub Text26_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text27.SetFocus
    End If
End Sub
Private Sub Text27_GotFocus()
    If Text26.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text26.SetFocus
    End If
End Sub

Private Sub Text27_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text28.SetFocus
    End If
End Sub
Private Sub Text28_GotFocus()
    If Text27.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text27.SetFocus
    End If
End Sub
Private Sub Text28_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text29.SetFocus
    End If
End Sub

Private Sub Text29_GotFocus()
    If Text28.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text28.SetFocus
    End If
End Sub
Private Sub Text29_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text29.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text29.SetFocus
        Else
        Text30.SetFocus
        End If
    End If
End Sub

Private Sub Text30_GotFocus()
    If Text29.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text29.SetFocus
    End If
End Sub

Private Sub Text30_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text31.SetFocus
    End If
End Sub

Private Sub Text31_GotFocus()
    If Text30.Text = Empty Then
        MsgBox "Mohon Inputkan Nama Himpunan"
        Text30.SetFocus
    End If
End Sub

Private Sub Text31_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text32.SetFocus
    End If
End Sub

Private Sub Text32_GotFocus()
    If Text31.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text31.SetFocus
    End If
End Sub

Private Sub Text32_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text33.SetFocus
    End If
End Sub

Private Sub Text33_GotFocus()
    If Text32.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text32.SetFocus
    End If
End Sub

Private Sub Text33_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text34.SetFocus
    End If
End Sub
Private Sub Text34_GotFocus()
    If Text33.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text33.SetFocus
    End If
End Sub
Private Sub Text34_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text35.SetFocus
    End If
End Sub
Private Sub Text35_GotFocus()
    If Text34.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text34.SetFocus
    End If
End Sub

Private Sub Text35_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text35.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text35.SetFocus
        Else
        Text36.SetFocus
        End If
    End If
End Sub
Private Sub Text36_GotFocus()
    If Text35.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text35.SetFocus
    End If
End Sub
Private Sub Text36_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text37.SetFocus
    End If
End Sub

Private Sub Text37_GotFocus()
    If Text36.Text = Empty Then
        MsgBox "Mohon Inputkan Nama Himpunan"
        Text36.SetFocus
    End If
End Sub
Private Sub Text37_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text38.SetFocus
    End If
End Sub
Private Sub Text38_GotFocus()
    If Text37.Text = Empty Then
        MsgBox "Mohon Inputkan Domain"
        Text37.SetFocus
    End If
End Sub

Private Sub Text38_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text39.SetFocus
    End If
End Sub

Private Sub Text39_GotFocus()
    If Text38.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text38.SetFocus
    End If
End Sub

Private Sub Text39_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text40.SetFocus
    End If
End Sub

Private Sub Text40_GotFocus()
    If Text39.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text39.SetFocus
    End If
End Sub

Private Sub Text40_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text41.SetFocus
    End If
End Sub

Private Sub Text41_GotFocus()
    If Text40.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text40.SetFocus
    End If
End Sub

Private Sub Text41_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text41.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text41.SetFocus
        End If
    End If
End Sub

Private Sub Text42_GotFocus()
    If Text1.Text = Empty Then
        MsgBox "Mohon Inputkan Nama Kasus"
        Text1.SetFocus
    End If
End Sub
Private Sub Text42_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text43.SetFocus
    End If
End Sub

Private Sub Text43_GotFocus()
    If Text42.Text = Empty Then
        MsgBox "Mohon Inputkan Nama Variabel"
        Text42.SetFocus
    End If
End Sub

Private Sub Text43_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text44.SetFocus
    End If
End Sub

Private Sub Text44_GotFocus()
    If Text43.Text = Empty Then
        MsgBox "Mohon Inputkan Semesta Pembicaraan"
        Text43.SetFocus
    End If
End Sub

Private Sub Text44_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text45.SetFocus
    End If
End Sub

Private Sub Text45_GotFocus()
    If Text44.Text = Empty Then
        MsgBox "Mohon Inputkan Nama Himpunan"
        Text44.SetFocus
    End If
End Sub

Private Sub Text45_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text46.SetFocus
    End If
End Sub
Private Sub Text46_GotFocus()
    If Text45.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text45.SetFocus
    End If
End Sub
Private Sub Text46_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text47.SetFocus
    End If
End Sub
Private Sub Text47_GotFocus()
    If Text46.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text46.SetFocus
    End If
End Sub

Private Sub Text47_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text48.SetFocus
    End If
End Sub
Private Sub Text48_GotFocus()
    If Text47.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text47.SetFocus
    End If
End Sub
Private Sub Text48_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text49.SetFocus
    End If
End Sub

Private Sub Text49_GotFocus()
    If Text48.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text48.SetFocus
    End If
End Sub
Private Sub Text49_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text49.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text49.SetFocus
        Else
        Text50.SetFocus
        End If
    End If
End Sub

Private Sub Text50_GotFocus()
    If Text49.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text49.SetFocus
    End If
End Sub

Private Sub Text50_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text51.SetFocus
    End If
End Sub

Private Sub Text51_GotFocus()
    If Text50.Text = Empty Then
        MsgBox "Mohon Inputkan Nama Himpunan"
        Text50.SetFocus
    End If
End Sub

Private Sub Text51_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text52.SetFocus
    End If
End Sub

Private Sub Text52_GotFocus()
    If Text51.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text51.SetFocus
    End If
End Sub

Private Sub Text52_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text53.SetFocus
    End If
End Sub

Private Sub Text53_GotFocus()
    If Text52.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text52.SetFocus
    End If
End Sub

Private Sub Text53_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text54.SetFocus
    End If
End Sub
Private Sub Text54_GotFocus()
    If Text53.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text53.SetFocus
    End If
End Sub
Private Sub Text54_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text55.SetFocus
    End If
End Sub
Private Sub Text55_GotFocus()
    If Text54.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text54.SetFocus
    End If
End Sub

Private Sub Text55_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text55.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text55.SetFocus
        Else
        Text56.SetFocus
        End If
    End If
End Sub
Private Sub Text56_GotFocus()
    If Text55.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text55.SetFocus
    End If
End Sub
Private Sub Text56_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text57.SetFocus
    End If
End Sub

Private Sub Text57_GotFocus()
    If Text56.Text = Empty Then
        MsgBox "Mohon Inputkan Nama Himpunan"
        Text56.SetFocus
    End If
End Sub
Private Sub Text57_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text58.SetFocus
    End If
End Sub
Private Sub Text58_GotFocus()
    If Text57.Text = Empty Then
        MsgBox "Mohon Inputkan Domain"
        Text57.SetFocus
    End If
End Sub

Private Sub Text58_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text59.SetFocus
    End If
End Sub

Private Sub Text59_GotFocus()
    If Text58.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text58.SetFocus
    End If
End Sub

Private Sub Text59_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text60.SetFocus
    End If
End Sub

Private Sub Text60_GotFocus()
    If Text59.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text59.SetFocus
    End If
End Sub

Private Sub Text60_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text61.SetFocus
    End If
End Sub

Private Sub Text61_GotFocus()
    If Text60.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text60.SetFocus
    End If
End Sub

Private Sub Text61_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text61.Text = Empty Then
        MsgBox "Mohon Lengkapi Domain"
        Text61.SetFocus
        End If
    End If
End Sub

