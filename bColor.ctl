VERSION 5.00
Begin VB.UserControl uColor 
   BackColor       =   &H0080FF80&
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1695
   ScaleWidth      =   2955
   Begin VB.HScrollBar scrColor 
      Height          =   255
      Index           =   2
      LargeChange     =   50
      Left            =   720
      Max             =   255
      TabIndex        =   9
      Top             =   1320
      Width           =   2205
   End
   Begin VB.HScrollBar scrColor 
      Height          =   255
      Index           =   1
      LargeChange     =   50
      Left            =   720
      Max             =   255
      TabIndex        =   8
      Top             =   960
      Width           =   2205
   End
   Begin VB.HScrollBar scrColor 
      Height          =   255
      Index           =   0
      LargeChange     =   50
      Left            =   720
      Max             =   255
      TabIndex        =   7
      Top             =   600
      Width           =   2205
   End
   Begin VB.TextBox txtColor 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   2
      Left            =   225
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "255"
      Top             =   1320
      Width           =   510
   End
   Begin VB.TextBox txtColor 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   225
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "255"
      Top             =   960
      Width           =   510
   End
   Begin VB.TextBox txtColor 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   225
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "255"
      Top             =   600
      Width           =   510
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   15
      ScaleHeight     =   480
      ScaleWidth      =   2925
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   2925
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   195
      Index           =   2
      Left            =   75
      TabIndex        =   3
      Top             =   1335
      Width           =   105
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   975
      Width           =   105
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   615
      Width           =   105
   End
End
Attribute VB_Name = "uColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
