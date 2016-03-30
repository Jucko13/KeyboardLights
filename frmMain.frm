VERSION 5.00
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "NTSVC.ocx"
Begin VB.Form frmMain 
   Caption         =   "K70RGB Led setup by Jucko13"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   ScaleHeight     =   602
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   981
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frMonitor 
      Caption         =   "Monitor Brightness"
      Height          =   1095
      Left            =   135
      TabIndex        =   42
      Top             =   7440
      Width           =   2130
      Begin VB.CommandButton cmdTurnScreenOff 
         Caption         =   "Turn Screen Off"
         Height          =   270
         Left            =   120
         TabIndex        =   45
         Top             =   675
         Width           =   1890
      End
      Begin VB.CommandButton cmdMonitorBrightnessSet 
         Caption         =   "Set"
         Height          =   285
         Left            =   1260
         TabIndex        =   44
         Top             =   270
         Width           =   750
      End
      Begin VB.TextBox txtMonitorBrightness 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   135
         TabIndex        =   43
         Text            =   "100"
         Top             =   270
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdQuickAdd 
      Caption         =   "Quick Add"
      Height          =   270
      Left            =   150
      TabIndex        =   40
      Top             =   945
      Width           =   1200
   End
   Begin VB.CheckBox chk16Million 
      Caption         =   "16.8m Colors"
      Height          =   210
      Left            =   105
      TabIndex        =   39
      Top             =   660
      Value           =   1  'Checked
      Width           =   2745
   End
   Begin VB.ListBox lstKeydownEffect 
      Height          =   1035
      ItemData        =   "frmMain.frx":0000
      Left            =   12075
      List            =   "frmMain.frx":0019
      TabIndex        =   35
      Top             =   7800
      Width           =   2265
   End
   Begin VB.Timer tmrTimers 
      Interval        =   100
      Left            =   6090
      Top             =   7530
   End
   Begin VB.CheckBox chkCombine 
      Caption         =   "Combine all Data in one package"
      Height          =   210
      Left            =   105
      TabIndex        =   34
      Top             =   450
      Value           =   1  'Checked
      Width           =   2745
   End
   Begin VB.TextBox txtLoopSpeed 
      Height          =   285
      Left            =   5790
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "0,0000"
      Top             =   945
      Width           =   1185
   End
   Begin VB.HScrollBar scrSpeed 
      Height          =   195
      LargeChange     =   150
      Left            =   4815
      Max             =   1999
      Min             =   1500
      TabIndex        =   14
      Top             =   450
      Value           =   1995
      Width           =   2145
   End
   Begin VB.CommandButton cmdSolidAdd 
      Caption         =   "Add"
      Height          =   480
      Left            =   10095
      TabIndex        =   23
      Top             =   8370
      Width           =   1065
   End
   Begin VB.CommandButton cmdSolidEdit 
      Caption         =   "Edit"
      Height          =   480
      Left            =   10095
      TabIndex        =   24
      Top             =   7905
      Width           =   1065
   End
   Begin VB.ListBox lstSolidKeys 
      Height          =   1425
      Left            =   7605
      TabIndex        =   22
      Top             =   7440
      Width           =   2505
   End
   Begin VB.ListBox lstCommands 
      Height          =   1425
      ItemData        =   "frmMain.frx":0090
      Left            =   2520
      List            =   "frmMain.frx":0097
      TabIndex        =   21
      Top             =   7455
      Width           =   1740
   End
   Begin VB.ListBox lstAnimationList 
      Height          =   1425
      Left            =   4275
      TabIndex        =   20
      Top             =   7455
      Width           =   1740
   End
   Begin VB.HScrollBar scrWidth 
      Height          =   210
      LargeChange     =   6
      Left            =   4815
      Max             =   25
      Min             =   2
      TabIndex        =   17
      Top             =   255
      Value           =   4
      Width           =   2145
   End
   Begin NTService.NTService NTService 
      Left            =   7080
      Top             =   7995
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      DisplayName     =   "K70RGB_Lights"
      ServiceName     =   "KeyboardLights"
      StartMode       =   3
   End
   Begin VB.PictureBox picStatusBarMemory 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   150
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   959
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   14385
   End
   Begin VB.PictureBox picStatusBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   165
      Picture         =   "frmMain.frx":00A7
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   959
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6435
      Width           =   14385
   End
   Begin VB.TextBox txtSendSpeed 
      Height          =   285
      Left            =   5790
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0,0000"
      Top             =   660
      Width           =   1185
   End
   Begin VB.ListBox lstAnimations 
      Appearance      =   0  'Flat
      Height          =   1005
      ItemData        =   "frmMain.frx":4B55
      Left            =   7575
      List            =   "frmMain.frx":4B92
      TabIndex        =   12
      Top             =   60
      Width           =   4140
   End
   Begin VB.Timer tmrVolume 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6525
      Top             =   7485
   End
   Begin VB.HScrollBar scrSolid 
      Height          =   165
      Index           =   2
      Left            =   12390
      Max             =   255
      TabIndex        =   11
      Top             =   1020
      Value           =   7
      Width           =   1965
   End
   Begin VB.HScrollBar scrSolid 
      Height          =   165
      Index           =   1
      Left            =   12390
      Max             =   255
      TabIndex        =   10
      Top             =   810
      Value           =   7
      Width           =   1965
   End
   Begin VB.HScrollBar scrSolid 
      Height          =   165
      Index           =   0
      Left            =   12390
      Max             =   255
      TabIndex        =   6
      Top             =   600
      Value           =   7
      Width           =   1965
   End
   Begin VB.PictureBox picSolid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   11775
      ScaleHeight     =   435
      ScaleWidth      =   2850
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   60
      Width           =   2880
   End
   Begin VB.CheckBox chkAnimation 
      Caption         =   "Maxout Colors"
      Height          =   285
      Index           =   0
      Left            =   7575
      TabIndex        =   3
      Top             =   1065
      Width           =   4140
   End
   Begin VB.PictureBox picKeyboard 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   5040
      Left            =   165
      Picture         =   "frmMain.frx":4CFC
      ScaleHeight     =   336
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   959
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1395
      Width           =   14385
      Begin VB.Shape shpKey 
         BorderColor     =   &H000000FF&
         Height          =   870
         Left            =   5475
         Top             =   240
         Visible         =   0   'False
         Width           =   930
      End
   End
   Begin VB.PictureBox picResize 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1065
      Left            =   195
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1275
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Timer tmrAnimation 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7020
      Top             =   7515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Info about the Keyboard"
      Height          =   255
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   2115
   End
   Begin VB.PictureBox picMemory 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   540
      Left            =   3900
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   314
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1290
      Visible         =   0   'False
      Width           =   4710
   End
   Begin VB.CommandButton cmdSolidSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   480
      Left            =   10095
      TabIndex        =   25
      Top             =   7440
      Width           =   1065
   End
   Begin VB.HScrollBar scrTransparency 
      CausesValidation=   0   'False
      Height          =   210
      LargeChange     =   50
      Left            =   4815
      Max             =   255
      TabIndex        =   37
      Top             =   60
      Value           =   127
      Width           =   2145
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "KeyPaint Transparency"
      Height          =   195
      Index           =   8
      Left            =   3105
      TabIndex        =   41
      Top             =   30
      Width           =   1740
   End
   Begin VB.Label lblKeyboardTransparency 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "127"
      Height          =   195
      Left            =   7005
      TabIndex        =   38
      Top             =   60
      Width           =   525
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ReactiveTyping"
      Height          =   195
      Index           =   7
      Left            =   12120
      TabIndex        =   36
      Top             =   7515
      Width           =   1125
   End
   Begin VB.Label lblColor 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      Height          =   195
      Index           =   2
      Left            =   14385
      TabIndex        =   33
      Top             =   1005
      Width           =   300
   End
   Begin VB.Label lblColor 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      Height          =   195
      Index           =   1
      Left            =   14385
      TabIndex        =   32
      Top             =   795
      Width           =   300
   End
   Begin VB.Label lblColor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "255"
      Height          =   195
      Index           =   0
      Left            =   14385
      TabIndex        =   31
      Top             =   585
      Width           =   300
   End
   Begin VB.Label lblAnimationInterval 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   195
      Left            =   7215
      TabIndex        =   30
      Top             =   450
      Width           =   105
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Animation Interval (ms):"
      Height          =   195
      Index           =   6
      Left            =   3165
      TabIndex        =   29
      Top             =   435
      Width           =   1635
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Program Loop Speed:"
      Height          =   195
      Index           =   5
      Left            =   4185
      TabIndex        =   27
      Top             =   990
      Width           =   1560
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Send Speed:"
      Height          =   195
      Index           =   4
      Left            =   4410
      TabIndex        =   26
      Top             =   705
      Width           =   1380
   End
   Begin VB.Label lblAnimationWidth 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   195
      Left            =   7005
      TabIndex        =   19
      Top             =   255
      Width           =   525
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Animation Width:"
      Height          =   195
      Index           =   3
      Left            =   3600
      TabIndex        =   18
      Top             =   240
      Width           =   1200
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
      Height          =   195
      Index           =   2
      Left            =   11820
      TabIndex        =   9
      Top             =   1005
      Width           =   300
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Green"
      Height          =   195
      Index           =   1
      Left            =   11820
      TabIndex        =   8
      Top             =   795
      Width           =   435
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Red"
      Height          =   195
      Index           =   0
      Left            =   11820
      TabIndex        =   7
      Top             =   585
      Width           =   300
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Tube
    X As Long
    y As Long
    Visible As Boolean
End Type

Private Type Stars
    X As Single
    y As Single
    nStep As Long
    Angle As Single
    R As Byte
    G As Byte
    B As Byte
End Type

Private Type POINT
    X As Long
    y As Long
End Type

Dim my_descriptor As UsbDeviceDescriptor
Dim dev_config As UsbConfigDescriptor
Dim my_interface As UsbInterfaceDescriptor
Dim my_endpoint As UsbEndPointDescriptor

Dim keyboard As Long

Dim Data(0 To 1152) As Byte '832+5*64=1152

Private Const Red As Long = 0
Private Const Green As Long = 1
Private Const Blue As Long = 2

Private Type KeyPlace
    X As Byte
    y As Byte
End Type

Private Type Key_RGB
    R As Byte
    G As Byte
    B As Byte
End Type

Dim KeyColor(0 To 144, Red To Blue) As Byte
Dim LastKeyColor(0 To 144, Red To Blue) As Byte
Dim KeyMatrix(0 To 23, 0 To 6) As Byte
Dim KeyMatrixReversed(0 To 144) As KeyPlace

Dim KeyNames() As String
Dim KeySelected() As Boolean
Dim KeySelectedSemi() As Boolean
Dim KeyMouseOver As Byte
Dim KeyMouseSelectionCount As Byte
Dim KeyVKCodes() As Byte
Dim KeyVKDown() As Boolean
Dim KeyVKDownPrevious() As Boolean
Dim KeyVKDownColors() As Key_RGB


Private Enum Timers
    t_ProgramLoop = 0
    t_SendData = 1
End Enum

Dim TimeQuerys(0 To 3) As Single

Private Type SolidKeys
    s_Keylist() As Byte
    lColor As Key_RGB
End Type

Dim KeySolidList() As SolidKeys


Dim KeySelectorX As Long
Dim KeySelectorY As Long
Dim KeySelectorDragging As Boolean

Private Type KeyPosition
    X As Long
    y As Long
    Width As Long
    Height As Long
End Type

Dim KeyPlaces(0 To 144) As KeyPosition

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINT, ByVal nCount As Long) As Long


Private Const AC_SRC_OVER = &H0
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type


Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, _
    ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, _
    ByVal heightSrc As Long, ByVal BLENDFUNCT As Long) As Boolean
         
         
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)


Dim m_LonVolume(0 To 10) As Long
Dim m_LonVolumePrevious(0 To 10) As Long
Dim m_bVolumeEffect(0 To 10) As Boolean
Dim m_bVolumeInitialized(0 To 10) As Boolean
Dim m_bVolumeMute(0 To 10) As Boolean



''''''''''''''''''''''''''''''''''''''''''''''''
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
''''''''''''''''''''''''''''''''''''''''''''''''




Private Enum Keys
    KEY_ESCAPE = 0
    KEY_TILDE = 1
    KEY_TAB = 2
    KEY_CAPSLOCK = 3
    KEY_SHIFT_LEFT = 4
    KEY_CTRL_LEFT = 5
    KEY_F12 = 6
    KEY_EQUALSIGN = 7
    KEY_WINDOWS_LOCK = 8
    KEY_NUM_7 = 9
    KEY_F1 = 12
    KEY_1 = 13
    KEY_Q = 14
    KEY_A = 15
    KEY_WINDOWS_LEFT = 17
    KEY_PRINTSCREEN = 18
    KEY_MEDIA_MUTE = 20
    KEY_NUM_8 = 21
    KEY_F2 = 24
    KEY_2 = 25
    KEY_W = 26
    KEY_S = 27
    KEY_Z = 28
    KEY_ALT_LEFT = 29
    KEY_SCROLLLOCK = 30
    KEY_BACKSPACE = 31
    KEY_MEDIA_STOP = 32
    KEY_NUM_9 = 33
    KEY_F3 = 36
    KEY_3 = 37
    KEY_E = 38
    KEY_D = 39
    KEY_X = 40
    KEY_PAUSE = 42
    KEY_DELETE = 43
    KEY_MEDIA_PREVIOUS = 44
    KEY_F4 = 48
    KEY_4 = 49
    KEY_R = 50
    KEY_F = 51
    KEY_C = 52
    KEY_SPACE = 53
    KEY_INSERT = 54
    KEY_END = 55
    KEY_MEDIA_PLAY = 56
    KEY_NUM_4 = 57
    KEY_F5 = 60
    KEY_5 = 61
    KEY_T = 62
    KEY_G = 63
    KEY_V = 64
    KEY_HOME = 66
    KEY_PAGEDOWN = 67
    KEY_MEDIA_NEXT = 68
    KEY_NUM_5 = 69
    KEY_F6 = 72
    KEY_6 = 73
    KEY_Y = 74
    KEY_H = 75
    KEY_B = 76
    KEY_PAGEUP = 78
    KEY_SHIFT_RIGHT = 79
    KEY_NUM_NUMLOCK = 80
    KEY_NUM_6 = 81
    KEY_F7 = 84
    KEY_7 = 85
    KEY_U = 86
    KEY_J = 87
    KEY_N = 88
    KEY_ALT_RIGHT = 89
    KEY_BLOK_HAAK_SLUIT = 90
    KEY_CTRL_RIGHT = 91
    KEY_NUM_SLASH = 92
    KEY_NUM_1 = 93
    KEY_F8 = 96
    KEY_8 = 97
    KEY_I = 98
    KEY_K = 99
    KEY_M = 100
    KEY_WINDOWS_RIGHT = 101
    KEY_SLASH_BACK = 102
    KEY_ARROW_UP = 103
    KEY_NUM_ASTERIX = 104
    KEY_NUM_2 = 105
    KEY_F9 = 108
    KEY_9 = 109
    KEY_O = 110
    KEY_L = 111
    KEY_COMMA = 112
    KEY_MENU = 113
    KEY_ARROW_LEFT = 115
    KEY_NUM_MIN = 116
    KEY_NUM_3 = 117
    KEY_F10 = 120
    KEY_0 = 121
    KEY_P = 122
    KEY_SEMICOLON = 123
    KEY_PERIOD = 124
    KEY_ENTER = 126
    KEY_ARROW_DOWN = 127
    KEY_NUM_PLUS = 128
    KEY_NUM_0 = 129
    KEY_F11 = 132
    KEY_MIN = 133
    KEY_BLOK_HAAK_OPEN = 134
    KEY_QUOTE = 135
    KEY_SLASH_FORWARD = 136
    KEY_BRIGHTNESS = 137
    KEY_ARROW_RIGHT = 139
    KEY_NUM_ENTER = 140
    KEY_NUM_DELETE = 141
End Enum
Dim key As Keys


Dim KeyNUMberPerRow(0 To 6) As Byte



Private Const PACKET1 As Long = 0
Private Const PACKET2 As Long = 64
Private Const PACKET3 As Long = 128
Private Const PACKET4 As Long = 192
Private Const PACKET5 As Long = 256
Private Const PACKET6 As Long = 320
Private Const PACKET7 As Long = 384
Private Const PACKET8 As Long = 448
Private Const PACKET9 As Long = 512
Private Const PACKET10 As Long = 576
Private Const PACKET11 As Long = 640
Private Const PACKET12 As Long = 704

Private Const THREEBITPACKET As Long = PACKET12 + PACKET2

Dim interface As Long


Dim bLoop As Boolean
Dim bLooping As Boolean
Private Declare Function GetTickCount Lib "Kernel32" () As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
         ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, _
         ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
         ByVal ySrc As Long, ByVal dwRop As Long) As Long
         
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
    

Private Sub chk16Million_Click()
    SaveSetting "KeyboardLights", "Animation", "16MillionColors", chk16Million.value

End Sub

Private Sub cmdMonitorBrightnessSet_Click()
    Dim i As Long
    
    Dim newBrightness As Long
    Dim percentage As Long
    
    
    percentage = Val(txtMonitorBrightness.Text)
    If percentage > 100 Then percentage = 100
    If percentage < 0 Then percentage = 0
    
    
    
    For i = 0 To MonitorCount - 1
        newBrightness = percentage '(Monitors(i).lBrightnessMaximum - Monitors(i).lBrightnessMinimum) / 100 * percentage + Monitors(i).lBrightnessMinimum

        Monitors(i).lBrightnessCurrent = newBrightness
        
        SetMonitorBrightness Monitors(i).pPhysicalInfo(0).hPhysicalMonitor, newBrightness
    Next i
End Sub

Private Sub cmdQuickAdd_Click()
        
        
    scrSolid(0).value = Fix(Rnd * 255)
    scrSolid(1).value = Fix(Rnd * 255)
    scrSolid(2).value = Fix(Rnd * 255)
    
    cmdSolidAdd_Click
    
    scrSolid(0).value = 0
    scrSolid(1).value = 0
    scrSolid(2).value = 0
End Sub

Private Sub cmdSolidAdd_Click()

    Dim i As Long
    Dim lcount As Long
    
    lcount = KeyMouseSelectionCount
    
    If lcount = 0 Then Exit Sub
    
    lstSolidKeys.AddItem lcount & " " & Hex(picSolid.BackColor)
    
    Dim ListSize As Long
    
    ListSize = lstSolidKeys.ListCount - 1
    
    ReDim Preserve KeySolidList(0 To ListSize)
    
    ReDim KeySolidList(ListSize).s_Keylist(0 To lcount - 1)
    
    Dim j As Long
    
    For i = 0 To 144
        If KeySelectedSemi(i) Then
            KeySolidList(ListSize).s_Keylist(j) = i
            j = j + 1
        End If
    Next i
    
    KeySolidList(ListSize).lColor.R = scrSolid(0).value
    KeySolidList(ListSize).lColor.G = scrSolid(1).value
    KeySolidList(ListSize).lColor.B = scrSolid(2).value
    
End Sub

Private Sub cmdSolidEdit_Click()
    Dim i As Long
    Dim ListPlace As Long
    
    If cmdSolidEdit.Caption = "Cancel" Then
        cmdSolidEdit.Caption = "Edit"
        cmdSolidAdd.Enabled = True
        cmdSolidSave.Enabled = False
        ReDim KeySelected(0 To 144)
        ReDim KeySelectedSemi(0 To 144)
        lstSolidKeys.Enabled = True
        KeyMouseSelectionCount = 0
        Exit Sub
    End If
    
    ListPlace = lstSolidKeys.ListIndex
    
    If ListPlace = -1 Then Exit Sub
    ReDim KeySelected(0 To 144)
    ReDim KeySelectedSemi(0 To 144)
    lstSolidKeys.Enabled = False
    
    KeyMouseSelectionCount = 0
    
    With KeySolidList(ListPlace)
        For i = 0 To UBound(.s_Keylist)
            KeySelectedSemi(.s_Keylist(i)) = True
            KeyMouseSelectionCount = KeyMouseSelectionCount + 1
        Next i
        
        scrSolid(0).value = .lColor.R
        scrSolid(1).value = .lColor.G
        scrSolid(2).value = .lColor.B
        
    End With
    
    cmdSolidAdd.Enabled = False
    cmdSolidSave.Enabled = True
    cmdSolidEdit.Caption = "Cancel"
    
    
    
End Sub

Private Sub cmdSolidSave_Click()
    Dim i As Long
    Dim lcount As Long
    Dim ListPlace As Long
    
    For i = 0 To 144
        If KeySelectedSemi(i) Then
            lcount = lcount + 1
        End If
    Next i
    
    If lcount = 0 Then Exit Sub
    ListPlace = lstSolidKeys.ListIndex
    
    lstSolidKeys.List(ListPlace) = lcount & " " & Hex(picSolid.BackColor)
    
    
    'ReDim Preserve KeySolidList(0 To ListSize)
    
    ReDim KeySolidList(ListPlace).s_Keylist(0 To lcount - 1)
    
    Dim j As Long
    
    For i = 0 To 144
        If KeySelectedSemi(i) Then
            KeySolidList(ListPlace).s_Keylist(j) = i
            j = j + 1
        End If
    Next i
    
    KeySolidList(ListPlace).lColor.R = scrSolid(0).value
    KeySolidList(ListPlace).lColor.G = scrSolid(1).value
    KeySolidList(ListPlace).lColor.B = scrSolid(2).value
    
    
    cmdSolidEdit.Caption = "Edit"
    cmdSolidAdd.Enabled = True
    cmdSolidSave.Enabled = False
    ReDim KeySelected(0 To 144)
    ReDim KeySelectedSemi(0 To 144)
    lstSolidKeys.Enabled = True
End Sub

Private Sub cmdTurnScreenOff_Click()
    Dim i As Long
    
    For i = 0 To MonitorCount - 1
        SetVCPFeature Monitors(i).pPhysicalInfo(0).hPhysicalMonitor, CByte(&HD6), 5
    Next i
End Sub

'
'Sub KeyboardToPicture(lPic As PictureBox, lResizedPic As PictureBox)
'    Dim PicBits() As Byte, PicInfo As BITMAP
'    Dim Cnt As Long, BytesPerLine As Long
'
'
'    lPic.Picture = lPic.Image
'
'    GetObject lPic.Image, Len(PicInfo), PicInfo
'
'    BytesPerLine = (PicInfo.bmWidth * 4)
'    ReDim PicBits(1 To BytesPerLine * PicInfo.bmHeight) As Byte
'    'GetBitmapBits lPic.Picture, UBound(PicBits), PicBits(1)
'
'    Dim i As Long
'
'    'PicBits (1) 'B
'    'PicBits (2) 'G
'    'PicBits (3) 'R
'    Dim X As Long
'    Dim Y As Long
'
'    For i = 1 To UBound(PicBits) Step 4
'        Y = (i \ BytesPerLine)
'        X = (i - Y * BytesPerLine - 1) / 4
'        GetLed X, Y, PicBits(i + 2), PicBits(i + 1), PicBits(i)
'    Next i
'
'    SetBitmapBits lPic.Image, UBound(PicBits), PicBits(1)
'    lPic.Picture = lPic.Image
'
'    lResizedPic.PaintPicture lPic.Picture, 0, 0, lResizedPic.ScaleWidth, lResizedPic.ScaleHeight, 0, 0, lPic.ScaleWidth, lPic.ScaleHeight
'
'    lResizedPic.Picture = lResizedPic.Image
'
'    Paint_Keyboard
'End Sub

Private Sub Command1_Click()
    If keyboard = -1 Then
        MsgBox "Keyboard Not Connected!", vbCritical
        Exit Sub
    End If
    
    GetBusInfo
    
    frmInfo.Show
End Sub

Sub Random_Lines_Vertical(Optional MaxColors As Boolean = True)
    Dim i As Long
    Dim j As Long

    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    
    For i = 0 To 23
        GetRandomColor R, G, B, MaxColors
        
        For j = 0 To 6
            SetLed i, j, R, G, B
        Next j
    Next i
End Sub


Sub Random_Lines_Horizontal(Optional MaxColors As Boolean = True)
    Dim i As Long
    Dim j As Long
    
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    
    For i = 0 To 6
        GetRandomColor R, G, B, MaxColors
        
        For j = 0 To 23
            SetLed j, i, R, G, B
        Next j
    Next i
End Sub


Sub Random_Key_Color(Optional MaxColors As Boolean = True)
    
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    
    GetRandomColor R, G, B, MaxColors
    
    SetLed Fix(Rnd * 24 + 1), Fix(Rnd * 7 + 1), R, G, B
End Sub

Sub Random_Keyboard_Color(Optional MaxColors As Boolean = True)
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    
    Dim X As Long
    Dim y As Long
    
    Randomize
    For X = 0 To 23
        For y = 0 To 6
            
            GetRandomColor R, G, B, MaxColors
            
            SetLed X, y, R, G, B
        Next y
    Next X
    
    
End Sub

Sub Random_Circle_Move(Optional MaxColors As Boolean = True)
    Static X As Long
    Static R As Byte
    Static G As Byte
    Static B As Byte
    
    'ClearLeds
    X = X + 1
    
    If X > 30 Then
        X = -5
        GetRandomColor R, G, B, MaxColors
    End If
    
    Make_Circle X, 3, 3, R * 1, G * 1, B * 1
    Make_Circle X, 3, 2, 255 - R, 255 - G, 255 - B
End Sub

Sub Make_Circle(X As Long, y As Long, sRadius As Single, R As Byte, G As Byte, B As Byte)
    Dim tmpSteps As Double
    
    Dim tX As Long
    Dim tY As Long
    If sRadius = 0 Then sRadius = 0.1
    
    tmpSteps = (2 * 3.14159) / (sRadius * 34)
    Dim i As Double
    
    'picKeyboard.Cls
    
    For i = 0 To 2 * 3.14159 Step tmpSteps
        tX = Sin(i) * sRadius + X
        tY = Cos(i) * sRadius + y

        SetLed tX, tY, R, G, B
    Next i
End Sub



Sub Random_Spiral_Fill(Optional MaxColors As Boolean = True)
    Static X As Long
    Static y As Long
    
    Static X_Dir As Long
    Static Y_Dir As Long
    
    Static R As Byte
    Static G As Byte
    Static B As Byte
    
    Static CurrentStep As Long
    
    Const StepTable As String = "0,0|23,0|23,6|0,6|" & _
                                "0,1|22,1|22,5|1,5|" & _
                                "1,2|21,2|21,4|2,4|" & _
                                "2,3|20,3|20,3|3,3|0,0"
          
    Dim Steps() As String
    Steps = Split(StepTable, "|")
    
    If CurrentStep = 0 Then
        GetRandomColor R, G, B, MaxColors
        
        X = 0
        y = 0
        CurrentStep = CurrentStep + 1
    Else
        If Steps(CurrentStep + 1) = "0,0" Then
            CurrentStep = 0
        Else
            Dim tmpSplit() As String
            
            tmpSplit = Split(Steps(CurrentStep + 1), ",")
            If tmpSplit(0) < X Then
                X = X - 1
            ElseIf tmpSplit(0) > X Then
                X = X + 1
            Else 'already on the X
                If tmpSplit(1) < y Then
                    y = y - 1
                ElseIf tmpSplit(1) > y Then
                    y = y + 1
                Else 'already on the Y
                    CurrentStep = CurrentStep + 1
                End If
            End If
        End If
        
        
    End If
    
    SetLed X, y, R, G, B
End Sub



Sub Static_Spiral_Fill_Rainbow(Optional MaxColors As Boolean = True)
    Dim X As Long
    Dim y As Long
    
    Dim X_Dir As Long
    Dim Y_Dir As Long
    
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    
    Dim CurrentStep As Long
    Static Initialized As Boolean
    Dim i As Long
    
    Const StepTable As String = "0,0|23,0|23,6|0,6|" & _
                                "0,1|22,1|22,5|1,5|" & _
                                "1,2|21,2|21,4|2,4|" & _
                                "2,3|20,3|20,3|0,0"
    Static numLoops As Long
    
    Static Steps() As String
    If Initialized = False Then
        Steps = Split(StepTable, "|")
        Initialized = True
    End If
    
    numLoops = numLoops + 1
    If numLoops = 181 Then numLoops = 0
    
    
    For i = 0 To 181
        
        
        If CurrentStep = 0 Then
            'GetRandomColor R, G, B, MaxColors
            X = 0
            y = 0
            
            CurrentStep = CurrentStep + 1
        Else
            'If CurrentStep = 14 Then
            '    Debug.Print "lol"
            'End If
            
            If Steps(CurrentStep + 1) = "0,0" Then
                CurrentStep = 0
            Else
                Dim tmpSplit() As String
                
                tmpSplit = Split(Steps(CurrentStep + 1), ",")
                If tmpSplit(0) < X Then
                    X = X - 1
                ElseIf tmpSplit(0) > X Then
                    X = X + 1
                Else 'already on the X
                    If tmpSplit(1) < y Then
                        y = y - 1
                    ElseIf tmpSplit(1) > y Then
                        y = y + 1
                    Else 'already on the Y
                        CurrentStep = CurrentStep + 1
                    End If
                End If
            End If
            
            'totalSteps = totalSteps + 1
        End If
        
        getRainbowColor R, G, B, i + numLoops, 180
        
        SetLed X, y, R, G, B
        
        'If i = 180 Then
        '    Debug.Print "lol"
        'End If
        
    Next i
End Sub

Sub getRainbowColor(ByRef R As Byte, ByRef G As Byte, ByRef B As Byte, ByVal lStep As Long, lNumSteps As Long)
    Dim tR As Byte
    Dim tG As Byte
    Dim tB As Byte
    
    Dim devider As Long
    
    devider = lNumSteps / 6
    
    Dim multiplier As Long
    
    lStep = (lStep Mod lNumSteps)
    
    multiplier = devider
    multiplier = lStep Mod multiplier
    multiplier = Abs(multiplier)
    
    Select Case lStep
        Case Is < devider
            tR = 255
            tG = 255 / devider * multiplier
            tB = 0
        Case Is < devider * 2
            tR = 255 - (255 / devider * multiplier)
            tG = 255
            tB = 0
        Case Is < devider * 3
            tR = 0
            tG = 255
            tB = 255 / devider * multiplier
        
        Case Is < devider * 4
            tR = 0
            tG = 255 - (255 / devider * multiplier)
            tB = 255
        
        Case Is < devider * 5
            tR = 255 / devider * multiplier
            tG = 0
            tB = 255
        
        Case Else
            tR = 255
            tG = 0
            tB = 255 - (255 / devider * multiplier)
        
    End Select
    
    
    R = tR
    G = tG
    B = tB
End Sub



Sub Random_Arrow_Move(Optional MaxColors As Boolean = True)
    Static X As Long
    Static R As Byte
    Static G As Byte
    Static B As Byte
    
    ClearLeds
    
    X = X + 1
    If X > 30 Then
        X = -5
        GetRandomColor R, G, B, MaxColors
        
        
        
        
        
    End If
    
    'R = 0
    'G = 0
    'B = 0
    Dim i As Long
    Dim w As Long
    
    For w = 0 To 1
        For i = 0 To 3
            SetLed X + i + w, i, R, G, B
            SetLed X + i + w, 7 - i, R, G, B
            
            SetLed 21 - (X + i + w), i, R, G, B
            SetLed 21 - (X + i + w), 7 - i, R, G, B
            
        Next i
    Next w
    
    'picAni.Picture = LoadPicture()
    'picAni.DrawWidth = 1
    'picAni.Line (X - 4, 0)-(X, 4), RGB(R, G, B)
    'picAni.Line (X - 4, 6)-(X, 2), RGB(R, G, B)
    
    'picAni.DrawWidth = 1
    'picAni.Line (picAni.ScaleWidth - X + 4 - 1, 0)-(picAni.ScaleWidth - X - 1, 4), RGB(R, G, B)
    'picAni.Line (picAni.ScaleWidth - X + 4 - 1, 6)-(picAni.ScaleWidth - X - 1, 2), RGB(R, G, B)
    
End Sub

Sub ClearLeds()
    Dim i As Long
    
    For i = 0 To 144
        KeyColor(i, Red) = 0
        KeyColor(i, Blue) = 0
        KeyColor(i, Green) = 0
    Next i
End Sub


Private Sub Static_Wave_Rainbow(ByVal lFase As Long, lWidth As Long, bWave As Boolean)
    Dim R As Long
    Dim G As Long
    Dim B As Long
    Dim cadd As Long
    Dim cadd2 As Long
    Dim cadd3 As Long
    Dim frmscw As Long
    Dim devide As Long
    Dim FrmSh As Long
    Dim X As Long
    Dim y As Long
    Dim tmpX As Long
    
    Dim clr1 As Long
    Dim clr2 As Long
    Dim clr3 As Long
    Dim clr4 As Long
    Dim clr5 As Long
    Dim clr6 As Long
    
    R = 255: G = 0: B = 0
    cadd = 3
    frmscw = lWidth
    devide = Int((frmscw \ 6))
    cadd = 255 / devide: cadd2 = 0
    
    '(X + lFase) Mod frmscw
    Dim tl As Long
    'ClearLeds
    
    Dim y_Array() As Long
    
    ReDim y_Array(0 To lWidth) As Long
    
    If bWave Then
        For X = 0 To lWidth
            y_Array(X) = Sin(2 * 3.14159 / lWidth * (X + lFase)) * 4
        Next X
        
        lFase = lFase + (lWidth / 6)
    End If
    
    FrmSh = 6
    For X = 0 To devide ' section '1 6th of form size
        cadd3 = Int(cadd2) ' cut off fraction for byte
        clr1 = RGB(255, cadd3, 0) 'red to yellow
        
        
        
        
        'If bWave Then
            tmpX = (X + lFase) Mod frmscw
            For y = 0 To FrmSh: SetLed tmpX, y + y_Array(tmpX), 255, cadd3 * 1, 0: Next y
            
            tmpX = (X + (devide) + lFase) Mod frmscw
            For y = 0 To (FrmSh): SetLed tmpX, y + y_Array(tmpX), 255 - cadd3, 255, 0: Next y
    
            tmpX = (X + (devide * 2) + lFase) Mod frmscw
            For y = 0 To FrmSh: SetLed tmpX, y + y_Array(tmpX), 0, 255, cadd3 * 1: Next y
    
            tmpX = (X + (devide * 3) + lFase) Mod frmscw
            For y = 0 To FrmSh:
            SetLed tmpX, y + y_Array(tmpX), 0, 255 - cadd3, 255:
            Next y
    
            tmpX = (X + (devide * 4) + lFase) Mod frmscw
            For y = 0 To FrmSh: SetLed tmpX, y + y_Array(tmpX), cadd3 * 1, 0, 255: Next y
            
            tmpX = (X + (devide * 5) + lFase) Mod frmscw
            For y = 0 To FrmSh: SetLed tmpX, y + y_Array(tmpX), 255, 0, 255 - cadd3: Next y
            
'        Else
'            tmpX = (X + (frm2 * 5) + lFase) Mod frmscw
'            For Y = 0 To FrmSh: SetLed tmpX, Y + y_Array(tmpX), 255, 255 - cadd3 * 1, 0: Next Y
'
'            tmpX = (X + (frm2 * 4) + lFase) Mod frmscw
'            For Y = 0 To FrmSh: SetLed tmpX, Y + y_Array(tmpX), cadd3 * 1, 255, 0: Next Y
'
'            tmpX = (X + (frm2 * 3) + lFase) Mod frmscw
'            For Y = 0 To FrmSh: SetLed tmpX, Y + y_Array(tmpX), 0, 255, 255 - cadd3 * 1: Next Y
'
'            tmpX = (X + (frm2 * 2) + lFase) Mod frmscw
'            For Y = 0 To FrmSh: SetLed tmpX, Y + y_Array(tmpX), 0, cadd3 * 1, 255: Next Y
'
'            tmpX = (X + (frm2 * 1) + lFase) Mod frmscw
'            For Y = 0 To FrmSh: SetLed tmpX, Y + y_Array(tmpX), 255 - cadd3 * 1, 0, 255: Next Y
'
'
'            tmpX = (X + (frm2 * 0) + lFase) Mod frmscw
'            For Y = 0 To FrmSh: SetLed tmpX, Y + y_Array(tmpX), 255, 0, cadd3 * 1: Next Y
'
'        End If
        
        cadd2 = cadd2 + cadd 'accumulate
        If cadd2 > 255 Then cadd2 = 255
    Next X ' each point in section
    
End Sub

Sub SetLed(X As Long, y As Long, R As Byte, G As Byte, B As Byte)
    Dim tmpKey As Byte
    If y < 0 Or X < 0 Then Exit Sub
    If X > 23 Or y > 6 Then Exit Sub
    
    tmpKey = KeyMatrix(X, y)
    If tmpKey = 144 Then
        'KeyColor(tmpKey, Red) = 7
        'KeyColor(tmpKey, Green) = 7
        'KeyColor(tmpKey, Blue) = 7
        Exit Sub
    End If
    
    KeyColor(tmpKey, Red) = R
    KeyColor(tmpKey, Green) = G
    KeyColor(tmpKey, Blue) = B
End Sub

Sub SetLedByName(Name As Byte, R As Byte, G As Byte, B As Byte)

    If Name = 144 Then
        'Exit Sub
    End If
    
    KeyColor(Name, Red) = R
    KeyColor(Name, Green) = G
    KeyColor(Name, Blue) = B
    
End Sub

Function GetLed(X As Long, y As Long, ByRef R As Byte, ByRef G As Byte, ByRef B As Byte) As Byte
    Dim tmpKey As Byte
    
    If y < 0 Or X < 0 Then Exit Function
    If X > 23 Or y > 6 Then Exit Function
    
    tmpKey = KeyMatrix(X, y)
    
    If tmpKey <> 144 Then
        R = KeyColor(tmpKey, Red)
        G = KeyColor(tmpKey, Green)
        B = KeyColor(tmpKey, Blue)
        GetLed = tmpKey
    Else
        R = 0
        G = 0
        B = 0
        GetLed = 144
    End If
End Function

Function GetLedR(X As Long, y As Long) As Byte
    Dim tmpKey As Byte
    GetLedR = 0
    
    If y < 0 Or X < 0 Then Exit Function
    If X > 23 Or y > 6 Then Exit Function
    
    tmpKey = KeyMatrix(X, y)
    
    If tmpKey <> 144 Then
        GetLedR = KeyColor(tmpKey, Red)
    Else
        GetLedR = 0
    End If
End Function


Private Sub Form_Unload(Cancel As Integer)
    'NTService.StopService
    bLoop = False
    
    CloseAllMonitorHandles
    
    tmrAnimation.Enabled = False
    tmrVolume.Enabled = False
    
    'Debug.Print UsbReset(keyboard)
    
    If keyboard <> -1 Then
        Debug.Print UsbReleaseInterface(keyboard, interface)
        Debug.Print UsbClose(keyboard)
    End If
    
    RemoveKeyboardHook
    
    keyboard = -1
    End
End Sub

Private Sub lstAnimations_Click()
    Select Case lstAnimations.ListIndex
        Case 0 'solid color
            scrSpeed.value = 2000 - 10
            
        Case 1, 2 'rainbow, rainbow wave
            scrSpeed.value = 2000 - 80
            
        Case 3 ' circle cycling
            scrSpeed.value = 2000 - 50
        
        Case 4 'random horizontal lines
            scrSpeed.value = 2000 - 100
            
        Case 5 'random vertical lines
            scrSpeed.value = 2000 - 100
            
        Case 6 'random colors per key
            scrSpeed.value = 2000 - 100
            
        Case 7 'random key color change
            scrSpeed.value = 2000 - 10
            
        Case 8 'random colored arrows
            scrSpeed.value = 2000 - 50
        
        Case 9 'random colored spiral fill
            scrSpeed.value = 2000 - 10
            
        Case 10 'flappy bird
            scrSpeed.value = 2000 - 20
            
        Case 11 'solid color cycle
            scrSpeed.value = 2000 - 100
            
        Case 12 'game of life maybe?
            scrSpeed.value = 2000 - 500
            
        Case 13 'starfield
            scrSpeed.value = 2000 - 200
        
        Case 14 'fireworks
            scrSpeed.value = 2000 - 80
        
        Case 15 'raindrops
            scrSpeed.value = 2000 - 20
        
        Case 16 'storm/thunder/rain
            scrSpeed.value = 2000 - 5
        
        Case 17 'rainbow spiral fill
            scrSpeed.value = 2000 - 5
            
        Case 18 'smooth rainbow
            scrSpeed.value = 2000 - 50
    End Select
    
    SaveSetting "KeyboardLights", "Animation", "AnimationsIndex", lstAnimations.ListIndex
End Sub


Sub Random_Weather_Storm()
    Static Tick As Long
    
    
    Tick = Tick + 1
    
    Static NewFlashes As Long
    
    
    If Tick > NewFlashes Then 'perform lightning flashes (whole keyboard
        
        Select Case Tick - NewFlashes
            Case 0 To 2, 5 To 7, 10, 13
                SolidLights 0, 0, 0
            Case 3, 4, 8, 9, 11
                SolidLights 255, 255, 255
            
            Case Is > 49
                NewFlashes = Rnd * 450
                
            Case Else
                
                
                
                
        End Select
        
        
    End If
    
    
    
    
    
    If Tick > 500 Then
        Tick = 0
    End If
    
End Sub





Private Sub lstKeydownEffect_Click()
    SaveSetting "KeyboardLights", "Animation", "KeydownEffectIndex", lstKeydownEffect.ListIndex
End Sub

Private Sub NTService_Continue(Success As Boolean)
    Success = True
End Sub

Private Sub NTService_Pause(Success As Boolean)
    Success = True
End Sub

Private Sub NTService_Start(Success As Boolean)
    '-- Add code to do what you want to do
    '   when the service starts
    '-- Set the success flag if successful
    Success = True
End Sub



Private Sub NTService_Stop()
    Unload Me
End Sub

Private Sub picKeyboard_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    KeySelectorX = X
    KeySelectorY = y
    shpKey.Left = X
    shpKey.Top = y
    shpKey.Width = 2
    shpKey.Height = 2
    shpKey.Visible = True
    KeySelectorDragging = True
    
    If Shift = 2 Then
        KeySelected = KeySelectedSemi
    Else
        ReDim KeySelected(0 To 144)
    End If
    
End Sub

Private Sub picKeyboard_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim i As Long
    Dim totalselected As Long
    
    If KeySelectorDragging Then
        If X < KeySelectorX Then
            shpKey.Left = X
            shpKey.Width = KeySelectorX - X
        Else
            shpKey.Left = KeySelectorX
            shpKey.Width = X - KeySelectorX
        End If
        
        If y < KeySelectorY Then
            shpKey.Top = y
            shpKey.Height = KeySelectorY - y
        Else
            shpKey.Top = KeySelectorY
            shpKey.Height = y - KeySelectorY
        End If
        
        For i = 0 To UBound(KeyPlaces)
            If KeyPlaces(i).X >= shpKey.Left And KeyPlaces(i).X + KeyPlaces(i).Width <= shpKey.Left + shpKey.Width And _
               KeyPlaces(i).y >= shpKey.Top And KeyPlaces(i).y + KeyPlaces(i).Height <= shpKey.Top + shpKey.Height Then
                
                If KeySelected(i) = True Then
                    KeySelectedSemi(i) = False
                Else
                    KeySelectedSemi(i) = True
                    totalselected = totalselected + 1
                End If
                
            ElseIf KeySelected(i) = True Then
                KeySelectedSemi(i) = True
                totalselected = totalselected + 1
            Else
                KeySelectedSemi(i) = False
            End If
        Next i
        
        KeyMouseSelectionCount = totalselected
    Else
        KeyMouseOver = 144
        For i = 0 To UBound(KeyPlaces)
            If X >= KeyPlaces(i).X And X <= KeyPlaces(i).X + KeyPlaces(i).Width Then
                If y >= KeyPlaces(i).y And y <= KeyPlaces(i).y + KeyPlaces(i).Height Then
                    KeyMouseOver = i
                    Exit For
                End If
            End If
        Next i
    End If
    
    Paint_Statusbar
End Sub

Function ColorToHex(R As Byte, G As Byte, B As Byte) As String
    Dim tR As String
    Dim tG As String
    Dim tB As String
    
    tR = Hex(R)
    tG = Hex(G)
    tB = Hex(B)
    
    If Len(tR) = 1 Then tR = "0" & tR
    If Len(tG) = 1 Then tG = "0" & tG
    If Len(tB) = 1 Then tB = "0" & tB
    
    ColorToHex = tR & tG & tB
End Function


Private Sub picKeyboard_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    KeySelectorDragging = False
    shpKey.Visible = False



' piece of code that helps with positioning the keys in the picture.

    'Static Results(0 To 200) As String
'    Static keyCount As Long
'    Static Initialized As Boolean
'
'    If Initialized = False Then
'        keyCount = 0
'        Initialized = True
'        Results(0) = "Split("""
'
'        picKeyboard.Cls
'        picKeyboard.ForeColor = vbRed
'        picKeyboard.Print KeyNames(keyCount)
'        picKeyboard.Print CStr(keyCount)
'        Exit Sub
'    End If
'
'    If Button = 1 Then
'
'        keyCount = keyCount + 1
'
'    ElseIf Button = 2 Then
'        keyCount = keyCount - 1
'    End If
    
    
    
'    picKeyboard.Cls
'    picKeyboard.ForeColor = vbRed
'    picKeyboard.Print KeyNames(keyCount)
'    picKeyboard.Print CStr(keyCount)
'
'    Dim i As Long
'    Results(0) = "Split("""
'    For i = 0 To 144
'        Results(i + 1) = InputBox(KeyNames(i) & vbCrLf & vbCrLf & "width,height")
'    Next i
'    Clipboard.Clear
'    Clipboard.SetText Join(Results, "|") & """,""|"")"
'




' piece of code that helps with REpositioning the keys in the picture.
'    Dim i As Long
'    Dim tmpStr() As String
'    If Button = 1 Then
'        For i = 0 To UBound(KeyPlaces)
'            If X >= KeyPlaces(i).X And X <= KeyPlaces(i).X + KeyPlaces(i).Width Then
'                If Y >= KeyPlaces(i).Y And Y <= KeyPlaces(i).Y + KeyPlaces(i).Height Then
'                    tmpStr = Split(InputBox("x,y,width,height for the key: " & KeyNames(i), , KeyPlaces(i).X & "," & KeyPlaces(i).Y & "," & KeyPlaces(i).Width & "," & KeyPlaces(i).Height), ",")
'                    KeyPlaces(i).X = CInt(tmpStr(0))
'                    KeyPlaces(i).Y = CInt(tmpStr(1))
'                    KeyPlaces(i).Width = CInt(tmpStr(2))
'                    KeyPlaces(i).Height = CInt(tmpStr(3))
'                    Exit For
'                End If
'            End If
'        Next i
'    ElseIf Button = 2 Then
'        Dim tmpStrCopy As String
'        For i = 0 To UBound(KeyPlaces)
'            tmpStrCopy = tmpStrCopy & KeyPlaces(i).X & "," & KeyPlaces(i).Y & "|"
'        Next i
'        tmpStrCopy = tmpStrCopy & vbCrLf & vbCrLf
'
'        For i = 0 To UBound(KeyPlaces)
'            tmpStrCopy = tmpStrCopy & KeyPlaces(i).Width & "," & KeyPlaces(i).Height & "|"
'        Next i
'
'        tmpStrCopy = tmpStrCopy & vbCrLf & vbCrLf
'
'        Clipboard.Clear
'        Clipboard.SetText tmpStrCopy
'    End If
    
    
End Sub

Private Sub picStatusBar_Click()
    Paint_Statusbar
End Sub

Private Sub scrSpeed_Change()
    tmrAnimation.Interval = 2000 - scrSpeed.value
    tmrAnimation.Enabled = False
    tmrAnimation.Enabled = True
    
    lblAnimationInterval.Caption = tmrAnimation.Interval
End Sub

Private Sub scrTransparency_Change()
    lblKeyboardTransparency.Caption = scrTransparency.value
End Sub

Private Sub scrTransparency_Scroll()
    scrTransparency_Change
    DoEvents
    Paint_Keyboard
End Sub

Private Sub scrWidth_Change()
    lblAnimationWidth.Caption = scrWidth.value
End Sub

Private Sub scrWidth_Scroll()
    scrWidth_Change
End Sub

Sub Init_Volume_Control()
    Dim lRes As Long
    Dim mayInit As Boolean
    Dim i As Long
    
    mayInit = False
    
    For i = 0 To 9
        lRes = OpenMixer(SPEAKER, i)
        If lRes = True Then
            m_bVolumeInitialized(i) = True
            m_LonVolume(i) = GetVolume(SPEAKER, i)
            m_LonVolumePrevious(i) = m_LonVolume(i)
            m_bVolumeMute(i) = GetMute(SPEAKER, i)
            
            'SelectMicrophone i
            
            'MsgBox "Result: " & lRes & vbCrLf & "Number: " & i & vbCrLf & "DeviceName: " & GetDeviceNameType(i) & vbCrLf & "Volume: " & m_LonVolume(i) & vbCrLf & "Mute: " & m_bVolumeMute(i)
            
            mayInit = True
        End If
    Next i
    
    If mayInit = False Then
        MsgBox "Could Not Initialize a single sound output interface!" & vbCrLf & vbCrLf & "Volume Control Disabled!", vbCritical
        Exit Sub
    End If
    
    tmrVolume.Enabled = True
    
End Sub


Sub Init_Keyboard_Matrix()
    Dim tmpFirstSplit() As String
    Dim tmpSecondSplit() As String
    Dim i As Long
    Dim j As Long

    tmpFirstSplit = Split("144,144,144,144,144,144,144,144,144,144,144,144,144,144,144,144,137,  8,144,144,144, 20,144,144" & vbCrLf & _
                          "  0,144, 12, 24, 36, 48, 60, 72, 84, 96,144,108,120,132,  6,144, 18, 30, 42,144, 32, 44, 56, 68" & vbCrLf & _
                          "  1, 13, 25, 37, 49, 61, 73, 85, 97,109,121,133,  7,144, 31,144, 54, 66, 78,144, 80, 92,104,116" & vbCrLf & _
                          "  2, 14, 26, 38, 50, 62, 74, 86, 98,110,122,134, 90,144,102,144, 43, 55, 67,144,  9, 21, 33,128" & vbCrLf & _
                          "  3,144, 15, 27, 39, 51, 63, 75, 87, 99,111,123,135,144,126,144,144,144,144,144, 57, 69, 81,144" & vbCrLf & _
                          "144,  4, 28, 40, 52, 64, 76, 88,100,112,124,136,144, 79,144,144,144,103,144,144, 93,105,117,140" & vbCrLf & _
                          "  5, 17, 29,144,144,144,144, 53,144,144, 89,101,113,144, 91,144,115,127,139,144,129,144,141,144", vbCrLf)
                          
    KeyNames = Split("ESCAPE,TILDE,TAB,CAPSLOCK,SHIFT_LEFT,CTRL_LEFT,F12,EQUALSIGN,WINDOWS_LOCK,NUM_7,NONE,NONE,F1,1,Q,A,NONE,WINDOWS_LEFT,PRINTSCREEN,NONE,MEDIA_MUTE,NUM_8,NONE,NONE,F2,2,W,S,Z,ALT_LEFT,SCROLLLOCK,BACKSPACE,MEDIA_STOP,NUM_9,NONE,NONE,F3,3,E,D,X,NONE,PAUSE,DELETE,MEDIA_PREVIOUS,NONE,NONE,NONE,F4,4,R,F,C,SPACE,INSERT,END,MEDIA_PLAY,NUM_4,NONE,NONE,F5,5,T,G,V,NONE,HOME,PAGEDOWN,MEDIA_NEXT,NUM_5,NONE,NONE,F6,6,Y,H,B,NONE,PAGEUP,SHIFT_RIGHT,NUM_NUMLOCK,NUM_6,NONE,NONE,F7,7,U,J,N,ALT_RIGHT,BLOK_HAAK_SLUIT,CTRL_RIGHT,NUM_SLASH,NUM_1,NONE,NONE,F8,8,I,K,M,WINDOWS_RIGHT,SLASH_BACK,ARROW_UP,NUM_ASTERIX,NUM_2,NONE,NONE,F9,9,O,L,COMMA,MENU,NONE,ARROW_LEFT,NUM_MIN,NUM_3,NONE,NONE,F10,0,P,SEMICOLON,PERIOD,NONE,ENTER,ARROW_DOWN,NUM_PLUS,NUM_0,NONE,NONE,F11,MIN,BLOK_HAAK_OPEN,QUOTE,SLASH_FORWARD,BRIGHTNESS,NONE,ARROW_RIGHT,NUM_ENTER,NUM_DELETE,NONE,NONE,NONE", ",")

    For i = 0 To UBound(tmpFirstSplit)
        tmpSecondSplit = Split(tmpFirstSplit(i), ",")
        For j = 0 To UBound(tmpSecondSplit)
            KeyMatrix(j, i) = CByte(tmpSecondSplit(j))
            KeyMatrixReversed(CByte(tmpSecondSplit(j))).X = j
            KeyMatrixReversed(CByte(tmpSecondSplit(j))).y = i
        Next j
    Next i


    tmpFirstSplit = Split("&H1B,&HC0,&H9,&H14,&HA0,&HA2,&H7B,&HBB,&H0,&H67,&H0,&H0,&H70,&H31,&H51,&H41,&H0,&H5B,&H2C,&H0,&HAD,&H68,&H0,&H0,&H71,&H32,&H57,&H53,&H5A,&HA4,&H91,&H8,&HB2,&H69,&H0,&H0,&H72,&H33,&H45,&H44,&H58,&H0,&H13,&H2E,&HB1,&H0,&H0,&H0,&H73,&H34,&H52,&H46,&H43,&H20,&H2D,&H23,&HB3,&H64,&H0,&H0,&H74,&H35,&H54,&H47,&H56,&H0,&H24,&H22,&HB0,&H65,&H0,&H0,&H75,&H36,&H59,&H48,&H42,&H0,&H21,&HA1,&H90,&H66,&H0,&H0,&H76,&H37,&H55,&H4A,&H4E,&HA5,&HDD,&HA3,&H6F,&H61,&H0,&H0,&H77,&H38,&H49,&H4B,&H4D,&H5C,&HDC,&H26,&H6A,&H62,&H0,&H0,&H78,&H39,&H4F,&H4C,&HBC,&H5D,&H0,&H25,&H6D,&H63,&H0,&H0,&H79,&H30,&H50,&HBA,&HBE,&H0,&HD,&H28,&H6B,&H60,&H0,&H0,&H7A,&HBD,&HDB,&HDE,&HBF,&H0,&H0,&H27,&HD,&H6E,&H0,&H0,&H0", ",")

    For i = 0 To 144
        KeyVKCodes(i) = Val(tmpFirstSplit(i))
    Next i
    
End Sub

Sub Init_Keyboard_Connection()
    On Error GoTo init_failed:
    UsbInit
    
    'UsbSetDebug (255)
    
    keyboard = UsbOpen(0, &H1B1C, &H1B13)
    
    If keyboard = 0 Then
        MsgBox "Could not connect to the Keyboard. Do you have the drivers installed that come with this program?", vbCritical
        keyboard = -1
        Exit Sub
    End If
    
    'interface = 0
    
'SET THE COLOR OF MY LOGITECH G9 LASER MOUSE
'    keyboard = UsbOpen(0, &H46D, &HC048)
'    interface = 1
'
'    Debug.Print UsbClaimInterface(keyboard, interface)
'    Dim byteswritten As Long
'    Dim l(0 To 100) As Byte
'
'    l(0) = &H10
'    l(1) = &H0
'    l(2) = &H80
'    l(3) = &H57
'    l(4) = &HFF
'    l(5) = &HFF
'    l(6) = &HFF
'
'    byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H210, &H1, l(0), 7, 500)
'    Debug.Print "byteswritten: " & byteswritten
'
'    UsbClearHalt keyboard, interface
'
'    Debug.Print "Release: " & UsbReleaseInterface(keyboard, interface)
'    Debug.Print "Close: " & UsbClose(keyboard)
 
        
Exit Sub

'    l(0) = &H11
'    l(1) = &H0
'    l(2) = &H93
'    l(3) = &H10
'    l(4) = &HFF
'    l(5) = &HFF
'    l(6) = &HFF
'    l(7) = &HFF
'    l(8) = &HFF
'    l(9) = &HFF
'    l(10) = &HFF
'    l(11) = &H6E
'    l(12) = &H0
'    l(13) = &H0
'    l(14) = &H0
'    l(15) = &H3C
'
'    l(16) = &H0
'    l(17) = &H82
'    l(18) = &H0
'    l(19) = &H0


init_failed:
    MsgBox "Could not initialize USBLIBVB0.dll please place it in the same folder as the exe or place it in the system32 or in syswow64!", vbCritical
    keyboard = -1
End Sub

Sub Init_Keyboard_KeyPositions()
    Dim tmpSplit1() As String
    Dim tmpSplit2() As String
    Dim i As Long
    Dim tS() As String
    

    
    tmpSplit1 = Split("16,65,27,30|16,111,27,30|16,152,49,31|16,194,59,30|16,236,80,29|16,277,48,30|600,65,27,30|517,111,26,30|696,14,24,24|790,152,26,31|||86,65,26,30|58,111,26,30|79,152,27,31|89,194,27,30||79,277,26,30|653,65,26,30||825,14,41,25|831,152,27,31|||128,65,26,30|100,111,26,30|121,152,27,31|131,194,26,30|110,236,27,29|120,277,38,30|695,65,26,30|559,111,68,30|786,63,38,25|873,152,26,31|||170,65,26,30|142,111,26,30|163,152,27,31|173,194,26,30|152,236,26,29||737,65,26,30|653,152,26,31|826,63,38,25||||211,65,26,30|184,111,26,30|204,152,27,31|215,194,26,30|194,236,26,29|173,277,256,30|653,111,26,30|695,152,26,31|866,63,38,25|790,194,26,30|||281,65,26,30|225,111,26,30|246,152,27,31|257,194,26,30|236,236,26,30||695,111,26,30|737,152,26,31|906,63,39,25|831,194,27,30|||323,65,26,30|267,111,26,30|288,152,27,31|298,194,26,30|277,236,26,29||737,111,26,30|527,236,100,29|790,111,26,30|873,194,26,30|||364,65,26,30|" & _
                      "309,111,26,30|329,152,27,31|340,194,26,30|318,236,26,29|444,277,37,30|538,152,27,31|580,277,47,30|831,111,27,30|790,235,26,30|||406,65,26,30|350,111,26,30|371,152,27,31|381,194,26,30|361,236,26,29|496,277,26,30|579,152,48,31|695,236,26,29|873,111,26,30|831,235,27,30|||475,65,26,30|392,111,26,30|413,152,27,31|423,194,26,30|402,236,26,29|537,277,27,30||653,277,26,30|914,111,27,31|873,235,26,30|||517,65,26,30|434,111,26,30|454,152,27,31|465,194,26,30|444,236,26,29||548,194,79,30|695,277,26,30|914,152,27,72|790,277,68,30|||559,65,26,30|475,111,26,30|496,152,27,31|506,194,27,30|486,236,26,29|655,14,24,24||737,277,26,30|914,235,27,72|873,277,26,30|||", "|")

    Dim newArray As String
    Dim lcount As Long
    
    newArray = "Split("""
    For i = 0 To UBound(tmpSplit1)
        
        If tmpSplit1(i) <> "" And KeyNames(i) <> "NONE" Then
            tS = Split(tmpSplit1(i), ",")
            KeyPlaces(i).X = tS(0)
            KeyPlaces(i).y = tS(1)
            KeyPlaces(i).Width = tS(2)
            KeyPlaces(i).Height = tS(3)
            lcount = lcount + 1
            'newArray = newArray & tS(0) & "," & tS(1) & "," & tS(2) & "," & tS(3) & "|"
        'Else
            'newArray = newArray & "|"
        End If
        
'        If lcount >= 10 Then
'            lcount = 0
'            newArray = newArray & """ & _" & vbCrLf & """"
'        End If
        
        
    Next i
    
    'newArray = newArray & """,""|"")"
    
    '''''''CODE FOR PRINTING EVERYTHING OUT TO REORDER THE KEYS
'    newArray = ""
'
'    For i = 0 To 144
'        newArray = newArray & "MATRIX_X:" & KeyMatrixReversed(i).X & vbTab & "MATRIX_Y:" & KeyMatrixReversed(i).Y
'
'        newArray = newArray & vbTab & "VK_CODE:&H" & Hex(KeyVKCodes(i))
'
'        newArray = newArray & vbTab & "POSITION:" & KeyPlaces(i).X & "," & KeyPlaces(i).Y & "," & KeyPlaces(i).Width & "," & KeyPlaces(i).Height
'
'        newArray = newArray & vbTab & KeyNames(i)
'
'        newArray = newArray & vbCrLf & vbCrLf
'
'    Next i
'


'''''''CODE FOR GETTING FROM CLIPBOARD AND CONVERT BACK TO SPLIT ARRAYS TO REORDER THE KEYS
'    newArray = ""
'
'    newArray = Clipboard.GetText
'
'    tmpSplit1 = Split(newArray, vbCrLf & vbCrLf)
'
'    Dim lKeyMatrix(0 To 23, 0 To 6) As Byte
'
'    Dim tmpSplit3() As String
'    Dim VKCODES As String
'    Dim KEYPOSITIONS As String
'    Dim SKEYNAMES As String
'    Dim PIXELMETRIX As String
'    Dim KEYENUM As String
'
'    KEYENUM = "Private Enum Keys" & vbCrLf
'    PIXELMETRIX = "tmpFirstSplit = Split("
'    SKEYNAMES = "KeyNames = Split("""
'    KEYPOSITIONS = "tmpSplit1 = Split("""
'    VKCODES = "tmpFirstSplit = Split("""
'
'    Dim tmpX As Long
'    Dim tmpY As Long
'
'    For tmpY = 0 To 6
'        For tmpX = 0 To 23
'            lKeyMatrix(tmpX, tmpY) = 144
'        Next tmpX
'    Next tmpY
'
'    For i = 0 To UBound(tmpSplit1) - 1
'        tmpSplit2 = Split(tmpSplit1(i), vbTab)
'
'        tmpSplit3 = Split(tmpSplit2(0), ":")
'        tmpX = CLng(tmpSplit3(1))
'
'        tmpSplit3 = Split(tmpSplit2(1), ":")
'        tmpY = CLng(tmpSplit3(1))
'
'        lKeyMatrix(tmpX, tmpY) = i
'
'        tmpSplit3 = Split(tmpSplit2(2), ":")
'        VKCODES = VKCODES & tmpSplit3(1) & ","
'
'        tmpSplit3 = Split(tmpSplit2(3), ":")
'        If tmpSplit3(1) = "0,0,0,0" Then
'            KEYPOSITIONS = KEYPOSITIONS & "|"
'        Else
'            KEYPOSITIONS = KEYPOSITIONS & tmpSplit3(1) & "|"
'        End If
'
'        tmpSplit3 = Split(tmpSplit2(4), ":")
'        SKEYNAMES = SKEYNAMES & tmpSplit3(0) & ","
'        If tmpSplit3(0) = "NONE" Then
'            lKeyMatrix(tmpX, tmpY) = 144
'        Else
'            KEYENUM = KEYENUM & vbTab & "KEY_" & tmpSplit3(0) & " = " & i & vbCrLf
'        End If
'
'    Next i
'
'
'    For tmpY = 0 To 6
'
'        PIXELMETRIX = PIXELMETRIX & """"
'
'        For tmpX = 0 To 23
'            PIXELMETRIX = PIXELMETRIX & somestring(CStr(lKeyMatrix(tmpX, tmpY)), 3) & ","
'        Next tmpX
'
'        PIXELMETRIX = Left$(PIXELMETRIX, Len(PIXELMETRIX) - 1)
'
'        If tmpY < 6 Then
'            PIXELMETRIX = PIXELMETRIX & """ & vbCrLf & _" & vbCrLf
'        End If
'
'    Next tmpY
'
'    KEYENUM = KEYENUM & "End Enum" & vbCrLf
'
'    PIXELMETRIX = PIXELMETRIX & """, vbCrLf)"
'
'    VKCODES = Left$(VKCODES, Len(VKCODES) - 1)
'    VKCODES = VKCODES & """, "","")"
'
'    KEYPOSITIONS = Left$(KEYPOSITIONS, Len(KEYPOSITIONS) - 1)
'    KEYPOSITIONS = KEYPOSITIONS & """, ""|"")"
'
'    SKEYNAMES = Left$(SKEYNAMES, Len(SKEYNAMES) - 1)
'    SKEYNAMES = SKEYNAMES & """, "","")"
'
'    Clipboard.Clear
'    Clipboard.SetText VKCODES & vbCrLf & vbCrLf & KEYPOSITIONS & vbCrLf & vbCrLf & SKEYNAMES & vbCrLf & vbCrLf & PIXELMETRIX & vbCrLf & KEYENUM & vbCrLf
    
End Sub

Function somestring(sinput As String, ilength As Long) As String
    
    If Len(sinput) < ilength Then
        somestring = String(ilength - Len(sinput), " ") & sinput
    Else
        somestring = sinput
    End If
    
End Function




Sub Paint_Statusbar()
    Dim BF As BLENDFUNCTION
    Dim lBF As Long
    
    Dim tmpStr As String
    Dim lTextCenter As Long
    
    lTextCenter = picStatusBarMemory.ScaleHeight / 2 - picStatusBarMemory.TextHeight("WQR") / 2
    
    
    
    picStatusBarMemory.Picture = LoadPicture
    
    picStatusBarMemory.CurrentX = lTextCenter
    picStatusBarMemory.CurrentY = lTextCenter
    
    tmpStr = "Key under mouse: " & KeyNames(KeyMouseOver)
    picStatusBarMemory.Print tmpStr
    
    
    
    picStatusBarMemory.CurrentX = lTextCenter + 200
    picStatusBarMemory.CurrentY = lTextCenter
    
    tmpStr = "Key Color: " & ColorToHex(KeyColor(KeyMouseOver, Red), KeyColor(KeyMouseOver, Green), KeyColor(KeyMouseOver, Blue)) '& " " & KeyPlaces(KeyMouseOver).X & " " & KeyPlaces(KeyMouseOver).y
    picStatusBarMemory.Print tmpStr
    
    
    If KeyMouseSelectionCount > 0 Then
        picStatusBarMemory.CurrentX = lTextCenter + 300
        picStatusBarMemory.CurrentY = lTextCenter
        
        tmpStr = "Keys Selected: " & KeyMouseSelectionCount
        picStatusBarMemory.Print tmpStr
    End If
    
    
    Set picStatusBarMemory.Picture = picStatusBarMemory.Image
    
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = 128
        .AlphaFormat = 0
    End With
    
    RtlMoveMemory lBF, BF, 4

    picStatusBar.Cls
    picStatusBar.BackColor = vbBlack
    AlphaBlend picStatusBar.hdc, 0, 0, picStatusBar.ScaleWidth, picStatusBar.ScaleHeight, picStatusBarMemory.hdc, 0, 0, picStatusBar.ScaleWidth, picStatusBar.ScaleHeight, lBF
    picStatusBar.Refresh
    
    
    
End Sub


Sub Paint_Keyboard()
    Dim polyDraw() As POINT
    Dim i As Long
    ReDim polyDraw(0 To 3)
    
    Dim pBufferHDC As Long
    Dim pBufferBitmap As Long
    Dim pHDC As Long
    Dim pBitmap As Long
    Dim pOldBitmap As Long
    
    
    picKeyboard.Cls
    picMemory.ForeColor = vbWhite
    
    Dim BF As BLENDFUNCTION
    Dim lBF As Long
    
    With BF
        .BlendOp = 0 'AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = scrTransparency.value '128
        .AlphaFormat = 0
    End With
    
    RtlMoveMemory lBF, BF, 4
    
    For i = 0 To 144
        'picKeyboard.Cls
        
        'picMemory.BackColor = vbBlack
        
        polyDraw(0).X = 0
        polyDraw(0).y = 0
        
        polyDraw(1).X = KeyPlaces(i).Width
        polyDraw(1).y = 0
        
        polyDraw(2).X = polyDraw(1).X
        polyDraw(2).y = KeyPlaces(i).Height
        
        polyDraw(3).X = 0
        polyDraw(3).y = polyDraw(2).y
        
        
        
        picMemory.FillColor = RGB(KeyColor(i, 0), KeyColor(i, 1), KeyColor(i, 2))
        'picMemory.DrawMode = 13
        If KeySelectedSemi(i) Then
            picMemory.DrawWidth = 3
        Else
            picMemory.DrawWidth = 1
        End If
        
        Polygon picMemory.hdc, polyDraw(0), 4
        'picKeyboard.Refresh
        
        Select Case scrTransparency.value
        
            Case 255
                BitBlt picKeyboard.hdc, KeyPlaces(i).X, KeyPlaces(i).y, KeyPlaces(i).Width + 1, KeyPlaces(i).Height + 1, picMemory.hdc, 0, 0, vbSrcCopy
            
            Case Is > 0
                AlphaBlend picKeyboard.hdc, KeyPlaces(i).X, KeyPlaces(i).y, KeyPlaces(i).Width + 1, KeyPlaces(i).Height + 1, picMemory.hdc, 0, 0, KeyPlaces(i).Width + 1, KeyPlaces(i).Height + 1, lBF
        End Select
        
        
    Next i
    
    
    
    Paint_Statusbar
End Sub


''''Sub Paint_Keyboard()
''''    Dim polyDraw() As Point
''''    Dim i As Long
''''    ReDim polyDraw(0 To 3)
''''
''''    Dim pBufferHDC As Long
''''    Dim pBufferBitmap As Long
''''    Dim pHDC As Long
''''    Dim pBitmap As Long
''''    Dim pOldBitmap As Long
''''
''''
''''    picKeyboard.Cls
''''    picMemory.Picture = LoadPicture
''''    picMemory.Cls
''''    picMemory.BackColor = vbBlack
''''
''''    picMemory.FillStyle = 0
''''    picMemory.ForeColor = vbWhite
''''
''''    For i = 0 To 144
''''        polyDraw(0).X = KeyPlaces(i).X
''''        polyDraw(0).y = KeyPlaces(i).y
''''
''''        polyDraw(1).X = polyDraw(0).X + KeyPlaces(i).Width
''''        polyDraw(1).y = polyDraw(0).y
''''
''''        polyDraw(2).X = polyDraw(1).X
''''        polyDraw(2).y = polyDraw(0).y + KeyPlaces(i).Height
''''
''''        polyDraw(3).X = KeyPlaces(i).X
''''        polyDraw(3).y = polyDraw(2).y
''''
''''        picMemory.FillColor = RGB(KeyColor(i, 0), KeyColor(i, 1), KeyColor(i, 2))
''''        picMemory.DrawMode = 13
''''        If KeySelectedSemi(i) Then
''''            picMemory.DrawWidth = 3
''''        Else
''''            picMemory.DrawWidth = 1
''''        End If
''''
''''        Polygon picMemory.hdc, polyDraw(0), 4
''''        'Polygon newPicHDC, polyDraw(0), 4
''''    Next i
''''
''''    Set picMemory.Picture = picMemory.Image
''''
''''    Dim BF As BLENDFUNCTION
''''    Dim lBF As Long
''''
''''    With BF
''''        .BlendOp = 0 'AC_SRC_OVER
''''        .BlendFlags = 0
''''        .SourceConstantAlpha = scrTransparency.value '128
''''        .AlphaFormat = 0
''''    End With
''''
''''    RtlMoveMemory lBF, BF, 4
''''
''''    pBufferHDC = CreateCompatibleDC(picMemory.hdc)
''''    pBufferBitmap = CreateCompatibleBitmap(picMemory.hdc, picMemory.ScaleWidth, picMemory.ScaleHeight)
''''    pOldBitmap = SelectObject(pBufferHDC, pBufferBitmap)
''''
''''
''''    pHDC = CreateCompatibleDC(pBufferHDC)
''''    'GetObject
''''
''''
''''    'pBitmap = CreateCompatibleBitmap(newPicHDC, picKeyboard.ScaleWidth, picKeyboard.ScaleHeight)
''''    'http://forums.codeguru.com/showthread.php?498661-can-i-combine-the-transparentblt%28%29-with-alphablend%28%29-api-functions
''''    'http://www.vbforums.com/showthread.php?652136-RESOLVED-Draw-semi-transparent-rectangle-in-picture-box
''''
''''
''''    AlphaBlend pBufferHDC, 0, 0, picKeyboard.ScaleWidth, picKeyboard.ScaleHeight, picMemory.hdc, 0, 0, picKeyboard.ScaleWidth, picKeyboard.ScaleHeight, lBF
''''
''''    Debug.Print TransparentBlt(picKeyboard.hdc, 0, 0, picKeyboard.ScaleWidth, picKeyboard.ScaleHeight, pBufferHDC, 0, 0, picKeyboard.ScaleWidth, picKeyboard.ScaleHeight, vbBlack)
''''
''''    picKeyboard.Refresh
''''
''''    DeleteDC pBufferHDC
''''    DeleteObject pBufferBitmap
''''
''''    Paint_Statusbar
''''End Sub

Sub slp()
    Sleep 10
    DoEvents
End Sub

Sub ShowForm()
        
    ScanForMonitors
    
    Select Case UCase(Command)
        Case "-I", "/I"
            NTService.Interactive = True
            If NTService.Install Then
                MsgBox NTService.DisplayName & " installed successfully."
            Else
                MsgBox NTService.DisplayName & " did not install successfully."
            End If
            End
        Case "-U", "/U"
            If NTService.Uninstall Then
                MsgBox NTService.DisplayName & " uninstalled successfully."
            Else
                MsgBox NTService.DisplayName & " did not uninstall successfully."
            End If
            End
        Case ""
            '-- Application is started without parameters
            NTService.Interactive = True
            NTService.StartService
        Case Else
            MsgBox "The parameter: " & Command & " was not understood. Try -I (install) or -U (uninstall)."
            End
    End Select
    
    ReDim KeySelected(0 To 144)
    ReDim KeySelectedSemi(0 To 144)
    ReDim KeyVKCodes(0 To 144)
    ReDim KeyVKDown(0 To 144)
    ReDim KeyVKDownPrevious(0 To 144)
    ReDim KeyVKDownColors(0 To 144)
    
    
    Init_Volume_Control

    Init_Keyboard_Matrix
    
    Init_Keyboard_Messages
    
    Init_Keyboard_Connection

    Init_Keyboard_KeyPositions

    Paint_Keyboard

    Paint_Statusbar

    picMemory.Left = picKeyboard.Left
    picMemory.Top = picKeyboard.Top
    picMemory.Width = picKeyboard.Width
    picMemory.Height = picKeyboard.Height

    
                
    Randomize
    
    'lstAnimations.ListIndex = 0
    
    tmrVolume.Enabled = True
    
    Me.Visible = True
    bLoop = True
    
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    Dim i As Long
    
    scrSolid(0).value = 0
    scrSolid(1).value = 127
    scrSolid(2).value = 255
    
    Dim ListSize As Long
    Dim j As Long
    
    Dim t As TimerData
    InitPerformanceTimer t
    
    lstAnimations.ListIndex = GetSetting("KeyboardLights", "Animation", "AnimationsIndex", 0)
    lstKeydownEffect.ListIndex = GetSetting("KeyboardLights", "Animation", "KeydownEffectIndex", 0)
    chk16Million.value = GetSetting("KeyboardLights", "Animation", "16MillionColors", 0)
    
    
    SaveSetting "KeyboardLights", "Animation", "AnimationsIndex", lstAnimations.ListIndex
    SaveSetting "KeyboardLights", "Animation", "16MillionColors", chk16Million.value
    SaveSetting "KeyboardLights", "Animation", "KeydownEffectIndex", lstKeydownEffect.ListIndex
    
    SetKeyboardHook
    
    Do
        QueryPerformanceCounter t.StartCount
    
        'bLooping = True
        'tmrAnimation_Timer
        'ClearLeds
        'SolidLights 7 - scrSolid(0).value, 7 - scrSolid(1).value, 7 - scrSolid(2).value
        
        Call Paint_Volume: slp
        
        Call Paint_KeyDown ': slp
        
        ListSize = lstSolidKeys.ListCount - 1
        If ListSize > -1 Then
            For i = 0 To ListSize
                With KeySolidList(i)
                    For j = 0 To UBound(.s_Keylist)
                        KeyColor(.s_Keylist(j), Red) = .lColor.R
                        KeyColor(.s_Keylist(j), Green) = .lColor.G
                        KeyColor(.s_Keylist(j), Blue) = .lColor.B
                    Next j

                End With
            Next i
        End If
        
        Call Paint_Keyboard ': slp
        
        QueryPerformanceCounter t.StopCount
        TimeQuerys(Timers.t_ProgramLoop) = ElapsedTime(t)
        
        Call SendData: slp
        
        'bLooping = False
    Loop While bLoop
    
    Unload Me
End Sub

Sub Paint_KeyDown()
    Dim lByte(0 To 255) As Byte
    Dim i As Long
    
    GetKeyboardState lByte(0)
    
'    Dim tmpTest As String
'
'    tmpTest = ""
'    For i = 0 To 144
'
'        tmpTest = tmpTest & KeyNames(i) & "=" & lByte(KeyVKCodes(i)) & vbCrLf
'    Next i
'
'    Debug.Print tmpTest
    
    Dim keyEffect As Long
    
    
    keyEffect = lstKeydownEffect.ListIndex
    
    
    For i = 0 To 144
        KeyVKDown(i) = (lByte(KeyVKCodes(i)) And 128)
        If KeyVKDown(i) Then
            If KeyVKDownPrevious(i) = False Then
                GetRandomColor KeyVKDownColors(i).R, KeyVKDownColors(i).G, KeyVKDownColors(i).B, chkAnimation(0).value = vbChecked
                KeyVKDownPrevious(i) = True
            End If
            
            Select Case keyEffect
                Case 1
                    SetLed CLng(KeyMatrixReversed(i).X), CLng(KeyMatrixReversed(i).y), KeyVKDownColors(i).R, KeyVKDownColors(i).G, KeyVKDownColors(i).B
                
                Case 2
                    FillRow CLng(KeyMatrixReversed(i).y), KeyVKDownColors(i).R, KeyVKDownColors(i).G, KeyVKDownColors(i).B
                
                Case 3
                    FillCol CLng(KeyMatrixReversed(i).X), KeyVKDownColors(i).R, KeyVKDownColors(i).G, KeyVKDownColors(i).B
                
                Case 4
                    FillRow CLng(KeyMatrixReversed(i).y), KeyVKDownColors(i).R, KeyVKDownColors(i).G, KeyVKDownColors(i).B
                    FillCol CLng(KeyMatrixReversed(i).X), KeyVKDownColors(i).R, KeyVKDownColors(i).G, KeyVKDownColors(i).B
                
                Case 5
                    Make_Circle CLng(KeyMatrixReversed(i).X), CLng(KeyMatrixReversed(i).y), 2, KeyVKDownColors(i).R, KeyVKDownColors(i).G, KeyVKDownColors(i).B
                
                Case 6
                    Make_Circle CLng(KeyMatrixReversed(i).X), CLng(KeyMatrixReversed(i).y), 0.8, KeyVKDownColors(i).R, KeyVKDownColors(i).G, KeyVKDownColors(i).B
                    
            End Select
            

            '

            '
            'Make_Circle CLng(KeyMatrixReversed(i).X), CLng(KeyMatrixReversed(i).Y), 1, keyVKDownColors(i).R, keyVKDownColors(i).G, keyVKDownColors(i).B
        Else
            KeyVKDownPrevious(i) = False
        End If
    Next i
    
End Sub

Sub Paint_Volume()
    Dim i As Long
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    
    For i = 0 To 8
        If m_bVolumeMute(i) Then
            SetLedByName Keys.KEY_MEDIA_MUTE, 255, 0, 0
                
        ElseIf m_bVolumeEffect(i) Then
            If m_LonVolume(i) > 0 Then
                If m_LonVolume(i) < 50 Then
                    R = 255
                    G = 255 / 50 * m_LonVolume(i)
                    B = 0
                ElseIf m_LonVolume(i) >= 50 And m_LonVolume(i) < 100 Then
                    R = 255 - (255 / 50 * (m_LonVolume(i) - 50))
                    G = 255
                    B = 0
                Else
                    R = 255
                    G = 0
                    B = 255
                End If
            Else
                R = 0
                G = 0
                B = 255
            End If
            
            SetLedByName Keys.KEY_MEDIA_MUTE, R, G, B
        End If
    Next i
End Sub


Sub SolidLights(R As Byte, G As Byte, B As Byte)
    Dim i As Long
    For i = 0 To 144
        KeyColor(i, Red) = R
        KeyColor(i, Green) = G
        KeyColor(i, Blue) = B
    Next i
    
End Sub

Sub RandomLights()
    Randomize
    Dim i As Long
    Dim j As Long
    
    For i = 0 To 144
        For j = Red To Blue
            KeyColor(i, j) = Rnd * 255
        Next j
    Next i
End Sub


Sub NUMberToNUMpad(lNUMber As Long, R As Byte, G As Byte, B As Byte)
    Dim lNUM As Long
    Dim i As Long
    Dim LightsToFill As Long
    
    
    LightsToFill = lNUMber \ 10
    lNUM = lNUMber Mod 10
    
    Select Case lNUM
        Case "0"
            SetLedByName Keys.KEY_NUM_0, R, G, B
        
        Case "1"
            SetLedByName Keys.KEY_NUM_1, R, G, B
            
        Case "2"
            SetLedByName Keys.KEY_NUM_2, R, G, B
            
        Case "3"
            SetLedByName Keys.KEY_NUM_3, R, G, B
            
        Case "4"
            SetLedByName Keys.KEY_NUM_4, R, G, B
            
        Case "5"
            SetLedByName Keys.KEY_NUM_5, R, G, B
            
        Case "6"
            SetLedByName Keys.KEY_NUM_6, R, G, B
            
        Case "7"
            SetLedByName Keys.KEY_NUM_7, R, G, B
            
        Case "8"
            SetLedByName Keys.KEY_NUM_8, R, G, B
            
        Case "9"
            SetLedByName Keys.KEY_NUM_9, R, G, B
            
    End Select
    
    Select Case LightsToFill
        Case 1
            SetLedByName Keys.KEY_NUM_NUMLOCK, R, G, B
        
        Case 2
            SetLedByName Keys.KEY_NUM_SLASH, R, G, B
        
        Case 3
            SetLedByName Keys.KEY_NUM_ASTERIX, R, G, B
        
        Case 4
            SetLedByName Keys.KEY_NUM_MIN, R, G, B
        
        Case 5
            SetLedByName Keys.KEY_NUM_PLUS, R, G, B
        
        Case 6
            SetLedByName Keys.KEY_NUM_ENTER, R, G, B
    End Select
        
End Sub

Sub Init_Keyboard_Messages()
    'RED PACKET
    Data(PACKET1) = &H7F
    Data(PACKET1 + 1) = &H1
    Data(PACKET1 + 2) = &H3C
    Data(PACKET1 + 3) = &H0
    
    Data(PACKET2) = &H7F
    Data(PACKET2 + 1) = &H2
    Data(PACKET2 + 2) = &H3C
    Data(PACKET2 + 3) = &H0
    
    Data(PACKET3) = &H7F
    Data(PACKET3 + 1) = &H3
    Data(PACKET3 + 2) = &H18
    Data(PACKET3 + 3) = &H0
     
    Data(PACKET4) = &H7
    Data(PACKET4 + 1) = &H28
    Data(PACKET4 + 2) = &H1
    Data(PACKET4 + 3) = &H3
    Data(PACKET4 + 4) = &H2
    
    
    'GREEN PACKET
    Data(PACKET5) = &H7F
    Data(PACKET5 + 1) = &H1
    Data(PACKET5 + 2) = &H3C
    Data(PACKET5 + 3) = &H0
    
    Data(PACKET6) = &H7F
    Data(PACKET6 + 1) = &H2
    Data(PACKET6 + 2) = &H3C
    Data(PACKET6 + 3) = &H0
    
    Data(PACKET7) = &H7F
    Data(PACKET7 + 1) = &H3
    Data(PACKET7 + 2) = &H18
    Data(PACKET7 + 3) = &H0
     
    Data(PACKET8) = &H7
    Data(PACKET8 + 1) = &H28
    Data(PACKET8 + 2) = &H2
    Data(PACKET8 + 3) = &H3
    Data(PACKET8 + 4) = &H2
    
    'BLUE PACKET
    Data(PACKET9) = &H7F
    Data(PACKET9 + 1) = &H1
    Data(PACKET9 + 2) = &H3C
    Data(PACKET9 + 3) = &H0
    
    Data(PACKET10) = &H7F
    Data(PACKET10 + 1) = &H2
    Data(PACKET10 + 2) = &H3C
    Data(PACKET10 + 3) = &H0
    
    Data(PACKET11) = &H7F
    Data(PACKET11 + 1) = &H3
    Data(PACKET11 + 2) = &H18
    Data(PACKET11 + 3) = &H0
    
    Data(PACKET12) = &H7
    Data(PACKET12 + 1) = &H28
    Data(PACKET12 + 2) = &H3
    Data(PACKET12 + 3) = &H3
    Data(PACKET12 + 4) = &H2
    
    
    'red
    Data(THREEBITPACKET + PACKET1) = &H7F
    Data(THREEBITPACKET + PACKET1 + 1) = &H1
    Data(THREEBITPACKET + PACKET1 + 2) = &H3C
    Data(THREEBITPACKET + PACKET1 + 3) = 0
    
    'red + green
    Data(THREEBITPACKET + PACKET2) = &H7F
    Data(THREEBITPACKET + PACKET2 + 1) = &H2
    Data(THREEBITPACKET + PACKET2 + 2) = &H3C
    Data(THREEBITPACKET + PACKET2 + 3) = 0

    'green + blue
    Data(THREEBITPACKET + PACKET3) = &H7F
    Data(THREEBITPACKET + PACKET3 + 1) = &H3
    Data(THREEBITPACKET + PACKET3 + 2) = &H3C

    'blue
    Data(THREEBITPACKET + PACKET4) = &H7F
    Data(THREEBITPACKET + PACKET4 + 1) = &H4
    Data(THREEBITPACKET + PACKET4 + 2) = &H24

    Data(THREEBITPACKET + PACKET5) = &H7
    Data(THREEBITPACKET + PACKET5 + 1) = &H27
    Data(THREEBITPACKET + PACKET5 + 2) = &H2
    Data(THREEBITPACKET + PACKET5 + 3) = &H0
    Data(THREEBITPACKET + PACKET5 + 4) = &H1
    Data(THREEBITPACKET + PACKET5 + 5) = &H1
    
End Sub


Sub SendData()
    Dim byteswritten As Long
    Dim j As Long
    Dim i As Long
    
    If keyboard = -1 Then Exit Sub
    
    Dim t As TimerData
    
    Dim KeysChanged As Boolean


    For i = 0 To 144
        If LastKeyColor(i, Red) <> KeyColor(i, Red) Then
            LastKeyColor(i, Red) = KeyColor(i, Red)
            KeysChanged = True
        End If
        
        If LastKeyColor(i, Green) <> KeyColor(i, Green) Then
            LastKeyColor(i, Green) = KeyColor(i, Green)
            KeysChanged = True
        End If
        
        If LastKeyColor(i, Blue) <> KeyColor(i, Blue) Then
            LastKeyColor(i, Blue) = KeyColor(i, Blue)
            KeysChanged = True
        End If
    Next i
    
    If KeysChanged = False Then GoTo EndOfTiming
    
    If chk16Million.value = vbChecked Then
        For i = 0 To 59
            
            Data(PACKET1 + 4 + i) = KeyColor(i, Red) '* 30
            Data(PACKET5 + 4 + i) = KeyColor(i, Green) '* 30
            Data(PACKET9 + 4 + i) = KeyColor(i, Blue) '* 30
            
            Data(PACKET2 + 4 + i) = KeyColor(i + 60, Red) '* 30
            Data(PACKET6 + 4 + i) = KeyColor(i + 60, Green) '* 30
            Data(PACKET10 + 4 + i) = KeyColor(i + 60, Blue) '* 30
        Next i
        
        For i = 0 To 23
            Data(PACKET3 + 4 + i) = KeyColor(i + 120, Red) '* 30
            
            Data(PACKET7 + 4 + i) = KeyColor(i + 120, Green) ' * 30
            
            Data(PACKET11 + 4 + i) = KeyColor(i + 120, Blue)  '* 30
        Next i
        
        
        If KeysChanged Then
            InitPerformanceTimer t
            QueryPerformanceCounter t.StartCount
        
            If chkCombine.value = vbChecked Then
                byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H200, &H3, Data(PACKET1), THREEBITPACKET, 500)
                'byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET1), 64 * 12, 1500)
            Else
                byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H200, &H3, Data(PACKET1), PACKET5, 500)
                byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H200, &H3, Data(PACKET5), PACKET5, 500)
                byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H200, &H3, Data(PACKET9), PACKET5, 500)
    
                'byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET1), 64 * 4, 1500)
                'byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET5), 64 * 4, 1500)
                'byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET9), 64 * 4, 1500)
            End If
            
            QueryPerformanceCounter t.StopCount
            TimeQuerys(Timers.t_SendData) = ElapsedTime(t)
            
        End If
    
    
    Else
        j = 0
        For i = 0 To 118 Step 2
            Data(THREEBITPACKET + PACKET1 + 4 + j) = Round(7 - 7 / 255 * KeyColor(i + 1, Red)) * 16 + Round(7 - 7 / 255 * KeyColor(i, Red))
            j = j + 1
        Next i: j = 0
        
        For i = 120 To 142 Step 2
            Data(THREEBITPACKET + PACKET2 + 4 + j) = Round(7 - 7 / 255 * KeyColor(i + 1, Red)) * 16 + Round(7 - 7 / 255 * KeyColor(i, Red))
            j = j + 1
        Next i
        
        For i = 0 To 92 Step 2
            Data(THREEBITPACKET + PACKET2 + 4 + j) = Round(7 - 7 / 255 * KeyColor(i + 1, Green)) * 16 + Round(7 - 7 / 255 * KeyColor(i, Green))
            j = j + 1
        Next i: j = 0
        
        For i = 96 To 142 Step 2
            Data(THREEBITPACKET + PACKET3 + 4 + j) = Round(7 - 7 / 255 * KeyColor(i + 1, Green)) * 16 + Round(7 - 7 / 255 * KeyColor(i, Green))
            j = j + 1
        Next i
        
        For i = 0 To 68 Step 2
            Data(THREEBITPACKET + PACKET3 + 4 + j) = Round(7 - 7 / 255 * KeyColor(i + 1, Blue)) * 16 + Round(7 - 7 / 255 * KeyColor(i, Blue))
            j = j + 1
        Next i: j = 0
        
        For i = 72 To 142 Step 2
            Data(THREEBITPACKET + PACKET4 + 4 + j) = Round(7 - 7 / 255 * KeyColor(i + 1, Blue)) * 16 + Round(7 - 7 / 255 * KeyColor(i, Blue))
            j = j + 1
        Next i
        
        
        If KeysChanged Then
            InitPerformanceTimer t
            QueryPerformanceCounter t.StartCount
        
            If chkCombine.value = vbChecked Then
                byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H200, &H3, Data(PACKET1 + THREEBITPACKET), PACKET6, 500)
                'byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET1), 64 * 12, 1500)
            Else
                byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H200, &H3, Data(PACKET1 + THREEBITPACKET), PACKET2, 500)
                byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H200, &H3, Data(PACKET2 + THREEBITPACKET), PACKET2, 500)
                byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H200, &H3, Data(PACKET3 + THREEBITPACKET), PACKET2, 500)
                byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H200, &H3, Data(PACKET4 + THREEBITPACKET), PACKET2, 500)
                byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H200, &H3, Data(PACKET5 + THREEBITPACKET), PACKET2, 500)
                
                'byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET1), 64 * 4, 1500)
                'byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET5), 64 * 4, 1500)
                'byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET9), 64 * 4, 1500)
            End If
            
            QueryPerformanceCounter t.StopCount
            TimeQuerys(Timers.t_SendData) = ElapsedTime(t)
            
        End If
    End If
    
    
'''''
'    For i = 0 To 118 Step 2
'        Data(PACKET1 + 4 + j) = KeyColor(i, Red) * 16 + KeyColor(i + 1, Red)
'        j = j + 1
'    Next i: j = 0
'
'    For i = 120 To 142 Step 2
'        Data((PACKET2 + 4 + j)) = KeyColor(i, Red) * 16 + KeyColor(i + 1, Red)
'        j = j + 1
'    Next i
'
'    For i = 0 To 92 Step 2
'        Data((PACKET2 + 4 + j)) = KeyColor(i, Green) * 16 + KeyColor(i + 1, Green)
'        j = j + 1
'    Next i: j = 0
'
'    For i = 96 To 142 Step 2
'        Data((PACKET3 + 4 + j)) = KeyColor(i, Green) * 16 + KeyColor(i + 1, Green)
'        j = j + 1
'    Next i
'
'    For i = 0 To 68 Step 2
'        Data((PACKET3 + 4 + j)) = KeyColor(i, Blue) * 16 + KeyColor(i + 1, Blue)
'        j = j + 1
'    Next i: j = 0
'
'    For i = 72 To 142 Step 2
'        Data((PACKET4 + 4 + j)) = KeyColor(i, Blue) * 16 + KeyColor(i + 1, Blue)
'        j = j + 1
'    Next i
    
    
    
    
    
    
'    If chkCombine.value = vbChecked Then
        'byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H200, &H3, Data(PACKET1), 832, 1)
'    Else
'        byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H200, &H3, Data(PACKET1), 64, 500)
'        byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H200, &H3, Data(PACKET2), 64, 500)
'        byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H200, &H3, Data(PACKET3), 64, 500)
'        byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H200, &H3, Data(PACKET4), 64, 500)
'        byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H200, &H3, Data(PACKET5), 64, 500)
'        byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H300, &H3, Data(PACKET6), 64, 500)
'        byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H300, &H3, Data(PACKET7), 64, 500)
'        byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H300, &H3, Data(PACKET8), 64, 500)
'        byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H300, &H3, Data(PACKET9), 64, 500)
'        byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H300, &H3, Data(PACKET10), 64, 500)
'        byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H300, &H3, Data(PACKET11), 64, 500)
'        byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H300, &H3, Data(PACKET12), 64, 500)
'    End If
    
    'For i = 0 To 250
        'byteswritten = UsbBulkWrite(keyboard, 0, Data(PACKET1), 832, 1500)
        'Debug.Print i; byteswritten
    'Next i
    
    'byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET1), 768, 1500)
    
'    byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET1), 64, 1500)
'    byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET2), 64, 1500)
'    byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET3), 64, 1500)
'    byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET4), 64, 1500)
    

    
    'Wait 20
    
'    byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET5), 64, 1500)
'    byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET6), 64, 1500)
'    byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET7), 64, 1500)
'    byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET8), 64, 1500)
    
    
    
    'Wait 20
    
'    byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET9), 64, 1500)
'    byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET10), 64, 1500)
'    byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET11), 64, 1500)
'    byteswritten = UsbBulkWrite(keyboard, 4, Data(PACKET12), 64, 1500)
'
    
    'Wait 20
    
EndOfTiming:
    
    'byteswritten = UsbControlMsg(keyboard, &H21, &H9, &H300, &H3, Data(PACKET1), 64 * 13, 500)
    
    
    
    
    'MsgBox ElapsedTime(t)
End Sub

Public Sub Wait(ByVal dblMilliseconds As Double)
    Dim dblStart As Double
    Dim dblEnd As Double
    Dim dblTickCount As Double
    
    dblTickCount = GetTickCount()
    dblStart = GetTickCount()
    dblEnd = GetTickCount + dblMilliseconds
    
    Do
    DoEvents
    dblTickCount = GetTickCount()
    Loop Until dblTickCount > dblEnd Or dblTickCount < dblStart
       
    
End Sub

Private Sub scrSolid_Change(Index As Integer)
    Dim i As Long
    
    
    lblColor(Index).Caption = scrSolid(Index).value
    
    picSolid.BackColor = RGB(scrSolid(0).value, scrSolid(1).value, scrSolid(2).value)
    
    For i = 0 To 144
        KeyColor(i, Red) = scrSolid(0).value
        KeyColor(i, Green) = scrSolid(1).value
        KeyColor(i, Blue) = scrSolid(2).value
    Next i
    
    
End Sub

Private Sub scrSolid_Scroll(Index As Integer)
    scrSolid_Change Index
End Sub

Private Sub GetBusInfo()
    Dim buffer(0 To 255) As Byte
    Dim i, X As Long
    Dim dev As Long
    Dim usbver As String
    
    frmInfo.txtInfo.Text = ""
    i = 0
    dev = keyboard

    If UsbGetDeviceDescriptor(dev, my_descriptor) Then
        Msg "Device " & i
        Msg "-- VID             : " & Hex$(my_descriptor.idVendor)
        Msg "-- PID             : " & Hex$(my_descriptor.idProduct) & " "

        If (my_descriptor.iManufacturer) Then
            If (UsbGetStringSimple(dev, my_descriptor.iManufacturer, buffer(0), UBound(buffer))) Then
                Msg "-- Manufacturer    : " & StrConv(buffer, vbUnicode)
                Msg " "
            End If
        Else
            Msg "-- Manufacturer    : not specified"
        End If

        If (my_descriptor.iProduct) Then
            If (UsbGetStringSimple(dev, my_descriptor.iProduct, buffer(0), UBound(buffer))) Then
                Msg "-- Product         : " & StrConv(buffer, vbUnicode)
                Msg " "
            End If
        Else
                Msg "-- Product         : not specified"
        End If

        If (my_descriptor.iSerialNumber) Then
            If (UsbGetStringSimple(dev, my_descriptor.iSerialNumber, buffer(0), UBound(buffer))) Then
                Msg "- Serial  nmbr    : " & StrConv(buffer, vbUnicode)
                Msg " "
            End If
        Else
            Msg "-- Serial  nmbr    : not specified"
        End If

        ' retrieve the USB version
        usbver = Hex$(my_descriptor.bcdUSB)
        Mid$(usbver, 3, 1) = Mid$(usbver, 2, 1)
        Mid$(usbver, 2, 1) = "."
        Msg "-- USB version     : " & usbver

        Msg "-- Device Class    : " & my_descriptor.bDeviceClass
        Msg "-- Subclass        : " & my_descriptor.bDeviceSubClass
'       msg "-- Max Packet size : " & my_descriptor.bMaxPacketSize0
'       msg "-- Protocol        : " & my_descriptor.bDeviceProtocol
        For X = 0 To my_descriptor.bNumConfigurations - 1
            print_configuration dev, X
        Next X
    End If

End Sub

Sub print_configuration(handle As Long, Index)
    Dim X
    If UsbGetConfigurationDescriptor(handle, Index, dev_config) Then
        Msg "--- Configuration   : " & Index
        Msg "--- Total Length    : " & dev_config.wTotalLength
        Msg "--- NUM interfaces  : " & dev_config.bNumInterfaces
        Msg "--- Config. Value   : " & dev_config.bConfigurationValue
        Msg "--- Configuration   : " & dev_config.iConfiguration
        Msg "--- Attributes      : " & Hex$(dev_config.bmAttributes)
        Msg "--- Max Power       : " & dev_config.MaxPower
    End If
    For X = 0 To dev_config.bNumInterfaces - 1
        print_interface handle, Index, X
    Next X
End Sub

Sub print_interface(handle As Long, config_index, interface_index)
    Dim X
    Dim alt

    alt = 0

    Do While UsbGetInterfaceDescriptor(handle, config_index, interface_index, alt, my_interface)
       Msg "---- Interface         : " & interface_index & "/" & alt
       Msg "---- Alternate setting : " & my_interface.bAlternateSetting
       Msg "---- NUMEndpoints      : " & my_interface.bNumEndpoints
       Msg "---- InterfaceClass    : " & my_interface.bInterfaceClass
       Msg "---- InterfaceSubClass : " & my_interface.bInterfaceSubClass
       Msg "---- InterfaceProtocol : " & my_interface.bInterfaceProtocol
       Msg "---- Interface         : " & my_interface.iInterface
       Msg "---- DescriptorType    : " & my_interface.bDescriptorType
       
       For X = 0 To my_interface.bNumEndpoints - 1
           print_endpoint handle, config_index, interface_index, alt, X
       Next X

       alt = alt + 1
    Loop
End Sub


Sub print_endpoint(handle As Long, config_index, interface_index, alt_setting, Index)
    If UsbGetEndpointDescriptor(handle, config_index, interface_index, alt_setting, Index, my_endpoint) Then
       Msg "----- Endpoint      : " & Index
       Msg "----- Address       : " & Hex$(my_endpoint.bEndpointAddress)
       Msg "----- Attributes    : " & Hex$(my_endpoint.bmAttributes)
       Msg "----- MaxPacketSize : " & Hex$(my_endpoint.wMaxPacketSize)
       Msg "----- Interval      : " & Hex$(my_endpoint.bInterval)
       Msg "----- Refresh       : " & Hex$(my_endpoint.bRefresh)
       Msg "----- Syncaddress   : " & Hex$(my_endpoint.bSynchAddress)
    End If
End Sub

Sub Msg(Msg)
   frmInfo.txtInfo.Text = frmInfo.txtInfo.Text & Msg & vbCrLf
End Sub


Sub Play_Flappy_Bird()
    Static Tubes(0 To 15) As Tube
    
    Static Game_X As Long
    'Static Game_Y As Long
    Static FrontTube As Long
    Const Tube_Spacing As Long = 10
    Const Tube_Max_Height As Long = 2
    Const Tube_Speed_Delay As Long = 6
    
    Static Fish_Y As Long
    Static Fish_Dead As Long
    Static Fish_Score As Long
    
    
    Dim TubesVisible As Long
    Dim i As Long
    
    
    Static Ticks As Long
    
    Ticks = Ticks + 1
    
    ClearLeds
    Static usedArrow As Boolean
    
    
    If KeyVKDown(Keys.KEY_ARROW_UP) Then
        If usedArrow = False Then Fish_Y = Fish_Y - 1
        
        usedArrow = True
    Else
        usedArrow = False
        If Ticks >= Tube_Speed_Delay Then
            Fish_Y = Fish_Y + 1
        End If
    End If
    
    If Fish_Y > 6 Then Fish_Y = 6
    If Fish_Y < 1 Then Fish_Y = 1
    
    For i = 0 To UBound(Tubes)
        If Tubes(i).Visible Then
            TubesVisible = TubesVisible + 1
            If Tubes(i).X <= 0 Then
                If Ticks >= Tube_Speed_Delay Then
                    Tubes(i).Visible = False
                End If
                'FrontTube = FrontTube + 1
                'If FrontTube > UBound(Tubes) Then FrontTube = 0
                
            Else
                If Fish_Dead = 0 And Ticks >= Tube_Speed_Delay Then
                    Tubes(i).X = Tubes(i).X - 1
                End If
                
                If Tubes(i).X < 24 And Tubes(i).X >= 0 Then
                    'Tube_Max_Height
                    
                    SetLed Tubes(i).X, Tubes(i).y - 1, 0, 255, 0
                    SetLed Tubes(i).X, Tubes(i).y - 2, 0, 255, 0
                    SetLed Tubes(i).X, Tubes(i).y, 0, 255, 0
                    
                    SetLed Tubes(i).X, Tubes(i).y + 3, 0, 255, 0
                    SetLed Tubes(i).X, Tubes(i).y + 4, 0, 255, 0
                    SetLed Tubes(i).X, Tubes(i).y + 5, 0, 255, 0
                    
                    If Tubes(i).X = 2 And Fish_Dead = 0 Then
                        If Fish_Y <= Tubes(i).y Or Fish_Y >= Tubes(i).y + 3 Then
                            Fish_Dead = 1
                        Else
                            If Ticks >= Tube_Speed_Delay Then
                                Fish_Score = Fish_Score + 1
                            End If
                        End If
                    End If
                    
                End If
            End If
        Else
            If i > 0 Then
                If Tubes(i - 1).X < 24 And Tubes(i - 1).Visible Then
                    Tubes(i).X = 24 + Tube_Spacing
                    Tubes(i).y = (Rnd * Tube_Max_Height) + 1
                    Tubes(i).Visible = True
                End If
            Else
                If Tubes(UBound(Tubes)).X < 24 And Tubes(UBound(Tubes)).Visible Then
                    Tubes(i).X = 24 + Tube_Spacing
                    Tubes(i).y = (Rnd * Tube_Max_Height) + 1
                    Tubes(i).Visible = True
                End If
            End If
            
        End If
        
    Next i
    
    If Fish_Dead > 0 Then
        If Ticks >= Tube_Speed_Delay Then
            Fish_Dead = Fish_Dead + 1
        End If
        
        Select Case Fish_Dead
            Case 1
                SetLed 2, Fish_Y, 255, 0, 0
            Case 2
                
                SetLed 2, Fish_Y - 1, 255, 0, 0
            Case 3
                If Fish_Y = 4 Then
                    SetLed 0, Fish_Y, 255, 0, 0
                Else
                    SetLed 1, Fish_Y, 255, 0, 0
                End If
            Case 4, 5, 6, 7, 8, 9, 10, 11
                If (Fish_Y + (Fish_Dead - 3)) = 5 Then
                
                    SetLed 1, 5, 255, 0, 0
                Else
                    SetLed 0, (Fish_Y + (Fish_Dead - 3)), 255, 0, 0
                End If
            Case 12
                For i = 0 To UBound(Tubes)
                    Tubes(i).X = 0
                    Tubes(i).y = 0
                    Tubes(i).Visible = False
                    Fish_Y = 4
                    Fish_Score = 0
                Next i
                Fish_Dead = 0
                Exit Sub
                
        End Select
        
    Else
        SetLed 2, Fish_Y, 255, 100, 0
    End If
    
    
    If TubesVisible = 0 Then
        Tubes(0).X = 24
        Tubes(0).y = (Rnd * Tube_Max_Height) + 1
        Tubes(0).Visible = True
    End If
    
    
    NUMberToNUMpad Fish_Score, 0, 0, 255
    
    
    If Ticks >= Tube_Speed_Delay Then Ticks = 0
End Sub



    
Private Sub tmrAnimation_Timer()

    
    Dim X As Long
    Dim y As Long
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    Dim i As Long
    Dim P As Long
    
    Dim lAnimationLength As Long
    
    lAnimationLength = scrWidth.value * 6
    
    Static a As Long
    
    Select Case lstAnimations.ListIndex
        Case 0
            scrSolid_Change 0
            tmrAnimation.Interval = 10
            
        Case 1, 2
            a = a + 1
            If a >= lAnimationLength Then a = 0
            
'            For i = 0 To 144
'                KeyColor(i, Red) = scrSolid(0).value
'                KeyColor(i, Green) = scrSolid(1).value
'                KeyColor(i, Blue) = scrSolid(2).value
'            Next i
            
            Static_Wave_Rainbow a, lAnimationLength, False
            
            'ClearLeds
            
            'If a >= 1500 Then a = 0
            
            'getRainbowColor R, G, B, a, 1500
            
            'SolidLights R, G, B
            
            Static_Wave_Rainbow a, lAnimationLength, lstAnimations.ListIndex = 2

        Case 3
            'picAni.Picture = LoadPicture()
            Random_Circle_Move chkAnimation(0).value = vbChecked
        
        Case 4
            Random_Lines_Horizontal chkAnimation(0).value = vbChecked
            
        Case 5
            Random_Lines_Vertical chkAnimation(0).value = vbChecked

        Case 6
            Random_Keyboard_Color chkAnimation(0).value = vbChecked

        Case 7
            Random_Key_Color chkAnimation(0).value = vbChecked

        Case 8
            'picAni.Picture = LoadPicture()
            Random_Arrow_Move chkAnimation(0).value = vbChecked

        Case 9
            Random_Spiral_Fill chkAnimation(0).value = vbChecked
            
        Case 10
            Play_Flappy_Bird

        Case 11 'solid color cycle
            Random_Color_Cycle

        Case 12 'game of life maybe?
            Static_Game_Of_Life

        Case 13 'starfield
            Static_Star_Field

        Case 14 'fireworks
            Static_Fireworks
        
        Case 15 'RainDrops
            Random_RainDrops
            
        Case 16 'Storm/Rain/Thunder
            Random_Weather_Storm
            
        Case 17 'rainbow spiral fill
            Static_Spiral_Fill_Rainbow
            
        Case 18 'smooth rainbow
            a = a + 1
            If a >= 24 * 8 Then a = 0
            
            For X = 0 To 24
                getRainbowColor R, G, B, a + X * 8, 24 * 8
                For y = 0 To 6
                    SetLed X, y, R, G, B
                Next y
            Next X
            
            
            
            
    End Select
    
    
    
    'a = a + 1
    'If a = picAni.ScaleWidth * 1 Then a = 0
    'Static_Wave_Rainbow picAni, a, picAni.ScaleWidth * 1
    'tmrAnimation.Interval = 50
    
    'Random_Arrow_Move chkAnimation(0).value = vbChecked
    

    
    
    'KeyboardToPicture picAni, picResize
    'Paint_Keyboard
    
    'imgAni.Picture = picAni.Image
    
    'SendData
End Sub

Sub Random_RainDrops()
    
    Dim i As Long
    Static FadeState(0 To 144) As Byte
    Static FadeColor(0 To 144, 0 To 2) As Byte
    Static FadeCurrentColor(0 To 144, 0 To 2) As Single
    
    
    For i = 0 To 144
        If FadeState(i) = 1 Then 'fade in state
            If KeyColor(i, Red) >= FadeColor(i, Red) And KeyColor(i, Green) >= FadeColor(i, Green) And KeyColor(i, Blue) >= FadeColor(i, Blue) Then
                FadeState(i) = 2
            Else
                If KeyColor(i, Red) < FadeColor(i, Red) Then FadeCurrentColor(i, Red) = FadeCurrentColor(i, Red) + ((FadeColor(i, Red)) / 25)
                If KeyColor(i, Green) < FadeColor(i, Green) Then FadeCurrentColor(i, Green) = FadeCurrentColor(i, Green) + ((FadeColor(i, Green)) / 25)
                If KeyColor(i, Blue) < FadeColor(i, Blue) Then FadeCurrentColor(i, Blue) = FadeCurrentColor(i, Blue) + ((FadeColor(i, Blue)) / 25)
                
                If FadeCurrentColor(i, Red) >= 255 Then FadeCurrentColor(i, Red) = 255
                If FadeCurrentColor(i, Green) >= 255 Then FadeCurrentColor(i, Green) = 255
                If FadeCurrentColor(i, Blue) >= 255 Then FadeCurrentColor(i, Blue) = 255
                
                KeyColor(i, Red) = FadeCurrentColor(i, Red)
                KeyColor(i, Green) = FadeCurrentColor(i, Green)
                KeyColor(i, Blue) = FadeCurrentColor(i, Blue)
                
            End If
        ElseIf FadeState(i) = 2 Then 'fade out state
            If KeyColor(i, Red) = 0 And KeyColor(i, Green) = 0 And KeyColor(i, Blue) = 0 Then
                FadeState(i) = 0
            Else
                If KeyColor(i, Red) > 0 Then FadeCurrentColor(i, Red) = FadeCurrentColor(i, Red) - 4
                If KeyColor(i, Green) > 0 Then FadeCurrentColor(i, Green) = FadeCurrentColor(i, Green) - 4
                If KeyColor(i, Blue) > 0 Then FadeCurrentColor(i, Blue) = FadeCurrentColor(i, Blue) - 4
                
                If FadeCurrentColor(i, Red) <= 0 Then FadeCurrentColor(i, Red) = 0
                If FadeCurrentColor(i, Green) <= 0 Then FadeCurrentColor(i, Green) = 0
                If FadeCurrentColor(i, Blue) <= 0 Then FadeCurrentColor(i, Blue) = 0
                
                KeyColor(i, Red) = FadeCurrentColor(i, Red)
                KeyColor(i, Green) = FadeCurrentColor(i, Green)
                KeyColor(i, Blue) = FadeCurrentColor(i, Blue)
            End If
        End If
        
    Next i
    
    Dim newKey As Byte
    Dim tries As Long
    
    For i = 0 To 10
        tries = 0
try_key:
        newKey = Rnd * 144
        tries = tries + 1
        
        If tries > 10 Then
            GoTo next_i
        End If
        
        If FadeState(newKey) = 0 Then
            KeyColor(newKey, Red) = 0
            KeyColor(newKey, Green) = 0
            KeyColor(newKey, Blue) = 0
            FadeCurrentColor(newKey, Red) = 0
            FadeCurrentColor(newKey, Green) = 0
            FadeCurrentColor(newKey, Blue) = 0
            
            FadeColor(newKey, Red) = 0
            FadeColor(newKey, Green) = 255
            FadeColor(newKey, Blue) = 255 * Rnd
            
            'GetRandomKeyColor FadeColor(newKey, Red), FadeColor(newKey, Green), FadeColor(newKey, Blue), chkAnimation(0).value = vbChecked
            
            FadeState(newKey) = 1
        End If
        
next_i:
    Next i
    
    
    
'
'    For i = 0 To 9
'        tries = 0
'try_key:
'        newKey = Rnd * 144
'        tries = tries + 1
'        If tries > 10 Then
'            GoTo next_i
'        End If
'        If KeyColor(newKey, Red) < 7 Or KeyColor(newKey, Green) < 7 Or KeyColor(newKey, Blue) < 7 Then GoTo try_key
'
'        KeyColor(newKey, Red) = 7
'        KeyColor(newKey, Green) = 0
'        KeyColor(newKey, Blue) = 1
'
'next_i:
'    Next i
    
End Sub

Sub FillRow(lRow As Long, R As Byte, G As Byte, B As Byte)
    Dim i As Long
    
    For i = 0 To 26
        SetLed i, lRow, R, G, B
    Next i
End Sub

Sub FillCol(lCol As Long, R As Byte, G As Byte, B As Byte)
    Dim i As Long
    
    For i = 0 To 26
        SetLed lCol, i, R, G, B
    Next i
End Sub


Sub Static_Game_Of_Life()
    Static Game_World(0 To 23, 0 To 6) As Byte
    
    
    
    
    
    
End Sub




Sub Random_Color_Cycle()
    Static R1 As Byte
    Static B1 As Byte
    Static G1 As Byte
    
    Static R2 As Byte
    Static B2 As Byte
    Static G2 As Byte

    Static CurrentStep As Long
    Static Initialized As Boolean
    
    Dim R As Long
    Dim G As Long
    Dim B As Long
    
    If Initialized = False Then
        GetRandomColor R1, G1, B1, True
        GetRandomColor R2, G2, B2, True
        Initialized = True
        
        CurrentStep = 1
    End If
    
    If CurrentStep = -1 Then
        R1 = R2
        G1 = G2
        B1 = B2
        GetRandomColor R2, G2, B2, True
        CurrentStep = 1
    End If
    
    If CurrentStep >= 0 And CurrentStep <= 30 Then
        R = CInt(R1) - CInt(R2)
        G = CInt(G1) - CInt(G2)
        B = CInt(B1) - CInt(B2)
        If R < 0 Then
            R = R1 + (-R / 30 * CurrentStep)
        ElseIf R > 0 Then
            R = R1 - (R / 30 * CurrentStep)
        ElseIf R = 0 Then
            R = R1
        End If
        
        If G < 0 Then
            G = G1 + (-G / 30 * CurrentStep)
        ElseIf G > 0 Then
            G = G1 - (G / 30 * CurrentStep)
        ElseIf G = 0 Then
            G = G1
        End If
        
        If B < 0 Then
            B = B1 + (-B / 30 * CurrentStep)
        ElseIf B > 0 Then
            B = B1 - (B / 30 * CurrentStep)
        ElseIf B = 0 Then
            B = B1
        End If
        
        SolidLights CByte(R), CByte(G), CByte(B)
    Else
        SolidLights CByte(R2), CByte(G2), CByte(B2)
    End If
    
    CurrentStep = CurrentStep + 1
    
    If CurrentStep > 45 Then CurrentStep = -1
    
End Sub


Sub Static_Fireworks()
    Const sRocket_Path As String = "23,2|11,2|9,3|5,3|0,0"
    
    Static Rocket_Path() As POINT
    Static Rocket_Trail(0 To 7) As POINT
    
    Static X As Long
    Static y As Long
    Static R As Byte
    Static B As Byte
    Static G As Byte
    
    Dim tR As Byte
    Dim tB As Byte
    Dim tG As Byte
    
    Static eR As Byte
    Static eG As Byte
    Static eB As Byte
    
    Static CurrentStep As Long
    
    Dim tmpSplit1() As String
    Dim tmpSplit2() As String
    Dim i As Long
    
    Static sFire_Size As Long
    
    Static Initialized As Boolean
    
    ClearLeds
    
    If Initialized = False Then
        Initialized = True
        tmpSplit1 = Split(sRocket_Path, "|")
        ReDim Rocket_Path(0 To UBound(tmpSplit1))
        For i = 0 To UBound(tmpSplit1)
            tmpSplit2 = Split(tmpSplit1(i), ",")
            Rocket_Path(i).X = CInt(tmpSplit2(0))
            Rocket_Path(i).y = CInt(tmpSplit2(1))
        Next i
        CurrentStep = 0
    End If
    
    
    If CurrentStep = 0 Then
        'GetRandomColor R, G, B, True
        
        sFire_Size = (7 * Rnd) + 3
        
        X = Rocket_Path(0).X
        y = Rocket_Path(0).y
        
        GetRandomColor R, G, B, True
        GetRandomColor eR, eG, eB, True
        
        For i = 0 To UBound(Rocket_Trail)
            Rocket_Trail(i).X = 0
            Rocket_Trail(i).y = 0
        Next i
        
        
        CurrentStep = CurrentStep + 1
    ElseIf CurrentStep < 0 Then
        If CurrentStep <= -sFire_Size Then
            For i = -sFire_Size To CurrentStep + sFire_Size
                Make_Circle X, y, sFire_Size + i, eR / sFire_Size * -i, eG / sFire_Size * -i, eB / sFire_Size * -i
                'Make_Circle X, Y, 10 + CurrentStep + 1, eB, eG, eR
            Next i
        Else
            For i = -sFire_Size To 0
                Make_Circle X, y, sFire_Size + CSng(i), (eR / sFire_Size * -i) / sFire_Size * -CurrentStep, (eG / sFire_Size * -i) / sFire_Size * -CurrentStep, (eB / sFire_Size * -i) / sFire_Size * -CurrentStep
                'Make_Circle X, Y, 10 + CurrentStep + 1, eB, eG, eR
            Next i
        End If
        
        CurrentStep = CurrentStep + 1
    Else
        
        
        For i = UBound(Rocket_Trail) To 1 Step -1
            Rocket_Trail(i) = Rocket_Trail(i - 1)
        Next i
        

        If Rocket_Path(CurrentStep).X < X Then
            X = X - 1
        ElseIf Rocket_Path(CurrentStep).X > X Then
            X = X + 1
        Else 'already on the X
            If Rocket_Path(CurrentStep).y < y Then
                y = y - 1
            ElseIf Rocket_Path(CurrentStep).y > y Then
                y = y + 1
            Else 'already on the Y
                If Rocket_Path(CurrentStep + 1).X = 0 And Rocket_Path(CurrentStep + 1).y = 0 Then
                    CurrentStep = -(sFire_Size * 2)
                Else
                    X = Rocket_Path(CurrentStep + 1).X
                    y = Rocket_Path(CurrentStep + 1).y
                    CurrentStep = CurrentStep + 2
                End If
            End If
        End If
        
        Rocket_Trail(0).X = X
        Rocket_Trail(0).y = y
        
        tR = 255 / 10 * sFire_Size
        tG = 255 / 10 * sFire_Size
        tB = 255 / 10 * sFire_Size
        'R = 255
        'G = 255
        'B = 255
        
        
        SetLed X, y, R / 10 * sFire_Size, G / 10 * sFire_Size, B / 10 * sFire_Size
        
        For i = 1 To UBound(Rocket_Trail)
            LowerColorBrightness tR, tG, tB
            SetLed Rocket_Trail(i).X, Rocket_Trail(i).y, tR, tG, tB
        Next i
        
        
        
    End If
    
    
    
End Sub

Sub LowerColorBrightness(ByRef R As Byte, ByRef G As Byte, ByRef B As Byte)
    Dim ir As Long
    Dim ib As Long
    Dim ig As Long
    
    ir = R - (255 / 7)
    ig = G - (255 / 7)
    ib = B - (255 / 7)
    
    If ir < 0 Then ir = 0
    If ig < 0 Then ig = 0
    If ib < 0 Then ib = 0
    
    R = ir
    G = ig
    B = ib
End Sub

Sub HigherColorBrightness(ByRef R As Byte, ByRef G As Byte, ByRef B As Byte)
    Dim ir As Long
    Dim ib As Long
    Dim ig As Long
    
    
    ir = R + (255 / 7)
    ig = G + (255 / 7)
    ib = B + (255 / 7)
    
    If ir > 255 Then ir = 255
    If ig > 255 Then ig = 255
    If ib > 255 Then ib = 255
    
    R = ir
    G = ig
    B = ib
End Sub


Sub Static_Star_Field()
    Static CurrentKeys(0 To 10) As Stars
    
    
    Static X As Long
    Static y As Long
    Static Initialized As Boolean
    Dim i As Long
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    
    R = 255
    G = 255
    B = 255
    
    Dim lCenterX As Long
    Dim lCenterY As Long
    
    lCenterX = 8
    lCenterY = 3
    
    SolidLights 7, 7, 7
    
    If Initialized = False Then
        Initialized = True
        
        For i = 0 To UBound(CurrentKeys)
            CurrentKeys(i).X = 0 '(Rnd - 0.48) * 1
            CurrentKeys(i).y = 0 '(Rnd - 0.5) * 1
            CurrentKeys(i).Angle = Rnd * 2 * 3.1415926
            CurrentKeys(i).nStep = Rnd * 20
            GetRandomColor CurrentKeys(i).R, CurrentKeys(i).G, CurrentKeys(i).B, True
        Next i
        
    End If
    
    For i = 0 To UBound(CurrentKeys)
        With CurrentKeys(i)
            
            If .nStep >= 20 Then
                .X = 0 '(Rnd - 0.5) * 1
                .y = 0 '(Rnd - 0.5) * 1
                .Angle = Rnd * 2 * 3.1415926
                .nStep = -1 'Rnd * 5
                
                GetRandomColor .R, .G, .B, True
            End If
            
            .nStep = .nStep + 1
            
            X = Sin(.Angle) * .nStep + (.X + lCenterX)
            y = Cos(.Angle) * .nStep + (.y + lCenterY)
            
            SetLed X, y, .R, .G, .B
        End With
    Next i
    
    'For X = 0 To 23
        'For Y = 0 To 6
            'SetLed X, Y, R, G, B
    
    'KeyMatrix ' = CurrentKeys
End Sub



Private Sub tmrTimers_Timer()
    txtLoopSpeed.Text = Round(TimeQuerys(Timers.t_ProgramLoop), 6)
    txtSendSpeed.Text = Round(TimeQuerys(Timers.t_SendData), 6)
End Sub

Private Sub tmrVolume_Timer()
    Static a(0 To 8) As Long
    
    Dim i As Long
    
    For i = 0 To 8
        If m_bVolumeInitialized(i) Then
            m_LonVolume(i) = GetVolume(SPEAKER, i)
            m_bVolumeMute(i) = GetMute(SPEAKER, i)
            If m_LonVolumePrevious(i) <> m_LonVolume(i) Then 'show volume effect
                m_LonVolumePrevious(i) = m_LonVolume(i)
                a(i) = 1
                m_bVolumeEffect(i) = True
            End If
            
            If a(i) > 0 Then
                a(i) = a(i) + 1
                
                If a(i) >= 200 Then
                    a(i) = 0
                    m_bVolumeEffect(i) = False
                End If
                
            End If
        End If
    Next i
End Sub

Sub GetRandomColor(ByRef R As Byte, ByRef G As Byte, ByRef B As Byte, Optional MaxColor As Boolean = False)
    Dim tmpR As Byte
    Dim tmpG As Byte
    Dim tmpB As Byte
    
    'Randomize
    
GenerateColor:
    If MaxColor Then
        
        Select Case CInt(Rnd * 2)
            Case 0
                tmpR = Rnd * 255
                tmpG = Rnd * 255
                tmpB = 255
                
            Case 1
                tmpR = 255
                tmpG = Rnd * 255
                tmpB = Rnd * 255
            Case 2
                tmpR = Rnd * 255
                tmpG = 255
                tmpB = Rnd * 255
        End Select
        
    Else
        tmpR = Rnd * 230 + 25
        tmpG = Rnd * 230 + 25
        tmpB = Rnd * 230 + 25
    End If
    
    If tmpR = R And tmpG = G And tmpB = B Then GoTo GenerateColor
    R = tmpR
    G = tmpG
    B = tmpB
End Sub


Sub GetRandomKeyColor(ByRef R As Byte, ByRef G As Byte, ByRef B As Byte, Optional MaxColor As Boolean = False)
    Dim tmpR As Byte
    Dim tmpG As Byte
    Dim tmpB As Byte
    
GenerateColor:
    GetRandomColor tmpR, tmpG, tmpB, MaxColor
    
    tmpR = 7 - (tmpR / 7)
    tmpG = 7 - (tmpG / 7)
    tmpB = 7 - (tmpB / 7)
    
    If tmpR = R And tmpG = G And tmpB = B Then GoTo GenerateColor
    R = tmpR
    G = tmpG
    B = tmpB
    
End Sub




