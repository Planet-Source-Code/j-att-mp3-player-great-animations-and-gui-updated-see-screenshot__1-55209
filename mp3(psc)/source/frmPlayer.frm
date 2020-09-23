VERSION 5.00
Object = "{4FC6B314-E09D-4BBE-9F62-0D60FDC597E4}#1.0#0"; "FMOD.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlayer 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Media Player"
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   Icon            =   "frmPlayer.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmPlayer.frx":08CA
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   499
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraOpen 
      Appearance      =   0  'Flat
      BackColor       =   &H00D7D7D7&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2820
      Left            =   5070
      TabIndex        =   18
      Top             =   225
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox picOpen 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   2850
      Left            =   5070
      ScaleHeight     =   2850
      ScaleWidth      =   1425
      TabIndex        =   17
      Top             =   210
      Width           =   1425
   End
   Begin VB.PictureBox picPlaylist 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2820
      Left            =   5070
      Picture         =   "frmPlayer.frx":7254
      ScaleHeight     =   188
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   162
      TabIndex        =   31
      Top             =   210
      Visible         =   0   'False
      Width           =   2430
      Begin VB.ListBox lstPlaylist 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   2010
         ItemData        =   "frmPlayer.frx":9135
         Left            =   60
         List            =   "frmPlayer.frx":9137
         TabIndex        =   32
         Top             =   330
         Width           =   2310
      End
      Begin VB.Label lblPlaylist 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " X "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   2040
         TabIndex        =   39
         ToolTipText     =   "Exit"
         Top             =   60
         Width           =   195
      End
      Begin VB.Label lblPlaylist 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "REMOVE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   1620
         TabIndex        =   36
         Top             =   2460
         Width           =   675
      End
      Begin VB.Label lblPlaylist 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " ADD "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1095
         TabIndex        =   35
         Top             =   2460
         Width           =   450
      End
      Begin VB.Label lblPlaylist 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   34
         Top             =   2460
         Width           =   420
      End
      Begin VB.Label lblPlaylist 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "OPEN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   33
         Top             =   2460
         Width           =   420
      End
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   3360
      Top             =   3360
   End
   Begin VB.Timer tmrOpen2 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   2880
      Top             =   3360
   End
   Begin VB.Timer tmrOpen1 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   2400
      Top             =   3360
   End
   Begin VB.PictureBox picPeak 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   270
      Picture         =   "frmPlayer.frx":9139
      ScaleHeight     =   1665
      ScaleWidth      =   225
      TabIndex        =   14
      Top             =   300
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picVolume 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   4005
      Picture         =   "frmPlayer.frx":99EB
      ScaleHeight     =   1665
      ScaleWidth      =   225
      TabIndex        =   13
      Top             =   300
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   390
      Picture         =   "frmPlayer.frx":A2B0
      ScaleHeight     =   225
      ScaleWidth      =   15
      TabIndex        =   12
      Top             =   1845
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Timer tmrRIntro 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1920
      Top             =   3360
   End
   Begin VB.Timer tmrFIntro 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1440
      Top             =   3360
   End
   Begin VB.Timer tmrVolume 
      Interval        =   300
      Left            =   960
      Top             =   3360
   End
   Begin VB.Timer tmrFmod 
      Interval        =   35
      Left            =   480
      Top             =   3360
   End
   Begin Project1.ocxFMOD ocxFMOD 
      Height          =   615
      Left            =   510
      TabIndex        =   11
      Top             =   2295
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
   End
   Begin VB.PictureBox picSpec 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      Picture         =   "frmPlayer.frx":AE2D
      ScaleHeight     =   615
      ScaleWidth      =   255
      TabIndex        =   10
      Top             =   3360
      Width           =   255
   End
   Begin MP3.Trans Trans 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1085
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   1560
      Picture         =   "frmPlayer.frx":B4DA
      ScaleHeight     =   1740
      ScaleWidth      =   1635
      TabIndex        =   16
      Top             =   315
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.PictureBox picTime 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DS-Digital"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      Picture         =   "frmPlayer.frx":149BC
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   22
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DS-Digital"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   570
      Picture         =   "frmPlayer.frx":14ACE
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   23
      Top             =   360
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox blnCD 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6210
      TabIndex        =   29
      Top             =   1245
      Width           =   210
   End
   Begin VB.CheckBox chkLoop 
      Height          =   195
      Left            =   6210
      TabIndex        =   28
      Top             =   1020
      Width           =   210
   End
   Begin VB.CheckBox chkVMeter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Check1"
      Height          =   195
      Left            =   6210
      TabIndex        =   25
      Top             =   795
      Width           =   210
   End
   Begin VB.TextBox txtTitleSpeed 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5925
      TabIndex        =   24
      Text            =   "2"
      Top             =   540
      Width           =   495
   End
   Begin MSComctlLib.Slider sldBalance 
      Height          =   120
      Left            =   5190
      TabIndex        =   30
      ToolTipText     =   "Audio Balance"
      Top             =   1560
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   212
      _Version        =   393216
      Max             =   255
      SelStart        =   127
      TickStyle       =   3
      TickFrequency   =   20
      Value           =   127
   End
   Begin VB.CheckBox chkPL 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   255
      Left            =   6210
      TabIndex        =   40
      Top             =   1740
      Width           =   210
   End
   Begin VB.Line lnPL 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      Visible         =   0   'False
      X1              =   344
      X2              =   344
      Y1              =   116
      Y2              =   132
   End
   Begin VB.Line lnPL 
      BorderColor     =   &H00808080&
      Index           =   2
      Visible         =   0   'False
      X1              =   344
      X2              =   410
      Y1              =   132
      Y2              =   132
   End
   Begin VB.Line lnPL 
      BorderColor     =   &H00808080&
      Index           =   1
      Visible         =   0   'False
      X1              =   410
      X2              =   410
      Y1              =   116
      Y2              =   132
   End
   Begin VB.Line lnPL 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   344
      X2              =   410
      Y1              =   116
      Y2              =   116
   End
   Begin VB.Shape shpBalBorder 
      Height          =   150
      Left            =   5175
      Top             =   1545
      Width           =   1215
   End
   Begin VB.Line lnAbout 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      Visible         =   0   'False
      X1              =   361
      X2              =   361
      Y1              =   179
      Y2              =   197
   End
   Begin VB.Line lnAbout 
      BorderColor     =   &H00808080&
      Index           =   2
      Visible         =   0   'False
      X1              =   361
      X2              =   412
      Y1              =   197
      Y2              =   197
   End
   Begin VB.Line lnAbout 
      BorderColor     =   &H00808080&
      Index           =   1
      Visible         =   0   'False
      X1              =   412
      X2              =   412
      Y1              =   179
      Y2              =   197
   End
   Begin VB.Line lnAbout 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   361
      X2              =   412
      Y1              =   179
      Y2              =   179
   End
   Begin VB.Label lblAbout 
      BackStyle       =   0  'Transparent
      Caption         =   " ABOUT "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5415
      TabIndex        =   26
      Top             =   2700
      Width           =   750
   End
   Begin VB.Label lblCurPos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "DS-Digital"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   21
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDuration 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "DS-Digital"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblslash 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "DS-Digital"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   8
      Left            =   4290
      TabIndex        =   9
      ToolTipText     =   "Exit"
      Top             =   2670
      Width           =   255
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   7
      Left            =   4290
      TabIndex        =   8
      ToolTipText     =   "Minimize"
      Top             =   2385
      Width           =   255
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   6
      Left            =   4290
      TabIndex        =   7
      ToolTipText     =   "Open (for CD see options)"
      Top             =   2100
      Width           =   255
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Height          =   390
      Index           =   5
      Left            =   4605
      TabIndex        =   6
      ToolTipText     =   "Volume Down"
      Top             =   2535
      Width           =   240
   End
   Begin VB.Image imgButton 
      Height          =   255
      Index           =   8
      Left            =   4290
      Picture         =   "frmPlayer.frx":15296
      Top             =   2670
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgButton 
      Height          =   255
      Index           =   7
      Left            =   4290
      Picture         =   "frmPlayer.frx":15544
      Top             =   2385
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgButton 
      Height          =   255
      Index           =   6
      Left            =   4290
      Picture         =   "frmPlayer.frx":157E2
      Top             =   2100
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgButton 
      Height          =   390
      Index           =   5
      Left            =   4605
      Picture         =   "frmPlayer.frx":15A90
      Top             =   2535
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Height          =   390
      Index           =   4
      Left            =   4605
      TabIndex        =   5
      ToolTipText     =   "Volume Up"
      Top             =   2115
      Width           =   240
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Height          =   390
      Index           =   3
      Left            =   4290
      TabIndex        =   4
      ToolTipText     =   "Rewind (Back)"
      Top             =   1680
      Width           =   555
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Height          =   390
      Index           =   2
      Left            =   4320
      TabIndex        =   3
      ToolTipText     =   "Fast Forward (Next)"
      Top             =   1290
      Width           =   555
   End
   Begin VB.Image imgButton 
      Height          =   390
      Index           =   4
      Left            =   4605
      Picture         =   "frmPlayer.frx":15D56
      Top             =   2115
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Height          =   390
      Index           =   1
      Left            =   4290
      TabIndex        =   2
      ToolTipText     =   "Stop"
      Top             =   885
      Width           =   555
   End
   Begin VB.Image imgButton 
      Height          =   390
      Index           =   3
      Left            =   4290
      Picture         =   "frmPlayer.frx":16022
      Top             =   1695
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image imgButton 
      Height          =   390
      Index           =   2
      Left            =   4290
      Picture         =   "frmPlayer.frx":16528
      Top             =   1290
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image imgButton 
      Height          =   390
      Index           =   1
      Left            =   4290
      Picture         =   "frmPlayer.frx":16A2D
      Top             =   885
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Height          =   390
      Index           =   0
      Left            =   4290
      TabIndex        =   1
      ToolTipText     =   "Play / Pause"
      Top             =   480
      Width           =   555
   End
   Begin VB.Image imgButton 
      Height          =   390
      Index           =   0
      Left            =   4290
      Picture         =   "frmPlayer.frx":16DA0
      Top             =   480
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblButton 
      BackStyle       =   0  'Transparent
      Height          =   165
      Index           =   9
      Left            =   4290
      TabIndex        =   15
      ToolTipText     =   "Options"
      Top             =   270
      Width           =   540
   End
   Begin VB.Image imgButton 
      Height          =   165
      Index           =   9
      Left            =   4290
      Picture         =   "frmPlayer.frx":172A1
      Top             =   285
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblBAbout 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   5160
      TabIndex        =   27
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblPL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PLAYLIST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5160
      TabIndex        =   37
      ToolTipText     =   "Show Playlist"
      Top             =   1755
      Width           =   975
   End
   Begin VB.Label lblBPL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5040
      TabIndex        =   38
      Top             =   1560
      Width           =   1455
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const WM_NCLBUTTONDOWN      As Long = &HA1
Private Const HTCAPTION             As Integer = 2
Private blnPlay                     As Boolean  'if playing
Private intVolume                   As Integer  'volume value(0-255)
Private blnUpVolume                 As Boolean  'if turning volume up
Private blnDownVolume               As Boolean  'if turning volume down
Private blnFF                       As Boolean  'if fastforwarding
Private blnRR                       As Boolean  'if rewinding
Private Filename                    As String   'file for fmod to open
Private panIntro                    As polyPAN  'shutter pan
Private panOpen                     As polyPAN  'options pan
Private lngFrame                    As Long     'shutter pan frame
Private lngOpenframe                As Long     'options pan frame
Private blnShutter                  As Boolean  'if shutter is open
Private blnOpen                     As Boolean  'if options bar is open
Private intOpenWidth                As Integer  'with for frame to open to
Private tmrx                        As Integer  'title position
Private s                           As String   'tmrtime temp time variable
Private intStage                    As Integer  'value for options animation
Private intTitleSpeed               As Integer  'obvious
Private PLstring                    As String   'temp string for playlist drag
Private PLindex                     As Integer  'temp index for playlist drag
Private PLsongindex                 As Integer
Private blnPL                       As Boolean  'variable to prevent double endofstream event
Private blnstop                     As Boolean
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub Bars(ByRef blnDisplay As Boolean)

    picPeak.Visible = blnDisplay
    picVolume.Visible = blnDisplay

End Sub

Private Sub Display(ByRef blnDisplay As Boolean)

    picProgress.Visible = blnDisplay
    lblslash.Visible = blnDisplay
    lblCurPos.Visible = blnDisplay
    lblDuration.Visible = blnDisplay
    picTime.Visible = blnDisplay
    picTitle.Visible = blnDisplay

End Sub

Private Sub blnCD_Click()
On Error Resume Next
If blnCD Then
    ocxFMOD.StopStream
Else
    ocxFMOD.CDStop
End If
blnPlay = False
chkLoop.Value = 0
ocxFMOD.Filename = ""
intVolume = 255
picVolume.Height = 255
If blnCD.Value = 1 Then 'if its checked
    If Not blnShutter Then
        tmrFIntro.Enabled = True
    End If
    setCDName
    chkVMeter.Enabled = False
    sldBalance.Enabled = False
Else 'if its not
    If blnShutter Then 'if shutter is open then close & hide stuff
        Display (False)
        Bars (False)
        tmrRIntro.Enabled = True
    End If
    tmrTime.Enabled = False
    chkVMeter.Enabled = True
    sldBalance.Enabled = True
End If
End Sub

Private Sub chkPL_Click()
If Not blnShutter Then
    tmrFIntro.Enabled = True
End If
PLStart
End Sub

Private Sub chkVMeter_Click()
picPeak.Height = 111
End Sub

Private Sub Form_Load()

    With ocxFMOD
        .ShowErrMsg = True
        'Initialize FMOD (Buffer size, output frequency, max channels, flages (use 4 or 0), output driver,
        'Mixer type 4 = autodetect), device(get device number from FMOD
        .Initiate 500, 44100, 32, 4, 2, 4, 0
        .EnableSpectrum True
        .SetGraphProperties &HFF8000, &H0, picSpec.Picture
    End With 'ocxFMOD
    intVolume = 255
    LoadPAN App.Path & "\player.pan", panIntro, App.Path & "\bg.gif"
    LoadPAN App.Path & "\open.pan", panOpen
    lngFrame = 1
    lngOpenframe = 1
    intTitleSpeed = 2
    intOpenWidth = 95
    
End Sub

Private Sub Form_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)

    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub

Private Sub Form_Paint()

    Trans.remap
    If Not blnShutter Then
        DrawPANFrame 1, panIntro, Me.hdc, True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    UnloadPAN panIntro
    fSTOP
    ocxFMOD.Terminate

End Sub

Private Function FormatTime(TimeInSec As Long) As String

  Dim Sec As Single

  Dim Min As Single
    TimeInSec = Fix(TimeInSec)
    Min = IIf(TimeInSec > 60, Fix(TimeInSec / 60), 0)
    Sec = IIf(Min >= 0, TimeInSec - (60 * Min), 0)
    FormatTime = Format$(Min & ":" & Sec, "hh:mm")

End Function

Private Sub lblBPL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LinesPL (False)
End Sub

Private Sub lblButton_Click(Index As Integer)

    Select Case Index
     Case 0 'play
            If blnPlay Then
                fPAUSE
            Else 'BLNPLAY = FALSE/0
                fPLAY
                setDuration
                tmrTime.Enabled = True
            End If
            blnPlay = Not (blnPlay) '
     Case 1 'stop
        chkLoop.Value = 0
        fSTOP
        ocxFMOD.Filename = ""
        tmrTime.Enabled = False
        If blnShutter Then
            Display (False)
            Bars (False)
            tmrRIntro.Enabled = True
        End If
     Case 6 'open
        frmOpen.Show vbModal
        Filename = frmOpen.Filename
        blnPlay = False
        fSTOP
        If LenB(Filename) And (Right(Filename, 4) <> ".cda") Then
                ocxFMOD.OpenFile Filename, 16
            If Not blnShutter Then
                tmrFIntro.Enabled = True
            End If
        End If
     Case 7 'minimize
        frmPlayer.WindowState = 1
     Case 8 'exit
        fSTOP
        ocxFMOD.Terminate
        If blnShutter Then
            Display (False)
            Bars (False)
            tmrRIntro.Enabled = True
        End If
        End
     Case 9
        If picPlaylist.Visible = False Then intOpenWidth = 95
        If blnOpen Then
            tmrOpen2.Enabled = True
         Else 'BLNOPEN = FALSE/0
            tmrOpen1.Enabled = True
        End If
    End Select

End Sub

Private Sub lblButton_DblClick(Index As Integer)

If blnCD Then
    Select Case Index
        Case 2  'ff
            ocxFMOD.CDNext
            setDuration
            setCDName
        Case 3  'rr
            ocxFMOD.CDBack
            setDuration
            setCDName
    End Select
End If

End Sub

Private Sub lblButton_MouseDown(Index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    Select Case Index
     Case 2 'fast forward
        blnFF = True
     Case 3 'rewind
        blnRR = True
     Case 4 'volume up
        blnUpVolume = True
     Case 5 'volume down
        blnDownVolume = True
    End Select
    If Index <> 6 Then
        imgButton(Index).Visible = True
    End If

End Sub

Private Sub lblAbout_Click()
If blnShutter Then
    picLogo.Visible = Not (picLogo.Visible)
    Display Not (picLogo.Visible)
End If
End Sub

Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LinesAbout (True)
End Sub

Private Sub lblBAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LinesAbout (False)
End Sub

Private Sub lblButton_MouseUp(Index As Integer, _
                              Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)

    imgButton(Index).Visible = False
    blnUpVolume = False
    blnDownVolume = False
    blnFF = False
    blnRR = False

End Sub

Private Sub lblPL_Click()
intOpenWidth = 162
tmrOpen2.Enabled = True
End Sub

Private Sub lblPL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LinesPL (True)
End Sub

Private Sub lblPlaylist_Click(Index As Integer)
Dim I As Integer
Select Case Index
    Case 0
        Dim strtemp As String
        frmOpen.Show vbModal
        If LenB(frmOpen.Filename) > 0 Then
            Open frmOpen.Filename For Input As #1
                While Not EOF(1)
                    Line Input #1, strtemp
                    lstPlaylist.AddItem strtemp
                Wend
            Close #1
        End If
    Case 1
        frmOpen.Caption = "Save"
        frmOpen.Show vbModal
        If LenB(frmOpen.Filename) > 0 Then
            Open frmOpen.Filename For Output As #1
                For I = 0 To lstPlaylist.ListCount - 1
                    Print #1, lstPlaylist.List(I)
                Next I
            Close #1
        End If
        frmOpen.Caption = "Open"
    Case 2
        frmOpen.Show vbModal
        lstPlaylist.AddItem frmOpen.Filename
    Case 3
        If lstPlaylist.ListIndex <> -1 Then
            lstPlaylist.RemoveItem (lstPlaylist.ListIndex)
        End If
    Case 4
        If lstPlaylist.List(0) = "" Then
            chkPL.Enabled = False
        Else
            fSTOP
            chkPL.Enabled = True
            blnCD.Value = 0
        End If
        tmrOpen2.Enabled = True
End Select
End Sub

Private Sub lblPlaylist_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPlaylist(Index).BorderStyle = 1
End Sub

Private Sub lstPlaylist_DblClick()
PLsongindex = lstPlaylist.ListIndex
If chkPL.Value = 1 Then PLStart
End Sub

Private Sub lstPlaylist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PLstring = lstPlaylist.List(lstPlaylist.ListIndex)
    PLindex = lstPlaylist.ListIndex
End Sub

Private Sub lstPlaylist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PLindex <> lstPlaylist.ListIndex Then
    lstPlaylist.AddItem PLstring, lstPlaylist.ListIndex
    lstPlaylist.RemoveItem (PLindex + 1)
End If
End Sub

Private Sub ocxFMOD_EndOfStream(ByVal Result As Long)
    
    If chkLoop.Value Then
        fPLAY
    ElseIf blnCD Then
        blnPlay = False
    ElseIf chkPL.Value = 1 Then
        'if not stop then
        blnPL = Not blnPL
        If blnPL Then
            PLsongindex = PLsongindex + 1
            If PLsongindex = lstPlaylist.ListCount Then PLsongindex = 0
            ocxFMOD.OpenFile lstPlaylist.List(PLsongindex), 16
            frmOpen.ShortName = Right(lstPlaylist.List(PLsongindex), Len(lstPlaylist.List(PLsongindex)) - InStrRev(lstPlaylist.List(PLsongindex), "\", , vbTextCompare))
        End If
        fPLAY
        setDuration
    End If
    
End Sub

Private Sub picPlaylist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Integer
    For I = 0 To 4
        lblPlaylist(I).BorderStyle = 0
    Next I
End Sub

Private Sub sldBalance_Change()
    ocxFMOD.SetBalance sldBalance.Value
End Sub

Private Sub sldBalance_Scroll()
    sldBalance_Change
End Sub

Private Sub tmrOpen1_Timer()
        DrawPANFrame lngOpenframe, panOpen, picOpen.hdc, False
        lngOpenframe = lngOpenframe + 1
        If lngOpenframe = 15 Then
            lngOpenframe = 1
            fraOpen.Visible = True
            tmrOpen1.Enabled = False
            tmrOpen2.Enabled = True
        End If
End Sub

Private Sub tmrOpen2_Timer()
     
     If intStage = 0 Then
        fraOpen.Visible = True
        With fraOpen
            .Width = .Width + 15
            If .Width >= intOpenWidth - 15 Then
                .Width = intOpenWidth
                intStage = 1
                If intOpenWidth = 95 Then
                    picOpen.Visible = Not picOpen.Visible
                Else
                    picPlaylist.Visible = Not picPlaylist.Visible
                End If
            End If
        End With 'fraOpen
     Else
        With fraOpen
            .Width = .Width - 15
            If .Width <= 15 Then
                .Width = 1
                .Visible = False
                tmrOpen2.Enabled = False
                blnOpen = Not blnOpen
                intStage = 0
                fraOpen.Visible = False
            End If
        End With 'fraOpen

    End If

End Sub

Private Sub tmrFIntro_Timer()

    DrawPANFrame lngFrame, panIntro, Me.hdc, True
    If lngFrame = 14 Then
        tmrFIntro.Enabled = False
        blnShutter = True
        frmPlayer.Refresh
        Display (True)
        Bars (True)
     Else 'NOT LNGFRAME...
        lngFrame = lngFrame + 1
    End If

End Sub

Private Sub tmrFmod_Timer()

    ocxFMOD.FMODTimer
    If chkVMeter.Value Then
        If ocxFMOD.PlayState = 3 Then
            picPeak.Height = ((ocxFMOD.VolumeLevelLeft + ocxFMOD.VolumeLevelRight) / 2) * 111
        End If
    End If
    
End Sub

Private Sub tmrRIntro_Timer()

    DrawPANFrame lngFrame, panIntro, Me.hdc, True
    If lngFrame = 1 Then
        tmrRIntro.Enabled = False
        blnShutter = False
     Else 'NOT LNGFRAME...
        lngFrame = lngFrame - 1
    End If

End Sub

Private Sub tmrTime_Timer()

    picProgress.Width = ((ocxFMOD.CurrentPosition / ocxFMOD.Duration) * 244) + 1
    s = FormatTime(ocxFMOD.CurrentPosition)
    With picTime
        .Cls
        .CurrentX = 105 - .TextWidth(s)
        .CurrentY = 0
        picTime.Print (s)
    End With 'picTime
    tmrx = tmrx - intTitleSpeed
    With picTitle
        .Cls
        .CurrentX = tmrx
        .CurrentY = 0
        picTitle.Print (frmOpen.ShortName)
        If tmrx <= -.TextWidth(frmOpen.ShortName) Then
            tmrx = .ScaleWidth
        End If
    End With 'picTitle

End Sub

Private Sub tmrVolume_Timer()

    If blnUpVolume Then
        intVolume = intVolume + 16
        If intVolume > 255 Then
            intVolume = 255
        End If
        fSETVOLUME (intVolume)
    End If
    If blnDownVolume Then
        intVolume = intVolume - 16
        If intVolume < 0 Then
            intVolume = 0
        End If
        fSETVOLUME (intVolume)
    End If
    If blnFF Then
        If ocxFMOD.CurrentPosition < (ocxFMOD.Duration - 5) Then
            fCHANGEPOS (ocxFMOD.CurrentPosition + 5)
        End If
    End If
    If blnRR Then
        If ocxFMOD.CurrentPosition > 5 Then
            fCHANGEPOS (ocxFMOD.CurrentPosition - 5)
        End If
    End If
    picVolume.Height = (intVolume * 111) / 255

End Sub

Private Sub LinesAbout(b As Boolean)
Dim I As Integer
For I = 0 To 3
    lnAbout(I).Visible = b
Next I
End Sub

Private Sub LinesPL(b As Boolean)
Dim I As Integer
For I = 0 To 3
    lnPL(I).Visible = b
Next I
End Sub

Private Sub txtTitleSpeed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (txtTitleSpeed.Text >= 0) And (txtTitleSpeed.Text < 13) Then
        intTitleSpeed = txtTitleSpeed.Text
    Else
        MsgBox "Value must be between 0 and 12", vbCritical, "Error"
    End If
End If
End Sub

Private Sub fPLAY()
    If blnCD Then
        ocxFMOD.CDPlay
    Else
        ocxFMOD.Play
    End If
End Sub

Private Sub fSTOP()
    If blnCD Then
        ocxFMOD.CDStop
    Else
        ocxFMOD.StopStream
    End If
End Sub

Private Sub fPAUSE()
    If blnCD Then
        ocxFMOD.CDPause
    Else
        ocxFMOD.Pause
    End If
End Sub

Private Sub fCHANGEPOS(POS As Long)
    If blnCD Then
        ocxFMOD.CDChangePos (POS)
    Else
        ocxFMOD.ChangePos (POS)
    End If
End Sub

Private Sub fSETVOLUME(VOLUME As Integer)
    If blnCD Then
        ocxFMOD.CDSetVolume (VOLUME)
    Else
        ocxFMOD.SetVolume (VOLUME)
    End If
End Sub

Private Sub setCDName()
frmOpen.ShortName = "Track " & ocxFMOD.CDTrack & " \ " & ocxFMOD.CDNumTracks
End Sub

Private Sub setDuration()
lblDuration.Caption = FormatTime(ocxFMOD.Duration)
End Sub

Private Sub PLStart()
ocxFMOD.OpenFile lstPlaylist.List(PLsongindex), 16
frmOpen.ShortName = Right(lstPlaylist.List(PLsongindex), Len(lstPlaylist.List(PLsongindex)) - InStrRev(lstPlaylist.List(PLsongindex), "\", , vbTextCompare))
setDuration
End Sub
