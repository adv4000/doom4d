VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form FormMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Doom Fourth Dimension"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   5805
   Icon            =   "FormMenu.frx":0000
   LinkTopic       =   "Form"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FormMenu.frx":08CA
   MousePointer    =   99  'Custom
   Picture         =   "FormMenu.frx":0BD4
   ScaleHeight     =   3990
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox VideoMode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Mortal Kombat 4 Incomplete"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Select Video Mode"
      Top             =   3600
      Width           =   2055
   End
   Begin MCI.MMControl MMControl 
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin VB.CommandButton Command2 
      Height          =   700
      Left            =   700
      Picture         =   "FormMenu.frx":4D272
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   4500
   End
   Begin VB.CommandButton Command1 
      Height          =   700
      Left            =   700
      Picture         =   "FormMenu.frx":5B738
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   4500
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Battle Area"
      BeginProperty Font 
         Name            =   "Mortal Kombat 4 Incomplete"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   540
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "FormMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  GameType = 1  ' Game type is -> EARTH
  Call SetSelectedMode ' Set VIdeo Settings
  PlaySound (App.Path & "\Sound\Select.wav")
  FormMenu.Hide
  FormMain.Show
End Sub

Private Sub Command2_Click()
  GameType = 2  ' Game type is -> DEATH
  Call SetSelectedMode ' Set VIdeo Settings
  PlaySound (App.Path & "\Sound\Select.wav")
  FormMenu.Hide
  FormMain.Show
End Sub

Private Sub Form_Load()
 GameType = 0
 
 VideoMode.AddItem "640x480x16Bit", 0
 VideoMode.AddItem "640x480x32Bit", 1
 VideoMode.AddItem "800x600x16Bit", 2
 VideoMode.AddItem "800x600x32Bit", 3
 VideoMode.AddItem "1024x768x16Bit", 4
 VideoMode.AddItem "1024x768x32Bit", 5
 VideoMode.AddItem "1280x1024x16Bit", 6
 VideoMode.AddItem "1280x1024x32Bit", 7
 VideoMode.Text = VideoMode.List(2) ' Default Video Mode

 PlaySound (App.Path & "\Sound\Start.wav")
End Sub

Private Sub SetSelectedMode()

Select Case (VideoMode.Text)
  Case "640x480x16Bit": Call SetVideoSettings(640, 480, 16)
  Case "640x480x32Bit": Call SetVideoSettings(640, 480, 32)
  Case "800x600x16Bit": Call SetVideoSettings(800, 600, 16)
  Case "800x600x32Bit": Call SetVideoSettings(800, 600, 32)
  Case "1024x768x16Bit": Call SetVideoSettings(1024, 768, 16)
  Case "1024x768x32Bit": Call SetVideoSettings(1024, 768, 32)
  Case "1280x1024x16Bit": Call SetVideoSettings(1280, 1024, 16)
  Case "1280x1024x32Bit": Call SetVideoSettings(1280, 1024, 32)
End Select

End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub MMControl_Done(NotifyCode As Integer)
 MMControl.Command = "Close"
End Sub

Private Sub PlaySound(MusicFile As String) ' Play MP3 -> WAV File by MCI32
 MMControl.Command = "Close"
 MMControl.FileName = MusicFile
 MMControl.Command = "Open"
 MMControl.Command = "Play"
End Sub

