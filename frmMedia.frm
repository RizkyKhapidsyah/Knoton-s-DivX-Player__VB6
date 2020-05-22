VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMedia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Knoton´s API DivX Player"
   ClientHeight    =   1155
   ClientLeft      =   5430
   ClientTop       =   8865
   ClientWidth     =   5775
   Icon            =   "frmMedia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "DivX Player"
   MaxButton       =   0   'False
   ScaleHeight     =   77
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   Begin VB.CommandButton cmdMedia 
      Appearance      =   0  'Flat
      Caption         =   "Help"
      Height          =   255
      Index           =   4
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   615
   End
   Begin VB.CheckBox chkFullscreen 
      Caption         =   "Fullscreen"
      Height          =   195
      Left            =   2400
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   840
      Width           =   1035
   End
   Begin VB.CommandButton cmdSize 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1380
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmdSize 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   840
      Width           =   255
   End
   Begin ComctlLib.Slider VolumeSlider 
      Height          =   315
      Left            =   4020
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   780
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      _Version        =   327682
      LargeChange     =   100
      SmallChange     =   10
      Max             =   100
      SelStart        =   100
      TickStyle       =   3
      Value           =   100
   End
   Begin VB.CommandButton cmdMedia 
      Appearance      =   0  'Flat
      Caption         =   "Open"
      Height          =   255
      Index           =   3
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   615
   End
   Begin ComctlLib.Slider SeekSlider 
      Height          =   315
      Left            =   780
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   420
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   556
      _Version        =   327682
      LargeChange     =   0
      SmallChange     =   0
      TickStyle       =   3
      TickFrequency   =   1000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4140
      Top             =   1200
   End
   Begin VB.CommandButton cmdMedia 
      Appearance      =   0  'Flat
      Caption         =   "Stop"
      Height          =   255
      Index           =   2
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdMedia 
      Appearance      =   0  'Flat
      Caption         =   "Play"
      Height          =   255
      Index           =   0
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdMedia 
      Appearance      =   0  'Flat
      Caption         =   "Pause"
      Height          =   255
      Index           =   1
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   615
   End
   Begin VB.Label lblSize 
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   840
      Width           =   435
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Seek"
      Height          =   195
      Left            =   840
      TabIndex        =   10
      Top             =   180
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      Height          =   195
      Left            =   3480
      TabIndex        =   9
      Top             =   840
      Width           =   555
   End
   Begin VB.Label lblCurTime 
      BackStyle       =   0  'Transparent
      Caption         =   "0:00:00"
      Height          =   195
      Left            =   4020
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblDuration 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No media"
      Height          =   195
      Left            =   780
      TabIndex        =   4
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Resize"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   555
   End
End
Attribute VB_Name = "frmMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************Information********************************'
'* This is a application to play .avi, .asf, .mpg, .mpeg, .wmv videos.      *'
'* You need to have all necessary codecs already installed on your computer.*'
'* This player can increase/decrease movie size by 25 % up to 300 %         *'
'* And of course Fullscreen.                                                *'
'* The application is tested on Windows XP/Win2K/NT                         *'
'* If it has unexpected behaviours on other systems please let me know      *'
'* what system you run it on and what function failed.                      *'
'* I am new at MCI programming and if you have some tips/info to share      *'
'* I would be very grateful.                                                *'
'****************************************************************************'
'****************************Contact Developer*******************************'
'*              Developer: Kenneth Hedman alias Knoton                      *'
'*              Mail:      knoton@hotmail.com                               *'
'*              webpage:   http://www.knoton.dns2go.com                     *'
'****************************************************************************'

Option Explicit
Private CD As CommonDialog

'Fullscreen
Private Sub chkFullscreen_Click()
If chkFullscreen.Value = 1 Then
    chkFullscreen.Value = 0
    Fullscreen = True
    PlayMedia
Else
    Fullscreen = False
    PlayMedia
End If
blnPause = True
End Sub

'The Commandobuttons
Private Sub cmdMedia_Click(Index As Integer)
Dim ret As Long
Dim tmp As String
Select Case Index
    Case 0 'Play
        If blnMediaChoosen Then
            PlayMedia
            blnPause = True
        End If
    Case 1 'Pause
        If blnMediaChoosen Then
            PauseMedia
            blnPause = False
        End If
    Case 2 'Stop
        If blnMediaChoosen Then
            Timer1.Enabled = False
            lblCurTime.Caption = "0:00:00"
            SeekSlider.Value = 0
            Call MoveMedia(0)
            Call PauseMedia
            blnPause = False
        End If
    Case 3 'Open
        If blnMediaChoosen Then
            Clear
            CloseMedia
        End If
        intSize = 100
        lblSize.Caption = intSize & " %"
        CD.ShowOpen
        tmp = CD.FileName
        If tmp <> "" Then
            MediaPath = """" & tmp & """"
            ret = OpenMedia
            If ret = 0 Then
                Call GetSize
                lblDuration.Caption = Format(MediaDuration, "h:mm:ss")
                SeekSlider.Max = MediaLengthMS
                ResizeMovie
                Call SetVolume(VolumeSlider.Value)
                PlayMedia
                Timer1.Enabled = True
                blnPause = True
            Else
                MsgBox "Media cant be played!", vbCritical
            End If
        End If
    Case 4
        frmHelp.Show
End Select
End Sub

'Increase/Decrease Moviesize
Private Sub cmdSize_Click(Index As Integer)
If blnMediaChoosen Then
    Select Case Index
        Case 0
            If intSize < 300 Then intSize = intSize + 25
        Case 1
            If intSize > 50 Then intSize = intSize - 25
    End Select
    If intSize = 100 Then
        Call ResizeMovie
    Else
        Call ResizeMovie(CCur(intSize))
    End If
    lblSize.Caption = intSize & " %"
End If
End Sub

Private Sub Form_Load()
Dim ret As Long
'Set initial values
intSize = 100
MediaVolume = 100
Set CD = New CommonDialog
CD.Filter = "Supported Media Files|*.avi;*.asf;*.mpg;*.mpeg;*.wmv|DivX File (*.avi)|*.avi"
CD.DialogTitle = "Choose media to play"

'Disable the screensaver
Call ScreenSaverActive(False)

'Register the hotkeys
ret = RegisterHotKey(Me.hwnd, 0, MOD_CTRL, VK_ADD)
ret = RegisterHotKey(Me.hwnd, 1, MOD_CTRL, VK_SUBTRACT)
ret = RegisterHotKey(Me.hwnd, 2, MOD_CTRL, VK_F7)
ret = RegisterHotKey(Me.hwnd, 3, MOD_CTRL, VK_F5)
ret = RegisterHotKey(Me.hwnd, 4, MOD_CTRL, VK_F6)
ret = RegisterHotKey(Me.hwnd, 5, MOD_CTRL, VK_DOWN)
ret = RegisterHotKey(Me.hwnd, 6, MOD_CTRL, VK_UP)

' Subclassing the form to get the Windows callback msgs.
glWinRet = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf CallbackMsgs)

'Incase the player get associated with a movie format
If Command <> "" Then
    frmMedia.Show
    lblSize.Caption = intSize & " %"
    MediaPath = """" & Command & """"
    ret = OpenMedia
    If ret = 0 Then
        Call GetSize
        lblDuration.Caption = Format(MediaDuration, "h:mm:ss")
        SeekSlider.Max = MediaLengthMS
        ResizeMovie
        Call SetVolume(VolumeSlider.Value)
        PlayMedia
        Timer1.Enabled = True
        blnPause = True
    Else
        MsgBox "Media cant be played!", vbCritical
    End If
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer
CloseMedia
Unload frmHelp

'Enable the screensaver
Call ScreenSaverActive(True)

'Unregister the hotkeys
For i = 0 To 6
    UnregisterHotKey Me.hwnd, i
Next
End Sub

'Seek to a choosen point in the movie
Private Sub SeekSlider_Click()
If blnMediaChoosen Then Call MoveMedia(SeekSlider.Value)
End Sub

Private Sub SeekSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If blnMediaChoosen Then Timer1.Enabled = False
End Sub

Private Sub SeekSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If blnMediaChoosen Then
    Timer1.Enabled = True
    blnPause = True
End If
End Sub

'Choose Volume 0 to 100 %
Private Sub VolumeSlider_Click()
If blnMediaChoosen Then Call SetVolume(VolumeSlider.Value)
MediaVolume = VolumeSlider.Value
End Sub

'Get current position in media
'If end of movie rewind and pause
'Set On/Off Play/Pause commandbuttons
Private Sub Timer1_Timer()
RunTime = GetCurrentMediaPos
If RunTime < MediaLengthMS Then
    SeekSlider.Value = RunTime
    lblCurTime.Caption = Format(FormatCount(RunTime), "h:mm:ss")
Else
    Timer1.Enabled = False
    lblCurTime.Caption = "0:00:00"
    SeekSlider.Value = 0
    Call MoveMedia(0)
    Call PauseMedia
    blnPause = False
End If

Select Case blnPause
    Case False
        cmdMedia(1).Visible = False
        cmdMedia(0).Visible = True
    Case True
        cmdMedia(1).Visible = True
        cmdMedia(0).Visible = False
End Select
End Sub
'Disable Timer and set some initial values
Private Sub Clear()
Timer1.Enabled = False
SeekSlider.Value = 0
lblDuration.Caption = "No Media"
lblCurTime.Caption = "0:00:00"
cmdMedia(1).Visible = False
cmdMedia(0).Visible = True
End Sub

