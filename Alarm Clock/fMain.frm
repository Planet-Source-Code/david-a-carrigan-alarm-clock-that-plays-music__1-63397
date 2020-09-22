VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fMain 
   Caption         =   "Alarm Clock"
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   3405
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tFileName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Please load media."
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   2400
      TabIndex        =   10
      Top             =   1070
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Wake Time"
      Height          =   690
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3130
      Begin VB.ComboBox cmbTimeOfDay 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "fMain.frx":0000
         Left            =   1560
         List            =   "fMain.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   220
         Width           =   615
      End
      Begin VB.TextBox tWakeTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox tWakeTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   600
         TabIndex        =   5
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox tWakeTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "00"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdLoadSong 
      Caption         =   "..."
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1680
      Width           =   495
   End
   Begin MSComDlg.CommonDialog LoadSong 
      Left            =   2160
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSnooze 
      Caption         =   "Snooze!   (+10)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Timer tUpdate 
      Interval        =   1000
      Left            =   2400
      Top             =   360
   End
   Begin VB.Label lCurrentTime 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "##:##:## XX"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3165
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private MediaControl As IMediaControl
Private MediaPosition As IMediaPosition
Private BasicAudio As IBasicAudio
Private MediaEvent As IMediaEvent

Private SongPlaying As Boolean
Private TimerStart As Boolean
Private StopTime As String
Private FileName As String


Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub cmdGO_Click()
    If cmdGO.Caption = "Start" Then
        If tWakeTime(0).Text <> "" Then
            If tWakeTime(1).Text <> "" Then
                If cmbTimeOfDay.Text <> "" Then
                    If FileName = "" Then
                        LoadSong.Filter = "All Files | *.*"
                        LoadSong.ShowOpen
                        FileName = LoadSong.FileName
                    End If
                    TimerStart = True
                    StopTime = tWakeTime(0).Text & ":" & Format(tWakeTime(1).Text, "00") & ":00" & " " & cmbTimeOfDay.Text
                    cmdGO.Caption = "Stop"
                    tWakeTime(0).Enabled = False
                    tWakeTime(1).Enabled = False
                    tWakeTime(2).Enabled = False
                    cmbTimeOfDay.Enabled = False
                    Height = 4500
                Else
                    MsgBox "Please select a time of day.", vbCritical
                End If
            Else
                MsgBox "Please type in a correct time!", vbCritical
            End If
        Else
            MsgBox "Please type in a correct time!", vbCritical
        End If
    Else
        TimerStart = False
        StopTime = ""
        cmdGO.Caption = "Start"
        tWakeTime(0).Enabled = True
        tWakeTime(1).Enabled = True
        tWakeTime(2).Enabled = True
        cmbTimeOfDay.Enabled = True
        Height = 2790
    End If
End Sub

Private Sub cmdSnooze_Click()
Dim Buffer As Integer

    Buffer = tWakeTime(1).Text + 10
    If Buffer >= 60 Then
        Buffer = Buffer - 60
        If tWakeTime(0).Text = 12 Then
            If cmbTimeOfDay = "AM" Then
                cmbTimeOfDay = "PM"
            Else
                cmbTimeOfDay = "AM"
            End If
            tWakeTime(0).Text = 1
        Else
            tWakeTime(0).Text = tWakeTime(0).Text + 1
        End If
    End If
    tWakeTime(1).Text = Buffer
    MediaControl.Stop
    TimerStart = True
    SongPlaying = False
    StopTime = tWakeTime(0).Text & ":" & Format(tWakeTime(1).Text, "00") & ":00" & " " & cmbTimeOfDay.Text
End Sub

Private Sub cmdLoadSong_Click()
    LoadSong.Filter = "All Files | *.*"
    LoadSong.ShowOpen
    FileName = LoadSong.FileName
    tFileName.Text = FileName
End Sub

Private Sub Form_Load()
    lCurrentTime.Caption = Format(Now, "Long Time")
End Sub

Public Sub RunVideo(sFilePath As String, FullScreen As Boolean)
On Error Resume Next

    Set MediaControl = New FilgraphManager
    Set BasicAudio = MediaControl
    Set MediaPosition = MediaControl
    MediaControl.RenderFile sFilePath
    MediaControl.Run
    MediaPosition.CurrentPosition = 0
End Sub

Private Sub tUpdate_Timer()
    If (TimerStart) Then
        If Format(Now, "Long Time") = StopTime Then
            If Not (SongPlaying) Then
                RunVideo FileName, False
                SongPlaying = True
            End If
        End If
    End If
    lCurrentTime.Caption = Format(Now, "Long Time")
End Sub

Private Sub tWakeTime_Change(Index As Integer)
On Error Resume Next

    Select Case Index
        Case 0
            If tWakeTime(0).Text > 12 Then
                MsgBox "Hours must be Equal to or less then 12, Dumbass!", vbCritical
                tWakeTime(0).Text = "00"
                tWakeTime(0).SelStart = 0
                tWakeTime(0).SelLength = 2
            End If
        Case 1
            If tWakeTime(1).Text > 59 Then
                MsgBox "Hours must be Equal to or less then 59, Dumbass!", vbCritical
                tWakeTime(1).Text = "00"
                tWakeTime(1).SelStart = 0
                tWakeTime(1).SelLength = 2
            End If
    End Select
End Sub

Private Sub tWakeTime_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub tWakeTime_LostFocus(Index As Integer)
    If Left(tWakeTime(0).Text, 1) = "0" Then
        tWakeTime(0).Text = Right(tWakeTime(0).Text, 1)
    End If
    If tWakeTime(1).Text = "" Then
        tWakeTime(1).Text = "00"
    End If
    tWakeTime(1).Text = Format(tWakeTime(1).Text, "00")
End Sub
