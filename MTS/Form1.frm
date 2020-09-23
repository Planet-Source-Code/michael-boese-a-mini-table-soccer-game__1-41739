VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Mic's Table Soccer"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9030
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9030
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox Picture3 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   0
      MousePointer    =   3  'I-Cursor
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   5985
      ScaleWidth      =   8985
      TabIndex        =   10
      Top             =   0
      Width           =   9015
      Begin MTS.SkinControl SpRot 
         Height          =   450
         Left            =   6600
         TabIndex        =   12
         Top             =   2985
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   794
         MaskColor       =   16711935
      End
      Begin MTS.SkinControl SpBlau 
         Height          =   450
         Left            =   1950
         TabIndex        =   14
         Top             =   2625
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   794
         MaskColor       =   16711935
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5730
         Left            =   2020
         Picture         =   "Form1.frx":B0104
         ScaleHeight     =   5700
         ScaleWidth      =   150
         TabIndex        =   15
         Top             =   120
         Width           =   180
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5730
         Left            =   6650
         Picture         =   "Form1.frx":B30C6
         ScaleHeight     =   5700
         ScaleWidth      =   150
         TabIndex        =   13
         Top             =   120
         Width           =   180
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Spielen"
         Height          =   495
         Left            =   3360
         MousePointer    =   1  'Pfeil
         TabIndex        =   11
         Top             =   2760
         Width           =   2415
      End
      Begin MTS.SkinControl SpBall 
         Height          =   225
         Left            =   4200
         TabIndex        =   16
         Top             =   120
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   397
         MaskColor       =   16711935
      End
   End
   Begin VB.PictureBox PicRot 
      Height          =   495
      Index           =   3
      Left            =   6360
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PicRot 
      Height          =   495
      Index           =   2
      Left            =   5880
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PicRot 
      Height          =   495
      Index           =   1
      Left            =   5400
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PicRot 
      Height          =   495
      Index           =   0
      Left            =   4920
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Left            =   6960
      Top             =   5880
   End
   Begin VB.PictureBox PicBlau 
      Height          =   495
      Index           =   3
      Left            =   4200
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PicBlau 
      Height          =   495
      Index           =   2
      Left            =   3720
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PicBlau 
      Height          =   495
      Index           =   1
      Left            =   3240
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PicBlau 
      Height          =   495
      Index           =   0
      Left            =   2760
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Left            =   1560
      Top             =   5880
   End
   Begin VB.PictureBox picball 
      Height          =   495
      Left            =   2160
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   17
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   7680
      TabIndex        =   9
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Menu mnuSpiel 
      Caption         =   "Spiel"
      Begin VB.Menu mnuNeu 
         Caption         =   "Neues Spiel"
      End
      Begin VB.Menu mnuStrich1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnde 
         Caption         =   "Spiel beenden"
      End
   End
   Begin VB.Menu mnuOptionen 
      Caption         =   "Optionen"
      Begin VB.Menu mnuOptionChange 
         Caption         =   "Optionen ändern"
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "Info"
      Begin VB.Menu mnuInfoAnzeigen 
         Caption         =   "Info anzeigen"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************
'
' Mic's Table Soccer
'
' Mini Table Soccer Game
'
' © Copyright 2002 by M.Boese
'
' Teile des Programms wurden von Dave Scarmozzino (www.TheScarms.com)
' entwickelt und von Herfried K. Wagner erweitert.
'
' Der  Autor übernimmt keine Haftung für Schäden, die durch dieses
' Programm verursacht wurden.
' Sie sind nicht berechtigt, diesen Code weiterzugeben, ausser in Form
' einer kompilierten Anwendung!
'**********************************************************************

Private Sub Command1_Click()

Command1.Visible = False
SpBall.Visible = True

1

BallX = BallX + (Cos(BallAngle * Pi) * (BallSpeed))
BallY = BallY + (Sin(BallAngle * Pi) * (BallSpeed))

If BallX < 120 Or BallX > 8640 Then TorAbfrage
    If Command1.Visible = True Then GoTo 2
    
    If BallY > 5620 + BallSpeed Then
        If BallAngle >= 0 And BallAngle < 90 Then Angle = 360 - BallAngle
        If BallAngle >= 90 And BallAngle < 180 Then Angle = 270 - (BallAngle - 90)
    End If
    
    If BallY < 150 - BallSpeed Then
        If BallAngle >= 270 And BallAngle < 360 Then Angle = 0 + (360 - BallAngle)
        If BallAngle >= 180 And BallAngle < 270 Then Angle = 90 + (270 - BallAngle)
    End If
    
    If BallX > 8640 + BallSpeed Then
        If BallAngle >= 270 And BallAngle < 360 Then Angle = 270 - (BallAngle - 270)
        If BallAngle >= 0 And BallAngle < 90 Then Angle = 180 - BallAngle
End If

If Angle = 90 Then Angle = 96
If Angle = 270 Then Angle = 276

If BallX < 150 - BallSpeed Then
    If BallAngle >= 180 And BallAngle < 270 Then Angle = 270 + (270 - BallAngle)
    If BallAngle >= 90 And BallAngle < 180 Then Angle = 90 - (BallAngle - 90)
End If

If BallAngle <> Angle Then BallSpeed = BallSpeed - 2
If BallSpeed < 5 Then BallSpeed = 5

BallAngle = Angle
SpBall.Left = BallX
SpBall.Top = BallY

If BallX > RotZoneA And BallX < RotZoneB Then SpielerRot

If BallSpeed < 5 Then BallSpeed = 5

For t = 1 To Ii
    DoEvents
Next t

If BallX > 4000 Then Getroffen = False
If BallX < 4000 Then GetroffenRot = False


If BallX - BallSpeed < 2280 And BallX - BallSpeed > 1800 And Getroffen = False Then
    If Timer1.Interval > 0 Then
        Collision
    Else
        If BallY > SpBlau.Top - SpBall.Height And BallY < SpBlau.Top + SpBlau.Height Then
            If BallAngle >= 180 And BallAngle < 270 Then Angle = 270 + (270 - BallAngle)
            If BallAngle >= 90 And BallAngle < 180 Then Angle = 90 - (BallAngle - 90)
            If BallAngle >= 270 And BallAngle < 360 Then Angle = 270 - (BallAngle - 270)
            If BallAngle >= 0 And BallAngle < 90 Then Angle = 180 - BallAngle
            
            Getroffen = True
        End If
    End If
End If

If BallX - BallSpeed > 6300 And BallX < 6900 And BallY > SpRot.Top And BallY < SpRot.Top + SpRot.Height Then Timer2.Interval = 30

If BallX - BallSpeed > 6360 And BallX - BallSpeed < 6840 And GetroffenRot = False Then
    If Timer2.Interval > 0 Then
        CollisionRot
    Else
        If BallY > SpRot.Top - SpBall.Height And BallY < SpRot.Top + SpBlau.Height Then
            If BallAngle >= 180 And BallAngle < 270 Then Angle = 270 + (270 - BallAngle)
            If BallAngle >= 90 And BallAngle < 180 Then Angle = 90 - (BallAngle - 90)
            If BallAngle >= 270 And BallAngle < 360 Then Angle = 270 - (BallAngle - 270)
            If BallAngle >= 0 And BallAngle < 90 Then Angle = 180 - BallAngle
            
            GetroffenRot = True
        End If
    End If
End If


GoTo 1

2

If ToreBlau = 10 Then
    Label1.Caption = "Sieger": ToreBlau = 0: ToreRot = 0
End If

If ToreRot = 10 Then
    Label2.Caption = "Sieger": ToreBlau = 0: ToreRot = 0
End If

BallX = 4200: BallY = 120

End Sub

Private Sub Form_Activate()

Randomize Timer

BallX = SpBall.Left
BallY = SpBall.Top

BallAngle = Int(360 * Rnd(1)) + 1
If BallAngle = 90 Then BallAngle = 96

Angle = BallAngle

BallSpeed = 20

Pi = (4 * Atn(1)) / 180

Label1.Caption = "0"
Label2.Caption = "0"




End Sub

Private Sub Form_Load()
PicBlau(0).Picture = LoadPicture(App.Path & "\Grafik\SpBlau1.bmp")
PicBlau(1).Picture = LoadPicture(App.Path & "\Grafik\SpBlau2.bmp")
PicBlau(2).Picture = LoadPicture(App.Path & "\Grafik\SpBlau3.bmp")
PicBlau(3).Picture = LoadPicture(App.Path & "\Grafik\SpBlau4.bmp")
picball.Picture = LoadPicture(App.Path & "\Grafik\ball.bmp")
PicRot(0).Picture = LoadPicture(App.Path & "\Grafik\SpRot1.bmp")
PicRot(1).Picture = LoadPicture(App.Path & "\Grafik\SpRot2.bmp")
PicRot(2).Picture = LoadPicture(App.Path & "\Grafik\SpRot3.bmp")
PicRot(3).Picture = LoadPicture(App.Path & "\Grafik\SpRot4.bmp")


BildBlau = 0
BildRot = 0

RotSpeed = 30
RotZoneA = 4320
RotZoneB = 8640

Ii = 1000

SpBlau.Picture = PicBlau(0).Picture
SpRot.Picture = PicRot(0).Picture
SpBall.Picture = picball.Picture

End Sub

Private Sub mnuEnde_Click()

End

End Sub

Private Sub mnuInfoAnzeigen_Click()

Form3.Show vbModal

End Sub

Private Sub mnuNeu_Click()

Label1.Caption = "0": ToreBlau = 0: ToreRot = 0
Label2.Caption = "0": ToreBlau = 0: ToreRot = 0

BallX = 4200: BallY = 120
SpBall.Left = BallX: SpBall.Top = BallY

Command1.Visible = True

BallSpeed = 20

End Sub

Private Sub mnuOptionChange_Click()

Form2.Show vbModal

End Sub

Private Sub Picture3_Click()

Timer1.Interval = 30

End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    Picture3.Top = 100
    Picture3.Left = 100
    DoEvents
    
    Picture3.Top = 50
    Picture3.Left = 50
    DoEvents
    
    Picture3.Top = 200
    Picture3.Left = 200
    DoEvents
    
    Picture3.Top = 100
    Picture3.Left = 100
    DoEvents
    
    Picture3.Top = 0
    Picture3.Left = 0
    BallAngle = BallAngle + 20
    
    If BallAngle > 360 Then BallAngle = 20
    Angle = BallAngle
End If


End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Y > 135 And Y < 5400 Then
    SpBlau.Top = Y
End If

End Sub

Private Sub Timer1_Timer()

If BildBlau = 3 Then BildBlau = 4
If BildBlau = 2 Then BildBlau = 3
If BildBlau = 1 Then BildBlau = 2
If BildBlau = 0 Then BildBlau = 1

If BildBlau = 1 Or BildBlau = 3 Then
    SpBlau.Left = 1560
Else
    SpBlau.Left = 1950
End If

'Framewechsel verlangsamen, damit die Drehung echt wirkt
If BildBlau <> 4 Then
    SpBlau.Picture = PicBlau(BildBlau).Picture
Else
    BildBlau = 0
    SpBlau.Picture = PicBlau(BildBlau).Picture
    Timer1.Interval = Timer1.Interval + 30
    
    If Timer1.Interval = 90 Then Timer1.Interval = 0
End If


End Sub

Public Sub Collision()

If BallY > SpBlau.Top - SpBall.Height And BallY < SpBlau.Top + SpBlau.Height Then
    BallAngle = Int(180 * Rnd(1))
    
    If BallAngle > 90 Then
        BallAngle = BallAngle - 90
    Else
        BallAngle = 360 - BallAngle
    End If
    
    Getroffen = True: BallSpeed = 50
End If

Angle = BallAngle

End Sub

Public Sub SpielerRot()

If SpRot.Top < BallY Then SpRot.Top = SpRot.Top + RotSpeed
If SpRot.Top > BallY Then SpRot.Top = SpRot.Top - RotSpeed

End Sub

Public Sub CollisionRot()

If BallY > SpRot.Top - SpBall.Height And BallY < SpRot.Top + SpBlau.Height Then
    BallAngle = Int(180 * Rnd(1))
    BallAngle = 90 + BallAngle
    
    Getroffen = True: BallSpeed = 50
End If

Angle = BallAngle

End Sub

Private Sub Timer2_Timer()

If BildRot = 3 Then BildRot = 4
If BildRot = 2 Then BildRot = 3
If BildRot = 1 Then BildRot = 2
If BildRot = 0 Then BildRot = 1

If BildRot = 1 Or BildRot = 3 Then
    SpRot.Left = 6120
Else
    SpRot.Left = 6600
End If

If BildRot <> 4 Then
    SpRot.Picture = PicRot(BildRot).Picture
Else
    BildRot = 0
    SpRot.Picture = PicRot(BildRot).Picture
    Timer2.Interval = Timer2.Interval + 30
    
    If Timer2.Interval = 90 Then Timer2.Interval = 0
End If

End Sub

Public Sub TorAbfrage()

If BallY >= 2280 And BallY <= 3500 Then
    If BallX < 120 Then
        ToreRot = ToreRot + 1: SpBall.Visible = False
    Else
        ToreBlau = ToreBlau + 1: SpBall.Visible = False
    End If

    Command1.Visible = True
    Label1.Caption = Str(ToreBlau)
    Label2.Caption = Str(ToreRot)
    
    BallAngle = Int(360 * Rnd(1)) + 1
    If BallAngle = 90 Then BallAngle = 85
    Angle = BallAngle
    
    BallSpeed = 20
End If

End Sub
