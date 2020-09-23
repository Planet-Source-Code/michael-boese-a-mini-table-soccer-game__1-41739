VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Info"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   3120
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label4 
      Caption         =   "Der Autor dieses Programms übernimmt keinerlei Haftung für Schäden oder Datenverluste, die durch dieses Programm entstehen."
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   3120
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label3 
      Caption         =   $"Form3.frx":0442
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3120
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3120
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "Dieses Programm ist Freeware und darf in komilierter Form, ohne weitere Bedingungen kopiert, geändert und verbreitet werden."
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Mic's Table Soccer V1.0 von M.Boese in 2002"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "Form3.frx":0516
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

