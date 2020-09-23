VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Optionen"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4575
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2280
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2280
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Abbrechen"
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Übernehmen"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      Max             =   100
      Min             =   1
      TabIndex        =   4
      Top             =   2280
      Value           =   10
      Width           =   4215
   End
   Begin VB.Label Label8 
      Caption         =   "000"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "langsam"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "schnell"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Spielgeschwindigkeit:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Schnelligkeit:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Reaktion:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Parameter für Gegenspieler:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

If Combo1.ListIndex = 0 Then RotSpeed = 10
If Combo1.ListIndex = 1 Then RotSpeed = 30
If Combo1.ListIndex = 2 Then RotSpeed = 50

If Combo2.ListIndex = 2 Then RotZoneA = 2880
If Combo2.ListIndex = 1 Then RotZoneA = 4320
If Combo2.ListIndex = 1 Then RotZoneA = 5760

Ii = HScroll1.Value * 100

Unload Me

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Form_Activate()

Label8.Caption = Str(HScroll1.Value)

End Sub

Private Sub Form_Load()

Combo1.AddItem "Langsam"
Combo1.AddItem "Normal"
Combo1.AddItem "Schnell"

Combo2.AddItem "Langsam"
Combo2.AddItem "Normal"
Combo2.AddItem "Schnell"

If RotSpeed = 10 Then Combo1.ListIndex = 0
If RotSpeed = 30 Then Combo1.ListIndex = 1
If RotSpeed = 50 Then Combo1.ListIndex = 2

If RotZoneA = 2880 Then Combo2.ListIndex = 2
If RotZoneA = 4320 Then Combo2.ListIndex = 1
If RotZoneA = 5760 Then Combo2.ListIndex = 0

HScroll1.Value = Ii / 100

End Sub

Private Sub HScroll1_Change()

Label8.Caption = Str(HScroll1.Value)

End Sub
