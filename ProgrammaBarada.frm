VERSION 5.00
Begin VB.Form ProgrammaBarada 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programma Barada"
   ClientHeight    =   2580
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7515
   Icon            =   "ProgrammaBarada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdYza 
      Caption         =   "Yza"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   120
      Picture         =   "ProgrammaBarada.frx":030A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Websayt: https://sites.google.com/view/tps-com-tm"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label lblProgrammanyDuzen 
      Caption         =   "Oyuny duzen: Guwandykow Seyitbay"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblTurkmenProgramStudy 
      Caption         =   "Turkmen Program Study"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   21.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   5055
   End
   Begin VB.Image ImgOMB 
      Height          =   480
      Left            =   240
      Picture         =   "ProgrammaBarada.frx":48B5
      Top             =   840
      Width           =   480
   End
   Begin VB.Label lblProgrammaWersiyasy 
      Caption         =   "Oyun wersiyasy: 1.0.0"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblProgrammaAdy 
      Caption         =   "Oyun ady: Turkmen War"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "ProgrammaBarada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdYza_Click()
Unload Me
End Sub

