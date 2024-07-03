VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Turkmen War"
   ClientHeight    =   7845
   ClientLeft      =   2370
   ClientTop       =   1755
   ClientWidth     =   11925
   FillColor       =   &H8000000F&
   ForeColor       =   &H80000010&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11925
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5520
      Top             =   840
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   6840
      Top             =   240
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5520
      Top             =   240
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6240
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4800
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4080
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Basla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   6
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Line Line6 
      X1              =   240
      X2              =   0
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line5 
      X1              =   11520
      X2              =   11880
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line4 
      X1              =   3000
      X2              =   2520
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line3 
      X1              =   8520
      X2              =   9000
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line2 
      X1              =   6240
      X2              =   5520
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line1 
      X1              =   5880
      X2              =   5880
      Y1              =   5880
      Y2              =   7680
   End
   Begin VB.Image Image10 
      Height          =   1635
      Left            =   9000
      Picture         =   "Form1.frx":030A
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   2595
   End
   Begin VB.Image Image9 
      Height          =   1635
      Left            =   3000
      Picture         =   "Form1.frx":2204
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   2595
   End
   Begin VB.Image Image8 
      Height          =   1605
      Left            =   6240
      Picture         =   "Form1.frx":40FE
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   2355
   End
   Begin VB.Image Image7 
      Height          =   1605
      Left            =   240
      Picture         =   "Form1.frx":5BAA
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   2355
   End
   Begin VB.Image Image6 
      Height          =   1995
      Left            =   0
      Picture         =   "Form1.frx":7656
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   11985
   End
   Begin VB.Image Image4 
      Height          =   1965
      Left            =   5640
      Picture         =   "Form1.frx":9EA9
      Top             =   1200
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Image Image3 
      Height          =   1965
      Left            =   3120
      Picture         =   "Form1.frx":CD96
      Top             =   1200
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label Label10 
      Caption         =   "1"
      Height          =   375
      Left            =   10080
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "1"
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      Height          =   375
      Left            =   10800
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "5"
      Height          =   375
      Left            =   8640
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "100"
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "6"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "105"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Esger (Level 1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7560
      TabIndex        =   1
      Top             =   720
      Width           =   3600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Esger (Level 2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   3600
   End
   Begin VB.Image Image2 
      Height          =   4695
      Left            =   7560
      Picture         =   "Form1.frx":FC83
      Top             =   1320
      Width           =   3585
   End
   Begin VB.Image Image1 
      Height          =   4695
      Left            =   360
      Picture         =   "Form1.frx":1423C
      Top             =   1200
      Width           =   3585
   End
   Begin VB.Image Image5 
      Height          =   6990
      Left            =   0
      Picture         =   "Form1.frx":1879F
      Top             =   0
      Width           =   11940
   End
   Begin VB.Menu Fayl 
      Caption         =   "Fayl"
      Begin VB.Menu Cyk 
         Caption         =   "Cyk"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu SazlamalarB 
      Caption         =   "Sazlamalar"
      Begin VB.Menu Sazlamalar1 
         Caption         =   "Sazlamalar"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu Komek 
      Caption         =   "Komek"
      Begin VB.Menu Klawisler1 
         Caption         =   "Klawisler"
         Shortcut        =   {F5}
      End
      Begin VB.Menu OyunBarada 
         Caption         =   "Oyun Barada"
         Shortcut        =   {F6}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
Timer2.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Command1.Visible = False
End Sub

Private Sub Cyk_Click()
End
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyA Then
JanAlma1
End If
If KeyCode = vbKeyK Then
JanAlma2
End If
If KeyCode = vbKeyS Then
JanGittirmeNo1
End If
If KeyCode = vbKeyL Then
JanGittirmeNo2
End If
End Sub

Private Sub Form_Load()
Call Jan1
Call Jan2
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image10_Click()
JanGittirmeNo2
End Sub

Private Sub Image7_Click()
JanAlma1
End Sub

Private Sub Image8_Click()
JanAlma2
End Sub

Private Sub Image9_Click()
JanGittirmeNo1
End Sub

Private Sub Klawisler1_Click()
Klawisler.Show
End Sub

Private Sub OyunBarada_Click()
ProgrammaBarada.Show
End Sub

Private Sub Sazlamalar1_Click()
Sazlamalar.Show
End Sub

Private Sub Timer1_Timer()
Call Jan1
Call Jan2
Call Yeniji
End Sub

Private Sub Timer2_Timer()
Call HereketlenmeX1
Call HereketlenmeX2
Call JanHereketi1
Call JanHereketi2
End Sub

Private Sub Timer3_Timer()
Call Attack1
Call Attack2
End Sub

Private Sub Timer4_Timer()
Call HereketlenmeY1
Call HereketlenmeY2
End Sub

Private Sub Timer5_Timer()
Timer3.Enabled = True
Timer6.Enabled = True
End Sub

Private Sub Timer6_Timer()
Dim sn1, sn2 As Integer
sn1 = Form1.Label7.Caption
sn2 = Form1.Label8.Caption
Call AttackHereket1
Call AttackHereket2
If sn1 = 1 Then
    Form1.Image3.Visible = False
    sn1 = 0
    Form1.Label7.Caption = sn1
End If

If sn2 = 1 Then
    Form1.Image4.Visible = False
    sn2 = 0
    Form1.Label8.Caption = sn2
End If
End Sub
