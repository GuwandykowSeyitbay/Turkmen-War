VERSION 5.00
Begin VB.Form Sazlamalar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sazlamalar"
   ClientHeight    =   2745
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5115
   Icon            =   "Sazlamalar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Awtomatik Sazla"
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sazla"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Text            =   "0"
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Text            =   "0"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Text            =   "0"
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Text            =   "0"
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "2-nji oyuncynyn jan gittirmesi:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "2-nji oyuncynyn jany:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "1-nji oyuncynyn jan gittirmesi:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "1-nji oyuncynyn jany:"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Sazlamalar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Jn1, Jn2, Jngtrm1, Jngtrm2 As Integer
Jn1 = Text1.Text
Jn2 = Text3.Text
Jngtrm1 = Text2.Text
Jngtrm2 = Text4.Text
Form1.Label3.Caption = Jn1
Form1.Label5.Caption = Jn2
Form1.Label4.Caption = Jngtrm1
Form1.Label6.Caption = Jngtrm2
Jan1
Jan2
Unload Me
End Sub

Private Sub Command2_Click()
Text1.Text = 105
Text2.Text = 6
Text3.Text = 100
Text4.Text = 5
End Sub
