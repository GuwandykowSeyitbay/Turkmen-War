Attribute VB_Name = "Finish"
Function Yeniji()
Dim Yeniji1, Yeniji2 As Integer
Yeniji1 = Form1.Label3.Caption
Yeniji2 = Form1.Label5.Caption
If Yeniji1 <= 0 Then
TimerDurduruju
If MsgBox("Yeniji: Oyuncy 2. Oyuna gaytadan baslajakmy?", vbYesNo, _
"Turkmen War") = vbYes Then
Form1.Command1.Visible = True
Gaytadan
JanBeriji
Sazlamalar.Show
Jan1
Jan2
Else
End
End If
End If
If Yeniji2 <= 0 Then
TimerDurduruju
If MsgBox("Yeniji: Oyuncy 1. Oyuna gaytadan baslajakmy?", vbYesNo, _
"Turkmen War") = vbYes Then
Form1.Command1.Visible = True
Gaytadan
JanBeriji
Sazlamalar.Show
Jan1
Jan2
Else
End
End If
End If
End Function
Function TimerDurduruju()
Form1.Timer1.Enabled = False
Form1.Timer2.Enabled = False
Form1.Timer3.Enabled = False
Form1.Timer4.Enabled = False
Form1.Timer5.Enabled = False
Form1.Timer6.Enabled = False
End Function
Function Gaytadan()
Form1.Image1.Top = 1200
Form1.Image1.Left = 360
Form1.Image2.Top = 1320
Form1.Image2.Left = 7560
Form1.Image3.Visible = False
Form1.Image4.Visible = False
JanHereketi1
JanHereketi2
End Function
Function JanBeriji()
Form1.Label3.Caption = 105
Form1.Label5.Caption = 100
End Function
