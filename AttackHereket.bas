Attribute VB_Name = "AttackHereket"
Function AttackHereket1()
Dim san As Byte
Dim sn As Integer
sn = Form1.Label7.Caption
Randomize
san = Int(Rnd(1) * 3) + 1
If san = 1 Then
    sn = sn + 1
    Form1.Label7.Caption = sn
    Form1.Image3.Visible = True
    Form1.Image3.Top = 1200
    Form1.Image3.Left = 3120
End If
If san = 2 Then
    sn = sn + 1
    Form1.Label7.Caption = sn
    Form1.Image3.Visible = True
    Form1.Image3.Top = 2640
    Form1.Image3.Left = 1800
End If
If san = 3 Then
    sn = sn + 1
    Form1.Label7.Caption = sn
    Form1.Image3.Visible = True
    Form1.Image3.Top = 3960
    Form1.Image3.Left = 3360
End If
End Function
Function AttackHereket2()
Dim san As Byte
Dim sn As Integer
sn = Form1.Label8.Caption
Randomize
san = Int(Rnd(1) * 3) + 1
If san = 1 Then
    sn = sn + 1
    Form1.Label8.Caption = sn
    Form1.Image4.Visible = True
    Form1.Image4.Top = 1200
    Form1.Image4.Left = 5640
End If
If san = 2 Then
    sn = sn + 1
    Form1.Label8.Caption = sn
    Form1.Image4.Visible = True
    Form1.Image4.Top = 2520
    Form1.Image4.Left = 7080
End If
If san = 3 Then
    sn = sn + 1
    Form1.Label8.Caption = sn
    Form1.Image4.Visible = True
    Form1.Image4.Top = 4080
    Form1.Image4.Left = 5280
End If
End Function
