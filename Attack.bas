Attribute VB_Name = "Attack"
Function Attack1()
Dim sn As Integer
sn = Form1.Label7.Caption
Call AttackHereket1
Call JanGittirme1
If sn = 1 Then
    Form1.Image3.Visible = False
    sn = 0
    Form1.Label7.Caption = sn
End If
End Function
Function Attack2()
Dim sn As Integer
sn = Form1.Label8.Caption
Call AttackHereket2
Call JanGittirme2
If sn = 1 Then
    Form1.Image4.Visible = False
    sn = 0
    Form1.Label8.Caption = sn
End If
End Function
