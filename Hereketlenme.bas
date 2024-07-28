Attribute VB_Name = "Hereketlenme"
Function HereketlenmeX1()
Dim x As Integer
x = Form1.Image1.Left
x = x + 10
Form1.Image1.Left = x
If x = 2280 Then
Form1.Timer2.Enabled = False
Form1.Timer4.Enabled = False
End If
End Function
Function HereketlenmeX2()
Dim x As Integer
x = Form1.Image2.Left
x = x - 10
Form1.Image2.Left = x
If x = 5800 Then
Form1.Timer2.Enabled = False
Form1.Timer4.Enabled = False
End If
End Function
Function HereketlenmeY1()
Dim y As Integer
y = Form1.Image1.Top
y = y - 100
Form1.Image1.Top = y
If y = 700 Then
y = y + 500
Form1.Image1.Top = y
End If
End Function
Function HereketlenmeY2()
y = Form1.Image2.Top
y = y - 100
Form1.Image2.Top = y
If y = 720 Then
y = y + 500
Form1.Image2.Top = y
End If
End Function
