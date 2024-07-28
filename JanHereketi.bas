Attribute VB_Name = "JanHereketi"
Function JanHereketi1()
Dim JanX, JanY As Integer
JanX = Form1.Image1.Left
JanY = Form1.Image1.Top
Form1.Label1.Left = JanX
Form1.Label1.Top = JanY - 480
End Function
Function JanHereketi2()
Dim JanX, JanY As Integer
JanX = Form1.Image2.Left
JanY = Form1.Image2.Top
Form1.Label2.Left = JanX
Form1.Label2.Top = JanY - 480
End Function
