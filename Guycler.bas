Attribute VB_Name = "Guycler"
Function JanAlma1()
Dim Jan As Integer
Jan = Form1.Label3.Caption
Jan = Jan + 50
Form1.Label3.Caption = Jan
End Function

Function JanAlma2()
Dim Jan As Integer
Jan = Form1.Label5.Caption
Jan = Jan + 50
Form1.Label5.Caption = Jan
End Function

Function JanGittirmeNo1()
Dim JG As Integer
JG = Form1.Label5.Caption
JG = JG - 20
Form1.Label5.Caption = JG
If JG <= 0 Then
JG = 0
Form1.Label5.Caption = JG
End If
End Function

Function JanGittirmeNo2()
Dim JG As Integer
JG = Form1.Label3.Caption
JG = JG - 20
Form1.Label3.Caption = JG
If JG <= 0 Then
JG = 0
Form1.Label3.Caption = JG
End If
End Function
