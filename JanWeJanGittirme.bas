Attribute VB_Name = "JanWeJanGittirme"
Function Jan1()
Dim JN As String
JN = Form1.Label3.Caption
Form1.Label1.Caption = "Esger (Level 2) Jan: " & JN
End Function
Function Jan2()
Dim JN As String
JN = Form1.Label5.Caption
Form1.Label2.Caption = "Esger (Level 1) Jan: " & JN
End Function
Function JanGittirme1()
Dim Jan, JanGittirme, Wagt As String
Jan = Form1.Label5.Caption
JanGittirme = Form1.Label4.Caption
Wagt = Form1.Label9.Caption
Jan = Jan - JanGittirme
Form1.Label5.Caption = Jan
If Wagt = "2" Then
Jan = Jan - JanGittirme
Form1.Label5.Caption = Jan
End If
If Jan <= 0 Then
Jan = 0
Form1.Label5.Caption = Jan
End If
End Function
Function JanGittirme2()
Dim Jan, JanGittirme, Wagt As String
Jan = Form1.Label3.Caption
JanGittirme = Form1.Label6.Caption
Wagt = Form1.Label10.Caption
Jan = Jan - JanGittirme
Form1.Label3.Caption = Jan
If Wagt = "2" Then
Jan = Jan - JanGittirme
Form1.Label3.Caption = Jan
End If
If Jan <= 0 Then
Jan = 0
Form1.Label3.Caption = Jan
End If
End Function
