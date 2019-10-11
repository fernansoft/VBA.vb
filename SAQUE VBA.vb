Dim n50, n20, n10, n05, n02 as Integer
Dim saque as Integer
InputBox("Isira o valor a ser sacado: ")
While saque > 13
	While saque > 53
		saque = saque - 50
		n50 = n50 + 1
	Wend
	While saque > 23
		saque = saque - 20
		n20 = n20 + 1
	Wend
	While saque > 13
		saque = saque - 10
		n10 = n10 + 1
	Wend
Wend
If ((saque = 11) or (saque = 13)) Then
	saque = saque - 5
	n05 = n05 + 1
End If
If ((saque = 6) or (saque = 8)) Then
	While saque >= 2
		saque = saque - 2
		n02 = n02 + 1
	Wend
	Else
		While saque >= 10
			saque = saque - 10
			n10 = n10 + 1
		Wend
		While saque >= 5
			saque = saque - 5
			n05 = n05 + 1
		Wend
		While saque >= 2
			saque = saque - 2
			n05 = n02 + 1
		Wend
End If
MsgBox("Você irá sacar:" & )

