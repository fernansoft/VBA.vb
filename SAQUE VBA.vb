Private Sub CommandButton1_Click()
Dim n50, n20, n10, n05, n02 As Integer
Dim saque As Integer
n50 = 0
n20 = 0
n10 = 0
n05 = 0
n02 = 0
saque = InputBox("Isira o valor a ser sacado: ")
    While ((saque > 53) or (saque = 50))
        saque = saque - 50
        n50 = n50 + 1
    Wend
    While ((saque >= 23) or (saque = 20))
        saque = saque - 20
        n20 = n20 + 1
    Wend
    While ((saque > 13) or (saque = 10))
        saque = saque - 10
        n10 = n10 + 1
    Wend
If ((saque = 11) Or (saque = 13)) Then
    saque = saque - 5
    n05 = n05 + 1
End If
If ((saque = 6) Or (saque = 8)) Then
    While saque >= 2
        saque = saque - 2
        n02 = n02 + 1
    Wend
    Else
        While (saque >= 10)
            saque = saque - 10
            n10 = n10 + 1
        Wend
        While (saque >= 5)
            saque = saque - 5
            n05 = n05 + 1
        Wend
        While (saque >= 2)
            saque = saque - 2
            n02 = n02 + 1
        Wend
End If
MsgBox (("Você irá sacar:") & vbCrLf & n50 & (" notas de R$50,00") & vbCrLf & n20 & (" notas de R$20,00") & vbCrLf & n10 & (" notas de R$10,00") & vbCrLf & n05 & (" notas de R$05,00") & vbCrLf & n02 & (" notas de R$02,00"))
End Sub
