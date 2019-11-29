Dim n100 As Integer
Dim n50 As Integer
Dim n20 As Integer
Dim n10 As Integer
Dim n05 As Integer
Dim n02 As Integer
Dim valorsaque As Integer
Dim opcmenu As Integer
'Procedimento opcmenu principal (gerencia qual vai ser a função chamada)
Dim vnotasnocaixa(6) As Variant
Private Sub CaixaEletronico_Click()
Call fmenu
End Sub
Sub fmenu()
    opcmenu = InputBox("Escolha uma opção para prosseguir:" & vbCrLf & "1 - Carregar notas" & vbCrLf & "2 - Retirar notas" & vbCrLf & "3 - Estatísticas" & vbCrLf & "9 - Fim" & vbCrLf & ">>> ")
    While opcmenu <> 1 And opcmenu <> 2 And opcmenu <> 3 And opcmenu <> 9
        MsgBox ("Não existe essa opção, favor selecionar uma das citadas no opcmenu!")
        opcmenu = InputBox("Escolha uma opção para prosseguir:" & vbCrLf & "1 - Carregar notas" & vbCrLf & "2 - Retirar notas" & vbCrLf & "3 - Estatísticas" & vbCrLf & "9 - Fim" & vbCrLf & ">>> ")
    Wend
    Select Case opcmenu
        Case 1
            Call fcarregarnotas
        Case 2
            Call fsaquedecrescente
        Case 3
        Case 9
    End Select
End Sub
'Procedimento carregar notas insere por input quantas notas de cada cédula o caixa terá
Sub fcarregarnotas()
    vnotasnocaixa(0) = InputBox("Insira a quantidade de notas de R$100,00: ")
    vnotasnocaixa(1) = InputBox("Insira a quantidade de notas de R$50,00: ")
    vnotasnocaixa(2) = InputBox("Insira a quantidade de notas de R$20,00: ")
    vnotasnocaixa(3) = InputBox("Insira a quantidade de notas de R$10,00: ")
    vnotasnocaixa(4) = InputBox("Insira a quantidade de notas de R$05,00: ")
    vnotasnocaixa(5) = InputBox("Insira a quantidade de notas de R$02,00: ")
    ' vnotasnocaixa(1) = notas de R$100,00
    ' vnotasnocaixa(2) = notas de R$50,00
    ' vnotasnocaixa(3) = notas de R$20,00
    ' vnotasnocaixa(4) = notas de R$10,00
    ' vnotasnocaixa(5) = notas de R$05,00
    ' vnotasnocaixa(6) = notas de R$02,00
    MsgBox ("Pronto! Agora o caixa tem:" & vbCrLf & vnotasnocaixa(0) & " notas de R$100,00" & vbCrLf & vnotasnocaixa(1) & " notas de R$50,00" & vbCrLf & vnotasnocaixa(2) & " notas de R$20,00" & vbCrLf & vnotasnocaixa(3) & " notas de R$10,00" & vbCrLf & vnotasnocaixa(4) & " notas de R$05,00" & vbCrLf & vnotasnocaixa(5) & " notas de R$02,00")
    Call fmenu
End Sub
Sub fsaquedecrescente()
    If vnotasnocaixa(0) = 0 And vnotasnocaixa(1) = 0 And vnotasnocaixa(2) = 0 And vnotasnocaixa(3) = 0 And vnotasnocaixa(4) = 0 And vnotasnocaixa(5) = 0 Then
        MsgBox ("O caixa não possuí notas carregadas no momento! Favor carregar as notas.")
        Call fmenu
    End If
    valorsaque = InputBox("insira o valor a ser sacado: ")
    While valorsaque > 103 Or valorsaque = 100
        valorsaque = valorsaque - 100
        n100 = n100 + 1
    Wend
    While valorsaque > 53 Or valorsaque = 50
        valorsaque = valorsaque - 50
        n50 = n50 + 1
    Wend
    While valorsaque > 23 Or valorsaque = 20
        valorsaque = valorsaque - 20
        n20 = n20 + 1
    Wend
    While valorsaque > 13 Or valorsaque = 10
        valorsaque = valorsaque - 10
        n10 = n10 + 1
    Wend
    If valorsaque = 11 Or valorsaque = 13 Then
        valorsaque = valorsaque - 5
        n05 = n05 + 1
    End If
    If valorsaque = 6 Or valorsaque = 8 Then
        While valorsaque >= 2
            valorsaque = valorsaque - 2
            n02 = n02 + 1
        Wend
    Else
        While valorsaque >= 5
            valorsaque = valorsaque - 5
            n05 = n05 + 1
        Wend
        While valorsaque >= 2
            valorsaque = valorsaque - 2
            n02 = n02 + 1
        Wend
    End If
    MsgBox ("As notas a serem sacadas são: " & vbCrLf & (n100) & " notas de R$100,00" & vbCrLf & (n50) & " notas de R$50,00" & vbCrLf & (n20) & " notas de R$20,00 " & vbCrLf & (n10) & " notas de R$10,00 " & vbCrLf & (n05) & " notas de R$05,00 " & vbCrLf & (n02) & " notas de R$02,00")
    Call fmenu
End Sub
