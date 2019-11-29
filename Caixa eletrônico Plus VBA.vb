Dim n100 As Integer
Dim n50 As Integer
Dim n20 As Integer
Dim n10 As Integer
Dim n05 As Integer
Dim n02 As Integer
Dim saldocaixa As Integer
Dim codigobanco As Integer
Dim valorsaque As Integer
Dim opcmenu As Integer
'Procedimento opcmenu principal (gerencia qual vai ser a função chamada)
Dim vbancos()
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
            ffim
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
    ' vnotasnocaixa(0) = notas de R$100,00
    ' vnotasnocaixa(1) = notas de R$50,00
    ' vnotasnocaixa(2) = notas de R$20,00
    ' vnotasnocaixa(3) = notas de R$10,00
    ' vnotasnocaixa(4) = notas de R$05,00
    ' vnotasnocaixa(5) = notas de R$02,00
    MsgBox ("Pronto! Agora o caixa tem:" & vbCrLf & vnotasnocaixa(0) & " notas de R$100,00" & vbCrLf & vnotasnocaixa(1) & " notas de R$50,00" & vbCrLf & vnotasnocaixa(2) & " notas de R$20,00" & vbCrLf & vnotasnocaixa(3) & " notas de R$10,00" & vbCrLf & vnotasnocaixa(4) & " notas de R$05,00" & vbCrLf & vnotasnocaixa(5) & " notas de R$02,00")
    saldocaixa = saldocaixa + vnotasnocaixa(0) * 100 + vnotasnocaixa(1) * 50 + vnotasnocaixa(2) * 20 + vnotasnocaixa(3) * 10 + vnotasnocaixa(4) * 5 + vnotasnocaixa(5) * 2
    Call fmenu
End Sub
Sub fsaquedecrescente()
    Call fcodigodobanco
    If vnotasnocaixa(0) = 0 And vnotasnocaixa(1) = 0 And vnotasnocaixa(2) = 0 And vnotasnocaixa(3) = 0 And vnotasnocaixa(4) = 0 And vnotasnocaixa(5) = 0 Then
        MsgBox ("O caixa não possuí notas carregadas no momento! Favor carregar as notas.")
        Call fmenu
    End If
    valorsaque = InputBox("insira o valor a ser sacado: ")
    If saque > saldocaixa Then
        MsgBox ("Você excedeu o limite do caixa!")
        Call fmenu
    End If
    If vnotasnocaixa(0) >= 1 Then
        While valorsaque > 103 Or valorsaque = 100
            valorsaque = valorsaque - 100
            n100 = n100 + 1
            vnotasnocaixa(0) = vnotasnocaixa(0) - 1
        Wend
    ElseIf vnotasnocaixa(0) = 0 Then
        MsgBox ("Ops, acabou as notas de R$100,00")
    End If
    If vnotasnocaixa(1) >= 1 Then
        While valorsaque > 53 Or valorsaque = 50
            valorsaque = valorsaque - 50
            n50 = n50 + 1
        Wend
    ElseIf vnotasnocaixa(1) = 0 Then
        MsgBox ("Ops, acabou as notas de R$50,00")
    End If
    If vnotasnocaixa(2) > 1 Then
        While valorsaque > 23 Or valorsaque = 20
            valorsaque = valorsaque - 20
            n20 = n20 + 1
        Wend
    ElseIf vnotasnocaixa(2) = 0 Then
        MsgBox ("Ops, acabou as notas de R$20,00")
    End If
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
Sub ffim()
    MsgBox ("Programa finalizado!")
End Sub
Sub fcodigodobanco()
'A função fcodigodobanco coleta o código do banco que o cliente tem conta.
    codigobanco = InputBox("Qual o banco na qual deseja fazer o saque?" & vbCrLf & "Digite de acordo com o código de cada um:" & vbCrLf & "1 - Banco do Brasil" & vbCrLf & "2 - Santander" & vbCrLf & "3 - Itaú" & vbCrLf & "4 - Caixa" & vbCrLf & ">>> ")
    ' Limitar input do usuário para apenas os códigos dos bancos
    While codigobanco <> 1 And codigobanco <> 2 And codigobanco <> 3 And codigobanco <> 4
        MsgBox ("Banco não cadastrado! Favor digitar apenas um dos citados no menu de seleção.")
        codigobanco = InputBox("Qual o banco na qual deseja fazer o saque?" & vbCrLf & "Digite de acordo com o código de cada um:" & vbCrLf & "1 - Banco do Brasil" & vbCrLf & "2 - Santander" & vbCrLf & "3 - Itaú" & vbCrLf & "4 - Caixa" & vbCrLf & ">>> ")
    Wend
End Sub
