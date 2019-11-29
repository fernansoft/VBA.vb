Private Sub CaixaEletronico_Click()
Dim opcmenu As Integer
Call fmenu
Dim vnotasnocaixa(6) As Integer
End Sub
'Procedimento opcmenu principal (gerencia qual vai ser a função chamada)
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
        Case 3
        Case 9
    End Select
End Sub
'Procedimento carregar notas insere por input quantas notas de cada cédula o caixa terá
Sub fcarregarnotas()
    Dim vnotasnocaixa(6) As Integer
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
