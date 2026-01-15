VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormularioFuncionario 
   Caption         =   "Cadastro de Funcionário"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14250
   OleObjectBlob   =   "FormularioFuncionario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormularioFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btLivro_Click()

    ' Fecha o formulário de Cadastro de Livros completamente
    Unload Me

    ' Abre o formulário de Cadastro de Funcionários
    FormularioCadastro.Show

End Sub

Private Sub btVendas_Click()
      ' Fecha o formulário de Vendas de Livros completamente
    Unload Me

    ' Abre o formulário de Vendas de Funcionários
    FormularioVendas.Show

End Sub
Private Sub btCadastrar_Click()

    Dim UltimaLinha As Long
    
    ' Validação básica dos campos
    If tbNome_Funcionario.Value = "" Or tbCPF.Value = "" Or tbEmail.Value = "" Then
        MsgBox "Por favor, preencha todos os campos antes de cadastrar.", vbExclamation
        Exit Sub
    End If
        
    ' Verificar se 'CPF' contém números inteiros válidos
    If Val(Replace(tbCPF.Value, ".", "")) <= 0 Then
        MsgBox "'CPF' deve conter apenas números.", vbExclamation
        Exit Sub
    End If
    
     ' Verifica o valor atual de CI na célula O1 e incrementa para o próximo número
    NovoCI = Range("O1").Value + 1
    Range("O1").Value = NovoCI  ' Atualiza a célula O1 com o próximo número
    
    ' Atribui um número de CI (Cadastro Interno) a partir da célula "O1"
    Me.tbCI.Value = Range("O1").Value

    ' Encontra a última linha vazia na planilha para adicionar o novo registro
    UltimaLinha = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    ' Registra os dados na planilha
    Cells(UltimaLinha, 1).Value = tbCI.Value
    Cells(UltimaLinha, 2).Value = tbNome_Funcionario.Value
    Cells(UltimaLinha, 3).Value = tbCPF.Value
    Cells(UltimaLinha, 4).Value = tbEmail.Value
  
    ' Limpar campos após o cadastro
    tbCI.Value = ""
    tbNome_Funcionario.Value = ""
    tbCPF.Value = ""
    tbEmail.Value = ""
        
    MsgBox "Cadastro realizado com sucesso!", vbInformation
End Sub

Private Sub btLimparDados_Click()

    ' Limpa os TextBoxes
    tbCI.Value = ""
    tbNome_Funcionario.Value = ""
    tbCPF.Value = ""
    tbEmail.Value = ""

    ' Limpa a TextBox de pesquisa
    tbPesquisa.Value = ""

    ' Limpa o ListBox de resultados
    lbResultados.Clear

End Sub

Private Sub btAlterar_Click()
    
    ' Verifica se um item está selecionado no ListBox
    If lbResultados.ListIndex = 0 Then
        MsgBox "Por favor, selecione um item para alterar.", vbExclamation
        Exit Sub
    End If
    
    ' Define a planilha onde os dados estão armazenados
    Set ws = ThisWorkbook.Sheets("Funcionario")

    ' Obtém o índice do item selecionado
    selectedIndex = lbResultados.ListIndex

    ' Obtém a linha correspondente ao item selecionado
    selectedRow = selectedIndex + 1 '

    ' Verifica se a linha selecionada é válida
    If selectedRow < 2 Or selectedRow > ws.Cells(ws.Rows.Count, 1).End(xlUp).Row Then
        MsgBox "Seleção inválida. Por favor, tente novamente.", vbExclamation
        Exit Sub
    End If

    ' Atualiza os valores na planilha com base nos TextBoxes
    ws.Cells(selectedRow, 1).Value = tbCI.Value
    ws.Cells(selectedRow, 2).Value = tbNome_Funcionario.Value
    ws.Cells(selectedRow, 3).Value = tbCPF.Value
    ws.Cells(selectedRow, 4).Value = tbEmail.Value

    ' Exibe uma mensagem confirmando a alteração
    MsgBox "Registro alterado com sucesso.", vbInformation

    ' Atualiza o ListBox após a alteração
    Call btPesquisar_Click

End Sub
Private Sub btExcluir_Click()

    Dim ws As Worksheet
    Dim selectedIndex As Long
    Dim selectedRow As Long
    Dim lastRow As Long

    ' Verifica se um item está selecionado
    If lbResultados.ListIndex = -1 Then
        MsgBox "Selecione um item da lista para excluir."
        Exit Sub
    End If

    ' Adiciona uma confirmação antes de excluir
    If MsgBox("Tem certeza de que deseja excluir este registro?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    ' Define a planilha onde a exclusão será feita
    Set ws = ThisWorkbook.Sheets("Funcionario")

    ' Obtém o índice do item selecionado
    selectedIndex = lbResultados.ListIndex

    ' Obtém o número da última linha com dados na planilha
    'Comentar bem
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' A ListBox começa na linha 2
    ' O índice da ListBox pode estar ajustado para começar a partir da linha 2 da planilha
    'Comentar bem
    selectedRow = selectedIndex + 1

    ' Verifica se a linha a ser excluída está dentro do intervalo válido
    If selectedRow < 2 Or selectedRow > lastRow Then
        MsgBox "Não foi possível localizar a linha para exclusão. Linha inválida: " & selectedRow
        Exit Sub
    End If

    ' Exclui a linha da planilha
    ws.Rows(selectedRow).Delete

    ' Atualiza a ListBox e as caixas de texto após a exclusão
    Call btPesquisar_Click

    ' Limpa as caixas de texto
    tbCI.Value = ""
    tbNome_Funcionario.Value = ""
    tbCPF.Value = ""
    tbEmail.Value = ""
    
End Sub


Private Sub btPesquisar_Click()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim palavra As String
    Dim found As Boolean

    ' Define a planilha onde a pesquisa será feita
    Set ws = ThisWorkbook.Sheets("Funcionario")

    ' Obtém a palavra digitada na TextBox
    palavra = Trim(tbPesquisa.Value) ' Remove espaços antes e depois

    ' Limpa a ListBox de resultados anteriores
    lbResultados.Clear

    ' Configura o número de colunas no ListBox e a largura de cada coluna
    lbResultados.ColumnCount = 8
    lbResultados.ColumnWidths = "100;0;150;80;100;100;80;50" 'Ajuste as larguras das colunas na listbox de pesquisa

    ' Adiciona os cabeçalhos das colunas ao ListBox
    lbResultados.AddItem "CI" ' Cabeçalhos das colunas
    lbResultados.List(0, 2) = "Nome do Funcionario"
    lbResultados.List(0, 3) = "CPF"
    lbResultados.List(0, 4) = "Email"

    ' Define o intervalo da pesquisa como o intervalo utilizado da planilha
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Encontrar a última linha com dados

    found = False ' Variável para verificar se algum resultado foi encontrado

    ' Percorre todas as linhas da planilha
    For i = 2 To lastRow ' Começa na linha 2
        ' Verifica se o termo pesquisado está presente em qualquer coluna da linha
        If InStr(1, ws.Cells(i, 1).Value, palavra, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 2).Value, palavra, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 3).Value, palavra, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 4).Value, palavra, vbTextCompare) > 0 Then

           ' Adiciona os dados da linha correspondente ao ListBox
           lbResultados.AddItem ws.Cells(i, 1).Value ' CI
           lbResultados.List(lbResultados.ListCount - 1, 2) = ws.Cells(i, 2).Value ' Nome do Funcionario
           lbResultados.List(lbResultados.ListCount - 1, 3) = ws.Cells(i, 3).Value ' CPF
           lbResultados.List(lbResultados.ListCount - 1, 4) = ws.Cells(i, 4).Value ' Email

           found = True ' Marca que um resultado foi encontrado
        End If
    Next i

    ' Se nenhum registro foi encontrado, exibe uma mensagem na ListBox
    If Not found Then
        lbResultados.AddItem "Registro não encontrado."
        lbResultados.List(0, 1) = "" ' Limpa o restante das colunas na linha de mensagem
        lbResultados.List(0, 2) = ""
        lbResultados.List(0, 3) = ""
        lbResultados.List(0, 4) = ""
    End If

End Sub


Private Sub Label12_Click()

End Sub

Private Sub Label25_Click()

End Sub

Private Sub Label32_Click()

End Sub

Private Sub Label8_Click()

End Sub
Private Sub lbResultados_Click()
    Dim ws As Worksheet
    Dim selectedRow As Range
    Dim selectedID As Long
    
     ' Verifica se um item foi selecionado e se não foi o cabeçalho (linha 0)
    If lbResultados.ListIndex = -1 Or lbResultados.ListIndex = 0 Then
        Exit Sub ' Sai se não tiver selecionado nada ou se o cabeçalho foi clicado
    End If

    ' Define a planilha onde a pesquisa foi feita
    Set ws = ThisWorkbook.Sheets("Funcionario")

    ' Obtém o ID do item selecionado (primeira coluna oculta)
    selectedID = lbResultados.Column(0, lbResultados.ListIndex)

    ' Localiza a linha na planilha correspondente ao ID selecionado
    Set selectedRow = ws.Columns(1).Find(What:=selectedID, LookIn:=xlValues, LookAt:=xlWhole)

    ' Se a linha for encontrada, preenche os TextBoxes
    If Not selectedRow Is Nothing Then
        tbCI.Value = selectedRow.Cells(1, 1).Value
        tbNome_Funcionario.Value = selectedRow.Cells(1, 2).Value
        tbCPF.Value = selectedRow.Cells(1, 3).Value
        tbEmail.Value = selectedRow.Cells(1, 4).Value
    Else
        MsgBox "Registro não encontrado.", vbExclamation
    End If
End Sub

Private Sub tbCPF_Change()
    Dim cpf As String
    Dim i As Integer
    
    ' Captura o texto digitado no campo CPF
    cpf = tbCPF.Value
    
    ' Remove tudo que não for número
    cpf = Replace(cpf, ".", "")
    cpf = Replace(cpf, "-", "")
    
    ' Verifica se o número digitado tem mais de 11 dígitos
    If Len(cpf) > 11 Then
        cpf = Left(cpf, 11)
    End If
    
    ' Aplica a máscara
    If Len(cpf) <= 3 Then
        tbCPF.Value = cpf
    ElseIf Len(cpf) <= 6 Then
        tbCPF.Value = Mid(cpf, 1, 3) & "." & Mid(cpf, 4, 3)
    ElseIf Len(cpf) <= 9 Then
        tbCPF.Value = Mid(cpf, 1, 3) & "." & Mid(cpf, 4, 3) & "." & Mid(cpf, 7, 3)
    Else
        tbCPF.Value = Mid(cpf, 1, 3) & "." & Mid(cpf, 4, 3) & "." & Mid(cpf, 7, 3) & "-" & Mid(cpf, 10, 2)
    End If
End Sub

Private Sub UserForm_Click()

End Sub
