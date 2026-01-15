VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormularioCadastro 
   Caption         =   "Cadastro de Livro"
   ClientHeight    =   8445.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14250
   OleObjectBlob   =   "FormularioCadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormularioCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btCadastrar_Click()
    Dim UltimaLinha As Long
    
    ' Validação básica dos campos
    If tbNome_Livro.Value = "" Or tbISBN.Value = "" Or tbAutoria.Value = "" Or tbEditora.Value = "" Or tbCategoria.Value = "" Or tbPreco.Value = "" Then
        MsgBox "Por favor, preencha todos os campos antes de cadastrar.", vbExclamation
        Exit Sub
    End If
        
    ' Verificar se 'ISBN' contém números inteiros válidos
    If Val(tbISBN.Value) <= 0 Then
        MsgBox "'ISBN' deve conter apenas números .", vbExclamation
        Exit Sub
    End If
    
    
    ' Verificar se 'Preço' é um número válido
    If Not IsNumeric(tbPreco.Value) Then
        MsgBox "'Preço' deve ser um valor numérico.", vbExclamation
        Exit Sub
    End If
    
    Me.tbID.Value = Range("O1").Value

    UltimaLinha = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    Cells(UltimaLinha, 1).Value = tbID.Value
    Cells(UltimaLinha, 2).Value = tbNome_Livro.Value
    Cells(UltimaLinha, 3).Value = tbISBN.Value
    Cells(UltimaLinha, 4).Value = tbAutoria.Value
    Cells(UltimaLinha, 5).Value = tbEditora.Value
    Cells(UltimaLinha, 6).Value = tbCategoria.Value
    Cells(UltimaLinha, 7).Value = tbPreco.Value
    Cells(UltimaLinha, 8).Value = tbUnidades.Value

    'Limpar campos após o cadastro
    tbID.Value = ""
    tbNome_Livro.Value = ""
    tbISBN.Value = ""
    tbAutoria.Value = ""
    tbEditora.Value = ""
    tbCategoria.Value = ""
    tbPreco.Value = ""
    tbUnidades.Value = ""
        
    MsgBox "Cadastro realizado com sucesso!", vbInformation
    
End Sub
Private Sub btLimparDados_Click()

    ' Limpa os TextBoxes
    tbID.Value = ""
    tbNome_Livro.Value = ""
    tbISBN.Value = ""
    tbAutoria.Value = ""
    tbEditora.Value = ""
    tbCategoria.Value = ""
    tbPreco.Value = ""
    tbUnidades.Value = ""

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
    Set ws = ThisWorkbook.Sheets("Banco de Dados")

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
    ws.Cells(selectedRow, 1).Value = tbID.Value
    ws.Cells(selectedRow, 2).Value = tbNome_Livro.Value
    ws.Cells(selectedRow, 3).Value = tbISBN.Value
    ws.Cells(selectedRow, 4).Value = tbAutoria.Value
    ws.Cells(selectedRow, 5).Value = tbEditora.Value
    ws.Cells(selectedRow, 6).Value = tbCategoria.Value
    ws.Cells(selectedRow, 7).Value = tbPreco.Value
     ws.Cells(selectedRow, 8).Value = tbUnidades.Value

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
    Set ws = ThisWorkbook.Sheets("Banco de Dados")

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
    tbID.Value = ""
    tbNome_Livro.Value = ""
    tbISBN.Value = ""
    tbAutoria.Value = ""
    tbEditora.Value = ""
    tbCategoria.Value = ""
    tbPreco.Value = ""
    
End Sub
Private Sub btPesquisar_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim palavra As String
    Dim found As Boolean

    ' Define a planilha onde a pesquisa será feita
    Set ws = ThisWorkbook.Sheets("Banco de Dados")

    ' Obtém a palavra digitada na TextBox (sem espaços extras)
    palavra = Trim(tbPesquisa.Value) ' Remove espaços antes e depois

    ' Se a palavra de pesquisa estiver vazia, mostra mensagem e sai da sub
    If palavra = "" Then
        MsgBox "Digite um termo para pesquisa.", vbExclamation
        Exit Sub
    End If

    ' Limpa a ListBox de resultados anteriores
    lbResultados.Clear

    ' Configura o número de colunas no ListBox e a largura de cada coluna
    lbResultados.ColumnCount = 8 ' Temos 8 colunas, incluindo unidades
    lbResultados.ColumnWidths = "50;150;100;100;100;100;80;60" ' Ajuste a largura para cada coluna conforme necessário

    ' Adiciona os cabeçalhos das colunas ao ListBox
    lbResultados.AddItem "ID" ' Cabeçalhos das colunas
    lbResultados.List(0, 1) = "Nome do Livro"
    lbResultados.List(0, 2) = "ISBN"
    lbResultados.List(0, 3) = "Autoria"
    lbResultados.List(0, 4) = "Editora"
    lbResultados.List(0, 5) = "Categoria"
    lbResultados.List(0, 6) = "Preço"
    lbResultados.List(0, 7) = "Unidades" ' Cabeçalho para unidades

    ' Define o intervalo da pesquisa como o intervalo utilizado da planilha
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Encontrar a última linha com dados

    found = False ' Variável para verificar se algum resultado foi encontrado

    ' Percorre todas as linhas da planilha (começa da linha 2, ignorando o cabeçalho)
    For i = 2 To lastRow ' Começa na linha 2
        ' Verifica se o termo pesquisado está presente em qualquer coluna da linha
        If InStr(1, ws.Cells(i, 1).Value, palavra, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 2).Value, palavra, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 3).Value, palavra, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 4).Value, palavra, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 5).Value, palavra, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 6).Value, palavra, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 7).Value, palavra, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 8).Value, palavra, vbTextCompare) > 0 Then

            ' Adiciona os dados da linha correspondente ao ListBox
            lbResultados.AddItem ws.Cells(i, 1).Value ' ID
            lbResultados.List(lbResultados.ListCount - 1, 1) = ws.Cells(i, 2).Value ' Nome do Livro
            lbResultados.List(lbResultados.ListCount - 1, 2) = ws.Cells(i, 3).Value ' ISBN
            lbResultados.List(lbResultados.ListCount - 1, 3) = ws.Cells(i, 4).Value ' Autoria
            lbResultados.List(lbResultados.ListCount - 1, 4) = ws.Cells(i, 5).Value ' Editora
            lbResultados.List(lbResultados.ListCount - 1, 5) = ws.Cells(i, 6).Value ' Categoria
            lbResultados.List(lbResultados.ListCount - 1, 6) = ws.Cells(i, 7).Value ' Preço
            lbResultados.List(lbResultados.ListCount - 1, 7) = ws.Cells(i, 8).Value ' Unidades

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
        lbResultados.List(0, 5) = ""
        lbResultados.List(0, 6) = ""
        lbResultados.List(0, 7) = ""
    End If

End Sub

Private Sub Label25_Click()

End Sub

Private Sub Label32_Click()

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
    Set ws = ThisWorkbook.Sheets("Banco de Dados")

    ' Obtém o ID do item selecionado (primeira coluna oculta)
    selectedID = lbResultados.Column(0, lbResultados.ListIndex)

    ' Localiza a linha na planilha correspondente ao ID selecionado
    Set selectedRow = ws.Columns(1).Find(What:=selectedID, LookIn:=xlValues, LookAt:=xlWhole)

    ' Se a linha for encontrada, preenche os TextBoxes
    If Not selectedRow Is Nothing Then
        tbID.Value = selectedRow.Cells(1, 1).Value
        tbNome_Livro.Value = selectedRow.Cells(1, 2).Value
        tbISBN.Value = selectedRow.Cells(1, 3).Value
        tbAutoria.Value = selectedRow.Cells(1, 4).Value
        tbEditora.Value = selectedRow.Cells(1, 5).Value
        tbCategoria.Value = selectedRow.Cells(1, 6).Value
        tbPreco.Value = selectedRow.Cells(1, 7).Value
        tbUnidades.Value = selectedRow.Cells(1, 8).Value
    Else
        MsgBox "Registro não encontrado.", vbExclamation
    End If
End Sub
Private Sub btFuncionario_Click()
      ' Fecha o formulário  de Funcionario completamente
    Unload Me

    ' Abre o formulário de Funcionários
    FormularioFuncionario.Show

End Sub
Private Sub btVendas_Click()
      ' Fecha o formulário de Vendas de Livros completamente
    Unload Me

    ' Abre o formulário de Vendas de Funcionários
    FormularioVendas.Show

End Sub
