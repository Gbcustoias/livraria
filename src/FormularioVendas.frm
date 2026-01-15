VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormularioVendas 
   Caption         =   "Vendas"
   ClientHeight    =   10755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17805
   OleObjectBlob   =   "FormularioVendas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormularioVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btFuncionario_Click()
      ' Fecha o formulário  de Funcionario completamente
    Unload Me

    ' Abre o formulário de Funcionários
    FormularioFuncionario.Show

End Sub


Private Sub btLivro_Click()

    ' Fecha o formulário de Cadastro de Livros completamente
    Unload Me

    ' Abre o formulário de Cadastro de Funcionários
    FormularioCadastro.Show

End Sub

Private Sub btFinalizar_Click()

    Dim ws As Worksheet
    Dim wsEstoque As Worksheet
    Dim lastRow As Long
    Dim i As Integer
    Dim produto As String
    Dim quantidade As Integer
    Dim preco As Double
    Dim total As Double
    Dim dataVenda As String
    Dim totalCarrinho As Double
    Dim nfForm As Object ' Referência para o UserForm NF
    Dim produtoLinha As Variant ' Variável como Variant para evitar erro de tipo
    Dim estoqueProduto As Double ' Estoque disponível do produto
    
    ' Verifica se todos os campos obrigatórios estão preenchidos
    If tbCPF.Value = "" Or tbCliente.Value = "" Or tbEmail.Value = "" Or tbTelefone.Value = "" Then
        MsgBox "Preencha todos os campos obrigatórios!", vbExclamation
        Exit Sub
    End If
    
    ' Valida o email
    If Not IsEmailValid(tbEmail.Value) Then
        MsgBox "Por favor, insira um email válido.", vbExclamation
        Exit Sub
    End If
    
    ' Verifica se o carrinho está vazio
    If lbCarrinho.ListCount = 0 Then
        MsgBox "Adicione produtos ao carrinho antes de finalizar a venda.", vbExclamation
        Exit Sub
    End If
    
    ' Acessa a planilha "Vendas"
    Set ws = ThisWorkbook.Sheets("Vendas")
    ' Acessa a planilha "Banco de Dados" para atualizar o estoque
    Set wsEstoque = ThisWorkbook.Sheets("Banco de Dados")
    
    ' Encontra a última linha disponível para adicionar a venda
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Salva as informações do cliente
    ws.Cells(lastRow, 1).Value = tbCPF.Value
    ws.Cells(lastRow, 2).Value = tbCliente.Value
    ws.Cells(lastRow, 3).Value = tbEmail.Value
    ws.Cells(lastRow, 4).Value = tbTelefone.Value
    ws.Cells(lastRow, 5).Value = cbFuncionario.Value
    
    ' A data da venda
    dataVenda = Date ' Data de hoje
    
    ' Inicializa o total do carrinho
    totalCarrinho = 0
    
    ' Salva os produtos e quantidades no carrinho
    For i = 0 To lbCarrinho.ListCount - 1
        produto = Split(lbCarrinho.List(i), " | ")(0) ' Nome do produto
        quantidade = CInt(Split(Split(lbCarrinho.List(i), " | ")(1), " un.")(0)) ' Quantidade
        preco = CDbl(Split(Split(lbCarrinho.List(i), " | ")(2), ": R$")(1)) ' Preço unitário
        total = preco * quantidade ' Total do item
        
        ' Salva na planilha "Vendas", linha por linha
        ws.Cells(lastRow, 6).Value = produto
        ws.Cells(lastRow, 7).Value = quantidade
        ws.Cells(lastRow, 8).Value = preco
        ws.Cells(lastRow, 9).Value = total
        ws.Cells(lastRow, 10).Value = dataVenda ' Data da venda
        
        totalCarrinho = totalCarrinho + total ' Acumulando o total da compra
        
        ' Atualiza o banco de dados subtraindo as unidades vendidas
        produtoLinha = Application.Match(produto, wsEstoque.Range("B:B"), 0) ' Coluna B: Nome do Livro
        
        ' Se o produto foi encontrado no estoque, atualiza o estoque
        If Not IsError(produtoLinha) Then
            ' Obtém o estoque atual do produto (coluna 8: Unidades)
            estoqueProduto = Val(Trim(wsEstoque.Cells(produtoLinha, 8).Value)) ' Usar Val e Trim para garantir que o valor seja numérico
            
            ' Subtrai a quantidade vendida do estoque
            wsEstoque.Cells(produtoLinha, 8).Value = estoqueProduto - quantidade
        End If
        
        lastRow = lastRow + 1 ' Para a próxima linha de venda
    Next i
    
    ' Abre o UserForm NF e preenche com os dados da venda
    Set nfForm = NF ' Referencia o UserForm NF
    
    ' Preenche os campos com os dados da venda
    nfForm.lbDia.Caption = Day(dataVenda)
    nfForm.lbMes.Caption = MonthName(Month(dataVenda))
    nfForm.lbAno.Caption = Year(dataVenda)
    nfForm.lbCliente.Caption = tbCliente.Value
    nfForm.lbCPF.Caption = tbCPF.Value
    
    ' Preenche os campos de produtos no UserForm NF
    For i = 0 To lbCarrinho.ListCount - 1
        produto = Split(lbCarrinho.List(i), " | ")(0)
        quantidade = CInt(Split(Split(lbCarrinho.List(i), " | ")(1), " un.")(0))
        preco = CDbl(Split(Split(lbCarrinho.List(i), " | ")(2), ": R$")(1))
        total = preco * quantidade
        
        ' Preenche os campos de Quantidade, Livro, Preço e Total
        nfForm.Controls("lbQuant" & (i + 1)).Caption = quantidade
        nfForm.Controls("lbLivro" & (i + 1)).Caption = produto
        nfForm.Controls("lbPreço" & (i + 1)).Caption = Format(preco, "R$ 0.00")
        nfForm.Controls("lbTotal" & (i + 1)).Caption = Format(total, "R$ 0.00")
    Next i
    
    ' Preenche o valor total da compra no campo lbValor
    nfForm.lbValor.Caption = "R$ " & Format(totalCarrinho, "0.00")
    
    ' Exibe o UserForm NF
    nfForm.Show
    
    MsgBox "Venda finalizada com sucesso!", vbInformation
    ' Limpar os campos após finalizar
    btLimparDados_Click
End Sub

Private Sub btLimparDados_Click()

    ' Limpar todos os campos
    tbCPF.Value = ""
    tbCliente.Value = ""
    tbEmail.Value = ""
    tbTelefone.Value = ""
    cbFuncionario.Value = ""
    tbPesquisa.Value = ""
    tbUnidades.Value = ""
    lbResultados.Clear
    lbCarrinho.Clear
    tbTotal.Value = ""
    
End Sub


Private Sub btPesquisar_Click()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim palavra As String
    Dim found As Boolean

    ' Define a planilha onde a pesquisa será feita
    Set ws = ThisWorkbook.Sheets("Banco de Dados")

    ' Obtém a palavra digitada na TextBox
    palavra = Trim(tbPesquisa.Value) ' Remove espaços antes e depois

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

    ' Percorre todas as linhas da planilha
    For i = 2 To lastRow ' Começa na linha 2
        ' Verifica se o termo pesquisado está presente em qualquer coluna da linha
        If InStr(1, ws.Cells(i, 1).Value, palavra, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 2).Value, palavra, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 3).Value, palavra, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 4).Value, palavra, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 5).Value, palavra, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 6).Value, palavra, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, 7).Value, palavra, vbTextCompare) > 0 Then

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


Private Sub btRemover_Click()
    ' Verifica se há um item selecionado no carrinho
    If lbCarrinho.ListIndex = -1 Then
        MsgBox "Selecione um item para remover do carrinho.", vbExclamation
        Exit Sub
    End If
    
    ' Remove o item selecionado no ListBox
    lbCarrinho.RemoveItem lbCarrinho.ListIndex
    
    ' Recalcula o total do carrinho após remoção
    Dim totalCarrinho As Double
    totalCarrinho = 0 ' Resetando o total
    
    Dim i As Integer
    Dim valorTexto As String
    Dim valorItem As Double
    Dim valorFormatado As String
    
    ' Percorre os itens restantes no carrinho e calcula o total
    For i = 0 To lbCarrinho.ListCount - 1
        ' Extrai o valor do item no formato "Total: R$xx.xx"
        valorTexto = Split(lbCarrinho.List(i), "Total: R$")(1)
        valorFormatado = Trim(valorTexto) ' Remove espaços extras
        
        ' Verifica se o valor extraído é numérico
        If IsNumeric(valorFormatado) Then
            valorItem = CDbl(valorFormatado)
            totalCarrinho = totalCarrinho + valorItem
        End If
    Next i
    
    ' Exibe o novo total no campo tbTotal
    tbTotal.Value = "R$ " & Format(totalCarrinho, "0.00")
    
    MsgBox "Item removido com sucesso!", vbInformation
End Sub


Private Sub btSelecionar_Click()
    Dim selectedRow As Integer
    Dim nomeLivro As String
    Dim precoUnitario As Double
    Dim quantidade As Integer
    Dim precoTotal As Double
    Dim itemCarrinho As String
    Dim totalCarrinho As Double
    Dim i As Integer
    Dim valorTexto As String
    Dim valorItem As Double
    Dim valorFormatado As String
    Dim estoqueDisponivel As Integer
    
    ' 1. Verifique se há uma seleção em lbResultados
    If lbResultados.ListIndex = -1 Then
        MsgBox "Selecione um livro na lista de resultados.", vbExclamation
        Exit Sub
    End If

    ' 2. Obtenha dados do livro selecionado
    selectedRow = lbResultados.ListIndex
    nomeLivro = lbResultados.List(selectedRow, 1)  ' Nome do Livro
    
    ' Verifica se o preço unitário é numérico
    If IsNumeric(lbResultados.List(selectedRow, 6)) Then ' Coluna 6 corresponde ao preço
        precoUnitario = CDbl(lbResultados.List(selectedRow, 6))
    Else
        MsgBox "O preço do livro selecionado não é válido.", vbExclamation
        Exit Sub
    End If

    ' 3. Verifique a quantidade informada e calcule o total
    If IsNumeric(tbUnidades.Text) Then
        quantidade = CInt(tbUnidades.Text)
        If quantidade <= 0 Then
            MsgBox "Informe uma quantidade válida.", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "Informe uma quantidade válida.", vbExclamation
        Exit Sub
    End If

    ' 4. Verifique o estoque disponível na planilha "Banco de Dados"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Banco de Dados")
    
    ' Encontra a linha correspondente ao livro selecionado
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Procura o livro selecionado na planilha "Banco de Dados"
    Dim linhaLivro As Long
    linhaLivro = -1
    For i = 2 To lastRow
        If ws.Cells(i, 2).Value = nomeLivro Then
            linhaLivro = i
            Exit For
        End If
    Next i
    
    ' Verifica se o livro foi encontrado
    If linhaLivro = -1 Then
        MsgBox "Livro não encontrado no banco de dados.", vbExclamation
        Exit Sub
    End If
    
    ' Verifica o estoque disponível do livro
    estoqueDisponivel = ws.Cells(linhaLivro, 8).Value ' Coluna 8 = Estoque
    
    ' 5. Verifica se a quantidade informada não excede o estoque
    If quantidade > estoqueDisponivel Then
        MsgBox "Quantidade solicitada excede o estoque disponível (" & estoqueDisponivel & " unidades).", vbExclamation
        Exit Sub
    End If

    ' 6. Calcule o preço total
    precoTotal = precoUnitario * quantidade

    ' 7. Adicione ao lbCarrinho
    itemCarrinho = nomeLivro & " | " & quantidade & " un. | Preço Unitário: R$" & Format(precoUnitario, "0.00") & " | Total: R$" & Format(precoTotal, "0.00")
    lbCarrinho.AddItem itemCarrinho

    ' 8. Recalcule o total do carrinho
    totalCarrinho = 0
    For i = 0 To lbCarrinho.ListCount - 1
        ' Extrai o valor do item no formato "Total: R$xx.xx"
        valorTexto = Split(lbCarrinho.List(i), "Total: R$")(1)
        valorFormatado = Trim(valorTexto) ' Remove espaços extras
        
        ' Verifica se o valor extraído é numérico
        If IsNumeric(valorFormatado) Then
            valorItem = CDbl(valorFormatado)
            totalCarrinho = totalCarrinho + valorItem
        End If
    Next i

    ' 9. Exibe o total no tbTotal
    tbTotal.Value = "R$ " & Format(totalCarrinho, "0.00")

    ' Opcional: Limpar seleção e campos após adicionar ao carrinho
    lbResultados.ListIndex = -1
    tbUnidades.Text = ""
End Sub


Private Sub Label45_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Define a planilha onde os dados dos funcionários estão
    Set ws = ThisWorkbook.Sheets("Funcionario")

    ' Encontra a última linha com dados na coluna B (onde estão os nomes dos funcionários)
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    ' Limpa qualquer valor pré-existente no ComboBox
    cbFuncionario.Clear

    ' Preenche o ComboBox com os nomes dos funcionários na coluna B (da linha 2 até a última linha)
    For i = 2 To lastRow
        cbFuncionario.AddItem ws.Cells(i, 2).Value
    Next i
End Sub

Private Sub lbRemover_Click()
    Dim selectedIndex As Integer
    Dim totalCarrinho As Double
    Dim i As Integer
    Dim valorTexto As String
    Dim valorItem As Double
    Dim valorFormatado As String
    
    ' Verifica se algum item está selecionado no carrinho
    selectedIndex = lbCarrinho.ListIndex
    If selectedIndex = -1 Then
        MsgBox "Selecione um item no carrinho para remover.", vbExclamation
        Exit Sub
    End If

    ' Remove o item selecionado do lbCarrinho
    lbCarrinho.RemoveItem selectedIndex

    ' Recalcula o total do carrinho após remoção
    totalCarrinho = 0
    For i = 0 To lbCarrinho.ListCount - 1
        ' Extrai o valor do item no formato "Total: R$xx.xx"
        valorTexto = Split(lbCarrinho.List(i), "Total: R$")(1)
        valorFormatado = Trim(valorTexto) ' Remove espaços extras
        
        ' Verifica se o valor extraído é numérico
        If IsNumeric(valorFormatado) Then
            valorItem = CDbl(valorFormatado)
            totalCarrinho = totalCarrinho + valorItem
        End If
    Next i

    ' Atualiza o valor total no tbTotal
    tbTotal.Value = "R$ " & Format(totalCarrinho, "0.00")
End Sub

Private Sub tbCliente_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
' Impede a digitação de números no campo Nome do Cliente
    If (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0 ' Cancela a tecla pressionada
        MsgBox "O campo 'Cliente' não pode conter números.", vbExclamation
    End If
End Sub
Private Sub tbCPF_Change()
    ' Limitar para 11 caracteres e aplicar a máscara
    If Len(tbCPF.Value) > 14 Then
        tbCPF.Value = Left(tbCPF.Value, 11)
    End If

    ' Aplica a máscara no CPF
    If Len(tbCPF.Value) = 3 Or Len(tbCPF.Value) = 7 Then
        tbCPF.Value = tbCPF.Value & "."
    ElseIf Len(tbCPF.Value) = 11 Then
        tbCPF.Value = tbCPF.Value & "-"
    End If
End Sub

Function IsEmailValid(email As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.IgnoreCase = True
    regex.Global = True
    regex.Pattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    
    IsEmailValid = regex.Test(email)
End Function
Private Sub tbTelefone_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
' Permite apenas números
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0 ' Cancela a tecla pressionada
        MsgBox "O campo 'Telefone' só pode conter números.", vbExclamation
    End If
End Sub

