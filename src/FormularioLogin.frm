VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormularioLogin 
   Caption         =   "Login"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5580
   OleObjectBlob   =   "FormularioLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormularioLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btLogin_Click()
    ' Verifica se o campo de usuário está vazio
    If tbUsuario.Value = "" Then
        MsgBox "Digite o usuário.", , "Digite o Usuário!"
        Exit Sub
    End If

    ' Verifica se o campo de senha está vazio
    If tbSenha.Value = "" Then
        MsgBox "Digite a senha.", , "Digite a Senha!"
        Exit Sub
    End If

    ' Chama a sub Logar passando usuário e senha
    Call Logar(tbUsuario.Value, tbSenha.Value)
End Sub


Private Sub btSair_Click()
'Essa Sub é atribuida ao botao de Sair, quando o usuario clica ela executa o comando de Quit.

    Application.Quit

End Sub

Private Sub tbSenha_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Essa sub devolve o valor da senha com uma mascara, no caso se tratando de uma senha o valor será "*".
'Dentro do boxsenha, o valor se mascara.
    tbSenha.PasswordChar = "*"

End Sub

Private Sub tbSenha_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    'Nessa sub, ele chama um comando quando o usuário passa o mouse por cima da senha, acaba limpando o valor de "*" e_
    'o valor que o usário digitou.
    tbSenha.PasswordChar = ""
    
End Sub

Private Sub tbSenha_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    'Novamente, quando sairmos fora da caixa da senha, o valor retorna a ficar com a mascara.
    tbSenha.PasswordChar = "*"
    
End Sub

Private Sub tbUsuario_Change()
    'Essa função é simples, questão de estilização, enquanto o usuário não clicar na box do usuário, a palavra "Usuário" ficara_
    'visivel, quando ele clicar para digitar, essa palavra some "".
    
    If Len(Me.tbUsuario.Value) = 0 Then
        Me.lbUsuario.Caption = "Usuário"
        
    Else
        Me.lbUsuario.Caption = ""
        
    End If
       
End Sub

Private Sub tbSenha_Change()
    'Essa função também é bem simples, enquanto o usuário não clicar na box da senha, a palavra "Senha" ficara visivel_
    'quando ele clicar para digitar, essa palavra some "".
    
    If Len(Me.tbSenha.Value) = 0 Then
        Me.lbSenha.Caption = "Senha"
        
    Else
        Me.lbSenha.Caption = ""
        
    End If
       
End Sub

Private Sub Logar(Usuario As String, Senha As String)
    Dim linha As Integer
    Dim encontrado As Boolean

    ' Começa a partir da linha 2, pois a linha 1 é o cabeçalho
    linha = 2
    encontrado = False ' Variável de controle de login

    ' Loop para verificar se o usuário e senha existem na planilha
    Do Until Planilha2.Cells(linha, 1).Value = "" ' Verifica até o final dos dados
        ' Se encontrar o usuário e senha na planilha, faz o login
        If Planilha2.Cells(linha, 2).Value = Usuario And Planilha2.Cells(linha, 3).Value = Senha Then
            MsgBox "Logado com sucesso", vbInformation

    ' Fecha completamente o formulário de login (ou qualquer outro formulário)
    Unload Me
    
    ' Exibe o formulário de cadastro
    FormularioCadastro.Show vbModal

            

            encontrado = True
            Exit Sub ' Sai da sub já que o login foi realizado com sucesso
        End If
        linha = linha + 1 ' Avança para a próxima linha
    Loop

    ' Se o login não for encontrado, exibe uma mensagem
    If Not encontrado Then
        MsgBox "Usuário ou senha inválidos.", vbExclamation
    End If
End Sub

Private Sub UserForm1_Initialize()

    'Essa é a sub de inicialização do formulário
    
    'Esse comando deixa invisivel a aplicação, ou seja, as planilhas de dados.
    'Application.Visible = False

End Sub


'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Essa sub caso o usuario clique no x do formulario de entrada, toda a planilha fecha.

'Se caso for 0, que o ponto de fechamento falado no comentário acima, ele salva o formulário e fecha.
'If CloseMode = 0 Then
    
    'Comandos de salvar e fechar.
    'ThisWorkbook.Save
    'ThisWorkbook.Close
    
'End If
      
'End Sub


Private Sub UserForm_Click()

End Sub
