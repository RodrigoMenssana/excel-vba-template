Attribute VB_Name = "ModProdutos"
Option Compare Text 'Não vai comparar Minúscula e Maiúscula

'Salvar Produtos
Sub SalvarProdutos()
    'Como será utilizado muitas vezes o Formulário, utilizamos o "With"
    With FrmProdutos
    
        'Cor Original da Borda do Text Box
        .TxtDescricao.BorderColor = &H80000005
        .TxtPreco.BorderColor = &H80000005
        
        'Validações de campos
            If .TxtDescricao.Text = Empty Then
                'Campo muda Cor da Borda
                .TxtDescricao.BorderColor = &H80FF&
                
                MsgBox "Digite o campo descrição!", vbExclamation, "Cadastro de Produtos"
                
                'Campo recebe o Foco
                .TxtDescricao.SetFocus
                
                'Sair da Rotina (Sub)
                Exit Sub
            End If
            
            If .TxtPreco.Text = Empty Then
                'Campo muda Cor da Borda
                .TxtPreco.BorderColor = &H80FF&
                
                MsgBox "Digite o Preço!", vbExclamation, "Cadastro de Produtos"
                
                'Campo recebe o Foco
                .TxtPreco.SetFocus
                
                'Sair da Rotina (Sub)
                Exit Sub
            End If
            
            'Cadastro do ID
                'Definindo variável
                    Dim lin As Long
                    
            'VERIFICAÇÃO SE É PARA SALVAR OU EDITAR
                If .BtnSalvar.Caption = "  Alterar" Then
                    'Procurar o nº da linha correspondente ao nº do Id pela correspondencia exata "LookAt:=xlWhole"
                    lin = PlanProdutos.Range("A:A").Find(.LstDados.Column(0), LookAt:=xlWhole).Row
                Else
                'Procurar Próxima Célula vazia
                lin = PlanProdutos.Range("A:A").Find(Empty).Row
                
                End If
                
                'Atribuir valor para ID
                If lin = 2 Then
                    'Colocar Valor em uma determinada Celula
                    PlanProdutos.Cells(lin, "A").Value = 1
                Else
                    'Pegar valor da linha anterior e somar com mais 1
                    PlanProdutos.Cells(lin, "A").Value = (PlanProdutos.Cells(lin - 1, "A").Value) + 1
                End If
                
                
                PlanProdutos.Cells(lin, "B").Value = .TxtDescricao.Text
                
                'Usando a função CDbl(Double) para converter para um número
                PlanProdutos.Cells(lin, "C").Value = VBA.CDbl(.TxtPreco.Text)
            
    End With
    
    'Chamar Função de novo Produto
    Call NovoProduto
    
End Sub

'BUSCAR DADOS PARA LIST BOX E PARA FILTROS NO TEXT BOX
Sub BuscarProduto()

    'Definindo variável
    Dim lin As Long
    Dim ultimaLinha As Long
    Dim i As Long
    
    'Inicializar variável i com valor zero
    i = 0
    
    'Limpar List Box
    FrmProdutos.LstDados.Clear
    
                  'O final da linha de dados
    ultimaLinha = PlanProdutos.UsedRange.Rows.Count
    
    'Estrutura de Repetição FOR
    For lin = 2 To ultimaLinha
    
        'Buscar termo digitado na Text Box corresponde aos dados da Planilha, e chamando a função AcSQL para não utilizar acentuação
        If PlanProdutos.Cells(lin, "B").Text Like "*" & AcSQL(FrmProdutos.TxtPesquisa.Text) & "*" Then
        
            FrmProdutos.LstDados.AddItem
            
            'Pegar dados da Primeira Coluna
            FrmProdutos.LstDados.List(i, 0) = PlanProdutos.Cells(lin, "A").Text
            
            'Inserir Hífen
            FrmProdutos.LstDados.List(i, 1) = "-"
            
            'Pegar dados da Terceira Coluna
            FrmProdutos.LstDados.List(i, 2) = PlanProdutos.Cells(lin, "B").Text
            
            'Pegar dados da Quarta Coluna
            FrmProdutos.LstDados.List(i, 3) = PlanProdutos.Cells(lin, "C").Text
            
            'Ir para próxima Linha da Planilha
            i = i + 1
            
        End If
    Next
    
End Sub

'CARREGAMENTO DE DADOS DO LIST BOX
Sub EditarProduto()

    'Como será utilizado muitas vezes o Formulário, utilizamos o "With"
    With FrmProdutos
        
        'Carregando Controles com dados do List Box
        .TxtDescricao.Text = .LstDados.Column(2) '3ª Coluna do List Box
        .TxtPreco.Text = .LstDados.Column(3) '4ª Coluna do List Box
        
        'Modificar Texto do Botão Salvar
        .BtnSalvar.Caption = "  Alterar"
        
        'Modificar Imagem do Botão Salvar
        .BtnSalvar.Picture = .PicEditar.Picture
        
        'Desabilitar List Box
        .LstDados.Enabled = False
        
        'Campo recebe o Foco
        .TxtDescricao.SetFocus
        
        
    
    End With
    
End Sub

'FUNÇÃO PARA NOVO PRODUTO
Sub NovoProduto()

'Como será utilizado muitas vezes o Formulário, utilizamos o "With"
    With FrmProdutos
        'Limpar Text Boxes
        .TxtDescricao.Text = Empty
        .TxtPesquisa.Text = Empty
        .TxtPreco.Text = Empty
        
        'Cor Original da Borda do Text Box
        .TxtDescricao.BorderColor = &H80000006
        .TxtPreco.BorderColor = &H80000006
        
        'Ativar List Box
        .LstDados.Enabled = True
        
        'Mudar configurações do botão Salvar/Editar
        .BtnSalvar.Caption = "  Salvar"
        .BtnSalvar.Picture = .PicSalvar.Picture
        
        'Campo recebe o Foco
        .TxtDescricao.SetFocus
    
    
    End With


End Sub
