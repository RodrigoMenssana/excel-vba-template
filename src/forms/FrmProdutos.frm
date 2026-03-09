VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmProdutos 
   Caption         =   "UserForm1"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8115
   OleObjectBlob   =   "FrmProdutos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BOTÃO SALVAR
Private Sub BtnSalvar_Click()

'Chamar Função do Módulo
    Call ModProdutos.SalvarProdutos
    
'Chamar
    Call ModProdutos.BuscarProduto

End Sub

'AO CLICAR NA LABEL O TEXT BOX RECEBE O FOCO
Private Sub LblPesquisa_Click()

    TxtPesquisa.SetFocus
    
End Sub


'DUPLO CLIQUE LIST BOX
Private Sub LstDados_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Call ModProdutos.EditarProduto

End Sub

Private Sub TxtPesquisa_Change()

    'Mostrar e esconder a Label de Pesquisa
    If TxtPesquisa.Text <> Empty Then
        LblPesquisa.Visible = False
    Else
        LblPesquisa.Visible = True
    End If
    
    'Chamar a Função Buscar Produtos
    Call ModProdutos.BuscarProduto

End Sub

'INICIALIZAR FORMULÁRIO
Private Sub UserForm_Initialize()

    'Chamar função de carregamento de dados da planilha
    Call ModProdutos.BuscarProduto

End Sub
