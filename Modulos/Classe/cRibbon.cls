VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRibbon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub Inicio(frm As Form, Ribbon As ACPRibbon, iml As ImageList)
20    On Error Resume Next
30    Ribbon.Theme = 2 '0 cinza escuro, 1 azul, 2 cinza claro
40    With Ribbon
50        frm.Picture = .LoadBackground
         '# OPCIONAL - Carrega as cores do Background do formulario
60        frm.BackColor = .BackColor
          
         '# Configura a imagelist desejada
70       .ImageList = iml
         '# Configura os bot?es    (True = Centtro, False(Padr?p) = Alinhado ao Topo)
80       .ButtonCenter = False
'------Começa aqui--------------------------------------------------

         '# Add Tabs ---   ID - Caption
90       .AddTab "1", "Administrativo"
100      .AddTab "2", "Comercial"
101      .AddTab "3", "Adm Financeira"
102      .AddTab "4", "Operacional"
         .AddTab "5", "Suprimentos"
         .AddTab "6", "Supervisão"
         .AddTab "7", "Encerrar"
         
         'Add Categorias   ID - Tab - Caption - ShowDialogButton
110      .AddCat "1", "1", "Pessoa", False 'Administrativo
         .AddCat "2", "1", "Logística", False 'Administrativo
         .AddCat "3", "1", "ASO", False 'Administrativo
         .AddCat "4", "1", "Treinamento", False 'Administrativo
         
         
120      .AddCat "5", "2", "Parâmetros", False 'Comercial
         .AddCat "6", "2", "Negociação", False 'Comercial
         .AddCat "7", "2", "Proposta de Locação", False 'Comercial"
                    
         .AddCat "8", "3", "Parâmetros", False 'Adm. Financeira
         .AddCat "9", "3", "Mov. Financeiro", False 'Adm. Financeira
         .AddCat "10", "3", "Lanc. Financ", False 'Adm. Financeira
         .AddCat "11", "3", "Gerar Excel", False 'Adm. Financeira
         .AddCat "12", "3", "Custo", False 'Adm. Financeira
         .AddCat "13", "3", "Consulta", False 'Adm. Financeira
         
         .AddCat "14", "4", "Equipamentos", False 'Operacional
         .AddCat "15", "4", "Registro e Movimentação de equipamentos", False 'Operacional
 
         .AddCat "16", "5", "Parâmetros", False 'Suprimentos
         .AddCat "17", "5", "Supervisão", False 'Suprimentos
         .AddCat "18", "5", "Movimentação de Materiais", False 'Suprimentos
         .AddCat "19", "5", "Almoxarifado", False 'Suprimentos
         .AddCat "20", "5", "Consulta", False 'Suprimentos
         
         .AddCat "21", "6", "Supervisão", False 'Supervisão

         .AddCat "22", "7", "Sair", False 'Sair

        
         'Administrativo-----------------------------------------------------------
         '# Add Button    ID - Cat - Caption    - Icons -More Arrow - ToolTip
130      .AddButton "0", "1", "Cadastro", 31, False, "Cadastro de pessoa"
         .AddButton "1", "1", "Consulta", 24, False, "Consulta Resumo Cliente"
         .AddButton "2", "2", "Atualiza eventos", 14, False, "Registro e Atualização de eventos"
         .AddButton "3", "2", "Agenda", 52, False, "Registro e Atualização de Escala de pessoal"
         .AddButton "4", "2", "Consulta Escala", 24, False, "Consulta Escala de Funcionário"
         .AddButton "5", "3", "Cad. Exames", 51, False, "Registro e Atualização de Exames"
         .AddButton "6", "3", "Agenda", 52, False, "Agenda de Exames Médicos"
         .AddButton "7", "3", "Consulta", 24, False, "Consulta a Agenda de Exames Por Funcionário"
         .AddButton "8", "4", "Cad. treinamento", 24, False, "Registro e Atualização de treinamentos"
         .AddButton "9", "4", "Agenda", 52, False, "Registro e Atualização de Agenda de treinamento"
         .AddButton "10", "4", "Cad. treinamento", 20, False, "Registro e Atualização de treinamentos Funcionais"

         'Comercial-------------------------------------------------------------
140      .AddButton "21", "5", "Unid. Oper.", 2, False, "Registro e Atualização de Unidade Operacional"
         .AddButton "22", "5", "Contrato/Produtos", 9, False, "Registro de Contratos e Produtos"
         .AddButton "23", "5", "Atu. Preços", 32, False, "Atualiza Preços de Contratos e Produtos"
         .AddButton "24", "5", "Consulta Preços", 36, False, "Preços Praticados Por Contratos"
         .AddButton "25", "6", "Medição", 25, False, "Registro e Processamento de Medições"
         .AddButton "26", "6", "Mapa de Medição", 15, False, "Controle e Emissão de Mapa de Medição"
         .AddButton "27", "6", "Fatura", 17, False, "Emissão de Fatura de Locação"
         .AddButton "28", "6", "Consulta Medições", 19, False, "Consulta Analítica de Medições"
         .AddButton "29", "7", "Indeniz.de Equiptos", 19, False, "Indenização de equipamentos"
         .AddButton "30", "7", "Proposta", 21, False, "Criar proposta de Negociação"
         .AddButton "31", "7", "O.S.", 23, False, "Registro de Ordem de Serviço"
         
         'Financeiro-----------------------------------------------------------
         .AddButton "41", "8", "Prod Fornec", 10, False, "Cadastramento de Produtos e fornecedores"
         .AddButton "42", "8", "C. Custo", 10, False, "Registro e Atualização de Centro de Custo"
         .AddButton "43", "9", "Recebtos", 10, False, "Recebimento de Fatura de Clientes"
         .AddButton "44", "9", "Pagtos", 10, False, "Pagamentos de Contas de fornecedores e Serviços"
         .AddButton "45", "9", "Reprog", 10, False, "Reprogramação Financeira"
         .AddButton "46", "10", "N Fiscal", 10, False, "Lançamento de Notas Fiscais de fornecedores e Serviços"
         .AddButton "47", "10", "Reemb.", 10, False, "Lançamento de reembolso de Despesas"
         .AddButton "48", "10", "Recebto. Val.", 10, False, "Contas a Receber Fora Medições"
         .AddButton "49", "11", "Contab", 10, False, "Gerar Excel Para Contabilidade"
         .AddButton "50", "11", "Cta.Receber", 10, False, "Gerar Excel Para Contabilidade"
         .AddButton "51", "12", "C.Custo", 10, False, "Consulta a Custo Por Centro De Custo"
         .AddButton "52", "12", "Custo Excel", 10, False, "Gerar Custo Em Excel"
         .AddButton "53", "13", "Lanc Financ", 10, False, "Lançamentos Financeiros do Dia"
         .AddButton "54", "13", "Analítico", 10, False, "Consulta Finaceiro Analítico Por Semana"
         .AddButton "55", "13", "Consolid", 10, False, "Consulta Finaceiro Analítico Consolidado"
         .AddButton "56", "13", "Projeção", 10, False, "Projeção Financeira"
         .AddButton "57", "13", "Cta Pag Rec", 10, False, "Contas Pagas e Recebidas Até a Data"
         .AddButton "58", "13", "Pagar/Receber", 10, False, "Consulta a Contas a Pagar e a Receber"
         .AddButton "59", "13", "Pagtos Realiz.", 10, False, "Pagamentos Realizados no Período Informado"
         .AddButton "60", "13", "Faturamento", 10, False, "Consulta a Faturamento Por Período"
         .AddButton "61", "13", "Loc n/Faturado", 10, False, "Locações e Serviços Realizados e Não Faturados"

         'Operacional-------------------------------------------------------------
         .AddButton "70", "14", "Cad equipamentos", 10, False, "Registro e Atualização de equipamentos"
         .AddButton "71", "15", "Logística de equipamentos", 10, False, "Locações e Serviços Realizados e Não Faturados"
         
         'Suprimentos----------------------------------------------------------
         .AddButton "80", "16", "Inventário", 26, False, "Inventário de Material em Estoque"
         .AddButton "81", "16", "Unidade de Embalagem", 26, False, "Registro de Unidades de Embalagem"
         .AddButton "82", "16", "Unidade de Medida", 26, False, "Registro de Unidades de Medida"
         .AddButton "83", "17", "Produtos", 26, False, "Registro de Atualização de Produtos"
         .AddButton "84", "17", "Contrato", 26, False, "Registro de Atualização de Contratos Comerciais"
         .AddButton "85", "18", "Req. Materiais", 26, False, "Requisição de Materiais (Baixa de Estoque)"
         .AddButton "86", "18", "Req. Compras", 26, False, "Requisição de Compras. Reposição de Estoque"
         .AddButton "87", "18", "Ordem de Compra", 26, False, "Ordem de Compra de Materiais.(PO)"
         .AddButton "88", "19", "Recebimento de Materiais", 26, False, "Recebimento de Compras Efetuadas"
         .AddButton "89", "19", "Entrega de Requisição", 26, False, "Entrega de Material Requisitado.(Baixa de Estoque)"
         .AddButton "90", "20", "Posição de Estoque", 26, False, "Posição de Estoque Atualizada"
         .AddButton "91", "20", "Ordens de Compra", 26, False, "Posição Atual de Ordens de Compras"
         .AddButton "92", "20", "Requisições", 26, False, "Posição Atual de Requisições de Baixa de Estoque"

         'Supervisão-------------------------------------------------------------
         .AddButton "100", "21", "Abertura e Encerramento do Sistema", 10, False, "Abertura e Encerramento do Sistema"
         .AddButton "101", "21", "Usuário/Senha", 10, False, "Registro e Atualização de Usuário/Senha"
         .AddButton "102", "21", "Endereços", 10, False, "Registro e Atualização de Endereços"
         .AddButton "103", "21", "Parâmetros Especiais", 10, False, "Parâmetros Especiais"
         .AddButton "104", "21", "Valores por Trimestre", 10, False, "Valores Pagos Por Trimestre"
         
         'SairConfigur??es--------------------------------------------------------
         .AddButton "200", "22", "Sair", 26, False, "Sair"
       
         
150      .Refresh
160   End With
      ChecarVersao
End Sub


Sub Ribbon_ButtonClick(ByVal Id As String, ByVal Caption As String, Ribbon As ACPRibbon, frm As Form)
Dim i As Integer
Dim ctl As Control
'Administrativo----------------------------------------------------------------------
With mdiSHB
   If Id = 0 Then .mdiCadCli_Click: GoTo Fim
   If Id = 2 Then .mdiEventos_Click: GoTo Fim
   If Id = 3 Then .mdiAtuEscala_Click: GoTo Fim
   If Id = 4 Then .mdiConsultalogisticaGeral_Click: GoTo Fim
   If Id = 5 Then .mdiAsoExames_Click: GoTo Fim
   If Id = 6 Then .mdiAsoProgramacao_Click: GoTo Fim
   If Id = 7 Then .mdiAsoConsulta_Click: GoTo Fim
   If Id = 8 Then .mdiTreinamentos_Click: GoTo Fim
   If Id = 9 Then .mdiAtuCursosTreinamentos_Click: GoTo Fim
   If Id = 10 Then .mdiConsultaProgCusrsos_Click: GoTo Fim
   If Id = 21 Then .mdiUnidadeOperacional_Click: GoTo Fim
   If Id = 22 Then .mdiProduto_Click: GoTo Fim
   If Id = 23 Then .mdiAtuTabPrecoProduto_Click: GoTo Fim
   If Id = 24 Then .mdiProdutoAtividadePreco_Click: GoTo Fim
   If Id = 25 Then .mdiPedido_Click: GoTo Fim
   If Id = 26 Then .mdiMedicao_Click: GoTo Fim
   If Id = 27 Then .mdiFaturaLocacao_Click: GoTo Fim
   If Id = 28 Then .mdiSerieHistoricaMedicao_Click: GoTo Fim
   
   'Comercial-----------------------------------------------------------------
   If Id = 29 Then .mdiInizEquip_Click: GoTo Fim
   If Id = 30 Then .mdiProposta_Click: GoTo Fim
   If Id = 31 Then .mdiOS_Click: GoTo Fim
   
   'Financeiro---------------------------------------------------------------
   If Id = 41 Then .mdiProdutosIn_Click: GoTo Fim
   If Id = 42 Then .mdiCentroDeCusto_Click: GoTo Fim
   If Id = 43 Then .mdiRecebimentos_Click: GoTo Fim
   If Id = 44 Then .mdiPagamentos_Click: GoTo Fim
   If Id = 45 Then .mdiReprogFinanc_Click: GoTo Fim
   If Id = 46 Then .mdiNfEntrada_Click: GoTo Fim
   If Id = 47 Then .mdiReembolo_Click: GoTo Fim
   If Id = 48 Then .mdiGeraCredito_Click: GoTo Fim
   If Id = 49 Then .mdiGeraExcelDebito_Click: GoTo Fim
   If Id = 50 Then .mdiContasReceber_Click: GoTo Fim
   If Id = 51 Then .mdiConsultaCentroDeCusto_Click: GoTo Fim
   If Id = 52 Then .mdiGeraCustoExcel_Click: GoTo Fim
   If Id = 53 Then .mdiFinancVendas_Click: GoTo Fim
   If Id = 54 Then .mdiFinancAnalitico_Click: GoTo Fim
   If Id = 55 Then .mdiConsolidSemanal_Click: GoTo Fim
   If Id = 56 Then .mdiEmpenho_Click: GoTo Fim
   If Id = 57 Then .mdiCtaPgRec_Click: GoTo Fim
   If Id = 58 Then .mdiCtaPagarReceber_Click: GoTo Fim
   If Id = 59 Then .mdiPagamentosRecebimentos_Click: GoTo Fim
   If Id = 60 Then .mdiConsultaFaturamento_Click: GoTo Fim
   If Id = 61 Then .mdiSevicoNaoFaturado_Click: GoTo Fim
   
   
'Operacional-----------------------------------------------------------------
  If Id = 70 Then .mdiCadEquipto_Click: GoTo Fim
  If Id = 71 Then .mdiEquipamento_Click: GoTo Fim
   
'Suprimentos-----------------------------------------------------------------
  If Id = 80 Then .mdiInventario_Click: GoTo Fim
  If Id = 81 Then .mdiUnidadeEmbalagem_Click: GoTo Fim
  If Id = 82 Then .mdiUnidadeMedida_Click: GoTo Fim
  If Id = 84 Then .mdiFornecProduto_Click: GoTo Fim
  If Id = 85 Then .mdiAcordoComercial_Click: GoTo Fim
  If Id = 86 Then .mdiReqMateriais_Click: GoTo Fim
  If Id = 87 Then .mdiPedidoDeCompra_Click: GoTo Fim
  If Id = 88 Then .mdiRecebeMateriais_Click: GoTo Fim
  If Id = 89 Then .mdiBaixaReqEstoque_Click: GoTo Fim
  If Id = 90 Then .mdiMovimentacaoMateriais_Click: GoTo Fim
  If Id = 91 Then .mdiConsultaPO_Click: GoTo Fim
  If Id = 91 Then .mdiConsultaReq_Click: GoTo Fim
  
  
  If Id = 200 Then .mdiSair_Click: GoTo Fim
   'mdiConsultaReq_Click
   'If Id = 8 Then .mdiFornecProduto_Click: GoTo Fim
      
   'Alterações--------------------------------------------------------------------
   'If Id = 21 If Id = 71 Then .mdiEquipamento_Click: GoTo FimThen frmAlteraFuncionario.Show: GoTo Fim
   'If Id = 22 Then frmAlteraCargo.Show: GoTo Fim
   'If Id = 23 Then frmAlteraVeiculo.Show: GoTo Fim
   'If Id = 24 Then frmAlteraAlojamento.Show: GoTo Fim
   
   'Documentos--------------------------------------------------------------------
   'If Id = 41 Then frmDocumento.Show: GoTo Fim
   
   'Financeiro--------------------------------------------------------------------
   'If Id = 51 Then frmPagamentos.Show: GoTo Fim
   
   'Configura??es-----------------------------------------------------------------
   'If Id = 61 Then frmDDD.Show: GoTo Fim
   'If Id = 63 Then frmEmpresa.Show: GoTo Fim
   
   'Referente ao in?cio, antes do login-------------------------------------------
   'If Id = 150 Then frmLogin.Show: GoTo Fim
   'If Id = 151 Then frmRecuperarSenha.Show: GoTo Fim
   'If Id = 152 Then frmCadPadrao.Show: GoTo Fim
   'If Id = 153 Then AtendimentoOnline: GoTo Fim
   
   'Mudar o Tema------------------------------------------------------------------
   If Id = 62 Then
      If Ribbon.Theme = 2 Then
         Ribbon.Theme = 1
      Else
         Ribbon.Theme = Ribbon.Theme + 1
      End If
      Ribbon.Refresh
   End If
End With
Fim:
For i = 0 To Forms.Count - 1
   Forms(i).Picture = Ribbon.LoadBackground
   Forms(i).BackColor = Ribbon.BackColor
   For Each ctl In Forms(i)
     If TypeOf ctl Is Label Then
        ctl.ForeColor = Ribbon.ForeColor
        ctl.BackColor = Ribbon.BackColor
     End If
     If TypeOf ctl Is Shape Then
        If ctl.Name <> "Shape2" Then
           ctl.BackColor = Ribbon.BackColor
        Else
           ctl.BorderColor = &H80000005
           ctl.Width = Forms(i).Width
        End If
     End If
     If TypeOf ctl Is Frame And ctl.Tag <> "*" Then _
        ctl.BackColor = Ribbon.BackColor: ctl.ForeColor = Ribbon.ForeColor
     If TypeOf ctl Is CheckBox Then ctl.BackColor = Ribbon.BackColor
     If TypeOf ctl Is OptionButton Then ctl.ForeColor = Ribbon.ForeColor
     If TypeOf ctl Is OptionButton Then ctl.BackColor = Ribbon.BackColor
   Next
Next
End Sub




