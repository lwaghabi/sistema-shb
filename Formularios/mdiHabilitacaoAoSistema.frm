VERSION 5.00
Begin VB.MDIForm mdiSHB 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "SHB - Semi Hermatics do Brasil"
   ClientHeight    =   8190
   ClientLeft      =   1395
   ClientTop       =   -1170
   ClientWidth     =   15960
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "mdiHabilitacaoAoSistema.frx":0000
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Menu mdiAdministrativo 
      Caption         =   "Administrativo"
      Begin VB.Menu mdiPessoa 
         Caption         =   "Pessoa"
         Begin VB.Menu mdiCadastro 
            Caption         =   "Cadastro"
            Begin VB.Menu mdiCadCli 
               Caption         =   "Pessoa"
            End
         End
         Begin VB.Menu mdiConsulta 
            Caption         =   "Consultas"
            Begin VB.Menu mdiConsultaCliente 
               Caption         =   "Resumo Cliente"
            End
         End
      End
      Begin VB.Menu mdiLogistica 
         Caption         =   "Log�stica de Pessoal"
         Begin VB.Menu mdiEventos 
            Caption         =   "Atualiza��o de Eventos"
         End
         Begin VB.Menu mdiAtuEscala 
            Caption         =   "Atualiza��o de Escalas de Pessoal"
         End
         Begin VB.Menu mdiConsultaLogisticaGeral 
            Caption         =   "Consulta Log�stica Geral"
         End
      End
      Begin VB.Menu mdiAso 
         Caption         =   "ASO"
         Begin VB.Menu mdiAsoExames 
            Caption         =   "Registro e atualiza��o de Exames"
         End
         Begin VB.Menu mdiAsoProgramacao 
            Caption         =   "Registro de Agenda de Exames - ASO Programa��o"
         End
         Begin VB.Menu mdiAsoConsulta 
            Caption         =   "Consulta Agenda de Funcion�rios por Exames"
         End
      End
      Begin VB.Menu mdiTreinamento 
         Caption         =   "Treinamento"
         Begin VB.Menu mdiTreinamentos 
            Caption         =   "Registro e Atualiza��o de Cursos e Treinamentos"
         End
         Begin VB.Menu mdiAtuCursosTreinamentos 
            Caption         =   "Programa��o de Cursos e Treinamentos por Funcion�rio"
         End
         Begin VB.Menu mdiConsultaProgCusrsos 
            Caption         =   "Consulta a Programa��o de Cursos"
         End
      End
   End
   Begin VB.Menu mdiComercial 
      Caption         =   "Comercial"
      Begin VB.Menu mdiParametros 
         Caption         =   "Par�metros de Negocia��o"
         Begin VB.Menu mdiUnidadeOperacional 
            Caption         =   "Registro e Atualiza��o de Unidade Operacional"
         End
         Begin VB.Menu mdiProduto 
            Caption         =   "Registro de Contratos e Produtos"
         End
         Begin VB.Menu mdiAtuTabPrecoProduto 
            Caption         =   "Atualiza��o da Tabela de Pre�os de Produtos e Contratos"
         End
         Begin VB.Menu mdiProdutoAtividadePreco 
            Caption         =   "Consulta de Pre�os Praticados por Contrato"
         End
      End
      Begin VB.Menu mdiNeg 
         Caption         =   "Negocia��o"
         Begin VB.Menu mdiProcesEControles 
            Caption         =   "Processamento e Controles"
            Begin VB.Menu mdiPedido 
               Caption         =   "Registro e Processamento de Medi��o"
            End
            Begin VB.Menu mdiMedicao 
               Caption         =   "Controle e Emiss�o de Medi��o"
            End
            Begin VB.Menu mdiFaturaLocacao 
               Caption         =   "Emiss�o de Fatura de Loca��o"
            End
            Begin VB.Menu mdiPagtoEmCheque 
               Caption         =   "Pagtos. de Frete de Sa�da em Cheques"
               Visible         =   0   'False
            End
            Begin VB.Menu mdiDevolucaoNegociacao 
               Caption         =   "Devolu��o de Compras"
               Visible         =   0   'False
            End
            Begin VB.Menu mdiControleFaturamento 
               Caption         =   "Emiss�o de Controle de Faturamento"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mdiLocacoesServicos 
            Caption         =   "Propostas de Loca��es e Servi�os"
            Begin VB.Menu mdiInizEquip 
               Caption         =   "Indeniza��o de Equipamentos"
            End
            Begin VB.Menu mdiProposta 
               Caption         =   "Proposta"
            End
            Begin VB.Menu mdiOS 
               Caption         =   "Registro de O.S."
            End
         End
         Begin VB.Menu mdiConsultaNegociacao 
            Caption         =   "Consultas"
            Begin VB.Menu mdiExtratoNotaFiscal 
               Caption         =   "Extrato de Nota Fiscal"
            End
            Begin VB.Menu mdiSerieHistoricaMedicao 
               Caption         =   "Consulta Anal�tica de Medi��es"
            End
            Begin VB.Menu mdiMovCli 
               Caption         =   "Movimenta��o de Pedidos por Cliente"
               Visible         =   0   'False
            End
            Begin VB.Menu mdiPedidosPendentes 
               Caption         =   "Movimenta��o de Pedidos no mes atual"
               Visible         =   0   'False
            End
            Begin VB.Menu mdiInformaoesFinanceiras 
               Caption         =   "Consulta a Informa��es Financeiras"
               Visible         =   0   'False
            End
            Begin VB.Menu mdiPrazoEntrega 
               Caption         =   "Estat�stica de Prazo de  Atendimento na Entrega  de Pedidos"
               Visible         =   0   'False
            End
            Begin VB.Menu mdiProdutoPeriodo 
               Caption         =   "Sa�da de Produto por Per�odo"
               Visible         =   0   'False
            End
            Begin VB.Menu mdiSaidasprodutos 
               Caption         =   "Sa�das de Produtos"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mdiRelatoriosNegociacao 
            Caption         =   "Relat�rios de Negocia��o"
            Visible         =   0   'False
            Begin VB.Menu mdiNotaFiscal 
               Caption         =   "Nota Fiscal"
            End
            Begin VB.Menu mdiImpTabPrecos 
               Caption         =   "Tabela de Pre�os de Produtos no SIM"
            End
            Begin VB.Menu mdiImpTabFrete 
               Caption         =   "Tabela de Fretes de Produtos no SIM"
            End
            Begin VB.Menu mdiNegMes 
               Caption         =   "Mapa de Negocia��es Faturadas no M�s"
            End
            Begin VB.Menu mdiNegMesConsig 
               Caption         =   "Mapa de Negocia��es Consignadas no M�s"
            End
            Begin VB.Menu mdiAcompVendasAnual 
               Caption         =   "Mapa de Acompanhamento de Vendas Anual"
            End
         End
      End
      Begin VB.Menu mdiColaboradores 
         Caption         =   "Colaboradores"
         Visible         =   0   'False
         Begin VB.Menu mdiCarteira 
            Caption         =   "Carteira de Clientes"
            Begin VB.Menu mdiAtuCarteira 
               Caption         =   "Atualiza��o de Carteira de Clientes"
            End
         End
         Begin VB.Menu mdiConsultaColab 
            Caption         =   "Consulta"
            Visible         =   0   'False
            Begin VB.Menu mdiClientesCarteira 
               Caption         =   "Clientes por Carteira"
            End
            Begin VB.Menu mdiClientePromot 
               Caption         =   "Clientes por Promotor"
            End
            Begin VB.Menu mdiCliRep 
               Caption         =   "Mapa de Clientes por Representantes"
            End
            Begin VB.Menu mdiMapaComissao 
               Caption         =   "Mapa de Comiss�es de Representantes"
            End
         End
      End
      Begin VB.Menu mdiRelatorios 
         Caption         =   "Relat�rios"
         Visible         =   0   'False
         Begin VB.Menu mdiCliProdRep 
            Caption         =   "Clientes e Produtos por Representante"
         End
         Begin VB.Menu mdiImpPerformanceCliAnual 
            Caption         =   "Performance de Negocia��o de Grupo nos �ltimos 12 Meses"
         End
         Begin VB.Menu mdiNegLinhaProd 
            Caption         =   "Performance de Negocia��o Por Modelo de Produto"
         End
         Begin VB.Menu mdiContatoCliRep 
            Caption         =   "Contatos de Clientes por Representantes"
         End
         Begin VB.Menu mdiMapaInativos 
            Caption         =   "Mapa de Clientes Inativos por Representante"
         End
         Begin VB.Menu mdiImpProdConsig 
            Caption         =   "Mapa de Produtos Consignados Pendentes de Apura��o"
         End
         Begin VB.Menu mdiProdConsig 
            Caption         =   "Mapa de Produtos Consignados por Nota Fiscal"
         End
         Begin VB.Menu mdiImpProdAnual 
            Caption         =   "Mapa de Performance de Produtos "
         End
         Begin VB.Menu mdiImpNegUF 
            Caption         =   "Mapa de Faturamento por Regi�o "
         End
         Begin VB.Menu mdiMovProdCliPeriodo 
            Caption         =   "Movimenta��o de Produtos por Clientes por Per�odo"
         End
         Begin VB.Menu mdiRelPessoa 
            Caption         =   "Rela��o de Clientes, Ender��os e Contatos"
         End
         Begin VB.Menu mdiPerformanceRepres 
            Caption         =   "Performance de Representantes e Produtos"
         End
         Begin VB.Menu mdiCliCidade 
            Caption         =   "Clientes por Cidade"
         End
      End
      Begin VB.Menu mdiConsultsEspeciais 
         Caption         =   "Cons. Especiais"
         Visible         =   0   'False
         Begin VB.Menu mdiClassifcPrdEntregue 
            Caption         =   "Classifica��o de Entrega Por Produto"
         End
         Begin VB.Menu mdiPrazoEntregaCE 
            Caption         =   "Estat�stica de Prazo de  Atendimento na Entrega  de Pedidos"
         End
         Begin VB.Menu mdiEstatisticaNeg 
            Caption         =   "Estat�stica de Negocia��o por UF"
         End
         Begin VB.Menu mdiEstatisticaNegRep 
            Caption         =   "Estatistica de Negocia��o por Representante"
         End
         Begin VB.Menu mdiPerformancePrdRegiao 
            Caption         =   "Performance de Produtos por Regi�o"
         End
         Begin VB.Menu mdiHistoricoProd 
            Caption         =   "Hist�rico de Performance de Produtos"
         End
         Begin VB.Menu mdiestatisticaporregiao 
            Caption         =   "Volume Negociado por Regi�o nos �ltimos 12 meses"
         End
         Begin VB.Menu mdiEvolucaoEntregas 
            Caption         =   "Evolu��o de Entregas de Produtos por Per�odo"
         End
      End
   End
   Begin VB.Menu mdiAdmFinanc 
      Caption         =   "Adm.Financeira"
      Begin VB.Menu mdiRecursosHumanos 
         Caption         =   "Recursos Humanos"
         Visible         =   0   'False
      End
      Begin VB.Menu mdiFinanceiro 
         Caption         =   "Financeiro"
         Begin VB.Menu mdiParmFinanc 
            Caption         =   "Par�metros Financeiros"
            Begin VB.Menu mdiProdutosIn 
               Caption         =   "Cadastramento de Servi�os e Produtos de Fornecedores"
            End
            Begin VB.Menu mdiCentroDeCusto 
               Caption         =   "Registro e Atualiza��o de Centro de Custo"
            End
         End
         Begin VB.Menu mdiControleFinanceiro 
            Caption         =   "Controle Financeiro"
            Begin VB.Menu mdiRecebimentos 
               Caption         =   "Recebimentos "
            End
            Begin VB.Menu mdiPagamentos 
               Caption         =   "Pagamentos"
            End
            Begin VB.Menu mdiAtualizaPrecoProdutoConsignado 
               Caption         =   "Atualiza Pre�o de Produto Consignado"
               Visible         =   0   'False
            End
            Begin VB.Menu mdiReprogFinanc 
               Caption         =   "Reprograma��o Financeira"
            End
            Begin VB.Menu mdiAjusteComissao 
               Caption         =   "Ajuste de Comiss�es"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mdiLancamentos 
            Caption         =   "Lan�amentos de Contas "
            Begin VB.Menu mdiCtaPagar 
               Caption         =   "a Pagar"
               Begin VB.Menu mdiNfEntrada 
                  Caption         =   "Notas Fiscais de Entrada"
               End
               Begin VB.Menu mdiNFSuprimentos 
                  Caption         =   "Nota Fiscal de Suprimentos"
               End
               Begin VB.Menu mdiReembolo 
                  Caption         =   "Reembolso de Pagamentos de Despesas"
               End
               Begin VB.Menu mdiReciboPagamento 
                  Caption         =   "Recibo de Pagamentos"
               End
            End
            Begin VB.Menu mdiCtaReceber 
               Caption         =   "a Receber"
               Begin VB.Menu mdiGeraCredito 
                  Caption         =   "Recebimentos"
               End
            End
         End
         Begin VB.Menu mdiGerarExcel 
            Caption         =   "Gerar Excel"
            Begin VB.Menu mdiGeraExcelDebito 
               Caption         =   "Gerar Contabilidade"
            End
            Begin VB.Menu mdiContasReceber 
               Caption         =   "Gerar Excel a Receber"
            End
         End
         Begin VB.Menu mdiCusto 
            Caption         =   "Custo"
            Begin VB.Menu mdiConsultaCentroDeCusto 
               Caption         =   "Consulta Centro de Custo"
            End
            Begin VB.Menu mdiGeraCustoExcel 
               Caption         =   "Custo em Excel"
            End
         End
         Begin VB.Menu mdiConsultaFinanc 
            Caption         =   "Consulta Financeiro"
            Begin VB.Menu mdiFinancVendas 
               Caption         =   "Lan�amentos Financeiros no Dia (Compra e Venda)"
            End
            Begin VB.Menu mdiFinancAnalitico 
               Caption         =   "Financeiro Anal�tico"
            End
            Begin VB.Menu mdiConsolidSemanal 
               Caption         =   "Financeiro Consolidado"
            End
            Begin VB.Menu mdiEmpenho 
               Caption         =   "Proje��o Financeiro"
            End
            Begin VB.Menu mdiConsultaFinanceiro 
               Caption         =   "Consulta a Informa��es Financeiras"
               Visible         =   0   'False
            End
            Begin VB.Menu mdiFinancCliente 
               Caption         =   "Financeiro por Cliente/Colaborador"
               Visible         =   0   'False
            End
            Begin VB.Menu mdiCtaPgRec 
               Caption         =   "Contas Pagas e Recebidas at� a Data"
            End
            Begin VB.Menu mdiCentroDeCustoNew 
               Caption         =   "Contas Pagas por Centro de Custo"
            End
            Begin VB.Menu mdiCtaPagarReceber 
               Caption         =   "Contas a Pagar e a Receber"
            End
            Begin VB.Menu mdiPagamentosRecebimentos 
               Caption         =   "Consulta a Pagamentos Realizados"
            End
            Begin VB.Menu mdiConsultaFaturamento 
               Caption         =   "Consulta Faturamento por Per�odo"
            End
            Begin VB.Menu mdiSevicoNaoFaturado 
               Caption         =   "Loca��es e Servi�os Prestados n�o Faturados"
            End
         End
         Begin VB.Menu mdiRelFinanc 
            Caption         =   "Relat�rios Financeiros"
            Visible         =   0   'False
            Begin VB.Menu mdiMapaPagtos 
               Caption         =   "Mapa de Pagamentos Di�rios"
            End
            Begin VB.Menu mdiMapaRecebimentos 
               Caption         =   "Mapa de Recebimentos no Per�odo"
            End
            Begin VB.Menu mdiClientesEmAtraso 
               Caption         =   "Rela��o de Clientes em Atraso"
            End
            Begin VB.Menu mdiImpConsignacaoPendente 
               Caption         =   "Mapa de Consigna��es Pendentes de Apura��o"
            End
            Begin VB.Menu mdiConsigApurada 
               Caption         =   "Mapa de Consigna��es Apuradas no Per�odo"
            End
            Begin VB.Menu mdiFaturamentoAnual 
               Caption         =   "Faturamento nos �ltimos 12 Meses"
            End
         End
      End
   End
   Begin VB.Menu mdiOperacional 
      Caption         =   "Operacional"
      Begin VB.Menu mdiCadEquipamentos 
         Caption         =   "Equipamentos"
         Begin VB.Menu mdiCadEquipto 
            Caption         =   "Cadastro de Equipamentos"
         End
         Begin VB.Menu mdiEquipamento 
            Caption         =   "Registro e Atualiza��o de Equipamentos"
         End
      End
   End
   Begin VB.Menu mdiIndustrial 
      Caption         =   "Industrial"
      Visible         =   0   'False
      Begin VB.Menu mdiProducao 
         Caption         =   "Produ��o"
         Begin VB.Menu mdiMovProducao 
            Caption         =   "Movimenta��o de Produ��o Di�ria"
         End
         Begin VB.Menu mdiMoveEspecial 
            Caption         =   "Movimento Especial de Produ��o"
         End
         Begin VB.Menu mdiConsultaProducao 
            Caption         =   "Consulta a Produ��o"
            Begin VB.Menu mdiPosGeralNeg 
               Caption         =   "Posi��o Geral de Estoque"
            End
            Begin VB.Menu mdiEstoquePrdAcabado 
               Caption         =   "Estoque de Produto Acabado"
            End
            Begin VB.Menu mdiEstoquePedido 
               Caption         =   "Posi��o de Estoques e Pedidos"
            End
            Begin VB.Menu mdiProdUnidadeFabril 
               Caption         =   "Produ��o por Unidade Fabril"
            End
            Begin VB.Menu mdiHistProducao 
               Caption         =   "Historico de Produ��o"
            End
         End
         Begin VB.Menu mdiRelProd 
            Caption         =   "Relat�rios de Produ��o"
            Begin VB.Menu mdiApoioProd 
               Caption         =   "Mapa de Apoio a Produ��o"
            End
            Begin VB.Menu mdiProdDiaria 
               Caption         =   "Mapa de Produ��o Di�ria"
            End
         End
      End
   End
   Begin VB.Menu mdiMateriaisEst 
      Caption         =   "Suprimentos"
      Begin VB.Menu mdiCadastroProdutos 
         Caption         =   "Cadastro de Produtos"
         Begin VB.Menu mdiClassificaProduto 
            Caption         =   "Classifica��o de Produtos em Estoque"
         End
         Begin VB.Menu mdiFornecProduto 
            Caption         =   "Fornecedores/Produtos "
         End
      End
      Begin VB.Menu mdiConsultaMateriais 
         Caption         =   "Movimenta��o de Materiais"
         Begin VB.Menu mdiReqMateriais 
            Caption         =   "Requisi��o de Materiais"
         End
         Begin VB.Menu mdiAcordoComercial 
            Caption         =   "Registro de Acordo Comercial"
         End
         Begin VB.Menu mdiPedidoDeCompra 
            Caption         =   "Pedido de Compra"
         End
         Begin VB.Menu mdiRecebeMateriais 
            Caption         =   "Recebimento de Materiais"
         End
      End
      Begin VB.Menu mdiParameto 
         Caption         =   "Par�metros"
         Begin VB.Menu mdiUnidadeEmbalagem 
            Caption         =   "Unidade de Embalagem"
         End
         Begin VB.Menu mdiUnidadeMedida 
            Caption         =   "Unidade de Medida"
         End
      End
   End
   Begin VB.Menu mdiSupervisao 
      Caption         =   "Supervis�o"
      Begin VB.Menu mdiSupervisor 
         Caption         =   "Supervisor"
         Begin VB.Menu mdiAbreFecha 
            Caption         =   "Abertura e Encerramento do Sistema"
         End
         Begin VB.Menu mdiEndereco 
            Caption         =   "Registro e Atualiza��o de Endere�os"
         End
         Begin VB.Menu mdiUsuario 
            Caption         =   "Cadastramento e Altera��o de Usu�rio e Senha"
         End
         Begin VB.Menu mdiParEsp 
            Caption         =   "Par�metros Especiais"
         End
         Begin VB.Menu mdiValoresPagosRecebidosTrimestre 
            Caption         =   "Valores Pagos e Recebidos por Trimestre"
         End
         Begin VB.Menu mdiComposicao 
            Caption         =   "Composi��o de Pre�os"
            Visible         =   0   'False
         End
         Begin VB.Menu mdiISS 
            Caption         =   "ISS"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu mdiHabilitacao 
      Caption         =   "Habilita��o"
      Begin VB.Menu mdiHabilitacaoSistema 
         Caption         =   "Habilita��o ao Sistema"
      End
      Begin VB.Menu mdiCalculadora 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu mdiUsuarioSenha 
         Caption         =   "Usu�rio Senha"
      End
      Begin VB.Menu mdiModelo 
         Caption         =   "Modelo"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mdiSair 
      Caption         =   "Encerrar"
   End
End
Attribute VB_Name = "mdiSHB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Data_Hoje_Mdi As Date
Dim ano As Integer
Dim Mes As Integer
Dim Dia As Integer
Dim AnoDb As Integer
Dim MesDb As Integer
Dim DiaDb As Integer
Dim DataHojeInvertida As String
Dim DataInvertida As String
Dim DataDias As String
Dim DataBase As String
Dim UsuarioLocal As String
Dim EncontreiAviso As Integer
Dim PessoaAnterior As String

Private Sub mdiConsltaParametros_Click()

End Sub

Private Sub mdiAcordoComercial_Click()
frmAcordoComercial.Show
End Sub

Private Sub mdiEndereco_Click()
frmEndereco.Show
End Sub

Private Sub mdiGeraCustoExcel_Click()
frmGeraCustoExcel.Show
End Sub

Private Sub mdiAtuCursosTreinamentos_Click()
frmTreinamentoAgenda.Show
End Sub

Private Sub mdiCadEquipto_Click()
frmEquipamentoTipo.Show
End Sub

Private Sub mdiConsultaCentroDeCusto_Click()
frmConsultaCentroDeCusto.Show
End Sub

Private Sub mdiConsultaFaturamento_Click()
frmConsultaFaturamento.Show
End Sub

Private Sub mdiConsultaProgCusrsos_Click()
frmTreinamentoConsulta.Show
End Sub

Private Sub mdiContasReceber_Click()
frmGeraExcelCredito.Show
End Sub

Private Sub mdiEquipamento_Click()
frmEquipamento.Show
End Sub

Private Sub mdiGeraExcelDebito_Click()
frmGeraExcelDebito.Show
End Sub

Private Sub mdiInizEquip_Click()
frmIndenizEquip.Show
End Sub

Private Sub mdiNFSuprimentos_Click()
frmNFSuprimentos.Show
End Sub

Private Sub mdiOS_Click()
frmOS.Show
End Sub

Private Sub mdiPedidoDeCompra_Click()
frmPO.Show
End Sub

Private Sub mdiProposta_Click()
frmProposta.Show
End Sub

Private Sub mdiRecebeMateriais_Click()
frmRecebProdutos.Show
End Sub

Private Sub mdiReqMateriais_Click()
frmRequisicao.Show
End Sub

Private Sub mdiSerieHistoricaMedicao_Click()
frmSerieHistoricaMedicao.Show
End Sub

Private Sub mdiSevicoNaoFaturado_Click()
frmLocacaoeServicoNaoFaturado.Show
End Sub

Private Sub mdiTreinamentos_Click()
frmTreinamentos.Show
End Sub

Private Sub mdiAsoConsulta_Click()
frmAsoConsulta.Show
End Sub

Private Sub mdiAsoExames_Click()
frmAsoExames.Show
End Sub

Private Sub mdiAsoProgramacao_Click()
frmAsoAgenda.Show
End Sub

Private Sub mdiAtuEscala_Click()
frmEscalaDePessoal.Show
End Sub

Private Sub mdiConsultalogisticaGeral_Click()
frmConsultaLogGeral.Show
End Sub

Private Sub mdiEmpenho_Click()
frmEmpenho.Show
End Sub

Private Sub mdiEventos_Click()
frmEventoDeLogistica.Show
End Sub

Private Sub mdiPagamentosRecebimentos_Click()
frmPagamentosRecebimentos.Show
End Sub

Private Sub mdiParEsp_Click()
If glbUsuario = "Pablo" Or glbUsuario = "pablo" Then
   frmFaturaLocacaoEsp.Show
Else
   MsgBox ("Fun��o em Desenvolvimento."), vbInformation
End If
End Sub

Private Sub mdiProdutoAtividadePreco_Click()
frmProdutoAtividadePreco.Show
End Sub

Private Sub mdiReembolo_Click()
frmReembolso.Show
End Sub
Private Sub mdiAcompVendasAnual_Click()
MsgBox ("Fun��o n�o Dispon�vel") 'frmAcompVendasAnual.Show
End Sub

Private Sub mdiApoioProd_Click()
MsgBox ("Fun��o n�o Dispon�vel") 'frmMapaApoioProducao.Show
End Sub

'Private Sub mdiApuraConsignacao_Click()
'MsgBox ("Fun��o n�o Dispon�vel") 'frmApuraConsignacao.Show
'End Sub

'Private Sub mdiAtualizaPrecoProdutoConsignado_Click()
'MsgBox "Fun��o em desenvolvimento" 'frmAtualizaPrecoConsignado.Show
'End Sub

'Private Sub mdiAtuCarteira_Click()
'MsgBox "Fun��o em desenvolvimento" 'frmCarteiraRepresentante.Show
'End Sub

Private Sub mdiCalculadora_Click()
frmCalculadora.Show
End Sub

Private Sub mdiCentroDeCusto_Click()
frmCentroDeCusto.Show
End Sub

Private Sub mdiCentroDeCustoNew_Click()
frmCentroDeCustoNew.Show
End Sub

Private Sub mdiClassifcPrdEntregue_Click()
MsgBox "Fun��o em desenvolvimento" 'frmClassifcPrdEntregue.Show
End Sub

Private Sub mdiClassificaProduto_Click()
'frmClassificaProdutosEstoque.Show
End Sub

Private Sub mdiClientesEmAtraso_Click()
'frmImpClientesEmAtraso.Show
MsgBox ("Relat�rio n�o Dispon�vel")
End Sub

Private Sub mdiCliProdRep_Click()
MsgBox "Fun��o em desenvolvimento" 'frmImpCliProdRep.Show
End Sub

Private Sub mdiConsigApurada_Click()
MsgBox "Fun��o em desenvolvimento" 'frmConsigApurada.Show
End Sub

'Private Sub mdiConsultaFinanceiro_Click()
'frmConsultaFinanceiro.Show
'End Sub

Private Sub mdiContatoCliRep_Click()
MsgBox "Fun��o em desenvolvimento" 'frmContatoCliRep.Show
End Sub

'Private Sub mdiEmpenho_Click()
'MsgBox "Fun��o em desenvolvimento" 'frmEmpenho.Show
'End Sub

Private Sub mdiEstatisticaNeg_Click()
MsgBox "Fun��o em desenvolvimento" 'frmEstatisticaUF.Show vbModal
End Sub

Private Sub mdiEstatisticaNegRep_Click()
MsgBox "Fun��o em desenvolvimento" 'frmEstatisticaNegRep.Show vbModal
End Sub

Private Sub mdiestatisticaporregiao_Click()
MsgBox "Fun��o em desenvolvimento" 'frmEstatisticaPorRegiao.Show vbModal
End Sub

Private Sub mdiExtratoNotaFiscal_Click()
frmExtratoNotaFiscal.Show
End Sub

Private Sub mdiFinancCliPeriodo_Click()
MsgBox "Fun��o em desenvolvimento" 'frmImpFinancCliPeriodo.Show
End Sub

Private Sub mdiFaturaLocacao_Click()
frmFaturaLocacao.Show
End Sub

Private Sub mdiFaturamentoAnual_Click()
MsgBox "Fun��o em desenvolvimento" 'frmImpFaturamentoAnual.Show vbModal
End Sub

Private Sub mdiFornecProduto_Click()
frmSupProduto.Show
End Sub

Private Sub mdiGeraCredito_Click()
frmGeraCredito.Show vbModal
End Sub

Private Sub mdiGrupoProdutos_Click()
MsgBox "Fun��o n�o dispon�vel " 'frmGrupoProdutoEstoque.Show
End Sub

Private Sub mdiHistoricoProd_Click()
MsgBox "Fun��o em desenvolvimento" 'frmHistoricoProd.Show
End Sub

Private Sub mdiHistProducao_Click()
MsgBox "Fun��o n�o dispon�vel " 'frmHistProducao.Show
End Sub

Private Sub mdiImpConsignacaoPendente_Click()
'frmImpConsigPendente.Show
MsgBox ("Fun��o n�o Dispon�vel")
End Sub

Private Sub mdiImpNegUF_Click()
MsgBox "Fun��o em desenvolvimento" 'frmImpNegUF.Show
End Sub

Private Sub mdiImpPerformanceCliAnual_Click()
MsgBox "Fun��o em desenvolvimento" 'frmImpPerformanceCliAnual.Show
End Sub

Private Sub mdiImpProdAnual_Click()
MsgBox "Fun��o em desenvolvimento" 'frmImpProdAnual.Show
End Sub

Private Sub mdiImpProdConsig_Click()
'frmImpProdConsig.Show
MsgBox ("Fun��o n�o Dispon�vel")
End Sub

Private Sub mdiImpTabFrete_Click()
MsgBox "Fun��o em desenvolvimento" 'frmImpTabFrete.Show
End Sub

Private Sub mdiImpTabPrecos_Click()
MsgBox "Fun��o em desenvolvimento" 'frmImpTabPrecos.Show
End Sub

Private Sub mdiInformaoesFinanceiras_Click()
frmConsultaFinanceiro.Show
End Sub

Private Sub mdiMapaPagtos_Click()
MsgBox "Fun��o em desenvolvimento" 'frmMapaPagamentos.Show
End Sub

Private Sub mdiAbreFecha_Click()
frmAbre_Fecha.Show
End Sub

'Private Sub mdiAjusteComissao_Click()
'frmAjusteComissao.Show
'End Sub
'Private Sub mdiatualizaTabPrecoFrete_Click()
'MsgBox "Fun��o em desenvolvimento" 'frmAtualizaPrecoFrete.Show
'End Sub
Private Sub mdiAtuTabPrecoProduto_Click()
frmAtualizaPrecoProd.Show
End Sub
Private Sub mdiCadCli_Click()

Call Rotina_AbrirBanco

glb.Open "Select * from Global where chDataAbertura = ('" & DataHojeInvertida & "')", db, 3, 3

If glb.EOF Then
   MsgBox ("Aten��o: O sistema encontra-se fechado. Esta fun��o s� pode ser usada ap�s a sua abertura."), vbInformation
   Exit Sub
End If
frmPessoa.Show
End Sub
Private Sub mdiCidadeBairro_Click()
MsgBox "Fun��o n�o dispon�vel"  'frmCidadeBairro.Show
End Sub

Private Sub mdiCliCidade_Click()
MsgBox "Fun��o em desenvolvimento" 'impCliCidade.Show
End Sub

Private Sub mdiClienteRepresentante_Click()
MsgBox "Fun��o em desenvolvimento"
End Sub

Private Sub mdiCliRep_Click()
MsgBox "Fun��o em desenvolvimento" 'frmCliRep.Show
End Sub

Private Sub mdiConsolidSemanal_Click()
frmConsolidadoFinanc.Show
End Sub
'Private Sub mdiConsultaCliente_Click()
'frmResumoCliente.Show
'End Sub

Private Sub mdiControleFaturamento_Click()
MsgBox "Fun��o em desenvolvimento" 'frmControleFaturamentoNew.Show
End Sub

Private Sub mdiCtaPgRec_Click()
frmCtaReceb.Show
End Sub
Private Sub mdiCtaPagarReceber_Click()
frmCtaPagar.Show
End Sub
'Private Sub mdiDevolucaoNegociacao_Click()
'frmDevolucaoNegociacao.Show
'End Sub

Private Sub mdiEstoquePedido_Click()
MsgBox "Fun��o n�o dispon�vel" 'frmEstoqueProdutoAcabado.Show
End Sub

Private Sub mdiEvolucaoEntregas_Click()
MsgBox "Fun��o em desenvolvimento" 'frmEvolucaoEntregas.Show
End Sub

Private Sub mdiFinancAnalitico_Click()
frmFinancAnalitico.Show
End Sub

Private Sub mdiFinancCliente_Click()
frmFinancCliente.Show
End Sub

Private Sub mdiFinancVendas_Click()
frmConsultaMovFinanc.Show
End Sub

Private Sub MDIForm_Load()
Dim UsuarioLocal As String
frmUsuarioSenha.Show vbModal
If Not glbUsuario = Empty Then
    If Not Compilando Then _
    MsgBox ("Voc� esta logado no SHB atrav�s da M�quina ") & glbMaquina & ", endere�o IP " & glbEnderecoIP

    mdiHabilitacaoSistema.Enabled = True
    ano = Year(Date)
    Mes = Month(Date)
    Dia = Day(Date)
    Data_Hoje = Date

    
    Call Rotina_AbrirBanco
       
    DataHojeInvertida = ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")
       
    usu.Open "Select * from Usuario where chNome = ('" & glbUsuario & "')", db, 3, 3
    If usu.EOF Then
       MsgBox ("Erro no acesso a Usuario na rotina de atualiza��o de mostrar aviso. Comunicar Analista respons�vel"), vbCritical
       End
    End If
    
    UsuarioLocal = usu!usuMostrarAviso
    
    EncontreiAviso = 0
    
    If usu!usuMostrarAviso = 1 Then
       Call VerificaASO
       If EncontreiAviso = 1 Then
          frmAviso.Show vbModal
       End If
    End If
    
    EncontreiAviso = 0
    
    If usu!usuAvisoTreinamento = 1 Then
       Call VerificaTreinamento
      
       If EncontreiAviso = 1 Then
          frmAvisoTreinamento.Show vbModal
       End If
    End If
    
    EncontreiAviso = 0
    
    If usu!usuAvisoReembolso = 1 Then
      Call VerificaReembolso
      
      If EncontreiAviso = 1 Then
         frmAvisoReembolso.Show vbModal
      End If
    End If
    
    EncontreiAviso = 0
    
    Call VerificaEquipamento
    
    If EncontreiAviso = 1 Then
       frmAvisoEquipamentos.Show vbModal
    End If
        
    ' MsgBox ("Data Hoje Invertida "), DataHojeInvertida
    frmControleTempo.Show
    frmControleTempo.Visible = False
Else
   End
End If
End Sub

'Private Sub mdiHabilitacaoSistema_Click()

'Set TabUsuario = dbSHB.OpenRecordset("Usuario")
'                    TabUsuario.Index = "IndUsuario"
                    
'TabUsuario.Seek "=", glbUsuario
'If TabUsuario.NoMatch Then
'   MsgBox "Usuario n�o Cadastrado"
'   mdiHabilitacaoSistema.Enabled = False
'   Exit Sub
'End If

'If TabUsuario("usustatus") = 1 Then
'   MsgBox "Este usuario esta ativo atraves de outro equipamento. Verificar."
'   Exit Sub
'End If


'End Sub

'Private Sub mdiICMS_Click()
'ICMS.Show
'End Sub

'Private Sub mdiMapaComissao_Click()
'MsgBox ("Fun��o n�o Dispon�vel") 'frmMapaComissao.Show
'End Sub

'Private Sub mdiMapaRecebimentos_Click()
'MsgBox ("Fun��o n�o Dispon�vel") 'frmMapaRecebimentos.Show
'End Sub

Private Sub mdiMostruario_Click()
MsgBox "Fun��o n�o dispon�vel" 'frmMostruario.Show
End Sub

Private Sub mdiMedicao_Click()
frmMedicao.Show
End Sub

Private Sub mdiMovCli_Click()
MsgBox "Fun��o n�o dispon�vel" 'frmMovCli.Show
End Sub

Private Sub mdiMoveEspecial_Click()
MsgBox ("Fun��o n�o Dispon�vel") 'frmMovEspecial.Show
End Sub

Private Sub mdiMovMostruario_Click()
'MsgBox "Fun��o em desenvolvimento. Maiores informa��es com o Sr. Luiz."
MsgBox "Fun��o n�o dispon�vel " 'frmControleMostruario.Show
End Sub

Private Sub mdiMovProdCliPeriodo_Click()
MsgBox ("Fun��o n�o Dispon�vel") 'frmMovProdCliPeriodo.Show
End Sub

Private Sub mdiMapaInativos_Click()
MsgBox ("Fun��o n�o Dispon�vel") 'frmMapaInativos.Show
End Sub
'Private Sub mdiMovProducao_Click()
'TabGlobal.Seek "=", Date
'If TabGlobal.NoMatch Then
'   MsgBox ("Aten��o: Sistema encontra-se fechado.")
'End If
'MsgBox "Fun��o n�o dispon�vel " 'frmProducao.Show vbModal
'End Sub

Private Sub mdiNegLinhaProd_Click()
'
MsgBox ("Fun��o n�o Dispon�vel") 'frmNegLinhaProd.Show

End Sub

Private Sub mdiNegMes_Click()
MsgBox ("Fun��o n�o Dispon�vel") 'impNegMes.Show vbModal
'deNegMes.rscmdNegMes.Close
End Sub

Private Sub mdiNegMesConsig_Click()
MsgBox ("Fun��o n�o Dispon�vel") 'impNegMesConsig.Show vbModal
'deNegMesConsig.rscmdNegMesConsig.Close
End Sub

Private Sub mdiNfEntrada_Click()
frmNotaFiscalEntrada.Show
End Sub

Private Sub mdiNotaFiscal_Click()
MsgBox "Fun��o n�o dispon�vel" 'frmImpNotaFiscal.Show
End Sub

Private Sub mdiPagamentos_Click()
Call Rotina_AbrirBanco

glb.Open "Select * from Global where chDataAbertura = ('" & DataHojeInvertida & "')", db, 3, 3

If glb.EOF Then
   MsgBox ("Aten��o: O sistema encontra-se fechado. Esta fun��o s� pode ser usada ap�s a abertura do sitema."), vbInformation
   Exit Sub
End If

Call FechaDB

frmControlePagamentos.Show
End Sub

Private Sub mdiPagtoEmCheque_Click()
MsgBox ("Fun��o n�o dispon�vel. Em manuten��o.")
'frmPagtoEmCheque.Show
End Sub

Private Sub mdiPedCompra_Click()
MsgBox ("Fun��o n�o Dispon�vel") 'frmPedidoDeCompra.Show
End Sub

Private Sub mdiPedido_Click()
Call Rotina_AbrirBanco
glb.Open "Select * from Global where chDataAbertura = ('" & DataHojeInvertida & "')", db, 3, 3
If glb.EOF Then
   MsgBox ("Aten��o: O sistema encontra-se fechado. Esta fun��o s� pode ser usada ap�s a abertura do sitema."), vbInformation
   Exit Sub
End If

Call FechaDB

frmPedido.Show

End Sub

Private Sub mdiPedidosPendentes_Click()
MsgBox "Fun��o n�o dispon�vel" 'frmPedidosPendentes.Show
End Sub

Private Sub mdiPerformancePrdRegiao_Click()
MsgBox "Fun��o n�o dispon�vel" 'frmEstatisticaRegiaoProduto.Show
End Sub

Private Sub mdiPerformanceRepres_Click()

MsgBox ("Relat�rio n�o Dispon�vel. ")
'frmPerformRepres.Show
End Sub

'Private Sub mdiPedidosProcessados_Click()
'frmPedidosProcessados.Show
'End Sub

Private Sub mdiPosGeralNeg_Click()
MsgBox "Fun��o n�o dispon�vel" 'frmPosGeralNegociacao.Show
End Sub

Private Sub mdiPrazoEntrega_Click()
MsgBox "Fun��o n�o dispon�vel" 'frmPrazoEntrega.Show vbModal
End Sub

Private Sub mdiPrazoEntregaCE_Click()
MsgBox "Fun��o n�o dispon�vel" 'frmPrazoEntrega.Show vbModal
End Sub

Private Sub mdiProdConsig_Click()
MsgBox "Fun��o n�o dispon�vel" 'impConsigProd.Show vbModal
End Sub

Private Sub mdiProdDiaria_Click()
MsgBox "Fun��o n�o dispon�vel" 'frmImpProdDiaria.Show
End Sub

Private Sub mdiProdUnidadeFabril_Click()
'MsgBox "Rotina em Manuten��o"
MsgBox "Fun��o n�o dispon�vel" 'frmProdGalpao.Show
End Sub

Private Sub mdiProduto_Click()
frmProduto.Show
End Sub

Private Sub mdiProdutoEstoque_Click()
MsgBox "Fun��o n�o dispon�vel " 'frmCadastroProdutosEstoque.Show
End Sub

Private Sub mdiProdutoPeriodo_Click()
MsgBox "Fun��o n�o dispon�vel " 'frmProdutoPeriodo.Show
End Sub

Private Sub mdiProdutosIn_Click()
frmProdutosDeEntrada.Show
End Sub

Private Sub mdiRecebimentos_Click()

Call Rotina_AbrirBanco

glb.Open "Select * from Global where chDataAbertura = ('" & DataHojeInvertida & "')", db, 3, 3
If glb.EOF Then
   MsgBox ("Aten��o: O sistema encontra-se fechado. Esta fun��o s� pode ser usada ap�s a abertura do sitema."), vbInformation
   Exit Sub
End If

frmControleFinanceiro.Show
End Sub

'Private Sub mdiReciboPagamento_Click()
'MsgBox "Fun��o n�o dispon�vel" 'frmReciboPagamento.Show vbModal
'End Sub

Private Sub mdiRecursosHumanos_Click()
MsgBox "Em Desenvolvimento"
End Sub

Private Sub mdiRelPessoa_Click()
MsgBox "Fun��o n�o dispon�vel" 'impRelPessoa.Show
End Sub

Private Sub mdiReprogFinanc_Click()
'MsgBox ("Fun��o em manuten��o")
frmReprogFinanc.Show
End Sub

Private Sub mdiRoteirizador_Click()
MsgBox "Fun��o n�o dispon�vel" 'frmRoteirizador.Show
End Sub

Private Sub mdiSaidasprodutos_Click()
'MsgBox ("Fun��o n�o dispon�vel. Em manuten��o.")
MsgBox "Fun��o n�o dispon�vel" 'frmProdutoSaida.Show
End Sub

Private Sub mdiSair_Click()
Dim Resp As String

Resp = MsgBox("Sa�da do sistema solicitada. Confirma???", vbExclamation + vbYesNo)

If Resp = vbYes Then
   Call Rotina_AbrirBanco
      
   rs.Open "select * from Usuario where chNome = ('" & glbUsuario & "')", db, 3, 3
   If rs.EOF Then
      Call FechaDB
   Else
      rs!usustatus = 0
      rs!usuMaquina = Empty
      rs.Update
      FechaDB
      End
   End If
Else
   MsgBox ("Sa�da cancelada")
End If
End Sub

Private Sub mdiTesteGrid_Click()
Dim senha As String
senha = InputBox("Informe Senha para entrar nesta fun��o.")
If Not senha = "Goiaba" Then
   MsgBox ("Voce n�o esta habilitado para esta fun��o.")
Else
   MsgBox "Fun��o n�o dispon�vel " 'TesteGrid.Show
End If
End Sub

Private Sub mdiUnidadeEmbalagem_Click()
MsgBox ("Fun��o n�o Dispon�vel") 'frmUnidadeEmbalagem.Show
End Sub

Private Sub mdiUnidadeMedida_Click()
MsgBox "Fun��o n�o dispon�vel " 'frmUnidadeMedida.Show
End Sub

Private Sub mdiUnidadeOperacional_Click()
frmUnidadeOperacional.Show
End Sub

Private Sub mdiUsuario_Click()

frmUsuario.Show
End Sub

Private Sub mdiUsuarioSenha_Click()

frmUsuarioSenha.Show vbModal
'Set TabUsuario = dbSHB.OpenRecordset("Usuario")
'    TabUsuario.Index = "IndUsuario"
Call Rotina_AbrirBanco
usu.Open "Select * from Usuario where chNome = ('" & glbUsuario & "')", db, 3, 3
If usu.EOF Then
   MsgBox ("Usuario n�o cadastrado"), vbCritical
   Call FechaDB
   Exit Sub
End If

mdiHabilitacaoSistema.Enabled = True

mdiPessoa.Enabled = False
mdiNeg.Enabled = False
mdiColaboradores = False
mdiParametros.Enabled = False
mdiFinanceiro.Enabled = False
mdiProducao.Enabled = False
If ChaveCompilando = 1 Then
   mdiMateriaisEst.Enabled = True
Else
   mdiMateriaisEst.Enabled = False
End If
mdiRelatorios.Enabled = False
mdiHabilitacao.Enabled = True
mdiSupervisor.Enabled = False
'TabUsuario.Close
End Sub

Private Sub mdiValePedagio_Click()
MsgBox "Fun��o n�o dispon�vel" 'frmValePedagio.Show
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Resp As String
    Resp = MsgBox("Sa�da do sistema solicitada. Confirma???", vbExclamation + vbYesNo)
    
If Resp = vbYes Then
   Call Rotina_AbrirBanco
      
   usu.Open "select * from Usuario where chNome = ('" & glbUsuario & "')", db, 3, 3
   acUsu = acUsu + 1
   If usu.EOF Then
      Call FechaDB
   Else
      usu!usustatus = 0
      usu!usuMaquina = Empty
      usu.Update
      FechaDB
      End
   End If
Else
   MsgBox ("Sa�da cancelada")
End If
End Sub

Public Sub VerificaASO()

PessoaAnterior = Empty

ano = Year(Date)
Mes = Month(Date)
Dia = Day(Date)

DataHojeInvertida = ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")

EncontreiAviso = 0

'Call Rotina_AbrirBanco

asoa.Open "Select * from AsoAgenda where asoaStatus = ('" & 0 & "')", db, 3, 3
If asoa.EOF Then
   Call FechaDB
   Exit Sub
End If

PessoaAnterior = Empty

asoa.MoveFirst

Do While (Not asoa.EOF) And EncontreiAviso = 0

   If Not asoa!chPessoa = PessoaAnterior Then
      If Not asoa!chPessoa = PessoaAnterior Then
         If pes.State = 1 Then
            pes.Close: Set pes = Nothing
         End If
         
         pes.Open "Select * from Pessoa where pesRazaoSocial = ('" & asoa!chPessoa & "')", db, 3, 3
         If pes.EOF Then
            MsgBox ("Pessoa n�o encontrado. Comunicar ao analista respons�vel."), vbCritical
            Call FechaDB
            Exit Sub
         End If
      End If
   End If
   
   If pes!pesStatusPessoa = 0 Then
      If asoe.State = 1 Then
         asoe.Close: Set asoe = Nothing
      End If
      
      asoe.Open "Select * from AsoExame where chNomeExame = ('" & asoa!chNomeExame & "')", db, 3, 3
      If Not asoe.EOF Then
         If asoe!exmUnidTempo = 0 Then
            DataDias = Date + asoe!exmPrazoAviso
            ano = Year(DataDias)
            Mes = Month(DataDias)
            Dia = Day(DataDias)
            DataBase = ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")
         Else
            If asoe!exmUnidTempo = 1 Then
               ano = Year(Date)
               Mes = Month(Date)
               Mes = Mes + asoe!exmPrazoAviso
               If Mes > 12 Then
                  ano = Year(Date)
                  ano = ano + 1
                  Mes = Mes - 12
               End If
               Dia = Day(Date)
               DataBase = ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")
            Else
               ano = Year(Date)
               ano = ano + asoe!exmPrazoAviso
               Mes = Month(Date)
               Dia = Day(Date)
               DataBase = ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")
            End If
         End If
         
         'AnoDb = Year(asoa!asoaDataProxExame)
         'MesDb = Month(asoa!asoaDataProxExame)
         'DiaDb = Day(asoa!asoaDataProxExame)
          DataInvertida = Format$(asoa!asoaDataProxExame, "yyyy-mm-dd")
         'DataInvertida = AnoDb & "-" & Format(MesDb, "00") & "-" & Format$(DiaDb, "00")
         
         If (DataInvertida > DataHojeInvertida) Or ((DataInvertida < DataHojeInvertida) And asoa!asoaStatus = 0) Then
            If Not (DataInvertida > DataBase) Then
               EncontreiAviso = 1
            End If
         End If
      
      
      End If
        
   End If
   
   'PessoaAnterior = asoa!chPessoa
   
   asoa.MoveNext

Loop

'Call FechaDB

 
End Sub

Public Sub VerificaTreinamento()
ano = Year(Date)
Mes = Month(Date)
Dia = Day(Date)

DataHojeInvertida = ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")

EncontreiAviso = 0

'Call Rotina_AbrirBanco

agcto.Open "Select * from TreinamentoAgenda where agctoStatus = ('" & 0 & "')", db, 3, 3
If agcto.EOF Then
   Call FechaDB
   Exit Sub
End If

agcto.MoveFirst

Do While Not agcto.EOF And EncontreiAviso = 0

   If Not agcto!chPessoa = PessoaAnterior Then
      
      If cto.State = 1 Then
         cto.Close: Set cto = Nothing
      End If
      
      cto.Open "Select * from Treinamento where chNomeCurso = ('" & agcto!chNomeCurso & "')", db, 3, 3
      If Not cto.EOF Then
         DataDias = Date + cto!ctoAvisoEm
         ano = Year(DataDias)
         Mes = Month(DataDias)
         Dia = Day(DataDias)
         DataBase = ano & "-" & Format$(Mes, "00") & "-" & Format$(Dia, "00")
   
         AnoDb = Year(agcto!agctoDataProxCurso)
         MesDb = Month(agcto!agctoDataProxCurso)
         DiaDb = Day(agcto!agctoDataProxCurso)
         
         DataInvertida = AnoDb & "-" & Format(MesDb, "00") & "-" & Format$(DiaDb, "00")
         
         If (DataInvertida > DataHojeInvertida) Or (DataInvertida < DataHojeInvertida) Then
            If Not (DataInvertida > DataBase) Then
               If pes.State = 1 Then
                  pes.Close: Set pes = Nothing
               End If
         
               pes.Open "Select * from Pessoa where pesRazaoSocial = ('" & agcto!chPessoa & "')", db, 3, 3
               If Not pes.EOF Then
                  If Not pes!pesStatusPessoa = 3 Then
                        EncontreiAviso = 1
                  End If
               End If
            End If
          End If
          PessoaAnterior = agcto!chPessoa
      End If
   End If

   agcto.MoveNext

Loop
End Sub

Public Sub VerificaReembolso()

EncontreiAviso = 0

'Call Rotina_AbrirBanco

Rmb.Open "Select * from Reembolso where rmbStatusReembolso = ('" & 0 & "')", db, 3, 3
If Rmb.EOF Then
   Call FechaDB
   Exit Sub
Else
   EncontreiAviso = 1
End If

'Call FechaDB

End Sub

Public Sub VerificaEquipamento()

Dim EquipTipoAnterior As String
Dim ChaveAuxiliar As String
Dim DataProxManut As Date

EncontreiAviso = 0

ChaveAuxiliar = "VENCIDO"

Call Rotina_AbrirBanco

eqpt.Open "Select * from Equipamento", db, 3, 3
If eqpt.EOF Then
   MsgBox ("Cadastro de Equipamentos vazio. Comunicar ao analista respons�vel."), vbCritical
   Call FechaDB
   Exit Sub
End If

eqpt.MoveFirst

Do While Not eqpt.EOF And EncontreiAviso = 0
   If Not EquipTipoAnterior = eqpt!eqptTipoEquipamento Then
      EquipTipoAnterior = eqpt!eqptTipoEquipamento
      If teq.State = 1 Then
         teq.Close: Set teq = Nothing
      End If
      teq.Open "Select * from EquipamentoTipo where chTipoDeEquipamento = ('" & eqpt!eqptTipoEquipamento & "')", db, 3, 3
      If teq.EOF Then
         MsgBox ("Erro no acesso a Tipo de Equipamento."), vbCritical
         Call FechaDB
         Exit Sub
      End If
   End If

   If Not IsNull(eqpt!eqptDataValidade) Then
      'DataProxManut = eqpt!eqptDataValidade
      If ((eqpt!eqptDataValidade - teq!teqDiasAntecedencia) < Date) And Not (eqpt!eqptStatusCalibracao = "EM CALIBRA��O") Then
         EncontreiAviso = 1
      End If
   Else
      MsgBox ("Data validade invalida o processo."), vbCritical
      Call FechaDB
      Exit Sub
   End If

   eqpt.MoveNext

Loop

Call FechaDB

End Sub

Private Sub mdiValoresPagosRecebidosTrimestre_Click()
frmValoresPagosRecebidosTrimestre.Show
End Sub
